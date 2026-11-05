use std::{
    fs,
    path::{Path, PathBuf},
    sync::OnceLock,
};

use anyhow::{Context, Result, ensure};
use image::{DynamicImage, Rgb, RgbImage};
use pdfium_render::prelude::*;

mod office;
#[cfg(test)]
mod tests;
pub use office::find_libreoffice;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FileKind {
    Word,
    Odt,
    Pdf,
    Image,
}

pub const EXTENSIONS: &[&str] = &[
    "doc", "docx", "odt", "pdf", "png", "jpg", "jpeg", "bmp", "webp", "tiff", "tif",
];

pub fn extension(path: &Path) -> String {
    path.extension()
        .unwrap_or_default()
        .to_string_lossy()
        .to_ascii_lowercase()
}

pub fn file_kind(path: &Path) -> Option<FileKind> {
    Some(match extension(path).as_str() {
        "doc" | "docx" => FileKind::Word,
        "odt" => FileKind::Odt,
        "pdf" => FileKind::Pdf,
        "png" | "jpg" | "jpeg" | "bmp" | "webp" | "tiff" | "tif" => FileKind::Image,
        _ => return None,
    })
}

pub fn convert_office(input: &Path, output: &Path) -> Result<()> {
    let from = extension(input);
    let to = extension(output);
    ensure!(
        matches!(
            (from.as_str(), to.as_str()),
            ("doc" | "docx", "pdf") | ("docx", "odt") | ("odt" | "pdf", "docx") | ("pdf", "doc")
        ),
        "不支援從 .{from} 轉為 .{to}"
    );
    if from == "pdf" && to == "doc" {
        let temp = tempfile::tempdir()?;
        let docx = temp.path().join("converted.docx");
        office::convert(input, &docx)?;
        office::convert(&docx, output)
    } else {
        office::convert(input, output)
    }
}

pub fn merge_documents(inputs: &[PathBuf], output: &Path) -> Result<()> {
    ensure!(inputs.len() >= 2, "合併至少需要兩個檔案");
    let kind = file_kind(&inputs[0]);
    ensure!(
        matches!(kind, Some(FileKind::Word | FileKind::Odt))
            && inputs.iter().all(|input| file_kind(input) == kind)
            && file_kind(output) == kind,
        "請選擇相容的文件格式"
    );
    let temp = tempfile::tempdir()?;
    let mut docx_inputs = Vec::with_capacity(inputs.len());
    for (index, input) in inputs.iter().enumerate() {
        if extension(input) == "docx" {
            docx_inputs.push(input.clone());
        } else {
            let docx = temp.path().join(format!("input_{index}.docx"));
            office::convert(input, &docx)?;
            docx_inputs.push(docx);
        }
    }
    if extension(output) == "docx" {
        office::merge_docx(&docx_inputs, output)
    } else {
        let docx = temp.path().join("merged.docx");
        office::merge_docx(&docx_inputs, &docx)?;
        office::convert(&docx, output)
    }
}

fn pdfium() -> Result<&'static Pdfium> {
    static ENGINE: OnceLock<Result<Pdfium, String>> = OnceLock::new();
    ENGINE
        .get_or_init(|| open_pdfium().map_err(|error| format!("{error:#}")))
        .as_ref()
        .map_err(|error| anyhow::anyhow!("{error}"))
}

fn open_pdfium() -> Result<Pdfium> {
    let executable = std::env::current_exe()?;
    let directories = [
        executable.parent(),
        executable.parent().and_then(Path::parent),
    ]
    .into_iter()
    .flatten();
    let dll = std::env::var_os("PDFIUM_PATH")
        .map(PathBuf::from)
        .or_else(|| {
            directories
                .map(|path| path.join("pdfium.dll"))
                .find(|path| path.is_file())
        });
    let bindings = match dll {
        Some(path) => Pdfium::bind_to_library(path),
        None => Pdfium::bind_to_system_library(),
    }
    .context("無法載入 PDFium。請將相容的 pdfium.dll 放在程式旁")?;
    Ok(Pdfium::new(bindings))
}

// Stage an output in its destination directory so a failure leaves the previous file intact.
fn write_output(path: &Path, write: impl FnOnce(&Path) -> Result<()>) -> Result<()> {
    let parent = path
        .parent()
        .filter(|dir| !dir.as_os_str().is_empty())
        .unwrap_or(Path::new("."));
    fs::create_dir_all(parent)?;
    let suffix = format!(".{}", extension(path));
    let temp = tempfile::Builder::new()
        .suffix(&suffix)
        .tempfile_in(parent)?
        .into_temp_path();
    write(&temp)?;
    temp.persist(path)?;
    Ok(())
}

pub fn merge_pdfs(inputs: &[PathBuf], output: &Path) -> Result<()> {
    ensure!(inputs.len() >= 2, "合併至少需要兩個 PDF");
    let pdfium = pdfium()?;
    let mut merged = pdfium.create_new_pdf()?;
    for input in inputs {
        let source = pdfium
            .load_pdf_from_file(input, None)
            .with_context(|| format!("無法開啟 {}", input.display()))?;
        merged.pages_mut().append(&source)?;
    }
    write_output(output, |path| Ok(merged.save_to_file(path)?))
}

pub fn images_to_pdf(inputs: &[PathBuf], output: &Path) -> Result<()> {
    ensure!(!inputs.is_empty(), "請至少加入一張圖片");
    let pdfium = pdfium()?;
    let mut document = pdfium.create_new_pdf()?;
    for input in inputs {
        let decoded =
            image::open(input).with_context(|| format!("無法讀取 {}", input.display()))?;
        let rgba = decoded.to_rgba8();
        let rgb = RgbImage::from_fn(rgba.width(), rgba.height(), |x, y| {
            let [r, g, b, a] = rgba.get_pixel(x, y).0;
            let a = u16::from(a);
            Rgb([r, g, b].map(|c| ((u16::from(c) * a + 255 * (255 - a) + 127) / 255) as u8))
        });
        let width = PdfPoints::new(rgb.width() as f32);
        let height = PdfPoints::new(rgb.height() as f32);
        let object =
            PdfPageImageObject::new_with_width(&document, &DynamicImage::ImageRgb8(rgb), width)?;
        let mut page = document
            .pages_mut()
            .create_page_at_end(PdfPagePaperSize::Custom(width, height))?;
        page.objects_mut().add_image_object(object)?;
    }
    write_output(output, |path| Ok(document.save_to_file(path)?))
}

pub fn pdf_to_images(input: &Path, output: &Path) -> Result<Vec<PathBuf>> {
    let pdfium = pdfium()?;
    let document = pdfium
        .load_pdf_from_file(input, None)
        .with_context(|| format!("無法開啟 {}", input.display()))?;
    let stem = input
        .file_stem()
        .context("無法取得 PDF 檔名")?
        .to_string_lossy();
    let config = PdfRenderConfig::new().scale_page_by_factor(2.0);
    document
        .pages()
        .iter()
        .enumerate()
        .map(|(index, page)| {
            let path = output.join(format!("{stem}_page_{:03}.png", index + 1));
            let image = page.render_with_config(&config)?.as_image()?;
            write_output(&path, |temp| Ok(image.save(temp)?))?;
            Ok(path)
        })
        .collect()
}
