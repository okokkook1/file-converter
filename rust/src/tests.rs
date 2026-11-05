use std::{fs, io::Read};

use super::*;

#[test]
fn extension_groups_match_supported_files() {
    assert_eq!(file_kind(Path::new("chapter.DOC")), Some(FileKind::Word));
    assert_eq!(file_kind(Path::new("photo.TIF")), Some(FileKind::Image));
    assert_eq!(file_kind(Path::new("archive.zip")), None);
}

#[test]
#[ignore = "requires the PDFium runtime"]
fn pdf_images_and_page_order() -> Result<()> {
    let temp = tempfile::tempdir()?;
    let root = temp.path();
    let first = root.join("紅.png");
    let second = root.join("藍.bmp");
    let mut red = image::RgbaImage::from_pixel(24, 40, image::Rgba([255, 0, 0, 255]));
    for y in 17..=23 {
        for x in 9..=15 {
            red.put_pixel(x, y, image::Rgba([0, 0, 0, 0]));
        }
    }
    red.save(&first)?;
    RgbImage::from_pixel(40, 24, Rgb([0, 0, 255])).save(&second)?;
    let combined = root.join("圖片.pdf");
    images_to_pdf(&[first.clone(), second.clone()], &combined)?;
    let rendered = pdf_to_images(&combined, root)?;
    assert_eq!(rendered.len(), 2);
    assert!(rendered[0].ends_with("圖片_page_001.png"));
    let page_one = image::open(&rendered[0])?.to_rgb8();
    let page_two = image::open(&rendered[1])?.to_rgb8();
    assert_eq!(page_one.dimensions(), (48, 80));
    assert_eq!(page_two.dimensions(), (80, 48));
    assert_eq!(page_one.get_pixel(8, 8).0, [255, 0, 0]);
    assert_eq!(page_one.get_pixel(24, 40).0, [255, 255, 255]);
    assert_eq!(page_two.get_pixel(40, 24).0, [0, 0, 255]);

    let blue_pdf = root.join("blue.pdf");
    let red_pdf = root.join("red.pdf");
    let merged = root.join("merged.pdf");
    images_to_pdf(&[second], &blue_pdf)?;
    images_to_pdf(&[first], &red_pdf)?;
    merge_pdfs(&[blue_pdf, red_pdf], &merged)?;
    let pages = pdf_to_images(&merged, root)?;
    assert_eq!(
        image::open(&pages[0])?.to_rgb8().get_pixel(8, 8).0,
        [0, 0, 255]
    );
    assert_eq!(
        image::open(&pages[1])?.to_rgb8().get_pixel(8, 8).0,
        [255, 0, 0]
    );

    let corrupt = root.join("corrupt.png");
    fs::write(&corrupt, b"invalid image")?;
    let previous = fs::read(&merged)?;
    assert!(images_to_pdf(&[corrupt], &merged).is_err());
    assert_eq!(fs::read(&merged)?, previous);
    Ok(())
}

fn xml(path: &Path, name: &str) -> Result<String> {
    let mut archive = zip::ZipArchive::new(fs::File::open(path)?)?;
    let mut contents = String::new();
    archive.by_name(name)?.read_to_string(&mut contents)?;
    Ok(contents)
}

#[test]
#[ignore = "requires LibreOffice"]
fn office_conversion_and_merge_order() -> Result<()> {
    let temp = tempfile::tempdir()?;
    let root = temp.path();
    let first = root.join("第一 章.html");
    let second = root.join("第二 章.html");
    fs::write(
        &first,
        "<html><body><h1>ALPHA START</h1><p>First chapter.</p></body></html>",
    )?;
    fs::write(
        &second,
        "<html><body><h1>BETA START</h1><table border='1'><tr><td>Second table</td></tr></table></body></html>",
    )?;
    let docx = first.with_extension("docx");
    let doc = second.with_extension("doc");
    office::convert(&first, &docx)?;
    office::convert(&second, &doc)?;

    let merged = root.join("合併結果.docx");
    merge_documents(&[doc.clone(), docx.clone()], &merged)?;
    let body = xml(&merged, "word/document.xml")?;
    assert!(body.find("BETA START").unwrap() < body.find("ALPHA START").unwrap());
    assert!(body.contains("Second table") && body.contains("<w:tbl>"));

    let first_odt = first.with_extension("odt");
    let second_odt = second.with_extension("odt");
    convert_office(&docx, &first_odt)?;
    office::convert(&second, &second_odt)?;
    let merged_odt = root.join("合併結果.odt");
    merge_documents(&[first_odt.clone(), second_odt], &merged_odt)?;
    let body = xml(&merged_odt, "content.xml")?;
    assert!(
        matches!((body.find("ALPHA"), body.find("BETA")), (Some(first), Some(second)) if first < second),
        "{body}"
    );
    let roundtrip = root.join("roundtrip.docx");
    convert_office(&first_odt, &roundtrip)?;
    assert!(xml(&roundtrip, "word/document.xml")?.contains("ALPHA START"));

    let pdf = root.join("legacy.pdf");
    convert_office(&doc, &pdf)?;
    let editable = root.join("editable.docx");
    convert_office(&pdf, &editable)?;
    assert!(xml(&editable, "word/document.xml")?.contains("BETA START"));
    let legacy = root.join("editable.doc");
    convert_office(&pdf, &legacy)?;
    assert!(fs::read(&legacy)?.starts_with(&[0xd0, 0xcf, 0x11, 0xe0]));
    Ok(())
}
