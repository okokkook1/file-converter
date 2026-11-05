#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::{
    path::{Path, PathBuf},
    process::Command,
    sync::mpsc,
    thread,
};

use anyhow::{Context, Result, ensure};
use core::{EXTENSIONS, FileKind};
use eframe::egui::{self, FontData, FontDefinitions, FontFamily};
use file_converter as core;
use rfd::FileDialog;

#[derive(Clone, Copy)]
enum Action {
    Convert(&'static str, &'static str),
    PdfImages,
    IndividualImages,
    Merge(FileKind),
    AutoMerge,
}

impl Action {
    fn accepts(self, path: &Path) -> bool {
        match self {
            Self::Convert("word", _) => core::file_kind(path) == Some(FileKind::Word),
            Self::Convert(from, _) => core::extension(path) == from,
            Self::PdfImages => core::file_kind(path) == Some(FileKind::Pdf),
            Self::IndividualImages => core::file_kind(path) == Some(FileKind::Image),
            Self::Merge(kind) => core::file_kind(path) == Some(kind),
            Self::AutoMerge => true,
        }
    }

    fn save_dialog(self) -> Option<FileDialog> {
        let Self::Merge(kind) = self else { return None };
        let (name, formats): (_, &[&str]) = match kind {
            FileKind::Pdf => ("merged_output.pdf", &["pdf"]),
            FileKind::Image => ("images_combined.pdf", &["pdf"]),
            FileKind::Word => ("merged_output.docx", &["docx", "doc"]),
            FileKind::Odt => ("merged_output.odt", &["odt"]),
        };
        let mut dialog = FileDialog::new()
            .set_title("儲存合併結果")
            .set_file_name(name);
        for ext in formats {
            dialog = dialog.add_filter(ext.to_uppercase(), &[*ext]);
        }
        Some(dialog)
    }

    fn run(self, inputs: Vec<PathBuf>, output: Option<PathBuf>) -> Result<(PathBuf, String)> {
        if let Self::Merge(kind) = self {
            let output = output.context("請選擇輸出位置")?;
            match kind {
                FileKind::Pdf => core::merge_pdfs(&inputs, &output)?,
                FileKind::Image => core::images_to_pdf(&inputs, &output)?,
                FileKind::Word | FileKind::Odt => core::merge_documents(&inputs, &output)?,
            }
            return Ok((
                output.parent().context("缺少輸出資料夾")?.into(),
                format!("已依序合併 {} 個檔案：{}", inputs.len(), output.display()),
            ));
        }
        let mut folder = PathBuf::new();
        let mut pages = 0;
        for input in &inputs {
            let parent = input.parent().context("無法取得輸入資料夾")?;
            folder = parent.to_path_buf();
            let result = match self {
                Self::Convert(_, to) => core::convert_office(input, &input.with_extension(to)),
                Self::IndividualImages => {
                    core::images_to_pdf(std::slice::from_ref(input), &input.with_extension("pdf"))
                }
                Self::PdfImages => {
                    let stem = input.file_stem().context("無法取得檔名")?.to_string_lossy();
                    folder = parent.join(format!("{stem}_images"));
                    core::pdf_to_images(input, &folder).map(|images| pages += images.len())
                }
                _ => unreachable!(),
            };
            result.with_context(|| format!("處理 {} 時失敗", input.display()))?;
        }
        let message = match self {
            Self::PdfImages => format!("已從 {} 個 PDF 匯出 {pages} 張 PNG", inputs.len()),
            _ => format!("已轉換 {} 個檔案", inputs.len()),
        };
        Ok((folder, message))
    }
}

type JobResult = Result<(PathBuf, String)>;

struct App {
    files: Vec<PathBuf>,
    selected: Option<usize>,
    merge_tab: bool,
    status: String,
    last_output: Option<PathBuf>,
    worker: Option<mpsc::Receiver<JobResult>>,
}

impl App {
    fn new(ctx: &egui::Context) -> Self {
        if let Ok(font) = std::fs::read(r"C:\Windows\Fonts\msjh.ttc") {
            let mut fonts = FontDefinitions::default();
            fonts
                .font_data
                .insert("jhenghei".into(), FontData::from_owned(font).into());
            for family in [FontFamily::Proportional, FontFamily::Monospace] {
                fonts
                    .families
                    .entry(family)
                    .or_default()
                    .push("jhenghei".into());
            }
            ctx.set_fonts(fonts);
        }
        let status = match core::find_libreoffice() {
            Some(_) => "就緒，已偵測到 LibreOffice",
            None => "未偵測到 LibreOffice，Office 文件相關操作需要安裝 LibreOffice",
        };
        Self {
            files: vec![],
            selected: None,
            merge_tab: false,
            status: status.into(),
            last_output: None,
            worker: None,
        }
    }

    fn add_files(&mut self) {
        if let Some(paths) = FileDialog::new()
            .set_title("新增檔案")
            .add_filter("所有支援檔案", EXTENSIONS)
            .pick_files()
        {
            for path in paths {
                if !self.files.contains(&path) {
                    self.files.push(path);
                }
            }
            self.status = format!("清單共有 {} 個檔案", self.files.len());
        }
    }

    fn move_file(&mut self, step: isize) {
        if let Some(index) = self.selected
            && let Some(target) = index
                .checked_add_signed(step)
                .filter(|&i| i < self.files.len())
        {
            self.files.swap(index, target);
            self.selected = Some(target);
        }
    }

    fn start(&mut self, ctx: &egui::Context, action: Action) -> Result<()> {
        let action = if let Action::AutoMerge = action {
            let kind = self
                .files
                .first()
                .and_then(|path| core::file_kind(path))
                .context("請先加入支援的檔案")?;
            ensure!(
                self.files
                    .iter()
                    .all(|path| core::file_kind(path) == Some(kind)),
                "自動合併需要清單中的檔案屬於同一類型"
            );
            Action::Merge(kind)
        } else {
            action
        };
        let inputs: Vec<_> = self
            .files
            .iter()
            .filter(|path| action.accepts(path))
            .cloned()
            .collect();
        let minimum = if matches!(
            action,
            Action::Merge(FileKind::Pdf | FileKind::Word | FileKind::Odt)
        ) {
            2
        } else {
            1
        };
        ensure!(
            inputs.len() >= minimum,
            "請至少加入 {minimum} 個符合格式的檔案"
        );
        let output = if let Some(dialog) = action.save_dialog() {
            let Some(path) = dialog.save_file() else {
                return Ok(());
            };
            let path = if path.extension().is_none() {
                path.with_extension(match action {
                    Action::Merge(FileKind::Word) => "docx",
                    Action::Merge(FileKind::Odt) => "odt",
                    _ => "pdf",
                })
            } else {
                path
            };
            let valid = match action {
                Action::Merge(FileKind::Image | FileKind::Pdf) => core::extension(&path) == "pdf",
                Action::Merge(kind) => core::file_kind(&path) == Some(kind),
                _ => unreachable!(),
            };
            ensure!(valid, "輸出檔案的副檔名與選取格式不符");
            Some(path)
        } else {
            None
        };
        let (send, receive) = mpsc::channel();
        let ctx = ctx.clone();
        self.status = "正在處理檔案…".into();
        self.worker = Some(receive);
        thread::spawn(move || {
            let _ = send.send(action.run(inputs, output));
            ctx.request_repaint();
        });
        Ok(())
    }
}

impl eframe::App for App {
    fn update(&mut self, ctx: &egui::Context, _: &mut eframe::Frame) {
        if let Some(worker) = &self.worker {
            match worker.try_recv() {
                Ok(Ok((dir, message))) => {
                    self.last_output = Some(dir);
                    self.status = message;
                    self.worker = None;
                }
                Ok(Err(error)) => {
                    self.status = format!("{error:#}");
                    self.worker = None;
                }
                Err(mpsc::TryRecvError::Disconnected) => {
                    self.status = "工作意外停止".into();
                    self.worker = None;
                }
                Err(mpsc::TryRecvError::Empty) => {}
            }
        }

        egui::TopBottomPanel::bottom("status").show(ctx, |ui| {
            ui.horizontal_wrapped(|ui| {
                if ui
                    .add_enabled(
                        self.last_output.is_some(),
                        egui::Button::new("開啟最近輸出資料夾"),
                    )
                    .clicked()
                    && let Some(path) = &self.last_output
                    && let Err(error) = Command::new("explorer.exe").arg(path).spawn()
                {
                    self.status = format!("無法開啟資料夾：{error}");
                }
                if self.worker.is_some() {
                    ui.spinner();
                }
                ui.label(&self.status);
            });
        });

        egui::TopBottomPanel::bottom("operations").show(ctx, |ui| {
            ui.horizontal(|ui| {
                ui.selectable_value(&mut self.merge_tab, false, "檔案互相轉換");
                ui.selectable_value(&mut self.merge_tab, true, "依序合併");
            });
            let actions: &[(&str, Action)] = if self.merge_tab {
                &[
                    ("自動辨識合併", Action::AutoMerge),
                    ("合併 PDF", Action::Merge(FileKind::Pdf)),
                    ("圖片合併成 PDF", Action::Merge(FileKind::Image)),
                    ("合併 Word", Action::Merge(FileKind::Word)),
                    ("合併 ODT", Action::Merge(FileKind::Odt)),
                ]
            } else {
                &[
                    ("DOCX → ODT", Action::Convert("docx", "odt")),
                    ("ODT → DOCX", Action::Convert("odt", "docx")),
                    ("Word → PDF", Action::Convert("word", "pdf")),
                    ("PDF → Word", Action::Convert("pdf", "docx")),
                    ("圖片 → 個別 PDF", Action::IndividualImages),
                    ("圖片 → 單一 PDF", Action::Merge(FileKind::Image)),
                    ("PDF → 逐頁 PNG", Action::PdfImages),
                ]
            };
            ui.add_enabled_ui(self.worker.is_none(), |ui| {
                ui.horizontal_wrapped(|ui| {
                    for &(label, action) in actions {
                        if ui.button(label).clicked()
                            && let Err(error) = self.start(ctx, action)
                        {
                            self.status = format!("{error:#}");
                        }
                    }
                });
            });
        });

        egui::CentralPanel::default().show(ctx, |ui| {
            ui.heading("檔案轉換與排序合併工具");
            ui.horizontal_wrapped(|ui| {
                if ui.button("新增檔案…").clicked() {
                    self.add_files();
                }
                if ui
                    .add_enabled(self.selected.is_some(), egui::Button::new("移除選取"))
                    .clicked()
                    && let Some(index) = self.selected
                {
                    self.files.remove(index);
                    self.selected = self.files.len().checked_sub(1).map(|last| index.min(last));
                }
                if ui.button("全部清空").clicked() {
                    self.files.clear();
                    self.selected = None;
                }
                if ui
                    .add_enabled(
                        self.selected.is_some_and(|i| i > 0),
                        egui::Button::new("上移"),
                    )
                    .clicked()
                {
                    self.move_file(-1);
                }
                if ui
                    .add_enabled(
                        self.selected.is_some_and(|i| i + 1 < self.files.len()),
                        egui::Button::new("下移"),
                    )
                    .clicked()
                {
                    self.move_file(1);
                }
            });
            ui.separator();
            egui::ScrollArea::both().show(ui, |ui| {
                for (index, path) in self.files.iter().enumerate() {
                    let label = format!("[{:02}] {}", index + 1, path.display());
                    if ui
                        .selectable_label(self.selected == Some(index), label)
                        .clicked()
                    {
                        self.selected = Some(index);
                    }
                }
                if self.files.is_empty() {
                    ui.label("加入檔案後，可調整清單順序再合併。");
                }
            });
        });
    }
}

fn main() -> eframe::Result {
    eframe::run_native(
        "多功能檔案轉換與合併工具",
        eframe::NativeOptions {
            viewport: egui::ViewportBuilder::default()
                .with_inner_size([800.0, 580.0])
                .with_min_inner_size([680.0, 480.0]),
            ..Default::default()
        },
        Box::new(|cc| Ok(Box::new(App::new(&cc.egui_ctx)))),
    )
}
