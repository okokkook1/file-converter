# 檔案轉換與排序合併工具

Windows x64 桌面工具。可先調整檔案清單順序，再執行批次轉換或合併。

## Windows x64 下載

[下載 Rust 免安裝版](https://github.com/okokkook1/file-converter/releases/download/v2.0.0/file-converter-rust-windows-x64.zip)。完整解壓後雙擊 `FileConverter.exe`，並保留同一資料夾內的執行期資源。

- 支援 DOCX ⇄ ODT、DOC／DOCX → PDF、PDF → DOCX、圖片 → PDF，以及 PDF 逐頁匯出 PNG。
- 可依指定順序合併 PDF、Word、ODT，或將多張圖片合併成一份 PDF。
- Office 文件相關操作需要 LibreOffice。可透過 `LIBREOFFICE_PATH` 指定 `soffice.com` 或 `soffice.exe`。
- PDF 轉 DOCX 會以可編輯的文字方塊重建版面，複雜表格可能需要手動整理。PDF 以頁面為單位合併，原檔書籤需重新建立。

## 從 Rust 原始碼編譯

在 Windows x64 安裝 Rust MSVC 工具鏈與 Visual Studio C++ Build Tools，並在專案根目錄執行：

```powershell
cargo build --manifest-path rust/Cargo.toml --release --locked
```

在 `rust` 資料夾內，將下載檔內的 `pdfium.dll` 複製到 `target\release\`，即可執行 `target\release\file-converter.exe`。

要執行含實際引擎的測試，請先將執行期資源放到上述位置，並在 `rust` 資料夾內執行：

```powershell
$env:PDFIUM_PATH = (Resolve-Path .\target\release\pdfium.dll).Path
cargo test --release --locked -- --include-ignored
```

## Python 舊版

[Python 版使用說明](README-python.md) · [Python 版 Windows 下載](https://github.com/okokkook1/file-converter/releases/tag/v1.0.0)
