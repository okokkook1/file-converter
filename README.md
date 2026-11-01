# 檔案轉換與排序合併工具 (File Converter & Merger)

支援 **Word (DOC/DOCX) ⇄ PDF 互轉**、**DOCX ⇄ ODT 互轉**、**圖片 ⇄ PDF 互轉**，以及**同檔案類型自訂排序合併**。

---

## 專案位置
`C:\Users\okokk\.gemini\antigravity\scratch\file_converter`

建議可將此目錄設定為 Antigravity 的當前工作區。

---

## 功能特點

1. **Word (DOC / DOCX) ⇄ PDF 互相轉換**
   - **Word ➔ PDF**：支援現代 `.docx` 與舊版 `.doc`，藉由 LibreOffice Headless 核心轉換，版面 100% 精確還原。
   - **PDF ➔ Word**：使用 `pdf2docx` 模組，完整解析並還原文字、段落、格式、字體與表格排版為可編輯的 `.docx` 文件。
2. **DOCX ⇄ ODT 雙向互相轉換**
   - 整合 LibreOffice Headless 引擎，轉換排版零跑版、高保真。
3. **圖片檔 ⇄ PDF 雙向互相轉換**
   - **圖片 ➔ PDF**：支援 JPG、PNG、BMP、WEBP、TIFF 等常見格式，可單檔轉換或批次合併。
   - **PDF ➔ 圖片**：使用高解析度渲染引擎（pypdfium2），將 PDF 各頁精確匯出為獨立高品質圖片檔（PNG）。
4. **同檔案類型自選排序與合併**
   - **清單排序功能**：支援「▲ 上移一層」與「▼ 下移一層」，任意排列合併前後順序。
   - **Word 合併**：依照自選順序合併多份 Word 文件（支援 `.docx` 與 `.doc` 混合合併）。
   - **PDF 合併**：依照自選順序將多個 PDF 檔合併為單一 PDF 檔。
   - **圖片合併**：依照自選順序將多張圖片合併成一份多頁 PDF 檔。
   - **ODT 合併**：依照自選順序合併多份 ODT 文件。
   - **⚡ 自動辨識合併**：一鍵判斷清單格式並執行對應合併。

---

## 快速啟動方式

### 方式 0：直接下載 Windows 執行檔 (免安裝 Python)
至 [GitHub Releases 頁面](https://github.com/okokkook1/file-converter/releases) 下載編譯完成的免安裝單一檔案：
👉 **[下載最新版 FileConverter.exe](https://github.com/okokkook1/file-converter/releases/latest/download/FileConverter.exe)**

### 方式 1：雙擊批次檔啟動 (推薦)
直接在檔案總管中雙擊：
```
run_app.bat
```

### 方式 2：使用終端機啟動
```powershell
cd C:\Users\okokk\.gemini\antigravity\scratch\file_converter
python app.py
```

---

## 程式碼呼叫範例 (Python API)

您也可以在其它 Python 程式碼中直接調用核心模組 `file_converter.py`：

```python
import file_converter as fc

# 1. Word (doc / docx) 轉 PDF
fc.convert_doc_to_pdf("report.docx", "report.pdf")
fc.convert_doc_to_pdf("legacy.doc", "legacy.pdf")

# 2. PDF 轉 Word (.docx)
fc.convert_pdf_to_docx("document.pdf", "editable.docx")

# 3. DOCX ⇄ ODT 互轉
fc.convert_docx_to_odt("report.docx", "report.odt")
fc.convert_odt_to_docx("note.odt", "note.docx")

# 4. 圖片轉多頁 PDF (依自選順序)
fc.convert_images_to_pdf(["page1.png", "page2.jpg"], "output.pdf")

# 5. PDF 轉圖片 (每頁匯出為 PNG)
fc.convert_pdf_to_images("document.pdf", output_dir="./extracted_images")

# 6. 排序合併 PDF
fc.merge_pdfs(["cover.pdf", "body.pdf", "appendix.pdf"], "final.pdf")

# 7. 排序合併 Word (支援 docx 與 doc)
fc.merge_docx(["chapter1.docx", "chapter2.doc"], "full_book.docx")
```
