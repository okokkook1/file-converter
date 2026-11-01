"""
test_converter.py - 自動化功能測試
驗證：
1. DOCX <-> ODT 雙向轉換
2. 圖片 -> PDF -> 圖片 雙向轉換
3. 多 PDF 排序合併
4. 多 DOCX 排序合併
5. 多 ODT 排序合併
6. 多圖片 排序合併為單一多頁 PDF
"""

import os
from pathlib import Path
from PIL import Image, ImageDraw
from docx import Document
import pypdf
import file_converter as fc


def run_tests():
    test_dir = Path("test_workspace")
    if test_dir.exists():
        import shutil
        shutil.rmtree(test_dir)
    test_dir.mkdir(parents=True, exist_ok=True)

    print("=== [1] 建立測試資料 ===")
    # 建立 2 個 DOCX
    docx1 = test_dir / "doc1.docx"
    d1 = Document()
    d1.add_heading("文件 1 標題", level=1)
    d1.add_paragraph("這是文件 1 的內容段落。")
    d1.save(str(docx1))

    docx2 = test_dir / "doc2.docx"
    d2 = Document()
    d2.add_heading("文件 2 標題", level=1)
    d2.add_paragraph("這是文件 2 的內容段落。")
    d2.save(str(docx2))

    # 建立 2 張圖片
    img1 = test_dir / "img1.png"
    im1 = Image.new("RGB", (200, 100), color=(255, 100, 100))
    d = ImageDraw.Draw(im1)
    d.text((10, 40), "Image 1", fill=(255, 255, 255))
    im1.save(str(img1))

    img2 = test_dir / "img2.png"
    im2 = Image.new("RGB", (200, 100), color=(100, 100, 255))
    d = ImageDraw.Draw(im2)
    d.text((10, 40), "Image 2", fill=(255, 255, 255))
    im2.save(str(img2))

    print("=== [2] 測試 DOCX <-> ODT 互轉 ===")
    odt1 = test_dir / "doc1.odt"
    fc.convert_docx_to_odt(str(docx1), str(odt1))
    assert odt1.exists(), "DOCX -> ODT 轉換失敗"
    print("  [V] DOCX -> ODT 成功")

    docx_back = test_dir / "doc1_back.docx"
    fc.convert_odt_to_docx(str(odt1), str(docx_back))
    assert docx_back.exists(), "ODT -> DOCX 轉換失敗"
    print("  [V] ODT -> DOCX 成功")

    # 製作第二個 ODT 用於合併測試
    odt2 = test_dir / "doc2.odt"
    fc.convert_docx_to_odt(str(docx2), str(odt2))

    print("=== [3] 測試 圖片 <-> PDF 互轉 ===")
    img_pdf = test_dir / "from_img.pdf"
    fc.convert_images_to_pdf([str(img1)], str(img_pdf))
    assert img_pdf.exists(), "圖片 -> PDF 轉換失敗"
    print("  [V] 圖片 -> PDF 成功")

    pdf_imgs = fc.convert_pdf_to_images(str(img_pdf), str(test_dir / "pdf_extracted"))
    assert len(pdf_imgs) == 1 and Path(pdf_imgs[0]).exists(), "PDF -> 圖片 轉換失敗"
    print("  [V] PDF -> 圖片 成功 (匯出圖片數:", len(pdf_imgs), ")")

    print("=== [4] 測試 檔案排序合併功能 ===")
    # 4.1 合併 PDF
    pdf_a = test_dir / "a.pdf"
    pdf_b = test_dir / "b.pdf"
    fc.convert_images_to_pdf([str(img1)], str(pdf_a))
    fc.convert_images_to_pdf([str(img2)], str(pdf_b))
    merged_pdf = test_dir / "merged.pdf"
    fc.merge_pdfs([str(pdf_a), str(pdf_b)], str(merged_pdf))
    reader = pypdf.PdfReader(str(merged_pdf))
    assert len(reader.pages) == 2, f"PDF 合併頁數錯誤: {len(reader.pages)}"
    print("  [V] 多個 PDF 排序合併成功 (頁數: 2)")

    # 4.2 合併 圖片 為單一 PDF
    merged_img_pdf = test_dir / "merged_images.pdf"
    fc.merge_images([str(img1), str(img2)], str(merged_img_pdf))
    reader_img = pypdf.PdfReader(str(merged_img_pdf))
    assert len(reader_img.pages) == 2, f"圖片合併 PDF 頁數錯誤: {len(reader_img.pages)}"
    print("  [V] 多張圖片排序合併為多頁 PDF 成功 (頁數: 2)")

    # 4.3 合併 DOCX
    merged_docx = test_dir / "merged.docx"
    fc.merge_docx([str(docx1), str(docx2)], str(merged_docx))
    assert merged_docx.exists(), "DOCX 合併失敗"
    print("  [V] 多個 DOCX 排序合併成功")

    # 4.4 合併 ODT
    merged_odt = test_dir / "merged.odt"
    fc.merge_odt([str(odt1), str(odt2)], str(merged_odt))
    assert merged_odt.exists(), "ODT 合併失敗"
    print("  [V] 多個 ODT 排序合併成功")

    print("=== [5] 測試 Word (.doc / .docx) <-> PDF 互轉 ===")
    # 5.1 DOCX -> PDF
    doc_to_pdf_file = test_dir / "docx_to_pdf.pdf"
    fc.convert_doc_to_pdf(str(docx1), str(doc_to_pdf_file))
    assert doc_to_pdf_file.exists(), "Word -> PDF 轉換失敗"
    print("  [V] Word (DOCX) -> PDF 成功")

    # 5.2 PDF -> DOCX
    pdf_to_docx_file = test_dir / "pdf_to_docx.docx"
    fc.convert_pdf_to_docx(str(doc_to_pdf_file), str(pdf_to_docx_file))
    assert pdf_to_docx_file.exists(), "PDF -> Word (DOCX) 轉換失敗"
    print("  [V] PDF -> Word (DOCX) 成功")

    # 5.3 製作 .doc (舊版) 測試 .doc -> PDF
    legacy_doc = test_dir / "legacy.doc"
    fc._run_soffice_convert(str(docx1), "doc", str(legacy_doc))
    assert legacy_doc.exists(), "產生 .doc 失敗"
    legacy_doc_to_pdf = test_dir / "legacy_doc_to_pdf.pdf"
    fc.convert_doc_to_pdf(str(legacy_doc), str(legacy_doc_to_pdf))
    assert legacy_doc_to_pdf.exists(), "舊版 .doc -> PDF 轉換失敗"
    print("  [V] 舊版 Word (.doc) -> PDF 成功")

    # 5.4 測試 DOC 與 DOCX 混合排序合併
    merged_mixed_word = test_dir / "merged_mixed.docx"
    fc.merge_docx([str(docx1), str(legacy_doc)], str(merged_mixed_word))
    assert merged_mixed_word.exists(), "DOC + DOCX 混合合併失敗"
    print("  [V] DOC + DOCX 混合排序合併成功")

    print("\n>>> 所有測試全部通過！<<<")


if __name__ == "__main__":
    run_tests()
