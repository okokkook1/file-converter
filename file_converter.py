"""
file_converter.py - 核心文件轉換與合併工具庫
支援：
1. DOCX <-> ODT 互相轉換 (使用 LibreOffice)
2. 圖片 <-> PDF 互相轉換 (Pillow, pypdfium2)
3. 同類型檔案自訂排序並合併：
   - 多個 PDF -> 單一 PDF (pypdf)
   - 多張圖片 -> 單一多頁 PDF (Pillow)
   - 多個 DOCX -> 單一 DOCX (docxcompose)
   - 多個 ODT -> 單一 ODT (LibreOffice + docxcompose)
"""

import os
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List, Optional
from PIL import Image
import pypdf
import pypdfium2 as pdfium
from docx import Document
from docxcompose.composer import Composer


def find_libreoffice() -> Optional[str]:
    """尋找系統中的 LibreOffice (soffice) 執行檔路徑"""
    env_path = os.environ.get("LIBREOFFICE_PATH")
    if env_path and os.path.isfile(env_path):
        return env_path

    # PATH 搜尋
    found = shutil.which("soffice") or shutil.which("libreoffice")
    if found:
        return found

    # Windows 常見路徑
    candidate_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in candidate_paths:
        if os.path.isfile(p):
            return p
    return None


def _run_soffice_convert(input_path: str, target_format: str, output_path: Optional[str] = None) -> str:
    soffice = find_libreoffice()
    if not soffice:
        raise RuntimeError("未找到 LibreOffice (soffice.exe)，請確認已安裝 LibreOffice 以支援 docx/odt 轉換。")

    in_file = Path(input_path).resolve()
    if not in_file.exists():
        raise FileNotFoundError(f"找不到輸入檔案: {in_file}")

    if output_path:
        out_target = Path(output_path).resolve()
    else:
        out_target = in_file.parent / f"{in_file.stem}.{target_format}"

    out_target.parent.mkdir(parents=True, exist_ok=True)

    # 使用獨立臨時目錄執行 LibreOffice 轉換，徹底避免覆蓋同目錄下現有檔案
    with tempfile.TemporaryDirectory() as tmp_dir:
        cmd = [
            soffice,
            "--headless",
            "--convert-to", target_format,
            str(in_file),
            "--outdir", str(tmp_dir)
        ]
        result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice 轉換失敗: {result.stderr or result.stdout}")

        produced_file = Path(tmp_dir) / f"{in_file.stem}.{target_format}"
        if not produced_file.exists():
            candidates = list(Path(tmp_dir).glob(f"*.{target_format}"))
            if candidates:
                produced_file = candidates[0]
            else:
                raise FileNotFoundError(f"預期的輸出檔案未生成: {produced_file}")

        if out_target.exists():
            out_target.unlink()
        shutil.move(str(produced_file), str(out_target))

    return str(out_target)


# ---------------- DOCX <-> ODT ----------------

def convert_docx_to_odt(docx_path: str, output_path: Optional[str] = None) -> str:
    """將 DOCX 轉換為 ODT"""
    return _run_soffice_convert(docx_path, "odt", output_path)


def convert_odt_to_docx(odt_path: str, output_path: Optional[str] = None) -> str:
    """將 ODT 轉換為 DOCX"""
    return _run_soffice_convert(odt_path, "docx", output_path)


# ---------------- Word (DOC / DOCX) <-> PDF ----------------

def convert_doc_to_pdf(doc_path: str, output_path: Optional[str] = None) -> str:
    """
    將 Word 文件 (.doc 或 .docx) 轉換為 PDF 檔案。
    使用 LibreOffice Headless 轉換，完美保持原始排版。
    """
    return _run_soffice_convert(doc_path, "pdf", output_path)


def convert_pdf_to_docx(pdf_path: str, output_path: Optional[str] = None) -> str:
    """
    將 PDF 檔案轉換為 Word DOCX 檔案。
    使用 pdf2docx 提取並還原文字、段落、樣式與表格排版。
    """
    in_pdf = Path(pdf_path).resolve()
    if not in_pdf.exists():
        raise FileNotFoundError(f"找不到 PDF 檔案: {in_pdf}")

    if output_path:
        out_docx = Path(output_path).resolve()
    else:
        out_docx = in_pdf.parent / f"{in_pdf.stem}.docx"

    out_docx.parent.mkdir(parents=True, exist_ok=True)

    from pdf2docx import Converter
    cv = Converter(str(in_pdf))
    cv.convert(str(out_docx))
    cv.close()
    return str(out_docx)


def convert_pdf_to_doc(pdf_path: str, output_path: Optional[str] = None) -> str:
    """
    將 PDF 轉換為舊版 Word .doc 檔案。
    先轉為 .docx 後透過 LibreOffice 轉為 .doc。
    """
    with tempfile.TemporaryDirectory() as tmp_dir:
        temp_docx = Path(tmp_dir) / f"{Path(pdf_path).stem}.docx"
        convert_pdf_to_docx(pdf_path, str(temp_docx))
        return _run_soffice_convert(str(temp_docx), "doc", output_path)


# ---------------- 圖片 <-> PDF ----------------

def convert_images_to_pdf(image_paths: List[str], output_pdf_path: str) -> str:
    """
    將一張或多張圖片（依給定順序）合併轉換為單一多頁 PDF。
    支援 JPG, PNG, BMP, WEBP 等 Pillow 支援的圖片格式。
    """
    if not image_paths:
        raise ValueError("請至少提供一個圖片檔案。")

    out_file = Path(output_pdf_path).resolve()
    out_file.parent.mkdir(parents=True, exist_ok=True)

    pil_images = []
    for p in image_paths:
        img_path = Path(p).resolve()
        if not img_path.exists():
            raise FileNotFoundError(f"找不到圖片檔案: {img_path}")
        im = Image.open(img_path)
        # 轉換為 RGB 避免透明通道在 PDF 發生相容性問題
        if im.mode in ("RGBA", "LA", "P"):
            rgb_im = Image.new("RGB", im.size, (255, 255, 255))
            if im.mode == "P":
                im = im.convert("RGBA")
            rgb_im.paste(im, mask=im.split()[-1] if im.mode in ("RGBA", "LA") else None)
            pil_images.append(rgb_im)
        else:
            pil_images.append(im.convert("RGB"))

    first_image = pil_images[0]
    rest_images = pil_images[1:] if len(pil_images) > 1 else []
    first_image.save(str(out_file), save_all=True, append_images=rest_images)
    return str(out_file)


def convert_pdf_to_images(pdf_path: str, output_dir: Optional[str] = None, fmt: str = "png", scale: float = 2.0) -> List[str]:
    """
    將 PDF 各頁轉換為圖片檔。
    scale=2.0 約相當於 144 DPI 高解析度渲染。
    """
    in_pdf = Path(pdf_path).resolve()
    if not in_pdf.exists():
        raise FileNotFoundError(f"找不到 PDF 檔案: {in_pdf}")

    out_folder = Path(output_dir).resolve() if output_dir else in_pdf.parent / f"{in_pdf.stem}_images"
    out_folder.mkdir(parents=True, exist_ok=True)

    pdf = pdfium.PdfDocument(str(in_pdf))
    saved_images = []
    total_pages = len(pdf)

    for i in range(total_pages):
        page = pdf.get_page(i)
        rendered = page.render(scale=scale).to_pil()
        out_image_path = out_folder / f"{in_pdf.stem}_page_{i+1:03d}.{fmt.lower()}"
        rendered.save(str(out_image_path))
        saved_images.append(str(out_image_path))

    return saved_images


# ---------------- 同類型檔案排序合併 ----------------

def merge_pdfs(pdf_paths: List[str], output_pdf_path: str) -> str:
    """依列表順序合併多個 PDF"""
    if len(pdf_paths) < 2:
        raise ValueError("合併至少需要 2 個 PDF 檔案。")

    out_file = Path(output_pdf_path).resolve()
    out_file.parent.mkdir(parents=True, exist_ok=True)

    writer = pypdf.PdfWriter()
    for p in pdf_paths:
        fpath = Path(p).resolve()
        if not fpath.exists():
            raise FileNotFoundError(f"找不到 PDF 檔案: {fpath}")
        writer.append(str(fpath))

    with open(out_file, "wb") as f:
        writer.write(f)
    writer.close()
    return str(out_file)


def merge_docx(docx_paths: List[str], output_docx_path: str) -> str:
    """依列表順序合併多個 Word 檔案 (支援 .docx 與 .doc)"""
    if len(docx_paths) < 2:
        raise ValueError("合併至少需要 2 個 Word 檔案。")

    out_file = Path(output_docx_path).resolve()
    out_file.parent.mkdir(parents=True, exist_ok=True)

    with tempfile.TemporaryDirectory() as tmp_dir:
        normalized_paths = []
        for i, p in enumerate(docx_paths):
            fpath = Path(p).resolve()
            if not fpath.exists():
                raise FileNotFoundError(f"找不到 Word 檔案: {fpath}")
            if fpath.suffix.lower() == ".doc":
                temp_docx = Path(tmp_dir) / f"converted_{i}.docx"
                _run_soffice_convert(str(fpath), "docx", str(temp_docx))
                normalized_paths.append(temp_docx)
            else:
                normalized_paths.append(fpath)

        master = Document(str(normalized_paths[0]))
        composer = Composer(master)

        for p in normalized_paths[1:]:
            doc = Document(str(p))
            composer.append(doc)

        composer.save(str(out_file))
        return str(out_file)


def merge_odt(odt_paths: List[str], output_odt_path: str) -> str:
    """
    依列表順序合併多個 ODT 檔案。
    透過 LibreOffice 轉為暫存 DOCX 合併後再轉回 ODT，確保排版樣式相容。
    """
    if len(odt_paths) < 2:
        raise ValueError("合併至少需要 2 個 ODT 檔案。")

    with tempfile.TemporaryDirectory() as tmpdir:
        temp_dir_path = Path(tmpdir)
        temp_docxs = []
        for i, odt in enumerate(odt_paths):
            tmp_docx = temp_dir_path / f"temp_{i}.docx"
            convert_odt_to_docx(odt, str(tmp_docx))
            temp_docxs.append(str(tmp_docx))

        merged_docx = temp_dir_path / "merged_temp.docx"
        merge_docx(temp_docxs, str(merged_docx))
        final_odt = convert_docx_to_odt(str(merged_docx), output_odt_path)
        return final_odt


def merge_images(image_paths: List[str], output_pdf_path: str) -> str:
    """依列表順序將多張圖片合併成一個多頁 PDF"""
    return convert_images_to_pdf(image_paths, output_pdf_path)
