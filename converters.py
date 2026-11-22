import base64
import csv
import html
import io
import logging
import os
import re
import shutil
import statistics
import subprocess
import tempfile
import zipfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

import pymupdf as fitz  # PyMuPDF
from PIL import Image
import img2pdf
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

from docx import Document
from docx.enum.section import WD_ORIENTATION
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.table import Table
from docx.text.paragraph import Paragraph

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

import markdown
from bs4 import BeautifulSoup
from pdf2docx import Converter

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    PageBreak,
    Paragraph as RLParagraph,
    SimpleDocTemplate,
    Spacer,
    Table as RLTable,
    TableStyle,
)

logger = logging.getLogger("velmo")
logging.basicConfig(level=logging.INFO)

SUPPORTED_FORMATS = {
    "pdf", "docx", "xlsx", "pptx", "csv", "txt", "md", "html", "png", "jpg", "jpeg", "webp"
}

# --- Reusable Lazy Singleton OCR Engine ---

_ocr_engine = None

def get_ocr_engine():
    """Returns the lazy-loaded singleton instance of RapidOCR."""
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        logger.info("Initializing lazy singleton RapidOCR engine...")
        _ocr_engine = RapidOCR()
    return _ocr_engine


# --- LibreOffice Detection ---

def get_libreoffice_command() -> Optional[str]:
    """Finds the soffice/libreoffice executable path if installed on the system."""
    cmd = shutil.which("soffice") or shutil.which("libreoffice")
    if cmd:
        return cmd

    common_paths = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for p in common_paths:
        if os.path.isfile(p):
            return p
    return None


LIBREOFFICE_PATH = get_libreoffice_command()
if LIBREOFFICE_PATH:
    logger.info(f"LibreOffice detected at: {LIBREOFFICE_PATH}")
else:
    logger.info("LibreOffice NOT detected. Pure-Python ReportLab fallback will be used.")


# --- Common Utilities ---

def sanitize_filename(filename: str, fallback: str = "converted") -> str:
    """
    Sanitizes a filename by:
    - Stripping path components (preventing directory traversal).
    - Removing null bytes and control characters.
    - Replacing unsafe/illegal characters (<>:"/\\|?*) with underscores.
    - Stripping leading and trailing periods and whitespace.
    """
    if not filename:
        return fallback

    clean = filename.replace("\\", "/")
    base = os.path.basename(clean)
    base = re.sub(r'[\x00-\x1f\x7f]', '', base)
    base = re.sub(r'[<>:"/\\|?*]', '_', base)
    base = base.strip(". ")
    return base or fallback


def sanitize_base_name(filename: str, fallback: str = "converted") -> str:
    """
    Extracts and sanitizes the stem/base name of a filename.
    """
    safe_name = sanitize_filename(filename, fallback=fallback)
    stem = Path(safe_name).stem
    stem = re.sub(r'[<>:"/\\|?*]', '_', stem).strip(". ")
    return stem or fallback


def sanitize_sheet_title(name: str, used_names: Set[str]) -> str:
    """Ensures an Excel worksheet title is valid, <= 31 chars, and unique."""
    cleaned = re.sub(r'[\x00-\x1f\\/*?\[\]:]', '_', name).strip()
    if not cleaned:
        cleaned = "Sheet"
    base = cleaned[:31]
    candidate = base
    counter = 1
    while candidate.lower() in used_names:
        suffix = f"_{counter}"
        candidate = f"{base[:31 - len(suffix)]}{suffix}"
        counter += 1
    used_names.add(candidate.lower())
    return candidate


def parse_numeric(text: str):
    """Parses numeric strings into int or float when appropriate."""
    if not text:
        return ""
    stripped = text.strip()
    try:
        val = int(stripped)
        if stripped == "0" or not stripped.startswith("0"):
            return val
    except ValueError:
        pass
    try:
        val = float(stripped)
        return val
    except ValueError:
        pass
    return text


def extract_fill_hex(cell) -> Optional[str]:
    """Extracts a 6-character hex color code from an openpyxl cell fill."""
    if not cell or not cell.fill or cell.fill.fill_type in (None, "none"):
        return None
    color = cell.fill.start_color or cell.fill.fgColor
    if not color:
        return None
    rgb = getattr(color, "rgb", None)
    if isinstance(rgb, str) and rgb:
        hex_code = rgb[-6:] if len(rgb) >= 6 else rgb
        if len(hex_code) == 6:
            if hex_code.upper() in ("000000", "FFFFFF") and getattr(color, "theme", None) is None:
                if cell.fill.fill_type == "solid":
                    return hex_code.upper()
                return None
            return hex_code.upper()
    return None


def set_cell_background(cell, hex_color: str):
    """Sets the background fill color of a python-docx table cell."""
    try:
        tcPr = cell._tc.get_or_add_tcPr()
        for child in list(tcPr):
            if child.tag.endswith("shd"):
                tcPr.remove(child)
        shd_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
        tcPr.append(parse_xml(shd_xml))
    except Exception:
        pass


def validate_input_file(filename: str, raw_bytes: bytes) -> None:
    """
    Validates that an uploaded file is not empty, has a supported format extension,
    and is structurally valid (not corrupted or garbage bytes).
    Raises ValueError on any violation so callers return clean HTTP 400 errors.
    """
    if not raw_bytes:
        raise ValueError(f"Uploaded file '{filename}' is empty.")

    ext = Path(filename).suffix.lower().lstrip(".").replace("jpeg", "jpg")
    if not ext:
        raise ValueError(f"Uploaded file '{filename}' has no file extension.")

    if ext not in SUPPORTED_FORMATS:
        raise ValueError(
            f"Unsupported file format '{ext}' in '{filename}'. Supported formats: {', '.join(sorted(SUPPORTED_FORMATS))}"
        )

    # Check structural integrity per format
    if ext == "pdf":
        try:
            doc = fitz.open(stream=raw_bytes, filetype="pdf")
            if len(doc) == 0:
                doc.close()
                raise ValueError("PDF document has 0 pages or is empty.")
            doc.close()
        except Exception as e:
            raise ValueError(f"Invalid or corrupted PDF file '{filename}': {e}")

    elif ext == "docx":
        try:
            Document(io.BytesIO(raw_bytes))
        except Exception as e:
            raise ValueError(f"Invalid or corrupted DOCX file '{filename}': {e}")

    elif ext == "xlsx":
        try:
            wb = openpyxl.load_workbook(io.BytesIO(raw_bytes), data_only=True)
            wb.close()
        except Exception as e:
            raise ValueError(f"Invalid or corrupted XLSX file '{filename}': {e}")

    elif ext == "pptx":
        try:
            Presentation(io.BytesIO(raw_bytes))
        except Exception as e:
            raise ValueError(f"Invalid or corrupted PPTX file '{filename}': {e}")

    elif ext in ("png", "jpg", "webp"):
        try:
            with Image.open(io.BytesIO(raw_bytes)) as im:
                im.verify()
        except Exception as e:
            raise ValueError(f"Invalid or corrupted image file '{filename}': {e}")

    elif ext in ("csv", "txt", "md", "html"):
        try:
            raw_bytes.decode("utf-8")
        except UnicodeDecodeError:
            try:
                raw_bytes.decode("latin-1")
            except Exception as e:
                raise ValueError(f"Invalid text encoding in file '{filename}': {e}")


def process_image_for_pdf(image_bytes: bytes, filename: str) -> bytes:
    """Prepares an image for lossless embedding via img2pdf, flattening any alpha channel."""
    if not image_bytes:
        raise ValueError(f"File '{filename}' is empty.")

    try:
        with Image.open(io.BytesIO(image_bytes)) as im:
            format_name = (im.format or "").upper()
            mode = im.mode
            has_alpha = mode in ("RGBA", "LA") or (mode == "P" and "transparency" in im.info)

            if format_name == "WEBP" or has_alpha or mode not in ("RGB", "L", "1"):
                if has_alpha:
                    im_rgba = im.convert("RGBA")
                    background = Image.new("RGBA", im_rgba.size, (255, 255, 255, 255))
                    composite = Image.alpha_composite(background, im_rgba)
                    rgb_img = composite.convert("RGB")
                else:
                    rgb_img = im.convert("RGB")

                out_buf = io.BytesIO()
                save_kwargs = {"format": "PNG"}
                if "dpi" in im.info:
                    save_kwargs["dpi"] = im.info["dpi"]
                rgb_img.save(out_buf, **save_kwargs)
                return out_buf.getvalue()

            try:
                img2pdf.convert([image_bytes])
                return image_bytes
            except Exception:
                rgb_img = im.convert("RGB")
                out_buf = io.BytesIO()
                save_kwargs = {"format": "PNG"}
                if "dpi" in im.info:
                    save_kwargs["dpi"] = im.info["dpi"]
                rgb_img.save(out_buf, **save_kwargs)
                return out_buf.getvalue()
    except Exception as e:
        raise ValueError(f"Cannot process image '{filename}': {str(e)}")


def extract_ocr_items(image_bytes: bytes) -> List[dict]:
    """Runs RapidOCR and returns normalized parsed box items."""
    ocr = get_ocr_engine()
    res, _ = ocr(image_bytes)
    if not res:
        return []

    items = []
    for box, text, score in res:
        text_clean = text.strip()
        if not text_clean:
            continue
        xs = [pt[0] for pt in box]
        ys = [pt[1] for pt in box]
        x_min, x_max = min(xs), max(xs)
        y_min, y_max = min(ys), max(ys)
        items.append({
            'x_min': x_min,
            'x_max': x_max,
            'y_min': y_min,
            'y_max': y_max,
            'x_center': (x_min + x_max) / 2.0,
            'y_center': (y_min + y_max) / 2.0,
            'height': max(y_max - y_min, 1.0),
            'width': max(x_max - x_min, 1.0),
            'text': text_clean,
            'score': score
        })
    return items


# --- Image Formats (PNG ↔ JPG ↔ WebP) ---

def parse_page_range(range_str: Optional[str], total_pages: int) -> List[int]:
    """
    Parses a 1-indexed page range string (e.g. '1-3, 5') into 0-indexed page indices.
    Ignores out-of-range pages gracefully if at least one valid page remains.
    Raises ValueError for malformed syntax (e.g. 'abc') or if no valid pages exist.
    """
    if not range_str or not range_str.strip():
        return list(range(total_pages))

    s = range_str.strip()
    if not re.match(r'^[0-9\s,\-]+$', s):
        raise ValueError(f"Invalid page_range syntax: '{range_str}'. Use comma-separated numbers and ranges (e.g. '1-3, 5').")

    selected: List[int] = []
    seen = set()
    parts = [p.strip() for p in s.split(",") if p.strip()]
    if not parts:
        raise ValueError(f"Invalid page_range syntax: '{range_str}'.")

    for part in parts:
        if "-" in part:
            bounds = [b.strip() for b in part.split("-")]
            if len(bounds) != 2 or not bounds[0].isdigit() or not bounds[1].isdigit():
                raise ValueError(f"Invalid range '{part}' in page_range.")
            start, end = int(bounds[0]), int(bounds[1])
            if start < 1 or end < 1 or start > end:
                raise ValueError(f"Invalid range '{part}'. Start must be >= 1 and <= end.")
            for p in range(start, end + 1):
                if 1 <= p <= total_pages:
                    idx = p - 1
                    if idx not in seen:
                        seen.add(idx)
                        selected.append(idx)
        else:
            if not part.isdigit():
                raise ValueError(f"Invalid page number '{part}' in page_range.")
            p = int(part)
            if p < 1:
                raise ValueError(f"Page numbers must be >= 1, got {p}.")
            if 1 <= p <= total_pages:
                idx = p - 1
                if idx not in seen:
                    seen.add(idx)
                    selected.append(idx)

    if not selected:
        raise ValueError(f"Requested page range '{range_str}' contains no valid pages for this {total_pages}-page document.")

    return selected


def filter_pdf_pages(pdf_bytes: bytes, page_range_str: str) -> bytes:
    """Filters a PDF byte string to retain only the pages specified in page_range_str."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    total_pages = len(doc)
    if total_pages == 0:
        doc.close()
        raise ValueError("PDF document has no pages.")

    selected_indices = parse_page_range(page_range_str, total_pages)
    if selected_indices != list(range(total_pages)):
        doc.select(selected_indices)
        out_bytes = doc.tobytes()
    else:
        out_bytes = pdf_bytes
    doc.close()
    return out_bytes


def convert_image_format(
    image_bytes: bytes,
    target_fmt: str,
    base_name: str,
    image_quality: int = 95
) -> Tuple[bytes, str, str, int]:
    """Converts directly between PNG, JPG, and WebP images."""
    target_fmt = target_fmt.lower().replace("jpeg", "jpg")
    with Image.open(io.BytesIO(image_bytes)) as im:
        buf = io.BytesIO()
        if target_fmt == "jpg":
            # Flatten transparency onto white background at quality
            if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
                im_rgba = im.convert("RGBA")
                bg = Image.new("RGBA", im_rgba.size, (255, 255, 255, 255))
                comp = Image.alpha_composite(bg, im_rgba).convert("RGB")
            else:
                comp = im.convert("RGB")
            comp.save(buf, format="JPEG", quality=image_quality)
            return buf.getvalue(), "image/jpeg", f"{base_name}.jpg", 1

        elif target_fmt == "png":
            im.save(buf, format="PNG")
            return buf.getvalue(), "image/png", f"{base_name}.png", 1

        elif target_fmt == "webp":
            im.save(buf, format="WEBP", quality=image_quality)
            return buf.getvalue(), "image/webp", f"{base_name}.webp", 1

        else:
            raise ValueError(f"Unsupported image target format: {target_fmt}")


def pdf_to_images(
    pdf_bytes: bytes,
    target_fmt: str,
    base_name: str,
    dpi: int = 300,
    image_quality: int = 95
) -> Tuple[bytes, str, str, int]:
    """Renders PDF pages to PNG, JPG, or WebP at custom DPI and quality."""
    target_fmt = target_fmt.lower().replace("jpeg", "jpg")
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page_count = len(doc)
    if page_count == 0:
        doc.close()
        raise ValueError("PDF contains no pages.")

    rendered_images: List[Tuple[str, bytes]] = []
    for idx, page in enumerate(doc, start=1):
        pix = page.get_pixmap(dpi=dpi, alpha=(target_fmt != "jpg"))
        if target_fmt == "png":
            img_bytes = pix.tobytes("png")
            fname = f"{base_name}_page_{idx:02d}.png"
        elif target_fmt == "jpg":
            img_bytes = pix.tobytes("jpg", jpg_quality=image_quality)
            fname = f"{base_name}_page_{idx:02d}.jpg"
        elif target_fmt == "webp":
            buf = io.BytesIO()
            pil_img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            pil_img.save(buf, format="WEBP", quality=image_quality)
            img_bytes = buf.getvalue()
            fname = f"{base_name}_page_{idx:02d}.webp"
        else:
            raise ValueError(f"Unsupported image format: {target_fmt}")
        rendered_images.append((fname, img_bytes))

    doc.close()

    if page_count == 1:
        fname, data = rendered_images[0]
        mime_map = {"png": "image/png", "jpg": "image/jpeg", "webp": "image/webp"}
        single_fname = f"{base_name}.{target_fmt}"
        return data, mime_map.get(target_fmt, "application/octet-stream"), single_fname, 1
    else:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for fname, data in rendered_images:
                zf.writestr(fname, data)
        return zip_buf.getvalue(), "application/zip", f"{base_name}_{target_fmt}_pages.zip", page_count


def pdf_to_png(pdf_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    return pdf_to_images(pdf_bytes, "png", base_name)


def images_to_pdf(image_list: List[Tuple[str, bytes]], base_name: str) -> Tuple[bytes, str, str, int]:
    """Combines one or multiple images losslessly into a single PDF."""
    if not image_list:
        raise ValueError("No images provided.")
    prepared = [process_image_for_pdf(raw, name) for name, raw in image_list]
    pdf_bytes = img2pdf.convert(prepared)
    out_name = f"{base_name}.pdf" if len(image_list) == 1 else f"{base_name}_combined.pdf"
    return pdf_bytes, "application/pdf", out_name, len(image_list)


# --- Word & Excel Core Converters ---

def xlsx_to_docx(xlsx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    wb_data = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    wb_raw = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=False)
    doc = Document()

    max_cols_overall = 0
    for ws in wb_data.worksheets:
        for row in ws.iter_rows(values_only=False):
            non_empty_indices = [idx for idx, cell in enumerate(row) if cell.value is not None and str(cell.value).strip() != ""]
            if non_empty_indices:
                max_cols_overall = max(max_cols_overall, max(non_empty_indices) + 1)

    if max_cols_overall > 6:
        section = doc.sections[0]
        section.orientation = WD_ORIENTATION.LANDSCAPE
        if section.page_width < section.page_height:
            section.page_width, section.page_height = section.page_height, section.page_width

    for ws_data in wb_data.worksheets:
        ws_raw = wb_raw[ws_data.title]
        doc.add_heading(ws_data.title, level=1)

        rows_data = []
        max_col_in_sheet = 0

        for row in ws_data.iter_rows(values_only=False):
            row_cells = []
            row_has_value = False
            for cell in row:
                val = cell.value
                raw_cell = ws_raw.cell(row=cell.row, column=cell.column)
                if val is None and raw_cell.value is not None:
                    val = raw_cell.value
                if val is not None and str(val).strip() != "":
                    row_has_value = True
                row_cells.append((cell, raw_cell, val))

            if row_has_value:
                last_idx = max(i for i, (_, _, v) in enumerate(row_cells) if v is not None and str(v).strip() != "")
                max_col_in_sheet = max(max_col_in_sheet, last_idx + 1)
                rows_data.append(row_cells)

        if not rows_data or max_col_in_sheet == 0:
            doc.add_paragraph("(Empty sheet)")
            continue

        num_rows = len(rows_data)
        num_cols = max_col_in_sheet
        table = doc.add_table(rows=num_rows, cols=num_cols)
        table.style = 'Table Grid'

        for r_i, row_cells in enumerate(rows_data):
            for c_i in range(num_cols):
                table_cell = table.cell(r_i, c_i)
                cell_data, raw_cell, val = row_cells[c_i] if c_i < len(row_cells) else (None, None, None)
                val_str = "" if val is None else str(val)
                p = table_cell.paragraphs[0]
                run = p.add_run(val_str)

                if cell_data and cell_data.font and cell_data.font.bold:
                    run.bold = True
                if cell_data:
                    fill_hex = extract_fill_hex(cell_data)
                    if fill_hex:
                        set_cell_background(table_cell, fill_hex)

        doc.add_paragraph()

    out_io = io.BytesIO()
    doc.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", len(wb_data.worksheets)


def docx_to_xlsx(docx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    doc = Document(io.BytesIO(docx_bytes))
    wb = openpyxl.Workbook()
    default_sheet = wb.active

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1'),
    )
    header_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    used_sheet_titles: Set[str] = set()
    table_index = 0
    last_heading_text: Optional[str] = None
    document_paragraphs = []

    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            p = Paragraph(child, doc)
            text = p.text.strip()
            if not text:
                continue
            is_heading = (
                p.style.name.startswith("Heading") or
                (any(r.bold for r in p.runs) and len(text) < 80)
            )
            document_paragraphs.append((p.style.name, text, is_heading))
            if is_heading:
                last_heading_text = text

        elif child.tag.endswith('tbl'):
            table = Table(child, doc)
            table_index += 1
            raw_title = last_heading_text if last_heading_text else f"Table {table_index}"
            sheet_title = sanitize_sheet_title(raw_title, used_sheet_titles)
            last_heading_text = None

            ws = wb.create_sheet(title=sheet_title)
            for r_idx, row in enumerate(table.rows, start=1):
                is_header_row = (r_idx == 1)
                for c_idx, cell in enumerate(row.cells, start=1):
                    c_text = cell.text.strip()
                    parsed_val = parse_numeric(c_text)
                    excel_cell = ws.cell(row=r_idx, column=c_idx, value=parsed_val)
                    excel_cell.border = thin_border

                    is_cell_bold = is_header_row or any(run.bold for p in cell.paragraphs for run in p.runs)
                    if is_cell_bold:
                        excel_cell.font = Font(bold=True)
                        if is_header_row:
                            excel_cell.fill = header_fill

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for c in col:
                    if c.value is not None:
                        lines = str(c.value).split("\n")
                        max_len = max(max_len, max(len(l) for l in lines))
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)

    text_sheet_title = sanitize_sheet_title("Document Text", used_sheet_titles)
    ws_text = wb.create_sheet(title=text_sheet_title, index=0)
    ws_text.cell(row=1, column=1, value="Type").font = Font(bold=True)
    ws_text.cell(row=1, column=1).fill = header_fill
    ws_text.cell(row=1, column=1).border = thin_border
    ws_text.cell(row=1, column=2, value="Content").font = Font(bold=True)
    ws_text.cell(row=1, column=2).fill = header_fill
    ws_text.cell(row=1, column=2).border = thin_border

    if document_paragraphs:
        for idx, (p_style, p_text, is_heading) in enumerate(document_paragraphs, start=2):
            c1 = ws_text.cell(row=idx, column=1, value=p_style)
            c2 = ws_text.cell(row=idx, column=2, value=p_text)
            c1.border = thin_border
            c2.border = thin_border
            if is_heading:
                c1.font = Font(bold=True)
                c2.font = Font(bold=True)
    else:
        c1 = ws_text.cell(row=2, column=1, value="Info")
        c2 = ws_text.cell(row=2, column=2, value="No standalone text found outside tables.")
        c1.border = thin_border
        c2.border = thin_border

    for col in ws_text.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for c in col:
            if c.value is not None:
                lines = str(c.value).split("\n")
                max_len = max(max_len, max(len(l) for l in lines))
        ws_text.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 65)

    if default_sheet in wb.worksheets and len(wb.worksheets) > 1:
        wb.remove(default_sheet)

    out_io = io.BytesIO()
    wb.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{base_name}.xlsx", table_index


def pdf_to_docx(pdf_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    base_name = sanitize_base_name(base_name)
    with tempfile.TemporaryDirectory() as tmp_dir:
        temp_pdf_path = os.path.join(tmp_dir, "input.pdf")
        temp_docx_path = os.path.join(tmp_dir, "output.docx")
        with open(temp_pdf_path, "wb") as f:
            f.write(pdf_bytes)
        cv = None
        try:
            cv = Converter(temp_pdf_path)
            cv.convert(temp_docx_path, start=0, end=None)
        finally:
            if cv:
                cv.close()
        with open(temp_docx_path, "rb") as f:
            docx_bytes = f.read()
    return docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", 1


def pdf_to_xlsx(pdf_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    wb = openpyxl.Workbook()
    default_sheet = wb.active

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1'),
    )
    header_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    used_sheet_titles: Set[str] = set()
    table_count = 0
    document_text_blocks = []

    for p_no, page in enumerate(doc, start=1):
        tabs = page.find_tables()
        table_bboxes = []

        for t_no, tab in enumerate(tabs, start=1):
            extracted = tab.extract()
            if not extracted:
                continue

            table_count += 1
            table_bboxes.append(tab.bbox)
            sheet_title = sanitize_sheet_title(f"Page_{p_no}_Table_{t_no}", used_sheet_titles)
            ws = wb.create_sheet(title=sheet_title)

            for r_idx, row in enumerate(extracted, start=1):
                is_header = (r_idx == 1)
                for c_idx, cell_val in enumerate(row, start=1):
                    c_str = "" if cell_val is None else str(cell_val).strip()
                    parsed = parse_numeric(c_str)
                    excel_cell = ws.cell(row=r_idx, column=c_idx, value=parsed)
                    excel_cell.border = thin_border
                    if is_header:
                        excel_cell.font = Font(bold=True)
                        excel_cell.fill = header_fill

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for c in col:
                    if c.value is not None:
                        lines = str(c.value).split("\n")
                        max_len = max(max_len, max(len(l) for l in lines))
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)

        blocks = page.get_text("blocks")
        for b in blocks:
            b_rect = fitz.Rect(b[0], b[1], b[2], b[3])
            b_text = b[4].strip()
            if not b_text:
                continue
            in_table = any(b_rect.intersects(fitz.Rect(t_box)) for t_box in table_bboxes)
            if not in_table or table_count == 0:
                document_text_blocks.append((p_no, b_text))

    text_sheet_title = sanitize_sheet_title("Document Text", used_sheet_titles)
    ws_text = wb.create_sheet(title=text_sheet_title, index=0)
    ws_text.cell(row=1, column=1, value="Page").font = Font(bold=True)
    ws_text.cell(row=1, column=1).fill = header_fill
    ws_text.cell(row=1, column=1).border = thin_border
    ws_text.cell(row=1, column=2, value="Content").font = Font(bold=True)
    ws_text.cell(row=1, column=2).fill = header_fill
    ws_text.cell(row=1, column=2).border = thin_border

    if document_text_blocks:
        for idx, (p_no, txt) in enumerate(document_text_blocks, start=2):
            c1 = ws_text.cell(row=idx, column=1, value=p_no)
            c2 = ws_text.cell(row=idx, column=2, value=txt)
            c1.border = thin_border
            c2.border = thin_border
    else:
        c1 = ws_text.cell(row=2, column=1, value=1)
        c2 = ws_text.cell(row=2, column=2, value="No standalone text found outside tables.")
        c1.border = thin_border
        c2.border = thin_border

    for col in ws_text.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for c in col:
            if c.value is not None:
                lines = str(c.value).split("\n")
                max_len = max(max_len, max(len(l) for l in lines))
        ws_text.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 65)

    if default_sheet in wb.worksheets and len(wb.worksheets) > 1:
        wb.remove(default_sheet)

    doc.close()
    out_io = io.BytesIO()
    wb.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{base_name}.xlsx", table_count


def convert_via_libreoffice(input_bytes: bytes, suffix: str) -> Optional[bytes]:
    soffice = get_libreoffice_command()
    if not soffice:
        return None

    with tempfile.TemporaryDirectory() as tmp_dir:
        input_path = os.path.join(tmp_dir, f"input{suffix}")
        with open(input_path, "wb") as f:
            f.write(input_bytes)

        try:
            cmd = [soffice, "--headless", "--convert-to", "pdf", "--outdir", tmp_dir, input_path]
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=60)
            output_pdf = os.path.join(tmp_dir, "input.pdf")
            if os.path.exists(output_pdf):
                with open(output_pdf, "rb") as f:
                    return f.read()
        except Exception as e:
            logger.warning(f"LibreOffice conversion failed: {e}. Falling back to ReportLab.")
            return None
    return None


def docx_to_pdf(docx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    pdf_bytes = convert_via_libreoffice(docx_bytes, ".docx")
    if not pdf_bytes:
        pdf_bytes = _docx_to_pdf_reportlab(docx_bytes)
    return pdf_bytes, "application/pdf", f"{base_name}.pdf", 1


def _docx_to_pdf_reportlab(docx_bytes: bytes) -> bytes:
    doc = Document(io.BytesIO(docx_bytes))
    buffer = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buffer, pagesize=letter, leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)

    styles = getSampleStyleSheet()
    h1_style = ParagraphStyle('DocxH1', parent=styles['Heading1'], fontSize=16, leading=20, spaceAfter=8, textColor=colors.HexColor('#0F172A'))
    h2_style = ParagraphStyle('DocxH2', parent=styles['Heading2'], fontSize=13, leading=16, spaceAfter=6, textColor=colors.HexColor('#1E293B'))
    body_style = ParagraphStyle('DocxBody', parent=styles['Normal'], fontSize=10, leading=14, spaceAfter=6, textColor=colors.HexColor('#334155'))
    cell_style = ParagraphStyle('DocxCell', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#0F172A'))
    cell_bold_style = ParagraphStyle('DocxCellBold', parent=styles['Normal'], fontSize=9, leading=12, fontName='Helvetica-Bold', textColor=colors.HexColor('#0F172A'))

    story = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            p = Paragraph(child, doc)
            text = p.text.strip()
            if not text:
                continue
            if p.runs and any(r.bold for r in p.runs):
                run_parts = [f"<b>{html.escape(r.text)}</b>" if r.bold else html.escape(r.text) for r in p.runs]
                formatted_text = "".join(run_parts)
            else:
                formatted_text = html.escape(text)

            if p.style.name.startswith("Heading 1"):
                story.append(RLParagraph(formatted_text, h1_style))
            elif p.style.name.startswith("Heading"):
                story.append(RLParagraph(formatted_text, h2_style))
            else:
                story.append(RLParagraph(formatted_text, body_style))

        elif child.tag.endswith('tbl'):
            table = Table(child, doc)
            table_data = []
            for r_idx, row in enumerate(table.rows):
                row_data = []
                for cell in row.cells:
                    c_text = html.escape(cell.text.strip())
                    is_bold = (r_idx == 0) or any(r.bold for p in cell.paragraphs for r in p.runs)
                    st = cell_bold_style if is_bold else cell_style
                    row_data.append(RLParagraph(c_text, st))
                table_data.append(row_data)

            if table_data:
                rl_table = RLTable(table_data)
                rl_table.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F1F5F9')),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 0), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                ]))
                story.append(rl_table)
                story.append(Spacer(1, 10))

    if not story:
        story.append(RLParagraph("Empty Document", body_style))

    pdf_doc.build(story)
    return buffer.getvalue()


def xlsx_to_pdf(xlsx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    pdf_bytes = convert_via_libreoffice(xlsx_bytes, ".xlsx")
    if not pdf_bytes:
        pdf_bytes = _xlsx_to_pdf_reportlab(xlsx_bytes)
    return pdf_bytes, "application/pdf", f"{base_name}.pdf", 1


def _xlsx_to_pdf_reportlab(xlsx_bytes: bytes) -> bytes:
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    buffer = io.BytesIO()

    max_cols = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            non_empty = [idx for idx, val in enumerate(row) if val is not None and str(val).strip() != ""]
            if non_empty:
                max_cols = max(max_cols, max(non_empty) + 1)

    pagesize = landscape(letter) if max_cols > 6 else letter
    pdf_doc = SimpleDocTemplate(buffer, pagesize=pagesize, leftMargin=30, rightMargin=30, topMargin=30, bottomMargin=30)

    styles = getSampleStyleSheet()
    h1_style = ParagraphStyle('XlsxH1', parent=styles['Heading1'], fontSize=15, leading=18, spaceAfter=8, textColor=colors.HexColor('#0F172A'))
    cell_style = ParagraphStyle('XlsxCell', parent=styles['Normal'], fontSize=8, leading=10, textColor=colors.HexColor('#0F172A'))
    header_style = ParagraphStyle('XlsxHdr', parent=styles['Normal'], fontSize=8, leading=10, fontName='Helvetica-Bold', textColor=colors.HexColor('#0F172A'))

    story = []
    for s_idx, ws in enumerate(wb.worksheets):
        if s_idx > 0:
            story.append(PageBreak())

        story.append(RLParagraph(html.escape(ws.title), h1_style))
        table_data = []
        for r_idx, row in enumerate(ws.iter_rows(values_only=True)):
            if any(v is not None and str(v).strip() != "" for v in row):
                last_idx = max(i for i, v in enumerate(row) if v is not None and str(v).strip() != "")
                row_slice = row[:last_idx + 1]
                r_data = [RLParagraph(html.escape("" if val is None else str(val)), header_style if r_idx == 0 else cell_style) for val in row_slice]
                table_data.append(r_data)

        if table_data:
            rl_table = RLTable(table_data)
            rl_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F1F5F9')),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('TOPPADDING', (0, 0), (-1, -1), 3),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
            ]))
            story.append(rl_table)
        else:
            story.append(RLParagraph("(Empty sheet)", cell_style))

        story.append(Spacer(1, 10))

    if not story:
        story.append(RLParagraph("Empty Workbook", cell_style))

    pdf_doc.build(story)
    return buffer.getvalue()


def docx_to_png(docx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    pdf_bytes, _, _, _ = docx_to_pdf(docx_bytes, base_name)
    return pdf_to_png(pdf_bytes, base_name)


def xlsx_to_png(xlsx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    pdf_bytes, _, _, _ = xlsx_to_pdf(xlsx_bytes, base_name)
    return pdf_to_png(pdf_bytes, base_name)


# --- OCR Converters ---

def images_to_docx(image_list: List[Tuple[str, bytes]], base_name: str) -> Tuple[bytes, str, str, int]:
    if not image_list:
        raise ValueError("No images provided.")

    doc = Document()
    for img_idx, (filename, img_bytes) in enumerate(image_list):
        if img_idx > 0:
            doc.add_page_break()

        items = extract_ocr_items(img_bytes)
        if not items:
            doc.add_paragraph(f"(No text detected on image: {filename})")
            continue

        median_h = statistics.median([it['height'] for it in items])
        items_sorted = sorted(items, key=lambda it: it['y_center'])
        lines: List[List[dict]] = []

        for item in items_sorted:
            placed = False
            for line in lines:
                line_y = sum(it['y_center'] for it in line) / len(line)
                line_h = sum(it['height'] for it in line) / len(line)
                if abs(item['y_center'] - line_y) <= max(line_h, item['height']) * 0.55:
                    line.append(item)
                    placed = True
                    break
            if not placed:
                lines.append([item])

        lines.sort(key=lambda ln: sum(it['y_center'] for it in ln) / len(ln))

        for line in lines:
            line.sort(key=lambda it: it['x_min'])
            line_text = " ".join(it['text'] for it in line)
            avg_h = sum(it['height'] for it in line) / len(line)

            if avg_h >= 1.35 * median_h and len(line_text) < 100:
                doc.add_heading(line_text, level=1 if avg_h >= 1.6 * median_h else 2)
            else:
                doc.add_paragraph(line_text)

    out_io = io.BytesIO()
    doc.save(out_io)
    output_filename = f"{base_name}.docx" if len(image_list) == 1 else f"{base_name}_ocr.docx"
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", output_filename, len(image_list)


def images_to_xlsx(image_list: List[Tuple[str, bytes]], base_name: str) -> Tuple[bytes, str, str, int]:
    if not image_list:
        raise ValueError("No images provided.")

    wb = openpyxl.Workbook()
    default_sheet = wb.active

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1'),
    )
    header_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    used_sheet_titles: Set[str] = set()

    for img_idx, (filename, img_bytes) in enumerate(image_list, start=1):
        raw_title = Path(filename).stem or f"Image {img_idx}"
        sheet_title = sanitize_sheet_title(raw_title, used_sheet_titles)
        ws = wb.create_sheet(title=sheet_title)

        items = extract_ocr_items(img_bytes)
        if not items:
            ws.cell(row=1, column=1, value="(No text detected on image)")
            continue

        items_sorted = sorted(items, key=lambda it: it['y_center'])
        rows: List[List[dict]] = []
        for item in items_sorted:
            placed = False
            for r in rows:
                r_y = sum(it['y_center'] for it in r) / len(r)
                r_h = sum(it['height'] for it in r) / len(r)
                if abs(item['y_center'] - r_y) <= max(r_h, item['height']) * 0.6:
                    r.append(item)
                    placed = True
                    break
            if not placed:
                rows.append([item])

        rows.sort(key=lambda r: sum(it['y_center'] for it in r) / len(r))

        all_x = sorted([it['x_center'] for it in items])
        col_clusters: List[List[float]] = []
        col_tol = 50.0
        for r in rows:
            if len(r) > 1:
                r_s = sorted(r, key=lambda it: it['x_center'])
                for k in range(len(r_s) - 1):
                    dist = r_s[k+1]['x_center'] - r_s[k]['x_center']
                    if 15 < dist < col_tol * 2:
                        col_tol = min(col_tol, dist * 0.45)

        for x in all_x:
            if not col_clusters:
                col_clusters.append([x])
            else:
                closest_idx = min(range(len(col_clusters)), key=lambda i: abs(sum(col_clusters[i])/len(col_clusters[i]) - x))
                closest_avg = sum(col_clusters[closest_idx]) / len(col_clusters[closest_idx])
                if abs(x - closest_avg) <= col_tol:
                    col_clusters[closest_idx].append(x)
                else:
                    col_clusters.append([x])

        sorted_cols = sorted([sum(c)/len(c) for c in col_clusters])

        for r_idx, r in enumerate(rows, start=1):
            is_header_row = (r_idx == 1)
            row_cells = [""] * len(sorted_cols)
            r.sort(key=lambda it: it['x_min'])

            for it in r:
                c_idx = min(range(len(sorted_cols)), key=lambda i: abs(sorted_cols[i] - it['x_center']))
                if row_cells[c_idx]:
                    row_cells[c_idx] += " " + it['text']
                else:
                    row_cells[c_idx] = it['text']

            for c_idx, val_str in enumerate(row_cells, start=1):
                parsed = parse_numeric(val_str)
                cell = ws.cell(row=r_idx, column=c_idx, value=parsed)
                cell.border = thin_border
                if is_header_row:
                    cell.font = Font(bold=True)
                    cell.fill = header_fill

        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for c in col:
                if c.value is not None:
                    max_len = max(max_len, len(str(c.value)))
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)

    if default_sheet in wb.worksheets and len(wb.worksheets) > 1:
        wb.remove(default_sheet)

    out_io = io.BytesIO()
    wb.save(out_io)
    output_filename = f"{base_name}.xlsx" if len(image_list) == 1 else f"{base_name}_ocr.xlsx"
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", output_filename, len(image_list)


# --- PowerPoint (.pptx) Converters ---

def _extract_slide_title(slide, default_title: str) -> Tuple[str, Any, bool]:
    """
    Extracts the most appropriate title for a slide.
    Checks slide.shapes.title first; if not found or empty, searches for the topmost text shape.
    Returns: (title_text, title_shape, skip_first_paragraph_only)
    """
    if slide.shapes.title and slide.shapes.title.has_text_frame and slide.shapes.title.text_frame.text.strip():
        return slide.shapes.title.text_frame.text.strip(), slide.shapes.title, False

    text_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
    if text_shapes:
        text_shapes.sort(key=lambda s: getattr(s, 'top', 0))
        top_s = text_shapes[0]
        non_empty_paras = [p for p in top_s.text_frame.paragraphs if p.text.strip()]
        if non_empty_paras:
            candidate = non_empty_paras[0].text.strip()
            if len(candidate) < 150:
                return candidate, top_s, (len(non_empty_paras) > 1)
    return default_title, None, False


def pptx_to_docx(pptx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Converts PowerPoint slides into an editable Word document with headings, bullets, and tables."""
    prs = Presentation(io.BytesIO(pptx_bytes))
    doc = Document()
    slide_count = len(prs.slides)

    for idx, slide in enumerate(prs.slides, start=1):
        title_text, title_shape, skip_first_p = _extract_slide_title(slide, f"Slide {idx}")
        doc.add_heading(title_text, level=1)

        for shape in slide.shapes:
            if shape == title_shape and not skip_first_p:
                continue
            if shape.has_text_frame:
                paragraphs = shape.text_frame.paragraphs
                if shape == title_shape and skip_first_p:
                    paragraphs = paragraphs[1:]
                for p in paragraphs:
                    t = p.text.strip()
                    if t:
                        doc.add_paragraph(t, style='List Bullet' if p.level > 0 else 'Normal')
            elif shape.has_table:
                tbl = shape.table
                doc_tbl = doc.add_table(rows=len(tbl.rows), cols=len(tbl.columns))
                doc_tbl.style = 'Table Grid'
                for r_i, row in enumerate(tbl.rows):
                    for c_i, cell in enumerate(row.cells):
                        cell_p = doc_tbl.cell(r_i, c_i).paragraphs[0]
                        cell_p.text = cell.text.strip()
                        if r_i == 0 and cell_p.runs:
                            cell_p.runs[0].bold = True
                doc.add_paragraph()

        doc.add_paragraph()

    out_io = io.BytesIO()
    doc.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", slide_count


def pptx_to_xlsx(pptx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Extracts tables from PPTX slides to sheets and collects slide text in Document Text."""
    docx_bytes, _, _, cnt = pptx_to_docx(pptx_bytes, base_name)
    return docx_to_xlsx(docx_bytes, base_name)


def pptx_to_pdf(pptx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Converts PPTX to PDF via LibreOffice headless or clean landscape ReportLab slide deck."""
    pdf_bytes = convert_via_libreoffice(pptx_bytes, ".pptx")
    if not pdf_bytes:
        pdf_bytes = _pptx_to_pdf_reportlab(pptx_bytes)
    return pdf_bytes, "application/pdf", f"{base_name}.pdf", 1


def _pptx_to_pdf_reportlab(pptx_bytes: bytes) -> bytes:
    prs = Presentation(io.BytesIO(pptx_bytes))
    buffer = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buffer, pagesize=landscape(letter), leftMargin=36, rightMargin=36, topMargin=36, bottomMargin=36)

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('SlideTitle', parent=styles['Heading1'], fontSize=20, leading=24, spaceAfter=14, textColor=colors.HexColor('#0F172A'))
    bullet_style = ParagraphStyle('SlideBullet', parent=styles['Normal'], fontSize=12, leading=16, spaceAfter=8, textColor=colors.HexColor('#334155'))
    cell_style = ParagraphStyle('SlideCell', parent=styles['Normal'], fontSize=9, leading=12, textColor=colors.HexColor('#0F172A'))
    header_style = ParagraphStyle('SlideHeader', parent=styles['Normal'], fontSize=9, leading=12, fontName='Helvetica-Bold', textColor=colors.HexColor('#0F172A'))

    story = []
    for idx, slide in enumerate(prs.slides, start=1):
        if idx > 1:
            story.append(PageBreak())

        title_text, title_shape, skip_first_p = _extract_slide_title(slide, f"Slide {idx}")
        story.append(RLParagraph(html.escape(title_text), title_style))

        for shape in slide.shapes:
            if shape == title_shape and not skip_first_p:
                continue
            if shape.has_text_frame:
                paragraphs = shape.text_frame.paragraphs
                if shape == title_shape and skip_first_p:
                    paragraphs = paragraphs[1:]
                for p in paragraphs:
                    t = p.text.strip()
                    if t:
                        prefix = "• " if p.level > 0 else ""
                        story.append(RLParagraph(html.escape(prefix + t), bullet_style))
            elif shape.has_table:
                tbl = shape.table
                table_data = []
                for r_idx, row in enumerate(tbl.rows):
                    row_data = [RLParagraph(html.escape(cell.text.strip()), header_style if r_idx == 0 else cell_style) for cell in row.cells]
                    table_data.append(row_data)
                if table_data:
                    rl_tbl = RLTable(table_data)
                    rl_tbl.setStyle(TableStyle([
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CBD5E1')),
                        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#F1F5F9')),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('TOPPADDING', (0, 0), (-1, -1), 4),
                        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
                        ('LEFTPADDING', (0, 0), (-1, -1), 6),
                        ('RIGHTPADDING', (0, 0), (-1, -1), 6),
                    ]))
                    story.append(rl_tbl)
                    story.append(Spacer(1, 10))

    if not story:
        story.append(RLParagraph("Empty Presentation", bullet_style))

    pdf_doc.build(story)
    return buffer.getvalue()


def docx_to_pptx(docx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Creates PowerPoint slides from Word headings, paragraphs, and tables."""
    doc = Document(io.BytesIO(docx_bytes))
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    current_slide = None
    slide_count = 0
    tx_box = None
    tf = None

    def start_new_slide(title="Slide"):
        nonlocal current_slide, slide_count, tx_box, tf
        current_slide = prs.slides.add_slide(blank_layout)
        slide_count += 1
        t_box = current_slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        p = t_box.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(28)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 23, 42)

        tx_box = current_slide.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(5.0))
        tf = tx_box.text_frame
        tf.word_wrap = True

    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            p = Paragraph(child, doc)
            text = p.text.strip()
            if not text:
                continue
            is_heading = p.style.name.startswith("Heading") or (any(r.bold for r in p.runs) and len(text) < 70)
            if is_heading or current_slide is None:
                start_new_slide(text if is_heading else "Introduction")
                if not is_heading and tf:
                    para = tf.add_paragraph()
                    para.text = text
                    para.font.size = Pt(14)
            else:
                if tf:
                    para = tf.add_paragraph()
                    para.text = text
                    para.font.size = Pt(14)
                    para.font.color.rgb = RGBColor(51, 65, 85)

        elif child.tag.endswith('tbl'):
            tbl = Table(child, doc)
            if current_slide is None:
                start_new_slide("Data Table")
            num_rows = len(tbl.rows)
            num_cols = len(tbl.columns)
            if num_rows > 0 and num_cols > 0:
                top_pos = Inches(2.0)
                shape = current_slide.shapes.add_table(num_rows, num_cols, Inches(0.8), top_pos, Inches(11.7), Inches(min(num_rows * 0.4 + 0.5, 4.8)))
                pptx_tbl = shape.table
                for r_i, row in enumerate(tbl.rows):
                    for c_i, cell in enumerate(row.cells):
                        c_cell = pptx_tbl.cell(r_i, c_i)
                        c_cell.text = cell.text.strip()
                        if r_i == 0 and c_cell.text_frame.paragraphs:
                            c_cell.text_frame.paragraphs[0].font.bold = True

    if slide_count == 0:
        start_new_slide("Document Content")

    out_io = io.BytesIO()
    prs.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation", f"{base_name}.pptx", slide_count


def xlsx_to_pptx(xlsx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Turns each Excel sheet into a widescreen presentation slide with an editable table."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    sheet_count = len(wb.worksheets)

    for ws in wb.worksheets:
        slide = prs.slides.add_slide(blank_layout)
        t_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        p = t_box.text_frame.paragraphs[0]
        p.text = ws.title
        p.font.size = Pt(28)
        p.font.bold = True

        rows_data = []
        for row in ws.iter_rows(values_only=True):
            if any(v is not None and str(v).strip() != "" for v in row):
                last_idx = max(i for i, v in enumerate(row) if v is not None and str(v).strip() != "")
                rows_data.append(row[:last_idx + 1])

        if rows_data:
            num_rows = len(rows_data)
            num_cols = max(len(r) for r in rows_data)
            shape = slide.shapes.add_table(num_rows, num_cols, Inches(0.8), Inches(1.8), Inches(11.7), Inches(min(num_rows * 0.4 + 0.5, 5.0)))
            t = shape.table
            for r_i, r in enumerate(rows_data):
                for c_i, val in enumerate(r):
                    cell = t.cell(r_i, c_i)
                    cell.text = "" if val is None else str(val)
                    if r_i == 0 and cell.text_frame.paragraphs:
                        cell.text_frame.paragraphs[0].font.bold = True
        else:
            tx = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.7), Inches(1.0))
            tx.text_frame.paragraphs[0].text = "(Empty worksheet)"

    out_io = io.BytesIO()
    prs.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation", f"{base_name}.pptx", sheet_count


def images_to_pptx(image_list: List[Tuple[str, bytes]], base_name: str) -> Tuple[bytes, str, str, int]:
    """Places each image centered on a widescreen (16:9) presentation slide."""
    if not image_list:
        raise ValueError("No images provided.")

    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for filename, img_bytes in image_list:
        slide = prs.slides.add_slide(blank_layout)
        with Image.open(io.BytesIO(img_bytes)) as pil_img:
            img_w, img_h = pil_img.size
            img_format = (pil_img.format or "").upper()
            if img_format not in ("PNG", "JPEG", "JPG", "BMP", "GIF", "TIFF"):
                png_buf = io.BytesIO()
                pil_img.save(png_buf, format="PNG")
                slide_img_bytes = png_buf.getvalue()
            else:
                slide_img_bytes = img_bytes

        aspect = img_w / img_h
        max_w = Inches(12.0)
        max_h = Inches(6.5)

        if aspect >= (12.0 / 6.5):
            w = max_w
            h = max_w / aspect
        else:
            h = max_h
            w = max_h * aspect

        left = (prs.slide_width - w) / 2
        top = (prs.slide_height - h) / 2
        slide.shapes.add_picture(io.BytesIO(slide_img_bytes), left, top, w, h)

    out_io = io.BytesIO()
    prs.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation", f"{base_name}.pptx", len(image_list)


def pdf_to_pptx(pdf_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Creates presentation slides per PDF page with extracted text/tables or crisp image fallback."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    for p_no, page in enumerate(doc, start=1):
        slide = prs.slides.add_slide(blank_layout)
        text = page.get_text().strip()
        tabs = page.find_tables()
        table_list = tabs.tables if hasattr(tabs, "tables") else list(tabs)

        if text or len(table_list) > 0:
            # Text / table slide
            t_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(0.8))
            t_box.text_frame.paragraphs[0].text = f"Page {p_no}"
            t_box.text_frame.paragraphs[0].font.size = Pt(24)
            t_box.text_frame.paragraphs[0].font.bold = True

            content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.7), Inches(5.2))
            ctf = content_box.text_frame
            ctf.word_wrap = True

            lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
            for ln in lines[:15]:
                p = ctf.add_paragraph()
                p.text = ln
                p.font.size = Pt(13)

            if len(table_list) > 0:
                extracted = table_list[0].extract()
                if extracted:
                    n_r = min(len(extracted), 6)
                    n_c = min(len(extracted[0]), 6)
                    shape = slide.shapes.add_table(n_r, n_c, Inches(0.8), Inches(4.5), Inches(11.7), Inches(2.2))
                    for r_i in range(n_r):
                        for c_i in range(n_c):
                            shape.table.cell(r_i, c_i).text = str(extracted[r_i][c_i] or "")
        else:
            # Image slide fallback
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            slide.shapes.add_picture(io.BytesIO(pix.tobytes("png")), Inches(1.5), Inches(0.5), Inches(10.33), Inches(6.5))

    doc.close()
    out_io = io.BytesIO()
    prs.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation", f"{base_name}.pptx", len(prs.slides)


# --- CSV (.csv) Converters ---

def csv_to_xlsx(csv_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Converts CSV file to a formatted Excel workbook with numeric parsing."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            text = csv_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = csv_bytes.decode("utf-8", errors="replace")

    reader = csv.reader(io.StringIO(text))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sanitize_sheet_title(base_name, set())

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1'),
    )
    header_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    for r_idx, row in enumerate(reader, start=1):
        is_hdr = (r_idx == 1)
        for c_idx, cell_str in enumerate(row, start=1):
            parsed = parse_numeric(cell_str)
            cell = ws.cell(row=r_idx, column=c_idx, value=parsed)
            cell.border = thin_border
            if is_hdr:
                cell.font = Font(bold=True)
                cell.fill = header_fill

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for c in col:
            if c.value is not None:
                max_len = max(max_len, len(str(c.value)))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)

    out_io = io.BytesIO()
    wb.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{base_name}.xlsx", 1


def xlsx_to_csv(xlsx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Exports Excel workbook to single CSV or ZIP of CSVs if multi-sheet."""
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    sheet_count = len(wb.worksheets)

    if sheet_count == 1:
        ws = wb.active
        sio = io.StringIO()
        writer = csv.writer(sio)
        for row in ws.iter_rows(values_only=True):
            if any(v is not None and str(v).strip() != "" for v in row):
                last_idx = max(i for i, v in enumerate(row) if v is not None and str(v).strip() != "")
                writer.writerow([("" if v is None else str(v)) for v in row[:last_idx + 1]])
        csv_bytes = sio.getvalue().encode("utf-8-sig")
        return csv_bytes, "text/csv; charset=utf-8", f"{base_name}.csv", 1
    else:
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for ws in wb.worksheets:
                sio = io.StringIO()
                writer = csv.writer(sio)
                for row in ws.iter_rows(values_only=True):
                    if any(v is not None and str(v).strip() != "" for v in row):
                        last_idx = max(i for i, v in enumerate(row) if v is not None and str(v).strip() != "")
                        writer.writerow([("" if v is None else str(v)) for v in row[:last_idx + 1]])
                clean_title = re.sub(r'[\\/*?:"<>|]', '_', ws.title).strip() or "sheet"
                zf.writestr(f"{clean_title}.csv", sio.getvalue().encode("utf-8-sig"))
        return zip_buf.getvalue(), "application/zip", f"{base_name}_csv_sheets.zip", sheet_count


# --- Text & Web Formats (.txt, .md, .html) ---

def text_to_docx(text_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Turns plain text into a structured Word document."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            text = text_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        text = text_bytes.decode("utf-8", errors="replace")

    doc = Document()
    for line in text.splitlines():
        doc.add_paragraph(line)
    out_io = io.BytesIO()
    doc.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", 1


def md_to_docx(md_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Converts Markdown into an editable Word document with headings, bold/italic, lists, and tables."""
    for enc in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            raw_text = md_bytes.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        raw_text = md_bytes.decode("utf-8", errors="replace")

    html_text = markdown.markdown(raw_text, extensions=['tables', 'fenced_code'])
    return html_to_docx(html_text.encode('utf-8'), base_name)


def html_to_docx(html_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Converts HTML markup into a clean Word document preserving headings, paragraphs, and tables."""
    soup = BeautifulSoup(html_bytes, "html.parser")
    doc = Document()

    for element in soup.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'ul', 'ol', 'table', 'pre']):
        if element.name in ['h1', 'h2', 'h3', 'h4', 'h5', 'h6']:
            level = int(element.name[1])
            doc.add_heading(element.get_text().strip(), level=min(level, 4))
        elif element.name == 'p':
            t = element.get_text().strip()
            if t:
                doc.add_paragraph(t)
        elif element.name in ['ul', 'ol']:
            for li in element.find_all('li'):
                doc.add_paragraph(li.get_text().strip(), style='List Bullet' if element.name == 'ul' else 'List Number')
        elif element.name == 'table':
            rows = element.find_all('tr')
            if rows:
                col_cnt = max(len(r.find_all(['td', 'th'])) for r in rows)
                tbl = doc.add_table(rows=len(rows), cols=col_cnt)
                tbl.style = 'Table Grid'
                for r_i, r in enumerate(rows):
                    cells = r.find_all(['td', 'th'])
                    for c_i, c in enumerate(cells):
                        cell_p = tbl.cell(r_i, c_i).paragraphs[0]
                        cell_p.text = c.get_text().strip()
                        if c.name == 'th' or r_i == 0:
                            if cell_p.runs:
                                cell_p.runs[0].bold = True
                doc.add_paragraph()
        elif element.name == 'pre':
            p = doc.add_paragraph(element.get_text())
            for r in p.runs:
                r.font.name = 'Consolas'

    out_io = io.BytesIO()
    doc.save(out_io)
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", 1


# --- Any Format → Text / Markdown / HTML Exporters ---

def any_to_txt(data_bytes: bytes, input_fmt: str, base_name: str) -> Tuple[bytes, str, str, int]:
    """Extracts clean plain text from any format (running OCR if input is an image)."""
    input_fmt = input_fmt.lower().replace("jpeg", "jpg")

    if input_fmt in ("png", "jpg", "webp"):
        items = extract_ocr_items(data_bytes)
        items.sort(key=lambda it: it['y_center'])
        lines = []
        for it in items:
            lines.append(it['text'])
        text = "\n".join(lines) if lines else "(No text detected on image)"

    elif input_fmt == "pdf":
        doc = fitz.open(stream=data_bytes, filetype="pdf")
        text = "\n\n".join(f"--- Page {i} ---\n" + page.get_text() for i, page in enumerate(doc, start=1))
        doc.close()

    elif input_fmt == "docx":
        doc = Document(io.BytesIO(data_bytes))
        parts = [p.text for p in doc.paragraphs if p.text.strip()]
        for tbl in doc.tables:
            for r in tbl.rows:
                parts.append("\t".join(c.text.strip() for c in r.cells))
        text = "\n".join(parts)

    elif input_fmt in ("xlsx", "csv"):
        if input_fmt == "csv":
            text = data_bytes.decode("utf-8", errors="replace")
        else:
            wb = openpyxl.load_workbook(io.BytesIO(data_bytes), data_only=True)
            parts = []
            for ws in wb.worksheets:
                parts.append(f"=== Sheet: {ws.title} ===")
                for row in ws.iter_rows(values_only=True):
                    if any(v is not None for v in row):
                        parts.append("\t".join("" if v is None else str(v) for v in row))
            text = "\n".join(parts)

    elif input_fmt == "pptx":
        prs = Presentation(io.BytesIO(data_bytes))
        parts = []
        for i, s in enumerate(prs.slides, start=1):
            parts.append(f"--- Slide {i} ---")
            for shp in s.shapes:
                if shp.has_text_frame:
                    parts.append(shp.text_frame.text)
        text = "\n".join(parts)

    elif input_fmt in ("html", "md"):
        text = BeautifulSoup(data_bytes, "html.parser").get_text()

    else:
        text = data_bytes.decode("utf-8", errors="replace")

    return text.encode("utf-8"), "text/plain; charset=utf-8", f"{base_name}.txt", 1


def any_to_md(data_bytes: bytes, input_fmt: str, base_name: str) -> Tuple[bytes, str, str, int]:
    """Generates structured Markdown from any document."""
    input_fmt = input_fmt.lower().replace("jpeg", "jpg")

    if input_fmt in ("png", "jpg", "webp"):
        items = extract_ocr_items(data_bytes)
        items.sort(key=lambda it: it['y_center'])
        md_text = f"# OCR Extracted Content: {base_name}\n\n" + "\n\n".join(it['text'] for it in items)

    elif input_fmt == "docx":
        doc = Document(io.BytesIO(data_bytes))
        parts = []
        for child in doc.element.body.iterchildren():
            if child.tag.endswith('p'):
                p = Paragraph(child, doc)
                t = p.text.strip()
                if not t:
                    continue
                if p.style.name.startswith("Heading 1"):
                    parts.append(f"# {t}\n")
                elif p.style.name.startswith("Heading 2"):
                    parts.append(f"## {t}\n")
                elif p.style.name.startswith("Heading"):
                    parts.append(f"### {t}\n")
                else:
                    parts.append(f"{t}\n")
            elif child.tag.endswith('tbl'):
                tbl = Table(child, doc)
                if tbl.rows:
                    header = [c.text.strip() for c in tbl.rows[0].cells]
                    parts.append("| " + " | ".join(header) + " |")
                    parts.append("| " + " | ".join(["---"] * len(header)) + " |")
                    for row in tbl.rows[1:]:
                        parts.append("| " + " | ".join(c.text.strip() for c in row.cells) + " |")
                    parts.append("")
        md_text = "\n".join(parts)

    elif input_fmt in ("xlsx", "csv"):
        if input_fmt == "csv":
            xlsx_data, _, _, _ = csv_to_xlsx(data_bytes, base_name)
            wb = openpyxl.load_workbook(io.BytesIO(xlsx_data), data_only=True)
        else:
            wb = openpyxl.load_workbook(io.BytesIO(data_bytes), data_only=True)

        parts = []
        for ws in wb.worksheets:
            parts.append(f"## Sheet: {ws.title}\n")
            rows = [r for r in ws.iter_rows(values_only=True) if any(v is not None for v in r)]
            if rows:
                header = ["" if v is None else str(v) for v in rows[0]]
                parts.append("| " + " | ".join(header) + " |")
                parts.append("| " + " | ".join(["---"] * len(header)) + " |")
                for r in rows[1:]:
                    parts.append("| " + " | ".join("" if v is None else str(v) for v in r) + " |")
                parts.append("")
        md_text = "\n".join(parts)

    elif input_fmt == "pdf":
        docx_data, _, _, _ = pdf_to_docx(data_bytes, base_name)
        return any_to_md(docx_data, "docx", base_name)

    elif input_fmt == "pptx":
        docx_data, _, _, _ = pptx_to_docx(data_bytes, base_name)
        return any_to_md(docx_data, "docx", base_name)

    else:
        md_text = data_bytes.decode("utf-8", errors="replace")

    return md_text.encode("utf-8"), "text/markdown; charset=utf-8", f"{base_name}.md", 1


def any_to_html(data_bytes: bytes, input_fmt: str, base_name: str) -> Tuple[bytes, str, str, int]:
    """Generates a clean responsive HTML5 document from any file format."""
    md_bytes, _, _, _ = any_to_md(data_bytes, input_fmt, base_name)
    body_html = markdown.markdown(md_bytes.decode('utf-8'), extensions=['tables', 'fenced_code'])

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>{html.escape(base_name)}</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; line-height: 1.6; max-width: 860px; margin: 40px auto; padding: 0 20px; color: #1e293b; }}
    h1, h2, h3 {{ color: #0f172a; border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; }}
    table {{ border-collapse: collapse; width: 100%; margin: 20px 0; }}
    th, td {{ border: 1px solid #cbd5e1; padding: 8px 12px; text-align: left; }}
    th {{ background: #f1f5f9; font-weight: 600; }}
    code {{ background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-family: monospace; }}
  </style>
</head>
<body>
{body_html}
</body>
</html>"""
    return full_html.encode("utf-8"), "text/html; charset=utf-8", f"{base_name}.html", 1


# --- Universal Hub-and-Spoke Dispatcher ---

def universal_convert(
    input_files: List[Tuple[str, bytes]],
    target_fmt: str,
    base_name: str,
    dpi: int = 300,
    image_quality: int = 95,
    page_range: Optional[str] = None,
) -> Tuple[bytes, str, str, int]:
    """
    Central router that accepts any of the 11 supported formats and routes
    cleanly to the target format using the hub-and-spoke graph.
    Supports resolution (dpi), quality (image_quality), and PDF page slicing (page_range).
    """
    target = target_fmt.lower().strip().replace("jpeg", "jpg")
    base_name = sanitize_base_name(base_name)

    # Validate integrity and format of all input files
    for fname, fbytes in input_files:
        validate_input_file(fname, fbytes)

    primary_filename, primary_bytes = input_files[0]
    source_ext = Path(primary_filename).suffix.lower().lstrip(".").replace("jpeg", "jpg")

    if not source_ext:
        raise ValueError(f"Could not determine source format from filename '{primary_filename}'")

    if source_ext == target and len(input_files) == 1:
        raise ValueError(f"File is already in {target.upper()} format.")

    # Apply PDF page range filtering if input is PDF
    if source_ext == "pdf" and page_range and page_range.strip():
        primary_bytes = filter_pdf_pages(primary_bytes, page_range.strip())
        input_files = [(primary_filename, primary_bytes)]

    # 1. Image Sources (png, jpg, webp)
    if source_ext in ("png", "jpg", "webp"):
        if target in ("png", "jpg", "webp"):
            return convert_image_format(primary_bytes, target, base_name, image_quality=image_quality)
        elif target == "pdf":
            return images_to_pdf(input_files, base_name)
        elif target == "pptx":
            return images_to_pptx(input_files, base_name)
        elif target == "docx":
            return images_to_docx(input_files, base_name)
        elif target == "xlsx":
            return images_to_xlsx(input_files, base_name)
        elif target == "csv":
            xlsx_data, _, _, _ = images_to_xlsx(input_files, base_name)
            return xlsx_to_csv(xlsx_data, base_name)
        elif target == "txt":
            return any_to_txt(primary_bytes, source_ext, base_name)
        elif target == "md":
            return any_to_md(primary_bytes, source_ext, base_name)
        elif target == "html":
            return any_to_html(primary_bytes, source_ext, base_name)

    # 2. PDF Source
    elif source_ext == "pdf":
        if target in ("png", "jpg", "webp"):
            return pdf_to_images(primary_bytes, target, base_name, dpi=dpi, image_quality=image_quality)
        elif target == "docx":
            return pdf_to_docx(primary_bytes, base_name)
        elif target == "xlsx":
            return pdf_to_xlsx(primary_bytes, base_name)
        elif target == "pptx":
            return pdf_to_pptx(primary_bytes, base_name)
        elif target == "csv":
            xlsx_data, _, _, _ = pdf_to_xlsx(primary_bytes, base_name)
            return xlsx_to_csv(xlsx_data, base_name)
        elif target == "txt":
            return any_to_txt(primary_bytes, "pdf", base_name)
        elif target == "md":
            return any_to_md(primary_bytes, "pdf", base_name)
        elif target == "html":
            return any_to_html(primary_bytes, "pdf", base_name)

    # 3. Word (.docx) Source
    elif source_ext == "docx":
        if target == "pdf":
            return docx_to_pdf(primary_bytes, base_name)
        elif target in ("png", "jpg", "webp"):
            pdf_data, _, _, _ = docx_to_pdf(primary_bytes, base_name)
            return pdf_to_images(pdf_data, target, base_name, dpi=dpi, image_quality=image_quality)
        elif target == "xlsx":
            return docx_to_xlsx(primary_bytes, base_name)
        elif target == "pptx":
            return docx_to_pptx(primary_bytes, base_name)
        elif target == "csv":
            xlsx_data, _, _, _ = docx_to_xlsx(primary_bytes, base_name)
            return xlsx_to_csv(xlsx_data, base_name)
        elif target == "txt":
            return any_to_txt(primary_bytes, "docx", base_name)
        elif target == "md":
            return any_to_md(primary_bytes, "docx", base_name)
        elif target == "html":
            return any_to_html(primary_bytes, "docx", base_name)

    # 4. Excel (.xlsx) Source
    elif source_ext == "xlsx":
        if target == "csv":
            return xlsx_to_csv(primary_bytes, base_name)
        elif target == "docx":
            return xlsx_to_docx(primary_bytes, base_name)
        elif target == "pdf":
            return xlsx_to_pdf(primary_bytes, base_name)
        elif target in ("png", "jpg", "webp"):
            pdf_data, _, _, _ = xlsx_to_pdf(primary_bytes, base_name)
            return pdf_to_images(pdf_data, target, base_name, dpi=dpi, image_quality=image_quality)
        elif target == "pptx":
            return xlsx_to_pptx(primary_bytes, base_name)
        elif target == "txt":
            return any_to_txt(primary_bytes, "xlsx", base_name)
        elif target == "md":
            return any_to_md(primary_bytes, "xlsx", base_name)
        elif target == "html":
            return any_to_html(primary_bytes, "xlsx", base_name)

    # 5. PowerPoint (.pptx) Source
    elif source_ext == "pptx":
        if target == "docx":
            return pptx_to_docx(primary_bytes, base_name)
        elif target == "xlsx":
            return pptx_to_xlsx(primary_bytes, base_name)
        elif target == "pdf":
            return pptx_to_pdf(primary_bytes, base_name)
        elif target in ("png", "jpg", "webp"):
            pdf_data, _, _, _ = pptx_to_pdf(primary_bytes, base_name)
            return pdf_to_images(pdf_data, target, base_name, dpi=dpi, image_quality=image_quality)
        elif target == "csv":
            xlsx_data, _, _, _ = pptx_to_xlsx(primary_bytes, base_name)
            return xlsx_to_csv(xlsx_data, base_name)
        elif target == "txt":
            return any_to_txt(primary_bytes, "pptx", base_name)
        elif target == "md":
            return any_to_md(primary_bytes, "pptx", base_name)
        elif target == "html":
            return any_to_html(primary_bytes, "pptx", base_name)

    # 6. CSV (.csv) Source (Hubs through XLSX)
    elif source_ext == "csv":
        xlsx_data, _, _, _ = csv_to_xlsx(primary_bytes, base_name)
        if target == "xlsx":
            return xlsx_data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{base_name}.xlsx", 1
        return universal_convert([(f"{base_name}.xlsx", xlsx_data)], target, base_name, dpi=dpi, image_quality=image_quality, page_range=page_range)

    # 7. Text & Web (.txt, .md, .html) Source (Hubs through DOCX)
    elif source_ext in ("txt", "md", "html"):
        if source_ext == "txt":
            docx_data, _, _, _ = text_to_docx(primary_bytes, base_name)
        elif source_ext == "md":
            docx_data, _, _, _ = md_to_docx(primary_bytes, base_name)
        else:
            docx_data, _, _, _ = html_to_docx(primary_bytes, base_name)

        if target == "docx":
            return docx_data, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", 1
        return universal_convert([(f"{base_name}.docx", docx_data)], target, base_name, dpi=dpi, image_quality=image_quality, page_range=page_range)

    raise ValueError(f"Conversion from '{source_ext.upper()}' to '{target.upper()}' is not supported.")


# --- Batch Conversion Hub ---

def get_unique_zip_name(desired_name: str, used_names: Set[str]) -> str:
    """Ensures a filename inside a ZIP archive is unique by appending _2, _3, etc."""
    safe_name = sanitize_filename(desired_name)
    if safe_name.lower() not in used_names:
        used_names.add(safe_name.lower())
        return safe_name

    stem = Path(safe_name).stem
    ext = Path(safe_name).suffix
    counter = 2
    while True:
        candidate = f"{stem}_{counter}{ext}"
        if candidate.lower() not in used_names:
            used_names.add(candidate.lower())
            return candidate
        counter += 1


def get_unique_folder_name(desired_folder: str, used_folders: Set[str]) -> str:
    """Ensures a subfolder name inside a ZIP archive is unique."""
    safe_folder = sanitize_base_name(desired_folder)
    if safe_folder.lower() not in used_folders:
        used_folders.add(safe_folder.lower())
        return safe_folder

    counter = 2
    while True:
        candidate = f"{safe_folder}_{counter}"
        if candidate.lower() not in used_folders:
            used_folders.add(candidate.lower())
            return candidate
        counter += 1


def batch_convert(
    input_files: List[Tuple[str, bytes]],
    target_fmt: str,
    dpi: int = 300,
    image_quality: int = 95,
    page_range: Optional[str] = None,
) -> Tuple[bytes, str, str, int]:
    """
    Converts multiple files of any mixed formats into target_fmt.
    Bundles all outputs into a single flat/subfoldered ZIP archive without double-zipping.
    Deduplicates filenames, handles same-format files gracefully, and notes any warnings.
    """
    if not input_files:
        raise ValueError("No files provided for batch conversion.")

    target = target_fmt.lower().strip().replace("jpeg", "jpg")
    zip_buffer = io.BytesIO()
    used_entry_names: Set[str] = set()
    used_folders: Set[str] = set()
    warnings: List[str] = []
    converted_count = 0

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as batch_zip:
        for original_filename, file_bytes in input_files:
            safe_orig_name = sanitize_filename(original_filename)
            stem = sanitize_base_name(safe_orig_name)
            source_ext = Path(safe_orig_name).suffix.lower().lstrip(".").replace("jpeg", "jpg")

            # If file is already in the target format
            if source_ext == target:
                unique_name = get_unique_zip_name(f"{stem}.{target}", used_entry_names)
                batch_zip.writestr(unique_name, file_bytes)
                converted_count += 1
                warnings.append(f"'{original_filename}': File was already in {target.upper()} format; included directly as '{unique_name}'.")
                continue

            try:
                content, media_type, out_name, count = universal_convert(
                    [(safe_orig_name, file_bytes)],
                    target,
                    base_name=stem,
                    dpi=dpi,
                    image_quality=image_quality,
                    page_range=page_range,
                )

                # Check if the individual conversion produced a ZIP archive (e.g. multi-page PDF -> PNG)
                if media_type == "application/zip" or out_name.lower().endswith(".zip") or zipfile.is_zipfile(io.BytesIO(content)):
                    subfolder = get_unique_folder_name(stem, used_folders)
                    inner_used_names: Set[str] = set()
                    with zipfile.ZipFile(io.BytesIO(content), "r") as inner_zip:
                        for inner_name in inner_zip.namelist():
                            if inner_name.endswith("/") or inner_name.endswith("\\"):
                                continue
                            inner_data = inner_zip.read(inner_name)
                            clean_inner_name = sanitize_filename(os.path.basename(inner_name))
                            unique_inner_name = get_unique_zip_name(clean_inner_name, inner_used_names)
                            zip_path = f"{subfolder}/{unique_inner_name}"
                            batch_zip.writestr(zip_path, inner_data)
                            converted_count += 1
                else:
                    # Single file output
                    unique_out_name = get_unique_zip_name(out_name, used_entry_names)
                    batch_zip.writestr(unique_out_name, content)
                    converted_count += 1

            except Exception as e:
                logger.warning(f"Batch conversion failed for file '{original_filename}': {e}")
                warnings.append(f"'{original_filename}': Conversion to {target.upper()} failed: {str(e)}")

        if warnings:
            report_text = "Velmo Batch Processing Report\n"
            report_text += "=" * 40 + "\n"
            report_text += f"Target Format: {target.upper()}\n"
            report_text += f"Total Files Processed: {len(input_files)}\n"
            report_text += f"Successful Outputs: {converted_count}\n\n"
            report_text += "Notes & Warnings:\n"
            report_text += "\n".join(f"- {w}" for w in warnings) + "\n"
            batch_zip.writestr("_batch_report.txt", report_text.encode("utf-8"))

    zip_bytes = zip_buffer.getvalue()
    zip_filename = f"batch_converted_{target}.zip"
    return zip_bytes, "application/zip", zip_filename, converted_count


# --- In-Browser Editable Preview & Export Engine ---

def preview_editable(
    input_files: List[Tuple[str, bytes]],
    target_fmt: str,
    page_range: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Extracts structured content from uploaded files for in-browser editing:
    - If target_fmt is tabular (xlsx, csv), returns mode="table" with sheets and rows.
    - If target_fmt is document/text (docx, pdf, pptx, md, txt, html, png, jpg, webp),
      returns mode="document" with structured editable Markdown.
    """
    if not input_files:
        raise ValueError("No files provided for preview.")

    target = target_fmt.lower().strip().replace("jpeg", "jpg")
    if target not in SUPPORTED_FORMATS:
        raise ValueError(f"Unsupported target format '{target_fmt}'.")

    # Validate integrity and format of all input files
    for fname, fbytes in input_files:
        validate_input_file(fname, fbytes)

    primary_filename, primary_bytes = input_files[0]
    base_name = sanitize_base_name(primary_filename)
    source_ext = Path(primary_filename).suffix.lower().lstrip(".").replace("jpeg", "jpg")

    if target in ("xlsx", "csv"):
        # Tabular mode
        if source_ext == "xlsx":
            xlsx_bytes = primary_bytes
        elif source_ext == "csv":
            xlsx_bytes, _, _, _ = csv_to_xlsx(primary_bytes, base_name)
        else:
            xlsx_bytes, _, _, _ = universal_convert(
                input_files,
                "xlsx",
                base_name=base_name,
                page_range=page_range,
            )

        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
        sheets = []
        for ws in wb.worksheets:
            sheet_rows = []
            for row in ws.iter_rows(values_only=True):
                if any(v is not None and str(v).strip() != "" for v in row):
                    last_idx = max(i for i, v in enumerate(row) if v is not None and str(v).strip() != "")
                    sheet_rows.append([("" if v is None else str(v)) for v in row[:last_idx + 1]])
                else:
                    sheet_rows.append([])
            while sheet_rows and not sheet_rows[-1]:
                sheet_rows.pop()
            if not sheet_rows:
                sheet_rows = [["", ""]]
            sheets.append({"name": ws.title, "rows": sheet_rows})

        if not sheets:
            sheets = [{"name": "Sheet 1", "rows": [["", ""]]}]

        return {
            "mode": "table",
            "filename": base_name,
            "target_format": target,
            "sheets": sheets,
        }
    else:
        # Document mode
        if source_ext == "md":
            md_text = primary_bytes.decode("utf-8", errors="replace")
        elif source_ext == "txt":
            md_text = primary_bytes.decode("utf-8", errors="replace")
        else:
            md_bytes, _, _, _ = universal_convert(
                input_files,
                "md",
                base_name=base_name,
                page_range=page_range,
            )
            md_text = md_bytes.decode("utf-8", errors="replace")

        return {
            "mode": "document",
            "filename": base_name,
            "target_format": target,
            "content": md_text,
        }


def export_edited(
    mode: str,
    filename: str,
    target_format: str,
    content: Optional[str] = None,
    sheets: Optional[List[Dict[str, Any]]] = None,
    dpi: int = 300,
    image_quality: int = 95,
) -> Tuple[bytes, str, str, int]:
    """
    Exports edited document or spreadsheet table data directly into the requested target format.
    - If mode == "table": builds a styled openpyxl Workbook with numeric parsing and exports cleanly.
    - If mode == "document": takes Markdown text and routes through the hub-and-spoke pipeline.
    """
    target = target_format.lower().strip().replace("jpeg", "jpg")
    if target not in SUPPORTED_FORMATS:
        raise ValueError(f"Unsupported target format '{target_format}'. Choose from: {', '.join(sorted(SUPPORTED_FORMATS))}")

    base_name = sanitize_base_name(filename or "edited_document")

    if mode == "table":
        sheets_data = sheets or []
        if not sheets_data:
            sheets_data = [{"name": "Sheet 1", "rows": [[""]]}]

        wb = openpyxl.Workbook()
        default_sheet = wb.active
        used_titles: Set[str] = set()

        thin_border = Border(
            left=Side(style='thin', color='CBD5E1'),
            right=Side(style='thin', color='CBD5E1'),
            top=Side(style='thin', color='CBD5E1'),
            bottom=Side(style='thin', color='CBD5E1'),
        )
        header_fill = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

        for idx, sheet_info in enumerate(sheets_data):
            title = sanitize_sheet_title(sheet_info.get("name") or f"Sheet {idx+1}", used_titles)
            if idx == 0:
                ws = default_sheet
                ws.title = title
            else:
                ws = wb.create_sheet(title=title)

            rows = sheet_info.get("rows", [])
            for r_idx, row in enumerate(rows, start=1):
                is_hdr = (r_idx == 1)
                for c_idx, cell_val in enumerate(row, start=1):
                    parsed = parse_numeric(str(cell_val) if cell_val is not None else "")
                    cell = ws.cell(row=r_idx, column=c_idx, value=parsed)
                    cell.border = thin_border
                    if is_hdr:
                        cell.font = Font(bold=True)
                        cell.fill = header_fill

            for col in ws.columns:
                max_len = 0
                col_letter = get_column_letter(col[0].column)
                for c in col:
                    if c.value is not None:
                        max_len = max(max_len, len(str(c.value)))
                ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)

        out_io = io.BytesIO()
        wb.save(out_io)
        xlsx_bytes = out_io.getvalue()

        if target == "xlsx":
            return xlsx_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{base_name}.xlsx", len(wb.worksheets)
        elif target == "csv":
            return xlsx_to_csv(xlsx_bytes, base_name)
        else:
            return universal_convert([(f"{base_name}.xlsx", xlsx_bytes)], target, base_name, dpi=dpi, image_quality=image_quality)

    elif mode == "document":
        md_text = content if content is not None else ""
        md_bytes = md_text.encode("utf-8")

        if target == "md":
            return md_bytes, "text/markdown; charset=utf-8", f"{base_name}.md", 1
        elif target == "txt":
            return md_bytes, "text/plain; charset=utf-8", f"{base_name}.txt", 1
        elif target == "docx":
            return md_to_docx(md_bytes, base_name)
        elif target == "html":
            return any_to_html(md_bytes, "md", base_name)
        else:
            return universal_convert([(f"{base_name}.md", md_bytes)], target, base_name, dpi=dpi, image_quality=image_quality)

    else:
        raise ValueError(f"Unknown edit mode '{mode}'. Must be 'table' or 'document'.")


# --- Fast Inline PDF Page Preview ---

def preview_pdf_pages(pdf_bytes: bytes, max_pages: int = 15, dpi: int = 130) -> Dict[str, Any]:
    """
    Renders up to max_pages of a PDF document at 130 DPI JPEG as base64 data URLs for fast inline preview.
    Returns:
        {"total_pages": int, "pages": List[str]}
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Invalid or corrupted PDF file: {e}")

    try:
        total_pages = len(doc)
        if total_pages == 0:
            raise ValueError("PDF document has 0 pages or is empty.")

        pages_b64: List[str] = []
        limit = min(total_pages, max_pages)
        for i in range(limit):
            page = doc[i]
            # Render page at specified DPI (130) without alpha for crisp, compact JPEG
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            img_bytes = pix.tobytes("jpeg", jpg_quality=88)
            b64_str = base64.b64encode(img_bytes).decode("ascii")
            pages_b64.append(f"data:image/jpeg;base64,{b64_str}")

        return {
            "total_pages": total_pages,
            "pages": pages_b64,
        }
    finally:
        doc.close()

