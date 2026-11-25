import base64
import csv
import gc
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

# Limit native BLAS/OpenMP CPU threads to prevent memory spikes on 512MB RAM
os.environ.setdefault("OMP_NUM_THREADS", "1")
os.environ.setdefault("OPENBLAS_NUM_THREADS", "1")
os.environ.setdefault("MKL_NUM_THREADS", "1")

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
from pptx.enum.text import MSO_ANCHOR

import markdown
from bs4 import BeautifulSoup
from pdf2docx import Converter

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image as RLImage,
    NextPageTemplate,
    PageBreak,
    PageTemplate,
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
    """Returns the lazy-loaded singleton instance of RapidOCR tuned for 512MB RAM safety."""
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        logger.info("Initializing lazy singleton RapidOCR engine (1 thread, bounded tensor size)...")
        _ocr_engine = RapidOCR(
            intra_op_num_threads=1,
            inter_op_num_threads=1,
            use_cls=False,
            max_side_len=1120,
            det_limit_side_len=736,
            det_limit_type="min",
            det_box_thresh=0.35,
            det_thresh=0.25,
            rec_batch_num=1,
        )
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
    """Parses numeric strings (including comma-grouped numbers like '1,420.50') into int or float."""
    if not text:
        return ""
    stripped = text.strip()
    if not stripped:
        return ""

    # Standard or comma-grouped integer (preserve leading-zero codes like '007' or '0123')
    clean_candidate = stripped
    if re.match(r"^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$", stripped):
        clean_candidate = stripped.replace(",", "")

    if re.match(r"^[+-]?\d+$", clean_candidate):
        unsigned = clean_candidate.lstrip("+-")
        if unsigned == "0" or not unsigned.startswith("0"):
            try:
                return int(clean_candidate)
            except ValueError:
                pass

    if re.match(r"^[+-]?(\d+\.\d+|\.\d+)$", clean_candidate):
        try:
            return float(clean_candidate)
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
    try:
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
    finally:
        gc.collect()


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
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Invalid or corrupted PDF file: {e}")

    try:
        total_pages = len(doc)
        if total_pages == 0:
            raise ValueError("PDF document has no pages.")

        selected_indices = parse_page_range(page_range_str, total_pages)
        if selected_indices != list(range(total_pages)):
            doc.select(selected_indices)
            out_bytes = doc.tobytes()
        else:
            out_bytes = pdf_bytes
        return out_bytes
    finally:
        doc.close()
        _release_pdf_memory()


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


def _release_pdf_memory():
    """Flushes PyMuPDF's internal global cache and triggers Python GC for 512MB RAM safety."""
    try:
        fitz.TOOLS.store_shrink(100)
    except Exception:
        pass
    gc.collect()


def pdf_to_images(
    pdf_bytes: bytes,
    target_fmt: str,
    base_name: str,
    dpi: int = 300,
    image_quality: int = 95
) -> Tuple[bytes, str, str, int]:
    """Renders PDF pages to opaque white-paper PNG, JPG, or WebP at custom DPI and quality."""
    target_fmt = target_fmt.lower().replace("jpeg", "jpg")
    if target_fmt not in ("png", "jpg", "webp"):
        raise ValueError(f"Unsupported image format: {target_fmt}")

    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Invalid or corrupted PDF file: {e}")

    def _render_page(page, idx: int) -> Tuple[str, bytes]:
        # Always flatten page canvas onto white paper (alpha=False) to prevent
        # transparent/black-box artifacts on dark viewers, JPG, and WEBP.
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        try:
            if target_fmt == "png":
                return f"{base_name}_page_{idx:02d}.png", pix.tobytes("png")
            elif target_fmt == "jpg":
                return f"{base_name}_page_{idx:02d}.jpg", pix.tobytes("jpg", jpg_quality=image_quality)
            else:
                buf = io.BytesIO()
                with Image.frombytes("RGB", [pix.width, pix.height], pix.samples) as pil_img:
                    pil_img.save(buf, format="WEBP", quality=image_quality)
                return f"{base_name}_page_{idx:02d}.webp", buf.getvalue()
        finally:
            del pix

    try:
        page_count = len(doc)
        if page_count == 0:
            raise ValueError("PDF contains no pages.")

        if page_count == 1:
            _, data = _render_page(doc[0], 1)
            mime_map = {"png": "image/png", "jpg": "image/jpeg", "webp": "image/webp"}
            return data, mime_map[target_fmt], f"{base_name}.{target_fmt}", 1

        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx, page in enumerate(doc, start=1):
                fname, img_bytes = _render_page(page, idx)
                zf.writestr(fname, img_bytes)
                del img_bytes
                if idx % 3 == 0:
                    _release_pdf_memory()
        return zip_buf.getvalue(), "application/zip", f"{base_name}_{target_fmt}_pages.zip", page_count
    finally:
        doc.close()
        _release_pdf_memory()


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
    list_counters: Dict[Tuple[str, int], int] = {}

    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            p_info = _extract_oxml_paragraph_data(child, doc)
            text = p_info["plain_text"].strip()
            if not text:
                continue
            p = Paragraph(child, doc)
            style_name = p.style.name if p.style else "Normal"
            hl = p_info["heading_level"]
            is_heading = (
                hl > 0
                or style_name.startswith("Heading")
                or (any(s.get("bold") for s in p_info["segments"]) and len(text) < 80 and not p_info["list_info"] and not p_info["is_callout"])
            )
            display_text = text
            if p_info["list_info"] is not None:
                li = p_info["list_info"]
                ilvl = li["ilvl"]
                indent = "  " * ilvl
                if li["ordered"]:
                    key = (li["num_id"], ilvl)
                    list_counters[key] = list_counters.get(key, 0) + 1
                    display_text = f"{indent}{list_counters[key]}. {text}"
                else:
                    bullet = "•" if ilvl == 0 else ("◦" if ilvl == 1 else "▪")
                    display_text = f"{indent}{bullet} {text}"

            document_paragraphs.append((style_name, display_text, is_heading))
            if is_heading:
                last_heading_text = text

        elif child.tag.endswith('tbl'):
            table = Table(child, doc)
            table_index += 1
            raw_title = last_heading_text if last_heading_text else f"Table {table_index}"
            sheet_title = sanitize_sheet_title(raw_title, used_sheet_titles)
            last_heading_text = None

            ws = wb.create_sheet(title=sheet_title)
            raw_rows, span_cmds, repeat_hdr_rows, _ = _extract_docx_table_grid_and_spans(table, doc)

            for r_idx, row in enumerate(raw_rows, start=1):
                is_header_row = (r_idx <= max(repeat_hdr_rows, 1))
                for c_idx, cell in enumerate(row, start=1):
                    excel_cell = ws.cell(row=r_idx, column=c_idx)
                    excel_cell.border = thin_border
                    if not cell or not cell["is_origin"]:
                        continue

                    c_text = "\n".join(
                        p["plain_text"].strip() for p in cell["paras"] if p["plain_text"].strip()
                    ).strip()
                    parsed_val = parse_numeric(c_text)
                    excel_cell.value = parsed_val

                    fill_hex = cell.get("fill")
                    is_cell_bold = is_header_row or any(
                        s.get("bold") for p in cell["paras"] for s in p["segments"]
                    )
                    if fill_hex and fill_hex not in ("FFFFFF", "AUTO"):
                        excel_cell.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
                        if _is_dark_hex(fill_hex):
                            excel_cell.font = Font(bold=is_cell_bold, color="FFFFFF")
                        elif is_cell_bold:
                            excel_cell.font = Font(bold=True)
                    else:
                        if is_cell_bold:
                            excel_cell.font = Font(bold=True)
                        if is_header_row:
                            excel_cell.fill = header_fill

            for c0, r0, c1, r1 in span_cmds:
                ws.merge_cells(start_row=r0 + 1, start_column=c0 + 1, end_row=r1 + 1, end_column=c1 + 1)

            _auto_fit_xlsx_columns(ws)

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

    _auto_fit_xlsx_columns(ws_text, min_width=12, max_width=65)

    if default_sheet in wb.worksheets and len(wb.worksheets) > 1:
        wb.remove(default_sheet)

    out_io = io.BytesIO()
    wb.save(out_io)
    wb.close()
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{base_name}.xlsx", table_index


def _is_scanned_pdf_page(page) -> bool:
    """Returns True if a PDF page has minimal/no digital text and contains a large raster image."""
    txt = page.get_text("text").strip()
    if len(txt) >= 40:
        return False
    imgs = page.get_images(full=True)
    if not imgs:
        return False
    if len(txt) == 0:
        return True
    page_area = max(abs(page.rect.width * page.rect.height), 1.0)
    try:
        for info in page.get_image_info():
            bbox = fitz.Rect(info["bbox"])
            if bbox.get_area() / page_area >= 0.25:
                return True
    except Exception:
        pass
    return len(txt) < 15


def _ocr_pdf_page(page, base_dpi: int = 150) -> List[dict]:
    """Renders a single PDF page with bounded pixel dimensions and extracts OCR items."""
    max_pt = max(page.rect.width, page.rect.height, 1.0)
    eff_dpi = min(base_dpi, max(110, int(1600 * 72.0 / max_pt)))
    pix = page.get_pixmap(dpi=eff_dpi, alpha=False)
    img_bytes = pix.tobytes("png")
    del pix
    try:
        return extract_ocr_items(img_bytes)
    finally:
        del img_bytes
        gc.collect()


def _cluster_ocr_table_rows(multi_lines: List[List[dict]]) -> List[List[str]]:
    """Aligns multi-span OCR lines into a consistent 2D table grid via horizontal column clustering."""
    if not multi_lines:
        return []

    # Build horizontal column intervals [x_min, x_max] from all spans
    all_spans = [sp for ln in multi_lines for sp in ln]
    avg_h = statistics.median([sp["height"] for sp in all_spans]) if all_spans else 20.0
    all_spans_sorted = sorted(all_spans, key=lambda s: s["x_min"])

    col_bounds: List[List[float]] = []  # [min_x, max_x]
    for sp in all_spans_sorted:
        placed = False
        for cb in col_bounds:
            # Overlap or tight horizontal alignment within 1.2 * avg_h
            overlap = min(cb[1], sp["x_max"]) - max(cb[0], sp["x_min"])
            if overlap > -0.4 * avg_h or abs(sp["x_min"] - cb[0]) <= 1.2 * avg_h:
                cb[0] = min(cb[0], sp["x_min"])
                cb[1] = max(cb[1], sp["x_max"])
                placed = True
                break
        if not placed:
            col_bounds.append([sp["x_min"], sp["x_max"]])

    col_bounds.sort(key=lambda cb: cb[0])
    num_cols = len(col_bounds)
    if num_cols < 2:
        return []

    grid: List[List[str]] = []
    for ln in multi_lines:
        row_cells = [""] * num_cols
        for sp in ln:
            # Assign span to column with maximum overlap or nearest interval
            best_c = 0
            best_score = float("-inf")
            for c_i, (cmin, cmax) in enumerate(col_bounds):
                overlap = min(cmax, sp["x_max"]) - max(cmin, sp["x_min"])
                if overlap > 0:
                    score = overlap
                else:
                    dist = min(abs(sp["x_min"] - cmax), abs(cmin - sp["x_max"]))
                    score = -dist
                if score > best_score:
                    best_score = score
                    best_c = c_i
            if row_cells[best_c]:
                row_cells[best_c] += " " + sp["text"]
            else:
                row_cells[best_c] = sp["text"]
        grid.append(row_cells)
    return grid


def _structure_ocr_blocks(items: List[dict]) -> List[Dict[str, Any]]:
    """
    Reconstructs RapidOCR bounding boxes into structured blocks:
    - {"type": "heading", "level": 1|2, "text": str}
    - {"type": "paragraph", "text": str}
    - {"type": "table", "rows": List[List[str]]}
    """
    if not items:
        return []

    items_sorted = sorted(items, key=lambda it: it["y_center"])
    raw_lines: List[List[dict]] = []
    for item in items_sorted:
        placed = False
        for line in raw_lines:
            line_y = sum(it["y_center"] for it in line) / len(line)
            line_h = sum(it["height"] for it in line) / len(line)
            if abs(item["y_center"] - line_y) <= max(line_h, item["height"]) * 0.55:
                line.append(item)
                placed = True
                break
        if not placed:
            raw_lines.append([item])

    raw_lines.sort(key=lambda ln: sum(it["y_center"] for it in ln) / len(ln))

    # Merge tight intra-sentence items on each line; keep wide column gaps separated
    span_lines: List[Dict[str, Any]] = []
    for ln in raw_lines:
        ln.sort(key=lambda it: it["x_min"])
        line_h = sum(it["height"] for it in ln) / len(ln)
        merged_spans: List[dict] = []
        for it in ln:
            if not merged_spans:
                merged_spans.append(dict(it))
            else:
                prev = merged_spans[-1]
                gap = it["x_min"] - prev["x_max"]
                if gap < line_h * 1.1:
                    prev["text"] = f"{prev['text']} {it['text']}"
                    prev["x_max"] = max(prev["x_max"], it["x_max"])
                    prev["y_min"] = min(prev["y_min"], it["y_min"])
                    prev["y_max"] = max(prev["y_max"], it["y_max"])
                    prev["y_center"] = (prev["y_min"] + prev["y_max"]) / 2.0
                    prev["x_center"] = (prev["x_min"] + prev["x_max"]) / 2.0
                    prev["height"] = max(prev["height"], it["height"])
                else:
                    merged_spans.append(dict(it))

        span_lines.append({
            "spans": merged_spans,
            "y_center": sum(s["y_center"] for s in merged_spans) / len(merged_spans),
            "height": line_h,
            "is_multi": len(merged_spans) >= 2,
        })

    all_heights = [sp["height"] for sl in span_lines for sp in sl["spans"]]
    median_h = statistics.median(all_heights) if all_heights else 20.0

    blocks: List[Dict[str, Any]] = []
    i = 0
    n = len(span_lines)
    para_buf: List[Dict[str, Any]] = []

    def flush_para():
        if not para_buf:
            return
        text = " ".join(" ".join(s["text"] for s in ln["spans"]) for ln in para_buf).strip()
        if text:
            blocks.append({"type": "paragraph", "text": text})
        para_buf.clear()

    while i < n:
        sl = span_lines[i]
        # Detect contiguous multi-column table block (>= 2 consecutive multi-span lines)
        if sl["is_multi"]:
            j = i + 1
            while j < n and span_lines[j]["is_multi"]:
                j += 1
            if (j - i) >= 2:
                table_grid = _cluster_ocr_table_rows([span_lines[k]["spans"] for k in range(i, j)])
                if table_grid and len(table_grid[0]) >= 2:
                    flush_para()
                    blocks.append({"type": "table", "rows": table_grid})
                    i = j
                    continue

        # Prose or heading line
        line_text = " ".join(s["text"] for s in sl["spans"]).strip()
        avg_h = sl["height"]
        if avg_h >= 1.28 * median_h and len(line_text) < 110:
            flush_para()
            lvl = 1 if avg_h >= 1.55 * median_h else 2
            blocks.append({"type": "heading", "level": lvl, "text": line_text})
        else:
            if para_buf:
                prev_sl = para_buf[-1]
                dy = sl["y_center"] - prev_sl["y_center"]
                ref_h = max(prev_sl["height"], sl["height"], 1.0)
                if dy > 1.65 * ref_h:
                    flush_para()
            para_buf.append(sl)
        i += 1

    flush_para()
    return blocks


def _append_ocr_blocks_to_docx(doc: Document, blocks: List[Dict[str, Any]], insert_before_elem=None):
    """Appends structured OCR headings, paragraphs, and Word tables to a python-docx Document."""
    created_elems = []
    for blk in blocks:
        btype = blk["type"]
        if btype == "heading":
            el = doc.add_heading(blk["text"], level=blk.get("level", 2))
            created_elems.append(el._p)
        elif btype == "paragraph":
            el = doc.add_paragraph(blk["text"])
            created_elems.append(el._p)
        elif btype == "table":
            rows = blk["rows"]
            if rows:
                num_rows = len(rows)
                num_cols = max(len(r) for r in rows)
                tbl = doc.add_table(rows=num_rows, cols=num_cols)
                tbl.style = "Table Grid"
                for r_i, row in enumerate(rows):
                    for c_i in range(num_cols):
                        val = row[c_i] if c_i < len(row) else ""
                        cell = tbl.cell(r_i, c_i)
                        p = cell.paragraphs[0]
                        run = p.add_run(val)
                        if r_i == 0:
                            run.bold = True
                            set_cell_background(cell, "F1F5F9")
                created_elems.append(tbl._tbl)
                sp_p = doc.add_paragraph()
                created_elems.append(sp_p._p)

    if insert_before_elem is not None:
        for xml_el in created_elems:
            insert_before_elem.addprevious(xml_el)


def pdf_to_docx(pdf_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    base_name = sanitize_base_name(base_name)
    try:
        src_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Invalid or corrupted PDF file: {e}")

    ocr_blocks_by_page: Dict[int, List[Dict[str, Any]]] = {}
    all_scanned = False
    page_count = len(src_doc)

    try:
        scanned_indices = [idx for idx, page in enumerate(src_doc) if _is_scanned_pdf_page(page)]
        all_scanned = (len(scanned_indices) == page_count and page_count > 0)

        for idx in scanned_indices:
            items = _ocr_pdf_page(src_doc[idx], base_dpi=150)
            blocks = _structure_ocr_blocks(items)
            if blocks:
                ocr_blocks_by_page[idx] = blocks

        if all_scanned and ocr_blocks_by_page:
            word_doc = Document()
            for idx in range(page_count):
                if idx > 0:
                    word_doc.add_page_break()
                blocks = ocr_blocks_by_page.get(idx, [])
                if blocks:
                    _append_ocr_blocks_to_docx(word_doc, blocks)
                else:
                    word_doc.add_paragraph(f"(No text detected on page {idx + 1})")
            out_io = io.BytesIO()
            word_doc.save(out_io)
            return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", 1

        # Prepare temp PDF (replacing OCR-extracted scanned pages with lightweight slot markers)
        with tempfile.TemporaryDirectory() as tmp_dir:
            temp_pdf_path = os.path.join(tmp_dir, "input.pdf")
            temp_docx_path = os.path.join(tmp_dir, "output.docx")

            if ocr_blocks_by_page:
                for idx in sorted(ocr_blocks_by_page.keys()):
                    rect = src_doc[idx].rect
                    w, h = rect.width, rect.height
                    src_doc.delete_page(idx)
                    marker_page = src_doc.new_page(pno=idx, width=w, height=h)
                    marker_page.insert_text((50, 72), f"__VELMO_OCR_PAGE_{idx}__", fontsize=11)
                src_doc.save(temp_pdf_path, deflate=True)
            else:
                with open(temp_pdf_path, "wb") as f:
                    f.write(pdf_bytes)

            src_doc.close()
            src_doc = None
            gc.collect()

            cv = None
            try:
                cv = Converter(temp_pdf_path)
                cv.convert(temp_docx_path, start=0, end=None)
            finally:
                if cv:
                    cv.close()
                gc.collect()

            if not ocr_blocks_by_page:
                with open(temp_docx_path, "rb") as f:
                    docx_bytes = f.read()
            else:
                word_doc = Document(temp_docx_path)
                for idx, blocks in ocr_blocks_by_page.items():
                    token = f"__VELMO_OCR_PAGE_{idx}__"
                    target_p = None
                    for p in word_doc.paragraphs:
                        if token in p.text:
                            target_p = p
                            break
                    if target_p is not None:
                        _append_ocr_blocks_to_docx(word_doc, blocks, insert_before_elem=target_p._p)
                        parent = target_p._p.getparent()
                        if parent is not None:
                            parent.remove(target_p._p)
                    else:
                        word_doc.add_page_break()
                        _append_ocr_blocks_to_docx(word_doc, blocks)
                out_io = io.BytesIO()
                word_doc.save(out_io)
                docx_bytes = out_io.getvalue()
    finally:
        if src_doc is not None:
            src_doc.close()
        _release_pdf_memory()

    return docx_bytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", 1


def _find_pdf_tables(page):
    """Finds PDF tables without triggering PyMuPDF 2.x ONNX layout models (saves ~150MB RAM)."""
    try:
        return page.find_tables(use_layout=False)
    except TypeError:
        return page.find_tables()


def _extract_pdf_table_grid(page, tab) -> List[List[str]]:
    """
    Extracts table cell strings using exact cell clip rects to prevent PyMuPDF
    baseline-splitting artifacts on decimals ('1,420.50') and underscores ('REG_CTRL_01').
    """
    try:
        rows_obj = getattr(tab, "rows", None)
        if rows_obj:
            grid: List[List[str]] = []
            for r in rows_obj:
                row_vals: List[str] = []
                for cbox in r.cells:
                    if cbox is None:
                        row_vals.append("")
                    else:
                        txt = page.get_text("text", clip=fitz.Rect(cbox)).strip()
                        row_vals.append(re.sub(r"\s*\n\s*", " ", txt))
                if any(v != "" for v in row_vals):
                    grid.append(row_vals)
            if grid:
                return grid
    except Exception:
        pass

    raw = tab.extract() or []
    return [[re.sub(r"\s*\n\s*", " ", "" if c is None else str(c).strip()) for c in row] for row in raw]


def _auto_fit_xlsx_columns(ws, min_width: int = 10, max_width: int = 50):
    """Adjusts openpyxl column widths cleanly based on cell contents."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for c in col:
            if c.value is not None:
                lines = str(c.value).split("\n")
                max_len = max(max_len, max(len(l) for l in lines))
        ws.column_dimensions[col_letter].width = min(max(max_len + 3, min_width), max_width)


def _write_table_rows_to_sheet(ws, rows: List[List[str]], start_row: int, thin_border, header_fill, style_first_row: bool = True) -> int:
    """Writes parsed rows into an openpyxl worksheet and returns the next free row index."""
    r_idx = start_row
    for row in rows:
        is_header = style_first_row and (r_idx == 1)
        for c_idx, cell_val in enumerate(row, start=1):
            c_str = "" if cell_val is None else str(cell_val).strip()
            parsed = parse_numeric(c_str)
            excel_cell = ws.cell(row=r_idx, column=c_idx, value=parsed)
            excel_cell.border = thin_border
            if is_header:
                excel_cell.font = Font(bold=True)
                excel_cell.fill = header_fill
        r_idx += 1
    return r_idx


def pdf_to_xlsx(pdf_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Invalid or corrupted PDF file: {e}")

    try:
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
        document_text_blocks: List[Tuple[int, str]] = []

        # Track previous table state across consecutive pages to merge multi-page tables cleanly
        last_ws = None
        last_header_sig: Optional[Tuple[str, ...]] = None
        last_next_row = 1
        last_table_page = -1

        for p_no, page in enumerate(doc, start=1):
            if _is_scanned_pdf_page(page):
                items = _ocr_pdf_page(page, base_dpi=150)
                ocr_blocks = _structure_ocr_blocks(items)
                t_no = 0
                for blk in ocr_blocks:
                    if blk["type"] == "table" and blk["rows"]:
                        t_no += 1
                        table_count += 1
                        sheet_title = sanitize_sheet_title(f"Page_{p_no}_Table_{t_no}", used_sheet_titles)
                        ws = wb.create_sheet(title=sheet_title)
                        next_r = _write_table_rows_to_sheet(ws, blk["rows"], 1, thin_border, header_fill, style_first_row=True)
                        _auto_fit_xlsx_columns(ws)
                        last_ws = ws
                        last_header_sig = tuple(c.strip() for c in blk["rows"][0])
                        last_next_row = next_r
                        last_table_page = p_no
                    elif blk["type"] in ("heading", "paragraph") and blk.get("text"):
                        document_text_blocks.append((p_no, blk["text"]))
                continue

            tabs = _find_pdf_tables(page)
            table_bboxes = []
            valid_tables: List[Tuple[Any, List[List[str]]]] = []
            for tab in tabs:
                extracted = _extract_pdf_table_grid(page, tab)
                if not extracted:
                    continue
                # Skip trivial 1-row layout/callout boxes so prose flows into Document Text
                if len(extracted) < 2:
                    continue
                valid_tables.append((tab, extracted))
                table_bboxes.append(tab.bbox)

            for t_no, (tab, extracted) in enumerate(valid_tables, start=1):
                cur_header_sig = tuple(c.strip() for c in extracted[0])
                can_merge_with_prev = (
                    t_no == 1
                    and last_ws is not None
                    and last_table_page == p_no - 1
                    and last_header_sig is not None
                    and len(cur_header_sig) == len(last_header_sig)
                    and cur_header_sig == last_header_sig
                    and len(extracted) > 1
                )
                if can_merge_with_prev:
                    last_next_row = _write_table_rows_to_sheet(
                        last_ws, extracted[1:], last_next_row, thin_border, header_fill, style_first_row=False
                    )
                    _auto_fit_xlsx_columns(last_ws)
                    last_table_page = p_no
                else:
                    table_count += 1
                    sheet_title = sanitize_sheet_title(f"Page_{p_no}_Table_{t_no}", used_sheet_titles)
                    ws = wb.create_sheet(title=sheet_title)
                    next_r = _write_table_rows_to_sheet(ws, extracted, 1, thin_border, header_fill, style_first_row=True)
                    _auto_fit_xlsx_columns(ws)
                    last_ws = ws
                    last_header_sig = cur_header_sig
                    last_next_row = next_r
                    last_table_page = p_no

            blocks = page.get_text("blocks")
            for b in blocks:
                if len(b) > 6 and b[6] != 0:
                    continue
                b_rect = fitz.Rect(b[0], b[1], b[2], b[3])
                b_text = b[4].strip()
                if not b_text:
                    continue
                b_area = max(b_rect.get_area(), 1.0)
                in_table = any(
                    ((b_rect & fitz.Rect(t_box)).get_area() / b_area) > 0.35
                    for t_box in table_bboxes
                )
                if not in_table:
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

        _auto_fit_xlsx_columns(ws_text, min_width=12, max_width=65)

        if default_sheet in wb.worksheets and len(wb.worksheets) > 1:
            wb.remove(default_sheet)

        out_io = io.BytesIO()
        wb.save(out_io)
        wb.close()
        return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{base_name}.xlsx", table_count
    finally:
        doc.close()
        _release_pdf_memory()


_REPORTLAB_FONTS_INITIALIZED = False
_REPORTLAB_SANS_FAMILY = "Helvetica"
_REPORTLAB_MONO_FAMILY = "Courier"


def _ensure_reportlab_unicode_fonts() -> Tuple[str, str]:
    """
    Registers bundled DejaVu Sans / DejaVu Sans Mono TrueType fonts (with OS fallback)
    so ReportLab renders ₹, €, £, —, ±, α, β, μ, ≤, ≥, °C, Ω cleanly without tofu boxes.
    Returns: (sans_family, mono_family)
    """
    global _REPORTLAB_FONTS_INITIALIZED, _REPORTLAB_SANS_FAMILY, _REPORTLAB_MONO_FAMILY
    if _REPORTLAB_FONTS_INITIALIZED:
        return _REPORTLAB_SANS_FAMILY, _REPORTLAB_MONO_FAMILY

    base_dir = Path(__file__).resolve().parent
    candidate_dirs = [
        base_dir / "fonts",
        Path("/usr/share/fonts/truetype/dejavu"),
        Path("/usr/share/fonts/truetype/liberation"),
        Path("C:/Windows/Fonts"),
    ]

    sans_specs = [
        ("DejaVuSans.ttf", "DejaVuSans-Bold.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans-BoldOblique.ttf"),
        ("LiberationSans-Regular.ttf", "LiberationSans-Bold.ttf", "LiberationSans-Italic.ttf", "LiberationSans-BoldItalic.ttf"),
        ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf"),
    ]
    mono_specs = [
        ("DejaVuSansMono.ttf", "DejaVuSansMono-Bold.ttf"),
        ("LiberationMono-Regular.ttf", "LiberationMono-Bold.ttf"),
        ("consola.ttf", "consolab.ttf"),
    ]

    for d in candidate_dirs:
        if not d.is_dir():
            continue
        for reg, bld, ita, bld_ita in sans_specs:
            p_reg = d / reg
            if p_reg.is_file():
                try:
                    p_bld = d / bld if (d / bld).is_file() else p_reg
                    p_ita = d / ita if (d / ita).is_file() else p_reg
                    p_bi = d / bld_ita if (d / bld_ita).is_file() else p_bld
                    pdfmetrics.registerFont(TTFont("VelmoSans", str(p_reg)))
                    pdfmetrics.registerFont(TTFont("VelmoSans-Bold", str(p_bld)))
                    pdfmetrics.registerFont(TTFont("VelmoSans-Italic", str(p_ita)))
                    pdfmetrics.registerFont(TTFont("VelmoSans-BoldItalic", str(p_bi)))
                    pdfmetrics.registerFontFamily(
                        "VelmoSans",
                        normal="VelmoSans",
                        bold="VelmoSans-Bold",
                        italic="VelmoSans-Italic",
                        boldItalic="VelmoSans-BoldItalic",
                    )
                    _REPORTLAB_SANS_FAMILY = "VelmoSans"
                    break
                except Exception as e:
                    logger.warning(f"Failed registering sans font from {p_reg}: {e}")
        if _REPORTLAB_SANS_FAMILY == "VelmoSans":
            break

    for d in candidate_dirs:
        if not d.is_dir():
            continue
        for reg, bld in mono_specs:
            p_reg = d / reg
            if p_reg.is_file():
                try:
                    p_bld = d / bld if (d / bld).is_file() else p_reg
                    pdfmetrics.registerFont(TTFont("VelmoMono", str(p_reg)))
                    pdfmetrics.registerFont(TTFont("VelmoMono-Bold", str(p_bld)))
                    pdfmetrics.registerFontFamily(
                        "VelmoMono",
                        normal="VelmoMono",
                        bold="VelmoMono-Bold",
                        italic="VelmoMono",
                        boldItalic="VelmoMono-Bold",
                    )
                    _REPORTLAB_MONO_FAMILY = "VelmoMono"
                    break
                except Exception as e:
                    logger.warning(f"Failed registering mono font from {p_reg}: {e}")
        if _REPORTLAB_MONO_FAMILY == "VelmoMono":
            break

    _REPORTLAB_FONTS_INITIALIZED = True
    return _REPORTLAB_SANS_FAMILY, _REPORTLAB_MONO_FAMILY


class _VelmoNumberedCanvas(rl_canvas.Canvas):
    """Two-pass ReportLab canvas supporting accurate 'Page X of Y' running footers."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states: List[dict] = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_deferred_page_footer(total_pages)
            super().showPage()
        super().save()

    def _draw_deferred_page_footer(self, total_pages: int):
        cb = getattr(self, "_velmo_footer_cb", None)
        if cb is not None:
            cb(self, self._pageNumber, total_pages)


def _flatten_image_for_reportlab(img_bytes: bytes) -> Tuple[bytes, int, int]:
    """Composites any transparent PNG/RGBA/LA/P image onto opaque white and returns (png_bytes, w, h)."""
    with Image.open(io.BytesIO(img_bytes)) as im:
        iw, ih = max(im.width, 1), max(im.height, 1)
        fmt = (im.format or "").upper()
        mode = im.mode
        has_alpha = mode in ("RGBA", "LA") or (mode == "P" and "transparency" in im.info)
        if has_alpha:
            im_rgba = im.convert("RGBA")
            bg = Image.new("RGBA", im_rgba.size, (255, 255, 255, 255))
            flat = Image.alpha_composite(bg, im_rgba).convert("RGB")
            buf = io.BytesIO()
            flat.save(buf, format="PNG")
            return buf.getvalue(), iw, ih
        if fmt not in ("PNG", "JPEG", "JPG") or mode not in ("RGB", "L"):
            rgb = im.convert("RGB")
            buf = io.BytesIO()
            rgb.save(buf, format="PNG")
            return buf.getvalue(), iw, ih
        return img_bytes, iw, ih


def _extract_oxml_paragraph_data(p_elem, doc: Document) -> Dict[str, Any]:
    """
    Walks a <w:p> XML element cleanly to extract:
    - ordered inline segments (text, bold, italic, underline, color, mono, link_url)
    - embedded images [(rId, cx_pt, cy_pt)]
    - explicit page breaks (before / inline)
    - heading level, list bullet/number info, alignment, callout shading/borders
    """
    ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    ns_r = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    ns_a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    ns_wp = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    ns_v = "{urn:schemas-microsoft-com:vml}"

    p_obj = Paragraph(p_elem, doc)
    style_name = (p_obj.style.name if p_obj.style else "") or ""
    style_id = (p_obj.style.style_id if p_obj.style else "") or ""
    pPr = p_elem.find(f"{ns_w}pPr")

    page_break_before = False
    has_section_break = False
    is_callout = False
    callout_bg = None
    callout_border = None
    list_info: Optional[Dict[str, Any]] = None
    align_val = TA_LEFT

    if p_obj.alignment == 1:
        align_val = TA_CENTER
    elif p_obj.alignment == 2:
        align_val = TA_RIGHT
    elif p_obj.alignment == 3:
        align_val = TA_JUSTIFY

    if pPr is not None:
        if pPr.find(f"{ns_w}pageBreakBefore") is not None:
            page_break_before = True
        if pPr.find(f"{ns_w}sectPr") is not None:
            has_section_break = True

        shd = pPr.find(f"{ns_w}shd")
        if shd is not None:
            fill = (shd.get(f"{ns_w}fill") or "").strip().upper()
            if re.match(r"^[0-9A-F]{6}$", fill) and fill not in ("FFFFFF", "AUTO"):
                callout_bg = fill
                is_callout = True

        pBdr = pPr.find(f"{ns_w}pBdr")
        if pBdr is not None:
            left_b = pBdr.find(f"{ns_w}left")
            if left_b is not None:
                bcol = (left_b.get(f"{ns_w}color") or "").strip().upper()
                if re.match(r"^[0-9A-F]{6}$", bcol):
                    callout_border = bcol
                    is_callout = True

        numPr = pPr.find(f"{ns_w}numPr")
        if numPr is not None:
            ilvl_el = numPr.find(f"{ns_w}ilvl")
            numId_el = numPr.find(f"{ns_w}numId")
            ilvl = int(ilvl_el.get(f"{ns_w}val", "0")) if ilvl_el is not None else 0
            num_id = numId_el.get(f"{ns_w}val", "0") if numId_el is not None else "0"
            if num_id != "0":
                list_info = {"num_id": num_id, "ilvl": ilvl, "ordered": "number" in style_name.lower()}

    if list_info is None:
        sn_low = style_name.lower()
        if "list bullet" in sn_low:
            lvl = 1 if ("2" in sn_low or "3" in sn_low) else 0
            list_info = {"num_id": "style_bullet", "ilvl": lvl, "ordered": False}
        elif "list number" in sn_low:
            lvl = 1 if ("2" in sn_low or "3" in sn_low) else 0
            list_info = {"num_id": "style_number", "ilvl": lvl, "ordered": True}

    # Resolve ordered vs bullet from numbering.xml if available
    if list_info is not None and list_info["num_id"] not in ("style_bullet", "style_number"):
        try:
            num_part = doc.part.numbering_part
            if num_part is not None:
                num_el = num_part.element.xpath(f'.//w:num[@w:numId="{list_info["num_id"]}"]')
                if num_el:
                    abs_id_el = num_el[0].find(f"{ns_w}abstractNumId")
                    if abs_id_el is not None:
                        abs_id = abs_id_el.get(f"{ns_w}val")
                        lvl_els = num_part.element.xpath(
                            f'.//w:abstractNum[@w:abstractNumId="{abs_id}"]/w:lvl[@w:ilvl="{list_info["ilvl"]}"]/w:numFmt'
                        )
                        if lvl_els:
                            fmt_val = (lvl_els[0].get(f"{ns_w}val") or "").lower()
                            list_info["ordered"] = fmt_val != "bullet"
        except Exception:
            pass

    if "quote" in style_name.lower():
        is_callout = True

    heading_level = 0
    m_h = re.match(r"^(?:Heading|Titre)\s*([1-6])$", style_name, re.I) or re.match(r"^Heading([1-6])$", style_id, re.I)
    if m_h:
        heading_level = int(m_h.group(1))
    elif style_name.lower() == "title":
        heading_level = 1

    segments: List[Dict[str, Any]] = []
    images: List[Tuple[str, Optional[float], Optional[float]]] = []
    inline_page_breaks = 0
    mono_chars = 0
    total_chars = 0

    def _parse_rpr(r_el) -> Dict[str, Any]:
        rPr = r_el.find(f"{ns_w}rPr")
        bold = False
        italic = False
        underline = False
        color = None
        font_name = ""
        font_size = None
        if rPr is not None:
            b_el = rPr.find(f"{ns_w}b")
            if b_el is not None and (b_el.get(f"{ns_w}val") or "true").lower() not in ("0", "false"):
                bold = True
            i_el = rPr.find(f"{ns_w}i")
            if i_el is not None and (i_el.get(f"{ns_w}val") or "true").lower() not in ("0", "false"):
                italic = True
            u_el = rPr.find(f"{ns_w}u")
            if u_el is not None and (u_el.get(f"{ns_w}val") or "single").lower() != "none":
                underline = True
            c_el = rPr.find(f"{ns_w}color")
            if c_el is not None:
                cval = (c_el.get(f"{ns_w}val") or "").strip().upper()
                if re.match(r"^[0-9A-F]{6}$", cval):
                    color = cval
            rFonts = rPr.find(f"{ns_w}rFonts")
            if rFonts is not None:
                font_name = rFonts.get(f"{ns_w}ascii") or rFonts.get(f"{ns_w}hAnsi") or rFonts.get(f"{ns_w}cs") or ""
            sz_el = rPr.find(f"{ns_w}sz")
            if sz_el is not None:
                try:
                    font_size = int(sz_el.get(f"{ns_w}val")) / 2.0
                except Exception:
                    pass
        return {
            "bold": bold,
            "italic": italic,
            "underline": underline,
            "color": color,
            "mono": _is_mono_font(font_name),
            "size": font_size,
        }

    def _walk_run(r_el, link_url: Optional[str] = None):
        nonlocal inline_page_breaks, mono_chars, total_chars
        meta = _parse_rpr(r_el)
        if link_url and not meta["color"]:
            meta["color"] = "2563EB"
            meta["underline"] = True

        for ch in r_el:
            tag = ch.tag
            if tag == f"{ns_w}t":
                txt = ch.text or ""
                if txt:
                    segments.append({"text": txt, "url": link_url, **meta})
                    n_ns = len(txt.strip())
                    total_chars += n_ns
                    if meta["mono"]:
                        mono_chars += n_ns
            elif tag == f"{ns_w}tab":
                segments.append({"text": "    ", "url": link_url, **meta})
            elif tag == f"{ns_w}br":
                btype = (ch.get(f"{ns_w}type") or "").lower()
                if btype == "page":
                    inline_page_breaks += 1
                else:
                    segments.append({"text": "\n", "url": link_url, **meta})
            elif tag in (f"{ns_w}drawing", f"{ns_w}pict"):
                # Extract DrawingML blip or VML imagedata
                for blip in ch.iter(f"{ns_a}blip"):
                    rid = blip.get(f"{ns_r}embed") or blip.get(f"{ns_r}link")
                    if rid:
                        cx_pt, cy_pt = None, None
                        for ext_el in ch.iter(f"{ns_wp}extent"):
                            try:
                                cx = int(ext_el.get("cx", "0"))
                                cy = int(ext_el.get("cy", "0"))
                                if cx > 0 and cy > 0:
                                    cx_pt = cx / 12700.0
                                    cy_pt = cy / 12700.0
                            except Exception:
                                pass
                            break
                        images.append((rid, cx_pt, cy_pt))
                for vimg in ch.iter(f"{ns_v}imagedata"):
                    rid = vimg.get(f"{ns_r}id")
                    if rid:
                        images.append((rid, None, None))

    for child in p_elem:
        if child.tag == f"{ns_w}r":
            _walk_run(child, None)
        elif child.tag == f"{ns_w}hyperlink":
            rid = child.get(f"{ns_r}id")
            anchor = child.get(f"{ns_w}anchor")
            url = None
            if rid and rid in doc.part.rels:
                url = doc.part.rels[rid].target_ref
            elif anchor:
                url = f"#{anchor}"
            for r_sub in child.findall(f"{ns_w}r"):
                _walk_run(r_sub, url)
        elif child.tag == f"{ns_w}fldSimple":
            instr = (child.get(f"{ns_w}instr") or "").upper()
            if "PAGE" in instr and "NUMPAGES" not in instr:
                segments.append({"text": "__VELMO_PAGE__", "url": None, "bold": False, "italic": False, "underline": False, "color": None, "mono": False, "size": None})
            elif "NUMPAGES" in instr:
                segments.append({"text": "__VELMO_NUMPAGES__", "url": None, "bold": False, "italic": False, "underline": False, "color": None, "mono": False, "size": None})
            else:
                for r_sub in child.findall(f"{ns_w}r"):
                    _walk_run(r_sub, None)

    is_mono_block = total_chars > 0 and (mono_chars / total_chars) >= 0.85 and ("\n" in "".join(s["text"] for s in segments) or is_callout)
    plain_text = "".join(s["text"] for s in segments)
    return {
        "segments": segments,
        "plain_text": plain_text,
        "images": images,
        "heading_level": heading_level,
        "list_info": list_info,
        "align": align_val,
        "page_break_before": page_break_before,
        "inline_page_breaks": inline_page_breaks,
        "has_section_break": has_section_break,
        "is_callout": is_callout,
        "callout_bg": callout_bg,
        "callout_border": callout_border,
        "is_mono_block": is_mono_block,
    }


def _segments_to_rl_markup(
    segments: List[Dict[str, Any]],
    mono_family: str,
    preserve_newlines: bool = False,
    default_color: Optional[str] = None,
    strip_edges: bool = True,
) -> str:
    """Converts extracted paragraph segments into ReportLab XML markup with links and inline formatting."""
    parts: List[str] = []
    # Merge adjacent segments with same URL so ReportLab renders a single crisp <a> link
    merged: List[Dict[str, Any]] = []
    for seg in segments:
        txt = seg["text"]
        if not txt:
            continue
        if (
            merged
            and seg.get("url") is not None
            and merged[-1].get("url") == seg.get("url")
            and merged[-1].get("bold") == seg.get("bold")
            and merged[-1].get("italic") == seg.get("italic")
        ):
            merged[-1]["text"] += txt
        else:
            merged.append(dict(seg))

    for seg in merged:
        txt = html.escape(seg["text"])
        if preserve_newlines:
            txt = txt.replace("  ", "&nbsp;&nbsp;").replace("\n", "<br/>")
        else:
            txt = txt.replace("\n", "<br/>")

        if seg.get("mono"):
            txt = f'<font name="{mono_family}">{txt}</font>'
        if seg.get("bold"):
            txt = f"<b>{txt}</b>"
        if seg.get("italic"):
            txt = f"<i>{txt}</i>"
        if seg.get("underline") and not seg.get("url"):
            txt = f"<u>{txt}</u>"

        col = seg.get("color") or default_color
        if seg.get("url"):
            esc_url = html.escape(seg["url"], quote=True)
            link_col = f"#{col}" if col else "#2563EB"
            txt = f'<a href="{esc_url}" color="{link_col}"><u>{txt}</u></a>'
        elif col and col.upper() not in ("000000", "AUTO"):
            txt = f'<font color="#{col}">{txt}</font>'
        parts.append(txt)

    joined = "".join(parts)
    return joined.strip() if strip_edges else joined


def _extract_docx_table_grid_and_spans(table: Table, doc: Document):
    """
    Parses a python-docx Table accurately handling horizontal (gridSpan) and vertical (vMerge) merges,
    background cell fills, repeating header rows (w:tblHeader), and gridCol relative widths.
    """
    ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    tbl_el = table._tbl

    # Extract gridCol widths (in twips -> points)
    grid_widths: List[float] = []
    tblGrid = tbl_el.find(f"{ns_w}tblGrid")
    if tblGrid is not None:
        for gc_el in tblGrid.findall(f"{ns_w}gridCol"):
            w_tw = gc_el.get(f"{ns_w}w")
            if w_tw and w_tw.isdigit():
                grid_widths.append(int(w_tw) / 20.0)

    tr_list = tbl_el.findall(f"{ns_w}tr")
    num_rows = len(tr_list)
    if num_rows == 0:
        return [], [], 0, []

    raw_rows: List[List[Optional[dict]]] = []
    repeat_hdr_rows = 0
    max_cols = len(grid_widths)

    for r_idx, tr in enumerate(tr_list):
        trPr = tr.find(f"{ns_w}trPr")
        if trPr is not None and trPr.find(f"{ns_w}tblHeader") is not None and r_idx == repeat_hdr_rows:
            repeat_hdr_rows += 1

        row_cells: List[Optional[dict]] = []
        for tc in tr.findall(f"{ns_w}tc"):
            tcPr = tc.find(f"{ns_w}tcPr")
            grid_span = 1
            v_merge = None
            fill_hex = None
            if tcPr is not None:
                gs_el = tcPr.find(f"{ns_w}gridSpan")
                if gs_el is not None:
                    try:
                        grid_span = max(1, int(gs_el.get(f"{ns_w}val", "1")))
                    except Exception:
                        grid_span = 1
                vm_el = tcPr.find(f"{ns_w}vMerge")
                if vm_el is not None:
                    v_merge = (vm_el.get(f"{ns_w}val") or "continue").lower()
                shd_el = tcPr.find(f"{ns_w}shd")
                if shd_el is not None:
                    fval = (shd_el.get(f"{ns_w}fill") or "").strip().upper()
                    if re.match(r"^[0-9A-F]{6}$", fval) and fval != "AUTO":
                        fill_hex = fval

            para_infos = [_extract_oxml_paragraph_data(p_el, doc) for p_el in tc.findall(f"{ns_w}p")]
            cell_info = {
                "paras": para_infos,
                "grid_span": grid_span,
                "v_merge": v_merge,
                "fill": fill_hex,
                "is_origin": True,
            }
            row_cells.append(cell_info)
            for _ in range(grid_span - 1):
                row_cells.append({"paras": [], "grid_span": 1, "v_merge": v_merge, "fill": fill_hex, "is_origin": False})

        max_cols = max(max_cols, len(row_cells))
        raw_rows.append(row_cells)

    if max_cols == 0:
        return [], [], 0, []

    for r in raw_rows:
        while len(r) < max_cols:
            r.append({"paras": [], "grid_span": 1, "v_merge": None, "fill": None, "is_origin": False})

    # Compute SPAN commands for both horizontal and vertical merges
    span_cmds: List[Tuple[int, int, int, int]] = []  # (c0, r0, c1, r1)
    visited = [[False] * max_cols for _ in range(num_rows)]

    for r0 in range(num_rows):
        for c0 in range(max_cols):
            if visited[r0][c0]:
                continue
            cell = raw_rows[r0][c0]
            if not cell or not cell["is_origin"]:
                visited[r0][c0] = True
                continue
            c1 = min(max_cols - 1, c0 + cell["grid_span"] - 1)
            r1 = r0
            if cell["v_merge"] == "restart":
                while r1 + 1 < num_rows:
                    below = raw_rows[r1 + 1][c0]
                    if below and below["v_merge"] == "continue":
                        below["is_origin"] = False
                        r1 += 1
                    else:
                        break
            for rr in range(r0, r1 + 1):
                for cc in range(c0, c1 + 1):
                    visited[rr][cc] = True
                    if rr != r0 or cc != c0:
                        raw_rows[rr][cc]["is_origin"] = False
            if c1 > c0 or r1 > r0:
                span_cmds.append((c0, r0, c1, r1))

    if repeat_hdr_rows == 0 and num_rows > 3:
        repeat_hdr_rows = 1

    return raw_rows, span_cmds, repeat_hdr_rows, grid_widths


def _compute_smart_table_col_widths(
    raw_rows: List[List[Optional[dict]]],
    grid_widths: List[float],
    avail_width: float,
) -> Tuple[List[float], float]:
    """Computes balanced ReportLab column widths strictly within avail_width and picks clean font size."""
    num_cols = len(raw_rows[0]) if raw_rows else 0
    if num_cols == 0:
        return [], 8.5

    if num_cols >= 10:
        font_pt = 6.5
        pad_pt = 4.5
    elif num_cols >= 8:
        font_pt = 7.0
        pad_pt = 5.5
    elif num_cols >= 6:
        font_pt = 8.0
        pad_pt = 7.0
    else:
        font_pt = 8.8
        pad_pt = 9.0

    # Estimate minimum natural width (longest unbreakable word) and desired width per column
    min_w = [26.0] * num_cols
    pref_w = [36.0] * num_cols

    for r in raw_rows:
        for c_i, cell in enumerate(r):
            if not cell or not cell["is_origin"] or cell["grid_span"] > 1:
                continue
            txt = " ".join(p["plain_text"] for p in cell["paras"]).strip()
            if not txt:
                continue
            words = re.split(r"\s+", txt)
            longest_word = max((len(w) for w in words), default=0)
            avg_char = font_pt * 0.56
            min_w[c_i] = max(min_w[c_i], longest_word * avg_char + pad_pt)
            pref_w[c_i] = max(pref_w[c_i], min(len(txt), 34) * avg_char + pad_pt)

    # Blend with Word gridCol hints if valid
    if len(grid_widths) == num_cols and sum(grid_widths) > 0:
        g_scale = avail_width / sum(grid_widths)
        for c_i in range(num_cols):
            pref_w[c_i] = max(min_w[c_i], 0.55 * pref_w[c_i] + 0.45 * (grid_widths[c_i] * g_scale))

    sum_min = sum(min_w)
    if sum_min >= avail_width:
        # Shrink font slightly if even minimum words exceed avail_width
        ratio = avail_width / max(sum_min, 1.0)
        if ratio < 0.92:
            font_pt = max(5.6, round(font_pt * max(ratio, 0.78), 1))
        widths = [w * ratio for w in min_w]
    else:
        extra = avail_width - sum_min
        deltas = [max(pref_w[i] - min_w[i], 0.5) for i in range(num_cols)]
        sum_delta = sum(deltas)
        widths = [min_w[i] + extra * (deltas[i] / sum_delta) for i in range(num_cols)]

    return widths, font_pt


def _is_dark_hex(hex_str: Optional[str]) -> bool:
    if not hex_str or not re.match(r"^[0-9A-F]{6}$", hex_str.upper()):
        return False
    r = int(hex_str[0:2], 16)
    g = int(hex_str[2:4], 16)
    b = int(hex_str[4:6], 16)
    lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return lum < 115.0


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
    sans_family, mono_family = _ensure_reportlab_unicode_fonts()
    doc = Document(io.BytesIO(docx_bytes))
    buffer = io.BytesIO()

    # Collect per-section geometry and header/footer text
    sections_meta = []
    for sec in doc.sections:
        pw = sec.page_width.pt if sec.page_width else letter[0]
        ph = sec.page_height.pt if sec.page_height else letter[1]
        if sec.orientation == WD_ORIENTATION.LANDSCAPE and pw < ph:
            pw, ph = ph, pw
        elif sec.orientation == WD_ORIENTATION.PORTRAIT and pw > ph:
            pw, ph = ph, pw

        lm = max(sec.left_margin.pt if sec.left_margin else 42.0, 28.0)
        rm = max(sec.right_margin.pt if sec.right_margin else 42.0, 28.0)
        tm = max(sec.top_margin.pt if sec.top_margin else 44.0, 42.0)
        bm = max(sec.bottom_margin.pt if sec.bottom_margin else 44.0, 42.0)

        hdr_text = ""
        if sec.header:
            hdr_parts = []
            for p in sec.header.paragraphs:
                p_info = _extract_oxml_paragraph_data(p._p, doc)
                t = p_info["plain_text"].strip()
                if t:
                    hdr_parts.append(t)
            hdr_text = " | ".join(hdr_parts)

        ftr_text = ""
        if sec.footer:
            ftr_parts = []
            for p in sec.footer.paragraphs:
                p_info = _extract_oxml_paragraph_data(p._p, doc)
                t = p_info["plain_text"].strip()
                if t:
                    ftr_parts.append(t)
            ftr_text = " | ".join(ftr_parts)

        sections_meta.append({
            "pw": pw,
            "ph": ph,
            "lm": lm,
            "rm": rm,
            "tm": tm,
            "bm": bm,
            "avail_w": max(pw - lm - rm, 120.0),
            "avail_h": max(ph - tm - bm, 120.0),
            "hdr": hdr_text,
            "ftr": ftr_text,
        })

    if not sections_meta:
        sections_meta.append({
            "pw": letter[0], "ph": letter[1], "lm": 42.0, "rm": 42.0, "tm": 44.0, "bm": 44.0,
            "avail_w": letter[0] - 84.0, "avail_h": letter[1] - 88.0, "hdr": "", "ftr": "",
        })

    def _make_on_page(sec_m: dict):
        def _on_page(canvas_obj, doc_obj):
            canvas_obj.saveState()
            if sec_m["hdr"]:
                canvas_obj.setFont(sans_family, 8.0)
                canvas_obj.setFillColor(colors.HexColor("#64748B"))
                canvas_obj.drawString(sec_m["lm"], sec_m["ph"] - 26.0, sec_m["hdr"][:120])
                canvas_obj.setStrokeColor(colors.HexColor("#E2E8F0"))
                canvas_obj.setLineWidth(0.5)
                canvas_obj.line(sec_m["lm"], sec_m["ph"] - 31.0, sec_m["pw"] - sec_m["rm"], sec_m["ph"] - 31.0)
            canvas_obj.restoreState()

            def _deferred_footer(c_obj, p_no: int, tot_p: int):
                if not sec_m["ftr"]:
                    return
                c_obj.saveState()
                c_obj.setStrokeColor(colors.HexColor("#E2E8F0"))
                c_obj.setLineWidth(0.5)
                c_obj.line(sec_m["lm"], 32.0, sec_m["pw"] - sec_m["rm"], 32.0)
                c_obj.setFont(sans_family, 8.0)
                c_obj.setFillColor(colors.HexColor("#64748B"))
                ftr_str = sec_m["ftr"].replace("__VELMO_PAGE__", str(p_no)).replace("__VELMO_NUMPAGES__", str(tot_p))
                c_obj.drawString(sec_m["lm"], 20.0, ftr_str[:135])
                c_obj.restoreState()

            canvas_obj._velmo_footer_cb = _deferred_footer

        return _on_page

    page_templates: List[PageTemplate] = []
    for s_idx, sm in enumerate(sections_meta):
        frame = Frame(
            sm["lm"],
            sm["bm"],
            sm["avail_w"],
            sm["avail_h"],
            id=f"frame_{s_idx}",
            leftPadding=0,
            rightPadding=0,
            topPadding=0,
            bottomPadding=0,
        )
        page_templates.append(
            PageTemplate(
                id=f"sec_{s_idx}",
                frames=[frame],
                pagesize=(sm["pw"], sm["ph"]),
                onPage=_make_on_page(sm),
            )
        )

    pdf_doc = BaseDocTemplate(
        buffer,
        pagesize=(sections_meta[0]["pw"], sections_meta[0]["ph"]),
        leftMargin=sections_meta[0]["lm"],
        rightMargin=sections_meta[0]["rm"],
        topMargin=sections_meta[0]["tm"],
        bottomMargin=sections_meta[0]["bm"],
        pageTemplates=page_templates,
    )

    styles = getSampleStyleSheet()
    h1_style = ParagraphStyle(
        "DocxH1", parent=styles["Heading1"], fontName=f"{sans_family}-Bold" if sans_family == "VelmoSans" else "Helvetica-Bold",
        fontSize=16, leading=20, spaceBefore=6, spaceAfter=6, textColor=colors.HexColor("#0F172A"), keepWithNext=True,
    )
    h2_style = ParagraphStyle(
        "DocxH2", parent=styles["Heading2"], fontName=f"{sans_family}-Bold" if sans_family == "VelmoSans" else "Helvetica-Bold",
        fontSize=13, leading=17, spaceBefore=6, spaceAfter=4, textColor=colors.HexColor("#1E293B"), keepWithNext=True,
    )
    h3_style = ParagraphStyle(
        "DocxH3", parent=styles["Heading3"], fontName=f"{sans_family}-Bold" if sans_family == "VelmoSans" else "Helvetica-Bold",
        fontSize=11, leading=14.5, spaceBefore=5, spaceAfter=3, textColor=colors.HexColor("#1E293B"), keepWithNext=True,
    )
    h4_style = ParagraphStyle(
        "DocxH4", parent=styles["Heading4"], fontName=f"{sans_family}-Bold" if sans_family == "VelmoSans" else "Helvetica-Bold",
        fontSize=10, leading=13.5, spaceBefore=4, spaceAfter=3, textColor=colors.HexColor("#334155"), keepWithNext=True,
    )
    body_style = ParagraphStyle(
        "DocxBody", parent=styles["Normal"], fontName=sans_family,
        fontSize=9.8, leading=13.6, spaceAfter=5, textColor=colors.HexColor("#1E293B"),
    )
    code_style = ParagraphStyle(
        "DocxCode", parent=styles["Normal"], fontName=mono_family,
        fontSize=8.0, leading=10.8, textColor=colors.HexColor("#0F172A"),
    )

    story = []
    keepalive_bufs: List[io.BytesIO] = []
    list_counters: Dict[Tuple[str, int], int] = {}
    cur_sec_idx = 0

    for child in doc.element.body.iterchildren():
        sm = sections_meta[min(cur_sec_idx, len(sections_meta) - 1)]

        if child.tag.endswith("p"):
            info = _extract_oxml_paragraph_data(child, doc)
            if info["page_break_before"]:
                story.append(PageBreak())

            # Embedded images inside paragraph
            for rid, cx_pt, cy_pt in info["images"]:
                rel_part = doc.part.related_parts.get(rid)
                if rel_part is None:
                    continue
                try:
                    flat_bytes, iw, ih = _flatten_image_for_reportlab(rel_part.blob)
                    target_w = cx_pt if cx_pt else (sm["avail_w"] * 0.85)
                    target_h = cy_pt if cy_pt else (target_w * (ih / iw))
                    max_w = sm["avail_w"] - 4.0
                    max_h = sm["avail_h"] * 0.72
                    scale = min(max_w / max(target_w, 1.0), max_h / max(target_h, 1.0), 1.0)
                    draw_w = target_w * scale
                    draw_h = target_h * scale
                    buf = io.BytesIO(flat_bytes)
                    keepalive_bufs.append(buf)
                    rl_img = RLImage(buf, width=draw_w, height=draw_h)
                    rl_img.hAlign = "CENTER" if info["align"] == TA_CENTER else ("RIGHT" if info["align"] == TA_RIGHT else "LEFT")
                    story.append(rl_img)
                    story.append(Spacer(1, 4))
                except Exception as img_err:
                    logger.warning(f"Failed rendering DOCX embedded image {rid}: {img_err}")

            # Text content
            if info["plain_text"].strip():
                if info["is_mono_block"]:
                    markup = _segments_to_rl_markup(info["segments"], mono_family, preserve_newlines=True, strip_edges=False)
                    box = RLTable([[RLParagraph(markup, code_style)]], colWidths=[sm["avail_w"] - 12])
                    box.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F1F5F9")),
                        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                        ("LINEBEFORE", (0, 0), (0, -1), 3.0, colors.HexColor("#475569")),
                        ("TOPPADDING", (0, 0), (-1, -1), 6),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                        ("LEFTPADDING", (0, 0), (-1, -1), 8),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ]))
                    story.append(box)
                    story.append(Spacer(1, 6))
                elif info["is_callout"]:
                    markup = _segments_to_rl_markup(info["segments"], mono_family)
                    bg_hex = f"#{info['callout_bg']}" if info["callout_bg"] else "#EFF6FF"
                    bd_hex = f"#{info['callout_border']}" if info["callout_border"] else "#2563EB"
                    box = RLTable([[RLParagraph(markup, body_style)]], colWidths=[sm["avail_w"] - 12])
                    box.setStyle(TableStyle([
                        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(bg_hex)),
                        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                        ("LINEBEFORE", (0, 0), (0, -1), 3.5, colors.HexColor(bd_hex)),
                        ("TOPPADDING", (0, 0), (-1, -1), 6),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
                        ("LEFTPADDING", (0, 0), (-1, -1), 9),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
                    ]))
                    story.append(box)
                    story.append(Spacer(1, 6))
                else:
                    markup = _segments_to_rl_markup(info["segments"], mono_family)
                    hl = info["heading_level"]
                    if hl == 1:
                        st = ParagraphStyle("H1Dyn", parent=h1_style, alignment=info["align"])
                        story.append(RLParagraph(markup, st))
                    elif hl == 2:
                        st = ParagraphStyle("H2Dyn", parent=h2_style, alignment=info["align"])
                        story.append(RLParagraph(markup, st))
                    elif hl == 3:
                        st = ParagraphStyle("H3Dyn", parent=h3_style, alignment=info["align"])
                        story.append(RLParagraph(markup, st))
                    elif hl >= 4:
                        st = ParagraphStyle("H4Dyn", parent=h4_style, alignment=info["align"])
                        story.append(RLParagraph(markup, st))
                    elif info["list_info"] is not None:
                        li = info["list_info"]
                        ilvl = li["ilvl"]
                        if li["ordered"]:
                            key = (li["num_id"], ilvl)
                            list_counters[key] = list_counters.get(key, 0) + 1
                            prefix = f"{list_counters[key]}. "
                        else:
                            bullets = ["• ", "◦ ", "▪ "]
                            prefix = bullets[min(ilvl, 2)]
                        st = ParagraphStyle(
                            f"ListL{ilvl}", parent=body_style,
                            leftIndent=16 + ilvl * 16, firstLineIndent=-10, spaceAfter=3, alignment=info["align"],
                        )
                        story.append(RLParagraph(f"{prefix}{markup}", st))
                    else:
                        st = ParagraphStyle("BodyDyn", parent=body_style, alignment=info["align"])
                        story.append(RLParagraph(markup, st))

            for _ in range(info["inline_page_breaks"]):
                story.append(PageBreak())

            if info["has_section_break"] and cur_sec_idx + 1 < len(sections_meta):
                cur_sec_idx += 1
                story.append(NextPageTemplate(f"sec_{cur_sec_idx}"))
                story.append(PageBreak())

        elif child.tag.endswith("tbl"):
            table = Table(child, doc)
            raw_rows, span_cmds, repeat_hdr_rows, grid_widths = _extract_docx_table_grid_and_spans(table, doc)
            if not raw_rows:
                continue

            col_widths, tbl_font_pt = _compute_smart_table_col_widths(raw_rows, grid_widths, sm["avail_w"])
            tbl_leading = round(tbl_font_pt * 1.28, 1)

            table_data = []
            ts_cmds = [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#CBD5E1")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 3),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                ("LEFTPADDING", (0, 0), (-1, -1), 3.5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 3.5),
            ]

            for r_idx, row in enumerate(raw_rows):
                row_data = []
                is_hdr = r_idx < repeat_hdr_rows
                for c_idx, cell in enumerate(row):
                    if not cell or not cell["is_origin"]:
                        row_data.append("")
                        continue

                    fill_hex = cell["fill"] or ("0F172A" if is_hdr else ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF"))
                    dark_bg = _is_dark_hex(fill_hex)
                    ts_cmds.append(("BACKGROUND", (c_idx, r_idx), (c_idx, r_idx), colors.HexColor(f"#{fill_hex}")))

                    def_text_hex = "FFFFFF" if dark_bg else "0F172A"
                    cell_paras = []
                    for p_info in cell["paras"]:
                        if p_info["plain_text"].strip():
                            # Override dark default run color when cell fill is dark
                            segs = []
                            for s in p_info["segments"]:
                                s_copy = dict(s)
                                if dark_bg and s_copy.get("color") in (None, "0F172A", "1E293B", "334155", "000000"):
                                    s_copy["color"] = "FFFFFF"
                                segs.append(s_copy)
                            mk = _segments_to_rl_markup(segs, mono_family, default_color=def_text_hex)
                            p_align = p_info["align"]
                            is_b = is_hdr or any(s.get("bold") for s in segs)
                            st = ParagraphStyle(
                                "TblCellDyn",
                                fontName=(f"{sans_family}-Bold" if (is_b and sans_family == "VelmoSans") else sans_family),
                                fontSize=tbl_font_pt,
                                leading=tbl_leading,
                                textColor=colors.HexColor(f"#{def_text_hex}"),
                                alignment=p_align,
                            )
                            cell_paras.append(RLParagraph(mk, st))
                    row_data.append(cell_paras if len(cell_paras) > 1 else (cell_paras[0] if cell_paras else ""))
                table_data.append(row_data)

            for c0, r0, c1, r1 in span_cmds:
                ts_cmds.append(("SPAN", (c0, r0), (c1, r1)))

            if table_data:
                rl_table = RLTable(
                    table_data,
                    colWidths=col_widths,
                    repeatRows=repeat_hdr_rows if len(table_data) > repeat_hdr_rows else 0,
                    splitByRow=1,
                )
                rl_table.setStyle(TableStyle(ts_cmds))
                story.append(rl_table)
                story.append(Spacer(1, 7))

    if not story:
        story.append(RLParagraph("Empty Document", body_style))

    try:
        pdf_doc.build(story, canvasmaker=_VelmoNumberedCanvas)
        return buffer.getvalue()
    finally:
        keepalive_bufs.clear()
        story.clear()
        gc.collect()


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


def _estimate_textbox_height_in(lines_info: List[Tuple[str, float]], width_in: float) -> float:
    """Estimates vertical height (inches) needed by wrapped paragraphs inside a PPTX textbox."""
    total_h = 0.10
    for text, pt in lines_info:
        chars_per_line = max(int(width_in * 10.2 * (12.0 / max(pt, 7.0))), 24)
        sub_lines = text.split("\n")
        wrapped = 0
        for sl in sub_lines:
            wrapped += max(1, (len(sl) + chars_per_line - 1) // chars_per_line)
        line_h_in = (pt * 1.35) / 72.0
        total_h += wrapped * line_h_in + (pt * 0.45) / 72.0
    return total_h


def _render_styled_pptx_table_slice(
    slide,
    raw_rows_slice: List[List[Optional[dict]]],
    span_cmds_slice: List[Tuple[int, int, int, int]],
    repeat_hdr_rows: int,
    left,
    top,
    width,
    max_height,
    col_weights: List[float],
) -> int:
    """Renders a slice of a DOCX table onto a PPTX slide with proportional widths, merges, and fills."""
    n_r = len(raw_rows_slice)
    n_c = len(raw_rows_slice[0]) if n_r else 0
    if n_r == 0 or n_c == 0:
        return 0

    if n_c >= 10:
        font_pt = 8.0
    elif n_c >= 8:
        font_pt = 8.5
    elif n_r > 12:
        font_pt = 8.8
    else:
        font_pt = 9.8

    # Estimate required height per row based on wrapped text lines
    total_w_in = width / Inches(1)
    sum_wt = max(sum(col_weights), 1.0)
    col_widths_in = [total_w_in * (w / sum_wt) for w in col_weights]

    row_heights_in: List[float] = []
    for r in raw_rows_slice:
        max_lines = 1
        for c_i, cell in enumerate(r):
            if not cell or not cell["is_origin"]:
                continue
            txt = " ".join(p["plain_text"].strip() for p in cell["paras"] if p["plain_text"].strip())
            cw_in = max(col_widths_in[c_i] - 0.12, 0.35)
            cpl = max(int(cw_in * 10.0 * (12.0 / font_pt)), 6)
            lines_needed = max(1, (len(txt) + cpl - 1) // cpl)
            max_lines = max(max_lines, lines_needed)
        rh = max(0.26, max_lines * ((font_pt * 1.35) / 72.0) + 0.10)
        row_heights_in.append(rh)

    est_h = min(max_height, Inches(sum(row_heights_in)))
    shape = slide.shapes.add_table(n_r, n_c, left, top, width, est_h)
    tbl = shape.table

    for c_i, cw_in in enumerate(col_widths_in):
        tbl.columns[c_i].width = Inches(cw_in)
    scale_h = min(1.0, (max_height / Inches(1)) / max(sum(row_heights_in), 0.1))
    for r_i, rh in enumerate(row_heights_in):
        tbl.rows[r_i].height = Inches(max(0.22, rh * scale_h))

    for r_i, row in enumerate(raw_rows_slice):
        is_hdr = r_i < repeat_hdr_rows
        for c_i, cell_info in enumerate(row):
            pptx_cell = tbl.cell(r_i, c_i)
            pptx_cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            pptx_cell.margin_top = Pt(1.5)
            pptx_cell.margin_bottom = Pt(1.5)
            pptx_cell.margin_left = Pt(3)
            pptx_cell.margin_right = Pt(3)

            fill_hex = (cell_info.get("fill") if cell_info else None) or (
                "0F172A" if is_hdr else ("F8FAFC" if r_i % 2 == 1 else "FFFFFF")
            )
            dark_bg = _is_dark_hex(fill_hex)
            pptx_cell.fill.solid()
            pptx_cell.fill.fore_color.rgb = RGBColor(
                int(fill_hex[0:2], 16), int(fill_hex[2:4], 16), int(fill_hex[4:6], 16)
            )

            if not cell_info or not cell_info["is_origin"]:
                pptx_cell.text = ""
                continue

            txt = "\n".join(
                p["plain_text"].strip() for p in cell_info["paras"] if p["plain_text"].strip()
            ).strip()
            pptx_cell.text = txt
            is_bold = is_hdr or any(
                s.get("bold") for p in cell_info["paras"] for s in p["segments"]
            )
            fg = RGBColor(255, 255, 255) if dark_bg else RGBColor(15, 23, 42)
            for p in pptx_cell.text_frame.paragraphs:
                p.font.size = Pt(font_pt)
                p.font.bold = is_bold
                p.font.color.rgb = fg

    for c0, r0, c1, r1 in span_cmds_slice:
        try:
            tbl.cell(r0, c0).merge(tbl.cell(r1, c1))
        except Exception:
            pass

    return int(shape.height)


def docx_to_pptx(docx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Creates clean, non-overlapping widescreen PowerPoint slides from Word headings, prose, lists, code blocks, images, and tables."""
    doc = Document(io.BytesIO(docx_bytes))
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    margin_x_in = 0.55
    content_w_in = 12.23
    top_body_in = 1.05
    max_y_in = 7.08

    current_slide = None
    current_title_shape = None
    slide_count = 0
    current_title = "Document Overview"
    cur_y_in = top_body_in
    slide_has_table_or_img = False
    para_buffer: List[dict] = []
    list_counters: Dict[Tuple[str, int], int] = {}

    def start_new_slide(title: str = "Slide", flush: bool = True):
        nonlocal current_slide, current_title_shape, slide_count, current_title, cur_y_in, slide_has_table_or_img
        if flush:
            flush_para_buffer()
        current_slide = prs.slides.add_slide(blank_layout)
        slide_count += 1
        current_title = title.strip() or f"Slide {slide_count}"
        slide_has_table_or_img = False

        t_box = current_slide.shapes.add_textbox(
            Inches(margin_x_in), Inches(0.30), Inches(content_w_in), Inches(0.65)
        )
        current_title_shape = t_box
        t_box.text_frame.word_wrap = True
        p = t_box.text_frame.paragraphs[0]
        p.text = current_title[:120]
        p.font.size = Pt(22 if len(current_title) < 60 else 18)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 23, 42)
        cur_y_in = top_body_in

    def flush_para_buffer():
        nonlocal cur_y_in, current_slide
        if not para_buffer:
            return
        if current_slide is None:
            start_new_slide(current_title, flush=False)

        idx = 0
        while idx < len(para_buffer):
            avail_h_in = max_y_in - cur_y_in
            if avail_h_in < 0.75:
                start_new_slide(f"{current_title} (cont.)" if "(cont.)" not in current_title else current_title, flush=False)
                avail_h_in = max_y_in - cur_y_in

            # Group contiguous items of same box type (plain vs code/callout) that fit into avail_h_in
            first = para_buffer[idx]
            is_box = first["kind"] in ("code", "callout")
            chunk: List[dict] = []
            used_h = 0.12

            while idx < len(para_buffer):
                item = para_buffer[idx]
                item_is_box = item["kind"] in ("code", "callout")
                if chunk and (is_box != item_is_box):
                    break
                item_h = _estimate_textbox_height_in([(item["text"], item["pt"])], content_w_in - (0.3 if item_is_box else 0.0))
                if chunk and used_h + item_h > avail_h_in:
                    break
                chunk.append(item)
                used_h += item_h
                idx += 1
                if item_is_box:
                    break

            box_h_in = min(max(used_h + 0.08, 0.38), avail_h_in)
            tb = current_slide.shapes.add_textbox(
                Inches(margin_x_in), Inches(cur_y_in), Inches(content_w_in), Inches(box_h_in)
            )
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_top = Pt(3)
            tf.margin_bottom = Pt(3)
            tf.margin_left = Pt(6 if is_box else 2)
            tf.margin_right = Pt(6 if is_box else 2)

            if is_box:
                bg_hex = first.get("bg") or ("F1F5F9" if first["kind"] == "code" else "EFF6FF")
                tb.fill.solid()
                tb.fill.fore_color.rgb = RGBColor(int(bg_hex[0:2], 16), int(bg_hex[2:4], 16), int(bg_hex[4:6], 16))

            for p_i, item in enumerate(chunk):
                p = tf.paragraphs[0] if p_i == 0 else tf.add_paragraph()
                p.space_after = Pt(4)
                p.level = item.get("level", 0)
                if item.get("prefix"):
                    r_pre = p.add_run()
                    r_pre.text = item["prefix"]
                    r_pre.font.size = Pt(item["pt"])
                    r_pre.font.bold = True
                    r_pre.font.color.rgb = RGBColor(30, 41, 59)

                for seg in item["segments"]:
                    if not seg["text"]:
                        continue
                    r = p.add_run()
                    r.text = seg["text"]
                    r.font.size = Pt(item["pt"])
                    r.font.bold = bool(item.get("bold") or seg.get("bold"))
                    r.font.italic = bool(seg.get("italic"))
                    if item["kind"] == "code" or seg.get("mono"):
                        r.font.name = "Consolas"
                    if seg.get("url"):
                        r.hyperlink.address = seg["url"]
                        r.font.color.rgb = RGBColor(37, 99, 235)
                        r.font.underline = True
                    elif seg.get("color") and seg["color"] not in ("000000", "AUTO"):
                        hx = seg["color"]
                        r.font.color.rgb = RGBColor(int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))
                    else:
                        r.font.color.rgb = RGBColor(15, 23, 42) if item.get("bold") else RGBColor(51, 65, 85)

            cur_y_in += box_h_in + 0.08

        para_buffer.clear()

    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            info = _extract_oxml_paragraph_data(child, doc)
            text = info["plain_text"].strip()

            if info["page_break_before"] and current_slide is not None:
                start_new_slide(current_title)

            # Handle embedded images in paragraph
            for rid, cx_pt, cy_pt in info["images"]:
                rel_part = doc.part.related_parts.get(rid)
                if rel_part is None:
                    continue
                flush_para_buffer()
                if current_slide is None:
                    start_new_slide(current_title)
                avail_h_in = max_y_in - cur_y_in
                if avail_h_in < 2.2 or slide_has_table_or_img:
                    start_new_slide(current_title)
                    avail_h_in = max_y_in - cur_y_in

                try:
                    flat_bytes, iw, ih = _flatten_image_for_reportlab(rel_part.blob)
                    target_w_in = (cx_pt / 72.0) if cx_pt else content_w_in * 0.8
                    target_h_in = (cy_pt / 72.0) if cy_pt else target_w_in * (ih / iw)
                    scale = min(content_w_in / max(target_w_in, 0.1), avail_h_in / max(target_h_in, 0.1), 1.35)
                    # Cap upscaling so small logos stay crisp and large plates fit
                    if target_h_in * scale > avail_h_in:
                        scale = avail_h_in / max(target_h_in, 0.1)
                    if target_w_in * scale > content_w_in:
                        scale = content_w_in / max(target_w_in, 0.1)
                    box_w = Inches(min(content_w_in, max(target_w_in * scale, 1.0)))
                    box_h = Inches(min(avail_h_in, max(target_h_in * scale, 0.7)))
                    box_left = Inches(margin_x_in) + int((Inches(content_w_in) - box_w) / 2)
                    _add_aspect_fit_picture(current_slide, flat_bytes, box_left, Inches(cur_y_in), box_w, box_h)
                    cur_y_in += (box_h / Inches(1)) + 0.10
                    slide_has_table_or_img = True
                except Exception as img_err:
                    logger.warning(f"DOCX->PPTX image render failed for {rid}: {img_err}")

            if text:
                hl = info["heading_level"]
                if hl == 1:
                    if current_slide is not None and (cur_y_in > top_body_in + 0.05 or para_buffer or slide_has_table_or_img):
                        start_new_slide(text)
                    else:
                        if current_slide is None:
                            start_new_slide(text)
                        else:
                            # Replace empty slide title
                            current_title = text
                            if current_title_shape is not None:
                                current_title_shape.text_frame.paragraphs[0].text = text[:120]
                elif hl == 2:
                    flush_para_buffer()
                    if current_slide is None or cur_y_in > 2.35 or slide_has_table_or_img:
                        start_new_slide(text)
                    else:
                        para_buffer.append({
                            "kind": "heading",
                            "text": text,
                            "segments": info["segments"],
                            "pt": 15.0,
                            "bold": True,
                            "level": 0,
                        })
                elif hl >= 3:
                    para_buffer.append({
                        "kind": "heading",
                        "text": text,
                        "segments": info["segments"],
                        "pt": 13.0 if hl == 3 else 11.5,
                        "bold": True,
                        "level": 0,
                    })
                elif info["is_mono_block"]:
                    para_buffer.append({
                        "kind": "code",
                        "text": info["plain_text"].strip("\n"),
                        "segments": info["segments"],
                        "pt": 9.5,
                        "bold": False,
                        "bg": "F1F5F9",
                    })
                elif info["is_callout"]:
                    para_buffer.append({
                        "kind": "callout",
                        "text": text,
                        "segments": info["segments"],
                        "pt": 11.0,
                        "bold": False,
                        "bg": info["callout_bg"] or "EFF6FF",
                    })
                elif info["list_info"] is not None:
                    li = info["list_info"]
                    ilvl = li["ilvl"]
                    if li["ordered"]:
                        key = (li["num_id"], ilvl)
                        list_counters[key] = list_counters.get(key, 0) + 1
                        prefix = f"{list_counters[key]}. "
                    else:
                        prefix = "• " if ilvl == 0 else ("◦ " if ilvl == 1 else "▪ ")
                    para_buffer.append({
                        "kind": "list",
                        "text": text,
                        "segments": info["segments"],
                        "prefix": prefix,
                        "pt": 11.5,
                        "bold": False,
                        "level": min(ilvl, 2),
                    })
                else:
                    para_buffer.append({
                        "kind": "body",
                        "text": text,
                        "segments": info["segments"],
                        "pt": 11.5,
                        "bold": False,
                        "level": 0,
                    })

            if (info["inline_page_breaks"] > 0 or info["has_section_break"]) and current_slide is not None:
                flush_para_buffer()
                # Prepare cleanly for next slide only if more content follows
                current_slide = None

        elif child.tag.endswith('tbl'):
            flush_para_buffer()
            tbl = Table(child, doc)
            raw_rows, span_cmds, repeat_hdr_rows, grid_widths = _extract_docx_table_grid_and_spans(tbl, doc)
            if not raw_rows:
                continue

            if current_slide is None:
                start_new_slide(current_title)

            avail_h_in = max_y_in - cur_y_in
            if avail_h_in < 1.8 or slide_has_table_or_img:
                start_new_slide(current_title)
                avail_h_in = max_y_in - cur_y_in

            num_cols = len(raw_rows[0])
            col_weights = []
            for c_i in range(num_cols):
                max_len = 4
                for r in raw_rows:
                    c = r[c_i]
                    if c and c["is_origin"] and c["grid_span"] == 1:
                        t = " ".join(p["plain_text"].strip() for p in c["paras"]).strip()
                        max_len = max(max_len, min(len(t), 34))
                if len(grid_widths) == num_cols and grid_widths[c_i] > 0:
                    col_weights.append(0.6 * max_len + 0.4 * (grid_widths[c_i] / 6.0))
                else:
                    col_weights.append(float(max_len))

            # Paginate tall tables cleanly across slides (e.g., 38-row register tables)
            hdr_slice = raw_rows[:repeat_hdr_rows] if repeat_hdr_rows > 0 else []
            data_rows = raw_rows[repeat_hdr_rows:] if repeat_hdr_rows > 0 else raw_rows
            hdr_spans = [(c0, r0, c1, r1) for (c0, r0, c1, r1) in span_cmds if r1 < repeat_hdr_rows]

            if not data_rows:
                data_rows = raw_rows
                hdr_slice = []
                hdr_spans = []
                effective_hdr_cnt = 0
            else:
                effective_hdr_cnt = len(hdr_slice)

            max_data_per_slide = 11 if num_cols <= 7 else 10
            if len(data_rows) <= max_data_per_slide + 2:
                # Single table slide
                if avail_h_in < min(max_y_in - top_body_in, len(raw_rows) * 0.34 + 0.3) and cur_y_in > top_body_in + 0.5:
                    start_new_slide(current_title)
                    avail_h_in = max_y_in - cur_y_in
                actual_h_emu = _render_styled_pptx_table_slice(
                    current_slide,
                    raw_rows,
                    span_cmds,
                    effective_hdr_cnt,
                    Inches(margin_x_in),
                    Inches(cur_y_in),
                    Inches(content_w_in),
                    Inches(avail_h_in),
                    col_weights,
                )
                cur_y_in += (actual_h_emu / Inches(1)) + 0.14
                slide_has_table_or_img = True
            else:
                # Split across slides with repeating headers
                base_title = current_title
                start_idx = 0
                part_no = 1
                while start_idx < len(data_rows):
                    if part_no > 1 or avail_h_in < 3.5:
                        title_lbl = base_title if part_no == 1 else f"{base_title} (Part {part_no})"
                        start_new_slide(title_lbl)
                        avail_h_in = max_y_in - cur_y_in
                    end_idx = min(start_idx + max_data_per_slide, len(data_rows))
                    chunk_rows = hdr_slice + data_rows[start_idx:end_idx]
                    chunk_spans = list(hdr_spans)
                    # Include spans wholly inside [start_idx, end_idx)
                    orig_start = effective_hdr_cnt + start_idx
                    orig_end = effective_hdr_cnt + end_idx - 1
                    for c0, r0, c1, r1 in span_cmds:
                        if r0 >= orig_start and r1 <= orig_end:
                            chunk_spans.append((c0, r0 - start_idx, c1, r1 - start_idx))

                    actual_h_emu = _render_styled_pptx_table_slice(
                        current_slide,
                        chunk_rows,
                        chunk_spans,
                        effective_hdr_cnt,
                        Inches(margin_x_in),
                        Inches(cur_y_in),
                        Inches(content_w_in),
                        Inches(avail_h_in),
                        col_weights,
                    )
                    cur_y_in += (actual_h_emu / Inches(1)) + 0.14
                    slide_has_table_or_img = True
                    start_idx = end_idx
                    part_no += 1

    flush_para_buffer()

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


def _add_aspect_fit_picture(slide, img_bytes: bytes, box_left, box_top, box_w, box_h):
    """Places an image centered inside the target box while strictly preserving aspect ratio."""
    with Image.open(io.BytesIO(img_bytes)) as pil_img:
        iw, ih = max(pil_img.width, 1), max(pil_img.height, 1)
    scale = min(box_w / iw, box_h / ih)
    draw_w = int(iw * scale)
    draw_h = int(ih * scale)
    left = int(box_left + (box_w - draw_w) / 2)
    top = int(box_top + (box_h - draw_h) / 2)
    slide.shapes.add_picture(io.BytesIO(img_bytes), left, top, draw_w, draw_h)


def _add_styled_pptx_table(slide, rows: List[List[str]], left, top, width, max_height):
    """Renders a complete non-clipped styled table onto a slide without row/column truncation."""
    if not rows:
        return
    n_r = len(rows)
    n_c = max(len(r) for r in rows)
    if n_r == 0 or n_c == 0:
        return

    if n_r <= 8 and n_c <= 6:
        font_pt = 10.5
    elif n_r <= 14 and n_c <= 9:
        font_pt = 8.5
    elif n_r <= 22:
        font_pt = 7.5
    else:
        font_pt = 6.5

    est_h = min(max_height, Inches(max(1.2, n_r * (0.24 if n_r > 14 else 0.36))))
    shape = slide.shapes.add_table(n_r, n_c, left, top, width, est_h)
    tbl = shape.table

    for r_i, row in enumerate(rows):
        is_hdr = (r_i == 0)
        row_bg = RGBColor(15, 23, 42) if is_hdr else (RGBColor(248, 250, 252) if r_i % 2 == 1 else RGBColor(255, 255, 255))
        for c_i in range(n_c):
            cell = tbl.cell(r_i, c_i)
            val = str(row[c_i]).strip() if c_i < len(row) and row[c_i] is not None else ""
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_top = Pt(1.5)
            cell.margin_bottom = Pt(1.5)
            cell.margin_left = Pt(3)
            cell.margin_right = Pt(3)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_bg
            for p in cell.text_frame.paragraphs:
                p.font.size = Pt(font_pt)
                p.font.bold = is_hdr
                p.font.color.rgb = RGBColor(255, 255, 255) if is_hdr else RGBColor(15, 23, 42)


def _add_fitted_textbox(slide, lines: List[str], left, top, width, height, max_pt: float = 12.0, min_pt: float = 7.5):
    """Places all text lines inside a bounded textbox, scaling font smoothly to prevent clipping."""
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_top = Pt(2)
    tf.margin_bottom = Pt(2)
    tf.margin_left = Pt(4)
    tf.margin_right = Pt(4)

    width_in = max(width / Inches(1), 1.0)
    height_in = max(height / Inches(1), 0.5)

    chosen_pt = min_pt
    for pt in [max_pt, 11.0, 10.0, 9.0, 8.2, 7.5]:
        if pt > max_pt:
            continue
        chars_per_line = max(int(width_in * 10.5 * (12.0 / pt)), 20)
        est_lines = sum(max(1, (len(ln) + chars_per_line - 1) // chars_per_line) for ln in lines)
        line_h_in = (pt * 1.32) / 72.0
        if est_lines * line_h_in <= height_in * 0.96:
            chosen_pt = pt
            break

    for idx, ln in enumerate(lines):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = ln
        p.font.size = Pt(chosen_pt)
        p.font.color.rgb = RGBColor(30, 41, 59)
        p.space_after = Pt(2 if chosen_pt <= 9.5 else 4)


def pdf_to_pptx(pdf_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Creates clean, non-overlapping 16:9 presentation slides per PDF page with full tables and OCR fallback."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Invalid or corrupted PDF file: {e}")

    try:
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_layout = prs.slide_layouts[6]

        margin_x = Inches(0.55)
        content_w = Inches(12.23)

        for p_no, page in enumerate(doc, start=1):
            slide = prs.slides.add_slide(blank_layout)

            if _is_scanned_pdf_page(page):
                items = _ocr_pdf_page(page, base_dpi=150)
                blocks = _structure_ocr_blocks(items)
                pix = page.get_pixmap(dpi=130, alpha=False)
                scan_thumb = pix.tobytes("png")
                del pix

                if blocks:
                    headings = [b["text"] for b in blocks if b["type"] == "heading"]
                    tables = [b["rows"] for b in blocks if b["type"] == "table"]
                    paras = [b["text"] for b in blocks if b["type"] == "paragraph"]

                    title_txt = headings[0] if headings else f"Page {p_no} (Scanned)"
                    t_box = slide.shapes.add_textbox(margin_x, Inches(0.30), content_w, Inches(0.60))
                    t_p = t_box.text_frame.paragraphs[0]
                    t_p.text = title_txt[:110]
                    t_p.font.size = Pt(20)
                    t_p.font.bold = True
                    t_p.font.color.rgb = RGBColor(15, 23, 42)

                    ocr_lines = headings[1:] + paras
                    if tables:
                        if ocr_lines:
                            _add_fitted_textbox(slide, ocr_lines, margin_x, Inches(0.95), content_w, Inches(1.25), max_pt=11.0)
                            _add_styled_pptx_table(slide, tables[0], margin_x, Inches(2.30), content_w, Inches(4.85))
                        else:
                            _add_styled_pptx_table(slide, tables[0], margin_x, Inches(1.05), content_w, Inches(6.05))
                    else:
                        if ocr_lines:
                            _add_fitted_textbox(slide, ocr_lines, margin_x, Inches(0.98), Inches(7.40), Inches(6.15), max_pt=12.0)
                        _add_aspect_fit_picture(slide, scan_thumb, Inches(8.15), Inches(0.98), Inches(4.63), Inches(6.15))

                    # Also store full OCR transcript in speaker notes
                    all_ocr_text = "\n".join(
                        [b["text"] if b["type"] != "table" else "\n".join(" | ".join(r) for r in b["rows"]) for b in blocks]
                    )
                    slide.notes_slide.notes_text_frame.text = all_ocr_text
                else:
                    _add_aspect_fit_picture(slide, scan_thumb, margin_x, Inches(0.40), content_w, Inches(6.70))

                del scan_thumb
                gc.collect()
                continue

            # Digital page
            tabs = _find_pdf_tables(page)
            valid_tables: List[List[List[str]]] = []
            table_bboxes = []
            for tab in tabs:
                extracted = _extract_pdf_table_grid(page, tab)
                if not extracted or len(extracted) < 2:
                    continue
                valid_tables.append(extracted)
                table_bboxes.append(tab.bbox)

            raw_blocks = page.get_text("blocks")
            non_table_lines: List[str] = []
            for b in raw_blocks:
                if len(b) > 6 and b[6] != 0:
                    continue
                b_rect = fitz.Rect(b[0], b[1], b[2], b[3])
                b_text = re.sub(r"\s*\n\s*", " ", b[4].strip())
                if not b_text:
                    continue
                b_area = max(b_rect.get_area(), 1.0)
                in_table = any(((b_rect & fitz.Rect(tb)).get_area() / b_area) > 0.35 for tb in table_bboxes)
                if not in_table:
                    non_table_lines.append(b_text)

            has_images = len(page.get_images()) > 0
            title_txt = f"Page {p_no}"
            for ln in non_table_lines:
                if 6 <= len(ln) <= 95 and "CONFIDENTIAL" not in ln and "Symbols:" not in ln and "Currencies:" not in ln:
                    title_txt = ln
                    break

            t_box = slide.shapes.add_textbox(margin_x, Inches(0.30), content_w, Inches(0.58))
            t_box.text_frame.word_wrap = True
            tp = t_box.text_frame.paragraphs[0]
            tp.text = f"Page {p_no} — {title_txt}" if title_txt != f"Page {p_no}" else title_txt
            tp.font.size = Pt(18)
            tp.font.bold = True
            tp.font.color.rgb = RGBColor(15, 23, 42)

            if valid_tables:
                primary_tbl = valid_tables[0]
                n_rows = len(primary_tbl)
                if n_rows >= 14:
                    prose_h = Inches(1.15)
                    tbl_top = Inches(2.22)
                    tbl_max_h = Inches(4.95)
                elif has_images:
                    prose_h = Inches(1.90)
                    tbl_top = Inches(2.98)
                    tbl_max_h = Inches(4.18)
                else:
                    prose_h = Inches(1.85)
                    tbl_top = Inches(2.93)
                    tbl_max_h = Inches(4.22)

                if has_images and n_rows < 14:
                    if non_table_lines:
                        _add_fitted_textbox(slide, non_table_lines, margin_x, Inches(0.95), Inches(8.15), prose_h, max_pt=10.5)
                    pix = page.get_pixmap(dpi=120, alpha=False)
                    snap = pix.tobytes("png")
                    del pix
                    _add_aspect_fit_picture(slide, snap, Inches(8.90), Inches(0.95), Inches(3.88), prose_h)
                    del snap
                elif non_table_lines:
                    _add_fitted_textbox(slide, non_table_lines, margin_x, Inches(0.95), content_w, prose_h, max_pt=10.5)
                else:
                    tbl_top = Inches(1.00)
                    tbl_max_h = Inches(6.15)

                _add_styled_pptx_table(slide, primary_tbl, margin_x, tbl_top, content_w, tbl_max_h)

            elif non_table_lines or has_images:
                if has_images and non_table_lines:
                    _add_fitted_textbox(slide, non_table_lines, margin_x, Inches(0.98), Inches(7.35), Inches(6.15), max_pt=12.0)
                    pix = page.get_pixmap(dpi=130, alpha=False)
                    snap = pix.tobytes("png")
                    del pix
                    _add_aspect_fit_picture(slide, snap, Inches(8.10), Inches(0.98), Inches(4.68), Inches(6.15))
                    del snap
                elif non_table_lines:
                    _add_fitted_textbox(slide, non_table_lines, margin_x, Inches(0.98), content_w, Inches(6.15), max_pt=12.5)
                else:
                    pix = page.get_pixmap(dpi=140, alpha=False)
                    snap = pix.tobytes("png")
                    del pix
                    _add_aspect_fit_picture(slide, snap, margin_x, Inches(0.98), content_w, Inches(6.15))
                    del snap

        out_io = io.BytesIO()
        prs.save(out_io)
        return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation", f"{base_name}.pptx", len(prs.slides)
    finally:
        doc.close()
        _release_pdf_memory()


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

def _segments_to_txt(segments: List[Dict[str, Any]]) -> str:
    """Renders paragraph segments to plain text, preserving link URLs inline as 'anchor (url)'."""
    merged: List[Dict[str, Any]] = []
    for seg in segments:
        txt = seg.get("text", "")
        if not txt:
            continue
        if merged and seg.get("url") is not None and merged[-1].get("url") == seg.get("url"):
            merged[-1]["text"] += txt
        else:
            merged.append(dict(seg))

    out: List[str] = []
    for seg in merged:
        raw = seg["text"]
        if seg.get("url"):
            lbl = raw.strip()
            lead = raw[: len(raw) - len(raw.lstrip())]
            trail = raw[len(raw.rstrip()):]
            if lbl and seg["url"] not in lbl:
                out.append(f"{lead}{lbl} ({seg['url']}){trail}")
            else:
                out.append(raw)
        else:
            out.append(raw)
    return "".join(out)


def _docx_to_txt(docx_bytes: bytes) -> str:
    """Extracts clean, reading-order plain text from a DOCX file including headers, links, lists, and tables."""
    doc = Document(io.BytesIO(docx_bytes))
    parts: List[str] = []
    list_counters: Dict[Tuple[str, int], int] = {}

    # Collect unique non-empty headers
    seen_hf: Set[str] = set()
    for sec in doc.sections:
        if sec.header:
            for p in sec.header.paragraphs:
                t = _extract_oxml_paragraph_data(p._p, doc)["plain_text"].strip()
                t = re.sub(r"__VELMO_(?:PAGE|NUMPAGES)__", "", t).strip(" |\t-")
                if t and t not in seen_hf:
                    seen_hf.add(t)
                    parts.append(t)
                    parts.append("=" * min(len(t), 72))

    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            info = _extract_oxml_paragraph_data(child, doc)
            plain = info["plain_text"].strip()
            if info["page_break_before"] and parts:
                parts.append("\n" + "-" * 40)

            if plain:
                txt_line = _segments_to_txt(info["segments"]).strip()
                if info["heading_level"] > 0:
                    parts.append(f"\n{txt_line}")
                elif info["is_mono_block"]:
                    parts.append("\n" + info["plain_text"].strip("\n") + "\n")
                elif info["list_info"] is not None:
                    li = info["list_info"]
                    ilvl = li["ilvl"]
                    indent = "  " * ilvl
                    if li["ordered"]:
                        key = (li["num_id"], ilvl)
                        list_counters[key] = list_counters.get(key, 0) + 1
                        parts.append(f"{indent}{list_counters[key]}. {txt_line}")
                    else:
                        parts.append(f"{indent}- {txt_line}")
                else:
                    parts.append(txt_line)

            if info["inline_page_breaks"] > 0 or info["has_section_break"]:
                parts.append("-" * 40)

        elif child.tag.endswith('tbl'):
            tbl = Table(child, doc)
            raw_rows, _, _, _ = _extract_docx_table_grid_and_spans(tbl, doc)
            if raw_rows:
                parts.append("")
                for r in raw_rows:
                    row_vals = []
                    for c in r:
                        if not c or not c["is_origin"]:
                            row_vals.append("")
                        else:
                            c_txt = " ".join(
                                _segments_to_txt(p["segments"]).strip()
                                for p in c["paras"]
                                if p["plain_text"].strip()
                            ).strip()
                            row_vals.append(re.sub(r"\s*\n\s*", " ", c_txt))
                    while row_vals and row_vals[-1] == "":
                        row_vals.pop()
                    if row_vals:
                        parts.append("\t".join(row_vals))
                parts.append("")

    for sec in doc.sections:
        if sec.footer:
            for p in sec.footer.paragraphs:
                t = _extract_oxml_paragraph_data(p._p, doc)["plain_text"].strip()
                t = re.sub(r"Page\s*__VELMO_PAGE__\s*of\s*__VELMO_NUMPAGES__", "", t, flags=re.I)
                t = re.sub(r"__VELMO_(?:PAGE|NUMPAGES)__", "", t).strip(" |\t")
                if t and t not in seen_hf:
                    seen_hf.add(t)
                    parts.append(f"\n[{t}]")

    return "\n".join(parts).strip() + "\n"


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
        try:
            page_chunks: List[str] = []
            for i, page in enumerate(doc, start=1):
                if _is_scanned_pdf_page(page):
                    items = _ocr_pdf_page(page, base_dpi=150)
                    blocks = _structure_ocr_blocks(items)
                    if blocks:
                        lines_out: List[str] = []
                        for blk in blocks:
                            if blk["type"] == "table" and blk.get("rows"):
                                for row in blk["rows"]:
                                    lines_out.append("\t".join(str(c).strip() for c in row))
                            elif blk.get("text"):
                                lines_out.append(blk["text"])
                        page_txt = "\n".join(lines_out)
                    else:
                        page_txt = page.get_text() or "(No text detected on scanned page)"
                else:
                    page_txt = page.get_text()
                page_chunks.append(f"--- Page {i} ---\n{page_txt.strip()}")
            text = "\n\n".join(page_chunks)
        finally:
            doc.close()
            _release_pdf_memory()

    elif input_fmt == "docx":
        text = _docx_to_txt(data_bytes)

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


def _format_md_table(rows: List[List[str]]) -> str:
    """Formats a 2D cell grid into a valid GFM Markdown pipe table."""
    if not rows:
        return ""
    max_cols = max(len(r) for r in rows)
    if max_cols == 0:
        return ""

    def clean_cell(val: Any) -> str:
        s = "" if val is None else str(val).strip()
        s = re.sub(r"\s*\n\s*", " ", s)
        return s.replace("|", "\\|")

    lines: List[str] = []
    hdr = [clean_cell(rows[0][c] if c < len(rows[0]) else "") for c in range(max_cols)]
    lines.append("| " + " | ".join(hdr) + " |")
    lines.append("| " + " | ".join(["---"] * max_cols) + " |")
    for row in rows[1:]:
        cells = [clean_cell(row[c] if c < len(row) else "") for c in range(max_cols)]
        lines.append("| " + " | ".join(cells) + " |")
    return "\n".join(lines)


def _is_mono_font(font_name: str) -> bool:
    fn = (font_name or "").lower()
    return any(k in fn for k in ("consola", "courier", "mono", "menlo", "cascadia", "fira", "inconsolata", "dejavusansmono", "liberationmono"))


def _render_pdf_block_text_with_links(blk: dict, page_links: List[Tuple[fitz.Rect, str]]) -> Tuple[str, float, bool, bool]:
    """
    Reconstructs text from a PyMuPDF dict block while preserving clickable hyperlinks
    and detecting dominant font size, bold status, and monospace code blocks.
    Returns: (text, dom_size, all_bold, is_monospace)
    """
    mono_chars = 0
    total_chars = 0
    size_weights: Dict[float, int] = {}
    bold_chars = 0
    nonspace_chars = 0

    raw_lines: List[str] = []
    for ln in blk.get("lines", []):
        spans = ln.get("spans", [])
        if not spans:
            continue

        # Build line segments merging consecutive spans sharing the same link URI
        segments: List[Tuple[Optional[str], str, bool]] = []  # (uri, text, is_bold)
        for sp in spans:
            stxt = (sp.get("text") or "").replace("\xa0", " ")
            if not stxt:
                continue
            stripped_len = len(stxt.strip())
            total_chars += len(stxt)
            nonspace_chars += stripped_len
            if _is_mono_font(sp.get("font", "")):
                mono_chars += stripped_len
            sz = round(float(sp.get("size", 10.0)), 1)
            size_weights[sz] = size_weights.get(sz, 0) + max(stripped_len, 1)
            is_bold = bool(sp.get("flags", 0) & 16) or ("bold" in (sp.get("font", "").lower()))
            if is_bold:
                bold_chars += stripped_len

            sp_rect = fitz.Rect(sp["bbox"])
            sp_area = max(sp_rect.get_area(), 1.0)
            matched_uri = None
            for lrect, uri in page_links:
                if ((sp_rect & lrect).get_area() / sp_area) > 0.35:
                    matched_uri = uri
                    break

            if segments and matched_uri is not None and segments[-1][0] == matched_uri:
                prev_uri, prev_txt, prev_b = segments[-1]
                segments[-1] = (prev_uri, prev_txt + stxt, prev_b and is_bold)
            else:
                segments.append((matched_uri, stxt, is_bold))

        line_parts = []
        for uri, seg_txt, _ in segments:
            if uri:
                clean_lbl = seg_txt.strip()
                if clean_lbl:
                    lead_ws = seg_txt[: len(seg_txt) - len(seg_txt.lstrip())]
                    trail_ws = seg_txt[len(seg_txt.rstrip()):]
                    line_parts.append(f"{lead_ws}[{clean_lbl}]({uri}){trail_ws}")
                else:
                    line_parts.append(seg_txt)
            else:
                line_parts.append(seg_txt)
        raw_lines.append("".join(line_parts).rstrip())

    is_mono = (nonspace_chars > 0 and (mono_chars / nonspace_chars) >= 0.70)
    all_bold = (nonspace_chars > 0 and (bold_chars / nonspace_chars) >= 0.80)
    dom_size = max(size_weights.items(), key=lambda kv: kv[1])[0] if size_weights else 10.0

    if is_mono:
        return "\n".join(raw_lines).strip("\n"), dom_size, all_bold, True

    # Join lines smartly: keep distinct list/bullet/footnote items on separate lines
    joined_lines: List[str] = []
    for ln in raw_lines:
        s = ln.strip()
        if not s:
            continue
        is_list_start = bool(re.match(r"^([•◦▪\-]|\d+[\.\)]|[a-zA-Z]\)|\[\d+\])\s+", s))
        if not joined_lines:
            joined_lines.append(s)
        elif is_list_start or joined_lines[-1].endswith(":"):
            joined_lines.append(s)
        else:
            joined_lines[-1] = f"{joined_lines[-1]} {s}"

    return "\n".join(joined_lines), dom_size, all_bold, False


def _order_page_blocks_reading_order(blocks: List[dict], page_w: float) -> List[dict]:
    """Orders text blocks respecting 2-column page layouts before falling back to top-to-bottom."""
    if len(blocks) < 4:
        return sorted(blocks, key=lambda b: (round(b["bbox"][1], 1), round(b["bbox"][0], 1)))

    mid_x = page_w / 2.0
    full_top = []
    left_col = []
    right_col = []
    other = []

    for b in blocks:
        x0, y0, x1, y1 = b["bbox"]
        if x0 < mid_x - 20 and x1 > mid_x + 20:
            other.append(b)
        elif x1 <= mid_x + 12:
            left_col.append(b)
        elif x0 >= mid_x - 12:
            right_col.append(b)
        else:
            other.append(b)

    # Check if there is genuine vertical overlap between left and right columns (2-column article)
    if len(left_col) >= 2 and len(right_col) >= 2 and len(other) <= 2:
        l_min_y = min(b["bbox"][1] for b in left_col)
        l_max_y = max(b["bbox"][3] for b in left_col)
        r_min_y = min(b["bbox"][1] for b in right_col)
        r_max_y = max(b["bbox"][3] for b in right_col)
        vert_overlap = min(l_max_y, r_max_y) - max(l_min_y, r_min_y)
        if vert_overlap > 60.0:
            col_top = min(l_min_y, r_min_y)
            full_top = sorted([b for b in other if b["bbox"][1] < col_top], key=lambda b: b["bbox"][1])
            full_bot = sorted([b for b in other if b["bbox"][1] >= col_top], key=lambda b: b["bbox"][1])
            left_col.sort(key=lambda b: (round(b["bbox"][1], 1), round(b["bbox"][0], 1)))
            right_col.sort(key=lambda b: (round(b["bbox"][1], 1), round(b["bbox"][0], 1)))
            return full_top + left_col + right_col + full_bot

    return sorted(blocks, key=lambda b: (round(b["bbox"][1], 1), round(b["bbox"][0], 1)))


def pdf_to_md(pdf_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Extracts structured Markdown directly from PDF text spans, links, tables, code blocks, and scanned-page OCR."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        raise ValueError(f"Invalid or corrupted PDF file: {e}")

    try:
        total_pages = len(doc)
        if total_pages == 0:
            raise ValueError("PDF document has no pages.")

        # Pass 1: Cache digital page tables and determine median prose body font size
        page_tables_cache: Dict[int, Tuple[List[Tuple[float, List[List[str]]]], List[Any]]] = {}
        size_hist: Dict[float, int] = {}
        for p_idx, page in enumerate(doc):
            if _is_scanned_pdf_page(page):
                continue
            tabs = _find_pdf_tables(page)
            table_items: List[Tuple[float, List[List[str]]]] = []
            table_bboxes = []
            for tab in tabs:
                extracted = _extract_pdf_table_grid(page, tab)
                if not extracted or len(extracted) < 2:
                    continue
                table_items.append((tab.bbox[1], extracted))
                table_bboxes.append(tab.bbox)
            table_items.sort(key=lambda item: item[0])
            page_tables_cache[p_idx] = (table_items, table_bboxes)

            ph = page.rect.height
            d = page.get_text("dict")
            for b in d.get("blocks", []):
                if b.get("type") != 0:
                    continue
                b_rect = fitz.Rect(b["bbox"])
                if b_rect.y1 <= 36 or b_rect.y0 >= ph - 36:
                    continue
                b_area = max(b_rect.get_area(), 1.0)
                if any(((b_rect & fitz.Rect(tb)).get_area() / b_area) > 0.35 for tb in table_bboxes):
                    continue
                for ln in b.get("lines", []):
                    for sp in ln.get("spans", []):
                        if _is_mono_font(sp.get("font", "")):
                            continue
                        t = (sp.get("text") or "").strip()
                        if t:
                            sz = round(float(sp.get("size", 10.0)), 1)
                            size_hist[sz] = size_hist.get(sz, 0) + len(t)

        body_size = max(size_hist.items(), key=lambda kv: kv[1])[0] if size_hist else 10.0

        md_parts: List[str] = []
        seen_running_hf: Set[str] = set()

        for p_no, page in enumerate(doc, start=1):
            if p_no > 1:
                md_parts.append("\n---\n")

            if _is_scanned_pdf_page(page):
                items = _ocr_pdf_page(page, base_dpi=150)
                blocks = _structure_ocr_blocks(items)
                if not blocks:
                    md_parts.append(f"*(Scanned Page {p_no}: no text detected)*\n")
                    continue
                for blk in blocks:
                    if blk["type"] == "heading":
                        prefix = "#" if blk.get("level", 2) == 1 else "##"
                        md_parts.append(f"{prefix} {blk['text']}\n")
                    elif blk["type"] == "paragraph":
                        md_parts.append(f"{blk['text']}\n")
                    elif blk["type"] == "table" and blk.get("rows"):
                        md_parts.append(_format_md_table(blk["rows"]) + "\n")
                continue

            pw, ph = page.rect.width, page.rect.height
            page_links = [(fitz.Rect(l["from"]), l["uri"]) for l in page.get_links() if l.get("uri")]
            table_items, table_bboxes = page_tables_cache.get(p_no - 1, ([], []))

            d = page.get_text("dict")
            prose_blocks: List[dict] = []
            hf_lines: List[str] = []

            for b in d.get("blocks", []):
                if b.get("type") != 0:
                    continue
                b_rect = fitz.Rect(b["bbox"])
                b_area = max(b_rect.get_area(), 1.0)
                if any(((b_rect & fitz.Rect(tb)).get_area() / b_area) > 0.35 for tb in table_bboxes):
                    continue
                if b_rect.y1 <= 36 or b_rect.y0 >= ph - 36:
                    hf_txt, _, _, _ = _render_pdf_block_text_with_links(b, page_links)
                    hf_clean = re.sub(r"\s+", " ", hf_txt).strip()
                    if hf_clean:
                        hf_lines.append(hf_clean)
                    continue
                prose_blocks.append(b)

            # Include unique header on Page 1; deduplicate repeated headers across pages
            for hf in hf_lines:
                if hf not in seen_running_hf and not re.search(r"Page\s+\d+\s*(of|/)\s*\d+", hf, re.I):
                    seen_running_hf.add(hf)
                    md_parts.append(f"*{hf}*\n")

            ordered_blocks = _order_page_blocks_reading_order(prose_blocks, pw)
            t_idx = 0
            mono_buf: List[str] = []

            def flush_mono():
                if mono_buf:
                    md_parts.append("```\n" + "\n\n".join(mono_buf) + "\n```\n")
                    mono_buf.clear()

            for b in ordered_blocks:
                by0 = b["bbox"][1]
                # Emit any tables situated above this block
                while t_idx < len(table_items) and table_items[t_idx][0] <= by0:
                    flush_mono()
                    md_parts.append(_format_md_table(table_items[t_idx][1]) + "\n")
                    t_idx += 1

                txt, dom_size, all_bold, is_mono = _render_pdf_block_text_with_links(b, page_links)
                if not txt.strip():
                    continue

                if is_mono:
                    mono_buf.append(txt)
                    continue

                flush_mono()

                # Classify headings vs lists vs paragraphs
                single_line = "\n" not in txt.strip()
                clean_s = txt.strip()
                ratio = dom_size / max(body_size, 1.0)
                is_numbered_h = bool(re.match(r"^\d+(\.\d+){0,4}\.?\s+[A-Z]", clean_s))
                looks_like_bullet = bool(re.match(r"^([•◦▪]|\[\d+\]|_+)", clean_s))
                looks_like_kpi_val = (
                    len(clean_s) <= 18
                    and bool(re.match(r"^[+\-₹€$£]?\s*[\d,\.]+\s*(%|Cr|MWh|GWh|MW|ms|Ω|°C|kWh)?$", clean_s, re.I))
                )

                if (
                    single_line
                    and len(clean_s) <= 115
                    and not looks_like_bullet
                    and not looks_like_kpi_val
                    and not clean_s.endswith(",")
                    and not (clean_s.endswith(".") and not is_numbered_h)
                    and (ratio >= 1.14 or (all_bold and (is_numbered_h or ratio >= 1.02)))
                ):
                    if ratio >= 1.62:
                        lvl = 1
                    elif ratio >= 1.32:
                        lvl = 2
                    elif ratio >= 1.14:
                        lvl = 3
                    else:
                        lvl = 4
                    md_parts.append(f"{'#' * lvl} {clean_s}\n")
                else:
                    # Format bullet/list items cleanly in Markdown
                    out_lines = []
                    for ln in txt.splitlines():
                        ls = ln.strip()
                        if ls.startswith("• "):
                            out_lines.append("- " + ls[2:].strip())
                        elif ls.startswith("◦ "):
                            out_lines.append("  - " + ls[2:].strip())
                        else:
                            out_lines.append(ls)
                    md_parts.append("\n".join(out_lines) + "\n")

            flush_mono()
            while t_idx < len(table_items):
                md_parts.append(_format_md_table(table_items[t_idx][1]) + "\n")
                t_idx += 1

        md_text = "\n".join(md_parts).strip() + "\n"
        return md_text.encode("utf-8"), "text/markdown; charset=utf-8", f"{base_name}.md", 1
    finally:
        doc.close()
        _release_pdf_memory()


def _segments_to_md(segments: List[Dict[str, Any]], include_inline_styles: bool = True) -> str:
    """Renders extracted paragraph segments to clean Markdown with [anchor](url) links and inline styles."""
    merged: List[Dict[str, Any]] = []
    for seg in segments:
        txt = seg.get("text", "")
        if not txt:
            continue
        if (
            merged
            and seg.get("url") is not None
            and merged[-1].get("url") == seg.get("url")
        ):
            merged[-1]["text"] += txt
        elif (
            merged
            and seg.get("url") is None
            and merged[-1].get("url") is None
            and merged[-1].get("bold") == seg.get("bold")
            and merged[-1].get("italic") == seg.get("italic")
            and merged[-1].get("mono") == seg.get("mono")
        ):
            merged[-1]["text"] += txt
        else:
            merged.append(dict(seg))

    out: List[str] = []
    for seg in merged:
        raw = seg["text"]
        if seg.get("url"):
            lbl = raw.strip()
            lead = raw[: len(raw) - len(raw.lstrip())]
            trail = raw[len(raw.rstrip()):]
            if lbl:
                out.append(f"{lead}[{lbl}]({seg['url']}){trail}")
            else:
                out.append(raw)
            continue

        if not include_inline_styles:
            out.append(raw)
            continue

        core = raw.strip()
        if not core:
            out.append(raw)
            continue
        lead = raw[: len(raw) - len(raw.lstrip())]
        trail = raw[len(raw.rstrip()):]

        if seg.get("mono"):
            formatted = f"`{core}`"
        elif seg.get("bold") and seg.get("italic"):
            formatted = f"***{core}***"
        elif seg.get("bold"):
            formatted = f"**{core}**"
        elif seg.get("italic"):
            formatted = f"*{core}*"
        else:
            formatted = core
        out.append(f"{lead}{formatted}{trail}")

    return "".join(out)


def _docx_to_md(docx_bytes: bytes) -> str:
    """Converts a DOCX file to structured GFM Markdown preserving H1-H6, lists, code blocks, callouts, links, and tables."""
    doc = Document(io.BytesIO(docx_bytes))
    parts: List[str] = []
    list_counters: Dict[Tuple[str, int], int] = {}
    in_list_block = False

    def close_list_block():
        nonlocal in_list_block
        if in_list_block:
            parts.append("")
            in_list_block = False

    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            info = _extract_oxml_paragraph_data(child, doc)
            plain = info["plain_text"].strip()

            if info["page_break_before"] and parts:
                close_list_block()
                parts.append("---\n")

            if plain:
                hl = info["heading_level"]
                if hl > 0:
                    close_list_block()
                    h_md = _segments_to_md(info["segments"], include_inline_styles=False).strip()
                    parts.append(f"{'#' * min(hl, 6)} {h_md}\n")
                elif info["is_mono_block"]:
                    close_list_block()
                    code_txt = info["plain_text"].strip("\n")
                    parts.append(f"```\n{code_txt}\n```\n")
                elif info["is_callout"]:
                    close_list_block()
                    q_md = _segments_to_md(info["segments"], include_inline_styles=True).strip()
                    q_lines = "\n".join(f"> {ln}" if ln.strip() else ">" for ln in q_md.splitlines())
                    parts.append(f"{q_lines}\n")
                elif info["list_info"] is not None:
                    li = info["list_info"]
                    ilvl = li["ilvl"]
                    indent = "  " * ilvl
                    item_md = _segments_to_md(info["segments"], include_inline_styles=True).strip()
                    if li["ordered"]:
                        key = (li["num_id"], ilvl)
                        list_counters[key] = list_counters.get(key, 0) + 1
                        parts.append(f"{indent}{list_counters[key]}. {item_md}")
                    else:
                        parts.append(f"{indent}- {item_md}")
                    in_list_block = True
                else:
                    close_list_block()
                    p_md = _segments_to_md(info["segments"], include_inline_styles=True).strip()
                    parts.append(f"{p_md}\n")

            if info["inline_page_breaks"] > 0 or info["has_section_break"]:
                close_list_block()
                if parts and not parts[-1].endswith("---\n"):
                    parts.append("---\n")

        elif child.tag.endswith('tbl'):
            close_list_block()
            tbl = Table(child, doc)
            raw_rows, _, _, _ = _extract_docx_table_grid_and_spans(tbl, doc)
            if raw_rows:
                rows_grid = []
                for r in raw_rows:
                    row_strs = []
                    for c in r:
                        if not c or not c["is_origin"]:
                            row_strs.append("")
                        else:
                            cell_md = " ".join(
                                _segments_to_md(p["segments"], include_inline_styles=False).strip()
                                for p in c["paras"]
                                if p["plain_text"].strip()
                            ).strip()
                            row_strs.append(cell_md)
                    rows_grid.append(row_strs)
                parts.append(_format_md_table(rows_grid) + "\n")

    close_list_block()
    return "\n".join(parts).strip() + "\n"


def any_to_md(data_bytes: bytes, input_fmt: str, base_name: str) -> Tuple[bytes, str, str, int]:
    """Generates structured Markdown from any document."""
    input_fmt = input_fmt.lower().replace("jpeg", "jpg")

    if input_fmt in ("png", "jpg", "webp"):
        items = extract_ocr_items(data_bytes)
        items.sort(key=lambda it: it['y_center'])
        md_text = f"# OCR Extracted Content: {base_name}\n\n" + "\n\n".join(it['text'] for it in items)

    elif input_fmt == "docx":
        md_text = _docx_to_md(data_bytes)

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
                rows_grid = [["" if v is None else str(v) for v in r] for r in rows]
                parts.append(_format_md_table(rows_grid) + "\n")
        md_text = "\n".join(parts)

    elif input_fmt == "pdf":
        return pdf_to_md(data_bytes, base_name)

    elif input_fmt == "pptx":
        docx_data, _, _, _ = pptx_to_docx(data_bytes, base_name)
        return any_to_md(docx_data, "docx", base_name)

    else:
        md_text = data_bytes.decode("utf-8", errors="replace")

    return md_text.encode("utf-8"), "text/markdown; charset=utf-8", f"{base_name}.md", 1


def _segments_to_html(segments: List[Dict[str, Any]], include_inline_styles: bool = True) -> str:
    """Renders paragraph segments to clean HTML5 inline markup with clickable links, colors, bold/italic/underline/code."""
    merged: List[Dict[str, Any]] = []
    for seg in segments:
        txt = seg.get("text", "")
        if not txt:
            continue
        if (
            merged
            and seg.get("url") is not None
            and merged[-1].get("url") == seg.get("url")
            and merged[-1].get("bold") == seg.get("bold")
            and merged[-1].get("italic") == seg.get("italic")
        ):
            merged[-1]["text"] += txt
        else:
            merged.append(dict(seg))

    out: List[str] = []
    for seg in merged:
        esc = html.escape(seg["text"]).replace("\n", "<br>")
        if seg.get("url"):
            esc_url = html.escape(seg["url"], quote=True)
            out.append(f'<a href="{esc_url}" target="_blank" rel="noopener">{esc}</a>')
            continue

        if not include_inline_styles:
            out.append(esc)
            continue

        if seg.get("mono"):
            esc = f"<code>{esc}</code>"
        if seg.get("bold"):
            esc = f"<strong>{esc}</strong>"
        if seg.get("italic"):
            esc = f"<em>{esc}</em>"
        if seg.get("underline"):
            esc = f"<u>{esc}</u>"
        col = seg.get("color")
        if col and col.upper() not in ("000000", "AUTO"):
            esc = f'<span style="color:#{html.escape(col)};">{esc}</span>'
        out.append(esc)
    return "".join(out)


def _docx_to_html_body(docx_bytes: bytes) -> str:
    """Converts a DOCX file into rich responsive HTML5 body content including embedded images, merges, lists, code, callouts, and links."""
    doc = Document(io.BytesIO(docx_bytes))
    parts: List[str] = []
    list_stack: List[Tuple[str, int]] = []

    def close_lists(target_level: int = -1, next_tag: Optional[str] = None):
        while list_stack:
            top_tag, top_lvl = list_stack[-1]
            if top_lvl > target_level or (top_lvl == target_level and next_tag is not None and top_tag != next_tag):
                parts.append(f"</{top_tag}>")
                list_stack.pop()
            else:
                break

    # Optional running header banner
    seen_hf: Set[str] = set()
    for sec in doc.sections:
        if sec.header:
            for p in sec.header.paragraphs:
                t = _extract_oxml_paragraph_data(p._p, doc)["plain_text"].strip()
                if t and t not in seen_hf:
                    seen_hf.add(t)
                    parts.append(f'<div class="doc-header">{html.escape(t)}</div>')

    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            info = _extract_oxml_paragraph_data(child, doc)
            if info["page_break_before"] and parts:
                close_lists()
                parts.append('<hr class="page-break">')

            # Embedded images inside paragraph
            for rid, cx_pt, cy_pt in info["images"]:
                rel_part = doc.part.related_parts.get(rid)
                if rel_part is None:
                    continue
                close_lists()
                try:
                    raw_blob = rel_part.blob
                    with Image.open(io.BytesIO(raw_blob)) as im:
                        fmt = (im.format or "").upper()
                        mode = im.mode
                        has_alpha = mode in ("RGBA", "LA") or (mode == "P" and "transparency" in im.info)
                    if fmt == "JPEG" and not has_alpha:
                        mime = "image/jpeg"
                        img_out = raw_blob
                    else:
                        img_out, _, _ = _flatten_image_for_reportlab(raw_blob)
                        mime = "image/png"
                    b64 = base64.b64encode(img_out).decode("ascii")
                    max_px = int(cx_pt * 1.333) if cx_pt else 820
                    align_cls = "center" if info["align"] == TA_CENTER else ("right" if info["align"] == TA_RIGHT else "left")
                    parts.append(
                        f'<figure class="doc-figure align-{align_cls}">'
                        f'<img src="data:{mime};base64,{b64}" alt="Embedded visual" style="max-width: min(100%, {max_px}px); height: auto;">'
                        f'</figure>'
                    )
                    del b64, img_out
                except Exception as img_err:
                    logger.warning(f"DOCX->HTML image embed failed for {rid}: {img_err}")

            plain = info["plain_text"].strip()
            if plain:
                hl = info["heading_level"]
                if hl > 0:
                    close_lists()
                    lvl = min(hl, 6)
                    h_html = _segments_to_html(info["segments"], include_inline_styles=False).strip()
                    parts.append(f"<h{lvl}>{h_html}</h{lvl}>")
                elif info["is_mono_block"]:
                    close_lists()
                    code_esc = html.escape(info["plain_text"].strip("\n"))
                    parts.append(f"<pre><code>{code_esc}</code></pre>")
                elif info["is_callout"]:
                    close_lists()
                    bg = info["callout_bg"] or "EFF6FF"
                    bd = info["callout_border"] or "2563EB"
                    c_html = _segments_to_html(info["segments"], include_inline_styles=True).strip()
                    parts.append(
                        f'<blockquote class="doc-callout" style="background:#{html.escape(bg)}; border-left:4px solid #{html.escape(bd)};">{c_html}</blockquote>'
                    )
                elif info["list_info"] is not None:
                    li = info["list_info"]
                    tag = "ol" if li["ordered"] else "ul"
                    ilvl = li["ilvl"]
                    close_lists(ilvl, tag)
                    if not list_stack or list_stack[-1][1] < ilvl:
                        parts.append(f"<{tag}>")
                        list_stack.append((tag, ilvl))
                    item_html = _segments_to_html(info["segments"], include_inline_styles=True).strip()
                    parts.append(f"<li>{item_html}</li>")
                else:
                    close_lists()
                    align_style = ""
                    if info["align"] == TA_CENTER:
                        align_style = ' style="text-align:center;"'
                    elif info["align"] == TA_RIGHT:
                        align_style = ' style="text-align:right;"'
                    p_html = _segments_to_html(info["segments"], include_inline_styles=True).strip()
                    parts.append(f"<p{align_style}>{p_html}</p>")

            if info["inline_page_breaks"] > 0 or info["has_section_break"]:
                close_lists()
                parts.append('<hr class="page-break">')

        elif child.tag.endswith('tbl'):
            close_lists()
            tbl = Table(child, doc)
            raw_rows, span_cmds, repeat_hdr_rows, _ = _extract_docx_table_grid_and_spans(tbl, doc)
            if not raw_rows:
                continue

            span_map = {(r0, c0): (c1 - c0 + 1, r1 - r0 + 1) for c0, r0, c1, r1 in span_cmds}
            tbl_html = ['<div class="table-wrap"><table>']
            for r_idx, row in enumerate(raw_rows):
                is_hdr = r_idx < max(repeat_hdr_rows, 1)
                tbl_html.append("<tr>")
                for c_idx, cell in enumerate(row):
                    if not cell or not cell["is_origin"]:
                        continue
                    cs, rs = span_map.get((r_idx, c_idx), (1, 1))
                    attrs = []
                    if cs > 1:
                        attrs.append(f'colspan="{cs}"')
                    if rs > 1:
                        attrs.append(f'rowspan="{rs}"')

                    fill_hex = cell.get("fill") or ("0F172A" if is_hdr else ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF"))
                    dark_bg = _is_dark_hex(fill_hex)
                    fg_hex = "FFFFFF" if dark_bg else "0F172A"
                    attrs.append(f'style="background-color:#{html.escape(fill_hex)}; color:#{fg_hex};"')

                    inner_parts = []
                    for p_info in cell["paras"]:
                        if p_info["plain_text"].strip():
                            segs = []
                            for s in p_info["segments"]:
                                sc = dict(s)
                                if dark_bg and sc.get("color") in (None, "0F172A", "1E293B", "334155", "000000"):
                                    sc["color"] = "FFFFFF"
                                segs.append(sc)
                            inner_parts.append(_segments_to_html(segs, include_inline_styles=True).strip())
                    cell_content = "<br>".join(inner_parts)
                    tag = "th" if is_hdr else "td"
                    attr_str = (" " + " ".join(attrs)) if attrs else ""
                    tbl_html.append(f"<{tag}{attr_str}>{cell_content}</{tag}>")
                tbl_html.append("</tr>")
            tbl_html.append("</table></div>")
            parts.append("".join(tbl_html))

    close_lists()
    for sec in doc.sections:
        if sec.footer:
            for p in sec.footer.paragraphs:
                t = _extract_oxml_paragraph_data(p._p, doc)["plain_text"].strip()
                t = re.sub(r"Page\s*__VELMO_PAGE__\s*of\s*__VELMO_NUMPAGES__", "", t, flags=re.I)
                t = re.sub(r"__VELMO_(?:PAGE|NUMPAGES)__", "", t).strip(" |\t")
                if t and t not in seen_hf:
                    seen_hf.add(t)
                    parts.append(f'<div class="doc-footer">{html.escape(t)}</div>')

    gc.collect()
    return "\n".join(parts)


def any_to_html(data_bytes: bytes, input_fmt: str, base_name: str) -> Tuple[bytes, str, str, int]:
    """Generates a clean responsive HTML5 document from any file format."""
    norm_fmt = input_fmt.lower().replace("jpeg", "jpg")
    if norm_fmt == "docx":
        body_html = _docx_to_html_body(data_bytes)
    else:
        md_bytes, _, _, _ = any_to_md(data_bytes, input_fmt, base_name)
        body_html = markdown.markdown(md_bytes.decode('utf-8'), extensions=['tables', 'fenced_code'])

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>{html.escape(base_name)}</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "DejaVu Sans", sans-serif; line-height: 1.6; max-width: 960px; margin: 40px auto; padding: 0 24px; color: #1e293b; background: #ffffff; }}
    h1, h2, h3, h4 {{ color: #0f172a; margin-top: 1.4em; }}
    h1, h2 {{ border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; }}
    .doc-header, .doc-footer {{ font-size: 0.82rem; color: #64748b; padding: 8px 0; border-bottom: 1px solid #e2e8f0; margin-bottom: 18px; }}
    .doc-footer {{ border-bottom: none; border-top: 1px solid #e2e8f0; margin-top: 28px; }}
    .doc-figure {{ margin: 18px 0; }}
    .doc-figure.align-center {{ text-align: center; }}
    .doc-figure.align-right {{ text-align: right; }}
    .doc-figure img {{ border-radius: 6px; background: #ffffff; }}
    .doc-callout {{ margin: 14px 0; padding: 12px 16px; border-radius: 6px; color: #1e293b; }}
    .page-break {{ border: none; border-top: 1px dashed #cbd5e1; margin: 28px 0; }}
    .table-wrap {{ width: 100%; overflow-x: auto; margin: 18px 0; }}
    table {{ border-collapse: collapse; width: 100%; font-size: 0.92rem; }}
    th, td {{ border: 1px solid #cbd5e1; padding: 7px 10px; text-align: left; vertical-align: middle; }}
    th {{ font-weight: 600; }}
    code {{ background: #f1f5f9; padding: 2px 6px; border-radius: 4px; font-family: "Consolas", "DejaVu Sans Mono", monospace; }}
    pre {{ background: #f1f5f9; padding: 14px 16px; border-radius: 6px; overflow-x: auto; border: 1px solid #e2e8f0; }}
    pre code {{ background: none; padding: 0; }}
    a {{ color: #2563eb; text-decoration: underline; }}
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
            finally:
                gc.collect()

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
            del pix
            b64_str = base64.b64encode(img_bytes).decode("ascii")
            pages_b64.append(f"data:image/jpeg;base64,{b64_str}")

        return {
            "total_pages": total_pages,
            "pages": pages_b64,
        }
    finally:
        doc.close()
        _release_pdf_memory()

