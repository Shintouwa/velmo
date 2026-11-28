import base64
import csv
import datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
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
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.datetime import from_excel as xl_from_excel

from docx import Document
from docx.enum.section import WD_ORIENTATION, WD_SECTION
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as DOCX_RT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Inches as DocxInches, Pt as DocxPt, RGBColor as DocxRGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN

import markdown
from bs4 import BeautifulSoup
from pdf2docx import Converter

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.utils import ImageReader
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


# --- Shared XLSX Parsing, Formula Evaluation & Formatting Helpers ---

_XLSX_THEME_COLORS = [
    "FFFFFF", "000000", "EEECE1", "1F497D", "4F81BD",
    "C0504D", "9BBB59", "8064A2", "4BACC6", "F79646",
]

_XLSX_INDEXED_COLORS = {
    0: "000000", 1: "FFFFFF", 2: "FF0000", 3: "00FF00", 4: "0000FF", 5: "FFFF00", 6: "FF00FF", 7: "00FFFF",
    8: "000000", 9: "FFFFFF", 10: "FF0000", 11: "00FF00", 12: "0000FF", 13: "FFFF00", 14: "FF00FF", 15: "00FFFF",
    16: "800000", 17: "008000", 18: "000080", 19: "808000", 20: "800080", 21: "008080", 22: "C0C0C0", 23: "808080",
    24: "9999FF", 25: "993366", 26: "FFFFCC", 27: "CCFFFF", 28: "660066", 29: "FF8080", 30: "0066CC", 31: "CCCCFF",
    40: "00CCFF", 41: "CCFFFF", 42: "CCFFCC", 43: "FFFF99", 44: "99CCFF", 45: "FF99CC", 46: "CC99FF", 47: "FFCC99",
    48: "3366FF", 49: "33CCCC", 50: "99CC00", 51: "FFCC00", 52: "FF9900", 53: "FF6600", 54: "666699", 55: "969696",
    56: "003366", 57: "339966", 58: "003300", 59: "333300", 60: "993300", 61: "993366", 62: "333399", 63: "333333",
}


def _apply_excel_tint(hex_rgb: str, tint: float) -> str:
    """Applies an Excel theme tint (-1.0 to 1.0) to a 6-char RGB hex string."""
    if not tint:
        return hex_rgb
    r = int(hex_rgb[0:2], 16)
    g = int(hex_rgb[2:4], 16)
    b = int(hex_rgb[4:6], 16)
    if tint < 0:
        factor = 1.0 + tint
        r, g, b = int(round(r * factor)), int(round(g * factor)), int(round(b * factor))
    else:
        r = int(round(r + (255 - r) * tint))
        g = int(round(g + (255 - g) * tint))
        b = int(round(b + (255 - b) * tint))
    r = max(0, min(255, r))
    g = max(0, min(255, g))
    b = max(0, min(255, b))
    return f"{r:02X}{g:02X}{b:02X}"


def _extract_xlsx_color_hex(color_obj, is_fill: bool = False) -> Optional[str]:
    """Extracts a 6-char hex color from an openpyxl Color (rgb, theme+tint, or indexed)."""
    if color_obj is None:
        return None
    ctype = getattr(color_obj, "type", None)
    rgb = getattr(color_obj, "rgb", None)
    if ctype == "rgb" and isinstance(rgb, str):
        if len(rgb) == 8 and rgb.upper().startswith("00"):
            return None
        hex_code = rgb[-6:].upper() if len(rgb) >= 6 else ""
        if re.match(r"^[0-9A-F]{6}$", hex_code):
            return hex_code
    if ctype == "theme":
        t_idx = getattr(color_obj, "theme", None)
        if isinstance(t_idx, int) and 0 <= t_idx < len(_XLSX_THEME_COLORS):
            tint = float(getattr(color_obj, "tint", 0.0) or 0.0)
            return _apply_excel_tint(_XLSX_THEME_COLORS[t_idx], tint)
    if ctype == "indexed":
        idx = getattr(color_obj, "indexed", None)
        if isinstance(idx, int) and idx in _XLSX_INDEXED_COLORS:
            return _XLSX_INDEXED_COLORS[idx]
    if isinstance(rgb, str) and len(rgb) >= 6:
        if len(rgb) == 8 and rgb.upper().startswith("00"):
            return None
        hex_code = rgb[-6:].upper()
        if re.match(r"^[0-9A-F]{6}$", hex_code):
            if is_fill and hex_code == "000000" and getattr(color_obj, "theme", None) is None:
                return None
            return hex_code
    return None


class _FormulaFallback(Exception):
    """Raised when an Excel formula uses unsupported syntax so the caller keeps the raw formula string."""


def _split_formula_top_level(expr: str, sep: str = ",") -> List[str]:
    """Splits an Excel formula expression on top-level separator outside parentheses and quotes."""
    parts: List[str] = []
    buf: List[str] = []
    depth = 0
    in_dq = False
    in_sq = False
    i = 0
    n = len(expr)
    while i < n:
        ch = expr[i]
        if ch == '"' and not in_sq:
            if in_dq and i + 1 < n and expr[i + 1] == '"':
                buf.append('""')
                i += 2
                continue
            in_dq = not in_dq
            buf.append(ch)
        elif ch == "'" and not in_dq:
            in_sq = not in_sq
            buf.append(ch)
        elif not in_dq and not in_sq:
            if ch == "(":
                depth += 1
            elif ch == ")" and depth > 0:
                depth -= 1
            elif depth == 0 and expr.startswith(sep, i):
                parts.append("".join(buf).strip())
                buf = []
                i += len(sep)
                continue
            buf.append(ch)
        else:
            buf.append(ch)
        i += 1
    parts.append("".join(buf).strip())
    return parts


def _parse_cell_coord(coord_str: str) -> Tuple[int, int]:
    """Parses '$B$5' or 'B5' into (row_1idx, col_1idx)."""
    clean = coord_str.replace("$", "").strip().upper()
    m = re.match(r"^([A-Z]{1,3})(\d+)$", clean)
    if not m:
        raise _FormulaFallback(coord_str)
    return int(m.group(2)), column_index_from_string(m.group(1))


def _resolve_xlsx_cell_value(
    sheet_title: str,
    row: int,
    col: int,
    wb_data,
    wb_raw,
    memo: Dict[Tuple[str, int, int], Any],
    visiting: Set[Tuple[str, int, int]],
) -> Any:
    """Returns cached value when available, or evaluates uncached standard formulas safely."""
    key = (sheet_title, row, col)
    if key in memo:
        return memo[key]
    if sheet_title not in wb_raw.sheetnames:
        raise _FormulaFallback(sheet_title)

    c_data = wb_data[sheet_title].cell(row=row, column=col)
    if c_data.value is not None:
        memo[key] = c_data.value
        return c_data.value

    c_raw = wb_raw[sheet_title].cell(row=row, column=col)
    raw_val = c_raw.value
    if not (isinstance(raw_val, str) and raw_val.startswith("=")):
        memo[key] = raw_val
        return raw_val

    if key in visiting:
        return raw_val
    visiting.add(key)
    try:
        res = _eval_formula_expr(raw_val[1:].strip(), sheet_title, wb_data, wb_raw, memo, visiting)
        memo[key] = res
        return res
    except _FormulaFallback:
        memo[key] = raw_val
        return raw_val
    except Exception:
        memo[key] = raw_val
        return raw_val
    finally:
        visiting.discard(key)


def _expand_range_ref(
    token: str,
    cur_sheet: str,
    wb_data,
    wb_raw,
    memo: Dict[Tuple[str, int, int], Any],
    visiting: Set[Tuple[str, int, int]],
) -> List[Any]:
    """Expands 'A1:B5' or 'Sheet!A1:B5' into a flat list of evaluated cell values."""
    t = token.strip()
    target_sheet = cur_sheet
    if "!" in t:
        s_part, ref_part = t.split("!", 1)
        s_clean = s_part.strip()
        if len(s_clean) >= 2 and s_clean[0] == "'" and s_clean[-1] == "'":
            s_clean = s_clean[1:-1].replace("''", "'")
        target_sheet = s_clean
        t = ref_part.strip()

    if ":" not in t:
        r, c = _parse_cell_coord(t)
        return [_resolve_xlsx_cell_value(target_sheet, r, c, wb_data, wb_raw, memo, visiting)]

    left_ref, right_ref = t.split(":", 1)
    r1, c1 = _parse_cell_coord(left_ref)
    r2, c2 = _parse_cell_coord(right_ref)
    if r1 > r2:
        r1, r2 = r2, r1
    if c1 > c2:
        c1, c2 = c2, c1
    if (r2 - r1 + 1) * (c2 - c1 + 1) > 20000:
        raise _FormulaFallback(token)
    vals: List[Any] = []
    for rr in range(r1, r2 + 1):
        for cc in range(c1, c2 + 1):
            vals.append(_resolve_xlsx_cell_value(target_sheet, rr, cc, wb_data, wb_raw, memo, visiting))
    return vals


def _to_number(v: Any) -> float:
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if v is None or v == "":
        return 0.0
    if isinstance(v, str):
        try:
            return float(v.replace(",", "").strip())
        except ValueError:
            pass
    raise _FormulaFallback(str(v))


def _eval_formula_expr(
    expr: str,
    cur_sheet: str,
    wb_data,
    wb_raw,
    memo: Dict[Tuple[str, int, int], Any],
    visiting: Set[Tuple[str, int, int]],
) -> Any:
    """Lightweight pure-Python evaluator for common Excel formulas (SUM, AVERAGE, IF, cross-sheet refs, arithmetic)."""
    s = expr.strip()
    if not s:
        raise _FormulaFallback("empty")

    # Strip matching outer parentheses
    while s.startswith("(") and s.endswith(")"):
        depth = 0
        in_dq = False
        in_sq = False
        matched = True
        for idx, ch in enumerate(s):
            if ch == '"' and not in_sq:
                in_dq = not in_dq
            elif ch == "'" and not in_dq:
                in_sq = not in_sq
            elif not in_dq and not in_sq:
                if ch == "(":
                    depth += 1
                elif ch == ")":
                    depth -= 1
                    if depth == 0 and idx < len(s) - 1:
                        matched = False
                        break
        if matched:
            s = s[1:-1].strip()
        else:
            break

    # 1. Comparisons (lowest precedence)
    for op in (">=", "<=", "<>", "!=", "=", ">", "<"):
        parts = _split_formula_top_level(s, op)
        if len(parts) == 2 and parts[0] and parts[1]:
            lv = _eval_formula_expr(parts[0], cur_sheet, wb_data, wb_raw, memo, visiting)
            rv = _eval_formula_expr(parts[1], cur_sheet, wb_data, wb_raw, memo, visiting)
            if isinstance(lv, str) and isinstance(rv, str):
                lcmp, rcmp = lv.lower(), rv.lower()
            else:
                lcmp, rcmp = _to_number(lv), _to_number(rv)
            if op == "=":
                return lcmp == rcmp
            if op in ("<>", "!="):
                return lcmp != rcmp
            if op == ">":
                return lcmp > rcmp
            if op == "<":
                return lcmp < rcmp
            if op == ">=":
                return lcmp >= rcmp
            if op == "<=":
                return lcmp <= rcmp

    # 2. String concatenation '&'
    amp_parts = _split_formula_top_level(s, "&")
    if len(amp_parts) > 1:
        out_parts: List[str] = []
        for p in amp_parts:
            v = _eval_formula_expr(p, cur_sheet, wb_data, wb_raw, memo, visiting)
            out_parts.append("" if v is None else str(v))
        return "".join(out_parts)

    # 3. Addition / Subtraction outside parens/quotes/scientific notation
    depth = 0
    in_dq = False
    in_sq = False
    for idx in range(len(s) - 1, 0, -1):
        ch = s[idx]
        if ch == '"' and not in_sq:
            in_dq = not in_dq
        elif ch == "'" and not in_dq:
            in_sq = not in_sq
        elif not in_dq and not in_sq:
            if ch == ")":
                depth += 1
            elif ch == "(":
                depth -= 1
            elif depth == 0 and ch in ("+", "-"):
                prev = s[:idx].rstrip()
                if prev and prev[-1] not in ("+", "-", "*", "/", "(", ",", "=", ">", "<", "&") and not prev.lower().endswith("e"):
                    lv = _to_number(_eval_formula_expr(prev, cur_sheet, wb_data, wb_raw, memo, visiting))
                    rv = _to_number(_eval_formula_expr(s[idx + 1:], cur_sheet, wb_data, wb_raw, memo, visiting))
                    return (lv + rv) if ch == "+" else (lv - rv)

    # 4. Multiplication / Division
    depth = 0
    in_dq = False
    in_sq = False
    for idx in range(len(s) - 1, 0, -1):
        ch = s[idx]
        if ch == '"' and not in_sq:
            in_dq = not in_dq
        elif ch == "'" and not in_dq:
            in_sq = not in_sq
        elif not in_dq and not in_sq:
            if ch == ")":
                depth += 1
            elif ch == "(":
                depth -= 1
            elif depth == 0 and ch in ("*", "/"):
                lv = _to_number(_eval_formula_expr(s[:idx], cur_sheet, wb_data, wb_raw, memo, visiting))
                rv = _to_number(_eval_formula_expr(s[idx + 1:], cur_sheet, wb_data, wb_raw, memo, visiting))
                if ch == "*":
                    return lv * rv
                if rv == 0:
                    raise _FormulaFallback("div0")
                return lv / rv

    # Unary +/-
    if s[0] == "-":
        return -_to_number(_eval_formula_expr(s[1:], cur_sheet, wb_data, wb_raw, memo, visiting))
    if s[0] == "+":
        return _to_number(_eval_formula_expr(s[1:], cur_sheet, wb_data, wb_raw, memo, visiting))

    # 5. Literals
    if len(s) >= 2 and s[0] == '"' and s[-1] == '"':
        return s[1:-1].replace('""', '"')
    if s.upper() == "TRUE":
        return True
    if s.upper() == "FALSE":
        return False
    if re.match(r"^\d+(\.\d+)?$", s):
        return int(s) if "." not in s else float(s)

    # 6. Function Call
    m_fn = re.match(r"^([A-Z0-9_.]+)\((.*)\)$", s, re.DOTALL | re.I)
    if m_fn:
        fn_name = m_fn.group(1).upper()
        arg_str = m_fn.group(2).strip()
        args = _split_formula_top_level(arg_str, ",") if arg_str else []

        if fn_name == "IF":
            if len(args) < 2:
                raise _FormulaFallback(s)
            cond = _eval_formula_expr(args[0], cur_sheet, wb_data, wb_raw, memo, visiting)
            truthy = bool(cond) if not isinstance(cond, str) else (cond.strip().upper() not in ("", "FALSE", "0"))
            if truthy:
                return _eval_formula_expr(args[1], cur_sheet, wb_data, wb_raw, memo, visiting)
            return _eval_formula_expr(args[2], cur_sheet, wb_data, wb_raw, memo, visiting) if len(args) >= 3 else False

        if fn_name == "TODAY":
            return datetime.date.today()

        def _collect_args(numeric_only: bool = True) -> List[Any]:
            out_vals: List[Any] = []
            for a in args:
                a_st = a.strip()
                if re.match(r"^(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?\d+$", a_st):
                    expanded = _expand_range_ref(a_st, cur_sheet, wb_data, wb_raw, memo, visiting)
                else:
                    expanded = [_eval_formula_expr(a_st, cur_sheet, wb_data, wb_raw, memo, visiting)]
                for item in expanded:
                    if numeric_only:
                        if isinstance(item, (int, float)) and not isinstance(item, bool):
                            out_vals.append(float(item))
                    else:
                        if item is not None and str(item) != "":
                            out_vals.append(item)
            return out_vals

        if fn_name == "SUM":
            return sum(_collect_args(True))
        if fn_name in ("AVERAGE", "AVG"):
            nums = _collect_args(True)
            if not nums:
                raise _FormulaFallback(s)
            return sum(nums) / len(nums)
        if fn_name == "MIN":
            nums = _collect_args(True)
            if not nums:
                raise _FormulaFallback(s)
            return min(nums)
        if fn_name == "MAX":
            nums = _collect_args(True)
            if not nums:
                raise _FormulaFallback(s)
            return max(nums)
        if fn_name == "COUNT":
            return len(_collect_args(True))
        if fn_name == "COUNTA":
            return len(_collect_args(False))
        if fn_name == "ROUND" and len(args) == 2:
            v = _to_number(_eval_formula_expr(args[0], cur_sheet, wb_data, wb_raw, memo, visiting))
            nd = int(_to_number(_eval_formula_expr(args[1], cur_sheet, wb_data, wb_raw, memo, visiting)))
            return round(v, nd)
        if fn_name == "ABS" and len(args) == 1:
            return abs(_to_number(_eval_formula_expr(args[0], cur_sheet, wb_data, wb_raw, memo, visiting)))
        if fn_name in ("CONCAT", "CONCATENATE"):
            return "".join("" if x is None else str(x) for x in _collect_args(False))
        if fn_name == "LEN" and len(args) == 1:
            return len(str(_eval_formula_expr(args[0], cur_sheet, wb_data, wb_raw, memo, visiting) or ""))
        if fn_name == "TRIM" and len(args) == 1:
            return str(_eval_formula_expr(args[0], cur_sheet, wb_data, wb_raw, memo, visiting) or "").strip()
        if fn_name == "LOWER" and len(args) == 1:
            return str(_eval_formula_expr(args[0], cur_sheet, wb_data, wb_raw, memo, visiting) or "").lower()
        if fn_name == "UPPER" and len(args) == 1:
            return str(_eval_formula_expr(args[0], cur_sheet, wb_data, wb_raw, memo, visiting) or "").upper()
        if fn_name == "AND":
            return all(bool(_eval_formula_expr(a, cur_sheet, wb_data, wb_raw, memo, visiting)) for a in args)
        if fn_name == "OR":
            return any(bool(_eval_formula_expr(a, cur_sheet, wb_data, wb_raw, memo, visiting)) for a in args)
        if fn_name == "NOT" and len(args) == 1:
            return not bool(_eval_formula_expr(args[0], cur_sheet, wb_data, wb_raw, memo, visiting))
        raise _FormulaFallback(fn_name)

    # 7. Single Cell Reference (same-sheet or cross-sheet)
    if re.match(r"^(?:(?:'[^']+'|[A-Za-z0-9_]+)!)?\$?[A-Z]{1,3}\$?\d+$", s):
        vals = _expand_range_ref(s, cur_sheet, wb_data, wb_raw, memo, visiting)
        val = vals[0] if vals else None
        if isinstance(val, str) and val.startswith("="):
            raise _FormulaFallback(val)
        return 0 if val is None else val

    raise _FormulaFallback(s)


def _split_numfmt_sections(fmt: str) -> List[str]:
    """Splits an Excel number_format string on unquoted semicolons."""
    secs: List[str] = []
    buf: List[str] = []
    in_q = False
    i = 0
    while i < len(fmt):
        ch = fmt[i]
        if ch == "\\" and i + 1 < len(fmt):
            buf.append(fmt[i:i + 2])
            i += 2
            continue
        if ch == '"':
            in_q = not in_q
            buf.append(ch)
        elif ch == ";" and not in_q:
            secs.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    secs.append("".join(buf))
    return secs


def _format_excel_value(val: Any, number_format: Optional[str]) -> str:
    """Formats Python/Excel cell values cleanly according to Excel number formats and dates."""
    if val is None:
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"

    fmt = (number_format or "General").strip()

    if isinstance(val, datetime.datetime):
        has_time_fmt = bool(re.search(r"(?:h|s|am/pm|:mm|mm:)", fmt, re.I))
        if has_time_fmt and (val.hour or val.minute or val.second):
            return val.strftime("%Y-%m-%d %H:%M")
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.time):
        return val.strftime("%H:%M:%S")
    if isinstance(val, str):
        return val

    if not isinstance(val, (int, float)):
        return str(val)

    # Excel serial date check (date format tokens without numeric '#'/'0' placeholders)
    fmt_clean_tags = re.sub(r'"[^"]*"|\[[^\]]*\]', "", fmt)
    if re.search(r"[yd]", fmt_clean_tags, re.I) and not re.search(r"[0#]", fmt_clean_tags):
        try:
            dt_val = xl_from_excel(float(val))
            if isinstance(dt_val, (datetime.datetime, datetime.date)):
                return dt_val.strftime("%Y-%m-%d")
        except Exception:
            pass

    if fmt in ("", "General", "@"):
        if isinstance(val, int):
            return str(val)
        if float(val).is_integer() and abs(val) < 1e12:
            return str(int(val))
        return str(val)

    sections = _split_numfmt_sections(fmt)
    is_neg = val < 0
    if is_neg and len(sections) > 1:
        sec = sections[1]
        explicit_neg_section = True
    elif val == 0 and len(sections) > 2:
        sec = sections[2]
        explicit_neg_section = False
    else:
        sec = sections[0]
        explicit_neg_section = False

    # Extract currency/literal tokens before stripping format tags
    sec_work = sec
    # Replace [$₹-4009] or [$€-2] with literal symbol
    sec_work = re.sub(r"\[\$([^-\]]+)(?:-[^\]]*)?\]", r"\1", sec_work)
    # Strip remaining bracket tags like [Red], [h], etc. and alignment skips _X / *X
    sec_work = re.sub(r"\[[^\]]*\]", "", sec_work)
    sec_work = re.sub(r"_.|\*.", "", sec_work)
    sec_work = sec_work.replace("\\", "").replace('"', "")

    has_parens = "(" in sec_work and ")" in sec_work
    is_pct = "%" in sec_work

    num_match = re.search(r"[#0,.]+", sec_work)
    if not num_match:
        return str(val)

    prefix_raw = sec_work[:num_match.start()]
    num_pat = num_match.group(0)
    suffix_raw = sec_work[num_match.end():]

    # Clean leading '-' or '()' from prefix/suffix so we apply sign consistently
    prefix = prefix_raw.replace("(", "").replace(")", "").replace("-", "").strip()
    suffix = suffix_raw.replace("(", "").replace(")", "").replace("-", "").strip()

    mag = abs(float(val)) * (100.0 if is_pct else 1.0)
    dec_places = 0
    if "." in num_pat:
        dec_part = num_pat.split(".", 1)[1]
        dec_places = sum(1 for ch in dec_part if ch in ("0", "#"))

    try:
        q_exp = Decimal("1") if dec_places == 0 else Decimal("1." + "0" * dec_places)
        d_mag = Decimal(str(mag)).quantize(q_exp, rounding=ROUND_HALF_UP)
        num_str = f"{d_mag:,f}" if "," in num_pat else f"{d_mag:f}"
    except (InvalidOperation, ValueError):
        num_str = f"{mag:,.{dec_places}f}" if "," in num_pat else f"{mag:.{dec_places}f}"

    core = f"{prefix}{num_str}{suffix}"
    if is_neg:
        if has_parens or (explicit_neg_section and "(" in sec_work):
            return f"({core})"
        return f"-{core}"
    return core


def _extract_xlsx_workbook_model(xlsx_bytes: bytes) -> List[Dict[str, Any]]:
    """
    Extracts every sheet of an XLSX workbook into a normalized, trimmed, styled model:
    - Evaluates uncached formulas (with cycle/fallback protection) and formats numbers/dates cleanly
    - Trims leading empty rows/columns (e.g. C4 offset tables) while marking completely blank sheets
    - Maps Excel merged ranges into trimmed grid (col0, row0, col1, row1) spans
    - Extracts background fills, high-contrast font colors, alignment, wrap, hyperlinks & flattened images
    """
    wb_data = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
    wb_raw = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=False)
    eval_memo: Dict[Tuple[str, int, int], Any] = {}
    visiting: Set[Tuple[str, int, int]] = set()
    sheets_out: List[Dict[str, Any]] = []

    try:
        for ws_raw in wb_raw.worksheets:
            title = ws_raw.title
            ws_data = wb_data[title]

            # Extract and alpha-flatten anchored images
            images_info: List[Dict[str, Any]] = []
            for img_obj in getattr(ws_raw, "_images", []):
                try:
                    raw_img_bytes = img_obj._data()
                    flat_bytes, iw, ih = _flatten_image_for_reportlab(raw_img_bytes)
                    disp_w = float(getattr(img_obj, "width", 0) or iw)
                    disp_h = float(getattr(img_obj, "height", 0) or ih)
                    if disp_w <= 0:
                        disp_w = float(iw)
                    if disp_h <= 0:
                        disp_h = float(ih)
                    anchor = getattr(img_obj, "anchor", None)
                    anc_row = getattr(getattr(anchor, "_from", None), "row", 0) if anchor else 0
                    images_info.append({
                        "bytes": flat_bytes,
                        "iw": iw,
                        "ih": ih,
                        "disp_w_pt": disp_w * 0.75,
                        "disp_h_pt": disp_h * 0.75,
                        "anchor_row": int(anc_row or 0),
                    })
                except Exception as img_err:
                    logger.warning(f"XLSX image extraction failed on '{title}': {img_err}")

            # Scan only populated cells (O(populated), safe against 1M-row blank formatting!)
            non_empty_coords: List[Tuple[int, int]] = []
            for (r_c, c_c), cell_raw in ws_raw._cells.items():
                c_data = ws_data._cells.get((r_c, c_c))
                has_val = (c_data is not None and c_data.value is not None and str(c_data.value) != "") or (
                    cell_raw.value is not None and str(cell_raw.value) != ""
                )
                has_hl = cell_raw.hyperlink is not None and bool(getattr(cell_raw.hyperlink, "target", None))
                if has_val or has_hl:
                    non_empty_coords.append((r_c, c_c))

            if not non_empty_coords:
                sheets_out.append({
                    "title": title,
                    "is_empty": True,
                    "rows": [],
                    "span_cmds": [],
                    "repeat_hdr_rows": 0,
                    "banner_rows": 0,
                    "col_width_hints": [],
                    "images": images_info,
                })
                continue

            min_r = min(r for r, _ in non_empty_coords)
            max_r = max(r for r, _ in non_empty_coords)
            min_c = min(c for _, c in non_empty_coords)
            max_c = max(c for _, c in non_empty_coords)

            # Expand bounding box if any merged range containing a populated top-left extends further
            coord_set = set(non_empty_coords)
            active_merges = []
            for mr in ws_raw.merged_cells.ranges:
                if (mr.min_row, mr.min_col) in coord_set or any(
                    min_r <= rr <= max_r and min_c <= cc <= max_c
                    for rr in range(mr.min_row, mr.max_row + 1)
                    for cc in range(mr.min_col, mr.max_col + 1)
                ):
                    min_r = min(min_r, mr.min_row)
                    max_r = max(max_r, mr.max_row)
                    min_c = min(min_c, mr.min_col)
                    max_c = max(max_c, mr.max_col)
                    active_merges.append(mr)

            n_rows = max_r - min_r + 1
            n_cols = max_c - min_c + 1

            # Map merged ranges into trimmed 0-indexed coordinates
            span_cmds: List[Tuple[int, int, int, int]] = []
            merged_child_map: Dict[Tuple[int, int], Tuple[int, int]] = {}
            for mr in active_merges:
                r0 = mr.min_row - min_r
                c0 = mr.min_col - min_c
                r1 = mr.max_row - min_r
                c1 = mr.max_col - min_c
                if c1 > c0 or r1 > r0:
                    span_cmds.append((c0, r0, c1, r1))
                for rr in range(r0, r1 + 1):
                    for cc in range(c0, c1 + 1):
                        if rr != r0 or cc != c0:
                            merged_child_map[(rr, cc)] = (r0, c0)

            col_hints: List[float] = []
            for abs_c in range(min_c, max_c + 1):
                cd = ws_raw.column_dimensions.get(get_column_letter(abs_c))
                col_hints.append(float(cd.width) if (cd and cd.width) else 12.0)

            grid: List[List[Dict[str, Any]]] = []
            for r_off in range(n_rows):
                abs_r = min_r + r_off
                row_list: List[Dict[str, Any]] = []
                for c_off in range(n_cols):
                    abs_c = min_c + c_off
                    c_raw = ws_raw.cell(row=abs_r, column=abs_c)
                    is_origin = (r_off, c_off) not in merged_child_map

                    if not is_origin:
                        orig_r, orig_c = merged_child_map[(r_off, c_off)]
                        fill_hex = _extract_xlsx_color_hex(
                            c_raw.fill.fgColor if c_raw.fill and c_raw.fill.fill_type == "solid" else None,
                            is_fill=True,
                        )
                        row_list.append({
                            "text": "",
                            "is_origin": False,
                            "origin": (orig_r, orig_c),
                            "fill": fill_hex,
                            "color": None,
                            "bold": False,
                            "italic": False,
                            "align": "left",
                            "wrap": False,
                            "url": None,
                        })
                        continue

                    val = _resolve_xlsx_cell_value(title, abs_r, abs_c, wb_data, wb_raw, eval_memo, visiting)
                    fmt = c_raw.number_format
                    text_str = _format_excel_value(val, fmt)

                    hl_url: Optional[str] = None
                    if c_raw.hyperlink and getattr(c_raw.hyperlink, "target", None):
                        hl_url = str(c_raw.hyperlink.target).strip()
                        if not text_str and getattr(c_raw.hyperlink, "display", None):
                            text_str = str(c_raw.hyperlink.display)
                        elif not text_str:
                            text_str = hl_url

                    fill_hex = None
                    if c_raw.fill and c_raw.fill.fill_type == "solid":
                        fill_hex = _extract_xlsx_color_hex(c_raw.fill.fgColor or c_raw.fill.start_color, is_fill=True)

                    font_obj = c_raw.font
                    is_bold = bool(font_obj and font_obj.bold)
                    is_italic = bool(font_obj and font_obj.italic)
                    font_hex = _extract_xlsx_color_hex(font_obj.color, is_fill=False) if font_obj else None

                    dark_bg = _is_dark_hex(fill_hex)
                    if dark_bg and font_hex in (None, "000000", "0F172A", "1E293B", "334155"):
                        font_hex = "FFFFFF"
                    elif not dark_bg and fill_hex is None and font_hex == "FFFFFF":
                        font_hex = "0F172A"

                    al = c_raw.alignment
                    h_align = (al.horizontal if al and al.horizontal else "") or ""
                    if h_align not in ("left", "center", "right"):
                        h_align = "right" if isinstance(val, (int, float)) and not isinstance(val, bool) else "left"
                    is_wrap = bool(al and al.wrap_text) or ("\n" in text_str)

                    row_list.append({
                        "text": text_str,
                        "is_origin": True,
                        "origin": (r_off, c_off),
                        "fill": fill_hex,
                        "color": font_hex,
                        "bold": is_bold,
                        "italic": is_italic,
                        "align": h_align,
                        "wrap": is_wrap,
                        "url": hl_url,
                    })
                grid.append(row_list)

            # Detect full-width banner rows at top & safe repeating header row count
            banner_rows = 0
            for r_chk in range(min(2, n_rows)):
                if any(c0 == 0 and r0 == r_chk and c1 >= max(n_cols - 2, 1) and r1 == r_chk for c0, r0, c1, r1 in span_cmds):
                    banner_rows += 1
                else:
                    break

            repeat_hdr_rows = 0
            if n_rows > 1:
                # Start after top banners if table has data below; include contiguous bold/shaded header rows
                start_hdr = banner_rows if banner_rows < n_rows - 1 else 0
                hdr_end = start_hdr
                while hdr_end < min(n_rows - 1, start_hdr + 3):
                    row_cells = grid[hdr_end]
                    orig_cells = [c for c in row_cells if c["is_origin"] and c["text"].strip()]
                    if orig_cells and (
                        all(c["bold"] or c["fill"] is not None for c in orig_cells) or hdr_end == start_hdr
                    ):
                        hdr_end += 1
                        if hdr_end == start_hdr + 1 and not any(
                            r0 < hdr_end <= r1 for _, r0, _, r1 in span_cmds
                        ):
                            break
                    else:
                        break
                repeat_hdr_rows = hdr_end
                # Never allow repeatRows to split across a vertical merge
                for _, r0, _, r1 in span_cmds:
                    if r0 < repeat_hdr_rows <= r1:
                        repeat_hdr_rows = r1 + 1
                if repeat_hdr_rows >= n_rows:
                    repeat_hdr_rows = 1 if not any(r0 == 0 and r1 > 0 for _, r0, _, r1 in span_cmds) else 0

            sheets_out.append({
                "title": title,
                "is_empty": False,
                "rows": grid,
                "span_cmds": span_cmds,
                "repeat_hdr_rows": repeat_hdr_rows,
                "banner_rows": banner_rows,
                "col_width_hints": col_hints,
                "images": images_info,
            })
    finally:
        wb_data.close()
        wb_raw.close()
        eval_memo.clear()
        gc.collect()

    return sheets_out


def _compute_xlsx_col_widths(
    sheet: Dict[str, Any],
    avail_width: float,
    sans_family: Optional[str] = None,
    bold_family: Optional[str] = None,
) -> Tuple[List[float], float]:
    """Computes exact-glyph-aware proportional column widths within avail_width so header/cell words don't mid-wrap."""
    reg_sans, _ = _ensure_reportlab_unicode_fonts()
    if sans_family is None:
        sans_family = reg_sans
    if bold_family is None:
        bold_family = f"{reg_sans}-Bold" if reg_sans == "VelmoSans" else "Helvetica-Bold"
    try:
        font_sans_obj = pdfmetrics.getFont(sans_family)
        font_bold_obj = pdfmetrics.getFont(bold_family)
    except Exception:
        font_sans_obj = font_bold_obj = None

    rows = sheet["rows"]
    num_cols = len(rows[0]) if rows else 0
    if num_cols == 0:
        return [], 8.5

    if num_cols >= 11:
        font_pt, pad_pt = 6.0, 6.0
    elif num_cols >= 9:
        font_pt, pad_pt = 6.8, 7.0
    elif num_cols >= 7:
        font_pt, pad_pt = 7.4, 8.0
    else:
        font_pt, pad_pt = 8.4, 9.0

    def _measure_min_pref(f_pt: float) -> Tuple[List[float], List[float]]:
        mn = [24.0] * num_cols
        pf = [34.0] * num_cols
        hints = sheet.get("col_width_hints", [])
        for r_idx, r in enumerate(rows):
            for c_i, cell in enumerate(r):
                if not cell["is_origin"]:
                    continue
                if any(c0 == c_i and r0 == r_idx and c1 > c0 for c0, r0, c1, _ in sheet["span_cmds"]):
                    continue
                txt = cell["text"].strip()
                if not txt:
                    continue
                f_obj = font_bold_obj if cell["bold"] else font_sans_obj
                words = [w for w in re.split(r"\s+", txt) if w]
                # Measure exact longest unbreakable token (up to 18 chars so pathological tokens don't crush neighbors)
                max_w_pt = 0.0
                for w in words:
                    sample_w = w[:18]
                    if f_obj is not None:
                        w_pt = f_obj.stringWidth(sample_w, f_pt)
                    else:
                        w_pt = len(sample_w) * f_pt * 0.64
                    if w_pt > max_w_pt:
                        max_w_pt = w_pt
                mn[c_i] = max(mn[c_i], max_w_pt + pad_pt)

                longest_ln = max(txt.splitlines(), key=len, default="")[:34]
                if f_obj is not None:
                    ln_pt = f_obj.stringWidth(longest_ln, f_pt)
                else:
                    ln_pt = len(longest_ln) * f_pt * 0.58
                pf[c_i] = max(pf[c_i], ln_pt + pad_pt + 2.0)

        if len(hints) == num_cols and sum(hints) > 0:
            h_scale = avail_width / sum(hints)
            for c_i in range(num_cols):
                pf[c_i] = max(mn[c_i], 0.70 * pf[c_i] + 0.30 * (hints[c_i] * h_scale))
        return mn, pf

    min_w, pref_w = _measure_min_pref(font_pt)
    while sum(min_w) > avail_width and font_pt > 5.4:
        font_pt = round(font_pt - 0.3, 1)
        min_w, pref_w = _measure_min_pref(font_pt)

    sum_min = sum(min_w)
    if sum_min >= avail_width:
        ratio = avail_width / max(sum_min, 1.0)
        widths = [w * ratio for w in min_w]
    else:
        extra = avail_width - sum_min
        deltas = [max(pref_w[i] - min_w[i], 0.5) for i in range(num_cols)]
        sum_delta = sum(deltas)
        widths = [min_w[i] + extra * (deltas[i] / sum_delta) for i in range(num_cols)]

    return widths, font_pt


# --- Word & Excel Core Converters ---

def _add_docx_hyperlink(paragraph, url: str, text: str, font_pt: float, hex_color: str = "2563EB", bold: bool = False):
    """Appends a real clickable Word <w:hyperlink> element to a python-docx paragraph."""
    r_id = paragraph.part.relate_to(url, DOCX_RT.HYPERLINK, is_external=True)
    sz_half_pt = int(round(font_pt * 2))
    b_tag = "<w:b/>" if bold else ""
    safe_txt = html.escape(text, quote=False)
    hl_xml = (
        f'<w:hyperlink {nsdecls("w", "r")} r:id="{r_id}">'
        f'<w:r><w:rPr>{b_tag}<w:color w:val="{hex_color}"/><w:u w:val="single"/>'
        f'<w:sz w:val="{sz_half_pt}"/><w:szCs w:val="{sz_half_pt}"/></w:rPr>'
        f'<w:t xml:space="preserve">{safe_txt}</w:t></w:r></w:hyperlink>'
    )
    paragraph._p.append(parse_xml(hl_xml))


def _set_docx_cell_width_and_margins(cell, width_pt: float, v_pad_tw: int = 45, h_pad_tw: int = 70):
    """Sets explicit Word cell width (dxa) and compact cell margins in O(1)."""
    tcPr = cell._tc.get_or_add_tcPr()
    for child in list(tcPr):
        if child.tag.endswith(("tcW", "tcMar")):
            tcPr.remove(child)
    twips = max(int(round(width_pt * 20.0)), 180)
    tcPr.append(parse_xml(f'<w:tcW {nsdecls("w")} w:w="{twips}" w:type="dxa"/>'))
    tcPr.append(parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{v_pad_tw}" w:type="dxa"/>'
        f'<w:bottom w:w="{v_pad_tw}" w:type="dxa"/>'
        f'<w:left w:w="{h_pad_tw}" w:type="dxa"/>'
        f'<w:right w:w="{h_pad_tw}" w:type="dxa"/>'
        f'</w:tcMar>'
    ))


def xlsx_to_docx(xlsx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    sheets = _extract_xlsx_workbook_model(xlsx_bytes)
    doc = Document()
    sheet_count = len(sheets)

    def _configure_section(sec, use_landscape: bool) -> float:
        if use_landscape:
            sec.orientation = WD_ORIENTATION.LANDSCAPE
            if sec.page_width < sec.page_height:
                sec.page_width, sec.page_height = sec.page_height, sec.page_width
            sec.left_margin = DocxInches(0.45)
            sec.right_margin = DocxInches(0.45)
        else:
            sec.orientation = WD_ORIENTATION.PORTRAIT
            if sec.page_width > sec.page_height:
                sec.page_width, sec.page_height = sec.page_height, sec.page_width
            sec.left_margin = DocxInches(0.55)
            sec.right_margin = DocxInches(0.55)
        sec.top_margin = DocxInches(0.50)
        sec.bottom_margin = DocxInches(0.50)
        return float(sec.page_width.pt - sec.left_margin.pt - sec.right_margin.pt)

    align_map = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
    }
    prev_landscape: Optional[bool] = None

    for s_idx, sh in enumerate(sheets):
        n_cols = len(sh["rows"][0]) if sh["rows"] else 1
        est_hint_pt = sum(sh.get("col_width_hints", [])) * 5.2
        use_landscape = n_cols >= 7 or est_hint_pt > 530.0

        if s_idx == 0:
            avail_w_pt = _configure_section(doc.sections[0], use_landscape)
        elif use_landscape != prev_landscape:
            new_sec = doc.add_section(WD_SECTION.NEW_PAGE)
            avail_w_pt = _configure_section(new_sec, use_landscape)
        else:
            doc.add_page_break()
            avail_w_pt = _configure_section(doc.sections[-1], use_landscape)
        prev_landscape = use_landscape

        h_para = doc.add_heading(sh["title"], level=1)
        h_para.paragraph_format.space_after = DocxPt(6)

        if sh["is_empty"] and not sh["images"]:
            p_emp = doc.add_paragraph("(Empty worksheet — no data)")
            if p_emp.runs:
                p_emp.runs[0].italic = True
            continue

        if not sh["is_empty"] and sh["rows"]:
            num_rows = len(sh["rows"])
            num_cols = len(sh["rows"][0])
            col_widths_pt, base_font_pt = _compute_xlsx_col_widths(sh, avail_w_pt)
            docx_font_pt = max(base_font_pt + 1.1, 7.5 if num_cols >= 11 else 8.5)

            table = doc.add_table(rows=num_rows, cols=num_cols)
            table.style = "Table Grid"
            table.autofit = False

            # Set tblGrid widths once
            ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            tblGrid = table._tbl.find(f"{ns_w}tblGrid")
            if tblGrid is not None:
                for g_col, w_pt in zip(tblGrid.findall(f"{ns_w}gridCol"), col_widths_pt):
                    g_col.set(f"{ns_w}w", str(max(int(round(w_pt * 20.0)), 180)))

            # Cache row.cells once per row so cell lookup is strictly O(1)
            row_cells_grid = [list(row.cells) for row in table.rows]
            v_pad_tw = 35 if num_rows > 40 else 50
            h_pad_tw = 45 if num_cols >= 10 else 75

            for r_i, row_data in enumerate(sh["rows"]):
                trPr = table.rows[r_i]._tr.get_or_add_trPr()
                trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
                if r_i < sh["repeat_hdr_rows"] and num_rows > sh["repeat_hdr_rows"]:
                    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))

                row_cells = row_cells_grid[r_i]
                for c_i in range(num_cols):
                    cell_info = row_data[c_i]
                    t_cell = row_cells[c_i]
                    _set_docx_cell_width_and_margins(t_cell, col_widths_pt[c_i], v_pad_tw=v_pad_tw, h_pad_tw=h_pad_tw)
                    t_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

                    if not cell_info["is_origin"]:
                        continue

                    fill_hex = cell_info["fill"] or ("F8FAFC" if r_i % 2 == 1 else "FFFFFF")
                    set_cell_background(t_cell, fill_hex)
                    dark_bg = _is_dark_hex(fill_hex)
                    fg_hex = cell_info["color"] or ("FFFFFF" if dark_bg else "0F172A")

                    p = t_cell.paragraphs[0]
                    p.alignment = align_map.get(cell_info["align"], WD_ALIGN_PARAGRAPH.LEFT)
                    p.paragraph_format.space_before = DocxPt(0)
                    p.paragraph_format.space_after = DocxPt(0)

                    txt_val = cell_info["text"]
                    if cell_info["url"]:
                        link_hex = "93C5FD" if dark_bg else "2563EB"
                        _add_docx_hyperlink(p, cell_info["url"], txt_val, docx_font_pt, hex_color=link_hex, bold=cell_info["bold"])
                    elif txt_val != "":
                        run = p.add_run(txt_val)
                        run.font.size = DocxPt(docx_font_pt)
                        run.bold = cell_info["bold"]
                        run.italic = cell_info["italic"]
                        run.font.color.rgb = DocxRGBColor.from_string(fg_hex)

            # Apply Excel merged cell spans (gridSpan + vMerge) after populating contents
            for c0, r0, c1, r1 in sh["span_cmds"]:
                try:
                    row_cells_grid[r0][c0].merge(row_cells_grid[r1][c1])
                except Exception:
                    pass

        if sh["images"]:
            for im_info in sh["images"]:
                try:
                    target_w_in = im_info["disp_w_pt"] / 72.0
                    target_h_in = im_info["disp_h_pt"] / 72.0
                    max_w_in = (avail_w_pt / 72.0) - 0.1
                    max_h_in = 5.2 if use_landscape else 6.0
                    scale = min(max_w_in / max(target_w_in, 0.1), max_h_in / max(target_h_in, 0.1), 1.0)
                    p_img = doc.add_paragraph()
                    p_img.paragraph_format.space_before = DocxPt(6)
                    p_img.paragraph_format.space_after = DocxPt(4)
                    p_img.add_run().add_picture(io.BytesIO(im_info["bytes"]), width=DocxInches(target_w_in * scale))
                except Exception as img_err:
                    logger.warning(f"Failed adding XLSX image to DOCX on '{sh['title']}': {img_err}")

    out_io = io.BytesIO()
    try:
        doc.save(out_io)
        return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", sheet_count
    finally:
        sheets.clear()
        gc.collect()


def _reset_style_list_counters(list_counters: Dict[Tuple[str, int], int]) -> None:
    """Resets style-based list counters when entering a new heading section."""
    for key in [k for k in list_counters if k[0] in ("style_number", "style_bullet")]:
        del list_counters[key]


def _next_ordered_list_num(list_counters: Dict[Tuple[str, int], int], num_id: str, ilvl: int) -> int:
    """Increments an ordered list counter at ilvl and resets any deeper sub-level counters."""
    for key in [k for k in list_counters if k[0] == num_id and k[1] > ilvl]:
        del list_counters[key]
    key = (num_id, ilvl)
    list_counters[key] = list_counters.get(key, 0) + 1
    return list_counters[key]


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
            if hl > 0:
                _reset_style_list_counters(list_counters)
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
                    num_val = _next_ordered_list_num(list_counters, li["num_id"], ilvl)
                    display_text = f"{indent}{num_val}. {text}"
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
    base_font_size: Optional[float] = None,
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
            and merged[-1].get("size") == seg.get("size")
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
        if base_font_size is not None and seg.get("size") and abs(float(seg["size"]) - base_font_size) >= 0.8:
            txt = f'<font size="{float(seg["size"]):.1f}">{txt}</font>'
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
                    if hl > 0:
                        _reset_style_list_counters(list_counters)
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
                            num_val = _next_ordered_list_num(list_counters, li["num_id"], ilvl)
                            prefix = f"{num_val}. "
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
    sans_family, _ = _ensure_reportlab_unicode_fonts()
    bold_family = f"{sans_family}-Bold" if sans_family == "VelmoSans" else "Helvetica-Bold"
    italic_family = f"{sans_family}-Italic" if sans_family == "VelmoSans" else "Helvetica-Oblique"

    sheets = _extract_xlsx_workbook_model(xlsx_bytes)
    buffer = io.BytesIO()

    # Build per-sheet Portrait/Landscape geometry
    page_templates: List[PageTemplate] = []
    sheet_geoms: List[dict] = []

    for s_idx, sh in enumerate(sheets):
        n_cols = len(sh["rows"][0]) if sh["rows"] else 1
        est_hint_pt = sum(sh.get("col_width_hints", [])) * 5.2
        use_landscape = n_cols >= 7 or est_hint_pt > 530.0
        pw, ph = landscape(letter) if use_landscape else letter
        lm = rm = 24.0 if use_landscape else 28.0
        tm = bm = 34.0
        avail_w = pw - lm - rm
        avail_h = ph - tm - bm
        geom = {"pw": pw, "ph": ph, "lm": lm, "rm": rm, "tm": tm, "bm": bm, "avail_w": avail_w, "avail_h": avail_h, "title": sh["title"]}
        sheet_geoms.append(geom)

        def _make_on_page(g: dict):
            def _on_page(canvas_obj, doc_obj):
                canvas_obj.saveState()
                canvas_obj.setFont(sans_family, 7.5)
                canvas_obj.setFillColor(colors.HexColor("#64748B"))
                canvas_obj.drawString(g["lm"], g["ph"] - 20.0, f"Sheet: {g['title']}"[:90])
                canvas_obj.setStrokeColor(colors.HexColor("#E2E8F0"))
                canvas_obj.setLineWidth(0.5)
                canvas_obj.line(g["lm"], g["ph"] - 24.0, g["pw"] - g["rm"], g["ph"] - 24.0)
                canvas_obj.restoreState()

                def _deferred_footer(c_obj, p_no: int, tot_p: int):
                    c_obj.saveState()
                    c_obj.setStrokeColor(colors.HexColor("#E2E8F0"))
                    c_obj.setLineWidth(0.5)
                    c_obj.line(g["lm"], 24.0, g["pw"] - g["rm"], 24.0)
                    c_obj.setFont(sans_family, 7.5)
                    c_obj.setFillColor(colors.HexColor("#64748B"))
                    c_obj.drawString(g["lm"], 13.0, f"Page {p_no} of {tot_p}")
                    c_obj.restoreState()

                canvas_obj._velmo_footer_cb = _deferred_footer
            return _on_page

        frame = Frame(lm, bm, avail_w, avail_h, id=f"xlsx_frame_{s_idx}", leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        page_templates.append(PageTemplate(id=f"xlsx_sec_{s_idx}", frames=[frame], pagesize=(pw, ph), onPage=_make_on_page(geom)))

    if not page_templates:
        pw, ph = letter
        frame = Frame(28, 34, pw - 56, ph - 68, id="xlsx_empty")
        page_templates.append(PageTemplate(id="xlsx_sec_0", frames=[frame], pagesize=(pw, ph)))
        sheet_geoms.append({"pw": pw, "ph": ph, "lm": 28.0, "rm": 28.0, "tm": 34.0, "bm": 34.0, "avail_w": pw - 56, "avail_h": ph - 68, "title": "Workbook"})

    pdf_doc = BaseDocTemplate(
        buffer,
        pagesize=(sheet_geoms[0]["pw"], sheet_geoms[0]["ph"]),
        leftMargin=sheet_geoms[0]["lm"],
        rightMargin=sheet_geoms[0]["rm"],
        topMargin=sheet_geoms[0]["tm"],
        bottomMargin=sheet_geoms[0]["bm"],
        pageTemplates=page_templates,
    )

    styles = getSampleStyleSheet()
    h1_style = ParagraphStyle(
        "XlsxH1", parent=styles["Heading1"], fontName=bold_family,
        fontSize=13.5, leading=17, spaceAfter=6, textColor=colors.HexColor("#0F172A"), keepWithNext=True,
    )
    empty_style = ParagraphStyle(
        "XlsxEmpty", parent=styles["Normal"], fontName=italic_family,
        fontSize=9, leading=12, textColor=colors.HexColor("#64748B"),
    )

    story = []
    keepalive_bufs: List[io.BytesIO] = []
    align_map = {"left": TA_LEFT, "center": TA_CENTER, "right": TA_RIGHT}

    for s_idx, sh in enumerate(sheets):
        geom = sheet_geoms[s_idx]
        if s_idx > 0:
            story.append(NextPageTemplate(f"xlsx_sec_{s_idx}"))
            story.append(PageBreak())

        story.append(RLParagraph(html.escape(sh["title"]), h1_style))

        if sh["is_empty"] and not sh["images"]:
            story.append(RLParagraph("(Empty worksheet — no data)", empty_style))
            continue

        if not sh["is_empty"] and sh["rows"]:
            col_widths, font_pt = _compute_xlsx_col_widths(sh, geom["avail_w"], sans_family=sans_family, bold_family=bold_family)
            leading_pt = round(font_pt * 1.26, 1)
            h_pad = 2.5 if len(col_widths) >= 10 else 3.8
            v_pad = 2.5 if len(sh["rows"]) > 40 else 3.5

            ts_cmds: List[Any] = [
                ("GRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#CBD5E1")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), v_pad),
                ("BOTTOMPADDING", (0, 0), (-1, -1), v_pad),
                ("LEFTPADDING", (0, 0), (-1, -1), h_pad),
                ("RIGHTPADDING", (0, 0), (-1, -1), h_pad),
            ]

            table_data = []
            for r_idx, row in enumerate(sh["rows"]):
                r_cells = []
                for c_idx, cell in enumerate(row):
                    if not cell["is_origin"]:
                        r_cells.append("")
                        continue

                    fill_hex = cell["fill"] or ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
                    ts_cmds.append(("BACKGROUND", (c_idx, r_idx), (c_idx, r_idx), colors.HexColor(f"#{fill_hex}")))
                    dark_bg = _is_dark_hex(fill_hex)
                    fg_hex = cell["color"] or ("FFFFFF" if dark_bg else "0F172A")

                    esc_txt = html.escape(cell["text"]).replace("\n", "<br/>")
                    if cell["url"]:
                        esc_url = html.escape(cell["url"], quote=True)
                        link_hex = "93C5FD" if dark_bg else "2563EB"
                        markup = f'<a href="{esc_url}" color="#{link_hex}"><u>{esc_txt}</u></a>'
                    else:
                        markup = esc_txt

                    if cell["italic"]:
                        fname = italic_family
                    elif cell["bold"]:
                        fname = bold_family
                    else:
                        fname = sans_family

                    c_style = ParagraphStyle(
                        "XlsxCellDyn",
                        fontName=fname,
                        fontSize=font_pt,
                        leading=leading_pt,
                        textColor=colors.HexColor(f"#{fg_hex}"),
                        alignment=align_map.get(cell["align"], TA_LEFT),
                    )
                    r_cells.append(RLParagraph(markup, c_style))
                table_data.append(r_cells)

            for c0, r0, c1, r1 in sh["span_cmds"]:
                ts_cmds.append(("SPAN", (c0, r0), (c1, r1)))

            rep_rows = sh["repeat_hdr_rows"] if len(sh["rows"]) > max(sh["repeat_hdr_rows"], 1) else 0
            rl_table = RLTable(table_data, colWidths=col_widths, repeatRows=rep_rows, splitByRow=1)
            rl_table.setStyle(TableStyle(ts_cmds))
            story.append(rl_table)

        if sh["images"]:
            story.append(Spacer(1, 10))
            for im_info in sh["images"]:
                try:
                    target_w = im_info["disp_w_pt"]
                    target_h = im_info["disp_h_pt"]
                    max_w = geom["avail_w"] - 4.0
                    max_h = geom["avail_h"] * 0.52
                    scale = min(max_w / max(target_w, 1.0), max_h / max(target_h, 1.0), 1.0)
                    buf = io.BytesIO(im_info["bytes"])
                    keepalive_bufs.append(buf)
                    rl_img = RLImage(buf, width=target_w * scale, height=target_h * scale)
                    rl_img.hAlign = "LEFT"
                    story.append(rl_img)
                    story.append(Spacer(1, 8))
                except Exception as im_err:
                    logger.warning(f"Failed adding XLSX image to PDF on '{sh['title']}': {im_err}")

    if not story:
        story.append(RLParagraph("Empty Workbook", empty_style))

    try:
        pdf_doc.build(story, canvasmaker=_VelmoNumberedCanvas)
        return buffer.getvalue()
    finally:
        keepalive_bufs.clear()
        story.clear()
        gc.collect()


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


# --- PowerPoint (.pptx) Shared Extraction Model & Converters ---

_PPTX_NS_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
_PPTX_NS_P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
_PPTX_NS_R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"

_DEFAULT_PPTX_THEME = {
    "dk1": "000000",
    "lt1": "FFFFFF",
    "dk2": "1F497D",
    "lt2": "EEECE1",
    "accent1": "4F81BD",
    "accent2": "C0504D",
    "accent3": "9BBB59",
    "accent4": "8064A2",
    "accent5": "4BACC6",
    "accent6": "F79646",
    "hlink": "0000FF",
    "folhlink": "800080",
    "tx1": "000000",
    "bg1": "FFFFFF",
    "tx2": "1F497D",
    "bg2": "EEECE1",
}


def _extract_pptx_theme_colors(prs: Presentation) -> Dict[str, str]:
    """Extracts the presentation's theme color scheme and clrMap aliases."""
    colors_map = dict(_DEFAULT_PPTX_THEME)
    try:
        if not prs.slide_masters:
            return colors_map
        master = prs.slide_masters[0]
        # Read clrScheme from master's theme relationship
        for rel in master.part.rels.values():
            if "theme" in (rel.reltype or "").lower():
                theme_el = parse_xml(rel.target_part.blob)
                clr_scheme = theme_el.find(f".//{_PPTX_NS_A}clrScheme")
                if clr_scheme is not None:
                    for child in clr_scheme:
                        key = child.tag.split("}")[-1].lower()
                        srgb = child.find(f"{_PPTX_NS_A}srgbClr")
                        sys_clr = child.find(f"{_PPTX_NS_A}sysClr")
                        val = None
                        if srgb is not None:
                            val = (srgb.get("val") or "").upper()
                        elif sys_clr is not None:
                            val = (sys_clr.get("lastClr") or sys_clr.get("val") or "").upper()
                        if val and re.match(r"^[0-9A-F]{6}$", val):
                            colors_map[key] = val
                break

        clr_map = master._element.find(f"{_PPTX_NS_P}clrMap")
        if clr_map is not None:
            for alias in ("bg1", "tx1", "bg2", "tx2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"):
                target = clr_map.get(alias)
                if target and target.lower() in colors_map:
                    colors_map[alias.lower()] = colors_map[target.lower()]
    except Exception:
        pass
    return colors_map


def _apply_pptx_xml_color_mods(hex_rgb: str, color_el) -> str:
    """Applies DrawingML tint/shade/lumMod/lumOff color modifiers to a 6-char hex RGB."""
    if color_el is None or not re.match(r"^[0-9A-F]{6}$", hex_rgb):
        return hex_rgb
    r, g, b = int(hex_rgb[0:2], 16) / 255.0, int(hex_rgb[2:4], 16) / 255.0, int(hex_rgb[4:6], 16) / 255.0

    def _pct(tag: str) -> Optional[float]:
        el = color_el.find(f"{_PPTX_NS_A}{tag}")
        if el is None:
            return None
        try:
            return int(el.get("val", "0")) / 100000.0
        except Exception:
            return None

    tint = _pct("tint")
    shade = _pct("shade")
    lum_mod = _pct("lumMod")
    lum_off = _pct("lumOff")

    if tint is not None:
        r = r * tint + (1.0 - tint)
        g = g * tint + (1.0 - tint)
        b = b * tint + (1.0 - tint)
    if shade is not None:
        r *= shade
        g *= shade
        b *= shade
    if lum_mod is not None or lum_off is not None:
        import colorsys
        h_c, l_c, s_c = colorsys.rgb_to_hls(r, g, b)
        if lum_mod is not None:
            l_c *= lum_mod
        if lum_off is not None:
            l_c += lum_off
        l_c = max(0.0, min(1.0, l_c))
        r, g, b = colorsys.hls_to_rgb(h_c, l_c, s_c)

    ri = max(0, min(255, int(round(r * 255))))
    gi = max(0, min(255, int(round(g * 255))))
    bi = max(0, min(255, int(round(b * 255))))
    return f"{ri:02X}{gi:02X}{bi:02X}"


def _resolve_pptx_color_hex(color_fmt, theme_colors: Dict[str, str], solid_fill_el=None) -> Optional[str]:
    """Resolves a python-pptx ColorFormat or DrawingML <a:solidFill> element to a 6-char RGB hex string."""
    if solid_fill_el is not None:
        srgb = solid_fill_el.find(f"{_PPTX_NS_A}srgbClr")
        if srgb is not None:
            val = (srgb.get("val") or "").strip().upper()
            if re.match(r"^[0-9A-F]{6}$", val):
                return _apply_pptx_xml_color_mods(val, srgb)
        sch = solid_fill_el.find(f"{_PPTX_NS_A}schemeClr")
        if sch is not None:
            sval = (sch.get("val") or "").strip().lower()
            base = theme_colors.get(sval)
            if base:
                return _apply_pptx_xml_color_mods(base, sch)

    if color_fmt is None:
        return None
    try:
        ctype = color_fmt.type
        if ctype == MSO_COLOR_TYPE.RGB and color_fmt.rgb is not None:
            hx = str(color_fmt.rgb).upper()
            return hx if re.match(r"^[0-9A-F]{6}$", hx) else None
        if ctype == MSO_COLOR_TYPE.SCHEME:
            tname = str(color_fmt.theme_color).split(".")[-1].lower()
            alias_map = {
                "dark_1": "dk1", "light_1": "lt1", "dark_2": "dk2", "light_2": "lt2",
                "text_1": "tx1", "background_1": "bg1", "text_2": "tx2", "background_2": "bg2",
                "accent_1": "accent1", "accent_2": "accent2", "accent_3": "accent3",
                "accent_4": "accent4", "accent_5": "accent5", "accent_6": "accent6",
                "hyperlink": "hlink", "followed_hyperlink": "folhlink",
            }
            key = alias_map.get(tname, tname)
            base = theme_colors.get(key)
            if base:
                br = float(getattr(color_fmt, "brightness", 0.0) or 0.0)
                return _apply_excel_tint(base, br) if br else base
    except Exception:
        pass
    return None


def _resolve_pptx_fill_hex(fill_obj, theme_colors: Dict[str, str], slide_bg_hex: Optional[str] = None) -> Optional[str]:
    """Resolves solid/gradient/background fill from a PPTX FillFormat."""
    if fill_obj is None:
        return None
    try:
        ftype = fill_obj.type
        if ftype == 1:  # SOLID
            return _resolve_pptx_color_hex(fill_obj.fore_color, theme_colors)
        if ftype == 3:  # GRADIENT
            stops = getattr(fill_obj, "gradient_stops", None)
            if stops:
                return _resolve_pptx_color_hex(stops[0].color, theme_colors)
        if ftype == 5:  # BACKGROUND (transparent over underlying canvas/cards)
            return None
    except Exception:
        pass
    return None


def _resolve_slide_bg_hex(slide, theme_colors: Dict[str, str]) -> str:
    """Resolves the effective background fill color of a slide (checking slide -> layout -> master)."""
    for container in (slide, getattr(slide, "slide_layout", None), getattr(slide, "slide_master", None)):
        if container is None:
            continue
        try:
            bg = container.background
            if bg is not None and bg.fill is not None:
                ftype = bg.fill.type
                if ftype == 1:
                    hx = _resolve_pptx_color_hex(bg.fill.fore_color, theme_colors)
                    if hx:
                        return hx
                elif ftype == 3 and bg.fill.gradient_stops:
                    hx = _resolve_pptx_color_hex(bg.fill.gradient_stops[0].color, theme_colors)
                    if hx:
                        return hx
        except Exception:
            pass
    return "FFFFFF"


def _format_pptx_autonum(num_type: str, n: int) -> str:
    """Formats DrawingML buAutoNum numbers (arabicPeriod, alphaLcPeriod, romanLcPeriod, etc.)."""
    nt = (num_type or "arabicPeriod").lower()
    if "alphalc" in nt:
        return f"{chr(ord('a') + ((n - 1) % 26))}."
    if "alphauc" in nt:
        return f"{chr(ord('A') + ((n - 1) % 26))}."
    if "romanlc" in nt:
        vals = [(10, "x"), (9, "ix"), (5, "v"), (4, "iv"), (1, "i")]
        rem = n
        res = []
        for v, sym in vals:
            while rem >= v:
                res.append(sym)
                rem -= v
        return f"{''.join(res) or str(n)}."
    if "parenr" in nt:
        return f"{n})"
    return f"{n}."


def _extract_pptx_text_frame(
    tf,
    part_rels,
    theme_colors: Dict[str, str],
    default_fg_hex: str,
    is_placeholder: bool = False,
) -> Tuple[List[Dict[str, Any]], bool]:
    """Extracts ordered paragraphs, inline styled runs, hyperlinks, bullets/numbers, and code-block flag."""
    tx_body = tf._txBody
    lst_style = tx_body.find(f"{_PPTX_NS_A}lstStyle")
    level_defaults: Dict[int, Dict[str, Any]] = {}
    if lst_style is not None:
        for lvl_idx in range(9):
            lvl_el = lst_style.find(f"{_PPTX_NS_A}lvl{lvl_idx + 1}pPr")
            if lvl_el is not None:
                d_rpr = lvl_el.find(f"{_PPTX_NS_A}defRPr")
                sz = None
                col = None
                if d_rpr is not None:
                    if d_rpr.get("sz") and d_rpr.get("sz").isdigit():
                        sz = int(d_rpr.get("sz")) / 100.0
                    sf = d_rpr.find(f"{_PPTX_NS_A}solidFill")
                    if sf is not None:
                        col = _resolve_pptx_color_hex(None, theme_colors, sf)
                level_defaults[lvl_idx] = {"size": sz, "color": col}

    paragraphs_out: List[Dict[str, Any]] = []
    ord_counters: Dict[int, int] = {}
    mono_chars = 0
    total_chars = 0

    for p_el in tx_body.findall(f"{_PPTX_NS_A}p"):
        pPr = p_el.find(f"{_PPTX_NS_A}pPr")
        lvl = 0
        align_val = TA_LEFT
        list_info: Optional[Dict[str, Any]] = None
        def_rpr = None
        space_after_pt = 3.0
        space_before_pt = 0.0

        if pPr is not None:
            try:
                lvl = max(0, min(8, int(pPr.get("lvl", "0"))))
            except Exception:
                lvl = 0
            algn = (pPr.get("algn") or "").lower()
            if algn == "ctr":
                align_val = TA_CENTER
            elif algn == "r":
                align_val = TA_RIGHT
            elif algn == "just":
                align_val = TA_JUSTIFY

            spc_a = pPr.find(f"{_PPTX_NS_A}spcAft/{_PPTX_NS_A}spcPts")
            if spc_a is not None and (spc_a.get("val") or "").isdigit():
                space_after_pt = int(spc_a.get("val")) / 100.0
            spc_b = pPr.find(f"{_PPTX_NS_A}spcBef/{_PPTX_NS_A}spcPts")
            if spc_b is not None and (spc_b.get("val") or "").isdigit():
                space_before_pt = int(spc_b.get("val")) / 100.0

            def_rpr = pPr.find(f"{_PPTX_NS_A}defRPr")

            has_bu_none = pPr.find(f"{_PPTX_NS_A}buNone") is not None
            bu_auto = pPr.find(f"{_PPTX_NS_A}buAutoNum")
            bu_char = pPr.find(f"{_PPTX_NS_A}buChar")
            if not has_bu_none:
                if bu_auto is not None:
                    num_type = bu_auto.get("type", "arabicPeriod")
                    start_at = 1
                    try:
                        start_at = int(bu_auto.get("startAt", "1"))
                    except Exception:
                        start_at = 1
                    for k in [k for k in ord_counters if k > lvl]:
                        del ord_counters[k]
                    if lvl not in ord_counters:
                        ord_counters[lvl] = start_at
                    else:
                        ord_counters[lvl] += 1
                    num_val = ord_counters[lvl]
                    list_info = {
                        "ordered": True,
                        "level": lvl,
                        "prefix": _format_pptx_autonum(num_type, num_val) + " ",
                        "num": num_val,
                    }
                elif bu_char is not None:
                    ch = bu_char.get("char", "•")
                    if ord(ch[0]) > 0xF000 or ch in ("n", "q", "v", "l", "Ø"):
                        ch = ["•", "◦", "▪"][min(lvl, 2)]
                    list_info = {"ordered": False, "level": lvl, "prefix": f"{ch} "}
                elif lvl > 0:
                    ch = ["•", "◦", "▪"][min(lvl, 2)]
                    list_info = {"ordered": False, "level": lvl, "prefix": f"{ch} "}

        if list_info is None:
            ord_counters.clear()

        def_sz = level_defaults.get(lvl, {}).get("size") or 13.0
        def_col = level_defaults.get(lvl, {}).get("color") or default_fg_hex
        def_bold = False
        def_italic = False
        if def_rpr is not None:
            if def_rpr.get("sz") and def_rpr.get("sz").isdigit():
                def_sz = int(def_rpr.get("sz")) / 100.0
            if def_rpr.get("b") in ("1", "true"):
                def_bold = True
            if def_rpr.get("i") in ("1", "true"):
                def_italic = True
            sf = def_rpr.find(f"{_PPTX_NS_A}solidFill")
            if sf is not None:
                c_res = _resolve_pptx_color_hex(None, theme_colors, sf)
                if c_res:
                    def_col = c_res

        segments: List[Dict[str, Any]] = []
        p_max_pt = def_sz

        def _parse_run_xml(r_el, override_text: Optional[str] = None):
            nonlocal mono_chars, total_chars, p_max_pt
            t_el = r_el.find(f"{_PPTX_NS_A}t")
            txt = override_text if override_text is not None else ((t_el.text or "") if t_el is not None else "")
            if not txt:
                return
            rPr = r_el.find(f"{_PPTX_NS_A}rPr")
            bold = def_bold
            italic = def_italic
            underline = False
            sz_pt = def_sz
            col_hex = def_col
            font_name = ""
            url = None

            if rPr is not None:
                if rPr.get("b") is not None:
                    bold = rPr.get("b") not in ("0", "false")
                if rPr.get("i") is not None:
                    italic = rPr.get("i") not in ("0", "false")
                if rPr.get("u") is not None and rPr.get("u") != "none":
                    underline = True
                if rPr.get("sz") and rPr.get("sz").isdigit():
                    sz_pt = int(rPr.get("sz")) / 100.0

                sf = rPr.find(f"{_PPTX_NS_A}solidFill")
                if sf is not None:
                    c_res = _resolve_pptx_color_hex(None, theme_colors, sf)
                    if c_res:
                        col_hex = c_res

                for f_tag in ("latin", "cs", "ea"):
                    f_el = rPr.find(f"{_PPTX_NS_A}{f_tag}")
                    if f_el is not None and f_el.get("typeface"):
                        font_name = f_el.get("typeface")
                        break

                hlink = rPr.find(f"{_PPTX_NS_A}hlinkClick")
                if hlink is not None:
                    rid = hlink.get(f"{_PPTX_NS_R}id")
                    if rid and part_rels and rid in part_rels:
                        url = part_rels[rid].target_ref
                    if url and not underline:
                        underline = True
                    if url and sf is None:
                        col_hex = "93C5FD" if _is_dark_hex(default_fg_hex) is False and default_fg_hex == "FFFFFF" else "2563EB"

            is_mono = _is_mono_font(font_name)
            n_ns = len(txt.strip())
            total_chars += n_ns
            if is_mono:
                mono_chars += n_ns
            p_max_pt = max(p_max_pt, sz_pt)
            segments.append({
                "text": txt,
                "url": url,
                "bold": bold,
                "italic": italic,
                "underline": underline,
                "color": col_hex,
                "mono": is_mono,
                "size": sz_pt,
            })

        for ch in p_el:
            tag = ch.tag
            if tag == f"{_PPTX_NS_A}r":
                _parse_run_xml(ch)
            elif tag == f"{_PPTX_NS_A}fld":
                _parse_run_xml(ch)
            elif tag == f"{_PPTX_NS_A}br":
                segments.append({
                    "text": "\n",
                    "url": None,
                    "bold": False,
                    "italic": False,
                    "underline": False,
                    "color": def_col,
                    "mono": False,
                    "size": def_sz,
                })

        plain = "".join(s["text"] for s in segments)
        paragraphs_out.append({
            "segments": segments,
            "plain_text": plain,
            "level": lvl,
            "align": align_val,
            "list_info": list_info,
            "size_pt": p_max_pt,
            "space_before": space_before_pt,
            "space_after": space_after_pt,
        })

    non_empty_count = sum(1 for p in paragraphs_out if p["plain_text"].strip())
    all_plain = "\n".join(p["plain_text"] for p in paragraphs_out)
    is_mono_block = (
        total_chars > 0
        and (mono_chars / total_chars) >= 0.75
        and (non_empty_count >= 2 or "\n" in all_plain)
    )
    return paragraphs_out, is_mono_block


def _extract_pptx_table_data(
    tbl_shape,
    part_rels,
    theme_colors: Dict[str, str],
    slide_bg_hex: str,
) -> Dict[str, Any]:
    """Extracts a PPTX table preserving merged cells (gridSpan/rowSpan), cell shading, contrast colors, and hyperlinks."""
    tbl = tbl_shape.table
    n_r = len(tbl.rows)
    n_c = len(tbl.columns)
    col_widths_pt = [max((c.width or 914400) / 12700.0, 12.0) for c in tbl.columns]
    row_heights_pt = [max((r.height or 228600) / 12700.0, 14.0) for r in tbl.rows]

    tbl_pr = tbl._tbl.find(f"{_PPTX_NS_A}tblPr")
    has_first_row_hdr = tbl_pr is not None and tbl_pr.get("firstRow") == "1"

    raw_rows: List[List[Dict[str, Any]]] = []
    span_cmds: List[Tuple[int, int, int, int]] = []

    for r_i in range(n_r):
        row_cells: List[Dict[str, Any]] = []
        for c_i in range(n_c):
            cell = tbl.cell(r_i, c_i)
            is_spanned = bool(getattr(cell, "is_spanned", False))
            is_origin = not is_spanned
            c_span = max(1, int(getattr(cell, "span_width", 1))) if getattr(cell, "is_merge_origin", False) else 1
            r_span = max(1, int(getattr(cell, "span_height", 1))) if getattr(cell, "is_merge_origin", False) else 1
            if is_origin and (c_span > 1 or r_span > 1):
                span_cmds.append((c_i, r_i, min(n_c - 1, c_i + c_span - 1), min(n_r - 1, r_i + r_span - 1)))

            fill_hex = _resolve_pptx_fill_hex(cell.fill, theme_colors, slide_bg_hex)
            if not fill_hex:
                if r_i == 0 and has_first_row_hdr:
                    fill_hex = "1E3A8A"
                else:
                    fill_hex = "F8FAFC" if r_i % 2 == 1 else "FFFFFF"

            dark_cell = _is_dark_hex(fill_hex)
            def_fg = "FFFFFF" if dark_cell else "0F172A"
            paras, _ = _extract_pptx_text_frame(cell.text_frame, part_rels, theme_colors, def_fg)

            # Guarantee legible contrast if source text inherited dark color on dark cell fill
            for p in paras:
                for seg in p["segments"]:
                    if dark_cell and seg["color"] in (None, "000000", "0F172A", "1E293B", "334155"):
                        seg["color"] = "FFFFFF"

            row_cells.append({
                "is_origin": is_origin,
                "col_span": c_span,
                "row_span": r_span,
                "fill": fill_hex,
                "paras": paras,
                "plain_text": "\n".join(p["plain_text"].strip() for p in paras if p["plain_text"].strip()).strip(),
            })
        raw_rows.append(row_cells)

    return {
        "col_widths_pt": col_widths_pt,
        "row_heights_pt": row_heights_pt,
        "raw_rows": raw_rows,
        "span_cmds": span_cmds,
        "repeat_hdr_rows": 1 if (n_r > 1 and (has_first_row_hdr or len(span_cmds) > 0)) else 1,
    }


def _extract_pptx_deck_model(pptx_bytes: bytes) -> Dict[str, Any]:
    """
    Parses a PPTX presentation into a rich slide-by-slide structural model.
    Recursively flattens GroupShapes with exact world coordinates, resolves theme/RGB fills,
    extracts bullets/numbering/inline styling/hyperlinks/code blocks, merged tables, images, and speaker notes.
    """
    prs = Presentation(io.BytesIO(pptx_bytes))
    slide_w_pt = (prs.slide_width or Inches(13.333)) / 12700.0
    slide_h_pt = (prs.slide_height or Inches(7.5)) / 12700.0
    theme_colors = _extract_pptx_theme_colors(prs)

    slides_out: List[Dict[str, Any]] = []

    for s_idx, slide in enumerate(prs.slides, start=1):
        slide_bg_hex = _resolve_slide_bg_hex(slide, theme_colors)
        slide_dark = _is_dark_hex(slide_bg_hex)
        part_rels = slide.part.rels
        items: List[Dict[str, Any]] = []

        def _walk_shapes(shapes, ox_emu: float, oy_emu: float, sx: float, sy: float, in_group: bool):
            for shp in shapes:
                lx = float(getattr(shp, "left", 0) or 0)
                ly = float(getattr(shp, "top", 0) or 0)
                lw = float(getattr(shp, "width", 0) or 0)
                lh = float(getattr(shp, "height", 0) or 0)

                wx_emu = ox_emu + lx * sx
                wy_emu = oy_emu + ly * sy
                ww_emu = lw * sx
                wh_emu = lh * sy

                if shp.shape_type == MSO_SHAPE_TYPE.GROUP:
                    grp_el = shp._element
                    xfrm = grp_el.find(f"{_PPTX_NS_P}grpSpPr/{_PPTX_NS_A}xfrm")
                    if xfrm is not None:
                        off = xfrm.find(f"{_PPTX_NS_A}off")
                        ext = xfrm.find(f"{_PPTX_NS_A}ext")
                        ch_off = xfrm.find(f"{_PPTX_NS_A}chOff")
                        ch_ext = xfrm.find(f"{_PPTX_NS_A}chExt")
                        gx = float(off.get("x", "0")) if off is not None else lx
                        gy = float(off.get("y", "0")) if off is not None else ly
                        gcx = float(ext.get("cx", "0")) if ext is not None else lw
                        gcy = float(ext.get("cy", "0")) if ext is not None else lh
                        cx0 = float(ch_off.get("x", "0")) if ch_off is not None else 0.0
                        cy0 = float(ch_off.get("y", "0")) if ch_off is not None else 0.0
                        ccx = float(ch_ext.get("cx", "0")) if ch_ext is not None else gcx
                        ccy = float(ch_ext.get("cy", "0")) if ch_ext is not None else gcy

                        wgx = ox_emu + gx * sx
                        wgy = oy_emu + gy * sy
                        nsx = sx * (gcx / ccx if ccx else 1.0)
                        nsy = sy * (gcy / ccy if ccy else 1.0)
                        _walk_shapes(shp.shapes, wgx - cx0 * nsx, wgy - cy0 * nsy, nsx, nsy, True)
                    else:
                        _walk_shapes(shp.shapes, ox_emu, oy_emu, sx, sy, True)
                    continue

                x_pt = wx_emu / 12700.0
                y_pt = wy_emu / 12700.0
                w_pt = max(ww_emu / 12700.0, 1.0)
                h_pt = max(wh_emu / 12700.0, 1.0)

                # Picture (including populated Picture Placeholder)
                img_blob = None
                try:
                    if shp.shape_type == MSO_SHAPE_TYPE.PICTURE or (hasattr(shp, "image") and shp.image is not None):
                        img_blob = shp.image.blob
                except Exception:
                    img_blob = None

                if img_blob is not None:
                    has_alpha = False
                    img_fmt = "PNG"
                    try:
                        with Image.open(io.BytesIO(img_blob)) as im:
                            img_fmt = (im.format or "PNG").upper()
                            has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
                    except Exception:
                        pass
                    items.append({
                        "kind": "picture",
                        "x": x_pt, "y": y_pt, "w": w_pt, "h": h_pt,
                        "blob": img_blob,
                        "fmt": img_fmt,
                        "has_alpha": has_alpha,
                        "slide_bg_hex": slide_bg_hex,
                        "in_group": in_group,
                    })
                    continue

                if shp.has_table:
                    tbl_data = _extract_pptx_table_data(shp, part_rels, theme_colors, slide_bg_hex)
                    items.append({
                        "kind": "table",
                        "x": x_pt, "y": y_pt, "w": w_pt, "h": h_pt,
                        "in_group": in_group,
                        **tbl_data,
                    })
                    continue

                fill_hex = _resolve_pptx_fill_hex(getattr(shp, "fill", None), theme_colors, slide_bg_hex)
                line_hex = None
                line_w_pt = 0.0
                try:
                    if shp.line and shp.line.fill and shp.line.fill.type == 1:
                        line_hex = _resolve_pptx_color_hex(shp.line.fill.fore_color, theme_colors)
                        line_w_pt = max((shp.line.width or Pt(1)) / 12700.0, 0.5)
                except Exception:
                    pass

                auto_type = None
                try:
                    auto_type = shp.auto_shape_type
                except Exception:
                    auto_type = None

                eff_bg = fill_hex or slide_bg_hex
                eff_dark = _is_dark_hex(eff_bg)
                def_fg = "FFFFFF" if eff_dark else "0F172A"

                paras: List[Dict[str, Any]] = []
                is_mono_block = False
                anchor = MSO_ANCHOR.TOP
                margins = (6.0, 6.0, 4.0, 4.0)
                if shp.has_text_frame:
                    tf = shp.text_frame
                    paras, is_mono_block = _extract_pptx_text_frame(
                        tf, part_rels, theme_colors, def_fg, is_placeholder=bool(getattr(shp, "is_placeholder", False))
                    )
                    try:
                        if tf.vertical_anchor is not None:
                            anchor = tf.vertical_anchor
                    except Exception:
                        pass
                    ml = (tf.margin_left / 12700.0) if tf.margin_left is not None else 6.0
                    mr = (tf.margin_right / 12700.0) if tf.margin_right is not None else 6.0
                    mt = (tf.margin_top / 12700.0) if tf.margin_top is not None else 4.0
                    mb = (tf.margin_bottom / 12700.0) if tf.margin_bottom is not None else 4.0
                    margins = (max(ml, 2.0), max(mr, 2.0), max(mt, 2.0), max(mb, 2.0))

                plain_text = "\n".join(p["plain_text"].strip() for p in paras if p["plain_text"].strip()).strip()
                if fill_hex is None and line_hex is None and not plain_text:
                    continue

                # Ensure contrast if text color matches effective background (e.g. unstyled run on dark bg)
                if eff_dark:
                    for p in paras:
                        for seg in p["segments"]:
                            if seg["color"] in (None, "000000", "0F172A", "1E293B"):
                                seg["color"] = "FFFFFF"

                is_title_placeholder = False
                try:
                    is_title_placeholder = (shp == slide.shapes.title)
                except Exception:
                    pass

                items.append({
                    "kind": "shape",
                    "x": x_pt, "y": y_pt, "w": w_pt, "h": h_pt,
                    "shape_type": shp.shape_type,
                    "auto_shape_type": auto_type,
                    "fill_hex": fill_hex,
                    "line_hex": line_hex,
                    "line_w_pt": line_w_pt,
                    "anchor": anchor,
                    "margins": margins,
                    "paras": paras,
                    "plain_text": plain_text,
                    "is_mono_block": is_mono_block,
                    "is_title_placeholder": is_title_placeholder,
                    "in_group": in_group,
                    "eff_dark": eff_dark,
                })

        _walk_shapes(slide.shapes, 0.0, 0.0, 1.0, 1.0, False)

        # Determine slide title + optional subtitle
        title_item_idx: Optional[int] = None
        slide_title = ""
        for idx, it in enumerate(items):
            if it["kind"] == "shape" and it["is_title_placeholder"] and it["plain_text"]:
                title_item_idx = idx
                first_line = next((p["plain_text"].strip() for p in it["paras"] if p["plain_text"].strip()), "")
                slide_title = first_line
                break

        if not slide_title:
            # Topmost non-group text shape with large font or near top band
            best_score = None
            for idx, it in enumerate(items):
                if it["kind"] != "shape" or not it["plain_text"] or it["in_group"] or it["is_mono_block"]:
                    continue
                first_p = next((p for p in it["paras"] if p["plain_text"].strip()), None)
                if not first_p:
                    continue
                cand = first_p["plain_text"].strip()
                if len(cand) > 150:
                    continue
                sz = first_p.get("size_pt", 13.0)
                if it["y"] < slide_h_pt * 0.45 or sz >= 18.0:
                    score = (0 if it["y"] < 100.0 else 1, -sz, it["y"], it["x"])
                    if best_score is None or score < best_score:
                        best_score = score
                        title_item_idx = idx
                        slide_title = cand

        # Extract speaker notes
        speaker_notes = ""
        try:
            if slide.has_notes_slide and slide.notes_slide.notes_text_frame:
                speaker_notes = slide.notes_slide.notes_text_frame.text.strip()
        except Exception:
            speaker_notes = ""

        # Check if completely blank or image-only
        has_text = any(it["kind"] == "shape" and it["plain_text"] for it in items) or any(it["kind"] == "table" for it in items)
        has_pics = any(it["kind"] == "picture" for it in items)
        is_blank = (not items) or (not has_text and not has_pics and not any(it.get("fill_hex") for it in items))
        is_image_only = has_pics and not has_text

        slides_out.append({
            "index": s_idx,
            "bg_hex": slide_bg_hex,
            "is_dark": slide_dark,
            "items": items,
            "title_item_idx": title_item_idx,
            "title": slide_title or f"Slide {s_idx}",
            "has_explicit_title": bool(slide_title),
            "speaker_notes": speaker_notes,
            "is_blank": is_blank,
            "is_image_only": is_image_only,
        })

    return {
        "slide_w_pt": slide_w_pt,
        "slide_h_pt": slide_h_pt,
        "slides": slides_out,
    }


def _pptx_to_pdf_reportlab(pptx_bytes: bytes) -> bytes:
    """
    Renders PowerPoint slides onto an exact-dimension ReportLab PDF canvas preserving:
    - True 16:9 vs 4:3 (or custom) slide geometry
    - Dark & light slide backgrounds, shape fills, rounded cards, arrows & callouts
    - Recursively unwrapped GroupShapes with exact world coordinates
    - Transparent RGBA PNGs via PDF SMask (zero black boxes / zero white patches on dark slides)
    - Bundled DejaVu Sans & DejaVu Sans Mono TTF fonts for crisp ₹, €, £, —, ±, α, β, μ, ≤, ≥, °C, Ω
    - Merged & shaded tables with auto-fit wrapping and clickable hyperlinks
    """
    sans_family, mono_family = _ensure_reportlab_unicode_fonts()
    model = _extract_pptx_deck_model(pptx_bytes)
    sw = model["slide_w_pt"]
    sh = model["slide_h_pt"]
    slides = model["slides"]

    buffer = io.BytesIO()
    c = rl_canvas.Canvas(buffer, pagesize=(sw, sh))
    img_reader_cache: Dict[int, Tuple[ImageReader, io.BytesIO]] = {}

    def _get_image_reader(blob: bytes, has_alpha: bool, fmt: str) -> ImageReader:
        key = hash(blob)
        if key in img_reader_cache:
            return img_reader_cache[key][0]
        if has_alpha or fmt not in ("PNG", "JPEG", "JPG"):
            with Image.open(io.BytesIO(blob)) as im:
                conv = im.convert("RGBA") if has_alpha else im.convert("RGB")
                ibuf = io.BytesIO()
                conv.save(ibuf, format="PNG")
                ibuf.seek(0)
        else:
            ibuf = io.BytesIO(blob)
        reader = ImageReader(ibuf)
        img_reader_cache[key] = (reader, ibuf)
        return reader

    def _draw_autoshape_geometry(item: Dict[str, Any], x: float, yb: float, w: float, h: float):
        fill_hex = item.get("fill_hex")
        line_hex = item.get("line_hex")
        line_w = item.get("line_w_pt", 0.75)
        do_fill = 1 if fill_hex else 0
        do_stroke = 1 if line_hex else 0
        if not do_fill and not do_stroke:
            return

        c.saveState()
        if do_fill:
            c.setFillColor(colors.HexColor(f"#{fill_hex}"))
        if do_stroke:
            c.setStrokeColor(colors.HexColor(f"#{line_hex}"))
            c.setLineWidth(line_w)

        atype = item.get("auto_shape_type")
        if atype == MSO_SHAPE.OVAL:
            c.ellipse(x, yb, x + w, yb + h, stroke=do_stroke, fill=do_fill)
        elif atype in (MSO_SHAPE.ROUNDED_RECTANGLE, MSO_SHAPE.RECTANGULAR_CALLOUT, MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT):
            rad = max(min(w, h) * 0.10, 4.0)
            rad = min(rad, 16.0)
            c.roundRect(x, yb, w, h, radius=rad, stroke=do_stroke, fill=do_fill)
        elif atype == MSO_SHAPE.RIGHT_ARROW:
            p = c.beginPath()
            hw = x + w * 0.62
            p.moveTo(x, yb + h * 0.25)
            p.lineTo(hw, yb + h * 0.25)
            p.lineTo(hw, yb)
            p.lineTo(x + w, yb + h * 0.50)
            p.lineTo(hw, yb + h)
            p.lineTo(hw, yb + h * 0.75)
            p.lineTo(x, yb + h * 0.75)
            p.close()
            c.drawPath(p, stroke=do_stroke, fill=do_fill)
        elif atype == MSO_SHAPE.LEFT_ARROW:
            p = c.beginPath()
            hw = x + w * 0.38
            p.moveTo(x + w, yb + h * 0.25)
            p.lineTo(hw, yb + h * 0.25)
            p.lineTo(hw, yb)
            p.lineTo(x, yb + h * 0.50)
            p.lineTo(hw, yb + h)
            p.lineTo(hw, yb + h * 0.75)
            p.lineTo(x + w, yb + h * 0.75)
            p.close()
            c.drawPath(p, stroke=do_stroke, fill=do_fill)
        elif atype == MSO_SHAPE.DOWN_ARROW:
            p = c.beginPath()
            hy = yb + h * 0.38
            p.moveTo(x + w * 0.25, yb + h)
            p.lineTo(x + w * 0.25, hy)
            p.lineTo(x, hy)
            p.lineTo(x + w * 0.50, yb)
            p.lineTo(x + w, hy)
            p.lineTo(x + w * 0.75, hy)
            p.lineTo(x + w * 0.75, yb + h)
            p.close()
            c.drawPath(p, stroke=do_stroke, fill=do_fill)
        elif atype == MSO_SHAPE.UP_ARROW:
            p = c.beginPath()
            hy = yb + h * 0.62
            p.moveTo(x + w * 0.25, yb)
            p.lineTo(x + w * 0.25, hy)
            p.lineTo(x, hy)
            p.lineTo(x + w * 0.50, yb + h)
            p.lineTo(x + w, hy)
            p.lineTo(x + w * 0.75, hy)
            p.lineTo(x + w * 0.75, yb)
            p.close()
            c.drawPath(p, stroke=do_stroke, fill=do_fill)
        else:
            c.rect(x, yb, w, h, stroke=do_stroke, fill=do_fill)
        c.restoreState()

    def _draw_textbox(item: Dict[str, Any], x: float, yb: float, w: float, h: float):
        paras = item.get("paras") or []
        non_empty = [p for p in paras if p["plain_text"].strip()]
        if not non_empty:
            return

        ml, mr, mt, mb = item.get("margins", (6.0, 6.0, 4.0, 4.0))
        inner_w = max(w - ml - mr, 24.0)
        # Keep text inside slide bottom even if shape extends slightly
        eff_yb = max(yb, 4.0)
        eff_h = max((yb + h) - eff_yb, 16.0)
        inner_h = max(eff_h - mt - mb, 12.0)

        def_color = "FFFFFF" if item.get("eff_dark") else "0F172A"
        is_mono_blk = item.get("is_mono_block", False)

        def _build_flowables(scale: float):
            flows: List[Tuple[RLParagraph, float, float, float]] = []
            tot_h = 0.0
            for p_idx, p_info in enumerate(non_empty):
                segs = []
                for s in p_info["segments"]:
                    sc = dict(s)
                    if sc.get("size"):
                        sc["size"] = max(6.5, sc["size"] * scale)
                    segs.append(sc)
                dom_pt = max(6.5, p_info.get("size_pt", 13.0) * scale)
                mk = _segments_to_rl_markup(
                    segs,
                    mono_family,
                    preserve_newlines=is_mono_blk,
                    default_color=def_color,
                    strip_edges=not is_mono_blk,
                    base_font_size=dom_pt,
                )
                if not mk:
                    continue
                li = p_info.get("list_info")
                lvl = p_info.get("level", 0)
                left_ind = 0.0
                first_ind = 0.0
                if li is not None:
                    prefix = html.escape(li["prefix"])
                    mk = f"<b>{prefix}</b>{mk}" if li["ordered"] else f"{prefix}{mk}"
                    left_ind = 12.0 + lvl * 14.0
                    first_ind = -10.0
                elif lvl > 0:
                    left_ind = lvl * 14.0

                leading = round(dom_pt * 1.24, 2)
                st = ParagraphStyle(
                    f"PptxP_{p_idx}",
                    fontName=mono_family if is_mono_blk else sans_family,
                    fontSize=dom_pt,
                    leading=leading,
                    alignment=p_info.get("align", TA_LEFT),
                    leftIndent=left_ind,
                    firstLineIndent=first_ind,
                    textColor=colors.HexColor(f"#{def_color}"),
                )
                rl_p = RLParagraph(mk, st)
                _, ph = rl_p.wrapOn(c, inner_w, 10000.0)
                sb = p_info.get("space_before", 0.0) * scale if p_idx > 0 else 0.0
                sa = min(p_info.get("space_after", 2.5), 6.0) * scale
                flows.append((rl_p, ph, sb, sa))
                tot_h += sb + ph + sa
            return flows, tot_h

        scale = 1.0
        flows, tot_h = _build_flowables(scale)
        if tot_h > inner_h + 2.0 and tot_h > 0:
            scale = max(0.62, (inner_h / tot_h) * 0.96)
            flows, tot_h = _build_flowables(scale)
            if tot_h > inner_h + 2.0:
                scale = max(0.55, scale * (inner_h / max(tot_h, 1.0)) * 0.96)
                flows, tot_h = _build_flowables(scale)

        if not flows:
            return

        avail_top = eff_yb + eff_h - mt
        avail_bot = eff_yb + mb
        anchor = item.get("anchor", MSO_ANCHOR.TOP)
        if anchor == MSO_ANCHOR.MIDDLE and tot_h < inner_h:
            cur_y = avail_top - (inner_h - tot_h) / 2.0
        elif anchor == MSO_ANCHOR.BOTTOM and tot_h < inner_h:
            cur_y = avail_bot + tot_h
        else:
            cur_y = avail_top

        for rl_p, ph, sb, sa in flows:
            cur_y -= sb
            draw_y = cur_y - ph
            if draw_y >= 2.0:
                rl_p.drawOn(c, x + ml, draw_y)
            cur_y -= (ph + sa)

    def _draw_table(item: Dict[str, Any], x: float, yb: float, w: float, h: float):
        raw_rows = item.get("raw_rows") or []
        if not raw_rows:
            return
        n_c = len(raw_rows[0])
        col_ws = item.get("col_widths_pt") or [w / max(n_c, 1)] * n_c
        sum_cw = max(sum(col_ws), 1.0)
        max_avail_w = min(w, max(sw - x - 6.0, 40.0))
        scaled_ws = [cw * (max_avail_w / sum_cw) for cw in col_ws]
        top_y = min(yb + h, sh - 4.0)
        max_avail_h = max(top_y - 6.0, 40.0)

        def _build_rl_table(font_scale: float, v_pad: float, h_pad: float) -> Tuple[RLTable, float]:
            table_data = []
            ts_cmds = [
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), v_pad),
                ("BOTTOMPADDING", (0, 0), (-1, -1), v_pad),
                ("LEFTPADDING", (0, 0), (-1, -1), h_pad),
                ("RIGHTPADDING", (0, 0), (-1, -1), h_pad),
            ]
            for r_i, row in enumerate(raw_rows):
                r_out = []
                for c_i, cell in enumerate(row):
                    if not cell.get("is_origin"):
                        r_out.append("")
                        continue
                    fill_hex = cell.get("fill") or "FFFFFF"
                    ts_cmds.append(("BACKGROUND", (c_i, r_i), (c_i, r_i), colors.HexColor(f"#{fill_hex}")))
                    dark_cell = _is_dark_hex(fill_hex)
                    def_fg = "FFFFFF" if dark_cell else "0F172A"
                    cell_paras = []
                    for p_info in cell.get("paras", []):
                        if not p_info["plain_text"].strip():
                            continue
                        segs = []
                        for s in p_info["segments"]:
                            sc = dict(s)
                            if sc.get("size"):
                                sc["size"] = max(6.0, sc["size"] * font_scale)
                            segs.append(sc)
                        f_pt = max(6.0, p_info.get("size_pt", 9.5) * font_scale)
                        mk = _segments_to_rl_markup(segs, mono_family, default_color=def_fg, base_font_size=f_pt)
                        is_b = any(s.get("bold") for s in segs)
                        st = ParagraphStyle(
                            f"PptxTC_{r_i}_{c_i}",
                            fontName=f"{sans_family}-Bold" if (is_b and sans_family == "VelmoSans") else sans_family,
                            fontSize=f_pt,
                            leading=round(f_pt * 1.24, 1),
                            alignment=p_info.get("align", TA_LEFT),
                            textColor=colors.HexColor(f"#{def_fg}"),
                        )
                        cell_paras.append(RLParagraph(mk, st))
                    r_out.append(cell_paras if len(cell_paras) > 1 else (cell_paras[0] if cell_paras else ""))
                table_data.append(r_out)

            for c0, r0, c1, r1 in item.get("span_cmds", []):
                ts_cmds.append(("SPAN", (c0, r0), (c1, r1)))

            rl_tbl = RLTable(table_data, colWidths=scaled_ws)
            rl_tbl.setStyle(TableStyle(ts_cmds))
            _, act_h = rl_tbl.wrapOn(c, max_avail_w, sh)
            return rl_tbl, act_h

        rl_tbl, act_h = _build_rl_table(1.0, 3.5, 4.5)
        if act_h > max_avail_h:
            for f_sc, vp, hp in [(0.86, 2.5, 3.5), (0.76, 2.0, 3.0), (0.68, 1.5, 2.5)]:
                rl_tbl, act_h = _build_rl_table(f_sc, vp, hp)
                if act_h <= max_avail_h:
                    break

        draw_y = max(top_y - act_h, 4.0)
        rl_tbl.drawOn(c, x, draw_y)

    try:
        if not slides:
            c.setFillColor(colors.white)
            c.rect(0, 0, sw, sh, fill=1, stroke=0)
            c.showPage()
        else:
            for s_data in slides:
                c.setPageSize((sw, sh))
                bg_hex = s_data["bg_hex"]
                c.saveState()
                c.setFillColor(colors.HexColor(f"#{bg_hex}"))
                c.rect(0, 0, sw, sh, fill=1, stroke=0)
                c.restoreState()

                for item in s_data["items"]:
                    x = item["x"]
                    w = item["w"]
                    h = item["h"]
                    yb = sh - item["y"] - h
                    kind = item["kind"]

                    if kind == "picture":
                        try:
                            reader = _get_image_reader(item["blob"], item["has_alpha"], item["fmt"])
                            c.drawImage(
                                reader, x, yb, width=w, height=h,
                                mask="auto" if item["has_alpha"] else None,
                                preserveAspectRatio=False,
                            )
                        except Exception as img_err:
                            logger.warning(f"PPTX->PDF image draw failed on slide {s_data['index']}: {img_err}")
                    elif kind == "table":
                        _draw_table(item, x, yb, w, h)
                    elif kind == "shape":
                        _draw_autoshape_geometry(item, x, yb, w, h)
                        if item.get("plain_text"):
                            _draw_textbox(item, x, yb, w, h)

                c.showPage()

        c.save()
        return buffer.getvalue()
    finally:
        img_reader_cache.clear()
        gc.collect()


def pptx_to_pdf(pptx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Converts PPTX to PDF via LibreOffice headless or exact-geometry ReportLab slide renderer."""
    pdf_bytes = convert_via_libreoffice(pptx_bytes, ".pptx")
    if not pdf_bytes:
        pdf_bytes = _pptx_to_pdf_reportlab(pptx_bytes)
    return pdf_bytes, "application/pdf", f"{base_name}.pdf", 1


def _flatten_pptx_image_for_doc(blob: bytes, has_alpha: bool, bg_hex: str = "FFFFFF") -> Tuple[bytes, int, int]:
    """Composites RGBA/LA/P images onto bg_hex (e.g., dark slide bg for white logos, white otherwise) as clean PNG."""
    with Image.open(io.BytesIO(blob)) as im:
        iw, ih = max(im.width, 1), max(im.height, 1)
        fmt = (im.format or "").upper()
        mode = im.mode
        if has_alpha or mode in ("RGBA", "LA") or (mode == "P" and "transparency" in im.info):
            hx = bg_hex if re.match(r"^[0-9A-F]{6}$", (bg_hex or "").upper()) else "FFFFFF"
            r_bg, g_bg, b_bg = int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16)
            im_rgba = im.convert("RGBA")
            bg = Image.new("RGBA", im_rgba.size, (r_bg, g_bg, b_bg, 255))
            flat = Image.alpha_composite(bg, im_rgba).convert("RGB")
            buf = io.BytesIO()
            flat.save(buf, format="PNG")
            return buf.getvalue(), iw, ih
        if fmt not in ("PNG", "JPEG", "JPG") or mode not in ("RGB", "L"):
            rgb = im.convert("RGB")
            buf = io.BytesIO()
            rgb.save(buf, format="PNG")
            return buf.getvalue(), iw, ih
        return blob, iw, ih


def _apply_docx_para_callout_style(p, bg_hex: str, border_hex: str):
    """Applies Word XML paragraph shading and left accent border to a python-docx Paragraph."""
    try:
        pPr = p._p.get_or_add_pPr()
        for tag in ("shd", "pBdr"):
            old = pPr.find(f"{{http://schemas.openxmlformats.org/wordprocessingml/2006/main}}{tag}")
            if old is not None:
                pPr.remove(old)
        pPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{bg_hex}"/>'))
        pPr.append(parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'<w:left w:val="single" w:sz="18" w:space="6" w:color="{border_hex}"/>'
            f'</w:pBdr>'
        ))
    except Exception:
        pass


def _append_pptx_segments_to_docx_para(
    p,
    segments: List[Dict[str, Any]],
    default_pt: float = 10.5,
    bg_is_dark: bool = False,
    force_mono: bool = False,
):
    """Populates a python-docx paragraph from PPTX segments preserving styles, colors, and clickable hyperlinks."""
    for seg in segments:
        txt = seg.get("text", "")
        if not txt:
            continue
        col = (seg.get("color") or "").upper()
        if bg_is_dark:
            if col in ("", "000000", "0F172A", "1E293B", "334155"):
                col = "FFFFFF"
        else:
            # On white Word background, avoid invisible white/near-white runs
            if not col or col in ("FFFFFF", "F8FAFC", "F1F5F9", "E2E8F0"):
                col = "0F172A"

        if seg.get("url"):
            link_col = "93C5FD" if bg_is_dark else "2563EB"
            _add_docx_hyperlink(p, seg["url"], txt, default_pt, hex_color=link_col, bold=bool(seg.get("bold")))
            continue

        run = p.add_run(txt)
        run.font.size = DocxPt(default_pt)
        run.bold = bool(seg.get("bold"))
        run.italic = bool(seg.get("italic"))
        run.underline = bool(seg.get("underline"))
        if force_mono or seg.get("mono"):
            run.font.name = "Consolas"
        if re.match(r"^[0-9A-F]{6}$", col) and col not in ("000000", "0F172A"):
            run.font.color.rgb = DocxRGBColor(int(col[0:2], 16), int(col[2:4], 16), int(col[4:6], 16))
        elif bg_is_dark:
            run.font.color.rgb = DocxRGBColor(255, 255, 255)


def pptx_to_docx(pptx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """
    Converts PowerPoint slides into a rich editable Word document preserving:
    - Slide titles (Heading 1) with clean page breaks per slide
    - Recursively unwrapped GroupShape KPI callout cards
    - Multi-level bulleted & numbered lists, bold/italic/underline/colored runs, monospace code blocks
    - Clickable hyperlinks, embedded images (RGBA composited cleanly), merged & shaded tables, and speaker notes
    """
    model = _extract_pptx_deck_model(pptx_bytes)
    slides = model["slides"]
    doc = Document()

    for sec in doc.sections:
        sec.left_margin = DocxInches(0.75)
        sec.right_margin = DocxInches(0.75)
        sec.top_margin = DocxInches(0.70)
        sec.bottom_margin = DocxInches(0.70)

    avail_w_in = 7.0
    ns_w = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

    for s_idx, s_data in enumerate(slides, start=1):
        h = doc.add_heading(s_data["title"], level=1)
        if s_idx > 1:
            h.paragraph_format.page_break_before = True

        if s_data["title_item_idx"] is not None:
            t_item = s_data["items"][s_data["title_item_idx"]]
            t_paras = [p for p in t_item.get("paras", []) if p["plain_text"].strip()]
            if len(t_paras) > 1:
                for p_extra in t_paras[1:]:
                    dp_sub = doc.add_paragraph()
                    _append_pptx_segments_to_docx_para(dp_sub, p_extra["segments"], default_pt=10.5, bg_is_dark=False)

        if s_data["is_blank"]:
            p_blank = doc.add_paragraph()
            r_b = p_blank.add_run("(Blank slide)")
            r_b.italic = True
            r_b.font.color.rgb = DocxRGBColor(100, 116, 139)

        # Sort body items in natural top-to-bottom, left-to-right reading order
        body_items = [
            it for idx, it in enumerate(s_data["items"])
            if idx != s_data["title_item_idx"]
        ]
        body_items.sort(key=lambda it: (round(it["y"] / 16.0), it["x"]))

        for item in body_items:
            kind = item["kind"]
            if kind == "picture":
                try:
                    bg_for_img = item["slide_bg_hex"] if _is_dark_hex(item["slide_bg_hex"]) else "FFFFFF"
                    flat_bytes, iw, ih = _flatten_pptx_image_for_doc(item["blob"], item["has_alpha"], bg_for_img)
                    w_in = min(avail_w_in, max(item["w"] / 72.0 * 0.72, 1.4))
                    if item["w"] / 72.0 > 7.5:
                        w_in = avail_w_in
                    p_img = doc.add_paragraph()
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(io.BytesIO(flat_bytes), width=DocxInches(w_in))
                except Exception as img_err:
                    logger.warning(f"PPTX->DOCX image embed failed on slide {s_idx}: {img_err}")

            elif kind == "table":
                raw_rows = item["raw_rows"]
                n_r = len(raw_rows)
                n_c = len(raw_rows[0]) if n_r else 0
                if n_r == 0 or n_c == 0:
                    continue
                doc_tbl = doc.add_table(rows=n_r, cols=n_c)
                doc_tbl.style = "Table Grid"

                col_ws = item["col_widths_pt"]
                sum_cw = max(sum(col_ws), 1.0)
                doc_col_pts = [(cw / sum_cw) * (avail_w_in * 72.0) for cw in col_ws]

                for r_i, row in enumerate(raw_rows):
                    if r_i == 0:
                        trPr = doc_tbl.rows[0]._tr.get_or_add_trPr()
                        trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
                    for c_i, cell_info in enumerate(row):
                        dcell = doc_tbl.cell(r_i, c_i)
                        _set_docx_cell_width_and_margins(dcell, doc_col_pts[c_i], v_pad_tw=40, h_pad_tw=65)
                        dcell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                        fill_hex = cell_info.get("fill") or "FFFFFF"
                        set_cell_background(dcell, fill_hex)
                        if not cell_info.get("is_origin"):
                            continue
                        dark_c = _is_dark_hex(fill_hex)
                        non_empty_paras = [p for p in cell_info["paras"] if p["plain_text"].strip()]
                        for p_i, p_info in enumerate(non_empty_paras):
                            dp = dcell.paragraphs[0] if p_i == 0 else dcell.add_paragraph()
                            dp.paragraph_format.space_before = DocxPt(0)
                            dp.paragraph_format.space_after = DocxPt(2)
                            _append_pptx_segments_to_docx_para(dp, p_info["segments"], default_pt=9.5, bg_is_dark=dark_c)

                for c0, r0, c1, r1 in item.get("span_cmds", []):
                    try:
                        doc_tbl.cell(r0, c0).merge(doc_tbl.cell(r1, c1))
                    except Exception:
                        pass
                doc.add_paragraph()

            elif kind == "shape":
                non_empty_paras = [p for p in item.get("paras", []) if p["plain_text"].strip()]
                if not non_empty_paras:
                    continue
                fill_hex = item.get("fill_hex")
                is_card = bool(fill_hex and fill_hex not in ("FFFFFF", "AUTO"))
                card_dark = is_card and _is_dark_hex(fill_hex)

                if item.get("is_mono_block"):
                    p_code = doc.add_paragraph()
                    _apply_docx_para_callout_style(p_code, "F1F5F9", "475569")
                    code_txt = "\n".join(p["plain_text"] for p in non_empty_paras)
                    r_c = p_code.add_run(code_txt)
                    r_c.font.name = "Consolas"
                    r_c.font.size = DocxPt(9.0)
                    r_c.font.color.rgb = DocxRGBColor(15, 23, 42)
                    continue

                if is_card and (item.get("in_group") or len(non_empty_paras) <= 4):
                    p_card = doc.add_paragraph()
                    border_col = item.get("line_hex") or ("38BDF8" if card_dark else "2563EB")
                    _apply_docx_para_callout_style(p_card, fill_hex, border_col)
                    for p_i, p_info in enumerate(non_empty_paras):
                        if p_i > 0:
                            p_card.add_run("\n")
                        _append_pptx_segments_to_docx_para(
                            p_card, p_info["segments"], default_pt=10.5, bg_is_dark=card_dark
                        )
                    continue

                for p_info in non_empty_paras:
                    li = p_info.get("list_info")
                    lvl = p_info.get("level", 0)
                    if li is not None:
                        base_s = "List Number" if li["ordered"] else "List Bullet"
                        style_cand = base_s if lvl == 0 else (f"{base_s} 2" if lvl == 1 else f"{base_s} 3")
                        try:
                            dp = doc.add_paragraph(style=style_cand)
                        except Exception:
                            dp = doc.add_paragraph(style=base_s)
                        if lvl > 0:
                            dp.paragraph_format.left_indent = DocxInches(0.25 * (lvl + 1))
                    else:
                        dp = doc.add_paragraph()
                        if lvl > 0:
                            dp.paragraph_format.left_indent = DocxInches(0.25 * lvl)
                    if is_card:
                        _apply_docx_para_callout_style(dp, fill_hex, item.get("line_hex") or "2563EB")
                    _append_pptx_segments_to_docx_para(
                        dp, p_info["segments"], default_pt=10.5, bg_is_dark=card_dark
                    )

        if s_data.get("speaker_notes"):
            p_note = doc.add_paragraph()
            _apply_docx_para_callout_style(p_note, "FEFCE8", "D97706")
            rn_lbl = p_note.add_run("Speaker Notes: ")
            rn_lbl.bold = True
            rn_lbl.font.size = DocxPt(9.5)
            rn_lbl.font.color.rgb = DocxRGBColor(146, 64, 14)
            rn_txt = p_note.add_run(s_data["speaker_notes"])
            rn_txt.italic = True
            rn_txt.font.size = DocxPt(9.5)
            rn_txt.font.color.rgb = DocxRGBColor(71, 85, 105)

    out_io = io.BytesIO()
    doc.save(out_io)
    gc.collect()
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document", f"{base_name}.docx", len(slides)


def _pptx_numeric_with_format(raw_text: str) -> Tuple[Any, Optional[str]]:
    """Parses PPTX numeric string (including comma-grouped numbers) and returns (value, excel_number_format)."""
    stripped = (raw_text or "").strip()
    val = parse_numeric(stripped)
    if isinstance(val, (int, float)) and not isinstance(val, bool):
        if re.match(r"^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$", stripped):
            if "." in stripped:
                dec_len = len(stripped.split(".", 1)[1])
                return val, f"#,##0.{'0' * dec_len}"
            return val, "#,##0"
    return val, None


def pptx_to_xlsx(pptx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """
    Exports PowerPoint slides to a multi-sheet Excel workbook with:
    - 'Slide Content' worksheet (titles, grouped KPI cards, lists, code blocks, clickable links)
    - 'Speaker Notes' worksheet
    - One styled worksheet per PPTX table preserving merges, contrast fills, comma-grouped numbers, and hyperlinks
    """
    model = _extract_pptx_deck_model(pptx_bytes)
    slides = model["slides"]

    wb = openpyxl.Workbook()
    default_sheet = wb.active

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1'),
    )
    hdr_fill = PatternFill(start_color="0F172A", end_color="0F172A", fill_type="solid")
    hdr_font = Font(bold=True, color="FFFFFF")

    used_titles: Set[str] = set()
    content_title = sanitize_sheet_title("Slide Content", used_titles)
    notes_title = sanitize_sheet_title("Speaker Notes", used_titles)

    ws_content = wb.create_sheet(title=content_title)
    ws_notes = wb.create_sheet(title=notes_title)

    for c_i, h_lbl in enumerate(["Slide", "Type", "Content"], start=1):
        cell = ws_content.cell(row=1, column=c_i, value=h_lbl)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.border = thin_border

    for c_i, h_lbl in enumerate(["Slide", "Slide Title", "Speaker Notes"], start=1):
        cell = ws_notes.cell(row=1, column=c_i, value=h_lbl)
        cell.fill = hdr_fill
        cell.font = hdr_font
        cell.border = thin_border

    content_row = 2
    notes_row = 2
    table_count = 0

    def _add_content_row(s_lbl: str, kind_lbl: str, text_val: str, is_bold: bool = False, url: Optional[str] = None):
        nonlocal content_row
        c1 = ws_content.cell(row=content_row, column=1, value=s_lbl)
        c2 = ws_content.cell(row=content_row, column=2, value=kind_lbl)
        c3 = ws_content.cell(row=content_row, column=3, value=text_val)
        for cc in (c1, c2, c3):
            cc.border = thin_border
        if is_bold:
            c1.font = Font(bold=True)
            c3.font = Font(bold=True)
        if url:
            c3.hyperlink = url
            c3.font = Font(bold=is_bold, underline="single", color="2563EB")
        content_row += 1

    for s_data in slides:
        s_idx = s_data["index"]
        s_lbl = f"Slide {s_idx}"
        _add_content_row(s_lbl, "Slide Title", s_data["title"], is_bold=True)

        if s_data["title_item_idx"] is not None:
            t_item = s_data["items"][s_data["title_item_idx"]]
            t_paras = [p for p in t_item.get("paras", []) if p["plain_text"].strip()]
            for p_extra in t_paras[1:]:
                url = next((s.get("url") for s in p_extra["segments"] if s.get("url")), None)
                _add_content_row(s_lbl, "Subtitle", _segments_to_txt(p_extra["segments"]).strip(), url=url)

        if s_data["is_blank"]:
            _add_content_row(s_lbl, "Blank Slide", "(Blank slide — no content)")

        body_items = [
            it for idx, it in enumerate(s_data["items"])
            if idx != s_data["title_item_idx"]
        ]
        body_items.sort(key=lambda it: (round(it["y"] / 16.0), it["x"]))

        for item in body_items:
            kind = item["kind"]
            if kind == "table":
                table_count += 1
                t_sheet_name = sanitize_sheet_title(s_data["title"] or f"Slide {s_idx} Table", used_titles)
                ws_t = wb.create_sheet(title=t_sheet_name)
                raw_rows = item["raw_rows"]
                for r_i, row in enumerate(raw_rows, start=1):
                    for c_i, cell_info in enumerate(row, start=1):
                        xl_cell = ws_t.cell(row=r_i, column=c_i)
                        xl_cell.border = thin_border
                        if not cell_info.get("is_origin"):
                            continue
                        raw_txt = cell_info.get("plain_text", "")
                        parsed_val, num_fmt = _pptx_numeric_with_format(raw_txt)
                        xl_cell.value = parsed_val
                        if num_fmt:
                            xl_cell.number_format = num_fmt

                        first_url = next(
                            (s.get("url") for p in cell_info.get("paras", []) for s in p.get("segments", []) if s.get("url")),
                            None,
                        )
                        if first_url:
                            xl_cell.hyperlink = first_url

                        fill_hex = cell_info.get("fill")
                        is_cell_bold = (r_i == 1) or any(
                            s.get("bold") for p in cell_info.get("paras", []) for s in p.get("segments", [])
                        )
                        dark_c = _is_dark_hex(fill_hex)
                        if fill_hex and fill_hex not in ("FFFFFF", "AUTO"):
                            xl_cell.fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
                        if first_url:
                            xl_cell.font = Font(bold=is_cell_bold, underline="single", color="93C5FD" if dark_c else "2563EB")
                        elif dark_c:
                            xl_cell.font = Font(bold=is_cell_bold, color="FFFFFF")
                        elif is_cell_bold:
                            xl_cell.font = Font(bold=True, color="0F172A")

                for c0, r0, c1, r1 in item.get("span_cmds", []):
                    ws_t.merge_cells(start_row=r0 + 1, start_column=c0 + 1, end_row=r1 + 1, end_column=c1 + 1)
                _auto_fit_xlsx_columns(ws_t)
                _add_content_row(s_lbl, "Table Sheet", f"Table extracted to '{t_sheet_name}'")

            elif kind == "picture":
                if s_data["is_image_only"]:
                    _add_content_row(s_lbl, "Full-Bleed Visual", f"Embedded diagram ({item['fmt']}, {int(round(item['w']))}x{int(round(item['h']))} pt)")

            elif kind == "shape":
                non_empty_paras = [p for p in item.get("paras", []) if p["plain_text"].strip()]
                if not non_empty_paras:
                    continue
                if item.get("is_mono_block"):
                    code_txt = "\n".join(p["plain_text"] for p in non_empty_paras)
                    _add_content_row(s_lbl, "Code Block", code_txt)
                    continue

                kind_lbl = "KPI Card (Group)" if item.get("in_group") else ("Callout / Shape" if item.get("fill_hex") else "Text")
                for p_info in non_empty_paras:
                    txt_line = _segments_to_txt(p_info["segments"]).strip()
                    li = p_info.get("list_info")
                    lvl = p_info.get("level", 0)
                    if li is not None:
                        txt_line = ("  " * lvl) + li["prefix"] + txt_line
                    first_url = next((s.get("url") for s in p_info["segments"] if s.get("url")), None)
                    _add_content_row(s_lbl, kind_lbl, txt_line, url=first_url)

        if s_data.get("speaker_notes"):
            c1 = ws_notes.cell(row=notes_row, column=1, value=s_lbl)
            c2 = ws_notes.cell(row=notes_row, column=2, value=s_data["title"])
            c3 = ws_notes.cell(row=notes_row, column=3, value=s_data["speaker_notes"])
            for cc in (c1, c2, c3):
                cc.border = thin_border
            notes_row += 1

    if notes_row == 2:
        c1 = ws_notes.cell(row=2, column=1, value="-")
        c2 = ws_notes.cell(row=2, column=2, value="Info")
        c3 = ws_notes.cell(row=2, column=3, value="No speaker notes present in presentation.")
        for cc in (c1, c2, c3):
            cc.border = thin_border

    _auto_fit_xlsx_columns(ws_content, min_width=12, max_width=85)
    _auto_fit_xlsx_columns(ws_notes, min_width=12, max_width=85)

    if default_sheet in wb.worksheets and len(wb.worksheets) > 1:
        wb.remove(default_sheet)

    out_io = io.BytesIO()
    wb.save(out_io)
    wb.close()
    gc.collect()
    return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", f"{base_name}.xlsx", table_count


def _estimate_textbox_height_in(lines_info: List[Tuple[str, float]], width_in: float) -> float:
    """Estimates vertical height (inches) needed by wrapped paragraphs inside a PPTX textbox."""
    total_h = 0.06
    for text, pt in lines_info:
        chars_per_line = max(int(width_in * 10.6 * (12.0 / max(pt, 7.0))), 24)
        sub_lines = text.split("\n")
        wrapped = 0
        for sl in sub_lines:
            wrapped += max(1, (len(sl) + chars_per_line - 1) // chars_per_line)
        line_h_in = (pt * 1.24) / 72.0
        total_h += wrapped * line_h_in + (pt * 0.28) / 72.0
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
    for r_i, r in enumerate(raw_rows_slice):
        max_lines = 1
        for c_i, cell in enumerate(r):
            if not cell or not cell["is_origin"]:
                continue
            # Skip wide horizontal merges when computing single-column wrapping height
            is_h_merge = any(c0 == c_i and r0 == r_i and c1 > c0 for c0, r0, c1, _ in span_cmds_slice)
            span_w_in = col_widths_in[c_i]
            if is_h_merge:
                for c0, r0, c1, _ in span_cmds_slice:
                    if c0 == c_i and r0 == r_i:
                        span_w_in = sum(col_widths_in[c0:c1 + 1])
                        break
            txt = "\n".join(p["plain_text"].strip() for p in cell["paras"] if p["plain_text"].strip())
            cw_in = max(span_w_in - 0.10, 0.35)
            cpl = max(int(cw_in * 10.0 * (12.0 / font_pt)), 6)
            lines_needed = 0
            for ln in (txt.splitlines() or [""]):
                lines_needed += max(1, (len(ln) + cpl - 1) // cpl)
            max_lines = max(max_lines, lines_needed)
        rh = max(0.24, max_lines * ((font_pt * 1.32) / 72.0) + 0.09)
        row_heights_in.append(rh)

    est_h = min(max_height, Inches(sum(row_heights_in)))
    shape = slide.shapes.add_table(n_r, n_c, left, top, width, est_h)
    tbl = shape.table

    for c_i, cw_in in enumerate(col_widths_in):
        tbl.columns[c_i].width = Inches(cw_in)
    scale_h = min(1.0, (max_height / Inches(1)) / max(sum(row_heights_in), 0.1))
    for r_i, rh in enumerate(row_heights_in):
        tbl.rows[r_i].height = Inches(max(0.21, rh * scale_h))

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

            non_empty_paras = [p for p in cell_info["paras"] if p["plain_text"].strip()]
            tf = pptx_cell.text_frame
            tf.word_wrap = True
            tf.paragraphs[0].text = ""

            has_urls = any(s.get("url") for p in non_empty_paras for s in p.get("segments", []))
            if not has_urls:
                txt = "\n".join(p["plain_text"].strip() for p in non_empty_paras).strip()
                pptx_cell.text = txt
                is_bold = is_hdr or any(
                    s.get("bold") for p in non_empty_paras for s in p.get("segments", [])
                )
                first_col = next(
                    (s.get("color") for p in non_empty_paras for s in p.get("segments", []) if s.get("color")),
                    None,
                )
                if dark_bg and first_col in (None, "000000", "0F172A", "1E293B", "334155"):
                    fg = RGBColor(255, 255, 255)
                elif first_col and re.match(r"^[0-9A-F]{6}$", first_col.upper()):
                    hx = first_col.upper()
                    fg = RGBColor(int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))
                else:
                    fg = RGBColor(255, 255, 255) if dark_bg else RGBColor(15, 23, 42)
                for p in pptx_cell.text_frame.paragraphs:
                    p.font.size = Pt(font_pt)
                    p.font.bold = is_bold
                    p.font.color.rgb = fg
            else:
                for p_idx, p_info in enumerate(non_empty_paras):
                    p = tf.paragraphs[0] if p_idx == 0 else tf.add_paragraph()
                    for seg in p_info.get("segments", []):
                        if not seg.get("text"):
                            continue
                        run = p.add_run()
                        run.text = seg["text"]
                        run.font.size = Pt(font_pt)
                        run.font.bold = bool(is_hdr or seg.get("bold"))
                        run.font.italic = bool(seg.get("italic"))
                        if seg.get("url"):
                            run.hyperlink.address = seg["url"]
                            run.font.underline = True
                            run.font.color.rgb = RGBColor(147, 197, 253) if dark_bg else RGBColor(37, 99, 235)
                        elif seg.get("color") and re.match(r"^[0-9A-F]{6}$", seg["color"].upper()) and not (
                            dark_bg and seg["color"].upper() in ("000000", "0F172A", "1E293B", "334155")
                        ):
                            hx = seg["color"].upper()
                            run.font.color.rgb = RGBColor(int(hx[0:2], 16), int(hx[2:4], 16), int(hx[4:6], 16))
                        else:
                            run.font.color.rgb = RGBColor(255, 255, 255) if dark_bg else RGBColor(15, 23, 42)

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

    body_items: List[Tuple[str, Any]] = []
    for child in doc.element.body.iterchildren():
        if child.tag.endswith('p'):
            body_items.append(('p', _extract_oxml_paragraph_data(child, doc)))
        elif child.tag.endswith('tbl'):
            body_items.append(('tbl', Table(child, doc)))

    doc_cover_title = "Document Overview"
    for b_kind, b_data in body_items[:6]:
        if b_kind == 'p' and b_data["heading_level"] > 0 and b_data["plain_text"].strip():
            doc_cover_title = b_data["plain_text"].strip()
            break

    current_slide = None
    current_title_shape = None
    slide_count = 0
    current_title = doc_cover_title
    cur_y_in = top_body_in
    slide_has_table_or_img = False
    has_cover_logo_only = False
    para_buffer: List[dict] = []
    list_counters: Dict[Tuple[str, int], int] = {}

    def start_new_slide(title: str = "Slide", flush: bool = True):
        nonlocal current_slide, current_title_shape, slide_count, current_title, cur_y_in, slide_has_table_or_img, has_cover_logo_only
        if flush:
            flush_para_buffer()
        current_slide = prs.slides.add_slide(blank_layout)
        slide_count += 1
        current_title = title.strip() or f"Slide {slide_count}"
        slide_has_table_or_img = False
        has_cover_logo_only = False

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
            if avail_h_in < 0.50:
                start_new_slide(f"{current_title} (cont.)" if "(cont.)" not in current_title else current_title, flush=False)
                avail_h_in = max_y_in - cur_y_in

            # Group contiguous items of same box type (plain vs code/callout) that fit into avail_h_in
            first = para_buffer[idx]
            is_box = first["kind"] in ("code", "callout")
            chunk: List[dict] = []
            used_h = 0.08

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

            # Keep trailing heading / Listing labels paired with the following code/list/paragraph
            if idx < len(para_buffer) and chunk and chunk[-1].get("keep_with_next"):
                has_prefix = any(not c.get("keep_with_next") for c in chunk)
                if has_prefix or cur_y_in > top_body_in + 0.2:
                    while chunk and chunk[-1].get("keep_with_next"):
                        if not has_prefix and len(chunk) == 0:
                            break
                        chunk.pop()
                        idx -= 1
                    if not chunk:
                        start_new_slide(f"{current_title} (cont.)" if "(cont.)" not in current_title else current_title, flush=False)
                        continue

            box_h_in = min(max(used_h + 0.06, 0.34), avail_h_in)
            tb = current_slide.shapes.add_textbox(
                Inches(margin_x_in), Inches(cur_y_in), Inches(content_w_in), Inches(box_h_in)
            )
            tf = tb.text_frame
            tf.word_wrap = True
            tf.margin_top = Pt(2.5)
            tf.margin_bottom = Pt(2.5)
            tf.margin_left = Pt(6 if is_box else 2)
            tf.margin_right = Pt(6 if is_box else 2)

            if is_box:
                bg_hex = first.get("bg") or ("F1F5F9" if first["kind"] == "code" else "EFF6FF")
                tb.fill.solid()
                tb.fill.fore_color.rgb = RGBColor(int(bg_hex[0:2], 16), int(bg_hex[2:4], 16), int(bg_hex[4:6], 16))

            for p_i, item in enumerate(chunk):
                p = tf.paragraphs[0] if p_i == 0 else tf.add_paragraph()
                p.space_after = Pt(2.5)
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

            cur_y_in += box_h_in + 0.06

        para_buffer.clear()

    def _is_caption_para(p_data: dict) -> bool:
        t = p_data["plain_text"].strip()
        if not t or p_data["images"] or p_data["heading_level"] > 0 or p_data["list_info"] is not None or p_data["is_callout"]:
            return False
        if len(t) > 180:
            return False
        return t.startswith(("Figure ", "Plate ", "Chart ", "Diagram ", "Table ")) or any(s.get("italic") for s in p_data["segments"])

    for b_idx, (b_kind, b_data) in enumerate(body_items):
        if b_kind == 'p':
            info = b_data
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

                    is_cover_logo = (
                        slide_count == 1
                        and not slide_has_table_or_img
                        and not para_buffer
                        and cur_y_in <= top_body_in + 0.05
                        and target_h_in <= 1.25
                        and b_idx <= 2
                    )

                    # Reserve space below the image if a caption follows right after
                    reserve_cap = 0.0
                    if not is_cover_logo and b_idx + 1 < len(body_items) and body_items[b_idx + 1][0] == 'p':
                        if _is_caption_para(body_items[b_idx + 1][1]):
                            reserve_cap = 0.85

                    eff_avail_h = max(avail_h_in - reserve_cap, 1.5) if avail_h_in > 2.35 else avail_h_in
                    max_scale = 1.05 if is_cover_logo else 1.35
                    scale = min(content_w_in / max(target_w_in, 0.1), eff_avail_h / max(target_h_in, 0.1), max_scale)
                    if target_h_in * scale > eff_avail_h:
                        scale = eff_avail_h / max(target_h_in, 0.1)
                    if target_w_in * scale > content_w_in:
                        scale = content_w_in / max(target_w_in, 0.1)

                    box_w = Inches(min(content_w_in, max(target_w_in * scale, 1.0)))
                    box_h = Inches(min(eff_avail_h, max(target_h_in * scale, 0.65)))
                    box_left = Inches(margin_x_in) if is_cover_logo else (Inches(margin_x_in) + int((Inches(content_w_in) - box_w) / 2))
                    _add_aspect_fit_picture(current_slide, flat_bytes, box_left, Inches(cur_y_in), box_w, box_h)
                    cur_y_in += (box_h / Inches(1)) + 0.10
                    if is_cover_logo:
                        has_cover_logo_only = True
                    else:
                        slide_has_table_or_img = True
                except Exception as img_err:
                    logger.warning(f"DOCX->PPTX image render failed for {rid}: {img_err}")

            if text:
                hl = info["heading_level"]
                if hl > 0:
                    _reset_style_list_counters(list_counters)

                next_is_code = (
                    b_idx + 1 < len(body_items)
                    and body_items[b_idx + 1][0] == 'p'
                    and body_items[b_idx + 1][1]["is_mono_block"]
                )

                if hl == 1:
                    if has_cover_logo_only:
                        current_title = text
                        if current_title_shape is not None:
                            current_title_shape.text_frame.paragraphs[0].text = text[:120]
                        has_cover_logo_only = False
                    elif current_slide is not None and (cur_y_in > top_body_in + 0.05 or para_buffer or slide_has_table_or_img):
                        start_new_slide(text)
                    else:
                        if current_slide is None:
                            start_new_slide(text)
                        else:
                            current_title = text
                            if current_title_shape is not None:
                                current_title_shape.text_frame.paragraphs[0].text = text[:120]
                elif hl == 2:
                    has_cover_logo_only = False
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
                            "keep_with_next": True,
                        })
                elif hl >= 3:
                    has_cover_logo_only = False
                    para_buffer.append({
                        "kind": "heading",
                        "text": text,
                        "segments": info["segments"],
                        "pt": 13.0 if hl == 3 else 11.5,
                        "bold": True,
                        "level": 0,
                        "keep_with_next": True,
                    })
                elif info["is_mono_block"]:
                    has_cover_logo_only = False
                    para_buffer.append({
                        "kind": "code",
                        "text": info["plain_text"].strip("\n"),
                        "segments": info["segments"],
                        "pt": 9.5,
                        "bold": False,
                        "bg": "F1F5F9",
                    })
                elif info["is_callout"]:
                    has_cover_logo_only = False
                    para_buffer.append({
                        "kind": "callout",
                        "text": text,
                        "segments": info["segments"],
                        "pt": 11.0,
                        "bold": False,
                        "bg": info["callout_bg"] or "EFF6FF",
                    })
                elif info["list_info"] is not None:
                    has_cover_logo_only = False
                    li = info["list_info"]
                    ilvl = li["ilvl"]
                    if li["ordered"]:
                        num_val = _next_ordered_list_num(list_counters, li["num_id"], ilvl)
                        prefix = f"{num_val}. "
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
                    has_cover_logo_only = False
                    para_buffer.append({
                        "kind": "body",
                        "text": text,
                        "segments": info["segments"],
                        "pt": 11.5,
                        "bold": False,
                        "level": 0,
                        "keep_with_next": next_is_code,
                    })

            if (info["inline_page_breaks"] > 0 or info["has_section_break"]) and current_slide is not None:
                flush_para_buffer()
                current_slide = None

        elif b_kind == 'tbl':
            has_cover_logo_only = False
            # Separate trailing keep_with_next headings so a table heading never gets stranded on the previous slide
            trailing_keep: List[dict] = []
            while para_buffer and para_buffer[-1].get("keep_with_next"):
                trailing_keep.insert(0, para_buffer.pop())

            flush_para_buffer()
            if current_slide is None:
                start_new_slide(current_title, flush=False)

            avail_h_in = max_y_in - cur_y_in
            needed_min_h = 2.0 + (0.45 * len(trailing_keep))
            if avail_h_in < needed_min_h or slide_has_table_or_img:
                start_new_slide(current_title, flush=False)
                avail_h_in = max_y_in - cur_y_in

            if trailing_keep:
                para_buffer.extend(trailing_keep)
                flush_para_buffer()
                avail_h_in = max_y_in - cur_y_in

            tbl = b_data
            raw_rows, span_cmds, repeat_hdr_rows, grid_widths = _extract_docx_table_grid_and_spans(tbl, doc)
            if not raw_rows:
                continue

            if avail_h_in < 1.8 or slide_has_table_or_img:
                start_new_slide(current_title, flush=False)
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
    """Turns each Excel sheet into clean, paginated 16:9 presentation slides with styled tables, merges, links, and flattened images."""
    sheets = _extract_xlsx_workbook_model(xlsx_bytes)
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    slide_count = 0

    margin_x_in = 0.50
    content_w_in = 12.33
    top_body_in = 1.00
    max_y_in = 7.12

    def _new_sheet_slide(title_txt: str):
        nonlocal slide_count
        sl = prs.slides.add_slide(blank_layout)
        slide_count += 1
        t_box = sl.shapes.add_textbox(Inches(margin_x_in), Inches(0.28), Inches(content_w_in), Inches(0.60))
        t_box.text_frame.word_wrap = True
        p = t_box.text_frame.paragraphs[0]
        p.text = title_txt[:115]
        p.font.size = Pt(22 if len(title_txt) < 55 else 18)
        p.font.bold = True
        p.font.color.rgb = RGBColor(15, 23, 42)
        return sl

    try:
        for sh in sheets:
            if sh["is_empty"] and not sh["images"]:
                sl = _new_sheet_slide(sh["title"])
                tx = sl.shapes.add_textbox(Inches(margin_x_in), Inches(1.4), Inches(content_w_in), Inches(1.0))
                p_emp = tx.text_frame.paragraphs[0]
                p_emp.text = "(Empty worksheet — no data)"
                p_emp.font.size = Pt(14)
                p_emp.font.italic = True
                p_emp.font.color.rgb = RGBColor(100, 116, 139)
                continue

            cur_slide = _new_sheet_slide(sh["title"])
            cur_y_in = top_body_in

            if not sh["is_empty"] and sh["rows"]:
                # Adapt sheet grid to _render_styled_pptx_table_slice cell schema
                adapted_rows: List[List[dict]] = []
                for r in sh["rows"]:
                    a_row: List[dict] = []
                    for c in r:
                        a_row.append({
                            "is_origin": c["is_origin"],
                            "fill": c["fill"],
                            "paras": [{
                                "plain_text": c["text"],
                                "segments": [{
                                    "text": c["text"],
                                    "bold": c["bold"],
                                    "italic": c["italic"],
                                    "color": c["color"],
                                    "url": c["url"],
                                }],
                            }] if c["is_origin"] and c["text"] != "" else [],
                        })
                    adapted_rows.append(a_row)

                col_widths_pt, _ = _compute_xlsx_col_widths(sh, content_w_in * 72.0)
                num_cols = len(col_widths_pt)
                rep_cnt = sh["repeat_hdr_rows"]
                hdr_slice = adapted_rows[:rep_cnt] if rep_cnt > 0 else []
                data_rows = adapted_rows[rep_cnt:] if rep_cnt > 0 else adapted_rows
                hdr_spans = [(c0, r0, c1, r1) for (c0, r0, c1, r1) in sh["span_cmds"] if r1 < rep_cnt]

                if not data_rows:
                    data_rows = adapted_rows
                    hdr_slice = []
                    hdr_spans = []
                    rep_cnt = 0

                max_data_per_slide = 12 if num_cols <= 7 else 10
                if len(data_rows) <= max_data_per_slide + 2:
                    avail_h_in = max_y_in - cur_y_in
                    act_h_emu = _render_styled_pptx_table_slice(
                        cur_slide,
                        adapted_rows,
                        sh["span_cmds"],
                        rep_cnt,
                        Inches(margin_x_in),
                        Inches(cur_y_in),
                        Inches(content_w_in),
                        Inches(avail_h_in),
                        col_widths_pt,
                    )
                    cur_y_in += (act_h_emu / Inches(1)) + 0.16
                else:
                    start_idx = 0
                    part_no = 1
                    while start_idx < len(data_rows):
                        if part_no > 1:
                            cur_slide = _new_sheet_slide(f"{sh['title']} (Part {part_no})")
                            cur_y_in = top_body_in
                        end_idx = min(start_idx + max_data_per_slide, len(data_rows))
                        chunk_rows = hdr_slice + data_rows[start_idx:end_idx]
                        chunk_spans = list(hdr_spans)
                        orig_start = rep_cnt + start_idx
                        orig_end = rep_cnt + end_idx - 1
                        for c0, r0, c1, r1 in sh["span_cmds"]:
                            if r0 >= orig_start and r1 <= orig_end:
                                chunk_spans.append((c0, r0 - start_idx, c1, r1 - start_idx))
                        act_h_emu = _render_styled_pptx_table_slice(
                            cur_slide,
                            chunk_rows,
                            chunk_spans,
                            rep_cnt,
                            Inches(margin_x_in),
                            Inches(cur_y_in),
                            Inches(content_w_in),
                            Inches(max_y_in - cur_y_in),
                            col_widths_pt,
                        )
                        cur_y_in += (act_h_emu / Inches(1)) + 0.16
                        start_idx = end_idx
                        part_no += 1

            # Place sheet images cleanly below table if generous space remains, else on a dedicated visual slide
            if sh["images"]:
                imgs = sh["images"]
                rem_h_in = max_y_in - cur_y_in
                if rem_h_in >= 2.15:
                    slot_w_in = (content_w_in - 0.25 * (len(imgs) - 1)) / len(imgs)
                    for idx_im, im_info in enumerate(imgs):
                        bx_left = Inches(margin_x_in + idx_im * (slot_w_in + 0.25))
                        _add_aspect_fit_picture(cur_slide, im_info["bytes"], bx_left, Inches(cur_y_in), Inches(slot_w_in), Inches(rem_h_in))
                else:
                    for idx_im, im_info in enumerate(imgs):
                        vis_slide = _new_sheet_slide(f"{sh['title']} — Visual {idx_im + 1}" if len(imgs) > 1 else f"{sh['title']} — Visual")
                        _add_aspect_fit_picture(vis_slide, im_info["bytes"], Inches(margin_x_in), Inches(top_body_in), Inches(content_w_in), Inches(max_y_in - top_body_in))

        out_io = io.BytesIO()
        prs.save(out_io)
        return out_io.getvalue(), "application/vnd.openxmlformats-officedocument.presentationml.presentation", f"{base_name}.pptx", slide_count
    finally:
        sheets.clear()
        gc.collect()


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


def _render_xlsx_sheet_csv_bytes(sheet: Dict[str, Any]) -> bytes:
    """Renders a normalized XLSX sheet model to RFC-4180 UTF-8-SIG CSV bytes."""
    sio = io.StringIO()
    writer = csv.writer(sio, dialect="excel", quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
    for row in sheet["rows"]:
        vals = [c["text"] if c["is_origin"] else "" for c in row]
        if any(v != "" for v in vals):
            writer.writerow(vals)
    return sio.getvalue().encode("utf-8-sig")


def xlsx_to_csv(xlsx_bytes: bytes, base_name: str) -> Tuple[bytes, str, str, int]:
    """Exports Excel workbook to RFC-4180 UTF-8-SIG CSV (or ZIP of CSVs if multi-sheet, skipping blank sheets)."""
    sheets = _extract_xlsx_workbook_model(xlsx_bytes)
    try:
        sheet_count = len(sheets)
        if sheet_count <= 1:
            if not sheets or sheets[0]["is_empty"]:
                return b"\xef\xbb\xbf", "text/csv; charset=utf-8", f"{base_name}.csv", 1
            return _render_xlsx_sheet_csv_bytes(sheets[0]), "text/csv; charset=utf-8", f"{base_name}.csv", 1

        non_empty_sheets = [sh for sh in sheets if not sh["is_empty"] and sh["rows"]]
        if not non_empty_sheets:
            return b"\xef\xbb\xbf", "text/csv; charset=utf-8", f"{base_name}.csv", 1

        zip_buf = io.BytesIO()
        used_names: Set[str] = set()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for sh in non_empty_sheets:
                clean_title = re.sub(r'[\\/*?:"<>|]', "_", sh["title"]).strip() or "sheet"
                entry_name = get_unique_zip_name(f"{clean_title}.csv", used_names)
                zf.writestr(entry_name, _render_xlsx_sheet_csv_bytes(sh))
        return zip_buf.getvalue(), "application/zip", f"{base_name}_csv_sheets.zip", len(non_empty_sheets)
    finally:
        sheets.clear()
        gc.collect()


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
                    _reset_style_list_counters(list_counters)
                    parts.append(f"\n{txt_line}")
                elif info["is_mono_block"]:
                    parts.append("\n" + info["plain_text"].strip("\n") + "\n")
                elif info["list_info"] is not None:
                    li = info["list_info"]
                    ilvl = li["ilvl"]
                    indent = "  " * ilvl
                    if li["ordered"]:
                        num_val = _next_ordered_list_num(list_counters, li["num_id"], ilvl)
                        parts.append(f"{indent}{num_val}. {txt_line}")
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


def _pptx_to_txt(pptx_bytes: bytes) -> str:
    """Extracts reading-order plain text from PPTX including GroupShapes, tables (TSV), hyperlinks, lists, code, and speaker notes."""
    model = _extract_pptx_deck_model(pptx_bytes)
    parts: List[str] = []

    for s_data in model["slides"]:
        s_idx = s_data["index"]
        parts.append(f"--- Slide {s_idx}: {s_data['title']} ---")

        if s_data["title_item_idx"] is not None:
            t_item = s_data["items"][s_data["title_item_idx"]]
            t_paras = [p for p in t_item.get("paras", []) if p["plain_text"].strip()]
            for p_extra in t_paras[1:]:
                t_line = _segments_to_txt(p_extra["segments"]).strip()
                if t_line:
                    parts.append(t_line)

        if s_data["is_blank"]:
            parts.append("(Blank slide)")
        elif s_data["is_image_only"]:
            parts.append("(Full-bleed visual slide)")

        body_items = [
            it for idx, it in enumerate(s_data["items"])
            if idx != s_data["title_item_idx"]
        ]
        body_items.sort(key=lambda it: (round(it["y"] / 16.0), it["x"]))

        for item in body_items:
            kind = item["kind"]
            if kind == "table":
                parts.append("")
                for row in item["raw_rows"]:
                    row_vals: List[str] = []
                    for cell in row:
                        if not cell.get("is_origin"):
                            row_vals.append("")
                        else:
                            c_txt = " ".join(
                                _segments_to_txt(p["segments"]).strip()
                                for p in cell.get("paras", [])
                                if p["plain_text"].strip()
                            ).strip()
                            row_vals.append(re.sub(r"\s*\n\s*", " ", c_txt))
                    while row_vals and row_vals[-1] == "":
                        row_vals.pop()
                    if row_vals:
                        parts.append("\t".join(row_vals))
                parts.append("")
            elif kind == "shape":
                non_empty_paras = [p for p in item.get("paras", []) if p["plain_text"].strip()]
                if not non_empty_paras:
                    continue
                if item.get("is_mono_block"):
                    parts.append("")
                    parts.extend(p["plain_text"] for p in non_empty_paras)
                    parts.append("")
                    continue
                for p_info in non_empty_paras:
                    txt_line = _segments_to_txt(p_info["segments"]).strip()
                    if not txt_line:
                        continue
                    li = p_info.get("list_info")
                    lvl = p_info.get("level", 0)
                    if li is not None:
                        prefix = li["prefix"] if li["ordered"] else "- "
                        parts.append(f"{'  ' * lvl}{prefix}{txt_line}")
                    else:
                        parts.append(txt_line)

        if s_data.get("speaker_notes"):
            parts.append(f"[Speaker Notes]: {s_data['speaker_notes']}")
        parts.append("")

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
            sheets = _extract_xlsx_workbook_model(data_bytes)
            try:
                parts: List[str] = []
                for sh in sheets:
                    parts.append(f"=== Sheet: {sh['title']} ===")
                    if sh["is_empty"] or not sh["rows"]:
                        parts.append("(Empty sheet)")
                        continue
                    for r in sh["rows"]:
                        row_vals: List[str] = []
                        for c in r:
                            if not c["is_origin"]:
                                row_vals.append("")
                            else:
                                cell_txt = re.sub(r"[\t\r\n]+", " ", c["text"]).strip()
                                if c["url"] and c["url"] not in cell_txt:
                                    cell_txt = f"{cell_txt} ({c['url']})".strip() if cell_txt else c["url"]
                                row_vals.append(cell_txt)
                        while row_vals and row_vals[-1] == "":
                            row_vals.pop()
                        if row_vals:
                            parts.append("\t".join(row_vals))
                text = "\n".join(parts)
            finally:
                sheets.clear()
                gc.collect()

    elif input_fmt == "pptx":
        text = _pptx_to_txt(data_bytes)

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
                    _reset_style_list_counters(list_counters)
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
                        num_val = _next_ordered_list_num(list_counters, li["num_id"], ilvl)
                        parts.append(f"{indent}{num_val}. {item_md}")
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


def _pptx_to_md(pptx_bytes: bytes) -> str:
    """Converts PowerPoint slides directly into structured GFM Markdown with GroupShapes, multi-level lists, fenced code, tables, links, and speaker notes."""
    model = _extract_pptx_deck_model(pptx_bytes)
    parts: List[str] = []

    for s_idx, s_data in enumerate(model["slides"], start=1):
        if s_idx > 1:
            parts.append("\n---\n")
        parts.append(f"# {s_data['title']}\n")

        if s_data["title_item_idx"] is not None:
            t_item = s_data["items"][s_data["title_item_idx"]]
            t_paras = [p for p in t_item.get("paras", []) if p["plain_text"].strip()]
            for p_extra in t_paras[1:]:
                p_md = _segments_to_md(p_extra["segments"], include_inline_styles=True).strip()
                if p_md:
                    parts.append(f"{p_md}\n")

        if s_data["is_blank"]:
            parts.append("*(Blank slide)*\n")
        elif s_data["is_image_only"]:
            parts.append("*(Full-bleed visual slide)*\n")

        body_items = [
            it for idx, it in enumerate(s_data["items"])
            if idx != s_data["title_item_idx"]
        ]
        body_items.sort(key=lambda it: (round(it["y"] / 16.0), it["x"]))

        in_list = False

        def close_list():
            nonlocal in_list
            if in_list:
                parts.append("")
                in_list = False

        for item in body_items:
            kind = item["kind"]
            if kind == "table":
                close_list()
                rows_grid: List[List[str]] = []
                for row in item["raw_rows"]:
                    row_strs: List[str] = []
                    for cell in row:
                        if not cell.get("is_origin"):
                            row_strs.append("")
                        else:
                            cell_md = " ".join(
                                _segments_to_md(p["segments"], include_inline_styles=False).strip()
                                for p in cell.get("paras", [])
                                if p["plain_text"].strip()
                            ).strip()
                            row_strs.append(cell_md)
                    rows_grid.append(row_strs)
                if rows_grid:
                    parts.append(_format_md_table(rows_grid) + "\n")

            elif kind == "shape":
                non_empty_paras = [p for p in item.get("paras", []) if p["plain_text"].strip()]
                if not non_empty_paras:
                    continue

                if item.get("is_mono_block"):
                    close_list()
                    code_txt = "\n".join(p["plain_text"] for p in non_empty_paras).strip("\n")
                    parts.append(f"```\n{code_txt}\n```\n")
                    continue

                fill_hex = item.get("fill_hex")
                is_card = bool(fill_hex and fill_hex not in ("FFFFFF", "AUTO"))
                if is_card and (item.get("in_group") or len(non_empty_paras) <= 4):
                    close_list()
                    q_lines = []
                    for p_info in non_empty_paras:
                        pm = _segments_to_md(p_info["segments"], include_inline_styles=True).strip()
                        if pm:
                            q_lines.append(f"> {pm}")
                    if q_lines:
                        parts.append("\n".join(q_lines) + "\n")
                    continue

                for p_info in non_empty_paras:
                    li = p_info.get("list_info")
                    lvl = p_info.get("level", 0)
                    p_md = _segments_to_md(p_info["segments"], include_inline_styles=True).strip()
                    if not p_md:
                        continue
                    if li is not None:
                        indent = "  " * lvl
                        prefix = li["prefix"] if li["ordered"] else "- "
                        parts.append(f"{indent}{prefix}{p_md}")
                        in_list = True
                    else:
                        close_list()
                        parts.append(f"{p_md}\n")

        close_list()
        if s_data.get("speaker_notes"):
            parts.append(f"> **Speaker Notes:** *{s_data['speaker_notes']}*\n")

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
            sheets = _extract_xlsx_workbook_model(xlsx_data)
        else:
            sheets = _extract_xlsx_workbook_model(data_bytes)

        try:
            parts: List[str] = []
            for sh in sheets:
                parts.append(f"## Sheet: {sh['title']}\n")
                if sh["is_empty"] or not sh["rows"]:
                    parts.append("*(Empty worksheet)*\n")
                    continue

                b_rows = sh["banner_rows"] if sh["banner_rows"] < len(sh["rows"]) else 0
                for b_idx in range(b_rows):
                    b_txt = " ".join(c["text"].strip() for c in sh["rows"][b_idx] if c["is_origin"] and c["text"].strip()).strip()
                    if b_txt:
                        parts.append(f"### {b_txt}\n" if b_idx == 0 else f"> {b_txt}\n")

                tbl_rows = sh["rows"][b_rows:]
                if tbl_rows:
                    rows_grid: List[List[str]] = []
                    for r in tbl_rows:
                        r_strs: List[str] = []
                        for c in r:
                            if not c["is_origin"]:
                                r_strs.append("")
                            elif c["url"]:
                                lbl = re.sub(r"\s*\n\s*", " ", c["text"]).strip() or c["url"]
                                lbl = lbl.replace("[", "\\[").replace("]", "\\]")
                                r_strs.append(f"[{lbl}]({c['url']})")
                            else:
                                r_strs.append(c["text"])
                        rows_grid.append(r_strs)
                    parts.append(_format_md_table(rows_grid) + "\n")
            md_text = "\n".join(parts)
        finally:
            sheets.clear()
            gc.collect()

    elif input_fmt == "pdf":
        return pdf_to_md(data_bytes, base_name)

    elif input_fmt == "pptx":
        md_text = _pptx_to_md(data_bytes)

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
    list_stack: List[List[Any]] = []

    def close_lists(target_level: int = -1, next_tag: Optional[str] = None):
        while list_stack:
            top_tag, top_lvl, open_li = list_stack[-1]
            if top_lvl > target_level or (top_lvl == target_level and next_tag is not None and top_tag != next_tag):
                if open_li:
                    parts.append("</li>")
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
                        list_stack.append([tag, ilvl, False])
                    elif list_stack[-1][2]:
                        parts.append("</li>")
                        list_stack[-1][2] = False
                    item_html = _segments_to_html(info["segments"], include_inline_styles=True).strip()
                    parts.append(f"<li>{item_html}")
                    list_stack[-1][2] = True
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


def _xlsx_to_html_body(xlsx_bytes: bytes) -> str:
    """Converts an XLSX workbook into responsive HTML5 with sheet navigation, colspan/rowspan, shading, links, and base64 images."""
    sheets = _extract_xlsx_workbook_model(xlsx_bytes)
    parts: List[str] = []
    try:
        if sheets:
            nav_links = [
                f'<a class="sheet-pill" href="#sheet-{idx}">{html.escape(sh["title"])}</a>'
                for idx, sh in enumerate(sheets)
            ]
            parts.append(f'<nav class="sheet-nav" aria-label="Worksheet navigation">{"".join(nav_links)}</nav>')

        for s_idx, sh in enumerate(sheets):
            parts.append(f'<section id="sheet-{s_idx}" class="xlsx-sheet">')
            parts.append(f'<h2>{html.escape(sh["title"])}</h2>')

            if sh["is_empty"] and not sh["images"]:
                parts.append('<p class="sheet-empty">(Empty worksheet — no data)</p>')
                parts.append("</section>")
                continue

            if not sh["is_empty"] and sh["rows"]:
                span_map = {(r0, c0): (c1 - c0 + 1, r1 - r0 + 1) for c0, r0, c1, r1 in sh["span_cmds"]}
                tbl_html: List[str] = ['<div class="table-wrap"><table>']
                for r_idx, row in enumerate(sh["rows"]):
                    is_hdr = r_idx < sh["repeat_hdr_rows"]
                    tbl_html.append("<tr>")
                    for c_idx, cell in enumerate(row):
                        if not cell["is_origin"]:
                            continue
                        cs, rs = span_map.get((r_idx, c_idx), (1, 1))
                        attrs: List[str] = []
                        if cs > 1:
                            attrs.append(f'colspan="{cs}"')
                        if rs > 1:
                            attrs.append(f'rowspan="{rs}"')

                        fill_hex = cell["fill"] or ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
                        dark_bg = _is_dark_hex(fill_hex)
                        fg_hex = cell["color"] or ("FFFFFF" if dark_bg else "0F172A")
                        style_parts = [
                            f"background-color:#{html.escape(fill_hex)}",
                            f"color:#{html.escape(fg_hex)}",
                            f"text-align:{html.escape(cell['align'])}",
                        ]
                        if cell["bold"] or is_hdr:
                            style_parts.append("font-weight:700")
                        if cell["italic"]:
                            style_parts.append("font-style:italic")
                        attrs.append(f'style="{"; ".join(style_parts)};"')

                        esc_txt = html.escape(cell["text"]).replace("\n", "<br>")
                        if cell["url"]:
                            esc_url = html.escape(cell["url"], quote=True)
                            link_col = "93C5FD" if dark_bg else "2563EB"
                            inner_html = f'<a href="{esc_url}" target="_blank" rel="noopener" style="color:#{link_col};">{esc_txt}</a>'
                        else:
                            inner_html = esc_txt

                        tag = "th" if is_hdr else "td"
                        tbl_html.append(f'<{tag} {" ".join(attrs)}>{inner_html}</{tag}>')
                    tbl_html.append("</tr>")
                tbl_html.append("</table></div>")
                parts.append("".join(tbl_html))

            if sh["images"]:
                for im_info in sh["images"]:
                    b64 = base64.b64encode(im_info["bytes"]).decode("ascii")
                    max_px = max(int(im_info["disp_w_pt"] * 1.333), 240)
                    parts.append(
                        f'<figure class="doc-figure">'
                        f'<img src="data:image/png;base64,{b64}" alt="Sheet diagram" style="max-width: min(100%, {max_px}px); height: auto;">'
                        f'</figure>'
                    )
                    del b64

            parts.append("</section>")

        return "\n".join(parts)
    finally:
        sheets.clear()
        gc.collect()


def _pptx_to_html_body(pptx_bytes: bytes) -> str:
    """Converts a PPTX presentation into rich responsive HTML5 slides with nav pills, GroupShape cards, colspan/rowspan tables, base64 images, code blocks, links, and speaker notes."""
    model = _extract_pptx_deck_model(pptx_bytes)
    slides = model["slides"]
    parts: List[str] = []

    if slides:
        nav_links = [
            f'<a class="sheet-pill" href="#slide-{s["index"]}">Slide {s["index"]}</a>'
            for s in slides
        ]
        parts.append(f'<nav class="sheet-nav" aria-label="Slide navigation">{"".join(nav_links)}</nav>')

    for s_data in slides:
        s_idx = s_data["index"]
        is_dark = s_data["is_dark"]
        bg_hex = s_data["bg_hex"]
        slide_cls = "pptx-slide is-dark" if is_dark else "pptx-slide"
        slide_style = f' style="background-color:#{html.escape(bg_hex)}; color:#FFFFFF;"' if is_dark else ""
        parts.append(f'<section id="slide-{s_idx}" class="{slide_cls}"{slide_style}>')
        parts.append(f'<div class="slide-kicker">Slide {s_idx} of {len(slides)}</div>')
        parts.append(f'<h2>{html.escape(s_data["title"])}</h2>')

        if s_data["title_item_idx"] is not None:
            t_item = s_data["items"][s_data["title_item_idx"]]
            t_paras = [p for p in t_item.get("paras", []) if p["plain_text"].strip()]
            for p_extra in t_paras[1:]:
                segs = []
                for s in p_extra["segments"]:
                    sc = dict(s)
                    if is_dark and sc.get("color") in (None, "000000", "0F172A", "1E293B"):
                        sc["color"] = "E2E8F0"
                    segs.append(sc)
                p_html = _segments_to_html(segs, include_inline_styles=True).strip()
                if p_html:
                    parts.append(f'<p class="slide-subtitle">{p_html}</p>')

        if s_data["is_blank"]:
            parts.append('<p class="sheet-empty">(Blank slide)</p>')

        body_items = [
            it for idx, it in enumerate(s_data["items"])
            if idx != s_data["title_item_idx"]
        ]
        body_items.sort(key=lambda it: (round(it["y"] / 16.0), it["x"]))

        list_stack: List[List[Any]] = []

        def close_lists(target_level: int = -1, next_tag: Optional[str] = None):
            while list_stack:
                top_tag, top_lvl, open_li = list_stack[-1]
                if top_lvl > target_level or (top_lvl == target_level and next_tag is not None and top_tag != next_tag):
                    if open_li:
                        parts.append("</li>")
                    parts.append(f"</{top_tag}>")
                    list_stack.pop()
                else:
                    break

        for item in body_items:
            kind = item["kind"]
            if kind == "picture":
                close_lists()
                try:
                    raw_blob = item["blob"]
                    fmt = item["fmt"]
                    has_alpha = item["has_alpha"]
                    if fmt == "JPEG" and not has_alpha:
                        mime = "image/jpeg"
                        img_out = raw_blob
                    elif fmt == "PNG":
                        # Preserve clean RGBA alpha in HTML5 so logos/diagrams composite cleanly onto dark or light slide cards
                        mime = "image/png"
                        img_out = raw_blob
                    else:
                        with Image.open(io.BytesIO(raw_blob)) as im:
                            conv = im.convert("RGBA") if has_alpha else im.convert("RGB")
                            ibuf = io.BytesIO()
                            conv.save(ibuf, format="PNG")
                            img_out = ibuf.getvalue()
                        mime = "image/png"
                    b64 = base64.b64encode(img_out).decode("ascii")
                    max_px = max(int(item["w"] * 1.333), 200)
                    parts.append(
                        f'<figure class="doc-figure align-center">'
                        f'<img src="data:{mime};base64,{b64}" alt="Slide {s_idx} visual" style="max-width: min(100%, {max_px}px); height: auto;">'
                        f'</figure>'
                    )
                    del b64
                except Exception as img_err:
                    logger.warning(f"PPTX->HTML image embed failed on slide {s_idx}: {img_err}")

            elif kind == "table":
                close_lists()
                raw_rows = item["raw_rows"]
                if not raw_rows:
                    continue
                span_map = {(r0, c0): (c1 - c0 + 1, r1 - r0 + 1) for c0, r0, c1, r1 in item.get("span_cmds", [])}
                tbl_html = ['<div class="table-wrap"><table>']
                for r_idx, row in enumerate(raw_rows):
                    is_hdr = r_idx == 0
                    tbl_html.append("<tr>")
                    for c_idx, cell in enumerate(row):
                        if not cell.get("is_origin"):
                            continue
                        cs, rs = span_map.get((r_idx, c_idx), (1, 1))
                        attrs = []
                        if cs > 1:
                            attrs.append(f'colspan="{cs}"')
                        if rs > 1:
                            attrs.append(f'rowspan="{rs}"')

                        fill_hex = cell.get("fill") or ("0F172A" if is_hdr else ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF"))
                        dark_c = _is_dark_hex(fill_hex)
                        fg_hex = "FFFFFF" if dark_c else "0F172A"
                        attrs.append(f'style="background-color:#{html.escape(fill_hex)}; color:#{fg_hex};"')

                        inner_parts = []
                        for p_info in cell.get("paras", []):
                            if p_info["plain_text"].strip():
                                segs = []
                                for s in p_info["segments"]:
                                    sc = dict(s)
                                    if sc.get("url"):
                                        sc["color"] = "93C5FD" if dark_c else "2563EB"
                                    elif dark_c and sc.get("color") in (None, "0F172A", "1E293B", "334155", "000000"):
                                        sc["color"] = "FFFFFF"
                                    segs.append(sc)
                                inner_parts.append(_segments_to_html(segs, include_inline_styles=True).strip())
                        cell_content = "<br>".join(inner_parts)
                        tag = "th" if is_hdr else "td"
                        tbl_html.append(f'<{tag} {" ".join(attrs)}>{cell_content}</{tag}>')
                    tbl_html.append("</tr>")
                tbl_html.append("</table></div>")
                parts.append("".join(tbl_html))

            elif kind == "shape":
                non_empty_paras = [p for p in item.get("paras", []) if p["plain_text"].strip()]
                if not non_empty_paras:
                    continue

                if item.get("is_mono_block"):
                    close_lists()
                    code_esc = html.escape("\n".join(p["plain_text"] for p in non_empty_paras).strip("\n"))
                    parts.append(f"<pre><code>{code_esc}</code></pre>")
                    continue

                fill_hex = item.get("fill_hex")
                is_card = bool(fill_hex and fill_hex not in ("FFFFFF", "AUTO"))
                card_dark = (is_card and _is_dark_hex(fill_hex)) or (not is_card and is_dark)

                if is_card and (item.get("in_group") or len(non_empty_paras) <= 4):
                    close_lists()
                    fg = "FFFFFF" if card_dark else "0F172A"
                    bd = item.get("line_hex") or ("38BDF8" if card_dark else "2563EB")
                    card_lines = []
                    for p_info in non_empty_paras:
                        segs = []
                        for s in p_info["segments"]:
                            sc = dict(s)
                            if card_dark and sc.get("color") in (None, "000000", "0F172A", "1E293B", "334155"):
                                sc["color"] = "FFFFFF"
                            segs.append(sc)
                        card_lines.append(f'<div>{_segments_to_html(segs, include_inline_styles=True).strip()}</div>')
                    parts.append(
                        f'<div class="pptx-card" style="background-color:#{html.escape(fill_hex)}; color:#{fg}; border-left:4px solid #{html.escape(bd)};">'
                        + "".join(card_lines)
                        + "</div>"
                    )
                    continue

                for p_info in non_empty_paras:
                    li = p_info.get("list_info")
                    lvl = p_info.get("level", 0)
                    segs = []
                    for s in p_info["segments"]:
                        sc = dict(s)
                        if card_dark and sc.get("color") in (None, "000000", "0F172A", "1E293B", "334155"):
                            sc["color"] = "FFFFFF"
                        segs.append(sc)
                    p_html = _segments_to_html(segs, include_inline_styles=True).strip()
                    if not p_html:
                        continue
                    if li is not None:
                        tag = "ol" if li["ordered"] else "ul"
                        close_lists(lvl, tag)
                        if not list_stack or list_stack[-1][1] < lvl:
                            parts.append(f"<{tag}>")
                            list_stack.append([tag, lvl, False])
                        elif list_stack[-1][2]:
                            parts.append("</li>")
                            list_stack[-1][2] = False
                        parts.append(f"<li>{p_html}")
                        list_stack[-1][2] = True
                    else:
                        close_lists()
                        parts.append(f"<p>{p_html}</p>")

        close_lists()
        if s_data.get("speaker_notes"):
            parts.append(
                f'<aside class="pptx-notes"><strong>Speaker Notes:</strong> {html.escape(s_data["speaker_notes"])}</aside>'
            )
        parts.append("</section>")

    gc.collect()
    return "\n".join(parts)


def any_to_html(data_bytes: bytes, input_fmt: str, base_name: str) -> Tuple[bytes, str, str, int]:
    """Generates a clean responsive HTML5 document from any file format."""
    norm_fmt = input_fmt.lower().replace("jpeg", "jpg")
    if norm_fmt == "docx":
        body_html = _docx_to_html_body(data_bytes)
    elif norm_fmt == "xlsx":
        body_html = _xlsx_to_html_body(data_bytes)
    elif norm_fmt == "pptx":
        body_html = _pptx_to_html_body(data_bytes)
    else:
        md_bytes, _, _, _ = any_to_md(data_bytes, input_fmt, base_name)
        body_html = markdown.markdown(md_bytes.decode('utf-8'), extensions=['tables', 'fenced_code'])

    full_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>{html.escape(base_name)}</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "DejaVu Sans", sans-serif; line-height: 1.6; max-width: 1120px; margin: 36px auto; padding: 0 24px; color: #1e293b; background: #ffffff; }}
    h1, h2, h3, h4 {{ color: #0f172a; margin-top: 1.4em; }}
    h1, h2 {{ border-bottom: 1px solid #e2e8f0; padding-bottom: 6px; }}
    .sheet-nav {{ position: sticky; top: 8px; z-index: 20; display: flex; flex-wrap: wrap; gap: 8px; padding: 10px 14px; margin-bottom: 22px; background: #f8fafc; border: 1px solid #cbd5e1; border-radius: 10px; }}
    .sheet-pill {{ display: inline-block; padding: 4px 12px; border-radius: 9999px; background: #ffffff; border: 1px solid #cbd5e1; color: #0f172a; font-size: 0.85rem; font-weight: 600; text-decoration: none; }}
    .sheet-pill:hover {{ background: #eff6ff; border-color: #2563eb; color: #2563eb; }}
    .xlsx-sheet {{ margin-bottom: 34px; }}
    .pptx-slide {{ margin-bottom: 28px; padding: 26px 28px; border: 1px solid #cbd5e1; border-radius: 12px; background: #ffffff; box-shadow: 0 2px 8px rgba(15,23,42,0.05); }}
    .pptx-slide.is-dark h2, .pptx-slide.is-dark p {{ color: #ffffff; border-color: rgba(255,255,255,0.2); }}
    .pptx-slide.is-dark a {{ color: #7dd3fc; }}
    .slide-kicker {{ font-size: 0.78rem; font-weight: 700; letter-spacing: 0.06em; text-transform: uppercase; opacity: 0.72; margin-bottom: 4px; }}
    .slide-subtitle {{ opacity: 0.9; margin-top: 4px; }}
    .pptx-card {{ margin: 12px 0; padding: 14px 18px; border-radius: 8px; box-shadow: 0 1px 4px rgba(15,23,42,0.08); }}
    .pptx-notes {{ margin-top: 16px; padding: 10px 14px; border-radius: 6px; background: #fefce8; border-left: 4px solid #d97706; color: #334155; font-size: 0.9rem; }}
    .sheet-empty {{ color: #64748b; font-style: italic; }}
    .doc-header, .doc-footer {{ font-size: 0.82rem; color: #64748b; padding: 8px 0; border-bottom: 1px solid #e2e8f0; margin-bottom: 18px; }}
    .doc-footer {{ border-bottom: none; border-top: 1px solid #e2e8f0; margin-top: 28px; }}
    .doc-figure {{ margin: 18px 0; }}
    .doc-figure.align-center {{ text-align: center; }}
    .doc-figure.align-right {{ text-align: right; }}
    .doc-figure img {{ border-radius: 6px; background: #ffffff; }}
    .pptx-slide .doc-figure img {{ background: transparent; }}
    .doc-callout {{ margin: 14px 0; padding: 12px 16px; border-radius: 6px; color: #1e293b; }}
    .page-break {{ border: none; border-top: 1px dashed #cbd5e1; margin: 28px 0; }}
    .table-wrap {{ width: 100%; overflow-x: auto; margin: 18px 0; }}
    table {{ border-collapse: collapse; width: 100%; font-size: 0.88rem; }}
    th, td {{ border: 1px solid #cbd5e1; padding: 6px 10px; text-align: left; vertical-align: middle; }}
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

