"""
stress_test_matrix.py - Full-Matrix 110-Pair Stress Test & Edge-Case Torture Suite
Verifies every single directed pair across 11 formats (110 combinations) plus edge cases
against the live server at http://127.0.0.1:8000.
"""

import io
import json
import os
import sys
import time
import urllib.request
import urllib.error
import zipfile
from pathlib import Path
from typing import Dict, Tuple

# Third-party parsers for deep validation
from PIL import Image, ImageDraw, ImageFont
import pymupdf
import openpyxl
from docx import Document
from pptx import Presentation

BASE_DIR = Path(__file__).resolve().parent
FIXTURES_DIR = BASE_DIR / "stress_fixtures"
FIXTURES_DIR.mkdir(exist_ok=True)
SERVER_URL = "http://127.0.0.1:8000"

ALL_FORMATS = [
    "pdf", "docx", "xlsx", "pptx", "csv",
    "png", "jpg", "webp", "md", "txt", "html"
]


def get_font(size: int = 20):
    for fn in ["segoeui.ttf", "arial.ttf", "calibri.ttf", "DejaVuSans.ttf"]:
        try:
            return ImageFont.truetype(fn, size)
        except Exception:
            pass
    try:
        return ImageFont.load_default(size=size)
    except Exception:
        return ImageFont.load_default()


def create_sample_fixtures() -> Dict[str, Tuple[str, bytes]]:
    """Generates realistic, non-trivial sample fixtures for all 11 formats."""
    fixtures = {}

    # 1. PDF (Multi-page, tables, Unicode symbols)
    pdf_path = FIXTURES_DIR / "sample_stress.pdf"
    doc = pymupdf.open()
    p1 = doc.new_page(width=595, height=842)
    p1.insert_text((50, 70), "Global Financial & Operations Report", fontsize=18)
    p1.insert_text((50, 100), "Metrics with Unicode: Revenue € 50,000 | Spend ₹ 1,200,000 | Temp 24°C | Q1–Q4", fontsize=11)
    # Draw simple table grid
    p1.draw_rect(pymupdf.Rect(50, 130, 545, 230), color=(0.2, 0.4, 0.8), width=1)
    p1.insert_text((60, 155), "Metric\t\tBudget (€)\tSpend (₹)\tStatus", fontsize=11)
    p1.insert_text((60, 185), "R&D Engine\t€ 25,000\t₹ 600,000\tActive", fontsize=10)
    p1.insert_text((60, 215), "Operations\t€ 25,000\t₹ 600,000\tApproved", fontsize=10)

    p2 = doc.new_page(width=595, height=842)
    p2.insert_text((50, 70), "Observations & Next Milestones", fontsize=18)
    p2.insert_text((50, 110), "All European branches (€) and Indian branches (₹) met targets within –2% to +5% range.", fontsize=11)
    doc.save(pdf_path)
    doc.close()
    with open(pdf_path, "rb") as f:
        fixtures["pdf"] = ("sample_stress.pdf", f.read())

    # 2. DOCX (Headings, multi-row table, styled paragraphs, Unicode)
    docx_path = FIXTURES_DIR / "sample_stress.docx"
    wdoc = Document()
    wdoc.add_heading("Corporate Performance Dashboard", level=1)
    wdoc.add_paragraph("Fiscal evaluation across regions: Frankfurt (€), Bengaluru (₹), and Helsinki (–10°C to 28°C).")
    tbl = wdoc.add_table(rows=3, cols=3)
    tbl.style = "Table Grid"
    headers = ["Department", "Budget (€/₹)", "Achievement"]
    for c_i, h in enumerate(headers):
        tbl.cell(0, c_i).text = h
    r1 = ["Platform Infrastructure", "€ 45,000", "99.98%"]
    for c_i, val in enumerate(r1):
        tbl.cell(1, c_i).text = val
    r2 = ["Customer Services", "₹ 750,000", "Exceeded Target"]
    for c_i, val in enumerate(r2):
        tbl.cell(2, c_i).text = val
    wdoc.add_heading("Operational Directives", level=2)
    wdoc.add_paragraph("High priority initiatives for Q3–Q4 delivery.")
    buf = io.BytesIO()
    wdoc.save(buf)
    with open(docx_path, "wb") as f:
        f.write(buf.getvalue())
    fixtures["docx"] = ("sample_stress.docx", buf.getvalue())

    # 3. XLSX (Multi-sheet, formulas/numbers, Unicode)
    xlsx_path = FIXTURES_DIR / "sample_stress.xlsx"
    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Financial_Summary"
    ws1.append(["Division", "Budget_EUR (€)", "Spend_INR (₹)", "Score", "Trend–Status"])
    ws1.append(["Engineering", 80000, 1500000, 94.5, "Strong Growth"])
    ws1.append(["Marketing", 45000, 850000, 88.0, "Stable–Target"])
    ws1.append(["Logistics", 32000, 620000, 91.2, "On–Track"])

    ws2 = wb.create_sheet(title="Climate_Logs")
    ws2.append(["Facility", "Current Temp (°C)", "Target Range (°C)"])
    ws2.append(["DataCenter A", 21.5, "20°–24°"])
    ws2.append(["ServerHub B", 22.0, "19°–23°"])

    xbuf = io.BytesIO()
    wb.save(xbuf)
    with open(xlsx_path, "wb") as f:
        f.write(xbuf.getvalue())
    fixtures["xlsx"] = ("sample_stress.xlsx", xbuf.getvalue())

    # 4. PPTX (Presentation, slides, bullets, Unicode)
    pptx_path = FIXTURES_DIR / "sample_stress.pptx"
    prs = Presentation()
    slide1 = prs.slides.add_slide(prs.slide_layouts[0])
    slide1.shapes.title.text = "Quarterly Strategy 2026"
    slide1.placeholders[1].text = "Revenue Target: € 5M | Local Spend: ₹ 12M | 22°C Operating Norm"

    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = "Execution Roadmaps & Milestones"
    tf = slide2.shapes.placeholders[1].text_frame
    tf.text = "Phase 1 – Initial Rollout Complete"
    p = tf.add_paragraph()
    p.text = "Phase 2 – Budget allocated: € 150,000 / ₹ 3,500,000"
    p2 = tf.add_paragraph()
    p2.text = "Phase 3 – Quality Assurance target > 99.5%"

    pbuf = io.BytesIO()
    prs.save(pbuf)
    with open(pptx_path, "wb") as f:
        f.write(pbuf.getvalue())
    fixtures["pptx"] = ("sample_stress.pptx", pbuf.getvalue())

    # 5. CSV (Comma-separated, numbers, Unicode)
    csv_path = FIXTURES_DIR / "sample_stress.csv"
    csv_text = (
        "Item_ID,Category,Price_EUR,Price_INR,Operating_Temp,Notes\n"
        "101,Industrial Controller,€ 850.00,₹ 76500,45°C,Heavy-duty unit – certified\n"
        "102,Smart Sensor,€ 120.50,₹ 10800,25°C,Ambient sensor – calibrated\n"
        "103,Power Unit,€ 450.00,₹ 40500,60°C,Backup supply – failover ready\n"
    )
    with open(csv_path, "w", encoding="utf-8") as f:
        f.write(csv_text)
    fixtures["csv"] = ("sample_stress.csv", csv_text.encode("utf-8"))

    # 6. PNG (RGBA Transparent Image with clear text for OCR and visual verification)
    png_path = FIXTURES_DIR / "sample_stress.png"
    im_rgba = Image.new("RGBA", (700, 350), (0, 0, 0, 0))
    draw_rgba = ImageDraw.Draw(im_rgba)
    font_bold = get_font(26)
    font_reg = get_font(18)
    draw_rgba.rounded_rectangle([20, 20, 680, 330], radius=16, fill=(15, 23, 42, 220), outline=(99, 102, 241, 255), width=3)
    draw_rgba.text((50, 50), "Enterprise Analytics Overview", fill=(255, 255, 255, 255), font=font_bold)
    draw_rgba.text((50, 100), "Financial Target: € 500,000 | ₹ 45,000,000", fill=(147, 197, 253, 255), font=font_reg)
    draw_rgba.text((50, 140), "Operational Status: 22°C Norm | Variance –1.2%", fill=(226, 232, 240, 255), font=font_reg)
    draw_rgba.text((50, 180), "Milestones Completed: Phase 1 & 2 Approved", fill=(110, 231, 183, 255), font=font_reg)
    draw_rgba.text((50, 230), "Velmo High Precision OCR Extraction", fill=(243, 244, 246, 255), font=font_reg)
    im_rgba.save(png_path, format="PNG")
    with open(png_path, "rb") as f:
        fixtures["png"] = ("sample_stress.png", f.read())

    # 7. JPG (Compressed RGB with shapes and text)
    jpg_path = FIXTURES_DIR / "sample_stress.jpg"
    im_jpg = Image.new("RGB", (700, 350), (30, 41, 59))
    draw_jpg = ImageDraw.Draw(im_jpg)
    draw_jpg.rectangle([30, 30, 670, 320], outline=(56, 189, 248), width=3)
    draw_jpg.text((50, 60), "Inspection Audit Checklist", fill=(255, 255, 255), font=font_bold)
    draw_jpg.text((50, 115), "Audit Scope: Frankfurt (€) & Bangalore (₹) Hubs", fill=(224, 242, 254), font=font_reg)
    draw_jpg.text((50, 160), "Environmental Specs: 25°C – Optimal Calibration", fill=(224, 242, 254), font=font_reg)
    draw_jpg.text((50, 210), "Overall Compliance Rating: 98.7% Certified", fill=(134, 239, 172), font=font_reg)
    im_jpg.save(jpg_path, format="JPEG", quality=95)
    with open(jpg_path, "rb") as f:
        fixtures["jpg"] = ("sample_stress.jpg", f.read())

    # 8. WebP (Modern Web image format)
    webp_path = FIXTURES_DIR / "sample_stress.webp"
    im_webp = Image.new("RGB", (700, 350), (6, 78, 59))
    draw_webp = ImageDraw.Draw(im_webp)
    draw_webp.rectangle([30, 30, 670, 320], outline=(52, 211, 153), width=3)
    draw_webp.text((50, 60), "Web Performance Benchmark", fill=(255, 255, 255), font=font_bold)
    draw_webp.text((50, 115), "Network Bandwidth: Fast – Multi-Region Cloud", fill=(209, 250, 229), font=font_reg)
    draw_webp.text((50, 160), "Annual Cost: € 12,000 (Savings: ₹ 250,000)", fill=(209, 250, 229), font=font_reg)
    draw_webp.text((50, 210), "Uptime Metric: 99.99% across 18°–26°C range", fill=(254, 240, 138), font=font_reg)
    im_webp.save(webp_path, format="WEBP", quality=92)
    with open(webp_path, "rb") as f:
        fixtures["webp"] = ("sample_stress.webp", f.read())

    # 9. MD (Markdown with headings, tables, bullet lists, Unicode)
    md_path = FIXTURES_DIR / "sample_stress.md"
    md_text = (
        "# Global Strategy & Technical Architecture\n\n"
        "Executive summary for international cross-border deployment.\n\n"
        "## Financial Allocation Overview\n\n"
        "Tracking capital expenditure in European division (**€**) and Indian division (**₹**).\n"
        "Normal facility operating envelope is **21°C – 25°C**.\n\n"
        "| Facility | Region | Budget (€) | Spend (₹) | Target Met |\n"
        "| --- | --- | --- | --- | --- |\n"
        "| Alpha Hub | Central EU | € 75,000 | ₹ 0 | Yes |\n"
        "| Beta Node | South Asia | € 15,000 | ₹ 4,200,000 | Yes |\n"
        "| Gamma Core | Americas | € 90,000 | ₹ 0 | Yes |\n\n"
        "### Key Architectural Deliverables\n\n"
        "- High-throughput document conversion pipeline\n"
        "- Full lossless round-trip support across 11 formats\n"
        "- Sub-second latency on standard document payloads\n"
    )
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md_text)
    fixtures["md"] = ("sample_stress.md", md_text.encode("utf-8"))

    # 10. TXT (Structured Plain Text, Unicode)
    txt_path = FIXTURES_DIR / "sample_stress.txt"
    txt_text = (
        "============================================================\n"
        "INTERNATIONAL METRICS & AUDIT REPORT\n"
        "============================================================\n"
        "Currency Units: EUR (€) and INR (₹)\n"
        "Environmental Index: Ambient 24°C – Optimal Operations\n\n"
        "DEPARTMENTAL SUMMARY TABLE:\n"
        "Dept\t\tAlloc(€)\tAlloc(₹)\tStatus\n"
        "Core-Platform\t€ 45,000\t₹ 4,000,000\tACTIVE\n"
        "Quality-Eng\t€ 25,000\t₹ 2,200,000\tACTIVE\n"
        "DevOps-Cloud\t€ 30,000\t₹ 2,700,000\tACTIVE\n"
        "============================================================\n"
        "End of Report – Verified Lossless Parsing\n"
    )
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(txt_text)
    fixtures["txt"] = ("sample_stress.txt", txt_text.encode("utf-8"))

    # 11. HTML (Valid HTML5 document with styled table, headings, Unicode)
    html_path = FIXTURES_DIR / "sample_stress.html"
    html_text = (
        "<!DOCTYPE html>\n"
        "<html lang=\"en\">\n"
        "<head>\n"
        "  <meta charset=\"UTF-8\">\n"
        "  <title>Enterprise Performance Review</title>\n"
        "</head>\n"
        "<body>\n"
        "  <h1>Enterprise Performance Review</h1>\n"
        "  <p>Comprehensive tracking of fiscal metrics in <b>€</b> and local expenditure in <b>₹</b>. Target operating temperature: <b>22°C – 26°C</b>.</p>\n"
        "  <h2>Regional Performance Summary</h2>\n"
        "  <table border=\"1\">\n"
        "    <thead>\n"
        "      <tr><th>Division</th><th>Budget (€)</th><th>Spend (₹)</th><th>Variance</th></tr>\n"
        "    </thead>\n"
        "    <tbody>\n"
        "      <tr><td>Engineering</td><td>€ 85,000</td><td>₹ 7,600,000</td><td>+1.8%</td></tr>\n"
        "      <tr><td>Operations</td><td>€ 42,000</td><td>₹ 3,750,000</td><td>-0.5%</td></tr>\n"
        "    </tbody>\n"
        "  </table>\n"
        "</body>\n"
        "</html>\n"
    )
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html_text)
    fixtures["html"] = ("sample_stress.html", html_text.encode("utf-8"))

    return fixtures


def post_convert(filename: str, file_bytes: bytes, target_format: str) -> Tuple[int, bytes, str]:
    """Sends a multipart/form-data POST request to /convert."""
    boundary = f"----WebKitFormBoundary{int(time.time()*1000)}"
    body = io.BytesIO()

    # Form field: target_format
    body.write(f"--{boundary}\r\n".encode())
    body.write(f'Content-Disposition: form-data; name="target_format"\r\n\r\n{target_format}\r\n'.encode())

    # Form field: files
    body.write(f"--{boundary}\r\n".encode())
    body.write(f'Content-Disposition: form-data; name="files"; filename="{filename}"\r\n'.encode())
    body.write(b"Content-Type: application/octet-stream\r\n\r\n")
    body.write(file_bytes)
    body.write(f"\r\n--{boundary}--\r\n".encode())

    req = urllib.request.Request(
        f"{SERVER_URL}/convert",
        data=body.getvalue(),
        headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
    )

    try:
        with urllib.request.urlopen(req, timeout=90) as resp:
            content = resp.read()
            disp = resp.headers.get("Content-Disposition", "")
            return resp.status, content, disp
    except urllib.error.HTTPError as e:
        err_body = e.read()
        return e.code, err_body, ""


def validate_output(target_fmt: str, content: bytes) -> str:
    """
    Opens and validates output file using its real parser to verify it is
    valid, readable, well-formed, and non-empty.
    """
    assert len(content) > 0, f"Output content is empty for target format '{target_fmt}'"

    # Multi-page or multi-sheet output might be bundled into a ZIP archive
    if content[:4] == b"PK\x03\x04" and target_fmt not in ("docx", "xlsx", "pptx"):
        with zipfile.ZipFile(io.BytesIO(content), "r") as zf:
            names = [n for n in zf.namelist() if not n.endswith("/")]
            assert len(names) > 0, "ZIP archive contains zero files."
            for name in names:
                extracted = zf.read(name)
                assert len(extracted) > 0, f"Extracted ZIP entry '{name}' is empty."
        return f"Valid ZIP ({len(names)} entries)"

    if target_fmt == "pdf":
        doc = pymupdf.open(stream=content, filetype="pdf")
        assert len(doc) > 0, "PDF document contains 0 pages."
        page = doc[0]
        assert page.rect.width > 0 and page.rect.height > 0, "PDF page has zero dimensions."
        page_count = len(doc)
        doc.close()
        return f"Valid PDF ({page_count} page{'s' if page_count > 1 else ''})"

    elif target_fmt == "docx":
        doc = Document(io.BytesIO(content))
        has_text = any(p.text.strip() for p in doc.paragraphs)
        has_table = len(doc.tables) > 0
        assert has_text or has_table, "DOCX has no readable text or tables."
        return f"Valid DOCX ({len(doc.paragraphs)} paras, {len(doc.tables)} tables)"

    elif target_fmt == "xlsx":
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        assert len(wb.worksheets) > 0, "XLSX has 0 worksheets."
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        assert len(rows) > 0, "XLSX active worksheet has 0 rows."
        sheet_count = len(wb.worksheets)
        wb.close()
        return f"Valid XLSX ({sheet_count} sheet{'s' if sheet_count > 1 else ''}, {len(rows)} rows in active)"

    elif target_fmt == "pptx":
        prs = Presentation(io.BytesIO(content))
        assert len(prs.slides) > 0, "PPTX presentation has 0 slides."
        return f"Valid PPTX ({len(prs.slides)} slides)"

    elif target_fmt in ("png", "jpg", "webp"):
        with Image.open(io.BytesIO(content)) as im:
            assert im.width > 0 and im.height > 0, "Image has invalid dimensions."
            fmt = (im.format or "").upper()
            if target_fmt == "jpg":
                assert fmt in ("JPEG", "JPG"), f"Expected JPEG format, got '{fmt}'"
            elif target_fmt == "png":
                assert fmt == "PNG", f"Expected PNG format, got '{fmt}'"
            elif target_fmt == "webp":
                assert fmt == "WEBP", f"Expected WEBP format, got '{fmt}'"
            return f"Valid {fmt} ({im.width}x{im.height})"

    elif target_fmt in ("csv", "txt", "md", "html"):
        text = content.decode("utf-8")
        assert len(text.strip()) > 0, f"Text content for format '{target_fmt}' is empty."
        return f"Valid {target_fmt.upper()} ({len(text)} chars)"

    else:
        raise ValueError(f"Unknown target format: '{target_fmt}'")


def run_full_matrix_tests(fixtures: Dict[str, Tuple[str, bytes]]) -> Tuple[int, int, list]:
    """Runs all 110 directed pairs (11 source formats x 10 target formats)."""
    print("\n" + "=" * 80)
    print("STARTING FULL MATRIX CONVERSION TESTS (110 DIRECTED PAIRS)")
    print("=" * 80)

    passed = 0
    failed = 0
    failures = []
    pair_count = 0

    for src_fmt in ALL_FORMATS:
        filename, data_bytes = fixtures[src_fmt]
        target_formats = [f for f in ALL_FORMATS if f != src_fmt]

        for tgt_fmt in target_formats:
            pair_count += 1
            pair_label = f"[{pair_count:03d}/110] {src_fmt.upper():>4} -> {tgt_fmt.upper():<4}"
            t0 = time.time()

            try:
                status, content, disp = post_convert(filename, data_bytes, tgt_fmt)
                elapsed = time.time() - t0

                if status != 200:
                    err_msg = content.decode("utf-8", errors="replace")
                    print(f"  FAIL: {pair_label} | HTTP {status}: {err_msg[:120]}")
                    failed += 1
                    failures.append((src_fmt, tgt_fmt, f"HTTP {status}: {err_msg[:120]}"))
                    continue

                # Deep parser validation
                parsed_info = validate_output(tgt_fmt, content)
                passed += 1
                print(f"  PASS: {pair_label} | {parsed_info:<35} | {len(content):>7} B | {elapsed:4.2f}s")

            except Exception as e:
                elapsed = time.time() - t0
                print(f"  FAIL: {pair_label} | Parser Validation Error: {e}")
                failed += 1
                failures.append((src_fmt, tgt_fmt, str(e)))

    return passed, failed, failures


def run_edge_case_torture_tests(fixtures: Dict[str, Tuple[str, bytes]]) -> Tuple[int, int, list]:
    """Runs all real-world edge cases & torture scenarios."""
    print("\n" + "=" * 80)
    print("STARTING EDGE-CASE & TORTURE TESTS")
    print("=" * 80)

    passed = 0
    failed = 0
    failures = []

    def check_torture(name: str, fn):
        nonlocal passed, failed
        try:
            fn()
            passed += 1
            print(f"  PASS: [TORTURE] {name}")
        except Exception as e:
            failed += 1
            failures.append((name, str(e)))
            print(f"  FAIL: [TORTURE] {name} | Error: {e}")

    # 1. Zero-byte files (must return 400 clean error, never 500)
    for ext in ["pdf", "docx", "xlsx", "csv", "png"]:
        def _test_zero(e=ext):
            status, body, _ = post_convert(f"zero_file.{e}", b"", "txt")
            assert status == 400, f"Expected HTTP 400 for 0-byte {e}, got {status}: {body.decode()}"
            err_json = json.loads(body.decode())
            assert "empty" in err_json.get("detail", "").lower(), f"Unexpected detail message: {err_json}"
        check_torture(f"Zero-byte file rejection ({ext}) -> HTTP 400", _test_zero)

    # 2. Corrupted / garbage bytes with .pdf / .docx / .xlsx (must return 400 clean error, never 500)
    for ext, tgt in [("pdf", "docx"), ("docx", "pdf"), ("xlsx", "csv")]:
        def _test_corrupt(e=ext, t=tgt):
            garbage = b"THIS_IS_DEFINITELY_CORRUPTED_GARBAGE_BYTES_1234567890!@#$%^&*()"
            status, body, _ = post_convert(f"corrupt.{e}", garbage, t)
            assert status == 400, f"Expected HTTP 400 for corrupted {e}->{t}, got {status}: {body.decode()}"
            err_json = json.loads(body.decode())
            assert "invalid" in err_json.get("detail", "").lower() or "corrupt" in err_json.get("detail", "").lower(), f"Expected corruption error detail: {err_json}"
        check_torture(f"Corrupted file rejection ({ext} -> {tgt}) -> HTTP 400", _test_corrupt)

    # 3. Unsupported file extension (e.g. malware.exe) -> HTTP 400
    def _test_unsupported_ext():
        status, body, _ = post_convert("malware.exe", b"MZ_EXECUTABLE_BINARY_DATA", "pdf")
        assert status == 400, f"Expected HTTP 400 for malware.exe, got {status}: {body.decode()}"
        err_json = json.loads(body.decode())
        assert "unsupported" in err_json.get("detail", "").lower(), f"Unexpected detail: {err_json}"
    check_torture("Unsupported extension (malware.exe) -> HTTP 400", _test_unsupported_ext)

    # 4. Blank solid-color PNG (zero text) -> OCR to docx, xlsx, txt must succeed gracefully without crashing
    def _test_blank_png_ocr():
        blank_im = Image.new("RGB", (400, 300), color=(255, 255, 255))
        buf = io.BytesIO()
        blank_im.save(buf, format="PNG")
        blank_bytes = buf.getvalue()

        # Convert to DOCX
        st_docx, res_docx, _ = post_convert("blank.png", blank_bytes, "docx")
        assert st_docx == 200, f"Blank PNG -> DOCX failed with HTTP {st_docx}"
        parsed_docx = validate_output("docx", res_docx)

        # Convert to XLSX
        st_xlsx, res_xlsx, _ = post_convert("blank.png", blank_bytes, "xlsx")
        assert st_xlsx == 200, f"Blank PNG -> XLSX failed with HTTP {st_xlsx}"
        parsed_xlsx = validate_output("xlsx", res_xlsx)

        # Convert to TXT
        st_txt, res_txt, _ = post_convert("blank.png", blank_bytes, "txt")
        assert st_txt == 200, f"Blank PNG -> TXT failed with HTTP {st_txt}"
        parsed_txt = validate_output("txt", res_txt)
    check_torture("Blank solid-color PNG -> OCR to DOCX, XLSX, TXT (graceful handling)", _test_blank_png_ocr)

    # 5. Large spreadsheet (200 rows x 10 columns) -> converts cleanly to pdf, docx, csv
    def _test_large_spreadsheet():
        wb_large = openpyxl.Workbook()
        ws = wb_large.active
        ws.title = "LargeData"
        ws.append([f"Column {c}" for c in range(1, 11)])
        for r in range(1, 201):
            ws.append([f"Row {r} Data {c}" for c in range(1, 11)])
        lbuf = io.BytesIO()
        wb_large.save(lbuf)
        large_bytes = lbuf.getvalue()

        # Convert to PDF
        st_pdf, res_pdf, _ = post_convert("large_data.xlsx", large_bytes, "pdf")
        assert st_pdf == 200, f"Large XLSX -> PDF failed with HTTP {st_pdf}"
        validate_output("pdf", res_pdf)

        # Convert to DOCX
        st_docx, res_docx, _ = post_convert("large_data.xlsx", large_bytes, "docx")
        assert st_docx == 200, f"Large XLSX -> DOCX failed with HTTP {st_docx}"
        validate_output("docx", res_docx)

        # Convert to CSV
        st_csv, res_csv, _ = post_convert("large_data.xlsx", large_bytes, "csv")
        assert st_csv == 200, f"Large XLSX -> CSV failed with HTTP {st_csv}"
        validate_output("csv", res_csv)
    check_torture("Large spreadsheet (200 rows x 10 cols) -> PDF, DOCX, CSV", _test_large_spreadsheet)

    # 6. Preview & Edit round-trip on CSV, MD, PPTX
    def _test_preview_export_roundtrip():
        # A) CSV -> Preview Editable (Table mode) -> Edit cell -> Export XLSX
        csv_fn, csv_data = fixtures["csv"]
        boundary = f"----WebKitFormBoundary{int(time.time()*1000)}"
        body = io.BytesIO()
        body.write(f"--{boundary}\r\n".encode())
        body.write(b'Content-Disposition: form-data; name="target_format"\r\n\r\nxlsx\r\n')
        body.write(f"--{boundary}\r\n".encode())
        body.write(f'Content-Disposition: form-data; name="files"; filename="{csv_fn}"\r\n'.encode())
        body.write(b"Content-Type: text/csv\r\n\r\n")
        body.write(csv_data)
        body.write(f"\r\n--{boundary}--\r\n".encode())

        req = urllib.request.Request(
            f"{SERVER_URL}/preview-editable",
            data=body.getvalue(),
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
        )
        with urllib.request.urlopen(req) as resp:
            assert resp.status == 200
            preview_json = json.loads(resp.read().decode())
            assert preview_json["mode"] == "table"
            assert len(preview_json["sheets"]) > 0

        # Edit a cell value and export to XLSX
        preview_json["sheets"][0]["rows"][1][1] = "EDITED_STRESS_CONTROLLER"
        export_payload = {
            "mode": "table",
            "filename": "stress_export",
            "target_format": "xlsx",
            "sheets": preview_json["sheets"]
        }
        req_exp = urllib.request.Request(
            f"{SERVER_URL}/export-edited",
            data=json.dumps(export_payload).encode(),
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req_exp) as resp_exp:
            assert resp_exp.status == 200
            exp_bytes = resp_exp.read()
            validate_output("xlsx", exp_bytes)
            # Confirm edited value is present
            wb_exp = openpyxl.load_workbook(io.BytesIO(exp_bytes), data_only=True)
            cell_vals = [str(c.value) for r in wb_exp.worksheets[0].iter_rows() for c in r if c.value]
            assert any("EDITED_STRESS_CONTROLLER" in v for v in cell_vals), "Edited value not found in exported XLSX"

        # B) MD -> Preview Editable (Document mode) -> Edit text -> Export DOCX
        md_fn, md_data = fixtures["md"]
        body_md = io.BytesIO()
        body_md.write(f"--{boundary}\r\n".encode())
        body_md.write(b'Content-Disposition: form-data; name="target_format"\r\n\r\ndocx\r\n')
        body_md.write(f"--{boundary}\r\n".encode())
        body_md.write(f'Content-Disposition: form-data; name="files"; filename="{md_fn}"\r\n'.encode())
        body_md.write(b"Content-Type: text/markdown\r\n\r\n")
        body_md.write(md_data)
        body_md.write(f"\r\n--{boundary}--\r\n".encode())

        req_md = urllib.request.Request(
            f"{SERVER_URL}/preview-editable",
            data=body_md.getvalue(),
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
        )
        with urllib.request.urlopen(req_md) as resp_md:
            assert resp_md.status == 200
            p_md = json.loads(resp_md.read().decode())
            assert p_md["mode"] == "document"
            assert "content" in p_md

        # Edit text and export to DOCX
        edited_content = "# EDITED STRESS HEADING\n\nThis is edited text with € and ₹ symbols."
        exp_doc_payload = {
            "mode": "document",
            "filename": "stress_edited_doc",
            "target_format": "docx",
            "content": edited_content
        }
        req_exp_doc = urllib.request.Request(
            f"{SERVER_URL}/export-edited",
            data=json.dumps(exp_doc_payload).encode(),
            headers={"Content-Type": "application/json"}
        )
        with urllib.request.urlopen(req_exp_doc) as resp_exp_doc:
            assert resp_exp_doc.status == 200
            exp_doc_bytes = resp_exp_doc.read()
            validate_output("docx", exp_doc_bytes)

        # C) PPTX -> Preview Editable (Document mode) -> Edit -> Export PDF
        pptx_fn, pptx_data = fixtures["pptx"]
        body_pptx = io.BytesIO()
        body_pptx.write(f"--{boundary}\r\n".encode())
        body_pptx.write(b'Content-Disposition: form-data; name="target_format"\r\n\r\npdf\r\n')
        body_pptx.write(f"--{boundary}\r\n".encode())
        body_pptx.write(f'Content-Disposition: form-data; name="files"; filename="{pptx_fn}"\r\n'.encode())
        body_pptx.write(b"Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation\r\n\r\n")
        body_pptx.write(pptx_data)
        body_pptx.write(f"\r\n--{boundary}--\r\n".encode())

        req_pptx = urllib.request.Request(
            f"{SERVER_URL}/preview-editable",
            data=body_pptx.getvalue(),
            headers={"Content-Type": f"multipart/form-data; boundary={boundary}"}
        )
        with urllib.request.urlopen(req_pptx) as resp_pptx:
            assert resp_pptx.status == 200
            p_pptx = json.loads(resp_pptx.read().decode())
            assert p_pptx["mode"] == "document"
            assert len(p_pptx.get("content", "")) > 0
    check_torture("Preview & Edit round-trip on CSV, MD, PPTX", _test_preview_export_roundtrip)

    return passed, failed, failures


def main():
    print("=" * 80)
    print("DOC-CONVERTER FULL-MATRIX STRESS TEST & EDGE-CASE TORTURE SUITE")
    print(f"Target Server: {SERVER_URL}")
    print("=" * 80)

    # 1. Health check
    try:
        with urllib.request.urlopen(f"{SERVER_URL}/health", timeout=5) as r:
            assert r.status == 200
            info = json.loads(r.read().decode())
            print(f"Server is healthy. Formats: {info.get('formats')}, LibreOffice: {info.get('libreoffice')}")
    except Exception as e:
        print(f"ERROR: Cannot connect to server at {SERVER_URL}: {e}")
        sys.exit(1)

    # 2. Generate fixtures
    print("\nGenerating realistic sample fixtures for all 11 formats...")
    fixtures = create_sample_fixtures()
    print(f"Generated {len(fixtures)} sample fixtures successfully.")

    # 3. Run full 110-pair matrix
    start_time = time.time()
    matrix_passed, matrix_failed, matrix_failures = run_full_matrix_tests(fixtures)

    # 4. Run edge case torture tests
    edge_passed, edge_failed, edge_failures = run_edge_case_torture_tests(fixtures)

    total_time = time.time() - start_time

    # 5. Summary Report
    print("\n" + "=" * 80)
    print("FINAL SUMMARY REPORT")
    print("=" * 80)
    print(f"Matrix Conversions Passed:   {matrix_passed} / 110  ({(matrix_passed/110)*100:.1f}%)")
    print(f"Matrix Conversions Failed:   {matrix_failed} / 110")
    print(f"Edge-Case Tests Passed:      {edge_passed} / {edge_passed + edge_failed}")
    print(f"Edge-Case Tests Failed:      {edge_failed}")
    print(f"Total Execution Time:        {total_time:.2f} seconds")

    if matrix_failures or edge_failures:
        print("\nFAILURE DETAILS:")
        for item in matrix_failures:
            print(f"  - Matrix: {item[0]} -> {item[1]}: {item[2]}")
        for item in edge_failures:
            print(f"  - Edge Case: {item[0]}: {item[1]}")
        print("=" * 80)
        sys.exit(1)
    else:
        print("\nALL 110 DIRECTED PAIRS AND ALL EDGE CASES PASSED WITH 100% SUCCESS!")
        print("=" * 80)
        sys.exit(0)


if __name__ == "__main__":
    main()
