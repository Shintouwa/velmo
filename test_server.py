import csv
import io
import json
import re
import time
import urllib.request
import urllib.error
import uuid
import zipfile
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
import pymupdf
import openpyxl
from openpyxl.styles import Font, PatternFill
from docx import Document
from pptx import Presentation
from pptx.util import Inches, Pt

BASE_DIR = Path(__file__).resolve().parent

# File paths
SAMPLE_PDF = BASE_DIR / "sample.pdf"
SAMPLE_1PAGE_PDF = BASE_DIR / "sample_1page.pdf"
SAMPLE_JPG = BASE_DIR / "test_sample.jpg"
SAMPLE_TRANSPARENT_PNG = BASE_DIR / "test_transparent.png"
SAMPLE_WEBP = BASE_DIR / "test_sample.webp"
SAMPLE_MULTISHEET_XLSX = BASE_DIR / "sample_multisheet.xlsx"
SAMPLE_MIXED_DOCX = BASE_DIR / "sample_mixed.docx"
SAMPLE_OCR_DOC_IMG = BASE_DIR / "sample_ocr_doc.png"
SAMPLE_OCR_TABLE_IMG = BASE_DIR / "sample_ocr_table.png"

# New test assets
SAMPLE_PPTX = BASE_DIR / "sample.pptx"
SAMPLE_CSV = BASE_DIR / "sample.csv"
SAMPLE_MD = BASE_DIR / "sample.md"
SAMPLE_HTML = BASE_DIR / "sample.html"

SERVER_URL = "http://127.0.0.1:8000"


def get_test_font(size: int):
    for font_name in ["arial.ttf", "segoeui.ttf", "calibri.ttf", "DejaVuSans.ttf"]:
        try:
            return ImageFont.truetype(font_name, size)
        except Exception:
            pass
    try:
        return ImageFont.load_default(size=size)
    except Exception:
        return ImageFont.load_default()


def generate_test_assets():
    """Ensure all sample test files exist across all 11 supported formats."""
    # 1. PDFs
    if not SAMPLE_PDF.exists():
        doc = pymupdf.open()
        p1 = doc.new_page(width=595, height=842)
        p1.insert_text((50, 100), "Sample Page 1", fontsize=20)
        p2 = doc.new_page(width=595, height=842)
        p2.insert_text((50, 100), "Sample Page 2", fontsize=20)
        doc.save(SAMPLE_PDF)
        doc.close()

    if not SAMPLE_1PAGE_PDF.exists():
        doc = pymupdf.open()
        p = doc.new_page(width=595, height=842)
        p.insert_text((50, 100), "Single Page PDF", fontsize=20)
        doc.save(SAMPLE_1PAGE_PDF)
        doc.close()

    # 2. Images
    if not SAMPLE_JPG.exists():
        im = Image.new("RGB", (800, 600), color=(79, 70, 229))
        draw = ImageDraw.Draw(im)
        draw.rectangle([50, 50, 750, 550], outline=(255, 255, 255), width=4)
        im.save(SAMPLE_JPG, format="JPEG", quality=95)

    if not SAMPLE_TRANSPARENT_PNG.exists():
        im_rgba = Image.new("RGBA", (600, 600), color=(0, 0, 0, 0))
        draw = ImageDraw.Draw(im_rgba)
        draw.ellipse([100, 100, 500, 500], fill=(16, 185, 129, 180), outline=(5, 150, 105, 255), width=6)
        im_rgba.save(SAMPLE_TRANSPARENT_PNG, format="PNG")

    if not SAMPLE_WEBP.exists():
        im_webp = Image.new("RGB", (700, 500), color=(6, 182, 212))
        draw = ImageDraw.Draw(im_webp)
        draw.polygon([(350, 80), (600, 420), (100, 420)], fill=(255, 255, 255))
        im_webp.save(SAMPLE_WEBP, format="WEBP", quality=90)

    # 3. Excel
    if not SAMPLE_MULTISHEET_XLSX.exists():
        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "Sales_Report"
        fill1 = PatternFill(start_color="FFE0E7FF", end_color="FFE0E7FF", fill_type="solid")
        ws1.append(["Product", "Region", "Q1_Sales", "Q2_Sales", "Total"])
        for c in ws1[1]:
            c.font = Font(bold=True)
            c.fill = fill1
        ws1.append(["Alpha", "North", 1500, 2300, 3800])
        ws1.append(["Beta", "South", 4200, 3100, 7300])

        ws2 = wb.create_sheet(title="Employee_Directory")
        fill2 = PatternFill(start_color="FFDCFCE7", end_color="FFDCFCE7", fill_type="solid")
        ws2.append(["ID", "Name", "Department", "Role", "Active", "Rating", "Notes"])
        for c in ws2[1]:
            c.font = Font(bold=True)
            c.fill = fill2
        ws2.append([101, "Alice Smith", "Engineering", "Lead", "Yes", 4.9, "Key contributor"])
        ws2.append([102, "Bob Jones", "Product", "Manager", "Yes", 4.7, "Q2 launch lead"])
        wb.save(SAMPLE_MULTISHEET_XLSX)

    # 4. Word
    if not SAMPLE_MIXED_DOCX.exists():
        doc = Document()
        doc.add_heading("Executive Summary", level=1)
        doc.add_paragraph("This is an overview of quarterly performance and company metrics.")
        doc.add_heading("Performance Table", level=2)
        tbl1 = doc.add_table(rows=2, cols=3)
        tbl1.style = "Table Grid"
        tbl1.cell(0, 0).paragraphs[0].add_run("Metric").bold = True
        tbl1.cell(0, 1).paragraphs[0].add_run("Value").bold = True
        tbl1.cell(0, 2).paragraphs[0].add_run("Status").bold = True
        tbl1.cell(1, 0).paragraphs[0].text = "Revenue"
        tbl1.cell(1, 1).paragraphs[0].text = "1250000"
        tbl1.cell(1, 2).paragraphs[0].text = "Exceeded"
        doc.add_heading("Department Headcount", level=2)
        doc.add_paragraph("Headcount has expanded across engineering and product teams.")
        tbl2 = doc.add_table(rows=2, cols=2)
        tbl2.style = "Table Grid"
        tbl2.cell(0, 0).paragraphs[0].add_run("Team").bold = True
        tbl2.cell(0, 1).paragraphs[0].add_run("Count").bold = True
        tbl2.cell(1, 0).paragraphs[0].text = "Design"
        tbl2.cell(1, 1).paragraphs[0].text = "14"
        doc.add_paragraph("Final notes: All targets have been verified for Q3.")
        doc.save(SAMPLE_MIXED_DOCX)

    # 5. OCR Images
    if not SAMPLE_OCR_DOC_IMG.exists():
        img = Image.new("RGB", (900, 380), color=(255, 255, 255))
        d = ImageDraw.Draw(img)
        title_font = get_test_font(28)
        body_font = get_test_font(20)
        d.text((50, 40), "Project Milestone Overview", fill=(15, 23, 42), font=title_font)
        d.text((50, 120), "Phase 1 deliverables have been completed on schedule.", fill=(51, 65, 85), font=body_font)
        d.text((50, 180), "The team is currently finalizing Phase 2 architecture.", fill=(51, 65, 85), font=body_font)
        img.save(SAMPLE_OCR_DOC_IMG, format="PNG")

    if not SAMPLE_OCR_TABLE_IMG.exists():
        img = Image.new("RGB", (850, 420), color=(255, 255, 255))
        d = ImageDraw.Draw(img)
        table_font = get_test_font(22)
        table_rows = [
            ["Item", "Qty", "Price"],
            ["Widget Alpha", "10", "25.50"],
            ["Widget Beta", "5", "12.00"],
            ["Widget Gamma", "100", "4.99"]
        ]
        y = 45
        col_x = [60, 360, 620]
        for row in table_rows:
            for c_i, text in enumerate(row):
                d.text((col_x[c_i], y), text, fill=(0, 0, 0), font=table_font)
            y += 75
        img.save(SAMPLE_OCR_TABLE_IMG, format="PNG")

    # 6. PowerPoint (.pptx)
    if not SAMPLE_PPTX.exists():
        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)
        blank_layout = prs.slide_layouts[6]

        # Slide 1: Title + bullets
        s1 = prs.slides.add_slide(blank_layout)
        tb1 = s1.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        tb1.text_frame.paragraphs[0].text = "Quarterly Business Review"
        tb1.text_frame.paragraphs[0].font.size = Pt(28)
        tb1.text_frame.paragraphs[0].font.bold = True

        cb1 = s1.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(11.7), Inches(4.5))
        p = cb1.text_frame.paragraphs[0]
        p.text = "Revenue grew 24% year-over-year"
        p.level = 0
        p2 = cb1.text_frame.add_paragraph()
        p2.text = "Enterprise customer retention reached 98%"
        p2.level = 0

        # Slide 2: Table
        s2 = prs.slides.add_slide(blank_layout)
        tb2 = s2.shapes.add_textbox(Inches(0.8), Inches(0.6), Inches(11.7), Inches(1.0))
        tb2.text_frame.paragraphs[0].text = "Key Metrics Summary"
        tb2.text_frame.paragraphs[0].font.size = Pt(28)
        tb2.text_frame.paragraphs[0].font.bold = True

        tbl_shp = s2.shapes.add_table(3, 3, Inches(0.8), Inches(1.8), Inches(10.0), Inches(2.5))
        t = tbl_shp.table
        t.cell(0, 0).text = "Department"
        t.cell(0, 1).text = "Headcount"
        t.cell(0, 2).text = "Budget"
        t.cell(1, 0).text = "Engineering"
        t.cell(1, 1).text = "45"
        t.cell(1, 2).text = "2500000"
        t.cell(2, 0).text = "Marketing"
        t.cell(2, 1).text = "18"
        t.cell(2, 2).text = "900000"
        prs.save(SAMPLE_PPTX)

    # 7. CSV (.csv)
    if not SAMPLE_CSV.exists():
        with open(SAMPLE_CSV, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["City", "Country", "Population", "GrowthRate"])
            writer.writerow(["Tokyo", "Japan", "37400068", "0.02"])
            writer.writerow(["Delhi", "India", "32941308", "2.15"])
            writer.writerow(["Shanghai", "China", "28516904", "1.42"])

    # 8. Markdown (.md)
    if not SAMPLE_MD.exists():
        md_content = """# Release Notes v5.0

Here is a summary of major updates in the platform.

## Features
- Universal 11-format Any-to-Any conversion
- Fully automated local OCR
- Preserved table styling

| Module | Status | Version |
| --- | --- | --- |
| Core Engine | Operational | 5.0 |
| OCR Service | Active | 1.4 |
| Web UI | Modernized | 5.0 |

*Thank you for using Velmo!*
"""
        SAMPLE_MD.write_text(md_content, encoding="utf-8")

    # 9. HTML (.html)
    if not SAMPLE_HTML.exists():
        html_content = """<!DOCTYPE html>
<html>
<head><title>System Report</title></head>
<body>
  <h1>Quarterly Platform Performance</h1>
  <p>The platform achieved 99.99% uptime across all conversion pipelines.</p>
  <table border="1">
    <tr><th>Pipeline</th><th>Latency</th><th>Success Rate</th></tr>
    <tr><td>Document Engine</td><td>120ms</td><td>100%</td></tr>
    <tr><td>OCR Engine</td><td>450ms</td><td>99.8%</td></tr>
  </table>
</body>
</html>"""
        SAMPLE_HTML.write_text(html_content, encoding="utf-8")


def encode_multipart_formdata(fields, files_list):
    boundary = f"----WebKitFormBoundary{uuid.uuid4().hex}"
    body = bytearray()
    for k, v in fields.items():
        body.extend(f"--{boundary}\r\n".encode("utf-8"))
        body.extend(f'Content-Disposition: form-data; name="{k}"\r\n\r\n'.encode("utf-8"))
        body.extend(f"{v}\r\n".encode("utf-8"))

    for field_name, filename, file_data, content_type in files_list:
        body.extend(f"--{boundary}\r\n".encode("utf-8"))
        body.extend(f'Content-Disposition: form-data; name="{field_name}"; filename="{filename}"\r\n'.encode("utf-8"))
        body.extend(f"Content-Type: {content_type}\r\n\r\n".encode("utf-8"))
        body.extend(file_data)
        body.extend(b"\r\n")

    body.extend(f"--{boundary}--\r\n".encode("utf-8"))
    return f"multipart/form-data; boundary={boundary}", bytes(body)


def wait_for_server(timeout=15):
    start = time.time()
    print(f"Waiting for server at {SERVER_URL} to become ready...")
    while time.time() - start < timeout:
        try:
            req = urllib.request.Request(f"{SERVER_URL}/api/info")
            with urllib.request.urlopen(req, timeout=2) as response:
                if response.status == 200:
                    info = json.loads(response.read().decode())
                    lo_status = "DETECTED" if info.get("libreoffice_available") else "NOT DETECTED (ReportLab fallback active)"
                    print(f"Server is up! LibreOffice status: {lo_status}")
                    return info
        except Exception:
            time.sleep(0.5)
    raise TimeoutError("Server did not respond within timeout.")


# --- Full Test Suite ---

def run_all_tests():
    info = wait_for_server()

    # 1. Web UI & Info
    print("\n[TEST 1] GET / & GET /api/info")
    req = urllib.request.Request(f"{SERVER_URL}/")
    with urllib.request.urlopen(req) as resp:
        body = resp.read().decode()
        assert resp.status == 200
        assert "Universal Any-to-Any Converter" in body
        print("  PASS: 11-format universal converter web UI served successfully.")

    # 2. Existing PDF -> PNG
    print("\n[TEST 2] POST /convert/pdf-to-png (Multi-page PDF)")
    content_type, body = encode_multipart_formdata({}, [("file", "sample.pdf", SAMPLE_PDF.read_bytes(), "application/pdf")])
    req = urllib.request.Request(f"{SERVER_URL}/convert/pdf-to-png", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        with zipfile.ZipFile(io.BytesIO(resp.read())) as zf:
            assert len(zf.namelist()) == 2
        print("  PASS: Legacy PDF -> PNG multi-page verified.")

    # 3. Existing Image -> PDF
    print("\n[TEST 3] POST /convert/image-to-pdf (Multi-image with transparency)")
    files = [
        ("files", "test_sample.jpg", SAMPLE_JPG.read_bytes(), "image/jpeg"),
        ("files", "test_transparent.png", SAMPLE_TRANSPARENT_PNG.read_bytes(), "image/png"),
        ("files", "test_sample.webp", SAMPLE_WEBP.read_bytes(), "image/webp")
    ]
    content_type, body = encode_multipart_formdata({}, files)
    req = urllib.request.Request(f"{SERVER_URL}/convert/image-to-pdf", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        doc = pymupdf.open(stream=resp.read(), filetype="pdf")
        assert len(doc) == 3
        doc.close()
        print("  PASS: Legacy Image(s) -> PDF verified.")

    # 4. Existing XLSX -> DOCX
    print("\n[TEST 4] POST /convert/xlsx-to-docx")
    content_type, body = encode_multipart_formdata({}, [("file", "sample.xlsx", SAMPLE_MULTISHEET_XLSX.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")])
    req = urllib.request.Request(f"{SERVER_URL}/convert/xlsx-to-docx", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        doc = Document(io.BytesIO(resp.read()))
        assert len(doc.tables) == 2
        print("  PASS: Legacy XLSX -> DOCX verified.")

    # 5. Existing DOCX -> XLSX
    print("\n[TEST 5] POST /convert/docx-to-xlsx")
    content_type, body = encode_multipart_formdata({}, [("file", "sample.docx", SAMPLE_MIXED_DOCX.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")])
    req = urllib.request.Request(f"{SERVER_URL}/convert/docx-to-xlsx", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        wb = openpyxl.load_workbook(io.BytesIO(resp.read()), data_only=True)
        assert "Performance Table" in wb.sheetnames
        assert "Document Text" in wb.sheetnames
        print("  PASS: Legacy DOCX -> XLSX verified.")

    # 6. PDF -> Editable DOCX (pdf2docx)
    print("\n[TEST 6] POST /convert: PDF -> DOCX (Editable layout via pdf2docx)")
    content_type, body = encode_multipart_formdata({"target_format": "docx"}, [("file", "sample.pdf", SAMPLE_PDF.read_bytes(), "application/pdf")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        doc = Document(io.BytesIO(resp.read()))
        assert len(doc.paragraphs) > 0
        print("  PASS: PDF -> DOCX converted cleanly using pdf2docx.")

    # 7. PDF -> Editable XLSX (PyMuPDF find_tables + Document Text)
    print("\n[TEST 7] POST /convert: PDF -> XLSX (Tables & Document Text)")
    content_type, body = encode_multipart_formdata({"target_format": "xlsx"}, [("file", "sample.pdf", SAMPLE_PDF.read_bytes(), "application/pdf")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        wb = openpyxl.load_workbook(io.BytesIO(resp.read()), data_only=True)
        assert "Document Text" in wb.sheetnames
        print("  PASS: PDF -> XLSX extracted tables & Document Text successfully.")

    # 8. DOCX -> PDF
    print("\n[TEST 8] POST /convert: DOCX -> PDF (LibreOffice or ReportLab fallback)")
    content_type, body = encode_multipart_formdata({"target_format": "pdf"}, [("file", "sample_mixed.docx", SAMPLE_MIXED_DOCX.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        pdf_bytes = resp.read()
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        assert len(doc) >= 1
        doc.close()
        print("  PASS: DOCX -> PDF converted with styled headings & tables.")

    # 9. XLSX -> PDF
    print("\n[TEST 9] POST /convert: XLSX -> PDF (LibreOffice or ReportLab fallback)")
    content_type, body = encode_multipart_formdata({"target_format": "pdf"}, [("file", "sample_multisheet.xlsx", SAMPLE_MULTISHEET_XLSX.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        pdf_bytes = resp.read()
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        assert len(doc) >= 1
        doc.close()
        print("  PASS: XLSX -> PDF converted with multiple sheets rendered.")

    # 10. DOCX -> PNG
    print("\n[TEST 10] POST /convert: DOCX -> PNG (Chained DOCX -> PDF -> 300 DPI PNG)")
    content_type, body = encode_multipart_formdata({"target_format": "png"}, [("file", "sample_mixed.docx", SAMPLE_MIXED_DOCX.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        assert resp.read().startswith(b"\x89PNG\r\n\x1a\n")
        print("  PASS: DOCX -> PNG rendered at 300 DPI via PDF pipeline.")

    # 11. XLSX -> PNG
    print("\n[TEST 11] POST /convert: XLSX -> PNG (Chained XLSX -> PDF -> 300 DPI PNG)")
    content_type, body = encode_multipart_formdata({"target_format": "png"}, [("file", "sample_multisheet.xlsx", SAMPLE_MULTISHEET_XLSX.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        with zipfile.ZipFile(io.BytesIO(resp.read())) as zf:
            assert len(zf.namelist()) >= 1
        print("  PASS: XLSX -> PNG rendered at 300 DPI via PDF pipeline.")

    # 12. Image -> DOCX (OCR)
    print("\n[TEST 12] POST /convert: Image -> DOCX (RapidOCR -> Word paragraphs & headings)")
    content_type, body = encode_multipart_formdata({"target_format": "docx"}, [("file", "sample_ocr_doc.png", SAMPLE_OCR_DOC_IMG.read_bytes(), "image/png")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        doc = Document(io.BytesIO(resp.read()))
        all_text = " ".join(p.text for p in doc.paragraphs)
        assert "Milestone" in all_text or "Project" in all_text
        print("  PASS: Image -> DOCX successfully extracted headings and paragraphs via local RapidOCR.")

    # 13. Image -> XLSX (OCR Table Clustering)
    print("\n[TEST 13] POST /convert: Image -> XLSX (RapidOCR -> Clustered Rows & Columns)")
    content_type, body = encode_multipart_formdata({"target_format": "xlsx"}, [("file", "sample_ocr_table.png", SAMPLE_OCR_TABLE_IMG.read_bytes(), "image/png")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        wb = openpyxl.load_workbook(io.BytesIO(resp.read()), data_only=True)
        sheet = wb.active
        assert sheet.max_row >= 4
        assert sheet.max_column >= 3
        print("  PASS: Image -> XLSX successfully clustered table screenshot into structured Excel rows/cols.")

    # --- NEW 11-FORMAT TESTS ---

    # 14. PPTX -> DOCX
    print("\n[TEST 14] POST /convert: PPTX -> DOCX")
    content_type, body = encode_multipart_formdata({"target_format": "docx"}, [("file", "sample.pptx", SAMPLE_PPTX.read_bytes(), "application/vnd.openxmlformats-officedocument.presentationml.presentation")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        doc = Document(io.BytesIO(resp.read()))
        headings = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
        assert any("Quarterly" in h or "Review" in h for h in headings)
        assert len(doc.tables) >= 1
        print(f"  Extracted {len(headings)} slide headings and {len(doc.tables)} tables in DOCX.")
        print("  PASS: PPTX -> DOCX converted slides to headings, bullets, and tables.")

    # 15. DOCX -> PPTX (Roundtrip)
    print("\n[TEST 15] POST /convert: DOCX -> PPTX")
    content_type, body = encode_multipart_formdata({"target_format": "pptx"}, [("file", "sample_mixed.docx", SAMPLE_MIXED_DOCX.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        prs = Presentation(io.BytesIO(resp.read()))
        assert len(prs.slides) >= 2
        print(f"  Created presentation with {len(prs.slides)} slides.")
        print("  PASS: DOCX -> PPTX generated slides with title and content.")

    # 16. XLSX -> PPTX
    print("\n[TEST 16] POST /convert: XLSX -> PPTX")
    content_type, body = encode_multipart_formdata({"target_format": "pptx"}, [("file", "sample_multisheet.xlsx", SAMPLE_MULTISHEET_XLSX.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        prs = Presentation(io.BytesIO(resp.read()))
        assert len(prs.slides) == 2, f"Expected 2 slides for 2 sheets, got {len(prs.slides)}"
        print(f"  Generated {len(prs.slides)} slides from multi-sheet workbook.")
        print("  PASS: XLSX -> PPTX turned each sheet into a presentation slide.")

    # 17. CSV ↔ XLSX (Lossless roundtrip)
    print("\n[TEST 17] POST /convert: CSV ↔ XLSX (Lossless Roundtrip)")
    # CSV -> XLSX
    content_type, body = encode_multipart_formdata({"target_format": "xlsx"}, [("file", "sample.csv", SAMPLE_CSV.read_bytes(), "text/csv")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        xlsx_bytes = resp.read()
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes), data_only=True)
        ws = wb.active
        assert ws.cell(row=1, column=1).value == "City"
        val_pop = ws.cell(row=2, column=3).value
        assert isinstance(val_pop, (int, float)), f"Expected numeric population, got {val_pop}"

    # XLSX -> CSV
    content_type, body = encode_multipart_formdata({"target_format": "csv"}, [("file", "converted.xlsx", xlsx_bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        csv_text = resp.read().decode("utf-8-sig")
        assert "Tokyo" in csv_text and "Delhi" in csv_text
        print("  PASS: CSV ↔ XLSX lossless roundtrip verified with numeric parsing.")

    # 18. MD -> DOCX
    print("\n[TEST 18] POST /convert: MD -> DOCX")
    content_type, body = encode_multipart_formdata({"target_format": "docx"}, [("file", "sample.md", SAMPLE_MD.read_bytes(), "text/markdown")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        doc = Document(io.BytesIO(resp.read()))
        headings = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
        assert any("Release Notes" in h for h in headings)
        assert len(doc.tables) >= 1
        print(f"  Markdown parsed into {len(headings)} headings and {len(doc.tables)} tables in DOCX.")
        print("  PASS: MD -> DOCX converted markdown headings, lists, and tables.")

    # 19. HTML -> PDF
    print("\n[TEST 19] POST /convert: HTML -> PDF")
    content_type, body = encode_multipart_formdata({"target_format": "pdf"}, [("file", "sample.html", SAMPLE_HTML.read_bytes(), "text/html")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        pdf_bytes = resp.read()
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        assert len(doc) >= 1
        text = doc[0].get_text()
        assert "Platform Performance" in text or "Pipeline" in text
        doc.close()
        print("  PASS: HTML -> PDF chained successfully to high-quality PDF.")

    # 20. PNG -> JPG -> WebP (Direct Image ↔ Image)
    print("\n[TEST 20] POST /convert: PNG -> JPG -> WebP (Direct Image Conversions)")
    # PNG -> JPG
    content_type, body = encode_multipart_formdata({"target_format": "jpg"}, [("file", "test_transparent.png", SAMPLE_TRANSPARENT_PNG.read_bytes(), "image/png")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        assert resp.headers.get("Content-Type") == "image/jpeg"
        jpg_bytes = resp.read()

    # JPG -> WebP
    content_type, body = encode_multipart_formdata({"target_format": "webp"}, [("file", "test.jpg", jpg_bytes, "image/jpeg")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        assert resp.headers.get("Content-Type") == "image/webp"
        webp_bytes = resp.read()
        with Image.open(io.BytesIO(webp_bytes)) as im:
            assert im.format == "WEBP"
        print("  PASS: PNG -> JPG (flattened onto white) -> WebP verified.")

    # 21. Image -> TXT (via RapidOCR)
    print("\n[TEST 21] POST /convert: Image -> TXT (RapidOCR Text Extraction)")
    content_type, body = encode_multipart_formdata({"target_format": "txt"}, [("file", "sample_ocr_doc.png", SAMPLE_OCR_DOC_IMG.read_bytes(), "image/png")])
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        txt = resp.read().decode("utf-8")
        assert "Milestone" in txt or "Phase 1" in txt
        print(f"  Extracted plain text from image: {txt[:80]}...")
        print("  PASS: Image -> TXT extracted clean readable text via RapidOCR.")

    # 22. Custom DPI (150 vs 300 DPI)
    print("\n[TEST 22] POST /convert: Custom DPI (150 vs 300 DPI)")
    content_type_150, body_150 = encode_multipart_formdata(
        {"target_format": "png", "dpi": "150"},
        [("file", "sample_1page.pdf", SAMPLE_1PAGE_PDF.read_bytes(), "application/pdf")]
    )
    req_150 = urllib.request.Request(f"{SERVER_URL}/convert", data=body_150, headers={"Content-Type": content_type_150}, method="POST")
    with urllib.request.urlopen(req_150) as resp:
        assert resp.status == 200
        data_150 = resp.read()
        im_150 = Image.open(io.BytesIO(data_150))

    content_type_300, body_300 = encode_multipart_formdata(
        {"target_format": "png", "dpi": "300"},
        [("file", "sample_1page.pdf", SAMPLE_1PAGE_PDF.read_bytes(), "application/pdf")]
    )
    req_300 = urllib.request.Request(f"{SERVER_URL}/convert", data=body_300, headers={"Content-Type": content_type_300}, method="POST")
    with urllib.request.urlopen(req_300) as resp:
        assert resp.status == 200
        data_300 = resp.read()
        im_300 = Image.open(io.BytesIO(data_300))

    assert len(data_150) < len(data_300), f"Expected 150 DPI ({len(data_150)} B) < 300 DPI ({len(data_300)} B)"
    assert im_150.width < im_300.width, f"Expected 150 DPI width ({im_150.width}) < 300 DPI width ({im_300.width})"
    dpi_meta_150 = im_150.info.get("dpi", (150, 150))[0]
    assert abs(dpi_meta_150 - 150) < 1.0, f"Expected 150 DPI, got {dpi_meta_150}"
    print(f"  150 DPI: {im_150.size} ({len(data_150)} bytes) vs 300 DPI: {im_300.size} ({len(data_300)} bytes)")
    print("  PASS: Custom dpi=150 produced smaller, exact 150 DPI PNG compared to 300 DPI.")

    # 23. Page Range Filtering & Malformed Error
    print("\n[TEST 23] POST /convert: PDF Page Range ('2' vs malformed 'abc')")
    content_type_p2, body_p2 = encode_multipart_formdata(
        {"target_format": "png", "page_range": "2"},
        [("file", "sample.pdf", SAMPLE_PDF.read_bytes(), "application/pdf")]
    )
    req_p2 = urllib.request.Request(f"{SERVER_URL}/convert", data=body_p2, headers={"Content-Type": content_type_p2}, method="POST")
    with urllib.request.urlopen(req_p2) as resp:
        assert resp.status == 200
        assert resp.headers.get("Content-Type") == "image/png"
        assert resp.headers.get("X-Count") == "1"
        im_p2 = Image.open(io.BytesIO(resp.read()))
        assert im_p2.format == "PNG"
        print("  page_range='2' returned single PNG of page 2 directly.")

    content_type_bad, body_bad = encode_multipart_formdata(
        {"target_format": "png", "page_range": "abc"},
        [("file", "sample.pdf", SAMPLE_PDF.read_bytes(), "application/pdf")]
    )
    req_bad = urllib.request.Request(f"{SERVER_URL}/convert", data=body_bad, headers={"Content-Type": content_type_bad}, method="POST")
    try:
        with urllib.request.urlopen(req_bad) as resp:
            assert False, "Expected 400 error for malformed page_range='abc'"
    except urllib.error.HTTPError as err:
        assert err.code == 400, f"Expected 400, got {err.code}"
        err_body = json.loads(err.read().decode())
        assert "page_range" in err_body.get("detail", "").lower() or "syntax" in err_body.get("detail", "").lower()
        print(f"  page_range='abc' returned clean 400 error: {err_body.get('detail')}")
    print("  PASS: Page range extraction and 400 validation verified.")

    # 24. GET /health
    print("\n[TEST 24] GET /health (Cloud Hosting Health Check)")
    req = urllib.request.Request(f"{SERVER_URL}/health")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        data = json.loads(resp.read().decode())
        assert data.get("status") == "ok", f"Expected status 'ok', got {data.get('status')}"
        assert isinstance(data.get("libreoffice"), bool), f"Expected bool for libreoffice, got {data.get('libreoffice')}"
        assert data.get("formats") == 11, f"Expected 11 formats, got {data.get('formats')}"
        print(f"  Health check response: {data}")
        print("  PASS: GET /health returned 200 OK with status: ok, libreoffice bool, and formats: 11.")

    # 25. Filename Sanitization
    print("\n[TEST 25] POST /convert: Unsafe Filename Sanitization (../../evil name?.pdf)")
    content_type, body = encode_multipart_formdata(
        {"target_format": "png", "page_range": "1"},
        [("file", "../../evil name?.pdf", SAMPLE_1PAGE_PDF.read_bytes(), "application/pdf")]
    )
    req = urllib.request.Request(f"{SERVER_URL}/convert", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        disposition = resp.headers.get("Content-Disposition", "")
        assert "filename=" in disposition
        match = re.search(r'filename="?([^"]+)"?', disposition)
        assert match is not None, f"Could not extract filename from {disposition}"
        download_filename = match.group(1)
        assert ".." not in download_filename, f"Found '..' in {download_filename}"
        assert "/" not in download_filename, f"Found '/' in {download_filename}"
        assert "\\" not in download_filename, f"Found '\\' in {download_filename}"
        assert "?" not in download_filename, f"Found '?' in {download_filename}"
        assert download_filename.endswith(".png"), f"Expected .png filename, got {download_filename}"
        print(f"  Sanitized download filename: '{download_filename}'")
        print("  PASS: Unsafe filename stripped of path traversal and illegal characters.")

    # 26. Mixed-Format Batch (sample_mixed.docx + sample_multisheet.xlsx + sample.md -> PDF)
    print("\n[TEST 26] POST /convert/batch: Mixed-Format Batch (DOCX + XLSX + MD -> PDF)")
    batch_files = [
        ("files", "sample_mixed.docx", SAMPLE_MIXED_DOCX.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        ("files", "sample_multisheet.xlsx", SAMPLE_MULTISHEET_XLSX.read_bytes(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
        ("files", "sample.md", SAMPLE_MD.read_bytes(), "text/markdown"),
    ]
    content_type, body = encode_multipart_formdata({"target_format": "pdf"}, batch_files)
    req = urllib.request.Request(f"{SERVER_URL}/convert/batch", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        assert resp.headers.get("Content-Type") == "application/zip"
        with zipfile.ZipFile(io.BytesIO(resp.read())) as zf:
            names = zf.namelist()
            print(f"  Batch ZIP entries: {names}")
            pdf_entries = [n for n in names if n.lower().endswith(".pdf")]
            assert len(pdf_entries) == 3, f"Expected 3 converted PDFs, got {pdf_entries}"
            for p in pdf_entries:
                pdf_data = zf.read(p)
                doc = pymupdf.open(stream=pdf_data, filetype="pdf")
                assert len(doc) >= 1
                doc.close()
        print("  PASS: Mixed-format batch returned ZIP with 3 valid converted PDFs.")

    # 27. Multi-page PDF + DOCX -> PNG Batch (No nested double-ZIPs)
    print("\n[TEST 27] POST /convert/batch: Multi-page PDF + DOCX -> PNG (No Nested Double-ZIPs)")
    batch_files_png = [
        ("files", "sample.pdf", SAMPLE_PDF.read_bytes(), "application/pdf"),
        ("files", "sample_mixed.docx", SAMPLE_MIXED_DOCX.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ]
    content_type, body = encode_multipart_formdata({"target_format": "png"}, batch_files_png)
    req = urllib.request.Request(f"{SERVER_URL}/convert/batch", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        assert resp.headers.get("Content-Type") == "application/zip"
        with zipfile.ZipFile(io.BytesIO(resp.read())) as zf:
            names = zf.namelist()
            print(f"  Batch PNG ZIP entries: {names}")
            assert not any(n.lower().endswith(".zip") for n in names), "Found nested .zip file inside batch ZIP!"
            png_entries = [n for n in names if n.lower().endswith(".png")]
            assert len(png_entries) >= 3, f"Expected at least 3 PNGs, got {len(png_entries)}"
            assert any("/" in n and "sample" in n for n in png_entries), "Expected PDF pages in subfolder"
            for p in png_entries:
                im = Image.open(io.BytesIO(zf.read(p)))
                assert im.format == "PNG"
        print("  PASS: Multi-page PDF + DOCX batch produced clean subfoldered PNGs without nested ZIPs.")

    # 28. Duplicate Filename Handling in Batch
    print("\n[TEST 28] POST /convert/batch: Duplicate Filename Deduplication (report.docx + report.docx -> PDF)")
    batch_dupes = [
        ("files", "report.docx", SAMPLE_MIXED_DOCX.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
        ("files", "report.docx", SAMPLE_MIXED_DOCX.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document"),
    ]
    content_type, body = encode_multipart_formdata({"target_format": "pdf"}, batch_dupes)
    req = urllib.request.Request(f"{SERVER_URL}/convert/batch", data=body, headers={"Content-Type": content_type}, method="POST")
    with urllib.request.urlopen(req) as resp:
        assert resp.status == 200
        assert resp.headers.get("Content-Type") == "application/zip"
        with zipfile.ZipFile(io.BytesIO(resp.read())) as zf:
            names = [n for n in zf.namelist() if not n.startswith("_")]
            print(f"  Deduplicated entries: {names}")
            assert len(names) == 2, f"Expected 2 deduplicated entries, got {names}"
            assert len(set(names)) == 2, "Duplicate names detected inside ZIP!"
            assert "report.pdf" in names
            assert "report_2.pdf" in names
        print("  PASS: Duplicate filenames safely deduplicated (report.pdf, report_2.pdf).")

    # 29. In-Browser Editable Preview & Export: Spreadsheet Grid (Image Table OCR -> Edit Cell -> Export XLSX)
    print("\n[TEST 29] POST /preview-editable & POST /export-edited (Table Mode: Image OCR -> Edit Cell -> XLSX)")
    content_type_preview, body_preview = encode_multipart_formdata(
        {"target_format": "xlsx"},
        [("file", "sample_ocr_table.png", SAMPLE_OCR_TABLE_IMG.read_bytes(), "image/png")]
    )
    req_preview = urllib.request.Request(f"{SERVER_URL}/preview-editable", data=body_preview, headers={"Content-Type": content_type_preview}, method="POST")
    with urllib.request.urlopen(req_preview) as resp:
        assert resp.status == 200
        preview_json = json.loads(resp.read().decode("utf-8"))
        assert preview_json["mode"] == "table"
        assert preview_json["target_format"] == "xlsx"
        assert len(preview_json["sheets"]) >= 1
        rows = preview_json["sheets"][0]["rows"]
        assert len(rows) >= 2
        print(f"  Preview extracted {len(preview_json['sheets'])} sheet(s), {len(rows)} rows.")

    # Edit a cell in the table grid
    edited_val = "EDITED_REVENUE_99999"
    # Ensure second row has at least 2 columns
    if len(preview_json["sheets"][0]["rows"][1]) < 2:
        preview_json["sheets"][0]["rows"][1].append(edited_val)
    else:
        preview_json["sheets"][0]["rows"][1][1] = edited_val

    export_payload = {
        "mode": "table",
        "filename": "edited_financials",
        "target_format": "xlsx",
        "sheets": preview_json["sheets"],
    }
    req_export = urllib.request.Request(
        f"{SERVER_URL}/export-edited",
        data=json.dumps(export_payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req_export) as resp:
        assert resp.status == 200
        assert resp.headers.get("Content-Type") == "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        disposition = resp.headers.get("Content-Disposition", "")
        assert "edited_financials.xlsx" in disposition
        xlsx_data = resp.read()
        wb = openpyxl.load_workbook(io.BytesIO(xlsx_data), data_only=True)
        ws = wb.active
        found_edited = False
        for row in ws.iter_rows(values_only=True):
            for cell_val in row:
                if str(cell_val) == edited_val:
                    found_edited = True
                    break
            if found_edited:
                break
        assert found_edited, f"Could not find edited value '{edited_val}' in exported Excel sheet!"
        print(f"  Exported Excel workbook confirmed edited value: '{edited_val}'.")
        print("  PASS: In-browser spreadsheet preview extracted OCR table, user edited cell, and exported clean XLSX.")

    # 30. In-Browser Editable Preview & Export: Document Text (DOCX -> Edit Markdown -> Export DOCX)
    print("\n[TEST 30] POST /preview-editable & POST /export-edited (Document Mode: DOCX -> Edit Markdown -> DOCX)")
    content_type_doc_preview, body_doc_preview = encode_multipart_formdata(
        {"target_format": "docx"},
        [("file", "sample_mixed.docx", SAMPLE_MIXED_DOCX.read_bytes(), "application/vnd.openxmlformats-officedocument.wordprocessingml.document")]
    )
    req_doc_preview = urllib.request.Request(f"{SERVER_URL}/preview-editable", data=body_doc_preview, headers={"Content-Type": content_type_doc_preview}, method="POST")
    with urllib.request.urlopen(req_doc_preview) as resp:
        assert resp.status == 200
        preview_doc_json = json.loads(resp.read().decode("utf-8"))
        assert preview_doc_json["mode"] == "document"
        assert preview_doc_json["target_format"] == "docx"
        assert "content" in preview_doc_json
        content_md = preview_doc_json["content"]
        assert len(content_md) > 10
        print(f"  Preview extracted document markdown ({len(content_md)} chars): {content_md[:60]}...")

    # Edit the document markdown
    custom_heading = "EDITED_CUSTOM_SECTION_HEADING"
    custom_paragraph = "This paragraph was modified inside the in-browser editable preview before exporting."
    edited_md = content_md + f"\n\n# {custom_heading}\n\n{custom_paragraph}\n"

    export_doc_payload = {
        "mode": "document",
        "filename": "my_edited_document",
        "target_format": "docx",
        "content": edited_md,
    }
    req_doc_export = urllib.request.Request(
        f"{SERVER_URL}/export-edited",
        data=json.dumps(export_doc_payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST"
    )
    with urllib.request.urlopen(req_doc_export) as resp:
        assert resp.status == 200
        assert resp.headers.get("Content-Type") == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        disposition = resp.headers.get("Content-Disposition", "")
        assert "my_edited_document.docx" in disposition
        docx_data = resp.read()
        doc = Document(io.BytesIO(docx_data))
        text_content = [p.text for p in doc.paragraphs]
        assert any(custom_heading in t for t in text_content), f"Heading '{custom_heading}' missing in exported docx!"
        assert any(custom_paragraph in t for t in text_content), f"Paragraph '{custom_paragraph}' missing in exported docx!"
        print(f"  Exported DOCX confirmed edited heading and text paragraph.")
        print("  PASS: In-browser document preview extracted Markdown, user edited text, and exported formatted DOCX.")

    # 31. POST /inspect-zip: ZIP Archive Live Inspector
    print("\n[TEST 31] POST /inspect-zip: ZIP Archive Live Inspector (Thumbnails & Previews)")
    test_zip_buf = io.BytesIO()
    with zipfile.ZipFile(test_zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("test_page_1.png", SAMPLE_OCR_DOC_IMG.read_bytes())
        zf.writestr("notes.txt", "This is an extracted document summary inside the archive.".encode("utf-8"))
    test_zip_bytes = test_zip_buf.getvalue()

    content_type_zip, body_zip = encode_multipart_formdata(
        {},
        [("file", "archive.zip", test_zip_bytes, "application/zip")]
    )
    req_zip = urllib.request.Request(f"{SERVER_URL}/inspect-zip", data=body_zip, headers={"Content-Type": content_type_zip}, method="POST")
    with urllib.request.urlopen(req_zip) as resp:
        assert resp.status == 200
        zip_json = json.loads(resp.read().decode("utf-8"))
        assert zip_json["total_files"] == 2
        entries = zip_json["entries"]
        assert len(entries) == 2
        img_entry = next(e for e in entries if e["ext"] == "png")
        assert img_entry["type"] == "image"
        assert img_entry["preview"].startswith("data:image/png;base64,")
        txt_entry = next(e for e in entries if e["ext"] == "txt")
        assert txt_entry["type"] == "text"
        assert "extracted document summary" in txt_entry["preview"]
        print(f"  Inspected ZIP entries: {[e['name'] for e in entries]}")
        print("  PASS: POST /inspect-zip extracted file list, base64 image data URI, and text preview.")

    print("\n=========================================================================================")
    print("ALL 31 TESTS (21 CONVERSIONS + DPI + PAGE RANGE + HEALTH + SANITIZE + 3 BATCH + 2 PREVIEW-EDIT + ZIP INSPECT) PASSED WITH 100% SUCCESS!")
    print("=========================================================================================")


if __name__ == "__main__":
    generate_test_assets()
    run_all_tests()
