from pathlib import Path
import pymupdf

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_PDF_PATH = BASE_DIR / "sample.pdf"
SAMPLE_1PAGE_PDF_PATH = BASE_DIR / "sample_1page.pdf"

def generate_samples():
    print(f"Creating 2-page sample PDF at {SAMPLE_PDF_PATH}...")
    doc = pymupdf.open()

    # Page 1
    page1 = doc.new_page(width=595, height=842)  # A4
    page1.draw_rect(pymupdf.Rect(40, 40, 555, 802), color=(0.2, 0.4, 0.8), width=3)
    page1.insert_text((60, 110), "Velmo Test Document", fontsize=24, color=(0.1, 0.1, 0.2))
    page1.insert_text((60, 150), "Page 1 of 2 - 300 DPI Rendering Test", fontsize=15, color=(0.2, 0.4, 0.8))
    page1.insert_text((60, 190), "This sample document verifies multi-page PDF conversion into a ZIP archive.", fontsize=11, color=(0.4, 0.4, 0.4))
    page1.draw_circle(pymupdf.Point(297, 450), 120, color=(0.1, 0.7, 0.4), fill=(0.92, 0.98, 0.94))

    # Page 2
    page2 = doc.new_page(width=595, height=842)  # A4
    page2.draw_rect(pymupdf.Rect(40, 40, 555, 802), color=(0.8, 0.3, 0.2), width=3)
    page2.insert_text((60, 110), "Velmo Test Document", fontsize=24, color=(0.1, 0.1, 0.2))
    page2.insert_text((60, 150), "Page 2 of 2 - Second Page Verification", fontsize=15, color=(0.8, 0.3, 0.2))
    page2.insert_text((60, 190), "Rendered as page_02.png inside the returned ZIP package.", fontsize=11, color=(0.4, 0.4, 0.4))
    page2.draw_rect(pymupdf.Rect(180, 360, 415, 540), color=(0.9, 0.5, 0.1), fill=(1.0, 0.96, 0.88))

    doc.save(SAMPLE_PDF_PATH)
    doc.close()
    print(f"Successfully generated 2-page sample PDF: {SAMPLE_PDF_PATH.name} ({SAMPLE_PDF_PATH.stat().st_size} bytes)")

    # 1-Page PDF
    doc1 = pymupdf.open()
    p1 = doc1.new_page(width=595, height=842)
    p1.insert_text((60, 110), "Single Page Sample PDF", fontsize=24, color=(0.1, 0.1, 0.2))
    p1.insert_text((60, 150), "Verifies single-page PDF conversion directly to a single .png file.", fontsize=13, color=(0.3, 0.3, 0.3))
    doc1.save(SAMPLE_1PAGE_PDF_PATH)
    doc1.close()
    print(f"Successfully generated 1-page sample PDF: {SAMPLE_1PAGE_PDF_PATH.name} ({SAMPLE_1PAGE_PDF_PATH.stat().st_size} bytes)")

if __name__ == "__main__":
    generate_samples()
