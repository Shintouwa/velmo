import io
import json
import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional
from xml.sax.saxutils import escape as xml_escape

from PIL import Image, ImageDraw, ImageFont


@dataclass
class SeoRoute:
    slug: str
    from_fmt: str
    to_fmt: str
    title: str
    description: str
    h1: str
    subtitle: str
    badge: str
    tech_overview: str
    how_to_steps: List[str]
    faqs: List[Dict[str, str]]
    reverse_slug: Optional[str]
    related_slugs: List[str] = field(default_factory=list)


FAVICON_SVG = """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 32 32">
  <defs>
    <linearGradient id="g" x1="0%" y1="0%" x2="100%" y2="100%">
      <stop offset="0%" stop-color="#FF4D1C"/>
      <stop offset="100%" stop-color="#E53E10"/>
    </linearGradient>
  </defs>
  <rect width="32" height="32" rx="8" fill="#09090B"/>
  <rect x="0.5" y="0.5" width="31" height="31" rx="7.5" fill="none" stroke="rgba(255,255,255,0.12)" stroke-width="1"/>
  <path d="M8 9l8 14 8-14h-4.5l-3.5 6.5L12.5 9z" fill="url(#g)"/>
</svg>"""

SEO_ROUTES: Dict[str, SeoRoute] = {
    "pdf-to-word": SeoRoute(
        slug="pdf-to-word",
        from_fmt="pdf",
        to_fmt="docx",
        title="PDF to Word Converter (Free & Editable DOCX) — Velmo",
        description="Convert PDF documents to editable Word (.docx) files online for free. Preserves layout, tables, fonts, and inline images with zero signup required.",
        h1="Convert PDF to Editable Word Document",
        subtitle="Transform fixed-layout PDF files into fully editable Microsoft Word documents while preserving original formatting, tables, and typography.",
        badge="Layout-Preserving Document Reconstruction",
        tech_overview="Our document reconstruction pipeline analyzes PDF text streams, bounding boxes, and table lines. Text blocks are synthesized into semantic paragraphs and headings, preserving native font styling and table cells for clean editing in Word.",
        how_to_steps=[
            "Drop your PDF file into the upload zone above or browse your computer.",
            "Velmo automatically pre-configures Word (.docx) as the active target format.",
            "Click 'Convert Document' to process text blocks, tables, and images in seconds.",
            "Download your editable .docx file or preview and refine extracted text directly in your browser."
        ],
        faqs=[
            {
                "q": "Will my tables and column layouts remain intact in Word?",
                "a": "Yes. Our layout reconstruction engine detects table borders and alignment grids, generating native Word table cells rather than loose floating text."
            },
            {
                "q": "Are scanned PDF documents supported?",
                "a": "Scanned documents containing raster images are processed through our local neural OCR pipeline to extract readable text into your Word document."
            },
            {
                "q": "Is there any file size limit or account requirement?",
                "a": "No account or registration is required. Conversions are completely free, unlimited, and processed in ephemeral memory without saving your files."
            }
        ],
        reverse_slug="word-to-pdf",
        related_slugs=["pdf-to-excel", "pdf-to-png", "pdf-to-txt", "image-to-word"]
    ),
    "word-to-pdf": SeoRoute(
        slug="word-to-pdf",
        from_fmt="docx",
        to_fmt="pdf",
        title="Word to PDF Converter (Print-Ready PDF) — Velmo",
        description="Convert Microsoft Word (.docx) documents to print-ready PDF files instantly. Preserves exact margins, headings, tables, and vector quality online.",
        h1="Convert Word (.docx) to Print-Ready PDF",
        subtitle="Export Microsoft Word documents to universally compatible, fixed-layout PDF files with vector typography and clean pagination.",
        badge="High-Precision PDF Typesetting",
        tech_overview="The document rendering engine converts DOCX paragraphs, run properties, bullet structures, and table grids into exact PDF vector streams. Margins, line heights, and table borders are rendered with high geometric precision.",
        how_to_steps=[
            "Upload your Microsoft Word document (.docx) to the converter.",
            "Target format is automatically set to PDF (.pdf).",
            "Click 'Convert Document' to generate a standardized, print-ready PDF.",
            "Review your document in our integrated multi-page PDF viewer and download instantly."
        ],
        faqs=[
            {
                "q": "Can I view the generated PDF before downloading?",
                "a": "Yes. Velmo includes an integrated page-by-page document viewer that lets you review every page directly inside your browser."
            },
            {
                "q": "Does the conversion preserve embedded images and custom fonts?",
                "a": "Yes. Embedded high-resolution images and document headings are rendered cleanly without lossy re-compression or layout shifting."
            },
            {
                "q": "Is my confidential document stored on your servers?",
                "a": "No. All conversions happen entirely in ephemeral server memory and are wiped immediately after your session ends."
            }
        ],
        reverse_slug="pdf-to-word",
        related_slugs=["excel-to-pdf", "pptx-to-pdf", "markdown-to-pdf", "html-to-pdf"]
    ),
    "pdf-to-excel": SeoRoute(
        slug="pdf-to-excel",
        from_fmt="pdf",
        to_fmt="xlsx",
        title="PDF to Excel Converter (Extract Tables to XLSX) — Velmo",
        description="Extract tables from PDF documents into editable Excel (.xlsx) spreadsheets. Fast, accurate data extraction with zero cloud storage or registration.",
        h1="Extract PDF Tables to Excel Spreadsheet",
        subtitle="Detect and convert financial reports, statements, and tabular data from PDF pages into multi-column Excel (.xlsx) spreadsheets.",
        badge="Geometric Table Grid Extraction",
        tech_overview="Our tabular extraction engine scans PDF vector coordinate streams for horizontal and vertical ruling lines. Intersecting points form tabular cell matrices, which are converted into native Excel sheets with automatic numeric data typing.",
        how_to_steps=[
            "Drag and drop your PDF containing statements or tables into the studio.",
            "Excel (.xlsx) is automatically selected as your target format.",
            "Click 'Convert Document' to extract tabular data into rows and columns.",
            "Download your Excel workbook or click 'Inspect / Edit' to modify cell values in our in-browser grid."
        ],
        faqs=[
            {
                "q": "Are numerical values converted into actual numbers or text strings?",
                "a": "Our data engine automatically casts numeric entries, currencies, and percentages into native numeric types so formulas and calculations work immediately."
            },
            {
                "q": "Can I edit the extracted spreadsheet before downloading?",
                "a": "Yes. Velmo includes an interactive in-browser spreadsheet editor allowing you to add rows, edit cell contents, and preview before saving."
            },
            {
                "q": "How does the extractor handle multi-page tables?",
                "a": "Continuous tables spanning multiple PDF pages are recognized and organized into clean sequential sheets or structured workbooks."
            }
        ],
        reverse_slug="excel-to-pdf",
        related_slugs=["pdf-to-word", "image-to-excel", "csv-to-excel", "word-to-excel"]
    ),
    "excel-to-pdf": SeoRoute(
        slug="excel-to-pdf",
        from_fmt="xlsx",
        to_fmt="pdf",
        title="Excel to PDF Converter (Auto-Fit Sheets to PDF) — Velmo",
        description="Convert Excel (.xlsx) spreadsheets to clean, professional PDF documents. Features auto-landscape fitting and multi-sheet page breaks for clean print.",
        h1="Convert Excel Spreadsheets to Professional PDF",
        subtitle="Transform wide data tables and multi-sheet workbooks into beautifully formatted, printable PDF documents with automatic orientation.",
        badge="Dynamic Tabular Layout Engine",
        tech_overview="Velmo evaluates spreadsheet column counts and content dimensions dynamically. Wide sheets are automatically laid out in landscape orientation with proportional cell wrapping, grid lines, and styled header rows.",
        how_to_steps=[
            "Upload your Excel spreadsheet (.xlsx) into the converter.",
            "Target format is pre-configured to PDF.",
            "Click 'Convert Document' to render formatted pages with auto-fitted columns.",
            "Inspect the layout using our crisp page previewer and save your PDF."
        ],
        faqs=[
            {
                "q": "How are wide spreadsheets prevented from being cut off?",
                "a": "The rendering engine automatically toggles landscape mode and calculates proportional column widths so all data fits comfortably on the page."
            },
            {
                "q": "Are multiple sheets included in the output PDF?",
                "a": "Yes. Each worksheet in your Excel workbook is rendered sequentially with clear section dividers and header formatting."
            },
            {
                "q": "Can I edit the spreadsheet data before creating the PDF?",
                "a": "Yes! Click 'Preview & Edit' to open our live spreadsheet editor, adjust cell values, and export directly to PDF."
            }
        ],
        reverse_slug="pdf-to-excel",
        related_slugs=["word-to-pdf", "csv-to-pdf", "excel-to-word", "excel-to-csv"]
    ),
    "image-to-excel": SeoRoute(
        slug="image-to-excel",
        from_fmt="png",
        to_fmt="xlsx",
        title="Image to Excel Converter (Free OCR Table Extraction) — Velmo",
        description="Convert PNG, JPG, or WebP images to editable Excel (.xlsx) spreadsheets. Fast neural OCR table clustering with private on-device memory processing.",
        h1="Convert Image to Excel Spreadsheet with OCR",
        subtitle="Extract tables, invoices, receipts, and photos of documents into structured, editable Excel (.xlsx) spreadsheets via local neural OCR.",
        badge="Neural OCR Table Clustering Engine",
        tech_overview="Our image-to-spreadsheet pipeline runs an on-device neural character recognition model. Text boxes are mapped into a 2D spatial coordinate plane. An intelligent clustering algorithm identifies row baselines and column boundaries, outputting structured Excel tables.",
        how_to_steps=[
            "Upload a PNG, JPG, or WebP image containing a table, bill, or balance sheet.",
            "Velmo automatically activates the neural OCR engine and selects Excel (.xlsx).",
            "Click 'Convert Document' to extract cell rows and columns in seconds.",
            "Download your .xlsx workbook or use the in-browser table editor to fine-tune data immediately."
        ],
        faqs=[
            {
                "q": "Can it extract tables from photos taken on a phone?",
                "a": "Yes. Clear, well-lit smartphone photos and document scans are analyzed by our neural OCR engine to reconstruct clean table rows and columns."
            },
            {
                "q": "Are my financial photos or receipts sent to cloud AI services?",
                "a": "No. The neural OCR engine runs locally in ephemeral server memory. No image data is ever logged, retained, or transmitted to third parties."
            },
            {
                "q": "Can I combine multiple receipt or table images into one workbook?",
                "a": "Yes. You can upload multiple images simultaneously and combine their extracted tables into a single organized spreadsheet."
            }
        ],
        reverse_slug="excel-to-pdf",
        related_slugs=["image-to-word", "pdf-to-excel", "csv-to-excel", "image-to-pdf"]
    ),
    "image-to-word": SeoRoute(
        slug="image-to-word",
        from_fmt="png",
        to_fmt="docx",
        title="Image to Word Converter (OCR Photo to DOCX) — Velmo",
        description="Extract text and headings from images into editable Word (.docx) documents. Free neural OCR document conversion with zero sign-up or file retention.",
        h1="Convert Image to Editable Word Document with OCR",
        subtitle="Extract readable, editable paragraphs, titles, and lists from scanned documents and photos directly into Microsoft Word (.docx).",
        badge="Neural Character Recognition Engine",
        tech_overview="Visual features and text lines are detected across image buffers. Spatial grouping algorithms categorize detected text runs into hierarchical titles, body paragraphs, and itemized lists, generating native Word styles.",
        how_to_steps=[
            "Drop your document image (PNG, JPG, or WebP) into the upload area.",
            "Target format is automatically pre-selected to Word (.docx).",
            "Click 'Convert Document' to process visual text into editable copy.",
            "Download your formatted Word document or review and edit the text in our live editor."
        ],
        faqs=[
            {
                "q": "Does this work on screenshots of articles and code snippets?",
                "a": "Yes. High-resolution screenshots of text, articles, and book pages are accurately transcribed into editable Word paragraphs."
            },
            {
                "q": "Can I edit the OCR output before downloading?",
                "a": "Yes. Click 'Preview & Edit' to open our live in-browser editor where you can correct text, format headings, and export directly."
            },
            {
                "q": "What languages are supported by the OCR engine?",
                "a": "The neural engine supports standard Latin-based scripts including English, Spanish, French, German, Italian, and common alphanumeric text."
            }
        ],
        reverse_slug="word-to-pdf",
        related_slugs=["image-to-excel", "pdf-to-word", "image-to-pdf", "word-to-txt"]
    ),
    "image-to-pdf": SeoRoute(
        slug="image-to-pdf",
        from_fmt="png",
        to_fmt="pdf",
        title="Image to PDF Converter (Combine Images to PDF) — Velmo",
        description="Combine PNG, JPG, and WebP images into a single clean PDF document. Lossless image embedding with custom page orientation and zero compression loss.",
        h1="Convert and Combine Images to PDF",
        subtitle="Merge multiple photos, scans, and screenshots into a single, cohesive PDF document with pristine original image quality.",
        badge="Lossless Stream PDF Embedding",
        tech_overview="Image byte streams are mapped directly into PDF vector container objects without lossy re-encoding or color degradation. Page dimensions adapt automatically to match the source aspect ratio of each uploaded image.",
        how_to_steps=[
            "Select one or more PNG, JPG, or WebP images from your device.",
            "Choose 'Combine into 1 file' or 'Keep separate' in the batch controls.",
            "Click 'Convert Document' to assemble your images into a structured PDF.",
            "Inspect the final document in our page viewer and download instantly."
        ],
        faqs=[
            {
                "q": "Can I reorder images before generating the PDF?",
                "a": "Yes! Our queued files list includes intuitive up/down reorder buttons so you can arrange your pages in exact sequence."
            },
            {
                "q": "Does converting reduce the resolution or clarity of my photos?",
                "a": "No. Images are embedded directly into the PDF structure at their native resolution without unwanted compression artifacts."
            },
            {
                "q": "How many images can I merge at once?",
                "a": "You can merge dozens of images in a single batch. All processing completes in seconds inside server memory."
            }
        ],
        reverse_slug="pdf-to-png",
        related_slugs=["pdf-to-jpg", "image-to-excel", "image-to-word", "batch-converter"]
    ),
    "pdf-to-png": SeoRoute(
        slug="pdf-to-png",
        from_fmt="pdf",
        to_fmt="png",
        title="PDF to PNG Converter (300 DPI High-Res Image) — Velmo",
        description="Convert PDF pages into high-resolution 300 DPI PNG images. Crisp vector rendering with transparent background support and instant ZIP archive download.",
        h1="Convert PDF Pages to High-Resolution PNG Images",
        subtitle="Render every page of your PDF into crisp, lossless PNG graphics at 150, 300, or 600 DPI for presentations, print, and web design.",
        badge="High-DPI Vector Rasterization",
        tech_overview="The rasterization engine renders PDF vector typography, shapes, and embedded graphics into pixel-perfect PNG bitmaps. Supports configurable resolution (150 Web, 300 Print, 600 Ultra) with anti-aliased font rendering.",
        how_to_steps=[
            "Upload your PDF document to the studio dropzone.",
            "Select PNG as the target format and choose your desired DPI resolution (default 300 DPI).",
            "Click 'Convert Document' to rasterize every page into a crisp image.",
            "Download your images bundled neatly into a single ZIP archive or single image file."
        ],
        faqs=[
            {
                "q": "What resolution is recommended for print quality?",
                "a": "We recommend selecting the 300 DPI (Print) or 600 DPI (Ultra) pill in our smart conversion settings for razor-sharp physical printing."
            },
            {
                "q": "How are multi-page PDFs delivered?",
                "a": "Multi-page PDFs are automatically rendered page by page and bundled into a organized ZIP archive with sequential filenames."
            },
            {
                "q": "Can I convert only specific pages?",
                "a": "Yes. Use the 'PDF Page Range' input in the settings to specify pages, such as '1-3, 5'."
            }
        ],
        reverse_slug="image-to-pdf",
        related_slugs=["pdf-to-jpg", "pdf-to-word", "pdf-to-excel", "png-to-webp"]
    ),
    "pdf-to-jpg": SeoRoute(
        slug="pdf-to-jpg",
        from_fmt="pdf",
        to_fmt="jpg",
        title="PDF to JPG Converter (Fast & Compressed Images) — Velmo",
        description="Convert PDF pages to compressed JPEG images online. Select custom rendering DPI and image quality with fast multi-page batch conversion and zero login.",
        h1="Convert PDF Pages to Compressed JPG Images",
        subtitle="Export PDF documents into lightweight, universally compatible JPEG images with customizable compression quality and resolution.",
        badge="Optimized JPEG Rasterization Engine",
        tech_overview="Each PDF page is converted to an RGB bitmap buffer and encoded with fine-tuned DCT quantization matrices. Configurable quality sliders (60% to 100%) allow the ideal balance between small file size and visual fidelity.",
        how_to_steps=[
            "Drop your PDF document into the converter.",
            "Select JPG from the Images category.",
            "Adjust the quality slider (e.g. 95% default) and target DPI if needed.",
            "Click 'Convert Document' and download your packaged JPG images."
        ],
        faqs=[
            {
                "q": "How small can the JPG files be made?",
                "a": "You can slide the compression quality down to 70% or 60% and select 150 DPI to produce ultra-compact images ideal for email or mobile apps."
            },
            {
                "q": "Are white backgrounds guaranteed for PDF pages?",
                "a": "Yes. Transparent PDF canvas backgrounds are automatically composited onto clean solid white backgrounds without dark halos."
            },
            {
                "q": "Is there a limit on how many pages can be exported?",
                "a": "Velmo processes large multi-page documents quickly, packaging all generated JPG files into a clean ZIP archive."
            }
        ],
        reverse_slug="image-to-pdf",
        related_slugs=["pdf-to-png", "jpg-to-png", "pdf-to-word", "image-to-pdf"]
    ),
    "excel-to-word": SeoRoute(
        slug="excel-to-word",
        from_fmt="xlsx",
        to_fmt="docx",
        title="Excel to Word Converter (Spreadsheet to DOCX) — Velmo",
        description="Convert Excel (.xlsx) spreadsheet tables into formatted Word (.docx) reports. Preserves table structure, headers, and column widths with clean styling.",
        h1="Convert Excel Spreadsheets to Formatted Word Reports",
        subtitle="Turn numerical tables, financial summaries, and data sheets into clean, professional Word documents ready for executive reporting.",
        badge="Tabular Style Transformation",
        tech_overview="Spreadsheet worksheet cells, merge spans, header fills, and column ratios are mapped to native Word table properties. Numeric values are cleanly formatted, and multiple sheets are separated with clear headings.",
        how_to_steps=[
            "Upload your Excel file (.xlsx) containing tables or data.",
            "Target format is automatically set to Word (.docx).",
            "Click 'Convert Document' to generate a beautifully styled document report.",
            "Download your .docx file or use the live editor to draft surrounding commentary."
        ],
        faqs=[
            {
                "q": "Are multiple sheets preserved?",
                "a": "Yes. Each sheet is converted into a distinct section within the Word document, titled with the original worksheet name."
            },
            {
                "q": "Can I edit the table columns and rows once in Word?",
                "a": "Yes. Tables are created as native Word tables, allowing you to resize columns, add rows, and apply standard Word table styles."
            },
            {
                "q": "Does the conversion support merged cells?",
                "a": "Yes. Merged horizontal and vertical ranges from Excel are translated accurately into merged Word table cells."
            }
        ],
        reverse_slug="word-to-excel",
        related_slugs=["excel-to-pdf", "word-to-pdf", "csv-to-excel", "excel-to-csv"]
    ),
    "word-to-excel": SeoRoute(
        slug="word-to-excel",
        from_fmt="docx",
        to_fmt="xlsx",
        title="Word to Excel Converter (Extract Tables to XLSX) — Velmo",
        description="Extract structured tables from Word (.docx) documents into Excel (.xlsx) spreadsheets. Converts table cells into clean rows and columns automatically.",
        h1="Extract Tables from Word Documents to Excel",
        subtitle="Pull tables, data lists, and figures out of Microsoft Word (.docx) files and compile them into structured, formula-ready Excel spreadsheets.",
        badge="Document Table Parsing Engine",
        tech_overview="Our parser navigates Word document XML hierarchies to isolate table structures. Table rows and nested cells are flattened into structured tabular matrices, casting numeric strings into numbers for instant spreadsheet calculations.",
        how_to_steps=[
            "Drop your Word document (.docx) into the upload area.",
            "Select Excel (.xlsx) as your desired output format.",
            "Click 'Convert Document' to extract all embedded tables.",
            "Download your structured Excel file or inspect the data inside our in-browser grid."
        ],
        faqs=[
            {
                "q": "What happens if my Word document has multiple tables?",
                "a": "Each table detected in your Word document is placed on its own clearly labeled worksheet or cleanly stacked in sequential order."
            },
            {
                "q": "Does it extract regular paragraph text?",
                "a": "The converter focuses primarily on structured tables. If you want full text extraction, choose TXT or Markdown as the target."
            },
            {
                "q": "Can I modify the extracted cells before downloading?",
                "a": "Yes. Click 'Preview & Edit' to open our interactive grid editor, modify values or add formulas, and export directly."
            }
        ],
        reverse_slug="excel-to-word",
        related_slugs=["pdf-to-excel", "csv-to-excel", "word-to-pdf", "image-to-excel"]
    ),
    "csv-to-excel": SeoRoute(
        slug="csv-to-excel",
        from_fmt="csv",
        to_fmt="xlsx",
        title="CSV to Excel Converter (Formatted Multi-Sheet XLSX) — Velmo",
        description="Convert CSV data files into formatted multi-sheet Excel (.xlsx) workbooks. Automatically formats columns, detects delimiters, and parses numbers cleanly.",
        h1="Convert CSV Data Files to Formatted Excel Workbooks",
        subtitle="Upgrade raw comma-separated data into styled, multi-sheet Excel workbooks with automated delimiter detection and numeric typing.",
        badge="Smart Dialect Delimiter Engine",
        tech_overview="A dialect sniffing engine automatically detects comma, semicolon, tab, or pipe delimiters. Field values are typed as integers, floating point numbers, dates, or strings, creating an open-ready Excel workbook with styled header rows.",
        how_to_steps=[
            "Upload your CSV file (.csv) to the converter.",
            "Excel (.xlsx) is automatically selected as the target format.",
            "Click 'Convert Document' to transform raw records into a polished spreadsheet.",
            "Download your workbook or open it in our in-browser editor to add columns and rows."
        ],
        faqs=[
            {
                "q": "Does it support European semicolon-delimited CSVs?",
                "a": "Yes. The delimiter sniffer identifies commas, semicolons, tabs, and pipes automatically without requiring manual configuration."
            },
            {
                "q": "Will leading zeros in ZIP codes or phone numbers be preserved?",
                "a": "Yes. Fields formatted as alphanumeric identifiers with leading zeros are preserved as strings to prevent data truncation."
            },
            {
                "q": "Can I convert multiple CSV files at the same time?",
                "a": "Yes. Upload multiple CSV files together to have each file converted into individual workbooks or sheets in one quick action."
            }
        ],
        reverse_slug="excel-to-csv",
        related_slugs=["excel-to-pdf", "csv-to-pdf", "pdf-to-excel", "word-to-excel"]
    ),
    "excel-to-csv": SeoRoute(
        slug="excel-to-csv",
        from_fmt="xlsx",
        to_fmt="csv",
        title="Excel to CSV Converter (Export Sheets to Clean CSV) — Velmo",
        description="Export Excel (.xlsx) spreadsheets into clean, comma-separated CSV text files. Handles multi-sheet workbooks with instant ZIP packaging and zero tracking.",
        h1="Export Excel Workbooks to Clean CSV Text Files",
        subtitle="Convert complex Excel (.xlsx) spreadsheets into lightweight, universally compatible CSV files ready for database ingestion and data pipelines.",
        badge="High-Speed Delimiter Export Engine",
        tech_overview="Cells across active worksheets are extracted with formula evaluation and formatted as standard RFC-4180 compliant CSV lines with UTF-8 encoding. Multi-sheet workbooks are automatically exported as individual CSV files bundled in a ZIP.",
        how_to_steps=[
            "Upload your Excel workbook (.xlsx) to the studio.",
            "Select CSV (.csv) as your target format.",
            "Click 'Convert Document' to extract clean comma-delimited data.",
            "Download your CSV file or the bundled ZIP containing all exported worksheets."
        ],
        faqs=[
            {
                "q": "What happens if my workbook contains multiple sheets?",
                "a": "Every worksheet is exported as its own CSV file named after the sheet, bundled together into a clean, ready-to-download ZIP archive."
            },
            {
                "q": "Is the output encoded in UTF-8?",
                "a": "Yes. Output CSV files use standard UTF-8 encoding, ensuring special characters, accents, and non-English text remain intact."
            },
            {
                "q": "Are formulas calculated or exported as code?",
                "a": "Calculated cell values are exported directly so your database or analytics tools receive clean raw numbers."
            }
        ],
        reverse_slug="csv-to-excel",
        related_slugs=["excel-to-pdf", "excel-to-word", "csv-to-pdf", "word-to-excel"]
    ),
    "pptx-to-pdf": SeoRoute(
        slug="pptx-to-pdf",
        from_fmt="pptx",
        to_fmt="pdf",
        title="PowerPoint to PDF Converter (PPTX Slides to PDF) — Velmo",
        description="Convert PowerPoint (.pptx) slide decks to high-resolution PDF documents. Preserves 16:9 and 4:3 slide geometry, typography, and graphics for sharing.",
        h1="Convert PowerPoint (.pptx) Slides to PDF Document",
        subtitle="Turn presentation decks into fixed-layout, portable PDF files with sharp typography, preserved graphics, and consistent slide proportions.",
        badge="Presentation Vector Geometry Engine",
        tech_overview="Slide masters, shapes, text boxes, and background color fills are rendered to vector PDF pages. The conversion preserves widescreen 16:9 and traditional 4:3 slide dimensions, ensuring slides look identical across every device.",
        how_to_steps=[
            "Drop your PowerPoint presentation (.pptx) into the converter.",
            "Target format is automatically set to PDF (.pdf).",
            "Click 'Convert Document' to generate a fixed-layout slide document.",
            "Preview every slide using our integrated viewer and download your PDF."
        ],
        faqs=[
            {
                "q": "Are widescreen 16:9 slide proportions retained?",
                "a": "Yes. Slide aspect ratios are automatically calculated and matched in the output PDF pages so slides never look squished or stretched."
            },
            {
                "q": "Can recipients view the PDF without having PowerPoint installed?",
                "a": "Yes! PDFs open natively in any browser, mobile phone, or desktop operating system without requiring presentation software."
            },
            {
                "q": "Are speaker notes included?",
                "a": "By default, slides are rendered cleanly as presentation view. Slide notes are kept private and excluded from the public PDF output."
            }
        ],
        reverse_slug="pdf-to-pptx",
        related_slugs=["pdf-to-word", "word-to-pdf", "excel-to-pdf", "batch-converter"]
    ),
    "pdf-to-pptx": SeoRoute(
        slug="pdf-to-pptx",
        from_fmt="pdf",
        to_fmt="pptx",
        title="PDF to PowerPoint Converter (PDF Slides to PPTX) — Velmo",
        description="Convert PDF slides into editable PowerPoint (.pptx) presentations. Reconstructs slide layouts and text frames for easy editing and presentation mode.",
        h1="Convert PDF Slides to Editable PowerPoint Presentations",
        subtitle="Reconstruct slide presentations from static PDF documents into editable Microsoft PowerPoint (.pptx) decks with movable text boxes.",
        badge="Slide Frame Reconstruction Engine",
        tech_overview="The engine analyzes PDF page orientations and dimensions, mapping vector graphics and text clusters into native PowerPoint slide shapes. Text runs become editable text boxes, allowing you to revise bullet points and titles with ease.",
        how_to_steps=[
            "Upload your PDF slide deck or presentation document.",
            "Select PowerPoint (.pptx) from the Documents category.",
            "Click 'Convert Document' to transform static pages into presentation slides.",
            "Download your editable .pptx file and customize inside PowerPoint or Google Slides."
        ],
        faqs=[
            {
                "q": "Can I edit the text on the slides in PowerPoint?",
                "a": "Yes! Text is converted into native PowerPoint text boxes rather than flat screenshots, allowing full typography and wording edits."
            },
            {
                "q": "Does it work with Google Slides and Keynote?",
                "a": "Yes. The generated .pptx files conform to standard OpenXML presentation formats, compatible with Google Slides, Keynote, and LibreOffice."
            },
            {
                "q": "How does it handle high-resolution presentation graphics?",
                "a": "Embedded vector figures and images are extracted and placed as distinct slide elements to maintain maximum sharpness."
            }
        ],
        reverse_slug="pptx-to-pdf",
        related_slugs=["pdf-to-word", "pdf-to-png", "word-to-pdf", "image-to-pdf"]
    ),
    "png-to-jpg": SeoRoute(
        slug="png-to-jpg",
        from_fmt="png",
        to_fmt="jpg",
        title="PNG to JPG Converter (Transparent to Clean JPEG) — Velmo",
        description="Convert transparent PNG images to clean, compressed JPG files. Automatically renders clean white backgrounds without dark artifacts or quality loss.",
        h1="Convert Transparent PNG Images to Compressed JPG",
        subtitle="Transform lossless PNG graphics into lightweight, universally compatible JPEG images with solid white backgrounds and no black box artifacts.",
        badge="Alpha Compositing & Compression Engine",
        tech_overview="RGBA channels are extracted and composited onto a pure white background layer to eliminate dark borders or black transparency artifacts. Output files are encoded with high-fidelity chroma subsampling at your preferred quality level.",
        how_to_steps=[
            "Drop one or more PNG images into the converter.",
            "Select JPG from the Images category.",
            "Adjust the quality slider (e.g. 95%) to balance file size and visual fidelity.",
            "Click 'Convert Document' and download your clean JPG images instantly."
        ],
        faqs=[
            {
                "q": "Will transparent areas turn black in the JPG?",
                "a": "No! Velmo uses alpha compositing to render transparency against a clean, crisp white background, avoiding black box artifacts."
            },
            {
                "q": "How much smaller will the JPG file be compared to the PNG?",
                "a": "For photos and complex graphics, converting PNG to JPG typically reduces file size by 60% to 80% without noticeable visual difference."
            },
            {
                "q": "Can I convert multiple PNG files at once?",
                "a": "Yes. Upload all your PNGs at once to convert them simultaneously into individual JPGs bundled in a single ZIP download."
            }
        ],
        reverse_slug="jpg-to-png",
        related_slugs=["png-to-webp", "webp-to-png", "image-to-pdf", "pdf-to-png"]
    ),
    "jpg-to-png": SeoRoute(
        slug="jpg-to-png",
        from_fmt="jpg",
        to_fmt="png",
        title="JPG to PNG Converter (Lossless Image Output) — Velmo",
        description="Convert JPG images to lossless PNG format online for free. Preserves full color fidelity and prepare graphics for editing with zero file uploads stored.",
        h1="Convert JPG Images to Lossless PNG Graphics",
        subtitle="Transform compressed JPEG photos into lossless PNG images ready for graphic design, transparent layering, and artifact-free archiving.",
        badge="Lossless Bitmap Transcoding Engine",
        tech_overview="JPEG compressed DCT coefficients are decoded into uncompressed 24-bit RGB bitmap arrays. The pixel stream is re-encoded using PNG lossless DEFLATE compression with adaptive filtering, preventing further generational compression loss.",
        how_to_steps=[
            "Upload your JPG or JPEG photos to the converter.",
            "Select PNG from the Images category.",
            "Click 'Convert Document' to process your files into lossless PNG format.",
            "Download your clean PNG images individually or in a batch ZIP archive."
        ],
        faqs=[
            {
                "q": "Does converting JPG to PNG make the image sharper?",
                "a": "While it cannot restore detail lost during original JPEG compression, converting to PNG freezes the image in lossless format so editing won't cause further degradation."
            },
            {
                "q": "Can I convert high-resolution camera photos?",
                "a": "Yes. Velmo handles high-megapixel camera photos and mobile shots easily, processing them directly in server memory."
            },
            {
                "q": "Is this tool completely free to use?",
                "a": "Yes. All image conversions on Velmo are 100% free with no daily limits, watermarks, or account registrations."
            }
        ],
        reverse_slug="png-to-jpg",
        related_slugs=["webp-to-png", "png-to-webp", "image-to-pdf", "image-to-excel"]
    ),
    "webp-to-png": SeoRoute(
        slug="webp-to-png",
        from_fmt="webp",
        to_fmt="png",
        title="WebP to PNG Converter (Save WebP as PNG Free) — Velmo",
        description="Convert modern WebP images into universal lossless PNG files. Preserves full transparency, alpha channels, and crisp details across all web devices.",
        h1="Convert WebP Images to Universal Lossless PNG",
        subtitle="Save WebP graphics from web browsers into universally supported PNG files with complete transparency and alpha channel preservation.",
        badge="Universal Image Format Transcoding",
        tech_overview="Our image pipeline decodes VP8 and VP8L bitstreams, extracting both RGB color data and 8-bit alpha channels. The image is saved as a 32-bit RGBA PNG, guaranteeing compatibility across older image viewers, editing software, and operating systems.",
        how_to_steps=[
            "Drop your WebP image files into the upload zone.",
            "Select PNG as your target format.",
            "Click 'Convert Document' to transform your WebP files into universal PNG format.",
            "Download your transparent PNG images instantly."
        ],
        faqs=[
            {
                "q": "Why do WebP images sometimes fail to open in Photoshop or older software?",
                "a": "WebP is a modern web format not natively supported by some desktop software. Converting to PNG ensures universal compatibility with every editing tool."
            },
            {
                "q": "Will transparency be preserved from the WebP image?",
                "a": "Yes! Full alpha transparency is preserved with zero loss, allowing you to use the PNG on any background."
            },
            {
                "q": "Can I convert animated WebP images?",
                "a": "The primary frame of the WebP is extracted as a high-resolution still PNG graphic."
            }
        ],
        reverse_slug="png-to-webp",
        related_slugs=["jpg-to-png", "png-to-jpg", "image-to-pdf", "pdf-to-png"]
    ),
    "png-to-webp": SeoRoute(
        slug="png-to-webp",
        from_fmt="png",
        to_fmt="webp",
        title="PNG to WebP Converter (Compress Images for Web) — Velmo",
        description="Compress PNG images into lightweight next-gen WebP format. Reduces file sizes by up to 50% while preserving transparent backgrounds and sharp edges.",
        h1="Compress PNG Images into Next-Gen WebP Graphics",
        subtitle="Drastically speed up web page load times by converting heavy PNG graphics into lightweight, high-performance WebP images with full transparency.",
        badge="Next-Gen Image Compression Engine",
        tech_overview="Converts RGB and RGBA buffers into WebP bitstreams utilizing predictive intra-frame coding. Delivers 25% to 50% smaller file sizes than standard PNGs while maintaining crisp typography, sharp icon edges, and transparent alpha layers.",
        how_to_steps=[
            "Upload your PNG graphics, logos, or icons into the converter.",
            "Select WebP from the Images category.",
            "Adjust the quality slider (e.g. 95% for sharp graphics).",
            "Click 'Convert Document' and download your optimized WebP assets."
        ],
        faqs=[
            {
                "q": "How does WebP improve website Core Web Vitals?",
                "a": "Smaller WebP image payloads reduce Largest Contentful Paint (LCP) and bandwidth consumption, boosting your Google PageSpeed scores."
            },
            {
                "q": "Is transparency maintained in the WebP output?",
                "a": "Yes. WebP natively supports 8-bit alpha transparency while producing significantly smaller files than transparent PNGs."
            },
            {
                "q": "Are WebP images supported across modern web browsers?",
                "a": "Yes. WebP is supported natively across Google Chrome, Apple Safari, Mozilla Firefox, Microsoft Edge, and modern mobile browsers."
            }
        ],
        reverse_slug="webp-to-png",
        related_slugs=["png-to-jpg", "jpg-to-png", "image-to-pdf", "pdf-to-png"]
    ),
    "markdown-to-pdf": SeoRoute(
        slug="markdown-to-pdf",
        from_fmt="md",
        to_fmt="pdf",
        title="Markdown to PDF Converter (Styled Developer Docs) — Velmo",
        description="Convert Markdown (.md) documents into beautifully styled PDF files. Formats code blocks, tables, task lists, and typography with clean print styling.",
        h1="Convert Markdown (.md) to Styled PDF Documents",
        subtitle="Transform developer READMEs, technical documentation, and notes into publication-quality PDF documents with syntax styling and clean tables.",
        badge="Markdown Typography & Print Engine",
        tech_overview="Markdown AST structures are parsed into styled document elements. Headings, monospaced code blocks, bulleted lists, and tables are formatted with clean typographical margins and proportional spacing for executive-grade PDF output.",
        how_to_steps=[
            "Upload your Markdown (.md) file or README.",
            "Target format is automatically selected as PDF (.pdf).",
            "Click 'Convert Document' to generate a styled document.",
            "Download your PDF or use our in-browser editor to make live edits before exporting."
        ],
        faqs=[
            {
                "q": "Are code blocks and syntax highlighting preserved?",
                "a": "Yes. Monospaced code blocks, backtick spans, and indentation are formatted cleanly inside highlighted code boxes."
            },
            {
                "q": "Can I edit the Markdown copy in the browser before generating the PDF?",
                "a": "Yes! Click 'Preview & Edit' to open our live Markdown editor, update copy, check real-time word counts, and export to PDF."
            },
            {
                "q": "Does it format Markdown tables and lists?",
                "a": "Yes. GitHub-flavored Markdown tables, ordered lists, and blockquotes are rendered into clean, print-ready document elements."
            }
        ],
        reverse_slug="pdf-to-markdown",
        related_slugs=["markdown-to-word", "html-to-pdf", "word-to-pdf", "pdf-to-word"]
    ),
    "markdown-to-word": SeoRoute(
        slug="markdown-to-word",
        from_fmt="md",
        to_fmt="docx",
        title="Markdown to Word Converter (MD to DOCX Document) — Velmo",
        description="Convert Markdown (.md) files to editable Word (.docx) documents. Translates headings, bullet lists, bold text, and code blocks into native Word styles.",
        h1="Convert Markdown (.md) to Editable Word Document",
        subtitle="Translate technical Markdown files into polished Microsoft Word (.docx) documents with native heading hierarchies and formatted tables.",
        badge="Semantic Document AST Transpiler",
        tech_overview="The Markdown document parser maps tokens directly to Word OpenXML styles: `#` becomes Heading 1, `##` becomes Heading 2, code fences map to shaded monospaced tables, and pipes map to editable Word grid cells.",
        how_to_steps=[
            "Drop your .md or .txt file into the upload zone.",
            "Word (.docx) is automatically selected as the target format.",
            "Click 'Convert Document' to transform your Markdown copy into a Word file.",
            "Download your editable document or adjust the text in our live editor."
        ],
        faqs=[
            {
                "q": "Are native Word heading styles applied?",
                "a": "Yes. Markdown headings are converted into native Word 'Heading 1', 'Heading 2', and 'Heading 3' styles, allowing automatic Table of Contents creation."
            },
            {
                "q": "What happens to inline links and bold text?",
                "a": "Hyperlinks, bold styling, italic emphasis, and blockquotes are preserved as native rich text formatting inside Word."
            },
            {
                "q": "Can I export back to Markdown later?",
                "a": "Yes. Velmo supports bidirectional conversion, allowing you to convert Word documents back into clean Markdown whenever needed."
            }
        ],
        reverse_slug="word-to-html",
        related_slugs=["markdown-to-pdf", "word-to-pdf", "pdf-to-markdown", "word-to-txt"]
    ),
    "html-to-pdf": SeoRoute(
        slug="html-to-pdf",
        from_fmt="html",
        to_fmt="pdf",
        title="HTML to PDF Converter (Webpage Markup to PDF) — Velmo",
        description="Convert HTML files and web markup into print-ready PDF documents. Renders semantic tables, headings, and CSS styling into clean paginated documents.",
        h1="Convert HTML and Web Markup to Print-Ready PDF",
        subtitle="Transform HTML files, email templates, and web snippets into clean, paginated PDF documents with preserved typography and styled tables.",
        badge="HTML5 Semantic Layout Engine",
        tech_overview="Parses HTML DOM elements, inline CSS styles, and structure into paginated PDF flow objects. Handles tables, headings, ordered lists, and text paragraphs with clean page margins and automatic line breaks.",
        how_to_steps=[
            "Upload your HTML file (.html) or web markup into the converter.",
            "PDF is automatically configured as your target format.",
            "Click 'Convert Document' to paginate and render your markup.",
            "Inspect the rendered PDF in our page viewer and download immediately."
        ],
        faqs=[
            {
                "q": "Does it support inline CSS styles?",
                "a": "Yes. Standard inline CSS styles for fonts, colors, borders, and margins are parsed and rendered into the resulting PDF."
            },
            {
                "q": "Are HTML tables paginated across multiple pages?",
                "a": "Yes. Large tables spanning beyond a single page are automatically paginated with repeated header rows for clear reading."
            },
            {
                "q": "Can I convert raw HTML code without saving a file first?",
                "a": "You can paste your text or HTML into our live in-browser editor and export directly to PDF in one click."
            }
        ],
        reverse_slug="word-to-html",
        related_slugs=["markdown-to-pdf", "word-to-pdf", "excel-to-pdf", "pdf-to-word"]
    ),
    "pdf-to-txt": SeoRoute(
        slug="pdf-to-txt",
        from_fmt="pdf",
        to_fmt="txt",
        title="PDF to Text Converter (Extract Raw Text from PDF) — Velmo",
        description="Extract plain text from PDF documents instantly. Strips layout clutter to deliver clean, unformatted UTF-8 text for data extraction and LLM ingestion.",
        h1="Extract Clean Plain Text from PDF Documents",
        subtitle="Strip visual formatting and pull clean, unformatted UTF-8 text from PDF files for data processing, text analysis, and AI pipelines.",
        badge="High-Speed Unicode Text Extractor",
        tech_overview="Traverses PDF content streams, sorting text spans by vertical and horizontal reading order. Resolves hyphenated line wraps, strips decorative graphics, and outputs structured, clean plain text encoded in UTF-8.",
        how_to_steps=[
            "Upload your PDF document to the studio.",
            "Select Plain Text (.txt) from the Text & Markup category.",
            "Click 'Convert Document' to extract all readable text instantly.",
            "Download your .txt file or inspect the extracted text inside our live browser editor."
        ],
        faqs=[
            {
                "q": "Is the extracted text suitable for Large Language Models and AI training?",
                "a": "Yes! The output is clean UTF-8 text free of decorative layout tags, ideal for RAG pipelines, LLM prompts, and NLP processing."
            },
            {
                "q": "How does it handle two-column or multi-column PDF layouts?",
                "a": "The text extractor analyzes spatial reading coordinates to ensure columns are read in natural reading order without interleaving text."
            },
            {
                "q": "Does it support password-protected PDFs?",
                "a": "PDFs must be unlocked before uploading so the text stream can be decoded safely in server memory."
            }
        ],
        reverse_slug="markdown-to-pdf",
        related_slugs=["pdf-to-word", "pdf-to-markdown", "word-to-txt", "pdf-to-excel"]
    ),
    "word-to-txt": SeoRoute(
        slug="word-to-txt",
        from_fmt="docx",
        to_fmt="txt",
        title="Word to Text Converter (Extract DOCX to TXT) — Velmo",
        description="Extract clean raw text from Word (.docx) documents. Strips document formatting while preserving paragraph breaks and table content in plain text format.",
        h1="Extract Clean Plain Text from Word Documents",
        subtitle="Convert Microsoft Word (.docx) files into clean, lightweight plain text (.txt) files without formatting clutter or proprietary XML tags.",
        badge="Semantic Paragraph Parsing Engine",
        tech_overview="Navigates OpenXML paragraph nodes and runs, stripping styling tags, fonts, and colors while preserving essential structural spacing, paragraph line breaks, and table cell divisions.",
        how_to_steps=[
            "Upload your Word document (.docx) to the converter.",
            "Select Plain Text (.txt) from the Text & Markup category.",
            "Click 'Convert Document' to pull raw text out in milliseconds.",
            "Download your clean .txt file or view the copy in our live text editor."
        ],
        faqs=[
            {
                "q": "Will paragraph breaks and blank lines be preserved?",
                "a": "Yes. Logical paragraph divisions, bullet item lines, and table rows are preserved with clean newline characters."
            },
            {
                "q": "Can I convert large Word manuscripts or legal briefs?",
                "a": "Yes. Velmo processes multi-hundred-page Word documents in a fraction of a second inside ephemeral memory."
            },
            {
                "q": "Is any personal data or metadata left in the text file?",
                "a": "No. Author names, revision histories, and hidden metadata tags are completely stripped, leaving only pure text content."
            }
        ],
        reverse_slug="markdown-to-word",
        related_slugs=["pdf-to-txt", "word-to-html", "word-to-pdf", "pdf-to-markdown"]
    ),
    "pdf-to-markdown": SeoRoute(
        slug="pdf-to-markdown",
        from_fmt="pdf",
        to_fmt="md",
        title="PDF to Markdown Converter (Clean MD for LLMs) — Velmo",
        description="Convert PDF documents into structured Markdown (.md) files. Extracts headings, paragraphs, and tables for technical documentation and LLM processing.",
        h1="Convert PDF Documents into Structured Markdown",
        subtitle="Extract document text, headings, and tables into clean GitHub Flavored Markdown (.md) optimized for documentation and AI context windows.",
        badge="Structural Markdown Syntax Engine",
        tech_overview="Analyzes font weights and point sizes across PDF text spans to classify document headers (#, ##, ###). Table grids are converted into clean Markdown pipe syntax, and lists are formatted with standard hyphen bullets.",
        how_to_steps=[
            "Upload your PDF document into the converter.",
            "Select Markdown (.md) from the Text & Markup category.",
            "Click 'Convert Document' to synthesize structured Markdown copy.",
            "Download your .md file or edit it directly in our live browser Markdown workspace."
        ],
        faqs=[
            {
                "q": "How does it detect headings versus body text?",
                "a": "Our layout engine evaluates relative font sizes, line weights, and spacing gaps to assign appropriate Markdown header levels (#, ##, ###)."
            },
            {
                "q": "Are PDF tables converted into Markdown tables?",
                "a": "Yes. Rectangular table structures are automatically transformed into standard pipe-delimited Markdown tables."
            },
            {
                "q": "Can I load the output directly into ChatGPT or Claude?",
                "a": "Yes! Markdown is the ideal format for LLM prompting because it preserves document structure while consuming minimal tokens."
            }
        ],
        reverse_slug="markdown-to-pdf",
        related_slugs=["pdf-to-txt", "pdf-to-word", "markdown-to-word", "pdf-to-excel"]
    ),
    "word-to-html": SeoRoute(
        slug="word-to-html",
        from_fmt="docx",
        to_fmt="html",
        title="Word to HTML Converter (Clean Web-Ready HTML) — Velmo",
        description="Convert Word (.docx) documents to clean, semantic HTML5 markup. Strips bloated Microsoft Office tags to generate lightweight, web-ready HTML code.",
        h1="Convert Word Documents to Clean Semantic HTML5",
        subtitle="Transform Microsoft Word documents into lightweight, web-ready HTML5 code without proprietary mso tags or bloated styling junk.",
        badge="Semantic HTML5 Transpiler",
        tech_overview="Translates Word document XML elements into standard HTML5 semantic elements (<h1>, <h2>, <p>, <ul>, <table>). Removes obsolete Microsoft Office XML namespaces and inline styles to output clean, standards-compliant web markup.",
        how_to_steps=[
            "Upload your Microsoft Word document (.docx).",
            "Select HTML (.html) as the target format.",
            "Click 'Convert Document' to generate clean web markup.",
            "Download your HTML file or preview the formatted webpage in our viewer."
        ],
        faqs=[
            {
                "q": "Does it remove Microsoft Office junk tags and mso classes?",
                "a": "Yes! The output is clean, valid HTML5 with zero bloated Office namespaces or unnecessary conditional comments."
            },
            {
                "q": "Are document tables converted into clean HTML tables?",
                "a": "Yes. Word tables are converted into clean <table>, <thead>, <tbody>, <tr>, and <td> elements ready for modern CSS styling."
            },
            {
                "q": "Can I copy the HTML code directly into WordPress or a CMS?",
                "a": "Yes. The generated markup pastes cleanly into WordPress, Webflow, Notion, or any modern content management system without styling conflicts."
            }
        ],
        reverse_slug="html-to-pdf",
        related_slugs=["word-to-pdf", "markdown-to-word", "word-to-txt", "word-to-excel"]
    ),
    "csv-to-pdf": SeoRoute(
        slug="csv-to-pdf",
        from_fmt="csv",
        to_fmt="pdf",
        title="CSV to PDF Converter (Print-Ready Data Tables) — Velmo",
        description="Convert CSV data tables into formatted, printable PDF documents. Automatically paginates tabular data with repeated header rows for clear readability.",
        h1="Convert CSV Data Tables into Print-Ready PDF Documents",
        subtitle="Transform raw CSV records into polished, publication-ready PDF tables featuring clean grid lines, alternating row shading, and repeated headers.",
        badge="Tabular Pagination & Vector Rendering",
        tech_overview="Parses CSV records and computes optimal column widths based on cell text length. Uses a vector layout engine to render bordered data tables with repeating headers across page breaks and auto-switches to landscape for wide datasets.",
        how_to_steps=[
            "Upload your CSV file (.csv) into the converter.",
            "Target format is automatically pre-configured to PDF.",
            "Click 'Convert Document' to generate a cleanly formatted document.",
            "Inspect the table in our multi-page PDF viewer and download your PDF."
        ],
        faqs=[
            {
                "q": "What happens if my CSV has dozens of columns?",
                "a": "The layout engine evaluates table width and automatically switches to landscape orientation to fit all columns comfortably."
            },
            {
                "q": "Do table headers repeat on every page?",
                "a": "Yes. For multi-page datasets, the top header row repeats on every page so readers never lose track of column definitions."
            },
            {
                "q": "Can I edit the data values before creating the PDF?",
                "a": "Yes. Click 'Preview & Edit' to open our live spreadsheet grid, adjust cell values, and export directly to PDF."
            }
        ],
        reverse_slug="excel-to-csv",
        related_slugs=["csv-to-excel", "excel-to-pdf", "pdf-to-excel", "batch-converter"]
    ),
    "batch-converter": SeoRoute(
        slug="batch-converter",
        from_fmt="all",
        to_fmt="zip",
        title="Batch File Converter (Convert Multiple Files at Once) — Velmo",
        description="Convert multiple documents and images simultaneously across 11 formats. Free universal batch conversion with instant ZIP packaging and zero file limits.",
        h1="Universal Batch Document & Image Converter",
        subtitle="Queue dozens of files across PDF, Word, Excel, PPTX, CSV, and Images. Convert them all at once and download a clean, organized ZIP archive.",
        badge="Parallel Multi-File Conversion Engine",
        tech_overview="Our batch pipeline queues multiple uploaded files and processes them concurrently in ephemeral server memory. Converted assets are compiled on the fly into an organized ZIP archive with verified checksums and sequential naming.",
        how_to_steps=[
            "Drag and drop multiple files (PDFs, Office docs, spreadsheets, or images) together.",
            "Choose your desired output format from the Target Format selector.",
            "Click 'Convert Document' to process all queued files in one batch.",
            "Download your packaged ZIP archive containing all converted files."
        ],
        faqs=[
            {
                "q": "Can I convert mixed file types in the same batch?",
                "a": "Yes! You can upload a mix of JPGs, PNGs, and PDFs together, converting them to your selected target format in one single operation."
            },
            {
                "q": "How are the converted files downloaded?",
                "a": "All converted files are neatly organized and bundled into a single ZIP archive for fast, one-click downloading."
            },
            {
                "q": "Is there a limit on how many files I can convert at once?",
                "a": "There are no artificial file count restrictions. Large batches are converted rapidly in parallel inside server memory."
            }
        ],
        reverse_slug=None,
        related_slugs=["image-to-pdf", "pdf-to-png", "pdf-to-word", "image-to-excel"]
    ),
}

# Cache for the generated OG preview image
_OG_IMAGE_BYTES: Optional[bytes] = None


def generate_og_image() -> bytes:
    """
    Generates a crisp 1200x630 high-contrast social preview card for Velmo
    matching the Obsidian + Vermilion Ember aesthetic using Pillow, caching the result in memory.
    """
    global _OG_IMAGE_BYTES
    if _OG_IMAGE_BYTES is not None:
        return _OG_IMAGE_BYTES

    width, height = 1200, 630
    # Background #09090B (Deep Obsidian)
    im = Image.new("RGB", (width, height), color=(9, 9, 11))
    draw = ImageDraw.Draw(im)

    # Architectural hairline outer frame
    draw.rectangle([20, 20, width - 20, height - 20], outline=(30, 30, 36), width=1)
    
    # Elevated Carbon Card (#111114) with subtle specular top edge
    draw.rounded_rectangle([36, 36, width - 36, height - 36], radius=16, fill=(17, 17, 20), outline=(42, 42, 50), width=1)
    draw.line([(44, 37), (width - 44, 37)], fill=(255, 77, 28), width=2)

    # Brand icon box: Vermilion Ember pill
    icon_x, icon_y = 80, 80
    draw.rounded_rectangle([icon_x, icon_y, icon_x + 88, icon_y + 88], radius=18, fill=(255, 77, 28))
    draw.polygon(
        [
            (icon_x + 22, icon_y + 24),
            (icon_x + 66, icon_y + 24),
            (icon_x + 55, icon_y + 44),
            (icon_x + 44, icon_y + 64),
            (icon_x + 33, icon_y + 44),
        ],
        fill=(255, 255, 255)
    )

    # Typography fonts fallback
    font_large = None
    font_sub = None
    font_small = None
    for name in ["segoeui.ttf", "arial.ttf", "calibri.ttf", "DejaVuSans.ttf"]:
        try:
            font_large = ImageFont.truetype(name, 54)
            font_sub = ImageFont.truetype(name, 26)
            font_small = ImageFont.truetype(name, 18)
            break
        except Exception:
            pass

    if font_large is None:
        font_large = ImageFont.load_default()
        font_sub = font_large
        font_small = font_large

    # Brand name
    draw.text((icon_x + 115, icon_y + 14), "Velmo Studio", fill=(250, 250, 250), font=font_large)

    # Tagline
    draw.text(
        (80, 205),
        "Universal 11-Format Any-to-Any Document Converter",
        fill=(250, 250, 250),
        font=font_sub
    )

    # Secondary explanation
    draw.text(
        (80, 250),
        "High-craft, private conversion with in-browser editing & local neural OCR.",
        fill=(161, 161, 170),
        font=font_small
    )

    # 11 Format badges
    formats = ["PDF", "DOCX", "XLSX", "PPTX", "CSV", "PNG", "JPG", "WEBP", "MD", "TXT", "HTML"]
    bx, by = 80, 315
    for fmt in formats:
        badge_w = max(68, len(fmt) * 16 + 24)
        draw.rounded_rectangle([bx, by, bx + badge_w, by + 40], radius=8, fill=(23, 23, 27), outline=(45, 45, 54), width=1)
        draw.text((bx + 14, by + 10), fmt, fill=(255, 77, 28), font=font_small)
        bx += badge_w + 14

    # Value proposition pills at bottom (No raw emojis)
    features = [
        "Private In-Memory Stream Processing",
        "Zero File Retention & No Signup",
        "Local Neural OCR Table Extraction",
        "Instant ZIP Batch Bundling"
    ]
    fx, fy = 80, 420
    for feat in features:
        draw.rounded_rectangle([fx, fy, fx + 500, fy + 48], radius=10, fill=(23, 23, 27), outline=(42, 42, 50), width=1)
        draw.text((fx + 20, fy + 14), feat, fill=(228, 228, 231), font=font_small)
        fy += 60
        if fy > 530:
            fx += 520
            fy = 420

    # Domain watermark
    draw.text((width - 340, height - 75), "velmo.onrender.com", fill=(113, 113, 122), font=font_small)

    buf = io.BytesIO()
    im.save(buf, format="PNG", optimize=True)
    _OG_IMAGE_BYTES = buf.getvalue()
    return _OG_IMAGE_BYTES


def get_robots_txt(site_url: str = "https://velmo.onrender.com") -> str:
    """Generates production robots.txt crawl directives."""
    site_url = site_url.rstrip("/")
    return f"""User-agent: *
Allow: /
Allow: /tools
Disallow: /convert
Disallow: /convert/batch
Disallow: /preview-editable
Disallow: /export-edited
Disallow: /inspect-zip
Disallow: /preview-pdf-pages

Sitemap: {site_url}/sitemap.xml
"""


def get_sitemap_xml(site_url: str = "https://velmo.onrender.com") -> str:
    """Generates dynamic XML sitemap covering root, /tools, and all 28 registered programmatic routes."""
    site_url = site_url.rstrip("/")
    xml_lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
        '  <url>',
        f'    <loc>{site_url}/</loc>',
        '    <changefreq>daily</changefreq>',
        '    <priority>1.0</priority>',
        '  </url>',
        '  <url>',
        f'    <loc>{site_url}/tools</loc>',
        '    <changefreq>weekly</changefreq>',
        '    <priority>0.9</priority>',
        '  </url>',
    ]

    for slug in SEO_ROUTES.keys():
        xml_lines.extend([
            '  <url>',
            f'    <loc>{site_url}/{slug}</loc>',
            '    <changefreq>weekly</changefreq>',
            '    <priority>0.9</priority>',
            '  </url>',
        ])

    xml_lines.append('</urlset>')
    return "\n".join(xml_lines)


def get_llms_txt(site_url: str = "https://velmo.onrender.com") -> str:
    """
    Generates standardized /llms.txt for AI search engines (ChatGPT Search, Perplexity, Gemini, Claude).
    Conforms strictly to the copy rule: no internal Python package names.
    """
    site_url = site_url.rstrip("/")
    lines = [
        "# Velmo Document & Image Converter",
        "",
        f"> Velmo ({site_url}) is a free, privacy-first, universal 11-format document and image conversion studio with in-browser editing and local neural OCR.",
        "",
        "## Core Capabilities",
        "- **Universal Format Matrix**: Seamless any-to-any conversion across 11 formats: PDF, DOCX, XLSX, PPTX, CSV, PNG, JPG, WEBP, MD, TXT, and HTML.",
        "- **Local Neural OCR Engine**: On-device table and character recognition running entirely in ephemeral server memory with zero external cloud API calls.",
        "- **In-Browser Studio Workspace**: Interactive Markdown text editor and dynamic spreadsheet grid allowing document inspection and editing before final export.",
        "- **High-DPI PDF Rasterization**: Vector rendering up to 600 DPI for razor-sharp physical print quality and image extraction.",
        "- **Zero File Retention & Total Privacy**: Ephemeral in-memory stream processing. No files are saved to disk or persistent databases.",
        "- **Batch Processing & ZIP Bundling**: Multi-file queue with automated packaging of converted assets into organized ZIP archives.",
        "",
        "## Supported Formats",
        "- **Documents**: PDF (.pdf), Microsoft Word (.docx), PowerPoint (.pptx)",
        "- **Spreadsheets & Data**: Microsoft Excel (.xlsx), Comma-Separated Values (.csv)",
        "- **Images**: PNG (.png), JPEG (.jpg, .jpeg), WebP (.webp)",
        "- **Text & Web**: Markdown (.md), Plain Text (.txt), HTML (.html)",
        "",
        "## High-Intent Conversion Routes",
        f"- [All 28 Conversion Tools]({site_url}/tools): Complete catalog of all 28 specialized document, spreadsheet, image, OCR, and developer conversion routes.",
    ]

    for slug, r in SEO_ROUTES.items():
        lines.append(f"- [{r.title}]({site_url}/{slug}): {r.description}")

    lines.extend([
        "",
        "## API Endpoints",
        "- `POST /convert`: Universal file conversion endpoint accepting single or multiple files.",
        "- `POST /convert/batch`: High-performance batch conversion endpoint returning bundled ZIP archives.",
        "- `POST /preview-editable`: Extracts document Markdown text or spreadsheet table grids for live browser preview.",
        "- `POST /export-edited`: Compiles user edits made in the in-browser workspace into the selected format.",
        "- `POST /inspect-zip`: Inspects batch archive contents with thumbnail previews.",
        "- `POST /preview-pdf-pages`: Renders page-by-page document previews.",
    ])

    return "\n".join(lines)


def get_json_ld_schemas(route_slug: Optional[str], site_url: str = "https://velmo.onrender.com") -> str:
    """Generates valid JSON-LD schema blocks (WebApplication, BreadcrumbList, FAQPage)."""
    site_url = site_url.rstrip("/")

    web_app_schema = {
        "@context": "https://schema.org",
        "@type": "WebApplication",
        "name": "Velmo",
        "alternateName": "Velmo Studio",
        "url": f"{site_url}/",
        "applicationCategory": "UtilitiesApplication",
        "operatingSystem": "All (Web Browser, Windows, macOS, Linux, iOS, Android)",
        "browserRequirements": "Requires JavaScript. Requires HTML5.",
        "softwareVersion": "5.0.0",
        "offers": {
            "@type": "Offer",
            "price": "0",
            "priceCurrency": "USD",
            "availability": "https://schema.org/InStock",
        },
        "description": "Universal 11-format any-to-any document and image converter with in-browser editing and local neural OCR.",
        "featureList": [
            "Universal conversion across 11 formats: PDF, DOCX, XLSX, PPTX, CSV, PNG, JPG, WEBP, MD, TXT, HTML",
            "Local neural OCR table and text extraction without third-party cloud APIs",
            "In-browser document and spreadsheet editing before export",
            "High-resolution 300 & 600 DPI PDF rendering",
            "Batch multi-file conversion with automatic ZIP packaging",
            "100% private in-memory processing with no file retention"
        ],
        "creator": {
            "@type": "Person",
            "name": "Shintouwa",
            "url": "https://github.com/Shintouwa"
        }
    }

    schemas = [web_app_schema]

    if route_slug == "tools":
        breadcrumb_schema = {
            "@context": "https://schema.org",
            "@type": "BreadcrumbList",
            "itemListElement": [
                {
                    "@type": "ListItem",
                    "position": 1,
                    "name": "Velmo Home",
                    "item": f"{site_url}/",
                },
                {
                    "@type": "ListItem",
                    "position": 2,
                    "name": "All 28 Conversion Tools",
                    "item": f"{site_url}/tools",
                }
            ]
        }
        schemas.append(breadcrumb_schema)
    elif route_slug and route_slug in SEO_ROUTES:
        route = SEO_ROUTES[route_slug]
        breadcrumb_schema = {
            "@context": "https://schema.org",
            "@type": "BreadcrumbList",
            "itemListElement": [
                {
                    "@type": "ListItem",
                    "position": 1,
                    "name": "Velmo Home",
                    "item": f"{site_url}/",
                },
                {
                    "@type": "ListItem",
                    "position": 2,
                    "name": route.h1,
                    "item": f"{site_url}/{route.slug}",
                }
            ]
        }
        schemas.append(breadcrumb_schema)

        if route.faqs:
            faq_schema = {
                "@context": "https://schema.org",
                "@type": "FAQPage",
                "mainEntity": [
                    {
                        "@type": "Question",
                        "name": f["q"],
                        "acceptedAnswer": {
                            "@type": "Answer",
                            "text": f["a"]
                        }
                    }
                    for f in route.faqs
                ]
            }
            schemas.append(faq_schema)

    json_blocks = []
    for s in schemas:
        json_str = json.dumps(s, ensure_ascii=False, indent=2)
        json_blocks.append(f'<script type="application/ld+json">\n{json_str}\n</script>')

    return "\n".join(json_blocks)


def render_below_the_fold_content(route_slug: Optional[str], site_url: str = "https://velmo.onrender.com") -> str:
    """
    Renders clean, high-contrast SEO sections below the fold.
    For root (/): returns "" (homepage ends cleanly right after studio converter).
    For /tools: renders dedicated clean catalog of all 28 tools.
    For conversion routes: renders sleek, airy breadcrumbs, how-to steps, tech specs, FAQs, and reverse/related links.
    """
    site_url = site_url.rstrip("/")

    # Homepage has ZERO congestion below the fold
    if not route_slug:
        return ""

    # Dedicated /tools directory
    if route_slug == "tools":
        categories = {
            "Document Converters": (
                '<svg class="ui-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>',
                [
                    "pdf-to-word", "word-to-pdf", "pdf-to-pptx", "pptx-to-pdf",
                    "pdf-to-txt", "word-to-txt", "pdf-to-markdown", "word-to-html"
                ]
            ),
            "Spreadsheet & Data Tools": (
                '<svg class="ui-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="3" y1="9" x2="21" y2="9"/><line x1="3" y1="15" x2="21" y2="15"/><line x1="9" y1="3" x2="9" y2="21"/><line x1="15" y1="3" x2="15" y2="21"/></svg>',
                [
                    "pdf-to-excel", "excel-to-pdf", "image-to-excel", "excel-to-word",
                    "word-to-excel", "csv-to-excel", "excel-to-csv", "csv-to-pdf"
                ]
            ),
            "Image & OCR Workflows": (
                '<svg class="ui-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><polyline points="21 15 16 10 5 21"/></svg>',
                [
                    "image-to-word", "image-to-pdf", "pdf-to-png", "pdf-to-jpg",
                    "png-to-jpg", "jpg-to-png", "webp-to-png", "png-to-webp"
                ]
            ),
            "Markdown & Developer Utilities": (
                '<svg class="ui-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="16 18 22 12 16 6"/><polyline points="8 6 2 12 8 18"/></svg>',
                [
                    "markdown-to-pdf", "markdown-to-word", "html-to-pdf", "batch-converter"
                ]
            )
        }

        html = [
            '<section class="seo-section" id="toolsDirectorySection">',
            '  <div class="seo-container">',
            '    <nav class="seo-breadcrumbs" aria-label="Breadcrumb">',
            '      <a href="/">Velmo Studio</a> &rsaquo;',
            '      <span>All 28 Conversion Tools</span>',
            '    </nav>',
            '    <div class="seo-tool-header">',
            '      <span class="seo-badge">Universal Directory</span>',
            '      <h2 class="seo-h2">All 28 Document &amp; Image Conversion Tools</h2>',
            '      <p class="seo-lead">Dedicated high-speed conversion tools across all 11 supported formats. 100% free, private in-memory processing with zero file retention.</p>',
            '    </div>',
            '    <div class="seo-categories-grid">',
        ]

        for cat_name, (icon_svg, slugs) in categories.items():
            html.append('      <div class="seo-category-card">')
            html.append(f'        <h3 class="seo-cat-title">{icon_svg} <span>{cat_name}</span></h3>')
            html.append('        <ul class="seo-links-list">')
            for s in slugs:
                if s in SEO_ROUTES:
                    r = SEO_ROUTES[s]
                    from_label = r.from_fmt.upper() if r.from_fmt else "MULTI"
                    to_label = r.to_fmt.upper() if r.to_fmt else "ZIP"
                    html.append(f'          <li><a href="/{s}" class="seo-route-link" title="{xml_escape(r.title)}"><span class="seo-tool-pills">{from_label} &rarr; {to_label}</span><span class="seo-tool-title">{xml_escape(r.h1)}</span><span class="seo-tool-arrow">&nearr;</span></a></li>')
            html.append('        </ul>')
            html.append('      </div>')

        html.extend([
            '    </div>',
            '    <div style="text-align: center; margin-top: 36px;">',
            '      <a href="/" class="btn-secondary" style="text-decoration: none; display: inline-flex; align-items: center; gap: 8px; padding: 10px 20px;">&larr; Return to Studio Converter</a>',
            '    </div>',
            '  </div>',
            '</section>',
        ])
        return "\n".join(html)

    if route_slug not in SEO_ROUTES:
        return ""

    # Conversion route specific landing section
    route = SEO_ROUTES[route_slug]

    reverse_link_html = ""
    if route.reverse_slug and route.reverse_slug in SEO_ROUTES:
        rev = SEO_ROUTES[route.reverse_slug]
        reverse_link_html = f"""
        <div class="seo-reverse-box">
          <span class="seo-reverse-label">
            <svg class="ui-icon" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M7 16V4m0 0L3 8m4-4l4 4m6 4v12m0 0l4-4m-4 4l-4-4"/></svg>
            Reverse Conversion:
          </span>
          <a href="/{rev.slug}" class="seo-reverse-link" title="{xml_escape(rev.title)}">{xml_escape(rev.h1)} &rarr;</a>
        </div>
        """

    related_links_html = []
    for rel_slug in route.related_slugs:
        if rel_slug in SEO_ROUTES:
            rel = SEO_ROUTES[rel_slug]
            related_links_html.append(
                f'<a href="/{rel.slug}" class="seo-pill-link" title="{xml_escape(rel.title)}">{xml_escape(rel.h1)}</a>'
            )

    steps_html = []
    for idx, step_text in enumerate(route.how_to_steps, 1):
        steps_html.append(f"""
          <div class="seo-step-card">
            <div class="seo-step-num">{idx:02d}</div>
            <div class="seo-step-body">{xml_escape(step_text)}</div>
          </div>
        """)

    faqs_html = []
    for f in route.faqs:
        faqs_html.append(f"""
          <details class="seo-faq-item">
            <summary class="seo-faq-question">
              <span>{xml_escape(f['q'])}</span>
              <svg class="faq-chevron" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="6 9 12 15 18 9"/></svg>
            </summary>
            <div class="seo-faq-answer">{xml_escape(f['a'])}</div>
          </details>
        """)

    html = f"""
    <section class="seo-section" id="seoPairSection">
      <div class="seo-container">
        <!-- Breadcrumbs -->
        <nav class="seo-breadcrumbs" aria-label="Breadcrumb">
          <a href="/">Velmo Studio</a> &rsaquo;
          <a href="/tools">All Tools</a> &rsaquo;
          <span>{xml_escape(route.h1)}</span>
        </nav>

        <!-- Tool Header -->
        <div class="seo-tool-header">
          <span class="seo-badge">{xml_escape(route.badge)}</span>
          <h2 class="seo-h2">{xml_escape(route.h1)}</h2>
          <p class="seo-lead">{xml_escape(route.subtitle)}</p>
        </div>

        <!-- How To Steps -->
        <div class="seo-block">
          <h3 class="seo-h3">How to Convert in 4 Simple Steps</h3>
          <div class="seo-steps-grid">
            {''.join(steps_html)}
          </div>
        </div>

        <!-- Technical Deep Dive -->
        <div class="seo-block seo-tech-block">
          <h3 class="seo-h3">Engineering &amp; Processing Specifications</h3>
          <p class="seo-tech-text">{xml_escape(route.tech_overview)}</p>
        </div>

        <!-- FAQs Accordion -->
        <div class="seo-block">
          <h3 class="seo-h3">Frequently Asked Questions</h3>
          <div class="seo-faqs-accordion">
            {''.join(faqs_html)}
          </div>
        </div>

        <!-- Reverse & Related Links -->
        <div class="seo-block seo-links-block">
          {reverse_link_html}
          <div class="seo-related-group">
            <span class="seo-related-title">Related Conversion Tools:</span>
            <div class="seo-pills-wrap">
              {''.join(related_links_html)}
            </div>
          </div>
        </div>
      </div>
    </section>
    """
    return html


def render_html_page(
    base_html: str,
    route_slug: Optional[str] = None,
    site_url: str = "https://velmo.onrender.com",
    google_verification: str = "",
    bing_verification: str = "",
) -> str:
    """
    Renders full server-rendered HTML for root (/), /tools, or any of the 28 programmatic routes.
    Injects dynamic title, meta description, canonical tag, OpenGraph tags, JSON-LD schemas,
    site verification meta tags, preselection attributes, and below-the-fold content.
    """
    site_url = site_url.rstrip("/")

    if route_slug == "tools":
        page_title = "All 28 Document & Image Converters — Velmo Tools Directory"
        page_description = "Browse all 28 specialized document, spreadsheet, image, OCR, and developer file conversion tools available on Velmo. Free, private, and unlimited."
        canonical_url = f"{site_url}/tools"
        preselect_from = ""
        preselect_to = ""
    elif route_slug and route_slug in SEO_ROUTES:
        route = SEO_ROUTES[route_slug]
        page_title = route.title
        page_description = route.description
        canonical_url = f"{site_url}/{route.slug}"
        preselect_from = route.from_fmt
        preselect_to = route.to_fmt
    else:
        page_title = "Velmo Studio — Universal 11-Format Any-to-Any Document Converter"
        page_description = "Velmo — Universal 11-Format Any-to-Any Document Converter with in-browser editing & local neural OCR. 100% free, private, with zero file retention."
        canonical_url = f"{site_url}/"
        preselect_from = ""
        preselect_to = ""

    # Verification tags
    verification_tags = []
    if google_verification:
        verification_tags.append(f'<meta name="google-site-verification" content="{xml_escape(google_verification)}">')
    if bing_verification:
        verification_tags.append(f'<meta name="msvalidate.01" content="{xml_escape(bing_verification)}">')
    verification_html = "\n  ".join(verification_tags)

    # OpenGraph & Twitter tags
    og_html = f"""
  <meta property="og:type" content="website">
  <meta property="og:site_name" content="Velmo">
  <meta property="og:title" content="{xml_escape(page_title)}">
  <meta property="og:description" content="{xml_escape(page_description)}">
  <meta property="og:url" content="{canonical_url}">
  <meta property="og:image" content="{site_url}/og-image.png">
  <meta property="og:image:width" content="1200">
  <meta property="og:image:height" content="630">
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="{xml_escape(page_title)}">
  <meta name="twitter:description" content="{xml_escape(page_description)}">
  <meta name="twitter:image" content="{site_url}/og-image.png">"""

    # Structured Data
    json_ld_html = get_json_ld_schemas(route_slug, site_url=site_url)

    # Below the fold content
    below_the_fold_html = render_below_the_fold_content(route_slug, site_url=site_url)

    # Clean injection into base_html
    rendered = base_html

    # Replace <title>...</title>
    rendered = re.sub(
        r'<title>.*?</title>',
        f'<title>{xml_escape(page_title)}</title>',
        rendered,
        flags=re.IGNORECASE | re.DOTALL,
    )

    # Replace or insert meta description
    if '<meta name="description"' in rendered:
        rendered = re.sub(
            r'<meta name="description" content=".*?">',
            f'<meta name="description" content="{xml_escape(page_description)}">',
            rendered,
            flags=re.IGNORECASE,
        )
    else:
        rendered = rendered.replace('</head>', f'  <meta name="description" content="{xml_escape(page_description)}">\n</head>')

    # Inject Canonical, Verification, OpenGraph, JSON-LD before </head>
    head_injections = f"""
  <link rel="canonical" href="{canonical_url}">
  <link rel="icon" type="image/svg+xml" href="data:image/svg+xml,{xml_escape(FAVICON_SVG)}">
  {verification_html}
  {og_html}
  {json_ld_html}
"""
    rendered = rendered.replace('</head>', f"{head_injections}\n</head>")

    # Inject data attributes onto <body>
    rendered = re.sub(
        r'<body(.*?)>',
        f'<body\\1 data-preselect-from="{preselect_from}" data-preselect-to="{preselect_to}">',
        rendered,
        count=1,
    )

    # Inject below-the-fold content right before footer (if any content to inject)
    if below_the_fold_html:
        if '<footer class="footer">' in rendered:
            rendered = rendered.replace(
                '<footer class="footer">',
                f"{below_the_fold_html}\n  <footer class=\"footer\">"
            )
        else:
            rendered = rendered.replace('</body>', f"{below_the_fold_html}\n</body>")

    return rendered
