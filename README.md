# Velmo — Universal 11-Format Any-to-Any Document Converter

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/Shintouwa/velmo)
[![GitHub Repository](https://img.shields.io/badge/GitHub-Shintouwa%2Fvelmo-6366f1?style=flat&logo=github)](https://github.com/Shintouwa/velmo)
[![FastAPI](https://img.shields.io/badge/FastAPI-5.0.0-009688?style=flat&logo=fastapi)](https://fastapi.tiangolo.com/)
[![Python](https://img.shields.io/badge/Python-3.11.9-3776AB?style=flat&logo=python)](https://python.org)

A modern, fast, and unrestricted document and image conversion web service and API built with FastAPI, PyMuPDF, pdf2docx, ReportLab, RapidOCR, openpyxl, python-docx, and python-pptx.

**Repository:** [https://github.com/Shintouwa/velmo](https://github.com/Shintouwa/velmo)

---

## 🚀 Supported Formats (11 Formats Any-to-Any)

| Category | Formats | Description & Engine |
| :--- | :--- | :--- |
| **Documents** | **PDF** (`.pdf`), **Word** (`.docx`), **PowerPoint** (`.pptx`) | PyMuPDF, pdf2docx, python-docx, python-pptx, ReportLab fallback, and LibreOffice headless when present. |
| **Data** | **Excel** (`.xlsx`), **CSV** (`.csv`) | openpyxl, Python CSV engine, numeric parsing, multi-sheet workbook generation & table clustering. |
| **Images** | **PNG** (`.png`), **JPG** (`.jpg`/`.jpeg`), **WebP** (`.webp`) | Pillow, img2pdf, 72–600 DPI custom rasterization, lossless WebP/PNG, alpha-channel flattening onto white. |
| **Text & Web** | **Markdown** (`.md`), **Plain Text** (`.txt`), **HTML** (`.html`) | Markdown + BeautifulSoup4 parser into native Word/PDF layouts; structured clean text export. |

---

## ✨ Key Features

- **In-Browser Editable Preview & Direct Export:**
  - **Spreadsheet Grid Editor (`POST /preview-editable` & `POST /export-edited`):** Inspect OCR or extracted tabular data across multiple sheet tabs, edit cell values directly in an interactive grid, dynamically add/delete rows and columns, and export directly to `.xlsx`, `.csv`, `.docx`, or `.pdf`.
  - **Document Text Editor:** Inspect extracted text with headings, lists, and tables as clean Markdown, edit copy directly with live word and character counters, switch target formats on the fly, and export directly to `.docx`, `.pdf`, `.pptx`, `.html`, or `.txt`.
- **Hub-and-Spoke Any-to-Any Architecture:** Every supported format converts to every other supported format seamlessly.
- **Batch Processing of Mixed Formats (`POST /convert/batch`):** Convert any combination of different files (e.g., `.docx` + `.xlsx` + `.md`) into a target format at once, automatically bundled into a single ZIP archive with automatic filename deduplication and clean subfolders for multi-page documents (zero nested double-ZIPs).
- **Unrestricted Production Usage:** Zero artificial file-size caps, zero rate limits, and full streaming upload handling.
- **Smart Conversion Options:**
  - **Resolution Control (DPI):** Select `150 DPI (Web)`, `300 DPI (Print)`, or `600 DPI (Ultra)` when rendering documents to images.
  - **Quality Compression:** Custom slider from 60% to 100% for JPG and WebP outputs.
  - **PDF Page Range:** Extract specific pages (e.g. `1-3, 5`) with graceful out-of-range filtering and 400 validation on malformed inputs.
- **Local On-Premises OCR:** Embedded `RapidOCR` engine automatically extracts editable paragraphs, headings, and tabular rows/columns from scanned documents or screenshots without cloud APIs.
- **Zero Disk Footprint:** All conversions process in-memory (`io.BytesIO`) or within scoped `tempfile.TemporaryDirectory()` contexts that clean up automatically on exit.
- **Production Filename Sanitization:** Traversal characters (`../`), null bytes, and illegal symbols are stripped before generating `Content-Disposition` download headers.
- **Interactive Web UI:**
  - Modern studio interface with high-contrast Dark & Light mode toggle.
  - Drag-and-drop file upload with format auto-detection.
  - Live image thumbnails with Move Up (`↑`) / Move Down (`↓`) reordering controls.
  - Instant in-browser result preview (images, inline PDF viewer, scrollable text/code viewer).
  - Session history with 1-click preview and re-download.

---

## 🌐 Deploy to Render (100% Free Native Python Tier)

Velmo is optimized to deploy seamlessly on [Render](https://render.com)'s free tier as a native Python web service — with **zero Docker overhead and no LibreOffice dependencies required**.

### Option A: 1-Click Blueprint Deploy (Recommended)

1. Click the button below:

   [![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/Shintouwa/velmo)

2. Render automatically parses [`render.yaml`](render.yaml) from this repository and pre-configures:
   - **Service Name:** `velmo`
   - **Runtime:** `python`
   - **Plan:** `free`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`
   - **Health Check Path:** `/health`
   - **Python Version:** `3.11.9`
3. Click **Apply** to deploy your live instance!

### Option B: Manual Web Service Setup on Render

1. Log in to [Render Dashboard](https://dashboard.render.com/) and click **New +** → **Web Service**.
2. Connect your GitHub repository: `https://github.com/Shintouwa/velmo`.
3. Configure the settings:
   - **Name:** `velmo`
   - **Region:** Any (e.g., Oregon, Frankfurt)
   - **Branch:** `main`
   - **Runtime:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `uvicorn main:app --host 0.0.0.0 --port $PORT`
   - **Plan:** `Free`
4. Expand **Advanced**:
   - Add Environment Variable: `PYTHON_VERSION` = `3.11.9`
   - Set **Health Check Path:** `/health`
5. Click **Create Web Service**.

---

## 🛠️ Local Development

### 1. Set Up Virtual Environment

```bash
cd doc-converter
python -m venv .venv

# Windows
.venv\Scripts\activate

# macOS / Linux
source .venv/bin/activate
```

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Run the Development Server

```bash
# Starts on http://localhost:8000 (respects $PORT environment variable)
python main.py

# Or via Uvicorn with auto-reload
uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

---

## 🧪 Automated Testing

Run the full end-to-end automated test suite:

```bash
python test_server.py
```

The test suite validates:
- API metadata (`GET /api/info`) and health checks (`GET /health`)
- All 21 core document, data, presentation, image, and OCR conversions
- Custom resolution rasterization (`dpi=150` vs `dpi=300`)
- Page-range extraction (`page_range="2"`) and 400 error handling (`page_range="abc"`)
- Path traversal & filename sanitization (`../../evil name?.pdf`)
- Mixed-format batch conversion with deduplication and subfolder packaging
- In-browser editable preview & direct export for spreadsheet tables and markdown documents
- Live ZIP inspection with image thumbnails and text snippet extraction (all 31 tests passing with 100% success)

---

## 🐳 Docker Deployment (Optional)

The application also includes a Debian-slim Dockerfile with LibreOffice and standard fonts pre-installed for environments where LibreOffice rendering is desired:

```bash
# Build the Docker image
docker build -t velmo .

# Run container (mapped to port 8000)
docker run -p 8000:8000 -e PORT=8000 velmo
```

Check health:
```bash
curl http://localhost:8000/health
```

---

## ☁️ Cloud Deployment (Google Cloud Run / AWS / Generic Host)

### Google Cloud Run

```bash
gcloud run deploy velmo \
  --source . \
  --platform managed \
  --region us-central1 \
  --allow-unauthenticated \
  --memory 2Gi \
  --cpu 2 \
  --timeout 300
```

### Generic Cloud Host / Docker Compose

Set the `PORT` environment variable (e.g. `PORT=8080`), allocate at least 1 GB of RAM, and point health probes to `GET /health`.
