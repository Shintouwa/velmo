import base64
import io
import logging
import os
from pathlib import Path
from typing import List, Optional
import zipfile

from fastapi import FastAPI, File, Form, UploadFile, HTTPException
from fastapi.responses import FileResponse, Response, JSONResponse
from pydantic import BaseModel

import converters

logger = logging.getLogger("velmo")

app = FastAPI(title="Velmo - 11-Format Universal Any-to-Any Converter", version="5.0.0")

BASE_DIR = Path(__file__).resolve().parent
INDEX_HTML_PATH = BASE_DIR / "index.html"

SUPPORTED_FORMATS = {
    "pdf", "docx", "xlsx", "pptx", "csv", "txt", "md", "html", "png", "jpg", "jpeg", "webp"
}

CATEGORIES = {
    "Documents": ["pdf", "docx", "pptx"],
    "Data": ["xlsx", "csv"],
    "Images": ["png", "jpg", "webp"],
    "Text & Web": ["md", "txt", "html"],
}


@app.get("/")
async def get_index():
    if not INDEX_HTML_PATH.exists():
        raise HTTPException(status_code=404, detail="index.html not found")
    return FileResponse(INDEX_HTML_PATH, media_type="text/html")


@app.get("/api/info")
async def get_info():
    return {
        "libreoffice_available": converters.LIBREOFFICE_PATH is not None,
        "libreoffice_path": converters.LIBREOFFICE_PATH,
        "supported_formats": sorted(list(SUPPORTED_FORMATS)),
        "categories": CATEGORIES,
    }


@app.get("/health")
async def health_check():
    return {
        "status": "ok",
        "libreoffice": converters.LIBREOFFICE_PATH is not None,
        "formats": 11,
    }


# --- Unified Any-to-Any Endpoint ---

@app.post("/convert")
async def universal_convert_endpoint(
    target_format: Optional[str] = Form(None),
    target: Optional[str] = None,
    dpi: Optional[int] = Form(300),
    image_quality: Optional[int] = Form(95),
    page_range: Optional[str] = Form(None),
    files: List[UploadFile] = File(default=[]),
    file: Optional[UploadFile] = File(default=None),
):
    fmt = (target_format or target or "").strip().lower()
    if not fmt:
        raise HTTPException(status_code=400, detail="Missing target_format parameter.")

    fmt_normalized = fmt.replace("jpeg", "jpg")
    if fmt_normalized not in SUPPORTED_FORMATS and fmt not in SUPPORTED_FORMATS:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported target format '{fmt}'. Choose from: {', '.join(sorted(SUPPORTED_FORMATS))}",
        )

    # Validate DPI
    effective_dpi = 300 if dpi is None else dpi
    if effective_dpi < 72 or effective_dpi > 600:
        raise HTTPException(status_code=400, detail="dpi must be between 72 and 600.")

    # Validate Image Quality
    effective_quality = 95 if image_quality is None else image_quality
    if effective_quality < 60 or effective_quality > 100:
        raise HTTPException(status_code=400, detail="image_quality must be between 60 and 100.")

    upload_list = list(files) if files else []
    if file and file.filename:
        upload_list.append(file)

    if not upload_list:
        raise HTTPException(status_code=400, detail="No files uploaded.")

    # Read uploaded file contents into memory
    input_files = []
    for f in upload_list:
        fname = f.filename or "document"
        raw_bytes = await f.read()
        if not raw_bytes:
            raise HTTPException(status_code=400, detail=f"Uploaded file '{fname}' is empty.")
        input_files.append((fname, raw_bytes))

    primary_filename = input_files[0][0]
    base_name = converters.sanitize_base_name(primary_filename)

    try:
        content, media_type, out_name, count = converters.universal_convert(
            input_files,
            fmt_normalized,
            base_name,
            dpi=effective_dpi,
            image_quality=effective_quality,
            page_range=page_range,
        )
        safe_out_name = converters.sanitize_filename(out_name)
        return Response(
            content=content,
            media_type=media_type,
            headers={
                "Content-Disposition": f'attachment; filename="{safe_out_name}"',
                "X-Count": str(count),
            },
        )
    except ValueError as val_err:
        raise HTTPException(status_code=400, detail=str(val_err))
    except Exception as e:
        logger.error(f"Conversion error to '{fmt}': {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Conversion error: {str(e)}")


# --- Batch Conversion Endpoint ---

@app.post("/convert/batch")
async def batch_convert_endpoint(
    target_format: Optional[str] = Form(None),
    target: Optional[str] = None,
    dpi: Optional[int] = Form(300),
    image_quality: Optional[int] = Form(95),
    page_range: Optional[str] = Form(None),
    files: List[UploadFile] = File(default=[]),
    file: Optional[UploadFile] = File(default=None),
):
    fmt = (target_format or target or "").strip().lower()
    if not fmt:
        raise HTTPException(status_code=400, detail="Missing target_format parameter.")

    fmt_normalized = fmt.replace("jpeg", "jpg")
    if fmt_normalized not in SUPPORTED_FORMATS and fmt not in SUPPORTED_FORMATS:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported target format '{fmt}'. Choose from: {', '.join(sorted(SUPPORTED_FORMATS))}",
        )

    effective_dpi = 300 if dpi is None else dpi
    if effective_dpi < 72 or effective_dpi > 600:
        raise HTTPException(status_code=400, detail="dpi must be between 72 and 600.")

    effective_quality = 95 if image_quality is None else image_quality
    if effective_quality < 60 or effective_quality > 100:
        raise HTTPException(status_code=400, detail="image_quality must be between 60 and 100.")

    upload_list = list(files) if files else []
    if file and file.filename:
        upload_list.append(file)

    if not upload_list:
        raise HTTPException(status_code=400, detail="No files uploaded.")

    input_files = []
    for f in upload_list:
        fname = f.filename or "document"
        raw_bytes = await f.read()
        if not raw_bytes:
            raise HTTPException(status_code=400, detail=f"Uploaded file '{fname}' is empty.")
        input_files.append((fname, raw_bytes))

    try:
        content, media_type, out_name, count = converters.batch_convert(
            input_files,
            fmt_normalized,
            dpi=effective_dpi,
            image_quality=effective_quality,
            page_range=page_range,
        )
        safe_out_name = converters.sanitize_filename(out_name)
        return Response(
            content=content,
            media_type=media_type,
            headers={
                "Content-Disposition": f'attachment; filename="{safe_out_name}"',
                "X-Batch-Count": str(count),
            },
        )
    except ValueError as val_err:
        raise HTTPException(status_code=400, detail=str(val_err))
    except Exception as e:
        logger.error(f"Batch conversion error to '{fmt}': {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Batch conversion error: {str(e)}")


# --- In-Browser Editable Preview & Export Endpoints ---

class ExportEditedPayload(BaseModel):
    mode: str
    filename: Optional[str] = "edited_document"
    target_format: str
    content: Optional[str] = None
    sheets: Optional[List[dict]] = None
    dpi: Optional[int] = 300
    image_quality: Optional[int] = 95


@app.post("/preview-editable")
async def preview_editable_endpoint(
    files: List[UploadFile] = File(default=[]),
    file: Optional[UploadFile] = File(default=None),
    target_format: Optional[str] = Form(default=None),
    target: Optional[str] = Form(default=None),
    page_range: Optional[str] = Form(default=None),
):
    fmt = target_format or target
    if not fmt:
        raise HTTPException(status_code=400, detail="Missing target_format parameter.")

    fmt_normalized = fmt.lower().strip().replace("jpeg", "jpg")
    if fmt_normalized not in SUPPORTED_FORMATS:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported target format '{fmt}'. Choose from: {', '.join(sorted(SUPPORTED_FORMATS))}",
        )

    upload_list = list(files) if files else []
    if file and file.filename:
        upload_list.append(file)

    if not upload_list:
        raise HTTPException(status_code=400, detail="No files uploaded.")

    input_files = []
    for f in upload_list:
        fname = f.filename or "document"
        raw_bytes = await f.read()
        if not raw_bytes:
            raise HTTPException(status_code=400, detail=f"Uploaded file '{fname}' is empty.")
        input_files.append((fname, raw_bytes))

    try:
        preview_data = converters.preview_editable(
            input_files,
            fmt_normalized,
            page_range=page_range,
        )
        return JSONResponse(content=preview_data)
    except ValueError as val_err:
        raise HTTPException(status_code=400, detail=str(val_err))
    except Exception as e:
        logger.error(f"Preview extraction error for '{fmt}': {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Preview extraction error: {str(e)}")


@app.post("/export-edited")
async def export_edited_endpoint(payload: ExportEditedPayload):
    fmt = payload.target_format
    if not fmt:
        raise HTTPException(status_code=400, detail="Missing target_format.")

    fmt_normalized = fmt.lower().strip().replace("jpeg", "jpg")
    if fmt_normalized not in SUPPORTED_FORMATS:
        raise HTTPException(
            status_code=400,
            detail=f"Unsupported target format '{fmt}'. Choose from: {', '.join(sorted(SUPPORTED_FORMATS))}",
        )

    effective_dpi = 300 if payload.dpi is None else payload.dpi
    if effective_dpi < 72 or effective_dpi > 600:
        raise HTTPException(status_code=400, detail="dpi must be between 72 and 600.")

    effective_quality = 95 if payload.image_quality is None else payload.image_quality
    if effective_quality < 60 or effective_quality > 100:
        raise HTTPException(status_code=400, detail="image_quality must be between 60 and 100.")

    try:
        content, media_type, out_name, count = converters.export_edited(
            mode=payload.mode,
            filename=payload.filename or "edited_document",
            target_format=fmt_normalized,
            content=payload.content,
            sheets=payload.sheets,
            dpi=effective_dpi,
            image_quality=effective_quality,
        )
        safe_out_name = converters.sanitize_filename(out_name)
        return Response(
            content=content,
            media_type=media_type,
            headers={
                "Content-Disposition": f'attachment; filename="{safe_out_name}"',
                "X-Count": str(count),
            },
        )
    except ValueError as val_err:
        raise HTTPException(status_code=400, detail=str(val_err))
    except Exception as e:
        logger.error(f"Export edited error for '{fmt}': {e}", exc_info=True)
        raise HTTPException(status_code=500, detail=f"Export edited error: {str(e)}")


@app.post("/inspect-zip")
async def inspect_zip_endpoint(file: UploadFile = File(...)):
    """Inspects contents of a ZIP archive and extracts previews for images/texts."""
    raw = await file.read()
    if not raw or not zipfile.is_zipfile(io.BytesIO(raw)):
        raise HTTPException(status_code=400, detail="Uploaded file is not a valid ZIP archive.")

    entries = []
    with zipfile.ZipFile(io.BytesIO(raw), "r") as zf:
        for info in zf.infolist():
            if info.is_dir() or info.filename.endswith("/") or info.filename.endswith("\\"):
                continue
            name = converters.sanitize_filename(os.path.basename(info.filename))
            full_path = info.filename
            size = info.file_size
            ext = Path(name).suffix.lower().lstrip(".")

            entry_data = {
                "name": name,
                "path": full_path,
                "size": size,
                "ext": ext,
            }

            # Generate preview if image
            if ext in ("png", "jpg", "jpeg", "webp"):
                img_bytes = zf.read(info.filename)
                mime = f"image/{ext if ext != 'jpg' else 'jpeg'}"
                b64 = base64.b64encode(img_bytes).decode("ascii")
                entry_data["type"] = "image"
                entry_data["preview"] = f"data:{mime};base64,{b64}"
            # Text/data preview
            elif ext in ("txt", "csv", "md", "html"):
                txt_bytes = zf.read(info.filename)
                try:
                    txt = txt_bytes.decode("utf-8")
                except UnicodeDecodeError:
                    txt = txt_bytes.decode("latin-1", errors="replace")
                entry_data["type"] = "text"
                entry_data["preview"] = txt[:3000]
            else:
                entry_data["type"] = "other"
                entry_data["preview"] = None

            entries.append(entry_data)

    return JSONResponse(content={
        "filename": converters.sanitize_filename(file.filename or "archive.zip"),
        "total_files": len(entries),
        "entries": entries,
    })


# --- Backward-Compatible Endpoints ---

@app.post("/convert/pdf-to-png")
async def convert_pdf_to_png_endpoint(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Uploaded file must be a PDF document.")
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    base_name = converters.sanitize_base_name(file.filename)
    content, media_type, out_name, page_cnt = converters.pdf_to_png(data, base_name)
    safe_out_name = converters.sanitize_filename(out_name)
    return Response(
        content=content,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{safe_out_name}"', "X-Page-Count": str(page_cnt)},
    )


@app.post("/convert/image-to-pdf")
async def convert_image_to_pdf_endpoint(
    files: List[UploadFile] = File(default=[]),
    file: Optional[UploadFile] = File(default=None),
):
    upload_list = list(files) if files else []
    if file and file.filename:
        upload_list.append(file)
    if not upload_list:
        raise HTTPException(status_code=400, detail="No image file(s) provided.")

    allowed_exts = {".png", ".jpg", ".jpeg", ".webp"}
    image_tuples = []
    for f in upload_list:
        f_name = f.filename or "image"
        if Path(f_name).suffix.lower() not in allowed_exts:
            raise HTTPException(status_code=400, detail=f"Unsupported format '{Path(f_name).suffix}' for '{f_name}'")
        raw = await f.read()
        if not raw:
            raise HTTPException(status_code=400, detail=f"File '{f_name}' is empty.")
        image_tuples.append((f_name, raw))

    base_name = converters.sanitize_base_name(image_tuples[0][0])
    content, media_type, out_name, cnt = converters.images_to_pdf(image_tuples, base_name)
    safe_out_name = converters.sanitize_filename(out_name)
    return Response(
        content=content,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{safe_out_name}"', "X-Page-Count": str(cnt)},
    )


@app.post("/convert/xlsx-to-docx")
async def convert_xlsx_to_docx_endpoint(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Uploaded file must be an Excel (.xlsx) spreadsheet.")
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    base_name = converters.sanitize_base_name(file.filename)
    content, media_type, out_name, cnt = converters.xlsx_to_docx(data, base_name)
    safe_out_name = converters.sanitize_filename(out_name)
    return Response(
        content=content,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{safe_out_name}"', "X-Sheet-Count": str(cnt)},
    )


@app.post("/convert/docx-to-xlsx")
async def convert_docx_to_xlsx_endpoint(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="Uploaded file must be a Word (.docx) document.")
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Uploaded file is empty.")
    base_name = converters.sanitize_base_name(file.filename)
    content, media_type, out_name, cnt = converters.docx_to_xlsx(data, base_name)
    safe_out_name = converters.sanitize_filename(out_name)
    return Response(
        content=content,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{safe_out_name}"', "X-Table-Count": str(cnt)},
    )


if __name__ == "__main__":
    import os
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run("main:app", host="0.0.0.0", port=port, reload=True)
