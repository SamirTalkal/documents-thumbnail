"""
On-Prem Document Thumbnail Service
----------------------------------

Flow:
1. Client uploads a file (pdf/doc/docx/ppt/pptx).
2. Service:
   - Generates a doc_id.
   - Saves the original file locally (for RAG / retention).
   - Generates a PNG screenshot of the first page/slide.
   - Saves that PNG locally.
3. Responds with:
   - doc_id
   - original_file_url
   - thumbnail_url (frontend can <img src=...>)

Extras:
- File size limit check
- Logging
- /files endpoint to list uploaded docs
- /delete/{doc_id} to clean up
- metadata.json to remember uploads

This file is meant to be portable.
Change only the CONFIG section for different machines/environments.
"""

import os
import uuid
import tempfile
import subprocess
import json
import logging
from datetime import datetime
from typing import Literal

import fitz  # PyMuPDF
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles


# ─────────────────────────────────────────
# CONFIG (edit this part per environment)
# ─────────────────────────────────────────

# Where to store stuff on disk (documents, thumbnails, metadata.json)
BASE_STORAGE_DIR = os.getenv("BASE_STORAGE_DIR", "./storage")
DOCUMENTS_DIR = os.path.join(BASE_STORAGE_DIR, "documents")
THUMBNAILS_DIR = os.path.join(BASE_STORAGE_DIR, "thumbnails")
METADATA_FILE = os.path.join(BASE_STORAGE_DIR, "metadata.json")

# Max allowed upload size in bytes (50 MB default)
MAX_FILE_SIZE = int(os.getenv("MAX_FILE_SIZE", str(50 * 1024 * 1024)))

# Path to LibreOffice CLI. 
# macOS: /Applications/LibreOffice.app/Contents/MacOS/soffice
# Linux: /usr/bin/libreoffice
# Windows: C:\\Program Files\\LibreOffice\\program\\soffice.exe
SOFFICE_CMD = os.getenv(
    "SOFFICE_CMD",
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe" if os.name == 'nt' else "/Applications/LibreOffice.app/Contents/MacOS/soffice"
)

# CORS allowed origins. In prod you can set a specific frontend URL.
CORS_ALLOW_ORIGINS = os.getenv("CORS_ALLOW_ORIGINS", "*").split(",")


# ─────────────────────────────────────────
# Setup filesystem + logging
# ─────────────────────────────────────────

os.makedirs(DOCUMENTS_DIR, exist_ok=True)
os.makedirs(THUMBNAILS_DIR, exist_ok=True)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("thumbnail-service")


# ─────────────────────────────────────────
# App + static mounts
# ─────────────────────────────────────────

app = FastAPI(
    title="On-Prem Document Thumbnail API",
    version="1.0.0",
    description="Upload a doc/pdf/pptx/etc and get a first-page thumbnail"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=CORS_ALLOW_ORIGINS,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Static file serving for browser access:
#   /documents/<doc_id>.<ext>  (original file)
#   /thumbnails/<doc_id>.png   (preview image)
app.mount("/documents", StaticFiles(directory=DOCUMENTS_DIR), name="documents")
app.mount("/thumbnails", StaticFiles(directory=THUMBNAILS_DIR), name="thumbnails")


# ─────────────────────────────────────────
# Metadata helpers
# ─────────────────────────────────────────

def load_metadata():
    """Load metadata (doc_id -> info) from disk."""
    if os.path.exists(METADATA_FILE):
        with open(METADATA_FILE, "r") as f:
            return json.load(f)
    return {}

def save_metadata(metadata):
    """Persist metadata back to disk."""
    with open(METADATA_FILE, "w") as f:
        json.dump(metadata, f, indent=2)

def add_file_metadata(doc_id, filename, extension, uploaded_at):
    metadata = load_metadata()
    metadata[doc_id] = {
        "filename": filename,
        "extension": extension,
        "uploaded_at": uploaded_at
    }
    save_metadata(metadata)

def check_duplicate_filename(filename):
    """Check if a filename already exists in metadata."""
    metadata = load_metadata()
    for doc_id, info in metadata.items():
        if info["filename"] == filename:
            return doc_id, info
    return None, None

def remove_file_metadata(doc_id):
    metadata = load_metadata()
    if doc_id in metadata:
        del metadata[doc_id]
        save_metadata(metadata)


# ─────────────────────────────────────────
# PDF -> PNG (first page only)
# ─────────────────────────────────────────

def pdf_first_page_to_png(pdf_bytes: bytes) -> bytes:
    """
    Convert first page of a PDF (bytes) into PNG bytes.
    """
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception as e:
        logger.error(f"Failed to open PDF: {e}")
        raise HTTPException(
            status_code=400,
            detail="File is not a valid PDF or cannot be opened as PDF."
        )

    if doc.page_count == 0:
        doc.close()
        raise HTTPException(status_code=400, detail="PDF has 0 pages.")

    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # 2x zoom for quality
    png_bytes = pix.tobytes("png")
    doc.close()
    return png_bytes


# ─────────────────────────────────────────
# Office -> PDF via LibreOffice/soffice -> PNG
# ─────────────────────────────────────────

def office_first_page_to_png(
    file_bytes: bytes,
    original_extension: Literal[".doc", ".docx", ".ppt", ".pptx"]
) -> bytes:
    """
    1. Save uploaded Office file to a temp file.
    2. Use soffice (LibreOffice CLI) to convert to PDF.
    3. Render first page of that PDF to PNG bytes.
    4. Cleanup temp files.
    """
    # 1. write to temp file with correct extension
    with tempfile.NamedTemporaryFile(delete=False, suffix=original_extension) as tmp_in:
        tmp_in.write(file_bytes)
        tmp_in_path = tmp_in.name

    # 2. output dir for the converted PDF
    tmp_out_dir = tempfile.mkdtemp()

    try:
        logger.info(f"[{tmp_in_path}] Converting {original_extension} -> PDF with {SOFFICE_CMD}")
        result = subprocess.run(
            [
                SOFFICE_CMD,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", tmp_out_dir,
                tmp_in_path,
            ],
            capture_output=True,
            check=True,
            timeout=30,
        )

        base_no_ext = os.path.splitext(os.path.basename(tmp_in_path))[0]
        guessed_pdf_path = os.path.join(tmp_out_dir, base_no_ext + ".pdf")

        if not os.path.exists(guessed_pdf_path):
            pdf_candidates = [
                os.path.join(tmp_out_dir, f)
                for f in os.listdir(tmp_out_dir)
                if f.lower().endswith(".pdf")
            ]
            if not pdf_candidates:
                logger.error("No PDF produced by soffice")
                raise HTTPException(
                    status_code=500,
                    detail=(
                        "LibreOffice/soffice conversion succeeded but no PDF was found. "
                        "stderr=" + result.stderr.decode("utf-8", "ignore")
                    ),
                )
            guessed_pdf_path = max(pdf_candidates, key=os.path.getmtime)

        with open(guessed_pdf_path, "rb") as pdf_file:
            converted_pdf_bytes = pdf_file.read()

        png_bytes = pdf_first_page_to_png(converted_pdf_bytes)
        logger.info(f"[{tmp_in_path}] Thumbnail generated OK")
        return png_bytes

    except subprocess.TimeoutExpired:
        logger.error("LibreOffice/soffice timed out")
        raise HTTPException(
            status_code=500,
            detail="LibreOffice/soffice conversion timed out."
        )
    except subprocess.CalledProcessError as e:
        logger.error(f"LibreOffice/soffice failed: {e.stderr.decode('utf-8','ignore')}")
        raise HTTPException(
            status_code=500,
            detail=(
                "LibreOffice/soffice failed to convert file to PDF. "
                "stderr=" + e.stderr.decode("utf-8", "ignore")
            ),
        )
    finally:
        # cleanup tmp files/dirs
        if os.path.exists(tmp_in_path):
            try:
                os.unlink(tmp_in_path)
            except OSError:
                pass
        if os.path.isdir(tmp_out_dir):
            for f in os.listdir(tmp_out_dir):
                fp = os.path.join(tmp_out_dir, f)
                try:
                    os.unlink(fp)
                except OSError:
                    pass
            try:
                os.rmdir(tmp_out_dir)
            except OSError:
                pass


# ─────────────────────────────────────────
# Routing logic based on extension
# ─────────────────────────────────────────

def generate_thumbnail_png_bytes(file_bytes: bytes, filename: str) -> bytes:
    lower = filename.lower()

    if lower.endswith(".pdf"):
        return pdf_first_page_to_png(file_bytes)
    if lower.endswith(".docx"):
        return office_first_page_to_png(file_bytes, ".docx")
    if lower.endswith(".doc"):
        return office_first_page_to_png(file_bytes, ".doc")
    if lower.endswith(".pptx"):
        return office_first_page_to_png(file_bytes, ".pptx")
    if lower.endswith(".ppt"):
        return office_first_page_to_png(file_bytes, ".ppt")

    raise HTTPException(
        status_code=400,
        detail={
            "error": "Unsupported file type",
            "filename": filename,
            "supported_formats": ["PDF", "DOC", "DOCX", "PPT", "PPTX"],
            "message": "Please upload a PDF or Office document."
        }
    )


# ─────────────────────────────────────────
# POST /upload
# ─────────────────────────────────────────

@app.post("/upload")
async def upload_document(file: UploadFile = File(...), replace_existing: bool = False):
    """
    1. Read upload.
    2. Check for duplicate filename.
    3. Enforce file size limit.
    4. Generate doc_id.
    5. Save original file to disk.
    6. Generate first-page thumbnail and save it to disk.
    7. Store metadata (for /files list).
    8. Return URLs so frontend can render thumbnail immediately.

    Parameters:
    - file: The uploaded file
    - replace_existing: If True, replace existing file with same name

    Response example:
    {
      "doc_id": "abc123...",
      "original_filename": "slides.pptx",
      "uploaded_at": "2025-10-27T10:15:23.123Z",
      "original_file_url": "/documents/abc123.pptx",
      "thumbnail_url": "/thumbnails/abc123.png",
      "duplicate_replaced": false
    }
    """
    logger.info(f"Upload started: {file.filename}")

    content = await file.read()

    # 1. Validate not empty
    if not content:
        raise HTTPException(status_code=400, detail="Empty file upload.")

    # 2. Check for duplicate filename
    existing_doc_id, existing_info = check_duplicate_filename(file.filename)
    duplicate_replaced = False
    
    if existing_doc_id and not replace_existing:
        raise HTTPException(
            status_code=409,
            detail={
                "error": "File with this name already exists",
                "filename": file.filename,
                "existing_doc_id": existing_doc_id,
                "existing_uploaded_at": existing_info["uploaded_at"],
                "message": "Use replace_existing=true to replace the existing file"
            }
        )
    
    # 3. Validate size
    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(
            status_code=413,
            detail=f"File too large. Max {MAX_FILE_SIZE // (1024*1024)}MB."
        )

    # 4. Handle replacement if needed
    if existing_doc_id and replace_existing:
        # Delete existing files
        existing_ext = existing_info["extension"]
        existing_doc_path = os.path.join(DOCUMENTS_DIR, f"{existing_doc_id}{existing_ext}")
        existing_thumb_path = os.path.join(THUMBNAILS_DIR, f"{existing_doc_id}.png")
        
        if os.path.exists(existing_doc_path):
            os.unlink(existing_doc_path)
        if os.path.exists(existing_thumb_path):
            os.unlink(existing_thumb_path)
        
        remove_file_metadata(existing_doc_id)
        duplicate_replaced = True
        logger.info(f"Replaced existing file: {file.filename}")

    # 5. Create doc_id and derive file paths
    doc_id = uuid.uuid4().hex
    _, ext = os.path.splitext(file.filename)
    ext = ext if ext else ""

    stored_doc_path = os.path.join(DOCUMENTS_DIR, f"{doc_id}{ext}")
    stored_thumb_path = os.path.join(THUMBNAILS_DIR, f"{doc_id}.png")

    # 4. Save original file
    with open(stored_doc_path, "wb") as f:
        f.write(content)
    logger.info(f"Saved original file: {stored_doc_path}")

    # 5. Generate and save thumbnail
    try:
        thumbnail_png_bytes = generate_thumbnail_png_bytes(content, file.filename)
    except HTTPException as e:
        # cleanup doc if thumbnail gen fails
        if os.path.exists(stored_doc_path):
            os.unlink(stored_doc_path)
        raise e

    with open(stored_thumb_path, "wb") as t:
        t.write(thumbnail_png_bytes)
    logger.info(f"Saved thumbnail: {stored_thumb_path}")

    # 6. Metadata
    uploaded_at = datetime.utcnow().isoformat() + "Z"
    add_file_metadata(doc_id, file.filename, ext, uploaded_at)

    # 7. Build response payload
    original_file_url = f"/documents/{doc_id}{ext}"
    thumbnail_url = f"/thumbnails/{doc_id}.png"

    resp = {
        "doc_id": doc_id,
        "original_filename": file.filename,
        "uploaded_at": uploaded_at,
        "original_file_url": original_file_url,
        "thumbnail_url": thumbnail_url,
        "duplicate_replaced": duplicate_replaced
    }

    logger.info(f"Upload completed: doc_id={doc_id}")
    return JSONResponse(resp)


# ─────────────────────────────────────────
# GET /check-duplicate/{filename}
# ─────────────────────────────────────────

@app.get("/check-duplicate/{filename}")
async def check_duplicate(filename: str):
    """
    Check if a filename already exists in the system.
    Returns information about the existing file if found.
    """
    existing_doc_id, existing_info = check_duplicate_filename(filename)
    
    if existing_doc_id:
        return {
            "exists": True,
            "filename": filename,
            "existing_doc_id": existing_doc_id,
            "existing_uploaded_at": existing_info["uploaded_at"],
            "existing_file_url": f"/documents/{existing_doc_id}{existing_info['extension']}",
            "existing_thumbnail_url": f"/thumbnails/{existing_doc_id}.png"
        }
    else:
        return {
            "exists": False,
            "filename": filename,
            "message": "Filename is available for upload"
        }


# ─────────────────────────────────────────
# GET /files  (list all stored docs for UI grid)
# ─────────────────────────────────────────

@app.get("/files")
async def list_files():
    metadata = load_metadata()
    files = []
    for doc_id, info in metadata.items():
        files.append({
            "doc_id": doc_id,
            "filename": info["filename"],
            "uploaded_at": info["uploaded_at"],
            "thumbnail_url": f"/thumbnails/{doc_id}.png",
            "file_url": f"/documents/{doc_id}{info['extension']}"
        })

    return {
        "files": files,
        "total": len(files)
    }


# ─────────────────────────────────────────
# DELETE /delete/{doc_id}
# ─────────────────────────────────────────

@app.delete("/delete/{doc_id}")
async def delete_document(doc_id: str):
    logger.info(f"Delete requested: doc_id={doc_id}")
    removed_paths = []

    # remove original(s)
    for fname in os.listdir(DOCUMENTS_DIR):
        if fname.startswith(doc_id):
            full_path = os.path.join(DOCUMENTS_DIR, fname)
            if os.path.isfile(full_path):
                try:
                    os.remove(full_path)
                    removed_paths.append(full_path)
                    logger.info(f"Deleted: {full_path}")
                except OSError as e:
                    logger.error(f"Failed to delete {full_path}: {e}")

    # remove thumbnail
    thumb_path = os.path.join(THUMBNAILS_DIR, f"{doc_id}.png")
    if os.path.isfile(thumb_path):
        try:
            os.remove(thumb_path)
            removed_paths.append(thumb_path)
            logger.info(f"Deleted: {thumb_path}")
        except OSError as e:
            logger.error(f"Failed to delete {thumb_path}: {e}")

    if not removed_paths:
        raise HTTPException(status_code=404, detail="No files found for that doc_id.")

    # cleanup metadata
    remove_file_metadata(doc_id)

    return {
        "doc_id": doc_id,
        "deleted": True,
        "removed_files": removed_paths
    }


# ─────────────────────────────────────────
# Health / root
# ─────────────────────────────────────────

@app.get("/health")
async def health():
    return {
        "status": "ok",
        "time": datetime.utcnow().isoformat() + "Z",
        "storage_dir": BASE_STORAGE_DIR,
        "total_files": len(load_metadata()),
        "soffice_cmd": SOFFICE_CMD
    }


@app.get("/")
async def root():
    return {
        "service": "On-Prem Document Thumbnail API",
        "version": "1.0.0",
        "endpoints": {
            "POST /upload": "Upload a document and generate thumbnail",
            "GET /files": "List uploaded files + thumbnails",
            "DELETE /delete/{doc_id}": "Delete a file and its thumbnail",
            "GET /thumbnails/{doc_id}.png": "Thumbnail image",
            "GET /documents/{doc_id}.ext": "Original file",
            "GET /health": "Health check"
        },
        "supported_formats": ["PDF", "DOC", "DOCX", "PPT", "PPTX"]
    }


# Run local (for dev)


if __name__ == "__main__":
    import uvicorn
    # host 0.0.0.0 so others on your LAN can hit it if needed
    uvicorn.run(app, host="0.0.0.0", port=8000)
