"""
Stateless Document → Thumbnail API
----------------------------------

✅ Accepts PDF, DOC, DOCX, PPT, PPTX
✅ Converts first page/slide → PNG
✅ Returns Base64 + metadata (no file saved)
✅ Uses LibreOffice for Office → PDF conversion
✅ Fully portable — works on Windows, macOS, Linux
"""

import os
import tempfile
import subprocess
import base64
import hashlib
from datetime import datetime
from typing import Literal

import fitz  # PyMuPDF
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, JSONResponse

app = FastAPI(
    title="Stateless Thumbnail Generator",
    version="1.0",
    description="Upload a document and get first-page thumbnail (base64).",
)

# Allow frontend integration (adjust origins in prod)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace with your frontend domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Path to LibreOffice executable (update per OS)
SOFFICE_CMD = (
    "C:\\Program Files\\LibreOffice\\program\\soffice.exe"  # Windows
    if hasattr(__import__("os"), "name") and __import__("os").name == "nt"
    else "/Applications/LibreOffice.app/Contents/MacOS/soffice"  # macOS
)


# ───────────────────────────────
# PDF → PNG (first page)
# ───────────────────────────────
def pdf_first_page_to_png(pdf_bytes: bytes) -> bytes:
    """Convert first page of PDF to PNG bytes."""
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    except Exception:
        raise HTTPException(status_code=400, detail="Invalid PDF file.")
    if doc.page_count == 0:
        raise HTTPException(status_code=400, detail="PDF has no pages.")
    page = doc.load_page(0)
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))  # Zoom x2 for better quality
    img_bytes = pix.tobytes("png")
    doc.close()
    return img_bytes


# ───────────────────────────────
# Office → PDF → PNG (its for ppt and word docs)
# ───────────────────────────────
def office_first_page_to_png(file_bytes: bytes, ext: Literal[".doc", ".docx", ".ppt", ".pptx"]) -> bytes:
    """Convert Office file to PNG bytes using LibreOffice + PyMuPDF."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp_in:
        tmp_in.write(file_bytes)
        tmp_in_path = tmp_in.name

    tmp_out_dir = tempfile.mkdtemp()
    try:
        result = subprocess.run(
            [SOFFICE_CMD, "--headless", "--convert-to", "pdf", "--outdir", tmp_out_dir, tmp_in_path],
            capture_output=True,
            timeout=40,
            check=True,
        )

        # Locate the converted PDF
        from os import listdir, path
        pdf_candidates = [path.join(tmp_out_dir, f) for f in listdir(tmp_out_dir) if f.lower().endswith(".pdf")]
        if not pdf_candidates:
            raise HTTPException(status_code=500, detail="LibreOffice conversion failed.")
        pdf_path = pdf_candidates[0]

        with open(pdf_path, "rb") as f:
            pdf_bytes = f.read()

        return pdf_first_page_to_png(pdf_bytes)
    except subprocess.TimeoutExpired:
        raise HTTPException(status_code=500, detail="LibreOffice conversion timed out.")
    except subprocess.CalledProcessError as e:
        raise HTTPException(status_code=500, detail=f"LibreOffice error: {e.stderr.decode()}")
    finally:
        import os, shutil
        if os.path.exists(tmp_in_path):
            os.unlink(tmp_in_path)
        shutil.rmtree(tmp_out_dir, ignore_errors=True)


# ───────────────────────────────
# Unified converter
# ───────────────────────────────
def to_png_bytes(file_bytes: bytes, filename: str) -> bytes:
    """Route to correct conversion based on extension."""
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
        detail="Unsupported file type. Supported: PDF, DOCX, DOC, PPTX, PPT",
    )


# ───────────────────────────────
# API Endpoint: /thumbnail-base64
# ───────────────────────────────
@app.post("/thumbnail", summary="Return PNG bytes", response_description="image/png")
async def thumbnail(file: UploadFile = File(...)):
    """
    Request: multipart/form-data with `file`.
    Response: 200 OK with `image/png` body (no persistence).
    """
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Empty file.")
    png = to_png_bytes(data, file.filename)
    return Response(content=png, media_type="image/png")

@app.post("/thumbnail-base64")
async def generate_thumbnail(file: UploadFile = File(...)):
    """
    Upload a document and return Base64 thumbnail with metadata.
    """
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Empty file uploaded.")

    # Convert file → PNG
    png_bytes = to_png_bytes(data, file.filename)

    # Prepare metadata
    encoded = base64.b64encode(png_bytes).decode("utf-8")
    sha256 = hashlib.sha256(png_bytes).hexdigest()

    return {
        "filename": file.filename,
        "mime": "image/png",
        "size_bytes": len(png_bytes),
        "sha256": sha256,
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "png_base64": encoded,
    }



# ───────────────────────────────
# Run local (for dev)
# ───────────────────────────────
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
