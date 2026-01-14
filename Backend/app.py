# backend/app.py

import io
import time
import logging
from typing import Any, Dict

from fastapi import FastAPI, UploadFile, File, HTTPException, Form, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from dotenv import load_dotenv
load_dotenv()

from tr_2 import process_cad_from_upload
from tr_2 import process_doc_from_stream


# -----------------------
# App + logging
# -----------------------
app = FastAPI(title="AutoCAD BOQ Web API")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s: %(message)s"
)

# -----------------------
# CORS (FIXED)
# IMPORTANT:
# - Don't use allow_origins=["*"] with allow_credentials=True
# - Add your exact Vercel + localhost origins
# -----------------------
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "http://localhost:3000",
        "http://127.0.0.1:3000",
        "https://autocadwebapp.vercel.app",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# -----------------------
# Helpers
# -----------------------
async def _run_pipeline_from_upload(file: UploadFile) -> Dict[str, Any]:
    """
    Reads uploaded ASCII DXF and calls process_doc_from_stream()
    (DXF-only path)
    """
    filename = file.filename or "drawing"
    if not filename.lower().endswith(".dxf"):
        raise HTTPException(
            status_code=400,
            detail="Only ASCII DXF files are supported by this pipeline."
        )

    contents = await file.read()
    if not contents:
        raise HTTPException(status_code=400, detail="Empty file upload")

    head = contents[:64]
    if b"AutoCAD Binary DXF" in head:
        raise HTTPException(
            status_code=415,
            detail="Binary DXF not supported. Please export ASCII DXF."
        )

    try:
        text = contents.decode("utf-8")
    except UnicodeDecodeError:
        try:
            text = contents.decode("latin-1")
        except Exception:
            raise HTTPException(
                status_code=400,
                detail="Unable to decode DXF text (utf-8/latin-1)."
            )

    try:
        summary = process_doc_from_stream(io.StringIO(text))
    except ValueError as ve:
        raise HTTPException(status_code=415, detail=str(ve))
    except Exception as ex:
        raise HTTPException(status_code=500, detail=f"Server error: {ex}")

    gsheet_id = summary.get("gsheet_id")
    sheet_tab = summary.get("sheet_tab")
    sheet_url = f"https://docs.google.com/spreadsheets/d/{gsheet_id}" if gsheet_id else None

    return {
        "ok": True,
        "message": "BOQ generated and pushed to Google Sheets.",
        "sheetUrl": sheet_url,
        "sheetName": sheet_tab,
        "uploadId": summary.get("upload_id") or gsheet_id,
        "rawSummary": summary,
    }


# -----------------------
# Core endpoint (instrumented)
# This is the one your Vercel app calls:
# POST https://autocad-boq-webapp.onrender.com/process-cad
# -----------------------
@app.post("/process-cad")
async def process_cad(
    request: Request,
    file: UploadFile = File(...),
    settings: str = Form("")  # FormData.append("settings", ...)
):
    t0 = time.time()
    origin = request.headers.get("origin", "")
    ua = request.headers.get("user-agent", "")
    try:
        data = await file.read()
        logging.info(
            "START /process-cad origin=%s bytes=%d filename=%s ua=%s",
            origin, len(data), file.filename, ua[:80]
        )

        out = process_cad_from_upload(
            file.filename or "upload.dwg",
            data,
            settings_json=settings
        )

        logging.info(
            "END   /process-cad origin=%s took=%.1fs",
            origin, (time.time() - t0)
        )
        return JSONResponse(out)

    except Exception as e:
        logging.exception(
            "ERR   /process-cad origin=%s took=%.1fs err=%s",
            origin, (time.time() - t0), str(e)
        )
        raise HTTPException(status_code=400, detail=str(e))


# -----------------------
# Optional aliases (if your frontend calls these)
# -----------------------
@app.post("/upload")
async def upload_drawing(
    request: Request,
    file: UploadFile = File(...),
    settings: str = Form("")
):
    # Just route to the same logic
    return await process_cad(request=request, file=file, settings=settings)


@app.post("/process")
async def process_drawing(
    request: Request,
    file: UploadFile = File(...),
    settings: str = Form("")
):
    # Just route to the same logic
    return await process_cad(request=request, file=file, settings=settings)


# -----------------------
# Optional DXF-stream pipeline endpoint
# -----------------------
@app.post("/process-dxf-stream")
async def process_dxf_stream(file: UploadFile = File(...)):
    return JSONResponse(await _run_pipeline_from_upload(file))


# -----------------------
# Health
# -----------------------
@app.get("/health")
def health():
    return {"status": "ok"}
