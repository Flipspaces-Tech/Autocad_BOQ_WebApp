# backend/app.py

import io
from typing import Any, Dict

from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from dotenv import load_dotenv
load_dotenv()

from tr_2 import process_cad_from_upload
from tr_2 import process_doc_from_stream


app = FastAPI(title="AutoCAD BOQ Web API")

# CORS: allow your React dev server / Vercel to call this API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # restrict later if needed
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/process-cad")
async def process_cad(
    file: UploadFile = File(...),
    settings: str = Form("")  # ✅ receives FormData.append("settings", ...)
):
    try:
        data = await file.read()
        return JSONResponse(
            process_cad_from_upload(file.filename or "upload.dwg", data, settings_json=settings)
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


async def _run_pipeline_from_upload(file: UploadFile) -> Dict[str, Any]:
    """
    Reads the uploaded DXF and calls process_doc_from_stream()
    from tr_2.py, which writes to Google Sheets and returns a summary.
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


# OPTIONAL: expose the DXF-stream pipeline explicitly (only if you actually use it)
@app.post("/process-dxf-stream")
async def process_dxf_stream(file: UploadFile = File(...)):
    return JSONResponse(await _run_pipeline_from_upload(file))


@app.post("/upload")
async def upload_drawing(
    file: UploadFile = File(...),
    settings: str = Form("")
):
    try:
        data = await file.read()
        return JSONResponse(
            process_cad_from_upload(file.filename or "upload.dwg", data, settings_json=settings)
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.post("/process")
async def process_drawing(
    file: UploadFile = File(...),
    settings: str = Form("")
):
    try:
        data = await file.read()
        return JSONResponse(
            process_cad_from_upload(file.filename or "upload.dwg", data, settings_json=settings)
        )
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


@app.get("/health")
def health():
    return {"status": "ok"}
