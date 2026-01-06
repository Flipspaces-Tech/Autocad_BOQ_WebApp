# backend/app.py

import io
from typing import Any, Dict

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse


from tr_2 import process_cad_from_upload  # import from wherever you placed it

# 👇 import your DXF → Sheets pipeline function from tr_2.py
from tr_2 import process_doc_from_stream

from dotenv import load_dotenv
load_dotenv()


app = FastAPI(title="AutoCAD BOQ Web API")

# CORS: allow your React dev server / Vercel to call this API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # you can restrict later
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# app = FastAPI()

@app.post("/process-cad")
async def process_cad(file: UploadFile = File(...)):
    try:
        data = await file.read()
        return process_cad_from_upload(file.filename, data)
    except Exception as e:
        raise HTTPException(status_code=400, detail=str(e))


async def _run_pipeline_from_upload(file: UploadFile) -> Dict[str, Any]:
  """
  Reads the uploaded DXF and calls process_doc_from_stream()
  from tr_2.py, which writes to Google Sheets and returns a summary.
  """

  filename = file.filename or "drawing"
  if not filename.lower().endswith(".dxf"):
      # Your pipeline is DXF-only, so be explicit
      raise HTTPException(
          status_code=400,
          detail="Only ASCII DXF files are supported by this pipeline."
      )

  # Read file bytes
  contents = await file.read()
  if not contents:
      raise HTTPException(status_code=400, detail="Empty file upload")

  # Quick binary DXF sniff (your tr_2.py probably does more)
  head = contents[:64]
  if b"AutoCAD Binary DXF" in head:
      raise HTTPException(
          status_code=415,
          detail="Binary DXF not supported. Please export ASCII DXF."
      )

  # Decode text
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

  # Call your big processing function from tr_2.py
  try:
      summary = process_doc_from_stream(io.StringIO(text))
  except ValueError as ve:
      # e.g. invalid DXF
      raise HTTPException(status_code=415, detail=str(ve))
  except Exception as ex:
      raise HTTPException(status_code=500, detail=f"Server error: {ex}")

  # Summary is whatever tr_2 returns; we shape it for frontend
  gsheet_id = summary.get("gsheet_id")
  sheet_tab = summary.get("sheet_tab")

  # Build a Google Sheets URL if we have the ID
  sheet_url = (
      f"https://docs.google.com/spreadsheets/d/{gsheet_id}"
      if gsheet_id else None
  )

  return {
      "ok": True,
      "message": "BOQ generated and pushed to Google Sheets.",
      "sheetUrl": sheet_url,             # 👈 React uses this
      "sheetName": sheet_tab,            # 👈 React shows this as tag
      "uploadId": summary.get("upload_id") or gsheet_id,
      "rawSummary": summary,             # Optional extra info
  }


@app.post("/upload")
async def upload_drawing(file: UploadFile = File(...)):
  try:
    data = await file.read()
    return JSONResponse(process_cad_from_upload(file.filename or "upload.dwg", data))
  except Exception as e:
    raise HTTPException(status_code=400, detail=str(e))

@app.post("/process")
async def process_drawing(file: UploadFile = File(...)):
  try:
    data = await file.read()
    return JSONResponse(process_cad_from_upload(file.filename or "upload.dwg", data))
  except Exception as e:
    raise HTTPException(status_code=400, detail=str(e))



# Simple health check
@app.get("/health")
def health():
  return {"status": "ok"}
