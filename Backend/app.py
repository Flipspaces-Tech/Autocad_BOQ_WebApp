# app.py  (Backend)

import os
import shutil
from typing import List, Any, Dict

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import requests

# ====== CONFIG ======

# Folder where uploaded DWG/DXF files will be stored temporarily
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# TODO: replace this with your real Apps Script Web App URL
APP_SCRIPT_URL = "https://script.google.com/macros/s/YOUR_SCRIPT_ID_HERE/exec"


# ====== FASTAPI APP SETUP ======

app = FastAPI(title="AutoCAD BOQ Web API")

# Allow frontend (React) to call this API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # in production, restrict to your domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ====== HELPER: YOUR BOQ PIPELINE WRAPPER ======

def run_boq_pipeline(file_path: str) -> (List[List[Any]], Dict[str, Any]):
    """
    This is the wrapper around your existing DXF/DWG → BOQ logic.

    For now it's a dummy implementation so the flow works end-to-end.
    Later, replace the dummy code with imports from VIZ_AUTOCAD_NEW and
    call your real functions there.

    Must return:
      rows: list of lists for Google Sheets
      meta: dict with metadata (projectName, etc.)
    """

    # ---- TODO: REPLACE THIS WITH YOUR REAL LOGIC ----
    header = ["Category", "Item", "Description", "Qty", "Unit", "Remarks"]
    rows = [header]

    # Example dummy BOQ row
    rows.append([
        "Seating",
        "Workstation Chair",
        "Standard task chair",
        42,
        "Nos",
        f"Dummy data for file {os.path.basename(file_path)}",
    ])

    meta = {
        "projectName": "Demo Project",
        "drawingFile": os.path.basename(file_path),
        "carpetAreaSft": 9500,
        "floor": "4F",
        "runTimestamp": "2025-11-21T13:45:00+05:30",
        "uploadId": os.path.splitext(os.path.basename(file_path))[0],
    }
    # ---- END DUMMY ----

    return rows, meta


# ====== HELPER: SEND TO APPS SCRIPT ======

def send_boq_to_apps_script(rows: List[List[Any]], meta: Dict[str, Any]) -> Dict[str, Any]:
    if not APP_SCRIPT_URL or "YOUR_SCRIPT_ID_HERE" in APP_SCRIPT_URL:
        raise RuntimeError("APP_SCRIPT_URL is not configured. Edit app.py and set it correctly.")

    payload = {
        "action": "writeBoq",
        "uploadId": meta.get("uploadId"),
        "rows": rows,
        "meta": meta,
    }

    try:
        res = requests.post(APP_SCRIPT_URL, json=payload, timeout=60)
    except Exception as e:
        raise RuntimeError(f"Failed to reach Apps Script: {e}")

    if not res.ok:
        raise RuntimeError(f"Apps Script HTTP error {res.status_code}: {res.text}")

    try:
        data = res.json()
    except Exception:
        raise RuntimeError(f"Apps Script returned non-JSON response: {res.text}")

    if not data.get("ok"):
        raise RuntimeError(f"Apps Script error: {data}")

    return data


# ====== API ENDPOINT ======

@app.post("/upload")
async def upload_drawing(file: UploadFile = File(...)):
    # Basic extension check
    filename = file.filename or "drawing"
    if not filename.lower().endswith((".dwg", ".dxf")):
        raise HTTPException(status_code=400, detail="Only DWG/DXF files are supported")

    # Save file to temporary folder
    save_path = os.path.join(UPLOAD_DIR, filename)
    try:
        with open(save_path, "wb") as f:
            shutil.copyfileobj(file.file, f)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save upload: {e}")

    try:
        # 1) Run your BOQ pipeline on this file
        rows, meta = run_boq_pipeline(save_path)

        # 2) Send the result to Apps Script → Google Sheets
        result = send_boq_to_apps_script(rows, meta)

        # 3) Return a clean response for the React frontend
        return JSONResponse({
            "ok": True,
            "message": "BOQ generated successfully",
            "sheetUrl": result.get("sheetUrl"),
            "sheetName": result.get("sheetName"),
            "uploadId": result.get("uploadId"),
        })

    except Exception as e:
        # Surface any pipeline or Apps Script errors
        raise HTTPException(status_code=500, detail=f"Processing failed: {e}")
    finally:
        # Optional cleanup
        try:
            if os.path.exists(save_path):
                os.remove(save_path)
        except Exception:
            pass


# ====== SIMPLE HEALTH CHECK ======

@app.get("/health")
def health():
    return {"status": "ok"}
