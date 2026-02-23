#!/usr/bin/env python3
"""
Uploads AutoCAD export JSONs to Google Sheets (BOM-safe).
- vis_export_visibility.json -> export_visibility tab
- vis_export_all.json        -> export_all tab
"""

import json
from pathlib import Path
import requests

EXPORTS_DIR = Path(r"C:\Users\admin\Documents\AUTOCAD_WEBAPP\EXPORTS")

FILES = {
    "vis_export_visibility.json": "export_visibility",
    "vis_export_all.json": "export_all",
    "vis_export_sheet_like.json": "vis_export_sheet_like",
}



GS_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbwTTg9SzLo70ICTbpr2a5zNw84CG6kylNulVONenq4BADQIuCq7GuJqtDq7H_QfV0pe/exec"
GSHEET_ID = "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM"
MODE = "replace"
TIMEOUT = 180


def flatten_items_json(data: dict):
    items = data.get("items", [])
    if not items:
        return [], []

    # ✅ if file provides headers, use them in that exact order
    headers = data.get("headers")
    if headers and isinstance(headers, list) and len(headers) > 0:
        rows = [[it.get(h, "") for h in headers] for it in items]
        return headers, rows

    # fallback: discover headers from items
    headers = []
    for it in items:
        for k in it.keys():
            if k not in headers:
                headers.append(k)

    rows = [[it.get(h, "") for h in headers] for it in items]
    return headers, rows


def upload(tab_name: str, headers, rows):
    payload = {
        "sheetId": GSHEET_ID,
        "tab": tab_name,
        "mode": MODE,
        "headers": headers,
        "rows": rows,
        "images": [""] * len(rows),
        "bgColors": [""] * len(rows),
        "colorOnly": False,
        "embedImages": False,
        "driveFolderId": "",
        "vAlign": "",
        "sparseAnchor": "last",
    }
    r = requests.post(GS_WEBAPP_URL, json=payload, timeout=TIMEOUT)
    if not r.ok:
        raise RuntimeError(f"Upload failed ({tab_name}): {r.status_code}\n{r.text}")


def read_json_bom_safe(path: Path) -> dict:
    # ✅ BOM-safe read
    txt = path.read_text(encoding="utf-8-sig", errors="strict")
    return json.loads(txt)


def main():
    print("📦 EXPORTS DIR:", EXPORTS_DIR)

    for filename, sheet_tab in FILES.items():
        path = EXPORTS_DIR / filename

        if not path.exists():
            print(f"⚠️ Missing file, skipping: {path}")
            continue

        print(f"⬆ Uploading {filename} → tab '{sheet_tab}'")

        data = read_json_bom_safe(path)
        headers, rows = flatten_items_json(data)

        if not rows:
            print(f"⚠️ No items[] found in {filename}, skipping upload")
            continue

        upload(sheet_tab, headers, rows)
        print(f"✅ Uploaded {len(rows)} rows → {sheet_tab}")

    print("🎉 Done")
    print("Sheet:", f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}")


if __name__ == "__main__":
    main()
