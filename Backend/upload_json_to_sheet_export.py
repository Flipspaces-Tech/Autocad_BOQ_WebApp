#!/usr/bin/env python3
"""
Uploads AutoCAD export JSONs to Google Sheets (BOM-safe).

Supported JSON schemas:
1) Standard exports:
   A) { "headers": [...], "items": [ {...}, ... ] }
   B) { "items": [ {...}, ... ] }
2) Planner dims export:
   C) { "dwg": "...", "plannerLayer": "...", "overall": {...}, "zones": [ {...}, ... ] }
3) Layer export variants (common):
   D) { "layers": [ {...}, ... ] }
   E) [ {...}, {...} ]  (top-level list)
   F) { "headers": [...], "rows": [ [...], ... ] }
   G) { "LayerA": 12, "LayerB": 5, ... } (dict summary) -> flattened to rows

Mappings:
- vis_export_visibility.json     -> export_visibility tab
- vis_export_all.json            -> export_all tab
- vis_export_sheet_like.json     -> vis_export_sheet_like tab
- vis_export_planner_dims.json   -> PLANNER tab
- vis_export_layers.json         -> LAYER tab
"""

import json
from pathlib import Path
import requests

EXPORTS_DIR = Path(r"C:\Users\admin\Documents\AUTOCAD_WEBAPP\EXPORTS")

FILES = {
    "vis_export_visibility.json": "export_visibility",
    "vis_export_all.json": "export_all",
    "vis_export_sheet_like.json": "vis_export_sheet_like",
    "vis_export_planner_dims.json": "PLANNER",
    "vis_export_layers_sheet_like.json": "LAYER",  # ✅ your request
}

GS_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbwTTg9SzLo70ICTbpr2a5zNw84CG6kylNulVONenq4BADQIuCq7GuJqtDq7H_QfV0pe/exec"
GSHEET_ID = "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM"
MODE = "replace"
TIMEOUT = 180


def read_json_bom_safe(path: Path):
    txt = path.read_text(encoding="utf-8-sig", errors="strict")
    return json.loads(txt)


def _discover_headers_from_list_of_dicts(records):
    headers = []
    for r in records:
        if not isinstance(r, dict):
            continue
        for k in r.keys():
            if k not in headers:
                headers.append(k)
    return headers


def flatten_items_json(data):
    """
    Returns (headers, rows) for upload.
    Handles a wide set of export schemas, including layer exports.
    """

    # -------------------------
    # E) top-level list
    # -------------------------
    if isinstance(data, list):
        records = [x for x in data if isinstance(x, dict)]
        if not records:
            return [], []
        headers = _discover_headers_from_list_of_dicts(records)
        rows = [[r.get(h, "") for h in headers] for r in records]
        return headers, rows

    if not isinstance(data, dict):
        return [], []

    # -------------------------
    # F) headers + rows already
    # -------------------------
    if isinstance(data.get("headers"), list) and isinstance(data.get("rows"), list):
        headers = data["headers"]
        rows = data["rows"]
        # rows might be list-of-lists already
        if rows and isinstance(rows[0], list):
            return headers, rows
        # or list-of-dicts
        if rows and isinstance(rows[0], dict):
            return headers, [[r.get(h, "") for h in headers] for r in rows if isinstance(r, dict)]
        return headers, []

    # -----------------------------
    # A/B) Standard items[] schema
    # -----------------------------
    items = data.get("items", [])
    if isinstance(items, list) and items:
        headers = data.get("headers")
        if isinstance(headers, list) and headers:
            rows = [[it.get(h, "") for h in headers] for it in items if isinstance(it, dict)]
            return headers, rows

        records = [x for x in items if isinstance(x, dict)]
        if not records:
            return [], []
        headers = _discover_headers_from_list_of_dicts(records)
        rows = [[r.get(h, "") for h in headers] for r in records]
        return headers, rows

    # -----------------------------
    # D) Layer export: layers[]
    # -----------------------------
    layers = data.get("layers", [])
    if isinstance(layers, list) and layers:
        records = [x for x in layers if isinstance(x, dict)]
        if not records:
            return [], []
        headers = _discover_headers_from_list_of_dicts(records)
        rows = [[r.get(h, "") for h in headers] for r in records]
        return headers, rows

    # --------------------------------
    # C) Planner dims zones[] schema
    # --------------------------------
    zones = data.get("zones", [])
    if isinstance(zones, list) and zones:
        dwg = data.get("dwg", "")
        planner_layer = data.get("plannerLayer", "")

        overall = data.get("overall", {}) or {}
        overall_w = overall.get("width_ft", "")
        overall_h = overall.get("height_ft", "")

        enriched = []
        for z in zones:
            if not isinstance(z, dict):
                continue
            row = dict(z)
            row["dwg"] = dwg
            row["plannerLayer"] = planner_layer
            row["overall_width_ft"] = overall_w
            row["overall_height_ft"] = overall_h
            enriched.append(row)

        preferred = [
            "name",
            "width_ft",
            "height_ft",
            "area_sqft",
            "dwg",
            "plannerLayer",
            "overall_width_ft",
            "overall_height_ft",
        ]

        present_keys = set()
        for r in enriched:
            present_keys.update(r.keys())

        headers = [k for k in preferred if k in present_keys]
        for k in sorted(present_keys):
            if k not in headers:
                headers.append(k)

        rows = [[r.get(h, "") for h in headers] for r in enriched]
        return headers, rows

    # ----------------------------------------
    # G) dict summary: {"Layer": qty, ...}
    # ----------------------------------------
    # If it's a "simple" dict (values not dict/list), flatten key-value pairs.
    simple = True
    for v in data.values():
        if isinstance(v, (dict, list)):
            simple = False
            break
    if simple and data:
        headers = ["key", "value"]
        rows = [[k, data.get(k, "")] for k in data.keys()]
        return headers, rows

    # nothing usable
    return [], []


def healthcheck():
    """
    Quick reachability check. Many Apps Scripts will return HTML;
    we just confirm it's reachable and show status.
    """
    try:
        r = requests.get(GS_WEBAPP_URL, timeout=30)
        print(f"🌐 WebApp reachability: {r.status_code}")
        # show a small snippet for debugging (avoid huge HTML)
        snip = (r.text or "")[:200].replace("\n", " ")
        print(f"🌐 WebApp response snippet: {snip!r}")
    except Exception as e:
        raise RuntimeError(f"Cannot reach GS WebApp URL. Error: {e}")


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

    # Always show server response when not OK
    if not r.ok:
        raise RuntimeError(
            f"Upload failed ({tab_name}): HTTP {r.status_code}\n"
            f"Response:\n{r.text}"
        )

    # Some webapps reply 200 but with JSON indicating failure.
    # Try parsing JSON if possible.
    ctype = (r.headers.get("content-type") or "").lower()
    if "application/json" in ctype:
        try:
            resp = r.json()
        except Exception:
            resp = None

        # If your GAS returns {ok:false, error:"..."} or similar, catch it.
        if isinstance(resp, dict):
            if resp.get("ok") is False or resp.get("success") is False:
                raise RuntimeError(f"Upload reported failure ({tab_name}): {resp}")
    else:
        # If it's not JSON, still useful to log a short snippet
        snip = (r.text or "")[:250].replace("\n", " ")
        print(f"ℹ️ Upload response (non-JSON) snippet: {snip!r}")


def main():
    print("📦 EXPORTS DIR:", EXPORTS_DIR)
    print("📄 Sheet ID:", GSHEET_ID)
    print("🔗 WebApp URL:", GS_WEBAPP_URL)
    print("🧪 Running healthcheck...")
    healthcheck()
    print("")

    for filename, sheet_tab in FILES.items():
        path = EXPORTS_DIR / filename

        if not path.exists():
            print(f"⚠️ Missing file, skipping: {path}")
            continue

        print(f"⬆ Uploading {filename} → tab '{sheet_tab}'")

        data = read_json_bom_safe(path)
        headers, rows = flatten_items_json(data)

        print(f"   ↳ Detected columns: {len(headers)}")
        print(f"   ↳ Detected rows:    {len(rows)}")

        if not rows:
            print(f"⚠️ No uploadable rows found in {filename}.")
            print("   👉 This usually means the JSON schema is different than expected.")
            print("   👉 Open the JSON and confirm it contains items[] / layers[] / zones[] / rows[].")
            continue

        upload(sheet_tab, headers, rows)
        print(f"✅ Uploaded {len(rows)} rows → {sheet_tab}\n")

    print("🎉 Done")
    print("Sheet:", f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}")


if __name__ == "__main__":
    main()