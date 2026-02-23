#!/usr/bin/env python3
"""
Upload JSON to Google Sheets (Apps Script Web App) into a NEW tab named: export

✅ What this script does
- Reads a JSON file from disk (your generated output)
- Flattens it into a table
  - If JSON is { ... } → writes key/value rows
  - If JSON is { "items": [ {...}, {...} ] } → writes items as a proper table
- Uploads to Google Sheet tab "export" using your existing Apps Script Web App (same payload style you already use)

Usage (Windows):
  py upload_json_to_sheet_export.py --json "C:\Users\admin\Documents\AUTOCAD_WEBAPP\EXPORTS\vis_export.json"

Optional:
  py upload_json_to_sheet_export.py --json "...\vis_export_visibility.json" --tab export_visibility
  py upload_json_to_sheet_export.py --json "...\vis_export_all.json" --tab export_all
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Tuple

import requests


# ======= CONFIG (set these) =======
GS_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbwTTg9SzLo70ICTbpr2a5zNw84CG6kylNulVONenq4BADQIuCq7GuJqtDq7H_QfV0pe/exec"
GSHEET_ID = "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM"
DEFAULT_TAB = "export"
MODE = "replace"  # replace | append
TIMEOUT_SEC = 300
# =================================


def load_json(p: Path) -> Any:
    raw = p.read_text(encoding="utf-8", errors="replace")
    return json.loads(raw)


def flatten_kv(obj: Dict[str, Any]) -> Tuple[List[str], List[List[str]]]:
    """Simple key/value table."""
    headers = ["key", "value"]
    rows: List[List[str]] = []
    for k, v in obj.items():
        if isinstance(v, (dict, list)):
            v = json.dumps(v, ensure_ascii=False)
        rows.append([str(k), "" if v is None else str(v)])
    return headers, rows


def items_table(obj: Dict[str, Any]) -> Tuple[List[str], List[List[str]]]:
    """
    If obj contains "items": [ {..}, {..} ], write a real table.
    Also supports obj itself being like {..} but items is the key.
    """
    items = obj.get("items")
    if not isinstance(items, list) or not items:
        raise ValueError("No items[] list found")

    # collect all columns across items
    cols: List[str] = []
    colset = set()
    for it in items:
        if isinstance(it, dict):
            for k in it.keys():
                if k not in colset:
                    colset.add(k)
                    cols.append(k)

    # add a few top-level fields as prefix columns (handy)
    prefix_cols = []
    for top_key in ("dwg", "totalInserts", "dynamicProcessed"):
        if top_key in obj:
            prefix_cols.append(top_key)

    headers = prefix_cols + cols

    rows: List[List[str]] = []
    for it in items:
        if not isinstance(it, dict):
            continue
        r: List[str] = []
        for pk in prefix_cols:
            r.append("" if obj.get(pk) is None else str(obj.get(pk)))
        for c in cols:
            v = it.get(c)
            if isinstance(v, (dict, list)):
                v = json.dumps(v, ensure_ascii=False)
            r.append("" if v is None else str(v))
        rows.append(r)

    return headers, rows


def build_sheet_payload(sheet_id: str, tab: str, mode: str, headers: List[str], rows: List[List[str]]) -> dict:
    return {
        "sheetId": sheet_id,
        "tab": tab,
        "mode": mode,
        "headers": headers if mode == "replace" else [],
        "rows": rows,
        "images": [""] * len(rows),
        "bgColors": [""] * len(rows),
        "colorOnly": False,
        "embedImages": False,
        "driveFolderId": "",
        "vAlign": "",
        "sparseAnchor": "last",
    }


def upload_to_sheet(webapp_url: str, payload: dict, timeout: int = TIMEOUT_SEC) -> None:
    r = requests.post(webapp_url, json=payload, timeout=timeout, allow_redirects=True)
    if not r.ok:
        raise RuntimeError(f"Upload failed: HTTP {r.status_code}\n{r.text[:2000]}")
    print(f"✅ Uploaded to sheet tab '{payload.get('tab')}' ({len(payload.get('rows') or [])} rows)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--json", required=True, help="Path to JSON file to upload")
    ap.add_argument("--sheet-id", default=GSHEET_ID)
    ap.add_argument("--webapp-url", default=GS_WEBAPP_URL)
    ap.add_argument("--tab", default=DEFAULT_TAB)
    ap.add_argument("--mode", choices=["replace", "append"], default=MODE)
    ap.add_argument(
        "--format",
        choices=["auto", "items", "kv"],
        default="auto",
        help="auto: if items[] exists → table; else key/value",
    )
    args = ap.parse_args()

    p = Path(args.json)
    if not p.exists():
        raise SystemExit(f"❌ JSON not found: {p}")

    obj = load_json(p)

    # Decide format
    headers: List[str]
    rows: List[List[str]]

    if args.format == "items":
        headers, rows = items_table(obj if isinstance(obj, dict) else {"items": obj})
    elif args.format == "kv":
        if not isinstance(obj, dict):
            obj = {"value": obj}
        headers, rows = flatten_kv(obj)
    else:
        # auto
        if isinstance(obj, dict) and isinstance(obj.get("items"), list):
            headers, rows = items_table(obj)
        else:
            if not isinstance(obj, dict):
                obj = {"value": obj}
            headers, rows = flatten_kv(obj)

    payload = build_sheet_payload(args.sheet_id, args.tab, args.mode, headers, rows)
    upload_to_sheet(args.webapp_url, payload)


if __name__ == "__main__":
    main()
