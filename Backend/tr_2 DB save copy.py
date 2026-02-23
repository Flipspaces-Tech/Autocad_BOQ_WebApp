#!/usr/bin/env python3
# DXF → CSV + Google Sheets (Apps Script Web App)
# - Aggregates INSERTs by (block_name, category(upload layer), category1(original dwg layer), zone) using median bbox (L/W)
# - Sends category1 = original DWG layer (lowercased) to Detail
# - Zone detection from PLANNER (INSERT bbox or closed LWPOLYLINE + nearest/inside label)
# - Layer totals with dominant color vote → ByLayer tab (color swatches)
# - Detail tab: Apps Script reorders columns and can merge zones under same (category1 + BOQ name)
#
# IMPORTANT:
# - Keep "headers" names stable: category1, BOQ name, zone, qty_type, qty_value, length (ft), width (ft), Description, Preview, remarks
# - Apps Script handles the final shaping/merging.

from __future__ import annotations

import argparse
import base64
import csv
import io
import json
import logging
import math
import os
import tempfile
import time
import uuid
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import numpy as np
import requests

import re

import ezdxf
from ezdxf import colors as ezcolors
from ezdxf import recover
from ezdxf.addons.drawing import Frontend, RenderContext
from ezdxf.addons.drawing.matplotlib import MatplotlibBackend
from contextlib import contextmanager

from collections import Counter

# --- OpenCV (installed via opencv-python-headless in requirements.txt) ---
try:
    import cv2
except ModuleNotFoundError:
    cv2 = None
    logging.warning("cv2 not installed. Previews will be disabled unless you add opencv-python-headless.")

# Matplotlib headless
os.environ.setdefault("MPLCONFIGDIR", "/tmp/matplotlib")
os.makedirs(os.environ["MPLCONFIGDIR"], exist_ok=True)
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt


# ===== Defaults you can edit =====
DXF_FOLDER = r"C:\Users\admin\Documents\AUTOCAD_WEBAPP\DXF"
OUT_ROOT   = r"C:\Users\admin\Documents\AUTOCAD_WEBAPP\EXPORTS"

GS_WEBAPP_URL       = "https://script.google.com/macros/s/AKfycbwTTg9SzLo70ICTbpr2a5zNw84CG6kylNulVONenq4BADQIuCq7GuJqtDq7H_QfV0pe/exec"
GSHEET_ID           = "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM"
GSHEET_TAB          = "t1"            # Detail tab
GSHEET_SUMMARY_TAB  = ""                # blank → auto "<GSHEET_TAB>_ByLayer"
GSHEET_MODE         = "replace"
GS_DRIVE_FOLDER_ID  = ""

# Which attribute tags we consider as "description"
DESC_TAGS = {"DESC", "DESCRIPTION", "NOTE", "REM", "REMARK", "INFO", "META_DESC"}

ENABLE_PREVIEWS = True
if cv2 is None:
    ENABLE_PREVIEWS = False


# =========================
# ONE.PY PREVIEW ENGINE
# =========================
PREVIEW_TARGET_SIZE = 128
PREVIEW_DPI         = 240
PREVIEW_PAD_PCT     = 0.04
PREVIEW_MARGIN_PCT  = 0.10
PREVIEW_SUPERSAMPLE = 1

PREVIEW_SKIP_LAYERS    = {"DEFPOINTS","DIM","DIMENSIONS","ANNOTATION","TEXT","NOTES"}
PREVIEW_SKIP_DXF_TYPES = {"TEXT","MTEXT","DIMENSION"}

PREVIEW_INK_THRESHOLD     = 100
PREVIEW_THICKEN_PX        = 2
PREVIEW_THICKEN_ITER      = 2
PREVIEW_CLOSE_GAPS_KSIZE  = 1

PREVIEW_EDGE_BG_THRESH    = 250
PREVIEW_MIN_VISIBLE_ALPHA = 8

ENABLE_NESTED_INSERTS = True
NESTED_MAX_DEPTH = 10


PREVIEW_STROKE = os.getenv("PREVIEW_STROKE", "black").strip().lower()
if PREVIEW_STROKE not in ("black","white"):
    PREVIEW_STROKE = "black"


# ===== Headers =====
CSV_HEADERS = [
    "entity_type","category","zone","category1",
    "BOQ name","qty_type","qty_value",
    "length (ft)","width (ft)","perimeter","area (ft2)",
    "Description",
    "Preview","remarks"
]

DETAIL_HEADERS = [
    "entity_type","category","zone","category1",
    "BOQ name","qty_type","qty_value",
    "length (ft)","width (ft)",
    "Description",
    "Preview",
    "remarks"
]

LAYER_HEADERS = [
    "category",
    "zone",
    "length (ft)",
    "width (ft)",
    "perimeter",
    "area (ft2)",
    "Preview",
]

# ===== Formatting & switches =====
DEC_PLACES = 2
FORCE_PLANNER_CATEGORY = True  # If zone exists, send "category" as PLANNER in Detail


def _parse_settings_json(settings_json: str) -> dict:
    if not settings_json:
        return {}
    try:
        if isinstance(settings_json, (dict, list)):
            return settings_json
        return json.loads(settings_json)
    except Exception:
        logging.warning("Invalid settings JSON, ignoring.")
        return {}

def _to_int(v, default=None):
    try:
        if v is None or v == "":
            return default
        return int(float(v))
    except Exception:
        return default

def _to_float(v, default=None):
    try:
        if v is None or v == "":
            return default
        return float(v)
    except Exception:
        return default

@contextmanager
def _apply_runtime_overrides(cfg: dict):
    """
    Quick override mechanism to control render + sheet settings per request.
    NOTE: This mutates module globals (not ideal for high concurrency).
    """
    keys = {
        # sheet
        "GSHEET_ID": "gsheetId",
        "GSHEET_TAB": "gsheetTab",

        # preview params
        "PREVIEW_TARGET_SIZE": "previewTargetSize",
        "PREVIEW_DPI": "previewDpi",
        "PREVIEW_INK_THRESHOLD": "previewInkThreshold",
        "PREVIEW_THICKEN_PX": "previewThickenPx",
        "PREVIEW_THICKEN_ITER": "previewThickenIter",
        "PREVIEW_CLOSE_GAPS_KSIZE": "previewCloseGapsKsize",
        "PREVIEW_EDGE_BG_THRESH": "previewEdgeBgThresh",
        "PREVIEW_MIN_VISIBLE_ALPHA": "previewMinVisibleAlpha",
        "PREVIEW_PAD_PCT": "previewPadPct",
        "PREVIEW_MARGIN_PCT": "previewMarginPct",
        "PREVIEW_SUPERSAMPLE": "previewSupersample",
    }

    old = {}
    try:
        for py_name, cfg_name in keys.items():
            if cfg_name in cfg and cfg[cfg_name] is not None and cfg[cfg_name] != "":
                old[py_name] = globals().get(py_name)

                # cast properly
                if py_name in ("PREVIEW_TARGET_SIZE","PREVIEW_DPI","PREVIEW_INK_THRESHOLD",
                               "PREVIEW_THICKEN_PX","PREVIEW_THICKEN_ITER","PREVIEW_CLOSE_GAPS_KSIZE",
                               "PREVIEW_EDGE_BG_THRESH","PREVIEW_MIN_VISIBLE_ALPHA","PREVIEW_SUPERSAMPLE"):
                    globals()[py_name] = _to_int(cfg[cfg_name], globals().get(py_name))
                elif py_name in ("PREVIEW_PAD_PCT","PREVIEW_MARGIN_PCT"):
                    globals()[py_name] = _to_float(cfg[cfg_name], globals().get(py_name))
                else:
                    globals()[py_name] = cfg[cfg_name]

        yield
    finally:
        for k, v in old.items():
            globals()[k] = v


# ===== Utilities =====
def layer_or_misc(name: str) -> str:
    s = (name or "").strip()
    return s if s else "misc"



def _coerce_int(v):
    try:
        if v is None:
            return None
        if isinstance(v, bool):
            return int(v)
        if isinstance(v, (int, float)):
            return int(round(float(v)))
        s = str(v).strip()
        if not s:
            return None
        if s.replace(".", "", 1).isdigit():
            return int(round(float(s)))
    except Exception:
        pass
    return None

def chair_count_from_visibility_string(vis: str) -> int:
    if not vis:
        return 0
    s = str(vis).upper().strip()

    # Any standalone number
    m = re.search(r"\b(\d{1,2})\b", s)
    if m:
        return int(m.group(1))

    # C8 / CH8 / S8 patterns
    m = re.search(r"\b(?:C|CH|S)(\d{1,2})\b", s)
    if m:
        return int(m.group(1))

    return 0



def get_dynamic_props(ins) -> dict:
    """
    Best-effort extraction of dynamic block properties from INSERT's extension dictionary.
    Returns {propName: value}.
    """
    props = {}
    try:
        if not ins.has_extension_dict:
            return props

        ed = ins.get_extension_dict()
        for k, handle in ed.items():
            try:
                obj = ed.doc.entitydb.get(handle)
                if obj is None or obj.dxftype() != "XRECORD":
                    continue

                tags = list(obj.tags)

                for i, t in enumerate(tags):
                    if isinstance(t.value, str):
                        name = t.value.strip()
                        if not name:
                            continue

                        val = None

                        # Look ahead for numeric OR string values
                        for j in range(i + 1, min(i + 10, len(tags))):
                            v = tags[j].value

                            # numeric
                            if isinstance(v, (int, float)):
                                val = v
                                break

                            # numeric stored as string
                            if isinstance(v, str):
                                s = v.strip()
                                if not s:
                                    continue
                                if s.replace(".", "", 1).isdigit():
                                    val = float(s) if "." in s else int(s)
                                    break

                                # ALSO allow non-numeric strings (Visibility labels like "8 CHAIRS")
                                # Stop at first meaningful string value
                                if len(s) <= 60:
                                    val = s
                                    break

                        if val is not None:
                            props.setdefault(name, val)

            except Exception:
                pass
    except Exception:
        pass

    return props

def units_scale_to_meters(doc, unitless_units: str = "m") -> float:
    try:
        code = int(doc.header.get("$INSUNITS", 0) or 0)
    except Exception:
        code = 0
    mapping = {1:0.0254, 2:0.3048, 4:0.001, 5:0.01, 6:1.0}
    if code in mapping:
        scale = mapping[code]
        logging.info("Detected $INSUNITS=%s → %s m/unit", code, scale)
        return scale
    unitless_map = {"m":1.0, "cm":0.01, "mm":0.001, "in":0.0254, "ft":0.3048}
    scale = unitless_map.get((unitless_units or "m").lower().strip(), 1.0)
    logging.warning("Unitless/unknown $INSUNITS=%s. Assuming %s per unit → %s m/unit.", code, unitless_units, scale)
    return scale

def to_target_units(v_m: float, target: str, kind: str) -> float:
    t = (target or "m").lower().strip()
    if kind == "length":
        return {"m":v_m, "mm":v_m*1000, "cm":v_m*100, "ft":v_m/0.3048}.get(t, v_m)
    return {"m":v_m, "mm":v_m*1_000_000, "cm":v_m*10_000, "ft":v_m/(0.3048**2)}.get(t, v_m)

def dist2d(p1, p2) -> float:
    return math.hypot(p2[0]-p1[0], p2[1]-p1[1])

def polyline_length_xy(pts: list[tuple[float,float]], closed: bool) -> float:
    if len(pts) < 2: return 0.0
    L = sum(dist2d(pts[i], pts[i+1]) for i in range(len(pts)-1))
    if closed: L += dist2d(pts[-1], pts[0])
    return L

def polygon_area_xy(pts: list[tuple[float,float]]) -> float:
    n = len(pts)
    if n < 3: return 0.0
    s = 0.0
    for i in range(n):
        x1,y1 = pts[i]; x2,y2 = pts[(i+1)%n]
        s += x1*y2 - x2*y1
    return abs(s)*0.5

def _fmt_num(val, places: int | None = None) -> str:
    if val is None or val == "": return ""
    try:
        p = DEC_PLACES if places is None else places
        num = float(str(val).strip())
        return f"{num:.{p}f}"
    except Exception:
        return ""

def _rgb_to_hex(rgb: tuple[int,int,int]) -> str:
    r, g, b = rgb
    return "#{:02X}{:02X}{:02X}".format(r, g, b)


# ===== Curve sampling helpers =====
def _sample_arc_pts(cx, cy, r, start_deg: Optional[float], end_deg: Optional[float]):
    if r <= 0: return []
    if start_deg is None or end_deg is None:
        start_deg, end_deg = 0.0, 360.0
    sweep = (end_deg - start_deg) % 360.0
    steps = max(16, int(max(16, sweep/6.0)))
    for i in range(steps+1):
        a = math.radians(start_deg + sweep*(i/steps))
        yield (cx + r*math.cos(a), cy + r*math.sin(a))

def _bulge_arc_points(p1, p2, bulge: float, min_steps: int = 8):
    if abs(bulge) < 1e-12: return [p1, p2]
    x1,y1 = p1; x2,y2 = p2
    dx, dy = x2-x1, y2-y1
    c = math.hypot(dx, dy)
    if c < 1e-12: return [p1]
    theta = 4.0 * math.atan(bulge)
    s_half = 2*bulge / (1 + bulge*bulge)
    if abs(s_half) < 1e-12: return [p1, p2]
    R = c / (2.0 * s_half)
    nx, ny = (-dy/c, dx/c)
    cos_half = (1 - bulge*bulge) / (1 + bulge*bulge)
    d = R * cos_half
    mx, my = (x1+x2)/2.0, (y1+y2)/2.0
    cx, cy = mx + nx*d, my + ny*d
    a1 = math.atan2(y1 - cy, x1 - cx)
    a2 = math.atan2(y2 - cy, x2 - cx)
    raw_ccw = (a2 - a1) % (2*math.pi)
    sweep = raw_ccw if theta >= 0 else raw_ccw - 2*math.pi
    steps = max(min_steps, int(abs(sweep) / (6*math.pi/180)))
    return [(cx + R*math.cos(a1 + sweep*(i/steps)),
             cy + R*math.sin(a1 + sweep*(i/steps))) for i in range(steps+1)]

def _collect_points_from_entity(e):
    et = e.dxftype()
    if et == "LINE":
        yield (float(e.dxf.start.x), float(e.dxf.start.y))
        yield (float(e.dxf.end.x), float(e.dxf.end.y))
    elif et == "LWPOLYLINE":
        verts = list(e); n = len(verts)
        if n == 0: return
        closed = bool(getattr(e, "closed", False))
        for i in range(n if closed else n-1):
            j = (i + 1) % n
            x1,y1 = float(verts[i][0]), float(verts[i][1])
            x2,y2 = float(verts[j][0]), float(verts[j][1])
            try: b = float(verts[i][4])
            except Exception: b = 0.0
            for p in _bulge_arc_points((x1,y1),(x2,y2),b)[:-1]:
                yield p
        yield (float(verts[-1][0]), float(verts[-1][1]))
        if closed:
            yield (float(verts[0][0]), float(verts[0][1]))
    elif et == "POLYLINE":
        vs = list(e.vertices())
        if not vs: return
        pts = []
        for v in vs:
            loc = getattr(v.dxf, "location", None)
            pts.append((float(loc.x), float(loc.y)) if loc is not None
                       else (float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
        closed = bool(getattr(e,"is_closed",getattr(e,"closed",False)))
        n = len(pts)
        for i in range(n - (0 if closed else 1)):
            j = (i + 1) % n
            try: b = float(vs[i].dxf.bulge)
            except Exception: b = 0.0
            for p in _bulge_arc_points(pts[i], pts[j], b)[:-1]:
                yield p
        yield pts[-1]
        if closed:
            yield pts[0]
    elif et == "CIRCLE":
        cx, cy = float(e.dxf.center.x), float(e.dxf.center.y); r = float(e.dxf.radius)
        yield from _sample_arc_pts(cx, cy, r, None, None)
    elif et == "ARC":
        cx, cy = float(e.dxf.center.x), float(e.dxf.center.y); r = float(e.dxf.radius)
        sa, ea = float(e.dxf.start_angle), float(e.dxf.end_angle)
        yield from _sample_arc_pts(cx, cy, r, sa, ea)
    elif et == "HATCH":
        paths = getattr(e, "paths", None)
        if paths:
            for path in paths:
                verts = getattr(path, "polyline_path", None)
                if verts:
                    for v in verts:
                        x = float(getattr(v, "x", v[0])); y = float(getattr(v, "y", v[1]))
                        yield (x, y)

from collections import Counter

def _ins_name(ins) -> str:
    return (getattr(ins, "effective_name", None)
            or getattr(ins, "block_name", None)
            or getattr(ins.dxf, "name", "") or "").strip()

TABLE_SET_NAME_HINTS  = {"TABLE", "DINING", "MEETING", "CONFERENCE", "RECTANGULAR", "CIRCULAR"}
TABLE_SET_LAYER_HINTS = {"MEETING", "CONFERENCE", "DINING"}

def is_table_set_insert(block_name: str, layer_name: str) -> bool:
    n  = (block_name or "").upper().strip()
    ly = (layer_name or "").upper().strip()

    # Never treat chairs as table-sets
    if "CHAIR" in n or n == "CHAIR":
        return False

    # Heuristic: table-ish name OR table-ish layer
    if any(h in n for h in TABLE_SET_NAME_HINTS):
        return True
    if any(h in ly for h in TABLE_SET_LAYER_HINTS):
        return True

    return False

def count_chairs_from_virtual_entities(ins) -> int:
    """
    Works when ARRAY expands into child INSERTs in virtual_entities().
    """
    try:
        n = 0
        for ve in ins.virtual_entities():
            if ve.dxftype() == "INSERT":
                child_name = (_ins_name(ve) or "").upper()
                if "CHAIR" in child_name:
                    n += 1
        return n
    except Exception:
        return 0



def _iter_nested_inserts_in_block(doc, block_name: str, depth: int = 0, max_depth: int = 6, seen=None):
    """Yield INSERT entities inside a block definition (recursive)."""
    if not block_name or depth > max_depth:
        return
    if seen is None:
        seen = set()
    key = (block_name, depth)
    if key in seen:
        return
    seen.add(key)

    if block_name not in doc.blocks:
        return

    blk = doc.blocks.get(block_name)
    for e in blk:
        if e.dxftype() == "INSERT":
            yield e
            child_name = _ins_name(e)
            yield from _iter_nested_inserts_in_block(doc, child_name, depth + 1, max_depth, seen)

def count_components_for_insert(doc, ins):
    """
    Meeting table set logic:
    - table_count = 1 per parent insert
    - chair_count = number of nested INSERTs inside the parent block definition (recursive)
    """
    parent_name = _ins_name(ins)

    c = Counter()
    for child_ins in _iter_nested_inserts_in_block(doc, parent_name):
        child_name = _ins_name(child_ins)
        if child_name:
            c[child_name] += 1

    table_count = 1
    chair_count = sum(c.values())   # <-- key change: don't depend on 'CHAIR' keyword

    return table_count, chair_count, c







# ===== Oriented bbox helpers =====
def _convex_hull(points: list[tuple[float,float]]) -> list[tuple[float,float]]:
    pts = sorted(set(points))
    if len(pts) <= 1:
        return pts
    def cross(o,a,b): return (a[0]-o[0])*(b[1]-o[1]) - (a[1]-o[1])*(b[0]-o[0])
    lower=[]
    for p in pts:
        while len(lower)>=2 and cross(lower[-2],lower[-1],p) <= 0:
            lower.pop()
        lower.append(p)
    upper=[]
    for p in reversed(pts):
        while len(upper)>=2 and cross(upper[-2],upper[-1],p) <= 0:
            upper.pop()
        upper.append(p)
    return lower[:-1] + upper[:-1]

def _oriented_bbox_lengths(points: list[tuple[float,float]]) -> tuple[float,float]:
    if not points:
        return (0.0, 0.0)
    hull = _convex_hull(points)
    if len(hull) == 1:
        return (0.0, 0.0)
    if len(hull) == 2:
        a,b = hull[0], hull[1]
        d = math.hypot(b[0]-a[0], b[1]-a[1])
        return (d, 0.0)

    def proj_extents(pts, cos_t, sin_t):
        xs = []; ys = []
        for x,y in pts:
            xr =  x*cos_t + y*sin_t
            yr = -x*sin_t + y*cos_t
            xs.append(xr); ys.append(yr)
        return (min(xs), max(xs), min(ys), max(ys))

    best_area = float("inf"); best_dims=(0.0,0.0)
    n = len(hull)
    for i in range(n):
        x1,y1 = hull[i]; x2,y2 = hull[(i+1)%n]
        dx,dy = (x2-x1, y2-y1)
        edge_len = math.hypot(dx,dy)
        if edge_len < 1e-12:
            continue
        cos_t = dx/edge_len
        sin_t = dy/edge_len
        minx,maxx,miny,maxy = proj_extents(hull, cos_t, sin_t)
        w = maxx - minx
        h = maxy - miny
        area = w*h
        if area < best_area:
            best_area = area
            L,W = (w,h) if w>=h else (h,w)
            best_dims = (L,W)
    return best_dims

def _bbox_of_insert_xy(ins) -> Optional[Tuple[float,float]]:
    try:
        pts = []
        for ve in ins.virtual_entities():
            for p in _collect_points_from_entity(ve) or []:
                pts.append((float(p[0]), float(p[1])))
        if not pts:
            return None
        L, W = _oriented_bbox_lengths(pts)
        return (L, W)
    except Exception:
        return None


# ===== ZONES =====
def _insert_bbox(ins) -> Optional[tuple[float,float,float,float]]:
    try:
        minx=miny=float("inf"); maxx=maxy=float("-inf"); anyp=False
        for ve in ins.virtual_entities():
            for (x,y) in _collect_points_from_entity(ve) or []:
                anyp=True
                minx=min(minx,x); miny=min(miny,y)
                maxx=max(maxx,x); maxy=max(maxy,y)
        if not anyp: return None
        return (minx,miny,maxx,maxy)
    except Exception:
        return None

def _bbox_center(b: tuple[float,float,float,float]) -> tuple[float,float]:
    minx, miny, maxx, maxy = b
    return ((minx+maxx)*0.5, (miny+maxy)*0.5)

def point_in_polygon(pt: tuple[float,float], poly: list[tuple[float,float]]) -> bool:
    x, y = pt; inside = False
    n = len(poly)
    for i in range(n):
        x1, y1 = poly[i]
        x2, y2 = poly[(i+1) % n]
        if ((y1 > y) != (y2 > y)) and (x < (x2 - x1) * (y - y1) / (y2 - y1 + 1e-12) + x1):
            inside = not inside
    return inside

def _poly_from_lwpoly(e) -> list[tuple[float,float]]:
    verts = list(e)
    if not verts: return []
    pts = []
    n = len(verts)
    closed = bool(getattr(e, "closed", False))
    for i in range(n if closed else n-1):
        j = (i + 1) % n
        x1,y1 = float(verts[i][0]), float(verts[i][1])
        x2,y2 = float(verts[j][0]), float(verts[j][1])
        try: b = float(verts[i][4])
        except Exception: b = 0.0
        seg = _bulge_arc_points((x1,y1),(x2,y2),b)
        if pts and seg:
            pts.extend(seg[1:])
        else:
            pts.extend(seg)
    if closed and pts and pts[0] != pts[-1]:
        pts.append(pts[0])
    return pts

def _normalize_poly(poly: list[tuple[float, float]]) -> list[tuple[float, float]]:
    """Remove duplicate consecutive points + drop explicit closing point for stable PIP."""
    if not poly:
        return poly
    cleaned = []
    for x, y in poly:
        if not cleaned:
            cleaned.append((x, y))
        else:
            px, py = cleaned[-1]
            if abs(x - px) > 1e-9 or abs(y - py) > 1e-9:
                cleaned.append((x, y))
    if len(cleaned) > 2:
        x0, y0 = cleaned[0]
        xN, yN = cleaned[-1]
        if abs(x0 - xN) < 1e-9 and abs(y0 - yN) < 1e-9:
            cleaned.pop()
    return cleaned

def _insert_ref_point(ins) -> Optional[tuple[float, float]]:
    """Best point to classify an INSERT into a PLANNER zone."""
    try:
        ip = getattr(ins.dxf, "insert", None)
        if ip is not None:
            return (float(ip.x), float(ip.y))
    except Exception:
        pass
    return None

@dataclass
class Zone:
    name: str
    poly: list  # list[(x,y)]

def _collect_planner_zones(msp) -> list[Zone]:
    zones: list[Zone] = []

    # ---------- A) PLANNER INSERT zones (bbox -> rectangle poly) ----------
    for ins in msp.query('INSERT[layer=="PLANNER"]'):
        try:
            b = _insert_bbox(ins)
            if not b:
                continue
            minx, miny, maxx, maxy = b
            poly = _normalize_poly([(minx, miny), (maxx, miny), (maxx, maxy), (minx, maxy), (minx, miny)])

            zname = None
            try:
                cand_tags = {"NAME", "ROOM", "ZONE", "LABEL", "TITLE"}
                for att in getattr(ins, "attribs", lambda: [])() or []:
                    tag = (getattr(att.dxf, "tag", "") or "").upper()
                    if tag in cand_tags:
                        txt = (getattr(att.dxf, "text", "") or "").strip()
                        if txt:
                            zname = txt
                            break
            except Exception:
                pass

            if not zname:
                zname = (
                    (getattr(ins, "effective_name", None) or "")
                    or (getattr(ins, "block_name", None) or "")
                    or (getattr(ins.dxf, "name", "") or "")
                ).strip()

            if not zname:
                zname = "Zone"

            zones.append(Zone(name=zname, poly=poly))
        except Exception:
            pass

    # ---------- B) PLANNER closed polylines + labels ----------
    tmp: list[Zone] = []
    for e in msp.query('LWPOLYLINE[layer=="PLANNER"]'):
        try:
            if not bool(getattr(e, "closed", False)):
                continue
            poly = _normalize_poly(_poly_from_lwpoly(e))
            if len(poly) >= 3:
                tmp.append(Zone(name="", poly=poly))
        except Exception:
            pass

    if tmp:
        labels: list[tuple[str, tuple[float, float]]] = []
        for t in msp.query("TEXT"):
            try:
                labels.append(((t.dxf.text or "").strip(), (float(t.dxf.insert.x), float(t.dxf.insert.y))))
            except Exception:
                pass
        for mt in msp.query("MTEXT"):
            try:
                raw = (mt.text or "").strip()
                labels.append((raw.split("\n", 1)[0].strip(), (float(mt.dxf.insert.x), float(mt.dxf.insert.y))))
            except Exception:
                pass

        used = set()
        for i, z in enumerate(tmp, start=1):
            zname = None

            # label inside polygon
            for idx, (txt, pt) in enumerate(labels):
                if idx in used or not txt:
                    continue
                if point_in_polygon(pt, z.poly):
                    zname = txt
                    used.add(idx)
                    break

            # fallback: nearest label to centroid
            if not zname and labels:
                xs = [p[0] for p in z.poly]
                ys = [p[1] for p in z.poly]
                cx = sum(xs) / len(xs)
                cy = sum(ys) / len(ys)

                best = None
                for idx, (txt, (x, y)) in enumerate(labels):
                    if idx in used or not txt:
                        continue
                    d = (x - cx) * (x - cx) + (y - cy) * (y - cy)
                    if best is None or d < best[0]:
                        best = (d, idx, txt)
                if best:
                    _, idx, txt = best
                    zname = txt
                    used.add(idx)

            if not zname:
                zname = f"Zone {i:02d}"

            zones.append(Zone(name=zname, poly=z.poly))

    # de-dupe
    seen = set()
    out = []
    for z in zones:
        key = (z.name, tuple(z.poly))
        if key not in seen:
            out.append(z)
            seen.add(key)

    return out

def _zone_for_point(pt: tuple[float,float], zones: list[Zone]) -> Optional[str]:
    for z in zones:
        if point_in_polygon(pt, z.poly):
            return z.name
    return None

def _entity_center_xy(e) -> Optional[tuple[float,float]]:
    pts = list(_collect_points_from_entity(e) or [])
    if not pts:
        return None
    xs = [p[0] for p in pts]; ys = [p[1] for p in pts]
    return ((min(xs)+max(xs))*0.5, (min(ys)+max(ys))*0.5)

def _zone_for_entity(e, zones: list[Zone]) -> str:
    c = _entity_center_xy(e)
    if not c:
        return ""
    z = _zone_for_point(c, zones)
    return z or ""


# ===== Row builder =====
def make_row(entity_type, qty_type, qty_value,
             block_name="", layer="", handle="", remarks="",
             bbox_length=None, bbox_width=None,
             preview_b64:str="", preview_hex:str="",
             perimeter=None, area=None, zone:str="", category1:str="",
             description:str="") -> dict:
    return {
        "entity_type": entity_type,
        "qty_type": qty_type,
        "qty_value": _fmt_num(qty_value),
        "block_name": block_name or "",
        "layer": layer_or_misc(layer),
        "zone": (zone or ""),
        "category1": category1 or "",
        "handle": handle or "",
        "remarks": remarks or "",
        "bbox_length": _fmt_num(bbox_length),
        "bbox_width":  _fmt_num(bbox_width),
        "preview_b64": preview_b64 or "",
        "preview_hex": preview_hex or "",
        "perimeter": _fmt_num(perimeter),
        "area": _fmt_num(area),
        "description": description or ""
    }


# ===== Previews (ONE.PY) =====
def _preview_build_frontend(doc, ax):
    ctx     = RenderContext(doc)
    backend = MatplotlibBackend(ax)
    try:
        from ezdxf.addons.drawing.config import Configuration
        cfg = Configuration.defaults()
        return Frontend(ctx, backend, cfg)
    except Exception:
        return Frontend(ctx, backend)

def _preview_render_entities_png_bytes(doc, entities, dpi: Optional[int] = None, pad_pct: Optional[float] = None) -> bytes:
    # IMPORTANT: dpi/pad_pct must be read at runtime to respect overrides
    if dpi is None:
        dpi = int(PREVIEW_DPI)
    if pad_pct is None:
        pad_pct = float(PREVIEW_PAD_PCT)

    fig = plt.figure()
    ax  = fig.add_axes([0,0,1,1])
    ax.set_facecolor((0,0,0,0))
    fig.patch.set_alpha(0.0)
    fig.set_dpi(dpi)

    frontend = _preview_build_frontend(doc, ax)
    msp = doc.modelspace()

    for e in entities:
        if hasattr(e, "dxf") and hasattr(e.dxf, "layer"):
            if (e.dxf.layer or "").upper() in PREVIEW_SKIP_LAYERS:
                continue
        if hasattr(e, "dxftype") and e.dxftype() in PREVIEW_SKIP_DXF_TYPES:
            continue
        try:
            frontend.draw_entity(e, msp)
        except Exception:
            pass

    ax.set_aspect("equal", "box")
    ax.autoscale(True, "both", tight=True)

    x0, x1 = ax.get_xlim()
    y0, y1 = ax.get_ylim()
    w = max(x1 - x0, 1e-6)
    h = max(y1 - y0, 1e-6)
    ax.set_xlim(x0 - w*pad_pct, x1 + w*pad_pct)
    ax.set_ylim(y0 - h*pad_pct, y1 + h*pad_pct)

    ax.axis("off")

    buf = io.BytesIO()
    fig.savefig(buf, format="png", transparent=True, bbox_inches="tight", pad_inches=0)
    plt.close(fig)
    return buf.getvalue()

def _ensure_rgba(img):
    if img is None:
        return None
    if img.ndim == 2:
        img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGRA)
    elif img.ndim == 3:
        c = img.shape[2]
        if c == 4:
            pass
        elif c == 3:
            img = np.dstack([img, np.full(img.shape[:2], 255, np.uint8)])
        else:
            img = np.dstack([img[..., 0:3], np.full(img.shape[:2], 255, np.uint8)])
    return np.ascontiguousarray(img, np.uint8)

def _as_single_channel(mask):
    if mask is None:
        return None
    if mask.ndim == 3 and mask.shape[2] == 1:
        mask = mask[..., 0]
    elif mask.ndim == 3 and mask.shape[2] == 3:
        mask = cv2.cvtColor(mask, cv2.COLOR_BGR2GRAY)
    if mask.dtype != np.uint8:
        mask = (mask > 0).astype(np.uint8)
    return np.ascontiguousarray(mask)

def _add_transparent_margin_rgba(img_rgba, margin):
    img_rgba = _ensure_rgba(img_rgba)
    if margin <= 0:
        return img_rgba
    h, w = img_rgba.shape[:2]
    canvas = np.zeros((h + 2*margin, w + 2*margin, 4), np.uint8)
    canvas[margin:margin+h, margin:margin+w] = img_rgba
    return canvas

def _resize_to_square_rgba(img_rgba, size: int, supersample: Optional[int] = None):
    img_rgba = _ensure_rgba(img_rgba)
    if supersample is None:
        supersample = int(PREVIEW_SUPERSAMPLE)

    h, w = img_rgba.shape[:2]
    if h == 0 or w == 0:
        return np.zeros((size, size, 4), np.uint8)

    target_big = size * max(1, supersample)
    scale = min(target_big / w, target_big / h)
    new_w = int(round(w * scale))
    new_h = int(round(h * scale))

    big = cv2.resize(img_rgba, (new_w, new_h), interpolation=cv2.INTER_NEAREST)
    big = _ensure_rgba(big)

    canvas = np.zeros((target_big, target_big, 4), np.uint8)
    x0 = (target_big - new_w) // 2
    y0 = (target_big - new_h) // 2
    canvas[y0:y0+new_h, x0:x0+new_w] = big

    final = cv2.resize(canvas, (size, size), interpolation=cv2.INTER_AREA)
    return final

def _dilate_mask(mask, ksize, iters):
    m = _as_single_channel(mask)
    if m is None or ksize <= 0 or iters <= 0:
        return m
    k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (ksize, ksize))
    out = cv2.dilate(m, k, iterations=iters)
    return _as_single_channel(out)

def _close_small_gaps(mask, ksize=2):
    m = _as_single_channel(mask)
    if m is None or ksize <= 0:
        return m
    k = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (ksize, ksize))
    out = cv2.morphologyEx(m, cv2.MORPH_CLOSE, k, iterations=1)
    return _as_single_channel(out)

def _labels_touching_border(labels):
    h, w = labels.shape
    return set(np.unique(np.r_[labels[0, :], labels[-1, :], labels[:, 0], labels[:, -1]]))

def _onepy_trim_to_icon_png_bytes(
    png_bytes: bytes,
    target_size: Optional[int] = None,
    stroke: Optional[str] = None
) -> bytes:
    if target_size is None:
        target_size = int(PREVIEW_TARGET_SIZE)
    if stroke is None:
        stroke = PREVIEW_STROKE

    arr = np.frombuffer(png_bytes, dtype=np.uint8)
    img = cv2.imdecode(arr, cv2.IMREAD_UNCHANGED)
    img = _ensure_rgba(img)
    if img is None:
        return b""

    rgb = img[..., :3]
    a   = img[..., 3]
    gray = cv2.cvtColor(rgb, cv2.COLOR_BGR2GRAY)

    non_white = (gray < 250).astype(np.uint8) * 255
    trim_mask = cv2.bitwise_or((a > 0).astype(np.uint8) * 255, non_white)

    coords = cv2.findNonZero(trim_mask)
    if coords is not None:
        x, y, w, h = cv2.boundingRect(coords)
        if w > 0 and h > 0:
            img = _ensure_rgba(img[y:y+h, x:x+w])

    size_max = max(img.shape[0], img.shape[1])
    margin = int(round(size_max * float(PREVIEW_MARGIN_PCT)))
    img = _add_transparent_margin_rgba(img, margin)

    rgb = img[..., :3]
    a   = img[..., 3]
    gray = cv2.cvtColor(rgb, cv2.COLOR_BGR2GRAY)

    bg_candidates = _as_single_channel(((gray >= int(PREVIEW_EDGE_BG_THRESH)) | (a <= int(PREVIEW_MIN_VISIBLE_ALPHA))).astype(np.uint8))
    _, labels = cv2.connectedComponents(bg_candidates, 8)
    border_lbls = _labels_touching_border(labels)
    bg_mask = np.isin(labels, list(border_lbls)).astype(np.uint8) * 255

    ink = ((bg_mask == 0) & ((a >= int(PREVIEW_MIN_VISIBLE_ALPHA)) | (gray < int(PREVIEW_INK_THRESHOLD)))).astype(np.uint8) * 255
    ink = _close_small_gaps(ink, int(PREVIEW_CLOSE_GAPS_KSIZE))
    ink = _dilate_mask(ink, int(PREVIEW_THICKEN_PX), int(PREVIEW_THICKEN_ITER))

    h, w = img.shape[:2]
    out = np.zeros((h, w, 4), np.uint8)
    mask = _as_single_channel(ink) > 0

    if stroke == "black":
        out[mask, :3] = 0
    else:
        out[mask, :3] = 255
    out[mask, 3] = 255

    out = _resize_to_square_rgba(out, int(target_size), supersample=int(PREVIEW_SUPERSAMPLE))

    ok, enc = cv2.imencode(".png", out, [int(cv2.IMWRITE_PNG_COMPRESSION), 3])
    return enc.tobytes() if ok else b""

def _render_onepy_preview_base64(
    doc,
    ins,
    target_size: Optional[int] = None,
    stroke: Optional[str] = None,
    dpi: Optional[int] = None,
    pad_pct: Optional[float] = None
) -> str:
    try:
        if target_size is None:
            target_size = int(PREVIEW_TARGET_SIZE)
        if stroke is None:
            stroke = PREVIEW_STROKE

        png0 = _preview_render_entities_png_bytes(doc, [ins], dpi=dpi, pad_pct=pad_pct)
        png1 = _onepy_trim_to_icon_png_bytes(png0, target_size=target_size, stroke=stroke)
        if not png1:
            return ""
        return base64.b64encode(png1).decode("ascii")
    except Exception:
        logging.exception("ONE.PY preview render failed for INSERT=%s", getattr(getattr(ins, "dxf", None), "name", "<?>"))
        return ""


# ===== Description helpers =====
def _description_from_insert(ins) -> str:
    try:
        for att in getattr(ins, "attribs", lambda: [])() or []:
            tag = (getattr(att.dxf, "tag", "") or "").upper()
            if tag in DESC_TAGS:
                txt = (getattr(att.dxf, "text", "") or "").strip()
                if txt:
                    return txt
    except Exception:
        pass
    return ""

def _description_from_blockrecord(msp, base_name: str) -> str:
    try:
        if base_name and hasattr(msp, "doc") and base_name in msp.doc.blocks:
            blkrec = msp.doc.blocks.get(base_name)
            return (getattr(getattr(blkrec, "dxf", None), "description", "") or "").strip()
    except Exception:
        pass
    return ""

def _attdef_default_for_desc(msp, base_name: str) -> str:
    try:
        if base_name and hasattr(msp, "doc") and base_name in msp.doc.blocks:
            blk = msp.doc.blocks.get(base_name)
            for e in blk:
                if e.dxftype() == "ATTDEF":
                    tag = (getattr(e.dxf, "tag", "") or "").upper()
                    if tag in DESC_TAGS:
                        return (getattr(e.dxf, "text", "") or "").strip()
    except Exception:
        pass
    return ""


def _safe_effective_name(ins) -> str:
    return (getattr(ins, "effective_name", None)
            or getattr(ins, "block_name", None)
            or getattr(getattr(ins, "dxf", None), "name", "") or "")

def _resolve_child_layer(child_layer: str, parent_layer: str) -> str:
    cl = (child_layer or "").strip()
    if cl == "" or cl.upper() == "0":
        return parent_layer
    return cl

def _iter_nested_inserts_from_block(doc, parent_ins, include_xrefs: bool,
                                   parent_layer: str, parent_zone: str, parent_cat1: str,
                                   depth: int = 0, max_depth: int = 10):
    if depth >= max_depth:
        return
    try:
        blk_name = getattr(parent_ins.dxf, "name", "") or ""
        if not blk_name or blk_name not in doc.blocks:
            return

        blk = doc.blocks.get(blk_name)

        for e in blk:
            if e.dxftype() != "INSERT":
                continue

            child_name = _safe_effective_name(e)
            if not include_xrefs and ("|" in (child_name or "")):
                continue

            child_layer = _resolve_child_layer(getattr(e.dxf, "layer", ""), parent_layer)

            yield (e, child_layer, parent_zone, parent_cat1)

            yield from _iter_nested_inserts_from_block(
                doc=doc,
                parent_ins=e,
                include_xrefs=include_xrefs,
                parent_layer=child_layer,
                parent_zone=parent_zone,
                parent_cat1=parent_cat1,
                depth=depth + 1,
                max_depth=max_depth
            )
    except Exception:
        return



def _canonical_block_name(ins) -> str:
    eff = (getattr(ins, "effective_name", None) or "").strip()
    raw = (getattr(getattr(ins, "dxf", None), "name", "") or "").strip()

    # If effective_name exists, prefer it (it is usually the "real" name)
    if eff:
        return eff

    # If raw is anonymous, try to resolve via dynamic block record
    if raw.startswith("*U"):
        try:
            # In many cases, anonymous blocks have an "xref" / owner pointing to the base.
            # ezdxf doesn't always expose the full dynamic mapping, so we fallback to raw.
            return raw
        except Exception:
            return raw

    return raw



# ===== Rows: INSERT detail (NO previews here; deferred later) =====
def iter_block_rows(
    msp,
    include_xrefs: bool,
    scale_to_m: float,
    target_units: str,
    zones: list[Zone] | None = None
) -> tuple[list[dict], dict[str, object]]:
    out: list[dict] = []
    reps: dict[str, object] = {}
    zones = zones or []
    doc = getattr(msp, "doc", None)

    def _zone_for_insert(ins) -> str:
        center_zone = ""
        pt = _insert_ref_point(ins)
        if pt:
            zname = _zone_for_point(pt, zones)
            if zname:
                center_zone = zname
        if not center_zone:
            b = _insert_bbox(ins)
            if b:
                cx, cy = _bbox_center(b)
                zname = _zone_for_point((cx, cy), zones)
                if zname:
                    center_zone = zname
        return center_zone or ""

    def _emit_row(ins, resolved_layer: str, zone_name: str, category1: str):
        ly = resolved_layer or ""
        name = _canonical_block_name(ins)


        if layer_or_misc(ly).upper() == "PLANNER":
            return
        if not include_xrefs and ("|" in (name or "")):
            return

        if name and name not in reps:
            reps[name] = ins

        desc_txt = _description_from_insert(ins)
        if not desc_txt:
            desc_txt = _description_from_blockrecord(msp, name)
        if not desc_txt:
            desc_txt = _attdef_default_for_desc(msp, name)

        bbox_du = _bbox_of_insert_xy(ins)
        if bbox_du:
            L_m = bbox_du[0] * scale_to_m
            W_m = bbox_du[1] * scale_to_m
            L_out = to_target_units(L_m, target_units, "length")
            W_out = to_target_units(W_m, target_units, "length")
        else:
            L_out = W_out = None

        upload_layer = "PLANNER" if (FORCE_PLANNER_CATEGORY and zone_name) else ly

        out.append(make_row(
            "INSERT", "count", 1.0,
            block_name=name,
            layer=upload_layer,
            handle=getattr(ins.dxf, "handle", ""),
            bbox_length=L_out, bbox_width=W_out,
            preview_b64="",
            zone=zone_name,
            category1=(category1 or "").strip().lower(),
            remarks="",
            description=desc_txt
        ))

    for ins in msp.query("INSERT"):
        try:
            ly = getattr(ins.dxf, "layer", "")
            if layer_or_misc(ly).upper() == "PLANNER":
                continue

            zone_name = _zone_for_insert(ins)
            cat1 = (ly or "").strip().lower()

            _emit_row(ins, ly, zone_name, cat1)

            if ENABLE_NESTED_INSERTS and doc is not None:
                for (child_ins, child_layer, child_zone, child_cat1) in _iter_nested_inserts_from_block(
                    doc=doc,
                    parent_ins=ins,
                    include_xrefs=include_xrefs,
                    parent_layer=ly,
                    parent_zone=zone_name,
                    parent_cat1=cat1,
                    depth=0,
                    max_depth=int(NESTED_MAX_DEPTH),
                ):
                    _emit_row(child_ins, child_layer, child_zone, child_cat1)

        except Exception as ex:
            logging.exception("INSERT failed: %s", ex)

    return out, reps


# ===== Colors =====
def _layer_rgb_map(doc) -> Dict[str, tuple[int,int,int]]:
    m: Dict[str, tuple[int,int,int]] = {}
    try:
        for layer in doc.layers:
            name = layer.dxf.name or ""
            key = layer_or_misc(name)
            tc = getattr(layer.dxf, "true_color", 0)
            if tc:
                rgb = ezcolors.int2rgb(tc)
            else:
                aci = int(getattr(layer.dxf, "color", 7) or 7)
                rgb = ezcolors.aci2rgb(aci if 0 <= aci <= 256 else 7)
            m[key] = rgb
    except Exception:
        pass
    return m

def _resolve_entity_rgb(e, layer_rgb_map: Dict[str, tuple[int,int,int]]) -> tuple[int,int,int]:
    ly = layer_or_misc(getattr(e.dxf, "layer", ""))
    try:
        tc = int(getattr(e.dxf, "true_color", 0) or 0)
        if tc: return ezcolors.int2rgb(tc)
    except Exception:
        pass
    try:
        aci = int(getattr(e.dxf, "color", 256) or 256)
    except Exception:
        aci = 256
    if 1 <= aci <= 255:
        return ezcolors.aci2rgb(aci)
    return layer_rgb_map.get(ly, (200,200,200))

def _entity_weight_for_colorvote(e) -> float:
    et = e.dxftype()
    try:
        if et == "LINE":
            p1=(e.dxf.start.x,e.dxf.start.y); p2=(e.dxf.end.x,e.dxf.end.y)
            return dist2d(p1,p2)
        if et == "ARC":
            r=float(e.dxf.radius)
            sweep=(float(e.dxf.end_angle)-float(e.dxf.start_angle))%360.0
            return (2.0*math.pi*r)*(sweep/360.0)
        if et == "CIRCLE":
            r=float(e.dxf.radius); return 2.0*math.pi*r
        if et == "LWPOLYLINE":
            verts=list(e)
            if not verts: return 0.0
            dense=[]; n=len(verts)
            closed=bool(getattr(e,"closed",False))
            for i in range(n if closed else n-1):
                j=(i+1)%n
                try: b=float(verts[i][4])
                except Exception: b=0.0
                seg=_bulge_arc_points((float(verts[i][0]),float(verts[i][1])),
                                      (float(verts[j][0]),float(verts[j][1])), b)
                dense.extend(seg[:-1])
            dense.append((float(verts[-1][0]),float(verts[-1][1])))
            return polyline_length_xy(dense, closed=False)
        if et == "POLYLINE":
            vs=list(e.vertices())
            if not vs: return 0.0
            coords=[]
            for v in vs:
                loc=getattr(v.dxf,"location",None)
                coords.append((float(loc.x),float(loc.y)) if loc is not None
                              else (float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
            n=len(coords); closed=bool(getattr(e,"is_closed",getattr(e,"closed",False)))
            dense=[]
            for i in range(n - (0 if closed else 1)):
                j=(i+1)%n
                try: b=float(vs[i].dxf.bulge)
                except Exception: b=0.0
                seg=_bulge_arc_points(coords[i], coords[j], b)
                dense.extend(seg[:-1])
            dense.append(coords[-1])
            return polyline_length_xy(dense, closed=False)
        if et == "HATCH":
            if hasattr(e,"get_filled_area"):
                try: return float(e.get_filled_area()) or 0.0
                except Exception: pass
            return 0.0
    except Exception:
        pass
    return 0.0

def _dominant_layer_rgb_map(msp, base_layer_rgb: Dict[str, tuple[int,int,int]], scale_to_m: float) -> Dict[str, tuple[int,int,int]]:
    votes: Dict[str, Dict[tuple[int,int,int], float]] = {}
    def _acc(e):
        ly = layer_or_misc(getattr(e.dxf,"layer",""))
        rgb = _resolve_entity_rgb(e, base_layer_rgb)
        w   = _entity_weight_for_colorvote(e)
        w *= (scale_to_m**2) if e.dxftype()=="HATCH" else scale_to_m
        if w <= 0: w = 1.0
        d = votes.setdefault(ly, {})
        d[rgb] = d.get(rgb, 0.0) + w

    for et in ("LINE","LWPOLYLINE","POLYLINE","ARC","CIRCLE","HATCH"):
        for e in msp.query(et):
            try: _acc(e)
            except Exception: pass

    out = dict(base_layer_rgb)
    for ly, hist in votes.items():
        if hist:
            out[ly] = max(hist.items(), key=lambda kv: kv[1])[0]
    return out


# ===== Layer metrics =====
def solve_rect_dims_from_perimeter_area(P: float, A: float) -> Tuple[Optional[float], Optional[float]]:
    try:
        if P is None or A is None or P <= 0 or A <= 0:
            return (None, None)
        S = P / 2.0
        D = S*S - 4.0*A
        if D < -1e-9: return (None, None)
        D = max(D, 0.0)
        root = math.sqrt(D)
        a = 0.5*(S + root)
        b = 0.5*(S - root)
        if a <= 0 or b <= 0: return (None, None)
        return (a, b) if a >= b else (b, a)
    except Exception:
        return (None, None)

def compute_layer_metrics(msp, scale_to_m: float, target_units: str, zones: list[Zone]):
    open_len: Dict[tuple[str,str], float] = {}
    peri:     Dict[tuple[str,str], float] = {}
    area:     Dict[tuple[str,str], float] = {}

    def add_open(layer, L_du, z):
        if L_du <= 0: return
        L_out = to_target_units(L_du * scale_to_m, target_units, "length")
        k = (z, layer_or_misc(layer))
        open_len[k] = open_len.get(k, 0.0) + L_out

    def add_peri(layer, P_du, z):
        if P_du <= 0: return
        P_out = to_target_units(P_du * scale_to_m, target_units, "length")
        k = (z, layer_or_misc(layer))
        peri[k] = peri.get(k, 0.0) + P_out

    def add_area(layer, A_du, z):
        if A_du <= 0: return
        A_out = to_target_units(A_du * (scale_to_m**2), target_units, "area")
        k = (z, layer_or_misc(layer))
        area[k] = area.get(k, 0.0) + A_out

    for e in msp.query("LINE"):
        try:
            z = _zone_for_entity(e, zones)
            p1=(e.dxf.start.x,e.dxf.start.y); p2=(e.dxf.end.x,e.dxf.end.y)
            add_open(e.dxf.layer, dist2d(p1,p2), z)
        except Exception:
            pass

    for e in msp.query("LWPOLYLINE"):
        try:
            z = _zone_for_entity(e, zones)
            verts = list(e)
            if not verts: continue
            closed = bool(getattr(e,"closed",False))
            dense=[]; n=len(verts)
            for i in range(n if closed else n-1):
                j=(i+1)%n
                try: b=float(verts[i][4])
                except Exception: b=0.0
                seg=_bulge_arc_points((float(verts[i][0]), float(verts[i][1])),
                                      (float(verts[j][0]), float(verts[j][1])), b)
                dense.extend(seg[:-1])
            dense.append((float(verts[-1][0]), float(verts[-1][1])))
            if closed: dense.append((float(verts[0][0]), float(verts[0][1])))

            L = polyline_length_xy(dense, closed=False)
            if closed:
                add_peri(e.dxf.layer, L, z)
                if len(dense) >= 3:
                    add_area(e.dxf.layer, polygon_area_xy(dense[:-1]), z)
            else:
                add_open(e.dxf.layer, L, z)
        except Exception:
            pass

    for e in msp.query("POLYLINE"):
        try:
            z = _zone_for_entity(e, zones)
            vs = list(e.vertices())
            if not vs: continue
            coords=[]
            for v in vs:
                loc=getattr(v.dxf,"location",None)
                coords.append((float(loc.x),float(loc.y)) if loc is not None
                              else (float(getattr(v.dxf,"x",0.0)), float(getattr(v.dxf,"y",0.0))))
            closed = bool(getattr(e,"is_closed",getattr(e,"closed",False)))
            n=len(coords); dense=[]
            for i in range(n - (0 if closed else 1)):
                j=(i+1)%n
                try: b=float(vs[i].dxf.bulge)
                except Exception: b=0.0
                seg=_bulge_arc_points(coords[i], coords[j], b)
                dense.extend(seg[:-1])
            dense.append(coords[-1])
            if closed: dense.append(coords[0])

            L = polyline_length_xy(dense, closed=False)
            if closed:
                add_peri(e.dxf.layer, L, z)
                if len(dense) >= 3:
                    add_area(e.dxf.layer, polygon_area_xy(dense[:-1]), z)
            else:
                add_open(e.dxf.layer, L, z)
        except Exception:
            pass

    for e in msp.query("ARC"):
        try:
            z = _zone_for_entity(e, zones)
            r=float(e.dxf.radius)
            sweep=(float(e.dxf.end_angle)-float(e.dxf.start_angle))%360.0
            add_open(e.dxf.layer, (2.0*math.pi*r)*(sweep/360.0), z)
        except Exception:
            pass

    for e in msp.query("CIRCLE"):
        try:
            z = _zone_for_entity(e, zones)
            r=float(e.dxf.radius)
            add_peri(e.dxf.layer, (2.0*math.pi*r), z)
            add_area(e.dxf.layer, math.pi*(r**2), z)
        except Exception:
            pass

    for e in msp.query("HATCH"):
        try:
            z = _zone_for_entity(e, zones)
            A_du=None
            if hasattr(e,"get_filled_area"):
                try:
                    v=e.get_filled_area()
                    if v and v>0: A_du=float(v)
                except Exception:
                    A_du=None
            if A_du and A_du>0:
                add_area(e.dxf.layer, A_du, z)
        except Exception:
            pass

    return open_len, peri, area

def make_layer_total_rows(open_by, peri_by, area_by, layer_rgb=None):
    rows = []
    layer_rgb = layer_rgb or {}
    all_keys = sorted(set(open_by.keys()) | set(peri_by.keys()) | set(area_by.keys()))

    def hex_for(ly: str) -> str:
        rgb = layer_rgb.get(ly)
        return _rgb_to_hex(rgb) if rgb else ""

    for (z, ly) in all_keys:
        if open_by.get((z, ly), 0.0) > 0:
            rows.append(make_row(
                "LAYER_SUMMARY","layer",None,
                layer=ly, zone=z,
                remarks="OPEN length only",
                bbox_length=open_by[(z, ly)],
                preview_hex=hex_for(ly),
            ))

        P = peri_by.get((z, ly), 0.0)
        A = area_by.get((z, ly), 0.0)
        if P > 0 or A > 0:
            L_rec, W_rec = solve_rect_dims_from_perimeter_area(P, A)
            rows.append(make_row(
                "LAYER_SUMMARY","layer",None,
                layer=ly, zone=z,
                remarks="CLOSED (rectangle): length/width + perimeter & area",
                bbox_length=L_rec, bbox_width=W_rec,
                perimeter=P, area=A,
                preview_hex=hex_for(ly),
            ))
    return rows


# ===== Sorting (keeps category blocks tidy) =====
def _norm_cat(s: str) -> str:
    s = (s or "").strip()
    s = " ".join(s.split())
    return s.upper()

def sort_rows_for_category_blocks(rows: list[dict]) -> None:
    def _key(r):
        cat  = _norm_cat(r.get("layer", ""))
        zone = (r.get("zone", "") or "").lower()
        et   = r.get("entity_type", "") or ""
        et_rank = 1 if et == "LAYER_SUMMARY" else 0
        cat1 = (r.get("category1", "") or "").lower()
        boq  = (r.get("block_name", "") or "").lower()
        return (cat, et_rank, zone, cat1, boq)
    rows.sort(key=_key)


# ===== CSV =====
def write_csv(rows: list[dict], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=CSV_HEADERS)
        writer.writeheader()
        for r in rows:
            writer.writerow({
                "entity_type": r.get("entity_type",""),
                "category":    r.get("layer",""),
                "zone":        r.get("zone",""),
                "category1":   r.get("category1",""),
                "BOQ name":    r.get("block_name",""),
                "qty_type":    r.get("qty_type",""),
                "qty_value":   r.get("qty_value",""),
                "length (ft)": r.get("bbox_length",""),
                "width (ft)":  r.get("bbox_width",""),
                "perimeter":   r.get("perimeter",""),
                "area (ft2)":  r.get("area",""),
                "Description": r.get("description",""),
                "Preview":     "",
                "remarks":     r.get("remarks",""),
            })


# ===== Sheets upload =====
def split_rows_for_upload(rows: list[dict]) -> tuple[list[dict], list[dict]]:
    detail, layer = [], []
    for r in rows:
        if (r.get("entity_type") == "LAYER_SUMMARY") and (r.get("qty_type") == "layer"):
            layer.append(r)
        else:
            detail.append(r)
    return detail, layer

def push_rows_to_webapp(rows: list[dict], webapp_url: str, spreadsheet_id: str,
                        tab: str, mode: str = "replace", summary_tab: str = "",
                        batch_rows: int = 300, timeout: int = 300,
                        valign_middle: bool = False, sparse_anchor: str = "last",
                        drive_folder_id: str = "") -> None:
    if not webapp_url or not spreadsheet_id or not tab:
        logging.info("WebApp push not configured (missing url/id/tab). Skipping upload.")
        return

    sess = requests.Session()
    first_mode = (mode or "replace").lower()

    def post_with_retries(payload, tries=4, backoff=2.0):
        for attempt in range(1, tries+1):
            try:
                return sess.post(webapp_url, json=payload, timeout=timeout, allow_redirects=True)
            except requests.exceptions.ReadTimeout:
                if attempt == tries:
                    raise
                time.sleep(backoff ** attempt)

    seen_preview_boq: set[str] = set()

    total = len(rows); sent = 0
    for idx, i in enumerate(range(0, total, batch_rows), start=1):
        chunk = rows[i:i+batch_rows]
        if not chunk:
            break

        is_layer = (chunk[0].get("entity_type") == "LAYER_SUMMARY")

        if is_layer:
            headers = LAYER_HEADERS
            data_rows = [[
                r.get("layer",""),
                r.get("zone",""),
                r.get("bbox_length",""),
                r.get("bbox_width",""),
                r.get("perimeter",""),
                r.get("area",""),
                "",
            ] for r in chunk]
            images     = [""] * len(chunk)
            bg_colors  = [r.get("preview_hex","") for r in chunk]
            color_only = True

        else:
            headers = DETAIL_HEADERS
            data_rows = [[
                r.get("entity_type",""),
                r.get("layer",""),
                r.get("zone",""),
                r.get("category1",""),
                r.get("block_name",""),
                r.get("qty_type",""),
                r.get("qty_value",""),
                r.get("bbox_length",""),
                r.get("bbox_width",""),
                r.get("description",""),
                "",
                r.get("remarks",""),
            ] for r in chunk]

            if ENABLE_PREVIEWS:
                images = []
                for r in chunk:
                    boq = (r.get("block_name", "") or "").strip()
                    if boq and boq not in seen_preview_boq:
                        images.append(r.get("preview_b64", ""))
                        seen_preview_boq.add(boq)
                    else:
                        images.append("")
            else:
                images = [""] * len(chunk)

            bg_colors  = [""] * len(chunk)
            color_only = False

        payload = {
            "sheetId": spreadsheet_id,
            "tab": tab,
            "mode": "replace" if (i == 0 and first_mode == "replace") else "append",
            "headers": headers if (i == 0 and first_mode == "replace") else [],
            "rows": data_rows,
            "images": images,
            "bgColors": bg_colors,
            "colorOnly": color_only,
            "embedImages": False,
            "driveFolderId": (drive_folder_id or ""),
            "vAlign": "middle" if valign_middle else "",
            "sparseAnchor": (sparse_anchor or "last"),
        }

        if summary_tab and i == 0:
            payload["summaryTab"] = summary_tab
            payload["summaryRows"] = []

        resp = post_with_retries(payload)
        if not resp.ok:
            raise RuntimeError(f"WebApp upload failed (batch {idx}): HTTP {resp.status_code} {resp.text}")

        sent += len(data_rows)
        logging.info("WebApp batch %d: uploaded %d/%d rows", idx, sent, total)


# ===== Main pipeline =====
def collect_dxf_files(path: Path, recursive: bool) -> List[Path]:
    if path.is_file():
        if path.suffix.lower() == ".dxf": return [path]
        logging.error("Provided file is not a .dxf: %s", path); return []
    if path is None or not path.exists():
        logging.error("Path does not exist: %s", path); return []
    pattern = "**/*.dxf" if recursive else "*.dxf"
    files = sorted(path.glob(pattern))
    if not files: logging.warning("No DXF files found in %s (recursive=%s)", path, recursive)
    return files

def derive_out_path(dxf_path: Path, out_dir: Path | None) -> Path:
    return (out_dir / f"{dxf_path.stem}_raw_extract.csv") if out_dir else dxf_path.with_name(f"{dxf_path.stem}_raw_extract.csv")

def process_one_dxf(
    dxf_path: Path,
    out_dir: Path | None,
    target_units: str,
    include_xrefs: bool,
    layer_metrics: bool,
    aggregate_inserts: bool,
    unitless_units: str
) -> list[dict]:
    logging.info("Processing DXF: %s", dxf_path)

    try:
        doc = ezdxf.readfile(str(dxf_path))
    except ezdxf.DXFStructureError:
        logging.warning("DXFStructureError — attempting recover.readfile()")
        doc, _auditor = recover.readfile(str(dxf_path))

    msp = doc.modelspace()
    logging.info("DWG $INSUNITS: %s", doc.header.get("$INSUNITS", "n/a"))
    scale_to_m = units_scale_to_meters(doc, unitless_units=unitless_units)

    # 1) zones
    zones = _collect_planner_zones(msp)

    rows: list[dict] = []

    # 2) INSERT rows (no previews yet) + representative INSERT per block_name
    insert_rows, rep_inserts = iter_block_rows(
        msp=msp,
        include_xrefs=include_xrefs,
        scale_to_m=scale_to_m,
        target_units=target_units,
        zones=zones,
    )

    # 3) Aggregate INSERTs (IMPORTANT: fixed indentation + actually emit aggregated rows)
    if aggregate_inserts:
        # key: (BOQ name, category(upload layer), category1(original dwg layer), zone)
        groups: Dict[tuple[str, str, str, str], dict] = {}

        for r in insert_rows:
            name = (r.get("block_name") or "").strip()
            layer = (r.get("layer") or "").strip()
            cat1 = (r.get("category1") or "").strip().lower()
            zone = (r.get("zone") or "").strip()

            key = (name, layer, cat1, zone)
            g = groups.get(key)
            if g is None:
                g = groups[key] = {
                    "count": 0.0,
                    "xs": [],
                    "ys": [],
                    "desc": (r.get("description") or ""),
                    "remarks": (r.get("remarks") or ""),
                }

            # sum qty_value (NOT row count)
            try:
                g["count"] += float(r.get("qty_value") or 0)
            except Exception:
                pass

            # median bbox (if present)
            try:
                if r.get("bbox_length") and r.get("bbox_width"):
                    g["xs"].append(float(r["bbox_length"]))
                    g["ys"].append(float(r["bbox_width"]))
            except Exception:
                pass

            # keep first non-empty description
            if (not g.get("desc")) and r.get("description"):
                g["desc"] = r.get("description") or ""

        # 3A) Render previews once per BOQ name (optional)
        previews_by_boq: Dict[str, str] = {}
        if ENABLE_PREVIEWS and rep_inserts:
            for boq_name, ins in rep_inserts.items():
                try:
                    # Respect runtime overrides via globals
                    previews_by_boq[boq_name] = _render_onepy_preview_base64(
                        doc=doc,
                        ins=ins,
                        target_size=int(PREVIEW_TARGET_SIZE),
                        stroke=PREVIEW_STROKE,
                        dpi=int(PREVIEW_DPI),
                        pad_pct=float(PREVIEW_PAD_PCT),
                    )
                except Exception:
                    previews_by_boq[boq_name] = ""

        # 3B) Emit aggregated rows
        for (name, layer, cat1, zone), g in groups.items():
            qty = g.get("count", 0.0)

            L = ""
            W = ""
            try:
                if g.get("xs"):
                    L = _fmt_num(float(np.median(g["xs"])))
                if g.get("ys"):
                    W = _fmt_num(float(np.median(g["ys"])))
            except Exception:
                L, W = "", ""

            rows.append(make_row(
                entity_type="INSERT",
                qty_type="count",
                qty_value=qty,
                block_name=name,
                layer=layer,
                handle="",
                remarks=g.get("remarks", ""),
                bbox_length=L if L != "" else None,
                bbox_width=W if W != "" else None,
                preview_b64=previews_by_boq.get(name, ""),
                zone=zone,
                category1=cat1,
                description=g.get("desc", ""),
            ))

    else:
        # Non-aggregated: still attach preview per row BOQ name (sparse upload will only embed first occurrence)
        if ENABLE_PREVIEWS and rep_inserts:
            previews_by_boq: Dict[str, str] = {}
            for boq_name, ins in rep_inserts.items():
                try:
                    previews_by_boq[boq_name] = _render_onepy_preview_base64(
                        doc=doc,
                        ins=ins,
                        target_size=int(PREVIEW_TARGET_SIZE),
                        stroke=PREVIEW_STROKE,
                        dpi=int(PREVIEW_DPI),
                        pad_pct=float(PREVIEW_PAD_PCT),
                    )
                except Exception:
                    previews_by_boq[boq_name] = ""

            for r in insert_rows:
                bn = (r.get("block_name") or "").strip()
                r["preview_b64"] = previews_by_boq.get(bn, "")

        rows.extend(insert_rows)

    # 4) Layer metrics
    if layer_metrics:
        open_by, peri_by, area_by = compute_layer_metrics(msp, scale_to_m, target_units, zones)
        base_layer_rgb = _layer_rgb_map(doc)
        dom_layer_rgb = _dominant_layer_rgb_map(msp, base_layer_rgb, scale_to_m)
        rows.extend(make_layer_total_rows(open_by, peri_by, area_by, layer_rgb=dom_layer_rgb))

    # 5) Sort + write CSV
    sort_rows_for_category_blocks(rows)

    out_path = derive_out_path(dxf_path, out_dir)
    write_csv(rows, out_path)
    logging.info("CSV written to: %s", out_path)

    return rows


# ===== CloudConvert (DWG→DXF) =====
CLOUDCONVERT_API_KEY = os.getenv("CLOUDCONVERT_API_KEY", "").strip()
CLOUDCONVERT_API_BASE = os.getenv("CLOUDCONVERT_API_BASE", "https://api.cloudconvert.com/v2").strip()

def cloudconvert_dwg_to_dxf_bytes(dwg_path: Path) -> bytes:
    if not CLOUDCONVERT_API_KEY:
        raise RuntimeError("Missing CLOUDCONVERT_API_KEY in environment")

    headers = {"Authorization": f"Bearer {CLOUDCONVERT_API_KEY}"}
    sess = requests.Session()

    job_payload = {
        "tasks": {
            "import-1": {"operation": "import/upload"},
            "convert-1": {
                "operation": "convert",
                "input": "import-1",
                "input_format": "dwg",
                "output_format": "dxf",
            },
            "export-1": {"operation": "export/url", "input": "convert-1"},
        }
    }

    r = sess.post(f"{CLOUDCONVERT_API_BASE}/jobs", json=job_payload, headers=headers, timeout=60)
    r.raise_for_status()
    job = r.json()["data"]
    job_id = job["id"]

    import_task = next(t for t in job["tasks"] if t["name"] == "import-1")
    form = import_task["result"]["form"]
    upload_url = form["url"]
    params = form["parameters"]

    with dwg_path.open("rb") as f:
        up = sess.post(upload_url, data=params, files={"file": (dwg_path.name, f)}, timeout=300)
    up.raise_for_status()

    while True:
        j = sess.get(f"{CLOUDCONVERT_API_BASE}/jobs/{job_id}", headers=headers, timeout=60)
        j.raise_for_status()
        data = j.json()["data"]
        status = data.get("status")

        if status == "finished":
            export_task = next(t for t in data["tasks"] if t["name"] == "export-1")
            file_url = export_task["result"]["files"][0]["url"]
            out = sess.get(file_url, timeout=300)
            out.raise_for_status()
            return out.content

        if status == "error":
            raise RuntimeError("CloudConvert job failed: " + json.dumps(data, indent=2)[:5000])

        time.sleep(1.5)


# =========================
# FastAPI/Render entrypoints
# =========================
def process_doc_from_stream(stream) -> dict:
    try:
        dxf_text = stream.read()
    except Exception as e:
        raise ValueError(f"Could not read DXF stream: {e}")

    if not dxf_text or not str(dxf_text).strip():
        raise ValueError("Empty DXF stream")

    upload_id = str(uuid.uuid4())[:12]

    with tempfile.TemporaryDirectory() as td:
        td_path = Path(td)
        dxf_path = td_path / f"upload_{upload_id}.dxf"
        dxf_path.write_text(dxf_text, encoding="utf-8", errors="replace")

        rows = process_one_dxf(
            dxf_path=dxf_path,
            out_dir=td_path,
            target_units="ft",
            include_xrefs=False,
            layer_metrics=True,
            aggregate_inserts=True,
            unitless_units="m",
        )

        detail_rows, layer_rows = split_rows_for_upload(rows)

        if detail_rows and GS_WEBAPP_URL and GSHEET_ID and GSHEET_TAB:
            push_rows_to_webapp(
                detail_rows,
                webapp_url=GS_WEBAPP_URL,
                spreadsheet_id=GSHEET_ID,
                tab=GSHEET_TAB,
                mode=GSHEET_MODE,
                summary_tab="",
                batch_rows=300,
                valign_middle=True,
                sparse_anchor="last",
                drive_folder_id=GS_DRIVE_FOLDER_ID,
            )

        summary_tab_name = GSHEET_SUMMARY_TAB.strip() if GSHEET_SUMMARY_TAB.strip() else (GSHEET_TAB + "_ByLayer")
        if layer_rows and GS_WEBAPP_URL and GSHEET_ID and summary_tab_name:
            push_rows_to_webapp(
                layer_rows,
                webapp_url=GS_WEBAPP_URL,
                spreadsheet_id=GSHEET_ID,
                tab=summary_tab_name,
                mode="replace" if GSHEET_MODE == "replace" else "append",
                summary_tab="",
                batch_rows=300,
                valign_middle=True,
                sparse_anchor="last",
                drive_folder_id=GS_DRIVE_FOLDER_ID,
            )

        return {
            "ok": True,
            "upload_id": upload_id,
            "gsheet_id": GSHEET_ID,
            "sheet_tab": GSHEET_TAB,
            "sheet_summary_tab": summary_tab_name,
            "total_rows": len(rows),
            "detail_rows": len(detail_rows),
            "layer_rows": len(layer_rows),
        }

def process_cad_from_upload(filename: str, file_bytes: bytes, settings_json: str = "") -> dict:
    if not filename:
        raise ValueError("Missing filename")

    ext = Path(filename).suffix.lower().strip()
    if ext not in (".dwg", ".dxf"):
        raise ValueError(f"Unsupported file type: {ext}. Upload .dwg or .dxf")

    if not file_bytes or len(file_bytes) < 10:
        raise ValueError("Empty upload")

    cfg = _parse_settings_json(settings_json)

    with _apply_runtime_overrides(cfg):
        upload_id = str(uuid.uuid4())[:12]

        with tempfile.TemporaryDirectory() as td:
            td_path = Path(td)

            in_path = td_path / f"upload_{upload_id}{ext}"
            in_path.write_bytes(file_bytes)

            if ext == ".dwg":
                dxf_bytes = cloudconvert_dwg_to_dxf_bytes(in_path)
                dxf_path = td_path / f"converted_{upload_id}.dxf"
                dxf_path.write_bytes(dxf_bytes)
            else:
                dxf_path = in_path

            rows = process_one_dxf(
                dxf_path=dxf_path,
                out_dir=td_path,
                target_units="ft",
                include_xrefs=False,
                layer_metrics=True,
                aggregate_inserts=True,
                unitless_units="m",
            )

            detail_rows, layer_rows = split_rows_for_upload(rows)

            if detail_rows and GS_WEBAPP_URL and GSHEET_ID and GSHEET_TAB:
                push_rows_to_webapp(
                    detail_rows,
                    webapp_url=GS_WEBAPP_URL,
                    spreadsheet_id=GSHEET_ID,
                    tab=GSHEET_TAB,
                    mode=GSHEET_MODE,
                    summary_tab="",
                    batch_rows=300,
                    valign_middle=True,
                    sparse_anchor="last",
                    drive_folder_id=GS_DRIVE_FOLDER_ID,
                )

            summary_tab_name = GSHEET_SUMMARY_TAB.strip() if GSHEET_SUMMARY_TAB.strip() else (GSHEET_TAB + "_ByLayer")
            if layer_rows and GS_WEBAPP_URL and GSHEET_ID and summary_tab_name:
                push_rows_to_webapp(
                    layer_rows,
                    webapp_url=GS_WEBAPP_URL,
                    spreadsheet_id=GSHEET_ID,
                    tab=summary_tab_name,
                    mode="replace" if GSHEET_MODE == "replace" else "append",
                    summary_tab="",
                    batch_rows=300,
                    valign_middle=True,
                    sparse_anchor="last",
                    drive_folder_id=GS_DRIVE_FOLDER_ID,
                )

            sheet_url = f"https://docs.google.com/spreadsheets/d/{GSHEET_ID}"

            return {
                "ok": True,
                "message": "✅ BOQ generated successfully.",
                "sheetUrl": sheet_url,
                "sheetName": GSHEET_TAB,
                "uploadId": upload_id,
                "total_rows": len(rows),
                "detail_rows": len(detail_rows),
                "layer_rows": len(layer_rows),
            }


def main():
    ap = argparse.ArgumentParser(description="DXF → CSV + Sheets upload (previews + zones + category1 + description).")
    ap.add_argument("--dxf"); ap.add_argument("--name")
    ap.add_argument("--decimals", type=int, default=None)
    ap.add_argument("--out-dir"); ap.add_argument("--out")
    ap.add_argument("--recursive", action="store_true")
    ap.add_argument("--target-units", default="ft")
    ap.add_argument("--include-xrefs", action="store_true")
    ap.add_argument("--no-layer-metrics", action="store_true")
    ap.add_argument("--no-aggregate-inserts", action="store_true")
    ap.add_argument("--gs-webapp", default=None); ap.add_argument("--gsheet-id", default=None)
    ap.add_argument("--gsheet-tab", default=None); ap.add_argument("--gsheet-summary-tab", default=None)
    ap.add_argument("--gsheet-mode", choices=["replace","append"], default=None)
    ap.add_argument("--batch-rows", type=int, default=300)
    ap.add_argument("--align-middle", action="store_true")
    ap.add_argument("--sparse-anchor", choices=["first","last","middle"], default="last")
    ap.add_argument("--drive-folder-id", default=None)
    ap.add_argument("--unitless-units", choices=["m","cm","mm","in","ft"], default="m")
    ap.add_argument("--verbose", action="store_true")
    args = ap.parse_args()

    global DEC_PLACES
    if args.decimals is not None:
        DEC_PLACES = max(0, min(10, args.decimals))

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)s: %(message)s")

    dxf_input = Path(args.dxf) if args.dxf else (Path(DXF_FOLDER)/f"{args.name}.dxf" if args.name else Path(DXF_FOLDER))
    out_dir   = Path(args.out_dir) if args.out_dir else Path(OUT_ROOT)
    explicit_out = Path(args.out) if args.out else None

    layer_metrics = not args.no_layer_metrics
    aggregate_inserts = not args.no_aggregate_inserts

    gs_webapp = (args.gs_webapp if args.gs_webapp is not None else GS_WEBAPP_URL).strip()
    gsheet_id = (args.gsheet_id if args.gsheet_id is not None else GSHEET_ID).strip()
    gsheet_tab = (args.gsheet_tab if args.gsheet_tab is not None else GSHEET_TAB).strip()
    gsheet_summary_tab = (args.gsheet_summary_tab if args.gsheet_summary_tab is not None else GSHEET_SUMMARY_TAB).strip()
    gsheet_mode = (args.gsheet_mode if args.gsheet_mode is not None else GSHEET_MODE).strip().lower()
    batch_rows = int(args.batch_rows)
    align_middle = args.align_middle
    sparse_anchor = args.sparse_anchor
    drive_folder_id = (args.drive_folder_id if args.drive_folder_id is not None else GS_DRIVE_FOLDER_ID).strip()

    if not dxf_input.exists():
        logging.error("DXF input path not found: %s", dxf_input); return

    files = collect_dxf_files(dxf_input, recursive=args.recursive)
    if not files: return

    def _summary_tab_name():
        return gsheet_summary_tab if gsheet_summary_tab else (gsheet_tab + "_ByLayer")

    if explicit_out:
        if len(files) != 1:
            logging.error("--out is for a single file. For folders, use --out-dir."); return
        f = files[0]
        rows = process_one_dxf(f, explicit_out.parent, args.target_units, args.include_xrefs,
                               layer_metrics, aggregate_inserts,
                               unitless_units=args.unitless_units)
        write_csv(rows, explicit_out)

        if gs_webapp and gsheet_id:
            detail_rows, layer_rows = split_rows_for_upload(rows)
            if detail_rows:
                push_rows_to_webapp(detail_rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, "",
                                    batch_rows=batch_rows, valign_middle=align_middle,
                                    sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)
            if layer_rows:
                push_rows_to_webapp(layer_rows, gs_webapp, gsheet_id, _summary_tab_name(),
                                    "replace" if gsheet_mode=="replace" else "append", "",
                                    batch_rows=batch_rows, valign_middle=align_middle,
                                    sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)
        return

    out_dir = out_dir if str(out_dir).strip() else None
    all_rows: list[dict] = []
    for f in files:
        try:
            rows = process_one_dxf(f, out_dir, args.target_units, args.include_xrefs,
                                   layer_metrics, aggregate_inserts,
                                   unitless_units=args.unitless_units)
            all_rows.extend(rows or [])
        except Exception as ex:
            logging.exception("Failed processing %s: %s", f, ex)

    if all_rows and gs_webapp and gsheet_id:
        detail_rows, layer_rows = split_rows_for_upload(all_rows)
        if detail_rows:
            push_rows_to_webapp(detail_rows, gs_webapp, gsheet_id, gsheet_tab, gsheet_mode, "",
                                batch_rows=batch_rows, valign_middle=align_middle,
                                sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)
        if layer_rows:
            push_rows_to_webapp(layer_rows, gs_webapp, gsheet_id, _summary_tab_name(),
                                "replace" if gsheet_mode=="replace" else "append", "",
                                batch_rows=batch_rows, valign_middle=align_middle,
                                sparse_anchor=sparse_anchor, drive_folder_id=drive_folder_id)

if __name__ == "__main__":
    main()
