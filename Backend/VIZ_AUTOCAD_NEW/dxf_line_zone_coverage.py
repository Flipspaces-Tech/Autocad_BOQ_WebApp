#!/usr/bin/env python3
# ==========================================================
# DXF → Line-vs-Zone Coverage (auto-detect, no layer assumptions)
# - Zones: closed LWPOLYLINE, closed POLYLINE, HATCH outer loops
# - Walls: LINE + any OPEN LWPOLYLINE / OPEN POLYLINE (split into segments)
# - Traverses INSERTs (virtual_entities) with transforms (depth-limited)
# - Outputs "long format": one row per (line_segment × zone) with inside length & %
# - Writes one CSV per DXF into OUT_ROOT
#
# Requires: pip install ezdxf shapely requests   (requests only if PUSH_TO_SHEET=True)
# Python: 3.8+
# ==========================================================

from __future__ import annotations
import csv, math
from pathlib import Path
from typing import List, Dict, Tuple, Iterable, Optional

import ezdxf
from ezdxf.entities import DXFEntity, LWPolyline, Polyline, Hatch, Line, Insert
from ezdxf.math import Matrix44
from shapely.geometry import LineString, Polygon, LinearRing, MultiLineString
from shapely.ops import unary_union

# --------- CONFIG (paths; Sheets push optional) ----------
DXF_FOLDER = r"C:\Users\admin\Documents\VIZ_AUTOCAD_NEW\DXF"
OUT_ROOT   = r"C:\Users\admin\Documents\VIZ_AUTOCAD_NEW\EXPORTS"
CSV_SUFFIX = "_line_zone_coverage.csv"

PUSH_TO_SHEET = False  # set True if you want to POST to Google Sheet
GS_WEBAPP_URL = "https://script.google.com/macros/s/AKfycbzdajkMohJJnlwbCKLQIp6imQe8VVCkkLQD4fB1sa0_2MfN7yhPONo8j3IacxIWna8u/exec"
GSHEET_ID     = "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM"
GSHEET_TAB    = "CRED_UPDATED"
GSHEET_MODE   = "replace"

# --------- TUNING ----------
TOUCH_TOL = 1e-6       # treat boundary-touch as inside
MIN_ZONE_AREA = 1e-8   # drop microscopic polygons
MAX_BLOCK_DEPTH = 4    # how deep to traverse INSERT nesting
MAX_ARC_SEGMENTS = 64  # upper bound when approximating bulge arcs

# =========================================================
#                    GEOMETRY HELPERS
# =========================================================
def _as_polygon_from_lwpoly(pl: LWPolyline, m: Matrix44) -> Optional[Polygon]:
    if not pl.closed:
        return None
    pts = [m.transform((p[0], p[1], 0.0)) for p in pl.get_points("xy")]
    ring = [(x, y) for x, y, _ in pts]
    poly = Polygon(ring)
    if not poly.is_valid:
        poly = poly.buffer(0)
    return poly if (poly and poly.area > MIN_ZONE_AREA) else None

def _as_polygon_from_polyline(pl: Polyline, m: Matrix44) -> Optional[Polygon]:
    if not pl.is_closed:
        return None
    pts = [m.transform((v.dxf.location.x, v.dxf.location.y, 0.0)) for v in pl.vertices]
    ring = [(x, y) for x, y, _ in pts]
    poly = Polygon(ring)
    if not poly.is_valid:
        poly = poly.buffer(0)
    return poly if (poly and poly.area > MIN_ZONE_AREA) else None

def _as_polygons_from_hatch(h: Hatch, m: Matrix44) -> List[Polygon]:
    polys: List[Polygon] = []
    try:
        for path in h.paths:
            if not path.is_outer:
                continue
            pts: List[Tuple[float, float]] = []
            for e in path.edges:
                if e.TYPE == "LineEdge":
                    p0 = m.transform((e.start.x, e.start.y, 0.0))
                    p1 = m.transform((e.end.x, e.end.y, 0.0))
                    if not pts:
                        pts.append((p0[0], p0[1]))
                    pts.append((p1[0], p1[1]))
                elif e.TYPE in ("ArcEdge", "EllipseEdge", "SplineEdge"):
                    # Coarse discretization
                    for t in [i/32 for i in range(33)]:
                        p = e.point(t)
                        P = m.transform((p.x, p.y, 0.0))
                        pts.append((P[0], P[1]))
            if len(pts) >= 3:
                poly = Polygon(pts)
                if not poly.is_valid:
                    poly = poly.buffer(0)
                if poly.area > MIN_ZONE_AREA:
                    polys.append(poly)
    except Exception:
        pass
    return polys

def _as_linestring_from_line(e: Line, m: Matrix44) -> Optional[LineString]:
    s = m.transform(e.dxf.start)
    t = m.transform(e.dxf.end)
    ls = LineString([(s[0], s[1]), (t[0], t[1])])
    return ls if ls.length > 0 else None

# ------ Open LWPOLYLINE segmentation (with bulge support) ------
def _lwpoly_points_xyb(pl: LWPolyline):
    # returns [(x,y,bulge), ...] (ezdxf provides "xyb" order)
    raw = pl.get_points("xyb")
    out = []
    for tup in raw:
        # ezdxf returns (x, y, start_width, end_width, bulge)
        x, y = tup[0], tup[1]
        bulge = tup[4] if len(tup) > 4 else 0.0
        out.append((x, y, float(bulge)))
    return out

def _bulge_arc_points(p0, p1, bulge):
    # Approximate an arc (bulge) with piecewise segments.
    # Bulge = tan(theta/4), theta is signed central angle.
    if bulge == 0.0:
        return [p0, p1]
    theta = 4.0 * math.atan(bulge)
    steps = max(2, min(MAX_ARC_SEGMENTS, int(math.ceil(abs(theta) * 12.0 / math.pi))))
    x0, y0 = p0; x1, y1 = p1
    # chord
    cx, cy = (x0 + x1) / 2.0, (y0 + y1) / 2.0
    vx, vy = (x1 - x0), (y1 - y0)
    L = math.hypot(vx, vy)
    if L == 0:
        return [p0, p1]
    r = (L / 2.0) / abs(math.sin(theta / 2.0))
    # unit perpendicular (left of chord)
    perp_x, perp_y = (-vy / L, vx / L)
    h = r * math.cos(theta / 2.0)
    sx, sy = cx + math.copysign(h, bulge) * perp_x, cy + math.copysign(h, bulge) * perp_y
    a0 = math.atan2(y0 - sy, x0 - sx)
    pts = []
    for i in range(steps + 1):
        a = a0 + theta * (i / steps)
        pts.append((sx + r * math.cos(a), sy + r * math.sin(a)))
    return pts

def _segments_from_open_lwpoly(pl: LWPolyline, m: Matrix44) -> List[LineString]:
    if pl.closed:
        return []
    xyb = _lwpoly_points_xyb(pl)
    segs: List[LineString] = []
    for i in range(len(xyb) - 1):
        (x0, y0, b0), (x1, y1, _) = xyb[i], xyb[i + 1]
        pts = _bulge_arc_points((x0, y0), (x1, y1), b0)
        W = [m.transform((px, py, 0.0)) for (px, py) in pts]
        for j in range(len(W) - 1):
            (xa, ya, _), (xb, yb, _) = W[j], W[j + 1]
            if xa == xb and ya == yb:
                continue
            segs.append(LineString([(xa, ya), (xb, yb)]))
    return segs

def _segments_from_open_polyline(pl: Polyline, m: Matrix44) -> List[LineString]:
    if pl.is_closed:
        return []
    verts = [(v.dxf.location.x, v.dxf.location.y) for v in pl.vertices]
    segs: List[LineString] = []
    for i in range(len(verts) - 1):
        (x0, y0), (x1, y1) = verts[i], verts[i + 1]
        (xa, ya, _) = m.transform((x0, y0, 0.0))
        (xb, yb, _) = m.transform((x1, y1, 0.0))
        if xa == xb and ya == yb:
            continue
        segs.append(LineString([(xa, ya), (xb, yb)]))
    return segs

def _buffered(poly: Polygon) -> Polygon:
    return poly.buffer(TOUCH_TOL, cap_style=2, join_style=2)

# =========================================================
#             FLATTEN MODELSPACE + INSERTS
# =========================================================
def walk_entities(space, m: Matrix44, depth=0) -> Iterable[Tuple[DXFEntity, Matrix44]]:
    """Yield (entity, world_matrix). Traverses INSERTs up to MAX_BLOCK_DEPTH."""
    for e in space:
        if isinstance(e, Insert) and depth < MAX_BLOCK_DEPTH:
            try:
                # ezdxf virtual_entities() returns already transformed in newer versions,
                # but we still multiply by the insert's matrix for safety.
                im = e.matrix if hasattr(e, "matrix") else Matrix44()
                bm = m @ im
                for ve in e.virtual_entities():
                    yield from walk_entities([ve], bm, depth + 1)
            except Exception:
                # If expansion fails, skip this INSERT
                continue
        else:
            yield e, m

# =========================================================
#           COLLECT ZONES + WALL SEGMENTS (AUTO)
# =========================================================
def collect_zones_and_lines(doc: ezdxf.EzDxf):
    zones: List[Polygon] = []
    lines: List[Tuple[str, LineString]] = []  # (handle, geometry)

    ms = doc.modelspace()
    for ent, m in walk_entities(ms, Matrix44()):
        try:
            if isinstance(ent, LWPolyline):
                # zone (closed)
                poly = _as_polygon_from_lwpoly(ent, m)
                if poly:
                    zones.append(poly)
                # walls (open segments)
                for ls in _segments_from_open_lwpoly(ent, m):
                    if ls.length > 0:
                        lines.append((ent.dxf.handle or "", ls))

            elif isinstance(ent, Polyline):
                # zone (closed)
                poly = _as_polygon_from_polyline(ent, m)
                if poly:
                    zones.append(poly)
                # walls (open segments)
                for ls in _segments_from_open_polyline(ent, m):
                    if ls.length > 0:
                        lines.append((ent.dxf.handle or "", ls))

            elif isinstance(ent, Hatch):
                zones.extend(_as_polygons_from_hatch(ent, m))

            elif isinstance(ent, Line):
                ls = _as_linestring_from_line(ent, m)
                if ls:
                    lines.append((ent.dxf.handle or "", ls))

        except Exception:
            continue

    # Merge overlapping/touching zones to keep topology sane
    if zones:
        try:
            merged = unary_union(zones)
            zones = [merged] if isinstance(merged, Polygon) else list(merged.geoms)
        except Exception:
            pass

    return zones, lines

# =========================================================
#                  COVERAGE + OUTPUT
# =========================================================
def coverage(line: LineString, zones: List[Polygon]) -> List[Tuple[int, float, float]]:
    """Return [(zone_index, inside_len, pct), ...] for this line segment."""
    total = float(line.length)
    out: List[Tuple[int, float, float]] = []
    if total <= 0 or not zones:
        return out
    for i, poly in enumerate(zones, start=1):
        z = _buffered(poly)
        inter = line.intersection(z)
        if inter.is_empty:
            inside = 0.0
        else:
            parts = [inter] if inter.geom_type == "LineString" else list(getattr(inter, "geoms", []))
            inside = float(sum(g.length for g in parts if g.length > 0))
        if inside > 0:
            out.append((i, inside, (inside / total) * 100.0))
    return out

def rows_long_format(dxf_path: Path) -> List[Dict[str, object]]:
    doc = ezdxf.readfile(str(dxf_path))
    zones, walls = collect_zones_and_lines(doc)
    print(f"  · {dxf_path.name}: zones={len(zones)} wall_segments={len(walls)}")

    rows: List[Dict[str, object]] = []
    for handle, ls in walls:
        total = float(ls.length)
        cov = coverage(ls, zones)
        if not cov:
            continue
        for (idx, inside, pct) in cov:
            rows.append({
                "dxf_file": dxf_path.name,
                "line_handle": handle,
                "line_total_len": round(total, 6),
                "zone_id": f"R{idx}",
                "inside_len": round(inside, 6),
                "inside_pct": round(pct, 6),
            })
    return rows

def write_csv(rows: List[Dict[str, object]], out_csv: Path):
    out_csv.parent.mkdir(parents=True, exist_ok=True)
    fields = ["dxf_file","line_handle","line_total_len","zone_id","inside_len","inside_pct"]
    with out_csv.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        for r in rows:
            w.writerow(r)

def post_to_sheet(rows: List[Dict[str, object]]):
    import requests
    if not rows:
        return {"status": "no_rows"}
    payload = {
        "sheet_id": GSHEET_ID,
        "tab": GSHEET_TAB,
        "mode": GSHEET_MODE,
        "rows": rows,
        "schema": ["dxf_file","line_handle","line_total_len","zone_id","inside_len","inside_pct"],
        "meta": {"source":"auto_line_zone_coverage"}
    }
    r = requests.post(GS_WEBAPP_URL, json=payload, timeout=60)
    try:
        return r.json()
    except Exception:
        return {"status": r.status_code, "text": r.text[:300]}

# =========================================================
#                         MAIN
# =========================================================
def main():
    in_dir = Path(DXF_FOLDER)
    out_dir = Path(OUT_ROOT)
    dxf_files = sorted([*in_dir.glob("*.dxf"), *in_dir.glob("*.DXF")])

    all_rows: List[Dict[str, object]] = []
    for dxf in dxf_files:
        rows = rows_long_format(dxf)
        out_csv = out_dir / f"{dxf.stem}{CSV_SUFFIX}"
        write_csv(rows, out_csv)
        print(f"[OK] {dxf.name}: {len(rows)} rows → {out_csv}")
        all_rows.extend(rows)

    if PUSH_TO_SHEET and all_rows:
        print("[SHEET]", post_to_sheet(all_rows))
    elif not all_rows:
        print("No rows produced (either no closed zones found or no LINE/open-polylines).")

if __name__ == "__main__":
    main()
