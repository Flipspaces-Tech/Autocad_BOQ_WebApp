#!/usr/bin/env python3
# ==========================================================
# DXF Metrics Checker (exterior footprint + zones + units)
# - Auto-picks DXF from DEFAULT_DXF_FOLDER if not provided
# - Reports EXTERIOR FOOTPRINT area & perimeter (gross)
# - Reports SUM OF CLOSED POLYS (zones-ish)
# - Normalizes units and prints in ft², in², m²
# Requires: pip install ezdxf shapely
# ==========================================================

from __future__ import annotations
import argparse, sys, math
from pathlib import Path

import ezdxf
from ezdxf import recover
from ezdxf.entities import LWPolyline, Polyline, Line
from shapely.geometry import Polygon, MultiPolygon
from shapely.ops import unary_union

# ---------- Defaults ----------
DEFAULT_DXF_FOLDER = Path(r"C:\Users\admin\Documents\VIZ_AUTOCAD_NEW\DXF")

# ---------- Units helpers ----------
INSUNITS_MAP = {
    0: ("unitless", 1.0),           # no scaling
    1: ("inches", 1.0),             # base = inches
    2: ("feet", 12.0),              # 1 ft = 12 in
    3: ("miles", 12.0*5280.0),
    4: ("millimeters", 1.0/25.4),   # 1 mm = 1/25.4 in
    5: ("centimeters", 1.0/2.54),
    6: ("meters", 39.37007874),     # 1 m in inches
    7: ("kilometers", 39370.07874),
}

def dxf_linear_units_in_inches(doc) -> float:
    code = int(doc.header.get("$INSUNITS", 0))
    _, to_in = INSUNITS_MAP.get(code, ("unitless", 1.0))
    return float(to_in)

def to_pretty_area(in2: float) -> dict:
    # Convert square inches → other units
    ft2 = in2 / 144.0
    m2  = in2 * (0.0254**2)
    return {"in2": in2, "ft2": ft2, "m2": m2}

def to_pretty_length(inches: float) -> dict:
    feet = inches / 12.0
    meters = inches * 0.0254
    return {"in": inches, "ft": feet, "m": meters}

# ---------- DXF loading ----------
def get_default_dxf() -> Path | None:
    if not DEFAULT_DXF_FOLDER.exists():
        print(f"❌ Folder not found: {DEFAULT_DXF_FOLDER}")
        return None
    dxfs = sorted(DEFAULT_DXF_FOLDER.glob("*.dxf"))
    if not dxfs:
        print(f"⚠️ No DXF files found in {DEFAULT_DXF_FOLDER}")
        return None
    print(f"✅ Using DXF: {dxfs[0].name}")
    return dxfs[0]

def load_doc(path: Path):
    try:
        doc = ezdxf.readfile(str(path))
        return doc, None
    except ezdxf.DXFError:
        print("⚠️ DXF error, attempting recover()...", file=sys.stderr)
        doc, auditor = recover.readfile(str(path))
        if auditor.has_errors:
            print(f"⚠️ recover() found {len(auditor.errors)} issues.", file=sys.stderr)
        return doc, auditor

# ---------- Geometry helpers ----------
def is_closed_poly(e) -> bool:
    try:
        if isinstance(e, LWPolyline):
            return bool(e.closed)
        if isinstance(e, Polyline):
            return bool(e.is_closed)
        return False
    except Exception:
        return False

def poly_to_coords(e):
    if isinstance(e, LWPolyline):
        return [(p[0], p[1]) for p in e.get_points("xy")]
    if isinstance(e, Polyline):
        return [(v.dxf.location.x, v.dxf.location.y) for v in e.vertices]
    return []

def collect_closed_polys(msp, layer_filter: set[str] | None = None):
    polys = []
    for e in msp:
        try:
            if layer_filter and e.dxf.layer not in layer_filter:
                continue
            if isinstance(e, (LWPolyline, Polyline)) and is_closed_poly(e):
                coords = poly_to_coords(e)
                if len(coords) >= 3:
                    if coords[0] != coords[-1]:
                        coords = coords + [coords[0]]
                    poly = Polygon(coords)
                    if poly.is_valid and not poly.is_empty:
                        polys.append(poly)
        except Exception:
            continue
    return polys

def sum_open_lengths(msp, layer_filter: set[str] | None = None) -> float:
    """Sum open segments (walls-as-lines), returned in DXF native units."""
    total = 0.0
    for e in msp:
        try:
            if layer_filter and e.dxf.layer not in layer_filter:
                continue
            if isinstance(e, Line):
                p1, p2 = e.dxf.start, e.dxf.end
                total += math.hypot(p2.x - p1.x, p2.y - p1.y)
            elif isinstance(e, (LWPolyline, Polyline)) and not is_closed_poly(e):
                coords = poly_to_coords(e)
                for (x1, y1), (x2, y2) in zip(coords, coords[1:]):
                    total += math.hypot(x2 - x1, y2 - y1)
        except Exception:
            continue
    return total

# ---------- Main ----------
def main():
    ap = argparse.ArgumentParser(description="DXF metrics (exterior footprint + zones + unit conversions)")
    ap.add_argument("dxf", nargs="?", help="Path to DXF (optional).")
    ap.add_argument("--zones-layers", nargs="*", default=[],
                    help="Layers to consider for zones (closed polys). Default: ALL")
    ap.add_argument("--walls-layers", nargs="*", default=[],
                    help="Layers to consider for wall length (open lines). Default: ALL")
    ap.add_argument("--verbose", "-v", action="store_true")
    args = ap.parse_args()

    dxf_path = Path(args.dxf) if args.dxf else get_default_dxf()
    if not dxf_path or not dxf_path.exists():
        print("❌ No DXF file found. Exiting.")
        return

    doc, _ = load_doc(dxf_path)
    msp = doc.modelspace()

    # Normalize units → inches
    to_in = dxf_linear_units_in_inches(doc)

    # 1) ZONES (sum of all closed polylines)
    zones_filter = set(args.zones_layers) if args.zones_layers else None
    polys = collect_closed_polys(msp, zones_filter)
    zones_area_native = sum(p.area for p in polys)                      # native^2
    zones_area_in2 = zones_area_native * (to_in ** 2)                   # → in²

    # 2) EXTERIOR FOOTPRINT (gross): union of all closed polys, take largest shell
    exterior_area_in2 = 0.0
    exterior_perim_in = 0.0
    if polys:
        merged = unary_union(polys)
        # merged can be Polygon or MultiPolygon
        candidates = []
        if isinstance(merged, Polygon):
            candidates = [merged]
        elif isinstance(merged, MultiPolygon):
            candidates = list(merged.geoms)
        else:
            candidates = []

        if candidates:
            # choose the largest outer shell (gross)
            biggest = max(candidates, key=lambda g: g.area)
            exterior_area_in2 = biggest.area * (to_in ** 2)
            exterior_perim_in = biggest.exterior.length * to_in

    # 3) WALL LENGTH (optional—open segments)
    walls_filter = set(args.walls_layers) if args.walls_layers else None
    wall_len_native = sum_open_lengths(msp, walls_filter)               # native
    wall_len_in = wall_len_native * to_in                               # → inches

    # 4) Report nicely
    zones = to_pretty_area(zones_area_in2)
    gross = to_pretty_area(exterior_area_in2)
    perim = to_pretty_length(exterior_perim_in)

    print("==== DXF METRICS ====")
    print(f"File:   {dxf_path.name}")
    print(f"Folder: {dxf_path.parent}")
    print(f"Entities in modelspace: {len(msp)}\n")

    # Exterior footprint (gross)
    if exterior_area_in2 > 0:
        print("EXTERIOR FOOTPRINT (gross):")
        print(f"  Area: {gross['in2']:.0f} in²  | {gross['ft2']:.3f} ft²  | {gross['m2']:.3f} m²")
        # perimeter in feet + inches pretty
        ft_total = perim["ft"]
        ft_whole = int(ft_total)
        inches = round((ft_total - ft_whole) * 12.0, 2)
        print(f"  Perimeter: {perim['in']:.1f} in  | {perim['ft']:.3f} ft  (~ {ft_whole}′-{inches:.2f}″)")
    else:
        print("EXTERIOR FOOTPRINT (gross): not found (no closed polys to merge)")

    print("")

    # Zones (sum of all closed polys)
    print("SUM OF CLOSED POLYS (zones-ish):")
    print(f"  Area: {zones['in2']:.0f} in²  | {zones['ft2']:.3f} ft²  | {zones['m2']:.3f} m²")
    print(f"  Polygons counted: {len(polys)}")
    if args.verbose:
        # quick peek
        biggest5 = sorted((p.area for p in polys), reverse=True)[:5]
        biggest5_in2 = [a * (to_in**2) for a in biggest5]
        print("  Top 5 poly areas (in²):", [f"{a:.0f}" for a in biggest5_in2])

    print("")
    # Wall length
    print("WALL LENGTH (open segments):")
    print(f"  {wall_len_in:.1f} in  | {wall_len_in/12.0:.2f} ft  | {wall_len_in*0.0254:.2f} m")
    print("\nCompare EXTERIOR FOOTPRINT with AutoCAD AREA on the outer boundary.")
    print("If they match, your units + exterior merge are good.")
    print("=====================")

if __name__ == "__main__":
    main()
