#!/usr/bin/env python3
# ==========================================================
# Universal DXF Envelope Metrics (INSERT-aware, resilient, XREF-aware)
# - Expands INSERTs; warns if inserts are XREFs (external) which ezdxf can't resolve
# - Optional debug scan to show where your geometry actually is
# - Optional paperspace scan fallback
#
# Requires: pip install ezdxf shapely numpy
# Python: 3.7+
# ==========================================================

from __future__ import annotations
import argparse, math, re, sys, csv
from pathlib import Path
from typing import List, Tuple, Optional, Dict, Set

import numpy as np
import ezdxf
from ezdxf import recover
from ezdxf.entities import DXFGraphic, Insert
from ezdxf.math import Vec2
from shapely.geometry import LineString, MultiLineString
from shapely.ops import unary_union, polygonize
import typing  # add this near the top

# ======= Default Paths (edit) =======
DEFAULT_DXF_FOLDER = r"C:\Users\admin\Documents\VIZ_AUTOCAD_NEW\DXF"
DEFAULT_OUT_FOLDER = r"C:\Users\admin\Documents\VIZ_AUTOCAD_NEW\EXPORTS"
# ====================================

# -------------------- Text parsing --------------------
import re
HEIGHT_RE = re.compile(r"(?:facade|glaz(?:e|ing)|ext(?:ernal)?)?\s*(?:height|ht)\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)", re.I)
SILL_RE   = re.compile(r"(?:sill)\s*(?:height|ht)?\s*[:=]?\s*([0-9]+(?:\.[0-9]+)?)", re.I)

def split_csv(s: str) -> List[str]:
    return [t.strip().lower() for t in s.split(",") if t.strip()] if s else []

def any_token_in(name: str, tokens: List[str]) -> bool:
    n = (name or "").lower()
    return any(t in n for t in tokens)

def parse_heights(doc, text_layer_hints: List[str]) -> Tuple[Optional[float], Optional[float]]:
    h = s = None
    msp = doc.modelspace()
    def layer_ok(ent) -> bool:
        if not text_layer_hints: return True
        try: return any_token_in(ent.dxf.layer, text_layer_hints)
        except Exception: return False
    def scan(val: str):
        nonlocal h, s
        if val:
            if h is None:
                m = HEIGHT_RE.search(val);  h = float(m.group(1)) if m else h
            if s is None:
                m = SILL_RE.search(val);    s = float(m.group(1)) if m else s
    for e in msp:
        t = e.dxftype()
        if t in ("TEXT", "MTEXT") and layer_ok(e):
            scan(e.dxf.text if t=="TEXT" else e.text)
        elif t == "INSERT" and layer_ok(e):
            try:
                for a in e.attribs():
                    scan(a.dxf.text or "")
            except Exception:
                pass
        if h is not None and s is not None:
            break
    return h, s

# -------------------- Units --------------------
def insunits_to_m(doc) -> float:
    u = int(doc.header.get("$INSUNITS", 0))
    return {
        0: 1.0,   # unitless
        1: 0.0254, # inches
        2: 0.3048, # feet
        4: 0.001, # mm
        5: 0.01,  # cm
        6: 1.0,   # meters
    }.get(u, 1.0)

def out_scale_factor(out_units: str) -> float:
    return {"m":1.0,"cm":100.0,"mm":1000.0,"ft":3.28084}[out_units]

# -------------------- DXF Reader (permissive + helpful) --------------------
def load_dxf_or_explain(path: Path):
    try:
        return ezdxf.readfile(path)
    except Exception:
        pass
    try:
        doc, auditor = recover.readfile(path)
        if auditor.has_errors:
            print(f"[WARN] DXF recovered with issues for {path.name}.")
        return doc
    except Exception:
        pass
    try:
        with open(path, "rb") as f:
            head = f.read(4096)
    except Exception as e:
        raise IOError(f"Cannot open '{path}': {e}")
    low = head.lower()
    if head.startswith(b"AC10"):
        raise IOError(f"'{path.name}' looks like a DWG (AC10xx). Please BIND/SAVE AS DXF 2013.")
    if b"<html" in low or b"<!doctype html" in low:
        raise IOError(f"'{path.name}' is an HTML download, not a DXF. Re-download the real file.")
    if (b"SECTION" in head[:1024]) and (b"\n0" in head[:1024]):
        raise IOError(f"'{path.name}' looks like a DXF but is unreadable. RECOVER in AutoCAD, then Save As DXF 2013.")
    raise IOError(f"'{path.name}' does not look like a DXF. Save As AutoCAD 2013 DXF and retry.")

# -------------------- Helpers --------------------
def sample_entity_lines(e) -> List[Tuple[Tuple[float,float],Tuple[float,float]]]:
    segs = []
    try:
        t = e.dxftype()
        if t == "LINE":
            p1, p2 = (e.dxf.start.x, e.dxf.start.y), (e.dxf.end.x, e.dxf.end.y)
            segs.append((p1,p2))
        elif t in ("LWPOLYLINE", "POLYLINE"):
            closed = bool(getattr(e, "closed", False)) if t=="LWPOLYLINE" else bool(getattr(e, "is_closed", False))
            flat = list(e.flattening(0.2))
            pts = [(p[0],p[1]) for p in flat]
            for i in range(len(pts)-1):
                segs.append((pts[i], pts[i+1]))
            if closed and len(pts)>1:
                segs.append((pts[-1], pts[0]))
        elif t == "ARC":
            cx, cy, r = e.dxf.center.x, e.dxf.center.y, e.dxf.radius
            a0 = math.radians(e.dxf.start_angle); a1 = math.radians(e.dxf.end_angle)
            if a1 < a0: a1 += 2*math.pi
            steps = max(8, int(abs(a1-a0)/ (math.pi/18)))
            last = (cx + r*math.cos(a0), cy + r*math.sin(a0))
            for k in range(1,steps+1):
                a = a0 + (a1-a0)*k/steps
                cur = (cx + r*math.cos(a), cy + r*math.sin(a))
                segs.append((last, cur)); last = cur
        elif t in ("CIRCLE","ELLIPSE","SPLINE"):
            flat = list(e.flattening(0.2))
            pts = [(p[0],p[1]) for p in flat]
            for i in range(len(pts)-1):
                segs.append((pts[i], pts[i+1]))
            if len(pts)>1:
                segs.append((pts[-1], pts[0]))
    except Exception:
        pass
    return segs

def is_closed_entity(e) -> bool:
    t = e.dxftype()
    if t == "LWPOLYLINE": return bool(getattr(e, "closed", False))
    if t == "POLYLINE":   return bool(getattr(e, "is_closed", False))
    return t in ("CIRCLE","ELLIPSE")

def entity_bbox(e):
    try:
        b = e.bbox()
        if b.has_data:
            (x1,y1,_),(x2,y2,_) = b.extmin, b.extmax
            return (float(x1),float(y1),float(x2),float(y2))
    except Exception:
        pass
    return None

# ------------- XREF detection & INSERT expansion -------------
def is_xref_insert(doc, ins: Insert) -> bool:
    try:
        # effective/explicit block name
        name = getattr(ins, "effective_name", None) or getattr(ins, "block_name", None) or ins.dxf.name
        if not name: return False
        blk = doc.blocks.get(name)
        # According to DXF spec, XREF flag bit (4) on block record marks external reference
        flags = int(getattr(blk.block_record.dxf, "flags", 0) or 0)
        return bool(flags & 4)
    except Exception:
        # If block not found in this DXF, also treat as XREF-like
        return True

def iter_drawables(space, doc, xref_counter: list) -> "typing.Iterator[DXFGraphic]":
    """
    Yield drawable entities from a layout space, expanding INSERTs.
    If an INSERT is an XREF, count it (can't expand) and skip.
    """
    for e in space:
        t = e.dxftype()
        if t in ("HATCH","DIMENSION","IMAGE"):
            continue
        if t == "INSERT":
            if is_xref_insert(doc, e):
                xref_counter.append(1)
                continue
            try:
                for ve in e.virtual_entities():
                    vt = ve.dxftype()
                    if vt in ("HATCH","DIMENSION","IMAGE"):
                        continue
                    yield ve
            except Exception:
                # could not expand; treat as empty
                continue
        else:
            yield e

# -------------------- Wall / Column detection --------------------
def footprint_perimeter(e) -> float:
    try:
        t = e.dxftype()
        if t in ("LWPOLYLINE","POLYLINE"):
            return float(e.length())
        elif t == "CIRCLE":
            return 2*math.pi*e.dxf.radius
        elif t == "ELLIPSE":
            pts = [Vec2(p) for p in e.flattening(0.2)]
            return sum(math.hypot(pts[i+1].x-pts[i].x, pts[i+1].y-pts[i].y) for i in range(len(pts)-1))
        else:
            segs = sample_entity_lines(e)
            return sum(math.hypot(x2-x1,y2-y1) for (x1,y1),(x2,y2) in segs)
    except Exception:
        return 0.0

def classify_columns(closed_entities: List[DXFGraphic], size_band: Tuple[float,float]) -> Tuple[float,int]:
    perim = 0.0; count = 0
    smin, smax = size_band
    for e in closed_entities:
        bb = entity_bbox(e)
        if not bb: continue
        x1,y1,x2,y2 = bb
        w, h = abs(x2-x1), abs(y2-y1)
        size = max(min(w,h), 0.0)
        ar = max(w,h) / (min(w,h)+1e-9)
        if smin <= size <= smax and ar <= 1.6:
            perim += footprint_perimeter(e)
            count += 1
    return perim, count

def detect_wall_polys(closed_entities: List[DXFGraphic], thick_band: Tuple[float,float], min_len=1.0, ar_min=3.0) -> float:
    s = 0.0
    tmin, tmax = thick_band
    for e in closed_entities:
        bb = entity_bbox(e)
        if not bb: continue
        x1,y1,x2,y2 = bb
        w, h = abs(x2-x1), abs(y2-y1)
        thick = min(w,h); long_side = max(w,h)
        if long_side < min_len: continue
        ar = long_side / max(thick,1e-9)
        if tmin <= thick <= tmax and ar >= ar_min:
            s += long_side
    return s

def pair_parallel_open_lines(open_segs, thick_band, ang_tol_deg=5.0, gap_tol=0.15) -> float:
    if not open_segs: return 0.0
    angs, mids, vecs = [], [], []
    for (x1,y1),(x2,y2) in open_segs:
        vx, vy = (x2-x1, y2-y1)
        L = math.hypot(vx,vy)
        if L < 0.5:
            angs.append(None); mids.append((x1,y1)); vecs.append((vx,vy)); continue
        angs.append(math.degrees(math.atan2(vy,vx)) % 180.0)
        mids.append(((x1+x2)/2.0, (y1+y2)/2.0))
        vecs.append((vx/L, vy/L))
    used: Set[int] = set()
    total = 0.0
    tmin, tmax = thick_band
    def seg_len(i):
        (a,b) = open_segs[i]
        return math.hypot(b[0]-a[0], b[1]-a[1])
    for i in range(len(open_segs)):
        if i in used: continue
        ai = angs[i]
        if ai is None: continue
        best = None; best_gap = 1e9
        for j in range(i+1, len(open_segs)):
            if j in used: continue
            aj = angs[j]
            if aj is None: continue
            dang = min(abs(ai-aj), 180-abs(ai-aj))
            if dang > ang_tol_deg: continue
            mx,my = mids[i]; nx,ny = mids[j]
            vx,vy = vecs[i]
            px,py = -vy, vx
            gap = abs((nx-mx)*px + (ny-my)*py)
            if tmin <= gap <= tmax*(1.0+gap_tol) and gap < best_gap:
                best = j; best_gap = gap
        if best is not None:
            used.add(i); used.add(best)
            total += 0.5*(seg_len(i)+seg_len(best))
    return total

# -------------------- Facade via polygonize (INSERT-aware) --------------------
def facade_length_from_all(space, doc, tiny_buffer: float) -> float:
    segs = []
    xrefs = []
    for e in iter_drawables(space, doc, xrefs):
        if not isinstance(e, DXFGraphic): continue
        segs.extend(sample_entity_lines(e))
    if not segs:
        return 0.0
    ml = MultiLineString([LineString([p1,p2]) for p1,p2 in segs])
    fused = unary_union(ml.buffer(tiny_buffer, join_style=2).boundary)
    polys = list(polygonize(fused))
    if not polys:
        return 0.0
    biggest = max(polys, key=lambda p: p.area)
    return float(biggest.exterior.length)

# -------------------- Measurement core --------------------
def measure_space(space, doc, units: str, in_to_m: float,
                  wall_tmin: float, wall_tmax: float,
                  col_smin: float, col_smax: float,
                  facade_gap: float, walls_layers: str,
                  columns_layers: str, facade_layers: str,
                  block_hints: str, text_layers: str,
                  facade_height: Optional[float], sill_height: Optional[float],
                  debug_scan: bool=False, label: str="Model") -> Dict[str, object]:

    out_mul = out_scale_factor(units)
    closed_entities: List[DXFGraphic] = []
    open_segments: List[Tuple[Tuple[float,float],Tuple[float,float]]] = []
    hint_blocks = split_csv(block_hints)
    layer_wall = split_csv(walls_layers)
    layer_col  = split_csv(columns_layers)
    layer_fac  = split_csv(facade_layers)

    # Collect geometry with INSERT expansion; count xrefs
    xref_counter: List[int] = []
    type_tally: Dict[str,int] = {}
    layer_set: Set[str] = set()

    for e in iter_drawables(space, doc, xref_counter):
        if not isinstance(e, DXFGraphic): continue
        t = e.dxftype(); type_tally[t] = type_tally.get(t,0)+1
        try: layer_set.add((e.dxf.layer or "").strip())
        except Exception: pass

        if is_closed_entity(e) or t in ("SPLINE","ELLIPSE","CIRCLE"):
            closed_entities.append(e)
        segs = sample_entity_lines(e)
        if segs: open_segments.extend(segs)

    # Debug scan output
    if debug_scan:
        total_segs = len(open_segments)
        print(f"[DEBUG/{label}] expanded types={type_tally}  xref_inserts={len(xref_counter)}  open_segs={total_segs}  closed_candidates={len(closed_entities)}")
        if layer_set:
            L = sorted(layer_set)
            print(f"[DEBUG/{label}] sample layers ({min(10,len(L))}/{len(L)}): " + ", ".join(L[:10]))

    # Columns
    col_perim, col_count = classify_columns(closed_entities, (col_smin / in_to_m, col_smax / in_to_m))
    if layer_col or hint_blocks:
        extra = 0.0; extra_count = 0
        for e in closed_entities:
            ok = False
            try:
                if layer_col and any_token_in(e.dxf.layer, layer_col): ok = True
                if not ok and e.dxftype()=="INSERT":
                    if any_token_in(e.dxf.name, hint_blocks): ok = True
            except Exception:
                pass
            if ok:
                extra += footprint_perimeter(e); extra_count += 1
        col_perim += extra; col_count += extra_count

    # Walls
    walls_len = detect_wall_polys(closed_entities, (wall_tmin / in_to_m, wall_tmax / in_to_m))
    walls_len += pair_parallel_open_lines(open_segments, (wall_tmin / in_to_m, wall_tmax / in_to_m))
    if layer_wall:
        for e in iter_drawables(space, doc, []):
            try:
                if any_token_in(e.dxf.layer, layer_wall):
                    for (x1,y1),(x2,y2) in sample_entity_lines(e):
                        walls_len += math.hypot(x2-x1,y2-y1)
            except Exception:
                pass

    # Facade
    facade_len = facade_length_from_all(space, doc, facade_gap / in_to_m)
    if layer_fac:
        for e in iter_drawables(space, doc, []):
            try:
                if any_token_in(e.dxf.layer, layer_fac):
                    for (x1,y1),(x2,y2) in sample_entity_lines(e):
                        facade_len += math.hypot(x2-x1,y2-y1)
            except Exception:
                pass

    # Heights from MODELSPACE only (text usually there)
    auto_h = auto_s = None
    try:
        auto_h, auto_s = parse_heights(doc, split_csv(text_layers))
    except Exception:
        pass

    fh = facade_height if facade_height is not None else auto_h
    sh = sill_height   if sill_height   is not None else auto_s

    # Scale to output units
    res = {
        "units": units,
        "walls_length": round(walls_len * in_to_m * out_mul, 3),
        "columns_length": round(col_perim * in_to_m * out_mul, 3),
        "facade_length": round(facade_len * in_to_m * out_mul, 3),
        "facade_height": "" if fh is None else round((fh) * out_mul, 3),
        "facade_sill_height": "" if sh is None else round((sh) * out_mul, 3),
        "columns_count": col_count,
        "_debug": {
            "types": type_tally,
            "xref_inserts": len(xref_counter),
            "open_segs": len(open_segments),
            "closed_candidates": len(closed_entities),
            "layers_seen": sorted(layer_set)[:20],
        }
    }
    return res

def measure_file(dxf_path: Path, units: str, in_scale: Optional[float],
                 wall_tmin: float, wall_tmax: float,
                 col_smin: float, col_smax: float,
                 facade_gap: float, walls_layers: str,
                 columns_layers: str, facade_layers: str,
                 block_hints: str, text_layers: str,
                 facade_height: Optional[float], sill_height: Optional[float],
                 debug_scan: bool=False, from_paperspace: bool=False) -> Dict[str, object]:

    doc = load_dxf_or_explain(dxf_path)
    in_to_m = in_scale if in_scale is not None else insunits_to_m(doc)

    # Choose layout space
    space = doc.paperspace() if from_paperspace else doc.modelspace()
    label = "Paper" if from_paperspace else "Model"

    res = measure_space(space, doc, units, in_to_m,
                        wall_tmin, wall_tmax, col_smin, col_smax,
                        facade_gap, walls_layers, columns_layers, facade_layers,
                        block_hints, text_layers, facade_height, sill_height,
                        debug_scan=debug_scan, label=label)

    # If modelspace looked empty but paperspace has data, try paperspace automatically (debug only)
    if (not from_paperspace) and debug_scan and res["_debug"]["open_segs"] == 0 and res["_debug"]["closed_candidates"] == 0:
        try:
            pres = measure_space(doc.paperspace(), doc, units, in_to_m,
                                 wall_tmin, wall_tmax, col_smin, col_smax,
                                 facade_gap, walls_layers, columns_layers, facade_layers,
                                 block_hints, text_layers, facade_height, sill_height,
                                 debug_scan=True, label="Paper(Auto)")
            # don’t replace results silently; just show debug. If you want, enable --from-paperspace.
        except Exception:
            pass

    res.pop("_debug", None)
    res.update({"file": dxf_path.name})
    return res

# -------------------- CLI / Batch --------------------
def main():
    ap = argparse.ArgumentParser(description="Universal DXF envelope metrics (INSERT & XREF aware).")
    ap.add_argument("--dxf", default="", help="DXF file or folder. Defaults to embedded folder.")
    ap.add_argument("--out", default=DEFAULT_OUT_FOLDER, help="Output folder for CSVs.")
    ap.add_argument("--units", choices=["m","cm","mm","ft"], default="m")
    ap.add_argument("--in-scale", type=float, default=None, help="Override drawing-units→meters (e.g., 0.001 for mm).")
    # Heuristic bands (meters)
    ap.add_argument("--wall-thickness-min", type=float, default=0.075)
    ap.add_argument("--wall-thickness-max", type=float, default=0.300)
    ap.add_argument("--column-size-min", type=float, default=0.25)
    ap.add_argument("--column-size-max", type=float, default=1.20)
    ap.add_argument("--facade-gap-buffer", type=float, default=0.05)
    # Optional hints
    ap.add_argument("--walls-layers", default="")
    ap.add_argument("--columns-layers", default="")
    ap.add_argument("--facade-layers", default="")
    ap.add_argument("--block-hints", default="col,column,pier,wall,ext,facade")
    ap.add_argument("--text-layers", default="text,anno,dim,notes")
    ap.add_argument("--facade-height", type=float, default=None)
    ap.add_argument("--sill-height", type=float, default=None)
    # New debug/space controls
    ap.add_argument("--debug-scan", action="store_true", help="Print a diagnostic scan of what geometry was found.")
    ap.add_argument("--from-paperspace", action="store_true", help="Measure from paperspace instead of modelspace.")
    args = ap.parse_args()

    # Determine input target(s)
    in_arg = args.dxf.strip() or DEFAULT_DXF_FOLDER
    in_path = Path(in_arg)
    if in_path.is_dir():
        dxf_files = sorted([p for p in in_path.glob("*.dxf")])
        if not dxf_files:
            print(f"No DXF files found in folder: {in_path}")
            sys.exit(0)
    else:
        if not in_path.exists():
            print(f"DXF path not found: {in_path}")
            sys.exit(1)
        dxf_files = [in_path]

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    summary_rows: List[Dict[str,object]] = []

    for f in dxf_files:
        print(f"\n====== UNIVERSAL ENVELOPE METRICS ======")
        print(f"DXF                     : {f.name}")
        try:
            res = measure_file(
                dxf_path=f,
                units=args.units,
                in_scale=args.in_scale,
                wall_tmin=args.wall_thickness_min,
                wall_tmax=args.wall_thickness_max,
                col_smin=args.column_size_min,
                col_smax=args.column_size_max,
                facade_gap=args.facade_gap_buffer,
                walls_layers=args.walls_layers,
                columns_layers=args.columns_layers,
                facade_layers=args.facade_layers,
                block_hints=args.block_hints,
                text_layers=args.text_layers,
                facade_height=args.facade_height,
                sill_height=args.sill_height,
                debug_scan=args.debug_scan,
                from_paperspace=args.from_paperspace
            )
        except IOError as e:
            print(f"[SKIP] {e}")
            continue
        except Exception as e:
            print(f"[SKIP] Unexpected error on '{f.name}': {e}")
            continue

        # Print to console
        u = res["units"]
        print(f"Output units            : {u}")
        print(f"1) EXISTING WALLS LENGTH: {res['walls_length']:,.3f} {u}")
        print(f"2) COLUMNS LENGTH (ΣP)  : {res['columns_length']:,.3f} {u}  (count≈{res['columns_count']})")
        print(f"3) FACADE LENGTH        : {res['facade_length']:,.3f} {u}")
        fh_out = res['facade_height'] if res['facade_height'] != "" else "N/A"
        sh_out = res['facade_sill_height'] if res['facade_sill_height'] != "" else "N/A"
        print(f"4) FACADE HEIGHT        : {fh_out if isinstance(fh_out,str) else f'{fh_out:.3f} {u}'}")
        print(f"5) FACADE SILL HEIGHT   : {sh_out if isinstance(sh_out,str) else f'{sh_out:.3f} {u}'}")
        print("========================================")

        # Guidance if we found nothing
        if args.debug_scan:
            print("[HINT] If open_segs and closed_candidates are 0 and xref_inserts > 0, your geometry is inside XREFs.")
            print("       In AutoCAD: BIND Xrefs (or use ODA File Converter with 'Bind Xrefs'), then Save As DXF 2013 and re-run.")
            if args.from_paperspace:
                print("[HINT] You measured from paperspace; try without --from-paperspace if your model is in modelspace.")

        # Write per-file CSV
        csv_path = out_dir / f"{f.stem}_metrics.csv"
        try:
            with csv_path.open("w", newline="", encoding="utf-8") as cf:
                writer = csv.DictWriter(cf, fieldnames=list(res.keys()))
                writer.writeheader(); writer.writerow(res)
            print(f"Saved: {csv_path}")
        except Exception as e:
            print(f"[WARN] Could not write CSV for {f.name}: {e}")

        summary_rows.append(res)

    # Write summary CSV (all files)
    if len(summary_rows) > 1:
        summary_path = out_dir / "metrics_summary.csv"
        try:
            with summary_path.open("w", newline="", encoding="utf-8") as sf:
                writer = csv.DictWriter(sf, fieldnames=list(summary_rows[0].keys()))
                writer.writeheader(); writer.writerows(summary_rows)
            print(f"\nSummary saved: {summary_path}")
        except Exception as e:
            print(f"[WARN] Could not write summary CSV: {e}")

if __name__ == "__main__":
    main()
