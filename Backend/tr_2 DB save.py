import os
import csv
from collections import Counter, defaultdict

import ezdxf

DXF_FOLDER = r"C:\Users\admin\Documents\AUTOCAD_WEBAPP\DXF"
OUT_TOP = os.path.join(DXF_FOLDER, "top_level_inserts.csv")
OUT_NESTED = os.path.join(DXF_FOLDER, "nested_inserts_expanded.csv")


def safe(name):
    return "" if name is None else str(name)


def walk_block_inserts(doc, block_name, depth=0, max_depth=20, _stack=None):
    """
    Recursively count nested INSERT names inside a block definition.
    Returns Counter({nested_block_name: count_within_one_outer_block}).

    Prevents infinite loops for self-referencing blocks.
    """
    if _stack is None:
        _stack = set()

    c = Counter()

    if depth > max_depth:
        return c

    bn = safe(block_name)
    if not bn:
        return c

    if bn in _stack:
        # cyclic reference safeguard
        return c

    # Block might not exist (corrupt refs)
    if bn not in doc.blocks:
        return c

    _stack.add(bn)
    try:
        blk = doc.blocks.get(bn)
        for e in blk:
            if e.dxftype() == "INSERT":
                child_name = safe(e.dxf.name)
                if child_name:
                    c[child_name] += 1

                    # recurse into that child block too
                    c += walk_block_inserts(
                        doc,
                        child_name,
                        depth=depth + 1,
                        max_depth=max_depth,
                        _stack=_stack,
                    )
    except Exception:
        pass
    finally:
        _stack.remove(bn)

    return c


def main():
    if not os.path.isdir(DXF_FOLDER):
        raise FileNotFoundError(f"Folder not found: {DXF_FOLDER}")

    dxf_files = [f for f in os.listdir(DXF_FOLDER) if f.lower().endswith(".dxf")]
    if not dxf_files:
        print(f"❌ No DXF files in {DXF_FOLDER}")
        return

    top_level_counts = Counter()
    expanded_nested_counts = Counter()

    for fname in dxf_files:
        path = os.path.join(DXF_FOLDER, fname)
        print(f"Reading: {path}")

        try:
            doc = ezdxf.readfile(path)
            msp = doc.modelspace()
        except Exception as e:
            print(f"  ❌ Skipped: {e}")
            continue

        # 1) Count top-level inserts in modelspace
        model_inserts = [ins for ins in msp.query("INSERT")]
        for ins in model_inserts:
            outer_name = safe(ins.dxf.name)
            if not outer_name:
                continue
            top_level_counts[(fname, outer_name)] += 1

        # 2) Expand nested inserts for each outer insert placement
        #    Approach: for each distinct outer_name, compute nested breakdown once,
        #    then multiply by how many times that outer_name appears in modelspace.
        per_outer_model_count = Counter(safe(ins.dxf.name) for ins in model_inserts if safe(ins.dxf.name))

        for outer_name, outer_qty in per_outer_model_count.items():
            nested_per_one = walk_block_inserts(doc, outer_name)

            # Add nested counts multiplied by number of occurrences of outer block
            for child_name, child_qty_in_one in nested_per_one.items():
                expanded_nested_counts[(fname, outer_name, child_name)] += outer_qty * child_qty_in_one

    # Write top-level CSV
    with open(OUT_TOP, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["source_dxf", "top_level_block_name", "count_in_modelspace"])
        for (src, name), cnt in sorted(top_level_counts.items(), key=lambda x: (-x[1], x[0][0], x[0][1])):
            w.writerow([src, name, cnt])

    # Write expanded nested CSV
    with open(OUT_NESTED, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["source_dxf", "outer_block_name", "nested_block_name", "expanded_count"])
        for (src, outer, child), cnt in sorted(expanded_nested_counts.items(), key=lambda x: (-x[1], x[0][0], x[0][1], x[0][2])):
            w.writerow([src, outer, child, cnt])

    print("\n✅ Done")
    print(f"Top-level CSV:   {OUT_TOP}")
    print(f"Nested CSV:      {OUT_NESTED}")


if __name__ == "__main__":
    main()
