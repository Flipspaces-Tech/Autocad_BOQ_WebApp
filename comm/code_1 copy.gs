/** ====== CONFIG (optional) ====== */
const one = ""; // keep blank if you always pass sheetId from backend

/**
 * ✅ Quick deploy verification
 * Open your /exec URL in browser (GET) and you should see this text.
 * If you don't, your Python is calling an OLD deployment URL.
 */
function doGet() {
  return ContentService.createTextOutput(
    "DXF Sheets WebApp ✅ merge-v3 (category bands + zone-in-category) " + new Date().toISOString()
  );
}

/** Web App: Detail vs ByLayer handling */
function doPost(e) {
  try {
    // ✅ SAFE: editor Run() sends undefined e; real webapp calls send e
    const p = JSON.parse((e && e.postData && e.postData.contents) ? e.postData.contents : "{}");

    const sheetId = String(p.sheetId || DEFAULT_SHEET_ID).trim();
    if (!sheetId) throw new Error("Missing sheetId");

    const ss  = SpreadsheetApp.openById(sheetId);
    const tab = String(p.tab || "Detail").trim();

    // Get/Create tab
    const sh = ss.getSheetByName(tab) || ss.insertSheet(tab);

    const mode       = String(p.mode || "replace").toLowerCase(); // replace | append
    const headersIn  = Array.isArray(p.headers) ? p.headers : null;
    const rowsIn     = Array.isArray(p.rows) ? p.rows : [];
    const images     = Array.isArray(p.images) ? p.images : [];
    const colors     = Array.isArray(p.bgColors) ? p.bgColors : [];
    const colorOnly  = !!p.colorOnly; // true → ByLayer
    const vAlign     = String(p.vAlign || "");
    const runId      = String(p.runId || "run");
    const folderId   = String(p.driveFolderId || "");

    // Preview sizing
    const IMG_W = Number(p.imageW || 42);
    const IMG_H = Number(p.imageH || 42);
    const PAD_W = 8, PAD_H = 8;

    // ---------- 0) Prepare headers/rows (branch on colorOnly) ----------
    let headers = headersIn ? headersIn.slice() : null;
    let rows    = rowsIn.slice();

    if (!colorOnly && headers) {
      // DETAIL: drop entity_type, category
      const kill = new Set(["entity_type", "category"]);
      const keepIdx = headers
        .map((h, i) => ({ i, keep: !kill.has(String(h).trim().toLowerCase()) }))
        .filter(x => x.keep)
        .map(x => x.i);

      headers = keepIdx.map(i => headersIn[i]);
      rows    = rows.map(r => keepIdx.map(i => r[i]));
    }
    // BYLAYER: keep as-is

    // ---------- 1) Replace/Append: clear & write header ----------
    let startRow;
    if (mode === "replace") {
      sh.clearContents();
      sh.clearFormats(); // clears old merges + formats
      if (headers && headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        startRow = 2;
      } else {
        startRow = 1;
      }
    } else {
      const last = sh.getLastRow();
      startRow = last ? last + 1 : (headers ? 2 : 1);
      // If appending and sheet is empty, write header
      if (startRow === 1 && headers && headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        startRow = 2;
      }
    }

    // ---------- 2) Write rows ----------
    if (rows.length) {
      const nCols = Math.max(...rows.map(r => r.length));
      if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
      if (sh.getMaxRows() < startRow - 1 + rows.length) {
        sh.insertRowsAfter(sh.getMaxRows(), startRow - 1 + rows.length - sh.getMaxRows());
      }
      sh.getRange(startRow, 1, rows.length, nCols).setValues(rows);
    }

    // ---------- 3) Basic formatting ----------
    if (rows.length) {
      const rng = sh.getRange(startRow, 1, rows.length, sh.getLastColumn());
      rng.setHorizontalAlignment("center");
      if (vAlign === "middle") rng.setVerticalAlignment("middle");
    }

    // ---------- 4) Ensure Preview column ----------
    const previewCol = ensurePreviewColumn_(sh);

    // ---------- 5) Size preview column / new rows ----------
    if (rows.length && previewCol) {
      sh.setColumnWidth(previewCol, IMG_W + PAD_W);
      sh.setRowHeights(startRow, rows.length, IMG_H + PAD_H);
    }

    // ---------- 6) Previews / Color swatches ----------
    if (rows.length && previewCol) {
      if (colorOnly) {
        // BYLAYER: set background color only (swatch)
        for (let i = 0; i < rows.length; i++) {
          const hex = (colors[i] || "").toString().trim();
          if (hex) sh.getRange(startRow + i, previewCol).setBackground(hex);
        }
      } else {
        // DETAIL: upload PNGs & write =IMAGE(url)
        if (images.length) {
          const folder = getOrCreateFolder_(folderId);
          for (let i = 0; i < rows.length; i++) {
            const b64 = images[i] || "";
            if (!b64) continue;
            const r = startRow + i;
            const fileName = (runId + "_" + r).replace(/[^\w\-\.]/g, "_") + ".png";
            const blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);
            const file = folder.createFile(blob);
            try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
            catch (_) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch (_) {} }
            const url = "https://drive.google.com/uc?export=view&id=" + file.getId();
            sh.getRange(r, previewCol)
              .setFormula('=IMAGE("' + url + '")')
              .setHorizontalAlignment("center")
              .setVerticalAlignment("middle");
          }
        }
      }
    }

    // ---------- 7) Post-write shaping ----------
    if (colorOnly) {
      // ✅ BYLAYER: force Image2 layout (merged category vertical bands)

      // Remove filter if any (filters can block merges in some cases)
      try {
        const f = sh.getFilter();
        if (f) f.remove();
      } catch (_) {}

      const colCategory = colIndexByHeader_(sh, "category");
      const colZone     = colIndexByHeader_(sh, "zone");

      const r1 = 2;
      const rN = sh.getLastRow();
      if (rN >= r1 && colCategory > 0) {
        // (A) Sort by category then zone so duplicates are contiguous
        const lastCol = sh.getLastColumn();
        const sortSpec = [{ column: colCategory, ascending: true }];
        if (colZone > 0) sortSpec.push({ column: colZone, ascending: true });
        sh.getRange(r1, 1, rN - r1 + 1, lastCol).sort(sortSpec);

        // (B) Merge CATEGORY runs (this is what you want)
        // (B) Merge CATEGORY vertical bands (exact green-highlight behaviour)
        mergeCategoryBands_(sh, colCategory, r1, rN);


        // (C) Merge ZONE within CATEGORY (optional but matches your Image2)
        if (colZone > 0) {
          mergeRunsInColumnWithinGroup_(sh, colZone, colCategory, r1, rN, {
            anchor: "first",
            normalize: v => String(v || "").trim()
          });
        }

        // (D) Cosmetics like Image2
        sh.getRange(r1, colCategory, rN - r1 + 1, 1)
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("left");
      }

    } else {
      // DETAIL: your existing logic (zone normalization + sort + merges + drop entity_type/category)
      const colZone = colIndexByHeader_(sh, "zone");
      if (colZone > 0) {
        // blanks → "misc"
        const r1 = 2, rN = sh.getLastRow();
        if (rN >= r1) {
          const zoneRng = sh.getRange(r1, colZone, rN - r1 + 1, 1);
          const Z = zoneRng.getValues();
          let changed = false;
          for (let i = 0; i < Z.length; i++) {
            const s = String(Z[i][0] || "").trim();
            if (!s) { Z[i][0] = "misc"; changed = true; }
          }
          if (changed) zoneRng.setValues(Z);
        }

        // sort by zone, forcing "misc" last
        const lastCol = sh.getLastColumn();
        const r1s = 2, rNs = sh.getLastRow();
        if (rNs >= r1s) {
          sh.insertColumnAfter(lastCol);
          const skCol = lastCol + 1;
          sh.getRange(1, skCol).setValue("__sort_zone__");

          const vals = sh.getRange(r1s, colZone, rNs - r1s + 1, 1).getValues();
          const keys = vals.map(v => {
            const s = String(v[0] || "").toLowerCase().trim();
            return (s === "misc") ? "zzzzzz" : s;
          }).map(k => [k]);

          sh.getRange(r1s, skCol, keys.length, 1).setValues(keys);
          sh.getRange(r1s, 1, rNs - r1s + 1, skCol).sort([{ column: skCol, ascending: true }]);
          sh.deleteColumn(skCol);
        }

        // merges
        mergeBandsByHeaders_(sh, ["zone"], "first");
        mergeColumnByGroup_(sh, "category1", "zone", "first");
      }

      // remove any lingering entity_type/category columns if present
      removeColumnsByHeader_(sh, ["entity_type", "category"]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        wrote: rows.length,
        tab,
        colorOnly,
        version: "merge-v3"
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ===== Helpers ===== */

function ensurePreviewColumn_(sh) {
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() >= 1 && lastCol > 0) {
    const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const idx = hdr.findIndex(h => (h || "").toString().trim().toLowerCase() === "preview");
    if (idx >= 0) return idx + 1;
    sh.insertColumnAfter(lastCol);
    sh.getRange(1, lastCol + 1).setValue("Preview");
    return lastCol + 1;
  }
  sh.getRange(1, 1).setValue("Preview");
  return 1;
}

function colIndexByHeader_(sh, name) {
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() < 1 || lastCol < 1) return 0;
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = hdr.findIndex(h => String(h).trim().toLowerCase() === String(name).trim().toLowerCase());
  return idx >= 0 ? idx + 1 : 0;
}

function removeColumnsByHeader_(sh, names) {
  if (!names || !names.length) return;
  const set = new Set(names.map(n => String(n).trim().toLowerCase()));
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() < 1 || lastCol < 1) return;
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim().toLowerCase());
  for (let c = hdr.length - 1; c >= 0; c--) {
    if (set.has(hdr[c])) sh.deleteColumn(c + 1);
  }
}

/** Merge vertical bands for the given header names (equal → same band). */
function mergeBandsByHeaders_(sheet, headerNames, anchor) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const r1 = 2, rN = values.length;

  headerNames.forEach(name => {
    const idx = headers.indexOf(String(name).trim().toLowerCase());
    if (idx < 0) return;

    const col = idx + 1;
    sheet.getRange(r1, col, rN - r1 + 1, 1).breakApart();

    let s = r1;
    let vn = "";
    let vRaw = "";

    for (let r = r1; r <= rN + 1; r++) {
      const raw = (r <= rN) ? String(values[r - 1][idx] || "") : "\u0000__END__";
      const norm = raw.trim().toUpperCase();

      if (!vn) {
        if (norm) { s = r; vn = norm; vRaw = raw; }
        continue;
      }

      const cont = (r <= rN) && (!norm || norm === vn);
      if (cont) continue;

      const e = r - 1;
      if (e > s) applyMerge_(sheet, col, s, e, vRaw, anchor);

      vn = ""; vRaw = "";
      if (r <= rN && norm) { s = r; vn = norm; vRaw = raw; }
    }
  });
}

/** Merge contiguous duplicates in target column, within the same group column. */
function mergeColumnByGroup_(sheet, targetHeader, groupHeader, anchor) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const tIdx = headers.indexOf(String(targetHeader).trim().toLowerCase());
  const gIdx = headers.indexOf(String(groupHeader).trim().toLowerCase());
  if (tIdx < 0 || gIdx < 0) return;

  const tCol = tIdx + 1;
  const r1 = 2, rN = values.length;
  sheet.getRange(r1, tCol, rN - r1 + 1, 1).breakApart();

  let runStart = r1;
  let lastGroup = String(values[r1 - 1][gIdx] || "").trim().toUpperCase();
  let lastVal   = String(values[r1 - 1][tIdx] || "").trim().toUpperCase();

  for (let r = r1 + 1; r <= rN + 1; r++) {
    const cg = (r <= rN) ? String(values[r - 1][gIdx] || "").trim().toUpperCase() : "\u0000__END__";
    const cv = (r <= rN) ? String(values[r - 1][tIdx] || "").trim().toUpperCase() : "\u0000__END__";

    if (r <= rN && cg === lastGroup && cv === lastVal) continue;

    const runEnd = r - 1;
    if (lastVal && runEnd > runStart) {
      applyMerge_(sheet, tCol, runStart, runEnd, values[runStart - 1][tIdx], anchor);
    }

    if (r <= rN) {
      runStart = r;
      lastGroup = cg;
      lastVal = cv;
    }
  }
}

function applyMerge_(sheet, col, s, e, value, anchor) {
  const band = sheet.getRange(s, col, e - s + 1, 1);
  band.clearContent();
  const anchorRow =
    (anchor === "first") ? s :
    (anchor === "middle") ? Math.floor((s + e) / 2) : e;

  sheet.getRange(anchorRow, col).setValue(value);
  band.merge().setVerticalAlignment("middle").setHorizontalAlignment("center");
}

function getOrCreateFolder_(folderId) {
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) {}
  }
  const it = DriveApp.getFoldersByName("DXF-Previews");
  return it.hasNext() ? it.next() : DriveApp.createFolder("DXF-Previews");
}

/* =========================
   ✅ NEW (strong merge helpers)
   ========================= */

/**
 * Merge runs in one column using values from that column only.
 * - carryBlanks=true means blank cells continue previous run (good for "category" bands).
 */
function mergeRunsInColumn_(sheet, col, r1, rN, opt) {
  const anchor = (opt && opt.anchor) || "first";
  const carryBlanks = !!(opt && opt.carryBlanks);
  const normalize = (opt && opt.normalize) || (v => String(v || "").trim());

  if (rN < r1) return;

  // break only this column
  sheet.getRange(r1, col, rN - r1 + 1, 1).breakApart();

  const vals = sheet.getRange(r1, col, rN - r1 + 1, 1).getDisplayValues().map(r => r[0]);

  let runStart = r1;
  let last = normalize(vals[0]);

  for (let i = 1; i <= vals.length; i++) {
    const v = (i < vals.length) ? normalize(vals[i]) : "\u0000__END__";
    const vEff = (carryBlanks && !v) ? last : v;

    if (i < vals.length && vEff === last) continue;

    const runEnd = r1 + i - 1;
    if (last && runEnd > runStart) {
      const rawAnchor = sheet.getRange(runStart, col).getDisplayValue();
      safeMerge_(sheet, col, runStart, runEnd, rawAnchor, anchor);
    }

    if (i < vals.length) {
      runStart = r1 + i;
      last = vEff;
    }
  }
}

/**
 * Merge runs in target column, but ONLY within the same group column value.
 * (Used for merging zone inside a category band.)
 */
function mergeRunsInColumnWithinGroup_(sheet, targetCol, groupCol, r1, rN, opt) {
  const anchor = (opt && opt.anchor) || "first";
  const normalize = (opt && opt.normalize) || (v => String(v || "").trim());

  if (rN < r1) return;

  sheet.getRange(r1, targetCol, rN - r1 + 1, 1).breakApart();

  const tVals = sheet.getRange(r1, targetCol, rN - r1 + 1, 1).getDisplayValues().map(r => r[0]);
  const gVals = sheet.getRange(r1, groupCol,  rN - r1 + 1, 1).getDisplayValues().map(r => r[0]);

  let runStart = r1;
  let lastKey = (String(gVals[0] || "").trim().toUpperCase()) + "||" + normalize(tVals[0]).toUpperCase();

  for (let i = 1; i <= tVals.length; i++) {
    const key = (i < tVals.length)
      ? (String(gVals[i] || "").trim().toUpperCase()) + "||" + normalize(tVals[i]).toUpperCase()
      : "\u0000__END__";

    if (i < tVals.length && key === lastKey) continue;

    const runEnd = r1 + i - 1;
    const zoneVal = normalize(tVals[i - 1]);
    if (zoneVal && runEnd > runStart) {
      const rawAnchor = sheet.getRange(runStart, targetCol).getDisplayValue();
      safeMerge_(sheet, targetCol, runStart, runEnd, rawAnchor, anchor);
    }

    if (i < tVals.length) {
      runStart = r1 + i;
      lastKey = key;
    }
  }
}

function safeMerge_(sheet, col, s, e, value, anchor) {
  try {
    const band = sheet.getRange(s, col, e - s + 1, 1);
    band.clearContent();
    const anchorRow =
      (anchor === "first") ? s :
      (anchor === "middle") ? Math.floor((s + e) / 2) : e;

    sheet.getRange(anchorRow, col).setValue(value);
    band.merge().setVerticalAlignment("middle").setHorizontalAlignment("center");
  } catch (err) {
    // If merge fails for any reason, restore at least the top cell so data isn't lost
    try { sheet.getRange(s, col).setValue(value); } catch (_) {}
  }
}

/**
 * ✅ Merges Column A category into tall vertical bands (green-highlight behavior)
 * - Merges only when adjacent rows have same category
 * - Keeps value only in the FIRST row of band; clears the rest
 * - Aligns left + vertically middle like your screenshot
 */
function mergeCategoryBands_(sh, col, r1, rN) {
  if (rN < r1) return;

  // Unmerge first so fresh merges work
  sh.getRange(r1, col, rN - r1 + 1, 1).breakApart();

  const rng = sh.getRange(r1, col, rN - r1 + 1, 1);
  const vals = rng.getValues().map(r => String(r[0] || "").trim());

  let start = r1;
  let last = vals[0];

  for (let i = 1; i <= vals.length; i++) {
    const v = (i < vals.length) ? vals[i] : "__END__";

    if (i < vals.length && v === last) continue;

    const end = r1 + i - 1;

    // merge band if more than 1 row and not blank
    if (last && end > start) {
      // clear content except top cell (so you don't lose category text)
      if (end > start) {
        sh.getRange(start + 1, col, end - start, 1).clearContent();
      }

      const band = sh.getRange(start, col, end - start + 1, 1);
      band.merge();
      band.setVerticalAlignment("middle").setHorizontalAlignment("left");
    }

    start = r1 + i;
    last = v;
  }
}

