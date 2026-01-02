// ====== GLOBAL CONFIG ======
const DEFAULT_SHEET_ID = "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM";

/**
 * ✅ Quick deploy verification
 * Open your /exec URL in browser (GET) and you should see this text.
 */
function doGet() {
  return ContentService.createTextOutput(
    "DXF Sheets WebApp ✅ single-doPost (fixed colorOnly + always-merge-colA) " + new Date().toISOString()
  );
}

// ====== SMALL UTILS ======
function parseBool_(v) {
  if (v === true || v === 1) return true;
  const s = String(v || "").trim().toLowerCase();
  return s === "true" || s === "1" || s === "yes";
}
function normalizeHeader_(s) {
  return String(s || "").toLowerCase().replace(/\s+/g, " ").trim();
}
function normKey_(s) {
  return String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");
}

function ensurePreviewColumn_(sh) {
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() >= 1 && lastCol > 0) {
    const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const idx = hdr.findIndex(h => normalizeHeader_(h) === "preview");
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
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const idx = hdr.findIndex(h => normalizeHeader_(h) === normalizeHeader_(name));
  return idx >= 0 ? idx + 1 : 0;
}

function removeColumnsByHeader_(sh, names) {
  if (!names || !names.length) return;
  const want = new Set(names.map(n => normalizeHeader_(n)));
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() < 1 || lastCol < 1) return;
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  for (let c = hdr.length - 1; c >= 0; c--) {
    if (want.has(hdr[c])) sh.deleteColumn(c + 1);
  }
}

function getOrCreateFolder_(folderId) {
  if (folderId) {
    try { return DriveApp.getFolderById(folderId); } catch (e) {}
  }
  const it = DriveApp.getFoldersByName("DXF-Previews");
  return it.hasNext() ? it.next() : DriveApp.createFolder("DXF-Previews");
}

// ====== MERGE HELPERS ======

function applyMerge_(sheet, col, s, e, value, anchor) {
  const band = sheet.getRange(s, col, e - s + 1, 1);
  band.clearContent();
  const anchorRow =
    (anchor === "first") ? s :
    (anchor === "middle") ? Math.floor((s + e) / 2) : e;

  sheet.getRange(anchorRow, col).setValue(value);
  band.merge().setVerticalAlignment("middle").setHorizontalAlignment("center");
}

/**
 * Merge vertical bands for the given header names (equal → same band).
 * NOTE: this one treats blanks as "continue previous".
 */
function mergeBandsByHeaders_(sheet, headerNames, anchor) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;

  const headers = values[0].map(normalizeHeader_);
  const r1 = 2, rN = values.length;

  headerNames.forEach(name => {
    const idx = headers.indexOf(normalizeHeader_(name));
    if (idx < 0) return;

    const col = idx + 1;
    sheet.getRange(r1, col, rN - r1 + 1, 1).breakApart();

    let s = r1;
    let vn = "";
    let vRaw = "";

    for (let r = r1; r <= rN + 1; r++) {
      const raw = (r <= rN) ? String(values[r - 1][idx] || "") : "\u0000__END__";
      const norm = normalizeHeader_(raw);

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

  const headers = values[0].map(normalizeHeader_);
  const tIdx = headers.indexOf(normalizeHeader_(targetHeader));
  const gIdx = headers.indexOf(normalizeHeader_(groupHeader));
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

/**
 * ✅ Vertical band merge for ONE column (your green highlight behavior)
 * - carries blanks forward (blank continues previous band)
 * - keeps value only in top cell
 */
function mergeCategoryBands_(sh, col, r1, rN) {
  if (rN < r1) return;

  sh.getRange(r1, col, rN - r1 + 1, 1).breakApart();

  const vals = sh.getRange(r1, col, rN - r1 + 1, 1)
                 .getDisplayValues()
                 .map(r => String(r[0] || "").trim());

  let start = r1;
  let current = vals[0]; // may be blank

  for (let i = 1; i <= vals.length; i++) {
    const raw = (i < vals.length) ? vals[i] : "__END__";

    // carry blanks forward
    const vEff = raw ? raw : current;

    if (i < vals.length && vEff === current) continue;

    const end = r1 + i - 1;

    if (current && end > start) {
      // clear content except top
      if (end > start) sh.getRange(start + 1, col, end - start, 1).clearContent();

      const band = sh.getRange(start, col, end - start + 1, 1);
      band.merge();
      band.setVerticalAlignment("middle").setHorizontalAlignment("center");

    }

    start = r1 + i;
    // IMPORTANT: new band only starts when raw is non-blank
    current = raw;
  }
}

/**
 * Merge runs in target column, but ONLY within the same group column value.
 * (Used for merging zone inside category in ByLayer.)
 */
function mergeRunsInColumnWithinGroup_(sheet, targetCol, groupCol, r1, rN) {
  if (rN < r1) return;

  sheet.getRange(r1, targetCol, rN - r1 + 1, 1).breakApart();

  const tVals = sheet.getRange(r1, targetCol, rN - r1 + 1, 1).getDisplayValues().map(r => String(r[0]||"").trim());
  const gVals = sheet.getRange(r1, groupCol,  rN - r1 + 1, 1).getDisplayValues().map(r => String(r[0]||"").trim());

  let runStart = r1;
  let lastKey = (gVals[0].toUpperCase()) + "||" + (tVals[0].toUpperCase());

  for (let i = 1; i <= tVals.length; i++) {
    const key = (i < tVals.length)
      ? (gVals[i].toUpperCase()) + "||" + (tVals[i].toUpperCase())
      : "\u0000__END__";

    if (i < tVals.length && key === lastKey) continue;

    const runEnd = r1 + i - 1;
    const val = tVals[i - 1];
    if (val && runEnd > runStart) {
      // merge + keep top cell only
      sheet.getRange(runStart + 1, targetCol, runEnd - runStart, 1).clearContent();
      const band = sheet.getRange(runStart, targetCol, runEnd - runStart + 1, 1);
      band.merge().setVerticalAlignment("middle").setHorizontalAlignment("left");
    }

    if (i < tVals.length) {
      runStart = r1 + i;
      lastKey = key;
    }
  }
}

// ====== PREVIEW-BY-NAME HANDLER ======
function handlePreviewByName_(p, sheetId) {
  const ss  = SpreadsheetApp.openById(sheetId);
  const tab = String(p.tab || "Fts");
  const sh  = ss.getSheetByName(tab) || ss.insertSheet(tab);
  const folder = getOrCreateFolder_(String(p.driveFolderId || ""));

  const baseKey = s => normKey_(s).replace(/\d+$/, "");

  const header = sh.getRange(1,1,1,Math.max(1, sh.getLastColumn()))
                   .getValues()[0].map(String);
  const cBOQ   = header.findIndex(h => normalizeHeader_(h) === "boq name") + 1;
  let   cPrev  = header.findIndex(h => normalizeHeader_(h) === "preview") + 1;
  if (!cPrev) cPrev = ensurePreviewColumn_(sh);
  if (!cBOQ) throw new Error('Column "BOQ name" not found');

  const r1 = 2, rN = sh.getLastRow();
  if (rN < r1) return { matched: 0, wrote: 0 };

  const names = sh.getRange(r1, cBOQ, rN - r1 + 1, 1)
                  .getDisplayValues().map(r => String(r[0]||""));

  const idxBase = new Map();
  names.forEach((n, i) => {
    const row = r1 + i;
    const kb  = baseKey(n);
    if (kb) { if (!idxBase.has(kb)) idxBase.set(kb, []); idxBase.get(kb).push(row); }
  });

  const items = Array.isArray(p.items) ? p.items : [];
  let matched = 0, wrote = 0;

  items.forEach((it, j) => {
    const rawName = it && it.name;
    const b64 = String(it && it.imageB64 || "");
    if (!rawName || !b64) return;

    const kb = baseKey(rawName);
    const rows = idxBase.get(kb) || [];
    if (!rows.length) return;

    const fileName = ("boq_" + kb + "_" + (j+1)).slice(0,120) + ".png";
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);
    const file = folder.createFile(blob);
    try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }
    catch (_) { try { file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW); } catch(_){} }
    const url = "https://drive.google.com/uc?export=view&id=" + file.getId();

    rows.forEach(r => {
      sh.getRange(r, cPrev).setFormula('=IMAGE("' + url + '")')
        .setHorizontalAlignment("center").setVerticalAlignment("middle");
      wrote++;
    });
    matched++;
  });

  if (wrote) {
    sh.setColumnWidth(cPrev, 50);
    sh.setRowHeights(r1, rN - r1 + 1, 50);
  }

  return { matched, wrote };
}

// ====== MAIN WEB APP ENTRYPOINT ======
function doPost(e) {
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) ? e.postData.contents : "{}");

    const sheetId = String(p.sheetId || DEFAULT_SHEET_ID).trim();
    if (!sheetId || sheetId === "undefined" || sheetId === "null") {
      throw new Error("Missing/invalid sheetId");
    }

    // Route: previews by BOQ name
    if (String(p.op || "") === "previewByName") {
      const res = handlePreviewByName_(p, sheetId);
      return ContentService.createTextOutput(JSON.stringify({ ok: true, ...res }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const ss  = SpreadsheetApp.openById(sheetId);
    const tab = String(p.tab || "Detail").trim();
    const sh  = ss.getSheetByName(tab) || ss.insertSheet(tab);

    const mode       = String(p.mode || "replace").toLowerCase(); // replace | append
    const headersIn  = Array.isArray(p.headers) ? p.headers : null;
    const rowsIn     = Array.isArray(p.rows) ? p.rows : [];
    const images     = Array.isArray(p.images) ? p.images : [];
    const colors     = Array.isArray(p.bgColors) ? p.bgColors : [];
    const colorOnly  = parseBool_(p.colorOnly); // ✅ FIXED
    const vAlign     = String(p.vAlign || "");
    const runId      = String(p.runId || "run");
    const folderId   = String(p.driveFolderId || "");

    const IMG_W = Number(p.imageW || 42);
    const IMG_H = Number(p.imageH || 42);
    const PAD_W = 8, PAD_H = 8;

    // Prepare headers/rows
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

    // Replace/Append
    let startRow;
    if (mode === "replace") {
      sh.clearContents();
      sh.clearFormats();
      if (headers && headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        startRow = 2;
      } else startRow = 1;
    } else {
      const last = sh.getLastRow();
      startRow = last ? last + 1 : (headers ? 2 : 1);
      if (startRow === 1 && headers && headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        startRow = 2;
      }
    }

    // Write rows
    if (rows.length) {
      const nCols = Math.max(...rows.map(r => r.length));
      if (sh.getMaxColumns() < nCols) sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
      if (sh.getMaxRows() < startRow - 1 + rows.length) {
        sh.insertRowsAfter(sh.getMaxRows(), startRow - 1 + rows.length - sh.getMaxRows());
      }
      sh.getRange(startRow, 1, rows.length, nCols).setValues(rows);
    }

    // Formatting
    if (rows.length) {
      const rng = sh.getRange(startRow, 1, rows.length, sh.getLastColumn());
      rng.setHorizontalAlignment("center");
      if (vAlign === "middle") rng.setVerticalAlignment("middle");
    }

    // Preview column + sizing
    const previewCol = ensurePreviewColumn_(sh);
    if (rows.length && previewCol) {
      sh.setColumnWidth(previewCol, IMG_W + PAD_W);
      sh.setRowHeights(startRow, rows.length, IMG_H + PAD_H);
    }

    // Previews / swatches
    if (rows.length && previewCol) {
      if (colorOnly) {
        for (let i = 0; i < rows.length; i++) {
          const hex = (colors[i] || "").toString().trim();
          if (hex) sh.getRange(startRow + i, previewCol).setBackground(hex);
        }
      } else if (images.length) {
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

    // Post-write shaping
    if (colorOnly) {
      // BYLAYER: sort + merge category bands + merge zone inside category
      try { const f = sh.getFilter(); if (f) f.remove(); } catch (_) {}

      const colCategory = colIndexByHeader_(sh, "category");
      const colZone     = colIndexByHeader_(sh, "zone");
      const r1 = 2, rN = sh.getLastRow();

      if (rN >= r1 && colCategory > 0) {
        const lastCol = sh.getLastColumn();
        const sortSpec = [{ column: colCategory, ascending: true }];
        if (colZone > 0) sortSpec.push({ column: colZone, ascending: true });
        sh.getRange(r1, 1, rN - r1 + 1, lastCol).sort(sortSpec);

        mergeCategoryBands_(sh, colCategory, r1, rN);

        if (colZone > 0) {
          mergeRunsInColumnWithinGroup_(sh, colZone, colCategory, r1, rN);
        }

        sh.getRange(r1, colCategory, rN - r1 + 1, 1)
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("left");
      }
    } else {
      // DETAIL: normalize blank zone → misc, sort zones (misc last), merges
      const colZone = colIndexByHeader_(sh, "zone");
      if (colZone > 0) {
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

        // sort by zone (misc last)
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

        mergeBandsByHeaders_(sh, ["zone"], "first");
        mergeColumnByGroup_(sh, "category1", "zone", "first");
      }

      removeColumnsByHeader_(sh, ["entity_type", "category"]);
    }

    // ✅ ALWAYS merge Column A into vertical bands (your green highlight)
    try { const f = sh.getFilter(); if (f) f.remove(); } catch (_) {}
    const r1A = 2, rNA = sh.getLastRow();
    if (rNA >= r1A) mergeCategoryBands_(sh, 1, r1A, rNA);

    return ContentService.createTextOutput(JSON.stringify({
      ok: true,
      wrote: rows.length,
      tab,
      colorOnly
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("Error in doPost: " + err);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
