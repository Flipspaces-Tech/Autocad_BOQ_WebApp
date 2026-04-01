// ====== GLOBAL CONFIG ======
const DEFAULT_SHEET_ID = "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM";
const UNMARKED_ZONE = "Unmarked Area";

/**
 * ✅ Quick deploy verification
 * Open your /exec URL in browser (GET) and you should see this text.
 */
function doGet() {
  return ContentService.createTextOutput(
    "DXF Sheets WebApp ✅ single-doPost (detail: agg by cat+boq+zone; unmarked zone fill; bylayer zone fill; merge boq then cat; bylayer merge; preview dedupe+merge; Params column enabled; NO setSharing) " +
      new Date().toISOString()
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
function stripZoneCountSuffix_(z) {
  // "Conference (5)" -> "Conference"
  return String(z || "").replace(/\s*\(\d+\)\s*$/g, "").trim();
}
function toNumberOr_(v, fallback) {
  const n = parseFloat(String(v ?? "").trim());
  return isFinite(n) ? n : fallback;
}
function isBlank_(v) {
  return String(v ?? "").trim() === "";
}

function ensurePreviewColumn_(sh) {
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() >= 1 && lastCol > 0) {
    const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const idx = hdr.findIndex((h) => normalizeHeader_(h) === "preview");
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
  const idx = hdr.findIndex((h) => normalizeHeader_(h) === normalizeHeader_(name));
  return idx >= 0 ? idx + 1 : 0;
}

function removeColumnsByHeader_(sh, names) {
  if (!names || !names.length) return;
  const want = new Set(names.map((n) => normalizeHeader_(n)));
  const lastCol = sh.getLastColumn();
  if (sh.getLastRow() < 1 || lastCol < 1) return;
  const hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(normalizeHeader_);
  for (let c = hdr.length - 1; c >= 0; c--) {
    if (want.has(hdr[c])) sh.deleteColumn(c + 1);
  }
}

function getOrCreateFolder_(folderId) {
  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (e) {}
  }
  const it = DriveApp.getFoldersByName("DXF-Previews");
  return it.hasNext() ? it.next() : DriveApp.createFolder("DXF-Previews");
}

// ====== MERGE HELPERS ======
function applyMerge_(sheet, col, s, e, value, anchor, hAlign) {
  const band = sheet.getRange(s, col, e - s + 1, 1);
  band.clearContent();
  const anchorRow =
    anchor === "first" ? s : anchor === "middle" ? Math.floor((s + e) / 2) : e;

  sheet.getRange(anchorRow, col).setValue(value);
  band.merge().setVerticalAlignment("middle").setHorizontalAlignment(hAlign || "center");
}

/**
 * ✅ Vertical band merge for ONE column
 * - carries blanks forward (blank continues previous band)
 * - keeps value only in top cell
 */
function mergeCategoryBands_(sh, col, r1, rN) {
  if (rN < r1) return;

  sh.getRange(r1, col, rN - r1 + 1, 1).breakApart();

  const vals = sh
    .getRange(r1, col, rN - r1 + 1, 1)
    .getDisplayValues()
    .map((r) => String(r[0] || "").trim());

  let start = r1;
  let current = vals[0]; // may be blank

  for (let i = 1; i <= vals.length; i++) {
    const raw = i < vals.length ? vals[i] : "__END__";

    // carry blanks forward
    const vEff = raw ? raw : current;

    if (i < vals.length && vEff === current) continue;

    const end = r1 + i - 1;

    if (current && end > start) {
      if (end > start) sh.getRange(start + 1, col, end - start, 1).clearContent();
      const band = sh.getRange(start, col, end - start + 1, 1);
      band.merge();
      band.setVerticalAlignment("middle").setHorizontalAlignment("center");
    }

    start = r1 + i;
    // new band only starts when raw is non-blank
    current = raw;
  }
}

/**
 * Merge runs in target column, but ONLY within the same group column value.
 * ✅ supports optional horizontal alignment
 */
function mergeRunsInColumnWithinGroup_(sheet, targetCol, groupCol, r1, rN, hAlign) {
  if (rN < r1) return;
  hAlign = hAlign || "left";

  sheet.getRange(r1, targetCol, rN - r1 + 1, 1).breakApart();

  const tVals = sheet
    .getRange(r1, targetCol, rN - r1 + 1, 1)
    .getDisplayValues()
    .map((r) => String(r[0] || "").trim());

  const gValsRaw = sheet
    .getRange(r1, groupCol, rN - r1 + 1, 1)
    .getDisplayValues()
    .map((r) => String(r[0] || "").trim());

  // ✅ carry blanks forward
  const gVals = [];
  let lastG = "";
  for (let i = 0; i < gValsRaw.length; i++) {
    const g = gValsRaw[i];
    if (g) lastG = g;
    gVals.push(g ? g : lastG);
  }

  let runStart = r1;
  let lastKey = gVals[0].toUpperCase() + "||" + tVals[0].toUpperCase();

  for (let i = 1; i <= tVals.length; i++) {
    const key =
      i < tVals.length
        ? gVals[i].toUpperCase() + "||" + tVals[i].toUpperCase()
        : "\u0000__END__";

    if (i < tVals.length && key === lastKey) continue;

    const runEnd = r1 + i - 1;
    const val = tVals[i - 1];

    if (val && runEnd > runStart) {
      sheet.getRange(runStart + 1, targetCol, runEnd - runStart, 1).clearContent();
      const band = sheet.getRange(runStart, targetCol, runEnd - runStart + 1, 1);
      band.merge().setVerticalAlignment("middle").setHorizontalAlignment(hAlign);
    }

    if (i < tVals.length) {
      runStart = r1 + i;
      lastKey = key;
    }
  }
}

/**
 * ✅ Merge Preview column ONLY within the same BOQ name run.
 * - BOQ blanks are carried forward
 * - Keeps the FIRST non-empty preview (formula/value) in the run
 * - Clears the rest and merges the preview cells for that BOQ run
 */
function mergePreviewWithinBoq_(sh, colPreview, colBoq, r1, rN) {
  if (rN < r1) return;

  sh.getRange(r1, colPreview, rN - r1 + 1, 1).breakApart();

  const boqRaw = sh
    .getRange(r1, colBoq, rN - r1 + 1, 1)
    .getDisplayValues()
    .map((r) => String(r[0] || "").trim());

  // carry BOQ blanks forward
  const boq = [];
  let lastBoq = "";
  for (let i = 0; i < boqRaw.length; i++) {
    if (boqRaw[i]) lastBoq = boqRaw[i];
    boq.push(boqRaw[i] ? boqRaw[i] : lastBoq);
  }

  const prevFormulas = sh.getRange(r1, colPreview, rN - r1 + 1, 1).getFormulas();
  const prevDisp = sh.getRange(r1, colPreview, rN - r1 + 1, 1).getDisplayValues();

  function hasPreviewAt(i) {
    const f = prevFormulas[i] && prevFormulas[i][0] ? String(prevFormulas[i][0]).trim() : "";
    const v = prevDisp[i] && prevDisp[i][0] ? String(prevDisp[i][0]).trim() : "";
    return !!(f || v);
  }

  let runStart = 0;
  let runKey = (boq[0] || "").toLowerCase();

  for (let i = 1; i <= boq.length; i++) {
    const key = i < boq.length ? (boq[i] || "").toLowerCase() : "\u0000__END__";

    if (i < boq.length && key === runKey) continue;

    const runEnd = i - 1;
    const startRow = r1 + runStart;
    const endRow = r1 + runEnd;

    if (runKey) {
      // find first row inside run that has a preview
      let srcIdx = -1;
      for (let k = runStart; k <= runEnd; k++) {
        if (hasPreviewAt(k)) {
          srcIdx = k;
          break;
        }
      }

      if (srcIdx >= 0 && endRow > startRow) {
        const srcRow = r1 + srcIdx;

        // move preview to top of run if needed
        if (srcRow !== startRow) {
          const srcCell = sh.getRange(srcRow, colPreview);
          const f = srcCell.getFormula();
          const v = srcCell.getValue();
          const dstCell = sh.getRange(startRow, colPreview);

          if (f) dstCell.setFormula(f);
          else if (v) dstCell.setValue(v);
        }

        // clear remaining previews in run
        sh.getRange(startRow + 1, colPreview, endRow - startRow, 1).clearContent();

        // merge preview band
        sh
          .getRange(startRow, colPreview, endRow - startRow + 1, 1)
          .merge()
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("center");
      }
    }

    runStart = i;
    runKey = key;
  }
}

// ====== DETAIL TRANSFORM (AGGREGATE BY category1 + BOQ name + zone) ======
function transformDetail_(headersIn, rowsIn, imagesIn) {
  const UNMARKED = UNMARKED_ZONE;

  // 1) Drop entity_type + category
  const kill = new Set(["entity_type", "category"]);
  const keepIdx = headersIn
    .map((h, i) => ({ i, keep: !kill.has(normalizeHeader_(h)) }))
    .filter((x) => x.keep)
    .map((x) => x.i);

  const headers2 = keepIdx.map((i) => headersIn[i]);
  const rows2 = rowsIn.map((r) => keepIdx.map((i) => r[i]));
  const images2 = (imagesIn || []).slice(0, rowsIn.length);

  // 2) Find indices
  const hnorm = headers2.map(normalizeHeader_);
  const idx = (name) => hnorm.indexOf(normalizeHeader_(name));

  const jCat1 = idx("category1");
  const jBoq  = idx("BOQ name");
  const jZone = idx("zone");
  const jQtyV = idx("qty_value");
  const jLen  = idx("length (ft)");
  const jWid  = idx("width (ft)");
  const jPrev = idx("Preview");
  const jParams = idx("Params"); // ✅ NEW

  if (jCat1 < 0 || jBoq < 0 || jZone < 0) {
    return { headers: headers2, rows: rows2, images: images2 };
  }

  // 3) Aggregate by (category1 + BOQ name + zone)
  const groups = new Map();

  for (let i = 0; i < rows2.length; i++) {
    const r = rows2[i];

    const cat1 = String(r[jCat1] || "").trim();
    const boq  = String(r[jBoq] || "").trim();
    if (!cat1 && !boq) continue;

    const zoneRaw = String(r[jZone] || "").trim();

    let zoneList = zoneRaw
      .split(/\r?\n+/)
      .map(stripZoneCountSuffix_)
      .map((s) => s.trim())
      .filter(Boolean);

    if (!zoneList.length) zoneList = [UNMARKED];

    let qty = toNumberOr_(jQtyV >= 0 ? r[jQtyV] : "", NaN);
    if (!isFinite(qty)) qty = 1;

    zoneList.forEach((z) => {
      const zEff = (!z || z.toLowerCase() === "misc") ? UNMARKED : z;
      const key = cat1 + "||" + boq + "||" + zEff;

      if (!groups.has(key)) {
        groups.set(key, {
          cat1,
          boq,
          zone: zEff,
          qty_total: 0,
          len: jLen >= 0 ? r[jLen] : "",
          wid: jWid >= 0 ? r[jWid] : "",
          params: jParams >= 0 ? String(r[jParams] || "").trim() : "", // ✅ NEW
          img: images2[i] || "",
        });
      }

      const g = groups.get(key);
      g.qty_total += qty;

      if (jLen >= 0 && isBlank_(g.len) && !isBlank_(r[jLen])) g.len = r[jLen];
      if (jWid >= 0 && isBlank_(g.wid) && !isBlank_(r[jWid])) g.wid = r[jWid];
      if (!g.params && jParams >= 0 && !isBlank_(r[jParams])) g.params = String(r[jParams]).trim(); // ✅ NEW
      if (!g.img && images2[i]) g.img = images2[i];
    });
  }

  // 4) Output columns ONLY (whitelist) + rename category1->Product
  const outHeaders = ["Product", "BOQ name", "zone", "qty_value", "length (ft)", "width (ft)", "Params", "Preview"]; // ✅ NEW
  const outRows = [];
  const outImgs = [];

  const entries = Array.from(groups.values()).sort((a, b) => {
    const c1 = a.cat1.localeCompare(b.cat1);
    if (c1) return c1;
    const c2 = a.boq.localeCompare(b.boq);
    if (c2) return c2;

    const za = a.zone.toLowerCase() === UNMARKED.toLowerCase() ? "zzzzzz" : a.zone.toLowerCase();
    const zb = b.zone.toLowerCase() === UNMARKED.toLowerCase() ? "zzzzzz" : b.zone.toLowerCase();
    return za.localeCompare(zb);
  });

  entries.forEach((g) => {
    outRows.push([
      g.cat1,
      g.boq,
      stripZoneCountSuffix_(g.zone),
      g.qty_total ? Number(g.qty_total) : "",
      g.len || "",
      g.wid || "",
      g.params || "", // ✅ NEW
      "", // Preview placeholder
    ]);
    outImgs.push(g.img || "");
  });

  return { headers: outHeaders, rows: outRows, images: outImgs };
}


// ====== PREVIEW-BY-NAME HANDLER ======
function handlePreviewByName_(p, sheetId) {
  const ss = SpreadsheetApp.openById(sheetId);
  const tab = String(p.tab || "Fts");
  const sh = ss.getSheetByName(tab) || ss.insertSheet(tab);
  const folder = getOrCreateFolder_(String(p.driveFolderId || ""));

  const baseKey = (s) => normKey_(s).replace(/\d+$/, "");

  const header = sh
    .getRange(1, 1, 1, Math.max(1, sh.getLastColumn()))
    .getValues()[0]
    .map(String);
  const cBOQ = header.findIndex((h) => normalizeHeader_(h) === "boq name") + 1;
  let cPrev = header.findIndex((h) => normalizeHeader_(h) === "preview") + 1;
  if (!cPrev) cPrev = ensurePreviewColumn_(sh);
  if (!cBOQ) throw new Error('Column "BOQ name" not found');

  const r1 = 2,
    rN = sh.getLastRow();
  if (rN < r1) return { matched: 0, wrote: 0 };

  const names = sh
    .getRange(r1, cBOQ, rN - r1 + 1, 1)
    .getDisplayValues()
    .map((r) => String(r[0] || ""));

  const idxBase = new Map();
  names.forEach((n, i) => {
    const row = r1 + i;
    const kb = baseKey(n);
    if (kb) {
      if (!idxBase.has(kb)) idxBase.set(kb, []);
      idxBase.get(kb).push(row);
    }
  });

  const items = Array.isArray(p.items) ? p.items : [];
  let matched = 0,
    wrote = 0;

  items.forEach((it, j) => {
    const rawName = it && it.name;
    const b64 = String((it && it.imageB64) || "");
    if (!rawName || !b64) return;

    const kb = baseKey(rawName);
    const rows = idxBase.get(kb) || [];
    if (!rows.length) return;

    const fileName = ("boq_" + kb + "_" + (j + 1)).slice(0, 120) + ".png";
    const blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);
    const file = folder.createFile(blob);

    // ❌ NO file.setSharing() (keeps Drive permissions as-is)

    const url = "https://drive.google.com/uc?export=view&id=" + file.getId();

    // write preview only once
    const firstRow = rows[0];
    sh.getRange(firstRow, cPrev)
      .setFormula('=IMAGE("' + url + '")')
      .setHorizontalAlignment("center")
      .setVerticalAlignment("middle");

    if (rows.length > 1) {
      sh.getRange(rows[1], cPrev, rows.length - 1, 1).clearContent();
    }

    wrote++;
    matched++;
  });

  if (wrote) {
    sh.setColumnWidth(cPrev, 50);
    sh.setRowHeights(r1, rN - r1 + 1, 50);
    mergePreviewWithinBoq_(sh, cPrev, cBOQ, r1, rN);
  }

  return { matched, wrote };
}

// ====== MAIN WEB APP ENTRYPOINT ======
function doPost(e) {
  try {
    const p = JSON.parse(e && e.postData && e.postData.contents ? e.postData.contents : "{}");

    const sheetId = String(p.sheetId || DEFAULT_SHEET_ID).trim();
    if (!sheetId || sheetId === "undefined" || sheetId === "null") {
      throw new Error("Missing/invalid sheetId");
    }

    // Route: previews by BOQ name
    if (String(p.op || "") === "previewByName") {
      const res = handlePreviewByName_(p, sheetId);
      return ContentService.createTextOutput(JSON.stringify({ ok: true, ...res })).setMimeType(
        ContentService.MimeType.JSON
      );
    }

    const ss = SpreadsheetApp.openById(sheetId);
    const tab = String(p.tab || "Detail").trim();
    const sh = ss.getSheetByName(tab) || ss.insertSheet(tab);

    const mode = String(p.mode || "replace").toLowerCase(); // replace | append
    const headersIn = Array.isArray(p.headers) ? p.headers : null;
    const rowsIn = Array.isArray(p.rows) ? p.rows : [];
    const imagesIn = Array.isArray(p.images) ? p.images : [];
    const colors = Array.isArray(p.bgColors) ? p.bgColors : [];
    const colorOnly = parseBool_(p.colorOnly);
    const vAlign = String(p.vAlign || "");
    const runId = String(p.runId || "run");
    const folderId = String(p.driveFolderId || "");

    const IMG_W = Number(p.imageW || 42);
    const IMG_H = Number(p.imageH || 42);
    const PAD_W = 8,
      PAD_H = 8;

    // Prepare headers/rows
    let headers = headersIn ? headersIn.slice() : null;
    let rows = rowsIn.slice();
    let images = imagesIn.slice();

    // DETAIL transform
    if (!colorOnly) {
      if (!headers) {
        const lastCol = sh.getLastColumn();
        if (sh.getLastRow() >= 1 && lastCol > 0) {
          headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
        }
      }
      if (headers && rows.length) {
        const t = transformDetail_(headers, rows, images);
        headers = t.headers;
        rows = t.rows;
        images = t.images;
      }
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
      startRow = last ? last + 1 : headers ? 2 : 1;
      if (startRow === 1 && headers && headers.length) {
        sh.getRange(1, 1, 1, headers.length).setValues([headers]);
        startRow = 2;
      }
    }

    // Write rows
    if (rows.length) {
      const nCols = Math.max(...rows.map((r) => r.length));
      if (sh.getMaxColumns() < nCols)
        sh.insertColumnsAfter(sh.getMaxColumns(), nCols - sh.getMaxColumns());
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

        // ✅ DEDUPE: only write preview once per BOQ name
        const colBoqNow = colIndexByHeader_(sh, "BOQ name");
        const seen = new Set();

        for (let i = 0; i < rows.length; i++) {
          const b64 = images[i] || "";
          if (!b64) continue;

          const boq = colBoqNow > 0 ? String(rows[i][colBoqNow - 1] || "").trim() : "";
          const boqKey = boq.toLowerCase();
          if (!boqKey) continue;

          if (seen.has(boqKey)) continue;
          seen.add(boqKey);

          const r = startRow + i;
          const fileName = (runId + "_" + r).replace(/[^\w\-\.]/g, "_") + ".png";
          const blob = Utilities.newBlob(Utilities.base64Decode(b64), "image/png", fileName);
          const file = folder.createFile(blob);

          // ❌ NO file.setSharing() (keeps Drive permissions as-is)

          const url = "https://drive.google.com/uc?export=view&id=" + file.getId();
          sh
            .getRange(r, previewCol)
            .setFormula('=IMAGE("' + url + '")')
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        }
      }
    }

    // Post-write shaping
    if (colorOnly) {
      try {
        const f = sh.getFilter();
        if (f) f.remove();
      } catch (_) {}

      const colCategory = colIndexByHeader_(sh, "category");
      const colZone = colIndexByHeader_(sh, "zone");
      const r1 = 2,
        rN = sh.getLastRow();

      // ✅ Fill blank/misc zones with "Unmarked Area" in ByLayer
      if (colZone > 0 && rN >= r1) {
        const zRange = sh.getRange(r1, colZone, rN - r1 + 1, 1);
        const zVals = zRange.getValues();
        for (let i = 0; i < zVals.length; i++) {
          const s = String(zVals[i][0] || "").trim();
          if (!s || s.toLowerCase() === "misc") zVals[i][0] = UNMARKED_ZONE;
        }
        zRange.setValues(zVals);
      }

      if (rN >= r1 && colCategory > 0) {
        const lastCol = sh.getLastColumn();
        const sortSpec = [{ column: colCategory, ascending: true }];
        if (colZone > 0) sortSpec.push({ column: colZone, ascending: true });
        sh.getRange(r1, 1, rN - r1 + 1, lastCol).sort(sortSpec);

        mergeCategoryBands_(sh, colCategory, r1, rN);

        if (colZone > 0) {
          mergeRunsInColumnWithinGroup_(sh, colZone, colCategory, r1, rN, "left");
        }

        sh
          .getRange(r1, colCategory, rN - r1 + 1, 1)
          .setVerticalAlignment("middle")
          .setHorizontalAlignment("center");
      }
    } else {
      try {
        const f = sh.getFilter();
        if (f) f.remove();
      } catch (_) {}

      const colCat1 = colIndexByHeader_(sh, "Product") || colIndexByHeader_(sh, "category1") || 1;

      const colBoq = colIndexByHeader_(sh, "BOQ name");
      const colZone = colIndexByHeader_(sh, "zone");
      const r1 = 2,
        rN = sh.getLastRow();

      // Sort by category1, then BOQ name, then zone (Unmarked Area last)
      if (rN >= r1 && colCat1 > 0) {
        const lastCol = sh.getLastColumn();

        if (colZone > 0) {
          sh.insertColumnAfter(lastCol);
          const skCol = lastCol + 1;
          sh.getRange(1, skCol).setValue("__sort_zone__");
          const zVals = sh.getRange(r1, colZone, rN - r1 + 1, 1).getValues();
          const keys = zVals.map((v) => {
            const s = String(v[0] || "").toLowerCase().trim();
            const un = UNMARKED_ZONE.toLowerCase();
            return [(!s || s === "misc" || s === un) ? "zzzzzz" : s];
          });
          sh.getRange(r1, skCol, keys.length, 1).setValues(keys);

          const spec = [{ column: colCat1, ascending: true }];
          if (colBoq > 0) spec.push({ column: colBoq, ascending: true });
          spec.push({ column: skCol, ascending: true });
          sh.getRange(r1, 1, rN - r1 + 1, skCol).sort(spec);

          sh.deleteColumn(skCol);
        } else {
          const spec = [{ column: colCat1, ascending: true }];
          if (colBoq > 0) spec.push({ column: colBoq, ascending: true });
          sh.getRange(r1, 1, rN - r1 + 1, lastCol).sort(spec);
        }
      }

      // Center A-C and vertically middle
      if (rN >= r1) {
        sh
          .getRange(r1, 1, rN - r1 + 1, 3)
          .setHorizontalAlignment("center")
          .setVerticalAlignment("middle");
      }

      // merge BOQ BEFORE category1
      if (colBoq > 0 && colCat1 > 0 && rN >= r1) {
        mergeRunsInColumnWithinGroup_(sh, colBoq, colCat1, r1, rN, "center");
      }

      // merge category1 bands (column 1 = Product)
      if (rN >= r1) {
        mergeCategoryBands_(sh, 1, r1, rN);
      }

      removeColumnsByHeader_(sh, ["entity_type", "category"]);

      // merge Preview blanks ONLY within same BOQ name
      const cBoq2 = colIndexByHeader_(sh, "BOQ name");
      const cPrev2 = colIndexByHeader_(sh, "Preview");
      const rN2 = sh.getLastRow();
      if (cBoq2 > 0 && cPrev2 > 0 && rN2 >= r1) {
        mergePreviewWithinBoq_(sh, cPrev2, cBoq2, r1, rN2);
      }
    }

    return ContentService.createTextOutput(JSON.stringify({ ok: true, wrote: rows.length, tab, colorOnly }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log("Error in doPost: " + err);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(
      ContentService.MimeType.JSON
    );
  }
}
