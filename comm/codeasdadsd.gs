// ====== GLOBAL CONFIG ======
const DEFAULT_SHEET_ID = "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM";

/**
 * ✅ Quick deploy verification
 * Open your /exec URL in browser (GET) and you should see this text.
 */
function doGet() {
  return ContentService.createTextOutput(
    "DXF Sheets WebApp ✅ single-doPost (detail: category1→boq→zone + aggregate zones; bylayer merge) " +
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

function applyMerge_(sheet, col, s, e, value, anchor) {
  const band = sheet.getRange(s, col, e - s + 1, 1);
  band.clearContent();
  const anchorRow =
    anchor === "first" ? s : anchor === "middle" ? Math.floor((s + e) / 2) : e;

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
  const r1 = 2,
    rN = values.length;

  headerNames.forEach((name) => {
    const idx = headers.indexOf(normalizeHeader_(name));
    if (idx < 0) return;

    const col = idx + 1;
    sheet.getRange(r1, col, rN - r1 + 1, 1).breakApart();

    let s = r1;
    let vn = "";
    let vRaw = "";

    for (let r = r1; r <= rN + 1; r++) {
      const raw = r <= rN ? String(values[r - 1][idx] || "") : "\u0000__END__";
      const norm = normalizeHeader_(raw);

      if (!vn) {
        if (norm) {
          s = r;
          vn = norm;
          vRaw = raw;
        }
        continue;
      }

      const cont = r <= rN && (!norm || norm === vn);
      if (cont) continue;

      const e = r - 1;
      if (e > s) applyMerge_(sheet, col, s, e, vRaw, anchor);

      vn = "";
      vRaw = "";
      if (r <= rN && norm) {
        s = r;
        vn = norm;
        vRaw = raw;
      }
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
  const r1 = 2,
    rN = values.length;
  sheet.getRange(r1, tCol, rN - r1 + 1, 1).breakApart();

  let runStart = r1;
  let lastGroup = String(values[r1 - 1][gIdx] || "")
    .trim()
    .toUpperCase();
  let lastVal = String(values[r1 - 1][tIdx] || "")
    .trim()
    .toUpperCase();

  for (let r = r1 + 1; r <= rN + 1; r++) {
    const cg =
      r <= rN
        ? String(values[r - 1][gIdx] || "")
            .trim()
            .toUpperCase()
        : "\u0000__END__";
    const cv =
      r <= rN
        ? String(values[r - 1][tIdx] || "")
            .trim()
            .toUpperCase()
        : "\u0000__END__";

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
    // IMPORTANT: new band only starts when raw is non-blank
    current = raw;
  }
}

/**
 * Merge runs in target column, but ONLY within the same group column value.
 */
function mergeRunsInColumnWithinGroup_(sheet, targetCol, groupCol, r1, rN) {
  if (rN < r1) return;

  sheet.getRange(r1, targetCol, rN - r1 + 1, 1).breakApart();

  const tVals = sheet
    .getRange(r1, targetCol, rN - r1 + 1, 1)
    .getDisplayValues()
    .map((r) => String(r[0] || "").trim());
  const gVals = sheet
    .getRange(r1, groupCol, rN - r1 + 1, 1)
    .getDisplayValues()
    .map((r) => String(r[0] || "").trim());

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
      band.merge().setVerticalAlignment("middle").setHorizontalAlignment("left");
    }

    if (i < tVals.length) {
      runStart = r1 + i;
      lastKey = key;
    }
  }
}

// ====== DETAIL TRANSFORM (your requirement) ======
/**
 * Make Detail sheet look like your Image1 style:
 * 1) Column order: category1, BOQ name, zone, qty_type, qty_value, length (ft), width (ft), Description, Preview, remarks (others kept at end)
 * 2) Aggregate by (category1 + BOQ name): combine zones into one cell and SUM qty_value.
 *    Zone cell format:
 *      Conference (2)
 *      LoungeRoom (1)
 */
function transformDetail_(headersIn, rowsIn, imagesIn) {
  const headers = headersIn.slice();
  const rows = rowsIn.slice();
  const images = imagesIn ? imagesIn.slice() : [];

  // drop entity_type, category if present
  const kill = new Set(["entity_type", "category"]);
  const keepIdx = headers
    .map((h, i) => ({ i, keep: !kill.has(String(h).trim().toLowerCase()) }))
    .filter((x) => x.keep)
    .map((x) => x.i);

  const headers2 = keepIdx.map((i) => headers[i]);
  const rows2 = rows.map((r) => keepIdx.map((i) => r[i]));
  const images2 = images.length ? images.slice(0, rows.length).filter((_, i) => true) : [];

  // index map
  const hnorm = headers2.map((h) => normalizeHeader_(h));
  const idx = (name) => hnorm.indexOf(normalizeHeader_(name));

  const iCat1 = idx("category1");
  const iBoq  = idx("boq name");
  const iZone = idx("zone");
  const iQtyT = idx("qty_type");
  const iQtyV = idx("qty_value");
  const iLen  = idx("length (ft)");
  const iWid  = idx("width (ft)");
  const iDesc = idx("description");
  const iRem  = idx("remarks");
  const iPrev = idx("preview"); // may be missing in incoming rows; preview handled later

  // desired lead order (only those that exist)
  const desired = [
    "category1",
    "BOQ name",
    "zone",
    "qty_type",
    "qty_value",
    "length (ft)",
    "width (ft)",
    "Description",
    "Preview",
    "remarks",
  ];
  const desiredIdx = [];
  desired.forEach((n) => {
    const j = idx(n);
    if (j >= 0) desiredIdx.push(j);
  });

  // keep the rest after desired
  const restIdx = [];
  for (let j = 0; j < headers2.length; j++) {
    if (!desiredIdx.includes(j)) restIdx.push(j);
  }
  const order = desiredIdx.concat(restIdx);

  const headers3 = order.map((j) => headers2[j]);
  const rows3 = rows2.map((r) => order.map((j) => r[j]));

  // re-find indices after reorder
  const hnorm3 = headers3.map((h) => normalizeHeader_(h));
  const jCat1 = hnorm3.indexOf("category1");
  const jBoq  = hnorm3.indexOf("boq name");
  const jZone = hnorm3.indexOf("zone");
  const jQtyT = hnorm3.indexOf("qty_type");
  const jQtyV = hnorm3.indexOf("qty_value");
  const jLen  = hnorm3.indexOf("length (ft)");
  const jWid  = hnorm3.indexOf("width (ft)");
  const jDesc = hnorm3.indexOf("description");
  const jRem  = hnorm3.indexOf("remarks");

  // aggregate by cat1+boq
  if (jCat1 < 0 || jBoq < 0) {
    // cannot aggregate safely; still return reordered
    return { headers: headers3, rows: rows3, images: images2 };
  }

  const groups = new Map(); // key -> group obj
  for (let i = 0; i < rows3.length; i++) {
    const r = rows3[i];
    const cat1 = String(r[jCat1] || "").trim();
    const boq = String(r[jBoq] || "").trim();
    if (!cat1 && !boq) continue;
    const key = (cat1 || "") + "||" + (boq || "");

    const zone = jZone >= 0 ? String(r[jZone] || "").trim() : "";
    const qtyv = jQtyV >= 0 ? parseFloat(String(r[jQtyV] || "").trim()) : 0;
    const qty = isFinite(qtyv) ? qtyv : 0;

    if (!groups.has(key)) {
      groups.set(key, {
        cat1,
        boq,
        qty_type: jQtyT >= 0 ? r[jQtyT] : "count",
        qty_total: 0,
        zones: new Map(), // zone -> count
        len: jLen >= 0 ? r[jLen] : "",
        wid: jWid >= 0 ? r[jWid] : "",
        desc: jDesc >= 0 ? r[jDesc] : "",
        rem: jRem >= 0 ? r[jRem] : "",
        baseRow: r.slice(),
        img: images2[i] || "",
      });
    }
    const g = groups.get(key);

    g.qty_total += qty > 0 ? qty : 0;

    if (zone) {
      g.zones.set(zone, (g.zones.get(zone) || 0) + (qty > 0 ? qty : 1));
    }

    // Prefer first non-empty for these fields
    if (jLen >= 0 && !g.len && r[jLen]) g.len = r[jLen];
    if (jWid >= 0 && !g.wid && r[jWid]) g.wid = r[jWid];
    if (jDesc >= 0 && !g.desc && r[jDesc]) g.desc = r[jDesc];
    if (jRem >= 0 && !g.rem && r[jRem]) g.rem = r[jRem];
    if (!g.img && images2[i]) g.img = images2[i];
  }

  // build aggregated rows
  const outRows = [];
  const outImgs = [];
  const keys = Array.from(groups.keys()).sort((a, b) => a.localeCompare(b));
  keys.forEach((key) => {
    const g = groups.get(key);
    const row = g.baseRow.slice();

    // zone cell: list zones with counts
    if (jZone >= 0) {
      const zlist = Array.from(g.zones.entries())
        .sort((a, b) => a[0].localeCompare(b[0]))
        .map(([z, c]) => `${z} (${Number(c)})`);
      row[jZone] = zlist.join("\n");
    }

    // qty_value total
    if (jQtyV >= 0) row[jQtyV] = g.qty_total ? Number(g.qty_total) : "";

    // overwrite preferred fields
    if (jQtyT >= 0) row[jQtyT] = g.qty_type || row[jQtyT];
    if (jLen >= 0) row[jLen] = g.len || row[jLen];
    if (jWid >= 0) row[jWid] = g.wid || row[jWid];
    if (jDesc >= 0) row[jDesc] = g.desc || row[jDesc];
    if (jRem >= 0) row[jRem] = g.rem || row[jRem];

    outRows.push(row);
    outImgs.push(g.img || "");
  });

  return { headers: headers3, rows: outRows, images: outImgs };
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
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (_) {
      try {
        file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (_) {}
    }
    const url = "https://drive.google.com/uc?export=view&id=" + file.getId();

    rows.forEach((r) => {
      sh.getRange(r, cPrev)
        .setFormula('=IMAGE("' + url + '")')
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle");
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
    const colorOnly = parseBool_(p.colorOnly); // ✅
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

    // DETAIL transform (this is your requested fix)
    if (!colorOnly) {
      // if headers not provided (append batches), use current sheet headers
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
          try {
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          } catch (_) {
            try {
              file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
            } catch (_) {}
          }
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
      try {
        const f = sh.getFilter();
        if (f) f.remove();
      } catch (_) {}

      const colCategory = colIndexByHeader_(sh, "category");
      const colZone = colIndexByHeader_(sh, "zone");
      const r1 = 2,
        rN = sh.getLastRow();

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
      // DETAIL: after transform we want Category1 bands (col A) like your Image1,
      // and Zone should wrap because it can have multiple lines now.
      const colCat1 = colIndexByHeader_(sh, "category1") || 1;
      const colZone = colIndexByHeader_(sh, "zone");
      const r1 = 2,
        rN = sh.getLastRow();

      // Sort by category1 then BOQ name
      const colBoq = colIndexByHeader_(sh, "BOQ name");
      if (rN >= r1 && colCat1 > 0) {
        const lastCol = sh.getLastColumn();
        const spec = [{ column: colCat1, ascending: true }];
        if (colBoq > 0) spec.push({ column: colBoq, ascending: true });
        sh.getRange(r1, 1, rN - r1 + 1, lastCol).sort(spec);
      }

      // Wrap zone text + align left for readability
      if (colZone > 0 && rN >= r1) {
        const zr = sh.getRange(r1, colZone, rN - r1 + 1, 1);
        zr.setWrap(true);
        zr.setHorizontalAlignment("left");
        zr.setVerticalAlignment("middle");

        // increase row height if multiple zones
        const zVals = zr.getDisplayValues().map((a) => String(a[0] || ""));
        for (let i = 0; i < zVals.length; i++) {
          const lines = Math.max(1, zVals[i].split("\n").length);
          const h = Math.max(50, 18 * lines + 14);
          sh.setRowHeight(r1 + i, h);
        }
      }

      // Clean old columns if any
      removeColumnsByHeader_(sh, ["entity_type", "category"]);
    }

    // ✅ ALWAYS merge Column A into vertical bands (your green highlight)
    try {
      const f = sh.getFilter();
      if (f) f.remove();
    } catch (_) {}
    const r1A = 2,
      rNA = sh.getLastRow();
    if (rNA >= r1A) mergeCategoryBands_(sh, 1, r1A, rNA);

    return ContentService.createTextOutput(
      JSON.stringify({
        ok: true,
        wrote: rows.length,
        tab,
        colorOnly,
      })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log("Error in doPost: " + err);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(err) })).setMimeType(
      ContentService.MimeType.JSON
    );
  }
}
