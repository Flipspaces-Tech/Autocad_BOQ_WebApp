/**************************************************
 * BOQ-LAYER → MASTER
 *
 * MEASUREMENT: from LAYER tab (optional per master tab)
 * QTY: from vis_export_sheet_like tab (required)
 *
 * ✅ Fills QTY even if MEASUREMENT column is missing in that master tab
 * ✅ Strong normalization for name matching (punctuation/dashes/spaces/newlines)
 * ✅ After run: hides rows based on final rule:
 *    - If MEAS col exists: hide row ONLY when (MEAS is blank/0) AND (QTY is blank/0)
 *    - If MEAS col missing: hide row when (QTY is blank/0)
 * ✅ Avoids un-hiding everything; only unhides rows that should be visible
 **************************************************/

const SHEETS = SpreadsheetApp;

const BOQ_SYNC = {
  // Mapping sheet (script hosted here)
  MAP_TAB: "BOQ-LAYER",

  // Mapping columns (1-based)
  MAP_COL_TARGETED: 2,      // Targeted BOQ Name
  MAP_COL_GENERATED: 5,     // Generated Layer Name (often in E)
  MAP_COL_BLOCKNAME: 6,     // Generated-Block Block Name (green column)

  // Export spreadsheet
  EXPORT_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",

  // Measurement source
  EXPORT_MEASURE_TAB: "LAYER",
  LAYER_HDR_LAYER: "layer",
  LAYER_HDR_ZONE: "zone",
  LAYER_HDR_AREA: "area (ft2)",
  LAYER_HDR_PERIM: "perimeter",
  LAYER_HDR_LENGTH: "length (ft)",

  // QTY source
  EXPORT_QTY_TAB: "vis_export_sheet_like",
  QTY_HDR_BOQNAME_PATTERNS: ["BOQ name", "boq_name", "name"],
  QTY_HDR_ZONE_PATTERNS: ["zone", "location"],
  QTY_HDR_QTYVALUE_PATTERNS: ["qty_value", "qty value", "qty", "quantity"],

  // Master
  MASTER_SS_ID: "1CVibwjRFz4gTATAeOFUlYzlGZybXILO60OrxwGaFLeY",
  MASTER_START_TAB: "Civil",

  MASTER_MATCH_COL_FALLBACK: 2,  // scope of work fallback col B
  MASTER_QTY_COL_FALLBACK: 9,    // QTY default column I

  MASTER_HDR_SCOPE: ["scope of work"],
  MASTER_HDR_LOCATION: ["location"],
  MASTER_HDR_MEASUREMENT: ["measurement", "qty measured", "measured"], // broaden
  MASTER_HDR_QTY: ["qty", "quantity"],
  MASTER_HDR_SRNO: ["sr. no.", "sr no", "sr.no", "sr"],

  NUMBER_FORMAT: "0.############",
  SHOW_DIALOG: true,
  DIALOG_TITLE: "Vizdom Sync — BOQ-LAYER → MASTER",

  // Skip write row only if BOTH meas+qty are 0 (if meas col missing => decided by qty only)
  SKIP_WHEN_BOTH_ZERO: true,

  // Hide behavior after sync
  APPLY_ROW_HIDING: true,
  HIDE_START_ROW: 2, // start applying hiding from row 2 (keep header safe)
};

function onOpen() {
  SHEETS.getUi()
    .createMenu("Vizdom Sync")
    .addItem("Sync BOQ-LAYER → MASTER", "syncBoqLayerToMaster")
    .addToUi();
}

function syncBoqLayerToMaster() {
  const ui = SHEETS.getUi();
  const startedAt = new Date();

  const report = {
    startedAt,
    finishedAt: null,
    mappingsTotal: 0,
    tabsProcessed: 0,
    tabsSkippedMissingCols: [],
    targetsMatched: 0,
    rowsInserted: 0,
    rowsUpdated: 0,
    skippedBothZero: 0,
    notFoundTargets: 0,

    rowsHidden: 0,
    rowsUnhidden: 0,

    debug: {
      masterDetectedCols: [],
      qtyDetectedCols: [],
      qtyRowsScanned: 0,
      qtyUniqueBoqKeys: 0,
      qtyMatchesFound: 0,
      sampleMappingBlockKeys: [],
      sampleQtyBoqKeys: [],
    },

    errors: [],
    notes: [],
  };

  try {
    const mapSS = SHEETS.getActiveSpreadsheet();
    const mapSh = mapSS.getSheetByName(BOQ_SYNC.MAP_TAB);
    if (!mapSh) throw new Error(`Mapping tab not found: ${BOQ_SYNC.MAP_TAB}`);

    const exportSS = SHEETS.openById(BOQ_SYNC.EXPORT_SS_ID);
    const masterSS = SHEETS.openById(BOQ_SYNC.MASTER_SS_ID);

    // 1) Read mappings
    const mappings = readMappings_(mapSh, report);
    report.mappingsTotal = mappings.length;

    if (!mappings.length) {
      report.notes.push("No mappings found in BOQ-LAYER.");
      finalize_(report);
      if (BOQ_SYNC.SHOW_DIALOG) showReportDialog_(ui, report);
      return;
    }

    // 2) Build lookup by targeted BOQ name
    const mappingByTarget = new Map();
    for (const m of mappings) {
      const tKey = normKey_advanced_(m.targeted);
      if (!mappingByTarget.has(tKey)) mappingByTarget.set(tKey, []);
      mappingByTarget.get(tKey).push(m);
    }

    // 3) Build QTY index once
    const qtyIndex = buildQtyIndex_(exportSS, report);

    // 4) Measurement cache
    const measCache = new Map();

    // 5) Iterate master tabs
    const tabs = getMasterTabsFromStart_(masterSS, BOQ_SYNC.MASTER_START_TAB);
    if (!tabs.length) throw new Error(`Start tab "${BOQ_SYNC.MASTER_START_TAB}" not found in MASTER.`);

    const targetsFoundSomewhere = new Set();

    for (const sh of tabs) {
      const tabName = sh.getName();
      const lastCol = sh.getLastColumn();
      let lastRow = sh.getLastRow();
      if (lastRow < 2 || lastCol < 2) continue;

      report.tabsProcessed++;

      const matchCol =
        findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_SCOPE, 15) || BOQ_SYNC.MASTER_MATCH_COL_FALLBACK;

      const locationCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_LOCATION, 15);
      const measurementCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_MEASUREMENT, 15); // may be null
      const qtyColDetected = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_QTY, 15);
      const qtyCol = qtyColDetected || BOQ_SYNC.MASTER_QTY_COL_FALLBACK;
      const srNoCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_SRNO, 15);

      report.debug.masterDetectedCols.push(
        `${tabName}: matchCol=${matchCol}, locationCol=${locationCol}, measurementCol=${measurementCol}, qtyCol=${qtyCol} (detected=${qtyColDetected || "no"})`
      );

      // Require LOCATION + QTY col
      if (!locationCol || !qtyCol) {
        report.tabsSkippedMissingCols.push(`${tabName} (missing LOCATION or QTY column)`);
        continue;
      }

      const scopeVals = sh.getRange(1, matchCol, lastRow, 1).getDisplayValues();

      let r = 1;
      while (r <= scopeVals.length) {
        const scopeText = String((scopeVals[r - 1] && scopeVals[r - 1][0]) || "").trim();
        if (!scopeText) { r++; continue; }

        // handle merged
        const cell = sh.getRange(r, matchCol);
        if (cell.isPartOfMerge()) {
          const mr = cell.getMergedRanges()[0];
          if (!(mr.getRow() === r && mr.getColumn() === matchCol)) { r++; continue; }
        }

        const rowKey = normKey_advanced_(scopeText);
        const mapList = mappingByTarget.get(rowKey);
        if (!mapList || !mapList.length) { r++; continue; }

        report.targetsMatched++;
        targetsFoundSomewhere.add(rowKey);

        // ---- zones union from MEAS + QTY
        const zoneTotalsMeas = new Map();
        const zoneTotalsQty = new Map();
        const zoneOrder = [];

        // MEAS (optional)
        if (measurementCol) {
          for (const m of mapList) {
            const cacheKey = `${normKey_advanced_(m.generated)}||${normKey_advanced_(m.measure)}`;
            let breakdown = measCache.get(cacheKey);
            if (!breakdown) {
              breakdown = computeMeasurementFromLayer_(exportSS, m.generated, m.measure);
              measCache.set(cacheKey, breakdown);
            }
            for (const z of breakdown.order) pushZone_(zoneOrder, z);
            for (const [z, v] of breakdown.byZone.entries()) {
              const n = toNumber_(v);
              if (!Number.isFinite(n)) continue;
              zoneTotalsMeas.set(z, (zoneTotalsMeas.get(z) || 0) + n);
            }
          }
        }

        // QTY (from index)
        for (const m of mapList) {
          if (!m.blockName) continue;
          const bKey = normKey_advanced_(m.blockName);

          const rec = qtyIndex.get(bKey);
          if (!rec) continue;

          report.debug.qtyMatchesFound++;

          for (const z of rec.order) pushZone_(zoneOrder, z);
          for (const [z, v] of rec.byZone.entries()) {
            const n = toNumber_(v);
            if (!Number.isFinite(n)) continue;
            zoneTotalsQty.set(z, (zoneTotalsQty.get(z) || 0) + n);
          }
        }

        if (!zoneOrder.length) {
          r++;
          continue;
        }

        // Filter zones where both are 0 (or if meas col missing => decided by qty only)
        const finalZones = [];
        for (const z of zoneOrder) {
          const meas = toNumber_(zoneTotalsMeas.get(z) || 0);
          const qty = toNumber_(zoneTotalsQty.get(z) || 0);

          const measZero = !measurementCol || (!Number.isFinite(meas) || meas === 0); // if no meas col => treat as 0
          const qtyZero = !Number.isFinite(qty) || qty === 0;

          let skip = false;
          if (BOQ_SYNC.SKIP_WHEN_BOTH_ZERO) {
            if (!measurementCol) {
              // no measurement col => skip if qty is zero
              skip = qtyZero;
            } else {
              skip = measZero && qtyZero;
            }
          }

          if (skip) {
            report.skippedBothZero++;
            continue;
          }

          finalZones.push(z);
        }

        if (!finalZones.length) { r++; continue; }

        const needed = finalZones.length;

        // Insert rows if needed
        if (needed > 1) {
          sh.insertRowsAfter(r, needed - 1);
          report.rowsInserted += (needed - 1);

          const blanks = Array.from({ length: needed - 1 }, () => [""]);
          scopeVals.splice(r, 0, ...blanks);

          lastRow += (needed - 1);

          const baseRowRange = sh.getRange(r, 1, 1, lastCol);
          for (let i = 1; i < needed; i++) {
            const newRowRange = sh.getRange(r + i, 1, 1, lastCol);
            baseRowRange.copyTo(newRowRange, { contentsOnly: false });

            const clearUpto = Math.max(1, locationCol - 1);
            sh.getRange(r + i, 1, 1, clearUpto).clearContent();
          }

          safeMergeAndCenter_(sh, r, matchCol, needed);
          if (srNoCol) safeMergeAndCenter_(sh, r, srNoCol, needed);
        }

        // Write rows
        for (let i = 0; i < needed; i++) {
          const zone = finalZones[i];

          sh.getRange(r + i, locationCol).setValue(zone);

          if (measurementCol) {
            const meas = toNumber_(zoneTotalsMeas.get(zone) || 0);
            const mCell = sh.getRange(r + i, measurementCol);
            mCell.setValue(Number.isFinite(meas) ? meas : 0);
            mCell.setNumberFormat(BOQ_SYNC.NUMBER_FORMAT);
          }

          const qty = toNumber_(zoneTotalsQty.get(zone) || 0);
          const qCell = sh.getRange(r + i, qtyCol);
          qCell.setValue(Number.isFinite(qty) ? qty : 0);
          qCell.setNumberFormat(BOQ_SYNC.NUMBER_FORMAT);
        }

        report.rowsUpdated += needed;
        r += needed;
      }

      // ✅ AFTER SHEET UPDATE: Apply hiding rule for that sheet
      if (BOQ_SYNC.APPLY_ROW_HIDING) {
        const res = applyRowHiding_(sh, measurementCol, qtyCol);
        report.rowsHidden += res.hidden;
        report.rowsUnhidden += res.unhidden;
      }
    }

    // Not found targets
    let notFound = 0;
    for (const m of mappings) {
      const tKey = normKey_advanced_(m.targeted);
      if (!targetsFoundSomewhere.has(tKey)) notFound++;
    }
    report.notFoundTargets = notFound;

    finalize_(report);
    if (BOQ_SYNC.SHOW_DIALOG) showReportDialog_(ui, report);
  } catch (e) {
    report.errors.push(String(e && e.stack ? e.stack : e));
    finalize_(report);
    if (BOQ_SYNC.SHOW_DIALOG) showReportDialog_(SHEETS.getUi(), report, true);
    throw e;
  }
}

/* -------------------------
   Read mappings
------------------------- */
function readMappings_(mapSh, report) {
  const lastRow = mapSh.getLastRow();
  const lastCol = mapSh.getLastColumn();
  if (lastRow < 2) return [];

  const grid = mapSh.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  const out = [];
  for (let r = 1; r < grid.length; r++) {
    const targeted = String(grid[r][BOQ_SYNC.MAP_COL_TARGETED - 1] || "").trim();
    const generated = String(grid[r][BOQ_SYNC.MAP_COL_GENERATED - 1] || "").trim();
    const blockName = String(grid[r][BOQ_SYNC.MAP_COL_BLOCKNAME - 1] || "").trim();

    if (!targeted || !generated) continue;

    out.push({
      targeted,
      generated,
      blockName,
      measure: "Area",
    });
  }

  for (const m of out.slice(0, 8)) {
    report.debug.sampleMappingBlockKeys.push(`${m.blockName}  =>  ${normKey_advanced_(m.blockName)}`);
  }

  return out;
}

/* -------------------------
   Build QTY index
------------------------- */
function buildQtyIndex_(exportSS, report) {
  const sh = exportSS.getSheetByName(BOQ_SYNC.EXPORT_QTY_TAB);
  if (!sh) throw new Error(`QTY tab not found: ${BOQ_SYNC.EXPORT_QTY_TAB}`);

  const dr = sh.getDataRange().getValues();
  if (dr.length < 2) throw new Error(`QTY tab "${BOQ_SYNC.EXPORT_QTY_TAB}" has no data`);

  const header = dr[0].map(v => String(v || "").trim());

  const idxBoq = findHeaderIndexContains_(header, BOQ_SYNC.QTY_HDR_BOQNAME_PATTERNS);
  const idxZone = findHeaderIndexContains_(header, BOQ_SYNC.QTY_HDR_ZONE_PATTERNS);
  const idxQty = findHeaderIndexContains_(header, BOQ_SYNC.QTY_HDR_QTYVALUE_PATTERNS);

  report.debug.qtyDetectedCols.push(
    `QTY tab "${BOQ_SYNC.EXPORT_QTY_TAB}": idxBoq=${idxBoq}, idxZone=${idxZone}, idxQty=${idxQty} (0-based)`
  );

  if (idxBoq === -1 || idxZone === -1 || idxQty === -1) {
    throw new Error(
      `QTY tab header detect failed. Found idxBoq=${idxBoq}, idxZone=${idxZone}, idxQty=${idxQty}. Check row 1 headers.`
    );
  }

  report.debug.qtyRowsScanned = dr.length - 1;

  const index = new Map();
  let currentBoq = "";

  for (let r = 1; r < dr.length; r++) {
    const boqCell = String(dr[r][idxBoq] || "").trim();
    if (boqCell) currentBoq = boqCell;
    if (!currentBoq) continue;

    const key = normKey_advanced_(currentBoq);
    const zone = String(dr[r][idxZone] || "").trim() || "misc";
    const qty = toNumber_(dr[r][idxQty]);
    if (!Number.isFinite(qty)) continue;

    if (!index.has(key)) index.set(key, { order: [], byZone: new Map() });
    const rec = index.get(key);

    if (!rec.byZone.has(zone)) rec.order.push(zone);
    rec.byZone.set(zone, (rec.byZone.get(zone) || 0) + qty);
  }

  report.debug.qtyUniqueBoqKeys = index.size;

  let i = 0;
  for (const [k] of index.entries()) {
    report.debug.sampleQtyBoqKeys.push(k);
    i++;
    if (i >= 12) break;
  }

  return index;
}

function findHeaderIndexContains_(headerRow, patterns) {
  const pats = patterns.map(p => normKey_advanced_(p));
  for (let i = 0; i < headerRow.length; i++) {
    const h = normKey_advanced_(headerRow[i]);
    if (!h) continue;
    if (pats.some(p => h.includes(p))) return i;
  }
  return -1;
}

/* -------------------------
   Measurement from LAYER
------------------------- */
function computeMeasurementFromLayer_(exportSS, layerName, measure) {
  const sh = exportSS.getSheetByName(BOQ_SYNC.EXPORT_MEASURE_TAB);
  if (!sh) throw new Error(`Export MEAS tab not found: ${BOQ_SYNC.EXPORT_MEASURE_TAB}`);

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { order: [], byZone: new Map() };

  const header = values[0].map(v => String(v || "").trim().toLowerCase());

  const idxLayer = header.indexOf(BOQ_SYNC.LAYER_HDR_LAYER.toLowerCase());
  const idxZone = header.indexOf(BOQ_SYNC.LAYER_HDR_ZONE.toLowerCase());
  const idxArea = header.indexOf(BOQ_SYNC.LAYER_HDR_AREA.toLowerCase());
  const idxPerim = header.indexOf(BOQ_SYNC.LAYER_HDR_PERIM.toLowerCase());
  const idxLen = header.indexOf(BOQ_SYNC.LAYER_HDR_LENGTH.toLowerCase());

  if (idxLayer === -1) throw new Error(`LAYER tab missing header: ${BOQ_SYNC.LAYER_HDR_LAYER}`);
  if (idxZone === -1) throw new Error(`LAYER tab missing header: ${BOQ_SYNC.LAYER_HDR_ZONE}`);

  const m = normKey_advanced_(measure);
  let idxMeasure = idxArea;
  let isCount = false;

  if (m === "perimeter") idxMeasure = idxPerim;
  else if (m === "length") idxMeasure = idxLen;
  else if (m === "count") isCount = true;

  const targetKey = normKey_advanced_(layerName);

  const order = [];
  const seen = new Set();
  const byZone = new Map();

  let currentLayer = "";

  for (let r = 1; r < values.length; r++) {
    const layerCell = String(values[r][idxLayer] || "").trim();
    if (layerCell) currentLayer = layerCell;
    if (!currentLayer) continue;

    if (normKey_advanced_(currentLayer) !== targetKey) continue;

    const zone = String(values[r][idxZone] || "").trim() || "misc";

    if (!seen.has(zone)) {
      seen.add(zone);
      order.push(zone);
    }

    let add = 0;
    if (isCount) add = 1;
    else {
      const v = toNumber_(values[r][idxMeasure]);
      if (!Number.isFinite(v)) continue;
      add = v;
    }

    byZone.set(zone, (byZone.get(zone) || 0) + add);
  }

  return { order, byZone };
}

/* -------------------------
   Hiding logic (FINAL RULE)
------------------------- */
function applyRowHiding_(sh, measurementCol, qtyCol) {
  const lastRow = sh.getLastRow();
  const startRow = BOQ_SYNC.HIDE_START_ROW || 2;
  if (lastRow < startRow) return { hidden: 0, unhidden: 0 };

  const numRows = lastRow - startRow + 1;

  // Display values are best for blanks coming from formulas
  const qtyDisp = sh.getRange(startRow, qtyCol, numRows, 1).getDisplayValues();
  let measDisp = null;
  if (measurementCol) measDisp = sh.getRange(startRow, measurementCol, numRows, 1).getDisplayValues();

  const isZeroOrBlank = (x) => {
    const t = String(x ?? "").trim();
    if (t === "" || t === "-") return true;
    const n = Number(t.replace(/,/g, ""));
    return Number.isFinite(n) && n === 0;
  };

  const rowsToHide = [];
  const rowsToUnhide = [];

  for (let i = 0; i < numRows; i++) {
    const rowNum = startRow + i;

    const qtyZeroBlank = isZeroOrBlank(qtyDisp[i][0]);

    let shouldHide = false;
    if (!measurementCol) {
      // No measurement column -> deciding factor = QTY
      shouldHide = qtyZeroBlank;
    } else {
      // Has measurement column -> hide only when BOTH are blank/zero
      const measZeroBlank = isZeroOrBlank(measDisp[i][0]);
      shouldHide = qtyZeroBlank && measZeroBlank;
    }

    const currentlyHidden = sh.isRowHiddenByUser(rowNum);

    if (shouldHide) rowsToHide.push(rowNum);
    else if (currentlyHidden) rowsToUnhide.push(rowNum);
  }

  const hidden = setRowsHidden_(sh, rowsToHide, true);
  const unhidden = setRowsHidden_(sh, rowsToUnhide, false);

  SpreadsheetApp.flush();
  return { hidden, unhidden };
}

function setRowsHidden_(sh, rows, hide) {
  if (!rows || !rows.length) return 0;

  rows.sort((a, b) => a - b);

  let count = 0;
  let start = rows[0];
  let prev = rows[0];

  const flush = (s, e) => {
    const n = e - s + 1;
    if (hide) sh.hideRows(s, n);
    else sh.showRows(s, n);
    count += n;
  };

  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r === prev + 1) {
      prev = r;
      continue;
    }
    flush(start, prev);
    start = r;
    prev = r;
  }
  flush(start, prev);

  return count;
}

/* -------------------------
   Helpers
------------------------- */
function getMasterTabsFromStart_(masterSS, startName) {
  const sheets = masterSS.getSheets();
  const startKey = normKey_advanced_(startName);

  let startIdx = -1;
  for (let i = 0; i < sheets.length; i++) {
    if (normKey_advanced_(sheets[i].getName()) === startKey) {
      startIdx = i;
      break;
    }
  }
  if (startIdx === -1) return [];
  return sheets.slice(startIdx);
}

function findHeaderColContains_(sheet, candidates, scanRows) {
  const lastCol = sheet.getLastColumn();
  const rows = Math.min(scanRows || 15, sheet.getLastRow());
  if (rows < 1 || lastCol < 1) return null;

  const grid = sheet.getRange(1, 1, rows, lastCol).getDisplayValues();
  const wanted = candidates.map(c => normKey_advanced_(c));

  for (let r = 0; r < grid.length; r++) {
    for (let c = 0; c < grid[r].length; c++) {
      const cell = normKey_advanced_(grid[r][c]);
      if (!cell) continue;
      if (wanted.some(w => cell.includes(w))) return c + 1;
    }
  }
  return null;
}

// ✅ Avoid merge paste errors by never merging if range intersects existing merges
// ✅ Avoid merge paste errors by never merging if range intersects existing merges
function safeMergeAndCenter_(sh, startRow, col, numRows) {
  if (numRows <= 1) return;

  const rng = sh.getRange(startRow, col, numRows, 1);

  // If ANY merge on the sheet intersects this column+rows block, skip merging
  if (rangeIntersectsAnyMerge_(sh, startRow, col, numRows, 1)) {
    rng.setHorizontalAlignment("center");
    rng.setVerticalAlignment("middle");
    return;
  }

  rng.merge();
  rng.setHorizontalAlignment("center");
  rng.setVerticalAlignment("middle");
}

function rangeIntersectsAnyMerge_(sh, row, col, numRows, numCols) {
  const lastRow = Math.max(sh.getLastRow(), 1);
  const lastCol = Math.max(sh.getLastColumn(), 1);

  // ✅ getMergedRanges() works on Range, not Sheet
  const merges = sh.getRange(1, 1, lastRow, lastCol).getMergedRanges();
  if (!merges || !merges.length) return false;

  const r1 = row, r2 = row + numRows - 1;
  const c1 = col, c2 = col + numCols - 1;

  for (const mr of merges) {
    const mr1 = mr.getRow();
    const mr2 = mr.getRow() + mr.getNumRows() - 1;
    const mc1 = mr.getColumn();
    const mc2 = mr.getColumn() + mr.getNumColumns() - 1;

    const rowOverlap = !(r2 < mr1 || r1 > mr2);
    const colOverlap = !(c2 < mc1 || c1 > mc2);

    if (rowOverlap && colOverlap) return true;
  }

  return false;
}

function pushZone_(arr, zone) {
  const z = (zone && String(zone).trim()) ? String(zone).trim() : "misc";
  if (!arr.includes(z)) arr.push(z);
}

function finalize_(report) {
  report.finishedAt = new Date();
}

function showReportDialog_(ui, report, isError) {
  const durSec = Math.round((report.finishedAt - report.startedAt) / 1000);

  const lines = [];
  lines.push(`Started: ${report.startedAt.toLocaleString()}`);
  lines.push(`Finished: ${report.finishedAt.toLocaleString()}`);
  lines.push(`Duration: ${durSec}s`);
  lines.push("");
  lines.push(`Mappings loaded: ${report.mappingsTotal}`);
  lines.push(`Tabs processed: ${report.tabsProcessed}`);
  lines.push(`Targets matched: ${report.targetsMatched}`);
  lines.push(`Rows inserted: ${report.rowsInserted}`);
  lines.push(`Rows updated: ${report.rowsUpdated}`);
  lines.push(`Skipped zones (both 0): ${report.skippedBothZero}`);
  lines.push(`Targets not found: ${report.notFoundTargets}`);
  lines.push("");
  lines.push(`Rows hidden: ${report.rowsHidden}`);
  lines.push(`Rows unhidden: ${report.rowsUnhidden}`);
  lines.push("");

  lines.push("DEBUG — QTY sheet:");
  for (const s of report.debug.qtyDetectedCols) lines.push(` - ${s}`);
  lines.push(` - rows scanned: ${report.debug.qtyRowsScanned}`);
  lines.push(` - unique BOQ keys: ${report.debug.qtyUniqueBoqKeys}`);
  lines.push(` - qty matches found: ${report.debug.qtyMatchesFound}`);
  lines.push("");

  lines.push("DEBUG — sample mapping block keys (raw => normalized):");
  for (const s of report.debug.sampleMappingBlockKeys) lines.push(` - ${s}`);
  lines.push("");

  lines.push("DEBUG — sample QTY BOQ keys (normalized):");
  for (const s of report.debug.sampleQtyBoqKeys) lines.push(` - ${s}`);
  lines.push("");

  lines.push("DEBUG — Master detected columns:");
  for (const s of report.debug.masterDetectedCols) lines.push(` - ${s}`);

  if (report.tabsSkippedMissingCols.length) {
    lines.push("");
    lines.push("Skipped tabs:");
    for (const t of report.tabsSkippedMissingCols) lines.push(` - ${t}`);
  }

  if (report.errors.length) {
    lines.push("");
    lines.push("Errors:");
    for (const e of report.errors) {
      lines.push("--------------------------------------------------");
      lines.push(e);
    }
  }

  const title = isError ? `${BOQ_SYNC.DIALOG_TITLE} (ERROR)` : BOQ_SYNC.DIALOG_TITLE;
  ui.alert(title, lines.join("\n"), ui.ButtonSet.OK);
}

/**
 * Strong normalization:
 * - lowercase
 * - replace all non-alphanum with spaces
 * - normalize all dash types, newlines, NBSP
 * - collapse spaces
 */
function normKey_advanced_(s) {
  let t = String(s || "")
    .replace(/\u00A0/g, " ")
    .replace(/[\r\n]+/g, " ")
    .replace(/[–—−]/g, "-")
    .toLowerCase()
    .trim();

  t = t.replace(/[^a-z0-9]+/g, " ");
  t = t.replace(/\s+/g, " ").trim();
  return t;
}

function toNumber_(v) {
  if (v == null || v === "") return NaN;
  const n = Number(v);
  return Number.isFinite(n) ? n : NaN;
}