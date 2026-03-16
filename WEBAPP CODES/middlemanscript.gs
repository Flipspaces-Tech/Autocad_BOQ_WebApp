/**************************************************
 * BOQ-LAYER → MASTER (MEAS + QTY + HIDE ZEROS)
 *
 * MEASUREMENT: from EXPORT "LAYER" tab (optional per master tab)
 * QTY: from EXPORT "vis_export_sheet_like" tab (required)
 *
 * ✅ Writes QTY even if MEASUREMENT column is missing in a master tab
 * ✅ QTY indexing no longer depends on "BOQ name" merged/carry-forward cells
 *    - Uses Product as primary key (works even if BOQ name is blank)
 *    - Also indexes by BOQ name as secondary key (backward compatible)
 * ✅ Strong normalization for matching (punctuation/dashes/spaces/newlines)
 * ✅ Skips creating/writing zone rows where BOTH meas & qty are zero/blank (configurable)
 * ✅ Merges & centers "SCOPE OF WORK" (and SR NO if present) across inserted rows
 * ✅ After sync: hides rows where BOTH (meas AND qty) are blank/0
 *    - If measurement column is missing in that tab → deciding factor is only QTY
 * ✅ Detailed dialog after every run
 *
 * Fixes:
 * - No sh.getMergedRanges() usage (Sheets don’t have it). Only Range.getMergedRanges()
 * - Avoids “paste partially intersects a merge” by breaking merge safely before insert/copy
 **************************************************/

const SHEETS = SpreadsheetApp;

const BOQ_SYNC = {
  // Mapping sheet (script hosted here)
  MAP_TAB: "BOQ-LAYER",

  // Mapping columns (1-based) — adjust if your sheet differs
  MAP_COL_TARGETED: 2,   // Targeted BOQ Name (master match)
  MAP_COL_GENERATED: 4,  // Generated Layer Name (for LAYER measurement)
  MAP_COL_BLOCKNAME: 6,  // Generated-Block / Block Name (for QTY matching)

  // Export spreadsheet (Auto-QA Output)
  EXPORT_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",

  // Measurement source tab
  EXPORT_MEASURE_TAB: "LAYER",
  LAYER_HDR_LAYER: "layer",
  LAYER_HDR_ZONE: "zone",
  LAYER_HDR_AREA: "area (ft2)",
  LAYER_HDR_PERIM: "perimeter",
  LAYER_HDR_LENGTH: "length (ft)",

  // QTY source tab
  EXPORT_QTY_TAB: "vis_export_sheet_like",
  QTY_HDR_PRODUCT_PATTERNS: ["product"],
  QTY_HDR_BOQNAME_PATTERNS: ["boq name", "boq_name", "name"],
  QTY_HDR_ZONE_PATTERNS: ["zone", "location"],
  QTY_HDR_QTYVALUE_PATTERNS: ["qty_value", "qty value", "qty", "quantity"],

  // Master spreadsheet
  MASTER_SS_ID: "12sJ3s0W8QkLXAwUKEJhPD-ydmQPCCHsv7sau3cOQszY",//1CVibwjRFz4gTATAeOFUlYzlGZybXILO60OrxwGaFLeY
  MASTER_START_TAB: "Civil",

  // Header detection in master
  MASTER_MATCH_COL_FALLBACK: 2, // usually col B
  MASTER_QTY_COL_FALLBACK: 9,   // usually col I (QTY)

  MASTER_HDR_SCOPE: ["scope of work"],
  MASTER_HDR_LOCATION: ["location"],
  // Measurement column might be named differently across tabs
  MASTER_HDR_MEASUREMENT: ["measurement", "qty measured", "measured"],
  MASTER_HDR_QTY: ["qty", "quantity"],
  MASTER_HDR_SRNO: ["sr. no.", "sr no", "sr.no", "sr"],

  // Rules
  // When building zone rows: skip if BOTH meas & qty are blank/0
  SKIP_WHEN_BOTH_ZERO: true,

  // After finishing a tab: hide if BOTH meas & qty are blank/0
  // If measurement column is missing in a tab → hide if qty is blank/0
  HIDE_ZERO_ROWS_AFTER_SYNC: true,

  NUMBER_FORMAT: "0.############",
  SHOW_DIALOG: true,
  DIALOG_TITLE: "Vizdom Sync — BOQ-LAYER → MASTER",
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
    tabsSkipped: [],

    targetsMatched: 0,
    rowsInserted: 0,
    rowsUpdated: 0,
    skippedBothZero: 0,

    rowsHidden: 0,
    rowsUnhidden: 0,

    notFoundTargets: 0,

    debug: {
      masterDetectedCols: [],
      qtyDetectedCols: [],
      qtyRowsScanned: 0,
      qtyUniqueBoqKeys: 0,
      qtyMatchesFound: 0,
      sampleMappingQtyKeys: [],
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

    // 2) Build mapping lookup by targeted (master scope of work)
    const mappingByTarget = new Map();
    for (const m of mappings) {
      const tKey = normKey_advanced_(m.targeted);
      if (!mappingByTarget.has(tKey)) mappingByTarget.set(tKey, []);
      mappingByTarget.get(tKey).push(m);
    }

    // 3) Build QTY index once (by normalized Product primary, BOQ name secondary)
    const qtyIndex = buildQtyIndex_(exportSS, report);

    // 4) Measurement cache (layerName+measure -> breakdown)
    const measCache = new Map();

    // 5) Iterate master tabs from start
    const tabs = getMasterTabsFromStart_(masterSS, BOQ_SYNC.MASTER_START_TAB);
    if (!tabs.length) throw new Error(`Start tab "${BOQ_SYNC.MASTER_START_TAB}" not found in MASTER.`);

    const targetsFoundSomewhere = new Set();

    for (const sh of tabs) {
      const tabName = sh.getName();
      const lastCol = sh.getLastColumn();
      let lastRow = sh.getLastRow();
      if (lastRow < 2 || lastCol < 2) continue;

      report.tabsProcessed++;

      // detect columns
      const matchCol =
        findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_SCOPE, 20) ||
        BOQ_SYNC.MASTER_MATCH_COL_FALLBACK;

      const locationCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_LOCATION, 20);

      // measurement may be absent in many tabs
      const measurementCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_MEASUREMENT, 20);

      const qtyColDetected = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_QTY, 20);
      const qtyCol = qtyColDetected || BOQ_SYNC.MASTER_QTY_COL_FALLBACK;

      const srNoCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_SRNO, 20);

      report.debug.masterDetectedCols.push(
        `${tabName}: matchCol=${matchCol}, locationCol=${locationCol}, measurementCol=${measurementCol || "null"}, qtyCol=${qtyCol} (detected=${qtyColDetected || "no"})`
      );

      // Minimal requirement: LOCATION + QTY must exist
      if (!locationCol || !qtyCol) {
        report.tabsSkipped.push(`${tabName} (missing LOCATION or QTY column)`);
        continue;
      }

      // snapshot scope values (keep aligned after insert)
      const scopeVals = sh.getRange(1, matchCol, lastRow, 1).getDisplayValues();

      let r = 1;
      while (r <= scopeVals.length) {
        const scopeText = String((scopeVals[r - 1] && scopeVals[r - 1][0]) || "").trim();
        if (!scopeText) { r++; continue; }

        const cell = sh.getRange(r, matchCol);

        // If merged: only process top-left; also store merged range to break safely if needed
        let mergedRange = null;
        if (cell.isPartOfMerge()) {
          const mrs = cell.getMergedRanges();
          mergedRange = (mrs && mrs.length) ? mrs[0] : null;

          if (mergedRange) {
            const isTopLeft =
              mergedRange.getRow() === r && mergedRange.getColumn() === matchCol;
            if (!isTopLeft) { r++; continue; }
          }
        }

        const rowKey = normKey_advanced_(scopeText);
        const mapList = mappingByTarget.get(rowKey);
        if (!mapList || !mapList.length) { r++; continue; }

        report.targetsMatched++;
        targetsFoundSomewhere.add(rowKey);

        // ---- Build union zones from MEAS + QTY (in stable order)
        const zoneOrder = [];
        const zoneTotalsMeas = new Map();
        const zoneTotalsQty = new Map();

        // MEAS (optional)
        // if (measurementCol) {
        //   for (const m of mapList) {
        //     const cacheKey = `${normKey_advanced_(m.generated)}||${normKey_advanced_(m.measure)}`;
        //     let breakdown = measCache.get(cacheKey);
        //     if (!breakdown) {
        //       breakdown = computeMeasurementFromLayer_(exportSS, m.generated, m.measure);
        //       measCache.set(cacheKey, breakdown);
        //     }

        //     for (const z of breakdown.order) pushZone_(zoneOrder, z);

        //     for (const [z, v] of breakdown.byZone.entries()) {
        //       const n = toNumber_(v);
        //       if (!Number.isFinite(n)) continue;
        //       zoneTotalsMeas.set(z, (zoneTotalsMeas.get(z) || 0) + n);
        //     }
        //   }
        // }
        if (measurementCol) {
            for (const m of mapList) {
              if (!m.generated) continue; // ✅ no layer name → no measurement

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

        // QTY (does NOT depend on BOQ-name presence anymore)
        for (const m of mapList) {
          const candidateQtyKeys = getCandidateQtyKeys_(m); // targeted -> blockName -> generated
          for (const key of candidateQtyKeys) {
            if (!key) continue;

            const rec = qtyIndex.get(key);
            if (!rec) continue;

            report.debug.qtyMatchesFound++;

            for (const z of rec.order) pushZone_(zoneOrder, z);

            for (const [z, v] of rec.byZone.entries()) {
              const n = toNumber_(v);
              if (!Number.isFinite(n)) continue;
              zoneTotalsQty.set(z, (zoneTotalsQty.get(z) || 0) + n);
            }

            // once matched with one candidate key, don’t double-add by other keys
            break;
          }
        }

        // If we got no zones from either source, nothing to write
        if (!zoneOrder.length) { r++; continue; }

        // ---- Filter zones: skip if BOTH meas & qty are blank/0
        const finalZones = [];
        for (const z of zoneOrder) {
          const meas = toNumber_(zoneTotalsMeas.get(z) || 0);
          const qty = toNumber_(zoneTotalsQty.get(z) || 0);

          const measZero = !Number.isFinite(meas) || meas === 0;
          const qtyZero = !Number.isFinite(qty) || qty === 0;

          if (BOQ_SYNC.SKIP_WHEN_BOTH_ZERO && measZero && qtyZero) {
            report.skippedBothZero++;
            continue;
          }
          finalZones.push(z);
        }

        if (!finalZones.length) { r++; continue; }

        const needed = finalZones.length;

        // ---- Insert rows if needed
        if (needed > 1) {
          if (mergedRange) mergedRange.breakApart();

          if (srNoCol) {
            const srCell = sh.getRange(r, srNoCol);
            if (srCell.isPartOfMerge()) {
              const srMr = srCell.getMergedRanges();
              if (srMr && srMr.length) srMr[0].breakApart();
            }
          }

          sh.insertRowsAfter(r, needed - 1);
          report.rowsInserted += (needed - 1);

          // keep scopeVals aligned
          const blanks = Array.from({ length: needed - 1 }, () => [""]);
          scopeVals.splice(r, 0, ...blanks);

          lastRow += (needed - 1);

          // Copy formatting/formulas row → new rows
          const baseRowRange = sh.getRange(r, 1, 1, lastCol);

          for (let i = 1; i < needed; i++) {
            const newRowRange = sh.getRange(r + i, 1, 1, lastCol);
            baseRowRange.copyTo(newRowRange, { contentsOnly: false });

            // Clear left-side columns before LOCATION so split rows look clean
            const clearUpto = Math.max(1, locationCol - 1);
            sh.getRange(r + i, 1, 1, clearUpto).clearContent();
          }

          // Merge & center scope / sr no across newly created block
          mergeAndCenterSafe_(sh, r, matchCol, needed);
          if (srNoCol) mergeAndCenterSafe_(sh, r, srNoCol, needed);
        }

        // ---- Write rows (LOCATION, MEAS if present, QTY always)
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

      // ---- After finishing this tab: hide rows based on final logic
      if (BOQ_SYNC.HIDE_ZERO_ROWS_AFTER_SYNC) {
        const hideRes = hideZeroRowsInTab_(sh, matchCol, measurementCol, qtyCol);
        report.rowsHidden += hideRes.hidden;
        report.rowsUnhidden += hideRes.unhidden;
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

/* =========================================================
   Mapping read
========================================================= */
function readMappings_(mapSh, report) {
  const lastRow = mapSh.getLastRow();
  const lastCol = mapSh.getLastColumn();
  if (lastRow < 2) return [];

  const grid = mapSh.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  const out = [];
  for (let r = 1; r < grid.length; r++) {
    const targeted = String(grid[r][BOQ_SYNC.MAP_COL_TARGETED - 1] || "").trim();
    const generated = String(grid[r][BOQ_SYNC.MAP_COL_GENERATED - 1] || "").trim();  // may be blank
    const blockName = String(grid[r][BOQ_SYNC.MAP_COL_BLOCKNAME - 1] || "").trim();  // may be blank

    // Must have the master match text
    if (!targeted) continue;

    // ✅ allow qty-only mappings (generated can be blank)
    // But if BOTH are blank, we can’t do anything useful.
    if (!generated && !blockName) continue;

    out.push({
      targeted,
      generated,     // can be ""
      blockName,     // can be ""
      measure: "Area",
    });
  }

  // samples for debug
  for (const m of out.slice(0, 10)) {
    const keys = getCandidateQtyKeys_(m).filter(Boolean);
    report.debug.sampleMappingQtyKeys.push(
      `${m.blockName || "(blank)"} | ${m.generated || "(blank)"} | ${m.targeted || "(blank)"}  =>  [${keys.join(" , ")}]`
    );
  }

  return out;
}

function getCandidateQtyKeys_(m) {
  // ✅ Prefer targeted first (matches your QTY sheet’s BOQ/Product better)
  const keys = [];
  if (m.targeted) keys.push(normKey_advanced_(m.targeted));
  if (m.blockName) keys.push(normKey_advanced_(m.blockName));
  if (m.generated) keys.push(normKey_advanced_(m.generated));

  const out = [];
  const seen = new Set();
  for (const k of keys) {
    if (!k) continue;
    if (seen.has(k)) continue;
    seen.add(k);
    out.push(k);
  }
  return out;
}

/* =========================================================
   Build QTY index (NOT dependent on BOQ merged cells)
========================================================= */
function buildQtyIndex_(exportSS, report) {
  const sh = exportSS.getSheetByName(BOQ_SYNC.EXPORT_QTY_TAB);
  if (!sh) throw new Error(`QTY tab not found: ${BOQ_SYNC.EXPORT_QTY_TAB}`);

  const dr = sh.getDataRange().getValues();
  if (dr.length < 2) throw new Error(`QTY tab "${BOQ_SYNC.EXPORT_QTY_TAB}" has no data`);

  const header = dr[0].map(v => String(v || "").trim());

  const idxProduct = findHeaderIndexContains_(header, BOQ_SYNC.QTY_HDR_PRODUCT_PATTERNS);
  const idxBoq = findHeaderIndexContains_(header, BOQ_SYNC.QTY_HDR_BOQNAME_PATTERNS);
  const idxZone = findHeaderIndexContains_(header, BOQ_SYNC.QTY_HDR_ZONE_PATTERNS);
  const idxQty = findHeaderIndexContains_(header, BOQ_SYNC.QTY_HDR_QTYVALUE_PATTERNS);

  report.debug.qtyDetectedCols.push(
    `QTY tab "${BOQ_SYNC.EXPORT_QTY_TAB}": idxProduct=${idxProduct}, idxBoq=${idxBoq}, idxZone=${idxZone}, idxQty=${idxQty} (0-based)`
  );

  if (idxZone === -1 || idxQty === -1) {
    throw new Error(
      `QTY tab header detect failed. Found idxZone=${idxZone}, idxQty=${idxQty}. Check row 1 headers.`
    );
  }
  if (idxProduct === -1 && idxBoq === -1) {
    throw new Error(
      `QTY tab must have either "Product" or "BOQ name" header. Found idxProduct=${idxProduct}, idxBoq=${idxBoq}.`
    );
  }

  report.debug.qtyRowsScanned = dr.length - 1;

  const index = new Map();

  function addToIndex(keyRaw, zone, qty) {
    const key = normKey_advanced_(keyRaw);
    if (!key) return;

    if (!index.has(key)) index.set(key, { order: [], byZone: new Map() });
    const rec = index.get(key);

    if (!rec.byZone.has(zone)) rec.order.push(zone);
    rec.byZone.set(zone, (rec.byZone.get(zone) || 0) + qty);
  }

  for (let r = 1; r < dr.length; r++) {
    const product = idxProduct !== -1 ? String(dr[r][idxProduct] || "").trim() : "";
    const boqName = idxBoq !== -1 ? String(dr[r][idxBoq] || "").trim() : "";

    const zone = String(dr[r][idxZone] || "").trim() || "misc";
    const qty = toNumber_(dr[r][idxQty]);
    if (!Number.isFinite(qty)) continue;

    // ✅ Primary: Product (works even if BOQ name is blank)
    if (product) addToIndex(product, zone, qty);

    // ✅ Secondary: BOQ name (backward compatibility)
    if (boqName) addToIndex(boqName, zone, qty);
  }

  report.debug.qtyUniqueBoqKeys = index.size;

  let i = 0;
  for (const [k] of index.entries()) {
    report.debug.sampleQtyBoqKeys.push(k);
    i++;
    if (i >= 14) break;
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

/* =========================================================
   Measurement from LAYER
========================================================= */
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

/* =========================================================
   HIDE rows logic (post-sync)
========================================================= */
function hideZeroRowsInTab_(sh, matchCol, measurementCol, qtyCol) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { hidden: 0, unhidden: 0 };

  const qtyVals = sh.getRange(1, qtyCol, lastRow, 1).getValues();
  const measVals = measurementCol ? sh.getRange(1, measurementCol, lastRow, 1).getValues() : null;

  let hidden = 0;
  let unhidden = 0;

  for (let r = 2; r <= lastRow; r++) {
    const scope = String(sh.getRange(r, matchCol).getDisplayValue() || "").trim();
    if (!scope) continue;

    // ✅ Never hide Total / Sheet Total rows
    if (isProtectedTotalRow_(scope)) {
      if (sh.isRowHiddenByUser(r)) {
        sh.showRows(r);
        unhidden++;
      }
      continue;
    }

    const qty = toNumber_(qtyVals[r - 1][0]);
    const qtyZero = !Number.isFinite(qty) || qty === 0;

    if (!measurementCol) {
      if (qtyZero) {
        if (!sh.isRowHiddenByUser(r)) { sh.hideRows(r); hidden++; }
      } else {
        if (sh.isRowHiddenByUser(r)) { sh.showRows(r); unhidden++; }
      }
      continue;
    }

    const meas = toNumber_(measVals[r - 1][0]);
    const measZero = !Number.isFinite(meas) || meas === 0;

    if (qtyZero && measZero) {
      if (!sh.isRowHiddenByUser(r)) { sh.hideRows(r); hidden++; }
    } else {
      if (sh.isRowHiddenByUser(r)) { sh.showRows(r); unhidden++; }
    }
  }

  return { hidden, unhidden };
}

function isProtectedTotalRow_(scopeText) {
  const s = normKey_advanced_(scopeText);
  return s === "total" || s === "sheet total";
}

/* =========================================================
   Master helpers
========================================================= */
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
  const rows = Math.min(scanRows || 20, sheet.getLastRow());
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

/* =========================================================
   Formatting helpers
========================================================= */
function mergeAndCenterSafe_(sh, startRow, col, numRows) {
  if (numRows <= 1) return;

  const rng = sh.getRange(startRow, col, numRows, 1);

  if (rng.isPartOfMerge()) {
    const mrs = rng.getMergedRanges();
    for (const mr of mrs) mr.breakApart();
  }

  rng.merge();
  rng.setHorizontalAlignment("center");
  rng.setVerticalAlignment("middle");
}

/* =========================================================
   Dialog helpers
========================================================= */
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
  lines.push(`Rows hidden: ${report.rowsHidden}`);
  lines.push(`Rows unhidden: ${report.rowsUnhidden}`);
  lines.push(`Targets not found: ${report.notFoundTargets}`);
  lines.push("");

  lines.push("DEBUG — QTY sheet:");
  for (const s of report.debug.qtyDetectedCols) lines.push(` - ${s}`);
  lines.push(` - rows scanned: ${report.debug.qtyRowsScanned}`);
  lines.push(` - unique BOQ keys: ${report.debug.qtyUniqueBoqKeys}`);
  lines.push("");

  lines.push("DEBUG — sample mapping qty keys (raw => candidates):");
  for (const s of report.debug.sampleMappingQtyKeys.slice(0, 10)) lines.push(` - ${s}`);
  lines.push("");

  lines.push("DEBUG — sample QTY keys (normalized):");
  for (const s of report.debug.sampleQtyBoqKeys) lines.push(` - ${s}`);
  lines.push("");

  lines.push("DEBUG — Master detected columns:");
  for (const s of report.debug.masterDetectedCols) lines.push(` - ${s}`);

  if (report.tabsSkipped.length) {
    lines.push("");
    lines.push("Skipped tabs:");
    for (const t of report.tabsSkipped) lines.push(` - ${t}`);
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

/* =========================================================
   Utils
========================================================= */
function pushZone_(arr, zone) {
  const z = (zone && String(zone).trim()) ? String(zone).trim() : "misc";
  if (!arr.includes(z)) arr.push(z);
}

function toNumber_(v) {
  if (v == null || v === "") return NaN;
  const n = Number(v);
  return Number.isFinite(n) ? n : NaN;
}

/**
 * Strong normalization:
 * - lowercase
 * - replace NBSP/newlines
 * - normalize dash variants
 * - replace all non-alphanum with space
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