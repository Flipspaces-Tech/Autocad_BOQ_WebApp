/**************************************************
 * BOQ-LAYER MAPPER (Image1) → MASTER BOQ (Image2)
 *
 * Image1 = mapping sheet (BOQ-LAYER) [script hosted here]
 * Image3 = export sheet (LAYER/PLANNER/...)
 * Image2 = master BOQ (tabs MASTER_START_TAB → last)
 *
 * Current supported export tab: LAYER
 *
 * ✅ Splits output rows by LOCATION (Zone) from Image3 "zone" column
 * ✅ SKIPS any zone row where total value === 0 (won’t write / won’t create rows)
 * ✅ Merges & centers "SCOPE OF WORK" (and SR NO if present) across inserted rows
 * ✅ Shows a detailed dialog summary after every run
 **************************************************/

const SHEETS = SpreadsheetApp; // ✅ prevents "SpreadsheetApp overwritten" bugs

const BOQ_SYNC = {
  // ---- IMAGE 1 (script hosted here) ----
  MAP_TAB: "BOQ-LAYER",

  // Mapping columns in Image1 tab (1-based)
  MAP_COL_TARGETED: 2, // B = Targeted BOQ Name (to find in MASTER)
  MAP_COL_GENERATED: 3, // C = Generated Layer Name (to find in EXPORT)
  MAP_COL_MEASURE: 6, // F = Measurement (Area/Perimeter/Length/Count)
  MAP_COL_UNITS: 8, // H = Units (optional for future)

  // ---- IMAGE 3 (Export Output) ----
  EXPORT_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",
  EXPORT_DEFAULT_TAB: "LAYER",

  // LAYER tab header names (case-insensitive)
  LAYER_HDR_LAYER: "layer",
  LAYER_HDR_ZONE: "zone",
  LAYER_HDR_AREA: "area (ft2)",
  LAYER_HDR_PERIM: "perimeter",
  LAYER_HDR_LENGTH: "length (ft)",

  // ---- IMAGE 2 (Master BOQ) ----
  MASTER_SS_ID: "1CVibwjRFz4gTATAeOFUlYzlGZybXILO60OrxwGaFLeY",
  MASTER_START_TAB: "Civil",

  // MASTER matching/writing behavior
  MASTER_MATCH_COL_FALLBACK: 2, // fallback = col B if header search fails

  // Formatting
  NUMBER_FORMAT: "0.############",

  // Normalization
  NORMALIZE_SPACES: true,

  // ✅ NEW: Skip zero rows
  SKIP_ZERO_VALUES: true, // turn off if needed

  // Dialog behavior
  SHOW_DIALOG: true,
  DIALOG_TITLE: "Vizdom Sync — BOQ-LAYER → MASTER",
};

function onOpen() {
  SHEETS.getUi()
    .createMenu("Vizdom Sync")
    .addItem("Sync BOQ-LAYER → MASTER", "syncBoqLayerToMaster")
    .addToUi();
}

/**
 * Main runner
 */
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
    rowsSkippedZero: 0,
    notFoundTargets: 0,
    errors: [],
    notes: [],
  };

  try {
    const mapSS = SHEETS.getActiveSpreadsheet();
    const mapSh = mapSS.getSheetByName(BOQ_SYNC.MAP_TAB);
    if (!mapSh) throw new Error(`Mapping tab not found: ${BOQ_SYNC.MAP_TAB}`);

    const exportSS = SHEETS.openById(BOQ_SYNC.EXPORT_SS_ID);
    const masterSS = SHEETS.openById(BOQ_SYNC.MASTER_SS_ID);

    // 1) Read mapping rows
    const mappings = readMappings_(mapSh);
    report.mappingsTotal = mappings.length;

    if (!mappings.length) {
      SHEETS.getActive().toast("No mappings found in BOQ-LAYER.", "Vizdom Sync", 6);
      report.notes.push("No mappings found in BOQ-LAYER.");
      finalize_(report);
      if (BOQ_SYNC.SHOW_DIALOG) showReportDialog_(ui, report);
      return;
    }

    // 2) Cache export computations
    const exportBreakdownCache = new Map();

    // 3) Process master tabs from MASTER_START_TAB → last
    const tabs = getMasterTabsFromStart_(masterSS, BOQ_SYNC.MASTER_START_TAB);
    if (!tabs.length) {
      throw new Error(`Start tab "${BOQ_SYNC.MASTER_START_TAB}" not found in MASTER.`);
    }

    // Build lookup: targetedKey -> array of mappings
    const mappingByTarget = new Map();
    for (const m of mappings) {
      const tKey = normKey_(m.targeted);
      if (!mappingByTarget.has(tKey)) mappingByTarget.set(tKey, []);
      mappingByTarget.get(tKey).push(m);
    }

    const targetsFoundSomewhere = new Set();

    for (const sh of tabs) {
      const tabName = sh.getName();
      const lastCol = sh.getLastColumn();
      let lastRow = sh.getLastRow();
      if (lastRow < 2 || lastCol < 2) continue;

      report.tabsProcessed++;

      // Match column: "SCOPE OF WORK" or fallback to col B
      const matchCol =
        findHeaderCol_(sh, ["scope of work"], 6) || BOQ_SYNC.MASTER_MATCH_COL_FALLBACK;

      // Output columns: LOCATION + MEASUREMENT
      const locationCol = findHeaderCol_(sh, ["location"], 6);
      const measurementCol = findHeaderCol_(sh, ["measurement"], 6);

      if (!locationCol || !measurementCol) {
        report.tabsSkippedMissingCols.push(`${tabName} (missing LOCATION/MEASUREMENT headers)`);
        continue;
      }

      // Optional SR NO column if header exists
      const srNoCol = findHeaderCol_(sh, ["sr. no.", "sr no", "sr.no", "sr"], 6);

      // Read match column values once (keep aligned when inserting)
      const scopeVals = sh.getRange(1, matchCol, lastRow, 1).getDisplayValues();

      let r = 1;
      while (r <= scopeVals.length) {
        const rowArr = scopeVals[r - 1];
        if (!rowArr) break;

        const scopeText = String(rowArr[0] || "").trim();
        if (!scopeText) {
          r++;
          continue;
        }

        // If row is part of merge, only handle top-left
        const cell = sh.getRange(r, matchCol);
        if (cell.isPartOfMerge()) {
          const mr = cell.getMergedRanges()[0];
          if (!(mr.getRow() === r && mr.getColumn() === matchCol)) {
            r++;
            continue;
          }
        }

        const rowKey = normKey_(scopeText);
        const mapList = mappingByTarget.get(rowKey);
        if (!mapList || !mapList.length) {
          r++;
          continue;
        }

        report.targetsMatched++;
        targetsFoundSomewhere.add(rowKey);

        // Build zone totals for this Targeted
        const zoneOrder = [];
        const zoneSeen = new Set();
        const zoneTotals = new Map();

        for (const m of mapList) {
          const cacheKey = `${m.sourceTab}||${normKey_(m.measure)}||${normKey_(m.generated)}`;

          let breakdown;
          if (exportBreakdownCache.has(cacheKey)) {
            breakdown = exportBreakdownCache.get(cacheKey);
          } else {
            breakdown = computeFromExport_(exportSS, m.sourceTab, m.generated, m.measure);
            exportBreakdownCache.set(cacheKey, breakdown);
          }

          for (const z of breakdown.order) {
            if (!zoneSeen.has(z)) {
              zoneSeen.add(z);
              zoneOrder.push(z);
            }
          }

          for (const [z, v] of breakdown.byZone.entries()) {
            const n = toNumber_(v);
            if (!Number.isFinite(n)) continue;
            zoneTotals.set(z, (zoneTotals.get(z) || 0) + n);
          }
        }

        // ✅ FILTER: keep only non-zero zones (in order)
        const filteredZones = [];
        for (const z of zoneOrder.length ? zoneOrder : ["misc"]) {
          const val = toNumber_(zoneTotals.get(z) || 0);
          const isZero = !Number.isFinite(val) || val === 0;
          if (BOQ_SYNC.SKIP_ZERO_VALUES && isZero) {
            report.rowsSkippedZero++;
            continue;
          }
          filteredZones.push(z);
        }

        // If everything is zero → do nothing (don’t insert, don’t write)
        if (!filteredZones.length) {
          r++;
          continue;
        }

        const needed = filteredZones.length;

        // Insert rows if needed and copy base row formatting/formulas
        if (needed > 1) {
          sh.insertRowsAfter(r, needed - 1);
          report.rowsInserted += (needed - 1);

          // keep scopeVals aligned
          const blanks = Array.from({ length: needed - 1 }, () => [""]);
          scopeVals.splice(r, 0, ...blanks);

          lastRow += (needed - 1);

          const baseRowRange = sh.getRange(r, 1, 1, lastCol);

          for (let i = 1; i < needed; i++) {
            const newRowRange = sh.getRange(r + i, 1, 1, lastCol);
            baseRowRange.copyTo(newRowRange, { contentsOnly: false });

            // Clear left-side columns (before LOCATION) so split rows look clean
            const clearUpto = Math.max(1, locationCol - 1);
            sh.getRange(r + i, 1, 1, clearUpto).clearContent();
          }

          // Merge & center SCOPE OF WORK + SR NO
          mergeAndCenter_(sh, r, matchCol, needed);
          if (srNoCol) mergeAndCenter_(sh, r, srNoCol, needed);
        }

        // Fill LOCATION + MEASUREMENT per non-zero zone
        for (let i = 0; i < needed; i++) {
          const zone = filteredZones[i];
          const val = toNumber_(zoneTotals.get(zone) || 0);

          sh.getRange(r + i, locationCol).setValue(zone);

          const mCell = sh.getRange(r + i, measurementCol);
          mCell.setValue(val);
          mCell.setNumberFormat(BOQ_SYNC.NUMBER_FORMAT);
        }

        report.rowsUpdated += needed;

        // skip past inserted rows
        r += needed;
      }
    }

    // Count targets not found anywhere
    let notFound = 0;
    for (const m of mappings) {
      const tKey = normKey_(m.targeted);
      if (!targetsFoundSomewhere.has(tKey)) notFound++;
    }
    report.notFoundTargets = notFound;

    SHEETS.getActive().toast(
      `Done. Updated ${report.rowsUpdated} row(s). Zero rows skipped: ${report.rowsSkippedZero}. Targets not found: ${report.notFoundTargets}`,
      "BOQ-LAYER → MASTER",
      10
    );

    finalize_(report);
    if (BOQ_SYNC.SHOW_DIALOG) showReportDialog_(ui, report);
  } catch (e) {
    report.errors.push(String(e && e.stack ? e.stack : e));
    finalize_(report);
    if (BOQ_SYNC.SHOW_DIALOG) showReportDialog_(SHEETS.getUi(), report, true);
    throw e;
  }
}

/* =========================
   Mapping read
========================= */

function readMappings_(mapSh) {
  const lastRow = mapSh.getLastRow();
  const lastCol = mapSh.getLastColumn();
  if (lastRow < 2) return [];

  const values = mapSh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const out = [];

  for (let r = 1; r < values.length; r++) {
    const targeted = String(values[r][BOQ_SYNC.MAP_COL_TARGETED - 1] || "").trim();
    const generated = String(values[r][BOQ_SYNC.MAP_COL_GENERATED - 1] || "").trim();
    const measure = String(values[r][BOQ_SYNC.MAP_COL_MEASURE - 1] || "").trim();
    const units = String(values[r][BOQ_SYNC.MAP_COL_UNITS - 1] || "").trim();

    if (!targeted || !generated) continue;

    out.push({
      targeted,
      generated,
      measure: measure || "Area",
      units,
      sourceTab: BOQ_SYNC.EXPORT_DEFAULT_TAB,
    });
  }

  return out;
}

/* =========================
   Export compute
========================= */

function computeFromExport_(exportSS, tabName, generatedName, measure) {
  const sh = exportSS.getSheetByName(tabName);
  if (!sh) throw new Error(`Export tab not found: ${tabName}`);

  if (normKey_(tabName) === "layer") {
    return computeFromLayerTab_(sh, generatedName, measure);
  }

  throw new Error(`Export tab "${tabName}" not supported yet.`);
}

function computeFromLayerTab_(sh, layerName, measure) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return { order: [], byZone: new Map() };

  const values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0].map(v => String(v || "").trim().toLowerCase());

  const idxLayer = header.indexOf(BOQ_SYNC.LAYER_HDR_LAYER.toLowerCase());
  const idxZone = header.indexOf(BOQ_SYNC.LAYER_HDR_ZONE.toLowerCase());
  const idxArea = header.indexOf(BOQ_SYNC.LAYER_HDR_AREA.toLowerCase());
  const idxPerim = header.indexOf(BOQ_SYNC.LAYER_HDR_PERIM.toLowerCase());
  const idxLen = header.indexOf(BOQ_SYNC.LAYER_HDR_LENGTH.toLowerCase());

  if (idxLayer === -1) throw new Error(`LAYER tab missing header: ${BOQ_SYNC.LAYER_HDR_LAYER}`);
  if (idxZone === -1) throw new Error(`LAYER tab missing header: ${BOQ_SYNC.LAYER_HDR_ZONE}`);

  const m = normKey_(measure);
  let idxMeasure = idxArea;
  let isCount = false;

  if (m === "perimeter") idxMeasure = idxPerim;
  else if (m === "length") idxMeasure = idxLen;
  else if (m === "count") isCount = true;

  if (!isCount && idxMeasure === -1) return { order: [], byZone: new Map() };

  const targetKey = normKey_(layerName);

  const order = [];
  const seen = new Set();
  const byZone = new Map();

  let currentLayer = "";

  for (let r = 1; r < values.length; r++) {
    const layerCell = String(values[r][idxLayer] || "").trim();
    if (layerCell) currentLayer = layerCell;
    if (!currentLayer) continue;

    if (normKey_(currentLayer) !== targetKey) continue;

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

/* =========================
   Master helpers
========================= */

function getMasterTabsFromStart_(masterSS, startName) {
  const sheets = masterSS.getSheets();
  const startKey = normKey_(startName);

  let startIdx = -1;
  for (let i = 0; i < sheets.length; i++) {
    if (normKey_(sheets[i].getName()) === startKey) {
      startIdx = i;
      break;
    }
  }
  if (startIdx === -1) return [];
  return sheets.slice(startIdx);
}

function findHeaderCol_(sheet, headerCandidatesLower, scanRows) {
  const lastCol = sheet.getLastColumn();
  const rows = Math.min(scanRows || 5, sheet.getLastRow());
  if (rows < 1 || lastCol < 1) return null;

  const grid = sheet.getRange(1, 1, rows, lastCol).getDisplayValues();
  const wanted = headerCandidatesLower.map(h => String(h).trim().toLowerCase());

  for (let r = 0; r < grid.length; r++) {
    const row = grid[r].map(v => String(v || "").trim().toLowerCase());
    for (let c = 0; c < row.length; c++) {
      if (!row[c]) continue;
      if (wanted.includes(row[c])) return c + 1;
    }
  }
  return null;
}

/* =========================
   Formatting helpers
========================= */

function mergeAndCenter_(sh, startRow, col, numRows) {
  if (numRows <= 1) return;

  const rng = sh.getRange(startRow, col, numRows, 1);
  if (rng.isPartOfMerge()) rng.breakApart();

  rng.merge();
  rng.setHorizontalAlignment("center");
  rng.setVerticalAlignment("middle");
}

/* =========================
   Dialog helpers
========================= */

function finalize_(report) {
  report.finishedAt = new Date();
}

function showReportDialog_(ui, report, isError) {
  const durMs = report.finishedAt - report.startedAt;
  const durSec = Math.round(durMs / 1000);

  const lines = [];
  lines.push(`Started:  ${report.startedAt.toLocaleString()}`);
  lines.push(`Finished: ${report.finishedAt.toLocaleString()}`);
  lines.push(`Duration: ${durSec}s`);
  lines.push("");
  lines.push(`Mappings loaded: ${report.mappingsTotal}`);
  lines.push(`Tabs processed:  ${report.tabsProcessed}`);
  lines.push(`Targets matched: ${report.targetsMatched}`);
  lines.push(`Rows inserted:   ${report.rowsInserted}`);
  lines.push(`Rows updated:    ${report.rowsUpdated}`);
  lines.push(`Zero rows skipped: ${report.rowsSkippedZero}`);
  lines.push(`Targets not found: ${report.notFoundTargets}`);

  if (report.tabsSkippedMissingCols.length) {
    lines.push("");
    lines.push("Skipped tabs (missing LOCATION/MEASUREMENT):");
    for (const t of report.tabsSkippedMissingCols) lines.push(` - ${t}`);
  }

  if (report.notes.length) {
    lines.push("");
    lines.push("Notes:");
    for (const n of report.notes) lines.push(` - ${n}`);
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

/* =========================
   Utils
========================= */

function normKey_(s) {
  let t = String(s || "").trim().toLowerCase();
  if (BOQ_SYNC.NORMALIZE_SPACES) t = t.replace(/\s+/g, " ");
  return t;
}

function toNumber_(v) {
  if (v == null || v === "") return NaN;
  const n = Number(v);
  return Number.isFinite(n) ? n : NaN;
}