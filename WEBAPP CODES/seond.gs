/**************************************************
 * LAYER → MASTER (Zones vertical + per-zone Area)
 * - Source: Auto-QA Output sheet, tab "LAYER"
 * - Target: MASTER sheet tab "SITE RECEE CALCULATION"
 *
 * Behavior:
 * - Column A (S. No) merged vertically to match zones count
 * - Column B (Item) merged vertically to match zones count
 * - Column C lists zones vertically
 * - Column D lists EACH zone’s area (NOT summed)
 * - Auto EXPANDS or SHRINKS rows to fit zones count (adaptable)
 * - Exact decimals + display format 0.############
 * - Hides row if LENGTH (Ft) is blank or 0
 * - Does NOT touch row 1
 **************************************************/

const LAYER_TO_MASTER = {
  // ---- SOURCE (Auto-QA Output) ----
  SRC_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",
  SRC_TAB: "LAYER",
  SRC_HDR_LAYER: "layer",
  SRC_HDR_ZONE: "zone",
  SRC_HDR_AREA: "area (ft2)",

  // ---- TARGET (MASTER) ----
  TGT_TAB: "SITE RECEE CALCULATION",

  // Target columns
  TGT_SNO_COL: 1,     // A
  TGT_ITEM_COL: 2,    // B
  TGT_ZONES_COL: 3,   // C
  TGT_AREA_COL: 4,    // D
  TGT_LENGTH_COL: 4,  // D = LENGTH (Ft) in your sheet

  // Scan limits
  TGT_SCAN_START_ROW: 2,   // don't touch row 1
  TGT_SCAN_END_ROW: 1200,

  // Rules
  NORMALIZE_UNMARKED_TO_MISC: true,

  // Formatting copy width when inserting rows
  COPY_FORMAT_COLS: 20,
};

function syncLayerToMasterZonesArea() {
  const srcSS = SpreadsheetApp.openById(LAYER_TO_MASTER.SRC_SS_ID);
  const srcSh = srcSS.getSheetByName(LAYER_TO_MASTER.SRC_TAB);
  if (!srcSh) throw new Error(`Source tab not found: ${LAYER_TO_MASTER.SRC_TAB}`);

  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const tgtSh = masterSS.getSheetByName(LAYER_TO_MASTER.TGT_TAB);
  if (!tgtSh) throw new Error(`Target tab not found in MASTER: ${LAYER_TO_MASTER.TGT_TAB}`);

  const lookup = buildLayerAggMap_(srcSh); // key -> { zoneArea: Map(zone -> area) }

  const lastRow = Math.min(LAYER_TO_MASTER.TGT_SCAN_END_ROW, tgtSh.getLastRow());
  const startRow = Math.max(2, LAYER_TO_MASTER.TGT_SCAN_START_ROW);
  if (lastRow < startRow) return;

  const items = tgtSh
    .getRange(startRow, LAYER_TO_MASTER.TGT_ITEM_COL, lastRow - startRow + 1, 1)
    .getDisplayValues();

  let updated = 0;

  for (let i = 0; i < items.length; i++) {
    const r = startRow + i;
    const rawItem = String(items[i][0] || "").trim();
    if (!rawItem) continue;

    // Only process top-left cell of merged block in column B
    const itemCell = tgtSh.getRange(r, LAYER_TO_MASTER.TGT_ITEM_COL);
    if (itemCell.isPartOfMerge()) {
      const merged = itemCell.getMergedRanges()[0];
      if (!(merged.getRow() === r && merged.getColumn() === LAYER_TO_MASTER.TGT_ITEM_COL)) continue;
    }

    const key = normKey_(rawItem);
    const agg = lookup.get(key);
    if (!agg || !agg.zoneArea) continue;

    const zones = Array.from(agg.zoneArea.keys());
    const areas = zones.map(z => agg.zoneArea.get(z) || 0);

    const block = getItemBlockRange_(tgtSh, r, LAYER_TO_MASTER.TGT_ITEM_COL);
    const currentRows = block.numRows;
    const desiredRows = Math.max(1, zones.length);

    // ---------- ADAPT HEIGHT: expand or shrink ----------
    if (desiredRows > currentRows) {
      const add = desiredRows - currentRows;

      insertRowsWithFormat_(
        tgtSh,
        block.startRow + currentRows - 1,
        add,
        LAYER_TO_MASTER.COPY_FORMAT_COLS
      );

      remakeMerge_(tgtSh, block.startRow, LAYER_TO_MASTER.TGT_SNO_COL, currentRows, desiredRows);
      remakeMerge_(tgtSh, block.startRow, LAYER_TO_MASTER.TGT_ITEM_COL, currentRows, desiredRows);

    } else if (desiredRows < currentRows) {
      breakMergeIfAny_(tgtSh, block.startRow, LAYER_TO_MASTER.TGT_SNO_COL);
      breakMergeIfAny_(tgtSh, block.startRow, LAYER_TO_MASTER.TGT_ITEM_COL);

      const deleteFrom = block.startRow + desiredRows;
      const deleteCount = currentRows - desiredRows;
      tgtSh.deleteRows(deleteFrom, deleteCount);

      if (desiredRows > 1) {
        tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_SNO_COL, desiredRows, 1)
          .merge()
          .setVerticalAlignment("middle");
        tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_ITEM_COL, desiredRows, 1)
          .merge()
          .setVerticalAlignment("middle");
      } else {
        tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_SNO_COL).setVerticalAlignment("middle");
        tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_ITEM_COL).setVerticalAlignment("middle");
      }
    } else {
      const snoCell = tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_SNO_COL);
      if (!snoCell.isPartOfMerge() && desiredRows > 1) {
        tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_SNO_COL, desiredRows, 1)
          .merge()
          .setVerticalAlignment("middle");
      } else if (snoCell.isPartOfMerge()) {
        snoCell.getMergedRanges()[0].setVerticalAlignment("middle");
      }

      if (itemCell.isPartOfMerge()) {
        itemCell.getMergedRanges()[0].setVerticalAlignment("middle");
      }
    }

    const writeRows = desiredRows;

    // ---------- Write zones ----------
    const zoneWrite = [];
    for (let z = 0; z < writeRows; z++) zoneWrite.push([zones[z] || ""]);
    tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_ZONES_COL, writeRows, 1).setValues(zoneWrite);

    // ---------- Write per-zone areas ----------
    const areaWrite = [];
    for (let z = 0; z < writeRows; z++) areaWrite.push([areas[z] ?? ""]);
    const areaRange = tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_AREA_COL, writeRows, 1);
    areaRange.setValues(areaWrite);
    areaRange.setNumberFormat("0.############");

    updated++;
  }

  hideRowsWithBlankOrZeroLength_();

  SpreadsheetApp.getActive().toast(
    `Updated ${updated} BOQ item(s).`,
    "LAYER → MASTER",
    8
  );
}

/* ---------- Build lookup from LAYER tab (merge-aware) ---------- */

function buildLayerAggMap_(srcSh) {
  const lastRow = srcSh.getLastRow();
  const lastCol = srcSh.getLastColumn();
  if (lastRow < 2) return new Map();

  const values = srcSh.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0].map(v => String(v || "").trim().toLowerCase());

  const idxLayer = header.indexOf(LAYER_TO_MASTER.SRC_HDR_LAYER.toLowerCase());
  const idxZone = header.indexOf(LAYER_TO_MASTER.SRC_HDR_ZONE.toLowerCase());
  const idxArea = header.indexOf(LAYER_TO_MASTER.SRC_HDR_AREA.toLowerCase());

  if (idxLayer === -1) throw new Error(`Missing header in LAYER: ${LAYER_TO_MASTER.SRC_HDR_LAYER}`);
  if (idxZone === -1) throw new Error(`Missing header in LAYER: ${LAYER_TO_MASTER.SRC_HDR_ZONE}`);
  if (idxArea === -1) throw new Error(`Missing header in LAYER: ${LAYER_TO_MASTER.SRC_HDR_AREA}`);

  const map = new Map();
  let currentLayer = "";

  for (let r = 1; r < values.length; r++) {
    const layerCell = String(values[r][idxLayer] || "").trim();
    if (layerCell) currentLayer = layerCell;
    if (!currentLayer) continue;

    let zoneRaw = String(values[r][idxZone] || "").trim();
    zoneRaw = normalizeZone_(zoneRaw);
    if (!zoneRaw) continue;

    const area = toNumber_(values[r][idxArea]) || 0;

    const key = normKey_(currentLayer);
    let agg = map.get(key);
    if (!agg) {
      agg = { zoneArea: new Map() };
      map.set(key, agg);
    }

    const prev = agg.zoneArea.get(zoneRaw) || 0;
    agg.zoneArea.set(zoneRaw, prev + (area > 0 ? area : 0));
  }

  return map;
}

/* ---------- Hide rows where LENGTH (Ft) is blank or 0 ---------- */

function hideRowsWithBlankOrZeroLength_() {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const tgtSh = masterSS.getSheetByName(LAYER_TO_MASTER.TGT_TAB);
  if (!tgtSh) throw new Error(`Target tab not found in MASTER: ${LAYER_TO_MASTER.TGT_TAB}`);

  const startRow = Math.max(2, LAYER_TO_MASTER.TGT_SCAN_START_ROW);
  const lastRow = Math.min(LAYER_TO_MASTER.TGT_SCAN_END_ROW, tgtSh.getLastRow());
  if (lastRow < startRow) return;

  const itemVals = tgtSh
    .getRange(startRow, LAYER_TO_MASTER.TGT_ITEM_COL, lastRow - startRow + 1, 1)
    .getDisplayValues();

  const lenVals = tgtSh
    .getRange(startRow, LAYER_TO_MASTER.TGT_LENGTH_COL, lastRow - startRow + 1, 1)
    .getDisplayValues();

  for (let i = 0; i < itemVals.length; i++) {
    const rowNum = startRow + i;
    if (rowNum === 1) continue;

    const itemText = String(itemVals[i][0] || "").trim();
    const lenText = String(lenVals[i][0] || "").trim();
    const lenNum = toNumber_(lenText);

    if (!itemText) continue;

    if (/^\(.*\)$/.test(itemText)) continue;

    if (lenText === "" || lenNum === 0 || lenNum == null) {
      tgtSh.hideRows(rowNum);
    } else {
      tgtSh.showRows(rowNum);
    }
  }
}

/* ---------- Target helpers ---------- */

function getItemBlockRange_(sheet, row, col) {
  const cell = sheet.getRange(row, col);
  if (cell.isPartOfMerge()) {
    const m = cell.getMergedRanges()[0];
    return { startRow: m.getRow(), numRows: m.getNumRows() };
  }
  return { startRow: row, numRows: 1 };
}

function insertRowsWithFormat_(sheet, afterRow, howMany, copyCols) {
  sheet.insertRowsAfter(afterRow, howMany);
  const srcFmt = sheet.getRange(afterRow, 1, 1, copyCols);
  const dstFmt = sheet.getRange(afterRow + 1, 1, howMany, copyCols);
  srcFmt.copyTo(dstFmt, { formatOnly: true });
}

function breakMergeIfAny_(sheet, row, col) {
  const cell = sheet.getRange(row, col);
  if (cell.isPartOfMerge()) cell.getMergedRanges()[0].breakApart();
}

function remakeMerge_(sheet, startRow, col, oldRows, newRows) {
  const oldRange = sheet.getRange(startRow, col, oldRows, 1);
  if (oldRange.isPartOfMerge()) oldRange.getMergedRanges()[0].breakApart();

  if (newRows > 1) {
    sheet.getRange(startRow, col, newRows, 1).merge().setVerticalAlignment("middle");
  } else {
    sheet.getRange(startRow, col).setVerticalAlignment("middle");
  }
}

/* ---------- Utils ---------- */

function normalizeZone_(zone) {
  const z = String(zone || "").trim();
  if (!z) return "misc";
  if (LAYER_TO_MASTER.NORMALIZE_UNMARKED_TO_MISC && z.toLowerCase() === "unmarked area") return "misc";
  return z;
}

function normKey_(s) {
  return String(s || "").trim().toLowerCase().replace(/\s+/g, " ");
}

function toNumber_(v) {
  if (v == null || v === "") return null;
  const n = Number(String(v).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}