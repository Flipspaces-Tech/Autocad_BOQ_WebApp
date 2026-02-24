/**************************************************
 * LAYER → MASTER (Zones vertical + per-zone Area)
 * - Source: Auto-QA Output sheet, tab "LAYER"
 * - Target: MASTER sheet tab "CALCULATION SHEET"
 *
 * Target behavior:
 * - Column A (S. No) merged vertically to match the item block
 * - Column B (BOQ item) merged vertically to match the item block
 * - Column C lists zones vertically
 * - Column D lists EACH zone’s area (NOT summed)
 * - Writes exact decimals (no rounding) + formats display as 0.############
 **************************************************/

const LAYER_TO_MASTER = {
  // ---- SOURCE (Auto-QA Output) ----
  SRC_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",
  SRC_TAB: "LAYER",
  SRC_HDR_LAYER: "layer",
  SRC_HDR_ZONE: "zone",
  SRC_HDR_AREA: "area (ft2)",

  // ---- TARGET (MASTER) ----
  TGT_TAB: "CALCULATION SHEET",

  // Target columns
  TGT_SNO_COL: 1,   // A  ✅ merge this too
  TGT_ITEM_COL: 2,  // B
  TGT_ZONES_COL: 3, // C
  TGT_AREA_COL: 4,  // D

  // Scan limits
  TGT_SCAN_START_ROW: 1,
  TGT_SCAN_END_ROW: 1200,

  // Rules
  NORMALIZE_UNMARKED_TO_MISC: true,

  // Formatting copy width when inserting rows
  COPY_FORMAT_COLS: 20,
};

function syncLayerToMasterZonesArea() {
  // ✅ Source
  const srcSS = SpreadsheetApp.openById(LAYER_TO_MASTER.SRC_SS_ID);
  const srcSh = srcSS.getSheetByName(LAYER_TO_MASTER.SRC_TAB);
  if (!srcSh) throw new Error(`Source tab not found: ${LAYER_TO_MASTER.SRC_TAB}`);

  // ✅ Target = MASTER (script hosted in MASTER)
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const tgtSh = masterSS.getSheetByName(LAYER_TO_MASTER.TGT_TAB);
  if (!tgtSh) throw new Error(`Target tab not found in MASTER: ${LAYER_TO_MASTER.TGT_TAB}`);

  const lookup = buildLayerAggMap_(srcSh); // key -> { zoneArea: Map(zone -> area) }

  const lastRow = Math.min(LAYER_TO_MASTER.TGT_SCAN_END_ROW, tgtSh.getLastRow());
  const startRow = Math.max(1, LAYER_TO_MASTER.TGT_SCAN_START_ROW);
  if (lastRow < startRow) return;

  const items = tgtSh
    .getRange(startRow, LAYER_TO_MASTER.TGT_ITEM_COL, lastRow - startRow + 1, 1)
    .getDisplayValues();

  let updated = 0;

  for (let i = 0; i < items.length; i++) {
    const r = startRow + i;
    const rawItem = String(items[i][0] || "").trim();
    if (!rawItem) continue;

    // Only process top-left cell of a merged item block in column B
    const itemCell = tgtSh.getRange(r, LAYER_TO_MASTER.TGT_ITEM_COL);
    if (itemCell.isPartOfMerge()) {
      const mergedRanges = itemCell.getMergedRanges();
      if (mergedRanges && mergedRanges.length) {
        const merged = mergedRanges[0];
        if (!(merged.getRow() === r && merged.getColumn() === LAYER_TO_MASTER.TGT_ITEM_COL)) continue;
      }
    }

    const key = normKey_(rawItem);
    const agg = lookup.get(key);
    if (!agg || !agg.zoneArea) continue;

    // Zones in Auto-QA order (Map preserves insertion order)
    const zones = Array.from(agg.zoneArea.keys());
    const areas = zones.map(z => agg.zoneArea.get(z) || 0);

    // Identify the current block height based on merged B (or 1)
    const block = getItemBlockRange_(tgtSh, r, LAYER_TO_MASTER.TGT_ITEM_COL);
    let blockRows = block.numRows;

    // Expand rows if needed
    if (zones.length > blockRows) {
      const add = zones.length - blockRows;

      // Insert rows after the block and copy formatting
      insertRowsWithFormat_(
        tgtSh,
        block.startRow + blockRows - 1,
        add,
        LAYER_TO_MASTER.COPY_FORMAT_COLS
      );

      // ✅ Re-merge A (S.No) to match new height
      const snoOld = tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_SNO_COL, blockRows, 1);
      if (snoOld.isPartOfMerge()) snoOld.getMergedRanges()[0].breakApart();
      const snoNew = tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_SNO_COL, zones.length, 1);
      snoNew.merge().setVerticalAlignment("middle"); // ✅ center the number

      // ✅ Re-merge B (Item) to match new height
      const itemOld = tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_ITEM_COL, blockRows, 1);
      if (itemOld.isPartOfMerge()) itemOld.getMergedRanges()[0].breakApart();
      const itemNew = tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_ITEM_COL, zones.length, 1);
      itemNew.merge().setVerticalAlignment("middle"); // (optional: keep item centered too)

      blockRows = zones.length;
    } else {
      // Even if we didn't expand, ensure A merge matches B merge for that block
      // (prevents cases where B is merged but A isn't)
      const snoCell = tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_SNO_COL);
      const snoBlock = tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_SNO_COL, blockRows, 1);
      if (!snoCell.isPartOfMerge() && blockRows > 1) {
        snoBlock.merge().setVerticalAlignment("middle");
      } else if (snoCell.isPartOfMerge()) {
        // keep vertical alignment consistent
        snoCell.getMergedRanges()[0].setVerticalAlignment("middle");
      }
    }

    // Write zones (Column C)
    const zoneWrite = [];
    for (let z = 0; z < blockRows; z++) zoneWrite.push([zones[z] || ""]);
    tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_ZONES_COL, blockRows, 1).setValues(zoneWrite);

    // Write per-zone areas (Column D) with exact decimals + display format
    const areaWrite = [];
    for (let z = 0; z < blockRows; z++) areaWrite.push([areas[z] ?? ""]);
    const areaRange = tgtSh.getRange(block.startRow, LAYER_TO_MASTER.TGT_AREA_COL, blockRows, 1);
    areaRange.setValues(areaWrite);
    areaRange.setNumberFormat("0.############"); // ✅ no visual rounding

    // Clear extra rows if zones are fewer than existing block height
    if (zones.length < blockRows) {
      const clearFrom = block.startRow + zones.length;
      const clearCount = blockRows - zones.length;

      tgtSh.getRange(clearFrom, LAYER_TO_MASTER.TGT_ZONES_COL, clearCount, 1).clearContent();
      tgtSh.getRange(clearFrom, LAYER_TO_MASTER.TGT_AREA_COL, clearCount, 1).clearContent();
    }

    updated++;
  }

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
  const idxZone  = header.indexOf(LAYER_TO_MASTER.SRC_HDR_ZONE.toLowerCase());
  const idxArea  = header.indexOf(LAYER_TO_MASTER.SRC_HDR_AREA.toLowerCase());

  if (idxLayer === -1) throw new Error(`Missing header in LAYER: ${LAYER_TO_MASTER.SRC_HDR_LAYER}`);
  if (idxZone  === -1) throw new Error(`Missing header in LAYER: ${LAYER_TO_MASTER.SRC_HDR_ZONE}`);
  if (idxArea  === -1) throw new Error(`Missing header in LAYER: ${LAYER_TO_MASTER.SRC_HDR_AREA}`);

  const map = new Map();
  let currentLayer = ""; // ✅ carry merged layer value downward

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
      agg = { zoneArea: new Map() }; // zone -> area
      map.set(key, agg);
    }

    // If same zone repeats multiple times, accumulate within the zone
    const prev = agg.zoneArea.get(zoneRaw) || 0;
    agg.zoneArea.set(zoneRaw, prev + (area > 0 ? area : 0));
  }

  return map;
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
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}