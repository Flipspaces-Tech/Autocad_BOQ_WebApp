/**************************************************
 * DOORS / OPENINGS → MASTER (Nos + Length + Zones vertical)
 *
 * MASTER: "CALCULATION SHEET" (Doors table)
 * MAPPING: Auto-QA Output Template, tab "BOQ-LAYER"
 * GENERATED: Auto-QA Output, tab "vis_export_sheet_like"
 *
 * Behavior:
 * ✅ Expands/Shrinks rows per item based on number of VALID zones found
 * ✅ Merges Column A (S. No) vertically to match valid zones count
 * ✅ Merges Column B (DOORS OR OPENINGS) vertically to match valid zones count
 * ✅ Writes Zones vertically into target zone column
 * ✅ Writes Nos per zone row
 * ✅ Writes Length(ft) per zone row
 * ✅ SKIPS / REMOVES rows where Nos is blank or 0 BEFORE merge/write
 **************************************************/

const DOORS_SYNC = {
  // ---- MASTER (script hosted here) ----
  MASTER_TAB: "DOOR & WINDOW CALCULATION",
  MASTER_TABLE_HEADER_TEXT: "doors or openings",

  // ✅ PUT YOUR GREEN COLUMN NUMBER HERE (A=1, B=2 ... I=9, J=10)
  TGT_ZONE_COL: 3,

  // ---- MAPPING (Auto-QA Output Template) ----
  MAP_SS_ID: "1wat9koeZC9puQOvH9gC9zxrkJvZxIn5in7EiwlHleus",
  MAP_TAB: "BOQ-LAYER",

  // ---- GENERATED (Auto-QA Output) ----
  GEN_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",
  GEN_TAB: "vis_export_sheet_like",

  // Formatting copy width when inserting rows
  COPY_FORMAT_COLS: 20,

  // Stop after N blank items in master doors list
  STOP_BLANK_RUN: 8,
};

// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu("Vizdom Sync")
//     .addItem("Sync DOORS/OPENINGS → MASTER (Nos + Length + Zones)", "syncDoorsOpeningsToMasterWithZones")
//     .addToUi();
// }

function hideDoorRowsWithBlankOrZeroNos_() {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const masterSh = masterSS.getSheetByName(DOORS_SYNC.MASTER_TAB);
  if (!masterSh) throw new Error(`MASTER tab not found: ${DOORS_SYNC.MASTER_TAB}`);

  const info = findDoorsTable_(masterSh);
  const { headerRow, colItem, colNos } = info;

  const lastRow = masterSh.getLastRow();
  if (lastRow <= headerRow) return;

  const itemVals = masterSh.getRange(headerRow + 1, colItem, lastRow - headerRow, 1).getDisplayValues();
  const nosVals = masterSh.getRange(headerRow + 1, colNos, lastRow - headerRow, 1).getDisplayValues();

  for (let i = 0; i < itemVals.length; i++) {
    const rowNum = headerRow + 1 + i;

    const itemText = String(itemVals[i][0] || "").trim();
    const nosText = String(nosVals[i][0] || "").trim();
    const nosNum = toNumber_(nosText);

    // don't hide completely blank rows or bracket sublabels
    if (!itemText) continue;
    if (/^\(.*\)$/.test(itemText)) continue;

    if (nosText === "" || nosNum === 0) {
      masterSh.hideRows(rowNum);
    } else {
      masterSh.showRows(rowNum);
    }
  }
}
function syncDoorsOpeningsToMasterWithZones() {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const masterSh = masterSS.getSheetByName(DOORS_SYNC.MASTER_TAB);
  if (!masterSh) throw new Error(`MASTER tab not found: ${DOORS_SYNC.MASTER_TAB}`);

  const mapSS = SpreadsheetApp.openById(DOORS_SYNC.MAP_SS_ID);
  const mapSh = mapSS.getSheetByName(DOORS_SYNC.MAP_TAB);
  if (!mapSh) throw new Error(`Mapping tab not found: ${DOORS_SYNC.MAP_TAB}`);

  const genSS = SpreadsheetApp.openById(DOORS_SYNC.GEN_SS_ID);
  const genSh = genSS.getSheetByName(DOORS_SYNC.GEN_TAB);
  if (!genSh) throw new Error(`Generated tab not found: ${DOORS_SYNC.GEN_TAB}`);

  // 1) Build mapping lookup: master item -> {blockName, generatedLayerName, measurementTokens}
  const mapping = buildDoorMapping_(mapSh);

  // 2) Build generated lookup: key -> {zones[], qtyArr[], lenArr[]}
  const genAgg = buildGeneratedAggWithZones_(genSh);

  // 3) Find doors table columns in master
  const masterInfo = findDoorsTable_(masterSh);
  const { headerRow, colSno, colItem, colNos, colLen } = masterInfo;

  // 4) Walk down the doors list dynamically
  let r = headerRow + 1;
  let blankRun = 0;
  let updatedItems = 0;

  while (r <= masterSh.getLastRow()) {
    const itemCell = masterSh.getRange(r, colItem);
    const rawItem = String(itemCell.getDisplayValue() || "").trim();

    if (!rawItem) {
      blankRun++;
      if (blankRun >= DOORS_SYNC.STOP_BLANK_RUN) break;
      r++;
      continue;
    }
    blankRun = 0;

    // Skip sublabels like "(RECEPTION)"
    if (/^\(.*\)$/.test(rawItem)) {
      r++;
      continue;
    }

    // Only process top-left of merged item block
    if (itemCell.isPartOfMerge()) {
      const merged = itemCell.getMergedRanges()[0];
      if (!(merged.getRow() === r && merged.getColumn() === colItem)) {
        r++;
        continue;
      }
    }

    const key = normKey_(rawItem);
    const mapRow = mapping.get(key);
    if (!mapRow) {
      const blk = getItemBlockRange_(masterSh, r, colItem);
      r = blk.startRow + blk.numRows;
      continue;
    }

    const needsCount = mapRow.measurementTokens.has("count");
    const needsLen = mapRow.measurementTokens.has("length");

    // Choose search keys (priority)
    const keysToTry = [];
    if (mapRow.blockName) keysToTry.push(mapRow.blockName);
    if (mapRow.generatedLayerName) keysToTry.push(mapRow.generatedLayerName);
    keysToTry.push(rawItem);

    const hit = findBestGenHit_(genAgg, keysToTry);
    if (!hit) {
      const blk = getItemBlockRange_(masterSh, r, colItem);
      r = blk.startRow + blk.numRows;
      continue;
    }

    // -----------------------------
    // FILTER VALID ROWS FIRST
    // keep only rows where Nos > 0
    // -----------------------------
    const filteredRows = [];
    const rawZones = hit.zones || [];
    const rawQtyArr = hit.qtyArr || [];
    const rawLenArr = hit.lenArr || [];

    for (let z = 0; z < rawZones.length; z++) {
      const zone = String(rawZones[z] || "").trim() || "misc";
      const qty = toNumber_(rawQtyArr[z]) ?? 0;
      const len = toNumber_(rawLenArr[z]);

      if (qty > 0) {
        filteredRows.push({
          zone: zone,
          qty: qty,
          len: len
        });
      }
    }

    // If nothing valid remains, make the block 1 row and clear outputs
    const block = getItemBlockRange_(masterSh, r, colItem);
    const currentRows = block.numRows;

    if (filteredRows.length === 0) {
      // shrink to 1 row if needed
      if (currentRows > 1) {
        breakMergeIfAny_(masterSh, block.startRow, colSno);
        breakMergeIfAny_(masterSh, block.startRow, colItem);

        masterSh.deleteRows(block.startRow + 1, currentRows - 1);

        // keep top row unmerged
        masterSh.getRange(block.startRow, colSno).setVerticalAlignment("middle");
        masterSh.getRange(block.startRow, colItem).setVerticalAlignment("middle");
      }

      // clear target values on the remaining row
      masterSh.getRange(block.startRow, DOORS_SYNC.TGT_ZONE_COL).clearContent();
      if (colNos) masterSh.getRange(block.startRow, colNos).clearContent();
      if (colLen) masterSh.getRange(block.startRow, colLen).clearContent();

      updatedItems++;
      r = block.startRow + 1;
      continue;
    }

    const desiredRows = filteredRows.length;

    // ---------- ADAPT HEIGHT (expand / shrink) ----------
    if (desiredRows > currentRows) {
      const add = desiredRows - currentRows;

      insertRowsWithFormat_(
        masterSh,
        block.startRow + currentRows - 1,
        add,
        DOORS_SYNC.COPY_FORMAT_COLS
      );

      remakeMerge_(masterSh, block.startRow, colSno, currentRows, desiredRows);
      remakeMerge_(masterSh, block.startRow, colItem, currentRows, desiredRows);

    } else if (desiredRows < currentRows) {
      breakMergeIfAny_(masterSh, block.startRow, colSno);
      breakMergeIfAny_(masterSh, block.startRow, colItem);

      const deleteFrom = block.startRow + desiredRows;
      const deleteCount = currentRows - desiredRows;
      masterSh.deleteRows(deleteFrom, deleteCount);

      if (desiredRows > 1) {
        masterSh.getRange(block.startRow, colSno, desiredRows, 1)
          .merge()
          .setVerticalAlignment("middle");
        masterSh.getRange(block.startRow, colItem, desiredRows, 1)
          .merge()
          .setVerticalAlignment("middle");
      } else {
        masterSh.getRange(block.startRow, colSno).setVerticalAlignment("middle");
        masterSh.getRange(block.startRow, colItem).setVerticalAlignment("middle");
      }
    } else {
      // ensure merge/alignment is correct
      if (desiredRows > 1) {
        ensureMergedExactly_(masterSh, block.startRow, colSno, desiredRows);
        ensureMergedExactly_(masterSh, block.startRow, colItem, desiredRows);
      } else {
        breakMergeIfAny_(masterSh, block.startRow, colSno);
        breakMergeIfAny_(masterSh, block.startRow, colItem);
        masterSh.getRange(block.startRow, colSno).setVerticalAlignment("middle");
        masterSh.getRange(block.startRow, colItem).setVerticalAlignment("middle");
      }
    }

    // ---------- Write values ----------
    const zoneWrite = [];
    const nosWrite = [];
    const lenWrite = [];

    for (let i = 0; i < filteredRows.length; i++) {
      zoneWrite.push([filteredRows[i].zone]);
      nosWrite.push([filteredRows[i].qty]);
      lenWrite.push([filteredRows[i].len != null ? filteredRows[i].len : ""]);
    }

    masterSh.getRange(block.startRow, DOORS_SYNC.TGT_ZONE_COL, desiredRows, 1).setValues(zoneWrite);

    if (needsCount) {
      masterSh.getRange(block.startRow, colNos, desiredRows, 1).setValues(nosWrite);
    } else {
      masterSh.getRange(block.startRow, colNos, desiredRows, 1).clearContent();
    }

    if (needsLen) {
      masterSh.getRange(block.startRow, colLen, desiredRows, 1).setValues(lenWrite);
    } else {
      masterSh.getRange(block.startRow, colLen, desiredRows, 1).clearContent();
    }

    updatedItems++;
    r = block.startRow + desiredRows;
  }
  hideDoorRowsWithBlankOrZeroNos_();

  SpreadsheetApp.getActive().toast(
    `Doors Sync complete. Updated ${updatedItems} item(s).`,
    "DOORS → MASTER",
    8
  );


}

/* =========================
   MASTER: find doors table
   ========================= */

function findDoorsTable_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = Math.min(30, sh.getLastColumn());
  if (lastRow < 1) throw new Error("MASTER sheet looks empty.");

  const scanRows = Math.min(lastRow, 2000);
  const scan = sh.getRange(1, 1, scanRows, lastCol).getDisplayValues();

  let headerRow = -1;
  const target = DOORS_SYNC.MASTER_TABLE_HEADER_TEXT;

  for (let r = 0; r < scan.length; r++) {
    const row = scan[r].map(v => String(v || "").trim().toLowerCase());
    if (row.includes(target)) {
      headerRow = r + 1;
      break;
    }
  }
  if (headerRow === -1) throw new Error(`Could not find "${target}" in MASTER.`);

  const header = sh.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || "").trim().toLowerCase());

  const colSno = header.findIndex(v => v.replace(/\./g, "") === "s no" || v === "s. no") + 1;
  const colItem = header.findIndex(v => v === "doors or openings") + 1;
  const colNos = header.findIndex(v => v === "nos") + 1;
  const colLen = header.findIndex(v => v.includes("length")) + 1;

  if (!colSno || !colItem || !colNos || !colLen) {
    throw new Error(`MASTER doors headers not detected. Need: S. No, DOORS OR OPENINGS, Nos, LENGTH (Ft)`);
  }

  return { headerRow, colSno, colItem, colNos, colLen };
}

/* =========================
   MAPPING: BOQ-LAYER
   ========================= */

function buildDoorMapping_(mapSh) {
  const lastRow = mapSh.getLastRow();
  const lastCol = mapSh.getLastColumn();
  if (lastRow < 2) return new Map();

  const headerScanRows = Math.min(10, lastRow);
  const scan = mapSh.getRange(1, 1, headerScanRows, lastCol).getDisplayValues();

  let headerRowIdx = -1;
  for (let r = 0; r < scan.length; r++) {
    const row = scan[r].map(v => String(v || "").trim().toLowerCase());
    if (row.includes("boq name") && row.includes("measurement")) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx === -1) throw new Error(`Mapping: couldn't find header row containing "BOQ Name" + "Measurement"`);

  const hdr = scan[headerRowIdx].map(v => String(v || "").trim().toLowerCase());

  const idxTarget = hdr.indexOf("boq name");
  const idxMeas = hdr.indexOf("measurement");

  const layerNameIdxs = [];
  for (let i = 0; i < hdr.length; i++) {
    if (hdr[i] === "layer name") layerNameIdxs.push(i);
  }
  const idxGeneratedLayer = layerNameIdxs.length >= 2 ? layerNameIdxs[1] : -1;

  const idxBlock = hdr.indexOf("block name");

  if (idxTarget === -1) throw new Error(`Mapping: missing "BOQ Name" header`);
  if (idxBlock === -1) throw new Error(`Mapping: missing "Block Name" header`);
  if (idxMeas === -1) throw new Error(`Mapping: missing "Measurement" header`);

  const dataStart = headerRowIdx + 2;
  const numRows = lastRow - (dataStart - 1);
  if (numRows <= 0) return new Map();

  const values = mapSh.getRange(dataStart, 1, numRows, lastCol).getDisplayValues();
  const map = new Map();

  for (let r = 0; r < values.length; r++) {
    const targeted = String(values[r][idxTarget] || "").trim();
    if (!targeted) continue;

    const blockName = String(values[r][idxBlock] || "").trim();
    const generatedLayerName =
      idxGeneratedLayer !== -1 ? String(values[r][idxGeneratedLayer] || "").trim() : "";

    const measRaw = String(values[r][idxMeas] || "").trim();
    const measurementTokens = new Set(
      measRaw.toLowerCase().split(/[^a-z]+/g).map(s => s.trim()).filter(Boolean)
    );

    map.set(normKey_(targeted), {
      targeted,
      blockName,
      generatedLayerName,
      measurementTokens,
    });
  }

  return map;
}

/* =========================
   GENERATED: vis_export_sheet_like
   ========================= */

function buildGeneratedAggWithZones_(genSh) {
  const lastRow = genSh.getLastRow();
  const lastCol = genSh.getLastColumn();
  if (lastRow < 2) return { byProduct: new Map(), byBoq: new Map() };

  const values = genSh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const hdr = values[0].map(v => String(v || "").trim().toLowerCase());

  const idxProduct = hdr.indexOf("product");
  const idxBoq = hdr.indexOf("boq name");
  const idxZone = hdr.indexOf("zone");
  const idxQty = hdr.indexOf("qty_value");
  const idxLen = hdr.indexOf("length (ft)");

  if (idxBoq === -1) throw new Error(`Generated: missing "BOQ name" header`);
  if (idxZone === -1) throw new Error(`Generated: missing "zone" header`);
  if (idxQty === -1) throw new Error(`Generated: missing "qty_value" header`);
  if (idxLen === -1) throw new Error(`Generated: missing "length (ft)" header`);

  let currentBoq = "";
  let currentProduct = "";

  const byBoq = new Map();
  const byProduct = new Map();

  function push_(map, key, zone, qty, len) {
    const k = normKey_(key);
    if (!k) return;

    let agg = map.get(k);
    if (!agg) {
      agg = { zoneQty: new Map(), zoneLen: new Map(), zones: [] };
      map.set(k, agg);
    }

    if (!agg.zoneQty.has(zone)) agg.zones.push(zone);

    const prevQ = agg.zoneQty.get(zone) || 0;
    agg.zoneQty.set(zone, prevQ + (qty > 0 ? qty : 0));

    if (!agg.zoneLen.has(zone) && len > 0) {
      agg.zoneLen.set(zone, len);
    }
  }

  for (let r = 1; r < values.length; r++) {
    const boqCell = String(values[r][idxBoq] || "").trim();
    if (boqCell) currentBoq = boqCell;

    if (idxProduct !== -1) {
      const prodCell = String(values[r][idxProduct] || "").trim();
      if (prodCell) currentProduct = prodCell;
    }

    if (!currentBoq && !currentProduct) continue;

    const zone = String(values[r][idxZone] || "").trim() || "misc";
    const qty = toNumber_(values[r][idxQty]) ?? 0;
    const len = toNumber_(values[r][idxLen]) ?? 0;

    if (currentBoq) push_(byBoq, currentBoq, zone, qty, len);
    if (currentProduct) push_(byProduct, currentProduct, zone, qty, len);
  }

  function finalize(map) {
    const out = new Map();
    for (const [k, agg] of map.entries()) {
      const zones = uniqKeepOrder_(agg.zones);
      const qtyArr = zones.map(z => agg.zoneQty.get(z) || 0);
      const lenArr = zones.map(z => agg.zoneLen.get(z) || null);
      out.set(k, { zones, qtyArr, lenArr });
    }
    return out;
  }

  return { byProduct: finalize(byProduct), byBoq: finalize(byBoq) };
}

function findBestGenHit_(genAgg, keysToTry) {
  for (const k of keysToTry) {
    const key = normKey_(k);
    if (!key) continue;

    if (genAgg.byBoq.has(key)) return genAgg.byBoq.get(key);
    if (genAgg.byProduct.has(key)) return genAgg.byProduct.get(key);
  }
  return null;
}

function uniqKeepOrder_(arr) {
  const seen = new Set();
  const out = [];
  for (const x of arr || []) {
    const v = String(x || "").trim();
    if (!v) continue;
    if (seen.has(v)) continue;
    seen.add(v);
    out.push(v);
  }
  return out;
}

/* =========================
   Block helpers
   ========================= */

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

function ensureMergedExactly_(sheet, startRow, col, numRows) {
  const rng = sheet.getRange(startRow, col, numRows, 1);
  if (rng.isPartOfMerge()) {
    const m = rng.getMergedRanges()[0];
    if (m.getRow() === startRow && m.getColumn() === col && m.getNumRows() === numRows) {
      m.setVerticalAlignment("middle");
      return;
    }
    m.breakApart();
  }

  if (numRows > 1) {
    rng.merge().setVerticalAlignment("middle");
  } else {
    sheet.getRange(startRow, col).setVerticalAlignment("middle");
  }
}

/* =========================
   Utils
   ========================= */

function normKey_(s) {
  return String(s || "")
    .trim()
    .toLowerCase()
    .replace(/[\u00A0]/g, " ")
    .replace(/[^\w\s-]/g, " ")
    .replace(/[_-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function toNumber_(v) {
  if (v == null || v === "") return null;
  const n = Number(String(v).replace(/,/g, "").trim());
  return Number.isFinite(n) ? n : null;
}

function roundSmart_(n) {
  if (n == null) return null;
  if (Math.abs(n - Math.round(n)) < 1e-9) return Math.round(n);
  return Math.round(n * 1000) / 1000;
}