/**************************************************
 * DOORS / OPENINGS → MASTER (Nos + Length from Export via Mapping)
 *
 * ✅ Master sheet: "CALCULATION SHEET"
 *    Uses the table that starts with header "DOORS OR OPENINGS"
 *    Fills:
 *      - "Nos" column (COUNT)
 *      - "LENGTH (Ft)" column (if Export length differs / master is blank)
 *
 * ✅ Mapping sheet (Auto-QA Output Template): tab "BOQ-LAYER"
 *    Matches Master item name (Targeted BOQ Name) → picks:
 *      - Generated Layer Name (optional)
 *      - Generated-Block Block Name (used as primary search key)
 *      - Measurement (Count / Length)
 *
 * ✅ Generated sheet (Auto-QA Output): tab "vis_export_sheet_like"
 *    Searches by:
 *      1) Product  (usually block / product name)
 *      2) BOQ name (fallback)
 *    Aggregates:
 *      - qty_value SUM (for Count)
 *      - length (ft) representative (mode, else first non-zero)
 *
 * ------------------------------------------------
 * SET THESE 3 CONSTANTS BEFORE RUNNING
 **************************************************/
const DOORS_SYNC = {
  // MASTER (script is hosted in MASTER spreadsheet)
  MASTER_TAB: "CALCULATION SHEET",

  // MAPPING (Auto-QA Output Template)
  MAP_SS_ID: "1wat9koeZC9puQOvH9gC9zxrkJvZxIn5in7EiwlHleus", // <-- e.g. 1wat9koeZC9puQOvH9c9zrxlvZxln5in7EiwlHleus
  MAP_TAB: "BOQ-LAYER",

  // GENERATED (Auto-QA Output)
  GEN_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",
  GEN_TAB: "vis_export_sheet_like",

  // Master table header text to locate the doors table
  MASTER_TABLE_HEADER_TEXT: "doors or openings",

  // How far down to scan for the table in master
  MASTER_SCAN_ROWS: 2000,
  MASTER_SCAN_COLS: 20,
};

// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu("Vizdom Sync")
//     .addItem("Sync DOORS/OPENINGS → MASTER (Nos + Length)", "syncDoorsOpeningsToMaster")
//     .addToUi();
// }

function syncDoorsOpeningsToMaster() {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const masterSh = masterSS.getSheetByName(DOORS_SYNC.MASTER_TAB);
  if (!masterSh) throw new Error(`MASTER tab not found: ${DOORS_SYNC.MASTER_TAB}`);

  const mapSS = SpreadsheetApp.openById(DOORS_SYNC.MAP_SS_ID);
  const mapSh = mapSS.getSheetByName(DOORS_SYNC.MAP_TAB);
  if (!mapSh) throw new Error(`Mapping tab not found: ${DOORS_SYNC.MAP_TAB}`);

  const genSS = SpreadsheetApp.openById(DOORS_SYNC.GEN_SS_ID);
  const genSh = genSS.getSheetByName(DOORS_SYNC.GEN_TAB);
  if (!genSh) throw new Error(`Generated tab not found: ${DOORS_SYNC.GEN_TAB}`);

  // 1) Build mapping: Targeted BOQ Name -> { blockName, genLayerName, measurement }
  const mapping = buildDoorMapping_(mapSh);

  // 2) Build generated aggregates: by Product and by BOQ name
  const genAgg = buildGeneratedAgg_(genSh);

  // 3) Locate doors table in master + get rows
  const masterInfo = findDoorsTable_(masterSh);
  const { headerRow, colItem, colNos, colLen } = masterInfo;

  const rows = readDoorItems_(masterSh, headerRow, colItem, DOORS_SYNC.MASTER_SCAN_ROWS);
  if (!rows.length) {
    SpreadsheetApp.getActive().toast("No DOORS/OPENINGS items found under table.", "Doors Sync", 6);
    return;
  }

  let updated = 0;
  let updatedNos = 0;
  let updatedLen = 0;

  // Batch read current Nos/Len (faster + fewer calls)
  const itemRange = masterSh.getRange(rows[0].row, colItem, rows.length, 1).getDisplayValues();
  const nosRange  = masterSh.getRange(rows[0].row, colNos,  rows.length, 1).getValues();
  const lenRange  = masterSh.getRange(rows[0].row, colLen,  rows.length, 1).getValues();

  const outNos = [];
  const outLen = [];

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i].row;
    const rawName = String(itemRange[i][0] || "").trim();
    const key = normKey_(rawName);

    let newNos = nosRange[i][0];
    let newLen = lenRange[i][0];

    const m = mapping.get(key);
    if (!m) {
      // no mapping found → keep as is
      outNos.push([newNos]);
      outLen.push([newLen]);
      continue;
    }

    // Measurement requirement (like: "Count", "Length", "Count Length")
    const needsCount = m.measurementTokens.has("count");
    const needsLen   = m.measurementTokens.has("length");

    // Search in generated using prioritized keys:
    // 1) blockName (strongest)
    // 2) generatedLayerName (optional)
    // 3) original master name
    const keysToTry = [];
    if (m.blockName) keysToTry.push(m.blockName);
    if (m.generatedLayerName) keysToTry.push(m.generatedLayerName);
    keysToTry.push(rawName);

    const hit = findBestGenHit_(genAgg, keysToTry);

    // If we found something in generated sheet, apply per measurement requirement
    if (hit) {
      if (needsCount) {
        const qty = hit.qtySum;
        if (qty !== null && qty !== "" && qty !== undefined) {
          if (String(newNos) !== String(qty)) {
            newNos = qty;
            updatedNos++;
            updated++;
          }
        }
      }

      if (needsLen) {
        const gotLen = hit.lenValue;
        if (gotLen != null && gotLen !== "" && Number(gotLen) > 0) {
          const masterLenNum = toNumber_(newLen);
          const gotLenNum = toNumber_(gotLen);

          // Update if master length is blank/0 OR differs (tolerance 0.01 ft)
          const shouldUpdate =
            masterLenNum == null ||
            masterLenNum === 0 ||
            Math.abs(masterLenNum - gotLenNum) > 0.01;

          if (shouldUpdate) {
            newLen = gotLenNum;
            updatedLen++;
            updated++;
          }
        }
      }
    }

    outNos.push([newNos]);
    outLen.push([newLen]);
  }

  // Write back in 2 setValues calls
  masterSh.getRange(rows[0].row, colNos, rows.length, 1).setValues(outNos);
  masterSh.getRange(rows[0].row, colLen, rows.length, 1).setValues(outLen);

  SpreadsheetApp.getActive().toast(
    `Done. Updated rows: ${updated} (Nos: ${updatedNos}, Length: ${updatedLen})`,
    "Doors Sync",
    8
  );
}

/* =========================
   MASTER: find doors table
   ========================= */

function findDoorsTable_(sh) {
  const lastRow = Math.min(DOORS_SYNC.MASTER_SCAN_ROWS, sh.getLastRow());
  const lastCol = Math.min(DOORS_SYNC.MASTER_SCAN_COLS, sh.getLastColumn());
  if (lastRow < 1 || lastCol < 1) throw new Error("MASTER sheet looks empty.");

  const scan = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const target = DOORS_SYNC.MASTER_TABLE_HEADER_TEXT;

  // Find header row containing "DOORS OR OPENINGS"
  let headerRow = -1;
  for (let r = 0; r < scan.length; r++) {
    const row = scan[r].map(v => String(v || "").trim().toLowerCase());
    if (row.some(v => v === target)) {
      headerRow = r + 1;
      break;
    }
  }
  if (headerRow === -1) throw new Error(`Could not find "${target}" in MASTER tab.`);

  // The header row is the row that also has columns like: S. No | DOORS OR OPENINGS | Nos | LENGTH (Ft)
  const header = sh.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0]
    .map(v => String(v || "").trim().toLowerCase());

  const colItem = header.findIndex(v => v === "doors or openings") + 1;
  const colNos  = header.findIndex(v => v === "nos") + 1;
  const colLen  = header.findIndex(v => v.includes("length")) + 1;

  if (!colItem || !colNos || !colLen) {
    throw new Error(
      `MASTER doors table columns not detected properly. Need headers: "DOORS OR OPENINGS", "Nos", "LENGTH (Ft)".`
    );
  }

  return { headerRow, colItem, colNos, colLen };
}

function readDoorItems_(sh, headerRow, colItem, maxScanRows) {
  // Read below header until we hit a long blank run
  const startRow = headerRow + 1;
  const lastRow = Math.min(maxScanRows, sh.getLastRow());
  if (lastRow < startRow) return [];

  const vals = sh.getRange(startRow, colItem, lastRow - startRow + 1, 1).getDisplayValues();

  const rows = [];
  let blankRun = 0;

  for (let i = 0; i < vals.length; i++) {
    const name = String(vals[i][0] || "").trim();
    const r = startRow + i;

    if (!name) {
      blankRun++;
      if (blankRun >= 8) break; // stop after 8 consecutive blanks
      continue;
    }

    blankRun = 0;

    // Skip sublabels like "(RECEPTION)" etc — you can remove this if you want them included
    if (/^\(.*\)$/.test(name)) continue;

    rows.push({ row: r });
  }

  return rows;
}

/* =========================
   MAPPING: BOQ-LAYER tab
   ========================= */

function buildDoorMapping_(mapSh) {
  const lastRow = mapSh.getLastRow();
  const lastCol = mapSh.getLastColumn();
  if (lastRow < 2) return new Map();

  // ✅ Find the real header row (because your sheet has multi-row headers)
  const headerScanRows = Math.min(10, lastRow);
  const scan = mapSh.getRange(1, 1, headerScanRows, lastCol).getDisplayValues();

  let headerRowIdx = -1; // 0-based
  for (let r = 0; r < scan.length; r++) {
    const row = scan[r].map(v => String(v || "").trim().toLowerCase());
    if (row.includes("boq name") && row.includes("measurement")) {
      headerRowIdx = r;
      break;
    }
  }
  if (headerRowIdx === -1) {
    throw new Error(`Mapping: couldn't find header row containing "BOQ Name" + "Measurement" in ${DOORS_SYNC.MAP_TAB}`);
  }

  const hdr = scan[headerRowIdx].map(v => String(v || "").trim().toLowerCase());

  // Columns
  const idxTarget = hdr.indexOf("boq name");        // Targeted BOQ Name
  const idxMeas   = hdr.indexOf("measurement");    // Measurement

  // There are 2 "Layer Name" columns; we want the one under "Generated"
  const layerNameIdxs = [];
  for (let i = 0; i < hdr.length; i++) if (hdr[i] === "layer name") layerNameIdxs.push(i);
  const idxGeneratedLayer = layerNameIdxs.length >= 2 ? layerNameIdxs[1] : -1;

  // "Block Name" under "Generated-Block"
  const idxBlock = hdr.indexOf("block name");

  if (idxTarget === -1) throw new Error(`Mapping: missing "BOQ Name" header in ${DOORS_SYNC.MAP_TAB}`);
  if (idxBlock === -1)  throw new Error(`Mapping: missing "Block Name" header in ${DOORS_SYNC.MAP_TAB}`);
  if (idxMeas === -1)   throw new Error(`Mapping: missing "Measurement" header in ${DOORS_SYNC.MAP_TAB}`);

  // ✅ Data starts AFTER the header row we found
  const dataStart = headerRowIdx + 2; // convert 0-based to sheet row + next row
  const values = mapSh.getRange(dataStart, 1, lastRow - (dataStart - 1), lastCol).getDisplayValues();

  const map = new Map();

  for (let r = 0; r < values.length; r++) {
    const targeted = String(values[r][idxTarget] || "").trim();
    if (!targeted) continue;

    const blockName = String(values[r][idxBlock] || "").trim();
    const generatedLayerName =
      idxGeneratedLayer !== -1 ? String(values[r][idxGeneratedLayer] || "").trim() : "";

    const measRaw = String(values[r][idxMeas] || "").trim();
    const measurementTokens = new Set(
      measRaw
        .toLowerCase()
        .split(/[^a-z]+/g)
        .map(s => s.trim())
        .filter(Boolean)
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

function buildGeneratedAgg_(genSh) {
  const lastRow = genSh.getLastRow();
  const lastCol = genSh.getLastColumn();
  if (lastRow < 2) return { byProduct: new Map(), byBoq: new Map() };

  const values = genSh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const hdr = values[0].map(v => String(v || "").trim().toLowerCase());

  const idxProduct = hdr.indexOf("product");
  const idxBoq = hdr.indexOf("boq name");
  const idxQty = hdr.indexOf("qty_value");
  const idxLen = hdr.indexOf("length (ft)");

  if (idxProduct === -1) throw new Error(`Generated: missing header "Product" in ${DOORS_SYNC.GEN_TAB}`);
  if (idxBoq === -1) throw new Error(`Generated: missing header "BOQ name" in ${DOORS_SYNC.GEN_TAB}`);
  if (idxQty === -1) throw new Error(`Generated: missing header "qty_value" in ${DOORS_SYNC.GEN_TAB}`);
  if (idxLen === -1) throw new Error(`Generated: missing header "length (ft)" in ${DOORS_SYNC.GEN_TAB}`);

  const byProduct = new Map(); // key -> { qtySum, lenValues[] }
  const byBoq = new Map();

  for (let r = 1; r < values.length; r++) {
    const product = String(values[r][idxProduct] || "").trim();
    const boq = String(values[r][idxBoq] || "").trim();
    const qty = toNumber_(values[r][idxQty]) ?? 0;
    const len = toNumber_(values[r][idxLen]) ?? 0;

    if (product) {
      const k = normKey_(product);
      const agg = byProduct.get(k) || { qtySum: 0, lenValues: [] };
      agg.qtySum += qty;
      if (len > 0) agg.lenValues.push(len);
      byProduct.set(k, agg);
    }

    if (boq) {
      const k = normKey_(boq);
      const agg = byBoq.get(k) || { qtySum: 0, lenValues: [] };
      agg.qtySum += qty;
      if (len > 0) agg.lenValues.push(len);
      byBoq.set(k, agg);
    }
  }

  // Convert lenValues[] -> representative lenValue (mode if possible else first)
  function finalize(map) {
    const out = new Map();
    for (const [k, v] of map.entries()) {
      out.set(k, {
        qtySum: roundSmart_(v.qtySum),
        lenValue: pickLenRepresentative_(v.lenValues),
      });
    }
    return out;
  }

  return { byProduct: finalize(byProduct), byBoq: finalize(byBoq) };
}

function findBestGenHit_(genAgg, keysToTry) {
  for (const k of keysToTry) {
    const key = normKey_(k);
    if (!key) continue;

    if (genAgg.byProduct.has(key)) return genAgg.byProduct.get(key);
    if (genAgg.byBoq.has(key)) return genAgg.byBoq.get(key);
  }
  return null;
}

function pickLenRepresentative_(arr) {
  if (!arr || !arr.length) return null;

  // Mode with tolerance bucket (0.01 ft)
  const buckets = new Map();
  for (const x of arr) {
    const b = (Math.round(x * 100) / 100).toFixed(2);
    buckets.set(b, (buckets.get(b) || 0) + 1);
  }

  let best = null;
  let bestCount = -1;
  for (const [b, c] of buckets.entries()) {
    if (c > bestCount) {
      bestCount = c;
      best = b;
    }
  }
  const bestNum = Number(best);
  return Number.isFinite(bestNum) ? bestNum : arr[0];
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
  // keep integers as integers
  if (Math.abs(n - Math.round(n)) < 1e-9) return Math.round(n);
  return Math.round(n * 1000) / 1000;
}