/**************************************************
 * PLANNER → CALCULATION SHEET (Rooms + Carpet Area)
 * - Source can be SAME spreadsheet or DIFFERENT via SOURCE_SS_ID
 * - Auto-finds source sheet by headers (name + area_sqft)
 * - Auto adds/removes rows BEFORE TOTAL
 * - Preserves TOTAL row + its formula
 * - Copies formatting from template row (full width)
 * - ✅ TOTAL row detection works even if TOTAL is merged (A:B etc.)
 **************************************************/

const PLANNER_SYNC = {
  // ✅ If PLANNER is in another spreadsheet, paste that ID here.
  // If left "", source = ACTIVE spreadsheet.
  SOURCE_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",

  SOURCE_TAB: "PLANNER", // preferred source tab name (optional)
  TARGET_TAB: "CALCULATION SHEET",

  TOTAL_TEXT: "total",

  // how many columns to scan in each row to locate TOTAL
  // (TOTAL is typically in merged A:B, and carpet is in C)
  TOTAL_SCAN_COLS: 6, // scans A..F for the word "TOTAL"
};

function onOpen() { SpreadsheetApp.getUi() .createMenu("Vizdom Sync") .addItem("Sync PLANNER → CALCULATION SHEET", "syncPlannerToMaster") .addToUi(); }

function syncPlannerToMaster() {
  const targetSS = SpreadsheetApp.getActiveSpreadsheet();

  // ✅ Source spreadsheet: either same as target or openById
  const sourceSS =
    (PLANNER_SYNC.SOURCE_SS_ID || "").trim()
      ? SpreadsheetApp.openById(PLANNER_SYNC.SOURCE_SS_ID.trim())
      : targetSS;

  const src = findPlannerSourceSheet_(sourceSS, PLANNER_SYNC.SOURCE_TAB);
  const tgt = targetSS.getSheetByName(PLANNER_SYNC.TARGET_TAB);
  if (!tgt) throw new Error(`Target tab not found: ${PLANNER_SYNC.TARGET_TAB}`);

  // Read planner data (name + area_sqft)
  const planner = readPlannerRows_(src); // [{name, area}]
  if (!planner.length) {
    SpreadsheetApp.getActive().toast("No planner rows found.", "PLANNER Sync", 6);
    return;
  }

  // Locate header row & columns + TOTAL row in target (merge-safe)
  const targetInfo = findRoomsBlock_(tgt);
  let {
    dataStartRow,
    totalRow,
    colSno,
    colRooms,
    colCarpet,
  } = targetInfo;

  // Template row = first data row style
  const templateRow = dataStartRow;
  if (templateRow >= totalRow) throw new Error("Template row invalid (no data row before TOTAL).");

  // How many rows exist currently between headers and TOTAL?
  const existingDataCount = Math.max(0, totalRow - dataStartRow);

  // Need this many rows:
  const neededDataCount = planner.length;
  const delta = neededDataCount - existingDataCount;

  // Width to format-copy (whole sheet width so blue area matches too)
  const lastCol = tgt.getLastColumn();

  if (delta > 0) {
    // Insert rows before TOTAL
    tgt.insertRowsBefore(totalRow, delta);

    // Copy formatting from template row into newly inserted rows
    const srcFmt = tgt.getRange(templateRow, 1, 1, lastCol);
    const dstFmt = tgt.getRange(templateRow + existingDataCount, 1, delta, lastCol);
    srcFmt.copyTo(dstFmt, { formatOnly: true });

  } else if (delta < 0) {
    // Delete extra rows (keep TOTAL)
    tgt.deleteRows(dataStartRow + neededDataCount, -delta);
  }

  // Re-find positions (TOTAL row shifts after insert/delete)
  const info2 = findRoomsBlock_(tgt);
  dataStartRow = info2.dataStartRow;
  totalRow = info2.totalRow;
  colSno = info2.colSno;
  colRooms = info2.colRooms;
  colCarpet = info2.colCarpet;

  // Write rows
  const sNoVals = planner.map((_, i) => [i + 1]);
  const roomVals = planner.map(p => [p.name]);
  const areaVals = planner.map(p => [p.area]);

  tgt.getRange(dataStartRow, colSno, planner.length, 1).setValues(sNoVals);
  tgt.getRange(dataStartRow, colRooms, planner.length, 1).setValues(roomVals);
  tgt.getRange(dataStartRow, colCarpet, planner.length, 1).setValues(areaVals);

  // TOTAL formula in carpet col:
  // Keep existing formula if present. If blank, set SUM over the new range.
  const totalCell = tgt.getRange(totalRow, colCarpet);
  // Always rebuild TOTAL formula safely
const startA1 = a1_(dataStartRow, colCarpet);
const endA1 = a1_(dataStartRow + planner.length - 1, colCarpet);
totalCell.setFormula(`=SUM(${startA1}:${endA1})`);

  SpreadsheetApp.getActive().toast(
    `Copied ${planner.length} room(s) from ${src.getName()} → ${tgt.getName()}`,
    "PLANNER Sync",
    8
  );
}

/* -------------------- SOURCE -------------------- */

function findPlannerSourceSheet_(ss, preferredName) {
  // 1) try preferred name
  if (preferredName) {
    const sh = ss.getSheetByName(preferredName);
    if (sh) return sh;
  }

  // 2) auto-find by headers: name + area_sqft
  for (const sh of ss.getSheets()) {
    const lastCol = sh.getLastColumn();
    const lastRow = sh.getLastRow();
    if (lastRow < 1 || lastCol < 1) continue;

    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(v => String(v || "").trim().toLowerCase());

    if (header.includes("name") && header.includes("area_sqft")) return sh;
  }

  throw new Error(
    `Source tab not found. Tried "${preferredName}" and could not auto-detect (need headers: name, area_sqft).`
  );
}

function readPlannerRows_(srcSheet) {
  const lastRow = srcSheet.getLastRow();
  const lastCol = srcSheet.getLastColumn();
  if (lastRow < 2) return [];

  const values = srcSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const header = values[0].map(v => String(v || "").trim().toLowerCase());

  const idxName = header.indexOf("name");
  const idxArea = header.indexOf("area_sqft");

  if (idxName === -1 || idxArea === -1) {
    throw new Error(`Source sheet "${srcSheet.getName()}" must contain headers: name, area_sqft`);
  }

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][idxName] || "").trim();
    if (!name) continue;

    const area = toNumber_(values[r][idxArea]);
    out.push({ name, area: area ?? "" });
  }
  return out;
}

/* -------------------- TARGET (merge-safe TOTAL) -------------------- */

function findRoomsBlock_(tgtSheet) {
  const lastRow = tgtSheet.getLastRow();
  const lastCol = tgtSheet.getLastColumn();

  // Search top area for header row: S. No | ROOMS | Carpet Area
  const scanRows = Math.min(lastRow, 60);
  const scanCols = Math.min(lastCol, 20);
  const scan = tgtSheet.getRange(1, 1, scanRows, scanCols).getValues();

  let headerRow = -1;
  let colSno = -1, colRooms = -1, colCarpet = -1;

  for (let r = 0; r < scan.length; r++) {
    const row = scan[r].map(v => String(v || "").trim().toLowerCase());

    const snoIdx = row.findIndex(v => v.replace(/\./g, "") === "s no" || v === "s. no");
    const roomsIdx = row.findIndex(v => v === "rooms");
    const carpetIdx = row.findIndex(v => v === "carpet area");

    if (snoIdx !== -1 && roomsIdx !== -1 && carpetIdx !== -1) {
      headerRow = r + 1;
      colSno = snoIdx + 1;
      colRooms = roomsIdx + 1;
      colCarpet = carpetIdx + 1;
      break;
    }
  }

  if (headerRow === -1) throw new Error("Header row not found (need: S. No, ROOMS, Carpet Area).");

  const dataStartRow = headerRow + 1;

  // ✅ Find TOTAL row by scanning A..F (or configured) for the word "TOTAL"
  const scanToCol = Math.min(lastCol, Math.max(PLANNER_SYNC.TOTAL_SCAN_COLS, colCarpet));
  const rowsBelow = Math.max(1, lastRow - headerRow);
  const block = tgtSheet.getRange(headerRow + 1, 1, rowsBelow, scanToCol).getValues();

  let totalRow = -1;
  for (let i = 0; i < block.length; i++) {
    const rowText = block[i]
      .map(v => String(v || "").trim().toLowerCase())
      .filter(Boolean);

    if (rowText.some(t => t === PLANNER_SYNC.TOTAL_TEXT)) {
      totalRow = (headerRow + 1) + i;
      break;
    }
  }

  if (totalRow === -1) throw new Error("TOTAL row not found.");

  return { headerRow, dataStartRow, totalRow, colSno, colRooms, colCarpet };
}

/* -------------------- UTILS -------------------- */

function toNumber_(v) {
  if (v == null || v === "") return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function a1_(row, col) {
  return colToA1_(col) + row;
}
function colToA1_(n) {
  let s = "";
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = ((n - 1) / 26) >> 0;
  }
  return s;
}