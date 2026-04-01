/**************************************************
 * MASTER SCRIPT (hosted in MASTER spreadsheet)
 * Menu:
 * 1) Sync PLANNER → CALCULATION SHEET
 * 2) Sync LAYER  → MASTER (Zones + Area)
 *
 * NOTE:
 * - Keep ONLY one onOpen() in entire project (this one).
 * - Shared utilities live here (toNumber_, a1_, colToA1_).
 **************************************************/

/**************************************************
 * PLANNER → CALCULATION SHEET (Rooms + Carpet Area)
 *
 * MODES:
 * - dynamic = old behavior
 *   -> target fully follows planner, inserts/deletes rows
 *
 * - fixed = keep existing target room rows as-is
 *   -> only fill matching Carpet Area values
 *
 * - hybrid = requested behavior
 *   -> all non-fixed planner rooms first
 *   -> fixed rooms always present at bottom in fixed order
 *   -> if fixed room matches planner, fill its area
 *   -> if not, keep it present anyway
 *   -> fixed rooms forced to stay RED
 *   -> planner room text color also transferred
 **************************************************/

const PLANNER_SYNC = {
  // If PLANNER is in another spreadsheet, paste that ID here.
  // If left "", source = ACTIVE spreadsheet (Master).
  SOURCE_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",

  SOURCE_TAB: "PLANNER", // preferred source tab name (optional)
  TARGET_TAB: "CARPET AREA",

  TOTAL_TEXT: "total",

  // how many columns to scan in each row to locate TOTAL
  TOTAL_SCAN_COLS: 6, // scans A..F for word "TOTAL"

  // "dynamic" | "fixed" | "hybrid"
  MODE: "hybrid",

  // These rooms must always remain present in hybrid mode
  FIXED_ROOMS: [
    "ELECTRICAL ROOM",
    "SERVER ROOM",
    "MALE WASHROOM",
    "FEMALE WASHROOM"
  ],

  // Fixed rooms should remain red like your template
  FIXED_ROOM_FONT_COLOR: "#ff0000"
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Vizdom Sync")
    .addItem("CARPET AREA 1", "syncPlannerToMaster")
    .addSeparator()
    .addItem("SITE RECEE CALCULATION", "syncLayerToMasterZonesArea")
    .addSeparator()
    .addItem("DOOR & WINDOW CALCULATION", "syncDoorsOpeningsToMasterWithZones")
    .addSeparator()
    .addItem("CARPET AREA 2", "syncFloorClToCalculationSheet")
    .addToUi();
}

function syncPlannerToMaster() {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();

  const sourceSS =
    (PLANNER_SYNC.SOURCE_SS_ID || "").trim()
      ? SpreadsheetApp.openById(PLANNER_SYNC.SOURCE_SS_ID.trim())
      : masterSS;

  const src = findPlannerSourceSheet_(sourceSS, PLANNER_SYNC.SOURCE_TAB);
  const tgt = masterSS.getSheetByName(PLANNER_SYNC.TARGET_TAB);
  if (!tgt) throw new Error(`Target tab not found in MASTER: ${PLANNER_SYNC.TARGET_TAB}`);

  const planner = readPlannerRows_(src); // [{name, area, color}]
  if (!planner.length) {
    SpreadsheetApp.getActive().toast("No planner rows found.", "PLANNER Sync", 6);
    return;
  }

  const mode = String(PLANNER_SYNC.MODE || "dynamic").toLowerCase();

  /**************************************************
   * HYBRID MODE
   * - non-fixed planner rooms first
   * - fixed rooms always at bottom in fixed order
   * - fixed rooms filled if matched, otherwise kept blank
   * - fixed rooms forced red
   * - planner room font colors transferred
   **************************************************/
  if (mode === "hybrid") {
    const fixedRooms = Array.isArray(PLANNER_SYNC.FIXED_ROOMS)
      ? PLANNER_SYNC.FIXED_ROOMS
      : [];

    const fixedKeySet = new Set(fixedRooms.map(normalizeRoomName_));

    // Build planner map with area + color
    const plannerMap = {};
    planner.forEach((p) => {
      const key = normalizeRoomName_(p.name);
      if (key) {
        plannerMap[key] = {
          area: p.area,
          color: p.color || "#000000"
        };
      }
    });

    // Non-fixed rooms come first and retain planner color
    const normalPlannerRows = planner
      .filter((p) => {
        const key = normalizeRoomName_(p.name);
        return key && !fixedKeySet.has(key);
      })
      .map((p) => ({
        name: p.name,
        area: p.area,
        color: p.color || "#000000"
      }));

    // Fixed rooms always come at bottom and stay red
    const fixedRows = fixedRooms.map((roomName) => {
      const key = normalizeRoomName_(roomName);
      return {
        name: roomName,
        area: Object.prototype.hasOwnProperty.call(plannerMap, key)
          ? plannerMap[key].area
          : "",
        color: PLANNER_SYNC.FIXED_ROOM_FONT_COLOR || "#ff0000"
      };
    });

    const finalRows = [...normalPlannerRows, ...fixedRows];

    const targetInfo = findRoomsBlock_(tgt);
    let { dataStartRow, totalRow, colSno, colRooms, colCarpet } = targetInfo;

    const templateRow = dataStartRow;
    if (templateRow >= totalRow) {
      throw new Error("Template row invalid (no data row before TOTAL).");
    }

    const existingDataCount = Math.max(0, totalRow - dataStartRow);
    const neededDataCount = finalRows.length;
    const delta = neededDataCount - existingDataCount;
    const lastCol = tgt.getLastColumn();

    if (delta > 0) {
      tgt.insertRowsBefore(totalRow, delta);

      const srcFmt = tgt.getRange(templateRow, 1, 1, lastCol);
      const dstFmt = tgt.getRange(templateRow + existingDataCount, 1, delta, lastCol);
      srcFmt.copyTo(dstFmt, { formatOnly: true });
    } else if (delta < 0) {
      tgt.deleteRows(dataStartRow + neededDataCount, -delta);
    }

    // Re-find after row changes
    const info2 = findRoomsBlock_(tgt);
    dataStartRow = info2.dataStartRow;
    totalRow = info2.totalRow;
    colSno = info2.colSno;
    colRooms = info2.colRooms;
    colCarpet = info2.colCarpet;

    const sNoVals = finalRows.map((_, i) => [i + 1]);
    const roomVals = finalRows.map((r) => [r.name]);
    const areaVals = finalRows.map((r) => [r.area]);

    tgt.getRange(dataStartRow, colSno, finalRows.length, 1).setValues(sNoVals);
    tgt.getRange(dataStartRow, colRooms, finalRows.length, 1).setValues(roomVals);
    tgt.getRange(dataStartRow, colCarpet, finalRows.length, 1).setValues(areaVals);

    // Apply font colors to ROOMS column
    const roomFontColors = finalRows.map((r) => [r.color || "#000000"]);
    tgt.getRange(dataStartRow, colRooms, finalRows.length, 1).setFontColors(roomFontColors);

    const totalCell = tgt.getRange(totalRow, colCarpet);
    const startA1 = a1_(dataStartRow, colCarpet);
    const endA1 = a1_(dataStartRow + finalRows.length - 1, colCarpet);
    totalCell.setFormula(`=SUM(${startA1}:${endA1})`);

    SpreadsheetApp.getActive().toast(
      `Hybrid sync done: ${normalPlannerRows.length} normal room(s) + ${fixedRows.length} fixed room(s)`,
      "PLANNER Sync",
      8
    );
    return;
  }

  /**************************************************
   * FIXED MODE
   * - keep target room list as-is
   * - only fill matching Carpet Area
   **************************************************/
  if (mode === "fixed") {
    const info = findRoomsBlock_(tgt);
    const { dataStartRow, totalRow, colRooms, colCarpet } = info;

    const targetRowCount = totalRow - dataStartRow;
    if (targetRowCount <= 0) {
      throw new Error("No room rows found between header and TOTAL.");
    }

    const targetRooms = tgt
      .getRange(dataStartRow, colRooms, targetRowCount, 1)
      .getValues()
      .map((r) => String(r[0] || "").trim());

    const existingCarpetValues = tgt
      .getRange(dataStartRow, colCarpet, targetRowCount, 1)
      .getValues();

    const plannerMap = {};
    planner.forEach((p) => {
      const key = normalizeRoomName_(p.name);
      if (key) {
        plannerMap[key] = {
          area: p.area,
          color: p.color || "#000000"
        };
      }
    });

    const carpetOut = [];
    const colorOut = [];
    let matchedCount = 0;

    for (let i = 0; i < targetRooms.length; i++) {
      const roomName = targetRooms[i];
      const key = normalizeRoomName_(roomName);

      if (key && Object.prototype.hasOwnProperty.call(plannerMap, key)) {
        carpetOut.push([plannerMap[key].area]);
        colorOut.push([plannerMap[key].color || "#000000"]);
        matchedCount++;
      } else {
        carpetOut.push([existingCarpetValues[i][0]]);
        // keep existing font color if no match
        colorOut.push([tgt.getRange(dataStartRow + i, colRooms).getFontColor()]);
      }
    }

    tgt.getRange(dataStartRow, colCarpet, carpetOut.length, 1).setValues(carpetOut);
    tgt.getRange(dataStartRow, colRooms, colorOut.length, 1).setFontColors(colorOut);

    const totalCell = tgt.getRange(totalRow, colCarpet);
    const startA1 = a1_(dataStartRow, colCarpet);
    const endA1 = a1_(totalRow - 1, colCarpet);
    totalCell.setFormula(`=SUM(${startA1}:${endA1})`);

    SpreadsheetApp.getActive().toast(
      `Fixed-mode sync: matched ${matchedCount} room(s)`,
      "PLANNER Sync",
      8
    );
    return;
  }

  /**************************************************
   * DYNAMIC MODE (old behavior)
   * - target fully follows planner
   * - planner room font colors transferred
   **************************************************/
  const targetInfo = findRoomsBlock_(tgt);
  let { dataStartRow, totalRow, colSno, colRooms, colCarpet } = targetInfo;

  const templateRow = dataStartRow;
  if (templateRow >= totalRow) throw new Error("Template row invalid (no data row before TOTAL).");

  const existingDataCount = Math.max(0, totalRow - dataStartRow);
  const neededDataCount = planner.length;
  const delta = neededDataCount - existingDataCount;

  const lastCol = tgt.getLastColumn();

  if (delta > 0) {
    tgt.insertRowsBefore(totalRow, delta);

    const srcFmt = tgt.getRange(templateRow, 1, 1, lastCol);
    const dstFmt = tgt.getRange(templateRow + existingDataCount, 1, delta, lastCol);
    srcFmt.copyTo(dstFmt, { formatOnly: true });
  } else if (delta < 0) {
    tgt.deleteRows(dataStartRow + neededDataCount, -delta);
  }

  const info2 = findRoomsBlock_(tgt);
  dataStartRow = info2.dataStartRow;
  totalRow = info2.totalRow;
  colSno = info2.colSno;
  colRooms = info2.colRooms;
  colCarpet = info2.colCarpet;

  const sNoVals = planner.map((_, i) => [i + 1]);
  const roomVals = planner.map((p) => [p.name]);
  const areaVals = planner.map((p) => [p.area]);
  const colorVals = planner.map((p) => [p.color || "#000000"]);

  tgt.getRange(dataStartRow, colSno, planner.length, 1).setValues(sNoVals);
  tgt.getRange(dataStartRow, colRooms, planner.length, 1).setValues(roomVals);
  tgt.getRange(dataStartRow, colCarpet, planner.length, 1).setValues(areaVals);
  tgt.getRange(dataStartRow, colRooms, planner.length, 1).setFontColors(colorVals);

  const totalCell = tgt.getRange(totalRow, colCarpet);
  const startA1 = a1_(dataStartRow, colCarpet);
  const endA1 = a1_(dataStartRow + planner.length - 1, colCarpet);
  totalCell.setFormula(`=SUM(${startA1}:${endA1})`);

  SpreadsheetApp.getActive().toast(
    `Dynamic sync: copied ${planner.length} room(s)`,
    "PLANNER Sync",
    8
  );
}

/* -------------------- PLANNER SOURCE -------------------- */

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

    const header = sh
      .getRange(1, 1, 1, lastCol)
      .getValues()[0]
      .map((v) => String(v || "").trim().toLowerCase());

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
  const header = values[0].map((v) => String(v || "").trim().toLowerCase());

  const idxName = header.indexOf("name");
  const idxArea = header.indexOf("area_sqft");

  if (idxName === -1 || idxArea === -1) {
    throw new Error(`Source sheet "${srcSheet.getName()}" must contain headers: name, area_sqft`);
  }

  const fontColors = srcSheet.getRange(2, 1, lastRow - 1, lastCol).getFontColors();

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const name = String(values[r][idxName] || "").trim();
    if (!name) continue;

    const area = toNumber_(values[r][idxArea]);
    const nameColor = fontColors[r - 1][idxName];

    out.push({
      name,
      area: area ?? "",
      color: nameColor || "#000000"
    });
  }
  return out;
}

/* -------------------- PLANNER TARGET (merge-safe TOTAL) -------------------- */

function findRoomsBlock_(tgtSheet) {
  const lastRow = tgtSheet.getLastRow();
  const lastCol = tgtSheet.getLastColumn();

  // Search top area for header row: S. No | ROOMS | Carpet Area
  const scanRows = Math.min(lastRow, 60);
  const scanCols = Math.min(lastCol, 20);
  const scan = tgtSheet.getRange(1, 1, scanRows, scanCols).getValues();

  let headerRow = -1;
  let colSno = -1;
  let colRooms = -1;
  let colCarpet = -1;

  for (let r = 0; r < scan.length; r++) {
    const row = scan[r].map((v) => String(v || "").trim().toLowerCase());

    const snoIdx = row.findIndex((v) => v.replace(/\./g, "") === "s no" || v === "s. no");
    const roomsIdx = row.findIndex((v) => v === "rooms");
    const carpetIdx = row.findIndex((v) => v === "carpet area");

    if (snoIdx !== -1 && roomsIdx !== -1 && carpetIdx !== -1) {
      headerRow = r + 1;
      colSno = snoIdx + 1;
      colRooms = roomsIdx + 1;
      colCarpet = carpetIdx + 1;
      break;
    }
  }

  if (headerRow === -1) {
    throw new Error("Header row not found (need: S. No, ROOMS, Carpet Area).");
  }

  const dataStartRow = headerRow + 1;

  // Find TOTAL row by scanning A..F (or configured) for the word "TOTAL"
  const scanToCol = Math.min(lastCol, Math.max(PLANNER_SYNC.TOTAL_SCAN_COLS, colCarpet));
  const rowsBelow = Math.max(1, lastRow - headerRow);
  const block = tgtSheet.getRange(headerRow + 1, 1, rowsBelow, scanToCol).getValues();

  let totalRow = -1;
  for (let i = 0; i < block.length; i++) {
    const rowText = block[i]
      .map((v) => String(v || "").trim().toLowerCase())
      .filter(Boolean);

    if (rowText.some((t) => t === PLANNER_SYNC.TOTAL_TEXT)) {
      totalRow = headerRow + 1 + i;
      break;
    }
  }

  if (totalRow === -1) throw new Error("TOTAL row not found.");

  return { headerRow, dataStartRow, totalRow, colSno, colRooms, colCarpet };
}

/* -------------------- SHARED UTILS -------------------- */

function normalizeRoomName_(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/&/g, "and")
    .replace(/[^a-z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

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