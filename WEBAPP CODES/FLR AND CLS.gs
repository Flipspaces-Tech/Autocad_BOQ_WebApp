const FLOOR_CL_SYNC = {
  GENERATED_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",
  GENERATED_TAB: "LAYER",

  MAPPING_SS_ID: "1wat9koeZC9puQOvH9gC9zxrkJvZxIn5in7EiwlHleus",
  MAPPING_TAB: "BOQ-LAYER",

  TARGET_TAB: "CALCULATION SHEET",
  PREFIXES: ["FS-FLR", "FS-CL"],
  TOTAL_TEXT: "total",
  TOTAL_SCAN_COLS: 8,
};

function cleanHeader_(v) {
  return String(v || "")
    .toLowerCase()
    .replace(/\n/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function nz_(n) {
  const x = toNumber_(n);
  return x == null ? 0 : x;
}

function round2_(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}
function syncFloorClToCalculationSheet() {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  const targetSh = masterSS.getSheetByName(FLOOR_CL_SYNC.TARGET_TAB);
  if (!targetSh) throw new Error(`Target tab not found: ${FLOOR_CL_SYNC.TARGET_TAB}`);

  const mappingSS = SpreadsheetApp.openById(FLOOR_CL_SYNC.MAPPING_SS_ID);
  const mappingSh = mappingSS.getSheetByName(FLOOR_CL_SYNC.MAPPING_TAB);
  if (!mappingSh) throw new Error(`Mapping tab not found: ${FLOOR_CL_SYNC.MAPPING_TAB}`);

  const generatedSS = SpreadsheetApp.openById(FLOOR_CL_SYNC.GENERATED_SS_ID);
  const generatedSh = generatedSS.getSheetByName(FLOOR_CL_SYNC.GENERATED_TAB);
  if (!generatedSh) throw new Error(`Generated tab not found: ${FLOOR_CL_SYNC.GENERATED_TAB}`);

  const mappings = readFloorClMappingsRobust_(mappingSh);

  Logger.log("Total filtered mappings = " + mappings.list.length);
  Logger.log(JSON.stringify(mappings.list.slice(0, 20), null, 2));

  if (!mappings.list.length) {
    SpreadsheetApp.getUi().alert(
      "FLR/CL Sync",
      "No FS-FLR / FS-CL mappings found.\n\nPlease check:\n1) mapping spreadsheet ID\n2) mapping tab name\n3) Layer Name column values",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const generatedRows = readGeneratedLayerRowsRobust_(generatedSh);
  Logger.log("Generated rows read = " + generatedRows.length);

  const agg = {};

  for (const row of generatedRows) {
    const layerNorm = normKey_(row.layer);
    const zoneNorm = normKey_(row.zone);
    if (!layerNorm || !zoneNorm) continue;

    const mapEntry = mappings.byAlias[layerNorm];
    if (!mapEntry) continue;

    const val = pickMeasurementValue_(row, mapEntry);
    if (val == null) continue;

    const boqNorm = normKey_(mapEntry.boqName);
    if (!agg[boqNorm]) agg[boqNorm] = {};
    if (!agg[boqNorm][zoneNorm]) agg[boqNorm][zoneNorm] = 0;

    agg[boqNorm][zoneNorm] += val;
  }

  const block = findTopRoomMatrixBlock_(targetSh);
  const roomRowMap = buildRoomRowMap_(targetSh, block);
  const itemColMap = buildItemColumnMap_(targetSh, block);

  const affectedCols = [];
  mappings.list.forEach((m) => {
    const c = itemColMap[normKey_(m.boqName)];
    if (c) affectedCols.push(c);
  });

  clearNonContiguousCols_(targetSh, block.dataStartRow, block.totalRow - 1, affectedCols);

  let writes = 0;
  Object.keys(agg).forEach((boqNorm) => {
    const col = itemColMap[boqNorm];
    if (!col) return;

    Object.keys(agg[boqNorm]).forEach((zoneNorm) => {
      const row = roomRowMap[zoneNorm];
      if (!row) return;

      targetSh.getRange(row, col).setValue(round2_(agg[boqNorm][zoneNorm]));
      writes++;
    });
  });

  rebuildTotalsForCols_(targetSh, block.dataStartRow, block.totalRow, affectedCols);

  SpreadsheetApp.getActive().toast(
    `FLR/CL sync done. Cells written: ${writes}`,
    "FLR/CL Sync",
    8
  );
}

function readFloorClMappingsRobust_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return { list: [], byAlias: {} };

  const vals = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  let headerRow = -1;
  let idxBoq = -1;
  let idxUnits = -1;
  let idxLayer = -1;
  let idxLayerOld = -1;
  let idxMeasurement = -1;

  for (let r = 0; r < Math.min(10, vals.length); r++) {
    const row = vals[r].map((v) => cleanHeader_(v));

    for (let c = 0; c < row.length; c++) {
      const h = row[c];
      if (h === "boq name") idxBoq = c;
      else if (h === "units") idxUnits = c;
      else if (h === "layer name" && idxLayer === -1) idxLayer = c;
      else if ((h === "layer name-old" || h === "layer name old") && idxLayerOld === -1) idxLayerOld = c;
      else if (h === "measurement") idxMeasurement = c;
    }

    if (idxBoq !== -1 && idxLayer !== -1) {
      headerRow = r;
      break;
    }
  }

  if (headerRow === -1) {
    throw new Error(`Could not detect mapping headers in ${sh.getName()}`);
  }

  Logger.log(
    JSON.stringify({
      headerRow: headerRow + 1,
      idxBoq,
      idxUnits,
      idxLayer,
      idxLayerOld,
      idxMeasurement,
    })
  );

  const out = [];
  const byAlias = {};

  for (let r = headerRow + 1; r < vals.length; r++) {
    const row = vals[r];

    const boqName = String(row[idxBoq] || "").trim();
    const layerName = idxLayer >= 0 ? String(row[idxLayer] || "").trim() : "";
    const layerOld = idxLayerOld >= 0 ? String(row[idxLayerOld] || "").trim() : "";
    const units = idxUnits >= 0 ? String(row[idxUnits] || "").trim().toUpperCase() : "";
    const measurement = idxMeasurement >= 0 ? String(row[idxMeasurement] || "").trim() : "";

    if (!boqName || !layerName) continue;

    const canonical = layerName.toUpperCase().trim();
    if (!FLOOR_CL_SYNC.PREFIXES.some((p) => canonical.startsWith(p))) continue;

    const entry = {
      boqName,
      layerName,
      layerOld,
      units,
      measurement,
    };
    out.push(entry);

    [layerName, layerOld]
      .filter(Boolean)
      .forEach((alias) => {
        byAlias[normKey_(alias)] = entry;
      });
  }

  return { list: out, byAlias };
}

function readGeneratedLayerRowsRobust_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return [];

  const vals = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  let headerRow = -1;
  let idxLayer = -1;
  let idxZone = -1;
  let idxLength = -1;
  let idxPerimeter = -1;
  let idxArea = -1;

  for (let r = 0; r < Math.min(10, vals.length); r++) {
    const row = vals[r].map(v => cleanHeader_(v));

    for (let c = 0; c < row.length; c++) {
      const h = row[c];
      if (h === "layer") idxLayer = c;
      else if (h === "zone") idxZone = c;
      else if (h === "length (ft)" || h === "length") idxLength = c;
      else if (h === "perimeter") idxPerimeter = c;
      else if (h === "area (ft2)" || h === "area") idxArea = c;
    }

    if (idxLayer !== -1 && idxZone !== -1) {
      headerRow = r;
      break;
    }
  }

  if (headerRow === -1) {
    throw new Error(`Could not detect generated headers in ${sh.getName()}`);
  }

  const out = [];
  let currentLayer = "";

  for (let r = headerRow + 1; r < vals.length; r++) {
    const row = vals[r];

    const layerCell = String(row[idxLayer] || "").trim();
    if (layerCell) currentLayer = layerCell;

    const zone = String(row[idxZone] || "").trim();
    if (!currentLayer || !zone) continue;

    out.push({
      layer: currentLayer,
      zone: zone.toLowerCase() === "unmarked area" ? "misc" : zone,
      length: idxLength >= 0 ? toNumber_(row[idxLength]) : null,
      perimeter: idxPerimeter >= 0 ? toNumber_(row[idxPerimeter]) : null,
      area: idxArea >= 0 ? toNumber_(row[idxArea]) : null,
    });
  }

  return out;
}

function pickMeasurementValue_(row, mapEntry) {
  const m = String(mapEntry.measurement || "").toLowerCase();
  const u = String(mapEntry.units || "").toUpperCase();

  const area = nz_(row.area);
  const perimeter = nz_(row.perimeter);
  const length = nz_(row.length);

  // Best-fit rule for your flooring/ceiling sync:
  // if area exists, prefer it first because your expected master output
  // (e.g. FS-FLR-MDS 75 -> 100, 232, 434, 56) is clearly coming from area column.
  if (area > 0) return area;

  // Then follow unit/measurement hints
  if (u === "SQFT") return area;
  if (u === "RFT") {
    if (m.indexOf("perimeter") !== -1 && perimeter > 0) return perimeter;
    if (length > 0) return length;
    return perimeter;
  }
  if (u === "COUNT") return 1;

  if (m.indexOf("area") !== -1 && area > 0) return area;
  if (m.indexOf("perimeter") !== -1 && perimeter > 0) return perimeter;
  if (m.indexOf("length") !== -1 && length > 0) return length;
  if (m.indexOf("count") !== -1) return 1;

  // Final fallback
  if (perimeter > 0) return perimeter;
  if (length > 0) return length;

  return null;
}

function findTopRoomMatrixBlock_(tgtSheet) {
  const lastRow = tgtSheet.getLastRow();
  const lastCol = tgtSheet.getLastColumn();
  const scan = tgtSheet
    .getRange(1, 1, Math.min(60, lastRow), Math.min(80, lastCol))
    .getDisplayValues();

  let headerRow = -1;
  let colRooms = -1;
  let colCarpet = -1;

  for (let r = 0; r < scan.length; r++) {
    const row = scan[r].map((v) => cleanHeader_(v));
    for (let c = 0; c < row.length; c++) {
      if (row[c] === "rooms") colRooms = c + 1;
      if (row[c] === "carpet area") colCarpet = c + 1;
    }
    if (colRooms !== -1 && colCarpet !== -1) {
      headerRow = r + 1;
      break;
    }
  }

  if (headerRow === -1) throw new Error("Top master header row not found.");

  const dataStartRow = headerRow + 1;
  const block = tgtSheet
    .getRange(dataStartRow, 1, lastRow - dataStartRow + 1, Math.min(10, lastCol))
    .getDisplayValues();

  let totalRow = -1;
  for (let i = 0; i < block.length; i++) {
    const row = block[i].map((v) => cleanHeader_(v));
    if (row.indexOf("total") !== -1) {
      totalRow = dataStartRow + i;
      break;
    }
  }

  if (totalRow === -1) throw new Error("TOTAL row not found.");

  return { headerRow, dataStartRow, totalRow, colRooms, colCarpet, lastCol };
}

function buildRoomRowMap_(sh, block) {
  const vals = sh
    .getRange(block.dataStartRow, block.colRooms, block.totalRow - block.dataStartRow, 1)
    .getDisplayValues();

  const map = {};
  for (let i = 0; i < vals.length; i++) {
    const room = String(vals[i][0] || "").trim();
    if (room) map[normKey_(room)] = block.dataStartRow + i;
  }
  return map;
}

function buildItemColumnMap_(sh, block) {
  const vals = sh.getRange(block.headerRow, 1, 1, block.lastCol).getDisplayValues()[0];
  const map = {};

  for (let c = block.colCarpet + 1; c <= block.lastCol; c++) {
    const txt = String(vals[c - 1] || "").trim();
    if (txt) map[normKey_(txt)] = c;
  }
  return map;
}

function clearNonContiguousCols_(sh, rowStart, rowEnd, cols) {
  const numRows = rowEnd - rowStart + 1;
  if (numRows <= 0) return;

  [...new Set(cols)].forEach((col) => {
    sh.getRange(rowStart, col, numRows, 1).clearContent();
  });
}

function rebuildTotalsForCols_(sh, dataStartRow, totalRow, cols) {
  [...new Set(cols)].forEach((col) => {
    sh.getRange(totalRow, col).setFormula(`=SUM(${a1_(dataStartRow, col)}:${a1_(totalRow - 1, col)})`);
  });
}