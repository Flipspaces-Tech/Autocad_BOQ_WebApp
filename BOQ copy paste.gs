/**************************************************
 * BOQ QTY Sync (All Tabs) — formula-safe, with Cell + Zone in log
 * Target: this spreadsheet (active)
 * Source: DXF pipeline output sheet
 **************************************************/

const CONFIG = {
  // SOURCE (DXF pipeline output)
  SYNC_LOG_MODE: 'replace',
  SOURCE_SS_ID: '12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM', // ← your GSHEET_ID
  SOURCE_TAB_NAME: 'CRED NEW',                                   // ← your GSHEET_TAB

  // Behavior
  SUM_DUPLICATES: true,                 // sum qty for duplicate BOQ names in source
  IGNORE_TABS: ['Summary', 'Sync Log'], // tabs to skip in target
  MIN_HEADER_ROW: 1                     // header assumed on row 1
};

// Flexible, case-insensitive header matching
const HEADER_ALIASES = {
  // Source columns
  SOURCE_BOQ:   [/^boq\s*name$/i, /^boq$/i, /^item\s*name$/i, /^description$/i, /^description\s*name$/i],
  SOURCE_QTY:   [/^qty[_\s-]*value$/i, /^quantity$/i, /^qty$/i, /^count$/i, /^quantity\s*value$/i],
  SOURCE_ZONE:  [/^zone$/i],

  // Target columns (include “SCOPE OF WORK”)
  TARGET_BOQ:   [/^boq\s*name$/i, /^boq$/i, /^item\s*name$/i, /^description$/i, /^description\s*name$/i, /^scope\s*of\s*work$/i],
  TARGET_QTY:   [/^qty$/i, /^quantity$/i, /^qty\s*value$/i, /^qty[_\s-]*value$/i]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sync')
    .addItem('Fill QTY in all tabs', 'syncQtyAllTabs')
    .addToUi();
}

/** Main entry: scans all tabs and fills QTY where BOQ matches (formula-safe). */
function syncQtyAllTabs() {
  const targetSS = SpreadsheetApp.getActiveSpreadsheet();
  const t0 = Date.now();

  // Build name -> { qty, zones:Set } map from source
  const nameMap = loadSourceMap_(); // Map<string, {qty:number, zones:Set<string>}>

  let totalHits = 0, totalMisses = 0, totalSheets = 0;
  const matchedAll = []; // {sheet, cell, name, qty, zonesStr}

  targetSS.getSheets().forEach(sh => {
    const sheetName = sh.getName();
    if (CONFIG.IGNORE_TABS.includes(sheetName)) return;

    const { hits, misses, updated, matched } = updateOneSheetQty_(sh, nameMap);
    if (updated) totalSheets++;
    totalHits += hits;
    totalMisses += misses;
    matchedAll.push(...matched);
  });

  // Aggregate for toast
  const agg = new Map(); // display name -> {sumQty, sheets:Set<string>}
  matchedAll.forEach(m => {
    const key = m.name;
    if (!agg.has(key)) agg.set(key, { sumQty: m.qty, sheets: new Set([m.sheet]) });
    else { const o = agg.get(key); o.sumQty += m.qty; o.sheets.add(m.sheet); }
  });

  const list = Array.from(agg.entries()).map(([n,o]) => `${n} = ${o.sumQty} (${Array.from(o.sheets).join(', ')})`);
  const preview = list.slice(0, 10).join('; ');
  const more = list.length > 10 ? ` … +${list.length - 10} more` : '';
  const ms = Date.now() - t0;

  // Persist full log (now with Cell + Zone)
  writeSyncLog_(matchedAll, totalHits, totalMisses, totalSheets, ms);

  const msg = `QTY sync → Sheets: ${totalSheets}, Matches: ${totalHits}, Missed: ${totalMisses}` +
              (list.length ? `\nMatched: ${preview}${more}` : `\nMatched: —`);
  SpreadsheetApp.getActive().toast(msg, 'BOQ QTY Sync', 10);
  console.log(msg);
}

/* ------------ Internals ------------ */

// Read source sheet into a Map of normalized name -> {qty, zones:Set}
function loadSourceMap_() {
  const srcSh = SpreadsheetApp.openById(CONFIG.SOURCE_SS_ID).getSheetByName(CONFIG.SOURCE_TAB_NAME);
  if (!srcSh) throw new Error(`Source tab "${CONFIG.SOURCE_TAB_NAME}" not found in ${CONFIG.SOURCE_SS_ID}`);

  const { data, headerMap } = readTable_(srcSh);
  const idxBOQ  = findHeaderIndex_(headerMap, HEADER_ALIASES.SOURCE_BOQ,  'SOURCE: BOQ name');
  const idxQTY  = findHeaderIndex_(headerMap, HEADER_ALIASES.SOURCE_QTY,  'SOURCE: qty_value');
  const idxZONE = tryFindHeaderIndex_(headerMap, HEADER_ALIASES.SOURCE_ZONE); // optional

  const out = new Map();                               // nameNorm -> { qty:number, zones:Set<string> }
  let lastZone = '';                                    // <- forward-fill cache

  for (const row of data) {
    const nameNorm = normalizeName_(row[idxBOQ]);
    const qty = toNumber_(row[idxQTY]);
    if (!nameNorm || qty == null) continue;

    let z = '';
    if (idxZONE != null) {
      const raw = String(row[idxZONE] ?? '').trim();
      if (raw) {
        z = raw;
        lastZone = raw;                                 // remember newest non-empty zone
      } else {
        z = lastZone;                                   // forward-fill zone for merged blanks
      }
    }

    if (!out.has(nameNorm)) out.set(nameNorm, { qty: 0, zones: new Set() });

    if (CONFIG.SUM_DUPLICATES) out.get(nameNorm).qty += qty;
    else out.get(nameNorm).qty = qty;

    if (z) out.get(nameNorm).zones.add(z);
  }
  return out;
}


// Update one target sheet (formula-safe). Returns hits/misses/updated + matched writes list.
function updateOneSheetQty_(sheet, nameMap) {
  const result = { hits: 0, misses: 0, updated: false, matched: [] };

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return result;

  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getValues();
  const formulas = range.getFormulas();
  if (!values.length) return result;

  const headers = (values[0] || []).map(h => String(h || '').trim());
  const headerMap = new Map(headers.map((h, i) => [h.toLowerCase(), i]));

  let idxBOQ, idxQTY;
  try {
    idxBOQ = findHeaderIndex_(headerMap, HEADER_ALIASES.TARGET_BOQ, `TARGET: BOQ (${sheet.getName()})`);
    idxQTY = findHeaderIndex_(headerMap, HEADER_ALIASES.TARGET_QTY, `TARGET: QTY (${sheet.getName()})`);
  } catch (_e) { return result; }

  const dataVals = values.slice(CONFIG.MIN_HEADER_ROW);
  const dataForm = formulas.slice(CONFIG.MIN_HEADER_ROW);

  const qtyColOut = new Array(dataVals.length);
  let changed = false;
  const gid = sheet.getSheetId();              // <— add gid for hyperlinking
  const firstDataRow = CONFIG.MIN_HEADER_ROW + 1;

  for (let r = 0; r < dataVals.length; r++) {
    const rowVals = dataVals[r];
    const rowForm = dataForm[r];
    const rawName = rowVals[idxBOQ];
    const nameNorm = normalizeName_(rawName);

    const existingFormula = rowForm[idxQTY] && rowForm[idxQTY].toString().trim();
    if (existingFormula) {
      qtyColOut[r] = [existingFormula];
      result.misses++;
      continue;
    }

    if (!nameNorm) {
      qtyColOut[r] = [rowVals[idxQTY]];
      result.misses++;
      continue;
    }

    if (nameMap.has(nameNorm)) {
      const { qty, zones } = nameMap.get(nameNorm);
      qtyColOut[r] = [qty];
      if (rowVals[idxQTY] !== qty) changed = true;
      result.hits++;

      // log the exact address + zones from source
      const a1 = a1For_(firstDataRow + r, idxQTY + 1);
      const zonesStr = zones && zones.size ? Array.from(zones).join(', ') : '';
      result.matched.push({
        sheet: sheet.getName(),
        gid: gid,                 // <— will be used to build the hyperlink
        cell: a1,
        name: String(rawName || '').trim(),
        qty: qty,
        zones: zonesStr
      });
    } else {
      qtyColOut[r] = [rowVals[idxQTY]];
      result.misses++;
    }
  }

  if (changed) {
    sheet.getRange(firstDataRow, idxQTY + 1, qtyColOut.length, 1).setValues(qtyColOut);
    result.updated = true;
  }
  return result;
}


// Persist a full run log to "Sync Log" sheet (includes Cell + Zones)
function writeSyncLog_(matchedAll, hits, misses, sheets, ms) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const name = 'Sync Log';
  const sh   = ss.getSheetByName(name) || ss.insertSheet(name);

  // Ensure header
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, 9).setValues([[
      'Timestamp', 'Sheet', 'Cell', 'BOQ Name', 'Zone(s)', 'QTY', 'Run Hits', 'Run Misses', 'Run (ms)'
    ]]);
    sh.setFrozenRows(1);
  }

  const ts      = new Date();
  const baseUrl = `https://docs.google.com/spreadsheets/d/${ssId}/edit`;

  // Build rows (Cell column becomes hyperlink formula in a second pass)
  const rows = matchedAll.length
    ? matchedAll.map(m => {
        const link = `${baseUrl}#gid=${m.gid}&range=${encodeURIComponent(m.cell)}`;
        return [ts, m.sheet, link, m.name, m.zones || '', m.qty, hits, misses, ms];
      })
    : [[ts, '-', '-', '-', '-', '-', hits, misses, ms]];

  // -------- REFRESH vs APPEND handling --------
  let startRow;
  if ((CONFIG.SYNC_LOG_MODE || 'append').toLowerCase() === 'replace') {
    // Clear everything below the header
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, sh.getMaxColumns()).clearContent();
    startRow = 2;  // always write from row 2
  } else {
    // Append mode
    startRow = sh.getLastRow() + 1;
  }

  // Write raw rows
  sh.getRange(startRow, 1, rows.length, 9).setValues(rows);

  // Convert column C to HYPERLINK(CellURL, A1Label)
  if (matchedAll.length) {
    const labels   = matchedAll.map(m => m.cell);
    const linkCol  = sh.getRange(startRow, 3, matchedAll.length, 1); // column C
    const formulas = labels.map((label, i) => {
      const url = rows[i][2];
      return [`=HYPERLINK("${url}","${label}")`];
    });
    linkCol.setFormulas(formulas);
  }
}



/* ------------ Shared utils ------------ */

function readTable_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(1, 1, lastRow, lastCol);
  const values = range.getValues();
  const headers = (values[0] || []).map(h => String(h || '').trim());
  const headerMap = new Map(headers.map((h, i) => [h.toLowerCase(), i]));
  const data = values.slice(CONFIG.MIN_HEADER_ROW);
  return { headers, headerMap, data };
}

function findHeaderIndex_(headerMap, patterns, labelForError) {
  for (const [hLower, idx] of headerMap.entries()) {
    for (const rx of patterns) {
      if (rx.test(hLower)) return idx;
    }
  }
  throw new Error(`Could not find header for ${labelForError}. Searched: ${patterns.map(r => r.toString()).join(', ')}`);
}

function tryFindHeaderIndex_(headerMap, patterns) {
  if (!patterns) return null;
  for (const [hLower, idx] of headerMap.entries()) {
    for (const rx of patterns) {
      if (rx.test(hLower)) return idx;
    }
  }
  return null; // optional
}

function normalizeName_(s) {
  if (s == null) return '';
  return String(s).trim().replace(/\s+/g, ' ').toLowerCase();
}

function toNumber_(v) {
  if (v == null || v === '') return null;
  const n = Number(v);
  return Number.isFinite(n) ? n : null;
}

function a1For_(row1, col1) {
  return colToA1_(col1) + String(row1);
}
function colToA1_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = (n - 1) / 26 >> 0;
  }
  return s;
}

/* -------- Optional: quick header scan debug -------- */
function debugHeaderScan() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    const lastCol = sh.getLastColumn();
    if (sh.getLastRow() < 1 || lastCol < 1) { console.log(name, 'empty'); return; }
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h || '').toString().trim());
    const lower = headers.map(h => h.toLowerCase());
    const hasBOQ = HEADER_ALIASES.TARGET_BOQ.some(rx => lower.some(h => rx.test(h)));
    const hasQTY = HEADER_ALIASES.TARGET_QTY.some(rx => lower.some(h => rx.test(h)));
    console.log(`[${name}] BOQ? ${hasBOQ}  QTY? ${hasQTY}  headers=`, headers);
  });
}
