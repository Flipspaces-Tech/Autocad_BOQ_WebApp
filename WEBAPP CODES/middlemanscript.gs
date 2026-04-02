/**************************************************
 * BOQ-LAYER → MASTER (MEAS + QTY + HIDE ZEROS)
 *
 * FLOW:
 * 1) Middleman sheet hosts script + BOQ-LAYER mapping
 * 2) Every run creates a NEW COPY of the master template spreadsheet
 * 3) Sync runs on that new copied spreadsheet
 * 4) Popup shows running state, then success + clickable link
 * 5) New copy ID is stored in Script Properties for downstream scripts
 *
 * EXTRA MERGE LOGIC:
 * After split rows are created, all columns are vertically merged
 * EXCEPT:
 *   LOCATION, MEASUREMENT, HEIGHT, QTY
 **************************************************/

const SHEETS = SpreadsheetApp;

const BOQ_SYNC = {
  // Mapping sheet (script hosted here)
  MAP_TAB: "BOQ-LAYER",

  // Mapping columns (1-based)
  MAP_COL_TARGETED: 2,
  MAP_COL_GENERATED: 4,
  MAP_COL_BLOCKNAME: 6,
  MAP_COL_GENERATED_BLOCK: 5,    // ✅ Column E: Generated-Block (check FIRST)
  MAP_COL_MEASURE: 7, // put the correct column number here



  // Export spreadsheet
  EXPORT_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",

  // SHEET_WHITELIST: ["Flooring", "Ceiling"],
  SHEET_WHITELIST: null,
  // SHEET_WHITELIST: ["Civil"],

  // Measurement source tab
  EXPORT_MEASURE_TAB: "LAYER",
  LAYER_HDR_LAYER: "layer",
  LAYER_HDR_ZONE: "zone",
  LAYER_HDR_AREA: "area (ft2)",
  LAYER_HDR_PERIM: "perimeter",
  LAYER_HDR_LENGTH: "length (ft)",

  VIS_HDR_LENGTH: "length (ft)",
  VIS_HDR_ZONE: "zone",
  VIS_HDR_PRODUCT: "product",
  VIS_HDR_BOQ: "boq name",
  VIS_HDR_QTY: "qty_value",

  // QTY source tab
  EXPORT_QTY_TAB: "vis_export_sheet_like",
  QTY_HDR_PRODUCT_PATTERNS: ["product"],
  QTY_HDR_BOQNAME_PATTERNS: ["boq name", "boq_name", "name"],
  QTY_HDR_ZONE_PATTERNS: ["zone", "location"],
  QTY_HDR_QTYVALUE_PATTERNS: ["qty_value", "qty value", "qty", "quantity","qty_value"],

  // Master template spreadsheet
  MASTER_TEMPLATE_SS_ID: "1rFikkFgJ84wl9Mqft1WG2UUT2PpqQK8se0aMNsKPG70",
  MASTER_START_TAB: "Civil",

  // Header detection in master
  MASTER_MATCH_COL_FALLBACK: 2,
  MASTER_QTY_COL_FALLBACK: 9,

  MASTER_HDR_SCOPE: ["scope of work"],
  MASTER_HDR_LOCATION: ["location"],
  MASTER_HDR_MEASUREMENT: ["measurement", "qty measured", "measured"],
  MASTER_HDR_HEIGHT: ["height"],
  MASTER_HDR_QTY: ["qty", "quantity"],
  MASTER_HDR_TOTAL_QTY: ["total qty", "total quantity"],
  MASTER_HDR_SRNO: ["sr. no.", "sr no", "sr.no", "sr"],

  SKIP_WHEN_BOTH_ZERO: true,
  HIDE_ZERO_ROWS_AFTER_SYNC: true,

  NUMBER_FORMAT: "0.############",
  SHOW_DIALOG: false,
  DIALOG_TITLE: "Vizdom Sync — BOQ-LAYER → MASTER",

  // latest generated working copy
  LATEST_MASTER_COPY_ID_KEY: "LATEST_MASTER_COPY_ID",
};

function onOpen() {
  SHEETS.getUi()
    .createMenu("AUTO QA")
    .addItem("BOQ layer to master copy", "launchSyncBoqLayerToMasterUi")
    .addToUi();
}

/* =========================================================
   UI launcher with progress popup
========================================================= */
function launchSyncBoqLayerToMasterUi() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        <style>
          body {
            font-family: Arial, sans-serif;
            padding: 18px;
            color: #1f1f1f;
            margin: 0;
          }
          .wrap {
            display: flex;
            flex-direction: column;
            gap: 14px;
          }
          .title {
            font-size: 18px;
            font-weight: 700;
            color: #0b57d0;
          }
          .status {
            font-size: 14px;
            line-height: 1.5;
          }
          .row {
            display: flex;
            align-items: center;
            gap: 12px;
          }
          .spinner {
            width: 20px;
            height: 20px;
            border: 3px solid #d7e3fc;
            border-top: 3px solid #0b57d0;
            border-radius: 50%;
            animation: spin 1s linear infinite;
            flex: 0 0 auto;
          }
          @keyframes spin {
            0% { transform: rotate(0deg); }
            100% { transform: rotate(360deg); }
          }
          .barWrap {
            width: 100%;
            height: 12px;
            background: #e6e6e6;
            border-radius: 999px;
            overflow: hidden;
            box-shadow: inset 0 1px 2px rgba(0,0,0,0.08);
          }
          .bar {
            width: 0%;
            height: 100%;
            background: linear-gradient(90deg, #34a853, #66bb6a);
            border-radius: 999px;
            transition: width 0.45s ease;
          }
          .muted {
            color: #5f6368;
            font-size: 12px;
          }
          .success {
            color: #137333;
            font-weight: 700;
            font-size: 15px;
          }
          .error {
            color: #c5221f;
            font-weight: 700;
            white-space: pre-wrap;
            font-size: 13px;
          }
          .btn {
            display: inline-block;
            padding: 10px 14px;
            background: #0b57d0;
            color: white;
            text-decoration: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
          }
          .btnSecondary {
            display: inline-block;
            padding: 10px 14px;
            background: #f1f3f4;
            color: #1f1f1f;
            text-decoration: none;
            border-radius: 8px;
            font-size: 14px;
            font-weight: 600;
            border: 1px solid #dadce0;
            cursor: pointer;
          }
          .btnRow {
            display: flex;
            gap: 10px;
            flex-wrap: wrap;
            margin-top: 8px;
          }
          .card {
            background: #f8f9fa;
            border: 1px solid #e0e0e0;
            border-radius: 10px;
            padding: 12px;
          }
          .kv {
            font-size: 13px;
            line-height: 1.6;
            word-break: break-word;
          }
        </style>
      </head>
      <body>
        <div class="wrap">
          <div class="title">AUTO QA in Progress</div>

          <div id="runningView">
            <div class="row">
              <div class="spinner"></div>
              <div class="status" id="statusText">Starting sync and creating a new master copy...</div>
            </div>

            <div class="barWrap">
              <div class="bar" id="progressBar"></div>
            </div>

            <div class="muted" id="hintText">
              Please wait. This popup will update automatically when the process finishes.
            </div>
          </div>

          <div id="doneView" style="display:none;"></div>
        </div>

        <script>
          const statusText = document.getElementById("statusText");
          const runningView = document.getElementById("runningView");
          const doneView = document.getElementById("doneView");
          const progressBar = document.getElementById("progressBar");

          const stagedMessages = [
            { text: "Starting sync and creating a new master copy...", pct: 10 },
            { text: "Reading BOQ-LAYER mappings...", pct: 25 },
            { text: "Reading export measurement and quantity data...", pct: 45 },
            { text: "Processing master tabs and writing rows...", pct: 72 },
            { text: "Applying hide/show logic and finalizing report...", pct: 92 }
          ];

          let i = 0;
          progressBar.style.width = "6%";
          statusText.textContent = stagedMessages[0].text;

          const timer = setInterval(() => {
            i++;
            if (i >= stagedMessages.length) {
              i = stagedMessages.length - 1;
            }
            statusText.textContent = stagedMessages[i].text;
            progressBar.style.width = stagedMessages[i].pct + "%";
          }, 1800);

          google.script.run
            .withSuccessHandler((result) => {
              clearInterval(timer);
              progressBar.style.width = "100%";
              statusText.textContent = "Sync completed successfully.";

              setTimeout(() => {
                runningView.style.display = "none";
                doneView.style.display = "block";

                const safe = (v) => (v == null ? "" : String(v))
                  .replace(/&/g, "&amp;")
                  .replace(/</g, "&lt;")
                  .replace(/>/g, "&gt;")
                  .replace(/"/g, "&quot;");

                doneView.innerHTML = \`
                  <div class="success">Sync completed successfully.</div>

                  <div class="card">
                    <div class="kv"><b>Master Copy Name:</b> \${safe(result.masterCopyName)}</div>
                    <div class="kv"><b>Master Copy ID:</b> \${safe(result.masterCopyId)}</div>
                    <div class="kv"><b>Rows Updated:</b> \${safe(result.rowsUpdated)}</div>
                    <div class="kv"><b>Rows Inserted:</b> \${safe(result.rowsInserted)}</div>
                    <div class="kv"><b>Rows Hidden:</b> \${safe(result.rowsHidden)}</div>
                  </div>

                  <div class="btnRow">
                    <a class="btn" href="\${safe(result.masterCopyUrl)}" target="_blank">Open New Master Copy</a>
                    <button class="btnSecondary" onclick="google.script.host.close()">Close</button>
                  </div>
                \`;
              }, 350);
            })
            .withFailureHandler((err) => {
              clearInterval(timer);
              runningView.style.display = "none";
              doneView.style.display = "block";

              const msg = err && err.message ? err.message : String(err);

              doneView.innerHTML = \`
                <div class="error">Sync failed.</div>
                <div class="card">
                  <div class="error">\${String(msg).replace(/</g, "&lt;").replace(/>/g, "&gt;")}</div>
                </div>
                <div class="btnRow">
                  <button class="btnSecondary" onclick="google.script.host.close()">Close</button>
                </div>
              \`;
            })
            .runSyncBoqLayerToMasterForUi();
        </script>
      </body>
    </html>
  `).setWidth(520).setHeight(320);

  SpreadsheetApp.getUi().showModelessDialog(html, "AUTO QA");
}

/* =========================================================
   Backend method called from popup
========================================================= */
function runSyncBoqLayerToMasterForUi() {
  const result = syncBoqLayerToMasterCore_();
  return {
    masterCopyId: result.report.debug.masterCopyId || "",
    masterCopyName: result.report.debug.masterCopyName || "",
    masterCopyUrl: result.report.debug.masterCopyUrl || "",
    rowsUpdated: result.report.rowsUpdated || 0,
    rowsInserted: result.report.rowsInserted || 0,
    rowsHidden: result.report.rowsHidden || 0
  };
}

/* =========================================================
   Main sync
========================================================= */
function syncBoqLayerToMaster() {
  const out = syncBoqLayerToMasterCore_();
  if (BOQ_SYNC.SHOW_DIALOG) {
    showReportDialog_(SHEETS.getUi(), out.report);
  }
}

function testFloringCeilingDetection() {
  const testSheets = ["Flooring", "Ceiling", "Civil"];
  for (const tabName of testSheets) {
    const isLocationSplitOnly = false;
    Logger.log(`${tabName}: ${isLocationSplitOnly ? "✅ SPLIT" : "❌ FILL"}`);
  }
}

function syncBoqLayerToMasterCore_() {
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
      masterCopyId: "",
      masterCopyName: "",
      masterCopyUrl: "",
    },

    errors: [],
    notes: [],
  };

  try {
    const mapSS = SHEETS.getActiveSpreadsheet();
    const mapSh = mapSS.getSheetByName(BOQ_SYNC.MAP_TAB);
    if (!mapSh) throw new Error(`Mapping tab not found: ${BOQ_SYNC.MAP_TAB}`);

    const exportSS = SHEETS.openById(BOQ_SYNC.EXPORT_SS_ID);

    const masterCopy = createMasterCopy_();
    const masterSS = SHEETS.openById(masterCopy.id);

    PropertiesService.getScriptProperties().setProperty(
      BOQ_SYNC.LATEST_MASTER_COPY_ID_KEY,
      masterCopy.id
    );

    report.debug.masterCopyId = masterCopy.id;
    report.debug.masterCopyName = masterCopy.name;
    report.debug.masterCopyUrl = masterCopy.url;

    const mappings = readMappings_(mapSh, report);
    report.mappingsTotal = mappings.length;

    if (!mappings.length) {
      report.notes.push("No mappings found in BOQ-LAYER.");
      finalize_(report);
      return { report, masterCopy };
    }

    const mappingByTarget = new Map();
    for (const m of mappings) {
      const tKey = normKey_advanced_(m.targeted);
      if (!mappingByTarget.has(tKey)) mappingByTarget.set(tKey, []);
      mappingByTarget.get(tKey).push(m);
    }

    const qtyIndex = buildQtyIndex_(exportSS, report);
    const measCache = new Map();

    const tabs = getMasterTabsFromStart_(masterSS, BOQ_SYNC.MASTER_START_TAB);
    if (!tabs.length) throw new Error(`Start tab "${BOQ_SYNC.MASTER_START_TAB}" not found in MASTER.`);

    const targetsFoundSomewhere = new Set();

    for (const sh of tabs) {
      const tabName = sh.getName();

      if (BOQ_SYNC.SHEET_WHITELIST && BOQ_SYNC.SHEET_WHITELIST.length > 0) {
        const isInWhitelist = BOQ_SYNC.SHEET_WHITELIST.some(name =>
          name.toLowerCase() === tabName.toLowerCase()
        );
        if (!isInWhitelist) {
          report.tabsSkipped.push(`${tabName} (not in whitelist)`);
          continue;
        }
      }

      const isLocationSplitOnly = false;
      const lastCol = sh.getLastColumn();
      let lastRow = sh.getLastRow();
      if (lastRow < 2 || lastCol < 2) continue;

      report.tabsProcessed++;

      const matchCol =
        findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_SCOPE, 20) ||
        BOQ_SYNC.MASTER_MATCH_COL_FALLBACK;

      const locationCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_LOCATION, 20);
      const measurementCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_MEASUREMENT, 20);
      const heightCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_HEIGHT, 20);

      const qtyColDetected = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_QTY, 20);
      const qtyCol = qtyColDetected || BOQ_SYNC.MASTER_QTY_COL_FALLBACK;

      const totalQtyCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_TOTAL_QTY, 20);
      const srNoCol = findHeaderColContains_(sh, BOQ_SYNC.MASTER_HDR_SRNO, 20);

      report.debug.masterDetectedCols.push(
        `${tabName}: matchCol=${matchCol}, locationCol=${locationCol}, measurementCol=${measurementCol || "null"}, heightCol=${heightCol || "null"}, qtyCol=${qtyCol} (detected=${qtyColDetected || "no"}), totalQtyCol=${totalQtyCol || "null"}, srNoCol=${srNoCol || "null"}, isLocationSplitOnly=${isLocationSplitOnly}`
      );

      if (!locationCol || !qtyCol) {
        report.tabsSkipped.push(`${tabName} (missing LOCATION or QTY column)`);
        continue;
      }

      const scopeVals = sh.getRange(1, matchCol, lastRow, 1).getDisplayValues();

      let r = 1;
      while (r <= scopeVals.length) {
        const scopeText = String((scopeVals[r - 1] && scopeVals[r - 1][0]) || "").trim();
        if (!scopeText) {
          r++;
          continue;
        }

        const cell = sh.getRange(r, matchCol);

        let mergedRange = null;
        if (cell.isPartOfMerge()) {
          const mrs = cell.getMergedRanges();
          mergedRange = (mrs && mrs.length) ? mrs[0] : null;

          if (mergedRange) {
            const isTopLeft =
              mergedRange.getRow() === r && mergedRange.getColumn() === matchCol;
            if (!isTopLeft) {
              r++;
              continue;
            }
          }
        }

        const rowKey = normKey_advanced_(scopeText);
        const mapList = mappingByTarget.get(rowKey);
        if (!mapList || !mapList.length) {
          r++;
          continue;
        }

        report.targetsMatched++;
        targetsFoundSomewhere.add(rowKey);

        const zoneOrder = [];
        const zoneTotalsMeas = new Map();
        const zoneTotalsQty = new Map();

        if (measurementCol) {
          for (const m of mapList) {
            if (!m.generated) continue;

            const cacheKey = `${normKey_advanced_(m.generated)}||${normKey_advanced_(m.measure)}`;
            let breakdown = measCache.get(cacheKey);
            if (!breakdown) {
              breakdown = computeMeasurementFromSource_(exportSS, m);
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

        for (const m of mapList) {
          const candidateQtyKeys = getCandidateQtyKeys_(m);
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

            break;
          }
        }

        if (!zoneOrder.length) {
          r++;
          continue;
        }

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

        if (!finalZones.length) {
          r++;
          continue;
        }

        const needed = finalZones.length;

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

          const blanks = Array.from({ length: needed - 1 }, () => [""]);
          scopeVals.splice(r, 0, ...blanks);

          lastRow += (needed - 1);

          const baseRowRange = sh.getRange(r, 1, 1, lastCol);

          for (let i = 1; i < needed; i++) {
            const newRowRange = sh.getRange(r + i, 1, 1, lastCol);
            baseRowRange.copyTo(newRowRange, { contentsOnly: false });

            const clearUpto = Math.max(1, locationCol - 1);
            sh.getRange(r + i, 1, 1, clearUpto).clearContent();
          }

          mergeNonDetailColumnsForBlock_(
            sh,
            r,
            needed,
            lastCol,
            {
              locationCol,
              measurementCol,
              heightCol,
              qtyCol
            }
          );
        }

        // ✅ decide from BOQ-LAYER Measurement column, not from template header
const writeQtyMode = mapList.some(m => isCountMeasure_(m.measure));

for (let i = 0; i < needed; i++) {
  const rowNum = r + i;
  const zone = finalZones[i];

  sh.getRange(rowNum, locationCol).setValue(zone);

  if (writeQtyMode) {
    const qty = toNumber_(zoneTotalsQty.get(zone) || 0);
    const safeQty = Number.isFinite(qty) ? qty : 0;

    if (qtyCol && !hasFormulaInCell_(sh, rowNum, qtyCol)) {
      const qCell = sh.getRange(rowNum, qtyCol);
      qCell.setValue(safeQty);
      qCell.setNumberFormat(BOQ_SYNC.NUMBER_FORMAT);
    }

    // optional: keep MEASUREMENT blank for Count rows
    if (measurementCol && !hasFormulaInCell_(sh, rowNum, measurementCol)) {
      sh.getRange(rowNum, measurementCol).clearContent();
    }

  } else {
    const meas = toNumber_(zoneTotalsMeas.get(zone) || 0);
    const safeMeas = Number.isFinite(meas) ? meas : 0;

    if (measurementCol && !hasFormulaInCell_(sh, rowNum, measurementCol)) {
      const mCell = sh.getRange(rowNum, measurementCol);
      mCell.setValue(safeMeas);
      mCell.setNumberFormat(BOQ_SYNC.NUMBER_FORMAT);
    }

    // optional: keep QTY blank for non-Count rows
    if (qtyCol && !hasFormulaInCell_(sh, rowNum, qtyCol)) {
      sh.getRange(rowNum, qtyCol).clearContent();
    }
  }
}

writeTotalQtyIfNoFormula_(
  sh,
  r,
  needed,
  totalQtyCol,
  qtyCol,
  BOQ_SYNC.NUMBER_FORMAT
);

        report.rowsUpdated += needed;
        r += needed;
      }

      // if (BOQ_SYNC.HIDE_ZERO_ROWS_AFTER_SYNC) {
      //   const hideRes = hideZeroRowsInTab_(sh, matchCol, measurementCol, qtyCol, locationCol, isLocationSplitOnly);
      //   report.rowsHidden += hideRes.hidden;
      //   report.rowsUnhidden += hideRes.unhidden;
      // }
    }

    let notFoundTargets = 0;
    for (const m of mappings) {
      const tKey = normKey_advanced_(m.targeted);
      if (!targetsFoundSomewhere.has(tKey)) notFoundTargets++;
    }
    report.notFoundTargets = notFoundTargets;

    finalize_(report);
    return { report, masterCopy };

  } catch (e) {
    report.errors.push(String(e && e.stack ? e.stack : e));
    finalize_(report);
    throw e;
  }
}

/* =========================================================
   Create fresh master copy
========================================================= */
function createMasterCopy_() {
  const templateFile = DriveApp.getFileById(BOQ_SYNC.MASTER_TEMPLATE_SS_ID);

  const exportSS = SHEETS.openById(BOQ_SYNC.EXPORT_SS_ID);
  const plannerSh = exportSS.getSheetByName("PLANNER");

  let desiredName = "";
  if (plannerSh) {
    const data = plannerSh.getDataRange().getDisplayValues();
    if (data.length >= 2) {
      const header = data[0].map(h => String(h || "").trim().toLowerCase());
      const idx = header.indexOf("dwg");
      if (idx !== -1) {
        desiredName = String(data[1][idx] || "").trim();
      }
    }
  }

  if (!desiredName) {
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    );
    desiredName = `${templateFile.getName()} - Copy - ${timestamp}`;
  }

  const parents = templateFile.getParents();
  let newFile;

  if (parents.hasNext()) {
    const parentFolder = parents.next();
    newFile = templateFile.makeCopy(desiredName, parentFolder);
  } else {
    newFile = templateFile.makeCopy(desiredName);
  }

  return {
    id: newFile.getId(),
    name: newFile.getName(),
    url: `https://docs.google.com/spreadsheets/d/${newFile.getId()}/edit`,
  };
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
    
    
    // ✅ NEW LOGIC: Check Generated-Block (E) FIRST, then fall back to Generated (D)
    let generated = String(grid[r][BOQ_SYNC.MAP_COL_GENERATED_BLOCK - 1] || "").trim();
    if (!generated) {
      generated = String(grid[r][BOQ_SYNC.MAP_COL_GENERATED - 1] || "").trim();
    }
    
    const blockName = String(grid[r][BOQ_SYNC.MAP_COL_BLOCKNAME - 1] || "").trim();

    if (!targeted) continue;
    if (!generated && !blockName) continue;

    let measure = "Area";
    if (BOQ_SYNC.MAP_COL_MEASURE) {
      measure = String(grid[r][BOQ_SYNC.MAP_COL_MEASURE - 1] || "").trim() || "Area";
    }

    out.push({
      targeted,
      generated,
      blockName,
      measure,
    });
  }

  return out;
}

function getCandidateQtyKeys_(m) {
  const keys = [];

  // ✅ 1. HIGHEST PRIORITY → Generated Block (Column E)
  if (m.generated) {
    keys.push(normKey_advanced_(m.generated));
  }

  // ✅ 2. Block Name (fallback)
  if (m.blockName) {
    keys.push(normKey_advanced_(m.blockName));
  }

  // ❌ DO NOT USE TARGETED FOR QTY MATCHING (causes mismatch)
  // if (m.targeted) keys.push(normKey_advanced_(m.targeted));

  // remove duplicates
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
   Build QTY index
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

  // ✅ carry forward last non-blank Product / BOQ name
  let currentProduct = "";
  let currentBoqName = "";

  for (let r = 1; r < dr.length; r++) {
    const rowProduct = idxProduct !== -1 ? String(dr[r][idxProduct] || "").trim() : "";
    const rowBoqName = idxBoq !== -1 ? String(dr[r][idxBoq] || "").trim() : "";

    if (rowProduct) currentProduct = rowProduct;
    if (rowBoqName) currentBoqName = rowBoqName;

    const zone = String(dr[r][idxZone] || "").trim() || "misc";
    const qty = toNumber_(dr[r][idxQty]);
    if (!Number.isFinite(qty)) continue;

    if (currentProduct) addToIndex(currentProduct, zone, qty);
    if (currentBoqName) addToIndex(currentBoqName, zone, qty);
  }

  report.debug.qtyUniqueBoqKeys = index.size;

  let i = 0;
  for (const [k, rec] of index.entries()) {
    report.debug.sampleQtyBoqKeys.push(`${k} => [${rec.order.join(" | ")}]`);
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
function computeMeasurementFromSource_(exportSS, mappingObj) {
  const measureType = normKey_advanced_(mappingObj.measure || "area");

  // 1) try vis_export_sheet_like first
  const visBreakdown = computeMeasurementFromVisExport_(exportSS, mappingObj, measureType);
  if (visBreakdown && visBreakdown.order && visBreakdown.order.length) {
    return visBreakdown;
  }

  // 2) fallback to LAYER
  return computeMeasurementFromLayer_(exportSS, mappingObj.generated, measureType);
}

function computeMeasurementFromLayer_(exportSS, layerName, measure) {
  const sh = exportSS.getSheetByName(BOQ_SYNC.EXPORT_MEASURE_TAB);
  if (!sh) return { order: [], byZone: new Map() };

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { order: [], byZone: new Map() };

  const header = values[0].map(v => String(v || "").trim().toLowerCase());

  const idxLayer = header.indexOf(BOQ_SYNC.LAYER_HDR_LAYER.toLowerCase());
  const idxZone = header.indexOf(BOQ_SYNC.LAYER_HDR_ZONE.toLowerCase());
  const idxArea = header.indexOf(BOQ_SYNC.LAYER_HDR_AREA.toLowerCase());
  const idxPerim = header.indexOf(BOQ_SYNC.LAYER_HDR_PERIM.toLowerCase());
  const idxLen = header.indexOf(BOQ_SYNC.LAYER_HDR_LENGTH.toLowerCase());

  if (idxLayer === -1 || idxZone === -1) {
    return { order: [], byZone: new Map() };
  }

  const m = normKey_advanced_(measure);
  let idxMeasure = idxArea;
  let isCount = false;

  if (m === "perimeter") idxMeasure = idxPerim;
  else if (m === "length") idxMeasure = idxLen;
  else if (m === "count") isCount = true;
  else idxMeasure = idxArea;

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

    let add = NaN;

    if (isCount) {
      add = 1;
    } else if (idxMeasure !== -1) {
      add = toNumber_(values[r][idxMeasure]);
    }

    if (!Number.isFinite(add)) continue;

    byZone.set(zone, (byZone.get(zone) || 0) + add);
  }

  return { order, byZone };
}

function computeMeasurementFromVisExport_(exportSS, mappingObj, measureType) {
  const sh = exportSS.getSheetByName(BOQ_SYNC.EXPORT_QTY_TAB);
  if (!sh) return { order: [], byZone: new Map() };

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { order: [], byZone: new Map() };

  const header = values[0].map(v => String(v || "").trim().toLowerCase());

  const idxProduct = findHeaderIndexContains_(header, [BOQ_SYNC.VIS_HDR_PRODUCT]);
  const idxBoq = findHeaderIndexContains_(header, [BOQ_SYNC.VIS_HDR_BOQ]);
  const idxZone = findHeaderIndexContains_(header, [BOQ_SYNC.VIS_HDR_ZONE]);
  const idxQty = findHeaderIndexContains_(header, [BOQ_SYNC.VIS_HDR_QTY]);
  const idxLength = findHeaderIndexContains_(header, [BOQ_SYNC.VIS_HDR_LENGTH]);

  // optional columns for future-proofing
  const idxArea = findHeaderIndexContains_(header, ["area (ft2)", "area"]);
  const idxPerim = findHeaderIndexContains_(header, ["perimeter"]);

  if (idxZone === -1) return { order: [], byZone: new Map() };

  const candidateKeys = getCandidateQtyKeys_(mappingObj);

  const order = [];
  const seen = new Set();
  const byZone = new Map();

  let currentProduct = "";
  let currentBoqName = "";

  for (let r = 1; r < values.length; r++) {
    const rowProduct = idxProduct !== -1 ? String(values[r][idxProduct] || "").trim() : "";
    const rowBoqName = idxBoq !== -1 ? String(values[r][idxBoq] || "").trim() : "";

    if (rowProduct) currentProduct = rowProduct;
    if (rowBoqName) currentBoqName = rowBoqName;

    const currentKeys = [
      normKey_advanced_(currentProduct),
      normKey_advanced_(currentBoqName)
    ].filter(Boolean);

    const isMatch = candidateKeys.some(k => currentKeys.includes(k));
    if (!isMatch) continue;

    const zone = String(values[r][idxZone] || "").trim() || "misc";

    if (!seen.has(zone)) {
      seen.add(zone);
      order.push(zone);
    }

    let add = NaN;

    if (measureType === "count") {
      if (idxQty !== -1) add = toNumber_(values[r][idxQty]);
    } else if (measureType === "length") {
      if (idxLength !== -1) add = toNumber_(values[r][idxLength]);
    } else if (measureType === "area") {
      if (idxArea !== -1) add = toNumber_(values[r][idxArea]);
    } else if (measureType === "perimeter") {
      if (idxPerim !== -1) add = toNumber_(values[r][idxPerim]);
    }

    if (!Number.isFinite(add)) continue;

    byZone.set(zone, (byZone.get(zone) || 0) + add);
  }

  return { order, byZone };
}

/* =========================================================
   HIDE rows logic
========================================================= */
// function hideZeroRowsInTab_(sh, matchCol, measurementCol, qtyCol, locationCol, isLocationSplitOnly) {
//   const lastRow = sh.getLastRow();
//   if (lastRow < 2) return { hidden: 0, unhidden: 0 };

//   const qtyVals = sh.getRange(1, qtyCol, lastRow, 1).getValues();
//   const measVals = measurementCol ? sh.getRange(1, measurementCol, lastRow, 1).getValues() : null;
//   const srVals = sh.getRange(1, 1, lastRow, 1).getDisplayValues();

//   let hidden = 0;
//   let unhidden = 0;

//   for (let r = 2; r <= lastRow; r++) {
//     const scope = String(sh.getRange(r, matchCol).getDisplayValue() || "").trim();
//     if (!scope) continue;

//     const cell = sh.getRange(r, matchCol);
//     if (cell.isPartOfMerge()) {
//       const mrs = cell.getMergedRanges();
//       if (mrs && mrs.length) {
//         const mergedRange = mrs[0];
//         const isTopLeft = mergedRange.getRow() === r && mergedRange.getColumn() === matchCol;
//         if (!isTopLeft) {
//           continue;
//         }
//       }
//     }

//     const srText = String((srVals[r - 1] && srVals[r - 1][0]) || "").trim();
//     if (isAlphaSerialRow_(srText)) {
//       if (sh.isRowHiddenByUser(r)) {
//         sh.showRows(r);
//         unhidden++;
//       }
//       continue;
//     }

//     if (isProtectedTotalRow_(scope)) {
//       if (sh.isRowHiddenByUser(r)) {
//         sh.showRows(r);
//         unhidden++;
//       }
//       continue;
//     }

//     if (isLocationSplitOnly && locationCol) {
//       const location = String(sh.getRange(r, locationCol).getDisplayValue() || "").trim();

//       const scopeCell = sh.getRange(r, matchCol);
//       const isMergedParent = scopeCell.isPartOfMerge() &&
//         scopeCell.getMergedRanges()[0].getRow() === r;

//       if (isMergedParent) {
//         if (sh.isRowHiddenByUser(r)) {
//           sh.showRows(r);
//           unhidden++;
//         }
//       } else if (!location) {
//         if (!sh.isRowHiddenByUser(r)) {
//           sh.hideRows(r);
//           hidden++;
//         }
//       } else {
//         if (sh.isRowHiddenByUser(r)) {
//           sh.showRows(r);
//           unhidden++;
//         }
//       }
//       continue;
//     }

//     const qty = toNumber_(qtyVals[r - 1][0]);
//     const qtyZero = !Number.isFinite(qty) || qty === 0;

//     if (!measurementCol) {
//       if (qtyZero) {
//         if (!sh.isRowHiddenByUser(r)) {
//           sh.hideRows(r);
//           hidden++;
//         }
//       } else {
//         if (sh.isRowHiddenByUser(r)) {
//           sh.showRows(r);
//           unhidden++;
//         }
//       }
//       continue;
//     }

//     const meas = toNumber_(measVals[r - 1][0]);
//     const measZero = !Number.isFinite(meas) || meas === 0;

//     if (qtyZero && measZero) {
//       if (!sh.isRowHiddenByUser(r)) {
//         sh.hideRows(r);
//         hidden++;
//       }
//     } else {
//       if (sh.isRowHiddenByUser(r)) {
//         sh.showRows(r);
//         unhidden++;
//       }
//     }
//   }

//   return { hidden, unhidden };
// }

// function isProtectedTotalRow_(scopeText) {
//   const s = normKey_advanced_(scopeText);
//   return s === "total" || s === "sheet total";
// }

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
  Logger.log(`[MERGE] Attempting: Row=${startRow}, Col=${col}, NumRows=${numRows}`);

  if (numRows <= 1) {
    Logger.log(`[MERGE] SKIPPED: NumRows=${numRows} (need > 1)`);
    return;
  }

  const rng = sh.getRange(startRow, col, numRows, 1);

  if (rng.isPartOfMerge()) {
    const mrs = rng.getMergedRanges();
    for (const mr of mrs) mr.breakApart();
  }

  rng.merge();
  Logger.log(`[MERGE] SUCCESS: Merged Row ${startRow} to ${startRow + numRows - 1}`);

  rng.setHorizontalAlignment("center");
  rng.setVerticalAlignment("middle");
}

function mergeNonDetailColumnsForBlock_(sh, startRow, numRows, lastCol, exemptColsObj) {
  if (numRows <= 1) return;

  const exemptCols = new Set(
    Object.values(exemptColsObj).filter(c => Number.isInteger(c) && c > 0)
  );

  for (let col = 1; col <= lastCol; col++) {
    if (exemptCols.has(col)) continue;

    const rng = sh.getRange(startRow, col, numRows, 1);

    if (rng.isPartOfMerge()) {
      const mergedRanges = rng.getMergedRanges();
      for (const mr of mergedRanges) {
        mr.breakApart();
      }
    }

    rng.merge();
    rng.setVerticalAlignment("middle");
    rng.setHorizontalAlignment("center");
  }
}

function hasFormulaInCell_(sh, row, col) {
  if (!col || row < 1) return false;
  const formula = sh.getRange(row, col).getFormula();
  return !!(formula && String(formula).trim());
}

function writeTotalQtyIfNoFormula_(sh, startRow, numRows, totalQtyCol, qtyCol, numberFormat) {
  if (!totalQtyCol || !qtyCol || startRow < 1 || numRows < 1) return;

  const totalCell = sh.getRange(startRow, totalQtyCol);

  // do not overwrite formula
  if (hasFormulaInCell_(sh, startRow, totalQtyCol)) return;

  const qtyValues = sh.getRange(startRow, qtyCol, numRows, 1).getValues();

  let total = 0;
  for (let i = 0; i < qtyValues.length; i++) {
    const n = toNumber_(qtyValues[i][0]);
    if (Number.isFinite(n)) total += n;
  }

  totalCell.setValue(total);
  totalCell.setNumberFormat(numberFormat || "0.############");
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
  lines.push(`Master Copy Name: ${report.debug.masterCopyName || "-"}`);
  lines.push(`Master Copy ID: ${report.debug.masterCopyId || "-"}`);
  lines.push(`Master Copy URL: ${report.debug.masterCopyUrl || "-"}`);
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
   Helper for downstream scripts
========================================================= */
function getLatestMasterCopyId_() {
  const id = PropertiesService.getScriptProperties().getProperty(
    BOQ_SYNC.LATEST_MASTER_COPY_ID_KEY
  );
  if (!id) throw new Error("No latest generated master copy ID found.");
  return id;
}

/* =========================================================
   Utils
========================================================= */
function isCountMeasure_(measure) {
  return normKey_advanced_(measure) === "count";
}
function pushZone_(arr, zone) {
  const z = (zone && String(zone).trim()) ? String(zone).trim() : "misc";
  if (!arr.includes(z)) arr.push(z);
}

function isAlphaSerialRow_(value) {
  const s = String(value || "").trim();
  if (!s) return false;

  return /^[A-Za-z]+(?:[.\-]?\d+)?$/.test(s);
}

function toNumber_(v) {
  if (v == null || v === "") return NaN;
  const n = Number(v);
  return Number.isFinite(n) ? n : NaN;
}

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