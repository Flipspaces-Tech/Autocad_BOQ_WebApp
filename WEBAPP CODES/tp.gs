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
 *
 * CONFIG POST-SYNC:
 * Furniture sheet CONFIGURATION column is overwritten after main sync.
 * Mapping:
 * Generated Furniture Scope of Work (col B)
 * -> BOQ-LAYER Targeted (col B)
 * -> BOQ-LAYER Generated-Block (col E) [fallback Generated col D]
 * -> PerInstanceData BOQ name (col B)
 * -> same row Config
 * -> clean with Config sheet
 * -> write to Furniture col D location-wise
 *
 * DYNAMIC NORMALIZATION:
 * Add a sheet named NORMALIZATION_RULES in the same spreadsheet where script is hosted.
 * Columns:
 *   A = Type
 *   B = Find
 *   C = Replace
 *
 * Example rows:
 *   name | wks        | workstation
 *   name | ws         | workstation
 *   name | conf       | conference
 *   zone | work hall  | workhall
 **************************************************/

const SHEETS = SpreadsheetApp;

const BOQ_SYNC = {
  // Mapping sheet (script hosted here)
  MAP_TAB: "BOQ-LAYER",

  // Mapping columns (1-based)
  MAP_COL_TARGETED: 2,
  MAP_COL_GENERATED: 4,
  MAP_COL_BLOCKNAME: 6,
  MAP_COL_GENERATED_BLOCK: 5,
  MAP_COL_MEASURE: 7,

  // Export spreadsheet
  EXPORT_SS_ID: "12AsC0b7_U4dxhfxEZwtrwOXXALAnEEkQm5N8tg_RByM",

  SHEET_WHITELIST: null,

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
  QTY_HDR_QTYVALUE_PATTERNS: ["qty_value", "qty value", "qty", "quantity", "qty_value"],

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

  LATEST_MASTER_COPY_ID_KEY: "LATEST_MASTER_COPY_ID",

  // Dynamic normalization
  NORMALIZATION_TAB: "NORMALIZATION_RULES",
  NORMALIZATION_HDR_TYPE: "type",
  NORMALIZATION_HDR_FIND: "find",
  NORMALIZATION_HDR_REPLACE: "replace",
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

    const normalizationRules = readNormalizationRules_(mapSS, report);

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
    }

    let notFoundTargets = 0;
    for (const m of mappings) {
      const tKey = normKey_advanced_(m.targeted);
      if (!targetsFoundSomewhere.has(tKey)) notFoundTargets++;
    }
    report.notFoundTargets = notFoundTargets;

    syncFurnitureConfigurationAfterMainSync_(exportSS, masterSS, mappings, normalizationRules, report);

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

  if (m.generated) {
    keys.push(normKey_advanced_(m.generated));
  }

  if (m.blockName) {
    keys.push(normKey_advanced_(m.blockName));
  }

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
   Dynamic normalization rules
========================================================= */
function readNormalizationRules_(mapSS, report) {
  const sh = mapSS.getSheetByName(BOQ_SYNC.NORMALIZATION_TAB);
  const out = {
    name: [],
    zone: []
  };

  if (!sh) {
    report.notes.push(`Normalization sheet "${BOQ_SYNC.NORMALIZATION_TAB}" not found. Using generic normalization only.`);
    return out;
  }

  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return out;

  const header = values[0].map(v => String(v || "").trim().toLowerCase());
  const idxType = header.indexOf(BOQ_SYNC.NORMALIZATION_HDR_TYPE.toLowerCase());
  const idxFind = header.indexOf(BOQ_SYNC.NORMALIZATION_HDR_FIND.toLowerCase());
  const idxReplace = header.indexOf(BOQ_SYNC.NORMALIZATION_HDR_REPLACE.toLowerCase());

  if (idxType === -1 || idxFind === -1 || idxReplace === -1) {
    report.notes.push(`Normalization sheet "${BOQ_SYNC.NORMALIZATION_TAB}" is missing Type/Find/Replace headers. Using generic normalization only.`);
    return out;
  }

  for (let r = 1; r < values.length; r++) {
    const type = String(values[r][idxType] || "").trim().toLowerCase();
    const find = String(values[r][idxFind] || "").trim().toLowerCase();
    const replace = String(values[r][idxReplace] || "").trim().toLowerCase();

    if (!type || !find) continue;
    if (type !== "name" && type !== "zone") continue;

    out[type].push({ find, replace });
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
   Measurement from source
========================================================= */
function computeMeasurementFromSource_(exportSS, mappingObj) {
  const measureType = normKey_advanced_(mappingObj.measure || "area");

  const visBreakdown = computeMeasurementFromVisExport_(exportSS, mappingObj, measureType);
  if (visBreakdown && visBreakdown.order && visBreakdown.order.length) {
    return visBreakdown;
  }

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
  if (numRows <= 1) return;

  const rng = sh.getRange(startRow, col, numRows, 1);

  if (rng.isPartOfMerge()) {
    const mrs = rng.getMergedRanges();
    for (const mr of mrs) mr.breakApart();
  }

  rng.merge();
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

  if (report.notes.length) {
    lines.push("Notes:");
    for (const n of report.notes) lines.push(` - ${n}`);
    lines.push("");
  }

  if (report.errors.length) {
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
   FURNITURE CONFIGURATION POST-SYNC
========================================================= */
function syncFurnitureConfigurationAfterMainSync_(exportSS, masterSS, mappings, normalizationRules, report) {
  const furnitureSh = masterSS.getSheetByName("Furniture");
  if (!furnitureSh) {
    report.notes.push('Furniture sheet not found in generated master. CONFIGURATION sync skipped.');
    return;
  }

  const perInstSh = exportSS.getSheetByName("PerInstanceData");
  if (!perInstSh) {
    report.notes.push('PerInstanceData sheet not found in export file. CONFIGURATION sync skipped.');
    return;
  }

  const configRulesSh = exportSS.getSheetByName("Config");
  if (!configRulesSh) {
    report.notes.push('Config sheet not found in export file. CONFIGURATION sync skipped.');
    return;
  }

  const ruleMap = buildFurnitureConfigRuleMap_(configRulesSh);
  const perInstanceIndex = buildPerInstanceFurnitureConfigIndex_ByBoqName_(perInstSh, ruleMap, normalizationRules, report);

  if (!perInstanceIndex.size) {
    report.notes.push('No cleaned config data built from PerInstanceData. CONFIGURATION sync skipped.');
    return;
  }

  const lastRow = furnitureSh.getLastRow();
  const lastCol = furnitureSh.getLastColumn();
  if (lastRow < 2) {
    report.notes.push('Furniture sheet has no data rows. CONFIGURATION sync skipped.');
    return;
  }

  const scopeCol = 2;    // B
  const configCol = 4;   // D
  const locationCol = 7; // G

  let r = 2;
  let updatedCount = 0;
  let splitCount = 0;

  while (r <= furnitureSh.getLastRow()) {
    const scopeText = String(furnitureSh.getRange(r, scopeCol).getDisplayValue() || "").trim();
    if (!scopeText) {
      r++;
      continue;
    }

    const blockInfo = getFurnitureBlockRange_(furnitureSh, r, scopeCol);
    const startRow = blockInfo.startRow;
    const numRows = blockInfo.numRows;

    const mapList = mappings.filter(m =>
      canonicalNameForConfigMatch_(m.targeted, normalizationRules) === canonicalNameForConfigMatch_(scopeText, normalizationRules)
    );

    if (!mapList.length) {
      r = startRow + numRows;
      continue;
    }

    const locVals = furnitureSh.getRange(startRow, locationCol, numRows, 1).getDisplayValues();
    const existingLocations = [];
    for (let i = 0; i < locVals.length; i++) {
      const z = String(locVals[i][0] || "").trim();
      if (z) existingLocations.push(z);
    }

    const extraLocations = getPerInstanceLocationsForMapList_(mapList, perInstanceIndex, normalizationRules);

    const allLocations = [];
    const seenLoc = new Set();

    for (const z of existingLocations.concat(extraLocations)) {
      const k = canonicalZoneForConfigMatch_(z, normalizationRules);
      if (!k || seenLoc.has(k)) continue;
      seenLoc.add(k);
      allLocations.push(z);
    }

    const desiredRows = [];
    for (const location of allLocations) {
      const finalConfig = findFurnitureConfigForScopeAndLocation_(
        mapList,
        location,
        perInstanceIndex,
        normalizationRules
      );

      desiredRows.push({
        location,
        config: finalConfig || ""
      });
    }

    const dedupedRows = dedupeLocationConfigRows_(desiredRows, normalizationRules);

    if (!dedupedRows.length) {
      r = startRow + numRows;
      continue;
    }

    const uniqueConfigs = Array.from(
      new Set(dedupedRows.map(x => normKey_advanced_(x.config)))
    ).filter(Boolean);

    if (uniqueConfigs.length <= 1) {
      const configToWrite = dedupedRows[0].config || "";
      if (configToWrite) {
        furnitureSh.getRange(startRow, configCol).setValue(configToWrite);
        furnitureSh.getRange(startRow, configCol).setWrap(true);
        updatedCount++;
      }
      r = startRow + numRows;
      continue;
    }

    splitFurnitureBlockByConfig_(
      furnitureSh,
      startRow,
      numRows,
      lastCol,
      locationCol,
      configCol,
      dedupedRows
    );

    updatedCount += dedupedRows.length;
    splitCount++;
    r = startRow + dedupedRows.length;
  }

  report.notes.push(`Furniture CONFIGURATION updated for ${updatedCount} row(s).`);
  report.notes.push(`Furniture blocks split by config difference: ${splitCount}.`);
}

function findFurnitureConfigForScopeAndLocation_(mapList, rowLocation, perInstanceIndex, normalizationRules) {
  for (const m of mapList) {
    const lookupKeys = getConfigLookupKeysFromMapping_(m, normalizationRules);

    for (const lookupKey of lookupKeys) {
      const boqKey = canonicalNameForConfigMatch_(lookupKey, normalizationRules);
      if (!boqKey) continue;

      const zoneMap = perInstanceIndex.get(boqKey);
      if (!zoneMap) continue;

      const locationKey = canonicalZoneForConfigMatch_(rowLocation, normalizationRules);

      let blocks = zoneMap.get(locationKey);

      if ((!blocks || !blocks.length) && zoneMap.has("misc")) {
        blocks = zoneMap.get("misc");
      }

      if ((!blocks || !blocks.length) && zoneMap.size === 1) {
        const firstKey = Array.from(zoneMap.keys())[0];
        blocks = zoneMap.get(firstKey) || [];
      }

      if (!blocks || !blocks.length) continue;

      const finalConfig = dedupeConfigBlocks_(blocks);
      if (finalConfig) return finalConfig;
    }
  }

  return "";
}

function getPerInstanceLocationsForMapList_(mapList, perInstanceIndex, normalizationRules) {
  const out = [];
  const seen = new Set();

  for (const m of mapList) {
    const lookupKeys = getConfigLookupKeysFromMapping_(m, normalizationRules);

    for (const lookupKey of lookupKeys) {
      const boqKey = canonicalNameForConfigMatch_(lookupKey, normalizationRules);
      if (!boqKey) continue;

      const zoneMap = perInstanceIndex.get(boqKey);
      if (!zoneMap) continue;

      for (const zoneKey of zoneMap.keys()) {
        if (!zoneKey || zoneKey === "misc") continue;
        if (seen.has(zoneKey)) continue;

        seen.add(zoneKey);
        out.push(zoneKey);
      }
    }
  }

  return out;
}

function dedupeLocationConfigRows_(rows, normalizationRules) {
  const out = [];
  const seen = new Set();

  for (const row of rows) {
    const locKey = canonicalZoneForConfigMatch_(row.location, normalizationRules);
    if (!locKey || seen.has(locKey)) continue;
    seen.add(locKey);

    out.push({
      location: row.location,
      config: row.config
    });
  }

  return out;
}

function getFurnitureBlockRange_(sh, rowNum, scopeCol) {
  const cell = sh.getRange(rowNum, scopeCol);

  if (cell.isPartOfMerge()) {
    const mergedRanges = cell.getMergedRanges();
    if (mergedRanges && mergedRanges.length) {
      const mr = mergedRanges[0];
      return {
        startRow: mr.getRow(),
        numRows: mr.getNumRows()
      };
    }
  }

  return {
    startRow: rowNum,
    numRows: 1
  };
}

function splitFurnitureBlockByConfig_(sh, startRow, oldNumRows, lastCol, locationCol, configCol, rowsData) {
  const needed = rowsData.length;

  const blockRange = sh.getRange(startRow, 1, oldNumRows, lastCol);
  if (blockRange.isPartOfMerge()) {
    const mergedRanges = blockRange.getMergedRanges();
    for (const mr of mergedRanges) {
      mr.breakApart();
    }
  }

  if (needed > oldNumRows) {
    sh.insertRowsAfter(startRow + oldNumRows - 1, needed - oldNumRows);

    const baseRowRange = sh.getRange(startRow, 1, 1, lastCol);
    for (let i = oldNumRows; i < needed; i++) {
      const newRowRange = sh.getRange(startRow + i, 1, 1, lastCol);
      baseRowRange.copyTo(newRowRange, { contentsOnly: false });
    }
  } else if (needed < oldNumRows) {
    sh.deleteRows(startRow + needed, oldNumRows - needed);
  }

  for (let i = 0; i < needed; i++) {
    const rowNum = startRow + i;
    sh.getRange(rowNum, locationCol).setValue(rowsData[i].location);
    sh.getRange(rowNum, configCol).setValue(rowsData[i].config);
    sh.getRange(rowNum, configCol).setWrap(true);
  }

  mergeNonDetailColumnsForBlock_(
    sh,
    startRow,
    needed,
    lastCol,
    {
      locationCol: locationCol,
      measurementCol: null,
      heightCol: null,
      qtyCol: configCol
    }
  );
}

function getConfigLookupKeysFromMapping_(m, normalizationRules) {
  const out = [];
  const seen = new Set();

  function push(v) {
    const s = String(v || "").trim();
    if (!s) return;
    const k = canonicalNameForConfigMatch_(s, normalizationRules);
    if (!k || seen.has(k)) return;
    seen.add(k);
    out.push(s);
  }

  push(m.generated);
  push(m.blockName);

  return out;
}

/* =========================================================
   BUILD PerInstanceData INDEX USING COLUMN B = BOQ name
========================================================= */
function buildPerInstanceFurnitureConfigIndex_ByBoqName_(perInstSh, ruleMap, normalizationRules, report) {
  const values = perInstSh.getDataRange().getDisplayValues();
  if (values.length < 2) return new Map();

  const header = values[0].map(v => String(v || "").trim());

  const idxBoq = findHeaderIndexContains_(header, ["boq name", "boq_name", "name"]);
  const idxZone = findHeaderIndexContains_(header, ["zone", "location"]);
  const idxConfig = findHeaderIndexContains_(header, ["config"]);

  if (idxBoq === -1) {
    report.notes.push('BOQ name column not found in PerInstanceData.');
    return new Map();
  }

  if (idxConfig === -1) {
    report.notes.push('CONFIG column not found in PerInstanceData.');
    return new Map();
  }

  const index = new Map();
  let currentBoqName = "";

  for (let r = 1; r < values.length; r++) {
    const rowBoqName = String(values[r][idxBoq] || "").trim();
    if (rowBoqName) currentBoqName = rowBoqName;

    if (!currentBoqName) continue;

    const zoneRaw = idxZone !== -1 ? String(values[r][idxZone] || "").trim() : "";
    const zone = canonicalZoneForConfigMatch_(zoneRaw || "misc", normalizationRules);

    const rawConfig = String(values[r][idxConfig] || "").trim();
    if (!rawConfig) continue;

    const cleanedBlock = cleanPerInstanceRawConfig_(rawConfig, ruleMap);
    if (!cleanedBlock) continue;

    const boqKey = canonicalNameForConfigMatch_(currentBoqName, normalizationRules);
    if (!boqKey) continue;

    if (!index.has(boqKey)) index.set(boqKey, new Map());

    const zoneMap = index.get(boqKey);
    if (!zoneMap.has(zone)) zoneMap.set(zone, []);

    zoneMap.get(zone).push(cleanedBlock);
  }

  return index;
}

/* =========================================================
   BUILD RULE MAP FROM Config SHEET
========================================================= */
function buildFurnitureConfigRuleMap_(configRulesSh) {
  const values = configRulesSh.getDataRange().getDisplayValues();
  const out = new Map();

  if (values.length < 2) return out;

  for (let r = 1; r < values.length; r++) {
    const rawConfigName = String(values[r][0] || "").trim();
    const dimension = String(values[r][1] || "").trim();
    const comment = String(values[r][2] || "").trim();

    if (!rawConfigName) continue;

    const configKey = normalizeConfigLabelKey_(rawConfigName);

    out.set(configKey, {
      rawConfigName,
      dimension,
      comment
    });
  }

  return out;
}

/* =========================================================
   CLEAN ONE RAW CONFIG STRING
========================================================= */
function cleanPerInstanceRawConfig_(rawText, ruleMap) {
  if (!rawText) return "";

  const parts = splitRawConfigIntoParts_(rawText);
  if (!parts.length) return "";

  const cleanedLines = [];

  for (const part of parts) {
    const parsed = parseOneConfigPart_(part);
    if (!parsed) continue;

    const rawLabel = parsed.label;
    const rawValue = parsed.value;

    const normalizedLabel = normalizeConfigLabelKey_(rawLabel);
    const rule = ruleMap.get(normalizedLabel);
    const fallbackRule = rule || findBestRuleForRawLabel_(normalizedLabel, ruleMap);
    if (!fallbackRule) continue;

    const dim = String(fallbackRule.dimension || "").trim();

    if (normKey_advanced_(dim) === "ignore") {
      continue;
    }

    const formattedLine = formatConvertedConfigLine_(rawValue, dim);
    if (!formattedLine) continue;

    cleanedLines.push(formattedLine);
  }

  return dedupeConfigLines_(cleanedLines).join("\n");
}

function splitRawConfigIntoParts_(rawText) {
  const s = String(rawText || "").trim();
  if (!s) return [];

  const out = [];
  let cur = "";
  let depthParen = 0;

  for (let i = 0; i < s.length; i++) {
    const ch = s[i];

    if (ch === "(") depthParen++;
    if (ch === ")") depthParen = Math.max(0, depthParen - 1);

    if (ch === "," && depthParen === 0) {
      const t = cur.trim();
      if (t) out.push(t);
      cur = "";
      continue;
    }

    cur += ch;
  }

  if (cur.trim()) out.push(cur.trim());

  return out;
}

function parseOneConfigPart_(part) {
  const s = String(part || "").trim();
  if (!s) return null;

  const eqIdx = s.indexOf("=");
  if (eqIdx === -1) return null;

  const left = s.slice(0, eqIdx).trim();
  const right = s.slice(eqIdx + 1).trim();

  if (!left || !right) return null;

  return {
    label: left,
    value: right
  };
}

function normalizeConfigLabelKey_(label) {
  let s = String(label || "")
    .replace(/\u00A0/g, " ")
    .replace(/[–—−]/g, "-")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();

  s = s.replace(/_/g, " ");
  s = s.replace(/\s*-\s*/g, "-");

  if (/^position\d*\s*[xy]?$/.test(s) || /^position\s*[xy]$/.test(s)) return "position";
  if (/^origin$/.test(s)) return "origin";
  if (/^angle\d*$/.test(s)) return "angle";

  let m;

  m = s.match(/^length[- ]?(\d+)$/);
  if (m) return `length ${m[1]}`;

  m = s.match(/^depth[- ]?(\d+)$/);
  if (m) return `depth ${m[1]}`;

  m = s.match(/^width[- ]?(\d+)$/);
  if (m) return `width ${m[1]}`;

  if (/^length$/.test(s)) return "length";
  if (/^depth$/.test(s)) return "depth";
  if (/^width$/.test(s)) return "width";
  if (/^diameter$/.test(s)) return "diameter";
  if (/^height$/.test(s)) return "height";
  if (/^capacity\/visibility$/.test(s)) return "capacity visibility";
  if (/^capacity visibility$/.test(s)) return "capacity visibility";
  if (/^x$/.test(s)) return "capacity visibility";

  return s.replace(/[^a-z0-9 ]+/g, " ").replace(/\s+/g, " ").trim();
}

function findBestRuleForRawLabel_(normalizedLabel, ruleMap) {
  if (!normalizedLabel || !ruleMap || !ruleMap.size) return null;

  if (ruleMap.has(normalizedLabel)) return ruleMap.get(normalizedLabel);

  const compact = normalizedLabel.replace(/\s+/g, " ").trim();

  for (const [key, rule] of ruleMap.entries()) {
    if (key === compact) return rule;
  }

  let m = compact.match(/^(length|depth|width)\s+(\d+)$/);
  if (m) {
    const base = m[1];
    const n = m[2];
    const tries = [
      `${base} ${n}`,
      `${base}${n}`,
      base
    ];
    for (const t of tries) {
      if (ruleMap.has(t)) return ruleMap.get(t);
    }
  }

  if (/^position/.test(compact) && ruleMap.has("position")) return ruleMap.get("position");
  if (/^origin/.test(compact) && ruleMap.has("origin")) return ruleMap.get("origin");
  if (/^angle/.test(compact) && ruleMap.has("angle")) return ruleMap.get("angle");

  return null;
}

function formatConvertedConfigLine_(rawValue, dimensionTag) {
  const dim = String(dimensionTag || "").trim();
  if (!dim) return "";

  const dimNorm = normKey_advanced_(dim);
  if (dimNorm === "ignore") return "";

  const num = extractFirstNumber_(rawValue);
  const dimLabel = dim.replace(/\s+/g, "");

  if (Number.isFinite(num)) {
    const rounded = Math.round(num);
    return `${rounded}mm${dimLabel}`;
  }

  const fallback = String(rawValue || "").trim();
  if (!fallback) return "";

  return `${fallback}${dimLabel}`;
}

function dedupeConfigBlocks_(blocks) {
  const out = [];
  const seen = new Set();

  for (const block of blocks) {
    const lines = String(block || "")
      .split("\n")
      .map(s => String(s || "").trim())
      .filter(Boolean);

    for (const line of lines) {
      const k = normKey_advanced_(line);
      if (!k || seen.has(k)) continue;
      seen.add(k);
      out.push(line);
    }
  }

  return out.join(" x\n");
}

function dedupeConfigLines_(lines) {
  const out = [];
  const seen = new Set();

  for (const line of lines) {
    const t = String(line || "").trim();
    if (!t) continue;

    const k = normKey_advanced_(t);
    if (seen.has(k)) continue;

    seen.add(k);
    out.push(t);
  }

  return out;
}

function normalizeZoneKey_(zone) {
  const z = String(zone || "").trim();
  return z ? normKey_advanced_(z) : "misc";
}

function extractFirstNumber_(v) {
  const s = String(v || "").trim();
  if (!s) return NaN;

  const m = s.match(/-?\d+(?:\.\d+)?/);
  if (!m) return NaN;

  const n = Number(m[0]);
  return Number.isFinite(n) ? n : NaN;
}

function trimTrailingZeros_(n) {
  if (!Number.isFinite(n)) return "";
  return String(Number(n));
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

function canonicalNameForConfigMatch_(s, normalizationRules) {
  let t = String(s || "").toLowerCase().trim();

  t = t.replace(/[_]/g, " ");
  t = t.replace(/[–—−-]/g, " ");
  t = t.replace(/[()]/g, " ");
  t = t.replace(/&/g, " and ");
  t = t.replace(/[^a-z0-9 ]+/g, " ");
  t = t.replace(/\s+/g, " ").trim();

  const rules = (normalizationRules && normalizationRules.name) || [];
  for (const rule of rules) {
    if (!rule.find) continue;
    const escaped = escapeRegex_(rule.find);
    t = t.replace(new RegExp(`\\b${escaped}\\b`, "g"), rule.replace || "");
    t = t.replace(/\s+/g, " ").trim();
  }

  return t;
}

function canonicalZoneForConfigMatch_(s, normalizationRules) {
  let t = String(s || "").toLowerCase().trim();

  t = t.replace(/[_]/g, " ");
  t = t.replace(/[–—−-]/g, " ");
  t = t.replace(/[^a-z0-9 ]+/g, " ");
  t = t.replace(/\s+/g, " ").trim();

  const rules = (normalizationRules && normalizationRules.zone) || [];
  for (const rule of rules) {
    if (!rule.find) continue;
    const escaped = escapeRegex_(rule.find);
    t = t.replace(new RegExp(`\\b${escaped}\\b`, "g"), rule.replace || "");
    t = t.replace(/\s+/g, " ").trim();
  }

  return t || "misc";
}

function escapeRegex_(s) {
  return String(s || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
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