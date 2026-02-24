// google apps script

function getSheetDataByName(name, sheet_type=null) {
    var ss = null
    const masterSheetId = "1X-ImBMhJnsR4MQZYhoMB-fOJ6UqLqNg3NbViz7vWZDI"
    if (sheet_type == "master") {
        ss = SpreadsheetApp.openById(masterSheetId)
    }
    else {
        ss = SpreadsheetApp.getActiveSpreadsheet();
    }
    Logger.log("Reading sheet - " + ss.getName());
  var sheet = ss.getSheetByName(name);
  var data = sheet.getDataRange().getValues();
  return data;
}

function getOrCreateSubSheetByName(name) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(name);
    if (sheet == null) {
        sheet = ss.insertSheet(name);
    }
    return sheet;
}

function appendToSheet(sheetName, row) {
  var sheet = getOrCreateSubSheetByName(sheetName)
  sheet.appendRow(row);
}

function insertRowAfterCurrentRowToSheet(sheetName, row, currentRow) {
  var sheet = getOrCreateSubSheetByName(sheetName)
  sheet.insertRowAfter(currentRow);
  sheet.getRange(currentRow + 1, 1, 1, row.length).setValues([row]);
}

function getSrNoValType(val){
    const floatVal = parseFloat(val)
    if (isNaN(floatVal)) {
        if (val.length == 1) {
            return "main"
        }
        else if (val.length == 2) {
            return "sub"
        }
    }
    else {
        return "number"
    }
    return "unknown"
}

function getSubSectionOptions(subSection, masterSubSheet){
  Logger.log("getSubSectionOptions inputs - " + subSection + "  - "+ masterSubSheet)
    const masterData = getSheetDataByName(masterSubSheet, "master")
    const data = []
    let recordData = false
    masterData.forEach(function(row){
        if (row[0] == subSection) {
            recordData = true
        }
        else if (recordData && getSrNoValType(row[0]) == "number") {
            data.push(row)
        }
        else {
            recordData = false
        }
    });
    return data.reverse();
}

function copySubSheetFromMasterToCurrent(subSheetName){
    const existingSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(subSheetName);
    const masterSheetId = "1X-ImBMhJnsR4MQZYhoMB-fOJ6UqLqNg3NbViz7vWZDI";
    var sheetToCopy = SpreadsheetApp.openById(masterSheetId).getSheetByName(subSheetName);
    var destinationSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    if (existingSheet) {
        Logger.log("Sheet already exists - " + subSheetName)
        destinationSpreadsheet.deleteSheet(existingSheet);
    }
    sheetToCopy.copyTo(destinationSpreadsheet).setName(subSheetName);
    Logger.log("Sheet " + subSheetName + " copied successfully.");
}

function iterateInputSheet(startPoint=null) {
    Logger.log("running iterateInputSheet with start point- " + startPoint)
  const inputSheet = "Sheet7"
  const data = getSheetDataByName(inputSheet)
  const srNoCol = 1
  const headerCol = 2
  let mainSheet = null
  let mainSubSheet = null
  let mainSheetName = null
  const selectedRows = {};
    if (startPoint != null && startPoint >= data.length) {
        return
    }
  for (var i=0; i<data.length; i++) {
    
    var row = data[i]
    var srNoValType = getSrNoValType(row[srNoCol])

    if (srNoValType == "main") {
        mainSheet = row[srNoCol]; mainSubSheet = null; mainSheetName = row[headerCol];
        selectedRows[mainSheet] = row[0] == true;
    }
    else if (srNoValType == "sub") {
        mainSubSheet = row[srNoCol];
        selectedRows[mainSubSheet] = row[0] == true || selectedRows[mainSheet];
        if (startPoint != null && i < startPoint) {
            continue
        }
        if (selectedRows[mainSubSheet]) {
            Logger.log("Got a selected Row- " + mainSubSheet + " - " + mainSheetName + " - " + mainSheet)
            getSubSectionOptions(mainSubSheet, mainSheetName).forEach(function(optionRow){
                insertRowAfterCurrentRowToSheet(inputSheet, [false, ...optionRow.slice(0,2)], i+1);
            });
            break;
        }
    }
  }
  iterateInputSheet(i+1)
}

function getCostTypeToColMapping(costType, costCol){
    const floatVal = parseFloat(costType[1])
    return costCol + floatVal - 1
}
function getDetailTypeToColMapping(detailType, detailsCol){
    const floatVal = parseFloat(detailType[1])
    return detailsCol + floatVal - 1
}
function callValidationToCell(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CIVIL");
    addValidationToCell(sheet, 6, 9, 1, 10);

}

function addValidationToCell(sheet, cellRow, cellCol, minVal, maxVal) {
    var cell = sheet.getRange(cellRow, cellCol);
    var rule = SpreadsheetApp.newDataValidation().requireNumberBetween(minVal, maxVal).build();
    cell.setDataValidation(rule);
  }

function deleteUnselectedRows(subSheetToCopy, selectedSubSheets, classificationSet, costSet, detailsSet, sheetType=null){
    // subSheetToCopy = "Civil"
    // selectedSubSheets = {"A1": [1,2]}
    let mainSubSheet = null
  let mainSection = null
  const srNoCol = 0
  const classificationCol = 7
  const qtCol = 10 
  const costCol = 10
  const detailsCol = 2
  const minQtCol = 19
  const maxQtCol = 20

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(subSheetToCopy);
    const data = getSheetDataByName(subSheetToCopy)
    for (var i=0; i<data.length; i++) {
        row = data[i]
        var srNoValType = getSrNoValType(row[srNoCol])
        if (srNoValType == "sub") {
            mainSubSheet = row[srNoCol]; mainSection = null
            Logger.log('found subsheet - ' + mainSubSheet)
        }
        else if (srNoValType == "number" && (mainSubSheet != null || sheetType)){
            const floatVal = parseFloat(row[srNoCol])
            mainSection = row[srNoCol]
            if (!sheetType && mainSubSheet in selectedSubSheets && classificationSet[mainSubSheet] == row[classificationCol]){
                Logger.log('Row - ' + mainSubSheet + ' - ' + mainSection + ' is selected')
                row[costCol] = row[getCostTypeToColMapping(costSet[mainSubSheet], costCol)] 
                row[detailsCol] = row[getDetailTypeToColMapping(detailsSet[mainSubSheet], detailsCol)]
                sheet.getRange(i+1, 1, 1, row.length).setValues([row]);
                addValidationToCell(sheet, i+1, qtCol, row[minQtCol], row[maxQtCol])
            }
            else if (sheetType){
                row[5] = row[getCostTypeToColMapping(costSet[sheetType], 5)]
                sheet.getRange(i+1, 1, 1, row.length).setValues([row]);
            }
            else if (!sheetType) {
                sheet.deleteRow(i+1);
                return deleteUnselectedRows(subSheetToCopy, selectedSubSheets, classificationSet, costSet, detailsSet)
            }
        }

    }
}

function deleteEmptyAndUnwantedRows(subSheetToCopy, selectedSubSheets, emptyOnly=false){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(subSheetToCopy);
    const srNoCol = 0
    const data = getSheetDataByName(subSheetToCopy)
    for (var i=3; i<data.length; i++) {
        row = data[i]
        var srNoValType = getSrNoValType(row[srNoCol])
        if (!row[srNoCol]){
            sheet.deleteRow(i+1);
            return deleteEmptyAndUnwantedRows(subSheetToCopy, selectedSubSheets)
        }
        if (srNoValType == "sub" && !emptyOnly) {
            mainSubSheet = row[srNoCol];
            if (mainSubSheet in selectedSubSheets){
                Logger.info('Can not delete sub sheet - ' + mainSubSheet + ' as it is selected')
            }
            else{
                sheet.deleteRow(i+1);
                return deleteEmptyAndUnwantedRows(subSheetToCopy, selectedSubSheets)
            }
        }
    }
    // delete 3rd col from the sheet
    // sheet.deleteColumn(3);
}

function deleteColumnsFromSheet(sheetName, colNo, colCount){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    sheet.deleteColumns(colNo, colCount);

}

function iterateOptionsSheet(){
  Logger.log("running iterateOptionsSheet")
  const inputSheet = "Sheet7"
  const data = getSheetDataByName(inputSheet)
  const srNoCol = 1
  const headerCol = 2

  const selectedSheets = {}
  const selectedSubSheets = {}

  let mainSheet = null
  let mainSheetName = null
  let mainSubSheet = null
  let mainSection = null
  const selectedRows = {};

  data.forEach(function(row){
    var srNoValType = getSrNoValType(row[srNoCol])
    if (srNoValType == "main") {
        mainSheet = row[srNoCol]; mainSubSheet = null; mainSection = null;
        mainSheetName = row[headerCol];
        selectedRows[mainSheet] = row[0] == true;
        row[0] = selectedRows[mainSheet]
    }
    else if (srNoValType == "sub") {
        mainSubSheet = row[srNoCol];
        mainSection = null;
        selectedRows[mainSubSheet] = row[0] == true || selectedRows[mainSheet];
        row[0] = selectedRows[mainSubSheet]
    }
    else if (srNoValType == "number") {
        mainSection = row[srNoCol];
        selectedRows[mainSection] = row[0] == true || selectedRows[mainSubSheet];
        row[0] = selectedRows[mainSection]
    }

    if (row[0] == true){
        selectedSheets[mainSheet] = mainSheetName
        if (mainSubSheet ) {
            if (mainSubSheet in selectedSubSheets && mainSection){
                selectedSubSheets[mainSubSheet].push(mainSection)
            }
            else if (mainSubSheet){
                selectedSubSheets[mainSubSheet] = []
                if (mainSection) {
                    selectedSubSheets[mainSubSheet].push(mainSection)
                }
            }
        }
        Logger.log("Got a selected Row- " + mainSheet + " - " + mainSubSheet + " - " + mainSection)
    }
  });

  for (var sheet in selectedSheets) {
    copySubSheetFromMasterToCurrent(selectedSheets[sheet]);
    deleteUnselectedRows(selectedSheets[sheet], selectedSubSheets);
    deleteEmptyAndUnwantedRows(selectedSheets[sheet], selectedSubSheets);
    // iterate over data of the newly copied sheet and remove entries
    // that are not in selectedSubSheets[sheet]
  }
}


function setValToFormula(sheet, row, col, colIdentifiers){
    sheet.getRange(row, col).setValue("=INT(MULTIPLY(" + colIdentifiers[0]+ row.toString() + ", "+ colIdentifiers[1] + row.toString() + "))");
    return
}

function insertAmountCol(sheetName, amountCol=11, colIdentifiers=null){
    const srNoCol=0
    if (!colIdentifiers){
        colIdentifiers = ["I", "J", "L"]
    }
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const data = getSheetDataByName(sheetName);
    for (var i=0; i<data.length; i++) {
        row = data[i]
        var srNoValType = getSrNoValType(row[srNoCol])
        if (srNoValType == "number"){
            mainSection = row[srNoCol]
            setValToFormula(sheet, i+1, amountCol, [colIdentifiers[0], colIdentifiers[1]])
            setValToFormula(sheet, i+1, amountCol+2, [colIdentifiers[0], colIdentifiers[2]])
            // sheet.getRange(i+1, amountCol).setValue("=INT(MULTIPLY("+ colIdentifiers[0]+ (i+1).toString() + ", "+ colIdentifiers[1] + (i+1).toString() + "))");
            // sheet.getRange(i+1, amountCol +2).setValue("=INT(MULTIPLY("+ colIdentifiers[0]+ (i+1).toString() + ", " +colIdentifiers[2] + (i+1).toString() + "))");
        }
    }
}

function insertSumFormula(sheet, start, end, amountCol, colId){
    sheet.getRange(end, amountCol).setValue("=SUM("+colId+ start + ":" + colId + (end-1) + ")");
    return
}

function addSumFormulas(sheetName, amountCol=11, colIdentifiers=null, sheetType=null){
    const srNoCol = 0
    if (!colIdentifiers){
        colIdentifiers = ["K", "M"]
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    const data = getSheetDataByName(sheetName);
    var currentSubSheet = null
    var start = null
    var end = null
    var rowsInserted = 0
    const totalIndexes = []

    if (sheetType){
        
        end = data.length
        sheet.insertRowAfter(end);
        sheet.getRange(end+1, 1).setValue("Sheet Total");
        
        insertSumFormula(sheet, 2, end+1, amountCol, colIdentifiers[0])
        insertSumFormula(sheet, 2, end+1, amountCol+2, colIdentifiers[1])
        return colIdentifiers[0] + (end+1).toString()
    }

    for (var i=0; i<data.length; i++) {
        row = data[i]
        var srNoValType = getSrNoValType(row[srNoCol])
        if (srNoValType == "sub") {
            if (!currentSubSheet){
                currentSubSheet=row[srNoCol]
                // start = i+1
            }
            else{
                // Add new row in top of existing row and add sum formula
                end = i+1 + rowsInserted
                Logger.log('found end, inserting sheet= ' + end + ' - ' + currentSubSheet)
                sheet.insertRowBefore(end);
                sheet.getRange(end, 1).setValue("Total");
                insertSumFormula(sheet, start, end, amountCol, colIdentifiers[0])
                insertSumFormula(sheet, start, end, amountCol+2, colIdentifiers[1])
                // sheet.getRange(end, amountCol).setValue("=SUM("+colIdentifiers[0]+ (start).toString() + ":" + colIdentifiers[0] + (end-1) + ")");
                // sheet.getRange(end, amountCol+2).setValue("=SUM("+colIdentifiers[1]+ (start).toString() + ":" + colIdentifiers[1] + (end-1) + ")");
                totalIndexes.push(end)
                currentSubSheet=row[srNoCol]
                rowsInserted = rowsInserted + 1 // 1 row inserted
            }
            start = i+1 + rowsInserted
        }
    }
    // Add new row in top of existing row and add sum formula
    end = i+1 + rowsInserted

    sheet.insertRowBefore(end);
    sheet.getRange(end, 1).setValue("Total");
    insertSumFormula(sheet, start, end, amountCol, colIdentifiers[0])
    insertSumFormula(sheet, start, end, amountCol+2, colIdentifiers[1])
    
    // sheet.getRange(end, amountCol).setValue("=SUM(" + colIdentifiers[0]+ (start) + ":" + colIdentifiers[0] + (end-1) + ")");
    // sheet.getRange(end, amountCol+2).setValue("=SUM(" + colIdentifiers[1]+ (start) + ":" + colIdentifiers[1] + (end-1) + ")");

    totalIndexes.push(end)
    rowsInserted = rowsInserted + 1
    Logger.log("Found the end of sheet, current index = " + end)
    finalEnd = end + 1
    sheet.insertRowAfter(end);
    sheet.getRange(finalEnd, 1).setValue("Sheet Total");
    const sumString = totalIndexes.map(function(index){return colIdentifiers[0] + index}).join(",")
    const sumStringBCS = totalIndexes.map(function(index){return colIdentifiers[1] + index}).join(",")
    
    sheet.getRange(finalEnd, amountCol).setValue("=SUM("+ sumString + ")");
    sheet.getRange(finalEnd, amountCol+2).setValue("=SUM("+ sumStringBCS + ")");
    
    Logger.log(sumString);
    return colIdentifiers[0] + (finalEnd).toString()

}

function updateSummarySheet(totalSumRecords){
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary");
    data = getSheetDataByName("Summary");
    for (var i=0; i<data.length; i++) {
        row = data[i]
        const floatVal = parseFloat(row[0])
        if (floatVal && totalSumRecords.hasOwnProperty(row[1])) {
            const formulaStr = "='"+ row[1] + "'!" + totalSumRecords[row[1]]
            Logger.log(formulaStr);
            sheet.getRange(i+1, 3).setFormula(formulaStr);
        }
    }

}

function generateBOQ(){
    const STATIC_SHEETS = [
        "Electrical Basic", "Electrical Detailed", "HVAC - VRF", "HVAC - Ductable", "HVAC - Ductable - Low Side Only",
        "Networking", "Fire Alarm System", "Fire Sprinkler System - Ext", "Fire Sprinkler System", "Access Control", "CCTV",
        "PA System"
    ]
    Logger.log("running generateBOQ")
    const inputSheet = "input"
    const data = getSheetDataByName(inputSheet)
    const srNoCol = 1
    const headerCol = 2
    const classificationCol = 4
    const costCol = 3
    const detailsCol = 5

    const classificationSet = {}
    const costSet = {}
    const detailsSet = {}
    const selectedSheets = {}
    const selectedSubSheets = {}
    const selectedRows = {};
    let mainSheet = null
    let mainSheetName = null
    let mainSubSheet = null
    
    data.forEach(function(row){
      var srNoValType = getSrNoValType(row[srNoCol])
      if (srNoValType == "main") {
          mainSheet = row[srNoCol]; mainSubSheet = null;
          mainSheetName = row[headerCol];
          selectedRows[mainSheet] = row[0] == true;
          row[0] = selectedRows[mainSheet]
          if(row[0]){
            selectedSubSheets[mainSheet + "1"] = []
            classificationSet[mainSheet + "1"] = row[classificationCol] || "S1"
            costSet[mainSheet + "1"] = row[costCol] || "C1"
            detailsSet[mainSheet + "1"] = row[detailsCol] || "BASIC"
          }
          classificationSet[mainSheet] = row[classificationCol] || "S1"
          costSet[mainSheet] = row[costCol] || "C1"
          detailsSet[mainSheet] = row[detailsCol] || "BASIC"
      }
      else if (srNoValType == "sub") {
          mainSubSheet = row[srNoCol];
          selectedRows[mainSubSheet] = row[0] == true || selectedRows[mainSheet];
          row[0] = selectedRows[mainSubSheet]
          classificationSet[mainSubSheet] = row[classificationCol] || classificationSet[mainSheet]
            costSet[mainSubSheet] = row[costCol] || costSet[mainSheet]
            detailsSet[mainSubSheet] = row[detailsCol] || detailsSet[mainSheet]
      }
      
  
      if (row[0] == true){
          selectedSheets[mainSheet] = mainSheetName
            if (mainSubSheet ) {
                selectedSubSheets[mainSubSheet] = []
            }
        //   Logger.log("Got a selected Row- " + mainSheet + " - " + mainSubSheet + " - " + mainSheetName)
      }
    });

    const totalSumRecords = {}
    for (var sheet in selectedSheets) {
        Logger.log('Executing for sheet- ' + selectedSheets[sheet])
        copySubSheetFromMasterToCurrent(selectedSheets[sheet]);
        if (STATIC_SHEETS.includes(selectedSheets[sheet])){
            Logger.log("Skipping " + selectedSheets[sheet] + " where sheet was " + sheet)
            deleteUnselectedRows(selectedSheets[sheet], selectedSubSheets, classificationSet, costSet, detailsSet, sheetType=sheet);
        
            deleteColumnsFromSheet(selectedSheets[sheet], 7, 4);
            sheetObj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(selectedSheets[sheet]);
            sheetObj.getRange(1, 6).setValue("UNIT RATE");
            sheetObj.getRange(2, 6).setValue("");

            insertAmountCol(selectedSheets[sheet], amountCol=7, colIdentifiers=["E", "F", "H"])
            totalSumRecords[selectedSheets[sheet]] = addSumFormulas(selectedSheets[sheet], amountCol=7, colIdentifiers=["G", "I"], sheetType=sheet);

            continue
        }
        Logger.log("Deleting unselected rows for " + selectedSheets[sheet])
        deleteUnselectedRows(selectedSheets[sheet], selectedSubSheets, classificationSet, costSet, detailsSet);
        deleteEmptyAndUnwantedRows(selectedSheets[sheet], selectedSubSheets);
        deleteColumnsFromSheet(selectedSheets[sheet], 3, 1);
        deleteColumnsFromSheet(selectedSheets[sheet], 11, 4);
        deleteColumnsFromSheet(selectedSheets[sheet], 15, 2);
        // deleteColumnsFromSheet(selectedSheets[sheet], 10);
        // deleteColumnsFromSheet(selectedSheets[sheet], 10);
        // deleteColumnsFromSheet(selectedSheets[sheet], 10);
        sheetObj = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(selectedSheets[sheet]);
        sheetObj.getRange(1, 3).setValue("SPECIFICATION");
        sheetObj.getRange(1, 10).setValue("UNIT RATE");
        sheetObj.getRange(2, 10).setValue("");
        // sheet.getRange(0, col).setValue("Hello World");
        // Add formulas in the sheet 
        insertAmountCol(selectedSheets[sheet])
        // const finalRow = addSumFormulas(selectedSheets[sheet]);
        totalSumRecords[selectedSheets[sheet]] = addSumFormulas(selectedSheets[sheet]);
      }
    Logger.log(totalSumRecords)

    updateSummarySheet(totalSumRecords)
    return;
  }

function testData(){
    // sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CIVIL");
    updateSummarySheet({"CEILING": "K23"})
}


// flooring ÷
// Partitions and Doors
// White Goods
