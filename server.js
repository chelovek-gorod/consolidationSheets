'use strict';

const ui = SpreadsheetApp.getUi();
const sidebar = HtmlService.createHtmlOutputFromFile("sidebar").setTitle("Consolidate sheets");

const cache = CacheService.getDocumentCache();

//onOpen();
ui.createMenu("Consolidation").addItem("New consolidation", "showSidebar").addToUi();

// UI
function onOpen() {
  cache.remove("cachedArr");
  ui.createMenu("Consolidation").addItem("New consolidation", "showSidebar").addToUi();
}

function showSidebar() {
  cache.remove("cachedArr");
  ui.showSidebar(sidebar);
}

function showModal(apdate) {
  if (apdate) cache.put("cachedArr", JSON.stringify(apdate));
  const htmlOutput = HtmlService.createHtmlOutputFromFile("modal")
  .setWidth(400).setHeight(500);
  ui.showModalDialog(htmlOutput, 'Add spreadsheets');
}

function showMessage(message) {
  ui.alert(message);
}

/*
[{
  ssI: spreadshit id,
  ssN: spreadshit name,
  siArr: [sheets id's for consolidation arr],
  snArr: [sheets name's for consolidation arr],
  inArr: [icluding for consolidation sheets arr],
  cnArr: [ceils for consolidation number arr]
}]
*/
const consolidateArr = updateConsolidateArr();

function updateConsolidateArr() {
  let cachedArr = cache.get("cachedArr");
  if (cachedArr !== null) return JSON.parse(cachedArr);
  else {
    let assApp = SpreadsheetApp.getActiveSpreadsheet();
    let obj = getConsolidateArrObject(assApp.getId(), assApp.getName(), assApp.getSheets()); 
    cache.put("cachedArr", JSON.stringify([obj]));
    return [obj];
  }
}

function getConsolidateArrObject (ssId, ssName, ssSheets) {
  let sArr = []; // put in siArr
  let nArr = []; // put in snArr
  let iArr = []; // put inArr
  let cArr = []; // put in cnArr
  for (let s of ssSheets) {
    let sId = s.getSheetId();
    let sName = s.getName();
    let cn = getSheetCeilsNumber(ssId, sId);
    sArr.push(sId);
    nArr.push(sName);
    iArr.push(1); // 1 = included true; 0 = included false
    cArr.push(cn);
  }
  return {ssI: ssId, ssN: ssName, siArr: sArr, snArr: nArr, inArr: iArr, cnArr: cArr};
}

function getSheet(spreadSheetId, sheetId) {
  return SpreadsheetApp.openById(spreadSheetId).getSheets().filter((s) => s.getSheetId() == sheetId)[0]; // not === !!! becouse id 0 != '0'
}
function getSheetCeilsNumber(spreadSheetId, sheetId) {
  let s = getSheet(spreadSheetId, sheetId);
  let rc = s.getDataRange().getValues();
  let cols = rc.length - 1;
  let rows = rc[0].length - 1;
  return cols * rows;
}

function clearConsolidateArr() {
  cache.put("cachedArr", JSON.stringify([]));
}

function startConsolidation() {
  cache.remove("cachedArr");
  let cachedResult = JSON.stringify({tittle: '', lines: [], props: []})
  cache.put("cachedResult", cachedResult);
}

/////////////////////////////////////////////////////////////////

// CALLBACK FUNCTIONS

function getConsolidateArr() {
  return consolidateArr;
}

function getSpraedsheetsList() {
  let filesArr = [];
  let files = DriveApp.getFilesByType(MimeType.GOOGLE_SHEETS);
  while(files.hasNext()) {
    let f = files.next();
    let included = false;
    if (~consolidateArr.findIndex(obj => obj.ssI === f.getId())) included = true;
    filesArr.push({name: f.getName(), id: f.getId(), inc: included});
  }
  return filesArr;
}

function apdateSpreadsheets(addSSarr) {
  for (let id of addSSarr) {
    let ss = SpreadsheetApp.openById(id).getSheets();
    let sn = SpreadsheetApp.openById(id).getName();
    let obj = getConsolidateArrObject (id, sn, ss);
    consolidateArr.push(obj);
  }
  cache.put("cachedArr", JSON.stringify(consolidateArr));
  ui.showSidebar(sidebar);
}

function consolidation(object) {
  // object = {ssId: id, sId: id} 

  // result = {tittle: s, lines: [{name, []},{name, []}], props: [s,s,s]} 
  let result = JSON.parse(cache.get("cachedResult"));
  let sheet = getSheet(object.ssId, object.sId);
  // update row with propertys
  let prop = 1; // column iteration number in parsing sheet
  let propsArr = []; // array of sheet propertys [A1,B1,C1...]
  while (sheet.getRange(1, prop).getValue() !== '') {
    let property = sheet.getRange(1, prop).getValue();
    // update title
    if (prop === 1) {
      if (result.tittle === '') result.tittle = property;
    } else {
      propsArr.push(property);
      // update result props arr if necessary
      if (result.props.indexOf(property) < 0) result.props.push(property);
    }
    prop++;
  }
  // update lines
  let line = 2; // row iteration number in parsing sheet
  while (sheet.getRange(line, 1).getValue() !== '') {
    let resultLineIndex; // index in result array lines [i]
    for (let ceil = 1; ceil < prop; ceil++) {
      let value = sheet.getRange(line, ceil).getValue();
      if (ceil === 1) {
        resultLineIndex = result.lines.findIndex(i => i.name === value);
        // update result lines arr if necessary and resultLineIndex
        if (resultLineIndex < 0) {
          resultLineIndex = result.lines.length;
          result.lines.push({name: value, lineArr: []});
        }
      } else {
        // consolidate ceils
        if (value) {
          let resultPropIndex = result.props.indexOf(propsArr[ceil-2]);
          let resultCeil = result.lines[resultLineIndex].lineArr[resultPropIndex];
          if (resultCeil) resultCeil = resultCeil + value;
          else resultCeil = value;
          result.lines[resultLineIndex].lineArr[resultPropIndex] = resultCeil;
        }
      }
    }
    line++;
  }
  cache.put("cachedResult", JSON.stringify(result));
}

function endConsolidation() {
  // result = {tittle: s, lines: [{name, []},{name, []}], props: [s,s,s]} 
  let result = JSON.parse(cache.get("cachedResult"));

  let assApp = SpreadsheetApp.getActiveSpreadsheet();
  let resultSheet;
  // create or clear resultSheet 
  if (assApp.getSheetByName("Results") !== null) {
    resultSheet = assApp.getSheetByName("Results");
    resultSheet.clear();
  } else {
    resultSheet = assApp.insertSheet(assApp.getSheets().length+1);
    resultSheet.setName("Results");
  }

  resultSheet.getRange(1, 1).setValue(result.title).setFontWeight("bold");

  for (let prop = 0; prop < result.props.length; prop++) {
    resultSheet.getRange(1, prop + 2).setValue(result.props[prop]).setFontWeight("bold");
  }
  for (let line = 0; line < result.lines.length; line++) {
    resultSheet.getRange(line + 2, 1).setValue(result.lines[line].name).setFontWeight("bold");
    for (let ceil = 0; ceil < result.lines[line].lineArr.length; ceil++) {
      let value = result.lines[line].lineArr[ceil];
      if (value) resultSheet.getRange(line + 2, ceil + 2).setValue(value);
    }
  }
  return assApp.getUrl() + '#gid=' + assApp.getSheetId();
}
