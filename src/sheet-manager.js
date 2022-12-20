/**
 * Gsuite Morph Tools - Sheet Manager 1.7.0
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

/**
 * getWorksheetNames
 * Returns a list of sheetnames, used for refreshing the table. 
 */
function getWorksheetNames() {
  let sheetNames = new Array();
  let sheets = ss().getSheets();
  for (let i = 0 ; i < sheets.length ; i++) sheetNames.push( { 'id': i, 'name': sheets[i].getName(), 'button': `<button type="button" id="goToSheetButton${i}" class="btn btn-primary background-none btn-sm viewSheetButton"><i class="fa-solid fa-right-long"></i>`, 'ishidden': sheets[i].isSheetHidden() } )
  
  return sheetNames;
}

/**
 * getWorksheetNamesArray
 * Returns a list of sheetnames, used for the first load after opening the sidebar.
 */
function getWorksheetNamesArray() {
  let sheetNames = new Array();
  let sheets = ss().getSheets();
  sheets.forEach(sh => {
    sheetNames.push( sh.getName());
  });
  return sheetNames;
}

/**
 * worksheetManagement
 * Main function of Worksheet Manager; responds to execute button.
 */
function worksheetManagement(sheetNamesToDeleteAsString, rowData) {
  const ss = SpreadsheetApp.getActive();

  let formData = [rowData.sdAction]; let [sdAction] = formData;
  let sheetsToManage = []; let shtm;

  let sheetNamesParsed = JSON.parse(sheetNamesToDeleteAsString);
  sheetNamesParsed = sheetNamesParsed.forEach((obj) => { sheetsToManage.push(obj.name) });
  
  Logger.log(`Selected sheets: ${sheetsToManage}`);

  sheetsToManage.forEach((sheet) => {
    shtm = ss.getSheetByName(sheet);
    let shIndex = shtm.getIndex();

    // This switch responds to the selected action for the selected sheets in frontend.
    switch (sdAction) {
      case 'actDelete':
        ss.deleteSheet(shtm);
        break;
      case 'actHide':
        shtm.hideSheet();
        break;
      case 'actShow':
        shtm.showSheet();
        break;
      case 'actClear':
        shtm.clear().clearNotes();
        break;
      case 'actDuplicate':
        shtm.copyTo(ss).activate();
        ss.moveActiveSheet(shIndex + 1);
        break;
      default:
    }
  });

  // Special switch for the partial freezer.
  switch (sdAction) {
    case 'actPartialFreezer':
      morphFreezer(sdAction, sheetsToManage);
      break;
    default:
  }

}

/**
 * rearrangeSheets
 * Rearrange sheets based on the main sheet table. It's based in Drag & Drop mode.
 */
function rearrangeSheets(rowReorderedsend) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let rowReorderedid = rowReorderedsend.map(function (row) { return row.id });
  let rowReorderedname = rowReorderedsend.map(function (row) { return row.name });

  let sheet; let hiddenList = [];

  rowReorderedname.forEach((sh, index) => {
    // Logger.log(`Sheet Index: ${index}`);
    sheet = ss.getSheetByName(sh);
    sheet.activate();
    ss.moveActiveSheet(index + 1);
    if(sheet.isSheetHidden()) { hiddenList.push(sh) }
  });

  SpreadsheetApp.flush();

  Logger.log(`List of Hidden Sheets: ${hiddenList}`);

  // This code tries to avoid the bug when you can't hide the active sheet.
  let myArray = rowReorderedname.filter( function( el ) { return !hiddenList.includes( el ) });
  ss.getSheetByName(myArray[0]).activate();

  // The code is duplicated because of a GAS bug that doesn't hide the sheets if it's in the last position (sometimes). Maybe we could find another fix here.
  hiddenList.forEach(sh => {
    sheet = ss.getSheetByName(sh);
    SpreadsheetApp.flush();
    sheet.hideSheet();
  });
  hiddenList.forEach(sh => {
    sheet = ss.getSheetByName(sh);
    SpreadsheetApp.flush();
    sheet.hideSheet();
  });

}

/**
 * goToSheet
 * Fast-travel between sheets to use in the sheets table.
 */
function goToSheet(sheetname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetname);
  sh.activate();
}
