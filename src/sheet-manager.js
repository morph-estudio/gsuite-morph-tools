/**
 * Gsuite Morph Tools - Sheet Manager 1.8.0
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
  let activeSheet = SpreadsheetApp.getActiveSheet();
  let rowReorderedname = rowReorderedsend.map(function (row) { return row.name });
  let hiddenList = [];
  let sheet;

  for(let i = 0; i < rowReorderedname.length; i++) {
    sheet = ss.getSheetByName(rowReorderedname[i]);
    if (sheet.isSheetHidden()) { 
        hiddenList.push(rowReorderedname[i]) 
    }
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(i + 1);
  }

  SpreadsheetApp.flush();

  for (let i = 0; i < hiddenList.length; i++) {
    Utilities.sleep(50);
    ss.getSheetByName(hiddenList[i]).hideSheet();
  }

  activeSheet.activate();
}

/**
 * moveSingleSheet
 * Faster code to move one single sheet.
 */
function moveSingleSheet(sh, shAfter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let shActive = ss.getActiveSheet();
  sh = ss.getSheetByName(sh);
  let isHidden = sh.isSheetHidden(); Logger.log(isHidden)
  sh.activate();
  let shIndex = sh.getIndex()
  shAfter = ss.getSheetByName(shAfter);
  let shAfterIndex = shAfter.getIndex();
  shAfterIndex > shIndex ? ss.moveActiveSheet(shAfterIndex) : ss.moveActiveSheet(shAfterIndex + 1); 
  shActive.activate();
  SpreadsheetApp.flush(); waiting(500);
  if (isHidden == true) {
    sh.hideSheet();
  }
  if (isHidden == true) {
    sh.hideSheet();
  }
}

/**
 * goToSheet
 * Fast-travel between sheets to use in the sheets table.
 */
function goToSheet(sheetname) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetname);
  sh.getRange("A1");
  sh.activate();
}
