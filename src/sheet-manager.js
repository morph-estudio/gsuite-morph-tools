/**
 * Gsuite Morph Tools - Sheet Manager 1.5
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

function getWorksheetNames() {
  let sheetNames = new Array();
  let sheets = ss().getSheets();
  for (let i = 0 ; i < sheets.length ; i++) sheetNames.push( { name: sheets[i].getName() } )
  return sheetNames;
}

function getWorksheetNamesArray() {
  let sheetNames = new Array();
  let sheets = ss().getSheets();
  sheets.forEach(sh => {
    sheetNames.push( sh.getName());
  });
  return sheetNames;
}

function deleteWorksheets(sheetNamesToDeleteAsString, rowData) {
  const ss = SpreadsheetApp.getActive();

  let formData = [rowData.sdAction];
  let [sdAction] = formData;
  let sheetsToDelete = [];

  let sheetNamesParsed = JSON.parse(sheetNamesToDeleteAsString);
  sheetNamesParsed = sheetNamesParsed.forEach((obj) => {
    sheetsToDelete.push(obj.name);
  });

  sheetsToDelete.forEach((sheet) => {
    let shtd = ss.getSheetByName(sheet);
    let shIndex = shtd.getIndex();
    switch (sdAction) {
      case 'act-delete':
        ss.deleteSheet(shtd);
        break;
      case 'act-hide':
        shtd.hideSheet();
        break;
      case 'act-clear':
        shtd.clear().clearNotes();
        break;
      case 'act-duplicate':
        shtd.copyTo(ss).activate();
        ss.moveActiveSheet(shIndex + 1);
        break;
      default:
    }
  });
}
