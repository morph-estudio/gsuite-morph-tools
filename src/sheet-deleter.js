function doGet() {
  return HtmlService.createHtmlOutputFromFile('test/sheetDeleterIndex');
}

function getWorksheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  let sheetNames = sheets.map((sheet) => [sheet.getSheetName()]);
  return sheetNames;
}

function deleteWorksheets(sheetNamesToDeleteAsString, rowData) {
  let sheetNamesToDelete = JSON.parse(sheetNamesToDeleteAsString);
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();

    let formData = [
    rowData.gmtAction,
  ];

  let [gmtAction] = formData;



  let sheetsToDelete = sheets.filter((sheet) => sheetNamesToDelete.includes(sheet.getSheetName()));

sheetsToDelete.forEach((sheet) => {
switch (gmtAction) {
  case 'act-delete':
    ss.deleteSheet(sheet);
    break;
  case 'act-hide':
    ss.getSheetByName(sheet.getSheetName()).hideSheet();
    break;
  case 'act-clear':
    ss.getSheetByName(sheet.getSheetName()).clear().clearNotes();
    break;
  default:
}
});

/*
  sheetsToDelete.forEach((sheet) => {
    ss.deleteSheet(sheet);
  });
*/






}



