function doGet() {
  return HtmlService.createHtmlOutputFromFile("sheetDeleterIndex");
}

function getWorksheetNames(){

  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = ss.getSheets()

  var sheetNames = sheets.map(sheet =>{
    return [sheet.getSheetName()]
  })

  return sheetNames
}

function deleteWorksheets(sheetNamesToDeleteAsString){
  var sheetNamesToDelete = JSON.parse(sheetNamesToDeleteAsString)
  var ss = SpreadsheetApp.getActiveSpreadsheet()
  var sheets = ss.getSheets()

  var sheetsToDelete = sheets.filter(sheet => sheetNamesToDelete.includes(sheet.getSheetName()))

  sheetsToDelete.forEach(sheet =>{
  ss.deleteSheet(sheet)
  })
  
}