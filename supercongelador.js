function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}


// 📜 SUPERCONGELADOR 🟢🔴

function congeladorMorph() {

  var ss = SpreadsheetApp.getActive();
  var ws = SpreadsheetApp.getActiveSheet();
  var spreadsheetId = SpreadsheetApp.getActive().getId();
  var dateGMT = Utilities.formatDate(new Date(), "GMT+1", "yyyyMMdd")
  var file = DriveApp.getFileById(spreadsheetId);
  var parentFolder = file.getParents();
  var parentFolder_ID = parentFolder.next().getId();
  var parentFolderName = file.getParents().next().getName();


  // Copy each sheet in the source Spreadsheet by removing the formulas as the temporal sheets. 🟢

    var ss = SpreadsheetApp.openById(spreadsheetId);
    var tempSheets = ss.getSheets().filter(sh => !sh.isSheetHidden()).map(function(sheet) {
      var dstSheet = sheet.copyTo(ss).setName(sheet.getSheetName() + "_temp");
      var src = dstSheet.getDataRange();
      src.copyTo(src, {contentsOnly: true});
      return dstSheet;
    });


  // Copy the source Spreadsheet. 🟢

    var destination = ss.copy(ss.getName() + " - " + dateGMT + " - " + "CONGELADO");
    var destinationId = destination.getId();
    var destinationFile = DriveApp.getFileById(destinationId);


  // Delete the temporal sheets in the source Spreadsheet. 🟢

    tempSheets.forEach(function(sheet) {ss.deleteSheet(sheet)});
    SpreadsheetApp.flush();


  // Delete the original sheets from the copied Spreadsheet and rename the copied sheets. 🟢

    destination.getSheets().forEach(function(sheet) {
    if (sheet.isSheetHidden()) {
      destination.deleteSheet(sheet);
      SpreadsheetApp.flush();
    }
    else {
      var sheetName = sheet.getSheetName();
      //Logger.log(sheetName);
      if (sheetName.indexOf("_temp") == -1) {
        destination.deleteSheet(sheet);
        SpreadsheetApp.flush();
      }
      else {
        //sheet.setName(sheetName.slice(0, -1));
        sheet.setName(sheetName.replace("_temp", ""));
      }
    }
    });

  // Move file to the destination folder. 🟢

    var file = DriveApp.getFileById(destinationId);
    DriveApp.getFolderById(parentFolder_ID).addFile(file);
    file.getParents().next().removeFile(file);


  // Export to XLSX. 🟢

    try {

    var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + destination.getId() + "&exportFormat=xlsx";

    var params = {
      method      : "get",
      headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    };

    var blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(ss.getName() + " - " + dateGMT + " - " + "CONGELADO" + ".xlsx");

    var ui = SpreadsheetApp.getUi();
    var confirm = Browser.msgBox('Documento Excel','¿Quieres crear una copia en formato XLSX en la misma carpeta?', Browser.Buttons.OK_CANCEL);
    if(confirm=='ok'){ DriveApp.getFolderById(parentFolder_ID).createFile(blob); };

    //MailApp.sendEmail('asanchez@morphestudio.es', 'Conversión de Google Sheet a Excel', 'El archivo XLSX aparece adjunto a este correo.', { attachments: [blob] });

    } catch (f) {
      Logger.log(f.toString());
    }



}
