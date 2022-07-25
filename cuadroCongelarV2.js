function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}


// 📜 CONGELAR CUADRO DE SUPERFICIES ✔️

function congelarCuadroNew() {


  // BASE 🟢

    var ss = SpreadsheetApp.getActive();
    var sh = SpreadsheetApp.getActiveSheet();
    var ws = SpreadsheetApp.getActive().getSheetByName('ACTUALIZAR');
    var ss_id = SpreadsheetApp.getActive().getId();
    var dateGMT = Utilities.formatDate(new Date(), "GMT+1", "yyyyMMdd");
    var file = DriveApp.getFileById(ss_id);
    var parentFolder = file.getParents();
    var parentFolder_ID = parentFolder.next().getId();
    var backupFolderSearch = DriveApp.getFolderById(parentFolder_ID);
    var originalDocName = ss.getName();


  // DESTINATION FOLDER (Específico cuadro de superficies) 🟢

    var searchFor ='title contains "Congelados"';
    var names =[];
    var backupFolderIds=[];
    var backupFolder = backupFolderSearch.searchFolders(searchFor);
      while (backupFolder.hasNext()) {
    var backupFolderDef = backupFolder.next();
    var backupFolderId = backupFolderDef.getId();
    backupFolderIds.push(backupFolderId);
    var backupFolderName = backupFolderDef.getName();
    var backupFolderUrl = backupFolderDef.getUrl();
    names.push(backupFolderName);
    }

    ws.getRange('B7')
    .setFontWeight("bold")
    .setFontColor("#B7B7B7")
    ;
    ws.getRange('A7').setValue('CARPETA BACKUP');
    ws.getRange('A8').setValue('ID CARPETA BACKUP');
    ws.getRange('A9').setValue('DESCARGAR ARCHIVO XLSX');
    ws.getRange('B7').setValue('=hyperlink("https://drive.google.com/corp/drive/folders/'+ backupFolderId +'";"' + backupFolderName + '")').setFontColor('#4A86E8');
    //ws.getRange('D7').setValue(backupFolderName);
    ws.getRange('B8').setValue(backupFolderId);


  // Copy the source Spreadsheet. 🟢

    var destination = ss.copy(ss.getName());

    file.setName(ss.getName() + " - " + dateGMT + " - " + "CONGELADO");

    var destinationId = destination.getId();
    var destinationFile = DriveApp.getFileById(destinationId);
    var ws_2 = destination.getSheetByName('ACTUALIZAR');

    destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    SpreadsheetApp.flush();

    var tempSheets = ss.getSheets().filter(sht => !sht.isSheetHidden()).map(function(sheet) {
      var dstSheet = sheet.getDataRange();
      dstSheet.copyTo(dstSheet, {contentsOnly: true});
      return dstSheet;
    });

    var tempSheets = ss.getSheets().filter(sht => sht.isSheetHidden()).map(function(sheet) {
      ss.deleteSheet(sheet);
    });
/*
    ss.getSheets().forEach(function(sheet) {
    if (sheet.isSheetHidden()) {
      ss.deleteSheet(sheet);
    } else{
      var dstSheet = sheet.getDataRange();
      dstSheet.copyTo(dstSheet, {contentsOnly: true});
      return dstSheet;
    }
    });
*/
    var fileOriginal = DriveApp.getFileById(destinationId);

    SpreadsheetApp.flush();

 // Delete the main DATA SHEET. 🟢

    ss.getSheets().forEach(function(sheet) {
    var sheetName = sheet.getSheetName();
    if (sheetName.indexOf("ACTUALIZAR") == -1) {
    }
       else {
        ss.deleteSheet(sheet);
      }
    });



  // Export to XLSX. 🟢

    try {

    var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + ss_id + "&exportFormat=xlsx";

    var params = {
      method      : "get",
      headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    };

    var blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(ss.getName() + ".xlsx");

    var blobFile = UrlFetchApp.getRequest(url, params);

    ws_2.getRange('B9').setValue(url).setFontColor('#4A86E8');

  } catch (f) {
    Logger.log(f.toString());
  }

  SpreadsheetApp.flush();


  // Move files 🟢

    file.moveTo(backupFolderDef);
    fileOriginal.moveTo(backupFolderSearch);
    var fileoriginalUrl = fileOriginal.getUrl();

    var html = "<script>window.open('" + fileoriginalUrl + "');google.script.host.close();</script>";
    var userInterface = HtmlService.createHtmlOutput(html);
    SpreadsheetApp.getUi().showModalDialog(userInterface, 'Abriendo el documento original...');



}
