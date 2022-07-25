function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}


function tplActualizarCuadroManual() {

  var ss = SpreadsheetApp.getActive();
  var sh = SpreadsheetApp.getActiveSheet();
  var ws = ss.getSheetByName('ACTUALIZAR') || ss.insertSheet('ACTUALIZAR', 1);
  var ss_id = ss.getId();
  var file = DriveApp.getFileById(ss_id);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  ws.clear().clearFormats();

  ws.getRange('A1').setValue('URL PANEL DE CONTROL');
  ws.getRange('A3').setValue('CARPETA CUADRO SUP.');
  ws.getRange('A4').setValue('ID CARPETA CUADRO SUP.');
  ws.getRange('A5').setValue('CARPETA PANEL DE CONTROL');
  ws.getRange('A6').setValue('ID CARPETA PANEL DE CONTROL');
  ws.getRange('A7').setValue('CARPETA BACKUP');
  ws.getRange('A8').setValue('ID CARPETA BACKUP');
  ws.getRange('A9').setValue('DESCARGAR ARCHIVO XLSX');

  ws.getRange('C2').setValue('Archivos exportados');
  ws.getRange('D2').setValue('IDs');
  ws.getRange('C3').setValue('TXT Sheets Falsos techos.txt');
  ws.getRange('C4').setValue('TXT Sheets Superficies.txt');
  ws.getRange('C5').setValue('TXT Sheets Ventanas.txt');

  // Estilo global
  ws.getRange(1, 1, ws.getMaxRows(), ws.getMaxColumns()).setFontSize(13).setFontFamily("Montserrat").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
  .setVerticalAlignment("middle");
  // Col A
  ws.getRange(1,1,9,1).setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor("#B7B7B7").setFontWeight("bold");
  // Row 2
  ws.getRange(2,1,1,4).setFontFamily("Inconsolata").setFontSize(16).setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setFontColor("#999999").setVerticalAlignment("middle").setFontWeight("bold").setHorizontalAlignment("center");
  // Col B
  ws.getRange(3,2,7,1).setBackground("#F3F3F3").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor("#B7B7B7");
  // ImportRanges
  ws.getRange(1,3,1,2).setBackground("#EAD1DC").setBorder(true, true, true, true, true, true, "#A64D79", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setFontColor("#A64D79").setFontWeight   ("bold").setHorizontalAlignment("center");
  // Control Panel
  ws.getRange('B1').setBackground("#FFF2CC").setBorder(true, true, true, true, true, true, "#BF9000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setFontColor("#BF9000").setFontWeight("bold");
  // Folders Bold
  ws.getRange('B3').setFontWeight("bold").setFontColor("#B7B7B7");
  ws.getRange('B5').setFontWeight("bold").setFontColor("#B7B7B7");
  ws.getRange('B7').setFontWeight("bold").setFontColor("#B7B7B7");
  ws.getRange('B8').setBackground("#FFF2CC").setBorder(true, true, true, true, true, true, "#BF9000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setFontColor("#BF9000");
  ws.getRange('B9').setFontWeight("bold").setFontColor("#B7B7B7");
  // Column/Row Size

  ws.setColumnWidth(1, 330);
  ws.setColumnWidth(2, 380);
  ws.setColumnWidth(3, 285);
  ws.setColumnWidth(4, 400);

  var maxRows = ws.getMaxRows();
  for (var i=1; i < maxRows+1; i++) {
  ws.setRowHeight(i,27)
  };
  ws.setRowHeight(2, 50);

  // FORMAT FILELIST
  ws.getRange(3,3,3,2).setFontSize(13).setFontFamily("Montserrat").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setFontColor('#B7B7B7');
  ws.getRange(3,3,3,1).setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor("#B7B7B7").setFontWeight("bold");
  ws.getRange(3,4,3,1).setBackground("#FFF2CC").setBorder(true, true, true, true, true, true, "#BF9000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setFontColor('#BF9000');


  // IMPORTRANGE PERMISSION 🟢

  var sectoresID = "1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";
  var sectores = DriveApp.getFileById(sectoresID);

  ws.getRange('C1').setValue('=IMPORTRANGE(B1;"Instrucciones!A1")');
  ws.getRange('D1').setValue('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"DB-SI!B2")');


  deleteEmptyRows();
  removeEmptyColumns();
  ws.activate();

}




function actualizarCuadroManual() {

  var ss = SpreadsheetApp.getActive();
  var sh = SpreadsheetApp.getActiveSheet();
  var ws = ss.getSheetByName('ACTUALIZAR') || ss.insertSheet('ACTUALIZAR', 1);
  var ss_id = ss.getId();
  var file = DriveApp.getFileById(ss_id);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  var filePanelUrl = ws.getRange('B1').getValue();
  var filePanelId = filePanelUrl.replace(/.*\/d\//, '').replace(/\/.*/,  '');
  var panelControl = DriveApp.getFileById(filePanelId);
  var folderPanelc = panelControl.getParents();
  var folderPanelcDef = folderPanelc.next();
  var folderPanelcId = folderPanelcDef.getId();
  ws.getRange('B5').setValue('=hyperlink("https://drive.google.com/corp/drive/folders/'+ folderPanelcId +'";"' + folderPanelcDef + '")').setFontColor('#4A86E8');
  ws.getRange('B6').setValue(folderPanelcId);


  var parents = file.getParents();
  var carpetaBaseID = parents.next().getId();
  var carpetaBase = DriveApp.getFolderById(carpetaBaseID);
  ws.getRange('B3').setValue('=hyperlink("https://drive.google.com/corp/drive/folders/'+ carpetaBaseID +'";"' + carpetaBase + '")').setFontColor('#4A86E8');
  ws.getRange('B4').setValue(carpetaBaseID);


  var txtFileId_FT = ws.getRange(3,4).getValue();
  var txtFileId_SP = ws.getRange(4,4).getValue();
  var txtFileId_VN = ws.getRange(5,4).getValue();

  var txtFile_FT = DriveApp.getFileById(txtFileId_FT);
  txtFile_FT.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var txtFile_SP = DriveApp.getFileById(txtFileId_SP);
  txtFile_SP.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var txtFile_VN = DriveApp.getFileById(txtFileId_VN);
  txtFile_VN.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);


  // FALSOS TECHOS

  var tsvUrl_FT = "https://drive.google.com/uc?id="+txtFileId_FT+"&x=.tsv";
  var tsvContent_FT = UrlFetchApp.fetch(tsvUrl_FT,{muteHttpExceptions: true }).getContentText();
  var tsvData_FT = Utilities.parseCsv(tsvContent_FT, '\t');

  var sheet_FT = ss.getSheetByName('TXT FALSOS TECHOS') || ss.insertSheet('TXT FALSOS TECHOS', 200);
  sheet_FT.setTabColor("F1C232");
  sheet_FT.clear();
  sheet_FT.getRange(1, 1, tsvData_FT.length, tsvData_FT[0].length).setValues(tsvData_FT);

  // TXT SUPERFICIES

  var tsvUrl = "https://drive.google.com/uc?id="+txtFileId_SP+"&x=.tsv";
  var tsvContent = UrlFetchApp.fetch(tsvUrl,{muteHttpExceptions: true }).getContentText();
  var tsvData = Utilities.parseCsv(tsvContent, '\t');

  var sheet_SP = ss.getSheetByName('TXT SUPERFICIES') || ss.insertSheet('TXT SUPERFICIES', 200);
  sheet_SP.setTabColor("F1C232");
  sheet_SP.clear();
  sheet_SP.getRange(1, 1, tsvData.length, tsvData[0].length).setValues(tsvData);


  // VENTANAS

  var tsvUrl_VN = "https://drive.google.com/uc?id="+txtFileId_VN+"&x=.tsv";
  var tsvContent_VN = UrlFetchApp.fetch(tsvUrl_VN,{muteHttpExceptions: true }).getContentText();
  var tsvData_VN = Utilities.parseCsv(tsvContent_VN, '\t');

  var sheet_VN = ss.getSheetByName('TXT VENTANAS') || ss.insertSheet('TXT VENTANAS', 200);
  sheet_VN.setTabColor("F1C232");
  sheet_VN.clear();
  sheet_VN.getRange(1, 1, tsvData_VN.length, tsvData_VN[0].length).setValues(tsvData_VN);


  ws.activate();



/*
//var credentials = ss.getRange("D3:D5").getValues(); // [5, 7, 9]

  var list = [[txtFileId_FN,txtFileId_SP,txtFileId_VN],['TXT FALSOS TECHOS', 'TXT SUPERFICIES', 'TXT VENTANAS']];
  Logger.log(list);
  Logger.log(Array.isArray(list));

    for (n in list) {
      var txtFileId = list[n][0];
      var txtFileSheet = list[n][1];

      var txt_file = DriveApp.getFileById(txtFileId);
      txt_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var tsvUrl = "https://drive.google.com/uc?id="+txtFileId+"&x=.tsv";
      var tsvContent = UrlFetchApp.fetch(tsvUrl,{muteHttpExceptions: true }).getContentText();
      var tsvData = Utilities.parseCsv(tsvContent, '\t');

      var sheetPaste = ss.getSheetByName(''+txtFileSheet) || ss.insertSheet(''+txtFileSheet, 200);
      sheetPaste.setTabColor("F1C232");
      sheetPaste.clear();
      sheetPaste.getRange(1, 1, tsvData.length, tsvData[0].length).setValues(tsvData);
    }

    ws.activate();


  var list = [];
  var txtNames = ws.getRange("C3:C5").getValues();
  var txtIDs = ws.getRange("D3:D5").getValues();
  var list2 = [];
  for (n in txtIDs) {
  var filess = DriveApp.getFileById(''+n);
  var filenames = filess.getName();
    list2.push([filenames]);
  Logger.log(n);

  }

  //list.push([txtNames, txtIDs, list2]);


  Logger.log(list2);

  // COPIA DE DATOS A HOJAS 🟢


    for (n in list) {
      var txtFileName = list[n][0];
      var txtFileId = list[n][1];
      var txtFileSheet = list[n][2];

      var txt_file = DriveApp.getFileById(txtFileId);
      txt_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var tsvUrl = "https://drive.google.com/uc?id="+txtFileId+"&x=.tsv";
      var tsvContent = UrlFetchApp.fetch(tsvUrl,{muteHttpExceptions: true }).getContentText();
      var tsvData = Utilities.parseCsv(tsvContent, '\t');

      var sheetPaste = ss.getSheetByName(''+txtFileSheet) || ss.insertSheet(''+txtFileSheet, 200);
      sheetPaste.setTabColor("F1C232");
      sheetPaste.clear();
      sheetPaste.getRange(1, 1, tsvData.length, tsvData[0].length).setValues(tsvData);
    }
*/


}








function congelarCuadroManual() {

  // BASE 🟢

    var ss = SpreadsheetApp.getActive();
    var sh = SpreadsheetApp.getActiveSheet();
    var ws = SpreadsheetApp.getActive().getSheetByName('ACTUALIZAR');
    var spreadsheetId = SpreadsheetApp.getActive().getId();
    var dateGMT = Utilities.formatDate(new Date(), "GMT+1", "yyyyMMdd");


  // DESTINATION FOLDER (Específico cuadro de superficies) 🟢

    var backupFolderId = ws.getRange(8,2).getValue();
    var backupFolderDef = DriveApp.getFolderById(backupFolderId);
    var backupFolderName = backupFolderDef.getName();
    ws.getRange('B7').setValue('=hyperlink("https://drive.google.com/corp/drive/folders/'+ backupFolderId +'";"' + backupFolderName + '")').setFontColor('#4A86E8');


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

    destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);


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

  // Delete the main DATA SHEET. 🟢

    destination.getSheets().forEach(function(sheet) {
    var sheetName = sheet.getSheetName();
    if (sheetName.indexOf("ACTUALIZAR") == -1) {
    }
       else {
        destination.deleteSheet(sheet);
      }
    });


  // Move file to the destination folder. 🟢

    var file = DriveApp.getFileById(destinationId);
    DriveApp.getFolderById(backupFolderId).addFile(file);
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
    blob.setName(ss.getName() + ".xlsx");

    var blobFile = UrlFetchApp.getRequest(url, params);

    ws.getRange('B9').setValue(url).setFontColor('#4A86E8');


  } catch (f) {
    Logger.log(f.toString());
  }


  ws.activate();

}
