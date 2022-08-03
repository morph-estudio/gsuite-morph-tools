function doGet() {
  return HtmlService.createHtmlOutputFromFile('html/index');
}

function tplActualizarCuadroManual() {
  const ss = SpreadsheetApp.getActive();
  const sh = SpreadsheetApp.getActiveSheet();
  let ws = ss.getSheetByName('ACTUALIZAR') || ss.insertSheet('ACTUALIZAR', 1);
  let ss_id = ss.getId();
  let file = DriveApp.getFileById(ss_id);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  ws.clear().clearFormats();

  tplText(ws);

  ws.getRange('C2').setValue('Archivos exportados');
  ws.getRange('D2').setValue('IDs');
  ws.getRange('C3').setValue('TXT Sheets Falsos techos.txt');
  ws.getRange('C4').setValue('TXT Sheets Superficies.txt');
  ws.getRange('C5').setValue('TXT Sheets Ventanas.txt');

  sheetFormatter(ws); // Formato de celdas

  // Control Panel
  ws.getRange('B1').setBackground('#FFF2CC').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#BF9000')
    .setFontWeight('bold');
  ws.getRange('B8').setBackground('#FFF2CC').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#BF9000');
  // Filelist Style
  ws.getRange(3, 3, 3, 2).setFontSize(13).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setFontColor('#B7B7B7');
  ws.getRange(3, 3, 3, 1).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor('#B7B7B7').setFontWeight('bold');
  ws.getRange(3, 4, 3, 1).setBackground('#FFF2CC').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#BF9000');

  // ImportRange Cells

  ws.getRange('C1').setValue('=IMPORTRANGE(B1;"Instrucciones!A1")');
  ws.getRange('D1').setValue('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"DB-SI!B2")');

  deleteEmptyRows();
  removeEmptyColumns();
  ws.activate();
}

function actualizarCuadroManual() {
  const ss = SpreadsheetApp.getActive();
  const sh = SpreadsheetApp.getActiveSheet();
  let ws = ss.getSheetByName('ACTUALIZAR') || ss.insertSheet('ACTUALIZAR', 1);
  let ss_id = ss.getId();
  let file = DriveApp.getFileById(ss_id);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  let filePanelUrl = ws.getRange('B1').getValue();
  let filePanelId = filePanelUrl.replace(/.*\/d\//, '').replace(/\/.*/,  '');
  let panelControl = DriveApp.getFileById(filePanelId);
  let folderPanelc = panelControl.getParents();
  let folderPanelcDef = folderPanelc.next();
  let folderPanelcId = folderPanelcDef.getId();
  let parents = file.getParents();
  let carpetaBaseID = parents.next().getId();
  let carpetaBase = DriveApp.getFolderById(carpetaBaseID);

  ws.getRange('B5').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${folderPanelcId}";"${folderPanelcDef}")`).setFontColor('#4A86E8');
  ws.getRange('B6').setValue(folderPanelcId);
  ws.getRange('B3').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${carpetaBaseID}";"${carpetaBase}")`).setFontColor('#4A86E8');
  ws.getRange('B4').setValue(carpetaBaseID);

  // ImportRange Permission

  let sectoresID = '1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro';

  importRangeToken(ss_id, carpetaBaseID);
  importRangeToken(ss_id, sectoresID);

  let txtFileId_FT = ws.getRange(3, 4).getValue();
  let txtFileId_SP = ws.getRange(4, 4).getValue();
  let txtFileId_VN = ws.getRange(5, 4).getValue();

  let txtFile_FT = DriveApp.getFileById(txtFileId_FT);
  txtFile_FT.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  let txtFile_SP = DriveApp.getFileById(txtFileId_SP);
  txtFile_SP.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  let txtFile_VN = DriveApp.getFileById(txtFileId_VN);
  txtFile_VN.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // FALSOS TECHOS

  let tsvUrl_FT = `https://drive.google.com/uc?id=${txtFileId_FT}&x=.tsv`;
  let tsvContent_FT = UrlFetchApp.fetch(tsvUrl_FT, { muteHttpExceptions: true }).getContentText();
  let tsvData_FT = Utilities.parseCsv(tsvContent_FT, '\t');

  let sheet_FT = ss.getSheetByName('TXT FALSOS TECHOS') || ss.insertSheet('TXT FALSOS TECHOS', 200);
  sheet_FT.setTabColor('F1C232');
  sheet_FT.clear();
  sheet_FT.getRange(1, 1, tsvData_FT.length, tsvData_FT[0].length).setValues(tsvData_FT);

  // TXT SUPERFICIES

  let tsvUrl = `https://drive.google.com/uc?id=${txtFileId_SP}&x=.tsv`;
  let tsvContent = UrlFetchApp.fetch(tsvUrl, {muteHttpExceptions: true }).getContentText();
  let tsvData = Utilities.parseCsv(tsvContent, '\t');

  let sheet_SP = ss.getSheetByName('TXT SUPERFICIES') || ss.insertSheet('TXT SUPERFICIES', 200);
  sheet_SP.setTabColor('F1C232');
  sheet_SP.clear();
  sheet_SP.getRange(1, 1, tsvData.length, tsvData[0].length).setValues(tsvData);

  // VENTANAS

  let tsvUrl_VN = `https://drive.google.com/uc?id=${txtFileId_VN}&x=.tsv`;
  let tsvContent_VN = UrlFetchApp.fetch(tsvUrl_VN, {muteHttpExceptions: true }).getContentText();
  let tsvData_VN = Utilities.parseCsv(tsvContent_VN, '\t');

  let sheet_VN = ss.getSheetByName('TXT VENTANAS') || ss.insertSheet('TXT VENTANAS', 200);
  sheet_VN.setTabColor("F1C232");
  sheet_VN.clear();
  sheet_VN.getRange(1, 1, tsvData_VN.length, tsvData_VN[0].length).setValues(tsvData_VN);

  ws.activate();
}

function congelarCuadroManual() {
  const ss = SpreadsheetApp.getActive();
  const sh = SpreadsheetApp.getActiveSheet();
  const dateGMT = Utilities.formatDate(new Date(), 'GMT+1', 'yyyyMMdd');
  let ws = SpreadsheetApp.getActive().getSheetByName('ACTUALIZAR');
  let ss_id = ss.getId();

  // Destination Folder (Específico cuadro de superficies)

  let backupFolderId = ws.getRange(8, 2).getValue();
  let backupFolderDef = DriveApp.getFolderById(backupFolderId);
  let backupFolderName = backupFolderDef.getName();
  ws.getRange('B7').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${backupFolderId}";"${backupFolderName}")`).setFontColor('#4A86E8');

  // Copy each sheet in the source Spreadsheet by removing the formulas as the temporal sheets

  let tempSheets = ss.getSheets().filter(sh => !sh.isSheetHidden()).map((sheet) => {
    let dstSheet = sheet.copyTo(ss).setName(`${sheet.getSheetName()}_temp`);
    let src = dstSheet.getDataRange();
    src.copyTo(src, { contentsOnly: true });
    return dstSheet;
  });

  // Copy the source Spreadsheet

  let destination = ss.copy(ss.getName() + " - " + dateGMT + " - " + "CONGELADO");
  let destinationId = destination.getId();
  let destinationFile = DriveApp.getFileById(destinationId);

  destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Delete the temporal sheets in the source Spreadsheet

  tempSheets.forEach((sheet) => { ss.deleteSheet(sheet); });
  SpreadsheetApp.flush();

  // Delete the original sheets from the copied Spreadsheet and rename the copied sheets

  destination.getSheets().forEach((sheet) => {
    if (sheet.isSheetHidden()) {
      destination.deleteSheet(sheet);
      SpreadsheetApp.flush();
    } else {
      var sheetName = sheet.getSheetName();
      //Logger.log(sheetName);
      if (sheetName.indexOf('_temp') == -1) {
        destination.deleteSheet(sheet);
        SpreadsheetApp.flush();
      } else {
        sheet.setName(sheetName.replace('_temp', ''));
      }
    }
  });

  // Delete the main DATA SHEET

  destination.getSheets().forEach((sheet) => {
    var sheetName = sheet.getSheetName();
    if (sheetName.indexOf("ACTUALIZAR") == -1) {
    } else {
      destination.deleteSheet(sheet);
    }
  });

  // Move file to the destination folder.

  var file = DriveApp.getFileById(destinationId);
  DriveApp.getFolderById(backupFolderId).addFile(file);
  file.getParents().next().removeFile(file);

  // Export to XLSX.

  try {
    let url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + destination.getId() + "&exportFormat=xlsx";
    let params = {
      method: "get",
      headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    };

    let blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(ss.getName() + ".xlsx");
    let blobFile = UrlFetchApp.getRequest(url, params);

    ws.getRange('B9').setValue(url).setFontColor('#4A86E8');
  } catch (f) {
    Logger.log(f.toString());
  }

  ws.activate();
}
