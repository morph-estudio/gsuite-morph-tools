function doGet() {
  return HtmlService.createHtmlOutputFromFile('client/index');
}

/*
 * Gsuite Morph Tools - CS Superfreezer
 * Developed by alsanchezromero
 * Created on Mon Jul 25 2022
 *
 * Copyright (c) 2022 Morph Estudio
*/

function congeladorMorph() {
  const ss = SpreadsheetApp.getActive();
  const ws = SpreadsheetApp.getActiveSheet();
  const dateGMT = Utilities.formatDate(new Date(), 'GMT+1', 'yyyyMMdd');
  let ss_id = SpreadsheetApp.getActive().getId();
  let file = DriveApp.getFileById(ss_id);
  let parentFolder = file.getParents();
  let parentFolder_ID = parentFolder.next().getId();
  let parentFolderName = file.getParents().next().getName();

  // Copy each sheet in the source Spreadsheet by removing the formulas as the temporal sheets

  let tempSheets = ss.getSheets().filter(sh => !sh.isSheetHidden()).map((sheet) => {
    let dstSheet = sheet.copyTo(ss).setName(`${sheet.getSheetName()}_temp`);
    let src = dstSheet.getDataRange();
    src.copyTo(src, { contentsOnly: true });
    return dstSheet;
  });

  // Copy the source Spreadsheet

  let destination = ss.copy(`${ss.getName()} - ${dateGMT} - ` + `CONGELADO`);
  let destinationId = destination.getId();
  let destinationFile = DriveApp.getFileById(destinationId);

  // Delete the temporal sheets in the source Spreadsheet

  tempSheets.forEach((sheet) => { ss.deleteSheet(sheet); });
  SpreadsheetApp.flush();

  // Delete the original sheets from the copied Spreadsheet and rename the copied sheets

  destination.getSheets().forEach((sheet) => {
    if (sheet.isSheetHidden()) {
      destination.deleteSheet(sheet);
      SpreadsheetApp.flush();
    } else {
      let sheetName = sheet.getSheetName();
      //Logger.log(sheetName);
      if (sheetName.indexOf('_temp') == -1) {
        destination.deleteSheet(sheet);
        SpreadsheetApp.flush();
      } else {
        sheet.setName(sheetName.replace('_temp', ''));
      }
    }
  });

  // Move file to the destination folder.

  file = DriveApp.getFileById(destinationId);
  DriveApp.getFolderById(parentFolder_ID).addFile(file);
  file.getParents().next().removeFile(file);

  // Export to XLSX

  try {
    let url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${destination.getId()}&exportFormat=xlsx`;

    let params = {
      method: "get",
      headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true,
    };

    let blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(`${ss.getName()} - ${dateGMT} - ` + `CONGELADO` + `.xlsx`);

    let ui = SpreadsheetApp.getUi();
    let confirm = Browser.msgBox('Documento Excel', '¿Quieres crear una copia en formato XLSX en la misma carpeta?', Browser.Buttons.OK_CANCEL);
    if (confirm == 'ok') { DriveApp.getFolderById(parentFolder_ID).createFile(blob); }

    // MailApp.sendEmail('asanchez@morphestudio.es', 'Conversión de Google Sheet a Excel', 'El archivo XLSX aparece adjunto a este correo.', { attachments: [blob] });
  } catch (f) {
    Logger.log(f.toString());
  }
}
