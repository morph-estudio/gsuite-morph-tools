function doGet() {
  return HtmlService.createHtmlOutputFromFile('client/index');
}

/*
 * Gsuite Morph Tools - CS Freezer 1.1
 * Developed by alsanchezromero
 * Created on Mon Jul 25 2022
 *
 * Copyright (c) 2022 Morph Estudio
*/

function congelarCuadro() {
  let ss = SpreadsheetApp.getActive();
  let sh = SpreadsheetApp.getActiveSheet();
  let ws = SpreadsheetApp.getActive().getSheetByName('ACTUALIZAR');
  let ss_id = ss.getId();
  let file = DriveApp.getFileById(ss_id);
  let parentFolder = file.getParents();
  let parentFolder_ID = parentFolder.next().getId();
  let backupFolderSearch = DriveApp.getFolderById(parentFolder_ID);
  const dateGMT = Utilities.formatDate(new Date(), "GMT+1", "yyyyMMdd");

  // Destination Folder (Específico cuadro de superficies)

  let searchFor = 'title contains "Congelados"';
  let names = [];
  let backupFolderIds= [];
  let backupFolder = backupFolderSearch.searchFolders(searchFor);
  let backupFolderDef; let backupFolderId; let backupFolderName; let backupFolderUrl;

  while (backupFolder.hasNext()) {
    backupFolderDef = backupFolder.next();
    backupFolderId = backupFolderDef.getId();
    backupFolderIds.push(backupFolderId);
    backupFolderName = backupFolderDef.getName();
    backupFolderUrl = backupFolderDef.getUrl();
    names.push(backupFolderName);
  }

  ws.getRange('B7').setFontWeight('bold').setFontColor('#B7B7B7');
  ws.getRange('A7').setValue('CARPETA BACKUP');
  ws.getRange('A8').setValue('ID CARPETA BACKUP');
  ws.getRange('A9').setValue('DESCARGAR ARCHIVO XLSX');
  ws.getRange('B7').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${backupFolderId}";"${backupFolderName}")`).setFontColor('#4A86E8');
  ws.getRange('B8').setValue(backupFolderId);

  // Copy each sheet in the source Spreadsheet by removing the formulas as the temporal sheets

  let dstSheet; let src;
  let tempSheets = ss.getSheets().filter((sh) => !sh.isSheetHidden()).map((sheet) => {
    dstSheet = sheet.copyTo(ss).setName(`${sheet.getSheetName()}_temp`);
    src = dstSheet.getDataRange();
    src.copyTo(src, { contentsOnly: true });
    return dstSheet;
  });

  // Copy the source Spreadsheet

  let destination = ss.copy(`${ss.getName()} - ${dateGMT} - ` + `CONGELADO`);
  let destinationId = destination.getId();
  let destinationFile = DriveApp.getFileById(destinationId);

  destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Delete the temporal sheets in the source Spreadsheet

  tempSheets.forEach((sheet) => {
    ss.deleteSheet(sheet);
  });

  SpreadsheetApp.flush();

  // Delete the original sheets from the copied Spreadsheet and rename the copied sheets

  let sheetName;
  destination.getSheets().forEach((sheet) => {
    if (sheet.isSheetHidden()) {
      destination.deleteSheet(sheet);
      SpreadsheetApp.flush();
    }
    else {
      sheetName = sheet.getSheetName();
      if (sheetName.indexOf('_temp') === -1) {
        destination.deleteSheet(sheet);
        SpreadsheetApp.flush();
      }
      else {
        sheet.setName(sheetName.replace('_temp', ''));
      }
    }
  });

  // Delete the main DATA SHEET

  destination.getSheets().forEach((sheet) => {
    sheetName = sheet.getSheetName();
    if (sheetName.indexOf('ACTUALIZAR') === -1) {
    } else {
      destination.deleteSheet(sheet);
    }
  });

  // Move file to the destination folder

  file = DriveApp.getFileById(destinationId);
  DriveApp.getFolderById(backupFolderId).addFile(file);
  file.getParents().next().removeFile(file);

  // Export to XLSX

  try {
    let url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${destination.getId()}&exportFormat=xlsx`;

    let params = {
      method: 'get',
      headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true,
    };

    let blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(`${ss.getName()}.xlsx`);

    //let blobFile = UrlFetchApp.getRequest(url, params);
    UrlFetchApp.getRequest(url, params);
    ws.getRange('B9').setValue(url).setFontColor('#4A86E8');
    // eslint-disable-next-line max-len
    // MailApp.sendEmail('asanchez@morphestudio.es', 'Conversión de Google Sheet a Excel', 'El archivo XLSX aparece adjunto a este correo.', { attachments: [blob] });
  } catch (f) {
    Logger.log(f.toString());
  }

  ws.activate();
}
