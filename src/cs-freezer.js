/**
 * Gsuite Morph Tools - CS Freezer 1.5.0
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

function morphFreezer(btnID) {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('LINK');
  let ss_id = ss.getId();
  let file = DriveApp.getFileById(ss_id);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  let parentFolder = file.getParents();
  let parentFolderID = parentFolder.next().getId();
  let backupFolderSearch = DriveApp.getFolderById(parentFolderID);
  let userMail = Session.getActiveUser().getEmail();
  let dateNow = Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy - HH:mm:ss');
  let freezerDate = Utilities.formatDate(new Date(), 'GMT+2', 'yyyyMMdd');

  let backupFolderId;
  if (btnID === 'csFreezer' || btnID === 'csManual3') {
    backupFolderId = freezerCS(sh, backupFolderSearch, btnID); // Destination Folder (Específico cuadro de superficies)
  }

  let tempSheets = createTemporalSheets(ss) // Copy each sheet in the source Spreadsheet by removing the formulas as the temporal sheets

  // Copy the source Spreadsheet

  let destination = ss.copy(`${ss.getName()} - ${freezerDate} - CONGELADO`);
  let destinationId = destination.getId();
  let destinationFile = DriveApp.getFileById(destinationId);
  destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Delete the temporal sheets in the source Spreadsheet

  tempSheets.forEach((sheet) => {
    ss.deleteSheet(sheet);
  });

  SpreadsheetApp.flush();

  // Delete the original sheets from the copied Spreadsheet and rename the copied sheets

  let formUrl = destination.getSheets()[0].getFormUrl(); // Remove form links
  if (formUrl) {
    FormApp.openByUrl(formUrl).removeDestination();
    let formID = getIdFromUrl(formUrl);
    DriveApp.getFileById(formID).setTrashed(true);
  };

  deleteAndRenameSheets(destination);

  // Delete the main DATA SHEET

  destination.getSheets().forEach((sheet) => {
    sheetName = sheet.getSheetName();
    if (sheetName.indexOf('LINK') > -1) {
      destination.deleteSheet(sheet);
    }
  });

  // Move file to the destination folder

  if (btnID === 'superFreezerButton') {
    DriveApp.getFolderById(parentFolderID).addFile(destinationFile);
    destinationFile.getParents().next().removeFile(destinationFile);
  };
  if (btnID === 'csFreezer' || btnID === 'csManual3') {
    DriveApp.getFolderById(backupFolderId).addFile(destinationFile);
    destinationFile.getParents().next().removeFile(destinationFile);
  };

  // Export to XLSX let xlsDownloadUrl = exportToXLSS(ss, destinationId);

  let url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${destinationId}&exportFormat=xlsx`;

  if (btnID === 'superFreezerButton') {
    let params = {
      method: "get",
      headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true,
    };
    let blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(`${ss.getName()} - ${freezerDate} - CONGELADO.xlsx`);
    const ui = SpreadsheetApp.getUi();
    let confirm = Browser.msgBox('Documento Excel', '¿Quieres crear una copia en formato Excel en la misma carpeta?', Browser.Buttons.OK_CANCEL);
    if (confirm == 'ok') { DriveApp.getFolderById(parentFolderID).createFile(blob); }
  } else if (btnID === 'csFreezer' || 'csManual3') {
    sh.getRange('B9').setValue(url).setFontColor('#0000FF'); // Add XLSX download url to sheet
    sh.getRange('B7').setNote(null).setNote(`Último congelado: ${dateNow} por ${userMail}`); // Last Update Note
    sh.activate();
  }

}

/**
 * freezerCS
 * Opciones específicas del congelador para el Cuadro de Superficies
 */
function freezerCS(sh, backupFolderSearch, btnID) {
  let backupFolderId;
  let backupFolderName;

  if (btnID === 'csFreezer') {
    let searchFor = 'title contains "Congelados"';
    let backupFolder = backupFolderSearch.searchFolders(searchFor);
    let backupFolderDef;

    while (backupFolder.hasNext()) {
      backupFolderDef = backupFolder.next();
      backupFolderId = backupFolderDef.getId();
      backupFolderName = backupFolderDef.getName();
      backupFolderUrl = backupFolderDef.getUrl();
    }
  } else if (btnID === 'csManual3') {
    backupFolderId = sh.getRange(8, 2).getValue();
    backupFolderDef = DriveApp.getFolderById(backupFolderId);
    backupFolderName = backupFolderDef.getName();
  }

  sh.getRange('B7').setFontWeight('bold').setFontColor('#B7B7B7');
  sh.getRange('A7').setValue('CARPETA BACKUP');
  sh.getRange('A8').setValue('ID CARPETA BACKUP');
  sh.getRange('A9').setValue('DESCARGAR ARCHIVO XLSX');
  sh.getRange('B7').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${backupFolderId}";"${backupFolderName}")`).setFontColor('#0000FF');
  sh.getRange('B8').setValue(backupFolderId);

  return backupFolderId;
}

/**
 * createTemporalSheets
 * Copia todas las hojas eliminando todas las fórmulas para dejar resultados planos
 */
function createTemporalSheets(ss) {
  let dstSheet; let src;
  let tempSheets = ss.getSheets().filter((sh) => !sh.isSheetHidden()).map((sheet) => {
    dstSheet = sheet.copyTo(ss).setName(`${sheet.getSheetName()}_temp`);
    src = dstSheet.getDataRange();
    src.copyTo(src, { contentsOnly: true });
    return dstSheet;
  });
  return tempSheets;
}

/**
 * deleteAndRenameSheets
 * Crea un archivo XLSS a partir del ID de un archivo de Google Sheets
 */
function deleteAndRenameSheets(destination) {
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
}

/**
 * exportToXLSS
 * Crea un archivo XLSS a partir del ID de un archivo de Google Sheets
 */
function exportToXLSS(ss, destinationId) {
  try {
    let url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${destinationId}&exportFormat=xlsx`;

    let params = {
      method: 'get',
      headers: {Authorization: `Bearer ${ScriptApp.getOAuthToken()}`},
      muteHttpExceptions: true,
    };

    let blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(`${ss.getName()}.xlsx`);
    UrlFetchApp.getRequest(url, params);

    return url;
  } catch (f) {
    // Logger.log(f.toString());
  }
}

function morphFastFreezer() {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('LINK');
  let ss_id = ss.getId();
  let file = DriveApp.getFileById(ss_id);
  let fileURL = file.getUrl();
  let parentFolder = file.getParents();
  let parentFolderID = parentFolder.next().getId();
  let parentFolderDef = DriveApp.getFolderById(parentFolderID);
  let freezerDate = Utilities.formatDate(new Date(), 'GMT+2', 'yyyyMMdd');

  var destination = ss.copy(ss.getName());
  file.setName(`${ss.getName()} - ${freezerDate} - CONGELADO`);

  let destinationId = destination.getId();
  let destinationFile = DriveApp.getFileById(destinationId);
  destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  SpreadsheetApp.flush();

  let tempSheets = ss.getSheets().filter((sht) => !sht.isSheetHidden()).map((sheet) => {
    let dstSheet = sheet.getDataRange();
    dstSheet.copyTo(dstSheet, { contentsOnly: true });
    return dstSheet;
  });

  tempSheets = ss.getSheets().filter((sht) => sht.isSheetHidden()).map((sheet) => {
    ss.deleteSheet(sheet);
  });

  let fileOriginal = DriveApp.getFileById(destinationId);

  SpreadsheetApp.flush();

  let url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${destinationId}&exportFormat=xlsx`;

  let params = {
    method: "get",
    headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
    muteHttpExceptions: true,
  };
  let blob = UrlFetchApp.fetch(url, params).getBlob();
  blob.setName(`${ss.getName()}.xlsx`);
  const ui = SpreadsheetApp.getUi();
  let confirm = Browser.msgBox('Documento Excel', '¿Quieres crear una copia en formato Excel en la misma carpeta?', Browser.Buttons.OK_CANCEL);
  if (confirm == 'ok') { DriveApp.getFolderById(parentFolderID).createFile(blob); }

  SpreadsheetApp.flush();

  file.moveTo(parentFolderDef);
  fileOriginal.moveTo(parentFolderDef);
  let fileOriginalUrl = fileOriginal.getUrl();

  openExternalUrlFromMenu(fileOriginalUrl);

}
