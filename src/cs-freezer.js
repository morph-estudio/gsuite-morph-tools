/**
 * Gsuite Morph Tools - CS Freezer 1.2
 * Developed by alsanchezromero
 * Created on Mon Jul 25 2022
 *
 * Copyright (c) 2022 Morph Estudio
 */

function congeladorMorph(btnID) {
  let ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('ACTUALIZAR');
  let ss_id = ss.getId();
  let file = DriveApp.getFileById(ss_id);
  let parentFolder = file.getParents();
  let parentFolderID = parentFolder.next().getId();
  let backupFolderSearch = DriveApp.getFolderById(parentFolderID);
  const dateNow = Utilities.formatDate(new Date(), 'GMT+1', 'yyyyMMdd');
  
  let backupFolderId;
  if (btnID === 'csFreezer' || 'csManual3') {
    backupFolderId = freezerCS(sh, backupFolderSearch, btnID); // Destination Folder (Específico cuadro de superficies)
  }

  let tempSheets = createTemporalSheets(ss) // Copy each sheet in the source Spreadsheet by removing the formulas as the temporal sheets

  // Copy the source Spreadsheet

  let destination = ss.copy(`${ss.getName()} - ${dateNow} - CONGELADO`);
  let destinationId = destination.getId();
  let destinationFile = DriveApp.getFileById(destinationId);
  destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Delete the temporal sheets in the source Spreadsheet

  tempSheets.forEach((sheet) => {
    ss.deleteSheet(sheet);
  });

  SpreadsheetApp.flush();

  // Delete the original sheets from the copied Spreadsheet and rename the copied sheets

  deleteAndRenameSheets(destination);

  // Delete the main DATA SHEET

  destination.getSheets().forEach((sheet) => {
    sheetName = sheet.getSheetName();
    if (sheetName.indexOf('ACTUALIZAR') > -1) {
      destination.deleteSheet(sheet);
    }
  });

  // Move file to the destination folder

  if (btnID === 'superFreezerButton') {
    file = DriveApp.getFileById(destinationId);
    DriveApp.getFolderById(parentFolderID).addFile(file);
    file.getParents().next().removeFile(file);
  } 
  if (btnID == 'csFreezer' || 'csManual3') {
    file = DriveApp.getFileById(destinationId);
    DriveApp.getFolderById(backupFolderId).addFile(file);
    file.getParents().next().removeFile(file);
  }

  // Export to XLSX let xlsDownloadUrl = exportToXLSS(ss, destinationId);

  let url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${destinationId}&exportFormat=xlsx`;

  if (btnID === 'superFreezerButton') {
    let params = {
      method: "get",
      headers: {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true,
    };
    let blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(`${ss.getName()} - ${dateNow} - ` + `CONGELADO` + `.xlsx`);
    const ui = SpreadsheetApp.getUi();
    let confirm = Browser.msgBox('Documento Excel', '¿Quieres crear una copia en formato Excel en la misma carpeta?', Browser.Buttons.OK_CANCEL);
    if (confirm == 'ok') { DriveApp.getFolderById(parentFolderID).createFile(blob); }
  } else if (btnID === 'csFreezer') {
      sh.getRange('B9').setValue(url).setFontColor('#4A86E8'); // Add XLSX download url to sheet
  }

  sh.activate();
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
      Logger.log('brutas: ' + backupFolderId)
    }
  } else if (btnID === 'csManual3') {
      backupFolderId = sh.getRange(8, 2).getValue();
      backupFolderDef = DriveApp.getFolderById(backupFolderId);
      backupFolderName = backupFolderDef.getName();
      Logger.log('frutas: ' + backupFolderId)
  }

  sh.getRange('B7').setFontWeight('bold').setFontColor('#B7B7B7');
  sh.getRange('A7').setValue('CARPETA BACKUP');
  sh.getRange('A8').setValue('ID CARPETA BACKUP');
  sh.getRange('A9').setValue('DESCARGAR ARCHIVO XLSX');
  sh.getRange('B7').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${backupFolderId}";"${backupFolderName}")`).setFontColor('#4A86E8');
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
    Logger.log(f.toString());
  }
}
