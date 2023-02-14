/**
 * Gsuite Morph Tools - CS Freezer 1.8.0
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

function morphFreezer(btnID, sheetSelection) {

  var startTime = new Date().getTime(); var elapsedTime;

  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('LINK');
  var ss_id = ss.getId();

  var userMail = Session.getActiveUser().getEmail();
  var dateNow = Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy - HH:mm:ss');
  var freezerDate = Utilities.formatDate(new Date(), 'GMT+2', 'yyyyMMdd');
  var destName = `${ss.getName()} - ${freezerDate} - CONGELADO`;
  var functionErrors = [];
  
  var file = DriveApp.getFileById(ss_id);
  Logger.log(`FILE: ${file.getName()}, FILEURL: ${file.getUrl()}`);
  var parentFolder = file.getParents();
  var parentFolderID = parentFolder.next().getId();
  var backupFolderSearch = DriveApp.getFolderById(parentFolderID);

  // Start freezing process for each button type

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before get backup folder: ${elapsedTime} seconds.`);

  var destination, destinationId, destinationFile, destinationSheets, controlPanelID;

  if (btnID === 'csFreezer') {

  // Automatically get the backup folder

  if (btnID === 'csFreezer') {
    var backupFolder, backupFolderId, backupFolderName;
    let searchFor = 'title contains "Congelados"';
    backupFolder = backupFolderSearch.searchFolders(searchFor);
    backupFolder = backupFolder.next();
    backupFolderId = backupFolder.getId();
    backupFolderName = backupFolder.getName();
  }

  // Creating destination file

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before copy: ${elapsedTime} seconds.`);
  destination = ss.copy(destName);
  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time after copy: ${elapsedTime} seconds.`);
  destinationId = destination.getId();
  destinationFile = DriveApp.getFileById(destinationId);
  destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  destinationSheets = destination.getSheets();
  SpreadsheetApp.flush();

  // DBB Variables ImportRange Permission

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before permission: ${elapsedTime} seconds.`);

  var bddVariablesID = '1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro';
  var controlPanelID = getIdFromUrl(sh.getRange('B1').getValue());
  var permissionFilesIDs = [bddVariablesID, controlPanelID];
  var token = ScriptApp.getOAuthToken();

  var promises = permissionFilesIDs.map(function(fileID) {
    let url = `https://docs.google.com/spreadsheets/d/${destinationId}/externaldata/addimportrangepermissions?donorDocId=${fileID}`;
    let params = {
      method: 'post',
      headers: {
        Authorization: `Bearer ${token}`,
      },
      muteHttpExceptions: true,
    };
    return UrlFetchApp.fetch(url, params);
  });

  Promise.allSettled(promises)
    .then(function(results) {
      var errors = results.filter(result => result.status === 'rejected');
      if (errors.length > 0) {
        Logger.log('Se ha producido un error con las peticiones ImportRange: ', errors);
      } else {
        Logger.log('Las peticiones ImportRange se han enviado correctamente.');
      }
    }); 

  SpreadsheetApp.flush();

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time after permission: ${elapsedTime} seconds.`);

  // Check Google Forms links and delete them from destination file

  var formUrl = destinationSheets[0].getFormUrl();
  Logger.log(`HAS FORM?: ${formUrl}`);
  if (formUrl) {
    Logger.log(`Se ha ejecutado el FormURL`);
    FormApp.openByUrl(formUrl).removeDestination();
    let formID = getIdFromUrl(formUrl);
    DriveApp.getFileById(formID).setTrashed(true);
  };

  // FREEZE!

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before frozen proccess: ${elapsedTime} seconds.`);

  var visibleSheets, hiddenSheets;
  var excludedTabColors = ["#00ff00", "#ff0000", "#ffff00"];

  visibleSheets = destinationSheets.filter(tempsheet => !tempsheet.isSheetHidden() && !excludedTabColors.includes(tempsheet.getTabColor()));
  hiddenSheets = destinationSheets.filter(tempsheet => tempsheet.isSheetHidden() || excludedTabColors.includes(tempsheet.getTabColor()));

  var requests = [
    ...visibleSheets.map((sheet) => {
      var sheetid = sheet.getSheetId();
      var lastrow = sheet.getLastRow();
      var lastcol = sheet.getLastColumn();
      return {
        copyPaste: {
          source: {
            sheetId: sheetid,
            startRowIndex: 0,
            endRowIndex: lastrow,
            startColumnIndex: 0,
            endColumnIndex: lastcol
          },
          destination: {
            sheetId: sheetid,
            startRowIndex: 0,
            endRowIndex: 1,
            startColumnIndex: 0,
            endColumnIndex: 1
          },
          pasteType: "PASTE_VALUES",
          pasteOrientation: "NORMAL"
        }
      };
    }),
    ...hiddenSheets.map(sheet => {
      return {
        deleteSheet: {
          sheetId: sheet.getSheetId()
        }
      };
    })
  ];

  Sheets.Spreadsheets.batchUpdate({spreadsheetId: ss_id, requests}, destinationId);

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time after frozen proccess: ${elapsedTime} seconds.`);

  } else {

    // Alternative Workflow for Superfreezer and Partial Freezer

    destination = SpreadsheetApp.create(destName);
    destinationId = destination.getId();
    destinationFile = DriveApp.getFileById(destinationId);
    destination = SpreadsheetApp.openById(destination.getId());

    var sheetArray = ss.getSheets();
    var visibleSheets;

    if (btnID === 'superFreezerButton') {
      visibleSheets = sheetArray.filter(tempsheet => !tempsheet.isSheetHidden());
    } else if (btnID === 'actPartialFreezer') {
      visibleSheets = sheetArray.filter(tempsheet => {
        return sheetSelection.includes(tempsheet.getName());
      });
    }

    let visibleSheetsNames = visibleSheets.map((sheet) => {
      return sheet.getName();
    });
    Logger.log(`Visible Sheet Array: ${visibleSheetsNames}`);

    visibleSheets.map((sheet) => {
      let src = sheet.getDataRange();
      let a1Notation = src.getA1Notation();
      let values = src.getValues();
      let dstSheet = sheet.copyTo(destination).setName(`${sheet.getSheetName()}`);
      dstSheet.getRange(a1Notation).setValues(values);
    });

    destination.deleteSheet(destination.getSheetByName('Hoja 1'));
  }

  // Move file to the destination folder

  if (btnID === 'csFreezer') {
    if (backupFolderId !== undefined) {
      DriveApp.getFolderById(backupFolderId).addFile(destinationFile);
      destinationFile.getParents().next().removeFile(destinationFile);
    } else {
      DriveApp.getFolderById(parentFolderID).addFile(destinationFile);
      destinationFile.getParents().next().removeFile(destinationFile);
      functionErrors.push(`No se ha encontrado la carpeta 'PXXXXX-A-CS-Congelados' del proyecto, por lo que el archivo se ha guardado en la misma carpeta que el cuadro.`)
    }
  } else if (btnID === 'superFreezerButton' || btnID === 'actPartialFreezer') {
    DriveApp.getFolderById(parentFolderID).addFile(destinationFile);
    destinationFile.getParents().next().removeFile(destinationFile);
  };

  // EXCEL Conversion and LINK Sheet Data

  let url = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${destinationId}&exportFormat=xlsx`;

  if (btnID === 'csFreezer') {
    let backupFolderText = backupFolderId == undefined ? '' : `=hyperlink("https://drive.google.com/corp/drive/folders/${backupFolderId}";"${backupFolderName}")`;
    sh.getRange('B8').setValue(backupFolderText).setFontWeight('bold').setFontColor('#0000FF').setNote(null).setNote(`Último congelado: ${dateNow} por ${userMail}`); // Last Update Note
    sh.getRange('B9').setValue(backupFolderId);
    sh.getRange('B10').setValue(url).setFontColor('#0000FF').setFontWeight('normal'); // Add XLSX download url to sheet
    sh.activate();
  } else if (btnID === 'superFreezerButton' || btnID === 'actPartialFreezer') {
    let confirm = Browser.msgBox('Documento Excel', '¿Quieres crear una copia en formato Excel en la misma carpeta?', Browser.Buttons.OK_CANCEL);
    if (confirm == 'ok') {
      exportToXLSS(ss, url, freezerDate, parentFolderID);
    }
  }

  if (functionErrors.length > 0) {
    var ui = SpreadsheetApp.getUi();
    functionErrors.forEach(element => ui.alert('Advertencia', element, ui.ButtonSet.OK));
  }
}

/**
 * exportToXLSS
 * Crea un archivo XLSS a partir del ID de un archivo de Google Sheets
 */
function exportToXLSS(ss, url, freezerDate, parentFolderID) {
  try {

    let params = {
      method: 'get',
      headers: {Authorization: `Bearer ${ScriptApp.getOAuthToken()}`},
      muteHttpExceptions: true,
    };

    let blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(`${ss.getName()} - ${freezerDate} - CONGELADO.xlsx`);
    DriveApp.getFolderById(parentFolderID).createFile(blob);
  } catch (f) {
    // Logger.log(f.toString());
  }
}

/**
 * morphFastFreezer
 * Función de pruebas para el congelador Morph
 */
function morphFastFreezer() {

}
