/**
 * Gsuite Morph Tools - CS Freezer 1.8.0
 * Developed by alsanchezromero
 *
 * Morph Estudio, 2023
 */

function morphFreezer(btnID, sheetSelection) {

  var startTime = new Date().getTime(); var elapsedTime;

  const ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('LINK');
  var ss_id = ss.getId();
  var ss_name = ss.getName();

  var userMail = Session.getActiveUser().getEmail();
  var dateNow = Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy - HH:mm:ss');
  var freezerDate = Utilities.formatDate(new Date(), 'GMT+2', 'yyyyMMdd');
  var destName = `${ss_name} - ${freezerDate} - CONGELADO`;
  var functionErrors = [];
  
  var file = DriveApp.getFileById(ss_id);
  Logger.log(`FILE: ${ss_name}, FILEURL: ${file.getUrl()}`);
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

    if (!backupFolder.hasNext()) { // La carpeta no fue encontrada
      var response = Browser.msgBox("Atención", "No se ha encontrado la carpeta para los archivos congelados. ¿Deseas crearla automáticamente?", Browser.Buttons.OK_CANCEL);
      if (response == "cancel") {
        throw new Error(`No se ha podido encontrar la carpeta de archivos congelados.`);
      }

      var folderName = `${ss_name.substring(0, 6)}-A-CS-${searchFor.substring(16, searchFor.length - 1)}`; // Obtener el nombre de la carpeta desde la cadena de búsqueda
      var newFolder = backupFolderSearch.createFolder(folderName);

      backupFolderId = newFolder.getId();
      backupFolder = DriveApp.getFolderById(backupFolderId);
      backupFolder = [backupFolder]; // Actualizar la variable expFolder para usar la carpeta recién creada
      
    } else { // La carpeta fue encontrada
      backupFolder = backupFolder.next();
      backupFolderId = backupFolder.getId();
    }

    var backupFolderName = backupFolder.getName();
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

  var { hyperlinkFontColor } = formatVariables();

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

  var requests;

  requests = [
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

  SpreadsheetApp.flush();

  if (ss_name.toLowerCase().indexOf("exportación superficies") !== -1) {
    formatearFileExportacion(destination);
  } else { Logger.log("El archivo no es un cuadro de exportación"); }

  } else {

    // Alternative Workflow for Superfreezer and Partial Freezer

/**/
  var sheetArray = ss.getSheets();
  var visibleSheets;

  if (btnID === 'superFreezerButton') {
    visibleSheets = sheetArray.filter(tempsheet => !tempsheet.isSheetHidden());
  } else if (btnID === 'actPartialFreezer') {
    visibleSheets = sheetArray.filter(tempsheet => {
      return sheetSelection.includes(tempsheet.getName());
    });
  }

  var tempSheets = visibleSheets.map(function(sheet) {
    var dstSheet = sheet.copyTo(ss).setName(sheet.getSheetName() + "_temp");
    var src = dstSheet.getDataRange();
    src.copyTo(src, {contentsOnly: true});
    return dstSheet;
  });
  
  // Copy the source Spreadsheet.
  var destination = ss.copy(ss_name + " - " + freezerDate);

    destinationId = destination.getId();
    destinationFile = DriveApp.getFileById(destinationId);
    destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    destination = SpreadsheetApp.openById(destination.getId());
  
  // Delete the temporal sheets in the source Spreadsheet.
  tempSheets.forEach(function(sheet) {ss.deleteSheet(sheet)});
  
  // Delete the original sheets from the copied Spreadsheet and rename the copied sheets.
  var destsheets = destination.getSheets();
  for (var i = 0; i < destsheets.length; i++) {
    var sheet = destsheets[i];
    var sheetName = sheet.getSheetName();
    if (sheetName.indexOf("_temp") == -1) {
      destination.deleteSheet(sheet);
    } else {
      SpreadsheetApp.flush()
      sheet.setName(sheetName.replace('_temp',''));
    }
  }

/*

    destination = SpreadsheetApp.create(destName);
    destinationId = destination.getId();
    destinationFile = DriveApp.getFileById(destinationId);
    destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    destination = SpreadsheetApp.openById(destination.getId());

    Logger.log(`destinationId: ${destinationId}`);

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
      let a1Notation = src.getA1Notation(); Logger.log(`a1Notation: ${a1Notation}`);
      let values = src.getValues();
      sheet.copyTo(destination).setName(`${sheet.getSheetName()}`);
      
      let newSheet = destination.getSheetByName(`${sheet.getSheetName()}`); Logger.log(`newSheet: ${newSheet.getName()}`);
      newSheet.getRange(a1Notation).setValues(values);
    });

    destination.deleteSheet(destination.getSheetByName('Hoja 1'));

*/

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

  let excelUrl = `https://docs.google.com/feeds/download/spreadsheets/Export?key=${destinationId}&exportFormat=xlsx`;

  if (btnID === 'csFreezer') {
    let backupFolderText = backupFolderId == undefined ? '' : `=hyperlink("https://drive.google.com/corp/drive/folders/${backupFolderId}";"${backupFolderName}")`;

  var data = sh.getRange("A:A").getValues();
  for (var i = 0; i < data.length; i++) {
    var cellValue = data[i][0];
    if (cellValue === "CARPETA CONGELADOS" || cellValue === "CARPETA BACKUP") {
      sh.getRange(i + 1, 2).setValue(backupFolderText).setFontWeight('bold').setFontColor(hyperlinkFontColor).setNote(null).setNote(`Último congelado: ${dateNow} por ${userMail}`); // Last Update Note;
    } else if (cellValue === "ID CARPETA CONGELADOS" || cellValue === "ID CARPETA BACKUP") {
      sh.getRange(i + 1, 2).setValue(backupFolderId);
    } else if (cellValue === "ÚLTIMO ARCHIVO CONGELADO") {
      sh.getRange(i + 1, 2).setValue(destinationFile.getUrl()).setFontColor(hyperlinkFontColor).setFontWeight('normal');
    } else if (cellValue === "DESCARGAR ARCHIVO XLSX") {
      sh.getRange(i + 1, 2).setValue(excelUrl).setFontColor(hyperlinkFontColor).setFontWeight('normal'); // Add XLSX download url to sheet;
    }
  }

    sh.activate();
  } else if (btnID === 'superFreezerButton' || btnID === 'actPartialFreezer') {
    let confirm = Browser.msgBox('Documento Excel', '¿Quieres crear una copia en formato Excel en la misma carpeta?', Browser.Buttons.OK_CANCEL);
    if (confirm == 'ok') {
      var fileName = `${ss.getName()} - ${freezerDate} - CONGELADO.xlsx`;
      exportToXLSS(fileName, excelUrl, parentFolderID);
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
function exportToXLSS(fileName, url, parentFolderID) {
  try {

    let params = {
      method: 'get',
      headers: {Authorization: `Bearer ${ScriptApp.getOAuthToken()}`},
      muteHttpExceptions: true,
    };

    let blob = UrlFetchApp.fetch(url, params).getBlob();
    blob.setName(fileName);
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

function formatearFileExportacion(destinationFile) {

  var destinationSheets = destinationFile.getSheets();
  
  for (var i = 0; i < destinationSheets.length; i++) {
    var sheet = destinationSheets[i];

    var cellA1 = sheet.getRange("A1").getValue().toLowerCase();
    if (cellA1.indexOf("selecciona resumen") !== -1) {
      deleteColumnsInsideColumnGroup(sheet);
      removeEmptyColumns(sheet);
      deleteEmptyRows(sheet);
      formatUnidadesColumn(sheet)
    }
  }
}

function deleteColumnsInsideColumnGroup(sheet) {
  var maxColumn = sheet.getMaxColumns();
  for (let c = maxColumn; c >= 1; c--) {
    const d = sheet.getColumnGroupDepth(c);
    if (d > 0) {
      sheet.deleteColumn(c);
    }
  }
}

function formatUnidadesColumn(sheet) {
  var lastColumn = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var columnNamesToFind = ['UNIDADES', 'PLANTAS', 'NIVEL', 'SUBTIPO TERRAZA', 'TIPO'];

  for (var i = 0; i < columnNamesToFind.length; i++) {
    var columnName = columnNamesToFind[i];
    var columnIndex = headers.indexOf(columnName);
    if (columnIndex !== -1) { break; }
  }

  Logger.log(`columnIndex: ${columnIndex}`)

  // Encontramos una columna del array


  var contieneComputo = false;

  // Verificar si algún encabezado contiene "CÓMPUTO"
  for (var i = 0; i < headers.length; i++) {
    var header = headers[i];
    if (header.indexOf('CÓMPUTO') !== -1) {
      contieneComputo = true;
      break;
    }
  }

  Logger.log(`sheetName: ${sheet.getName()}, columnIndex: ${columnIndex}, contieneComputo: ${contieneComputo}, lastColumn: ${lastColumn}`)

  if (contieneComputo) {

    for (var newColumnIndex = columnIndex + 2; newColumnIndex <= lastColumn; newColumnIndex++) {
      // Verificar si el índice está dentro de los límites de la matriz de encabezados
        var header = headers[newColumnIndex - 1];

        Logger.log(`newColumnIndex: ${newColumnIndex}, header: ${header}, conputoIndex: ${header.indexOf('CÓMPUTO')}`)
        
        // Verificar si el encabezado contiene "CÓMPUTO"
        if (header.indexOf('CÓMPUTO') !== -1) {
        } else {
          var range = sheet.getRange(2, newColumnIndex, lastRow - 1, 1); // Ignorar el encabezado
          range.setNumberFormat('0.00');
        }
    }

  } else {
    var range = sheet.getRange(2, columnIndex + 2, lastRow - 1, lastColumn - columnIndex - 1); // Ignorar el encabezado
    range.setNumberFormat('0.00');
  }
  
}
