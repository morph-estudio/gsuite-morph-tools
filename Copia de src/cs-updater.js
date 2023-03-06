/**
 * Gsuite Morph Tools - CS Updater 1.8.0
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

function morphCSUpdaterOldie(btnID, rowData) {

  // Main variables

  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('LINK') || ss.insertSheet('LINK', 1).setTabColor('#FFFF00');
  const ss_id = ss.getId();
  const ss_name = ss.getName();
  var userMail = Session.getActiveUser().getEmail();
  var dateNow = Utilities.formatDate(new Date(), 'GMT+1', 'dd/MM/yyyy - HH:mm:ss');
  const txtFolderName = `ExpTXT`;
  var file = DriveApp.getFileById(ss_id);
  Logger.log(`FILE: ${file.getName()}`);
  var fileSharing = file.getSharingAccess();
  var functionErrors = [];
  
  // Set Sharing Access for the main file

  if (fileSharing != 'ANYONE_WITH_LINK') file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Variables obtained from user configuration
  var formData = [
    rowData.updatePrefix,
    rowData.prefixAll,
    rowData.updateLinkPage,
    rowData.keepSheetVisibility,
    rowData.updateDebugReport
  ];

  var [updatePrefix, prefixAll, updateLinkPage, keepSheetVisibility, updateDebugReport] = formData;

  let backupOriginalFormula, backupOriginalValues;

  // Backup Checker: eliminar cuando ya todos los cuadros tengan la hoja actualizada a la última versión (con CARPETA BACKUP en A8)
  if (sh.getRange('A7').getValue().toString().toLowerCase().includes('backup')) {
    backupOriginalFormula = sh.getRange('B7').getFormula();
    backupOriginalValues = sh.getRange(8, 2, 2, 1).getValues();
  } else {
    backupOriginalFormula = sh.getRange('B8').getFormula();
    backupOriginalValues = sh.getRange(9, 2, 2, 1).getValues();
  }

  // Main Format and Column A Text

/*
  var textColumnA = [['URL PANEL DE CONTROL'], [null], ['CARPETA PANEL DE CONTROL'], ['ID CARPETA PANEL DE CONTROL'], ['CARPETA CUADRO'], ['ID CARPETA CUADRO'], ['CARPETA EXPORTACIONES'], ['CARPETA BACKUP'], ['ID CARPETA BACKUP'], ['DESCARGAR ARCHIVO XLSX']];
  sh.getRange(1, 1, textColumnA.length, 1).setValues(textColumnA);

  linkPageTemplateFormat(sh); // Cell Format
*/
  sh.getDataRange().clear();
  formatLinkSheetOld(ss);

  // File Folder

  let carpetaCuadroBaseID = file.getParents().next().getId();
  let carpetaCuadroBase = DriveApp.getFolderById(carpetaCuadroBaseID);
  Logger.log(`SHARING: ${fileSharing}, BTN_ID: ${btnID}, FILE FOLDER: ${carpetaCuadroBaseID}`);

  // Control Panel and Control Panel Folder

  var tipodearchivo, controlPanelFileName, controlPanelFileURL, controlPanelFolderID, controlPanelFolder;
  var tiposDeCuadros = ['superficies', 'mediciones', 'exportaciones']

  try {

    if (ss_name.toLowerCase().includes(tiposDeCuadros[0])) {
      tipodearchivo = tiposDeCuadros[0];
    } else if (ss_name.toLowerCase().includes(tiposDeCuadros[1])) {
      tipodearchivo = tiposDeCuadros[1];
    }

    let controlPanelData = getControlPanel(file, tipodearchivo);
    [controlPanelFileName, controlPanelFileID, controlPanelFileURL, controlPanelFolder] = controlPanelData;
    controlPanelFolderID = controlPanelFolder.getId();
    let controlPanelFile = DriveApp.getFileById(controlPanelFileID);
    if(controlPanelFile.getSharingAccess().toString() != 'ANYONE_WITH_LINK') controlPanelFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    functionErrors.push(`No se ha encontrado el panel de control del proyecto, pero no te preocupes, no es un error relevante. Revisa si está en la carpeta correcta y vuelve a actualizar o introduce la URL manualmente.`)
  }

  // ImportRange Permission

  let sectoresID = '1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro';

  importRangeToken(ss_id, carpetaCuadroBaseID);
  importRangeToken(ss_id, sectoresID);

  Utilities.sleep(100);

  // Locate exported TXT folder

  let list = [];

  //if (sh.getDataRange().isBlank() == false) sh.getRange(3, 3, sh.getLastRow() - 2, 2).clear();
  let searchFor = 'title contains "ExpTXT"';
  let expFolder = carpetaCuadroBase.searchFolders(searchFor);
  let expFolderDef = expFolder.next();
  let expFolderID = expFolderDef.getId();
  let expFolderName = expFolderDef.getName();

  // Column B and ImportRange Values

  let controlPanelFolderText = controlPanelFolder == undefined ? '' : `=hyperlink("https://drive.google.com/corp/drive/folders/${controlPanelFolderID}";"${controlPanelFolder}")`;

  var textImportRange = [['=IMPORTRANGE(B1;"Instrucciones!A1")', '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"DB-SI!B2")']];
  var textColumnB = [
    [controlPanelFileURL],
    [null],
    [controlPanelFolderText],
    [controlPanelFolderID],
    [`=hyperlink("https://drive.google.com/corp/drive/folders/${carpetaCuadroBaseID}";"${carpetaCuadroBase}")`],
    [carpetaCuadroBaseID],
    [`=hyperlink("https://drive.google.com/corp/drive/folders/${expFolderID}";"${expFolderName}")`]
  ];
  
  sh.getRange(1, 2, textColumnB.length, 1).setValues(textColumnB);
  sh.getRange(1, 3, 1, 2).setValues(textImportRange);
  sh.getRange('B8').setFormula(backupOriginalFormula);
  sh.getRange(9, 2, 2, 1).setValues(backupOriginalValues);

  // Array List of export files

  keepNewestFilesOfEachNameInAFolder(expFolderDef); // Delete duplicated files in Exports Folder

  let prefijo = [updatePrefix] || ['TXT','MED']; // prefix mask
  let files = expFolderDef.getFiles();
  while (files.hasNext()) {
    file = files.next();
    filename = file.getName();
    if (prefixAll === true) {
      if (filename.includes('.txt')) {
        list.push([file.getName(), file.getId(), file.getName().slice(0, -4).replace('Sheets ', '').toUpperCase()]);
      }
    } else {
      if (prefijo.some(prefix => filename.includes(prefix))) {
        list.push([file.getName(), file.getId(), file.getName().slice(0, -4).replace('Sheets ', '').toUpperCase()]);
      }
    }
  }

  if (list.length < 1) throw new Error(`No se ha encontrado ningun archivo en la carpeta ${txtFolderName} que cumpla los criterios seleccionados.`);

  let result = [['Archivos importados', 'Archivos importados: IDs', 'Hoja'], ...list.filter((r) => r[0].includes('.txt')).sort()];
  let resultCrop = result.map((val) => val.slice(0, -1));
  //sh.getRange(2, 3, getLastDataRow(sh,"C") - 1, 2).clearContent();
  sh.getRange(2, 3, result.length, 2).setValues(resultCrop);
  Logger.log(`Export Folder: ${expFolderDef.getName()}, Number of files to import: ${result.length - 1}`)

  // Export List Format

  sh.getRange(3, 3, list.length, 2).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  //sh.getRange(3, 3, list.length, 1).setFontWeight('bold');
  //sh.getRange(3, 4, list.length, 1).setBackground('#fafafa');
  //sh.getRange(3, 2, 8, 1).setBackground('#fafafa'); // Fondo de la columna B, se colcoca aquí para evitar que se desplace el color a celdas sobrantes

  // Copy TXT data in Sheets

  let txt_file, txt_sharing, tsvUrl, tsvContent, sheetPaste, sheetId;
  let loggerData = [];

  if (sh.getRange('B7').getNote() != '') sh.getRange('B7').clearNote(); // Línea temporal para borrar las notas de la antigua versión (BORRAR EN EL FUTURO)

  sh.getRange('C2').setNote(null).setNote(`Última actualización: ${dateNow} por ${userMail}`); // Last Update Note

  SpreadsheetApp.flush();

  const requests = list.map(([txtFileName, txtFileId, txtFileSheet]) => {
    txt_file = DriveApp.getFileById(txtFileId);
    txt_sharing = txt_file.getSharingAccess();
    if (txt_sharing.toString() != 'ANYONE_WITH_LINK') txt_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    sheetPaste = ss.getSheetByName(txtFileSheet) || ss.insertSheet(`${txtFileSheet}`, 200); sh.activate(); sheetPaste.setTabColor('00FF00');
    if (keepSheetVisibility != true) sheetPaste.hideSheet();
    sheetId = sheetPaste.getSheetId();
    tsvUrl = `https://drive.google.com/uc?id=${txtFileId}&x=.tsv`;
    tsvContent = UrlFetchApp.fetch(tsvUrl).getContentText();

    if (updateDebugReport) {
      let tsvData = Utilities.parseCsv(tsvContent, '\t');
      let obj = { k1_Sheetname: txtFileSheet, k2_dataRows: tsvData.length, k4_dataCols: tsvData[0].length };
      loggerData.push(obj);
    }

    return [
      { updateCells: { 
        range: {
          sheetId: sheetId,
          startRowIndex: 0,
          endRowIndex: sheetPaste.getLastRow(),
          startColumnIndex: 0,
          endColumnIndex: sheetPaste.getLastColumn()
        },
        fields: "userEnteredValue",
        rows: [],
        }},
      { pasteData: { 
        data: tsvContent,
        coordinate: { sheetId },
        delimiter: '\t' 
        }}
      ];
  });

  Sheets.Spreadsheets.batchUpdate({ requests }, ss_id);

  deleteEmptyRows(); removeEmptyColumns();

  if (updateDebugReport) {
    list.map(([txtFileSheet], index) => {
      sheetPaste = ss.getSheetByName(txtFileSheet).getLastRow();
      loggerData[index].k3_copiedRows = sheetPaste;
    });
    Logger.log(loggerData);
  }

  if (functionErrors.length > 0) {
    var ui = SpreadsheetApp.getUi();
    var message = functionErrors.join('\n\n');
    ui.alert('Advertencia', message, ui.ButtonSet.OK);
  }
}

/**
 * morphCSUpdaterDirect(
 * Created for faster successive updates
 */
function morphCSUpdaterDirectOldie(rowData) {

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('LINK');
  let ss_id = ss.getId();
  let file = DriveApp.getFileById(ss_id);
  let userMail = Session.getActiveUser().getEmail();
  let dateNow = Utilities.formatDate(new Date(), 'GMT+1', 'dd/MM/yyyy - HH:mm:ss');
  const txtFolderName = `ExpTXT`;
  ScriptApp.getOAuthToken();
  var functionErrors = [];

  // Variables obtained from user configuration
  var formData = [
    rowData.updateDebugReport
  ];

  var [updateDebugReport] = formData;

  let listValues = sh.getRange("C3:C").getValues();
  let listLength = listValues.filter(String).length;
  if (listLength < 1) throw new Error('No hay ningún archivo en la lista de archivos importados.');
  Logger.log(`FILENAME: ${file.getName()}, Number of files to import: ${listLength}`);
  let listRaw = sh.getRange(3,3,listLength,2).getValues();
  let list = listRaw.map(function(item) {
    let newString = item[0];
    let newItem = [newString, item[1]];
    return newItem;
 });

  let tsvUrl, tsvContent, sheetPaste, sheetId;
  var loggerData = [];
  sh.getRange('C2').setNote(null).setNote(`Última actualización: ${dateNow} por ${userMail}`); // Last Update Note

  const requests = list.map(([txtFileSheet, txtFileId]) => {
    try {
      let sheetName = txtFileSheet.slice(0, -4).replace('Sheets ', '').toUpperCase();
      sheetPaste = ss.getSheetByName(sheetName);
      sh.activate();
      sheetId = sheetPaste.getSheetId();
      tsvUrl = `https://drive.google.com/uc?id=${txtFileId}&x=.tsv`;
      tsvContent = UrlFetchApp.fetch(tsvUrl).getContentText();

      if (updateDebugReport) {
        let tsvData = Utilities.parseCsv(tsvContent, '\t');
        let obj = { k1_Sheetname: sheetName, k2_dataRows: tsvData.length, k4_dataCols: tsvData[0].length };
        loggerData.push(obj);
      }

      return [
        { updateCells: { 
          range: {
            sheetId: sheetId,
            startRowIndex: 0,
            endRowIndex: sheetPaste.getLastRow(),
            startColumnIndex: 0,
            endColumnIndex: sheetPaste.getLastColumn()
          },
          fields: "userEnteredValue",
          rows: [],
          }},
        { pasteData: { 
          data: tsvContent,
          coordinate: { sheetId },
          delimiter: '\t' 
          }}
        ];

    } catch (e) {
      functionErrors.push(`No se ha encontrado el archivo "${txtFileSheet}". Comprueba que exista en la carpeta "${txtFolderName}" o que esté bien escrito el nombre/ID en la lista de archivos importados de la hoja LINK y luego vuelve a actualizar.`)
    }

  });

  SpreadsheetApp.flush();
  Sheets.Spreadsheets.batchUpdate({ requests }, ss_id);

  if (updateDebugReport) {
    list.map(([txtFileSheet], index) => {
      let sheetName = txtFileSheet.slice(0, -4).replace('Sheets ', '').toUpperCase();
      sheetPaste = ss.getSheetByName(sheetName).getLastRow();
      loggerData[index].k3_copiedRows = sheetPaste;
    });
    Logger.log(loggerData);
  }

  if (functionErrors.length > 0) {
    var ui = SpreadsheetApp.getUi();
    var message = functionErrors.join('\n\n');
    ui.alert('Advertencia', message, ui.ButtonSet.OK);
  }
}

/**
 * getControlPanel
 * Search the Control Panel Document in the Folder Structure
 */
function getControlPanelOldie(file, tipodearchivo) {
  let controlPanelData = [];
  let parents;
  var pcMask = 'Panel de control';

  Logger.log(`Se buscará el panel de control asociado a este cuadro de ${tipodearchivo}...`);

  switch (tipodearchivo) {
    case 'superficies':

      parents = file.getParents();
      while (parents.hasNext() && controlPanelData.length == 0) {
        let tempFolder = parents.next();
        let files = DriveApp.getFolderById(tempFolder.getId()).getFilesByType(MimeType.GOOGLE_SHEETS);
        while (files.hasNext()) {
          let fileCP = files.next();
          if (fileCP.getName().toLowerCase().includes(pcMask.toLowerCase())) {
            controlPanelData.push(fileCP.getName(), fileCP.getId(), fileCP.getUrl(), tempFolder);
            break;
          }
        }
        parents = tempFolder.getParents();
      }
      break;
    case 'mediciones':

      parents = file.getParents();
      while (parents.hasNext() && controlPanelData.length == 0) {

        let tempFolder = parents.next();
        let tempFolderFolders = tempFolder.getFolders();

        Logger.log(tempFolder.getName());

        while (tempFolderFolders.hasNext()) {

          /**/
          let grandFolder = tempFolderFolders.next();
          let folderNames = ["Proyecto Base", "Arquitectura"];

          Logger.log(grandFolder.getName())

          if (folderNames.some(word => grandFolder.getName().includes(word))) {

            Logger.log('he encontrado la puta carpeta' + grandFolder.getName())
            
            let tempFolderFoldersLv2 = grandFolder.getFolders();

            while (tempFolderFoldersLv2.hasNext()) {
              let parentFolder = tempFolderFoldersLv2.next();
              if (parentFolder.getName().includes('Doc Escrita')) {
                Logger.log('he entrado en' + parentFolder.getName())
                let parentFolderID = parentFolder.getId();
                let folderFiles = DriveApp.getFolderById(parentFolderID).getFilesByType(MimeType.GOOGLE_SHEETS);
                
                while (folderFiles.hasNext()) {
                  let fileCP = folderFiles.next();
                  if (fileCP.getName().toLowerCase().includes(pcMask.toLowerCase())) {
                    Logger.log('panel de control name' + fileCP.getName())
                    controlPanelData.push(fileCP.getName(), fileCP.getId(), fileCP.getUrl(), parentFolder);
                    break
                  }
                }
              }
            }
          }
        }
        parents = tempFolder.getParents();
      }
      break;
  }
  Logger.log(`CONTROL_PANEL_INFO: ${controlPanelData}`);
  return controlPanelData;
}
