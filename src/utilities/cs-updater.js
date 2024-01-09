/**
 * Gsuite Morph Tools - CS Updater
 * Developed by alsanchezromero
 *
 * Morph Estudio, 2023
 */

function morphCSUpdater(btnID, rowData) {

  // Main variables

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('LINK') || ss.insertSheet('LINK', 1).setTabColor('#FFFF00');
  var ss_id = ss.getId();
  var ss_name = ss.getName();
  var userMail = Session.getActiveUser().getEmail();
  var dateNow = Utilities.formatDate(new Date(), 'GMT+3', 'dd/MM/yyyy - HH:mm:ss');
  const txtFolderName = `ExpTXT`;
  var file = DriveApp.getFileById(ss_id);
  Logger.log(`FILENAME: ${file.getName()}, URL: ${file.getUrl()}`);
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
    rowData.updateDebugReport,
    rowData.updateHistorico,
    rowData.updateBasic,
    rowData.updateAI,
    rowData.updateAC
  ];

  var [updatePrefix, prefixAll, updateLinkPage, keepSheetVisibility, updateDebugReport, updateHistorico, updateBasic, updateAI, updateAC] = formData;

  if (updateHistorico) historicoDeSuperficies();

  // Main Format and Column A Text

  formatLinkSheet(updateBasic, updateAI, updateAC, ss);

  // File Folder
  let carpetaCuadroBaseID = file.getParents().next().getId();
  let carpetaCuadroBase = DriveApp.getFolderById(carpetaCuadroBaseID);
  Logger.log(`SHARING: ${fileSharing}, BTN_ID: ${btnID}, FILE FOLDER: ${carpetaCuadroBaseID}`);

  // Locate exported TXT folder

  if (updateBasic === true || updateAI  === true) {

    var list = []; var expFolder; var expFolderDef; var expFolderID; var expFolderName;

    var searchFor = 'title contains "ExpTXT"';
    expFolder = carpetaCuadroBase.searchFolders(searchFor);

    if (!expFolder.hasNext()) { // La carpeta no fue encontrada
      var response = Browser.msgBox("Atención", "No se ha encontrado la carpeta para los archivos de exportación. ¿Deseas crearla automáticamente?", Browser.Buttons.OK_CANCEL);
      if (response == "cancel") {
        throw new Error(`No se ha podido encontrar la carpeta con los archivos de exportación.`);
      }

      // Crear la carpeta
      expFolderName = `${ss.getName().substring(0, 6)}-A-CS-${searchFor.substring(16, searchFor.length - 1)}`; // Obtener el nombre de la carpeta desde la cadena de búsqueda
      expFolder = carpetaCuadroBase.createFolder(expFolderName);
      expFolderDef = expFolder.next()

      throw new Error(`Se ha creado la carpeta ExpTXT, introduce los archivos en la carpeta y vuelve a actualizar.`);

      /*
      var expFolderID = newFolder.getId();
      var expFolderDef = DriveApp.getFolderById(expFolderID);
      expFolder = [expFolderDef]; // Actualizar la variable expFolder para usar la carpeta recién creada
      */
      
    } else { // La carpeta fue encontrada
      expFolderDef = expFolder.next();
      expFolderID = expFolderDef.getId();
      expFolderName = expFolderDef.getName();
    }

    Logger.log(`EXPFOLDER: ${expFolderName}`);

    try {    } catch (error) {
    }

  }

  if (updateBasic) {

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

    // Column B and ImportRange Values

      let controlPanelFolderText = controlPanelFolder == undefined ? '' : `=hyperlink("https://drive.google.com/corp/drive/folders/${controlPanelFolderID}";"${controlPanelFolder}")`;

      var textImportRange = [['=IMPORTRANGE(B1;"Instrucciones!A1")', '=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"DB-SI!B2")']];

      var textColumnB = [
        [controlPanelFileURL],
        [controlPanelFolderText],
        [controlPanelFolderID],
        [`=hyperlink("https://drive.google.com/corp/drive/folders/${carpetaCuadroBaseID}";"${carpetaCuadroBase}")`],
        [carpetaCuadroBaseID],
        [`=hyperlink("https://drive.google.com/corp/drive/folders/${expFolderID}";"${expFolderName}")`]
      ];
      
      sh.getRange(1, 2, textColumnB.length, 1).setValues(textColumnB);

    try {    } catch (error) {
    }
  }

  // Array List of export files

  if (updateAI) {

      keepNewestFilesOfEachNameInAFolder(expFolderDef); // Delete duplicated files in Exports Folder

      let prefijo = (updatePrefix.trim() !== "") ? [updatePrefix] : ['TXT', 'MED', 'CSV', 'SUP']; // prefix mask
      let files = expFolderDef.getFiles();
      while (files.hasNext()) {
        file = files.next();
        filename = file.getName();

        if (prefixAll === true) {
          if (filename.includes('.txt')) {
            list.push([file.getName(), file.getId(), file.getName().toUpperCase().slice(0, -4).replace('SHEETS ', '').trim()]);
          }
        } else {
          let filenamePrefix = filename.split(' ')[0]; // Obtenemos el prefijo antes del primer espacio
          if (prefijo.some(prefix => filenamePrefix.includes(prefix))) {
            list.push([file.getName(), file.getId(), file.getName().toUpperCase().slice(0, -4).replace('SHEETS ', '').trim()]);
          }
        }
      }

      if (list.length < 1) throw new Error(`No se ha encontrado ningun archivo en la carpeta ${txtFolderName} que cumpla los criterios seleccionados.`);

      var allowedExtensions = ['.txt', '.tsv', '.csv'];
      let result = list.filter((r) => allowedExtensions.some(ext => r[0].includes(ext))).sort();
      let resultCrop = result.map((val) => val.slice(0, -1));
      sh.getRange(2, 4, result.length, 2).setValues(resultCrop);
      Logger.log(`Export Folder: ${expFolderDef.getName()}, Number of files to import: ${result.length - 1}`)

      // Export List Format

      sh.getRange(2, 4, list.length, 2).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

      // Copy TXT data in Sheets

      let txt_file, txt_sharing, tsvUrl, tsvContent, sheetPaste, sheetId;
      let loggerData = [];

      // if (sh.getRange('B8').getNote() != '') sh.getRange('B7').clearNote(); // Línea temporal para borrar las notas de la antigua versión (BORRAR EN EL FUTURO)

      sh.getRange('D1').setNote(null).setNote(`Última actualización: ${dateNow} por ${userMail}`); // Last Update Note

      SpreadsheetApp.flush();

      var changeTabColors = []; var tabColor = '#00ff00';

      const requests = list.map(([txtFileName, txtFileId, txtFileSheet]) => {
        txt_file = DriveApp.getFileById(txtFileId);
        txt_sharing = txt_file.getSharingAccess();
        if (txt_sharing.toString() != 'ANYONE_WITH_LINK') txt_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        sheetPaste = ss.getSheetByName(txtFileSheet) || ss.insertSheet(`${txtFileSheet}`, 200); sh.activate();

        var currentTabColor = sheetPaste.getTabColorObject(); Logger.log(`Tab color: ${currentTabColor}`)
        if (currentTabColor != tabColor) changeTabColors.push(sheetPaste);

        if (keepSheetVisibility != true) sheetPaste.hideSheet();
        sheetId = sheetPaste.getSheetId();
        tsvUrl = `https://drive.google.com/uc?id=${txtFileId}&x=.tsv`;
        tsvContent = UrlFetchApp.fetch(tsvUrl).getContentText();

        let delimiter = '\t';
        if (txtFileName.endsWith('.csv')) {
          delimiter = ',';
        } else if (txtFileName.endsWith('.asv')) {
          delimiter = '*';
        }

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
            delimiter: delimiter 
            }}
          ];
      });

      Sheets.Spreadsheets.batchUpdate({ requests }, ss_id);

    try {    } catch (error) {
    }

  }

  if (updateAC) getConectedSheetList(0, 6, 'LINK');

  if (updateAI) {
    for (var i = 0; i < changeTabColors.length; i++) {
      changeTabColors[i].setTabColor(tabColor);
    }
  }

  deleteEmptyRows(sh); removeEmptyColumns(sh);

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
function morphCSUpdaterDirect(rowData) {

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('LINK');
  let ss_id = ss.getId();
  let file = DriveApp.getFileById(ss_id);
  Logger.log(`FILENAME: ${file.getName()}, URL: ${file.getUrl()}`);
  let userMail = Session.getActiveUser().getEmail();
  let dateNow = Utilities.formatDate(new Date(), 'GMT+1', 'dd/MM/yyyy - HH:mm:ss');
  const txtFolderName = `ExpTXT`;
  ScriptApp.getOAuthToken();
  var functionErrors = [];

  // Variables obtained from user configuration
  var formData = [
    rowData.updateDebugReport,
    rowData.updateHistorico
  ];

  var [updateDebugReport, updateHistorico] = formData;

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var acIndex = headers.findIndex(header => header.trim().toLowerCase() === 'archivos importados (ai)');
  
  if (acIndex >= 0) {
    var cLetter = numToCol(acIndex + 1); Logger.log(cLetter);
    var listValues = sh.getRange(`${cLetter}2:${cLetter}`).getValues();
  }

  let listLength = listValues.filter(String).length;
  if (listLength < 1) throw new Error('No hay ningún archivo en la lista de archivos importados.');
  Logger.log(`Number of files to import: ${listLength}`);

  if (updateHistorico) historicoDeSuperficies();

  let listRaw = sh.getRange(2, acIndex + 1, listLength, 2).getValues();

  let list = listRaw.map(function(item) {
    let newString = item[0];
    let newItem = [newString, item[1]];
    return newItem;
 });

  let tsvUrl, tsvContent, sheetPaste, sheetId;
  var loggerData = [];
  sh.getRange(`${cLetter}1`).setNote(null).setNote(`Última actualización: ${dateNow} por ${userMail}`); // Last Update Note

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
function getControlPanel(file, tipodearchivo) {
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

/**
 * getConectedSheetList
 * Construct a list of sheets connected with ImportRange Formulas
 */
 function getConectedSheetList(rowShift, colShift, sheetName) {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var importedSheets = [];

  // Iterar sobre todas las hojas del documento

  for (var i = 0; i < sheets.length; i++) {
    let sheet = sheets[i];
    let range = sheet.getRange("A1");
    let formula = range.getFormula();
    
    // Si la celda A1 contiene una fórmula IMPORTRANGE, guarda la información en un objeto

    if (formula.indexOf("IMPORTRANGE") !== -1) {
      Logger.log(`formula: ${formula}`)
      let sheetName = sheet.getName();
      let targetGID = sheet.getSheetId();
      let originFirstCellA1 = getFirstCellA1Notation(sheet); Logger.log(`originFirstCellA1: ${originFirstCellA1}`);

      let formulaMatch, targetSpreadsheetURL, originSheetRange, arrayValues, originSheetName, referencedSheetRange

      if (formula.toString().includes('https://')) {

        formulaMatch = formula.match(/\bIMPORTRANGE\("([^"]+)";"([^"]+)"\)/i);
        targetSpreadsheetURL = formulaMatch[1].trim();
        originSheetRange = formulaMatch[2].trim();

        arrayValues = originSheetRange.split('!');
        originSheetName = arrayValues[0];
        referencedSheetRange = arrayValues[1];

      } else {

        formulaMatch = formula.match(/\bIMPORTRANGE\(([^;)]+);[\"\']?([^\"\');]+)/i);
        let firstArgument = formulaMatch[1].trim();
        originSheetRange = formulaMatch[2].trim();

        arrayValues = firstArgument.split('!');
        originSheetName = arrayValues[0];
        let originSheet = ss.getSheetByName(originSheetName)
        targetSpreadsheetURL = originSheet.getRange(arrayValues[1]).getValue();

        arrayValues = originSheetRange.split('!');
        originSheetName = arrayValues[0];
        referencedSheetRange = arrayValues[1];

      }

      Logger.log(`targetSpreadsheetURL: ${targetSpreadsheetURL}`); Logger.log(`originSheetName: ${originSheetName}`);

      let originSpreadsheet = SpreadsheetApp.openById(getIdFromUrl(targetSpreadsheetURL.toString()));
      let originSpreadsheetName = originSpreadsheet.getName();
      let originGID = originSpreadsheet.getSheetByName(originSheetName).getSheetId();

      importedSheets.push({
        "originSpreadsheetName": originSpreadsheetName,
        "originSheetName": originSheetName,
        "originGID": originGID,
        "targetSheetName": sheetName,
        "targetGID": targetGID,
        "targetSpreadsheetURL": targetSpreadsheetURL,
        "originFirstCellA1": originFirstCellA1
      });
    }
  }

  Logger.log(`importedSheets: ${importedSheets}`);
  
  // Construir la lista de hojas conectadas

  var list = [];
  for (var i = 0; i < importedSheets.length; i++) {
    let importedSheet = importedSheets[i];
    targetSpreadsheetURL = importedSheet["targetSpreadsheetURL"]; Logger.log(`originSpreadsheetURLBeforeOpen: ${targetSpreadsheetURL}`);
    var importedSpreadsheet = SpreadsheetApp.openById(getIdFromUrl(importedSheet["targetSpreadsheetURL"]));
    var importedSpreadsheetName = importedSpreadsheet.getName();

    var row = [
      importedSpreadsheetName,
      targetSpreadsheetURL,
      `=HYPERLINK("${targetSpreadsheetURL}#gid=${importedSheet["originGID"]}";"${importedSheet["originSheetName"]}")`,
      `=HYPERLINK("#gid=${importedSheet["targetGID"]}";"${importedSheet["targetSheetName"]}")`,
      importedSheet["originFirstCellA1"],
    ];

    list.push(row);
  }
  
  // Ordenar la lista alfabéticamente por nombre

  list.sort(function(a, b) {
    var nameA = a[0].toUpperCase();
    var nameB = b[0].toUpperCase();
    if (nameA < nameB) {
      return -1;
    }
    if (nameA > nameB) {
      return 1;
    }
    return 0;
  });

  
  // Pegar la lista en la hoja "LINK"

  var sh_link = ss.getSheetByName(sheetName) || ss.getSheetByName("LINKs");

  let [textHeadersAC, colArray] = connectedListFormat(sh_link, rowShift, colShift);
  list = [textHeadersAC.slice(0, -2), ...list];
  let connectedSheetsListColumns = textHeadersAC.length;

  let lastRow = sh_link.getLastRow();
  let emptyChecker = checkIfSheetIsEmpty(sh_link); Logger.log(emptyChecker)

  if (emptyChecker != true) {
    sh_link.getRange(2 + rowShift, 1 + colShift, lastRow - rowShift, connectedSheetsListColumns).clearContent();
    sh_link.getRange(1 + rowShift, 1 + colShift, lastRow - rowShift, connectedSheetsListColumns)
      .setBorder(false, false, false, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor('#607D8B');
  }

  let listRange = sh_link.getRange(1 + rowShift, 1 + colShift, list.length, 5);

  listRange.setValues(list);
  deleteDuplicatedListObjects(sh_link, listRange, rowShift, colShift);

  if(importedSheets.length > 0) {
    let n;
    for (var i = 0; i < importedSheets.length; i++) {
      n = i + 2;
      sh_link.getRange(rowShift + n, 6 + colShift)
      .setFormula(`=IF(${numToCol(1 + colShift)}${n}<>"";IMPORTRANGE(CHAR(34)&${colArray[1]}${n}&CHAR(34);CHAR(39)&${colArray[2]}${n}&"'!"&${colArray[4]}${n});)`);
      sh_link.getRange(rowShift + n, 7 + colShift)
      .setFormula(`=IF(${colArray[0]}${rowShift + n}:${colArray[0]}<>"";IF(ISERROR(${colArray[5]}${rowShift + n}:${colArray[5]});"🟥";"🟩");"")`);
    }
  }

  // List Format

  sh_link.getRange(1 + rowShift, 1 + colShift, list.length, connectedSheetsListColumns).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
  // Limpiar filas y columnas sobrantes

  let deleteRowIndex = getLastDataRowIndex(sh_link);
  let maxRows = sh_link.getMaxRows();
  if (maxRows > deleteRowIndex) sh_link.deleteRows(deleteRowIndex + 1, maxRows - deleteRowIndex);
  removeEmptyColumns(sh_link);
}

/**
 * deleteDuplicatedListObjects
 * Borrar elementos duplicados en la lista de archivos conectados
 */
function deleteDuplicatedListObjects(sh, dataRange, rowShift, colShift) {
  var data = dataRange.getValues();
  var uniqueValues = {}; // Objeto para almacenar los valores únicos de "Archivos conectados"
  
  // Recorrer el array de datos
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var value = row[1];
    
    // Si el valor no está en el objeto, almacenarlo
    if (!uniqueValues[value]) {
      uniqueValues[value] = true;
    } else {
      // Si el valor ya está en el objeto, modificar las celdas
      sh.getRange(i + rowShift + 1, 1 + colShift, 1, 1).clearContent(); // Borrar la columna "Archivos duplicados"
      sh.getRange(i + rowShift + 1, 2 + colShift, 1, 1).setFontColor('#EFEFEF'); // Cambiar el color de la columna "Acción"
    }
  }
}
