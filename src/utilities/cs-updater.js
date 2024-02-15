/**
 * Gsuite Morph Tools - CS Updater
 * Developed by alsanchezromero
 *
 * Morph Estudio, 2023
 */

function morphCSUpdater(btnID, rowData) {

  // Main variables

  const linkPageName = `LINK`;
  const txtFolderName = `ExpTXT`;
  const linkPageTabColor = '#FFFF00';
  const exportedTabColor = '#00ff00';
  
  ScriptApp.getOAuthToken();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(linkPageName) || ss.insertSheet(linkPageName, 1).setTabColor(linkPageTabColor);
  const ss_id = ss.getId();
  const ss_name = ss.getName();
  const userMail = Session.getActiveUser().getEmail();
  const dateNow = new Date().toLocaleString('es-ES', {timeZone: 'Europe/Madrid', hour12: false});
  var file = DriveApp.getFileById(ss_id);
  const fileSharing = file.getSharingAccess();
    const functionErrors = [];

  Logger.log(`FILENAME: ${file.getName()}, URL: ${file.getUrl()}`);

  ScriptApp.getOAuthToken();
  if (fileSharing != 'ANYONE_WITH_LINK') file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Variables obtained from user configuration

  const {
    updatePrefix,
    updateLinkPage,
    keepSheetVisibility,
    updateDebugReport,
    updateHistorico,
    updateBasic,
    updateAI,
    updateAC,
    quotaControl,
  } = rowData;

  /* Activar cuando el histórico esté disponible
  if (updateHistorico) historicoDeSuperficies();
  */

  Logger.log(`updateLinkPage: ${updateLinkPage}`)

  var aiColumnLetter;
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var acIndex = headers.findIndex(header => header.trim().toLowerCase() === 'archivos importados (ai)');
  if (acIndex >= 0) aiColumnLetter = numToCol(acIndex + 1);

  if (updateLinkPage) {

    // Basic Page Format

    formatLinkSheet(updateBasic, updateAI, updateAC, ss);

    // File Folder

    let cuadroFolder = file.getParents().next();
    let cuadroFolderID = cuadroFolder.getId();

    Logger.log(`SHARING: ${fileSharing}, BTN_ID: ${btnID}, FILE FOLDER: ${cuadroFolderID}`);

    // Locate exported TXT folder

    if (updateBasic === true || updateAI  === true) {

      var exportFolder, exportFolder, exportFolderID, exportFolderName;

      let searchFor = 'title contains "ExpTXT"';
      exportFolder = cuadroFolder.searchFolders(searchFor);

      if (!exportFolder.hasNext()) { // La carpeta no fue encontrada
        var response = Browser.msgBox("Atención", "No se ha encontrado la carpeta para los archivos de exportación. ¿Deseas crearla automáticamente?", Browser.Buttons.OK_CANCEL);
        if (response == "cancel") {
          throw new Error(`No se ha podido encontrar la carpeta con los archivos de exportación.`);
        }

        exportFolderName = `${ss.getName().substring(0, 6)}-A-CS-${searchFor.substring(16, searchFor.length - 1)}`; // Obtener el nombre de la carpeta desde la cadena de búsqueda
        exportFolder = cuadroFolder.createFolder(exportFolderName);
        exportFolder = exportFolder.next()

        throw new Error(`Se ha creado la carpeta ExpTXT, introduce los archivos en la carpeta y vuelve a actualizar.`);
        
      } else { // La carpeta fue encontrada

        exportFolder = exportFolder.next();

        let prefijo = (updatePrefix.trim() !== "") ? true : false;

        if (prefijo) {
          let subCarpetas = exportFolder.getFolders();
          while (subCarpetas.hasNext()) {
            let subCarpeta = subCarpetas.next();
            if (subCarpeta.getName().includes(updatePrefix)) {
                exportFolder = subCarpeta;
                break;
            }
          }
        }

        exportFolderID = exportFolder.getId();
        exportFolderName = exportFolder.getName();
      }
    }

    // Control Panel & Control Panel Folder

    if (updateBasic) {

      try {

        let tiposDeCuadros = ['superficies', 'mediciones', 'exportaciones'];
        if (ss_name.toLowerCase().includes(tiposDeCuadros[0])) {
          var tipodearchivo = tiposDeCuadros[0];
        } else if (ss_name.toLowerCase().includes(tiposDeCuadros[1])) {
          var tipodearchivo = tiposDeCuadros[1];
        }

        let controlPanelData = getControlPanel(file, tipodearchivo);
        [ controlPanelFileName, controlPanelFileID, controlPanelFileURL, controlPanelFolder ] = controlPanelData;
        controlPanelFolderID = controlPanelFolder.getId();
        const controlPanelFile = DriveApp.getFileById(controlPanelFileID);
        if(controlPanelFile.getSharingAccess().toString() != 'ANYONE_WITH_LINK') controlPanelFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      } catch (err) {
        functionErrors.push(`No se ha encontrado el panel de control del proyecto, pero no te preocupes, no es un error relevante. Revisa si está en la carpeta correcta y vuelve a actualizar o introduce la URL manualmente.`)
      }

      // ImportRange Permission for external files

      let sectoresID = '1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro';
      importRangeToken(ss_id, cuadroFolderID);
      importRangeToken(ss_id, sectoresID);

      Utilities.sleep(100);

      // Basic Data Column

      var controlPanelFolderLink = controlPanelFolder == undefined ? '' : `=hyperlink("https://drive.google.com/corp/drive/folders/${controlPanelFolderID}";"${controlPanelFolder}")`;
      var cuadroFolderLink = `=hyperlink("https://drive.google.com/corp/drive/folders/${cuadroFolderID}";"${cuadroFolder}")`;
      var exportFolderLink = `=hyperlink("https://drive.google.com/corp/drive/folders/${exportFolderID}";"${exportFolderName}")`

      var basicDataColumn = linkColumnA();
      var basicDataColumnRange = sh.getRange(1, 1, sh.getLastRow(), 1).getValues(); // Obtenemos solo la primera columna
      var finalDataColumn = [];

      for (var i = 0; i < basicDataColumnRange.length; i++) {
        var key = basicDataColumnRange[i][0];

        if (basicDataColumn.hasOwnProperty(key)) {
          var value = basicDataColumn[key];
          finalDataColumn.push([value])
        }
      }
      
      sh.getRange(1, 2, finalDataColumn.length, 1).setValues(finalDataColumn);
    }

    function linkColumnA() {
      return {
        "URL PANEL DE CONTROL": controlPanelFileURL,
        "CARPETA PANEL DE CONTROL": controlPanelFolderLink,
        "ID CARPETA PANEL DE CONTROL": controlPanelFolderID,
        "CARPETA CUADRO": cuadroFolderLink,
        "ID CARPETA CUADRO": cuadroFolderID,
        "CARPETA EXPORTACIONES": exportFolderLink,
        "CARPETA CONGELADOS": null,
        "ÚLTIMO ARCHIVO CONGELADO": null,
        "DESCARGAR ARCHIVO XLSX": null,
      };
    }
  }

  if (updateAI || updateLinkPage) {

    var exportedFilesList = []; 

    let txt_file, txt_sharing, tsvUrl, tsvContent, sheetPaste, sheetId;
    let loggerData = [];

    if (updateLinkPage) {

      var changeTabColors = []
      keepNewestFilesOfEachNameInAFolder(exportFolder); // Delete duplicated files in Exports Folder

      let files = exportFolder.getFiles();
      while (files.hasNext()) {
        file = files.next();
        filename = file.getName();
        exportedFilesList.push([file.getName(), file.getId(), file.getName().toUpperCase().replace(/\.[^/.]+$/, "").replace('SHEETS', '').replace(/\s+/g, ' ').trim()]);
      }

      if (exportedFilesList.length < 1) throw new Error(`No se ha encontrado ningun archivo en la carpeta ${txtFolderName} que cumpla los criterios seleccionados.`);

      var allowedExtensions = ['.txt', '.tsv', '.csv'];
      let result = exportedFilesList.filter((r) => allowedExtensions.some(ext => r[0].includes(ext))).sort();
      let resultCrop = result.map((val) => val.slice(0, -1));
      sh.getRange(2, 4, result.length, 2).setValues(resultCrop);

      Logger.log(`Export Folder: ${exportFolder.getName()}, Number of files to import: ${result.length - 1}`)

      // Exported List Format in Link Page

      sh.getRange(2, 4, exportedFilesList.length, 2).setBorder(false, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

      SpreadsheetApp.flush();
    }

    if (!updateLinkPage) {
      var listValues = sh.getRange(`${aiColumnLetter}2:${aiColumnLetter}`).getValues();

      let listLength = listValues.filter(String).length;
      if (listLength < 1) throw new Error('No hay ningún archivo en la lista de archivos importados.');
      Logger.log(`UpdateType: Direct, Number of files to import: ${listLength}`);

      let rawList = sh.getRange(2, acIndex + 1, listLength, 2).getValues();

      for (let i = 0; i < rawList.length; i++) {
        let item = rawList[i];
        exportedFilesList.push([item[0], item[1], item[0].toUpperCase().replace(/\.[^/.]+$/, "").replace('SHEETS', '').replace(/\s+/g, ' ').trim()]);
      }
    }

    // Copy TXT data in Spreadsheet
    
    const requests = exportedFilesList.map(([txtFileName, txtFileId, txtFileSheet]) => {
      txt_file = DriveApp.getFileById(txtFileId);
      txt_sharing = txt_file.getSharingAccess();
      if (txt_sharing.toString() != 'ANYONE_WITH_LINK') txt_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

      sheetPaste = ss.getSheetByName(txtFileSheet) || ss.insertSheet(`${txtFileSheet}`, 200); //sh.activate();

      if (updateLinkPage) {
        var currentTabColor = sheetPaste.getTabColorObject();
        if (currentTabColor != exportedTabColor) changeTabColors.push(sheetPaste);
        if (keepSheetVisibility != true) sheetPaste.hideSheet();
      }

      if(!quotaControl) {

        try {
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

        } catch (e) {
          functionErrors.push(`No se ha encontrado el archivo "${txtFileSheet}". Comprueba que exista en la carpeta "${txtFolderName}" o que esté bien escrito el nombre/ID en la lista de archivos importados de la hoja LINK y luego vuelve a actualizar.`)
        }

      } else {

        let archivo = DriveApp.getFileById(txtFileId);
        let contenido = archivo.getBlob().getDataAsString();
        let filas = contenido.split('\n');
        let tsvContent = [];

        for (var i = 0; i < filas.length; i++) {
          var fila = filas[i].split('\t');
          tsvContent.push(fila);
        }

        try {
          let targetRange = sheetPaste.getRange(1, 1, tsvContent.length, tsvContent[0].length);
          targetRange.clearContent();
          targetRange.setValues(tsvContent);
        } catch (error) {
        }
      }
    });

    if(!quotaControl) { Sheets.Spreadsheets.batchUpdate({ requests }, ss_id); }
  }

  // if (updateAC) getConectedSheetList(0, 6, 'LINK'); // Sin usar hasta que no se programe la lista automática de hojas conectadas

  if (updateLinkPage) {
    if (updateAI) {
      for (var i = 0; i < changeTabColors.length; i++) {
        changeTabColors[i].setTabColor(exportedTabColor);
      }
    }
    deleteEmptyRows(sh);
    removeEmptyColumns(sh);
  }

  if (updateDebugReport) {
    exportedFilesList.map(([txtFileSheet], index) => {
      sheetPaste = ss.getSheetByName(txtFileSheet).getLastRow();
      loggerData[index].k3_copiedRows = sheetPaste;
    });
    Logger.log(loggerData);
  }

  sh.getRange(`${aiColumnLetter}1`).setNote(null).setNote(`Última actualización: ${dateNow} por ${userMail}`); // Last Update Note

  if (functionErrors.length > 0) {
    let ui = SpreadsheetApp.getUi();
    let message = functionErrors.join('\n\n');
    ui.alert('Advertencia', message, ui.ButtonSet.OK);
  }
}





















/**
 * getControlPanel
 * Search the Control Panel Document in the Folder Structure
 */
function getControlPanel(file, tipodearchivo) {

  const controlPanelData = [];
  var parents;
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

