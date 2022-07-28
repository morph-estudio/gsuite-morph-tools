function doGet() {
  return HtmlService.createHtmlOutputFromFile('html/index');
}

/*
 * Gsuite Morph Tools - CS Updater 1.3
 * Developed by alsanchezromero
 * Created on Mon Jul 25 2022
 *
 * Copyright (c) 2022 Morph Estudio
*/

function actualizarCuadro() {
  const ss = SpreadsheetApp.getActive();
  let ws = ss.getSheetByName('ACTUALIZAR') || ss.insertSheet('ACTUALIZAR', 1);
  let ss_id = ss.getId();

  /* TESTvar ss = SpreadsheetApp.openById('1v5f3X1ShmVCGdT6NdWvmJHcfeP01ptuwHfT1iqM6UQI');
  var ws = ss.getSheetByName('ACTUALIZAR')
  let ss_id = '1v5f3X1ShmVCGdT6NdWvmJHcfeP01ptuwHfT1iqM6UQI';*/

  let file = DriveApp.getFileById(ss_id);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  ws.clear().clearFormats();
  sheetFormatter(ws); // Formato de celdas

  // Carpeta Cuadro de Superficies

  let parents = file.getParents();
  let carpetaBaseID = parents.next().getId();
  let carpetaBase = DriveApp.getFolderById(carpetaBaseID);

  // Panel de control

  let fldrA = [];
  parents = file.getParents();
  while (parents.hasNext()) {
    let f = parents.next();
    let f_id = f.getId();
    fldrA.push(f_id);
    parents = f.getParents();
  };

  let filA = [];
  let pcMask = 'Panel de control';
  for (let i = 0; i < fldrA.length; i++) {
    let files = DriveApp.getFolderById(fldrA[i]).getFilesByType(MimeType.GOOGLE_SHEETS);
    while (files.hasNext()) {
      let filePC = files.next();
      if (filePC.getName().includes(pcMask)) {
        filA.push([filePC.getName()], [filePC.getId()], [filePC.getUrl()], [fldrA[i]]);
      }
    }
  }

  let [filePanelName, filePanelId, filePanelUrl, folderPanelcId] = filA;
  let panelControl = DriveApp.getFileById(filePanelId);
  let folderPanelcName = DriveApp.getFolderById(folderPanelcId);

  panelControl.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Set Template Values

  tplText(ws);

  ws.getRange('B1').setValue(filePanelUrl);
  ws.getRange('B3').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${carpetaBaseID}";"${carpetaBase}")`).setFontColor('#4A86E8');
  ws.getRange('B4').setValue(carpetaBaseID);
  ws.getRange('B5').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${folderPanelcId}";"${folderPanelcName}")`).setFontColor('#4A86E8');
  ws.getRange('B6').setValue(folderPanelcId);

  /**/
  /* Versión obsoleta del buscador
  var searchFor ='title contains "Panel"';
  var names =[];
  var filePanelIds=[];
  var filesPanel = folderPanelcName.searchFiles(searchFor);
    while (filesPanel.hasNext()) {
  var filePanel = filesPanel.next();
  var filePanelId = filePanel.getId();// To get FileId of the file
  filePanelIds.push(filePanelId);
  var filePanelname = filePanel.getName();
  var filePanelUrl = filePanel.getUrl();
  names.push(filePanelname);
  }
  */

  // ImportRange Permission

  let sectoresID = '1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro';

  importRangeToken(ss_id, carpetaBaseID);
  importRangeToken(ss_id, sectoresID);

  ws.getRange('C1').setValue('=IMPORTRANGE(B1;"Instrucciones!A1")');
  ws.getRange('D1').setValue('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"DB-SI!B2")');

  Utilities.sleep(250);

  // Localizar archivos TXT exportados

  let rangeClear = ws.getRange(3, 3, 6, 2);
  rangeClear.clearContent().clearFormat();

  let searchFor = 'title contains "Exportaciones"';
  let names =[];
  let expFolderIds=[];
  let expFolder = carpetaBase.searchFolders(searchFor);
  let expFolderDef; let expFolderId; let expFolderName;
  while (expFolder.hasNext()) {
    expFolderDef = expFolder.next();
    expFolderId = expFolderDef.getId();
    expFolderIds.push(expFolderId);
    expFolderName = expFolderDef.getName();
    names.push(expFolderName);
  }

  let sufix = 'TXT'; // mask
  let list = [];
  let files = expFolderDef.getFiles();
  while (files.hasNext()) {
    file = files.next();
    list.push([file.getName(), file.getId(), file.getName().slice(0, -4).replace('Sheets ', '').toUpperCase()]);
  }

  let result = [['Archivos exportados', 'IDs', 'Hoja'], ...list.filter((r) => r[0].includes(sufix)).sort()];
  let resultCrop = result.map((val) => val.slice(0, -1));
  ws.getRange(2, 3, result.length, 2).setValues(resultCrop);

  // Filelist Styling

  ws.getRange(3, 3, list.length, 1).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#B7B7B7').setFontWeight('bold');
  ws.getRange(3, 4, list.length, 1).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#B7B7B7').setBackground('#F3F3F3');
  ws.getRange(3, 3, list.length, 2).setFontSize(13).setFontFamily('Montserrat').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle');

  Utilities.sleep(100);

  // Copy data in Sheets

  for (const [txtFileName, txtFileId, txtFileSheet] of list) {
    let txt_file = DriveApp.getFileById(txtFileId);
    txt_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    let tsvUrl = `https://drive.google.com/uc?id=${txtFileId}&x=.tsv`;
    let tsvContent = UrlFetchApp.fetch(tsvUrl, {muteHttpExceptions: true }).getContentText();
    let tsvData = Utilities.parseCsv(tsvContent, '\t');
    let sheetPaste = ss.getSheetByName(`${txtFileSheet}`) || ss.insertSheet(`${txtFileSheet}`, 100);
    ws.activate();
    sheetPaste.setTabColor('F1C232').hideSheet();
    sheetPaste.clear();
    sheetPaste.getRange(1, 1, tsvData.length, tsvData[0].length).setValues(tsvData);
  };
}

function tplText(ws) {
  ws.getRange('A1').setValue('URL PANEL DE CONTROL');
  ws.getRange('B2').setValue('Carpetas referentes');
  ws.getRange('A3').setValue('CARPETA CUADRO SUP.');
  ws.getRange('A4').setValue('ID CARPETA CUADRO SUP.');
  ws.getRange('A5').setValue('CARPETA PANEL DE CONTROL');
  ws.getRange('A6').setValue('ID CARPETA PANEL DE CONTROL');
  ws.getRange('A7').setValue('CARPETA BACKUP');
  ws.getRange('A8').setValue('ID CARPETA BACKUP');
  ws.getRange('A9').setValue('DESCARGAR ARCHIVO XLSX');
}

function sheetFormatter(ws) {
  // Estilo global
  ws.getRange(1, 1, ws.getMaxRows(), ws.getMaxColumns()).setFontSize(13).setFontFamily('Montserrat').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle')
    .setFontColor('#B7B7B7');
  // Col A
  ws.getRange(1, 1, 9, 1).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontWeight('bold');
  // Row 2
  ws.getRange(2, 1, 1, 4).setFontFamily('Inconsolata').setFontSize(16).setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#999999')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  // Col B
  ws.getRange(3, 2, 7, 1).setBackground('#F3F3F3').setBorder(true, true, true, true, true, true, '#CCCCCC', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // ImportRanges
  ws.getRange(1, 3, 1, 2).setBackground('#EAD1DC').setBorder(true, true, true, true, true, true, '#A64D79', SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor('#A64D79')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  // Control Panel
  ws.getRange('B1').setBackground('#D9EAD3').setBorder(true, true, true, true, true, true, '#6AA886', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#6AA886')
    .setFontWeight('bold');
  // Folders Bold
  let rangeListBold = ws.getRangeList(['B3', 'B5', 'B7', 'B9']);
  rangeListBold.setFontWeight('bold');
  /*
  ws.getRange('B3').setFontWeight('bold');
  ws.getRange('B5').setFontWeight('bold');
  ws.getRange('B7').setFontWeight('bold');
  ws.getRange('B9').setFontWeight('bold');*/
  // Column/Row Size
  ws.setColumnWidth(1, 330);
  ws.setColumnWidth(2, 380);
  ws.setColumnWidth(3, 285);
  ws.setColumnWidth(4, 400);

  let maxRows = ws.getMaxRows();
  for (let i = 1; i < maxRows + 1; i++) {
    ws.setRowHeight(i, 27);
  }
  ws.setRowHeight(2, 50);
}

function importRangeToken(ss_id, tokenID) { // TokenID is the ID of destination Google Sheet
  let url = `https://docs.google.com/spreadsheets/d/${ss_id}/externaldata/addimportrangepermissions?donorDocId=${tokenID}`;
  let token = ScriptApp.getOAuthToken();
  let params = {
    method: 'post',
    headers: {
      Authorization: `Bearer ${token}`,
    },
    muteHttpExceptions: true,
  };

  UrlFetchApp.fetch(url, params);
}
