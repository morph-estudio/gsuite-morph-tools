/**
 * Gsuite Morph Tools - CS Updater 1.5.0
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

function morphCSUpdater(btnID, rowData) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('LINK') || ss.insertSheet('LINK', 1);
  let ss_id = ss.getId();
  let userMail = Session.getActiveUser().getEmail();
  let dateNow = Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy - HH:mm:ss');

  let file = DriveApp.getFileById(ss_id);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  let formData = [
    rowData.updatePrefix,
    rowData.prefixAll,
  ];

  let [updatePrefix, prefixAll] = formData;  

  if (btnID === 'csUpdater') {
    sh.clear().clearFormats();
    templateFormat(sh); // Formato de celdas
  }

  // Carpeta Cuadro de Superficies

  let parents = file.getParents();
  let carpetaBaseID = parents.next().getId();
  let carpetaBase = DriveApp.getFolderById(carpetaBaseID);

  // Panel de control

  let filA = getControlPanel(sh, file, btnID);
  let [filePanelName, filePanelId, filePanelUrl, folderPanelcId] = filA;

  let panelControl = DriveApp.getFileById(filePanelId);
  let folderPanelcName = DriveApp.getFolderById(folderPanelcId);
  panelControl.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // Set Template Values

  templateText(sh);

  sh.getRange('B1').setValue(filePanelUrl);
  sh.getRange('B3').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${carpetaBaseID}";"${carpetaBase}")`).setFontColor('#0000FF');
  sh.getRange('B4').setValue(carpetaBaseID);
  sh.getRange('B5').setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${folderPanelcId}";"${folderPanelcName}")`).setFontColor('#0000FF');
  sh.getRange('B6').setValue(folderPanelcId);

  // ImportRange Permission

  let sectoresID = '1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro';

  importRangeToken(ss_id, carpetaBaseID);
  importRangeToken(ss_id, sectoresID);

  sh.getRange('C1').setValue('=IMPORTRANGE(B1;"Instrucciones!A1")');
  sh.getRange('D1').setValue('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"DB-SI!B2")');

  Utilities.sleep(250);

  // Localizar archivos TXT exportados

  let list = [];

  if (btnID === 'csUpdater') {
    sh.getRange(3, 3, 6, 2).clear();

    let searchFor = 'title contains "ExpTXT"';
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

    keepNewestFilesOfEachNameInAFolder(expFolderDef); // Delete duplicated files in Exports Folder

    let sufix = updatePrefix || 'TXT'; // mask
    let files = expFolderDef.getFiles();
    while (files.hasNext()) {
      file = files.next();
      filename = file.getName();
      if (prefixAll === true) {
        if (filename.includes('.txt')) {
          list.push([file.getName(), file.getId(), file.getName().slice(0, -4).replace('Sheets ', '').toUpperCase()]);
        }
      } else {
        if (filename.includes(sufix)) {
          list.push([file.getName(), file.getId(), file.getName().slice(0, -4).replace('Sheets ', '').toUpperCase()]);
        }
      }
    }

    let result = [['Archivos exportados', 'IDs', 'Hoja'], ...list.filter((r) => r[0].includes('.txt')).sort()];

    
    let resultCrop = result.map((val) => val.slice(0, -1));
    sh.getRange(2, 3, result.length, 2).setValues(resultCrop);

    // Filelist Styling

    sh.getRange(3, 3, list.length, 2).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontSize(14).setFontFamily('Montserrat').setFontColor('#78909c').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment('middle');
    sh.getRange(3, 3, list.length, 1).setFontWeight('bold');
    sh.getRange(3, 4, list.length, 1).setBackground('#fafafa');
  } else if (btnID === 'csManual2') {
    let txtFileId_FT = sh.getRange(3, 4).getValue();
    let txtFile_FT = DriveApp.getFileById(txtFileId_FT);
    let txtFileId_SP = sh.getRange(4, 4).getValue();
    let txtFile_SP = DriveApp.getFileById(txtFileId_SP);
    let txtFileId_VN = sh.getRange(5, 4).getValue();
    let txtFile_VN = DriveApp.getFileById(txtFileId_VN);
    let files = [txtFile_FT, txtFile_SP, txtFile_VN];

    txtFile_FT.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    txtFile_SP.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    txtFile_VN.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    files.forEach((file) => {
      list.push([file.getName(), file.getId(), file.getName().slice(0, -4).replace('Sheets ', '').toUpperCase()]);
    });
  }

  // Copy data in Sheets

  SpreadsheetApp.flush();

  for (const [txtFileName, txtFileId, txtFileSheet] of list) {
    let txt_file = DriveApp.getFileById(txtFileId);
    txt_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    let tsvUrl = `https://drive.google.com/uc?id=${txtFileId}&x=.tsv`;
    let tsvContent = UrlFetchApp.fetch(tsvUrl, {muteHttpExceptions: true }).getContentText();
    let tsvData = Utilities.parseCsv(tsvContent, '\t');
    let sheetPaste = ss.getSheetByName(`${txtFileSheet}`) || ss.insertSheet(`${txtFileSheet}`, 100);
    sh.activate();
    sheetPaste.setTabColor('F1C232').hideSheet();
    sheetPaste.clear();
    sheetPaste.getRange(1, 1, tsvData.length, tsvData[0].length).setValues(tsvData);
    SpreadsheetApp.flush();
  };

  sh.getRange('C2').setNote(null).setNote(`Última actualización: ${dateNow} por ${userMail}`); // Last Update Note
  deleteEmptyRows(); removeEmptyColumns();
  sh.activate();
}

function getControlPanel(sh, file, btnID) {
  let filA = [];

  if (btnID === 'csUpdater') {
    let parents = file.getParents();
    let fldrA = [];
    while (parents.hasNext()) {
      let f = parents.next();
      let f_id = f.getId();
      fldrA.push(f_id);
      parents = f.getParents();
    };

    let pcMask = 'Panel de control';
    for (let i = 0; i < fldrA.length; i++) {
      let files = DriveApp.getFolderById(fldrA[i]).getFilesByType(MimeType.GOOGLE_SHEETS);
      while (files.hasNext()) {
        let filePC = files.next();
        if (filePC.getName().includes(pcMask)) {
          filA.push(filePC.getName(), filePC.getId(), filePC.getUrl(), fldrA[i]);
        }
      }
    }
  } else if (btnID === 'csManual2') {
    let filePanelUrl = sh.getRange(1, 2).getValue();
    let filePanelId = getIdFromUrl(filePanelUrl);
    let filePC = DriveApp.getFileById(filePanelId);
    filA.push(filePC.getName(), filePanelId, filePanelUrl, filePC.getParents().next().getId());
  }
  return filA;
}

function templateText(sh) {
  sh.getRange('A1').setValue('URL PANEL DE CONTROL');
  sh.getRange('B2').setValue('Carpetas referentes');
  sh.getRange('A3').setValue('CARPETA CUADRO SUP.');
  sh.getRange('A4').setValue('ID CARPETA CUADRO SUP.');
  sh.getRange('A5').setValue('CARPETA PANEL DE CONTROL');
  sh.getRange('A6').setValue('ID CARPETA PANEL DE CONTROL');
  sh.getRange('A7').setValue('CARPETA BACKUP');
  sh.getRange('A8').setValue('ID CARPETA BACKUP');
  sh.getRange('A9').setValue('DESCARGAR ARCHIVO XLSX');
  sh.getRange('C1').setValue('=IMPORTRANGE(B1;"Instrucciones!A1")');
  sh.getRange('D1').setValue('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"DB-SI!B2")');
  sh.getRange('C2').setValue('Archivos importados');
  sh.getRange('D2').setValue('IDs');
}

function templateFormat(sh) {
  // Estilo global
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontSize(14).setFontFamily('Montserrat').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle')
    .setFontColor('#78909c');
  // Col A
  sh.getRange(1, 1, 9, 1).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontWeight('bold');
  // Row 2
  sh.getRange(2, 1, 1, 4).setFontFamily('Inconsolata').setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  // Col B
  sh.getRange(3, 2, 7, 1).setBackground('#fafafa').setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // ImportRanges
  sh.getRange(1, 3, 1, 2).setBackground('#e0f7fa').setBorder(true, true, true, true, true, true, '#26c6da', SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor('#26c6da')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  // Control Panel
  sh.getRange('B1').setBackground('#ECFDF5').setBorder(true, true, true, true, true, true, '#00C853', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#00C853')
    .setFontWeight('bold');
  // Folders Bold
  sh.getRangeList(['B3', 'B5', 'B7', 'B9']).setFontWeight('bold');
  // Column/Row Size
  sh.setColumnWidth(1, 340);
  sh.setColumnWidth(2, 380);
  sh.setColumnWidth(3, 340);
  sh.setColumnWidth(4, 380);

  for (let i = 1; i < sh.getMaxRows() + 1; i++) {
    sh.setRowHeight(i, 28);
  }
  sh.setRowHeight(1, 35); sh.setRowHeight(2, 50);
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

function manualUpdaterTemplate() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName('LINK') || ss.insertSheet('LINK', 1);

  sh.clear().clearFormats();
  templateText(sh);

  sh.getRange('C2').setValue('Archivos exportados');
  sh.getRange('D2').setValue('IDs');
  sh.getRange('C3').setValue('TXT Sheets Falsos techos.txt');
  sh.getRange('C4').setValue('TXT Sheets Superficies.txt');
  sh.getRange('C5').setValue('TXT Sheets Ventanas.txt');

  templateFormat(sh); // Formato de celdas

  // Control Panel 
  sh.getRangeList(['B1', 'B8']).setBackground('#FFFDE7').setBorder(true, true, true, true, true, true, '#FBC02D', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#FBC02D');
  sh.getRange('B1').setFontWeight('bold');
  // Filelist Style
  sh.getRange(3, 3, 3, 2).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  sh.getRange(3, 3, 3, 1).setFontWeight('bold');
  sh.getRange(3, 4, 3, 1).setBackground('#FFFDE7').setBorder(true, true, true, true, true, true, '#FBC02D', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#FBC02D');
  // ImportRange Cells
  sh.getRange('C1').setValue('=IMPORTRANGE(B1;"Instrucciones!A1")');
  sh.getRange('D1').setValue('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"DB-SI!B2")');

  deleteEmptyRows(); removeEmptyColumns();
  sh.activate();
}
