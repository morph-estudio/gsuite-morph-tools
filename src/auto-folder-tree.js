/**
 * Gsuite Morph Tools - Morph autoFolderTree 1.3
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

/* eslint-disable guard-for-in */
/* eslint-disable no-restricted-syntax */

function autoFolderTree() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getActiveSheet();
  let userMail = Session.getActiveUser().getEmail();
  let dateNow = Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy - HH:mm:ss');
  let niveles = [1, 2, 3, 4, 5, 6, 7];

  let result = ui().alert(
    '¿Quieres crear una copia de la hoja?',
    'Las fórmulas de la plantilla actual se sustituirán por las nuevas carpetas creadas. Si no haces una copia perderás la plantilla personalizada.',
    ui().ButtonSet.YES_NO,
  );

  if (result == ui().Button.YES) {
    let sheetName = sh.getSheetName();
    let copiedSheetIndex = sh.getIndex() + 1;
    sh.setName(`${sheetName} - Final`);
    sh.copyTo(ss).setName(sheetName).activate();
    ss.moveActiveSheet(copiedSheetIndex);
    sh.activate();
  }

  for (n in niveles) {
    if (n == 0) Logger.log('holaaaa');
    let levelInput = niveles[n];
    let Level = levelInput * 2 + 1;
    let numRows = sh.getLastRow(); // Number of rows to process
    let dataRange = sh.getRange(2, Number(Level) - 1, numRows, Number(Level)); // startRow, startCol, endRow, endCol
    let data = dataRange.getValues();
    let parentFolderID = new Array();
    let theParentFolder;

    for (let i in data) {
      parentFolderID[i] = data [i][0];
      if (data [i][0] == '') {
        parentFolderID[i] = parentFolderID[i - 1];
      }
    }

    for (let i in data) {
      
      if (data [i][1] !== '') {
        if (n == 0) {
          theParentFolder = DriveApp.getFolderById(getIdFromUrl(parentFolderID[i]));
          Logger.log('cosasidtheparent ' + theParentFolder)
        } else {
          theParentFolder = DriveApp.getFolderById(parentFolderID[i]);
        }
        let folderName = data[i][1];
        let theChildFolder = theParentFolder.createFolder(folderName);
        let newFolderID = sh.getRange(Number(i) + 2, Number(Level) + 1);
        let folderIdValue = theChildFolder.getId();
        newFolderID.setValue(folderIdValue);
        let addLink = sh.getRange(Number(i) + 2, Number(Level));
        let value = addLink.getDisplayValue();
        addLink.setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${folderIdValue}";"${value}")`);
        SpreadsheetApp.flush();
      }
    }
    sh.getRange('B2').clearNote().setNote(`Estructura creada el ${dateNow} por ${userMail}`);
    SpreadsheetApp.flush();
  }
}

function autoFolderTreeTpl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  sh.clear().clearFormats();

  // Copy Data from TSV

  let externalFolderId = '1BwVkhZsDQh-FO3Jgj4pwPu-p4WFG9wr-';
  let fileName = 'autoFolderTree.txt';
  let fileId;
  let filesFound = searchFile(fileName, externalFolderId);
  for (let file of filesFound) {
    fileId = file.getId();
  }
  let tsvUrl = `https://drive.google.com/uc?id=${fileId}&x=.tsv`;
  let tsvContent = UrlFetchApp.fetch(tsvUrl, {}).getContentText();
  let tsvData = Utilities.parseCsv(tsvContent, '\t');
  sh.getRange(1, 1, tsvData.length, tsvData[0].length).setValues(tsvData);

  // Global Style
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontSize(12).setFontFamily('Inter').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle');
  // Levels of Structure
  sh.getRange(1, 3, 1, 13).setBackground('#546E7A').setFontColor('#fff');
  sh.getRange('B1').setBackground('#FFAB00').setBorder(true, true, true, true, true, true, '#FFAB00', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#fff');
  sh.getRange('B2').setBackground('#FFFDE7').setBorder(true, true, true, true, true, true, '#FFAB00', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#FFAB00').setNote(null).setNote(`Introduce en esta celda la dirección URL de la carpeta inicial de la estructura.`);
  // Style of Morph Project Template
  sh.getRange(1, 18, 1, 6).setBackground('#FFAB00').setBorder(true, true, true, true, true, true, '#FFAB00', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#fff');
  sh.getRange(2, 18, 1, 6).setBackground('#FFFDE7').setBorder(true, true, true, true, true, true, '#FFAB00', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#FFAB00').setHorizontalAlignment('center');

  let cell = sh.getRange('T2');
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(['AEI', 'E', 'I', 'IINT', 'I+D']).build();
  cell.setDataValidation(rule);
  
  sh.getRange(1, 1, 1, sh.getMaxColumns()).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  // Column Style
  sh.setFrozenRows(1);
  sh.hideColumns(4); sh.hideColumns(6); sh.hideColumns(8); sh.hideColumns(10);
  sh.hideColumns(12); sh.hideColumns(14); sh.hideColumns(16);
  sh.setColumnWidth(1, 25);
  sh.setColumnWidth(2, 250);
  sh.setColumnWidth(3, 230);
  sh.setColumnWidth(5, 230);
  sh.setColumnWidth(7, 230);
  sh.setColumnWidth(9, 230);
  sh.setColumnWidth(11, 230);
  sh.setColumnWidth(13, 230);
  sh.setColumnWidth(15, 230);
  sh.setColumnWidth(17, 40);
  sh.setColumnWidth(21, 150);
  sh.setColumnWidth(22, 200);
  sh.setColumnWidth(23, 200);

  removeEmptyColumns();
  deleteEmptyRows();
  SpreadsheetApp.flush();
}
