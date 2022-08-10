/**
 * Gsuite Morph Tools - Morph autoFolderTree 1.2
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

/* eslint-disable guard-for-in */
/* eslint-disable no-restricted-syntax */

function autoFolderTree() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  let niveles = [1, 2, 3, 4, 5, 6, 7];

  sh.activate();

  let cell = sh.getRange('B3');
  if (cell.isBlank()) {
    const ui = SpreadsheetApp.getUi();
    let result = ui.prompt(
      'ID Carpeta',
      'Introduce el ID de la carpeta donde crear la estructura:',
      ui.ButtonSet.OK_CANCEL,
    );

    let button = result.getSelectedButton();
    let userGetID = result.getResponseText();
    if (button == ui.Button.OK) {
      // call function and pass the value
      cell.setValue(userGetID);
    }
  }

  for (n in niveles) {
    let levelInput = niveles[n];
    let Level = levelInput * 2 + 1;
    let numRows = sh.getLastRow(); // Number of rows to process
    let dataRange = sh.getRange(3, Number(Level) - 1, numRows, Number(Level)); // startRow, startCol, endRow, endCol
    let data = dataRange.getValues();
    let parentFolderID = new Array();

    for (let i in data) {
      parentFolderID[i] = data [i][0];
      if (data [i][0] == '') {
        parentFolderID[i] = parentFolderID[i - 1];
      }
    }

    for (let i in data) {
      if (data [i][1] !== '') {
        let theParentFolder = DriveApp.getFolderById(parentFolderID[i]);
        let folderName = data[i][1];
        let theChildFolder = theParentFolder.createFolder(folderName);
        let newFolderID = sh.getRange(Number(i) + 3, Number(Level) + 1);
        let folderIdValue = theChildFolder.getId();
        newFolderID.setValue(folderIdValue);
        let addLink = sh.getRange(Number(i) + 3, Number(Level));
        let value = addLink.getDisplayValue();
        addLink.setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${folderIdValue}","${value}")`);
      }
    }
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
  let tsvContent = UrlFetchApp.fetch(tsvUrl, {muteHttpExceptions: true }).getContentText();
  let tsvData = Utilities.parseCsv(tsvContent, '\t');
  sh.getRange(1, 1, tsvData.length, tsvData[0].length).setValues(tsvData);

  // Global Style
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontSize(12).setFontFamily('Montserrat').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle');

  // Levels of Structure
  sh.getRange(1, 3, 1, 13).setBackground('#434343').setFontColor('#fff');
  sh.getRange('B1').setBackground('#BF9000').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#fff');
  sh.getRange('B3').setBackground('#FFF2CC').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#BF9000');

  // Style of Morph Project Template
  sh.getRange(1, 18, 1, 6).setBackground('#BF9000').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#fff');
  sh.getRange(2, 18, 1, 6).setBackground('#FFF2CC').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#BF9000').setHorizontalAlignment('center');

  let cell = sh.getRange('T2');
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(['AEI', 'E', 'I', 'IINT', 'I+D']).build();

  cell.setDataValidation(rule);
  sh.getRange(1, 1, 1, sh.getMaxColumns()).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  // Column Style
  sh.setFrozenRows(1);
  sh.hideColumns(4); sh.hideColumns(6); sh.hideColumns(8); sh.hideColumns(10);
  sh.hideColumns(12); sh.hideColumns(14); sh.hideColumns(16);
  sh.setColumnWidth(1, 25);
  sh.setColumnWidth(2, 220);
  sh.setColumnWidth(3, 200);
  sh.setColumnWidth(5, 200);
  sh.setColumnWidth(7, 200);
  sh.setColumnWidth(9, 200);
  sh.setColumnWidth(11, 200);
  sh.setColumnWidth(13, 200);
  sh.setColumnWidth(15, 200);
  sh.setColumnWidth(17, 40);
  sh.setColumnWidth(21, 150);
  sh.setColumnWidth(22, 200);
  sh.setColumnWidth(23, 200);

  removeEmptyColumns();
  deleteEmptyRows();
  SpreadsheetApp.flush();
}
