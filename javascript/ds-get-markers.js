function doGet() {
  return HtmlService.createHtmlOutputFromFile('html/document-studio');
}

/*
* Gsuite Morph Tools - Morph Document Studio
* Developed by alsanchezromero
* Created on Mon Jul 25 2022
*
* Copyright (c) 2022 Morph Estudio
*/

// GET MARKERS MAIN FUNCTION

function getMarkers(rowData) {
  const ss = SpreadsheetApp.getActive();
  const sh = SpreadsheetApp.getActiveSheet();
  let lastColmn;

  /* TEST
    //var ss = SpreadsheetApp.openById('1v5f3X1ShmVCGdT6NdWvmJHcfeP01ptuwHfT1iqM6UQI');
    //var sh = ss.getSheetByName('EXEC')
    //var docID = "1Gs4Gd4JtVMrI-nu6ELZtS71kebUU9CN3II2Xz5-Q2F8";
  */

  // Data + Variables

  const formData = [rowData.templateID, rowData.greenCells1, rowData.purgeMarkers];
  let [docURL, greenCells1, purgeMarkers] = formData;

  let urlCell = sh.getLastColumn();
  let nameCell = urlCell - 1;
  let docID;

  if (greenCells1) {
    docURL = sh.getRange(1, nameCell).getValue();
    docID = getIdFromUrl(docURL);
  } else {
    docID = getIdFromUrl(docURL);
  }

  let identifier = {
    start: '{{',
    start_include: true,
    end: '}}',
    end_include: true,
  };

  let gDocTemplate = DriveApp.getFileById(docID);
  let fileType = gDocTemplate.getMimeType();
  let docMarkers;

  switch (fileType) {
    case MimeType.GOOGLE_DOCS:
      docMarkers = getDocItems(docID, identifier);
      break;
    case MimeType.GOOGLE_SLIDES:
      docMarkers = getSlidesItems(docID, identifier);
      break;
    default:
  }

  let updatedValues = [];
  let driverArray = docMarkers.flat(); // Slicing {} markers for enhancing interface.
  driverArray.forEach((el) => {
    let sliced = el.slice(2, -2);
    updatedValues.push(sliced);
  });

  // Purge Markers

  let isGrCell; let semiLastCell; let semiIsGrCell;

  if (purgeMarkers == true) {
    lastColmn = sh.getLastColumn();
    isGrCell = isGreenCell(lastColmn);
    semiLastCell = lastColmn - 1;
    semiIsGrCell = isGreenCell(semiLastCell);

    if (isGrCell == true && semiIsGrCell == true) {
      columnRemover(sh, updatedValues, 2);
    } else if (isGrCell == false && semiIsGrCell == false) {
      columnRemover(sh, updatedValues, 0);
    } else if (isGrCell == true && semiIsGrCell == false) {
      columnRemover(sh, updatedValues, 1);
    }
  }

  // Add New Markers

  lastColmn = sh.getLastColumn();
  if (lastColmn >= 1) {
    let newHeaderRange = sh.getRange(1, 1, 1, sh.getLastColumn());
    let headerValuesNew = newHeaderRange.getValues()[0];

    updatedValues.forEach((a, index) => {
      let i = headerValuesNew.indexOf(a);
      if (i === -1) {
        if (index === 0) {
          sh.insertColumns(index + 1);
        } else {
          sh.insertColumnAfter(index);
        }

        sh.setColumnWidth(index + 1, 150);
        let headerCell = sh.getRange(1, index + 1, 1, 1);
        headerCell.setValue(a);
      }
    });
  } else {
    updatedValues.forEach((a, index) => {
      sh.insertColumns(index + 1);
      sh.setColumnWidth(index + 1, 150);
      let firstCell= sh.getRange(1, index + 1, 1, 1);
      firstCell.setValue(a);
    });
  }

  // Style

  lastColmn = sh.getLastColumn();
  isGrCell = isGreenCell(lastColmn);
  semiLastCell = lastColmn - 1;
  semiIsGrCell = isGreenCell(semiLastCell);

  sh.getRange(1, 1, 1, sh.getMaxColumns()).setFontWeight('bold');
  sh.setFrozenRows(1);

  if (lastColmn === 0) {
    sh.insertColumns(1);
    sh.getRange(1, 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] File-links');
    sh.insertColumns(1);
    sh.getRange(1, 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] Files')
      // eslint-disable-next-line max-len
      .setNote('Celdas verdes: para utilizar la opción "usar celdas verdes" debes introducir en esta celda el LINK de la plantilla y en la siguiente columna el LINK de la carpeta de destino.');
  } else if (isGrCell == false) {
    sh.insertColumnAfter(lastColmn);
    sh.getRange(1, lastColmn + 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] File-links');
    sh.insertColumnAfter(lastColmn);
    sh.getRange(1, lastColmn + 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] Files')
      // eslint-disable-next-line max-len
      .setNote('Celdas verdes: para utilizar la opción "usar celdas verdes" debes introducir en esta celda el LINK de la plantilla y en la siguiente columna el LINK de la carpeta de destino.');
  } else if (isGrCell == true && semiIsGrCell == false) {
    sh.insertColumnAfter(lastColmn);
    sh.getRange(1, lastColmn + 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS]');
  }

  /*
    var docMarkersLenght = updatedValues.length;
    if (lastColmn === 0){
      sh.insertColumns(1)
      sh.getRange(1, 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] File-links')
      sh.insertColumns(1)
      sh.getRange(1, 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] Files')
      .setNote('Celdas verdes: para utilizar la opción "usar celdas verdes" debes introducir en esta celda el LINK de la plantilla y en la siguiente columna el LINK de la carpeta de destino.')
    } else if (lastColmn === docMarkersLenght){
      sh.insertColumnAfter(lastColmn)
      sh.getRange(1, lastColmn+1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] File-links')
      sh.insertColumnAfter(lastColmn)
      sh.getRange(1, lastColmn+1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] Files')
      .setNote('Celdas verdes: para utilizar la opción "usar celdas verdes" debes introducir en esta celda el LINK de la plantilla y en la siguiente columna el LINK de la carpeta de destino.')
    } else if (lastColmn === docMarkersLenght + 1){
      sh.insertColumnAfter(lastColmn)
      sh.getRange(1, lastColmn+1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS]')
    };
  */

  lastColmn = sh.getLastColumn();
  let lastCol2 = sh.getLastColumn() - 1;
  sh.setColumnWidth(lastColmn, 300); sh.setColumnWidth(lastCol2, 300);
  removeEmptyColumns();
}
