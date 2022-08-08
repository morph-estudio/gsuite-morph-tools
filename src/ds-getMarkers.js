/**
 * Gsuite Morph Tools - Morph Document Studio (Get Markers)
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

function getMarkers(rowData) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  // Data + Variables

  let formData = [rowData.templateID, rowData.greenCells, rowData.purgeMarkers];
  let [docURL, greenCells, purgeMarkers] = formData;

  let filenameField = '[DS] Files'; let fileurlField = '[DS] File-links'; let mailurlField = '[DS] Email-links';

  let dataReturn = getGreenColumns(sh, filenameField, fileurlField);

  let indexData = [
    dataReturn.indexNameCell,
    dataReturn.indexUrlCell,
  ];

  let [indexNameCell, indexUrlCell] = indexData;

  let docID;
  let sheetEmpty = sh.getLastColumn();

  if (greenCells) {
    docURL = sh.getRange(1, indexNameCell + 1).getNotes();
    docID = getIdFromUrl(docURL[0][0]);
  } else {
    docID = getIdFromUrl(docURL);
  }

  dataReturn = getInternallyMarkers(docID)

  indexData = [
    dataReturn.docMarkers,
    dataReturn.gDocTemplate,
    dataReturn.fileType,
  ];

  let [docMarkers, gDocTemplate,fileType] = indexData;

  let notAllMarkersChanged; let headerValues; let updatedValues = [];
  
  let driverArray = docMarkers.flat(); // Slicing {} markers for enhancing interface.
  driverArray.forEach((el) => {
    let sliced = el.slice(2, -2);
    updatedValues.push(sliced);
  });

  // Purge Markers

  if (purgeMarkers) {

    headerValues = flatten(emailDropdown()).filter(e => e !== filenameField && e !== fileurlField && e !== mailurlField);

    if (headerValues.length != 0) {
      notAllMarkersChanged = updatedValues.some(element => {
        return headerValues.includes(element);
      }); Logger.log('allmakers: ' + notAllMarkersChanged)

      if (notAllMarkersChanged) {
        columnRemover(sh, updatedValues, headerValues);
      } else {
        indexNameCell = fieldIndex(sh, filenameField);
        sh.deleteColumns(1, indexNameCell);
      }
    }
  }

  // Add New Markers

  headerValues = flatten(emailDropdown()).filter(e => e !== filenameField && e !== fileurlField && e !== mailurlField);

  updatedValues.forEach((a, index) => {
    if (headerValues.indexOf(a) === -1) {
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

  // Style

  indexNameCell = fieldIndex(sh, filenameField);

  if (sheetEmpty <= 2 || notAllMarkersChanged === false) {
    sh.getRange(1, 1, 1, indexNameCell).clearFormat();
  }

  sh.getRange(1, 1, 1, sh.getMaxColumns()).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  removeEmptyColumns();
}
