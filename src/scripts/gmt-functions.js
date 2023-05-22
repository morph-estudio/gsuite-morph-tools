// CATEGORY: RANGES

/**
 * deleteNamedRanges
 * Remove all Named Ranges in a Spreadsheet
 */
 function deleteNamedRanges() {
  var ss = SpreadsheetApp.getActive();
  var namedRanges = ss.getNamedRanges();
  for (var i = 0; i < namedRanges.length; i++) {
    namedRanges[i].remove();
  }
}

/**
 * refreshNamedRanges
 * Refresh Named Ranges in a Morph Table based on TXT_OP_TPL
 */
function refreshNamedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('X_Variables');
  var sh_txtop = ss.getSheetByName('TXT OPERACIONES');

  if (!sh || !sh_txtop) {
    throw new Error(`No se encontraron las hojas 'TXT OPERACIONES' o 'X_Variables'`);
  }

  var headers = sh_txtop.getRange(1, 1, 1, sh_txtop.getLastColumn()).getValues()[0];

  var headers_template = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var parametrosColumnIndex = headers_template.indexOf("PLANTILLA PARAMETRO ARCHI");

  if (parametrosColumnIndex !== -1 && parametrosColumnIndex < headers_template.length - 1) {
    var columnARange = sh.getRange(2, parametrosColumnIndex + 1, sh.getLastRow() - 1, 1);
    var columnBRange = sh.getRange(2, parametrosColumnIndex + 2, sh.getLastRow() - 1, 1);

    var columnA = columnARange.getValues();
    var columnB = columnBRange.getValues();
  } else {

  }

  // Caché de la lista de rangos nombrados
  var namedRanges = ss.getNamedRanges();
  var namedRangesDict = {};
  for (var i = 0; i < namedRanges.length; i++) {
    namedRangesDict[namedRanges[i].getName()] = namedRanges[i];
  }

  var namedRangeName, namedRange, txtvalue;

  for (var i = 0; i < columnB.length; i++) {
    namedRangeName = columnB[i][0];
    txtvalue = columnA[i][0];

    if (!namedRangeName) {
      continue;
    }

    var columnIndex = headers.indexOf(txtvalue);

    if (columnIndex === -1) {
      continue;
    }

    columnIndex += 1;

    namedRange = sh_txtop.getRange(2, columnIndex, sh_txtop.getMaxRows() - 1, 1);

    // Buscar si el rango nombrado existe utilizando el diccionario
    var existingNamedRange = namedRangesDict[namedRangeName];
    if (existingNamedRange) {
      existingNamedRange.setRange(namedRange);
    } else {
      ss.setNamedRange(namedRangeName, namedRange);
    }
  }
}

// CATEGORY: ROW/COLUMN DELETION AND SPREADSHEET OPTIMIZATION

/**
 * deleteEmptyRows, removeEmptyColumns
 * Delete the rows and columns up to the last one that contains data
 */
function deleteEmptyRows(sh) {
  sh = sh || SpreadsheetApp.getActiveSheet();
  let maxRows = sh.getMaxRows();
  let lastRow = sh.getLastRow();
  let checker2 = maxRows - (maxRows - lastRow);
  if (checker2 != maxRows) {
    sh.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

function removeEmptyColumns(sh) {
  sh = sh || SpreadsheetApp.getActiveSheet();
  let maxCol = sh.getMaxColumns();
  let lastCol = sh.getLastColumn();
  let checker = maxCol - (maxCol - lastCol);
  if (checker != maxCol) {
    sh.deleteColumns(lastCol + 1, maxCol - lastCol);
  }
}

/**
 * deleteAllEmptyRows
 * Delete any row without data in a specified sheet
 */
 function deleteAllEmptyRows() {
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getActiveSheet();
  let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  let row = sheet.getLastRow();

  while (row > 2) {
    let rec = data.pop();
    if (rec.join('').length === 0) {
      sheet.deleteRow(row);
    }
    row--;
  }
  let maxRows = sheet.getMaxRows();
  let lastRow = sheet.getLastRow();
  if (maxRows - lastRow != 0) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

/**
 * deleteUntilLastDataRow
 * Delete any row without data in a specified sheet.
 * DON'T KNOW IF "deleteAllEmptyRows" IS THE SAME FUNCTION <- REVISAR
 */
function deleteUntilLastDataRow(sh) {
  sh = sh || SpreadsheetApp.getActiveSheet();
  var maxRows = sh.getMaxRows();
  var lastDataRow = 0;
  
  for (var i = maxRows; i >= 1; i--) {
    var rowRange = sh.getRange(i, 1, 1, sh.getLastColumn());
    var rowValues = rowRange.getValues()[0];
    var rowIsEmpty = true;
    
    for (var j = 0; j < rowValues.length; j++) {
      if (rowValues[j]) {
        rowIsEmpty = false;
        break;
      }
    }
    
    if (rowIsEmpty) {
      sh.deleteRow(i);
    } else {
      lastDataRow = i;
      break;
    }
  }
  
  return lastDataRow;
}

/**
 * deleteExcessiveRows, optimizeSpreadsheet
 * Spreadsheet Optimization based on excessive empty rows
 */
function deleteExcessiveRows(rowsInput, sh) {
  sh = sh || SpreadsheetApp.getActiveSheet();
  let maxRows = sh.getMaxRows();
  if (maxRows < rowsInput) throw new Error(`La hoja ya tiene menos de ${rowsInput} filas.`)
  sh.deleteRows(rowsInput, maxRows - rowsInput);
}

function optimizeSpreadsheet(rowsInput) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var sheetList = { sheets: [], sheetNames: [] };

  sheets.forEach(function(sh) {
    let maxRows = sh.getMaxRows();
    if (maxRows > rowsInput) {
      sheetList.sheets.push(sh);
      sheetList.sheetNames.push(sh.getName());
    }
  });

  if (sheetList.sheets.length === 0) throw new Error(`No hay ninguna hoja con más de ${rowsInput} filas.`)

  var sheetNamesWithFormat = sheetList.sheetNames.map(name => ` ${name}`);

  var confirm = Browser.msgBox('Optimización del documento', `Las siguientes hojas tienen más de ${rowsInput} filas:\\n\\n${sheetNamesWithFormat}\\n\\nSuponiendo que las celdas sobrantes están vacías, ¿Quieres borrar las celdas vacías hasta la fila ${rowsInput} y así optimizar el documento?`, Browser.Buttons.OK_CANCEL);
  if (confirm == 'ok') {
    sheetList.sheets.forEach(function(sh) {
      let maxRows = sh.getMaxRows();
      sh.deleteRows(rowsInput, maxRows - rowsInput);
    });
  } else {
    throw new Error(`No se ha optimizado el documento.`)
  }
}

/**
 * purgeDocumentCache
 * Borrar la memoria caché de la hoja de cálculo
 */
 function purgeDocumentCache() {
  SpreadsheetApp.flush()
}

// CATEGORY: SHEET FORMATTING

/**
 * increaseFontSize, decreaseFontSize
 * Modify the text font size proportionally throughout the sheet
 */
function increaseFontSize() {
  const range = sh().getDataRange();
  let fontsizes = range.getFontSizes();
  let numRows = range.getNumRows();
  let numCols = range.getNumColumns();

  for (let i = 0; i < numRows; i++) {
    for (let j = 0; j < numCols; j++) {
      range.getCell(i + 1,j + 1).setFontSize(fontsizes[i][j] + 2)
    }
  }
}

function decreaseFontSize() {
  const range = sh().getDataRange();
  let fontsizes = range.getFontSizes();
  let numRows = range.getNumRows();
  let numCols = range.getNumColumns();

  for (let i = 0; i < numRows; i++) {
    for (let j = 0; j < numCols; j++) {
      range.getCell(i + 1,j + 1).setFontSize(fontsizes[i][j] - 2)
    }
  }
}

// CATEGORY: UTILITIES

/**
 * getDomainUsersList
 * Genera una lista con los usuarios del servidor Morph
 */
function getDomainUsersList() {
  let users = [];
  let options = {
    domain: 'morphestudio.es', // Google Workspace domain name
    customer: 'my_customer',
    maxResults: 200,
    projection: 'basic', // Fetch basic details of users
    viewType: 'domain_public',
    orderBy: 'email', // Sort results by users
  };

  do {
    var response = AdminDirectory.Users.list(options);
    response.users.forEach((user) => {
      users.push([user.name.fullName, user.primaryEmail]);
    });

    // For domains with many users, the results are paged
    if (response.nextPageToken) {
      options.pageToken = response.nextPageToken;
    }
  } while (response.nextPageToken);

  // Insert data in a spreadsheet
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getActiveSheet();
  let a1cell = sh.getActiveCell().getA1Notation();
  let splitArray = getSplitA1Notation(a1cell);
  sh.getRange(splitArray[1], letterToColumn(splitArray[0]), users.length, users[0].length).setValues(users);
  //sh.getRange(1, 1, users.length, users[0].length).setValues(users);
  sh.setColumnWidth(letterToColumn(splitArray[0]), 250);
  sh.setColumnWidth(letterToColumn(splitArray[0]) + 1, 280);
}

/**
 * uniqueIdentifier
 * Genera identificadores únicos en las celdas seleccionadas.
 */
function uniqueIdentifier() {
  const selection = sh().getActiveRange();
  const columns = selection.getNumColumns();
  const rows = selection.getNumRows();
  for (let column = 1; column <= columns; column++) {
    for (let row = 1; row <= rows; row++) {
      const cell = selection.getCell(row, column);
      cell.setValue(Utilities.getUuid());
    }
  }
}

// FUNCTIONS IN FILES AND FOLDERS SECTION

/**
 * listFilesInFolder
 * Crea una lista de archivos dentro de una carpeta de Google Drive
 */
function listFilesInFolder(rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  let a1cell = sh.getActiveCell().getA1Notation();
  let splitArray = getSplitA1Notation(a1cell);

  let formData = [rowData.listFolderID, rowData.useA1]; let [listFolderID, useA1] = formData;

  let fldr_id;
  if (useA1) {
    fldr_URL = sh.getRange(1, 1).getNote();
    fldr_id = getIdFromUrl(fldr_URL);
  } else {
    fldr_id = getIdFromUrl(listFolderID);
  }

  Logger.log(`File: ${ss.getName()}, Sheetname: ${sh.getName()}, UseA1: ${useA1} FolderID: ${fldr_id}`)

  let fldr = DriveApp.getFolderById(fldr_id);
  let files = fldr.getFiles();
  let names = [];
  while (files.hasNext()) {
    f = files.next();
    f_url = f.getUrl();
    f_name = f.getName().replace(/\.[^/.]+$/, '');
    f_mime = shortenMimetype(f.getMimeType());
    names.push([f_name,f_url,f_mime]);
  }
  let result = [['Filename', 'File URL', 'Type'], ...names.sort()];
  sh.getRange(splitArray[1], letterToColumn(splitArray[0]), names.length + 1, 3).setValues(result);
}

function shortenMimetype(mimetype) {

  const mimetypes = {
    "application/vnd.google-apps.script": "GOOGLE_APPS_SCRIPT",
    "application/vnd.google-apps.drawing": "GOOGLE_DRAWINGS",
    "application/vnd.google-apps.document": "GOOGLE_DOCS",
    "application/vnd.google-apps.form": "GOOGLE_FORMS",
    "application/vnd.google-apps.spreadsheet": "GOOGLE_SHEETS",
    "application/vnd.google-apps.site": "GOOGLE_SITES",
    "application/vnd.google-apps.presentation": "GOOGLE_SLIDES",
    "application/vnd.google-apps.folder": "FOLDER",
    "application/vnd.google-apps.shortcut": "SHORTCUT",
    "image/bmp": "BMP",
    "image/gif": "GIF",
    "image/jpeg": "JPEG",
    "image/png": "PNG",
    "image/svg+xml": "SVG",
    "application/pdf": "PDF",
    "text/css": "CSS",
    "text/csv": "CSV",
    "text/html": "HTML",
    "application/javascript": "JAVASCRIPT",
    "text/plain": "PLAIN_TEXT",
    "application/rtf": "RTF",
    "application/vnd.oasis.opendocument.graphics": "OPENDOCUMENT_GRAPHICS",
    "application/vnd.oasis.opendocument.presentation": "OPENDOCUMENT_PRESENTATION",
    "application/vnd.oasis.opendocument.spreadsheet": "OPENDOCUMENT_SPREADSHEET",
    "application/vnd.oasis.opendocument.text": "OPENDOCUMENT_TEXT",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": "MICROSOFT_EXCEL",
    "application/vnd.ms-excel": "MICROSOFT_EXCEL_LEGACY",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": "MICROSOFT_POWERPOINT",
    "application/vnd.ms-powerpoint": "MICROSOFT_POWERPOINT_LEGACY",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": "MICROSOFT_WORD",
    "application/msword": "MICROSOFT_WORD_LEGACY",
    "application/zip": "ZIP",
    "video/mp4": "VIDEO_MP4"
  };

  return mimetypes[mimetype] || "UNKNOWN";
}

/**
 * insertImagesOfFolder
 * Insert the images from a folder into the selected cells
 */
function insertImagesOfFolder(rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  let formData = [rowData.listFolderID, rowData.useA1, rowData.imageFolderFileID, rowData.imageFolderFileName, rowData.imageFolderImage, rowData.imageFolderArrayFormula];
  let [folderUrl, useA1, imageFolderFileID, imageFolderFileName, imageFolderImage, imageFolderArrayFormula] = formData;

  let folderID;
  if (useA1) folderUrl = sh.getRange(1, 1).getNote();
  folderID = getIdFromUrl(folderUrl);

  let folder = DriveApp.getFolderById(folderID);
  let contents = folder.getFiles();
  let file; let downloadList = []; let cnt = 0;

  let selectedCell = sh.getActiveCell().getA1Notation();
  let a1NotationSplitArray = getSplitA1Notation(selectedCell);
  let baseURL = 'https://drive.google.com/uc?id='

  while (contents.hasNext()) {
    file = contents.next();
    cnt++;
    Logger.log(file.getMimeType())
    if ([MimeType.JPEG, MimeType.PNG, MimeType.GIF].includes(file.getMimeType())) {
      downloadList.push(file)
      Logger.log('fileperm ' + file.getSharingAccess())
      if(file.getSharingAccess() != 'ANYONE_WITH_LINK') file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
  };

  downloadList.sort().forEach((el, i) => {

    let listData = [ [] ];

    if (imageFolderFileName) listData[0].push(el.getName());
    if (imageFolderFileID) listData[0].push(el.getId())

    let count = Number(a1NotationSplitArray[1]) + Number(i);

    if (imageFolderImage) {

      sh.getRange(count, letterToColumn(a1NotationSplitArray[0]), 1, listData[0].length).setValues(listData); // Paste of not-image columns
      let image = SpreadsheetApp
                  .newCellImage()
                  .setSourceUrl(baseURL + el.getId())
                  .build();

      sh.getRange(count, letterToColumn(a1NotationSplitArray[0]) + listData[0].length, 1, 1).setValue(image);

    } else {

      listData[0].push(baseURL + el.getId()); // Push Public-URL to list
      sh.getRange(count, letterToColumn(a1NotationSplitArray[0]), 1, listData[0].length).setValues(listData); // Paste of not-image columns

      if (imageFolderArrayFormula) {
        let formulaRange = sh.getRange(Number(a1NotationSplitArray[1]), letterToColumn(a1NotationSplitArray[0]) + listData[0].length, 1, 1);
        let shiftedLetter = getShiftedLetter(a1NotationSplitArray[0], listData[0].length - 1);
        formulaRange.setFormula(`=ARRAYFORMULA(IMAGE($${shiftedLetter}$${Number(a1NotationSplitArray[1])}:$${shiftedLetter}$${Number(a1NotationSplitArray[1]) + downloadList.length - 1}))`);
      }

    }
    SpreadsheetApp.flush();
  });
  
}

// FUNCTIONS IN INTEROPERABILITY SECTION

/**
 * sheetConnect
 * Conecta hojas entre distintos documentos de Google Sheets.
 */
function sheetConnect(rowData) {

  let formData = [
    rowData[`sheetConnectSheetname`],
    rowData[`sheetConnectTargetURL`],
    rowData[`sheetConnectLinkList`]
  ];

  let [sheetConnectSheetname, sheetConnectTargetURL, sheetConnectLinkList] = formData;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  let ss_url = ss.getUrl();

  let sourceSheet = ss.getSheetByName(sheetConnectSheetname);
  var target = SpreadsheetApp.openById(getIdFromUrl(sheetConnectTargetURL));
  var targetSheets = target.getSheets();
  var targetSheet;
  
  var hojaEncontrada = false;
  
  for (var i = 0; i < targetSheets.length; i++) {
    if (targetSheets[i].getName() == sheetConnectSheetname) {
      hojaEncontrada = true;
      break;
    }
  }
  
  if (hojaEncontrada) {
    targetSheet = target.getSheetByName(sheetConnectSheetname);
  } else {
    targetSheet = sourceSheet.copyTo(target);
    targetSheet.setName(sheetConnectSheetname).setTabColor(sourceSheet.getTabColor());
    targetSheet.clearContents();
  }

  let targetSheetLink;

  if (sheetConnectLinkList == true) {
    targetSheetLink = target.getSheetByName('LINK');
 
    if (targetSheetLink.getRange('E2').getValue() != textMark) getEmptyLinkSheet();

    let sheetLink = '#gid=' + targetSheet.getSheetId();
    let lastRow = getLastDataRow(sh, 'E');

    let linkRange_1 = targetSheetLink.getRange(lastRow, 5, 1, 2);
    let linkRange_2 = targetSheetLink.getRange(lastRow, 5, 1, 1);

/*
    targetSheetLink.getRange(lastRow, 5, sh.getLastRow(), 2).clearFormat().setFontFamily('Montserrat').setFontSize(14).setFontWeight('normal').setHorizontalAlignment('left').setFontColor('#0000FF');
    targetSheetLink.getRange(1, 5, 2, 2).setBackground(null).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    targetSheetLink.getRange(1, 4, 1, 1).setBorder(true, true, true, true, true, true, '#26c6da', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    if (targetSheetLink.getRange('E2').getValue() != `Hojas conectadas`) {
      targetSheetLink.getRange(2, 5, 1, 2).setBackground('#fff').setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
        .setFontFamily('Inconsolata').setFontWeight('bold').setHorizontalAlignment('center');
    }
*/

    linkRange_1.setValues([[`=hyperlink("${sheetLink}";"${targetSheet.getName()}"& " / ${target.getName()}")`, `=hyperlink("${ss_url}";"${ss.getName()}")`]])
      .setBackground('#fafafa').setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    linkRange_2.setBackground('#fff').setFontColor('#78909c');

  }
  targetSheet.getRange('A1').setFormula(`=IMPORTRANGE("${ss_url}";"${sheetConnectSheetname}!A1:AZZ10000")`)
}

/**
 * getCSVFilesData
 * Import CSV data by sheet and range
 */
function getCSVFilesData(rowData, counter) {

  for (let i = 0; i <= counter; i++) {

    let formData = [
      rowData[`csvURL${i}`],
      rowData[`csvCELL${i}`],
      rowData[`selSht${i}`]
    ];

    let [csvURL, csvCELL, selSht] = formData;

    let sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(selSht);

    let fileURL = csvURL;
    let fileID = getIdFromUrl(fileURL);
    let file = DriveApp.getFileById(fileID);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    let fetchURL = `https://drive.google.com/uc?id=${fileID}&x=.csv`;

    let csvContent = UrlFetchApp.fetch(fetchURL);
    let csvData = Utilities.parseCsv(csvContent);

    SpreadsheetApp.flush();
    sh.getRange(sh.getRange(csvCELL).getRowIndex(), sh.getRange(csvCELL).getColumn(), csvData.length, csvData[0].length).setValues(csvData);
    SpreadsheetApp.flush();
  }
}

// FUNCTIONS IN HELP SECTION

/**
 * cellCounter
 * Cell counter for Google Sheets document
 */
let size;
function cellCounter() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  let cells = 0;
  sheets.forEach((sheet) => {
    cells = cells + sheet.getMaxRows() * sheet.getMaxColumns();
  });
  let division = cells / 10000000 * 100;
  let percentage = +division.toFixed(0);
  return percentage;
}

function cellCounter2() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  let cells = 0;
  sheets.forEach((sheet) => {
    cells = cells + sheet.getMaxRows() * sheet.getMaxColumns();
  });
  let division = cells / 10000000 * 100;
  let percentage = +division.toFixed(0);
  return (`📈 Cada Google Sheets tiene capacidad para diez millones de celdas. Has usado el <strong>${percentage}%</strong> del total con <strong>${cells} celdas</strong>.`);
}
