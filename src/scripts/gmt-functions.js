// CATEGORY: Named Range Utilities

/**
 * Elimina todos los rangos nombrados de una hoja de cálculo de Google Sheets
 */
function deleteNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  var namedRanges = ss.getNamedRanges();
  for (var i = 0; i < namedRanges.length; i++) {
    namedRanges[i].remove();
  }
}

/**
 * Elimina todos los rangos nombrados que apuntan a un intervalo #REF! en una hoja de cálculo de Google Sheets.
 */
function deleteInvalidNamedRanges() {
  const ss = SpreadsheetApp.getActive();
  const namedRanges = ss.getNamedRanges();
  const namesR = namedRanges.map(nr=>nr.getName());   
  namesR.forEach(name=>{
      ss.removeNamedRange(name);
  });
}

/**
 * Actualiza los rangos nombrados en la hoja de cálculo basados en una plantilla Morph.
 */
function refreshNamedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssID = ss.getId();

  var shVarName = 'X Variables';
  var shOpeName = 'TXT OPERACIONES';
  var shVarColumnIndexName = 'VA1340';

  var shVar = ss.getSheetByName(shVarName);
  var shOpe = ss.getSheetByName(shOpeName);

  if (!shVar || !shOpe) {
    throw new Error(`No se encontraron las hojas ${shOpeName} o ${shVarName}.`);
  }

  // Filtrar encabezados y valores de la plantilla de rangos nombrados

  var shOpeHeaders = flattenArray(shOpe.getRange(1, 1, 1, shOpe.getLastColumn()).getValues());
  Logger.log(`Headers: ${shOpeHeaders}`);

  var shVarHeaders = shVar.getRange(1, 1, 1, shVar.getLastColumn()).getValues()[0];
  var shVarColumnIndex = shVarHeaders.indexOf(shVarColumnIndexName);

  if (shVarColumnIndex !== -1) {
    var columnIDRange = shVar.getRange(2, shVarColumnIndex + 1, shVar.getMaxRows() - 1, 1).getValues();
    var columnArchiRange = shVar.getRange(2, shVarColumnIndex + 2, shVar.getMaxRows() - 1, 1).getValues();
    var columnTXTNameRange = shVar.getRange(2, shVarColumnIndex + 3, shVar.getMaxRows() - 1, 1).getValues();
    var columnAltIDRange = shVar.getRange(2, shVarColumnIndex + 4, shVar.getMaxRows() - 1, 1).getValues();
    var columnAltArchiRange = shVar.getRange(2, shVarColumnIndex + 5, shVar.getMaxRows() - 1, 1).getValues();

    var filteredColumnArchiRange = [];
    var filteredColumnTXTNameRange = [];
    var filteredColumnIDRange = [];

    var altCodesObject = {};

    for (var i = 0; i < columnAltIDRange.length; i++) {
      var key = columnAltIDRange[i][0].substring(3); // Obtenemos el valor de columnAltIDRange en la fila i
      var value = columnAltArchiRange[i][0]; // Obtenemos el valor de columnAltArchiRange en la misma fila
      altCodesObject[key] = value; // Asignamos el valor al objeto usando el valor de columnAltIDRange como clave
    }

    Logger.log(`altCodesObject: ${JSON.stringify(altCodesObject)}`)

    for (var i = 0; i < columnTXTNameRange.length; i++) {
      var valueB = columnTXTNameRange[i][0];
      if (valueB) {
        filteredColumnArchiRange.push(columnArchiRange[i][0]);
        filteredColumnIDRange.push(columnIDRange[i][0]);
        filteredColumnTXTNameRange.push(valueB);
      }
    }
  }

  Logger.log(`filteredColumnArchiRange: ${filteredColumnArchiRange}, filteredColumnTXTNameRange: ${filteredColumnTXTNameRange}`)

  // Get Named Ranges using Google Sheets API

  var response = Sheets.Spreadsheets.get(ssID);
  var namedRanges = response.namedRanges; Logger.log(`namedRanges: ${namedRanges}`);
  var namedRangesDict = {};

  // Verificar si namedRanges está vacío o es nulo

  if (namedRanges || namedRanges != undefined) {
    for (var i = 0; i < namedRanges.length; i++) {
      namedRangesDict[namedRanges[i].name] = namedRanges[i];
    }
  }

  var batchRequests = [];
  var headerFound;
  var columnIndex;

  for (var i = 0; i < filteredColumnTXTNameRange.length; i++) {
    SpreadsheetApp.flush();
    var namedRangeName = filteredColumnTXTNameRange[i].toString().trim();
    var idHeader = filteredColumnIDRange[i].toString().trim();
    var txtHeader = filteredColumnArchiRange[i].toString().trim();
    var columnIndex = shOpeHeaders.indexOf(txtHeader);

    if (columnIndex == -1) {
      var altCodesObjectFiltered = {};
      for (var key in altCodesObject) {
        if (key === idHeader.substring(3)) { // Quita los tres primeros caracteres y compara
          altCodesObjectFiltered[key] = altCodesObject[key];
        }
      }
      
      var found = false;
      for (var key in altCodesObjectFiltered) {
        columnIndex = shOpeHeaders.indexOf(altCodesObjectFiltered[key]);
        if (columnIndex !== -1) {
          found = true;
          break;
        }
      }

      Logger.log(`idHeader: ${idHeader}, txtHeader: ${txtHeader}, altCodesObjectFiltered: ${JSON.stringify(altCodesObjectFiltered)}, found: ${found}`)
      
      if (found) {
        headerFound = true;
      } else {
        headerFound = false;
      }
    } else {
      headerFound = true;
    }

    if (headerFound === false) {
      continue;
    } else {
      var namedRange = {
        range: {
          sheetId: shOpe.getSheetId(),
          startRowIndex: 1,
          endRowIndex: 7000,
          startColumnIndex: columnIndex,
          endColumnIndex: columnIndex + 1
        }
      };

      var existingNamedRange = namedRangesDict[namedRangeName];

      if (existingNamedRange != undefined) {

        // Update existing Named Range
        batchRequests.push({
          updateNamedRange: {
            namedRange: {
              name: namedRangeName.trim(),
              namedRangeId: existingNamedRange.namedRangeId,
              range: namedRange.range
            },
            fields: 'range'
          }
        });
      } else {

        // Add new Named Range
        batchRequests.push({
          addNamedRange: {
            namedRange: {
              name: namedRangeName,
              range: namedRange.range
            }
          }
        });
      }
    }
  }

  Logger.log(`batchRequests: ${JSON.stringify(batchRequests)}`);

  // Send batch update request

  SpreadsheetApp.flush();

  Sheets.Spreadsheets.batchUpdate({
    requests: batchRequests
  }, ssID);
}

// CATEGORY: Main Tools & Utilities

/**
 * Elimina las filas vacías en una hoja de cálculo hasta la última fila que contiene datos.
 *
 * @param {Sheet} sh - La hoja de cálculo en la que se eliminarán las filas vacías (opcional, se utiliza la hoja activa por defecto).
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

/**
 * Elimina las columnas vacías en una hoja de cálculo hasta la última columna que contiene datos.
 *
 * @param {Sheet} sh - La hoja de cálculo en la que se eliminarán las columnas vacías (opcional, se utiliza la hoja activa por defecto).
 */
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
 * Obtiene una lista de usuarios del servidor Morph.
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
 * Incrementa el tamaño de fuente proporcionalmente en toda la hoja.
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

/**
 * Reduce el tamaño de fuente proporcionalmente en toda la hoja.
 */
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

/**
 * Genera identificadores únicos en las celdas seleccionadas.
 */
function uniqueIdentifier() {
  var selection = sh().getActiveRange();
  var columns = selection.getNumColumns();
  var rows = selection.getNumRows();
  for (let column = 1; column <= columns; column++) {
    for (let row = 1; row <= rows; row++) {
      var cell = selection.getCell(row, column);
      cell.setValue(Utilities.getUuid());
    }
  }
}

// CATEGORY: Files and Folders Section

/**
 * Genera una lista de archivos dentro de una carpeta de Google Drive.
 *
 * @param {object} rowData - Los datos de la fila que contiene información sobre la carpeta.
 * @param {string} rowData.listFolderID - El ID de la carpeta de Google Drive o la URL de la celda A1 que contiene la URL de la carpeta.
 * @param {boolean} rowData.useA1 - Indica si se debe usar la celda A1 para obtener la URL de la carpeta.
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

/**
 * Reduce el nombre del tipo de archivo MIME a un valor corto.
 *
 * @param {string} mimetype - El tipo de archivo MIME a reducir.
 * @return {string} - El valor corto del tipo de archivo MIME.
 */
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
 * Inserta las imágenes de una carpeta en las celdas seleccionadas.
 *
 * @param {object} rowData - Los datos de la fila que contiene información sobre la carpeta y la configuración de la inserción.
 * @param {string} rowData.listFolderID - El ID de la carpeta de Google Drive o la URL de la celda A1 que contiene la URL de la carpeta.
 * @param {boolean} rowData.useA1 - Indica si se debe usar la celda A1 para obtener la URL de la carpeta.
 * @param {boolean} rowData.imageFolderFileID - Indica si se debe incluir el ID de archivo en la lista de datos.
 * @param {boolean} rowData.imageFolderFileName - Indica si se debe incluir el nombre de archivo en la lista de datos.
 * @param {boolean} rowData.imageFolderImage - Indica si se deben insertar las imágenes en las celdas.
 * @param {boolean} rowData.imageFolderArrayFormula - Indica si se debe aplicar una fórmula de matriz para mostrar las imágenes.
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

// CATEGORÍA: Funciones de la sección de Interoperabilidad

/**
 * Conecta hojas entre distintos documentos de Google Sheets.
 *
 * @param {object} rowData - Los datos de la fila que contiene información sobre la conexión de hojas.
 * @param {string} rowData.sheetConnectSheetname - El nombre de la hoja que se va a conectar.
 * @param {string} rowData.sheetConnectTargetURL - La URL del documento de Google Sheets de destino.
 * @param {boolean} rowData.sheetConnectLinkList - Indica si se va a vincular la hoja de forma predeterminada en la lista de hojas.
 */
function vincularHojasImportrange(rowData) {

  let formData = [
    rowData[`sheetConnectSheetname`],
    rowData[`sheetConnectTargetURL`],
    rowData[`sheetConnectLinkList`]
  ];

  let [sheetConnectSheetname, sheetConnectTargetURL, sheetConnectLinkList] = formData;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ss_url = ss.getUrl();

  let sourceSheet = ss.getSheetByName(sheetConnectSheetname);
  if(sheetConnectLinkList){
    var target = ss;
  } else {
    var target = SpreadsheetApp.openById(getIdFromUrl(sheetConnectTargetURL));
  }
  
  targetSheet = sourceSheet.copyTo(target);
  targetSheet.setName(`X ${sheetConnectSheetname}`).setTabColor('#00ff00');
  targetSheet.clearContents();

  targetSheet.getRange('A1').setFormula(`=IMPORTRANGE("${ss_url}";"${sheetConnectSheetname}!A1:AZZ10000")`)
}

/**
 * Importa datos CSV por hoja y rango.
 *
 * @param {object} rowData - Los datos de la fila que contiene información sobre la importación de datos CSV.
 * @param {number} counter - El contador que especifica la cantidad de datos CSV a importar.
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

// CATEGORY: Functions in Help Section

/**
 * Contador de celdas para el documento de Google Sheets.
 *
 * @return {number} - El porcentaje de celdas utilizadas en relación con el límite de 10 millones de celdas.
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

/**
 * Contador de celdas con mensaje descriptivo.
 *
 * @return {string} - Un mensaje que indica el porcentaje de celdas utilizadas y la cantidad total de celdas.
 */
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

/**
 * Elimina filas vacías hasta la última fila de datos en una hoja especificada.
 *
 * @param {Sheet} sh - La hoja en la que se eliminarán las filas vacías.
 * @return {number} - El número de la última fila de datos después de la eliminación.
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
 * Optimización de la hoja de cálculo eliminando filas excesivas.
 *
 * @param {number} rowsInput - El número máximo de filas permitidas en cada hoja.
 * @throws {Error} - Si la hoja ya tiene menos de 'rowsInput' filas.
 */
function deleteExcessiveRows(rowsInput, sh) {
  sh = sh || SpreadsheetApp.getActiveSheet();
  let maxRows = sh.getMaxRows();
  if (maxRows < rowsInput) throw new Error(`La hoja ya tiene menos de ${rowsInput} filas.`)
  sh.deleteRows(rowsInput, maxRows - rowsInput);
}

/**
 * Optimiza el documento de Google Sheets eliminando filas excesivas en todas las hojas que tienen más de 'rowsInput' filas.
 *
 * @param {number} rowsInput - El número máximo de filas permitidas en cada hoja.
 * @throws {Error} - Si no hay ninguna hoja con más de 'rowsInput' filas o si el usuario cancela la optimización.
 */
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
 * Borra la memoria caché del documento de hojas de cálculo.
 */
 function purgeDocumentCache() {
  SpreadsheetApp.flush()
}
