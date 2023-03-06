/**
 * getConectedSheetList
 * Genera una lista de las hojas conectadas con ImportRange en el documento
 */
function getConectedSheetList(rowShift, columnShift, sheetName) {

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

  let [textHeadersAC, colArray] = connectedListFormat(sh_link, rowShift, columnShift);
  list = [textHeadersAC.slice(0, -2), ...list];
  let connectedSheetsListColumns = textHeadersAC.length;

  let lastRow = sh_link.getLastRow();
  let emptyChecker = checkIfSheetIsEmpty(sh_link); Logger.log(emptyChecker)

  if (emptyChecker != true) {
    sh_link.getRange(2 + rowShift, 1 + columnShift, lastRow - rowShift, connectedSheetsListColumns).clearContent();
    sh_link.getRange(1 + rowShift, 1 + columnShift, lastRow - rowShift, connectedSheetsListColumns)
      .setBorder(false, false, false, false, false, false, 'black', SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor('#607D8B');
  }

  let listRange = sh_link.getRange(1 + rowShift, 1 + columnShift, list.length, 5);

  listRange.setValues(list);
  deleteDuplicatedListObjects(sh_link, listRange, rowShift, columnShift);

  if(importedSheets.length > 0) {
    let n;
    for (var i = 0; i < importedSheets.length; i++) {
      n = i + 2;
      sh_link.getRange(rowShift + n, 6 + columnShift)
      .setFormula(`=IF(${numToCol(1 + columnShift)}${n}<>"";IMPORTRANGE(CHAR(34)&${colArray[1]}${n}&CHAR(34);CHAR(39)&${colArray[2]}${n}&"'!"&${colArray[4]}${n});)`);
    }
    sh_link.getRange(2 + rowShift, connectedSheetsListColumns + columnShift)
      .setFormula(`=ARRAYFORMULA(IF(${colArray[0]}${2 + rowShift}:${colArray[0]}<>"";IF(ISERROR(${colArray[5]}${2 + rowShift}:${colArray[5]});"🟥";"🟩");""))`)
  }

  // List Format

  sh_link.getRange(1 + rowShift, 1 + columnShift, list.length, connectedSheetsListColumns).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  
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

function deleteDuplicatedListObjects(sh, dataRange, rowShift, columnShift) {
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
      sh.getRange(i + rowShift + 1, 1 + columnShift, 1, 1).clearContent(); // Borrar la columna "Archivos duplicados"
      sh.getRange(i + rowShift + 1, 2 + columnShift, 1, 1).setFontColor('#EFEFEF'); // Cambiar el color de la columna "Acción"
    }
  }
}

/**
 * basicFormat
 * Aplica una plantilla base a la hoja LINK de los cuadros Morph
 */

function formatVariables() {

  var rowData = {
    mainFontFamily: 'Inter',
    mainFontSize: 14,
    mainFontColor: '#607d8b',
    mainBorderColor: '#b0bec5',
    hyperlinkFontColor: '#0000FF'
  };

  return rowData;
}

function basicFormat(rango) {

  var { mainFontFamily, mainFontSize, mainFontColor } = formatVariables();

  rango
    .setFontFamily(mainFontFamily)
    .setFontSize(mainFontSize)
    .setFontWeight('normal')
    .setFontColor(mainFontColor)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
}

/**
 * formatLinkSheet
 * Aplica una plantilla base a la hoja LINK de los cuadros Morph
 */
function formatLinkSheet(ss) {

  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheet_link = ss.getSheetByName('LINK') || ss.insertSheet('LINK', 0).setTabColor('#FFFF00');

  var { mainFontFamily, mainFontSize, mainFontColor, mainBorderColor } = formatVariables();

  var colors = leerColoresColumna(sheet_link, `H1:H${sheet_link.getLastRow()}`, mainFontColor) // Mantiene los colores en la columna de archivos conectados
    Logger.log(`colors: ${JSON.stringify(colors)}`);

  sheet_link.getDataRange().clearFormat();

  var textColumnA = [
    ['URL PANEL DE CONTROL'],
    ['CARPETA PANEL DE CONTROL'],
    ['ID CARPETA PANEL DE CONTROL'],
    ['CARPETA CUADRO'],
    ['ID CARPETA CUADRO'],
    ['CARPETA EXPORTACIONES'],
    ['CARPETA BACKUP'],
    ['ID CARPETA BACKUP'],
    ['DESCARGAR ARCHIVO XLSX']
  ];

  sheet_link.getRange(1, 1, textColumnA.length, 1).setValues(textColumnA);

  // Añadir bloque de archivos importados
  importedListFormat(sheet_link, 0, 3)

  var maxRows = sheet_link.getMaxRows();
  var maxColumns = sheet_link.getMaxColumns();

  // Global Style
  sheet_link.getRange(1, 1, maxRows, maxColumns).setFontFamily(mainFontFamily).setFontSize(mainFontSize).setFontColor(mainFontColor)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // Col A-B Border
  sheet_link.getRange(1, 1, 9, 2)
  .setBorder(true, true, true, true, true, true, mainBorderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setHorizontalAlignment('left');

  // Col A
  sheet_link.getRange(1, 1, 9, 1).setFontWeight('bold')

  // Control Panel
  sheet_link.getRange('B1').setBackground('#ECFDF5').setFontWeight('bold').setFontColor('#00C853')
    .setBorder(true, true, true, true, true, true, '#00C853', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Folder Text Style
  sheet_link.getRangeList(['B2', 'B4', 'B6', 'B7']).setFontWeight('bold').setFontColor('#0000FF');
  sheet_link.getRange('B9').setFontColor('#0000FF');
  sheet_link.getRangeList(['A1:A', 'A1:B1']).setFontWeight('bold');

  // Añadir bloque de archivos conectados
  connectedListFormat(sheet_link, 0, 6)
  aplicarColoresGuardados(sheet_link, colors);

  // Rows and Columns Size
  const columnWidths = {
    "1": 340,
    "2": 340,
    "3": 28,
    "6": 28,
  };

  setColumnWidths(sheet_link, columnWidths);
  setCustomRowHeight(30, sheet_link);
  sheet_link.setRowHeight(1, 45);
  sheet_link.setFrozenRows(1);

  sheet_link.hideColumns(3);
  sheet_link.hideColumns(6);
  
  // Remove Empty Rows and Columns

  //let deleteRowIndex = getLastDataRowIndex(sheet_link); Logger.log(`deleteRowIndex: ${deleteRowIndex}`);
  //if (maxRows > deleteRowIndex) sheet_link.deleteRows(deleteRowIndex + 1, sheet_link.getMaxRows() - deleteRowIndex);
  removeEmptyColumns(sheet_link);
}


/**
 * importedListFormat
 * Aplica formato al bloque de Archivos Importados
 */
function importedListFormat(sh, rowShift, columnShift) {

  var { mainBorderColor } = formatVariables();

  let col1 = numToCol(1 + columnShift);
  let col2 = numToCol(2 + columnShift);
  let colArray = [col1, col2];
  let colArrayLength = colArray.length;

  let prefix_one = 'AI';

  let textHeadersAI = [[
    `Archivos Importados (${prefix_one})`, 
    `${prefix_one}: IDs`
  ]];

  sh.getRange(1, 1 + columnShift, 1, textHeadersAI[0].length).setValues(textHeadersAI);

  var columnWidths = {
    [1 + columnShift]: 340,
    [2 + columnShift]: 340,
  };

  setColumnWidths(sh, columnWidths);

  basicFormat(sh.getRange(1 + rowShift, 1 + columnShift, sh.getMaxRows(), colArrayLength));

  sh.getRange(`${col1}:${col1}`).setFontWeight('bold');
  
  sh.getRange(1 + rowShift, 1 + columnShift, 1, colArrayLength)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');

  sh.getRange(1 + rowShift, 1 + columnShift, getLastDataRow(sh, col1), colArrayLength).setBorder(true, true, true, true, true, true, mainBorderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  return textHeadersAI;
}

/**
 * connectedListFormat
 * Aplica formato al bloque de Archivos Conectados
 */
function connectedListFormat(sh, rowShift, columnShift) {

  var { mainBorderColor } = formatVariables();

  let col1 = numToCol(1 + columnShift);
  let col2 = numToCol(2 + columnShift);
  let col3 = numToCol(3 + columnShift);
  let col4 = numToCol(4 + columnShift);
  let col5 = numToCol(5 + columnShift);
  let col6 = numToCol(6 + columnShift);
  let col7 = numToCol(7 + columnShift);

  let colArray = [col1, col2, col3, col4, col5, col6, col7];
  let colArrayLength = colArray.length;

  let prefix_two = 'AC';

  let textHeadersAC = [
    `Archivos conectados (${prefix_two})`,
    `${prefix_two}: URL`,
    `${prefix_two}: Hoja origen`,
    `${prefix_two}: Hoja destino`,
    `${prefix_two}: Cell`,
    `${prefix_two}: Test`,
    `⬜`
  ];

  let textHeadersACpaste = [textHeadersAC]

  sh.getRange(1, 1 + columnShift, 1, textHeadersAC.length).setValues(textHeadersACpaste);

  var columnWidths = {
    [1 + columnShift]: 400,
    [2 + columnShift]: 440,
    [3 + columnShift]: 225,
    [4 + columnShift]: 225,
    [5 + columnShift]: 75,
    [6 + columnShift]: 150,
    [7 + columnShift]: 28
  };

  setColumnWidths(sh, columnWidths);

  if (sh.getRange('A1').getValue() === textHeadersAC[0]) setCustomRowHeight(30, sh); // Si la lista se aplica en una hoja nueva, modificar la altura de las filas

  basicFormat(sh.getRange(1 + rowShift, 1 + columnShift, sh.getMaxRows(), colArrayLength));

  sh.getRange(`${col1}:${col1}`).setFontWeight('bold');
  sh.getRange(1 + rowShift, 1 + columnShift, 1, 7).setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange(1 + rowShift, 1 + columnShift, getLastDataRow(sh, col4), colArrayLength).setBorder(true, true, true, true, true, true, mainBorderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(`${col5}:${col5}`).setHorizontalAlignment('right');
  sh.getRange(`${col6}:${col6}`).setHorizontalAlignment('left');
  sh.getRange(`${col7}:${col7}`).setHorizontalAlignment('center');
  sh.getRange(`${col5}1:${col7}1`).setFontSize(10);

  createCollapseGroup(sh, 5 + columnShift, 1, `${col5}:${col6}`);
  sh.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);

  return [textHeadersAC, colArray];
}

function cambiarColor(sh, range, baseColor) {
  var rango = sh.getRange(range);
  var valores = rango.getValues();
  var celdasAfectadas = [];
  var targetColor;

  for (var i = 0; i < valores.length; i++) {
    var celda = rango.getCell(i + 1, 1);
    var color = celda.getFontColor();

    if (color != baseColor) {
      if (targetColor != undefined) targetColor = color;
      celdasAfectadas.push(celda);
    }
  }

  for (var i = 0; i < celdasAfectadas.length; i++) {
    celdasAfectadas[i].setFontColor(targetColor);
  }
}

function leerColoresColumna(sh, rangeReference, baseColor) {

  Logger.log(`rangeReference: ${rangeReference}`);

  var range = sh.getRange(rangeReference);
  var data = range.getValues();
  var fontColors = range.getFontColors();
  var colors = {};
  for (var i = 0; i < data.length; i++) {
    if (fontColors[i][0] != baseColor) {
      colors[range.getCell(i+1,1).getA1Notation()] = fontColors[i][0];
    }
  }
  return colors;
}

function aplicarColoresGuardados(sh, colors) {
  for (var cell in colors) {
    sh.getRange(cell).setFontColor(colors[cell]);
  }
}






























/**
 * formatLinkSheetOld
 * Aplica una plantilla base a la hoja LINK de los cuadros Morph
 */
function formatLinkSheetOld(ss) {

  ss = ss || SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('LINK') || ss.insertSheet('LINK', 0).setTabColor('#FFFF00');
  sh.getDataRange().clearFormat();

  // Headers

  let prefix_one = 'Archivos importados'; `${prefix_one}`

  let result = [[`Carpetas referentes`, `${prefix_one}`, `${prefix_one}: IDs`]];
  sh.getRange(2, 2, 1, result[0].length).setValues(result);

  // Column A Titles

  var textColumnA = [['URL PANEL DE CONTROL'], [''], ['CARPETA PANEL DE CONTROL'], ['ID CARPETA PANEL DE CONTROL'], ['CARPETA CUADRO'], ['ID CARPETA CUADRO'], ['CARPETA EXPORTACIONES'], ['CARPETA BACKUP'], ['ID CARPETA BACKUP'], ['DESCARGAR ARCHIVO XLSX']];

  sh.getRange(1, 1, textColumnA.length, 1).setValues(textColumnA);

  let maxRows = sh.getMaxRows();
  let maxColumns = sh.getMaxColumns();

  // Global Style
  sh.getRange(1, 1, maxRows, maxColumns).setFontFamily('Montserrat').setFontSize(14).setFontWeight('normal').setFontColor('#607D8B')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  // First Row Titles
  sh.getRange(1, 3, 1, maxColumns - 2).setBorder(null, null, null, null, true, null, '#fff', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Second Row Titles
  sh.getRange(2, 1, 1, 4).setFontFamily('Inconsolata')
    .setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setHorizontalAlignment('center');

  // Col A
  sh.getRange(1, 1, 10, 1).setFontWeight('bold')
    .setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Col B
  sh.getRange(1, 2, 10, 1).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Archivos importados List
  sh.getRange(2, 3, getLastDataRow(sh,"C") - 1, 2).setBorder(true, true, true, true, true, true, '#b0bec5', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  //sh.getRange(1, 3, 1, 2).setFontColor('#64DD17');

  // ImportRange
  sh.getRange(1, 3, 1, 2).setBackground('#e0f7fa').setFontWeight('bold').setBorder(true, true, true, true, true, true, '#26c6da', SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor('#26c6da')
    .setHorizontalAlignment('center');

  // Control Panel
  sh.getRange('B1').setBackground('#ECFDF5').setFontWeight('bold').setFontColor('#00C853')
    .setBorder(true, true, true, true, true, true, '#00C853', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Folder Text Style
  sh.getRangeList(['B3', 'B5', 'B7', 'B8']).setFontWeight('bold').setFontColor('#0000FF');
  sh.getRange('B10').setFontColor('#0000FF');
  sh.getRangeList(['A1:A', 'A2:H2', 'C2:C']).setFontWeight('bold');

  // Rows and Columns Size
  const columnWidths = {
    "1": 340, "2": 340, "3:4": 340, "5": 30
  };

  setColumnWidths(sh, columnWidths);
  setCustomRowHeight(30, sh);
  sh.setRowHeight(1, 35);
  sh.setRowHeight(2, 50);

  // Remove Empty Rows and Columns
  deleteEmptyRows(sh); removeEmptyColumns(sh);
}
