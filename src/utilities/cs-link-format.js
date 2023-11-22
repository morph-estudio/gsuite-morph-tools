/**
 * basicFormat
 * Global variables and Basic Formattings functions
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
 * Apply basic template formatting to the LINK sheet in Morph Tables
 */
function formatLinkSheet(updateBasic, updateAI, updateAC, ss) {

  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  var sheet_link = ss.getSheetByName('LINK') || ss.insertSheet('LINK', 0).setTabColor('#FFFF00');

  var { mainFontFamily, mainFontSize, mainFontColor, mainBorderColor } = formatVariables();

  if(updateBasic === "general") {
    sheet_link.getRange(1,1,sheet_link.getMaxRows(),sheet_link.getMaxColumns()).clear();
    sheet_link.clearNotes();
    basicListFormat(sheet_link, 0, 0);
    importedListFormat(sheet_link, 0, 3);
    connectedListFormat(sheet_link, 0, 6);
  }

  // Añadir bloque de archivos importados
  if(updateBasic) {
    basicListFormat(sheet_link, 0, 0);
  }
  // Añadir bloque de archivos importados
  if(updateAI) {
    importedListFormat(sheet_link, 0, 3)
  }
  // Añadir bloque de archivos conectados
  if(updateAC) {
    try {
      var colors = leerColoresColumnaAC(sheet_link, `H1:H${sheet_link.getLastRow()}`, mainFontColor) // Mantiene los colores en la columna de archivos conectados
    } catch (error) {
    }
    connectedListFormat(sheet_link, 0, 6);

    aplicarColoresColumnaAC(sheet_link, colors);
  }

  if(updateBasic != "general") {
  // Global Style
  sheet_link.getRange(1, 1, sheet_link.getMaxRows(), sheet_link.getMaxColumns()).setFontFamily(mainFontFamily).setFontSize(mainFontSize)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  // Row and Column Size
  var columnWidths = {
    "3": 28,
    "6": 28,
  };

  setColumnWidths(sheet_link, columnWidths);
  sheet_link.hideColumns(3);
  sheet_link.hideColumns(6);
  setCustomRowHeight(30, sheet_link);
  sheet_link.setRowHeight(1, 45);
  sheet_link.setFrozenRows(1);

  removeEmptyColumns(sheet_link);
  deleteEmptyRows(sheet_link);
}

/**
 * basicListFormat
 * LINK Sheet "Primary Data" Formatting
 */
function basicListFormat(sh, rowShift, colShift) {

  var { mainBorderColor } = formatVariables();

  let col1 = numToCol(1 + colShift);
  let col2 = numToCol(2 + colShift);
  let colArray = [col1, col2];
  let colArrayLength = colArray.length;

  if(sh.getMaxColumns() < colArrayLength) sh.insertColumns(sh.getMaxColumns(), colArrayLength); // If there are not enough columns, insert a new one

  sh.getRange(1 + rowShift, 1 + colShift, sh.getMaxRows(), 2).clearFormat();

  var textColumnA = [
    ['URL PANEL DE CONTROL'],
    ['CARPETA PANEL DE CONTROL'],
    ['ID CARPETA PANEL DE CONTROL'],
    ['CARPETA CUADRO'],
    ['ID CARPETA CUADRO'],
    ['CARPETA EXPORTACIONES'],
    ['CARPETA CARPETA CONGELADOS'],
    ['ÚLTIMO ARCHIVO CONGELADO'],
    ['DESCARGAR ARCHIVO XLSX']
  ];

  basicFormat(sh.getRange(1 + rowShift, 1 + colShift, sh.getMaxRows(), colArrayLength));

  sh.getRange(1 + rowShift, 1 + colShift, textColumnA.length, 1).setValues(textColumnA);

  // Specific Formatting

  // Col A
  sh.getRange(1 + rowShift, 1 + colShift, 9, 1).setFontWeight('bold');
  // Col A-B Border
  sh.getRange(1 + rowShift, 1 + colShift, 9, 2)
  .setBorder(true, true, null, true, true, true, mainBorderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  .setHorizontalAlignment('left');
  // Control Panel
  sh.getRange(1 + rowShift, 2 + colShift, 1, 1).setBackground('#ECFDF5').setFontWeight('bold').setFontColor('#00C853')
  .setBorder(true, true, null, true, true, true, '#00C853', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  // Folder Text Style
  sh.getRangeList([`${col2}2`, `${col2}4`, `${col2}6`, `${col2}7`]).setFontWeight('bold').setFontColor('#0000FF');
  sh.getRange(`${col2}9`).setFontColor('#0000FF');
  sh.getRangeList([`${col1}1:${col1}`, `${col1}1:${col2}1`]).setFontWeight('bold');

  var columnWidths = {
    [1 + colShift]: 340,
    [2 + colShift]: 340,
  };

  setColumnWidths(sh, columnWidths);
}

/**
 * basicListFormat
 * LINK Sheet "Imported Files Block" Formatting
 */
function importedListFormat(sh, rowShift, colShift) {

  var { mainBorderColor, mainFontColor } = formatVariables();

  sh.getRange(1 + rowShift, 1 + colShift, sh.getLastRow(), 2).clear().setFontColor(mainFontColor);

  let col1 = numToCol(1 + colShift);
  let col2 = numToCol(2 + colShift);
  let colArray = [col1, col2];
  let colArrayLength = colArray.length;

  if(sh.getMaxColumns() < colArrayLength) sh.insertColumns(sh.getMaxColumns(), colArrayLength); // If there are not enough columns, insert a new one

  let prefix_one = 'AI';

  let textHeadersAI = [[
    `Archivos Importados (${prefix_one})`, 
    `${prefix_one}: IDs`
  ]];

  sh.getRange(1, 1 + colShift, 1, textHeadersAI[0].length).setValues(textHeadersAI);

  var columnWidths = {
    [1 + colShift]: 340,
    [2 + colShift]: 340,
  };

  setColumnWidths(sh, columnWidths);

  basicFormat(sh.getRange(1 + rowShift, 1 + colShift, sh.getMaxRows(), colArrayLength));

  // Specific Formatting

  sh.getRange(`${col1}:${col1}`).setFontWeight('bold');
  sh.getRange(1 + rowShift, 1 + colShift, 1, colArrayLength)
    .setHorizontalAlignment('center')
    .setFontWeight('bold');
  sh.getRange(1 + rowShift, 1 + colShift, getLastDataRow(sh, col1), colArrayLength).setBorder(true, true, true, true, true, true, mainBorderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  return textHeadersAI;
}

/**
 * connectedListFormat
 * LINK Sheet "Connected Files" Formatting
 */
function connectedListFormat(sh, rowShift, colShift) {

  var { mainBorderColor, mainFontColor } = formatVariables();

  sh.getRange(1 + rowShift, 1 + colShift, sh.getLastRow(), 7).clear().setFontColor(mainFontColor);

  let col1 = numToCol(1 + colShift);
  let col2 = numToCol(2 + colShift);
  let col3 = numToCol(3 + colShift);
  let col4 = numToCol(4 + colShift);
  let col5 = numToCol(5 + colShift);
  let col6 = numToCol(6 + colShift);
  let col7 = numToCol(7 + colShift);

  let colArray = [col1, col2, col3, col4, col5, col6, col7];
  let colArrayLength = colArray.length;

  if(sh.getMaxColumns() < colArrayLength) sh.insertColumns(sh.getMaxColumns(), colArrayLength); // If there are not enough columns, insert a new one

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

  sh.getRange(1, 1 + colShift, 1, textHeadersAC.length).setValues(textHeadersACpaste);

  var columnWidths = {
    [1 + colShift]: 400,
    [2 + colShift]: 440,
    [3 + colShift]: 225,
    [4 + colShift]: 225,
    [5 + colShift]: 85,
    [6 + colShift]: 150,
    [7 + colShift]: 35
  };

  setColumnWidths(sh, columnWidths);

  if (sh.getRange('A1').getValue() === textHeadersAC[0]) setCustomRowHeight(30, sh); // Si la lista se aplica en una hoja nueva, modificar la altura de las filas

  basicFormat(sh.getRange(1 + rowShift, 1 + colShift, sh.getMaxRows(), colArrayLength));

  // Specific Formatting
  sh.getRange(`${col1}:${col1}`).setFontWeight('bold');
  sh.getRange(1 + rowShift, 1 + colShift, 1, 7).setHorizontalAlignment('center').setFontWeight('bold');
  sh.getRange(1 + rowShift, 1 + colShift, getLastDataRow(sh, col4), colArrayLength).setBorder(true, true, true, true, true, true, mainBorderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(`${col5}:${col5}`).setHorizontalAlignment('right');
  sh.getRange(`${col6}:${col6}`).setHorizontalAlignment('left');
  sh.getRange(`${col7}:${col7}`).setHorizontalAlignment('center');
  sh.getRange(`${col7}:${col7}`).setHorizontalAlignment('center');
  sh.getRange(`${col5}1:${col7}1`).setFontSize(10);

  createCollapseGroup(sh, 5 + colShift, 1, `${col5}:${col6}`);
  sh.setColumnGroupControlPosition(SpreadsheetApp.GroupControlTogglePosition.AFTER);

  return [textHeadersAC, colArray];
}

/**
 * leerColoresColumnaAC
 * Save "Connected Files" column colors
 */
function leerColoresColumnaAC(sh, rangeReference, baseColor) {
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

/**
 * aplicarColoresColumnaAC
 * Apply saved "Connected Files" column colors
 */
function aplicarColoresColumnaAC(sh, colors) {
  for (var cell in colors) {
    sh.getRange(cell).setFontColor(colors[cell]);
  }
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
