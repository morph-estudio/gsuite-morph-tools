/**
 * formulaDatabaseImport
 * Importa las fórmulas de la base de datos Morph al documento actual
 */
function formulaDatabaseImport(rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  var { formulaInteropSelectFileType, formulaInteropAllDocument, formulaInteropFormat } = rowData;

  var finalFolder = formulaDatabaseProjectFolder(ss, formulaInteropSelectFileType, 'import');
  var filesList = { files: [] };

  if (!formulaInteropAllDocument) {
    let searchFor = `${ss.getActiveSheet().getName()}.gse`; // Elimina la extensión de archivo del nombre de la hoja activa
    Logger.log(`Buscando archivo con título '${searchFor}' en carpeta ${finalFolder.getName()} (${finalFolder.getUrl()}) ...`);
    let filesIterator = finalFolder.searchFiles("title='" + searchFor + "'");
    if (filesIterator.hasNext()) {
      let file = filesIterator.next();
      let fileName = file.getName().split('.')[0].trim();
      let contenido = file.getBlob().getDataAsString().replace(/\r\n|\r|\n/g, "\n");
      let filesData = { name: fileName, content: contenido };
      filesList.files.push(filesData);
      Logger.log(`Se encontró un archivo con título '${file.getName()}' en la carpeta asociada al documento.`);
    } else {
      throw new Error(`No se ha encontrado ningún archivo en la carpeta '${finalFolder.getName()}' correspondiente a la hoja seleccionada.`);
    }
  } else {
    let files = finalFolder.getFiles();
    while (files.hasNext()) {
      let file = files.next();
      let fileName = file.getName().split('.')[0].trim();
      let contenido = file.getBlob().getDataAsString().replace(/\r\n|\r|\n/g, "\n");
      let filesData = { name: fileName, content: contenido };
      filesList.files.push(filesData);
    }
  }

  var sheet = []
  var cells = [];
  var formulas = [];
  var formulaList = {};

  filesList.files.forEach(function(file) {
    formulaList = {};
    sheet.push(file.name)

    var lines = file.content.split('\n');

    for (var i = 0; i < lines.length; i++) {
      var line = lines[i];
      if (!line.trim()) continue; // Saltar a la siguiente iteración si la línea está vacía
      if (line.startsWith('// CELL=')) {
        // Extraer la celda y guardarla en el array de celdas
        var cell = line.split('=')[1].trim();
        cells.push(cell);
      
        // Buscar la fórmula de la celda y guardarla en el array de fórmulas
        var formula = '';
        var inComment = false;

        for (var j = i + 1; j < lines.length; j++) {
          var nextLine = lines[j];

          if (nextLine.startsWith('// CELL=')) {
              break; // Se ha encontrado otra celda, salir del bucle
            } else if (nextLine.trim().startsWith('/*')) {
              inComment = true;
            } else if (nextLine.trim().startsWith('*/')) {
              inComment = false;
            } else if (nextLine.trim().startsWith('//')) {
              continue;
            } else if (nextLine.trim().length < 1) {
              continue;
            } else if (!inComment && !nextLine.startsWith('/*')) {
              // No estamos dentro de un comentario y la línea no empieza por /*
              // Añadir la línea al array de fórmulas
              formula += nextLine + '\n';
            }
        }

        var finalFormula;
        if (formulaInteropFormat) {
          finalFormula = cleanFormulaFormat(formula.substring(1).slice(0, -1));
        } else {
          finalFormula = formula.substring(1).slice(0, -1);
        }

        formulas.push(finalFormula);
      }
    }

    formulaList = { cell: cells, formula: formulas };

    var sh = ss.getSheetByName(file.name);

    for (var i = 0; i < formulaList.cell.length; i++) {
      sh.getRange(formulaList.cell[i]).clearContent();
    }

    SpreadsheetApp.flush();

    for (var i = 0; i < formulaList.cell.length; i++) {
      sh.getRange(formulaList.cell[i]).setFormula(formulaList.formula[i])
    }

    // Limpiar los arrays para el siguiente archivo
    cells = [];
    formulas = [];
  });

}

/**
 * formulaDatabaseProjectFolder
 * Encuentra la carpeta del documento actual en la base de datos de fórmulas
 */
function formulaDatabaseProjectFolder(ss, formulaInteropSelectFileType, buttonClicked) {

  var ss_name = ss.getName();
  
  const cache = CacheService.getScriptCache();
  const cacheKey = `folder-${ss_name}`;

  const cachedFolder = cache.get(cacheKey);
  if (cachedFolder !== null) {
    var finalFolder = DriveApp.getFolderById(cachedFolder);
    Logger.log(`Aviso: se ha recuperado la carpeta del documento de la memoria caché.`)
    return finalFolder;
  }

  const codePattern = /^P\d{5}/;
  var ss_code;
  if (codePattern.test(ss_name)) ss_code = ss_name.substring(0, 6);
  
  const baseFolder = DriveApp.getFolderById('1vEX2Z9rcJ-ZUMqHosHYYnEfVHzt1ymJ4');
  var projectFolder;

  switch(formulaInteropSelectFileType) {
    case "plantilla":
      projectFolder = baseFolder.getFoldersByName(`_PLANTILLAS`).next();
      break;
    case "proyecto":
      if (!ss_code) throw new Error(`No se ha encontrado un código de proyecto en el nombre del documento. Añade el código correspondiente o elige otra opción en la configuración.`);
      projectFolder = baseFolder.getFoldersByName(ss_code).next();
      break;
    case "otros":
      projectFolder = baseFolder.getFoldersByName(`_OTROS`).next();
      break;
    default:
      // Acción por defecto si no se cumple ninguna condición
  }

  var folders = projectFolder.getFolders();
  var finalFolder;
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() == ss_name) {
      finalFolder = folder;
      break;
    }
  }

  if (buttonClicked === 'import') {
    if (!finalFolder) throw new Error(`Has seleccionado que el archivo tiene código de proyecto, pero el código no se ha encontrado en el nombre de archivo.`);
  } else {
    if (!finalFolder) {
      finalFolder = projectFolder.createFolder(ss_name);
    }
  }

  // Guarda la carpeta en la caché para futuras ejecuciones
  cache.put(cacheKey, finalFolder.getId());

  return finalFolder;

}

/**
 * formulaDatabaseExport
 * Exporta las fórmulas del documento actual a la base de datos Morph
 */
function formulaDatabaseExport(rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var { formulaInteropSelectFileType, formulaInteropAllDocument, formulaInteropFormat } = rowData;

  var finalFolder = formulaDatabaseProjectFolder(ss, formulaInteropSelectFileType, 'export');

  var sheets; formulaInteropAllDocument ? sheets = ss.getSheets() : sheets = [ ss.getActiveSheet() ];

  for (let i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    var sheetName = sheet.getName();

    var cellFormulas = extractCellFormulas(sheet, formulaInteropFormat);
    var fileContent = composeFileContent(cellFormulas);

    var searchFor = `${sheetName}.gse`;
    let filesIterator = finalFolder.searchFiles(`title='${searchFor}'`);
    if (filesIterator.hasNext()) {
      let file = filesIterator.next();
      file.setContent(fileContent);
    } else {
      var template_file = DriveApp.getFileById('1voTFxQuHs_8Gr_pGFp17mhXhZ9B8Awkw');
      var newFile = template_file.makeCopy(`${sheetName}.gse`, finalFolder);
      var fileBlob = Utilities.newBlob(fileContent, "application/octet-stream");
      newFile.setContent(fileBlob.getDataAsString())
    }
  }
}

function extractCellFormulas(sheet, formulaInteropFormat) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, lastRow, lastColumn);
  var formulas = range.getFormulas();

  var cellFormulas = {};

  var excludedFunctions = ["=HYPERLINK(", "=IMAGE(", "=IMPORTRANGE("]; // lista de funciones excluidas

  for (let row = 0; row < formulas.length; row++) {
    for (let col = 0; col < formulas[row].length; col++) {

      var formula = formulaInteropFormat ? formatFormula(formulas[row][col]) : formulas[row][col];
      var cell = range.getCell(row + 1, col + 1).getA1Notation();

      if (formula !== "") {
        var excludeFormula = false;
        for (let i = 0; i < excludedFunctions.length; i++) {
          if (formula.toUpperCase().startsWith(excludedFunctions[i])) {
            excludeFormula = true;
            break;
          }
        }
        if (!excludeFormula) {
          cellFormulas[cell] = formula;
        }
      }
    }
  }
  return cellFormulas;
}

/*
function extractCellFormulas(sheet, formulaInteropFormat) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(1, 1, lastRow, lastColumn);
  var formulas = range.getFormulas();

  var cellFormulas = {};

  for (let row = 0; row < formulas.length; row++) {
    for (let col = 0; col < formulas[row].length; col++) {

      formula = formulaInteropFormat ? formatFormula(formulas[row][col]) : formulas[row][col];
      //const formula = formulas[row][col];
      if (formula !== "") {
        var cell = range.getCell(row + 1, col + 1).getA1Notation();
        cellFormulas[cell] = formula;
      }
    }
  }
  return cellFormulas;
}
*/

function composeFileContent(cellFormulas) {
  let fileContent = "";
  const cellList = Object.keys(cellFormulas);
  cellList.sort();

  let cellListLength = cellList.length;

  for (let i = 0; i < cellListLength; i++) {
    const cell = cellList[i];
    const formula = cellFormulas[cell];
    fileContent += `// CELL=${cell}\n${formula}${i !== cellListLength - 1 ? '\n\n\n' : '\n'}`;
  }

  return fileContent;
}

/**
 * formulaLogger
 * Script that logs changes made to formulas in a Google Spreadsheet and records details like user, date, and file/cell location in a "Formula Changelog" sheet.
 */
function formulaLogger(rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  let sheetName = sh.getName();
  let formulaSheetName = `Formula Changelog`;

  // Main Variables
  var formData = [
    rowData.formulaText,
    rowData.sendToTemplate
  ];
  var [formulaText, sendToTemplate] = formData;

  //Logger.log(`FORMLATEXT: ${formulaText}, SENDTEMPLATE: ${sendToTemplate}`)

  let user = Session.getActiveUser().getEmail();
  let dateNow = Utilities.formatDate(new Date(), 'GMT+1', 'dd/MM/yyyy - HH:mm');

  let fileName = ss.getName();
  let fileURL = `${ss.getUrl()}#gid=${sh.getSheetId()}`;
  let selectedCell = ss.getCurrentCell();
  let selectedCellNotation = selectedCell.getA1Notation();
  let newFormula = selectedCell.getFormula();

  let fileType; 
  let includesArray = ['Superficies', 'Mediciones', 'Exportación'];
  let foundString = includesArray.find(function(element) {
    return fileName.includes(element);
  });

  if (foundString) {
    fileType = `Cuadro de ${foundString}`;
  } else {
    fileType = `Undefined`;
  }
  
  let controlPanelURL = ss.getSheetByName('LINK').getRange('B1').getValue();
  let controlPanelID = getIdFromUrl(controlPanelURL); Logger.log(`CPURL: ${controlPanelURL}, CPID: ${controlPanelID}`)
  let controlPanelFile = SpreadsheetApp.openById(controlPanelID);
  let controlPanelLoggerSheet = controlPanelFile.getSheetByName(formulaSheetName) || controlPanelFile.insertSheet(formulaSheetName, 200).setTabColor('00FF00');

  if (controlPanelLoggerSheet.getDataRange().isBlank()) {
    controlPanelLoggerSheet.getRange(1, 1, controlPanelLoggerSheet.getMaxRows(), controlPanelLoggerSheet.getMaxColumns()).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("bottom");
    let rangeList = controlPanelLoggerSheet.getRangeList(["B1:B", "G1:G", "H1:H"]);
    rangeList.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    let columns = [{column: 1, width: 180},{column: 2, width: 225},{column: 3, width: 160},{column: 4, width: 100},{column: 5, width: 185},{column: 6, width: 150},{column: 7, width: 60},{column: 8, width: 225},{column: 9, width: 225},{column: 10, width: 100}];
    columns.forEach(function(column) {
      controlPanelLoggerSheet.setColumnWidth(column.column, column.width);
    });

    let headerList = ['Filename', 'Sheet Link', 'Type', 'Date', 'User', 'Sheet', 'Cell', 'Old Formula', 'New Formula', 'Send to Nave Nodriza'];
    let headerRange = controlPanelLoggerSheet.getRange(1, 1, 1, headerList.length);
    headerRange.setValues([headerList]);
    deleteEmptyRows(controlPanelLoggerSheet); removeEmptyColumns(controlPanelLoggerSheet); controlPanelLoggerSheet.appendRow([null]);
    headerRange.setFontWeight('bold');
  }

  formulaText = "'" + formulaText;
  newFormula = "'" + newFormula;

  let dataList = [];
  dataList.push(fileName, fileURL, fileType, dateNow, user, sheetName, selectedCellNotation, formulaText, newFormula, sendToTemplate);

  controlPanelLoggerSheet.getRange(controlPanelLoggerSheet.getLastRow() + 1, 1, 1, dataList.length).setValues([dataList]);
}

/**
 * formulaDropper
 * Returns formula of selected cell in current Google Spreadsheet.
 */
function formulaDropper() {
  let ss = SpreadsheetApp.getActive();
  var selectedCell = ss.getCurrentCell();
  var formula = selectedCell.getFormula();
  Logger.log(formula)
  return formula;
}

/**
 * formatFormula
 * Formatea una formula dada para que sea más sencilla de leer
 * @param {string} formula - La formula a formatear
 * @param {Array} noIndentChars - Caracteres de apertura que no deben causar indentación
 */
function formatFormula(formula) {
  let indentationLevel = 0;
  let formattedFormula = "";
  var indentation = "  "

  for (let i = 0; i < formula.length; i++) {
    const char = formula[i];

    switch (char) {
      case "(":
        formattedFormula += char + "\n" + indentation.repeat(indentationLevel + 1);
        indentationLevel++;
        break;
      case ")":
        indentationLevel--;
        formattedFormula += "\n" + indentation.repeat(indentationLevel) + char;
        break;
      case ",":
        formattedFormula += char + "\n" + indentation.repeat(indentationLevel);
        break;
      case ";":
        formattedFormula += char + "\n" + indentation.repeat(indentationLevel);
        break;
      case ".":
        formattedFormula += char;
        break;
      case "&":
        if (formattedFormula.slice(-1) !== " ") {
          formattedFormula += " ";
        }
        formattedFormula += char;
        if (formula[i + 1] !== " ") {
          formattedFormula += " ";
        }
        break;
      default:
        formattedFormula += char;
        break;
    }
  }

  formattedFormula = cleanFormulas(formattedFormula);

  return formattedFormula;
}

function formatFormulaold(formula) {
  let indentationLevel = 0;
  let formattedFormula = "";

  for (let i = 0; i < formula.length; i++) {
    const char = formula[i];

    switch (char) {
      case "(":

        if (formattedFormula.slice(-1) === "\n") {
          formattedFormula += char;
        } else {
          formattedFormula += char + "\n" + "\t".repeat(indentationLevel + 1);
          indentationLevel++;
        }

        break;
      case ")":
        
        if (formattedFormula.slice(-1) === "\n") {
          formattedFormula += char;
        } else {
          indentationLevel--;
          formattedFormula += "\n" + "\t".repeat(indentationLevel) + char;
        }
        break;
      case ",":
        if (formattedFormula.slice(-1) === "\n") {
          formattedFormula += char;
        } else {
          formattedFormula += char + "\n" + "\t".repeat(indentationLevel);
        }
        break;
      case ";":
        if (formattedFormula.slice(-1) === "\n") {
          formattedFormula += char;
        } else {
          formattedFormula += char + "\n" + "\t".repeat(indentationLevel);
        }
        break;
      case ".":
        formattedFormula += char;
        break;
      case "&":
        if (formattedFormula.slice(-1) !== " ") {
          formattedFormula += " ";
        }
        formattedFormula += char;
        if (formula[i + 1] !== " ") {
          formattedFormula += " ";
        }
        break;
      default:
        formattedFormula += char;
        break;
    }
  }

  formattedFormula = cleanFormulas(formattedFormula);
  
  return formattedFormula;
}



function cleanFormulas(formula) {

  var excludedFormulas = ["CHAR(", "NOW(", "TODAY("];

  var cleanFormula = formula;
  for (var i = 0; i < excludedFormulas.length; i++) {
    var excludedFormula = excludedFormulas[i];
    var formulaIndex = cleanFormula.indexOf(excludedFormula);
    while (formulaIndex !== -1) {
      var startIndex = formulaIndex;
      var endIndex = cleanFormula.indexOf(")", startIndex) + 1;
      var formulaToClean = cleanFormula.substring(startIndex, endIndex);
      var cleanSubFormula = formulaToClean.replace(/^[ \t]+|[ \t]+$/gm, '').replace(/\n/g, '');
      
      cleanFormula = cleanFormula.substring(0, startIndex) + cleanSubFormula + cleanFormula.substring(endIndex);
      formulaIndex = cleanFormula.indexOf(excludedFormula, startIndex + cleanSubFormula.length);
    }
  }
  cleanFormula = cleanFormula.replace(/^\s*$(?:\r\n?|\n)/gm, ""); // Eliminar líneas vacías

  return cleanFormula;
}

function cleanFormulaFormat(formula) {
  var clean = formula.replace(/^[ \t]+|[ \t]+$/gm, '').replace(/\n/g, '');
  return clean;
}


function eliminarEspaciosEnFormulas() {
  var hoja = SpreadsheetApp.getActiveSheet();
  var rango = hoja.getDataRange();
  var formulas = rango.getFormulas();
  for (var i = 0; i < formulas.length; i++) {
    for (var j = 0; j < formulas[i].length; j++) {
      var formula = formulas[i][j];
      if (formula != "") {
        var nuevaFormula = formula.replace(/\s/g, '').replace(/\t/g, '').replace(/\n/g, '');
        hoja.getRange(i+1, j+1).setFormula(nuevaFormula);
      }
    }
  }
}

