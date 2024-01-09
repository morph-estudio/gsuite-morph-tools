// SECTION = FUNCTIONS IN DEVELOPMENT

function refreshFormulasInSelectedRow() {
  // Mostrar un cuadro de diálogo para que el usuario ingrese el número de fila
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Introduce el número de fila:');

  // Obtener el número de fila ingresado por el usuario
  var rowNumber = parseInt(result.getResponseText());

  // Obtener la hoja de cálculo actual
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Obtener la última columna en la fila especificada
  var lastColumn = sheet.getLastColumn();

  // Obtener la gama de celdas en la fila especificada
  var rowRange = sheet.getRange(rowNumber, 1, 1, lastColumn);

  // Obtener los valores en la fila
  var values = rowRange.getValues();

  // Obtener las fórmulas en la fila
  var formulas = rowRange.getFormulas();

  // Recorrer las celdas en la fila y refrescar las fórmulas
  for (var i = 0; i < values[0].length; i++) {
    if (formulas[0][i] !== "") {
      // Borrar el contenido de la celda
      sheet.getRange(rowNumber, i + 1).clearContent();

      // Volver a pegar la fórmula
      sheet.getRange(rowNumber, i + 1).setFormula(formulas[0][i]);
    }
  }
}



/**
 * getWorksheetNamesArray
 * Returns a list of sheetnames, used for the first load after opening the sidebar.
 */
function getSelectedCuadroSheets(tipoCuadro) {
  let sheetNames = new Array();
  let cuadroID = naveNodrizaIDS(tipoCuadro);
  let ss = SpreadsheetApp.openById(cuadroID);
    
  let sheets = ss.getSheets();
  sheets.forEach(sh => {
    sheetNames.push( sh.getName());
  });
  return sheetNames;
}

function obtenerValoresUnicosDeRangoNombrado(hoja, rangoOHeader, header) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(hoja); // Reemplaza 'NombreDeTuHoja' con el nombre de tu hoja
  var rangoNombrado;

  if (header) {
    // Busca la columna que tenga el encabezado proporcionado
    var headers = hoja.getRange(2, 1, 1, hoja.getLastColumn()).getValues()[0];
    var columnIndex = headers.indexOf(header);
    if (columnIndex !== -1) {
      rangoNombrado = hoja.getRange(3, columnIndex + 1, hoja.getLastRow() - 1, 1); // +1 para ajustar el índice de columna
    } else {
      throw new Error("No se encontró el encabezado proporcionado en la hoja.");
    }
  } else {
    // Utiliza el argumento rango si no se proporciona el encabezado
    rangoNombrado = hoja.getRange(rangoOHeader); // Reemplaza 'NombreDeTuRango' con el nombre de tu rango nombrado
  }

  var valores = rangoNombrado.getValues(); // Obtiene los valores del rango nombrado
  var valoresUnicos = obtenerValoresUnicos(valores); // Llama a la función para obtener valores únicos
  // Browser.msgBox(valoresUnicos)
  return valoresUnicos;
}

function obtenerValoresUnicos(array) {
  var valoresUnicos = [];
  for (var i = 0; i < array.length; i++) {
    for (var j = 0; j < array[i].length; j++) {
      if (valoresUnicos.indexOf(array[i][j]) === -1) { // Verifica si el valor ya está en el array de valores únicos
        valoresUnicos.push(array[i][j]); // Agrega el valor único al array
      }
    }
  }
  return eliminarValoresVacios(valoresUnicos);
}

function eliminarValoresVacios(array) {
  var newArray = array.filter(function(valor) {
    return valor !== "" && valor !== "";
  });
  return newArray;
}

function translatePresentation() {
  var presentationId = '1g6hOkkSEw7K2RJlAtsZasFnnET8dVHAXtH9LklPiY2s'; // Asegúrate de poner aquí el ID de tu presentación.
  var presentation = SlidesApp.openById(presentationId);
  var slides = presentation.getSlides();
  
  for (var i = 0; i < slides.length; i++) {
    var shapes = slides[i].getShapes();
    for (var j = 0; j < shapes.length; j++) {
      if (shapes[j].getText) { // Verificamos que el objeto tenga texto
        var originalText = shapes[j].getText().asString();
        //var originalText = textRange.getText();
        if (originalText !== '') { // Verificamos que el texto no esté vacío
          var translatedText = LanguageApp.translate(originalText, 'es', 'en'); // Traducimos el texto
          shapes[j].setText(translatedText); // Establecemos el texto traducido
        }
      }
    }
  }
}

function updateCSDB() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('CS-OP');
  var data = sh.getRange('B4:B').getValues();
  var data_2 = sh.getRange('M4:M').getValues();

  var folderIdsByProjectCode = {
    P05: "0AHyQttPCl3bJUk9PVA",
    P06: "0AMR_I9BV2NhPUk9PVA",
    P07: "0AB4Gq6lkAaEqUk9PVA",
    P08: "0AJxCS-3KvMXEUk9PVA",
    P09: "0AKuz_Ey_m-O4Uk9PVA",
    P10: "0AApa1UiYkMUIUk9PVA",
    P11: "0ADEoIVQFg9Y8Uk9PVA",
    P12: "0AOG8vnC0w0qsUk9PVA",
    P13: "0AKAUEcok-cppUk9PVA",
    P16: "0AKMKawyNToCKUk9PVA"
  };

  var folderPairs = [
    {
      folderStructure: {
        a: ['Trabajo'],
        b: ['Arquitectura'],
        c: ['Doc Escrita'],
        d: ['Cuadro Superficies', 'CS']
      },
      filesPattern: {
        "superficies": {
          pattern: { a: `Cuadro Superficies` },
          columnToSetValue: 'O',
          filetype: 'google_sheets'
        },
        "exportacion": {
          pattern: { a: `Exportación Superficies` },
          columnToSetValue: 'S',
          filetype: 'google_sheets'
        },
        "cliente": {
          pattern: { a: `Cliente`, b: `AVINTIA` },
          columnToSetValue: 'W',
          filetype: 'google_sheets'
        }
      }
    },
    {
      folderStructure: {
        a: ['Trabajo'],
        b: ['Arquitectura'],
        c: ['Maq']
      },
      filesPattern: {
        "slides": {
          pattern: { a: `CD` },
          columnToSetValue: 'N',
          filetype: 'google_slides'
        }
      }
    }
  ];

  for (var i = 0; i < data.length; i++) {
    var projectCode = data[i][0];
    var projectFolderURL = data_2[i][0].toString().trim();
    var projectFolder = null;

    if (projectFolderURL !== "") {
      Logger.log(`${projectCode}: Value in COLUMN M already exists. Skipping root folder search.`);
      var projectFolderID = getFolderIdFromUrl(projectFolderURL);
      projectFolder = DriveApp.getFolderById(projectFolderID);
    } else {
      var folderId = folderIdsByProjectCode[projectCode.substring(0, 3)];
      var folders = DriveApp.getFolderById(folderId).getFolders();
      //Logger.log(`Share Drive is: ${projectCode.substring(0, 3)} and its URL: ${folderId}`);

      while (folders.hasNext()) {
        var folder = folders.next();
        if (folder.getName().includes(projectCode)) {
          projectFolderURL = folder.getUrl();
          sh.getRange('M' + (i + 4)).setValue(projectFolderURL);
          projectFolder = folder;
          break;
        }
      }
    }

    if (projectFolder !== null) {
      for (var j = 0; j < folderPairs.length; j++) {
        var folderStructure = folderPairs[j].folderStructure;
        var filesPattern = folderPairs[j].filesPattern;
        var rootFolderId = projectFolder.getId();
        var finalFolderID = followFolderStructure(rootFolderId, folderStructure);
        var finalFolder = DriveApp.getFolderById(finalFolderID);

        for (var key in filesPattern) {
          var patterns = filesPattern[key].pattern;
          var columnToSetValue = filesPattern[key].columnToSetValue;
          var filetype = filesPattern[key].filetype;
          
          // Check if the value already exists in the column
          var currentValue = sh.getRange(columnToSetValue + (i + 4)).getValue();
          if (currentValue !== "") {
            Logger.log(`${projectCode}: Value in COLUMN ${columnToSetValue} already exists. Skipping file search.`);
            continue;
          }

          var matchingFiles = findMatchingFiles(finalFolder, patterns, filetype);

          if (matchingFiles.length > 0) {
            matchingFiles.sort(function(a, b) {
              return a.getLastUpdated() < b.getLastUpdated() ? 1 : -1;
            });
            var definitiveFile = matchingFiles[0];
            Logger.log(`${projectCode}: Found file: ${definitiveFile.getName()}, Last updated: ${definitiveFile.getLastUpdated()}`);
            sh.getRange(columnToSetValue + (i + 4)).setValue(definitiveFile.getUrl());
          } else {
            Logger.log(`${projectCode}: No files found matching the pattern.`);
          }
        }
      }
    }
  }
}

function followFolderStructure(rootFolderId, folderStructure) {
  var folder = DriveApp.getFolderById(rootFolderId);
  for (var key in folderStructure) {
    var folderNamesToMatch = folderStructure[key];
    var subFolders = folder.getFolders();
    var foundFolder = null;

    while (subFolders.hasNext()) {
      var subFolder = subFolders.next();
      var folderName = subFolder.getName();
      var folderNameMatches = folderNamesToMatch.some(name => folderName.includes(name));
      
      if (folderNameMatches) {
        foundFolder = subFolder;
        break;
      }
    }

    if (foundFolder) {
      folder = foundFolder;
    } else {
      continue;
      //throw new Error('No se encuentra la carpeta con el nombre: ' + folderNamesToMatch.join(' / '));
    }
  }

  return folder.getId();
}

function findMatchingFiles(rootFolder, patterns, filetype) {
  var matchingFiles = [];
  var rootFiles = rootFolder.getFilesByType(getMimeTypeByFileType(filetype));

  while (rootFiles.hasNext()) {
    var rootFile = rootFiles.next();
    if (matchesPatterns(rootFile.getName(), patterns)) {
      matchingFiles.push(rootFile);
    }
  }

  if (matchingFiles.length === 0) {
    var subFolders = rootFolder.getFolders();

    while (subFolders.hasNext()) {
      var subFolder = subFolders.next();
      var subFiles = subFolder.getFilesByType(getMimeTypeByFileType(filetype));

      while (subFiles.hasNext()) {
        var subFile = subFiles.next();
        if (matchesPatterns(subFile.getName(), patterns)) {
          matchingFiles.push(subFile);
        }
      }
    }
  }

  return matchingFiles;
}

function matchesPatterns(filename, patterns) {
  for (var key in patterns) {
    var pattern = patterns[key];
    if (filename.includes(pattern)) {
      return true;
    }
  }
  return false;
}

function getMimeTypeByFileType(filetype) {
  //Logger.log(`Mimetype: ${filetype}`);
  switch (filetype) {
    case 'google_sheets':
      return 'application/vnd.google-apps.spreadsheet';
    case 'google_slides':
      return 'application/vnd.google-apps.presentation';
    // Agregar más tipos de archivos según sea necesario
    default:
      return '';
  }
}

function getFolderIdFromUrl(url) {
  var id = "";
  var match = url.match(/[-\w]{25,}/);
  if (match) {
    id = match[0];
  }
  return id;
}










































function getVideoId(url) {
  var videoId = extractVideoId(url);
  return videoId;
}

function extractVideoId(url) {
  var videoId = "";
  var regExp = /^.*((youtu.be\/)|(v\/)|(\/u\/\w\/)|(embed\/)|(watch\?))\??v?=?([^#\&\?]*).*/;
  var match = url.match(regExp);
  if (match && match[7].length == 11) {
    videoId = match[7];
  } else {
    videoId = "Invalid URL";
  }
  return videoId;
}

function convertirReferenciasAbsolutas() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var datarange = sh.getDataRange();
  var backgroundColors = datarange.getBackgrounds();
  var values = datarange.getValues();
  var numRows = backgroundColors.length;
  var numCols = backgroundColors[0].length;
  
  for (var i = 0; i < numRows; i++) {
    for (var j = 0; j < numCols; j++) {
      var color = backgroundColors[i][j];
      switch (color) {
        case "#CCCCCC":
          // Hacer algo si la celda tiene fondo rojo
          sh.getRange(i+1, j+1).setBackground();
          break;
        default:

          break;
      }
    }
  }
}

function convertirReferenciasAbsolutas2() {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rangoSeleccionado = hoja.getActiveRange();
  var formulas = rangoSeleccionado.getFormulas();
  var filas = formulas.length;
  var columnas = formulas[0].length;

  for (var fila = 0; fila < filas; fila++) {
    for (var columna = 0; columna < columnas; columna++) {
      var formula = formulas[fila][columna];
      if (formula) {
        // Restar 1 al número de fila en cada celda con fórmula
        var nuevaFormula = formula.replace(/\$?([A-Z]+)\$?(\d+)/g, function(match, col, row) {
          return col + (parseInt(row) - 1);
        });
        rangoSeleccionado.getCell(fila+1, columna+1).setFormula(nuevaFormula);
      }
    }
  }
}



function helperToMoveFormulas() {

  var rowToMoveFormulas = 3;

  var sheet = SpreadsheetApp.getActiveSheet();
  var selectedRange = sheet.getActiveRange();
  var startRow = selectedRange.getRow();
  var startColumn = selectedRange.getColumn();
  var numRows = selectedRange.getNumRows();
  var numColumns = selectedRange.getNumColumns();
  var formulas = sheet.getRange(startRow, startColumn, numRows, numColumns).getFormulas()[0];
  
  for (var i = 0; i < formulas.length; i++) {
    if (formulas[i] !== "") {
      var colIndex = startColumn + i;
      var range = sheet.getRange(rowToMoveFormulas, colIndex);
      var formula = '={"";' + formulas[i].slice(1) + '}';
      range.setFormula(formula);
    }
  }

}

function addComponent() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var rangeList = sh.getRangeList([ 'P1884', 'AQ1884', 'AS1884', 'AU1884', 'AX1884', 'BA1884', 'BE1884', 'BF1884' ]);
  var values = [['m2'], ['Fachada'], ['Modular'], ['Prueba2'], ['Notas generales al capítulo aislamiento horizontal'], ['0'], ['Poliuretano proyectado 30 celda cerrada CCC4']];

  rangeList.getRanges().forEach(function(range, index) {
    range.setValues([values[index]]);
  });
}

function listNamedRanges() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var namedRanges = ss.getNamedRanges();
  var data = [];

  var name, sheetName, notation, headerRange, header;
  
  for (var i = 0; i < namedRanges.length; i++) {
    name = namedRanges[i].getName();
    range = namedRanges[i].getRange()
    notation = range.getA1Notation();
    sheetName = range.getSheet().getName();
    headerRange = sheetName + "!" + notation.split(":")[0].replace(/[0-9]/g, '') + "1";
    //header = ss.getSheetByName(sheetName).getRange(headerRange).getValue();
    data.push([name, sheetName, notation]);
  }

  var cell = sheet.getActiveCell();
  
  var numRows = data.length;
  var numCols = data[0].length;
  sheet.getRange(cell.getRow(), cell.getColumn(), numRows, numCols).setValues(data);
  
  return data;
}


/**
 * macroModificarCuadros
 * Función para realizar macros rápidas para cualquier necesidad en un documento.
 */
function macroModificarCuadros(a) {

  var nombreHoja = "X Variables"; // Nombre de la hoja donde se encuentran las variables
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombreHoja);
  if (hoja.getName() != "X_Variables") hoja.setName("X_Variables"); // Cambia el nombre de la hoja a "X_VARIABLES"
  var ultimaColumnaDatos = hoja.getLastColumn() + 1;

  hoja.insertColumns(hoja.getMaxColumns(), 12);
  
  var formula = `=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"TXT_OP_TEMPLATE!A1:G")`;
  var columnaFormula = ultimaColumnaDatos + 5; // Columna siguiente a la última con datos
  var rangoFormula = hoja.getRange(1, columnaFormula, 1, 1);
  
  rangoFormula.setFormula(formula); // Aplica la fórmula en la columna correspondiente

}

/**
 * substringsColorTool
 * Format text for specific fragments of text in the cells.
 */
function substringsColorTool(sctRowData, sctDivCounterArray) {

  Logger.log(sctDivCounterArray)

  for (let i = 0; i < sctDivCounterArray.length; i++) {

    var formData = [
      sctRowData[`sctSheetSelector-${sctDivCounterArray[i]}`],
      sctRowData[`sheetRange-${sctDivCounterArray[i]}`],
      sctRowData[`textSubstring-${sctDivCounterArray[i]}`],
      sctRowData[`colorpicker-${sctDivCounterArray[i]}`],
      sctRowData[`stylecheckBold-${sctDivCounterArray[i]}`],
      sctRowData[`stylecheckItalic-${sctDivCounterArray[i]}`]
    ];

    var [sheetSelector, sheetRange, textSubstring, colorpicker, stylecheckBold, stylecheckItalic] = formData;

    Logger.log(sheetSelector); Logger.log(colorpicker); Logger.log(stylecheckBold); Logger.log(stylecheckItalic);

    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetSelector);
    var numRows = sh.getLastRow(); var numCols = sh.getLastColumn();

    for (var r = 1; r <= numRows; r++) {
      for (var c = 1; c <= numCols; c++) {
        var cell = sh.getRange(r, c);
        var text = cell.getValue().toString();

        if (text.indexOf(textSubstring) > -1) {
          var substring = textSubstring;
          
          var startIndex = text.indexOf(substring);
          var endIndex = startIndex + substring.length;
          var textStyle = SpreadsheetApp.newTextStyle().setForegroundColor(colorpicker).setBold(stylecheckBold).setItalic(stylecheckItalic).build();
          var value = SpreadsheetApp.newRichTextValue()
              .setText(text)
              .setTextStyle(startIndex, endIndex, textStyle)
              .build();
          cell.setRichTextValue(value);
        }
      }
    }

  }
}

/**
 * printValidation
 * Create a series of PDF files based on the values of a drop-down cell.
 */
function printValidation(rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getActiveSheet();

  let formData = [
    rowData.secuencialFilename,
    rowData.secuencialDestFolder,
    rowData.secuencialSize,
    rowData.secuencialOrientation,
    rowData.secuencialMarTopBottom,
    rowData.secuencialMarLeftRight,
    rowData.secuencialFitw,
    rowData.secuencialGridlines,
    rowData.secuencialFzr
  ];

  let [secuencialFilename, secuencialDestFolder, secuencialSize, secuencialOrientation, secuencialMarTopBottom, secuencialMarLeftRight, secuencialFitw, secuencialGridlines, secuencialFzr] = formData;

  let selection = ss.getSelection();
  let currentCell = selection.getCurrentCell().getA1Notation();
  let ss_id = ss.getId();
  let ss_url = ss.getUrl();
  let file = DriveApp.getFileById(ss_id);
  let parentFolder;
  secuencialDestFolder == '' ? parentFolder = file.getParents().next() : parentFolder = DriveApp.getFolderById(getIdFromUrl(secuencialDestFolder));

  let dv = sh.getRange(currentCell).getDataValidation();
  let critVal = dv.getCriteriaValues();
  let validationValues = critVal[0].getValues();
  validationValues = transpose(validationValues)[0];
  validationValues = validationValues.filter(n => n)

  let fileName; let blob; Logger.log(validationValues);
  
  validationValues.forEach((name) => {
    sh.getRange(currentCell).setValue(name);
    SpreadsheetApp.flush();
    if(secuencialFilename.includes('{{cell}}')) {
      fileName = secuencialFilename.replace('{{cell}}', name);
    } else {
      fileName = secuencialFilename;
    }
    blob = _getAsBlob(ss_url, sh, secuencialSize, secuencialOrientation, secuencialMarTopBottom, secuencialMarLeftRight, secuencialFitw, secuencialGridlines, secuencialFzr);
    blob = blob.setName(fileName);
    parentFolder.createFile(blob);
  });
}

function _getAsBlob(url, sheet, secuencialSize, secuencialOrientation, secuencialMarTopBottom, secuencialMarLeftRight, secuencialFitw, secuencialGridlines, secuencialFzr, range) {
  var rangeParam = ''
  var sheetParam = ''
  if (range) {
    rangeParam =
      '&r1=' + (range.getRow() - 1)
      + '&r2=' + range.getLastRow()
      + '&c1=' + (range.getColumn() - 1)
      + '&c2=' + range.getLastColumn()
  }
  if (sheet) {
    sheetParam = '&gid=' + sheet.getSheetId()
  }
  // A credit to https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  // these parameters are reverse-engineered (not officially documented by Google)
  // they may break overtime.
  var exportUrl = url.replace(/\/edit.*$/, '')
      + '/export?exportFormat=pdf&format=pdf'
      + '&size=' + secuencialSize
      + '&portrait=' + secuencialOrientation
      + '&fitw=' + secuencialFitw
      + '&top_margin=' + secuencialMarTopBottom
      + '&bottom_margin=' + secuencialMarTopBottom
      + '&left_margin=' + secuencialMarLeftRight
      + '&right_margin=' + secuencialMarLeftRight
      + '&sheetnames=false&printtitle=false'
      + '&pagenum=UNDEFINED' // change it to CENTER to print page numbers
      + '&gridlines=' + secuencialGridlines
      + '&fzr=' + secuencialFzr
      + sheetParam
      + rangeParam
      
  // Logger.log('exportUrl=' + exportUrl)
  var response
  var i = 0
  for (; i < 5; i += 1) {
    response = UrlFetchApp.fetch(exportUrl, {
      muteHttpExceptions: true,
      headers: { 
        Authorization: 'Bearer ' +  ScriptApp.getOAuthToken(),
      },
    })
    if (response.getResponseCode() === 429) {
      // printing too fast, retrying
      Utilities.sleep(3000)
    } else {
      break
    }
  }
  
  if (i === 5) {
    throw new Error('Printing failed. Too many sheets to print.')
  }
  
  return response.getBlob()
}

/**
 * historicoDeSuperficies
 * Crea un nuevo histórico en el histórico del cuadro de superficies
 */
function historicoDeSuperficies() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateNow = Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy');

  var sheets = [
    { name: 'Histórico CONST', sheetRef: 0 },
  ];

  var cellMappings = {
    0: {
      mainCell: 'D1',
      firstCell: 'G1',
      secondCell: 'H1',
      thirdCell: 'E1',
      groupRange: 'H:I'
    },
    1: {
      mainCell: 'E1',
      firstCell: 'H1',
      secondCell: 'I1',
      thirdCell: 'F1',
      groupRange: 'I:J'
    }
  };

  sheets.forEach(function(sheetInfo) {
    var sh = ss.getSheetByName(sheetInfo.name);
    if (!sh) {
      throw new Error('No se encontró la hoja correspondiente al histórico de superficies: ' + sheetInfo.name);
    }

    var mapping = cellMappings[sheetInfo.sheetRef];

    if (!mapping) {
      throw new Error('Valor de sheetRef no válido para la hoja ' + sheetInfo.name);
    }

    var mainCell = mapping.mainCell;
    var firstCell = mapping.firstCell;
    var secondCell = mapping.secondCell;
    var thirdCell = mapping.thirdCell;
    var groupRange = mapping.groupRange;

    var mainRange = sh.getRange(mainCell);
    var secondRange = sh.getRange(secondCell);
    var firstRange = sh.getRange(firstCell);
    var mainColumnIndex = mainRange.getColumn();
    var firstColumnIndex = firstRange.getColumn();

    // Coger la fórmula original para luego reutilizarla

    var originalFormulaRange = sh.getRange(thirdCell);
    var originalFormula = originalFormulaRange.getFormulas();

    // Comprobar si han cambiado los datos desde el último histórico

    var freezeRange;
    var lastRow = sh.getLastRow();

    if (firstRange.isBlank()) {
      freezeRange = sh.getRange(1, mainColumnIndex, lastRow, 1);
      freezeRange.copyTo(sh.getRange(1, firstColumnIndex), { contentsOnly: true });
      sh.getRange(firstCell).setValue(dateNow);
    } else {
      var mainColumnRange = sh.getRange(2, mainColumnIndex, lastRow, 1).getValues();
      var firstColumnRange = sh.getRange(2, firstColumnIndex, lastRow, 1).getValues();
      var isEqual = mainColumnRange.every(function(row, i) {
        return row[0] === firstColumnRange[i][0];
      });

      if (isEqual) {
        throw new Error('Los valores de superficies no han cambiado desde el último histórico.');
      }

      // Insertar el histórico de datos

      freezeRange = sh.getRange(1, mainColumnIndex, lastRow, 3);

      sh.insertColumns(firstColumnIndex, 3);
      freezeRange.copyTo(sh.getRange(1, firstColumnIndex), { contentsOnly: true });
      sh.getRange(1, firstColumnIndex + 1, lastRow, 1).clearContent();
      sh.getRange(2, firstColumnIndex - 1, lastRow, 1).clearContent();
      firstRange.setValue(dateNow);

      // Añadir formato

      var columns = [
        { column: firstColumnIndex, width: 100 },
        { column: firstColumnIndex + 1, width: 100 },
        { column: firstColumnIndex + 2, width: 150 }
      ];

      columns.forEach(function(column) {
        sh.setColumnWidth(column.column, column.width);
      });

      sh.getRange(1, firstColumnIndex - 2, 1, 2).copyTo(sh.getRange(1, firstColumnIndex + 1), { formatOnly: true }); // Formato a los nuevos encabezados
      sh.getRange(2, firstColumnIndex + 1, sh.getMaxRows() - 1, 2).setBackgroundColor(null).setFontColor('black'); // Formato a las nuevas columnas 2 y 3
      secondRange.setBorder(false, false, false, false, false, false);
      sh.getRange(groupRange).shiftRowGroupDepth(1);

      // Restaurar fórmulas

      var newFormula = '={"Diferencia con " & IF(TO_TEXT(' + (sheetInfo.sheetRef === 0 ? 'J1' : 'K1') + ') <> ""; TO_TEXT(' + (sheetInfo.sheetRef === 0 ? 'J1' : 'K1') + '); "última fecha"); ARRAYFORMULA(IF(B2:B <> ""; IF(TO_TEXT(' + (sheetInfo.sheetRef === 0 ? 'J2:J' : 'K2:K') + ') <> ""; ' + (sheetInfo.sheetRef === 0 ? 'G2:G-J2:J' : 'H2:H-K2:K') + '; 0);))}';
      originalFormulaRange.setFormula(originalFormula);
      secondRange.setFormula(newFormula);
    }
  });
}


/**
 * adaptarCuadroAntiguo
 * Script to adapt old surface tables to the new automatic structure.
 */
function adaptarCuadroAntiguo() {
  let ss = SpreadsheetApp.getActive();
  let sheetnames = getSheetnames(ss);
  let ui = SpreadsheetApp.getUi();

  // Change Sheet Names and create LINK sheet

  if (sheetnames.indexOf('ACTUALIZAR') > -1) {
    let sh_act = ss.getSheetByName('ACTUALIZAR');
    sh_act.setName('LINK');
    let sh_link = ss.insertSheet('LINK_temp', 0);
    linkPageTemplateFormat(sh_link); linkPageTemplateText(sh_link); deleteEmptyRows(); removeEmptyColumns();
    ss.deleteSheet(sh_act);
    sh_link.setName('LINK').setTabColor('FFFF00');
  }

  let oldSheets = ['TXT LIMPIO','TXT FT','TXT VN'];
  let newSheets = ['TXT SUPERFICIES','TXT FALSOS TECHOS','TXT VENTANAS'];

  for (let i = 0; i < oldSheets.length; i++) {
    if (sheetnames.indexOf(oldSheets[i]) > -1) {
      ss.getSheetByName(oldSheets[i]).setName(newSheets[i]).setTabColor('00FF00');
    }
  }

  // Change Export Folder Name

  let ss_id = ss.getId();
  let file = DriveApp.getFileById(ss_id);
  let parents = file.getParents();
  let carpetaBase = parents.next();
  let searchFor = `title contains 'Exportaciones' or title contains 'Exportación' or title contains 'Exportar' or title contains 'Exportados'`;
  let expFolder = carpetaBase.searchFolders(searchFor); Logger.log(expFolder)
  let a;

  try {
    let expFolderDef = expFolder.next();
    expFolderDef.setName(expFolderDef.getName().replace('Exportaciones', 'ExpTXT').replace('Exportación', 'ExpTXT').replace('Exportar', 'ExpTXT').replace('Exportados', 'ExpTXT'))
  } catch (e) {
    a = true;
  }

  if (a == true) {
    ui.alert('Aviso', 'No se ha encontrado la carpeta de Exportaciones .txt dentro de la carpeta del Cuadro de Superficies. Debes modificarlo manualmente añadiendo "ExpTXT" en el nombre (siguiendo la estructura PXXXXX-A-CS-ExpTXT)', ui.ButtonSet.OK)
  }

  setCuadroAdaptado();
}

function setCuadroAdaptado() {
  PropertiesService.getDocumentProperties().setProperties({
    'adaptedSpreadsheet': true,
  });

  sidebarIndex();
}

/**
 * fExportXML
 * Función en desarrollo
 */
function fExportXML() {
  const values = sh().getDataRange().getValues();
  return '<sheet>' + values.map((row, i) => {
    return '<row>' + row.map(function(v) {
      return '<cell>' + v + '</cell>';
    }).join('') + '</row>';
  }).join('') + '</sheet>';
}

function exportXML() {
  var content;
  try {
    content = fExportXML();
  } catch(err) {
    content = '<error>' + (err.message || err) + '</error>';
  }
  return ContentService.createTextOutput(content)
    .setMimeType(ContentService.MimeType.XML).downloadAsFile('Hola');
}

/**
 * saveSheetAsTSV
 * Guarda la hoja en formato TSV manteniendo las fórmulas
 */
function saveSheetAsTSV() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getActiveSheet();

  const ui = SpreadsheetApp.getUi();
  let result = ui.prompt(
    'Carpeta de destino',
    'Introduce el LINK de la carpeta donde guardar el archivo.\nSi se deja en blanco se creará una nueva carpeta en Mi Unidad.',
    ui.ButtonSet.OK_CANCEL,
  );

  let button = result.getSelectedButton();
  let userResponse = result.getResponseText();
  let folder; let externalFolderId;
  if (button == ui.Button.OK) {
    if (userResponse === '') {
      folder = rootFolder.createFolder('TSV Exports');
      externalFolderId = folder.getid();
    } else {
      externalFolderId = getIdFromUrl(userResponse);
      folder = DriveApp.getFolderById(externalFolderId);
    }
  }

  fileName = sh.getName() + ".txt";
  var tsvFile = convertRangeTotsvFile(fileName, sh);
  folder.createFile(fileName, tsvFile);
  // Browser.msgBox('Files are waiting in a folder named ' + folder.getName());
}

function convertRangeTotsvFile(tsvFileName, sheet) {
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var formula = activeRange.getFormulas();
    var tsvFile = undefined;
    // loop through the data in the range and build a string with the tsv  data
    if (data.length > 1) {
      var tsv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (formula[row][col] !== '') {
            data[row][col] = formula[row][col]
          }
          if (data[row][col].toString().indexOf("\t") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }
        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          tsv += data[row].join("\t") + "\r\n";
        }
        else {
          tsv += data[row].join("\t");
        }
      }
      tsvFile = tsv;
    }
    return tsvFile;
  }
  catch(err) {
    Browser.msgBox(err);
  }
}

// SECTION = AUTOFOLDERTREE AND PROJECT FOLDER INIT

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

// SECTION = MORPH CHATBOT DEVELOPMENT

/**
 * botBrainSave
 * Guarda las hojas de un documento de Sheets en formato CSV y las sube a un bucket de Cloud Storage
 * Funciones dependientes: uploadFivaroGCS(), getService(params), authCallback(request)
 */
 function botBrainSave() {
  const sheets = ss().getSheets();
  var fileName;

  sheets.forEach((sheet) => { 
    fileName = sheet.getName() + ".csv";
    var url = null; var blob;
    url = `https://docs.google.com/spreadsheets/d/${ss().getId()}/gviz/tq?tqx=out:csv&gid=${sheet.getSheetId()}`;
    if (url) {
      blob = UrlFetchApp.fetch(url, {
        headers: { authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      }).getBlob();
    }

    // var file = folder.createFile(blob).setName(fileName);

    var params = {
      CLIENT_ID: '443000249830-g093ottepd29j2s13j8ki8ts5lfdou24.apps.googleusercontent.com',
      CLIENT_SECRET: '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',
      BUCKET_NAME: 'morph-bot-brain',
      FILE_PATH: `knowledge-base/${fileName}`,
      DRIVE_FILE: blob,
    };
    uploadFivaroGCS(params);
  });
}

function uploadFivaroGCS(params) {
  var service = getService(params);
  if (!service.hasAccess()) {
    // Logger.log('Please authorize %s', service.getAuthorizationUrl());
    openExternalUrlFromMenu(service.getAuthorizationUrl());
    return;
  }

  var blob = params.DRIVE_FILE;
  var bytes = blob.getBytes();

  var url = 'https://www.googleapis.com/upload/storage/v1/b/BUCKET/o?uploadType=media&name=FILE'
    .replace('BUCKET', params.BUCKET_NAME)
    .replace('FILE', encodeURIComponent(params.FILE_PATH)); Logger.log('fileurl: ' + url)

  var response = UrlFetchApp.fetch(url, {
    method: 'POST',
    contentLength: bytes.length,
    contentType: blob.getContentType(),
    payload: bytes,
    headers: {
      Authorization: `Bearer ${service.getAccessToken()}`,
    },
  });

  var result = JSON.parse(response.getContentText());
  Logger.log(JSON.stringify(result, null, 2));
}

function getService(params) {
  return OAuth2.createService('ctrlq')
    .setAuthorizationBaseUrl('https://accounts.google.com/o/oauth2/auth')
    .setTokenUrl('https://accounts.google.com/o/oauth2/token')
    .setClientId(params.CLIENT_ID)
    .setClientSecret(params.CLIENT_SECRET)
    .setCallbackFunction('authCallback')
    .setPropertyStore(PropertiesService.getUserProperties())
    .setScope('https://www.googleapis.com/auth/devstorage.read_write')
    .setParam('access_type', 'offline')
    .setParam('approval_prompt', 'force')
    .setParam('login_hint', Session.getActiveUser().getEmail());
}

function authCallback(request) {
  var service = getService();
  var authorized = service.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput('Connected to Google Cloud Storage');
  }
  return HtmlService.createHtmlOutput('Access Denied');
}

// SECTION = DEPRECATED FUNCTIONS

/**
 * drawPalette
 * Genera la paleta de colores del Cuadro de Superficies
 */
function drawPalette() {
  SpreadsheetApp.getActiveSpreadsheet().insertSheet('chart');
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('chart').getRange(1, 1, 1, 104).setBackgrounds([
    [

    "#F2F2F2","#FFF3F3","#EEECE1","#D4C9C6","#CDC8CE","#BDBDBD","#A5A5A5","#7F7F7F", // PALETA GREY / BROWN
    "#FFF1CB","#FFE9AD","#FFFAB1","#FFFFAE","#FFE699","#FFE499","#FFFF99","#FCF58F","#FFEB84","#FFFB84","#FFD666","#FED166","#F7CB4D","#ECFF49", // PALETA YELLOW
    "#FFD7AE","#ED956F","#ED7D31","#FF6F31","#F57C00", // PALETA ORANGE
    "#FFEEF6","#FFE0EF","#E7CFCF","#F9CCC4","#E6BFD2","#FAC5DC","#EEBBD1","#E6B3B3","#F1ADC9","#E2ABC5","#EEA3C6","#F097A3","#E78BB6","#FD7D8F","#F08090","#DD808D","#E67C73","#FF6565","#D16969","#E04C4C", // PALETA RED
    "#D7D7FF","#FFDDFF","#E2C6FF","#FFC4FF","#E0A7EB","#D797E4","#FF8FFF","#FF62FF", // PALETA PINK
    "#F0FDFF","#DBE5F1","#E0F7FA","#ECFFF5","#E0FFFC","#DEFFFF","#D7FDFF","#CCF2FF","#C4F7F7","#B9EBFD","#B7D4F0","#C5FFFF","#C4FDF8","#A8DFF3","#A3C3C9","#AFF5EF","#94E6DF","#7FCFEC","#74BFDB","#7BDFD6","#71B9D3","#62D1C7","#5CA4BD","#31AA9F","#3D8DA8","#1B8B81","#1B6F8B", // PALETA BLUE
    "#EAF8D0","#C7D9B7","#D5E6B6","#E4EBA9","#CCFFCC","#A9DFC5","#CCCF96","#BEE2A6","#B8F1B8","#CDF39C","#CEED8A","#BCDA85","#A7DB85","#CEFF70","#A4C26E","#9DE063","#41CF7F","#8FA369","#63BE7B","#57BB8A","#6AA369","#008000" // PALETA GREEN

    ],
  ]);
  let ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('chart');
  ss.deleteSheet(sheet);
}

/**
 * mergeColumns
 * Combina los datos de las columnas con mismo encabezado
 */
function mergeColumns() {
  let transpose = (ar) => ar[0].map((_, c) => ar.map((r) => r[c]));
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getActiveSheet();
  let values = sh.getDataRange().getValues();
  let temp = [
    ...transpose(values)
      .reduce(
        (m, [a, ...b]) => m.set(a, m.has(a) ? [...m.get(a), ...b] : [a, ...b]),
        new Map(),
      )
      .values(),
  ];
  let res = transpose(temp);
  sh.clearContents();
  sh.getRange(1, 1, res.length, res[0].length).setValues(res);
}

function rowToDict(sheet, rownumber) { // Ni idea de para qué era esta función
  var columns = sheet.getRange(1,1,1, sheet.getMaxColumns()).getValues()[0];
  var data = sheet.getDataRange().getValues()[rownumber-1];
  var dict_data = {};
  for (var keys in columns) {
    var key = columns[keys];
    dict_data[key] = data[keys];
  }
  return dict_data;
}

// SECTION = CUSTOM FUNCTIONS FOR THE DEVELOPER SECTION

/**
 * getDatabaseColumn
 * Devuelve los valores de una columna en documento Sheets externo a través de su título.
 */
function getDatabaseColumn(headerName) {
  let params = {
    muteHttpExceptions: true,
  };
  const parsedDB = JSON.parse(UrlFetchApp.fetch('https://opensheet.elk.sh/1lcymggGAbACfKuG0ceMDWIIB9zWuxgVtSR9qpgNq4Ng/Permissions', params).getContentText());
  const dbColumn = parsedDB.map(i => i[headerName]);
  return dbColumn;
}

/**
 * getUserRolePermission
 * Comprueba en la base de datos si el usuario tiene acceso a la información.
 */
function getDevPermission() {
  const userMail = Session.getActiveUser().getEmail();

  // Permission Database: https://docs.google.com/spreadsheets/d/1lcymggGAbACfKuG0ceMDWIIB9zWuxgVtSR9qpgNq4Ng/edit#gid=0

  const userDevPermission = getDatabaseColumn(`devAreaPermission`);
  let devAreaPermission = userDevPermission !== '' && userDevPermission.indexOf(userMail) > -1 ? true : false;

  const userformulaMODPermission = getDatabaseColumn(`formulaModPermission`);
  let formulaModPermission = userformulaMODPermission !== '' && userformulaMODPermission.indexOf(userMail) > -1 ? true : false;

  const userLoggerMODPermission = getDatabaseColumn(`loggerModPermission`);
  let loggerModPermission = userformulaMODPermission !== '' && userLoggerMODPermission.indexOf(userMail) > -1 ? true : false;

  const devGlobalKeys = getDatabaseColumn(`devGlobalKeys`);

  const databaseManualKeys = getDatabaseColumn(`databaseManualKeys`);

  var permission = {
    devAreaPermission: devAreaPermission,
    devGlobalKeys: devGlobalKeys,
    formulaModPermission: formulaModPermission,
    databaseManualKeys: databaseManualKeys,
    loggerModPermission: loggerModPermission
  };

  Logger.log(permission);
  return permission;
}

/**
 * getDevPassword
 * Comprueba en la base de datos si la contraseña es correcta.
 */
function getDevPassword(headerName) {
  const devPassArray = getDatabaseColumn(headerName);
  return devPassArray;
}
