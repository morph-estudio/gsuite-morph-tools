/**
 * Gsuite Morph Tools - Morph Spreadsheet Configurator
 * Developed by alsanchezromero
 *
 * Morph Estudio, 2023
 */

/**
 * naveNodrizaIDS
 * Devuelve el ID de diferentes archivos clave de Morph
 */
function naveNodrizaIDS(file) {
  switch (file) {
    case 'naveNodrizasuperficies':
      return '1_Qq8y_cC5V9lSThCypq0Qpbj6JXrtv-QBBlt7ghy0pI';
    default:
  }
}

/**
 * mainCopySheetFromTemplate
 * Trae hojas de plantilla al documento actual
 */
function mainCopySheetFromTemplate(rowData, ovewriteSwitch) {
  Logger.log(`ovewriteSwitch: ${ovewriteSwitch}`)
  var templateSpreadsheet = SpreadsheetApp.openById(naveNodrizaIDS('naveNodrizasuperficies'));
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = activeSpreadsheet.getActiveSheet();
  var functionErrors = [];

  function processSheet(sheet) {
    if (!sheet.isTrue) return;

    var templateSheet = templateSpreadsheet.getSheetByName(sheet.name);
    if (!templateSheet) {
      functionErrors.push('La hoja ' + sheet.name + ' no se encontró en la plantilla.');
      return;
    }

    var newSheet;

    if (ovewriteSwitch) {

      var existingSheet = activeSpreadsheet.getSheetByName(sheet.name);

      if (existingSheet) {
        var sheetIndex = existingSheet.getIndex();
        newSheet = templateSheet.copyTo(activeSpreadsheet);
        activeSpreadsheet.setActiveSheet(newSheet);
        activeSpreadsheet.moveActiveSheet(sheetIndex);
        activeSpreadsheet.deleteSheet(existingSheet);
        newSheet.setName(sheet.name);
      } else {
        functionErrors.push('No existe una hoja llamada ' + sheet.name + ' en el documento actual, por lo que no se ha sobrescrito.');
      }
    } else {
      templateSheet.copyTo(activeSpreadsheet);
    }
  }

  [...rowData.secondarySheets, ...rowData.masterSheets].forEach(processSheet);
  activeSheet.activate();

  if (functionErrors.length > 0) {
    var ui = SpreadsheetApp.getUi();
    var message = functionErrors.join('\n\n');
    ui.alert('Advertencia', message, ui.ButtonSet.OK);
  }
}

/**
 * mainGenerateTemplateSheets
 * Genera hojas secundarias en el documento a partir de sus hojas maestras
 */
function mainGenerateTemplateSheets(rowData) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = ss.getActiveSheet();
  Logger.log('INTERFACE RECEIVED DATA: ' + JSON.stringify(rowData));

  var configJSONData = templateSheetConfigObject();

  var filteredSecondarySheets = configJSONData.secondarySheets.filter(function(secondarySheet) {
    var correspondingMasterSheet = configJSONData.masterSheets.find(function(masterSheet) {
      return masterSheet.name === secondarySheet.masterSheet;
    });

    if (!correspondingMasterSheet) return false;

    var masterSheetIsTrue = rowData.masterSheets.some(function(masterSheetData) {
      return masterSheetData.name === correspondingMasterSheet.name && masterSheetData.isTrue;
    });

    if (!masterSheetIsTrue) return false;

    var correspondingRowDataHoja = rowData.secondarySheets.find(function(rowDataHoja) {
      return rowDataHoja.name === secondarySheet.name && rowDataHoja.isTrue;
    });

    return !!correspondingRowDataHoja;
  });

  filteredSecondarySheets.forEach(function(secondarySheet) {
    var masterSheetName = secondarySheet.masterSheet;
    var masterSheetObj = configJSONData.masterSheets.find(function(maestra) {
      return maestra.name === masterSheetName;
    });
    var tabColor = masterSheetObj ? masterSheetObj.tabColor : null;
    var masterSheetConfig = getSettingsByName(configJSONData, masterSheetName);
    var newSheetConfig = getSettingsByName(configJSONData, secondarySheet.name);
    var newSheetName = secondarySheet.name + '_temp';

    var copiedSheet = copyMasterSheetAsSecondary(ss, masterSheetName, newSheetName);
    aplicarConfiguracion(ss, masterSheetName, newSheetName, masterSheetConfig, newSheetConfig);

    if (tabColor) {
      copiedSheet.setTabColor(tabColor);
    }

    var relativePosition = secondarySheet.relativePosition;
    var masterSheetIndex = ss.getSheetByName(masterSheetName).getIndex();
    var candidateSheetIndex = masterSheetIndex;

    var sheetsToCheck = configJSONData.secondarySheets.filter(function(sheet) {
      return sheet.masterSheet === masterSheetName && sheet.relativePosition < relativePosition;
    });

    if (sheetsToCheck.length > 0) {
      var closestSheet = sheetsToCheck.reduce(function(prev, current) {
        return Math.abs(current.relativePosition - relativePosition) < Math.abs(prev.relativePosition - relativePosition) ? current : prev;
      });

      var sheetToMove = ss.getSheetByName(closestSheet.name);
      if (sheetToMove) {
        candidateSheetIndex = sheetToMove.getIndex();
        Logger.log('closestSheet.name: ' + closestSheet.name);
      } else {
        Logger.log('Sheet ' + closestSheet.name + ' does not exist.');
      }
    }

    ss.setActiveSheet(copiedSheet);
    ss.moveActiveSheet(candidateSheetIndex + 1);
  });

  activeSheet.activate();
}

/**
 * aplicarConfiguracion
 * Aplicar la configuración determinada a cada masterSheet o secondarySheet
 */
function aplicarConfiguracion(ss, masterSheetName, secondarySheetName, masterSheetConfig, newSheetConfig) {
  var secondarySheet = ss.getSheetByName(secondarySheetName);

  Logger.log(`CONFIG - masterSheetName: ${masterSheetName}, secondarySheetName: ${secondarySheetName}, masterSheetConfig: ${JSON.stringify(masterSheetConfig)}, newSheetConfig: ${JSON.stringify(newSheetConfig)},`)

  // Aplicar configuraciónNuevaHoja a la hoja secundaria

  var keys = []
  
  for (var key in newSheetConfig) {
    keys.push(key)
    var valor = newSheetConfig[key];
    var rango = masterSheetConfig[key + "-range"];
    var accion = masterSheetConfig[key + "-range-action"];
    try { var rangoValorArray = masterSheetConfig[key + "-range-value"]; } catch (error) { }
    var columnIndexArray;

    switch (accion) {
      case "setValue":
        secondarySheet.getRange(rango).setValue(valor);
        break;
      case "setBoolean":
        secondarySheet.getRange(rango).setValue(valor);
        break;
      case "deleteRange":
        columnIndexArray = getColumnsRangeObject(rango, true);
        Logger.log(JSON.stringify(`columnIndexArray/deleteRange: ${columnIndexArray}`));
        deleteColumnsFromIndices(secondarySheet, columnIndexArray);
        break;
      case "deleteColumnGroup":
        columnIndexArray = getColumnsRangeObject(rango, true);
        Logger.log(JSON.stringify(`columnIndexArray/deleteColumnGroup: ${columnIndexArray}`));
        deleteColumnGroup(secondarySheet, columnIndexArray);
        break;
      case "addColumnGroup":
        columnIndexArray = getColumnsRangeObject(rango, false);
        Logger.log(JSON.stringify(`columnIndexArray/addColumnGroup: ${columnIndexArray}`));
        addColumnGroup(secondarySheet, columnIndexArray);
        break;
      case "resizeColumns":
        columnIndexArray = getColumnsRangeObject(rango, true);
        Logger.log(JSON.stringify(`columnIndexArray/resizeColumns: ${columnIndexArray}`));
        resizeColumns(secondarySheet, columnIndexArray, rangoValorArray);
        break;
      default:
        break;
    }
  }
  Logger.log(`actionKeys for ${secondarySheetName}: ${keys}`)
}

/**
 * getSettingsByName
 * Obtiene el objeto de configuración de cada hoja
 */
function getSettingsByName(configJSONData, name) {
  const sheetConfig = configJSONData.masterSheets.find(sheet => sheet.name === name) ||
                      configJSONData.secondarySheets.find(sheet => sheet.name === name);
  
  return sheetConfig ? sheetConfig.settings : null;
}

/**
 * copyMasterSheetAsSecondary
 * Copia la hoja maestra y la renombra como secundaria
 */
function copyMasterSheetAsSecondary(ss, masterSheetName, newSheetName) {
  var masterSheet = ss.getSheetByName(masterSheetName);
  var newSheet = masterSheet.copyTo(ss);
  newSheet.setName(newSheetName);
  return newSheet;
}

/**
 * getColumnsRangeObject
 * Devuelve un array de rangos sobre los que aplicar una acción
 */
function getColumnsRangeObject(rangesObject, convertToIndex) {
  columnIndexArray = [];

  for (var key in rangesObject) {
    var value = rangesObject[key];
    var startColumn;
    var endColumn;

    if (key === "range") {
      var rangeParts = value.split(":");
      if (rangeParts.length === 2) {

        if (convertToIndex === true) {
          startColumn = getColumnIndexFromLetter(rangeParts[0]);
          endColumn = rangeParts[1]; // Almacenar el valor para comprobar si es un número o letra

          if (!isNaN(parseInt(endColumn))) {
            // Si es un número, mantener el mismo valor
            var endColumn = parseInt(endColumn);
          } else {
            // Si es una letra, calcular la diferencia de columnas
            var endColumn = getColumnIndexFromLetter(endColumn) - startColumn + 1;
          }
        } else {
          startColumn = rangeParts[0];
          endColumn = rangeParts[1];
        }
        columnIndexArray.push(`${startColumn}:${endColumn}`);
      }
    } else if (key === "multiple") {
      var letters = value.split(",");
      letters.forEach(letter => {
        columnIndexArray.push(getColumnIndexFromLetter(letter));
      });
    }
  }
  return columnIndexArray;
}


function getColumnIndexFromLetter(letter) {
  return letter.charCodeAt(0) - "A".charCodeAt(0) + 1;
}

/**
 * deleteColumnsFromIndices
 * Borra columnas de una hoja a través del getColumnsRangeObject
 */
function deleteColumnsFromIndices(secondarySheet, columnIndexArray) {

  columnIndexArray.forEach(item => {
    if (item.includes(":")) {
      var [startColumn, numColumns] = item.split(":").map(value => parseInt(value, 10));
      secondarySheet.deleteColumns(startColumn, numColumns);
    } else if (!isNaN(item)) {
      var columnIndex = parseInt(item, 10);
      secondarySheet.deleteColumn(columnIndex);
    }
  });
}

function deleteColumnGroup(sheet, columnIndexArray) {

  columnIndexArray.forEach(item => {
    if (item.includes(":")) {
      var [startColumn, columnGroupShift] = item.split(":").map(value => parseInt(value, 10));
      sheet.getColumnGroup(startColumn, columnGroupShift).remove();
    }
  });
}

function addColumnGroup(sheet, columnIndexArray) {

  columnIndexArray.forEach(item => {
    sheet.getRange(item).shiftRowGroupDepth(1);
    if (item.includes(":")) {
      var [startColumn, columnGroupShift] = item.split(":").map(value => getColumnIndexFromLetter(value));
      sheet.getColumnGroup(startColumn, 1).collapse();
    }
  });
}

function resizeColumns(sheet, columnIndexArray, resizeColumnsRangeValue) {
  columnIndexArray.forEach((item, index) => {
    if (!isNaN(item)) {
      var columnIndex = parseInt(item, 10);
      var width = resizeColumnsRangeValue[index];
      sheet.setColumnWidth(columnIndex, width);
    }
  });
}

/**
 * saveDataToDocumentProperties
 * Guarda un objeto rowData en las propiedades del documento
 */
function saveDataToDocumentProperties(rowData) {
  for (let key in rowData) {
    PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(rowData[key]));
  }
}

/**
 * removeDataFromDocumentProperties
 * Elimina un objeto rowData de las propiedades del documento
 */
function removeDataFromDocumentProperties(rowData) {
  var properties = PropertiesService.getDocumentProperties();
  var keysToRemove = Object.keys(properties.getProperties());
  var rowDataKeys = Object.keys(rowData);

  keysToRemove.forEach(function(key) {
    if (rowDataKeys.indexOf(key) > -1) {
      properties.deleteProperty(key);
    }
  });

  var script = `
    <script>
      google.script.run.handlePropertiesRemoved();
    </script>
  `;

  return script;
}

/**
 * templateSheetConfigObject
 * Ofrece el objeto JSON de configuración de plantilla
 */
function templateSheetConfigObject() {

  var templateSheetConfigObject = {
    "projectTypes": [
      {
        "nombre": "Plurifamiliar",
        "code": "plurif",
        "settings": { }
      },
      {
        "nombre": "Unifamiliar",
        "code": "unifam",
        "settings": { }
      },
      {
        "nombre": "Torre de viviendas",
        "code": "torrev",
        "settings": { }
      },
      {
        "nombre": "Hospitality",
        "code": "hospit",
        "settings": { }
      },
      {
        "nombre": "Comercial",
        "code": "comerc",
        "settings": { }
      },
    ],
    "masterSheets": [
      {
        "name": "CHIVATOS",
        "description": "Hoja resumen de los principales chivatos del cuadro",
        "tabColor": "#ff0000",
        "settings": {
        }
      },
      {
        "name": "Construida_Edificable",
        "tabColor": "#3c78d8",
        "settings": {
        }
      },
      {
        "name": "SUPER-M",
        "tabColor": "#3c78d8",
        "settings": {
          "portalBoolean-range": "D2",
          "portalBoolean-range-action": "setBoolean",
          "deleteColumnGroup-range": { "range": "A:1", },
          "deleteColumnGroup-range-action": "deleteColumnGroup",
          "addColumnGroup-range": { "range": "A:C", },
          "addColumnGroup-range-action": "addColumnGroup",
          "resizeColumns-range": { "multiple": "J", },
          "resizeColumns-range-value": [50],
          "resizeColumns-range-action": "resizeColumns",
        }
      },
      {
        "name": "SUPER-S",
        "tabColor": "#3c78d8",
        "settings": {
          "portalBoolean-range": "D2",
          "portalBoolean-range-action": "setBoolean",
          "deleteColumnGroup-range": { "range": "A:1", },
          "deleteColumnGroup-range-action": "deleteColumnGroup",
          "addColumnGroup-range": { "range": "A:C", },
          "addColumnGroup-range-action": "addColumnGroup",
          "resizeColumns-range": { "multiple": "O", },
          "resizeColumns-range-value": [50],
          "resizeColumns-range-action": "resizeColumns",
        }
      },
      {
        "name": "Aparcamiento_Trastero",
        "tabColor": "#3c78d8",
        "settings": {
        }
      },
      {
        "name": "SUPER PLANTAS",
        "tabColor": "#ffd966",
        "settings": {
          "portal-range": "D2",
          "portal-range-action": "setBoolean",
        }
      },
      {
        "name": "SUPER ROSAS",
        "tabColor": "#ff62ff",
        "settings": {
          "type-range": "A5",
          "type-range-action": "setValue",
        }
      },
      {
        "name": "SI SECTORES Y LRE",
        "tabColor": "#cc4125",
        "settings": {
        }
      },
      {
        "name": "SI OCUPACIÓN",
        "tabColor": "#cc4125",
        "settings": {
        }
      },
      {
        "name": "CÓMPUTO VIV",
        "tabColor": "#6aa84f",
        "settings": {
        }
      },
      {
        "name": "SUPER PLAZAS",
        "tabColor": "#6aa84f",
        "settings": {
          "type-range": "A2",
          "type-range-action": "setValue",
        }
      },
      {
        "name": "Justificación ventanas",
        "tabColor": "#6aa84f",
        "settings": {
        }
      },
      {
        "name": "Justificación FT",
        "tabColor": "#6aa84f",
        "settings": {
        }
      },
      {
        "name": "SUP AUTO",
        "tabColor": "#6aa84f",
        "settings": {
        }
      },
      {
        "name": "Cuadros justificativos AYTO",
        "tabColor": "#b45f06",
        "settings": {
        }
      },
      {
        "name": "AYTO CONST USO",
        "tabColor": "#b45f06",
        "settings": {
          "type-range": "A1",
          "type-range-action": "setValue",
        }
      },
      {
        "name": "AYTO MADRID",
        "tabColor": "#b45f06",
        "settings": {
          "portal-range": "B1",
          "portal-range-action": "setValue",
          "srbr-range": "D1",
          "srbr-range-action": "setValue",
          "use-range": "C4",
          "use-range-action": "setBoolean",
          "portalresume-range": { "range": "T:Y", },
          "portalresume-range-action": "deleteRange"
        }
      },
      {
        "name": "IT 03 CB Cuarto Basuras",
        "tabColor": "#b45f06",
        "settings": {
        }
      },
      {
        "name": "Histórico CONST",
        "tabColor": "#9900ff",
        "settings": {
        }
      },
    ],
    "secondarySheets": [
      {
        "name": "Portal-Letra M",
        "description": "Cuadro de superficies M desglosado por bloques",
        "masterSheet": "SUPER-M",
        "relativePosition": 1,
        "settings": {
          "portalBoolean": true,
          "deleteColumnGroup": true,
          "addColumnGroup": true,
          "resizeColumns": true
        }
      },
      {
        "name": "Portal-Letra S",
        "description": "Cuadro de superficies S desglosado por bloques",
        "masterSheet": "SUPER-S",
        "relativePosition": 1,
        "settings": {
          "portalBoolean": true,
          "deleteColumnGroup": true,
          "addColumnGroup": true,
          "resizeColumns": true
        }
      },
      {
        "name": "PLANTAS SR",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 1,
        "settings": {
        }
      },
      {
        "name": "PLANTAS BR",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 2,
        "settings": {
        }
      },
      {
        "name": "PLANTAS URB",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 3,
        "settings": {
        }
      },
      {
        "name": "ZZCC SR",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 4,
        "settings": {
        }
      },
      {
        "name": "CONST SR",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 5,
        "settings": {
        }
      },
      {
        "name": "COMP SR",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 6,
        "settings": {
        }
      },
      {
        "name": "REP SR",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 7,
        "settings": {
        }
      },
      {
        "name": "CONST SR-BR",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 8,
        "settings": {
        }
      },
      {
        "name": "COMP SR-BR",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 9,
        "settings": {
        }
      },
      {
        "name": "REP SR-BR",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 10,
        "settings": {
        }
      },
      {
        "name": "CONST SR BLOQUES",
        "masterSheet": "SUPER PLANTAS",
        "relativePosition": 11,
        "settings": {
        }
      },
      {
        "name": "RESUMEN VIV TIPO",
        "masterSheet": "SUPER ROSAS",
        "relativePosition": 1,
        "settings": {
          "type": "RESUMEN VIV TIPO",
        }
      },
      {
        "name": "ÚTIL VIV TIPO interior",
        "masterSheet": "SUPER ROSAS",
        "relativePosition": 2,
        "settings": {
          "type": "ÚTIL VIV TIPO interior",
        }
      },
      {
        "name": "ÚTIL VIV SUBTIPO",
        "masterSheet": "SUPER ROSAS",
        "relativePosition": 3,
        "settings": {
          "type": "ÚTIL VIV SUBTIPO",
        }
      },
      {
        "name": "ÚTIL VIVIENDAS BLOQUE",
        "masterSheet": "SUPER ROSAS",
        "relativePosition": 4,
        "settings": {
        }
      },
      {
        "name": "Cuadro trasteros",
        "masterSheet": "SUPER PLAZAS",
        "relativePosition": 1,
        "settings": {
          "type": "Trasteros",
        }
      },
      {
        "name": "AYTO COMP USO",
        "masterSheet": "AYTO CONST USO",
        "relativePosition": 1,
        "settings": {
          "type": "COMPUTABLE POR USOS",
        }
      },
      {
        "name": "AYTO MADRID SR",
        "masterSheet": "AYTO MADRID",
        "relativePosition": 1,
        "settings": {
          "portal": "TODO",
          "srbr": "SR",
          "use": true,
          "portalresume": false
        }
      },
      {
        "name": "AYTO MADRID BR",
        "masterSheet": "AYTO MADRID",
        "relativePosition": 2,
        "settings": {
          "portal": "TODO",
          "srbr": "BR",
          "use": true,
          "portalresume": false
        }
      },
      {
        "name": "AYTO MADRID Portal [x]",
        "masterSheet": "AYTO MADRID",
        "relativePosition": 3,
        "settings": {
        }
      },
      {
        "name": "AYTO MADRID Comercial",
        "masterSheet": "AYTO MADRID",
        "relativePosition": 4,
        "settings": {
        }
      },
      {
        "name": "Histórico CONST Resumen",
        "masterSheet": "Histórico CONST",
        "relativePosition": 1,
        "settings": {
        }
      },
    ]
  }

  return templateSheetConfigObject;
}
