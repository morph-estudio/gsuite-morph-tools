/**
 * Gsuite Morph Tools - Morph Spreadsheet Configurator
 * Developed by alsanchezromero
 *
 * Morph Estudio, 2023
 */

function mscPruebas() {
  Logger.log(getColumnsRangeObject({ "multiple": "AE:BM,CI:CX" }, true))
}

/**
 * Copia y configura un cuadro de superficies en la carpeta seleccionada.
 *
 * @param {Object} rowData - Los datos que se utilizarán para determinar la copia del cuadro.
 */
function crearCuadroSuperficies(rowData) {
  
  var [projectCategory, folderURL, csFolderURL, pcFolderURL, cuadroType, cuadroSoftware, projectShortname, projectCode, projectLocalidad, cuadroAmpliadoSwitch] = [rowData.projectCategory, rowData.cuadroFolder, rowData.cuadroFolderSup, rowData.cuadroFolderPanel, rowData.cuadroType, rowData.cuadroSoftware, rowData.projectShortname, rowData.projectCode, rowData.projectLocalidad, rowData.cuadroAmpliadoSwitch];

  var cuadroTypeName;
  var cuadroTypeID;
  var includedSheetNames;
  var destinationFolder;

  var idArray = {};
  var urlArray = {};

  const cuadroTypes = getCuadroTypes(); Logger.log(`cuadroTypes: ${cuadroTypes}`);

  Logger.log(`cuadroType: ${cuadroType}, projectCategory: ${projectCategory}, projectCode: ${projectCode}, projectShortname: ${projectShortname}, folderURL: ${folderURL}`);

  // Comprobaciones iniciales 

  var isExpediente = cuadroType === "selectExpedienteCompleto" ? true : false;
  var cuadrosArray = isExpediente === true ? ["selectPanelControl", "selectCuadroSuperficies", "selectCuadroExportacion"] : [cuadroType];

  cuadrosArray.forEach(function(cuadro) {

    const cuadroTypeName = getTipoArchivo(null, cuadroTypes[cuadro]);
    if (cuadroTypeName) {
      cuadroTypeID = naveNodrizaIDS(cuadroTypeName);
      folderByCuadro = cuadroTypes[cuadro] === 1 ? pcFolderURL : csFolderURL;
      destinationFolder = isExpediente ? folderByCuadro : folderURL;
    }

    try {
      destinationFolder = DriveApp.getFolderById(getIdFromUrl(destinationFolder));
    } catch (error) {
        throw new Error('No se ha encontrado la carpeta especificada.');
    }

    // Traer la configuración del objeto JSON

    var configJSONData = templateSheetConfigObject(false, cuadroTypeName);
    var projectTypeConfig = configJSONData.projectTypes.find((type) => type.code === projectCategory);

    if (!projectTypeConfig) { throw new Error('No se ha encontrado una configuración para este tipo de proyecto.'); }

    // Generar la copia de plantilla

    var newFilename = `${projectCode} ${cuadroTypeName} ${projectShortname}`;
    var newFile = DriveApp.getFileById(cuadroTypeID).makeCopy(newFilename, destinationFolder);
    var newFileID = newFile.getId();

    idArray[cuadroTypeName] = newFileID;
    urlArray[cuadroTypeName] = newFile.getUrl();
  });

  cuadrosArray.forEach(function(cuadro) {

    Logger.log(`cuadroTypes[cuadro]: ${cuadroTypes[cuadro]}`);

    cuadroTypeName = getTipoArchivo(null, cuadroTypes[cuadro] || null);

    var cuadroFileID = idArray[cuadroTypeName];
    var archivoCopiado = SpreadsheetApp.openById(cuadroFileID);

    Logger.log(`Tipo de cuadro: ${cuadroTypeName}, ID Archivo: ${cuadroFileID}, URL Cuadro: ${urlArray[cuadroTypeName]}, ¿Es ampliado?: ${cuadroAmpliadoSwitch}`);

    var configJSONData = templateSheetConfigObject(false, cuadroTypeName);
    var projectTypeConfig = configJSONData.projectTypes.find(
      (type) => type.code === projectCategory
    );
    
    // Eliminar los rangos nombrados en función del tipo de cuadro

    if (cuadro === "selectCuadroSuperficies" && cuadroSoftware !== "Revit") {
      eliminarRangosNombradosPorPrefijo(archivoCopiado, "RVT");
    }

    SpreadsheetApp.flush();

    // Hojas excluidas

    var includedMasterSheetsCodes;

    includedMasterSheetsCodes = cuadroAmpliadoSwitch ? projectTypeConfig.settings.fullCuadro : projectTypeConfig.settings.minimumCuadro;

    if (cuadroSoftware === "Revit") {
      var revitSheetsCodes = projectTypeConfig.settings.revitCuadro;
      Object.assign(includedMasterSheetsCodes, revitSheetsCodes);
    }

    includedSheetNames = new Set(
      includedMasterSheetsCodes.map((code) => {
        var masterSheet = configJSONData.masterSheets.find((sheet) => sheet.code === code);
        return masterSheet ? masterSheet.name : '';
      })
    );

    Logger.log(`includedSheetNames: ${JSON.stringify(includedSheetNames)}`)

    if (includedSheetNames.size > 0) {
      var hojasArchivoCopiado = archivoCopiado.getSheets();
      var sheetsToDelete = hojasArchivoCopiado.filter((sheet) => {
        var sheetName = sheet.getName();
        return !includedSheetNames.has(sheetName);
      });

      var requests = sheetsToDelete.map((sheet) => {
        return {
          deleteSheet: {
            sheetId: sheet.getSheetId(),
          },
        };
      });

      if (requests.length > 0) {
        Sheets.Spreadsheets.batchUpdate({ requests }, archivoCopiado.getId());
      }
    }

    SpreadsheetApp.flush();

    // Ocultar las hojas definidas en "hiddenSheets"

    var hiddenSheetCodes = projectTypeConfig.settings.hiddenSheets || [];
    var hiddenSheetNames = hiddenSheetCodes.map((code) => {
      var sheet = configJSONData.masterSheets.find((masterSheet) => masterSheet.code === code);
      return sheet ? sheet.name : null;
    });

    hiddenSheetNames.forEach(function(sheetName) {
      if (sheetName) {
        var sheetToHide = archivoCopiado.getSheetByName(sheetName);
        if (sheetToHide) {
          sheetToHide.hideSheet();
        }
      }
    });

    // Configuración específica de la plantilla

    var instruccionesSheet; var linkRange;

    switch (cuadro) {
      case 'selectPanelControl':

        instruccionesSheet = archivoCopiado.getSheetByName('Instrucciones');

        var linkRange = 'B3';
        if (isExpediente) {
          instruccionesSheet.getRange(linkRange).setValue(urlArray['Cuadro Superficies']);
        }
        
        var values = instruccionesSheet.getRange("A1:B" + instruccionesSheet.getLastRow()).getValues();
        var found;

        for (var i = 0; i < values.length; i++) {
          var cellA = values[i][0]; // Valor de la columna A
          var cellB = values[i][1]; // Valor de la columna B

          found = false;

          if (cellA === 'Expediente') {
            cellB = projectCode; found = true;
          } else if (cellA === 'Título abreviado') {
            cellB = projectShortname; found = true;
          } else if (cellA === 'Localidad') {
            cellB = projectLocalidad; found = true;
          } else { cellB = null; }
          
          if (found) { instruccionesSheet.getRange(i + 1, 2, 1, 1).setValue(cellB); }
        }

        let BDDPC_ID = '1xWhvOUZPGgkRr8Emtt2lm4n6U67VCiO4h1PqpX72Dbs';
        importRangeToken(idArray['Panel de control'], idArray['Cuadro Superficies']);
        importRangeToken(idArray['Panel de control'], BDDPC_ID);


        break;
      case 'selectCuadroSuperficies':

        instruccionesSheet = archivoCopiado.getSheetByName('LINK');
        var linkRange = 'B1';
        if (isExpediente) {
          instruccionesSheet.getRange(linkRange).setValue(urlArray['Panel de control']);
        } else {
          instruccionesSheet.getRange(linkRange).clearContent();
        }

        let BDDCS_ID = '1P5R7Gw22DeTjyaCHLqXngQDYA5WdenVGYMPoaYM1OrQ';
        importRangeToken(idArray['Cuadro Superficies'], BDDCS_ID);
        importRangeToken(idArray['Cuadro Superficies'], idArray['Panel de control']);

        break;
      case 'selectCuadroExportacion':

        var instruccionesSheet = archivoCopiado.getSheetByName('LINK');
        var linkRange = 'B2:B3';
        if (isExpediente) {
          var cuadroSuperficiesValue = [urlArray['Cuadro Superficies']];
          var panelControlValue = [urlArray['Panel de control']];
          instruccionesSheet.getRange(linkRange).setValues([cuadroSuperficiesValue, panelControlValue]);
        } else {
          instruccionesSheet.getRange(linkRange).clearContent();
        }

        importRangeToken(idArray['Exportación Superficies'], idArray['Panel de control']);
        importRangeToken(idArray['Exportación Superficies'], idArray['Cuadro Superficies']);

        break;
      case 'selectCuadroMediciones':

        var instruccionesSheet = archivoCopiado.getSheetByName('LINK');
        var headers = instruccionesSheet.getRange(1, 1, 1, instruccionesSheet.getLastColumn()).getValues()[0];
        var columnIndex = headers.indexOf("AC: URL");
        if (columnIndex !== -1) {
          var dataColumn = instruccionesSheet.getRange(2, columnIndex + 1, instruccionesSheet.getLastRow() - 1, 1).getValues();
          var bddidsarray = [];

          // Iterar a través de los valores de la columna "AC:URL"
          for (var i = 0; i < dataColumn.length; i++) {
            var url = dataColumn[i][0];
            if (url !== "") {
              var bddID = getIdFromUrl(url);
              bddidsarray.push(bddID);
            }
          }

          for (var u = 0; u < bddidsarray.length; u++) {
            bddID = bddidsarray[u];
            importRangeToken(idArray['Cuadro Mediciones'], bddID);
          }
        } else {
          Logger.log("La columna 'AC:URL' no se encontró en la hoja.");
        }

        break;
      default:
        break;
    }

  });

}






















function vincularHojaExportacion(rowData, overwriteSwitch, rowDataExportacion) {
  Logger.log(`overwriteSwitch: ${overwriteSwitch}, selectPanelType: ${selectPanelType}, selectPanelTypeSheets: ${selectPanelTypeSheets}`)
  var [selectPanelType, selectPanelTypeSheets] = [rowDataExportacion.selectPanelType, rowDataExportacion.selectPanelTypeSheets];

  var panelID = naveNodrizaIDS(selectPanelType);

  switch (selectPanelType) {
    case 'naveNodrizaSuperficies':
      var string = 'Cuadro Superficies';
      break;
    case 'naveNodrizaPanelControl':
      var string = 'Panel de control';
      break;
    default:
      break;
  }
  
  
  // Accede al documento con el ID panelID
  var ss = SpreadsheetApp.getActive();
  var panelDoc = SpreadsheetApp.openById(panelID);
  var panelDocURL = panelDoc.getUrl();

  // Obtiene la hoja actual en el documento actual

  // Verifica si ya existe una hoja con el mismo nombre en el documento panelDoc
  var sheetExists = ss.getSheetByName(selectPanelTypeSheets) !== null;

  if (sheetExists && !overwriteSwitch) {
    Logger.log("Advertencia: La hoja con el mismo nombre ya existe en el documento destino. No se sobrescribirá.");
    return;
  }

  var sheetToCopy = panelDoc.getSheetByName(selectPanelTypeSheets);

  // Copia la hoja actual al documento panelDoc
  var copiedSheet = sheetToCopy.copyTo(ss);
  
  // Cambia el nombre de la hoja copiada al mismo nombre que en el documento panelDoc
  copiedSheet.setName(selectPanelTypeSheets);

  // Elimina el contenido de la hoja copiada
  copiedSheet.clearContents();
  copiedSheet.clearNotes();
  copiedSheet.getRange(1, 1, copiedSheet.getMaxRows(), copiedSheet.getMaxColumns()).clearDataValidations();

  // En la celda A1 de la hoja copiada, coloca una fórmula con importrange
  var formula = `=IMPORTRANGE(LINK!$B$2;"${selectPanelTypeSheets}!A1:AAZ")`;
  copiedSheet.getRange("A1").setFormula(formula);



  var configObject = templateSheetConfigObject(false, string);

  var hojaEncontrada = configObject.masterSheets.find(function(sheet) {
    return sheet.name === selectPanelTypeSheets;
  });

  // Si no se encuentra en las hojas maestras, busca en las hojas secundarias
  if (!hojaEncontrada) {
    hojaEncontrada = configObject.secondarySheets.find(function(sheet) {
      return sheet.name === selectPanelTypeSheets;
    });
  }

  // Verifica si se encontró la hoja y si tiene la clave "filaFormulas" en "settings"
  if (hojaEncontrada && hojaEncontrada.settings && hojaEncontrada.settings.filaFormulas === true) {
    // Oculta la fila 3 de la hoja actual
    copiedSheet.hideRows(3);
  }



}





/**
 * Copia hojas desde una plantilla al documento actual.
 *
 * @param {Object} rowData - Los datos que se utilizarán para determinar qué hojas copiar.
 * @param {boolean} overwriteSwitch - Un interruptor booleano que indica si se deben sobrescribir las hojas existentes.
 */
function mainCopySheetFromTemplate(rowData, overwriteSwitch, templateFile) {
  
  var plantillaID = naveNodrizaIDS(templateFile.toString().trim());
  Logger.log(`overwriteSwitch: ${overwriteSwitch}, templateFile: ${templateFile}, plantillaID: ${plantillaID}`)
  var templateSpreadsheet = SpreadsheetApp.openById(plantillaID);
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

    var existingSheet = activeSpreadsheet.getSheetByName(sheet.name);

    if (overwriteSwitch) {

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
      if (existingSheet) {
        templateSheet.copyTo(activeSpreadsheet);
      } else {
        newSheet = templateSheet.copyTo(activeSpreadsheet);
        newSheet.setName(sheet.name);
      }
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
 * Genera hojas secundarias en el documento a partir de sus hojas maestras.
 *
 * @param {Object} rowData - Los datos que se utilizarán para determinar qué hojas copiar.
 * @param {boolean} ovewriteSwitch - Un interruptor booleano que indica si se deben sobrescribir las hojas existentes.
 */
function mainGenerateTemplateSheets(rowData, overwriteSwitch, templateType) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var configJSONData = templateSheetConfigObject(true);

  Logger.log(`templateType: ${templateType}, INTERFACE RECEIVED DATA: ${JSON.stringify(rowData)}, INTERFACE configJSONData: ${JSON.stringify(configJSONData)}`);

  var filteredSecondarySheets = configJSONData.secondarySheets.filter(function (secondarySheet) {
    var correspondingMasterSheet = configJSONData.masterSheets.find(function (masterSheet) {
      return masterSheet.name === secondarySheet.masterSheet;
    });

    if (!correspondingMasterSheet) return false;

    var masterSheetIsTrue = rowData.masterSheets.some(function (masterSheetData) {
      return masterSheetData.name === correspondingMasterSheet.name && masterSheetData.isTrue;
    });

    if (!masterSheetIsTrue) return false;

    var correspondingRowDataHoja = rowData.secondarySheets.find(function (rowDataHoja) {
      return rowDataHoja.name === secondarySheet.name && rowDataHoja.isTrue;
    });

    return !!correspondingRowDataHoja;
  });

  if (filteredSecondarySheets.length === 0) {
    throw new Error('No se ha seleccionado ninguna hoja secundaria.');
  }

  filteredSecondarySheets.forEach(function (secondarySheet) {
    var masterSheetName = secondarySheet.masterSheet;
    var masterSheetObj = configJSONData.masterSheets.find(function (maestra) {
      return maestra.name === masterSheetName;
    });
    var tabColor = masterSheetObj ? masterSheetObj.tabColor : null;

    var masterSheetConfig = configJSONData.masterSheets.find(function (masterSheet) {
      return masterSheet.name === masterSheetName;
    });

    var newSheetConfig = configJSONData.secondarySheets.find(function (sheet) {
      return sheet.name === secondarySheet.name;
    });

    var newSheetName = secondarySheet.name;

    // Verificar si la hoja ya existe en el documento
    var existingSheet = ss.getSheetByName(newSheetName);
    if (existingSheet) {
      if (overwriteSwitch) {
        ss.deleteSheet(existingSheet);
      } else {
        var response = SpreadsheetApp.getUi().alert(
          'Advertencia: la hoja ya existe',
          'La hoja "' + newSheetName + '" ya existe en el documento, por lo que no se generará. Para sobrescribir la hoja por una nueva, debes seleccionar la opción "Sobrescribir hojas" antes de ejecutar.',
          SpreadsheetApp.getUi().ButtonSet.OK
        );

        if (response == SpreadsheetApp.getUi().Button.OK) {
          // El usuario seleccionó "No", se salta esta iteración del bucle
          return;
        }
      }
    }

    var copiedSheet = copyMasterSheetAsSecondary(ss, masterSheetName, newSheetName);
    aplicarConfiguracion(ss, masterSheetName, newSheetName, masterSheetConfig.settings, newSheetConfig.settings);

    if (tabColor) {
      copiedSheet.setTabColor(tabColor);
    }

    if (templateType != "Exportación Superficies") {
      var relativePosition = secondarySheet.relativePosition;
      var masterSheetIndex = ss.getSheetByName(masterSheetName).getIndex();
      var candidateSheetIndex = masterSheetIndex;

      var sheetsToCheck = configJSONData.secondarySheets.filter(function (sheet) {
        return sheet.masterSheet === masterSheetName && sheet.relativePosition < relativePosition;
      });

      if (sheetsToCheck.length > 0) {
        var closestSheet = sheetsToCheck.reduce(function (prev, current) {
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

      SpreadsheetApp.flush()

      ss.setActiveSheet(copiedSheet);
      ss.moveActiveSheet(candidateSheetIndex + 1);
    }
  });
}

/**
 * Aplica la configuración determinada a cada masterSheet o secondarySheet.
 *
 * @param {Spreadsheet} ss - El objeto Spreadsheet activo al que se aplicará la configuración.
 * @param {string} masterSheetName - El nombre de la hoja maestra.
 * @param {string} secondarySheetName - El nombre de la hoja secundaria.
 * @param {Object} masterSheetConfig - La configuración de la hoja maestra.
 * @param {Object} newSheetConfig - La configuración de la hoja secundaria.
 */
function aplicarConfiguracion(ss, masterSheetName, secondarySheetName, masterSheetConfig, newSheetConfig) {
  var secondarySheet = ss.getSheetByName(secondarySheetName);

  Logger.log(`CONFIG - masterSheetName: ${masterSheetName}, secondarySheetName: ${secondarySheetName}, masterSheetConfig: ${JSON.stringify(masterSheetConfig)}, newSheetConfig: ${JSON.stringify(newSheetConfig)}`);

  // Aplicar configuraciónNuevaHoja a la hoja secundaria
  for (var key in masterSheetConfig) {
    var masterSetting = masterSheetConfig[key];
    var newSetting = newSheetConfig[key];
    if (newSetting == undefined) continue;
    var action = masterSetting.action;
    var range = masterSetting.range;
    var value = newSetting.value;
    var columnIndexArray;

    switch (action) {
      case "setValue":
        if (range.includes(';')) {
          var headerParts = range.split(";");
          var columnHeader = headerParts[0];
          var rowNumber = parseInt(headerParts[1]);

          var headerRow = secondarySheet.getRange(1, 1, 1, secondarySheet.getLastColumn()).getValues()[0];
          var columnIndex = headerRow.indexOf(columnHeader) + 1;

          // Verificar si se encontró el encabezado
          if (columnIndex > 0) {
            secondarySheet.getRange(rowNumber, columnIndex).setValue(value);
          } else {
            Logger.log("El encabezado '" + columnHeader + "' no se encontró en la hoja.");
          }
        } else { secondarySheet.getRange(range).setValue(value); }
        break;
      case "setBoolean":
        if (range.includes(';')) {
          var headerParts = range.split(";");
          var columnHeader = headerParts[0];
          var rowNumber = parseInt(headerParts[1]);

          var headerRow = secondarySheet.getRange(1, 1, 1, secondarySheet.getLastColumn()).getValues()[0];
          var columnIndex = headerRow.indexOf(columnHeader) + 1;

          // Verificar si se encontró el encabezado
          if (columnIndex > 0) {
            secondarySheet.getRange(rowNumber, columnIndex).setValue(value);
          } else {
            Logger.log("El encabezado '" + columnHeader + "' no se encontró en la hoja.");
          }
        } else { secondarySheet.getRange(range).setValue(value); }
        break;
      case "deleteRange":
        columnIndexArray = getColumnsRangeObject(range, true);
        Logger.log(`columnIndexArray/deleteRange: ${JSON.stringify(columnIndexArray)}`);
        deleteColumnsFromIndices(secondarySheet, columnIndexArray);
        break;
      case "deleteColumnGroup":
        columnIndexArray = getColumnsRangeObject(range, true);
        Logger.log(`columnIndexArray/deleteColumnGroup: ${JSON.stringify(columnIndexArray)}`);
        deleteColumnGroup(secondarySheet, columnIndexArray);
        break;
      case "addColumnGroup":
        columnIndexArray = getColumnsRangeObject(range, false);
        Logger.log(`columnIndexArray/addColumnGroup: ${JSON.stringify(columnIndexArray)}`);
        addColumnGroup(secondarySheet, columnIndexArray);
        break;
      case "collapseColumnGroup":
        columnIndexArray = getColumnsRangeObject(range, false);
        Logger.log(`columnIndexArray/addColumnGroup: ${JSON.stringify(columnIndexArray)}`);
        collapseColumnGroup(secondarySheet, columnIndexArray);
        break;
      case "resizeColumns":
        columnIndexArray = getColumnsRangeObject(range, true);
        Logger.log(`columnIndexArray/resizeColumns: ${JSON.stringify(columnIndexArray)}`);
        resizeColumns(secondarySheet, columnIndexArray, value);
        break;
      default:
        break;
    }
  }
  Logger.log(`Applied settings for ${secondarySheetName}`);
  SpreadsheetApp.flush()
}


/**
 * Obtiene el objeto de configuración de una hoja por su nombre.
 *
 * @param {Object} configJSONData - El objeto de datos de configuración que contiene las hojas maestras y secundarias.
 * @param {string} name - El nombre de la hoja para la cual se desea obtener la configuración.
 * @return {Object|null} El objeto de configuración de la hoja o null si no se encuentra.
 */
function getSettingsByName(configJSONData, name) {
  const sheetConfig = configJSONData.masterSheets.find(sheet => sheet.name === name) ||
                      configJSONData.secondarySheets.find(sheet => sheet.name === name);
  
  return sheetConfig ? sheetConfig.settings : null;
}

/**
 * Copia la hoja maestra y la renombra como hoja secundaria.
 *
 * @param {Spreadsheet} ss - El objeto Spreadsheet activo donde se realizará la copia.
 * @param {string} masterSheetName - El nombre de la hoja maestra que se copiará.
 * @param {string} newSheetName - El nombre que se asignará a la nueva hoja secundaria.
 * @return {Sheet} La hoja secundaria recién creada.
 */
function copyMasterSheetAsSecondary(ss, masterSheetName, newSheetName) {
  var masterSheet = ss.getSheetByName(masterSheetName);
  var newSheet = masterSheet.copyTo(ss);
  newSheet.setName(newSheetName);
  return newSheet;
}

////////////////////////////
// HELPER FUNCTIONS
////////////////////////////

/**
 * Devuelve un array de rangos sobre los que aplicar una acción.
 *
 * @param {Object} rangesObject - El objeto que contiene los rangos a procesar.
 * @param {boolean} convertToIndex - Indica si los rangos deben convertirse en índices de columnas.
 * @return {string[]} Un array de rangos en formato de texto o índices de columnas si se ha especificado la conversión.
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
        startColumn = getColumnIndexFromLetter(rangeParts[0]);
        endColumn = getColumnIndexFromLetter(rangeParts[1]);

        if (convertToIndex === true) {
          columnIndexArray.push(`${startColumn}:${endColumn - startColumn + 1}`);
        } else {
          columnIndexArray.push(`${rangeParts[0]}:${rangeParts[1]}`);
        }
      }
    } else if (key === "multiple") {
      var ranges = value.split(",");
      ranges.forEach(range => {
        if (convertToIndex === true) {
          var rangeParts = range.split(":");
          if (rangeParts.length === 2) {
            startColumn = getColumnIndexFromLetter(rangeParts[0]);
            endColumn = getColumnIndexFromLetter(rangeParts[1]);
            columnIndexArray.push(`${startColumn}:${endColumn - startColumn + 1}`);
          }
        } else {
          columnIndexArray.push(range);
        }
      });
    }
  }
  return columnIndexArray;
}

/**
 * Obtiene el índice de columna a partir de una letra de columna o número de columna.
 *
 * @param {string|number} letterOrNumber - La letra de columna o número de columna a convertir.
 * @return {number} El índice de columna correspondiente.
 */
function getColumnIndexFromLetter(letterOrNumber) {
  if (!isNaN(letterOrNumber)) {
    // Si es un número, devolverlo sin cambios
    return parseInt(letterOrNumber);
  }

  var index = 0;
  for (var i = 0; i < letterOrNumber.length; i++) {
    index = index * 26 + (letterOrNumber.charCodeAt(i) - "A".charCodeAt(0) + 1);
  }
  return index;
}

////////////////////////////
// MORPH SPREADSHEET CONFIGURATOR ACTIONS
////////////////////////////

/**
 * Borra columnas de una hoja utilizando los índices de columna especificados.
 *
 * @param {Sheet} secondarySheet - La hoja en la que se borrarán las columnas.
 * @param {number[]} columnIndexArray - Un array de índices de columna a borrar.
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

/**
 * Agrega un grupo de columnas en una hoja utilizando los índices de columna especificados.
 *
 * @param {Sheet} sheet - La hoja en la que se agregará el grupo de columnas.
 * @param {string[]} columnIndexArray - Un array de índices de columna que representan el grupo a agregar.
 */
function addColumnGroup(sheet, columnIndexArray) {

  columnIndexArray.forEach(item => {
    sheet.getRange(item).shiftRowGroupDepth(1);
    if (item.includes(":")) {
      var [startColumn, columnGroupShift] = item.split(":").map(value => getColumnIndexFromLetter(value));
      sheet.getColumnGroup(startColumn, 1);
      // sheet.getColumnGroup(startColumn, 1).collapse();
    }
  });
}

/**
 * Elimina un grupo de columnas en una hoja utilizando los índices de columna especificados.
 *
 * @param {Sheet} sheet - La hoja en la que se eliminará el grupo de columnas.
 * @param {string[]} columnIndexArray - Un array de índices de columna que representan el grupo a eliminar.
 */
function deleteColumnGroup(sheet, columnIndexArray) {

  columnIndexArray.forEach(item => {
    if (item.includes(":")) {
      var [startColumn, columnGroupShift] = item.split(":").map(value => parseInt(value, 10));
      sheet.getColumnGroup(startColumn, columnGroupShift).remove();
    }
  });
}

/**
 * Contrae un grupo de columnas en una hoja utilizando los índices de columna especificados.
 *
 * @param {Sheet} sheet - La hoja en la que se contraerá el grupo de columnas.
 * @param {string[]} columnIndexArray - Un array de índices de columna que representan el grupo a contraer.
 */
function collapseColumnGroup(sheet, columnIndexArray) {

  columnIndexArray.forEach(item => {
    if (item.includes(":")) {
      var [startColumn, columnGroupShift] = item.split(":").map(value => getColumnIndexFromLetter(value));
      sheet.getColumnGroup(startColumn, columnGroupShift).collapse();
    }
  });
}

/**
 * Cambia el ancho de columnas en una hoja utilizando los índices de columna y los anchos especificados.
 *
 * @param {Sheet} sheet - La hoja en la que se cambiará el ancho de columnas.
 * @param {number[]} columnIndexArray - Un array de índices de columna a redimensionar.
 * @param {number[]} resizeColumnsRangeValue - Un array de anchos de columna correspondientes.
 */
function resizeColumns(sheet, columnIndexArray, resizeColumnsRangeValue) {
  columnIndexArray.forEach((item, index) => {
    if (!isNaN(item)) {
      var columnIndex = parseInt(item, 10);
      var width = resizeColumnsRangeValue[index];
      sheet.setColumnWidth(columnIndex, width);
    }
  });
}

////////////////////////////
// SAVE AND LOAD FUNCTIONS
////////////////////////////

/**
 * Guarda un objeto rowData en las propiedades del documento.
 *
 * @param {Object} rowData - El objeto de datos a guardar en las propiedades del documento.
 */
function saveDataToDocumentProperties(rowData) {
  for (let key in rowData) {
    PropertiesService.getDocumentProperties().setProperty(key, JSON.stringify(rowData[key]));
  }
}

/**
 * Elimina un objeto rowData de las propiedades del documento y llama a una función adicional después de eliminar las propiedades.
 *
 * @param {Object} rowData - El objeto de datos a eliminar de las propiedades del documento.
 * @return {string} Un fragmento de script HTML que llama a la función "handlePropertiesRemoved" de Google Apps Script.
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


function agruparColumnasDeExportacion() {
  var hoja = SpreadsheetApp.getActiveSheet();
  var numColumnas = hoja.getLastColumn();
  var numRows = hoja.getLastRow();

  // Recorre todas las columnas
  for (var i = numColumnas; i >= 1; i--) {
    var columna = hoja.getRange(1, i, numRows, 1);
    var valores = columna.getValues().flat(); // Obtén los valores de la columna como un array

    // Verifica si la columna contiene valores numéricos
    var contieneValoresNumericos = valores.some(function(valor) {
      return typeof valor === 'number' && !isNaN(valor);
    });

    // Calcula la suma de los valores numéricos en la columna
    var suma = valores.reduce(function(acc, valor) {
      return acc + (typeof valor === 'number' ? valor : 0);
    }, 0);

    var d = hoja.getColumnGroupDepth(i);

    // Verifica si la columna está agrupada
    if (d > 0) {
      // Si la suma no es cero y la columna contiene valores numéricos, desagrupa la columna y expande el grupo
      if (suma !== 0 || contieneValoresNumericos) {
        columna.shiftColumnGroupDepth(0);
      }
    } else {
      // Si la suma es cero y la columna contiene valores numéricos, agrupa la columna con nivel 1 y colapsa la agrupación
      if (suma === 0 && contieneValoresNumericos) {
        columna.shiftColumnGroupDepth(1).collapseGroups();
      }
    }
  }
}

function eliminarRangosNombradosPorPrefijo(ss, prefijo) {
  var rangosNombrados = ss.getNamedRanges();
  
  for (var i = 0; i < rangosNombrados.length; i++) {
    var nombre = rangosNombrados[i].getName();
    
    if (nombre.startsWith(prefijo)) {
      rangosNombrados[i].remove();
    }
  }
}




function isDate(value) {
  // Comprueba si el valor es una instancia válida de Date en JavaScript.
  return Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime());
}
