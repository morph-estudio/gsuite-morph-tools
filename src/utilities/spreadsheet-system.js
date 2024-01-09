  /**
 * Gsuite Morph Tools - Morph Spreadsheet System 1.8.3
 * Developed by alsanchezromero
 *
 * Copyright (c) 2023 Morph Estudio
 * 
 * Morph Spreadsheet System es un sistema de emulación de fórmulas offline para Google Sheets que evita la recalculación automática.
 */
function bddcompuestos(mssUpdateColumnReferenceText, mssUpdateAllCheck) {

  var startTime = new Date().getTime(); var elapsedTime;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var numRows = sheet.getLastRow();

  //var cellRefs = ["AP;Q;AR;AT;BF|concat;/"];
  //var cellRefs = ["BJ-3-12|concat;+"];
  //var cellRefs = ["BJ-3-12|filter;'BDD Componentes'!$AR$;'BDD Componentes'!$AV$~concat;/"];
  //var cellRefs = ["AP;Q;AR;AT;BF|concat;/", "BJ-3-12|concat;+", "BJ-3-12|filter;'BDD Componentes'!$AR$;'BDD Componentes'!$AV$~concat;/"];
  //var cellRefs = sheet.getRange(1, ColumnBDD('ET'), 1, sheet.getLastColumn() - ColumnBDD('ET')).getValues()[0];

  // Variables que definen a partir de qué fila se ejecutarán las acciones

  let mssUpdateColumnLastDataRow = getLastDataRow(sheet, mssUpdateColumnReferenceText);
  var firstRowToApply = mssUpdateColumnLastDataRow + 1;
  var numberOfRowsToApply = numRows - mssUpdateColumnLastDataRow; Logger.log(`firstRowToApply: ${firstRowToApply}, numberOfRowsToApply: ${numberOfRowsToApply}`);

  var cellRefs = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0]; // Matriz nodriza con las formulas MSS
  var mainColumnSmallestIndex = smallestNumberInMatrix(cellRefs); // Index de la primera columna con formulas MSS

  var arrayCell = cellRefs.filter(n => n); // Matriz nodriza limpia de valores vacíos

  // Arrays de trabajo

  var arrayCellRefs = []; // Array base de rangos
  var arrayCellParams = []; // Array base de parámetros

  for (var i = 0; i < arrayCell.length; i++) {
    var cellParts = arrayCell[i].split("|");
    arrayCellRefs.push(cellParts[0]);
    arrayCellParams.push(cellParts[1]);
  }

  var arrayDelimiter = arrayCellRefs.map(function(elem) {
    if (elem.match(/[;-]/)) {
      return elem.match(/[;-]/)[0];
    } else {
      return null;
    }
  });

  var methods, cellRefs, baseDelimiter;
  var cellListData, cellList;

  var formulalist = []; // Objeto CLAVE: almacena los resultados de cada columna
  var filterCache = {}; // Objeto CLAVE: almacenar los datos de la caché local

  // SUPERBUCLE: aquí comienza el bucle por cada una de las formulas MSS base

  for (var i = 0; i < arrayCellRefs.length; i++) {

    cellRefs = arrayCellRefs[i]; Logger.log(`cellRefs${i}: ${JSON.stringify(cellRefs)}`);
    baseDelimiter = arrayDelimiter[i];

    /*
    if (cellRefs == 'cache') {
    }
    */

    switch (arrayDelimiter[i]) {
      case ";":

        Logger.log(`Lista de rangos aislados (cortados por delimitador ;)`);

        cellListData = cellRefs.split(";");
        cellList = cellListData.map(function(value) {
          return ColumnBDD(value);
        });

        break;
      case "-":

        Logger.log(`Lista de rangos recurrentes (cortados por delimitador -)`);

        cellListData = cellRefs.split("-");
        cellList = getColumnList(ColumnBDD(cellListData[0]), cellListData[1], cellListData[2]);

        break;
      default:
        cellListData = [ cellRefs ];
        cellList = cellListData.map(function(value) {
          return ColumnBDD(value);
        });

        break;
    }

    var cellListPrimary = cellList; Logger.log(`cellListPrimary${i}: ${JSON.stringify(cellListPrimary)}`); // Guardo la estructura primigenia para poder utilizarla más adelante en caso de ser necesario

    var methods = arrayCellParams[i]; // Logger.log(`methods${i}: ${methods}`);
    var methodList = methods.split("~"); Logger.log(`methodList${i}: ${methodList} y el methodCount está a ${methodCount}`); // Array con los métodos que tiene la formula
    var methodCount = 0;

    // Mapear en cellList los valores de cada columna
    
    cellList = mapColumnValues(cellList, sheet, firstRowToApply, numberOfRowsToApply);
    Logger.log(`cellList${i}: ${JSON.stringify(cellList)}`);

    // Itera a través de los métodos de la formula

    var param1, param2, param3, param4;

    for (var p = 0; p < methodList.length; p++) {

      var paramSplit = methodList[p].split(";");
      var methodType = paramSplit[0];
      
      Logger.log(`El método ${p} para la formula ${i} es ${methodType} y sus parámetros son ${JSON.stringify(paramSplit)} `);

      elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before ${methodType}-${i}-${p}: ${elapsedTime} seconds.`);

      // Este Switch maneja cada tipo de método posible

      switch (methodType) {
        case "cache":

          Logger.log(`Se ha entrado al switch-case CACHE`); methodCount += 1; 

          cellListPrimary


          break;
        case "sust":

          Logger.log(`Se ha entrado al switch-case SUST`); methodCount += 1; 

          param1 = paramSplit[1]; // Array de índices de rangos a aplicar ([0] = todos, [1, 3] = se aplica al primer y tercer rango de la lista)
          param2 = paramSplit[2]; // Tipo de mapeo a realizar (map, add, sustract, etc.)
          param3 = paramSplit[3]; // Texto a añadir

          param1 = param1.slice(1, -1).split(',').map(Number);

          // Logger.log(`cellList-${i}-${p}: ${JSON.stringify(cellList)}`);

          if(param1[0] > 0) { // Aplicar el método según el índice
            cellListTemp = cellList.filter((subarray, index) => param1.includes(index + 1));
            var mappedSubarrays = sustCellList(param3, cellListTemp); Logger.log(`mappedSubarrays${i}-${p}: ${JSON.stringify(mappedSubarrays)}`);
            var recomposedCellList = recomposeList(cellList, mappedSubarrays, param1);
            cellList = recomposedCellList;
          } else {
            cellList = sustCellList(param3, cellList);
          }

          Logger.log(`finalCellList-${i}-${p}: ${JSON.stringify(cellList)}`);

          if(methodCount == methodList.length) { formula = cellList };

          break;
        case "filter":

          param1 = paramSplit[1]; // Columna de búsqueda
          param2 = paramSplit[2]; // Columna de resultado
          param3 = paramSplit[3]; // Columnas a filtrar [0] o  [x,x,x...] //Logger.log(`Filter Parametro 2: ${param2}`);

          Logger.log(`Se ha entrado al switch-case FILTER: Search Column Param${i}-${p}: ${param1}, Result Column Param${i}-${p}: ${param2}`); methodCount += 1; 

          // Desglose de los parámetros

          param3 = param3.slice(1, -1).split(',').map(Number);

          let range1 = param1.replace(/['$]/g,"").split("!");
          let range1_sheetname = range1[0]; let range1_colLetter = range1[1];
          let range1_colarray = [ColumnBDD(range1_colLetter)];

          let range2 = param2.replace(/['$]/g,"").split("!");
          let range2_sheetname = range2[0]; let range2_colLetter = range2[1];
          let range2_colarray = [ColumnBDD(range2_colLetter)];

          Logger.log(`range1: ${range1_sheetname}+${range1_colarray}, range2: ${range2_sheetname}+${range2_colarray}`);

          if (`${range1_sheetname}+${range1_colarray}` == `${range2_sheetname}+${range2_colarray}`) { break; }

          // Las matrices principales se transforman en matriz unidimensional

          elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before mapping colToSearch and colToReturn-${i}-${p}: ${elapsedTime} seconds.`);

          var colToSearch, colToReturn;
          var cacheReference_SC = `${range1_sheetname},${range1_colLetter.toString()}`;
          var cacheReference_RC = `${range2_sheetname},${range2_colLetter.toString()}`;

          if (filterCache.hasOwnProperty(cacheReference_SC)) {
            colToSearch = filterCache[cacheReference_SC].colToSearch;
            Logger.log(`CACHE: Se ha activado el caché para colToSearch.`);
          } else {
            // Las variables no están en caché, se calculan y se almacenan en caché
            let range1_sheet = ss.getSheetByName(range1_sheetname);
            let range1_numRows = range1_sheet.getLastRow();
            colToSearch = mapColumnValues(range1_colarray, range1_sheet, 1, range1_numRows)[0];
            colToSearch = flattenArray(colToSearch);
            filterCache[cacheReference_SC] = {
              colToSearch: colToSearch,
            };
            Logger.log(`CACHE: Se ha almacenado en caché: colToSearch`);
          }

          if (filterCache.hasOwnProperty(cacheReference_RC)) {
            colToReturn = filterCache[cacheReference_RC].colToReturn;
            Logger.log(`CACHE: Se ha activado el caché para colToReturn.`);
          } else {
            let range1_sheet = ss.getSheetByName(range1_sheetname);
            let range1_numRows = range1_sheet.getLastRow();
            colToReturn = mapColumnValues(range2_colarray, range1_sheet, 1, range1_numRows)[0];
            elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time after DIRECT MAPPING colToSearch and colToReturn-${i}-${p}: ${elapsedTime} seconds.`);
            colToReturn = flattenArray(colToReturn);
            elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time after FLATTENNING colToSearch and colToReturn-${i}-${p}: ${elapsedTime} seconds.`);
            filterCache[cacheReference_RC] = {
              colToReturn: colToReturn
            };
            Logger.log(`CACHE: Se ha almacenado en caché: colToReturn.`);
          }

          elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time after mapping colToSearch and colToReturn-${i}-${p}: ${elapsedTime} seconds.`);

          // Logger.log(`colToReturnBugFinding: ${JSON.stringify(colToReturn)}`);

          if(param3[0] > 0) {

            var joinedReference = cellListPrimary.filter((subarray, index) => param3.includes(index + 1)); Logger.log(`newCellList: ${JSON.stringify(cellListTemp)}`);
            cacheReference = `${joinedReference.toString()};${param1}`; Logger.log(`cacheReference: ${JSON.stringify(cacheReference)}`);

            cellListTemp = cellList.filter((subarray, index) => param3.includes(index + 1));

            elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before mappedSubarrays-${i}-${p}: ${elapsedTime} seconds.`);

            var mappedSubarrays = filterColumns(cellListTemp, colToSearch, colToReturn, cacheReference, filterCache, startTime, methodType, i, p);

            elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before RecomposeCellList-${i}-${p}: ${elapsedTime} seconds.`);

            var recomposedCellList = recomposeList(cellList, mappedSubarrays, param3);

            cellList = recomposedCellList;
          
          } else {
            
            cacheReference = `${cellListPrimary.toString()};${param1}`; Logger.log(`cacheReference: ${JSON.stringify(cacheReference)}`);

            cellList = filterColumns(cellList, colToSearch, colToReturn, cacheReference, filterCache, startTime, methodType, i, p);
          }

          // Logger.log(`FilterCacheAfter-${i}-${p}: ${JSON.stringify(filterCache)}`);

          break;
        case "oper":

          param1 = paramSplit[1];
          Logger.log(`Se ha entrado al switch-case OPER. La operación es: ${param1}`); methodCount += 1; 

          cellList = applyOperationToNumbers(cellList, paramSplit[1]);

          break;

        case "concat":

          param1 = paramSplit[1];
          Logger.log(`Se ha entrado al switch-case CONCAT. El delimitador de concatenación es: ${param1}`); methodCount += 1;

          cellList = concatColumns(numRows, cellList, param1);

          break;

        default:
          break;
      }

      Logger.log(`cellListAfterMethod-${i}-${p}: ${JSON.stringify(cellList)}`);

      if(methodCount == methodList.length) { formula = cellList }; // Cuando ya no haya más métodos, añade el contenido a formula

    }

    formulalist.push(formula); // Una vez se ha iterado por todos los métodos, se envía la formula al array clave

  }
  
  Logger.log(`formulalist: ${JSON.stringify(formulalist)}`);
  
  // Transponer la matriz de formulas para insertarlas en la hoja de Sheets

  var transposedList = formulalist[0].map(function(_, i) {
    return formulalist.map(function(row) {
      return row[i];
    });
  });

  Logger.log(`El resultado final se insertará a partir de la columna: ${JSON.stringify(mainColumnSmallestIndex + 1)}`);
  
  sheet.getRange(firstRowToApply, mainColumnSmallestIndex + 1, transposedList.length, transposedList[0].length).setValues(transposedList);

}





// OPERATION FUNCTIONS

/**
 * sustCellList
 * 
 */
function sustCellList(string, cellList) {
  const cellListfinal = cellList.map(subarray1 =>
    subarray1.map(subarray2 =>
      subarray2.map(value =>
        value === "" ? "" : value + string
      )
    )
  );
  return cellListfinal;
}

/**
 * filterColumns
 * 
 */

function filterColumns(baseColumns, columnToSearch, columnToReturn, cacheReference, filterCache, startTime, methodType, i, p) {

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time ENTERING ${methodType}-${i}-${p} FUNCTION: ${elapsedTime} seconds.`);

  if (filterCache.hasOwnProperty(cacheReference)) {
    // Obtener el array de índices de la cache
    var indexArray = filterCache[cacheReference].indexarray; Logger.log(`IndexArrayCachedReturned-${i}-${p}: ${JSON.stringify(indexArray)}`);
    Logger.log(`CACHE: Se ha activado el caché para este filtrado.`);
    // Utilizar el array de índices para obtener directamente los resultados de la nueva columna
    var updatedArray = baseColumns.map((row, i) => row.map((val, j) => {
      if (indexArray[i][j][0] !== -1) {
        return [columnToReturn[indexArray[i][j][0]]];
      } else {
        return [""];
      }
    }));
    Logger.log(`CACHE: Resultado filtrado: ${JSON.stringify(updatedArray)}`);
    elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time EXIT ${methodType}-${i}-${p} FUNCTION CACHED: ${elapsedTime} seconds.`);
    return updatedArray;
  }

  Logger.log(`CACHE: Esta entrada no está en caché, se creará un nuevo registro.`);

  var indexArray = baseColumns.map(row => row.map(val => {
    if (val[0].trim().length > 0) {
      var index = columnToSearch.indexOf(val[0]);
      if (index >= 0) {
        return [index];
      } else {
        return [-1];
      }
    } else {
      return [-1];
    }
  }));

  Logger.log(`IndexArrayNoCached-${i}-${p}: ${JSON.stringify(indexArray)}`);

  // Almacenar el array de índices en la cache
  filterCache[cacheReference] = {
    indexarray: indexArray
  };

  // Utilizar el array de índices para obtener directamente los resultados de la nueva columna
  var updatedArray = baseColumns.map((row, i) => row.map((val, j) => {
    if (indexArray[i][j][0] >= 0) {
      return [columnToReturn[indexArray[i][j][0]]];
    } else {
      return [""];
    }
  }));

  Logger.log(`Resultado Filtrado: ${JSON.stringify(updatedArray)}`);
  // Logger.log(`Resultado Filter Cache: ${JSON.stringify(filterCache)}`);

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time EXIT ${methodType}-${i}-${p} FUNCTION NOT CACHED: ${elapsedTime} seconds.`);
  
  return updatedArray;
}






function filterColumnsv1(baseColumns, columnToSearch, columnToReturn, cacheReference, filterCache) {

  if (filterCache.hasOwnProperty(cacheReference)) {
    // Obtener el array de índices de la cache
    var indexArray = filterCache[cacheReference].indexarray;
    Logger.log(`CACHE: Se ha activado el caché para este filtrado.`);
    // Utilizar el array de índices para obtener directamente los resultados de la nueva columna
    var updatedArray = [];
    for (let i = 0; i < baseColumns.length; i++) {
      updatedArray.push([]);
      for (let j = 0; j < baseColumns[i].length; j++) {
        if (indexArray[i][j][0] >= 0) {
          updatedArray[i][j] = [columnToReturn[indexArray[i][j][0]]];
        } else {
          updatedArray[i][j] = [""];
        }
      }
    }
    Logger.log(`Resultado filtrado (desde cache): ${JSON.stringify(updatedArray)}`);
    return updatedArray;
  }

  Logger.log(`CACHE: Esta entrada no está en caché, se creará un nuevo registro.`);

  var updatedArray = [];
  var loggerArray = [];
  var indexArray = [];

  for (let i = 0; i < baseColumns.length; i++) {
    updatedArray.push([]);
    indexArray.push([]);
    for (let j = 0; j < baseColumns[i].length; j++) {
      var val = baseColumns[i][j];
      loggerArray.push(val);

      if (val[0].trim().length > 0) {
        var index = columnToSearch.indexOf(val[0]);

        // Si se encuentra una coincidencia, agregar el valor correspondiente en de la columna columnToReturn
        if (index >= 0) {
          updatedArray[i][j] = [columnToReturn[index]];
          indexArray[i][j] = [index];
        } else {
          updatedArray[i][j] = [""];
          indexArray[i][j] = [-1];
        }
      } else {
        updatedArray[i][j] = [""];
        indexArray[i][j] = [-1];
      }
    }
  }

  // Almacenar el array de índices en la cache
  filterCache[cacheReference] = {
    indexarray: indexArray
  };

  // Utilizar el array de índices para obtener directamente los resultados de la nueva columna
  for (let i = 0; i < baseColumns.length; i++) {
    for (let j = 0; j < baseColumns[i].length; j++) {
      if (indexArray[i][j][0] >= 0) {
        updatedArray[i][j] = [columnToReturn[indexArray[i][j][0]]];
      } else {
        updatedArray[i][j] = [""];
      }
    }
  }

  Logger.log(`Resultado Filtrado: ${JSON.stringify(updatedArray)}`);
  Logger.log(`Resultado Filter Cache: ${JSON.stringify(filterCache)}`);
  
  return updatedArray;
}













/**
 * applyOperationToNumbers (oper)
 * 
 */


function applyOperationToNumbers(cellList, operationString) {

  var operator = operationString[0];
  var numbers = parseFloat(operationString.substr(1));

  var operationFunction;
  switch (operator) {
    case "+":
      operationFunction = function(num) { return num + numbers; };
      break;
    case "-":
      operationFunction = function(num) { return num - numbers; };
      break;
    case "*":
      operationFunction = function(num) { return num * numbers; };
      break;
    case "/":
      operationFunction = function(num) { return num / numbers; };
      break;
    default:
      throw new Error("Operación no válida: " + operator);
  }

  // Utilizar map() para recorrer las filas de la lista
  var updatedCellList = cellList.map(function(row) {
    // Utilizar map() para recorrer las celdas de la fila
    return row.map(function(cell) {
      var num = parseFloat(cell[0]);
      // Verificar si el valor de la celda es un número y no es NaN
      if (!isNaN(num)) {
        // Aplicar la operación al número y actualizar el valor en la celda
        cell[0] = operationFunction(num).toString();
      }
      return cell;
    });
  });

  return updatedCellList;
}


function applyOperationToNumbersv1(cellList, operationString) {

  var operationFunction;

  var operator = operationString[0]; Logger.log(`operator: ${operator}`); 
  var numbers = parseFloat(operationString.substr(1));  Logger.log(`numbers: ${numbers}`); 
  
  switch (operator) {
    case "+":
      operationFunction = function(num) { return num + numbers; };
      break;
    case "-":
      operationFunction = function(num) { return num - numbers; };
      break;
    case "*":
      operationFunction = function(num) { return num * numbers; };
      break;
    case "/":
      operationFunction = function(num) { return num / numbers; };
      break;
    default:
      throw new Error("Operación no válida: " + operator);
  }

  // Recorremos las filas de la lista
  for (var i = 0; i < cellList.length; i++) {
    // Recorremos las celdas de la fila
    for (var j = 0; j < cellList[i].length; j++) {
      // Verificamos si el valor de la celda es un número
      if (!isNaN(cellList[i][j][0])) {
        // Verificamos si el valor de la celda es una cadena vacía
        if (cellList[i][j][0] !== "") {
          // Aplicamos la operación al número y actualizamos el valor en el array
          cellList[i][j][0] = operationFunction(cellList[i][j][0]);
        }
      }
    }
  }

  return cellList;
}

/**
 * concatColumns
 * Concatenación de valores de un array de columnas
 */
function concatColumns(numberOfRowsToApply, cellList, separator='+') {

  var values = [];
  
  // Obtener los valores de cada columna y concatenarlos por fila
  for (var i = 0; i < cellList.length; i++) {
    var columnValues = cellList[i]
    for (var j = 0; j < numberOfRowsToApply; j++) {
      if (typeof columnValues[j] === 'object' && columnValues[j][0] !== '') {
        if (!values[j]) {
          values[j] = columnValues[j][0];
        } else {
          values[j] += separator + columnValues[j][0];
        }
      }
    }
  }

  var output = [];
  for (var i = 0; i < values.length; i++) {
    output.push([values[i]]);
  }

  return output;
}







// HELPER FUNCTIONS

/**
 * ColumnBDD
 * Devuelve el index de una columna de sheets introduciendo la letra de columna
 */
function ColumnBDD(baseCell_mat) {
  var index = 0;
  for (var i = 0; i < baseCell_mat.length; i++) {
    var charCode = baseCell_mat.charCodeAt(i) - 64;
    index += charCode * Math.pow(26, baseCell_mat.length - i - 1);
  }
  return parseInt(index);
}

/**
 * getColumnList
 * List de números a partir de un inicial, final y offset
 */
function getColumnList(startnumber, offset, count) {
  var numberList = [];
  for (var i = 0; i < count; i++) {
    numberList.push(startnumber + i * offset);
  }
  return numberList;
}
/**
 * mapColumnValues
 * Coge un array de index de columnas y devuelve un array de valores de columnas
 */
function mapColumnValues(cellList, sheet, firstRowToApply, numberOfRowsToApply) {
  var columnValuesList = cellList.map(function(cell) {
    return sheet.getRange(firstRowToApply, cell, numberOfRowsToApply, 1).getValues();
  });

  Logger.log(`MappedColumnValues: ${JSON.stringify(columnValuesList)}`);
  
  return columnValuesList;
}

// Esta no funciona bien
function mapColumnValues2(cellList, sheet) {
  var data = sheet.getDataRange().getValues();
  var result = [];

  cellList.forEach(function(cell) {
    var column = data.map(function(row) {
      return [row[cell - 1]]; // ajustar índice de columna a base 0
    });
    result.push([column]);
  });

  Logger.log(`MappedColumnValues: ${JSON.stringify(result)}`);

  return result[0];
}





/**
 * recomposeList
 * 
 */
function recomposeList(cellList, mappedSubarrays, param3) {
  var recomposedCellList = [];
  let i = 0; let j = 0;

  while (i < cellList.length) {
    if (param3.includes(i + 1)) {
      recomposedCellList.push(mappedSubarrays[j]);
      j++;
    } else {
      recomposedCellList.push(cellList[i]);
      
    }
    i++;
    
  }
  return recomposedCellList;
}

/**
 * flattenArray
 * 
 */
function flattenArray(matriz) {
  return matriz.flat();
}


function smallestNumberInMatrix(array) {

  var nonEmptyColumns = [];
  for (var i = 0; i < array.length; i++) {
    if (array[i] !== "") {
      nonEmptyColumns.push(i);
    }
  }

  return Math.min.apply(null, nonEmptyColumns);
}
