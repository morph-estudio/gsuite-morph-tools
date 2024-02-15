/**
 * Elimina las columnas vacías en una hoja de cálculo hasta la última columna que contiene datos.
 *
 * @param {Sheet} sh - La hoja de cálculo en la que se eliminarán las columnas vacías (opcional, se utiliza la hoja activa por defecto).
 */
function sendToBigqueryDatabase() {

  var histSheetname = 'his_export';
  var projectId = 'alsan-scripts';
  var datasetId = 'medautocomparativo';
  var tableId = 'historico';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(histSheetname);
  var rango = hoja.getDataRange();
  var values = rango.getValues();

  var rowsCSV = values.join("\n");
  var data = Utilities.newBlob(rowsCSV, 'application/octet-stream');

  function convertValuesToRows(data) {
    var rows = [];
    var headers = values[0];

    Logger.log(`headers: ${headers}`)

    for (var i = 1, numColumns = data.length; i < numColumns; i++) {
      var row = BigQuery.newTableDataInsertAllRequestRows();
      row.json = data[i].reduce(function(obj, value, index) {
        obj[headers[index]] = value;
        return obj;
      }, {});
      rows.push(row);
    };
    Logger.log(`rows: ${rows}`)
    return rows;
  }

  function bigqueryInsertData(data, tableId) {
    var insertAllRequest = BigQuery.newTableDataInsertAllRequest();
    insertAllRequest.rows = convertValuesToRows(data);     
    var response = BigQuery.Tabledata.insertAll(insertAllRequest, projectId, datasetId, tableId);
    if (response.insertErrors) {
      Logger.log(response.insertErrors);
    }
  }

  try {
    bigqueryInsertData(Utilities.parseCsv(data.getDataAsString()), tableId);
  } catch (error) {
    Logger.log(`Error al enviar datos a BigQuery: ${error.message}`);
    throw new Error('Ha habido un error enviando los datos a la base de datos.');
  }
}

function fechaformat() {
  var fecha = `12/01/2024`
  var fechaNueva = Utilities.formatDate(fecha, 'GMT', 'yyyy-MM-dd');
  Logger.log(fechaNueva)
}

function sendToLocalCSV(dataSheetName, fileName) {
  
  // Get the main file folder
  var currentFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next();
  
  // Verify if CSV file exists in current folder
  var files = currentFolder.getFilesByName(fileName);
  var file;
  if (files.hasNext()) {
    file = files.next();
  } else {
    file = currentFolder.createFile(fileName, '');
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(dataSheetName);
  var data = sheet.getRange('A2:' + sheet.getLastColumn() + sheet.getLastRow()).getValues();
  
  // CSV Compose 
  var csvContent = '';
  for (var row = 0; row < data.length; row++) {
    csvContent += data[row].join(',') + '\n';
  }
  
  if (file.getSize() > 0) {
    var existingData = file.getBlob().getDataAsString();
    csvContent = existingData + csvContent;
  }
  
  file.setContent(csvContent);
}
