/**
 * Genera un archivo BC3 a partir de los datos proporcionados en una hoja de cálculo.
 *
 * Esta función toma los datos de una hoja de cálculo de Google Sheets y genera un archivo BC3
 * (presupuesto de obra) con el formato específico requerido. El archivo BC3 incluye información
 * sobre capítulos, subcapítulos, partidas, resúmenes y líneas de medición de una obra o proyecto.
 */
function generarArchivoBC3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Variables de versión

  var SoftwareName = "Morph Sheets to BC3";
  var Empresa = "Architecture Meets Engineering SL";
  var VersionFIEBDC = "FIEBDC-3/2020";
  var Rotulo_Identificacion = "00000";
  var Codification = "ANSI";

  // Variables iniciales

  var BC3Sheetname = "BC3";
  var VARSheetname = "Data";

  var originalFileName = ss.getName();
  var bc3Sheet = ss.getSheetByName(BC3Sheetname);
  var bc3Date = obtenerFechaHoraMadrid('ddMMyy');

  var finalFolder = DriveApp.getFolderById("1rnVSZZYHM7VHwxUVwISXBeres5kuiycQ");

  // Array para guardar las líneas del archivo BC3

  var bc3LinesInit = [];
  var bc3LinesRoot = [];
  var bc3LinesCaps = [];
  var bc3LinesSubs = [];
  var bc3LinesPart = [];
  var bc3LinesDesp = [];

  var bc3DesglCaps = [];
  var bc3DesglPart = [];

  // Leer los datos de la hoja

  var dataRange = bc3Sheet.getDataRange();
  var data = dataRange.getValues();

  // Mapear las columnas de la hoja Sheets

  var headerValues = bc3Sheet.getRange(3, 1, 1, bc3Sheet.getLastColumn()).getValues()[0];
  var columnMapping = headerValues.reduce((acc, val, idx) => {
    acc[val] = idx;
    return acc;
  }, {});

  var Capcode = columnMapping['Capcode'];
  var Subcode = columnMapping['Subcode'];
  var Codigo = columnMapping['Código'];
  var Nat = columnMapping['Nat'];
  var Ud = columnMapping['Ud'];
  var Resumen = columnMapping['Resumen'];
  var Comentario = columnMapping['Comentario'];
  var N = columnMapping['N'];
  var Longitud = columnMapping['Longitud'];
  var Anchura = columnMapping['Anchura'];
  var Altura = columnMapping['Altura'];
  var Cantidad = columnMapping['Cantidad'];
  var CanPres = columnMapping['CanPres'];

  // Generar líneas de inicio del documento

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VARSheetname);
  var datos = hoja.getDataRange().getValues();
  var variables = {};
  for (var i = 0; i < datos.length; i++) {
    var nombre = datos[i][0]; // Valor en la columna A
    var valor = datos[i][1]; // Valor en la columna B
    variables[nombre] = valor;
  }

  var baseLine = `~V|${Empresa}|${VersionFIEBDC}\\${bc3Date}|${SoftwareName}|${Rotulo_Identificacion}|${Codification}||||||`;
  var keyLine = `~K|2\\2\\2\\2\\2\\2\\2\\2\\EUR|${variables["CI"]}\\${variables["GG"]}\\${variables["BI"]}\\${variables["BAJA"]}\\${variables["IVA"]}|3\\2\\\\3\\2\\\\2\\2\\2\\2\\2\\2\\2\\2\\EUR||`;
  // example keyLine = `~K | DN \\ DD \\ DS \\ DR \\ DI \\ DP \\ DC \\ DM \\ DIVISA | CI \\ GG \\ BI \\ BAJA \\ IVA | DRC \\ DC \\ \\ DFS \\ DRS \\ \\ DUO \\ DI \\ DES \\ DN \\ DD \\ DS \\ DSP \\ DEC \\ DIVISA | [ n ] | `;
  bc3LinesInit.push(baseLine); bc3LinesInit.push(keyLine);

  var projectReference = data[0][3].toString();
  var projectCode = projectReference.split('-')[0].trim();
  var projectName = projectReference.split('-')[1].trim();

  var rootLine = `~C|${projectCode}##||${projectName}|1||0|`;
  bc3LinesRoot.push(rootLine);

  // Generar líneas de cuerpo del documento

  var line; var lineDesglose;
  
  data.forEach(row => {
    lineJump = "";
    switch (row[Nat]) {
      case 'Capítulo': // Línea de información del capítulo
        line = `~C|${row[Codigo]}#||${row[Resumen]}|0|${bc3Date}|0|`;
        Logger.log(`Capítulo: ${row[Resumen]}`);

        bc3LinesCaps.push(line); bc3DesglCaps.push(`${row[Codigo]}\\\\`);
        break;
      case 'Subcapítulo': // Línea de información del subcapítulo
        line = `~C|${row[Codigo]}#||${row[Resumen]}|0|${bc3Date}|0|`;
        lineDesglose = `~Y|${row[Capcode]}#|${row[Codigo]}#\\\\|`;

        bc3LinesSubs.push(line); bc3LinesSubs.push(lineDesglose);
        break;
        /**/
      case 'Partida': // Línea de información de la partida
        line = `~C|${row[Codigo]}|${row[Ud]}|${row[Resumen]}|0|${bc3Date}|PAR|`;

        bc3LinesPart.push(line); bc3DesglPart.push(`${row[Codigo]} \\ \\ `);
        break;
      case 'RES': // Resumen de la partida
        line = `~T|${row[Codigo]}|${row[Resumen]}|`;

        bc3LinesPart.push(line);
        break;
      case 'SLT': // Líneas de medición
        //line = `~N | ${row[Subcode]} \\ ${row[Codigo]} | | | LINO \\ ${row[Comentario]} \\ ${row[N]} \\ ${row[Longitud]} \\ ${row[Anchura]} \\ ${row[Altura]} \\ | LIN |`;
        line = `~N|${row[Subcode]}\\${row[Codigo]}|||LINO\\${row[Comentario]}\\${row[N].toString().trim()}\\${row[Longitud].toString().trim()}\\${row[Anchura].toString().trim()}\\${row[Altura].toString().trim()}\\|LIN|`;

        bc3LinesPart.push(line);
        break;
      case 'TZP': // Cuando termina un subcapítulo se añade la línea de desglose correspondiente a ese subcapítulo
        line = `~D|${row[Subcode].replace('.100000','')}|${bc3DesglPart.join('|')}|`;

        bc3LinesDesp.push(line);
        bc3DesglPart = [];
        break;
      default:
        break;
    }
  });

  rootLine = `~D|${projectCode}##|${bc3DesglCaps.join('|')}|`;
  bc3LinesRoot.push(rootLine);

  // Crear y guardar el archivo BC3
  var bc3Content = bc3LinesInit.concat(bc3LinesRoot, bc3LinesCaps, bc3LinesSubs, bc3LinesPart, bc3LinesDesp).join('\r\n');

  var fileName = originalFileName + "_BC3_" + obtenerFechaHoraMadrid() + ".bc3"; // Nombre del archivo XLSX 
  fileName = fileName.replace(/\s+/g, '_').replace(/:/g, '');
  // finalFolder.createFile(fileName, bc3Content, MimeType.PLAIN_TEXT);

  var blob = Utilities.newBlob('').setDataFromString(bc3Content, "ISO-8859-1");
  blob.setName(fileName);
  finalFolder.createFile(blob);
}

/**
 * Congela y formatea la hoja BC3 del Cuadro de Mediciones
 *
 */
function formatearHojaBC3() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var BC3Sheetname = "BC3";
  var originalSheet = spreadsheet.getSheetByName(BC3Sheetname);

  if (!originalSheet) {
    SpreadsheetApp.getUi().alert(`La hoja ${BC3Sheetname} no fue encontrada en el documento.`);
    return;
  }

  var duplicatedSheet = originalSheet.copyTo(spreadsheet);
  duplicatedSheet.clearContents();

  var originalData = originalSheet.getDataRange().getValues();
  var trimmedData = originalData.map(row => row.map(cell => typeof cell === 'string' ? cell.trim() : cell));
  duplicatedSheet.getRange(1, 1, trimmedData.length, trimmedData[0].length).setValues(trimmedData);

  var maxColumns = duplicatedSheet.getMaxColumns();
  var colAbackgrounds = duplicatedSheet.getRange("A:A").getBackgrounds();
  var rangesToFormat = [];

  colAbackgrounds.forEach((color, i) => {
    if (color[0] === "#c0c0c0") {
      var fila = i + 1;
      rangesToFormat.push({ range: duplicatedSheet.getRange(fila, 1, 1, maxColumns), row: fila, rowHeight: 1 });
    }
  });

  Logger.log(`rangesToFormat: ${rangesToFormat}`)

  rangesToFormat.forEach(r => {
    r.range.setBorder(true, true, true, true, true, true, "#c0c0c0", SpreadsheetApp.BorderStyle.SOLID);
    r.range.setBackground("#c0c0c0");
    duplicatedSheet.setRowHeight(r.row, r.rowHeight);
  });

  eliminarFilasSobrantes(duplicatedSheet);

  var headerValues = duplicatedSheet.getRange(3, 1, 1, duplicatedSheet.getLastColumn()).getValues()[0];
  var columnMapping = headerValues.reduce((acc, val, idx) => {
    acc[val] = idx + 1;
    return acc;
  }, {});

  var formatRanges = { azul: [], azulclaro: [], magenta: [], clear: [], bold: [] };

  var Capcode = columnMapping['Capcode'];
  var Subcode = columnMapping['Subcode'];
  var Codigo = columnMapping['Código'];
  var Nat = columnMapping['Nat'];
  var Ud = columnMapping['Ud'];
  var Resumen = columnMapping['Resumen'];
  var Comentario = columnMapping['Comentario'];
  var N = columnMapping['N'];
  var Longitud = columnMapping['Longitud'];
  var Anchura = columnMapping['Anchura'];
  var Altura = columnMapping['Altura'];
  var Cantidad = columnMapping['Cantidad'];
  var CanPres = columnMapping['CanPres'];
  var Pres = columnMapping['Pres'];
  var ImpPres = columnMapping['ImpPres'];

  duplicatedSheet.getRange("E:E").getValues().forEach((val, row) => {
    var fila = row + 1;
    switch (val[0]) {
      case 'Capítulo':
        formatRanges.azul.push({ fila: fila, startCol: 1, numCols: maxColumns });
        break;
      case 'Subcapítulo':
        formatRanges.azulclaro.push({ fila: fila, startCol: 1, numCols: maxColumns });
        break;
      case 'Partida':
        formatRanges.magenta.push({ fila: fila, startCol: CanPres, numCols: 3 });
        formatRanges.clear.push({ fila: fila, startCol: Cantidad, numCols: 1 });
        break;
      case 'RES':
        formatRanges.clear.push({ fila: fila, startCol: 1, numCols: 5 });
        break;
      case 'SLT':
        formatRanges.clear.push({ fila: fila, startCol: 1, numCols: 7 });
        formatRanges.magenta.push({ fila: fila, startCol: Cantidad, numCols: 1 });
        break;
      case 'TOT':
        formatRanges.bold.push({ fila: fila, startCol: Cantidad, numCols: 2 });
        formatRanges.bold.push({ fila: fila, startCol: ImpPres, numCols: 1 });
        formatRanges.magenta.push({ fila: fila, startCol: CanPres, numCols: 1 });
        formatRanges.magenta.push({ fila: fila, startCol: ImpPres, numCols: 1 });
        formatRanges.clear.push({ fila: fila, startCol: 1, numCols: 7 });
        break;
      case 'TZP':
        formatRanges.bold.push({ fila: fila, startCol: Pres, numCols: 2 });
        formatRanges.magenta.push({ fila: fila, startCol: Pres, numCols: 2 });
        formatRanges.clear.push({ fila: fila, startCol: 1, numCols: 7 });
        break;
      case 'TZZ':
        formatRanges.bold.push({ fila: fila, startCol: Pres, numCols: 2 });
        formatRanges.magenta.push({ fila: fila, startCol: Pres, numCols: 2 });
        formatRanges.clear.push({ fila: fila, startCol: 1, numCols: 7 });
        break;
      case 'VBK':
        formatRanges.clear.push({ fila: fila, startCol: 1, numCols: maxColumns });
        break;
    }
  });

  formatRanges.azul.forEach(r => duplicatedSheet.getRange(r.fila, r.startCol, 1, r.numCols).setBackground("#b4cbe0").setFontWeight("bold"));
  formatRanges.azulclaro.forEach(r => duplicatedSheet.getRange(r.fila, r.startCol, 1, r.numCols).setBackground("#c2d5e7").setFontWeight("bold"));
  formatRanges.magenta.forEach(r => duplicatedSheet.getRange(r.fila, r.startCol, 1, r.numCols).setFontColor("#ff00ff"));
  formatRanges.bold.forEach(r => duplicatedSheet.getRange(r.fila, r.startCol, 1, r.numCols).setFontWeight("bold"));
  formatRanges.clear.forEach(r => duplicatedSheet.getRange(r.fila, r.startCol, 1, r.numCols).clearContent());

  duplicatedSheet.deleteColumns(1, 3);
  duplicatedSheet.clearConditionalFormatRules();

  var ss_id = spreadsheet.getId();
  var file = DriveApp.getFileById(ss_id);
  var parentFolder = file.getParents();
  var parentFolderID = parentFolder.next().getId();
  var backupFolderSearch = DriveApp.getFolderById(parentFolderID);

  var backupFolder, backupFolderId;
  let searchFor = 'title contains "BC3"';
  backupFolder = backupFolderSearch.searchFolders(searchFor);

  if (!backupFolder.hasNext()) {
    var response = Browser.msgBox("Atención", "No se ha encontrado la carpeta para los archivos congelados. ¿Deseas crearla automáticamente?", Browser.Buttons.OK_CANCEL);
    if (response == "cancel") {
      throw new Error(`No se ha podido encontrar la carpeta de archivos congelados.`);
    }
    
    var folderName = `${ss_name.substring(0, 6)}-A-CS-${searchFor.substring(16, searchFor.length - 1)}`;
    var newFolder = backupFolderSearch.createFolder(folderName);

    backupFolderId = newFolder.getId();
    backupFolder = DriveApp.getFolderById(backupFolderId);
  } else {
    backupFolder = backupFolder.next();
    backupFolderId = backupFolder.getId();
  }

  var newSpreadsheet = generarXLSXdeDuplicatedSheet(backupFolder, duplicatedSheet, BC3Sheetname);

  var linkSheet = spreadsheet.getSheetByName("LINK");
  var data = linkSheet.getRange("A:A").getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][0] === "ÚLTIMO .BC3 CONGELADO") {
      linkSheet.getRange(i + 1, 2).setValue(newSpreadsheet.getUrl());
    } else if (data[i][0] === "CARPETA EXPORTACIÓN BC3") {
      linkSheet.getRange(i + 1, 2).setValue(backupFolder.getUrl());
    }
  }

  spreadsheet.deleteSheet(duplicatedSheet);
}

function generarXLSXdeDuplicatedSheet(backupFolder, duplicatedSheet, originalSheetName) {

  var spreadsheet = duplicatedSheet.getParent();
  var originalFileName = spreadsheet.getName();
  var duplicatedFileName = originalFileName + "_BC3_" + obtenerFechaHoraMadrid() + ".xlsx"; // Nombre del archivo XLSX 
  duplicatedFileName = duplicatedFileName.replace(/\s+/g, '_').replace(/:/g, '');

  // Crear una copia de la hoja duplicada en un nuevo libro
  var newSpreadsheet = SpreadsheetApp.create(duplicatedFileName)
  DriveApp.getFileById(newSpreadsheet.getId()).moveTo(backupFolder);

  var newSheet = newSpreadsheet.getActiveSheet();
  var finalSheet = duplicatedSheet.copyTo(newSpreadsheet);
  finalSheet.setName(originalSheetName)
  
  // Eliminar la hoja en blanco que se crea automáticamente en el nuevo libro
  newSpreadsheet.deleteSheet(newSheet);

/*
  // Abre una ventana del navegador con el enlace de descarga del archivo XLSX
  var newSpreadsheetId = newSpreadsheet.getId();
  var xlsxFileUrl = `https://docs.google.com/spreadsheets/d/${newSpreadsheetId}/export?format=xlsx`;
  var htmlOutput = HtmlService.createHtmlOutput(`<script>window.open("${xlsxFileUrl}");google.script.host.close();</script>`);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Descargar archivo XLSX");
*/

  return newSpreadsheet;
}

function obtenerFechaHoraMadrid(formato) {
  var formato = formato || 'yyyyMMdd HH:mm';
  var madridTimeZone = CalendarApp.getTimeZone();
  var fechaHoraMadrid = Utilities.formatDate(new Date(), madridTimeZone, formato);
  return fechaHoraMadrid;
}

function eliminarFilasSobrantes(sh) {
  sh = sh || SpreadsheetApp.getActiveSheet();
  let maxRows = sh.getMaxRows();
  let lastRow = sh.getLastRow();
  let checker2 = maxRows - (maxRows - lastRow);
  if (checker2 != maxRows) {
    sh.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}
