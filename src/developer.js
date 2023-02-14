
function fastInit() {
  Logger.log('Fast Init makes things faster.')
}

// CUSTOM FUNCTIONS FOR THE DEVELOPER SECTION

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
function getDevPermission(headerName) {
  const userPermission = getDatabaseColumn(headerName);
  const userMail = Session.getActiveUser().getEmail();
  let permission = userPermission !== '' && userPermission.indexOf(userMail) > -1 ? true : false; Logger.log(permission)
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

// FUNCTIONS IN DEVELOPMENT

/**
 * formulaLogger
 * Función para realizar macros rápidas para cualquier necesidad en un documento.
 */
function macroModificarCuadros(a) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();

  let maxColumns = sh.getMaxColumns();

  let firstRange = sh.getRange(4, 1, 1, maxColumns).getFormulas();
  let modifiedFormulas = [];
  let row = [];
  firstRange[0].forEach(function(formula) {
    row.push("={\"\";" + formula.slice(1) + "}");
  });
  modifiedFormulas.push(row);
  let pasteRange = sh.getRange(3, 1, 1, maxColumns).setFormulas(modifiedFormulas);
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
function historicoDeSuperficies(sheetRef) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sheetRef == 0 ? `Histórico CONSTRUIDAS` : `Histórico CONSTRUIDAS desglosado`;
  const sh = ss.getSheetByName(sheetName);
  const dateNow = Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy');

  let mainCell = sheetRef == 0 ? `D1` : `E1`;
  let firstCell = sheetRef == 0 ? `G1` : `H1`;
  let secondCell = sheetRef == 0 ? `H1` : `I1`;
  let thirdCell = sheetRef == 0 ? `E1` : `F1`;
  let groupRange = sheetRef == 0 ? `H:I` : `I:J`;
  
  let mainRange = sh.getRange(mainCell);
  let secondRange = sh.getRange(secondCell);
  let firstRange = sh.getRange(firstCell);
  let mainColumnIndex = mainRange.getColumn();
  let firstColumnIndex = firstRange.getColumn();

  let originalFormulaRange = sh.getRange(thirdCell);
  let originalFormula = originalFormulaRange.getFormulas();

  let freezeRange;
  let lastRow = sh.getLastRow();

  if (firstRange.isBlank()) {
    freezeRange = sh.getRange(1, mainColumnIndex, lastRow, 1);
    freezeRange.copyTo(sh.getRange(1, firstColumnIndex), {contentsOnly:true});
    sh.getRange(firstCell).setValue(dateNow);
  } else {

    // Comprobar si han cambiado los valores desde el último histórico

    var range1 = sh.getRange(2, mainColumnIndex, lastRow, 1).getValues();
    var range2 = sh.getRange(2, firstColumnIndex, lastRow, 1).getValues();
    var isEqual = true;
    for (var i = 0; i < lastRow; i++) {
      if (range1[i][0] != range2[i][0]) {
        isEqual = false;
        break;
      }
    }
    if (isEqual) {
      throw new Error('Los valores de superficies no han cambiado desde el último histórico.')
    }

    // Crear el histórico

    freezeRange = sh.getRange(1, mainColumnIndex, lastRow, 3);

    sh.insertColumns(firstColumnIndex, 3);
    freezeRange.copyTo(sh.getRange(1, firstColumnIndex), {contentsOnly:true});
    firstRange.setValue(dateNow);

    // Modificación de estilo / formato de texto

    let columns = [{column: firstColumnIndex, width: 100},{column: firstColumnIndex + 1, width: 100},{column: firstColumnIndex + 2, width: 130}];
    columns.forEach(function(column) {
      sh.setColumnWidth(column.column, column.width);
    });

    let freezeRangeFormat = sh.getRange(1, firstColumnIndex - 2, 1, 2);
    freezeRangeFormat.copyTo(sh.getRange(1, firstColumnIndex + 1), {formatOnly:true});
    secondRange.setBorder(false, false, false, false, false, false);
    let frozenRange2 = sh.getRange(2, firstColumnIndex + 1, lastRow - 1, 2);
    frozenRange2.setBackgroundColor(null).setFontColor('black');
    sh.getRange(1, firstColumnIndex, lastRow, 1).setBorder(null, true, null, null, null, null, firstRange.getBackground(), SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    sh.getRange(groupRange).shiftRowGroupDepth(1); // Agrupa las nuevas columnas

    // Arreglo de fórmulas, modificar newFormulas si cambian en la plantilla

    sh.getRange(2, mainColumnIndex + 2, lastRow - 1, 1).clearContent(); // Limpia columna original de justificación
    sh.getRange(2, firstColumnIndex + 1, lastRow - 1, 1).clearContent(); // Limpia nueva columna diferencia de construidas

    let newFormula = sheetRef == 0 ? `={"Diferencia con "&IF(TO_TEXT(J1)<>"";TO_TEXT(J1);"última fecha");ARRAYFORMULA(IF(B2:B<>"";IF(TO_TEXT(J2:J)<>"";G2:G-J2:J;0);))}` : `={"Diferencia con "&IF(TO_TEXT(K1)<>"";TO_TEXT(K1);"última fecha");ARRAYFORMULA(IF(B2:B<>"";IF(TO_TEXT(K2:K)<>"";H2:H-K2:K;0);))}`;
    originalFormulaRange.setFormula(originalFormula); 
    secondRange.setFormula(newFormula); 
  }
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

// AUTOFOLDERTREE AND PROJECT FOLDER INIT

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

// MORPH CHATBOT DEVELOPMENT

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

// DEPRECATED FUNCTIONS

/**
 * colorMe
 * Fondo de celdas con el color Morph
 */
function colorMe() {
  let ss = SpreadsheetApp.getActive();
  let selection = ss.getSelection();
  let currentCell = selection.getActiveRange();
  currentCell.setBackgroundColor('#f1cb50');
}

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
 * adjustRowsHeight
 * Ajusta la altura de las filas seleccionadas.
 */
function adjustRowsHeight(rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();

  let formData = [rowData.heightSelectedNumber];
  let [heightSelectedNumber] = formData;

  let a = sh.getSelection().getActiveRange().getValues();
  let ab = sh.getSelection().getActiveRange().getA1Notation();
  let abs = ab.split(':'); let abst = getSplitA1Notation(abs[0]);
  sh.setRowHeights(abst[1], a.length, heightSelectedNumber);
}

/**
 * autoResizeAllRows, autoResizeAllCols
 * Automatically adjust the size of rows and columns
 */
function autoResizeAllRows() {
  const sh = ss().getActiveSheet();
  const maxRows = sh.getLastRow();
  sh.autoResizeRows(1, maxRows)
}

function autoResizeAllCols() {
  const sh = ss().getActiveSheet();
  const numCols = sh.getLastColumn();

  for (let j = 1; j < numCols + 1; j++) {
    sh.autoResizeColumn(j);
    let colWidth = sh.getColumnWidth(j)
    sh.setColumnWidth(j, colWidth + 20)
  }
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
