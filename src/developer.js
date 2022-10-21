/**
 * adaptarCuadroAntiguo
 * Script para adaptar cuadros de superficies antiguos a la nueva estructura automática.
 */
function crearPuntoHistorico() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  let dateNow = Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy');

  let firstCell = 'G1';
  let columnIndex;
  columnIndex = sh.getRange(firstCell).getColumn();
  let lastColumn = sh.getLastColumn(); Logger.log(lastColumn)
  let checkRange = sh.getRange(1, lastColumn).getValue();

  if (checkRange == "N/A") {
    //sh.insertColumnAfter(6); sh.setColumnWidth(7, 100);
    
    let copyRange = sh.getRange(1, 4, sh.getLastRow(), 1);
    copyRange.copyTo(sh.getRange(1, columnIndex), {contentsOnly:true}); /**/
    sh.getRange(firstCell).setValue(dateNow);
    return;
  }
    

  let copyRange = sh.getRange(1, 4, sh.getLastRow(), 3);
  sh.insertColumns(columnIndex, 3);
  sh.setColumnWidth(columnIndex, 100); sh.setColumnWidth(columnIndex + 1, 110); sh.setColumnWidth(columnIndex + 2, 150);
  sh.getRange('H:I').shiftRowGroupDepth(1);

  copyRange.copyTo(sh.getRange(1, columnIndex), {contentsOnly:true});
  let dateHeader = sh.getRange(firstCell);
  dateHeader.setValue(dateNow);

  let copyRange2 = sh.getRange(1, columnIndex, sh.getLastRow(), 1);
  let sheetID = sh.getSheetId();
  copyRange2.copyFormatToRange(sheetID, columnIndex, columnIndex, 1, sh.getLastRow())
  //copyRange2.copyTo(sh.getRange(1, columnIndex), {formatOnly:true});
  let copyRange3 = sh.getRange(1, columnIndex - 2, sh.getLastRow(), 2);
  copyRange3.copyTo(sh.getRange(1, columnIndex + 1), {formatOnly:true});

  let boldColor = dateHeader.getBackground();
  copyRange2.setBorder(null, true, null, null, null, null, boldColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  copyRange2.setBorder(null, null, null, true, null, null, boldColor, SpreadsheetApp.BorderStyle.SOLID);

  

  sh.getRange(1, columnIndex + 1).setFormula('="diferencia con "&TO_TEXT($J$1)');

  

  sh.getRange(2, columnIndex + 1).setValue("").setFormula('=ARRAYFORMULA(IF(B2:B<>"";G2:G-J2:J;))');
  sh.getRange(3, columnIndex + 1, sh.getLastRow() - 2, 1).clearContent();
/**/

}


/**
 * adaptarCuadroAntiguo
 * Script para adaptar cuadros de superficies antiguos a la nueva estructura automática.
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
    templateFormat(sh_link); templateText(sh_link); deleteEmptyRows(); removeEmptyColumns();
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

  PropertiesService.getDocumentProperties().setProperties({
    'adaptedSpreadsheet': true,
  });

  sidebarIndex()

}

/**
 * getSheetnames
 * Devuelve una lista con el nombre de las hojas del documento Sheets.
 */
function getSheetnames(ss) { 
  var out = [];
  var sheets = ss.getSheets();
  for (var i = 0 ; i < sheets.length ; i++) out.push( sheets[i].getName() )
  return out;
}

/**
 * getDatabaseColumn
 * Devuelve los valores de una columna en documento Sheets externo a través de su título.
 */
function getDatabaseColumn(headerName) {
  const parsedDB = JSON.parse(UrlFetchApp.fetch('https://opensheet.elk.sh/1lcymggGAbACfKuG0ceMDWIIB9zWuxgVtSR9qpgNq4Ng/Permissions').getContentText());
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

/*

function jsonColumnArray_Deprecated(database, keyName) {
  let permission = []
  for (let i = 0; i < database.length; i++) {
    permission.push(database[i][keyName])
  }
  permission = permission.filter(function (el) {
    return el != null;
  });

  return permission;
}

function getPermission_Deprecated(database, keyName) {
  let databaseParsed = JSON.parse(UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/1lcymggGAbACfKuG0ceMDWIIB9zWuxgVtSR9qpgNq4Ng/gviz/tq?tqx=out:json&gid=0')
    .getContentText().match(/(?<=.*\().*(?=\);)/s)[0]);
  let columnNeeded = databaseParsed.table.cols.findIndex(obj => obj.label === headerName);
  let tableLength = Object.keys(databaseParsed.table.rows).length; Logger.log(tableLength);
  let permission; let userPermission = [];

  for (let i = 0; i < tableLength; i++) {
    userPermission.push(databaseParsed.table.rows[i].c[columnNeeded].v)
  }

  const userMail = Session.getActiveUser().getEmail();
  if (userPermission !== '' && userPermission.indexOf(userMail) > -1) { permission = true; } else { permission = false; }
  // ui().alert(permission)
  Logger.log(permission)

  return permission;

*/

/**
 * botBrainSave
 * Guarda las hojas de un documento de Sheets en formato CSV y las sube a un bucket de Cloud Storage.
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

/**
 * fExportXML
 * Función en desarrollo.
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
