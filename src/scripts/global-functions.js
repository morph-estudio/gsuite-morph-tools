/**
 * waiting
 * Pausa el script por milisegundos, lo que en ocasiones permite evitar bloqueos
 */
function waiting(ms) {
  Utilities.sleep(ms);
}

/**
 * Útiles de programación
 * Función para transponer o simplificar arrays
 */
 function transpose(a) {
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}

function flatten(arrayOfArrays) {
  return [].concat.apply([], arrayOfArrays);
}

/**
 * checkIfSheetIsEmpty
 * Empty Sheet Checker
 */
 function checkIfSheetIsEmpty(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow == 0) {
    return true;
  }
  return false;
}

/**
 * isValidHttpUrl
 * Chequea si una URL es válida
 */
function isValidHttpUrl(str) {
  let pattern = new RegExp('^(https?:\\/\\/)?' // protocol
    + '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' // domain name
    + '((\\d{1,3}\\.){3}\\d{1,3}))' // OR ip (v4) address
    + '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' // port and path
    + '(\\?[;&a-z\\d%_.~+=-]*)?' // query string
    + '(\\#[-a-z\\d_]*)?$', 'i'); // fragment locator
  return !!pattern.test(str);
}

/**
 * isImage
 * Checkea si un documento es imagen
 */
function isImage(url) {
  return /\.(jpg|jpeg|png|webp|avif|gif|svg)$/.test(url);
}

/**
 * openExternalUrlFromMenu
 * Abre una página externa desde Google Apps Script
 */
function openExternalUrlFromMenu(link) {
  let oeufmURL = `${link}`;
  let oeufmHTML = HtmlService.createHtmlOutput('<html><script>'
  + 'window.close = function(){window.setTimeout(function() {google.script.host.close()},9)};'
  + `var a = document.createElement("a"); a.href="${oeufmURL}"; a.target="_blank";`
  + 'if(document.createEvent){'
  + 'var event=document.createEvent("MouseEvents");'
  + 'if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'
  + 'event.initEvent("click",true,true); a.dispatchEvent(event);'
  + '}else{ a.click() }'
  + 'close();'
  + '</script>'
  // Offer URL as clickable link in case above code fails.
  // eslint-disable-next-line max-len
  + `<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="${oeufmURL}" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>`
  + '<script>google.script.host.setHeight(40);google.script.host.setWidth(410) </script>'
  + '</html>')
    .setWidth(110).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(oeufmHTML, 'Abriendo enlace externo... asegúrate de no tener las ventanas emergentes bloqueadas.');
}

/**
 * getIdFromUrl
 * Obtiene la ID de un documento de Google a partir de su dirección URL
 */
function getIdFromUrl(url) { return url.match(/[-\w]{25,}(?!.*[-\w]{25,})/); }

function getIdFromUrlDeprecated(url) {
  let match = url.match(/([a-z0-9_-]{25,})[$/&?]/i);
  return match ? match[1] : null;
}

/**
 * whatAmI
 * Retorna el tipo de objeto
 */
function whatAmI(ob) {
  try {
    // test for an object
    if (ob !== Object(ob)) {
      return {
        type: typeof ob,
        value: ob,
        length: typeof ob === 'string' ? ob.length : null,
      };
    }
    try {
      var stringify = JSON.stringify(ob);
    } catch (err) {
      var stringify = '{"result":"unable to stringify"}';
    }
    return {
      type: typeof ob,
      value: stringify,
      name: ob.constructor ? ob.constructor.name : null,
      nargs: ob.constructor ? ob.constructor.arity : null,
      length: Array.isArray(ob) ? ob.length : null,
    };
  } catch (err) {
    return {
      type: 'unable to figure out what I am',
    };
  }
}

/**
 * searchFile
 * Busca un archivo concreto dentro de una carpeta
 */
function searchFile(fileName, folderId) {
  let files = [];
  // Look for file in current folder
  const folderFiles = DriveApp.getFolderById(folderId).getFiles();
  while (folderFiles.hasNext()) {
    const folderFile = folderFiles.next(); 
    if (folderFile.getName() === fileName) {
      files.push(folderFile);
    }
  }
  // Recursively look for file in subfolders
  const subfolders = DriveApp.getFolderById(folderId).getFolders(); 
  while (subfolders.hasNext()) {
    files = files.concat(searchFile(fileName, subfolders.next().getId()));
  }
  return files;
}

/**
 * keepNewestFilesOfEachNameInAFolder
 * Borrar archivos duplicados en una carpeta (elimina el más antiguo)
 */
function keepNewestFilesOfEachNameInAFolder(folder) {
  const files = folder.getFiles();
  let fO = { pA: [] };
  let keep = [];
  while (files.hasNext()) {
    let file = files.next();
    let n = file.getName();
    //Organize file info in fO
    if (!fO.hasOwnProperty(n)) {
      fO[n] = [];
      fO[n].push(file);
      fO.pA.push(n);
    } else {
      fO[n].push(file);
    }
  }
  //Sort each group with same name
  fO.pA.forEach(n => {
    fO[n].sort((a, b) => {
      let va = new Date(a.getDateCreated()).valueOf();
      let vb = new Date(b.getDateCreated()).valueOf();
      return vb - va;
    });
    //Keep the newest one and delete the rest
    fO[n].forEach((f, i) => {
      if (i > 0) {
        f.setTrashed(true)
      }
    });
  });
}

/**
 * getSplitA1Notation
 * Separa las letras y números de una notación A1 de Google Sheets.
 */
function getSplitA1Notation(cell) {
  let splitArray = cell.split(/([0-9]+)/);
  return splitArray;
}

/**
 * numToCol
 * Return the letter corresponding to a column index in a sheet
 */
 function numToCol(num) {
  var col = "";
  while (num > 0) {
    var remainder = (num - 1) % 26;
    col = String.fromCharCode(65 + remainder) + col;
    num = Math.floor((num - remainder) / 26);
  }
  return col;
}

/**
 * letterToColumn
 * Return the column number based on the letter input for a Google Sheets sheet.
 */
function letterToColumn(letter) {
  let column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

/**
 * getShiftedLetter
 * Return the letter of a column shifted a number of times.
 */
function getShiftedLetter(letter, shiftNumber) {
  var charCode = letter.charCodeAt(0);
  var newCharCode = charCode + shiftNumber;
  return String.fromCharCode(newCharCode);
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
 * deleteAllSheetNotes
 * Delete all cell notes in the selected sheet
 */
function deleteAllSheetNotes(sh) {
  let notes = sh.getNotes();
  for (var i = 0; i < notes.length; i++) {
    sh.getRange(notes[i].getRow(), notes[i].getColumn()).clearNote();
  }
}

/**
 * setColumnWidths
 * Establece el ancho de las columnas en una hoja; el argumento columnWidths debe ofrecerse como objeto
 */
function setColumnWidths(sheet, columnWidths) {
  for (const column of Object.keys(columnWidths)) {
    const [start, end] = column.split(":").map(Number); 
    const width = columnWidths[column]; // Logger.log(`start: ${start}`); Logger.log(`end: ${end}`); Logger.log(`width: ${width}`);
    if (isNaN(end)) {
      sheet.setColumnWidth(start, width);
    } else {
      for (let i = start; i <= end; i++) {
        sheet.setColumnWidth(i, width);
      }
    }
  }
}

/**
 * setCustomRowHeight
 * Establece un alto para todas las filas de una hoja
 */
function setCustomRowHeight(height, sh) {
  sh = sh || SpreadsheetApp.getActiveSheet();
  for (let i = 1; i < sh.getMaxRows() + 1; i++) {
    sh.setRowHeight(i, height);
  }
}

/**
 * getFirstCellA1Notation
 * Return the A1 Notation of the first cell with values in a sheet
 */
function getFirstCellA1Notation(sh) {
  var sheet = sh;
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  for (var row = 0; row < values.length; row++) {
    for (var col = 0; col < values[row].length; col++) {
      if (values[row][col] != "") {
        var rowNumber = row + 1;
        var columnNumber = col + 1;
        var a1Notation = sheet.getRange(rowNumber, columnNumber).getA1Notation();
        return a1Notation;
      }
    }
  }
  return "";
}

/**
 * createCollapseGroup
 * Makes easy to create collapse groups in a sheet
 */
function createCollapseGroup(sh, columnIndex, shiftDepth, columnRange) {
  let group;
  try {
    group = sh.getColumnGroup(columnIndex, shiftDepth);
    group.collapse();

  } catch (error) {
    sh.getRange(columnRange).shiftRowGroupDepth(shiftDepth);
    group = sh.getColumnGroup(columnIndex, shiftDepth)
    group.collapse();
  }
}

/**
 * getLastDataRow
 * Get last row in a single column
 */
 function getLastDataRow(sh, column) {
  var lastRow = sh.getLastRow();
  var range = sh.getRange(column + lastRow);
  if (range.getValue() !== "") {
    return lastRow;
  } else {
    return range.getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
  }              
}

/**
 * getLastDataRowIndex
 * Index of the last row with data in a sheet
 */
function getLastDataRowIndex(sheet) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  var lastRow = sheet.getMaxRows();

  for (var i = lastRow - 1; i >= 0; i--) {
    var row = values[i]; Logger.log(`coso: ${row}`)
    if (row && row.length > 0 && row.join("").length > 0) {
      return i + 1;
    }
  }
  
  return 0;
}

/**
 * eliminarGruposColumnas
 * Delete all column groups in a sheet
 */
function eliminarGruposColumnas(hoja) {

}

/**
 * importRangeToken
 * This function set automatic permission for ImportRange functions. TokenID is the ID of destination Google Sheet
 */
function importRangeToken(ss_id, tokenID) { // 
  let url = `https://docs.google.com/spreadsheets/d/${ss_id}/externaldata/addimportrangepermissions?donorDocId=${tokenID}`;
  let token = ScriptApp.getOAuthToken();
  let params = {
    method: 'post',
    headers: {
      Authorization: `Bearer ${token}`,
    },
    muteHttpExceptions: true,
  };

  UrlFetchApp.fetch(url, params);
}
