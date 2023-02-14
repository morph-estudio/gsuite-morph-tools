/* eslint-disable no-only-tests/no-only-tests */

/**
 * waiting
 * Pausa el script por milisegundos, lo que en ocasiones permite evitar bloqueos
 */
function waiting(ms) {
  Utilities.sleep(ms);
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
 * getSplitA1Notation
 * Separa las letras y números de una notación A1 de Google Sheets.
 */
function getSplitA1Notation(cell) {
  let splitArray = cell.split(/([0-9]+)/);
  return splitArray;
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
