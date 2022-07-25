function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

// NEW GOOGLE SHEET

function createSpreadsheet() {
  SpreadsheetApp.create("SHEET CREADA CON MORPH TOOLS");
}

// WAIT FUNCTION

function waiting(ms){
  Utilities.sleep(ms);
}

// MORPH COLOR BACKGROUND

function colorMe() {
  var spreadsheet = SpreadsheetApp.getActive();
  var selection = spreadsheet.getSelection();
  var currentCell = selection.getActiveRange();
  {
  currentCell.setBackgroundColor('#f1cb50');
  }
}

// MORPH PALETTE

function drawPalette() {
  SpreadsheetApp.getActiveSpreadsheet().insertSheet("chart");
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("chart").getRange(1,1,1,104).setBackgrounds([
  [

  "#F2F2F2","#FFF3F3","#EEECE1","#D4C9C6","#CDC8CE","#BDBDBD","#A5A5A5","#7F7F7F", // PALETA GREY / BROWN
  "#FFF1CB","#FFE9AD","#FFFAB1","#FFFFAE","#FFE699","#FFE499","#FFFF99","#FCF58F","#FFEB84","#FFFB84","#FFD666","#FED166","#F7CB4D","#ECFF49", // PALETA YELLOW
  "#FFD7AE","#ED956F","#ED7D31","#FF6F31","#F57C00", // PALETA ORANGE
  "#FFEEF6","#FFE0EF","#E7CFCF","#F9CCC4","#E6BFD2","#FAC5DC","#EEBBD1","#E6B3B3","#F1ADC9","#E2ABC5","#EEA3C6","#F097A3","#E78BB6","#FD7D8F","#F08090","#DD808D","#E67C73","#FF6565","#D16969","#E04C4C", // PALETA RED
  "#D7D7FF","#FFDDFF","#E2C6FF","#FFC4FF","#E0A7EB","#D797E4","#FF8FFF","#FF62FF", // PALETA PINK
  "#F0FDFF","#DBE5F1","#E0F7FA","#ECFFF5","#E0FFFC","#DEFFFF","#D7FDFF","#CCF2FF","#C4F7F7","#B9EBFD","#B7D4F0","#C5FFFF","#C4FDF8","#A8DFF3","#A3C3C9","#AFF5EF","#94E6DF","#7FCFEC","#74BFDB","#7BDFD6","#71B9D3","#62D1C7","#5CA4BD","#31AA9F","#3D8DA8","#1B8B81","#1B6F8B", // PALETA BLUE
  "#EAF8D0","#C7D9B7","#D5E6B6","#E4EBA9","#CCFFCC","#A9DFC5","#CCCF96","#BEE2A6","#B8F1B8","#CDF39C","#CEED8A","#BCDA85","#A7DB85","#CEFF70","#A4C26E","#9DE063","#41CF7F","#8FA369","#63BE7B","#57BB8A","#6AA369","#008000" // PALETA GREEN

  ]
  ]);
  var ssch = SpreadsheetApp.getActive();
  var sheet = ssch.getSheetByName("chart");
  ssch.deleteSheet(sheet);
}

// MORPH USERLIST

function getDomainUsersList() {
  var users = [];
  var options = {
    domain: 'morphestudio.es', // Google Workspace domain name
    customer: 'my_customer',
    maxResults: 200,
    projection: 'basic', // Fetch basic details of users
    viewType: 'domain_public',
    orderBy: 'email', // Sort results by users
  };

  do {
    var response = AdminDirectory.Users.list(options);
    response.users.forEach(function (user) {
      users.push([user.name.fullName, user.primaryEmail]);
    });

    // For domains with many users, the results are paged
    if (response.nextPageToken) {
      options.pageToken = response.nextPageToken;
    }
  } while (response.nextPageToken);

  // Insert data in a spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('MORPH MEMBERS') || ss.insertSheet('MORPH MEMBERS', 1);
  //SpreadsheetApp.getActiveSpreadsheet().moveActiveSheet(ss.getNumSheets());
  ws.getRange(1, 1, users.length, users[0].length).setValues(users);
  ws.setColumnWidth(1, 250);
  ws.setColumnWidth(2, 280);

  var rangememb = ws.getRange("A1:B500")
  .setFontSize(12)
  .setFontFamily("Montserrat")
  ;
  var rangememb2 = ws.getRange("A1:A500")
  .setFontWeight("bold")
  ;

  ws.activate();
  deleteEmptyRows();
  removeEmptyColumns();
}

// REMOVE ROWS AND COLUMNS

function deleteFullEmptyRows() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getActiveSheet();
  var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  var row = sheet.getLastRow();

  while (row > 2) {
    var rec = data.pop();
    if (rec.join('').length === 0) {
      sheet.deleteRow(row);
    }
    row--;
  }

  var maxRows = sheet.getMaxRows();
  var lastRow = sheet.getLastRow();
  if (maxRows - lastRow != 0) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

function deleteEmptyRows() {
  var sh = SpreadsheetApp.getActiveSheet();
  var maxRows = sh.getMaxRows();
  var lastRow = sh.getLastRow();
  var checker2 = maxRows-(maxRows-lastRow);
  if (checker2 != maxRows){
  sh.deleteRows(lastRow+1, maxRows-lastRow);
  }
}

function removeEmptyColumns() {
  var sh = SpreadsheetApp.getActiveSheet();
  var maxCol = sh.getMaxColumns();
  var lastCol = sh.getLastColumn();
  var checker = maxCol-(maxCol-lastCol);
  if (checker != maxCol){
  sh.deleteColumns(lastCol+1, maxCol-lastCol);
  }
}

// MERGE COLUMNS BY HEADER

function merge_Columns() {
  var transpose = (ar) => ar[0].map((_, c) => ar.map((r) => r[c]));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var values = sh.getDataRange().getValues();
  var temp = [
    ...transpose(values)
      .reduce(
        (m, [a, ...b]) => m.set(a, m.has(a) ? [...m.get(a), ...b] : [a, ...b]),
        new Map()
      )
      .values(),
  ];
  var res = transpose(temp);
  sh.clearContents();
  sh.getRange(1, 1, res.length, res[0].length).setValues(res);
}

function getIdFromUrl(url) { return url.match(/[-\w]{25,}(?!.*[-\w]{25,})/);}

function getIdFromUrls(url) {
  var match = url.match(/([a-z0-9_-]{25,})[$/&?]/i);
  return match ? match[1] : null;
}

/* **************
FUNCIONES FUERA DEL SISTEMA / PRUEBAS / ETC
 ************** */

// SCRIPTING TRIGGERS

function startScript() {
  do {
    Logger.log('Script running');
    Utilities.sleep(5000);
  } while (keepRunning());
  return 'OK';
}

function keepRunning() {
  var status = PropertiesService.getScriptProperties().getProperty('run') || 'OK';
  return status === 'OK' ? true : false;
}

function stopScript() {
  PropertiesService.getScriptProperties(actualizarCuadro);
  PropertiesService.getScriptProperties(actualizarCuadro).setProperty("stop", "stop");
}

function doGet(e) {
  PropertiesService.getScriptProperties().setProperty('run', 'OK');
  return HtmlService.createHtmlOutputFromFile('index').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// IDENTIFY OBJECTS

function whatAmI (ob) {

  try {
    // test for an object
    if (ob !== Object(ob)) {
        return {
          type:typeof ob,
          value: ob,
          length:typeof ob === 'string' ? ob.length : null
        } ;
    }
    else {
      try {
        var stringify = JSON.stringify(ob);
      }
      catch (err) {
        var stringify = '{"result":"unable to stringify"}';
      }
      return {
        type:typeof ob ,
        value : stringify,
        name:ob.constructor ? ob.constructor.name : null,
        nargs:ob.constructor ? ob.constructor.arity : null,
        length:Array.isArray(ob) ? ob.length:null
      };
    }
  }
  catch (err) {
    return {
      type:'unable to figure out what I am'
    } ;
  }
}
