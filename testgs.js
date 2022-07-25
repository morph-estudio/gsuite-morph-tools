


function appendFormData(rowData) {

var ss = SpreadsheetApp.getActive();
var sh = SpreadsheetApp.getActiveSheet();
var ws = ss.getSheetByName('DS') || ss.insertSheet('DS', 1);

var list = [rowData.templateID,rowData.destFolder,rowData.exportFormat];
ws.appendRow(list);

return true;

}


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
  PropertiesService.getScriptProperties().setProperty('run', 'STOP');
  return 'Kill Signal Issued';
}

function doGet(e) {
  PropertiesService.getScriptProperties().setProperty('run', 'OK');
  return HtmlService.createHtmlOutputFromFile('index').setSandboxMode(HtmlService.SandboxMode.IFRAME);
}












function listFilesInFolder() {

  var ui = SpreadsheetApp.getUi();
  var userprompt1 = ui.prompt("Introduce el ID de la carpeta:");
  var srcFolderID = userprompt1.getResponseText();

  var folder = DriveApp.getFolderById(srcFolderID).next();
  var contents = folder.getFiles();

  var file,
    data,
    sheet = SpreadsheetApp.getActiveSheet();
  sheet.clear();

  sheet.appendRow(['Name', 'Date', 'Size', 'URL', 'Download', 'Description', 'Type']);

  for (var i = 0; i < contents.length; i++) {
    file = contents[i];

    if (file.getFileType() == 'SPREADSHEET') {
      continue;
    }

    data = [
      file.getName(),
      file.getDateCreated(),
      file.getSize(),
      file.getUrl(),
      'https://docs.google.com/uc?export=download&confirm=no_antivirus&id=' + file.getId(),
      file.getDescription(),
      file.getFileType().toString(),
    ];

    sheet.appendRow(data);
  }
}



function addImportrangePermission() {
  // id of the spreadsheet to add permission to import
  var ssId = SpreadsheetApp.getActiveSpreadsheet().getId();

  // donor or source spreadsheet id, you should get it somewhere
  var donorId = '1GrELZHlEKu_QbBVqv...';

  // adding permission by fetching this url
  var url = `https://docs.google.com/spreadsheets/d/${ssId}/externaldata/addimportrangepermissions?donorDocId=${carpetabaseid}`;

  const token = ScriptApp.getOAuthToken();

  const params = {
    method: 'post',
    headers: {
      Authorization: 'Bearer ' + token,
    },
    muteHttpExceptions: true
  };

  UrlFetchApp.fetch(url, params);
}


/*

function createMenu(){
  const ui = SpreadsheetApp.getUi()
  const menu = ui.createMenu("Utilities")
  menu.addItem("Delete Worksheets","loadSidebar")
  menu.addToUi()
}

function onOpen(){
  createMenu()
}



function doGet() {
  return HtmlService.createTemplateFromFile("index").evaluate()
}

function getData(){

  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const ws = ss.getSheetByName("Sheet1")
  const product = ws.getRange("A1").getValue()
  const qty = ws.getRange("B1").getValue()

  return {product:product,quantity:qty}
}

function importcogs() {
  Logger.log("import begin");
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var urlsheet = ss.getSheetByName("GetInfo");
  var request = urlsheet.getRange(5,2).getValue();
  Logger.log(request);

  var response = UrlFetchApp.fetch(request);
  Logger.log("download data finish");
  Logger.log(response.getContentText());


  var sheet = ss.getSheetByName("Data");
  var obj = JSON.parse(response);
  let vs = obj.data.map(o => Object.values(o));//data
  vs.unshift(Object.keys(obj.data[0]));//add header
  sheet.getRange(1,1,vs.length, vs[0].length).setValues(vs);//output to spreadsheet
}



function alertMessageEmoji() {
  SpreadsheetApp.getUi().alert("⚠️ You're about to share a file externally", "You're about to share this document with bob@example.com, who is not a pre-approved recipient. Are you sure?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
}


function getSuperHero() {
  return {name: "SuperGeek",  catch_phrase: "Don't worry ma'am, I come from the Internet" };
}



*/
