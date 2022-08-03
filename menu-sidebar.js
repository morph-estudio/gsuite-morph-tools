function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createAddonMenu()
    .addItem('Panel G-Suite Morph Tools', 'loadSidebar1')
    .addItem('Document Studio', 'loadSidebar2')
    .addItem('Gestor de hojas', 'loadSidebar3')
    .addItem('Email', 'loadSidebar4')
    .addSeparator()
    .addItem('Changelog', 'loadSidebarCL')
    .addToUi();
}

/**/
const ui = SpreadsheetApp.getUi();
var barTitle1 = '💛 G-Suite Morph Tools (I+D)';
var barTitle2 = '📑 Document Studio by Morph (I+D)';
var barTitle3 = '📑 Gestor de hojas by Morph (I+D)';

function loadSidebar1() {
  var hs1 = HtmlService.createTemplateFromFile('html/index').evaluate().setTitle(barTitle1);
  ui.showSidebar(hs1);
  
}
function loadSidebar2() {
  var hs2 = HtmlService.createTemplateFromFile('html/document-studio').evaluate().setTitle(barTitle2);
  ui.showSidebar(hs2);
}
function loadSidebar3() {
  var hs3 = HtmlService.createTemplateFromFile('test/sheetDeleterIndex').evaluate().setTitle(barTitle3);
  ui.showSidebar(hs3);
}
function loadSidebar4() {
  var hs4 = HtmlService.createTemplateFromFile('html/email').evaluate().setTitle(barTitle2);
  ui.showSidebar(hs4);
}

function loadSidebarCL() {
  var link = 'https://github.com/alsanmorph/gsuite-morph-tools/blob/d2f4a46ebcbea3daada6952819f66dd469fb55ac/CHANGELOG.md';
  openExternalUrlFromMenu(link);
}

function navBarTitle() {
  var navBarTitle = 'Morphies';
  return navBarTitle;
}

/**/

let size;
function cellCounter() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  let cells = 0;
  sheets.forEach((sheet) => {
    cells = cells + sheet.getMaxRows() * sheet.getMaxColumns();
  });
  let division = cells / 10000000 * 100;
  let percentage = +division.toFixed(0);
  return percentage;
}

function cellCounter2() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheets = ss.getSheets();
  let cells = 0;
  sheets.forEach((sheet) => {
    cells = cells + sheet.getMaxRows() * sheet.getMaxColumns();
  });
  let division = cells / 10000000 * 100;
  let percentage = +division.toFixed(0);
  return (`📈 Cada Google Sheets tiene capacidad para diez millones de celdas. Has usado el <strong>${percentage}%</strong> del total con <strong>${cells} celdas</strong>.`);
}
