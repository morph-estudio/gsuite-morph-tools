function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem('G-Suite Morph Tools Panel', 'sidebarIndex')
    .addItem('Morph Document Studio', 'comingSoon')
    //.addItem('Gestor de hojas', 'sidebarSD')
    .addSeparator()
    .addItem('Changelog', 'sidebarChangelog')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

let sh = function() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet() }
let ui = function() {
  return SpreadsheetApp.getUi() }
const ss = function() {
  return SpreadsheetApp.getActiveSpreadsheet() }

function sidebarIndex() {
  let barTitleIndex = '⚙️ G-Suite Morph Tools (I+D)';
  let hs1 = HtmlService.createTemplateFromFile('public/index').evaluate().setTitle(barTitleIndex);
  ui().showSidebar(hs1);
}
function sidebarDS() {
  let barTitleDS = '📑 Morph Document Studio (I+D)';
  let hs2 = HtmlService.createTemplateFromFile('public/document-studio').evaluate().setTitle(barTitleDS);
  ui().showSidebar(hs2);
}
function sidebarSD() {
  let barTitleSD = '📑 Gestor de hojas by Morph (I+D)';
  let hs3 = HtmlService.createTemplateFromFile('public/sheet-deleter').evaluate().setTitle(barTitleSD);
  ui().showSidebar(hs3);
}
function sidebarChangelog() {
  let link = 'https://github.com/morph-estudio/gsuite-morph-tools/blob/main/CHANGELOG.md';
  openExternalUrlFromMenu(link);
}
function comingSoon() {
  Browser.msgBox('Herramienta en desarrollo', 'Morph Document Studio estará disponible en la próxima versión de Gsuite Morph Tools.', Browser.Buttons.OK);
}

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
