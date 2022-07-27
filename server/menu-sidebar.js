function onOpen(e) {
  var menu = SpreadsheetApp.getUi().createAddonMenu(); // Or DocumentApp or SlidesApp or FormApp.
  menu.addItem('Panel G-Suite Morph Tools ', 'loadSidebar1');
  menu.addItem('Document Studio', 'loadSidebar2');
  //menu.addItem('Trituradora de papel', 'loadSidebar3');
  menu.addSeparator();
  menu.addItem('Changelog', 'loadSidebarX');
  menu.addToUi();
}

const ui = SpreadsheetApp.getUi(); /**/
const barTitle1 = '💛 G-Suite Morph Tools (I+D)';
const barTitle2 = '📑 Document Studio by Morph (I+D)';
const barTitle3 = '📑 Trituradora de papel by Morph (I+D)';

function loadSidebar1() {
  const hs1 = HtmlService.createTemplateFromFile('client/index').evaluate().setTitle(barTitle1);
  ui.showSidebar(hs1);
}

function loadSidebar2() {
  const hs2 = HtmlService.createTemplateFromFile('client/document-studio').evaluate().setTitle(barTitle2);
  ui.showSidebar(hs2);
}

function loadSidebar3() {
  const hs3 = HtmlService.createTemplateFromFile('test/sheetDeleterIndex').evaluate().setTitle(barTitle3);
  ui.showSidebar(hs3);
}

function loadSidebarX() {
  const hsx = HtmlService.createTemplateFromFile('client/changelog').evaluate().setTitle(barTitle1);
  ui.showSidebar(hsx);
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
