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
  const hs1 = HtmlService.createTemplateFromFile('html/index').evaluate().setTitle(barTitle1);
  ui.showSidebar(hs1);
}

function loadSidebar2() {
  const hs2 = HtmlService.createTemplateFromFile('html/document-studio').evaluate().setTitle(barTitle2);
  ui.showSidebar(hs2);
}

function loadSidebar3() {
  const hs3 = HtmlService.createTemplateFromFile('test/sheetDeleterIndex').evaluate().setTitle(barTitle3);
  ui.showSidebar(hs3);
}

function loadSidebarX() {
  let url = 'https://gist.github.com/juampynr/4c18214a8eb554084e21d6e288a18a2c';
  let html = HtmlService.createHtmlOutput('<html><script>'
  + 'window.close = function(){window.setTimeout(function() {google.script.host.close()},9)};'
  + `var a = document.createElement("a"); a.href="${url}"; a.target="_blank";`
  + 'if(document.createEvent){'
  + 'var event=document.createEvent("MouseEvents");'
  + 'if(navigator.userAgent.toLowerCase().indexOf("firefox")>-1){window.document.body.append(a)}'
  + 'event.initEvent("click",true,true); a.dispatchEvent(event);'
  + '}else{ a.click() }'
  + 'close();'
  + '</script>'
  // Offer URL as clickable link in case above code fails.
  // eslint-disable-next-line max-len
  + `<body style="word-break:break-word;font-family:sans-serif;">Failed to open automatically. <a href="${url}" target="_blank" onclick="window.close()">Click here to proceed</a>.</body>`
  + '<script>google.script.host.setHeight(40);google.script.host.setWidth(410) </script>'
  + '</html>')
    .setWidth(110).setHeight(1);
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening Changelog... make sure the pop up is not blocked.');
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
