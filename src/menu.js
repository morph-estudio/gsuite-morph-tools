/**
 * Global Variables + Global Functions
 */

const morphDivision = '(I+D)'; const morphDev = '(Devs)'; const gmtVersion = '1.5.7'
const titleIX = 'G-Suite Morph Tools'; const barTitleIX = `💡 ${titleIX} ${gmtVersion}`;
const titleDS = 'Morph Document Studio'; const barTitleDS = `✨ ${titleDS} ${morphDivision}`;
const titleSM = 'Gestor de hojas'; const barTitleSM = `📋 ${titleSM} ${morphDivision}`;
const titleCL = 'Guía de estilo'; const barTitleCL = `🎨 ${titleCL} ${morphDivision}`;


const ss = function() {
  return SpreadsheetApp.getActiveSpreadsheet() }
const sh = function() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet() }
const ui = function() {
  return SpreadsheetApp.getUi() }

/**
 * Menu Basic Engine
 */

function onOpen(e) {
  
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem(titleIX, 'sidebarIndex')
    .addItem(titleDS, 'sidebarDS')
    .addItem(titleSM, 'sidebarSM')
    //.addItem(titleCL, 'sidebarCL')
    .addSeparator()
    .addItem('Changelog', 'sidebarChangelog')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

function sidebarIndex() {
  let html = HtmlService.createTemplateFromFile('public/index');
  html.protection = getDevPermission('DevAreaMails'); // html.navBarHEX = '#FFCCBC';
  html.isAdapted = getDocProperty('adaptedSpreadsheet'); // html.navBarHEX = '#FFCCBC';
  html.wsNames = getWorksheetNamesArray();
  html = html.evaluate().setTitle(barTitleIX); ui().showSidebar(html);
}

function sidebarDS() {
  // Browser.msgBox('Herramienta en desarrollo', 'Morph Document Studio estará disponible en la próxima versión de G-Suite Morph Tools.', Browser.Buttons.OK);
  let html = HtmlService.createTemplateFromFile('public/document-studio');
  html.dsProperties = getDocProperties(); html.emailDropdown = emailDropdown();
  html = html.evaluate().setTitle(barTitleDS); ui().showSidebar(html);
}

function sidebarSM() {
  let html = HtmlService.createTemplateFromFile('public/sheet-manager');
  html.wsNames = getWorksheetNames();
  html = html.evaluate().setTitle(barTitleSM); ui().showSidebar(html)
}

function sidebarCL() {
  let html = HtmlService.createTemplateFromFile('public/styles-front');
  
  html.obj = cargarEstilos();
  html = html.evaluate().setTitle(barTitleCL); ui().showSidebar(html)
}

function sidebarDSDevs() {
  let html = HtmlService.createTemplateFromFile('public/document-studio');
  html.dsProperties = getDocProperties(); html.emailDropdown = emailDropdown();
  html = html.evaluate().setTitle(`${barTitleDS} Devs`); ui().showSidebar(html);
}

function sidebarSMDevs() {
  let html = HtmlService.createTemplateFromFile('public/sheet-manager');
  html.wsNames = getWorksheetNames();
  html = html.evaluate().setTitle(`${barTitleSM} Devs`); ui().showSidebar(html)
}

function sidebarCLDevs() {
  let html = HtmlService.createTemplateFromFile('public/styles-front');
  html.obj = cargarEstilos();
  var estilos_sheet = PropertiesService.getDocumentProperties();
  html = html.evaluate().setTitle(`${barTitleCL} Devs`); ui().showSidebar(html)
}

function sidebarChangelog() {
  let link = 'https://github.com/morph-estudio/gsuite-morph-tools/blob/main/CHANGELOG.md';
  openExternalUrlFromMenu(link);
}
