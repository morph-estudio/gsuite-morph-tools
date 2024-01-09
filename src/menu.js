/**
 * GLOBAL VARIABLES AND FUNCTIONS
 */

var gmtVersion = '3.0.0';
var morphDivision = '(I+D)';
var morphDev = '(Devs)';

var titleIX = 'Herramientas de Google Suite'; var barTitleIX = `💡 ${titleIX} ${gmtVersion}`;
var titleSM = 'Gestor de hojas de Sheets'; var barTitleSM = `📋 ${titleSM} ${gmtVersion} ${morphDivision}`;
var titleDS = 'Morph Document Studio'; var barTitleDS = `⚱️ ${titleDS} ${gmtVersion} ${morphDivision}`;
var titleLG = 'Registros Morph'; var barTitleLG = `📝 ${titleLG} ${gmtVersion} ${morphDivision}`;
var titlePC = 'Gestión de cuadros'; var barTitlePC = `✨ ${titlePC} ${gmtVersion} ${morphDivision}`;
var titleCL = 'Guía de estilo'; var barTitleCL = `🎨 ${titleCL} ${gmtVersion} ${morphDivision}`;

var ss = function() {
  return SpreadsheetApp.getActiveSpreadsheet() }
var sh = function() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet() }
var ui = function() {
  return SpreadsheetApp.getUi() }

/**
 * MAIN MENU ENGINE
 */

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem(titlePC, 'sidebarPC')
    .addItem(titleIX, 'sidebarIndex')
    .addItem(titleSM, 'sidebarSM')
    .addItem(titleDS, 'sidebarDS')
    .addSeparator()
    .addItem(titleLG, 'sidebarLG')
    .addSeparator()
    .addItem('Changelog', 'sidebarChangelog')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/**
 * SIDEBAR FUNCTIONS
 */

/* Browser.msgBox('Herramienta en desarrollo', 'Esta herramienta estará disponible en la próxima versión de Gsuite Morph Tools.', Browser.Buttons.OK); */

function sidebarIndex() {
  let html = HtmlService.createTemplateFromFile('public/main/index');
  html.permission = getDevPermission();
  //html.isAdapted = getDocProperty('adaptedSpreadsheet');
  html.wsNames = getWorksheetNamesArray();
  html = html.evaluate().setTitle(barTitleIX); ui().showSidebar(html);
}

function sidebarLG() {
  let html = HtmlService.createTemplateFromFile('public/main/logger-interno');
  html.permission = getDevPermission();
  html.loggerEntries = getLoggerEntries();
  html = html.evaluate().setTitle(barTitleLG); ui().showSidebar(html);
}

function sidebarPC() {
  let html = HtmlService.createTemplateFromFile('public/main/gestion-cuadros');
  html.savedProperties = getDocProperties();
  html.config = templateSheetConfigObject(true);
  html.wsNames = getWorksheetNamesArray();
  html = html.evaluate().setTitle(barTitlePC); ui().showSidebar(html);
}

function sidebarDS() {
  let html = HtmlService.createTemplateFromFile('public/main/document-studio');
  html.dsProperties = getDocProperties(); html.headerDropdownValues = headerDropdownValues();
  html = html.evaluate().setTitle(barTitleDS); ui().showSidebar(html);
}

function sidebarSM() {
  let html = HtmlService.createTemplateFromFile('public/main/sheet-manager');
  html.wsNames = getWorksheetNames();
  html.wsNamesArray = getWorksheetNamesArray();
  html = html.evaluate().setTitle(barTitleSM); ui().showSidebar(html)
}

function sidebarCL() {
  let html = HtmlService.createTemplateFromFile('public/main/gestion-estilos');
  html.obj = cargarEstilos();
  html = html.evaluate().setTitle(barTitleWIP); ui().showSidebar(html);
}

function sidebarCLDevs() {
  let html = HtmlService.createTemplateFromFile('public/main/gestion-estilos');
  html.obj = cargarEstilos();
  var estilos_sheet = PropertiesService.getDocumentProperties();
  html = html.evaluate().setTitle(`${barTitleCL} Devs`); ui().showSidebar(html);
}

function sidebarChangelog() {
  let link = 'https://github.com/morph-estudio/gsuite-morph-tools/blob/main/CHANGELOG.md';
  openExternalUrlFromMenu(link);
}

/**
 * DOCUMENT PROPERTIES AND DATA STORING FUNCTIONS
 */

function getSavedSheetProperties(rowData) {
  PropertiesService.getDocumentProperties().setProperties(rowData);
}

function saveSheetPropertiesWithArray(rowData, array, arrayName) {
  PropertiesService.getDocumentProperties().setProperties(rowData);
  let jarray = JSON.stringify(array);
  PropertiesService.getDocumentProperties().setProperty(arrayName, jarray);
}

function fastInit(optional) {
  Logger.log('Fast Init makes things faster.')
  if (optional) {
    Logger.log(`Properties: ${JSON.stringify(optional)}`);
  }
}
