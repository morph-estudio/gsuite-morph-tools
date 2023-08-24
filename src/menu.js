/** GLOBAL VARIABLES AND FUNCTIONS */

const gmtVersion = '1.9.0';
const morphDivision = '(I+D)';
const morphDev = '(Devs)';

const titleIX = 'Gsuite Morph Tools'; const barTitleIX = `💡 ${titleIX} ${gmtVersion}`;
const titleSM = 'Gestor de hojas'; const barTitleSM = `📋 ${titleSM} ${morphDivision}`;
const titleDS = 'Morph Document Studio'; const barTitleDS = `✨ ${titleDS} ${morphDivision}`;
const titleLG = 'Registros Morph'; const barTitleLG = `✨ ${titleLG} ${morphDivision}`;
const titleWIP = 'Morph Control Panel'; const barTitleWIP = `✨ ${titleWIP} ${morphDivision}`;
const titleCL = 'Guía de estilo'; const barTitleCL = `🎨 ${titleCL} ${morphDivision}`;

const ss = function() {
  return SpreadsheetApp.getActiveSpreadsheet() }
const sh = function() {
  return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet() }
const ui = function() {
  return SpreadsheetApp.getUi() }

/** MAIN MENU ENGINE */

function onOpen(e) {
  SpreadsheetApp.getUi().createAddonMenu()
    .addItem(titleIX, 'sidebarIndex')
    .addItem(titleSM, 'sidebarSM')
    .addItem(titleDS, 'sidebarDS')
    .addItem(titleLG, 'sidebarLG')
    // .addItem(titleWIP, 'sidebarWIP')
    //.addItem(titleCL, 'sidebarCL')
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

function sidebarIndex() {
  let html = HtmlService.createTemplateFromFile('public/main/index');
  html.permission = getDevPermission();
  //html.isAdapted = getDocProperty('adaptedSpreadsheet');
  html.wsNames = getWorksheetNamesArray();
  html = html.evaluate().setTitle(barTitleIX); ui().showSidebar(html);
}

function sidebarLG() {
  /* Browser.msgBox('Herramienta en desarrollo', 'Esta herramienta estará disponible en la próxima versión de Gsuite Morph Tools.', Browser.Buttons.OK); */
  
  let html = HtmlService.createTemplateFromFile('public/main/logger-interno');
  html.permission = getDevPermission();
  html.loggerEntries = getLoggerEntries();
  html = html.evaluate().setTitle(barTitleLG); ui().showSidebar(html);
  
}

function sidebarWIP() {
  let html = HtmlService.createTemplateFromFile('public/main/cuadroswip');
  html.savedProperties = getDocProperties();
  html.config = templateSheetConfigObject();
  html.wsNames = getWorksheetNamesArray();
  html = html.evaluate().setTitle(barTitleLG); ui().showSidebar(html);
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
  let html = HtmlService.createTemplateFromFile('public/main/styles-front');
  html.obj = cargarEstilos();
  html = html.evaluate().setTitle(barTitleWIP); ui().showSidebar(html);
}

function sidebarDSDevs() {
  let html = HtmlService.createTemplateFromFile('public/main/document-studio');
  html.dsProperties = getDocProperties(); html.headerDropdownValues = headerDropdownValues();
  html = html.evaluate().setTitle(`${barTitleDS} Devs`); ui().showSidebar(html);
}

function sidebarSMDevs() {
  let html = HtmlService.createTemplateFromFile('public/main/sheet-manager');
  html.wsNames = getWorksheetNames();
  html = html.evaluate().setTitle(`${barTitleSM} Devs`); ui().showSidebar(html)
}

function sidebarCLDevs() {
  let html = HtmlService.createTemplateFromFile('public/main/styles-front');
  html.obj = cargarEstilos();
  var estilos_sheet = PropertiesService.getDocumentProperties();
  html = html.evaluate().setTitle(`${barTitleCL} Devs`); ui().showSidebar(html)
}

function sidebarChangelog() {
  let link = 'https://github.com/morph-estudio/gsuite-morph-tools/blob/main/CHANGELOG.md';
  openExternalUrlFromMenu(link);
}

/**
 * DOCUMENT PROPERTIES AND DATA STORING FUNCTIONS
 */

function getSavedSheetProperties(rowData) {
  PropertiesService.getDocumentProperties()
    .setProperties(rowData);
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
