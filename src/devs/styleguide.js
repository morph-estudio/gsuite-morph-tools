// variable donde se guardan todos los estilos


function colorize(name, mode) {

  /*
  const [background, color] = cols.split('-');
  let sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let cells = sh.getActiveRange();
  cells.setBackgroundColor(`#${background}`)
       .setFontColor(`#${color}`);
*/

  
  var stl = JSON.parse(UrlFetchApp.fetch('https://opensheet.elk.sh/1v5f3X1ShmVCGdT6NdWvmJHcfeP01ptuwHfT1iqM6UQI/ESTILOS').getContentText());
  //var stl = UrlFetchApp.fetch('https://opensheet.elk.sh/1v5f3X1ShmVCGdT6NdWvmJHcfeP01ptuwHfT1iqM6UQI/ESTILOS').getContentText()//.slice(1, -1);
  //const dbColumn = parsedDB.map(i => i[headerName]);
  //let transposeCol = transpose(parsedDB)
  //stl = `{entries: ${stl}}`;





  let sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let cells = sh.getActiveRange();
  let response;

  Logger.log(mode)
  if (mode === 'v-reg') {
      response = stl.find(el => el.estilo === name).normal
      let [background, color] = response.split('-');

    cells.setBackgroundColor(background)
        .setFontColor(color);
  } else if (mode === 'v-bold') {
         response = stl.find(el => el.estilo === name).acentuado
      let [background, color] = response.split('-');
    cells.setBackgroundColor(background)
        .setFontColor(color);
  }
//Logger.log(fetchResponse)
  //Logger.log(fetchResponse.entries.find(el => el.estilo === 'Mariana Blue').normal)




  /*
    let payload = {
      "black": "#000",
      "white": "#fff"
      }
  
  const [background, color] = cols.split('-');
  Logger.log(background); Logger.log(color);
  let selection = ss().getSelection();
  let currentCell = selection.getActiveRange();
  currentCell.setBackgroundColor(`#${background}`).setFontColor(`#${color}`);
  */
}

function guardarEstilo(estilo, styleName) {
  const estilos_sheet = PropertiesService.getDocumentProperties();
  // borramos previamente los estilos
  eliminarEstilo(estilo);

  // obtenemos la celda activa
  var celda = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveCell();  

  //guardamos colores y tamaño
  estilos_sheet.setProperty('colorLetra'+estilo, celda.getFontColor())
               .setProperty('colorFondo'+estilo, celda.getBackground())
               .setProperty('colorNames'+estilo, styleName)
               .setProperty('size'+estilo, celda.getFontSize()+'')
               

  return {  colorFondo: estilos_sheet.getProperty('colorFondo'+estilo),
            colorLetra: estilos_sheet.getProperty('colorLetra'+estilo),
            colorNames: estilos_sheet.getProperty('colorNames'+estilo)        
            }
}

function aplicarEstilo(estilo) {
  const estilos_sheet = PropertiesService.getDocumentProperties();
  //borrarEstilos();
  let sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let celdas = sh.getActiveRange();

  celdas.setFontColor(estilos_sheet.getProperty('colorLetra'+estilo))
        .setBackground(estilos_sheet.getProperty('colorFondo'+estilo))
        .setFontSize(estilos_sheet.getProperty('size'+estilo))
}

function cargarEstilos() {
  const estilos_sheet = PropertiesService.getDocumentProperties();
  return estilos_sheet.getProperties();
}

function eliminarEstilo(estilo) {
  //colores
  const estilos_sheet = PropertiesService.getDocumentProperties();
  estilos_sheet.deleteProperty('colorLetra'+estilo);
  estilos_sheet.deleteProperty('colorFondo'+estilo);
  estilos_sheet.deleteProperty('sizeFuente'+estilo);
}

function borrarEstilos() {
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getActiveRange().clear({formatOnly: true});
}

function eliminarEstilos() {
  const estilos_sheet = PropertiesService.getDocumentProperties();
  estilos_sheet.deleteAllProperties();
  sidebarCL();
}

function autoLoader() {
  return null;
}