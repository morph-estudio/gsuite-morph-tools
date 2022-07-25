function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}


function autoFolderTree() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getActiveSheet();
  var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('FOLDERTREE');
  var niveles = [1,2,3,4,5,6,7];
  ws.activate();

  var cell = ws.getRange('B3');
  if (cell.isBlank()) {

    var ui = SpreadsheetApp.getUi();

    var result = ui.prompt(
        'ID Carpeta',
        'Introduce el ID de la carpeta donde crear la estructura:',
        ui.ButtonSet.OK_CANCEL);

    var button = result.getSelectedButton();
    var userGetID = result.getResponseText();
    if (button == ui.Button.OK) {
      // call function and pass the value
      cell.setValue(userGetID);
    }
  } else {}

  for (n in niveles){

    var levelInput = niveles[n];
    var Level =  levelInput * 2 + 1;
    var numRows = ws.getLastRow(); // Number of rows to process
    var dataRange = ws.getRange(3, Number(Level)-1, numRows, Number(Level)); //startRow, startCol, endRow, endCol
    var data = dataRange.getValues();
    var parentFolderID = new Array();

    for(var i in data)
      {
      parentFolderID[i] = data [i][0];
      if (data [i][0] == "")
        {
        parentFolderID[i] = parentFolderID[i-1];
        }
      }

    for(var i in data){
    if (data [i][1] !== "")
      {
      var theParentFolder = DriveApp.getFolderById(parentFolderID[i]);
      var folderName = data[i][1];
      var theChildFolder = theParentFolder.createFolder(folderName);
      var newFolderID = ws.getRange(Number(i)+3,Number(Level)+1);
      var folderIdValue = theChildFolder.getId();
      newFolderID.setValue(folderIdValue);
      var addLink = ws.getRange(Number(i)+3,Number(Level));
      var value = addLink.getDisplayValue();
      addLink.setValue('=hyperlink("https://drive.google.com/corp/drive/folders/'+ folderIdValue +'","' + value + '")');
      }
    }

  }


  }



function plantilla_AutoFolderTree() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName('FOLDERTREE') || ss.insertSheet('FOLDERTREE', 1);

  ws.clear().clearFormats();
  ws.setFrozenRows(0);
  ws.getRange(1, 1, ws.getMaxRows(), ws.getMaxColumns()).setFontSize(12).setFontFamily("Montserrat").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("middle");


// LEVELS OF STRUCTURE

  var list = [["LEVEL 1", "", "LEVEL 2", "", "LEVEL 3", "", "LEVEL 4", "", "LEVEL 5", "", "LEVEL 6", "", "LEVEL 7"]];
  ws.getRange(1,3,1,13).setValues(list).setBackground("#434343").setFontColor("#fff");
  ws.getRange('B1').setValue('ID BASE FOLDER').setBackground("#6AA84F").setBorder(true, true, true, true, true, true , "#6AA84F", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor("#fff");;
  ws.getRange('B3').setBackground("#D9EAD3").setBorder(true, true, true, true, true, true , "#6AA84F", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor("#6AA84F");

  ws.hideColumns(4);  ws.hideColumns(6);  ws.hideColumns(8);  ws.hideColumns(10);  ws.hideColumns(12);  ws.hideColumns(14);  ws.hideColumns(16);

  var list2 = [["CODE 1","CODE 2","CODE 3","CLIENT","LOCATION","PROJECT NAME"], ["P00000","'01","AEI","Cliente","Madrid","El Encinar"]];
  ws.getRange(1,18,2,6).setValues(list2);
  ws.getRange(1,18,1,6).setBackground("#BF9000").setBorder(true, true, true, true, true, true, "#BF9000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor("#fff");
  ws.getRange(2,18,1,6).setBackground("#FFF2CC").setBorder(true, true, true, true, true, true, "#BF9000", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor("#A61C00").setHorizontalAlignment("center");

  var cell = ws.getRange('T2');
  var range = ws.getRange('B1:B10');
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(["AEI","E","I","IINT","I+D"])
      .build();
  cell.setDataValidation(rule);

  ws.getRange(1, 1, 1, ws.getMaxColumns()).setFontWeight("bold").setFontSize(14).setHorizontalAlignment("center");


// TEMPLATE STRUCTURE

  ws.getRange('C3').setValue('=CONCATENATE(R2," (",S2,"-",T2,"-",U2,"-",V2,") ",W2)');

  ws.getRange('E4').setValue('=CONCATENATE(R2," 1 Trabajo")');
  ws.getRange('E11').setValue('=CONCATENATE(R2," 2 Doc Previa")');
  ws.getRange('E19').setValue('=CONCATENATE(R2," 3 Comunicación")');
  ws.getRange('E35').setValue('=CONCATENATE($R$2," 4 Proyectos entregados")');
  ws.getRange('E36').setValue('=CONCATENATE($R$2," 5 Publicacion")');
  ws.getRange('E37').setValue('=CONCATENATE($R$2," 6 Obra")');
  ws.getRange('E40').setValue('=CONCATENATE($R$2," 7 Press")');
  ws.getRange('E41').setValue('=CONCATENATE($R$2," 8 Asistencia")');

  var listnv3 = [['=CONCATENATE($R$2," ","Arquitectura")'],['=CONCATENATE($R$2," ","Breeam")'],['=CONCATENATE($R$2," ","Estructuras")'],['=CONCATENATE($R$2," ","Instalaciones")'],['=CONCATENATE($R$2," ","Interiorismo")'],['=CONCATENATE($R$2," ","Mediciones")']];

  ws.getRange(5,7,6,1).setValues(listnv3);

  var listnv3_2 = [['=CONCATENATE($R$2," ","01"," ","Doc recibida cliente")'],['=CONCATENATE($R$2," ","02"," ","Normativa")'],['=CONCATENATE($R$2," ","03"," ","Web")'],['=CONCATENATE($R$2," ","04"," ","Cartografía")'],['=CONCATENATE($R$2," ","05"," ","Fotos")'],['=CONCATENATE($R$2," ","06"," ","Estudio de mercado")'],['=CONCATENATE($R$2," ","07"," ","Doc recibida cliente")']];

  ws.getRange(12,7,7,1).setValues(listnv3_2);

  ws.getRange('G11').setValue('=CONCATENATE($R$2," ","01"," ","Cliente")')
  ws.getRange('G23').setValue('=CONCATENATE($R$2," ","02"," ","Morph Interno")')
  ws.getRange('G26').setValue('=CONCATENATE($R$2," ","Ayuntamiento")')
  ws.getRange('G29').setValue('=CONCATENATE($R$2," ","Renderista")')
  ws.getRange('G32').setValue('=CONCATENATE($R$2," ","Obra")')
  ws.getRange('G38').setValue('01_Fotografías')
  ws.getRange('G39').setValue('02_Actas de obra');

  var listnv4 = [['Enviado'],['Recibido']];

// FORMATTING

  ws.setFrozenRows(1);

  ws.getRange(21,9,2,1).setValues(listnv4)
  ws.getRange(24,9,2,1).setValues(listnv4)
  ws.getRange(27,9,2,1).setValues(listnv4)
  ws.getRange(30,9,2,1).setValues(listnv4)
  ws.getRange(33,9,2,1).setValues(listnv4);

  ws.setColumnWidth(1, 25)
  ws.setColumnWidth(2, 200)
  ws.setColumnWidth(3, 200)
  ws.setColumnWidth(5, 200)
  ws.setColumnWidth(7, 200)
  ws.setColumnWidth(9, 200)
  ws.setColumnWidth(11, 200)
  ws.setColumnWidth(13, 200)
  ws.setColumnWidth(15, 200)
  ws.setColumnWidth(17, 40)
  ws.setColumnWidth(22, 200)
  ws.setColumnWidth(23, 200);


  removeEmptyColumns();
  deleteEmptyRows();
  SpreadsheetApp.flush();

  ws.activate();

}
