/**
 * Gsuite Morph Tools - Morph autoFolderTree 1.0
 * Developed by alsanchezromero
 * Created on Mon Jul 25 2022
 *
 * Copyright (c) 2022 Morph Estudio
 */

/* eslint-disable guard-for-in */
/* eslint-disable no-restricted-syntax */

function autoFolderTree() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getActiveSheet();
  let niveles = [1, 2, 3, 4, 5, 6, 7];

  sh.activate();

  let cell = sh.getRange('B3');
  if (cell.isBlank()) {
    const ui = SpreadsheetApp.getUi();
    let result = ui.prompt(
      'ID Carpeta',
      'Introduce el ID de la carpeta donde crear la estructura:',
      ui.ButtonSet.OK_CANCEL,
    );

    let button = result.getSelectedButton();
    let userGetID = result.getResponseText();
    if (button == ui.Button.OK) {
      // call function and pass the value
      cell.setValue(userGetID);
    }
  }

  for (n in niveles) {
    let levelInput = niveles[n];
    let Level = levelInput * 2 + 1;
    let numRows = sh.getLastRow(); // Number of rows to process
    let dataRange = sh.getRange(3, Number(Level) - 1, numRows, Number(Level)); // startRow, startCol, endRow, endCol
    let data = dataRange.getValues();
    let parentFolderID = new Array();

    for (let i in data) {
      parentFolderID[i] = data [i][0];
      if (data [i][0] == '') {
        parentFolderID[i] = parentFolderID[i - 1];
      }
    }

    for (let i in data) {
      if (data [i][1] !== '') {
        let theParentFolder = DriveApp.getFolderById(parentFolderID[i]);
        let folderName = data[i][1];
        let theChildFolder = theParentFolder.createFolder(folderName);
        let newFolderID = sh.getRange(Number(i) + 3, Number(Level) + 1);
        let folderIdValue = theChildFolder.getId();
        newFolderID.setValue(folderIdValue);
        let addLink = sh.getRange(Number(i) + 3, Number(Level));
        let value = addLink.getDisplayValue();
        addLink.setValue(`=hyperlink("https://drive.google.com/corp/drive/folders/${folderIdValue}","${value}")`);
      }
    }
  }
}

// TEMPLATE FUNCTION

function autoFolderTreeTpl() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('FOLDERTREE') || ss.insertSheet('FOLDERTREE', 1);

  sh.clear().clearFormats();
  sh.setFrozenRows(0);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setFontSize(12).setFontFamily('Montserrat').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment('middle');

  // Levels of Structure

  let list = [['LEVEL 1', '', 'LEVEL 2', '', 'LEVEL 3', '', 'LEVEL 4', '', 'LEVEL 5', '', 'LEVEL 6', '', 'LEVEL 7']];
  sh.getRange(1, 3, 1, 13).setValues(list).setBackground('#434343').setFontColor('#fff');
  sh.getRange('B1').setValue('ID BASE FOLDER').setBackground('#BF9000').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#fff');
  sh.getRange('B3').setBackground('#FFF2CC').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#BF9000');

  sh.hideColumns(4); sh.hideColumns(6); sh.hideColumns(8); sh.hideColumns(10);
  sh.hideColumns(12); sh.hideColumns(14); sh.hideColumns(16);

  let list2 = [['CODE 1', 'CODE 2', 'CODE 3', 'CLIENT', 'LOCATION', 'PROJECT NAME'], ['P00000', '01', 'AEI', 'Cliente', 'Madrid', 'El Encinar']];
  sh.getRange(1, 18, 2, 6).setValues(list2);
  sh.getRange(1, 18, 1, 6).setBackground('#BF9000').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#fff');
  sh.getRange(2, 18, 1, 6).setBackground('#FFF2CC').setBorder(true, true, true, true, true, true, '#BF9000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor('#BF9000');

  let cell = sh.getRange('T2');
  let range = sh.getRange('B1:B10');
  let rule = SpreadsheetApp.newDataValidation().requireValueInList(['AEI', 'E', 'I', 'IINT', 'I+D']).build();

  cell.setDataValidation(rule);
  sh.getRange(1, 1, 1, sh.getMaxColumns()).setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  // Template Structure

  sh.getRange('C3').setValue('=CONCATENATE(R2," (",S2,"-",T2,"-",U2,"-",V2,") ",W2)');

  sh.getRange('E4').setValue('=CONCATENATE(R2," 1 Trabajo")');
  sh.getRange('E11').setValue('=CONCATENATE(R2," 2 Doc Previa")');
  sh.getRange('E19').setValue('=CONCATENATE(R2," 3 Comunicación")');
  sh.getRange('E35').setValue('=CONCATENATE($R$2," 4 Proyectos entregados")');
  sh.getRange('E36').setValue('=CONCATENATE($R$2," 5 Publicacion")');
  sh.getRange('E37').setValue('=CONCATENATE($R$2," 6 Obra")');
  sh.getRange('E40').setValue('=CONCATENATE($R$2," 7 Press")');
  sh.getRange('E41').setValue('=CONCATENATE($R$2," 8 Asistencia")');

  let listnv3 = [['=CONCATENATE($R$2," ","Arquitectura")'], ['=CONCATENATE($R$2," ","Breeam")'],
    ['=CONCATENATE($R$2," ","Estructuras")'], ['=CONCATENATE($R$2," ","Instalaciones")'],
    ['=CONCATENATE($R$2," ","Interiorismo")'], ['=CONCATENATE($R$2," ","Mediciones")']
  ];

  sh.getRange(5, 7, 6, 1).setValues(listnv3);

  let listnv3_2 = [['=CONCATENATE($R$2," ","01"," ","Doc recibida cliente")'], ['=CONCATENATE($R$2," ","02"," ","Normativa")'],
    ['=CONCATENATE($R$2," ","03"," ","Web")'], ['=CONCATENATE($R$2," ","04"," ","Cartografía")'],
    ['=CONCATENATE($R$2," ","05"," ","Fotos")'], ['=CONCATENATE($R$2," ","06"," ","Estudio de mercado")'],
    ['=CONCATENATE($R$2," ","07"," ","Doc recibida cliente")']
  ];

  sh.getRange(12, 7, 7, 1).setValues(listnv3_2);

  sh.getRange('G11').setValue('=CONCATENATE($R$2," ","01"," ","Cliente")');
  sh.getRange('G23').setValue('=CONCATENATE($R$2," ","02"," ","Morph Interno")');
  sh.getRange('G26').setValue('=CONCATENATE($R$2," ","Ayuntamiento")');
  sh.getRange('G29').setValue('=CONCATENATE($R$2," ","Renderista")');
  sh.getRange('G32').setValue('=CONCATENATE($R$2," ","Obra")');
  sh.getRange('G38').setValue('01_Fotografías');
  sh.getRange('G39').setValue('02_Actas de obra');

  let listnv4 = [['Enviado'], ['Recibido']];

  // Rows & Columns

  sh.setFrozenRows(1);

  sh.getRange(21, 9, 2, 1).setValues(listnv4);
  sh.getRange(24, 9, 2, 1).setValues(listnv4);
  sh.getRange(27, 9, 2, 1).setValues(listnv4);
  sh.getRange(30, 9, 2, 1).setValues(listnv4);
  sh.getRange(33, 9, 2, 1).setValues(listnv4);

  sh.setColumnWidth(1, 25);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 200);
  sh.setColumnWidth(5, 200);
  sh.setColumnWidth(7, 200);
  sh.setColumnWidth(9, 200);
  sh.setColumnWidth(11, 200);
  sh.setColumnWidth(13, 200);
  sh.setColumnWidth(15, 200);
  sh.setColumnWidth(17, 40);
  sh.setColumnWidth(22, 200);
  sh.setColumnWidth(23, 200);

  removeEmptyColumns();
  deleteEmptyRows();
  SpreadsheetApp.flush();

  sh.activate();
}
