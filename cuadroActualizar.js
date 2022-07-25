function doGet() {
  return HtmlService.createHtmlOutputFromFile("index");
}

// ACTUALIZAR CUADRO DE SUPERFICIES

function actualizarCuadro() {

  var ss = SpreadsheetApp.getActive();
  var sh = SpreadsheetApp.getActiveSheet();
  var ws = ss.getSheetByName('ACTUALIZAR') || ss.insertSheet('ACTUALIZAR', 1);
  var ss_id = ss.getId();
  var file = DriveApp.getFileById(ss_id);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);


  // FORMATO DE CELDAS

    ws.clear().clearFormats();

    // Estilo global
    ws.getRange(1, 1, ws.getMaxRows(), ws.getMaxColumns()).setFontSize(13).setFontFamily("Montserrat").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP)
    .setVerticalAlignment("middle").setFontColor('#B7B7B7');
    // Col A
    ws.getRange(1,1,9,1).setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontWeight("bold");
    // Row 2
    ws.getRange(2,1,1,4).setFontFamily("Inconsolata").setFontSize(16).setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor("#999999").setFontWeight("bold").setHorizontalAlignment("center");
    // Col B
    ws.getRange(3,2,7,1).setBackground("#F3F3F3").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    // ImportRanges
    ws.getRange(1,3,1,2).setBackground("#EAD1DC").setBorder(true, true, true, true, true, true, "#A64D79", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor("#A64D79").setFontWeight   ("bold").setHorizontalAlignment("center");
    // Control Panel
    ws.getRange('B1').setBackground("#D9EAD3").setBorder(true, true, true, true, true, true, "#6AA886", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
    .setFontColor("#6AA886").setFontWeight("bold");
    // Folders Bold
    ws.getRange('B3').setFontWeight("bold");
    ws.getRange('B5').setFontWeight("bold");
    ws.getRange('B7').setFontWeight("bold");
    ws.getRange('B9').setFontWeight("bold");
    // Column/Row Size
    ws.setColumnWidth(1, 330);
    ws.setColumnWidth(2, 380);
    ws.setColumnWidth(3, 285);
    ws.setColumnWidth(4, 400);

    var maxRows = ws.getMaxRows();
    for (var i=1; i < maxRows+1; i++) {
    ws.setRowHeight(i,27);
    }
    ws.setRowHeight(2, 50);


  // CARPETA CUADRO DE SUPERFICIES

    var parents = file.getParents();
    var carpetaBaseID = parents.next().getId();
    var carpetaBase = DriveApp.getFolderById(carpetaBaseID);
    ws.getRange('B2').setValue('Carpetas referentes');
    ws.getRange('A3').setValue('CARPETA CUADRO SUP.');
    ws.getRange('B3').setValue('=hyperlink("https://drive.google.com/corp/drive/folders/'+ carpetaBaseID +'";"' + carpetaBase + '")').setFontColor('#4A86E8');
    ws.getRange('A4').setValue('ID CARPETA CUADRO SUP.');
    ws.getRange('B4').setValue(carpetaBaseID);



  // PANEL DE CONTROL

    var fldrA = [];
    var filA = [];
    while (parents.hasNext()) {
      var f = parents.next();
      fldrA.push(f.getId());
      parents = f.getParents();
    }
    for(let i = 0;i<fldrA.length; i++) {
      let files = DriveApp.getFolderById(fldrA[i]).getFilesByType(MimeType.GOOGLE_SHEETS);
      while(files.hasNext()) {
        let file = files.next();
        if(file.getName().includes("Panel de control")) { // mask
          filA.push([file.getName()],[file.getId()],[file.getUrl()],[fldrA[i]]);
        }
      }
    }

    var [filePanelName, filePanelId, filePanelUrl, folderPanelcId] = filA;

    var panelControl = DriveApp.getFileById(filePanelId);
    var folderPanelcName = DriveApp.getFolderById(folderPanelcId);
    var folderPanelc = panelControl.getParents();

    panelControl.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    ws.getRange('A1').setValue('URL PANEL DE CONTROL');
    ws.getRange('B1').setValue(filePanelUrl);
    ws.getRange('A5').setValue('CARPETA PANEL DE CONTROL');
    ws.getRange('B5').setValue('=hyperlink("https://drive.google.com/corp/drive/folders/'+ folderPanelcId +'";"' + folderPanelcName + '")').setFontColor('#4A86E8');
    ws.getRange('A6').setValue('ID CARPETA PANEL DE CONTROL');
    ws.getRange('B6').setValue(folderPanelcId);
    ws.getRange('A7').setValue('CARPETA BACKUP');
    ws.getRange('A8').setValue('ID CARPETA BACKUP');
    ws.getRange('A9').setValue('DESCARGAR ARCHIVO XLSX');

    /* VERSIÓN OBSOLETA DEL BUSCADOR
    var searchFor ='title contains "Panel"';
    var names =[];
    var filePanelIds=[];
    var filesPanel = folderPanelcName.searchFiles(searchFor);
      while (filesPanel.hasNext()) {
    var filePanel = filesPanel.next();
    var filePanelId = filePanel.getId();// To get FileId of the file
    filePanelIds.push(filePanelId);
    var filePanelname = filePanel.getName();
    var filePanelUrl = filePanel.getUrl();
    names.push(filePanelname);
    }
    */


  // IMPORTRANGE PERMISSION

    var sectoresID = "1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";
    var sectores = DriveApp.getFileById(sectoresID);

    var url = `https://docs.google.com/spreadsheets/d/${ss_id}/externaldata/addimportrangepermissions?donorDocId=${carpetaBaseID}`;
    var tokent1 = ScriptApp.getOAuthToken();
    var paramst1 = {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + tokent1,
      },
      muteHttpExceptions: true
    };

    UrlFetchApp.fetch(url, paramst1);

    var urlt2 = `https://docs.google.com/spreadsheets/d/${ss_id}/externaldata/addimportrangepermissions?donorDocId=${sectoresID}`;
    const tokent2 = ScriptApp.getOAuthToken();
    const paramst2 = {
      method: 'post',
      headers: {
        Authorization: 'Bearer ' + tokent2,
      },
      muteHttpExceptions: true
    };

    UrlFetchApp.fetch(urlt2, paramst2);

    ws.getRange('C1').setValue('=IMPORTRANGE(B1;"Instrucciones!A1")');
    ws.getRange('D1').setValue('=IMPORTRANGE("https://docs.google.com/spreadsheets/d/1CuMcYrtT6NXwxa9fMEIOTgRfkPySnNwKvA_1dyarCro";"DB-SI!B2")');

    Utilities.sleep(250);


  // LOCALIZAR ARCHIVOS TXT EXPORTADOS

    var rangeClear = ws.getRange(3,3,6,2);
    rangeClear.clearContent().clearFormat();

    var searchFor ='title contains "Exportaciones"';
    var names =[];
    var expFolderIds=[];
    var expFolder = carpetaBase.searchFolders(searchFor);
      while (expFolder.hasNext()) {
    var expFolderDef = expFolder.next();
    var expFolderId = expFolderDef.getId();
    expFolderIds.push(expFolderId);
    var expFolderName = expFolderDef.getName();
    var expFolderUrl = expFolderDef.getUrl();
    names.push(expFolderName);
    }

    var sufix = 'TXT'; // mask
    var list = [];
    var files = expFolderDef.getFiles();
    while (files.hasNext()) {
      file = files.next();
      list.push([file.getName(),file.getId(),file.getName().slice(0,-4).replace("Sheets ", "").toUpperCase()]);
    }

    var result = [['Archivos exportados','IDs', 'Hoja'], ...list.filter(r => r[0].includes(sufix)).sort()];
    var resultCrop = result.map(function(val) {
    return val.slice(0, -1);
    });
    ws.getRange(2,3, result.length, 2).setValues(resultCrop);

    // FORMAT FILELIST
    ws.getRange(3,3,list.length,1).setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor("#B7B7B7").setFontWeight("bold");
    ws.getRange(3,4,list.length,1).setBackground("#F3F3F3").setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID_MEDIUM).setFontColor("#B7B7B7");
    ws.getRange(3,3,list.length,2).setFontSize(13).setFontFamily("Montserrat").setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP).setVerticalAlignment("middle");

    Utilities.sleep(100);
    Logger.log(list);

  // COPIA DE DATOS A HOJAS

    for (n in list) {
      var txtFileName = list[n][0];
      var txtFileId = list[n][1];
      var txtFileSheet = list[n][2];

      var txt_file = DriveApp.getFileById(txtFileId);
      txt_file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      var tsvUrl = "https://drive.google.com/uc?id="+txtFileId+"&x=.tsv";
      var tsvContent = UrlFetchApp.fetch(tsvUrl,{muteHttpExceptions: true }).getContentText();
      var tsvData = Utilities.parseCsv(tsvContent, '\t');

      var sheetPaste = ss.getSheetByName(''+txtFileSheet) || ss.insertSheet(''+txtFileSheet, 100);
      sheetPaste.setTabColor("F1C232").hideSheet();
      sheetPaste.clear();
      sheetPaste.getRange(1, 1, tsvData.length, tsvData[0].length).setValues(tsvData);
    }

    ws.activate();

}
