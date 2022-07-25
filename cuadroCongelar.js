function doGet() {
    return HtmlService.createHtmlOutputFromFile("index");
  }

  // CONGELAR CUADRO DE SUPERFICIES

  function congelarCuadro() {

    // BASE

      var ss = SpreadsheetApp.getActive();
      var sh = SpreadsheetApp.getActiveSheet();
      var ws = SpreadsheetApp.getActive().getSheetByName('ACTUALIZAR');
      var spreadsheetId = SpreadsheetApp.getActive().getId();
      var dateGMT = Utilities.formatDate(new Date(), "GMT+1", "yyyyMMdd");
      var file = DriveApp.getFileById(spreadsheetId);
      var parentFolder = file.getParents();
      var parentFolder_ID = parentFolder.next().getId();
      var backupFolderSearch = DriveApp.getFolderById(parentFolder_ID);


    // DESTINATION FOLDER (Específico cuadro de superficies)


      var searchFor ='title contains "Congelados"';
      var names =[];
      var backupFolderIds=[];
      var backupFolder = backupFolderSearch.searchFolders(searchFor);
        while (backupFolder.hasNext()) {
      var backupFolderDef = backupFolder.next();
      var backupFolderId = backupFolderDef.getId();
      backupFolderIds.push(backupFolderId);
      var backupFolderName = backupFolderDef.getName();
      var backupFolderUrl = backupFolderDef.getUrl();
      names.push(backupFolderName);
      }

      ws.getRange('B7')
      .setFontWeight("bold")
      .setFontColor("#B7B7B7")
      ;
      ws.getRange('A7').setValue('CARPETA BACKUP');
      ws.getRange('A8').setValue('ID CARPETA BACKUP');
      ws.getRange('A9').setValue('DESCARGAR ARCHIVO XLSX');
      ws.getRange('B7').setValue('=hyperlink("https://drive.google.com/corp/drive/folders/'+ backupFolderId +'";"' + backupFolderName + '")').setFontColor('#4A86E8');
      //ws.getRange('D7').setValue(backupFolderName);
      ws.getRange('B8').setValue(backupFolderId);


    // Copy each sheet in the source Spreadsheet by removing the formulas as the temporal sheets.

      var tempSheets = ss.getSheets().filter(sh => !sh.isSheetHidden()).map(function(sheet) {
        var dstSheet = sheet.copyTo(ss).setName(sheet.getSheetName() + "_temp");
        var src = dstSheet.getDataRange();
        src.copyTo(src, {contentsOnly: true});
        return dstSheet;
      });


    // Copy the source Spreadsheet.

      var destination = ss.copy(ss.getName() + " - " + dateGMT + " - " + "CONGELADO");
      var destinationId = destination.getId();
      var destinationFile = DriveApp.getFileById(destinationId);

      destinationFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);


    // Delete the temporal sheets in the source Spreadsheet.

      tempSheets.forEach(function(sheet) {
        ss.deleteSheet(sheet);
      });
      SpreadsheetApp.flush();


    // Delete the original sheets from the copied Spreadsheet and rename the copied sheets.

      destination.getSheets().forEach(function(sheet) {
      if (sheet.isSheetHidden()) {
        destination.deleteSheet(sheet);
        SpreadsheetApp.flush();
      }
      else {
        var sheetName = sheet.getSheetName();
        //Logger.log(sheetName);
        if (sheetName.indexOf("_temp") == -1) {
          destination.deleteSheet(sheet);
          SpreadsheetApp.flush();
        }
        else {
          //sheet.setName(sheetName.slice(0, -1));
          sheet.setName(sheetName.replace("_temp", ""));
        }
      }
      });

    // Delete the main DATA SHEET.

      destination.getSheets().forEach(function(sheet) {
      var sheetName = sheet.getSheetName();
      if (sheetName.indexOf("ACTUALIZAR") == -1) {
      }
         else {
          destination.deleteSheet(sheet);
        }
      });


    // Move file to the destination folder.

      file = DriveApp.getFileById(destinationId);
      DriveApp.getFolderById(backupFolderId).addFile(file);
      file.getParents().next().removeFile(file);


    // Export to XLSX.

      try {

      var url = "https://docs.google.com/feeds/download/spreadsheets/Export?key=" + destination.getId() + "&exportFormat=xlsx";

      var params = {
        method      : "get",
        headers     : {"Authorization": "Bearer " + ScriptApp.getOAuthToken()},
        muteHttpExceptions: true
      };

      var blob = UrlFetchApp.fetch(url, params).getBlob();
      blob.setName(ss.getName() + ".xlsx");

      var blobFile = UrlFetchApp.getRequest(url, params);

      ws.getRange('B9').setValue(url).setFontColor('#4A86E8');

      //MailApp.sendEmail('asanchez@morphestudio.es', 'Conversión de Google Sheet a Excel', 'El archivo XLSX aparece adjunto a este correo.', { attachments: [blob] });

    } catch (f) {
      Logger.log(f.toString());
    }


    ws.activate();

  }
