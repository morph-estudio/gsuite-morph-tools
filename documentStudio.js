function doGet() {
  return HtmlService.createHtmlOutputFromFile('documentStudioIndex');
}

/*
 * Gsuite Morph Tools - Morph Document Studio 1.9
 * Developed by alsanchezromero
 * Created on Mon Jul 25 2022
 *
 * Copyright (c) 2022 Morph Estudio
*/

// GET-MARKERS FUNCS

function getDocItems(docID, identifier){
  let body = DocumentApp.openById(docID).getBody();
  let docText = body.getText();

  // Check if search characters are to be included.
  let startLen = identifier.start_include ? 0 : identifier.start.length;
  let endLen = identifier.end_include ? 0 : identifier.end.length;

  // Set up the reference loop
  let textStart = 0;
  let doc = docText;
  let docList = [];

  // Loop through text grab the identifier items. Start loop from last set of end identfiers.
  while (textStart > -1){
    textStart = doc.indexOf(identifier.start);

    if (textStart === -1){
      break;
    } else {
      let textEnd = doc.indexOf(identifier.end) + identifier.end.length;
      let word = doc.substring(textStart, textEnd);

      doc = doc.substring(textEnd);
      docList.push(word.substring(startLen, word.length - endLen));
    }
  }

  // Return a unique set of identifiers.
  return [...new Set(docList)];
}

function getSlidesItems(docID, identifier){
  var slides = SlidesApp.openById(docID).getSlides();

  var sumaTextos = [];
  slides.forEach(function(slide){
    var shapes = (slide.getShapes());
    shapes.forEach(function(shape){

      var textito = shape.getText().asString();
      sumaTextos.push(textito);
    });
  });

  var docText = sumaTextos.toString();

  // Check if search characters are to be included.
  var startLen =  identifier.start_include ? 0 : identifier.start.length;
  var endLen = identifier.end_include ? 0 : identifier.end.length;

  // Set up the reference loop
  var textStart = 0;
  var doc = docText;
  var docList = [];

  // Loop through text grab the identifier items. Start loop from last set of end identfiers.
  while(textStart > -1){
    textStart = doc.indexOf(identifier.start);

    if(textStart === -1){
      break;
    } else{
      var textEnd = doc.indexOf(identifier.end) + identifier.end.length;
      var word = doc.substring(textStart,textEnd);

      doc = doc.substring(textEnd);
      docList.push(word.substring(startLen,word.length - endLen));
    }
  }

  // Return a unique set of identifiers.
  return [...new Set(docList)];
}

function isGreenCell (lastCell){
  var mycell = SpreadsheetApp.getActiveSheet().getRange(1,lastCell);
  var bghex = mycell.getBackground();
  if (bghex == "#ecfdf5"){
    return true;
  } else{
    return false;
  }
}

function columnRemover(sh, updatedValues, numberOfGreenCells){

  var lastColHeaders = sh.getLastColumn() - numberOfGreenCells;
  var headerRange = sh.getRange(1, 1, 1, lastColHeaders);
  var headerValues = headerRange.getValues()[0];

  var deleteColumn = [];
  headerValues.forEach(function(a, index) {
    var i = updatedValues.indexOf(a);
    if (i === -1) {
      deleteColumn.push(index+1);
    }
  });

  for (var j=lastColHeaders; j > 0; j--) {
    if (deleteColumn.indexOf(j) == -1) {
    }
    else{
      var lastCol = sh.getLastColumn();
      if (lastCol > 1){
        sh.deleteColumn(j);
      } else {
        sh.insertColumnAfter(lastCol);
        sh.deleteColumn(j);
      }

    }
  }
}


// DOCUMENTSTUDIO FUNCTIONS

function isValidHttpUrl(str) {
  const pattern = new RegExp('^(https?:\\/\\/)?'+ // protocol
    '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|'+ // domain name
    '((\\d{1,3}\\.){3}\\d{1,3}))'+ // OR ip (v4) address
    '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*'+ // port and path
    '(\\?[;&a-z\\d%_.~+=-]*)?'+ // query string
    '(\\#[-a-z\\d_]*)?$','i'); // fragment locator
  return !!pattern.test(str);
 }

function isImage(url) {
  return /\.(jpg|jpeg|png|webp|avif|gif|svg)$/.test(url);
}

function imageFromTextDocs(body, searchText, image) {
    var next = body.findText(searchText);
    var atts=body.getAttributes();
    if (!next) return;
    var r = next.getElement();
    r.asText().setText("");
    var img = r.getParent().asParagraph().insertInlineImage(0, image);

      var w = img.getWidth();
      var h = img.getHeight();
      // el ratio es por h/w o entre w/h
      var mr = atts['MARGIN_RIGHT'];
      var ml = atts['MARGIN_LEFT'];
      var sw = atts['PAGE_WIDTH'];
      var sh = atts['PAGE_HEIGHT'];

      if (w > sw){
      img.setWidth(sw);
      img.setHeight(sw * h / w);
      } else {
      }

    return next;
}

function imageFromTextDocsCustomWidth(body, searchText, image, width) {
    var next = body.findText(searchText);
    var atts=body.getAttributes();
    if (!next) return;
    var r = next.getElement();
    r.asText().setText("");
    var img = r.getParent().asParagraph().insertInlineImage(0, image);
    var w = img.getWidth();
    var h = img.getHeight();
    var sw = atts['PAGE_WIDTH'];

    if (width > sw){
      img.setWidth(sw);
      img.setHeight(sw * h / w);
    } else {
      img.setWidth(width);
      img.setHeight(width * h / w);
    }

    return next;

}

function replaceDocText(repText, newText, copyId) {

  var replaceRules = [
    {
      toReplace: repText,
      newValue: newText
    },];
  const requestBuild = replaceRules.map(rule => {
    var replaceAllTextRequest = Docs.newReplaceAllTextRequest();
    replaceAllTextRequest.replaceText = rule.newValue;
    replaceAllTextRequest.containsText = Docs.newSubstringMatchCriteria();
    replaceAllTextRequest.containsText.text = rule.toReplace;
    replaceAllTextRequest.containsText.matchCase = false;
    //Logger.log(replaceAllTextRequest)
    var request = Docs.newRequest();
    request.replaceAllText = replaceAllTextRequest;
    return request;
  });

    var batchUpdateRequest = Docs.newBatchUpdateDocumentRequest();
    batchUpdateRequest.requests = requestBuild;
    var result = Docs.Documents.batchUpdate(batchUpdateRequest, copyId);

}

function imageFromTextSlides (slides, searchText, imageUrl) {

  slides.forEach(function(slide){
    slide.getShapes().forEach(s => {
      if (s.getText().asString().includes(searchText)) {
      s.replaceWithImage(imageUrl, true);
      }
    });
  });

}

function replaceSlideText (slides, replaceText, markerText, copyId) {

  slides.forEach(function(slide){
    var pageElementId = slide.getObjectId();
    var resource = {
      requests: [{
        replaceAllText: {
          pageObjectIds: [pageElementId],
          replaceText: replaceText,
          containsText: { matchCase: false, text: markerText }
        }
      }]
    };
    var result = Slides.Presentations.batchUpdate(resource, copyId);
  });

}

function formatDropdown(){
  var ss = SpreadsheetApp.getActive();
  var sh = SpreadsheetApp.getActiveSheet();
  var ws = ss.getSheetByName('DS OPTIONS') || ss.insertSheet('DS OPTIONS', 1);

  return ws.getRange(2,1,ws.getLastRow()-1,1).getValues();
}

// GET MARKERS MAIN FUNCTION

function getMarkers(rowData){
  // var ss = SpreadsheetApp.getActive();
  const sh = SpreadsheetApp.getActiveSheet();
  let lastColmn;

  /* TEST
    //var ss = SpreadsheetApp.openById('1v5f3X1ShmVCGdT6NdWvmJHcfeP01ptuwHfT1iqM6UQI');
    //var sh = ss.getSheetByName('EXEC')
    //var docID = "1Gs4Gd4JtVMrI-nu6ELZtS71kebUU9CN3II2Xz5-Q2F8";
  */

  // DATA + VARIABLES

  const formData = [rowData.templateID, rowData.greenCells1, rowData.purgeMarkers];
  let [docURL, greenCells1, purgeMarkers] = formData;

  let urlCell = sh.getLastColumn();
  let nameCell = urlCell - 1;
  let docID;

  if (greenCells1) {
    docURL = sh.getRange(1, nameCell).getValue();
    docID = getIdFromUrl(docURL);
  } else{
    docID = getIdFromUrl(docURL);
  }

  var identifier = {
    start: `{{`,
    start_include: true,
    end: `}}`,
    end_include: true
  };

  var gDocTemplate = DriveApp.getFileById(docID);
  var fileType = gDocTemplate.getMimeType();
  var docMarkers;

  switch (fileType) {
  case MimeType.GOOGLE_DOCS:
    docMarkers = getDocItems(docID, identifier);
  break;
  case MimeType.GOOGLE_SLIDES:
    docMarkers = getSlidesItems(docID, identifier);
  break;
  default:
  }

  var updatedValues = [];
  var driverArray = docMarkers.flat(); // Slicing {} markers for enhancing interface.
  driverArray.forEach(function(el){
  var sliced = el.slice(2,-2);
  updatedValues.push(sliced);
  });

  // DELETE REMOVED MARKERS

  var isGrCell;
  var semiLastCell;
  var semiIsGrCell;

  if (purgeMarkers == true){
    lastColmn = sh.getLastColumn();
    isGrCell = isGreenCell(lastColmn);
    semiLastCell = lastColmn - 1;
    semiIsGrCell = isGreenCell(semiLastCell);

    if (isGrCell == true && semiIsGrCell == true){
      columnRemover(sh, updatedValues, 2);
    }
    else if (isGrCell == false && semiIsGrCell == false) {
      columnRemover(sh, updatedValues, 0);
    } else if (isGrCell == true && semiIsGrCell == false) {
      columnRemover(sh, updatedValues, 1);
    }
  }

  // ADD NEW MARKERS

  lastColmn = sh.getLastColumn();
  if (lastColmn >= 1){
    var newHeaderRange = sh.getRange(1, 1, 1, sh.getLastColumn());
    var headerValuesNew = newHeaderRange.getValues()[0];

    updatedValues.forEach(function(a, index) {

      var i = headerValuesNew.indexOf(a);
      if (i === -1) {

        if (index === 0){
          sh.insertColumns(index+1);
        } else{
        sh.insertColumnAfter(index);
        }

        sh.setColumnWidth(index+1, 150);
        var headerCell = sh.getRange(1,index+1,1,1);
        headerCell.setValue(a);

      }
    });
  } else {

    updatedValues.forEach(function(a, index) {

      sh.insertColumns(index+1);
      sh.setColumnWidth(index+1, 150);
      var firstCell= sh.getRange(1,index+1, 1, 1);
      firstCell.setValue(a);

    });

  }

  // STYLE FORMAT

  lastColmn = sh.getLastColumn();
  isGrCell = isGreenCell(lastColmn);
  semiLastCell = lastColmn - 1;
  semiIsGrCell = isGreenCell(semiLastCell);

  sh.getRange(1, 1, 1, sh.getMaxColumns()).setFontWeight("bold");
  sh.setFrozenRows(1);

  if (lastColmn === 0){
    sh.insertColumns(1);
    sh.getRange(1, 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] File-links');
    sh.insertColumns(1);
    sh.getRange(1, 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] Files')
    .setNote('Celdas verdes: para utilizar la opción "usar celdas verdes" debes introducir en esta celda el LINK de la plantilla y en la siguiente columna el LINK de la carpeta de destino.');
  } else if (isGrCell == false){
    sh.insertColumnAfter(lastColmn);
    sh.getRange(1, lastColmn+1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] File-links');
    sh.insertColumnAfter(lastColmn);
    sh.getRange(1, lastColmn+1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] Files')
    .setNote('Celdas verdes: para utilizar la opción "usar celdas verdes" debes introducir en esta celda el LINK de la plantilla y en la siguiente columna el LINK de la carpeta de destino.');
  } else if (isGrCell == true && semiIsGrCell == false){
    sh.insertColumnAfter(lastColmn);
    sh.getRange(1, lastColmn+1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS]');
  }

/*
  var docMarkersLenght = updatedValues.length;

  if (lastColmn === 0){
    sh.insertColumns(1)
    sh.getRange(1, 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] File-links')
    sh.insertColumns(1)
    sh.getRange(1, 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] Files')
    .setNote('Celdas verdes: para utilizar la opción "usar celdas verdes" debes introducir en esta celda el LINK de la plantilla y en la siguiente columna el LINK de la carpeta de destino.')
  } else if (lastColmn === docMarkersLenght){
    sh.insertColumnAfter(lastColmn)
    sh.getRange(1, lastColmn+1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] File-links')
    sh.insertColumnAfter(lastColmn)
    sh.getRange(1, lastColmn+1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS] Files')
    .setNote('Celdas verdes: para utilizar la opción "usar celdas verdes" debes introducir en esta celda el LINK de la plantilla y en la siguiente columna el LINK de la carpeta de destino.')
  } else if (lastColmn === docMarkersLenght + 1){
    sh.insertColumnAfter(lastColmn)
    sh.getRange(1, lastColmn+1).setBackground('#ECFDF5').setFontColor('#34a853').setValue('[DS]')
  };
*/

  lastColmn = sh.getLastColumn();
  var lastCol2 = sh.getLastColumn()-1;
  sh.setColumnWidth(lastColmn, 300); sh.setColumnWidth(lastCol2, 300);

  removeEmptyColumns();

}


function documentStudio(rowData) {

  // var ss = SpreadsheetApp.getActive();
  var sh = SpreadsheetApp.getActiveSheet();

  var formData = [
    rowData.templateID,
    rowData.destFolder,
    rowData.fileName,
    rowData.exportFormat,
    rowData.permission1,
    rowData.permission2,
    rowData.permission3,
    rowData.numSwitch,
    rowData.greenCells
    ];

  var [docURL,destFolderURL,fileName,exportFormat,permission1,permission2,permission3,numSwitch,greenCells] = formData;

  var rows = sh.getDataRange().getValues();
  var lastColmn = sh.getLastColumn();
  var urlCell = lastColmn;
  var nameCell = urlCell - 1;
  var fin = nameCell - 1;

  var userMail = Session.getActiveUser().getEmail();
  var dateNow = Utilities.formatDate(new Date(), "GMT+2", "dd/MM/yyyy - HH:mm:ss");
  var destFolderID; var docID;

  if (greenCells){
    docURL = sh.getRange(1, nameCell).getValue();
    docID = getIdFromUrl(docURL);
    destFolderURL = sh.getRange(1, urlCell).getValue();
    destFolderID = getIdFromUrl(destFolderURL);
  } else{
    docID = getIdFromUrl(docURL);
    destFolderID = getIdFromUrl(destFolderURL);
  }

  // Create temporal folder if destFolder is empty
  var destinationFolder;
  if (destFolderID === ''){
    var rootFolder = DriveApp.getRootFolder();
    destinationFolder = rootFolder.createFolder('Morph Tools Files');
  } else {
    destinationFolder = DriveApp.getFolderById(destFolderID);
  }

  // INTERNALLY GETTING MARKERS

  var identifier = {
    start: `{{`,
    start_include: true,
    end: `}}`,
    end_include: true
  };

  var gDocTemplate = DriveApp.getFileById(docID);
  var fileType = gDocTemplate.getMimeType();
  var docMarkers;

  switch (fileType) {
  case MimeType.GOOGLE_DOCS:

    docMarkers = getDocItems(docID, identifier).sort();

  break;
  case MimeType.GOOGLE_SLIDES:

    docMarkers = getSlidesItems(docID, identifier).sort();

  break;
  default:
  }

  var rw1 = [docMarkers];
  var rw2 = rw1[0].map((col, i) => rw1.map(row => row[i]));

  var doc; var newNames; var headerValues; var value;

  // BODY ITERATION

  rows.forEach(function(row, index){
    if (index === 0) return; // Check if this row is the headers, if so we skip it
    if (row[sh.getLastColumn()-1]) return; // Check if a document has already been generated by looking at 'Document Link', if so we skip it
    var copy = gDocTemplate.makeCopy(fileName, destinationFolder); // Copy of Template (`DS ${row[0]}, ${row[1]}` , destinationFolder)
    var copyId = copy.getId();

    var doc; var headerValues;
    switch (fileType) {
    case MimeType.GOOGLE_DOCS:

    doc = DocumentApp.openById(copy.getId()); // Open Copy using DocumentApp

    for (var i=0; i<lastColmn; i=i+1){
      headerValues = sh.getRange(1,i+1).getValue();
      newNames = doc.getName().replace(`${'{{'+headerValues+'}}'}`, row[i]);
      doc.setName(newNames);
    }

    var body = doc.getBody(); // Get content of doc

    // TEXT REPLACING

    var customWidth; var checkURL; var checkIMG; var imageID; var imageFile; var imageType; var image; var response;

    for (var i=0; i<fin; i=i+1){

      headerValues = sh.getRange(1,i+1).getValue();
      value = sh.getRange(index+1,i+1).getValue().toString();
      var output = [];

      if (value.includes("{w=")){
        var t = value.split('{w=');
        t.forEach(function(q){
          output.push([q]);
        });
        value = output[0].toString();
        width = parseInt(output[1].toString().replace("}",""), 10).toFixed(0);
        customWidth = true;
      } else {customWidth = false;}

      var replaceText = `${'{{'+headerValues+'}}'}`;
      checkURL = isValidHttpUrl(`${value}`);
      checkIMG = isImage(`${value}`);

      if (checkURL == true){

        if(value.indexOf("drive.google.com/file")>-1){
          imageID = getIdFromUrl(value);
          imageFile = DriveApp.getFileById(imageID);
          imageType = imageFile.getMimeType();

          if(imageType == 'JPEG','PNG','GIF'){
            imageFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            image = DriveApp.getFileById(imageID).getBlob();

            if (customWidth == false){
              imageFromTextDocs(body, replaceText, image);
            } else {
              imageFromTextDocsCustomWidth(body, replaceText, image, width);
            }

          }
        } else if (checkIMG == true) {
            response = UrlFetchApp.fetch(`${''+value}`);
            image = response.getBlob();

            if (customWidth == false){
              imageFromTextDocs(body, replaceText, image);
            } else {
              imageFromTextDocsCustomWidth(body, replaceText, image, width);
            }

        } else {
            replaceDocText(`${'{{'+headerValues+'}}'}`, `${row[i]}`, copyId);
        }
      } else {
        replaceDocText(`${'{{'+headerValues+'}}'}`, `${row[i]}`, copyId);
      }

    }

    break;
    case MimeType.GOOGLE_SLIDES:

    doc = SlidesApp.openById(copy.getId());

    for (var i=0; i<lastColmn; i=i+1){
      headerValues = sh.getRange(1,i+1).getValue();
      newNames = doc.getName().replace(`${'{{'+headerValues+'}}'}`, row[i]);
      doc.setName(newNames);
    }

    var slides = doc.getSlides();

    // TEXT REPLACING

    for (var i=0; i<fin; i=i+1){
      headerValues = sh.getRange(1,i+1).getValue();
      var searchText = `${'{{'+headerValues+'}}'}`;
      value = sh.getRange(index+1,i+1).getValue().toString();
      checkURL = isValidHttpUrl(`${value}`);
      checkIMG = isImage(`${value}`);

      if (checkURL == true){

        if(value.indexOf("drive.google.com/file")>-1){
          imageID = getIdFromUrl(value);
          imageFile = DriveApp.getFileById(imageID);
          imageType = imageFile.getMimeType();

          if(imageType == 'JPEG','PNG','GIF'){
            imageFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            image = DriveApp.getFileById(imageID).getBlob();
            imageFromTextSlides(slides, searchText, image);
          }
        } else if (checkIMG == true) {
            //var imageUrl = `${''+value}`;
            response = UrlFetchApp.fetch(`${''+value}`);
            image = response.getBlob();
            imageFromTextSlides(slides, searchText, image);
        } else {
            replaceSlideText(slides, `${row[i]}`, `${'{{'+headerValues+'}}'}`, copyId);
        }

      } else {
          replaceSlideText(slides, `${row[i]}`, `${'{{'+headerValues+'}}'}`, copyId);
      }

    }

    break;
    default:
    }

    // EXPORT OPTIONS

    var url; var newDocID;
    switch (exportFormat) {
      case 'PDF':

        doc.saveAndClose();
        newDocID = doc.getId();
        var templateFile = DriveApp.getFileById(newDocID);
        var theBlob = templateFile.getBlob().getAs('application/pdf');
        doc = destinationFolder.createFile(theBlob);
        copy.setTrashed(true); // Delete the original file

      break;
      case 'KEEP':

        url = doc.getUrl();
        newDocID = doc.getId();
        doc.saveAndClose();
        doc = DriveApp.getFileById(newDocID);

      break;
      default:
    }

    // AUTONUMERATION AND RESULTS

    if (numSwitch){
      newNames = newNames + "_" + String(index).padStart(3, '0');
      doc.setName(newNames);
    }

    url = doc.getUrl();
    sh.getRange(index + 1, nameCell).setValue('=hyperlink("'+url+'","'+newNames+'")'); // File-name link cell
    sh.getRange(index + 1, urlCell).setValue(url).setNote(null).setNote('Actualizado por ' + userMail + ' el ' + dateNow); // URL link cell

    // FILE PERMISSION

    if (permission2){
      doc.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
    }
    if (permission3){
      doc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

  });

}





