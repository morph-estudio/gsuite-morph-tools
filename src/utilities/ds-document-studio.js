/**
 * Gsuite Morph Tools - Morph Document Studio
 * Developed by alsanchezromero
 *
 * Morph Estudio, 2023
 */

function documentStudio(rowData) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  var startTime = new Date().getTime();
  var elapsedTime;

  const {
    dsActivate,
    emailActivate,
    templateID,
    greenCells,
    destinationFolder,
    fileName,
    exportFormat,
    permission1,
    permission2,
    permission3,
    numerationSwitch,
    allDocuments
  } = rowData;

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before Basic Data, Files and Folders: ${elapsedTime} seconds.`);

  // Basic Data, Files and Folders

  let userMail = Session.getActiveUser().getEmail();
  let dateNow = Utilities.formatDate(new Date(), 'GMT+2', 'dd/MM/yyyy - HH:mm:ss');

  const { indexNameCell, indexUrlCell } = getGreenColumns(sh);

  Logger.log(`El index de File-links es: ${indexUrlCell}`);

  if (dsActivate) {

    var doc;
    var headerValues; 
    var value;
    var columnaFechaArray = [];

    var {docID, destinationFolderFinale} = getDocumentData(greenCells, sh, templateID, destinationFolder, indexNameCell, indexUrlCell);
    var mainTemplateFile = DriveApp.getFileById(docID);
    var fileType = mainTemplateFile.getMimeType();

    var lastColmn = sh.getLastColumn();
    var rows = sh.getDataRange().getValues();

    elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before Body Iteration: ${elapsedTime} seconds. rows.length: ${rows.length}`);

    // Body Iteration

    for (var index = 1; index < rows.length; index++) {

      SpreadsheetApp.flush();
      var requests = [];
      var row = rows[index];
      
      if (row[indexUrlCell] !== "" && allDocuments === false) {
        Logger.log(`ATENCIÓN: No se generará el documento correspondiente a la fila ${index}`)
        continue;
      } else {
        Logger.log(`Se generará el documento correspondiente a la fila ${index}`)
      }

      var copy = mainTemplateFile.makeCopy(fileName, destinationFolderFinale); 
      var copyId = copy.getId();

      let headerRow = rows[0];
      var docNameWordArray = {};

      switch (fileType) {
        case MimeType.GOOGLE_DOCS:
          doc = DocumentApp.openById(copyId);
          var body = doc.getBody();

          var customWidth;
          var checkURL;
          var checkIMG;
          var imageID;
          var imageFile;
          var imageType;
          var image;
          var response;
          var width;

          for (var i = 0; i < indexNameCell - 1; i++) {
            var headerValues = rows[0][i].toString();
            if (headerValues === "" || headerValues === null || headerValues.startsWith('[')) {
                continue;
            }

            value = rows[index][i].toString();

            var imgCustomSizeCheck = parseImageDimension(value);

            if (imgCustomSizeCheck.customDimension) {
              value = imgCustomSizeCheck.value
            }

            let replaceText = `${`{{${headerValues}}}`}`;
            if (fileName.includes(replaceText)) { docNameWordArray[replaceText] = value; }
            checkURL = isValidHttpUrl(`${value}`);
            checkIMG = isImage(`${value}`);

            if (checkURL == true) {
              if (value.indexOf('drive.google.com/file') > -1) {
                imageID = getIdFromUrl(value);
                imageFile = DriveApp.getFileById(imageID);
                imageType = imageFile.getMimeType();

                if (imageType == 'JPEG', 'PNG', 'GIF') {
                  Logger.log(`Es una imagen con la ID ${imageID}`)
                  //imageFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
                  image = DriveApp.getFileById(imageID).getBlob();

                  imgCustomSizeCheck.customDimension ? imageFromTextDocsCustomDim(body, replaceText, image, imgCustomSizeCheck.dimension, imgCustomSizeCheck.number) : imageFromTextDocs(body, replaceText, image);

                }
              } else if (checkIMG == true) {
                response = UrlFetchApp.fetch(`${`${value}`}`);
                image = response.getBlob();

                if (customWidth == false) {
                  //Logger.log(`Entraste a la imagen externa ${value}`)
                  imageFromTextDocs(body, replaceText, image);
                } else {
                  //Logger.log(`Entraste a la imagen externa custom ${value}`)
                  imageFromTextDocsCustomDim(body, replaceText, image, width);
                }
              } else {
                createDocsTextRequests(replaceText, value, requests)
              }
            } else {
              createDocsTextRequests(replaceText, value, requests)
            }
          }

          if (requests.length > 0) {
            let batchUpdateRequest = Docs.newBatchUpdateDocumentRequest();
            batchUpdateRequest.requests = requests;
            Docs.Documents.batchUpdate(batchUpdateRequest, copyId);
          }

        break;
        case MimeType.GOOGLE_SLIDES:
          elapsedTime = (new Date().getTime() - startTime) / 1000; 
          // Logger.log(`Elapsed time before Enter Slides: ${elapsedTime} seconds.`);
          doc = SlidesApp.openById(copyId);
          
          let slides = doc.getSlides();

          requests = []; // Move requests array here to accumulate all requests before executing.

          for (let i = 0; i < lastColmn; i += 1) {
            headerValues = headerRow[i];

            if(columnaFechaArray.length < 1 && headerValues.toString().toLowerCase().includes('fecha')) { columnaFechaArray.push(i)}

            let searchText = `{{${headerValues}}}`;
            value = row[i].toString();
            if(columnaFechaArray.includes(i)) { value = Utilities.formatDate(new Date(value), "GMT+0100", "dd/MM/yyyy");}
            if (fileName.includes(searchText)) { docNameWordArray[searchText] = value; }
            checkURL = isValidHttpUrl(value);
            checkIMG = isImage(value);

            if (checkURL) {
              if (value.indexOf("drive.google.com/file") > -1) {
                image = fetchImage(value);
              } else if (checkIMG) {
                response = UrlFetchApp.fetch(value);
                image = response.getBlob();
              }
              if (image) {
                imageFromTextSlides(slides, searchText, image);
              } else {
                //elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before replaceSlideText: ${elapsedTime} seconds.`);
                createSlideTextRequests(slides, value, searchText, requests);
              }
            } else {
              //elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before replaceSlideText: ${elapsedTime} seconds.`);
              createSlideTextRequests(slides, value, searchText, requests);
            }
          }

          // Execute all requests in one batch after the loop.
          Slides.Presentations.batchUpdate({ requests: requests }, copyId);
        break;
        default:
      }

      elapsedTime = (new Date().getTime() - startTime) / 1000;
      // Logger.log(`Elapsed time before ROW LOOP ${index}: ${elapsedTime} seconds.`);

      /** Export Options **/
      
      var doc = exportDocumentStudio(doc, exportFormat, destinationFolderFinale, startTime);

      /** New File Naming **/

      // Logger.log(`docNameWordArray: ${JSON.stringify(docNameWordArray)}`);

      var finalFileName = fileName;

      for (var key in docNameWordArray) {
        if (docNameWordArray.hasOwnProperty(key)) {
          var regex = new RegExp(key, 'g'); // Expresión regular global para reemplazar todas las ocurrencias.
          finalFileName = finalFileName.replace(regex, docNameWordArray[key]);
        }
      }

      if (numerationSwitch) {
        finalFileName = `${finalFileName}_${String(index).padStart(3, '0')}`;
      }

      doc.setName(finalFileName);

      var url = doc.getUrl();
      
      /** Update URL Link Cell and Filename Link Cell **/
      let range = sh.getRange(index + 1, indexNameCell + 1, 1, 2);
        range.setValues([[
          `=hyperlink("${url}";"${doc.getName()}")`, 
          url
        ]]).setNotes([
          [null, `Actualizado por ${userMail} el ${dateNow}`]
        ]);

      /** File Permission **/

      if (permission2) { doc.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW); }
      if (permission3) { doc.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }

      //elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before replaceSlideText: ${elapsedTime} seconds.`);
    };
  }

  /** Email Sending **/

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time before Email Sending: ${elapsedTime} seconds.`);

  if (emailActivate) {
    emailSending(sh, rowData, indexUrlCell, userMail, dateNow);
  }

  elapsedTime = (new Date().getTime() - startTime) / 1000; Logger.log(`Elapsed time after Email Sending: ${elapsedTime} seconds.`);

}

/**
 * exportDocumentStudio
 * Export Document Studio File to different format
 */
function exportDocumentStudio(doc, exportFormat, destinationFolderFinale, startTime) {

  let newDocID;
  var exportedDoc;

  switch (exportFormat) {
    case 'PDF':
      doc.saveAndClose();
      newDocID = doc.getId();
      doc = DriveApp.getFileById(newDocID);
      elapsedTime = (new Date().getTime() - startTime) / 1000; //Logger.log(`Elapsed time before theBlob: ${elapsedTime} seconds.`);
      let theBlob = doc.getBlob().getAs('application/pdf');
      exportedDoc = destinationFolderFinale.createFile(theBlob);
      elapsedTime = (new Date().getTime() - startTime) / 1000; //Logger.log(`Elapsed time before setTrashed: ${elapsedTime} seconds.`);
      doc.setTrashed(true); // Delete the original file
      elapsedTime = (new Date().getTime() - startTime) / 1000; //Logger.log(`Elapsed time after setTrashed: ${elapsedTime} seconds.`);
      break;
    case 'KEEP':
      doc.saveAndClose();
      newDocID = doc.getId();
      exportedDoc = DriveApp.getFileById(newDocID);
      break;
    default:
  }

  return exportedDoc ;
}

////////////////////////////// FIND AND REPLACE FUNCTIONS

/**
 * replaceDocText
 * Find and Replace Markers in a Google Doc File
 */
function replaceDocText(repText, newText, copyId) {
  let replaceRules = [
    {
      toReplace: repText,
      newValue: newText,
    }];
  const requestBuild = replaceRules.map(rule => {
    let replaceAllTextRequest = Docs.newReplaceAllTextRequest();
    replaceAllTextRequest.replaceText = rule.newValue;
    replaceAllTextRequest.containsText = Docs.newSubstringMatchCriteria();
    replaceAllTextRequest.containsText.text = rule.toReplace;
    replaceAllTextRequest.containsText.matchCase = false;
    let request = Docs.newRequest();
    request.replaceAllText = replaceAllTextRequest;
    return request;
  });

  let batchUpdateRequest = Docs.newBatchUpdateDocumentRequest();
  batchUpdateRequest.requests = requestBuild;
  Docs.Documents.batchUpdate(batchUpdateRequest, copyId);
}

/**
 * imageFromTextDocs
 * Find and Replace Image Markers in a Google Doc File
 */
function imageFromTextDocs(body, searchText, image) {
  let next = body.findText(searchText);
  let atts = body.getAttributes();
  if (!next) return;
  let r = next.getElement();
  r.asText().setText('');
  let img = r.getParent().asParagraph().insertInlineImage(0, image);

  let w = img.getWidth();
  let h = img.getHeight();
  // el ratio es por h/w o entre w/h
  let mr = atts['MARGIN_RIGHT'];
  let ml = atts['MARGIN_LEFT'];
  let sw = atts['PAGE_WIDTH'];
  let sh = atts['PAGE_HEIGHT'];

  if (w > sw) {
    img.setWidth(sw);
    img.setHeight((sw * h) / w);
  }
}

/**
 * imageFromTextDocsCustomDim
 * Find and Replace Image Markers with Custom Size in a Google Doc File
 */
function imageFromTextDocsCustomDim(body, searchText, image, dimension, size) {
  let next = body.findText(searchText);
  let atts = body.getAttributes();
  if (!next) return;
  let r = next.getElement();
  r.asText().setText('');
  let img = r.getParent().asParagraph().insertInlineImage(0, image);
  let w = img.getWidth();
  let h = img.getHeight();
  let sw = atts['PAGE_WIDTH'];
  let sh = atts['PAGE_HEIGHT'];

  if (dimension === "h") {
    if (size > sh) {
      img.setHeight(size);
      img.setWidth((size * w) / h);
    } else {
      img.setHeight(size);
      img.setWidth((size * w) / h);
    }
  } else if (dimension === "w") {
    if (size > sw) {
      img.setWidth(sw);
      img.setHeight((sw * h) / w);
    } else {
      img.setWidth(size);
      img.setHeight((size * h) / w);
    }
  }
}

/*
  if (width > sw) {
    img.setWidth(sw);
    img.setHeight((sw * h) / w);
  } else {
    img.setWidth(width);
    img.setHeight((width * h) / w);
  }
*/
/**
 * createSlideTextRequests
 * Find and Replace Markers in a Google Slides File
 */
function createDocsTextRequests(replacementText, value, requests) {
  requests.push({
    replaceAllText: {
      containsText: {
        text: replacementText,
        matchCase: false,
      },
      replaceText: value,
    },
  })
}

/**
 * createSlideTextRequests
 * Find and Replace Markers in a Google Slides File
 */
function createSlideTextRequests(slides, replaceText, markerText, requests) {
  for(let slide of slides) {
    requests.push({
      replaceAllText: {
        pageObjectIds: [slide.getObjectId()],
        replaceText: replaceText,
        containsText: { matchCase: false, text: markerText },
      },
    });
  }
}

/**
 * imageFromTextSlides
 * Find and Replace Image Markers in a Google Slides File
 */
function imageFromTextSlides(slides, searchText, imageUrl) {
  for (var i = 0; i < slides.length; i++) {
    var slide = slides[i];
    var shapes = slide.getShapes();
    for (var j = 0; j < shapes.length; j++) {
      var shape = shapes[j];
      if (shape.getText().asString().includes(searchText)) {
        shape.replaceWithImage(imageUrl, true);
      }
    }
  }
}
























function fetchImage(url, customWidth = false) {
  let imageID = getIdFromUrl(url);
  let imageFile = DriveApp.getFileById(imageID);
  let imageType = imageFile.getMimeType();

  if (imageType == 'JPEG', 'PNG', 'GIF') {
    imageFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return DriveApp.getFileById(imageID).getBlob();
  }
}

function getWidth(value) {
  if (value.includes('{w=')) {
    let output = [];
    let t = value.split('{w=');
    t.forEach((q) => {
      output.push([q]);
    });
    return parseInt(output[1].toString().replace('}', ''), 10).toFixed(0);
  }
  return null;
}








/**
 * isGreenCell
 * Check if a column is a Document Studio Column
 */
function isGreenCell(lastCell) {
  let mycell = SpreadsheetApp.getActiveSheet().getRange(1, lastCell);
  let bgHEX = mycell.getBackground();
  if (bgHEX == '#ecfdf5') {
    return true;
  }
  return false;
}

/**
 * addGreenColumn
 * Genera las columnas propias de Document Studio
 */
function addGreenColumn(sh, headerTitle, cellNote) {
  let lastColmn = sh.getLastColumn();

  if (lastColmn === 0) {
    sh.insertColumns(1);
    sh.getRange(1, lastColmn + 1).setBackground('#ECFDF5').setFontColor('#00C853').setValue(headerTitle)
      .setNote(cellNote);
  } else {
    sh.insertColumnAfter(lastColmn);
    sh.getRange(1, lastColmn + 1).setBackground('#ECFDF5').setFontColor('#00C853').setValue(headerTitle)
      .setNote(cellNote);
  }
}

/**
 * getGreenColumns
 * Genera las columnas propias de Document Studio
 */
function getGreenColumns(sh) {

  var filenameField = '[DS] Files'; 
  var fileurlField = '[DS] File-links';
  var indexUrlCell;
  var indexNameCell;

  if (sh.getLastColumn() === 0) {
    addGreenColumn(sh, '[DS] Files', 'Celdas verdes: para usar la opción "usar celdas verdes" debes sustituir esta nota con la URL de la plantilla.');
    var indexNameCell = headerIndexes(sh, flatten(sheetHeaderValues()), filenameField);
    sh.setColumnWidth(indexNameCell + 1, 300);
  };

  let headerValues = flatten(sheetHeaderValues());

  if (headerValues.indexOf(filenameField) !== -1) { // Coincidencia exacta
    indexNameCell = headerIndexes(sh, headerValues, filenameField);
  } else {
    addGreenColumn(sh, '[DS] Files', 'Celdas verdes: para usar la opción "usar celdas verdes" debes sustituir esta nota con la URL de la plantilla.');
    indexNameCell = headerIndexes(sh, headerValues, filenameField);
    sh.setColumnWidth(indexNameCell + 1, 300);
  }
  if (headerValues.indexOf(fileurlField) !== -1) { // Coincidencia exacta
    indexUrlCell = headerIndexes(sh, headerValues, fileurlField);
  } else {
    addGreenColumn(sh, '[DS] File-links', 'Celdas verdes: para usar la opción "usar celdas verdes" debes sustituir esta nota con la URL de la carpeta de destino.');
    indexUrlCell = headerIndexes(sh, headerValues, fileurlField);
    sh.setColumnWidth(indexUrlCell + 1, 300);
  }

  let dataReturn = {
    indexNameCell,
    indexUrlCell,
  };
  return dataReturn;
}

////////////////////////////// DOCUMENT DATA

function getDocumentData(greenCells, sh, templateID, destinationFolder, indexNameCell, indexUrlCell) {
  let docID;
  let destFolderID;

  if (greenCells) {
    let data = getGreenCellData(sh, indexNameCell, indexUrlCell);
    templateID = data.templateID;
    docID = getIdFromUrl(templateID);
    destinationFolder = data.destinationFolder;
    destFolderID = getIdFromUrl(destinationFolder);
    destinationFolderFinale = DriveApp.getFolderById(destFolderID);
  } else {
    docID = getIdFromUrl(templateID);
    destinationFolderFinale = handleDestinationFolder(destinationFolder);
  }

  return {docID, destinationFolderFinale};
}

function getGreenCellData(sh, indexNameCell, indexUrlCell) {
  return {
    templateID: sh.getRange(1, indexNameCell + 1).getNotes()[0][0],
    destinationFolder: sh.getRange(1, indexUrlCell + 1).getNotes()[0][0],
  };
}

function handleDestinationFolder(destinationFolder) {
  if (destinationFolder === '') {
    let rootFolder = DriveApp.getRootFolder(); // Create temporal folder if destFolder is empty
    destinationFolder = rootFolder.createFolder('Morph Document Studio Files');
  } else {
    let destFolderID = getIdFromUrl(destinationFolder);
    destinationFolder = DriveApp.getFolderById(destFolderID);
  }

  return destinationFolder;
}


////////////////////////////// MAIN MARKERS FUNCTIONS

/**
 * getTemplateMarkers
 * Core Function to get the list of Markers in different Google Type of Files
 */
function getTemplateMarkers(docID) {
  let identifier = {
    start: '{{',
    start_include: true,
    end: '}}',
    end_include: true,
  };

  let mainTemplateFile = DriveApp.getFileById(docID);
  let fileType = mainTemplateFile.getMimeType();
  let docMarkers;

  switch (fileType) {
    case MimeType.GOOGLE_DOCS:
      docMarkers = getDocMarkers(docID, identifier);
      break;
    case MimeType.GOOGLE_SLIDES:
      docMarkers = getSlidesMarkers(docID, identifier);
      break;
    default:
  }

  let dataReturn = {
    docMarkers,
    mainTemplateFile,
    fileType,
  };
  return dataReturn;
}

/**
 * getDocMarkers
 * Returns a list of Markers in a Google Docs File
 */
function getDocMarkers(templateID, identifier) {
  var body = DocumentApp.openById(templateID).getBody();
  var docText = body.getText();
  var docList = new Set();

  let textStart = 0;
  while ((textStart = docText.indexOf(identifier.start, textStart)) !== -1) {
    var textEnd = docText.indexOf(identifier.end, textStart + identifier.start.length);
    if (textEnd !== -1) {
      var startLen = identifier.start_include ? textStart : textStart + identifier.start.length;
      var endLen = identifier.end_include ? textEnd + identifier.end.length : textEnd;
      docList.add(docText.slice(startLen, endLen));
      textStart = textEnd + identifier.end.length;
    } else {
      break;
    }
  }

  return Array.from(docList);
}

/**
 * getSlidesMarkers
 * Returns a list of Markers in a Google Slides File
 */
function getSlidesMarkers(templateID, identifier) {
  var slides = SlidesApp.openById(templateID).getSlides();
  var markers = new Set();

  slides.forEach((slide) => {
    slide.getShapes().forEach((shape) => {
      var text = shape.getText().asString();
      let startIndex = 0;
      let endIndex = 0;

      while ((startIndex = text.indexOf(identifier.start, endIndex)) !== -1) {
        endIndex = text.indexOf(identifier.end, startIndex + identifier.start.length);
        if (endIndex !== -1) {
          var startLen = identifier.start_include ? startIndex : startIndex + identifier.start.length;
          var endLen = identifier.end_include ? endIndex + identifier.end.length : endIndex;
          markers.add(text.slice(startLen, endLen));
        } else {
          break;
        }
      }
    });
  });

  return Array.from(markers);
}



////////////////////////////// HELPER FUNCTIONS

/**
 * sheetHeaderValues
 * Devuelve el header de las columnas relacionadas con e-mail
 */
function sheetHeaderValues() {
  let headerValues = sh().getRange(1, 1, 1, sh().getLastColumn()).getValues(); 
  headerValues = transpose(headerValues);

  return headerValues;
}

/**
 * headerDropdownValues
 * Devuelve el header de las columnas relacionadas con e-mail
 */
function headerDropdownValues(filter) {
  let headerValues = sh().getRange(1, 1, 1, sh().getLastColumn()).getValues();
  headerValues = transpose(headerValues);

  if (filter) {
    // Filtrar los valores que contienen la palabra "email" (case insensitive)
    filterHeaderValues = headerValues.filter(row => {
      return row.some(cell => {
        return cell.toString().toLowerCase().includes(filter);
      });
    });

    return filterHeaderValues;
  } else {
    // Si no se proporciona un filtro, devolver los valores originales sin filtrar
    return headerValues;
  }
}

/**
 * headerIndex
 * Returns the Index of a header by name
 */
function headerIndex(sh, fieldName) {
  let dropdownValues = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues().flat().filter(r=>r!="");

  idx = dropdownValues.findIndex(item => item.includes(fieldName));
  return idx;
}

function headerIndexes(sh, headers, fieldName) {
  idx = headers.findIndex(item => item.includes(fieldName));
  return idx;
}






































////////////////////////////// SAVE AND DELETE CONFIGURATION

/**
 * saveDsConfiguration
 * Save the current Document Studio Configuration
 */
function saveDsConfiguration(rowData) {
  let formData = [
    rowData.dsActivate,
    rowData.emailActivate,

    rowData.templateID,
    rowData.greenCells,

    rowData.destinationFolder,
    rowData.fileName,
    rowData.exportFormat,
    rowData.permission1,
    rowData.permission2,
    rowData.permission3,
    rowData.numerationSwitch,

    rowData.emailField,
    rowData.emailSpecific,
    rowData.emailSender,
    rowData.emailSubject,
    rowData.emailMoreFields,
    rowData.emailBCC,
    rowData.emailReplyTo,
    rowData.emailMessage,
    rowData.emailAttachSwitch,
    rowData.emailAttachField,

    rowData.allDocuments,
    rowData.allEmails,
  ];

  let [dsActivate, emailActivate, templateID, greenCells, destinationFolder, fileName, exportFormat, permission1, permission2, permission3, numerationSwitch, emailField, emailSpecific, emailSender, emailSubject, emailMoreFields, emailBCC, emailReplyTo, emailMessage, emailAttachSwitch, emailAttachField, allDocuments, allEmails] = formData;

  PropertiesService.getDocumentProperties().setProperties({
    'DS Activate': dsActivate,
    'Email Activate': emailActivate,

    'Template Link': templateID,
    'Green Cells': Boolean(greenCells),

    'Destination Folder': destinationFolder,
    'Filename': fileName,
    'Export Format': exportFormat,
    'Permission 1': permission1,
    'Permission 2': permission2,
    'Permission 3': permission3,
    'Numeration Switch': numerationSwitch,

    'Email Field': emailField,
    'Email Specific': emailSpecific,
    'Email Sender': emailSender,
    'Email Subject': emailSubject,
    'Email More Fields': emailMoreFields,
    'Email BCC': emailBCC,
    'Email Reply To': emailReplyTo,
    'Email Message': emailMessage,
    'Email Attach Switch': emailAttachSwitch,
    'Email Attach Field': emailAttachField,

    'All Documents': allDocuments,
    'All Emails': allEmails,
  });
}

/**
 * deleteDsConfiguration
 * Delete the current Document Studio Configuration
 */
function deleteDsConfiguration() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
}

/**
 * deleteProperty
 * Delete a single Document Property
 */
function deleteProperty(e) {
  PropertiesService.getDocumentProperties().deleteProperty(e);
}

/**
 * getDocProperties
 * Returns all Document Properties
 */
function getDocProperties(e) {
  let props = PropertiesService.getDocumentProperties().getProperties();
  return props;
}

/**
 * getDocProperty
 * Returns a single Document Property
 */
function getDocProperty(e) {
  let prop = PropertiesService.getDocumentProperties().getProperty(e)
  return prop;
}




























/**
 * Gsuite Morph Tools - Morph Document Studio - Get Markers
 * Developed by alsanchezromero
 *
 * Copyright (c) 2022 Morph Estudio
 */

function getMarkers(rowData) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();

  // Data + Variables

  let formData = [rowData.templateID, rowData.greenCells, rowData.purgeMarkers];
  let [docURL, greenCells, purgeMarkers] = formData;

  let filenameField = '[DS] Files'; let fileurlField = '[DS] File-links'; let mailurlField = '[DS] Email-links';

  let dataReturn = getGreenColumns(sh, filenameField, fileurlField);

  let indexData = [
    dataReturn.indexNameCell,
    dataReturn.indexUrlCell,
  ];

  let [indexNameCell, indexUrlCell] = indexData;

  let docID;
  let sheetEmpty = sh.getLastColumn();

  if (greenCells) {
    docURL = sh.getRange(1, indexNameCell + 1).getNotes();
    docID = getIdFromUrl(docURL[0][0]);
  } else {
    docID = getIdFromUrl(docURL);
  }

  dataReturn = getTemplateMarkers(docID)

  indexData = [
    dataReturn.docMarkers,
    dataReturn.mainTemplateFile,
    dataReturn.fileType,
  ];

  let [docMarkers, mainTemplateFile,fileType] = indexData;

  let notAllMarkersChanged; let headerValues; let updatedValues = [];
  
  let driverArray = docMarkers.flat(); // Slicing {} markers for enhancing interface.
  driverArray.forEach((el) => {
    let sliced = el.slice(2, -2);
    updatedValues.push(sliced);
  });

  // Purge Markers

  if (purgeMarkers) {

    headerValues = flatten(sheetHeaderValues()).filter(e => e !== filenameField && e !== fileurlField && e !== mailurlField);

    if (headerValues.length != 0) {
      notAllMarkersChanged = updatedValues.some(element => {
        return headerValues.includes(element);
      })

      if (notAllMarkersChanged) {
        columnRemover(sh, updatedValues, headerValues);
      } else {
        indexNameCell = headerIndex(sh, filenameField);
        sh.deleteColumns(1, indexNameCell);
      }
    }
  }

  // Add New Markers

  headerValues = flatten(sheetHeaderValues()).filter(e => e !== filenameField && e !== fileurlField && e !== mailurlField);

  updatedValues.forEach((a, index) => {
    if (headerValues.indexOf(a) === -1) {
      if (index === 0) {
        sh.insertColumns(index + 1);
      } else {
        sh.insertColumnAfter(index);
      }
      sh.setColumnWidth(index + 1, 150);
      let headerCell = sh.getRange(1, index + 1, 1, 1);
      headerCell.setValue(a);
    }
  });

  // Style

  indexNameCell = headerIndex(sh, filenameField);

  if (sheetEmpty <= 2 || notAllMarkersChanged === false) {
    sh.getRange(1, 1, 1, indexNameCell).clearFormat();
  }

  sh.getRange(1, 1, 1, sh.getMaxColumns()).setFontWeight('bold').setHorizontalAlignment('center');
  sh.setFrozenRows(1);
  removeEmptyColumns();
}


function columnRemover(sh, updatedValues, headerValues) {
  let deleteColumn = [];
  headerValues.forEach((a, index) => {
    let i = updatedValues.indexOf(a);
    if (i === -1) {
      deleteColumn.push(index + 1);
    }
  });

  let lastColmn;
  for (let j = headerValues.length; j > 0; j--) {
    if (deleteColumn.indexOf(j) === -1) {
    } else {
      lastColmn = sh.getLastColumn();
      if (lastColmn > 1) {
        sh.deleteColumn(j);
      } else {
        sh.insertColumnAfter(lastColmn);
        sh.deleteColumn(j);
      }
    }
  }
}



















/**
 * emailSending
 * Check if a column is a Document Studio Column
 */
function emailSending(sh, rowData, indexUrlCell, userMail, dateNow) {

  var {
    emailField,
    emailSpecific,
    emailSender,
    emailSubject,
    emailMessage,
    emailBCC,
    emailReplyTo,
    emailAttachSwitch,
    emailAttachField,
    allEmails,
  } = rowData;

  var mailurlField = '[DS] Email-links';
  let indexFile;
  var indexEmail;
  var indexEmailLinks;

  /** Get Green Column **/

  let dropdownValues = flatten(sheetHeaderValues());

  if (dropdownValues.indexOf('[DS] Email-links') > -1) {
    indexEmailLinks = headerIndex(sh, mailurlField);
  } else {
    addGreenColumn(sh, '[DS] Email-links', '');
    indexEmailLinks = headerIndex(sh, mailurlField);
    sh.setColumnWidth(indexUrlCell + 1, 300);
  }

  let dropdownValuesWithMarker = dropdownValues.map((i) => `{{${i}}}`);
  let lastColumn = sh.getLastColumn();

  /** Index of Main Columns (Email & File) **/

  indexEmail = headerIndex(sh, emailField);

  if (emailAttachSwitch) {
    if (emailAttachField === 'default') {
      indexFile = headerIndex(sh, fileurlField);
    } else {
      indexFile = headerIndex(sh, emailAttachField);
    }
  }

  let headerValues;
  let fileID;
  let file;
  let adress;
  let emailSubjectReplaced;
  let emailMessageReplaced;
  let linkMail;
  let mailId;

  let emailRows = sh.getDataRange().getValues();

  /** Loop Start **/

  for (var index = 1; index < emailRows.length; index++) {
    var row = emailRows[index];

    if (row[indexEmail] === '') { continue; }
    if (row[indexEmailLinks] !== "" && allEmails === false) { continue; }

    /** Replace Markers in Subject and Message **/

    if (dropdownValuesWithMarker.some((v) => emailSubject.includes(v))) {
      for (let i = 0; i < lastColumn; i += 1) {
        headerValues = sh.getRange(1, i + 1).getValue();
        headerValues = `${`{{${headerValues}}}`}`;
        if (emailSubject.indexOf(headerValues) > -1) {
          emailSubjectReplaced = emailSubject.replace(headerValues, row[i]);
        }
      }
    } else { emailSubjectReplaced = emailSubject; }

    if (dropdownValuesWithMarker.some((v) => emailMessage.includes(v))) {
      for (let i = 0; i < lastColumn; i += 1) {
        headerValues = sh.getRange(1, i + 1).getValue();
        headerValues = `${`{{${headerValues}}}`}`;
        if (emailMessage.indexOf(headerValues) > -1) {
          emailMessageReplaced = emailMessage.replace(headerValues, row[i]);
        }
      }
    } else { emailMessageReplaced = emailMessage; }

    /** Send E-Mails **/

    adress = row[indexEmail];

    let mailOptions = {
      name: emailSender || 'Morph Estudio',
      cc: emailSpecific,
      bcc: emailBCC,
      replyTo: emailReplyTo,
      htmlBody: emailMessageReplaced || 'Este correo ha sido enviado automáticamente desde el Workspace de Morph Estudio',
    };

    if (emailAttachSwitch) {
      try {
        fileID = getIdFromUrl(row[indexFile]);
        file = DriveApp.getFileById(fileID);
        mailOptions.attachments = [file];
      } catch (error) {
        // Añadir errores a un cuadro de errores al final de la ejecución.
      }
    }

    mailId = GmailApp.createDraft(adress, emailSubjectReplaced || 'Morph Document Studio', '', mailOptions).send().getId();

    /** Set Mail-Link in Google Sheet **/

    linkMail = `https://mail.google.com/mail/u/0/?tab=rm#sent/${mailId}`;
    sh.getRange(index + 1, indexEmailLinks + 1).setValue(linkMail).setNote(null).setNote(`Enviado por ${userMail} el ${dateNow}`);
  };
}

function isValidFile(url) {
url = 'https://drive.google.com/file/d/1MB-bbMI77augi246fLiPb0JmM2FrNP0b'
esArchivoValido(url).then(result => console.log(result)); // Devolverá true
}

async function esArchivoValido(url) {
  try {
    const response = await fetch(url, { method: 'HEAD' });
    return response.ok && response.headers.get('Content-Type').startsWith('application/pdf');
  } catch (error) {
    return false;
  }
}

function parseImageDimension(value) {

  let dimension = null;
  let number = null;
  let customDimension = false;

  // Expresión regular para buscar el patrón {w=number} o {h=number}
  const regex = /{(w|h)=(\d+)}/;

  // Verificar si el valor coincide con el patrón
  const match = value.match(regex);

  if (match) {
    dimension = match[1]; // 'w' para width o 'h' para height
    number = parseInt(match[2], 10);
    customDimension = true;

    value = value.replace(regex, '').trim();
  }

  return {
    value,
    dimension,
    number,
    customDimension
  };
}
