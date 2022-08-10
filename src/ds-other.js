function getDocItems(docID, identifier) {
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
  while (textStart > -1) {
    textStart = doc.indexOf(identifier.start);

    if (textStart === -1) {
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

function getSlidesItems(docID, identifier) {
  let slides = SlidesApp.openById(docID).getSlides();

  let sumaTextos = [];
  slides.forEach((slide) => {
    let shapes = (slide.getShapes());
    shapes.forEach((shape) => {
      let textito = shape.getText().asString();
      sumaTextos.push(textito);
    });
  });

  let docText = sumaTextos.toString();

  // Check if search characters are to be included.
  let startLen = identifier.start_include ? 0 : identifier.start.length;
  let endLen = identifier.end_include ? 0 : identifier.end.length;

  // Set up the reference loop
  let textStart = 0;
  let doc = docText;
  let docList = [];

  // Loop through text grab the identifier items. Start loop from last set of end identfiers.
  while (textStart > -1) {
    textStart = doc.indexOf(identifier.start);

    if (textStart === -1) {
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

function isGreenCell(lastCell) {
  let mycell = SpreadsheetApp.getActiveSheet().getRange(1, lastCell);
  let bgHEX = mycell.getBackground();
  if (bgHEX == '#ecfdf5') {
    return true;
  }
  return false;
}

function getInternallyMarkers(docID) {
  let identifier = {
    start: '{{',
    start_include: true,
    end: '}}',
    end_include: true,
  };

  let gDocTemplate = DriveApp.getFileById(docID);
  let fileType = gDocTemplate.getMimeType();
  let docMarkers;

  switch (fileType) {
    case MimeType.GOOGLE_DOCS:
      docMarkers = getDocItems(docID, identifier);
      break;
    case MimeType.GOOGLE_SLIDES:
      docMarkers = getSlidesItems(docID, identifier);
      break;
    default:
  }

  let dataReturn = {
    docMarkers,
    gDocTemplate,
    fileType,
  };
  return dataReturn;
}

// DOCUMENT-STUDIO FUNCTIONS

function getGreenColumns(sh, filenameField, fileurlField) {
  if (sh.getLastColumn() === 0) {
    addGreenColumn(sh, '[DS] Files', 'Celdas verdes: para usar la opción "usar celdas verdes" debes introducir en esta casilla una nota con la URL de la plantilla.');
    var indexNameCell = fieldIndex(sh, filenameField);
    sh.setColumnWidth(indexNameCell + 1, 300);
  };

  let dropdownValues = flatten(emailDropdown());

  if (dropdownValues.indexOf(filenameField) > -1) {
    var indexNameCell = fieldIndex(sh, filenameField);
  } else {
    addGreenColumn(sh, '[DS] Files', 'Celdas verdes: para usar la opción "usar celdas verdes" debes introducir en esta casilla una nota con la URL de la plantilla.');
    var indexNameCell = fieldIndex(sh, filenameField);
    sh.setColumnWidth(indexNameCell + 1, 300);
  }
  if (dropdownValues.indexOf(fileurlField) > -1) {
    var indexUrlCell = fieldIndex(sh, fileurlField);
  } else {
    addGreenColumn(sh, '[DS] File-links', 'Celdas verdes: para usar la opción "usar celdas verdes" debes introducir en esta casilla una nota con la URL de la carpeta de destino.');
    var indexUrlCell = fieldIndex(sh, fileurlField);
    sh.setColumnWidth(indexUrlCell + 1, 300);
  }

  let dataReturn = {
    indexNameCell,
    indexUrlCell,
  };
  return dataReturn;
}

function addGreenColumn(sh, headerTitle, cellNote) {
  let lastColmn = sh.getLastColumn();

  if (lastColmn === 0) {
  sh.insertColumns(1);
  sh.getRange(1, lastColmn + 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue(headerTitle)
    .setNote(cellNote);
  } else {
  sh.insertColumnAfter(lastColmn);
  sh.getRange(1, lastColmn + 1).setBackground('#ECFDF5').setFontColor('#34a853').setValue(headerTitle)
    .setNote(cellNote);
  }

}

function emailDropdown() { // If dropdown options are in a Google Sheet
  let sh = SpreadsheetApp.getActive().getActiveSheet();
  let dropdownValues = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues(); 
  dropdownValues = transpose(dropdownValues);
  return dropdownValues;
}

function fieldIndex(sh, fieldName) {
  SpreadsheetApp.flush();
  let dropdownValues = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues();
  dropdownValues = transpose(dropdownValues);
  dropdownValues = [].concat.apply([], dropdownValues);
  idx = dropdownValues.findIndex(item => item.includes(fieldName));
  return idx;
}

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

function imageFromTextDocsCustomWidth(body, searchText, image, width) {
  let next = body.findText(searchText);
  let atts = body.getAttributes();
  if (!next) return;
  let r = next.getElement();
  r.asText().setText('');
  let img = r.getParent().asParagraph().insertInlineImage(0, image);
  let w = img.getWidth();
  let h = img.getHeight();
  let sw = atts['PAGE_WIDTH'];

  if (width > sw) {
    img.setWidth(sw);
    img.setHeight((sw * h) / w);
  } else {
    img.setWidth(width);
    img.setHeight((width * h) / w);
  }
}

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
  let result = Docs.Documents.batchUpdate(batchUpdateRequest, copyId);
}

function imageFromTextSlides(slides, searchText, imageUrl) {
  slides.forEach((slide) => {
    slide.getShapes().forEach((s) => {
      if (s.getText().asString().includes(searchText)) {
        s.replaceWithImage(imageUrl, true);
      }
    });
  });
}

function replaceSlideText(slides, replaceText, markerText, copyId) {
  slides.forEach((slide) => {
    let pageElementId = slide.getObjectId();
    let resource = {
      requests: [{
        replaceAllText: {
          pageObjectIds: [pageElementId],
          replaceText,
          containsText: { matchCase: false, text: markerText },
        },
      }],
    };
    let result = Slides.Presentations.batchUpdate(resource, copyId);
  });
}

function setDocProperties(rowData){

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

function deleteProperties() {
  PropertiesService.getDocumentProperties().deleteAllProperties();
}

function getDocProperties(){

  /*
  let documentProperties = PropertiesService.getDocumentProperties().getProperties()
  let a = documentProperties['Email Message'];
  let b = documentProperties['All Emails'];
  Logger.log('all property: ' + a + b)
  */

  return PropertiesService.getDocumentProperties().getProperties();
}
