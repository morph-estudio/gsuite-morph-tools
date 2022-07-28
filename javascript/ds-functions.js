function doGet() {
  return HtmlService.createHtmlOutputFromFile('html/document-studio');
}

// GET-MARKERS FUNCTIONS

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

function columnRemover(sh, updatedValues, numberOfGreenCells) {
  let lastColHeaders = sh.getLastColumn() - numberOfGreenCells;
  let headerRange = sh.getRange(1, 1, 1, lastColHeaders);
  let headerValues = headerRange.getValues()[0];

  let deleteColumn = [];
  headerValues.forEach((a, index) => {
    let i = updatedValues.indexOf(a);
    if (i === -1) {
      deleteColumn.push(index + 1);
    }
  });

  let lastColmn;
  for (let j = lastColHeaders; j > 0; j--) {
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

// DOCUMENT-STUDIO FUNCTIONS

function isValidHttpUrl(str) {
  let pattern = new RegExp('^(https?:\\/\\/)?' // protocol
    + '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' // domain name
    + '((\\d{1,3}\\.){3}\\d{1,3}))' // OR ip (v4) address
    + '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' // port and path
    + '(\\?[;&a-z\\d%_.~+=-]*)?' // query string
    + '(\\#[-a-z\\d_]*)?$', 'i'); // fragment locator
  return !!pattern.test(str);
}

function isImage(url) {
  return /\.(jpg|jpeg|png|webp|avif|gif|svg)$/.test(url);
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
    // Logger.log(replaceAllTextRequest)
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

function formatDropdown() { // If dropdown options are in a Google Sheet
  const ss = SpreadsheetApp.getActive();
  let sh = SpreadsheetApp.getActiveSheet();
  let ws = ss.getSheetByName('DS OPTIONS') || ss.insertSheet('DS OPTIONS', 1);

  return ws.getRange(2, 1, ws.getLastRow() - 1, 1).getValues();
}
