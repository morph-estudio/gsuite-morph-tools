function doGet() {
  return HtmlService.createHtmlOutputFromFile('html/email');
}

/*
 * Gsuite Morph Tools - Morph Document Studio 2
 * Developed by alsanchezromero
 * Created on Mon Jul 25 2022
 *
 * Copyright (c) 2022 Morph Estudio
*/

function dsSendMail(rowData) {

const ss = SpreadsheetApp.getActive();
const sh = SpreadsheetApp.getActiveSheet();

let formData = [
  rowData.emailField,
  rowData.emailSpecific,
  rowData.emailSender,
  rowData.emailSubject,
  rowData.emailBCC,
  rowData.emailReplyTo,
  rowData.emailMessage,
  rowData.emailAttach,
  rowData.emailAttachField,
];

let [emailField, emailSpecific, emailSender, emailSubject, emailBCC, emailReplyTo, emailMessage, emailAttach, emailAttachField] = formData;
let userMail = Session.getActiveUser().getEmail();

// Get index of selected columns

let dropdownValues = emailDropdown();
dropdownValues = flatten(dropdownValues); Logger.log(dropdownValues);

let indexEmail; let indexFile; let indexNumber;

function fieldIndex(dropdownValues, fieldName) {
  dropdownValues.forEach((col, index) => {
    if (col.indexOf(fieldName) > -1) {
      indexNumber = index;
    }
  });
  return indexNumber;
}

indexEmail = fieldIndex(dropdownValues, emailField); Logger.log(indexEmail);
indexFile = fieldIndex(dropdownValues, emailAttachField); Logger.log(indexFile);


/*
function getColumn(sh, index) {
  let rg1 = sh.getRange(2,index1 + 1,sh.getLastRow() -1,1).getValues();
  let rg2 = sh.getRange(2,index2 + 1,sh.getLastRow() -1,1).getValues();

  let rf1 = [];
  let rf2 = [];

  rg1.forEach((m, index) => {
    if (m != '' && )
    rf1.push(m)

  })
*/

// Get array of values (email-document)

function getColumn(sh, index) {
  let values = sh.getRange(2, index + 1, sh.getLastRow() - 1, 1).getValues();
  values = values.filter(String).length;
  values = sh.getRange(2, index + 1, values, 1).getValues();
  values = flatten(values)
  return values;
}

let emailValues = getColumn(sh, indexEmail); Logger.log('valores 1 ' + emailValues);
let fileValues = getColumn(sh, indexFile); Logger.log('valores 2 ' + fileValues);

// Set emails

let docID; let file;

emailValues.forEach((adress, index) => {

docID = getIdFromUrl(fileValues[index]); Logger.log(docID);
file = DriveApp.getFileById(docID[0]);

/**/
  MailApp.sendEmail({
    to: adress,
    subject: emailSubject || 'Morph Document Studio',
    name: emailSender || 'Morph Estudio',
    cc: emailSpecific,
    bcc: emailBCC,
    replyTo: emailReplyTo,
    htmlBody: emailMessage || 'Este correo ha sido enviado automáticamente desde el Google Workspace de Morph Estudio.',
    attachments: [file],
  });

});

if (emailAttach) {
}

}


function emailDropdown() { // If dropdown options are in a Google Sheet
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getActiveSheet();
  //let ws = ss.getSheetByName('DS OPTIONS') || ss.insertSheet('DS OPTIONS', 1);
  let dropdownValues = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues(); 
  dropdownValues = transpose(dropdownValues);
  return dropdownValues;
}

function transpose(a) {
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}

function flatten(arrayOfArrays){
  return [].concat.apply([], arrayOfArrays);
}
