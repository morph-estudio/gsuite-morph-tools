function getCSVFilesData(rowData, counter) {

  const sh = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  for (let i = 0; i <= counter; i++) {

    let formData = [
      rowData[`csvURL${i}`],
      rowData[`csvCELL${i}`]
    ];

    let [csvURL, csvCELL] = formData;

    let fileURL = csvURL;
    let fileID = getIdFromUrl(fileURL);
    let file = DriveApp.getFileById(fileID);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    let fetchURL = `https://drive.google.com/uc?id=${fileID}&x=.csv`;

    let csvContent = UrlFetchApp.fetch(fetchURL);
    let csvData = Utilities.parseCsv(csvContent);

    SpreadsheetApp.flush();
    sh.getRange(sh.getRange(csvCELL).getRowIndex(), sh.getRange(csvCELL).getColumn(), csvData.length, csvData[0].length).setValues(csvData);
    SpreadsheetApp.flush();
  }
}

function getSavedSheetProperties(rowData) {
  PropertiesService.getDocumentProperties()
    .setProperties(rowData);
}
