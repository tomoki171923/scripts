// ----------------------------------------- //
// Create CSV file from Google Spread sheet .
// How to use    : execute the function "createCsv" on your Google Spread sheet.
// ----------------------------------------- //

function _createFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss.getId();
  var parentFolder = DriveApp.getFileById(ssId).getParents();
  var folderId = parentFolder.next().getId();
  var folder = DriveApp.getFolderById(folderId);
  var date = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyyMMddHHmmss");
  var foldername = "outputCsv_" + date;
  var newfolder = folder.createFolder(foldername);
  return newfolder;
}

function _loadData(sheet) {
  var data = sheet.getDataRange().getValues();
  var csv = '';
  for(var i = 0; i < data.length; i++) {
    csv += data[i].join(',') + "\r\n";
  }
  return csv;
}

function _writeDrive(csv,fileName,drive) {
  var contentType = 'text/csv';
  fileName += ".csv";
  var charset = 'utf-8';
  var blob = Utilities.newBlob('', contentType, fileName).setDataFromString(csv, charset);
  drive.createFile(blob);
}

function createCsv() {
  var objSheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var drive = _createFolder();
  for (i = 0; i < objSheets.length; i++) {
    var sheet = objSheets[i];
    var csvData = _loadData(sheet);
    _writeDrive(csvData, sheet.getName(),drive);
  } 
}



