//Set trigger to run on edit
var fileId = '1LcdXBRNGZO0AQ1z_8_gsD3ELggLDu5os'; // update this file with it's ID
var txtFile = DriveApp.getFileById(fileId);
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];

function export(){
  var stringContent = sheet.getRange("A:D").getValues()
  txtFile.setContent(stringContent);
}

function clearTXT() {
  txtFile.setContent("");
}
