function getBlob() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var outputsheet = ss.getSheetByName("AllWarrantyUnits");

  //build URL
  var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?";
  //URL parameters
  var params = 'format=xlsx' + '&gid=' + outputsheet.getSheetId();

  //get token for UrlFetch
  var token = ScriptApp.getOAuthToken();
  var response = UrlFetchApp.fetch(url + params, { 
    'headers': {
      'Authorization': 'Bearer ' + token
    }
  }); 
  var theBlob = response.getAs('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  //var theBlob = response.getBlob().setName("warrantyUnits.xlsx");
  return theBlob
}

function sendEmail() {
    MailApp.sendEmail('##REDACTED##', "test", "testinnggg", {
      attachments: getBlob()
  }
);
}
