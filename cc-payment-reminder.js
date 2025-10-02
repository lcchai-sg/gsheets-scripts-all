/*
  This script works with a Google Sheet I created to send a reminder via email when a credit card payment is due.
  
  My Google sheet allows me to select how to pay back an expense (in weekly, fortnightly or monthly frequencies)
  and uses this information to calculate when a payment is due (column A), how many days until a payment needs 
  to be made (column B) and how much to pay back (column C).
  
  A trigger has been set up so that this is checked everyday and if a payment is due on the day, an email will 
  sent to me with the total amount due for that day.
*/

function reminder() {
  // set the first sheet (index 0) in the spreadsheet as the active sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.setActiveSheet(spreadsheet.getSheets()[0]);
  var sheet = spreadsheet.getActiveSheet();

  // data to look at starts at row 2 (headings in row 1)
  var startRow = 2;

  // countif formula in cell of row 1 column 16 to count total number of rows
  var lastRow = sheet.getRange(1, 16).getValue();

  // get data from 2nd column excluding row 1 (containing heading)
  var range = sheet.getRange(2, 2, lastRow - startRow + 1, 1);
  var numRows = range.getNumRows();
  var days_left_values = range.getValues();

  // get data from 3rd column excluding row 1 (containing heading)
  range = sheet.getRange(2, 3, lastRow - startRow + 1, 1);
  var amount_owed_values = range.getValues();

  var html = "";

  // iterate through days_left_values to check if any are equal to 0 and if so append the html string
  for (var i = 0; i < numRows; i++) {
    var days_left = days_left_values[i][0];
    if (days_left == 0) {
      var amount_owed = amount_owed_values[i][0];
      html = "<body><h1>Reminder:</h1><p>Payment of <strong>&pound;" + amount_owed + "</strong> is due today.</p></body>";
    }
  }
  
  // if the html string has been changed, send email reminder
  if (html != "") {
    MailApp.sendEmail("myemail@address.com,myemail@address.com", "Credit Card Payment Due", html, { htmlBody: html });
  }
};
