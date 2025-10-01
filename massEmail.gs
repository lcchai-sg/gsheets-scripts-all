//function to send mass emails
//to do - fix truth values 

function sendEmails() {
  // Replace "Sheet1" with the name of your sheet
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("toSend");

// Get the data from the sheet
var data = sheet.getDataRange().getValues();

// Loop through each row of data
for (var i = 0; i < data.length; i++) {
  var row = data[i];
  var email = row[1]; // Replace 1 with the index of the email column
  var shouldSend = row[8]; // Replace 8 with the index of the "send email" column
  var subject = "INSERT SUBJECT " + row[10] + "s"; // Replace 10 with the index of the subject column
  var name = row[0].split(" ")[0]; // Replace 0 with the index of the name column and split the string to get the first word

  if (shouldSend && !row[9]) { // Replace 9 with the index of the "email sent" column
    // Send the email
    MailApp.sendEmail(email, subject, 
  'Hey ' + name + '!\n\ INSERT BODY TEXT\n\n-Noah Jacobs');

Logger.log("Before: " + sheet.getRange(i+1, 9).getValue());

    // Mark the email as sent
    sheet.getRange(i+1, 9).setValue("TRUE"); // Replace 9 with the index of the "email sent" column
    Logger.log("After: " + sheet.getRange(i+1, 9).getValue());

  }
  
}

}
