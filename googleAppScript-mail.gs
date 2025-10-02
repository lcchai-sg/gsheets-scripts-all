var EMAIL_SENT = "EMAIL_SENT";
function myFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;  // First row of data to process
  var numRows = 2;   // Number of rows to process
  
  var range = sheet.getRange(1, 1, 3, 8);
  var values = range.getValues();

  var i = 0;
 for (var row in values) {
   
   
   //for (var col in values[row]) {
     //Logger.log(values[row][col]);
     var emailAddress =  values[row][3];
      var message = "Hi " + values[row][0] +  ", \n" +  "Thank you \n"
   message +=  "\nName :" + values[row][4] + "\n";
   Logger.log(emailAddress); 
   Logger.log(message);
     
  // }
   Logger.log("--------------------");
   
   var subject = "Subject";
   MailApp.sendEmail(emailAddress, "no-reply@myemail.com", subject, message);
   //sheet.getRange(i, 8).setValue(EMAIL_SENT);
   SpreadsheetApp.flush();
   
   i += 1;
 }
  
}
