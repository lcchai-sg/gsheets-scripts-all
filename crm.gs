//code to send email when cell containing (days since last contact) / (days to contact)

function checkForAlerts() {
  
  var contactList = [];
  //get the sheet 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CRM");
  //find number of ppl in sheet
  var pplCountCell = sheet.getRange("D1");
  var entryCount = pplCountCell.getValue();

  // check for ppl to contact
  for (let i = 3; i < (entryCount + 3); i++){
    var cellRange = sheet.getRange("A" + i.toString());
    var value = cellRange.getValue();
    if (value > 1){
      var nameCell = cellRange.offset(0,1);
      var name = nameCell.getValue()
      contactList += name + ', ';
    }   
  }

  //compose email
  var address = 'jacobs@noahsarcpartners.com'
  var subject = 'Contact List'
  var body = 'Today, you must contact: ' + contactList;
 

  //if empty list, send empty email
  if (contactList.length < 1){
    var body = "No one to contact"
    MailApp.sendEmail(address,subject,body);
    return;
  }

  //otherwise, send email
  MailApp.sendEmail(address,subject,body);
}
