//checks if bday. if so, sends email

function checkForBday() {
  
  var bDayList = [];
  //get the sheet 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("importantDates");
  //find number of ppl in sheet
  var pplCountCell = sheet.getRange("D1");
  var entryCount = pplCountCell.getValue();

  // check for ppl to contact; i is starting row
  for (let i = 2; i < (entryCount + 2); i++){
    var bDayCell = sheet.getRange("A" + i.toString());
    var bDay = bDayCell.getValue();
    var currentDateCell = sheet.getRange("F1");
    var currentDate = currentDateCell.getValue();
    var bDayDay = new Date(bDay).getDate();
    var bDayMonth = new Date(bDay).getMonth();
    var currentDateDay = new Date(currentDate).getDate();
    var currentDateMonth = new Date(currentDate).getMonth();

    // Compare the day and month of the birthday and current date
    if (bDayDay === currentDateDay && bDayMonth === currentDateMonth){
      var nameCell = bDayCell.offset(0,1);
      var name = nameCell.getValue()
      bDayList += 'It is ' + name + '`s Birthday. ';
    }   
  }

  //compose email
  var address = 'jacobs@noahsarcpartners.com'
  var subject = 'Bday List'
  var body = bDayList
 

  //if empty list, send empty email
  if (bDayList.length < 1){
    return;
  }

  //otherwise, send email
  MailApp.sendEmail(address,subject,body);
}
