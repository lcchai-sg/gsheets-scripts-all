var calendarId = ""; // insert ID here
var calendarSetColNumber = 6;
var endDateColNumber = 5;
var startDateColNumber = 4;
var titleColNumber = 3;
var descriptionColNumber = 2;
var fromTimeStampColNumber = 1;

function getLatestAndSubmitToCalendar() { 
  var sheet = SpreadsheetApp.getActiveSheet();  
  var rows = sheet.getDataRange();  
  var lastRow = rows.getLastRow();   
  var calendarSetCell = sheet.getRange(lastRow, calendarSetColNumber,1,1);
  if (!calendarSetCell.isBlank()) return; // check if calendar event has already been created
   
  var startDate = sheet.getRange(lastRow,startDateColNumber,1,1).getValue();  
  var endDate = sheet.getRange(lastRow,endDateColNumber,1,1).getValue();  
  var submittedOn = "Submitted on :"+sheet.getRange(lastRow,fromTimeStampColNumber,1,1).getValue();  
  var description = "Reserved by :"+sheet.getRange(lastRow,descriptionColNumber,1,1).getValue()+"\n"+submittedOn;  
  var title = sheet.getRange(lastRow,titleColNumber,1,1).getValue()+" / "+sheet.getRange(lastRow,descriptionColNumber,1,1).getValue();
  
  createEvent(title,startDate,endDate,description);
  calendarSetCell.setValue("done");
}  

function createEvent(title,startDate,endDate,desc) {  
  var cal = CalendarApp.getCalendarById(calendarId);  
  var start = new Date(startDate);  
  var end = new Date(endDate);  
  var opts = {description: desc, location: 'TIU English Plaza' };
  cal.createEvent(title, start, end, opts);
}
