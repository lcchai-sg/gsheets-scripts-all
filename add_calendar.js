var moment = Moment.load();
var GLOBAL = {
  formUrl: 'https://formurl.com', //enter form url here
  calendarID: 'user@gmail.com', //enter the email of the account you want to add the event to here
  formMap : { //reflects the form column heads
    dateTime : "Tour Times",
    name: "Name (first, last)",
    phoneNumber: "Phone Number",
    email: "Email"
  }
}
function getFormResponse(){
  var form = FormApp.openByUrl(GLOBAL.formUrl);
  var responses = form.getResponses();
  var length = responses.length;
  var lastResponse = responses[length-1];
  var itemResponses = lastResponse.getItemResponses();
  var eventObject = {};
  for (var i = 0 ; i<itemResponses.length; i++) {
    //Get the title of the form item being iterated on
    var thisItem = itemResponses[i].getItem().getTitle();
        //get the submitted response to the form item being
        //iterated on
    var thisResponse = itemResponses[i].getResponse();
    switch (thisItem) {
      case GLOBAL.formMap.dateTime:
        var date = thisResponse.slice(5,14);
        var startTime = thisResponse.slice(15,23);
        var endTime = thisResponse.slice(26,37);
        var convertedStartDate = moment(date + ' ' + startTime, 'MM-DD-YY hh:mm a').format();
        var convertedEndDate = moment(date + ' ' + endTime, 'MM-DD-YY hh:mm a').format();
        eventObject.startTime = convertedStartDate;
        eventObject.endTime = convertedEndDate;
        break;
      case GLOBAL.formMap.name:
        eventObject.name = thisResponse;
        break;
      case GLOBAL.formMap.phoneNumber:
        eventObject.phone = thisResponse;
        break;
      case GLOBAL.formMap.email:
        eventObject.email = thisResponse;
        break;
    }
  }
  return eventObject;
}
function createCalendarEvent(eventObject){
  var startTime = eventObject.startTime;
  var endTime = eventObject.endTime;
  var name = eventObject.name + ' Tour';
  var email = eventObject.email;
  var phone = eventObject.phone;
  var cal = CalendarApp.getCalendarById(GLOBAL.calendarID)
  var event = cal.createEvent(name, new Date(startTime), new Date(endTime), {description: name, location: email});
  return event;
}
function onFormSubmit(){
  var eventObject = getFormResponse();
  var event = createCalendarEvent(eventObject);
}
