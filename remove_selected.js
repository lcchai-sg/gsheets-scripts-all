function availableSlots(){
  var form = FormApp.openByUrl('URL_OF_YOUR_FORM');
  var slots = SpreadsheetApp
                .getActiveSpreadsheet()
                .getRange("slots!A2:C10")
                .getValues();
  var choice = [];
  for (s in slots){
    if (slots[s][0] != "" && slots[s][2] > 0){
      choice.push(slots[s][0]);
    }
  }
  var formItems = form.getItems(FormApp.ItemType.LIST);
  formItems[0].asListItem().setChoiceValues(choice);
}

//blatantly taken from https://mashe.hawksey.info/2014/11/dynamically-remove-google-form-options-after-they-have-been-selected-by-someone-or-reach-defined-limits/
//who did a great job...it works, why fix it
