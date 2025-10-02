var warranty_units = [];
const sheets_to_parse = ["Apple", "Dell", "HP", "Lenovo"];

function compile() {
  if(is_weekend()) {
    return; //do not run if weekend
  }
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = activeSpreadsheet.getSheetByName("Apple");
  var data=sheet.getDataRange().getValues();
  warranty_units.push(data[0].slice(2)); //add header to array
  sheets_to_parse.forEach((element) => parseSheet(activeSpreadsheet.getSheetByName(element)));

  dumpToSheet();
  sendEmail();
}

function is_weekend() { //returns true if today is saturday or sunday
  var today = new Date();
  var day = today.getDay();
  if(day == 6 || day == 0) {
    return true;
  }
  return false;
}

function dumpToSheet() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var outputSheet = activeSpreadsheet.getSheetByName('AllWarrantyUnits');
  if (outputSheet != null) {
    outputSheet.clearContents();
  } else {
    outputSheet = activeSpreadsheet.insertSheet("AllWarrantyUnits");
  }
  range = outputSheet.getRange(1,1, warranty_units.length, warranty_units[0].length);
  range.setValues(warranty_units)
}


function parseSheet(sheet) {
  var data=sheet.getDataRange().getValues();
  var length = data
  var warranty_column = getWarrantyColumn(sheet)
  for(var i = 1; i < data.length; i++) {
    if(data[i][warranty_column].indexOf('y') > -1) {
      var processed = data[i][warranty_column + 1];
      if(processed == "" || processed == "AC+") {
        warranty_units.push(data[i].slice(2));
      }
    }
  }
}

function getWarrantyColumn(sheet) {
  var data=sheet.getDataRange().getValues();

  for (var i=0; i < data[0].length; i++) {
    if (data[0][i] == 'warranty') {
      console.log("found it");
      return i;
    }
  }
}
