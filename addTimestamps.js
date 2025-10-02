function onEdit(e) { 
  var timezone = "GMT+9"; 
  var timestamp_format = "HH:mm"; 

  var sheet = e.range.getSheet();
  if (sheet.getSheetName() == 'Sheet1') {    // Name of the sheet to track
    var editColumn = e.range.getColumn(); 
    var row = e.range.getRow(); 
    var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues();
    var updateCols= getAllIndexes(headers[0], "Student ID");
    var timeStampCols= getAllIndexes(headers[0], "Time");
    if (row > 1) {
      for (var i = 0; i < updateCols.length; i++) {      
        var dateCol = timeStampCols[i];
        var updateCol = updateCols[i] + 1; 
        if (dateCol > -1 && editColumn == updateCol) {
          var cell = sheet.getRange(row, dateCol + 1); 
          var date = Utilities.formatDate(new Date(), timezone, timestamp_format); 
          cell.setValue(date); 
        }
      }
    }
  }
}


function getAllIndexes(arr, val) {
    var indexes = [], i;
    for(i = 0; i < arr.length; i++)
        if (arr[i] === val)
            indexes.push(i);
    return indexes;
}
