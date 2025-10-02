// A function that adds an up or down arrow to a cell 
// depending on if its higher or lower than the cell in the 
// row previous.
function addChangeArrow() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  var startingRow = 3;
  var startingColumn = 2
  
  for (var i = startingRow; i < data.length; i++) {
    for (var j = startingColumn; j < data[i].length; j++){
      // If the current cell is empty ignore it.
      if(data[i][j] === "") {continue;}
      
      var currentValue = data[i][j].toString();
      var previousValue = data[i-1][j].toString();
      
      // Remove any previous arrows that might be in the cell.
      if(isNaN(currentValue)){ currentValue = currentValue.substring(1);} 
      if(isNaN(previousValue)){ previousValue = previousValue.substring(1);}
      
      // If the values still aren't a number skip this cell.
      if(isNaN(currentValue)) {continue;}
      if(isNaN(previousValue)) {continue;}
      
      // Get the number values, so we're not comparing strings.
      currentValue = parseFloat(currentValue, 10);
      previousValue = parseFloat(previousValue, 10);
      
      // Decide whether to use the up arrow or down arrow.
      if( currentValue > previousValue ) {
        sheet.getRange(i+1,j+1).setValue("▲" + currentValue);
        continue;
      }
      
      if( currentValue < previousValue ) {
        sheet.getRange(i+1,j+1).setValue("▼" + currentValue);
        continue;
      }
    }
  }

}
