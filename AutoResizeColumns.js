//This script with automatically resize columns on the active sheet, starting from cell A1 until the last cell with data in row 1
function autoResizeColumns() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastColumn = sheet.getLastColumn();
  
    for (var i = 1; i <= lastColumn; i++) {
      sheet.autoResizeColumn(i);
    }
  }
