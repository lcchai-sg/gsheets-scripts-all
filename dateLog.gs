function onEdit(e){
  var sheet = e.source.getActiveSheet();
  
  if (sheet.getName() == "Sheet1"){ // name of the sheet
    var actRng = sheet.getActiveRange();
    var editColumn = actRng.getColumn();
    var rowIndex = actRng.getRowIndex();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    var dateCol = headers[0].indexOf("DATE") + 1; // column header for insert timestamp
    var orderCol = headers[0].indexOf("ITEM") + 1; // column header for cell values to be filled
    
    if (dateCol > 0 && rowIndex > 1 && editColumn == orderCol){
      // optional if statement to check if the cell is empty to provend accidental overwrite
      if(sheet.getRange(rowIndex, dateCol).isBlank(){
        sheet.getRange(rowIndex, dateCol).setValue(Utilities.formatDate(new Date(), "UTC-8", "MM/dd/yyyy"));
      }
    }
  }
}
