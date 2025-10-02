function cp_below(){
  
  function current_cell(){
  var currentCell = SpreadsheetApp.getActive();
  return currentCell;
  };

  function current_range(){
    var currentrange = SpreadsheetApp.getActive().getActiveRange();
    return currentrange;
  };
  
  function copyandpaste() {
    current_cell().getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
    parameter1a = current_range().getNumRows();
    parameter1b = current_range().getNumColumns();
    current_cell().getSelection().getNextDataRange(SpreadsheetApp.Direction.DOWN).activate();
    parameter2a = current_range().getNumRows()-1;
    parameter2b = current_range().getNumColumns();
    current_cell().getCurrentCell().offset(0, 0,parameter2a,parameter2b).activate();
    current_cell().getCurrentCell().offset(0, 0,parameter1a, parameter1b).copyTo(current_range(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  };

  for (var i=0; i<10;i++){
    copyandpaste();
    current_cell().getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  };
};
