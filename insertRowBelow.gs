/** @OnlyCurrentDoc */
function prompt() {
  var i = SpreadsheetApp.getUi().prompt("Enter rows amount to insert below").getResponseText();
  return i
};
function InsertBelow() {
  var Spread_Sheet = SpreadsheetApp.getActive();
  Spread_Sheet.getActiveSheet().insertRowsAfter(Spread_Sheet.getCurrentCell().getRow(), prompt()); 
};
//Replace promt() with any number for fixed value
