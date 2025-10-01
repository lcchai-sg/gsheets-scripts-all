var sheet = SpreadsheetApp.getActive().getSheetByName('Complete Accounting');
var sheetBeforeNV = SpreadsheetApp.getActive().getSheetByName('Before NV cleared');

function ClearCells() {

  
  //Machine
  sheet.getRange('C19:C20').clearContent();
  sheet.getRange('C31:C36').clearContent();
  sheet.getRange('C39:C41').clearContent();
  sheet.getRange('C43:C46').clearContent();
  sheet.getRange('C54:C57').clearContent();
  sheet.getRange('C59:C62').clearContent();
  sheet.getRange('C64:C66').clearContent();
  sheet.getRange('B71:B79').clearContent();
  sheet.getRange('B81:B89').clearContent();
  sheet.getRange('B91:B98').clearContent();
  sheet.getRange('B100:B107').clearContent();

  //CSV
  sheet.getRange('G17:G18').clearContent();
  sheet.getRange('G21:G24').clearContent();
  sheet.getRange('G27:G38').clearContent();
  sheet.getRange('G43:G64').clearContent();

  //Clear 'Before NV Cleared' page
  sheetBeforeNV.getRange('C19:C20').clearContent();
  sheetBeforeNV.getRange('C31:C36').clearContent();
  sheetBeforeNV.getRange('C39:C41').clearContent();
  sheetBeforeNV.getRange('C43:C46').clearContent();
  sheetBeforeNV.getRange('C54:C57').clearContent();
  sheetBeforeNV.getRange('C59:C62').clearContent();
  sheetBeforeNV.getRange('C6:C66').clearContent();

  sheetBeforeNV.getRange('B71:B79').clearContent();
  sheetBeforeNV.getRange('B81:B89').clearContent();
  sheetBeforeNV.getRange('B91:B98').clearContent();
  sheetBeforeNV.getRange('B100:B107').clearContent();
}

function SetZerosJustCsv(){
  sheet.getRange('G17:G18').setValue(0);
  sheet.getRange('G21:G24').setValue(0);
  sheet.getRange('G27:G38').setValue(0);
  sheet.getRange('G43:G64').setValue(0);
}

function SetZerosJustMachine(){
  //Clear 'Before NV Cleared' page
  sheet.getRange('C19:C20').setValue(0);
  sheet.getRange('C31:C36').setValue(0);
  sheet.getRange('C39:C41').setValue(0);
  sheet.getRange('C43:C46').setValue(0);
  sheet.getRange('C54:C57').setValue(0);
  sheet.getRange('C59:C62').setValue(0);
  sheet.getRange('C64:C66').setValue(0);
  sheet.getRange('B71:B79').clearContent();
  sheet.getRange('B81:B89').clearContent();
  sheet.getRange('B91:B98').clearContent();
  sheet.getRange('B100:B107').clearContent();
}

function ReplaceWithZero() {
  //'Complete Accounting' page
    //Machine column
  sheet.getRange('C19:C20').setValue(0);
  sheet.getRange('C31:C36').setValue(0);
  sheet.getRange('C39:C41').setValue(0);
  sheet.getRange('C43:C46').setValue(0);
  sheet.getRange('C54:C57').setValue(0);
  sheet.getRange('C59:C62').setValue(0);
  sheet.getRange('C64:C66').setValue(0);
  sheet.getRange('B71:B79').clearContent();
  sheet.getRange('B81:B89').clearContent();
  sheet.getRange('B91:B98').clearContent();
  sheet.getRange('B100:B107').clearContent();;

    //Csv column
  sheet.getRange('G17:G18').setValue(0);
  sheet.getRange('G21:G24').setValue(0);
  sheet.getRange('G27:G38').setValue(0);
  sheet.getRange('G43:G64').setValue(0);


  //Clear 'Before NV Cleared' page
  sheetBeforeNV.getRange('C19:C20').setValue(0);
  sheetBeforeNV.getRange('C31:C36').setValue(0);
  sheetBeforeNV.getRange('C39:C41').setValue(0);
  sheetBeforeNV.getRange('C43:C46').setValue(0);
  sheetBeforeNV.getRange('C54:C57').setValue(0);
  sheetBeforeNV.getRange('C59:C62').setValue(0);
  sheetBeforeNV.getRange('C6:C66').setValue(0);
  sheet.getRange('B71:B79').clearContent();
  sheet.getRange('B81:B89').clearContent();
  sheet.getRange('B91:B98').clearContent();
  sheet.getRange('B100:B107').clearContent();
}
