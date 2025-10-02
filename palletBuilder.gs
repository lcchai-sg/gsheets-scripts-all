var ss = SpreadsheetApp.getActiveSpreadsheet();
var exportSheet = ss.getSheets()[5];
var fileId = '1vZ-Jv6NiC3G8f-TZfjl8pF9meTr3X7Ep'; // palletBuilder.txt
var txtFile = DriveApp.getFileById(fileId);
var allFileId = '1kay94-NXfV4x9TPwUv3jBPdDpVZe4Wfm'; // allPallets.txt
var allTxtFile = DriveApp.getFileById(allFileId);

// adds a function drop down menu
// set a trigger to run on open
function onOpen() {
   var ss = SpreadsheetApp.getActiveSpreadsheet();
   var menubuttons = [{name: "Export Red", functionName: "red"},
                      {name: "Export Yellow", functionName: "yellow"},
                      {name: "Export Green", functionName: "green"},
                      {name: "Export Blue", functionName: "blue"},
                      {name: "Export Purple", functionName: "purple"}];
   ss.addMenu("Functions", menubuttons);
}

function red() {clearAndExport("A",0);}
function yellow() {clearAndExport("B",1);}
function green() {clearAndExport("C",2);}
function blue() {clearAndExport("D",3);}
function purple() {clearAndExport("E",4);}

function clearAndExport(col, num){
  var values = exportSheet.getRange(col + "1:" + col + "253").getValues();
  var strValues = "Pallet " + values[0] + "\n";
  for (var i = 1; i < 253; i++) {
    if (values[i] != "") {
      strValues = strValues + values[i] + "\n"
    }
  }
  appendTXT(strValues);
  var sheet = ss.getSheets()[num];
  sheet.getRange("A1:A252").clearContent();
}

function appendTXT(content){
  var currentFileContent = "";
  var allCurrentFileContent = "";
  var stringContent = "";
  var delimiter = ";";
  stringContent = stringContent || content || "\n";
  if(stringContent){
    // get file's current text content
    currentFileContent = txtFile.getBlob().getDataAsString();
    allCurrentFileContent = allTxtFile.getBlob().getDataAsString();
    // update the txt file with its previous content
    txtFile.setContent(currentFileContent + stringContent);
    allTxtFile.setContent(allCurrentFileContent + stringContent);
  }
}

function clearTXT() {
  txtFile.setContent("");
}

// set a trigger to run on open
function renameSheets() {
  for (var i=0; i<5; i++) {
    var sheet = ss.getSheets()[i];
    var cell = sheet.getRange("C1");   
    var value = cell.getValue();
    sheet.setName(value);
  }
}
