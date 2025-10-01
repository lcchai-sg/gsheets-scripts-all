// graduating students.js

// ------- This is the new version for activating with clickable button (using macros) -------
function updateSheets() {

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var graduatingStudentsSheet = ss.getSheetByName('Graduating Students');
  
  // Get the 'Form Responses' and 'Slide Submissions' sheets by their IDs
  var formResponsesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RSVP Responses');
  var slideSubmissionsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Slide Submissions');

  // Add activation notification
  graduatingStudentsSheet.getRange('C1:F1').setValue('Please wait, update in progress');
  SpreadsheetApp.getActive().toast("Please wait, updating list with new responses");

  var formResponsesData = formResponsesSheet.getDataRange().getValues();
  var slideSubmissionsData = slideSubmissionsSheet.getDataRange().getValues();

  for (var i = 2; i < formResponsesData.length; i++) {
    var id = formResponsesData[i][4]; // Column E
    var tickets = formResponsesData[i][6]; // Column G

    // Find the rows in 'Graduating Students' that match the ID
    var graduatingStudentsData = graduatingStudentsSheet.getDataRange().getValues();
    for (var j = 2; j < graduatingStudentsData.length; j++) {
      if (graduatingStudentsData[j][2] === id) { // Column C
        // Copy the tickets value to column L
        graduatingStudentsSheet.getRange(j + 1, 12).setValue(tickets); // Column L

        // If tickets > 0, set column K to true
        if (tickets > 0) {
          graduatingStudentsSheet.getRange(j + 1, 11).setValue(true); // Column K
        }

        // Set column J to true
        graduatingStudentsSheet.getRange(j + 1, 10).setValue(true); // Column J
      }
    }
  }

  for (var i = 2; i < slideSubmissionsData.length; i++) {
    if (slideSubmissionsData[i][14] !== '') { // Column O
      var id = slideSubmissionsData[i][7]; // Column H

      // Find the rows in 'Graduating Students' that match the ID
      var graduatingStudentsData = graduatingStudentsSheet.getDataRange().getValues();
      for (var j = 2; j < graduatingStudentsData.length; j++) {
        if (graduatingStudentsData[j][2] === id) { // Column C
          // Set column M to true
          graduatingStudentsSheet.getRange(j + 1, 13).setValue(true); // Column M
        }
      }
    }
  }

  // Add completion notification
  SpreadsheetApp.getActive().toast("Update Complete");

  // adding timestamp for last update
  graduatingStudentsSheet.getRange('C1:F1').setValue('Last updated: ' + Utilities.formatDate(new Date(), "GMT-7", "MM/dd HH:mm"));
}

//=============================================================================================================================================================================================
// graduation to do list.js

// Adding a Timestamp to New Tasks

function addTimestamp(event){ 
  var sheet = event.source.getActiveSheet();
  var range = event.source.getActiveRange();

  if (event.oldValue == null && sheet.getSheetValues(1, range.getColumn(), 1, 1)[0] == "Remaining Tasks") {
    
    // Adding static date to date assigned
    range.offset(0,-2).setValue(new Date());

    // Adding static date to deadline (as default, and to force user to insert Date Assigned if necessary)
    range.offset(0,1).setValue(new Date());

    // Adding default progress to "New"
    range.offset(0,2).setValue("New");
  }
}


// Completed Tasks Actions

function migratingTasks(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;


  // Migrating to Completed Tasks

  // Check if the edit is in the 'Remaining Tasks' tab (Status)
  if (sheet.getName() == "Remaining Tasks" && e.value == "Completed") {

    // Get the entire row
    var sourceRow = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn());

    // Switch to "Completed Tasks" tab
    var archiveSheet = e.source.getSheetByName("Completed Tasks");

    // Find the row index where "Date Assigned" is in column A, row 1 on "Completed Tasks"
    var sigmaRowIndex = archiveSheet.getRange("A1:A").createTextFinder("Date Assigned").findNext().getRow();

    // Insert a new row below "sigma"
    archiveSheet.insertRowAfter(sigmaRowIndex);

    // Copy the formatting and values into the new row below "Date Assigned"
    sourceRow.copyTo(archiveSheet.getRange(sigmaRowIndex + 1, 1), { formatOnly: true });
    sourceRow.copyTo(archiveSheet.getRange(sigmaRowIndex + 1, 1));

    // Set the fill color to #ffffff (white)
    //archiveSheet.getRange(sigmaRowIndex + 1, 1, 1, sourceRow.getNumColumns()).setBackground("#ffffff");

    // Set row height to 21
    archiveSheet.setRowHeight(sigmaRowIndex + 1, 21);

    // Delete the original row from the "Remaining Tasks" tab
    sheet.deleteRow(range.getRow());
  } 

  // Moving back to 'Remaining Tasks' if marked as 'On Going' or 'In Progress' in Completed Tasks
  if (sheet.getName() == "Completed Tasks" && (e.value == "On Going" || e.value == "In Progress" || e.value == "Pending Review")) {

    // Get the entire row
    var sourceRow = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn());

    // Switch to "To Dos - Completed Tasks" tab
    var doingTasksSheet = e.source.getSheetByName("Remaining Tasks");

    // Find the row index where "Example Task" is in column C, row 2 on "Remaining Tasks" tab
    var alphaRowIndex = doingTasksSheet.getRange("C2:C").createTextFinder("Example Task").findNext().getRow();

    // Insert a new row below "beta"
    doingTasksSheet.insertRowAfter(alphaRowIndex);

    // Copy the formatting and values into the new row below "Example Task"
    sourceRow.copyTo(doingTasksSheet.getRange(alphaRowIndex + 1, 1), { formatOnly: true });
    sourceRow.copyTo(doingTasksSheet.getRange(alphaRowIndex + 1, 1));

    // Set the fill color to #ffffff for columns C to J
    // doingTasksSheet.getRange(alphaRowIndex + 1, 3, 1, sourceRow.getNumColumns() - 2).setBackground("#ffffff");

    // Set the row height to 21
    doingTasksSheet.setRowHeight(alphaRowIndex + 1, 21);

    // Delete the original row from the "Archived Tasks" tab
    sheet.deleteRow(range.getRow());
  }
}

//=============================================================================================================================================================================================
// to do list.js

// Move Completed Rows (if already in To Dos List, scroll for code if task was archived) + Add Fill Color

function onEdit1(e) {
  var src = SpreadsheetApp.getActiveSheet();
  const r = e.range;

  if(r.getColumn() == 6 && e.value == "Complete") {

    var data = src.getDataRange().getValues();
    for(var i=1; i<100; i++) {
      if(data[i][0] == "beta"){ 
      var completedrow = i + 1;
      break;
      }
    }

    current = src.getRange(r.getRow(), 1, 1, src.getLastColumn());
    src.moveRows(current, completedrow + 1);  

    // Set the fill color for the entire row starting from column C
    var lastColumn = src.getLastColumn();
    src.getRange(completedrow, 3, 1, lastColumn - 4).setBackground("#fff2cc");

  } 
}



// Insert Blank Row For Next New Task
// need to change it so it checks for null values (new inserts) and doesn't add a new one if there's one already
// might also adjust it so it keeps a new insert blank row at the top of the list? 
// need to fix that with moving tasks back and forth between archive/to dos

function insertBlankRow() {
  // Get the active spreadsheet
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Check if the active sheet is "To Dos"
  if (sheet.getName() == "To Dos") {
    // Find the last row in the sheet
    var lastRow = sheet.getLastRow();

    // Find the row index where column A = "alpha"
    var alphaRowIndex = 6;
    for (var i = lastRow; i > 6; i--) { // not needed for this situation, but accounts for shifting indexes when rows above are added/removed
      if (sheet.getRange(i, 1).getValue() == "alpha") {
        alphaRowIndex = i;
        break;
      }
    }

    // Check if the row below "alpha" has text in column E
    var valueBelowAlpha = sheet.getRange(alphaRowIndex + 1, 5).getValue();

    // If not empty, insert a row and set values in columns D and F
    if (valueBelowAlpha !== "") { // or maybe !== null ?
      sheet.insertRowBefore(alphaRowIndex+1);

      // // Set default drop-down values (Optional)
      // sheet.getRange(alphaRowIndex+1, 4).setValue("Low");
      sheet.getRange(alphaRowIndex+1, 6).setValue("New");
    }
  }
}




// Archived Tasks Actions

function onEdit2(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;


  // Migrating to Archived Tasks

  // Check if the edit is in the "To Dos" tab and in column F (Progress)
  if (sheet.getName() == "To Dos" && range.getColumn() == 6 && e.value == "Archive") {

    // Get the entire row
    var sourceRow = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn());

    // Switch to "Archived Tasks" tab
    var archiveSheet = e.source.getSheetByName("Archived Tasks");

    // Find the row index where "sigma" is in column A, row 3 on "Archived Tasks"
    var sigmaRowIndex = archiveSheet.getRange("A3:A").createTextFinder("sigma").findNext().getRow();

    // Insert a new row below "sigma"
    archiveSheet.insertRowAfter(sigmaRowIndex);

    // Copy the formatting and values into the new row below "sigma"
    sourceRow.copyTo(archiveSheet.getRange(sigmaRowIndex + 1, 1), { formatOnly: true });
    sourceRow.copyTo(archiveSheet.getRange(sigmaRowIndex + 1, 1));

    // Set the fill color to #ffffff (white)
    archiveSheet.getRange(sigmaRowIndex + 1, 1, 1, sourceRow.getNumColumns()).setBackground("#ffffff");

    // Delete the original row from the "To Dos" tab
    sheet.deleteRow(range.getRow());
  } 
  
  
  // Moving to "To Dos - Completed Tasks" if marked as "Complete" in Archived Tasks
  if (sheet.getName() == "Archived Tasks" && range.getColumn() == 6 && e.value == "Complete") {

    // Get the entire row
    var sourceRow = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn());

    // Switch to "To Dos - Completed Tasks" tab
    var completedTasksSheet = e.source.getSheetByName("To Dos");

    // Find the row index where "beta" is in column A, row 3 on "To Dos - Completed Tasks" tab
    var betaRowIndex = completedTasksSheet.getRange("A3:A").createTextFinder("beta").findNext().getRow();

    // Insert a new row below "beta"
    completedTasksSheet.insertRowAfter(betaRowIndex);

    // Copy the formatting and values into the new row below "beta"
    sourceRow.copyTo(completedTasksSheet.getRange(betaRowIndex + 1, 1), { formatOnly: true });
    sourceRow.copyTo(completedTasksSheet.getRange(betaRowIndex + 1, 1));

    // Set the fill color to #fff2cc (light yellow 3) for columns C to J
    completedTasksSheet.getRange(betaRowIndex + 1, 3, 1, sourceRow.getNumColumns() - 2).setBackground("#fff2cc");

    // Set the row height to 21
    completedTasksSheet.setRowHeight(betaRowIndex + 1, 21);

    // Delete the original row from the "Archived Tasks" tab
    sheet.deleteRow(range.getRow());
  }


  // Moving back to list of To Dos if marked as "In Motion" in Archived Tasks
  else if (sheet.getName() == "Archived Tasks" && range.getColumn() == 6 && e.value == "In Motion") {

    // Get the entire row
    var sourceRow = sheet.getRange(range.getRow(), 1, 1, sheet.getLastColumn());

    // Switch to "To Dos - Completed Tasks" tab
    var completedTasksSheet = e.source.getSheetByName("To Dos");

    // Find the row index where "beta" is in column A, row 3 on "To Dos - Completed Tasks" tab
    var alphaRowIndex = completedTasksSheet.getRange("A3:A").createTextFinder("alpha").findNext().getRow();

    // Insert a new row below "beta"
    completedTasksSheet.insertRowAfter(alphaRowIndex);

    // Copy the formatting and values into the new row below "beta"
    sourceRow.copyTo(completedTasksSheet.getRange(alphaRowIndex + 1, 1), { formatOnly: true });
    sourceRow.copyTo(completedTasksSheet.getRange(alphaRowIndex + 1, 1));

    // Set the fill color to #fff2cc (light yellow 3) for columns C to J
    completedTasksSheet.getRange((alphaRowIndex + 1), 3, 1, sourceRow.getNumColumns() - 2).setBackground("#ffffff");

    // Set the row height to 21
    completedTasksSheet.setRowHeight((alphaRowIndex + 1), 21);

    // Delete the original row from the "Archived Tasks" tab
    sheet.deleteRow(range.getRow());
  }
}

//=============================================================================================================================================================================================

