function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Employee Clock')
    .addItem('Clock In', 'clockIn')
    .addItem('Clock Out', 'clockOut')
    .addToUi();

  // Clear all protections on the sheet
  clearAllProtections();
  
  // Protect columns F to K
  protectColumnsFK();

  // Protect columns D and E for script editing
  protectColumnsDE();
}

function clearAllProtections() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Get all protections on the sheet and remove them
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  
  protections.forEach(function(protection) {
    protection.remove();
  });
}

function protectColumnsFK() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Protect columns F to K
  var rangeFK = sheet.getRange('F:K');
  var protectionFK = rangeFK.protect().setDescription('Protect columns F to K from manual edits.');

  // Set the protection to allow only admins to edit
  var userEmail = Session.getEffectiveUser().getEmail();
  protectionFK.removeEditors(protectionFK.getEditors());
  protectionFK.addEditor('admin@maxvilleheritage.org'); // Add admin
  
  protectionFK.setWarningOnly(false); // Fully lock the range
}

function protectColumnsDE() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Protect column D (Clock In)
  var rangeD = sheet.getRange('D:D'); 
  var protectionD = rangeD.protect().setDescription('Protect column D from manual edits.');
  
  // Protect column E (Clock Out)
  var rangeE = sheet.getRange('E:E');
  var protectionE = rangeE.protect().setDescription('Protect column E from manual edits.');

  // Remove all editors except the script runner
  var userEmail = Session.getEffectiveUser().getEmail(); // Get the email of the user running the script
  removeAllEditors(protectionD, userEmail);
  removeAllEditors(protectionE, userEmail);
  
  // Add admin email if necessary
  protectionD.addEditor('admin@maxvilleheritage.org');
  protectionE.addEditor('admin@maxvilleheritage.org');

  protectionD.setWarningOnly(false); // Fully lock the range
  protectionE.setWarningOnly(false);
}

function removeAllEditors(protection, userEmail) {
  if (protection) {
    protection.removeEditors(protection.getEditors());
    protection.addEditor(userEmail); // Allow the script owner
    var me = Session.getEffectiveUser();
    protection.removeEditor(me); // Ensure the current user is removed after adding the owner
    protection.setWarningOnly(false);
  }
}

function clockIn() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM d, yyyy');

  var data = sheet.getRange('A:A').getValues();
  for (var i = 0; i < data.length; i++) {
    var cellDateString = data[i][0];
    
    if (cellDateString !== "Date") {
      var cellDateFormatted = Utilities.formatDate(cellDateString, Session.getScriptTimeZone(), 'MMMM d, yyyy');

      if (cellDateFormatted === dateString) {
        var clockInCell = sheet.getRange(i + 1, 4); // Column D for Clock In
        
        Logger.log("Clock In Cell: " + clockInCell.getA1Notation()); // Log the A1 notation

        if (clockInCell.isBlank()) {
          unlockCellForUser(clockInCell); // Temporarily unlock the cell for the user
          SpreadsheetApp.flush(); // Ensure all pending changes are applied
          Utilities.sleep(500); // Delay to ensure the cell is unlocked
          clockInCell.setValue(today); // Set clock in time
          lockCell(clockInCell); // Lock the cell after editing
        } else {
          SpreadsheetApp.getUi().alert("Clock In time has already been recorded for today.");
        }
        return;
      }
    }
  }
  SpreadsheetApp.getUi().alert("Today's date not found in column A.");
}

function clockOut() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var today = new Date();
  var dateString = Utilities.formatDate(today, Session.getScriptTimeZone(), 'MMMM d, yyyy');

  var data = sheet.getRange('A:A').getValues();
  for (var i = 0; i < data.length; i++) {
    var cellDateString = data[i][0];
    
    if (cellDateString !== "Date") {
      var cellDateFormatted = Utilities.formatDate(cellDateString, Session.getScriptTimeZone(), 'MMMM d, yyyy');

      if (cellDateFormatted === dateString) {
        var clockOutCell = sheet.getRange(i + 1, 5); // Column E for Clock Out
        
        Logger.log("Clock Out Cell: " + clockOutCell.getA1Notation()); // Log the A1 notation

        if (clockOutCell.isBlank()) {
          unlockCellForUser(clockOutCell); // Temporarily unlock the cell for the user
          SpreadsheetApp.flush(); // Ensure all pending changes are applied
          Utilities.sleep(500); // Delay to ensure the cell is unlocked
          clockOutCell.setValue(today); // Set clock out time
          lockCell(clockOutCell); // Lock the cell after editing
        } else {
          SpreadsheetApp.getUi().alert("Clock Out time has already been recorded for today.");
        }
        return;
      }
    }
  }
  SpreadsheetApp.getUi().alert("Today's date not found in column A.");
}

function unlockCellForUser(cell) {
  if (!cell) {
    Logger.log("Error: cell is undefined in unlockCellForUser."); // Log if cell is undefined
    return;
  }
  
  Logger.log("Unlocking cell for user: " + cell.getA1Notation()); // Log the range being unlocked
  
  try {
    var protection = cell.protect().setDescription('Temporary unlock for user');
    var userEmail = Session.getEffectiveUser().getEmail(); // Get the email of the current user
    protection.addEditor(userEmail); // Temporarily grant the current user edit access
    protection.removeEditor('admin@maxvilleheritage.org'); // Ensure only the necessary users are allowed
    
    protection.setWarningOnly(false); // Fully unlock the cell for editing
  } catch (e) {
    Logger.log("Error unlocking cell: " + e.message);
  }
}

function lockCell(cell) {
  if (!cell) {
    Logger.log("Error: cell is undefined in lockCell."); // Log if cell is undefined
    return;
  }
  
  Logger.log("Locking cell: " + cell.getA1Notation()); // Log the range being locked
  
  // Protect the range
  var protection = cell.protect().setDescription('This cell is locked after editing.');
  
  // Allow only the admin to edit the range again after the user has clocked in/out
  protection.addEditor('admin@maxvilleheritage.org'); // Add admin
  
  var userEmail = Session.getEffectiveUser().getEmail(); // Get the current user email
  protection.removeEditor(userEmail); // Remove the current user's temporary access

  protection.setWarningOnly(false); // Fully lock the range after editing
}
