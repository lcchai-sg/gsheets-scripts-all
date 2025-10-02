// menu for testing script
function onOpen() {
  var ui = SpreadsheetApp.getUi()
  ui.createMenu("Daily BTC").addItem("Take Snapshot", "recordValues").addToUi()
}

// record history from a cell and append to next available row
function recordValues() {
  // force the spreadsheet to update
  SpreadsheetApp.flush()

  // get sheets and record date
  var summary = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Summary")
  var dailyA = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily BTC Assets")
  var dailyD = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Daily BTC Debts")
  var date = new Date()

  // get values from named ranges
  var assets = summary.getRange('assets').getValues()
  var debts = summary.getRange('debts').getValues()

  // calculated values
  var sumA = 0
  var sumD = 0
  var net

  for (var i = 0; i < assets.length; i++) {
    for (var j = 0; j < assets[i].length; j++) {
      sumA += assets[i][j]
    }
  }

  for (var i = 0; i < debts.length; i++) {
    for (var j = 0; j < debts[i].length; j++) {
      sumD += debts[i][j]
    }
  }

  net = sumA - sumD

  // record asset values
  var assetRow = [date]
  for (var i = 0; i < assets.length; i++) {
    for (var j = 0; j < assets[i].length; j++) {
      assetRow.push(assets[i][j])
    }
  }
  assetRow.push((Number(sumA).toFixed(2)))
  dailyA.appendRow(assetRow)

  // record debt values
  var debtRow = [date]
  for (var i = 0; i < debts.length; i++) {
    for (var j = 0; j < debts[i].length; j++) {
      debtRow.push(debts[i][j])
    }
  }
  debtRow.push((Number(sumD).toFixed(2)))
  debtRow.push((Number(net).toFixed(2)))
  dailyD.appendRow(debtRow)
}
