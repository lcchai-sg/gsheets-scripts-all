function myFunction() {
  var id = "Sheets' id here";
  var sheets = SpreadsheetApp.openById(id).getSheetByName("first");
  init()


  var candidates = sheets.getRange(2, 2, 200, 5).getValues()
  for (var r in candidates) {
    Logger.log(r)
    for (var c in candidates[r]) {
      var candidateCode = getCodeMap(candidates[r][c])
      if (undefined == candidateCode) {
        break
      }
      if (isRemain(candidateCode[2])) {
        Logger.log("name: " + r + "志願: " + candidateCode)
        setResult(r, candidateCode)
        minusQuota(candidateCode)
        var candidateCode = getCodeMap(candidates[r][c])
        break
      }
    }
  }


  function isRemain(qouta) {
    if (qouta > 0) {
      return true
    }
  }

  /**  return [code][unitname][qouta]*/
  function getCodeMap(code) {
    var range = sheets.getRange(2, 10, 18, 3)
    var maps = range.getValues()
    for (var r in maps) {
      if (code == maps[r][0]) {
        return maps[r]
      }
    }

  }
  function setResult(row, code) {
    var row = parseInt(row) + 2
    sheets.getRange(row, 7).setValue(code[1])

  }

  function minusQuota(code) {
    // Logger.log("minus:" + code)
    var range = sheets.getRange(2, 10, 18, 3)
    var maps = range.getValues()
    for (var r in maps) {
      for (var c in maps[r]) {
        if (code[0] == maps[r][0]) {
          var rr = parseInt(r) + 2
          var cc = parseInt(c) + 12
          var nowRange = sheets.getRange(rr, cc)
          var newQuota = code[2] - 1
          nowRange.setValue(newQuota)
          Logger.log("minus: " + code + "; remains: " + newQuota)
          break
        }
      }
    }

  }

  function init() {
    var originalValue = sheets.getRange(2, 13, 18).getValues()
    sheets.getRange(2, 12, 18).setValues(originalValue)
    sheets.getRange(2, 7, 200).setValue("無單位")
  }

}
