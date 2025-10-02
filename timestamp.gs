function onEdit(e){

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const timestamp = new Date();
  const username = Session.getEffectiveUser().getEmail();

  if(row > 1 && e.source.getActiveSheet().getName() == "Sheet1"){
    e.source.getActiveSheet().getRange(row, 8).setValue(timestamp).setNumberFormat('yyyy-MM-dd hh:mm:ss')
    e.source.getActiveSheet().getRange(row, 7).setValue(username)
    e.source.getActiveSheet().getRange(row, 1).setValue(row - 1)

    if(e.source.getActiveSheet().getRange(row, 6).getValue() == ""){
      e.source.getActiveSheet().getRange(row, 6).setValue(timestamp).setNumberFormat('yyyy-MM-dd hh:mm:ss')
      e.source.getActiveSheet().getRange(row, 5).setValue(username)
    }
  }
}
