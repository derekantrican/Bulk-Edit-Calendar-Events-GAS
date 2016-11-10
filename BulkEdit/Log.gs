function clearLog() {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log").getRange("A1").setValue("");
}

function appendToLog(addThis){
  var range = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log").getRange("A1");
  var value = range.getValue();
  var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "[MM/dd/YYYY HH:mm:ss ");
  timestamp += Session.getScriptTimeZone() + "]";
  
  if (value == "")
    value += timestamp + " " + addThis;
  else
    value += "\n" + timestamp + " " + addThis;
    
  range.setValue(value);
}
