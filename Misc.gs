function install(){
  SpreadsheetApp.getActive().getSheets();
  CalendarApp.getAllCalendars();
  ScriptApp.newTrigger(onOpen).forSpreadsheet(SpreadsheetApp.getActive()).onOpen();
}

function error(problem){
  Browser.msgBox(problem + " [TERMINATING SCRIPT....]");
  return;
}

function onOpen() {  
  var menu = [    
    {name: "Install", functionName: "install"},
    {name: "Run", functionName: "bulkEdit"},
    {name: "Reset Edit Fields", functionName: "resetFields"},
    {name: "Clear Log", functionName: "clearLog"},
    {name: "Check for Update", functionName: "checkUpdate"}
  ];  
  SpreadsheetApp.getActiveSpreadsheet()
    .addMenu("Program Functions", menu);
}

function checkUpdate() {
  var versionToCheck = SpreadsheetApp.openById("1YRylapVSejmJJiwzqwFxJvNK7_E__gSM8isaIf-XouU").getSheetByName("Welcome!").getRange("F3").getValue();
  var currentVersion = SpreadsheetApp.getActive().getSheetByName("Welcome!").getRange("F3").getValue();
  
  if (versionToCheck <= currentVersion)
    Browser.msgBox("You have the most current version. Congratulations!");
  else
    Browser.msgBox("There is a more current version available! The most recent version is v" + versionToCheck +" . Here is the link to download it: http://goo.gl/7xhuUw");
}

function resetFields() {
  var sheet = SpreadsheetApp.getActive().getSheetByName("Home");
  
  //Set required fields blank
  sheet.getRange("C5").setValue("");
  sheet.getRange("C6").setValue("");
  sheet.getRange("C7").setValue("");
  
  //Set optional fields "No"
  sheet.getRange("C9").setValue("No");
  sheet.getRange("C10").setValue("No");
  sheet.getRange("C11").setValue("No");
  sheet.getRange("C12").setValue("No");
  sheet.getRange("C13").setValue("No");
  sheet.getRange("C14").setValue("No");
  sheet.getRange("C15").setValue("No");
  sheet.getRange("C16").setValue("No");
  sheet.getRange("E16").setValue("No");
  sheet.getRange("C17").setValue("No");
  sheet.getRange("C22").setValue("No");
  
  //Set keyword/phrase blank
  sheet.getRange("C21").setValue("");
  
  //Set dependent fields blank
  sheet.getRange("E9").setValue("");
  sheet.getRange("E10").setValue("");
  sheet.getRange("E11").setValue("");
  sheet.getRange("E12").setValue("");
  sheet.getRange("E13").setValue("");
  sheet.getRange("E14").setValue("");
  sheet.getRange("E15").setValue("");
  sheet.getRange("E17").setValue("");
  sheet.getRange("E21").setValue("");
}
