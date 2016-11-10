function renameEvent(calendarID, event, renameTo){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var oldTitle = event.getTitle();
  event.setTitle(renameTo);
  
  appendToLog('"' + oldTitle + "\" was renamed to \"" + renameTo + '"');  
}

function changeLocation(calendarID, event, newLoc){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var oldLoc = event.getLocation();
  event.setLocation(newLoc);
  var title = event.getTitle();
  
  appendToLog("The location of \"" + title + "\" was changed from \"" + oldLoc + "\" to \"" + newLoc + '"');
}

function changeDescription(calendarID, event, newDesc){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var title = event.getTitle();
  var oldDesc = event.getDescription();
  event.setDescription(newDesc);
  
  appendToLog("The description of \"" + title + "\" was changed from \"" + oldDesc + "\" to \"" + newDesc + '"');
}

function appendDescription(calendarID, event, newDesc){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var title = event.getTitle();
  var oldDesc = event.getDescription();
  newDesc = oldDesc + " " + newDesc;
  event.setDescription(newDesc);
  
  appendToLog("The description of \"" + title + "\" was changed from \"" + oldDesc + "\" to \"" + newDesc + '"');
}

//Moving the start time is done in the main code because it's harder to do

//Changing the duration is done in the main code because it's harder to do

function addReminders(calendarID, event, reminderType, minutesBefore){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var title = event.getTitle();
    
  if (reminderType == "Email"){
    event.addEmailReminder(minutesBefore);
    appendToLog("An email reminder " + minutesBefore + " was added to \"" + title + '"');
  }
  else if (reminderType == "SMS"){
    event.addSmsReminder(minutesBefore);
    appendToLog("A SMS reminder " + minutesBefore + " was added to \"" + title + '"');    
  }
  else if (reminderType == "Popup"){
    event.addPopupReminder(minutesBefore);
    appendToLog("A Popup reminder " + minutesBefore + " was added to \"" + title + '"');
  }
}

function deleteReminders(calendarID, event, reminderType){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var title = event.getTitle();
  var emailReminders = event.getEmailReminders();
  var smsReminders = event.getSmsReminders();
  var popupReminders = event.getPopupReminders();
  
  if (reminderType == "Email"){
    event.removeAllReminders();
    
    //Add back the other reminders
    for (var i = 0; i < smsReminders.length; i++)
      event.addSmsReminder(smsReminders[i]);
      
    for (var j = 0; j < popupReminders.length; j++)
      event.addPopupReminder(popupReminders[j]);
    
    appendToLog("All the email reminders for \"" + title + "\" were removed.");
  }
  else if (reminderType == "SMS"){
    event.removeAllReminders();
    
    //Add back other reminders
    for (var i = 0; i < emailReminders.length; i++)
      event.addSmsReminder(emailReminders[i]);

    for (var j = 0; j < popupReminders.length; j++)
      event.addPopupReminder(popupReminders[j]);
    
    appendToLog("All the SMS reminders for \"" + title + "\" were removed.");
  }
  else if (reminderType == "Popup"){
    event.removeAllReminders();
    
    //Add back other reminders
    for (var i = 0; i < smsReminders.length; i++)
      event.addSmsReminder(smsReminders[i]);

    for (var j = 0; j < emailReminders.length; j++)
      event.addPopupReminder(emailReminders[j]);
    
    appendToLog("All the popup reminders for \"" + title + "\" were removed.");
  }
  else if (reminderType == "All"){
    event.removeAllReminders();
    appendToLog("All the reminders for \"" + title + "\" were removed.");
  }
}

function deleteEvent(calendarID, event){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var title = event.getTitle();
  event.deleteEvent();
  
  appendToLog('"' + title + "\" was deleted");
}

function keywordMatch(keyword, keywordLoc, event){
  var matchHere = '';

  if (keywordLoc == "Title")
    var match = event.getTitle().split(' ');
  else if (keywordLoc == "Description")
    var match = event.getDescription().split(' ');
  else if (keywordLoc == "Location")
    var match = event.getLocation().split(' ');
  
  for (var i = 0; i < match.length; i++)
    matchHere += match[i];
  
  if(matchHere.search(keyword) >= 0)
    return true;
  else
    return false;
}

function exactMatch(exact, exactLoc, event){
  var matchHere = '';

  if (exactLoc == "Title")
    var match = event.getTitle().split(' ');
  else if (exactLoc == "Description")
    var match = event.getDescription().split(' ');
  else if (exactLoc == "Location")
    var match = event.getLocation().split(' ');
  
  for (var i = 0; i < match.length; i++)
    matchHere += match[i];
  
  if(matchHere == exact)
    return true;
  else
    return false;
}

function changeVisibility(calendarID, event, visibilityTo){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var title = event.getTitle();
  
  if (visibilityTo == "Public")
    visibilityTo = CalendarApp.Visibility.PUBLIC;
  else if (visibilityTo == "Private")
    visibilityTo = CalendarApp.Visibility.PRIVATE;
  else if (visibilityTo == "Confidential")
    visibilityTo = CalendarApp.Visibility.CONFIDENTIAL;
  else
    visibilityTo = CalendarApp.Visibility.DEFAULT;
  
  event.setVisibility(visibilityTo);
  
  appendToLog('The visibility of "' + title + '" was changed to ' + visibilityTo);
}
