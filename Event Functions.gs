function renameEvent(calendarID, eventID, renameTo){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var event = calendar.getEventSeriesById(eventID);
  var oldTitle = event.getTitle();
  event.setTitle(renameTo);
  
  appendToLog('"' + oldTitle + "\" was renamed to \"" + renameTo + '"');  
}

function changeLocation(calendarID, eventID, newLoc){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var event = calendar.getEventSeriesById(eventID);
  var title = event.getTitle();
  var oldLoc = event.getLocation();
  event.setLocation(newLoc);
  
  appendToLog("The location of \"" + title + "\" was changed from \"" + oldLoc + "\" to \"" + newLoc + '"');
}

function changeDescription(calendarID, eventID, newDesc){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var event = calendar.getEventSeriesById(eventID);
  var title = event.getTitle();
  var oldDesc = event.getDescription();
  event.setDescription(newDesc);
  
  appendToLog("The description of \"" + title + "\" was changed from \"" + oldDesc + "\" to \"" + newDesc + '"');
}

function appendDescription(calendarID, eventID, newDesc){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var event = calendar.getEventSeriesById(eventID);
  var title = event.getTitle();
  var oldDesc = event.getDescription();
  newDesc = oldDesc + " " + newDesc;
  event.setDescription(newDesc);
  
  appendToLog("The description of \"" + title + "\" was changed from \"" + oldDesc + "\" to \"" + newDesc + '"');
}

//Moving the start time is done in the main code because it's harder to do

//Changing the duration is done in the main code because it's harder to do

function addReminders(calendarID, eventID, reminderType, minutesBefore){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var event = calendar.getEventSeriesById(eventID);
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

function deleteReminders(calendarID, eventID, reminderType){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var event = calendar.getEventSeriesById(eventID);
  var title = event.getTitle();
  var emailReminders = event.getEmailReminders();
  var smsReminders = event.getSmsReminders();
  var popupReminders = event.getPopupReminders();
  
  if (reminderType == "Email"){
    event.removeAllReminders();
    
    //Add back the other reminders
    for (var i = 0; i < smsReminders.length; i++){
      event.addSmsReminder(smsReminders[i]);
    }
    for (var j = 0; j < popupReminders.length; j++){
      event.addPopupReminder(popupReminders[j]);
    }
    
    appendToLog("All the email reminders for \"" + title + "\" were removed.");
  }
  else if (reminderType == "SMS"){
    event.removeAllReminders();
    
    //Add back other reminders
    for (var i = 0; i < emailReminders.length; i++){
      event.addSmsReminder(emailReminders[i]);
    }
    for (var j = 0; j < popupReminders.length; j++){
      event.addPopupReminder(popupReminders[j]);
    }
    
    appendToLog("All the SMS reminders for \"" + title + "\" were removed.");
  }
  else if (reminderType == "Popup"){
    event.removeAllReminders();
    
    //Add back other reminders
    for (var i = 0; i < smsReminders.length; i++){
      event.addSmsReminder(smsReminders[i]);
    }
    for (var j = 0; j < emailReminders.length; j++){
      event.addPopupReminder(emailReminders[j]);
    }
    
    appendToLog("All the popup reminders for \"" + title + "\" were removed.");
  }
  else if (reminderType == "All"){
    event.removeAllReminders();
    appendToLog("All the reminders for \"" + title + "\" were removed.");
  }
}

function deleteEvent(calendarID, eventID){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var event = calendar.getEventSeriesById(eventID);
  var title = event.getTitle();
  event.deleteEventSeries();
  
  appendToLog('"' + title + "\" was deleted");
}

function alternativeDeleteEvent(calendarID, event){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var title = event.getTitle();
  event.deleteEvent();
  
  appendToLog('"' + title + "\" was deleted");
}

function keywordMatch(keyword,keywordLoc,event){
  if (keywordLoc == "Title"){
    if(event.getTitle().search(keyword) >= 0)
      return true;
    else
      return false;
  }
  else if (keywordLoc == "Description"){
    if(event.getDescription().search(keyword) >= 0)
      return true;
    else
      return false;
  }
  else if (keywordLoc == "Location"){
    if(event.getLocation().search(keyword) >= 0)
      return true;
    else
      return false;
  }
}

function changeVisibility(calendarID, eventID, visibility){
  var calendar = CalendarApp.getCalendarById(calendarID);
  var event = calendar.getEventSeriesById(eventID);
  var title = event.getTitle();
  
  if (visibility == "Public")
    visibility = CalendarApp.Visibility.PUBLIC;
  else if (visibility == "Private")
    visibility = CalendarApp.Visibility.PRIVATE;
  else if (visibility == "Confidential")
    visibility = CalendarApp.Visibility.CONFIDENTIAL;
  else
    visibility = CalendarApp.Visibility.DEFAULT;
  
  event.setVisibility(visibility);
  
  appendToLog('The visibility of "' + title + '" was changed to ' + visibility);
}
