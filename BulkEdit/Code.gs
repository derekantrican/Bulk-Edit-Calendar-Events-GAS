function bulkEdit(){
  var sheet = SpreadsheetApp.getActive().getSheetByName("Home");
  var log = SpreadsheetApp.getActive().getSheetByName("Log");
  var calendar = sheet.getRange("C5").getValue();
  var startTime = sheet.getRange("C6").getValue();
  var endTime = sheet.getRange("C7").getValue();
  var timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  var rename = sheet.getRange("C9").getValue();
  var renameTo = sheet.getRange("E9").getValue();
  var changeLoc = sheet.getRange("C10").getValue();
  var newLoc = sheet.getRange("E10").getValue();
  var changeDesc = sheet.getRange("C11").getValue();
  var newDesc = sheet.getRange("E11").getValue();
  var moveStart = sheet.getRange("C12").getValue();
  var forwardBack = sheet.getRange("D12").getValue();
  var minutesToMove = sheet.getRange("E12").getValue();
  var changeDur = sheet.getRange("C13").getValue();
  var lengthenShorten = sheet.getRange("D13").getValue();
  var changeDurByMin = sheet.getRange("E13").getValue();
  var changeAbsStart = sheet.getRange("C14").getValue();
  var absStart = sheet.getRange("E14").getValue();
  var changeAbsEnd = sheet.getRange("C15").getValue();
  var absEnd = sheet.getRange("E15").getValue();
  var addRemind = sheet.getRange("C16").getValue();
  var minutesRemind = sheet.getRange("E16").getValue();
  var deleteRemind = sheet.getRange("C17").getValue();
  var deleteReminderType = sheet.getRange("E17").getValue();
  var deleteTheEvent = sheet.getRange("C18").getValue();
  var keyword = sheet.getRange("C23").getValue();
  var keywordLoc = sheet.getRange("E23").getValue();
  var allDays = sheet.getRange("C24").getValue();
  var alternativeDelete = sheet.getRange("E18").getValue();
  var visibility = sheet.getRange("C19").getValue();
  var visibilityTo = sheet.getRange("E19").getValue();
  
  if (checkRequiredFields() == false)
    return;
  
  if (checkDependentFields() == false)
    return;
  
  clearLog(); //Clear the log and start editing events

  var currentCalendar = CalendarApp.getCalendarsByName(calendar)[0];
      
  //If the calendar doesn't exist, throw an error and exit
  if (typeof currentCalendar == 'undefined'){
    error("I can't find the calendar named \"" + calendar + "\". If you're using your default calendar, the name should be your gmail address, not your name.");
    return;
  }
  
  var calendarID = currentCalendar.getId();

  var events = currentCalendar.getEvents(startTime, endTime);
  
  var currentEvent;
  for (var i = 0; i < events.length; i++){
    if (keyword != ""){
      if (keywordMatch(keyword,keywordLoc,events[i]) == true)
        currentEvent = events[i];
      else
        continue;
    }
    else
      currentEvent = events[i];
       
    if (allDays == "Yes"){
      if (currentEvent.isAllDayEvent() == false)
        continue;
    }   
       
    //START EDITING EVENTS
    
    if (deleteTheEvent == "Yes"){ //Delete the event
      if (alternativeDelete == "Yes"){
        alternativeDeleteEvent(calendarID, currentEvent);
        continue; //I do this first because it'd be pointless to modify an event you're going to delete
      }
      else{
        deleteEvent(calendarID,currentEvent.getId());
        continue; //I do this first because it'd be pointless to modify an event you're going to delete
      }
    }
    
    if (rename == "Yes") //Rename the event
      renameEvent(calendarID,currentEvent.getId(),renameTo);
    
    if (changeLoc == "Yes") //Change the event's location
      changeLocation(calendarID,currentEvent.getId(),newLoc);
      
    if (changeDesc == "Yes") //Change the event's description
      changeDescription(calendarID,currentEvent.getId(),newDesc);
    else if (changeDesc == "Yes (append)")
      appendDescription(calendarID,currentEvent.getId(),newDesc);
      
    if (moveStart == "Yes"){ //Change the event's start time
      if (currentEvent.isAllDayEvent() == true){ //If it is an all-day event, throw an error
        if (minutesToMove % 1440 != 0){
          error('"' + currentEvent.getTitle() + "\" is an all-day event and " + minutesToMove + " is not a multiple of 1440 (the # of minutes in a day)");
          continue;
        }
        else{
          if (forwardBack == "by... (minutes future->past)")
            var newAllDayDate = new Date(currentEvent.getAllDayStartDate().getTime() - (minutesToMove * 60 * 1000));
          else
            var newAllDayDate = new Date(currentEvent.getAllDayStartDate().getTime() + (minutesToMove * 60 * 1000));
          
          currentEvent.setAllDayDate(newAllDayDate);
        }
      }
      else{
        if (forwardBack == "by... (minutes future->past)"){ //Move the event backward in time
          var newStart = new Date(currentEvent.getStartTime().getTime() - (minutesToMove * 60 * 1000));
          var newEnd = new Date(currentEvent.getEndTime().getTime() - (minutesToMove * 60 * 1000));
          currentEvent.setTime(newStart,newEnd);
          appendToLog('"' + currentEvent.getTitle() + "\" was moved " + minutesToMove + " closer to the Beginning of All Time");
        }
        else{ //Move the event forward in time
          var newStart = new Date(currentEvent.getStartTime().getTime() + (minutesToMove * 60 * 1000));
          var newEnd = new Date(currentEvent.getEndTime().getTime() + (minutesToMove * 60 * 1000));
          currentEvent.setTime(newStart,newEnd);
          appendToLog('"' + currentEvent.getTitle() + "\" was moved " + minutesToMove + " closer to the End of All Time");
        }
      }
    }
    
    if (changeDur == "Yes"){ //Change the event's duration
      if (currentEvent.isAllDayEvent() == true){ //If it is an all-day event, throw an error
        appendToLog('"' + currentEvent.getTitle() + "\" is an all-day event. You can't change the duration of an all-day event.");
        continue;
      }
      
      if (lengthenShorten == "by... (lengthen minutes)"){ //Lengthen the duration of the event
        var newEnd = new Date(currentEvent.getEndTime().getTime() + (changeDurByMin * 60 * 1000));
        currentEvent.setTime(currentEvent.getStartTime(),newEnd);
        appendToLog('The duration of "' + currentEvent.getTitle() + "\" was lengthened by " + minutesToMove + '" minutes');
      }
      else{ //Shorten the duration of the event
        var newEnd = new Date(currentEvent.getEndTime().getTime() - (changeDurByMin * 60 * 1000));
        currentEvent.setTime(currentEvent.getStartTime(),newEnd);
        appendToLog('The duration of "' + currentEvent.getTitle() + "\" was shortened by " + minutesToMove + '" minutes');
      }
    }
    
    if (changeAbsStart == "Yes") {  // Change absolute beginning time of event
      appendToLog('Editing start time for "' + currentEvent.getTitle() + '" ...');
      if (currentEvent.isAllDayEvent() == true) {  // if it is an all-day event, throw an error
      appendToLog('"' + currentEvent.getTitle() + "\" is an all-day event. You can't change the absolute start time of an all-day event.");
      continue;
      }
      else {    
        var oldStart = new Date(currentEvent.getStartTime().getTime());
        var oldEnd = new Date(currentEvent.getEndTime().getTime());
        var newStart = new Date(currentEvent.getStartTime().getTime());
        newStart.setHours(absStart.getHours());
        newStart.setMinutes(absStart.getMinutes());
        var newEnd = new Date(newStart.getTime() + (oldEnd.getTime() - oldStart.getTime()));
        currentEvent.setTime(newStart,newEnd);
        appendToLog('Set start time for "' + currentEvent.getTitle() + '" from ' + oldStart + ' to ' + newStart);
      }
    }
    
    if (changeAbsEnd == "Yes") { // Change absolute end of time event
      appendToLog('Editing end time for "' + currentEvent.getTitle() + '"...');
      if (currentEvent.isAllDayEvent() == true) {  // if it is an all-day event, throw an error
        appendToLog('"' + currentEvent.getTitle() + "\" is an all-day event. You can't change the absolute start time of an all-day event.");
        continue;
      } 
      else {
        var oldStart = new Date(currentEvent.getStartTime().getTime());
        var oldEnd = new Date(currentEvent.getEndTime().getTime());
        var newStart = new Date(oldStart);
        var newEnd = new Date(oldEnd);
        newEnd.setHours(absEnd.getHours());
        newEnd.setMinutes(absEnd.getMinutes());
        
        if (newEnd.getTime() < newStart.getTime()) {
          appendToLog('Cannot set end time for "' + currentEvent.getTitle() + '" before start time.');
          continue;
        }
        
        currentEvent.setTime(newStart,newEnd);
        appendToLog('Set end time for "' + currentEvent.getTitle() + '" from ' + oldEnd + ' to ' + newEnd);
      }
    }
    
        
    if (addRemind == "Yes (Email)"){ //Add reminders to the event
      addReminders(calendarID,currentEvent.getId(),"Email",minutesRemind);
    }
    else if (addRemind == "Yes (SMS)"){
      addReminders(calendarID,currentEvent.getId(),"SMS",minutesRemind);
    }
    else if (addRemind == "Yes (Popup)"){
      addReminders(calendarID,currentEvent.getId(),"Popup",minutesRemind);
    }
    
    if (deleteRemind == "Yes"){
      if (deleteReminderType == "Email")
        deleteReminders(calendarID,currentEvent.getId(),"Email");
      else if (deleteReminderType == "SMS")
        deleteReminders(calendarID,currentEvent.getId(),"SMS");
      else if (deleteReminderType == "Popup")
        deleteReminders(calendarID,currentEvent.getId(),"Popup");
      else if (deleteReminderType == "All Reminders of All Types")
        deleteReminders(calendarID,currentEvent.getId(),"All");
    }
    
    if (visibility == "Yes"){
      changeVisibility(calendarID, currentEvent.getId(), visibilityTo);
    }
    
  }
}

function checkRequiredFields(){
  //Returns "true" if all required fields are filled
  //Otherwise returns "false"
  
  var sheet = SpreadsheetApp.getActive().getSheetByName("Home");
  var calendars = sheet.getRange("C5").getValue();
  var startTime = sheet.getRange("C6").getValue();
  var endTime = sheet.getRange("C7").getValue();
  
  if (calendars == ""){
    error('"' + sheet.getRange("B5").getValue() + '" is a required field');
    return false;
  }
  else if (startTime == ""){
    error('"' + sheet.getRange("B6").getValue() + '" is a required field');
    return false;
  }
  else if (endTime == ""){
    error('"' + sheet.getRange("B7").getValue() + '" is a required field');
    return false;
  }
  else if (endTime < startTime){
    error('"' + sheet.getRange("B7").getValue() + "\" can't be before \"" + sheet.getRange("B6").getValue() + '"');
    return false;
  }
  else
    return true;
}

function checkDependentFields(){
  //Returns "true" if all dependent fields for selected optional fields are filled
  //Otherwise returns "false"
  
  var sheet = SpreadsheetApp.getActive().getSheetByName("Home");
  
  if (sheet.getRange("C9").getValue() == "Yes" && sheet.getRange("E9").getValue() == ""){
    error('"' + sheet.getRange("B9").getValue() + '" + is selected, yet the "' + sheet.getRange("D9").getValue() + '" field is not filled');
    return false;
  }
  else if (sheet.getRange("C10").getValue() == "Yes" && sheet.getRange("E10").getValue() == ""){
    error('"' + sheet.getRange("B10").getValue() + '" + is selected, yet the "' + sheet.getRange("D10").getValue() + '" field is not filled');
    return false;
  }
  else if (sheet.getRange("C11").getValue() == "Yes" && sheet.getRange("E11").getValue() == ""){
    error('"' + sheet.getRange("B11").getValue() + '" + is selected, yet the "' + sheet.getRange("D11").getValue() + '" field is not filled');
    return false;
  }
  else if (sheet.getRange("C12").getValue() == "Yes" && sheet.getRange("E12").getValue() == ""){
    error('"' + sheet.getRange("B12").getValue() + '" + is selected, yet the "' + sheet.getRange("D12").getValue() + '" field is not filled');
    return false;
  }
  else if (sheet.getRange("C13").getValue() == "Yes" && sheet.getRange("E13").getValue() == ""){
    error('"' + sheet.getRange("B13").getValue() + '" + is selected, yet the "' + sheet.getRange("D13").getValue() + '" field is not filled');
    return false;
  }
  else if (sheet.getRange("C14").getValue() == "Yes" && sheet.getRange("E14").getValue() == ""){
    error('"' + sheet.getRange("B14").getValue() + '" + is selected, yet the "' + sheet.getRange("D14").getValue() + '" field is not filled');
    return false;
  }
  else if (sheet.getRange("C15").getValue() == "Yes" && sheet.getRange("E15").getValue() == ""){
    error('"' + sheet.getRange("B15").getValue() + '" + is selected, yet the "' + sheet.getRange("D15").getValue() + '" field is not filled');
    return false;
  }
  else if (sheet.getRange("C17").getValue() == "Yes" && sheet.getRange("E17").getValue() == ""){
    error('"' + sheet.getRange("B17").getValue() + '" + is selected, yet the "' + sheet.getRange("D17").getValue() + '" field is not filled');
    return false;
  }
  else if (sheet.getRange("C21").getValue() != "" && sheet.getRange("E21").getValue() == ""){
    error('A Keyword/Phrase is chosen, yet the "' + sheet.getRange("D21").getValue() + '" field is not filled');
    return false;
  }
  else
    return true;
}
