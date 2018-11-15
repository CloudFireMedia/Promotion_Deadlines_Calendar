function disableAutomation_() {

  var owner = SpreadsheetApp.getActive().getOwner().getEmail();
  var user = Session.getActiveUser().getEmail();
  if( user != owner){
    Browser.msgBox('Disable Automation', "Sorry.  Automation can only be disabled by the sheet owner.\\nPlease ask "+owner+" to run this.", Browser.Buttons.OK);
    return;
  }

  deleteTriggerByHandlerName_('dailyTrigger');
  
  Browser.msgBox('Disable Automation', "Automation has been disabled.\\nThe daily notifcations will not be sent until this is enabled again.", Browser.Buttons.OK);
}

function setupAutomation_() {

  var owner = SpreadsheetApp.getActive().getOwner().getEmail();
  var user = Session.getActiveUser().getEmail();
  var activeSpreadsheet = SpreadsheetApp.getActive();
  
  if (user != owner && activeSpreadsheet !== null){
    Browser.msgBox('Enable Automation', "Sorry.  Automation can only be enabled by the sheet owner.\\nPlease ask "+owner+" to run this.", Browser.Buttons.OK);
    return;
  }
  
  //setup daily trigger
  deleteTriggerByHandlerName_('dailyTrigger');//remove any existing triggers so we don't have conflicts
  ScriptApp.newTrigger('dailyTrigger').timeBased().everyDays(1).atHour(5).nearMinute(30).create();

  if (activeSpreadsheet !== null) {
    Browser.msgBox('Enable Automation', 'Done!', Browser.Buttons.OK) 
  }
}

function deleteTriggerByHandlerName_(handlerName){ // Delete a trigger based on the name of the function it excecutes

  var allTriggers = ScriptApp.getProjectTriggers(),
      deleted = false;
      
  for (var i=0; i < allTriggers.length; i++){
    if (allTriggers[i].getHandlerFunction() == handlerName){ // Found the trigger we're looking for
      ScriptApp.deleteTrigger(allTriggers[i]);
      deleted = true;
      //break;//might be more than one, keep looking
    }
  }
  return deleted;
}


