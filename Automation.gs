function disableAutomation_() {

  var owner = getSpreadsheet_().getOwner().getEmail();
  var user = Session.getActiveUser().getEmail();
  
  if (user != owner) {
  
    msgBox_(
      'Disable Automation', 
      "Sorry.  Automation can only be disabled by the sheet owner.\\nPlease ask " + owner + " to run this.", 
      Browser.Buttons.OK);
      
    return;
  }

  var response = msgBox_(
    'Disable Automation', 
    'Please confirm that you want to disable automation', 
    Browser.Buttons.OK_CANCEL);
  
  Log_.fine(response)

  if (response !== 'ok') {
    msgBox_('Disable Automation', 'Automation disable cancelled', Browser.Buttons.OK);
    return;
  }

  deleteTriggerByHandlerName_('dailyTrigger');
  
  msgBox_(
    'Disable Automation', 
    "Automation has been disabled.\\nThe daily notifcations will not be sent until this is enabled again.", 
    Browser.Buttons.OK);
}

function setupAutomation_() {

  var activeSpreadsheet = getSpreadsheet_()
  var owner = activeSpreadsheet.getOwner().getEmail();
  var user = Session.getActiveUser().getEmail();
  
  if (user != owner && activeSpreadsheet !== null) {
  
    msgBox_(
      'Enable Automation', 
      "Sorry.  Automation can only be enabled by the sheet owner.\\nPlease ask " + owner + " to run this.", 
      Browser.Buttons.OK);
      
    return;
  }
  
  var response = msgBox_(
    'Enable Automation', 
    'Please confirm that you want to enable automation', 
    Browser.Buttons.OK_CANCEL);
  
  Log_.fine(response);

  if (response !== 'ok') {
    msgBox_('Enable Automation', 'Automation enable cancelled', Browser.Buttons.OK);
    return;
  }
  
  deleteTriggerByHandlerName_('dailyTrigger');
  ScriptApp.newTrigger('dailyTrigger').timeBased().everyDays(1).atHour(5).nearMinute(30).create();

  if (activeSpreadsheet !== null) {
    msgBox_('Enable Automation', 'Done!', Browser.Buttons.OK) 
  }
}

// Delete a trigger based on the name of the function it excecutes
function deleteTriggerByHandlerName_(handlerName) { 

  var allTriggers = ScriptApp.getProjectTriggers();
  var deleted = false;
      
  for (var i=0; i < allTriggers.length; i++){
    if (allTriggers[i].getHandlerFunction() == handlerName){ // Found the trigger we're looking for
      ScriptApp.deleteTrigger(allTriggers[i]);
      deleted = true;
      //break;//might be more than one, keep looking
    }
  }
  return deleted;
}


