
function addRowBelow_() {

  var spreadSheet = getSpreadsheet_()
  
  toast_("Working...","",-1) 

  var selectedRange = spreadSheet.getActiveRange().getValues()
  var selectedCellRow = spreadSheet.getCurrentCell().getRow()
  
  if (selectedCellRow < 4) {
    toast_("","",0.01) // hack to close toast
    msgBox_('Error! \\n\\nPlease select only a cell or row in row 4 or more.')
    return  
  }
  
  // check either more than one row or merged cell is selected
  if (selectedRange.length > 1) {
    toast_("","",0.01) // hack to close toast
    msgBox_('Error! \\n\\nYou have selected more than one row. Please select only \\none cell or row and run the function again.')
    return
  }
  
  if ((selectedRange[0] + "").indexOf(',') === -1){
    spreadSheet.getRange(selectedCellRow+':'+selectedCellRow).activate() // select whole row of selected cell
  }
  
  spreadSheet.getActiveSheet().insertRowsAfter(spreadSheet.getActiveRange().getLastRow(), 1)
  var numRows = spreadSheet.getActiveRange().getNumRows()
  var numColumns = spreadSheet.getActiveRange().getNumColumns()
  spreadSheet.getActiveRange().offset(numRows, 0, 1, numColumns).activate() // Select the new row
  
  // Copy the start date to the new row
  spreadSheet.getCurrentCell().offset(0, 3).activate() // Select the start date in the new row
  spreadSheet.getCurrentCell().offset(-1, 0).copyTo(spreadSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
  
  spreadSheet.getCurrentCell().offset(0, 2).activate()
  spreadSheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
                                                .setText('No')
                                                .setTextStyle(0, 2, SpreadsheetApp.newTextStyle().setFontSize(7).build())
                                                .build())
  
  spreadSheet.getCurrentCell().offset(0, 1).activate()
  spreadSheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
                                                .setText('No')
                                                .setTextStyle(0, 2, SpreadsheetApp.newTextStyle().setFontSize(7).build())
                                                .build())
  // Copy the deadlines                                              
  spreadSheet.getCurrentCell().offset(0, 2).activate()
  spreadSheet.getCurrentCell().offset(-1, 0, 1, 3).copyTo(spreadSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false)
  
  // Copy the promo status - disabled for now as it has been replaced by sheet formula
  spreadSheet.getCurrentCell().offset(0, 3).activate()
  
  // go to top merged cell in column A of selected cell
  spreadSheet.getCurrentCell().offset(0, -10).activate()
  while (spreadSheet.getRange('A'+(selectedCellRow)).getValue()==''){ 
    selectedCellRow--
  }
  
  var newRow = spreadSheet.getCurrentCell().getRow() // get new added row number
  var nextValue = spreadSheet.getRange('A'+(newRow+1)).getValue()
  
  if (nextValue != '') {
    spreadSheet.getRange('A' + selectedCellRow + ':'+ 'A'+ newRow ).activate().merge()  // merge cell in column A of new added row
    spreadSheet.getRange('B' + selectedCellRow + ':'+ 'B'+ newRow ).activate().mergeVertically() // merge cell in column B of new added row
  }     
  
  spreadSheet.getRange('D'+newRow).activate() // select new added row's cell in column D  
  toast_("Row added.")
  
};

function deleteRow_() {
  var spreadSheet = getSpreadsheet_()
  spreadSheet.getActiveSheet().deleteRows(spreadSheet.getActiveRange().getRow(), spreadSheet.getActiveRange().getNumRows())
  toast_("Row deleted.")
};

//macro menu item 'Calculate Deadlines'
function calculateDeadlines_() {

  toast_("Working...","",-1)
  
  var promotionRequestResponsesSheetId = Config.get('PROMOTION_FORM_RESPONSES_GSHEET_ID')
  var tierDateSheet = SpreadsheetApp.openById(promotionRequestResponsesSheetId).getSheetByName(TIER_DUEDATE_SHEET_NAME_)
  
  var ss = getSpreadsheet_()
  var cdmSheet = ss.getSheetByName(CDM_SHEET_NAME_)
  var cdmValues = cdmSheet.getDataRange().getValues()
    
  var tier1Name = tierDateSheet.getRange('A2').getValue()
  var tier2Name = tierDateSheet.getRange('A3').getValue()
  var tier3Name = tierDateSheet.getRange('A4').getValue()
  
  var tier1DueDate = 7 * tierDateSheet.getRange('B2').getValue()
  var tier2DueDate = 7 * tierDateSheet.getRange('B3').getValue()
  var tier3DueDate = 7 * tierDateSheet.getRange('B4').getValue()
  
  cdmSheet.getRange('I3').setValue(tier1Name) 
  cdmSheet.getRange('J3').setValue(tier2Name) 
  cdmSheet.getRange('K3').setValue(tier3Name) 
  
  var rangeColumnJ = cdmSheet.getRange("J1")
  var rangeColumnK = cdmSheet.getRange("K1")
  
  if (tier2DueDate === '') {
    cdmSheet.hideColumn(rangeColumnJ)
  } else {
    cdmSheet.unhideColumn(rangeColumnJ)
  }
  
  if (tier3DueDate === '') {
    cdmSheet.hideColumn(rangeColumnK)
  } else{
    cdmSheet.unhideColumn(rangeColumnK)
  }
  
  // save value of tier1, tier2 and tier3 for later use
  PropertiesService
    .getDocumentProperties()
      .setProperty('tier1DueDate', tier1DueDate)
      .setProperty('tier2DueDate', tier2DueDate)
      .setProperty('tier3DueDate', tier3DueDate)
    
  var startRowIndex = cdmSheet.getFrozenRows()
  var numberOfRows = cdmValues.length
  
  // loop through each cell in column D to update column I, J K
  for (var rowIndex = startRowIndex; rowIndex < numberOfRows; rowIndex++) {  
    calculateRowDeadline_(cdmSheet, rowIndex, cdmValues) 
  }
  
  toast_("Success! Promotion Deadlines have been assigned to all events in the current sheet.");
  
} // calculateDeadlines()

//  macro to update tier1, tier2 and tier3 when value of cell in column D (EventDate) is changed
function onEdit_(e) {
  var sheet = spreadSheet.getSheetByName(CDM_SHEET_NAME_)
  var thisCol = e.range.getColumn(); // get column in which value is changed
  if (thisCol != 4) return; // if change is not in column D exist macro
  var rowIndex = sheet.getCurrentCell().getRow();  
  var values = sheet.getDataRange().getValues()
  calculateRowDeadline_(sheet, rowIndex, values)
  toast_("Success! Promotion Deadlines have been updated for the selected row.");
  
} // end of onEdit_ macro

function newEventPopup_() {
  var spreadSheet = getSpreadsheet_();
  var htmlService = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(750)  // width of form
    .setHeight(480) // height of form
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
  spreadSheet.show(htmlService);
} // end of NewEventPopup macro

// macro for getting sponsors to fill  sponsor dropdownlist of new event form
function fillSponsor_() {
  var staffSheetId = Config.get('STAFF_DATA_GSHEET_ID');
  var staffSheet = SpreadsheetApp.openById(staffSheetId).getSheetByName(SPONSOR_SHEET_NAME_);
  var lastRow = staffSheet.getLastRow();
  var optionsValue = staffSheet.getRange("L3:L" + lastRow).getValues(); 
  var optionsArray = new Array();
  // loop throgh each cell in range to fill optionArray with unique values
  for(var i = 0; i < lastRow; i++) {
    if(optionsValue[i]!= "N/A" && optionsValue[i]!= "" && optionsValue[i]!= null){
      // code to check duplication
      var duplicate = false; 
      // loop through optionArray to check if it already exist in optionArray or not
      for(var j = 0; j < optionsArray.length; j++){
        if(optionsArray[j][0] == optionsValue[i][0]){
          duplicate = true; 
        }
      } // end of inner loop
      if(!duplicate){
      optionsArray.push(optionsValue[i]);
      }
    } 
  } // end of outer loop
  return optionsArray; 
}; // end of fillSponsor function

// macro for getting tiers to fill  tier dropdownlist of new event form
function fillTier_() {
  var promotionRequestResponsesSheetId = Config.get('PROMOTION_FORM_RESPONSES_GSHEET_ID');
  var tierSheet = SpreadsheetApp.openById(promotionRequestResponsesSheetId).getSheetByName(TIER_DUEDATE_SHEET_NAME_);
  var lastRow = tierSheet.getLastRow();
  var optionsValue = tierSheet.getRange("A2:A" + lastRow).getValues();;
  var optionsArray = new Array();
  // loop throgh each cell in range to fill optionArray
  for(var i = 0; i < lastRow; i++) {
    optionsArray.push(optionsValue[i]);
  } // end of loop
  return optionsArray;
}; // end of fillTier macro

// macro that executes when New event form is submitted
function processResponse_(eventArray) {

  var eventTitle = eventArray[0];
  var eventTier = eventArray[1];
  var dateSplit = eventArray[2].split("/");
  var eventDate = new Date(dateSplit[2],dateSplit[0]-1,dateSplit[1]);
  var eventSponsor = eventArray[3];
  var spreadSheet = getSpreadsheet_()
  var cdmSheet = spreadSheet.getSheetByName(CDM_SHEET_NAME_)
  
  Log_.fine(
    'eventTitle: "%s", eventTier: %s, eventDate: %s, eventSponsor: %s', 
    eventTitle,
    eventTier,
    eventDate,
    eventSponsor);
  
  var lastRow = cdmSheet.getLastRow();
  
  // code for setting new event ordered chronologically by Col D.
  var dataValues = cdmSheet.getRange("D4:D" + lastRow).getValues(); 
  var flag = 0; // to check if row found in chronologically order, 0 = not found, 1 = found
  
  for (var i=0; i < dataValues.length; i++) {
  
    if (new Date(dataValues[i]) >= eventDate) { 
    
      flag = 1; 
      
      // equality operator is not possible for date
      if ((new Date(dataValues[i])).getTime() == (eventDate).getTime()) {
      
        // if date already exists
        index = i + 4;
        
      } else {
      
        // if date does not already exist      
        index = i + 3;
      }

      writeRow(index++);
      break;
    } 
  }
  
  if (flag === 0) {
  
  // No row found in chronologically order so place new event at the end, event is latest  
    writeRow(lastRow + 1);
  }
  
  return;
  
  // Private Functions
  // -----------------
  
  function writeRow(index) {
    cdmSheet.getRange("D" + lastRow).activate();
    addRowBelow_();
    cdmSheet.getRange("C"+ index).setValue(eventTier);
    cdmSheet.getRange("D"+ index).setValue(eventDate);
    cdmSheet.getRange("E"+ index).setValue(eventTitle); 
    cdmSheet.getRange("H"+ index).setValue(eventSponsor);
    cdmSheet.getRange("E"+ index).activate();  
  }
  
} // processResponse_()

function calculateRowDeadline_(sheet, rowIndex, values) {
  
  // getting value of tier1, tier2 and tier3 from properties service that is saved during initialization setup
  var tier1DueDate = PropertiesService.getDocumentProperties().getProperty('tier1DueDate');
  var tier2DueDate = PropertiesService.getDocumentProperties().getProperty('tier2DueDate');
  var tier3DueDate = PropertiesService.getDocumentProperties().getProperty('tier3DueDate');
  
  // check tier1, tier2 and tier3 value is set during initialization setup
  if (tier1DueDate == null) { 
    msgBox_(
      'Error! \\n\\nDefault values for tier1, tier2 and tier3 Promotion Deadlines have not been assigned, ' + 
        'and deadlines for the selected row cannot be automatically updated. \\n\\n' + 
        'Please run \'Macros > Calculate Deadlines\' to assign default Promotion Deadlines values.');
    return;
  }
  
  var rowNumber = rowIndex + 1
  Log_.fine(values) 
  var eventDate  = formatEventDate_(values, rowIndex)
    
  Log_.fine('eventDate ' + rowIndex + eventDate)
  Log_.fine('tier1DueDate' + tier1DueDate)
  
  // finding value for column I
  var newTier1 = new Date(eventDate.getTime()-tier1DueDate*3600000*24);
  var day = (newTier1+"").substring(0,3);
    
  // if day is Sunday than move it back to a friday
  if (day === 'Sun') {
    newTier1 = new Date(newTier1.getTime()-2*3600000*24);
    Log_.fine('This day is a Sunday, moving forward to a Friday')
  }
    
  // if day is Saturday than move it to Friday
  if (day === 'Sat') {
    newTier1 = new Date(newTier1.getTime()-1*3600000*24);
    Log_.fine('This day is a Saturday, moving forward to a Friday')
  }
    
  Log_.fine('newTier1 ' + newTier1)
  sheet.getRange('I' + rowNumber ).setValue(newTier1); // setting value of active cell for Column I = tier1
  
  // Finding Value for Column J if tier2 is  not empty (in weeks)
  if (tier2DueDate !== '') {
  
    var newTier2 = new Date(eventDate.getTime()-(tier2DueDate*3600000*24));
    day = (newTier2+"").substring(0,3);
    
    // if day is Sunday than move it to Monday
    if (day === 'Sun') {
      newTier2 = new Date(newTier2.getTime()+1*3600000*24);
      Log_.fine('This day is a Sunday, moving forward to a Friday')
    }
    
    // if day is Saturday than move it to Friday
    if (day === 'Sat') {
      newTier2 = new Date(newTier2.getTime()-(1*3600000*24));
      Log_.fine('This day is a Saturday, moving forward to a Friday')
    }
    
    sheet.getRange('J' + rowNumber ).setValue(newTier2); // setting value of active cell for Column J = tier2
  }
  
  //finding value for column K if tier3 is not empty (in weeks)
  if (tier3DueDate !== '') {
    
    var newTier3 = new Date(eventDate.getTime()-(tier3DueDate*3600000*24));
    day = (newTier3+"").substring(0,3);
    
    // if day is Sunday than move it to Monday
    if (day === 'Sun') {
      newTier3 = new Date(newTier3.getTime()+1*3600000*24);
      Log_.fine('This day is a Sunday, moving forward to a Friday')
    }
    
    // if day is Saturday than move it to Friday
    if (day === 'Sat') {
      var newTier3 = new Date(newTier3.getTime()-1*3600000*24);
      Log_.fine('This day is a Saturday, moving forward to a Friday')
    }
    
    sheet.getRange('K' + rowNumber ).setValue(newTier3); // setting value of active cell for Column K = tier2
  }
    
} // calculateRowDeadline_()