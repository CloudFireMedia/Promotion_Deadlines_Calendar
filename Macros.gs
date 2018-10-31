
function addRowBelow_() {

  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  spreadsheet.toast("Working...","",-1);
  var spreadSheet = SpreadsheetApp.getActive();
  var selectedRange= spreadSheet.getActiveRange().getValues(); 
  var selectedCellRow = spreadSheet.getCurrentCell().getRow();
  // check either more than one row or merged cell is selected
  if(selectedRange.length>1){
    spreadsheet.toast("","",0.01); // hack to close toast
    Browser.msgBox('Error! \\n\\nYou have selected more than one row. Please select only \\none cell or row and run the function again.');
  }
  else{ // if only one cell is selected
    if((selectedRange[0]+"").indexOf(',')==-1){
      spreadSheet.getRange(selectedCellRow+':'+selectedCellRow).activate(); // select whole row of selected cell
    }
    spreadSheet.getActiveSheet().insertRowsAfter(spreadSheet.getActiveRange().getLastRow(), 1);
    spreadSheet.getActiveRange().offset(spreadSheet.getActiveRange().getNumRows(), 0, 1, spreadSheet.getActiveRange().getNumColumns()).activate();
    spreadSheet.getCurrentCell().offset(0, 3).activate();
    spreadSheet.getCurrentCell().offset(-1, 0).copyTo(spreadSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    spreadSheet.getCurrentCell().offset(0, 2).activate();
    spreadSheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
                                                  .setText('No')
                                                  .setTextStyle(0, 2, SpreadsheetApp.newTextStyle().setFontSize(7).build())
                                                  .build());
    spreadSheet.getCurrentCell().offset(0, 1).activate();
    spreadSheet.getCurrentCell().setRichTextValue(SpreadsheetApp.newRichTextValue()
                                                  .setText('No')
                                                  .setTextStyle(0, 2, SpreadsheetApp.newTextStyle().setFontSize(7).build())
                                                  .build());
    spreadSheet.getCurrentCell().offset(0, 2).activate();
    spreadSheet.getCurrentCell().offset(-1, 0, 1, 3).copyTo(spreadSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    spreadSheet.getCurrentCell().offset(0, 3).activate();
    spreadSheet.getCurrentCell().offset(-1, 0).copyTo(spreadSheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    spreadSheet.getCurrentCell().offset(0, -8).activate();
    while(spreadSheet.getRange('A'+(selectedCellRow)).getValue()==''){ // go to top merged cell in column A of selected cell
      selectedCellRow =selectedCellRow-1;
    }
    var newRow = spreadSheet.getCurrentCell().getRow(); // get new added row
    var nextValue = spreadSheet.getRange('A'+(newRow+1)).getValue();
    if(nextValue != '') {
      spreadSheet.getRange('A' + selectedCellRow + ':'+ 'A'+ newRow ).activate().merge();  // merge cell in column A of new added row
      spreadSheet.getRange('B' + selectedCellRow + ':'+ 'B'+ newRow ).activate().mergeVertically(); // merge cell in column B of new added row
    }     
    spreadSheet.getRange('D'+newRow).activate(); // select new added row's cell in column D
    SpreadsheetApp.getActiveSpreadsheet().toast("Row added.");
  }
};

function deleteRow_() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
  SpreadsheetApp.getActiveSpreadsheet().toast("Row deleted.");
};

//macro menu item 'Initialize Promotion Deadlines'
function calculateDeadlines_(){
  SpreadsheetApp.getActiveSpreadsheet().toast("Working...","",-1);
  var tierDateSheet = SpreadsheetApp.openById(TIER_DUEDATE_SHEET_ID_).getSheetByName(TIER_DUEDATE_SHEET_NAME_);
  var spreadSheet = SpreadsheetApp.getActive();
  // getting tiers name from tiers due Date sheet
  var tier1Name = tierDateSheet.getRange('A2').getValue();
  var tier2Name = tierDateSheet.getRange('A3').getValue();
  var tier3Name = tierDateSheet.getRange('A4').getValue();
  // getting tiers due dates from tiers due Date sheet
  var tier1DueDate = tierDateSheet.getRange('B2').getValue();
  var tier2DueDate = tierDateSheet.getRange('B3').getValue();
  var tier3DueDate = tierDateSheet.getRange('B4').getValue();
  // writing teirs name to active sheet
  spreadSheet.getRange('I3').setValue(tier1Name);
  spreadSheet.getRange('J3').setValue(tier2Name);
  spreadSheet.getRange('K3').setValue(tier3Name);
  // getting range
  var rangeColumnJ = spreadSheet.getRange("J1");
  var rangeColumnK = spreadSheet.getRange("K1");
  // hiding/ unhiding column J 
  if (tier2DueDate ==''){
    spreadSheet.hideColumn(rangeColumnJ);
  }
  else{
    spreadSheet.unhideColumn(rangeColumnJ)
  }
   // hiding/ unhiding column K 
  if (tier3DueDate ==''){
    spreadSheet.hideColumn(rangeColumnK);
  }
  else{
    spreadSheet.unhideColumn(rangeColumnK)
  }
  // save value of tier1, tier2 and tier3 for later use
  PropertiesService.getScriptProperties().setProperty('tier1DueDate', tier1DueDate);
  PropertiesService.getScriptProperties().setProperty('tier2DueDate', tier2DueDate);
  PropertiesService.getScriptProperties().setProperty('tier3DueDate', tier3DueDate);
  var lastRow = spreadSheet.getLastRow();
  // loop through each cell in column D to update column I, J K
  for (var c = 4; c < lastRow+1; c++) {
    var startDate  = spreadSheet.getRange('D' + c ).getValue();
    //finding value for column I
    var newTier1 = new Date(startDate.getTime()-tier1DueDate*3600000*24);
    var day = (newTier1+"").substring(0,3);
    // if day is Sunday than move it to Monday
    if(day=='Sun'){
      newTier1 = new Date(newTier1.getTime()+1*3600000*24);
    }
    // if day is Saturday than move it to Friday
    if(day=='Sat'){
      newTier1 = new Date(newTier1.getTime()-1*3600000*24);
    }
    spreadSheet.getRange('I' + c ).setValue(newTier1); // setting value of active cell for Column I = tier1
    //Finding Value for Column J if tier2 is  not empty
    if(tier2DueDate != ''){
      var newTier2 = new Date(startDate.getTime()-tier2DueDate*3600000*24);
      day = (newTier2+"").substring(0,3);
      // if day is Sunday than move it to Monday
      if(day=='Sun'){
        newTier2 = new Date(newTier2.getTime()+1*3600000*24);
      }
      // if day is Saturday than move it to Friday
      if(day=='Sat'){
        newTier2 = new Date(newTier2.getTime()-1*3600000*24);
      }
      spreadSheet.getRange('J' + c ).setValue(newTier2); // setting value of active cell for Column J = tier2
    }
    
    //finding value for column K if tier3 is not empty
    if(tier3DueDate != ''){
      var newTier3 = new Date(startDate.getTime()-tier3DueDate*3600000*24);
      day = (newTier3+"").substring(0,3);
      // if day is Sunday than move it to Monday
      if(day=='Sun'){
        newTier3 = new Date(newTier3.getTime()+1*3600000*24);
      }
      // if day is Saturday than move it to Friday
      if(day=='Sat'){
        var newTier3 = new Date(newTier3.getTime()-1*3600000*24);
      }
      spreadSheet.getRange('K' + c ).setValue(newTier3); // setting value of active cell for Column K = tier2
    }
    
  } // end of loop
  
  SpreadsheetApp.getActiveSpreadsheet().toast("Success! Promotion Deadlines have been assigned to all events in the current sheet.");
  
}; // calculateDeadlines()

//  macro to update tier1, tier2 and tier3 when value of cell  in column D (Start Date) is changed
function onEdit_(e) 
{
  var spreadSheet = SpreadsheetApp.getActive();
  var thisCol = e.range.getColumn(); // get column in which value is changed
  if (thisCol != 4) return; // if change is not in column D exist macro
  // getting value of tier1, tier2 and tier3 from properties service that is saved during initialization setup
  var tier1Value=   PropertiesService.getScriptProperties().getProperty('tier1DueDate');
  var tier2Value =  PropertiesService.getScriptProperties().getProperty('tier2DueDate');
  var tier3Value = PropertiesService.getScriptProperties().getProperty('tier3DueDate');
  // check tier1, tier2 and tier3 value is set during initialization setup
  if(tier1Value == null){ 
    Browser.msgBox('Error! \\n\\nDefault values for tier1, tier2 and tier3 Promotion Deadlines have not been assigned, and deadlines for the selected row cannot be automatically updated. \\n\\nPlease run \'Macros > Initialize Promotion Deadlines\' to assign default Promotion Deadlines values.');
    return;
  }
  var eidtRow = spreadSheet.getCurrentCell().getRow();
  var newDate  = spreadSheet.getRange('D' + eidtRow ).getValue(); // getting updated date in column D
  //updating value for tier1 in Column I
  var newtier1Value = new Date(newDate.getTime()-tier1Value*3600000*24);
  var day = (newtier1Value+"").substring(0,3);
  // if day is Sunday than set it to next Monday
  if(day=='Sun'){
    newtier1Value = new Date(newtier1Value.getTime()+1*3600000*24);
  }
  // if day is Saturday than set it to  previous Friday
  if(day=='Sat'){
    newtier1Value = new Date(newtier1Value.getTime()-1*3600000*24);
  }
  spreadSheet.getRange('I' + eidtRow ).setValue(newtier1Value); // setting new value of tier1
  if(tier2Value!=''){
    //updating value for tier2 in Column J
    var newtier2Value = new Date(newDate.getTime()-tier2Value*3600000*24);
    var day = (newtier2Value+"").substring(0,3);
    // if day is Sunday than set it to next Monday
    if(day=='Sun'){
      newtier2Value = new Date(newtier2Value.getTime()+1*3600000*24);
    }
    // if day is Saturday than set it to  previous Friday
    if(day=='Sat'){
      newtier2Value = new Date(newtier2Value.getTime()-1*3600000*24);
    }
    spreadSheet.getRange('J' + eidtRow ).setValue(newtier2Value); // setting new value of tier2
  }
  if(tier3Value!=''){
    //updating value for tier3 in Column K
    var newtier3Value = new Date(newDate.getTime()-tier3Value*3600000*24);
    day = (newtier3Value+"").substring(0,3);
    // if day is Sunday than set it to next Monday
    if(day=='Sun'){
      newtier3Value = new Date(newtier3Value.getTime()+1*3600000*24);
    }
    // if day is Saturday than set it to  previous Friday
    if(day=='Sat'){
      newtier3Value = new Date(newtier3Value.getTime()-1*3600000*24);
    }
    spreadSheet.getRange('K' + eidtRow ).setValue(newtier3Value); // setting new value of tier3
  }
  SpreadsheetApp.getActiveSpreadsheet().toast("Success! Promotion Deadlines have been updated for the selected row.");
} // end of onEdit_ macro

function newEventPopup_() {
  var spreadSheet = SpreadsheetApp.getActive();
  var htmlService = HtmlService.createHtmlOutputFromFile('Form')
    .setWidth(750)  // width of form
    .setHeight(480) // height of form
    .setSandboxMode(HtmlService.SandboxMode.NATIVE);
  spreadSheet.show(htmlService);
} // end of NewEventPopup macro

// macro for getting sponsors to fill  sponsor dropdownlist of new event form
function fillSponsor_() {
  var staffSheet = SpreadsheetApp.openById(SPONSOR_SHEET_ID_).getSheetByName(SPONSOR_SHEET_NAME_);
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
  var tierSheet = SpreadsheetApp.openById(TIER_DUEDATE_SHEET_ID_).getSheetByName(TIER_DUEDATE_SHEET_NAME_);
  var lastRow = tierSheet.getLastRow();
  var optionsValue = tierSheet.getRange("A2:A" + lastRow).getValues();;
  var optionsArray = new Array();
  // loop throgh each cell in range to fill optionArray
  for(var i = 0; i < lastRow; i++) {
    optionsArray.push(optionsValue[i]);
  } // end of loop
  return optionsArray;
}; // end of fillTier macro