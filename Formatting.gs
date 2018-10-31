function formatSheet_(){
  colorBorders_();
  setWeeksFormat_();
  setEventsFormat_();
  removeEmptyRows_();
  hideRows_();
}

// Sets internal borders to white solid
function colorBorders_() { 
  if( ! SpreadsheetApp.getActive()) var sheet = getSheetDev_(); else //debug
  var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME_);
  var cell = sheet.getRange('A'+(sheet.getFrozenRows()+1)+':L'+sheet.getLastRow());
  cell.setBorder(false, false, false, false, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
}

//Remove all empty rows at the end (not in the data)
function removeEmptyRows_() { 
  if( ! SpreadsheetApp.getActive()) var sheet = getSheetDev_(); else //debug
  var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME);
  removeExtraRows(sheet);
}

//Hide all rows before today's date
function hideRows_() { 
  //hide rows prior to todays date
  if( ! SpreadsheetApp.getActive()) var sheet = getSheetDev_(); 
  else //debug
    var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME);
  var values = sheet.getRange(sheet.getFrozenRows()+1, 4, sheet.getLastRow(), 1).getValues();
  var today = new Date(new Date().setHours(0,0,0,0));//midnight today
  
  for(var v=0; v<values.length; v++)
    if(new Date(values[v]) >= today) break;//sets v to the last row to hide
  if(v)//if there's a value, hide the rows
    sheet.hideRows(sheet.getFrozenRows()+1, v);
}

function setWeeksFormat_() { 
  if( ! SpreadsheetApp.getActive()) var sheet = getSheetDev_(); else //debug
  var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME);
  var range = sheet.getRange(sheet.getFrozenRows()+1, 1, sheet.getMaxRows()-sheet.getFrozenRows());

  range
  .setFontFamily('Lato')
  .setFontSize('7')
  .setFontWeight('Bold')
  .setFontColor('#ACC5CF')
  ;
}

function setEventsFormat_() { 
  if( ! SpreadsheetApp.getActive()) var sheet = getSheetDev_(); else //debug
  var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME_);
  //range: everything but the first col
  var range = sheet.getRange(sheet.getFrozenRows()+1, 4, sheet.getMaxRows()-sheet.getFrozenRows(), sheet.getMaxColumns()-3);

  range
  .setFontFamily('Lato')
  .setFontSize('7')
  .setFontWeight('Normal')
  .setFontColor('#acc5cf')
  .setBackgroundColor('#ffffff')
  ;
}