function formatSheet_() {

  var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME_);

  colorBorders_();
  setWeeksFormat_();
  setEventsFormat_();
  removeEmptyRows_();
  hideRows_();
  
  // Private Functions
  // -----------------
  
  // Sets internal borders to white solid
  function colorBorders_() { 
    var cell = sheet.getRange('A'+(sheet.getFrozenRows()+1)+':L'+sheet.getLastRow());
    cell.setBorder(false, false, false, false, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
  }
  
  //Remove all empty rows at the end (not in the data)
  function removeEmptyRows_() {   
  
    var maxRows = sheet.getMaxRows();
    var lastRow = getLastPopulatedRow(sheet);
    if (maxRows - lastRow > 0) {
      sheet.deleteRows(lastRow+1, maxRows-lastRow);
    }
      
    return
      
    // Private Functions
    // -----------------
    
    function getLastPopulatedRow(sheet) {
    
      var values = sheet.getDataRange().getValues();
      
      for (var i=values.length-1; i>0; i--) {
      
        if (values[i].join('').length) {
          return ++i;
        }
      }
      
      return 0; 
    }
    
  }
  
  // Hide all rows before today's date
  function hideRows_() { 
  
    var values = sheet.getRange(sheet.getFrozenRows()+1, 4, sheet.getLastRow(), 1).getValues();
    var today = new Date(new Date().setHours(0,0,0,0));//midnight today
    
    for (var v=0; v<values.length; v++) {
    
      if (new Date(values[v]) >= today) {
        break;//sets v to the last row to hide
      }
      
      if (v) {//if there's a value, hide the rows
        sheet.hideRows(sheet.getFrozenRows()+1, v);
      }
    }
    
  } // hideRows_()
  
  function setWeeksFormat_() { 
  
    var range = sheet.getRange(sheet.getFrozenRows()+1, 1, sheet.getMaxRows()-sheet.getFrozenRows());
  
    range
      .setFontFamily('Lato')
      .setFontSize('7')
      .setFontWeight('Bold')
      .setFontColor('#ACC5CF');
      
  } // setWeeksFormat_()
  
  function setEventsFormat_() { 
  
    var range = sheet.getRange(sheet.getFrozenRows()+1, 4, sheet.getMaxRows()-sheet.getFrozenRows(), sheet.getMaxColumns()-3);
  
    range
      .setFontFamily('Lato')
      .setFontSize('7')
      .setFontWeight('Normal')
      .setFontColor('#acc5cf')
      .setBackgroundColor('#ffffff');
      
  } // setEventsFormat_()
  
} // formatSheet_()