//function updateRecordsOnEdit() {
//  var source = SpreadsheetApp.getActiveSpreadsheet();
//  var source_sheet = source.getSheetByName("Communications Director Master")
//  var row = source_sheet.getActiveCell().getRow();
//  var i=row-1;  
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getActiveSheet();
//  var dataRange = sheet.getDataRange();
//  var values = dataRange.getValues();
//  var fColumn=values[i][5];
//  var gColumn=values[i][6];
//  if(fColumn=="N/A" && gColumn=="N/A")
//  {
//    i=i+1;
//    sheet.getRange(i,3).setValue('N/A'); 
//  }
//}