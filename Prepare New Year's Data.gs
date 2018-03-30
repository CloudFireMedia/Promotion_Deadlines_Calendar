////prepares data in Columns A, B, F, G that has been copied from Calendar Planning Team.gsheet
//
//function prepareNewYearsData() {
//  var sheet = SpreadsheetApp.getActiveSheet();
//  var dataRange = sheet.getDataRange(); // Set Any Range
//  var data = dataRange.getValues();
//  var numRows = data.length;
//  var numColumns = data[0].length;
//  
//  var values = [];
//  for (var i = 3; i < numRows; i++) {
//      //ui.alert(data[i][2]);
//      //return;
//      if(data[i][2]=="N/A" && data[i][5]=="" && data[i][6]=="")
//      {
//        sheet.getRange(i+1,6).setValue('N/A'); 
//        sheet.getRange(i+1,7).setValue('N/A');         
//      }
//    else if(data[i][2]!="N/A" && data[i][5]=="" && data[i][6]=="")
//    {
//      sheet.getRange(i+1,6).setValue('No'); 
//      sheet.getRange(i+1,7).setValue('No');  
//    }
//    
//  }
//  getSundayServices();
//  calculateWeekNumber();
//}
//
//function getSundayServices()
//{
//  var sheet = SpreadsheetApp.getActiveSheet();
//  var dataRange = sheet.getDataRange(); // Set Any Range
//  var data = dataRange.getValues();
//  var numRows = data.length;
//  var numColumns = data[0].length;
//  
//  var values = [];
//  var indexesOfSunday=[];
//  
//  for (var i =4; i < numRows; i++) {
//    var eColumnValue=data[i][4];
//    if(eColumnValue.indexOf("Sunday Service")!="-1")
//    {
//      indexesOfSunday.push(i);
//    }
//  }
//  combineCell(indexesOfSunday);  
//}
//
//function combineCell(indexesOfSunday)
//{
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var s = ss.getSheetByName('Communications Director Master');
//  var maxValue=indexesOfSunday[indexesOfSunday.length-1];
//  
//  for(var k=0;k<indexesOfSunday.length;k++)
//  {
//    var from =indexesOfSunday[k];
//    var to =indexesOfSunday[k+1];
//    from=from+1
//    var vardiff=to-from;
//    vardiff=vardiff+1
//    if(from-1==maxValue)
//    {
//        s.getRange(from,1,1,2).mergeVertically();
//    }
//    else
//    {
//       s.getRange(from,1,vardiff,2).mergeVertically();  
//    }
//  }
//}
//
//function weeksBetween(d1, d2) {
//    return Math.round((d2 - d1) / (7 * 24 * 60 * 60 * 1000));
//}
//
//function calculateWeekNumber() {
//  var sheet = SpreadsheetApp.getActiveSheet();
//  var dataRange = sheet.getDataRange(); // Set Any Range
//  var data = dataRange.getValues();
//  var numRows = data.length;
//  var numColumns = data[0].length;  
//  var values = [];
//  for (var i = 3; i < numRows; i++) {
//    var dateFromSheet=data[i][3];
//    var year=dateFromSheet.getYear();
//    if(year=="2017")
//    {
//      var startDate=new Date("01/01/2017");
//    }
//    if(year=="2018")
//    {
//      var startDate=new Date("12/31/2017");
//    }
//    if(year=="2019")
//    {
//      var startDate=new Date("12/30/2019");
//    }
//    var weeks=weeksBetween(startDate,dateFromSheet);
//    if(weeks==52)
//    {
//      weeks=1;
//    }
//    else
//    {
//      weeks=weeks+1;
//    }
//    sheet.getRange(i+1,1).setValue(weeks);    
//  }
//}
