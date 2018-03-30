//// Redevelopment Note: 
//
//
//// Sets vertical and horizontal borders to white solid
//function colorBorders() {
//var ss = SpreadsheetApp.getActive();
//var sheet = ss.getActiveSheet();
//var cell = sheet.getRange("A4:L1000");
//cell.setBorder(false, false, false, false, true, true, "white", SpreadsheetApp.BorderStyle.SOLID);
//}
//
////Remove all empty rows
//function removeEmptyRows() {
//var ss = SpreadsheetApp.getActive();
//var sheet = ss.getActiveSheet();
//var maxRows = sheet.getMaxRows(); 
//var lastRow = sheet.getLastRow();
//if (maxRows-lastRow != 0){
//      sheet.deleteRows(lastRow+1, maxRows-lastRow);
//      }
//  }
//
//
////Hide all rows before today's date
//function hideRows() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getActiveSheet();
//  var dataRange = sheet.getDataRange();
//  var values = dataRange.getValues();
//  //var sheetDate="";
//  var today = new Date();
//  var dd = today.getDate();
//  var mm = today.getMonth()+1; //January is 0!
//  var yyyy = today.getFullYear();
//  //var todaysDate=dd+"/"+mm+"/"+yyyy;
//var todaysDate=yyyy+"/"+mm+"/"+dd;
// // var ui = SpreadsheetApp.getUi();
// // ui.alert(todaysDate);
//  for (var i = 3; i < values.length; i++) {
//    var row = "";
//    for (var j = 0; j < values[i].length; j++) {  
//      if(j==3)
//      {
//        //ui.alert("Value="+values[i][j]);
//        if(values[i][j]!="")
//        {
//         //ui.alert("In If");
//        var date = new Date(values[i][j]);
//        
//        var sheetDate=date.getFullYear()+"/"+(date.getMonth() + 1)+"/"+date.getDate(); 
//        //ui.alert(sheetDate+" >="+ todaysDate);
//          var d1 = new Date(sheetDate);
//          var d2 = new Date(todaysDate);
//          
//          row = values[i][j+1];
//          
//        if (d1 >= d2) {
//          //ui.alert(d1+" >="+ d2);
//         // ui.alert(i);
//          var k=4;
//          //for(var k=4; k<i; k++)
//          //{
//          try {
//            sheet.hideRows(k,i-3);
//          }
//          catch(err) {
//          }
//            //}
//          return;
//        }
//        } 
//    }    
//  }  
//}
//}
//
//function setWeeksFormat() {
//   var ss = SpreadsheetApp.openById('1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc'); 
//   var sheet = ss.getSheets()[0];                 
//   var cell = sheet.getRange('A4:A999');    
//  
//   cell.setFontFamily('Lato');                 
//   cell.setFontSize('7');  
//   cell.setFontWeight('Bold')
//   cell.setFontColor('#ACC5CF')}
//
//function setEventsFormat() {
//   var ss = SpreadsheetApp.openById('1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc');
//   var sheet = ss.getSheets()[0];                  
//   var cell = sheet.getRange('C4:L999');    
//  
//   cell.setFontFamily('Lato');                 
//   cell.setFontSize('7');  
//   cell.setFontWeight('Normal')
//   cell.setFontColor('#ACC5CF')
//   cell.setBackgroundColor('#FFFFFF')}