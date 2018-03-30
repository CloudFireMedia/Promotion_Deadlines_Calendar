////emails Communications Director if a Bronze deadline has passed with no action from the sponsor
//
//var ccnStaffDataSheet="1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc";
////var ui = SpreadsheetApp.getUi();
//
//function ChangeValues(rangeString) {
// var ss = SpreadsheetApp.getActiveSpreadsheet();
//    var sheet = ss.getActiveSheet();
//    var dataRange = sheet.getDataRange();
//    var values = dataRange.getValues();
//  
//    for (var i = 3; i < values.length; i++) {
//      var row = "";
//      if(i>=3)
//       {
//         var dateColumn=values[i][8];
//         var gColumn=values[i][6];
//         var EventName=values[i][4];
//         var todaysDate = new Date();
//         
//         //ui.alert(dateColumn.getTime());
//         //ui.alert(todaysDate.getTime());
//           if(isDate(dateColumn))
//           {
//             //return;
//             if(dateColumn.getTime()<todaysDate.getTime() && gColumn=="No")
//             {
//               var iUpdate=i+1;
//               sheet.getRange(iUpdate,7).setValue('N/A');             
//               updateRecordsOnEdit(i);
//               getCommunicationDirector(values[i][4],values[i][7]);
//			   
//             }
//           }
//         //return;
//       }
//      }
//}
//
//function isDate(dateString) {
//  return (new Date(dateString)) != "Invalid Date"; // notice not !==
//}
//
//function getCommunicationDirector(eventName,staffSponsor)
//{
//  //ui.alert(eventName+"Send Mail");
//  // ui.alert(staffSponsor+"Send Mail");
//  var CcNashStaffSheet = SpreadsheetApp.openById(ccnStaffDataSheet);
//  var lastRow = CcNashStaffSheet.getLastRow();
//  var staffRows = CcNashStaffSheet.getSheetValues(3,1,lastRow, 2);
//  var staffEmails = CcNashStaffSheet.getSheetValues(3,9,lastRow, 1);
//  var staffTitle = CcNashStaffSheet.getSheetValues(3,5,lastRow, 1);
//  for(var row in staffRows) {
//    if(staffTitle[row][0]=="Communications Director")
//    {
//      var emailTo=staffEmails[row][0];
//      emailTo=emailTo.trim();
//      //ui.alert(emailTo);
//      //Mail Sending Script
//      var subject="Attention: a final promotion deadline has lapsed"
//      var body="This is an automated notice that the Bronze deadline for {{event}} has lapsed without action by the staff sponsor, {{staffSponsor}}";
//      body = body.replace("{{event}}", eventName);
//      body = body.replace("{{staffSponsor}}", staffSponsor);
//      //ui.alert(subject);
//     // ui.alert(body);
//      MailApp.sendEmail({
//        name:"chbarlow@gmail.com",
//        to: emailTo,
//        subject: subject,
//        htmlBody: body        
//      });  
//    }
//    //staffFromSheet.push([staffRows[row][0] + ' ' + staffRows[row][1], staffEmails[row][0]]);
//  }
//}
//
//
//
//
//
//
//function updateRecordsOnEdit(i) {
//  
//  //var ui = SpreadsheetApp.getUi();
//  //ui.alert(i);
// // return;
//   var ss = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = ss.getActiveSheet();
//   var dataRange = sheet.getDataRange();
//   var values = dataRange.getValues();
//  
//  var fColumn=values[i][5];
//  var gColumn=values[i][6];
//  //ui.alert(gColumn);
// //return;
//  if(gColumn=="No")
//  {
//    var iUpdate=i+1;
//    sheet.getRange(iUpdate,12).setValue('Awaiting promotion request');
//  }
//  else if(gColumn=="Yes" && fColumn=="Yes")
//  {
//    var iUpdate=i+1;
//    sheet.getRange(iUpdate,12).setValue('Promotion Scheduled');
//  }
//  else if(gColumn=="Yes" && (fColumn=="N/A" || fColumn=="No"))  
//  {
//    var iUpdate=i+1;
//    sheet.getRange(iUpdate,12).setValue('Pending review');
//  }
//  else if(gColumn=="N/A")
//  {
//    var iUpdate=i+1;
//    sheet.getRange(iUpdate,12).setValue('N/A'); 
//    sheet.getRange(iUpdate,3).setValue('N/A'); 
//    sheet.getRange(iUpdate,6).setValue('N/A'); 
//  }
//           
//    }    
