///*  
//
//*****************************************************************************
//
//* DESCRIPTION * 
//--
//
//
//* USAGE NOTES *
//--
//
//
//* MAINTENANCE NOTES *
//
//If changes are made to the ordering of the columns in CCN Staff Data Sheet, 
//lines 35-38, 69, 72, 173-177 in this script need to be updated or script
//will not run!!
//
//If one wishes for the script to execute along a different timeline, "+3" in
//tomorrow.setDate(tomorrow.getDate() +3); must be updated.
//
//*****************************************************************************
//*/
//
////var ui = SpreadsheetApp.getUi();
//var spreadsheetUrl = "https://docs.google.com/spreadsheets/d/1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc";
//var cnnStaffDoc="1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI";
//
//
//function getSheets_ErrorReminder() {
//  var spreadsheet = SpreadsheetApp.openById('1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc');
//  //var spreadsheet = SpreadsheetApp.openByUrl(spreadsheetUrl);
//  var sheets = spreadsheet.getSheets();
//  
//  for(var i=0; i<sheets.length; i++) {
//    var spreadsheet = SpreadsheetApp.openById('1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc');
//    var sheets = spreadsheet.getSheets();  
//    for(var i=0; i<sheets.length; i++) {
//      var sheet = sheets[i];  
//      
//      var sheetName=sheet.getName();
//      //ui.alert(sheetName);
//      // return;
//      var staffFromSheet = getStaffFromMainSpreadsheetFour();
//      for(var j = 0; j < staffFromSheet.length; j += 1) {
//        var nameOfStaff=staffFromSheet[j][0];
//        var emailOfStaff=staffFromSheet[j][1];
//        var isTeamLeader1=staffFromSheet[j][2];
//        var staffTeam1=staffFromSheet[j][3];
//        if(staffTeam1==sheetName && isTeamLeader1=="Yes")
//        {
//          var sheetToCheck=spreadsheet.getSheetByName(sheetName);
//          var mysheet = sheetToCheck.getDataRange();
//          var values1 = mysheet.getValues();
//          for (var l = 0; l < values1.length; l++) {
//            if(values1[l][2]=="Error")
//            {
//              //ui.alert(values1[l][2]);
//              if(isTeamLeader1=="No")
//              {
//                var teamLeaderReturn=getTeamLeaderFour(staffTeam1);
//              }
//              
//              //Send Email Function
//              var subject = "Action required!";
//              var body = "{recipient} Leader:<br /><br />One or more of the events on your team's Event Sponsorship Page contains incorrect or incomplete information. PROMOTION WILL NOT BE SCHEDULED FOR YOUR TEAM'S EVENT UNLESS YOU TAKE ACTION. Please visit your team's <a href='{spreadsheetUrl}'>Event Sponsorship Page</a>, find the row(s) highlighted in red, and follow the instructions in the 'Promotion Status' column.<br><br>CCN Communications";
//              
//              if(isTeamLeader1=="No")
//              {
//                //ui.alert("In If"+emailOfStaff);
//                if(emailOfStaff!="")
//                {
//                  var arr=teamLeaderReturn.split(",");
//                  // body = body.replace("{recipient}", nameOfStaff);
//                  body = body.replace("{recipient}", nameOfStaff+"<br> cc "+arr[1]+", Team Leader");
//                  body = body.replace("{spreadsheetUrl}", spreadsheetUrl + "#gid=" + sheetToCheck.getSheetId());
//                  
//                  MailApp.sendEmail({
//                    name:"communications@ccnash.org",
//                    to: emailOfStaff,
//                    cc: arr[0],
//                    subject: subject,
//                    htmlBody: body        
//                  });
//                }
//              }
//              else
//              {
//                //ui.alert(i);
//                //return;
//                //ui.alert("In Else"+emailOfStaff);
//                if(emailOfStaff!="")
//                {
//                  body = body.replace("{recipient}", sheetName);
//                  body = body.replace("{spreadsheetUrl}", spreadsheetUrl + "#gid=" + sheetToCheck.getSheetId());
//                  
//                  MailApp.sendEmail({
//                    name:"communications@ccnash.org",
//                    to: emailOfStaff,
//                    subject: subject,
//                    htmlBody: body        
//                  });          
//                }
//                break;
//              }
//              
//            }
//          }
//        }
//      }      
//    }
//  }
//}
//
//
//function getTeamLeaderFour(staffTeam1)
//{
//    var CcNashStaffSheet = SpreadsheetApp.openById(cnnStaffDoc);
//    var dataRange1 = CcNashStaffSheet.getDataRange();
//    var values1 = dataRange1.getValues();
//    for (var l = 0; l < values1.length; l++) {
//      if(staffTeam1==values1[l][11] && values1[l][12]=="Yes")
//      {
//         return values1[l][8]+','+values1[l][0]+" "+values1[l][1];
//      }
//    }
//}
//
//
//function getStaffFromMainSpreadsheetFour() {
//  var CcNashStaffSheet = SpreadsheetApp.openById(cnnStaffDoc);
//  var lastRow = CcNashStaffSheet.getLastRow();
//  var staffRows = CcNashStaffSheet.getSheetValues(3,1,lastRow, 2);
//  var staffEmails = CcNashStaffSheet.getSheetValues(3,9,lastRow, 1)
//  var staffEmails = CcNashStaffSheet.getSheetValues(3,9,lastRow, 1)
//  var isTeamLeader = CcNashStaffSheet.getSheetValues(3,13,lastRow, 1)
//  var staffTeam = CcNashStaffSheet.getSheetValues(3,12,lastRow, 1)
//  var staffFromSheet = [];
//  
//  for(var row in staffRows) {
//    staffFromSheet.push([staffRows[row][0] + ' ' + staffRows[row][1], staffEmails[row][0], isTeamLeader[row][0], staffTeam[row][0]]);
//  } 
//  return staffFromSheet;;
//}
