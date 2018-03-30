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
//--
//
//*****************************************************************************
//
//*/
//
//
//var spreadsheetUrl = "https://docs.google.com/spreadsheets/d/1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc/edit#gid=0";
////var ui = SpreadsheetApp.getUi();
//function getSheets_OneDayReminder() {
//  var spreadsheet = SpreadsheetApp.openById('1d0-hBf96ilIpAO67LR86leEq09jYP2866uWC48bJloc');
//  var sheets = spreadsheet.getSheets();  
//  for(var i=0; i<sheets.length; i++) {
//    var sheet = sheets[i];  
//    
//    var sheetName=sheet.getName();
//    var staffFromSheet = getStaffFromMainSpreadsheet();
//    for(var j = 0; j < staffFromSheet.length; j += 1) {
//      var nameOfStaff=staffFromSheet[j][0];
//      var emailOfStaff=staffFromSheet[j][1];
//      var isTeamLeader1=staffFromSheet[j][2];
//      var staffTeam1=staffFromSheet[j][3];
//      if(staffTeam1==sheetName && isTeamLeader1=="Yes")
//      {
//        CheckSheetInitial(sheet,emailOfStaff,'Yes',sheetName);              
//      }
//      /* if(nameOfStaff==sheetName)
//      {
//        CheckSheet(sheet,emailOfStaff,isTeamLeader1,staffTeam1);       
//      }*/
//    }    
//  }        
//}
//
//function CheckSheetInitial(sheet,emailOfStaff,isTeamLeader1,staffTeam1) {
//  CheckColumn_OneDay(sheet, 5, "Reminder: Bronze deadline for [ {event date} ] {event}", "Hi {recipient}<br /><br /> This is an automated reminder that the Bronze promotion deadline for your team's event, [ {event date} ] {event}, is tomorrow. This is the last reminder for your event. <br><br> If you would like to schedule promotion for your team's event please submit a <a href='https://docs.google.com/forms/d/e/1FAIpQLSdin9MFXdUjCNqDsy3Th5a82-wgvGlIl6668NUmIggLSFULeg/viewform'>Bronze Promotion Request</a> or schedule a <a href='https://calendly.com/chad_barlow/promo'>Promotion Planning Meeting</a> by {promotion deadline}. <br /><br /> Need more information? Visit your team's <a href='{spreadsheetUrl}'>Event Sponsorship Page</a> or reply to this email.<br><br><b>N.B. Team Leader, please forward this email to the appropriate person on your team whom you would like to initiate a promotion request!</b><br><br>Chad",emailOfStaff,isTeamLeader1,staffTeam1);
//  CheckColumn_OneDay(sheet, 6, "Reminder: Silver deadline for [ {event date} ] {event}", "Hi {recipient}<br /><br /> This is an automated reminder that the Silver promotion deadline for your team's event, [ {event date} ] {event}, is tomorrow. <br><br>If you would like to request Silver promotion for your team's event, please submit a <a href='https://docs.google.com/forms/d/e/1FAIpQLSdin9MFXdUjCNqDsy3Th5a82-wgvGlIl6668NUmIggLSFULeg/viewform'>Silver Promotion Request</a> or schedule a <a href='https://calendly.com/chad_barlow/promo'>Promotion Planning Meeting</a> by {promotion deadline}. If you do not want to request Silver promotion for this event, you may ignore this email. <br /><br /> Need more information? Visit your team's <a href='{spreadsheetUrl}'>Event Sponsorship Page</a> or reply to this email.<br><br><b>N.B. Team Leader, please forward this email to the appropriate person on your team whom you would like to initiate a promotion request!</b><br><br>Chad",emailOfStaff,isTeamLeader1,staffTeam1);
//  CheckColumn_OneDay(sheet, 7, "Reminder: Gold deadline for [ {event date} ] {event}", "Hi {recipient}<br /><br /> This is an automated reminder that the Gold promotion deadline for your team's event, [ {event date} ] {event}, is tomorrow. <br><br>If you would like to request Gold promotion for your team's event, please submit a <a href='https://docs.google.com/forms/d/e/1FAIpQLSdin9MFXdUjCNqDsy3Th5a82-wgvGlIl6668NUmIggLSFULeg/viewform'>Gold Promotion Request</a> or schedule a <a href='https://calendly.com/chad_barlow/promo'>Promotion Planning Meeting</a> by {promotion deadline}. If you do not want to request Gold promotion for this event, you may ignore this email. <br /><br /> Need more information? Visit your team's <a href='{spreadsheetUrl}'>Event Sponsorship Page</a> or reply to this email.<br><br><b>N.B. Team Leader, please forward this email to the appropriate person on your team whom you would like to initiate a promotion request!</b><br><br>Chad",emailOfStaff,isTeamLeader1,staffTeam1);
//}
//
//
//
//function CheckColumn_OneDay(aSheet, column, subject, body,emailOfStaff,isTeamLeader1,staffTeam1) {
//  var subjectSchema=subject;
//  var bodySchema=body;
//     //var ui = SpreadsheetApp.getUi();
//  //ui.alert("Team="+staffTeam1);
//  //return
//  var range = aSheet.getRange(4, 1, aSheet.getMaxRows(),column);
//  
//    //ui.alert(isTeamLeader1);
//  if(isTeamLeader1=="No")
//  {
//    var teamLeaderReturn=getTeamLeaderInitial(staffTeam1);
//  }
//  
//  var data = range.getValues();
//  for (var i = 0; i < data.length; ++i) {
//    var record = data[i];
//    var event = record[1];
//    var eventdate = record[0];
//    var date = record[column-1];
//   
//    //return;
//    var requested = record[2];
//    if (date && (typeof date =="object")) {
//      //ui.alert("In If");
//      var tomorrow = new Date();
//      
//      tomorrow.setDate(tomorrow.getDate() +1);
//      tomorrow.setHours(0,0,0,0);    
//      date.setHours(0,0,0,0);
//      
//      var dateDiff=date - tomorrow;
//       
//      if(dateDiff==0)
//      {
//        //ui.alert("Sheet Date="+date);
//       //ui.alert("tomorrow"+tomorrow);
//        //ui.alert("Sheet Date="+date);
//        //ui.alert("tomorrow"+tomorrow);
//        //ui.alert("Diff="+dateDiff);
//        //ui.alert(requested);
//        //ui.alert(staffTeam1);
//        //ui.alert(record[0]);
//      }
//      if ((date - tomorrow === 0) && requested=="No") {
//        //ui.alert("In If");
//        subject = subjectSchema.replace("{event}", event);
//        subject = subject.replace("{event date}", Utilities.formatDate(eventdate,"GMT-5", "MM.dd"));
//        
//        if(isTeamLeader1=="No")
//        {
//          if(emailOfStaff!="")
//          {
//            var arr=teamLeaderReturn.split(",");
//            body = bodySchema.replace("{event}", event);
//            body = body.replace("{event date}", Utilities.formatDate(eventdate,"GMT-5", "MM.dd"));
//            body = body.replace("{promotion deadline}", Utilities.formatDate(date,"GMT-5", "MM.dd"));
//            body = body.replace("{recipient}", aSheet.getName()+"<br> cc "+arr[1]+", Team Leader");
//            body = body.replace("{spreadsheetUrl}", spreadsheetUrl + "#gid=" + aSheet.getSheetId());
//            
//            MailApp.sendEmail({
//              name:"communications@ccnash.org",
//              to: emailOfStaff,
//              cc: arr[0],
//              subject: subject,
//              htmlBody: body        
//            });
//          }
//        }
//        else
//        {
//          //ui.alert("In Else"+emailOfStaff);
//          if(emailOfStaff!="")
//          {
//            body = bodySchema.replace("{event}", event);
//            body = body.replace("{event date}", Utilities.formatDate(eventdate,"GMT-5", "MM.dd"));
//            body = body.replace("{promotion deadline}", Utilities.formatDate(date,"GMT-5", "MM.dd"));
//            body = body.replace("{recipient}", aSheet.getName());
//            body = body.replace("{spreadsheetUrl}", spreadsheetUrl + "#gid=" + aSheet.getSheetId());
//            
//            MailApp.sendEmail({
//              name:"communications@ccnash.org",
//              to: emailOfStaff,
//              subject: subject,
//              htmlBody: body        
//            });          
//          }
//        }
//      }
//    }
//  }
//}
//
//
//function getTeamLeaderInitial(staffTeam1)
//{
//  //var ui = SpreadsheetApp.getUi();
//  // ui.alert("in if");
//  //ui.alert(staffTeam1);
//  var CcNashStaffSheet = SpreadsheetApp.openById('1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI');
//  var dataRange1 = CcNashStaffSheet.getDataRange();
//  var values1 = dataRange1.getValues();
//  //ui.alert(values1);
//  for (var l = 0; l < values1.length; l++) {
//    //ui.alert(staffTeam1+"=="+values1[l][11]);
//    if(staffTeam1==values1[l][11] && values1[l][12]=="Yes")
//    {
//      //ui.alert("Found");
//      return values1[l][8]+','+values1[l][0]+" "+values1[l][1];
//    }
//  }
//  
//}
//
//
//
//function getStaffFromMainSpreadsheet() {
//  var CcNashStaffSheet = SpreadsheetApp.openById('1iiFmdqUd-CoWtUjZxVgGcNb74dPVh-l5kuU_G5mmiHI');
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
