/**
 * Check event date vs deadlines, send notifications at
 * three days before deadline, 1 day before deadline and day after deadline
 */

function checkDeadlines_() {

  // Get the Promotion Deadlines Calendar data
  // -----------------------------------------

  var ss = SpreadsheetApp.getActive();
  
  if (ss === null && !PRODUCTION_VERSION_) {
    ss = SpreadsheetApp.openById(TEST_PDC_SPREADSHEET_ID_);
  } else {
    throw new Error('No active spreadsheet');
  }
  
  var spreadsheetUrl = ss.getUrl();
  var sheet = ss.getSheetByName(DATA_SHEET_NAME_);
  var values = sheet.getDataRange().getValues();
  
  // Get the staff data
  // ------------------
  
  var staffDataSheetId = Config.get('STAFF_DATA_GSHEET_ID');
  var staffSpreadsheet = SpreadsheetApp.openById(staffDataSheetId);
  var staffDataRange = staffSpreadsheet.getDataRange();
  var staff = getStaff_(staffSpreadsheet);
  
  // remove non-team leaders
  var teamLeads = staff.filter(function(i) {
    return i.isTeamLeader;
  })
  
  var teams = teamLeads.reduce(function (out, cur) {
    if (cur.isTeamLeader) {    
      out[cur.team] = {
        name:cur.name,
        email:cur.email
      };
    }
    return out;
  }, {});
  
  // Check each row
  // --------------
  
  var today = getMidnight_();
  var startRowIndex = sheet.getFrozenRows();
  var numberOfRows = values.length;

  for (var rowIndex = startRowIndex; rowIndex < numberOfRows; rowIndex++) {
  
    var rowNumber = rowIndex + 1;
  
    //possble values "Yes" or "No". Only process if 'no'
    
    var promoRequested = values[rowIndex][PROMO_INITIATED_COLUMN_INDEX_].toLowerCase()
    
    if (promoRequested !== 'no') {
      bblogFine_('promoRequested not "no", skip this row: ' + promoRequested + ' (' + rowNumber + ')')
      continue; 
    }
    
    // Check each tier for a deadline match
    // ------------------------------------
    
    for (var tiersIndex = 0; tiersIndex < TIERS_.length; tiersIndex++) {
    
      var promoType      = TIERS_[tiersIndex];
      var promoIndex     = TIER_DEADLINES_[promoType];
      var promoDeadline  = values[rowIndex][promoIndex];
      
      bblogFine_('rowNumber: ' + rowNumber + '/' + promoType)
      
      if (!(promoDeadline instanceof Date)) {
        bblogFine_('promo deadline not date, skip this tier: ' + promoDeadline)
        continue;
      }
      
      var dateDiffInDays = DateDiff.inDays(today, promoDeadline);
      
      //if the data difference is not -1, 1 or 3 then move onto the next tier
      if ([-1, 1, 3].indexOf(dateDiffInDays) === -1) {
        bblogFine_('dateDiffInDays ignored, skip this tier: ' + dateDiffInDays)      
        continue;
      }
        
      if (dateDiffInDays === -1 && promoType !== 'Bronze') {
        bblogFine_('-1 dateDiffInDays ignored as not Bronze: ' + dateDiffInDays)            
        continue;
      }
      
      sendEmail();
            
    } // for each tier
    
  } // for each row
  
  return
  
  // Private Functions
  // -----------------
  
  function sendEmail() {
  
    bblogFine_('Sending email');
  
    // The is a large difference between the deadlines for the tiers so only one email per day will
    // ever get sent as only one deadline will match
  
    var eventDate      = values[rowIndex][START_DATE_COLUMN_INDEX_];
    var eventName      = values[rowIndex][EVENT_DATE_COLUMN_INDEX_];
    var staffSponsor   = values[rowIndex][SPONSOR_COLUMN_INDEX_]; //this is the sponsoring team, not the person
    
    var to = teams && teams[staffSponsor] && teams[staffSponsor].email;
    to = to || vlookup_('Communications Director', staffDataRange, STAFF_DATA_JOB_TITLE_COLUMN_INDEX_, 8);
        
    var toName = teams && teams[staffSponsor] && teams[staffSponsor].name;
    toName = toName || vlookup_('Communications Director', staffDataRange, STAFF_DATA_JOB_TITLE_COLUMN_INDEX_, 0);
    
    var body = '';
    var subject = '';
    
    bblogInfo_('dateDiffInDays: ' + dateDiffInDays + ' (' + rowNumber + ')');
    
    switch (dateDiffInDays) {
        
      case -1: //one day past final (Bronze) deadline - sent to communications director - they've just missed the last chance for any promotion
      
        sheet.getRange(rowNumber, PROMO_INITIATED_COLUMN_INDEX_ + 1).setValue('No');
        var staffRange = SpreadsheetApp.openById(staffDataSheetId).getDataRange();
        to = vlookup_('Communications Director', staffRange, STAFF_DATA_JOB_TITLE_COLUMN_INDEX_, 8);
        subject = CONFIG_.eventsCalendar.emails.expired.subject;
        body = CONFIG_.eventsCalendar.emails.expired.body;
        break;
        
      case 1: //day before promoType deadline - sent to team leader
        
        subject = Utilities.formatString(CONFIG_.eventsCalendar.emails.oneDay.subject, promoType, getFormattedDate_(eventDate), eventName);;
        body = CONFIG_.eventsCalendar.emails.oneDay.body;
        break;
        
      case 3: //3 days before promoType deadline
      
        subject = Utilities.formatString(CONFIG_.eventsCalendar.emails.threeDays.subject, promoType, getFormattedDate_(eventDate), eventName);;
        body = CONFIG_.eventsCalendar.emails.threeDays.body;
        break;
        
      default: 
        throw new Error('Bad date difference');
    }

    var promoFormId = Config.get('PROMOTION_REQUEST_FORM_ID');
    var promoFormUrl = FormApp.openById(promoFormId).getEditUrl().replace('/edit', '/viewform');
    var calendlyUrl = Config.get('ADMIN_CALENDLY_URL');

    body = body
      .replace(/{recipient}/g,      toName )
      .replace(/{staffSponsor}/g,   staffSponsor )
      .replace(/{eventDate}/g,      getFormattedDate_(eventDate) )
      .replace(/{eventName}/g,      eventName )
      .replace(/{promoType}/g,      promoType )
      .replace(/{promoDeadline}/g,  Utilities.formatDate(promoDeadline, Session.getScriptTimeZone(), 'E, MMMM d') )
      .replace(/{spreadsheetUrl}/g, spreadsheetUrl )
      .replace(/{formUrl}/g,        promoFormUrl)
      .replace(/{calendlyUrl}/g,    calendlyUrl);

    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: body
    });
  
    bblogInfo_('Email sent (' + rowNumber + '). Subject: ' + subject + '\nto: ' + to + '\nbody: ' + body); 

  } // checkDeadlines_.sendEmail()
  
} // checkDeadlines_()

function onEdit_eventsCalendar(e) {

  var range = e.range;
  
  //check for exit conditions
  var sheet = range.getSheet();
  if(sheet.getName() != DATA_SHEET_NAME_) return;//only run on dataSheetName
  var col = range.getColumn();
  var width = range.getWidth();
  var colsInRange = Array.apply(null, Array(width)).map(function(c,i){return col+i;});//return array of col numbers across range like [2,3,4]
  if(colsInRange.indexOf(6) <0 && colsInRange.indexOf(7) <0 ) return;//edit was not in a column we are watching, skedaddle
  
  //ok, we're good to go
  var row = range.getRow();
  var height = range.getHeight();
  var data = sheet.getRange(row, 1, height, sheet.getLastColumn()).getValues();
  
  for(var r in data){//handle multi-row edits (like paste or move)
    var onWebCal = data[r][6-1].toUpperCase();
    var promoRequested = data[r][7-1].toUpperCase();
    if(onWebCal=="N/A" && promoRequested=="N/A"){
      var promoTypeRange = sheet.getRange(row+parseInt(r),3);
      promoTypeRange.setValue('N/A');//Promo Type
      blinkRange(promoTypeRange);
    }
  }
  
} // onEdit_eventsCalendar()

function dailyTrigger_(){
//  formatSheet_(); // TODO - https://trello.com/c/zBmW3avV
  checkDeadlines_();//run this before checkTeamSheetsForErrors_() in case it affects the other sheets 
  if(new Date().getDay() == 1)//only run on Mondays
    checkTeamSheetsForErrors_();
}

/**
 * for each team leader
 *  check the team sheet for errors in column C
 *  notify person on sheet and cc team leader (if event owner)
 */

function checkTeamSheetsForErrors_() {

  var ss = SpreadsheetApp.getActive();
  var staff = getStaff_();
  
  //remove non-team leaders
  var teamLeads = staff.filter(function(i) { 
    return i.isTeamLeader
  })
  
  for (var t in teamLeads) {
  
    var team = teamLeads[t].team;
    var sheet = ss.getSheetByName(team);
    
    bblogFine_('team: ' + team);
    
    if (sheet === null) {
    
      var subject = Utilities.formatString('Error in %s on %s', ss.getName(), getFormattedDate_());
      var htmlBody = Utilities.formatString('Unable to find sheet "%s" in <a href="%s">%s</a> for Team Lead "%s <%s>".', 
                                          team, ss.getUrl(), ss.getName(), teamLeads[t].name, teamLeads[t].email);
    
      MailApp.sendEmail({
        to       : ADMIN_EMAIL_ADDRESS_,
        subject  : subject,
        htmlBody : htmlBody,
      })
      
      bblogInfo_('Email sent to:\nSubject: ' + subject + '\nto: ' + ADMIN_EMAIL_ADDRESS_ + '\nbody: ' + htmlBody);             
      continue;
    }
    
    //sheet found, get data and check for errors
    var data = sheet.getDataRange().getValues();
    
    for (var row in data) {
    
      var promoRequested = data[row][2];
      
      if (promoRequested.toLowerCase() !== "error") {
        continue;
      }
      
      //else, let team leader know about the error
      var to = teamLeads[t].email;
      var sheetUrl = ss.getUrl() + "#gid=" + sheet.getSheetId();
      var subject = "Action required!";
      var body = Utilities.formatString("\
%s Leader:<br><br>One or more of the events on your team's Event Sponsorship Page contains incorrect or incomplete information. PROMOTION WILL NOT BE SCHEDULED FOR YOUR TEAM'S EVENT UNLESS YOU TAKE ACTION. Please visit your team's <a href='\
%s'>Event Sponsorship Page</a>, find the row(s) highlighted in red, and follow the instructions in the 'Promotion Status' column.<br><br>Promotion Admin\
", team, sheetUrl);
      
      MailApp.sendEmail({
        name     : ADMIN_EMAIL_ADDRESS_,
        to       : to,
        subject  : subject,
        htmlBody : body
      });
      
      bblogInfo_('Email sent to:\nSubject: ' + subject + '\nto: ' + to + '\nbody: ' + body); 
            
    }//next row
  }//next teamlead   
  
} // checkTeamSheetsForErrors_()