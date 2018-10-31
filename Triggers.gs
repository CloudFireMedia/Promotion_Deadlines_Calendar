/**
 * Check event date vs deadlines, send notifications at
 * three days before deadline, 1 day before deadline and day after deadline
 */

function checkDeadlines_() {

  var ss = SpreadsheetApp.getActive();
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
        bblogFine_('promodeadline not date, skip this tier: ' + promoDeadline)
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
        
        subject = Utilities.formatString(CONFIG_.eventsCalendar.emails.oneDay.subject, promoType, fDate_LOCAL(eventDate), eventName);;
        body = CONFIG_.eventsCalendar.emails.oneDay.body;
        break;
        
      case 3: //3 days before promoType deadline
      
        subject = Utilities.formatString(CONFIG_.eventsCalendar.emails.threeDays.subject, promoType, fDate_LOCAL(eventDate), eventName);;
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
      .replace(/{eventDate}/g,      fDate_LOCAL(eventDate) )
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
  formatSheet_();
  checkDeadlines_();//run this before checkTeamSheetsForErrors_() in case it affects the other sheets 
  if(new Date().getDay() == 1)//only run on Mondays
    checkTeamSheetsForErrors_();
}

function checkTeamSheetsForErrors_() {

  //get staff
  //for each team leader
  //check the team sheet for errors in column C
  //notify person on sheet and cc team leader (if event owner)
  
  var promoCalendarSpreadsheet = SpreadsheetApp.getActive();
  var staff = getStaff_();
  var teamLeads = staff.filter(function(i){return i.isTeamLeader})//remove non-team leaders
  for(var t in teamLeads){
    var team = teamLeads[t].team;
    var sheet = promoCalendarSpreadsheet.getSheetByName(team);
    if(!sheet){
    
      MailApp.sendEmail({
        to       : ADMIN_EMAIL_ADDRESS_,
        subject  : Utilities.formatString('Error in %s on %s', ss.getName(), fDate_LOCAL()),
        htmlBody : Utilities.formatString('Unable to find sheet "%s" in <a href="%s">%s</a> for Team Lead "%s <%s>"', 
                                          team, ss.getUrl(), ss.getName(), teamLeads[t].name, teamLeads[t].email)
      })
      continue;
    }
    
    //sheet found, get data and check for errors
    var data = sheet.getDataRange().getValues();
    for(var row in data) {
      var promoRequested = data[row][2];
      if(promoRequested.toLowerCase() != "error") continue;//all good, next row please
      
      //else, let team leader know about the error
      var to = teamLeads[t].email;
      var sheetUrl = ss.getUrl() + "#gid=" + sheet.getSheetId();
      var subject = "Action required!";
      var body = Utilities.formatString("\
%s Leader:<br><br>One or more of the events on your team's Event Sponsorship Page contains incorrect or incomplete information. PROMOTION WILL NOT BE SCHEDULED FOR YOUR TEAM'S EVENT UNLESS YOU TAKE ACTION. Please visit your team's <a href='\
%s'>Event Sponsorship Page</a>, find the row(s) highlighted in red, and follow the instructions in the 'Promotion Status' column.<br><br>CCN Communications\
", team, sheetUrl);
      
      MailApp.sendEmail({
        name     : ADMIN_EMAIL_ADDRESS_,
        to       : to,
        subject  : subject,
        htmlBody : body
      });
    }//next row
  }//next teamlead    
}

function getStaff_(spreadsheet) {//returns [{},{},...]

  if (spreadsheet === undefined) {
    staffSpreadsheetId = Config.get('STAFF_DATA_GSHEET_ID')
    spreadsheet = SpreadsheetApp.openById(staffSpreadsheetId)
  } else {}

  var sheet = spreadsheet.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  values = values.slice(sheet.getFrozenRows());//remove headers if any
  
  var staff = values.map(function(c,i,a){
    return {
      name         : [c[0], c[1]].join(' '),
      email        : c[8],
      team         : c[11],
      isTeamLeader : (c[12].toLowerCase()=='yes'),
      jobTitle     : c[4],
    };
  },[]);

  return staff;
}

function fDate_LOCAL(date, format){ 
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'MM.dd')
}

function getMidnight_(date){
  date = date || new Date();//set default today if not provided
  date = new Date(date);//don't change original date
  date.setHours(0,0,0,0);//set time to midnight
  return date;
}

function vLookup_(needle, range, searchOffset, returnOffset) {

  if (typeof searchOffset == 'undefined') {
    searchOffset = 0;// what column to search, default to first
  }
  
  if (typeof returnOffset == 'undefined') {
    returnOffset = 1;// what column to return, default to second
  }
  
  var haystack = range.getValues();
  
  for(var i=0; i<haystack.length; i++) {
  
    if(haystack[i][searchOffset] && haystack[i][searchOffset].toLowerCase() == needle.toLowerCase()) {
      return haystack[i][returnOffset];
    }
  }
  
} // vLookup_()

