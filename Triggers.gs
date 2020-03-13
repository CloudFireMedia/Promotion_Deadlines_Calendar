/**
 * Check event date vs deadlines, send notifications at
 * three days before deadline, 1 day before deadline and day after deadline
 */

function checkDeadlines_() {

  // Get the Promotion Deadlines Calendar data
  // -----------------------------------------
  
  var ss = getSpreadsheet_()
  var sheet = ss.getSheetByName(CDM_SHEET_NAME_)
  
  var values = sheet.getDataRange().getValues()  
  Log_.fine(values)
  
  var testtimezone = ss.getSpreadsheetTimeZone()  
  Log_.fine('testtimezone' + testtimezone)
  
  // Get the staff data
  // ------------------
  
  var staffDataSheetId = Config.get('STAFF_DATA_GSHEET_ID')
  var staffSpreadsheet = SpreadsheetApp.openById(staffDataSheetId)
  var staffDataRange = staffSpreadsheet.getDataRange()
  var staff = getStaff_(staffSpreadsheet)
  
  // remove non-team leaders
  var teamLeads = staff.filter(function(i) {
    return i.isTeamLeader
  })
  
  var teams = teamLeads.reduce(function (out, cur) {
    if (cur.isTeamLeader) {    
      out[cur.team] = {
        name:cur.name,
        email:cur.email
      }
    }
    return out
  }, {})
  
  // Check each row
  // --------------
  
  var today = getMidday_()
  var startRowIndex = sheet.getFrozenRows()
  var numberOfRows = values.length

  Log_.fine('numberOfRows' + numberOfRows)
  for (var rowIndex = startRowIndex; rowIndex < numberOfRows; rowIndex++) {
  
    var rowNumber = rowIndex + 1
  
    //possble values "Yes" or "No". Only process if 'no'
    
    var promoRequested = values[rowIndex][PROMO_INITIATED_COLUMN_INDEX_].toLowerCase()
    
    if (promoRequested !== 'no') {
      Log_.fine('promoRequested not "no", skip this row: ' + promoRequested + ' (' + rowNumber + ')')
      continue 
    }
    
    // Check each tier for a deadline match
    // ------------------------------------
    
    for (var tiersIndex = 0; tiersIndex < TIERS_.length; tiersIndex++) {
    
      var promoType      = TIERS_[tiersIndex]
      var promoIndex     = TIER_DEADLINES_[promoType]
      var promoDeadline  = values[rowIndex][promoIndex]
      
      Log_.fine('rowNumber: ' + rowNumber + '/' + promoType)
      Log_.fine('promoDeadline: ' + promoDeadline)
      
      if (!(promoDeadline instanceof Date)) {
        Log_.fine('promo deadline not date, skip this tier: ' + promoDeadline)
        continue
      }
      
      var dateDiffInDays = Utils.DateDiff.inDays(today, promoDeadline)
      Log_.info('dateDiffInDays' + dateDiffInDays)
      
      //if the data difference is not -1, 1 or 3 then move onto the next tier
      if ([-1, 1, 3].indexOf(dateDiffInDays) === -1) {
        Log_.fine('dateDiffInDays ignored, skip this tier: ' + dateDiffInDays)      
        continue
      }
        
      if (dateDiffInDays === -1 && promoType !== 'Bronze') {
        Log_.fine('-1 dateDiffInDays ignored as not Bronze: ' + dateDiffInDays)            
        continue
      }
      
      sendEmail(teams, values)
            
    } // for each tier
    
  } // for each row
  
  return
  
  // Private Functions
  // -----------------
  
  function sendEmail(teams, values) {
  
    Log_.info('Sending email')
 
    // There is a large difference between the deadlines for the tiers so only one email per day will
    // ever get sent as only one deadline will match
  
    var eventDate      = getEventDate_(values)
    var eventName      = values[rowIndex][EVENT_DATE_COLUMN_INDEX_]
    var sponsorTeamFull    = values[rowIndex][SPONSOR_COLUMN_INDEX_]  

    //Split the team names into individual names
    var sponsorTeamSplit = sponsorTeamFull.split(" + ")
    
    Log_.fine(sponsorTeamSplit)
    
    for(var s in sponsorTeamSplit) {
            
      var sponsorTeam = sponsorTeamSplit[s]
      
      if (sponsorTeam == "N/A") {
        Log_.fine('Team assigned is N/A "' + eventName + '" - do not send email')
        return
      }
    
      var toEmail = teams && teams[sponsorTeam] && teams[sponsorTeam].email
    
      Log_.fine('toEmail' + toEmail)
      Log_.fine('sponsorTeam' + sponsorTeam)

      if (!toEmail) {
        Log_.warning('No recognised team assigned to "' + eventName + '"')
      }
    
      // If toEmail is undefined as no team leader assigned, send email to Communications Director
      toEmail = toEmail || vlookup_('Communications Director', staffDataRange, STAFF_DATA_JOB_TITLE_COLUMN_INDEX_, 8)
    
      if (!toEmail) {
        throw new Error('No recognised team assigned to "' + eventName + '", and no communications director assigned')
      }
        
      var toName = teams && teams[sponsorTeam] && teams[sponsorTeam].name
      toName = toName || vlookup_('Communications Director', staffDataRange, STAFF_DATA_JOB_TITLE_COLUMN_INDEX_, 0)
    
      var body = ''
      var subject = ''
    
      Log_.info('dateDiffInDays: ' + dateDiffInDays + ' (' + rowNumber + ')')
    
      switch (dateDiffInDays) {
        
        case -1: //one day past final (Bronze) deadline - sent to communications director - they've just missed the last chance for any promotion
      
          sheet.getRange(rowNumber, PROMO_INITIATED_COLUMN_INDEX_ + 1).setValue('No')
          var staffRange = SpreadsheetApp.openById(staffDataSheetId).getDataRange()
          to = vlookup_('Communications Director', staffRange, STAFF_DATA_JOB_TITLE_COLUMN_INDEX_, 8)
          subject = CONFIG_.eventsCalendar.emails.expired.subject
          body = CONFIG_.eventsCalendar.emails.expired.body;
          break;
        
        case 1: //day before promoType deadline - sent to team leader
        
          Log_.fine('eventdate' + eventDate)
          Log_.fine('getFormattedDate_eventdate' + getFormattedDate_(eventDate))
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

      var promoFormUrl = ''
    
      if (TEST_GET_FORM_URL_) {
        var promoFormId = Config.get('PROMOTION_REQUEST_FORM_ID');
        promoFormUrl = FormApp.openById(promoFormId).getEditUrl().replace('/edit', '/viewform');
      }
    
      var calendlyUrl = ''
    
      if (TEST_GET_CALENDLY_URL_) {    
        Config.get('ADMIN_CALENDLY_URL');
      }

      body = body
        .replace(/{recipient}/g,       toName )
        .replace(/{sponsorTeam}/g,     sponsorTeam)
        .replace(/{eventDate}/g,       getFormattedDate_(eventDate) )
        .replace(/{eventName}/g,       eventName )
        .replace(/{promoType}/g,       promoType )
        .replace(/{promoDeadline}/g,   Utilities.formatDate(promoDeadline, Session.getScriptTimeZone(), 'E, MMMM d') )
        .replace(/{sponsorSheetUrl}/g, getSponsorTeamSheet() )
        .replace(/{formUrl}/g,         promoFormUrl)
        .replace(/{calendlyUrl}/g,     calendlyUrl);

      MailApp.sendEmail({
        to: toEmail,
        subject: subject,
        htmlBody: body
      });
  
      Log_.info('Email sent (' + rowNumber + '). Subject: ' + subject + '\nto: ' + toEmail + '\nbody: ' + body); 

    } // for (var sponsorTeam in sponsorTeamSplit
    return
          
    // Private Functions
    // -----------------
    
    function getSponsorTeamSheet() {
    
      if (!sponsorTeam) {
        sponsorTeam = CDM_SHEET_NAME_
      }
    
      sponsorSheet = ss.getSheetByName(sponsorTeam)
      
      if (sponsorSheet === null) {
        sponsorSheet = ss.getSheetByName(CDM_SHEET_NAME_)
      } 
    
      var sponsorSheetId = sponsorSheet.getSheetId()
      var sponsorSheetUrl = ss.getUrl() + '#gid=' + sponsorSheetId
      return sponsorSheetUrl
    }
    
  } // checkDeadlines_.sendEmail()
  
} // checkDeadlines_()

function dailyTrigger_() {

  formatCommunicationsDirectorMaster_();
  
  // Run this before checkTeamSheetsForErrors_() in case it affects the other sheets
  checkDeadlines_(); 
  
  if (new Date().getDay() == 1) {
  
    //only run on Mondays
    checkTeamSheetsForErrors_();
  }
}

/**
 * for each team leader
 *  check the team sheet for errors in column C
 *  notify team leader
 */

function checkTeamSheetsForErrors_() {

  var ss = getSpreadsheet_();
  var staff = getStaff_();
  
  //remove non-team leaders
  var teamLeads = staff.filter(function(i) { 
    return i.isTeamLeader
  })
  
  for (var t in teamLeads) {
  
    var team = teamLeads[t].team;
    var sheet = ss.getSheetByName(team);    
    var toEmail = teamLeads[t].email

    Log_.fine('team: ' + team);
    
    if (!toEmail) {
      Log_.warning('No recognised team assigned to "' + eventName + '"')
    }
    
    // If toEmail is undefined as no team leader assigned, send email to Communications Director
    toEmail = toEmail || vlookup_('Communications Director', staffDataRange, STAFF_DATA_JOB_TITLE_COLUMN_INDEX_, 8);
    
    if (!toEmail) {
      throw new Error('No recognised team assigned to "' + eventName + '", and no communications director assigned')
    }    
    
    if (sheet === null) {
    
      var subject = Utilities.formatString('Error in %s on %s', ss.getName(), getFormattedDate_());
      var htmlBody = Utilities.formatString('Unable to find sheet "%s" in <a href="%s">%s</a> for Team Lead "%s <%s>".', 
                                          team, ss.getUrl(), ss.getName(), teamLeads[t].name, teamLeads[t].email);
    
      MailApp.sendEmail({
        to       : toEmail,
        subject  : subject,
        htmlBody : htmlBody,
      })
      
      Log_.info('Email sent to:\nSubject: ' + subject + '\nto: ' + toEmail + '\nbody: ' + htmlBody);             
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
", team, sheetUrl)
      
      MailApp.sendEmail({
        name     : ADMIN_EMAIL_ADDRESS_,
        to       : to,
        subject  : subject,
        htmlBody : body
      })
      
      Log_.info('Email sent to:\nSubject: ' + subject + '\nto: ' + to + '\nbody: ' + body) 
            
    }//next row
  }//next teamlead   
  
} // checkTeamSheetsForErrors_()