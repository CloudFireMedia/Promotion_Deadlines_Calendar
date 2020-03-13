function getSpreadsheet_() {
  var spreadsheet = SpreadsheetApp.getActive(); 
  if (spreadsheet === null) {
    Log_.fine('no spreadsheet')
    if (!PRODUCTION_VERSION_) {
      spreadsheet = SpreadsheetApp.openById(TEST_PDC_SPREADSHEET_ID_)
    } else {
      throw new Error('No active spreadsheet');
    }
  }
  return spreadsheet
}

function getFormattedDate_(date, format) { 
  date = date || new Date();
  format = format || 'MM.dd';
  return Utilities.formatDate(date, Session.getScriptTimeZone(), format);
}

function getMidnight_(date){
  date = date || new Date();//set default today if not provided
  date = new Date(date);//don't change original date
  date.setHours(0,0,0,0);//set time to midnight
  return date;
}

function formatEventDate_(values, rowIndex) {
  
  // Add 2 hours to the eventDate as the timezone used in getValues uses a different timezone (GMT-5) than the spreadsheet, 
  // (GMT-6) causing the date to change to the day before. Adding 2 hours stops the date being converted to the day before.
  var eventDateOriginal = values[rowIndex][START_DATE_COLUMN_INDEX_]
    
  const inc = 2 * 1000 * 60 * 60 // 2 hours
  var eventDate = new Date( eventDateOriginal.getTime() + inc )
  Log_.fine('original eventDate' + eventDateOriginal)
  Log_.fine('incremented eventDate' + eventDate)
  
  return eventDate
}

function vlookup_(needle, range, searchOffset, returnOffset) {

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
  
} // vlookup_()

function getStaff_(spreadsheet) {//returns [{},{},...]

  if (spreadsheet === undefined) {
    staffSpreadsheetId = Config.get('STAFF_DATA_GSHEET_ID')
    spreadsheet = SpreadsheetApp.openById(staffSpreadsheetId)
  } 

  var sheet = spreadsheet.getSheetByName(SPONSOR_SHEET_NAME_);
  
  var values = sheet.getDataRange().getValues();
  values = values.slice(sheet.getFrozenRows()); //remove headers if any
  
  var staff = values.map(function(nextStaffMember) {
    return {
      name         : [nextStaffMember[STAFF_DATA_FIRST_NAME_COLUMN_INDEX_], nextStaffMember[STAFF_DATA_LAST_NAME_COLUMN_INDEX_]].join(' '),
      email        : nextStaffMember[STAFF_DATA_EMAIL_COLUMN_INDEX_],
      team         : nextStaffMember[STAFF_DATA_TEAM_COLUMN_INDEX_],
      isTeamLeader : (nextStaffMember[STAFF_DATA_LEADER_COLUMN_INDEX_].toLowerCase() === 'yes'),
      jobTitle     : nextStaffMember[STAFF_DATA_JOB_TITLE_COLUMN_INDEX_],
    };
  },[]);

  return staff;
}

function toast_(msg, title, timeoutSeconds) {
  var spreadsheet = SpreadsheetApp.getActive()
  if (spreadsheet !== null) {
    var localTitle = title || ''
    var localTimeout = timeoutSeconds || null  
    spreadsheet.toast(msg, localTitle, localTimeout)
  }
}

function msgBox_(msg) {
  var spreadsheet = SpreadsheetApp.getActive()
  if (spreadsheet !== null) {  
    Browser.msgBox(msg)
  }
}