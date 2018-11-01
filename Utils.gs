var DateDiff = (function(ns) {

  // Get the number of whole days
  ns.inDays = function(d1, d2) {  
    checkParams(d1, d2)    
    return Math.floor((d2 - d1) / (24 * 3600 * 1000))
  }
  
  ns.inWeeks = function(d1, d2) {  
    checkParams(d1, d2)        
    return parseInt((d2 - d1)/(24 * 3600 * 1000 * 7));
  }
  
  ns.inMonths = function(d1, d2) {
  
    checkParams(d1, d2)    
    
    var d1Y = d1.getFullYear();
    var d2Y = d2.getFullYear();
    var d1M = d1.getMonth();
    var d2M = d2.getMonth();
    
    return (d2M + 12 * d2Y) - (d1M + 12 * d1Y);
  }
  
  inYears: function(d1, d2) {
    checkParams(d1, d2)    
    return d2.getFullYear() - d1.getFullYear();
  }
  
  function checkParams(d1, d2) {
    if (!(d1 instanceof Date) || !(d2 instanceof Date)) {
      throw new Error('DateDiff - bad args. d1: ' + d1 + ', d2:' + d2)
    }
  }
  
  return ns;
  
})(DateDiff || {})

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