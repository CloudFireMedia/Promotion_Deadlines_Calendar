// 34567890123456789012345678901234567890123456789012345678901234567890123456789

// JSHint - TODO
/* jshint asi: true */

(function() {"use strict"})()

// Code review all files - TODO
// JSHint review (see files) - TODO
// Unit Tests - TODO
// System Test (Dev) - TODO
// System Test (Prod) - TODO

// Config.gs
// =========
//
// All the constants and configuration settings

// Configuration
// =============

var SCRIPT_NAME = "Promotion Deadliness Calendar";
var SCRIPT_VERSION = "v1.3.dev_ajr";

var PRODUCTION_VERSION_ = false;

// Log Library
// -----------

var DEBUG_LOG_LEVEL_ = PRODUCTION_VERSION_ ? BBLog.Level.INFO : BBLog.Level.FINER;
var DEBUG_LOG_DISPLAY_FUNCTION_NAMES_ = PRODUCTION_VERSION_ ? BBLog.DisplayFunctionNames.NO : BBLog.DisplayFunctionNames.NO;

// Assert library
// --------------

var SEND_ERROR_EMAIL_ = PRODUCTION_VERSION_ ? true : false;
var HANDLE_ERROR_ = Assert.HandleError.THROW;
var ADMIN_EMAIL_ADDRESS_ = 'dev@cloudfire.media';

// Tests
// -----

var TEST_PDC_SPREADSHEET_ID_ = '1Zn4IFvkahOa2mlR8pVWgbmD6icuOQzZAniiVZA6jrCQ' // Test Copy of Promotion Deadlines Calendar v1.3 (live)

// Constants/Enums
// ===============

// Staff data sheet

var STAFF_DATA_FIRST_NAME_COLUMN_INDEX_   = 0;
var STAFF_DATA_LAST_NAME_COLUMN_INDEX_    = 1;
var STAFF_DATA_BIRTHDAY_COLUMN_INDEX_     = 2;
// var STAFF_DATA_LAST_NAME_COLUMN_INDEX_ = 3; // Leagacy
var STAFF_DATA_JOB_TITLE_COLUMN_INDEX_    = 4;
var STAFF_DATA_STATUS_COLUMN_INDEX_       = 5;
var STAFF_DATA_EXTENSION_COLUMN_INDEX_    = 6;
var STAFF_DATA_CELL_COLUMN_INDEX_         = 7;
var STAFF_DATA_EMAIL_COLUMN_INDEX_        = 8;
var STAFF_DATA_SECOND_PHONE_COLUMN_INDEX_ = 9;
var STAFF_DATA_TEAM_COLUMN_INDEX_         = 11;
var STAFF_DATA_LEADER_COLUMN_INDEX_       = 12;
var STAFF_DATA_MEMBER_COLUMN_INDEX_       = 13;

// Promotion Deadslines Calendar
var START_DATE_COLUMN_INDEX_      = 3;
var EVENT_DATE_COLUMN_INDEX_      = 4;
var PROMO_INITIATED_COLUMN_INDEX_ = 6;
var SPONSOR_COLUMN_INDEX_         = 7;

var DATA_SHEET_NAME_ = 'Communications Director Master';
var SPONSOR_SHEET_NAME_ = 'Staff';
var TIER_DUEDATE_SHEET_NAME_ ='Lookup: Tier Due Dates'; 

var TIERS_ = ['Gold','Silver','Bronze'];
var TIER_DEADLINES_ = {Bronze: 8, Silver: 9, Gold: 10}; // column index for matching promo type

var CONFIG_ = {

  eventsCalendar : {
  
    emails : {
    
      expired : { //one day past final (Bronze) deadline - sent to communications director - they've just missed the last chance for any promotion
        subject : "Attention: a final promotion deadline has lapsed",
        body : "\
This is an automated notice that the \
{promoType} deadline for \
{eventName} has lapsed without action by the staff sponsor, \
{staffSponsor}",
      },
      
      oneDay : { //day before promoType deadline - sent to team leader
        subject : "Reminder: %s deadline for [ %s ] %s",
        body : "\Hi \
{recipient},<br><br>This is an automated reminder that the \
{promoType} \
promotion deadline for your team's event, [ \
{eventDate} ] \
{eventName}, is tomorrow. This is the last reminder for a {promoType} promotion for your event.<br><br>If you would like to schedule promotion for your team's event please submit a <a href='\
{formUrl}\'>{promoType} Promotion Request</a> or schedule a <a href='{calendlyUrl}'>Promotion Planning Meeting</a> by \
{promoDeadline}. If you do not want to request \
{promoType} promotion for this event, you may ignore this email. </b><br><br>Promotions Admin<br><br>--<br>P.S. If appropriate, please forward this email to the team member responsible for initiating a promotion request. <br><br> Need more information? Visit your team's <a href='\
{spreadsheetUrl}'>Event Sponsorship Page</a> or reply to this email.\
",
      },
      
      threeDays : { //3 days before promoType deadline
        subject : "Reminder: %s deadline for [ %s ] %s",
        body : "Hi \
{recipient},<br><br> This is an automated reminder that the \
{promoType} promotion deadline for your team's event, [ \
{eventDate} ] \
{eventName}, is in three days. <br><br>If you would like to request \
{promoType} promotion for your team's event, please submit a <a href='\
{formUrl}\
'>Promotion Request</a> or schedule a <a href='{calendlyUrl}'>Promotion Planning Meeting</a> by \
{promoDeadline}. If you do not want to request \
{promoType} promotion for this event, you may ignore this email. </b><br><br>Promotions Admin<br><br>--<br>P.S. If appropriate, please forward this email to the team member responsible to initiating a promotion request. <br><br> Need more information? Visit your team's <a href='\
{spreadsheetUrl}'>Event Sponsorship Page</a> or reply to this email.\
",
      },
    },
  },
  
} // CONFIG_ 

// Function Template
// -----------------

/**
 *
 *
 * @param {object} 
 *
 * @return {object}
 */
/* 
function functionTemplate() {

  Log_.functionEntryPoint()
  
  

} // functionTemplate() 
*/