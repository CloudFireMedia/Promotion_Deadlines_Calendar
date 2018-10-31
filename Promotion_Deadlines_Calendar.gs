// JSHint - TODO
/* jshint asi: true */

(function() {"use strict"})()

// Promotion_Deadlines_Calendar.gs
// ===============================
//
// External interface to this script - all of the event handlers.
//
// This files contains all of the event handlers, plus miscellaneous functions 
// not worthy of their own files yet
//
// The filename is prepended with _API as the Github chrome extension won't 
// push a file with the same name as the project.

var Log_

// Public event handlers
// ---------------------
//
// All external event handlers need to be top-level function calls; they can't 
// be part of an object, and to ensure they are all processed similarily 
// for things like logging and error handling, they all go through 
// errorHandler_(). These can be called from custom menus, web apps, 
// triggers, etc
// 
// The main functionality of a call is in a function with the same name but 
// post-fixed with an underscore (to indicate it is private to the script)
//
// For debug, rather than production builds, lower level functions are exposed
// in the menu

var EVENT_HANDLERS_ = {

//                           Name                            onError Message - "Failed to..."  Main Functionality
//                           ----                            ---------------                   ------------------

  addRowBelow:               ['addRowBelow()',               'add row below',                  addRowBelow_],
  deleteRow:                 ['deleteRow()',                 'delete row',                     deleteRow_],
  calculateDeadlines:        ['calculateDeadlines()',        'calculate deadlines',            calculateDeadlines_],
  newEventPopup:             ['newEventPopup()',             'show event popup',               newEventPopup_],
  prepareNewYearsData:       ['prepareNewYearsData()',       'prepare New Years data',         prepareNewYearsData_],
  repopulateBulletins_:      ['repopulateBulletins_()',      'repopulate bulletins',           repopulateBulletins_],  
  formatSheet:               ['formatSheet()',               'format sheet',                   formatSheet_],
  hideRows:                  ['hideRows()',                  'hide rows',                      hideRows_],
  removeEmptyRows:           ['removeEmptyRows()',           'remove empty rows',              removeEmptyRows_],
  colorBorders:              ['colorBorders()',              'color borders',                  colorBorders_],
  setWeeksFormat:            ['setWeeksFormat()',            'setWeeksFormat',                 setWeeksFormat_],
  setEventsFormat:           ['setEventsFormat()',           'set events format',              setEventsFormat_],
  restoreHeader:             ['restoreHeader()',             'restore header',                 restoreHeader_],
  backupHeader:              ['backupHeader()',              'backup header',                  backupHeader_],
  setupAutomation:           ['setupAutomation()',           'setup automation',               setupAutomation_],
  disableAutomation:         ['disableAutomation()',         'disable automation',             disableAutomation_],
  checkDeadlines:            ['checkDeadlines()',            'check deadlines',                checkDeadlines_],
  checkTeamSheetsForErrors:  ['checkTeamSheetsForErrors()',  'check Team Sheets For Errors',   checkTeamSheetsForErrors_],
  dailyTrigger:              ['dailyTrigger()',              'run daily trigger',              dailyTrigger_],
  fillSponsor:               ['fillSponsor()',               'fill sponsor in popup',          fillSponsor_],
  fillTier:                  ['fillTier()',                  'fill tiers in popup',            fillTier_],
}

function addRowBelow(args)              {return eventHandler_(EVENT_HANDLERS_.addRowBelow, args)}
function deleteRow(args)                {return eventHandler_(EVENT_HANDLERS_.deleteRow, args)}
function calculateDeadlines(args)       {return eventHandler_(EVENT_HANDLERS_.calculateDeadlines, args)}
function newEventPopup(args)            {return eventHandler_(EVENT_HANDLERS_.newEventPopup, args)}
function prepareNewYearsData(args)      {return eventHandler_(EVENT_HANDLERS_.prepareNewYearsData, args)}
function repopulateBulletins(args)      {return eventHandler_(EVENT_HANDLERS_.repopulateBulletins, args)}
function formatSheet(args)              {return eventHandler_(EVENT_HANDLERS_.formatSheet, args)}
function hideRows(args)                 {return eventHandler_(EVENT_HANDLERS_.hideRows, args)}
function removeEmptyRows(args)          {return eventHandler_(EVENT_HANDLERS_.removeEmptyRows, args)}
function colorBorders(args)             {return eventHandler_(EVENT_HANDLERS_.colorBorders, args)}
function setWeeksFormat(args)           {return eventHandler_(EVENT_HANDLERS_.setWeeksFormat, args)}
function setEventsFormat(args)          {return eventHandler_(EVENT_HANDLERS_.setEventsFormat, args)}
function restoreHeader(args)            {return eventHandler_(EVENT_HANDLERS_.restoreHeader, args)}
function backupHeader(args)             {return eventHandler_(EVENT_HANDLERS_.backupHeader, args)}
function setupAutomation(args)          {return eventHandler_(EVENT_HANDLERS_.setupAutomation, args)}
function disableAutomation(args)        {return eventHandler_(EVENT_HANDLERS_.disableAutomation, args)}
function checkDeadlines(args)           {return eventHandler_(EVENT_HANDLERS_.checkDeadlines, args)}
function checkTeamSheetsForErrors(args) {return eventHandler_(EVENT_HANDLERS_.checkTeamSheetsForErrors, args)}
function dailyTrigger(args)             {return eventHandler_(EVENT_HANDLERS_.dailyTrigger, args)}
function fillSponsor(args)              {return eventHandler_(EVENT_HANDLERS_.fillSponsor, args)}
function fillTier(args)                 {return eventHandler_(EVENT_HANDLERS_.fillTier, args)}

// Private Functions
// =================

// General
// -------

/**
 * All external function calls should call this to ensure standard 
 * processing - logging, errors, etc - is always done.
 *
 * @param {Array} config:
 *   [0] {Function} prefunction
 *   [1] {String} eventName
 *   [2] {String} onErrorMessage
 *   [3] {Function} mainFunction
 *
 * @param {Object}   arg1       The argument passed to the top-level event handler
 */

function eventHandler_(config, args) {

  try {

    var userEmail = Session.getActiveUser().getEmail()

    Log_ = BBLog.getLog({
      level:                DEBUG_LOG_LEVEL_, 
      displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
    })
    
    Log_.info('Handling ' + config[0] + ' from ' + (userEmail || 'unknown email') + ' (' + SCRIPT_NAME + ' ' + SCRIPT_VERSION + ')')
    
    // Call the main function
    return config[2](args)
    
  } catch (error) {
  
    var handleError = Assert.HandleError.DISPLAY_FULL

    if (!PRODUCTION_VERSION_) {
      handleError = Assert.HandleError.THROW
    }

    var assertConfig = {
      error:          error,
      userMessage:    config[1],
      log:            Log_,
      handleError:    handleError, 
      sendErrorEmail: SEND_ERROR_EMAIL_, 
      emailAddress:   ADMIN_EMAIL_ADDRESS_,
      scriptName:     SCRIPT_NAME,
      scriptVersion:  SCRIPT_VERSION, 
    }

    Assert.handleError(assertConfig) 
  }
  
} // eventHandler_()

// Private event handlers
// ----------------------

/**
 *
 *
 * @param {object} 
 *
 * @return {object}
 */
 
function onInstall_() {

  Log_.functionEntryPoint()
  
  // TODO ...

} // onInstall() 