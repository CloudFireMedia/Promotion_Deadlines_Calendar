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

//                           Name                            onError Message                          Main Functionality
//                           ----                            ---------------                          ------------------

  addRowBelow:               ['addRowBelow()',               'Failed to add row below',               addRowBelow_],
  deleteRow:                 ['deleteRow()',                 'Failed to delete row',                  deleteRow_],
  calculateDeadlines:        ['calculateDeadlines()',        'Failed to calculate deadlines',         calculateDeadlines_],
  fillSponsor:               ['fillSponsor()',               'Failed to fill sponsor',                fillSponsor_],
  fillTier:                  ['fillTier()',                  'Failed to fill tier',                   fillTier_],
  newEventPopup:             ['newEventPopup()',             'Failed to show event popup',            newEventPopup_],
  prepareNewYearsData:       ['prepareNewYearsData()',       'Failed to prepare New Years data',      prepareNewYearsData_],
  repopulateBulletins_:      ['repopulateBulletins_()',      'Failed to repopulateBulletins_',        repopulateBulletins_],
/*  
  :             ['()',             'Failed to ',            _],
  :             ['()',             'Failed to ',            _],
  :             ['()',             'Failed to ',            _],
  :             ['()',             'Failed to ',            _],
  :             ['()',             'Failed to ',            _],
  :             ['()',             'Failed to ',            _],
  :             ['()',             'Failed to ',            _],
  :             ['()',             'Failed to ',            _],
  :             ['()',             'Failed to ',            _],
*/
}

function addRowBelow(args)          {return eventHandler_(EVENT_HANDLERS_.addRowBelow, args)}
function deleteRow(args)            {return eventHandler_(EVENT_HANDLERS_.deleteRow, args)}
function calculateDeadlines(args)   {return eventHandler_(EVENT_HANDLERS_.calculateDeadlines, args)}
function fillSponsor(args)          {return eventHandler_(EVENT_HANDLERS_.fillSponsor, args)}
function fillTier(args)             {return eventHandler_(EVENT_HANDLERS_.fillTier, args)}
function newEventPopup(args)        {return eventHandler_(EVENT_HANDLERS_.newEventPopup, args)}
function prepareNewYearsData(args)  {return eventHandler_(EVENT_HANDLERS_.prepareNewYearsData, args)}
function repopulateBulletins_(args) {return eventHandler_(EVENT_HANDLERS_.repopulateBulletins_, args)}
/*
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
function (args)      {return eventHandler_(EVENT_HANDLERS_., args)}
*/

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
    return config[2](arg1)
    
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