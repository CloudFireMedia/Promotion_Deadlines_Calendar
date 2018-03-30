var SCRIPT_NAME = ' CCN Events Promotion Calendar';
var SCRIPT_VERSION = 'v1.0'

//settings
var documentConfig = {
  dataSheetName : 'Communications Director Master',
};

//push config options to library
PL.updateConfig({eventsCalendar:documentConfig});

function onEdit(e)                 {}
function runMeToGrantPermissions() {}
function onOpen(e)                 {PL.onOpen()}
//function onEdit_Triggered(e)     {PL.onEdit_eventsCalendar_Triggered(e)}//not needed
function setupAutomation()         {PL.setupAutomation_eventsCalendar()}
function disableAutomation()       {PL.disableAutomation_eventsCalendar()}
function calendar_dailyTrigger()   {PL.calendar_dailyTrigger()}


//debug
function dumpConfig() {PL.dumpConfig()}
function foo() {PL.foo()}


// Original bound code, now in ChCS Github Repo - https://github.com/Christ-Church-Nashville/ChCS_Events_Promotion_Calendar v1.0