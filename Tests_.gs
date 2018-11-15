function test_misc() {
  var a = SpreadsheetApp.openById('15mrKpNDi7RZ8ddDPKGdJICOA2a1jJhmHPFZzO-4kL5Y').getRange('G175').getValue()
  debugger
}

function test_checkDeadlines() {
  checkDeadlines_()
}

function test_getStaff() {
  var s = getStaff_()
}


function test_processResponse() {
  processResponse_(['Test2053', 'Bronze', '12/10/2018', 'Worship Team'])
}

function test_addRowBelow() {
  addRowBelow_()
}

// BBLog
// -----

function bblogFine_(what, optSubject){ 

  if (PRODUCTION_VERSION_) {
    return;
  }

  BBLog
    .getLog({
      level: BBLog.Level.FINE,
      displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
    })
    .fine(what + ' - ' + (optSubject || ''))
  
}

function bblogInfo_(what, optSubject) { 

  BBLog
    .getLog({
      displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
    })
    .info(what + ' - ' + (optSubject || ''))
}

function bblogError_(what, optSubject){ 

  BBLog
    .getLog({
      level: BBLog.Level.SEVERE,
      displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
    })
    .severe(what + ' - ' + (optSubject || ''))

} // bblogError_()
