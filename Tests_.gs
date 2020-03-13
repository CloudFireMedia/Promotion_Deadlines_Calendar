function test_init() {
    Log_ = BBLog.getLog({
      level:                DEBUG_LOG_LEVEL_, 
      displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
      sheetId:              TEST_PDC_SPREADSHEET_ID_
    })
}

function test_misc() {
  var a = PropertiesService.getDocumentProperties().deleteAllProperties()
  return a
}

function test_checkDeadlines() {
  checkDeadlines_()
}

function test_getStaff() {
  var s = getStaff_()
}

function test_processResponse() {
  test_init()
  processResponse_(['Test2053', 'Bronze', '12/10/2018', 'Worship Team'])
}

function test_addRowBelow() {
  test_init()
  var ss = SpreadsheetApp.openById(TEST_PDC_SPREADSHEET_ID_)
  addRowBelow_(ss, 4)
}
