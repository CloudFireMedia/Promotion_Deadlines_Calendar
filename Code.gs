function prepareNewYearsData_() { 

  //mostly is NOT formatting but whatev, the menu action name should be updated
  var sheet = SpreadsheetApp.getActive().getSheetByName(config.eventsCalendar.dataSheetName);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  var startRow = sheet.getFrozenRows();
  var numRows = values.length;
  var numColumns = values[0].length;

  //populate LISTED ON WEB CAL and PROMO REQUESTED columns
  for(var i=startRow; i<numRows; i++) {
    if(values[i][5] || values[i][6]) continue;//something already set, next please!
    var val = values[i][2] == "N/A" ? 'N/A' : 'No';
    sheet.getRange(i + 1, 6).setValue(val);//LISTED ON WEB CAL
    sheet.getRange(i + 1, 7).setValue(val);//PROMO REQUESTED
  }
  
  combineSundayWeekCells();
  
  // Private Functions
  // -----------------
  
  function ccombineSundayWeekCells() {
    //search EVENT TITLE column for 'Sunday Service'
    var sheet = SpreadsheetApp.getActive().getSheetByName(config.eventsCalendar.dataSheetName);
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var sundayRows = [];
  
    for(var i=4; i<values.length; i++)//skip three header rows
      if(values[i][4].indexOf("Sunday Service") > -1)
        sundayRows.push(i);
    
    var maxValue = sundayRows[sundayRows.length - 1];//the last entry
  
    for(var k = 0; k < sundayRows.length; k++) {
      var from = sundayRows[k]+1;//+1 0-based array offset
      var to   = sundayRows[k+1];//+1 next array index, gives us the row before the next Sunday
      var numRows = to - from +1;//number of rows
      if(from > maxValue) break;
      //sheet.getRange(from, 1, 1, 2).mergeVertically();//uh, merging a single row doesn't do anything, just skip (or end in this case)
      //else
      sheet.getRange(from, 1, numRows, 2).mergeVertically();
    }
  }
  
} // prepareNewYearsData()

function applyFormula_(){ 
  var sheet = SpreadsheetApp.getActive().getSheetByName(config.eventsCalendar.dataSheetName);
  var values = sheet.getDataRange().getValues();
  //fix the WEEK column formula
  sheet.getRange('A3').setFormula('={"WEEK"; ArrayFormula( if(D4:D, WEEKNUM(D4:D), IFERROR(1/0)) ) }')
  //fix the Bronze, Silver, Gold column formulae
  sheet.getRange('I3:K3').setFormulas([[
    '={"Bronze";ArrayFormula(if(LEN(D4:D),if(GTE(D4:D, TODAY()+21), D4:D-21, "--"),IFERROR(1/0)))}',
    '={"Silver";ArrayFormula(if(LEN(D4:D),if(GTE(D4:D, TODAY()+42), D4:D-42, "--"),IFERROR(1/0)))}',
    '={"Gold";  ArrayFormula(if(LEN(D4:D),if(GTE(D4:D, TODAY()+70), D4:D-70, "--"),IFERROR(1/0)))}'
  ]]);
}

function getStaff_() {//returns [{},{},...]

  var sheet = SpreadsheetApp.openById(config.files.staffData).getActiveSheet();
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

function repopulateBulletins_() { 

  ////this collects bulletin names for makeRepeat_() which is probably unused
  // so there's really no reason for it to run either
  // Since I'm keeping the other script just-in-case, I'll leave this one too
  // --Bob 20181208
  //note: based on the menu action name, perhaps this is used during new yeat processing.
  
  var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME_);
  var range = sheet.getRange(sheet.getFrozenRows()+1, 2, sheet.getLastRow()-sheet.getFrozenRows());
  var bulletins = range.getValues()
  .reduce(function(bulletins,name){
    if(name[0]) bulletins.push(name[0]);
    return bulletins;
  },[]);
  
  makeRepeat(bulletins, sheet.getFrozenRows()+1);//readFrom...uh, was 6.  But I don't know why it would be so --Bob
  
  // Private Functions
  // -----------------
  
  function makeRepeat(arr, readFrom) {
    ////this appears to be doing nothing
    // at one time it made the bulletin names repeat but that was likely before the bulletin names were merged vertically
    // the script runs but makes no changes as there should never be a matcihng condition.
    // I'll leave it here in case there is a use for it that I'm unaware of
    // --Bob 20181208
    var sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET_NAME_);
    var range = sheet.getDataRange();
    var values = range.getValues();
    var k = 0;
    var startFrom = 0;
  
    //gets the first row with data inb col 1
    for(var j = readFrom; j <= values.length; j++) {
      if(range.getCell(j, 1).getValue()) {
        var startFrom = j;
        break;
      }
    }
    
    while(k < 52) {//yeah, only if there really /are/ 52... otherwise infinite loop
      if(startFrom > values.length) return;//happens during recursion
      if(range.getCell(startFrom, 1).getValue()) {
        range.getCell(startFrom, 2).setValue(arr[k]);
        k++
      }
      startFrom++;
    }
    makeRepeat(arr, startFrom)
  }
  
} // repopulateBulletins_()