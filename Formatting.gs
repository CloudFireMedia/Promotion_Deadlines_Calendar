/* 
 * NOTE - 2020.01.16
 *
 * The formatCommunicationsDirectorMaster function uses an installable daily trigger to:
 *
 *   a) break apart merged ranges
 *   b) sort sheet chronologically
 *   c) delete empty rows
 *   d) vertically merge identical and contiguous weeks, months, and bulletins across 1-2 calendar years 
 *   e) (re)set conditional and non-conditional formatting
 *   f) (re)set date formatting
 *   g) hide rows before today 
 *
 * Regarding pt d): vertically merging months, weeks, etc. in a grid that contains >1 calendar year requires either 1) API calls in multiple
 * loops, which is programmatically expensive, or B) finding the bitonic point in a bitonic series (e.g. the point of inflection in an 
 * increasing then decreasing series). In the code below, I opt for option B and look for the bitonic point in the PDC's week numbers array.
 *
 * The PDC's week numbers array does not represent a 'strict' bitonic series in which the difference between adjacent values always == 1. 
 * Therefore it is necessary to define the longest increasing subsequence (LIS) before looking for the bitonic point. In a non-strict bitonic 
 * series, the value of the last indexed element of the LIS is the first integer in the bitonic point. One way to find the 
 * the last element of the LIS is to find the 'biggest chronological drop' in the series. See: https://stackoverflow.com/questions/53481266/get-the-biggest-chronological-drop-min-and-max-from-an-array-with-on
 *
 * References:
 *
 * - Finding the bitonic point in bitonic sequence:
 *   - https://www.geeksforgeeks.org/find-the-maximum-element-in-an-array-which-is-first-increasing-and-then-decreasing/
 *   - https://www.geeksforgeeks.org/find-bitonic-point-given-bitonic-sequence/ 
 *   - https://stackoverflow.com/questions/11536123/finding-an-number-in-montonically-increasing-and-then-decreasing-sequencecera/11536172#11536172
 *   - https://gatestack.in/t/code-to-find-inflation-point-in-bitonic-series-javascript/1157
 *   - https://codereview.stackexchange.com/questions/180455/function-to-check-if-a-numbers-sequence-is-increasing/180486#180486
 *   - https://en.wikipedia.org/wiki/Bitonic_sorter
 *
 * - Solving longest increasing subsequence (LIS) problem: 
 *   - https://gist.github.com/wheresrhys/4497653
 *   - https://stackoverflow.com/questions/45314870/javascript-find-out-of-sequence-dates
 *   - https://codereview.stackexchange.com/questions/138455/hand-crafted-longest-common-sub-sequence
 *   - https://codereview.stackexchange.com/questions/29674/finding-the-longest-non-decreasing-subsequence-in-a-grid
 *
 * P.S. - I briefly entertained the idea that I could avoid dealing with both the bitonic point and the LIS problems by constructing an array of 
 * dates from Col D, filtering out all date properties except for the year property, and then simply finding the lastIndexOf the current year. 
 * Unfortunately, it turns out that dates are notoriously messy in javascript. Serial manipulation of a dates array such as filtering date properties 
 * appears to require regex that is above my level. See: https://stackoverflow.com/questions/47355050/filter-array-of-dates
 *
 * --cdb
 *
 */

function formatCommunicationsDirectorMaster_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Communications Director Master");  
  
  // Show toast popup
  SpreadsheetApp.getActiveSpreadsheet().toast('Reformatting sheet.', 'Please wait...', -1)

  var TwoDArray = sheet.getRange("B4:B").getValues(); // get 2D array
  
  // Flatten to 1D array
  var OneDArray = [] ;
  for (var i = 0; i < TwoDArray.length; i++) {
    OneDArray.push(TwoDArray[i][0]);
  }
  var firstYearWeeksArray = OneDArray;
  var lastWeekInFirstYear = (checkData(firstYearWeeksArray)[0]); //should be 29.0
  Log_.fine("lastWeekInFirstYear" + lastWeekInFirstYear); 
  Log_.fine("answer!!!!! " + firstYearWeeksArray.lastIndexOf(lastWeekInFirstYear));
  var lastRowofFirstYear = firstYearWeeksArray.lastIndexOf(lastWeekInFirstYear)+4;
  
/* 
 * NOTE: getLastRow() method counts trailing blank rows that are subject to an array formula, even when the array places
 * no values in those rows. Therefore, method sheet.deleteRows(lastRow+1, maxRows-lastRow) will not work without a helper
 * function to pass the correct, 'pseudo-blank' row number to var lastRow.
 */
  
  //delete empty rows at the bottom of the sheet
  var maxRows = sheet.getMaxRows(); 
  var lastRow = (function getLastPopulatedRow() {
    var data = sheet.getRange("E:E").getValues(); //check event title only
    for (var i = data.length-1; i > 0; i--) {
      for (var j = 0; j < data[0].length; j++) {
        if (data[i][j]) return i+1;}}
    return 0;})(); 
  if(maxRows > lastRow){
    sheet.deleteRows(lastRow+1, maxRows-lastRow);}
  
/* 
 * NOTE: GAS breakapart() is buggy. If we add a non-merged 'helper' row on the bottom of the sheet before
 * running breakapart(), this allows us to avoid the otherwise persistent and cryptic error: "You must select 
 * all cells in a merged range to merge or unmerge them." Further, if we delete the helper AFTER we have 
 * merged it with the rest of the range, this avoids leaving a remainder of unmerged rows at the bottom of the sheet.
 * ‾\_(ツ)_/‾
 */
  
  // Append 1 helper row to the end of year 1
  sheet.insertRowAfter(lastRowofFirstYear); 
  
  // Append 1 helper row to the end of year 2
  var rowLast = sheet.getLastRow();
  sheet.insertRowAfter(rowLast); 
  
  // UnMerge Cols A and B
  sheet.getRange(4, 1, sheet.getLastRow()-4, 2).activate().breakApart();
  
  // unMerge Col M
  sheet.getRange(4, 13, sheet.getLastRow()-4, 1).activate().breakApart();
  
  // Sort sheet by Col 4
  sheet.sort(4);
  
  // Clear conditional formatting only
  sheet.clearConditionalFormatRules();
  
/*  
 * Timestamps are haphazardly added to Col 9 by a separate 'Calculate Deadlines' script. Timestamps interfere with 
 * Google Sheet's conditional formatting rules and need to be stripped out before the present script can continue. 
 * The following code is hacky but after trying a bunch of other methods, this is the quickest / simplest one that I
 * could find.
 */
  
  // Strip out timestamps, preserve formatting
  sheet.getRange(4,9,sheet.getLastRow()-4,3).setNumberFormat("MM/dd/yyyy").setNumberFormat('@STRING@').setNumberFormat("MM/dd/yyyy").setNumberFormat("MM.DD");
  
  // Reformat spreadsheet
  var numColumns = sheet.getMaxColumns();  
  var numRows = sheet.getMaxRows();  
  
  ss.getRange('Format!A1:M3').copyTo(sheet.getRange('Communications Director Master!A1:M3'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); 
  ss.getRange('Format!A4:M4').copyTo(sheet.getRange('Communications Director Master!A4:M4'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); 
  ss.getRange('Format!A5:M5').copyTo(sheet.getRange('Communications Director Master!A5:M'), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false); 
  
  // Merge Col A for year 1
  var startAY1 = 4;
  var stopAY1 = lastRowofFirstYear;
  var cAY1 = {};
  var kAY1 = "";
  var offsetAY1 = 0;
  var dataAY1 = sheet.getRange("A" + startAY1 + ":" + "A" + stopAY1).getValues().filter(String);
  
  // Retrieve the number of duplication values
  dataAY1.forEach(function(eAY1){cAY1[eAY1[0]] = cAY1[eAY1[0]] ? cAY1[eAY1[0]] + 1 : 1;});
  
  // Merge cells
  dataAY1.forEach(function(eAY1){
    if (kAY1 != eAY1[0]) {
      sheet.getRange(startAY1 + offsetAY1, 1, cAY1[eAY1[0]], 1).merge();
      offsetAY1 += cAY1[eAY1[0]];
    }
    kAY1 = eAY1[0];
  });
  
  // Merge Col A for year 2
  var startAY2 = lastRowofFirstYear+1;
  var stopAY2 = sheet.getLastRow();
  var cAY2 = {};
  var kAY2 = "";
  var offsetAY2 = 0;
  var dataAY2 = sheet.getRange("A" + startAY2 + ":" + "A" + stopAY2).getValues().filter(String);
  
  // Retrieve the number of duplication values
  dataAY2.forEach(function(eAY2){cAY2[eAY2[0]] = cAY2[eAY2[0]] ? cAY2[eAY2[0]] + 1 : 1;});
  
  // Merge cells
  dataAY2.forEach(function(eAY2){
    if (kAY2 != eAY2[0]) {
      sheet.getRange(startAY2 + offsetAY2, 1, cAY2[eAY2[0]], 1).merge();
      offsetAY2 += cAY2[eAY2[0]];
    }
    kAY2 = eAY2[0];
  });
  
  // Merge Col B and Col M for year 1
  var startBY1 = 4;
  var stopBY1 = lastRowofFirstYear;
  var cBY1 = {};
  var kBY1 = "";
  var offsetBY1 = 0;
  var dataBY1 = sheet.getRange("B" + startBY1 + ":" + "B" + stopBY1).getValues().filter(String);
  
  // Retrieve the number of duplication values
  dataBY1.forEach(function(eBY1){cBY1[eBY1[0]] = cBY1[eBY1[0]] ? cBY1[eBY1[0]] + 1 : 1;});
  
  // Merge cells in Col B
  dataBY1.forEach(function(eBY1){
    if (kBY1 != eBY1[0]) {
      sheet.getRange(startBY1 + offsetBY1, 2, cBY1[eBY1[0]], 1).merge();
      // Impose Col B ('Weeks') merge pattern on Col M ('Bulletins')
      sheet.getRange(startBY1 + offsetBY1, 13, cBY1[eBY1[0]], 1).merge();
      offsetBY1 += cBY1[eBY1[0]];
    }
    kBY1 = eBY1[0];
  });
  
  // Merge Col B and Col M for year 2
  var startBY2 = lastRowofFirstYear+1;
  var stopBY2 = sheet.getLastRow();
  var cBY2 = {};
  var kBY2 = "";
  var offsetBY2 = 0;
  var dataBY2 = sheet.getRange("B" + startBY2 + ":" + "B" + stopBY2).getValues().filter(String);
  
  // Retrieve the number of duplication values
  dataBY2.forEach(function(eBY2){cBY2[eBY2[0]] = cBY2[eBY2[0]] ? cBY2[eBY2[0]] + 1 : 1;});
  
  // Merge cells in Col B
  dataBY2.forEach(function(eBY2){
    if (kBY2 != eBY2[0]) {
      sheet.getRange(startBY2 + offsetBY2, 2, cBY2[eBY2[0]], 1).merge();
      // Impose Col B ('Weeks') merge pattern on Col M ('Bulletins')
      sheet.getRange(startBY2 + offsetBY2, 13, cBY2[eBY2[0]], 1).merge();
      offsetBY2 += cBY2[eBY2[0]];
    }
    kBY2 = eBY2[0];
  });
  
  // Delete helper row plus any empty rows that were moved to the bottom of the sheet during sorting
  var maxRows = sheet.getMaxRows(); 
  var lastRow = (function getLastPopulatedRow() {
    var data = sheet.getRange("E:E").getValues(); //check Event Title column only
    for (var i = data.length-1; i > 0; i--) {
      for (var j = 0; j < data[0].length; j++) {
        if (data[i][j]) return i+1;}}
    return 0;})(); 
  if(maxRows > lastRow){
    sheet.deleteRows(lastRow+1, maxRows-lastRow);}
  
  // Activate no cells
  sheet.activate();
  
  // Hide rows before today
  var v = sheet.getRange("D:D").getValues();
  var today = new Date();
  today.setHours(0, 0, 0, 0);
  for (var i = sheet.getLastRow(); i > 2; i--) {
    var t = v[i - 1];
    if (t != "") {
      var u = new Date(t);
      if (u < today) {
        sheet.hideRows(i);
      }
    }
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Reformat complete.', 'Thanks!', 4);
  return;
  
  // Private Functions
  // -----------------

  function checkData(data) {
    var bestDropStart = 0
    var bestDropEnd = 0
    var bestDrop = 0
    var currentDropStart = 0
    var currentDropEnd = 0
    var currentDrop = 0
    for (var i = 1; i < data.length; i++) {
      if (data[i] < data[i - 1]) {
        // we are dropping
        currentDropEnd = i
        currentDrop = data[currentDropStart] - data[i]
      } else {
        // the current drop ended; check if it's better
        if (currentDrop > bestDrop) {
          bestDrop = currentDrop
          bestDropStart = currentDropStart
          bestDropEnd = currentDropEnd
        }
        // start a new drop
        currentDropStart = currentDropEnd = i
        currentDrop = 0
      }
    }
    // check for a best drop at end of data
    if (currentDrop > bestDrop) {
      bestDrop = currentDrop
      bestDropStart = currentDropStart
      bestDropEnd = currentDropEnd
    }
    // return the best drop data
    return [data[bestDropStart], data[bestDropEnd], bestDrop]
    
  } // formatCommunicationsDirectorMaster_.checkData()

} // formatCommunicationsDirectorMaster_()
