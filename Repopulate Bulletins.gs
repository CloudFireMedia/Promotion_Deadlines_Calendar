//function repopulateBulletins() {
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var s = ss.getActiveSheet();
//  var dataRange = s.getDataRange();
//  var values = dataRange.getValues();
//  var arr=new Array();
//  var p=0;
//  //Here set value of first 1 row id
//  var readFrom=6;
//  while(p<52)
//  {    
//    if(dataRange.getCell(readFrom,1).getValue()!="")
//    {
//      arr.push(dataRange.getCell(readFrom,2).getValue());
//      p++
//    }
//    readFrom++;
//  }
//  makeRepeat(arr,readFrom);
//}
//
//
//
//function makeRepeat(arr,readFrom)
//{
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var s = ss.getActiveSheet();
//  var dataRange = s.getDataRange();
//  var values = dataRange.getValues();
//  var k=0;
//  var startFrom="";
//  for(var j=readFrom; j<=values.length; j++)
//  {
//    if(dataRange.getCell(j,1).getValue()!="")
//    {
//      var startFrom=j;
//      break;
//    }
//  }
//  while(k<52)
//  {    
//    if(startFrom>values.length)
//    {
//      return;
//    }
//    if(dataRange.getCell(startFrom,1).getValue()!="")
//    {
//      dataRange.getCell(startFrom,2).setValue(arr[k]);
//      k++
//    }   
//    startFrom++;
//  }
//  makeRepeat(arr,startFrom)
//} 