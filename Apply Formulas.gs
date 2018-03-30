//function ApplyFormula()
//{
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var s = ss.getActiveSheet();
//  var dataRange = s.getDataRange();
//  var values = dataRange.getValues();
//  
//  
//  for(var k=4;k<=values.length;k++)
//  {
//    var formulaI="=D"+k+"-21";
//    var formulaJ="=D"+k+"-42";
//    var formulaK="=D"+k+"-70";
//    var formulaL='=if(and(F'+k+'="No",G'+k+'="No"),"Awaiting Promotion Request",if(and(F'+k+'="Yes",G'+k+'="Yes"),"Scheduled",if(and(F'+k+'="Yes",G'+k+'="No"),"Awaiting Promotion Request",if(and(F'+k+'="No",G'+k+'="Yes"),"Pending Review",if(and(F'+k+'="Yes",G'+k+'="N/A"),"ERROR: COLUMNS F & G MUST NOT CONTAIN ONLY ONE N/A",if(and(F'+k+'="N/A",G'+k+'="Yes"),"ERROR: COLUMNS F & G MUST NOT CONTAIN ONLY ONE N/A",if(and(F'+k+'="N/A",G'+k+'="No"),"ERROR: COLUMNS F & G MUST NOT CONTAIN ONLY ONE N/A",if(and(F'+k+'="No",G'+k+'="N/A"),"ERROR: COLUMNS F & G MUST NOT CONTAIN ONLY ONE N/A",if(and(F'+k+'="N/A",G'+k+'="N/A"),"N/A","N/A")))))))))';
//    
//    s.getRange(k,9).setFormula(formulaI); 
//    s.getRange(k,10).setFormula(formulaJ); 
//    s.getRange(k,11).setFormula(formulaK); 
//    s.getRange(k,12).setFormula(formulaL); 
//  }
//}
