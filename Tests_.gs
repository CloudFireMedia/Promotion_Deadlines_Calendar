function tets_misc() {
  var a = SpreadsheetApp.openById('15mrKpNDi7RZ8ddDPKGdJICOA2a1jJhmHPFZzO-4kL5Y').getRange('G175').getValue()
  debugger
}


// BBLog
// -----

function bblogFine_(what, optSubject){ 

  if (!config.debug) {
    return;
  }

  BBLog
    .getLog({level: BBLog.Level.FINE})
    .fine(what + ' - ' + (optSubject || ''))
  
}

function bblogInfo_(what, optSubject) { 

  BBLog
    .getLog()
    .info(what + ' - ' + (optSubject || ''))
}

function bblogError_(what, optSubject){ 

//  var logLevel = (config.debug) ? BBLog.Level.FINE : BBLog.Level.INFO;

  BBLog
    .getLog({level: BBLog.Level.SEVERE})
    .severe(what + ' - ' + (optSubject || ''))

/*
  try { 
    Tools.tellBob(what, config.clientName, config.projectName, optSubject) 
  }
  catch(e){ 
    write('Tools library inaccessible or not installed.');
    MailApp.sendEmail({
      to:'bob+christchurchnashville@rupholdt.com',
      subject:'Error on Christ Church Nashville app',
      body:'Tools library inaccessible or not installed in CCN Promotions Library.'
    })
  }
*/  

} // bblogError_()
