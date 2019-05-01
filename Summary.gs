function onEdit(e) {
  var sheet = e.source;
  if (sheet.getSheetName().toLowerCase() != 'summary') { return; }
  var range = e.range.getA1Notation();
  
  if (range != 'P2' && range != 'Q2') { return; }
  
  var month, year, honda, bmw, mini, current, dateCheck, formulas;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  
  if (range == 'P2') {
    month = e.value;
    year = sheet.getRange('Q2').getDisplayValue();
  }
  else {
    year = e.value;
    month = sheet.getRange('P2').getValue();
  }
  
  month = month.toLowerCase();
  year = year.toString();
  
  for (var i = 0; i < sheets.length; i++) {
    current = sheets[i].getSheetName().toLowerCase().split(' ');
    dateCheck = current.indexOf(month) != -1 && current.indexOf(year) != -1;
    
    if (dateCheck) { 
      if (current.indexOf('bmw') != -1) { bmw = sheets[i].getSheetName(); }
      else if (current.indexOf('honda') != -1) { honda = sheets[i].getSheetName(); }
      else if (current.indexOf('mini') != -1) { mini = sheets[i].getSheetName(); }
    }
  }
  
  if ([bmw, honda, mini].indexOf(undefined) !== -1) { return; }
  
  range = sheet.getRange('A3:K3');
  formulas = range.getFormulas();
  formulas[0][0] = "=SORT(agentRatio('" + bmw + "'!C:E),1,TRUE)";
  formulas[0][5] = "=SORT(agentRatio('" + honda + "'!C:E),1,TRUE)";
  formulas[0][10] = "=SORT(agentRatio('" + mini + "'!C:E),1,TRUE)";
  range.setFormulas(formulas);
}
