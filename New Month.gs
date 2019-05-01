function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
}

function newMonth() {
  var date = new Date();
  if (date.getDate() != 1) { return; }
  var name = Utilities.formatDate(date, 'MST', "MMM YYYY").toUpperCase();
  var dealers = ['Mini', 'Honda', 'BMW'];
  if (date.getMonth() == 8) { name = name.replace('SEP', 'SEPT'); }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var needsDuplicate;
  
  
  for (var i = 0; i < dealers.length; i++) {
    needsDuplicate = false;
    if (ss.getSheetByName(dealers[i] + ' ' + name) !== null) { needsDuplicate = true; }
    ss.getSheetByName(dealers[i] + ' Master').copyTo(ss).setName(dealers[i] + (needsDuplicate ? ' 2' : '') + ' ' + name).activate();
    SpreadsheetApp.flush();
    ss.moveActiveSheet(2);
    SpreadsheetApp.flush();
  }
  
  updatePaceLinkSheet();
  
  if (date.getMonth() == 0) { updateSummaryValidation(date.getYear()); }
}

function updatePaceLinkSheet() {
  SpreadsheetApp.flush();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('PACE_link');
  
  if (sheet === null) { throw 'PACE_link sheet was not found! Please correct the error or contact Kennen!'; }
  
  var month, year, range, formulas;
  var sheets = ss.getSheets();
  var name = Utilities.formatDate(new Date(), 'MST', "MMM YYYY").toUpperCase();
  if (name.indexOf('SEP') !== -1) { name = name.replace('SEP', 'SEPT'); }
  
  month = name.split(' ')[0];
  year = name.split(' ')[1].toString();
  sheet.getRange('P2:Q2').setValues([[month, year]]);
  
  range = sheet.getRange('A3:K3');
  formulas = range.getFormulas();
  formulas[0][0] = "=SORT(agentRatio('BMW " + name + "'!$C$1:E, P3),1,TRUE)";
  formulas[0][5] = "=SORT(agentRatio('Honda " + name + "'!$C$1:E, P3),1,TRUE)";
  formulas[0][10] = "=SORT(agentRatio('Mini " + name + "'!$C$1:E, P3),1,TRUE)";
  range.setFormulas(formulas);
}

function updateSummaryValidation(year) {
  year = year.toString();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Summary');
  var range = sheet.getRange('Q2');
  var validation = range.getDataValidation();
  var values = validation.getCriteriaValues();
  values[0].push(year);
  validation = validation.copy().requireValueInList(values[0]).setAllowInvalid(false).build();
  range.setDataValidation(validation);
}