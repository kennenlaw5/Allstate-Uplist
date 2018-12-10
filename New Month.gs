function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
}

function newMonth() {
  var date = new Date('January 1, 2019');
  if (date.getDate() != 1) { return; }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = Utilities.formatDate(date, 'MST', "MMM YYYY").toUpperCase();
  var dealers = ['Mini', 'Honda', 'BMW'];
  if (date.getMonth() == 8) { name = name.replace('SEP', 'SEPT'); }
  
  for (var i = 0; i < dealers.length; i++) {
    ss.getSheetByName(dealers[i] + ' Master').copyTo(ss).setName(dealers[i] + ' ' + name).activate();
    ss.moveActiveSheet(2);
    SpreadsheetApp.flush()
  }
}
