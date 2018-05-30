function initPeriodTable(color) {
  writePeriodHeader(color);
  writePeriodUserRows(color);
}

function writePeriodHeader(color) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  sheet.getRange(1, 1).setValue('Менеджер').setBackground(color);
  var columnI = 2;
  REPORT.forEach(function(k) {
    sheet.getRange(1, columnI++).setValue(k.name).setBackground(color);
  });
}

function writePeriodUserRows(color) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rowI = 2;
  performers = OPTIONS.performers;
  performers.forEach(function(user, i) {
    var userData = APIRequest('users', {query: [{key: 'name', value: user}]}).users[0];
    sheet.getRange(rowI++, 1).setValue(userData.firstname + ' ' + userData.lastname + ' (' + userData.login + ')').setBackground(color);
  });
}
