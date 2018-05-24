var OPTIONS = {};

function initWeeklyOptions() {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

  getOptionsData();

  var sheetName = formatDate(OPTIONS.startDate) + ' : ' + formatDate(OPTIONS.finalDate);
  var existingSheet = _ss.getSheetByName(sheetName);
  if (existingSheet) _ss.deleteSheet(existingSheet);
  createNewSheet(sheetName, '#93c47d');
}
