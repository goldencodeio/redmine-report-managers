var OPTIONS = {};

function initMonthlyOptions() {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

  getOptionsData();

  var sheetName = 'Итог Месяц';
  var existingSheet = _ss.getSheetByName(sheetName);
  if (existingSheet) _ss.deleteSheet(existingSheet);
  createNewSheet(sheetName, '#ffd966');
}
