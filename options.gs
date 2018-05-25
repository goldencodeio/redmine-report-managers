var OPTIONS = {};

function initOptions() {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

  getOptionsData();

  OPTIONS.currentDate = OPTIONS.startDate;
  Logger.log(OPTIONS.currentDate.toJSON());
  var sheetName = formatDate(OPTIONS.currentDate);
  var existingSheet = _ss.getSheetByName(sheetName);
  var tDate = new Date(OPTIONS.currentDate.getTime());
  if (existingSheet) {
    existingSheet.activate();
    initTable();
    processReports();
  } else {
    tDate.setDate(tDate.getDate() -1);
    var prevSheet = _ss.getSheetByName(formatDate(tDate));
    var isRangesEmpty = false;
    if (prevSheet) {
      prevSheet.activate();
      var namedRanges = prevSheet.getNamedRanges();
      for (var i = 0; i < namedRanges.length; i++) {
        if (namedRanges[i].getRange().getValue() === '') isRangesEmpty = true;
      }
    }
    if (!isRangesEmpty) {
      createNewSheet(sheetName, '#6d9eeb');
      initTable();
      processReports();
    } else Browser.msgBox('Не заполнены ручные поля за предыдущий день');
  }
}

function getOptionsData() {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();

  var optionsSheet = _ss.setActiveSheet(getOptionsSheet());

  var data = optionsSheet.getRange(1, 1, optionsSheet.getLastRow(), optionsSheet.getLastColumn()).getValues();
  data.forEach(function(row) {
    var key = row.shift();
    row = row.filter(function(a) {
      if (a === 0) return true;
      return a
    });
    OPTIONS[key] = row.length > 1 ? row : row[0];
  });

  OPTIONS.startDate.setHours(OPTIONS.startDate.getHours() - 1 * OPTIONS.startDate.getTimezoneOffset() / 60);
  OPTIONS.finalDate.setHours(OPTIONS.finalDate.getHours() - 1 * OPTIONS.finalDate.getTimezoneOffset() / 60);

  if (!Array.isArray(OPTIONS.performers)) OPTIONS.performers = [OPTIONS.performers];
  if (!Array.isArray(OPTIONS.performersWorkHours)) OPTIONS.performersWorkHours = [OPTIONS.performersWorkHours];
}

function getOptionsSheet() {
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().toLowerCase() === 'options')
      return sheets[i];
  }
  return null;
}

function createNewSheet(name, color) {
  var _ss = SpreadsheetApp.getActiveSpreadsheet();
  return _ss.insertSheet(name).setTabColor(color).setColumnWidth(1, 200).activate();
}
