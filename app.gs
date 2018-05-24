function createDailyReport() {
  initOptions();
}

function createWeeklyReport() {
  initWeeklyOptions();
  initPeriodTable('#d9ead3');
  processPeriodReports();
}

function createMonthlyReport() {
  initMonthlyOptions();
  initPeriodTable('#fff2cc');
  processPeriodReports();
}

function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.addMenu('GoldenCode Report', [
    {name: 'Создать Ежедневный Отчёт', functionName: 'createDailyReport'},
    {name: 'Создать Еженедельный Отчёт', functionName: 'createWeeklyReport'},
    {name: 'Создать Ежемесячный Отчёт', functionName: 'createMonthlyReport'}
  ]);
}
