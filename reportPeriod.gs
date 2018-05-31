function processPeriodReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetWeekly = ss.getActiveSheet();
  var allSheets = ss.getSheets();
  var dailySheets = [];
  var performers = [];

  allSheets.forEach(function(sheet, i) {
    var tDate = new Date(OPTIONS.startDate.getTime());
    while (tDate.getTime() <= OPTIONS.finalDate.getTime()) {
      if (sheet.getName() === formatDate(tDate)) dailySheets.push(sheet);
      tDate.setDate(tDate.getDate() + 1);
    }
  });

  if (dailySheets.length > 0) {
    dailySheets.forEach(function(sheet, iSheet) {
      var data = sheet.getRange(2, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
      data.forEach(function(row) {
        var login = row.shift();
        if (typeof(login) !== 'string') return;
        login = login.match(/\d{4}/);
        if (!login) return;
        login = login[0];

        OPTIONS.performers.forEach(function(user, iUser) {
          if (user == login){
            if (performers[iUser] === undefined) performers[iUser] = [];
            performers[iUser].push(row);
          }
        });
      });
    });
  }

  performers = performers.map(function(user) {
    var arrSum = [];
    for (var i = 0; i < user[0].length; i++) {
      arrSum.push(0);
    }

    var countWrittenTime = 0;

    arrSum.forEach(function(sum, i) {
      user.forEach(function(row) {
        if (typeof(row[i]) === 'string' && /\//.test(row[i])) {
          if (arrSum[i] === 0) arrSum[i] = '0 / 0';
          arrSum[i] = arrSum[i].split('/');
          row[i] = row[i].split('/');
          arrSum[i] = (parseInt(arrSum[i][0], 10) + parseInt(row[i][0], 10)) + ' / ' + (parseInt(arrSum[i][1], 10) + parseInt(row[i][1], 10));
        } else {
          if (i === 0 && parseFloat(row[i]) !== 0) countWrittenTime++;
          arrSum[i] += parseFloat(row[i]);
        }
      });
    });
    arrSum[1] = countWrittenTime ? Math.floor(arrSum[1] / countWrittenTime) : 0;
    return arrSum;
  });

  // print

  var rowI = 2;
  var columnI = 2;
  performers.forEach(function(user) {
    user.forEach(function(value) {
      sheetWeekly.getRange(rowI, columnI++).setValue(value);
    });
    columnI = 2;
    rowI++;
  });
}
