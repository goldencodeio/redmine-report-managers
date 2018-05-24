function processPeriodReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetWeekly = ss.getActiveSheet();
  var allSheets = ss.getSheets();
  var dailySheets = [];
  var performers = [];
  var attendants = [];

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
        login = login.match(/\d{4}/);
        if (!login) return;
        login = login[0];

        OPTIONS.performers.forEach(function(user, iUser) {
          // Browser.msgBox(user + ' | ' + login);
          if (user == login){
            if (performers[iUser] === undefined) performers[iUser] = [];
            performers[iUser].push(row);
          }
        });
        OPTIONS.attendants.forEach(function(user, iUser) {
          if (user == login){
            if (attendants[iUser] === undefined) attendants[iUser] = [];
            attendants[iUser].push(row);
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

    arrSum.forEach(function(sum, i) {
      user.forEach(function(row) {
          arrSum[i] += parseFloat(row[i]);
      });
    });
    arrSum[1] = arrSum[1] / user.length;
    arrSum[9] = arrSum[9] / user.length;
    arrSum[10] = arrSum[10] / user.length;
    return arrSum;
  });

  attendants = attendants.map(function(user) {
    var arrSum = [];
    for (var i = 0; i < user[0].length; i++) {
      arrSum.push(0);
    }

    arrSum.forEach(function(sum, i) {
      user.forEach(function(row) {
          arrSum[i] += parseFloat(row[i]);
      });
    });
    arrSum[1] = arrSum[1] / user.length;
    arrSum[9] = arrSum[9] / user.length;
    arrSum[10] = arrSum[10] / user.length;
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

  rowI += 2;

  attendants.forEach(function(user) {
    user.forEach(function(value) {
      sheetWeekly.getRange(rowI, columnI++).setValue(value);
    });
    columnI = 2;
    rowI++;
  });
}
