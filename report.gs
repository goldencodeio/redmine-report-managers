var REPORT = [
  {
    code: 'work_time',
    name: 'Рабочее\nвремя',
    manual: false
  },
  {
    code: 'written_time',
    name: '% Списанного\nвремени',
    manual: false
  },
  {
    code: 'total_tasks',
    name: 'Всего\nзадач',
    manual: false
  },
  {
    code: 'done_tasks',
    name: 'Выполнено/\nОценено',
    manual: false
  },
  {
    code: 'critical_tasks',
    name: 'Критических/\nОценено',
    manual: false
  },
 {
   code: 'overdue_tasks',
   name: 'Просроченных/\nОценено',
   manual: false
 },
 {
   code: 'paid_separately',
   name: 'Оплачивается\nотдельно/\nОценено',
   manual: false
 },
  {
    code: 'unsubscribed',
    name: 'Неотписано/\nОценено',
    manual: false
  },
 {
   code: 'claims',
   name: 'Претензий/\nОтработано',
   manual: false
 },
 {
   code: 'client_rating_avg',
   name: 'Ср. Оценка\nзаявителя',
   manual: false
 },
 {
   code: 'boss_rating_avg',
   name: 'Ср. Оценка\nведения задачи',
   manual: false
 },
 {
   code: 'forgotten',
   name: 'Забыто',
   manual: true
 },
  {
    code: 'delays',
    name: 'Опозданий\n(мин)',
    manual: true
  },
  {
    code: 'overtime_spent',
    name: 'Переработок\n(мин)',
    manual: true
  },
  {
    code: 'lies',
    name: 'Вранья',
    manual: true
  },
  {
    code: 'points_written_off',
    name: 'Баллов\nсписано по\nпретензиям',
    manual: true
  }
];

function processReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var rowI = 2;
  var columnI = 2;
  var doneIssues = [];
  OPTIONS.performers = OPTIONS.performers.map(function(user, userIndex) {
    user.reports = {};

    REPORT.forEach(function(report) {
      if (!report.manual) {
        var reportValue = getUserReport(report.code, user, userIndex, 'performers');
        user.reports[report] = reportValue;
        if ((Array.isArray(reportValue))) {
          var listUrl = '';
          if ((Array.isArray(reportValue[0]))) {
            reportValue[0].forEach(function(task) {
              listUrl += 'http://redmine.zolotoykod.ru/issues/' + task.id + '\n';
            });
            sheet.getRange(rowI, columnI++).setValue(reportValue[0].length + ' / '+ reportValue[1].length).setNote(listUrl);
          } else {
            reportValue.forEach(function(task) {
              listUrl += 'http://redmine.zolotoykod.ru/issues/' + task.id + '\n';
            });
            sheet.getRange(rowI, columnI++).setValue(reportValue.length).setNote(listUrl);
          }
        } else {
          sheet.getRange(rowI, columnI++).setValue(reportValue);
        }
      } else {
        ss.setNamedRange('manualRange' + rowI + columnI, sheet.getRange(sheet.getRange(rowI, columnI++).getA1Notation()));
      }
    });

    columnI = 2;
    rowI++;
    return user;
  });

  rowI += 2;

  OPTIONS.attendants = OPTIONS.attendants.map(function(user, userIndex) {
    user.reports = {};

    REPORT.forEach(function(report) {
      if (!report.manual) {
        var reportValue = getUserReport(report.code, user, userIndex, 'attendants');
        user.reports[report] = reportValue;
        if ((Array.isArray(reportValue))) {
          var listUrl = '';
          if ((Array.isArray(reportValue[0]))) {
            reportValue[0].forEach(function(task) {
              listUrl += 'http://redmine.zolotoykod.ru/issues/' + task.id + '\n';
            });
            sheet.getRange(rowI, columnI++).setValue(reportValue[0].length + ' / '+ reportValue[1].length).setNote(listUrl);
          } else {
            reportValue.forEach(function(task) {
              listUrl += 'http://redmine.zolotoykod.ru/issues/' + task.id + '\n';
            });
            sheet.getRange(rowI, columnI++).setValue(reportValue.length).setNote(listUrl);
          }
        } else {
          sheet.getRange(rowI, columnI++).setValue(reportValue);
        }
      } else {
        ss.setNamedRange('manualRange' + rowI + columnI, sheet.getRange(sheet.getRange(rowI, columnI++).getA1Notation()));
      }
    });

    columnI = 2;
    rowI++;
    return user;
  });
}

function getUserReport(report, user, userIndex, userType) {
  switch (report) {
    case 'work_time':
      return getWorkTime(userIndex, userType);
      break;

    case 'written_time':
      return getWrittenTime(user, userIndex, userType);
      break;

    case 'total_tasks':
      return getCountTotalTasks(user, userIndex, userType);
      break;

    case 'done_tasks':
      return getCountDoneTasks(user, userIndex, userType);
      break;

    case 'critical_tasks':
      return getCountCriticalTasks(user, userIndex, userType);
      break;

    case 'overdue_tasks':
      return getOverdueTasks(user, userIndex, userType);
      break;

    case 'paid_separately':
      return getPaidSeparatelyTasks(user, userIndex, userType);
      break;

    case 'unsubscribed':
      return getUnsubscribed(user, userIndex, userType);
      break;

    case 'claims':
      return getClaims(user, userIndex, userType);
      break;

    case 'client_rating_avg':
      return getClientRatingAverage(user, userIndex, userType);
      break;

    case 'boss_rating_avg':
      return getBossRatingAverage(user, userIndex, userType);
      break;
  }
}

function getWorkTime(i, userType) {
  if (userType === 'performers') return OPTIONS.performersWorkHours[i];
  if (userType === 'attendants')
    return getHoursByRange(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]);
}

function getWrittenTime(user, i, userType) {
  var date = (userType === 'attendants') ? getDateRangeWithTime(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]) : formatDate(OPTIONS.currentDate);
  var res = APIRequest('time_entries', {query: [
    {key: 'user_id', value: user.id},
    {key: 'spent_on', value: date}
  ]});

  var timeEntries = res.time_entries.reduce(function(a, c) {
    return a + c.hours;
  }, 0);

  if (userType === 'performers')
    if (!OPTIONS.performersWorkHours[i]) return 0;
    return Math.floor(100 / parseInt(OPTIONS.performersWorkHours[i], 10) * timeEntries);

  if (userType === 'attendants')
    return Math.floor(100 / getHoursByRange(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]) * timeEntries);
}

function getCountTotalTasks(user, i, userType) {
  var date = (userType === 'attendants') ? getDateRangeWithTime(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]) : formatDate(OPTIONS.currentDate);
  var res = APIRequest('issues', {query: [
    {key: 'tracker_id', value: '!5'},
    {key: 'assigned_to_id', value: user.id},
    {key: 'status_id', value: '*'},
    {key: 'created_on', value: date}
  ]});
  return res.issues;
}

function getCountDoneTasks(user, userIndex, userType) {
  var filterDate = (userType === 'attendants') ? formatDate(OPTIONS.attendantsStartDate[userIndex]) : formatDate(OPTIONS.currentDate);
  var res = APIRequest('issues', {query: [
    {key: 'tracker_id', value: '!5'},
    {key: 'assigned_to_id', value: user.id},
    {key: 'status_id', value: '*'},
    {key: 'created_on', value: '<=' + filterDate},
    {key: 'updated_on', value: '>=' + filterDate}
  ]});
  Logger.log(res.issues.length);
  var filteredIssues = res.issues.filter(function(task) {
    var resDetail = APIRequestIssueById(task.id, {query: [
      {key: 'include', value: 'journals'}
    ]});
    for (var j = 0; j < resDetail.issue.journals.length; j++) {
      var journal = resDetail.issue.journals[j];
      if (userType === 'attendants') {
        if (Date.parse(journal.created_on) > OPTIONS.attendantsStartDate[userIndex].getTime() && Date.parse(journal.created_on) < OPTIONS.attendantsFinalDate[userIndex].getTime()) {
          for (var d = 0; d < journal.details.length; d++) {
            var detail = journal.details[d];
            if (detail.name === 'status_id' && detail.new_value === '3') return true;
          }
        }
      } else {
        var journalCreateDate = journal.created_on.split('T').shift();
        if (journalCreateDate === formatDate(OPTIONS.currentDate)) {
          for (var d = 0; d < journal.details.length; d++) {
            var detail = journal.details[d];
            if (detail.name === 'status_id' && detail.new_value === '3') return true;
          }
        }
      }
    }
    // resDetail.issue.journals.forEach(function(journal) {
    //   if (userType === 'attendants') {
    //     if (Date.parse(journal.created_on) > OPTIONS.attendantsStartDate[userIndex].getTime() && Date.parse(journal.created_on) < OPTIONS.attendantsFinalDate[userIndex].getTime()) {
    //       journal.details.forEach(function(detail) {
    //         if (detail.name === 'status_id' && detail.new_value === '3') return true;
    //       });
    //     }
    //   } else {
    //     var journalCreateDate = journal.created_on.split('T').shift();
    //     if (journalCreateDate === formatDate(OPTIONS.currentDate)) {
    //       journal.details.forEach(function(detail) {
    //         if (detail.name === 'status_id' && detail.new_value === '3') return true;
    //       });
    //     }
    //   }
    // });
    return false;
  });

  doneIssues = filteredIssues;

  var filteredIssuesWithRate = filteredIssues.filter(function(item) {
    if (item.custom_fields.find(function(i) {return i.id === 7}).value !== '')
      return true;
  });
  return [filteredIssues, filteredIssuesWithRate];
}

function getCountCriticalTasks(user, i, userType) {
  // var date = (userType === 'attendants') ? getDateRangeWithTime(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]) : formatDate(OPTIONS.currentDate);
  // var res = APIRequest('issues', {query: [
  //   {key: 'assigned_to_id', value: user.id},
  //   {key: 'status_id', value: 'closed'},
  //   {key: 'closed_on', value: date},
  //   {key: 'priority_id', value: '5'}
  // ]});
  //
  // var res1 = APIRequest('issues', {query: [
  //   {key: 'assigned_to_id', value: user.id},
  //   {key: 'status_id', value: 'closed'},
  //   {key: 'closed_on', value: date},
  //   {key: 'priority_id', value: '4'}
  // ]});

  var criticalTasks = doneIssues.filter(function(item) {
    if (item.priority.id > 3) return true;
  });

  var criticalTasksWithRate = criticalTasks.filter(function(item) {
    if (item.custom_fields.find(function(i) {return i.id === 7}).value !== '')
      return true;
  });

  return [criticalTasks, criticalTasksWithRate];
}

function getOverdueTasks(user, i, userType) {
  // if (userType === 'attendants') {
  //     var date = getDateRangeWithTime(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]);
  // } else {
  //     var prevDate = new Date(OPTIONS.currentDate.getTime());
  //     prevDate.setDate(prevDate.getDate() - 1);
  //     var date = '<=' + formatDate(prevDate);
  // }

  // var res = APIRequest('issues', {query: [
  //   {key: 'assigned_to_id', value: user.id},
  //   {key: 'status_id', value: 'closed'},
  //   {key: 'closed_on', value: formatDate(OPTIONS.currentDate)},
  //   {key: 'due_date', value: date}
  // ]});

  // return res.issues;

  var overdueTasks = doneIssues.filter(function(item) {
    if (userType === 'attendants') {
      if (item.due_date && (Date.parse(item.due_date) + 1000 * 60 * 60 * 24) < OPTIONS.attendantsFinalDate[i].getTime())
        return true;
    } else {
      if (item.due_date && (Date.parse(item.due_date) + 1000 * 60 * 60 * 24) < OPTIONS.currentDate.getTime())
        return true;
    }
  });

  var overdueTasksWithRate = overdueTasks.filter(function(item) {
    if (item.custom_fields.find(function(i) {return i.id === 7}).value !== '')
      return true;
  });

  return [overdueTasks, overdueTasksWithRate];
}

function getPaidSeparatelyTasks(user, i, userType) {
  // var date = (userType === 'attendants') ? getDateRangeWithTime(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]) : formatDate(OPTIONS.currentDate);
  // var res = APIRequest('issues', {query: [
  //   {key: 'assigned_to_id', value: user.id},
  //   {key: 'status_id', value: 'closed'},
  //   {key: 'closed_on', value: date},
  //   {key: 'cf_24', value: 'Единовременная услуга (К оплате)'}
  // ]});
  //
  // return res.issues;

  var paidSeparatelyTasks = doneIssues.filter(function(item) {
    var tariff = item.custom_fields.find(function(i) {return i.id === 24});
    if (tariff && tariff.value === 'Единовременная услуга (К оплате)') return true;
  });

  var paidSeparatelyTasksWithRate = paidSeparatelyTasks.filter(function(item) {
    if (item.custom_fields.find(function(i) {return i.id === 7}).value !== '')
      return true;
  });

  return [paidSeparatelyTasks, paidSeparatelyTasksWithRate];
}

function getUnsubscribed(user, i, userType) {
  // var date = (userType === 'attendants') ? getDateRangeWithTime(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]) : formatDate(OPTIONS.currentDate);
  // var res = APIRequest('issues', {query: [
  //   {key: 'assigned_to_id', value: user.id},
  //   {key: 'status_id', value: 'closed'},
  //   {key: 'closed_on', value: date},
  //   {key: 'cf_1', value: '1'}
  // ]});

  var unsubscribed = doneIssues.filter(function(item) {
    if (item.custom_fields.find(function(i) {return i.id === 1}).value === '')
      return true;
  });

  var unsubscribedWithRate = unsubscribed.filter(function(item) {
    if (item.custom_fields.find(function(i) {return i.id === 7}).value !== '')
      return true;
  });

  return [unsubscribed, unsubscribedWithRate];
}

function getClaims(user, i, userType) {
  var date = (userType === 'attendants') ? getDateRangeWithTime(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]) : formatDate(OPTIONS.currentDate);
  var allClaims = APIRequest('issues', {query: [
    {key: 'tracker_id', value: 5},
    {key: 'assigned_to_id', value: user.id},
    {key: 'status_id', value: '*'},
    {key: 'created_on', value: date}
  ]});

  var closedClaims = APIRequest('issues', {query: [
    {key: 'tracker_id', value: 5},
    {key: 'assigned_to_id', value: user.id},
    {key: 'status_id', value: 'closed'},
    {key: 'created_on', value: date}
  ]});

  return [allClaims.issues, closedClaims.issues];
}

function getClientRatingAverage(user, i, userType) {
  // var date = (userType === 'attendants') ? getDateRangeWithTime(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]) : formatDate(OPTIONS.currentDate);
  // var res = APIRequest('issues', {query: [
  //   {key: 'assigned_to_id', value: user.id},
  //   {key: 'status_id', value: 'closed'},
  //   {key: 'closed_on', value: date},
  //   {key: 'cf_7', value: '*'}
  // ]});

  var doneIssuesWithRate = doneIssues.filter(function(item) {
    if (item.custom_fields.find(function(i) {return i.id === 7}).value !== '')
      return true;
  });

  var sum = doneIssuesWithRate.reduce(function(a, c) {
    return a + parseInt(c.custom_fields.find(function(i) {return i.id === 7}).value, 10);
  }, 0);
  return doneIssuesWithRate.length ? sum / doneIssuesWithRate.length : 0;
}

function getBossRatingAverage(user, i, userType) {
  // var date = (userType === 'attendants') ? getDateRangeWithTime(OPTIONS.attendantsStartDate[i], OPTIONS.attendantsFinalDate[i]) : formatDate(OPTIONS.currentDate);
  // var res = APIRequest('issues', {query: [
  //   {key: 'assigned_to_id', value: user.id},
  //   {key: 'status_id', value: 'closed'},
  //   {key: 'closed_on', value: date},
  //   {key: 'cf_8', value: '*'}
  // ]});

  var doneIssuesWithRate = doneIssues.filter(function(item) {
    if (item.custom_fields.find(function(i) {return i.id === 8}).value !== '')
      return true;
  });

  var sum = doneIssuesWithRate.reduce(function(a, c) {
    return a + parseInt(c.custom_fields.find(function(i) {return i.id === 8}).value, 10);
  }, 0);
  return doneIssuesWithRate.length ? sum / doneIssuesWithRate.length : 0;
}
