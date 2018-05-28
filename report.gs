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
    code: 'projects_processed',
    name: 'Проектов\nобработано',
    manual: false
  },
  {
    code: 'projects_dev',
    name: 'Разработка',
    manual: false
  },
  {
    code: 'projects_integration',
    name: 'Внедрение',
    manual: false
  },
  {
    code: 'done_tasks',
    name: 'Выполнено\nзадач/\nОценено',
    manual: false
  },
 {
   code: 'paid_separately',
   name: 'Задач\nобработано\n(опл. отдельно)\n/ Оценено',
   manual: false
 },
 {
   code: 'claims',
   name: 'Претензий/\nОтработано',
   manual: false
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
        var reportValue = getUserReport(report.code, user, userIndex);
        user.reports[report] = reportValue;
        if ((Array.isArray(reportValue))) {
          var listUrl = '';
          if ((Array.isArray(reportValue[0]))) {
            reportValue[0].forEach(function(item) {
              if (report.code.search(/projects_/) === 0)
                listUrl += 'http://redmine.zolotoykod.ru/projects/' + item.identifier + '\n';
              else
                listUrl += 'http://redmine.zolotoykod.ru/issues/' + item.id + '\n';
            });
            sheet.getRange(rowI, columnI++).setValue(reportValue[0].length + ' / '+ reportValue[1].length).setNote(listUrl);
          } else {
            reportValue.forEach(function(item) {
              if (report.code.search(/projects_/) === 0)
                listUrl += 'http://redmine.zolotoykod.ru/projects/' + item.identifier + '\n';
              else
                listUrl += 'http://redmine.zolotoykod.ru/issues/' + item.id + '\n';
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

function getUserReport(report, user, userIndex) {
  switch (report) {
    case 'work_time':
      return getWorkTime(userIndex);
      break;

    case 'written_time':
      return getWrittenTime(user, userIndex);
      break;

    case 'projects_processed':
      return getProjectsProcessed(user);
      break;

    case 'projects_dev':
      return getProjectsProcessedDev(user);
      break;

    case 'projects_integration':
      return getProjectsProcessedIntegrat(user);
      break;

    case 'done_tasks':
      return getDoneTasks(user);
      break;

    case 'paid_separately':
      return getPaidSeparatelyTasks(user);
      break;

    case 'claims':
      return getClaims(user);
      break;
  }
}

function getWorkTime(i) {
  return OPTIONS.performersWorkHours[i];
}

function getWrittenTime(user, i) {
  var res = APIRequest('time_entries', {query: [
    {key: 'user_id', value: user.id},
    {key: 'spent_on', value: formatDate(OPTIONS.currentDate)}
  ]});

  var timeEntries = res.time_entries.reduce(function(a, c) {
    return a + c.hours;
  }, 0);

  if (!OPTIONS.performersWorkHours[i]) return 0;
  return Math.floor(100 / parseInt(OPTIONS.performersWorkHours[i], 10) * timeEntries);
}

function getProjectsProcessed(user) {
  var res = APIRequest('time_entries', {query: [
    {key: 'user_id', value: user.id},
    {key: 'spent_on', value: formatDate(OPTIONS.currentDate)}
  ]});

  if (!res.time_entries.length) return 0;

  var projectsId = res.time_entries.map(function(item) {
    return item.project.id;
  });

  projectsId = projectsId.filter(function(item, pos) {
    return projectsId.indexOf(item) == pos;
  });

  return projectsId.map(function(id) {
    var resProject = APIRequestById('projects', id);
    return resProject.project;
  });
}

function getProjectsProcessedDev(user) {
  var res = APIRequest('time_entries', {query: [
    {key: 'user_id', value: user.id},
    {key: 'spent_on', value: formatDate(OPTIONS.currentDate)}
  ]});

  if (!res.time_entries.length) return 0;

  var issuesId = res.time_entries.map(function(item) {
    return item.issue.id;
  });

  issuesId = issuesId.filter(function(item, pos) {
    return issuesId.indexOf(item) == pos;
  });

  var issues = issuesId.map(function(id) {
    var resIssue = APIRequestById('issues', id);
    return resIssue.issue;
  });

  issues = issues.filter(function(item) {
    return item.tracker.id === 6;
  });

  var projectsId = issues.map(function(item) {
    return item.project.id;
  });

  projectsId = projectsId.filter(function(item, pos) {
    return projectsId.indexOf(item) == pos;
  });

  return projectsId.map(function(id) {
    var resProject = APIRequestById('projects', id);
    return resProject.project;
  });
}

function getProjectsProcessedIntegrat(user) {
  var res = APIRequest('time_entries', {query: [
    {key: 'user_id', value: user.id},
    {key: 'spent_on', value: formatDate(OPTIONS.currentDate)}
  ]});

  if (!res.time_entries.length) return 0;

  var issuesId = res.time_entries.map(function(item) {
    return item.issue.id;
  });

  issuesId = issuesId.filter(function(item, pos) {
    return issuesId.indexOf(item) == pos;
  });

  var issues = issuesId.map(function(id) {
    var resIssue = APIRequestById('issues', id);
    return resIssue.issue;
  });

  issues = issues.filter(function(item) {
    return item.tracker.id === 8;
  });

  var projectsId = issues.map(function(item) {
    return item.project.id;
  });

  projectsId = projectsId.filter(function(item, pos) {
    return projectsId.indexOf(item) == pos;
  });

  return projectsId.map(function(id) {
    var resProject = APIRequestById('projects', id);
    return resProject.project;
  });
}

function getDoneTasks(user) {
  var res = APIRequest('issues', {query: [
    {key: 'tracker_id', value: '!5'},
    {key: 'assigned_to_id', value: user.id},
    {key: 'status_id', value: '*'},
    {key: 'created_on', value: '<=' + formatDate(OPTIONS.currentDate)},
    {key: 'updated_on', value: '>=' + formatDate(OPTIONS.currentDate)}
  ]});

  var filteredIssues = res.issues.filter(function(task) {
    var resDetail = APIRequestById('issues', task.id, {query: [
      {key: 'include', value: 'journals'}
    ]});

    for (var j = 0; j < resDetail.issue.journals.length; j++) {
      var journal = resDetail.issue.journals[j];
      var journalCreateDate = journal.created_on.split('T').shift();
      if (journalCreateDate === formatDate(OPTIONS.currentDate)) {
        for (var d = 0; d < journal.details.length; d++) {
          var detail = journal.details[d];
          if (detail.name === 'status_id' && detail.new_value === '3') return true;
        }
      }
    }
    return false;
  });

  doneIssues = filteredIssues;

  var filteredIssuesWithRate = filteredIssues.filter(function(item) {
    if (item.custom_fields.find(function(i) {return i.id === 7}).value !== '')
      return true;
  });
  return [filteredIssues, filteredIssuesWithRate];
}

function getPaidSeparatelyTasks(user) {
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

function getClaims(user) {
  var allClaims = APIRequest('issues', {query: [
    {key: 'tracker_id', value: 5},
    {key: 'assigned_to_id', value: user.id},
    {key: 'status_id', value: '*'},
    {key: 'created_on', value: formatDate(OPTIONS.currentDate)}
  ]});

  var closedClaims = APIRequest('issues', {query: [
    {key: 'tracker_id', value: 5},
    {key: 'assigned_to_id', value: user.id},
    {key: 'status_id', value: 'closed'},
    {key: 'created_on', value: formatDate(OPTIONS.currentDate)}
  ]});

  return [allClaims.issues, closedClaims.issues];
}
