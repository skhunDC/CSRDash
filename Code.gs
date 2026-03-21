const APP_CONFIG = Object.freeze({
  spreadsheetName: 'CSR Dashboard Data',
  allowedUsers: Object.freeze([
    'skhun@dublincleaners.com',
    'skhun1@dublincleaners.com',
    'ss.sku@gmail.com',
  ]),
  refreshIntervalMs: 60 * 1000,
  logoUrl:
    'https://www.dublincleaners.com/wp-content/uploads/2024/12/Dublin-Logos-stacked.png',
  sheets: Object.freeze({
    sales: 'Sales',
    performance: 'CSR_Performance',
    tasks: 'Tasks',
    schedule: 'Schedule',
    competitions: 'Competitions',
    employeeOfWeek: 'Employee_Of_Week',
  }),
  competitionTabs: Object.freeze(['Conversions', 'Patio Signups', 'Alterations']),
  tasks: Object.freeze([
    'Front Counter',
    'Phones',
    'Counter Bags',
    'Wedding Gowns',
    'Shirt Hangers',
    'Shake Shirts',
    'Hand Pressed',
    'Sort Hangers',
    'Paper Hangers',
    'Lobby',
    'Breakroom',
    'Vacuum',
    'Trash',
    '# of Carts',
  ]),
});

function doGet(e) {
  if (e && e.parameter && e.parameter.view === 'print') {
    return doGetPrint();
  }

  const template = HtmlService.createTemplateFromFile('index');
  const auth = getAuthorizationState_();

  template.initialState = JSON.stringify({
    auth,
    refreshIntervalMs: APP_CONFIG.refreshIntervalMs,
    logoUrl: APP_CONFIG.logoUrl,
    printUrl: buildAppUrl_('print'),
  });

  return template
    .evaluate()
    .setTitle('Dublin Cleaners CSR Dashboard')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function doGetPrint() {
  const template = HtmlService.createTemplateFromFile('print');
  const auth = getAuthorizationState_();

  template.initialState = JSON.stringify({
    auth,
    refreshIntervalMs: APP_CONFIG.refreshIntervalMs,
    logoUrl: APP_CONFIG.logoUrl,
  });

  return template
    .evaluate()
    .setTitle('CSR Dashboard Print View')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getInitialDashboardData() {
  authorizeRequest_();
  return getDashboardData_();
}

function refreshDashboardData() {
  authorizeRequest_();
  return getDashboardData_();
}

function updateTaskStatus(taskName, completed, assignedCsr) {
  authorizeRequest_();
  if (!taskName) {
    throw new Error('Task name is required.');
  }

  const spreadsheet = getOrCreateSpreadsheet_();
  const sheet = spreadsheet.getSheetByName(APP_CONFIG.sheets.tasks);
  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    if (String(values[rowIndex][1]).trim() === taskName) {
      sheet.getRange(rowIndex + 1, 3).setValue(assignedCsr || '');
      sheet.getRange(rowIndex + 1, 4).setValue(Boolean(completed));
      sheet.getRange(rowIndex + 1, 5).setValue(new Date());
      return getTasks_();
    }
  }

  throw new Error('Task not found: ' + taskName);
}

function seedDashboardData() {
  authorizeRequest_();
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureAllSheets_(spreadsheet, true);
  return {
    ok: true,
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
  };
}

function getDashboardData_() {
  const spreadsheet = getOrCreateSpreadsheet_();
  ensureAllSheets_(spreadsheet, false);

  return {
    generatedAt: new Date().toISOString(),
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    hero: getHeroData_(),
    performance: getPerformanceData_(),
    tasks: getTasks_(),
    schedule: getScheduleData_(),
    competitions: getCompetitionData_(),
    employeeOfWeek: getEmployeeOfWeek_(),
  };
}

function getAuthorizationState_() {
  const email = normalizeEmail_(Session.getActiveUser().getEmail());
  const authorized = email && APP_CONFIG.allowedUsers.indexOf(email) !== -1;

  return {
    email: email || '',
    authorized,
    allowedUsers: APP_CONFIG.allowedUsers.slice(),
  };
}

function authorizeRequest_() {
  const auth = getAuthorizationState_();
  if (!auth.authorized) {
    throw new Error('Unauthorized');
  }
}

function getOrCreateSpreadsheet_() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let spreadsheetId = scriptProperties.getProperty('CSR_DASHBOARD_SPREADSHEET_ID');
  let spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : null;

  if (!spreadsheet) {
    const existingFiles = DriveApp.getFilesByName(APP_CONFIG.spreadsheetName);
    spreadsheet = existingFiles.hasNext()
      ? SpreadsheetApp.open(existingFiles.next())
      : SpreadsheetApp.create(APP_CONFIG.spreadsheetName);

    scriptProperties.setProperty('CSR_DASHBOARD_SPREADSHEET_ID', spreadsheet.getId());
  }

  return spreadsheet;
}

function ensureAllSheets_(spreadsheet, forceSeed) {
  ensureSalesSheet_(spreadsheet, forceSeed);
  ensurePerformanceSheet_(spreadsheet, forceSeed);
  ensureTasksSheet_(spreadsheet, forceSeed);
  ensureScheduleSheet_(spreadsheet, forceSeed);
  ensureCompetitionsSheet_(spreadsheet, forceSeed);
  ensureEmployeeSheet_(spreadsheet, forceSeed);
}

function ensureSalesSheet_(spreadsheet, forceSeed) {
  const headers = ['Week Label', 'Store', 'Current Year Sales', 'Last Year Sales'];
  const rows = [
    ['Week 12', 'Dublin', 18540, 17110],
    ['Week 12', 'Pleasanton', 16220, 15330],
    ['Week 12', 'San Ramon', 14860, 15010],
  ];
  ensureSheetWithData_(spreadsheet, APP_CONFIG.sheets.sales, headers, rows, forceSeed);
}

function ensurePerformanceSheet_(spreadsheet, forceSeed) {
  const headers = ['Date', 'CSR', 'Sales', 'Hours'];
  const yesterday = shiftDate_(new Date(), -1);
  const rows = [
    [yesterday, 'Sophia', 820, 6.5],
    [yesterday, 'Mateo', 770, 7],
    [yesterday, 'Avery', 690, 6],
    [yesterday, 'Jordan', 510, 6.5],
  ];
  ensureSheetWithData_(spreadsheet, APP_CONFIG.sheets.performance, headers, rows, forceSeed);
}

function ensureTasksSheet_(spreadsheet, forceSeed) {
  const headers = ['Display Order', 'Task', 'Assigned CSR', 'Completed', 'Updated At'];
  const rows = APP_CONFIG.tasks.map(function(task, index) {
    return [index + 1, task, '', false, ''];
  });
  ensureSheetWithData_(spreadsheet, APP_CONFIG.sheets.tasks, headers, rows, forceSeed);
}

function ensureScheduleSheet_(spreadsheet, forceSeed) {
  const headers = ['Date', 'CSR', 'Status', 'Notes'];
  const today = resetTime_(new Date());
  const tomorrow = shiftDate_(today, 1);
  const rows = [
    [today, 'Sophia', 'WORKING', 'Opening shift'],
    [today, 'Mateo', 'OFF', 'Vacation'],
    [today, 'Avery', 'WORKING', 'Mid shift'],
    [tomorrow, 'Jordan', 'OFF', 'Personal day'],
    [tomorrow, 'Sophia', 'WORKING', 'Closing shift'],
  ];
  ensureSheetWithData_(spreadsheet, APP_CONFIG.sheets.schedule, headers, rows, forceSeed);
}

function ensureCompetitionsSheet_(spreadsheet, forceSeed) {
  const headers = ['Category', 'CSR', 'Value', 'Goal', 'Notes'];
  const rows = [
    ['Conversions', 'Sophia', 18, 25, 'Strong close rate'],
    ['Conversions', 'Mateo', 16, 25, 'Steady'],
    ['Patio Signups', 'Avery', 12, 20, 'Leading signups'],
    ['Patio Signups', 'Jordan', 9, 20, 'Needs support'],
    ['Alterations', 'Sophia', 14, 18, 'Strong upsell'],
    ['Alterations', 'Mateo', 11, 18, 'Solid week'],
  ];
  ensureSheetWithData_(spreadsheet, APP_CONFIG.sheets.competitions, headers, rows, forceSeed);
}

function ensureEmployeeSheet_(spreadsheet, forceSeed) {
  const headers = ['Name', 'Quote', 'Image URL', 'Highlight'];
  const rows = [[
    'Sophia Ramirez',
    'Great service is about making every guest feel expected, not just welcomed.',
    '',
    'Top guest feedback and strongest conversion improvement this week.',
  ]];
  ensureSheetWithData_(spreadsheet, APP_CONFIG.sheets.employeeOfWeek, headers, rows, forceSeed);
}

function ensureSheetWithData_(spreadsheet, sheetName, headers, rows, forceSeed) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }

  const hasData = sheet.getLastRow() > 0 && sheet.getLastColumn() > 0;
  if (forceSeed || !hasData) {
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    if (sheetName === APP_CONFIG.sheets.performance || sheetName === APP_CONFIG.sheets.schedule) {
      sheet.getRange(2, 1, Math.max(rows.length, 1), 1).setNumberFormat('yyyy-mm-dd');
    }
    styleSheet_(sheet, headers.length);
  }
}

function styleSheet_(sheet, columnCount) {
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#102542')
    .setFontColor('#f4f7fb');
  sheet.autoResizeColumns(1, columnCount);
}

function getHeroData_() {
  const values = getOrCreateSpreadsheet_()
    .getSheetByName(APP_CONFIG.sheets.sales)
    .getDataRange()
    .getValues();

  const rows = values.slice(1).filter(function(row) {
    return row[0] && row[1];
  });

  const totals = rows.reduce(
    function(acc, row) {
      acc.current += Number(row[2]) || 0;
      acc.last += Number(row[3]) || 0;
      acc.stores.push({
        store: row[1],
        currentYear: Number(row[2]) || 0,
        lastYear: Number(row[3]) || 0,
      });
      acc.weekLabel = row[0] || acc.weekLabel;
      return acc;
    },
    { current: 0, last: 0, stores: [], weekLabel: 'Current Week' }
  );

  totals.stores = totals.stores.map(function(store) {
    store.deltaPercent = calculateDeltaPercent_(store.currentYear, store.lastYear);
    return store;
  }).sort(function(a, b) {
    return b.currentYear - a.currentYear;
  });

  return {
    weekLabel: totals.weekLabel,
    currentYearSales: totals.current,
    lastYearSales: totals.last,
    deltaPercent: calculateDeltaPercent_(totals.current, totals.last),
    stores: totals.stores,
  };
}

function getPerformanceData_() {
  const values = getOrCreateSpreadsheet_()
    .getSheetByName(APP_CONFIG.sheets.performance)
    .getDataRange()
    .getValues();
  const yesterdayKey = formatDateKey_(shiftDate_(new Date(), -1));

  const entries = values.slice(1).filter(function(row) {
    return row[0] && formatDateKey_(row[0]) === yesterdayKey;
  }).map(function(row) {
    const sales = Number(row[2]) || 0;
    const hours = Number(row[3]) || 0;
    const dollarsPerHour = hours ? sales / hours : 0;

    return {
      date: formatDisplayDate_(row[0]),
      csr: row[1],
      sales,
      hours,
      dollarsPerHour,
    };
  }).sort(function(a, b) {
    return b.dollarsPerHour - a.dollarsPerHour;
  }).map(function(entry, index, all) {
    const top = all[0] ? all[0].dollarsPerHour : entry.dollarsPerHour;
    const ratio = top ? entry.dollarsPerHour / top : 0;
    entry.rank = index + 1;
    entry.tier = ratio >= 0.9 ? 'elite' : ratio >= 0.75 ? 'strong' : 'developing';
    return entry;
  });

  return {
    dateLabel: formatDisplayDate_(shiftDate_(new Date(), -1)),
    entries,
  };
}

function getTasks_() {
  const values = getOrCreateSpreadsheet_()
    .getSheetByName(APP_CONFIG.sheets.tasks)
    .getDataRange()
    .getValues();

  const entries = values.slice(1).filter(function(row) {
    return row[1];
  }).map(function(row) {
    return {
      order: Number(row[0]) || 0,
      task: row[1],
      assignedCsr: row[2] || '',
      completed: normalizeBoolean_(row[3]),
      updatedAt: row[4] ? formatTimestamp_(row[4]) : '',
    };
  }).sort(function(a, b) {
    return a.order - b.order;
  });

  const summary = entries.reduce(
    function(acc, entry) {
      if (entry.completed) {
        acc.completed += 1;
      }
      if (!entry.assignedCsr) {
        acc.unassigned += 1;
      }
      return acc;
    },
    { completed: 0, unassigned: 0 }
  );

  summary.total = entries.length;
  summary.progressPercent = entries.length
    ? Math.round((summary.completed / entries.length) * 100)
    : 0;

  return {
    entries,
    summary,
  };
}

function getScheduleData_() {
  const values = getOrCreateSpreadsheet_()
    .getSheetByName(APP_CONFIG.sheets.schedule)
    .getDataRange()
    .getValues();
  const today = formatDateKey_(new Date());
  const tomorrow = formatDateKey_(shiftDate_(new Date(), 1));

  const base = {
    today: { label: 'Today', off: [] },
    tomorrow: { label: 'Tomorrow', off: [] },
  };

  values.slice(1).forEach(function(row) {
    if (!row[0] || String(row[2]).toUpperCase() !== 'OFF') {
      return;
    }

    const dateKey = formatDateKey_(row[0]);
    const entry = {
      csr: row[1],
      notes: row[3] || '',
    };

    if (dateKey === today) {
      base.today.off.push(entry);
    }
    if (dateKey === tomorrow) {
      base.tomorrow.off.push(entry);
    }
  });

  return base;
}

function getCompetitionData_() {
  const values = getOrCreateSpreadsheet_()
    .getSheetByName(APP_CONFIG.sheets.competitions)
    .getDataRange()
    .getValues();

  const categories = {};
  APP_CONFIG.competitionTabs.forEach(function(tab) {
    categories[tab] = [];
  });

  values.slice(1).forEach(function(row) {
    const category = row[0];
    if (!category || !categories[category]) {
      return;
    }

    const value = Number(row[2]) || 0;
    const goal = Number(row[3]) || 0;
    categories[category].push({
      csr: row[1],
      value,
      goal,
      progressPercent: goal ? Math.min(100, Math.round((value / goal) * 100)) : 0,
      notes: row[4] || '',
    });
  });

  Object.keys(categories).forEach(function(key) {
    categories[key].sort(function(a, b) {
      return b.value - a.value;
    });
  });

  return {
    tabs: APP_CONFIG.competitionTabs.slice(),
    categories,
  };
}

function getEmployeeOfWeek_() {
  const values = getOrCreateSpreadsheet_()
    .getSheetByName(APP_CONFIG.sheets.employeeOfWeek)
    .getDataRange()
    .getValues();
  const row = values[1] || [];

  return {
    name: row[0] || 'TBD',
    quote: row[1] || 'Add a quote in the Employee_Of_Week sheet.',
    imageUrl: row[2] || '',
    highlight: row[3] || '',
  };
}

function calculateDeltaPercent_(current, previous) {
  if (!previous) {
    return current ? 100 : 0;
  }
  return ((current - previous) / previous) * 100;
}

function normalizeEmail_(email) {
  return String(email || '').trim().toLowerCase();
}

function normalizeBoolean_(value) {
  return value === true || String(value).toLowerCase() === 'true';
}

function formatCurrency_(value) {
  return Utilities.formatString('$%,.0f', Number(value) || 0);
}

function formatPercent_(value) {
  const sign = value > 0 ? '+' : '';
  return sign + value.toFixed(1) + '%';
}

function formatDateKey_(value) {
  return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function formatDisplayDate_(value) {
  return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'EEE, MMM d');
}

function formatTimestamp_(value) {
  return Utilities.formatDate(new Date(value), Session.getScriptTimeZone(), 'MMM d, h:mm a');
}

function shiftDate_(date, days) {
  const copy = resetTime_(date);
  copy.setDate(copy.getDate() + days);
  return copy;
}

function resetTime_(date) {
  const copy = new Date(date);
  copy.setHours(0, 0, 0, 0);
  return copy;
}

function buildAppUrl_(view) {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    return '';
  }
  return view === 'print' ? url + '?view=print' : url;
}
