const CONFIG = {
  APP_NAME: 'PLT CSR Dashboard',
  ALLOWED_EMAILS: ['skhun@dublincleaners.com', 'ss.sku@gmail.com'],
  REFRESH_INTERVAL_MS: 5 * 60 * 1000,
  PERFORMANCE_GOAL: 150,
  DEFAULT_STORES: ['Dublin', 'Pleasanton', 'San Ramon'],
  DB_PROPERTY_KEY: 'CSR_DASHBOARD_DB_ID',
  CACHE_KEY_DASHBOARD: 'csr_dashboard_payload',
  CACHE_SECONDS: 60,
  SHEETS: {
    SALES_WEEKLY: 'Sales_Weekly',
    CSR_PERFORMANCE: 'CSR_Performance',
    CSR_SCHEDULE: 'CSR_Schedule',
    CSR_RECOGNITION: 'CSR_Recognition',
    CSR_COMPETITIONS: 'CSR_Competitions',
    CLEANING_CHECKLIST: 'Cleaning_Checklist',
    AUTH_LOG: 'Auth_Log'
  }
};

function doGet(e) {
  ensureDatabase();
  const page = e && e.parameter && e.parameter.page;
  const file = page === 'print' ? 'print' : 'index';
  return HtmlService.createTemplateFromFile(file)
    .evaluate()
    .setTitle(CONFIG.APP_NAME)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

function getAuthStatus() {
  const email = getCurrentUserEmail_();
  const authorized = isAuthorizedEmail_(email);
  logAuthAttempt(email || 'unknown', authorized);
  return {
    authorized: authorized,
    email: email
  };
}

function getClientConfig() {
  authorizeOrThrow_();
  return {
    appName: CONFIG.APP_NAME,
    refreshIntervalMs: CONFIG.REFRESH_INTERVAL_MS,
    performanceGoal: CONFIG.PERFORMANCE_GOAL,
    defaultStores: CONFIG.DEFAULT_STORES
  };
}

function getDashboardData() {
  authorizeOrThrow_();
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.CACHE_KEY_DASHBOARD);
  if (cached) {
    return JSON.parse(cached);
  }

  const db = openDatabase_();
  const salesComparisonPayload = buildSalesComparison_(db);

  const payload = {
    generatedAt: new Date().toISOString(),
    salesComparison: salesComparisonPayload.rows,
    salesComparisonYears: salesComparisonPayload.years,
    csrPerformanceYesterday: buildCsrPerformanceYesterday_(db),
    scheduleWeek: buildScheduleWeek_(db),
    csrOfWeek: buildCsrOfWeek_(db),
    competitionsSnapshot: buildCompetitionsSnapshot_(db),
    checklistSnapshot: buildChecklistSnapshot_(db),
    stores: CONFIG.DEFAULT_STORES,
    goalPerHour: CONFIG.PERFORMANCE_GOAL
  };

  cache.put(CONFIG.CACHE_KEY_DASHBOARD, JSON.stringify(payload), CONFIG.CACHE_SECONDS);
  return payload;
}

function saveSalesEntry(payload) {
  authorizeOrThrow_();
  if (!payload || !payload.weekStart || !payload.store || payload.year === undefined) {
    throw new Error('Invalid payload for sales entry.');
  }

  const db = openDatabase_();
  const sheet = db.getSheetByName(CONFIG.SHEETS.SALES_WEEKLY);
  upsertSheetRow_(sheet, {
    WeekStart: payload.weekStart,
    Store: payload.store,
    Year: Number(payload.year),
    Sales: Number(payload.sales || 0)
  }, ['WeekStart', 'Store', 'Year']);

  clearDashboardCache_();
  return { success: true };
}

function saveCsrRecognition(payload) {
  authorizeOrThrow_();
  if (!payload || !payload.weekStart) {
    throw new Error('Invalid payload for CSR recognition.');
  }

  const db = openDatabase_();
  const sheet = db.getSheetByName(CONFIG.SHEETS.CSR_RECOGNITION);
  const imageData = String(payload.imageData || '').trim();
  const imageUrl = String(payload.imageUrl || '').trim();

  upsertSheetRow_(sheet, {
    WeekStart: payload.weekStart,
    CSRName: payload.csrName || 'TBD',
    Store: payload.store || '-',
    Quote: payload.quote || 'Recognizing excellence in service every week.',
    ImageUrl: imageData || imageUrl
  }, ['WeekStart']);

  clearDashboardCache_();
  return { success: true };
}

function savePerformanceEntry(payload) {
  authorizeOrThrow_();
  if (!payload || !payload.date || !payload.store || !payload.csrName) {
    throw new Error('Invalid payload for performance entry.');
  }

  const db = openDatabase_();
  const sheet = db.getSheetByName(CONFIG.SHEETS.CSR_PERFORMANCE);
  upsertSheetRow_(sheet, {
    Date: payload.date,
    Store: payload.store,
    CSRName: payload.csrName,
    Sales: Number(payload.sales || 0),
    Hours: Number(payload.hours || 0)
  }, ['Date', 'Store', 'CSRName']);

  clearDashboardCache_();
  return { success: true };
}

function saveScheduleEntry(payload) {
  authorizeOrThrow_();
  if (!payload || !payload.date || !payload.store || !payload.csrName) {
    throw new Error('Invalid payload for schedule entry.');
  }

  const db = openDatabase_();
  const sheet = db.getSheetByName(CONFIG.SHEETS.CSR_SCHEDULE);
  upsertSheetRow_(sheet, {
    Date: payload.date,
    Store: payload.store,
    CSRName: payload.csrName,
    ShiftStatus: payload.shiftStatus || 'OFF'
  }, ['Date', 'Store', 'CSRName']);

  clearDashboardCache_();
  return { success: true };
}

function getCompetitionsData() {
  authorizeOrThrow_();
  const db = openDatabase_();
  const rows = getDataRows_(db.getSheetByName(CONFIG.SHEETS.CSR_COMPETITIONS));
  const metricsMap = {};
  rows.forEach(function (row) {
    const metric = row.metric;
    if (!metricsMap[metric]) {
      metricsMap[metric] = [];
    }
    metricsMap[metric].push({
      weekStart: row.weekStart,
      store: row.store,
      csrName: row.csrName,
      value: Number(row.value || 0),
      updatedAt: row.updatedAt
    });
  });

  const metrics = Object.keys(metricsMap).sort();
  const leaders = metrics.map(function (metric) {
    const sorted = metricsMap[metric].slice().sort(function (a, b) {
      return b.value - a.value;
    });
    return {
      metric: metric,
      leader: sorted[0] || null,
      entries: sorted
    };
  });

  return {
    metrics: metrics,
    leaders: leaders,
    rows: rows
  };
}

function saveCompetitionEntry(payload) {
  authorizeOrThrow_();
  if (!payload || !payload.metric || !payload.csrName) {
    throw new Error('Invalid payload. Expected metric and csrName.');
  }

  const db = openDatabase_();
  const sheet = db.getSheetByName(CONFIG.SHEETS.CSR_COMPETITIONS);
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  const now = new Date().toISOString();
  let foundRow = -1;

  for (var i = 1; i < values.length; i++) {
    if (
      values[i][0] === payload.weekStart &&
      values[i][1] === payload.store &&
      values[i][2] === payload.metric &&
      values[i][3] === payload.csrName
    ) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow > -1) {
    sheet.getRange(foundRow, headers.indexOf('Value') + 1).setValue(Number(payload.value || 0));
    sheet.getRange(foundRow, headers.indexOf('UpdatedAt') + 1).setValue(now);
  } else {
    sheet.appendRow([
      payload.weekStart || getWeekStartISO_(new Date()),
      payload.store || CONFIG.DEFAULT_STORES[0],
      payload.metric,
      payload.csrName,
      Number(payload.value || 0),
      now
    ]);
  }

  clearDashboardCache_();
  return { success: true };
}

function getChecklist(dateIso, store) {
  authorizeOrThrow_();
  const db = openDatabase_();
  const sheet = db.getSheetByName(CONFIG.SHEETS.CLEANING_CHECKLIST);
  const rows = getDataRows_(sheet);
  const targetDate = dateIso || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const targetStore = store || CONFIG.DEFAULT_STORES[0];

  const filtered = rows.filter(function (row) {
    return row.date === targetDate && row.store === targetStore;
  });

  if (!filtered.length) {
    seedChecklistForDateStore_(sheet, targetDate, targetStore);
    return getChecklistInternal_(sheet, targetDate, targetStore);
  }

  return getChecklistInternal_(sheet, targetDate, targetStore);
}

function getChecklistInternal_(sheet, targetDate, targetStore) {
  const rows = getDataRows_(sheet);
  const filtered = rows.filter(function (row) {
    return row.date === targetDate && row.store === targetStore;
  });

  return {
    date: targetDate,
    store: targetStore,
    items: filtered.map(function (row) {
      return {
        task: row.task,
        completed: normalizeBoolean_(row.completed),
        completedBy: row.completedBy,
        completedAt: row.completedAt
      };
    })
  };
}

function setChecklistItem(payload) {
  authorizeOrThrow_();
  if (!payload || !payload.date || !payload.store || !payload.task) {
    throw new Error('Invalid payload for checklist update.');
  }

  const db = openDatabase_();
  const sheet = db.getSheetByName(CONFIG.SHEETS.CLEANING_CHECKLIST);
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const email = getCurrentUserEmail_();
  const completed = !!payload.completed;
  const completedAt = completed ? new Date().toISOString() : '';
  let updated = false;

  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === payload.date && values[i][1] === payload.store && values[i][2] === payload.task) {
      sheet.getRange(i + 1, headers.indexOf('Completed') + 1).setValue(completed);
      sheet.getRange(i + 1, headers.indexOf('CompletedBy') + 1).setValue(completed ? email : '');
      sheet.getRange(i + 1, headers.indexOf('CompletedAt') + 1).setValue(completedAt);
      updated = true;
      break;
    }
  }

  if (!updated) {
    sheet.appendRow([
      payload.date,
      payload.store,
      payload.task,
      completed,
      completed ? email : '',
      completedAt
    ]);
  }

  clearDashboardCache_();
  return { success: true };
}

function ensureDatabase() {
  const db = openDatabase_();
  ensureSheet_(db, CONFIG.SHEETS.SALES_WEEKLY, ['WeekStart', 'Store', 'Year', 'Sales']);
  ensureSheet_(db, CONFIG.SHEETS.CSR_PERFORMANCE, ['Date', 'Store', 'CSRName', 'Sales', 'Hours']);
  ensureSheet_(db, CONFIG.SHEETS.CSR_SCHEDULE, ['Date', 'Store', 'CSRName', 'ShiftStatus']);
  ensureSheet_(db, CONFIG.SHEETS.CSR_RECOGNITION, ['WeekStart', 'CSRName', 'Store', 'Quote', 'ImageUrl']);
  ensureSheet_(db, CONFIG.SHEETS.CSR_COMPETITIONS, ['WeekStart', 'Store', 'Metric', 'CSRName', 'Value', 'UpdatedAt']);
  ensureSheet_(db, CONFIG.SHEETS.CLEANING_CHECKLIST, ['Date', 'Store', 'Task', 'Completed', 'CompletedBy', 'CompletedAt']);
  ensureSheet_(db, CONFIG.SHEETS.AUTH_LOG, ['Timestamp', 'Email', 'Authorized']);

  seedIfEmpty_(db);
  return {
    spreadsheetId: db.getId(),
    spreadsheetUrl: db.getUrl()
  };
}

function logAuthAttempt(email, authorized) {
  try {
    const db = openDatabase_();
    const sheet = db.getSheetByName(CONFIG.SHEETS.AUTH_LOG) || ensureSheet_(db, CONFIG.SHEETS.AUTH_LOG, ['Timestamp', 'Email', 'Authorized']);
    sheet.appendRow([new Date().toISOString(), email || 'unknown', !!authorized]);
    return { success: true };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

function openDatabase_() {
  const props = PropertiesService.getScriptProperties();
  const existingId = props.getProperty(CONFIG.DB_PROPERTY_KEY);
  if (existingId) {
    return SpreadsheetApp.openById(existingId);
  }

  const ss = SpreadsheetApp.create('PLT CSR Dashboard DB');
  props.setProperty(CONFIG.DB_PROPERTY_KEY, ss.getId());
  return ss;
}

function ensureSheet_(db, name, headers) {
  let sheet = db.getSheetByName(name);
  if (!sheet) {
    sheet = db.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    ensureHeaders_(sheet, headers);
  }
  return sheet;
}

function ensureHeaders_(sheet, expectedHeaders) {
  const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (header) {
    return String(header || '').trim();
  });
  expectedHeaders.forEach(function (header) {
    if (currentHeaders.indexOf(header) === -1) {
      sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
      currentHeaders.push(header);
    }
  });
}

function seedIfEmpty_(db) {
  seedSalesWeekly_(db.getSheetByName(CONFIG.SHEETS.SALES_WEEKLY));
  seedPerformance_(db.getSheetByName(CONFIG.SHEETS.CSR_PERFORMANCE));
  seedSchedule_(db.getSheetByName(CONFIG.SHEETS.CSR_SCHEDULE));
  seedRecognition_(db.getSheetByName(CONFIG.SHEETS.CSR_RECOGNITION));
  seedCompetitions_(db.getSheetByName(CONFIG.SHEETS.CSR_COMPETITIONS));
  seedChecklist_(db.getSheetByName(CONFIG.SHEETS.CLEANING_CHECKLIST));
}

function seedSalesWeekly_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const weekStart = getWeekStartISO_(new Date());
  const rows = [];
  CONFIG.DEFAULT_STORES.forEach(function (store, idx) {
    rows.push([weekStart, store, 2025, 24000 + idx * 2200]);
    rows.push([weekStart, store, 2026, 26500 + idx * 2600]);
  });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedPerformance_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yIso = Utilities.formatDate(yesterday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const rows = [
    [yIso, 'Dublin', 'Ana', 3900, 22],
    [yIso, 'Pleasanton', 'Chris', 4600, 24],
    [yIso, 'San Ramon', 'Taylor', 3600, 20]
  ];
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedSchedule_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const monday = getWeekStartDate_(new Date());
  const csrs = ['Ana', 'Chris', 'Taylor', 'Jordan'];
  const rows = [];
  for (var d = 0; d < 7; d++) {
    const date = new Date(monday);
    date.setDate(monday.getDate() + d);
    const iso = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    csrs.forEach(function (csr, idx) {
      const store = CONFIG.DEFAULT_STORES[idx % CONFIG.DEFAULT_STORES.length];
      const isOff = (d + idx) % 5 === 0;
      rows.push([iso, store, csr, isOff ? 'OFF' : '9a-5p']);
    });
  }
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedRecognition_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const weekStart = getWeekStartISO_(new Date());
  sheet.getRange(2, 1, 1, 5).setValues([[
    weekStart,
    'Chris',
    'Pleasanton',
    'Every customer leaves smiling because we listen first.',
    ''
  ]]);
}

function seedCompetitions_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const weekStart = getWeekStartISO_(new Date());
  const metrics = ['Conversions', 'Patio Signups', 'Alterations'];
  const csrs = ['Ana', 'Chris', 'Taylor'];
  const rows = [];
  metrics.forEach(function (metric, mIdx) {
    csrs.forEach(function (csr, cIdx) {
      rows.push([
        weekStart,
        CONFIG.DEFAULT_STORES[cIdx % CONFIG.DEFAULT_STORES.length],
        metric,
        csr,
        (mIdx + 1) * 4 + cIdx,
        new Date().toISOString()
      ]);
    });
  });
  sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
}

function seedChecklist_(sheet) {
  if (sheet.getLastRow() > 1) return;
  const dateIso = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  seedChecklistForDateStore_(sheet, dateIso, CONFIG.DEFAULT_STORES[0]);
}

function seedChecklistForDateStore_(sheet, dateIso, store) {
  const tasks = [
    'Vacuum lobby and rugs',
    'Sanitize counters and POS terminals',
    'Take out trash and recycling',
    'Lock chemical storage',
    'Final lights / alarm walkthrough'
  ];
  const rows = tasks.map(function (task) {
    return [dateIso, store, task, false, '', ''];
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function getDataRows_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map(function (row) {
    const obj = {};
    headers.forEach(function (header, idx) {
      const key = toCamelCase_(header);
      obj[key] = row[idx];
    });
    return obj;
  });
}

function buildSalesComparison_(db) {
  const rows = getDataRows_(db.getSheetByName(CONFIG.SHEETS.SALES_WEEKLY));
  const targetWeek = getWeekStartISO_(new Date());
  const availableYears = {};
  const stores = {};

  rows.forEach(function (row) {
    if (row.weekStart !== targetWeek) return;
    const year = Number(row.year);
    const store = row.store;
    if (!store || !year) return;
    availableYears[year] = true;
    stores[store] = true;
  });

  const yearList = Object.keys(availableYears)
    .map(function (year) { return Number(year); })
    .sort(function (a, b) { return a - b; });

  const thisYear = yearList.length ? yearList[yearList.length - 1] : new Date().getFullYear();
  const lastYear = yearList.length > 1 ? yearList[yearList.length - 2] : thisYear - 1;
  const grouped = {};

  rows.forEach(function (row) {
    if (row.weekStart !== targetWeek) return;
    const store = row.store;
    if (!store) return;
    if (!grouped[store]) grouped[store] = { store: store, salesThisYear: 0, salesLastYear: 0 };
    const year = Number(row.year);
    if (year === lastYear) grouped[store].salesLastYear += Number(row.sales || 0);
    if (year === thisYear) grouped[store].salesThisYear += Number(row.sales || 0);
  });

  const allStores = Object.keys(stores).length ? Object.keys(stores) : CONFIG.DEFAULT_STORES;
  const mappedRows = allStores.map(function (store) {
    if (!grouped[store]) {
      grouped[store] = { store: store, salesThisYear: 0, salesLastYear: 0 };
    }
    const item = grouped[store];
    const delta = item.salesThisYear - item.salesLastYear;
    const pct = item.salesLastYear ? delta / item.salesLastYear : 0;
    return {
      store: store,
      salesLastYear: item.salesLastYear,
      salesThisYear: item.salesThisYear,
      delta: delta,
      deltaPct: pct
    };
  });

  return {
    years: {
      thisYear: thisYear,
      lastYear: lastYear
    },
    rows: mappedRows
  };
}

function buildCsrPerformanceYesterday_(db) {
  const rows = getDataRows_(db.getSheetByName(CONFIG.SHEETS.CSR_PERFORMANCE));
  const y = new Date();
  y.setDate(y.getDate() - 1);
  const dateIso = Utilities.formatDate(y, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const mapped = rows
    .filter(function (row) {
      return row.date === dateIso;
    })
    .map(function (row) {
      const sales = Number(row.sales || 0);
      const hours = Number(row.hours || 0);
      return {
        date: row.date,
        store: row.store,
        csrName: row.csrName,
        sales: sales,
        hours: hours,
        dollarsPerHour: hours ? sales / hours : 0
      };
    })
    .sort(function (a, b) {
      return b.dollarsPerHour - a.dollarsPerHour;
    });

  return {
    date: dateIso,
    goal: CONFIG.PERFORMANCE_GOAL,
    topPerformer: mapped[0] || null,
    entries: mapped
  };
}

function buildScheduleWeek_(db) {
  const rows = getDataRows_(db.getSheetByName(CONFIG.SHEETS.CSR_SCHEDULE));
  const monday = getWeekStartDate_(new Date());
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  const startIso = Utilities.formatDate(monday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const endIso = Utilities.formatDate(sunday, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const entries = rows.filter(function (row) {
    return row.date >= startIso && row.date <= endIso;
  });

  const offList = entries.filter(function (row) {
    return String(row.shiftStatus).toUpperCase() === 'OFF';
  });

  return {
    weekStart: startIso,
    weekEnd: endIso,
    entries: entries,
    offList: offList
  };
}

function buildCsrOfWeek_(db) {
  const rows = getDataRows_(db.getSheetByName(CONFIG.SHEETS.CSR_RECOGNITION));
  const weekStart = getWeekStartISO_(new Date());
  const found = rows.find(function (row) {
    return row.weekStart === weekStart;
  });

  return found || {
    weekStart: weekStart,
    csrName: 'TBD',
    store: '-',
    quote: 'Recognizing excellence in service every week.',
    imageUrl: ''
  };
}

function buildCompetitionsSnapshot_(db) {
  const rows = getDataRows_(db.getSheetByName(CONFIG.SHEETS.CSR_COMPETITIONS));
  const weekStart = getWeekStartISO_(new Date());
  const filtered = rows.filter(function (row) {
    return row.weekStart === weekStart;
  });

  const byMetric = {};
  filtered.forEach(function (row) {
    if (!byMetric[row.metric]) byMetric[row.metric] = [];
    byMetric[row.metric].push(row);
  });

  return Object.keys(byMetric).map(function (metric) {
    const entries = byMetric[metric].map(function (row) {
      return {
        csrName: row.csrName,
        store: row.store,
        value: Number(row.value || 0)
      };
    }).sort(function (a, b) { return b.value - a.value; });
    return {
      metric: metric,
      leader: entries[0] || null,
      entries: entries
    };
  });
}

function buildChecklistSnapshot_(db) {
  const sheet = db.getSheetByName(CONFIG.SHEETS.CLEANING_CHECKLIST);
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const store = CONFIG.DEFAULT_STORES[0];
  const data = getChecklist(today, store);
  const total = data.items.length;
  const completed = data.items.filter(function (item) { return item.completed; }).length;

  return {
    date: today,
    store: store,
    total: total,
    completed: completed,
    progress: total ? completed / total : 0
  };
}


function upsertSheetRow_(sheet, record, keys) {
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const keyIndexes = keys.map(function (key) {
    return headers.indexOf(key);
  });

  if (keyIndexes.some(function (idx) { return idx === -1; })) {
    throw new Error('Sheet headers are missing required columns: ' + keys.join(', '));
  }

  let rowIndex = -1;
  for (var r = 1; r < values.length; r++) {
    const match = keys.every(function (key, idx) {
      return values[r][keyIndexes[idx]] === record[key];
    });
    if (match) {
      rowIndex = r + 1;
      break;
    }
  }

  const rowValues = headers.map(function (header, idx) {
    const hasValue = Object.prototype.hasOwnProperty.call(record, header);
    if (!hasValue) {
      if (rowIndex > -1) {
        return values[rowIndex - 1][idx];
      }
      return '';
    }
    return record[header];
  });

  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
}

function authorizeOrThrow_() {
  const email = getCurrentUserEmail_();
  const authorized = isAuthorizedEmail_(email);
  if (!authorized) {
    logAuthAttempt(email || 'unknown', false);
    throw new Error('Unauthorized access.');
  }
}

function isAuthorizedEmail_(email) {
  return CONFIG.ALLOWED_EMAILS.indexOf(String(email || '').toLowerCase()) !== -1;
}

function getCurrentUserEmail_() {
  const active = Session.getActiveUser().getEmail();
  const effective = Session.getEffectiveUser().getEmail();
  return (active || effective || '').toLowerCase();
}

function getWeekStartDate_(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getWeekStartISO_(date) {
  return Utilities.formatDate(getWeekStartDate_(date), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function toCamelCase_(text) {
  return String(text)
    .replace(/[^a-zA-Z0-9 ]/g, '')
    .split(' ')
    .map(function (chunk, idx) {
      if (!chunk) return '';
      const lower = chunk.toLowerCase();
      return idx === 0 ? lower : lower.charAt(0).toUpperCase() + lower.slice(1);
    })
    .join('');
}

function normalizeBoolean_(value) {
  if (value === true || value === false) return value;
  const normalized = String(value).toLowerCase();
  return normalized === 'true' || normalized === '1' || normalized === 'yes';
}

function clearDashboardCache_() {
  CacheService.getScriptCache().remove(CONFIG.CACHE_KEY_DASHBOARD);
}
