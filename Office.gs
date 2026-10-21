/**
 * Grace Baptist Church - Usher Attendance: the office page
 *
 * Back end for office/index.html. Through the week, the office fills in the
 * numbers the ushers don't count (Sunday School, the ministries and the
 * children's classes: the other columns of the History tabs) and checks or
 * corrects what the ushers logged. A week is a Sunday and the Wednesday after
 * it, as in the History tabs. Its office numbers are one row of the Weekly
 * tab, under the same column names as the History tabs.
 *
 * Saving from the office page sends no email. Weekly.gs emails the week's
 * report on Saturday morning.
 *
 * The page sends the office code (OFFICE_CODE, made by setup) with every
 * request. It's separate from the usher code because it can also read and
 * change what the ushers logged.
 */

const WEEKLY_NAME = 'Weekly';
const WEEKLY_STATUS = 'Weekly Report';
const OFFICE_EDIT = 'Office Edit';   // Attendance tab column: who entered or changed the row on the office page
const DAY_MS = 86400000;

const NURSERIES = ['Infant Nursery (0-1 yrs)', 'Baby Nursery (1-2 yrs)', 'Toddler Nursery (2-3 yrs)'];

// What the office page asks for, card by card. Each key is a History tab
// column. A card with a service starts its "from" boxes at the ushers' count
// for that service (those areas added up) until the office saves a number.
const OFFICE_GROUPS = [
  {
    title: 'Sunday School',
    total: 'SS Total',
    fields: [
      { key: 'SS Adult', label: 'Adults' },
      { key: 'SS Teen', label: 'Teens' },
      { key: 'SS Children', label: 'Children' },
    ],
  },
  {
    title: 'Sunday School Classes',
    fields: [
      { key: 'SS Nursery', label: 'Nursery' },
      { key: 'SS K4-K5', label: 'K4-K5' },
      { key: 'SS 1st-2nd', label: '1st-2nd' },
      { key: 'SS 3rd-4th', label: '3rd-4th' },
      { key: 'SS 5th-6th', label: '5th-6th' },
    ],
  },
  {
    title: 'Ministries',
    fields: [
      { key: 'Bus', label: 'Bus' },
      { key: 'Keenagers/HGM', label: 'Keenagers/HGM' },
      { key: 'K4T/AWANA', label: 'K4T/AWANA' },
      { key: 'Kidstown', label: 'Kidstown' },
    ],
  },
  {
    title: 'Sunday Morning Children',
    service: 'Sunday AM',
    fields: [
      { key: 'AM Nursery', label: 'Nursery', from: NURSERIES },
      { key: 'AM K4-K5', label: 'K4-K5', from: ['Pre-Primary (4-5 yrs)'] },
      { key: 'AM 1st-3rd', label: '1st-3rd', from: ['Junior Church (1st-3rd)'] },
    ],
  },
  {
    title: 'Sunday Evening Children',
    service: 'Sunday PM',
    fields: [
      { key: 'PM Nursery', label: 'Nursery', from: NURSERIES },
    ],
  },
  {
    title: 'Wednesday Children',
    service: 'Wednesday',
    fields: [
      { key: 'Wed Nursery', label: 'Nursery', from: NURSERIES },
      { key: 'Wed K4-K5', label: 'K4-K5', from: ['Pre-Primary (4-5 yrs)'] },
      { key: 'Wed 1st-3rd', label: '1st-3rd', from: ['Junior Church (1st-3rd)'] },
      { key: 'Wed 4th-6th', label: '4th-6th' },
      { key: 'Wed Teens', label: 'Teens' },
    ],
  },
];
// The old sheet's Children's Ministries tab repeated SS Teen as SS Teens, so
// the page asks once and fills both.
const OFFICE_COPIES = { 'SS Teens': 'SS Teen' };
// The services every week has, and how many days after its Sunday each falls.
const WEEK_SERVICES = [
  { service: 'Sunday AM', days: 0 },
  { service: 'Sunday PM', days: 0 },
  { service: 'Wednesday', days: 3 },
];

/** POST from the office page: {code, action: 'week' | 'save', ...}. */
function officeRequest_(body) {
  if (!codeIsRight_(body.code, 'OFFICE_CODE')) return { status: 'badcode' };
  if (body.action === 'week') return officeWeek_(parseSunday_(body.week));
  if (body.action === 'save') return saveOfficeWeek_(body.form || {});
  throw new Error('The page asked for something the script doesn\'t know. Reload the page.');
}

/** "2026-10-18" -> that Sunday at noon. */
function parseSunday_(value) {
  const d = parseDate_(value);
  if (d.getDay() !== 0) throw new Error('A week starts on a Sunday.');
  return d;
}

/** The Weekly tab's number formats (a function, so Code.gs's formats exist by the time it runs). */
function weeklyFormats_() {
  return { 'Week Of': DATE_FORMAT, 'Updated At': SUBMITTED_FORMAT };
}

function ymd_(d, tz) {
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

/** Everything the office page shows for the week starting on sunday. */
function officeWeek_(sunday) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const sheet = getSheet_();
  const office = officeValues_(sunday, tz);
  const services = weekServices_(sheet, headerIndex_(sheet), sunday, tz);
  // For weeks before the ushers' page, the old sheet's count of each service.
  services.forEach(function (s) { if (!s.logged) s.old = historyTotal_(s.service, sunday, tz); });
  return {
    status: 'ok',
    week: ymd_(sunday, tz),
    sections: SECTIONS,
    groups: OFFICE_GROUPS,
    services: services,
    office: office.values,
    notes: office.notes,
    updatedBy: office.by,
    updatedAt: office.at instanceof Date ? Utilities.formatDate(office.at, tz, 'EEE M/d h:mm a') : '',
    from: office.from,
  };
}

/**
 * The week's services from the Attendance tab, oldest first, plus a blank
 * Sunday AM, Sunday PM or Wednesday for any the ushers didn't log.
 */
function weekServices_(sheet, cols, sunday, tz) {
  const start = ymd_(sunday, tz);
  const end = ymd_(new Date(sunday.getTime() + 6 * DAY_MS), tz);
  const last = sheet.getLastRow();
  const rows = last > 1 ? sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues() : [];
  const get = function (r, h) { return cols[h] ? r[cols[h] - 1] : ''; };
  const out = [];
  rows.forEach(function (r) {
    const d = get(r, 'Date');
    if (!(d instanceof Date) || ymd_(d, tz) < start || ymd_(d, tz) > end) return;
    const counts = {};
    SECTIONS.forEach(function (s) {
      s.areas.forEach(function (a) { counts[a] = typeof get(r, a) === 'number' ? get(r, a) : ''; });
    });
    const at = get(r, 'Submitted At');
    out.push({
      date: ymd_(d, tz),
      service: String(get(r, 'Service')),
      otherName: String(get(r, 'Other Service')),
      logged: true,
      counts: counts,
      total: Number(get(r, 'Total Attendance')) || 0,
      by: String(get(r, 'Submitted By')),
      at: at instanceof Date ? Utilities.formatDate(at, tz, 'EEE M/d h:mm a') : '',
      edited: String(get(r, OFFICE_EDIT)),
      notes: String(get(r, 'Notes')),
    });
  });
  WEEK_SERVICES.forEach(function (w) {
    if (out.some(function (e) { return e.service === w.service; })) return;
    out.push({
      date: ymd_(new Date(sunday.getTime() + w.days * DAY_MS), tz), service: w.service, otherName: '',
      logged: false, counts: {}, total: 0, by: '', at: '', edited: '', notes: '',
    });
  });
  return out.sort(function (a, b) {
    return a.date < b.date ? -1 : a.date > b.date ? 1 : SERVICES.indexOf(a.service) - SERVICES.indexOf(b.service);
  });
}

// ---- The Weekly tab ---------------------------------------------------------

/** Every History column the office fills: the fields, the totals and the copies. */
function officeColumns_() {
  const out = [];
  OFFICE_GROUPS.forEach(function (g) {
    if (g.total) out.push(g.total);
    g.fields.forEach(function (f) { out.push(f.key); });
  });
  return out.concat(Object.keys(OFFICE_COPIES));
}

/** The Weekly tab's columns: the office's History columns, in the History tabs' order. */
function weeklyHeaders_() {
  const office = officeColumns_();
  return ['Week Of'].concat(HISTORY_HEADERS.concat(CHILDREN_HEADERS).filter(function (h) {
    return office.indexOf(h) > -1 || h === 'Notes';
  }), ['Updated By', 'Updated At', WEEKLY_STATUS]);
}

/** Returns the Weekly tab, creating it and adding any missing columns on the right. */
function weeklyTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tab = ss.getSheetByName(WEEKLY_NAME);
  if (!tab) tab = ss.insertSheet(WEEKLY_NAME, Math.min(2, ss.getNumSheets()));
  const lastCol = tab.getLastColumn();
  const current = lastCol ? tab.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  const missing = weeklyHeaders_().filter(function (h) { return current.indexOf(h) === -1; });
  if (missing.length) {
    const all = current.concat(missing);
    if (tab.getMaxColumns() < all.length) tab.insertColumnsAfter(tab.getMaxColumns(), all.length - tab.getMaxColumns());
    tab.getRange(1, current.length + 1, 1, missing.length).setValues([missing]);
    tab.getRange(1, 1, 1, all.length).setFontWeight('bold').setFontColor('#ffffff').setBackground('#233617')
      .setWrap(true).setVerticalAlignment('middle');
    tab.setFrozenRows(1);
    tab.setFrozenColumns(1);
    const formats = weeklyFormats_();
    all.forEach(function (h, i) {
      if (formats[h]) tab.getRange(2, i + 1, tab.getMaxRows() - 1, 1).setNumberFormat(formats[h]);
    });
  }
  return tab;
}

/** A tab's header row as {header: column}. */
function tabColumns_(tab) {
  const index = {};
  tab.getRange(1, 1, 1, tab.getLastColumn()).getValues()[0].forEach(function (h, i) {
    if (h) index[String(h)] = i + 1;
  });
  return index;
}

/** The Weekly tab row for the week starting on sunday, or 0. */
function weeklyRow_(tab, cols, sunday, tz) {
  const last = tab.getLastRow();
  if (last < 2) return 0;
  const day = ymd_(sunday, tz);
  const weeks = tab.getRange(2, cols['Week Of'], last - 1, 1).getValues();
  for (let i = 0; i < weeks.length; i++) {
    if (weeks[i][0] instanceof Date && ymd_(weeks[i][0], tz) === day) return i + 2;
  }
  return 0;
}

/**
 * A week's office numbers: {values: {column: value}, notes, by, at, from}.
 * From the Weekly tab when the office has saved that week ("weekly"), or
 * else from that year's History tab ("history"), or blank ("").
 */
function officeValues_(sunday, tz) {
  const keys = officeColumns_();
  const tab = weeklyTab_();
  const cols = tabColumns_(tab);
  const row = weeklyRow_(tab, cols, sunday, tz);
  const v = row ? tab.getRange(row, 1, 1, tab.getLastColumn()).getValues()[0] : null;
  // A row with no Updated By only records a weekly report sent before the office saved anything.
  if (v && v[cols['Updated By'] - 1]) {
    const values = {};
    keys.forEach(function (k) { if (cols[k]) values[k] = v[cols[k] - 1]; });
    return {
      values: values, notes: String(v[cols['Notes'] - 1]), by: String(v[cols['Updated By'] - 1]),
      at: v[cols['Updated At'] - 1], from: 'weekly',
    };
  }
  const old = historyWeek_(sunday, tz, keys);
  return old ? { values: old.values, notes: old.notes, by: '', at: '', from: 'history' }
    : { values: {}, notes: '', by: '', at: '', from: '' };
}

/**
 * The office columns from the History tab rows dated in the week starting on
 * sunday. The old sheet sometimes split a week over two rows (Easter), so
 * each column takes the first row that has something in it.
 */
function historyWeek_(sunday, tz, keys) {
  const tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HISTORY_PREFIX + Utilities.formatDate(sunday, tz, 'yyyy'));
  if (!tab || tab.getLastRow() < 2) return null;
  const values = tab.getRange(1, 1, tab.getLastRow(), tab.getLastColumn()).getValues();
  const head = values[0].map(String);
  const start = ymd_(sunday, tz);
  const end = ymd_(new Date(sunday.getTime() + 6 * DAY_MS), tz);
  const rows = values.slice(1).filter(function (r) {
    return r[0] instanceof Date && ymd_(r[0], tz) >= start && ymd_(r[0], tz) <= end;
  });
  if (!rows.length) return null;
  const out = {};
  keys.forEach(function (k) {
    const c = head.indexOf(k);
    if (c < 0) return;
    rows.some(function (r) {
      if (r[c] === '') return false;
      out[k] = r[c];
      return true;
    });
  });
  const notes = head.indexOf('Notes');
  return {
    values: out,
    notes: rows.map(function (r) { return notes < 0 ? '' : String(r[notes]); }).filter(String).join('; '),
  };
}

// ---- Saving -------------------------------------------------------------------

/**
 * Saves the office page: its numbers to the week's Weekly row, and any usher
 * counts it changed to the Attendance tab. A service the ushers didn't log
 * gets a new row. Nothing is emailed. Returns the week as saved.
 */
function saveOfficeWeek_(form) {
  const sunday = parseSunday_(form.week);
  const by = clean_(form.by, 60);
  if (!by) throw new Error('Add your name before saving.');

  // Check every number before writing any of them.
  const given = form.office || {};
  const office = {};
  OFFICE_GROUPS.forEach(function (g) {
    g.fields.forEach(function (f) { office[f.key] = count_(given[f.key], f.label + ' (' + g.title + ')'); });
  });
  const services = (form.services || []).map(function (s) {
    const service = String(s.service || '');
    if (SERVICES.indexOf(service) === -1) throw new Error('Unknown service: ' + service);
    const counts = s.counts || {};
    const record = {};
    fillCounts_(record, counts);
    return {
      date: parseDate_(s.date),
      service: service,
      otherName: service === 'Other' ? clean_(s.otherName, 60) : '',
      record: record,
      any: Object.keys(counts).some(function (a) { return counts[a] !== '' && counts[a] !== null; }),
    };
  });

  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const stamp = by + ', ' + Utilities.formatDate(new Date(), tz, 'M/d h:mm a');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    const cols = headerIndex_(sheet);
    services.forEach(function (s) {
      const row = findRow_(sheet, cols, s.date, s.service, s.otherName, tz);
      if (row) {
        s.record[OFFICE_EDIT] = 'Changed by ' + stamp;
        writeRow_(sheet, cols, row, s.record, ATTENDANCE_FORMATS);
      } else if (s.any) {
        s.record['Date'] = s.date;
        s.record['Service'] = s.service;
        s.record['Other Service'] = s.otherName;
        s.record['Submitted By'] = by;
        s.record['Submitted At'] = new Date();
        s.record[REPORT_STATUS] = 'Not emailed: entered on the office page';
        s.record[OFFICE_EDIT] = 'Entered by ' + stamp;
        writeRow_(sheet, cols, 0, s.record, ATTENDANCE_FORMATS);
      }
    });

    const record = { 'Week Of': sunday, 'Notes': clean_(form.notes, 500), 'Updated By': by, 'Updated At': new Date() };
    Object.keys(office).forEach(function (k) { record[k] = office[k]; });
    OFFICE_GROUPS.forEach(function (g) {
      if (!g.total) return;
      const parts = g.fields.map(function (f) { return office[f.key]; });
      record[g.total] = parts.some(function (n) { return n !== ''; })
        ? parts.reduce(function (t, n) { return t + (n || 0); }, 0) : '';
    });
    Object.keys(OFFICE_COPIES).forEach(function (k) { record[k] = office[OFFICE_COPIES[k]]; });
    const tab = weeklyTab_();
    const wcols = tabColumns_(tab);
    const row = weeklyRow_(tab, wcols, sunday, tz);
    writeRow_(tab, wcols, row, record, weeklyFormats_());
    if (!row && tab.getLastRow() > 2) {
      // Keep the weeks in order when an earlier one is filled in late.
      tab.getRange(2, 1, tab.getLastRow() - 1, tab.getLastColumn()).sort(wcols['Week Of']);
    }
  } finally {
    lock.releaseLock();
  }
  const out = officeWeek_(sunday);
  out.status = 'saved';
  return out;
}
