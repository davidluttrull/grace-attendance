/**
 * Grace Baptist Church - Usher Attendance
 *
 * Back end for the ushers' attendance page (index.html, hosted on GitHub
 * Pages). The page reads the room list from doGet and sends each service's
 * numbers to doPost, which logs them as one row in this spreadsheet.
 * See SETUP.md for how to deploy it.
 *
 * Report.gs emails a report of each submission, as the body of the email, to
 * the addresses on the Report Emails tab. Office.gs is the back end for the
 * office page (office/index.html), and Weekly.gs emails its weekly report.
 *
 * To add, rename or reorder an area, edit SECTIONS (or a service's list in
 * SERVICE_SECTIONS) below and deploy a new version. New areas get a new column at the right end of the sheet; old
 * columns are never deleted, so past data stays put.
 */

const SHEET_NAME = 'Attendance';
const SUMMARY_NAME = 'Summary';

const SERVICES = ['Sunday AM', 'Sunday PM', 'Wednesday', 'Other'];

// Areas counted at each service, in the order ushers see them. A service in
// SERVICE_SECTIONS below counts its own list instead. Keep the same three
// section titles in every list: each section's subtotal column is named after
// its title.
const SECTIONS = [
  {
    title: 'Auditorium',
    areas: ['South Wing', 'South Main', 'Platform', 'North Wing', 'North Main'],
  },
  {
    title: 'Children',
    areas: ['Infant Nursery (0-1 yrs)', 'Baby Nursery (1-2 yrs)', 'Toddler Nursery (2-3 yrs)',
            'Pre-Primary (4-5 yrs)', 'Junior Church (1st-3rd)'],
  },
  {
    title: 'Other',
    areas: ['Ushers', 'Spanish Ministry', 'Little Chapel', 'Fellowship Hall',
            'Kitchen', 'Late Comers'],
  },
];

// Services whose rooms differ from SECTIONS.
const SERVICE_SECTIONS = {
  // Wednesday: Baby and Toddler share the Main Nursery, 4th-6th grade and the
  // hallway classroom meet, and Fellowship Hall and the Kitchen aren't used.
  'Wednesday': [
    {
      title: 'Auditorium',
      areas: ['South Wing', 'South Main', 'Platform', 'North Wing', 'North Main'],
    },
    {
      title: 'Children',
      areas: ['Infant Nursery (0-1 yrs)', 'Main Nursery (1-3 yrs)', 'Pre-Primary (4-5 yrs)',
              'Junior Church (1st-3rd)', '4th-6th Grade'],
    },
    {
      title: 'Other',
      areas: ['Ushers', 'Spanish Ministry', 'Little Chapel', 'Hallway Classroom', 'Late Comers'],
    },
  ],
};

// Usher door assignments from the bottom of the paper form.
const DOORS = [
  { group: 'South Wing', doors: [1, 2] },
  { group: 'South Main', doors: [3, 4] },
  { group: 'North Main', doors: [5, 6] },
  { group: 'North Wing', doors: [7, 8] },
];

// How the Date and Submitted At columns show in the sheet.
const DATE_FORMAT = 'ddd m/d/yyyy';
const SUBMITTED_FORMAT = 'm/d/yyyy h:mm am/pm';
const ATTENDANCE_FORMATS = { 'Date': DATE_FORMAT, 'Submitted At': SUBMITTED_FORMAT };

const MAX_COUNT = 5000;
const MAX_BAD_CODES = 20; // wrong usher codes allowed per 10 minutes, across all phones

/** GET: the room list the page builds its form from. */
function doGet() {
  return json_(pageConfig_());
}

/**
 * POST from the page, body {form, replace}. The page sends it as plain text
 * with no Google login attached, so it works whatever account a phone uses.
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.action) return json_(officeRequest_(body));   // the office page (Office.gs)
    if (!codeIsRight_(body.code, 'USHER_CODE')) return json_({ status: 'badcode' });
    return json_(submitAttendance(body.form, body.replace));
  } catch (err) {
    return json_({ status: 'error', message: err.message });
  }
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Run this once from the editor to create the tabs, grant permissions and
 * schedule the weekly report. It also creates the usher code and the office
 * code (the first time) and shows them in the Execution log. To change a
 * code, edit USHER_CODE or OFFICE_CODE under Project Settings > Script
 * Properties.
 */
function setup() {
  const sheet = getSheet_();
  ensureEmailsTab_();
  weeklyTab_();
  ensureServiceViews_(SpreadsheetApp.getActiveSpreadsheet(), headerIndex_(sheet));
  ensureWeeklyTrigger_();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);

  const props = PropertiesService.getScriptProperties();
  ['USHER_CODE', 'OFFICE_CODE'].forEach(function (name) {
    let code = props.getProperty(name);
    if (!code) {
      code = String(Math.floor(100000 + Math.random() * 900000));
      props.setProperty(name, code);
    }
    Logger.log((name === 'USHER_CODE' ? 'Usher' : 'Office') + ' code: ' + code);
  });
}

/**
 * Checks a code against a Script Property (USHER_CODE or OFFICE_CODE),
 * refusing everything for 10 minutes after too many wrong ones. The codes
 * aren't in this file, so the public copy on GitHub doesn't give them away.
 */
function codeIsRight_(code, property) {
  const expected = PropertiesService.getScriptProperties().getProperty(property);
  if (!expected) throw new Error('The code hasn\'t been set up yet. Run setup in Apps Script.');
  const cache = CacheService.getScriptCache();
  const key = 'bad' + property;
  const misses = Number(cache.get(key) || 0);
  if (misses >= MAX_BAD_CODES) throw new Error('Too many wrong codes. Wait 10 minutes and try again.');
  if (String(code || '').trim() === expected.trim()) return true;
  cache.put(key, String(misses + 1), 600);
  return false;
}

/**
 * Saves one service's numbers as a row.
 * If that date + service is already logged, returns {status: 'duplicate'}
 * unless replace is true, in which case the old row is overwritten.
 */
function submitAttendance(form, replace) {
  form = form || {};
  const date = parseDate_(form.date);
  const service = String(form.service || '');
  if (SERVICES.indexOf(service) === -1) throw new Error('Please pick a service.');
  const otherName = service === 'Other' ? clean_(form.otherName, 60) : '';
  if (service === 'Other' && !otherName) throw new Error('Please name the service.');

  const counts = form.counts || {};
  // A phone still showing an old room list would send numbers for rooms this
  // service no longer counts. Refuse rather than drop them silently.
  const sections = sectionsFor_(service);
  const known = {};
  sections.forEach(function (section) {
    section.areas.forEach(function (area) { known[area] = true; });
  });
  Object.keys(counts).forEach(function (area) {
    if (!known[area] && counts[area] !== '' && counts[area] !== null && counts[area] !== undefined) {
      throw new Error('The room list has changed since this page was opened. Close it, open it again, ' +
        'and re-enter the number for ' + area + '.');
    }
  });
  const record = { 'Date': date, 'Service': service, 'Other Service': otherName };
  // This replaces the whole row, so rooms this service doesn't count are blanked.
  allSections_().forEach(function (section) {
    section.areas.forEach(function (area) { if (!known[area]) record[area] = ''; });
  });
  const total = fillCounts_(record, counts, sections);
  record['Bulletins Left'] = count_(form.bulletins, 'Bulletins Left');
  const doors = form.doors || {};
  allDoors_().forEach(function (n) {
    record[doorHeader_(n)] = clean_(doors[n], 40);
  });
  record['Submitted By'] = clean_(form.submittedBy, 60);
  record['Submitted At'] = new Date();
  record['Notes'] = clean_(form.notes, 500);
  record[REPORT_STATUS] = replace ? 'Pending (updated)' : 'Pending';
  record[OFFICE_EDIT] = '';

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = getSheet_();
    const cols = headerIndex_(sheet);
    const width = sheet.getLastColumn();
    const existing = findRow_(sheet, cols, date, service, otherName, tz);

    if (existing && !replace) {
      const old = sheet.getRange(existing, 1, 1, width).getValues()[0];
      const at = old[cols['Submitted At'] - 1];
      return {
        status: 'duplicate',
        by: String(old[cols['Submitted By'] - 1] || ''),
        at: at instanceof Date ? Utilities.formatDate(at, tz, 'EEE h:mm a') : '',
        total: old[cols['Total Attendance'] - 1],
      };
    }

    writeRow_(sheet, cols, existing, record, ATTENDANCE_FORMATS);
    try {
      queueReport_();
    } catch (err) {
      // The numbers are saved; the report can be sent later from the Attendance menu.
      console.error('Could not schedule the report: ' + err.message);
    }
    return { status: 'saved', replaced: !!existing, total: total };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Puts each area's count from counts ({area: number or ""}) into record, for
 * the areas in sections, with the section subtotals and Total Attendance.
 * Returns the total.
 */
function fillCounts_(record, counts, sections) {
  let total = 0;
  sections.forEach(function (section) {
    let subtotal = 0;
    section.areas.forEach(function (area) {
      const n = count_(counts[area], area);
      record[area] = n;
      subtotal += n || 0;
    });
    record[section.title + ' Subtotal'] = subtotal;
    total += subtotal;
  });
  record['Total Attendance'] = total;
  return total;
}

/**
 * Writes record ({header: value}) into row, or as a new row at the bottom
 * when row is 0, and returns the row. An existing row keeps every column not
 * in record. Writing a Date replaces its column's format with Sheets' default
 * ("10/14/2026 12:00:00"), so formats ({header: number format}) are put back.
 */
function writeRow_(sheet, cols, row, record, formats) {
  const keys = Object.keys(record);
  if (row) {
    const at = keys.map(function (k) { return cols[k]; });
    const first = Math.min.apply(null, at);
    const range = sheet.getRange(row, first, 1, Math.max.apply(null, at) - first + 1);
    const values = range.getValues()[0];
    keys.forEach(function (k) { values[cols[k] - first] = record[k]; });
    range.setValues([values]);
  } else {
    const values = new Array(sheet.getLastColumn()).fill('');
    keys.forEach(function (k) { values[cols[k] - 1] = record[k]; });
    sheet.appendRow(values);
    row = sheet.getLastRow();
  }
  Object.keys(formats).forEach(function (h) {
    if (cols[h]) sheet.getRange(row, cols[h]).setNumberFormat(formats[h]);
  });
  return row;
}

function pageConfig_() {
  return { sections: SECTIONS, serviceSections: SERVICE_SECTIONS, doors: DOORS, services: SERVICES };
}

/** The areas a service counts, by section. */
function sectionsFor_(service) {
  return SERVICE_SECTIONS[service] || SECTIONS;
}

/** Every area any service counts, by section: SECTIONS' areas first, then the others'. */
function allSections_() {
  return SECTIONS.map(function (section) {
    const areas = section.areas.slice();
    Object.keys(SERVICE_SECTIONS).forEach(function (service) {
      SERVICE_SECTIONS[service].forEach(function (s) {
        if (s.title !== section.title) return;
        s.areas.forEach(function (a) { if (areas.indexOf(a) === -1) areas.push(a); });
      });
    });
    return { title: section.title, areas: areas };
  });
}

/**
 * A service's sections, with any of the named areas it doesn't count added
 * to the end of their sections: rooms the service used to count, or another
 * service's rooms, that a row has numbers in.
 */
function sectionsWith_(service, names) {
  const all = allSections_();
  return sectionsFor_(service).map(function (own) {
    const extra = all.filter(function (s) { return s.title === own.title; })[0].areas.filter(function (a) {
      return own.areas.indexOf(a) === -1 && names.indexOf(a) > -1;
    });
    return { title: own.title, areas: own.areas.concat(extra) };
  });
}

/**
 * The sections to show for a logged row, so no number in it goes missing:
 * its service's areas plus any other area it has a number in.
 * get(header) returns that row's value.
 */
function rowSections_(service, get) {
  const names = [];
  allSections_().forEach(function (s) {
    s.areas.forEach(function (a) {
      const v = get(a);
      if (v !== '' && v !== null && v !== undefined) names.push(a);
    });
  });
  return sectionsWith_(service, names);
}

function headers_() {
  const h = ['Date', 'Service', 'Other Service', 'Total Attendance'];
  allSections_().forEach(function (section) {
    section.areas.forEach(function (area) { h.push(area); });
    h.push(section.title + ' Subtotal');
  });
  h.push('Bulletins Left');
  allDoors_().forEach(function (n) { h.push(doorHeader_(n)); });
  h.push('Submitted By', 'Submitted At', 'Notes', REPORT_STATUS, OFFICE_EDIT);
  return h;
}

function allDoors_() {
  return DOORS.reduce(function (all, g) { return all.concat(g.doors); }, []);
}

function doorHeader_(n) {
  return 'Door ' + n + ' Usher';
}

/** Returns the Attendance tab, creating it (and the Summary tab) if needed. */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    const sheets = ss.getSheets();
    // A brand-new spreadsheet has one empty "Sheet1": reuse it.
    if (sheets.length === 1 && sheets[0].getLastRow() === 0) {
      sheet = sheets[0].setName(SHEET_NAME);
    } else {
      sheet = ss.insertSheet(SHEET_NAME, 0);
    }
  }
  const cols = headerIndex_(sheet);
  ensureSummary_(ss, cols);
  return sheet;
}

/** Makes sure every header exists (adding missing ones on the right) and maps name -> column. */
function headerIndex_(sheet) {
  const lastCol = sheet.getLastColumn();
  const current = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  const missing = headers_().filter(function (h) { return current.indexOf(h) === -1; });

  if (missing.length) {
    const start = current.length + 1;
    if (sheet.getMaxColumns() < current.length + missing.length) {
      sheet.insertColumnsAfter(sheet.getMaxColumns(), current.length + missing.length - sheet.getMaxColumns());
    }
    sheet.getRange(1, start, 1, missing.length).setValues([missing]);
    const all = current.concat(missing);
    formatSheet_(sheet, all);
    current.push.apply(current, missing);
  }

  const index = {};
  current.forEach(function (h, i) { if (h) index[h] = i + 1; });
  // New columns can be new rooms, which the service tabs should show.
  if (missing.length) ensureServiceViews_(SpreadsheetApp.getActiveSpreadsheet(), index);
  return index;
}

function formatSheet_(sheet, headers) {
  const header = sheet.getRange(1, 1, 1, headers.length);
  header.setFontWeight('bold').setFontColor('#ffffff').setBackground('#233617')
    .setWrap(true).setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  headers.forEach(function (h, i) {
    const col = sheet.getRange(1, i + 1, sheet.getMaxRows(), 1);
    if (h === 'Date') col.offset(1, 0, sheet.getMaxRows() - 1).setNumberFormat(DATE_FORMAT);
    if (h === 'Submitted At') col.offset(1, 0, sheet.getMaxRows() - 1).setNumberFormat(SUBMITTED_FORMAT);
    if (h === 'Total Attendance' || /Subtotal$/.test(h)) {
      sheet.getRange(1, i + 1).setBackground('#4E7A35');
      col.offset(1, 0, sheet.getMaxRows() - 1).setFontWeight('bold');
    }
  });
}

/** Summary tab: one row per date, one column per service, showing totals. */
function ensureSummary_(ss, cols) {
  let summary = ss.getSheetByName(SUMMARY_NAME);
  if (!summary) summary = ss.insertSheet(SUMMARY_NAME);
  const ref = function (h) { return columnRef_(cols, h); };
  const formula = '=QUERY({' + ref('Date') + ', ' + ref('Service') + ', ' + ref('Total Attendance') + '}, ' +
    '"select Col1, sum(Col3) where Col1 is not null group by Col1 pivot Col2 order by Col1 desc", 1)';
  const cell = summary.getRange(1, 1);
  if (cell.getFormula() !== formula) {
    cell.setFormula(formula);
    summary.getRange(1, 1, 1, 6).setFontWeight('bold').setFontColor('#ffffff').setBackground('#233617');
    summary.getRange('A2:A').setNumberFormat(DATE_FORMAT);
    summary.setFrozenRows(1);
  }
}

/** "'Attendance'!D:D": a whole Attendance column, for formulas. */
function columnRef_(cols, h) {
  const letter = columnLetter_(cols[h]);
  return "'" + SHEET_NAME + "'!" + letter + ':' + letter;
}

/**
 * One tab per service (Sunday AM, Sunday PM, Wednesday, Other Services),
 * showing that service's rows from the Attendance tab with only its own
 * rooms, newest first. Each is a formula, so it fills itself in; type
 * changes on the Attendance tab or the office page, not here. Rebuilt when
 * the Attendance tab gets new columns.
 */
function ensureServiceViews_(ss, cols) {
  SERVICES.forEach(function (service, i) {
    const name = service === 'Other' ? 'Other Services' : service;
    const headers = ['Date'].concat(service === 'Other' ? ['Other Service'] : [], ['Total Attendance']);
    sectionsFor_(service).forEach(function (section) {
      headers.push.apply(headers, section.areas);
      headers.push(section.title + ' Subtotal');
    });
    headers.push('Bulletins Left', 'Submitted By', 'Notes');
    // Col1 is the Service column, to pick the rows; the rest are shown.
    const refs = [columnRef_(cols, 'Service')].concat(headers.map(function (h) { return columnRef_(cols, h); }));
    const shown = headers.map(function (h, c) { return 'Col' + (c + 2); });
    const formula = '=QUERY({' + refs.join(', ') + '}, "select ' + shown.join(', ') +
      " where Col1 = '" + service + "' order by Col2 desc\", 1)";

    let tab = ss.getSheetByName(name);
    if (!tab) {
      const summary = ss.getSheetByName(SUMMARY_NAME);
      tab = ss.insertSheet(name, summary ? summary.getIndex() + i : ss.getNumSheets());
    }
    const cell = tab.getRange(1, 1);
    if (cell.getFormula() === formula) return;
    tab.clear();
    cell.setFormula(formula);
    if (tab.getMaxColumns() < headers.length) tab.insertColumnsAfter(tab.getMaxColumns(), headers.length - tab.getMaxColumns());
    tab.getRange(1, 1, 1, headers.length).setFontWeight('bold').setFontColor('#ffffff').setBackground('#233617')
      .setWrap(true).setVerticalAlignment('middle');
    headers.forEach(function (h, c) {
      if (h === 'Total Attendance' || /Subtotal$/.test(h)) {
        tab.getRange(1, c + 1).setBackground('#4E7A35');
        tab.getRange(2, c + 1, tab.getMaxRows() - 1, 1).setFontWeight('bold');
      }
    });
    tab.getRange(2, 1, tab.getMaxRows() - 1, 1).setNumberFormat(DATE_FORMAT);
    tab.setFrozenRows(1);
    tab.setFrozenColumns(1);
    tab.setTabColor('#4E7A35');
  });
}

function columnLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Finds the row already logged for this date + service, if any. */
function findRow_(sheet, cols, date, service, otherName, tz) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;
  const n = lastRow - 1;
  const dates = sheet.getRange(2, cols['Date'], n, 1).getValues();
  const services = sheet.getRange(2, cols['Service'], n, 1).getValues();
  const others = sheet.getRange(2, cols['Other Service'], n, 1).getValues();
  const day = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  const other = otherName.toLowerCase();
  for (let i = n - 1; i >= 0; i--) {
    const d = dates[i][0];
    if (!(d instanceof Date)) continue;
    if (Utilities.formatDate(d, tz, 'yyyy-MM-dd') === day &&
        services[i][0] === service &&
        String(others[i][0]).toLowerCase() === other) {
      return i + 2;
    }
  }
  return 0;
}

/** "2026-10-11" -> Date at noon, so a time zone difference can't shift the day. */
function parseDate_(value) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(value || ''));
  if (!m) throw new Error('Please pick a date.');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 12);
}

function count_(value, label) {
  if (value === '' || value === null || value === undefined) return '';
  const n = Number(value);
  if (!isFinite(n) || n < 0 || n > MAX_COUNT || Math.floor(n) !== n) {
    throw new Error('The number for ' + label + ' doesn\'t look right.');
  }
  return n;
}

/** Trims text and stops it from being read as a spreadsheet formula. */
function clean_(value, max) {
  let s = String(value === null || value === undefined ? '' : value).trim().slice(0, max);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}
