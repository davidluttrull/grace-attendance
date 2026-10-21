/**
 * Grace Baptist Church - Usher Attendance
 *
 * Back end for the ushers' attendance page (index.html, hosted on GitHub
 * Pages). The page reads the room list from doGet and sends each service's
 * numbers to doPost, which logs them as one row in this spreadsheet.
 * See SETUP.md for how to deploy it.
 *
 * Report.gs emails a report of each submission, as the body of the email, to
 * the addresses on the Report Emails tab.
 *
 * To add, rename or reorder an area, edit SECTIONS below and deploy a new
 * version. New areas get a new column at the right end of the sheet; old
 * columns are never deleted, so past data stays put.
 */

const SHEET_NAME = 'Attendance';
const SUMMARY_NAME = 'Summary';

const SERVICES = ['Sunday AM', 'Sunday PM', 'Wednesday', 'Other'];

// Areas counted each service, in the order ushers see them.
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
    if (!codeIsRight_(body.code)) return json_({ status: 'badcode' });
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
 * Run this once from the editor to create the tabs and grant permissions.
 * It also creates the usher code (the first time) and shows it in the
 * Execution log. To change the code, edit USHER_CODE under Project Settings >
 * Script Properties.
 */
function setup() {
  const sheet = getSheet_();
  ensureEmailsTab_();
  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);

  const props = PropertiesService.getScriptProperties();
  let code = props.getProperty('USHER_CODE');
  if (!code) {
    code = String(Math.floor(100000 + Math.random() * 900000));
    props.setProperty('USHER_CODE', code);
  }
  Logger.log('Usher code: ' + code);
}

/**
 * Checks the usher code against the USHER_CODE Script Property, refusing
 * everything for 10 minutes after too many wrong ones. The code isn't in this
 * file, so the public copy on GitHub doesn't give it away.
 */
function codeIsRight_(code) {
  const expected = PropertiesService.getScriptProperties().getProperty('USHER_CODE');
  if (!expected) throw new Error('The usher code hasn\'t been set up yet. Run setup in Apps Script.');
  const cache = CacheService.getScriptCache();
  const misses = Number(cache.get('badUsherCode') || 0);
  if (misses >= MAX_BAD_CODES) throw new Error('Too many wrong codes. Wait 10 minutes and try again.');
  if (String(code || '').trim() === expected.trim()) return true;
  cache.put('badUsherCode', String(misses + 1), 600);
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
  // A phone still showing an old room list would send numbers for rooms that
  // no longer exist. Refuse rather than drop them silently.
  const known = {};
  SECTIONS.forEach(function (section) {
    section.areas.forEach(function (area) { known[area] = true; });
  });
  Object.keys(counts).forEach(function (area) {
    if (!known[area] && counts[area] !== '' && counts[area] !== null && counts[area] !== undefined) {
      throw new Error('The room list has changed since this page was opened. Close it, open it again, ' +
        'and re-enter the number for ' + area + '.');
    }
  });
  const record = { 'Date': date, 'Service': service, 'Other Service': otherName };
  let total = 0;
  SECTIONS.forEach(function (section) {
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
  record['Bulletins Left'] = count_(form.bulletins, 'Bulletins Left');
  const doors = form.doors || {};
  allDoors_().forEach(function (n) {
    record[doorHeader_(n)] = clean_(doors[n], 40);
  });
  record['Submitted By'] = clean_(form.submittedBy, 60);
  record['Submitted At'] = new Date();
  record['Notes'] = clean_(form.notes, 500);
  record[REPORT_STATUS] = replace ? 'Pending (updated)' : 'Pending';

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

    const row = new Array(width).fill('');
    Object.keys(record).forEach(function (key) {
      row[cols[key] - 1] = record[key];
    });
    if (existing) {
      sheet.getRange(existing, 1, 1, width).setValues([row]);
    } else {
      sheet.appendRow(row);
    }
    // Writing a Date replaces the column's format with Sheets' default
    // "10/14/2026 12:00:00", so put the column's format back on this row.
    const saved = existing || sheet.getLastRow();
    sheet.getRange(saved, cols['Date']).setNumberFormat(DATE_FORMAT);
    sheet.getRange(saved, cols['Submitted At']).setNumberFormat(SUBMITTED_FORMAT);
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

function pageConfig_() {
  return { sections: SECTIONS, doors: DOORS, services: SERVICES };
}

function headers_() {
  const h = ['Date', 'Service', 'Other Service', 'Total Attendance'];
  SECTIONS.forEach(function (section) {
    section.areas.forEach(function (area) { h.push(area); });
    h.push(section.title + ' Subtotal');
  });
  h.push('Bulletins Left');
  allDoors_().forEach(function (n) { h.push(doorHeader_(n)); });
  h.push('Submitted By', 'Submitted At', 'Notes', REPORT_STATUS);
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
  const ref = function (h) {
    const letter = columnLetter_(cols[h]);
    return "'" + SHEET_NAME + "'!" + letter + ':' + letter;
  };
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
