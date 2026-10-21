/**
 * Grace Baptist Church - Usher Attendance: history from the old sheet
 *
 * Before the ushers' page, attendance was kept by hand in the "Attendance
 * Report" spreadsheet: one row per Sunday from July 2006, with the Sunday AM,
 * Sunday PM and Wednesday counts, Sunday School and the specific ministries,
 * plus a Children's Ministries tab from 2024. importHistory copies it into
 * this spreadsheet, one "History 2006", "History 2007", ... tab per year.
 *
 * The report (Report.gs) reads those tabs beside the Attendance tab, so
 * recent services and comparisons with last year reach back before the page.
 *
 * Run importHistory once from the editor (or the Attendance menu). Running it
 * again rebuilds the History tabs from the old sheet, so edit the old sheet,
 * not the History tabs, if a past number needs fixing before the next import.
 *
 * Nothing is left out silently: every non-blank row of the old sheet that
 * isn't copied, and any cells cut from the right of a copied row, are listed
 * with the reason on a "History Import Log" tab.
 */

const OLD_SHEET_ID = '17uKsSNSYH2jGkFCK9bjijiRYlm2R8vZh9xJMxFiwWo8';
const OLD_MAIN_TAB = 'Attendance Report';
const OLD_CHILDREN_TAB = "Children's Ministries Attendance";
const HISTORY_PREFIX = 'History ';
const HISTORY_LOG = 'History Import Log';

// Old sheet columns A-M, then the notes in N.
const HISTORY_HEADERS = ['Date', 'Bus', 'Keenagers/HGM', 'K4T/AWANA', 'Sunday AM', 'Auditorium', 'Kidstown',
  'Sunday PM', 'Wednesday', 'SS Total', 'SS Children', 'SS Teen', 'SS Adult', 'Notes'];
// Children's Ministries columns B-P, added to the years that have them.
const CHILDREN_HEADERS = ['SS Nursery', 'SS K4-K5', 'SS 1st-2nd', 'SS 3rd-4th', 'SS 5th-6th', 'SS Teens',
  'AM Nursery', 'AM K4-K5', 'AM 1st-3rd', 'PM Nursery', 'Wed Nursery', 'Wed K4-K5', 'Wed 1st-3rd', 'Wed 4th-6th', 'Wed Teens'];
// Which History column holds each service's count. Wednesday's count sits on
// the Sunday row of its week, so its date is that Sunday plus three days.
const HISTORY_SERVICES = {
  'Sunday AM': { column: 'Sunday AM', days: 0 },
  'Sunday PM': { column: 'Sunday PM', days: 0 },
  'Wednesday': { column: 'Wednesday', days: 3 },
};
// Labels of the averages tables kept to the right of the old sheet; not notes.
const TABLE_LABELS = /^(SS|AM|PM|SUN AM|SUN PM|AWANA|WED|Wed|Wednesday|Kidstown|Monday|Teens:|Helpers:|Helper's Kids:|YOUTH RALLY.*|Attendance Averages.*)$/;

/** Copies the old sheet into one History tab per year. Shows what it did in a message and the Execution log. */
function importHistory() {
  const old = SpreadsheetApp.openById(OLD_SHEET_ID);
  const log = [];   // [old tab, old row, reason, what was in it]
  const years = parseOldMain_(old.getSheetByName(OLD_MAIN_TAB).getDataRange().getDisplayValues(), log);
  const kids = old.getSheetByName(OLD_CHILDREN_TAB);
  if (kids) mergeChildren_(years, parseOldChildren_(kids.getDataRange().getDisplayValues(), log), log);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = [];
  Object.keys(years).sort().forEach(function (year) {
    const rows = years[year].sort(function (a, b) { return a.date - b.date; });
    const withKids = rows.some(function (r) { return r.kids; });
    const headers = HISTORY_HEADERS.concat(withKids ? CHILDREN_HEADERS : []);
    const values = rows.map(function (r) {
      return [r.date].concat(r.main, [r.notes.join('; ')], withKids ? (r.kids || CHILDREN_HEADERS.map(function () { return ''; })) : []);
    });
    const name = HISTORY_PREFIX + year;
    let tab = ss.getSheetByName(name);
    if (tab) tab.clear(); else tab = ss.insertSheet(name, ss.getNumSheets());
    tab.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (values.length) tab.getRange(2, 1, values.length, headers.length).setValues(values);
    styleHistoryTab_(tab, headers.length, values.length);
    summary.push(name + ': ' + values.length + ' rows');
  });
  writeImportLog_(ss, log);
  orderHistoryTabs_(ss);
  const skipped = log.filter(function (l) { return l[2].indexOf('Copied') !== 0 && l[2].indexOf('Added') !== 0; }).length;
  const message = 'Imported ' + summary.length + ' years.\n' + summary.join('\n') +
    '\n\n' + log.length + (log.length === 1 ? ' entry' : ' entries') + ' on the ' + HISTORY_LOG + ' tab (' + skipped + ' rows not copied). Look them over for anything that should have been kept.';
  Logger.log(message);
  try { SpreadsheetApp.getUi().alert(message); } catch (e) { /* run from the editor */ }
  return summary;
}

/** A row's non-blank cells as "A: 7/2/06 | B: 93 | ...", for the import log. */
function rowText_(v) {
  const out = [];
  v.forEach(function (x, i) {
    const t = String(x || '').trim();
    if (t) out.push(columnLetter_(i + 1) + ': ' + t);
  });
  return out.join(' | ');
}

/** Puts the History tabs after the other tabs, newest year first, with the import log last. */
function orderHistoryTabs_(ss) {
  const history = ss.getSheets().filter(function (t) { return /^History \d{4}$/.test(t.getName()); })
    .sort(function (a, b) { return b.getName().localeCompare(a.getName()); });
  const log = ss.getSheetByName(HISTORY_LOG);
  const others = ss.getSheets().length - history.length - (log ? 1 : 0);
  history.concat(log ? [log] : []).forEach(function (tab, i) {
    ss.setActiveSheet(tab);
    ss.moveActiveSheet(others + i + 1);
  });
  ss.setActiveSheet(ss.getSheets()[0]);
}

/** Writes the import log tab, replacing an earlier one. */
function writeImportLog_(ss, log) {
  let tab = ss.getSheetByName(HISTORY_LOG);
  if (tab) tab.clear(); else tab = ss.insertSheet(HISTORY_LOG, ss.getNumSheets());
  const rows = [['Old tab', 'Old row', 'What happened', 'What was in the row']].concat(
    log.length ? log : [['', '', 'Every row with anything in it was copied.', '']]);
  tab.getRange(1, 1, rows.length, 4).setValues(rows).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('top').setWrap(true);
  tab.getRange(1, 1, 1, 4).setBackground('#233617').setFontColor('#FFFFFF').setFontWeight('bold');
  tab.setFrozenRows(1);
  [190, 70, 300, 620].forEach(function (w, i) { tab.setColumnWidth(i + 1, w); });
  tab.setTabColor('#8A5200');
}

/** "7/2/06" -> [7, 2, 2006]; anything else -> null. */
function parseMDY_(text) {
  const m = /^\s*(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s*$/.exec(String(text));
  if (!m) return null;
  const y = Number(m[3]);
  return [Number(m[1]), Number(m[2]), y < 100 ? 2000 + y : y];
}

/** A number typed as text ("1,038") becomes a number; everything else stays as typed. */
function cellValue_(text) {
  const t = String(text).trim();
  return /^-?\d[\d,]*(\.\d+)?$/.test(t) ? Number(t.replace(/,/g, '')) : t;
}

/**
 * A row's date, using the year of the section it's under: the old sheet has a
 * few typos ("2/4/06" in the 2007 section). Returns {date, note}.
 */
function sectionDate_(text, year) {
  const p = parseMDY_(text);
  if (!p || p[0] < 1 || p[0] > 12 || p[1] < 1 || p[1] > 31) return null;
  const date = new Date(year, p[0] - 1, p[1], 12);
  return { date: date, note: p[2] === year ? '' : 'Date entered as ' + String(text).trim() };
}

/** Why a row with no usable date in column A wasn't copied. */
function noDateReason_(a) {
  if (/^[A-Za-z]{3}-\d{2}$/.test(a)) return 'Not copied: a monthly average row ("' + a + '"), not a single Sunday';
  if (parseMDY_(a)) return 'Not copied: the date in column A couldn\'t be read';
  return a ? 'Not copied: column A isn\'t a date' : 'Not copied: no date in column A';
}

/** The Attendance Report tab -> {year: [{date, main: [12 values], notes: []}]}. Skipped rows go in log. */
function parseOldMain_(values, log) {
  const years = {};
  const skip = function (i, v, reason) { if (rowText_(v)) log.push([OLD_MAIN_TAB, i + 1, reason, rowText_(v)]); };
  let year = null;
  values.forEach(function (v, i) {
    const a = String(v[0] || '').trim();
    if (/^(19|20)\d\d$/.test(a)) {
      year = Number(a);
      if (rowText_(v.slice(1))) skip(i, v, 'Used as the year heading; the other cells in the row weren\'t copied');
      return;
    }
    if (!year) { skip(i, v, 'Not copied: above the first year heading (titles and column headings)'); return; }
    const d = sectionDate_(a, year);
    if (!d) { skip(i, v, noDateReason_(a)); return; }   // blank rows, monthly averages like "Aug-06", the totals at the bottom
    const main = v.slice(1, 13).map(cellValue_);
    while (main.length < 12) main.push('');
    const notes = [];
    if (d.note) notes.push(d.note);
    // Notes are in column N, sometimes running on into the next columns (like "Mon-184", "Tue-228", ... "Total-359").
    const n = String(v[13] || '').trim();
    let used = 13;   // the last column copied
    if (n && !TABLE_LABELS.test(n) && typeof cellValue_(n) !== 'number') {
      notes.push(n);
      for (let c = 14; c < v.length; c++) {
        const more = String(v[c] || '').trim();
        if (!more || TABLE_LABELS.test(more) || typeof cellValue_(more) === 'number') break;
        notes.push(more);
        used = c;
      }
    } else {
      used = 12;
    }
    // A row with nothing but blanks and zeros was a placeholder for a Sunday never filled in.
    const filled = main.some(function (x) { return x !== '' && x !== 0; }) || notes.some(function (x) { return x !== d.note; });
    if (!filled) { skip(i, v, 'Not copied: a date with no numbers or notes (a Sunday never filled in)'); return; }
    // Cells past the notes are the averages tables kept beside the rows.
    const extra = v.slice(used + 1).map(function (x, k) { return String(x || '').trim() ? columnLetter_(used + 2 + k) + ': ' + String(x).trim() : ''; })
      .filter(Boolean);
    if (extra.length) log.push([OLD_MAIN_TAB, i + 1, 'Copied, except these cells to the right (averages tables)', extra.join(' | ')]);
    (years[year] = years[year] || []).push({ date: d.date, main: main, notes: notes });
  });
  return years;
}

/** The Children's Ministries tab -> [{year, date, kids: [15 values], note, row}]. Its year headers are in column B. */
function parseOldChildren_(values, log) {
  const out = [];
  const skip = function (i, v, reason) { if (rowText_(v)) log.push([OLD_CHILDREN_TAB, i + 1, reason, rowText_(v)]); };
  let year = null;
  values.forEach(function (v, i) {
    const a = String(v[0] || '').trim();
    const b = String(v[1] || '').trim();
    if (!a && /^(19|20)\d\d$/.test(b)) {
      year = Number(b);
      if (rowText_(v.slice(2))) skip(i, v, 'Used as the year heading; the other cells in the row weren\'t copied');
      return;
    }
    if (!year) { skip(i, v, 'Not copied: above the first year heading (titles and column headings)'); return; }
    const d = sectionDate_(a, year);
    if (!d) { skip(i, v, noDateReason_(a)); return; }
    const kids = v.slice(1, 16).map(cellValue_);
    while (kids.length < 15) kids.push('');
    const note = [d.note ? 'Children: ' + d.note.toLowerCase() : '', String(v[16] || '').trim()].filter(Boolean).join('; ');
    if (!kids.some(function (x) { return x !== ''; }) && !note) { skip(i, v, 'Not copied: a date with no numbers or notes'); return; }
    const extra = v.slice(17).map(function (x, k) { return String(x || '').trim() ? columnLetter_(18 + k) + ': ' + String(x).trim() : ''; }).filter(Boolean);
    if (extra.length) log.push([OLD_CHILDREN_TAB, i + 1, 'Copied, except these cells to the right', extra.join(' | ')]);
    out.push({ year: year, date: d.date, kids: kids, note: note, row: i + 1, text: rowText_(v) });
  });
  return out;
}

/** Puts each children's row on the same date's row of that year, or a row of its own (noted in the log). */
function mergeChildren_(years, kids, log) {
  const key = function (d) { return d.getMonth() + '/' + d.getDate(); };
  kids.forEach(function (k) {
    const rows = years[k.year] = years[k.year] || [];
    const match = rows.filter(function (r) { return !r.kids && key(r.date) === key(k.date); })[0];
    const row = match || { date: k.date, main: HISTORY_HEADERS.slice(1, 13).map(function () { return ''; }), notes: [] };
    row.kids = k.kids;
    if (k.note) row.notes.push(k.note);
    if (!match) {
      rows.push(row);
      log.push([OLD_CHILDREN_TAB, k.row, 'Added as its own row: no ' + k.year + ' Sunday with this date in the main tab (check the date)', k.text]);
    }
  });
}

function styleHistoryTab_(tab, cols, rows) {
  const all = tab.getRange(1, 1, rows + 1, cols);
  all.setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  tab.getRange(1, 1, 1, cols).setBackground('#233617').setFontColor('#FFFFFF').setFontWeight('bold').setWrap(true)
    .setHorizontalAlignment('center');
  if (rows) {
    tab.getRange(2, 1, rows, 1).setNumberFormat('M/d/yyyy').setHorizontalAlignment('left');
    tab.getRange(2, 2, rows, 12).setHorizontalAlignment('right');
    tab.getRange(2, 14, rows, 1).setWrap(true).setFontColor('#5B6A50');
    if (cols > 14) tab.getRange(2, 15, rows, cols - 14).setHorizontalAlignment('right');
  }
  tab.setFrozenRows(1);
  tab.setFrozenColumns(1);
  tab.setColumnWidth(1, 90);
  for (let c = 2; c <= cols; c++) tab.setColumnWidth(c, c === 14 ? 260 : 86);
  tab.setTabColor('#7BAE58');
}

/**
 * Past counts of a service from the History tabs, for the years given:
 * [{date, total}], counting only rows where that service has a number.
 */
function historyCounts_(service, years) {
  const spec = HISTORY_SERVICES[service];
  if (!spec) return [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const out = [];
  years.forEach(function (year) {
    const tab = ss.getSheetByName(HISTORY_PREFIX + year);
    if (!tab || tab.getLastRow() < 2) return;
    const values = tab.getRange(1, 1, tab.getLastRow(), tab.getLastColumn()).getValues();
    const col = values[0].indexOf(spec.column);
    if (col < 0) return;
    values.slice(1).forEach(function (r) {
      const n = r[col];
      if (!(r[0] instanceof Date) || typeof n !== 'number' || n <= 0) return;
      out.push({ date: new Date(r[0].getTime() + spec.days * 86400000), total: n });
    });
  });
  return out;
}
