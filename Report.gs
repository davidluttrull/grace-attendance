/**
 * Grace Baptist Church - Usher Attendance: emailed report
 *
 * After each submission, Code.gs marks the row "Pending" in the Report
 * Emailed column and schedules sendPendingReports a few seconds later, so the
 * usher's phone doesn't wait on the email. sendPendingReports builds the
 * report and emails it to the addresses on the Report Emails tab. The report
 * is the body of the email, as HTML styled like the web page, with a plain-
 * text version for email apps that don't show HTML. Recent services and the
 * comparison with last year also read the History tabs (History.gs), so they
 * reach back before the ushers' page.
 */

const EMAILS_NAME = 'Report Emails';
const REPORT_STATUS = 'Report Emailed';
// The web page's logo, made for its dark green header (600 x 210).
const LOGO_URL = 'https://davidluttrull.github.io/grace-attendance/logo.png';
const HISTORY_COUNT = 8;

// The web page's light theme.
const THEME = {
  deep: '#233617',    // header bar and totals
  brand: '#7BAE58',   // heading underlines
  mid: '#4E7A35',     // heading text
  ink: '#1E2A16',
  muted: '#5B6A50',
  line: '#DCE7D3',    // borders
  soft: '#F3F8EF',    // page background and subtotal rows
  track: '#EEF2EA',   // empty part of a bar
};
// One color per section, in order, with the text color that reads on it.
const SECTION_COLORS = [
  { bg: '#4E7A35', ink: '#FFFFFF' },
  { bg: '#7BAE58', ink: '#10160C' },
  { bg: '#8A9A7E', ink: '#FFFFFF' },
  { bg: '#B8CFA6', ink: '#1E2A16' },
];

/** Adds an Attendance menu to the Sheet. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Attendance')
    .addItem('Email report for selected row', 'emailSelectedReport')
    .addItem('Send me a test report for selected row', 'emailSelectedReportToMe')
    .addSeparator()
    .addItem('Email weekly report', 'emailWeeklyReportMenu')
    .addItem('Send me a test weekly report', 'emailWeeklyReportToMe')
    .addSeparator()
    .addItem('Import history from the old attendance sheet', 'importHistory')
    .addToUi();
}

/** Called by submitAttendance after a row is saved. */
function queueReport_() {
  const pending = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'sendPendingReports';
  });
  if (!pending) {
    ScriptApp.newTrigger('sendPendingReports').timeBased().after(5 * 1000).create();
  }
}

/** Trigger handler: emails a report for every row marked Pending. */
function sendPendingReports() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'sendPendingReports') ScriptApp.deleteTrigger(t);
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sheet = getSheet_();
    const cols = headerIndex_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const status = sheet.getRange(2, cols[REPORT_STATUS], lastRow - 1, 1).getValues();
    status.forEach(function (s, i) {
      if (String(s[0]).indexOf('Pending') === 0) {
        emailReportForRow_(sheet, cols, i + 2, reportRecipients_(), String(s[0]).indexOf('update') > -1);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

/** Menu: (re)send the report for the row the cursor is on to everyone on the list. */
function emailSelectedReport() {
  sendSelected_(reportRecipients_());
}

/** Menu: send the selected row's report only to whoever is signed in, to check how it looks. */
function emailSelectedReportToMe() {
  sendSelected_([Session.getEffectiveUser().getEmail()]);
}

function sendSelected_(recipients) {
  const ui = SpreadsheetApp.getUi();
  const sheet = getSheet_();
  const row = SpreadsheetApp.getActiveRange().getRow();
  if (SpreadsheetApp.getActiveSheet().getName() !== SHEET_NAME || row < 2) {
    ui.alert('Click a row on the ' + SHEET_NAME + ' tab first, then try again.');
    return;
  }
  const result = emailReportForRow_(sheet, headerIndex_(sheet), row, recipients, false);
  ui.alert(result);
}

/** Builds and emails one row's report, and records the outcome in the Report Emailed column. */
function emailReportForRow_(sheet, cols, row, recipients, updated) {
  const cell = sheet.getRange(row, cols[REPORT_STATUS]);
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  if (!recipients.length) {
    cell.setValue('Not sent: no addresses on the ' + EMAILS_NAME + ' tab');
    return 'There are no addresses on the ' + EMAILS_NAME + ' tab, so nothing was sent.';
  }
  try {
    const data = reportData_(sheet, cols, row, tz);
    const label = serviceName_(data) + ', ' + Utilities.formatDate(data.date, tz, 'MMM d');
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: (updated ? 'Updated attendance: ' : 'Attendance: ') + label + ' (' + data.total + ')',
      htmlBody: reportHtml_(data, updated),
      body: emailText_(data, updated),
      name: 'Grace Attendance',
    });
    cell.setValue('Sent ' + Utilities.formatDate(new Date(), tz, 'M/d h:mm a') + ' to ' + recipients.length);
    return 'Report emailed to ' + recipients.join(', ') + '.';
  } catch (err) {
    cell.setValue('Failed: ' + err.message);
    return 'The report could not be sent: ' + err.message;
  }
}

/** Addresses from the Report Emails tab (column A, below the heading). */
function reportRecipients_() {
  const tab = ensureEmailsTab_();
  const last = tab.getLastRow();
  if (last < 2) return [];
  return tab.getRange(2, 1, last - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s); });
}

function ensureEmailsTab_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let tab = ss.getSheetByName(EMAILS_NAME);
  if (!tab) {
    tab = ss.insertSheet(EMAILS_NAME);
    tab.getRange(1, 1, 1, 2).setValues([['Email the report to', 'One address per row in column A. ' +
      'A Google Group or routing address works too. Leave the list empty to stop sending.']]);
    tab.getRange(1, 1).setFontWeight('bold').setFontColor('#ffffff').setBackground('#233617');
    tab.getRange(1, 2).setFontColor(THEME.muted).setFontStyle('italic');
    tab.setColumnWidth(1, 280);
    tab.setFrozenRows(1);
  }
  return tab;
}

// ---- Data -------------------------------------------------------------------

/** Everything the report shows, read from one row plus that service's history. */
function reportData_(sheet, cols, row, tz) {
  const width = sheet.getLastColumn();
  const all = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  const v = all[row - 2];
  const get = function (r, h) { return cols[h] ? r[cols[h] - 1] : ''; };
  const num = function (x) { return x === '' || x === null ? null : Number(x); };

  const data = {
    date: get(v, 'Date'),
    service: get(v, 'Service'),
    otherName: get(v, 'Other Service'),
    total: Number(get(v, 'Total Attendance')) || 0,
    bulletins: num(get(v, 'Bulletins Left')),
    submittedBy: get(v, 'Submitted By'),
    submittedAt: get(v, 'Submitted At'),
    notes: get(v, 'Notes'),
    tz: tz,
    sections: SECTIONS.map(function (s) {
      const areas = s.areas.map(function (a) { return { name: a, n: num(get(v, a)) }; });
      return {
        title: s.title,
        areas: areas,
        subtotal: areas.reduce(function (t, a) { return t + (a.n || 0); }, 0),
      };
    }),
    doors: allDoors_().map(function (n) { return { n: n, name: get(v, doorHeader_(n)) }; }),
    history: [],
    sameDay: [],
  };

  const day = function (d) { return d instanceof Date ? Utilities.formatDate(d, tz, 'yyyy-MM-dd') : ''; };
  const thisDay = day(data.date);
  all.forEach(function (r, i) {
    if (!(get(r, 'Date') instanceof Date)) return;
    const entry = { date: get(r, 'Date'), total: Number(get(r, 'Total Attendance')) || 0, current: i === row - 2 };
    if (get(r, 'Service') === data.service &&
        String(get(r, 'Other Service')).toLowerCase() === String(data.otherName).toLowerCase() &&
        day(entry.date) <= thisDay) {
      data.history.push(entry);
    }
    if (day(entry.date) === thisDay && i !== row - 2) {
      data.sameDay.push({ name: get(r, 'Service') === 'Other' ? get(r, 'Other Service') : get(r, 'Service'), total: entry.total });
    }
  });
  // Counts from before the ushers' page, from the History tabs (History.gs).
  // The Attendance tab wins when both have the same service on the same day.
  if (data.service !== 'Other' && data.date instanceof Date) {
    const year = data.date.getFullYear();
    const logged = {};
    data.history.forEach(function (e) { logged[day(e.date)] = true; });
    historyCounts_(data.service, [year - 2, year - 1, year]).forEach(function (e) {
      if (!logged[day(e.date)] && day(e.date) <= thisDay) data.history.push({ date: e.date, total: e.total, current: false, past: true });
    });
  }
  data.history.sort(function (a, b) { return a.date - b.date; });
  data.compare = compareYears_(data.history, data.date, data.total);
  data.history = data.history.slice(-HISTORY_COUNT);
  return data;
}

/**
 * This service against last year: the same week a year ago (the service
 * closest to 52 weeks back, within 3 days), and this year's average so far
 * beside last year's over the same weeks. Calendar years.
 */
function compareYears_(series, date, total) {
  if (!(date instanceof Date)) return null;
  const DAY = 86400000;
  const target = date.getTime() - 364 * DAY;
  let lastYear = null;
  series.forEach(function (e) {
    const gap = Math.abs(e.date.getTime() - target);
    if (!e.current && gap <= 3 * DAY && (!lastYear || gap < Math.abs(lastYear.date.getTime() - target))) lastYear = e;
  });
  const y = date.getFullYear();
  const sameDayLastYear = new Date(y - 1, date.getMonth(), date.getDate(), 23);
  const avg = function (list) {
    return list.length ? Math.round(list.reduce(function (t, e) { return t + e.total; }, 0) / list.length) : null;
  };
  const thisYear = series.filter(function (e) { return e.date.getFullYear() === y && e.date <= date; });
  const before = series.filter(function (e) { return e.date.getFullYear() === y - 1 && e.date <= sameDayLastYear; });
  return {
    year: y,
    lastYear: lastYear ? { date: lastYear.date, total: lastYear.total, change: total - lastYear.total } : null,
    average: avg(thisYear), count: thisYear.length,
    lastAverage: avg(before), lastCount: before.length,
  };
}

function serviceName_(data) {
  return data.service === 'Other' ? data.otherName : data.service;
}

// ---- Email HTML -------------------------------------------------------------
//
// The report is the email itself, styled like the web pages. Email apps ignore
// most CSS, so everything is tables, inline styles and inline-block slots:
// side by side on a computer, stacked and centered on a phone.

const FONT = "font-family:Poppins,'Helvetica Neue',Helvetica,Arial,sans-serif;";
const DASH = '<span style="color:#AAAAAA">—</span>';

function esc_(s) {
  return String(s === null || s === undefined ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

/** "Infant Nursery (0-1 yrs)" -> {name: "Infant Nursery", detail: "0-1 yrs"} */
function splitName_(s) {
  const m = /^(.*?)\s*\(([^)]*)\)$/.exec(s);
  return m ? { name: m[1], detail: m[2] } : { name: s, detail: '' };
}

function pct_(n, total) {
  return total ? (Math.round(n * 1000 / total) / 10).toFixed(1) + '%' : '';
}

function fmtDate_(d, tz, pattern) {
  return d instanceof Date ? Utilities.formatDate(d, tz, pattern) : '';
}

/** A full-width layout table. */
function eTable_(rows, style) {
  return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;' + (style || '') + '">' + rows + '</table>';
}

/** One white card on the page, like the web page's cards. */
function eCard_(inner) {
  return '<tr><td style="padding-top:12px">' +
    eTable_('<tr><td style="padding:16px 14px">' + inner + '</td></tr>',
      'background:#FFFFFF;border:1px solid ' + THEME.line + ';border-radius:16px;border-collapse:separate') + '</td></tr>';
}

function eHeading_(text, top) {
  return '<p style="' + FONT + 'margin:' + (top === undefined ? 0 : top) + 'px 0 4px;font-size:13px;font-weight:600;letter-spacing:1px;' +
    'text-transform:uppercase;color:' + THEME.mid + '">' + esc_(text) + '</p>';
}

function eCaption_(text) {
  return '<p style="' + FONT + 'margin:0 0 10px;font-size:12px;line-height:1.4;color:' + THEME.muted + '">' + esc_(text) + '</p>';
}

function eSubhead_(text, top) {
  return '<p style="' + FONT + 'margin:' + (top === undefined ? 10 : top) + 'px 0 4px;font-size:14px;font-weight:600;color:' + THEME.ink + '">' + esc_(text) + '</p>';
}

/**
 * One of two side-by-side slots in a card (eRow_). Each is up to 304px wide
 * with 8px of padding on both sides, so two fit across the 608px inside a
 * card with a 16px gap, and one alone sits centered on a phone.
 */
function eSlot_(html, style) {
  return '<div style="display:inline-block;vertical-align:top;width:100%;max-width:304px;text-align:left;font-size:14px;line-height:1.4;' + (style || '') + '">' +
    eTable_('<tr><td style="padding:0 8px">' + html + '</td></tr>') + '</div>';
}

/** A row of slots: side by side when there's room, stacked and centered when there isn't. */
function eRow_(slots) {
  return '<div style="font-size:0;line-height:0;text-align:center;margin:0 -8px">' + slots + '</div>';
}

/** A horizontal bar, n out of max, on a pale track. */
function eBar_(n, max, color) {
  const w = n ? Math.max(2, Math.round(n * 100 / Math.max(max, 1))) : 0;
  return eTable_('<tr>' +
    (w ? '<td width="' + w + '%" style="height:10px;background:' + color + ';font-size:0;line-height:0">&nbsp;</td>' : '') +
    (w < 100 ? '<td width="' + (100 - w) + '%" style="height:10px;background:' + THEME.track + ';font-size:0;line-height:0">&nbsp;</td>' : '') +
    '</tr>', 'border-radius:5px;overflow:hidden;border-collapse:separate');
}

/** Bar rows: label | bar | number. Blank areas read "not reported". */
function eBars_(items, color) {
  const max = Math.max.apply(null, items.map(function (a) { return a.n || 0; }).concat([1]));
  return eTable_(items.map(function (a) {
    const parts = splitName_(a.name);
    return '<tr><td style="' + FONT + 'width:44%;padding:6px 8px 6px 0;font-size:14px;line-height:1.25;color:' + THEME.ink + ';border-bottom:1px solid ' + THEME.line + '">' +
      esc_(parts.name) + (parts.detail ? '<br><span style="font-size:11px;color:' + THEME.muted + '">' + esc_(parts.detail) + '</span>' : '') + '</td>' +
      '<td style="padding:6px 0;border-bottom:1px solid ' + THEME.line + '">' +
      (a.n === null ? '<span style="' + FONT + 'font-size:11px;font-style:italic;color:' + THEME.muted + '">not reported</span>' : eBar_(a.n, max, color)) + '</td>' +
      '<td style="' + FONT + 'width:40px;padding:6px 0 6px 8px;font-size:14px;font-weight:600;text-align:right;color:' + THEME.ink + ';border-bottom:1px solid ' + THEME.line + '">' +
      (a.n === null ? '' : a.n) + '</td></tr>';
  }).join(''));
}

/** A quick box: label, big number, a line under it. */
function eBox_(label, value, sub, dark) {
  const T = THEME;
  return eSlot_(eTable_('<tr><td style="' + FONT + 'padding:12px 14px;line-height:1.3;background:' + (dark ? T.deep : '#FFFFFF') +
    ';border:1px solid ' + (dark ? T.deep : T.line) + ';border-radius:12px;height:84px;vertical-align:top">' +
    '<p style="margin:0;font-size:11px;letter-spacing:.6px;text-transform:uppercase;color:' + (dark ? '#CFE3C0' : T.muted) + '">' + esc_(label) + '</p>' +
    '<p style="margin:3px 0 0;font-size:30px;line-height:1.1;font-weight:700;color:' + (dark ? '#FFFFFF' : T.ink) + '">' + value + '</p>' +
    '<p style="margin:3px 0 0;font-size:12px;color:' + (dark ? '#CFE3C0' : T.muted) + '">' + sub + '</p></td></tr>', 'border-collapse:separate'), 'margin-top:10px');
}

/** The auditorium as a floor plan, when its areas are the usual five. */
function eFloorPlan_(section) {
  const T = THEME;
  const n = {};
  section.areas.forEach(function (a) { n[a.name] = a.n; });
  const names = ['North Wing', 'North Main', 'South Main', 'South Wing', 'Platform'];
  if (!names.every(function (k) { return k in n; })) return '';
  const total = section.subtotal;
  const block = function (name, bg) {
    const v = n[name];
    return '<td width="25%" style="' + FONT + 'padding:12px 4px;background:' + bg + ';text-align:center;border:3px solid #FFFFFF;border-radius:10px">' +
      '<p style="margin:0;font-size:24px;line-height:1.1;font-weight:700;color:#FFFFFF">' + (v === null ? '—' : v) + '</p>' +
      '<p style="margin:2px 0 0;font-size:12px;line-height:1.2;color:#FFFFFF">' + esc_(name) + '</p>' +
      '<p style="margin:2px 0 0;font-size:11px;color:#CFE3C0">' + (v === null ? 'not reported' : pct_(v, total)) + '</p></td>';
  };
  const plat = n['Platform'];
  return eTable_('<tr><td width="20%"></td><td width="60%" style="' + FONT + 'padding:6px;background:' + T.soft + ';border:1px solid ' + T.brand +
      ';border-radius:10px;text-align:center"><span style="font-size:11px;letter-spacing:.6px;color:' + T.muted + '">PLATFORM</span> ' +
      '<b style="font-size:16px;color:' + T.ink + '">' + (plat === null ? '—' : plat) + '</b>' +
      (plat === null ? '' : ' <span style="font-size:11px;color:' + T.muted + '">' + pct_(plat, total) + '</span>') + '</td><td width="20%"></td></tr>',
      'border-collapse:separate') +
    eTable_('<tr>' + block('North Wing', T.mid) + block('North Main', T.deep) + block('South Main', T.deep) + block('South Wing', T.mid) + '</tr>',
      'margin-top:6px;border-collapse:separate;table-layout:fixed');
}

/**
 * A whole email: the dark green header bar with the logo, kicker (small
 * capitals), title and subtitle, then cards (eCard_ rows), then footer.
 * preview is the line email apps show beside the subject.
 */
function ePage_(preview, kicker, title, subtitle, cards, footer) {
  const T = THEME;
  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">' +
    '<meta name="color-scheme" content="light"><meta name="supported-color-schemes" content="light">' +
    '<link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600;700&display=swap" rel="stylesheet"></head>' +
    '<body style="margin:0;padding:0;background:' + T.soft + '">' +
    '<div style="display:none;max-height:0;overflow:hidden;mso-hide:all">' + esc_(preview) + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" bgcolor="' + T.soft + '" style="background:' + T.soft + '">' +
    '<tr><td align="center" style="padding:16px 6px 24px">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width:640px;border-collapse:collapse">\n' +
    // Header bar, like the top of the web page
    '<tr><td bgcolor="' + T.deep + '" align="center" style="background:' + T.deep + ';border-radius:16px;padding:20px 18px 18px">' +
    '<img src="' + LOGO_URL + '" width="149" height="52" alt="Grace Baptist Church" style="display:block;margin:0 auto 10px;border:0;height:52px;width:149px">' +
    '<p style="' + FONT + 'margin:0;font-size:12px;letter-spacing:1.5px;text-transform:uppercase;color:#CFE3C0">' + esc_(kicker) + '</p>' +
    '<p style="' + FONT + 'margin:2px 0 0;font-size:22px;font-weight:700;color:#FFFFFF">' + esc_(title) + '</p>' +
    '<p style="' + FONT + 'margin:2px 0 0;font-size:14px;color:#CFE3C0">' + esc_(subtitle) + '</p>' +
    '</td></tr>\n' + cards + '\n' +
    '<tr><td style="' + FONT + 'padding:16px 8px 0;text-align:center;font-size:11px;color:#999999">' + esc_(footer) + '</td></tr>' +
    '</table></td></tr></table></body></html>';
}

/** The report, as the email's HTML body. */
function reportHtml_(d, updated) {
  const T = THEME;
  const svc = serviceName_(d);
  const prev = d.history.length > 1 ? d.history[d.history.length - 2] : null;
  const change = prev ? d.total - prev.total : null;
  const changeNote = prev
    ? (change === 0 ? 'Same as' : (change > 0 ? 'Up ' + change + ' from' : 'Down ' + (-change) + ' from')) + ' ' + fmtDate_(prev.date, d.tz, 'M/d')
    : 'All areas';
  // Up green, down amber, as in the offering report.
  const signedHtml = function (n) {
    return '<span style="color:' + (n < 0 ? '#8A5200' : '#3D7A28') + ';font-weight:600">' + (n > 0 ? '+' : n < 0 ? '−' : '±') + Math.abs(n) + '</span>';
  };
  const cmp = d.compare || {};
  const ly = cmp.lastYear;
  const out = [];
  const preview = d.total + ' at ' + svc + ' · ' + d.sections.map(function (s) { return s.title + ' ' + s.subtotal; }).join(' · ');


  // Quick boxes: the total, then one per section, two across
  const boxes = [eBox_('Total attendance', String(d.total), esc_(changeNote) +
    (ly ? '<br>' + (ly.change > 0 ? '+' : ly.change < 0 ? '−' : '±') + Math.abs(ly.change) + ' vs. last year (' + ly.total + ')' : ''), true)].concat(d.sections.map(function (s) {
    return eBox_(s.title, String(s.subtotal), pct_(s.subtotal, d.total) ? pct_(s.subtotal, d.total) + ' of total' : '&nbsp;');
  }));
  let quick = '';
  for (let i = 0; i < boxes.length; i += 2) quick += eRow_(boxes[i] + (boxes[i + 1] || ''));

  // Where everyone was: one stacked bar and its legend
  let where = '';
  if (d.total) {
    const cells = [], legend = [];
    // Widths in whole percents; the last segment takes what's left so they add up to 100.
    const shown = d.sections.filter(function (s) { return s.subtotal; });
    let used = 0;
    d.sections.forEach(function (s, i) {
      const color = SECTION_COLORS[i % SECTION_COLORS.length];
      if (s.subtotal) {
        const last = s === shown[shown.length - 1];
        const share = last ? 100 - used : Math.max(1, Math.round(s.subtotal * 100 / d.total));
        used += share;
        cells.push('<td width="' + share + '%" style="' + FONT + 'height:30px;background:' + color.bg + ';color:' + color.ink +
          ';text-align:center;font-size:13px;font-weight:600">' + (share >= 8 ? s.subtotal : '&nbsp;') + '</td>');
      }
      legend.push('<span style="white-space:nowrap;margin-right:14px"><span style="display:inline-block;width:10px;height:10px;border-radius:3px;background:' +
        color.bg + ';margin-right:5px"></span><b>' + esc_(s.title) + '</b> ' + s.subtotal + ' · ' + pct_(s.subtotal, d.total) + '</span> ');
    });
    where = eHeading_('Where everyone was', 18) + eCaption_('Share of the ' + d.total + ' people counted, by area of the building.') +
      eTable_('<tr>' + cells.join('') + '</tr>', 'border-radius:8px;overflow:hidden;border-collapse:separate') +
      '<p style="' + FONT + 'margin:8px 0 0;font-size:13px;line-height:1.8;color:' + T.ink + '">' + legend.join('') + '</p>';
  }
  out.push(eCard_(quick + where));

  // Auditorium floor plan (or bars if its areas were renamed)
  const aud = d.sections[0];
  out.push(eCard_(eHeading_(aud.title + ' by section') + eCaption_(aud.subtotal + ' in the ' + aud.title.toLowerCase() + '.') +
    (eFloorPlan_(aud) || eBars_(aud.areas, SECTION_COLORS[0].bg))));

  // The other sections, two across
  const rest = d.sections.slice(1);
  if (rest.length) {
    let html = '';
    for (let i = 0; i < rest.length; i += 2) {
      const col = function (s, idx) {
        if (!s) return '';
        return eSlot_(eSubhead_(s.title, 4) + eCaption_(s.subtotal + ' counted.') + eBars_(s.areas, SECTION_COLORS[(idx + 1) % SECTION_COLORS.length].bg), 'margin-bottom:12px');
      };
      html += eRow_(col(rest[i], i) + col(rest[i + 1], i + 1));
    }
    out.push(eCard_(eHeading_('Other areas') + html));
  }

  // Full count
  const cell = function (html, st) {
    return '<td style="' + FONT + 'padding:6px 5px;font-size:13px;color:' + T.ink + ';border-bottom:1px solid ' + T.line + ';' + (st || '') + '">' + html + '</td>';
  };
  const head = '<tr>' + ['Area', 'Count', 'Share'].map(function (t, i) {
    return '<td style="' + FONT + 'padding:4px 5px;font-size:11px;letter-spacing:.5px;text-transform:uppercase;color:' + T.muted +
      ';border-bottom:2px solid ' + T.brand + (i ? ';text-align:right' : '') + '">' + t + '</td>';
  }).join('') + '</tr>';
  const rows = [head];
  d.sections.forEach(function (s) {
    s.areas.forEach(function (a) {
      const parts = splitName_(a.name);
      rows.push('<tr>' + cell(esc_(parts.name) + (parts.detail ? ' <span style="font-size:11px;color:' + T.muted + '">(' + esc_(parts.detail) + ')</span>' : '')) +
        cell(a.n === null ? DASH : String(a.n), 'text-align:right') + cell(a.n === null ? '' : pct_(a.n, d.total), 'text-align:right;color:' + T.muted) + '</tr>');
    });
    const sub = 'background:' + T.soft + ';font-weight:600;';
    rows.push('<tr>' + cell(esc_(s.title) + ' subtotal', sub) + cell(String(s.subtotal), sub + 'text-align:right') +
      cell(pct_(s.subtotal, d.total), sub + 'text-align:right') + '</tr>');
  });
  const grand = 'background:' + T.deep + ';color:#FFFFFF;font-weight:600;font-size:14px;';
  rows.push('<tr>' + cell('Total attendance', grand) + cell(String(d.total), grand + 'text-align:right') + cell(d.total ? '100%' : '', grand + 'text-align:right') + '</tr>');
  const full = eSubhead_('Full count', 4) + eCaption_('Every area as entered. A dash means it was left blank.') + eTable_(rows.join(''));

  // Recent services of the same kind, and others the same day
  const recent = [eSubhead_('Recent ' + svc + ' services', 4)];
  if (d.history.length > 1) {
    recent.push(eCaption_('This service and the ones before it.'));
    const max = Math.max.apply(null, d.history.map(function (e) { return e.total; }).concat([1]));
    recent.push(eTable_(d.history.slice().reverse().map(function (e) {
      const b = e.current ? 'font-weight:700;' : '';
      return '<tr><td style="' + FONT + 'width:52px;padding:5px 8px 5px 0;font-size:13px;white-space:nowrap;color:' + T.ink + ';' + b + '">' +
        fmtDate_(e.date, d.tz, e.date.getFullYear() === d.date.getFullYear() ? 'MMM d' : 'MMM d, yyyy') + '</td>' +
        '<td style="padding:5px 0">' + eBar_(e.total, max, e.current ? T.deep : '#A8C496') + '</td>' +
        '<td style="' + FONT + 'width:44px;padding:5px 0 5px 8px;font-size:13px;text-align:right;color:' + T.ink + ';' + b + '">' + e.total + '</td></tr>';
    }).join('')));
  } else {
    recent.push(eCaption_('This is the first ' + svc + ' service recorded. Comparisons will appear here as more are logged.'));
  }
  // Compared with last year
  if (ly || cmp.average !== null && cmp.lastAverage !== null) {
    const rowsCmp = [];
    if (ly) rowsCmp.push(['Same week last year', ly.total + ' <span style="font-size:11px;color:' + T.muted + '">' + fmtDate_(ly.date, d.tz, 'MMM d, yyyy') + '</span>', signedHtml(ly.change)]);
    if (cmp.average !== null) rowsCmp.push([cmp.year + ' average so far', cmp.average + ' <span style="font-size:11px;color:' + T.muted + '">' + cmp.count + (cmp.count === 1 ? ' service' : ' services') + '</span>', '']);
    if (cmp.lastAverage !== null) rowsCmp.push([(cmp.year - 1) + ' average, same weeks', String(cmp.lastAverage),
      cmp.average !== null ? signedHtml(cmp.average - cmp.lastAverage) : '']);
    recent.push(eSubhead_('Compared with last year', 14));
    recent.push(eTable_(rowsCmp.map(function (r) {
      return '<tr>' + cell(r[0]) + cell(r[1], 'text-align:right;white-space:nowrap') + cell(r[2], 'text-align:right;width:44px') + '</tr>';
    }).join('')));
  }
  if (d.sameDay.length) {
    recent.push(eSubhead_('Same day', 14));
    recent.push(eTable_(d.sameDay.map(function (e) {
      return '<tr>' + cell(esc_(e.name)) + cell('<b>' + e.total + '</b>', 'text-align:right') + '</tr>';
    }).join('')));
  }
  out.push(eCard_(eRow_(eSlot_(full, 'margin-bottom:12px') + eSlot_(recent.join(''), 'margin-bottom:12px'))));

  // Service details: bulletins, door ushers, who counted, notes
  const doors = eTable_('<tr>' + ['Door', 'Usher'].map(function (t) {
    return '<td style="' + FONT + 'padding:4px 5px;font-size:11px;letter-spacing:.5px;text-transform:uppercase;color:' + T.muted + ';border-bottom:2px solid ' + T.brand + '">' + t + '</td>';
  }).join('') + '</tr>' + d.doors.map(function (x) {
    return '<tr>' + cell('Door ' + x.n, 'width:70px') + cell(x.name ? esc_(x.name) : DASH) + '</tr>';
  }).join(''));
  const by = d.submittedBy ? 'Counted by ' + d.submittedBy : 'Counted';
  const at = d.submittedAt instanceof Date ? ' on ' + fmtDate_(d.submittedAt, d.tz, 'M/d/yyyy') + ' at ' + fmtDate_(d.submittedAt, d.tz, 'h:mm a') : '';
  const left = eSubhead_('Bulletins left over', 4) +
    '<p style="' + FONT + 'margin:0;font-size:30px;font-weight:700;color:' + T.ink + '">' + (d.bulletins === null ? '—' : d.bulletins) + '</p>' +
    '<p style="' + FONT + 'margin:10px 0 0;font-size:13px;color:' + T.muted + '">' + esc_(by + at) + '.</p>' +
    (d.notes ? eSubhead_('Notes', 14) + '<p style="' + FONT + 'margin:0;font-size:14px;line-height:1.45;color:' + T.ink + '">' + esc_(d.notes) + '</p>' : '');
  out.push(eCard_(eHeading_('Service details') + eRow_(eSlot_(left, 'margin-bottom:12px') + eSlot_(eSubhead_('Door ushers', 4) + doors, 'margin-bottom:12px'))));

  return ePage_(preview, updated ? 'Updated attendance' : 'Attendance', svc, fmtDate_(d.date, d.tz, 'EEEE, MMMM d, yyyy'),
    out.join('\n'), 'Sent automatically by the Grace usher attendance page.');
}

// ---- Plain-text email -----------------------------------------------------------
//
// For the few email apps that don't show HTML.

function emailText_(d, updated) {
  const lines = [(updated ? 'Updated attendance for ' : 'Attendance for ') + serviceName_(d) + ', ' +
    fmtDate_(d.date, d.tz, 'EEEE, MMMM d') + ': ' + d.total, ''];
  d.sections.forEach(function (s) {
    lines.push(s.title + ': ' + s.subtotal);
    s.areas.forEach(function (a) { lines.push('  ' + a.name + ': ' + (a.n === null ? '—' : a.n)); });
  });
  const c = d.compare;
  if (c && c.lastYear) lines.push('', 'Same week last year: ' + c.lastYear.total + ' (' + fmtDate_(c.lastYear.date, d.tz, 'MMM d, yyyy') + '), ' +
    (c.lastYear.change >= 0 ? '+' : '') + c.lastYear.change);
  if (c && c.average !== null) lines.push(c.year + ' average so far: ' + c.average + (c.lastAverage !== null ? '; ' + (c.year - 1) + ' same weeks: ' + c.lastAverage : ''));
  if (d.bulletins !== null) lines.push('', 'Bulletins left over: ' + d.bulletins);
  if (d.submittedBy) lines.push('Counted by ' + d.submittedBy);
  if (d.notes) lines.push('', 'Notes: ' + d.notes);
  return lines.join('\n');
}
