/**
 * Grace Baptist Church - Usher Attendance: emailed PDF report
 *
 * After each submission, Code.gs marks the row "Pending" in the Report
 * Emailed column and schedules sendPendingReports a few seconds later, so the
 * usher's phone doesn't wait on the PDF. sendPendingReports builds a two-page
 * report, turns it into a PDF by way of a temporary Google Doc, and emails it
 * to the addresses on the Report Emails tab.
 *
 * The report is HTML made of tables only: the Google Docs import that turns
 * it into a PDF ignores flexbox, grid and SVG.
 */

const EMAILS_NAME = 'Report Emails';
const REPORT_STATUS = 'Report Emailed';
const REPORT_LOGO_URL = 'https://davidluttrull.github.io/grace-attendance/report-logo.png';
const HISTORY_COUNT = 8;

const C_DEEP = '#4E7A34';
const C_BRAND = '#7BAE59';
const C_TINT = '#E3EFD9';
const C_DARK = '#3A3A3A';
const C_MUTED = '#6B6B6B';
const C_RULE = '#E6E6E6';
const C_TRACK = '#F0F0F0';
const SECTION_COLORS = [C_DEEP, C_BRAND, C_DARK, '#9A9A9A', '#B8CFA6'];

/** Adds an Attendance menu to the Sheet. */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Attendance')
    .addItem('Email report for selected row', 'emailSelectedReport')
    .addItem('Send me a test report for selected row', 'emailSelectedReportToMe')
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
    const pdf = reportPdf_(data);
    const label = serviceName_(data) + ', ' + Utilities.formatDate(data.date, tz, 'MMM d');
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: (updated ? 'Updated attendance: ' : 'Attendance: ') + label + ' (' + data.total + ')',
      htmlBody: emailHtml_(data, updated),
      body: emailText_(data, updated),
      name: 'Grace Attendance',
      attachments: [pdf],
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
    tab.getRange(1, 2).setFontColor(C_MUTED).setFontStyle('italic');
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
  data.history.sort(function (a, b) { return a.date - b.date; });
  data.history = data.history.slice(-HISTORY_COUNT);
  return data;
}

function serviceName_(data) {
  return data.service === 'Other' ? data.otherName : data.service;
}

// ---- PDF --------------------------------------------------------------------

/**
 * HTML -> temporary Google Doc (Drive converts it) -> PDF. The Doc is trashed afterwards.
 * The conversion needs the Drive API advanced service (Services + > Drive API in the editor).
 */
function reportPdf_(data) {
  if (typeof Drive === 'undefined') {
    throw new Error('The Drive API service is off. In Apps Script, click + next to Services, ' +
      'choose Drive API, click Add, then try again.');
  }
  const html = Utilities.newBlob(reportHtml_(data), 'text/html', 'report.html');
  const title = 'Attendance report (temporary)';
  // Drive API v3 (the editor's default) has create; v2 has insert.
  const created = Drive.Files.create
    ? Drive.Files.create({ name: title, mimeType: MimeType.GOOGLE_DOCS }, html)
    : Drive.Files.insert({ title: title, mimeType: MimeType.GOOGLE_DOCS }, html, { convert: true });
  const file = DriveApp.getFileById(created.id);
  try {
    finishDoc_(DocumentApp.openById(created.id));
    const name = 'Attendance ' + Utilities.formatDate(data.date, data.tz, 'yyyy-MM-dd') + ' ' + serviceName_(data) + '.pdf';
    return file.getAs('application/pdf').setName(name);
  } finally {
    file.setTrashed(true);
  }
}

/** Page margins, the logo and the page break, which the HTML import can't carry. */
function finishDoc_(doc) {
  const body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(43).setMarginRight(43);

  const logo = takePlaceholder_(body, 'LOGO');
  if (logo) {
    try {
      const img = logo.insertInlineImage(0, UrlFetchApp.fetch(REPORT_LOGO_URL).getBlob());
      const h = 52;
      img.setWidth(Math.round(img.getWidth() * h / img.getHeight())).setHeight(h);
    } catch (e) {
      logo.appendText('GRACE BAPTIST CHURCH').setBold(true).setFontSize(16);
    }
  }
  const pageBreak = takePlaceholder_(body, 'PAGEBREAK');
  if (pageBreak) pageBreak.appendPageBreak();
  doc.saveAndClose();
}

/** Removes a [[NAME]] marker and returns the paragraph it was in. */
function takePlaceholder_(body, name) {
  const found = body.findText('\\[\\[' + name + '\\]\\]');
  if (!found) return null;
  const text = found.getElement().asText();
  text.deleteText(found.getStartOffset(), found.getEndOffsetInclusive());
  return text.getParent().asParagraph();
}

// ---- Report HTML ------------------------------------------------------------

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

/** A table cell. Docs draws a border on every cell unless told not to. */
function td_(content, style) {
  return '<td style="border:none;padding:0;vertical-align:top;' + (style || '') + '">' + content + '</td>';
}

function table_(rows, width) {
  return '<table style="border-collapse:collapse;width:' + width + 'pt">' + rows + '</table>';
}

function heading_(text, top) {
  return '<p style="margin:' + (top === undefined ? 14 : top) + 'pt 0 2pt;font-size:11pt;font-weight:bold;' +
    'color:' + C_DEEP + ';letter-spacing:1px">' + esc_(text.toUpperCase()) + '</p>';
}

function caption_(text) {
  return '<p style="margin:0 0 6pt;font-size:8.5pt;color:' + C_MUTED + '">' + esc_(text) + '</p>';
}

/**
 * A bar drawn as shaded non-breaking spaces. Nested tables would be neater
 * HTML, but Docs pads each one with an empty line.
 */
function barSpan_(n, max, color, width) {
  const SPACE = 2.74; // a non-breaking space at 8pt, as measured in the Docs PDF
  const chars = Math.max(1, Math.round(width / SPACE));
  const span = function (bg, count, text) {
    return '<span style="font-size:8pt;background-color:' + bg + '">' + (text || '') +
      new Array(Math.max(0, count) + 1).join('&nbsp;') + '</span>';
  };
  if (n === null) {
    return '<span style="font-size:8pt;font-style:italic;color:' + C_MUTED + ';background-color:' + C_TRACK + '">&nbsp;not&nbsp;reported</span>' +
      span(C_TRACK, chars - 20);
  }
  const fill = n ? Math.max(1, Math.round(chars * n / Math.max(max, 1))) : 0;
  return span(color, fill) + span(C_TRACK, chars - fill);
}

/** Horizontal bar rows: label | bar | number. Blank areas read "not reported". */
function bars_(items, color, width) {
  const labelW = Math.round(width * 0.42), numW = 30, barW = width - labelW - numW;
  const max = Math.max.apply(null, items.map(function (a) { return a.n || 0; }).concat([1]));
  return table_(items.map(function (a) {
    const parts = splitName_(a.name);
    const label = '<p style="margin:0;font-size:9.5pt;font-weight:bold">' + esc_(parts.name) + '</p>' +
      (parts.detail ? '<p style="margin:0;font-size:7.5pt;color:' + C_MUTED + '">' + esc_(parts.detail) + '</p>' : '');
    return '<tr>' + td_(label, 'width:' + labelW + 'pt;padding:3pt 4pt 3pt 0') +
      td_('<p style="margin:0">' + barSpan_(a.n, max, color, barW - 20) + '</p>', 'width:' + barW + 'pt;padding:5pt 0 3pt') +
      td_(a.n === null ? '' : '<b>' + a.n + '</b>', 'width:' + numW + 'pt;text-align:right;padding:3pt 0;font-size:10pt') +
      '</tr>';
  }).join(''), width);
}

/** The auditorium as a floor plan, when its areas are the usual five. */
function auditoriumPlan_(section) {
  const n = {};
  section.areas.forEach(function (a) { n[a.name] = a.n; });
  const names = ['North Wing', 'North Main', 'South Main', 'South Wing', 'Platform'];
  if (!names.every(function (k) { return k in n; })) return '';
  const total = section.subtotal;
  const block = function (name, bg, w) {
    const v = n[name];
    return td_('<p style="margin:0;font-size:22pt;font-weight:bold;color:#ffffff">' + (v === null ? '—' : v) + '</p>' +
      '<p style="margin:0;font-size:9.5pt;color:#ffffff">' + esc_(name) + '</p>' +
      '<p style="margin:0;font-size:7.5pt;color:' + C_TINT + '">' + (v === null ? 'not reported' : pct_(v, total) + ' of auditorium') + '</p>',
      'width:' + w + 'pt;background:' + bg + ';text-align:center;padding:12pt 4pt');
  };
  const gap = td_('', 'width:8pt');
  const plat = n['Platform'];
  return table_('<tr>' + td_('', 'width:131pt') + td_('', 'width:8pt') +
      td_('<p style="margin:0;font-size:7.5pt;color:' + C_MUTED + '">PLATFORM</p><p style="margin:0;font-weight:bold">' +
        (plat === null ? '—' : plat) + '</p>', 'width:248pt;background:' + C_TINT + ';text-align:center;padding:4pt;' +
        'border:1pt solid ' + C_BRAND) +
      td_('', 'width:8pt') + td_('', 'width:131pt') + '</tr>', 526) +
    '<p style="margin:0;font-size:4pt">&nbsp;</p>' +
    table_('<tr>' + block('North Wing', '#61894A', 115) + gap + block('North Main', C_DEEP, 124) + gap +
      block('South Main', C_DEEP, 124) + gap + block('South Wing', '#61894A', 115) + '</tr>', 526);
}

function fmtDate_(d, tz, pattern) {
  return d instanceof Date ? Utilities.formatDate(d, tz, pattern) : '';
}

/** The two-page report. Letter page, 526pt between the margins set in finishDoc_. */
function reportHtml_(d) {
  const W = 526;
  const svc = serviceName_(d);
  const longDate = fmtDate_(d.date, d.tz, 'EEEE, MMMM d, yyyy');
  const shortDate = fmtDate_(d.date, d.tz, 'MMMM d, yyyy');
  const prev = d.history.length > 1 ? d.history[d.history.length - 2] : null;
  const h = [];

  // Body padding becomes the page margins on import (finishDoc_ sets them again to be sure).
  h.push('<html><head><meta charset="utf-8"></head>' +
    '<body style="max-width:526pt;padding:36pt 43pt;font-family:Arial;font-size:10pt;color:#1A1A1A">');

  // Header: logo | title
  h.push(table_('<tr>' + td_('<p style="margin:0">[[LOGO]]</p>', 'width:200pt;vertical-align:middle;padding-bottom:8pt;border-bottom:3pt solid ' + C_BRAND) +
    td_('<p style="margin:0;text-align:right;font-size:17pt;font-weight:bold">' + esc_(svc.toUpperCase()) + ' ATTENDANCE</p>' +
      '<p style="margin:0;text-align:right;color:' + C_MUTED + '">' + esc_(longDate) + '</p>',
      'width:326pt;vertical-align:middle;padding-bottom:8pt;border-bottom:3pt solid ' + C_BRAND) + '</tr>', W));

  // KPI tiles: total, then one per section
  const tiles = [];
  const tileW = Math.floor((W - 8 * d.sections.length) / (d.sections.length + 1.3));
  const firstW = W - 8 * d.sections.length - tileW * d.sections.length;
  const change = prev ? d.total - prev.total : null;
  const totalNote = prev
    ? (change === 0 ? 'Same as' : (change > 0 ? 'Up ' + change + ' from' : 'Down ' + (-change) + ' from')) +
      ' ' + fmtDate_(prev.date, d.tz, 'M/d')
    : 'All areas';
  tiles.push(td_('<p style="margin:0;font-size:7.5pt;color:' + C_TINT + '">TOTAL ATTENDANCE</p>' +
    '<p style="margin:0;font-size:28pt;font-weight:bold;color:#ffffff">' + d.total + '</p>' +
    '<p style="margin:0;font-size:8pt;color:' + C_TINT + '">' + esc_(totalNote) + '</p>',
    'width:' + firstW + 'pt;background:' + C_DEEP + ';padding:8pt 10pt'));
  d.sections.forEach(function (s) {
    tiles.push(td_('', 'width:8pt'));
    tiles.push(td_('<p style="margin:0;font-size:7.5pt;color:' + C_MUTED + '">' + esc_(s.title.toUpperCase()) + '</p>' +
      '<p style="margin:0;font-size:28pt;font-weight:bold">' + s.subtotal + '</p>' +
      '<p style="margin:0;font-size:8pt;color:' + C_MUTED + '">' + (pct_(s.subtotal, d.total) ? pct_(s.subtotal, d.total) + ' of total' : '&nbsp;') + '</p>',
      'width:' + tileW + 'pt;padding:8pt 10pt;border:1pt solid #D9D9D9'));
  });
  h.push('<p style="margin:0;font-size:6pt">&nbsp;</p>');
  h.push(table_('<tr>' + tiles.join('') + '</tr>', W));

  // Where everyone was: one stacked bar
  h.push(heading_('Where everyone was'));
  h.push(caption_('Share of the ' + d.total + ' people counted, by area of the building.'));
  if (d.total) {
    const cells = [], legend = [];
    d.sections.forEach(function (s, i) {
      const color = SECTION_COLORS[i % SECTION_COLORS.length];
      if (s.subtotal) {
        const share = s.subtotal / d.total;
        cells.push(td_(share >= 0.07 ? '<b>' + s.subtotal + '</b>' : '&nbsp;',
          'width:' + Math.max(1, Math.round(W * share)) + 'pt;background:' + color + ';color:#ffffff;text-align:center;padding:6pt 0'));
      }
      legend.push('<span style="color:' + color + '">&#9632;</span> <b>' + esc_(s.title) + '</b> ' + s.subtotal +
        ' &middot; ' + pct_(s.subtotal, d.total));
    });
    h.push(table_('<tr>' + cells.join('') + '</tr>', W));
    h.push('<p style="margin:4pt 0 0;font-size:8.5pt">' + legend.join('&nbsp;&nbsp;&nbsp;&nbsp;') + '</p>');
  }

  // Auditorium floor plan (or bars if its areas were renamed)
  const aud = d.sections[0];
  const plan = auditoriumPlan_(aud);
  h.push(heading_(aud.title + ' by section'));
  h.push(caption_(aud.subtotal + ' in the ' + aud.title.toLowerCase() + '.'));
  h.push(plan || bars_(aud.areas, SECTION_COLORS[0], W));

  // Remaining sections side by side, two per row
  const rest = d.sections.slice(1);
  for (let i = 0; i < rest.length; i += 2) {
    const col = function (s, idx) {
      if (!s) return td_('', 'width:255pt');
      return td_(heading_(s.title, 12) + caption_(s.subtotal + ' counted.') +
        bars_(s.areas, SECTION_COLORS[(idx + 1) % SECTION_COLORS.length], 255), 'width:255pt');
    };
    h.push(table_('<tr>' + col(rest[i], i) + td_('', 'width:16pt') + col(rest[i + 1], i + 1) + '</tr>', W));
  }

  // ---- Page 2
  // The HTML import ignores page-break styles, so finishDoc_ puts the break here.
  h.push('<p style="margin:0;font-size:1pt">[[PAGEBREAK]]</p>');
  h.push(table_('<tr>' + td_('<b>Grace Baptist Church</b> &middot; ' + esc_(svc) + ' Attendance',
    'width:340pt;font-size:8pt;color:' + C_MUTED + ';padding-bottom:4pt;border-bottom:1.5pt solid ' + C_BRAND) +
    td_(esc_(shortDate), 'width:186pt;text-align:right;font-size:8pt;color:' + C_MUTED + ';padding-bottom:4pt;border-bottom:1.5pt solid ' + C_BRAND) +
    '</tr>', W));

  // Full count table
  // Docs sizes columns from the first row, so it carries the widths.
  const th = function (t, w, right) {
    return td_(t, 'width:' + w + 'pt;font-size:7.5pt;color:' + C_MUTED + ';padding:4pt 4pt;border-bottom:1.5pt solid ' + C_BRAND +
      (right ? ';text-align:right' : ''));
  };
  const rowsHtml = ['<tr>' + th('AREA', 152) + th('COUNT', 48, true) + th('SHARE', 48, true) + '</tr>'];
  const line = 'border-bottom:0.5pt solid ' + C_RULE + ';padding:3pt 4pt;font-size:9pt';
  d.sections.forEach(function (s) {
    s.areas.forEach(function (a) {
      const parts = splitName_(a.name);
      rowsHtml.push('<tr>' +
        td_(esc_(parts.name) + (parts.detail ? ' <span style="font-size:7.5pt;color:' + C_MUTED + '">(' + esc_(parts.detail) + ')</span>' : ''), line) +
        td_(a.n === null ? '<span style="color:#AAAAAA">—</span>' : String(a.n), line + ';text-align:right') +
        td_(a.n === null ? '' : pct_(a.n, d.total), line + ';text-align:right') + '</tr>');
    });
    const sub = 'background:' + C_TINT + ';font-weight:bold;padding:3pt 4pt;font-size:9pt';
    rowsHtml.push('<tr>' + td_(esc_(s.title) + ' subtotal', sub) + td_(String(s.subtotal), sub + ';text-align:right') +
      td_(pct_(s.subtotal, d.total), sub + ';text-align:right') + '</tr>');
  });
  const grand = 'background:' + C_DEEP + ';color:#ffffff;font-weight:bold;padding:4pt 4pt;font-size:10pt';
  rowsHtml.push('<tr>' + td_('Total attendance', grand) + td_(String(d.total), grand + ';text-align:right') +
    td_(d.total ? '100%' : '', grand + ';text-align:right') + '</tr>');
  const left = heading_('Full count', 12) + caption_('Every area as entered. A dash means it was left blank.') +
    table_(rowsHtml.join(''), 248);

  // Recent services of the same kind
  const right = [];
  right.push(heading_('Recent ' + svc + ' services', 12));
  if (d.history.length > 1) {
    right.push(caption_('This service and the ones before it.'));
    const max = Math.max.apply(null, d.history.map(function (e) { return e.total; }).concat([1]));
    right.push(table_(d.history.slice().reverse().map(function (e) {
      const b = e.current ? function (x) { return '<b>' + x + '</b>'; } : function (x) { return x; };
      return '<tr>' + td_(b(fmtDate_(e.date, d.tz, 'MMM d')), 'width:52pt;font-size:9pt;padding:2pt 0') +
        td_('<p style="margin:0">' + barSpan_(e.total, max, e.current ? C_DEEP : '#A8C496', 132) + '</p>', 'width:154pt;padding:3pt 0 2pt') +
        td_(b(e.total), 'width:40pt;text-align:right;font-size:9pt;padding:2pt 0') + '</tr>';
    }).join(''), 246));
  } else {
    right.push(caption_('This is the first ' + svc + ' service recorded. Comparisons will appear here as more are logged.'));
  }
  if (d.sameDay.length) {
    right.push(heading_('Same day', 12));
    right.push(table_(d.sameDay.map(function (e) {
      return '<tr>' + td_(esc_(e.name), 'width:200pt;font-size:9pt;padding:2pt 0') +
        td_('<b>' + e.total + '</b>', 'width:46pt;text-align:right;font-size:9pt;padding:2pt 0') + '</tr>';
    }).join(''), 246));
  }

  // Service details
  right.push(heading_('Service details', 14));
  right.push(table_('<tr>' + td_('<p style="margin:0;font-size:7.5pt;color:' + C_MUTED + '">BULLETINS LEFT OVER</p>' +
    '<p style="margin:0;font-size:16pt;font-weight:bold">' + (d.bulletins === null ? '—' : d.bulletins) + '</p>',
    'width:246pt;padding:6pt 8pt;border:1pt solid #D9D9D9') + '</tr>', 246));
  const doorRows = d.doors.map(function (x) {
    return '<tr>' + td_('Door ' + x.n, 'width:60pt;font-size:9pt;padding:2pt 4pt;border-bottom:0.5pt solid ' + C_RULE) +
      td_(x.name ? esc_(x.name) : '<span style="color:#AAAAAA">—</span>', 'width:186pt;font-size:9pt;padding:2pt 4pt;border-bottom:0.5pt solid ' + C_RULE) + '</tr>';
  });
  right.push('<p style="margin:0;font-size:6pt">&nbsp;</p>');
  right.push(table_('<tr>' + td_('DOOR', 'width:60pt;font-size:7.5pt;color:' + C_MUTED + ';padding:3pt 4pt;border-bottom:1.5pt solid ' + C_BRAND) +
    td_('USHER', 'width:186pt;font-size:7.5pt;color:' + C_MUTED + ';padding:3pt 4pt;border-bottom:1.5pt solid ' + C_BRAND) + '</tr>' +
    doorRows.join(''), 246));
  const by = d.submittedBy ? 'Counted by ' + d.submittedBy : 'Counted';
  const at = d.submittedAt instanceof Date ? ' on ' + fmtDate_(d.submittedAt, d.tz, 'M/d/yyyy') + ' at ' + fmtDate_(d.submittedAt, d.tz, 'h:mm a') : '';
  right.push('<p style="margin:8pt 0 0;font-size:8.5pt;color:' + C_MUTED + '">' + esc_(by + at) + '.</p>');
  if (d.notes) {
    right.push(heading_('Notes', 12));
    right.push('<p style="margin:0;font-size:9pt">' + esc_(d.notes) + '</p>');
  }

  h.push(table_('<tr>' + td_(left, 'width:262pt') + td_('', 'width:18pt') + td_(right.join(''), 'width:246pt') + '</tr>', W));
  h.push('<p style="margin:16pt 0 0;font-size:7.5pt;color:#999999;text-align:center">Sent automatically by the Grace usher attendance page.</p>');
  h.push('</body></html>');
  return h.join('\n');
}

// ---- Email body ---------------------------------------------------------------

function emailText_(d, updated) {
  const lines = [(updated ? 'Updated attendance for ' : 'Attendance for ') + serviceName_(d) + ', ' +
    fmtDate_(d.date, d.tz, 'EEEE, MMMM d') + ': ' + d.total];
  d.sections.forEach(function (s) { lines.push(s.title + ': ' + s.subtotal); });
  if (d.submittedBy) lines.push('Counted by ' + d.submittedBy);
  lines.push('', 'The full report is attached as a PDF.');
  return lines.join('\n');
}

function emailHtml_(d, updated) {
  const rows = d.sections.map(function (s) {
    return '<tr><td style="padding:2px 16px 2px 0">' + esc_(s.title) + '</td><td style="text-align:right"><b>' + s.subtotal + '</b></td></tr>';
  }).join('');
  return '<div style="font-family:Arial,sans-serif;font-size:14px;color:#1A1A1A">' +
    '<p style="margin:0 0 4px;color:' + C_MUTED + '">' + (updated ? 'Updated attendance' : 'Attendance') + ' &middot; ' +
    esc_(serviceName_(d)) + ' &middot; ' + esc_(fmtDate_(d.date, d.tz, 'EEEE, MMMM d')) + '</p>' +
    '<p style="margin:0 0 12px;font-size:32px;font-weight:bold;color:' + C_DEEP + '">' + d.total + '</p>' +
    '<table style="border-collapse:collapse;margin-bottom:12px">' + rows + '</table>' +
    (d.submittedBy ? '<p style="margin:0 0 12px;color:' + C_MUTED + '">Counted by ' + esc_(d.submittedBy) + '</p>' : '') +
    '<p style="margin:0">The full report is attached as a PDF.</p></div>';
}
