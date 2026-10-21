/**
 * Grace Baptist Church - Usher Attendance: the weekly report
 *
 * Every Saturday morning sendWeeklyReport emails the week just past (the
 * Sunday and the Wednesday after it) to the addresses on the Report Emails
 * tab: what the ushers counted at each service and what the office entered on
 * the office page (Office.gs), each beside the same week last year. setup
 * schedules it. The Weekly tab's Weekly Report column shows what happened.
 */

const WEEKLY_HOUR = 7;   // Saturday, in the script's time zone

/** Schedules sendWeeklyReport for Saturday mornings, once. */
function ensureWeeklyTrigger_() {
  const scheduled = ScriptApp.getProjectTriggers().some(function (t) {
    return t.getHandlerFunction() === 'sendWeeklyReport';
  });
  if (!scheduled) {
    ScriptApp.newTrigger('sendWeeklyReport').timeBased()
      .onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(WEEKLY_HOUR).create();
  }
}

/** Trigger handler: emails the report for the week that started last Sunday. */
function sendWeeklyReport() {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  emailWeeklyReport_(latestSunday_(tz), reportRecipients_(), true);
}

/** Menu: (re)send the weekly report to everyone on the list. */
function emailWeeklyReportMenu() {
  sendWeeklySelected_(reportRecipients_(), true);
}

/** Menu: send the weekly report only to whoever is signed in. */
function emailWeeklyReportToMe() {
  sendWeeklySelected_([Session.getEffectiveUser().getEmail()], false);
}

/** The week of the row selected on the Weekly tab, or else the latest week. */
function sendWeeklySelected_(recipients, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  let sunday = latestSunday_(tz);
  if (ss.getActiveSheet().getName() === WEEKLY_NAME) {
    const tab = ss.getActiveSheet();
    const row = SpreadsheetApp.getActiveRange().getRow();
    const week = row > 1 ? tab.getRange(row, tabColumns_(tab)['Week Of']).getValue() : '';
    if (week instanceof Date) sunday = parseSunday_(ymd_(week, tz));
  }
  SpreadsheetApp.getUi().alert(emailWeeklyReport_(sunday, recipients, record));
}

/** The most recent Sunday (today, on a Sunday), at noon. */
function latestSunday_(tz) {
  return parseDate_(weekOf_(new Date(), tz));
}

/** Builds and emails one week's report. record: note the outcome on the Weekly tab. */
function emailWeeklyReport_(sunday, recipients, record) {
  const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const note = function (text) {
    if (!record) return;
    const tab = weeklyTab_();
    const cols = tabColumns_(tab);
    const row = weeklyRow_(tab, cols, sunday, tz) || writeRow_(tab, cols, 0, { 'Week Of': sunday }, weeklyFormats_());
    tab.getRange(row, cols[WEEKLY_STATUS]).setValue(text);
  };
  if (!recipients.length) {
    note('Not sent: no addresses on the ' + EMAILS_NAME + ' tab');
    return 'There are no addresses on the ' + EMAILS_NAME + ' tab, so nothing was sent.';
  }
  try {
    const data = weeklyData_(sunday, tz);
    const am = data.services.filter(function (s) { return s.service === 'Sunday AM'; })[0];
    MailApp.sendEmail({
      to: recipients.join(','),
      subject: 'Weekly attendance: week of ' + Utilities.formatDate(sunday, tz, 'MMM d') +
        (am && am.total !== null ? ' (Sunday AM ' + am.total + ')' : ''),
      htmlBody: weeklyHtml_(data),
      body: weeklyText_(data),
      name: 'Grace Attendance',
    });
    note('Sent ' + Utilities.formatDate(new Date(), tz, 'M/d h:mm a') + ' to ' + recipients.length);
    return 'Weekly report for the week of ' + Utilities.formatDate(sunday, tz, 'MMM d') + ' emailed to ' + recipients.join(', ') + '.';
  } catch (err) {
    note('Failed: ' + err.message);
    return 'The weekly report could not be sent: ' + err.message;
  }
}

// ---- Data -------------------------------------------------------------------

/**
 * The week starting on sunday, and the same week last year (52 weeks back).
 * Services come from the Attendance tab, or the History tabs for weeks
 * before the ushers' page; office numbers from the Weekly tab, or History.
 */
function weeklyData_(sunday, tz) {
  const sheet = getSheet_();
  const cols = headerIndex_(sheet);
  const before = new Date(sunday.getTime() - 364 * DAY_MS);
  const lastServices = weekServices_(sheet, cols, before, tz);
  const week = weekServices_(sheet, cols, sunday, tz);
  const services = week.map(function (s) {
    const counts = SECTIONS.map(function (sec) {
      return {
        title: sec.title,
        subtotal: sec.areas.reduce(function (t, a) { return t + (Number(s.counts[a]) || 0); }, 0),
      };
    });
    const last = lastServices.filter(function (e) {
      return e.logged && e.service === s.service && e.otherName.toLowerCase() === s.otherName.toLowerCase();
    })[0];
    return {
      service: s.service,
      name: s.service === 'Other' ? s.otherName : s.service,
      date: parseDate_(s.date),
      logged: s.logged,
      total: s.logged ? s.total : historyTotal_(s.service, sunday, tz),
      sections: s.logged ? counts : null,
      lastYear: last ? last.total : historyTotal_(s.service, before, tz),
      notes: s.notes,
    };
  });
  return {
    sunday: sunday,
    tz: tz,
    services: services,
    office: officeValues_(sunday, tz),
    lastOffice: officeValues_(before, tz),
  };
}

/** A service's count in the History tabs for the week starting on sunday, or null. */
function historyTotal_(service, sunday, tz) {
  const start = ymd_(sunday, tz);
  const end = ymd_(new Date(sunday.getTime() + 6 * DAY_MS), tz);
  const year = Number(Utilities.formatDate(sunday, tz, 'yyyy'));
  const hit = historyCounts_(service, [year, year + 1]).filter(function (e) {
    return ymd_(e.date, tz) >= start && ymd_(e.date, tz) <= end;
  })[0];
  return hit ? hit.total : null;
}

// ---- Email HTML ---------------------------------------------------------------

/** A count for a table cell: a number, a dash for blank, or the old sheet's words ("No Awana"). */
function wCell_(v) {
  if (v === '' || v === null || v === undefined) return DASH;
  if (typeof v === 'number') return String(v);
  return '<span style="font-size:11px;color:' + THEME.muted + '">' + esc_(v) + '</span>';
}

/** "+12", "−5" or "±0". */
function wSigned_(c) {
  return (c > 0 ? '+' : c < 0 ? '−' : '±') + Math.abs(c);
}

/** The change from last year, up in green and down in amber, when both are numbers. */
function wChange_(n, last) {
  if (typeof n !== 'number' || typeof last !== 'number') return '';
  return '<span style="color:' + (n < last ? '#8A5200' : '#3D7A28') + ';font-weight:600">' + wSigned_(n - last) + '</span>';
}

/** A quick box's second line: "+12 vs. last year (400)". */
function wVersus_(n, last) {
  if (typeof last !== 'number') return '';
  return (typeof n === 'number' ? wSigned_(n - last) + ' ' : '') + 'vs. last year (' + last + ')';
}

/** Label | this week | last year | change. rows: [{label, n, last, strong}]. */
function wCompare_(rows) {
  const T = THEME;
  const cell = function (html, st) {
    return '<td style="' + FONT + 'padding:6px 5px;font-size:13px;color:' + T.ink + ';border-bottom:1px solid ' + T.line + ';' + (st || '') + '">' + html + '</td>';
  };
  const head = '<tr>' + ['', 'This week', 'Last year', ''].map(function (t, i) {
    return '<td style="' + FONT + 'padding:4px 5px;font-size:11px;letter-spacing:.5px;text-transform:uppercase;white-space:nowrap;color:' + T.muted +
      ';border-bottom:2px solid ' + T.brand + (i ? ';text-align:right' : '') + '">' + t + '</td>';
  }).join('') + '</tr>';
  return eTable_(head + rows.map(function (r) {
    const st = r.strong ? 'background:' + T.soft + ';font-weight:600;' : '';
    return '<tr>' + cell(esc_(r.label), st) +
      cell(wCell_(r.n), st + 'text-align:right;white-space:nowrap') +
      cell(wCell_(r.last), st + 'text-align:right;color:' + T.muted) +
      cell(wChange_(r.n, r.last), st + 'text-align:right;width:40px') + '</tr>';
  }).join(''));
}

/** A group's rows for wCompare_, with its total (when it has one) last. */
function wGroupRows_(d, g) {
  const rows = g.fields.map(function (f) {
    return { label: f.label, n: d.office.values[f.key], last: d.lastOffice.values[f.key] };
  });
  if (g.total) {
    // The total as saved, or added up here when the old sheet left it blank.
    const total = function (values, list) {
      const v = values[g.total];
      return v === undefined || v === '' ? sum_(list) : v;
    };
    rows.push({
      label: 'Total',
      n: total(d.office.values, rows.map(function (r) { return r.n; })),
      last: total(d.lastOffice.values, rows.map(function (r) { return r.last; })),
      strong: true,
    });
  }
  return rows;
}

/** The weekly report, as the email's HTML body. */
function weeklyHtml_(d) {
  const T = THEME;
  const tz = d.tz;
  const out = [];
  const end = new Date(d.sunday.getTime() + 6 * DAY_MS);
  const ss = wGroupRows_(d, officeGroup_('Sunday School'));
  const ssTotal = ss[ss.length - 1];

  // Quick boxes: each service, then Sunday School, two across
  const boxes = d.services.map(function (s, i) {
    const versus = wVersus_(s.total, s.lastYear);
    return eBox_(s.name, s.total === null ? '—' : String(s.total),
      esc_(Utilities.formatDate(s.date, tz, 'EEE M/d') + (s.total === null ? ', not logged' : '')) + (versus ? '<br>' + esc_(versus) : ''), i === 0);
  });
  boxes.push(eBox_('Sunday School', typeof ssTotal.n === 'number' ? String(ssTotal.n) : '—', esc_(wVersus_(ssTotal.n, ssTotal.last)) || '&nbsp;'));
  let quick = '';
  for (let i = 0; i < boxes.length; i += 2) quick += eRow_(boxes[i] + (boxes[i + 1] || ''));
  out.push(eCard_(quick));

  // Each service by section, as the ushers logged it
  const cell = function (html, st) {
    return '<td style="' + FONT + 'padding:6px 5px;font-size:13px;color:' + T.ink + ';border-bottom:1px solid ' + T.line + ';' + (st || '') + '">' + html + '</td>';
  };
  const head = '<tr>' + ['Service'].concat(SECTIONS.map(function (s) { return s.title; }), ['Total']).map(function (t, i) {
    return '<td style="' + FONT + 'padding:4px 5px;font-size:11px;letter-spacing:.5px;text-transform:uppercase;color:' + T.muted +
      ';border-bottom:2px solid ' + T.brand + (i ? ';text-align:right' : '') + '">' + esc_(t) + '</td>';
  }).join('') + '</tr>';
  const rows = d.services.map(function (s) {
    const name = esc_(s.name) + '<br><span style="font-size:11px;color:' + T.muted + '">' + Utilities.formatDate(s.date, tz, 'EEE M/d') + '</span>';
    if (!s.sections) {
      return '<tr>' + cell(name) + cell('<span style="font-size:12px;font-style:italic;color:' + T.muted + '">' +
        (s.total === null ? 'Not logged' : 'From the old attendance sheet') + '</span>', 'text-align:right" colspan="' + SECTIONS.length) +
        cell(s.total === null ? DASH : '<b>' + s.total + '</b>', 'text-align:right') + '</tr>';
    }
    return '<tr>' + cell(name) + s.sections.map(function (sec) { return cell(String(sec.subtotal), 'text-align:right'); }).join('') +
      cell('<b>' + s.total + '</b>', 'text-align:right') + '</tr>';
  });
  out.push(eCard_(eHeading_('Services') + eCaption_('What the ushers counted, by area of the building.') + eTable_(head + rows.join(''))));

  // Sunday School and the ministries beside the Sunday School classes
  out.push(eCard_(eRow_(
    eSlot_(eSubhead_('Sunday School', 4) + wCompare_(ss) +
      eSubhead_('Ministries', 14) + wCompare_(wGroupRows_(d, officeGroup_('Ministries'))), 'margin-bottom:12px') +
    eSlot_(eSubhead_('Sunday School classes', 4) + wCompare_(wGroupRows_(d, officeGroup_('Sunday School Classes'))), 'margin-bottom:12px'))));

  // Notes from the office and the ushers
  const notes = [];
  if (d.office.notes) notes.push(['Office', d.office.notes]);
  d.services.forEach(function (s) { if (s.notes) notes.push([s.name, s.notes]); });
  if (notes.length) {
    out.push(eCard_(eHeading_('Notes') + notes.map(function (n) {
      return '<p style="' + FONT + 'margin:6px 0 0;font-size:14px;line-height:1.45;color:' + T.ink + '"><b>' + esc_(n[0]) + ':</b> ' + esc_(n[1]) + '</p>';
    }).join('')));
  }

  const preview = d.services.map(function (s) { return s.name + ' ' + (s.total === null ? '—' : s.total); }).join(' · ') +
    (typeof ssTotal.n === 'number' ? ' · Sunday School ' + ssTotal.n : '');
  return ePage_(preview, 'Weekly attendance', 'Week of ' + Utilities.formatDate(d.sunday, tz, 'MMMM d'),
    Utilities.formatDate(d.sunday, tz, 'MMM d') + ' – ' + Utilities.formatDate(end, tz, 'MMM d, yyyy'),
    out.join('\n'), weeklyFooter_(d));
}

function weeklyFooter_(d) {
  const o = d.office;
  const who = o.from === 'weekly' && o.by
    ? 'Office numbers entered by ' + o.by + (o.at instanceof Date ? ', last saved ' + Utilities.formatDate(o.at, d.tz, 'EEE M/d h:mm a') : '') + '. '
    : o.from === 'history' ? 'Office numbers from the old attendance sheet. ' : 'The office hasn\'t entered numbers for this week. ';
  return who + 'Sent automatically every Saturday morning.';
}

// ---- Plain-text email -----------------------------------------------------------

function weeklyText_(d) {
  const tz = d.tz;
  const show = function (v) { return v === '' || v === null || v === undefined ? '—' : String(v); };
  const lines = ['Weekly attendance, week of ' + Utilities.formatDate(d.sunday, tz, 'MMMM d, yyyy'), ''];
  d.services.forEach(function (s) {
    lines.push(s.name + ' (' + Utilities.formatDate(s.date, tz, 'EEE M/d') + '): ' + (s.total === null ? 'not logged' : s.total) +
      (typeof s.lastYear === 'number' ? '; last year ' + s.lastYear : ''));
  });
  OFFICE_GROUPS.forEach(function (g) {
    lines.push('', g.title);
    wGroupRows_(d, g).forEach(function (r) {
      lines.push('  ' + r.label + ': ' + show(r.n) + (r.last !== undefined && r.last !== '' ? ' (last year ' + r.last + ')' : ''));
    });
  });
  if (d.office.notes) lines.push('', 'Office notes: ' + d.office.notes);
  d.services.forEach(function (s) { if (s.notes) lines.push('', s.name + ' notes: ' + s.notes); });
  lines.push('', weeklyFooter_(d));
  return lines.join('\n');
}
