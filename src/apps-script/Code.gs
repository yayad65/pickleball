/**
 * Get Pickled RSVP
 * Lives inside a Google Sheet in Yasser's Gmail account.
 *  - The pickleball app sends the regular players (with emails) and the game days here.
 *  - The day before each game day it asks Yasser to approve, then emails players a one-tap reply link,
 *    reminds anyone who hasn't replied at 3 PM and 6 PM, and sends Yasser a summary at 6 PM.
 *  - Nothing goes to players until Yasser approves that game day.
 *  - Test mode (on at first) sends every player email to Yasser instead.
 */
var TZ = 'America/New_York';
var TABS = {
  Settings: [['Setting', 'Value']],
  Players: [['Name', 'Email', 'Token', 'Active']],
  Days: [['Date', 'Approved', 'Approval asked', 'Invite sent', 'Reminder 1 sent', 'Reminder 2 sent', 'Summary sent']],
  Replies: [['Date', 'Token', 'Name', 'Answer', 'When']],
  Marked: [['Date', 'Name', 'Status', 'Sub for']],
  Log: [['When', 'What', 'To', 'Subject']]
};
var DEFAULTS = [
  ['Event name', 'Wed 7AM Get Pickled'],
  ['Your name', 'Yasser'],
  ['Your email', ''],
  ['Start time', '7:00 AM'],
  ['Courts', 2],
  ['Reply page', 'https://yayad65.github.io/pickleball/rsvp.html'],
  ['Test mode (all emails go to you)', 'TRUE'],
  ['Ask me to approve at (hour)', 8],
  ['Invite at (hour)', 10],
  ['First reminder at (hour)', 15],
  ['Second reminder and summary at (hour)', 18],
  ['Web app URL', ''],
  ['App connection code', '(appears after you deploy the web app)']
];

/* ---------- one-time setup ---------- */
function setup() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(TABS).forEach(function (name) {
    var sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (sh.getLastRow() === 0) { sh.appendRow(TABS[name][0]); sh.getRange(1, 1, 1, TABS[name][0].length).setFontWeight('bold'); sh.setFrozenRows(1); }
  });
  var st = ss.getSheetByName('Settings'), have = st.getDataRange().getValues().map(function (r) { return r[0]; });
  DEFAULTS.forEach(function (d) { if (have.indexOf(d[0]) < 0) st.appendRow(d); });
  if (!setting('Your email')) setSetting('Your email', Session.getEffectiveUser().getEmail());
  var first = ss.getSheets()[0]; if (first.getName() === 'Sheet1' && first.getLastRow() === 0) ss.deleteSheet(first);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('KEY')) props.setProperty('KEY', Utilities.getUuid().replace(/-/g, ''));
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(15).create();
  refreshConnectionCode();
  st.autoResizeColumns(1, 2);
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Get Pickled')
    .addItem('Show app connection code', 'showConnectionCode')
    .addItem('Send me a test email now', 'sendTestNow')
    .addToUi();
}
/* The real web app address. Google's own getUrl() can return a preview address that only works for you,
   so the address the app connects with (or one pasted into Settings) is used instead. */
var URL_RE = /^https:\/\/script\.google\.com\/macros\/s\/[\w-]+\/exec$/;
function webUrl() {
  var s = String(setting('Web app URL') || '').trim(); if (URL_RE.test(s)) return s;
  return '';
}
function refreshConnectionCode() {
  var url = webUrl();
  setSetting('App connection code', url ? url + '?k=' + key() : 'Copy the Web app URL from Deploy > Manage deployments, paste it in "Web app URL" above, then use Get Pickled > Show app connection code');
  return url ? url + '?k=' + key() : '';
}
function showConnectionCode() {
  var code = refreshConnectionCode();
  SpreadsheetApp.getUi().alert(code ? 'Paste this into the app (Setup > Players > Email invites):\n\n' + code
    : 'First copy the Web app URL: in Apps Script choose Deploy > Manage deployments and click Copy under Web app URL. Paste it into the Settings tab next to "Web app URL", then choose this menu item again.');
}

/* ---------- settings and sheet helpers ---------- */
function key() { return PropertiesService.getScriptProperties().getProperty('KEY'); }
function sheet(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function setting(label) {
  var rows = sheet('Settings').getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (rows[i][0] === label) return rows[i][1];
  return '';
}
function setSetting(label, value) {
  var sh = sheet('Settings'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (rows[i][0] === label) { sh.getRange(i + 1, 2).setValue(value); return; }
  sh.appendRow([label, value]);
}
/* Sheets turns "7:00 AM" into a time value; show it as plain text again. */
function startTime() {
  var v = setting('Start time');
  if (v instanceof Date) return Utilities.formatDate(v, SpreadsheetApp.getActive().getSpreadsheetTimeZone() || TZ, 'h:mm a');
  return String(v || '').replace(/^'/, '');
}
function testMode() { var v = String(setting('Test mode (all emails go to you)')).toUpperCase(); return v !== 'FALSE' && v !== 'NO' && v !== 'OFF'; }
function me() { return String(setting('Your email') || Session.getEffectiveUser().getEmail()); }
function iso(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function cellISO(v) { return v instanceof Date ? iso(v) : String(v).replace(/^'/, '').slice(0, 10); }
function nowNY() { var n = new Date(); return { date: iso(n), hour: +Utilities.formatDate(n, TZ, 'H'), stamp: Utilities.formatDate(n, TZ, 'yyyy-MM-dd HH:mm') }; }
function addDays(isoDate, n) { var p = isoDate.split('-').map(Number); var d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n, 12)); return d.toISOString().slice(0, 10); }
function niceDate(isoDate, opts) {
  var p = isoDate.split('-').map(Number), d = new Date(Date.UTC(p[0], p[1] - 1, p[2], 12));
  var days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var mons = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  if (opts === 'short') return days[d.getUTCDay()].slice(0, 3) + ', ' + mons[d.getUTCMonth()].slice(0, 3) + ' ' + d.getUTCDate();
  return days[d.getUTCDay()] + ', ' + mons[d.getUTCMonth()] + ' ' + d.getUTCDate();
}
function players() {
  var rows = sheet('Players').getDataRange().getValues().slice(1);
  return rows.map(function (r, i) { return { row: i + 2, name: String(r[0]).trim(), email: String(r[1]).trim(), token: String(r[2]), active: String(r[3]).toUpperCase() !== 'FALSE' }; })
    .filter(function (p) { return p.name; });
}
function activePlayers() { return players().filter(function (p) { return p.active && p.email; }); }
function gameDates() { return sheet('Days').getDataRange().getValues().slice(1).map(function (r) { return cellISO(r[0]); }).filter(Boolean); }
function dayState(d) {
  var sh = sheet('Days'), rows = sh.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) if (cellISO(rows[i][0]) === d) {
    return { row: i + 1, approved: String(rows[i][1] || ''), asked: rows[i][2], invite: rows[i][3], rem1: rows[i][4], rem2: rows[i][5], summary: rows[i][6] };
  }
  return null;
}
function markDay(d, col, value) { var s = dayState(d); if (s) sheet('Days').getRange(s.row, col).setValue(value); }
var COL = { approved: 2, asked: 3, invite: 4, rem1: 5, rem2: 6, summary: 7 };
function replies(d) {
  var out = {};
  sheet('Replies').getDataRange().getValues().slice(1).forEach(function (r) { if (cellISO(r[0]) === d) out[String(r[1])] = String(r[3]); });
  return out;
}
function marks(d) {
  var out = {}, subs = [], yes = {};
  var sh = sheet('Marked'); if (!sh) return { out: out, subs: subs, yes: yes };
  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    if (cellISO(r[0]) !== d) return;
    if (String(r[2]) === 'out') out[String(r[1])] = true;
    else if (String(r[2]) === 'in') yes[String(r[1])] = true;
    else if (String(r[2]) === 'sub') subs.push({ name: String(r[1]), forName: String(r[3] || '') });
  });
  return { out: out, subs: subs, yes: yes };
}
function tellChange(p, d, before, now) {
  var ev = String(setting('Event name')), word = function (x) { return x === 'in' ? 'coming' : 'not coming'; };
  var html = wrap('<p><b>' + esc(p.name) + '</b> changed their answer for <b>' + esc(niceDate(d)) + '</b>:</p>' +
    '<p style="font-size:20px"><span style="color:#5C6378;text-decoration:line-through">' + word(before) + '</span> &rarr; <b style="color:' + (now === 'in' ? '#1C5A40' : '#7A1F45') + '">' + word(now) + '</b></p>' +
    '<p>Where things stand now:</p>' + summaryHTML(d) + button(adminLink('day', d), 'Open the replies page'));
  MailApp.sendEmail({ to: me(), subject: p.name + ' changed to ' + word(now) + ' for ' + ev + ' (' + niceDate(d, 'short') + ')', htmlBody: html, name: String(setting('Your name') || 'Yasser') });
  log('change alert', me(), p.name + ': ' + before + ' -> ' + now + ' (' + d + ')');
}
function log(what, to, subject) { sheet('Log').appendRow([nowNY().stamp, what, to, subject]); }
function esc(s) { return String(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function first(name) { return String(name).split(/\s+/)[0]; }
function hourText(h) { h = +h; return (h % 12 || 12) + ' ' + (h < 12 ? 'AM' : 'PM'); }
function scriptId() { var m = webUrl().match(/\/s\/([^/]+)\//); return m ? m[1] : ''; }
function replyLink(token, d) { return setting('Reply page') + '?s=' + encodeURIComponent(scriptId()) + '&t=' + encodeURIComponent(token) + '&d=' + d; }
function adminLink(view, d) { return webUrl() + '?k=' + key() + (view ? '&view=' + view : '') + (d ? '&d=' + d : ''); }

/* ---------- emails ---------- */
function wrap(inner) { return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:16px;line-height:1.45;color:#1C2440;max-width:520px">' + inner + '</div>'; }
function button(href, label) { return '<p style="margin:22px 0"><a href="' + esc(href) + '" style="display:inline-block;background:#22408E;color:#FFFFFF;text-decoration:none;font-weight:bold;padding:14px 28px;border-radius:999px">' + esc(label) + '</a></p>'; }
function send(to, subject, html, what, forName) {
  var realTo = to, subj = subject, body = html;
  if (forName && testMode()) {
    realTo = me(); subj = '[Test for ' + forName + '] ' + subject;
    body = '<p style="font-family:Arial,sans-serif;font-size:13px;background:#FFF4D6;padding:8px 12px;border-radius:8px">Test mode: this would have gone to ' + esc(forName) + ' (' + esc(to) + ').</p>' + html;
  }
  MailApp.sendEmail({ to: realTo, subject: subj, htmlBody: body, name: String(setting('Your name') || 'Yasser'), replyTo: me() });
  log(what, realTo, subj);
}
function playerEmail(p, d, kind) {
  var ev = String(setting('Event name')), time = startTime(), sign = String(setting('Your name') || 'Yasser');
  var r1 = hourText(setting('First reminder at (hour)')), r2 = hourText(setting('Second reminder and summary at (hour)'));
  var when = '<b>' + esc(niceDate(d)) + ' at ' + esc(time) + '</b>';
  var subject, intro, note = '';
  if (kind === 'invite') {
    subject = ev + ': tomorrow (' + niceDate(d, 'short') + '). Are you in?';
    intro = '<p>We play tomorrow, ' + when + '. Can you make it?</p>';
    note = '<p style="font-size:14px;color:#5C6378">Heads up: if I don\'t hear back by ' + r1 + ' today, you\'ll get a reminder at ' + r1 + ' and another at ' + r2 + '. Replying stops them.</p>';
  } else if (kind === 'rem1') {
    subject = 'Reminder: ' + ev + ' tomorrow. Are you in?';
    intro = '<p>Just checking: can you make it tomorrow, ' + when + '? One tap below lets me know.</p>';
    note = '<p style="font-size:14px;color:#5C6378">If I don\'t hear back, you\'ll get one more reminder at ' + r2 + '.</p>';
  } else {
    subject = 'Last reminder: ' + ev + ' tomorrow. Are you in?';
    intro = '<p>Last check for tomorrow, ' + when + '. Please let me know so I can find a sub if needed.</p>';
  }
  var html = wrap('<p>Hi ' + esc(first(p.name)) + ',</p>' + intro + button(replyLink(p.token, d), 'Reply in one tap') + note +
    '<p style="font-size:13px;color:#5C6378">This link is just for you.</p><p>Thanks,<br>' + esc(sign) + '</p>');
  return { subject: subject, html: html };
}
function sendBatch(d, kind) {
  var rs = replies(d), mk = marks(d), n = 0;
  activePlayers().forEach(function (p) {
    if (mk.out[p.name]) return;
    if (mk.yes[p.name]) return;
    if (kind !== 'invite' && rs[p.token]) return;
    var m = playerEmail(p, d, kind); send(p.email, m.subject, m.html, kind, p.name); n++;
  });
  return n;
}
function tally(d) {
  var rs = replies(d), mk = marks(d), out = { yes: [], no: [], marked: [], none: [], subs: mk.subs, told: {} };
  players().filter(function (p) { return p.active; }).forEach(function (p) {
    if (mk.out[p.name]) { out.marked.push(p.name); return; }
    var a = rs[p.token];
    if (!a && mk.yes[p.name]) { out.yes.push(p.name); out.told[p.name] = true; return; }
    if (!p.email) return;
    (a === 'in' ? out.yes : a === 'out' ? out.no : out.none).push(p.name);
  });
  return out;
}
function summaryHTML(d) {
  var t = tally(d), courts = +setting('Courts') || 2, need = Math.max(0, courts * 4 - t.yes.length - t.subs.length);
  var line = need ? 'You need <b>' + need + ' sub' + (need > 1 ? 's' : '') + '</b> for ' + courts + ' full court' + (courts > 1 ? 's' : '') + '.' : 'You have enough for ' + courts + ' full court' + (courts > 1 ? 's' : '') + '.';
  var li = function (arr, color) { return arr.length ? arr.map(function (n) { return '<li style="color:' + color + '">' + esc(n) + '</li>'; }).join('') : '<li style="color:#5C6378">Nobody</li>'; };
  return '<p>' + line + '</p>' +
    '<p><b>Coming (' + t.yes.length + ')</b></p><ul>' + (t.yes.length ? t.yes.map(function (n) { return '<li style="color:#1C5A40">' + esc(n) + (t.told[n] ? ' <span style="color:#5C6378">(told you)</span>' : '') + '</li>'; }).join('') : li([], '')) + '</ul>' +
    (t.subs.length ? '<p><b>Subs (' + t.subs.length + ')</b></p><ul>' + t.subs.map(function (x) { return '<li style="color:#1C5A40">' + esc(x.name) + (x.forName ? ' <span style="color:#5C6378">(for ' + esc(x.forName) + ')</span>' : '') + '</li>'; }).join('') + '</ul>' : '') +
    '<p><b>Can\'t make it (' + (t.no.length + t.marked.length) + ')</b></p><ul>' + li(t.no, '#7A1F45').replace('<li style="color:#5C6378">Nobody</li>', t.marked.length ? '' : '<li style="color:#5C6378">Nobody</li>') + t.marked.map(function (n) { return '<li style="color:#7A1F45">' + esc(n) + ' <span style="color:#5C6378">(you marked)</span></li>'; }).join('') + '</ul>' +
    '<p><b>No reply (' + t.none.length + ')</b></p><ul>' + li(t.none, '#6B5200') + '</ul>';
}
function sendSummary(d) {
  var ev = String(setting('Event name'));
  MailApp.sendEmail({ to: me(), subject: ev + ' tomorrow (' + niceDate(d, 'short') + '): who\'s coming', name: String(setting('Your name') || 'Yasser'),
    htmlBody: wrap('<p>Replies for <b>' + esc(niceDate(d)) + '</b>:</p>' + summaryHTML(d) + button(adminLink('day', d), 'Open the replies page')) });
  log('summary', me(), 'Summary ' + d);
}
function askApproval(d) {
  var ev = String(setting('Event name')), ps = activePlayers();
  var missing = players().filter(function (p) { return p.active && !p.email; }).map(function (p) { return p.name; });
  var html = wrap('<p>Tomorrow is <b>' + esc(ev) + '</b>, ' + esc(niceDate(d)) + '.</p>' +
    '<p>Nothing has been sent yet. Approve to email ' + ps.length + ' player' + (ps.length === 1 ? '' : 's') + ' the reply link at ' + hourText(setting('Invite at (hour)')) +
    ', with reminders at ' + hourText(setting('First reminder at (hour)')) + ' and ' + hourText(setting('Second reminder and summary at (hour)')) + ' to anyone who hasn\'t replied.</p>' +
    (testMode() ? '<p style="background:#FFF4D6;padding:8px 12px;border-radius:8px">Test mode is on: every email will come to you instead of the players.</p>' : '') +
    (missing.length ? '<p style="color:#7A1F45">No email for: ' + esc(missing.join(', ')) + '. They won\'t get anything.</p>' : '') +
    button(adminLink('day', d), 'Review and approve'));
  MailApp.sendEmail({ to: me(), subject: 'Approve emails for ' + ev + ' (' + niceDate(d, 'short') + ')?', htmlBody: html, name: String(setting('Your name') || 'Yasser') });
  log('approval request', me(), 'Approve ' + d);
}

/* ---------- the schedule (runs every 15 minutes) ---------- */
function tick() {
  var lock = LockService.getScriptLock(); if (!lock.tryLock(20000)) return;
  try {
    refreshConnectionCode();
    var n = nowNY(), d = addDays(n.date, 1);
    if (gameDates().indexOf(d) < 0) return;
    var s = dayState(d), H = function (lbl) { return +setting(lbl); };
    var hInvite = H('Invite at (hour)'), h1 = H('First reminder at (hour)'), h2 = H('Second reminder and summary at (hour)');
    if (!s.asked && n.hour >= H('Ask me to approve at (hour)')) { askApproval(d); markDay(d, COL.asked, n.stamp); s.asked = n.stamp; }
    var ok = s.approved === 'all' || s.approved === 'invite';
    if (ok && !s.invite && n.hour >= hInvite) { sendBatch(d, 'invite'); markDay(d, COL.invite, n.stamp); s.invite = n.stamp; }
    var inviteHour = s.invite ? +String(s.invite instanceof Date ? Utilities.formatDate(s.invite, TZ, 'H') : String(s.invite).slice(11, 13)) : 99;
    if (s.approved === 'all' && s.invite && !s.rem1 && n.hour >= h1 && inviteHour < h1) { sendBatch(d, 'rem1'); markDay(d, COL.rem1, n.stamp); s.rem1 = n.stamp; }
    if (s.approved === 'all' && s.invite && !s.rem2 && n.hour >= h2 && inviteHour < h2) { sendBatch(d, 'rem2'); markDay(d, COL.rem2, n.stamp); }
    if (s.invite && !s.summary && n.hour >= h2) { sendSummary(d); markDay(d, COL.summary, n.stamp); }
  } finally { lock.releaseLock(); }
}
function approve(k, d, scope) {
  if (k !== key()) throw new Error('Not allowed');
  if (['all', 'invite', 'skip'].indexOf(scope) < 0) throw new Error('Unknown choice');
  if (!dayState(d)) throw new Error('That game day is not on the list');
  markDay(d, COL.approved, scope);
  if (scope !== 'skip') tick();
  return dayView(d);
}
function sendTestNow(k) {
  if (k !== undefined && k !== key()) throw new Error('Not allowed');
  var today = nowNY().date, d = gameDates().filter(function (x) { return x > today; }).sort()[0] || addDays(today, 1);
  var p = { name: String(setting('Your name') || 'Yasser') + ' (test)', email: me(), token: 'TEST' };
  ['invite', 'rem1', 'rem2'].forEach(function (kind) {
    var m = playerEmail(p, d, kind);
    MailApp.sendEmail({ to: me(), subject: '[Test] ' + m.subject, htmlBody: m.html, name: String(setting('Your name') || 'Yasser') });
  });
  log('test', me(), 'Test emails for ' + d);
  try { SpreadsheetApp.getActive().toast('Sent you the invite and both reminders as a test.'); } catch (e) {}
  return 'Sent you 3 test emails (the invite and both reminders) for ' + niceDate(d) + '.';
}

/* ---------- web app ---------- */
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function findPlayer(t) { if (t === 'TEST') return { name: String(setting('Your name') || 'Yasser') + ' (test)', token: 'TEST' }; return players().filter(function (p) { return p.token === t && p.token; })[0]; }
function doGet(e) {
  var q = e.parameter || {};
  if (q.api === 'info' || q.api === 'answer') {
    var p = findPlayer(String(q.t || '')), d = String(q.d || '');
    if (!p || !/^\d{4}-\d{2}-\d{2}$/.test(d)) return json({ ok: false, error: 'This link isn\'t valid any more. Please ask Yasser for a new one.' });
    if (p.token !== 'TEST' && gameDates().indexOf(d) < 0) return json({ ok: false, error: 'That game day isn\'t on the schedule.' });
    if (q.api === 'answer') {
      var a = q.a === 'in' ? 'in' : q.a === 'out' ? 'out' : '';
      if (!a) return json({ ok: false, error: 'Unknown answer.' });
      if (d < nowNY().date) return json({ ok: false, error: 'That game day has passed.' });
      var lock = LockService.getScriptLock(); lock.waitLock(20000);
      try {
        var sh = sheet('Replies'), rows = sh.getDataRange().getValues(), done = false, before = '';
        for (var i = 1; i < rows.length; i++) if (cellISO(rows[i][0]) === d && String(rows[i][1]) === p.token) { before = String(rows[i][3]); sh.getRange(i + 1, 4, 1, 2).setValues([[a, nowNY().stamp]]); done = true; break; }
        if (!done) sh.appendRow(["'" + d, p.token, p.name, a, nowNY().stamp]);
      } finally { lock.releaseLock(); }
      if (before && before !== a) { try { tellChange(p, d, before, a); } catch (x) {} }
    }
    return json({ ok: true, first: first(p.name), organizer: String(setting('Your name') || 'Yasser'), contact: me(), event: String(setting('Event name')), date: d, dateText: niceDate(d), time: startTime(), answer: replies(d)[p.token] || '' });
  }
  if (q.k !== key()) return HtmlService.createHtmlOutput('<p style="font-family:Arial,sans-serif">Nothing to see here.</p>');
  if (q.view === 'day' && q.d) return page(dayView(String(q.d)));
  return page(homeView());
}
function doPost(e) {
  var b = {}; try { b = JSON.parse(e.postData.contents); } catch (x) { return json({ ok: false, error: 'Bad request' }); }
  if (b.k !== key()) return json({ ok: false, error: 'This connection code is wrong. Copy it again from the Sheet.' });
  if (b.api === 'replies') { var d = String(b.d || ''); var t = tally(d), rs = replies(d), mk = marks(d), byName = {}; activePlayers().forEach(function (p) { byName[p.name] = mk.out[p.name] ? 'out' : (rs[p.token] || (mk.yes[p.name] ? 'told' : '')); }); return json({ ok: true, date: d, replies: byName, counts: { in: t.yes.length, out: t.no.length + t.marked.length, none: t.none.length, subs: t.subs.length } }); }
  if (b.api === 'day' || b.api === 'approve') {
    var dd = String(b.d || '');
    if (!dayState(dd)) return json({ ok: false, error: 'That game day is not on the list. Tap Update sheet first.' });
    if (b.api === 'approve') {
      if (['all', 'invite', 'skip'].indexOf(b.scope) < 0) return json({ ok: false, error: 'Unknown choice' });
      var cur = dayState(dd); if (cur.invite && !(b.scope === 'all' && cur.approved === 'invite')) return json({ ok: false, error: 'The invite already went out for that day.' });
      approve(key(), dd, String(b.scope));
    }
    var st = dayState(dd), at = function (v) { return !v ? '' : v instanceof Date ? Utilities.formatDate(v, TZ, "EEE h:mm a") : String(v).replace(/^'/, ''); };
    return json({ ok: true, date: dd, approved: st.approved, asked: !!st.asked, invite: at(st.invite), rem1: at(st.rem1), rem2: at(st.rem2), summary: at(st.summary), testMode: testMode(),
      hours: { invite: hourText(setting('Invite at (hour)')), rem1: hourText(setting('First reminder at (hour)')), rem2: hourText(setting('Second reminder and summary at (hour)')) } });
  }
  if (b.api !== 'sync') return json({ ok: false, error: 'Unknown request' });
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = sheet('Players'), cur = players(), byKey = {};
    cur.forEach(function (p) { byKey[(p.email || p.name).toLowerCase()] = p; });
    var list = (b.players || []).filter(function (p) { return p && p.name; }).slice(0, 60);
    var rows = list.map(function (p) {
      var em = String(p.email || '').trim(), old = byKey[(em || p.name).toLowerCase()] || byKey[String(p.name).toLowerCase()];
      return [String(p.name).trim(), em, old && old.token ? old.token : Utilities.getUuid().replace(/-/g, '').slice(0, 16), 'TRUE'];
    });
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
    if (rows.length) sh.getRange(2, 1, rows.length, 4).setValues(rows);
    var ds = sheet('Days'), have = gameDates();
    (b.dates || []).filter(function (d) { return /^\d{4}-\d{2}-\d{2}$/.test(d) && have.indexOf(d) < 0; }).forEach(function (d) { ds.appendRow(["'" + d, '', '', '', '', '', '']); });
    if (b.avail && typeof b.avail === 'object') {
      var ms = sheet('Marked') || (function () { var x = SpreadsheetApp.getActive().insertSheet('Marked'); x.appendRow(TABS.Marked[0]); return x; })();
      var keep = ms.getDataRange().getValues().slice(1).filter(function (r) { return r[0] !== '' && !(cellISO(r[0]) in b.avail); });
      var add = [];
      Object.keys(b.avail).forEach(function (d) {
        if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) return; var a = b.avail[d] || {};
        (a.out || []).slice(0, 60).forEach(function (n) { add.push(["'" + d, String(n), 'out', '']); });
        (a.subs || []).slice(0, 60).forEach(function (x) { add.push(["'" + d, String(x.name || ''), 'sub', String(x.forName || '')]); });
        (a.in || []).slice(0, 60).forEach(function (n) { add.push(["'" + d, String(n), 'in', '']); });
      });
      var all = keep.map(function (r) { return ["'" + cellISO(r[0]), r[1], r[2], r[3]]; }).concat(add);
      if (ms.getLastRow() > 1) ms.getRange(2, 1, ms.getLastRow() - 1, 4).clearContent();
      if (all.length) ms.getRange(2, 1, all.length, 4).setValues(all);
    }
    if (b.url && URL_RE.test(String(b.url))) { setSetting('Web app URL', String(b.url)); refreshConnectionCode(); }
    if (b.courts) setSetting('Courts', +b.courts);
    if (b.startTime) setSetting('Start time', "'" + String(b.startTime));
    var noEmail = rows.filter(function (r) { return !r[1]; }).map(function (r) { return r[0]; });
    return json({ ok: true, players: rows.length, noEmail: noEmail, dates: gameDates().length, testMode: testMode() });
  } finally { lock.releaseLock(); }
}

/* ---------- Yasser's pages (opened from his emails) ---------- */
function page(inner) {
  return HtmlService.createHtmlOutput('<!doctype html><html><head><meta name="viewport" content="width=device-width,initial-scale=1"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;background:#EEF0F4;color:#1C2440;margin:0;padding:16px}main{max-width:560px;margin:0 auto;display:flex;flex-direction:column;gap:14px}' +
    '.card{background:#fff;border-radius:18px;padding:16px;box-shadow:0 2px 0 #D4D8E0,0 12px 20px -12px rgba(28,40,80,.35)}h1{font-size:22px;margin:0}h2{font-size:17px;margin:0 0 6px}' +
    'button{font:inherit;font-weight:bold;border:0;border-radius:999px;padding:14px 18px;min-height:48px;cursor:pointer;width:100%}.pri{background:#22408E;color:#fff}.sec{background:#EEF1F6;color:#22408E}.warn{background:#FFF4D6;padding:10px 12px;border-radius:12px}' +
    '.row{display:flex;justify-content:space-between;gap:8px;padding:6px 0;border-top:1px solid #EEF1F6}.row:first-child{border-top:0}.btns{display:flex;flex-direction:column;gap:8px}#msg{font-weight:bold;color:#1C5A40}a{color:#22408E}' +
    '</style></head><body><main id="main">' + inner + '</main><script>var K=' + JSON.stringify(key()) + ';function go(d,s){document.getElementById("msg").textContent="Saving…";google.script.run.withSuccessHandler(function(h){document.getElementById("main").innerHTML=h;}).withFailureHandler(function(e){document.getElementById("msg").textContent=e.message;}).approve(K,d,s);}' +
    'function test(){document.getElementById("msg").textContent="Sending…";google.script.run.withSuccessHandler(function(t){document.getElementById("msg").textContent=t;}).withFailureHandler(function(e){document.getElementById("msg").textContent=e.message;}).sendTestNow(K);}</script></body></html>')
    .setTitle(String(setting('Event name')) + ' replies').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
function dayView(d) {
  var s = dayState(d), ev = String(setting('Event name')); if (!s) return '<div class="card"><p>That game day is not on the list.</p></div>';
  var t = tally(d), status = s.approved === 'all' ? 'Approved: invite and both reminders' : s.approved === 'invite' ? 'Approved: invite only' : s.approved === 'skip' ? 'Skipped: no emails this week' : 'Not approved yet. Nothing has been sent.';
  var at = function (v) { return v instanceof Date ? Utilities.formatDate(v, TZ, "EEE MMM d 'at' h:mm a") : String(v).replace(/^'/, ''); };
  var sent = [s.invite ? 'Invite sent ' + at(s.invite) : '', s.rem1 ? 'Reminder 1 sent ' + at(s.rem1) : '', s.rem2 ? 'Reminder 2 sent ' + at(s.rem2) : ''].filter(Boolean).join('<br>');
  if (s.invite && s.approved === 'invite') sent += '<br>The buttons go away once the invite is out.';
  var btns = (!s.invite && s.approved !== 'skip') ? '<div class="btns"><button class="pri" onclick="go(\'' + d + '\',\'all\')">Approve: invite at ' + hourText(setting('Invite at (hour)')) + ' + reminders at ' + hourText(setting('First reminder at (hour)')) + ' and ' + hourText(setting('Second reminder and summary at (hour)')) + '</button>' +
    '<button class="sec" onclick="go(\'' + d + '\',\'invite\')">Invite only, no reminders</button><button class="sec" onclick="go(\'' + d + '\',\'skip\')">Skip this week (send nothing)</button></div>' : '';
  return '<div class="card"><h1>' + esc(ev) + '</h1><p>' + esc(niceDate(d)) + ' · ' + esc(startTime()) + '</p>' +
    (testMode() ? '<p class="warn">Test mode is on: all emails go to you, not the players.</p>' : '') + '<p><b>' + status + '</b></p>' + (sent ? '<p style="font-size:14px;color:#5C6378">' + sent + '</p>' : '') + btns + '<p id="msg"></p></div>' +
    '<div class="card"><h2>Replies</h2>' + summaryHTML(d) + '</div>';
}
function homeView() {
  var today = nowNY().date, ds = gameDates().filter(function (d) { return d >= today; }).sort().slice(0, 8);
  var rows = ds.map(function (d) { var s = dayState(d), t = tally(d); return '<div class="row"><a href="' + esc(adminLink('day', d)) + '" target="_top">' + esc(niceDate(d, 'short')) + '</a><span>' + (s.approved || 'not approved') + ' · ' + t.yes.length + ' in, ' + t.no.length + ' out</span></div>'; }).join('');
  return '<div class="card"><h1>' + esc(setting('Event name')) + '</h1>' + (testMode() ? '<p class="warn">Test mode is on: all emails go to you, not the players.</p>' : '') +
    '<p>' + activePlayers().length + ' players with email · ' + gameDates().length + ' game days</p></div>' +
    '<div class="card"><h2>Coming up</h2>' + (rows || '<p>No game days yet. Send them from the app.</p>') + '</div>' +
    '<div class="card"><h2>Try it</h2><p>Sends you the invite and both reminders for the next game day. Nobody else gets anything.</p><button class="pri" onclick="test()">Send me a test now</button><p id="msg"></p></div>';
}
