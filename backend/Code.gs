/**
 * Ankur & Soumya — RSVP backend (Google Apps Script, bound to the RSVP Tracker sheet)
 *
 * The website (ankursoumya.com/rsvp) calls this as a web app:
 *   GET  ?action=unlock&pw=0525                    -> checks the website password
 *   GET  ?action=lookup&name=First Last&pw=0525    -> finds the guest's row by full name (guest or plus one)
 *   POST {action:"submit", row, token, attending, plusOne, message} -> writes the response
 *
 * Runs as a standalone Apps Script web app that opens the RSVP sheet by ID.
 * Day-to-day settings live in the sheet's "Website Settings" tab (live tab, website password, alert emails, reply-to, open/closed).
 *
 * Script properties (Project Settings > Script properties), technical only:
 *   FROM_NAME      sender name, default "Ankur & Soumya".
 *   BREVO_API_KEY  optional. When set, emails are sent through Brevo from FROM_EMAIL.
 *   FROM_EMAIL     sender address for Brevo, default rsvp@ankursoumya.com.
 *   SECRET         created automatically; signs the row tokens.
 */

var SHEET_ID = '1dhXm5g0u0N9XoH1SrlfSVOe9u_Qx7x8aUYbk92F5Go0';
var SETTINGS_TAB = 'Website Settings';

var HEADER_ROW = 3;
var DEADLINE = new Date('2027-01-01T00:00:00+08:00'); // responses accepted through Dec 31
var SITE = 'https://ankursoumya.com';

var COLS = {
  first: 'First Name', last: 'Last Name',
  p1first: 'Plus One First Name', p1last: 'Plus One Last Name',
  status: 'RSVP Status', p1status: 'Plus One RSVP', message: 'Message to Couple',
  respondedAt: 'Responded At'
};

/* ---------- web app entry points ---------- */

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'unlock') return json_({ ok: pwOk_(p.pw), closed: isClosed_() });
    if (p.action === 'lookup') return json_(lookup_(p.name || '', p.pw || ''));
    if (p.action === 'ping') return json_({ ok: true, sheet: sheetName_() });
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', detail: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'submit') return json_(submit_(body));
    return json_({ ok: false, error: 'unknown_action' });
  } catch (err) {
    return json_({ ok: false, error: 'server', detail: String(err) });
  }
}

/* ---------- lookup ---------- */

function lookup_(name, pw) {
  if (!pwOk_(pw)) return { ok: false, error: 'bad_password' };
  var q = norm_(name);
  if (q.split(' ').length < 2) return { ok: false, error: 'need_full_name' };
  var cache = CacheService.getScriptCache(), key = 'tries_' + Utilities.base64EncodeWebSafe(q).slice(0, 200);
  var tries = Number(cache.get(key) || 0);
  if (tries >= 10) return { ok: false, error: 'too_many' };
  var t = table_(), exact = [], loose = [];
  t.rows.forEach(function (r, i) {
    var first = norm_(r[t.c.first]), last = norm_(r[t.c.last]);
    if (!first && !last) return;
    var g = nameMatch_(q, first, last), p = nameMatch_(q, norm_(r[t.c.p1first]), norm_(r[t.c.p1last]));
    if (g === 2) exact.push({ i: i, role: 'guest' }); else if (p === 2) exact.push({ i: i, role: 'plus' });
    else if (g === 1) loose.push({ i: i, role: 'guest' }); else if (p === 1) loose.push({ i: i, role: 'plus' });
  });
  var hits = exact.length ? exact : loose;
  if (!hits.length) { cache.put(key, String(tries + 1), 3600); return { ok: false, error: 'no_match' }; }
  if (hits.length > 1) return { ok: false, error: 'duplicate' };
  cache.remove(key);
  var party = party_(t, hits[0].i);
  party.role = hits[0].role;
  return { ok: true, party: party, closed: isClosed_() };
}

/** 2 = exact full name; 1 = close enough (middle name left out, or no last name in the sheet); 0 = no. */
function nameMatch_(q, first, last) {
  if (!first) return 0;
  if (q === (first + ' ' + last).trim()) return 2;
  var f1 = first.split(' ')[0];
  if (last && q === f1 + ' ' + last) return 1;
  if (!last && q.split(' ')[0] === f1) return 1;
  return 0;
}

/** The one website password, kept in the Website Settings tab ("Website password"). */
function pwOk_(pw) {
  var want = String(setting_('Website password', '0525')).trim().toLowerCase();
  return want !== '' && String(pw || '').trim().toLowerCase() === want;
}

function party_(t, i) {
  var r = t.rows[i], row = i + HEADER_ROW + 1;
  var p1 = String(r[t.c.p1first] || '').trim();
  var p1l = String(r[t.c.p1last] || '').trim();
  return {
    row: row,
    token: sign_(row, r[t.c.first], r[t.c.last]),
    first: String(r[t.c.first]).trim(),
    last: String(r[t.c.last]).trim(),
    plusOne: p1 ? { first: p1, last: p1l } : null,
    status: String(r[t.c.status] || ''),
    plusOneStatus: String(r[t.c.p1status] || ''),
    message: String(r[t.c.message] || ''),
    respondedAt: r[t.c.respondedAt] ? String(r[t.c.respondedAt]) : ''
  };
}

/* ---------- submit ---------- */

function submit_(b) {
  if (isClosed_()) return { ok: false, error: 'closed' };
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var t = table_();
    var i = Number(b.row) - HEADER_ROW - 1;
    var r = t.rows[i];
    if (!r || sign_(b.row, r[t.c.first], r[t.c.last]) !== b.token) return { ok: false, error: 'stale' };

    var attending = b.attending === 'yes' ? 'Attending' : b.attending === 'no' ? 'Not attending' : '';
    if (!attending) return { ok: false, error: 'missing_answer' };
    var hasP1 = String(r[t.c.p1first] || '').trim() !== '';
    var p1 = '';
    if (hasP1) {
      p1 = b.plusOne === 'yes' ? 'Attending' : b.plusOne === 'no' ? 'Not attending' : '';
      if (!p1) return { ok: false, error: 'missing_answer' };
    }
    var msg = String(b.message || '').slice(0, 1000);
    var now = Utilities.formatDate(new Date(), 'Asia/Makassar', 'yyyy-MM-dd HH:mm') + ' (Bali)';
    var prev = String(r[t.c.status] || '');

    var sh = t.sheet, row = Number(b.row);
    sh.getRange(row, t.c.status + 1).setValue(attending);
    sh.getRange(row, t.c.p1status + 1).setValue(hasP1 ? p1 : '');
    sh.getRange(row, t.c.message + 1).setValue(msg);
    sh.getRange(row, t.c.respondedAt + 1).setValue(now);
    SpreadsheetApp.flush();

    var party = { first: String(r[t.c.first]).trim(), last: String(r[t.c.last]).trim(),
      p1: hasP1 ? (String(r[t.c.p1first]).trim() + ' ' + String(r[t.c.p1last] || '').trim()).trim() : '',
      status: attending, p1status: p1, message: msg, updated: !!prev,
      byPlus: hasP1 && b.by === 'plus' };

    try { notifyCouple_(party, now); } catch (err) { console.error(err); }

    return { ok: true, status: attending, plusOneStatus: p1 };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- emails ---------- */

function sendGuest_(p) {
  var yes = p.status === 'Attending';
  var subject = yes ? 'See you in Bali! Your RSVP is confirmed' : 'Your RSVP for Ankur & Soumya\'s wedding';
  var lines = [row_(p.first + ' ' + p.last, p.status)];
  if (p.p1) lines.push(row_(p.p1, p.p1status));
  var intro = yes
    ? 'We are so happy you\'ll be celebrating with us in Bali, May 23 to 26, 2027.'
    : 'We\'re sorry you can\'t make it, and we\'ll miss you. Thank you for letting us know.';
  var html = shell_(
    '<p style="margin:0 0 6px;font-size:18px">Dear ' + esc_(p.p1 ? p.first + ' & ' + p.p1.split(' ')[0] : p.first) + ',</p>' +
    '<p style="margin:0 0 18px">' + intro + (p.updated ? ' Your updated response is below.' : ' Here\'s what we received:') + '</p>' +
    '<table style="width:100%;border-collapse:collapse;margin:0 0 18px">' + lines.join('') + '</table>' +
    (p.message ? '<p style="margin:0 0 18px;font-style:italic">Your note: "' + esc_(p.message) + '"</p>' : '') +
    '<p style="margin:0 0 18px">Need to change anything? Just RSVP again at <a href="' + SITE + '/rsvp/" style="color:#5A1A1E">ankursoumya.com/rsvp</a> before December 31, 2026.</p>' +
    (yes ? '<p style="margin:0">Schedule, travel tips and everything else are on <a href="' + SITE + '" style="color:#5A1A1E">ankursoumya.com</a>.</p>' : '') +
    '<p style="margin:22px 0 0">With love,<br>Ankur &amp; Soumya</p>');
  send_(p.email, subject, html, plain_(html));
}

function notifyCouple_(p, when) {
  var to = setting_('Send new-RSVP alerts to', '');
  if (!to) return;
  var who = p.first + ' ' + p.last;
  var by = p.byPlus ? p.p1 + ' (plus one of ' + who + ')' : who;
  var subject = (p.updated ? 'RSVP updated: ' : 'New RSVP: ') + who + ' (' + p.status + (p.p1 ? ', +1 ' + (p.p1status === 'Attending' ? 'yes' : 'no') : '') + ')';
  var html = shell_(
    '<p style="margin:0 0 14px">' + esc_(by) + (p.updated ? ' updated their RSVP' : ' just RSVP\'d') + ' on ' + esc_(when) + '.</p>' +
    '<table style="width:100%;border-collapse:collapse;margin:0 0 14px">' + row_(who, p.status) + (p.p1 ? row_(p.p1 + ' (plus one)', p.p1status) : '') + '</table>' +
    (p.message ? '<p style="margin:0 0 14px;font-style:italic">"' + esc_(p.message) + '"</p>' : '') +
    '<p style="margin:0">Live count: ' + esc_(countLine_()) + '</p>' +
    '<p style="margin:12px 0 0"><a href="' + ss_().getUrl() + '" style="color:#5A1A1E">Open the RSVP sheet</a></p>');
  // one send per inbox, so a block on one address never stops the others
  String(to).split(',').map(function (x) { return x.trim(); }).filter(String).forEach(function (addr) {
    try { send_(addr, subject, html, plain_(html)); } catch (err) { console.error(addr + ': ' + err); }
  });
}

function countLine_() {
  var t = table_(), yes = 0, no = 0, pending = 0;
  t.rows.forEach(function (r) {
    if (!String(r[t.c.first]).trim()) return;
    var hasP1 = String(r[t.c.p1first] || '').trim() !== '';
    var s = String(r[t.c.status] || '');
    if (s === 'Attending') yes++; else if (s === 'Not attending') no++; else pending++;
    if (hasP1) {
      var s1 = String(r[t.c.p1status] || '');
      if (s1 === 'Attending') yes++; else if (s1 === 'Not attending') no++; else pending++;
    }
  });
  return yes + ' attending, ' + no + ' not attending, ' + pending + ' still to reply';
}

function send_(to, subject, html, text) {
  var name = prop_('FROM_NAME', 'Ankur & Soumya');
  var replyTo = setting_('Guest replies go to', '');
  var brevo = prop_('BREVO_API_KEY', '');
  if (brevo) {
    var payload = { sender: { name: name, email: prop_('FROM_EMAIL', 'rsvp@ankursoumya.com') },
      to: String(to).split(',').map(function (x) { return { email: x.trim() }; }),
      subject: subject, htmlContent: html, textContent: text };
    if (replyTo) payload.replyTo = { email: replyTo };
    var res = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
      method: 'post', contentType: 'application/json', headers: { 'api-key': brevo },
      payload: JSON.stringify(payload), muteHttpExceptions: true });
    if (res.getResponseCode() >= 300) throw new Error('Brevo ' + res.getResponseCode() + ' ' + res.getContentText().slice(0, 200));
    return;
  }
  var opts = { htmlBody: html, name: name };
  if (replyTo) opts.replyTo = replyTo;
  MailApp.sendEmail(to, subject, text, opts);
}

function shell_(inner) {
  return '<div style="background:#F4EBD6;padding:28px 12px;font-family:Georgia,\'Times New Roman\',serif;color:#5A1A1E">' +
    '<div style="max-width:520px;margin:0 auto;background:#FBF6EA;border:1px solid #E2CFA0;border-radius:6px;padding:30px 28px;font-size:16px;line-height:1.55">' +
    '<div style="text-align:center;font-size:26px;letter-spacing:1px;margin:0 0 4px">Ankur &amp; Soumya</div>' +
    '<div style="text-align:center;font-size:11px;letter-spacing:4px;color:#9C7A2A;margin:0 0 22px">BALI &middot; MAY 23 &ndash; 26, 2027</div>' +
    inner + '</div></div>';
}

function row_(who, status) {
  var c = status === 'Attending' ? '#2F6B3F' : '#8A3A3A';
  return '<tr><td style="padding:8px 0;border-bottom:1px solid #E8DAB8">' + esc_(who) + '</td>' +
    '<td style="padding:8px 0;border-bottom:1px solid #E8DAB8;text-align:right;color:' + c + '">' + esc_(status) + '</td></tr>';
}

/* ---------- sheet helpers ---------- */

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

var SETTINGS_CACHE = null;
function settings_() {
  if (SETTINGS_CACHE) return SETTINGS_CACHE;
  var out = {}, sh = ss_().getSheetByName(SETTINGS_TAB);
  if (sh) sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 2).getDisplayValues().forEach(function (r) { if (r[0]) out[String(r[0]).trim()] = String(r[1]).trim(); });
  return (SETTINGS_CACHE = out);
}
function setting_(k, d) { var v = settings_()[k]; return v ? v : d; }
function setSetting_(k, v) {
  var sh = ss_().getSheetByName(SETTINGS_TAB), vals = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]).trim() === k) { sh.getRange(i + 1, 2).setValue(v); SETTINGS_CACHE = null; return; }
}
function isClosed_() { return new Date() >= DEADLINE || /^no$/i.test(setting_('RSVPs open', 'Yes')); }
function sheetName_() { return setting_('Live tab', 'TEST RSVP Tracker'); }

function table_() {
  var sh = ss_().getSheetByName(sheetName_());
  if (!sh) throw new Error('Tab not found: ' + sheetName_());
  var last = Math.max(sh.getLastRow(), HEADER_ROW);
  var width = sh.getLastColumn();
  var head = sh.getRange(HEADER_ROW, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
  var c = {};
  Object.keys(COLS).forEach(function (k) {
    var idx = head.indexOf(COLS[k]);
    if (idx < 0) throw new Error('Missing column "' + COLS[k] + '" on ' + sheetName_());
    c[k] = idx;
  });
  var rows = last > HEADER_ROW ? sh.getRange(HEADER_ROW + 1, 1, last - HEADER_ROW, width).getDisplayValues() : [];
  return { sheet: sh, c: c, rows: rows };
}

function prop_(k, d) { var v = PropertiesService.getScriptProperties().getProperty(k); return v == null || v === '' ? d : v; }

function sign_(row, first, last) {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid(); props.setProperty('SECRET', s); }
  var raw = Utilities.computeHmacSha256Signature(row + '|' + norm_(first) + '|' + norm_(last) + '|' + sheetName_(), s);
  return Utilities.base64EncodeWebSafe(raw).slice(0, 22);
}

function norm_(s) {
  return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase()
    .replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function maskEmail_(e) {
  var m = String(e).split('@'); if (m.length !== 2) return '';
  return m[0].slice(0, 2) + '•••@' + m[1];
}

function esc_(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function plain_(html) { return html.replace(/<br>/g, '\n').replace(/<\/(p|tr|div)>/g, '\n').replace(/<[^>]+>/g, ' ').replace(/&amp;/g, '&').replace(/&middot;/g, '·').replace(/&ndash;/g, '–').replace(/[ \t]+/g, ' ').replace(/\n\s+/g, '\n').trim(); }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- sheet menu for Ankur & Soumya ---------- */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('RSVP website')
    .addItem('Which tab is live?', 'menuWhich')
    .addItem('Switch website to the TEST tab', 'menuTest')
    .addItem('Switch website to the real RSVP Tracker', 'menuLive')
    .addToUi();
}
function menuWhich() { SpreadsheetApp.getUi().alert('The website is reading and writing: "' + sheetName_() + '"\n\n' + countLine_()); }
function menuTest() { setSetting_('Live tab', 'TEST RSVP Tracker'); menuWhich(); }
function menuLive() { setSetting_('Live tab', 'RSVP Tracker'); menuWhich(); }

/** Old hourly passcode job: passcodes are gone, so it just removes its own trigger. */
function autoPasscodes() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'autoPasscodes') ScriptApp.deleteTrigger(t); });
}

/** Run once from the editor to grant permissions. */
function setup() {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('FROM_NAME')) p.setProperty('FROM_NAME', 'Ankur & Soumya');
  sign_(4, 'a', 'b');
  autoPasscodes();
  Logger.log('Live tab: ' + sheetName_() + ' | ' + countLine_() + ' | mail quota left: ' + MailApp.getRemainingDailyQuota());
}
