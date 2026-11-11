/**
 * Ankur & Soumya — RSVP backend (Google Apps Script, bound to the RSVP Tracker sheet)
 *
 * The website (ankursoumya.com/rsvp) calls this as a web app:
 *   GET  ?action=lookup&name=First Last&code=123456  -> finds the guest's row (name + passcode)
 *   POST {action:"submit", row, token, attending, plusOne, message} -> writes the response
 *
 * Script properties (Project Settings > Script properties):
 *   SHEET_NAME     which tab is live. "TEST RSVP Tracker" while testing, "RSVP Tracker" to go live.
 *   NOTIFY_TO      where the couple's "new RSVP" emails go (comma separated).
 *   REPLY_TO       reply-to address on guest confirmations (e.g. rsvp@ankursoumya.com).
 *   FROM_NAME      sender name, default "Ankur & Soumya".
 *   BREVO_API_KEY  optional. When set, emails are sent through Brevo from FROM_EMAIL.
 *   FROM_EMAIL     sender address for Brevo, default rsvp@ankursoumya.com.
 *   SECRET         created automatically; signs the row tokens.
 */

var HEADER_ROW = 3;
var DEADLINE = new Date('2026-12-16T00:00:00+08:00'); // responses accepted through Dec 15
var SITE = 'https://ankursoumya.com';

var COLS = {
  first: 'First Name', last: 'Last Name', nick: 'Nick Name', email: 'Email', phone: 'WhatsApp Phone Number',
  p1first: 'Plus One First Name', p1last: 'Plus One Last Name',
  status: 'RSVP Status', p1status: 'Plus One RSVP', message: 'Message to Couple',
  respondedAt: 'Responded At', confirmation: 'Confirmation Sent', code: 'Passcode'
};

/* ---------- web app entry points ---------- */

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'lookup') return json_(lookup_(p.name || '', p.code || ''));
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

function lookup_(name, code) {
  var q = norm_(name);
  code = String(code || '').replace(/\D/g, '');
  if (q.split(' ').length < 2) return { ok: false, error: 'need_full_name' };
  if (code.length !== 6) return { ok: false, error: 'need_code' };
  var cache = CacheService.getScriptCache(), key = 'tries_' + Utilities.base64EncodeWebSafe(q).slice(0, 200);
  var tries = Number(cache.get(key) || 0);
  if (tries >= 10) return { ok: false, error: 'too_many' };
  var t = table_();
  var hit = -1;
  t.rows.forEach(function (r, i) {
    if (hit > -1) return;
    var first = norm_(r[t.c.first]), last = norm_(r[t.c.last]), nick = norm_(r[t.c.nick]);
    var p1 = norm_(r[t.c.p1first] + ' ' + r[t.c.p1last]);
    if (!first && !last) return;
    var names = [first + ' ' + last];
    if (nick) names.push(nick + ' ' + last);
    if (p1.trim()) names.push(p1);
    if (names.indexOf(q) > -1 && String(r[t.c.code]).replace(/\D/g, '') === code) hit = i;
  });
  if (hit < 0) { cache.put(key, String(tries + 1), 3600); return { ok: false, error: 'no_match' }; }
  cache.remove(key);
  return { ok: true, party: party_(t, hit), closed: new Date() >= DEADLINE };
}

/** 6-digit passcode derived from the guest's phone number (stable), unique within the tab. */
function makeCode_(phone, salt, used) {
  var props = PropertiesService.getScriptProperties();
  var s = props.getProperty('SECRET');
  if (!s) { s = Utilities.getUuid(); props.setProperty('SECRET', s); }
  var base = String(phone || '').replace(/\D/g, '') || ('row' + salt);
  for (var n = 0; n < 50; n++) {
    var h = Utilities.computeHmacSha256Signature(base + '|' + n, s);
    var v = ((h[0] & 255) * 16777216 + (h[1] & 255) * 65536 + (h[2] & 255) * 256 + (h[3] & 255)) % 900000 + 100000;
    var c = String(v);
    if (!used[c]) { used[c] = true; return c; }
  }
  throw new Error('Could not make a unique passcode');
}

/** Fills the Passcode column for every guest that doesn't have one yet. */
function fillPasscodes_(sheetName) {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  if (!sh) return 0;
  var width = sh.getLastColumn(), last = sh.getLastRow();
  var head = sh.getRange(HEADER_ROW, 1, 1, width).getValues()[0].map(function (h) { return String(h).trim(); });
  var ci = head.indexOf(COLS.code), fi = head.indexOf(COLS.first), pi = head.indexOf(COLS.phone);
  if (ci < 0 || last <= HEADER_ROW) return 0;
  var vals = sh.getRange(HEADER_ROW + 1, 1, last - HEADER_ROW, width).getDisplayValues();
  var used = {}, made = 0;
  vals.forEach(function (r) { if (r[ci]) used[String(r[ci]).replace(/\D/g, '')] = true; });
  var out = vals.map(function (r, i) {
    if (r[ci] || !String(r[fi]).trim()) return [r[ci]];
    made++; return [makeCode_(r[pi], i, used)];
  });
  var rng = sh.getRange(HEADER_ROW + 1, ci + 1, out.length, 1);
  rng.setNumberFormat('@');
  rng.setValues(out);
  return made;
}

function party_(t, i) {
  var r = t.rows[i], row = i + HEADER_ROW + 1;
  var p1 = String(r[t.c.p1first] || '').trim();
  var p1l = String(r[t.c.p1last] || '').trim();
  var email = String(r[t.c.email] || '').trim();
  return {
    row: row,
    token: sign_(row, r[t.c.first], r[t.c.last]),
    first: String(r[t.c.first]).trim(),
    last: String(r[t.c.last]).trim(),
    plusOne: p1 ? { first: p1, last: p1l } : null,
    status: String(r[t.c.status] || ''),
    plusOneStatus: String(r[t.c.p1status] || ''),
    message: String(r[t.c.message] || ''),
    respondedAt: r[t.c.respondedAt] ? String(r[t.c.respondedAt]) : '',
    emailHint: email ? maskEmail_(email) : ''
  };
}

/* ---------- submit ---------- */

function submit_(b) {
  if (new Date() >= DEADLINE) return { ok: false, error: 'closed' };
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
      status: attending, p1status: p1, message: msg, email: String(r[t.c.email] || '').trim(), updated: !!prev };

    var sent = '';
    try { if (party.email) { sendGuest_(party); sent = 'Sent ' + now; } else { sent = 'No email on file'; } }
    catch (err) { sent = 'Failed: ' + String(err).slice(0, 120); }
    sh.getRange(row, t.c.confirmation + 1).setValue(sent);
    try { notifyCouple_(party, now); } catch (err) { console.error(err); }

    return { ok: true, status: attending, plusOneStatus: p1, emailed: sent.indexOf('Sent') === 0, emailHint: party.email ? maskEmail_(party.email) : '' };
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
    '<p style="margin:0 0 6px;font-size:18px">Dear ' + esc_(p.first) + ',</p>' +
    '<p style="margin:0 0 18px">' + intro + (p.updated ? ' Your updated response is below.' : ' Here\'s what we received:') + '</p>' +
    '<table style="width:100%;border-collapse:collapse;margin:0 0 18px">' + lines.join('') + '</table>' +
    (p.message ? '<p style="margin:0 0 18px;font-style:italic">Your note: "' + esc_(p.message) + '"</p>' : '') +
    '<p style="margin:0 0 18px">Need to change anything? Just RSVP again at <a href="' + SITE + '/rsvp/" style="color:#5A1A1E">ankursoumya.com/rsvp</a> before December 15, 2026.</p>' +
    (yes ? '<p style="margin:0">Schedule, travel tips and everything else are on <a href="' + SITE + '" style="color:#5A1A1E">ankursoumya.com</a>.</p>' : '') +
    '<p style="margin:22px 0 0">With love,<br>Ankur &amp; Soumya</p>');
  send_(p.email, subject, html, plain_(html));
}

function notifyCouple_(p, when) {
  var to = prop_('NOTIFY_TO', '');
  if (!to) return;
  var who = p.first + ' ' + p.last;
  var subject = (p.updated ? 'RSVP updated: ' : 'New RSVP: ') + who + ' (' + p.status + (p.p1 ? ', +1 ' + (p.p1status === 'Attending' ? 'yes' : 'no') : '') + ')';
  var html = shell_(
    '<p style="margin:0 0 14px">' + esc_(who) + (p.updated ? ' updated their RSVP' : ' just RSVP\'d') + ' on ' + esc_(when) + '.</p>' +
    '<table style="width:100%;border-collapse:collapse;margin:0 0 14px">' + row_(who, p.status) + (p.p1 ? row_(p.p1 + ' (plus one)', p.p1status) : '') + '</table>' +
    (p.message ? '<p style="margin:0 0 14px;font-style:italic">"' + esc_(p.message) + '"</p>' : '') +
    '<p style="margin:0">Live count: ' + esc_(countLine_()) + '</p>' +
    '<p style="margin:12px 0 0"><a href="' + SpreadsheetApp.getActive().getUrl() + '" style="color:#5A1A1E">Open the RSVP sheet</a></p>');
  send_(to, subject, html, plain_(html));
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
  var replyTo = prop_('REPLY_TO', '');
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

function sheetName_() { return prop_('SHEET_NAME', 'TEST RSVP Tracker'); }

function table_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(sheetName_());
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
    .addSeparator()
    .addItem('Create passcodes for new guests', 'menuCodes')
    .addToUi();
}
function menuWhich() { SpreadsheetApp.getUi().alert('The website is reading and writing: "' + sheetName_() + '"\n\n' + countLine_()); }
function menuTest() { PropertiesService.getScriptProperties().setProperty('SHEET_NAME', 'TEST RSVP Tracker'); menuWhich(); }
function menuCodes() {
  var n = fillPasscodes_(sheetName_());
  SpreadsheetApp.getUi().alert(n ? 'Created ' + n + ' new passcode(s) on "' + sheetName_() + '".' : 'Every guest on "' + sheetName_() + '" already has a passcode.');
}
function menuLive() { PropertiesService.getScriptProperties().setProperty('SHEET_NAME', 'RSVP Tracker'); menuWhich(); }

/** Run once from the editor to grant permissions and set defaults. */
function setup() {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SHEET_NAME')) p.setProperty('SHEET_NAME', 'TEST RSVP Tracker');
  if (!p.getProperty('FROM_NAME')) p.setProperty('FROM_NAME', 'Ankur & Soumya');
  sign_(4, 'a', 'b');
  Logger.log('Passcodes created: ' + fillPasscodes_('TEST RSVP Tracker'));
  Logger.log('Live tab: ' + sheetName_() + ' | ' + countLine_() + ' | mail quota left: ' + MailApp.getRemainingDailyQuota());
}
