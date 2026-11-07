/**
 * Valley View 2nd Hour Attendance — Code.gs (v6: 90-Day Email Auth + High-Volume Check-In Queue)
 *
 * v6 HIGHLIGHTS
 *   • Check-ins from QR/NFC scans are written to a "Check-In Log" sheet with one fast append (no sheet-wide
 *     lock). A batch "flush" merges them into the class sheets (every minute via trigger, and whenever a
 *     leader opens/refreshes the app). 200 people scanning at once no longer queue behind a lock.
 *   • The scan page's member/date list is cached for hours, so a crowd does not re-read the Sheet.
 *   • Sign-in code requests can no longer be used to lock a leader out.
 *   • After pasting: run authorizeEmailSetup once, then 2nd Hour Tools ▸ Turn On Auto-Sync, then deploy a New version.
 *
 * IMPORTANT ONE-TIME EMAIL AUTHORIZATION:
 *   After pasting this code into Extensions → Apps Script:
 *   1. Click the "▷ Run" button at the top of the Apps Script editor (with `authorizeEmailSetup` selected)
 *      and click "Review permissions → Allow" once. (This grants Google permission to send 6-digit login emails.)
 *   2. Click Deploy → Manage deployments → Edit (pencil) → New version → Deploy.
 */

const CLASSES = ['Sunday School', 'Elders Quorum', 'Relief Society', 'Gospel Foundation', 'Temple Prep'];
const MAX_ROWS = 600;        // class-sheet rows read (roster + header + total)
const MAX_MEMBERS = 600;     // Member Info rows read
const DATE_RE = /^\d{4}-\d{2}-\d{2}$/;
const SESSION_DAYS = 90;          // Leaders stay signed in for 90 days per device
const LOGIN_CODE_TTL_SEC = 600;   // 6-digit email login codes expire after 10 minutes
const MAIL_SENDER_NAME = 'Valley View 2nd Hour Attendance';
const CODE_COOLDOWN_SEC = 45;     // minimum gap between sign-in emails to the same address
const CHECKIN_LOG = 'Check-In Log';
const SCAN_CACHE_KEY = 'scan_base_v2';
const SCAN_CACHE_TTL = 21600;     // 6 hours (CacheService maximum)

/**
 * ▷ RUN THIS FUNCTION ONCE IN THE APPS SCRIPT EDITOR TO GRANT EMAIL PERMISSIONS!
 * Clicking "▷ Run" on `authorizeEmailSetup` prompts Google's OAuth dialog so `MailApp`
 * and `GmailApp` are authorized to send 6-digit verification codes when `doPost` runs.
 */
function authorizeEmailSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCheckInLog_(ss);
  ensureVisitorsSheet_(ss);
  const quota = MailApp.getRemainingDailyQuota();
  const leaders = Object.values(getAllowedLeaders_());
  const msg =
    '✅ Email & Sheet permissions are authorized!\n\n' +
    '• Remaining daily email quota: ' + quota + '\n' +
    '• Authorized emails found in "Report Contacts": ' + leaders.length + '\n\n' +
    'Next: 2nd Hour Tools ▸ Turn On Auto-Sync, then Deploy → Manage deployments → Edit (pencil) → New version → Deploy.';
  console.log(msg);
  try {
    if (ss) SpreadsheetApp.getUi().alert('2nd Hour Attendance Setup Complete', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {}
}

// ---------- Google Sheet Menu ----------
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('2nd Hour Tools')
      .addItem('🚀 Authorize Email Sign-In (Run Once)', 'authorizeEmailSetup')
      .addSeparator()
      .addItem('🔗 Show App & Classroom QR Links', 'showLinksDialog')
      .addItem('👥 Check Authorized Leadership Emails', 'showAllowedLeadersDialog')
      .addSeparator()
      .addItem('🌐 Set GitHub Pages App URL', 'manageGitHubUrlDialog')
      .addItem('🔒 Sign Out All Devices (Reset 90-Day Sessions)', 'signOutAllDevices')
      .addSeparator()
      .addItem('⚡ Turn On Auto-Sync (check-ins → class sheets)', 'installAutoSync')
      .addItem('🔄 Sync Check-Ins Now', 'syncCheckInsNow')
      .addItem('⏸ Turn Off Auto-Sync', 'removeAutoSync')
      .addSeparator()
      .addItem('🧹 Delete Load-Test Visitors', 'cleanupLoadTestVisitors')
      .addToUi();
  } catch (e) {}
}

function onEdit() {
  try {
    CacheService.getScriptCache().removeAll(['allowed_leaders_v1', SCAN_CACHE_KEY]);
  } catch (e) {}
}

function signOutAllDevices() {
  const props = PropertiesService.getScriptProperties();
  const nextEpoch = String(Number(props.getProperty('SESSION_EPOCH') || '1') + 1);
  props.setProperty('SESSION_EPOCH', nextEpoch);
  props.deleteProperty('ADMIN_KEY');
  try {
    CacheService.getScriptCache().remove('allowed_leaders_v1');
    SpreadsheetApp.getUi().alert('All signed-in devices have been signed out. Leaders will be prompted for a new 6-digit email code next time they open the app.');
  } catch (e) {}
}

function showAllowedLeadersDialog() {
  const ui = SpreadsheetApp.getUi();
  const leaders = Object.values(getAllowedLeaders_());
  if (!leaders.length) {
    ui.alert('No Authorized Emails Found', 'Add leadership names, callings, and email addresses to the "Report Contacts" sheet (Columns A–D) so they can sign in.', ui.ButtonSet.OK);
    return;
  }
  const lines = leaders.map(u => '• ' + u.name + ' (' + u.calling + ') — ' + u.email);
  ui.alert('Authorized Leaders (from Report Contacts)', 'These emails can request a 6-digit sign-in code:\n\n' + lines.join('\n'), ui.ButtonSet.OK);
}

function manageGitHubUrlDialog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const current = props.getProperty('GITHUB_PAGES_URL') || '';
  const resp = ui.prompt(
    'GitHub Pages Web App URL',
    (current ? 'Current GitHub Pages URL:\n' + current + '\n\n' : '') +
      'Paste your published GitHub Pages URL (e.g. https://yourname.github.io/2nd-hour-attendance/):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() === ui.Button.OK) {
    const val = String(resp.getResponseText() || '').trim();
    props.setProperty('GITHUB_PAGES_URL', val);
    clearScanCache_();
    ui.alert(val ? 'GitHub Pages URL saved:\n' + val : 'GitHub Pages URL cleared.');
  }
}

function showLinksDialog() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const execUrl = ScriptApp.getService().getUrl() || '<Deploy Web App first>';
  const ghUrl = (props.getProperty('GITHUB_PAGES_URL') || '').replace(/\/+$/, '');
  const base = ghUrl || execUrl;
  const lines = [
    'Leadership App URL (Sign in with Report Contacts email):',
    base,
    '',
    'Classroom QR / NFC Self Check-In Links:'
  ];
  CLASSES.forEach(c => {
    lines.push('• ' + c + ': ' + base + '?class=' + encodeURIComponent(c));
  });
  ui.alert('2nd Hour Attendance Links', lines.join('\n'), ui.ButtonSet.OK);
}

// ---------- Cryptographic & Session Helpers (90-Day Email Auth) ----------
function getAuthSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty('AUTH_HMAC_SECRET_V1');
  if (!secret) {
    secret = Utilities.getUuid() + '-' + Utilities.getUuid() + '-' + Utilities.getUuid();
    props.setProperty('AUTH_HMAC_SECRET_V1', secret);
  }
  return secret;
}

function hmacWebSafe_(message) {
  const sig = Utilities.computeHmacSha256Signature(String(message), getAuthSecret_());
  return Utilities.base64EncodeWebSafe(sig).replace(/=+$/g, '');
}

function safeEqual_(a, b) {
  a = String(a || '');
  b = String(b || '');
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) {
    diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  }
  return diff === 0;
}

function emailHash_(email) {
  return hmacWebSafe_('em|' + String(email || '').trim().toLowerCase()).slice(0, 24);
}

function userError_(msg) {
  const err = new Error(msg);
  err.name = 'UserError';
  err.code = 'invalid';
  return err;
}

function authError_(msg) {
  const err = new Error(msg || 'Please sign in again.');
  err.name = 'AuthError';
  err.code = 'auth';
  return err;
}

/**
 * Reads authorized leaders from the "Report Contacts" sheet (Columns A..D: Calling, Name, Phone, Email).
 */
function getAllowedLeaders_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('allowed_leaders_v1');
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }

  const users = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rcSheet = ss ? ss.getSheetByName('Report Contacts') : null;
  if (rcSheet && rcSheet.getLastRow() >= 2) {
    const rcRows = Math.min(100, rcSheet.getLastRow());
    rcSheet.getRange(2, 1, rcRows - 1, 4).getValues().forEach(r => {
      const calling = String(r[0] || '').trim();
      const name = String(r[1] || '').trim();
      const phone = cleanPhone_(r[2]);
      const email = cleanEmail_(r[3]).toLowerCase();
      if (!email || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) return;
      if (!users[email]) {
        users[email] = {
          email: email,
          name: name || email,
          calling: calling || 'Ward Leadership',
          phone: phone
        };
      } else if (calling && users[email].calling.indexOf(calling) === -1) {
        users[email].calling += ' / ' + calling;
      }
    });
  }

  try { cache.put('allowed_leaders_v1', JSON.stringify(users), 120); } catch (e) {}
  return users;
}

function signSession_(email) {
  const props = PropertiesService.getScriptProperties();
  const exp = Math.floor(Date.now() / 1000) + SESSION_DAYS * 86400;
  const epoch = props.getProperty('SESSION_EPOCH') || '1';
  const payload = Utilities.base64EncodeWebSafe(email + '|' + exp + '|' + epoch).replace(/=+$/g, '');
  return payload + '.' + hmacWebSafe_('session|' + payload).slice(0, 32);
}

function requireLeader_(req) {
  const token = String((req && req.token) || '').trim();
  const parts = token.split('.');
  if (parts.length !== 2) throw authError_();
  if (!safeEqual_(parts[1], hmacWebSafe_('session|' + parts[0]).slice(0, 32))) throw authError_();

  let decoded;
  try {
    decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString();
  } catch (e) {
    throw authError_();
  }
  const f = decoded.split('|');
  const epoch = PropertiesService.getScriptProperties().getProperty('SESSION_EPOCH') || '1';
  if (f.length !== 3 || Number(f[1]) < Date.now() / 1000 || f[2] !== epoch) {
    throw authError_('Your 90-day session has expired. Please sign in again.');
  }
  const email = String(f[0] || '').trim().toLowerCase();
  const user = getAllowedLeaders_()[email];
  if (!user) {
    throw authError_('Your email is no longer listed in Report Contacts.');
  }
  return user;
}

function apiRequestCode_(req) {
  const email = String((req && req.email) || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) {
    return { ok: false, code: 'bad_email', message: 'Please enter a valid email address.' };
  }
  const generic = {
    ok: true,
    message: 'If that address is listed on Report Contacts, a 6-digit code is on its way. (A new code can be sent once a minute.)'
  };

  const user = getAllowedLeaders_()[email];
  if (!user) {
    // Same rough response time as a real send, and nothing is written to the cache for unknown addresses.
    Utilities.sleep(900);
    return generic;
  }

  // The code only ever goes to the leader's own inbox, so a stranger requesting codes cannot lock the
  // leader out: if a fresh code is already out, we simply don't send another one yet.
  const cache = CacheService.getScriptCache();
  const eh = emailHash_(email);
  const key = 'code_att:' + eh;
  const existing = cache.get(key);
  if (existing) {
    try {
      const rec = JSON.parse(existing);
      if (rec && rec.s && (Date.now() - rec.s) < CODE_COOLDOWN_SEC * 1000 && rec.t < 5) return generic;
    } catch (e) {}
  }
  if (MailApp.getRemainingDailyQuota() < 5) {
    return { ok: false, code: 'rate', message: 'The daily email limit was reached. Please try again tomorrow.' };
  }

  const num = parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000;
  const code = String(num).padStart(6, '0');
  cache.put(
    key,
    JSON.stringify({ h: hmacWebSafe_('code|' + email + '|' + code), t: 0, s: Date.now() }),
    LOGIN_CODE_TTL_SEC
  );
  try {
    sendLoginCodeEmail_(user, code);
  } catch (err) {
    cache.remove(key);
    const errStr = String((err && err.message) || err || '');
    console.error('Sign-in code email failed: ' + errStr);
    if (/permission|authoriz|scope|MailApp|GmailApp/i.test(errStr)) {
      return {
        ok: false,
        code: 'mail_auth',
        message: 'One-time setup needed: In your Google Sheet, open Extensions ▸ Apps Script, click "▷ Run" at the top (on authorizeEmailSetup) to click "Allow", then deploy a New version.'
      };
    }
    return {
      ok: false,
      code: 'mail',
      message: 'Could not send email (' + errStr.slice(0, 90) + '). Check Apps Script permissions.'
    };
  }
  return generic;
}

function sendLoginCodeEmail_(user, code) {
  const firstName = extractFirstName_(user.name) || user.name;
  const subject = 'Your 2nd Hour Attendance sign-in code: ' + code;
  const plainBody =
    'Hi ' + firstName + ',\n\n' +
    'Your Valley View 2nd Hour Attendance sign-in code is:\n\n' +
    '    ' + code + '\n\n' +
    'This code expires in ' + Math.round(LOGIN_CODE_TTL_SEC / 60) + ' minutes. Once entered, your device will stay signed in for ' + SESSION_DAYS + ' days.\n\n' +
    'If you did not request this code, you can safely ignore this email.\n\n' +
    '— Valley View Ward 2nd Hour Attendance';

  const htmlBody =
    '<div style="font-family:\'Inter\',-apple-system,BlinkMacSystemFont,\'Segoe UI\',Roboto,sans-serif;max-width:420px;margin:0 auto;background:#232323;color:#ffffff;border-radius:18px;overflow:hidden;border:1px solid #383838;">' +
      '<div style="background:#555888;padding:20px 24px;">' +
        '<div style="font-size:11px;color:#e0e2f5;text-transform:uppercase;letter-spacing:0.08em;font-weight:700;">Valley View Ward</div>' +
        '<div style="font-family:Georgia,serif;font-style:italic;font-weight:900;font-size:22px;color:#ffffff;margin-top:2px;">2nd Hour Attendance</div>' +
      '</div>' +
      '<div style="padding:24px;">' +
        '<p style="margin:0 0 12px;font-size:14px;color:#e0e0e0;">Hi <b>' + safeHtml_(firstName) + '</b> (' + safeHtml_(user.calling) + '),</p>' +
        '<p style="margin:0 0 16px;font-size:13px;color:#bdbdbd;line-height:1.5;">Use this 6-digit verification code to sign in to the 2nd Hour Attendance app:</p>' +
        '<div style="background:#f5f5f5;color:#232323;border-radius:12px;padding:16px;text-align:center;font-size:28px;font-weight:800;letter-spacing:0.3em;margin-bottom:16px;">' +
          safeHtml_(code) +
        '</div>' +
        '<p style="margin:0;font-size:11.5px;color:#9e9e9e;line-height:1.5;">This code expires in ' + Math.round(LOGIN_CODE_TTL_SEC / 60) + ' minutes. Your phone will stay signed in for ' + SESSION_DAYS + ' days.</p>' +
      '</div>' +
    '</div>';

  // Use MailApp first (lighter OAuth scope), fallback to GmailApp
  try {
    MailApp.sendEmail({
      to: user.email,
      subject: subject,
      body: plainBody,
      htmlBody: htmlBody,
      name: MAIL_SENDER_NAME
    });
  } catch (mailErr) {
    GmailApp.sendEmail(user.email, subject, plainBody, {
      name: MAIL_SENDER_NAME,
      htmlBody: htmlBody
    });
  }
}

function apiVerifyCode_(req) {
  const email = String((req && req.email) || '').trim().toLowerCase();
  const code = String((req && req.code) || '').replace(/\D/g, '');
  const cache = CacheService.getScriptCache();
  const key = 'code_att:' + emailHash_(email);
  const raw = cache.get(key);
  if (!raw || code.length !== 6) {
    return { ok: false, code: 'bad_code', message: 'That code has expired or is not right. Please request a new one.' };
  }
  let rec;
  try { rec = JSON.parse(raw); } catch (e) {
    return { ok: false, code: 'bad_code', message: 'Invalid code state. Please request a new code.' };
  }
  if (rec.t >= 5) {
    cache.remove(key);
    return { ok: false, code: 'bad_code', message: 'Too many attempts. Please request a new code.' };
  }
  if (!safeEqual_(rec.h, hmacWebSafe_('code|' + email + '|' + code))) {
    rec.t++;
    cache.put(key, JSON.stringify(rec), LOGIN_CODE_TTL_SEC);
    return { ok: false, code: 'bad_code', message: 'That code is not right. Please check it and try again.' };
  }
  cache.remove(key);
  const user = getAllowedLeaders_()[email];
  if (!user) {
    return { ok: false, code: 'bad_code', message: 'That email is not listed in Report Contacts.' };
  }
  return {
    ok: true,
    token: signSession_(email),
    me: {
      email: user.email,
      name: user.name,
      calling: user.calling
    }
  };
}

// ---------- Data Guards & Format Helpers ----------
function normalizeClassName_(c) {
  const s = String(c || '').trim();
  const low = s.toLowerCase();
  if (!low) return '';
  if (low === 'sunday school' || low === 'ss') return 'Sunday School';
  if (low === 'elders quorum' || low === 'eq') return 'Elders Quorum';
  if (low === 'relief society' || low === 'rs') return 'Relief Society';
  if (low === 'gospel foundation' || low === 'gospel foundations' || low === 'gospel found.' || low === 'gospel found' || low === 'gf') return 'Gospel Foundation';
  if (low === 'temple prep' || low === 'temple preparation' || low === 'tp') return 'Temple Prep';
  return s;
}

function assertClass_(c) {
  const norm = normalizeClassName_(c);
  if (CLASSES.indexOf(norm) === -1) throw userError_('Unknown class');
  return norm;
}

function getClassSheet_(ss, className) {
  const norm = normalizeClassName_(className);
  let sheet = ss.getSheetByName(norm);
  if (!sheet && norm === 'Gospel Foundation') {
    sheet = ss.getSheetByName('Gospel Foundations');
  }
  return sheet;
}

function assertDate_(d) {
  if (!DATE_RE.test(String(d || ''))) throw userError_('Invalid date');
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function tz_(ss) {
  return ss.getSpreadsheetTimeZone() || 'America/Denver';
}

function clip_(v, n) {
  return String(v == null ? '' : v).trim().slice(0, n);
}

function cleanPhone_(v) {
  return clip_(v, 40).replace(/[^0-9+()\-.\sx]/gi, '');
}

function cleanEmail_(v) {
  return clip_(v, 100).replace(/[^A-Za-z0-9._%+\-@]/g, '');
}

function safeCell_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function safeHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, m => ({
    '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
  })[m]);
}

function isTrue_(v) {
  return v === true || String(v).toUpperCase() === 'TRUE';
}

function dateKey_(d, tz) {
  if (d instanceof Date && !isNaN(d.getTime())) {
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }
  return String(d || '').trim().split(' ')[0];
}

function readDateColumns_(sheet, tz) {
  const lastCol = sheet.getLastColumn();
  if (lastCol < 4) return [];
  const vals = sheet.getRange(1, 4, 1, lastCol - 3).getValues()[0];
  const out = [];
  for (let c = 0; c < vals.length; c++) {
    const d = vals[c];
    if (d instanceof Date && !isNaN(d.getTime())) {
      out.push({
        key: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
        label: Utilities.formatDate(d, tz, 'MMM d, yyyy'),
        colIndex: c + 4
      });
    } else if (typeof d === 'string' && DATE_RE.test(d.trim().split(' ')[0])) {
      const k = d.trim().split(' ')[0];
      out.push({ key: k, label: k, colIndex: c + 4 });
    }
  }
  return out;
}

function findDateCol_(sheet, dateStr, tz) {
  const m = readDateColumns_(sheet, tz).find(x => x.key === dateStr);
  return m ? m.colIndex : -1;
}

function extractFirstName_(fullName) {
  if (!fullName) return '';
  const s = String(fullName).trim();
  if (s.indexOf(',') > -1) {
    const parts = s.split(',');
    if (parts.length > 1) {
      const f = parts[1].trim().split(' ')[0];
      if (f) return f;
    }
  }
  return s.split(' ')[0];
}

function getOrCreateSheet_(ss, name, init) {
  let sheet = ss.getSheetByName(name);
  if (sheet) return sheet;
  const lock = LockService.getScriptLock();
  const got = lock.tryLock(15000);
  try {
    sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      init(sheet);
    }
  } catch (e) {
    sheet = ss.getSheetByName(name);   // another request created it first
    if (!sheet) throw e;
  } finally {
    if (got) lock.releaseLock();
  }
  return sheet;
}

function ensureVisitorsSheet_(ss) {
  return getOrCreateSheet_(ss, 'Visitors', function(sheet) {
    sheet.appendRow(['Timestamp', 'Visitor Name', 'Class', 'Date', 'Phone Number', 'Email', 'Total Visits', 'Notes']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#E2E8F0');
    sheet.getRange('D:D').setNumberFormat('@');
    sheet.setFrozenRows(1);
  });
}

function ensureCheckInLog_(ss) {
  return getOrCreateSheet_(ss, CHECKIN_LOG, function(sheet) {
    sheet.appendRow(['Timestamp', 'Class', 'Date', 'Member', 'Result']);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#E2E8F0');
    sheet.getRange('C:C').setNumberFormat('@');
    sheet.setFrozenRows(1);
  });
}

// ---------- Web App Entry Points (doGet + doPost) ----------
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (String(p.action || '') === 'scanBoot') {
    let out;
    try {
      const cls = assertClass_(p.className || p.class || 'Sunday School');
      out = Object.assign({ ok: true }, getScanBootData_(cls));
    } catch (err) {
      out = { ok: false, code: 'invalid', message: String((err && err.message) || err) };
    }
    return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
  }

  const rawClass = String(p.class || p.className || '').trim();
  const isScan = Boolean(rawClass);
  let targetClass = '';
  let scanBoot = {};
  if (isScan) {
    try {
      targetClass = assertClass_(rawClass);
    } catch (err) {
      return HtmlService.createHtmlOutput('Unknown class link.');
    }
    try {
      scanBoot = Object.assign({ ok: true }, getScanBootData_(targetClass));
    } catch (err) {
      console.error('Pre-fetch scanBoot failed: ' + err);
      scanBoot = { isScanMode: true, defaultClass: targetClass };
    }
  }

  const template = HtmlService.createTemplateFromFile('Index');
  template.targetClass = targetClass;
  template.scanBootJson = JSON.stringify(scanBoot).replace(/</g, '\u003c');
  return template.evaluate()
    .setTitle(isScan ? targetClass + ' Check-In' : 'Valley View 2nd Hour Attendance')
    .setFaviconUrl('https://drive.google.com/uc?id=1YYFXMz-fEyDiTYG1tDLuPx9679TQLCtI&.png')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover');
}

function doPost(e) {
  let out;
  try {
    const raw = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    out = routeApi_(JSON.parse(raw));
  } catch (err) {
    if (err && (err.name === 'AuthError' || err.code === 'auth')) {
      out = { ok: false, code: 'auth', message: err.message || 'Please sign in again.' };
    } else if (err && err.name === 'UserError') {
      out = { ok: false, code: 'invalid', message: err.message };
    } else {
      const msg = String((err && err.message) || err || 'Something went wrong.');
      console.error('API error: ' + msg);
      out = { ok: false, code: 'error', message: msg };
    }
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function routeApi_(req) {
  const action = String((req && req.action) || '');
  switch (action) {
    case 'requestCode':
      return apiRequestCode_(req);
    case 'verifyCode':
      return apiVerifyCode_(req);

    case 'scanBoot': {
      const targetClass = assertClass_(req.className || req.class || 'Sunday School');
      return Object.assign({ ok: true }, getScanBootData_(targetClass));
    }
    case 'selfCheckInMember': {
      const res = selfCheckInMember(req.className, req.dateStr, req.memberName);
      return Object.assign({ ok: true }, res);
    }
    case 'selfCheckInVisitor': {
      const res = selfCheckInVisitor(req.className, req.dateStr, req.visitorData);
      return Object.assign({ ok: true }, res);
    }

    case 'boot': {
      const user = requireLeader_(req);
      flushCheckInsSafe_(8000);
      const initialClass = assertClass_(req.className || 'Sunday School');
      const boot = getFastBootData_(initialClass, false);
      return Object.assign({ ok: true, me: user }, boot);
    }
    case 'getAttendanceData': {
      requireLeader_(req);
      const cls = assertClass_(req.className);
      assertDate_(req.dateStr);
      flushCheckInsSafe_(5000);
      const data = getRosterFast_(SpreadsheetApp.getActiveSpreadsheet(), cls, req.dateStr, null);
      return Object.assign({ ok: true }, data);
    }
    case 'saveAttendanceData': {
      requireLeader_(req);
      const saved = saveAttendanceData_(req.className, req.dateStr, req.changes);
      return Object.assign({ ok: true }, saved);
    }
    case 'getVisitorsData': {
      requireLeader_(req);
      if (req.dateStr) assertDate_(req.dateStr);
      const vData = getVisitorsFast_(SpreadsheetApp.getActiveSpreadsheet(), req.dateStr);
      return Object.assign({ ok: true }, vData);
    }
    case 'getWeeklyWardReport': {
      requireLeader_(req);
      flushCheckInsSafe_(5000);
      const rep = getWeeklyWardReport_(req.dateStr);
      return Object.assign({ ok: true }, rep);
    }
    default:
      return { ok: false, code: 'bad_request', message: 'Unknown request.' };
  }
}

function callServerApi(req) {
  try {
    return routeApi_(req || {});
  } catch (err) {
    if (err && (err.name === 'AuthError' || err.code === 'auth')) {
      return { ok: false, code: 'auth', message: err.message || 'Please sign in again.' };
    }
    if (err && err.name === 'UserError') {
      return { ok: false, code: 'invalid', message: err.message };
    }
    return { ok: false, code: 'error', message: String((err && err.message) || err) };
  }
}

// ---------- Cached data for the public scan page (built once, shared by every scanner) ----------
function nearestDateKey_(dates) {
  let best = '', minDiff = Infinity;
  (dates || []).forEach(d => {
    const diff = Math.abs(new Date(d.key + 'T12:00:00').getTime() - Date.now());
    if (diff < minDiff) { minDiff = diff; best = d.key; }
  });
  return best;
}

function buildScanBase_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = tz_(ss);
  const ssSheet = ss.getSheetByName('Sunday School');
  if (!ssSheet) throw new Error('Sheet "Sunday School" not found.');
  const dates = readDateColumns_(ssSheet, tz);

  const memberNames = [];
  const index = {};   // lower-case name -> official spelling (Member Info plus anyone already on a class roster)
  const mSheet = ss.getSheetByName('Member Info');
  if (mSheet) {
    const mRows = Math.min(MAX_MEMBERS, mSheet.getLastRow());
    if (mRows >= 2) {
      mSheet.getRange(2, 1, mRows - 1, 1).getValues().forEach(r => {
        const name = String(r[0] || '').trim();
        if (!name) return;
        memberNames.push(name);
        index[name.toLowerCase()] = name;
      });
    }
  }
  memberNames.sort();
  CLASSES.forEach(c => {
    const sh = getClassSheet_(ss, c);
    if (!sh) return;
    const n = Math.min(MAX_ROWS, sh.getLastRow()) - 1;
    if (n < 1) return;
    sh.getRange(2, 1, n, 1).getValues().forEach(r => {
      const nm = String(r[0] || '').trim();
      if (!nm || nm === 'Weekly Total') return;
      if (!index[nm.toLowerCase()]) index[nm.toLowerCase()] = nm;
    });
  });

  const props = PropertiesService.getScriptProperties();
  return {
    dates: dates,
    memberNames: memberNames,
    index: index,
    webAppUrl: ScriptApp.getService().getUrl() || '',
    githubPagesUrl: props.getProperty('GITHUB_PAGES_URL') || ''
  };
}

function getScanBase_() {
  const cache = CacheService.getScriptCache();
  let hit = cache.get(SCAN_CACHE_KEY);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }
  // Only one request rebuilds the cache; the rest wait briefly and then read it.
  const lock = LockService.getScriptLock();
  const got = lock.tryLock(10000);
  try {
    hit = cache.get(SCAN_CACHE_KEY);
    if (hit) {
      try { return JSON.parse(hit); } catch (e) {}
    }
    const base = buildScanBase_();
    try { cache.put(SCAN_CACHE_KEY, JSON.stringify(base), SCAN_CACHE_TTL); } catch (e) {}
    return base;
  } finally {
    if (got) lock.releaseLock();
  }
}

function clearScanCache_() {
  try { CacheService.getScriptCache().remove(SCAN_CACHE_KEY); } catch (e) {}
}

function getScanBootData_(targetClass) {
  const base = getScanBase_();
  return {
    classes: CLASSES,
    dates: base.dates,
    defaultDate: nearestDateKey_(base.dates),
    defaultClass: targetClass || 'Sunday School',
    memberNames: base.memberNames,
    webAppUrl: base.webAppUrl,
    githubPagesUrl: base.githubPagesUrl,
    isScanMode: true
  };
}

// ---------- Boot & Roster Builders ----------
function getFastBootData_(initialClass, isScanMode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = tz_(ss);
  const ssSheet = ss.getSheetByName('Sunday School');
  if (!ssSheet) throw new Error('Sheet "Sunday School" not found.');

  const dates = readDateColumns_(ssSheet, tz);
  let defaultDate = '', minDiff = Infinity;
  dates.forEach(d => {
    const diff = Math.abs(new Date(d.key + 'T12:00:00').getTime() - Date.now());
    if (diff < minDiff) {
      minDiff = diff;
      defaultDate = d.key;
    }
  });

  const memberInfoMap = {}, memberNames = [];
  const mSheet = ss.getSheetByName('Member Info');
  if (mSheet) {
    const mRows = Math.min(MAX_MEMBERS, mSheet.getLastRow());
    if (mRows >= 2) {
      mSheet.getRange(2, 1, mRows - 1, 4).getValues().forEach(r => {
        const name = String(r[0] || '').trim();
        if (!name) return;
        memberNames.push(name);
        memberInfoMap[name] = {
          gender: String(r[1] || '').trim().toUpperCase(),
          phone: cleanPhone_(r[2]),
          email: cleanEmail_(r[3])
        };
      });
    }
  }
  memberNames.sort();

  const props = PropertiesService.getScriptProperties();
  const base = {
    classes: CLASSES,
    dates: dates,
    defaultDate: defaultDate,
    defaultClass: initialClass || 'Sunday School',
    memberNames: memberNames,
    webAppUrl: ScriptApp.getService().getUrl() || '',
    githubPagesUrl: props.getProperty('GITHUB_PAGES_URL') || ''
  };
  if (isScanMode) return Object.assign(base, { isScanMode: true });

  const reportContacts = [];
  const rcSheet = ss.getSheetByName('Report Contacts');
  if (rcSheet) {
    const rcRows = Math.min(100, rcSheet.getLastRow());
    if (rcRows >= 2) {
      rcSheet.getRange(2, 1, rcRows - 1, 4).getValues().forEach(r => {
        const calling = String(r[0] || '').trim(), name = String(r[1] || '').trim();
        if (calling && name) {
          reportContacts.push({
            calling: calling,
            name: name,
            phone: cleanPhone_(r[2]),
            email: cleanEmail_(r[3])
          });
        }
      });
    }
  }
  ensureVisitorsSheet_(ss);
  return Object.assign(base, {
    memberInfoMap: memberInfoMap,
    reportContacts: reportContacts,
    isScanMode: false,
    initialRoster: getRosterFast_(ss, initialClass || 'Sunday School', defaultDate, memberInfoMap),
    visitorsData: getVisitorsFast_(ss, defaultDate)
  });
}

function getRosterFast_(ss, className, dateStr, memberInfoMap) {
  const normClass = normalizeClassName_(className);
  const empty = { members: [], stats: { total: 0, present: 0, absent: 0, dayRate: 0, avgRate: 0 } };
  const sheet = getClassSheet_(ss, normClass);
  if (!sheet) return empty;
  const col = findDateCol_(sheet, dateStr, tz_(ss));
  const n = Math.min(MAX_ROWS, sheet.getLastRow()) - 1;
  if (col === -1 || n < 1) return empty;

  if (!memberInfoMap) {
    memberInfoMap = {};
    const mSheet = ss.getSheetByName('Member Info');
    if (mSheet) {
      const mRows = Math.min(MAX_MEMBERS, mSheet.getLastRow());
      if (mRows >= 2) {
        mSheet.getRange(2, 1, mRows - 1, 4).getValues().forEach(r => {
          const nm = String(r[0] || '').trim();
          if (nm) {
            memberInfoMap[nm] = {
              gender: String(r[1] || '').trim().toUpperCase(),
              phone: cleanPhone_(r[2]),
              email: cleanEmail_(r[3])
            };
          }
        });
      }
    }
  }

  const meta = sheet.getRange(2, 1, n, 3).getValues();
  const att = sheet.getRange(2, col, n, 1).getValues();
  const members = [];
  let present = 0, sumRates = 0;
  for (let i = 0; i < meta.length; i++) {
    const name = String(meta[i][0] || '').trim();
    if (!name || name === 'Weekly Total') continue;
    const here = isTrue_(att[i][0]);
    if (here) present++;
    const rate = Number(meta[i][2]) || 0;
    sumRates += rate;
    const c = (memberInfoMap && memberInfoMap[name]) || { gender: '', phone: '', email: '' };
    members.push({
      name: name,
      gender: c.gender,
      phone: c.phone,
      email: c.email,
      attended: here,
      totalAttended: Number(meta[i][1]) || 0,
      rate: rate
    });
  }
  members.sort((a, b) => a.name.localeCompare(b.name));
  const total = members.length;
  return {
    className: normClass,
    dateStr: dateStr,
    members: members,
    stats: {
      total: total,
      present: present,
      absent: total - present,
      dayRate: total ? present / total : 0,
      avgRate: total ? sumRates / total : 0
    }
  };
}

function saveAttendanceData_(className, dateStr, changes) {
  const normClass = assertClass_(className);
  assertDate_(dateStr);
  changes = changes || {};
  return withLock_(() => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    try { flushCheckIns_(); } catch (e) { console.error('Pre-save flush failed: ' + e); }
    const sheet = getClassSheet_(ss, normClass);
    if (!sheet) throw new Error('Sheet "' + normClass + '" not found.');
    const col = findDateCol_(sheet, dateStr, tz_(ss));
    if (col === -1) throw new Error('Date ' + dateStr + ' column not found.');
    const n = Math.min(MAX_ROWS, sheet.getLastRow()) - 1;
    if (n < 1) return { success: true, className: normClass, dateStr: dateStr, present: 0, total: 0, rate: 0 };

    const names = sheet.getRange(2, 1, n, 1).getValues();
    const rng = sheet.getRange(2, col, n, 1);
    const vals = rng.getValues(), forms = rng.getFormulas();
    let present = 0, total = 0;
    const out = vals.map((v, i) => {
      const nm = String(names[i][0] || '').trim();
      if (!nm || nm === 'Weekly Total') return [forms[i][0] || v[0]];
      total++;
      const val = Object.prototype.hasOwnProperty.call(changes, nm) ? !!changes[nm] : isTrue_(v[0]);
      if (val) present++;
      return [val];
    });
    rng.setValues(out);
    return {
      success: true,
      className: normClass,
      dateStr: dateStr,
      present: present,
      total: total,
      rate: total ? present / total : 0
    };
  });
}

// ---------- Self Check-In (Public QR / NFC) — built for crowds ----------
// Member check-ins are validated against the cached roster and appended to the "Check-In Log" sheet.
// No sheet-wide lock is taken here. flushCheckIns_() merges the log into the class sheets in one batch.
function selfCheckInMember(className, dateStr, memberName) {
  const normClass = assertClass_(className);
  assertDate_(dateStr);
  const name = clip_(memberName, 100);
  if (!name) throw userError_('Name required');

  const base = getScanBase_();
  const official = base.index[name.toLowerCase()];
  if (!official) throw userError_('Member not found on ward list. Please use "I\'m a Visitor".');
  if (!(base.dates || []).some(d => d.key === dateStr)) throw userError_('That date is not on the schedule.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCheckInLog_(ss).appendRow([new Date(), normClass, dateStr, official, '']);
  return {
    success: true,
    memberName: official,
    firstName: extractFirstName_(official),
    className: normClass,
    dateStr: dateStr,
    queued: true
  };
}

function selfCheckInVisitor(className, dateStr, visitorData) {
  const normClass = assertClass_(className);
  assertDate_(dateStr);
  visitorData = visitorData || {};
  const name = clip_(visitorData.name, 80);
  if (!name) throw userError_('Please enter your name.');
  const phone = cleanPhone_(visitorData.phone);
  const email = cleanEmail_(visitorData.email);
  const notes = clip_(visitorData.notes, 200);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = tz_(ss);
  const sheet = ensureVisitorsSheet_(ss);
  const last = Math.min(2000, sheet.getLastRow());
  const rows = last > 1 ? sheet.getRange(2, 1, last - 1, 4).getValues() : [];

  const priorDates = new Set();
  for (let i = 0; i < rows.length; i++) {
    const rName = String(rows[i][1] || '').trim().toLowerCase();
    if (rName !== name.toLowerCase()) continue;
    const rClass = normalizeClassName_(rows[i][2]);
    const rDate = dateKey_(rows[i][3], tz);
    if (rDate === dateStr && rClass === normClass) {
      return {
        success: true,
        visitorName: name,
        firstName: extractFirstName_(name),
        totalVisits: priorDates.size || 1,
        alreadyCheckedIn: true
      };
    }
    if (rDate) priorDates.add(rDate);
  }
  priorDates.add(dateStr);
  const totalVisits = priorDates.size;
  // appendRow is a single atomic append, so simultaneous visitors cannot overwrite each other.
  sheet.appendRow([new Date(), safeCell_(name), normClass, dateStr, safeCell_(phone), safeCell_(email), totalVisits, safeCell_(notes)]);
  return {
    success: true,
    visitorName: name,
    firstName: extractFirstName_(name),
    totalVisits: totalVisits,
    alreadyCheckedIn: false
  };
}

// ---------- Check-In Log -> class sheets (batch merge) ----------
function flushCheckInsSafe_(waitMs) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(waitMs || 5000)) return false;
    try {
      flushCheckIns_();
      return true;
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    console.error('Check-in sync failed: ' + e);
    return false;
  }
}

// Caller must hold the script lock.
function flushCheckIns_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(CHECKIN_LOG);
  if (!log) return 0;
  const props = PropertiesService.getScriptProperties();
  const last = log.getLastRow();
  let cursor = Number(props.getProperty('CHECKIN_CURSOR') || 1);
  if (cursor > last) cursor = 1;           // log was cleared by hand
  if (last <= cursor) return 0;

  const tz = tz_(ss);
  const count = last - cursor;
  const rows = log.getRange(cursor + 1, 1, count, 4).getValues();
  const results = rows.map(() => ['ok']);
  const groups = {};
  rows.forEach((r, i) => {
    const cls = normalizeClassName_(r[1]);
    const dateStr = dateKey_(r[2], tz);
    const name = String(r[3] || '').trim();
    if (!name || CLASSES.indexOf(cls) === -1 || !DATE_RE.test(dateStr)) {
      results[i] = ['skipped: unreadable row'];
      return;
    }
    const gk = cls + '|' + dateStr;
    const grp = groups[gk] || (groups[gk] = { cls: cls, dateStr: dateStr, items: {} });
    const k = name.toLowerCase();
    if (!grp.items[k]) grp.items[k] = { name: name, rows: [] };
    grp.items[k].rows.push(i);
  });

  // A transient failure here throws before the cursor moves, so the same rows are retried next run.
  Object.keys(groups).forEach(gk => applyCheckInGroup_(ss, tz, groups[gk], results));

  log.getRange(cursor + 1, 5, count, 1).setValues(results);
  props.setProperty('CHECKIN_CURSOR', String(last));
  return count;
}

function applyCheckInGroup_(ss, tz, grp, results) {
  const items = Object.keys(grp.items).map(k => grp.items[k]);
  const failAll = msg => items.forEach(it => it.rows.forEach(i => { results[i] = [msg]; }));

  const sheet = getClassSheet_(ss, grp.cls);
  if (!sheet) return failAll('skipped: class sheet not found');
  const col = findDateCol_(sheet, grp.dateStr, tz);
  if (col === -1) return failAll('skipped: date column not found');

  const n = Math.min(MAX_ROWS, sheet.getLastRow()) - 1;
  const idx = {};
  let rng = null, vals = [], forms = [];
  if (n > 0) {
    sheet.getRange(2, 1, n, 1).getValues().forEach((r, i) => {
      const k = String(r[0] || '').trim().toLowerCase();
      if (k && k !== 'weekly total' && !(k in idx)) idx[k] = i;
    });
    rng = sheet.getRange(2, col, n, 1);
    vals = rng.getValues();
    forms = rng.getFormulas();
  }

  const touched = {}, missing = [];
  let needsWrite = false;
  items.forEach(it => {
    const k = it.name.toLowerCase();
    if (k in idx) {
      touched[idx[k]] = true;
      if (!isTrue_(vals[idx[k]][0])) needsWrite = true;
    } else {
      missing.push(it);
    }
  });
  if (needsWrite) {
    rng.setValues(vals.map((v, i) => touched[i] ? [true] : [forms[i][0] || v[0]]));
  }

  if (missing.length) {          // on the ward list but not yet on this class sheet: add them (as before)
    let newRow = sheet.getLastRow() + 1;
    missing.forEach(it => {
      sheet.getRange(newRow, 1).setValue(safeCell_(it.name));
      if (newRow >= 3) sheet.getRange(2, 2, 1, 2).copyTo(sheet.getRange(newRow, 2, 1, 2));
      sheet.getRange(newRow, col).setValue(true);
      it.rows.forEach(i => { results[i] = ['ok (added to class sheet)']; });
      newRow++;
    });
  }
}

// ---------- Auto-sync trigger + menu helpers ----------
function autoSyncCheckIns() {
  const lock = LockService.getScriptLock();
  if (lock.tryLock(20000)) {
    try { flushCheckIns_(); } finally { lock.releaseLock(); }
  }
  try { getScanBase_(); } catch (e) {}   // keeps the scan-page cache warm
}

function installAutoSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureCheckInLog_(ss);
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoSyncCheckIns') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('autoSyncCheckIns').timeBased().everyMinutes(1).create();
  try {
    SpreadsheetApp.getUi().alert('Auto-Sync is on. Self check-ins now merge into the class sheets every minute (and instantly whenever a leader opens or refreshes the app).');
  } catch (e) {}
}

function removeAutoSync() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'autoSyncCheckIns') ScriptApp.deleteTrigger(t);
  });
  try { SpreadsheetApp.getUi().alert('Auto-Sync is off. Check-ins still sync whenever a leader opens or refreshes the app.'); } catch (e) {}
}

// Removes the fake "ZZ Load Test ..." visitors created by loadtest-console.js
function cleanupLoadTestVisitors() {
  const sheet = ensureVisitorsSheet_(SpreadsheetApp.getActiveSpreadsheet());
  const last = sheet.getLastRow();
  let removed = 0;
  if (last >= 2) {
    const names = sheet.getRange(2, 2, last - 1, 1).getValues();
    for (let i = names.length - 1; i >= 0; i--) {
      if (/^ZZ Load Test/i.test(String(names[i][0] || ''))) {
        sheet.deleteRow(i + 2);
        removed++;
      }
    }
  }
  try { SpreadsheetApp.getUi().alert('Removed ' + removed + ' load-test visitor row(s).'); } catch (e) {}
  return removed;
}

function syncCheckInsNow() {
  const ok = flushCheckInsSafe_(20000);
  try {
    SpreadsheetApp.getUi().alert(ok ? 'Check-ins synced to the class sheets.' : 'Another update was in progress. Please try again in a moment.');
  } catch (e) {}
}

// ---------- Visitors & Ward Report Builders ----------
function getVisitorsFast_(ss, dateStr) {
  const tz = tz_(ss);
  const sheet = ensureVisitorsSheet_(ss);
  const last = Math.min(2000, sheet.getLastRow());
  if (last <= 1) return { dateVisitors: [], frequentVisitors: [], allRecent: [] };

  const rows = sheet.getRange(2, 1, last - 1, 8).getValues();
  const byVisitor = {}, dateVisitors = [], allRecent = [], seenVisit = {};
  rows.forEach(r => {
    const name = String(r[1] || '').trim();
    if (!name) return;
    const vClass = String(r[2] || '').trim();
    const vDate = dateKey_(r[3], tz);
    const visitKey = name.toLowerCase() + '|' + normalizeClassName_(vClass) + '|' + vDate;
    if (seenVisit[visitKey]) return;       // same person double-tapped / retried
    seenVisit[visitKey] = true;
    const phone = String(r[4] || '').trim();
    const email = String(r[5] || '').trim();
    const notes = String(r[7] || '').trim();
    const key = name.toLowerCase();
    if (!byVisitor[key]) {
      byVisitor[key] = { name: name, phone: phone, email: email, dates: new Set(), classes: new Set(), lastDate: vDate, notes: notes };
    }
    if (vDate) byVisitor[key].dates.add(vDate);
    if (vClass) byVisitor[key].classes.add(vClass);
    if (phone) byVisitor[key].phone = phone;
    if (email) byVisitor[key].email = email;
    if (vDate >= byVisitor[key].lastDate) byVisitor[key].lastDate = vDate;

    const item = { name: name, className: vClass, dateStr: vDate, phone: phone, email: email, notes: notes };
    allRecent.push(item);
    if (!dateStr || vDate === dateStr) dateVisitors.push(item);
  });

  const countFor = nm => (byVisitor[nm.toLowerCase()] ? byVisitor[nm.toLowerCase()].dates.size : 1);
  dateVisitors.forEach(v => (v.totalVisits = countFor(v.name)));
  allRecent.forEach(v => (v.totalVisits = countFor(v.name)));

  const frequentVisitors = Object.values(byVisitor)
    .map(v => ({
      name: v.name,
      phone: v.phone,
      email: v.email,
      totalVisits: v.dates.size,
      classes: Array.from(v.classes).join(', '),
      lastDate: v.lastDate,
      notes: v.notes
    }))
    .filter(v => v.totalVisits > 1)
    .sort((a, b) => b.totalVisits - a.totalVisits);

  return {
    dateVisitors: dateVisitors,
    frequentVisitors: frequentVisitors,
    allRecent: allRecent.reverse().slice(0, 40)
  };
}

function getWeeklyWardReport_(dateStr) {
  assertDate_(dateStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalPresent = 0, totalEnrolled = 0;
  const summary = CLASSES.map(c => {
    const r = getRosterFast_(ss, c, dateStr, null);
    totalPresent += r.stats.present;
    if (c === 'Sunday School') totalEnrolled = r.stats.total;
    return {
      className: c,
      present: r.stats.present,
      total: r.stats.total,
      rate: r.stats.dayRate
    };
  });
  const v = getVisitorsFast_(ss, dateStr);
  return {
    dateStr: dateStr,
    summary: summary,
    visitorsCount: v.dateVisitors.length,
    visitorsList: v.dateVisitors,
    totalPresent: totalPresent,
    totalEnrolled: totalEnrolled
  };
}
