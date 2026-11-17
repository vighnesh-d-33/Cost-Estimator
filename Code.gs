/**
 * Google Apps Script backend – Metal printing cost estimator  (USER NAME + PASSWORD LOGIN)
 * Storage: the Google Sheet this script is bound to (Extensions > Apps Script).
 * Tabs created automatically: Quotes, Images, Powders, Machines, Settings, Users, Audit
 *
 * SECURITY (read SECURITY.md):
 *  - Staff sign in with a user name and password. Passwords are stored only as salted hashes in the Users tab.
 *  - A successful login returns a random session token (kept in the script cache, 30 min idle timeout). Every other request needs it.
 *  - 5 wrong passwords for one user name in 10 minutes lock that user name for 10 minutes (40 wrong tries overall lock everyone).
 *  - Only POST requests are accepted for data; GET returns a plain "running" message and no data.
 *  - Input is size-limited and validated; text that could run as a spreadsheet formula (= + - @) is neutralised.
 *  - Logins, saves, deletes and master changes are written to the Audit tab with the user name.
 *
 * ADD USERS: see addUser() below (run it from the Apps Script editor – there is no sign-up page).
 * Deploy: Deploy > New deployment > Web app.  Execute as: Me.  Who has access: Anyone
 * After EVERY change to this file: Deploy > Manage deployments > pencil > Version: New version > Deploy.
 */
const QUOTE_HEADERS = ['ID', 'Saved', 'Customer', 'Part', 'Machine', 'Material', 'Laser', 'Parts/Build', 'Layer (um)', 'Print hrs',
  'Material kg', 'Sets', 'Units', 'Total cost', 'Round-off', 'Per part', 'State JSON', 'Saved by'];
const POWDER_HEADERS = ['Powder', 'Price / kg', 'Density g/cc', 'Wastage', 'HT time (hrs)', 'HT power (hrs)'];
const POWDER_KEYS = ['n', 'p', 'dens', 'waste', 'ht', 'fp'];
const MACHINE_HEADERS = ['Machine', 'Lasers', 'Price', 'Interest %', 'Years', 'Utilisation %', 'Fixed rate per hr'];
const MACHINE_KEYS = ['n', 'laser', 'price', 'rate', 'yrs', 'util', 'fixed'];
const USER_HEADERS = ['Username', 'Salt', 'Password hash', 'Role', 'Active', 'Created', 'Email', 'L1 name', 'L1 email', 'L2 name', 'L2 email', 'L3 name', 'L3 email'];
const APPROVAL_HEADERS = ['Quote ID', 'Status', 'Level', 'Prepared by', 'Prepared email', 'Customer', 'Part', 'Per part', 'Round-off', 'L1 name', 'L1 email', 'L2 name', 'L2 email', 'L3 name', 'L3 email', 'Token', 'Sent', 'Log JSON', 'State hash'];
const SITE_URL = '';   // <-- type the website address here (shown in e-mails as a link back to the site)
const EMAIL_RE = /^[^\s@<>",;]+@[^\s@<>",;]+\.[A-Za-z]{2,}$/;
const SETTINGS_HEADERS = ['Group', 'Key', 'Value'];
const IMAGE_HEADERS = ['Quote ID', 'Part', 'Data'];
const MAX_IMG = 150000, IMG_CHUNK = 40000, IMG_RE = /^data:image\/jpeg;base64,[A-Za-z0-9+\/=]+$/;   // model picture of a quote: small JPEG, stored in pieces in the Images tab
/* Default prices / rates live ONLY here (and then in the Sheet). They are copied into the Powders, Machines and Settings tabs the first time
   those tabs are empty, and are sent to the website only after a successful login. Edit the Sheet afterwards, not this code. */
const DEFAULT_POWDERS = [['SS304', 3200, 8, .3], ['AlSi10Mg', 4500, 2.7, .3], ['CuCrZr', 12000, 8.9, 0], ['SS316L', 2400, 8, .3], ['18Ni300', 5200, 8.1, .5], ['H13', 4500, 8.1, .5],
  ['StrengthAl', 15000, 2.5, 0], ['17-4PH', 2600, 7.8, .5], ['IN718', 8000, 8.1, .5], ['IN625', 8000, 8.1, .5], ['Ti6Al4V', 18500, 4.4, .6], ['CoCr', 8500, 10, .3]].map(r => r.concat([5.5, 3.5]));
const DEFAULT_MACHINES = [['STLR 220', 'Single laser', 12850000, 10.5, 5, 80, ''], ['STLR180', '', '', '', '', '', 1500], ['STLR400-1', 'Dual laser', '', '', '', '', 3000],
  ['STLR400-2', 'Dual laser', '', '', '', '', 3500], ['SLM', '', '', '', '', '', 3000]];
const DEFAULT_SETTINGS = {
  i: { prep: 3, wire: 10, supHrs: 160, supRate: 100, filter: 150, recoater: 525, kw: 4, pRate: 8, gasF: 1, gRate: 8, opF: 0.5, opRate: 150, margin: 15, adv: 50, pf: 750 },
  furn: { price: 1000000, rate: 10.5, yrs: 5, util: 90, kwh: 25, kwhRate: 8, other: 30 },
  wire: { price: 1300000, rate: 10.5, yrs: 5, util: 85, cons: 100, op: 150, pwr: 8 } };
const MAX_BODY = 300000, MAX_STATE = 45000, MAX_POWDERS = 200, MAX_MACHINES = 50;
const MAX_USER_FAILS = 5, MAX_ALL_FAILS = 40, LOCK_SECONDS = 600;   // lockout limits
const MAX_REQ_PER_MIN = 120;                                        // per signed-in user
const IDLE_SECONDS = 1800, SESSION_SECONDS = 21600;                 // 30 min idle sign-out (cache keeps a session at most 6 h after last use)
const HASH_ROUNDS = 300;                                           // password hash iterations (raise if logins stay fast)
const USER_RE = /^[a-z0-9._@-]{3,40}$/;
const ADMIN_EMAIL = '';   // <-- type the admin's e-mail here. Only this address receives the "forgot password" verification code.
const CODE_SECONDS = 600, MAX_CODE_REQUESTS = 3, MAX_CODE_TRIES = 5;
let CURRENT_USER = '';

function sheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(headers); sh.setFrozenRows(1); sh.getRange(1, 1, 1, headers.length).setFontWeight('bold'); }
  return sh;
}
function setup() { sheet_('Quotes', QUOTE_HEADERS); sheet_('Powders', POWDER_HEADERS); sheet_('Machines', MACHINE_HEADERS); sheet_('Users', USER_HEADERS); sheet_('Settings', SETTINGS_HEADERS); sheet_('Images', IMAGE_HEADERS); sheet_('Audit', ['Time', 'Action', 'Detail', 'User']); sheet_('Approvals', APPROVAL_HEADERS); }   // run once to authorise
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function audit_(action, detail, user) { try { sheet_('Audit', ['Time', 'Action', 'Detail', 'User']).appendRow([new Date(), action, String(detail || '').slice(0, 200), String(user || CURRENT_USER || '').slice(0, 60)]); } catch (e) { /* never block the request */ } }

/* ---------- sanitising ---------- */
function cell_(v, max) {          // value going into a sheet cell
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return isFinite(v) ? v : '';
  let t = String(v).slice(0, max || 500);
  if (/^[=+\-@\t\r]/.test(t)) t = "'" + t;      // spreadsheet-formula injection guard
  return t;
}
function num_(v) { v = (v === '' || v === null || v === undefined) ? null : +v; return v === null ? null : (isFinite(v) ? v : null); }
const ID_RE = /^Q\d{10,15}$/;

/* ---------- users & passwords ---------- */
function same_(a, b) {            // constant-time string compare
  a = String(a); b = String(b);
  let r = a.length ^ b.length;
  for (let i = 0; i < Math.max(a.length, b.length); i++) r |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return r === 0;
}
function pwOk_(p) { return typeof p === 'string' && p.length >= 10 && p.length <= 100 && /[A-Za-z]/.test(p) && /\d/.test(p); }
const PW_MSG = 'Password needs at least 10 characters with letters and numbers';
function hex_(bytes) { return bytes.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join(''); }
function hash_(salt, password) {  // salted, iterated SHA-256 (hex)
  const saltBytes = Utilities.newBlob(salt).getBytes();
  let h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, saltBytes.concat(Utilities.newBlob(password).getBytes()));
  for (let i = 0; i < HASH_ROUNDS; i++) h = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h.concat(saltBytes));
  return hex_(h);
}
function findUser_(username) {
  const sh = sheet_('Users', USER_HEADERS), rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === username)
      return { row: i + 1, salt: String(rows[i][1]), hash: String(rows[i][2]), role: String(rows[i][3] || 'user'), active: String(rows[i][4]).trim().toLowerCase() !== 'no', email: String(rows[i][6] || ''),
        ap: [[7, 8], [9, 10], [11, 12]].map(c => ({ name: String(rows[i][c[0]] || ''), email: String(rows[i][c[1]] || '') })) };
  }
  return null;
}
function setUser_(username, password, role, by, email, ap) {      // create the user, or reset password / role and re-enable
  username = String(username || '').trim().toLowerCase();
  if (!USER_RE.test(username)) throw new Error('User name: 3-40 characters, only a-z 0-9 . _ @ -');
  if (!pwOk_(password)) throw new Error(PW_MSG);
  role = role === 'admin' || role === 'approver' ? role : 'user';
  const salt = Utilities.getUuid().replace(/-/g, ''), hash = hash_(salt, password), sh = sheet_('Users', USER_HEADERS), u = findUser_(username);
  const extra = [email || '']; (ap || [{}, {}, {}]).forEach(a => { extra.push((a && a.name) || '', (a && a.email) || ''); });
  if (u) { sh.getRange(u.row, 1, 1, 5).setValues([[username, salt, hash, role, 'Yes']]); userCols_(sh); if (email) sh.getRange(u.row, 7).setValue(email); if (ap) sh.getRange(u.row, 8, 1, 6).setValues([extra.slice(1)]); }
  else { userCols_(sh); sh.appendRow([username, salt, hash, role, 'Yes', new Date()].concat(extra)); }
  audit_('USER_SET', username + ' (' + role + ')', by || 'editor');
}

/* ---------- run these from the Apps Script editor (Run button) ---------- */
function addUser() {              // 1) type the user name, password and role below  2) select addUser and click Run  3) delete the password again and save
  const USERNAME = '', PASSWORD = '', ROLE = 'user';   // ROLE: 'admin' or 'user'
  setUser_(USERNAME, PASSWORD, ROLE);   // adds the user, or resets the password if the user name already exists
}
function addAdmin() {             // creates the admin user. Type a password below, select addAdmin, click Run, then DELETE the password from the code and save
  const USERNAME = 'admin', PASSWORD = '', ROLE = 'admin';
  if (!pwOk_(PASSWORD)) throw new Error('Type a password (10+ characters, letters and numbers) between the quotes after PASSWORD =');
  setUser_(USERNAME, PASSWORD, ROLE);
  Logger.log('Admin user "' + USERNAME + '" saved. Now delete the password from this code.');
}
function disableUser() {          // type the user name, select disableUser and click Run
  const USERNAME = '';
  const u = findUser_(String(USERNAME).trim().toLowerCase()); if (!u) throw new Error('No such user');
  sheet_('Users', USER_HEADERS).getRange(u.row, 5).setValue('No'); audit_('USER_DISABLED', USERNAME, 'editor');
}

function checkAdminReset() {      // select this in the editor, click Run, then open "Execution log" at the bottom
  const USERNAME = 'admin';       // the user name you type on the sign-in page
  const u = findUser_(USERNAME);
  Logger.log(!u ? 'PROBLEM: no user named "' + USERNAME + '" in the Users tab' : 'User found. Role = ' + u.role + ', active = ' + u.active + ((u.role === 'admin' && u.active) ? '  (OK)' : '  -> PROBLEM: it must be role admin and active'));
  MailApp.sendEmail(ADMIN_EMAIL, 'Cost estimator - test e-mail', 'If you can read this, e-mail sending works.');
  Logger.log('Test e-mail sent to ' + ADMIN_EMAIL + '. Mails you can still send today: ' + MailApp.getRemainingDailyQuota());
}

/* ---------- authentication ---------- */
function authFail_(code, msg) { return { ok: false, code: code, error: msg }; }
function login_(d) {
  const cache = CacheService.getScriptCache(), u = String(d.username || '').trim().toLowerCase(), p = typeof d.password === 'string' ? d.password : '';
  const ukey = 'fails:' + (USER_RE.test(u) ? u : '-'), uf = +(cache.get(ukey) || 0), af = +(cache.get('fails:*') || 0);
  if (uf >= MAX_USER_FAILS || af >= MAX_ALL_FAILS) return authFail_('LOCKED', 'Too many wrong attempts. Try again in 10 minutes.');
  const un = resolveUser_(u), rec = un ? findUser_(un) : null;
  const calc = hash_(rec ? rec.salt : 'x', p);                      // always hash, so a missing user takes the same time
  if (!rec || !rec.active || !same_(calc, rec.hash)) {
    cache.put(ukey, String(uf + 1), LOCK_SECONDS); cache.put('fails:*', String(af + 1), LOCK_SECONDS); Utilities.sleep(1000);
    audit_('LOGIN_FAILED', u.slice(0, 40), u.slice(0, 40));
    return authFail_('AUTH', 'Wrong user name or password.');
  }
  cache.remove(ukey);
  const token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');   // 64 hex characters
  cache.put('s:' + token, JSON.stringify({ u: un, role: rec.role, t: Date.now(), h: rec.hash.slice(0, 16) }), SESSION_SECONDS);
  audit_('LOGIN', '', un);
  return { ok: true, token: token, user: un, role: rec.role };
}
function rate_(u) {              // simple per-user request limit (per minute)
  const c = CacheService.getScriptCache(), k = 'r:' + u + ':' + Math.floor(Date.now() / 60000), n = +(c.get(k) || 0) + 1;
  c.put(k, String(n), 90); return n <= MAX_REQ_PER_MIN;
}
function session_(token) {        // returns {u, role} or null
  if (typeof token !== 'string' || !/^[0-9a-f]{64}$/.test(token)) return null;
  const cache = CacheService.getScriptCache(), raw = cache.get('s:' + token); if (!raw) return null;
  let s; try { s = JSON.parse(raw); } catch (e) { return null; }
  const rec = findUser_(s.u);
  if (!rec || !rec.active || s.h !== rec.hash.slice(0, 16) || Date.now() - s.t > IDLE_SECONDS * 1000) { cache.remove('s:' + token); return null; }   // idle too long, user disabled, or password changed
  s.role = rec.role; s.t = Date.now(); cache.put('s:' + token, JSON.stringify(s), SESSION_SECONDS);
  return s;
}

/* ---------- forgot password (admin accounts only, code goes ONLY to ADMIN_EMAIL) ---------- */
function maskEmail_(e) { const p = String(e).split('@'); return (p[0].slice(0, 2) + '***@' + (p[1] || '')); }
function forgotRequest_(d) {
  if (!ADMIN_EMAIL || ADMIN_EMAIL.indexOf('@') < 1) return authFail_('AUTH', 'Admin e-mail is not set. Type it into ADMIN_EMAIL in Code.gs and redeploy.');
  const cache = CacheService.getScriptCache(), n = +(cache.get('fp:req') || 0), u = String(d.username || '').trim().toLowerCase();
  if (n >= MAX_CODE_REQUESTS) return authFail_('LOCKED', 'Too many requests. Try again in 10 minutes.');
  cache.put('fp:req', String(n + 1), CODE_SECONDS);
  const un = resolveUser_(u), rec = un ? findUser_(un) : null;
  if (rec && rec.active && rec.role === 'admin') {                      // same reply either way, so nobody can discover which names exist
    const code = ('000000' + (parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 1000000)).slice(-6), salt = Utilities.getUuid().replace(/-/g, '');
    cache.put('fp:' + u, JSON.stringify({ s: salt, h: hash_(salt, code), n: 0 }), CODE_SECONDS);
    MailApp.sendEmail(ADMIN_EMAIL, 'Cost estimator – password reset code', 'Your verification code is ' + code + '\n\nIt is valid for 10 minutes. If you did not ask for it, ignore this e-mail and nothing changes.');
    audit_('RESET_CODE_SENT', u, u);
  } else audit_('RESET_NOT_SENT', 'no active admin user named: ' + u.slice(0, 40), u.slice(0, 40));   // visible only to you in the Audit tab
  return { ok: true, message: 'If this is an admin account, a 6-digit code has been sent to ' + maskEmail_(ADMIN_EMAIL) + '. It is valid for 10 minutes.' };
}
function forgotReset_(d) {
  const cache = CacheService.getScriptCache(), u = String(d.username || '').trim().toLowerCase(), code = String(d.code || '').trim(), pw = d.password;
  if (!pwOk_(pw)) return { ok: false, error: PW_MSG };
  const raw = USER_RE.test(u) ? cache.get('fp:' + u) : null, rec = raw ? findUser_(u) : null;
  if (!raw || !rec || !rec.active || rec.role !== 'admin') return authFail_('AUTH', 'Code expired or not requested. Click Forgot password again.');
  const e = JSON.parse(raw);
  if (e.n >= MAX_CODE_TRIES) { cache.remove('fp:' + u); return authFail_('AUTH', 'Too many wrong codes. Request a new code.'); }
  if (!/^\d{6}$/.test(code) || !same_(hash_(e.s, code), e.h)) {
    e.n++; cache.put('fp:' + u, JSON.stringify(e), CODE_SECONDS); Utilities.sleep(1000);
    audit_('RESET_CODE_WRONG', u, u); return authFail_('AUTH', 'Wrong code.');
  }
  setUser_(u, pw, 'admin', u); cache.remove('fp:' + u); cache.remove('fails:' + u);
  audit_('PASSWORD_RESET', 'by e-mail code', u);
  return { ok: true, message: 'Password changed. Sign in with the new password.' };
}

/* ---------- master data: first-time fill from the defaults above, then everything is read from the Sheet ---------- */
function seed_(name, headers, rows) {
  const sh = sheet_(name, headers);
  if (sh.getLastRow() < 2 && rows.length) sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  return sh;
}
function settings_() {
  const flat = []; Object.keys(DEFAULT_SETTINGS).forEach(gp => Object.keys(DEFAULT_SETTINGS[gp]).forEach(k => flat.push([gp, k, DEFAULT_SETTINGS[gp][k]])));
  const rows = seed_('Settings', SETTINGS_HEADERS, flat).getDataRange().getValues().slice(1), out = { i: {}, furn: {}, wire: {} };
  rows.forEach(r => { const gp = String(r[0]), k = String(r[1]); if (out[gp] && Object.prototype.hasOwnProperty.call(DEFAULT_SETTINGS[gp], k) && num_(r[2]) !== null) out[gp][k] = num_(r[2]); });   // only known keys, numbers only
  return out;
}

/* ---------- model picture of a quote (stored as base64 pieces in the Images tab; each piece starts with a letter so Sheets never reads it as a formula) ---------- */
function deleteImg_(id) {
  const sh = sheet_('Images', IMAGE_HEADERS), n = sh.getLastRow(); if (n < 2) return;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) sh.deleteRow(i + 2);
}
function saveImg_(id, url) {
  deleteImg_(id); if (!url) return;
  const sh = sheet_('Images', IMAGE_HEADERS), b = url.slice(url.indexOf(',') + 1), rows = [];
  for (let p = 0, k = 0; p < b.length; p += IMG_CHUNK, k++) rows.push([id, k, 'A' + b.slice(p, p + IMG_CHUNK)]);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}
function getImg_(id) {
  const sh = sheet_('Images', IMAGE_HEADERS), n = sh.getLastRow(); if (n < 2) return '';
  const rows = sh.getRange(2, 1, n - 1, 3).getValues().filter(r => String(r[0]) === id).sort((a, b) => a[1] - b[1]);
  return rows.length ? 'data:image/jpeg;base64,' + rows.map(r => String(r[2]).slice(1)).join('') : '';
}
function imgIds_() {
  const sh = sheet_('Images', IMAGE_HEADERS), n = sh.getLastRow(), set = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 1).getValues().forEach(r => { set[String(r[0])] = 1; });
  return set;
}


/* ---------- approval workflow (3 levels; approvers sign in to the website) ---------- */
function userCols_(sh) {          // older Users sheets get the new approver column headers
  if (sh.getRange(1, 7).getValue() !== 'Email') { sh.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]).setFontWeight('bold'); }
}
function cleanAp_(ap) {           // validates 3 approvers {name,email}; returns array or throws
  if (!Array.isArray(ap) || ap.length !== 3) throw new Error('Three approvers are needed');
  return ap.map((a, i) => {
    const name = String((a && a.name) || '').trim().slice(0, 60), email = String((a && a.email) || '').trim().toLowerCase().slice(0, 100);
    if (!name || !EMAIL_RE.test(email)) throw new Error('Level ' + (i + 1) + ' approver: enter a name and a valid e-mail');
    return { name: cell_(name, 60), email: email };
  }).map((a, i, arr) => { if (arr.some((b, j) => j !== i && b.name.toLowerCase() === a.name.toLowerCase())) throw new Error('Each level needs a different person. ' + a.name + ' is used more than once.'); return a; });
}
function apAny_(ap) { return Array.isArray(ap) && ap.some(a => a && (String(a.name || '').trim() || String(a.email || '').trim())); }
function userByEmail_(email) {    // active user whose Email column matches -> user name, else ''
  email = String(email || '').trim().toLowerCase(); if (!email) return '';
  const rows = sheet_('Users', USER_HEADERS).getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) if (String(rows[i][6] || '').trim().toLowerCase() === email && String(rows[i][4]).trim().toLowerCase() !== 'no') return String(rows[i][0]).trim().toLowerCase();
  return '';
}
function resolveUser_(input) {    // sign in with the user name OR the e-mail saved for that user
  input = String(input || '').trim().toLowerCase();
  if (!USER_RE.test(input) && !EMAIL_RE.test(input)) return '';
  if (USER_RE.test(input) && findUser_(input)) return input;
  return userByEmail_(input);
}
function esc_(t) { return String(t == null ? '' : t).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c])); }
function inr_(n) { return 'Rs. ' + Number(n || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 }); }
function findApproval_(id) {
  const sh = sheet_('Approvals', APPROVAL_HEADERS), n = sh.getLastRow(); if (n < 2) return null;
  const rows = sh.getRange(2, 1, n - 1, APPROVAL_HEADERS.length).getValues();
  for (let i = rows.length - 1; i >= 0; i--) if (String(rows[i][0]) === String(id)) return { sh: sh, row: i + 2, v: rows[i] };
  return null;
}
function logOf_(v) { try { return JSON.parse(v[17] || '[]'); } catch (e) { return []; } }
function approvalView_(a) {       // safe object for the browser
  const v = a.v;
  return { id: String(v[0]), status: String(v[1]), level: +v[2] || 0, by: String(v[3]), sent: v[16] instanceof Date ? v[16].toISOString() : String(v[16]),
    customer: String(v[5]), part: String(v[6]), perPart: v[7], round: v[8],
    approvers: [[9, 10], [11, 12], [13, 14]].map(c => ({ name: String(v[c[0]]), email: maskEmail_(String(v[c[1]])) })), log: logOf_(v) };
}
function mailMsg_(title, color, lines, link, linkText) {
  const body = lines.map(l => '<p style="margin:6px 0">' + l + '</p>').join('');
  return '<div style="font-family:Arial,sans-serif;max-width:560px;border:1px solid #ddd;border-radius:8px;overflow:hidden">' +
    '<div style="background:' + color + ';color:#fff;padding:12px 16px;font-size:16px;font-weight:bold">' + esc_(title) + '</div>' +
    '<div style="padding:14px 16px;font-size:14px;color:#222">' + body +
    (link ? '<p style="margin:16px 0"><a href="' + esc_(link) + '" style="background:#14262b;color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:bold">' + esc_(linkText) + '</a></p>' : '') +
    '</div><div style="background:#f5f5f5;color:#777;padding:8px 16px;font-size:12px">amace solutions pvt.ltd. - Metal printing cost estimator (automatic message)</div></div>';
}
function send_(to, subject, html) { MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, name: 'Amace Cost Estimator' }); }
function summaryLines_(v) {
  return ['<b>Customer:</b> ' + esc_(v[5] || '-') + ' &nbsp; <b>Part:</b> ' + esc_(v[6] || '-'),
    '<b>Per part cost:</b> ' + inr_(v[7]) + ' &nbsp; <b>Total round-off:</b> ' + inr_(v[8]), '<b>Prepared by:</b> ' + esc_(v[3])];
}
function siteLink_() { return SITE_URL ? SITE_URL.replace(/#.*$/, '') + '#approvals' : ''; }
function mailLevel_(a, lvl) {     // ask the approver of level lvl (1-3) to decide: they sign in to the website
  const v = a.v, name = v[7 + lvl * 2], email = v[8 + lvl * 2];
  send_(email, 'Approval needed (Level ' + lvl + '): quote ' + (v[6] || v[0]),
    mailMsg_('Quote approval - Level ' + lvl, '#7a1f3d', ['Dear ' + esc_(name) + ',', 'A cost quote is waiting for your approval.'].concat(summaryLines_(v), ['Sign in to the website with your user name (or e-mail) and password, open the <b>Approvals</b> tab and select this quote.']), siteLink_(), 'Open the website'));
}
function sendApproval_(d, sess) {
  if (!ID_RE.test(String(d.id))) return { ok: false, error: 'Invalid id' };
  const q = sheet_('Quotes', QUOTE_HEADERS), n = q.getLastRow(); if (n < 2) return { ok: false, error: 'Save the quote first' };
  const r = q.getRange(2, 1, n - 1, 18).getValues().filter(x => String(x[0]) === String(d.id))[0];
  if (!r) return { ok: false, error: 'Save the quote first' };
  if (String(r[17] || '') !== sess.u && !(sess.role === 'admin' && !r[17])) return { ok: false, error: 'This quote belongs to another user' };
  const me = findUser_(sess.u);
  if (me.ap.some(x => !x.name)) return { ok: false, error: 'No approvers are set for you. Ask the admin to choose 3 approvers in Users > Edit.' };
  const ap = me.ap.map(x => { const nm = String(x.name).trim().toLowerCase(), u = findUser_(nm);
    return { name: nm, email: u ? String(u.email || '').trim().toLowerCase() : '', ok: !!u && u.active && u.role === 'approver' && EMAIL_RE.test(String(u.email || '').trim()) }; });
  for (let i = 0; i < 3; i++) if (!ap[i].ok) return { ok: false, error: 'Level ' + (i + 1) + ' approver "' + ap[i].name + '" is not an active Approver login with an e-mail. Ask the admin to choose the approvers again in Users > Edit.' };
  if (new Set(ap.map(x => x.name)).size !== 3) return { ok: false, error: 'Each level needs a different person. Ask the admin to fix your approvers.' };
  if (ap.some(x => x.name === sess.u)) return { ok: false, error: 'You cannot be your own approver. Ask the admin to fix your approvers.' };
  const hash = hex_(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(r[16]))).slice(0, 40);
  const old = findApproval_(d.id);
  if (old && old.v[1] === 'REJECTED' && String(old.v[18] || '') === hash) return { ok: false, error: 'This quote was rejected. Update it as the remarks ask, then send for approval again.' };
  const stale = !!old && old.v[1] === 'PENDING' && ![9, 11, 13].every(c => { const u = findUser_(String(old.v[c]).trim().toLowerCase()); return !!u && u.role === 'approver'; });
  if (old && old.v[1] === 'PENDING' && !stale) return { ok: false, error: 'This quote is already waiting for approval (Level ' + old.v[2] + ')' };
  if (old && old.v[1] === 'APPROVED') return { ok: false, error: 'This quote is already fully approved' };
  const now = new Date();
  const row = [String(d.id), 'PENDING', 1, sess.u, me.email, r[2], r[3], r[15], r[14], ap[0].name, ap[0].email, ap[1].name, ap[1].email, ap[2].name, ap[2].email, '', now,
    JSON.stringify([{ t: now.toISOString(), e: 'SENT', who: sess.u, note: 'Sent for approval' }]), hash];
  const sh = sheet_('Approvals', APPROVAL_HEADERS);
  if (old) sh.getRange(old.row, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  SpreadsheetApp.flush();
  const a = findApproval_(d.id); mailLevel_(a, 1);
  if (me.email) send_(me.email, 'Quote sent for approval: ' + (r[3] || r[0]), mailMsg_('Sent for approval', '#14262b', ['Your quote was sent to <b>' + esc_(ap[0].name) + '</b> (Level 1).'].concat(summaryLines_(row))));
  audit_('APPROVAL_SENT', d.id); return { ok: true, approval: approvalView_(a) };
}
function myEmail_(sess) { return String(sess.u || '').trim().toLowerCase(); }   // the signed-in user name (approvers are matched by login name, not by e-mail)
function isApproverNow_(v, email) { const lvl = +v[2]; return v[1] === 'PENDING' && !!email && lvl >= 1 && lvl <= 3 && String(v[7 + lvl * 2]).trim().toLowerCase() === email; }
function myStatus_(v, email) {    // what this approver (by e-mail) has done / must do on this approval row
  if (!email) return { mine: '', myLvl: 0, when: '' };
  if (isApproverNow_(v, email)) return { mine: 'PENDING', myLvl: +v[2], when: v[16] instanceof Date ? v[16].toISOString() : String(v[16]) };
  const log = logOf_(v); let mine = '', myLvl = 0, when = '';
  [1, 2, 3].filter(L => String(v[7 + L * 2]).trim().toLowerCase() === email).forEach(L => {
    const e = log.filter(x => x.lvl === L && (x.e === 'APPROVED' || x.e === 'REJECTED')).pop(); if (e) { mine = e.e; myLvl = L; when = e.t; } });
  return { mine: mine, myLvl: myLvl, when: String(when) };
}
function myApprovals_(sess) {     // quotes waiting for me (first) and quotes I already decided
  const email = myEmail_(sess), sh = sheet_('Approvals', APPROVAL_HEADERS), n = sh.getLastRow(), out = [];
  if (n > 1 && email) sh.getRange(2, 1, n - 1, APPROVAL_HEADERS.length).getValues().forEach(v => {
    const m = myStatus_(v, email); if (!m.mine) return;
    const it = approvalView_({ v: v }); it.mine = m.mine; it.myLevel = m.myLvl; it.when = m.when; out.push(it);
  });
  out.sort((x, y) => (x.mine === 'PENDING') !== (y.mine === 'PENDING') ? (x.mine === 'PENDING' ? -1 : 1) : String(y.when).localeCompare(String(x.when)));
  return { ok: true, items: out, pending: out.filter(x => x.mine === 'PENDING').length };
}
function approvalDetail_(d, sess) {
  if (!ID_RE.test(String(d.id))) return { ok: false, error: 'Invalid id' };
  const a = findApproval_(d.id); if (!a) return { ok: false, error: 'Not found' };
  const email = myEmail_(sess), v = a.v, isApprover = [9, 11, 13].some(c => String(v[c]).trim().toLowerCase() === email && !!email);
  if (!(isApprover || String(v[3]) === sess.u || sess.role === 'admin')) return { ok: false, error: 'Not allowed' };
  const q = sheet_('Quotes', QUOTE_HEADERS), n = q.getLastRow(), r = n > 1 ? q.getRange(2, 1, n - 1, 18).getValues().filter(x => String(x[0]) === String(d.id))[0] : null;
  if (!r) return { ok: false, error: 'The quote was deleted' };
  let state; try { state = JSON.parse(r[16]); } catch (e) { return { ok: false, error: 'Quote data unreadable' }; }
  const ms = myStatus_(v, email);
  return { ok: true, approval: approvalView_(a), canDecide: isApproverNow_(v, email), mine: ms.mine, myLevel: ms.myLvl, state: state, saved: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), image: getImg_(String(d.id)) };
}
function decide_(d, sess) {       // approve / reject by a signed-in approver; remarks are compulsory for both
  const a = findApproval_(d.id); if (!a) return { ok: false, error: 'Not found' };
  const v = a.v, email = myEmail_(sess), decision = d.decision === 'approve' ? 'approve' : d.decision === 'reject' ? 'reject' : '';
  if (!decision) return { ok: false, error: 'Choose approve or reject' };
  if (v[1] !== 'PENDING') return { ok: false, error: 'This quote is no longer waiting for approval (status: ' + v[1] + ').' };
  if (!isApproverNow_(v, email)) return { ok: false, error: 'This quote is not waiting for your approval.' };
  const remarks = String(d.remarks || '').trim().slice(0, 500);
  if (remarks.length < 3) return { ok: false, error: 'Remarks are compulsory. Please type them.' };
  const lvl = +v[2], log = logOf_(v), who = v[7 + lvl * 2], now = new Date(), mail = v[4], sum = summaryLines_(v);
  log.push({ t: now.toISOString(), e: decision === 'approve' ? 'APPROVED' : 'REJECTED', lvl: lvl, who: who, note: cell_(remarks, 500) });
  if (decision === 'reject') {
    a.sh.getRange(a.row, 2, 1, 2).setValues([['REJECTED', lvl]]); a.sh.getRange(a.row, 18).setValue(JSON.stringify(log));
    if (mail) send_(mail, 'Quote REJECTED at Level ' + lvl + ': ' + (v[6] || v[0]), mailMsg_('Quote rejected - Level ' + lvl, '#b3261e', ['Rejected by <b>' + esc_(who) + '</b> (Level ' + lvl + ').', '<b>Remarks / changes:</b><br>' + esc_(remarks).replace(/\n/g, '<br>')].concat(sum, ['Please make the changes, save the quote and click <b>Send approval</b> again.']), siteLink_(), 'Open the website'));
    audit_('APPROVAL_REJECTED', d.id + ' L' + lvl, sess.u); return { ok: true, message: 'Rejected. Your remarks were e-mailed to ' + v[3] + '.' };
  }
  if (lvl < 3) {
    a.sh.getRange(a.row, 3).setValue(lvl + 1); a.sh.getRange(a.row, 18).setValue(JSON.stringify(log));
    mailLevel_(findApproval_(d.id), lvl + 1);
    if (mail) send_(mail, 'Quote approved at Level ' + lvl + ': ' + (v[6] || v[0]), mailMsg_('Level ' + lvl + ' approved', '#1b7f3b', ['Approved by <b>' + esc_(who) + '</b>.<br><b>Remarks:</b> ' + esc_(remarks), 'Now waiting for Level ' + (lvl + 1) + ' (' + esc_(v[9 + lvl * 2]) + ').'].concat(sum)));
    audit_('APPROVAL_L' + lvl, d.id, sess.u); return { ok: true, message: 'Approved. The quote went to Level ' + (lvl + 1) + ' (' + v[9 + lvl * 2] + ').' };
  }
  a.sh.getRange(a.row, 2, 1, 2).setValues([['APPROVED', 3]]); a.sh.getRange(a.row, 18).setValue(JSON.stringify(log));
  const all = [mail, v[10], v[12], v[14]].filter((x, i, arr) => x && arr.indexOf(x) === i).join(',');
  if (all) send_(all, 'Quote FULLY APPROVED: ' + (v[6] || v[0]), mailMsg_('Quote fully approved', '#1b7f3b', ['All 3 levels have approved this quote.', '<b>Level 3 remarks:</b> ' + esc_(remarks)].concat(sum)));
  audit_('APPROVAL_DONE', d.id, sess.u); return { ok: true, message: 'Approved. This was the final level - everyone has been notified.' };
}
function doGet() { return json_({ ok: true, message: 'Cost estimator API is running' }); }   // no data over GET

function doPost(e) {
  try {
    if (!e || !e.postData || e.postData.contents.length > MAX_BODY) return json_({ ok: false, error: 'Request too large' });
    let d; try { d = JSON.parse(e.postData.contents); } catch (x) { return json_({ ok: false, error: 'Bad request' }); }
    if (!d || typeof d !== 'object' || Array.isArray(d) || typeof d.action !== 'string' || d.action.length > 30) return json_({ ok: false, error: 'Bad request' });
    if (d.action === 'login') return json_(login_(d));
    if (d.action === 'forgotRequest') return json_(forgotRequest_(d));
    if (d.action === 'forgotReset') return json_(forgotReset_(d));
    const sess = session_(d.token); if (!sess) return json_(authFail_('SESSION', 'Please sign in again.'));
    CURRENT_USER = sess.u;
    if (!rate_(sess.u)) return json_({ ok: false, code: 'RATE', error: 'Too many requests. Wait a minute.' });
    if (d.action === 'logout') { CacheService.getScriptCache().remove('s:' + d.token); audit_('LOGOUT', ''); return json_({ ok: true }); }
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { return json_(handle_(d, sess)); } finally { lock.releaseLock(); }
  } catch (err) { Logger.log(err); return json_({ ok: false, error: 'Server error' }); }   // details only in the log, never sent to the browser
}

function handle_(d, sess) {
  const admin = sess && sess.role === 'admin';
  const mine = owner => owner === sess.u || (!owner && admin);   // a quote belongs to the user who saved it (old quotes without an owner belong to the admin)
  switch (d.action) {
    case 'list': {         // each user gets ONLY the quotes they saved
      const rows = sheet_('Quotes', QUOTE_HEADERS).getDataRange().getValues().slice(1), quotes = [], imgs = imgIds_(); let skipped = 0;
      rows.forEach(r => {
        if (!r[0] || !mine(String(r[17] || ''))) return;
        try { quotes.push({ id: String(r[0]), saved: r[1] instanceof Date ? r[1].toISOString() : String(r[1]), total: r[13], round: r[14], perPart: r[15], by: String(r[17] || ''), hasImg: !!imgs[String(r[0])], state: JSON.parse(r[16]) }); }
        catch (x) { skipped++; }
      });
      return { ok: true, quotes: quotes.reverse(), skipped: skipped };
    }
    case 'image': {        // picture of one of your own quotes
      if (!ID_RE.test(String(d.id))) return { ok: false, error: 'Invalid id' };
      const sh = sheet_('Quotes', QUOTE_HEADERS), n = sh.getLastRow(); if (n < 2) return { ok: true, image: '' };
      const r = sh.getRange(2, 1, n - 1, 18).getValues().filter(x => String(x[0]) === String(d.id))[0];
      if (!r || !mine(String(r[17] || ''))) return { ok: false, error: 'Not found' };
      return { ok: true, image: getImg_(String(d.id)) };
    }
    case 'masters': {
      const mats = seed_('Powders', POWDER_HEADERS, DEFAULT_POWDERS).getDataRange().getValues().slice(1).filter(r => r[0])
        .map(r => Object.fromEntries(POWDER_KEYS.map((k, i) => [k, i === 0 ? String(r[i]) : (num_(r[i]) || 0)])));
      const mach = seed_('Machines', MACHINE_HEADERS, DEFAULT_MACHINES).getDataRange().getValues().slice(1).filter(r => r[0])
        .map(r => Object.fromEntries(MACHINE_KEYS.map((k, i) => [k, i < 2 ? String(r[i]) : num_(r[i])])));
      return { ok: true, materials: mats, machines: mach, settings: settings_() };
    }
    case 'save': {
      if (!ID_RE.test(String(d.id)) || !d.state || typeof d.state !== 'object' || !d.state.i) return { ok: false, error: 'Invalid quote' };
      const stateJson = JSON.stringify(d.state); if (stateJson.length > MAX_STATE) return { ok: false, error: 'Quote too large' };
      if (d.image !== undefined && (typeof d.image !== 'string' || (d.image !== '' && (d.image.length > MAX_IMG || !IMG_RE.test(d.image))))) return { ok: false, error: 'Picture not accepted (use a smaller JPG/PNG)' };
      const sh = sheet_('Quotes', QUOTE_HEADERS), i = d.state.i;
      const row = [String(d.id), cell_(d.saved, 40), cell_(i.customer, 200), cell_(i.part, 200), cell_(i.machine, 100), cell_(i.material, 100), cell_(i.laser, 50),
        num_(i.parts), num_(i.layer), (num_(i.days) || 0) * 24 + (num_(i.hrs) || 0), num_(i.kg), num_(i.sets), num_(i.units), num_(d.total), num_(d.round), num_(d.perPart), stateJson].map(v => v === null ? '' : v);
      if (!sh.getRange(1, 18).getValue()) sh.getRange(1, 18).setValue('Saved by');      // older sheets get the new column header
      const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => r[0]), at = ids.indexOf(String(d.id));
      row.push(sess.u);
      if (at >= 0) {
        const owner = String(sh.getRange(at + 2, 18).getValue() || '');
        if (!mine(owner)) return { ok: false, error: 'This quote belongs to another user' };
        sh.getRange(at + 2, 1, 1, row.length).setValues([row]);
      } else sh.appendRow(row);
      if (typeof d.image === 'string') saveImg_(String(d.id), d.image);     // '' removes the picture; no 'image' key keeps the old one
      SpreadsheetApp.flush(); audit_('SAVE', d.id); return { ok: true, id: d.id };
    }
    case 'delete': {
      if (!ID_RE.test(String(d.id))) return { ok: false, error: 'Invalid id' };
      const sh = sheet_('Quotes', QUOTE_HEADERS), ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues().map(r => r[0]), at = ids.indexOf(String(d.id));
      if (at >= 0) {
        const owner = String(sh.getRange(at + 2, 18).getValue() || '');
        if (!mine(owner)) return { ok: false, error: 'This quote belongs to another user' };
        sh.deleteRow(at + 2); deleteImg_(String(d.id));
      }
      SpreadsheetApp.flush(); audit_('DELETE', d.id); return { ok: true };
    }
    case 'saveMasters': {
      if (!admin) return { ok: false, error: 'Only an admin can change the price lists' };
      const mats = Array.isArray(d.materials) ? d.materials : null, mach = Array.isArray(d.machines) ? d.machines : null;
      if (!mats || !mach || mats.length > MAX_POWDERS || mach.length > MAX_MACHINES) return { ok: false, error: 'Invalid lists' };
      const p = sheet_('Powders', POWDER_HEADERS); p.clearContents(); p.getRange(1, 1, 1, POWDER_HEADERS.length).setValues([POWDER_HEADERS]).setFontWeight('bold');
      if (mats.length) p.getRange(2, 1, mats.length, POWDER_KEYS.length).setValues(mats.map(x => POWDER_KEYS.map((k, n) => n === 0 ? cell_(x && x[k], 60) : (num_(x && x[k]) === null ? '' : num_(x[k])))));
      const m = sheet_('Machines', MACHINE_HEADERS); m.clearContents(); m.getRange(1, 1, 1, MACHINE_HEADERS.length).setValues([MACHINE_HEADERS]).setFontWeight('bold');
      if (mach.length) m.getRange(2, 1, mach.length, MACHINE_KEYS.length).setValues(mach.map(x => MACHINE_KEYS.map((k, n) => n < 2 ? cell_(x && x[k], 60) : (num_(x && x[k]) === null ? '' : num_(x[k])))));
      SpreadsheetApp.flush(); audit_('SAVE_LISTS', mats.length + ' powders, ' + mach.length + ' machines');
      return { ok: true, powders: mats.length, machines: mach.length };
    }
    case 'users': {
      if (!admin) return { ok: false, error: 'Admin only' };
      const rows = sheet_('Users', USER_HEADERS).getDataRange().getValues().slice(1).filter(r => r[0]);
      return { ok: true, users: rows.map(r => ({ username: String(r[0]), role: String(r[3] || 'user'), active: String(r[4]).trim().toLowerCase() !== 'no', email: String(r[6] || ''),
        ap: [[7, 8], [9, 10], [11, 12]].map(c => ({ name: String(r[c[0]] || ''), email: String(r[c[1]] || '') })) })) };
    }
    case 'userSave': {     // add a user / approver, or set a new password for an existing one (admin only)
      if (!admin) return { ok: false, error: 'Admin only' };
      const name = String(d.username || '').trim().toLowerCase(), role = d.role === 'admin' || d.role === 'approver' ? d.role : 'user';
      if (name === sess.u && role !== 'admin') return { ok: false, error: 'You cannot remove your own admin role' };
      let ap, email = String(d.email || '').trim().toLowerCase();
      try {
        if (role === 'approver' && !EMAIL_RE.test(email)) throw new Error('An approver needs a valid e-mail (it is used for the e-mail alerts and for choosing them as an approver)');
        if (email && !EMAIL_RE.test(email)) throw new Error('Enter a valid e-mail for the user');
        const ex = findUser_(name), hasAp = !!ex && ex.ap.every(a => a.name && EMAIL_RE.test(a.email));
        if (role === 'user' && (apAny_(d.approvers) || !hasAp)) ap = cleanAp_(d.approvers);      // a User needs all 3 approver levels (kept as they are when only the password is reset)
        else if (role === 'admin' && apAny_(d.approvers)) ap = cleanAp_(d.approvers);
        setUser_(name, d.password, role, sess.u, email, ap);
      } catch (x) { return { ok: false, error: x.message }; }
      return { ok: true };
    }
    case 'userActive': {
      if (!admin) return { ok: false, error: 'Admin only' };
      const name = String(d.username || '').trim().toLowerCase(), u = USER_RE.test(name) ? findUser_(name) : null;
      if (!u) return { ok: false, error: 'No such user' };
      if (name === sess.u) return { ok: false, error: 'You cannot disable your own account' };
      sheet_('Users', USER_HEADERS).getRange(u.row, 5).setValue(d.active ? 'Yes' : 'No'); audit_(d.active ? 'USER_ENABLED' : 'USER_DISABLED', name);
      return { ok: true };
    }
    case 'sendApproval': return sendApproval_(d, sess);
    case 'myApprovals': return myApprovals_(sess);
    case 'approvalDetail': return approvalDetail_(d, sess);
    case 'decide': return decide_(d, sess);
    case 'approvals': {      // status of the signed-in user's own quote approvals
      const sh = sheet_('Approvals', APPROVAL_HEADERS), n = sh.getLastRow(), out = {};
      if (n > 1) sh.getRange(2, 1, n - 1, APPROVAL_HEADERS.length).getValues().forEach(v => { if (String(v[3]) === sess.u || (admin && !v[3])) out[String(v[0])] = approvalView_({ v: v }); });
      return { ok: true, approvals: out };
    }
    case 'userApprovers': {  // admin: set the user's e-mail and (optionally) the 3 approver levels; no password change
      if (!admin) return { ok: false, error: 'Admin only' };
      const name = String(d.username || '').trim().toLowerCase(), u = USER_RE.test(name) ? findUser_(name) : null;
      if (!u) return { ok: false, error: 'No such user' };
      const email = String(d.email || '').trim().toLowerCase();
      if (!EMAIL_RE.test(email)) return { ok: false, error: 'Enter a valid e-mail for the user' };
      let ap = [{ name: '', email: '' }, { name: '', email: '' }, { name: '', email: '' }];
      try { if (u.role === 'user' || (u.role === 'admin' && apAny_(d.approvers))) ap = cleanAp_(d.approvers); } catch (x) { return { ok: false, error: x.message }; }
      const sh = sheet_('Users', USER_HEADERS); userCols_(sh);
      sh.getRange(u.row, 7, 1, 7).setValues([[email, ap[0].name, ap[0].email, ap[1].name, ap[1].email, ap[2].name, ap[2].email]]);
      audit_('USER_APPROVERS', name); return { ok: true };
    }
    default: return { ok: false, error: 'Unknown action' };
  }
}
