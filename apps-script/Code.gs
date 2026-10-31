/**
 * English ♥ Español — private chat room backend
 *
 * HOW TO INSTALL
 * 1. Open the Google Sheet > Extensions > Apps Script.
 * 2. Delete anything in Code.gs, paste this whole file, click Save.
 * 3. Pick "setup" in the function dropdown and click Run (approve the permissions).
 * 4. Back in the sheet, use the menu  ♥ Translator > Set / change room password.
 * 5. Deploy > New deployment > type: Web app
 *      Execute as: Me      Who has access: Anyone
 *    Click Deploy and copy the Web app URL (ends in /exec).
 *
 * The password is never stored as plain text. Only a scrambled (hashed) copy
 * is kept in Script Properties, which nobody can see in the sheet.
 */

const SHEET_ID = '1Y8CXanG_hf_gnW6MDn3KLyQkOlN-9R0hj3yivEiOUHM';
const MSG = 'Messages';
const SES = 'Sessions';
const SET = 'Settings';
const MAX_FAILS = 8;            // wrong passwords allowed...
const FAIL_WINDOW_SEC = 900;    // ...per 15 minutes, then locked out
const MAX_TEXT = 2000;

/* ======================= SHEET SETUP ======================= */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('♥ Translator')
    .addItem('1. Set up the sheet', 'setup')
    .addItem('2. Set / change room password', 'setPassword')
    .addSeparator()
    .addItem('Sign everyone out', 'signEveryoneOut')
    .addItem('Clear all messages', 'clearMessages')
    .addToUi();
}

function setup() {
  const ss = ss_();

  // ---- Settings tab ----
  const st = getOrCreate_(ss, SET, 0);
  const oldName = st.getLastRow() >= 2 ? String(st.getRange('B2').getValue() || '') : '';
  st.clear();
  st.getRange('A1:C1').setValues([['Setting', 'Value', 'Notes']]);
  st.getRange('A2:C7').setValues([
    ['Room name', oldName || 'Us ♥ Nosotros', 'You can change this. It shows at the top of the chat.'],
    ['Room password', hasPassword_() ? 'Set ✓ (hidden)' : 'NOT SET', 'Use the menu: ♥ Translator > Set / change room password'],
    ['Password last changed', PropertiesService.getScriptProperties().getProperty('pwChanged') || '', ''],
    ['Signed-in devices', '=MAX(0,COUNTA(Sessions!A2:A))', 'Phones/computers currently signed in'],
    ['Messages', '', 'Messages in the chat (not counting deleted ones)'],
    ['Web app URL', '', 'Paste your /exec link here for safekeeping (optional)']
  ]);
  st.getRange('B6').setFormula('=COUNTIFS(Messages!A2:A,"<>",Messages!G2:G,"<>yes")');
  styleHeader_(st, 3);
  st.getRange('A2:A7').setFontWeight('bold');
  st.getRange('C2:C7').setFontColor('#7a6f7f').setFontStyle('italic');
  st.setColumnWidth(1, 190); st.setColumnWidth(2, 260); st.setColumnWidth(3, 420);
  st.getRange('B2').setBackground('#fff7d6');

  // ---- Messages tab ----
  const mh = ['ID', 'Time', 'Name', 'Language', 'Original', 'Translation', 'Deleted'];
  const ms = getOrCreate_(ss, MSG, 1);
  if (ms.getLastRow() === 0) ms.appendRow(mh); else ms.getRange(1, 1, 1, mh.length).setValues([mh]);
  styleHeader_(ms, mh.length);
  ms.getRange('B:B').setNumberFormat('mmm d, yyyy h:mm am/pm');
  ms.getRange('C:F').setNumberFormat('@');
  ms.getRange('E:F').setWrap(true);
  [60, 160, 110, 80, 360, 360, 70].forEach((w, i) => ms.setColumnWidth(i + 1, w));

  // ---- Sessions tab ----
  const sh = ['Device key (scrambled)', 'Name', 'Language', 'Signed in', 'Last seen'];
  const ss2 = getOrCreate_(ss, SES, 2);
  if (ss2.getLastRow() === 0) ss2.appendRow(sh); else ss2.getRange(1, 1, 1, sh.length).setValues([sh]);
  styleHeader_(ss2, sh.length);
  ss2.getRange('D:E').setNumberFormat('mmm d, yyyy h:mm am/pm');
  [260, 120, 90, 170, 170].forEach((w, i) => ss2.setColumnWidth(i + 1, w));

  // Remove the empty default tab if it is still there
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);

  ss.setActiveSheet(st);
  toast_('Sheet is set up. Next: ♥ Translator > Set / change room password');
}

function setPassword() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Room password', 'Type the new room password (at least 6 characters).\nShare it only with your girlfriend.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const pw = r.getResponseText();
  if (!pw || pw.length < 6) { ui.alert('Password must be at least 6 characters. Nothing was changed.'); return; }
  const salt = Utilities.getUuid();
  const props = PropertiesService.getScriptProperties();
  const when = new Date().toLocaleString();
  props.setProperties({ pwSalt: salt, pwHash: sha_(salt + pw), pwChanged: when });
  signEveryoneOut_(true);
  const st = ss_().getSheetByName(SET);
  if (st) { st.getRange('B3').setValue('Set ✓ (hidden)'); st.getRange('B4').setValue(when); }
  ui.alert('Password saved. Everyone was signed out and must use the new password.');
}

function signEveryoneOut() {
  signEveryoneOut_(false);
  toast_('Everyone is signed out.');
}

function clearMessages() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Delete ALL messages?', 'This cannot be undone.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  const ms = ss_().getSheetByName(MSG);
  if (ms.getLastRow() > 1) ms.deleteRows(2, ms.getLastRow() - 1);
  const p = PropertiesService.getScriptProperties();
  p.setProperty('deleted', '[]');
  p.setProperty('epoch', String(Date.now()));
  bumpRev_();
  toast_('All messages deleted.');
}

/* ======================= WEB API ======================= */

function doGet() {
  return json_({ ok: true, app: 'engspan' });
}

function doPost(e) {
  let req = {};
  try { req = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (x) { return json_({ ok: false, error: 'bad_request' }); }
  try {
    switch (req.action) {
      case 'login':  return json_(login_(req));
      case 'poll':   return json_(poll_(req));
      case 'send':   return json_(send_(req));
      case 'delete': return json_(delete_(req));
      case 'logout': return json_(logout_(req));
      default:       return json_({ ok: false, error: 'unknown_action' });
    }
  } catch (err) {
    return json_({ ok: false, error: err.code || 'server_error', message: String(err.message || err) });
  }
}

function login_(req) {
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('fails') || 0);
  if (fails >= MAX_FAILS) throw fail_('too_many_tries');

  const props = PropertiesService.getScriptProperties();
  const salt = props.getProperty('pwSalt'), hash = props.getProperty('pwHash');
  if (!salt || !hash) throw fail_('password_not_set');

  if (sha_(salt + String(req.password || '')) !== hash) {
    cache.put('fails', String(fails + 1), FAIL_WINDOW_SEC);
    Utilities.sleep(800);
    throw fail_('wrong_password');
  }

  const name = clean_(req.name, 30);
  if (!name) throw fail_('name_required');
  const lang = req.lang === 'es' ? 'es' : 'en';

  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  const now = new Date();
  ss_().getSheetByName(SES).appendRow([sha_(token), safe_(name), lang, now, now]);
  return { ok: true, token, name, lang, room: roomName_() };
}

function poll_(req) {
  session_(req.token);
  const props = PropertiesService.getScriptProperties();
  const rev = props.getProperty('rev') || '0';
  const epoch = props.getProperty('epoch') || '0';
  const since = Number(req.since);
  const fresh = !(req.since >= 0) || String(req.epoch) !== epoch;

  if (!fresh && String(req.rev) === rev) return { ok: true, rev, epoch, same: true };

  const ms = ss_().getSheetByName(MSG);
  const last = ms.getLastRow();
  let rows = [];
  if (fresh) {
    const n = Math.min(300, Math.max(0, last - 1));
    if (n > 0) rows = ms.getRange(last - n + 1, 1, n, 7).getValues();
  } else if (last - 1 > since) {
    rows = ms.getRange(since + 2, 1, last - 1 - since, 7).getValues();
  }
  const msgs = rows.filter(r => r[0] !== '' && r[6] !== 'yes').map(rowToMsg_);
  return {
    ok: true, rev, epoch, fresh, room: roomName_(), msgs,
    lastId: Math.max(0, last - 1),
    deleted: JSON.parse(props.getProperty('deleted') || '[]')
  };
}

function send_(req) {
  const s = session_(req.token);
  const text = clean_(req.text, MAX_TEXT);
  if (!text) throw fail_('empty');
  const from = req.from === 'es' ? 'es' : 'en';
  const to = from === 'es' ? 'en' : 'es';
  let translation = '';
  try { translation = LanguageApp.translate(text, from, to); } catch (x) { translation = ''; }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ms = ss_().getSheetByName(MSG);
    const id = ms.getLastRow();               // row (id + 1) is the next empty row
    const now = new Date();
    ms.getRange(id + 1, 1, 1, 7).setValues([[id, now, safe_(s.name), from, safe_(text), safe_(translation), '']]);
    SpreadsheetApp.flush();
    bumpRev_();
    return { ok: true, msg: { id, t: now.getTime(), name: s.name, from, text, tr: translation } };
  } finally {
    lock.releaseLock();
  }
}

function delete_(req) {
  const s = session_(req.token);
  const id = Number(req.id);
  const ms = ss_().getSheetByName(MSG);
  if (!(id >= 1) || id + 1 > ms.getLastRow()) throw fail_('not_found');
  const row = ms.getRange(id + 1, 1, 1, 7).getValues()[0];
  if (String(row[2]) !== s.name) throw fail_('not_yours');
  ms.getRange(id + 1, 7).setValue('yes');
  const props = PropertiesService.getScriptProperties();
  const del = JSON.parse(props.getProperty('deleted') || '[]');
  del.push(id);
  props.setProperty('deleted', JSON.stringify(del.slice(-500)));
  bumpRev_();
  return { ok: true };
}

function logout_(req) {
  if (!req.token) return { ok: true };
  const h = sha_(req.token);
  const sh = ss_().getSheetByName(SES);
  const vals = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues();
  for (let i = vals.length - 1; i >= 1; i--) if (vals[i][0] === h) sh.deleteRow(i + 1);
  CacheService.getScriptCache().remove(sessKey_(h));
  return { ok: true };
}

/* ======================= HELPERS ======================= */

function session_(token) {
  if (!token) throw fail_('no_session');
  const h = sha_(String(token));
  const cache = CacheService.getScriptCache();
  const key = sessKey_(h);
  const hit = cache.get(key);
  if (hit) return JSON.parse(hit);

  const sh = ss_().getSheetByName(SES);
  const last = sh.getLastRow();
  if (last < 2) throw fail_('no_session');
  const vals = sh.getRange(2, 1, last - 1, 3).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (vals[i][0] === h) {
      const s = { name: String(vals[i][1]), lang: String(vals[i][2]) };
      sh.getRange(i + 2, 5).setValue(new Date());
      cache.put(key, JSON.stringify(s), 21600);
      return s;
    }
  }
  throw fail_('no_session');
}

function signEveryoneOut_(quiet) {
  const sh = ss_().getSheetByName(SES);
  if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
  const p = PropertiesService.getScriptProperties();
  p.setProperty('gen', String(Number(p.getProperty('gen') || 0) + 1));
}

function sessKey_(h) {
  return 'S' + (PropertiesService.getScriptProperties().getProperty('gen') || '0') + '_' + h.slice(0, 40);
}

function rowToMsg_(r) {
  return { id: Number(r[0]), t: r[1] instanceof Date ? r[1].getTime() : 0, name: String(r[2]), from: String(r[3]), text: String(r[4]), tr: String(r[5]) };
}

function bumpRev_() {
  const p = PropertiesService.getScriptProperties();
  p.setProperty('rev', String(Number(p.getProperty('rev') || 0) + 1));
}

function roomName_() {
  const st = ss_().getSheetByName(SET);
  return (st && String(st.getRange('B2').getValue())) || 'Our room';
}

function hasPassword_() {
  return !!PropertiesService.getScriptProperties().getProperty('pwHash');
}

function sha_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}

function clean_(s, max) {
  return String(s == null ? '' : s).replace(/\r/g, '').trim().slice(0, max);
}

// Stop text like "=SUM(...)" from being treated as a formula in the sheet
function safe_(s) {
  s = String(s);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function fail_(code) { const e = new Error(code); e.code = code; return e; }

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function getOrCreate_(ss, name, index) {
  return ss.getSheetByName(name) || ss.insertSheet(name, index);
}

function styleHeader_(sh, cols) {
  sh.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#2b2230').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function toast_(msg) {
  try { SpreadsheetApp.getActiveSpreadsheet().toast(msg, '♥ Translator', 6); } catch (e) {}
}
