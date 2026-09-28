/**
 * OMG Rental — Google Sheets API (Google Apps Script)
 * ---------------------------------------------------
 * Paste this into the Google Sheet: Extensions → Apps Script → Code.gs
 * 1. Run  setup()  once (creates every tab + the Owner login, PIN 1234).
 * 2. Deploy → New deployment → Web app · Execute as: Me · Who has access: Anyone
 * 3. Copy the Web app URL into the OMG Rental app (Settings → Connection) or config.js.
 *
 * The sheet is the single source of truth. The app never deletes rows:
 * bookings / payments / sales change status instead.
 */

var SCHEMA = {
  Collections:    { key: 'collection_id', cols: ['collection_id', 'name', 'status', 'description'] },
  Items:          { key: 'item_code', cols: ['item_code', 'type_code', 'collection_id', 'name', 'colour', 'colour_hex', 'size', 'design_style', 'purchase_cost', 'rent_price', 'sale_price', 'security_deposit', 'mode', 'condition', 'physical_status', 'notes', 'prev_rentals', 'prev_revenue', 'repair_cost', 'added_on'] },
  Packages:       { key: 'package_id', cols: ['package_id', 'name', 'price', 'description', 'slots', 'includes_salon_offer'] },
  Customers:      { key: 'customer_id', cols: ['customer_id', 'name', 'mobile', 'whatsapp', 'email', 'address', 'customer_since', 'omg_salon_client', 'notes', 'id_type', 'id_last4', 'email_opt', 'whatsapp_opt', 'sms_opt', 'marketing_opt'] },
  Bookings:       { key: 'booking_no', cols: ['booking_no', 'customer_id', 'customer_name', 'pickup_at', 'event_date', 'return_by', 'rent_gross', 'extra_days', 'discount', 'voucher_code', 'deposit_due', 'status', 'picked_at', 'received_at', 'closed_at', 'cancel_reason', 'late_flag_amount', 'late_flag_at', 'notes', 'created_by', 'created_at', 'rent_net', 'deposit_paid', 'total_paid', 'balance_due', 'payment_mode'] },
  Booking_Items:  { key: 'line_id', cols: ['line_id', 'booking_no', 'item_code', 'package_no', 'package_id', 'package_name', 'package_price', 'list_price', 'allocated_price', 'deposit', 'blocked_from', 'blocked_until'] },
  Payments:       { key: 'payment_id', cols: ['payment_id', 'booking_no', 'kind', 'amount', 'mode', 'at', 'recorded_by'] },
  Booking_Events: { key: 'event_id', cols: ['event_id', 'booking_no', 'at', 'by', 'text'] },
  Returns:        { key: 'booking_no', cols: ['booking_no', 'received_at', 'inspected_at', 'late_minutes', 'late_fee', 'late_fee_waived', 'refund', 'refund_mode', 'extra_collected', 'collect_mode', 'rent_deducted'] },
  Return_Items:   { key: 'line_id', cols: ['line_id', 'booking_no', 'item_code', 'condition', 'charge', 'note'] },
  Sales:          { key: 'sale_id', cols: ['sale_id', 'item_code', 'customer_id', 'walk_in', 'list_price', 'discount', 'voucher_code', 'price', 'mode', 'at', 'sold_by', 'voided'] },
  Vouchers:       { key: 'code', cols: ['code', 'description', 'issuer', 'type', 'value', 'max_discount', 'min_amount', 'valid_from', 'valid_to', 'usage_limit', 'used', 'per_customer', 'collection', 'applies_to', 'status'] },
  Users:          { key: 'user_id', cols: ['user_id', 'name', 'role', 'login', 'active', 'pin_hash'] },
  Settings:       { key: 'key', cols: ['key', 'value'] },
  Item_History:   { key: 'history_id', cols: ['history_id', 'item_code', 'at', 'by', 'event'] },
  Audit_Log:      { key: 'audit_id', cols: ['audit_id', 'at', 'user', 'role', 'action', 'record', 'details'] },
  Item_Photos:    { key: 'photo_id', cols: ['photo_id', 'item_code', 'customer_id', 'booking_no', 'kind', 'url', 'path', 'size_kb', 'uploaded_at', 'uploaded_by', 'active'] },
  Email_Log:      { key: 'email_id', cols: ['email_id', 'dedup_key', 'template', 'channel', 'booking_no', 'sale_id', 'customer_id', 'to', 'subject', 'status', 'created_at', 'sent_at', 'attempts', 'error', 'meta'] }
};
var ADMIN_TABLES = ['Users', 'Settings', 'Collections', 'Packages'];
var DATE_COL = /(_at|_date|_from|_until|^at|^valid_from|^valid_to)$/;
var SESSION_SECONDS = 6 * 60 * 60;

/* ============ ONE-TIME SETUP (run from the Apps Script editor) ============ */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone('Asia/Kolkata');
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SALT')) props.setProperty('SALT', Utilities.getUuid());
  Object.keys(SCHEMA).forEach(function (name, i) {
    var sh = ss.getSheetByName(name);
    if (!sh) {
      var first = ss.getSheets()[0];
      if (i === 0 && first.getLastRow() === 0 && first.getName().indexOf('Sheet') === 0) { sh = first; sh.setName(name); }
      else sh = ss.insertSheet(name);
    }
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, SCHEMA[name].cols.length).setValues([SCHEMA[name].cols])
        .setFontWeight('bold').setFontColor('#ffffff').setBackground('#7a1f4b');
      sh.setFrozenRows(1);
    }
  });
  var users = sheetObjects('Users');
  if (!users.length) {
    appendObjects('Users', [{ user_id: 'u1', name: 'Shop Owner', role: 'admin', login: 'owner', active: 'yes', pin_hash: hashPin('1234') }]);
  }
  Logger.log('Setup complete. Owner login PIN is 1234 — change it in the app (Settings → Users & roles).');
}

/* First-run safety net: if the sheet has no tabs/users yet, set it up automatically. */
function ensureSetup() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (sh && sh.getLastRow() >= 2) return;
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try { sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users'); if (!sh || sh.getLastRow() < 2) setup(); }
  finally { lock.releaseLock(); }
}

/* ============ HTTP ENTRY POINTS ============ */
function doGet(e) {
  if (e && e.parameter && e.parameter.unsub) return unsubscribePage(e.parameter.unsub, e.parameter.t);   // link in email footers
  return json({ ok: true, app: 'OMG Rental API', version: 1 });
}

function doPost(e) {
  var req;
  try { req = JSON.parse(e.postData.contents); } catch (x) { return json({ ok: false, error: 'Bad request' }); }
  try {
    switch (req.action) {
      case 'ping': return json({ ok: true, serverTime: new Date().toISOString() });
      case 'users':
        ensureSetup();
        return json({ ok: true, users: sheetObjects('Users').filter(isActive).map(function (u) { return { user_id: u.user_id, name: u.name, role: u.role }; }) });
      case 'login':
        var lg = login(req.user_id, req.pin);
        if (lg.ok && req.withData) { lg.tables = loadAll(); lg.photosEnabled = photosEnabled(); lg.serverTime = new Date().toISOString(); }   // saves a second round trip
        return json(lg);
      case 'catalog': return json(catalog());   // public, no login: QR tags open this
      case 'logout': CacheService.getScriptCache().remove('t_' + req.token); return json({ ok: true });
      case 'load': var u = auth(req.token); return json({ ok: true, user: u, tables: loadAll(!!req.fresh), photosEnabled: photosEnabled(), serverTime: new Date().toISOString() });
      case 'commit': return json(commit(auth(req.token), req.changes || {}, req.newBookings || []));
      case 'verifyPin': auth(req.token); return json(verifyPin(req.pin));
      case 'setPin': return json(setPin(auth(req.token), req.user_id, req.pin));
      case 'resetData': return json(resetData(auth(req.token), req.pin));
      case 'uploadPhoto': return json(uploadPhoto(auth(req.token), req));
      case 'emailStatus': return json(emailStatus(adminOnly(auth(req.token))));
      case 'emailPreview': adminOnly(auth(req.token)); return json(emailPreview(req));
      case 'emailTest': adminOnly(auth(req.token)); return json(emailTest(req));
      case 'emailRun': adminOnly(auth(req.token)); emailTick(); return json({ ok: true });
      case 'emailSend': return json(emailSendNow(auth(req.token), req));
      default: return json({ ok: false, error: 'Unknown action' });
    }
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err), code: (err && err.code) || 'error' });
  }
}

/* ============ AUTH ============ */
function hashPin(pin) {
  var salt = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + String(pin), Utilities.Charset.UTF_8);
  return Utilities.base64Encode(raw);
}
function isActive(u) { return String(u.active).toLowerCase() !== 'no'; }
function fail(msg, code) { var e = new Error(msg); e.code = code || 'error'; throw e; }
function login(userId, pin) {
  var cache = CacheService.getScriptCache(), k = 'fail_' + userId, n = +(cache.get(k) || 0);
  if (n >= 5) return { ok: false, error: 'Too many wrong PINs. Try again in 10 minutes.' };
  var u = sheetObjects('Users').filter(function (x) { return x.user_id === userId && isActive(x); })[0];
  if (!u || u.pin_hash !== hashPin(pin)) { cache.put(k, String(n + 1), 600); return { ok: false, error: 'Wrong PIN' }; }
  cache.remove(k);
  var token = Utilities.getUuid(), user = { user_id: u.user_id, name: u.name, role: u.role };
  cache.put('t_' + token, JSON.stringify(user), SESSION_SECONDS);
  return { ok: true, token: token, user: user };
}
function auth(token) {
  var cache = CacheService.getScriptCache(), s = token && cache.get('t_' + token);
  if (!s) fail('Session expired — please log in again', 'auth');
  cache.put('t_' + token, s, SESSION_SECONDS);
  return JSON.parse(s);
}
function verifyPin(pin) {
  var h = hashPin(pin);
  var m = sheetObjects('Users').filter(function (u) { return isActive(u) && (u.role === 'manager' || u.role === 'admin') && u.pin_hash === h; })[0];
  return m ? { ok: true, name: m.name } : { ok: false, error: 'Incorrect manager PIN' };
}
function setPin(user, userId, pin) {
  if (user.role !== 'admin' && user.user_id !== userId) fail('Only an admin can change another user’s PIN', 'forbidden');
  if (!/^\d{4,6}$/.test(String(pin))) fail('PIN must be 4–6 digits');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try { upsertObjects('Users', [{ user_id: userId, pin_hash: hashPin(pin) }]); } finally { lock.releaseLock(); }
  return { ok: true };
}

/* Admin tool: empty every data tab (keeps Users + Settings). Needs the admin's own PIN.
   Google Sheets version history (File → Version history) can still restore old data. */
function resetData(user, pin) {
  if (user.role !== 'admin') fail('Only an admin can clear data', 'forbidden');
  var me = sheetObjects('Users').filter(function (u) { return u.user_id === user.user_id; })[0];
  if (!me || me.pin_hash !== hashPin(pin)) fail('Wrong PIN');
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(SCHEMA).forEach(function (t) {
      if (t === 'Users' || t === 'Settings') return;
      var sh = ss.getSheetByName(t);
      if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1);
    });
    clearCachedTables();
    upsertObjects('Audit_Log', [{ audit_id: 'A-' + Utilities.getUuid(), at: new Date().toISOString(), user: user.name, role: 'Admin', action: 'All data cleared', record: 'Google Sheet', details: 'Users and settings kept' }]);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

/* ============ PHOTOS → GitHub ============
   Photos go to the "photos" branch of the (public) app repository, so they never
   trigger a Netlify rebuild of the main branch.
   Apps Script → Project Settings → Script Properties:
     GITHUB_TOKEN  = fine-grained token with "Contents: Read and write" on this repo only  (required)
     GITHUB_REPO   = owner/repo      (optional, default below)
     GITHUB_BRANCH = branch name     (optional, default "photos")
   The token never leaves Apps Script. The app sends an already-compressed JPEG. */
var DEFAULT_PHOTO_REPO = 'sanskar-malviya/omg-rental';
var DEFAULT_PHOTO_BRANCH = 'photos';
/* Run this once from the editor: it asks Google for the "Connect to an external service"
   permission and checks that the GitHub token works. */
function connectGithub() {
  var P = PropertiesService.getScriptProperties();
  var token = P.getProperty('GITHUB_TOKEN'), repo = P.getProperty('GITHUB_REPO') || DEFAULT_PHOTO_REPO, branch = P.getProperty('GITHUB_BRANCH') || DEFAULT_PHOTO_BRANCH;
  if (!token) { Logger.log('❌ GITHUB_TOKEN is missing — add it in Project Settings → Script Properties.'); return; }
  var res = UrlFetchApp.fetch('https://api.github.com/repos/' + repo + '/branches/' + branch, { headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json' }, muteHttpExceptions: true });
  var code = res.getResponseCode();
  Logger.log(code === 200 ? '✅ GitHub connected: ' + repo + ' (branch ' + branch + '). Photo uploads are ready.'
                          : '❌ GitHub said ' + code + ': ' + res.getContentText().slice(0, 200));
}
function photosEnabled() {
  return !!PropertiesService.getScriptProperties().getProperty('GITHUB_TOKEN');
}
function uploadPhoto(user, req) {
  var P = PropertiesService.getScriptProperties();
  var token = P.getProperty('GITHUB_TOKEN'), repo = P.getProperty('GITHUB_REPO') || DEFAULT_PHOTO_REPO, branch = P.getProperty('GITHUB_BRANCH') || DEFAULT_PHOTO_BRANCH;
  if (!token) fail('Photo storage is not set up yet — add GITHUB_TOKEN in Apps Script → Project Settings → Script Properties', 'photos_off');
  var data = String(req.data || '');
  if (!data || !/^[A-Za-z0-9+\/=]+$/.test(data)) fail('Invalid image data');
  if (data.length > 2800000) fail('Image too large (max ~2 MB after compression)');
  var kind = req.kind === 'damage' ? 'damage' : req.kind === 'id' ? 'id' : 'item';
  var code = kind === 'id' ? '' : String(req.item_code || '').replace(/[^A-Za-z0-9-]/g, '');
  var custId = kind === 'id' ? String(req.customer_id || '').replace(/[^A-Za-z0-9-]/g, '') : '';
  if (kind === 'id' && !custId) fail('Customer is required');
  if (kind !== 'id' && !code) fail('Item code is required');
  var bno = kind === 'id' ? '' : String(req.booking_no || '').replace(/[^A-Za-z0-9-]/g, '');
  var stamp = Utilities.formatDate(new Date(), 'Asia/Kolkata', 'yyyyMMdd-HHmmss');
  var rand = Utilities.getUuid().slice(0, 4);
  // ID proofs get a long random file name so the link cannot be guessed
  var path = kind === 'id' ? 'ids/' + custId + '/' + Utilities.getUuid() + '.jpg'
           : kind === 'damage' ? 'damage/' + (bno || 'no-booking') + '/' + code + '-' + stamp + '-' + rand + '.jpg'
                               : 'items/' + code + '/' + code + '-' + stamp + '-' + rand + '.jpg';
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var res = UrlFetchApp.fetch('https://api.github.com/repos/' + repo + '/contents/' + path, {
      method: 'put', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' },
      payload: JSON.stringify({ message: (kind === 'id' ? 'ID proof ' + custId : 'Photo ' + path) + ' (' + user.name + ')', content: data, branch: branch })
    });
    var status = res.getResponseCode();
    if (status !== 201 && status !== 200) {
      var msg = ''; try { msg = JSON.parse(res.getContentText()).message; } catch (x) { }
      fail('GitHub upload failed (' + status + (msg ? ': ' + msg : '') + ')' + (status === 401 || status === 403 || status === 404 ? ' — check GITHUB_TOKEN permissions and GITHUB_REPO' : ''));
    }
    var photo = {
      photo_id: 'PH-' + Utilities.getUuid().slice(0, 8), item_code: code, customer_id: custId, booking_no: bno, kind: kind,
      url: 'https://raw.githubusercontent.com/' + repo + '/' + branch + '/' + path, path: path,
      size_kb: Math.round(data.length * 0.75 / 1024), uploaded_at: new Date().toISOString(), uploaded_by: user.name, active: 'yes'
    };
    upsertObjects('Item_Photos', [photo]);
    return { ok: true, photo: photo };
  } finally { lock.releaseLock(); }
}

/* ============ READ ============
   Reading 17 tabs from the sheet takes several seconds, so a ready copy of all tables
   is kept in the script cache. Every save through the app (upsertObjects) and every hand
   edit in the sheet (onEdit) clears it; the next load rebuilds it. */
var CACHE_META = 'db_meta', CACHE_CHUNK = 25000, CACHE_SECONDS = 6 * 60 * 60;
function loadAll(fresh) {
  var cache = CacheService.getScriptCache();
  if (!fresh) { var hit = readCachedTables(cache); if (hit) return hit; }
  var lock = LockService.getScriptLock(); lock.waitLock(25000);   // no save can run while the copy is built
  try {
    if (!fresh) { var again = readCachedTables(cache); if (again) return again; }
    var out = readAllSheets();
    writeCachedTables(cache, out);
    return out;
  } finally { lock.releaseLock(); }
}
function readAllSheets() {
  var byName = {};
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sh) { byName[sh.getName()] = sh; });
  var out = {};
  Object.keys(SCHEMA).forEach(function (t) {
    out[t] = byName[t] ? rowsToObjects(byName[t].getDataRange().getValues()).map(function (o) { delete o.pin_hash; return o; }) : [];
  });
  return out;
}
function readCachedTables(cache) {
  try {
    var meta = cache.get(CACHE_META); if (!meta) return null;
    meta = JSON.parse(meta);
    var keys = []; for (var i = 0; i < meta.n; i++) keys.push('db_' + meta.id + '_' + i);
    var got = cache.getAll(keys), s = '';
    for (var k = 0; k < keys.length; k++) { if (got[keys[k]] == null) return null; s += got[keys[k]]; }
    return JSON.parse(s);
  } catch (x) { return null; }
}
function writeCachedTables(cache, tables) {
  try {
    var s = JSON.stringify(tables), id = Utilities.getUuid().slice(0, 8), parts = {}, n = 0;
    for (var i = 0; i < s.length; i += CACHE_CHUNK) parts['db_' + id + '_' + (n++)] = s.slice(i, i + CACHE_CHUNK);
    cache.putAll(parts, CACHE_SECONDS);
    cache.put(CACHE_META, JSON.stringify({ id: id, n: n }), CACHE_SECONDS);
  } catch (x) { }   // too big for the cache: loads just read the sheet
}
function clearCachedTables() { try { CacheService.getScriptCache().remove(CACHE_META); } catch (x) { } }
/* Runs automatically when someone edits the sheet by hand, so the app sees the change. */
function onEdit() { clearCachedTables(); }
function sheetObjects(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  return sh ? rowsToObjects(sh.getDataRange().getValues()) : [];
}
function rowsToObjects(v) {
  if (v.length < 2) return [];
  var h = v[0], out = [];
  for (var r = 1; r < v.length; r++) {
    if (v[r].join('') === '') continue;
    var o = {};
    for (var c = 0; c < h.length; c++) if (h[c]) o[h[c]] = v[r][c];
    out.push(o);
  }
  return out;
}

/* ============ PUBLIC CATALOGUE (no login) ============
   Only what a customer may see: item looks, item photos and booked date ranges.
   Never prices, costs, notes, customers, bookings, ID proofs or settings like PINs. */
var PUBLIC_SETTINGS = ['shopName', 'city', 'pickupStart', 'pickupEnd', 'returnDeadline', 'bufferHours', 'festivalName', 'festivalStart', 'festivalNights'];
function catalog() {
  var t = loadAll(), now = new Date(), status = {}, colName = {}, codes = {};
  t.Bookings.forEach(function (b) { status[b.booking_no] = b.status; });
  t.Collections.forEach(function (c) { colName[c.collection_id] = c.name; });
  var items = t.Items.filter(function (i) { return i.item_code && ['sold', 'lost'].indexOf(String(i.physical_status)) < 0; }).map(function (i) {
    codes[i.item_code] = 1;
    return { code: String(i.item_code), type: String(i.type_code), name: String(i.name), colour: String(i.colour), hex: String(i.colour_hex), size: String(i.size), style: String(i.design_style), mode: String(i.mode || 'rent'), status: String(i.physical_status || 'in'), collection: colName[i.collection_id] || '' };
  });
  var photos = {};
  t.Item_Photos.forEach(function (p) {
    if (p.kind === 'item' && String(p.active).toLowerCase() !== 'no' && codes[p.item_code] && p.url) (photos[p.item_code] = photos[p.item_code] || []).push(String(p.url));
  });
  var blocks = {}, yesterday = now.getTime() - 864e5;
  t.Booking_Items.forEach(function (l) {
    var st = status[l.booking_no];
    if ((st !== 'reserved' && st !== 'picked') || !codes[l.item_code]) return;
    var f = new Date(l.blocked_from), u = new Date(l.blocked_until);
    if (st === 'picked' && now > u) u = now;                      // still out (overdue)
    if (u.getTime() < yesterday) return;
    (blocks[l.item_code] = blocks[l.item_code] || []).push([f.toISOString(), u.toISOString()]);
  });
  var settings = {};
  t.Settings.forEach(function (r) {
    if (PUBLIC_SETTINGS.indexOf(r.key) < 0) return;
    try { settings[r.key] = JSON.parse(String(r.value)); } catch (x) { settings[r.key] = r.value; }
  });
  return { ok: true, shop: { name: settings.shopName || 'OMG Rental', city: settings.city || '' }, settings: settings, items: items, photos: photos, blocks: blocks, serverTime: now.toISOString() };
}

/* ============ WRITE (inside a lock, so two tablets can't clash) ============ */
function commit(user, changes, newBookings) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    Object.keys(changes).forEach(function (t) {
      if (!SCHEMA[t]) fail('Unknown table ' + t);
      if (ADMIN_TABLES.indexOf(t) >= 0 && user.role !== 'admin') fail('Only an admin can change ' + t, 'forbidden');
    });
    if (changes.Email_Log) fail('The email log is written by the server only', 'forbidden');
    if (changes.Users) changes.Users.forEach(function (r) { delete r.pin_hash; });
    checkNewBookings(changes, newBookings);
    Object.keys(changes).forEach(function (t) { if (changes[t].length) upsertObjects(t, changes[t]); });
    var queued = 0;
    try { queued = queueEventEmails(changes); } catch (x) { }   // email problems never block a save
  } finally { lock.releaseLock(); }
  if (queued) scheduleEmailRun();
  return { ok: true, serverTime: new Date().toISOString() };
}

/* Double-booking guard: re-checked on the server for every new booking. */
function checkNewBookings(changes, newBookings) {
  if (!newBookings.length) return;
  var bookings = sheetObjects('Bookings'), status = {};
  bookings.forEach(function (b) { status[b.booking_no] = b.status; });
  newBookings.forEach(function (no) { if (status[no]) fail('Booking number ' + no + ' is already taken — refreshing, please confirm again', 'conflict'); });
  var existing = sheetObjects('Booking_Items').filter(function (l) { return status[l.booking_no] === 'reserved' || status[l.booking_no] === 'picked'; });
  var now = new Date();
  (changes.Booking_Items || []).forEach(function (l) {
    if (newBookings.indexOf(l.booking_no) < 0) return;
    var from = new Date(l.blocked_from), until = new Date(l.blocked_until);
    existing.forEach(function (x) {
      if (x.item_code !== l.item_code) return;
      var xf = new Date(x.blocked_from), xu = new Date(x.blocked_until);
      if (status[x.booking_no] === 'picked' && now > xu) xu = now;       // still out (overdue)
      if (from < xu && xf < until) fail(l.item_code + ' was just booked on ' + x.booking_no + ' — please choose another piece', 'conflict');
    });
  });
}

function upsertObjects(name, rows) {
  clearCachedTables();
  var ss =SpreadsheetApp.getActiveSpreadsheet(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, SCHEMA[name].cols.length).setValues([SCHEMA[name].cols]); sh.setFrozenRows(1); }
  var key = SCHEMA[name].key;
  var lastCol = Math.max(1, sh.getLastColumn());
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  rows.forEach(function (r) { Object.keys(r).forEach(function (k) { if (headers.indexOf(k) < 0) { headers.push(k); sh.getRange(1, headers.length).setValue(k).setFontWeight('bold'); } }); });
  var kc = headers.indexOf(key) + 1;
  var last = sh.getLastRow(), idx = {};
  if (last >= 2) sh.getRange(2, kc, last - 1, 1).getValues().forEach(function (v, i) { idx[String(v[0])] = i + 2; });
  var appends = [];
  rows.forEach(function (r) {
    var rowNo = idx[String(r[key])];
    if (rowNo) {
      var cur = sh.getRange(rowNo, 1, 1, headers.length).getValues()[0];
      headers.forEach(function (h, c) { if (h in r) cur[c] = cell(h, r[h]); });
      sh.getRange(rowNo, 1, 1, headers.length).setValues([cur]);
    } else {
      appends.push(headers.map(function (h) { return h in r ? cell(h, r[h]) : ''; }));
      idx[String(r[key])] = -1;
    }
  });
  if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, headers.length).setValues(appends);
}
function appendObjects(name, rows) { upsertObjects(name, rows); }
function cell(h, v) {
  if (v === null || v === undefined) return '';
  if (DATE_COL.test(h) && typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T/.test(v)) return new Date(v);
  return v;
}
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* =====================================================================
   EMAIL, RECEIPTS & REMINDERS
   ---------------------------------------------------------------------
   • Sent through Gmail of the account that owns this script. To send FROM your business
     Gmail, add it once in that Gmail: Settings → Accounts → "Send mail as" (SMTP
     smtp.gmail.com · port 587 · TLS · your business email + its App Password), then put the
     address in the app: Settings → Email & reminders → Sender email.
   • One-time: run setupEmail() in this editor → Allow → Deploy → Manage deployments → New version.
   • Every email is written to the Email_Log tab BEFORE sending. Its dedup_key stops duplicates.
   • Booking / return / sale emails are queued when the app saves and sent within ~1 minute.
     Timed reminders run every 15 minutes (emailTick), even when no tablet is open.
   ===================================================================== */
var TZ = 'Asia/Kolkata';
var EMAIL_DEFAULTS = {
  enabled: true, senderName: 'OMG Rental', senderEmail: '', replyTo: '', staffEmail: '',
  businessPhone: '7610544284', businessAddress: 'Indore, Madhya Pradesh', mapsLink: '', reviewLink: '', feedbackEmail: '',
  logoUrl: 'https://omgrental.netlify.app/omg-logo.jpg', appUrl: 'https://omgrental.netlify.app/',
  pickupTomorrowTime: '18:00', pickupTodayTime: '09:00', paymentDueTime: '11:00',
  returnTomorrowTime: '18:00', returnTodayTime: '08:00',
  overdue1Time: '12:15', overdue2Time: '15:00', overdue3Time: '19:00',
  reviewDelayHours: 24, staffDigestTime: '08:30'
};
/* kind: event = sent when the app saves · auto = timed rule · staff = to the shop · future = not active yet */
var TEMPLATES = {
  booking_confirmation: { name: 'Booking confirmation + receipt', trigger: 'A new booking is saved', kind: 'event', pdf: true,
    subject: 'Thank you for choosing OMG Rental — Booking {{booking_number}}', heading: 'Your booking is confirmed',
    intro: 'Hi {{customer_name}}, thank you for choosing OMG Rental! Your outfit is reserved just for you. Your receipt is attached to this email.',
    blocks: ['booking', 'dates', 'items', 'money', 'instructions'], cta: ['call', 'whatsapp', 'directions'] },
  booking_updated: { name: 'Booking updated', trigger: 'Dates or amount of a booking change', kind: 'event',
    subject: 'Your OMG Rental booking {{booking_number}} has been updated', heading: 'Your booking has been updated',
    intro: 'Hi {{customer_name}}, the details of your booking have changed. Please check the new details below.',
    blocks: ['changes', 'dates', 'items', 'money'], cta: ['call', 'whatsapp'] },
  pickup_tomorrow: { name: 'Pickup reminder — day before', trigger: 'Day before pickup, at the set time', kind: 'auto', timeKey: 'pickupTomorrowTime',
    subject: 'Your OMG Rental pickup is tomorrow — {{booking_number}}', heading: 'Your pickup is tomorrow',
    intro: 'Hi {{customer_name}}, your outfit will be ready for collection tomorrow. Pickup window: {{pickup_time}}.',
    blocks: ['booking', 'dates', 'items', 'balance'], cta: ['call', 'whatsapp', 'directions'] },
  pickup_today: { name: 'Pickup day — outfit ready', trigger: 'Morning of pickup, at the set time', kind: 'auto', timeKey: 'pickupTodayTime',
    subject: 'Your OMG Rental outfit is ready — pickup today', heading: 'Your outfit is ready',
    intro: 'Hi {{customer_name}}, your outfit is ready! Your pickup window is {{pickup_time}} today. Please bring a photo ID, and the balance or deposit if it is still pending.',
    blocks: ['booking', 'items', 'balance'], cta: ['call', 'whatsapp', 'directions'] },
  payment_due: { name: 'Balance due reminder', trigger: 'Day before pickup, if a balance is unpaid', kind: 'auto', timeKey: 'paymentDueTime',
    subject: 'Payment reminder — Booking {{booking_number}}', heading: 'A balance is pending',
    intro: 'Hi {{customer_name}}, a balance of {{balance_due}} is pending for your booking. You can pay at pickup by cash or UPI.',
    blocks: ['booking', 'money'], cta: ['call', 'whatsapp'] },
  return_tomorrow: { name: 'Return reminder — day before', trigger: 'Evening before the return day', kind: 'auto', timeKey: 'returnTomorrowTime',
    subject: 'Your OMG Rental return is tomorrow — {{booking_number}}', heading: 'Your return is tomorrow',
    intro: 'Hi {{customer_name}}, we hope you had a wonderful time! Please return your outfit tomorrow, {{return_date}}, before {{return_time}}.',
    blocks: ['booking', 'dates', 'items', 'returnrules'], cta: ['call', 'whatsapp', 'directions'] },
  return_today: { name: 'Return due today', trigger: 'Morning of the return day', kind: 'auto', timeKey: 'returnTodayTime',
    subject: 'Your OMG Rental return is due today before {{return_time}}', heading: 'Your return is due today',
    intro: 'Hi {{customer_name}}, your rented items are due back today. Return deadline: {{return_time}}. Please return them before the deadline to avoid late charges.',
    blocks: ['booking', 'items', 'returnrules'], cta: ['call', 'whatsapp', 'directions'] },
  overdue_1: { name: 'Overdue — 1st reminder', trigger: 'Return day, after the deadline + grace', kind: 'auto', timeKey: 'overdue1Time',
    subject: 'Your OMG Rental return is overdue — {{booking_number}}', heading: 'Your return is overdue',
    intro: 'Hi {{customer_name}}, your rental was due back before {{return_time}} on {{return_date}}. Please return the items or call us as soon as possible. Late charges apply as per the rental terms.',
    blocks: ['booking', 'items', 'returnrules'], cta: ['call', 'whatsapp'] },
  overdue_2: { name: 'Overdue — 2nd reminder', trigger: 'Return day, second reminder time', kind: 'auto', timeKey: 'overdue2Time',
    subject: 'Reminder: please return your OMG Rental items — {{booking_number}}', heading: 'Please return your items',
    intro: 'Hi {{customer_name}}, we have not yet received the items from booking {{booking_number}}, which were due before {{return_time}} on {{return_date}}. Please return them today or call us.',
    blocks: ['booking', 'items', 'returnrules'], cta: ['call', 'whatsapp'] },
  overdue_3: { name: 'Overdue — final reminder', trigger: 'Return day, final reminder time', kind: 'auto', timeKey: 'overdue3Time',
    subject: 'Final reminder: OMG Rental return overdue — {{booking_number}}', heading: 'Final reminder',
    intro: 'Hi {{customer_name}}, the items from booking {{booking_number}} are still not returned. Late charges are adding up. Please call us right away on {{business_phone}}. Our team will also call you.',
    blocks: ['booking', 'items', 'returnrules'], cta: ['call', 'whatsapp'] },
  late_fee: { name: 'Late return charge', trigger: 'A late fee is added to a booking', kind: 'event',
    subject: 'Late return charge — Booking {{booking_number}}', heading: 'Late return charge',
    intro: 'Hi {{customer_name}}, a late return charge has been added to your booking. Here is the breakdown.',
    blocks: ['booking', 'latefee'], cta: ['call', 'whatsapp'] },
  return_received: { name: 'Return received', trigger: 'Items received, inspection pending', kind: 'event',
    subject: 'We have received your OMG Rental items — {{booking_number}}', heading: 'We have received your items',
    intro: 'Hi {{customer_name}}, thank you for returning your outfit. Your items have been received and are currently undergoing inspection. We will email you the final settlement shortly.',
    blocks: ['booking', 'items'], cta: ['call'] },
  settlement: { name: 'Return settlement + deposit refund', trigger: 'Return closed after inspection', kind: 'event', pdf: true,
    subject: 'Your OMG Rental return is complete — Booking {{booking_number}}', heading: 'Return complete',
    intro: 'Hi {{customer_name}}, your return has been inspected and settled. Here is the full breakdown of your deposit. The detailed receipt is attached.',
    blocks: ['booking', 'inspection', 'settle'], cta: ['call'] },
  cancellation: { name: 'Booking cancellation', trigger: 'A booking is cancelled', kind: 'event',
    subject: 'Booking cancelled — {{booking_number}}', heading: 'Your booking has been cancelled',
    intro: 'Hi {{customer_name}}, your booking {{booking_number}} has been cancelled. If you have any question about a refund, please call us.',
    blocks: ['booking', 'items', 'cancel'], cta: ['call', 'whatsapp'] },
  sale_receipt: { name: 'Purchase receipt', trigger: 'An item is sold to a customer', kind: 'event', pdf: true,
    subject: 'Thank you for your purchase — OMG Rental', heading: 'Thank you for your purchase',
    intro: 'Hi {{customer_name}}, thank you for shopping with OMG Rental! Your receipt is attached.',
    blocks: ['sale'], cta: ['call', 'whatsapp'] },
  review_request: { name: 'Review request', trigger: 'After the return is closed (delay in hours)', kind: 'auto',
    subject: 'How was your OMG Rental experience?', heading: 'How was your experience?',
    intro: 'Hi {{customer_name}}, thank you for renting with us! We would love to hear how it went. It takes less than a minute.',
    blocks: ['review'], cta: [] },
  staff_digest: { name: 'Daily summary for staff', trigger: 'Every morning, to the staff email', kind: 'staff', timeKey: 'staffDigestTime',
    subject: 'OMG Rental today, {{today}}: {{overdue_count}} overdue · {{returns_today}} returns · {{pickups_today}} pickups', heading: 'Today at OMG Rental',
    intro: 'Good morning! Here is what needs attention today.',
    blocks: ['staff'], cta: [] },
  payment_failed: { name: 'Payment failed (online payments)', trigger: 'For future online payments', kind: 'future', off: true,
    subject: 'Payment could not be completed — {{booking_number}}', heading: 'Payment could not be completed',
    intro: 'Hi {{customer_name}}, your payment of {{payment_amount}} could not be completed. Please try again or call us.',
    blocks: ['booking'], cta: ['call'] }
};

/* ---------- small helpers ---------- */
function adminOnly(u) { if (u.role !== 'admin') fail('Only the owner (admin) can do this', 'forbidden'); return u; }
function escH(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
function rsE(n) { n = Math.round(+n || 0); var neg = n < 0; n = Math.abs(n); var s = String(n), last = s.slice(-3), rest = s.slice(0, -3); if (rest) last = ',' + last; rest = rest.replace(/\B(?=(\d{2})+(?!\d))/g, ','); return (neg ? '−' : '') + '₹' + rest + last; }
function toDate(v) { if (!v) return null; var d = v instanceof Date ? v : new Date(v); return isNaN(d.getTime()) ? null : d; }
function fmtD(v, f) { var d = toDate(v); return d ? Utilities.formatDate(d, TZ, f) : ''; }
function ymdIST(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function atIST(ymd, hm) { return new Date(ymd + 'T' + (hm || '00:00') + ':00+05:30'); }
function hmLabel(hm) { var p = String(hm || '12:00').split(':'), h = +p[0], m = p[1] || '00'; return (h % 12 || 12) + ':' + m + ' ' + (h < 12 ? 'AM' : 'PM'); }
function findRow(rows, k, v) { for (var i = 0; i < (rows || []).length; i++) if (String(rows[i][k]) === String(v)) return rows[i]; return null; }
function settingsObj(T) { var o = {}; (T.Settings || []).forEach(function (r) { try { o[r.key] = JSON.parse(String(r.value)); } catch (x) { o[r.key] = r.value; } }); return o; }
function emailSettings(T) {
  var s = settingsObj(T), em = {}, k;
  for (k in EMAIL_DEFAULTS) em[k] = EMAIL_DEFAULTS[k];
  var saved = s.email || {};
  for (k in saved) if (saved[k] !== '' && saved[k] != null) em[k] = saved[k];
  em.templates = s.emailTemplates || {};
  em.app = s;
  return em;
}
function tplFor(key, em, over) {
  var d = TEMPLATES[key]; if (!d) return null;
  var o = over || (em.templates || {})[key] || {}, t = {}, k;
  for (k in d) t[k] = d[k];
  ['subject', 'heading', 'intro'].forEach(function (f) { if (o[f]) t[f] = o[f]; });
  t.enabled = o.enabled != null ? !!o.enabled : !d.off;
  return t;
}
function validEmail(s) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(s || '').trim()); }
function unsubToken(cid) {
  var salt = PropertiesService.getScriptProperties().getProperty('SALT') || '';
  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':unsub:' + cid)).slice(0, 16);
}

/* ---------- data for one email ---------- */
function emailCtx(T, bookingNo, saleId, customerId) {
  var b = bookingNo ? findRow(T.Bookings, 'booking_no', bookingNo) : null;
  var sale = saleId ? findRow(T.Sales, 'sale_id', saleId) : null;
  var cid = customerId || (b && b.customer_id) || (sale && sale.customer_id);
  var c = findRow(T.Customers, 'customer_id', cid) || { name: (b && b.customer_name) || '' };
  var items = {}; (T.Items || []).forEach(function (i) { items[i.item_code] = i; });
  var no = b && b.booking_no;
  return {
    b: b, sale: sale, c: c, items: items,
    lines: no ? (T.Booking_Items || []).filter(function (l) { return l.booking_no === no; }) : [],
    pays: no ? (T.Payments || []).filter(function (p) { return p.booking_no === no; }) : [],
    ret: no ? findRow(T.Returns, 'booking_no', no) : null,
    retItems: no ? (T.Return_Items || []).filter(function (r) { return r.booking_no === no; }) : [],
    T: T
  };
}
function moneyOf(ctx) {
  var b = ctx.b || {}, P = function (k) { return ctx.pays.filter(function (p) { return p.kind === k; }).reduce(function (a, p) { return a + (+p.amount || 0); }, 0); };
  var rentNet = b.rent_net !== '' && b.rent_net != null ? +b.rent_net : (+b.rent_gross || 0) - (+b.discount || 0);
  var depPaid = b.deposit_paid !== '' && b.deposit_paid != null ? +b.deposit_paid : P('deposit');
  var paid = b.total_paid !== '' && b.total_paid != null ? +b.total_paid : P('rent') + P('deposit');
  var last = ctx.pays.filter(function (p) { return p.kind === 'rent' || p.kind === 'deposit'; }).slice(-1)[0];
  return { gross: +b.rent_gross || 0, discount: +b.discount || 0, rentNet: rentNet, depDue: +b.deposit_due || 0, depPaid: depPaid, paid: paid,
    balance: b.balance_due !== '' && b.balance_due != null ? +b.balance_due : Math.max(0, rentNet + (+b.deposit_due || 0) - paid),
    mode: b.payment_mode || (last && last.mode) || '' };
}
function emailVars(ctx, em) {
  var b = ctx.b || {}, c = ctx.c || {}, m = moneyOf(ctx), r = ctx.ret || {}, app = em.app || {};
  var damage = ctx.retItems.reduce(function (a, x) { return a + (+x.charge || 0); }, 0);
  var itemList = ctx.lines.map(function (l) { var it = ctx.items[l.item_code] || {}; return l.item_code + (it.name ? ' ' + it.name : ''); }).join(', ');
  var pickupEnd = app.pickupEnd || '15:00';
  var v = {
    customer_name: String(c.name || 'there').split(' ')[0], customer_full_name: c.name || '', booking_number: b.booking_no || (ctx.sale && ctx.sale.sale_id) || '',
    reference_number: b.booking_no || (ctx.sale && ctx.sale.sale_id) || '', item_list: itemList,
    pickup_date: fmtD(b.pickup_at, 'EEE, d MMMM yyyy'), pickup_time: b.pickup_at ? fmtD(b.pickup_at, 'h:mm a') + ' – ' + hmLabel(pickupEnd) : '',
    event_date: fmtD(b.event_date, 'EEE, d MMMM yyyy'), return_date: fmtD(b.return_by, 'EEE, d MMMM yyyy'), return_time: fmtD(b.return_by, 'h:mm a'),
    rental_amount: rsE(m.rentNet), discount: rsE(m.discount), deposit_amount: rsE(m.depDue), deposit_paid: rsE(m.depPaid),
    damage_charge: rsE(damage), late_fee: rsE(r.late_fee_waived === 'yes' ? 0 : (+r.late_fee || +b.late_flag_amount || 0)), refund_amount: rsE(+r.refund || 0),
    payment_amount: rsE(m.paid), payment_method: m.mode || '—', balance_due: rsE(m.balance),
    business_phone: em.businessPhone, business_address: em.businessAddress, receipt_link: '', today: Utilities.formatDate(new Date(), TZ, 'EEE d MMM')
  };
  if (ctx.sale) { v.payment_amount = rsE(ctx.sale.price); v.payment_method = ctx.sale.mode || '—'; var si = ctx.items[ctx.sale.item_code] || {}; v.item_list = ctx.sale.item_code + (si.name ? ' ' + si.name : ''); }
  if (ctx.staff) { v.overdue_count = ctx.staff.overdue.length; v.returns_today = ctx.staff.returns.length; v.pickups_today = ctx.staff.pickups.length; }
  return v;
}
function fillVars(s, v, html) { return String(s || '').replace(/\{\{\s*([a-z_\/]+)\s*\}\}/g, function (_, k) { var x = v[k] != null ? v[k] : ''; return html ? escH(x) : String(x); }); }

/* ---------- HTML building blocks (tables + inline styles = works in Gmail, Outlook, Apple Mail) ---------- */
var BR = { maroon: '#7a1f4b', deep: '#4a0f2c', gold: '#e0b04e', cream: '#fbf6ef', ink: '#2b1320', muted: '#7d6b73', line: '#eee2d6' };
function eRow(label, value, strong, color) {
  return '<tr><td style="padding:9px 0;border-bottom:1px solid ' + BR.line + ';font:14px/1.4 Arial,Helvetica,sans-serif;color:' + BR.muted + '">' + label + '</td>' +
    '<td align="right" style="padding:9px 0;border-bottom:1px solid ' + BR.line + ';font:' + (strong ? '700 15px' : '600 14px') + '/1.4 Arial,Helvetica,sans-serif;color:' + (color || BR.ink) + '">' + value + '</td></tr>';
}
function eCard(title, inner) {
  return '<tr><td style="padding:8px 28px"><table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + BR.cream + ';border:1px solid ' + BR.line + ';border-radius:14px"><tr><td style="padding:16px 18px">' +
    (title ? '<div style="font:700 11px/1.4 Arial,Helvetica,sans-serif;letter-spacing:2px;text-transform:uppercase;color:' + BR.maroon + ';margin-bottom:6px">' + title + '</div>' : '') +
    inner + '</td></tr></table></td></tr>';
}
function eTable(rows) { return '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + rows.join('') + '</table>'; }
/* small PNG icons hosted with the app (emoji and SVG do not show in every mail app) */
function eIcon(em, name, size) { var base = String(em.appUrl || 'https://omgrental.netlify.app/'); if (base.slice(-1) !== '/') base += '/'; size = size || 18; return '<img src="' + escH(base + 'email/' + name + '.png') + '" width="' + size + '" height="' + size + '" alt="" style="display:inline-block;width:' + size + 'px;height:' + size + 'px;vertical-align:middle;border:0;margin:-3px 8px 0 0">'; }
function eBtn(href, label, bg, fg) {
  return '<table role="presentation" cellpadding="0" cellspacing="0" style="display:inline-block;margin:4px 6px 4px 0"><tr><td bgcolor="' + bg + '" style="border-radius:10px;background:' + bg + '">' +
    '<a href="' + escH(href) + '" style="display:inline-block;padding:12px 20px;font:700 14px/1 Arial,Helvetica,sans-serif;color:' + (fg || '#ffffff') + ';text-decoration:none;border-radius:10px">' + label + '</a></td></tr></table>';
}
function itemRows(ctx) {
  return ctx.lines.map(function (l) {
    var it = ctx.items[l.item_code] || {};
    return '<tr><td style="padding:8px 0;border-bottom:1px solid ' + BR.line + '"><span style="display:inline-block;font:700 12px/1 Courier New,monospace;background:#ffffff;border:1px solid ' + BR.line + ';border-radius:6px;padding:5px 7px;color:' + BR.ink + '">' + escH(l.item_code) + '</span>' +
      '<span style="font:14px/1.4 Arial,Helvetica,sans-serif;color:' + BR.ink + ';padding-left:8px">' + escH(it.name || '') + '</span></td></tr>';
  }).join('');
}
function renderBlock(name, ctx, v, em, meta) {
  var b = ctx.b || {}, m = ctx.b ? moneyOf(ctx) : null, e = escH;
  switch (name) {
    case 'booking': return b.booking_no ? eCard('', eTable([eRow('Booking', '<span style="font-family:Courier New,monospace">#' + e(b.booking_no) + '</span>', true), eRow('Customer', e(ctx.c.name || ''))])) : '';
    case 'dates': return b.booking_no ? eCard('Your dates', eTable([eRow('Pickup', e(v.pickup_date) + '<br><span style="color:' + BR.maroon + '">' + e(v.pickup_time) + '</span>'), eRow('Event', e(v.event_date)), eRow('Return', e(v.return_date) + '<br><span style="color:' + BR.maroon + '">Before ' + e(v.return_time) + '</span>', true)])) : '';
    case 'items': return ctx.lines.length ? eCard('Your outfit', '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + itemRows(ctx) + '</table>') : '';
    case 'money': if (!m) return ''; return eCard('Payment', eTable([
      eRow('Rental amount', rsE(m.gross)), m.discount ? eRow('Discount', '−' + rsE(m.discount), false, '#1f7a3a') : '', m.discount ? eRow('Rental after discount', rsE(m.rentNet)) : '',
      eRow('Security deposit (refundable)', rsE(m.depDue)), eRow('Total paid' + (m.mode ? ' · ' + e(m.mode) : ''), rsE(m.paid), true),
      m.balance > 0 ? eRow('Balance due', rsE(m.balance), true, '#b3261e') : '']));
    case 'balance': if (!m || m.balance <= 0) return ''; return eCard('Payment', eTable([eRow('Balance to pay at pickup', rsE(m.balance), true, '#b3261e'), eRow('Accepted', 'Cash · UPI · Card')]));
    case 'instructions': return eCard('Good to know', '<ul style="margin:0;padding-left:18px;font:14px/1.7 Arial,Helvetica,sans-serif;color:' + BR.ink + '">' +
      '<li>Pickup on your event day, ' + e(v.pickup_time) + '. Please bring a photo ID.</li><li>Return the next day <b>before ' + e(v.return_time) + '</b>. Late returns are charged as per our rental terms.</li>' +
      '<li>Please avoid perfume, food and drink stains on the outfit. Do not wash or iron it; we take care of cleaning.</li><li>Your deposit is refunded after the items are returned and checked.</li></ul>');
    case 'returnrules': return eCard('Return', '<div style="font:15px/1.6 Arial,Helvetica,sans-serif;color:' + BR.ink + '">Return deadline: <b style="color:' + BR.maroon + '">' + e(v.return_date) + ', before ' + e(v.return_time) + '</b><br><span style="color:' + BR.muted + ';font-size:13px">Please bring every piece, including jewellery and accessories. Late returns are charged as per the rental terms.</span></div>');
    case 'latefee': var lf = +b.late_flag_amount || 0; return eCard('Late charge', eTable([eRow('Return deadline', e(v.return_date) + ', ' + e(v.return_time)), eRow('Late charge', rsE(lf), true, '#b3261e'), eRow('Security deposit paid', rsE(m.depPaid)), eRow('Deposit after late charge', rsE(Math.max(0, m.depPaid - lf)), true)]));
    case 'inspection':
      if (!ctx.retItems.length) return '';
      return eCard('Inspection', '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + ctx.retItems.map(function (x) {
        var it = ctx.items[x.item_code] || {}, cond = { good: 'Good', minor: 'Minor damage', stain: 'Stain', repair: 'Needs repair', major: 'Major damage', missing: 'Missing' }[x.condition] || x.condition || 'Good';
        return eRow('<b style="font-family:Courier New,monospace;color:' + BR.ink + '">' + e(x.item_code) + '</b> ' + e(it.name || ''), e(cond) + (+x.charge ? ' · ' + rsE(x.charge) : ''), false, +x.charge ? '#b3261e' : '#1f7a3a');
      }).join('') + '</table>');
    case 'settle':
      var r = ctx.ret || {}, dmg = ctx.retItems.reduce(function (a, x) { return a + (+x.charge || 0); }, 0), late = r.late_fee_waived === 'yes' ? 0 : (+r.late_fee || 0);
      return eCard('Deposit settlement', eTable([eRow('Deposit received', rsE(m.depPaid)), eRow('Damage deduction', dmg ? '−' + rsE(dmg) : rsE(0), false, dmg ? '#b3261e' : null),
        eRow('Late fee', late ? '−' + rsE(late) : rsE(0) + (r.late_fee_waived === 'yes' ? ' (waived)' : ''), false, late ? '#b3261e' : null),
        +r.rent_deducted ? eRow('Balance rent deducted', '−' + rsE(r.rent_deducted)) : '', +r.extra_collected ? eRow('Extra collected', rsE(r.extra_collected) + (r.collect_mode ? ' · ' + e(r.collect_mode) : '')) : '',
        eRow('Refunded to you', rsE(r.refund) + (r.refund_mode ? ' · ' + e(r.refund_mode) : ''), true, '#1f7a3a')]) +
        '<div style="font:13px/1.5 Arial,Helvetica,sans-serif;color:' + BR.muted + ';margin-top:10px">This booking is now fully settled. Thank you!</div>');
    case 'cancel': return m ? eCard('Refund', eTable([eRow('Amount paid', rsE(m.paid)), eRow('Status', 'Please call us for your refund details')])) : '';
    case 'changes':
      var p = (meta && meta.prev) || {};
      return eCard('What changed', eTable([eRow('Before', e([fmtD(p.pickup, 'd MMM, h:mm a') && 'Pickup ' + fmtD(p.pickup, 'd MMM, h:mm a'), fmtD(p.ret, 'd MMM, h:mm a') && 'Return ' + fmtD(p.ret, 'd MMM, h:mm a'), p.rent != null && p.rent !== '' ? 'Rent ' + rsE(p.rent) : ''].filter(String).join(' · ')) || '—'),
        eRow('Now', e(['Pickup ' + fmtD(b.pickup_at, 'd MMM, h:mm a'), 'Return ' + fmtD(b.return_by, 'd MMM, h:mm a'), 'Rent ' + rsE(m.rentNet)].join(' · ')), true)]));
    case 'sale':
      var s = ctx.sale || {}, si = ctx.items[s.item_code] || {};
      return eCard('Receipt', eTable([eRow('Receipt', '<span style="font-family:Courier New,monospace">#' + e(s.sale_id) + '</span>', true), eRow('Product', e(si.name || '') + ' <span style="font-family:Courier New,monospace">(' + e(s.item_code) + ')</span>'), eRow('Quantity', '1'),
        eRow('Price', rsE(s.list_price)), +s.discount ? eRow('Discount', '−' + rsE(s.discount), false, '#1f7a3a') : '', eRow('Total paid' + (s.mode ? ' · ' + e(s.mode) : ''), rsE(s.price), true), eRow('Date', e(fmtD(s.at, 'd MMM yyyy, h:mm a')))]));
    case 'review':
      return '<tr><td align="center" style="padding:10px 28px 4px">' + (em.reviewLink ? eBtn(em.reviewLink, '★ Leave a Google review', BR.maroon) : '') +
        eBtn('mailto:' + (em.feedbackEmail || em.replyTo || em.senderEmail || '') + '?subject=' + encodeURIComponent('Feedback on booking ' + v.booking_number), 'Send feedback', '#ffffff', BR.maroon) + '</td></tr>';
    case 'staff':
      var st = ctx.staff, list = function (title, rows, color) {
        return eCard(title + ' · ' + rows.length, rows.length ? eTable(rows.map(function (x) { return eRow('<b style="color:' + BR.ink + '">' + e(x.name) + '</b> · ' + e(x.no) + (x.mobile ? ' · <a href="tel:+91' + e(x.mobile) + '" style="color:' + BR.maroon + '">' + e(x.mobile) + '</a>' : ''), e(x.note), false, color); })) : '<div style="font:14px Arial,Helvetica,sans-serif;color:' + BR.muted + '">None</div>');
      };
      return list('Overdue — call now', st.overdue, '#b3261e') + list('Returns due today', st.returns) + list('Pickups today', st.pickups) + list('Balance to collect', st.balance, '#b3261e') + list('Awaiting inspection', st.inspect) +
        (st.failed ? eCard('Emails', '<div style="font:14px Arial,Helvetica,sans-serif;color:#b3261e">' + st.failed + ' email(s) failed in the last day. Check Settings → Email & reminders.</div>') : '');
  }
  return '';
}
function renderEmail(key, ctx, em, meta, over) {
  var t = tplFor(key, em, over); if (!t) return null;
  var v = emailVars(ctx, em), phone = String(em.businessPhone || '').replace(/\D/g, '');
  var subject = fillVars(t.subject, v, false);
  var intro = fillVars(t.intro, v, true).replace(/\n/g, '<br>');
  var cta = (t.cta || []).map(function (c) {
    if (c === 'call' && phone) return eBtn('tel:+91' + phone.slice(-10), eIcon(em, 'call-white') + 'Call ' + escH(em.businessPhone), BR.maroon);
    if (c === 'whatsapp' && phone) return eBtn('https://wa.me/91' + phone.slice(-10) + '?text=' + encodeURIComponent('Hi OMG Rental, about booking ' + v.booking_number), eIcon(em, 'wa-white') + 'WhatsApp us', '#1f9d55');
    if (c === 'directions' && em.mapsLink) return eBtn(em.mapsLink, eIcon(em, 'pin-maroon') + 'Get directions', '#ffffff', BR.maroon);
    return '';
  }).join('');
  var base = ''; try { base = ScriptApp.getService().getUrl() || ''; } catch (x) { }
  var unsub = base && ctx.c && ctx.c.customer_id ? base + '?unsub=' + encodeURIComponent(ctx.c.customer_id) + '&t=' + unsubToken(ctx.c.customer_id) : '';
  var html = '<!DOCTYPE html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="light only"><meta name="supported-color-schemes" content="light only"><title>' + escH(subject) + '</title></head>' +
    '<body style="margin:0;padding:0;background:#f3e9de;-webkit-text-size-adjust:100%">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0">' + escH(fillVars(t.intro, v, false)).slice(0, 140) + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" bgcolor="#f3e9de" style="background:#f3e9de"><tr><td align="center" style="padding:24px 10px">' +
    '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="width:100%;max-width:600px;background:#ffffff;border-radius:20px;overflow:hidden;border:1px solid ' + BR.line + '">' +
    '<tr><td align="center" bgcolor="' + BR.maroon + '" style="background:' + BR.maroon + ';background-image:linear-gradient(135deg,' + BR.deep + ',' + BR.maroon + ' 55%,#a8325f);padding:30px 24px 24px">' +
    (em.logoUrl ? '<img src="' + escH(em.logoUrl) + '" width="96" height="96" alt="OMG Rental" style="display:block;width:96px;height:96px;border-radius:48px;border:3px solid ' + BR.gold + ';background:#ffffff">' : '') +
    '<div style="font:700 22px/1.2 Georgia,\'Times New Roman\',serif;color:#ffffff;margin-top:14px;letter-spacing:1px">OMG Rental</div>' +
    '<div style="font:700 11px/1.4 Arial,Helvetica,sans-serif;letter-spacing:3px;color:' + BR.gold + ';margin-top:4px;text-transform:uppercase">Garba &amp; festive outfits · Indore</div></td></tr>' +
    '<tr><td style="height:4px;background:' + BR.gold + ';line-height:4px;font-size:0">&nbsp;</td></tr>' +
    '<tr><td style="padding:28px 28px 10px"><h1 style="margin:0 0 12px;font:700 25px/1.25 Georgia,\'Times New Roman\',serif;color:' + BR.ink + '">' + escH(fillVars(t.heading, v, false)) + '</h1>' +
    '<p style="margin:0;font:15px/1.65 Arial,Helvetica,sans-serif;color:#4b3a42">' + intro + '</p></td></tr>' +
    (t.blocks || []).map(function (bn) { return renderBlock(bn, ctx, v, em, meta); }).join('') +
    (cta ? '<tr><td style="padding:16px 28px 6px">' + cta + '</td></tr>' : '') +
    '<tr><td style="padding:22px 28px 26px"><div style="font:14px/1.6 Arial,Helvetica,sans-serif;color:#4b3a42">With love,<br><b>Team OMG Rental</b></div></td></tr>' +
    '<tr><td bgcolor="' + BR.cream + '" style="background:' + BR.cream + ';padding:20px 28px;border-top:1px solid ' + BR.line + '">' +
    '<div style="font:700 14px/1.5 Arial,Helvetica,sans-serif;color:' + BR.ink + '">OMG Rental · linked with OMG Salon</div>' +
    '<div style="font:13px/1.6 Arial,Helvetica,sans-serif;color:' + BR.muted + '">' + escH(em.businessAddress) + '<br>' + eIcon(em, 'call-maroon', 14) + '<a href="tel:+91' + phone.slice(-10) + '" style="color:' + BR.maroon + ';text-decoration:none;font-weight:700">' + escH(em.businessPhone) + '</a>' +
    (em.appUrl ? ' · <a href="' + escH(em.appUrl) + '#/catalog" style="color:' + BR.maroon + '">See our collection</a>' : '') + '</div>' +
    '<div style="font:11px/1.5 Arial,Helvetica,sans-serif;color:#a3949a;margin-top:10px">You are receiving this email about your booking with OMG Rental.' + (unsub ? ' <a href="' + escH(unsub) + '" style="color:#a3949a">Stop these emails</a>.' : '') + '</div></td></tr>' +
    '</table></td></tr></table></body></html>';
  var text = subject + '\n\n' + fillVars(t.intro, v, false) + '\n\nBooking: ' + v.booking_number + '\nCall: ' + em.businessPhone + '\n' + em.businessAddress;
  return { subject: subject, html: html, text: text, pdf: !!t.pdf, enabled: t.enabled };
}

/* ---------- PDF receipt ---------- */
function logoDataUri(em) {
  var cache = CacheService.getScriptCache(), hit = cache.get('logo_b64');
  if (hit) return hit;
  try { var r = UrlFetchApp.fetch(em.logoUrl, { muteHttpExceptions: true }); if (r.getResponseCode() !== 200) return ''; var u = 'data:image/jpeg;base64,' + Utilities.base64Encode(r.getContent()); if (u.length < 90000) cache.put('logo_b64', u, 21600); return u; } catch (x) { return ''; }
}
function receiptPdf(ctx, em, key) {
  var v = emailVars(ctx, em), b = ctx.b, e = escH, row = function (l, r, bold) { return '<tr><td style="padding:6px 0;border-bottom:1px solid #eee">' + l + '</td><td style="padding:6px 0;border-bottom:1px solid #eee;text-align:right;' + (bold ? 'font-weight:bold' : '') + '">' + r + '</td></tr>'; };
  var body = '';
  if (ctx.sale) {
    var s = ctx.sale, si = ctx.items[s.item_code] || {};
    body = '<h2>Purchase receipt #' + e(s.sale_id) + '</h2><p>Date: ' + e(fmtD(s.at, 'd MMM yyyy, h:mm a')) + '<br>Customer: ' + e(ctx.c.name || s.walk_in || '') + '</p><table width="100%" cellspacing="0">' +
      row('<b>' + e(s.item_code) + '</b> ' + e(si.name || '') + ' × 1', rsE(s.list_price)) + (+s.discount ? row('Discount', '−' + rsE(s.discount)) : '') + row('Total paid (' + e(s.mode || '') + ')', rsE(s.price), true) + '</table>';
  } else if (b) {
    var m = moneyOf(ctx);
    body = '<h2>' + (key === 'settlement' ? 'Settlement receipt' : 'Rental receipt') + ' #' + e(b.booking_no) + '</h2><p>Customer: <b>' + e(ctx.c.name || '') + '</b>' + (ctx.c.mobile ? ' · ' + e(ctx.c.mobile) : '') +
      '<br>Pickup: ' + e(v.pickup_date) + ', ' + e(v.pickup_time) + '<br>Event: ' + e(v.event_date) + '<br>Return: ' + e(v.return_date) + ', before ' + e(v.return_time) + '</p>' +
      '<table width="100%" cellspacing="0"><tr><th align="left" style="border-bottom:2px solid #7a1f4b;padding:6px 0">Item</th><th align="right" style="border-bottom:2px solid #7a1f4b;padding:6px 0">Rent</th></tr>' +
      ctx.lines.map(function (l) { var it = ctx.items[l.item_code] || {}; return row('<b>' + e(l.item_code) + '</b> ' + e(it.name || '') + (l.package_name ? ' <i>(' + e(l.package_name) + ')</i>' : ''), rsE(l.allocated_price)); }).join('') +
      row('Rental amount', rsE(m.gross)) + (m.discount ? row('Discount', '−' + rsE(m.discount)) : '') + row('Rental total', rsE(m.rentNet), true) + row('Security deposit (refundable)', rsE(m.depDue)) + '</table>' +
      '<h3>Payments</h3><table width="100%" cellspacing="0">' + (ctx.pays.map(function (p) { return row(e({ rent: 'Rent', deposit: 'Deposit', topup: 'Extra collected', refund: 'Deposit refund', damage: 'Damage charge', late: 'Late fee', rent_refund: 'Rent refund' }[p.kind] || p.kind) + ' · ' + e(p.mode || '') + ' · ' + e(fmtD(p.at, 'd MMM, h:mm a')), (p.kind === 'refund' || p.kind === 'rent_refund' ? '−' : '') + rsE(p.amount)); }).join('') || row('No payments yet', '')) +
      row('Total paid', rsE(m.paid), true) + (m.balance > 0 ? row('Balance due', rsE(m.balance), true) : '') + '</table>';
    if (key === 'settlement' && ctx.ret) {
      var r = ctx.ret, dmg = ctx.retItems.reduce(function (a, x) { return a + (+x.charge || 0); }, 0);
      body += '<h3>Deposit settlement</h3><table width="100%" cellspacing="0">' + ctx.retItems.map(function (x) { return row(e(x.item_code) + ' — ' + e(x.condition || 'good') + (x.note ? ' (' + e(x.note) + ')' : ''), +x.charge ? '−' + rsE(x.charge) : '—'); }).join('') +
        row('Deposit received', rsE(m.depPaid)) + row('Damage', '−' + rsE(dmg)) + row('Late fee' + (r.late_fee_waived === 'yes' ? ' (waived)' : ''), '−' + rsE(r.late_fee_waived === 'yes' ? 0 : r.late_fee)) + row('Refunded (' + e(r.refund_mode || '') + ')', rsE(r.refund), true) + '</table>';
    }
  }
  var logo = logoDataUri(em);
  var html = '<html><head><meta charset="utf-8"><style>body{font-family:Arial,Helvetica,sans-serif;color:#2b1320;font-size:12px;margin:28px}h1{font-family:Georgia,serif;color:#7a1f4b;margin:0}h2{font-size:17px;margin:18px 0 6px}h3{font-size:14px;margin:16px 0 4px;color:#7a1f4b}td,th{font-size:12px}</style></head><body>' +
    '<table width="100%"><tr><td>' + (logo ? '<img src="' + logo + '" width="70" height="70">' : '') + '</td><td align="right"><h1>OMG Rental</h1><div>' + e(em.businessAddress) + '<br>Phone ' + e(em.businessPhone) + '</div></td></tr></table><hr style="border:0;border-top:3px solid #e0b04e">' +
    body + '<p style="margin-top:24px;color:#7d6b73;font-size:11px">Issued ' + e(Utilities.formatDate(new Date(), TZ, 'd MMM yyyy, h:mm a')) + ' IST. Security deposits are refundable after the items are returned and inspected. Thank you for choosing OMG Rental!</p></body></html>';
  var name = 'OMG-Rental-' + (ctx.sale ? 'Receipt-' + ctx.sale.sale_id : (key === 'settlement' ? 'Settlement-' : 'Receipt-') + b.booking_no) + '.pdf';
  return Utilities.newBlob(html, 'text/html', 'receipt.html').getAs('application/pdf').setName(name);
}

/* ---------- queue: emails caused by a save (called inside commit, under the lock) ---------- */
function emailLogRows() { return sheetObjects('Email_Log'); }
function logIndex(rows) {
  var byKey = {}; rows.forEach(function (r) { if (r.status !== 'failed' || +r.attempts < 3) byKey[r.dedup_key] = r; }); return byKey;
}
function bookingFp(b) { return [toDate(b.pickup_at) ? toDate(b.pickup_at).getTime() : '', toDate(b.return_by) ? toDate(b.return_by).getTime() : '', +b.rent_gross || 0, +b.discount || 0].join('|'); }
function queueEventEmails(changes) {
  var B = changes.Bookings || [], R = changes.Returns || [], S = changes.Sales || [];
  if (!B.length && !R.length && !S.length) return 0;
  var em = emailSettings({ Settings: sheetObjects('Settings') }); if (!em.enabled) return 0;
  var rows = emailLogRows(), idx = logIndex(rows), now = Date.now(), q = [];
  var reach = {}; sheetObjects('Customers').forEach(function (c) { reach[c.customer_id] = validEmail(c.email) && String(c.email_opt).toLowerCase() !== 'no'; });
  (changes.Customers || []).forEach(function (c) { if (c.email != null) reach[c.customer_id] = validEmail(c.email) && String(c.email_opt).toLowerCase() !== 'no'; });
  var recent = function (v) { var d = toDate(v); return d && now - d.getTime() < 3 * 864e5 && d.getTime() - now < 864e5; };
  var add = function (tpl, key, o) { if (idx[key] || !tplFor(tpl, em).enabled || !reach[o.customer_id]) return; idx[key] = 1; q.push(newLogRow(tpl, key, o)); };
  var sheetB = null, fullB = function (b) { if (b.customer_id && b.created_at) return b; sheetB = sheetB || sheetObjects('Bookings'); var o = findRow(sheetB, 'booking_no', b.booking_no) || {}, x = {}, k; for (k in o) x[k] = o[k]; for (k in b) x[k] = b[k]; return x; };
  B.forEach(function (b0) {
    var b = fullB(b0), no = b.booking_no, fp = bookingFp(b), meta = { fp: fp, pickup: b.pickup_at, ret: b.return_by, rent: (+b.rent_gross || 0) - (+b.discount || 0) };
    var confirmed = idx['booking_confirmation|' + no];
    if ((b.status === 'reserved' || b.status === 'picked') && !confirmed && recent(b.created_at)) add('booking_confirmation', 'booking_confirmation|' + no, { booking_no: no, customer_id: b.customer_id, meta: meta });
    else if ((b.status === 'reserved' || b.status === 'picked') && confirmed) {
      var last = rows.filter(function (r) { return r.booking_no === no && (r.template === 'booking_confirmation' || r.template === 'booking_updated'); }).pop();
      var lm = {}; try { lm = JSON.parse(last.meta || '{}'); } catch (x) { }
      if (lm.fp && lm.fp !== fp) add('booking_updated', 'booking_updated|' + no + '|' + fp, { booking_no: no, customer_id: b.customer_id, meta: { fp: fp, prev: lm } });
    }
    if ((b.status === 'cancelled' || b.status === 'void') && confirmed) add('cancellation', 'cancellation|' + no, { booking_no: no, customer_id: b.customer_id });
    if (b.status === 'returned_pending' && recent(b.received_at)) add('return_received', 'return_received|' + no, { booking_no: no, customer_id: b.customer_id });
    if (+b.late_flag_amount > 0 && b.status === 'picked' && recent(b.late_flag_at)) add('late_fee', 'late_fee|' + no, { booking_no: no, customer_id: b.customer_id });
  });
  R.forEach(function (r) { if (r.inspected_at && recent(r.inspected_at)) { var bk = fullB(findRow(B, 'booking_no', r.booking_no) || { booking_no: r.booking_no }); add('settlement', 'settlement|' + r.booking_no, { booking_no: r.booking_no, customer_id: bk.customer_id || '' }); } });
  S.forEach(function (s) { if (s.customer_id && String(s.voided) !== 'yes' && recent(s.at)) add('sale_receipt', 'sale_receipt|' + s.sale_id, { sale_id: s.sale_id, customer_id: s.customer_id }); });
  if (q.length) upsertObjects('Email_Log', q);
  return q.length;
}
function newLogRow(tpl, key, o) {
  return { email_id: 'EM-' + Utilities.getUuid().slice(0, 12), dedup_key: key, template: tpl, channel: 'email', booking_no: o.booking_no || '', sale_id: o.sale_id || '', customer_id: o.customer_id || '',
    to: o.to || '', subject: '', status: 'pending', created_at: new Date().toISOString(), sent_at: '', attempts: 0, error: '', meta: o.meta ? JSON.stringify(o.meta) : '' };
}
/* run the queue ~20 s after a save (one-off trigger). If triggers are not allowed yet, the 15-minute run picks it up. */
function scheduleEmailRun() {
  try {
    var c = CacheService.getScriptCache(); if (c.get('email_soon')) return;
    var t = ScriptApp.newTrigger('emailTick').timeBased().after(20 * 1000).create();
    PropertiesService.getScriptProperties().setProperty('EMAIL_ONEOFF_' + t.getUniqueId(), '1');
    c.put('email_soon', '1', 60);
  } catch (x) { }
}
function cleanupOneOffTriggers() {
  try {
    var P = PropertiesService.getScriptProperties(), props = P.getProperties();
    ScriptApp.getProjectTriggers().forEach(function (t) { if (props['EMAIL_ONEOFF_' + t.getUniqueId()]) { ScriptApp.deleteTrigger(t); P.deleteProperty('EMAIL_ONEOFF_' + t.getUniqueId()); } });
  } catch (x) { }
}

/* ---------- timed rules (runs every 15 minutes) ---------- */
function automationCandidates(T, em, now, idx) {
  var out = [], today = ymdIST(now), tomorrow = ymdIST(new Date(now.getTime() + 864e5));
  var on = function (k) { return tplFor(k, em).enabled; };
  var cust = {}; (T.Customers || []).forEach(function (c) { cust[c.customer_id] = c; });
  var reachable = function (b) { var c = cust[b.customer_id]; return c && validEmail(c.email) && String(c.email_opt).toLowerCase() !== 'no'; };
  var add = function (tpl, b) { var key = tpl + '|' + b.booking_no; if (!idx[key]) { idx[key] = 1; out.push(newLogRow(tpl, key, { booking_no: b.booking_no, customer_id: b.customer_id })); } };
  (T.Bookings || []).forEach(function (b) {
    if (!reachable(b)) return;
    var pu = toDate(b.pickup_at), rb = toDate(b.return_by); if (!pu || !rb) return;
    var puDay = ymdIST(pu), rbDay = ymdIST(rb);
    if (b.status === 'reserved') {
      if (on('pickup_tomorrow') && puDay === tomorrow && now >= atIST(today, em.pickupTomorrowTime)) add('pickup_tomorrow', b);
      if (on('pickup_today') && puDay === today && now >= atIST(today, em.pickupTodayTime) && now < atIST(today, (em.app || {}).pickupEnd || '15:00')) add('pickup_today', b);
      if (on('payment_due') && puDay === tomorrow && +b.balance_due > 0 && now >= atIST(today, em.paymentDueTime)) add('payment_due', b);
    }
    if (b.status === 'picked') {
      if (on('return_tomorrow') && rbDay === tomorrow && now >= atIST(today, em.returnTomorrowTime)) add('return_tomorrow', b);
      if (on('return_today') && rbDay === today && now >= atIST(today, em.returnTodayTime) && now < rb) add('return_today', b);
      var grace = (+(em.app || {}).graceMin || 0) * 60000;
      if (now.getTime() > rb.getTime() + grace) {
        var times = [em.overdue1Time, em.overdue2Time, em.overdue3Time].map(function (t) { return atIST(rbDay, t); });
        var level = 0; times.forEach(function (t, i) { if (now >= t && t.getTime() >= rb.getTime()) level = i + 1; });
        if (!level && now.getTime() - rb.getTime() > 864e5) level = 3;
        var higherSent = false; for (var j = level; j <= 3; j++) if (idx['overdue_' + j + '|' + b.booking_no]) higherSent = true;
        if (level && !higherSent && on('overdue_' + level)) add('overdue_' + level, b);   // only the latest level: never a burst
      }
    }
    if (b.status === 'closed' && on('review_request') && em.reviewLink) {
      var cl = toDate(b.closed_at);
      if (cl && now.getTime() >= cl.getTime() + (+em.reviewDelayHours || 24) * 36e5 && now.getTime() < cl.getTime() + 7 * 864e5) add('review_request', b);
    }
  });
  if (em.staffEmail && on('staff_digest') && now >= atIST(today, em.staffDigestTime) && !idx['staff_digest|' + today]) {
    idx['staff_digest|' + today] = 1; out.push(newLogRow('staff_digest', 'staff_digest|' + today, { to: em.staffEmail, meta: { day: today } }));
  }
  return out;
}
function staffSummary(T, now) {
  var today = ymdIST(now), cust = {}; (T.Customers || []).forEach(function (c) { cust[c.customer_id] = c; });
  var x = function (b, note) { var c = cust[b.customer_id] || {}; return { name: c.name || b.customer_name || '', no: b.booking_no, mobile: String(c.mobile || '').replace(/\D/g, ''), note: note }; };
  var s = { overdue: [], returns: [], pickups: [], balance: [], inspect: [], failed: 0 };
  (T.Bookings || []).forEach(function (b) {
    var rb = toDate(b.return_by), pu = toDate(b.pickup_at);
    if (b.status === 'picked' && rb && rb < now) s.overdue.push(x(b, 'was due ' + fmtD(rb, 'd MMM h:mm a')));
    else if (b.status === 'picked' && rb && ymdIST(rb) === today) s.returns.push(x(b, 'before ' + fmtD(rb, 'h:mm a')));
    if (b.status === 'reserved' && pu && ymdIST(pu) === today) s.pickups.push(x(b, fmtD(pu, 'h:mm a')));
    if (['reserved', 'picked', 'returned_pending'].indexOf(b.status) >= 0 && +b.balance_due > 0) s.balance.push(x(b, rsE(b.balance_due)));
    if (b.status === 'returned_pending') s.inspect.push(x(b, 'received ' + fmtD(b.received_at, 'd MMM h:mm a')));
  });
  s.failed = (T.Email_Log || []).filter(function (r) { return r.status === 'failed' && toDate(r.created_at) && now - toDate(r.created_at) < 864e5; }).length;
  return s;
}

/* ---------- sending ---------- */
function sendOne(row, T, em) {
  var ctx = emailCtx(T, row.booking_no, row.sale_id, row.customer_id), to = row.to;
  if (row.template === 'staff_digest') { ctx.staff = staffSummary(T, new Date()); to = to || em.staffEmail; }
  else {
    if (!validEmail(ctx.c.email)) return { status: 'skipped', error: 'No email address for this customer' };
    if (String(ctx.c.email_opt).toLowerCase() === 'no') return { status: 'skipped', error: 'Customer turned off emails' };
    to = String(ctx.c.email).trim();
    if (row.template === 'payment_due' && ctx.b && moneyOf(ctx).balance <= 0) return { status: 'skipped', error: 'Already paid' };
    if (/^(pickup_|return_|overdue_)/.test(row.template) && ctx.b && ['reserved', 'picked'].indexOf(ctx.b.status) < 0) return { status: 'skipped', error: 'Booking status changed (' + ctx.b.status + ')' };
    if (!ctx.b && !ctx.sale) return { status: 'skipped', error: 'Booking not found' };
  }
  var meta = {}; try { meta = JSON.parse(row.meta || '{}'); } catch (x) { }
  var m = renderEmail(row.template, ctx, em, meta);
  if (!m) return { status: 'skipped', error: 'Unknown template' };
  var opts = { htmlBody: m.html, name: em.senderName || 'OMG Rental' };
  if (validEmail(em.replyTo)) opts.replyTo = em.replyTo;
  if (m.pdf) { try { opts.attachments = [receiptPdf(ctx, em, row.template)]; } catch (x) { } }
  deliverMail(to, m.subject, m.text, opts, em);
  return { status: 'sent', to: to, subject: m.subject };
}
/* Uses the business address as "From" when it is set up in Gmail → "Send mail as"; otherwise the script owner's address. */
function deliverMail(to, subject, text, opts, em) {
  if (em.senderEmail && validEmail(em.senderEmail)) {
    var aliases = GmailApp.getAliases().map(function (a) { return String(a).toLowerCase(); });
    var me = String(Session.getEffectiveUser().getEmail()).toLowerCase(), want = String(em.senderEmail).toLowerCase();
    if (want === me || aliases.indexOf(want) >= 0) { if (want !== me) opts.from = em.senderEmail; }
    else throw new Error('Sender ' + em.senderEmail + ' is not set up in Gmail "Send mail as" of ' + me);
  }
  GmailApp.sendEmail(to, subject, text, opts);
}
function emailTick() {
  cleanupOneOffTriggers();
  var ul = LockService.getUserLock(); if (!ul.tryLock(10000)) return;   // one run at a time
  try {
    var now = new Date();
    PropertiesService.getScriptProperties().setProperty('EMAIL_LAST_RUN', now.toISOString());
    var T = loadAll(), em = emailSettings(T); if (!em.enabled) return;
    var rows = T.Email_Log || [], idx = logIndex(rows);
    var fresh = automationCandidates(T, em, now, idx);
    if (fresh.length) { var l1 = LockService.getScriptLock(); l1.waitLock(25000); try { upsertObjects('Email_Log', fresh); } finally { l1.releaseLock(); } }
    var todo = rows.concat(fresh).filter(function (r) { return r.status === 'pending' || (r.status === 'failed' && +r.attempts < 3); }).slice(0, 40);
    var quota = 100; try { quota = MailApp.getRemainingDailyQuota(); } catch (x) { }
    var updates = [];
    todo.forEach(function (r) {
      if (quota < 1) { updates.push({ email_id: r.email_id, error: 'Daily Gmail sending limit reached — will retry' }); return; }
      var res;
      try { res = sendOne(r, T, em); if (res.status === 'sent') quota--; }
      catch (x) { res = { status: 'failed', error: String(x && x.message || x).slice(0, 300) }; }
      updates.push({ email_id: r.email_id, status: res.status, to: res.to || r.to || '', subject: res.subject || r.subject || '', sent_at: res.status === 'sent' ? new Date().toISOString() : '', attempts: (+r.attempts || 0) + 1, error: res.error || '' });
    });
    if (updates.length) { var l2 = LockService.getScriptLock(); l2.waitLock(25000); try { upsertObjects('Email_Log', updates); } finally { l2.releaseLock(); } }
  } finally { ul.releaseLock(); }
}

/* ---------- admin tools called from the app ---------- */
function emailStatus() {
  var out = { ok: true, automation: false, sender: '', aliases: [], quota: null, lastRun: PropertiesService.getScriptProperties().getProperty('EMAIL_LAST_RUN') || '', templates: {} };
  Object.keys(TEMPLATES).forEach(function (k) { var t = TEMPLATES[k]; out.templates[k] = { name: t.name, trigger: t.trigger, kind: t.kind, subject: t.subject, heading: t.heading, intro: t.intro, off: !!t.off, pdf: !!t.pdf }; });
  try {
    out.sender = Session.getEffectiveUser().getEmail();
    out.quota = MailApp.getRemainingDailyQuota();
    out.aliases = GmailApp.getAliases();
    out.automation = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'emailTick' && !PropertiesService.getScriptProperties().getProperty('EMAIL_ONEOFF_' + t.getUniqueId()); });
  } catch (x) { out.needsSetup = true; out.error = 'Email is not switched on yet: open Apps Script, run setupEmail() and click Allow, then deploy a new version.'; }
  return out;
}
function sampleCtx(T) {
  var b = (T.Bookings || []).filter(function (x) { return x.status !== 'void'; }).slice(-1)[0];
  if (b) return emailCtx(T, b.booking_no);
  var now = new Date(), d = function (days, hm) { return atIST(ymdIST(new Date(now.getTime() + days * 864e5)), hm); };
  var fake = { booking_no: 'OMG-00124', customer_id: 'CU-SAMPLE', customer_name: 'Neha Sharma', pickup_at: d(1, '14:00'), event_date: d(1, '00:00'), return_by: d(2, '12:00'), rent_gross: 997, discount: 0, deposit_due: 2000, rent_net: 997, deposit_paid: 2000, total_paid: 2997, balance_due: 0, payment_mode: 'UPI', status: 'reserved' };
  return { b: fake, c: { customer_id: '', name: 'Neha Sharma' }, items: { G021: { name: 'Parrot Green Mirror-work Ghagra' }, C043: { name: 'Parrot Green Choli' }, J027: { name: 'Oxidised Silver Set' } },
    lines: [{ item_code: 'G021', allocated_price: 499 }, { item_code: 'C043', allocated_price: 349 }, { item_code: 'J027', allocated_price: 149 }], pays: [{ kind: 'rent', amount: 997, mode: 'UPI', at: now }, { kind: 'deposit', amount: 2000, mode: 'UPI', at: now }],
    ret: { refund: 1700, refund_mode: 'UPI', late_fee: 0 }, retItems: [{ item_code: 'C043', condition: 'minor', charge: 300 }], T: T };
}
function previewCtx(T, req, key) {
  var ctx = req.booking_no ? emailCtx(T, req.booking_no) : sampleCtx(T);
  if (key === 'sale_receipt') { var s = (T.Sales || []).slice(-1)[0]; ctx = s ? emailCtx(T, null, s.sale_id) : ctx; if (!ctx.sale) ctx.sale = { sale_id: 'S-501', item_code: 'G021', list_price: 2499, discount: 200, price: 2299, mode: 'UPI', at: new Date() }; }
  if (key === 'staff_digest') ctx.staff = staffSummary(T, new Date());
  return ctx;
}
function emailPreview(req) {
  var T = loadAll(), em = emailSettings(T); if (req.email) for (var k in req.email) em[k] = req.email[k];
  var key = req.template; if (!TEMPLATES[key]) fail('Unknown template');
  var m = renderEmail(key, previewCtx(T, req, key), em, { prev: { pickup: new Date(Date.now() - 864e5), rent: 899 } }, req.override || null);
  return { ok: true, subject: m.subject, html: m.html };
}
function emailTest(req) {
  if (!validEmail(req.to)) fail('Enter a valid email address');
  var T = loadAll(), em = emailSettings(T); if (req.email) for (var k in req.email) em[k] = req.email[k];
  var key = req.template || 'booking_confirmation'; if (!TEMPLATES[key]) fail('Unknown template');
  var ctx = previewCtx(T, req, key), m = renderEmail(key, ctx, em, { prev: { pickup: new Date(Date.now() - 864e5), rent: 899 } }, req.override || null);
  var opts = { htmlBody: m.html, name: em.senderName || 'OMG Rental' }; if (validEmail(em.replyTo)) opts.replyTo = em.replyTo;
  if (m.pdf && (ctx.sale || (ctx.b && ctx.b.booking_no !== 'OMG-00124'))) { try { opts.attachments = [receiptPdf(ctx, em, key)]; } catch (x) { } }
  deliverMail(String(req.to).trim(), '[TEST] ' + m.subject, m.text, opts, em);
  return { ok: true };
}
/* staff (manager / admin) can resend an email for a booking, e.g. the receipt */
function emailSendNow(user, req) {
  if (user.role === 'staff') fail('Ask a manager to resend emails', 'forbidden');
  var key = req.template; if (!TEMPLATES[key] || TEMPLATES[key].kind === 'staff') fail('Unknown template');
  var T = loadAll(), em = emailSettings(T), row = newLogRow(key, key + '|' + (req.booking_no || req.sale_id) + '|manual|' + Date.now(), { booking_no: req.booking_no || '', sale_id: req.sale_id || '', customer_id: req.customer_id || '' });
  var res; try { res = sendOne(row, T, em); } catch (x) { res = { status: 'failed', error: String(x && x.message || x).slice(0, 300) }; }
  row.status = res.status; row.to = res.to || ''; row.subject = res.subject || ''; row.error = res.error || ''; row.attempts = 1; row.sent_at = res.status === 'sent' ? new Date().toISOString() : '';
  row.meta = JSON.stringify({ manual: user.name });
  var l = LockService.getScriptLock(); l.waitLock(25000); try { upsertObjects('Email_Log', [row]); } finally { l.releaseLock(); }
  return res.status === 'sent' ? { ok: true, to: res.to } : { ok: false, error: res.error };
}
function unsubscribePage(cid, t) {
  var ok = cid && t && t === unsubToken(cid), msg;
  if (ok) {
    var l = LockService.getScriptLock(); l.waitLock(20000);
    try { upsertObjects('Customers', [{ customer_id: cid, email_opt: 'no', marketing_opt: 'no' }]); } finally { l.releaseLock(); }
    msg = 'You will no longer receive emails from OMG Rental. To start again, just ask us at the shop or call 7610544284.';
  } else msg = 'This link is not valid. Please call OMG Rental on 7610544284.';
  return HtmlService.createHtmlOutput('<div style="font-family:Arial,sans-serif;max-width:480px;margin:60px auto;padding:24px;text-align:center;border:1px solid #eee;border-radius:16px"><h2 style="color:#7a1f4b">OMG Rental</h2><p>' + escH(msg) + '</p></div>').setTitle('OMG Rental');
}

/* ---------- ONE-TIME: run this in the editor to switch email on ---------- */
function setupEmail() {
  setup();   // makes sure the Email_Log tab exists
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'emailTick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('emailTick').timeBased().everyMinutes(15).create();
  var me = Session.getEffectiveUser().getEmail(), T = loadAll(true), em = emailSettings(T);
  var m = renderEmail('booking_confirmation', sampleCtx(T), em, {});
  GmailApp.sendEmail(me, '[TEST] ' + m.subject, m.text, { htmlBody: m.html, name: em.senderName });
  Logger.log('✅ Email is on. Reminders run every 15 minutes. A test email was sent to ' + me + '. Gmail quota left today: ' + MailApp.getRemainingDailyQuota() +
    '. "Send mail as" addresses: ' + (GmailApp.getAliases().join(', ') || 'none') + '. Now: Deploy → Manage deployments → ✏️ → New version → Deploy.');
}
