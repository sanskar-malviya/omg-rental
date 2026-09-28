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
  Customers:      { key: 'customer_id', cols: ['customer_id', 'name', 'mobile', 'whatsapp', 'email', 'address', 'customer_since', 'omg_salon_client', 'notes', 'id_type', 'id_last4'] },
  Bookings:       { key: 'booking_no', cols: ['booking_no', 'customer_id', 'customer_name', 'pickup_at', 'event_date', 'return_by', 'rent_gross', 'extra_days', 'discount', 'voucher_code', 'deposit_due', 'status', 'picked_at', 'received_at', 'closed_at', 'cancel_reason', 'late_flag_amount', 'late_flag_at', 'notes', 'created_by', 'created_at'] },
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
  Item_Photos:    { key: 'photo_id', cols: ['photo_id', 'item_code', 'customer_id', 'booking_no', 'kind', 'url', 'path', 'size_kb', 'uploaded_at', 'uploaded_by', 'active'] }
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
function doGet() { return json({ ok: true, app: 'OMG Rental API', version: 1 }); }

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
    if (changes.Users) changes.Users.forEach(function (r) { delete r.pin_hash; });
    checkNewBookings(changes, newBookings);
    Object.keys(changes).forEach(function (t) { if (changes[t].length) upsertObjects(t, changes[t]); });
    return { ok: true, serverTime: new Date().toISOString() };
  } finally { lock.releaseLock(); }
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
