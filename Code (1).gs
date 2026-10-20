/** EVENT APP — BACKEND API v1.0 (Google Apps Script). Hanya API JSON; tidak ada halaman UI.
 *  Jalankan setup() sekali, lalu buatSuperAdmin() sekali. Lihat panduan di jawaban Tahap 3. */
const VERSION = '1.0.0';
const COLS = {
  Settings: ['key', 'value', 'updated_at', 'updated_by'],
  Users: ['user_id', 'username', 'display_name', 'role', 'password_hash', 'is_active', 'must_change_password', 'failed_attempts', 'locked_until', 'last_login_at', 'created_at', 'updated_at', 'deleted_at'],
  Sessions: ['session_id', 'token_hash', 'user_id', 'created_at', 'expires_at', 'revoked_at', 'last_seen_at'],
  StaffAccess: ['user_id', 'event_id', 'granted_at', 'granted_by'],
  Events: ['event_id', 'event_name', 'title', 'tagline', 'description', 'organizer', 'contact_info', 'event_date', 'start_time', 'end_time', 'timezone', 'location', 'location_details', 'capacity', 'extra_info', 'status', 'registration_start', 'registration_end', 'created_at', 'updated_at', 'created_by', 'version', 'deleted_at'],
  EventDesigns: ['event_id', 'template_id', 'draft_config_json', 'published_config_json', 'draft_version', 'published_version', 'published_at', 'published_by', 'updated_at'],
  Registrations: ['registration_id', 'event_id', 'agency_name', 'brand_name', 'representative_name', 'participant_count', 'qr_token_hash', 'qr_version', 'registration_status', 'checkin_status', 'checkin_at', 'checked_in_by', 'request_id', 'source', 'registered_at', 'updated_at', 'version', 'deleted_at'],
  CheckIns: ['checkin_id', 'registration_id', 'event_id', 'participant_count', 'checked_in_at', 'checked_in_by', 'qr_version'],
  Assets: ['asset_id', 'event_id', 'drive_file_id', 'file_name', 'mime_type', 'size_bytes', 'url', 'uploaded_by', 'uploaded_at', 'deleted_at'],
  AuditLogs: ['log_id', 'at', 'user_id', 'action', 'entity_type', 'entity_id', 'summary']
};
const DEF_SET = { company_name: 'Perusahaan Anda', app_name: 'Event Manager', logo_asset_id: '', organizer_info: '', default_timezone: 'Asia/Jakarta', default_template_id: 'elegant-corporate', session_hours: '12', max_participants_per_registration: '100', register_rate_limit: '60' };
const EV_STATUS = ['draft', 'active', 'registration_closed', 'finished', 'cancelled'];
const TZS = ['Asia/Jakarta', 'Asia/Makassar', 'Asia/Jayapura'];
const FONTS = ['Inter', 'Poppins', 'Montserrat', 'Playfair Display', 'Lora', 'Merriweather', 'Raleway', 'Nunito', 'Cormorant Garamond', 'DM Serif Display'];
const LAYOUTS = ['hero-full', 'split', 'card', 'minimal', 'dark-premium', 'banner-detail', 'mobile-first'];
const TPL = { // [nama, layout, bg1, bg2, judul, isi, tombol, teksTombol, kartu, fontJudul, fontIsi]
  'elegant-corporate': ['Elegant Corporate', 'banner-detail', '#0f2a4a', '#1d4e89', '#ffffff', '#e2e8f0', '#c9a227', '#111111', '#ffffff', 'Playfair Display', 'Inter'],
  'luxury-dark': ['Luxury Dark', 'dark-premium', '#0b0b0f', '#1c1c24', '#f5e6b3', '#d1d5db', '#d4af37', '#111111', '#16161d', 'Cormorant Garamond', 'Raleway'],
  'minimalist-white': ['Minimalist White', 'minimal', '#ffffff', '#f3f4f6', '#111827', '#374151', '#111827', '#ffffff', '#ffffff', 'Inter', 'Inter'],
  'modern-gradient': ['Modern Gradient', 'hero-full', '#6d28d9', '#ec4899', '#ffffff', '#fdf2f8', '#ffffff', '#6d28d9', '#ffffff', 'Poppins', 'Poppins'],
  'formal-gala': ['Formal Gala', 'card', '#1a1a2e', '#4a1942', '#fff7e6', '#f3e8d0', '#b8860b', '#ffffff', '#fffaf0', 'DM Serif Display', 'Lora'],
  'nature-outdoor': ['Nature & Outdoor', 'split', '#14532d', '#65a30d', '#ffffff', '#ecfccb', '#facc15', '#14532d', '#f7fee7', 'Merriweather', 'Nunito'],
  'festive-gathering': ['Festive Gathering', 'mobile-first', '#be123c', '#f97316', '#ffffff', '#fff1f2', '#fde047', '#7f1d1d', '#ffffff', 'Montserrat', 'Nunito']
};

/* ===================== Helper dasar ===================== */
const props_ = () => PropertiesService.getScriptProperties();
function ss_() { const id = props_().getProperty('SPREADSHEET_ID'); return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet(); }
function sh_(n) { const s = ss_().getSheetByName(n); if (!s) throw err_('NOT_SETUP', 'Database belum disiapkan. Jalankan setup() di editor Apps Script.'); return s; }
function err_(c, m) { const e = new Error(m); e.code = c; return e; }
const now_ = () => new Date().toISOString();
const id_ = (p, n) => p + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, n || 10).toUpperCase();
const sha_ = s => Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8).map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
function pwhash_(pw, salt) { let h = salt + pw; for (let i = 0; i < 2000; i++) h = sha_(h + salt + pw); return h; }
const newToken_ = () => (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '').slice(0, 40);
const s_ = (v, m) => String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, m);
const l_ = (v, m) => s_(v, m).replace(/\s+/g, ' ');
const out_ = o => { const x = Object.assign({}, o); delete x._row; return x; };
const hex_ = (v, d) => /^#[0-9a-fA-F]{6}$/.test(v) ? v : d;
const num_ = (v, lo, hi, d) => { v = Number(v); return isFinite(v) ? Math.min(hi, Math.max(lo, v)) : d; };
const pick_ = (v, list, d) => list.indexOf(v) >= 0 ? v : d;

function all_(n) {
  const v = sh_(n).getDataRange().getValues(), h = v.shift(), o = [];
  v.forEach((r, i) => { if (r[0] === '') return; const x = { _row: i + 2 }; h.forEach((k, j) => x[k] = r[j] == null ? '' : String(r[j])); o.push(x); });
  return o;
}
function save_(n, o) {
  const c = COLS[n], s = sh_(n), row = c.map(k => o[k] == null ? '' : o[k]);
  if (o._row) s.getRange(o._row, 1, 1, c.length).setValues([row]); else { s.appendRow(row); o._row = s.getLastRow(); }
  return o;
}
function locked_(fn) {
  const l = LockService.getScriptLock();
  if (!l.tryLock(25000)) throw err_('BUSY', 'Server sedang sibuk, coba lagi sebentar.');
  try { return fn(); } finally { l.releaseLock(); }
}
function rate_(k, max) { const c = CacheService.getScriptCache(), n = Number(c.get(k) || 0); if (n >= max) throw err_('RATE_LIMIT', 'Terlalu banyak permintaan. Coba lagi sebentar.'); c.put(k, String(n + 1), 60); }
function aud_(u, action, type, id, summary) { try { save_('AuditLogs', { log_id: id_('LOG', 12), at: now_(), user_id: u ? u.user_id : '', action, entity_type: type, entity_id: id || '', summary: s_(summary, 300) }); } catch (e) { console.error(e); } }
function setting_(k) { const r = all_('Settings').find(x => x.key === k); return r ? r.value : DEF_SET[k]; }

/* ===================== Setup (dijalankan manual sekali) ===================== */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Buka Apps Script dari Spreadsheet: menu Extensions > Apps Script.');
  props_().setProperty('SPREADSHEET_ID', ss.getId());
  Object.keys(COLS).forEach(n => {
    const c = COLS[n], R = ['Registrations', 'CheckIns', 'AuditLogs', 'Sessions'].indexOf(n) >= 0 ? 5000 : 1000;
    const s = ss.getSheetByName(n) || ss.insertSheet(n);
    if (s.getMaxColumns() < c.length) s.insertColumnsAfter(s.getMaxColumns(), c.length - s.getMaxColumns());
    if (s.getMaxRows() < R) s.insertRowsAfter(s.getMaxRows(), R - s.getMaxRows());
    s.getRange(1, 1, s.getMaxRows(), c.length).setNumberFormat('@');
    s.getRange(1, 1, 1, c.length).setValues([c]).setFontWeight('bold').setBackground('#e8eefc');
    s.setFrozenRows(1);
  });
  const d = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1') || ss.getSheetByName('Lembar1');
  if (d && ss.getSheets().length > 1 && d.getLastRow() === 0) ss.deleteSheet(d);
  Logger.log('Setup selesai. Lanjut: isi data di buatSuperAdmin() lalu jalankan.');
}
/** Isi USERNAME & PASSWORD_SEMENTARA, jalankan sekali, lalu kembalikan nilainya. Password wajib diganti saat login pertama. */
function buatSuperAdmin() {
  const USERNAME = 'admin', PASSWORD_SEMENTARA = 'GANTI_PASSWORD_SEMENTARA';
  if (PASSWORD_SEMENTARA.indexOf('GANTI') === 0 || PASSWORD_SEMENTARA.length < 8) throw new Error('Isi PASSWORD_SEMENTARA (minimal 8 karakter) dulu.');
  const un = USERNAME.toLowerCase(), salt = id_('S', 16), ph = salt + '$' + pwhash_(PASSWORD_SEMENTARA, salt), t = now_();
  let u = all_('Users').find(x => x.username.toLowerCase() === un && !x.deleted_at);
  if (u) { u.password_hash = ph; u.is_active = 'TRUE'; u.role = 'super_admin'; u.must_change_password = 'TRUE'; u.failed_attempts = '0'; u.locked_until = ''; u.updated_at = t; }
  else u = { user_id: id_('USR'), username: un, display_name: 'Super Admin', role: 'super_admin', password_hash: ph, is_active: 'TRUE', must_change_password: 'TRUE', failed_attempts: '0', created_at: t, updated_at: t };
  save_('Users', u);
  Logger.log('Super admin "' + un + '" siap. Sekarang hapus password dari kode ini.');
}

/* ===================== Router ===================== */
// Level akses: 0 = publik, 1 = pengguna login (admin/petugas), 2 = hanya super_admin
const ACT = {
  'system.ping': [0, ping_],
  'auth.login': [0, login_], 'auth.logout': [1, logout_], 'auth.me': [1, (d, u) => pu_(u)], 'auth.changePassword': [1, changePw_],
  'events.list': [2, evList_], 'events.get': [2, evGet_], 'events.create': [2, evCreate_], 'events.update': [2, evUpdate_],
  'events.duplicate': [2, evDup_], 'events.setStatus': [2, evStatus_], 'events.archive': [2, evArchive_],
  'design.get': [2, dsGet_], 'design.saveDraft': [2, dsSave_], 'design.publish': [2, dsPublish_], 'design.reset': [2, dsReset_],
  'assets.upload': [2, asUpload_], 'assets.delete': [2, asDelete_],
  'participants.list': [2, pList_], 'participants.get': [2, pGet_], 'participants.create': [2, pCreate_], 'participants.update': [2, pUpdate_],
  'participants.archive': [2, pArchive_], 'participants.resetQr': [2, pResetQr_], 'participants.export': [2, pExport_],
  'public.event': [0, pubEvent_], 'public.register': [0, pubRegister_],
  'checkin.events': [1, ciEvents_], 'checkin.scan': [1, ciScan_],
  'users.list': [2, () => usersList_()], 'users.create': [2, uCreate_], 'users.update': [2, uUpdate_], 'users.resetPassword': [2, uResetPw_], 'users.setAccess': [2, uAccess_],
  'settings.get': [2, setGet_], 'settings.save': [2, setSave_], 'dashboard.stats': [2, dash_], 'audit.list': [2, auditList_]
};
function route_(req) {
  const a = ACT[req.action]; if (!a) throw err_('NOT_FOUND', 'Action tidak dikenal.');
  const d = req.data && typeof req.data === 'object' ? req.data : {};
  let u = null;
  if (a[0] > 0) {
    u = auth_(req.token);
    if (u.must_change_password === 'TRUE' && req.action.indexOf('auth.') !== 0) throw err_('MUST_CHANGE_PASSWORD', 'Anda harus mengganti password terlebih dahulu.');
    if (a[0] === 2 && u.role !== 'super_admin') throw err_('FORBIDDEN', 'Anda tidak memiliki izin untuk tindakan ini.');
  }
  return a[1](d, u, req);
}
function respond_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return respond_({ ok: true, data: { name: 'Event API', version: VERSION, note: 'Gunakan POST.' } }); }
function doPost(e) {
  try {
    let req; try { req = JSON.parse((e.postData && e.postData.contents) || '{}'); } catch (x) { throw err_('BAD_REQUEST', 'Format permintaan bukan JSON.'); }
    return respond_({ ok: true, data: route_(req) });
  } catch (x) {
    if (!x.code) console.error(x.stack || x);
    return respond_({ ok: false, error: { code: x.code || 'INTERNAL', message: x.code ? x.message : 'Terjadi kesalahan pada server.' } });
  }
}
function ping_() { let ready = true; try { Object.keys(COLS).forEach(n => sh_(n)); } catch (e) { ready = false; } return { name: 'Event API', version: VERSION, time: now_(), database_ready: ready }; }

/* ===================== Autentikasi ===================== */
const pu_ = u => ({ user_id: u.user_id, username: u.username, display_name: u.display_name, role: u.role, is_active: u.is_active === 'TRUE', must_change_password: u.must_change_password === 'TRUE', last_login_at: u.last_login_at, created_at: u.created_at });
function auth_(token) {
  if (!token) throw err_('UNAUTHORIZED', 'Silakan login terlebih dahulu.');
  const th = sha_(String(token)), t = now_();
  const s = all_('Sessions').find(x => x.token_hash === th && !x.revoked_at && x.expires_at > t);
  if (!s) throw err_('UNAUTHORIZED', 'Sesi berakhir, silakan login ulang.');
  const u = all_('Users').find(x => x.user_id === s.user_id && x.is_active === 'TRUE' && !x.deleted_at);
  if (!u) throw err_('UNAUTHORIZED', 'Akun tidak aktif.');
  return u;
}
function revoke_(uid, exceptHash) { all_('Sessions').forEach(s => { if (s.user_id === uid && !s.revoked_at && s.token_hash !== exceptHash) { s.revoked_at = now_(); save_('Sessions', s); } }); }
function login_(d) {
  const un = String(d.username || '').trim().toLowerCase(), pw = String(d.password || '');
  if (!un || !pw) throw err_('INVALID', 'Username dan password wajib diisi.');
  const u = all_('Users').find(x => x.username.toLowerCase() === un && !x.deleted_at);
  if (u && u.locked_until > now_()) throw err_('LOCKED', 'Akun dikunci sementara karena terlalu banyak percobaan. Coba lagi 10 menit lagi.');
  const p = (u ? u.password_hash : 'x$x').split('$'), ok = pwhash_(pw, p[0]) === p[1] && u && u.is_active === 'TRUE';
  if (!ok) {
    if (u) { const f = Number(u.failed_attempts || 0) + 1; u.failed_attempts = f >= 5 ? '0' : String(f); if (f >= 5) u.locked_until = new Date(Date.now() + 600000).toISOString(); save_('Users', u); }
    throw err_('INVALID_LOGIN', 'Username atau password salah.');
  }
  const token = newToken_() + newToken_(), t = now_(), hrs = Number(setting_('session_hours')) || 12;
  save_('Sessions', { session_id: id_('SES'), token_hash: sha_(token), user_id: u.user_id, created_at: t, expires_at: new Date(Date.now() + hrs * 3600000).toISOString(), last_seen_at: t });
  u.failed_attempts = '0'; u.locked_until = ''; u.last_login_at = t; save_('Users', u);
  if (Math.random() < 0.1) purge_();
  aud_(u, 'auth.login', 'user', u.user_id, 'Login');
  return { token, expires_in_hours: hrs, user: pu_(u) };
}
function purge_() { const lim = new Date(Date.now() - 2 * 86400000).toISOString(); all_('Sessions').filter(s => s.expires_at < lim).sort((a, b) => b._row - a._row).forEach(s => sh_('Sessions').deleteRow(s._row)); }
function logout_(d, u, req) { const th = sha_(String(req.token)), s = all_('Sessions').find(x => x.token_hash === th); if (s) { s.revoked_at = now_(); save_('Sessions', s); } return { logged_out: true }; }
function changePw_(d, u, req) {
  const np = String(d.new_password || ''), p = u.password_hash.split('$');
  if (pwhash_(String(d.old_password || ''), p[0]) !== p[1]) throw err_('INVALID', 'Password lama salah.');
  if (np.length < 8) throw err_('INVALID', 'Password baru minimal 8 karakter.');
  if (np === String(d.old_password)) throw err_('INVALID', 'Password baru harus berbeda.');
  const salt = id_('S', 16); u.password_hash = salt + '$' + pwhash_(np, salt); u.must_change_password = 'FALSE'; u.updated_at = now_(); save_('Users', u);
  revoke_(u.user_id, sha_(String(req.token))); aud_(u, 'auth.changePassword', 'user', u.user_id, 'Ganti password');
  return pu_(u);
}

/* ===================== Event ===================== */
const ev_ = (id, incl) => { const e = all_('Events').find(x => x.event_id === String(id)); if (!e || (e.deleted_at && !incl)) throw err_('NOT_FOUND', 'Event tidak ditemukan.'); return e; };
const EMPTY_ST = () => ({ registrations: 0, people: 0, checked_registrations: 0, checked_people: 0 });
function statsMap_() {
  const m = {};
  all_('Registrations').forEach(r => {
    if (r.deleted_at || r.registration_status !== 'registered') return;
    const s = m[r.event_id] || (m[r.event_id] = EMPTY_ST()), n = Number(r.participant_count) || 0;
    s.registrations++; s.people += n; if (r.checkin_status === 'checked_in') { s.checked_registrations++; s.checked_people += n; }
  });
  return m;
}
const st_ = (m, id) => m[id] || EMPTY_ST();
const evOut_ = (e, m) => { const x = out_(e); x.capacity = Number(x.capacity) || 0; x.version = Number(x.version) || 1; x.stats = st_(m, e.event_id); return x; };
function isoOpt_(v, label) { v = String(v || '').trim(); if (!v) return ''; const t = Date.parse(v); if (isNaN(t)) throw err_('INVALID', label + ' tidak valid.'); return new Date(t).toISOString(); }
function cleanEvent_(d, b) {
  const o = Object.assign({}, b), g = (k, m, ml) => d[k] !== undefined ? (ml ? s_(d[k], m) : l_(d[k], m)) : (o[k] || '');
  o.event_name = g('event_name', 120); if (!o.event_name) throw err_('INVALID', 'Nama event wajib diisi.');
  o.title = g('title', 150) || o.event_name; o.tagline = g('tagline', 200); o.description = g('description', 3000, 1); o.organizer = g('organizer', 120);
  o.contact_info = g('contact_info', 300, 1); o.location = g('location', 200); o.location_details = g('location_details', 1000, 1); o.extra_info = g('extra_info', 2000, 1);
  o.event_date = g('event_date', 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(o.event_date) || isNaN(Date.parse(o.event_date))) throw err_('INVALID', 'Tanggal event wajib diisi (format YYYY-MM-DD).');
  ['start_time', 'end_time'].forEach(k => { o[k] = g(k, 5); if (o[k] && !/^([01]\d|2[0-3]):[0-5]\d$/.test(o[k])) throw err_('INVALID', 'Format jam harus HH:mm.'); });
  o.timezone = pick_(g('timezone', 30), TZS, setting_('default_timezone'));
  if (d.capacity !== undefined) { const c = parseInt(d.capacity, 10); if (!(c >= 0 && c <= 100000)) throw err_('INVALID', 'Kapasitas harus 0–100000 (0 = tanpa batas).'); o.capacity = c; } else o.capacity = Number(o.capacity) || 0;
  if (d.status !== undefined) { if (EV_STATUS.indexOf(d.status) < 0) throw err_('INVALID', 'Status tidak valid.'); o.status = d.status; } else o.status = o.status || 'draft';
  if (d.registration_start !== undefined) o.registration_start = isoOpt_(d.registration_start, 'Awal pendaftaran');
  if (d.registration_end !== undefined) o.registration_end = isoOpt_(d.registration_end, 'Akhir pendaftaran');
  if (o.registration_start && o.registration_end && o.registration_end < o.registration_start) throw err_('INVALID', 'Akhir pendaftaran harus setelah awalnya.');
  return o;
}
function evList_(d) { const m = statsMap_(); return all_('Events').filter(e => d.include_archived || !e.deleted_at).map(e => evOut_(e, m)).reverse(); }
function evGet_(d) { return evOut_(ev_(d.event_id, true), statsMap_()); }
function evCreate_(d, u) {
  const t = now_(), o = cleanEvent_(d, { status: 'draft' }), tid = TPL[d.template_id] ? d.template_id : setting_('default_template_id');
  Object.assign(o, { event_id: id_('EVT'), created_at: t, updated_at: t, created_by: u.user_id, version: 1 }); delete o._row; o.deleted_at = '';
  save_('Events', o);
  save_('EventDesigns', { event_id: o.event_id, template_id: tid, draft_config_json: JSON.stringify(defCfg_(tid)), published_config_json: '', draft_version: 1, published_version: 0, updated_at: t });
  aud_(u, 'events.create', 'event', o.event_id, 'Buat event: ' + o.event_name);
  return evOut_(o, {});
}
function evUpdate_(d, u) {
  return locked_(() => {
    const e = ev_(d.event_id); if (d.version !== undefined && String(d.version) !== e.version) throw err_('CONFLICT', 'Event sudah diubah orang lain. Muat ulang data lalu coba lagi.');
    const o = cleanEvent_(d, e); o.version = Number(e.version) + 1; o.updated_at = now_(); save_('Events', o);
    aud_(u, 'events.update', 'event', o.event_id, 'Ubah event: ' + o.event_name); return evOut_(o, statsMap_());
  });
}
function evDup_(d, u) {
  const e = ev_(d.event_id), t = now_(), n = Object.assign({}, e, { event_id: id_('EVT'), event_name: s_(e.event_name + ' (Salinan)', 120), status: 'draft', created_at: t, updated_at: t, created_by: u.user_id, version: 1, deleted_at: '' }); delete n._row;
  save_('Events', n); const ds = dsRow_(e.event_id);
  save_('EventDesigns', { event_id: n.event_id, template_id: ds.template_id, draft_config_json: ds.draft_config_json, published_config_json: '', draft_version: 1, published_version: 0, updated_at: t });
  aud_(u, 'events.duplicate', 'event', n.event_id, 'Duplikasi dari ' + e.event_id); return evOut_(n, {});
}
function evStatus_(d, u) {
  if (EV_STATUS.indexOf(d.status) < 0) throw err_('INVALID', 'Status tidak valid.');
  return locked_(() => { const e = ev_(d.event_id); e.status = d.status; e.version = Number(e.version) + 1; e.updated_at = now_(); save_('Events', e); aud_(u, 'events.setStatus', 'event', e.event_id, 'Status: ' + d.status); return evOut_(e, statsMap_()); });
}
function evArchive_(d, u) {
  return locked_(() => {
    const e = ev_(d.event_id, true), has = all_('Registrations').some(r => r.event_id === e.event_id);
    if (d.permanent) {
      if (has) throw err_('HAS_DATA', 'Event sudah memiliki data pendaftaran, hanya bisa diarsipkan.');
      [['StaffAccess', 'event_id'], ['EventDesigns', 'event_id']].forEach(p => all_(p[0]).filter(x => x[p[1]] === e.event_id).sort((a, b) => b._row - a._row).forEach(x => sh_(p[0]).deleteRow(x._row)));
      sh_('Events').deleteRow(e._row); aud_(u, 'events.delete', 'event', e.event_id, 'Hapus permanen: ' + e.event_name); return { deleted: true };
    }
    e.deleted_at = now_(); e.updated_at = e.deleted_at; save_('Events', e); aud_(u, 'events.archive', 'event', e.event_id, 'Arsip: ' + e.event_name); return { archived: true };
  });
}

/* ===================== Desain undangan ===================== */
function defCfg_(tid) {
  if (!TPL[tid]) tid = 'elegant-corporate'; const p = TPL[tid];
  return { schema: 1, template_id: tid, layout: p[1],
    theme: { bg: { type: 'gradient', color: p[2], gradient: { angle: 135, stops: [p[2], p[3]] }, asset_id: '', position: 'center', fit: 'cover' }, overlay: { color: '#000000', opacity: 0.35 },
      colors: { title: p[4], body: p[5], muted: p[5], button: p[6], buttonText: p[7], card: p[8], border: '#e5e7eb' },
      font: { heading: p[9], body: p[10], titleSize: 44, subtitleSize: 20, bodySize: 16, weight: 700 }, style: { align: 'center', spacing: 'normal', radius: 12, shadow: 'soft' } },
    content: { welcome: '', button_text: 'Daftar Sekarang', closing: '', footer: '' },
    show: { logo: true, banner: true, description: true, date: true, location: true, contact: true, extra: true, footer: true }, assets: { logo: '', banner: '' } };
}
function cleanCfg_(c, tid) {
  c = c && typeof c === 'object' ? c : {};
  const D = defCfg_(tid), t = c.theme || {}, dt = D.theme, bg = t.bg || {}, ov = t.overlay || {}, co = t.colors || {}, fo = t.font || {}, sy = t.style || {}, ct = c.content || {}, sh = c.show || {}, as = c.assets || {};
  const A = all_('Assets').filter(a => !a.deleted_at).map(a => a.asset_id), ast = v => A.indexOf(String(v || '')) >= 0 ? String(v) : '';
  const gs = bg.gradient && Array.isArray(bg.gradient.stops) && bg.gradient.stops.length >= 2 ? bg.gradient.stops.slice(0, 4).map(x => hex_(x, '#000000')) : dt.bg.gradient.stops;
  const colors = {}; Object.keys(dt.colors).forEach(k => colors[k] = hex_(co[k], dt.colors[k]));
  const show = {}; Object.keys(D.show).forEach(k => show[k] = sh[k] === undefined ? D.show[k] : !!sh[k]);
  return { schema: 1, template_id: D.template_id, layout: pick_(c.layout, LAYOUTS, D.layout),
    theme: { bg: { type: pick_(bg.type, ['solid', 'gradient', 'image'], dt.bg.type), color: hex_(bg.color, dt.bg.color), gradient: { angle: num_(bg.gradient && bg.gradient.angle, 0, 360, 135), stops: gs },
        asset_id: ast(bg.asset_id), position: pick_(bg.position, ['center', 'top', 'bottom', 'left', 'right'], 'center'), fit: pick_(bg.fit, ['cover', 'contain'], 'cover') },
      overlay: { color: hex_(ov.color, '#000000'), opacity: num_(ov.opacity, 0, 1, dt.overlay.opacity) }, colors,
      font: { heading: pick_(fo.heading, FONTS, dt.font.heading), body: pick_(fo.body, FONTS, dt.font.body), titleSize: num_(fo.titleSize, 20, 72, 44), subtitleSize: num_(fo.subtitleSize, 12, 36, 20), bodySize: num_(fo.bodySize, 12, 24, 16), weight: pick_(Number(fo.weight), [400, 500, 600, 700, 800], 700) },
      style: { align: pick_(sy.align, ['left', 'center', 'right'], 'center'), spacing: pick_(sy.spacing, ['compact', 'normal', 'relaxed'], 'normal'), radius: num_(sy.radius, 0, 40, 12), shadow: pick_(sy.shadow, ['none', 'soft', 'strong'], 'soft') } },
    content: { welcome: s_(ct.welcome, 1000), button_text: l_(ct.button_text || D.content.button_text, 40), closing: s_(ct.closing, 500), footer: l_(ct.footer, 300) }, show, assets: { logo: ast(as.logo), banner: ast(as.banner) } };
}
function dsRow_(id) {
  let r = all_('EventDesigns').find(x => x.event_id === id);
  if (!r) r = save_('EventDesigns', { event_id: id, template_id: setting_('default_template_id'), draft_config_json: JSON.stringify(defCfg_(setting_('default_template_id'))), published_config_json: '', draft_version: 1, published_version: 0, updated_at: now_() });
  return r;
}
function assetUrls_(cfgs) {
  const ids = {}; cfgs.forEach(c => { if (!c) return; [c.theme && c.theme.bg && c.theme.bg.asset_id, c.assets && c.assets.logo, c.assets && c.assets.banner].forEach(i => { if (i) ids[i] = 1; }); });
  const m = {}; all_('Assets').forEach(a => { if (ids[a.asset_id] && !a.deleted_at) m[a.asset_id] = a.url; }); return m;
}
const parse_ = s => { try { return s ? JSON.parse(s) : null; } catch (e) { return null; } };
function dsOut_(r) {
  const dr = parse_(r.draft_config_json), pb = parse_(r.published_config_json);
  return { event_id: r.event_id, template_id: r.template_id, draft: dr, published: pb, draft_version: Number(r.draft_version), published_version: Number(r.published_version), published_at: r.published_at, has_unpublished_changes: Number(r.draft_version) !== Number(r.published_version), asset_urls: assetUrls_([dr, pb]), templates: Object.keys(TPL).map(k => ({ id: k, name: TPL[k][0], layout: TPL[k][1] })) };
}
function dsGet_(d) { ev_(d.event_id, true); return dsOut_(dsRow_(d.event_id)); }
function dsSave_(d, u) {
  return locked_(() => {
    ev_(d.event_id); const r = dsRow_(d.event_id), tid = TPL[d.template_id] ? d.template_id : r.template_id;
    const cfg = cleanCfg_(d.config, tid); const js = JSON.stringify(cfg); if (js.length > 40000) throw err_('INVALID', 'Konfigurasi desain terlalu besar.');
    r.template_id = tid; r.draft_config_json = js; r.draft_version = Number(r.draft_version) + 1; r.updated_at = now_(); save_('EventDesigns', r);
    aud_(u, 'design.saveDraft', 'event', d.event_id, 'Simpan draft desain'); return dsOut_(r);
  });
}
function dsPublish_(d, u) {
  return locked_(() => {
    ev_(d.event_id); const r = dsRow_(d.event_id);
    r.published_config_json = r.draft_config_json; r.published_version = r.draft_version; r.published_at = now_(); r.published_by = u.user_id; save_('EventDesigns', r);
    aud_(u, 'design.publish', 'event', d.event_id, 'Publikasi desain v' + r.published_version); return dsOut_(r);
  });
}
function dsReset_(d, u) {
  return locked_(() => {
    ev_(d.event_id); const r = dsRow_(d.event_id), tid = TPL[d.template_id] ? d.template_id : r.template_id;
    r.template_id = tid; r.draft_config_json = JSON.stringify(defCfg_(tid)); r.draft_version = Number(r.draft_version) + 1; r.updated_at = now_(); save_('EventDesigns', r);
    aud_(u, 'design.reset', 'event', d.event_id, 'Reset desain ke template ' + tid); return dsOut_(r);
  });
}

/* ===================== Aset (Google Drive) ===================== */
function asUpload_(d, u) {
  const mime = String(d.mime_type || ''), types = { 'image/jpeg': 'jpg', 'image/png': 'png', 'image/webp': 'webp' };
  if (!types[mime]) throw err_('INVALID', 'Format gambar harus JPG, PNG, atau WebP.');
  let bytes; try { bytes = Utilities.base64Decode(String(d.data_base64 || '')); } catch (e) { throw err_('INVALID', 'Data gambar tidak valid.'); }
  if (bytes.length < 100 || bytes.length > 600 * 1024) throw err_('INVALID', 'Ukuran gambar maksimal 600 KB (gambar dikompres otomatis oleh aplikasi).');
  const b = bytes, ok = mime === 'image/jpeg' ? (b[0] === -1 && b[1] === -40) : mime === 'image/png' ? (b[0] === -119 && b[1] === 80 && b[2] === 78 && b[3] === 71) : (b[0] === 82 && b[1] === 73 && b[2] === 70 && b[3] === 70 && b[8] === 87 && b[9] === 69 && b[10] === 66 && b[11] === 80);
  if (!ok) throw err_('INVALID', 'Isi file tidak cocok dengan format gambar.');
  if (d.event_id) ev_(d.event_id);
  let fid = props_().getProperty('ASSET_FOLDER_ID'), folder; try { folder = fid ? DriveApp.getFolderById(fid) : null; } catch (e) { folder = null; }
  if (!folder) { folder = DriveApp.createFolder('EventApp-Assets'); props_().setProperty('ASSET_FOLDER_ID', folder.getId()); }
  const aid = id_('AST'), name = l_(d.file_name, 80).replace(/[^\w.\- ]/g, '_') || ('gambar.' + types[mime]);
  const f = folder.createFile(Utilities.newBlob(bytes, mime, aid + '-' + name)); f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const a = { asset_id: aid, event_id: d.event_id || '', drive_file_id: f.getId(), file_name: name, mime_type: mime, size_bytes: bytes.length, url: 'https://lh3.googleusercontent.com/d/' + f.getId(), uploaded_by: u.user_id, uploaded_at: now_() };
  save_('Assets', a); aud_(u, 'assets.upload', 'asset', aid, name); return out_(a);
}
function asDelete_(d, u) {
  const a = all_('Assets').find(x => x.asset_id === d.asset_id && !x.deleted_at); if (!a) throw err_('NOT_FOUND', 'Aset tidak ditemukan.');
  try { DriveApp.getFileById(a.drive_file_id).setTrashed(true); } catch (e) { console.error(e); }
  a.deleted_at = now_(); save_('Assets', a); aud_(u, 'assets.delete', 'asset', a.asset_id, a.file_name); return { deleted: true };
}

/* ===================== Halaman publik ===================== */
function accept_(e, sm) {
  const t = now_();
  if (e.status !== 'active') return { ok: false, reason: e.status === 'registration_closed' ? 'Pendaftaran untuk event ini sudah ditutup.' : 'Event ini tidak menerima pendaftaran.' };
  if (e.registration_start && t < e.registration_start) return { ok: false, reason: 'Pendaftaran belum dibuka.' };
  if (e.registration_end && t > e.registration_end) return { ok: false, reason: 'Periode pendaftaran sudah berakhir.' };
  const cap = Number(e.capacity) || 0; if (cap && st_(sm, e.event_id).people >= cap) return { ok: false, reason: 'Kuota pendaftaran sudah penuh.' };
  return { ok: true };
}
function pubEvent_(d) {
  const e = all_('Events').find(x => x.event_id === String(d.event_id || '') && !x.deleted_at);
  if (!e || (e.status !== 'active' && e.status !== 'registration_closed')) throw err_('NOT_FOUND', 'Event tidak ditemukan.');
  const ds = dsRow_(e.event_id), cfg = parse_(ds.published_config_json) || defCfg_(ds.template_id), a = accept_(e, statsMap_()), x = {};
  ['event_id', 'title', 'tagline', 'description', 'organizer', 'contact_info', 'event_date', 'start_time', 'end_time', 'timezone', 'location', 'location_details', 'extra_info', 'status'].forEach(k => x[k] = e[k]);
  x.accepting = a.ok; x.closed_reason = a.ok ? '' : a.reason; x.config = cfg; x.asset_urls = assetUrls_([cfg]);
  x.app = { company_name: setting_('company_name'), app_name: setting_('app_name'), max_participants: Number(setting_('max_participants_per_registration')) || 100 };
  return x;
}
function ticket_(r, e, token) { return { registration_id: r.registration_id, qr_text: 'EVQ1.' + token, qr_version: Number(r.qr_version), event_id: e.event_id, event_title: e.title, agency_name: r.agency_name, brand_name: r.brand_name, representative_name: r.representative_name, participant_count: Number(r.participant_count) }; }
function regCreate_(d, u, manual) {
  const ag = l_(d.agency_name, 100), br = l_(d.brand_name, 100), rp = l_(d.representative_name, 100), n = Number(d.participant_count), max = Number(setting_('max_participants_per_registration')) || 100;
  if (!ag || !br || !rp) throw err_('INVALID', 'Nama agensi, brand, dan perwakilan wajib diisi.');
  if (!Number.isInteger(n) || n < 1 || n > max) throw err_('INVALID', 'Jumlah peserta harus bilangan bulat 1–' + max + '.');
  if (!manual) { if (d.website) throw err_('INVALID', 'Permintaan ditolak.'); rate_('reg_' + s_(d.event_id, 30), Number(setting_('register_rate_limit')) || 60); }
  const rid = s_(d.request_id, 64);
  return locked_(() => {
    const e = ev_(d.event_id), sm = statsMap_();
    if (!manual) { const a = accept_(e, sm); if (!a.ok) throw err_('CLOSED', a.reason); }
    if (!manual && rid.length >= 8) { const dup = all_('Registrations').find(r => r.event_id === e.event_id && r.request_id === rid && !r.deleted_at); if (dup) { const tk = newToken_(); dup.qr_token_hash = sha_(tk); dup.qr_version = Number(dup.qr_version) + 1; save_('Registrations', dup); return ticket_(dup, e, tk); } }
    const cap = Number(e.capacity) || 0; if (cap && st_(sm, e.event_id).people + n > cap) throw err_('CAPACITY', 'Kuota tidak mencukupi. Sisa kuota: ' + Math.max(0, cap - st_(sm, e.event_id).people) + ' orang.');
    const tk = newToken_(), t = now_(), r = { registration_id: id_('REG', 10), event_id: e.event_id, agency_name: ag, brand_name: br, representative_name: rp, participant_count: n, qr_token_hash: sha_(tk), qr_version: 1, registration_status: 'registered', checkin_status: 'not_checked_in', request_id: manual ? '' : rid, source: manual ? 'manual' : 'public', registered_at: t, updated_at: t, version: 1 };
    save_('Registrations', r); if (manual) aud_(u, 'participants.create', 'registration', r.registration_id, ag + ' / ' + br + ' (' + n + ')');
    return ticket_(r, e, tk);
  });
}
function pubRegister_(d) { return regCreate_(d, null, false); }

/* ===================== Check-in ===================== */
function canEvent_(u, eid) { if (u.role === 'super_admin') return true; return all_('StaffAccess').some(a => a.user_id === u.user_id && a.event_id === eid); }
function ciEvents_(d, u) {
  const sm = statsMap_();
  return all_('Events').filter(e => !e.deleted_at && (e.status === 'active' || e.status === 'registration_closed') && canEvent_(u, e.event_id))
    .map(e => ({ event_id: e.event_id, title: e.title, event_date: e.event_date, location: e.location, stats: st_(sm, e.event_id) }));
}
function ciScan_(d, u) {
  const eid = String(d.event_id || ''); ev_(eid); if (!canEvent_(u, eid)) throw err_('FORBIDDEN', 'Anda tidak memiliki akses ke event ini.');
  const m = /^EVQ1\.([0-9a-f]{40})$/.exec(String(d.qr || '').trim()); if (!m) return { result: 'INVALID', message: 'QR Code tidak dikenali.' };
  const h = sha_(m[1]);
  return locked_(() => {
    const r = all_('Registrations').find(x => x.qr_token_hash === h && !x.deleted_at);
    if (!r) return { result: 'INVALID', message: 'QR Code tidak dikenali.' };
    const ev = ev_(r.event_id, true), info = { event_title: ev.title, agency_name: r.agency_name, brand_name: r.brand_name, representative_name: r.representative_name, participant_count: Number(r.participant_count) };
    if (r.event_id !== eid) return { result: 'WRONG_EVENT', message: 'QR Code ini untuk event lain.', registration: info };
    if (r.registration_status !== 'registered') return { result: 'CANCELLED', message: 'Pendaftaran ini sudah dibatalkan.', registration: info };
    if (r.checkin_status === 'checked_in') return { result: 'DUPLICATE', message: 'QR Code sudah digunakan untuk check-in.', registration: info, checked_in_at: r.checkin_at };
    const t = now_(); r.checkin_status = 'checked_in'; r.checkin_at = t; r.checked_in_by = u.user_id; r.updated_at = t; save_('Registrations', r);
    save_('CheckIns', { checkin_id: id_('CHK', 12), registration_id: r.registration_id, event_id: eid, participant_count: r.participant_count, checked_in_at: t, checked_in_by: u.user_id, qr_version: r.qr_version });
    aud_(u, 'checkin.scan', 'registration', r.registration_id, r.agency_name + ' (' + r.participant_count + ' orang)');
    return { result: 'OK', message: 'Check-in berhasil.', registration: info, checked_in_at: t };
  });
}

/* ===================== Peserta (admin) ===================== */
const rOut_ = r => { const x = out_(r); delete x.qr_token_hash; delete x.request_id; x.participant_count = Number(x.participant_count); x.version = Number(x.version); x.qr_version = Number(x.qr_version); return x; };
const reg_ = id => { const r = all_('Registrations').find(x => x.registration_id === String(id)); if (!r) throw err_('NOT_FOUND', 'Pendaftaran tidak ditemukan.'); return r; };
function pList_(d) {
  const q = l_(d.q, 100).toLowerCase();
  return all_('Registrations').filter(r => (d.include_archived || !r.deleted_at) && (!d.event_id || r.event_id === d.event_id) && (!d.checkin_status || r.checkin_status === d.checkin_status) &&
    (!q || (r.agency_name + ' ' + r.brand_name + ' ' + r.representative_name).toLowerCase().indexOf(q) >= 0)).map(rOut_).reverse();
}
function pGet_(d) { return rOut_(reg_(d.registration_id)); }
function pCreate_(d, u) { return regCreate_(d, u, true); }
function pUpdate_(d, u) {
  return locked_(() => {
    const r = reg_(d.registration_id); if (r.deleted_at) throw err_('NOT_FOUND', 'Pendaftaran sudah diarsipkan.');
    if (d.version !== undefined && String(d.version) !== r.version) throw err_('CONFLICT', 'Data sudah diubah orang lain. Muat ulang lalu coba lagi.');
    const e = ev_(r.event_id, true), max = Number(setting_('max_participants_per_registration')) || 100;
    const ag = d.agency_name !== undefined ? l_(d.agency_name, 100) : r.agency_name, br = d.brand_name !== undefined ? l_(d.brand_name, 100) : r.brand_name, rp = d.representative_name !== undefined ? l_(d.representative_name, 100) : r.representative_name;
    const n = d.participant_count !== undefined ? Number(d.participant_count) : Number(r.participant_count);
    if (!ag || !br || !rp) throw err_('INVALID', 'Nama agensi, brand, dan perwakilan wajib diisi.');
    if (!Number.isInteger(n) || n < 1 || n > max) throw err_('INVALID', 'Jumlah peserta harus bilangan bulat 1–' + max + '.');
    if (n !== Number(r.participant_count)) {
      if (r.checkin_status === 'checked_in' && !d.confirm) throw err_('CONFIRM_REQUIRED', 'Pendaftaran ini sudah check-in. Kirim konfirmasi untuk mengubah jumlah peserta.');
      const cap = Number(e.capacity) || 0, used = st_(statsMap_(), e.event_id).people - Number(r.participant_count); if (cap && used + n > cap) throw err_('CAPACITY', 'Kuota tidak mencukupi.');
    }
    Object.assign(r, { agency_name: ag, brand_name: br, representative_name: rp, participant_count: n, updated_at: now_(), version: Number(r.version) + 1 }); save_('Registrations', r);
    aud_(u, 'participants.update', 'registration', r.registration_id, ag + ' / ' + br + ' (' + n + ')'); return rOut_(r);
  });
}
function pArchive_(d, u) {
  return locked_(() => {
    const r = reg_(d.registration_id); if (r.checkin_status === 'checked_in' && !d.confirm) throw err_('CONFIRM_REQUIRED', 'Pendaftaran ini sudah check-in. Kirim konfirmasi untuk mengarsipkan.');
    r.deleted_at = now_(); r.registration_status = 'cancelled'; r.updated_at = r.deleted_at; save_('Registrations', r); aud_(u, 'participants.archive', 'registration', r.registration_id, r.agency_name); return { archived: true };
  });
}
function pResetQr_(d, u) {
  return locked_(() => {
    const r = reg_(d.registration_id); if (r.deleted_at) throw err_('NOT_FOUND', 'Pendaftaran sudah diarsipkan.');
    if (r.checkin_status === 'checked_in') throw err_('CHECKED_IN', 'Pendaftaran sudah check-in; QR tidak dapat diganti.');
    const tk = newToken_(); r.qr_token_hash = sha_(tk); r.qr_version = Number(r.qr_version) + 1; r.updated_at = now_(); save_('Registrations', r);
    aud_(u, 'participants.resetQr', 'registration', r.registration_id, 'Reset QR v' + r.qr_version); return ticket_(r, ev_(r.event_id, true), tk);
  });
}
function pExport_(d, u) {
  const e = ev_(d.event_id, true), cv = v => { v = String(v == null ? '' : v); if (/^[=+\-@]/.test(v)) v = "'" + v; return '"' + v.replace(/"/g, '""') + '"'; };
  const rows = pList_({ event_id: e.event_id }).reverse(), head = ['ID Registrasi', 'Agensi', 'Brand', 'Perwakilan', 'Jumlah Orang', 'Status Check-in', 'Waktu Check-in (UTC)', 'Waktu Daftar (UTC)', 'Sumber'];
  const lines = [head.map(cv).join(',')].concat(rows.map(r => [r.registration_id, r.agency_name, r.brand_name, r.representative_name, r.participant_count, r.checkin_status === 'checked_in' ? 'Hadir' : 'Belum hadir', r.checkin_at, r.registered_at, r.source].map(cv).join(',')));
  aud_(u, 'participants.export', 'event', e.event_id, rows.length + ' baris'); return { file_name: 'peserta-' + e.event_id + '.csv', csv: '\uFEFF' + lines.join('\r\n') };
}

/* ===================== Pengguna ===================== */
function usersList_() { const acc = all_('StaffAccess'); return all_('Users').filter(u => !u.deleted_at).map(u => { const x = pu_(u); x.event_ids = acc.filter(a => a.user_id === u.user_id).map(a => a.event_id); return x; }); }
const user_ = id => { const u = all_('Users').find(x => x.user_id === String(id) && !x.deleted_at); if (!u) throw err_('NOT_FOUND', 'Pengguna tidak ditemukan.'); return u; };
const admins_ = (exceptId) => all_('Users').filter(x => !x.deleted_at && x.is_active === 'TRUE' && x.role === 'super_admin' && x.user_id !== exceptId).length;
function uCreate_(d, u) {
  const un = String(d.username || '').trim().toLowerCase(), pw = String(d.password || ''), role = pick_(d.role, ['super_admin', 'staff'], 'staff');
  if (!/^[a-z0-9._-]{3,30}$/.test(un)) throw err_('INVALID', 'Username 3–30 karakter: huruf kecil, angka, titik, strip, garis bawah.');
  if (pw.length < 8) throw err_('INVALID', 'Password minimal 8 karakter.');
  return locked_(() => {
    if (all_('Users').some(x => x.username.toLowerCase() === un && !x.deleted_at)) throw err_('DUPLICATE', 'Username sudah dipakai.');
    const salt = id_('S', 16), t = now_(), n = { user_id: id_('USR'), username: un, display_name: l_(d.display_name, 80) || un, role, password_hash: salt + '$' + pwhash_(pw, salt), is_active: 'TRUE', must_change_password: 'TRUE', failed_attempts: '0', created_at: t, updated_at: t };
    save_('Users', n); aud_(u, 'users.create', 'user', n.user_id, un + ' (' + role + ')'); return pu_(n);
  });
}
function uUpdate_(d, u) {
  return locked_(() => {
    const t = user_(d.user_id), role = d.role !== undefined ? pick_(d.role, ['super_admin', 'staff'], t.role) : t.role, act = d.is_active !== undefined ? !!d.is_active : t.is_active === 'TRUE';
    if (t.user_id === u.user_id && (!act || role !== 'super_admin')) throw err_('INVALID', 'Anda tidak dapat menonaktifkan atau menurunkan akun Anda sendiri.');
    if (t.role === 'super_admin' && (role !== 'super_admin' || !act) && admins_(t.user_id) < 1) throw err_('INVALID', 'Harus ada minimal satu super admin aktif.');
    const roleChanged = role !== t.role; t.display_name = d.display_name !== undefined ? (l_(d.display_name, 80) || t.username) : t.display_name; t.role = role; t.is_active = act ? 'TRUE' : 'FALSE'; t.updated_at = now_(); save_('Users', t);
    if (!act || roleChanged) revoke_(t.user_id); aud_(u, 'users.update', 'user', t.user_id, t.username); return pu_(t);
  });
}
function uResetPw_(d, u) {
  const pw = String(d.new_password || ''); if (pw.length < 8) throw err_('INVALID', 'Password minimal 8 karakter.');
  const t = user_(d.user_id), salt = id_('S', 16); t.password_hash = salt + '$' + pwhash_(pw, salt); t.must_change_password = 'TRUE'; t.failed_attempts = '0'; t.locked_until = ''; t.updated_at = now_(); save_('Users', t);
  revoke_(t.user_id); aud_(u, 'users.resetPassword', 'user', t.user_id, t.username); return { reset: true };
}
function uAccess_(d, u) {
  return locked_(() => {
    const t = user_(d.user_id), ids = Array.isArray(d.event_ids) ? d.event_ids.map(String) : []; ids.forEach(i => ev_(i));
    all_('StaffAccess').filter(a => a.user_id === t.user_id).sort((a, b) => b._row - a._row).forEach(a => sh_('StaffAccess').deleteRow(a._row));
    ids.filter((v, i) => ids.indexOf(v) === i).forEach(i => save_('StaffAccess', { user_id: t.user_id, event_id: i, granted_at: now_(), granted_by: u.user_id }));
    aud_(u, 'users.setAccess', 'user', t.user_id, ids.length + ' event'); return { event_ids: ids };
  });
}

/* ===================== Pengaturan, dashboard, audit ===================== */
function setGet_() { const o = Object.assign({}, DEF_SET); all_('Settings').forEach(r => { if (r.key in o) o[r.key] = r.value; }); o.templates = Object.keys(TPL).map(k => ({ id: k, name: TPL[k][0] })); o.fonts = FONTS; o.timezones = TZS; return o; }
function setSave_(d, u) {
  const v = d.values || {}, rules = { company_name: x => l_(x, 100), app_name: x => l_(x, 100), organizer_info: x => s_(x, 500), logo_asset_id: x => x, default_timezone: x => { if (TZS.indexOf(x) < 0) throw err_('INVALID', 'Zona waktu tidak valid.'); return x; },
    default_template_id: x => { if (!TPL[x]) throw err_('INVALID', 'Template tidak valid.'); return x; }, session_hours: x => String(Math.round(num_(x, 1, 168, 12))), max_participants_per_registration: x => String(Math.round(num_(x, 1, 1000, 100))), register_rate_limit: x => String(Math.round(num_(x, 5, 600, 60))) };
  return locked_(() => {
    Object.keys(v).forEach(k => {
      if (!rules[k]) return; const val = rules[k](String(v[k])); let r = all_('Settings').find(x => x.key === k);
      r = r || { key: k }; r.value = val; r.updated_at = now_(); r.updated_by = u.user_id; save_('Settings', r);
    });
    aud_(u, 'settings.save', 'settings', '', Object.keys(v).join(', ')); return setGet_();
  });
}
function dash_() {
  const sm = statsMap_(), ev = all_('Events').filter(e => !e.deleted_at), tot = EMPTY_ST();
  const per = ev.map(e => { const s = st_(sm, e.event_id); Object.keys(tot).forEach(k => tot[k] += s[k]); return { event_id: e.event_id, title: e.title, event_date: e.event_date, status: e.status, capacity: Number(e.capacity) || 0, stats: s }; }).reverse();
  return { events_total: ev.length, events_active: ev.filter(e => e.status === 'active').length, registrations: tot.registrations, people: tot.people, checked_people: tot.checked_people, not_checked_people: tot.people - tot.checked_people,
    attendance_percent: tot.people ? Math.round(tot.checked_people / tot.people * 100) : 0, per_event: per, recent: auditList_({ limit: 10 }) };
}
function auditList_(d) {
  const lim = Math.min(200, Math.max(1, parseInt(d.limit, 10) || 50)), us = {}; all_('Users').forEach(x => us[x.user_id] = x.username);
  return all_('AuditLogs').slice(-lim).reverse().map(a => ({ at: a.at, user: us[a.user_id] || a.user_id || '-', action: a.action, entity_type: a.entity_type, entity_id: a.entity_id, summary: a.summary }));
}
