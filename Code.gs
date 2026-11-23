/**
 * Backend Google Apps Script - Katalog RS Hermina Balikpapan
 * Database : Google Sheets
 *   Promos, MCU, Packages, Doctors, Contacts, Settings  -> bisa diubah dari Dashboard Admin di website
 *   Admins (login), Leads (pesanan masuk)               -> Admins hanya diubah langsung di Spreadsheet
 *
 * Fitur tambahan: Rekomendasi AI Groq (aksi POST 'recommend', rotasi model otomatis) - lihat bagian REKOMENDASI AI di bawah.
 * CORS: Web App Apps Script (akses "Anyone") otomatis membalas Access-Control-Allow-Origin: *.
 * Header kustom tidak bisa diatur dan doOptions() tidak didukung, jadi frontend mengirim POST
 * dengan Content-Type "text/plain" (request sederhana, tanpa preflight).
 */

const SPREADSHEET_ID = ''; // Kosongkan bila script dibuat dari Extensions > Apps Script di dalam Spreadsheet
const TOKEN_TTL = 21600;   // Sesi admin 6 jam (detik)
const MAX_LOGIN_FAIL = 5;  // Batas gagal login sebelum dikunci 10 menit
const MAX_CONTACTS = 10;
const PURPOSES = ['paket', 'telemedicine', 'booking', 'igd'];

// Sinkronisasi jadwal dokter dari API publik Hermina (lihat bagian "Sinkronisasi" di bawah)
const HERMINA_API = 'https://api.herminahospitals.com/api/v1/public';
const HOSPITAL_ID = 'hermina-balikpapan';
const SCHEDULE_TYPES = ['executive', 'regular']; // executive = Poliklinik Padma (Non BPJS); regular = dugaan untuk Kamala (BPJS), cek dengan testApi()
const DAY_NAMES = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu'];

const SCHEMA = {
  Promos:   ['id', 'title', 'description', 'bannerUrl', 'discountDetails', 'startDate', 'endDate', 'isActive', 'details', 'price', 'category'],
  MCU:      ['id', 'name', 'category', 'items', 'total', 'isActive'],
  Packages: ['id', 'name', 'category', 'description', 'details', 'options', 'terms', 'note', 'isActive'],
  Doctors:  ['id', 'name', 'specialization', 'photoUrl', 'schedule', 'day', 'payer'],
  Contacts: ['id', 'label', 'number', 'isActive', 'purposes'],
  Settings: ['key', 'value'],
  Admins:   ['username', 'password', 'name'],
  Tarif:    ['id', 'category', 'name', 'price', 'note', 'isActive'],
  Notifs:   ['id', 'title', 'body', 'date', 'isActive'],
  Leads:    ['timestamp', 'name', 'phone', 'service', 'notes', 'type', 'doctor', 'poli', 'staff', 'plan'],
  Visits:   ['timestamp', 'date', 'visitor', 'page'],
  Ratings:  ['timestamp', 'name', 'stars', 'kritik', 'kesan', 'pesan'],
  AiCache:  ['key', 'loose', 'json', 'timestamp']
};
const KEY = { Tarif: 'id', Notifs: 'id', Promos: 'id', MCU: 'id', Packages: 'id', Doctors: 'id', Contacts: 'id', Settings: 'key' }; // sheet yang boleh diubah admin
const PREFIX = { Tarif: 'TRF', Notifs: 'NTF', Promos: 'PROMO', MCU: 'MCU', Packages: 'PKG', Doctors: 'DOC', Contacts: 'WA' };
const NUM_FIELDS = ['price', 'total'];
const BOOL_FIELDS = ['isActive'];
const AREA_FIELDS = ['items', 'details', 'options', 'terms', 'description'];

/* ---------- Helper ---------- */

function ss_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet_(name) {
  const ss = ss_();
  let sheet = ss.getSheetByName(name);
  if (!sheet) { sheet = ss.insertSheet(name); sheet.appendRow(SCHEMA[name]); }
  else if (SCHEMA[name]) { // tambahkan kolom baru dari pembaruan versi (sekali per 6 jam)
    const c = CacheService.getScriptCache();
    if (!c.get('sch_' + name)) {
      const hdr = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
      const miss = SCHEMA[name].filter(h => hdr.indexOf(h) < 0);
      if (miss.length) sheet.getRange(1, hdr.length + 1, 1, miss.length).setValues([miss]);
      c.put('sch_' + name, '1', 21600);
    }
  }
  return sheet;
}

function json_(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
function fail_(msg) { return json_({ success: false, error: msg }); }
function toBool_(v) { return v === true || String(v).toLowerCase() === 'true'; }

function clean_(v, max) {
  let s = String(v === undefined || v === null ? '' : v).trim().slice(0, max || 500);
  if (/^[=+\-@]/.test(s)) s = "'" + s; // cegah formula injection di Spreadsheet
  return s;
}

function getData_(name) {
  const rows = getSheet_(name).getDataRange().getValues();
  if (rows.length < 2) return [];
  const tz = Session.getScriptTimeZone();
  const headers = rows[0];
  return rows.slice(1)
    .filter(r => r.some(c => c !== ''))
    .map(r => {
      const o = {};
      headers.forEach((h, j) => { o[h] = r[j] instanceof Date ? Utilities.formatDate(r[j], tz, 'yyyy-MM-dd') : r[j]; });
      return o;
    });
}

function hashPassword_(password) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  return raw.map(b => ((b < 0 ? b + 256 : b).toString(16)).padStart(2, '0')).join('');
}

function requireAuth_(token) { return token && CacheService.getScriptCache().get('sess_' + token); }

function normNumber_(raw) {
  let d = String(raw || '').replace(/\D/g, '');
  if (d.indexOf('0') === 0) d = '62' + d.slice(1);
  else if (d.indexOf('8') === 0) d = '62' + d;
  return /^62\d{8,13}$/.test(d) ? d : '';
}

/* ---------- Router ---------- */

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || '';
    const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    switch (action) {
      case 'getPromos':
        return json_({ success: true, data: getData_('Promos')
          .map(p => { const o = {}; SCHEMA.Promos.forEach(h => { o[h] = p[h]; }); o.isActive = toBool_(p.isActive); return o; }) // kolom tambahan tidak pernah publik
          .filter(p => p.isActive && (!p.startDate || p.startDate <= today) && (!p.endDate || p.endDate >= today)) });
      case 'getMcu':
        return json_({ success: true, data: getData_('MCU').filter(m => toBool_(m.isActive) && m.name)
          .map(m => ({ id: m.id, name: m.name, category: m.category, items: String(m.items || ''), total: Number(m.total) || 0 })) });
      case 'getPackages':
        return json_({ success: true, data: getData_('Packages').filter(p => toBool_(p.isActive) && p.name)
          .map(p => ({ id: p.id, name: p.name, category: p.category, description: p.description, details: p.details, options: p.options, terms: p.terms, note: p.note })) });
      case 'getDoctors':
        return json_({ success: true, data: getData_('Doctors') });
      case 'getContacts':
        return json_({ success: true, data: getData_('Contacts').filter(c => toBool_(c.isActive) && c.number)
          .slice(0, MAX_CONTACTS).map(c => ({ id: c.id, label: c.label, number: String(c.number), purposes: String(c.purposes || 'paket') })) });
      case 'getNotifs':
        return json_({ success: true, data: getData_('Notifs').filter(n => toBool_(n.isActive) && n.title)
          .map(n => ({ id: String(n.id), title: String(n.title), body: String(n.body || ''), date: String(n.date || '').slice(0, 10) }))
          .sort((a, b) => b.id.localeCompare(a.id)).slice(0, 15) });
      case 'getSettings': {
        const s = {};
        getData_('Settings').forEach(r => { if (r.key) s[r.key] = String(r.value); });
        return json_({ success: true, data: s });
      }
      default: return fail_('Aksi GET tidak dikenali.');
    }
  } catch (err) { return fail_(err.toString()); }
}

function doPost(e) {
  try {
    const up = e.postData.contents.indexOf('{"action":"adminUpload"') === 0; // hanya unggah gambar yang boleh besar
    if (e.postData.contents.length > (up ? 950000 : 20000)) return fail_('Data terlalu besar.');
    const data = JSON.parse(e.postData.contents);
    switch (data.action) {
      case 'login':       return handleLogin_(data);
      case 'logout':      CacheService.getScriptCache().remove('sess_' + data.token); return json_({ success: true });
      case 'adminList':   return handleAdminList_(data);
      case 'adminSave':   return handleAdminSave_(data);
      case 'adminDelete': return handleAdminDelete_(data);
      case 'adminSyncDoctors': {
        if (!requireAuth_(data.token)) return fail_('Sesi habis. Silakan login ulang.');
        const r = syncDoctors_();
        return r.ok ? json_({ success: true, message: r.msg }) : fail_(r.msg);
      }
      case 'submitLead':  return handleSubmitLead_(data);
      case 'trackVisit':  return handleVisit_(data);
      case 'submitRating': return handleRating_(data);
      case 'adminStats':  return handleStats_(data);
      case 'adminUpload': return handleUpload_(data);
      case 'adminColumn': return handleColumn_(data);
      case 'recommend':   return handleRecommend_(data);
      case 'ask':         return handleAsk_(data);
      default:            return fail_('Aksi POST tidak dikenali.');
    }
  } catch (err) { return fail_(err.toString()); }
}

/* ---------- Login ---------- */

function handleLogin_(data) {
  const user = String(data.username || data.email || '').trim().toLowerCase();
  const password = String(data.password || '');
  if (!user || !password) return fail_('Username dan kata sandi wajib diisi.');

  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + user.replace(/\W/g, '');
  if (Number(cache.get(failKey) || 0) >= MAX_LOGIN_FAIL) return fail_('Terlalu banyak percobaan gagal. Coba lagi dalam 10 menit.');

  const rows = getSheet_('Admins').getDataRange().getValues();
  if (rows.length < 2) return fail_('Belum ada akun admin. Isi sheet Admins.');
  const h = rows[0];
  const ui = h.indexOf('username') >= 0 ? h.indexOf('username') : h.indexOf('email');
  const pi = h.indexOf('password'), ni = h.indexOf('name');
  const hashed = hashPassword_(password);

  for (let i = 1; i < rows.length; i++) {
    const stored = String(rows[i][pi]);
    if (String(rows[i][ui]).trim().toLowerCase() === user && (stored === hashed || stored === password)) {
      const token = Utilities.getUuid();
      cache.put('sess_' + token, user, TOKEN_TTL);
      cache.remove(failKey);
      return json_({ success: true, token: token, admin: { name: rows[i][ni] || user, username: user } });
    }
  }
  cache.put(failKey, String(Number(cache.get(failKey) || 0) + 1), 600);
  return fail_('Username atau kata sandi salah.');
}

/* ---------- Admin: kelola semua sheet ---------- */

function adminSheet_(data) {
  const n = String(data.sheet || '');
  if (!KEY[n]) throw new Error('Sheet tidak dikenal.');
  return n;
}

function conv_(h, v) {
  if (BOOL_FIELDS.indexOf(h) >= 0) return v === undefined ? true : toBool_(v);
  if (NUM_FIELDS.indexOf(h) >= 0) { const d = String(v === undefined ? '' : v).replace(/\D/g, ''); return d ? Number(d) : ''; }
  if (AREA_FIELDS.indexOf(h) >= 0) {
    return String(v || '').split('\n').map(l => clean_(l.replace(/^[-•*\s]+/, ''), 300)).filter(Boolean).slice(0, 80).join('\n');
  }
  return clean_(v, 600);
}

function handleAdminList_(data) {
  if (!requireAuth_(data.token)) return fail_('Sesi habis. Silakan login ulang.');
  const n = adminSheet_(data);
  const list = getData_(n).map(r => {
    BOOL_FIELDS.forEach(f => { if (f in r) r[f] = toBool_(r[f]); });
    if (n === 'Contacts') r.number = String(r.number);
    return r;
  });
  const sh = getSheet_(n), hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  return json_({ success: true, data: list, extra: hdr.filter(h => h && SCHEMA[n].indexOf(h) < 0) });
}

function validateContacts_(rows, hdr) {
  const num = hdr.indexOf('number'), act = hdr.indexOf('isActive');
  if (rows.length > MAX_CONTACTS) return 'Maksimal ' + MAX_CONTACTS + ' kontak WhatsApp.';
  if (!rows.some(r => toBool_(r[act]))) return 'Minimal 1 kontak WhatsApp harus aktif.';
  const seen = {};
  for (let i = 0; i < rows.length; i++) {
    const k = String(rows[i][num]);
    if (k && seen[k]) return 'Nomor ' + k + ' terdaftar dua kali.';
    seen[k] = 1;
  }
  return '';
}

function handleAdminSave_(data) {
  if (!requireAuth_(data.token)) return fail_('Sesi habis. Silakan login ulang.');
  const n = adminSheet_(data), key = KEY[n], row = data.row || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet_(n);
    const all = sheet.getDataRange().getValues();
    const hdr = all[0], kc = hdr.indexOf(key);
    let id = clean_(row[key], 80), idx = -1;
    if (n === 'Settings' && !id) return fail_('Kunci pengaturan wajib diisi.');
    if (id) for (let i = 1; i < all.length; i++) if (String(all[i][kc]) === id) { idx = i; break; }
    if (!id) id = PREFIX[n] + '-' + Date.now();

    const rec = {};
    SCHEMA[n].forEach(h => { rec[h] = conv_(h, row[h]); });
    hdr.forEach(h => { if (h && SCHEMA[n].indexOf(h) < 0) rec[h] = row[h] !== undefined ? clean_(row[h], 600) : (idx >= 0 ? all[idx][hdr.indexOf(h)] : ''); }); // kolom tambahan
    rec[key] = id;

    if (n === 'Doctors' && (!rec.name || !rec.day || !rec.schedule)) return fail_('Nama dokter, hari, dan jam praktik wajib diisi.');
    if ((n === 'MCU' || n === 'Packages') && !rec.name) return fail_('Nama paket wajib diisi.');
    if (n === 'Promos' && !rec.title) return fail_('Judul promo wajib diisi.');
    if (n === 'Contacts') {
      rec.number = normNumber_(row.number);
      if (!rec.label) return fail_('Nama petugas wajib diisi.');
      if (!rec.number) return fail_('Nomor WhatsApp tidak valid. Contoh: 08123456789.');
      rec.purposes = String(row.purposes || 'paket').split(',').map(s => s.trim()).filter(p => PURPOSES.indexOf(p) >= 0).join(',') || 'paket';
    }
    const out = hdr.map(h => rec[h] === undefined ? '' : rec[h]);

    if (n === 'Contacts') {
      const after = all.slice(1).filter(r => r[kc] !== '');
      const pos = after.findIndex(r => String(r[kc]) === id);
      if (pos < 0) after.push(out); else after[pos] = out;
      const err = validateContacts_(after, hdr);
      if (err) return fail_(err);
    }

    const target = idx < 0 ? sheet.getLastRow() + 1 : idx + 1;
    const range = sheet.getRange(target, 1, 1, out.length);
    if (n === 'Contacts' || n === 'Settings') range.setNumberFormat('@');
    range.setValues([out]);
    try { notifyChange_(n, idx < 0, out, hdr); } catch (e) {}
    return json_({ success: true, id: id });
  } finally {
    lock.releaseLock();
  }
}

function handleAdminDelete_(data) {
  if (!requireAuth_(data.token)) return fail_('Sesi habis. Silakan login ulang.');
  const n = adminSheet_(data), key = KEY[n];
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet_(n);
    const all = sheet.getDataRange().getValues();
    const hdr = all[0], kc = hdr.indexOf(key);
    for (let i = 1; i < all.length; i++) {
      if (String(all[i][kc]) === String(data.id)) {
        if (n === 'Contacts') {
          const rest = all.slice(1).filter((r, j) => j !== i - 1 && r[kc] !== '');
          const err = validateContacts_(rest, hdr);
          if (err) return fail_(err);
        }
        sheet.deleteRow(i + 1);
        return json_({ success: true });
      }
    }
    return fail_('Data tidak ditemukan.');
  } finally {
    lock.releaseLock();
  }
}

/* ---------- Pesanan masuk (Leads) ---------- */

function handleSubmitLead_(data) {
  const name = clean_(data.name, 100), phone = String(data.phone || '').trim().slice(0, 20);
  if (!name) return fail_('Nama wajib diisi.');
  if (phone && !/^[0-9+\s\-]{8,20}$/.test(phone)) return fail_('Format nomor telepon tidak valid.');

  const cache = CacheService.getScriptCache();
  const key = 'lead_' + (phone || name).replace(/\W/g, '').slice(0, 40) + '_' + clean_(data.service, 40).replace(/\W/g, '');
  if (cache.get(key)) return fail_('Permintaan baru saja terkirim. Mohon tunggu beberapa saat.');
  cache.put(key, '1', 30);

  const sh = getSheet_('Leads');
  if (sh.getLastColumn() < SCHEMA.Leads.length) sh.getRange(1, 1, 1, SCHEMA.Leads.length).setValues([SCHEMA.Leads]);
  const svc = clean_(data.service, 150), notes = clean_(data.notes, 800), parts = svc.split(': ');
  const m = re => { const x = re.exec(notes); return x ? x[1].trim() : ''; };
  let doctor = '', poli = '';
  if (/^(Booking dokter|Telemedicine)/.test(svc)) {
    const rest = parts.slice(1).join(': ');
    if (rest.indexOf(' - ') > 0) { poli = rest.split(' - ')[0]; doctor = rest.split(' - ').slice(1).join(' - '); } else doctor = rest;
    if (!poli && doctor) { const d = getData_('Doctors').filter(x => x.name === doctor)[0]; poli = d ? String(d.specialization || '') : ''; }
  }
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    sh.appendRow([now_(), name, phone ? "'" + phone : '-', svc, notes, clean_(parts[0], 60), doctor, poli, m(/Petugas: ([^|]+)/), m(/Rencana: ([^|]+)/)]);
  } finally { lock.releaseLock(); }
  return json_({ success: true });
}

/* ---------- Sinkronisasi jadwal dokter dari web Hermina ---------- */
/*
 * Data diambil di server (Apps Script) agar tidak terkena batasan CORS browser dan agar ratusan
 * permintaan tidak dilakukan dari HP pengunjung. Hasilnya ditulis ke sheet Doctors, lalu website
 * membaca sheet itu seperti biasa. Jalankan installSyncTrigger() sekali agar otomatis tiap jam.
 * Mematikan: sheet Settings, key doctor_sync = off.
 */

function apiOpts_() {
  return { method: 'get', muteHttpExceptions: true, headers: { Accept: 'application/json', 'User-Agent': 'Mozilla/5.0 (compatible; KatalogHerminaBPP/1.0)' } };
}

function fetchAllDoctors_() {
  const list = [], seen = {};
  let page = 1, last = 1;
  do {
    const res = UrlFetchApp.fetch(HERMINA_API + '/doctors?hospital_id=' + HOSPITAL_ID + '&per_page=100&page=' + page, apiOpts_());
    if (res.getResponseCode() !== 200) throw new Error('API daftar dokter HTTP ' + res.getResponseCode());
    const j = JSON.parse(res.getContentText());
    (j.data || []).forEach(d => { const a = d.attributes; if (a && a.slug && !seen[a.slug]) { seen[a.slug] = 1; list.push(a); } });
    last = (j.pagination && j.pagination.last) || 1;
    page++;
  } while (page <= last && page <= 50);
  return list;
}

function fetchSchedules_(doctors) {
  const reqs = [], meta = [];
  doctors.forEach(d => SCHEDULE_TYPES.forEach(t => {
    reqs.push(Object.assign({ url: HERMINA_API + '/doctors/' + encodeURIComponent(d.slug) + '/schedules?schedule_type=' + t + '&type=table&lang=id' }, apiOpts_()));
    meta.push({ d: d, t: t });
  }));
  const out = [];
  let bad = 0;
  for (let i = 0; i < reqs.length; i += 20) {
    const resps = UrlFetchApp.fetchAll(reqs.slice(i, i + 20));
    resps.forEach((r, k) => {
      const code = r.getResponseCode(), m = meta[i + k];
      if (code === 200) { try { out.push({ d: m.d, t: m.t, json: JSON.parse(r.getContentText()) }); } catch (e) { bad++; } }
      else if (code !== 404) bad++; // 404 = dokter tidak punya jadwal tipe ini
    });
    Utilities.sleep(300); // jangan membebani server Hermina
  }
  return { out: out, bad: bad, total: reqs.length };
}

function cleanSpec_(s) {
  s = String(s || '').replace(/\s+/g, ' ').replace(/\s*-\s*/g, ' - ').trim();
  if (!s) return 'Lainnya';
  return /^(dokter|spesialis|poli)/i.test(s) ? s : 'Spesialis ' + s;
}

/** Ubah hasil API menjadi baris sheet Doctors: id, name, specialization, photoUrl, schedule, day, payer. */
function buildDoctorRows_(results) {
  const map = {}; // kunci: slug|payer|hari -> daftar slot
  const info = {};
  results.forEach(res => {
    const d = res.d, data = (res.json && res.json.data) || {};
    Object.keys(data).forEach(hosp => {
      if (hosp.toLowerCase().indexOf('balikpapan') < 0) return;
      const clinics = data[hosp] || {};
      Object.keys(clinics).forEach(clinic => {
        const payer = /kamala/i.test(clinic) ? 'BPJS' : /padma/i.test(clinic) ? 'Non BPJS' : (res.t === 'executive' ? 'Non BPJS' : 'BPJS');
        (clinics[clinic] || []).forEach(s => {
          let di = Number(s.day_integer) - 1;
          if (!(di >= 0 && di <= 6)) di = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'].indexOf(String(s.day || '').toLowerCase());
          if (di < 0 || !s.from_time || !s.to_time) return;
          const k = d.slug + '|' + payer + '|' + di;
          const cek = !!s.partial_holiday && String(s.partial_holiday_schedule || '').replace(/\s/g, '') === s.from_time + '-' + s.to_time;
          (map[k] = map[k] || []).push({ from: s.from_time, to: s.to_time, cek: cek });
          info[k] = { d: d, payer: payer, di: di };
        });
      });
    });
  });
  const rows = [];
  Object.keys(map).forEach(k => {
    const slots = map[k].sort((a, b) => a.from < b.from ? -1 : a.from > b.from ? 1 : 0), seen = {}, parts = [];
    slots.forEach(x => {
      const t = x.from.replace(':', '.') + '-' + x.to.replace(':', '.') + (x.cek ? ' (cek libur)' : '');
      if (!seen[t]) { seen[t] = 1; parts.push(t); }
    });
    const i = info[k], a = i.d;
    rows.push([
      'API-' + a.id + '-' + (i.payer === 'BPJS' ? 'B' : 'N') + '-' + (i.di + 1),
      String(a.full_name || '').replace(/\s+/g, ' ').trim(),
      cleanSpec_(a.speciality),
      (a.avatar && (a.avatar.thumb && a.avatar.thumb.url || a.avatar.url)) || '',
      parts.join(' / '),
      DAY_NAMES[i.di],
      i.payer
    ]);
  });
  rows.sort((x, y) => (x[2] + x[1]).localeCompare(y[2] + y[1]) || DAY_NAMES.indexOf(x[5]) - DAY_NAMES.indexOf(y[5]));
  return rows;
}

function setSetting_(key, value) {
  const sheet = getSheet_('Settings');
  const all = sheet.getDataRange().getValues();
  let row = -1;
  for (let i = 1; i < all.length; i++) if (String(all[i][0]) === key) { row = i + 1; break; }
  if (row < 0) row = sheet.getLastRow() + 1;
  const r = sheet.getRange(row, 1, 1, 2);
  r.setNumberFormat('@');
  r.setValues([[key, value]]);
}

function syncDoctors_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, msg: 'Sinkronisasi sedang berjalan. Coba lagi sebentar lagi.' };
  try {
    const set = {};
    getData_('Settings').forEach(r => { set[r.key] = String(r.value); });
    if (String(set.doctor_sync || 'on').toLowerCase() === 'off') return { ok: false, msg: 'Sinkronisasi dimatikan (Settings: doctor_sync = off).' };

    const docs = fetchAllDoctors_();
    if (!docs.length) return { ok: false, msg: 'API tidak mengembalikan dokter. Data lama dipertahankan.' };
    const f = fetchSchedules_(docs);
    if (f.bad > f.total * 0.2) return { ok: false, msg: 'Terlalu banyak permintaan gagal (' + f.bad + ' dari ' + f.total + '). Data lama dipertahankan.' };
    const rows = buildDoctorRows_(f.out);
    if (!rows.length) return { ok: false, msg: 'Tidak ada jadwal yang terbaca. Data lama dipertahankan.' };

    const ss = ss_(), sheet = getSheet_('Doctors');
    if (!ss.getSheetByName('Doctors_Excel')) { // arsipkan data lama satu kali
      const old = sheet.getDataRange().getValues();
      if (old.length > 1) { const b = ss.insertSheet('Doctors_Excel'); b.getRange(1, 1, old.length, old[0].length).setValues(old); }
    }
    const width = SCHEMA.Doctors.length;
    if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, width).clearContent();
    const range = sheet.getRange(2, 1, rows.length, width);
    range.setNumberFormat('@');
    range.setValues(rows);
    setSetting_('jadwal_update', Utilities.formatDate(new Date(), 'Asia/Makassar', 'd MMM yyyy HH:mm') + ' WITA (otomatis dari website Hermina)');
    const names = {};
    rows.forEach(r => { names[r[1]] = 1; });
    return { ok: true, msg: 'Sinkronisasi selesai: ' + Object.keys(names).length + ' dokter, ' + rows.length + ' baris jadwal.' };
  } finally {
    lock.releaseLock();
  }
}

/** Dipanggil oleh trigger waktu. */
function syncDoctorsJob() { Logger.log(JSON.stringify(syncDoctors_())); }

/** Jalankan SEKALI dari editor: menjadwalkan sinkronisasi otomatis tiap 1 jam. */
function installSyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'syncDoctorsJob') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncDoctorsJob').timeBased().everyHours(1).create();
  Logger.log('Trigger sinkronisasi terpasang (tiap 1 jam).');
}

/** Jalankan dari editor untuk memeriksa API: lihat hasilnya di Execution log. */
function testApi() {
  const docs = fetchAllDoctors_();
  Logger.log('Jumlah dokter dari API: ' + docs.length);
  const d = docs[0];
  ['executive', 'regular', ''].forEach(t => {
    const url = HERMINA_API + '/doctors/' + encodeURIComponent(d.slug) + '/schedules?' + (t ? 'schedule_type=' + t + '&' : '') + 'type=table&lang=id';
    const r = UrlFetchApp.fetch(url, apiOpts_());
    Logger.log('[' + (t || 'tanpa schedule_type') + '] HTTP ' + r.getResponseCode() + ' ' + r.getContentText().slice(0, 600));
  });
  Logger.log('Uji ubah ke baris Doctors (dokter pertama saja): ' + JSON.stringify(buildDoctorRows_([{ d: d, t: 'executive', json: JSON.parse(UrlFetchApp.fetch(HERMINA_API + '/doctors/' + encodeURIComponent(d.slug) + '/schedules?schedule_type=executive&type=table&lang=id', apiOpts_()).getContentText()) }])));
}

/* ---------- Setup & bantuan (jalankan dari editor) ---------- */

/** Membuat sheet yang belum ada beserta judul kolomnya. Tidak menghapus atau menimpa data. */
function setup() {
  const ss = ss_();
  Object.keys(SCHEMA).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(SCHEMA[name]);
    sh.getRange(1, 1, 1, SCHEMA[name].length).setFontWeight('bold').setBackground('#DDF3E7');
    sh.setFrozenRows(1);
  });
  const msg = 'Setup selesai. Pastikan sheet Admins berisi akun: username, password (hash SHA-256), name.';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { Logger.log(msg); }
}

/** Ganti teks di bawah dengan sandi baru, jalankan, lalu salin hash dari Logs ke kolom password sheet Admins. */
function generateHash() {
  Logger.log(hashPassword_('GANTI_DENGAN_SANDI_BARU'));
}

/* =====================================================================
 * REKOMENDASI AI (Groq) — ditambahkan 17 Nov 2026
 * ---------------------------------------------------------------------
 * Cara pasang API key (JANGAN tulis key di kode / di index.html):
 *   Apps Script editor > Project Settings (ikon gir) > Script Properties
 *   > Add script property:  Name = GROQ_API_KEY , Value = key dari
 *   https://console.groq.com/keys
 * Opsional: property GROQ_MODELS = daftar model dipisah koma, urut dari
 * yang paling diutamakan. Bila kosong, dipakai GROQ_MODELS_DEFAULT.
 *
 * Rotasi model: limit Groq dihitung PER MODEL. Bila model pertama
 * membalas 429 (limit) / 5xx / jawaban tidak valid, permintaan otomatis
 * diteruskan ke model berikutnya. Bila semua model habis, pengunjung
 * tetap menerima rekomendasi cadangan berbasis aturan (scoring katalog),
 * jadi fitur tidak pernah tampil kosong.
 * Hasil profil yang sama di-cache 1 jam di CacheService agar hemat kuota.
 * ===================================================================== */

const GROQ_API_URL = 'https://api.groq.com/openai/v1/chat/completions';
const GROQ_MODELS_DEFAULT = ['openai/gpt-oss-120b', 'openai/gpt-oss-20b', 'qwen/qwen3.8-27b'];

function groqKey_() {
  return String(PropertiesService.getScriptProperties().getProperty('GROQ_API_KEY') || '').trim();
}

function groqModels_() {
  const raw = String(PropertiesService.getScriptProperties().getProperty('GROQ_MODELS') || '').trim();
  if (raw) return raw.split(',').map(s => s.trim()).filter(Boolean);
  const c = CacheService.getScriptCache(), hit = c.get('groq_models');
  if (hit) return JSON.parse(hit);
  let list = GROQ_MODELS_DEFAULT.slice();
  try { // pakai SEMUA model chat aktif milik akun Groq Anda
    const r = UrlFetchApp.fetch('https://api.groq.com/openai/v1/models', { headers: { Authorization: 'Bearer ' + groqKey_() }, muteHttpExceptions: true });
    if (r.getResponseCode() === 200) {
      const ids = JSON.parse(r.getContentText()).data.filter(x => x.active !== false).map(x => x.id)
        .filter(id => !/whisper|tts|guard|orpheus|playai|embed|safeguard|compound/i.test(id));
      if (ids.length) list = GROQ_MODELS_DEFAULT.filter(x => ids.indexOf(x) >= 0).concat(ids.filter(id => GROQ_MODELS_DEFAULT.indexOf(id) < 0));
    }
  } catch (e) {}
  c.put('groq_models', JSON.stringify(list), 21600);
  return list;
}

function now_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'); }
function today_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'); }
function dstr_(x) { return x instanceof Date ? Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(x).slice(0, 10); }

/** Cache jawaban AI permanen di sheet AiCache (kolom 1 = kunci tepat, 2 = kunci longgar). */
function aiFind_(col, val) {
  const sh = getSheet_('AiCache'), n = sh.getLastRow(); if (n < 2) return null;
  const v = sh.getRange(2, col, n - 1, 1).getValues();
  for (let i = v.length - 1; i >= 0; i--) if (v[i][0] === val) { try { return JSON.parse(sh.getRange(i + 2, 3).getValue()); } catch (e) { return null; } }
  return null;
}
function aiStore_(key, loose, out) {
  try { const l = LockService.getScriptLock(); l.waitLock(5000);
    try { getSheet_('AiCache').appendRow([key, loose, JSON.stringify(out).slice(0, 45000), now_()]); } finally { l.releaseLock(); } } catch (e) {}
}

/* ---------- Statistik, kunjungan, penilaian ---------- */
function handleVisit_(data) {
  const vid = clean_(data.vid, 40).replace(/\W/g, ''); if (!vid) return fail_('Data tidak valid.');
  const c = CacheService.getScriptCache(); if (c.get('vis_' + vid)) return json_({ success: true });
  c.put('vis_' + vid, '1', 1800); // 1 kunjungan per sesi 30 menit
  const l = LockService.getScriptLock(); l.waitLock(10000);
  try { getSheet_('Visits').appendRow([now_(), today_(), vid, clean_(data.page, 40) || 'home']); } finally { l.releaseLock(); }
  return json_({ success: true });
}

function handleRating_(data) {
  if (data.website) return json_({ success: true }); // honeypot anti-bot
  const stars = Math.round(Number(data.stars)), name = clean_(data.name, 80);
  if (!(stars >= 1 && stars <= 5)) return fail_('Pilih bintang 1 sampai 5.');
  if (!name) return fail_('Nama wajib diisi.');
  const c = CacheService.getScriptCache(), k = 'rate_' + name.replace(/\W/g, '').slice(0, 30);
  if (c.get(k)) return fail_('Penilaian baru saja terkirim.'); c.put(k, '1', 60);
  const l = LockService.getScriptLock(); l.waitLock(10000);
  try { getSheet_('Ratings').appendRow([now_(), name, stars, clean_(data.kritik, 800), clean_(data.kesan, 500), clean_(data.pesan, 500)]); } finally { l.releaseLock(); }
  return json_({ success: true });
}

/** Hanya untuk admin login. Nama & nomor pasien TIDAK dikirim ke browser. */
function handleStats_(data) {
  if (!requireAuth_(data.token)) return fail_('Sesi habis. Silakan login ulang.');
  const v = {};
  getData_('Visits').forEach(r => { const d = dstr_(r.date), o = v[d] || (v[d] = { w: 0, u: {} }); o.w++; o.u[r.visitor] = 1; });
  const visits = Object.keys(v).sort().map(d => [d, v[d].w, Object.keys(v[d].u).length]);
  const leads = getData_('Leads').map(r => [dstr_(r.timestamp), r.type || String(r.service).split(':')[0], r.doctor || '', r.poli || '', r.staff || '', r.plan || '']);
  const rt = getData_('Ratings'), sum = rt.reduce((a, r) => a + (Number(r.stars) || 0), 0);
  return json_({ success: true, visits: visits, leads: leads, ratings: { count: rt.length, avg: rt.length ? Math.round(sum / rt.length * 100) / 100 : 0 } });
}

/** Katalog ringkas untuk AI & scoring cadangan. Hanya data aktif dari sheet. */
function catalogForAi_() {
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const out = [];
  getData_('MCU').filter(m => toBool_(m.isActive) && m.name).forEach(m => {
    out.push({ type: 'MCU', name: String(m.name), category: String(m.category || ''),
      price: Number(m.total) || 0,
      desc: String(m.items || '').split('\n').slice(0, 8).join(', ').slice(0, 400) });
  });
  getData_('Packages').filter(p => toBool_(p.isActive) && p.name).forEach(p => {
    const prices = String(p.options || '').split('\n').map(l => Number(String(l.split('|')[1] || '').replace(/\D/g, '')) || 0).filter(Boolean);
    out.push({ type: 'Paket', name: String(p.name), category: String(p.category || ''),
      price: prices.length ? Math.min.apply(null, prices) : 0,
      desc: (String(p.description || '') + ' ' + String(p.details || '').replace(/\n/g, ', ')).slice(0, 400) });
  });
  getData_('Promos').map(p => Object.assign(p, { isActive: toBool_(p.isActive) }))
    .filter(p => p.isActive && p.title && (!p.startDate || p.startDate <= today) && (!p.endDate || p.endDate >= today))
    .forEach(p => {
      out.push({ type: 'Promo', name: String(p.title), category: 'Promo',
        price: Number(p.price) || 0,
        desc: (String(p.description || '') + ' ' + String(p.details || '').replace(/\n/g, ', ')).slice(0, 300) });
    });
  return out;
}

/* Kata kunci tujuan -> penanda di nama/kategori/deskripsi katalog */
const RECO_KEYWORDS = {
  'mcu-rutin':   ['mcu simple', 'mcu silver', 'paket umum', 'medical check'],
  'jantung':     ['cardio', 'jantung', 'treadmill', 'echocardiography', 'ekg'],
  'wanita':      ['wanita', 'femina', 'pap smear', 'mammae', 'mammography', 'hpv', 'gardasil'],
  'pranikah':    ['premarital', 'pra-nikah', 'pranikah'],
  'kerja':       ['pre employment', 'field worker', 'kerja', 'karyawan', 'occupational'],
  'hamil':       ['persalinan', 'partus', 'sc eracs', 'kehamilan'],
  'vaksin':      ['vaksin', 'influenza', 'pneumonia', 'hpv', 'gardasil', 'meningitis', 'umroh', 'haji'],
  'anak':        ['anak', 'influenza anak', 'pneumonia anak'],
  'lansia':      ['lansia', 'femina', 'diamond', 'platinum'],
  'homecare':    ['home care', 'homecare', 'perawat'],
  'lainnya':     []
};

function recoScore_(item, prof) {
  const hay = (item.name + ' ' + item.category + ' ' + item.desc).toLowerCase();
  let score = 0;
  (RECO_KEYWORDS[prof.goal] || []).forEach(k => { if (hay.indexOf(k) >= 0) score += 30; });
  String(prof.concern || '').toLowerCase().split(/[^a-z0-9]+/).filter(w => w.length >= 4)
    .forEach(w => { if (hay.indexOf(w) >= 0) score += 6; });
  if (prof.gender === 'Wanita' && /wanita|femina/.test(hay)) score += 18;
  if (prof.gender === 'Pria' && /\bpria\b/.test(hay)) score += 18;
  if (prof.gender === 'Wanita' && /\bpria\b/.test(hay) && !/wanita/.test(hay)) score -= 25;
  if (prof.gender === 'Pria' && /wanita/.test(hay) && !/\bpria\b/.test(hay)) score -= 25;
  if (prof.age === 'Lansia 60 tahun ke atas' && /lansia|diamond|platinum|cardio/.test(hay)) score += 10;
  if (prof.age === 'Anak (di bawah 18 tahun)' && /anak/.test(hay)) score += 15;
  if (prof.budget > 0 && item.price > 0) {
    if (item.price <= prof.budget) score += 12; else score -= Math.min(20, Math.round((item.price - prof.budget) / 1000000));
  }
  return score;
}

function fallbackRecommend_(prof, catalog) {
  const ranked = catalog.map(it => ({ it: it, s: recoScore_(it, prof) }))
    .sort((a, b) => b.s - a.s || (a.it.price || 999999999) - (b.it.price || 999999999));
  const picked = [];
  const seenCat = {};
  ranked.forEach(r => { // utamakan variasi kategori, maksimal 3
    if (picked.length >= 3) return;
    if (picked.length >= 2 && seenCat[r.it.category]) return;
    seenCat[r.it.category] = 1; picked.push(r.it);
  });
  ranked.forEach(r => { if (picked.length < 3 && picked.indexOf(r.it) < 0) picked.push(r.it); });
  return {
    intro: 'Berdasarkan kebutuhan yang Anda pilih, berikut layanan RS Hermina Balikpapan yang paling sesuai untuk dipertimbangkan:',
    recommendations: picked.slice(0, 3).map(it => ({
      name: it.name, type: it.type,
      reason: 'Sesuai dengan tujuan "' + prof.goalLabel + '"' + (it.category ? ' pada kategori ' + it.category : '') +
        (it.price ? ', dengan tarif ' + 'Rp ' + Number(it.price).toLocaleString('id-ID') : '') +
        '. Silakan buka detailnya untuk melihat rincian isi paket sebelum memutuskan.'
    })),
    closing: 'Rekomendasi ini bersifat informasi umum, bukan diagnosis. Untuk keluhan atau kondisi khusus, konsultasikan dulu dengan dokter kami agar pilihan layanannya tepat.'
  };
}

function handleRecommend_(data) {
  const p = data.profile || {};
  const prof = {
    age: clean_(p.age, 60), gender: clean_(p.gender, 30),
    goal: Object.prototype.hasOwnProperty.call(RECO_KEYWORDS, String(p.goal || '')) ? String(p.goal) : 'lainnya',
    goalLabel: clean_(p.goalLabel, 80) || 'Cek kesehatan',
    concern: clean_(p.concern, 300),
    budget: Math.max(0, Math.min(100000000, Number(p.budget) || 0)),
    payer: clean_(p.payer, 40)
  };

  const cache = CacheService.getScriptCache();
  const cacheKey = 'reco_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,
    [prof.age, prof.gender, prof.goal, prof.concern.toLowerCase(), prof.budget, prof.payer].join('|'))
    .map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('').slice(0, 24);
  const loose = 'L_' + [prof.goal, prof.payer, prof.gender].join('|');
  const cached = cache.get(cacheKey);
  if (cached) { try { const j = JSON.parse(cached); j.cached = true; return json_(j); } catch (e) {} }

  const stored = aiFind_(1, cacheKey); // cache permanen di sheet AiCache
  if (stored) { stored.cached = true; return json_(stored); }
  const catalog = catalogForAi_();
  if (!catalog.length) return fail_('Katalog masih kosong.');
  const fallback = fallbackRecommend_(prof, catalog);

  const key = groqKey_();
  if (!key) {
    const res = { success: true, source: 'fallback', note: 'GROQ_API_KEY belum dipasang di Script Properties.', data: fallback };
    return json_(res);
  }

  const validNames = {};
  catalog.forEach(it => { validNames[it.name] = it; });

  const systemPrompt = [
    'Anda adalah Sales Marketing Professional RS Hermina Balikpapan: sopan, hangat, meyakinkan, dan sangat pandai membantu calon pasien mantap mengambil keputusan.',
    'Tugas Anda: merekomendasikan MAKSIMAL 3 layanan HANYA dari daftar katalog yang diberikan.',
    'Aturan wajib:',
    '1. Dilarang mengarang nama layanan, harga, atau fasilitas di luar katalog.',
    '2. Dilarang mendiagnosis penyakit dan dilarang menjanjikan kesembuhan/hasil pasti.',
    '3. Gaya sales profesional: tonjolkan manfaat nyata dari katalog, buat pengunjung merasa yakin, lalu SELALU ajak langsung appointment/hubungi petugas hari ini. Dilarang menakut-nakuti, membuat kelangkaan atau urgensi palsu, atau menghalangi pengunjung berkonsultasi dengan dokter.',
    '4. Bila pengunjung menyebut gejala berat/darurat (nyeri dada, sesak, perdarahan, tidak sadar), sarankan IGD 24 jam, bukan membeli paket.',
    '5. Tulis dalam Bahasa Indonesia yang natural.',
    '6. Bila pertanyaan di luar katalog, jangan menebak: arahkan dengan sopan agar menghubungi petugas Hermina lewat tombol WhatsApp.',
    'Balas HANYA dengan JSON valid tanpa teks lain, format:',
    '{"intro":"1-2 kalimat pembuka personal","recommendations":[{"name":"nama persis dari katalog","type":"MCU/Paket/Promo","reason":"alasan personal 1-3 kalimat, sebut tarif hanya bila ada di katalog"}],"closing":"1-2 kalimat ajakan appointment langsung hari ini + anjuran konsultasi dokter bila perlu"}'
  ].join('\n');

  const userPrompt = 'Profil pengunjung:\n' +
    '- Usia: ' + (prof.age || '-') + '\n- Untuk: ' + (prof.gender || '-') +
    '\n- Tujuan: ' + prof.goalLabel + '\n- Kebutuhan/keluhan umum (bukan untuk diagnosis): ' + (prof.concern || '-') +
    '\n- Budget maksimal: ' + (prof.budget ? 'Rp ' + prof.budget.toLocaleString('id-ID') : 'Fleksibel') +
    '\n- Penjaminan: ' + (prof.payer || '-') +
    '\n\nKatalog resmi (JSON):\n' + JSON.stringify(catalog);

  const m0 = groqModels_(), off = Math.floor(Math.random() * m0.length); // sebar beban antar model
  const models = m0.slice(off).concat(m0.slice(0, off)).slice(0, 6);
  const attempts = [];
  for (let i = 0; i < models.length; i++) {
    const model = models[i];
    try {
      const res = UrlFetchApp.fetch(GROQ_API_URL, {
        method: 'post', muteHttpExceptions: true,
        contentType: 'application/json',
        headers: { Authorization: 'Bearer ' + key },
        payload: JSON.stringify({
          model: model, temperature: 0.4, max_tokens: 900,
          messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userPrompt }]
        })
      });
      const code = res.getResponseCode();
      if (code === 429 || code >= 500) { attempts.push(model + ': HTTP ' + code); continue; } // limit/server sibuk -> rotasi
      if (code !== 200) { attempts.push(model + ': HTTP ' + code); continue; }
      const body = JSON.parse(res.getContentText());
      const text = body.choices && body.choices[0] && body.choices[0].message && body.choices[0].message.content || '';
      const start = text.indexOf('{'), end = text.lastIndexOf('}');
      if (start < 0 || end <= start) { attempts.push(model + ': jawaban bukan JSON'); continue; }
      const parsed = JSON.parse(text.slice(start, end + 1));
      const recs = (parsed.recommendations || []).filter(r => r && validNames[String(r.name || '')])
        .slice(0, 3).map(r => ({ name: String(r.name), type: validNames[String(r.name)].type, reason: clean_(r.reason, 500) }));
      if (!recs.length) { attempts.push(model + ': tidak ada nama yang cocok dengan katalog'); continue; }
      const out = { success: true, source: 'groq', model: model, attempts: attempts,
        data: { intro: clean_(parsed.intro, 400) || fallback.intro, recommendations: recs,
                closing: clean_(parsed.closing, 300) || fallback.closing } };
      try { cache.put(cacheKey, JSON.stringify(out), 3600); } catch (e) {}
      aiStore_(cacheKey, loose, out);
      return json_(out);
    } catch (err) {
      attempts.push(model + ': ' + err.toString().slice(0, 80));
    }
  }
  const alt = aiFind_(2, loose); // AI down: pakai jawaban tersimpan yang paling mirip
  if (alt) { alt.cached = true; alt.note = 'Jawaban tersimpan (AI sedang sibuk).'; return json_(alt); }
  return json_({ success: true, source: 'fallback', note: 'Semua model Groq sedang mencapai limit. Rekomendasi cadangan ditampilkan.', attempts: attempts, data: fallback });
}

/** Jalankan dari editor untuk memeriksa API key Groq (key TIDAK ditampilkan di log). */
function testGroq() {
  const key = groqKey_();
  if (!key) { Logger.log('GROQ_API_KEY belum dipasang. Project Settings > Script Properties > Add script property.'); return; }
  Logger.log('GROQ_API_KEY terpasang (panjang ' + key.length + ' karakter). Model rotasi: ' + groqModels_().join(' -> '));
  const res = UrlFetchApp.fetch('https://api.groq.com/openai/v1/models', { method: 'get', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + key } });
  Logger.log('GET /models HTTP ' + res.getResponseCode() + ' (200 = key valid)');
}


/* ---------- Unggah gambar highlight & kolom tambahan (admin) ---------- */
function handleUpload_(data) {
  if (!requireAuth_(data.token)) return fail_('Sesi habis. Silakan login ulang.');
  const b64 = String(data.image || '');
  if (!b64 || b64.length > 900000) return fail_('Gambar kosong atau terlalu besar.');
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length > 700000 || (bytes[0] & 255) !== 0xFF || (bytes[1] & 255) !== 0xD8) return fail_('Hanya gambar JPEG (maks 700 KB).');
  const P = PropertiesService.getScriptProperties();
  let folder;
  try { folder = DriveApp.getFolderById(P.getProperty('IMG_FOLDER')); } catch (e) { folder = DriveApp.createFolder('Katalog Hermina - Highlight'); P.setProperty('IMG_FOLDER', folder.getId()); }
  const f = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', 'highlight-' + Date.now() + '.jpg'));
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return json_({ success: true, url: 'https://lh3.googleusercontent.com/d/' + f.getId() });
}

function handleColumn_(data) {
  if (!requireAuth_(data.token)) return fail_('Sesi habis. Silakan login ulang.');
  const n = String(data.sheet || ''), name = String(data.name || '').trim();
  if (['Promos', 'MCU', 'Packages', 'Contacts'].indexOf(n) < 0) return fail_('Sheet ini tidak mendukung kolom tambahan.');
  if (!/^[A-Za-z][A-Za-z0-9_]{0,29}$/.test(name)) return fail_('Nama kolom: huruf/angka/_ (maks 30), diawali huruf.');
  if (SCHEMA[n].indexOf(name) >= 0) return fail_('Kolom inti tidak boleh diubah.');
  const lock = LockService.getScriptLock(); lock.waitLock(15000);
  try {
    const sh = getSheet_(n), hdr = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0], pos = hdr.indexOf(name);
    if (data.op === 'add') { if (pos >= 0) return fail_('Kolom sudah ada.'); sh.getRange(1, hdr.length + 1).setValue(name); }
    else if (data.op === 'delete') { if (pos < 0) return fail_('Kolom tidak ditemukan.'); sh.deleteColumn(pos + 1); }
    else return fail_('Aksi tidak dikenali.');
    return json_({ success: true });
  } finally { lock.releaseLock(); }
}


/* ================= TARIF, TANYA PASIEN & NOTIFIKASI ================= */
/*
 * Sheet Tarif (admin: tab "Tarif Harga") = database harga untuk pertanyaan pasien. Jawaban TIDAK dikarang AI:
 * dicocokkan ke sheet Tarif + katalog (MCU, Paket, Promo). Bila tidak ada yang cocok, website mengarahkan
 * pasien ke 5 petugas via WhatsApp, dan pertanyaan dicatat di sheet Leads agar bisa ditambahkan ke Tarif.
 * Sheet Notifs = pemberitahuan (lonceng + notifikasi browser). Terisi otomatis saat admin menambah/mengubah
 * Promo, MCU, Paket, Tarif; juga bisa diisi manual lewat tab "Notifikasi".
 */
const NOTIFY_LABEL = { Promos: 'Promo', MCU: 'Paket MCU', Packages: 'Paket layanan', Tarif: 'Tarif' };
const STOP_ = 'yang dan atau di ke dari untuk dengan apa berapa harga harganya tarif biaya ada nya itu ini saya mau tanya bagaimana gimana kah tolong info informasi paket layanan rumah sakit hermina balikpapan pak bu mas mba mbak'.split(' ');

function notifyChange_(sheet, isNew, row, hdr) {
  if (!NOTIFY_LABEL[sheet]) return;
  const o = {}; hdr.forEach((h, i) => { o[h] = row[i]; });
  const name = String(o.name || o.title || '').slice(0, 80);
  if (!name || (o.isActive !== undefined && !toBool_(o.isActive))) return;
  const c = CacheService.getScriptCache(), k = 'nt_' + sheet + '_' + name.replace(/\W/g, '').slice(0, 40);
  if (!isNew && c.get(k)) return; // ubahan beruntun pada item yang sama dijadikan satu pemberitahuan
  c.put(k, '1', 600);
  getSheet_('Notifs').appendRow(['NTF-' + Date.now(), (isNew ? 'Baru: ' : 'Diperbarui: ') + name, NOTIFY_LABEL[sheet], now_(), true]);
}

function tok_(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter(w => w.length > 2 && STOP_.indexOf(w) < 0); }

function handleAsk_(data) {
  const q = clean_(data.q, 200);
  if (q.length < 3) return fail_('Tulis pertanyaan Anda.');
  const qt = tok_(q), items = [];
  getData_('Tarif').filter(t => toBool_(t.isActive) && t.name).forEach(t =>
    items.push({ type: 'Tarif', name: String(t.name), cat: String(t.category || ''), price: Number(t.price) || 0, note: String(t.note || ''), b: 0.5 }));
  catalogForAi_().forEach(c => items.push({ type: c.type, name: c.name, cat: c.category, price: c.price, note: c.desc, b: 0 }));
  const need = Math.max(1, Math.ceil(qt.length / 2));
  const hits = qt.length ? items.map(it => {
    const h1 = (it.name + ' ' + it.cat).toLowerCase(), h2 = it.note.toLowerCase();
    let m = 0, sc = it.b; qt.forEach(w => { if (h1.indexOf(w) >= 0) { m++; sc += 2; } else if (h2.indexOf(w) >= 0) { m++; sc += 1; } });
    return { it: it, m: m, sc: sc };
  }).filter(x => x.m >= need).sort((a, b) => b.sc - a.sc).slice(0, 4).map(x => ({ type: x.it.type, name: x.it.name, price: x.it.price, note: x.it.note.slice(0, 220) })) : [];
  if (!hits.length) { // catat pertanyaan yang belum terjawab (maks. 1x/menit per pertanyaan)
    const c = CacheService.getScriptCache(), k = 'ask_' + q.toLowerCase().replace(/\W/g, '').slice(0, 40);
    if (!c.get(k)) { c.put(k, '1', 60); try { getSheet_('Leads').appendRow([now_(), '-', '-', 'Pertanyaan belum terjawab', q, 'Pertanyaan', '', '', '', '']); } catch (e) {} }
  }
  return json_({ success: true, found: hits.length > 0, data: hits });
}
