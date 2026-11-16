/**
 * CEK RAK - Backend (Google Apps Script)
 * Pasang di: Google Sheet > Extensions > Apps Script
 * Langkah: jalankan setupSheets() sekali, lalu Deploy > New deployment > Web app.
 */

const TZ = Session.getScriptTimeZone();
const HEADERS = {
  master:   ['sku', 'nama_barang', 'lokasi', 'qty_system'],
  sesi:     ['sesi_id', 'nama_sesi', 'status', 'dibuat_oleh', 'dibuat_pada', 'mode'],
  scan_log: ['scan_id', 'sesi_id', 'sku', 'qty', 'lokasi', 'petugas', 'waktu', 'status', 'tipe'],
  users:    ['nama', 'pin', 'role'],
  recycle_bin: ['bin_id', 'dihapus_pada', 'dihapus_oleh', 'scan_id', 'sesi_id', 'sku', 'qty', 'lokasi', 'petugas', 'waktu', 'status', 'tipe']
};
// Kolom yang harus disimpan sebagai teks (supaya barcode & PIN tidak berubah jadi angka)
const TEXT_COLS = { master: [1], sesi: [1, 5], scan_log: [1, 2, 3, 7], users: [2], recycle_bin: [1, 2, 4, 5, 6, 10] };

/* ---------- Setup awal ---------- */
function setupSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(HEADERS).forEach(function (name) {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const h = HEADERS[name];
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#FFD43B');
    sh.setFrozenRows(1);
    TEXT_COLS[name].forEach(function (c) {
      sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  });
  const users = ss.getSheetByName('users');
  if (users.getLastRow() < 2) {
    users.getRange(2, 1, 2, 3).setValues([['Admin', '1234', 'admin'], ['Petugas 1', '1111', 'petugas']]);
  }
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  SpreadsheetApp.getUi().alert('Selesai. Ganti PIN contoh di sheet "users" sebelum dipakai.');
}

/* ---------- Pintu masuk API ---------- */
function doGet() { return json_({ ok: true, data: 'Cek Rak API aktif' }); }

function doPost(e) {
  try {
    const req = JSON.parse(e.postData.contents);
    const user = auth_(req.pin);
    const fn = ACTIONS[req.action];
    if (!fn) throw new Error('Aksi tidak dikenal: ' + req.action);
    return json_({ ok: true, data: fn(req, user) });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- Helper ---------- */
const SH_MEMO = {};
function sh_(name) {
  if (SH_MEMO[name]) return SH_MEMO[name];
  const s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!s) throw new Error('Sheet "' + name + '" belum ada. Jalankan setupSheets() dulu.');
  return (SH_MEMO[name] = s);
}
function rows_(name) {
  const s = sh_(name), n = s.getLastRow() - 1;
  return n > 0 ? s.getRange(2, 1, n, HEADERS[name].length).getValues() : [];
}
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }

// Daftar user disimpan di cache 5 menit supaya tiap scan tidak membaca sheet users.
function usersMap_(fresh) {
  const c = CacheService.getScriptCache();
  if (!fresh) { const hit = c.get('users_v1'); if (hit) return JSON.parse(hit); }
  const map = {};
  rows_('users').forEach(function (r) {
    const p = String(r[1]).trim();
    if (p) map[p] = { nama: String(r[0]), role: String(r[2]).toLowerCase() === 'admin' ? 'admin' : 'petugas' };
  });
  c.put('users_v1', JSON.stringify(map), 300);
  return map;
}
function auth_(pin) {
  pin = String(pin || '').trim();
  if (!pin) throw new Error('PIN kosong');
  let u = usersMap_(false)[pin];
  if (!u) u = usersMap_(true)[pin];          // mungkin user baru ditambahkan: baca ulang sekali
  if (!u) throw new Error('PIN salah');
  return u;
}
function requireAdmin_(user) { if (user.role !== 'admin') throw new Error('Hanya admin yang boleh melakukan ini'); }

function getSesi_(id) {
  const rows = rows_('sesi');
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === String(id)) {
      return { id: String(rows[i][0]), nama: String(rows[i][1]), status: String(rows[i][2]), row: i + 2,
               mode: String(rows[i][5] || '').toLowerCase() === 'itemize' ? 'itemize' : 'opname' };
    }
  }
  return null;
}

// Cari SKU di master (jumlahkan qty jika SKU muncul di beberapa baris/lokasi)
function findMaster_(sku) {
  const s = sh_('master'), n = s.getLastRow() - 1;
  if (n < 1) return null;
  const hits = s.getRange(2, 1, n, 1).createTextFinder(sku).matchEntireCell(true).findAll();
  if (!hits.length) return null;
  let nama = '', qty = 0;
  const lok = [];
  hits.forEach(function (h) {
    const r = s.getRange(h.getRow(), 1, 1, 4).getValues()[0];
    if (!nama) nama = String(r[1]);
    const l = String(r[2]).trim();
    if (l && lok.indexOf(l) < 0) lok.push(l);
    qty += Number(r[3]) || 0;
  });
  return { sku: sku, nama: nama, qty_system: qty, lokasi: lok.join(', ') };
}

/* ---------- Cache master per SKU (mempercepat scan beruntun) ----------
   Hasil pencarian SKU di sheet master disimpan di cache 6 jam. Versi master (MVER) dinaikkan
   setiap upload data system, sehingga cache lama otomatis tidak terpakai. */
function masterVer_() { return PropertiesService.getScriptProperties().getProperty('MVER') || '0'; }
function bumpMasterVer_() { PropertiesService.getScriptProperties().setProperty('MVER', String(Date.now())); }
function findMasterMany_(skus) {
  const c = CacheService.getScriptCache(), ver = masterVer_(), out = {}, put = {};
  const keyOf = function (s) { return ('mk' + ver + '_' + s).slice(0, 240); };
  const set = {}; skus.forEach(function (s) { set[s] = 1; });
  const uniq = Object.keys(set);
  const got = uniq.length ? c.getAll(uniq.map(keyOf)) : {};
  uniq.forEach(function (s) {
    const hit = got[keyOf(s)];
    if (hit != null) { out[s] = hit === '-' ? null : JSON.parse(hit); return; }
    const m = findMaster_(s);
    out[s] = m; put[keyOf(s)] = m ? JSON.stringify(m) : '-';
  });
  try { if (Object.keys(put).length) c.putAll(put, 21600); } catch (e) {}
  return out;
}

// Info sesi dari cache (5 menit); dihapus saat sesi ditutup
function getSesiFast_(id) {
  const c = CacheService.getScriptCache(), k = 'sesi_' + id, hit = c.get(k);
  if (hit) return JSON.parse(hit);
  const s = getSesi_(id);
  if (s) c.put(k, JSON.stringify(s), 300);
  return s;
}
function markDirty_(sesiId) { CacheService.getScriptCache().put('dirty_' + sesiId, '1', 21600); }

// Cari baris scan_log milik satu SKU di satu sesi tanpa membaca seluruh sheet (pakai TextFinder)
function skuLog_(log, sesiId, sku) {
  const n = log.getLastRow() - 1;
  if (n < 1) return [];
  const out = [];
  log.getRange(2, 3, n, 1).createTextFinder(sku).matchEntireCell(true).findAll().forEach(function (h) {
    const row = h.getRow(), v = log.getRange(row, 1, 1, 9).getValues()[0];
    if (String(v[1]) === sesiId && String(v[2]) === sku) out.push({ row: row, v: v });
  });
  return out;
}
function sumFisik_(sesiId, sku) {
  return skuLog_(sh_('scan_log'), sesiId, sku).reduce(function (t, x) { return x.v[7] === 'valid' ? t + (Number(x.v[3]) || 0) : t; }, 0);
}
function adaValid_(sesiId, sku) {
  return skuLog_(sh_('scan_log'), sesiId, sku).some(function (x) { return x.v[7] === 'valid'; });
}

function isDuplicate_(sheet, scanId) {
  const n = sheet.getLastRow() - 1;
  if (n < 1) return false;
  const from = Math.max(2, sheet.getLastRow() - 299);
  const ids = sheet.getRange(from, 1, sheet.getLastRow() - from + 1, 1).getValues();
  return ids.some(function (r) { return String(r[0]) === scanId; });
}

/* ---------- Skema lama (sheet sesi belum punya kolom mode) ---------- */
function ensureSchema_() {
  const sesi = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('sesi');
  if (sesi && sesi.getRange(1, 6).getValue() === '') {
    sesi.getRange(1, 6).setValue('mode').setFontWeight('bold').setBackground('#FFD43B');
  }
}

/* ---------- Tulis baris ke sheet (dengan format teks agar barcode tidak jadi angka) ---------- */
function appendRows_(sh, ncol, rows, textCols) {
  if (!rows.length) return;
  const start = Math.max(sh.getLastRow(), 1) + 1;
  if (sh.getMaxRows() < start + rows.length - 1) sh.insertRowsAfter(sh.getMaxRows(), start + rows.length - 1 - sh.getMaxRows());
  textCols.forEach(function (c) { sh.getRange(start, c, rows.length, 1).setNumberFormat('@'); });
  sh.getRange(start, 1, rows.length, ncol).setValues(rows);
}
function rewrite_(sh, ncol, rows, textCols) {
  const n = sh.getLastRow() - 1;
  if (n > 0) sh.getRange(2, 1, n, ncol).clearContent();
  appendRows_(sh, ncol, rows, textCols);
}

/* ---------- Recycle bin ---------- */
function binSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName('recycle_bin');
  if (!sh) {
    sh = ss.insertSheet('recycle_bin');
    const h = HEADERS.recycle_bin;
    sh.getRange(1, 1, 1, h.length).setValues([h]).setFontWeight('bold').setBackground('#FFD43B');
    sh.setFrozenRows(1);
    TEXT_COLS.recycle_bin.forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  }
  return sh;
}
function binRows_() {
  const s = binSheet_(), n = s.getLastRow() - 1;
  return n > 0 ? s.getRange(2, 1, n, 12).getValues() : [];
}
function uniqValid_(rows, statusIdx, skuIdx) {
  const set = {};
  rows.forEach(function (r) { if (r[statusIdx] === 'valid') set[String(r[skuIdx])] = 1; });
  return Object.keys(set).length;
}
function binId_() { return 'RST-' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd-HHmmss'); }

/* ---------- Sheet hasil scan (otomatis, satu sheet per sesi) ---------- */
const HASIL_HEADERS = {
  opname:  ['sku', 'nama_barang', 'lokasi', 'qty_system', 'qty_fisik', 'selisih', 'status', 'petugas', 'waktu_scan'],
  itemize: ['sku', 'nama_barang', 'status', 'petugas', 'waktu_scan']
};
const HASIL_STATUS_COL = { opname: 7, itemize: 3 };
const STATUS_WARNA = {
  'Sesuai': '#D9F2E6', 'Sudah di-scan': '#D9F2E6', 'Kurang': '#F9D6D3', 'Lebih': '#FCE8C3',
  'Belum di-scan': '#E6E9ED', 'Tidak ada di system': '#E4DBF5'
};

// Sesi baru: ID sesi = nama output (mis. STOCK TAKE_JW5006_CGAR_20261114_153045). Sesi lama (OPN-/ITM-) tetap "Hasil <ID>".
function hasilName_(sesiId) {
  if (/^(OPN|ITM)-/.test(sesiId)) return 'Hasil ' + sesiId;
  return String(sesiId).replace(/[\\\/?*\[\]:]/g, '').slice(0, 100);
}

/* ---------- Data toko (kode & nama) ---------- */
function cleanPart_(v, max) { return String(v || '').replace(/[\\\/\?\*\[\]:'"]/g, '').replace(/\s+/g, ' ').trim().slice(0, max); }
function getToko_() {
  const p = PropertiesService.getScriptProperties();
  return { kode: p.getProperty('TOKO_KODE') || '', nama: p.getProperty('TOKO_NAMA') || '' };
}

/* ---------- Daftar SKU master (cache 6 jam, dihapus saat upload master) ---------- */
function masterSkus_() {
  const c = CacheService.getScriptCache(), cnt = c.get('msk_n');
  if (cnt !== null) {
    const keys = [];
    for (let i = 0; i < Number(cnt); i++) keys.push('msk_' + i);
    const got = keys.length ? c.getAll(keys) : {};
    if (keys.every(function (k) { return got[k] != null; })) {
      const set = {};
      keys.forEach(function (k) { got[k].split('\n').forEach(function (x) { if (x) set[x] = 1; }); });
      return set;
    }
  }
  const sh = sh_('master'), n = sh.getLastRow() - 1, set = {};
  if (n > 0) sh.getRange(2, 1, n, 1).getValues().forEach(function (r) { const k = String(r[0]).trim(); if (k) set[k] = 1; });
  const shards = []; let cur = [], len = 0;
  Object.keys(set).forEach(function (k) {
    if (len + k.length + 1 > 90000) { shards.push(cur.join('\n')); cur = []; len = 0; }
    cur.push(k); len += k.length + 1;
  });
  if (cur.length) shards.push(cur.join('\n'));
  try {
    const obj = { msk_n: String(shards.length) };
    shards.forEach(function (x, i) { obj['msk_' + i] = x; });
    c.putAll(obj, 21600);
  } catch (e) {}
  return set;
}
function modeOf_(sesi) { return sesi.mode === 'itemize' ? 'itemize' : 'opname'; }

function statusOf_(mode, inMaster, scanned, sys, fisik) {
  if (!inMaster) return 'Tidak ada di system';
  if (mode === 'itemize') return scanned ? 'Sudah di-scan' : 'Belum di-scan';
  if (!scanned) return 'Belum di-scan';
  const d = fisik - sys;
  return d === 0 ? 'Sesuai' : d < 0 ? 'Kurang' : 'Lebih';
}

function hasilRow_(mode, r) {
  return mode === 'itemize'
    ? [r.sku, r.nama, r.status, r.petugas, r.waktu]
    : [r.sku, r.nama, r.lokasi, r.qty_system, r.qty_fisik, r.selisih, r.status, r.petugas, r.waktu];
}

// Seluruh baris hasil: semua SKU master (sudah/belum discan) + SKU yang discan tapi tidak ada di master
function rekapFull_(sesi) {
  const mode = modeOf_(sesi);
  const master = {}, order = [];
  rows_('master').forEach(function (r) {
    const sku = String(r[0]).trim();
    if (!sku) return;
    if (!master[sku]) { master[sku] = { nama: String(r[1]), lok: [], sys: 0 }; order.push(sku); }
    const l = String(r[2]).trim();
    if (l && master[sku].lok.indexOf(l) < 0) master[sku].lok.push(l);
    master[sku].sys += Number(r[3]) || 0;
  });
  const fisik = {}, last = {};
  rows_('scan_log').forEach(function (r) {
    if (String(r[1]) === sesi.id && r[7] === 'valid') {
      const sku = String(r[2]);
      fisik[sku] = (fisik[sku] || 0) + (Number(r[3]) || 0);
      last[sku] = { lokasi: String(r[4]), petugas: String(r[5]), waktu: String(r[6]) };
    }
  });
  Object.keys(fisik).forEach(function (sku) { if (!master[sku]) order.push(sku); });

  return order.map(function (sku) {
    const m = master[sku], scanned = (sku in fisik), f = fisik[sku] || 0, sys = m ? m.sys : 0, l = last[sku];
    return {
      sku: sku,
      nama: m ? m.nama : '(Tidak ada di system)',
      lokasi: m ? m.lok.join(', ') : (l ? l.lokasi : ''),
      qty_system: sys, qty_fisik: f, selisih: f - sys,
      status: statusOf_(mode, !!m, scanned, sys, f),
      petugas: l ? l.petugas : '', waktu: l ? l.waktu : ''
    };
  });
}

// Tulis ulang seluruh sheet hasil sesi ini
function writeHasil_(sesi) {
  const dl = LockService.getDocumentLock();     // kunci terpisah dari kunci scan, jadi tidak menghambat scan
  dl.waitLock(30000);
  try { return writeHasilNow_(sesi); } finally { dl.releaseLock(); }
}
function writeHasilNow_(sesi) {
  const mode = modeOf_(sesi), head = HASIL_HEADERS[mode], nc = head.length, sc = HASIL_STATUS_COL[mode];
  const ss = SpreadsheetApp.getActiveSpreadsheet(), name = hasilName_(sesi.id);
  const rows = rekapFull_(sesi);
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.clear();
  const n = rows.length;
  if (sh.getMaxRows() < n + 1) sh.insertRowsAfter(sh.getMaxRows(), n + 1 - sh.getMaxRows());
  sh.getRange(1, 1, sh.getMaxRows(), 1).setNumberFormat('@');
  sh.getRange(1, nc, sh.getMaxRows(), 1).setNumberFormat('@');
  sh.getRange(1, 1, 1, nc).setValues([head]).setFontWeight('bold').setBackground('#FFD43B');
  sh.setFrozenRows(1);
  if (n) {
    sh.getRange(2, 1, n, nc).setValues(rows.map(function (r) { return hasilRow_(mode, r); }));
    sh.getRange(2, sc, n, 1).setBackgrounds(rows.map(function (r) { return [STATUS_WARNA[r.status] || '#FFFFFF']; }));
  }
  return { sheet: name, baris: n };
}

/* ---------- Aksi ---------- */
const ACTIONS = {
  login: function (req, user) { return user; },

  listSesi: function () {
    return rows_('sesi').map(function (r) {
      return { sesi_id: String(r[0]), nama_sesi: String(r[1]), status: String(r[2]), mode: String(r[5] || '').toLowerCase() === 'itemize' ? 'itemize' : 'opname' };
    }).reverse();
  },

  createSesi: function (req, user) {
    requireAdmin_(user);
    const nama = String(req.nama || '').trim();
    if (!nama) throw new Error('Nama sesi wajib diisi');
    const mode = String(req.mode || '').toLowerCase() === 'itemize' ? 'itemize' : 'opname';
    ensureSchema_();
    const toko = getToko_();
    if (!toko.kode || !toko.nama) throw new Error('Isi kode toko dan nama toko dulu di menu Admin');
    const id = (mode === 'itemize' ? 'ITEMIZE' : 'STOCK TAKE') + '_' + toko.kode + '_' + toko.nama + '_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd_HHmmss');
    sh_('sesi').appendRow([id, nama, 'aktif', user.nama, now_(), mode]);
    try { writeHasil_({ id: id, nama: nama, status: 'aktif', mode: mode }); } catch (err) { console.error('writeHasil_: ' + err); }
    return { sesi_id: id, sheet: hasilName_(id), mode: mode };
  },

  closeSesi: function (req, user) {
    requireAdmin_(user);
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    sh_('sesi').getRange(sesi.row, 3).setValue('selesai');
    CacheService.getScriptCache().remove('sesi_' + sesi.id);
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try { writeHasil_(sesi); } finally { lock.releaseLock(); }   // hasil akhir ditulis ulang penuh
    return { ok: true };
  },

  // Tulis ulang penuh sheet hasil (mis. setelah data system diupload ulang)
  // Tulis ulang sheet hasil hanya jika ada scan baru sejak sinkron terakhir (atau force)
  syncHasil: function (req) {
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    const c = CacheService.getScriptCache(), k = 'dirty_' + sesi.id;
    if (!req.force && !c.get(k)) return { dilewati: true };
    c.remove(k);
    return writeHasil_(sesi);
  },

  uploadMaster: function (req, user) {
    requireAdmin_(user);
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const s = sh_('master');
      if (!req.append) {
        const n = s.getLastRow() - 1;
        if (n > 0) s.getRange(2, 1, n, 4).clearContent();
      }
      const data = (req.rows || []).map(function (r) {
        return [String(r[0] || '').trim(), String(r[1] || '').trim(), String(r[2] || '').trim(), Number(r[3]) || 0];
      }).filter(function (r) { return r[0]; });
      if (data.length) s.getRange(Math.max(s.getLastRow(), 1) + 1, 1, data.length, 4).setValues(data);
      CacheService.getScriptCache().remove('msk_n');
      bumpMasterVer_();                              // cache master per SKU jadi kedaluwarsa
      return { masuk: data.length };
    } finally { lock.releaseLock(); }
  },

  scan: function (req, user) {
    const sku = String(req.sku || '').trim();
    if (!sku) throw new Error('SKU kosong');
    const sesi = getSesiFast_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    if (sesi.status !== 'aktif') throw new Error('Sesi sudah selesai, scan tidak bisa ditambah');
    const item = sesi.mode === 'itemize';
    const m = findMaster_(sku);                     // baca master di luar kunci supaya kunci cepat dilepas

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const log = sh_('scan_log');
      const id = String(req.scan_id || Utilities.getUuid());
      let baru = isDuplicate_(log, id), total;      // kiriman ulang dari antrean offline
      if (baru) total = sumFisik_(sesi.id, sku);
      else {
        const valid = skuLog_(log, sesi.id, sku).filter(function (x) { return x.v[7] === 'valid'; });
        if (valid.length) total = valid.reduce(function (t, x) { return t + (Number(x.v[3]) || 0); }, 0);
        else {                                       // SKU baru di sesi ini: simpan otomatis
          const qty = item ? 1 : (m ? m.qty_system : 1);
          log.appendRow([id, sesi.id, sku, qty, item ? '-' : String(req.lokasi || '-'), user.nama, now_(), 'valid', 'scan']);
          baru = true; total = qty;
        }
      }
      if (baru) markDirty_(sesi.id);                 // sheet hasil diperbarui terpisah (syncHasil), tidak menahan scan
      return {
        sku: sku,
        nama: m ? m.nama : '(Tidak ada di system)',
        ada_di_master: !!m,
        qty_system: m ? m.qty_system : 0,
        total_fisik: total,
        baru: baru,
        mode: sesi.mode
      };
    } finally { lock.releaseLock(); }
  },

  // Kirim banyak scan sekaligus (scan beruntun): satu kunci, satu pembacaan scan_log, satu penulisan.
  // Jawaban berisi satu hasil per scan, dengan bentuk sama seperti aksi "scan".
  scanBatch: function (req, user) {
    const items = (req.items || []).slice(0, 40), results = [];
    if (!items.length) return { results: results };
    const sesi = getSesiFast_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    if (sesi.status !== 'aktif') throw new Error('Sesi sudah selesai, scan tidak bisa ditambah');
    const item = sesi.mode === 'itemize';
    const skus = items.map(function (it) { return String(it.sku || '').trim(); }).filter(function (s) { return s; });
    const mm = findMasterMany_(skus);                // dari cache; di luar kunci

    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const log = sh_('scan_log'), n = log.getLastRow() - 1, ids = {}, tot = {}, have = {};
      if (n > 0) log.getRange(2, 1, n, 8).getValues().forEach(function (r) {
        if (String(r[1]) !== sesi.id) return;
        ids[String(r[0])] = 1;
        if (r[7] === 'valid') { const k = String(r[2]); have[k] = 1; tot[k] = (tot[k] || 0) + (Number(r[3]) || 0); }
      });
      const fresh = [], w = now_();
      let dirty = false;
      items.forEach(function (it) {
        const sku = String(it.sku || '').trim(), sid = String(it.scan_id || '');
        if (!sku) { results.push({ scan_id: sid, error: 'SKU kosong' }); return; }
        const m = mm[sku], id = sid || Utilities.getUuid();
        let baru, total;
        if (ids[id]) { baru = true; total = tot[sku] || 0; }             // kiriman ulang dari antrean offline
        else if (have[sku]) { baru = false; total = tot[sku] || 0; }      // sudah pernah discan di sesi ini
        else {
          const qty = item ? 1 : (m ? m.qty_system : 1);
          fresh.push([id, sesi.id, sku, qty, item ? '-' : String(it.lokasi || '-'), user.nama, w, 'valid', 'scan']);
          ids[id] = 1; have[sku] = 1; tot[sku] = qty; baru = true; total = qty;
        }
        if (baru) dirty = true;
        results.push({ scan_id: sid, data: {
          sku: sku, nama: m ? m.nama : '(Tidak ada di system)', ada_di_master: !!m,
          qty_system: m ? m.qty_system : 0, total_fisik: total, baru: baru, mode: sesi.mode
        } });
      });
      appendRows_(log, 9, fresh, TEXT_COLS.scan_log);
      if (dirty) markDirty_(sesi.id);
      return { results: results };
    } finally { lock.releaseLock(); }
  },

  // Progres + daftar SKU yang sudah discan (semua petugas). Dipakai untuk menyegarkan layar otomatis
  // supaya scan petugas lain langsung terlihat. Jika "sig" sama dengan keadaan terakhir, daftar tidak dikirim ulang.
  sync: function (req) {
    const sesi = getSesiFast_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    const master = masterSkus_(), total = Object.keys(master).length;
    const log = sh_('scan_log'), n = log.getLastRow() - 1, seen = {};
    if (n > 0) log.getRange(2, 2, n, 7).getValues().forEach(function (r) {
      if (String(r[0]) === sesi.id && r[6] === 'valid') seen[String(r[1])] = 1;
    });
    const list = Object.keys(seen);
    let sudah = 0, luar = 0;
    list.forEach(function (k) { if (master[k]) sudah++; else luar++; });
    const progress = { mode: sesi.mode, total: total, sudah: sudah, belum: Math.max(total - sudah, 0), luar: luar,
                       persen: total ? Math.round(sudah / total * 1000) / 10 : 0 };
    const sig = log.getLastRow() + ':' + list.length + ':' + sudah;
    if (req.sig && String(req.sig) === sig) return { same: true, sig: sig, progress: progress };
    return { same: false, sig: sig, progress: progress, scanned: list };
  },

  undo: function (req, user) {
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    if (sesi.status !== 'aktif') throw new Error('Sesi sudah selesai');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const log = sh_('scan_log'), n = log.getLastRow() - 1;
      if (n < 1) throw new Error('Belum ada scan');
      const vals = log.getRange(2, 1, n, 9).getValues();
      // Pilih baris yang dibatalkan: utamakan scan_id dari klien (scan teratas di riwayat),
      // lalu SKU terbaru milik petugas, terakhir baris terbawah milik petugas (perilaku lama).
      const wantId = String(req.scan_id || ''), wantSku = String(req.sku || '');
      const mine = function (r) { return String(r[1]) === sesi.id && String(r[5]) === user.nama && r[7] === 'valid'; };
      let pick = -1;
      if (wantId) for (let i = n - 1; i >= 0; i--) { if (String(vals[i][0]) === wantId && mine(vals[i])) { pick = i; break; } }
      if (pick < 0 && wantSku) for (let i = n - 1; i >= 0; i--) { if (String(vals[i][2]) === wantSku && mine(vals[i])) { pick = i; break; } }
      if (pick < 0 && !wantId && !wantSku) for (let i = n - 1; i >= 0; i--) { if (mine(vals[i])) { pick = i; break; } }
      if (pick >= 0) {
        const r = vals[pick];
        log.getRange(pick + 2, 8).setValue('dibatalkan');
        const sku = String(r[2]), m = findMaster_(sku);
        markDirty_(sesi.id);
        return {
          sku: sku, qty: Number(r[3]),
          nama: m ? m.nama : '(Tidak ada di system)',
          qty_system: m ? m.qty_system : 0,
          total_fisik: sumFisik_(sesi.id, sku),
          mode: sesi.mode
        };
      }
      throw new Error('Tidak ada scan Anda yang bisa dibatalkan di sesi ini');
    } finally { lock.releaseLock(); }
  },

  // Ubah qty fisik satu SKU: baris lama ditandai "diganti", baris baru dicatat (riwayat tetap ada)
  setQty: function (req, user) {
    const sku = String(req.sku || '').trim();
    const qty = Math.round(Number(req.qty));
    if (!sku) throw new Error('SKU kosong');
    if (isNaN(qty) || qty < 0 || qty > 9999) throw new Error('Qty harus 0 - 9999');
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    if (sesi.mode === 'itemize') throw new Error('Mode itemize tidak memakai qty');
    if (sesi.status !== 'aktif') throw new Error('Sesi sudah selesai');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const log = sh_('scan_log'), n = log.getLastRow() - 1;
      if (n > 0) log.getRange(2, 1, n, 9).getValues().forEach(function (r, i) {
        if (String(r[1]) === sesi.id && String(r[2]) === sku && r[7] === 'valid') log.getRange(i + 2, 8).setValue('diganti');
      });
      log.appendRow([Utilities.getUuid(), sesi.id, sku, qty, String(req.lokasi || '-'), user.nama, now_(), 'valid', 'edit']);
      markDirty_(sesi.id);
      const m = findMaster_(sku);
      return { sku: sku, nama: m ? m.nama : '(Tidak ada di system)', ada_di_master: !!m, qty_system: m ? m.qty_system : 0, total_fisik: qty };
    } finally { lock.releaseLock(); }
  },

  // Daftar SKU yang sudah tercatat di sesi (untuk menu Edit qty), terbaru di atas
  listCounted: function (req) {
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    const master = {};
    rows_('master').forEach(function (r) {
      const k = String(r[0]).trim();
      if (!k) return;
      if (!master[k]) master[k] = { nama: String(r[1]), sys: 0 };
      master[k].sys += Number(r[3]) || 0;
    });
    const agg = {};
    rows_('scan_log').forEach(function (r, i) {
      if (String(r[1]) !== sesi.id || r[7] !== 'valid') return;
      const sku = String(r[2]), m = master[sku];
      if (!agg[sku]) agg[sku] = { sku: sku, nama: m ? m.nama : '(Tidak ada di system)', qty_system: m ? m.sys : 0, qty_fisik: 0, idx: 0 };
      agg[sku].qty_fisik += Number(r[3]) || 0;
      agg[sku].idx = i;
    });
    const rows = Object.keys(agg).map(function (k) { return agg[k]; }).sort(function (a, b) { return b.idx - a.idx; });
    return { rows: rows };
  },

  rekap: function (req) {
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    return {
      sesi: { sesi_id: sesi.id, nama_sesi: sesi.nama, status: sesi.status },
      mode: sesi.mode,
      sheet: hasilName_(sesi.id),
      rows: rekapFull_(sesi)
    };
  },

  // Hapus sesi permanen: baris sesi, semua scan_log sesi, isi recycle_bin sesi, dan sheet hasilnya
  deleteSesi: function (req, user) {
    requireAdmin_(user);
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const log = sh_('scan_log'), n = log.getLastRow() - 1;
      const all = n > 0 ? log.getRange(2, 1, n, 9).getValues() : [];
      const rest = all.filter(function (r) { return String(r[1]) !== sesi.id; });
      const hapusScan = all.length - rest.length;
      if (hapusScan) rewrite_(log, 9, rest, TEXT_COLS.scan_log);

      const bin = binSheet_(), brows = binRows_();
      const keep = brows.filter(function (r) { return String(r[4]) !== sesi.id; });
      if (keep.length !== brows.length) rewrite_(bin, 12, keep, TEXT_COLS.recycle_bin);

      const ss = SpreadsheetApp.getActiveSpreadsheet(), name = hasilName_(sesi.id);
      const sistem = ['master', 'sesi', 'scan_log', 'users', 'recycle_bin'];
      const hs = ss.getSheetByName(name);
      let sheetDihapus = false;
      if (hs && sistem.indexOf(name) < 0 && ss.getSheets().length > 1) { ss.deleteSheet(hs); sheetDihapus = true; }

      const again = getSesi_(sesi.id);                       // baris sesi dihapus terakhir
      if (again) sh_('sesi').deleteRow(again.row);
      const c = CacheService.getScriptCache();
      c.remove('sesi_' + sesi.id); c.remove('dirty_' + sesi.id);
      return { sesi_id: sesi.id, scan: hapusScan, sheet: sheetDihapus ? name : '' };
    } finally { lock.releaseLock(); }
  },

  // Import hasil scan dari file teks: setiap SKU dicatat seperti hasil scan (tipe 'import'), SKU yang sudah ada di sesi dilewati
  importScans: function (req, user) {
    requireAdmin_(user);
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    if (sesi.status !== 'aktif') throw new Error('Sesi sudah selesai, tidak bisa import');
    const item = sesi.mode === 'itemize';
    const rows = (req.rows || []).slice(0, 500);
    if (!rows.length) return { masuk: 0, dilewati: [], dibuat: [] };
    const master = {};
    rows_('master').forEach(function (r) {
      const k = String(r[0]).trim();
      if (!k) return;
      if (!master[k]) master[k] = { nama: String(r[1]), sys: 0 };
      master[k].sys += Number(r[3]) || 0;
    });
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const log = sh_('scan_log'), n = log.getLastRow() - 1;
      const have = {}, ids = {};
      if (n > 0) log.getRange(2, 1, n, 8).getValues().forEach(function (r) {
        if (String(r[1]) !== sesi.id) return;
        ids[String(r[0])] = 1;
        if (r[7] === 'valid') have[String(r[2])] = 1;
      });
      const out = [], skip = [], dibuat = [], w = now_();
      rows.forEach(function (r) {
        const sku = String(r.sku || '').trim();
        if (!sku) return;
        const id = String(r.scan_id || Utilities.getUuid());
        const m = master[sku];
        const q = (r.qty === null || r.qty === undefined || r.qty === '') ? null : Math.round(Number(r.qty));
        const qty = item ? 1 : (q === null ? (m ? m.sys : 1) : q);
        if (isNaN(qty) || qty < 0 || qty > 9999) { skip.push(sku); return; }
        const info = { scan_id: id, sku: sku, nama: m ? m.nama : '(Tidak ada di system)', ada: !!m, qty: qty };
        if (ids[id]) { dibuat.push(info); return; }          // kiriman ulang dari klien: sudah tercatat
        if (have[sku]) { skip.push(sku); return; }
        out.push([id, sesi.id, sku, qty, item ? '-' : (String(r.lokasi || '').trim() || '-'), user.nama, w, 'valid', 'import']);
        have[sku] = 1; dibuat.push(info);
      });
      appendRows_(log, 9, out, TEXT_COLS.scan_log);
      if (out.length) markDirty_(sesi.id);
      return { masuk: dibuat.length, dilewati: skip, dibuat: dibuat };
    } finally { lock.releaseLock(); }
  },

  // Link Google Sheet (seluruh file) dan tab hasil sesi terpilih, untuk tombol di menu Admin
  sheetUrl: function (req, user) {
    requireAdmin_(user);
    const ss = SpreadsheetApp.getActiveSpreadsheet(), url = ss.getUrl();
    let hasil = '';
    if (req.sesi_id) { const hs = ss.getSheetByName(hasilName_(String(req.sesi_id))); if (hs) hasil = url + '#gid=' + hs.getSheetId(); }
    return { url: url, hasil: hasil };
  },

  getToko: function (req, user) { requireAdmin_(user); return getToko_(); },

  setToko: function (req, user) {
    requireAdmin_(user);
    const kode = cleanPart_(req.kode, 20).toUpperCase(), nama = cleanPart_(req.nama, 30);
    if (!kode || !nama) throw new Error('Kode toko dan nama toko wajib diisi');
    const p = PropertiesService.getScriptProperties();
    p.setProperty('TOKO_KODE', kode); p.setProperty('TOKO_NAMA', nama);
    return { kode: kode, nama: nama };
  },

  // Progres scan: sudah / belum / di luar system
  progress: function (req) {
    const sesi = getSesiFast_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    const master = masterSkus_(), total = Object.keys(master).length;
    const log = sh_('scan_log'), n = log.getLastRow() - 1, seen = {};
    if (n > 0) log.getRange(2, 2, n, 7).getValues().forEach(function (r) {
      if (String(r[0]) === sesi.id && r[6] === 'valid') seen[String(r[1])] = 1;
    });
    let sudah = 0, luar = 0;
    Object.keys(seen).forEach(function (k) { if (master[k]) sudah++; else luar++; });
    return { mode: sesi.mode, total: total, sudah: sudah, belum: Math.max(total - sudah, 0), luar: luar,
             persen: total ? Math.round(sudah / total * 1000) / 10 : 0 };
  },

  // Info data terakhir di recycle bin untuk sesi ini
  binInfo: function (req, user) {
    requireAdmin_(user);
    const rows = binRows_().filter(function (r) { return String(r[4]) === String(req.sesi_id); });
    if (!rows.length) return { ada: false };
    const last = rows[rows.length - 1], id = String(last[0]);
    const batch = rows.filter(function (r) { return String(r[0]) === id; });
    return { ada: true, bin_id: id, dihapus_pada: String(last[1]), dihapus_oleh: String(last[2]), baris: batch.length, sku: uniqValid_(batch, 10, 5) };
  },

  // Reset: pindahkan semua hasil scan sesi ke recycle_bin, lalu kosongkan
  resetSesi: function (req, user) {
    requireAdmin_(user);
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    if (sesi.status !== 'aktif') throw new Error('Sesi sudah selesai, tidak bisa direset');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const log = sh_('scan_log'), n = log.getLastRow() - 1;
      const all = n > 0 ? log.getRange(2, 1, n, 9).getValues() : [];
      const mine = all.filter(function (r) { return String(r[1]) === sesi.id; });
      const rest = all.filter(function (r) { return String(r[1]) !== sesi.id; });
      if (!mine.length) throw new Error('Belum ada hasil scan di sesi ini');
      const id = binId_(), waktu = now_();
      appendRows_(binSheet_(), 12, mine.map(function (r) { return [id, waktu, user.nama].concat(r); }), TEXT_COLS.recycle_bin);  // simpan dulu
      rewrite_(log, 9, rest, TEXT_COLS.scan_log);                                                                              // baru hapus
      writeHasil_(sesi);
      return { bin_id: id, baris: mine.length, sku: uniqValid_(mine, 7, 2) };
    } finally { lock.releaseLock(); }
  },

  // Pulihkan data terakhir di recycle bin. Hasil scan yang ada sekarang dipindah ke bin sebagai gantinya.
  restoreSesi: function (req, user) {
    requireAdmin_(user);
    const sesi = getSesi_(req.sesi_id);
    if (!sesi) throw new Error('Sesi tidak ditemukan');
    if (sesi.status !== 'aktif') throw new Error('Sesi sudah selesai, tidak bisa dipulihkan');
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const bin = binSheet_(), brows = binRows_();
      const mineBin = brows.filter(function (r) { return String(r[4]) === sesi.id; });
      if (!mineBin.length) throw new Error('Tidak ada data di recycle bin untuk sesi ini');
      const bid = String(mineBin[mineBin.length - 1][0]);
      const batch = mineBin.filter(function (r) { return String(r[0]) === bid; });

      const log = sh_('scan_log'), n = log.getLastRow() - 1;
      const all = n > 0 ? log.getRange(2, 1, n, 9).getValues() : [];
      const cur = all.filter(function (r) { return String(r[1]) === sesi.id; });
      const rest = all.filter(function (r) { return String(r[1]) !== sesi.id; });

      let backup = '';
      if (cur.length) {                                   // 1. amankan hasil scan saat ini
        backup = binId_(); const w = now_();
        appendRows_(bin, 12, cur.map(function (r) { return [backup, w, user.nama].concat(r); }), TEXT_COLS.recycle_bin);
      }
      rewrite_(log, 9, rest.concat(batch.map(function (r) { return r.slice(3); })), TEXT_COLS.scan_log);   // 2. pulihkan
      const keep = binRows_().filter(function (r) { return String(r[0]) !== bid; });                        // 3. keluarkan dari bin
      rewrite_(bin, 12, keep, TEXT_COLS.recycle_bin);
      writeHasil_(sesi);
      return { bin_id: bid, dipulihkan: batch.length, sku: uniqValid_(batch, 10, 5), backup: backup };
    } finally { lock.releaseLock(); }
  }
};
