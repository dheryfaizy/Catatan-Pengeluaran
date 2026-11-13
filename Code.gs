/**
 * Web App Rekapitulasi Pengeluaran Bulanan
 * Backend: Google Apps Script + Google Sheets
 *
 * Cara pakai: buka Spreadsheet > Extensions > Apps Script (script terikat ke Spreadsheet),
 * atau isi SPREADSHEET_ID jika script dibuat terpisah.
 */
const CONFIG = {
  SPREADSHEET_ID: '', // kosongkan jika script terikat ke Spreadsheet
  SHEET_RINGKASAN: 'Ringkasan & Limit Anggaran',
  SHEET_JURNAL: 'Jurnal Transaksi Harian',
  SHEET_LOG: 'Log Backup', // dibuat otomatis
  FOLDER_BUKTI: 'Bukti Struk Pengeluaran' // folder Google Drive, dibuat otomatis
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dompet Ayah Mima')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ---------- Helpers ---------- */

function getSheet_(name) {
  const ss = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet "' + name + '" tidak ditemukan.');
  return sh;
}

function fmt_(date, pattern) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), pattern);
}

function toNumber_(v) {
  if (typeof v === 'number') return v;
  const n = Number(String(v || '').replace(/[^\d-]/g, ''));
  return isNaN(n) ? 0 : n;
}

function parseDate_(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  const m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(v));
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : null;
}

function statusOf_(limit, sisa) {
  if (limit === 0) return 'WASPADA';
  if (sisa < 0) return 'OVER BUDGET';
  if (sisa === 0) return 'WASPADA';
  return 'AMAN';
}

/** Baca seluruh jurnal: [{row, date, kategori, jumlah, catatan}] */
function readJurnal_() {
  const sh = getSheet_(CONFIG.SHEET_JURNAL);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 5).getValues()
    .map(function (r, i) {
      return {
        row: i + 2,
        date: parseDate_(r[0]),
        kategori: String(r[1]).trim(),
        jumlah: toNumber_(r[2]),
        catatan: String(r[3] || ''),
        bukti: String(r[4] || '')
      };
    })
    .filter(function (t) { return t.date && t.kategori; });
}

/* ---------- API untuk Frontend ---------- */

/** Daftar nama kategori untuk dropdown form. */
function getKategori() {
  const sh = getSheet_(CONFIG.SHEET_RINGKASAN);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 2, last - 1, 1).getValues()
    .map(function (r) { return String(r[0]).trim(); })
    .filter(function (n) { return n && !/^total/i.test(n); });
}

/** Ringkasan limit, realisasi, sisa, dan status per kategori untuk bulan tertentu (yyyy-MM). */
function getDashboardData(bulan) {
  bulan = bulan || fmt_(new Date(), 'yyyy-MM');
  const spent = {};
  readJurnal_().forEach(function (t) {
    if (fmt_(t.date, 'yyyy-MM') === bulan) spent[t.kategori] = (spent[t.kategori] || 0) + t.jumlah;
  });

  const sh = getSheet_(CONFIG.SHEET_RINGKASAN);
  const last = sh.getLastRow();
  const rows = last < 2 ? [] : sh.getRange(2, 1, last - 1, 4).getValues();

  let totalLimit = 0, totalReal = 0;
  const items = rows
    .filter(function (r) {
      const n = String(r[1]).trim();
      return n && !/^total/i.test(n);
    })
    .map(function (r) {
      const nama = String(r[1]).trim();
      const limit = toNumber_(r[3]);
      const realisasi = spent[nama] || 0;
      const sisa = limit - realisasi;
      totalLimit += limit;
      totalReal += realisasi;
      return {
        no: r[0],
        nama: nama,
        tipe: String(r[2] || '-'),
        limit: limit,
        realisasi: realisasi,
        sisa: sisa,
        persen: limit > 0 ? Math.round(realisasi / limit * 100) : 0,
        status: statusOf_(limit, sisa)
      };
    });

  return {
    bulan: bulan,
    items: items,
    totals: { limit: totalLimit, realisasi: totalReal, sisa: totalLimit - totalReal }
  };
}

/** Tambah transaksi (hanya teks, cepat). Foto diunggah terpisah lewat unggahBukti(). Mengembalikan nomor baris. */
function tambahTransaksi(data) {
  const tgl = parseDate_(data.tanggal);
  const jumlah = toNumber_(data.jumlah);
  const kategori = String(data.kategori || '').trim();
  if (!tgl) throw new Error('Tanggal tidak valid.');
  if (jumlah <= 0) throw new Error('Jumlah harus lebih dari 0.');
  if (getKategori().indexOf(kategori) === -1) throw new Error('Kategori tidak ditemukan.');

  let baris = 0;
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet_(CONFIG.SHEET_JURNAL);
    if (!sh.getRange(1, 5).getValue()) sh.getRange(1, 5).setValue('Bukti Struk');
    sh.appendRow([tgl, kategori, jumlah, String(data.catatan || '').trim()]);
    baris = sh.getLastRow();
    sh.getRange(baris, 1).setNumberFormat('yyyy-mm-dd');
    sh.getRange(baris, 3).setNumberFormat('#,##0');
  } finally {
    lock.releaseLock();
  }
  log_('TAMBAH', fmt_(tgl, 'yyyy-MM-dd'), kategori, jumlah, data.catatan); // di luar lock agar respons cepat
  return { ok: true, row: baris };
}

/** Riwayat transaksi, difilter bulan (yyyy-MM) dan kategori (opsional), terbaru di atas. */
function getRiwayat(bulan, kategori) {
  const list = readJurnal_()
    .filter(function (t) {
      return (!bulan || fmt_(t.date, 'yyyy-MM') === bulan) && (!kategori || t.kategori === kategori);
    })
    .sort(function (a, b) { return (b.date.getTime() - a.date.getTime()) || (b.row - a.row); }) // tanggal terbaru di atas; dalam satu tanggal, transaksi terbaru di atas
    .map(function (t) {
      return { row: t.row, tanggal: fmt_(t.date, 'yyyy-MM-dd'), kategori: t.kategori, jumlah: t.jumlah, catatan: t.catatan, bukti: t.bukti };
    });
  const total = list.reduce(function (s, t) { return s + t.jumlah; }, 0);
  return { items: list, total: total };
}

/** Hapus transaksi berdasarkan nomor baris; kategori & jumlah dicek agar tidak salah hapus. */
function hapusTransaksi(row, kategori, jumlah) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(CONFIG.SHEET_JURNAL);
    if (row < 2 || row > sh.getLastRow()) throw new Error('Data tidak ditemukan. Muat ulang riwayat.');
    const r = sh.getRange(row, 1, 1, 4).getValues()[0];
    if (String(r[1]).trim() !== kategori || toNumber_(r[2]) !== Number(jumlah)) {
      throw new Error('Data sudah berubah. Muat ulang riwayat.');
    }
    log_('HAPUS', fmt_(parseDate_(r[0]) || new Date(), 'yyyy-MM-dd'), String(r[1]).trim(), toNumber_(r[2]), r[3]);
    sh.deleteRow(row);
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/* ---------- Backup otomatis (sheet "Log Backup") ---------- */

function ss_() {
  return CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

/** Mencatat setiap aksi (tambah/hapus transaksi, tambah/edit kategori) ke sheet log. Kegagalan log tidak membatalkan transaksi. */
function log_(aksi, tanggal, kategori, jumlah, catatan) {
  try {
    const ss = ss_();
    let sh = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (!sh) {
      sh = ss.insertSheet(CONFIG.SHEET_LOG);
      sh.appendRow(['Waktu Input', 'Aksi', 'Tanggal Transaksi', 'Kategori', 'Jumlah (Rp)', 'Catatan']);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, 6).setFontWeight('bold');
    }
    sh.appendRow([new Date(), aksi, tanggal, kategori, jumlah, String(catatan || '')]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
    SpreadsheetApp.flush();
  } catch (e) {
    console.error('Log backup gagal: ' + e.message);
  }
}

/* ---------- Kategori anggaran: tambah & edit ---------- */

/** data = {asal (kosong = kategori baru), nama, tipe, limit} */
function simpanKategori(data) {
  const nama = String(data.nama || '').trim();
  const tipe = String(data.tipe || '-').trim();
  const limit = Math.max(0, toNumber_(data.limit));
  const asal = String(data.asal || '').trim();
  if (!nama) throw new Error('Nama kategori wajib diisi.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(CONFIG.SHEET_RINGKASAN);
    const last = sh.getLastRow();
    const names = last < 2 ? [] : sh.getRange(2, 2, last - 1, 1).getValues()
      .map(function (r) { return String(r[0]).trim(); });
    const dup = names.indexOf(nama);

    if (asal) { // EDIT
      const idx = names.indexOf(asal);
      if (idx === -1) throw new Error('Kategori asal tidak ditemukan.');
      if (nama !== asal && dup !== -1) throw new Error('Nama kategori sudah dipakai.');
      sh.getRange(idx + 2, 2, 1, 3).setValues([[nama, tipe, limit]]);
      if (nama !== asal) { // ikutkan perubahan nama ke seluruh transaksi lama
        const jr = getSheet_(CONFIG.SHEET_JURNAL);
        const jl = jr.getLastRow();
        if (jl >= 2) {
          const rng = jr.getRange(2, 2, jl - 1, 1);
          rng.setValues(rng.getValues().map(function (x) { return [String(x[0]).trim() === asal ? nama : x[0]]; }));
        }
      }
      log_('KATEGORI-EDIT', '', nama, limit, 'Asal: ' + asal + ' | Tipe: ' + tipe);
    } else { // TAMBAH
      if (dup !== -1) throw new Error('Nama kategori sudah ada.');
      const r = Math.max(last, 1) + 1;
      const j = "'" + CONFIG.SHEET_JURNAL + "'";
      sh.getRange(r, 1, 1, 7).setValues([[
        r - 1, nama, tipe, limit,
        '=SUMIF(' + j + '!B:B,B' + r + ',' + j + '!C:C)',
        '=D' + r + '-E' + r,
        '=IF(D' + r + '=0,"WASPADA",IF(F' + r + '<0,"OVER BUDGET",IF(F' + r + '=0,"WASPADA","AMAN")))'
      ]]);
      sh.getRange(r, 4, 1, 3).setNumberFormat('#,##0');
      log_('KATEGORI-TAMBAH', '', nama, limit, 'Tipe: ' + tipe);
    }
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/* ---------- Foto bukti (Google Drive, diunggah di latar belakang) ---------- */

/** Jalankan SEKALI dari editor Apps Script untuk memberi izin akses Google Drive. */
function izinkanDrive() {
  DriveApp.getRootFolder();
}

function retry_(fn) {
  let err;
  for (let i = 0; i < 3; i++) {
    try { return fn(); } catch (e) { err = e; Utilities.sleep(500 * Math.pow(2, i)); }
  }
  throw err;
}

function folderBukti_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('FOLDER_BUKTI_ID');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { /* folder dihapus, buat ulang */ } }
  const it = DriveApp.getFoldersByName(CONFIG.FOLDER_BUKTI);
  const f = it.hasNext() ? it.next() : DriveApp.createFolder(CONFIG.FOLDER_BUKTI);
  props.setProperty('FOLDER_BUKTI_ID', f.getId());
  return f;
}

/** foto = {mime, data (base64)}; mengembalikan URL file di Drive. */
function simpanFoto_(foto, tgl, kategori) {
  const b64 = String(foto.data).replace(/^data:[^,]*,/, '');
  if (b64.length > 8000000) throw new Error('Ukuran foto terlalu besar.');
  const nama = fmt_(tgl, 'yyyyMMdd') + '_' + kategori.replace(/[^\w\-]+/g, '_') + '_' + Date.now() + '.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), foto.mime || 'image/jpeg', nama);
  return folderBukti_().createFile(blob).getUrl();
}

function cocok_(r, tglStr, kategori, jumlah) {
  const d = parseDate_(r[0]);
  return !!d && fmt_(d, 'yyyy-MM-dd') === tglStr && String(r[1]).trim() === kategori && toNumber_(r[2]) === jumlah;
}

/** Cari baris transaksi: cek nomor baris dulu, jika sudah bergeser cari dari bawah yang cocok & belum punya bukti. */
function cariBaris_(sh, row, tglStr, kategori, jumlah) {
  const last = sh.getLastRow();
  if (row >= 2 && row <= last && cocok_(sh.getRange(row, 1, 1, 3).getValues()[0], tglStr, kategori, jumlah)) return row;
  if (last < 2) return 0;
  const v = sh.getRange(2, 1, last - 1, 5).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    if (cocok_(v[i], tglStr, kategori, jumlah) && !v[i][4]) return i + 2;
  }
  return 0;
}

/** Dipanggil frontend setelah transaksi tersimpan. p = {row, tanggal, kategori, jumlah, foto:{mime,data}} */
function unggahBukti(p) {
  const tgl = parseDate_(p.tanggal);
  const kategori = String(p.kategori || '').trim();
  const jumlah = Number(p.jumlah);
  if (!tgl) throw new Error('Tanggal tidak valid.');
  if (!p.foto || !p.foto.data) throw new Error('Data foto kosong.');

  const url = retry_(function () { return simpanFoto_(p.foto, tgl, kategori); }); // unggah di luar lock

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(CONFIG.SHEET_JURNAL);
    const tglStr = fmt_(tgl, 'yyyy-MM-dd');
    const row = cariBaris_(sh, Number(p.row), tglStr, kategori, jumlah);
    if (!row) throw new Error('Baris transaksi tidak ditemukan. Foto ada di Drive: ' + url);
    sh.getRange(row, 5).setValue(url);
    log_('BUKTI', tglStr, kategori, jumlah, 'Bukti: ' + url);
  } finally {
    lock.releaseLock();
  }
  return { ok: true, url: url };
}

/* ---------- Hapus kategori ---------- */

/** Kategori hanya bisa dihapus jika belum dipakai transaksi apa pun. */
function hapusKategori(nama) {
  nama = String(nama || '').trim();
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = getSheet_(CONFIG.SHEET_RINGKASAN);
    const last = sh.getLastRow();
    const names = last < 2 ? [] : sh.getRange(2, 2, last - 1, 1).getValues()
      .map(function (r) { return String(r[0]).trim(); });
    const idx = names.indexOf(nama);
    if (idx === -1) throw new Error('Kategori tidak ditemukan.');
    const dipakai = readJurnal_().filter(function (t) { return t.kategori === nama; }).length;
    if (dipakai > 0) {
      throw new Error('Kategori masih dipakai ' + dipakai + ' transaksi. Hapus transaksinya dulu atau cukup ubah namanya.');
    }
    sh.deleteRow(idx + 2);
    log_('KATEGORI-HAPUS', '', nama, 0, '');
  } finally {
    lock.releaseLock();
  }
  return { ok: true };
}

/* ---------- Quote of the day (teks yang bisa diedit di Home) ---------- */

function getStatus() {
  return PropertiesService.getScriptProperties().getProperty('STATUS_TEXT') || 'Catat sekarang. Pantau kapan saja.';
}

/** Simpan teks (maks 160 karakter). Teks kosong = kembali ke teks bawaan. */
function simpanStatus(teks) {
  const t = String(teks || '').replace(/\s+/g, ' ').trim().slice(0, 160);
  const props = PropertiesService.getScriptProperties();
  if (t) props.setProperty('STATUS_TEXT', t); else props.deleteProperty('STATUS_TEXT');
  return getStatus();
}

/* ---------- Ubah urutan transaksi (drag & drop) ---------- */

/**
 * Mengubah urutan transaksi dalam SATU tanggal.
 * rows = nomor baris sheet milik transaksi, sesuai urutan tampilan baru (atas → bawah).
 * Tampilan mengurutkan baris terbesar di atas, jadi isi baris rows[0] dipindah ke nomor baris terbesar, dst.
 */
function urutkanTransaksi(tanggal, rows) {
  const tgl = parseDate_(tanggal);
  rows = (rows || []).map(Number);
  if (!tgl) throw new Error('Tanggal tidak valid.');
  if (rows.length < 2) return { ok: true };
  if (rows.some(isNaN) || new Set(rows).size !== rows.length) throw new Error('Data urutan tidak valid.');
  const tglStr = fmt_(tgl, 'yyyy-MM-dd');

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSheet_(CONFIG.SHEET_JURNAL);
    const last = sh.getLastRow();
    const all = last < 2 ? [] : sh.getRange(2, 1, last - 1, 5).getValues();
    const baru = rows.map(function (r) {
      const v = all[r - 2];
      const d = v && parseDate_(v[0]);
      if (!d || fmt_(d, 'yyyy-MM-dd') !== tglStr) throw new Error('Data sudah berubah. Muat ulang riwayat.');
      return v;
    });
    const slot = rows.slice().sort(function (a, b) { return b - a; });
    slot.forEach(function (r, i) { sh.getRange(r, 1, 1, 5).setValues([baru[i]]); });
  } finally {
    lock.releaseLock();
  }
  log_('URUT', tglStr, '', 0, rows.length + ' transaksi diurutkan ulang');
  return { ok: true };
}
