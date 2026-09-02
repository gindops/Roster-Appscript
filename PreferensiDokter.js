/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Preferensi Dokter (pendamping perawat)
 * File: PreferensiDokter.gs
 *
 * Dokter dapat menyatakan preferensi perawat pendamping berdasarkan:
 *   - Tipe "Nama"     → perawat spesifik (mis. "Rina")
 *   - Tipe "Level"    → grading perawat (Gold / Silver)
 *   - Tipe "Kategori" → New / Existing
 * Tiap preferensi diberi Prioritas:
 *   - "Wajib"     → filter keras saat auto-generate. Bila mustahil
 *                   (mis. perawat cuti / tak kapabel), SLOT DIBIARKAN
 *                   KOSONG dan dilaporkan (kosong = blocker finalisasi).
 *   - "Utamakan"  → bias skor kuat, tetapi tak pernah mengosongkan slot
 *                   atau melanggar aturan wajib (N-05/A-02).
 * Satu dokter boleh punya banyak baris preferensi.
 *
 * CAKUPAN WAKTU (kolom Berlaku/Tanggal):
 *   - "Selalu"  → berlaku setiap kali dokter praktik (perilaku lama).
 *   - "Tanggal" → hanya berlaku pada tanggal yang terdaftar (dd/MM/yyyy, dipisah koma).
 * Bila pada 1 tanggal ada preferensi bertanggal yang cocok, preferensi "Selalu"
 * milik dokter itu DIABAIKAN untuk tanggal tersebut (override menyeluruh).
 * Tujuannya mencegah kombinasi Wajib yang saling bentrok → slot kosong → blocker.
 *
 * Data disimpan permanen di sheet Preferensi_Dokter:
 *   ID Dokter | Nama Dokter | Tipe | Nilai | Prioritas | Keterangan | Berlaku | Tanggal
 * ============================================================
 */

const PD_KOL = 8; // jumlah kolom sheet Preferensi_Dokter

/* ================= MIGRASI ================= */

/**
 * Tambahkan kolom Berlaku/Tanggal pada sheet lama (6 kolom) tanpa merusak data.
 * Baris lama yang Berlaku-nya kosong diisi "Selalu" = persis perilaku sebelumnya.
 * Aman dijalankan berulang.
 */
function migrasiPreferensiBerlaku_(ss) {
  ss = ss || SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
  if (!sh) return;
  // Header kolom G & H
  const hdr = sh.getRange(1, 1, 1, PD_KOL).getValues()[0];
  if (String(hdr[6]).trim() !== 'Berlaku' || String(hdr[7]).trim() !== 'Tanggal') {
    sh.getRange(1, 7, 1, 2).setValues([['Berlaku', 'Tanggal']]);
    if (typeof formatHeader_ === 'function') formatHeader_(sh, PD_KOL);
    sh.setColumnWidth(7, 110); sh.setColumnWidth(8, 260);
  }
  sh.getRange(2, 8, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('@');
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const kol = sh.getRange(2, 7, n, 1).getValues();
  let ubah = false;
  for (let i = 0; i < kol.length; i++) {
    if (!String(kol[i][0]).trim()) { kol[i][0] = 'Selalu'; ubah = true; }
  }
  if (ubah) sh.getRange(2, 7, n, 1).setValues(kol);
}

/* ================= DROPDOWN & MENU ================= */

/** Pasang dropdown di sheet Preferensi_Dokter (dipanggil setup & saat sheet dibuka) */
function pasangValidasiPreferensi_(ss) {
  const sh = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
  if (!sh) return;
  const MAXR = 300;
  const dokter = getMasterData('dokter').filter(function (d) { return d.aktif !== 'Tidak'; }).map(function (d) { return d.nama; });
  if (dokter.length) {
    sh.getRange(2, 2, MAXR, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(dokter, true).setAllowInvalid(true).build());
  }
  sh.getRange(2, 3, MAXR, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Nama', 'Level', 'Kategori'], true).setAllowInvalid(true).build());
  sh.getRange(2, 5, MAXR, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Wajib', 'Utamakan'], true).setAllowInvalid(false).build());
  sh.getRange(2, 7, MAXR, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['Selalu', 'Tanggal'], true).setAllowInvalid(true).build());
  sh.getRange(2, 1, MAXR, 1).setNumberFormat('@'); // ID Dokter tetap teks
  sh.getRange(2, 8, MAXR, 1).setNumberFormat('@'); // Tanggal tetap TEKS dd/MM/yyyy
}

/** Menu: ⭐ Buka Sheet Preferensi Dokter */
function bukaPreferensiDokter() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
  if (!sh) {
    sh = ss.insertSheet(SHEETS.PREFERENSI_DOKTER);
    sh.getRange(1, 1, 1, HEADER.PREFERENSI_DOKTER.length).setValues([HEADER.PREFERENSI_DOKTER]);
    formatHeader_(sh, HEADER.PREFERENSI_DOKTER.length);
  }
  pasangValidasiPreferensi_(ss);
  ss.setActiveSheet(sh);
  SpreadsheetApp.getUi().alert(
    '⭐ Preferensi Dokter',
    'Catat preferensi perawat pendamping per dokter SEBELUM auto-generate.\n\n' +
    '• Nama Dokter: pilih dari dropdown\n' +
    '• Tipe: Nama (perawat spesifik) / Level (Gold/Silver) / Kategori (New/Existing)\n' +
    '• Nilai: isi sesuai Tipe (mis. "Rina", "Gold", atau "Existing")\n' +
    '• Prioritas: Wajib (keras — slot bisa kosong bila mustahil) atau Utamakan (soft)\n\n' +
    'Saat 🤖 Auto-Generate Asistensi, preferensi ini dihormati dan yang tak terpenuhi dilaporkan.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/* ================= BACA & COCOK (dipakai scheduler) ================= */

/**
 * Ubah teks daftar tanggal (dipisah koma / titik-koma / baris baru) menjadi
 * array "dd/MM/yyyy" yang sudah dinormalkan, unik, dan urut.
 * Nilai yang tak bisa diparse dikumpulkan di .buruk agar bisa dilaporkan.
 * @return {Object} {ok:[..], buruk:[..]}
 */
function pdParseTanggalList_(teks) {
  const ok = [], buruk = [], seen = {};
  String(teks || '').split(/[,;\n]/).forEach(function (bagian) {
    const s = String(bagian).trim();
    if (!s) return;
    const d = parseTgl_(s);
    if (!d || isNaN(d.getTime())) { buruk.push(s); return; }
    const norm = formatTgl_(d);
    if (seen[norm]) return;
    seen[norm] = true;
    ok.push(norm);
  });
  ok.sort(function (a, b) { return kunciTgl_(a) < kunciTgl_(b) ? -1 : 1; });
  return { ok: ok, buruk: buruk };
}

/**
 * Baca sheet Preferensi_Dokter.
 * @return {Object} { namaDokterLower: [ {tipe, nilai, prioritas, berlaku, tglSet} ] }
 *   berlaku = 'Selalu' | 'Tanggal'; tglSet = { 'dd/MM/yyyy': true }
 */
function bacaPreferensiDokter_() {
  const map = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
  if (!sh || sh.getLastRow() < 2) return map;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, PD_KOL).getValues();
  rows.forEach(function (r) {
    const nama = String(r[1]).trim();
    const tipe = String(r[2]).trim();
    const nilai = String(r[3]).trim();
    let prioritas = String(r[4]).trim();
    if (!nama || !tipe || !nilai) return;
    if (['Nama', 'Level', 'Kategori'].indexOf(tipe) === -1) return;
    prioritas = (prioritas === 'Wajib') ? 'Wajib' : 'Utamakan';
    // Cakupan waktu — baris lama (kolom kosong) diperlakukan sebagai 'Selalu'.
    let berlaku = String(r[6] || '').trim() === 'Tanggal' ? 'Tanggal' : 'Selalu';
    const tglSet = {};
    if (berlaku === 'Tanggal') {
      const p = pdParseTanggalList_(r[7]);
      // Bertanggal tapi tak punya tanggal sah = tidak bisa dipakai; abaikan barisnya
      // (lebih aman daripada diam-diam berlaku selamanya).
      if (!p.ok.length) return;
      p.ok.forEach(function (t) { tglSet[t] = true; });
    }
    const key = nama.toLowerCase();
    if (!map[key]) map[key] = [];
    map[key].push({ tipe: tipe, nilai: nilai, prioritas: prioritas, berlaku: berlaku, tglSet: tglSet });
  });
  return map;
}

/**
 * Preferensi yang berlaku untuk satu slot (tanggal tertentu).
 * Aturan: bila ada preferensi bertanggal yang cocok dengan tanggal ini,
 * preferensi "Selalu" milik dokter tsb DIABAIKAN (override menyeluruh).
 * @param {Array} semua daftar preferensi 1 dokter dari bacaPreferensiDokter_
 * @param {string} tanggal 'dd/MM/yyyy'
 */
function prefBerlakuUntuk_(semua, tanggal) {
  const arr = semua || [];
  const bertgl = arr.filter(function (p) { return p.berlaku === 'Tanggal' && p.tglSet[tanggal]; });
  if (bertgl.length) return bertgl;
  return arr.filter(function (p) { return p.berlaku !== 'Tanggal'; });
}

/** Apakah perawat p memenuhi 1 preferensi? */
function cocokPreferensi_(p, pref) {
  const nilai = String(pref.nilai || '').trim();
  if (pref.tipe === 'Nama') return String(p.nama || '').toLowerCase() === nilai.toLowerCase();
  if (pref.tipe === 'Level') return String(p.level || '') === nilai;
  if (pref.tipe === 'Kategori') return String(p.kategori || '') === nilai;
  return false;
}

/** Ringkas 1 preferensi jadi teks untuk laporan */
function ringkasPreferensi_(pref) {
  let s = pref.tipe + ' "' + pref.nilai + '" (' + pref.prioritas + ')';
  if (pref && pref.berlaku === 'Tanggal') {
    const tgl = Object.keys(pref.tglSet || {});
    s += ' [tanggal: ' + (tgl.length > 3 ? tgl.slice(0, 3).join(', ') + ' +' + (tgl.length - 3) : tgl.join(', ')) + ']';
  }
  return s;
}

/* ================= CRUD (dipakai web app) ================= */

/** Data awal untuk form web: daftar dokter aktif + daftar perawat aktif */
function getPreferensiFormData() {
  const dokter = getMasterData('dokter').filter(function (d) { return d.aktif !== 'Tidak'; })
    .map(function (d) { return { id: d.id, nama: d.nama, spes: d.spesialisasi }; });
  const perawat = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) {
      return {
        nama: p.nama, level: p.level, kategori: p.kategori,
        kapab: String(p.kapabilitas || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; })
      };
    });
  return { dokter: dokter, perawat: perawat };
}

/** Daftar preferensi tersimpan */
function getPreferensiList() {
  if (typeof wajibModul_ === 'function') wajibModul_('pref'); // hanya peran ber-akses modul pref
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
  if (!sh || sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, PD_KOL).getValues();
  const kiniKunci = kunciTgl_(formatTgl_(new Date()));
  const out = [];
  rows.forEach(function (r, i) {
    if (!String(r[1]).trim()) return;
    const berlaku = String(r[6] || '').trim() === 'Tanggal' ? 'Tanggal' : 'Selalu';
    const tgl = berlaku === 'Tanggal' ? pdParseTanggalList_(r[7]).ok : [];
    // "lewat" = bertanggal & SEMUA tanggalnya sudah lewat (preferensi Selalu tak pernah lewat)
    const lewat = berlaku === 'Tanggal' && tgl.length > 0 &&
      tgl.every(function (t) { return kunciTgl_(t) < kiniKunci; });
    out.push({
      baris: i + 2, id: String(r[0]), nama: String(r[1]), tipe: String(r[2]),
      nilai: String(r[3]), prioritas: String(r[4]) || 'Utamakan', keterangan: String(r[5]),
      berlaku: berlaku, tanggal: tgl, lewat: lewat
    });
  });
  return out;
}

/** Validasi + normalisasi 1 payload preferensi */
function validasiPreferensi_(data) {
  const nama = String(data.nama || '').trim();
  if (!nama) return { ok: false, pesan: 'Nama dokter wajib dipilih.' };
  const tipe = String(data.tipe || '').trim();
  if (['Nama', 'Level', 'Kategori'].indexOf(tipe) === -1) return { ok: false, pesan: 'Tipe harus Nama/Level/Kategori.' };
  let nilai = String(data.nilai || '').trim();
  if (!nilai) return { ok: false, pesan: 'Nilai preferensi wajib diisi.' };
  if (tipe === 'Level' && ['Gold', 'Silver'].indexOf(nilai) === -1) return { ok: false, pesan: 'Nilai Level harus Gold atau Silver.' };
  if (tipe === 'Kategori' && ['New', 'Existing'].indexOf(nilai) === -1) return { ok: false, pesan: 'Nilai Kategori harus New atau Existing.' };
  const prioritas = String(data.prioritas || '').trim() === 'Wajib' ? 'Wajib' : 'Utamakan';
  // --- Cakupan waktu ---
  const berlaku = String(data.berlaku || '').trim() === 'Tanggal' ? 'Tanggal' : 'Selalu';
  let tglTeks = '', catatan = '';
  if (berlaku === 'Tanggal') {
    const sumber = Array.isArray(data.tanggal) ? data.tanggal.join(',') : String(data.tanggal || '');
    const p = pdParseTanggalList_(sumber);
    if (p.buruk.length) return { ok: false, pesan: 'Tanggal tidak dikenali: ' + p.buruk.join(', ') + '. Format harus dd/MM/yyyy.' };
    if (!p.ok.length) return { ok: false, pesan: 'Pilih minimal 1 tanggal, atau ubah Berlaku menjadi "Selalu".' };
    tglTeks = p.ok.join(', ');
    // Di luar periode aktif bukan error — bisa jadi untuk periode berikutnya. Cukup diingatkan.
    const pm = parseTgl_(getConfig('PERIODE_MULAI')), pa = parseTgl_(getConfig('PERIODE_AKHIR'));
    if (pm && pa) {
      const luar = p.ok.filter(function (t) { const d = parseTgl_(t); return d && (d < pm || d > pa); });
      if (luar.length) catatan = ' Catatan: ' + luar.length + ' tanggal di luar periode aktif (' + luar.join(', ') + ') — tersimpan, baru berlaku saat periode itu digenerate.';
    }
  }
  // Lookup ID dokter dari master (bila cocok)
  let id = String(data.id || '').trim();
  if (!id) {
    const d = getMasterData('dokter').filter(function (x) { return x.nama.toLowerCase() === nama.toLowerCase(); })[0];
    if (d) id = d.id;
  }
  return { ok: true, id: id, nama: nama, tipe: tipe, nilai: nilai, prioritas: prioritas,
    keterangan: String(data.keterangan || ''), berlaku: berlaku, tanggal: tglTeks, catatan: catatan };
}

/** Tambah 1 preferensi */
function addPreferensi(data) {
  if (typeof wajibModul_ === 'function') wajibModul_('pref');
  const v = validasiPreferensi_(data);
  if (!v.ok) return v;
  let lock = null;
  try { lock = LockService.getDocumentLock(); lock.waitLock(8000); } catch (e) { lock = null; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
    if (!sh) return { ok: false, pesan: 'Sheet Preferensi_Dokter belum ada. Jalankan Setup Awal Template.' };
    sh.appendRow([v.id, v.nama, v.tipe, v.nilai, v.prioritas, v.keterangan, v.berlaku, v.tanggal]);
    const br = sh.getLastRow();
    sh.getRange(br, 1).setNumberFormat('@');
    sh.getRange(br, 8).setNumberFormat('@').setValue(v.tanggal); // pastikan tetap teks
    return { ok: true, pesan: 'Preferensi untuk ' + v.nama + ' tersimpan.' + v.catatan };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/** Update 1 preferensi berdasarkan nomor baris */
function updatePreferensi(baris, data) {
  if (typeof wajibModul_ === 'function') wajibModul_('pref');
  const v = validasiPreferensi_(data);
  if (!v.ok) return v;
  let lock = null;
  try { lock = LockService.getDocumentLock(); lock.waitLock(8000); } catch (e) { lock = null; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
    if (!sh || baris < 2 || baris > sh.getLastRow()) return { ok: false, pesan: 'Baris tidak valid.' };
    sh.getRange(baris, 1).setNumberFormat('@');
    sh.getRange(baris, 8).setNumberFormat('@');
    sh.getRange(baris, 1, 1, PD_KOL).setValues([[v.id, v.nama, v.tipe, v.nilai, v.prioritas, v.keterangan, v.berlaku, v.tanggal]]);
    return { ok: true, pesan: 'Preferensi ' + v.nama + ' diperbarui.' + v.catatan };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/** Hapus 1 preferensi berdasarkan nomor baris */
function deletePreferensi(baris) {
  if (typeof wajibModul_ === 'function') wajibModul_('pref');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
  if (!sh || baris < 2 || baris > sh.getLastRow()) return { ok: false, pesan: 'Baris tidak valid.' };
  sh.deleteRow(baris);
  return { ok: true, pesan: 'Preferensi dihapus.' };
}