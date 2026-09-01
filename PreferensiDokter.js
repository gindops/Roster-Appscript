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
 * Data disimpan permanen di sheet Preferensi_Dokter:
 *   ID Dokter | Nama Dokter | Tipe | Nilai | Prioritas | Keterangan
 * ============================================================
 */

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
  sh.getRange(2, 1, MAXR, 1).setNumberFormat('@'); // ID Dokter tetap teks
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
 * Baca sheet Preferensi_Dokter.
 * @return {Object} { namaDokterLower: [ {tipe, nilai, prioritas} ] }
 */
function bacaPreferensiDokter_() {
  const map = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
  if (!sh || sh.getLastRow() < 2) return map;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  rows.forEach(function (r) {
    const nama = String(r[1]).trim();
    const tipe = String(r[2]).trim();
    const nilai = String(r[3]).trim();
    let prioritas = String(r[4]).trim();
    if (!nama || !tipe || !nilai) return;
    if (['Nama', 'Level', 'Kategori'].indexOf(tipe) === -1) return;
    prioritas = (prioritas === 'Wajib') ? 'Wajib' : 'Utamakan';
    const key = nama.toLowerCase();
    if (!map[key]) map[key] = [];
    map[key].push({ tipe: tipe, nilai: nilai, prioritas: prioritas });
  });
  return map;
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
  return pref.tipe + ' "' + pref.nilai + '" (' + pref.prioritas + ')';
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
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
  const out = [];
  rows.forEach(function (r, i) {
    if (!String(r[1]).trim()) return;
    out.push({
      baris: i + 2, id: String(r[0]), nama: String(r[1]), tipe: String(r[2]),
      nilai: String(r[3]), prioritas: String(r[4]) || 'Utamakan', keterangan: String(r[5])
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
  // Lookup ID dokter dari master (bila cocok)
  let id = String(data.id || '').trim();
  if (!id) {
    const d = getMasterData('dokter').filter(function (x) { return x.nama.toLowerCase() === nama.toLowerCase(); })[0];
    if (d) id = d.id;
  }
  return { ok: true, id: id, nama: nama, tipe: tipe, nilai: nilai, prioritas: prioritas, keterangan: String(data.keterangan || '') };
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
    sh.appendRow([v.id, v.nama, v.tipe, v.nilai, v.prioritas, v.keterangan]);
    sh.getRange(sh.getLastRow(), 1).setNumberFormat('@');
    return { ok: true, pesan: 'Preferensi untuk ' + v.nama + ' tersimpan.' };
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
    sh.getRange(baris, 1, 1, 6).setValues([[v.id, v.nama, v.tipe, v.nilai, v.prioritas, v.keterangan]]);
    return { ok: true, pesan: 'Preferensi ' + v.nama + ' diperbarui.' };
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