/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Request Cuti / Libur
 * File: RequestCuti.gs
 *
 * Proses SEBELUM generate jadwal: SPV mencatat permintaan cuti/
 * libur pada tanggal tertentu untuk dokter, perawat, receptionist
 * di sheet Request_Cuti. Saat generate:
 *  - DOKTER cuti  → tidak dijadwalkan praktik pada tanggal tsb
 *                   (baris hilang dari Jadwal_Dokter)
 *  - PERAWAT/RECEPTIONIST cuti → otomatis jadi LIBUR pada tanggal
 *                   tsb (dihormati auto-generate & validator)
 *
 * Format sheet Request_Cuti (mulai baris 2):
 *   Nama | Peran | Tanggal Mulai | Tanggal Akhir (opsional) | Keterangan
 *   Tanggal: dd/MM/yyyy. Akhir kosong = 1 hari saja.
 * ============================================================
 */

/** Menu: 📝 Input Request Cuti/Libur — buka sheet & pastikan dropdown terpasang */
function bukaRequestCuti() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let rc = ss.getSheetByName(SHEETS.REQUEST_CUTI);
  if (!rc) {
    rc = ss.insertSheet(SHEETS.REQUEST_CUTI);
    rc.getRange(1, 1, 1, HEADER.REQUEST_CUTI.length).setValues([HEADER.REQUEST_CUTI]);
    formatHeader_(rc, HEADER.REQUEST_CUTI.length);
  }
  pasangValidasiRequest_(ss);   // segarkan dropdown nama (staf terbaru)
  ss.setActiveSheet(rc);
  SpreadsheetApp.getUi().alert(
    '📝 Request Cuti/Libur',
    'Isi permintaan cuti/libur di sheet ini SEBELUM generate jadwal.\n\n' +
    '• Nama: pilih dari dropdown (dokter/perawat/receptionist)\n' +
    '• Peran: pilih; bila dikosongkan sistem mendeteksi otomatis\n' +
    '• Tanggal Mulai wajib (dd/MM/yyyy). Tanggal Akhir opsional (untuk rentang beberapa hari)\n' +
    '• Keterangan: mis. Cuti tahunan, Sakit, Izin\n\n' +
    'Saat 🚀 Auto-Generate / ⚡ Generate Jadwal dijalankan:\n' +
    '– Dokter cuti tidak akan dijadwalkan praktik pada tanggal itu\n' +
    '– Perawat/Receptionist cuti otomatis ditandai libur pada tanggal itu',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Baca & olah Request_Cuti.
 * @return {Object} {
 *   dokter: { namaLower: { 'dd/MM/yyyy': true } },   // untuk hapus dari Jadwal_Dokter
 *   staf:   { namaProper: [ 'dd/MM/yyyy', ... ] },    // perawat+receptionist → libur wajib
 *   jumlah: n                                          // total baris request valid
 * }
 */
function bacaRequestCuti_() {
  const hasil = { dokter: {}, staf: {}, jumlah: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rc = ss.getSheetByName(SHEETS.REQUEST_CUTI);
  if (!rc || rc.getLastRow() < 2) return hasil;

  // Lookup master untuk deteksi peran & nama proper
  const map = {}; // namaLower → {peran, nama}
  getMasterData('dokter').forEach(function (d) { map[d.nama.toLowerCase()] = { peran: 'Dokter', nama: d.nama }; });
  getMasterData('perawat').forEach(function (p) { map[p.nama.toLowerCase()] = { peran: 'Perawat', nama: p.nama }; });
  getMasterData('receptionist').forEach(function (r) { map[r.nama.toLowerCase()] = { peran: 'Receptionist', nama: r.nama }; });

  const rows = rc.getRange(2, 1, rc.getLastRow() - 1, 5).getValues();
  rows.forEach(function (r) {
    const nama = String(r[0]).trim();
    if (!nama) return;
    let peran = String(r[1]).trim();
    const key = nama.toLowerCase();
    const ref = map[key];
    if (!peran && ref) peran = ref.peran;          // deteksi otomatis
    const namaProper = ref ? ref.nama : nama;

    const mulai = parseTgl_(normalTgl_(r[2]));
    if (!mulai) return;                            // tanggal mulai wajib & valid
    let akhir = parseTgl_(normalTgl_(r[3]));
    if (!akhir || akhir < mulai) akhir = mulai;    // kosong/invalid → 1 hari

    const tanggalArr = [];
    for (let d = new Date(mulai); d <= akhir; d.setDate(d.getDate() + 1)) {
      tanggalArr.push(formatTgl_(d));
    }
    if (!tanggalArr.length) return;
    hasil.jumlah++;

    if (peran === 'Dokter') {
      if (!hasil.dokter[key]) hasil.dokter[key] = {};
      tanggalArr.forEach(function (t) { hasil.dokter[key][t] = true; });
    } else {
      // Perawat / Receptionist (atau tak dikenal → anggap staf libur)
      if (!hasil.staf[namaProper]) hasil.staf[namaProper] = [];
      tanggalArr.forEach(function (t) {
        if (hasil.staf[namaProper].indexOf(t) === -1) hasil.staf[namaProper].push(t);
      });
    }
  });
  return hasil;
}

/** Normalisasi input tanggal: terima Date, "dd/MM/yyyy", atau "dd-MM-yyyy" → "dd/MM/yyyy" */
function normalTgl_(v) {
  if (v instanceof Date) return formatTgl_(v);
  const s = String(v || '').trim().replace(/-/g, '/');
  return s;
}

/* ================= SIDEBAR REQUEST CUTI ================= */

/** Menu: 📝 Input Request Cuti/Libur (form) */
function showRequestCutiSidebar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(SHEETS.REQUEST_CUTI)) {
    const rc = ss.insertSheet(SHEETS.REQUEST_CUTI);
    rc.getRange(1, 1, 1, HEADER.REQUEST_CUTI.length).setValues([HEADER.REQUEST_CUTI]);
    formatHeader_(rc, HEADER.REQUEST_CUTI.length);
  }
  const html = HtmlService.createHtmlOutputFromFile('SidebarRequestCuti')
    .setTitle('Request Cuti / Libur — Audy Dental');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Daftar staf aktif untuk dropdown (nama + peran + level/kategori) */
function getStafListForRequest() {
  const out = [];
  getMasterData('dokter').forEach(function (d) { if (d.aktif !== 'Tidak') out.push({ nama: d.nama, peran: 'Dokter', level: '', kategori: '' }); });
  getMasterData('perawat').forEach(function (p) { if (p.aktif !== 'Tidak') out.push({ nama: p.nama, peran: 'Perawat', level: p.level || '', kategori: p.kategori || '' }); });
  getMasterData('receptionist').forEach(function (r) { if (r.aktif !== 'Tidak') out.push({ nama: r.nama, peran: 'Resepsionis', level: r.level || '', kategori: '' }); });
  return out;
}

/**
 * Tambah 1 request cuti dari sidebar.
 * @param {Object} data {nama, peran, mulai(yyyy-mm-dd), akhir(yyyy-mm-dd|''), keterangan}
 */
function addRequestCuti(data) {
  let lock = null;
  try { lock = LockService.getDocumentLock(); lock.waitLock(8000); } catch (e) { lock = null; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rc = ss.getSheetByName(SHEETS.REQUEST_CUTI);
    if (!rc) return { ok: false, pesan: 'Sheet Request_Cuti belum ada. Jalankan Setup Awal Template.' };

    const nama = String(data.nama || '').trim();
    if (!nama) return { ok: false, pesan: 'Nama wajib dipilih.' };
    const mulai = isoKeTgl_(data.mulai);
    if (!mulai) return { ok: false, pesan: 'Tanggal mulai wajib diisi.' };
    let akhir = isoKeTgl_(data.akhir);
    // validasi rentang
    const dM = parseTgl_(mulai), dA = akhir ? parseTgl_(akhir) : null;
    if (dA && dA < dM) return { ok: false, pesan: 'Tanggal akhir tidak boleh sebelum tanggal mulai.' };

    const catatan = cekAturanRequest_(nama, mulai, akhir); // detail aturan saat pembuatan
    rc.appendRow([nama, String(data.peran || ''), '', '', String(data.keterangan || ''), catatan]);
    const rr = rc.getLastRow();
    rc.getRange(rr, 3, 1, 2).setNumberFormat('@').setValues([[mulai, akhir || '']]); // paksa teks agar tak jadi Date
    return { ok: true, pesan: 'Request cuti untuk ' + nama + ' tersimpan.', warning: catatan };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/** Daftar request cuti tersimpan (untuk ditampilkan & dihapus di sidebar) */
function getRequestCutiList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rc = ss.getSheetByName(SHEETS.REQUEST_CUTI);
  if (!rc || rc.getLastRow() < 2) return [];
  const rows = rc.getRange(2, 1, rc.getLastRow() - 1, HEADER.REQUEST_CUTI.length).getValues();
  const out = [];
  rows.forEach(function (r, i) {
    if (!String(r[0]).trim()) return;
    const mulai = fmtTglCell_(r[2]), akhir = fmtTglCell_(r[3]);
    const tersimpan = String(r[5] || '').trim(); // Catatan Aturan disimpan saat pembuatan
    out.push({
      baris: i + 2, nama: String(r[0]), peran: String(r[1]),
      mulai: mulai, akhir: akhir, keterangan: String(r[4]),
      catatan: tersimpan,
      warning: tersimpan || cekAturanRequest_(String(r[0]), mulai, akhir)
    });
  });
  return out;
}

/** Normalisasi sel tanggal (Date atau teks) → "dd/MM/yyyy" bersih untuk tampilan */
function fmtTglCell_(v) {
  if (v instanceof Date) return formatTgl_(v);
  const s = String(v || '').trim();
  if (!s) return '';
  const d = parseTgl_(normalTgl_(s));
  return d ? formatTgl_(d) : s;
}

/** Update 1 request cuti berdasarkan nomor baris (untuk aksi Edit) */
function updateRequestCuti(baris, data) {
  let lock = null;
  try { lock = LockService.getDocumentLock(); lock.waitLock(8000); } catch (e) { lock = null; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rc = ss.getSheetByName(SHEETS.REQUEST_CUTI);
    if (!rc || baris < 2 || baris > rc.getLastRow()) return { ok: false, pesan: 'Baris tidak valid.' };
    const nama = String(data.nama || '').trim();
    if (!nama) return { ok: false, pesan: 'Nama wajib dipilih.' };
    const mulai = isoKeTgl_(data.mulai);
    if (!mulai) return { ok: false, pesan: 'Tanggal mulai wajib diisi.' };
    let akhir = isoKeTgl_(data.akhir);
    const dM = parseTgl_(mulai), dA = akhir ? parseTgl_(akhir) : null;
    if (dA && dA < dM) return { ok: false, pesan: 'Tanggal akhir tidak boleh sebelum mulai.' };
    const catatan = cekAturanRequest_(nama, mulai, akhir); // detail aturan diperbarui saat edit
    rc.getRange(baris, 3, 1, 2).setNumberFormat('@');
    rc.getRange(baris, 1, 1, HEADER.REQUEST_CUTI.length).setValues([[nama, String(data.peran || ''), mulai, akhir || '', String(data.keterangan || ''), catatan]]);
    return { ok: true, pesan: 'Request cuti ' + nama + ' diperbarui.', warning: catatan };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/** Hapus 1 request cuti berdasarkan nomor baris */
function deleteRequestCuti(baris) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rc = ss.getSheetByName(SHEETS.REQUEST_CUTI);
  if (!rc || baris < 2 || baris > rc.getLastRow()) return { ok: false, pesan: 'Baris tidak valid.' };
  rc.deleteRow(baris);
  return { ok: true, pesan: 'Request dihapus.' };
}

/**
 * Cek aturan saat request cuti dibuat (keputusan #4): kembalikan pesan WARNING
 * bila permintaan melanggar aturan (mis. perawat Silver minta libur weekend).
 * Permintaan tetap disimpan & dihormati sebagai pengecualian — bukan blocker.
 */
function cekAturanRequest_(nama, mulaiTgl, akhirTgl) {
  const low = String(nama || '').toLowerCase();
  const prw = getMasterData('perawat').filter(function (p) { return p.nama.toLowerCase() === low; })[0];
  const rcp = getMasterData('receptionist').filter(function (r) { return r.nama.toLowerCase() === low; })[0];
  const weekend = String(getConfig('HARI_WEEKEND') || 'Sabtu,Minggu').split(',').map(function (h) { return h.trim().toLowerCase(); });
  const dM = parseTgl_(mulaiTgl); if (!dM) return '';
  const dA = akhirTgl ? (parseTgl_(akhirTgl) || dM) : dM;
  let adaWeekend = false, jml = 0;
  for (let d = new Date(dM); d <= dA; d.setDate(d.getDate() + 1)) {
    jml++;
    const hari = NAMA_HARI[d.getDay()].toLowerCase();
    if (weekend.indexOf(hari) !== -1) adaWeekend = true;
  }
  const temuan = [];
  // N-03 hanya berlaku untuk perawat (keputusan #11: weekend tak berlaku bagi resepsionis)
  if (prw && prw.level === 'Silver' && adaWeekend) temuan.push({ kode: 'N-03', teks: 'Perawat Silver diminta libur weekend — hanya Gold yang boleh.' });
  const jatah = Number(getConfig('LIBUR_WAJIB_PERAWAT')) || 4;
  if (jml > jatah) temuan.push({ kode: (rcp && !prw) ? 'R-06' : 'N-02', teks: 'Rentang ' + jml + ' hari melebihi jatah libur normal (' + jatah + ') per periode.' });
  if (!temuan.length) return '';
  return kelompokkanTemuan_(temuan) +
    '\n\nCatatan: permintaan tetap disimpan & dihormati sebagai pengecualian (tidak memblokir finalisasi) — mohon konfirmasi kebijakannya.';
}

/** Konversi input date sidebar "yyyy-mm-dd" → "dd/MM/yyyy" */
function isoKeTgl_(s) {
  const m = String(s || '').match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return '';
  return ('0' + m[3]).slice(-2) + '/' + ('0' + m[2]).slice(-2) + '/' + m[1];
}