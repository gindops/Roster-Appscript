/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Fase 3: Papan Penjadwalan
 * File: Scheduler.gs
 *
 * - Papan_Asistensi: grid tanggal × shift × dokter × slot,
 *   slot mengikuti kebutuhan asisten per dokter (A-05).
 *   SPV assign perawat via dropdown; validasi oleh Validator.gs.
 * - Papan_Libur: grid tanggal × kolom libur (keputusan desain),
 *   dropdown berisi perawat + receptionist aktif — dasar
 *   perhitungan N-02/N-03 (perawat) dan R-06 (receptionist).
 * ============================================================
 */

/** Kolom Papan_Asistensi */
const PA_HEADER = ['Tanggal', 'Hari', 'Shift', 'Nama Dokter', 'Spesialisasi', 'Jumlah Asisten', 'Perawat'];
const PA_KOL_PERAWAT = 7;

/** Papan_Libur: jumlah kolom slot libur per tanggal */
const PL_MAX_SLOT = 8;

/* ================= PAPAN ASISTENSI ================= */

/** Menu: 🗓️ Buat Papan Asistensi */
function buatPapanAsistensi() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const jd = ss.getSheetByName(SHEETS.JADWAL_DOKTER);
  if (!jd || jd.getLastRow() < 2) {
    laporOtomatis_('Papan Asistensi', 'Jadwal_Dokter masih kosong.\nJalankan dulu ⚡ Generate Jadwal dari Master atau 🔄 Import.');
    return;
  }

  if (!konfirmasiOtomatis_(
    'Buat Papan Asistensi',
    'Papan akan dibuat ulang dari Jadwal_Dokter.\nAssignment perawat yang sudah ada di papan lama akan HILANG. Lanjutkan?'
  )) return;

  // Baca jadwal dokter: Tanggal|Hari|Shift|ID|Nama|Spes|Kebutuhan
  const data = jd.getRange(2, 1, jd.getLastRow() - 1, 7).getValues()
    .filter(function (r) { return r[0]; });

  // Bangun baris papan: 1 baris per slot asisten
  const rows = [];
  data.forEach(function (r) {
    const n = parseInt(r[6], 10) || 1;
    for (let s = 1; s <= n; s++) {
      rows.push([r[0], r[1], r[2], r[4], r[5], s, '']);
    }
  });

  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  pa.clear();
  pa.getRange(1, 1, pa.getMaxRows(), pa.getMaxColumns()).clearDataValidations(); // buang dropdown lama
  pa.getRange(1, 1, 1, PA_HEADER.length).setValues([PA_HEADER]);
  formatHeader_(pa, PA_HEADER.length);
  if (rows.length) {
    pa.getRange(2, 1, rows.length, 1).setNumberFormat('@'); // kolom Tanggal jadi TEKS
    pa.getRange(2, 1, rows.length, PA_HEADER.length).setValues(rows);
  }

  // Dropdown perawat aktif di kolom Perawat
  const namaPerawat = getMasterData('perawat')
    .filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) { return p.nama; });
  if (namaPerawat.length) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(namaPerawat, true).setAllowInvalid(true).build();
    pa.getRange(2, PA_KOL_PERAWAT, Math.max(rows.length, 1)).setDataValidation(rule);
  }

  // Pewarnaan selang-seling per tanggal + garis pemisah antar tanggal
  warnaiPerTanggal_(pa, rows.length, 1);
  pa.setColumnWidth(4, 200);
  pa.setColumnWidth(5, 110);
  pa.setFrozenRows(1);

  laporOtomatis_(
    'Papan Asistensi',
    rows.length + ' slot asistensi dibuat dari ' + data.length + ' jadwal dokter.\n' +
    'Slot = kursi asisten: dokter berkebutuhan 2 asisten muncul 2 baris (Slot 1 & 2) per shift.'
  );
}

/* ================= PAPAN LIBUR ================= */

/** Menu: 🏖️ Buat Papan Libur */
function buatPapanLibur() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const pMulai = parseTgl_(getConfig('PERIODE_MULAI'));
  const pAkhir = parseTgl_(getConfig('PERIODE_AKHIR'));
  if (!pMulai || !pAkhir) {
    ui.alert('Periode aktif belum diset.\nJalankan dulu 📅 Set Periode Aktif.');
    return;
  }

  let pl = ss.getSheetByName('Papan_Libur');
  if (pl && pl.getLastRow() > 1) {
    const konf = ui.alert('Buat Papan Libur',
      'Papan libur lama akan DIGANTI (isian libur hilang). Lanjutkan?', ui.ButtonSet.YES_NO);
    if (konf !== ui.Button.YES) return;
  }
  if (!pl) pl = ss.insertSheet('Papan_Libur');
  pl.clear();

  // Header
  const header = ['Tanggal', 'Hari'];
  for (let i = 1; i <= PL_MAX_SLOT; i++) header.push('Libur ' + i);
  pl.getRange(1, 1, 1, header.length).setValues([header]);
  formatHeader_(pl, header.length);

  // Baris per tanggal periode
  const weekend = String(getConfig('HARI_WEEKEND') || 'Sabtu,Minggu').split(',')
    .map(function (h) { return h.trim().toLowerCase(); });
  const rows = [];
  for (let d = new Date(pMulai); d <= pAkhir; d.setDate(d.getDate() + 1)) {
    const hari = NAMA_HARI[d.getDay()].charAt(0) + NAMA_HARI[d.getDay()].slice(1).toLowerCase();
    const baris = [formatTgl_(d), hari];
    for (let i = 0; i < PL_MAX_SLOT; i++) baris.push('');
    rows.push(baris);
  }
  pl.getRange(2, 1, rows.length, 1).setNumberFormat('@'); // kolom Tanggal jadi TEKS
  pl.getRange(2, 1, rows.length, header.length).setValues(rows);

  // Dropdown: gabungan perawat + receptionist aktif
  const namaStaf = getMasterData('perawat')
    .filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) { return p.nama; })
    .concat(getMasterData('receptionist')
      .filter(function (r) { return r.aktif !== 'Tidak'; })
      .map(function (r) { return r.nama; }));
  if (namaStaf.length) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(namaStaf, true).setAllowInvalid(true).build();
    pl.getRange(2, 3, rows.length, PL_MAX_SLOT).setDataValidation(rule);
  }

  // Tandai baris weekend
  rows.forEach(function (r, i) {
    if (weekend.indexOf(String(r[1]).toLowerCase()) !== -1) {
      pl.getRange(i + 2, 1, 1, header.length).setBackground('#fff3e0');
    }
  });
  pl.setFrozenRows(1);
  pl.setFrozenColumns(2);

  ui.alert(
    'Papan Libur siap ✅',
    rows.length + ' tanggal periode dibuat.\n\n' +
    'Isi nama perawat/receptionist yang libur per tanggal via dropdown.\n' +
    'Aturan yang divalidasi: 4x libur per periode (N-02 perawat, R-06 receptionist), ' +
    'libur weekend hanya perawat Gold maks. 1x (N-03).',
    ui.ButtonSet.OK
  );
}

/* ================= HELPER ================= */

/** Warnai baris selang-seling berdasarkan perubahan nilai kolom tanggal */
function warnaiPerTanggal_(sh, nRows, kolTanggal) {
  if (!nRows) return;
  const tgl = sh.getRange(2, kolTanggal, nRows, 1).getValues();
  let ganjil = false;
  let prev = null;
  const bgs = [];
  for (let i = 0; i < nRows; i++) {
    const t = String(tgl[i][0]);
    if (t !== prev) { ganjil = !ganjil; prev = t; }
    bgs.push([ganjil ? '#f0f6ff' : '#ffffff']);
  }
  // Terapkan ke seluruh lebar papan (kecuali kolom Perawat agar warna validasi tak tertimpa)
  const nCols = sh.getLastColumn();
  for (let c = 1; c <= nCols; c++) {
    if (c === PA_KOL_PERAWAT) continue;
    sh.getRange(2, c, nRows, 1).setBackgrounds(bgs);
  }
}