/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Setup Template & Menu (v2)
 * File: Code.gs
 * Referensi aturan: PROJECT_CONTEXT.md v1.4 (N-xx, A-xx, R-xx)
 * v2: form master data direvisi — pola praktik dokter
 *     (hari/jam/shift) & kapabilitas spesialis perawat (N-05)
 * ============================================================
 */

/** Nama-nama sheet (jangan diubah tanpa menyesuaikan modul lain) */
const SHEETS = {
  MD_DOKTER: 'MD_Dokter',
  MD_PERAWAT: 'MD_Perawat',
  MD_RECEPTIONIST: 'MD_Receptionist',
  MD_SPESIALISASI: 'MD_Spesialisasi',
  IMPORT_CIS: 'Import_CIS',
  JADWAL_DOKTER: 'Jadwal_Dokter',
  PAPAN_ASISTENSI: 'Jadwal_Asistensi',
  PAPAN_RECEPTIONIST: 'Jadwal_Resepsionis',
  REQUEST_CUTI: 'Request_Cuti',
  REQUEST_JAGA: 'Request_Jaga',
  PREFERENSI_DOKTER: 'Preferensi_Dokter',
  DASHBOARD: 'Dashboard',
  CONFIG: 'Config',
  CONFIG_MODUL: 'Config_Modul'
};

/** Daftar modul web app + akses default per peran (Ya/Tidak) — dipakai seed Config_Modul & fallback */
const MODUL_DEFAULT = [
  ['dashboard', 'Dashboard', 'Ya', 'Ya', 'Ya'],
  ['master', 'Master Data', 'Ya', 'Tidak', 'Tidak'],
  ['cuti', 'Request Jadwal', 'Ya', 'Ya', 'Tidak'],
  ['pref', 'Preferensi Dokter', 'Ya', 'Ya', 'Tidak'],
  ['aksi', 'Aksi (Generate & Lock)', 'Ya', 'Ya', 'Tidak'],
  ['kalender', 'Jadwal Karyawan · Kalender', 'Ya', 'Ya', 'Ya'],
  ['kepatuhan', 'Kesesuaian (ikut menu Aksi)', 'Ya', 'Ya', 'Ya'],
  ['cetak', 'Jadwal Karyawan · List View', 'Ya', 'Ya', 'Ya']
];

/** Nama lama sheet → nama baru (untuk migrasi otomatis saat setup) */
const RENAME_SHEET = [
  ['Papan_Asistensi', 'Jadwal_Asistensi'],
  ['Papan_Receptionist', 'Jadwal_Resepsionis']
];

/** Warna tema */
const WARNA = {
  HEADER: '#1a73e8',
  HEADER_TEKS: '#ffffff',
  MASTER: '#e8f0fe',
  TERKUNCI: '#f3f3f3'
};

/** 8 spesialisasi terkonfirmasi (aturan A-03) */
const SPESIALISASI = [
  ['GP', 'General Practice'],
  ['Sp.KG', 'Spesialis Konservasi Gigi'],
  ['Sp.KGA', 'Spesialis Kedokteran Gigi Anak'],
  ['Sp.BM', 'Spesialis Bedah Mulut'],
  ['Sp.Perio', 'Spesialis Periodonsia'],
  ['Sp.Ortho', 'Spesialis Ortodonti'],
  ['Sp.Prostho', 'Spesialis Prostodonsia'],
  ['Sp.PM', 'Spesialis Penyakit Mulut']
];

/** Parameter default (semua aturan angka disimpan di sheet Config — bukan hard-coded) */
const CONFIG_DEFAULT = [
  ['PARAMETER', 'NILAI', 'KETERANGAN'],
  ['JAM_SHIFT_PAGI', '09:00-15:00', 'Jam shift pagi (terkonfirmasi)'],
  ['JAM_SHIFT_SIANG', '15:00-21:00', 'Jam shift siang (terkonfirmasi)'],
  ['LIBUR_WAJIB_PERAWAT', 4, 'N-02: jatah libur perawat per periode'],
  ['LIBUR_WAJIB_RECEPTIONIST', 4, 'R-06: jatah libur receptionist per periode'],
  ['LONGSHIFT_WAJIB_PERAWAT', 4, 'A-04: longshift wajib perawat per periode'],
  ['LONGSHIFT_WAJIB_RECEPT_SILVER', 4, 'R-07: longshift wajib receptionist Silver'],
  ['MAX_LIBUR_WEEKEND_GOLD', 1, 'N-03: batas libur weekend perawat Gold'],
  ['AMBANG_DU', 4, 'R-03/R-04: ambang jumlah DU aktif'],
  ['MIN_RECEPTIONIST_DU_KECIL', 2, 'R-03: DU aktif 1-4 → minimal receptionist'],
  ['MIN_RECEPTIONIST_DU_BESAR', 3, 'R-04: DU aktif >4 → jumlah receptionist'],
  ['HARI_WEEKEND', 'Sabtu,Minggu', 'Definisi weekend'],
  ['TANGGAL_MULAI_PERIODE', 21, 'Periode: tgl 21 s.d. 20 bulan berikutnya']
];

/** Header tiap sheet master & operasional (v2) */
const HEADER = {
  MD_DOKTER: ['ID', 'Nama Dokter', 'Spesialisasi', 'Kebutuhan Asisten', 'Hari Praktek', 'Jam Awal', 'Jam Akhir', 'Jam Shift', 'Kategori Shift', 'Aktif', 'Email'],
  MD_PERAWAT: ['ID', 'Nama Perawat', 'Level', 'Kategori', 'Kapabilitas Spesialis', 'Aktif', 'Email'],
  MD_RECEPTIONIST: ['ID', 'Nama Receptionist', 'Level', 'Aktif', 'Email'],
  MD_SPESIALISASI: ['Kode', 'Nama Spesialisasi'],
  IMPORT_CIS: ['<< PASTE DATA FORMAT 8 KOLOM DI SINI, MULAI BARIS 3: Dokter | Spesialisasi | Hari | Awal | Akhir | Jam | Shift | kebutuhan Nurse >>'],
  JADWAL_DOKTER: ['Tanggal', 'Hari', 'Shift', 'ID Dokter', 'Nama Dokter', 'Spesialisasi', 'Kebutuhan Asisten'],
  REQUEST_CUTI: ['Nama', 'Peran', 'Tanggal Mulai', 'Tanggal Akhir (opsional)', 'Keterangan', 'Catatan Aturan'],
  REQUEST_JAGA: ['Nama', 'Peran', 'Tanggal Mulai', 'Tanggal Akhir (opsional)', 'Keterangan', 'Shift', 'Catatan Aturan'],
  PREFERENSI_DOKTER: ['ID Dokter', 'Nama Dokter', 'Tipe', 'Nilai', 'Prioritas', 'Keterangan'],
  CONFIG_MODUL: ['Modul', 'Label', 'Admin', 'SPV', 'Viewer'],
  DASHBOARD: ['(Dashboard akan diisi otomatis oleh script pada Fase 5)']
};

/** Menu custom saat spreadsheet dibuka — urut sesuai alur kerja SPV */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🦷 Penjadwalan Audy')
    .addItem('🚀 Auto-Generate SEMUA (1 klik)', 'autoGenerateSemua')
    .addItem('🔍 Cek Kepatuhan Periode', 'cekKepatuhanPeriode')
    .addSeparator()
    .addSubMenu(ui.createMenu('1️⃣ Setup & Master Data')
      .addItem('⚙️ Setup Awal Template', 'setupTemplate')
      .addItem('📋 Input Master Data (form)', 'showMasterDataSidebar')
      .addItem('🧹 Rapikan Master Dokter (jam shift & spasi)', 'rapikanMasterDokter')
      .addItem('🧭 Rapikan Urutan Sheet', 'rapikanUrutanSheet'))
    .addSubMenu(ui.createMenu('2️⃣ Periode & Jadwal Dokter')
      .addItem('📅 Set Periode Aktif', 'setPeriodeAktif')
      .addItem('📝 Input Request Jadwal (Cuti/Libur)', 'showRequestCutiSidebar')
      .addItem('📄 Buka Sheet Request Jadwal', 'bukaRequestCuti')
      .addItem('⭐ Buka Sheet Preferensi Dokter', 'bukaPreferensiDokter')
      .addItem('⚡ Generate Jadwal dari Master Dokter', 'generateJadwalDariMaster')
      .addItem('🔄 Import Jadwal Dokter (dari Import_CIS)', 'importJadwalDokter'))
    .addSubMenu(ui.createMenu('3️⃣ Papan Penjadwalan (per langkah)')
      .addItem('🎲 Auto-Generate Papan Libur', 'autoGeneratePapanLibur')
      .addItem('🗓️ Buat Papan Asistensi (kosong)', 'buatPapanAsistensi')
      .addItem('🤖 Auto-Generate Asistensi', 'autoGenerateAsistensi')
      .addItem('🛎️ Generate Papan Receptionist', 'generatePapanReceptionist')
      .addItem('🏖️ Buat Papan Libur (kosong, isi manual)', 'buatPapanLibur'))
    .addSubMenu(ui.createMenu('4️⃣ Dashboard & Finalisasi')
      .addItem('📊 Perbarui Dashboard', 'buatDashboard')
      .addItem('✅ Cek Final & Kunci Jadwal', 'cekFinal')
      .addItem('🖨️ Buat Jadwal Final (print view)', 'buatJadwalFinal'))
    .addSubMenu(ui.createMenu('🌐 Web App')
      .addItem('🔗 Info & URL Web App', 'infoWebApp')
      .addItem('👤 Buka Sheet Kontrol Akses', 'bukaConfigAkses')
      .addItem('🧩 Buka Sheet Akses Modul', 'bukaConfigModul'))
    .addToUi();
}

/* ================= MODE AUTO-GENERATE SEMUA ================= */

/** Flag & penampung laporan saat 🚀 berjalan (menekan dialog per langkah) */
let MODE_OTOMATIS_ = false;
let LAPORAN_OTOMATIS_ = [];

/** Konfirmasi yang otomatis "Ya" saat mode 🚀 */
function konfirmasiOtomatis_(judul, pesan) {
  if (MODE_OTOMATIS_) return true;
  const ui = SpreadsheetApp.getUi();
  return ui.alert(judul, pesan, ui.ButtonSet.YES_NO) === ui.Button.YES;
}

/** Laporan yang dikumpulkan (mode 🚀) atau langsung tampil (mode satuan) */
function laporOtomatis_(judul, pesan) {
  if (MODE_OTOMATIS_) {
    LAPORAN_OTOMATIS_.push('══ ' + judul + ' ══\n' + pesan);
  } else {
    const ui = SpreadsheetApp.getUi();
    ui.alert(judul, pesan, ui.ButtonSet.OK);
  }
}

/**
 * Menu: 🚀 Auto-Generate SEMUA (1 klik).
 * Menjalankan berurutan: jadwal dokter dari master → papan libur →
 * papan asistensi (buat + isi) → papan receptionist (buat + isi).
 * Satu konfirmasi di awal, satu ringkasan gabungan di akhir.
 */
function autoGenerateSemua() {
  const ui = SpreadsheetApp.getUi();
  const pMulai = getConfig('PERIODE_MULAI');
  const pAkhir = getConfig('PERIODE_AKHIR');
  if (!pMulai || !pAkhir) {
    ui.alert('Periode aktif belum diset.\n\nJalankan dulu: 2️⃣ Periode & Jadwal Dokter → 📅 Set Periode Aktif.');
    return;
  }
  const konf = ui.alert('🚀 Auto-Generate SEMUA',
    'Sistem akan menjalankan berurutan untuk periode ' + pMulai + ' s.d. ' + pAkhir + ':\n\n' +
    '1. Generate Jadwal Dokter dari master\n' +
    '2. Auto-generate Papan Libur\n' +
    '3. Buat + auto-isi Papan Asistensi\n' +
    '4. Generate + auto-isi Papan Receptionist\n\n' +
    'SEMUA papan lama akan DIGANTI. Lanjutkan?',
    ui.ButtonSet.YES_NO);
  if (konf !== ui.Button.YES) return;

  MODE_OTOMATIS_ = true;
  LAPORAN_OTOMATIS_ = [];
  try {
    generateJadwalDariMaster();
    autoGeneratePapanLibur();
    buatPapanAsistensi();
    autoGenerateAsistensi();
    generatePapanReceptionist();
  } catch (err) {
    LAPORAN_OTOMATIS_.push('❌ Terhenti karena error: ' + err.message);
  } finally {
    MODE_OTOMATIS_ = false;
  }
  rapikanUrutanSheetInti_();
  ui.alert('🚀 Auto-Generate Selesai',
    LAPORAN_OTOMATIS_.join('\n\n').slice(0, 7500) +
    '\n\nLangkah akhir: review papan, lalu jalankan 🔍 Cek Kepatuhan Periode.',
    ui.ButtonSet.OK);
}

/**
 * SETUP AWAL — jalankan saat pertama kali memakai template
 * dan setiap kali upgrade versi script.
 * Membangun struktur sheet, header, dropdown, config, dan proteksi.
 * Aman dijalankan ulang: data yang ada tidak dihapus.
 */
function setupTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 0. Migrasi: rename sheet lama → nama baru (bila ada) ---
  RENAME_SHEET.forEach(function (r) {
    const lama = ss.getSheetByName(r[0]);
    const baru = ss.getSheetByName(r[1]);
    if (lama && !baru) lama.setName(r[1]);
  });

  // --- 1. Buat semua sheet sesuai urutan ---
  const urutan = [
    SHEETS.CONFIG, SHEETS.MD_DOKTER, SHEETS.MD_PERAWAT, SHEETS.MD_RECEPTIONIST,
    SHEETS.MD_SPESIALISASI, SHEETS.REQUEST_CUTI, SHEETS.REQUEST_JAGA, SHEETS.PREFERENSI_DOKTER, SHEETS.IMPORT_CIS, SHEETS.JADWAL_DOKTER,
    SHEETS.PAPAN_ASISTENSI, SHEETS.PAPAN_RECEPTIONIST, SHEETS.DASHBOARD
  ];
  urutan.forEach(function (nama, i) {
    let sh = ss.getSheetByName(nama);
    if (!sh) sh = ss.insertSheet(nama, i);
  });

  // --- 2. Config ---
  const cfg = ss.getSheetByName(SHEETS.CONFIG);
  if (cfg.getLastRow() < 2) {
    cfg.clear();
    cfg.getRange(1, 1, CONFIG_DEFAULT.length, 3).setValues(CONFIG_DEFAULT);
    formatHeader_(cfg, 3);
    cfg.setColumnWidths(1, 3, 220);
  }

  // --- 3. MD_Spesialisasi (pre-filled & dikunci) ---
  const spes = ss.getSheetByName(SHEETS.MD_SPESIALISASI);
  if (spes.getLastRow() < 2) {
    spes.clear();
    spes.getRange(1, 1, 1, 2).setValues([HEADER.MD_SPESIALISASI]);
    spes.getRange(2, 1, SPESIALISASI.length, 2).setValues(SPESIALISASI);
    formatHeader_(spes, 2);
    spes.setColumnWidth(2, 260);
    spes.getRange(1, 1, spes.getMaxRows(), 2).setBackground(WARNA.TERKUNCI);
    const prot = spes.protect().setDescription('Master spesialisasi — dikunci (acuan A-03)');
    prot.setWarningOnly(true); // ganti ke false + atur editor bila ingin kunci penuh
  }

  // --- 4. Sheet master data (header selalu disamakan dgn versi terbaru) ---
  setupSheetMaster_(ss, SHEETS.MD_DOKTER, HEADER.MD_DOKTER);
  setupSheetMaster_(ss, SHEETS.MD_PERAWAT, HEADER.MD_PERAWAT);
  setupSheetMaster_(ss, SHEETS.MD_RECEPTIONIST, HEADER.MD_RECEPTIONIST);

  // --- 5. Dropdown / data validation (anti-typo) ---
  pasangValidasi_(ss);

  // --- 5a. Sheet Config_Akses (kontrol akses web app) ---
  let ak = ss.getSheetByName('Config_Akses');
  if (!ak) ak = ss.insertSheet('Config_Akses');
  if (ak.getRange(1, 1).isBlank()) {
    ak.getRange(1, 1, 1, 3).setValues([['Email', 'Peran', 'Nama']]);
    formatHeader_(ak, 3);
    // Seed pemilik sebagai Admin
    const ownerEmail = Session.getEffectiveUser().getEmail();
    if (ownerEmail) ak.getRange(2, 1, 1, 3).setValues([[ownerEmail, 'Admin', 'Pemilik']]);
    const rulePeran = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Admin', 'SPV', 'Viewer'], true).setAllowInvalid(false).build();
    ak.getRange(2, 2, 300, 1).setDataValidation(rulePeran);
    ak.setColumnWidth(1, 240); ak.setColumnWidth(3, 160);
  }

  // --- 5a-2. Sheet Config_Modul (akses modul per peran untuk web app) ---
  let cm = ss.getSheetByName(SHEETS.CONFIG_MODUL);
  if (!cm) cm = ss.insertSheet(SHEETS.CONFIG_MODUL);
  if (cm.getRange(1, 1).isBlank()) {
    cm.getRange(1, 1, 1, HEADER.CONFIG_MODUL.length).setValues([HEADER.CONFIG_MODUL]);
    formatHeader_(cm, HEADER.CONFIG_MODUL.length);
    cm.getRange(2, 1, MODUL_DEFAULT.length, 5).setValues(MODUL_DEFAULT);
    const ruleYT = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Ya', 'Tidak'], true).setAllowInvalid(false).build();
    cm.getRange(2, 3, MODUL_DEFAULT.length, 3).setDataValidation(ruleYT);
    cm.setColumnWidth(2, 200);
    cm.getRange(1, 1, cm.getMaxRows(), 2).setBackground(WARNA.TERKUNCI);
  }

  // --- 5b. Sheet Request Cuti/Libur ---
  const rc = ss.getSheetByName(SHEETS.REQUEST_CUTI);
  rc.getRange(1, 1, 1, HEADER.REQUEST_CUTI.length).setValues([HEADER.REQUEST_CUTI]); // selalu samakan header (mendukung kolom Catatan Aturan)
  formatHeader_(rc, HEADER.REQUEST_CUTI.length);
  rc.getRange('C2:D').setNumberFormat('@'); // teks agar tanggal tak berubah format
  rc.setColumnWidth(1, 200); rc.setColumnWidth(5, 220); rc.setColumnWidth(6, 280);
  pasangValidasiRequest_(ss);

  // --- 5b-2. Sheet Request Jaga (permintaan dijadwalkan jaga) ---
  let rj = ss.getSheetByName(SHEETS.REQUEST_JAGA);
  if (!rj) rj = ss.insertSheet(SHEETS.REQUEST_JAGA);
  rj.getRange(1, 1, 1, HEADER.REQUEST_JAGA.length).setValues([HEADER.REQUEST_JAGA]); // selalu samakan header (mendukung kolom Shift)
  formatHeader_(rj, HEADER.REQUEST_JAGA.length);
  rj.getRange('C2:D').setNumberFormat('@');
  rj.setColumnWidth(1, 200); rj.setColumnWidth(5, 220); rj.setColumnWidth(7, 280);
  if (typeof pasangValidasiRequestJaga_ === 'function') pasangValidasiRequestJaga_(ss);

  // --- 5c. Sheet Preferensi Dokter ---
  const pref = ss.getSheetByName(SHEETS.PREFERENSI_DOKTER);
  if (pref.getRange(1, 1).isBlank()) {
    pref.getRange(1, 1, 1, HEADER.PREFERENSI_DOKTER.length).setValues([HEADER.PREFERENSI_DOKTER]);
    formatHeader_(pref, HEADER.PREFERENSI_DOKTER.length);
    pref.setColumnWidth(2, 200); pref.setColumnWidth(4, 160); pref.setColumnWidth(6, 220);
  }
  pasangValidasiPreferensi_(ss);

  // --- 6. Sheet operasional ---
  const imp = ss.getSheetByName(SHEETS.IMPORT_CIS);
  if (imp.getLastRow() < 1 || imp.getRange(1, 1).isBlank()) {
    imp.getRange(1, 1).setValue(HEADER.IMPORT_CIS[0])
      .setFontWeight('bold').setFontColor('#b45309');
    imp.getRange(2, 1).setValue('Setelah paste: menu 🦷 Penjadwalan Audy → 📅 Set Periode Aktif → 🔄 Import Jadwal Dokter.');
  }
  const jd = ss.getSheetByName(SHEETS.JADWAL_DOKTER);
  if (jd.getRange(1, 1).isBlank()) {
    jd.getRange(1, 1, 1, HEADER.JADWAL_DOKTER.length).setValues([HEADER.JADWAL_DOKTER]);
    formatHeader_(jd, HEADER.JADWAL_DOKTER.length);
  }
  const db = ss.getSheetByName(SHEETS.DASHBOARD);
  if (db.getRange(1, 1).isBlank()) {
    db.getRange(1, 1).setValue(HEADER.DASHBOARD[0]).setFontStyle('italic');
  }

  // --- 7. Hapus sheet default kosong bawaan Google + susun urutan tab ---
  const def = ss.getSheetByName('Sheet1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  rapikanUrutanSheetInti_();

  SpreadsheetApp.getUi().alert(
    'Setup selesai ✅',
    'Struktur template berhasil dibangun/diperbarui (v2).\n\nLangkah berikutnya: buka menu 🦷 Penjadwalan Audy → 📋 Input Master Data untuk mengisi data dokter, perawat, dan receptionist.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/** Helper: siapkan sheet master dengan header terformat.
 *  Header selalu disamakan dengan definisi terbaru (mendukung upgrade versi). */
function setupSheetMaster_(ss, nama, header) {
  const sh = ss.getSheetByName(nama);
  const cur = sh.getRange(1, 1, 1, header.length).getValues()[0];
  const sama = header.every(function (h, i) { return cur[i] === h; });
  if (!sama) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    formatHeader_(sh, header.length);
    sh.setColumnWidth(2, 220);
  }
}

/** Helper: format baris header */
function formatHeader_(sh, nCols) {
  sh.getRange(1, 1, 1, nCols)
    .setBackground(WARNA.HEADER)
    .setFontColor(WARNA.HEADER_TEKS)
    .setFontWeight('bold');
  sh.setFrozenRows(1);
}

/** Helper: pasang semua dropdown data-validation di sheet master.
 *  PENTING: validasi lama dibersihkan dulu — mencegah aturan versi
 *  sebelumnya menempel di kolom yang sudah berubah fungsi
 *  (mis. kolom 5 MD_Perawat: dulu "Aktif" Ya/Tidak, kini "Kapabilitas"). */
function pasangValidasi_(ss) {
  const MAXR = 500; // kapasitas baris master data

  // Bersihkan seluruh validasi lama di 3 sheet master
  [SHEETS.MD_DOKTER, SHEETS.MD_PERAWAT, SHEETS.MD_RECEPTIONIST].forEach(function (nama) {
    const sh = ss.getSheetByName(nama);
    sh.getRange(2, 1, MAXR, sh.getMaxColumns()).clearDataValidations();
  });

  const ruleDari = function (values) {
    return SpreadsheetApp.newDataValidation()
      .requireValueInList(values, true).setAllowInvalid(false).build();
  };
  const ruleYaTidak = ruleDari(['Ya', 'Tidak']);
  const ruleLevel = ruleDari(['Gold', 'Silver']);      // N-01, R-08
  const ruleKategori = ruleDari(['Existing', 'New']);  // N-04

  // Dokter: spesialisasi dari MD_Spesialisasi (range), kebutuhan asisten 1-5,
  // hari praktek, jam shift (dari Config), kategori shift
  const spesRange = ss.getSheetByName(SHEETS.MD_SPESIALISASI)
    .getRange(2, 1, SPESIALISASI.length, 1);
  const ruleSpes = SpreadsheetApp.newDataValidation()
    .requireValueInRange(spesRange, true).setAllowInvalid(false).build();
  const ruleAsisten = ruleDari(['1', '2', '3', '4', '5']);
  const ruleHari = ruleDari(['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu', 'Minggu']);
  const jamPagi = String(getConfig('JAM_SHIFT_PAGI') || '09:00-15:00');
  const jamSiang = String(getConfig('JAM_SHIFT_SIANG') || '15:00-21:00');
  const ruleJamShift = ruleDari([jamPagi, jamSiang]);
  const ruleKatShift = ruleDari(['Shift Pagi', 'Shift Siang']);

  const dok = ss.getSheetByName(SHEETS.MD_DOKTER);
  dok.getRange(2, 3, MAXR).setDataValidation(ruleSpes);      // Spesialisasi
  dok.getRange(2, 4, MAXR).setDataValidation(ruleAsisten);   // Kebutuhan Asisten (A-05)
  dok.getRange(2, 5, MAXR).setDataValidation(ruleHari);      // Hari Praktek
  dok.getRange(2, 8, MAXR).setDataValidation(ruleJamShift);  // Jam Shift (dari Config)
  dok.getRange(2, 9, MAXR).setDataValidation(ruleKatShift);  // Kategori Shift
  dok.getRange(2, 10, MAXR).setDataValidation(ruleYaTidak);  // Aktif

  // Kolom 5 (Kapabilitas): dropdown spesialis dengan multi-select via onEdit.
  // allowInvalid = true agar sel boleh berisi gabungan beberapa kode (dipisah koma).
  const ruleSpesMulti = SpreadsheetApp.newDataValidation()
    .requireValueInList(SPESIALISASI.map(function (s) { return s[0]; }), true)
    .setAllowInvalid(true).build();

  const prw = ss.getSheetByName(SHEETS.MD_PERAWAT);
  prw.getRange(2, 3, MAXR).setDataValidation(ruleLevel);      // Level
  prw.getRange(2, 4, MAXR).setDataValidation(ruleKategori);   // Kategori
  prw.getRange(2, 5, MAXR).setDataValidation(ruleSpesMulti);  // Kapabilitas (multi-select)
  prw.getRange(2, 6, MAXR).setDataValidation(ruleYaTidak);    // Aktif

  const rcp = ss.getSheetByName(SHEETS.MD_RECEPTIONIST);
  rcp.getRange(2, 3, MAXR).setDataValidation(ruleLevel);    // Level (R-08)
  rcp.getRange(2, 4, MAXR).setDataValidation(ruleYaTidak);  // Aktif
}

/** Dropdown Nama & Peran di sheet Request_Cuti (dipanggil setup & saat sheet dibuka) */
function pasangValidasiRequest_(ss) {
  const rc = ss.getSheetByName(SHEETS.REQUEST_CUTI);
  if (!rc) return;
  const MAXR = 300;
  // Daftar nama = semua staf aktif (dokter + perawat + receptionist)
  const nama = getMasterData('dokter').filter(function (d) { return d.aktif !== 'Tidak'; }).map(function (d) { return d.nama; })
    .concat(getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; }).map(function (p) { return p.nama; }))
    .concat(getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; }).map(function (r) { return r.nama; }));
  // Nama bisa sama antar peran; buat unik
  const unik = [];
  nama.forEach(function (n) { if (unik.indexOf(n) === -1) unik.push(n); });
  if (unik.length) {
    const ruleNama = SpreadsheetApp.newDataValidation()
      .requireValueInList(unik, true).setAllowInvalid(true).build();
    rc.getRange(2, 1, MAXR, 1).setDataValidation(ruleNama);
  }
  const rulePeran = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Dokter', 'Perawat', 'Receptionist'], true).setAllowInvalid(true).build();
  rc.getRange(2, 2, MAXR, 1).setDataValidation(rulePeran);
}

/**
 * Menu: 🧭 Rapikan Urutan Sheet.
 * Menyusun tab sesuai frekuensi pakai + warna per kelompok:
 *  🔵 Papan kerja harian → 🟢 Rekap → 🟡 Jadwal & Master → 🟠 Import → ⚫ Config
 */
function rapikanUrutanSheet() {
  rapikanUrutanSheetInti_();
  laporOtomatis_('Urutan Sheet',
    'Tab tersusun ulang sesuai alur kerja:\n' +
    '🔵 Papan kerja → 🟢 Rekap → 🟡 Jadwal & Master → 🟠 Import → ⚫ Config');
}

/** Inti penyusunan tab (tanpa dialog — dipakai setup & auto-generate) */
function rapikanUrutanSheetInti_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const urutan = [
    // Papan kerja harian (biru)
    ['Jadwal_Asistensi', '#1a73e8'],
    ['Jadwal_Resepsionis', '#1a73e8'],
    ['Papan_Libur', '#1a73e8'],
    // Dashboard, hasil audit & jadwal jadi (hijau)
    ['Dashboard', '#188038'],
    ['Jadwal_Final', '#188038'],
    ['Rekap_Kepatuhan', '#188038'],
    // Jadwal periode & request (kuning tua)
    ['Jadwal_Dokter', '#f9ab00'],
    ['Request_Cuti', '#f9ab00'],
    ['Request_Jaga', '#f9ab00'],
    ['Preferensi_Dokter', '#f9ab00'],
    // Master data (kuning)
    ['MD_Dokter', '#fbc02d'],
    ['MD_Perawat', '#fbc02d'],
    ['MD_Receptionist', '#fbc02d'],
    ['MD_Spesialisasi', '#fbc02d'],
    ['Map_Spesialisasi', '#fbc02d'],
    // Area import (oranye)
    ['Import_CIS', '#e8710a'],
    ['Import_Perawat', '#e8710a'],
    // Pengaturan (abu-abu)
    ['Config_Akses', '#5f6368'],
    ['Config_Modul', '#5f6368'],
    ['Config', '#5f6368']
  ];
  let pos = 1;
  urutan.forEach(function (u) {
    const sh = ss.getSheetByName(u[0]);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(pos);
    sh.setTabColor(u[1]);
    pos++;
  });
  // Kembali fokus ke papan utama
  const utama = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (utama) ss.setActiveSheet(utama);
}

/** Menu: 🔗 Info & URL Web App */
function infoWebApp() {
  const ui = SpreadsheetApp.getUi();
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  ui.alert('🌐 Web App Penjadwalan',
    (url ? 'URL Web App aktif:\n' + url + '\n\n' : 'Web App belum di-deploy.\n\n') +
    'Cara deploy (sekali):\n' +
    '1. Extensions → Apps Script → Deploy → New deployment\n' +
    '2. Type: Web app\n' +
    '3. Execute as: Me (pemilik)\n' +
    '4. Who has access: Anyone with Google account\n' +
    '5. Deploy → salin URL, bagikan ke SPV\n\n' +
    'Kontrol siapa yang boleh akses & perannya di sheet "Config_Akses" (Email | Peran | Nama). ' +
    'Peran: Admin (semua), SPV (jalankan alur & edit jadwal), Viewer (lihat saja).',
    ui.ButtonSet.OK);
}

/** Menu: 👤 Buka Sheet Kontrol Akses */
function bukaConfigAkses() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ak = ss.getSheetByName('Config_Akses');
  if (!ak) { setupTemplate(); ak = ss.getSheetByName('Config_Akses'); }
  ss.setActiveSheet(ak);
}

/** Menu: 🧩 Buka Sheet Akses Modul */
function bukaConfigModul() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let cm = ss.getSheetByName(SHEETS.CONFIG_MODUL);
  if (!cm) { setupTemplate(); cm = ss.getSheetByName(SHEETS.CONFIG_MODUL); }
  ss.setActiveSheet(cm);
  SpreadsheetApp.getUi().alert('🧩 Akses Modul',
    'Atur modul mana yang boleh diakses tiap peran (Ya/Tidak). Bisa juga diatur dari tab "Pengaturan" di web app (khusus Admin). Perubahan langsung berlaku setelah refresh web app.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/** Baca 1 nilai parameter dari sheet Config (dipakai modul lain) */
function getConfig(param) {
  const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  const data = cfg.getRange(2, 1, cfg.getLastRow() - 1, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === param) return data[i][1];
  }
  return null;
}