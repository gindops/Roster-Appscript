/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Fase 2: Import & Expand Jadwal Dokter (v2)
 * File: ImportCIS.gs
 *
 * Dua sumber Jadwal_Dokter (keputusan #19):
 *   A. Import dari Import_CIS — export CIS yang dirapikan SPV
 *      ke format 8 kolom, di-paste ke sheet Import_CIS.
 *   B. Generate dari Master — pola praktik langsung dari MD_Dokter.
 *
 * Format 8 kolom (mulai baris 3 di Import_CIS):
 *   Dokter | Spesialisasi | Hari | Awal | Akhir | Jam | Shift | kebutuhan Nurse
 *   contoh: drg. Hirania, Sp.KG | Sp.KG | 1. SENIN | 10.00 | 14.00 | 09.00 - 15.00 | Pagi | 2
 * ============================================================
 */

/** Baris pertama data di sheet Import_CIS (baris 1-2 = petunjuk) */
const IMPORT_BARIS_MULAI = 3;

/** Nama hari Indonesia sesuai indeks getDay() JavaScript (0=Minggu) */
const NAMA_HARI = ['MINGGU', 'SENIN', 'SELASA', 'RABU', 'KAMIS', 'JUMAT', 'SABTU'];

/** Alias kode spesialisasi yang sering muncul di data CIS → kode master (A-03) */
const ALIAS_SPESIALISASI_DEFAULT = [
  ['Sp.Pros', 'Sp.Prostho'],
  ['Sp.Prost', 'Sp.Prostho'],
  ['Sp.Orth', 'Sp.Ortho'],
  ['Sp.Ort', 'Sp.Ortho'],
  ['Sp.Kg', 'Sp.KG'],
  ['Sp.Kga', 'Sp.KGA'],
  ['Sp.Bm', 'Sp.BM'],
  ['Sp.Pm', 'Sp.PM'],
  ['Sp.Perio', 'Sp.Perio'],
  ['Gp', 'GP']
];

/* ================= PERIODE AKTIF ================= */

/**
 * Set periode aktif (menu: 📅 Set Periode Aktif).
 * SPV memasukkan bulan mulai, sistem menghitung 21 bln tsb s.d. 20 bln berikutnya.
 */
function setPeriodeAktif() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    'Set Periode Aktif',
    'Masukkan bulan & tahun MULAI periode (format: MM/YYYY).\nContoh: 07/2026 → periode 21 Jul 2026 s.d. 20 Agu 2026.',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const m = res.getResponseText().trim().match(/^(\d{1,2})[\/\-](\d{4})$/);
  if (!m) { ui.alert('Format tidak dikenali. Gunakan MM/YYYY, contoh: 07/2026.'); return; }
  const bulan = parseInt(m[1], 10), tahun = parseInt(m[2], 10);
  if (bulan < 1 || bulan > 12) { ui.alert('Bulan harus 1-12.'); return; }

  const tglMulai = Number(getConfig('TANGGAL_MULAI_PERIODE')) || 21;
  const mulai = new Date(tahun, bulan - 1, tglMulai);
  const akhir = new Date(tahun, bulan, tglMulai - 1); // tgl 20 bulan berikutnya

  setConfig_('PERIODE_MULAI', formatTgl_(mulai), 'Tanggal mulai periode aktif');
  setConfig_('PERIODE_AKHIR', formatTgl_(akhir), 'Tanggal akhir periode aktif');
  ui.alert('Periode aktif: ' + formatTgl_(mulai) + ' s.d. ' + formatTgl_(akhir));
}

/** Tulis/update 1 parameter di sheet Config */
function setConfig_(param, nilai, ket) {
  const cfg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
  const data = cfg.getRange(2, 1, Math.max(cfg.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === param) { cfg.getRange(i + 2, 2).setValue(nilai); return; }
  }
  cfg.appendRow([param, nilai, ket || '']);
}

/* ================= SUMBER A: IMPORT DARI IMPORT_CIS ================= */

/** Menu: 🔄 Import Jadwal Dokter (dari Import_CIS) */
function importJadwalDokter() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Prasyarat: periode aktif ---
  const pMulai = parseTgl_(getConfig('PERIODE_MULAI'));
  const pAkhir = parseTgl_(getConfig('PERIODE_AKHIR'));
  if (!pMulai || !pAkhir) {
    ui.alert('Periode aktif belum diset.\nJalankan dulu: 🦷 Penjadwalan Audy → 📅 Set Periode Aktif.');
    return;
  }

  // --- Baca data mentah ---
  const imp = ss.getSheetByName(SHEETS.IMPORT_CIS);
  const lastRow = imp.getLastRow();
  if (lastRow < IMPORT_BARIS_MULAI) {
    ui.alert('Tidak ada data.\nPaste data format 8 kolom ke sheet Import_CIS mulai baris ' + IMPORT_BARIS_MULAI + '.');
    return;
  }
  let raw = imp.getRange(IMPORT_BARIS_MULAI, 1, lastRow - IMPORT_BARIS_MULAI + 1, 8).getValues();

  // Lewati baris header jika ikut ter-paste
  raw = raw.filter(function (r) {
    return String(r[0]).trim() && String(r[0]).trim().toLowerCase() !== 'dokter';
  });
  if (!raw.length) { ui.alert('Tidak ada baris data valid di Import_CIS.'); return; }

  // --- Konfirmasi sebelum menimpa hasil lama ---
  const konf = ui.alert(
    'Import Jadwal Dokter',
    'Akan memproses ' + raw.length + ' baris pola mingguan menjadi jadwal per tanggal untuk periode ' +
    formatTgl_(pMulai) + ' s.d. ' + formatTgl_(pAkhir) + '.\n\nIsi lama sheet Jadwal_Dokter akan DIGANTI. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  if (konf !== ui.Button.YES) return;

  // --- Siapkan referensi ---
  const aliasMap = getAliasMap_(ss);
  const masterSpes = getSpesialisasiList().map(function (s) { return s.kode; });
  const dokterMap = buatIndexDokter_();  // nama lower → {id, ...}
  // Jam shift kanonik dari Config — JANGAN pakai teks mentah kolom Jam
  // (data CIS memakai format "09.00 - 15.00" + spasi tak konsisten)
  const jamPagiCfg = String(getConfig('JAM_SHIFT_PAGI') || '09:00-15:00');
  const jamSiangCfg = String(getConfig('JAM_SHIFT_SIANG') || '15:00-21:00');
  const errors = [];
  const polaValid = [];
  let dokterBaru = 0;

  // --- Validasi & normalisasi tiap baris pola ---
  raw.forEach(function (r, idx) {
    const barisSheet = idx + IMPORT_BARIS_MULAI;
    const nama = String(r[0]).trim();
    let spes = String(r[1]).trim();
    const hariRaw = String(r[2]).trim();
    let shift = String(r[6]).trim();
    const nurse = parseInt(r[7], 10);

    // Normalisasi spesialisasi via alias
    if (masterSpes.indexOf(spes) === -1) {
      const alias = aliasMap[spes.toLowerCase()];
      if (alias) spes = alias;
      else { errors.push('Baris ' + barisSheet + ': spesialisasi "' + spes + '" tidak dikenal.'); return; }
    }

    // Normalisasi hari: "1. SENIN" → indeks getDay()
    const hariNama = hariRaw.replace(/^[\d\.\s]+/, '').toUpperCase().trim();
    const hariIdx = NAMA_HARI.indexOf(hariNama);
    if (hariIdx === -1) { errors.push('Baris ' + barisSheet + ': hari "' + hariRaw + '" tidak dikenali.'); return; }

    // Normalisasi shift
    shift = shift.charAt(0).toUpperCase() + shift.slice(1).toLowerCase();
    if (shift !== 'Pagi' && shift !== 'Siang') {
      errors.push('Baris ' + barisSheet + ': shift "' + r[6] + '" harus Pagi atau Siang.'); return;
    }

    if (!nurse || nurse < 1) {
      errors.push('Baris ' + barisSheet + ': kebutuhan Nurse tidak valid ("' + r[7] + '").'); return;
    }

    // Dokter/pola belum ada di master → daftarkan otomatis lengkap dengan
    // pola praktiknya (nama sama → ID sama; duplikat nama+hari+shift dilewati)
    let dok = dokterMap[nama.toLowerCase()];
    const hasil = addMasterData('dokter', {
      nama: nama,
      spesialisasi: spes,
      kebutuhanAsisten: String(nurse),
      hariPraktek: hariNama.charAt(0) + hariNama.slice(1).toLowerCase(),
      jamAwal: String(r[3]).trim(),
      jamAkhir: String(r[4]).trim(),
      jamShift: shift === 'Pagi' ? jamPagiCfg : jamSiangCfg,
      kategoriShift: 'Shift ' + shift
    });
    if (hasil.ok) {
      if (!dok) { dokterBaru++; }
      dok = { id: hasil.id };
      dokterMap[nama.toLowerCase()] = dok;
    } else if (!dok) {
      // Gagal daftar dan dokter memang belum dikenal → baris tidak bisa diproses
      errors.push('Baris ' + barisSheet + ': gagal daftar dokter "' + nama + '" — ' + hasil.pesan);
      return;
    }
    // Bila gagal karena "jadwal duplikat" tapi dokter sudah ada → lanjut pakai ID lama

    polaValid.push({ idDokter: dok.id, nama: nama, spes: spes, hariIdx: hariIdx, shift: shift, nurse: nurse });
  });

  // --- Expand, saring cuti dokter, & tulis ---
  const hasilRaw = expandPola_(polaValid, pMulai, pAkhir);
  const saring = saringDokterCuti_(hasilRaw);
  const hasil = saring.rows;
  tulisJadwal_(ss, hasil);

  // --- Ringkasan ---
  let pesan = 'Import selesai ✅\n\n' +
    '• Pola valid: ' + polaValid.length + ' dari ' + raw.length + ' baris\n' +
    '• Jadwal per tanggal dihasilkan: ' + hasil.length + ' baris\n' +
    (saring.dibuang ? '• Dibuang karena dokter cuti (Request_Cuti): ' + saring.dibuang + ' baris\n' : '') +
    '• Dokter baru didaftarkan otomatis ke MD_Dokter: ' + dokterBaru + '\n';
  if (errors.length) {
    pesan += '\n⚠️ ' + errors.length + ' baris DILEWATI:\n' +
      errors.slice(0, 15).join('\n') +
      (errors.length > 15 ? '\n… dan ' + (errors.length - 15) + ' error lainnya.' : '');
  }
  ui.alert('Hasil Import', pesan, ui.ButtonSet.OK);
}

/* ================= SUMBER B: GENERATE DARI MASTER ================= */

/**
 * Menu: ⚡ Generate Jadwal dari Master Dokter.
 * Membaca pola praktik langsung dari MD_Dokter (baris Aktif=Ya)
 * lalu meng-expand ke tanggal aktual periode aktif → Jadwal_Dokter.
 */
function generateJadwalDariMaster() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const pMulai = parseTgl_(getConfig('PERIODE_MULAI'));
  const pAkhir = parseTgl_(getConfig('PERIODE_AKHIR'));
  if (!pMulai || !pAkhir) {
    ui.alert('Periode aktif belum diset.\nJalankan dulu: 🦷 Penjadwalan Audy → 📅 Set Periode Aktif.');
    return;
  }

  const rows = getMasterData('dokter').filter(function (d) { return d.aktif !== 'Tidak'; });
  if (!rows.length) { laporOtomatis_('Jadwal Dokter', 'MD_Dokter kosong. Isi master data dokter dulu.'); return; }

  const errors = [];
  const polaValid = [];
  rows.forEach(function (d) {
    const hariIdx = NAMA_HARI.indexOf(String(d.hariPraktek || '').toUpperCase().trim());
    if (hariIdx === -1) {
      errors.push(d.id + ' (' + d.nama + '): Hari Praktek "' + d.hariPraktek + '" tidak valid.'); return;
    }
    const shift = String(d.kategoriShift || '').replace(/^Shift\s*/i, '').trim();
    if (shift !== 'Pagi' && shift !== 'Siang') {
      errors.push(d.id + ' (' + d.nama + '): Kategori Shift "' + d.kategoriShift + '" tidak valid.'); return;
    }
    const nurse = parseInt(d.kebutuhanAsisten, 10);
    if (!nurse || nurse < 1) {
      errors.push(d.id + ' (' + d.nama + '): Kebutuhan Asisten tidak valid.'); return;
    }
    polaValid.push({ idDokter: d.id, nama: d.nama, spes: d.spesialisasi, hariIdx: hariIdx, shift: shift, nurse: nurse });
  });

  if (!polaValid.length) {
    laporOtomatis_('Jadwal Dokter', 'Tidak ada pola valid di MD_Dokter.\n' + errors.join('\n')); return;
  }

  if (!konfirmasiOtomatis_(
    'Generate Jadwal dari Master',
    'Akan meng-expand ' + polaValid.length + ' pola praktik dari MD_Dokter untuk periode ' +
    formatTgl_(pMulai) + ' s.d. ' + formatTgl_(pAkhir) + '.\n\nIsi lama sheet Jadwal_Dokter akan DIGANTI. Lanjutkan?'
  )) return;

  const hasilRaw = expandPola_(polaValid, pMulai, pAkhir);
  const saring = saringDokterCuti_(hasilRaw);
  tulisJadwal_(ss, saring.rows);

  let pesan = 'Generate selesai ✅\n' +
    '• Pola dari master: ' + polaValid.length + ' baris\n' +
    '• Jadwal per tanggal dihasilkan: ' + saring.rows.length + ' baris' +
    (saring.dibuang ? '\n• Dibuang karena dokter cuti (Request_Cuti): ' + saring.dibuang + ' baris' : '');
  if (errors.length) {
    pesan += '\n⚠️ ' + errors.length + ' baris master DILEWATI:\n' + errors.slice(0, 15).join('\n');
  }
  laporOtomatis_('Jadwal Dokter', pesan);
}

/* ================= HELPER ================= */

/**
 * Buang baris jadwal dokter yang bertepatan dengan request cuti dokter.
 * @return {Object} {rows: [...], dibuang: n}
 */
function saringDokterCuti_(rows) {
  const req = bacaRequestCuti_();
  if (!req || !req.dokter || !Object.keys(req.dokter).length) return { rows: rows, dibuang: 0 };
  let dibuang = 0;
  const sisa = rows.filter(function (r) {
    const tgl = String(r[0]);
    const namaDok = String(r[4]).toLowerCase();   // kolom 5 = Nama Dokter
    if (req.dokter[namaDok] && req.dokter[namaDok][tgl]) { dibuang++; return false; }
    return true;
  });
  return { rows: sisa, dibuang: dibuang };
}

/** Expand pola mingguan → baris jadwal per tanggal aktual, terurut */
function expandPola_(polaValid, pMulai, pAkhir) {
  const hasil = [];
  for (let d = new Date(pMulai); d <= pAkhir; d.setDate(d.getDate() + 1)) {
    const idx = d.getDay();
    polaValid.forEach(function (p) {
      if (p.hariIdx === idx) {
        hasil.push([
          formatTgl_(d),
          NAMA_HARI[idx].charAt(0) + NAMA_HARI[idx].slice(1).toLowerCase(),
          p.shift, p.idDokter, p.nama, p.spes, p.nurse
        ]);
      }
    });
  }
  hasil.sort(function (a, b) {
    if (a[0] !== b[0]) return a[0] < b[0] ? -1 : 1;
    if (a[2] !== b[2]) return a[2] === 'Pagi' ? -1 : 1;
    return a[4] < b[4] ? -1 : 1;
  });
  return hasil;
}

/** Ganti isi Jadwal_Dokter dengan hasil baru */
function tulisJadwal_(ss, hasil) {
  const jd = ss.getSheetByName(SHEETS.JADWAL_DOKTER);
  if (jd.getLastRow() > 1) jd.getRange(2, 1, jd.getLastRow() - 1, 7).clearContent();
  if (hasil.length) {
    jd.getRange(2, 1, hasil.length, 1).setNumberFormat('@'); // kolom Tanggal jadi TEKS agar tak dikonversi/tertukar
    jd.getRange(2, 1, hasil.length, 7).setValues(hasil);
  }
}

/** Buat/baca sheet Map_Spesialisasi (alias → kode master) */
function getAliasMap_(ss) {
  let sh = ss.getSheetByName('Map_Spesialisasi');
  if (!sh) {
    sh = ss.insertSheet('Map_Spesialisasi');
    sh.getRange(1, 1, 1, 2).setValues([['Alias (di data CIS)', 'Kode Master']]);
    sh.getRange(2, 1, ALIAS_SPESIALISASI_DEFAULT.length, 2).setValues(ALIAS_SPESIALISASI_DEFAULT);
    sh.getRange(1, 1, 1, 2).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  const map = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      if (r[0] && r[1]) map[String(r[0]).trim().toLowerCase()] = String(r[1]).trim();
    });
  }
  return map;
}

/** Index dokter by nama (lowercase) dari MD_Dokter */
function buatIndexDokter_() {
  const map = {};
  getMasterData('dokter').forEach(function (d) {
    map[d.nama.toLowerCase()] = d;
  });
  return map;
}

/** Format Date → "dd/MM/yyyy" */
function formatTgl_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
}

/** Parse "dd/MM/yyyy" → Date (atau null) */
function parseTgl_(s) {
  if (s instanceof Date) return s;
  const m = String(s || '').match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
}
