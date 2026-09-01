/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Import Massal Master Data & Utilitas
 * File: ImportMaster.gs
 *
 * 1. importMasterPerawat(): import massal perawat dari sheet
 *    Import_Perawat, format 4 kolom mulai baris 3:
 *      Nama | Level (Gold/Silver) | Kategori (Existing/New) | Kapabilitas
 *    Kapabilitas = kode spesialis dipisah koma (boleh kosong =
 *    dianggap mampu semua). Kode alias (Sp.Pros dll) otomatis
 *    diterjemahkan via Map_Spesialisasi.
 *
 * 2. rapikanMasterDokter(): perbaiki data MD_Dokter yang terlanjur
 *    ter-import dengan Jam Shift format mentah CIS ("09.00 - 15.00")
 *    → disamakan dengan format kanonik Config, plus trim spasi.
 * ============================================================
 */

const IMP_PRW_SHEET = 'Import_Perawat';

/** Menu: 👥 Import Master Perawat (massal) */
function importMasterPerawat() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Siapkan sheet bila belum ada
  let sh = ss.getSheetByName(IMP_PRW_SHEET);
  if (!sh) {
    sh = ss.insertSheet(IMP_PRW_SHEET);
    sh.getRange(1, 1).setValue(
      '<< PASTE DATA PERAWAT DI SINI, MULAI BARIS 3: Nama | Level (Gold/Silver) | Kategori (Existing/New) | Kapabilitas (kode spesialis dipisah koma, boleh kosong) >>'
    ).setFontWeight('bold').setFontColor('#b45309');
    sh.getRange(2, 1).setValue('Contoh: Ani Lestari | Gold | Existing | GP, Sp.KG, Sp.Ortho');
    ui.alert('Sheet "' + IMP_PRW_SHEET + '" baru dibuat.\nPaste data perawat di sana (mulai baris 3), lalu jalankan menu ini lagi.');
    return;
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 3) {
    ui.alert('Tidak ada data.\nPaste data perawat format 4 kolom ke sheet ' + IMP_PRW_SHEET + ' mulai baris 3.');
    return;
  }

  let raw = sh.getRange(3, 1, lastRow - 2, 4).getValues();
  raw = raw.filter(function (r) {
    return String(r[0]).trim() && String(r[0]).trim().toLowerCase() !== 'nama';
  });
  if (!raw.length) { ui.alert('Tidak ada baris data valid.'); return; }

  const aliasMap = getAliasMap_(ss);
  const masterSpes = getSpesialisasiList().map(function (s) { return s.kode; });
  const spesLower = {};
  masterSpes.forEach(function (k) { spesLower[k.toLowerCase()] = k; });

  let masuk = 0, duplikat = 0;
  const errors = [];

  raw.forEach(function (r, idx) {
    const baris = idx + 3;
    const nama = String(r[0]).trim();

    // Normalisasi level
    let level = String(r[1]).trim().toLowerCase();
    if (level === 'gold') level = 'Gold';
    else if (level === 'silver') level = 'Silver';
    else { errors.push('Baris ' + baris + ': level "' + r[1] + '" harus Gold/Silver.'); return; }

    // Normalisasi kategori
    let kategori = String(r[2]).trim().toLowerCase();
    if (kategori === 'existing' || kategori === 'exist') kategori = 'Existing';
    else if (kategori === 'new' || kategori === 'fresh') kategori = 'New';
    else { errors.push('Baris ' + baris + ': kategori "' + r[2] + '" harus Existing/New.'); return; }

    // Normalisasi kapabilitas: pisah koma/;, terjemahkan alias, validasi kode
    const kapabRaw = String(r[3] || '').split(/[,;]/)
      .map(function (s) { return s.trim(); })
      .filter(function (s) { return s; });
    const kapab = [];
    let gagalKode = null;
    kapabRaw.forEach(function (k) {
      let kode = spesLower[k.toLowerCase()] || aliasMap[k.toLowerCase()];
      if (kode && masterSpes.indexOf(kode) !== -1) {
        if (kapab.indexOf(kode) === -1) kapab.push(kode);
      } else {
        gagalKode = k;
      }
    });
    if (gagalKode) {
      errors.push('Baris ' + baris + ': kode kapabilitas "' + gagalKode + '" tidak dikenal.');
      return;
    }

    const hasil = addMasterData('perawat', {
      nama: nama, level: level, kategori: kategori, kapabilitas: kapab.join(', ')
    });
    if (hasil.ok) masuk++;
    else if (hasil.pesan.indexOf('sudah terdaftar') !== -1) duplikat++;
    else errors.push('Baris ' + baris + ': ' + hasil.pesan);
  });

  let pesan = 'Import perawat selesai ✅\n\n' +
    '• Berhasil masuk: ' + masuk + '\n' +
    '• Duplikat dilewati: ' + duplikat + '\n';
  if (errors.length) {
    pesan += '\n⚠️ ' + errors.length + ' baris DILEWATI:\n' + errors.slice(0, 15).join('\n') +
      (errors.length > 15 ? '\n… dan ' + (errors.length - 15) + ' lainnya.' : '');
  }
  ui.alert('Hasil Import Perawat', pesan, ui.ButtonSet.OK);
}

/** Menu: 🧹 Rapikan Master Dokter (jam shift & spasi) */
function rapikanMasterDokter() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.MD_DOKTER);
  if (!sh || sh.getLastRow() < 2) { ui.alert('MD_Dokter kosong.'); return; }

  const jamPagi = String(getConfig('JAM_SHIFT_PAGI') || '09:00-15:00');
  const jamSiang = String(getConfig('JAM_SHIFT_SIANG') || '15:00-21:00');

  // Kolom: 1 ID, 2 Nama, 3 Spes, 4 Kebutuhan, 5 Hari, 6 JamAwal, 7 JamAkhir, 8 JamShift, 9 KategoriShift, 10 Aktif
  const n = sh.getLastRow() - 1;
  const data = sh.getRange(2, 1, n, 10).getValues();
  let ubah = 0;

  const baru = data.map(function (r) {
    const asli = JSON.stringify(r);
    // Trim semua kolom teks
    for (let c = 0; c < 10; c++) {
      if (typeof r[c] === 'string') r[c] = r[c].trim();
    }
    // Normalisasi kategori shift ("Pagi", "shift pagi ", dll → "Shift Pagi")
    const kat = String(r[8]).toLowerCase().replace(/^shift\s*/, '').trim();
    if (kat === 'pagi') r[8] = 'Shift Pagi';
    else if (kat === 'siang') r[8] = 'Shift Siang';
    // Jam shift kanonik dari Config berdasarkan kategori shift
    if (r[8] === 'Shift Pagi') r[7] = jamPagi;
    else if (r[8] === 'Shift Siang') r[7] = jamSiang;
    if (JSON.stringify(r) !== asli) ubah++;
    return r;
  });

  sh.getRange(2, 1, n, 10).setValues(baru);
  ui.alert('Selesai 🧹', ubah + ' dari ' + n + ' baris MD_Dokter dirapikan.\n' +
    'Jam Shift kini seragam: "' + jamPagi + '" / "' + jamSiang + '".', ui.ButtonSet.OK);
}
