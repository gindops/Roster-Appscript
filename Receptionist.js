/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Fase 4: Papan Receptionist
 * File: Receptionist.gs
 *
 * Menu: 🛎️ Generate Papan Receptionist
 * Membangun papan tanggal × shift lengkap dengan:
 *  - DU Aktif per shift (R-05: dihitung dari jumlah dokter praktik)
 *  - Minimum receptionist (R-03: DU 1-4 → 2; R-04: DU >4 → 3)
 *  - Auto-assign receptionist:
 *      · menghormati Papan_Libur (R-06 divalidasi audit)
 *      · Silver diprioritaskan longshift (pagi+siang) sampai
 *        genap kewajiban R-07
 *      · beban diratakan; R-02 terpenuhi karena jumlah mengikuti DU
 *
 * Edit manual tetap divalidasi real-time oleh Validator.gs.
 * ============================================================
 */

/** Kolom Papan_Receptionist */
const PR_HEADER = ['Tanggal', 'Hari', 'Shift', 'DU Aktif', 'Min Receptionist', 'Recept 1', 'Recept 2', 'Recept 3'];
const PR_KOL_SLOT1 = 6;   // kolom Recept 1
const PR_MAX_SLOT = 3;    // R-04: kebutuhan maksimum 3

/** Menu: 🛎️ Generate Papan Receptionist */
function generatePapanReceptionist() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const jd = ss.getSheetByName(SHEETS.JADWAL_DOKTER);
  if (!jd || jd.getLastRow() < 2) {
    laporOtomatis_('Papan Receptionist', 'Jadwal_Dokter kosong.\nGenerate/import dulu jadwal dokter.');
    return;
  }
  const receptionist = getMasterData('receptionist')
    .filter(function (r) { return r.aktif !== 'Tidak'; })
    .map(function (r) { return { nama: r.nama, level: r.level }; });
  if (!receptionist.length) { laporOtomatis_('Papan Receptionist', 'MD_Receptionist kosong.'); return; }

  if (!konfirmasiOtomatis_('Generate Papan Receptionist',
    'Papan_Receptionist akan dibuat ulang dan diisi otomatis.\n' +
    'Isian lama akan hilang. Lanjutkan?')) return;

  // --- DU aktif per tanggal per shift (R-05) ---
  const du = {}; // tanggal → {Pagi: n, Siang: n, hari}
  jd.getRange(2, 1, jd.getLastRow() - 1, 7).getValues().forEach(function (r) {
    if (!r[0]) return;
    const t = String(r[0]);
    if (!du[t]) du[t] = { Pagi: 0, Siang: 0, hari: String(r[1]) };
    du[t][String(r[2])]++;
  });

  const ambang = Number(getConfig('AMBANG_DU')) || 4;
  const minKecil = Number(getConfig('MIN_RECEPTIONIST_DU_KECIL')) || 2;
  const minBesar = Number(getConfig('MIN_RECEPTIONIST_DU_BESAR')) || 3;
  const minDari = function (n) {
    if (n <= 0) return 0;
    return n <= ambang ? minKecil : minBesar;
  };

  const tanggalUrut = Object.keys(du).sort(function (a, b) {
    return kunciTgl_(a) < kunciTgl_(b) ? -1 : 1;
  });

  const hasil = susunReceptionist_({
    tanggalUrut: tanggalUrut,
    du: du,
    minDari: minDari,
    receptionist: receptionist,
    liburMap: petaLiburPerNama_(ss),
    longshiftTarget: Number(getConfig('LONGSHIFT_WAJIB_RECEPT_SILVER')) || 4,
    jagaMap: bacaRequestJaga_().map
  });

  // --- Tulis papan ---
  let pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (!pr) pr = ss.insertSheet(SHEETS.PAPAN_RECEPTIONIST);
  pr.clear();
  pr.getRange(1, 1, pr.getMaxRows(), pr.getMaxColumns()).clearDataValidations(); // buang dropdown lama
  pr.getRange(1, 1, 1, PR_HEADER.length).setValues([PR_HEADER]);
  formatHeader_(pr, PR_HEADER.length);

  if (hasil.rows.length) {
    // Dropdown nama receptionist dipasang DULU, baru tulis nama
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(receptionist.map(function (r) { return r.nama; }), true)
      .setAllowInvalid(true).build();
    pr.getRange(2, PR_KOL_SLOT1, hasil.rows.length, PR_MAX_SLOT).setDataValidation(rule);
    pr.getRange(2, 1, hasil.rows.length, 1).setNumberFormat('@'); // kolom Tanggal jadi TEKS
    pr.getRange(2, 1, hasil.rows.length, PR_HEADER.length).setValues(hasil.rows);
    // Tandai baris yang kekurangan
    hasil.baruKurang.forEach(function (idx) {
      pr.getRange(idx + 2, PR_KOL_SLOT1, 1, PR_MAX_SLOT).setBackground(WARNA_PELANGGARAN);
    });
  }
  warnaiPerTanggal_(pr, hasil.rows.length, 1);
  pr.setFrozenRows(1);

  // --- Ringkasan ---
  let pesan = 'Papan Receptionist selesai ✅\n\n' +
    '• Baris shift dibuat: ' + hasil.rows.length + '\n' +
    '• Shift kekurangan receptionist: ' + hasil.baruKurang.length + ' (ditandai merah)\n' +
    '• Silver longshift genap ' + hasil.longshiftTarget + 'x: ' +
    hasil.silverGenap + '/' + hasil.totalSilver + '\n';
  if (hasil.silverKurang.length) {
    pesan += '\n⚠️ Silver longshift belum genap:\n' + hasil.silverKurang
      .map(function (k) { return '• ' + k.nama + ': ' + k.longshift + 'x'; }).join('\n') + '\n';
  }
  pesan += '\nIni DRAF — edit manual tetap divalidasi real-time.';
  laporOtomatis_('Papan Receptionist', pesan);
}

/**
 * Algoritma inti (murni — bisa diuji terpisah).
 * Strategi per tanggal: butuh max(minPagi, minSiang) orang; yang bertugas
 * di kedua shift = longshift. Prioritas longshift: Silver yang belum genap
 * kewajibannya (R-07), lalu beban paling ringan.
 * @return {Object} {rows, baruKurang[], silver*, totalSilver}
 */
function susunReceptionist_(input) {
  const stat = {}; // nama → {total, longshift}
  input.receptionist.forEach(function (r) { stat[r.nama] = { total: 0, longshift: 0 }; });

  const rows = [];
  const baruKurang = [];
  const shuffle = function (arr) { const a = arr.slice(); for (let i=a.length-1;i>0;i--){ const j=Math.floor(Math.random()*(i+1)); const tmp=a[i];a[i]=a[j];a[j]=tmp; } return a; };

  // Berapa banyak longshift Silver yang boleh SENGAJA dibuat per hari (tersebar merata)
  const CAP_SILVER_LONG = 1;
  const jagaMap = input.jagaMap || {}; // namaLower → {tanggal→true} (request jaga)
  const mintaJaga = function (nm, t) { const k = String(nm).toLowerCase(); return !!(jagaMap[k] && jagaMap[k][t]); };

  // --- Satu pass: isi tiap shift; longshift MINIMAL, hanya untuk memenuhi kuota Silver (R-07) ---
  const dayAssign = {};
  input.tanggalUrut.forEach(function (t) {
    const d = input.du[t];
    const minP = input.minDari(d.Pagi), minS = input.minDari(d.Siang);
    const tersedia = input.receptionist.filter(function (r) { return !(input.liburMap[r.nama] && input.liburMap[r.nama][t]); });

    // Silver yang masih kurang kuota longshift → prioritas dijadwalkan agar bisa longshift
    const needSilver = shuffle(tersedia.filter(function (r) { return r.level === 'Silver' && stat[r.nama].longshift < input.longshiftTarget; }))
      .sort(function (a, b) { return stat[a.nama].longshift - stat[b.nama].longshift; });
    const others = shuffle(tersedia).sort(function (a, b) {
      const ja = mintaJaga(a.nama, t) ? 0 : 1, jb = mintaJaga(b.nama, t) ? 0 : 1;
      if (ja !== jb) return ja - jb;                            // Request Jaga didahulukan
      return stat[a.nama].total - stat[b.nama].total;
    });

    // PAGI: dahulukan hingga CAP Silver yang butuh kuota, lalu isi dari beban paling ringan
    const pagi = [];
    const sudah = {};
    needSilver.forEach(function (r) { if (pagi.length >= Math.min(minP, CAP_SILVER_LONG)) return; if (!sudah[r.nama.toLowerCase()]) { pagi.push(r.nama); sudah[r.nama.toLowerCase()] = true; } });
    others.forEach(function (r) { if (pagi.length >= minP) return; if (!sudah[r.nama.toLowerCase()]) { pagi.push(r.nama); sudah[r.nama.toLowerCase()] = true; } });
    const pagiSet = {}; pagi.forEach(function (n) { pagiSet[n.toLowerCase()] = true; });

    // SIANG: sengaja longshift-kan Silver-butuh-kuota yang tadi masuk pagi (maks CAP), sisanya orang BEDA
    const siang = [];
    needSilver.forEach(function (r) { if (siang.length >= Math.min(minS, CAP_SILVER_LONG)) return; if (pagiSet[r.nama.toLowerCase()]) siang.push(r.nama); });
    const siangSet0 = {}; siang.forEach(function (n) { siangSet0[n.toLowerCase()] = true; });
    others.forEach(function (r) { if (siang.length >= minS) return; const low = r.nama.toLowerCase(); if (!pagiSet[low] && !siangSet0[low]) { siang.push(r.nama); siangSet0[low] = true; } });
    // Bila orang beda tak cukup, baru pinjam dari pagi (longshift TERPAKSA karena staf kurang)
    if (siang.length < minS) { for (let k = 0; k < pagi.length && siang.length < minS; k++) { if (siang.indexOf(pagi[k]) === -1) siang.push(pagi[k]); } }
    const siangSet = {}; siang.forEach(function (n) { siangSet[n.toLowerCase()] = true; });

    tersedia.forEach(function (r) { const p = pagiSet[r.nama.toLowerCase()], s2 = siangSet[r.nama.toLowerCase()];
      if (p) stat[r.nama].total++; if (s2) stat[r.nama].total++; if (p && s2) stat[r.nama].longshift++; });
    dayAssign[t] = { du: d, minP: minP, minS: minS, pagi: pagi, siang: siang };
  });

  // --- Bangun baris papan ---
  input.tanggalUrut.forEach(function (t) {
    const da = dayAssign[t], d = da.du;
    const buatBaris = function (shift, duN, minN, list) {
      const b = [t, d.hari, shift, duN, minN];
      for (let i = 0; i < PR_MAX_SLOT; i++) b.push(list[i] || '');
      rows.push(b);
      if (list.length < minN) baruKurang.push(rows.length - 1);
    };
    buatBaris('Pagi', d.Pagi, da.minP, da.pagi);
    buatBaris('Siang', d.Siang, da.minS, da.siang);
  });

  // Statistik Silver R-07
  let silverGenap = 0, totalSilver = 0;
  const silverKurang = [];
  input.receptionist.forEach(function (r) {
    if (r.level !== 'Silver') return;
    totalSilver++;
    if (stat[r.nama].longshift >= input.longshiftTarget) silverGenap++;
    else silverKurang.push({ nama: r.nama, longshift: stat[r.nama].longshift });
  });

  return {
    rows: rows, baruKurang: baruKurang,
    longshiftTarget: input.longshiftTarget,
    silverGenap: silverGenap, totalSilver: totalSilver, silverKurang: silverKurang
  };
}