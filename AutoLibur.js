/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Auto-Generate Papan Libur
 * File: AutoLibur.gs
 *
 * Menu: 🎲 Auto-Generate Papan Libur
 * Menyusun DRAF pembagian libur acak yang patuh aturan:
 *  - N-02/R-06 : tepat 4 libur per perawat & receptionist
 *  - N-03      : Silver tidak weekend; Gold maks 1x weekend
 *  - Kapasitas : jumlah perawat tersisa per tanggal >= slot asistensi
 *                terbanyak (pagi/siang) dari Jadwal_Dokter
 *  - Kapabilitas: perawat tidak diliburkan bila sisa perawat yang
 *                mampu ke spesialis yang praktik hari itu tidak cukup
 *  - Receptionist tersisa per tanggal >= kebutuhan R-03/R-04
 *  - Sebaran   : libur diusahakan jatuh di minggu yang berbeda
 *
 * Hasil adalah DRAF — SPV tetap bisa mengubah manual, dan
 * 🔍 Cek Kepatuhan Periode tetap menjadi penentu akhir.
 * ============================================================
 */

/** Menu: 🎲 Auto-Generate Papan Libur */
function autoGeneratePapanLibur() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const pMulai = parseTgl_(getConfig('PERIODE_MULAI'));
  const pAkhir = parseTgl_(getConfig('PERIODE_AKHIR'));
  if (!pMulai || !pAkhir) { ui.alert('Periode aktif belum diset (📅 Set Periode Aktif).'); return; }

  const jd = ss.getSheetByName(SHEETS.JADWAL_DOKTER);
  if (!jd || jd.getLastRow() < 2) {
    laporOtomatis_('Papan Libur', 'Jadwal_Dokter kosong.\nGenerate/import dulu agar kapasitas harian bisa dihitung.');
    return;
  }

  if (!konfirmasiOtomatis_('Auto-Generate Papan Libur',
    'Papan_Libur akan DIBUAT ULANG dan diisi draf pembagian libur otomatis.\n' +
    'Isian libur yang ada sekarang akan hilang. Lanjutkan?')) return;

  // --- Susun input untuk algoritma ---
  const weekend = String(getConfig('HARI_WEEKEND') || 'Sabtu,Minggu').split(',')
    .map(function (h) { return h.trim().toLowerCase(); });

  // Daftar tanggal periode
  const tanggalList = [];
  for (let d = new Date(pMulai); d <= pAkhir; d.setDate(d.getDate() + 1)) {
    const hari = NAMA_HARI[d.getDay()].charAt(0) + NAMA_HARI[d.getDay()].slice(1).toLowerCase();
    tanggalList.push({
      tanggal: formatTgl_(d),
      hari: hari,
      isWeekend: weekend.indexOf(hari.toLowerCase()) !== -1,
      minggu: Math.floor((d - pMulai) / (7 * 24 * 3600 * 1000))
    });
  }

  // Demand per tanggal dari Jadwal_Dokter
  const demand = {}; // tanggal → {pagi, siang, spes: {kode → max slot per shift}, duPagi, duSiang}
  jd.getRange(2, 1, jd.getLastRow() - 1, 7).getValues().forEach(function (r) {
    if (!r[0]) return;
    const t = String(r[0]), shift = String(r[2]), spes = String(r[5]);
    const n = parseInt(r[6], 10) || 1;
    if (!demand[t]) demand[t] = { pagi: 0, siang: 0, spes: {}, spesPagi: {}, spesSiang: {}, duPagi: 0, duSiang: 0 };
    const dd = demand[t];
    if (shift === 'Pagi') {
      dd.pagi += n; dd.duPagi++;
      dd.spesPagi[spes] = (dd.spesPagi[spes] || 0) + n;
    } else {
      dd.siang += n; dd.duSiang++;
      dd.spesSiang[spes] = (dd.spesSiang[spes] || 0) + n;
    }
    dd.spes[spes] = Math.max(dd.spesPagi[spes] || 0, dd.spesSiang[spes] || 0);
  });

  const ambangDU = Number(getConfig('AMBANG_DU')) || 4;
  const minKecil = Number(getConfig('MIN_RECEPTIONIST_DU_KECIL')) || 2;
  const minBesar = Number(getConfig('MIN_RECEPTIONIST_DU_BESAR')) || 3;

  const perawat = getMasterData('perawat')
    .filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) {
      return {
        nama: p.nama, level: p.level,
        kapab: String(p.kapabilitas || '').split(',')
          .map(function (s) { return s.trim(); }).filter(function (s) { return s; })
      };
    });
  const receptionist = getMasterData('receptionist')
    .filter(function (r) { return r.aktif !== 'Tidak'; })
    .map(function (r) { return { nama: r.nama, level: r.level }; });

  // Libur WAJIB dari Request_Cuti (perawat & receptionist)
  const req = bacaRequestCuti_();
  const perawatNama = {}; perawat.forEach(function (p) { perawatNama[p.nama] = true; });
  const rcpNama = {}; receptionist.forEach(function (r) { rcpNama[r.nama] = true; });
  const wajibPerawat = {}, wajibRcp = {};
  Object.keys(req.staf).forEach(function (nm) {
    if (perawatNama[nm]) wajibPerawat[nm] = req.staf[nm];
    else if (rcpNama[nm]) wajibRcp[nm] = req.staf[nm];
  });

  const input = {
    tanggalList: tanggalList,
    demand: demand,
    perawat: perawat,
    receptionist: receptionist,
    jatahPerawat: Number(getConfig('LIBUR_WAJIB_PERAWAT')) || 4,
    jatahReceptionist: Number(getConfig('LIBUR_WAJIB_RECEPTIONIST')) || 4,
    maxWeekendGold: Number(getConfig('MAX_LIBUR_WEEKEND_GOLD')) || 1,
    maxSlot: PL_MAX_SLOT,
    ambangDU: ambangDU, minKecil: minKecil, minBesar: minBesar,
    wajibPerawat: wajibPerawat, wajibRcp: wajibRcp
  };

  // --- Jalankan algoritma (fungsi murni, teruji terpisah) ---
  const hasil = susunLibur_(input);

  // --- Bangun ulang Papan_Libur & tulis hasil ---
  buatPapanLiburTanpaKonfirmasi_(ss, tanggalList);
  const pl = ss.getSheetByName('Papan_Libur');
  tanggalList.forEach(function (t, i) {
    const names = hasil.liburMap[t.tanggal] || [];
    if (names.length) {
      pl.getRange(i + 2, 3, 1, names.length).setValues([names.slice(0, PL_MAX_SLOT)]);
    }
  });

  let pesan = 'Draf libur selesai ✅\n\n' +
    '• Perawat terjadwal penuh (' + input.jatahPerawat + 'x): ' + hasil.sukses.perawat + '/' + perawat.length + '\n' +
    '• Receptionist terjadwal penuh: ' + hasil.sukses.receptionist + '/' + receptionist.length + '\n';
  if (hasil.kurang.length) {
    pesan += '\n⚠️ Belum dapat jatah penuh (atur manual / kendurkan kendala):\n' +
      hasil.kurang.map(function (k) { return '• ' + k.nama + ': ' + k.dapat + '/' + k.target; }).join('\n');
  }
  pesan += '\nIni DRAF — silakan sesuaikan manual bila perlu.';
  laporOtomatis_('Papan Libur', pesan);
}

/**
 * Algoritma inti (murni, tanpa akses Spreadsheet — bisa diuji terpisah).
 * Greedy + skor: untuk tiap staf pilih tanggal terbaik satu per satu.
 * @return {Object} {liburMap: {tanggal: [nama]}, sukses:{}, kurang:[]}
 */
function susunLibur_(input) {
  const liburMap = {};                 // tanggal → [nama]
  const liburPerawatCount = {};        // tanggal → jumlah perawat libur
  const liburRcpCount = {};            // tanggal → jumlah receptionist libur
  input.tanggalList.forEach(function (t) {
    liburMap[t.tanggal] = [];
    liburPerawatCount[t.tanggal] = 0;
    liburRcpCount[t.tanggal] = 0;
  });

  const totalPerawat = input.perawat.length;
  const totalRcp = input.receptionist.length;

  // Kebutuhan minimum per tanggal
  const minPerawat = {}; // max(slot pagi, slot siang)
  const minRcp = {};     // dari DU aktif (R-03/R-04), 0 bila tak ada dokter
  input.tanggalList.forEach(function (t) {
    const d = input.demand[t.tanggal];
    minPerawat[t.tanggal] = d ? Math.max(d.pagi, d.siang) : 0;
    if (!d) { minRcp[t.tanggal] = 0; return; }
    const kebutuhan = function (du) {
      if (du <= 0) return 0;
      return du <= input.ambangDU ? input.minKecil : input.minBesar;
    };
    minRcp[t.tanggal] = Math.max(kebutuhan(d.duPagi), kebutuhan(d.duSiang));
  });

  // Jumlah perawat mampu per spesialis (kapab kosong = mampu semua)
  const mampu = {}; // kode spes → jumlah perawat mampu
  input.perawat.forEach(function (p) {
    const semua = p.kapab.length === 0;
    Object.keys(kumpulSpes_(input)).forEach(function (s) {
      if (semua || p.kapab.indexOf(s) !== -1) mampu[s] = (mampu[s] || 0) + 1;
    });
  });
  // Perawat mampu yang sedang libur, per tanggal per spes
  const liburMampu = {}; // tanggal → {spes → count}
  input.tanggalList.forEach(function (t) { liburMampu[t.tanggal] = {}; });

  const shuffled = function (arr) {
    const a = arr.slice();
    for (let i = a.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      const tmp = a[i]; a[i] = a[j]; a[j] = tmp;
    }
    return a;
  };

  /** Boleh tidak perawat p libur di tanggal t? */
  const bolehPerawat = function (p, t) {
    if (liburMap[t.tanggal].indexOf(p.nama) !== -1) return false;
    if (liburMap[t.tanggal].length >= input.maxSlot) return false;
    if (t.isWeekend && p.level !== 'Gold') return false;
    // Kapasitas total perawat
    if (totalPerawat - liburPerawatCount[t.tanggal] - 1 < minPerawat[t.tanggal]) return false;
    // Kapabilitas: sisa perawat mampu per spesialis yang praktik hari itu
    const d = input.demand[t.tanggal];
    if (d) {
      const spesList = Object.keys(d.spes);
      for (let i = 0; i < spesList.length; i++) {
        const s = spesList[i];
        const pMampu = p.kapab.length === 0 || p.kapab.indexOf(s) !== -1;
        if (!pMampu) continue;
        const sisaMampu = (mampu[s] || 0) - (liburMampu[t.tanggal][s] || 0) - 1;
        if (sisaMampu < d.spes[s]) return false;
      }
    }
    return true;
  };

  const catatPerawat = function (p, t) {
    liburMap[t.tanggal].push(p.nama);
    liburPerawatCount[t.tanggal]++;
    const d = input.demand[t.tanggal];
    if (d) {
      Object.keys(d.spes).forEach(function (s) {
        if (p.kapab.length === 0 || p.kapab.indexOf(s) !== -1) {
          liburMampu[t.tanggal][s] = (liburMampu[t.tanggal][s] || 0) + 1;
        }
      });
    }
  };

  /** Skor tanggal: makin kecil makin bagus (sepi libur, minggu belum terpakai, hari sepi praktik) */
  const skor = function (t, mingguDipakai) {
    let s = liburMap[t.tanggal].length * 10;
    if (mingguDipakai[t.minggu]) s += 15;
    s += minPerawat[t.tanggal];            // hari sepi praktik lebih diprioritaskan
    s += Math.random() * 4;                // unsur acak agar tiap generate berbeda
    return s;
  };

  const kurang = [];
  let suksesPerawat = 0;

  /** Pra-tempat libur WAJIB (request cuti) — abaikan aturan weekend & kapasitas
   *  karena ini permintaan eksplisit; understaffing akan ditandai audit. */
  const praTempat = function (nama, tglList, isPerawat, pObj) {
    let n = 0;
    (tglList || []).forEach(function (tgl) {
      const t = null;
      const tt = input.tanggalList.filter(function (x) { return x.tanggal === tgl; })[0];
      if (!tt) return;
      if (liburMap[tt.tanggal].indexOf(nama) !== -1) return;
      if (isPerawat) { catatPerawat(pObj, tt); }
      else { liburMap[tt.tanggal].push(nama); liburRcpCount[tt.tanggal]++; }
      n++;
    });
    return n;
  };

  // --- Bagikan jatah perawat ---
  shuffled(input.perawat).forEach(function (p) {
    const mingguDipakai = {};
    let dapat = 0;

    // Libur wajib dari request cuti dulu
    dapat += praTempat(p.nama, input.wajibPerawat && input.wajibPerawat[p.nama], true, p);

    // Gold: coba 1x weekend dulu (maks sesuai config)
    if (p.level === 'Gold' && input.maxWeekendGold > 0) {
      const kandidatW = shuffled(input.tanggalList.filter(function (t) {
        return t.isWeekend && bolehPerawat(p, t);
      }));
      if (kandidatW.length) {
        catatPerawat(p, kandidatW[0]);
        mingguDipakai[kandidatW[0].minggu] = true;
        dapat++;
      }
    }

    // Sisa jatah: hari kerja, pilih skor terbaik
    while (dapat < input.jatahPerawat) {
      const kandidat = input.tanggalList.filter(function (t) {
        return !t.isWeekend && bolehPerawat(p, t);
      });
      if (!kandidat.length) break;
      kandidat.sort(function (a, b) { return skor(a, mingguDipakai) - skor(b, mingguDipakai); });
      catatPerawat(p, kandidat[0]);
      mingguDipakai[kandidat[0].minggu] = true;
      dapat++;
    }

    if (dapat >= input.jatahPerawat) suksesPerawat++;
    else kurang.push({ nama: p.nama, dapat: dapat, target: input.jatahPerawat });
  });

  // --- Bagikan jatah receptionist ---
  let suksesRcp = 0;
  shuffled(input.receptionist).forEach(function (r) {
    const mingguDipakai = {};
    let dapat = 0;
    // Libur wajib dari request cuti dulu
    dapat += praTempat(r.nama, input.wajibRcp && input.wajibRcp[r.nama], false, null);
    while (dapat < input.jatahReceptionist) {
      const kandidat = input.tanggalList.filter(function (t) {
        if (liburMap[t.tanggal].indexOf(r.nama) !== -1) return false;
        if (liburMap[t.tanggal].length >= input.maxSlot) return false;
        return totalRcp - liburRcpCount[t.tanggal] - 1 >= minRcp[t.tanggal];
      });
      if (!kandidat.length) break;
      kandidat.sort(function (a, b) { return skor(a, mingguDipakai) - skor(b, mingguDipakai); });
      liburMap[kandidat[0].tanggal].push(r.nama);
      liburRcpCount[kandidat[0].tanggal]++;
      mingguDipakai[kandidat[0].minggu] = true;
      dapat++;
    }
    if (dapat >= input.jatahReceptionist) suksesRcp++;
    else kurang.push({ nama: r.nama, dapat: dapat, target: input.jatahReceptionist });
  });

  return {
    liburMap: liburMap,
    sukses: { perawat: suksesPerawat, receptionist: suksesRcp },
    kurang: kurang
  };
}

/** Kumpulkan semua kode spesialis yang muncul di demand */
function kumpulSpes_(input) {
  const set = {};
  Object.keys(input.demand).forEach(function (t) {
    Object.keys(input.demand[t].spes).forEach(function (s) { set[s] = true; });
  });
  return set;
}

/** Bangun ulang struktur Papan_Libur tanpa dialog (dipakai auto-generate) */
function buatPapanLiburTanpaKonfirmasi_(ss, tanggalList) {
  let pl = ss.getSheetByName('Papan_Libur');
  if (!pl) pl = ss.insertSheet('Papan_Libur');
  pl.clear();

  const header = ['Tanggal', 'Hari'];
  for (let i = 1; i <= PL_MAX_SLOT; i++) header.push('Libur ' + i);
  pl.getRange(1, 1, 1, header.length).setValues([header]);
  formatHeader_(pl, header.length);

  const rows = tanggalList.map(function (t) {
    const baris = [t.tanggal, t.hari];
    for (let i = 0; i < PL_MAX_SLOT; i++) baris.push('');
    return baris;
  });
  pl.getRange(2, 1, rows.length, 1).setNumberFormat('@'); // kolom Tanggal jadi TEKS
  pl.getRange(2, 1, rows.length, header.length).setValues(rows);

  const namaStaf = getMasterData('perawat')
    .filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) { return p.nama; })
    .concat(getMasterData('receptionist')
      .filter(function (r) { return r.aktif !== 'Tidak'; })
      .map(function (r) { return r.nama; }));
  if (namaStaf.length) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(namaStaf, true).setAllowInvalid(true).build(); // warn, jangan blokir keras (staf baru tetap bisa ditulis)
    pl.getRange(2, 3, rows.length, PL_MAX_SLOT).setDataValidation(rule);
  }

  tanggalList.forEach(function (t, i) {
    if (t.isWeekend) pl.getRange(i + 2, 1, 1, header.length).setBackground('#fff3e0');
  });
  pl.setFrozenRows(1);
  pl.setFrozenColumns(2);
}