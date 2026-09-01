/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Mesin Validasi (v2)
 * File: Validator.gs
 *
 * Mode kombinasi:
 *  1. REAL-TIME (onEdit): cek ringan per sel saat SPV mengedit —
 *     Papan_Asistensi : kapabilitas (N-05), dobel assign, A-02,
 *                       bentrok dengan jatah libur
 *     Papan_Libur     : duplikat, weekend (N-03)
 *     Papan_Receptionist: nama tak dikenal, bentrok libur,
 *                       duplikat 1 baris, kekurangan vs minimum
 *  2. AUDIT PERIODE (🔍 Cek Kepatuhan Periode): scan penuh —
 *     N-02, N-03, A-02, A-03 (+gugur otomatis & gap kapabilitas),
 *     A-04, R-03/R-04, R-06, R-07, bentrok libur, slot kosong.
 *
 * Warna sel: merah = pelanggaran hard, oranye = peringatan.
 * ============================================================
 */

const WARNA_VALID = '#e6f4ea';
const WARNA_WARNING = '#fff3cd';
const WARNA_PELANGGARAN = '#f8d7da';

/* ================= 1. VALIDASI REAL-TIME ================= */

/** Simple trigger: berjalan setiap kali user mengedit sel */
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sheet = e.range.getSheet();
    const nama = sheet.getName();
    if (nama === SHEETS.PAPAN_ASISTENSI && e.range.getColumn() === PA_KOL_PERAWAT) {
      validasiSelAsistensi_(sheet, e.range.getRow());
    } else if (nama === 'Papan_Libur' && e.range.getColumn() >= 3) {
      validasiSelLibur_(sheet, e.range.getRow(), e.range.getColumn());
    } else if (nama === SHEETS.PAPAN_RECEPTIONIST && e.range.getColumn() >= PR_KOL_SLOT1) {
      validasiSelReceptionist_(sheet, e.range.getRow(), e.range.getColumn());
    } else if (nama === SHEETS.MD_PERAWAT && e.range.getColumn() === 5 && e.range.getRow() >= 2) {
      multiSelectKapabilitas_(e);
    } else if ((nama === SHEETS.MD_PERAWAT || nama === SHEETS.MD_RECEPTIONIST) &&
               e.range.getColumn() === 2 && e.range.getRow() >= 2) {
      autoIdManual_(sheet, e.range.getRow(), nama);
    }
    tandaiDrafBilaFinal_(nama);
  } catch (err) {
    console.error(err);
  }
}

/**
 * Bila status jadwal sudah FINAL lalu ada edit MANUAL di papan kerja,
 * status otomatis dikembalikan ke "DRAF (diubah setelah final)" agar
 * tidak menyesatkan. Jalankan ✅ Cek Final lagi untuk memfinalkan ulang.
 * (Perubahan programatik/auto-generate tidak memicu simple trigger ini.)
 */
function tandaiDrafBilaFinal_(namaSheet) {
  const papan = [SHEETS.PAPAN_ASISTENSI, SHEETS.PAPAN_RECEPTIONIST, 'Papan_Libur'];
  if (papan.indexOf(namaSheet) === -1) return;
  const st = String(getConfig('STATUS_JADWAL') || '');
  if (st.indexOf('FINAL ✓') === 0) {
    setConfig_('STATUS_JADWAL', 'DRAF (diubah setelah final)', 'Status quality gate');
  }
}

/**
 * Auto-ID saat input MANUAL langsung di sheet (tanpa sidebar).
 * Ketik nama di kolom B pada baris ber-ID kosong → ID otomatis dibuat
 * (PRW-xxx / RCP-xxx) dan kolom Aktif diisi "Ya".
 * Jalur ini kebal terhadap error PERMISSION_DENIED sidebar karena
 * berjalan sebagai simple trigger onEdit.
 */
function autoIdManual_(sheet, row, namaSheet) {
  const selId = sheet.getRange(row, 1);
  if (String(selId.getValue()).trim()) return; // ID sudah ada → biarkan
  const nama = String(sheet.getRange(row, 2).getValue()).trim();
  if (!nama) return;

  const prefix = (namaSheet === SHEETS.MD_PERAWAT) ? 'PRW' : 'RCP';
  const last = sheet.getLastRow();
  let maxNum = 0;
  if (last >= 2) {
    sheet.getRange(2, 1, last - 1, 1).getValues().forEach(function (r) {
      const m = String(r[0]).match(/(\d+)$/);
      if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
    });
  }
  selId.setValue(prefix + '-' + ('000' + (maxNum + 1)).slice(-3));

  // Kolom Aktif: MD_Perawat = kolom 6, MD_Receptionist = kolom 4
  const kolAktif = (namaSheet === SHEETS.MD_PERAWAT) ? 6 : 4;
  const selAktif = sheet.getRange(row, kolAktif);
  if (!String(selAktif.getValue()).trim()) selAktif.setValue('Ya');
}

/**
 * Multi-select toggle untuk kolom Kapabilitas (MD_Perawat kolom E).
 * Saat SPV memilih 1 spesialis dari dropdown:
 *  - belum ada di sel  → ditambahkan ke daftar (dipisah ", ")
 *  - sudah ada         → dihapus (toggle off)
 * Sel dikosongkan → dibiarkan kosong (= mampu semua spesialis).
 */
function multiSelectKapabilitas_(e) {
  const sel = e.range;
  const dipilih = String(e.value || '').trim();     // nilai baru dari dropdown
  const lama = String(e.oldValue || '').trim();     // isi sebelumnya

  // Sel dikosongkan manual → biarkan kosong
  if (!dipilih) { sel.setNote(''); return; }

  // Bila nilai baru sudah berupa gabungan (mis. hasil paste), biarkan apa adanya
  if (dipilih.indexOf(',') !== -1) { sel.setNote(catatanKapab_(dipilih)); return; }

  const arr = lama ? lama.split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; }) : [];
  const idx = arr.indexOf(dipilih);
  if (idx === -1) arr.push(dipilih);   // tambah
  else arr.splice(idx, 1);             // toggle off

  sel.setValue(arr.join(', ')).setNote(catatanKapab_(arr.join(', ')));
}

/** Catatan sel kapabilitas */
function catatanKapab_(teks) {
  const n = teks ? teks.split(',').filter(function (s) { return s.trim(); }).length : 0;
  if (!n) return 'Kosong = dianggap mampu SEMUA spesialis.';
  return n + ' spesialis dipilih. Pilih lagi item yang sama untuk menghapusnya.';
}

/** Cek ringan 1 sel assignment di Papan_Asistensi */
function validasiSelAsistensi_(sheet, row) {
  if (row < 2) return;
  const sel = sheet.getRange(row, PA_KOL_PERAWAT);
  const nama = String(sel.getValue()).trim();
  if (!nama) { sel.setBackground(null).setNote(''); return; }

  const info = sheet.getRange(row, 1, 1, PA_KOL_PERAWAT).getValues()[0];
  const tanggal = String(info[0]), shift = String(info[2]);
  const dokter = String(info[3]), spes = String(info[4]);

  const perawatMap = petaPerawat_();
  const liburMap = petaLiburPerNama_(SpreadsheetApp.getActiveSpreadsheet());
  const p = perawatMap[nama.toLowerCase()];
  const masalah = [];
  let level = 'ok';

  if (!p) {
    masalah.push('Perawat tidak ditemukan di master.');
    level = 'merah';
  } else {
    if (p.aktif === 'Tidak') { masalah.push('Perawat berstatus NONAKTIF.'); level = 'merah'; }

    // Bentrok jatah libur
    if (liburMap[p.nama] && liburMap[p.nama][tanggal]) {
      masalah.push('Bentrok: ' + p.nama + ' tercatat LIBUR di Papan_Libur tanggal ini.');
      level = 'merah';
    }

    // N-05: kapabilitas spesialis
    if (p.kapabilitasSet.length && p.kapabilitasSet.indexOf(spes) === -1) {
      masalah.push('N-05: belum punya kapabilitas asistensi ' + spes + '.');
      if (level !== 'merah') level = 'oranye';
    }

    // Scan sekali untuk dobel-assign & A-02
    const last = sheet.getLastRow();
    if (last > 1) {
      const semua = sheet.getRange(2, 1, last - 1, PA_KOL_PERAWAT).getValues();
      let adaNewLain = false;
      for (let i = 0; i < semua.length; i++) {
        const r = semua[i], rRow = i + 2;
        if (rRow === row) continue;
        const rNama = String(r[PA_KOL_PERAWAT - 1]).trim();
        if (!rNama) continue;
        if (rNama.toLowerCase() === nama.toLowerCase() &&
            String(r[0]) === tanggal && String(r[2]) === shift) {
          masalah.push('Dobel assign: sudah terjadwal di baris ' + rRow +
            ' (' + String(r[3]) + ') pada tanggal & shift yang sama.');
          level = 'merah';
        }
        if (String(r[0]) === tanggal && String(r[2]) === shift && String(r[3]) === dokter) {
          const rekan = perawatMap[rNama.toLowerCase()];
          if (rekan && rekan.kategori === 'New') adaNewLain = true;
        }
      }
      if (p.kategori === 'New' && adaNewLain) {
        masalah.push('A-02: dua perawat NEW berdampingan pada dokter & shift yang sama.');
        level = 'merah';
      }
    }
  }

  terapkanWarna_(sel, level, masalah);
}

/** Cek ringan 1 sel di Papan_Libur */
function validasiSelLibur_(sheet, row, col) {
  if (row < 2) return;
  const sel = sheet.getRange(row, col);
  const nama = String(sel.getValue()).trim();
  if (!nama) { sel.setBackground(null).setNote(''); return; }

  const masalah = [];
  let level = 'ok';

  const baris = sheet.getRange(row, 3, 1, PL_MAX_SLOT).getValues()[0];
  let muncul = 0;
  baris.forEach(function (v) {
    if (String(v).trim().toLowerCase() === nama.toLowerCase()) muncul++;
  });
  if (muncul > 1) { masalah.push('Nama yang sama tercatat 2x di tanggal ini.'); level = 'merah'; }

  const hari = String(sheet.getRange(row, 2).getValue()).toLowerCase();
  const weekend = String(getConfig('HARI_WEEKEND') || 'Sabtu,Minggu').split(',')
    .map(function (h) { return h.trim().toLowerCase(); });
  if (weekend.indexOf(hari) !== -1) {
    const p = petaPerawat_()[nama.toLowerCase()];
    if (p && p.level !== 'Gold') {
      masalah.push('N-03: hanya perawat GOLD yang boleh libur weekend.');
      level = 'merah';
    } else if (p) {
      masalah.push('Libur weekend Gold: maks. 1x per periode — cek total di 🔍 Cek Kepatuhan.');
      if (level === 'ok') level = 'oranye';
    }
  }

  terapkanWarna_(sel, level, masalah);
}

/** Cek ringan 1 sel di Papan_Receptionist */
function validasiSelReceptionist_(sheet, row, col) {
  if (row < 2) return;
  const sel = sheet.getRange(row, col);
  const nama = String(sel.getValue()).trim();

  const info = sheet.getRange(row, 1, 1, PR_KOL_SLOT1 + PR_MAX_SLOT - 1).getValues()[0];
  const tanggal = String(info[0]);
  const minimal = parseInt(info[4], 10) || 0;
  const slotBaris = [];
  for (let i = 0; i < PR_MAX_SLOT; i++) {
    const v = String(info[PR_KOL_SLOT1 - 1 + i] || '').trim();
    if (v) slotBaris.push(v);
  }

  if (!nama) {
    sel.setBackground(null).setNote('');
    // Kekurangan vs minimum tetap perlu ditandai
    if (slotBaris.length < minimal) {
      sel.setBackground(WARNA_WARNING)
        .setNote('Shift ini baru terisi ' + slotBaris.length + ' dari minimal ' + minimal + ' receptionist (R-03/R-04).');
    }
    return;
  }

  const masalah = [];
  let level = 'ok';

  const rcpMap = {};
  getMasterData('receptionist').forEach(function (r) { rcpMap[r.nama.toLowerCase()] = r; });
  const r = rcpMap[nama.toLowerCase()];
  if (!r) { masalah.push('Receptionist tidak ditemukan di master.'); level = 'merah'; }
  else if (r.aktif === 'Tidak') { masalah.push('Receptionist berstatus NONAKTIF.'); level = 'merah'; }

  // Duplikat dalam baris (shift yang sama)
  let muncul = 0;
  slotBaris.forEach(function (v) {
    if (v.toLowerCase() === nama.toLowerCase()) muncul++;
  });
  if (muncul > 1) { masalah.push('Nama yang sama 2x di shift ini.'); level = 'merah'; }

  // Bentrok libur
  const liburMap = petaLiburPerNama_(SpreadsheetApp.getActiveSpreadsheet());
  if (liburMap[nama] && liburMap[nama][tanggal]) {
    masalah.push('Bentrok: ' + nama + ' tercatat LIBUR di Papan_Libur tanggal ini.');
    level = 'merah';
  }

  // Kekurangan vs minimum
  if (slotBaris.length < minimal) {
    masalah.push('Shift ini baru ' + slotBaris.length + '/' + minimal + ' receptionist (R-03/R-04).');
    if (level === 'ok') level = 'oranye';
  }

  terapkanWarna_(sel, level, masalah);
}

/** Terapkan warna & catatan hasil validasi ke sel */
function terapkanWarna_(sel, level, masalah) {
  if (level === 'merah') sel.setBackground(WARNA_PELANGGARAN);
  else if (level === 'oranye') sel.setBackground(WARNA_WARNING);
  else sel.setBackground(WARNA_VALID);
  sel.setNote(masalah.join('\n'));
}

/* ================= 2. AUDIT PERIODE ================= */

/** Menu: 🔍 Cek Kepatuhan Periode — jalankan audit + tulis Rekap_Kepatuhan */
function cekKepatuhanPeriode() {
  const ui = SpreadsheetApp.getUi();
  const h = auditInti_();
  if (!h.ok) { ui.alert(h.pesan); return; }
  tulisRekapKepatuhan_(h);
  ui.alert(
    'Audit selesai',
    '• Perawat direkap: ' + h.rekap.length + '\n' +
    '• Receptionist direkap: ' + h.rekapRcp.length + '\n' +
    '• Slot asistensi kosong: ' + h.slotKosong + '\n' +
    '• Dokter tanpa perawat: ' + (h.dokterTanpaPerawat || 0) + '\n' +
    '• Shift receptionist kekurangan: ' + h.shiftKurang + '\n' +
    '• Total temuan: ' + h.pelanggaran.length + '\n\n' +
    'Detail lengkap ada di sheet "Rekap_Kepatuhan".',
    ui.ButtonSet.OK
  );
}

/**
 * AUDIT INTI — dipakai oleh 🔍 Cek Kepatuhan, 📊 Dashboard, dan ✅ Cek Final.
 * @return {Object} {ok, pesan?, rekap, rekapRcp, pelanggaran, slotKosong, shiftKurang}
 */
function auditInti_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (!pa || pa.getLastRow() < 2) return { ok: false, pesan: 'Papan_Asistensi masih kosong.' };

  const LIBUR_PRW = Number(getConfig('LIBUR_WAJIB_PERAWAT')) || 4;
  const LIBUR_RCP = Number(getConfig('LIBUR_WAJIB_RECEPTIONIST')) || 4;
  const LONG_PRW = Number(getConfig('LONGSHIFT_WAJIB_PERAWAT')) || 4;
  const LONG_RCP_SILVER = Number(getConfig('LONGSHIFT_WAJIB_RECEPT_SILVER')) || 4;
  const MAX_WKND_GOLD = Number(getConfig('MAX_LIBUR_WEEKEND_GOLD')) || 1;
  const weekend = String(getConfig('HARI_WEEKEND') || 'Sabtu,Minggu').split(',')
    .map(function (h) { return h.trim().toLowerCase(); });

  const perawatMap = petaPerawat_();
  const receptionistMap = {};
  getMasterData('receptionist').forEach(function (r) {
    receptionistMap[r.nama.toLowerCase()] = r;
  });
  const liburMap = petaLiburPerNama_(ss);

  // Set tanggal cuti yang DIMINTA (Request Jadwal) per nama → pengecualian blocker (keputusan #4)
  const reqSet = {}; // namaLower → { 'dd/MM/yyyy': true }
  try {
    const rq = bacaRequestCuti_();
    Object.keys(rq.staf || {}).forEach(function (nm) {
      const low = nm.toLowerCase(); reqSet[low] = reqSet[low] || {};
      (rq.staf[nm] || []).forEach(function (t) { reqSet[low][t] = true; });
    });
  } catch (e) {}
  const diminta_ = function (namaLower, tgl) { return !!(reqSet[namaLower] && reqSet[namaLower][tgl]); };

  // --- Papan asistensi ---
  const rows = pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues();
  const stat = {};
  const pelanggaran = [];
  let slotKosong = 0;
  const spesPraktik = {};

  const stMulai = function (nama) {
    if (!stat[nama]) stat[nama] = {
      shiftPagi: {}, shiftSiang: {}, spesDidapat: {}, totalSlot: 0,
      libur: 0, liburWeekend: 0
    };
    return stat[nama];
  };
  const grupDokter = {};
  const dokterSlot = {}; // tanggal|shift|dokter → {total, terisi} untuk metrik "dokter tanpa perawat"

  rows.forEach(function (r, i) {
    const baris = i + 2;
    const tanggal = String(r[0]), shift = String(r[2]);
    const dokter = String(r[3]), spes = String(r[4]);
    const nama = String(r[6]).trim();
    if (spes) spesPraktik[spes] = true;

    // Lacak kelengkapan asisten per dokter (termasuk slot kosong)
    const dk = tanggal + '|' + shift + '|' + dokter;
    if (!dokterSlot[dk]) dokterSlot[dk] = { total: 0, terisi: 0, tanggal: tanggal, shift: shift, dokter: dokter };
    dokterSlot[dk].total++;
    if (nama) dokterSlot[dk].terisi++;

    if (!nama) { slotKosong++; return; }
    const key = nama.toLowerCase();
    const p = perawatMap[key];
    if (!p) {
      pelanggaran.push(['Papan_Asistensi baris ' + baris, 'Master', nama + ' tidak ada di MD_Perawat']);
      return;
    }
    const s = stMulai(key);
    s.totalSlot++;
    if (shift === 'Pagi') s.shiftPagi[tanggal] = true; else s.shiftSiang[tanggal] = true;
    s.spesDidapat[spes] = true;

    if (liburMap[p.nama] && liburMap[p.nama][tanggal]) {
      pelanggaran.push(['Papan_Asistensi baris ' + baris, 'Libur',
        p.nama + ' bertugas padahal tercatat libur tanggal ' + tanggal]);
    }
    if (p.kapabilitasSet.length && p.kapabilitasSet.indexOf(spes) === -1) {
      pelanggaran.push(['Papan_Asistensi baris ' + baris, 'N-05',
        nama + ' belum berkapabilitas ' + spes + ' (warning)']);
    }

    const gk = tanggal + '|' + shift + '|' + dokter;
    if (!grupDokter[gk]) grupDokter[gk] = [];
    grupDokter[gk].push({ nama: nama, kategori: p.kategori, baris: baris });
  });

  Object.keys(grupDokter).forEach(function (gk) {
    const news = grupDokter[gk].filter(function (x) { return x.kategori === 'New'; });
    if (news.length >= 2) {
      pelanggaran.push(['Grup ' + gk, 'A-02',
        news.map(function (x) { return x.nama; }).join(' & ') + ' — perawat NEW berdampingan']);
    }
  });

  // --- Papan libur ---
  const pl = ss.getSheetByName('Papan_Libur');
  const liburRcp = {};
  if (pl && pl.getLastRow() > 1) {
    const lrows = pl.getRange(2, 1, pl.getLastRow() - 1, 2 + PL_MAX_SLOT).getValues();
    lrows.forEach(function (r, i) {
      const hari = String(r[1]).toLowerCase();
      const isWeekend = weekend.indexOf(hari) !== -1;
      const tglL = fmtTglCell_(r[0]);
      for (let c = 2; c < 2 + PL_MAX_SLOT; c++) {
        const nama = String(r[c]).trim();
        if (!nama) continue;
        const key = nama.toLowerCase();
        if (perawatMap[key]) {
          const s = stMulai(key);
          s.libur++;
          if (isWeekend) {
            s.liburWeekend++;
            if (perawatMap[key].level !== 'Gold') {
              // Bila libur weekend ini berasal dari Request Jadwal → pengecualian (warning), bukan blocker
              const wr = diminta_(key, tglL) ? ' (warning)' : '';
              pelanggaran.push(['Papan_Libur baris ' + (i + 2), 'N-03',
                nama + ' (Silver) libur weekend — hanya Gold yang boleh' + wr]);
            }
          }
        } else if (receptionistMap[key]) {
          liburRcp[key] = (liburRcp[key] || 0) + 1;
        } else {
          pelanggaran.push(['Papan_Libur baris ' + (i + 2), 'Master', nama + ' tidak ada di master']);
        }
      }
    });
  }

  // --- Papan receptionist (R-03/R-04, R-07, R-02, bentrok libur) ---
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  const rcpStat = {}; // nama lower → {pagi:{tgl}, siang:{tgl}, total}
  let shiftKurang = 0;
  if (pr && pr.getLastRow() > 1) {
    const prows = pr.getRange(2, 1, pr.getLastRow() - 1, PR_KOL_SLOT1 + PR_MAX_SLOT - 1).getValues();
    const duPerTanggal = {}; // tanggal → {Pagi:{du, isi}, Siang:{du, isi}}
    prows.forEach(function (r, i) {
      const baris = i + 2;
      const tanggal = String(r[0]), shift = String(r[2]);
      const duN = parseInt(r[3], 10) || 0;
      const minimal = parseInt(r[4], 10) || 0;
      const isi = [];
      for (let c = PR_KOL_SLOT1 - 1; c < PR_KOL_SLOT1 - 1 + PR_MAX_SLOT; c++) {
        const nm = String(r[c] || '').trim();
        if (nm) isi.push(nm);
      }
      if (!duPerTanggal[tanggal]) duPerTanggal[tanggal] = {};
      duPerTanggal[tanggal][shift] = { du: duN, isi: isi.length };

      if (isi.length < minimal) {
        shiftKurang++;
        pelanggaran.push(['Papan_Receptionist baris ' + baris, 'R-03/R-04',
          tanggal + ' ' + shift + ': terisi ' + isi.length + '/' + minimal]);
      }
      const dupCek = {};
      isi.forEach(function (nm) {
        const key = nm.toLowerCase();
        if (dupCek[key]) {
          pelanggaran.push(['Papan_Receptionist baris ' + baris, 'Dobel', nm + ' tercatat 2x di shift yang sama']);
        }
        dupCek[key] = true;
        const rMaster = receptionistMap[key];
        if (!rMaster) {
          pelanggaran.push(['Papan_Receptionist baris ' + baris, 'Master', nm + ' tidak ada di MD_Receptionist']);
          return;
        }
        if (liburMap[nm] && liburMap[nm][tanggal]) {
          pelanggaran.push(['Papan_Receptionist baris ' + baris, 'Libur',
            nm + ' bertugas padahal tercatat libur tanggal ' + tanggal]);
        }
        if (!rcpStat[key]) rcpStat[key] = { pagi: {}, siang: {}, total: 0 };
        rcpStat[key].total++;
        if (shift === 'Pagi') rcpStat[key].pagi[tanggal] = true;
        else rcpStat[key].siang[tanggal] = true;
      });
    });
    // R-02: shift dengan DU lebih banyak seharusnya receptionist-nya tidak lebih sedikit
    Object.keys(duPerTanggal).forEach(function (t) {
      const d = duPerTanggal[t];
      if (d.Pagi && d.Siang) {
        if (d.Pagi.du > d.Siang.du && d.Pagi.isi < d.Siang.isi) {
          pelanggaran.push([t, 'R-02', 'DU pagi lebih banyak tetapi receptionist pagi lebih sedikit (warning)']);
        }
        if (d.Siang.du > d.Pagi.du && d.Siang.isi < d.Pagi.isi) {
          pelanggaran.push([t, 'R-02', 'DU siang lebih banyak tetapi receptionist siang lebih sedikit (warning)']);
        }
      }
    });
  }

  // --- Rekap perawat ---
  const daftarSpesPraktik = Object.keys(spesPraktik).sort();
  const rekap = [];
  Object.keys(perawatMap).forEach(function (key) {
    const p = perawatMap[key];
    if (p.aktif === 'Tidak') return;
    const s = stat[key] || stMulai(key);

    let longshift = 0;
    Object.keys(s.shiftPagi).forEach(function (t) { if (s.shiftSiang[t]) longshift++; });

    const kurang = daftarSpesPraktik.filter(function (sp) { return !s.spesDidapat[sp]; });
    const mustahil = kurang.filter(function (sp) {
      return p.kapabilitasSet.length && p.kapabilitasSet.indexOf(sp) === -1;
    });

    // w:false = blocker, w:true = warning. (keputusan #2: N-02/N-03/A-04 = blocker; A-03 & over-jatah = warning)
    const catatan = [];
    if (s.totalSlot === 0) catatan.push({ t: 'A-07: perawat aktif belum dijadwalkan asistensi sama sekali', w: false });
    if (s.libur < LIBUR_PRW) catatan.push({ t: 'N-02: libur ' + s.libur + '/' + LIBUR_PRW, w: false });
    if (s.libur > LIBUR_PRW) catatan.push({ t: 'Libur melebihi jatah (' + s.libur + ')', w: true });
    if (p.level === 'Gold' && s.liburWeekend > MAX_WKND_GOLD) {
      catatan.push({ t: 'N-03: libur weekend ' + s.liburWeekend + 'x (maks ' + MAX_WKND_GOLD + ')', w: false });
    }
    if (longshift < LONG_PRW) catatan.push({ t: 'A-04: longshift ' + longshift + '/' + LONG_PRW, w: false });
    if (kurang.length) {
      catatan.push({ t: 'A-03: belum ke ' + kurang.join(', ') +
        (mustahil.length ? ' [TIDAK MUNGKIN: ' + mustahil.join(', ') + ' di luar kapabilitas]' : ''), w: true });
    }
    catatan.forEach(function (c) { pelanggaran.push([p.nama, c.t.split(':')[0], c.t + (c.w ? ' (warning)' : '')]); });

    // "patuh" hanya ditentukan pelanggaran blocker (w:false). A-03 (belum asistensi)
    // & over-jatah adalah warning → tak membuat tidak patuh, hanya info di kolomnya.
    const nIsu = catatan.filter(function (c) { return !c.w; }).length;
    rekap.push([
      p.id, p.nama, p.level, p.kategori, s.totalSlot,
      longshift + ' / ' + LONG_PRW,
      s.libur + ' / ' + LIBUR_PRW,
      s.liburWeekend,
      Object.keys(s.spesDidapat).sort().join(', ') || '-',
      kurang.join(', ') || '✓ lengkap',
      nIsu ? '⚠️ ' + nIsu + ' isu' : '✅ sesuai'
    ]);
  });

  // --- Rekap receptionist (R-06 + R-07) ---
  const rekapRcp = [];
  Object.keys(receptionistMap).forEach(function (key) {
    const r = receptionistMap[key];
    if (r.aktif === 'Tidak') return;
    const nLibur = liburRcp[key] || 0;
    const rs = rcpStat[key] || { pagi: {}, siang: {}, total: 0 };
    let longshift = 0;
    Object.keys(rs.pagi).forEach(function (t) { if (rs.siang[t]) longshift++; });

    // R-06 (libur) = blocker (keputusan #2); R-07 (longshift Silver) = warning
    const catatan = [];
    if (rs.total === 0) catatan.push({ t: 'R-09: resepsionis aktif belum dijadwalkan shift sama sekali', w: false });
    if (nLibur < LIBUR_RCP) catatan.push({ t: 'R-06: libur ' + nLibur + '/' + LIBUR_RCP, w: false });
    if (r.level === 'Silver' && longshift < LONG_RCP_SILVER) {
      catatan.push({ t: 'R-07: longshift ' + longshift + '/' + LONG_RCP_SILVER, w: true });
    }
    catatan.forEach(function (c) { pelanggaran.push([r.nama, c.t.split(':')[0], c.t + (c.w ? ' (warning)' : '')]); });

    const nIsuRcp = catatan.filter(function (c) { return !c.w; }).length; // hanya blocker (R-07 warning tak menghitung)
    rekapRcp.push([r.id, r.nama, r.level, rs.total,
      longshift + (r.level === 'Silver' ? ' / ' + LONG_RCP_SILVER : ''),
      nLibur + ' / ' + LIBUR_RCP,
      nIsuRcp ? '⚠️ ' + nIsuRcp + ' isu' : '✅ sesuai']);
  });

  // --- Metrik dokter yang tidak mendapat perawat ---
  let dokterTanpaPerawat = 0;   // terisi = 0 dari kebutuhan
  let dokterAsistenKurang = 0;  // 0 < terisi < total
  const daftarDokterKurang = [];
  Object.keys(dokterSlot).forEach(function (dk) {
    const d = dokterSlot[dk];
    if (d.terisi === 0) {
      dokterTanpaPerawat++;
      daftarDokterKurang.push([d.tanggal + ' ' + d.shift, 'A-05',
        d.dokter + ' — ' + d.total + ' slot asisten kosong (0/' + d.total + ' terisi)']);
    } else if (d.terisi < d.total) {
      dokterAsistenKurang++;
      daftarDokterKurang.push([d.tanggal + ' ' + d.shift, 'A-05',
        d.dokter + ' — ' + (d.total - d.terisi) + ' slot asisten kosong (' + d.terisi + '/' + d.total + ' terisi)']);
    }
  });
  // Masukkan ke daftar temuan
  daftarDokterKurang.forEach(function (t) { pelanggaran.push(t); });

  return {
    ok: true, rekap: rekap, rekapRcp: rekapRcp, pelanggaran: pelanggaran,
    slotKosong: slotKosong, shiftKurang: shiftKurang,
    dokterTanpaPerawat: dokterTanpaPerawat, dokterAsistenKurang: dokterAsistenKurang
  };
}

/** Tulis hasil audit ke sheet Rekap_Kepatuhan */
function tulisRekapKepatuhan_(h) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rekap = h.rekap, rekapRcp = h.rekapRcp, pelanggaran = h.pelanggaran;
  let rk = ss.getSheetByName('Rekap_Kepatuhan');
  if (!rk) rk = ss.insertSheet('Rekap_Kepatuhan');
  rk.clear();
  rk.getRange(1, 1).setValue('REKAP KEPATUHAN PERIODE — dibuat ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
    .setFontWeight('bold');
  rk.getRange(2, 1).setValue('Slot asistensi belum terisi: ' + h.slotKosong +
    ' | Shift receptionist kekurangan: ' + h.shiftKurang);
  rk.getRange(3, 1).setValue('Dokter tanpa perawat: ' + (h.dokterTanpaPerawat || 0) +
    ' | Dokter asisten belum lengkap: ' + (h.dokterAsistenKurang || 0))
    .setFontColor((h.dokterTanpaPerawat || 0) > 0 ? '#c5221f' : '#188038');

  const hdrP = ['ID', 'Perawat', 'Level', 'Kategori', 'Total Asisten Dokter', 'Longshift',
    'Libur', 'Libur Wknd', 'Spesialis Didapat', 'Spesialis Kurang', 'Status'];
  rk.getRange(4, 1, 1, hdrP.length).setValues([hdrP]);
  rk.getRange(4, 1, 1, hdrP.length).setBackground('#1a73e8').setFontColor('#fff').setFontWeight('bold');
  if (rekap.length) rk.getRange(5, 1, rekap.length, hdrP.length).setValues(rekap);

  let baris = 5 + rekap.length + 1;
  const hdrR = ['ID', 'Receptionist', 'Level', 'Total Shift', 'Longshift', 'Libur', 'Status'];
  rk.getRange(baris, 1, 1, hdrR.length).setValues([hdrR]);
  rk.getRange(baris, 1, 1, hdrR.length).setBackground('#1a73e8').setFontColor('#fff').setFontWeight('bold');
  if (rekapRcp.length) rk.getRange(baris + 1, 1, rekapRcp.length, hdrR.length).setValues(rekapRcp);

  baris = baris + 1 + rekapRcp.length + 1;
  rk.getRange(baris, 1).setValue('DAFTAR TEMUAN (' + pelanggaran.length + ')').setFontWeight('bold');
  if (pelanggaran.length) {
    rk.getRange(baris + 1, 1, 1, 3).setValues([['Lokasi/Staf', 'Aturan', 'Detail']]);
    rk.getRange(baris + 1, 1, 1, 3).setBackground('#5f6368').setFontColor('#fff').setFontWeight('bold');
    rk.getRange(baris + 2, 1, pelanggaran.length, 3).setValues(pelanggaran);
  }
  rk.setColumnWidth(2, 180);
  rk.setColumnWidths(9, 2, 220);
  rk.setFrozenRows(4);
}

/* ================= HELPER ================= */

/** Peta perawat: nama lowercase → {id, nama, level, kategori, kapabilitasSet[], aktif} */
function petaPerawat_() {
  const map = {};
  getMasterData('perawat').forEach(function (p) {
    map[p.nama.toLowerCase()] = {
      id: p.id, nama: p.nama, level: p.level, kategori: p.kategori, aktif: p.aktif,
      kapabilitasSet: String(p.kapabilitas || '').split(',')
        .map(function (s) { return s.trim(); })
        .filter(function (s) { return s; })
    };
  });
  return map;
}