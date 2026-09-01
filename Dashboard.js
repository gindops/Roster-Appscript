/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Fase 5: Dashboard & Finalisasi
 * File: Dashboard.gs
 *
 * 📊 Perbarui Dashboard : ringkasan eksekutif periode (KPI)
 * ✅ Cek Final          : quality gate — jadwal dinyatakan FINAL
 *                         hanya bila 0 pelanggaran hard constraint
 * 🖨️ Buat Jadwal Final  : tampilan cetak per tanggal yang mudah
 *                         dibaca (untuk ditempel/dibagikan)
 * ============================================================
 */

/* ================= 📊 DASHBOARD ================= */

/** Menu: 📊 Perbarui Dashboard */
function buatDashboard() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const h = auditInti_();
  if (!h.ok) { ui.alert(h.pesan); return; }

  // Hitung KPI
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  const totalSlot = pa.getLastRow() - 1;
  const terisi = totalSlot - h.slotKosong;
  const prwPatuh = h.rekap.filter(function (r) { return String(r[10]).indexOf('✅') === 0; }).length;
  const rcpPatuh = h.rekapRcp.filter(function (r) { return String(r[6]).indexOf('✅') === 0; }).length;

  // Pelanggaran per kategori aturan
  const perAturan = {};
  h.pelanggaran.forEach(function (p) {
    perAturan[p[1]] = (perAturan[p[1]] || 0) + 1;
  });

  const blockers = hitungBlocker_(h);
  const statusJadwal = String(getConfig('STATUS_JADWAL') || 'DRAF');

  // Tulis Dashboard
  let db = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!db) db = ss.insertSheet(SHEETS.DASHBOARD);
  db.clear();

  const pMulai = getConfig('PERIODE_MULAI') || '-';
  const pAkhir = getConfig('PERIODE_AKHIR') || '-';

  db.getRange(1, 1).setValue('📊 DASHBOARD PENJADWALAN — PERIODE ' + pMulai + ' s.d. ' + pAkhir)
    .setFontSize(14).setFontWeight('bold');
  db.getRange(2, 1).setValue('Diperbarui: ' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
    .setFontColor('#5f6368');

  const kpi = [
    ['STATUS JADWAL', statusJadwal],
    ['Blocker finalisasi (hard constraint)', blockers.total],
    ['', ''],
    ['Slot asistensi terisi', terisi + ' / ' + totalSlot + '  (' + persen_(terisi, totalSlot) + ')'],
    ['Dokter tanpa perawat', (h.dokterTanpaPerawat || 0)],
    ['Dokter asisten belum lengkap', (h.dokterAsistenKurang || 0)],
    ['Shift receptionist kekurangan', h.shiftKurang],
    ['Perawat sesuai penuh', prwPatuh + ' / ' + h.rekap.length + '  (' + persen_(prwPatuh, h.rekap.length) + ')'],
    ['Receptionist sesuai penuh', rcpPatuh + ' / ' + h.rekapRcp.length + '  (' + persen_(rcpPatuh, h.rekapRcp.length) + ')'],
    ['Total temuan audit', h.pelanggaran.length]
  ];
  db.getRange(4, 1, kpi.length, 2).setValues(kpi);
  db.getRange(4, 1, kpi.length, 1).setFontWeight('bold');
  db.getRange(4, 2).setFontSize(13).setFontWeight('bold')
    .setFontColor(statusJadwal.indexOf('FINAL') === 0 ? '#188038' : '#b06000');
  db.getRange(5, 2).setFontWeight('bold')
    .setFontColor(blockers.total === 0 ? '#188038' : '#c5221f');

  // Tabel temuan per aturan
  let baris = 4 + kpi.length + 2;
  db.getRange(baris, 1).setValue('TEMUAN PER ATURAN').setFontWeight('bold');
  baris++;
  const aturanKeys = Object.keys(perAturan).sort();
  if (aturanKeys.length) {
    db.getRange(baris, 1, 1, 2).setValues([['Aturan', 'Jumlah']])
      .setBackground('#1a73e8').setFontColor('#fff').setFontWeight('bold');
    const rowsAturan = aturanKeys.map(function (k) { return [k, perAturan[k]]; });
    db.getRange(baris + 1, 1, rowsAturan.length, 2).setValues(rowsAturan);
    baris += 1 + rowsAturan.length;
  } else {
    db.getRange(baris, 1).setValue('✅ Tidak ada temuan — jadwal bersih.').setFontColor('#188038');
    baris++;
  }

  baris += 1;
  db.getRange(baris, 1).setValue('Detail per staf: sheet "Rekap_Kepatuhan" · Finalisasi: menu ✅ Cek Final')
    .setFontColor('#5f6368').setFontStyle('italic');

  db.setColumnWidth(1, 300);
  db.setColumnWidth(2, 260);
  ss.setActiveSheet(db);

  laporOtomatis_('Dashboard', 'Dashboard diperbarui.\nBlocker finalisasi: ' + blockers.total +
    ' | Temuan total: ' + h.pelanggaran.length);
}

/* ================= ✅ CEK FINAL (QUALITY GATE) ================= */

/**
 * Menu: ✅ Cek Final & Kunci Jadwal.
 * Jadwal dinyatakan FINAL hanya bila 0 blocker hard constraint.
 * Warning (N-05, R-02) tidak memblokir tetapi dilaporkan.
 */
function cekFinal() {
  const ui = SpreadsheetApp.getUi();
  const h = auditInti_();
  if (!h.ok) { ui.alert(h.pesan); return; }
  tulisRekapKepatuhan_(h);

  const blockers = hitungBlocker_(h);

  if (blockers.total > 0) {
    setConfig_('STATUS_JADWAL', 'BELUM FINAL — ' + blockers.total + ' blocker', 'Status quality gate');
    buatDashboard();
    ui.alert('❌ Jadwal BELUM layak final',
      'Masih ada ' + blockers.total + ' pelanggaran hard constraint:\n\n' +
      blockers.daftar.slice(0, 15).map(function (b) { return '• ' + b; }).join('\n') +
      (blockers.daftar.length > 15 ? '\n… dan ' + (blockers.daftar.length - 15) + ' lainnya.' : '') +
      '\n\nPerbaiki dulu (lihat Rekap_Kepatuhan), lalu jalankan ✅ Cek Final lagi.',
      ui.ButtonSet.OK);
    return;
  }

  const stempel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  setConfig_('STATUS_JADWAL', 'FINAL ✓ ' + stempel, 'Status quality gate');
  try { simpanArsipJadwal_(stempel); } catch (e) {}   // arsipkan snapshot periode ini
  buatDashboard();
  const wrn = h.pelanggaran.length; // blocker 0 → semua sisa temuan adalah warning
  ui.alert('✅ Jadwal FINAL',
    'Semua hard constraint terpenuhi — jadwal dinyatakan FINAL (' + stempel + ').' +
    (wrn > 0 ? '\n\nCatatan: masih ada ' + wrn + ' warning non-blocker (N-05/R-02) di Rekap_Kepatuhan.' : '') +
    '\n\nLangkah berikutnya: 🖨️ Buat Jadwal Final untuk dicetak/dibagikan.',
    ui.ButtonSet.OK);
}

/** Klasifikasi blocker: semua temuan kecuali warning/info + slot/shift kosong.
 *  Mengembalikan { total, daftar:[string], rinci:[{kode,keluarga,lokasi,detail}] }.
 *  daftar kini menyertakan lokasi/nama (dari p[0]) dan membuang kode ganda. */
function hitungBlocker_(h) {
  const daftar = [];
  const rinci = [];
  const tambah = function (kode, lokasi, detail) {
    const k = String(kode || '').trim();
    let det = String(detail || '').replace(/\s*\((warning|info)\)/g, '').trim();
    // buang prefiks "KODE:" / "KODE -" di awal detail agar tak tampil dua kali
    if (k) det = det.replace(new RegExp('^' + k.replace(/[-\/]/g, '\\$&') + '\\s*[:\\-–]?\\s*'), '').trim();
    const lok = String(lokasi || '').trim();
    const teks = k + (lok ? ' · ' + lok : '') + (det ? ' — ' + det : '');
    daftar.push(teks);
    rinci.push({ kode: k, keluarga: (typeof familiKode_ === 'function' ? familiKode_(k) : 'D'), lokasi: lok, detail: det });
  };
  h.pelanggaran.forEach(function (p) {
    const detail = String(p[2]);
    if (detail.indexOf('(warning)') !== -1) return; // N-05, R-02 → bukan blocker
    if (detail.indexOf('(info)') !== -1) return;     // metrik dokter (sudah tercakup A-05)
    tambah(p[1], p[0], detail);                       // p[0] = nama/lokasi, p[1] = kode
  });
  // Slot asisten kosong (A-05) & kekurangan shift resepsionis (R-03/R-04) TIDAK ditambah
  // sebagai agregat di sini — keduanya sudah tercatat rinci (tanggal · dokter/shift) di pelanggaran.
  return { total: daftar.length, daftar: daftar, rinci: rinci };
}

/* ================= 🖨️ JADWAL FINAL (PRINT VIEW) ================= */

/** Menu: 🖨️ Buat Jadwal Final (print view) */
function buatJadwalFinal() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (!pa || pa.getLastRow() < 2) { ui.alert('Papan_Asistensi masih kosong.'); return; }

  // Kumpulkan asistensi: tanggal → shift → dokter → [perawat]
  const agenda = {};
  pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues().forEach(function (r) {
    if (!r[0]) return;
    const t = String(r[0]), hari = String(r[1]), shift = String(r[2]);
    const dokter = String(r[3]), spes = String(r[4]);
    const perawat = String(r[6]).trim();
    if (!agenda[t]) agenda[t] = { hari: hari, Pagi: {}, Siang: {}, rcp: { Pagi: [], Siang: [] }, libur: [] };
    const key = dokter + ' (' + spes + ')';
    if (!agenda[t][shift][key]) agenda[t][shift][key] = [];
    agenda[t][shift][key].push(perawat || '—');
  });

  // Receptionist per tanggal-shift
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (pr && pr.getLastRow() > 1) {
    pr.getRange(2, 1, pr.getLastRow() - 1, PR_KOL_SLOT1 + PR_MAX_SLOT - 1).getValues().forEach(function (r) {
      const t = String(r[0]), shift = String(r[2]);
      if (!agenda[t]) return;
      for (let c = PR_KOL_SLOT1 - 1; c < PR_KOL_SLOT1 - 1 + PR_MAX_SLOT; c++) {
        const nm = String(r[c] || '').trim();
        if (nm) agenda[t].rcp[shift].push(nm);
      }
    });
  }

  // Libur per tanggal
  const pl = ss.getSheetByName('Papan_Libur');
  if (pl && pl.getLastRow() > 1) {
    pl.getRange(2, 1, pl.getLastRow() - 1, 2 + PL_MAX_SLOT).getValues().forEach(function (r) {
      const t = String(r[0]);
      if (!agenda[t]) return;
      for (let c = 2; c < 2 + PL_MAX_SLOT; c++) {
        const nm = String(r[c] || '').trim();
        if (nm) agenda[t].libur.push(nm);
      }
    });
  }

  // Tulis sheet Jadwal_Final
  let jf = ss.getSheetByName('Jadwal_Final');
  if (!jf) jf = ss.insertSheet('Jadwal_Final');
  jf.clear();

  const status = String(getConfig('STATUS_JADWAL') || 'DRAF');
  const judul = 'JADWAL ASISTENSI & RECEPTIONIST — PERIODE ' +
    (getConfig('PERIODE_MULAI') || '') + ' s.d. ' + (getConfig('PERIODE_AKHIR') || '') +
    '   [' + status + ']';
  const out = [[judul, '', '']];
  const formats = [{ r: 0, gaya: 'judul' }];

  Object.keys(agenda).sort(function (a, b) {
    return kunciTgl_(a) < kunciTgl_(b) ? -1 : 1;
  }).forEach(function (t) {
    const a = agenda[t];
    out.push(['', '', '']);
    formats.push({ r: out.length - 1, gaya: 'spasi' });
    out.push(['📅 ' + a.hari + ', ' + t, '', '']);
    formats.push({ r: out.length - 1, gaya: 'tanggal' });

    ['Pagi', 'Siang'].forEach(function (shift) {
      const dokters = Object.keys(a[shift]).sort();
      dokters.forEach(function (d, i) {
        out.push([i === 0 ? '  ' + shift.toUpperCase() : '', d, a[shift][d].join(', ')]);
        formats.push({ r: out.length - 1, gaya: 'isi' });
      });
      if (a.rcp[shift].length) {
        out.push([dokters.length ? '' : '  ' + shift.toUpperCase(), 'Receptionist', a.rcp[shift].join(', ')]);
        formats.push({ r: out.length - 1, gaya: 'rcp' });
      }
    });
    if (a.libur.length) {
      out.push(['  LIBUR', '', a.libur.join(', ')]);
      formats.push({ r: out.length - 1, gaya: 'libur' });
    }
  });

  jf.getRange(1, 1, out.length, 3).setValues(out);
  formats.forEach(function (f) {
    const rng = jf.getRange(f.r + 1, 1, 1, 3);
    if (f.gaya === 'judul') rng.setFontWeight('bold').setFontSize(13);
    else if (f.gaya === 'tanggal') rng.setFontWeight('bold').setBackground('#e8f0fe');
    else if (f.gaya === 'rcp') rng.setFontColor('#7b1fa2');
    else if (f.gaya === 'libur') rng.setFontColor('#b06000').setFontStyle('italic');
  });
  jf.setColumnWidth(1, 110);
  jf.setColumnWidth(2, 240);
  jf.setColumnWidth(3, 420);
  ss.setActiveSheet(jf);

  ui.alert('🖨️ Jadwal Final siap',
    'Sheet "Jadwal_Final" berisi jadwal per tanggal yang siap dicetak (File → Print) atau dibagikan.\n' +
    'Status jadwal: ' + status, ui.ButtonSet.OK);
}

/** Persentase aman */
function persen_(a, b) {
  if (!b) return '0%';
  return Math.round(a * 100 / b) + '%';
}