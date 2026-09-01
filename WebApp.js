/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Web App (Fase Web-1)
 * File: WebApp.gs
 *
 * Lapisan web app yang menjadikan spreadsheet sebagai DATABASE.
 * Di-deploy sebagai Web App (Deploy → New deployment → Web app,
 * "Execute as: Me", "Who has access: Anyone with Google account").
 * Kontrol akses per-email via sheet Config_Akses (Admin/SPV/Viewer).
 *
 * Semua endpoint di sini UI-FREE (tidak memanggil getUi()),
 * memakai ulang pure-function algoritma yang sudah teruji, sehingga
 * sistem menu spreadsheet tetap utuh sebagai fallback admin.
 * ============================================================
 */

const SHEET_AKSES = 'Config_Akses';

/** Entry point web app — file HTML bernama "index" */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Penjadwalan Audy Dental')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://ssl.gstatic.com/docs/spreadsheets/favicon3.ico');
}

/* ================= AKSES / ROLE ================= */

/** Peran user aktif berdasarkan Config_Akses. Default Viewer bila tak terdaftar. */
function getPeranUser_() {
  const email = String((Session.getActiveUser() && Session.getActiveUser().getEmail()) || '').toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_AKSES);
  if (!sh || sh.getLastRow() < 2) return { email: email, peran: 'Admin', nama: '', terdaftar: true }; // bootstrap: belum ada daftar → Admin
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === email && email) {
      const p = String(rows[i][1]).trim();
      return { email: email, peran: (p === 'Admin' || p === 'SPV' || p === 'Viewer') ? p : 'Viewer', nama: String(rows[i][2] || '').trim(), terdaftar: true };
    }
  }
  return { email: email, peran: '', nama: '', terdaftar: false }; // email tak terdaftar → akses ditolak
}

/** Konteks awal untuk frontend */
function getWebContext() {
  const u = getPeranUser_();
  if (!u.terdaftar) return { boleh: false, email: u.email || '(anonim)' };
  return {
    boleh: true,
    email: u.email || '(anonim)',
    peran: u.peran,
    nama: u.nama || '',
    modul: modulUntukPeran_(u.peran),
    periodeMulai: String(getConfig('PERIODE_MULAI') || ''),
    periodeAkhir: String(getConfig('PERIODE_AKHIR') || ''),
    status: String(getConfig('STATUS_JADWAL') || 'DRAF')
  };
}

/** Guard: lempar error bila peran user tidak termasuk yang diizinkan */
function wajibPeran_(daftar) {
  const u = getPeranUser_();
  if (daftar.indexOf(u.peran) === -1) {
    throw new Error('Akses ditolak untuk peran ' + u.peran + '. Butuh: ' + daftar.join('/') + '.');
  }
  return u;
}

/* ================= AKSES MODUL (per peran) ================= */

/** Baca matriks akses modul dari Config_Modul; fallback ke MODUL_DEFAULT. */
function bacaAksesModul_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEETS.CONFIG_MODUL);
  const akses = { Admin: {}, SPV: {}, Viewer: {} };
  const labels = {};
  const src = (sh && sh.getLastRow() > 1) ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : MODUL_DEFAULT;
  src.forEach(function (r) {
    const id = String(r[0]).trim(); if (!id) return;
    labels[id] = String(r[1]) || id;
    akses.Admin[id] = String(r[2]).trim().toLowerCase() === 'ya';
    akses.SPV[id] = String(r[3]).trim().toLowerCase() === 'ya';
    akses.Viewer[id] = String(r[4]).trim().toLowerCase() === 'ya';
  });
  return { akses: akses, labels: labels };
}

/** Daftar id modul yang boleh diakses peran tsb. */
function modulUntukPeran_(peran) {
  const m = bacaAksesModul_().akses[peran] || {};
  return Object.keys(m).filter(function (id) { return m[id]; });
}

/** Guard: pastikan peran user aktif punya akses modul tsb (modul aktif = akses penuh). */
function wajibModul_(modulId) {
  const u = getPeranUser_();
  const m = bacaAksesModul_().akses[u.peran] || {};
  if (!m[modulId]) throw new Error('Akses ditolak: modul "' + modulId + '" tidak diaktifkan untuk peran ' + u.peran + '.');
  return u;
}

/** Matriks akses modul untuk halaman Pengaturan (Admin). */
function getModulConfigWeb() {
  wajibPeran_(['Admin']);
  const d = bacaAksesModul_();
  const list = MODUL_DEFAULT.map(function (r) {
    const id = r[0];
    return { id: id, label: r[1], Admin: !!d.akses.Admin[id], SPV: !!d.akses.SPV[id], Viewer: !!d.akses.Viewer[id] };
  });
  return { ok: true, modul: list };
}

/** Simpan matriks akses modul (Admin). */
function setModulConfigWeb(rows) {
  wajibPeran_(['Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEETS.CONFIG_MODUL);
  if (!sh) sh = ss.insertSheet(SHEETS.CONFIG_MODUL);
  sh.clear();
  sh.getRange(1, 1, 1, 5).setValues([['Modul', 'Label', 'Admin', 'SPV', 'Viewer']]);
  const out = (rows || []).map(function (r) {
    return [r.id, r.label || r.id, r.Admin ? 'Ya' : 'Tidak', r.SPV ? 'Ya' : 'Tidak', r.Viewer ? 'Ya' : 'Tidak'];
  });
  if (out.length) sh.getRange(2, 1, out.length, 5).setValues(out);
  return { ok: true, pesan: 'Akses modul disimpan. Pengguna cukup me-refresh web app.' };
}

/* ================= EMPLOYEE LISTING (Config_Akses) ================= */

/** Validasi & normalisasi input akses. */
function normAkses_(data) {
  const email = String((data && data.email) || '').trim().toLowerCase();
  const peran = String((data && data.peran) || '').trim();
  const nama = String((data && data.nama) || '').trim();
  if (!email || email.indexOf('@') === -1 || email.indexOf('.') === -1) return { err: 'Email tidak valid.' };
  if (['Admin', 'SPV', 'Viewer'].indexOf(peran) === -1) return { err: 'Peran harus Admin, SPV, atau Viewer.' };
  return { email: email, peran: peran, nama: nama };
}

/** Baris berisi email (case-insensitive); -1 bila tak ada. */
function cariBarisAkses_(sh, email) {
  if (!sh || sh.getLastRow() < 2) return -1;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toLowerCase() === email) return i + 2;
  }
  return -1;
}

/** Jumlah baris berperan Admin. */
function jumlahAdmin_(sh) {
  if (!sh || sh.getLastRow() < 2) return 0;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  return rows.filter(function (r) { return String(r[0]).trim() && String(r[1]).trim() === 'Admin'; }).length;
}

/** Daftar email berhak akses (Admin). */
function getAksesListWeb() {
  wajibPeran_(['Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_AKSES);
  if (!sh || sh.getLastRow() < 2) return { ok: true, list: [] };
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const list = [];
  rows.forEach(function (r, i) {
    const email = String(r[0]).trim();
    if (!email) return;
    const p = String(r[1]).trim();
    list.push({ baris: i + 2, email: email, peran: (p === 'Admin' || p === 'SPV' || p === 'Viewer') ? p : 'Viewer', nama: String(r[2] || '').trim() });
  });
  return { ok: true, list: list };
}

/** Tambah email berhak akses (Admin). */
function addAksesWeb(data) {
  wajibPeran_(['Admin']);
  const n = normAkses_(data); if (n.err) return { ok: false, pesan: n.err };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_AKSES);
  if (!sh) { sh = ss.insertSheet(SHEET_AKSES); sh.getRange(1, 1, 1, 3).setValues([['Email', 'Peran', 'Nama']]); }
  if (cariBarisAkses_(sh, n.email) !== -1) return { ok: false, pesan: 'Email ' + n.email + ' sudah terdaftar.' };
  sh.appendRow([n.email, n.peran, n.nama]);
  return { ok: true, pesan: 'Akses ' + n.email + ' (' + n.peran + ') ditambahkan.' };
}

/** Perbarui 1 baris akses (Admin). */
function updateAksesWeb(baris, data) {
  wajibPeran_(['Admin']);
  const n = normAkses_(data); if (n.err) return { ok: false, pesan: n.err };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_AKSES);
  if (!sh || baris < 2 || baris > sh.getLastRow()) return { ok: false, pesan: 'Baris tidak valid.' };
  const lain = cariBarisAkses_(sh, n.email);
  if (lain !== -1 && lain !== baris) return { ok: false, pesan: 'Email ' + n.email + ' sudah dipakai baris lain.' };
  const lamaPeran = String(sh.getRange(baris, 2).getValue()).trim();
  if (lamaPeran === 'Admin' && n.peran !== 'Admin' && jumlahAdmin_(sh) <= 1) return { ok: false, pesan: 'Tidak bisa menurunkan peran Admin terakhir — tambah Admin lain dulu.' };
  sh.getRange(baris, 1, 1, 3).setValues([[n.email, n.peran, n.nama]]);
  return { ok: true, pesan: 'Akses ' + n.email + ' diperbarui.' };
}

/** Hapus 1 baris akses (Admin). Pengaman: tak boleh hapus Admin terakhir. */
function deleteAksesWeb(baris) {
  wajibPeran_(['Admin']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_AKSES);
  if (!sh || baris < 2 || baris > sh.getLastRow()) return { ok: false, pesan: 'Baris tidak valid.' };
  const peran = String(sh.getRange(baris, 2).getValue()).trim();
  if (peran === 'Admin' && jumlahAdmin_(sh) <= 1) return { ok: false, pesan: 'Tidak bisa menghapus Admin terakhir — sistem harus punya minimal 1 Admin.' };
  sh.deleteRow(baris);
  return { ok: true, pesan: 'Akses dihapus.' };
}

/**
 * Upsert email ke Config_Akses sebagai Viewer bila belum terdaftar.
 * Dipanggil otomatis saat master data (dokter/perawat/resepsionis) disimpan.
 * Tak menurunkan peran email yang sudah ada. Tanpa guard peran karena
 * dipicu oleh alur master-data yang sudah dijaga wajibModul_('master').
 * @return {boolean} true bila baris baru ditambahkan.
 */
function daftarkanAksesViewer_(email, nama) {
  const e = String(email || '').trim().toLowerCase();
  if (!e || e.indexOf('@') === -1 || e.indexOf('.') === -1) return false; // email tak valid → lewati
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_AKSES);
  if (!sh) { sh = ss.insertSheet(SHEET_AKSES); sh.getRange(1, 1, 1, 3).setValues([['Email', 'Peran', 'Nama']]); }
  if (cariBarisAkses_(sh, e) !== -1) return false; // sudah terdaftar → jangan ubah peran
  sh.appendRow([e, 'Viewer', String(nama || '').trim()]);
  return true;
}

/* ================= DASHBOARD ================= */

function getDashboardWeb() {
  const h = auditInti_();
  if (!h.ok) return { ok: false, pesan: h.pesan };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  const totalSlot = pa ? Math.max(pa.getLastRow() - 1, 0) : 0;
  const terisi = totalSlot - h.slotKosong;
  const prwPatuh = h.rekap.filter(function (r) { return String(r[10]).indexOf('✅') === 0; }).length;
  const rcpPatuh = h.rekapRcp.filter(function (r) { return String(r[6]).indexOf('✅') === 0; }).length;
  const blockers = hitungBlocker_(h);
  return {
    ok: true,
    status: String(getConfig('STATUS_JADWAL') || 'DRAF'),
    periode: (getConfig('PERIODE_MULAI') || '-') + ' s.d. ' + (getConfig('PERIODE_AKHIR') || '-'),
    blocker: blockers.total,
    slotTerisi: terisi, slotTotal: totalSlot,
    dokterTanpaPerawat: h.dokterTanpaPerawat || 0,
    dokterAsistenKurang: h.dokterAsistenKurang || 0,
    shiftKurang: h.shiftKurang,
    prwPatuh: prwPatuh, prwTotal: h.rekap.length,
    rcpPatuh: rcpPatuh, rcpTotal: h.rekapRcp.length,
    temuan: h.pelanggaran.length
  };
}

/* ================= REKAP KEPATUHAN (live audit) ================= */

/**
 * Rekap kepatuhan LANGSUNG dari kondisi papan terkini (auditInti_()).
 * Karena dihitung ulang tiap dipanggil, setiap perubahan manual di papan
 * (asistensi/libur/receptionist) otomatis tercermin saat halaman di-refresh.
 */
function getRekapKepatuhanWeb() {
  const h = auditInti_();
  if (!h.ok) return { ok: false, pesan: h.pesan };
  const blockers = hitungBlocker_(h);
  const perawat = h.rekap.map(function (r) {
    return {
      id: r[0], nama: r[1], level: r[2], kategori: r[3], totalSlot: r[4],
      longshift: String(r[5]), libur: String(r[6]), liburWknd: r[7],
      spesDidapat: String(r[8]), spesKurang: String(r[9]), status: String(r[10]),
      patuh: String(r[10]).indexOf('✅') === 0
    };
  });
  const rcp = h.rekapRcp.map(function (r) {
    return {
      id: r[0], nama: r[1], level: r[2], totalSlot: r[3],
      longshift: String(r[4]), libur: String(r[5]), status: String(r[6]),
      patuh: String(r[6]).indexOf('✅') === 0
    };
  });
  const temuan = h.pelanggaran.map(function (p) {
    const detail = String(p[2]);
    const jenis = (detail.indexOf('(warning)') !== -1 || detail.indexOf('(info)') !== -1) ? 'warning' : 'blocker';
    const kode = String(p[1]);
    return { lokasi: String(p[0]), aturan: kode, keluarga: familiKode_(kode), keluargaNama: NAMA_KELOMPOK[familiKode_(kode)], detail: detail, jenis: jenis };
  });
  return {
    ok: true,
    status: String(getConfig('STATUS_JADWAL') || 'DRAF'),
    periode: (getConfig('PERIODE_MULAI') || '-') + ' s.d. ' + (getConfig('PERIODE_AKHIR') || '-'),
    blocker: blockers.total,
    slotKosong: h.slotKosong, shiftKurang: h.shiftKurang,
    prwPatuh: perawat.filter(function (x) { return x.patuh; }).length,
    rcpPatuh: rcp.filter(function (x) { return x.patuh; }).length,
    perawat: perawat, rcp: rcp, temuan: temuan
  };
}

/* ================= SET PERIODE ================= */

function setPeriodeWeb(bulan, tahun) {
  wajibModul_('aksi');
  bulan = parseInt(bulan, 10); tahun = parseInt(tahun, 10);
  if (!bulan || bulan < 1 || bulan > 12 || !tahun) return { ok: false, pesan: 'Bulan/tahun tidak valid.' };
  const tglMulai = Number(getConfig('TANGGAL_MULAI_PERIODE')) || 21;
  const mulai = new Date(tahun, bulan - 1, tglMulai);
  const akhir = new Date(tahun, bulan, tglMulai - 1);
  setConfig_('PERIODE_MULAI', formatTgl_(mulai), 'Tanggal mulai periode aktif');
  setConfig_('PERIODE_AKHIR', formatTgl_(akhir), 'Tanggal akhir periode aktif');
  setConfig_('STATUS_JADWAL', 'DRAF', 'Status quality gate');
  return { ok: true, pesan: 'Periode ' + formatTgl_(mulai) + ' s.d. ' + formatTgl_(akhir) + ' diaktifkan.' };
}

/* ================= GENERATE SEMUA (WEB-SAFE) ================= */

/** Jalankan seluruh generasi memakai pure-core, tanpa dialog. */
function generateSemuaWeb() {
  wajibModul_('aksi');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pMulai = parseTgl_(getConfig('PERIODE_MULAI'));
  const pAkhir = parseTgl_(getConfig('PERIODE_AKHIR'));
  if (!pMulai || !pAkhir) return { ok: false, pesan: 'Periode belum diset.' };

  const log = [];

  // 1) Jadwal Dokter dari master + saring cuti dokter
  const dokRows = getMasterData('dokter').filter(function (d) { return d.aktif !== 'Tidak'; });
  const polaValid = [];
  dokRows.forEach(function (d) {
    const hariIdx = NAMA_HARI.indexOf(String(d.hariPraktek || '').toUpperCase().trim());
    const shift = String(d.kategoriShift || '').replace(/^Shift\s*/i, '').trim();
    const nurse = parseInt(d.kebutuhanAsisten, 10);
    if (hariIdx === -1 || (shift !== 'Pagi' && shift !== 'Siang') || !nurse) return;
    polaValid.push({ idDokter: d.id, nama: d.nama, spes: d.spesialisasi, hariIdx: hariIdx, shift: shift, nurse: nurse });
  });
  const jdRaw = expandPola_(polaValid, pMulai, pAkhir);
  const jdSaring = saringDokterCuti_(jdRaw);
  tulisJadwal_(ss, jdSaring.rows);
  log.push('Jadwal Dokter: ' + jdSaring.rows.length + ' baris' + (jdSaring.dibuang ? ' (' + jdSaring.dibuang + ' dibuang krn cuti)' : ''));

  // 2) Papan Libur (auto) — reuse pure susunLibur_
  const libur = generateLiburCore_(ss, pMulai, pAkhir);
  log.push('Papan Libur: perawat penuh ' + libur.sukses.perawat + ', receptionist penuh ' + libur.sukses.receptionist);

  // 3) Papan Asistensi (buat + isi)
  const asis = generateAsistensiCore_(ss);
  log.push('Jadwal Asistensi: ' + asis.terisi + '/' + asis.total + ' slot terisi');

  // 4) Papan Resepsionis (buat + isi)
  const rcp = generateResepsionisCore_(ss);
  log.push('Jadwal Resepsionis: ' + rcp.baris + ' baris, kekurangan ' + rcp.kurang);

  setConfig_('STATUS_JADWAL', 'DRAF', 'Status quality gate');
  return { ok: true, log: log };
}

/** Core auto-libur (UI-free) */
function generateLiburCore_(ss, pMulai, pAkhir) {
  const weekend = String(getConfig('HARI_WEEKEND') || 'Sabtu,Minggu').split(',')
    .map(function (h) { return h.trim().toLowerCase(); });
  const tanggalList = [];
  for (let d = new Date(pMulai); d <= pAkhir; d.setDate(d.getDate() + 1)) {
    const hari = NAMA_HARI[d.getDay()].charAt(0) + NAMA_HARI[d.getDay()].slice(1).toLowerCase();
    tanggalList.push({ tanggal: formatTgl_(d), hari: hari, isWeekend: weekend.indexOf(hari.toLowerCase()) !== -1,
      minggu: Math.floor((d - pMulai) / (7 * 24 * 3600 * 1000)) });
  }
  const jd = ss.getSheetByName(SHEETS.JADWAL_DOKTER);
  const demand = {};
  if (jd && jd.getLastRow() > 1) {
    jd.getRange(2, 1, jd.getLastRow() - 1, 7).getValues().forEach(function (r) {
      if (!r[0]) return;
      const t = String(r[0]), shift = String(r[2]), spes = String(r[5]);
      const n = parseInt(r[6], 10) || 1;
      if (!demand[t]) demand[t] = { pagi: 0, siang: 0, spes: {}, spesPagi: {}, spesSiang: {}, duPagi: 0, duSiang: 0 };
      const dd = demand[t];
      if (shift === 'Pagi') { dd.pagi += n; dd.duPagi++; dd.spesPagi[spes] = (dd.spesPagi[spes] || 0) + n; }
      else { dd.siang += n; dd.duSiang++; dd.spesSiang[spes] = (dd.spesSiang[spes] || 0) + n; }
      dd.spes[spes] = Math.max(dd.spesPagi[spes] || 0, dd.spesSiang[spes] || 0);
    });
  }
  const perawat = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) { return { nama: p.nama, level: p.level, kapab: String(p.kapabilitas || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; }) }; });
  const receptionist = getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; })
    .map(function (r) { return { nama: r.nama, level: r.level }; });
  const req = bacaRequestCuti_();
  const perawatNama = {}; perawat.forEach(function (p) { perawatNama[p.nama] = true; });
  const rcpNama = {}; receptionist.forEach(function (r) { rcpNama[r.nama] = true; });
  const wajibPerawat = {}, wajibRcp = {};
  Object.keys(req.staf).forEach(function (nm) { if (perawatNama[nm]) wajibPerawat[nm] = req.staf[nm]; else if (rcpNama[nm]) wajibRcp[nm] = req.staf[nm]; });

  const ambang = Number(getConfig('AMBANG_DU')) || 4;
  const hasil = susunLibur_({
    tanggalList: tanggalList, demand: demand, perawat: perawat, receptionist: receptionist,
    jatahPerawat: Number(getConfig('LIBUR_WAJIB_PERAWAT')) || 4,
    jatahReceptionist: Number(getConfig('LIBUR_WAJIB_RECEPTIONIST')) || 4,
    maxWeekendGold: Number(getConfig('MAX_LIBUR_WEEKEND_GOLD')) || 1,
    maxSlot: PL_MAX_SLOT, ambangDU: ambang,
    minKecil: Number(getConfig('MIN_RECEPTIONIST_DU_KECIL')) || 2,
    minBesar: Number(getConfig('MIN_RECEPTIONIST_DU_BESAR')) || 3,
    wajibPerawat: wajibPerawat, wajibRcp: wajibRcp
  });
  buatPapanLiburTanpaKonfirmasi_(ss, tanggalList);
  const pl = ss.getSheetByName('Papan_Libur');
  tanggalList.forEach(function (t, i) {
    const names = hasil.liburMap[t.tanggal] || [];
    if (names.length) pl.getRange(i + 2, 3, 1, Math.min(names.length, PL_MAX_SLOT)).setValues([names.slice(0, PL_MAX_SLOT)]);
  });
  return hasil;
}

/** Core buat + auto-isi Papan Asistensi (UI-free) */
function generateAsistensiCore_(ss) {
  const jd = ss.getSheetByName(SHEETS.JADWAL_DOKTER);
  const data = jd.getRange(2, 1, jd.getLastRow() - 1, 7).getValues().filter(function (r) { return r[0]; });
  const rows = [];
  data.forEach(function (r) {
    const n = parseInt(r[6], 10) || 1;
    for (let s = 1; s <= n; s++) rows.push([r[0], r[1], r[2], r[4], r[5], s, '']);
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
  const namaPerawat = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; }).map(function (p) { return p.nama; });
  if (namaPerawat.length && rows.length) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(namaPerawat, true).setAllowInvalid(true).build();
    pa.getRange(2, PA_KOL_PERAWAT, rows.length).setDataValidation(rule);
  }
  pa.setFrozenRows(1);

  // auto-isi
  const slots = rows.map(function (r, i) { return { idx: i, tanggal: String(r[0]), hari: String(r[1]), shift: String(r[2]), dokter: String(r[3]), spes: String(r[4]) }; });
  const perawat = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) { return { nama: p.nama, level: p.level, kategori: p.kategori, kapab: String(p.kapabilitas || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; }) }; });
  const hasil = susunAsistensi_({ slots: slots, perawat: perawat, liburMap: petaLiburPerNama_(ss), longshiftTarget: Number(getConfig('LONGSHIFT_WAJIB_PERAWAT')) || 4, prefMap: bacaPreferensiDokter_(), jagaMap: bacaRequestJaga_().map });
  if (rows.length) {
    const nilai = hasil.assign.map(function (nm) { return [nm || '']; });
    const bg = hasil.assign.map(function (nm) { return [nm ? WARNA_VALID : WARNA_PELANGGARAN]; });
    pa.getRange(2, PA_KOL_PERAWAT, rows.length, 1).setValues(nilai).setBackgrounds(bg);
  }
  return { terisi: hasil.terisi, total: slots.length };
}

/** Core buat + auto-isi Papan Resepsionis (UI-free) */
function generateResepsionisCore_(ss) {
  const jd = ss.getSheetByName(SHEETS.JADWAL_DOKTER);
  const du = {};
  jd.getRange(2, 1, jd.getLastRow() - 1, 7).getValues().forEach(function (r) {
    if (!r[0]) return;
    const t = String(r[0]);
    if (!du[t]) du[t] = { Pagi: 0, Siang: 0, hari: String(r[1]) };
    du[t][String(r[2])]++;
  });
  const ambang = Number(getConfig('AMBANG_DU')) || 4;
  const minKecil = Number(getConfig('MIN_RECEPTIONIST_DU_KECIL')) || 2;
  const minBesar = Number(getConfig('MIN_RECEPTIONIST_DU_BESAR')) || 3;
  const minDari = function (n) { if (n <= 0) return 0; return n <= ambang ? minKecil : minBesar; };
  const tanggalUrut = Object.keys(du).sort(function (a, b) { return kunciTgl_(a) < kunciTgl_(b) ? -1 : 1; });
  const receptionist = getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; }).map(function (r) { return { nama: r.nama, level: r.level }; });
  const hasil = susunReceptionist_({ tanggalUrut: tanggalUrut, du: du, minDari: minDari, receptionist: receptionist, liburMap: petaLiburPerNama_(ss), longshiftTarget: Number(getConfig('LONGSHIFT_WAJIB_RECEPT_SILVER')) || 4, jagaMap: bacaRequestJaga_().map });
  let pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (!pr) pr = ss.insertSheet(SHEETS.PAPAN_RECEPTIONIST);
  pr.clear();
  pr.getRange(1, 1, pr.getMaxRows(), pr.getMaxColumns()).clearDataValidations(); // buang dropdown lama
  pr.getRange(1, 1, 1, PR_HEADER.length).setValues([PR_HEADER]);
  formatHeader_(pr, PR_HEADER.length);
  if (hasil.rows.length) {
    // pasang validasi BARU dulu, baru tulis nama (agar tak bentrok aturan lama)
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(receptionist.map(function (r) { return r.nama; }), true).setAllowInvalid(true).build();
    pr.getRange(2, PR_KOL_SLOT1, hasil.rows.length, PR_MAX_SLOT).setDataValidation(rule);
    pr.getRange(2, 1, hasil.rows.length, 1).setNumberFormat('@'); // kolom Tanggal jadi TEKS
    pr.getRange(2, 1, hasil.rows.length, PR_HEADER.length).setValues(hasil.rows);
    hasil.baruKurang.forEach(function (idx) { pr.getRange(idx + 2, PR_KOL_SLOT1, 1, PR_MAX_SLOT).setBackground(WARNA_PELANGGARAN); });
  }
  pr.setFrozenRows(1);
  return { baris: hasil.rows.length, kurang: hasil.baruKurang.length };
}

/* ================= BOARD (READ) ================= */

/** Normalisasi kolom tanggal (Date/teks) → {tanggal:"dd/MM/yyyy", hari:"Selasa"} */
function boardTglHari_(v0, v1) {
  const tgl = fmtTglCell_(v0);
  const d = parseTgl_(tgl);
  let hari = String(v1);
  if (d) { const n = NAMA_HARI[d.getDay()]; hari = n.charAt(0) + n.slice(1).toLowerCase(); }
  return { tanggal: tgl, hari: hari };
}

function getBoardAsistensi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (!pa || pa.getLastRow() < 2) return { rows: [] };
  const vals = pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues();
  const rows = vals.map(function (r, i) {
    const th = boardTglHari_(r[0], r[1]);
    return { baris: i + 2, tanggal: th.tanggal, hari: th.hari, shift: String(r[2]),
      dokter: String(r[3]), spes: String(r[4]), jumlah: r[5], perawat: String(r[6]) };
  });
  return { rows: rows };
}

function getBoardResepsionis() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (!pr || pr.getLastRow() < 2) return { rows: [] };
  const vals = pr.getRange(2, 1, pr.getLastRow() - 1, PR_KOL_SLOT1 + PR_MAX_SLOT - 1).getValues();
  const rows = vals.map(function (r) {
    const slot = [];
    for (let i = 0; i < PR_MAX_SLOT; i++) { const v = String(r[PR_KOL_SLOT1 - 1 + i] || '').trim(); if (v) slot.push(v); }
    const th = boardTglHari_(r[0], r[1]);
    return { tanggal: th.tanggal, hari: th.hari, shift: String(r[2]), du: r[3], minimal: r[4], isi: slot };
  });
  return { rows: rows };
}

function getBoardLibur() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pl = ss.getSheetByName('Papan_Libur');
  if (!pl || pl.getLastRow() < 2) return { rows: [] };
  const vals = pl.getRange(2, 1, pl.getLastRow() - 1, 2 + PL_MAX_SLOT).getValues();
  const rows = vals.map(function (r) {
    const names = [];
    for (let i = 2; i < 2 + PL_MAX_SLOT; i++) { const v = String(r[i] || '').trim(); if (v) names.push(v); }
    const th = boardTglHari_(r[0], r[1]);
    return { tanggal: th.tanggal, hari: th.hari, names: names };
  });
  return { rows: rows };
}

/* ================= KALENDER GABUNGAN ================= */

/** Agregasi asistensi + resepsionis + libur per tanggal untuk tampilan kalender */
function getKalenderData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pMulai = parseTgl_(getConfig('PERIODE_MULAI'));
  const pAkhir = parseTgl_(getConfig('PERIODE_AKHIR'));
  if (!pMulai || !pAkhir) return { ok: false, pesan: 'Periode belum diset.' };

  const weekend = String(getConfig('HARI_WEEKEND') || 'Sabtu,Minggu').split(',').map(function (h) { return h.trim().toLowerCase(); });
  const dates = [], map = {};
  for (let d = new Date(pMulai); d <= pAkhir; d.setDate(d.getDate() + 1)) {
    const t = formatTgl_(d);
    const n = NAMA_HARI[d.getDay()];
    const hari = n.charAt(0) + n.slice(1).toLowerCase();
    map[t] = { tanggal: t, hari: hari, isWeekend: weekend.indexOf(hari.toLowerCase()) !== -1, asis: [], resep: [], libur: [] };
    dates.push(t);
  }

  // Peta perawat untuk flag kategori (New/Existing) & level
  const prwMap = {};
  getMasterData('perawat').forEach(function (p) { prwMap[p.nama.toLowerCase()] = { kategori: p.kategori, level: p.level }; });
  // Peta seluruh staf → peran (untuk penanda pada daftar libur)
  const stafMap = {};
  getMasterData('perawat').forEach(function (p) { if (p.aktif !== 'Tidak') stafMap[p.nama.toLowerCase()] = { nama: p.nama, peran: 'Perawat', level: p.level, kategori: p.kategori, email: String(p.email || '').trim().toLowerCase() }; });
  getMasterData('receptionist').forEach(function (r) { if (r.aktif !== 'Tidak') stafMap[r.nama.toLowerCase()] = { nama: r.nama, peran: 'Resepsionis', level: r.level, kategori: '', email: String(r.email || '').trim().toLowerCase() }; });

  // Akumulasi statistik per staf (total tugas, longshift, libur) untuk tooltip
  const stafStat = {};
  const ss_ = function (nm) { const k = nm.toLowerCase(); if (!stafStat[k]) stafStat[k] = { total: 0, pagi: {}, siang: {}, libur: 0 }; return stafStat[k]; };

  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (pa && pa.getLastRow() > 1) {
    pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues().forEach(function (r, i) {
      const t = fmtTglCell_(r[0]);
      const nm = String(r[6]);
      const info = prwMap[nm.toLowerCase()] || {};
      if (nm.trim()) { const st = ss_(nm); st.total++; if (String(r[2]) === 'Pagi') st.pagi[t] = 1; else st.siang[t] = 1; }
      if (map[t]) map[t].asis.push({ baris: i + 2, shift: String(r[2]), dokter: String(r[3]), spes: String(r[4]), jumlah: r[5], perawat: nm, perawatKat: info.kategori || '', perawatLevel: info.level || '' });
    });
  }
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (pr && pr.getLastRow() > 1) {
    pr.getRange(2, 1, pr.getLastRow() - 1, PR_KOL_SLOT1 + PR_MAX_SLOT - 1).getValues().forEach(function (r, i) {
      const t = fmtTglCell_(r[0]); if (!map[t]) return;
      const shift = String(r[2]);
      const isi = [], slots = [];
      for (let c = PR_KOL_SLOT1 - 1; c < PR_KOL_SLOT1 - 1 + PR_MAX_SLOT; c++) { const v = String(r[c] || '').trim(); slots.push({ c: c + 1, nama: v }); if (v) isi.push(v); }
      isi.forEach(function (nm) { const st = ss_(nm); st.total++; if (shift === 'Pagi') st.pagi[t] = 1; else st.siang[t] = 1; });
      map[t].resep.push({ shift: shift, du: r[3], minimal: r[4], baris: i + 2, isi: isi, slots: slots });
    });
  }
  const pl = ss.getSheetByName('Papan_Libur');
  if (pl && pl.getLastRow() > 1) {
    pl.getRange(2, 1, pl.getLastRow() - 1, 2 + PL_MAX_SLOT).getValues().forEach(function (r) {
      const t = fmtTglCell_(r[0]); if (!map[t]) return;
      const names = []; for (let c = 2; c < 2 + PL_MAX_SLOT; c++) { const v = String(r[c] || '').trim(); if (v) names.push(v); }
      names.forEach(function (nm) { ss_(nm).libur++; });
      map[t].libur = names;
    });
  }
  // Gabungkan statistik (total tugas, longshift, libur) ke stafMap untuk tooltip
  Object.keys(stafStat).forEach(function (k) {
    const st = stafStat[k];
    let ls = 0; Object.keys(st.pagi).forEach(function (t) { if (st.siang[t]) ls++; });
    if (stafMap[k]) { stafMap[k].total = st.total; stafMap[k].longshift = ls; stafMap[k].libur = st.libur; }
    else stafMap[k] = { nama: k, peran: '', total: st.total, longshift: ls, libur: st.libur };
  });
  // Info blocker terkini → ditampilkan di Mapping Staff agar bisa diperbaiki manual.
  let blocker = 0, blockerRinci = [];
  try { const h = auditInti_(); if (h && h.ok) { const b = hitungBlocker_(h); blocker = b.total; blockerRinci = b.rinci; } } catch (e) {}
  return { ok: true, mulai: formatTgl_(pMulai), akhir: formatTgl_(pAkhir), dates: dates, map: map, staf: stafMap, blocker: blocker, blockerRinci: blockerRinci };
}

/* ================= EDIT LIBUR (IN-APP) ================= */

/** Apakah `nama` sudah bertugas (asistensi/resepsionis) pada `tgl`? Kembalikan '' | 'asistensi' | 'resepsionis'. */
function bertugasTanggal_(ss, nama, tgl) {
  const low = String(nama).toLowerCase();
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (pa && pa.getLastRow() > 1) {
    const rows = pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (fmtTglCell_(rows[i][0]) === tgl && String(rows[i][PA_KOL_PERAWAT - 1]).trim().toLowerCase() === low) return 'asistensi';
    }
  }
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (pr && pr.getLastRow() > 1) {
    const rows = pr.getRange(2, 1, pr.getLastRow() - 1, PR_KOL_SLOT1 + PR_MAX_SLOT - 1).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (fmtTglCell_(rows[i][0]) !== tgl) continue;
      for (let c = PR_KOL_SLOT1 - 1; c < PR_KOL_SLOT1 - 1 + PR_MAX_SLOT; c++) {
        if (String(rows[i][c] || '').trim().toLowerCase() === low) return 'resepsionis';
      }
    }
  }
  return '';
}

/** Baris Papan_Libur untuk `tgl` (-1 bila tak ada). */
function cariBarisLibur_(pl, tgl) {
  if (!pl || pl.getLastRow() < 2) return -1;
  const rows = pl.getRange(2, 1, pl.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < rows.length; i++) { if (fmtTglCell_(rows[i][0]) === tgl) return i + 2; }
  return -1;
}

/** Segarkan dropdown nama Papan_Libur dgn staf aktif terkini (allowInvalid=true)
 *  supaya penulisan nama staf baru tak ditolak data validation lama. */
function refreshLiburValidation_(pl) {
  if (!pl || pl.getLastRow() < 2) return;
  const namaStaf = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; }).map(function (p) { return p.nama; })
    .concat(getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; }).map(function (r) { return r.nama; }));
  if (!namaStaf.length) { pl.getRange(2, 3, pl.getLastRow() - 1, PL_MAX_SLOT).clearDataValidations(); return; }
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(namaStaf, true).setAllowInvalid(true).build();
  pl.getRange(2, 3, pl.getLastRow() - 1, PL_MAX_SLOT).setDataValidation(rule);
  SpreadsheetApp.flush();
}

/** Tambah libur seharian utk 1 orang pd tgl tsb (Admin/SPV), dgn validasi anti-blocker. */
function tambahLiburWeb(tgl, nama) {
  wajibPeran_(['Admin', 'SPV']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nm = String(nama || '').trim();
  tgl = String(tgl || '').trim();
  if (!nm || !tgl) return { ok: false, pesan: 'Nama & tanggal wajib.' };
  const prw = getMasterData('perawat').filter(function (p) { return p.nama.toLowerCase() === nm.toLowerCase() && p.aktif !== 'Tidak'; })[0];
  const rcp = getMasterData('receptionist').filter(function (r) { return r.nama.toLowerCase() === nm.toLowerCase() && r.aktif !== 'Tidak'; })[0];
  if (!prw && !rcp) return { ok: false, pesan: nm + ' bukan perawat/resepsionis aktif.' };
  const pl = ss.getSheetByName('Papan_Libur');
  if (!pl) return { ok: false, pesan: 'Papan_Libur belum ada — generate jadwal dulu.' };
  const baris = cariBarisLibur_(pl, tgl);
  if (baris === -1) return { ok: false, pesan: 'Tanggal ' + tgl + ' tak ada di Papan_Libur.' };
  // N-03: perawat Silver tak boleh libur weekend (resepsionis boleh)
  const hari = String(pl.getRange(baris, 2).getValue()).toLowerCase();
  const weekend = String(getConfig('HARI_WEEKEND') || 'Sabtu,Minggu').split(',').map(function (h) { return h.trim().toLowerCase(); });
  if (prw && !rcp && prw.level !== 'Gold' && weekend.indexOf(hari) !== -1) {
    return { ok: false, pesan: 'Perawat Silver tak boleh libur weekend (N-03). Pilih hari kerja.' };
  }
  // Cari slot kosong & cegah duplikat
  const slotVals = pl.getRange(baris, 3, 1, PL_MAX_SLOT).getValues()[0];
  let kolomKosong = -1;
  for (let i = 0; i < PL_MAX_SLOT; i++) {
    const v = String(slotVals[i] || '').trim();
    if (v.toLowerCase() === nm.toLowerCase()) return { ok: false, pesan: nm + ' sudah libur tanggal itu.' };
    if (!v && kolomKosong === -1) kolomKosong = i;
  }
  if (kolomKosong === -1) return { ok: false, pesan: 'Slot libur tanggal itu penuh (maks ' + PL_MAX_SLOT + ').' };
  // Cegah bentrok: tak boleh libur bila masih bertugas hari itu
  const tugas = bertugasTanggal_(ss, nm, tgl);
  if (tugas) return { ok: false, pesan: nm + ' masih bertugas ' + tugas + ' tanggal itu — hapus penugasannya dulu, baru beri libur.' };
  refreshLiburValidation_(pl); // segarkan daftar validasi agar nama staf baru tak ditolak
  pl.getRange(baris, 3 + kolomKosong).setNumberFormat('@').setValue(nm);
  return { ok: true, pesan: 'Libur ' + nm + ' pada ' + tgl + ' ditambahkan.' };
}

/** Hapus libur 1 orang pd tgl tsb (Admin/SPV). */
function hapusLiburWeb(tgl, nama) {
  wajibPeran_(['Admin', 'SPV']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const nm = String(nama || '').trim();
  tgl = String(tgl || '').trim();
  const pl = ss.getSheetByName('Papan_Libur');
  const baris = cariBarisLibur_(pl, tgl);
  if (baris === -1) return { ok: false, pesan: 'Tanggal tak ditemukan.' };
  const slotVals = pl.getRange(baris, 3, 1, PL_MAX_SLOT).getValues()[0];
  for (let i = 0; i < PL_MAX_SLOT; i++) {
    if (String(slotVals[i] || '').trim().toLowerCase() === nm.toLowerCase()) {
      pl.getRange(baris, 3 + i).clearContent();
      return { ok: true, pesan: 'Libur ' + nm + ' pada ' + tgl + ' dihapus.' };
    }
  }
  return { ok: false, pesan: nm + ' tidak tercatat libur tanggal itu.' };
}

/* ================= EDIT ASISTENSI (IN-APP) ================= */

/** Daftar perawat eligible untuk 1 slot (kapabilitas + tidak libur + tidak dobel) */
function getEligiblePerawatWeb(baris) {
  wajibPeran_(['Admin', 'SPV']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  const info = pa.getRange(baris, 1, 1, PA_KOL_PERAWAT).getValues()[0];
  const tanggal = String(info[0]), shift = String(info[2]), dokter = String(info[3]), spes = String(info[4]);
  const perawatMap = petaPerawat_();
  const liburMap = petaLiburPerNama_(ss);
  // Nama yang sudah bertugas tanggal+shift ini + deteksi apakah dokter ini sudah punya perawat NEW (A-02)
  const dipakai = {};
  const totalMap = {}; // nama lower → total asistensi periode ini (semua tanggal/shift)
  let adaNewLain = false;
  pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues().forEach(function (r, i) {
    const nmAll = String(r[6]).trim();
    if (nmAll) totalMap[nmAll.toLowerCase()] = (totalMap[nmAll.toLowerCase()] || 0) + 1;
    if (i + 2 === baris) return;
    if (String(r[0]) === tanggal && String(r[2]) === shift) {
      const nm = nmAll;
      if (nm) {
        dipakai[nm.toLowerCase()] = true;
        // A-02: cek rekan New di dokter yang sama
        if (String(r[3]) === dokter) { const rk = perawatMap[nm.toLowerCase()]; if (rk && rk.kategori === 'New') adaNewLain = true; }
      }
    }
  });
  const out = [];
  Object.keys(perawatMap).forEach(function (k) {
    const p = perawatMap[k];
    if (p.aktif === 'Tidak') return;
    if (liburMap[p.nama] && liburMap[p.nama][tanggal]) return;
    if (dipakai[k]) return;
    const eligibleKapab = !p.kapabilitasSet.length || p.kapabilitasSet.indexOf(spes) !== -1;
    out.push({ nama: p.nama, level: p.level, kategori: p.kategori, kapabelSpes: eligibleKapab, total: totalMap[k] || 0 });
  });
  // Urut: pelanggaran paling SEDIKIT di atas (paling eligible), lalu asistensi
  // paling SEDIKIT di atas; yang paling banyak melanggar aturan di bawah.
  //   pelanggaran = luar kapabilitas (N-05) + (New saat dokter sudah punya New → A-02)
  const skorLanggar = function (x) {
    return (x.kapabelSpes ? 0 : 1) + ((adaNewLain && x.kategori === 'New') ? 1 : 0);
  };
  out.sort(function (a, b) {
    const sa = skorLanggar(a), sb = skorLanggar(b);
    if (sa !== sb) return sa - sb;                                    // makin patuh makin atas
    if ((a.total || 0) !== (b.total || 0)) return (a.total || 0) - (b.total || 0); // asistensi paling sedikit di atas
    return a.nama < b.nama ? -1 : 1;
  });
  return { spes: spes, dokter: dokter, adaNewLain: adaNewLain, list: out };
}

/** Daftar resepsionis eligible untuk slot papan resepsionis pada `baris`.
 *  Kecualikan yang libur & yang sudah di shift itu (akan bikin blocker);
 *  urut beban paling SEDIKIT di atas. */
function getEligibleResepsionisWeb(baris) {
  wajibPeran_(['Admin', 'SPV']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  const nCol = PR_KOL_SLOT1 + PR_MAX_SLOT - 1;
  const info = pr.getRange(baris, 1, 1, nCol).getValues()[0];
  const tanggal = String(info[0]), shift = String(info[2]);
  const liburMap = petaLiburPerNama_(ss);
  const total = {}, diShift = {};
  pr.getRange(2, 1, pr.getLastRow() - 1, nCol).getValues().forEach(function (r) {
    const t = String(r[0]), sh = String(r[2]);
    for (let c = PR_KOL_SLOT1 - 1; c < nCol; c++) {
      const nm = String(r[c] || '').trim(); if (!nm) continue; const low = nm.toLowerCase();
      total[low] = (total[low] || 0) + 1;
      if (t === tanggal && sh === shift) diShift[low] = true;
    }
  });
  const out = [];
  getMasterData('receptionist').forEach(function (r) {
    if (r.aktif === 'Tidak') return;
    const low = r.nama.toLowerCase();
    if (liburMap[r.nama] && liburMap[r.nama][tanggal]) return;   // libur → dikecualikan
    if (diShift[low]) return;                                     // sudah di shift ini → dikecualikan
    out.push({ nama: r.nama, level: r.level, total: total[low] || 0 });
  });
  out.sort(function (a, b) { return (a.total !== b.total) ? a.total - b.total : (a.nama < b.nama ? -1 : 1); });
  return { tanggal: tanggal, shift: shift, list: out };
}

/** Tulis nama resepsionis ke slot (baris,kolom). nama '' = kosongkan.
 *  Validasi: aktif, tidak libur, tidak dobel di shift sama. */
function setResepSlotWeb(baris, kolom, nama) {
  wajibPeran_(['Admin', 'SPV']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  refreshResepsionisValidation_(pr);
  const nCol = PR_KOL_SLOT1 + PR_MAX_SLOT - 1;
  const info = pr.getRange(baris, 1, 1, nCol).getValues()[0];
  const tanggal = String(info[0]), shift = String(info[2]);
  nama = String(nama || '').trim();
  if (!nama) { pr.getRange(baris, kolom).clearContent(); SpreadsheetApp.flush(); return { ok: true, pesan: 'Slot dikosongkan.' }; }
  const rc = getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak' && r.nama.toLowerCase() === nama.toLowerCase(); });
  if (!rc.length) return { ok: false, pesan: nama + ' bukan resepsionis aktif.' };
  const proper = rc[0].nama, low = proper.toLowerCase();
  const liburMap = petaLiburPerNama_(ss);
  if (liburMap[proper] && liburMap[proper][tanggal]) return { ok: false, pesan: proper + ' sedang libur tanggal ' + tanggal + '.' };
  const grid = pr.getRange(2, 1, pr.getLastRow() - 1, nCol).getValues();
  for (let i = 0; i < grid.length; i++) {
    const r = grid[i]; if (String(r[0]) !== tanggal || String(r[2]) !== shift) continue;
    for (let c = PR_KOL_SLOT1 - 1; c < nCol; c++) {
      if ((i + 2) === baris && (c + 1) === kolom) continue;
      if (String(r[c] || '').trim().toLowerCase() === low) return { ok: false, pesan: proper + ' sudah bertugas di shift ini.' };
    }
  }
  pr.getRange(baris, kolom).setNumberFormat('@').setValue(proper);
  SpreadsheetApp.flush();
  return { ok: true, pesan: proper + ' ditempatkan.' };
}

/**
 * Sarankan asistensi untuk SATU tanggal: isi hanya slot yang masih kosong,
 * seadil mungkin (beban terkecil didahulukan) & patuh aturan:
 * tidak libur, tidak dobel di shift sama, kapabilitas (N-05), A-02 (New tak berdampingan).
 * Assignment yang sudah ada TIDAK diubah.
 */
function suggestAsistensiHariWeb(tanggal) {
  wajibPeran_(['Admin', 'SPV']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (!pa || pa.getLastRow() < 2) return { ok: false, pesan: 'Papan asistensi kosong.' };
  const grid = pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues();
  const liburMap = petaLiburPerNama_(ss);
  const perawat = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) { return { nama: p.nama, level: p.level, kategori: p.kategori, kapab: String(p.kapabilitas || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; }) }; });
  if (!perawat.length) return { ok: false, pesan: 'Master perawat kosong.' };
  const pmap = {}; perawat.forEach(function (p) { pmap[p.nama.toLowerCase()] = p; });

  // Statistik saat ini: total beban (fairness), okupansi shift, jumlah New per dokter+shift (A-02)
  const total = {}; perawat.forEach(function (p) { total[p.nama.toLowerCase()] = 0; });
  const sudahShift = {}; const grupNew = {};
  grid.forEach(function (r) {
    const nm = String(r[6]).trim(); if (!nm) return; const low = nm.toLowerCase();
    total[low] = (total[low] || 0) + 1;
    sudahShift[low + '|' + String(r[0]) + '|' + String(r[2])] = true;
    const p = pmap[low];
    if (p && p.kategori === 'New') { const gk = String(r[0]) + '|' + String(r[2]) + '|' + String(r[3]); grupNew[gk] = (grupNew[gk] || 0) + 1; }
  });

  // Slot kosong tanggal ini, Pagi dulu
  const idxs = [];
  grid.forEach(function (r, i) { if (String(r[0]) === tanggal && !String(r[6]).trim()) idxs.push(i); });
  idxs.sort(function (a, b) { const sa = String(grid[a][2]), sb = String(grid[b][2]); if (sa !== sb) return sa === 'Pagi' ? -1 : 1; return a - b; });
  if (!idxs.length) return { ok: true, terisi: 0, sisa: 0, pesan: 'Tidak ada slot kosong pada ' + tanggal + '.' };

  let terisi = 0;
  idxs.forEach(function (i) {
    const r = grid[i]; const shift = String(r[2]), dokter = String(r[3]), spes = String(r[4]);
    const gk = tanggal + '|' + shift + '|' + dokter;
    const kand = perawat.filter(function (p) {
      const low = p.nama.toLowerCase();
      if (liburMap[p.nama] && liburMap[p.nama][tanggal]) return false;   // tidak libur
      if (sudahShift[low + '|' + tanggal + '|' + shift]) return false;   // tidak dobel
      if (p.kapab.length && p.kapab.indexOf(spes) === -1) return false;  // kapabilitas N-05
      if (p.kategori === 'New' && (grupNew[gk] || 0) > 0) return false;  // A-02
      return true;
    });
    if (!kand.length) return;
    // Paling adil: beban terkecil dulu, acak untuk seri
    kand.sort(function (a, b) { const ta = total[a.nama.toLowerCase()] || 0, tb = total[b.nama.toLowerCase()] || 0; return ta !== tb ? ta - tb : (Math.random() - 0.5); });
    const pil = kand[0]; const low = pil.nama.toLowerCase();
    grid[i][6] = pil.nama;
    total[low] = (total[low] || 0) + 1;
    sudahShift[low + '|' + tanggal + '|' + shift] = true;
    if (pil.kategori === 'New') grupNew[gk] = (grupNew[gk] || 0) + 1;
    pa.getRange(i + 2, PA_KOL_PERAWAT).setValue(pil.nama).setBackground(WARNA_VALID);
    terisi++;
  });
  const sisa = idxs.length - terisi;
  return { ok: true, terisi: terisi, sisa: sisa,
    pesan: terisi + ' slot terisi seadil mungkin' + (sisa ? ', ' + sisa + ' tetap kosong (tak ada perawat memenuhi syarat)' : '') + '.' };
}

/* ================= PERBAIKI BLOCKER (non-destruktif) ================= */

/** Isi SEMUA slot asisten kosong (semua tanggal) secara adil & patuh aturan.
 *  Non-destruktif: hanya mengisi sel kosong, tak mengubah yang sudah terisi. */
function isiAsistensiKosong_(ss) {
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (!pa || pa.getLastRow() < 2) return { terisi: 0, sisa: 0 };
  const grid = pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues();
  const liburMap = petaLiburPerNama_(ss);
  const perawat = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) { return { nama: p.nama, level: p.level, kategori: p.kategori, kapab: String(p.kapabilitas || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; }) }; });
  if (!perawat.length) return { terisi: 0, sisa: 0 };
  const pmap = {}; perawat.forEach(function (p) { pmap[p.nama.toLowerCase()] = p; });
  const total = {}; perawat.forEach(function (p) { total[p.nama.toLowerCase()] = 0; });
  const sudahShift = {}, grupNew = {};
  grid.forEach(function (r) {
    const nm = String(r[6]).trim(); if (!nm) return; const low = nm.toLowerCase();
    total[low] = (total[low] || 0) + 1;
    sudahShift[low + '|' + String(r[0]) + '|' + String(r[2])] = true;
    const p = pmap[low]; if (p && p.kategori === 'New') { const gk = String(r[0]) + '|' + String(r[2]) + '|' + String(r[3]); grupNew[gk] = (grupNew[gk] || 0) + 1; }
  });
  const idxs = []; grid.forEach(function (r, i) { if (!String(r[6]).trim()) idxs.push(i); });
  idxs.sort(function (a, b) { const ta = String(grid[a][0]), tb = String(grid[b][0]); if (ta !== tb) return ta < tb ? -1 : 1; const sa = String(grid[a][2]), sb = String(grid[b][2]); if (sa !== sb) return sa === 'Pagi' ? -1 : 1; return a - b; });
  let terisi = 0; const detail = [];
  idxs.forEach(function (i) {
    const r = grid[i]; const tanggal = String(r[0]), shift = String(r[2]), dokter = String(r[3]), spes = String(r[4]);
    const gk = tanggal + '|' + shift + '|' + dokter;
    const kand = perawat.filter(function (p) {
      const low = p.nama.toLowerCase();
      if (liburMap[p.nama] && liburMap[p.nama][tanggal]) return false;   // tidak libur
      if (sudahShift[low + '|' + tanggal + '|' + shift]) return false;   // tidak dobel
      if (p.kapab.length && p.kapab.indexOf(spes) === -1) return false;  // kapabilitas N-05
      if (p.kategori === 'New' && (grupNew[gk] || 0) > 0) return false;  // A-02
      return true;
    });
    if (!kand.length) return;
    kand.sort(function (a, b) { const ta = total[a.nama.toLowerCase()] || 0, tb = total[b.nama.toLowerCase()] || 0; return ta !== tb ? ta - tb : (Math.random() - 0.5); });
    const pil = kand[0]; const low = pil.nama.toLowerCase();
    grid[i][6] = pil.nama; total[low] = (total[low] || 0) + 1; sudahShift[low + '|' + tanggal + '|' + shift] = true;
    if (pil.kategori === 'New') grupNew[gk] = (grupNew[gk] || 0) + 1;
    pa.getRange(i + 2, PA_KOL_PERAWAT).setValue(pil.nama).setBackground(WARNA_VALID);
    terisi++; detail.push({ nama: pil.nama, tgl: tanggal, shift: shift, dokter: dokter });
  });
  return { terisi: terisi, sisa: idxs.length - terisi, detail: detail };
}

/** Tambah resepsionis ke shift yang di bawah minimal (R-03/R-04). Non-destruktif. */
function isiResepsionisKurang_(ss) {
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (!pr || pr.getLastRow() < 2) return { terisi: 0, sisa: 0 };
  const nCol = PR_KOL_SLOT1 + PR_MAX_SLOT - 1;
  const grid = pr.getRange(2, 1, pr.getLastRow() - 1, nCol).getValues();
  const liburMap = petaLiburPerNama_(ss);
  const rcp = getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; }).map(function (r) { return { nama: r.nama, level: r.level }; });
  if (!rcp.length) return { terisi: 0, sisa: 0 };
  const total = {}; rcp.forEach(function (r) { total[r.nama.toLowerCase()] = 0; });
  const sudah = {};
  grid.forEach(function (r) { const t = String(r[0]), sh = String(r[2]); for (let c = PR_KOL_SLOT1 - 1; c < nCol; c++) { const nm = String(r[c] || '').trim(); if (nm) { const low = nm.toLowerCase(); total[low] = (total[low] || 0) + 1; sudah[low + '|' + t + '|' + sh] = true; } } });
  let terisi = 0, sisa = 0; const detail = [];
  grid.forEach(function (r, i) {
    const t = String(r[0]), sh = String(r[2]); const min = parseInt(r[4], 10) || 0;
    const isi = [], kosongCol = [];
    for (let c = PR_KOL_SLOT1 - 1; c < nCol; c++) { const nm = String(r[c] || '').trim(); if (nm) isi.push(nm); else kosongCol.push(c); }
    let need = min - isi.length;
    while (need > 0 && kosongCol.length) {
      const kand = rcp.filter(function (x) { const low = x.nama.toLowerCase(); if (liburMap[x.nama] && liburMap[x.nama][t]) return false; if (sudah[low + '|' + t + '|' + sh]) return false; return true; });
      if (!kand.length) { sisa += need; break; }
      kand.sort(function (a, b) { const ta = total[a.nama.toLowerCase()] || 0, tb = total[b.nama.toLowerCase()] || 0; return ta !== tb ? ta - tb : (Math.random() - 0.5); });
      const pil = kand[0]; const low = pil.nama.toLowerCase(); const col = kosongCol.shift();
      pr.getRange(i + 2, col + 1).setNumberFormat('@').setValue(pil.nama);
      sudah[low + '|' + t + '|' + sh] = true; total[low] = (total[low] || 0) + 1;
      need--; terisi++; detail.push({ nama: pil.nama, tgl: t, shift: sh });
    }
  });
  return { terisi: terisi, sisa: sisa, detail: detail };
}

/** Top-up libur yang belum terpenuhi (N-02 perawat, R-06 resepsionis) secara ADITIF:
 *  hanya menambah libur, tak menghapus yang ada. Patuh N-03 (Silver tak weekend,
 *  Gold maks 1x), tak bentrok tugas, dan kapasitas tetap cukup untuk asistensi/shift. */
function topUpLibur_(ss) {
  const jatahPrw = Number(getConfig('LIBUR_WAJIB_PERAWAT')) || 4;
  const jatahRcp = Number(getConfig('LIBUR_WAJIB_RECEPTIONIST')) || 4;
  const maxWkndGold = Number(getConfig('MAX_LIBUR_WEEKEND_GOLD')) || 1;
  const weekend = String(getConfig('HARI_WEEKEND') || 'Sabtu,Minggu').split(',').map(function (h) { return h.trim().toLowerCase(); });
  const pl = ss.getSheetByName('Papan_Libur');
  if (!pl || pl.getLastRow() < 2) return { tambah: 0, gagal: 0, detail: [] };
  refreshLiburValidation_(pl); // pastikan dropdown mengizinkan penulisan nama

  const perawat = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; });
  const recept = getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; });
  const prwSet = {}; perawat.forEach(function (p) { prwSet[p.nama.toLowerCase()] = true; });
  const rcpSet = {}; recept.forEach(function (r) { rcpSet[r.nama.toLowerCase()] = true; });
  const nPrw = perawat.length, nRcp = recept.length;

  // Papan_Libur → info per tanggal
  const rows = pl.getRange(2, 1, pl.getLastRow() - 1, 2 + PL_MAX_SLOT).getValues();
  const dInfo = {}, dOrder = [];
  rows.forEach(function (r, i) {
    const t = fmtTglCell_(r[0]); if (!t) return;
    const names = [], kosong = [];
    for (let c = 2; c < 2 + PL_MAX_SLOT; c++) { const v = String(r[c] || '').trim(); if (v) names.push(v.toLowerCase()); else kosong.push(c); }
    let liburPrw = 0, liburRcp = 0; names.forEach(function (n) { if (prwSet[n]) liburPrw++; else if (rcpSet[n]) liburRcp++; });
    dInfo[t] = { row: i + 2, isWeekend: weekend.indexOf(String(r[1]).toLowerCase()) !== -1, names: names, kosong: kosong, liburPrw: liburPrw, liburRcp: liburRcp };
    dOrder.push(t);
  });

  // Kebutuhan asistensi & tugas perawat per tanggal
  const asisNeed = {}, prwTugas = {};
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (pa && pa.getLastRow() > 1) {
    const ar = pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues();
    const pg = {}, sg = {};
    ar.forEach(function (r) { const t = fmtTglCell_(r[0]); const nm = String(r[6]).trim();
      if (String(r[2]) === 'Pagi') pg[t] = (pg[t] || 0) + 1; else sg[t] = (sg[t] || 0) + 1;
      if (nm) { if (!prwTugas[t]) prwTugas[t] = {}; prwTugas[t][nm.toLowerCase()] = 1; } });
    dOrder.forEach(function (t) { asisNeed[t] = Math.max(pg[t] || 0, sg[t] || 0); });
  }
  // Kebutuhan resepsionis & tugas per tanggal
  const rcpNeed = {}, rcpTugas = {};
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (pr && pr.getLastRow() > 1) {
    const rr = pr.getRange(2, 1, pr.getLastRow() - 1, PR_KOL_SLOT1 + PR_MAX_SLOT - 1).getValues();
    const mp = {}, msg = {};
    rr.forEach(function (r) { const t = fmtTglCell_(r[0]); const min = parseInt(r[4], 10) || 0;
      if (String(r[2]) === 'Pagi') mp[t] = Math.max(mp[t] || 0, min); else msg[t] = Math.max(msg[t] || 0, min);
      for (let c = PR_KOL_SLOT1 - 1; c < PR_KOL_SLOT1 - 1 + PR_MAX_SLOT; c++) { const nm = String(r[c] || '').trim(); if (nm) { if (!rcpTugas[t]) rcpTugas[t] = {}; rcpTugas[t][nm.toLowerCase()] = 1; } } });
    dOrder.forEach(function (t) { rcpNeed[t] = Math.max(mp[t] || 0, msg[t] || 0); });
  }

  // Hitung libur & weekend per staf
  const liburCnt = {}, wkndCnt = {};
  dOrder.forEach(function (t) { dInfo[t].names.forEach(function (low) { liburCnt[low] = (liburCnt[low] || 0) + 1; if (dInfo[t].isWeekend) wkndCnt[low] = (wkndCnt[low] || 0) + 1; }); });

  let tambah = 0, gagal = 0; const detail = [];
  const proses = function (staf, isPrw) {
    const low = staf.nama.toLowerCase();
    let cur = liburCnt[low] || 0;
    const jatah = isPrw ? jatahPrw : jatahRcp;
    // urutkan kandidat: tanggal dengan libur peran tsb paling sedikit (sebar), weekday dulu
    const cand = dOrder.slice().sort(function (a, b) {
      const la = isPrw ? dInfo[a].liburPrw : dInfo[a].liburRcp, lb = isPrw ? dInfo[b].liburPrw : dInfo[b].liburRcp;
      if (la !== lb) return la - lb; return (dInfo[a].isWeekend ? 1 : 0) - (dInfo[b].isWeekend ? 1 : 0);
    });
    for (let k = 0; k < cand.length && cur < jatah; k++) {
      const t = cand[k], di = dInfo[t];
      if (di.names.indexOf(low) !== -1) continue;         // sudah libur
      if (!di.kosong.length) continue;                     // slot penuh
      const tugas = isPrw ? (prwTugas[t] && prwTugas[t][low]) : (rcpTugas[t] && rcpTugas[t][low]);
      if (tugas) continue;                                 // bertugas hari itu
      if (isPrw && di.isWeekend) { if (staf.level !== 'Gold') continue; if ((wkndCnt[low] || 0) >= maxWkndGold) continue; }
      // kapasitas: setelah 1 lagi libur, sisa staf peran ini masih cukup utk kebutuhan
      const need = isPrw ? (asisNeed[t] || 0) : (rcpNeed[t] || 0);
      const nTot = isPrw ? nPrw : nRcp;
      const liburRole = isPrw ? di.liburPrw : di.liburRcp;
      if (nTot - liburRole - 1 < need) continue;
      // tulis libur
      const col = di.kosong.shift();
      pl.getRange(di.row, col + 1).setNumberFormat('@').setValue(staf.nama);
      di.names.push(low); if (isPrw) di.liburPrw++; else di.liburRcp++;
      liburCnt[low] = ++cur; if (di.isWeekend) wkndCnt[low] = (wkndCnt[low] || 0) + 1;
      tambah++; detail.push({ nama: staf.nama, tgl: t, peran: isPrw ? 'Perawat' : 'Resepsionis' });
    }
    if (cur < jatah) gagal += (jatah - cur);
  };
  perawat.forEach(function (p) { proses(p, true); });
  recept.forEach(function (r) { proses(r, false); });
  return { tambah: tambah, gagal: gagal, detail: detail };
}

/** Jadwalkan resepsionis aktif yang belum punya shift sama sekali (R-09):
 *  tambahkan ke sebuah shift yang masih punya slot kosong & ia tak libur. */
function jadwalkanResepsionisTakTerjadwal_(ss) {
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (!pr || pr.getLastRow() < 2) return { terjadwal: 0, gagal: 0, alasan: ['papan Jadwal_Resepsionis kosong — generate dulu'], detail: [] };
  const nCol = PR_KOL_SLOT1 + PR_MAX_SLOT - 1;
  const grid = pr.getRange(2, 1, pr.getLastRow() - 1, nCol).getValues();
  const liburMap = petaLiburPerNama_(ss);
  const recept = getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; });
  const total = {}; recept.forEach(function (r) { total[r.nama.toLowerCase()] = 0; });
  const inShift = {};
  grid.forEach(function (r) {
    const t = String(r[0]), sh = String(r[2]);
    for (let c = PR_KOL_SLOT1 - 1; c < nCol; c++) {
      const nm = String(r[c] || '').trim();
      if (nm) { const low = nm.toLowerCase(); if (total[low] !== undefined) total[low]++; inShift[low + '|' + t + '|' + sh] = true; }
    }
  });
  const belum = recept.filter(function (r) { return (total[r.nama.toLowerCase()] || 0) === 0; });
  let terjadwal = 0, gagal = 0; const alasan = [], detail = [];
  belum.forEach(function (r) {
    const low = r.nama.toLowerCase();
    let placed = false, adaHariLuang = false;
    for (let i = 0; i < grid.length && !placed; i++) {
      const row = grid[i], t = String(row[0]), sh = String(row[2]);
      if (liburMap[r.nama] && liburMap[r.nama][t]) continue;    // jangan tabrak libur
      if (inShift[low + '|' + t + '|' + sh]) continue;          // sudah di shift ini
      adaHariLuang = true;
      for (let c = PR_KOL_SLOT1 - 1; c < nCol; c++) {
        if (!String(row[c] || '').trim()) {
          pr.getRange(i + 2, c + 1).setNumberFormat('@').setValue(r.nama);
          grid[i][c] = r.nama;                                  // cegah tabrakan antar-orang belum terjadwal
          inShift[low + '|' + t + '|' + sh] = true; placed = true; terjadwal++;
          detail.push({ nama: r.nama, tgl: t, shift: sh }); break;
        }
      }
    }
    if (!placed) {
      gagal++;
      alasan.push(r.nama + (adaHariLuang
        ? ' — semua slot shift sudah penuh (maks ' + PR_MAX_SLOT + '/shift)'
        : ' — libur di semua hari papan'));
    }
  });
  return { terjadwal: terjadwal, gagal: gagal, alasan: alasan, detail: detail };
}

/** Segarkan dropdown papan RESEPSIONIS: semua resepsionis aktif + allowInvalid(true)
 *  agar penulisan nama (mis. staf baru) tak ditolak data-validation. */
function refreshResepsionisValidation_(pr) {
  if (!pr || pr.getLastRow() < 2) return;
  const rng = pr.getRange(2, PR_KOL_SLOT1, pr.getLastRow() - 1, PR_MAX_SLOT);
  const nama = getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; })
    .map(function (r) { return r.nama; });
  if (!nama.length) { rng.clearDataValidations(); return; }
  rng.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(nama, true).setAllowInvalid(true).build());
  SpreadsheetApp.flush();
}

/** Segarkan dropdown papan ASISTENSI (kolom Perawat): semua perawat aktif + allowInvalid(true). */
function refreshAsistensiValidation_(pa) {
  if (!pa || pa.getLastRow() < 2) return;
  const rng = pa.getRange(2, PA_KOL_PERAWAT, pa.getLastRow() - 1, 1);
  const nama = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) { return p.nama; });
  if (!nama.length) { rng.clearDataValidations(); return; }
  rng.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(nama, true).setAllowInvalid(true).build());
  SpreadsheetApp.flush();
}

/** Nama (lowercase) resepsionis AKTIF yang belum muncul di papan sama sekali. */
function namaTanpaJadwalResepsionis_(ss) {
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  const ada = {};
  if (pr && pr.getLastRow() >= 2) {
    const nCol = PR_KOL_SLOT1 + PR_MAX_SLOT - 1;
    pr.getRange(2, 1, pr.getLastRow() - 1, nCol).getValues().forEach(function (row) {
      for (let c = PR_KOL_SLOT1 - 1; c < nCol; c++) { const nm = String(row[c] || '').trim(); if (nm) ada[nm.toLowerCase()] = true; }
    });
  }
  return getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; })
    .map(function (r) { return r.nama.toLowerCase(); })
    .filter(function (l) { return !ada[l]; });
}

/** Nama (lowercase) perawat AKTIF yang belum muncul di papan asistensi sama sekali. */
function namaTanpaJadwalPerawat_(ss) {
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  const ada = {};
  if (pa && pa.getLastRow() >= 2) {
    pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues().forEach(function (row) {
      const nm = String(row[6] || '').trim(); if (nm) ada[nm.toLowerCase()] = true;
    });
  }
  const pMap = petaPerawat_();
  return Object.keys(pMap).filter(function (k) { return pMap[k].aktif !== 'Tidak' && !ada[k]; });
}

/** Seimbangkan beban RESEPSIONIS: pindahkan shift dari staf di ATAS rata-rata
 *  ke staf 'prioritas' (yang mulai TANPA jadwal), sampai mendekati rata-rata.
 *  Menukar nama dalam slot yang sama → jumlah resepsionis per shift TETAP
 *  (tak melanggar R-03/R-04). Hormati libur & tak dobel di shift yang sama.
 *  @param {Array<string>=} prioritas nama lowercase yang wajib dinaikkan (default: none). */
function seimbangkanBebanResepsionis_(ss, prioritas) {
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (!pr || pr.getLastRow() < 2) return { pindah: 0, sisa: [], detail: [] };
  const nCol = PR_KOL_SLOT1 + PR_MAX_SLOT - 1;
  const grid = pr.getRange(2, 1, pr.getLastRow() - 1, nCol).getValues();
  const liburMap = petaLiburPerNama_(ss);
  const recept = getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; });
  if (recept.length < 2) return { pindah: 0, sisa: [], detail: [] };
  const namaAsli = {}, cnt = {};
  recept.forEach(function (r) { const l = r.nama.toLowerCase(); namaAsli[l] = r.nama; cnt[l] = 0; });
  const slots = [], inShift = {};
  for (let i = 0; i < grid.length; i++) {
    const t = String(grid[i][0]), sh = String(grid[i][2]);
    for (let c = PR_KOL_SLOT1 - 1; c < nCol; c++) {
      const nm = String(grid[i][c] || '').trim(); if (!nm) continue;
      const low = nm.toLowerCase();
      slots.push({ i: i, c: c, low: low, t: t, sh: sh });
      if (cnt[low] !== undefined) cnt[low]++;
      inShift[low + '|' + t + '|' + sh] = true;
    }
  }
  const totalAssign = slots.filter(function (s) { return cnt[s.low] !== undefined; }).length;
  const target = Math.floor(totalAssign / recept.length);
  if (target < 1) return { pindah: 0, sisa: [], detail: [] };
  let pindah = 0; const sisa = [], detail = [];
  // penerima = HANYA staf prioritas (mulai tanpa jadwal) yang masih di bawah rata-rata
  const fokus = prioritas || [];
  const penerima = fokus.filter(function (l) { return cnt[l] !== undefined && cnt[l] < target; })
    .sort(function (a, b) { return cnt[a] - cnt[b]; });
  penerima.forEach(function (U) {
    let guard = 0;
    while (cnt[U] < target && guard++ < 500) {
      let moved = false;
      for (let k = 0; k < slots.length; k++) {
        const s = slots[k];
        if (s.low === U || cnt[s.low] === undefined || cnt[s.low] <= target) continue;   // donor harus berlebih
        if (liburMap[namaAsli[U]] && liburMap[namaAsli[U]][s.t]) continue;                // U libur hari itu
        if (inShift[U + '|' + s.t + '|' + s.sh]) continue;                                // U sudah di shift itu
        const donor = namaAsli[s.low] || s.low;
        pr.getRange(s.i + 2, s.c + 1).setNumberFormat('@').setValue(namaAsli[U]);
        grid[s.i][s.c] = namaAsli[U];
        inShift[s.low + '|' + s.t + '|' + s.sh] = false;
        cnt[s.low]--; cnt[U]++;
        detail.push({ ke: namaAsli[U], dari: donor, tgl: s.t, shift: s.sh });
        s.low = U; inShift[U + '|' + s.t + '|' + s.sh] = true;
        pindah++; moved = true; break;
      }
      if (!moved) break;
    }
    if (cnt[U] < 1) sisa.push(namaAsli[U]);
  });
  return { pindah: pindah, sisa: sisa, detail: detail };
}

/** Seimbangkan beban PERAWAT (papan asistensi, 1 perawat/baris): pindahkan slot
 *  dari perawat di ATAS rata-rata ke perawat 'prioritas' (mulai tanpa jadwal),
 *  hormati kapabilitas (N-05), libur, tak-dobel, & aturan pasangan New (A-02).
 *  @param {Array<string>=} prioritas nama lowercase yang wajib dinaikkan (default: none). */
function seimbangkanBebanPerawat_(ss, prioritas) {
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (!pa || pa.getLastRow() < 2) return { pindah: 0, sisa: [], detail: [] };
  const grid = pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues();
  const pMap = petaPerawat_();
  const liburMap = petaLiburPerNama_(ss);
  const aktif = Object.keys(pMap).filter(function (k) { return pMap[k].aktif !== 'Tidak'; });
  if (aktif.length < 2) return { pindah: 0, sisa: [], detail: [] };
  const cnt = {}; aktif.forEach(function (k) { cnt[k] = 0; });
  const shiftKat = {}, slots = [];   // shiftKat[t|sh] = { low: kategori }
  for (let i = 0; i < grid.length; i++) {
    const t = String(grid[i][0]), sh = String(grid[i][2]), spes = String(grid[i][4]);
    const nm = String(grid[i][6] || '').trim(); if (!nm) continue;
    const low = nm.toLowerCase();
    slots.push({ i: i, low: low, t: t, sh: sh, spes: spes });
    if (cnt[low] !== undefined) cnt[low]++;
    const kk = t + '|' + sh; if (!shiftKat[kk]) shiftKat[kk] = {};
    shiftKat[kk][low] = pMap[low] ? pMap[low].kategori : 'Existing';
  }
  const totalAssign = slots.filter(function (s) { return cnt[s.low] !== undefined; }).length;
  const target = Math.floor(totalAssign / aktif.length);
  if (target < 1) return { pindah: 0, sisa: [], detail: [] };
  const adaExistingLain = function (kk, kecuali) {   // ada Existing selain 'kecuali' di shift?
    const set = shiftKat[kk] || {}; let ada = false;
    Object.keys(set).forEach(function (l) { if (l !== kecuali && set[l] === 'Existing') ada = true; });
    return ada;
  };
  let pindah = 0; const sisa = [], detail = [];
  const fokus = prioritas || [];
  const penerima = fokus.filter(function (k) { return cnt[k] !== undefined && cnt[k] < target; })
    .sort(function (a, b) { return cnt[a] - cnt[b]; });
  penerima.forEach(function (U) {
    const pu = pMap[U]; if (!pu) return;
    let guard = 0;
    while (cnt[U] < target && guard++ < 800) {
      let moved = false;
      for (let k = 0; k < slots.length; k++) {
        const s = slots[k];
        if (s.low === U || cnt[s.low] === undefined || cnt[s.low] <= target) continue;
        if (pu.kapabilitasSet.length && s.spes && pu.kapabilitasSet.indexOf(s.spes) === -1) continue; // N-05
        if (liburMap[pu.nama] && liburMap[pu.nama][s.t]) continue;
        const kk = s.t + '|' + s.sh;
        if (shiftKat[kk] && shiftKat[kk][U] !== undefined) continue;                       // sudah di shift itu
        if (pu.kategori === 'New' && !adaExistingLain(kk, s.low)) continue;                // A-02
        const donor = pMap[s.low] ? pMap[s.low].nama : s.low;
        pa.getRange(s.i + 2, PA_KOL_PERAWAT).setValue(pu.nama);
        grid[s.i][6] = pu.nama;
        if (shiftKat[kk]) { delete shiftKat[kk][s.low]; shiftKat[kk][U] = pu.kategori; }
        cnt[s.low]--; cnt[U]++;
        detail.push({ ke: pu.nama, dari: donor, tgl: s.t, shift: s.sh, spes: s.spes });
        s.low = U;
        pindah++; moved = true; break;
      }
      if (!moved) break;
    }
    if (cnt[U] < 1) sisa.push(pu.nama);
  });
  return { pindah: pindah, sisa: sisa, detail: detail };
}

/** Perbaiki SEMUA blocker yang aman diotomasi (non-destruktif):
 *  A-05 (slot asisten), R-03/R-04 (shift), N-02/R-06 (libur),
 *  R-09/A-07 (staf belum terjadwal → diseimbangkan dari staf yang kelebihan). */
function perbaikiBlockerWeb() {
  wajibPeran_(['Admin', 'SPV']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Tangkap siapa yang BELUM terjadwal SEBELUM pengisian, agar tetap jadi target
  // penyeimbangan walau sebagian sudah terisi di langkah berikutnya.
  const rcp0 = namaTanpaJadwalResepsionis_(ss);      // resepsionis 0 shift (R-09)
  const prw0 = namaTanpaJadwalPerawat_(ss);          // perawat 0 slot (A-07)
  // Segarkan dropdown papan agar nama staf baru (mis. yg belum terjadwal) tak ditolak validasi
  refreshResepsionisValidation_(ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST));
  refreshAsistensiValidation_(ss.getSheetByName(SHEETS.PAPAN_ASISTENSI));
  const lib = topUpLibur_(ss);                       // 1) top-up libur (memengaruhi ketersediaan)
  SpreadsheetApp.flush();
  const a = isiAsistensiKosong_(ss);                 // 2) isi slot asisten kosong (A-05)
  const r = isiResepsionisKurang_(ss);               // 3) isi shift resepsionis kurang (R-03/R-04)
  SpreadsheetApp.flush();
  const brc = seimbangkanBebanResepsionis_(ss, rcp0);// 4) seimbangkan resepsionis tanpa jadwal → rata (R-09)
  const brp = seimbangkanBebanPerawat_(ss, prw0);    // 5) seimbangkan perawat tanpa jadwal → rata (A-07)
  SpreadsheetApp.flush();
  const rj = jadwalkanResepsionisTakTerjadwal_(ss);  // 6) fallback: resepsionis yg MASIH 0-shift → slot kosong
  const pindahRcp = rj.terjadwal + brc.pindah;
  const log = [
    'Libur: +' + lib.tambah + ' hari' + (lib.gagal ? ' (' + lib.gagal + ' jatah belum terpenuhi — kapasitas/hari kerja tak cukup)' : '') + '.',
    'Asistensi: ' + a.terisi + ' slot kosong terisi' + (a.sisa ? ', ' + a.sisa + ' tetap kosong (tak ada perawat memenuhi syarat)' : '') + '.',
    'Resepsionis: ' + (r.terisi + pindahRcp) + ' penugasan disesuaikan' + (brc.sisa.length ? ' (belum tuntas: ' + brc.sisa.join(', ') + ')' : '') + '.',
    'Perawat: ' + brp.pindah + ' slot diseimbangkan ke perawat yg belum terjadwal' + (brp.sisa.length ? ' (belum tuntas: ' + brp.sisa.join(', ') + ')' : '') + '.'
  ];
  const rincian = {
    libur: lib.detail || [],                 // {nama, tgl, peran}
    asistensiIsi: a.detail || [],            // {nama, tgl, shift, dokter}
    resepsionisIsi: r.detail || [],          // {nama, tgl, shift}
    resepsionisSwap: brc.detail || [],       // {ke, dari, tgl, shift}
    resepsionisTaruh: rj.detail || [],       // {nama, tgl, shift}
    perawatSwap: brp.detail || []            // {ke, dari, tgl, shift, spes}
  };
  return { ok: true, log: log, rincian: rincian,
    pesan: 'Perbaikan otomatis selesai. ' + log.join(' ') +
      ' Catatan: penyeimbangan bisa menggeser jumlah longshift (A-04/R-07) — bila butuh presisi longshift, jalankan Generate ulang.' };
}

/** Assign perawat ke slot + validasi; kosongkan bila nama '' */
function assignPerawatWeb(baris, nama) {
  wajibPeran_(['Admin', 'SPV']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  pa.getRange(baris, PA_KOL_PERAWAT).setValue(nama || '');
  const v = validasiAssignWeb_(ss, baris);
  // warna sel
  const sel = pa.getRange(baris, PA_KOL_PERAWAT);
  sel.setBackground(v.level === 'merah' ? WARNA_PELANGGARAN : v.level === 'oranye' ? WARNA_WARNING : (nama ? WARNA_VALID : null));
  sel.setNote(v.masalah.join('\n'));
  return { ok: true, level: v.level, masalah: v.masalah };
}

/** Validasi 1 baris asistensi (UI-free) → {level, masalah[]} */
function validasiAssignWeb_(ss, baris) {
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  const info = pa.getRange(baris, 1, 1, PA_KOL_PERAWAT).getValues()[0];
  const nama = String(info[6]).trim();
  const masalah = []; let level = 'ok';
  if (!nama) return { level: 'ok', masalah: [] };
  const tanggal = String(info[0]), shift = String(info[2]), dokter = String(info[3]), spes = String(info[4]);
  const perawatMap = petaPerawat_();
  const liburMap = petaLiburPerNama_(ss);
  const p = perawatMap[nama.toLowerCase()];
  if (!p) { return { level: 'merah', masalah: ['Perawat tidak ada di master.'] }; }
  if (p.aktif === 'Tidak') { masalah.push('Perawat NONAKTIF.'); level = 'merah'; }
  if (liburMap[p.nama] && liburMap[p.nama][tanggal]) { masalah.push('Bentrok: sedang libur tanggal ini.'); level = 'merah'; }
  if (p.kapabilitasSet.length && p.kapabilitasSet.indexOf(spes) === -1) { masalah.push('N-05: di luar kapabilitas ' + spes + '.'); if (level !== 'merah') level = 'oranye'; }
  const semua = pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues();
  let adaNewLain = false;
  for (let i = 0; i < semua.length; i++) {
    const r = semua[i], rRow = i + 2; if (rRow === baris) continue;
    const rNama = String(r[6]).trim(); if (!rNama) continue;
    if (rNama.toLowerCase() === nama.toLowerCase() && String(r[0]) === tanggal && String(r[2]) === shift) { masalah.push('Dobel assign (baris ' + rRow + ').'); level = 'merah'; }
    if (String(r[0]) === tanggal && String(r[2]) === shift && String(r[3]) === dokter) { const rk = perawatMap[rNama.toLowerCase()]; if (rk && rk.kategori === 'New') adaNewLain = true; }
  }
  if (p.kategori === 'New' && adaNewLain) { masalah.push('A-02: dua NEW berdampingan.'); level = 'merah'; }
  return { level: level, masalah: masalah };
}

/* ================= FINALISASI ================= */

function cekFinalWeb() {
  wajibModul_('aksi');
  const h = auditInti_();
  if (!h.ok) return { ok: false, pesan: h.pesan };
  tulisRekapKepatuhan_(h);
  const b = hitungBlocker_(h);
  if (b.total > 0) {
    setConfig_('STATUS_JADWAL', 'BELUM FINAL — ' + b.total + ' blocker', 'Status quality gate');
    return { ok: true, final: false, blocker: b.total, daftar: b.daftar.slice(0, 30), rinci: b.rinci.slice(0, 30) };
  }
  const stempel = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  setConfig_('STATUS_JADWAL', 'FINAL ✓ ' + stempel, 'Status quality gate');
  try { simpanArsipJadwal_(stempel); } catch (e) {}   // arsipkan snapshot periode ini
  return { ok: true, final: true, stempel: stempel, warning: h.pelanggaran.length };
}

/* ================= MASTER DATA (dibungkus akses) ================= */

function webGetMaster(jenis) { wajibModul_('master'); return getMasterData(jenis); }
function webAddMaster(jenis, data) { wajibModul_('master'); return addMasterData(jenis, data); }
function webUpdateMaster(jenis, id, data) { wajibModul_('master'); return updateMasterData(jenis, id, data); }
function webSetAktif(jenis, id, status) { wajibModul_('master'); return setStatusAktif(jenis, id, status); }
function webGetSpesialisasi() { return getSpesialisasiList(); }

/* ================= REQUEST CUTI (dibungkus akses) ================= */

function webGetStaf() { wajibModul_('cuti'); return getStafListForRequest(); }
function webAddCuti(data) { wajibModul_('cuti'); return addRequestCuti(data); }
function webUpdateCuti(baris, data) { wajibModul_('cuti'); return updateRequestCuti(baris, data); }
function webListCuti() { wajibModul_('cuti'); return getRequestCutiList(); }
function webDeleteCuti(baris) { wajibModul_('cuti'); return deleteRequestCuti(baris); }

function webAddJaga(data) { wajibModul_('cuti'); return addRequestJaga(data); }
function webUpdateJaga(baris, data) { wajibModul_('cuti'); return updateRequestJaga(baris, data); }
function webListJaga() { wajibModul_('cuti'); return getRequestJagaList(); }
function webDeleteJaga(baris) { wajibModul_('cuti'); return deleteRequestJaga(baris); }