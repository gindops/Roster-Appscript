/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Request Jaga (permintaan dijadwalkan)
 * File: RequestJaga.gs
 *
 * Perawat/receptionist dapat MEMINTA dijadwalkan JAGA pada tanggal
 * tertentu (kebalikan dari cuti). Dicatat SPV/Admin di sheet
 * Request_Jaga. Saat auto-generate, permintaan ini menjadi
 * PREFERENSI (bias skor) — sistem berusaha menempatkan orang tsb
 * di tanggal itu selama tidak melanggar aturan.
 *
 * Format sheet Request_Jaga (mulai baris 2):
 *   Nama | Peran | Tanggal Mulai | Tanggal Akhir (opsional) | Keterangan
 * ============================================================
 */

/** Dropdown Nama di sheet Request_Jaga */
function pasangValidasiRequestJaga_(ss) {
  const rj = ss.getSheetByName(SHEETS.REQUEST_JAGA);
  if (!rj) return;
  const MAXR = 300;
  const nama = getMasterData('perawat').filter(function (p) { return p.aktif !== 'Tidak'; }).map(function (p) { return p.nama; })
    .concat(getMasterData('receptionist').filter(function (r) { return r.aktif !== 'Tidak'; }).map(function (r) { return r.nama; }));
  const unik = [];
  nama.forEach(function (n) { if (unik.indexOf(n) === -1) unik.push(n); });
  if (unik.length) {
    const rule = SpreadsheetApp.newDataValidation().requireValueInList(unik, true).setAllowInvalid(true).build();
    rj.getRange(2, 1, MAXR, 1).setDataValidation(rule);
  }
}

/**
 * Baca Request_Jaga → { map: { namaLower: { 'dd/MM/yyyy': true } }, jumlah }
 * Dipakai auto-generate sebagai preferensi penempatan.
 */
/** Normalisasi shift → "Pagi", "Siang", atau "Pagi, Siang" (kosong/Full → keduanya) */
function normShift_(s) {
  const str = String(s || '');
  const t = [];
  if (/pagi/i.test(str)) t.push('Pagi');
  if (/siang/i.test(str)) t.push('Siang');
  if (!t.length) { t.push('Pagi'); t.push('Siang'); }
  return t.join(', ');
}

function bacaRequestJaga_() {
  const hasil = { map: {}, jumlah: 0 };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rj = ss.getSheetByName(SHEETS.REQUEST_JAGA);
  if (!rj || rj.getLastRow() < 2) return hasil;
  const rows = rj.getRange(2, 1, rj.getLastRow() - 1, HEADER.REQUEST_JAGA.length).getValues();
  rows.forEach(function (r) {
    const nama = String(r[0]).trim();
    if (!nama) return;
    const mulai = parseTgl_(normalTgl_(r[2]));
    if (!mulai) return;
    let akhir = parseTgl_(normalTgl_(r[3]));
    if (!akhir || akhir < mulai) akhir = mulai;
    const shift = normShift_(r[5]); // 'Pagi', 'Siang', atau 'Pagi, Siang'
    const key = nama.toLowerCase();
    if (!hasil.map[key]) hasil.map[key] = {};
    for (let d = new Date(mulai); d <= akhir; d.setDate(d.getDate() + 1)) hasil.map[key][formatTgl_(d)] = shift; // tanggal → 'Pagi'|'Siang'|'Pagi, Siang'
    hasil.jumlah++;
  });
  return hasil;
}

/**
 * Cek aturan saat request jaga dibuat: WARNING (tidak memblokir) bila
 * bertabrakan dengan request CUTI orang yang sama pada tanggal tsb.
 */
function cekAturanRequestJaga_(nama, mulaiTgl, akhirTgl) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const low = String(nama || '').toLowerCase();
  const prw = getMasterData('perawat').filter(function (p) { return p.nama.toLowerCase() === low; })[0];
  const rcp = getMasterData('receptionist').filter(function (r) { return r.nama.toLowerCase() === low; })[0];
  if (!prw && !rcp) return '⚠ ' + nama + ' bukan perawat/resepsionis — hanya perawat & resepsionis yang bisa request jadwal jaga.';

  const dM = parseTgl_(mulaiTgl); if (!dM) return '';
  const dA = akhirTgl ? (parseTgl_(akhirTgl) || dM) : dM;
  const tgls = [];
  for (let d = new Date(dM); d <= dA; d.setDate(d.getDate() + 1)) tgls.push(formatTgl_(d));

  // Data pendukung
  const pMulai = parseTgl_(getConfig('PERIODE_MULAI'));
  const pAkhir = parseTgl_(getConfig('PERIODE_AKHIR'));
  const liburMap = petaLiburPerNama_(ss);
  const cuti = bacaRequestCuti_();
  const cutiSet = {};
  Object.keys(cuti.staf || {}).forEach(function (nm) { if (nm.toLowerCase() === low) (cuti.staf[nm] || []).forEach(function (t) { cutiSet[t] = true; }); });
  // Tanggal ada praktik dokter (dari Jadwal_Dokter) → dasar ada tidaknya slot asistensi / DU
  const praktikTgl = {};
  const jd = ss.getSheetByName(SHEETS.JADWAL_DOKTER);
  if (jd && jd.getLastRow() > 1) {
    jd.getRange(2, 1, jd.getLastRow() - 1, 1).getValues().forEach(function (r) { const t = fmtTglCell_(r[0]); if (t) praktikTgl[t] = (praktikTgl[t] || 0) + 1; });
  }

  const adaJadwalDokter = Object.keys(praktikTgl).length > 0; // cek praktik hanya bila jadwal dokter sudah ada
  const luarPeriode = [], bentrokCuti = [], bentrokLibur = [], takPraktik = [];
  tgls.forEach(function (t) {
    const d = parseTgl_(t);
    if (pMulai && pAkhir && (d < pMulai || d > pAkhir)) luarPeriode.push(t);
    if (cutiSet[t]) bentrokCuti.push(t);
    if (liburMap[nama] && liburMap[nama][t]) bentrokLibur.push(t);
    if (adaJadwalDokter && !praktikTgl[t]) takPraktik.push(t);
  });

  const pesan = [];
  if (luarPeriode.length) pesan.push('Di luar periode aktif (' + (getConfig('PERIODE_MULAI') || '?') + '–' + (getConfig('PERIODE_AKHIR') || '?') + '): ' + luarPeriode.join(', ') + ' → request tak berpengaruh.');
  if (bentrokCuti.length) pesan.push('Bentrok request CUTI: ' + bentrokCuti.join(', ') + ' → cuti diprioritaskan, jaga diabaikan.');
  if (bentrokLibur.length) pesan.push('Sudah dijadwalkan LIBUR: ' + bentrokLibur.join(', ') + ' → tak bisa jaga di tanggal itu.');
  if (takPraktik.length) pesan.push('Tak ada praktik dokter: ' + takPraktik.join(', ') + ' → ' + (prw ? 'tak ada slot asistensi' : 'DU 0, resepsionis tak dibutuhkan') + '.');

  if (!pesan.length) return '';
  return '⚠ Perhatian aturan:\n• ' + pesan.join('\n• ');
}

/* ================= CRUD (web) ================= */

function getRequestJagaList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rj = ss.getSheetByName(SHEETS.REQUEST_JAGA);
  if (!rj || rj.getLastRow() < 2) return [];
  const rows = rj.getRange(2, 1, rj.getLastRow() - 1, HEADER.REQUEST_JAGA.length).getValues();
  const out = [];
  rows.forEach(function (r, i) {
    if (!String(r[0]).trim()) return;
    const shift = normShift_(r[5]);
    const mulai = fmtTglCell_(r[2]), akhir = fmtTglCell_(r[3]);
    const tersimpan = String(r[6] || '').trim(); // Catatan Aturan disimpan saat pembuatan
    const warning = tersimpan || cekAturanRequestJaga_(String(r[0]), mulai, akhir);
    out.push({ baris: i + 2, nama: String(r[0]), peran: String(r[1]), mulai: mulai, akhir: akhir, keterangan: String(r[4]), shift: shift, catatan: tersimpan, warning: warning });
  });
  return out;
}

function addRequestJaga(data) {
  let lock = null;
  try { lock = LockService.getDocumentLock(); lock.waitLock(8000); } catch (e) { lock = null; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rj = ss.getSheetByName(SHEETS.REQUEST_JAGA);
    if (!rj) return { ok: false, pesan: 'Sheet Request_Jaga belum ada. Jalankan Setup Awal Template.' };
    const nama = String(data.nama || '').trim();
    if (!nama) return { ok: false, pesan: 'Nama wajib dipilih.' };
    const mulai = isoKeTgl_(data.mulai);
    if (!mulai) return { ok: false, pesan: 'Tanggal mulai wajib diisi.' };
    let akhir = isoKeTgl_(data.akhir);
    const dM = parseTgl_(mulai), dA = akhir ? parseTgl_(akhir) : null;
    if (dA && dA < dM) return { ok: false, pesan: 'Tanggal akhir tidak boleh sebelum mulai.' };
    const shift = normShift_(data.shift);
    const catatan = cekAturanRequestJaga_(nama, mulai, akhir); // detail aturan saat pembuatan
    rj.appendRow([nama, String(data.peran || ''), '', '', String(data.keterangan || ''), shift, catatan]);
    const rr = rj.getLastRow();
    rj.getRange(rr, 3, 1, 2).setNumberFormat('@').setValues([[mulai, akhir || '']]);
    return { ok: true, pesan: 'Request jaga untuk ' + nama + ' (' + shift + ') tersimpan.', warning: catatan };
  } finally { if (lock) { try { lock.releaseLock(); } catch (e) {} } }
}

function updateRequestJaga(baris, data) {
  let lock = null;
  try { lock = LockService.getDocumentLock(); lock.waitLock(8000); } catch (e) { lock = null; }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rj = ss.getSheetByName(SHEETS.REQUEST_JAGA);
    if (!rj || baris < 2 || baris > rj.getLastRow()) return { ok: false, pesan: 'Baris tidak valid.' };
    const nama = String(data.nama || '').trim();
    if (!nama) return { ok: false, pesan: 'Nama wajib dipilih.' };
    const mulai = isoKeTgl_(data.mulai);
    if (!mulai) return { ok: false, pesan: 'Tanggal mulai wajib diisi.' };
    let akhir = isoKeTgl_(data.akhir);
    const dM = parseTgl_(mulai), dA = akhir ? parseTgl_(akhir) : null;
    if (dA && dA < dM) return { ok: false, pesan: 'Tanggal akhir tidak boleh sebelum mulai.' };
    const shift = normShift_(data.shift);
    const catatan = cekAturanRequestJaga_(nama, mulai, akhir); // detail aturan diperbarui saat edit
    rj.getRange(baris, 3, 1, 2).setNumberFormat('@');
    rj.getRange(baris, 1, 1, HEADER.REQUEST_JAGA.length).setValues([[nama, String(data.peran || ''), mulai, akhir || '', String(data.keterangan || ''), shift, catatan]]);
    return { ok: true, pesan: 'Request jaga ' + nama + ' (' + shift + ') diperbarui.', warning: catatan };
  } finally { if (lock) { try { lock.releaseLock(); } catch (e) {} } }
}

function deleteRequestJaga(baris) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rj = ss.getSheetByName(SHEETS.REQUEST_JAGA);
  if (!rj || baris < 2 || baris > rj.getLastRow()) return { ok: false, pesan: 'Baris tidak valid.' };
  rj.deleteRow(baris);
  return { ok: true, pesan: 'Request jaga dihapus.' };
}