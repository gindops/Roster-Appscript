/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Arsip Jadwal Final
 * File: ArsipJadwal.gs
 *
 * Saat jadwal dinyatakan FINAL (menu ✅ Cek Final atau tombol web),
 * snapshot ringkas jadwal periode itu disimpan ke sheet tersembunyi
 * 'Arsip_Jadwal'. Dengan begitu tab "Cetak Jadwal" di web app bisa
 * menampilkan & mengunduh jadwal per periode bulan yang sudah final,
 * meski papan kerja sudah ditimpa periode berikutnya.
 *
 * Struktur baris (tanpa header): 1 baris per periode
 *   Periode | Mulai | Akhir | Tanggal Final | JSON(chunk1) | JSON(chunk2) | ...
 * JSON dipecah per 40.000 karakter agar aman dari batas 50.000/sel.
 * ============================================================
 */

const SHEET_ARSIP = 'Arsip_Jadwal';
const ARSIP_CHUNK = 40000;

/** Bentuk snapshot ringkas dari kondisi papan saat ini (via getKalenderData). */
function snapshotKalenderCompact_() {
  const k = getKalenderData();
  if (!k.ok) return null;
  const map = {};
  k.dates.forEach(function (t) {
    const d = k.map[t]; if (!d) return;
    map[t] = {
      hari: d.hari, isWeekend: d.isWeekend,
      asis: (d.asis || []).map(function (a) { return { shift: a.shift, dokter: a.dokter, spes: a.spes, perawat: a.perawat }; }),
      resep: (d.resep || []).map(function (r) { return { shift: r.shift, isi: r.isi }; }),
      libur: d.libur || []
    };
  });
  return { mulai: k.mulai, akhir: k.akhir, dates: k.dates, map: map };
}

/** Simpan/replace snapshot periode aktif ke Arsip_Jadwal. Dipanggil saat FINAL. */
function simpanArsipJadwal_(stempel) {
  const snap = snapshotKalenderCompact_();
  if (!snap) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_ARSIP);
  if (!sh) { sh = ss.insertSheet(SHEET_ARSIP); try { sh.hideSheet(); } catch (e) {} }

  const periode = snap.mulai + ' – ' + snap.akhir;
  const json = JSON.stringify(snap);
  const chunks = [];
  for (let i = 0; i < json.length; i += ARSIP_CHUNK) chunks.push(json.substr(i, ARSIP_CHUNK));

  // Cari baris existing dengan periode sama (upsert)
  let target = -1;
  if (sh.getLastRow() >= 1) {
    const col = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
    for (let r = 0; r < col.length; r++) { if (String(col[r][0]) === periode) { target = r + 1; break; } }
  }
  if (target === -1) target = sh.getLastRow() + 1;
  else sh.getRange(target, 1, 1, sh.getMaxColumns()).clearContent();

  const rowVals = [periode, snap.mulai, snap.akhir, stempel || ''].concat(chunks);
  sh.getRange(target, 1, 1, 4).setNumberFormat('@');
  sh.getRange(target, 1, 1, rowVals.length).setValues([rowVals]);
}

/** Daftar periode arsip final (terbaru dulu) — dipakai web. */
function getArsipList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_ARSIP);
  if (!sh || sh.getLastRow() < 1) return [];
  const rows = sh.getRange(1, 1, sh.getLastRow(), 4).getValues();
  const out = [];
  rows.forEach(function (r) {
    if (!String(r[0]).trim()) return;
    out.push({ periode: String(r[0]), mulai: String(r[1]), akhir: String(r[2]), final: String(r[3]) });
  });
  out.reverse();
  return out;
}

/** Ambil snapshot 1 periode arsip → struktur seperti getKalenderData. */
function getArsipData(periode) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET_ARSIP);
  if (!sh || sh.getLastRow() < 1) return { ok: false, pesan: 'Arsip kosong.' };
  const last = sh.getLastColumn();
  const rows = sh.getRange(1, 1, sh.getLastRow(), last).getValues();
  for (let r = 0; r < rows.length; r++) {
    if (String(rows[r][0]) === periode) {
      let json = '';
      for (let c = 4; c < last; c++) json += String(rows[r][c] || '');
      try {
        const snap = JSON.parse(json);
        return { ok: true, periode: periode, final: String(rows[r][3]), mulai: snap.mulai, akhir: snap.akhir, dates: snap.dates, map: snap.map };
      } catch (e) { return { ok: false, pesan: 'Data arsip rusak.' }; }
    }
  }
  return { ok: false, pesan: 'Periode tak ditemukan di arsip.' };
}