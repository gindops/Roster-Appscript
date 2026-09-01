/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Katalog Aturan (kode terpusat)
 * File: AturanKatalog.gs
 *
 * Satu sumber kebenaran untuk SEMUA kode aturan + helper yang
 * mengelompokkan temuan validasi berdasarkan KELUARGA KODE:
 *   N = Perawat (grading/libur/kapabilitas)
 *   A = Asistensi perawat
 *   R = Resepsionis
 *   D = Data & Operasional (lintas peran, tanpa kode N/A/R)
 *
 * Dipakai oleh: cekAturanRequestJaga_, cekAturanRequest_ (cuti),
 * audit Kesesuaian (WebApp.getKepatuhanWeb), dan dokumentasi.
 * ============================================================
 */

/**
 * Katalog kode → { label, kelompok, tingkat }.
 * tingkat: 'blocker' (memblokir finalisasi) | 'warning' (dilaporkan,
 * tak memblokir) | 'atribut' (definisi/master data) | 'info'.
 * Tingkat mengikuti perilaku KODE aktual (hitungBlocker_ di Dashboard.gs).
 */
const KATALOG_ATURAN = {
  // --- Keluarga N: Perawat ---
  'N-01': { label: 'Level grading perawat (Gold / Silver)', kelompok: 'N', tingkat: 'atribut' },
  'N-02': { label: 'Perawat wajib 4× libur per periode', kelompok: 'N', tingkat: 'blocker' },
  'N-03': { label: 'Libur weekend hanya Gold (maks 1×); Silver dilarang', kelompok: 'N', tingkat: 'blocker' },
  'N-04': { label: 'Kategori perawat (Existing / New)', kelompok: 'N', tingkat: 'atribut' },
  'N-05': { label: 'Kapabilitas asistensi spesialis (filter eligibility)', kelompok: 'N', tingkat: 'warning' },
  // --- Keluarga A: Asistensi perawat ---
  'A-01': { label: 'Struktur 2 shift per hari (pagi & siang)', kelompok: 'A', tingkat: 'atribut' },
  'A-02': { label: 'Perawat New tak boleh berdampingan (harus ada Existing)', kelompok: 'A', tingkat: 'blocker' },
  'A-03': { label: 'Rotasi ke semua spesialis yang praktik, min 1× (gugur bila tak praktik)', kelompok: 'A', tingkat: 'warning' },
  'A-04': { label: 'Perawat wajib 4× longshift per periode', kelompok: 'A', tingkat: 'blocker' },
  'A-05': { label: 'Jumlah asisten per dokter mengikuti kebutuhan di master', kelompok: 'A', tingkat: 'blocker' },
  'A-06': { label: 'Asistensi diturunkan dari jadwal praktik dokter (CIS)', kelompok: 'A', tingkat: 'atribut' },
  'A-07': { label: 'Perawat aktif wajib terjadwal asistensi (min 1×) — 0 tugas = blocker', kelompok: 'A', tingkat: 'blocker' },
  // --- Keluarga R: Resepsionis ---
  'R-01': { label: 'Jumlah resepsionis per shift ditentukan DU aktif', kelompok: 'R', tingkat: 'atribut' },
  'R-02': { label: 'Proporsional: shift ber-DU lebih banyak → resepsionis lebih banyak', kelompok: 'R', tingkat: 'warning' },
  'R-03': { label: 'DU aktif 1–4 → minimal 2 resepsionis', kelompok: 'R', tingkat: 'blocker' },
  'R-04': { label: 'DU aktif > 4 → minimal 3 resepsionis', kelompok: 'R', tingkat: 'blocker' },
  'R-05': { label: 'DU aktif dihitung dari jumlah dokter praktik', kelompok: 'R', tingkat: 'atribut' },
  'R-06': { label: 'Resepsionis wajib 4× libur per periode', kelompok: 'R', tingkat: 'blocker' },
  'R-07': { label: 'Resepsionis Silver wajib 4× longshift', kelompok: 'R', tingkat: 'warning' },
  'R-08': { label: 'Level grading resepsionis (Gold / Silver)', kelompok: 'R', tingkat: 'atribut' },
  'R-09': { label: 'Resepsionis aktif wajib terjadwal shift (min 1×) — 0 tugas = blocker', kelompok: 'R', tingkat: 'blocker' },
  // --- Keluarga D: Data & Operasional (lintas peran) ---
  'Dobel': { label: 'Nama tercatat 2× pada tanggal + shift yang sama', kelompok: 'D', tingkat: 'blocker' },
  'Libur': { label: 'Bertugas padahal tercatat libur (bentrok libur)', kelompok: 'D', tingkat: 'blocker' },
  'Master': { label: 'Nama tak ada di master data / staf nonaktif', kelompok: 'D', tingkat: 'blocker' },
  'Periode': { label: 'Tanggal di luar periode aktif', kelompok: 'D', tingkat: 'info' },
  'Cuti': { label: 'Bentrok dengan request cuti orang yang sama', kelompok: 'D', tingkat: 'info' },
  'Peran': { label: 'Bukan perawat/resepsionis (request tidak berlaku)', kelompok: 'D', tingkat: 'invalid' }
};

/** Urutan & label tampilan tiap keluarga */
var URUTAN_KELOMPOK = ['N', 'A', 'R', 'D'];
var NAMA_KELOMPOK = {
  N: 'Perawat (N)',
  A: 'Asistensi (A)',
  R: 'Resepsionis (R)',
  D: 'Data & Operasional'
};

/** Kode → keluarga (huruf depan; selain N/A/R = D). "R-03/R-04" → "R". */
function familiKode_(kode) {
  const k = String(kode || '').toUpperCase().charAt(0);
  if (k === 'N' || k === 'A' || k === 'R') return k;
  return 'D';
}

/** Kode → label katalog (string kosong bila tak dikenal). */
function labelAturan_(kode) {
  const e = KATALOG_ATURAN[String(kode)];
  return e ? e.label : '';
}

/** Kode → tingkat ('blocker'|'warning'|'atribut'|'info'|'invalid'|''). */
function tingkatAturan_(kode) {
  const e = KATALOG_ATURAN[String(kode)];
  return e ? e.tingkat : '';
}

/**
 * Render daftar temuan menjadi teks dikelompokkan per KELUARGA kode.
 * @param {Array<{kode:string, teks:string}>} items
 * @param {string=} judul  Header (default '⚠ Perhatian aturan:')
 * @return {string} '' bila items kosong.
 *
 * Contoh keluaran:
 *   ⚠ Perhatian aturan:
 *   ▸ Perawat (N)
 *     • [N-03] ...
 *   ▸ Data & Operasional
 *     • [Libur] ...
 */
function kelompokkanTemuan_(items, judul) {
  if (!items || !items.length) return '';
  const grup = { N: [], A: [], R: [], D: [] };
  items.forEach(function (it) {
    if (!it || !it.teks) return;
    const g = familiKode_(it.kode);
    grup[g].push('  • [' + it.kode + '] ' + it.teks);
  });
  const bagian = [];
  URUTAN_KELOMPOK.forEach(function (g) {
    if (grup[g].length) bagian.push('▸ ' + NAMA_KELOMPOK[g] + '\n' + grup[g].join('\n'));
  });
  if (!bagian.length) return '';
  return (judul || '⚠ Perhatian aturan:') + '\n' + bagian.join('\n');
}