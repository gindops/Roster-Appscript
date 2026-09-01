/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Master Data (CRUD) — v2
 * File: MasterData.gs
 * Dipanggil dari sidebar (SidebarMasterData.html)
 *
 * v2:
 *  - Dokter: field pola praktik (hari, jam awal/akhir, jam shift,
 *    kategori shift). Multi-baris per dokter dengan ID SAMA;
 *    duplikat ditolak hanya jika nama+hari+kategori shift sama.
 *  - Perawat: field Kapabilitas Spesialis (kode dipisah koma).
 *    Catatan: kapabilitas TIDAK mengubah kewajiban rotasi A-03
 *    (keputusan #17) — hanya filter eligibility saat assign.
 * ============================================================
 */

/** Konfigurasi tiap jenis master data (urutan = kolom sheet) */
const MD_CONFIG = {
  dokter: {
    sheet: 'MD_Dokter',
    prefix: 'DOK',
    kolom: ['id', 'nama', 'spesialisasi', 'kebutuhanAsisten', 'hariPraktek',
            'jamAwal', 'jamAkhir', 'jamShift', 'kategoriShift', 'aktif', 'email'],
    label: 'Dokter'
  },
  perawat: {
    sheet: 'MD_Perawat',
    prefix: 'PRW',
    kolom: ['id', 'nama', 'level', 'kategori', 'kapabilitas', 'aktif', 'email'],
    label: 'Perawat'
  },
  receptionist: {
    sheet: 'MD_Receptionist',
    prefix: 'RCP',
    kolom: ['id', 'nama', 'level', 'aktif', 'email'],
    label: 'Receptionist'
  }
};

/** Tampilkan sidebar input master data */
function showMasterDataSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('SidebarMasterData')
    .setTitle('Input Master Data — Audy Dental');
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Ambil daftar spesialisasi untuk dropdown & checkbox kapabilitas */
function getSpesialisasiList() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MD_SPESIALISASI);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) { return { kode: String(r[0]), nama: String(r[1]) }; });
}

/** Opsi jam shift untuk dropdown form dokter (dibaca dari Config) */
function getJamShiftOptions() {
  return [
    String(getConfig('JAM_SHIFT_PAGI') || '09:00-15:00'),
    String(getConfig('JAM_SHIFT_SIANG') || '15:00-21:00')
  ];
}

/**
 * Ambil semua data master untuk ditampilkan di sidebar.
 * @param {string} jenis - 'dokter' | 'perawat' | 'receptionist'
 */
function getMasterData(jenis) {
  const cfg = MD_CONFIG[jenis];
  if (!cfg) throw new Error('Jenis master data tidak dikenal: ' + jenis);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);
  if (!sh || sh.getLastRow() < 2) return [];
  const nCols = cfg.kolom.length;
  return sh.getRange(2, 1, sh.getLastRow() - 1, nCols).getValues()
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      const obj = {};
      cfg.kolom.forEach(function (k, i) { obj[k] = String(r[i]); });
      return obj;
    });
}

/**
 * Tambah 1 record master data.
 * Dokter: nama yang sudah ada → baris jadwal baru dengan ID sama;
 *         duplikat nama+hari+kategori shift ditolak.
 * Perawat/Receptionist: duplikat nama ditolak.
 * @return {Object} {ok, pesan, id}
 */
function addMasterData(jenis, data) {
  // Lock bersifat opsional — mencegah ID ganda saat 2 user input bersamaan.
  // LockService kadang melempar PERMISSION_DENIED (gangguan server Google);
  // bila gagal, tetap lanjut agar penyimpanan tidak terblokir.
  let lock = null;
  try {
    lock = LockService.getDocumentLock();
    lock.waitLock(8000);
  } catch (errLock) {
    lock = null;
  }
  try {
    const cfg = MD_CONFIG[jenis];
    if (!cfg) throw new Error('Jenis master data tidak dikenal: ' + jenis);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);
    if (!sh) throw new Error('Sheet ' + cfg.sheet + ' belum ada. Jalankan Setup Awal Template dulu.');

    // --- Validasi wajib isi ---
    const nama = String(data.nama || '').trim();
    if (!nama) return { ok: false, pesan: 'Nama wajib diisi.' };
    if (jenis === 'dokter') {
      if (!data.spesialisasi) return { ok: false, pesan: 'Spesialisasi wajib dipilih.' };
      if (!data.kebutuhanAsisten || Number(data.kebutuhanAsisten) < 1)
        return { ok: false, pesan: 'Kebutuhan asisten wajib diisi (minimal 1).' };
      if (!data.hariPraktek) return { ok: false, pesan: 'Hari praktek wajib dipilih.' };
      if (!data.kategoriShift) return { ok: false, pesan: 'Kategori shift wajib dipilih.' };
      if (!data.jamShift) return { ok: false, pesan: 'Jam shift wajib dipilih.' };
    }
    if ((jenis === 'perawat' || jenis === 'receptionist') && !data.level)
      return { ok: false, pesan: 'Level (Gold/Silver) wajib dipilih.' };
    if (jenis === 'perawat' && !data.kategori)
      return { ok: false, pesan: 'Kategori (Existing/New) wajib dipilih.' };

    const existing = getMasterData(jenis);
    let id = null;

    if (jenis === 'dokter') {
      // Cek duplikat jadwal: nama + hari + kategori shift
      const dupJadwal = existing.some(function (r) {
        return r.nama.toLowerCase() === nama.toLowerCase() &&
               r.hariPraktek === data.hariPraktek &&
               r.kategoriShift === data.kategoriShift;
      });
      if (dupJadwal) return {
        ok: false,
        pesan: 'Jadwal duplikat: ' + nama + ' sudah terdaftar di hari ' +
               data.hariPraktek + ' (' + data.kategoriShift + ').'
      };
      // Nama sudah ada → pakai ID yang sama (multi-baris per dokter)
      const sama = existing.filter(function (r) {
        return r.nama.toLowerCase() === nama.toLowerCase();
      });
      if (sama.length) id = sama[0].id;
    } else {
      // Perawat/Receptionist: duplikat nama ditolak
      const dup = existing.some(function (r) {
        return r.nama.toLowerCase() === nama.toLowerCase();
      });
      if (dup) return { ok: false, pesan: cfg.label + ' dengan nama "' + nama + '" sudah terdaftar.' };
    }

    // --- Generate ID bila belum ada ---
    if (!id) {
      let maxNum = 0;
      existing.forEach(function (r) {
        const m = String(r.id).match(/(\d+)$/);
        if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
      });
      id = cfg.prefix + '-' + ('000' + (maxNum + 1)).slice(-3);
    }

    data.id = id;
    data.aktif = data.aktif || 'Ya';
    const row = cfg.kolom.map(function (k) { return data[k] !== undefined ? data[k] : ''; });
    sh.appendRow(row);

    const notaAkses = daftarAksesDariMaster_(data.email, nama);
    return { ok: true, pesan: cfg.label + ' "' + nama + '" tersimpan (ID ' + id + ').' + notaAkses, id: id };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/**
 * Update record master (edit). Nama boleh diubah (dengan peringatan di UI).
 * ID & status Aktif dipertahankan.
 * @return {Object} {ok, pesan, namaBerubah}
 */
function updateMasterData(jenis, id, data) {
  let lock = null;
  try { lock = LockService.getDocumentLock(); lock.waitLock(8000); } catch (e) { lock = null; }
  try {
    const cfg = MD_CONFIG[jenis];
    if (!cfg) throw new Error('Jenis tidak dikenal: ' + jenis);
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);
    if (!sh || sh.getLastRow() < 2) return { ok: false, pesan: 'Data kosong.' };

    const nama = String(data.nama || '').trim();
    if (!nama) return { ok: false, pesan: 'Nama wajib diisi.' };
    if ((jenis === 'perawat' || jenis === 'receptionist') && !data.level) return { ok: false, pesan: 'Level wajib dipilih.' };
    if (jenis === 'perawat' && !data.kategori) return { ok: false, pesan: 'Kategori wajib dipilih.' };

    // Cari baris berdasarkan ID
    const nCols = cfg.kolom.length;
    const rng = sh.getRange(2, 1, sh.getLastRow() - 1, nCols);
    const vals = rng.getValues();
    let idx = -1;
    for (let i = 0; i < vals.length; i++) { if (String(vals[i][0]) === String(id)) { idx = i; break; } }
    if (idx === -1) return { ok: false, pesan: 'ID ' + id + ' tidak ditemukan.' };

    // Cegah duplikat nama dengan record lain
    const dup = vals.some(function (r) {
      return String(r[0]) !== String(id) && String(r[1]).trim().toLowerCase() === nama.toLowerCase();
    });
    if (dup) return { ok: false, pesan: 'Nama "' + nama + '" sudah dipakai record lain.' };

    const namaLama = String(vals[idx][1]).trim();
    const aktifLama = vals[idx][cfg.kolom.indexOf('aktif')];

    // Susun baris baru: id & aktif dipertahankan, field lain dari data
    const baru = cfg.kolom.map(function (k) {
      if (k === 'id') return id;
      if (k === 'aktif') return aktifLama;
      return data[k] !== undefined ? data[k] : vals[idx][cfg.kolom.indexOf(k)];
    });
    sh.getRange(idx + 2, 1, 1, nCols).setValues([baru]);

    const namaBerubah = namaLama.toLowerCase() !== nama.toLowerCase();
    const notaAkses = daftarAksesDariMaster_(data.email, nama);
    return { ok: true, namaBerubah: namaBerubah,
      pesan: cfg.label + ' "' + nama + '" diperbarui.' +
        (namaBerubah ? ' ⚠️ Nama berubah dari "' + namaLama + '" — generate ulang jadwal agar assignment lama ikut terupdate.' : '') + notaAkses };
  } finally {
    if (lock) { try { lock.releaseLock(); } catch (e) {} }
  }
}

/**
 * Aktif/nonaktifkan record (soft delete).
 * Untuk dokter yang punya beberapa baris jadwal, semua baris dengan ID
 * tersebut ikut diubah statusnya.
 */
function setStatusAktif(jenis, id, status) {
  const cfg = MD_CONFIG[jenis];
  if (!cfg) throw new Error('Jenis master data tidak dikenal: ' + jenis);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(cfg.sheet);
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  const kolAktif = cfg.kolom.indexOf('aktif') + 1;
  let n = 0;
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      sh.getRange(i + 2, kolAktif).setValue(status);
      n++;
    }
  }
  if (!n) return { ok: false, pesan: 'ID ' + id + ' tidak ditemukan.' };
  return {
    ok: true,
    pesan: n + ' baris ' + id + ' diubah menjadi ' + (status === 'Ya' ? 'Aktif' : 'Nonaktif') + '.'
  };
}

/**
 * Auto-daftar email staf master (dokter/perawat/resepsionis) ke Employee
 * Listing (Config_Akses) sebagai Viewer. Dipanggil saat master tersimpan.
 * Mengembalikan catatan untuk ditambahkan ke pesan sukses ('' bila tak ada).
 * Tak menurunkan peran yang sudah ada (lihat daftarkanAksesViewer_ di WebApp.gs).
 */
function daftarAksesDariMaster_(email, nama) {
  const e = String(email || '').trim();
  if (!e) return '';
  try {
    if (daftarkanAksesViewer_(e, nama)) return ' Email ' + e + ' didaftarkan ke Employee Listing sebagai Viewer.';
  } catch (err) {}
  return '';
}