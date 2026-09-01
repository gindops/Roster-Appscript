/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Auto-Generate Asistensi Perawat
 * File: AutoAsistensi.gs
 *
 * Menu: 🤖 Auto-Generate Asistensi
 * Mengisi kolom Perawat di Papan_Asistensi secara otomatis:
 *  - Menghormati Papan_Libur (yang libur tidak dijadwalkan)
 *  - N-05 : hanya perawat berkapabilitas spesialis tsb
 *  - A-02 : tidak ada 2 perawat New berdampingan (dokter+shift+tgl)
 *  - Tidak dobel: 1 perawat 1 slot per tanggal+shift
 *  - Mengejar A-04: pasangan pagi+siang di tanggal sama diprioritaskan
 *    bagi yang belum genap longshift wajib
 *  - Mengejar A-03: perawat yang belum pernah ke spesialis tsb
 *    di periode ini diprioritaskan (rotasi)
 *  - Beban kerja diratakan antar perawat
 *
 * Hasil adalah DRAF — edit manual tetap divalidasi real-time,
 * dan 🔍 Cek Kepatuhan Periode tetap penentu akhir.
 * ============================================================
 */

/** Menu: 🤖 Auto-Generate Asistensi */
function autoGenerateAsistensi() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (!pa || pa.getLastRow() < 2) {
    laporOtomatis_('Auto-Asistensi', 'Papan_Asistensi belum ada.\nJalankan dulu 🗓️ Buat Papan Asistensi.');
    return;
  }

  if (!konfirmasiOtomatis_('Auto-Generate Asistensi',
    'Seluruh kolom Perawat di Papan_Asistensi akan DIISI ULANG otomatis.\n' +
    'Assignment yang sudah ada akan ditimpa. Lanjutkan?')) return;

  // --- Susun input ---
  const n = pa.getLastRow() - 1;
  const grid = pa.getRange(2, 1, n, PA_KOL_PERAWAT).getValues();
  const slots = grid.map(function (r, i) {
    return {
      idx: i, tanggal: String(r[0]), hari: String(r[1]), shift: String(r[2]),
      dokter: String(r[3]), spes: String(r[4])
    };
  });

  const perawat = getMasterData('perawat')
    .filter(function (p) { return p.aktif !== 'Tidak'; })
    .map(function (p) {
      return {
        nama: p.nama, level: p.level, kategori: p.kategori,
        kapab: String(p.kapabilitas || '').split(',')
          .map(function (s) { return s.trim(); }).filter(function (s) { return s; })
      };
    });
  if (!perawat.length) { laporOtomatis_('Auto-Asistensi', 'MD_Perawat kosong.'); return; }

  const hasil = susunAsistensi_({
    slots: slots,
    perawat: perawat,
    liburMap: petaLiburPerNama_(ss),
    longshiftTarget: Number(getConfig('LONGSHIFT_WAJIB_PERAWAT')) || 4,
    prefMap: bacaPreferensiDokter_(),
    jagaMap: bacaRequestJaga_().map
  });

  // --- Tulis hasil ke papan ---
  const nilai = hasil.assign.map(function (nm) { return [nm || '']; });
  const bg = hasil.assign.map(function (nm) { return [nm ? WARNA_VALID : WARNA_PELANGGARAN]; });
  pa.getRange(2, PA_KOL_PERAWAT, n, 1).setValues(nilai).setBackgrounds(bg).setNote(null);

  // --- Ringkasan ---
  let pesan = 'Draf asistensi selesai ✅\n\n' +
    '• Slot terisi: ' + hasil.terisi + '/' + slots.length +
    (hasil.kosong ? ' (slot kosong ditandai merah)' : '') + '\n' +
    '• Perawat longshift genap ' + hasil.longshiftTarget + 'x: ' +
    hasil.longshiftGenap + '/' + perawat.length + '\n';
  if (hasil.longshiftKurang.length) {
    pesan += '\n⚠️ Longshift belum genap:\n' + hasil.longshiftKurang
      .map(function (k) { return '• ' + k.nama + ': ' + k.longshift + 'x'; }).join('\n') + '\n';
  }
  if (hasil.rotasiKurang.length) {
    pesan += '\n⚠️ Rotasi spesialis belum lengkap:\n' + hasil.rotasiKurang.slice(0, 10)
      .map(function (k) { return '• ' + k.nama + ': belum ke ' + k.kurang.join(', '); }).join('\n') + '\n';
  }
  if (hasil.prefUnmet && hasil.prefUnmet.length) {
    pesan += '\n⭐ Preferensi WAJIB tak terpenuhi (slot dibiarkan kosong):\n' + hasil.prefUnmet.slice(0, 10)
      .map(function (u) { return '• ' + u.tanggal + ' ' + u.shift + ' — ' + u.dokter + ': ' + u.detail; }).join('\n') + '\n';
  }
  pesan += '\nIni DRAF — sesuaikan manual bila perlu.';
  laporOtomatis_('Auto-Asistensi', pesan);
}

/**
 * Algoritma inti (murni — bisa diuji terpisah).
 * @param {Object} input {slots, perawat, liburMap: {nama→{tanggal→true}}, longshiftTarget}
 * @return {Object} {assign:[nama per idx], terisi, kosong, longshift*, rotasiKurang}
 */
function susunAsistensi_(input) {
  const assign = new Array(input.slots.length).fill('');
  const stat = {};   // nama → {total, spesSet, pagi:{tgl}, siang:{tgl}, longshift}
  input.perawat.forEach(function (p) {
    stat[p.nama] = { total: 0, spesSet: {}, pagi: {}, pagiDok: {}, siang: {}, longshift: 0 };
  });

  // Grup A-02 & anti-dobel
  const grupNew = {};    // tanggal|shift|dokter → jumlah New terpasang
  const sudahShift = {}; // nama|tanggal|shift → true
  const prefUnmet = [];  // preferensi Wajib yang tak terpenuhi → slot kosong
  const prefMap = input.prefMap || {}; // namaDokterLower → [ {tipe,nilai,prioritas} ]
  const jagaMap = input.jagaMap || {}; // namaPerawatLower → {tanggal→true} (request jaga)

  // Kumpulan spesialis yang praktik (dasar rotasi A-03)
  const spesPraktik = {};
  input.slots.forEach(function (s) { spesPraktik[s.spes] = true; });

  // Urutkan slot: tanggal, Pagi dulu (agar pasangan longshift bisa dibangun saat Siang)
  const urut = input.slots.slice().sort(function (a, b) {
    const ka = kunciTgl_(a.tanggal), kb = kunciTgl_(b.tanggal);
    if (ka !== kb) return ka < kb ? -1 : 1;
    if (a.shift !== b.shift) return a.shift === 'Pagi' ? -1 : 1;
    return a.idx - b.idx;
  });

  urut.forEach(function (slot) {
    const gKey = slot.tanggal + '|' + slot.shift + '|' + slot.dokter;

    const kandidat = input.perawat.filter(function (p) {
      const s = stat[p.nama];
      // Libur hari itu?
      if (input.liburMap[p.nama] && input.liburMap[p.nama][slot.tanggal]) return false;
      // Sudah bertugas di tanggal+shift ini?
      if (sudahShift[p.nama + '|' + slot.tanggal + '|' + slot.shift]) return false;
      // Kapabilitas (kosong = mampu semua)
      if (p.kapab.length && p.kapab.indexOf(slot.spes) === -1) return false;
      // A-02: New tidak berdampingan New
      if (p.kategori === 'New' && (grupNew[gKey] || 0) > 0) return false;
      return true;
    });

    if (!kandidat.length) return; // slot dibiarkan kosong, dilaporkan

    // Preferensi dokter untuk slot ini
    const prefs = prefMap[slot.dokter.toLowerCase()] || [];
    const wajib = prefs.filter(function (pr) { return pr.prioritas === 'Wajib'; });
    const utamakan = prefs.filter(function (pr) { return pr.prioritas === 'Utamakan'; });
    let pool = kandidat;
    if (wajib.length) {
      // Wajib = filter keras: kandidat harus memenuhi SEMUA preferensi wajib
      const poolW = kandidat.filter(function (p) {
        return wajib.every(function (w) { return cocokPreferensi_(p, w); });
      });
      if (poolW.length) {
        pool = poolW;
      } else {
        // Mustahil dipenuhi → slot dibiarkan kosong & dicatat (sesuai kebijakan Wajib)
        prefUnmet.push({
          tanggal: slot.tanggal, shift: slot.shift, dokter: slot.dokter,
          jenis: 'Wajib', detail: wajib.map(ringkasPreferensi_).join(' + ')
        });
        return;
      }
    }

    // Skor: makin kecil makin baik
    pool.forEach(function (p) {
      const s = stat[p.nama];
      let sk = s.total * 10;                                   // ratakan beban
      if (!s.spesSet[slot.spes]) sk -= 30;                     // kejar rotasi A-03
      if (slot.shift === 'Siang' && s.pagi[slot.tanggal] &&
          s.longshift < input.longshiftTarget) {
        sk -= 40;                                              // pasangkan longshift A-04
        if (s.pagiDok[slot.tanggal] === slot.dokter) sk += 25; // soft: hindari longshift di dokter yang sama
      }
      if (slot.shift === 'Pagi' && s.longshift < input.longshiftTarget) sk -= 5;
      // Bias preferensi Utamakan: perawat yang cocok didahulukan (soft)
      let cocok = 0;
      utamakan.forEach(function (u) { if (cocokPreferensi_(p, u)) cocok++; });
      sk -= cocok * 60;
      // Bias Request Jaga: perawat yang minta dijadwalkan di tanggal+shift ini didahulukan (soft)
      const jsh = jagaMap[p.nama.toLowerCase()] ? jagaMap[p.nama.toLowerCase()][slot.tanggal] : null;
      if (jsh && jsh.indexOf(slot.shift) !== -1) sk -= 70; // 'Pagi, Siang' cocok utk kedua shift
      sk += Math.random() * 4;                                 // variasi antar generate
      p._skor = sk;
    });
    pool.sort(function (a, b) { return a._skor - b._skor; });

    const pilih = pool[0];
    const s = stat[pilih.nama];
    assign[slot.idx] = pilih.nama;
    s.total++;
    s.spesSet[slot.spes] = true;
    if (slot.shift === 'Pagi') { s.pagi[slot.tanggal] = true; s.pagiDok[slot.tanggal] = slot.dokter; }
    else {
      s.siang[slot.tanggal] = true;
      if (s.pagi[slot.tanggal]) s.longshift++;
    }
    sudahShift[pilih.nama + '|' + slot.tanggal + '|' + slot.shift] = true;
    if (pilih.kategori === 'New') grupNew[gKey] = (grupNew[gKey] || 0) + 1;
  });

  // --- Statistik hasil ---
  const daftarSpes = Object.keys(spesPraktik).sort();
  let terisi = 0;
  assign.forEach(function (nm) { if (nm) terisi++; });
  const longshiftKurang = [];
  const rotasiKurang = [];
  let longshiftGenap = 0;
  input.perawat.forEach(function (p) {
    const s = stat[p.nama];
    if (s.longshift >= input.longshiftTarget) longshiftGenap++;
    else longshiftKurang.push({ nama: p.nama, longshift: s.longshift });
    const kurang = daftarSpes.filter(function (sp) {
      // Rotasi hanya relevan bila perawat memang mampu (di luar kapabilitas =
      // tetap dilaporkan oleh audit sebagai "tidak mungkin", bukan di sini)
      if (p.kapab.length && p.kapab.indexOf(sp) === -1) return false;
      return !s.spesSet[sp];
    });
    if (kurang.length) rotasiKurang.push({ nama: p.nama, kurang: kurang });
  });

  return {
    assign: assign, terisi: terisi, kosong: assign.length - terisi,
    longshiftTarget: input.longshiftTarget,
    longshiftGenap: longshiftGenap, longshiftKurang: longshiftKurang,
    rotasiKurang: rotasiKurang, prefUnmet: prefUnmet
  };
}

/** Kunci urut tanggal dd/MM/yyyy → yyyyMMdd */
function kunciTgl_(t) {
  const m = String(t).match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return String(t);
  return m[3] + ('0' + m[2]).slice(-2) + ('0' + m[1]).slice(-2);
}

/** Baca Papan_Libur → {nama → {tanggal → true}} */
function petaLiburPerNama_(ss) {
  const map = {};
  const pl = ss.getSheetByName('Papan_Libur');
  if (!pl || pl.getLastRow() < 2) return map;
  const rows = pl.getRange(2, 1, pl.getLastRow() - 1, 2 + PL_MAX_SLOT).getValues();
  rows.forEach(function (r) {
    const tanggal = String(r[0]);
    for (let c = 2; c < 2 + PL_MAX_SLOT; c++) {
      const nama = String(r[c]).trim();
      if (!nama) continue;
      if (!map[nama]) map[nama] = {};
      map[nama][tanggal] = true;
    }
  });
  return map;
}