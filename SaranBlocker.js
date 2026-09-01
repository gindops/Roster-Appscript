/**
 * ============================================================
 * PENJADWALAN AUDY DENTAL — Saran Perbaikan Blocker (dry-run)
 * File: SaranBlocker.gs
 *
 * Alur: "Sarankan perbaikan" (TIDAK menulis) → tampilkan tiap saran
 * sebagai ITEM dengan ikon Terapkan sendiri → klik ikon = terapkan
 * item itu saja. Semua saran divalidasi terhadap SEMUA aturan agar
 * tak memicu blocker baru:
 *   A-02 (New berdampingan, per tgl|shift|DOKTER), N-05 kapabilitas,
 *   libur, tak dobel di shift sama, A-04/R-07 longshift.
 *
 * A-02 diperbaiki dengan MENUKAR 2 perawat DALAM shift yang sama
 * (netral terhadap libur, longshift, & beban — hanya berpindah dokter).
 * Longshift diperbaiki dengan mengambil slot dari donor AMAN.
 *
 * Tiap item membawa 'moves' (sel yang ditulis). Antar-item TAK berbagi
 * sel (penanda 'touched') sehingga tiap ikon berdiri sendiri & aman.
 * ============================================================
 */

/** Muat model papan asistensi ke array {r,tgl,hari,shift,dokter,spes,low,asal}. */
function sb_modelAsis_(ss) {
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  if (!pa || pa.getLastRow() < 2) return { pa: pa, rows: [] };
  const g = pa.getRange(2, 1, pa.getLastRow() - 1, PA_KOL_PERAWAT).getValues();
  const rows = g.map(function (r, i) {
    const nm = String(r[6] || '').trim();
    return { r: i + 2, tgl: String(r[0]), hari: String(r[1]), shift: String(r[2]),
      dokter: String(r[3]), spes: String(r[4]), low: nm.toLowerCase(), asal: nm };
  });
  return { pa: pa, rows: rows };
}

/** Muat model papan resepsionis ke array slot {r,c,tgl,hari,shift,min,low,asal}. */
function sb_modelRcp_(ss) {
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  if (!pr || pr.getLastRow() < 2) return { pr: pr, slots: [] };
  const nCol = PR_KOL_SLOT1 + PR_MAX_SLOT - 1;
  const g = pr.getRange(2, 1, pr.getLastRow() - 1, nCol).getValues();
  const slots = [];
  g.forEach(function (r, i) {
    for (let c = PR_KOL_SLOT1 - 1; c < nCol; c++) {
      const nm = String(r[c] || '').trim();
      slots.push({ r: i + 2, c: c + 1, tgl: String(r[0]), hari: String(r[1]), shift: String(r[2]), low: nm.toLowerCase(), asal: nm });
    }
  });
  return { pr: pr, slots: slots };
}

/** Hitung statistik longshift & beban dari daftar {tgl,shift,low}. */
function sb_statLongshift_(items) {
  const byPerson = {};
  items.forEach(function (it) {
    if (!it.low) return;
    const p = byPerson[it.low] || (byPerson[it.low] = { total: 0, hari: {} });
    p.total++; (p.hari[it.tgl] || (p.hari[it.tgl] = {}))[it.shift] = true;
  });
  Object.keys(byPerson).forEach(function (low) {
    const p = byPerson[low]; let ls = 0;
    Object.keys(p.hari).forEach(function (t) { if (p.hari[t]['Pagi'] && p.hari[t]['Siang']) ls++; });
    p.longshift = ls;
  });
  return byPerson;
}

/** Kapabilitas cocok? (set kosong = universal). */
function sb_bisa_(kapabSet, spes) { return !kapabSet.length || !spes || kapabSet.indexOf(spes) !== -1; }

/**
 * PASS A-02: grup (tgl|shift|dokter) dengan ≥2 New → tukar satu New dengan
 * satu Existing dari grup lain pada shift sama. Menghasilkan ITEM {moves:[2]}.
 */
function sb_passA02_(asis, pMap, liburMap, items, touched, gagal, ctr) {
  const grup = {};
  asis.forEach(function (row, i) {
    if (!row.low) return;
    const k = row.tgl + '|' + row.shift + '|' + row.dokter;
    (grup[k] || (grup[k] = [])).push(i);
  });
  const kat = function (low) { return pMap[low] ? pMap[low].kategori : 'Existing'; };
  const jmlNew = function (gk) { return (grup[gk] || []).filter(function (i) { return kat(asis[i].low) === 'New'; }).length; };
  // indeks kehadiran per shift & beban per orang (utk Opsi B)
  const inShift = {}, load = {};
  asis.forEach(function (row) { if (!row.low) return; inShift[row.low + '|' + row.tgl + '|' + row.shift] = true; load[row.low] = (load[row.low] || 0) + 1; });
  const existingAktif = Object.keys(pMap).filter(function (k) { return pMap[k].aktif !== 'Tidak' && pMap[k].kategori === 'Existing'; });

  Object.keys(grup).forEach(function (gk) {
    let guard = 0;
    while (jmlNew(gk) >= 2 && guard++ < 20) {
      const parts = gk.split('|'); const tgl = parts[0], shift = parts[1];
      const idxNew = grup[gk].filter(function (i) { return kat(asis[i].low) === 'New' && !touched['A' + asis[i].r]; });
      if (!idxNew.length) { gagal.push({ kode: 'A-02', ket: 'Grup ' + gk + ' — slot New sudah dipakai saran lain' }); break; }
      // keluarkan New dgn beban TERBESAR duluan (rebalance yang over-load)
      idxNew.sort(function (a, b) { return (load[asis[b].low] || 0) - (load[asis[a].low] || 0); });
      const X = idxNew[0];
      const N = pMap[asis[X].low];
      let done = false;

      // --- Opsi A: TUKAR dgn Existing di grup lain, shift sama (beban tetap) ---
      const kandidatGrup = Object.keys(grup).filter(function (g2) {
        const p2 = g2.split('|'); return p2[0] === tgl && p2[1] === shift && g2 !== gk && jmlNew(g2) === 0;
      });
      for (let a = 0; a < kandidatGrup.length && !done; a++) {
        const anggota = grup[kandidatGrup[a]];
        for (let b = 0; b < anggota.length; b++) {
          const Y = anggota[b]; const E = pMap[asis[Y].low];
          if (!E || E.kategori !== 'Existing') continue;
          if (touched['A' + asis[Y].r]) continue;
          if (!sb_bisa_(N.kapabilitasSet, asis[Y].spes)) continue;   // New → dokter partner
          if (!sb_bisa_(E.kapabilitasSet, asis[X].spes)) continue;   // Existing → dokter grup ini
          items.push({
            id: 'i' + (++ctr.n), kat: 'A-02', tgl: tgl, shift: shift,
            ket: E.nama + ' ⇄ ' + N.nama,
            detail: E.nama + ' → ' + asis[X].dokter + '  •  ' + N.nama + ' → ' + asis[Y].dokter + ' (tukar, beban tetap)',
            moves: [
              { b: 'A', r: asis[X].r, c: PA_KOL_PERAWAT, dari: asis[X].asal, ke: E.nama },
              { b: 'A', r: asis[Y].r, c: PA_KOL_PERAWAT, dari: asis[Y].asal, ke: N.nama }
            ]
          });
          const lowX = asis[X].low, lowY = asis[Y].low;
          asis[X].low = lowY; asis[Y].low = lowX;                   // keduanya tetap di shift → inShift/load tak berubah
          touched['A' + asis[X].r] = true; touched['A' + asis[Y].r] = true;
          done = true; break;
        }
      }
      if (done) continue;

      // --- Opsi B: masukkan Existing LUANG di shift ini, gantikan si New → grup jadi CAMPUR ---
      const spesX = asis[X].spes;
      const kandE = existingAktif.filter(function (low) {
        if (inShift[low + '|' + tgl + '|' + shift]) return false;                              // belum ada di shift ini
        if (liburMap[pMap[low].nama] && liburMap[pMap[low].nama][tgl]) return false;           // tidak libur
        if (!sb_bisa_(pMap[low].kapabilitasSet, spesX)) return false;                          // mampu spesialis grup
        return true;
      }).sort(function (a, b) { return (load[a] || 0) - (load[b] || 0); });                    // Existing paling ringan
      if (kandE.length && ((load[asis[X].low] || 0) - 1) >= 1) {                               // New tetap ≥1 slot
        const Elow = kandE[0], E = pMap[Elow], lowN = asis[X].low;
        items.push({
          id: 'i' + (++ctr.n), kat: 'A-02', tgl: tgl, shift: shift,
          ket: E.nama + ' menggantikan ' + N.nama,
          detail: asis[X].dokter + ': ' + E.nama + ' (Existing) masuk, ' + N.nama + ' (New) keluar → grup jadi campur',
          moves: [{ b: 'A', r: asis[X].r, c: PA_KOL_PERAWAT, dari: asis[X].asal, ke: E.nama }]
        });
        inShift[lowN + '|' + tgl + '|' + shift] = false; load[lowN] = (load[lowN] || 1) - 1;
        asis[X].low = Elow; inShift[Elow + '|' + tgl + '|' + shift] = true; load[Elow] = (load[Elow] || 0) + 1;
        touched['A' + asis[X].r] = true;
        done = true;
      }
      if (done) continue;

      gagal.push({ kode: 'A-02', ket: 'Grup ' + gk + ' — tak ada Existing untuk ditukar (shift sama) maupun Existing luang di shift ini yang berkapabilitas' });
      break;
    }
  });
}

/**
 * PASS longshift asistensi (A-04): perawat longshift < target → tambahkan ke
 * shift pelengkap pada hari yang sudah ia kerja, mengambil slot donor AMAN.
 * Menghasilkan ITEM {moves:[1]}.
 */
function sb_passLongshiftAsis_(asis, pMap, target, items, touched, gagal, ctr) {
  const aktif = Object.keys(pMap).filter(function (k) { return pMap[k].aktif !== 'Tidak'; });
  const recompute = function () { return sb_statLongshift_(asis.map(function (r) { return { tgl: r.tgl, shift: r.shift, low: r.low }; })); };
  const kat = function (low) { return pMap[low] ? pMap[low].kategori : 'Existing'; };
  const grpNewCount = function (tgl, shift, dokter, kecualiIdx, tambahLow) {
    let n = 0;
    asis.forEach(function (row, i) { if (row.tgl === tgl && row.shift === shift && row.dokter === dokter && i !== kecualiIdx && kat(row.low) === 'New') n++; });
    if (tambahLow && kat(tambahLow) === 'New') n++;
    return n;
  };
  const punyaDiShift = function (low, tgl, shift) { return asis.some(function (r) { return r.low === low && r.tgl === tgl && r.shift === shift; }); };

  aktif.forEach(function (P) {
    const pu = pMap[P]; let stat = recompute(); let guard = 0;
    while ((stat[P] ? stat[P].longshift : 0) < target && guard++ < 12) {
      const st = stat[P]; if (!st) break; let dibuat = false;
      const hariList = Object.keys(st.hari);
      for (let h = 0; h < hariList.length && !dibuat; h++) {
        const D = hariList[h]; const sh = st.hari[D];
        if (!!sh['Pagi'] === !!sh['Siang']) continue;
        const target2 = sh['Pagi'] ? 'Siang' : 'Pagi';
        // dokter yang diasisteni P di shift lain hari itu → utamakan longshift dgn dokter BEDA (soft)
        let pDok = '';
        for (let k = 0; k < asis.length; k++) { if (asis[k].low === P && asis[k].tgl === D && asis[k].shift !== target2) { pDok = asis[k].dokter; break; } }
        const donorValid = function (i) {
          const row = asis[i];
          if (row.tgl !== D || row.shift !== target2 || !row.low) return false;
          if (touched['A' + row.r]) return false;
          const Q = row.low; if (Q === P) return false;
          if (!sb_bisa_(pu.kapabilitasSet, row.spes)) return false;
          if (punyaDiShift(P, D, target2)) return false;
          if (pu.kategori === 'New' && grpNewCount(row.tgl, row.shift, row.dokter, i, P) >= 2) return false; // A-02
          const sq = stat[Q]; if (!sq) return false;
          const qLoss = (sq.hari[D] && sq.hari[D]['Pagi'] && sq.hari[D]['Siang']) ? 1 : 0;
          if ((sq.longshift - qLoss) < target) return false; // jangan bikin A-04 baru utk donor
          if ((sq.total - 1) < 1) return false;              // donor tetap terjadwal
          return true;
        };
        const pasang = function (i) {
          const row = asis[i], Q = row.low;
          items.push({
            id: 'i' + (++ctr.n), kat: 'A-04', tgl: D, shift: target2, dokter: row.dokter,
            ke: pu.nama, dari: pMap[Q] ? pMap[Q].nama : row.asal,
            ket: pu.nama + ' menggantikan ' + (pMap[Q] ? pMap[Q].nama : row.asal) + (row.dokter === pDok ? ' (dokter sama)' : ''),
            moves: [{ b: 'A', r: row.r, c: PA_KOL_PERAWAT, dari: row.asal, ke: pu.nama }]
          });
          row.low = P; touched['A' + row.r] = true; dibuat = true; stat = recompute();
        };
        let fb = -1;
        for (let i = 0; i < asis.length && !dibuat; i++) {
          if (!donorValid(i)) continue;
          if (asis[i].dokter !== pDok) pasang(i);            // utamakan dokter beda
          else if (fb < 0) fb = i;                            // simpan fallback dokter sama
        }
        if (!dibuat && fb >= 0) pasang(fb);                  // soft: pakai dokter sama bila tak ada beda
      }
      if (!dibuat) { gagal.push({ kode: 'A-04', ket: pu.nama + ' — longshift ' + (st ? st.longshift : 0) + '/' + target + ', tak ada donor aman' }); break; }
    }
  });
}

/** PASS longshift resepsionis (R-07, Silver): serupa, tanpa kapabilitas/A-02. */
function sb_passLongshiftRcp_(slots, rcpMap, target, items, touched, gagal, ctr) {
  const silver = Object.keys(rcpMap).filter(function (k) { return rcpMap[k].level === 'Silver' && rcpMap[k].aktif !== 'Tidak'; });
  const recompute = function () { return sb_statLongshift_(slots.map(function (s) { return { tgl: s.tgl, shift: s.shift, low: s.low }; })); };
  const punyaDiShift = function (low, tgl, shift) { return slots.some(function (s) { return s.low === low && s.tgl === tgl && s.shift === shift; }); };

  silver.forEach(function (P) {
    let stat = recompute(); let guard = 0;
    while ((stat[P] ? stat[P].longshift : 0) < target && guard++ < 12) {
      const st = stat[P]; if (!st) break; let dibuat = false;
      const hariList = Object.keys(st.hari);
      for (let h = 0; h < hariList.length && !dibuat; h++) {
        const D = hariList[h]; const sh = st.hari[D];
        if (!!sh['Pagi'] === !!sh['Siang']) continue;
        const target2 = sh['Pagi'] ? 'Siang' : 'Pagi';
        for (let i = 0; i < slots.length && !dibuat; i++) {
          const s = slots[i];
          if (s.tgl !== D || s.shift !== target2 || !s.low) continue;
          if (touched['R' + s.r + 'c' + s.c]) continue;
          const Q = s.low; if (Q === P) continue;
          if (punyaDiShift(P, D, target2)) continue;
          const sq = stat[Q]; if (!sq) continue;
          const qLoss = (sq.hari[D] && sq.hari[D]['Pagi'] && sq.hari[D]['Siang']) ? 1 : 0;
          const qNeed = (rcpMap[Q] && rcpMap[Q].level === 'Silver') ? target : 0;
          if ((sq.longshift - qLoss) < qNeed) continue;
          if ((sq.total - 1) < 1) continue;
          items.push({
            id: 'i' + (++ctr.n), kat: 'R-07', tgl: D, shift: target2,
            ke: rcpMap[P].nama, dari: rcpMap[Q] ? rcpMap[Q].nama : s.asal,
            ket: rcpMap[P].nama + ' menggantikan ' + (rcpMap[Q] ? rcpMap[Q].nama : s.asal),
            moves: [{ b: 'R', r: s.r, c: s.c, dari: s.asal, ke: rcpMap[P].nama }]
          });
          s.low = P; touched['R' + s.r + 'c' + s.c] = true; dibuat = true; stat = recompute();
        }
      }
      if (!dibuat) { gagal.push({ kode: 'R-07', ket: rcpMap[P].nama + ' — longshift ' + (st ? st.longshift : 0) + '/' + target + ', tak ada donor aman' }); break; }
    }
  });
}

/**
 * ENTRY — Sarankan perbaikan blocker (TIDAK menulis). Mengembalikan daftar
 * item (tiap item punya 'moves' sendiri) untuk diterapkan per-item.
 */
function sarankanPerbaikanBlockerWeb() {
  wajibPeran_(['Admin', 'SPV']);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pMap = petaPerawat_();
  const liburMap = petaLiburPerNama_(ss);
  const rcpMap = {}; getMasterData('receptionist').forEach(function (r) { rcpMap[r.nama.toLowerCase()] = { nama: r.nama, level: r.level, aktif: r.aktif }; });
  const longPrw = Number(getConfig('LONGSHIFT_WAJIB_PERAWAT')) || 4;
  const longRcp = Number(getConfig('LONGSHIFT_WAJIB_RECEPT_SILVER')) || 4;

  const mA = sb_modelAsis_(ss), mR = sb_modelRcp_(ss);
  const items = [], gagal = [], touched = {}, ctr = { n: 0 };

  try { sb_passA02_(mA.rows, pMap, liburMap, items, touched, gagal, ctr); } catch (e) { gagal.push({ kode: 'A-02', ket: 'gagal proses: ' + e.message }); }
  try { sb_passLongshiftAsis_(mA.rows, pMap, longPrw, items, touched, gagal, ctr); } catch (e) { gagal.push({ kode: 'A-04', ket: 'gagal proses: ' + e.message }); }
  try { sb_passLongshiftRcp_(mR.slots, rcpMap, longRcp, items, touched, gagal, ctr); } catch (e) { gagal.push({ kode: 'R-07', ket: 'gagal proses: ' + e.message }); }

  const nA02 = items.filter(function (x) { return x.kat === 'A-02'; }).length;
  const nA04 = items.filter(function (x) { return x.kat === 'A-04'; }).length;
  const nR07 = items.filter(function (x) { return x.kat === 'R-07'; }).length;
  const ada = items.length > 0;
  const pesan = ada
    ? ('Ditemukan ' + items.length + ' saran (A-02: ' + nA02 + ', longshift perawat: ' + nA04 + ', longshift resepsionis: ' + nR07 + '). Tinjau, lalu tekan ikon ✅ pada baris yang ingin diterapkan. Tidak ada yang ditulis sampai kamu menekannya.')
    : 'Tidak ada saran perbaikan berbasis pertukaran yang aman ditemukan (mungkin sudah sesuai, atau butuh Generate ulang).';
  return { ok: true, ada: ada, items: items, takTertangani: gagal, pesan: pesan };
}

/**
 * ENTRY — Terapkan daftar moves (dari SATU item, atau beberapa). Menulis HANYA
 * bila sel saat ini masih == 'dari' (guard anti-timpa).
 */
function terapkanSaranBlockerWeb(moves) {
  wajibPeran_(['Admin', 'SPV']);
  if (!moves || !moves.length) return { ok: false, pesan: 'Tak ada saran untuk diterapkan.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pa = ss.getSheetByName(SHEETS.PAPAN_ASISTENSI);
  const pr = ss.getSheetByName(SHEETS.PAPAN_RECEPTIONIST);
  refreshAsistensiValidation_(pa);
  refreshResepsionisValidation_(pr);
  let tulis = 0, lewat = 0; const bentrok = [];
  moves.forEach(function (m) {
    const sh = m.b === 'A' ? pa : pr; if (!sh) { lewat++; return; }
    const sel = sh.getRange(m.r, m.c);
    const kini = String(sel.getValue() || '').trim();
    if (kini !== String(m.dari || '').trim()) { lewat++; bentrok.push((m.ke || '(kosong)') + ' @ ' + (m.tgl || '') + ' ' + (m.shift || '')); return; }
    sel.setNumberFormat('@').setValue(m.ke || '');
    tulis++;
  });
  SpreadsheetApp.flush();
  const ok = tulis > 0;
  return { ok: ok, tulis: tulis, lewat: lewat, bentrok: bentrok,
    pesan: ok ? ('Diterapkan: ' + tulis + ' sel' + (lewat ? ', ' + lewat + ' dilewati (sel sudah berubah)' : '') + '.')
      : 'Gagal: semua sel sasaran sudah berubah sejak saran dibuat. Tekan "Sarankan perbaikan" lagi untuk memperbarui.' };
}