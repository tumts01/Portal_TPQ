/**
 * PRESENSI MODEL BARU (HEMAT BARIS)
 * ---------------------------------------------------------------
 * Model lama  : Presensi_Santri -> 1 baris per santri per hari (termasuk yang hadir).
 * Model baru  : 2 sheet kecil
 *   - Pertemuan               -> 1 baris per kelas per hari (pengajar, jurnal, daftar santri
 *                                yang terdaftar di kelas saat itu).
 *   - Presensi_Ketidakhadiran -> hanya santri yang S / I / A.
 *   Hadir (H) = jumlah pertemuan santri tsb - (S + I + A).
 *
 * Peralihan aman lewat dua langkah:
 *   1. Mode LAMA (default)  : aplikasi tetap memakai Presensi_Santri. Migrasi bisa disimulasikan
 *                             dan dijalankan tanpa mengubah/menghapus data lama.
 *   2. Mode BARU            : setelah migrasi diverifikasi, aktifkan lewat presensiAktifkanModeBaru().
 *                             Bisa dikembalikan lewat presensiKembaliKeModeLama().
 *
 * Fungsi berakhiran "_" bersifat privat (tidak bisa dipanggil dari browser).
 */

const PRES_SH_PERTEMUAN = 'Pertemuan';
const PRES_SH_ABSEN = 'Presensi_Ketidakhadiran';
const PRES_HDR_PERTEMUAN = ['ID_Pertemuan', 'Tanggal', 'Kelas', 'Tahun_Ajaran', 'Pengajar', 'Jurnal', 'Jenis', 'Jml_Santri', 'Jml_Tidak_Hadir', 'Daftar_Santri', 'Created_At'];
const PRES_HDR_ABSEN = ['ID_Absen', 'ID_Pertemuan', 'Tanggal', 'NIS', 'Nama_Santri', 'Kelas', 'Status', 'Tahun_Ajaran', 'Created_At'];
// Kolom bertipe teks (agar tanggal / NIS tidak diubah otomatis oleh Sheets)
const PRES_TEXT_COLS_PERTEMUAN = [2, 10];
const PRES_TEXT_COLS_ABSEN = [3, 4];
const PRES_KEGIATAN_KHUSUS_ = ['RAPAT', 'ISTIGHOSAH', 'LAIN-LAIN'];

// ======================= HELPER DASAR =======================

function presensiModeBaru_() {
  return PropertiesService.getScriptProperties().getProperty('PRESENSI_MODE') === 'BARU';
}

function presensiTz_() {
  try { return SS.getSpreadsheetTimeZone() || 'GMT+7'; } catch (e) { return 'GMT+7'; }
}

function presensiPastikanAdmin_() {
  if (PropertiesService.getScriptProperties().getProperty('PRESENSI_ADMIN') !== 'YA') {
    throw new Error('Ditolak. Buka Apps Script > Project Settings > Script properties, tambahkan PRESENSI_ADMIN = YA, lalu jalankan lagi. Hapus properti itu setelah selesai.');
  }
}

function presensiTglStr_(v) {
  if (v instanceof Date) {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, presensiTz_(), 'yyyy-MM-dd');
  }
  const s = String(v == null ? '' : v).trim();
  if (!s) return '';
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (m) return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  const d = new Date(s);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, presensiTz_(), 'yyyy-MM-dd');
  return s;
}

function presensiNormStatus_(raw) {
  const st = String(raw == null || raw === '' ? 'H' : raw).toUpperCase().trim();
  if (st === 'S' || st.indexOf('SAKIT') !== -1) return 'S';
  if (st === 'I' || st.indexOf('IZIN') !== -1) return 'I';
  if (st === 'A' || st.indexOf('ALPA') !== -1 || st.indexOf('ALFA') !== -1) return 'A';
  return 'H';
}

// Identitas santri di daftar: NIS bila ada, jika tidak "#nama huruf kecil"
function presensiToken_(nis, nama) {
  const n = String(nis == null ? '' : nis).trim();
  if (n && n !== '-') return n;
  const nm = String(nama == null ? '' : nama).trim().toLowerCase();
  return nm ? '#' + nm : '';
}

function presensiTaCocok_(rowTa, targetTa) {
  const t = String(targetTa == null ? '' : targetTa).trim();
  const r = String(rowTa == null ? '' : rowTa).trim();
  return !t || !r || r === t;
}

function presensiSheet_(name, headers, textCols) {
  let sh = SS.getSheetByName(name);
  if (!sh) sh = SS.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.setFrozenRows(1);
    textCols.forEach(function (c) {
      sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@');
    });
  }
  return sh;
}

function presensiSheetPertemuan_() { return presensiSheet_(PRES_SH_PERTEMUAN, PRES_HDR_PERTEMUAN, PRES_TEXT_COLS_PERTEMUAN); }
function presensiSheetAbsen_() { return presensiSheet_(PRES_SH_ABSEN, PRES_HDR_ABSEN, PRES_TEXT_COLS_ABSEN); }

function presensiBaca_(sh, ncol) {
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, ncol).getValues();
}

function presensiTulis_(sh, startRow, rows, ncol, textCols) {
  const CHUNK = 5000;
  for (let i = 0; i < rows.length; i += CHUNK) {
    const part = rows.slice(i, i + CHUNK);
    textCols.forEach(function (c) {
      sh.getRange(startRow + i, c, part.length, 1).setNumberFormat('@');
    });
    sh.getRange(startRow + i, 1, part.length, ncol).setValues(part);
  }
}

// ======================= PENYIMPANAN =======================

function presensiSimpanBatch_(dataArray) {
  if (!dataArray || dataArray.length === 0) return { success: true, message: 'Tidak ada data untuk disimpan' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const first = dataArray[0];
    const kelas = String(first.kelas || '').trim();
    if (!kelas) throw new Error('Kelas wajib diisi');
    const tanggal = presensiTglStr_(first.tanggal);
    if (!tanggal) throw new Error('Tanggal tidak valid');
    const kelasU = kelas.toUpperCase();
    const khusus = PRES_KEGIATAN_KHUSUS_.indexOf(kelasU) !== -1;

    const shP = presensiSheetPertemuan_();
    const shA = presensiSheetAbsen_();

    // Cegah absen ganda (kecuali kegiatan khusus yang boleh berkali-kali)
    if (!khusus) {
      const last = shP.getLastRow();
      if (last > 1) {
        const kolom = shP.getRange(2, 2, last - 1, 2).getValues(); // Tanggal, Kelas
        for (let i = 0; i < kolom.length; i++) {
          if (presensiTglStr_(kolom[i][0]) === tanggal && String(kolom[i][1]).trim().toUpperCase() === kelasU) {
            throw new Error('GAGAL: Kelas ' + kelas + ' sudah diabsen pada tanggal tersebut! Tidak boleh dobel.');
          }
        }
      }
    }

    const roster = [];
    const absen = [];
    const seen = {};
    dataArray.forEach(function (d) {
      const token = presensiToken_(d.nis, d.nama_santri);
      if (!token || seen[token]) return;
      seen[token] = true;
      roster.push(token);
      const st = presensiNormStatus_(d.status);
      if (st !== 'H') {
        const nisRaw = String(d.nis == null ? '' : d.nis).trim();
        absen.push({ nis: nisRaw === '-' ? '' : nisRaw, nama: String(d.nama_santri || '').trim(), status: st });
      }
    });

    const now = new Date();
    const stamp = Utilities.formatDate(now, presensiTz_(), 'yyyyMMddHHmmss');
    const idP = 'PT' + stamp + '-' + Utilities.getUuid().slice(0, 4);
    const ta = String(first.tahun_ajaran || '').trim();

    const rowP = [
      idP, tanggal, kelas, ta,
      String(first.input_oleh || '').trim(),
      String(first.keterangan || '').trim(),
      khusus ? 'KEGIATAN' : 'KELAS',
      roster.length, absen.length, roster.join('|'), now
    ];
    const rowsA = absen.map(function (a, i) {
      return ['AB' + stamp + '-' + (i + 1) + '-' + Utilities.getUuid().slice(0, 4), idP, tanggal, a.nis, a.nama, kelas, a.status, ta, now];
    });

    const rP = shP.getLastRow() + 1;
    presensiTulis_(shP, rP, [rowP], PRES_HDR_PERTEMUAN.length, PRES_TEXT_COLS_PERTEMUAN);
    if (rowsA.length > 0) {
      try {
        presensiTulis_(shA, shA.getLastRow() + 1, rowsA, PRES_HDR_ABSEN.length, PRES_TEXT_COLS_ABSEN);
      } catch (err) {
        shP.deleteRow(rP); // batalkan agar tidak ada pertemuan tanpa data ketidakhadiran
        throw err;
      }
    }
    SpreadsheetApp.flush();

    return {
      success: true,
      message: roster.length > 0
        ? 'Presensi ' + roster.length + ' santri berhasil disimpan (' + absen.length + ' tidak hadir)'
        : 'Jurnal kegiatan berhasil disimpan'
    };
  } finally {
    lock.releaseLock();
  }
}

// ======================= PEMBACAAN =======================

/**
 * Hitung total pertemuan & S/I/A per santri (token) untuk filter kelas / tahun ajaran.
 */
function presensiHitung_(pRows, aRows, opt) {
  const res = {};
  const kelasCocok = function (k) { return !opt.kelas || isKelasMatch_(k, opt.kelas); };
  pRows.forEach(function (r) {
    if (!presensiTaCocok_(r[3], opt.ta) || !kelasCocok(r[2])) return;
    const roster = String(r[9] || '');
    if (!roster) return;
    roster.split('|').forEach(function (tok) {
      if (!tok) return;
      if (!res[tok]) res[tok] = { token: tok, nama: '', kelas: String(r[2] || '').trim(), total: 0, S: 0, I: 0, A: 0 };
      res[tok].total++;
    });
  });
  aRows.forEach(function (r) {
    if (!presensiTaCocok_(r[7], opt.ta) || !kelasCocok(r[5])) return;
    const tok = presensiToken_(r[3], r[4]);
    if (!tok || !res[tok]) return;
    const st = presensiNormStatus_(r[6]);
    if (st !== 'H') res[tok][st]++;
    if (!res[tok].nama && r[4]) res[tok].nama = String(r[4]).trim();
  });
  return res;
}

function getRekapPresensiSantriKelas(kelas, tahunAjaran) {
  if (!presensiModeBaru_()) return presensiRekapLegacy_(kelas, tahunAjaran);

  const pRows = presensiBaca_(presensiSheetPertemuan_(), PRES_HDR_PERTEMUAN.length);
  const aRows = presensiBaca_(presensiSheetAbsen_(), PRES_HDR_ABSEN.length);
  const hasil = presensiHitung_(pRows, aRows, { kelas: kelas, ta: tahunAjaran });

  const namaMap = {};
  try {
    (getSantriList() || []).forEach(function (s) {
      const nm = String(s.Nama_Santri || s.nama_santri || s.Nama || '').trim();
      const tk = presensiToken_(s.NIS || s.nis, nm);
      if (tk) namaMap[tk] = nm;
      if (nm) namaMap['#' + nm.toLowerCase()] = nm;
    });
  } catch (e) {}

  return Object.keys(hasil).map(function (tok) {
    const h = hasil[tok];
    const isNameToken = tok.charAt(0) === '#';
    return {
      nis: isNameToken ? '-' : tok,
      nama_santri: namaMap[tok] || h.nama || (isNameToken ? tok.slice(1) : tok),
      kelas: h.kelas,
      H: Math.max(0, h.total - h.S - h.I - h.A),
      S: h.S, I: h.I, A: h.A
    };
  }).sort(function (a, b) { return a.nama_santri.localeCompare(b.nama_santri); });
}

// Padanan model lama (Presensi_Santri) dengan bentuk keluaran yang sama
function presensiRekapLegacy_(kelas, ta) {
  const rows = getPresensiSantriByTA(kelas, ta) || [];
  const g = {};
  rows.forEach(function (d) {
    const nisOk = d.nis && d.nis !== '-';
    const namaOk = d.nama_santri && d.nama_santri !== '(Tanpa Nama)';
    if (!nisOk && !namaOk) return;
    const key = nisOk ? d.nis : (d.nama_santri + '_' + d.kelas);
    if (!g[key]) g[key] = { nis: nisOk ? d.nis : '-', nama_santri: d.nama_santri, kelas: d.kelas, H: 0, S: 0, I: 0, A: 0 };
    g[key][presensiNormStatus_(d.status)]++;
  });
  return Object.keys(g).map(function (k) { return g[k]; });
}

function presensiRekapByNis_(nis, namaSantri, tahunAjaran) {
  const pRows = presensiBaca_(presensiSheetPertemuan_(), PRES_HDR_PERTEMUAN.length);
  const aRows = presensiBaca_(presensiSheetAbsen_(), PRES_HDR_ABSEN.length);
  const tNis = String(nis || '').trim();
  const tNama = String(namaSantri || '').trim().toLowerCase();
  const tokNis = presensiToken_(tNis, '');
  const tokNama = tNama ? '#' + tNama : '';
  const targetTA = tahunAjaran ? String(tahunAjaran).trim() : '';

  function hitung(taFilter) {
    let total = 0, s = 0, iz = 0, a = 0;
    pRows.forEach(function (r) {
      if (!presensiTaCocok_(r[3], taFilter)) return;
      const roster = '|' + String(r[9] || '') + '|';
      if ((tokNis && roster.indexOf('|' + tokNis + '|') !== -1) || (tokNama && roster.indexOf('|' + tokNama + '|') !== -1)) total++;
    });
    aRows.forEach(function (r) {
      if (!presensiTaCocok_(r[7], taFilter)) return;
      const rn = String(r[3] || '').trim();
      const rm = String(r[4] || '').trim().toLowerCase();
      if (!((tNis && rn === tNis) || (tNama && rm === tNama))) return;
      const st = presensiNormStatus_(r[6]);
      if (st === 'S') s++; else if (st === 'I') iz++; else if (st === 'A') a++;
    });
    return { hadir: Math.max(0, total - s - iz - a), sakit: s, izin: iz, alpa: a, total: total };
  }

  let res = targetTA ? hitung(targetTA) : hitung('');
  if (res.total === 0 && targetTA) {
    const all = hitung('');
    if (all.total > 0) res = all;
  }
  const siaTotal = res.sakit + res.izin + res.alpa;
  return {
    success: true,
    nis: tNis,
    namaSantri: namaSantri,
    tahunAjaran: targetTA,
    hadir: res.hadir,
    sakit: res.sakit,
    izin: res.izin,
    alpa: res.alpa,
    totalSIA: siaTotal,
    totalPertemuan: res.total,
    persenKehadiran: res.total > 0 ? Math.round((res.hadir / res.total) * 100) : 100
  };
}

function presensiAbsenUntukRapor_(kelas, targetTA, santriNisSet, santriNamaSet) {
  const aRows = presensiBaca_(presensiSheetAbsen_(), PRES_HDR_ABSEN.length);
  const list = [];
  aRows.forEach(function (r) {
    const rNis = String(r[3] || '').trim();
    const rNama = String(r[4] || '').trim().toLowerCase();
    const rKelas = String(r[5] || '').trim();
    const rTA = String(r[7] || '').trim();
    const matchTA = !targetTA || !rTA || rTA === targetTA;
    const matchKelas = !kelas || isKelasMatch_(rKelas, kelas) || (rNis && santriNisSet.has(rNis)) || (rNama && santriNamaSet.has(rNama));
    if (matchTA && matchKelas) {
      list.push({
        nis: rNis || '-',
        nama_santri: r[4] || '',
        kelas: rKelas || '-',
        status: presensiNormStatus_(r[6]),
        tahun_ajaran: rTA
      });
    }
  });
  return list;
}

function presensiGuruByGuru_(namaGuru, bulan, tahun) {
  const pRows = presensiBaca_(presensiSheetPertemuan_(), PRES_HDR_PERTEMUAN.length);
  const result = [];
  const unik = {};
  pRows.forEach(function (r, i) {
    if (!r[0]) return;
    const guru = String(r[4] || '').trim();
    if (namaGuru && guru !== namaGuru) return;
    const m = presensiTglStr_(r[1]).match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return;
    if (bulan && parseInt(m[2], 10) !== parseInt(bulan, 10)) return;
    if (tahun && parseInt(m[1], 10) !== parseInt(tahun, 10)) return;
    const kelas = String(r[2] || '').trim();
    let key = m[0] + '_' + guru + '_' + kelas;
    if (PRES_KEGIATAN_KHUSUS_.indexOf(kelas.toUpperCase()) !== -1) key += '_' + i;
    if (unik[key]) return;
    unik[key] = true;
    result.push({
      tanggal: m[0],
      nama_guru: guru,
      kelas_kegiatan: kelas,
      jurnal_mengajar: String(r[5] || '').trim() || '-'
    });
  });
  return { total: result.length, data: result };
}

function presensiDateVal_(tglStr) {
  if (!tglStr) return 1;
  const t = new Date(tglStr).getTime();
  return isNaN(t) ? 1 : t;
}

function presensiJurnalKelas_() {
  const pRows = presensiBaca_(presensiSheetPertemuan_(), PRES_HDR_PERTEMUAN.length);
  const latest = {};
  pRows.forEach(function (r) {
    const kelas = String(r[2] || '').trim();
    const kU = kelas.toUpperCase();
    if (!kelas || PRES_KEGIATAN_KHUSUS_.indexOf(kU) !== -1) return;
    const jurnal = String(r[5] || '').trim();
    if (!jurnal) return;
    const tgl = presensiTglStr_(r[1]);
    const dv = presensiDateVal_(tgl);
    if (!latest[kU] || latest[kU].dateVal < dv) {
      latest[kU] = { kelas: kelas, materi: jurnal, pengajar: String(r[4] || '').trim(), tanggal: tgl, dateVal: dv };
    }
  });
  return Object.keys(latest).map(function (k) { return latest[k]; })
    .sort(function (a, b) { return a.kelas.localeCompare(b.kelas); });
}

// Dipakai getRekapDataKelas: { namaKelas: {dateVal, jurnal} }
function presensiJurnalTerakhirMap_() {
  const pRows = presensiBaca_(presensiSheetPertemuan_(), PRES_HDR_PERTEMUAN.length);
  const map = {};
  pRows.forEach(function (r) {
    const kelas = String(r[2] || '').trim();
    const jurnal = String(r[5] || '').trim();
    if (!kelas || !jurnal) return;
    const dv = presensiDateVal_(presensiTglStr_(r[1]));
    if (!map[kelas] || map[kelas].dateVal <= dv) map[kelas] = { dateVal: dv, jurnal: jurnal };
  });
  return map;
}

function presensiKelasUnik_() {
  const pRows = presensiBaca_(presensiSheetPertemuan_(), PRES_HDR_PERTEMUAN.length);
  const unik = {};
  pRows.forEach(function (r) {
    const kelas = String(r[2] || '').trim();
    if (kelas && String(r[9] || '')) unik[kelas] = true;
  });
  return Object.keys(unik).sort();
}

// ======================= MIGRASI =======================

/**
 * Menyusun rencana migrasi dari Presensi_Santri (tidak menulis apa pun).
 * Satu "pengiriman" absen (kelas + tanggal + menit input) menjadi 1 baris Pertemuan,
 * sehingga hasil hitung per santri identik dengan model lama.
 */
function presensiRencanaMigrasi_() {
  const sh = SHEETS.presensiSantri;
  const data = sh ? sh.getDataRange().getValues() : [];
  const groups = [];
  const index = {};
  const legacyAgg = {};
  let dilewati = 0, barisLegacy = 0, tanpaTanggal = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const kelas = String(row[4] || '').trim();
    const nisRaw = String(row[2] || '').trim();
    const nis = (nisRaw && nisRaw !== '-') ? nisRaw : '';
    const nama = String(row[3] || '').trim();
    const token = presensiToken_(nis, nama);
    // Baris tanpa kelas DAN tanpa identitas santri = sampah, dilewati.
    // Baris tanpa kelas tetapi punya santri tetap dimigrasi agar total per santri (rekap/rapor) tidak berubah.
    if (!kelas && !token) { dilewati++; continue; }
    barisLegacy++;

    const tgl = presensiTglStr_(row[1]);
    if (!tgl) tanpaTanggal++;
    const ta = String(row[9] || '').trim();
    const kelasU = kelas.toUpperCase();
    const khusus = PRES_KEGIATAN_KHUSUS_.indexOf(kelasU) !== -1;
    const created = row[8] instanceof Date ? row[8] : null;
    const bucket = created ? Math.floor(created.getTime() / 60000) : 'x';
    const key = khusus ? ('K|' + i) : ('R|' + tgl + '|' + kelasU + '|' + ta + '|' + bucket);

    let g = index[key];
    if (!g) {
      g = { tgl: tgl, kelas: kelas, ta: ta, pengajar: '', jurnal: '', khusus: khusus, created: created, tokens: {}, order: [] };
      index[key] = g;
      groups.push(g);
    }
    const pengajar = String(row[7] || '').trim();
    if (!g.pengajar && pengajar) g.pengajar = pengajar;
    const jurnal = String(row[6] || '').trim();
    if (!g.jurnal && jurnal) g.jurnal = jurnal;

    if (token) {
      const st = presensiNormStatus_(row[5]);
      const ex = g.tokens[token];
      if (!ex) { g.tokens[token] = { nis: nis, nama: nama, status: st }; g.order.push(token); }
      else if (ex.status === 'H' && st !== 'H') ex.status = st;
      const ak = token + '||' + ta;
      if (!legacyAgg[ak]) legacyAgg[ak] = { H: 0, S: 0, I: 0, A: 0 };
      legacyAgg[ak][st]++;
    }
  }

  const pRows = [];
  const aRows = [];
  const hariKelas = {};
  let tanpaPengajar = 0, kegiatan = 0;
  groups.forEach(function (g, n) {
    const idP = 'PT-MIG-' + (n + 1);
    const absen = g.order.filter(function (t) { return g.tokens[t].status !== 'H'; });
    const dibuat = g.created || new Date();
    pRows.push([idP, g.tgl, g.kelas, g.ta, g.pengajar, g.jurnal, g.khusus ? 'KEGIATAN' : 'KELAS',
      g.order.length, absen.length, g.order.join('|'), dibuat]);
    absen.forEach(function (t, k) {
      const e = g.tokens[t];
      aRows.push(['AB-MIG-' + (n + 1) + '-' + (k + 1), idP, g.tgl, e.nis, e.nama, g.kelas, e.status, g.ta, dibuat]);
    });
    if (g.khusus) kegiatan++;
    else {
      const hk = g.tgl + '|' + g.kelas.toUpperCase();
      hariKelas[hk] = (hariKelas[hk] || 0) + 1;
    }
    if (!g.pengajar) tanpaPengajar++;
  });

  const duplikat = Object.keys(hariKelas).filter(function (k) { return hariKelas[k] > 1; });
  return {
    pRows: pRows, aRows: aRows, legacyAgg: legacyAgg,
    ringkasan: {
      baris_lama_valid: barisLegacy,
      baris_lama_dilewati_kosong: dilewati,
      baris_lama_tanpa_tanggal: tanpaTanggal,
      pertemuan_baru: pRows.length,
      kegiatan_khusus: kegiatan,
      ketidakhadiran_baru: aRows.length,
      total_baris_baru: pRows.length + aRows.length,
      penghematan_baris_persen: barisLegacy > 0 ? Math.round((1 - (pRows.length + aRows.length) / barisLegacy) * 1000) / 10 : 0,
      pertemuan_tanpa_pengajar: tanpaPengajar,
      kelas_hari_diabsen_lebih_dari_sekali: duplikat.length,
      contoh_kelas_hari_ganda: duplikat.slice(0, 10)
    }
  };
}

function presensiAgregatPerToken_(pRows, aRows) {
  const m = {};
  pRows.forEach(function (r) {
    const ta = String(r[3] || '').trim();
    String(r[9] || '').split('|').forEach(function (tok) {
      if (!tok) return;
      const k = tok + '||' + ta;
      if (!m[k]) m[k] = { T: 0, S: 0, I: 0, A: 0 };
      m[k].T++;
    });
  });
  aRows.forEach(function (r) {
    const tok = presensiToken_(r[3], r[4]);
    if (!tok) return;
    const e = m[tok + '||' + String(r[7] || '').trim()];
    if (!e) return;
    const st = presensiNormStatus_(r[6]);
    if (st !== 'H') e[st]++;
  });
  Object.keys(m).forEach(function (k) {
    const e = m[k];
    e.H = Math.max(0, e.T - e.S - e.I - e.A);
  });
  return m;
}

// Bandingkan total H/S/I/A per santri (model lama vs model baru)
function presensiBandingkan_(legacyAgg, baruAgg) {
  const keys = {};
  Object.keys(legacyAgg).forEach(function (k) { keys[k] = true; });
  Object.keys(baruAgg).forEach(function (k) { keys[k] = true; });
  const beda = [];
  let cocok = 0;
  Object.keys(keys).forEach(function (k) {
    const a = legacyAgg[k] || { H: 0, S: 0, I: 0, A: 0 };
    const b = baruAgg[k] || { H: 0, S: 0, I: 0, A: 0 };
    if (a.H === b.H && a.S === b.S && a.I === b.I && a.A === b.A) cocok++;
    else beda.push({ santri_tahun: k, lama: a, baru: { H: b.H, S: b.S, I: b.I, A: b.A } });
  });
  return { santri_dibandingkan: cocok + beda.length, identik: cocok, berbeda: beda.length, contoh_berbeda: beda.slice(0, 10) };
}

/** LANGKAH 1 - Simulasi (hanya membaca, tidak menulis apa pun). Aman dijalankan kapan saja. */
function presensiMigrasiSimulasi() {
  const plan = presensiRencanaMigrasi_();
  const verifikasi = presensiBandingkan_(plan.legacyAgg, presensiAgregatPerToken_(plan.pRows, plan.aRows));
  const hasil = { ringkasan: plan.ringkasan, verifikasi_hitungan_per_santri: verifikasi };
  Logger.log(JSON.stringify(hasil, null, 2));
  return hasil;
}

function presensiTulisMigrasi_(timpa) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const plan = presensiRencanaMigrasi_();
    const shP = presensiSheetPertemuan_();
    const shA = presensiSheetAbsen_();
    const exP = presensiBaca_(shP, PRES_HDR_PERTEMUAN.length);
    const exA = presensiBaca_(shA, PRES_HDR_ABSEN.length);
    const sudahAda = exP.some(function (r) { return String(r[0]).indexOf('PT-MIG-') === 0; });
    if (sudahAda && !timpa) {
      throw new Error('Migrasi sudah pernah dijalankan. Gunakan presensiMigrasiUlang() bila ingin menimpa hasil migrasi sebelumnya.');
    }
    const keepP = exP.filter(function (r) { return String(r[0]).indexOf('PT-MIG-') !== 0; });
    const keepA = exA.filter(function (r) { return String(r[0]).indexOf('AB-MIG-') !== 0; });
    const finalP = plan.pRows.concat(keepP);
    const finalA = plan.aRows.concat(keepA);

    if (shP.getLastRow() > 1) shP.getRange(2, 1, shP.getLastRow() - 1, PRES_HDR_PERTEMUAN.length).clearContent();
    if (shA.getLastRow() > 1) shA.getRange(2, 1, shA.getLastRow() - 1, PRES_HDR_ABSEN.length).clearContent();
    presensiTulis_(shP, 2, finalP, PRES_HDR_PERTEMUAN.length, PRES_TEXT_COLS_PERTEMUAN);
    presensiTulis_(shA, 2, finalA, PRES_HDR_ABSEN.length, PRES_TEXT_COLS_ABSEN);
    SpreadsheetApp.flush();

    // Verifikasi dari isi sheet yang benar-benar tertulis (hanya bila tidak ada data baru di luar migrasi)
    const bacaP = presensiBaca_(shP, PRES_HDR_PERTEMUAN.length);
    const bacaA = presensiBaca_(shA, PRES_HDR_ABSEN.length);
    const verifikasi = (keepP.length === 0 && keepA.length === 0)
      ? presensiBandingkan_(plan.legacyAgg, presensiAgregatPerToken_(bacaP, bacaA))
      : 'dilewati (ada data baru di luar hasil migrasi)';
    const hasil = {
      pesan: 'Migrasi selesai. Data lama di sheet Presensi_Santri TIDAK diubah. Mode aplikasi belum berubah.',
      ringkasan: plan.ringkasan,
      baris_tertulis: { Pertemuan: bacaP.length, Presensi_Ketidakhadiran: bacaA.length },
      verifikasi_dari_sheet: verifikasi
    };
    Logger.log(JSON.stringify(hasil, null, 2));
    return hasil;
  } finally {
    lock.releaseLock();
  }
}

/** LANGKAH 2 - Jalankan migrasi (menulis ke sheet baru; data lama tidak disentuh). */
function presensiMigrasiJalankan() {
  presensiPastikanAdmin_();
  return presensiTulisMigrasi_(false);
}

/** Ulangi migrasi (menghapus baris hasil migrasi sebelumnya, data lama & data baru non-migrasi tetap). */
function presensiMigrasiUlang() {
  presensiPastikanAdmin_();
  return presensiTulisMigrasi_(true);
}

/** LANGKAH 3 - Alihkan aplikasi ke model baru setelah hasil migrasi diperiksa. */
function presensiAktifkanModeBaru() {
  presensiPastikanAdmin_();
  const adaMigrasi = presensiBaca_(presensiSheetPertemuan_(), PRES_HDR_PERTEMUAN.length).length > 0;
  const lamaKosong = !SHEETS.presensiSantri || SHEETS.presensiSantri.getLastRow() <= 1;
  if (!adaMigrasi && !lamaKosong) {
    throw new Error('Sheet Pertemuan masih kosong. Jalankan presensiMigrasiJalankan() terlebih dahulu.');
  }
  PropertiesService.getScriptProperties().setProperty('PRESENSI_MODE', 'BARU');
  return 'Mode presensi: BARU (memakai sheet Pertemuan + Presensi_Ketidakhadiran).';
}

/** Kembali ke model lama (data yang disimpan selama mode BARU tetap aman di sheet baru). */
function presensiKembaliKeModeLama() {
  presensiPastikanAdmin_();
  PropertiesService.getScriptProperties().setProperty('PRESENSI_MODE', 'LAMA');
  return 'Mode presensi: LAMA (memakai sheet Presensi_Santri).';
}

/** Ringkasan kondisi saat ini (hanya membaca). */
function presensiStatistik() {
  const lama = SHEETS.presensiSantri ? Math.max(0, SHEETS.presensiSantri.getLastRow() - 1) : 0;
  const shP = SS.getSheetByName(PRES_SH_PERTEMUAN);
  const shA = SS.getSheetByName(PRES_SH_ABSEN);
  const hasil = {
    mode: presensiModeBaru_() ? 'BARU' : 'LAMA',
    baris_Presensi_Santri: lama,
    baris_Pertemuan: shP ? Math.max(0, shP.getLastRow() - 1) : 0,
    baris_Presensi_Ketidakhadiran: shA ? Math.max(0, shA.getLastRow() - 1) : 0
  };
  Logger.log(JSON.stringify(hasil));
  return hasil;
}
