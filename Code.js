/**
 * APLIKASI TPQ AL MABRUR SINGOSARI
 * Main Backend Code
 */

// Global variables
const SPREADSHEET_ID = '1G7djwD3qq6R4Jp9agI0tnbKoc1QC1T6G0Y-s4eCrBCo';
const SS = (function() {
  try {
    const act = SpreadsheetApp.getActiveSpreadsheet();
    if (act && act.getId() === SPREADSHEET_ID) return act;
    if (act) return act;
  } catch (e) {}
  return SpreadsheetApp.openById(SPREADSHEET_ID);
})();
const SHEETS = {
  config: SS.getSheetByName('Config'),
  users: SS.getSheetByName('Users'),
  santri: SS.getSheetByName('Santri'),
  guru: SS.getSheetByName('Guru'),
  kelas: SS.getSheetByName('Kelas'),
  materiKelas: SS.getSheetByName('Materi_Kelas'),
  pembayaran: SS.getSheetByName('Pembayaran') || SS.insertSheet('Pembayaran'),
  pemasukan: SS.getSheetByName('Pemasukan') || SS.insertSheet('Pemasukan'),
  pengeluaran: SS.getSheetByName('Pengeluaran'),
  jenisTagihan: SS.getSheetByName('Jenis_Tagihan') || SS.insertSheet('Jenis_Tagihan'),
  tagihanConfig: SS.getSheetByName('Tagihan_Konfigurasi') || SS.insertSheet('Tagihan_Konfigurasi'),
  jenisKegiatanGuru: SS.getSheetByName('Jenis_Kegiatan_Guru') || SS.insertSheet('Jenis_Kegiatan_Guru'),
  presensiSantri: SS.getSheetByName('Presensi_Santri'),
  presensiGuru: SS.getSheetByName('Presensi_Guru'),
  nilaiRapor: SS.getSheetByName('Nilai_Rapor'),
  pengumuman: SS.getSheetByName('Pengumuman'),
  catatanKeseharian: SS.getSheetByName('Catatan_Keseharian'),
  ceklisConfig: SS.getSheetByName('Ceklis_Konfigurasi') || SS.insertSheet('Ceklis_Konfigurasi'),
  ceklisData: SS.getSheetByName('Ceklis_Data') || SS.insertSheet('Ceklis_Data'),
  nilaiBupres: SS.getSheetByName('Nilai_Bupres') || SS.insertSheet('Nilai_Bupres'),
};

function ensureSheetHeaders_() {
  // Nilai_Bupres
  if (SHEETS.nilaiBupres.getLastRow() === 0) {
    SHEETS.nilaiBupres.appendRow([
      'ID_Bupres', 'Tahun_Ajaran', 'Kelas', 'NIS', 'Nama_Santri', 
      'B', 'B_Minus', 'C_Plus', 'C', 'C_Minus', 'K', 
      'Total_Nilai', 'Tatap_Muka', 'Poin_BCK', 'Catatan', 'Input_Oleh', 'Updated_At'
    ]);
  }
  // Ceklis_Konfigurasi
  if (SHEETS.ceklisConfig.getLastRow() === 0) {
    SHEETS.ceklisConfig.appendRow(['ID_Ceklis', 'Nama_Ceklis', 'Deskripsi', 'Daftar_Kolom', 'Target_Kelas', 'Status', 'Created_At']);
  }
  // Ceklis_Data
  if (SHEETS.ceklisData.getLastRow() === 0) {
    SHEETS.ceklisData.appendRow(['ID_Data', 'ID_Ceklis', 'Kelas', 'NIS', 'Nama_Santri', 'Hasil_Ceklis', 'Catatan', 'Updated_By', 'Updated_At']);
  }
  // Pemasukan
  if (SHEETS.pemasukan.getLastRow() === 0) {
    SHEETS.pemasukan.appendRow(['ID_Pemasukan', 'Tanggal', 'Sumber', 'Keterangan', 'Nominal', 'Input_Oleh', 'Created_At']);
  }
  // Jenis_Tagihan
  if (SHEETS.jenisTagihan.getLastRow() === 0) {
    SHEETS.jenisTagihan.appendRow(['Jenis_Tagihan']);
    ['SPP', 'Pendaftaran', 'Seragam', 'Buku', 'Lainnya'].forEach(j => SHEETS.jenisTagihan.appendRow([j]));
  }
  // Tagihan_Konfigurasi
  if (SHEETS.tagihanConfig.getLastRow() === 0) {
    SHEETS.tagihanConfig.appendRow(['ID_Tagihan', 'Tahun_Ajaran', 'Jenis_Tagihan', 'Nominal', 'Keterangan']);
  }
  // Jenis_Kegiatan_Guru
  if (SHEETS.jenisKegiatanGuru.getLastRow() === 0) {
    SHEETS.jenisKegiatanGuru.appendRow(['Jenis_Kegiatan']);
    ['Mengajar', 'Rapat', 'Kegiatan Khusus'].forEach(j => SHEETS.jenisKegiatanGuru.appendRow([j]));
  }
  
  // Cek & Tambah Kolom Baru di Sheet Santri
  if (SHEETS.santri.getLastRow() > 0) {
    const sCols = SHEETS.santri.getLastColumn();
    const sHeaders = SHEETS.santri.getRange(1, 1, 1, sCols).getValues()[0];
    const sMissing = [];
    ['Tahun_Ajaran', 'Status', 'Alamat'].forEach(col => {
      if (!sHeaders.includes(col)) sMissing.push(col);
    });
    if (sMissing.length > 0) {
      SHEETS.santri.getRange(1, sCols + 1, 1, sMissing.length).setValues([sMissing]);
    }
  }

  // Cek & Tambah Kolom Baru di Sheet Guru
  if (SHEETS.guru.getLastRow() > 0) {
    const gCols = SHEETS.guru.getLastColumn();
    const gHeaders = SHEETS.guru.getRange(1, 1, 1, gCols).getValues()[0];
    const gMissing = [];
    ['Sekolah_Universitas', 'Kelas_Diampu', 'Status', 'Alamat_Lengkap'].forEach(col => {
      if (!gHeaders.includes(col)) gMissing.push(col);
    });
    if (gMissing.length > 0) {
      SHEETS.guru.getRange(1, gCols + 1, 1, gMissing.length).setValues([gMissing]);
    }
  }
}

// Serve HTML
function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Aplikasi TPQ Al Mabrur Singosari')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===== AUTHENTICATION =====

function login(username) {
  const searchUsername = String(username || '').trim();
  if (!searchUsername) {
    return { success: false, message: 'Username atau Nomor WhatsApp wajib diisi' };
  }

  const searchDigits = searchUsername.replace(/\D/g, '');

  // 1. Cek di Sheet Users (Admin, Ustadz, Bendahara, atau Akun Khusus)
  const userSheet = SHEETS.users;
  const userData = userSheet.getDataRange().getValues();
  
  for (let i = 1; i < userData.length; i++) {
    const row = userData[i];
    const rowUsername = String(row[1] || '').trim();
    const rowHp = String(row[5] || '').replace(/\D/g, '');
    const rowStatus = String(row[6] || '').toLowerCase();
    
    const matchUsername = (rowUsername.toLowerCase() === searchUsername.toLowerCase());
    const matchHp = searchDigits && rowHp && (searchDigits === rowHp || searchDigits.endsWith(rowHp) || rowHp.endsWith(searchDigits));
    
    if ((matchUsername || matchHp) && rowStatus === 'aktif') {
      userSheet.getRange(i + 1, 8).setValue(new Date());
      
      return {
        success: true,
        user: {
          id: row[0],
          username: row[1],
          role: row[3],
          nama: row[4],
          noHp: row[5],
          foto: (function(urlRaw) {
            const url = String(urlRaw || '');
            if (!url) return '';
            let m = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
            return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w500' : url;
          })(row[9]),
        }
      };
    }
  }

  // 2. Cek di Sheet Santri (Login Mandiri Orang Tua / Santri via No WA Ayah/Ibu atau NIS)
  const santriSheet = SHEETS.santri;
  const santriData = santriSheet.getDataRange().getValues();
  const santriHeaders = santriData[0];
  
  const idxNIS = santriHeaders.findIndex(h => h.toLowerCase() === 'nis');
  const idxNama = santriHeaders.findIndex(h => h.toLowerCase().includes('nama_santri') || h.toLowerCase() === 'nama');
  const idxWaAyah = santriHeaders.findIndex(h => h.toLowerCase().includes('wa_ayah'));
  const idxWaIbu = santriHeaders.findIndex(h => h.toLowerCase().includes('wa_ibu'));
  const idxNamaAyah = santriHeaders.findIndex(h => h.toLowerCase().includes('nama_ayah'));
  const idxNamaIbu = santriHeaders.findIndex(h => h.toLowerCase().includes('nama_ibu'));
  const idxFoto = santriHeaders.findIndex(h => h.toLowerCase().includes('foto'));
  
  // Normalisasi digit input nomor HP
  const cleanPhone = (num) => {
    let n = String(num || '').replace(/\D/g, '');
    if (n.startsWith('62')) n = n.slice(2);
    if (n.startsWith('0')) n = n.slice(1);
    return n;
  };
  
  const inputPhoneClean = cleanPhone(searchDigits);
  
  for (let i = 1; i < santriData.length; i++) {
    const sRow = santriData[i];
    const sNis = String(idxNIS !== -1 ? sRow[idxNIS] : sRow[0]).trim();
    const sNama = String(idxNama !== -1 ? sRow[idxNama] : sRow[1]).trim();
    const sWaAyah = cleanPhone(idxWaAyah !== -1 ? sRow[idxWaAyah] : '');
    const sWaIbu = cleanPhone(idxWaIbu !== -1 ? sRow[idxWaIbu] : '');
    const sNamaAyah = String(idxNamaAyah !== -1 ? sRow[idxNamaAyah] : '').trim();
    const sNamaIbu = String(idxNamaIbu !== -1 ? sRow[idxNamaIbu] : '').trim();
    const sFoto = String(idxFoto !== -1 ? sRow[idxFoto] : '').trim();
    
    // Cocokkan NIS
    const matchNis = sNis && (sNis.toLowerCase() === searchUsername.toLowerCase());
    
    // Cocokkan No WA
    let matchWa = false;
    let matchedParent = '';
    
    if (inputPhoneClean && inputPhoneClean.length >= 7) {
      if (sWaAyah && (inputPhoneClean === sWaAyah || inputPhoneClean.endsWith(sWaAyah) || sWaAyah.endsWith(inputPhoneClean))) {
        matchWa = true;
        matchedParent = sNamaAyah ? 'Bpk. ' + sNamaAyah : 'Wali Santri';
      } else if (sWaIbu && (inputPhoneClean === sWaIbu || inputPhoneClean.endsWith(sWaIbu) || sWaIbu.endsWith(inputPhoneClean))) {
        matchWa = true;
        matchedParent = sNamaIbu ? 'Ibu ' + sNamaIbu : 'Wali Santri';
      }
    }
    
    if (matchNis || matchWa) {
      const displayName = matchWa && matchedParent ? `${matchedParent} (${sNama})` : sNama;
      
      return {
        success: true,
        user: {
          id: 'WALI-' + sNis,
          username: sNis, // Tetapkan NIS sebagai username agar sinkron dengan seluruh fitur dasbor & tagihan
          role: 'wali_santri',
          nama: displayName,
          noHp: matchWa ? searchUsername : (sRow[idxWaAyah] || sRow[idxWaIbu] || ''),
          foto: (function(urlRaw) {
            const url = String(sFoto || '');
            if (!url) return '';
            let m = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
            return m ? 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w500' : url;
          })(),
        }
      };
    }
  }
  
  return { success: false, message: 'Username, NIS, atau Nomor WhatsApp tidak ditemukan di database TPQ.' };
}

// ===== CONFIG =====

function saveConfig(key, value) {
  try {
    const sheet = SHEETS.config;
    const data = sheet.getDataRange().getValues();
    let found = false;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(value);
        found = true;
        break;
      }
    }
    
    if (!found) {
      sheet.appendRow([key, value]);
    }
    
    return { success: true, message: 'Tahun Ajaran berhasil disimpan sebagai default' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getConfig() {
  const data = SHEETS.config.getDataRange().getValues();
  const config = {};
  
  for (let i = 1; i < data.length; i++) {
    config[data[i][0]] = data[i][1];
  }
  
  return config;
}

// ===== SANTRI =====

function getSantriList() {
  const data = SHEETS.santri.getDataRange().getValues();
  const headers = data[0];
  const santriList = [];
  
  // Get Kelas mapping for Lokasi_Kelas
  const kelasData = SHEETS.kelas.getDataRange().getValues();
  const kelasHeaders = kelasData[0];
  const idxNamaKelas = kelasHeaders.findIndex(h => h === 'Nama Kelas' || h === 'Nama_Kelas');
  const idxLokasiKelas = kelasHeaders.findIndex(h => h === 'Lokasi Kelas' || h === 'Lokasi_Kelas');
  
  const kelasMap = {};
  if (idxNamaKelas !== -1 && idxLokasiKelas !== -1) {
    for (let i = 1; i < kelasData.length; i++) {
      if (kelasData[i][idxNamaKelas]) {
        kelasMap[kelasData[i][idxNamaKelas]] = kelasData[i][idxLokasiKelas];
      }
    }
  }

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0] && !row[1] && !row[3]) continue;
    const santri = {};
    headers.forEach((header, index) => {
      let val = row[index];
      if (val instanceof Date) val = val.toISOString();
      santri[header] = val;
    });
    
    // Add Lokasi Kelas
    if (santri.Kelas && kelasMap[santri.Kelas]) {
      santri.Lokasi_Kelas = kelasMap[santri.Kelas];
    } else {
      santri.Lokasi_Kelas = '-';
    }
    
    santriList.push(santri);
  }
  
  return santriList;
}

function saveSantri(data) {
  try {
    const sheet = SHEETS.santri;
    const headers = sheet.getDataRange().getValues()[0];
    
    // Normalize keys
    const dataNorm = {};
    Object.keys(data).forEach(k => { dataNorm[k.toLowerCase()] = data[k]; });
    
    if (data.NIS) {
      // Update existing
      const range = sheet.getDataRange().getValues();
      for (let i = 1; i < range.length; i++) {
        if (range[i][0] === data.NIS) {
          headers.forEach((header, index) => {
            const hKey = header.toLowerCase();
            if (dataNorm[hKey] !== undefined) {
              sheet.getRange(i + 1, index + 1).setValue(dataNorm[hKey]);
            }
          });
          return { success: true, message: 'Data santri berhasil diupdate' };
        }
      }
    } else {
      // Insert new
      const nisInfo = getNextNIS();
      const nis = (dataNorm['nis'] && String(dataNorm['nis']).trim()) ? String(dataNorm['nis']).trim() : nisInfo.nextNIS;
      dataNorm['nis'] = nis;
      dataNorm['tanggal_masuk'] = new Date();
      dataNorm['created_at'] = new Date();
      if (!dataNorm['status']) dataNorm['status'] = 'aktif';
      
      const newRow = headers.map(header => {
        const hKey = header.toLowerCase();
        return dataNorm[hKey] !== undefined ? dataNorm[hKey] : '';
      });
      
      sheet.appendRow(newRow);
      return { success: true, message: 'Data santri berhasil ditambahkan', nis: nis };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteSantri(nis) {
  try {
    const sheet = SHEETS.santri;
    const range = sheet.getDataRange().getValues();
    const targetNis = String(nis || '').trim();
    
    for (let i = 1; i < range.length; i++) {
      if (String(range[i][0]).trim() === targetNis) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data santri berhasil dihapus' };
      }
    }
    
    return { success: false, message: 'Santri tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

/**
 * Upload dan perbarui foto profil santri ke Google Drive dan simpan URL-nya di sheet Santri
 */
function uploadFotoSantri(data) {
  try {
    if (!data || !data.nis) {
      throw new Error('NIS santri tidak ditemukan.');
    }
    if (!data.imageBase64) {
      throw new Error('File foto tidak ditemukan atau kosong.');
    }
    
    const nis = String(data.nis).trim();
    let base64Str = String(data.imageBase64);
    let mimeType = 'image/jpeg';
    
    // Parse data URL jika ada prefix base64
    if (base64Str.indexOf(';base64,') !== -1) {
      const parts = base64Str.split(';base64,');
      const mimeMatch = parts[0].match(/data:(.*?)$/);
      if (mimeMatch) mimeType = mimeMatch[1];
      base64Str = parts[1];
    }
    
    const decodedBytes = Utilities.base64Decode(base64Str);
    const fileName = 'FOTO_' + nis + '_' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss') + '.jpg';
    const blob = Utilities.newBlob(decodedBytes, mimeType, fileName);
    
    // Cari atau buat folder "Foto_Santri_TPQ" di Google Drive
    let folder;
    const folders = DriveApp.getFoldersByName('Foto_Santri_TPQ');
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('Foto_Santri_TPQ');
    }
    try {
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {}
    
    // Buat file foto
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {}
    const fileId = file.getId();
    
    // URL publik untuk thumbnail & direct view
    const fileUrl = 'https://drive.google.com/uc?export=view&id=' + fileId;
    const thumbUrl = 'https://drive.google.com/thumbnail?id=' + fileId + '&sz=w500';
    
    // Update kolom Foto_URL di sheet Santri
    const sheet = SHEETS.santri;
    const dataRows = sheet.getDataRange().getValues();
    const headers = dataRows[0];
    
    let idxFoto = headers.findIndex(h => h.toString().toLowerCase().includes('foto'));
    if (idxFoto === -1) {
      const nextCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, nextCol).setValue('Foto_URL');
      idxFoto = nextCol - 1;
    }
    
    const idxNIS = headers.findIndex(h => h.toString().toLowerCase() === 'nis');
    const nisCol = idxNIS !== -1 ? idxNIS : 0;
    
    let foundRow = -1;
    for (let i = 1; i < dataRows.length; i++) {
      if (String(dataRows[i][nisCol] || '').trim() === nis) {
        foundRow = i + 1;
        break;
      }
    }
    
    if (foundRow === -1) {
      throw new Error('Santri dengan NIS ' + nis + ' tidak ditemukan di database.');
    }
    
    sheet.getRange(foundRow, idxFoto + 1).setValue(fileUrl);
    
    return {
      success: true,
      fileUrl: fileUrl,
      thumbUrl: thumbUrl,
      message: 'Foto santri berhasil diupload dan disimpan!'
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || err.toString()
    };
  }
}

/**
 * Update biodata mandiri oleh Orang Tua / Wali Santri dari Dasbor Santri
 * Memperbarui Identitas Diri dan Data Orang Tua / Wali di sheet Santri
 */
function updateBiodataSantriWali(formData) {
  try {
    if (!formData || !formData.NIS) {
      throw new Error('NIS santri tidak ditemukan.');
    }
    
    const nis = String(formData.NIS).trim();
    const sheet = SHEETS.santri;
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idxNIS = headers.findIndex(h => h.toString().toLowerCase() === 'nis');
    const nisCol = idxNIS !== -1 ? idxNIS : 0;
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][nisCol] || '').trim() === nis) {
        rowIndex = i + 1; // 1-indexed for Sheets
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error('Data santri dengan NIS ' + nis + ' tidak ditemukan di database.');
    }
    
    // Mapping kolom yang diizinkan untuk diubah oleh Orang Tua / Wali
    const allowedFields = {
      'nama_santri': formData.Nama_Santri || formData.nama || '',
      'nama': formData.Nama_Santri || formData.nama || '',
      'jenis_kelamin': formData.Jenis_Kelamin || formData.jk || 'L',
      'l/p': formData.Jenis_Kelamin || formData.jk || 'L',
      'tempat_lahir': formData.Tempat_Lahir || '',
      'tanggal_lahir': formData.Tanggal_Lahir || '',
      'usia': formData.Usia || '',
      'nama_sekolah': formData.Nama_Sekolah || '',
      'alamat': formData.Alamat || formData.Alamat_Lengkap || '',
      'alamat_lengkap': formData.Alamat || formData.Alamat_Lengkap || '',
      'nama_ayah': formData.Nama_Ayah || '',
      'no_wa_ayah': formData.No_WA_Ayah || '',
      'wa_ayah': formData.No_WA_Ayah || '',
      'nama_ibu': formData.Nama_Ibu || '',
      'no_wa_ibu': formData.No_WA_Ibu || '',
      'wa_ibu': formData.No_WA_Ibu || ''
    };
    
    // Perbarui cell pada baris santri tersebut
    headers.forEach((header, colIdx) => {
      const hRaw = header.toString().toLowerCase().trim();
      const hKey = hRaw.replace(/\s+/g, '_');
      let valToSet = undefined;
      
      if (allowedFields.hasOwnProperty(hRaw)) {
        valToSet = allowedFields[hRaw];
      } else if (allowedFields.hasOwnProperty(hKey)) {
        valToSet = allowedFields[hKey];
      }
      
      if (valToSet !== undefined) {
        sheet.getRange(rowIndex, colIdx + 1).setValue(valToSet);
      }
    });
    
    // Ambil data santri terbaru setelah update
    const updatedRow = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
    const updatedSantri = {};
    headers.forEach((h, idx) => {
      let val = updatedRow[idx];
      if (val instanceof Date) val = Utilities.formatDate(val, 'GMT+7', 'yyyy-MM-dd');
      updatedSantri[h] = val;
    });
    
    return {
      success: true,
      message: 'Biodata santri dan data orang tua berhasil diperbarui!',
      santri: updatedSantri
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || err.toString()
    };
  }
}

// ===== PENDAFTARAN SANTRI BARU & NIS OTOMATIS =====

/**
 * Menghasilkan NIS baru otomatis dengan format YY + NNNN (contoh: 261066)
 * dihitung dari NIS terbesar yang ada di sheet Santri.
 */
function getNextNIS() {
  try {
    const sheet = SHEETS.santri;
    const data = sheet.getDataRange().getValues();
    const currentYear = new Date().getFullYear();
    const yearPrefix = String(currentYear).slice(-2); // e.g. "26"
    
    let maxSeq = 0;
    let maxNisOverall = 0;
    
    for (let i = 1; i < data.length; i++) {
      const val = data[i][0];
      if (val === undefined || val === null || val === '') continue;
      const strVal = String(val).trim();
      const numVal = parseInt(strVal.replace(/\D/g, ''), 10);
      if (!isNaN(numVal) && numVal > maxNisOverall) {
        maxNisOverall = numVal;
      }
      
      // Cek format 6 digit: YY + NNNN (contoh: 261065, 210581)
      const m = strVal.match(/^(\d{2})(\d{4})$/);
      if (m) {
        const seqNum = parseInt(m[2], 10);
        if (!isNaN(seqNum) && seqNum > maxSeq) {
          maxSeq = seqNum;
        }
      } else {
        const digits = strVal.replace(/\D/g, '');
        if (digits.length >= 6) {
          const last4 = parseInt(digits.slice(-4), 10);
          if (!isNaN(last4) && last4 > maxSeq) {
            maxSeq = last4;
          }
        }
      }
    }
    
    if (maxSeq === 0) {
      maxSeq = 1000;
    }
    
    const nextSeq = maxSeq + 1;
    const nextSeqPadded = String(nextSeq).padStart(4, '0');
    const nextNIS = yearPrefix + nextSeqPadded;
    const lastNIS = maxNisOverall > 0 ? String(maxNisOverall) : (yearPrefix + String(maxSeq).padStart(4, '0'));
    
    return {
      success: true,
      yearPrefix: yearPrefix,
      maxSeq: maxSeq,
      nextSeq: nextSeq,
      nextNIS: nextNIS,
      lastNIS: lastNIS
    };
  } catch (err) {
    const fallbackYear = String(new Date().getFullYear()).slice(-2);
    return {
      success: false,
      message: err.message || err.toString(),
      yearPrefix: fallbackYear,
      nextNIS: fallbackYear + '1066',
      lastNIS: fallbackYear + '1065'
    };
  }
}

/**
 * Mengambil Tahun Ajaran aktif pendaftaran santri baru dari Sheet Config.
 */
function getTahunAjaranPendaftaran() {
  try {
    const cfg = getConfig();
    return String(cfg.TAHUN_AJARAN_PENDAFTARAN || cfg.tahun_ajaran_pendaftaran || cfg.TAHUN_AJARAN || cfg.tahun_ajaran || '20262').trim();
  } catch (err) {
    return '20262';
  }
}

/**
 * Mengatur Tahun Ajaran aktif pendaftaran santri baru.
 */
function setTahunAjaranPendaftaran(tahunAjaran) {
  try {
    const val = String(tahunAjaran || '').trim();
    if (!val) return { success: false, message: 'Tahun Ajaran tidak boleh kosong' };
    
    saveConfig('TAHUN_AJARAN_PENDAFTARAN', val);
    return {
      success: true,
      message: 'Tahun Ajaran Pendaftaran berhasil diatur ke ' + val,
      tahunAjaran: val
    };
  } catch (err) {
    return { success: false, message: err.message || err.toString() };
  }
}

/**
 * Pendaftaran Santri Baru (Online dari Halaman Login atau dari Dashboard Admin)
 * Menyimpan ke sheet Santri di spreadsheet 1G7djwD3qq6R4Jp9agI0tnbKoc1QC1T6G0Y-s4eCrBCo
 */
function daftarSantriBaru(formData) {
  try {
    if (!formData) throw new Error('Data pendaftaran kosong');
    const namaSantri = String(formData.Nama_Santri || formData.nama || '').trim();
    if (!namaSantri) throw new Error('Nama Santri wajib diisi');
    
    const sheet = SHEETS.santri;
    const headers = sheet.getDataRange().getValues()[0];
    
    // Ambil tahun ajaran aktif pendaftaran
    const tahunAjaran = String(formData.Tahun_Ajaran || getTahunAjaranPendaftaran()).trim();
    
    // Tentukan NIS: gunakan input atau auto generate dari terbesar
    let nis = String(formData.NIS || '').trim();
    if (!nis) {
      const nisInfo = getNextNIS();
      nis = nisInfo.nextNIS;
    }
    
    // Normalisasi data sesuai nama kolom di sheet
    const dataNorm = {
      'nis': nis,
      'nama_santri': namaSantri,
      'jenis_kelamin': formData.Jenis_Kelamin || formData.jk || 'L',
      'tempat_lahir': formData.Tempat_Lahir || '',
      'tanggal_lahir': formData.Tanggal_Lahir || '',
      'usia': formData.Usia || '',
      'nama_sekolah': formData.Nama_Sekolah || '',
      'kelas': formData.Kelas || '-',
      'nama_ayah': formData.Nama_Ayah || '',
      'no_wa_ayah': formData.No_WA_Ayah || '',
      'nama_ibu': formData.Nama_Ibu || '',
      'no_wa_ibu': formData.No_WA_Ibu || '',
      'alamat': formData.Alamat || formData.Alamat_Lengkap || '',
      'alamat_lengkap': formData.Alamat || formData.Alamat_Lengkap || '',
      'kemampuan_mengaji': formData.Kemampuan_Mengaji || '-',
      'tanggal_daftar': new Date(),
      'status': formData.Status || 'Mendaftar',
      'foto_url': formData.Foto_URL || '',
      'tahun_ajaran': tahunAjaran,
      'lokasi_kelas': formData.Lokasi_Kelas || '-'
    };
    
    const newRow = headers.map(header => {
      const hKey = header.toLowerCase();
      return dataNorm[hKey] !== undefined ? dataNorm[hKey] : '';
    });
    
    sheet.appendRow(newRow);
    
    // Jika status Aktif dan diminta buat user login, tambahkan ke sheet Users
    if (formData.buatUser === true || formData.Status === 'Aktif') {
      try {
        const noHp = String(formData.No_WA_Ayah || formData.No_WA_Ibu || '').trim();
        createUserSantri_(nis, namaSantri, noHp);
      } catch (eUser) {
        console.warn('Gagal buat user santri:', eUser);
      }
    }
    
    return {
      success: true,
      message: 'Pendaftaran Santri Baru berhasil disimpan!',
      nis: nis,
      namaSantri: namaSantri,
      tahunAjaran: tahunAjaran,
      status: dataNorm.status
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || err.toString()
    };
  }
}

/**
 * Helper untuk membuat akun user login otomatis bagi santri/wali di sheet Users
 */
function createUserSantri_(nis, nama, noHp) {
  const sheetUsers = SHEETS.users;
  const data = sheetUsers.getDataRange().getValues();
  const searchUsername = String(nis).trim();
  
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === searchUsername) {
      return;
    }
  }
  
  const newId = 'U' + (data.length);
  sheetUsers.appendRow([
    newId,
    searchUsername,
    searchUsername, // password default = NIS
    'wali_santri',
    nama,
    noHp || '',
    'aktif',
    '',
    new Date(),
    ''
  ]);
}

/**
 * Mengambil daftar pendaftaran santri untuk dipantau di Dasbor Admin
 */
function getPendaftaranList(tahunFilter) {
  try {
    const santriList = getSantriList();
    const tahunAktif = getTahunAjaranPendaftaran();
    
    let result = santriList;
    if (tahunFilter && tahunFilter !== 'semua') {
      result = result.filter(s => String(s.Tahun_Ajaran || '').trim() === String(tahunFilter).trim());
    }
    
    // Urutkan: Pendaftar yang masih status 'mendaftar' (menunggu) berada di paling atas,
    // disusul oleh pendaftar yang sudah diverifikasi berdasarkan NIS terbesar / terbaru
    result.sort((a, b) => {
      const stA = String(a.Status || 'aktif').toLowerCase().trim();
      const stB = String(b.Status || 'aktif').toLowerCase().trim();
      const isMendaftarA = stA === 'mendaftar' || stA === 'pending' || stA.includes('daftar');
      const isMendaftarB = stB === 'mendaftar' || stB === 'pending' || stB.includes('daftar');
      
      if (isMendaftarA && !isMendaftarB) return -1;
      if (!isMendaftarA && isMendaftarB) return 1;
      
      const nisA = parseInt(String(a.NIS || '').replace(/\D/g, ''), 10) || 0;
      const nisB = parseInt(String(b.NIS || '').replace(/\D/g, ''), 10) || 0;
      return nisB - nisA;
    });
    
    return {
      success: true,
      data: result,
      tahunAjaranAktif: tahunAktif,
      nisInfo: getNextNIS()
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || err.toString(),
      data: []
    };
  }
}

/**
 * Mengambil statistik pendaftaran untuk dasbor admin
 */
function getPendaftaranStats() {
  try {
    const sheet = SHEETS.santri;
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const idxStatus = headers.findIndex(h => h.toLowerCase() === 'status');
    const idxTahun = headers.findIndex(h => h.toLowerCase() === 'tahun_ajaran');
    
    const tahunAktif = getTahunAjaranPendaftaran();
    let totalPendaftarBaru = 0;
    let totalAktifTahunIni = 0;
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0] && !row[1]) continue;
      
      const st = String(idxStatus !== -1 ? row[idxStatus] : 'aktif').toLowerCase().trim();
      const thn = String(idxTahun !== -1 ? row[idxTahun] : '').trim();
      
      if (st === 'mendaftar' || st === 'pending' || st.includes('daftar')) {
        totalPendaftarBaru++;
      }
      
      if (st === 'aktif' && thn === tahunAktif) {
        totalAktifTahunIni++;
      }
    }
    
    const totalPendaftar = totalPendaftarBaru + totalAktifTahunIni;
    
    return {
      success: true,
      tahunAjaranAktif: tahunAktif,
      nisInfo: getNextNIS(),
      totalPendaftarBaru: totalPendaftarBaru,
      totalAktifTahunIni: totalAktifTahunIni,
      totalPendaftar: totalPendaftar,
      totalSemua: totalPendaftar > 0 ? totalPendaftar : totalPendaftarBaru
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || err.toString(),
      tahunAjaranAktif: '20262',
      nisInfo: { nextNIS: '261066', lastNIS: '261065' },
      totalPendaftarBaru: 0,
      totalAktifTahunIni: 0,
      totalPendaftar: 0,
      totalSemua: 0
    };
  }
}

/**
 * Verifikasi dan Aktivasi Pendaftaran Santri oleh Admin:
 * Memberikan NIS definitif, menetapkan kelas, dan mengaktifkan status santri
 */
function verifikasiPendaftaranSantri(payload) {
  try {
    const sheet = SHEETS.santri;
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const nisLama = String(payload.nisLama || payload.NIS || '').trim();
    const namaSantri = String(payload.Nama_Santri || payload.nama || '').trim();
    const nisBaru = String(payload.nisBaru || payload.NIS || '').trim();
    const kelasBaru = String(payload.Kelas || payload.kelas || '-').trim();
    const statusBaru = String(payload.Status || 'Aktif').trim();
    const tahunAjaran = String(payload.Tahun_Ajaran || '').trim();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      const rowNis = String(data[i][0]).trim();
      const rowNama = String(data[i][1]).trim();
      if ((nisLama && rowNis === nisLama) || (namaSantri && rowNama.toLowerCase() === namaSantri.toLowerCase())) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) {
      throw new Error('Data santri tidak ditemukan di database.');
    }
    
    const idxNIS = headers.findIndex(h => h.toLowerCase() === 'nis');
    const idxKelas = headers.findIndex(h => h.toLowerCase() === 'kelas');
    const idxStatus = headers.findIndex(h => h.toLowerCase() === 'status');
    const idxTahun = headers.findIndex(h => h.toLowerCase() === 'tahun_ajaran');
    
    if (idxNIS !== -1 && nisBaru) sheet.getRange(rowIndex, idxNIS + 1).setValue(nisBaru);
    if (idxKelas !== -1 && kelasBaru) sheet.getRange(rowIndex, idxKelas + 1).setValue(kelasBaru);
    if (idxStatus !== -1 && statusBaru) sheet.getRange(rowIndex, idxStatus + 1).setValue(statusBaru);
    if (idxTahun !== -1 && tahunAjaran) sheet.getRange(rowIndex, idxTahun + 1).setValue(tahunAjaran);
    
    if (payload.buatUser === true || statusBaru.toLowerCase() === 'aktif') {
      try {
        const idxWaAyah = headers.findIndex(h => h.toLowerCase().includes('wa_ayah'));
        const idxWaIbu = headers.findIndex(h => h.toLowerCase().includes('wa_ibu'));
        const noHp = (idxWaAyah !== -1 ? data[rowIndex - 1][idxWaAyah] : '') || (idxWaIbu !== -1 ? data[rowIndex - 1][idxWaIbu] : '');
        createUserSantri_(nisBaru || nisLama, namaSantri || data[rowIndex - 1][1], String(noHp));
      } catch (eU) {
        console.warn('Gagal buat user santri:', eU);
      }
    }
    
    return {
      success: true,
      message: 'Pendaftaran santri berhasil diverifikasi & status diubah menjadi ' + statusBaru + '!',
      nis: nisBaru || nisLama,
      kelas: kelasBaru,
      status: statusBaru
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || err.toString()
    };
  }
}

// ===== GURU =====

function getGuruList() {
  const data = SHEETS.guru.getDataRange().getValues();
  const headers = data[0];
  const guruList = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const guru = {};
    headers.forEach((header, index) => {
      let val = row[index];
      if (val instanceof Date) val = val.toISOString();
      guru[header] = val;
    });
    guruList.push(guru);
  }
  
  return guruList;
}

function saveGuru(data) {
  try {
    const sheet = SHEETS.guru;
    const headers = sheet.getDataRange().getValues()[0];
    
    const dataNorm = {};
    Object.keys(data).forEach(k => { dataNorm[k.toLowerCase()] = data[k]; });
    
    if (data.ID_Guru) {
      // Update
      const range = sheet.getDataRange().getValues();
      for (let i = 1; i < range.length; i++) {
        if (range[i][0] === data.ID_Guru) {
          headers.forEach((header, index) => {
            const hKey = header.toLowerCase();
            if (dataNorm[hKey] !== undefined) {
              sheet.getRange(i + 1, index + 1).setValue(dataNorm[hKey]);
            }
          });
          return { success: true, message: 'Data guru berhasil diupdate' };
        }
      }
    } else {
      // Insert
      const idGuru = 'G' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');
      dataNorm['id_guru'] = idGuru;
      dataNorm['created_at'] = new Date();
      if (!dataNorm['status']) dataNorm['status'] = 'aktif';
      
      const newRow = headers.map(header => {
        const hKey = header.toLowerCase();
        return dataNorm[hKey] !== undefined ? dataNorm[hKey] : '';
      });
      
      sheet.appendRow(newRow);
      return { success: true, message: 'Data guru berhasil ditambahkan', id_guru: idGuru };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== KELAS =====

function getKelasList() {
  const data = SHEETS.kelas.getDataRange().getValues();
  const headers = data[0];
  const kelasList = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const kelas = {};
    headers.forEach((header, index) => {
      let val = row[index];
      if (val instanceof Date) val = val.toISOString();
      kelas[header] = val;
    });
    kelasList.push(kelas);
  }
  
  return kelasList;
}

function isKelasMatch_(k1, k2) {
  if (!k1 || !k2) return false;
  const s1 = String(k1).trim().toUpperCase();
  const s2 = String(k2).trim().toUpperCase();
  if (s1 === s2) return true;
  
  const clean1 = s1.replace(/^JILID\s*/i, '').replace(/[^A-Z0-9]/g, '');
  const clean2 = s2.replace(/^JILID\s*/i, '').replace(/[^A-Z0-9]/g, '');
  return clean1 === clean2;
}

function matchMateriKelas_(materiKelasName, targetKelasName) {
  if (!materiKelasName || !targetKelasName) return false;
  const m = String(materiKelasName).trim().toUpperCase();
  const t = String(targetKelasName).trim().toUpperCase();
  
  if (m === t) return true;
  
  // Normalize strings (strip "JILID " and non-alphanumeric)
  const cleanM = m.replace(/^JILID\s*/i, '').replace(/[^A-Z0-9]/g, '');
  const cleanT = t.replace(/^JILID\s*/i, '').replace(/[^A-Z0-9]/g, '');
  
  if (cleanM === cleanT) return true;
  
  // Strip trailing section letter (e.g. 'A', 'B', 'C') from target class
  // e.g. "4A" -> "4", "1B" -> "1"
  const tBase = cleanT.replace(/[A-Z]$/, '');
  const mBase = cleanM.replace(/[A-Z]$/, '');
  if (tBase && (tBase === cleanM || tBase === mBase)) {
    return true;
  }
  
  // Keyword checks for non-numbered classes
  if (cleanT.includes('PRA') && cleanM.includes('PRA')) return true;
  if (cleanT.includes('TAHFIDH') && cleanM.includes('TAHFIDH')) return true;
  if (cleanT.includes('DINIYAH') && cleanM.includes('DINIYAH')) return true;
  if ((cleanT.includes('JUZAMMA') || cleanT.includes('TADARUS') || cleanT.includes('TAKHOSSUS')) &&
      (cleanM.includes('JUZAMMA') || cleanM.includes('TADARUS') || cleanM.includes('TAKHOSSUS'))) return true;
      
  return false;
}

function getDaftarKelasMateri() {
  const data = SHEETS.materiKelas.getDataRange().getValues();
  const set = new Set();
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][1] || '').trim();
    if (k) set.add(k);
  }
  return Array.from(set);
}

function getMateriByKelas(namaKelas) {
  if (!namaKelas) return [];
  const data = SHEETS.materiKelas.getDataRange().getValues();
  const materiList = [];
  const target = String(namaKelas).trim();
  
  // 1. Try exact match first
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (String(row[1] || '').trim().toUpperCase() === target.toUpperCase()) {
      materiList.push({
        id_materi: row[0],
        kelas: row[1],
        materi: row[2],
        bobot: row[3],
        urutan: row[4],
        tipe: row[5],
      });
    }
  }
  
  // 2. Fallback to smart fuzzy / parent-level match
  if (materiList.length === 0) {
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      if (matchMateriKelas_(row[1], target)) {
        materiList.push({
          id_materi: row[0],
          kelas: row[1],
          materi: row[2],
          bobot: row[3],
          urutan: row[4],
          tipe: row[5],
        });
      }
    }
  }
  
  // Sort by urutan
  materiList.sort((a, b) => (Number(a.urutan) || 0) - (Number(b.urutan) || 0));
  
  return materiList;
}

function saveMateriKelas(data) {
  try {
    const sheet = SHEETS.materiKelas;
    
    if (data.id_materi) {
      // Update
      const range = sheet.getDataRange().getValues();
      for (let i = 1; i < range.length; i++) {
        if (range[i][0] === data.id_materi) {
          sheet.getRange(i + 1, 2).setValue(data.kelas);
          sheet.getRange(i + 1, 3).setValue(data.materi);
          sheet.getRange(i + 1, 4).setValue(data.bobot);
          sheet.getRange(i + 1, 5).setValue(data.urutan);
          sheet.getRange(i + 1, 6).setValue(data.tipe);
          return { success: true, message: 'Materi berhasil diupdate' };
        }
      }
    } else {
      // Insert
      const idMateri = 'M' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');
      sheet.appendRow([
        idMateri,
        data.kelas,
        data.materi,
        data.bobot,
        data.urutan,
        data.tipe
      ]);
      return { success: true, message: 'Materi berhasil ditambahkan', id_materi: idMateri };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteMateriKelas(idMateri) {
  try {
    const sheet = SHEETS.materiKelas;
    const range = sheet.getDataRange().getValues();
    
    for (let i = 1; i < range.length; i++) {
      if (range[i][0] === idMateri) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Materi berhasil dihapus' };
      }
    }
    return { success: false, message: 'Materi tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== PRESENSI =====

function savePresensiSantri(data) {
  try {
    const sheet = SHEETS.presensiSantri;
    const newRow = [
      'PS' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss'),
      data.tanggal,
      data.nis,
      data.nama_santri,
      data.kelas,
      data.status, // S/I/A
      data.keterangan,
      data.input_oleh,
      new Date(),
      data.tahun_ajaran || '' // Tambahkan Tahun Ajaran di Kolom ke-10
    ];
    sheet.appendRow(newRow);
    return { success: true, message: 'Presensi berhasil disimpan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function savePresensiSantriBatch(dataArray) {
  try {
    if (!dataArray || dataArray.length === 0) return { success: true, message: 'Tidak ada data untuk disimpan' };
    
    const kelasInput = dataArray[0].kelas;
    const tanggalInput = dataArray[0].tanggal; // format yyyy-mm-dd
    const sheet = SHEETS.presensiSantri;
    
    // CEK DOUBLE ABSEN (Kecuali untuk kegiatan khusus yang bisa diisi berkali-kali)
    const kegiatanKhusus = ['RAPAT', 'ISTIGHOSAH', 'LAIN-LAIN'];
    const kelasUpper = kelasInput.toString().trim().toUpperCase();
    
    if (!kegiatanKhusus.includes(kelasUpper)) {
      const existingData = sheet.getDataRange().getValues();
      for (let i = 1; i < existingData.length; i++) {
        const row = existingData[i];
        if (row[4] && row[4].toString().trim().toUpperCase() === kelasUpper) {
          let rowTgl = '';
          if (row[1] instanceof Date) {
            rowTgl = Utilities.formatDate(row[1], 'GMT+7', 'yyyy-MM-dd');
          } else if (row[1]) {
            try { rowTgl = new Date(row[1]).toISOString().split('T')[0]; } catch(e) {}
          }
          if (rowTgl === tanggalInput) {
            throw new Error('GAGAL: Kelas ' + kelasInput + ' sudah diabsen pada tanggal tersebut! Tidak boleh dobel.');
          }
        }
      }
    }
    
    const rows = [];
    const now = new Date();
    
    for (let i = 0; i < dataArray.length; i++) {
      const data = dataArray[i];
      // Sleep slightly or format date to ensure unique ID if needed? 
      // Utilities.formatDate only gives up to seconds. 
      // So let's append milliseconds or loop index for safety.
      const id = 'PS' + Utilities.formatDate(now, 'GMT+7', 'yyyyMMddHHmmss') + i;
      
      rows.push([
        id,
        data.tanggal,
        data.nis,
        data.nama_santri,
        data.kelas,
        data.status,
        data.keterangan,
        data.input_oleh,
        now,
        data.tahun_ajaran || ''
      ]);
    }
    
    // Batch insert
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return { success: true, message: 'Presensi ' + rows.length + ' santri berhasil disimpan' };
  } catch (err) {
    throw new Error(err.message);
  }
}

function getPresensiSantriByTA(kelas, tahunAjaran) {
  const data = SHEETS.presensiSantri.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    // Bypass restriction, check all rows that have at least Nama or Kelas
    if (!row[3] && !row[4]) continue; 
    
    let matchKelas = !kelas || isKelasMatch_(row[4], kelas);
    let matchTA = !tahunAjaran || !row[9] || (row[9].toString().trim() === '') || (row[9].toString().trim() === tahunAjaran.toString().trim());
    
    if (matchKelas && matchTA) {
      result.push({
        id_presensi: row[0] || ('PS_old_' + i),
        tanggal: (row[1] instanceof Date) ? row[1].toISOString() : (row[1] || ''),
        nis: row[2] || '-',
        nama_santri: row[3] || '(Tanpa Nama)',
        kelas: row[4] || '-',
        status: row[5] || 'H',
        keterangan: row[6] || '',
        tahun_ajaran: row[9] || ''
      });
    }
  }
  return result;
}

/**
 * Mengambil rekap presensi (Hadir, Sakit, Izin, Alpa, Total SIA) per santri
 */
function getRekapPresensiSantriByNIS(nis, namaSantri, tahunAjaran) {
  try {
    const sheet = SHEETS.presensiSantri;
    const data = sheet.getDataRange().getValues();
    const targetNis = String(nis || '').trim();
    const targetNama = String(namaSantri || '').trim().toLowerCase();
    const targetTA = tahunAjaran ? String(tahunAjaran).trim() : '';
    
    function countFor(taFilter) {
      let h = 0, s = 0, iz = 0, a = 0, tot = 0;
      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const rowNis = String(row[2] || '').trim();
        const rowNama = String(row[3] || '').trim().toLowerCase();
        
        const matchNis = targetNis && rowNis === targetNis;
        const matchNama = targetNama && rowNama === targetNama;
        if (!matchNis && !matchNama) continue;
        
        const rowTA = String(row[9] || '').trim();
        if (taFilter && rowTA && rowTA !== taFilter) continue;
        
        const st = String(row[5] || 'H').toUpperCase().trim();
        tot++;
        if (st === 'S' || st.includes('SAKIT')) s++;
        else if (st === 'I' || st.includes('IZIN')) iz++;
        else if (st === 'A' || st.includes('ALPA') || st.includes('ALFA')) a++;
        else h++;
      }
      return { hadir: h, sakit: s, izin: iz, alpa: a, total: tot };
    }
    
    let res = targetTA ? countFor(targetTA) : countFor('');
    if (res.total === 0 && targetTA) {
      const allRes = countFor('');
      if (allRes.total > 0) res = allRes;
    }
    
    const siaTotal = res.sakit + res.izin + res.alpa;
    const persenKehadiran = res.total > 0 ? Math.round((res.hadir / res.total) * 100) : 100;
    
    return {
      success: true,
      nis: targetNis,
      namaSantri: namaSantri,
      tahunAjaran: targetTA,
      hadir: res.hadir,
      sakit: res.sakit,
      izin: res.izin,
      alpa: res.alpa,
      totalSIA: siaTotal,
      totalPertemuan: res.total,
      persenKehadiran: persenKehadiran
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || err.toString(),
      hadir: 0, sakit: 0, izin: 0, alpa: 0, totalSIA: 0, totalPertemuan: 0, persenKehadiran: 100
    };
  }
}

function savePresensiGuru(data) {
  try {
    const sheet = SHEETS.presensiGuru;
    const newRow = [
      'PG' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss'),
      data.tanggal,
      data.id_guru,
      data.nama_guru,
      data.kelas_kegiatan,
      data.jurnal_mengajar,
      'hadir',
      data.input_oleh,
      new Date(),
    ];
    sheet.appendRow(newRow);
    return { success: true, message: 'Presensi guru berhasil disimpan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== NILAI RAPOR =====

function saveNilaiRapor(data) {
  try {
    const sheet = SHEETS.nilaiRapor;
    
    // Check if already exists
    const range = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    const dMateri = String(data.materi || '').trim().toLowerCase();
    const dNis = String(data.nis || '').trim();
    const dTA = String(data.tahun_ajaran || '').trim();
    
    for (let i = 1; i < range.length; i++) {
      const rMateri = String(range[i][7] || '').trim().toLowerCase();
      const rNis = String(range[i][3] || '').trim();
      const rTA = String(range[i][1] || '').trim();
      if (rMateri === dMateri && rNis === dNis && rTA === dTA) {
        rowIndex = i + 1;
        break;
      }
    }
    
    const rowData = [
      data.id_nilai || 'N' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss'),
      data.tahun_ajaran,
      data.semester || '-',
      data.nis,
      data.nama_santri,
      data.kelas,
      data.id_materi || '',
      data.materi,
      data.bobot || 0,
      data.jumlah_kesalahan || 0,
      data.nilai_materi || 0,
      data.tipe || 'regular',
      data.wali_kelas || '',
      data.kepala_tpq || '',
      data.tanggal_cetak || '',
      data.input_oleh || '',
      data.perilaku || 'B',
      data.kedisiplinan || 'B',
      data.keaktifan || 'B',
      data.catatan_penguji || '',
      new Date(),
    ];
    
    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    return { success: true, message: 'Nilai berhasil disimpan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function saveBatchNilaiRapor(nilaiArray) {
  try {
    if (!Array.isArray(nilaiArray) || nilaiArray.length === 0) {
      return { success: false, message: 'Data nilai kosong' };
    }
    
    const sheet = SHEETS.nilaiRapor;
    const range = sheet.getDataRange().getValues();
    
    const existingMap = {};
    for (let i = 1; i < range.length; i++) {
      const rMateri = String(range[i][7] || '').trim().toLowerCase();
      const rNis = String(range[i][3] || '').trim();
      const rTA = String(range[i][1] || '').trim();
      if (rMateri && rNis) {
        existingMap[`${rMateri}_${rNis}_${rTA}`] = i + 1;
      }
    }
    
    const toAppend = [];
    const now = new Date();
    
    for (let k = 0; k < nilaiArray.length; k++) {
      const data = nilaiArray[k];
      const dMateri = String(data.materi || '').trim().toLowerCase();
      const dNis = String(data.nis || '').trim();
      const dTA = String(data.tahun_ajaran || '').trim();
      const key = `${dMateri}_${dNis}_${dTA}`;
      
      const rowData = [
        data.id_nilai || ('N' + Utilities.formatDate(now, 'GMT+7', 'yyyyMMddHHmmss') + '_' + k),
        data.tahun_ajaran,
        data.semester || '-',
        data.nis,
        data.nama_santri,
        data.kelas,
        data.id_materi || '',
        data.materi,
        data.bobot || 0,
        data.jumlah_kesalahan || 0,
        data.nilai_materi || 0,
        data.tipe || 'regular',
        data.wali_kelas || '',
        data.kepala_tpq || '',
        data.tanggal_cetak || '',
        data.input_oleh || '',
        data.perilaku || 'B',
        data.kedisiplinan || 'B',
        data.keaktifan || 'B',
        data.catatan_penguji || '',
        now,
      ];
      
      if (existingMap[key]) {
        const rowIndex = existingMap[key];
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        toAppend.push(rowData);
      }
    }
    
    if (toAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
    }
    
    return { success: true, message: 'Berhasil menyimpan ' + nilaiArray.length + ' data nilai rapor!' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getNilaiRaporByKelas(kelas, tahunAjaran) {
  const data = SHEETS.nilaiRapor.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    // Bypass ID restriction for manually entered data
    if (!row[1] && !row[5]) continue;
    let matchKelas = !kelas || isKelasMatch_(row[5], kelas);
    let matchTA = !tahunAjaran || (row[1] && row[1].toString().trim() === tahunAjaran.toString().trim());
    if (matchKelas && matchTA) {
      result.push({
        id_nilai: row[0],
        tahun_ajaran: row[1],
        semester: row[2],
        nis: row[3],
        nama_santri: row[4],
        kelas: row[5],
        id_materi: row[6],
        materi: row[7],
        bobot: row[8],
        jumlah_kesalahan: row[9],
        nilai_materi: row[10],
        tipe: row[11],
        perilaku: row[16],
        kedisiplinan: row[17],
        keaktifan: row[18],
        catatan_penguji: row[19],
      });
    }
  }
  
  return result;
}

function hitungNilaiRapor(nis, kelas, tahunAjaran) {
  const nilaiList = getNilaiRaporByKelas(kelas, tahunAjaran);
  const santriNilai = nilaiList.filter(n => String(n.nis || '').trim() === String(nis || '').trim());
  
  let totalKesalahan = 0;
  let nilaiLagu = 0;
  let weightedTotal = 0;
  let weightedMax = 0;
  let manualValues = [];
  
  santriNilai.forEach(nilai => {
    if (nilai.tipe === 'regular') {
      totalKesalahan += (nilai.bobot * nilai.jumlah_kesalahan);
    } else if (nilai.tipe === 'plus') {
      nilaiLagu = nilai.nilai_materi || 5;
    } else if (nilai.tipe === 'weighted') {
      weightedTotal += nilai.nilai_materi;
      weightedMax += nilai.bobot;
    } else if (nilai.tipe === 'manual') {
      manualValues.push(nilai.nilai_materi);
    }
  });
  
  let nilaiAkhir = 0;
  const upperKelas = String(kelas || '').toUpperCase();
  
  if (upperKelas.includes('TAHFIDH')) {
    // Weighted score
    nilaiAkhir = weightedTotal;
  } else if (upperKelas.includes('DINIYAH')) {
    // Average of manual values
    if (manualValues.length > 0) {
      nilaiAkhir = manualValues.reduce((a, b) => a + b, 0) / manualValues.length;
    }
  } else {
    // Regular formula: 95 - totalKesalahan + nilaiLagu
    nilaiAkhir = 95 - totalKesalahan + nilaiLagu;
  }
  
  // Minimum 70
  if (nilaiAkhir < 70) nilaiAkhir = 70;
  
  return {
    nilai_akhir: Math.round(nilaiAkhir),
    total_kesalahan: totalKesalahan,
    nilai_lagu: nilaiLagu,
    status: nilaiAkhir > 80 ? 'NAIK/LULUS' : 'MENGULANG',
  };
}

// ===== NILAI BUPRES (BUKU PRESTASI) =====

function getTatapMukaConfig(tahunAjaran, namaKelas) {
  try {
    const config = getConfig();
    const ta = String(tahunAjaran || config.TAHUN_AJARAN || '20262').trim();
    const kelasKey = String(namaKelas || '').trim();
    
    // 1. Cek spesifik per kelas: TATAP_MUKA_{TA}_{KELAS}
    if (kelasKey) {
      const specificKey = ('TATAP_MUKA_' + ta + '_' + kelasKey).toUpperCase().replace(/\s+/g, '_');
      if (config[specificKey] !== undefined && config[specificKey] !== '') {
        return {
          tatap_muka: Number(config[specificKey]) || 0,
          is_specific: true,
          key: specificKey
        };
      }
    }
    
    // 2. Cek global per tahun ajaran: TATAP_MUKA_{TA}
    const taKey = ('TATAP_MUKA_' + ta).toUpperCase();
    if (config[taKey] !== undefined && config[taKey] !== '') {
      return {
        tatap_muka: Number(config[taKey]) || 0,
        is_specific: false,
        key: taKey
      };
    }
    
    // 3. Fallback standar default TPQ
    const defaultVal = Number(config['TATAP_MUKA_DEFAULT']) || 24;
    return {
      tatap_muka: defaultVal,
      is_specific: false,
      key: 'DEFAULT'
    };
  } catch (e) {
    return { tatap_muka: 24, is_specific: false, error: e.toString() };
  }
}

function saveTatapMukaConfig(tahunAjaran, namaKelas, jumlahTatapMuka, applyAll) {
  try {
    const ta = String(tahunAjaran || '').trim();
    const jml = Number(jumlahTatapMuka) || 0;
    if (jml <= 0) {
      return { success: false, message: 'Jumlah tatap muka harus lebih dari 0' };
    }
    
    if (applyAll) {
      const taKey = ('TATAP_MUKA_' + ta).toUpperCase();
      saveConfig(taKey, jml);
      saveConfig('TATAP_MUKA_DEFAULT', jml);
      return { success: true, message: 'Jumlah tatap muka (' + jml + ') berhasil disimpan untuk SEMUA kelas di TA ' + ta };
    } else {
      const kelasKey = String(namaKelas || '').trim();
      if (!kelasKey) {
        return { success: false, message: 'Pilih nama kelas terlebih dahulu' };
      }
      const specificKey = ('TATAP_MUKA_' + ta + '_' + kelasKey).toUpperCase().replace(/\s+/g, '_');
      saveConfig(specificKey, jml);
      return { success: true, message: 'Jumlah tatap muka (' + jml + ') berhasil disimpan untuk kelas ' + kelasKey + ' (TA ' + ta + ')' };
    }
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getBupresByKelas(kelas, tahunAjaran) {
  try {
    ensureSheetHeaders_();
    const sheet = SHEETS.nilaiBupres;
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    const targetTA = String(tahunAjaran || '').trim();
    const targetKelas = String(kelas || '').trim();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[3] && !row[4]) continue; // NIS & Nama kosong
      
      const rTA = String(row[1] || '').trim();
      const rKelas = String(row[2] || '').trim();
      
      const matchTA = !targetTA || rTA === targetTA;
      const matchKelas = !targetKelas || rKelas === targetKelas || isKelasMatch_(rKelas, targetKelas);
      
      if (matchTA && matchKelas) {
        result.push({
          id_bupres: row[0],
          tahun_ajaran: row[1],
          kelas: row[2],
          nis: String(row[3] || '').trim(),
          nama_santri: row[4],
          b: Number(row[5]) || 0,
          b_minus: Number(row[6]) || 0,
          c_plus: Number(row[7]) || 0,
          c: Number(row[8]) || 0,
          c_minus: Number(row[9]) || 0,
          k: Number(row[10]) || 0,
          total_nilai: Number(row[11]) || 0,
          tatap_muka: Number(row[12]) || 0,
          poin_bck: Number(row[13]) || 0,
          catatan: row[14] || '',
          input_oleh: row[15] || '',
          updated_at: row[16] instanceof Date ? row[16].toISOString() : String(row[16] || '')
        });
      }
    }
    
    return result;
  } catch (e) {
    return [];
  }
}

function saveBatchBupres(bupresArray) {
  try {
    ensureSheetHeaders_();
    if (!Array.isArray(bupresArray) || bupresArray.length === 0) {
      return { success: false, message: 'Data Bupres kosong' };
    }
    
    const sheet = SHEETS.nilaiBupres;
    const range = sheet.getDataRange().getValues();
    
    const existingMap = {};
    for (let i = 1; i < range.length; i++) {
      const rTA = String(range[i][1] || '').trim();
      const rNis = String(range[i][3] || '').trim();
      if (rTA && rNis) {
        existingMap[rTA + '_' + rNis] = i + 1;
      }
    }
    
    const now = new Date();
    const toAppend = [];
    
    for (let idx = 0; idx < bupresArray.length; idx++) {
      const d = bupresArray[idx];
      const ta = String(d.tahun_ajaran || '').trim();
      const nis = String(d.nis || '').trim();
      const key = ta + '_' + nis;
      
      const b = Math.max(0, parseInt(d.b) || 0);
      const bMinus = Math.max(0, parseInt(d.b_minus) || 0);
      const cPlus = Math.max(0, parseInt(d.c_plus) || 0);
      const c = Math.max(0, parseInt(d.c) || 0);
      const cMinus = Math.max(0, parseInt(d.c_minus) || 0);
      const k = Math.max(0, parseInt(d.k) || 0);
      
      const totalEntries = b + bMinus + cPlus + c + cMinus + k;
      const tatapMuka = Math.max(1, parseInt(d.tatap_muka) || 24);
      if (tatapMuka > 0 && totalEntries > tatapMuka) {
        return {
          success: false,
          message: 'Validasi Gagal: Total kategori nilai untuk ' + (d.nama_santri || d.nis) + ' (' + totalEntries + ') melebihi batas jumlah tatap muka (' + tatapMuka + '). Silakan periksa kembali!'
        };
      }
      
      const totalNilai = (b * 95) + (bMinus * 90) + (cPlus * 83) + (c * 80) + (cMinus * 73) + (k * 70);
      const poinBCK = Number((totalNilai / tatapMuka).toFixed(2));
      
      const idBupres = d.id_bupres || ('BP' + Utilities.formatDate(now, 'GMT+7', 'yyyyMMddHHmmss') + '_' + idx);
      
      const rowData = [
        idBupres,
        ta,
        d.kelas || '',
        nis,
        d.nama_santri || '',
        b,
        bMinus,
        cPlus,
        c,
        cMinus,
        k,
        totalNilai,
        tatapMuka,
        poinBCK,
        d.catatan || '',
        d.input_oleh || '',
        now
      ];
      
      if (existingMap[key]) {
        const rowIndex = existingMap[key];
        sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
      } else {
        toAppend.push(rowData);
      }
    }
    
    if (toAppend.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, toAppend[0].length).setValues(toAppend);
    }
    
    return { 
      success: true, 
      message: 'Berhasil menyimpan ' + bupresArray.length + ' data Buku Prestasi (Bupres)!' 
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== PEMBAYARAN =====

function savePembayaran(data) {
  try {
    const sheet = SHEETS.pembayaran;
    const idPembayaran = 'TR' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');
    
    const newRow = [
      idPembayaran,
      data.tanggal,
      data.nis,
      data.nama_santri,
      data.kelas,
      data.jenis_tagihan,
      data.periode,
      data.nominal,
      data.status,
      data.metode,
      data.keterangan,
      data.input_oleh,
      new Date(),
    ];
    sheet.appendRow(newRow);
    
    return { success: true, message: 'Pembayaran berhasil disimpan', id: idPembayaran };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getPembayaranBySantri(nis) {
  const data = SHEETS.pembayaran.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (row[2] === nis) {
      result.push({
        id_pembayaran: row[0],
        tanggal: row[1] instanceof Date ? row[1].toISOString() : row[1],
        nis: row[2],
        nama_santri: row[3],
        kelas: row[4],
        jenis_tagihan: row[5],
        periode: row[6],
        nominal: row[7],
        status: row[8],
        metode: row[9],
      });
    }
  }
  
  return result;
}

function getRekapPembayaran() {
  const data = SHEETS.pembayaran.getDataRange().getValues();
  const rekap = {
    total_pemasukan: 0,
    per_bulan: {},
    per_jenis: {},
  };
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (row[8] === 'lunas') {
      const nominal = parseFloat(row[7]) || 0;
      const tanggal = new Date(row[1]);
      const bulanKey = Utilities.formatDate(tanggal, 'GMT+7', 'yyyy-MM');
      const jenisKey = row[5];
      
      rekap.total_pemasukan += nominal;
      rekap.per_bulan[bulanKey] = (rekap.per_bulan[bulanKey] || 0) + nominal;
      rekap.per_jenis[jenisKey] = (rekap.per_jenis[jenisKey] || 0) + nominal;
    }
  }
  
  return rekap;
}

// ===== PENGELUARAN =====

function savePengeluaran(data) {
  try {
    const sheet = SHEETS.pengeluaran;
    const idPengeluaran = 'OUT' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');
    
    const newRow = [
      idPengeluaran,
      data.tanggal,
      data.kategori,
      data.deskripsi,
      data.nominal,
      data.bukti,
      data.input_oleh,
      new Date(),
    ];
    sheet.appendRow(newRow);
    
    return { success: true, message: 'Pengeluaran berhasil disimpan', id: idPengeluaran };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getTotalSaldo() {
  const pemasukanData = SHEETS.pembayaran.getDataRange().getValues();
  const pemasukanExtra = SHEETS.pemasukan.getLastRow() > 1 ? SHEETS.pemasukan.getDataRange().getValues() : [];
  const pengeluaranData = SHEETS.pengeluaran.getDataRange().getValues();
  
  let totalPemasukan = 0;
  let totalPengeluaran = 0;
  
  // Sum pemasukan dari Pembayaran (lunas)
  for (let i = 1; i < pemasukanData.length; i++) {
    if (!pemasukanData[i][0]) continue;
    if (pemasukanData[i][8] === 'lunas') {
      totalPemasukan += parseFloat(pemasukanData[i][7]) || 0;
    }
  }
  // Sum pemasukan dari Pemasukan (non-SPP)
  for (let i = 1; i < pemasukanExtra.length; i++) {
    if (!pemasukanExtra[i][0]) continue;
    totalPemasukan += parseFloat(pemasukanExtra[i][4]) || 0;
  }
  
  // Sum pengeluaran
  for (let i = 1; i < pengeluaranData.length; i++) {
    if (!pengeluaranData[i][0]) continue;
    totalPengeluaran += parseFloat(pengeluaranData[i][4]) || 0;
  }
  
  return {
    pemasukan: totalPemasukan,
    pengeluaran: totalPengeluaran,
    saldo: totalPemasukan - totalPengeluaran,
  };
}

// ===== JENIS TAGIHAN & KONFIGURASI =====

function getJenisTagihan() {
  ensureSheetHeaders_();
  const data = SHEETS.jenisTagihan.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) result.push(data[i][0]);
  }
  return result;
}

function saveJenisTagihan(jenis) {
  try {
    ensureSheetHeaders_();
    const data = SHEETS.jenisTagihan.getDataRange().getValues();
    const existing = data.slice(1).map(r => r[0]);
    if (!existing.includes(jenis)) {
      SHEETS.jenisTagihan.appendRow([jenis]);
    }
    return { success: true, list: getJenisTagihan() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function savePemasukan(data) {
  try {
    ensureSheetHeaders_();
    const id = 'IN' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');
    SHEETS.pemasukan.appendRow([id, data.tanggal, data.sumber, data.keterangan, data.nominal, data.input_oleh, new Date()]);
    return { success: true, message: 'Pemasukan berhasil disimpan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getAllPemasukan() {
  if (SHEETS.pemasukan.getLastRow() <= 1) return [];
  const data = SHEETS.pemasukan.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    result.push({
      id: row[0],
      tanggal: row[1] instanceof Date ? row[1].toISOString() : row[1],
      sumber: row[2],
      keterangan: row[3],
      nominal: row[4],
    });
  }
  return result;
}

function getTagihanKonfigurasi(tahunAjaran) {
  ensureSheetHeaders_();
  const data = SHEETS.tagihanConfig.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1] && !row[2]) continue; // Skip jika Tahun Ajaran dan Jenis Tagihan kosong
    if (!tahunAjaran || row[1] === tahunAjaran) {
      result.push({ id: row[0] || 'manual', tahun_ajaran: row[1], jenis_tagihan: row[2], nominal: row[3], keterangan: row[4] });
    }
  }
  return result;
}

function saveTagihanKonfigurasi(data) {
  try {
    ensureSheetHeaders_();
    const id = 'TK' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');
    SHEETS.tagihanConfig.appendRow([id, data.tahun_ajaran, data.jenis_tagihan, data.nominal, data.keterangan]);
    return { success: true, message: 'Tagihan konfigurasi berhasil disimpan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteTagihanKonfigurasi(id) {
  try {
    const data = SHEETS.tagihanConfig.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) { SHEETS.tagihanConfig.deleteRow(i + 1); return { success: true, message: 'Dihapus' }; }
    }
    return { success: false, message: 'Tidak ditemukan' };
  } catch (e) { return { success: false, message: e.toString() }; }
}

function getTunggakanByTA(tahunAjaran) {
  // Ambil konfigurasi tagihan wajib untuk TA tsb
  const configs = getTagihanKonfigurasi(tahunAjaran);
  if (!configs.length) return { configs: [], tunggakan: [] };
  
  const pembayaranData = SHEETS.pembayaran.getDataRange().getValues();
  const santriData = SHEETS.santri.getDataRange().getValues();
  
  // Ambil index kolom secara dinamis dari header Santri
  const sHeaders = santriData[0].map(h => h.toLowerCase());
  const idxNIS = sHeaders.indexOf('nis');
  const idxNama = sHeaders.indexOf('nama_santri');
  const idxKelas = sHeaders.indexOf('kelas');
  const idxStatus = sHeaders.indexOf('status');
  
  // Jika kolom belum ada atau error, kembalikan kosong
  if (idxNIS === -1 || idxStatus === -1) return { configs, tunggakan: [] };
  
  // Filter hanya santri aktif
  const santriAktif = santriData.slice(1).filter(r => r[0] && (r[idxStatus] || '').toString().toLowerCase() === 'aktif');
  
  const tunggakan = [];
  for (const s of santriAktif) {
    const nis = s[idxNIS], namaSantri = s[idxNama], kelas = s[idxKelas];
    for (const cfg of configs) {
      const sudahBayar = pembayaranData.slice(1).some(p =>
        p[2] === nis && p[5] === cfg.jenis_tagihan && p[6] === tahunAjaran && (p[8] || '').toString().toLowerCase() === 'lunas'
      );
      if (!sudahBayar) {
        tunggakan.push({ nis, nama_santri: namaSantri, kelas, jenis_tagihan: cfg.jenis_tagihan, nominal: cfg.nominal, tahun_ajaran: tahunAjaran });
      }
    }
  }
  return { configs, tunggakan };
}

function getGrafikData() {
  const pembData = SHEETS.pembayaran.getDataRange().getValues();
  const pemasukanExtra = SHEETS.pemasukan.getLastRow() > 1 ? SHEETS.pemasukan.getDataRange().getValues() : [];
  const pengeluaranData = SHEETS.pengeluaran.getDataRange().getValues();
  
  const pemasukanPerBulan = {};
  const pengeluaranPerBulan = {};
  
  for (let i = 1; i < pembData.length; i++) {
    if (!pembData[i][0]) continue;
    if (pembData[i][8] === 'lunas') {
      const tgl = new Date(pembData[i][1]);
      if (isNaN(tgl)) continue;
      const key = Utilities.formatDate(tgl, 'GMT+7', 'MMM yyyy');
      pemasukanPerBulan[key] = (pemasukanPerBulan[key] || 0) + (parseFloat(pembData[i][7]) || 0);
    }
  }
  for (let i = 1; i < pemasukanExtra.length; i++) {
    if (!pemasukanExtra[i][0]) continue;
    const tgl = new Date(pemasukanExtra[i][1]);
    if (isNaN(tgl)) continue;
    const key = Utilities.formatDate(tgl, 'GMT+7', 'MMM yyyy');
    pemasukanPerBulan[key] = (pemasukanPerBulan[key] || 0) + (parseFloat(pemasukanExtra[i][4]) || 0);
  }
  for (let i = 1; i < pengeluaranData.length; i++) {
    if (!pengeluaranData[i][0]) continue;
    const tgl = new Date(pengeluaranData[i][1]);
    if (isNaN(tgl)) continue;
    const key = Utilities.formatDate(tgl, 'GMT+7', 'MMM yyyy');
    pengeluaranPerBulan[key] = (pengeluaranPerBulan[key] || 0) + (parseFloat(pengeluaranData[i][4]) || 0);
  }
  
  const labels = [...new Set([...Object.keys(pemasukanPerBulan), ...Object.keys(pengeluaranPerBulan)])];
  return {
    labels,
    pemasukan: labels.map(l => pemasukanPerBulan[l] || 0),
    pengeluaran: labels.map(l => pengeluaranPerBulan[l] || 0),
  };
}

// ===== PENGUMUMAN =====

function getPengumuman(target) {
  const data = SHEETS.pengumuman.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    if (row[4] === target || row[4] === 'semua') {
      result.push({
        id_pengumuman: row[0],
        tanggal: row[1] instanceof Date ? row[1].toISOString() : row[1],
        judul: row[2],
        isi: row[3],
        target: row[4],
        prioritas: row[5],
        status: row[6],
        dibuat_oleh: row[7],
      });
    }
  }
  
  return result;
}

function savePengumuman(data) {
  try {
    const sheet = SHEETS.pengumuman;
    const idPengumuman = 'PGM' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');
    
    const newRow = [
      idPengumuman,
      new Date(),
      data.judul,
      data.isi,
      data.target,
      data.prioritas,
      'aktif',
      data.dibuat_oleh,
      new Date(),
    ];
    sheet.appendRow(newRow);
    
    return { success: true, message: 'Pengumuman berhasil disimpan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

// ===== CATATAN KESEHARIAN =====

function saveCatatanKeseharian(data) {
  try {
    const sheet = SHEETS.catatanKeseharian;
    const idCatatan = 'CK' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss');
    
    const newRow = [
      idCatatan,
      new Date(),
      data.nis,
      data.nama_santri,
      data.kelas,
      data.catatan,
      data.input_oleh,
      new Date(),
    ];
    sheet.appendRow(newRow);
    
    return { success: true, message: 'Catatan berhasil disimpan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getCatatanBySantri(nis) {
  const data = SHEETS.catatanKeseharian.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[2] === nis) {
      result.push({
        tanggal: row[1] instanceof Date ? row[1].toISOString() : row[1],
        catatan: row[5],
        input_oleh: row[6],
      });
    }
  }
  
  return result;
}

// ===== EXPORT EXCEL =====

function exportToExcel(sheetName, data) {
  try {
    // Data should be array of objects
    if (!data || data.length === 0) {
      return { success: false, message: 'No data to export' };
    }
    
    const headers = Object.keys(data[0]);
    const csvData = [
      headers.join(','),
      ...data.map(row => headers.map(h => `"${row[h] || ''}"`).join(','))
    ];
    
    const csvString = csvData.join('\n');
    const blob = Utilities.newBlob(csvString, 'text/csv', sheetName + '.csv');
    
    // Create file in Drive
    const file = DriveApp.createFile(blob);
    file.setName(sheetName + '_' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMdd_HHmmss') + '.xlsx');
    
    return {
      success: true,
      downloadUrl: file.getDownloadUrl(),
      fileId: file.getId(),
    };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
// ===== TAMBAHAN FUNGSI YANG BELUM ADA =====

function getDashboardStats() {
  const santriList = getSantriList();
  let santriCount = 0;
  for (let i = 0; i < santriList.length; i++) {
    if ((santriList[i].Status || 'aktif').toLowerCase() === 'aktif') {
      santriCount++;
    }
  }
  const guruCount = Math.max(0, SHEETS.guru.getLastRow() - 1);
  const saldo = getTotalSaldo();
  
  return { santri: santriCount, guru: guruCount, saldo: saldo };
}

function getAllPembayaran() {
  const data = SHEETS.pembayaran.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    result.push({
      id_pembayaran: row[0],
      tanggal: row[1] instanceof Date ? row[1].toISOString() : row[1],
      nis: row[2],
      nama_santri: row[3],
      kelas: row[4],
      jenis_tagihan: row[5],
      periode: row[6],
      nominal: row[7],
      status: row[8],
      metode: row[9],
    });
  }
  return result;
}

function getAllPengeluaran() {
  const data = SHEETS.pengeluaran.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    result.push({
      id_pengeluaran: row[0],
      tanggal: row[1] instanceof Date ? row[1].toISOString() : row[1],
      kategori: row[2],
      deskripsi: row[3],
      nominal: row[4],
    });
  }
  return result;
}

function getTunggakan() {
  const data = SHEETS.pembayaran.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[8] === 'belum') {
      result.push({
        nis: row[2],
        nama_santri: row[3],
        kelas: row[4],
        jenis_tagihan: row[5],
        periode: row[6],
        nominal: row[7],
      });
    }
  }
  return result;
}

function getPresensiGuruByMonth(bulan, tahun) {
  const data = SHEETS.presensiGuru.getDataRange().getValues();
  const result = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    const tglRaw = row[1];
    const tgl = tglRaw instanceof Date ? tglRaw : new Date(tglRaw);
    if (isNaN(tgl)) continue;
    
    if (tgl.getMonth() + 1 === parseInt(bulan) && 
        tgl.getFullYear() === parseInt(tahun)) {
      result.push({
        tanggal: tglRaw instanceof Date ? tglRaw.toISOString() : tglRaw,
        nama_guru: row[3],
        kelas_kegiatan: row[4],
        jurnal_mengajar: row[5],
      });
    }
  }
  return result;
}

function getPresensiGuruByGuru(namaGuru, bulan, tahun) {
  const data = SHEETS.presensiSantri.getDataRange().getValues();
  const result = [];
  const uniqueSessions = new Set();
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[0]) continue;
    
    const rowNamaGuru = row[7] ? row[7].toString().trim() : '';
    if (namaGuru && rowNamaGuru !== namaGuru) continue;
    
    const tglRaw = row[1];
    const tgl = tglRaw instanceof Date ? tglRaw : new Date(tglRaw);
    if (isNaN(tgl)) continue;
    
    const rowBulan = tgl.getMonth() + 1;
    const rowTahun = tgl.getFullYear();
    if (bulan && rowBulan !== parseInt(bulan)) continue;
    if (tahun && rowTahun !== parseInt(tahun)) continue;
    
    const dateStr = tglRaw instanceof Date ? tglRaw.toISOString().split('T')[0] : tglRaw.toString().split('T')[0];
    const kelas = row[4] ? row[4].toString().trim() : '';
    const upperKelas = kelas.toUpperCase();
    let key = dateStr + '_' + rowNamaGuru + '_' + kelas;
    
    if (upperKelas === 'RAPAT' || upperKelas === 'ISTIGHOSAH' || upperKelas === 'LAIN-LAIN') {
       key += '_' + i;
    }
    
    if (!uniqueSessions.has(key)) {
      uniqueSessions.add(key);
      result.push({
        tanggal: tglRaw instanceof Date ? tglRaw.toISOString() : tglRaw,
        nama_guru: rowNamaGuru,
        kelas_kegiatan: kelas,
        jurnal_mengajar: row[6] ? row[6].toString().trim() : '-',
      });
    }
  }
  return { total: result.length, data: result };
}

function getJenisKegiatanGuru() {
  ensureSheetHeaders_();
  const manualData = SHEETS.jenisKegiatanGuru.getDataRange().getValues();
  const result = new Set();
  
  for (let i = 1; i < manualData.length; i++) {
    if (manualData[i][0]) result.add(manualData[i][0].toString().trim());
  }
  
  const santriData = SHEETS.santri.getDataRange().getValues();
  for (let i = 1; i < santriData.length; i++) {
    const row = santriData[i];
    if (row[3] && row[7]) { 
      result.add(row[7].toString().trim());
    }
  }
  
  return Array.from(result).sort();
}

function saveJenisKegiatanGuru(jenis) {
  try {
    ensureSheetHeaders_();
    const data = SHEETS.jenisKegiatanGuru.getDataRange().getValues();
    const existing = data.slice(1).map(r => r[0]);
    if (!existing.includes(jenis)) {
      SHEETS.jenisKegiatanGuru.appendRow([jenis]);
    }
    return { success: true, list: getJenisKegiatanGuru() };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function deleteGuru(id) {
  try {
    const sheet = SHEETS.guru;
    const range = sheet.getDataRange().getValues();
    
    for (let i = 1; i < range.length; i++) {
      if (range[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data guru berhasil dihapus' };
      }
    }
    return { success: false, message: 'Guru tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}
/**
 * MENGAMBIL DATA PEMBAYARAN DARI SPREADSHEET (APLIKASI LAMA)
 */
function getPaymentStatus() {
  try {
    const ss = SpreadsheetApp.openById('1a2KFMWS6pUh3qZfZqz5BkUanVc0X4FD2__sRHDF4MmY');
    
    const sheet20262 = ss.getSheetByName('INPUT');
    const sheet20261 = ss.getSheetByName('rekap 20261');
    const sheet20252 = ss.getSheetByName('rekap 20252');
    
    if (!sheet20262 || !sheet20261 || !sheet20252) {
      throw new Error('Sheet "INPUT", "rekap 20261", atau "rekap 20252" tidak ditemukan.');
    }
    
    function parseHeaders(sheet) {
      const data = sheet.getDataRange().getValues();
      const headers = data[0];
      const cols = [];
      let limit = headers.findIndex(h => String(h).toLowerCase().includes('induk'));
      if (limit === -1) limit = headers.length;
      for (let i = 3; i < limit; i++) {
        let val = String(headers[i]).trim();
        if (val && !val.toLowerCase().includes('tanggal')) cols.push(val);
      }
      return { data, headers, cols };
    }

    const p20262 = parseHeaders(sheet20262);
    const p20261 = parseHeaders(sheet20261);
    const p20252 = parseHeaders(sheet20252);
    
    const studentsMap = {}; 

    function processSheet(dataObj, tagihanKey) {
      const data = dataObj.data;
      const headers = dataObj.headers;
      const columns = dataObj.cols;
      for (let i = 1; i < data.length; i++) {
        if (data[i][1]) { 
          const nama = data[i][1].toString().trim();
          const namaKey = nama.toLowerCase();

          if (!studentsMap[namaKey]) {
            studentsMap[namaKey] = {
              no: data[i][0],
              nama: nama,
              kelas: data[i][2] ? data[i][2].toString().trim() : '',
              tagihan20262: {},
              tagihan20261: {},
              tagihan20252: {}
            };
          }
          
          columns.forEach(bulan => {
            let bulanIndex = -1;
            for(let k = 0; k < headers.length; k++) {
               if(headers[k] && headers[k].toString().trim().toLowerCase() === bulan.toLowerCase()) {
                   bulanIndex = k;
                   break;
               }
            }

            if (bulanIndex !== -1) {
              let tanggalIndex = -1;
              for (let j = bulanIndex + 1; j < headers.length; j++) {
                if (headers[j] && headers[j].toString().toLowerCase().includes('tanggal')) {
                  tanggalIndex = j;
                  break;
                }
              }
              
              studentsMap[namaKey][tagihanKey][bulan] = {
                status: data[i][bulanIndex] ? data[i][bulanIndex].toString().trim() : '',
                tanggal: tanggalIndex !== -1 ? (data[i][tanggalIndex] ? data[i][tanggalIndex].toString().trim() : '') : ''
              };
            }
          });
        }
      }
    }

    processSheet(p20262, 'tagihan20262');
    processSheet(p20261, 'tagihan20261');
    processSheet(p20252, 'tagihan20252');
    
    return {
      students: Object.values(studentsMap),
      headers20262: p20262.cols,
      headers20261: p20261.cols,
      headers20252: p20252.cols
    };
  } catch (error) {
    console.error('Error:', error);
    return { error: error.message || error.toString() };
  }
}

/**
 * Update status pembayaran ke LUNAS
 */
function updatePayment(nama, tipeTagihan, bulan, tanggalPembayaran) {
  try {
    const ss = SpreadsheetApp.openById('1a2KFMWS6pUh3qZfZqz5BkUanVc0X4FD2__sRHDF4MmY');
    
    let sheetName = '';
    if (tipeTagihan === '20262') sheetName = 'INPUT';
    else if (tipeTagihan === '20261') sheetName = 'rekap 20261';
    else if (tipeTagihan === '20252') sheetName = 'rekap 20252';
    
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Sheet "${sheetName}" tidak ditemukan`);
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let foundRow = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toLowerCase() === nama.trim().toLowerCase()) {
        foundRow = i;
        break;
      }
    }
    
    if (foundRow === -1) throw new Error('Siswa tidak ditemukan');
    
    // Konversi parameter bulan menjadi array (mendukung single atau multiple select)
    let bulanArray = Array.isArray(bulan) ? bulan : [bulan];
    
    for (let b = 0; b < bulanArray.length; b++) {
      const namaBulan = bulanArray[b];
      let bulanIndex = -1;
      
      for (let i = 0; i < headers.length; i++) {
        if (headers[i] && headers[i].toString().trim().toLowerCase() === namaBulan.toLowerCase()) {
          bulanIndex = i;
          break;
        }
      }
      
      if (bulanIndex === -1) continue; // Skip jika kolom bulan ini tidak ditemukan
      
      let tanggalIndex = -1;
      for (let j = bulanIndex + 1; j < headers.length; j++) {
        if (headers[j] && headers[j].toString().toLowerCase().includes('tanggal pembayaran')) {
          tanggalIndex = j;
          break;
        }
      }
      
      sheet.getRange(foundRow + 1, bulanIndex + 1).setValue('LUNAS');
      if (tanggalIndex !== -1) {
         sheet.getRange(foundRow + 1, tanggalIndex + 1).setValue(tanggalPembayaran);
      }
    }
    
    return { success: true, message: 'Pembayaran berhasil disimpan' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}


function getJurnalKelas() {
  const data = SHEETS.presensiSantri.getDataRange().getValues();
  const latestByKelas = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const kelas = row[4] ? row[4].toString().trim() : '';
    const upperKelas = kelas.toUpperCase();
    
    if (kelas && upperKelas !== 'RAPAT' && upperKelas !== 'ISTIGHOSAH' && upperKelas !== 'LAIN-LAIN') {
       const tanggal = row[1];
       let dateVal = 0;
       
       if (tanggal instanceof Date) {
          dateVal = tanggal.getTime();
       } else if (tanggal) {
          let parsed = new Date(tanggal);
          if (isNaN(parsed.getTime())) {
             const parts = tanggal.toString().split(/[-/]/);
             if (parts.length === 3) {
                if (parts[0].length === 4) {
                   parsed = new Date(parts[0] + '-' + parts[1] + '-' + parts[2]);
                } else {
                   parsed = new Date(parts[2] + '-' + parts[1] + '-' + parts[0]);
                }
             }
          }
          if (!isNaN(parsed.getTime())) dateVal = parsed.getTime();
          else dateVal = 1; 
       } else {
          dateVal = 1; 
       }
       
       if (dateVal > 0) {
          const jurnal = row[6] ? row[6].toString().trim() : '';
          if (jurnal && (!latestByKelas[upperKelas] || latestByKelas[upperKelas].dateVal < dateVal)) {
             latestByKelas[upperKelas] = {
                kelas: kelas,
                materi: jurnal,
                pengajar: row[7] ? row[7].toString().trim() : '',
                tanggal: (tanggal && typeof tanggal.getTime === 'function') ? tanggal.toISOString() : tanggal,
                dateVal: dateVal
             };
          }
       }
    }
  }
  
  return Object.values(latestByKelas).sort((a, b) => a.kelas.localeCompare(b.kelas));
}



function getUniqueKelasFromPresensi() {
  const data = SHEETS.presensiSantri.getDataRange().getValues();
  const unique = new Set();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[3] && row[4]) {
      unique.add(row[4].toString().trim());
    }
  }
  return Array.from(unique).sort();
}


function debugPresensiSantri() {
  const data = SHEETS.presensiSantri.getDataRange().getValues();
  let count = 0;
  let matches = 0;
  let missingName = 0;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row[3]) count++;
    else missingName++;
    
    if (row[9] && row[9].toString().trim() === '20261') {
       matches++;
    }
  }
  return 'Total rows: ' + data.length + ', Valid names: ' + count + ', Missing names: ' + missingName + ', TA matches: ' + matches;
}






// ===== REKAP DATA KELAS =====
function getRekapDataKelas() {
  // 1. Ambil data kelas
  const dataKelas = SHEETS.kelas.getDataRange().getValues();
  const kelasInfo = {};
  const rekapList = [];
  const kelasHeaders = dataKelas[0].map(h => h.toString().trim().toLowerCase());
  const idxLokasiKelas = kelasHeaders.findIndex(h => h === 'lokasi kelas' || h === 'lokasi_kelas' || h === 'lokasi');
  
  for (let i = 1; i < dataKelas.length; i++) {
    const row = dataKelas[i];
    if (!row[0]) continue;
    const namaKelas = row[1] ? row[1].toString().trim() : '';
    const waliKelas = row[5] ? row[5].toString().trim() : '';
    const lokasiKelas = idxLokasiKelas !== -1 && row[idxLokasiKelas] ? row[idxLokasiKelas].toString().trim() : '-';
    
    const upperName = namaKelas.toUpperCase();
    if (namaKelas && upperName !== 'RAPAT' && upperName !== 'ISTIGHOSAH' && upperName !== 'LAIN-LAIN') {
      kelasInfo[namaKelas] = {
        nama_kelas: namaKelas,
        wali_kelas: waliKelas,
        lokasi_kelas: lokasiKelas,
        jml_santri: 0,
        jml_l: 0,
        jml_p: 0,
        jurnal_terakhir: '-'
      };
      rekapList.push(namaKelas);
    }
  }

  // 2. Ambil data santri untuk hitung L/P
  const dataSantri = SHEETS.santri.getDataRange().getValues();
  // Assume: row[3] is Kelas, row[2] is L/P. Let's check getSantriList headers.
  // Actually, header mapping might be different. Let's find index by header name.
  const santriHeaders = dataSantri[0].map(h => h.toString().toLowerCase());
  const idxKelas = santriHeaders.findIndex(h => h.includes('kelas'));
  const idxLP = santriHeaders.findIndex(h => h.includes('kelamin') || h === 'l/p' || h === 'jk');
  const idxStatus = santriHeaders.findIndex(h => h === 'status' || h === 'status santri');
  
  if (idxKelas !== -1 && idxLP !== -1) {
    for (let i = 1; i < dataSantri.length; i++) {
      const row = dataSantri[i];
      if (!row[0] && !row[1] && !row[3]) continue; // skip empty
      
      const status = idxStatus !== -1 && row[idxStatus] ? row[idxStatus].toString().trim().toLowerCase() : 'aktif';
      if (status !== 'aktif') continue;
      
      const kelas = row[idxKelas] ? row[idxKelas].toString().trim() : '';
      const jk = row[idxLP] ? row[idxLP].toString().trim().toUpperCase() : '';
      
      if (kelas && kelasInfo[kelas]) {
        kelasInfo[kelas].jml_santri++;
        if (jk === 'L' || jk === 'LAKI-LAKI') kelasInfo[kelas].jml_l++;
        else if (jk === 'P' || jk === 'PEREMPUAN') kelasInfo[kelas].jml_p++;
      }
    }
  }

  // 3. Ambil data jurnal terakhir dari Presensi_Santri
  const dataPresensi = SHEETS.presensiSantri.getDataRange().getValues();
  const latestJurnal = {};
  
  for (let i = 1; i < dataPresensi.length; i++) {
    const row = dataPresensi[i];
    const kelas = row[4] ? row[4].toString().trim() : '';
    if (kelas && kelasInfo[kelas]) {
      const tanggal = row[1];
      let dateVal = 0;
      if (tanggal instanceof Date) dateVal = tanggal.getTime();
      else if (tanggal) {
        let parsed = new Date(tanggal);
        if (!isNaN(parsed.getTime())) dateVal = parsed.getTime();
        else dateVal = 1; 
      }
      
      const jurnal = row[6] ? row[6].toString().trim() : '';
      if (jurnal && (!latestJurnal[kelas] || latestJurnal[kelas].dateVal <= dateVal)) {
        latestJurnal[kelas] = { dateVal: dateVal, jurnal: jurnal };
      }
    }
  }
  
  // Gabungkan hasil
  const finalResult = [];
  rekapList.forEach(k => {
    if (latestJurnal[k]) {
      kelasInfo[k].jurnal_terakhir = latestJurnal[k].jurnal;
    }
    finalResult.push(kelasInfo[k]);
  });
  
  return finalResult;
}

// ===== CEKLIS SANTRI =====

function getCeklisKonfigurasi(role) {
  try {
    ensureSheetHeaders_();
    const sheet = SHEETS.ceklisConfig;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const list = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      
      let daftarKolom = [];
      try {
        daftarKolom = typeof row[3] === 'string' && row[3].startsWith('[') ? JSON.parse(row[3]) : String(row[3] || '').split(',').map(s => s.trim()).filter(Boolean);
      } catch (e) {
        daftarKolom = String(row[3] || '').split(',').map(s => s.trim()).filter(Boolean);
      }
      
      let targetKelas = [];
      try {
        targetKelas = typeof row[4] === 'string' && row[4].startsWith('[') ? JSON.parse(row[4]) : (String(row[4] || '') === 'SEMUA' ? ['SEMUA'] : String(row[4] || '').split(',').map(s => s.trim()).filter(Boolean));
      } catch (e) {
        targetKelas = ['SEMUA'];
      }
      
      const status = String(row[5] || 'aktif').toLowerCase();
      if (role === 'ustadz' && status !== 'aktif') {
        continue;
      }
      
      list.push({
        id_ceklis: String(row[0]),
        nama_ceklis: String(row[1] || ''),
        deskripsi: String(row[2] || ''),
        daftar_kolom: daftarKolom,
        target_kelas: targetKelas,
        status: status,
        created_at: row[6] instanceof Date ? row[6].toISOString() : String(row[6] || '')
      });
    }
    
    return list;
  } catch (err) {
    Logger.log('Error getCeklisKonfigurasi: ' + err.toString());
    return [];
  }
}

function saveCeklisKonfigurasi(payload) {
  try {
    ensureSheetHeaders_();
    const sheet = SHEETS.ceklisConfig;
    const data = sheet.getDataRange().getValues();
    
    const idCeklis = payload.id_ceklis || ('CKL_' + Utilities.formatDate(new Date(), 'GMT+7', 'yyyyMMddHHmmss'));
    const namaCeklis = String(payload.nama_ceklis || '').trim();
    if (!namaCeklis) throw new Error('Nama/Topik ceklis tidak boleh kosong');
    
    const deskripsi = String(payload.deskripsi || '').trim();
    const daftarKolom = JSON.stringify(Array.isArray(payload.daftar_kolom) ? payload.daftar_kolom.map(k => String(k).trim()).filter(Boolean) : []);
    const targetKelas = JSON.stringify(Array.isArray(payload.target_kelas) ? payload.target_kelas : ['SEMUA']);
    const status = String(payload.status || 'aktif').toLowerCase();
    
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === idCeklis) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex !== -1) {
      sheet.getRange(rowIndex, 2, 1, 5).setValues([[namaCeklis, deskripsi, daftarKolom, targetKelas, status]]);
    } else {
      sheet.appendRow([idCeklis, namaCeklis, deskripsi, daftarKolom, targetKelas, status, new Date()]);
    }
    
    return { success: true, message: 'Konfigurasi Ceklis berhasil disimpan', id_ceklis: idCeklis };
  } catch (err) {
    return { success: false, message: err.message || err.toString() };
  }
}

function deleteCeklisKonfigurasi(idCeklis) {
  try {
    ensureSheetHeaders_();
    const sheet = SHEETS.ceklisConfig;
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(idCeklis)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Ceklis berhasil dihapus' };
      }
    }
    return { success: false, message: 'Ceklis tidak ditemukan' };
  } catch (err) {
    return { success: false, message: err.message || err.toString() };
  }
}

function getKelasForUstadz(userNama, userId, userRole, userUsername) {
  try {
    const kelasData = SHEETS.kelas.getDataRange().getValues();
    if (kelasData.length <= 1) return [];
    
    const headers = kelasData[0].map(h => String(h || '').trim().toLowerCase());
    const idxNamaKelas = headers.findIndex(h => h === 'nama kelas' || h === 'nama_kelas' || h.includes('nama'));
    const idxWali = headers.findIndex(h => h.includes('wali') || h.includes('ustadz') || h.includes('guru'));
    
    const allKelas = [];
    for (let i = 1; i < kelasData.length; i++) {
      const row = kelasData[i];
      const nKelas = idxNamaKelas !== -1 ? String(row[idxNamaKelas] || '').trim() : String(row[1] || '').trim();
      const wKelas = idxWali !== -1 ? String(row[idxWali] || '').trim() : String(row[5] || row[3] || '').trim();
      if (nKelas) {
        allKelas.push({
          nama_kelas: nKelas,
          wali: wKelas
        });
      }
    }
    
    if (userRole === 'super_admin' || userRole === 'admin') {
      return allKelas.map(k => k.nama_kelas);
    }
    
    const searchNama = String(userNama || '').trim().toLowerCase();
    const searchId = String(userId || '').trim().toLowerCase();
    const searchUser = String(userUsername || '').trim().toLowerCase();
    const matchedKelas = new Set();
    
    // 1. Cek dari Sheet Kelas
    allKelas.forEach(k => {
      const wali = k.wali.toLowerCase();
      if (wali) {
        if (wali === searchNama || wali === searchId || wali === searchUser ||
            (searchNama && (wali.includes(searchNama) || searchNama.includes(wali))) ||
            (searchId && (wali === searchId || wali.includes(searchId))) ||
            (searchUser && (wali === searchUser || wali.includes(searchUser)))) {
          matchedKelas.add(k.nama_kelas);
        }
      }
    });
    
    // 2. Cek dari Sheet Guru
    if (SHEETS.guru) {
      const guruData = SHEETS.guru.getDataRange().getValues();
      const gHeaders = guruData[0].map(h => String(h || '').trim().toLowerCase());
      const idxGNama = gHeaders.findIndex(h => h === 'nama' || h.includes('nama'));
      const idxGId = gHeaders.findIndex(h => h === 'id_guru' || h.includes('id'));
      const idxGKelas = gHeaders.findIndex(h => h.includes('kelas') || h.includes('diampu'));
      
      if (idxGKelas !== -1) {
        for (let i = 1; i < guruData.length; i++) {
          const gRow = guruData[i];
          const gNama = idxGNama !== -1 ? String(gRow[idxGNama] || '').trim().toLowerCase() : '';
          const gId = idxGId !== -1 ? String(gRow[idxGId] || '').trim().toLowerCase() : '';
          
          if ((searchNama && (gNama === searchNama || gNama.includes(searchNama) || searchNama.includes(gNama))) ||
              (searchId && gId === searchId) || (searchUser && gId === searchUser)) {
            const diampuRaw = String(gRow[idxGKelas] || '');
            diampuRaw.split(',').forEach(item => {
              const trimmed = item.trim();
              if (trimmed && trimmed !== '-') matchedKelas.add(trimmed);
            });
          }
        }
      }
    }
    
    const resultList = Array.from(matchedKelas);
    // Jika tidak ada kelas yang ter-match (misal beda ejaan nama), tampilkan semua kelas agar Ustadz tetap bisa memilih
    if (resultList.length === 0) {
      return allKelas.map(k => k.nama_kelas);
    }
    
    return resultList;
  } catch (err) {
    Logger.log('Error getKelasForUstadz: ' + err.toString());
    return [];
  }
}

function getMonitoringCeklis(idCeklis) {
  try {
    ensureSheetHeaders_();
    
    // 1. Ambil config ceklis
    const configs = getCeklisKonfigurasi('super_admin');
    if (!configs || configs.length === 0) {
      return { success: false, message: 'Belum ada formulir ceklis yang dibuat.' };
    }
    
    let targetConfig = null;
    if (idCeklis) {
      targetConfig = configs.find(c => String(c.id_ceklis) === String(idCeklis));
    }
    if (!targetConfig) {
      targetConfig = configs[0];
    }
    
    const targetIdCeklis = String(targetConfig.id_ceklis);
    const daftarKolom = targetConfig.daftar_kolom || [];
    const jmlKolom = daftarKolom.length;
    
    // 2. Ambil semua kelas dari Sheet Kelas
    const kelasData = SHEETS.kelas.getDataRange().getValues();
    const kHeaders = (kelasData[0] || []).map(h => String(h || '').trim().toLowerCase());
    const idxNamaK = kHeaders.findIndex(h => h === 'nama kelas' || h === 'nama_kelas' || h.includes('nama'));
    const idxWaliK = kHeaders.findIndex(h => h.includes('wali') || h.includes('ustadz') || h.includes('guru'));
    
    const kelasList = [];
    for (let i = 1; i < kelasData.length; i++) {
      const row = kelasData[i];
      const nK = idxNamaK !== -1 ? String(row[idxNamaK] || '').trim() : String(row[1] || '').trim();
      const wK = idxWaliK !== -1 ? String(row[idxWaliK] || '').trim() : String(row[5] || row[3] || '').trim();
      if (nK) {
        kelasList.push({
          nama_kelas: nK,
          wali_kelas: wK || '-'
        });
      }
    }
    
    // Filter target kelas berdasarkan targetConfig.target_kelas
    let relevantClasses = kelasList;
    const targetKConfig = targetConfig.target_kelas || ['SEMUA'];
    const isSemua = targetKConfig.includes('SEMUA') || targetKConfig.length === 0;
    if (!isSemua) {
      relevantClasses = kelasList.filter(k => {
        return targetKConfig.some(tk => isKelasMatch_(k.nama_kelas, tk) || k.nama_kelas.toLowerCase() === String(tk).toLowerCase());
      });
    }
    
    // 3. Ambil data santri aktif
    const santriData = SHEETS.santri.getDataRange().getValues();
    const sHeaders = santriData[0] || [];
    const idxSNis = sHeaders.findIndex(h => {
      const clean = String(h || '').trim().toLowerCase();
      return clean === 'nis' || clean.includes('nis');
    });
    const idxSNama = sHeaders.findIndex(h => {
      const clean = String(h || '').trim().toLowerCase();
      return clean === 'nama_santri' || clean === 'nama santri' || clean.includes('nama');
    });
    const idxSKelas = sHeaders.findIndex(h => {
      const clean = String(h || '').trim().toLowerCase();
      return clean === 'kelas' || clean.includes('kelas');
    });
    const idxSStatus = sHeaders.findIndex(h => {
      const clean = String(h || '').trim().toLowerCase();
      return clean === 'status';
    });
    
    // Kelompokkan santri aktif per kelas
    const santriByKelas = {};
    relevantClasses.forEach(k => {
      santriByKelas[k.nama_kelas] = [];
    });
    
    for (let i = 1; i < santriData.length; i++) {
      const row = santriData[i];
      if (!row[0] && !row[1] && !row[3]) continue;
      const status = idxSStatus !== -1 ? String(row[idxSStatus] || 'aktif').trim().toLowerCase() : 'aktif';
      if (status !== 'aktif' && status !== '') continue;
      
      const sKelas = idxSKelas !== -1 ? String(row[idxSKelas] || '').trim() : '';
      const sNis = idxSNis !== -1 ? String(row[idxSNis] || '').trim() : String(row[0] || '').trim();
      const sNama = idxSNama !== -1 ? String(row[idxSNama] || '').trim() : String(row[1] || '').trim();
      if (!sNis && !sNama) continue;
      
      for (let j = 0; j < relevantClasses.length; j++) {
        const kName = relevantClasses[j].nama_kelas;
        if (sKelas === kName || isKelasMatch_(sKelas, kName)) {
          santriByKelas[kName].push({ nis: sNis, nama: sNama });
          break;
        }
      }
    }
    
    // 4. Ambil data ceklis yang tersimpan
    const cDataSheet = SHEETS.ceklisData;
    const cDataRows = cDataSheet.getDataRange().getValues();
    
    const savedDataMap = {};
    const latestUpdateByKelas = {};
    
    for (let i = 1; i < cDataRows.length; i++) {
      const row = cDataRows[i];
      const rIdCeklis = String(row[1] || '');
      if (rIdCeklis !== targetIdCeklis) continue;
      
      const rKelas = String(row[2] || '').trim();
      const rNis = String(row[3] || '').trim();
      let rHasil = {};
      try {
        rHasil = typeof row[5] === 'string' && row[5].startsWith('{') ? JSON.parse(row[5]) : {};
      } catch (e) {
        rHasil = {};
      }
      const rUpdatedBy = String(row[7] || '');
      const rUpdatedAt = row[8] instanceof Date ? row[8].toISOString() : String(row[8] || '');
      
      savedDataMap[rNis] = {
        hasil: rHasil,
        updated_by: rUpdatedBy,
        updated_at: rUpdatedAt,
        kelas: rKelas
      };
      
      for (let j = 0; j < relevantClasses.length; j++) {
        const kName = relevantClasses[j].nama_kelas;
        if (rKelas === kName || isKelasMatch_(rKelas, kName)) {
          if (!latestUpdateByKelas[kName] || (rUpdatedAt && rUpdatedAt > latestUpdateByKelas[kName].updated_at)) {
            latestUpdateByKelas[kName] = {
              updated_by: rUpdatedBy,
              updated_at: rUpdatedAt
            };
          }
          break;
        }
      }
    }
    
    // 5. Kalkulasi metrik per kelas
    const perKelasList = [];
    let grandTotalSantri = 0;
    let grandTotalSantriTuntas = 0;
    let grandTotalSantriTerisi = 0;
    let grandTotalItemTarget = 0;
    let grandTotalItemTerceklis = 0;
    let totalKelasTuntas = 0;
    let totalKelasSebagian = 0;
    let totalKelasBelum = 0;
    
    relevantClasses.forEach(k => {
      const kName = k.nama_kelas;
      const santriList = santriByKelas[kName] || [];
      const totalSantri = santriList.length;
      const targetItems = totalSantri * jmlKolom;
      
      let totalItemTerceklis = 0;
      let santriTuntasCount = 0;
      let santriTerisiCount = 0;
      
      santriList.forEach(s => {
        const sData = savedDataMap[s.nis];
        if (sData && sData.hasil) {
          let checkedForSantri = 0;
          daftarKolom.forEach(col => {
            if (sData.hasil[col] === true || sData.hasil[col] === 'true' || sData.hasil[col] === 1) {
              checkedForSantri++;
            }
          });
          totalItemTerceklis += checkedForSantri;
          if (checkedForSantri > 0) santriTerisiCount++;
          if (jmlKolom > 0 && checkedForSantri === jmlKolom) santriTuntasCount++;
        }
      });
      
      const persentase = targetItems > 0 ? Math.round((totalItemTerceklis / targetItems) * 100) : 0;
      
      let statusKelas = 'Belum Diisi';
      if (totalSantri === 0) {
        statusKelas = 'Tidak Ada Santri';
      } else if (persentase === 100) {
        statusKelas = 'Tuntas';
        totalKelasTuntas++;
      } else if (persentase > 0 || santriTerisiCount > 0) {
        statusKelas = 'Sebagian';
        totalKelasSebagian++;
      } else {
        statusKelas = 'Belum Diisi';
        totalKelasBelum++;
      }
      
      const lastUpdate = latestUpdateByKelas[kName] || { updated_by: '-', updated_at: '' };
      
      perKelasList.push({
        nama_kelas: kName,
        wali_kelas: k.wali_kelas || '-',
        total_santri: totalSantri,
        santri_terisi: santriTerisiCount,
        santri_tuntas: santriTuntasCount,
        total_kolom: jmlKolom,
        target_item: targetItems,
        terceklis_item: totalItemTerceklis,
        persentase: persentase,
        status: statusKelas,
        updated_by: lastUpdate.updated_by || '-',
        updated_at: lastUpdate.updated_at || ''
      });
      
      grandTotalSantri += totalSantri;
      grandTotalSantriTuntas += santriTuntasCount;
      grandTotalSantriTerisi += santriTerisiCount;
      grandTotalItemTarget += targetItems;
      grandTotalItemTerceklis += totalItemTerceklis;
    });
    
    perKelasList.sort((a, b) => a.nama_kelas.localeCompare(b.nama_kelas));
    
    const overallPersentase = grandTotalItemTarget > 0 ? Math.round((grandTotalItemTerceklis / grandTotalItemTarget) * 100) : 0;
    
    return {
      success: true,
      config: {
        id_ceklis: targetConfig.id_ceklis,
        nama_ceklis: targetConfig.nama_ceklis,
        deskripsi: targetConfig.deskripsi,
        daftar_kolom: daftarKolom,
        total_kolom: jmlKolom,
        status: targetConfig.status
      },
      all_configs: configs.map(c => ({ id_ceklis: c.id_ceklis, nama_ceklis: c.nama_ceklis, status: c.status })),
      ringkasan: {
        total_kelas: relevantClasses.length,
        kelas_tuntas: totalKelasTuntas,
        kelas_sebagian: totalKelasSebagian,
        kelas_belum: totalKelasBelum,
        total_santri: grandTotalSantri,
        santri_terisi: grandTotalSantriTerisi,
        santri_tuntas: grandTotalSantriTuntas,
        target_item: grandTotalItemTarget,
        terceklis_item: grandTotalItemTerceklis,
        overall_persentase: overallPersentase
      },
      detail_kelas: perKelasList
    };
  } catch (err) {
    Logger.log('Error getMonitoringCeklis: ' + err.toString());
    return {
      success: false,
      message: err.message || err.toString()
    };
  }
}

function getCeklisData(idCeklis, namaKelas) {
  try {
    ensureSheetHeaders_();
    
    const santriData = SHEETS.santri.getDataRange().getValues();
    const santriHeaders = santriData[0] || [];
    const idxNis = santriHeaders.findIndex(h => {
      const clean = String(h || '').trim().toLowerCase();
      return clean === 'nis' || clean.includes('nis');
    });
    const idxNama = santriHeaders.findIndex(h => {
      const clean = String(h || '').trim().toLowerCase();
      return clean === 'nama_santri' || clean === 'nama santri' || clean.includes('nama');
    });
    const idxKelas = santriHeaders.findIndex(h => {
      const clean = String(h || '').trim().toLowerCase();
      return clean === 'kelas' || clean.includes('kelas');
    });
    const idxStatus = santriHeaders.findIndex(h => {
      const clean = String(h || '').trim().toLowerCase();
      return clean === 'status';
    });
    
    const targetSantri = [];
    for (let i = 1; i < santriData.length; i++) {
      const row = santriData[i];
      const sKelas = idxKelas !== -1 ? String(row[idxKelas] || '').trim() : '';
      const sStatus = idxStatus !== -1 ? String(row[idxStatus] || 'aktif').trim().toLowerCase() : 'aktif';
      
      if ((sKelas === String(namaKelas).trim() || isKelasMatch_(sKelas, namaKelas)) && (sStatus === 'aktif' || sStatus === '')) {
        const sNama = (idxNama !== -1 ? String(row[idxNama] || '').trim() : '') || String(row[1] || '').trim();
        const sNis = idxNis !== -1 ? String(row[idxNis] || '').trim() : String(row[0] || '').trim();
        targetSantri.push({
          nis: sNis,
          nama: sNama,
          kelas: sKelas,
          hasil: {},
          catatan: ''
        });
      }
    }
    
    targetSantri.sort((a, b) => a.nama.localeCompare(b.nama));
    
    const cDataSheet = SHEETS.ceklisData;
    const cDataRows = cDataSheet.getDataRange().getValues();
    const mapHasil = {};
    const mapCatatan = {};
    
    for (let i = 1; i < cDataRows.length; i++) {
      const row = cDataRows[i];
      const rowIdCeklis = String(row[1] || '');
      const rowKelas = String(row[2] || '');
      const rowNis = String(row[3] || '');
      
      if (rowIdCeklis === String(idCeklis) && rowKelas === String(namaKelas)) {
        let hasilObj = {};
        try {
          hasilObj = typeof row[5] === 'string' && row[5].startsWith('{') ? JSON.parse(row[5]) : {};
        } catch (e) {
          hasilObj = {};
        }
        mapHasil[rowNis] = hasilObj;
        mapCatatan[rowNis] = String(row[6] || '');
      }
    }
    
    targetSantri.forEach(s => {
      if (mapHasil[s.nis]) s.hasil = mapHasil[s.nis];
      if (mapCatatan[s.nis]) s.catatan = mapCatatan[s.nis];
    });
    
    return {
      success: true,
      santriList: targetSantri
    };
  } catch (err) {
    return {
      success: false,
      message: err.message || err.toString(),
      santriList: []
    };
  }
}

function saveCeklisData(idCeklis, namaKelas, dataList, updatedBy) {
  try {
    ensureSheetHeaders_();
    if (!idCeklis || !namaKelas || !Array.isArray(dataList)) {
      throw new Error('Data parameter tidak lengkap');
    }
    
    const sheet = SHEETS.ceklisData;
    const allData = sheet.getDataRange().getValues();
    const existingIndexMap = {};
    
    for (let i = 1; i < allData.length; i++) {
      const rowId = String(allData[i][0] || '');
      if (rowId) {
        existingIndexMap[rowId] = i + 1;
      }
    }
    
    const now = new Date();
    const newRows = [];
    
    dataList.forEach(item => {
      const nis = String(item.nis || '').trim();
      if (!nis) return;
      
      const rowId = `${idCeklis}_${nis}`;
      const hasilJson = JSON.stringify(item.hasil || {});
      const catatan = String(item.catatan || '').trim();
      const namaSantri = String(item.nama || '').trim();
      
      if (existingIndexMap[rowId]) {
        const rIndex = existingIndexMap[rowId];
        sheet.getRange(rIndex, 6, 1, 4).setValues([[hasilJson, catatan, updatedBy || 'Ustadz', now]]);
      } else {
        newRows.push([
          rowId,
          idCeklis,
          namaKelas,
          nis,
          namaSantri,
          hasilJson,
          catatan,
          updatedBy || 'Ustadz',
          now
        ]);
      }
    });
    
    if (newRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, 9).setValues(newRows);
    }
    
    return { success: true, message: 'Data Ceklis Kelas berhasil disimpan!' };
  } catch (err) {
    return { success: false, message: err.message || err.toString() };
  }
}

// ===== API ENDPOINT UNTUK GITHUB PAGES =====
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
       throw new Error("No payload");
    }
    
    var payload = JSON.parse(e.postData.contents);
    var actionName = payload.action;
    var args = payload.args || [];
    
    // Call the function dynamically
    if (typeof this[actionName] === 'function') {
      var result = this[actionName].apply(this, args);
      return ContentService.createTextOutput(JSON.stringify({data: result}))
                           .setMimeType(ContentService.MimeType.JSON);
    } else {
      throw new Error("Function " + actionName + " not found or not a function.");
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({error: error.message || error.toString()}))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
// ===========================================
