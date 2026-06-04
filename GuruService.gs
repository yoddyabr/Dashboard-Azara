// ==========================================
// GURUSERVICE.GS - CRUD Master Data Guru & Data Honor Guru
// v2: 1 guru = 1 ID, kelas disimpan comma-separated per baris mapel
// ==========================================

var CACHE_KEY_GURU = 'azara_guruData';

function invalidateGuruCache() {
  CacheService.getScriptCache().remove(CACHE_KEY_GURU);
}

// ==========================================
// 1. GET DATA GURU + HONOR (with cache)
// ==========================================
function getGuruData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY_GURU);
  if (cached) return JSON.parse(cached);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var guruList = [];
  var honorList = [];

  try {
    var sheetGuru = ss.getSheetByName(CONFIG.SHEET_GURU);
    if (sheetGuru) {
      var dataGuru = sheetGuru.getDataRange().getValues();
      for (var i = 1; i < dataGuru.length; i++) {
        if (!dataGuru[i][0] || dataGuru[i][0].toString().trim() === "") continue;
        guruList.push({
          rowIndex: i + 1,
          id   : dataGuru[i][0].toString().trim(),
          nama : dataGuru[i][1] ? dataGuru[i][1].toString().trim() : "",
          mapel: dataGuru[i][2] ? dataGuru[i][2].toString().trim() : "",
          kelas: dataGuru[i][3] ? dataGuru[i][3].toString().trim() : ""
          // kelas: comma-separated string, e.g. "SD 1, SD 2, SD 3"
        });
      }
    }
  } catch(e) { Logger.log("Error sheet guru: " + e.toString()); }

  // Honor sekarang ada di Data Mata Pelajaran (lihat MasterDataService.gs)
  var result = { guruList: guruList, honorList: honorList };
  try { cache.put(CACHE_KEY_GURU, JSON.stringify(result), 120); } catch(e) {}
  return result;
}

// ==========================================
// 2. SIMPAN GURU
// Payload:
//   - rowIndex  : isi jika EDIT baris tertentu
//   - idGuru    : isi jika tambah mapel ke guru YANG SUDAH ADA
//   - namaBaru  : nama guru baru (jika guru baru)
//   - mapel     : string mata pelajaran
//   - kelasArr  : array kelas yang dipilih, e.g. ["SD 1","SD 2"]
// ==========================================
function simpanGuru(payload) {
  // ── NEW: Multi-mapel mode (untuk tambah guru baru / tambah multiple mapel ke guru existing) ──
  if (payload.mapelGroups && Array.isArray(payload.mapelGroups) && payload.mapelGroups.length > 0 && !payload.rowIndex) {
    return simpanGuruMultiMapel_(payload);
  }
  if (!payload.mapel   || payload.mapel.toString().trim()   === "") return "❌ Mata pelajaran tidak boleh kosong!";
  if (!payload.kelasArr || payload.kelasArr.length === 0)           return "❌ Pilih minimal satu kelas!";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGuru = ss.getSheetByName(CONFIG.SHEET_GURU);
  if (!sheetGuru) {
    sheetGuru = ss.insertSheet(CONFIG.SHEET_GURU);
    sheetGuru.appendRow(["ID Guru","Nama Guru","Mata Pelajaran","Kelas"]);
  }
  // Paksa kolom Kelas (D) jadi format TEKS supaya "10, 11, 12" tidak diubah jadi tanggal
  sheetGuru.getRange(1, 4, sheetGuru.getMaxRows(), 1).setNumberFormat('@');

  var dataGuru  = sheetGuru.getDataRange().getValues();
  var kelasStr  = payload.kelasArr.join(', ');

  // ── MODE EDIT baris yang sudah ada ──
  if (payload.rowIndex && parseInt(payload.rowIndex) > 1) {
    var row = parseInt(payload.rowIndex);
    sheetGuru.getRange(row, 3).setValue(payload.mapel);
    sheetGuru.getRange(row, 4).setValue(kelasStr);
    invalidateGuruCache();
    return "✅ Data berhasil diperbarui!";
  }

  // ── MODE: tambah mapel ke guru yang sudah ada ──
  if (payload.idGuru && payload.idGuru.toString().trim() !== "") {
    var idTarget = payload.idGuru.toString().trim();
    var namaGuru = "";
    var existingRow = -1;

    for (var i = 1; i < dataGuru.length; i++) {
      if (!dataGuru[i][0] || dataGuru[i][0].toString().trim() !== idTarget) continue;
      if (!namaGuru) namaGuru = dataGuru[i][1] ? dataGuru[i][1].toString().trim() : "";
      // Cek apakah mapel ini sudah ada untuk guru ini
      if (dataGuru[i][2] && dataGuru[i][2].toString().trim().toLowerCase() === payload.mapel.toLowerCase()) {
        existingRow = i + 1;
      }
    }
    if (!namaGuru) return "❌ ID Guru tidak ditemukan.";

    if (existingRow > 1) {
      // Update kelas di baris yang sudah ada
      sheetGuru.getRange(existingRow, 4).setValue(kelasStr);
      invalidateGuruCache();
      return "✅ Kelas untuk mapel " + payload.mapel + " (" + namaGuru + ") berhasil diperbarui!";
    }
    // Tambah baris baru dengan ID yang sama
    sheetGuru.appendRow([idTarget, namaGuru, payload.mapel, kelasStr]);
    invalidateGuruCache();
    return "✅ Mata pelajaran " + payload.mapel + " berhasil ditambahkan untuk " + namaGuru + "!";
  }

  // ── MODE: guru baru ──
  if (!payload.namaBaru || payload.namaBaru.toString().trim() === "") return "❌ Nama guru tidak boleh kosong!";

  var maxId = 0;
  for (var m = 1; m < dataGuru.length; m++) {
    if (dataGuru[m][0]) {
      var cleanId = parseInt(dataGuru[m][0].toString().replace(/\D/g, ""));
      if (!isNaN(cleanId) && cleanId > maxId) maxId = cleanId;
    }
  }
  var nextId = "GRU-" + ("000" + (maxId + 1)).slice(-4);
  sheetGuru.appendRow([nextId, payload.namaBaru.toString().trim(), payload.mapel, kelasStr]);
  invalidateGuruCache();
  return "✅ Guru baru berhasil ditambahkan dengan ID: " + nextId;
}

// ==========================================

// ==========================================
// Multi-mapel: 1 guru, N (mapel,kelas) groups dalam 1 call
// ==========================================
function simpanGuruMultiMapel_(payload) {
  var validGroups = payload.mapelGroups.filter(function(g) {
    return g.mapel && g.mapel.toString().trim() && g.kelasArr && g.kelasArr.length > 0;
  });
  if (validGroups.length === 0) return "❌ Tambahkan minimal 1 mata pelajaran dengan kelas!";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGuru = ss.getSheetByName(CONFIG.SHEET_GURU);
  if (!sheetGuru) {
    sheetGuru = ss.insertSheet(CONFIG.SHEET_GURU);
    sheetGuru.appendRow(["ID Guru","Nama Guru","Mata Pelajaran","Kelas"]);
  }
  // Paksa kolom Kelas (D) jadi format TEKS supaya "10, 11, 12" tidak diubah jadi tanggal
  sheetGuru.getRange(1, 4, sheetGuru.getMaxRows(), 1).setNumberFormat('@');

  var dataGuru = sheetGuru.getDataRange().getValues();
  var idGuru   = payload.idGuru ? payload.idGuru.toString().trim() : "";
  var namaGuru = "";

  if (idGuru) {
    // Guru existing — cari namanya
    for (var i = 1; i < dataGuru.length; i++) {
      if (dataGuru[i][0] && dataGuru[i][0].toString().trim() === idGuru) {
        namaGuru = dataGuru[i][1] ? dataGuru[i][1].toString().trim() : "";
        break;
      }
    }
    if (!namaGuru) return "❌ ID Guru " + idGuru + " tidak ditemukan!";
  } else {
    // Guru baru — generate ID
    if (!payload.namaBaru || payload.namaBaru.toString().trim() === "") return "❌ Nama guru tidak boleh kosong!";
    namaGuru = payload.namaBaru.toString().trim();
    var maxId = 0;
    for (var m = 1; m < dataGuru.length; m++) {
      if (dataGuru[m][0]) {
        var cleanId = parseInt(dataGuru[m][0].toString().replace(/\D/g, ""));
        if (!isNaN(cleanId) && cleanId > maxId) maxId = cleanId;
      }
    }
    idGuru = "GRU-" + ("000" + (maxId + 1)).slice(-4);
  }

  var added = 0, updated = 0;
  validGroups.forEach(function(g) {
    var mapel = g.mapel.toString().trim();
    var kelasStr = g.kelasArr.join(', ');
    var existingRow = -1;
    for (var i = 1; i < dataGuru.length; i++) {
      if (dataGuru[i][0] && dataGuru[i][0].toString().trim() === idGuru &&
          dataGuru[i][2] && dataGuru[i][2].toString().trim().toLowerCase() === mapel.toLowerCase()) {
        existingRow = i + 1; break;
      }
    }
    if (existingRow > 0) {
      sheetGuru.getRange(existingRow, 4).setValue(kelasStr);
      updated++;
    } else {
      sheetGuru.appendRow([idGuru, namaGuru, mapel, kelasStr]);
      added++;
    }
  });

  invalidateGuruCache();
  var parts = [];
  if (added   > 0) parts.push(added + " mapel baru");
  if (updated > 0) parts.push(updated + " mapel diperbarui");
  return "✅ " + namaGuru + " (" + idGuru + "): " + parts.join(", ") + "!";
}

// 3. HAPUS SATU BARIS DATA GURU
// ==========================================
function hapusGuru(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGuru = ss.getSheetByName(CONFIG.SHEET_GURU);
  var row = parseInt(rowIndex);
  if (sheetGuru && row > 1) {
    var cellCheck = sheetGuru.getRange(row, 1).getValue();
    if (!cellCheck || cellCheck.toString().trim() === "") return "⚠️ Baris tidak ditemukan. Silakan refresh.";
    sheetGuru.deleteRow(row);
    invalidateGuruCache();
    return "✅ Data Guru berhasil dihapus!";
  }
  return "❌ Gagal menghapus data guru.";
}

// ==========================================

// ==========================================
// UPDATE GURU LENGKAP (replace semua mapel guru sekaligus)
// payload = { idGuru, namaBaru, mapelGroups: [{mapel, kelasArr}] }
// ==========================================
function updateGuruLengkap(payload) {
  if (!payload || !payload.idGuru) return '\u274c ID guru kosong.';
  if (!payload.mapelGroups || payload.mapelGroups.length === 0)
    return '\u274c Minimal 1 mata pelajaran.';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_GURU);
  if (!sheet) return '\u274c Sheet guru tidak ditemukan.';
  // Paksa kolom Kelas (D) jadi teks
  sheet.getRange(1, 4, sheet.getMaxRows(), 1).setNumberFormat('@');

  var idGuru = payload.idGuru.toString().trim();
  var data = sheet.getDataRange().getValues();

  // Ambil nama (pakai namaBaru, atau dari row existing)
  var nama = payload.namaBaru ? payload.namaBaru.toString().trim() : '';
  if (!nama) {
    for (var i = 1; i < data.length; i++) {
      if ((data[i][0]||'').toString().trim() === idGuru) { nama = data[i][1] ? data[i][1].toString() : ''; break; }
    }
  }

  // Hapus semua baris guru ini (dari bawah ke atas)
  for (var i = data.length - 1; i >= 1; i--) {
    if ((data[i][0]||'').toString().trim() === idGuru) sheet.deleteRow(i + 1);
  }

  // Insert ulang semua mapel group
  payload.mapelGroups.forEach(function(g) {
    sheet.appendRow([idGuru, nama, g.mapel || '', (g.kelasArr || []).join(', ')]);
  });

  invalidateGuruCache();
  return '\u2705 Data guru ' + nama + ' berhasil diperbarui.';
}

// ==========================================
// HAPUS GURU SECARA KESELURUHAN (semua mapel-nya)
// ==========================================
function hapusGuruByID(idGuru) {
  if (!idGuru) return '\u274c ID guru kosong.';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_GURU);
  if (!sheet) return '\u274c Sheet guru tidak ditemukan.';
  var data = sheet.getDataRange().getValues();
  var id = idGuru.toString().trim();
  var count = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if ((data[i][0]||'').toString().trim() === id) { sheet.deleteRow(i + 1); count++; }
  }
  invalidateGuruCache();
  return '\u2705 Guru dihapus (' + count + ' mata pelajaran).';
}