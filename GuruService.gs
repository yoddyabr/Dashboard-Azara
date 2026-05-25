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
  if (!payload.mapel   || payload.mapel.toString().trim()   === "") return "❌ Mata pelajaran tidak boleh kosong!";
  if (!payload.kelasArr || payload.kelasArr.length === 0)           return "❌ Pilih minimal satu kelas!";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetGuru = ss.getSheetByName(CONFIG.SHEET_GURU);
  if (!sheetGuru) {
    sheetGuru = ss.insertSheet(CONFIG.SHEET_GURU);
    sheetGuru.appendRow(["ID Guru","Nama Guru","Mata Pelajaran","Kelas"]);
  }

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