// ==========================================
// MASTERDATASERVICE.GS
// CRUD: Data Mata Pelajaran + Data Ruangan
// ==========================================

var CACHE_KEY_MASTER = 'azara_masterData';

function invalidateMasterCache() {
  CacheService.getScriptCache().remove(CACHE_KEY_MASTER);
}

// ==========================================
// 1. GET MASTER DATA (with cache)
// ==========================================
function getMasterData() {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY_MASTER);
  if (cached) return JSON.parse(cached);

  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var mapelList = [], ruanganList = [];

  try {
    var sheetMapel = ss.getSheetByName(CONFIG.SHEET_MAPEL);
    if (sheetMapel) {
      var dm = sheetMapel.getDataRange().getValues();
      for (var i = 1; i < dm.length; i++) {
        if (!dm[i][0] || dm[i][0].toString().trim() === '') continue;
        mapelList.push({
          rowIndex: i + 1,
          mapel   : dm[i][0].toString().trim(),
          kelas   : dm[i][1] ? dm[i][1].toString().trim() : ''
        });
      }
    }
  } catch(e) { Logger.log('Error mapel: ' + e); }

  try {
    var sheetRuangan = ss.getSheetByName(CONFIG.SHEET_RUANGAN);
    if (sheetRuangan) {
      var dr = sheetRuangan.getDataRange().getValues();
      for (var j = 1; j < dr.length; j++) {
        if (!dr[j][0] || dr[j][0].toString().trim() === '') continue;
        ruanganList.push({
          rowIndex : j + 1,
          ruangan  : dr[j][0].toString().trim(),
          kapasitas: (dr[j][1] && !isNaN(dr[j][1])) ? Number(dr[j][1]) : 0
        });
      }
    }
  } catch(e) { Logger.log('Error ruangan: ' + e); }

  var result = { mapelList: mapelList, ruanganList: ruanganList };
  try { cache.put(CACHE_KEY_MASTER, JSON.stringify(result), 120); } catch(e) {}
  return result;
}

// ==========================================
// 2. SIMPAN / EDIT MATA PELAJARAN
// ==========================================
function simpanMapel(payload) {
  if (!payload.mapel || payload.mapel.toString().trim() === '')
    return '❌ Nama mata pelajaran tidak boleh kosong!';
  if (!payload.kelas || payload.kelas.toString().trim() === '')
    return '❌ Kelas tidak boleh kosong!';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MAPEL);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_MAPEL);
    sheet.appendRow(['Mata Pelajaran', 'Kelas']);
  }

  if (payload.rowIndex && parseInt(payload.rowIndex) > 1) {
    var row = parseInt(payload.rowIndex);
    sheet.getRange(row, 1).setValue(payload.mapel);
    sheet.getRange(row, 2).setValue(payload.kelas);
    invalidateMasterCache();
    return '✅ Data Mata Pelajaran berhasil diperbarui!';
  }

  // Cek duplikat mapel + kelas
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1] &&
        data[i][0].toString().trim().toLowerCase() === payload.mapel.toLowerCase() &&
        data[i][1].toString().trim().toLowerCase() === payload.kelas.toLowerCase()) {
      return '⚠️ ' + payload.mapel + ' untuk kelas ' + payload.kelas + ' sudah ada!';
    }
  }
  sheet.appendRow([payload.mapel, payload.kelas]);
  invalidateMasterCache();
  return '✅ Mata Pelajaran berhasil ditambahkan!';
}

// ==========================================
// 3. HAPUS MATA PELAJARAN
// ==========================================
function hapusMapel(rowIndex) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MAPEL);
  var row   = parseInt(rowIndex);
  if (sheet && row > 1) {
    sheet.deleteRow(row);
    invalidateMasterCache();
    return '✅ Mata Pelajaran berhasil dihapus!';
  }
  return '❌ Gagal menghapus.';
}

// ==========================================
// 4. SIMPAN / EDIT RUANGAN
// ==========================================
function simpanRuanganMaster(payload) {
  if (!payload.ruangan || payload.ruangan.toString().trim() === '')
    return '❌ Nama ruangan tidak boleh kosong!';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_RUANGAN);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_RUANGAN);
    sheet.appendRow(['Ruangan', 'Kapasitas']);
  }

  if (payload.rowIndex && parseInt(payload.rowIndex) > 1) {
    var row = parseInt(payload.rowIndex);
    sheet.getRange(row, 1).setValue(payload.ruangan);
    sheet.getRange(row, 2).setValue(Number(payload.kapasitas) || 0);
    invalidateMasterCache();
    return '✅ Data Ruangan berhasil diperbarui!';
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === payload.ruangan.toLowerCase()) {
      return '⚠️ Ruangan ' + payload.ruangan + ' sudah ada!';
    }
  }
  sheet.appendRow([payload.ruangan, Number(payload.kapasitas) || 0]);
  invalidateMasterCache();
  return '✅ Ruangan berhasil ditambahkan!';
}

// ==========================================
// 5. HAPUS RUANGAN
// ==========================================
function hapusRuanganMaster(rowIndex) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_RUANGAN);
  var row   = parseInt(rowIndex);
  if (sheet && row > 1) {
    sheet.deleteRow(row);
    invalidateMasterCache();
    return '✅ Ruangan berhasil dihapus!';
  }
  return '❌ Gagal menghapus.';
}