// ==========================================
// MASTERDATASERVICE.GS
// Data Mata Pelajaran: Mapel | Kelas | Honor Offline | Honor Online
// Data Ruangan: Ruangan | Kapasitas
// ==========================================

var CACHE_KEY_MASTER = 'azara_masterData';

function invalidateMasterCache() {
  CacheService.getScriptCache().remove(CACHE_KEY_MASTER);
}

// ==========================================
// GET MASTER DATA (mapel 4 kolom + ruangan)
// ==========================================
function getMasterData() {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY_MASTER);
  if (cached) return JSON.parse(cached);

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var mapelList   = [];
  var ruanganList = [];

  try {
    var sheetMapel = ss.getSheetByName(CONFIG.SHEET_MAPEL);
    if (sheetMapel) {
      var dm = sheetMapel.getDataRange().getValues();
      for (var i = 1; i < dm.length; i++) {
        if (!dm[i][0] || dm[i][0].toString().trim() === '') continue;
        mapelList.push({
          rowIndex     : i + 1,
          mapel        : dm[i][0].toString().trim(),
          kelas        : dm[i][1] ? dm[i][1].toString().trim() : '',
          honorOffline : (dm[i][2] && !isNaN(dm[i][2])) ? Number(dm[i][2]) : 0,
          honorOnline  : (dm[i][3] && !isNaN(dm[i][3])) ? Number(dm[i][3]) : 0
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
// SIMPAN / EDIT MATA PELAJARAN (4 kolom)
// ==========================================
function simpanMapel(payload) {
  if (!payload.mapel || payload.mapel.toString().trim() === '')
    return '❌ Nama mata pelajaran tidak boleh kosong!';
  if (!payload.kelas || payload.kelas.toString().trim() === '')
    return '❌ Kelas tidak boleh kosong!';

  var honorOffline = Number(payload.honorOffline) || 0;
  var honorOnline  = Number(payload.honorOnline)  || 0;

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MAPEL);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_MAPEL);
    sheet.appendRow(['Mata Pelajaran','Kelas','Honor Offline per Pertemuan','Honor Online per Pertemuan']);
  }

  if (payload.rowIndex && parseInt(payload.rowIndex) > 1) {
    var row = parseInt(payload.rowIndex);
    sheet.getRange(row, 1).setValue(payload.mapel);
    sheet.getRange(row, 2).setValue(payload.kelas);
    sheet.getRange(row, 3).setValue(honorOffline);
    sheet.getRange(row, 4).setValue(honorOnline);
    invalidateMasterCache();
    return '✅ Data berhasil diperbarui!';
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
  sheet.appendRow([payload.mapel, payload.kelas, honorOffline, honorOnline]);
  invalidateMasterCache();
  return '✅ Mata Pelajaran berhasil ditambahkan!';
}

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
// SIMPAN / EDIT RUANGAN
// ==========================================
function simpanRuanganMaster(payload) {
  if (!payload.ruangan || payload.ruangan.toString().trim() === '')
    return '❌ Nama ruangan tidak boleh kosong!';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_RUANGAN);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_RUANGAN);
    sheet.appendRow(['Ruangan','Kapasitas']);
  }

  if (payload.rowIndex && parseInt(payload.rowIndex) > 1) {
    var row = parseInt(payload.rowIndex);
    sheet.getRange(row, 1).setValue(payload.ruangan);
    sheet.getRange(row, 2).setValue(Number(payload.kapasitas) || 0);
    invalidateMasterCache();
    return '✅ Ruangan berhasil diperbarui!';
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