// ==========================================
// MASTERDATASERVICE.GS
// Data Mata Pelajaran: Mapel | Tingkat | Honor Offline | Honor Online
// Data Ruangan: Ruangan | Kapasitas
// Data Kelas: Jenjang | Tingkat | Rombel
// ==========================================

var CACHE_KEY_MASTER = 'azara_masterData';

function invalidateMasterCache() {
  CacheService.getScriptCache().remove(CACHE_KEY_MASTER);
}

// ── Auto-seed Data Kelas kalau belum ada (pakai default 12 kelas lama) ──
function ensureDataKelasSheet_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_KELAS);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_KELAS);
    sheet.appendRow(['Jenjang','Tingkat','Rombel']);
    var defaults = [
      ['SD','1','SD 1'],['SD','2','SD 2'],['SD','3','SD 3'],
      ['SD','4','SD 4'],['SD','5','SD 5'],['SD','6','SD 6'],
      ['SMP','7','SMP 7'],['SMP','8','SMP 8'],['SMP','9','SMP 9'],
      ['SMA','10','SMA 10'],['SMA','11','SMA 11'],['SMA','12','SMA 12']
    ];
    defaults.forEach(function(r){ sheet.appendRow(r); });
  }
  return sheet;
}

// ==========================================
// GET MASTER DATA (mapel + ruangan + kelas)
// ==========================================
function getMasterData() {
  var cache  = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY_MASTER);
  if (cached) return JSON.parse(cached);

  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var mapelList   = [];
  var ruanganList = [];
  var kelasList   = [];

  // Data Mata Pelajaran (kolom B sekarang = TINGKAT, bukan rombel)
  try {
    var sheetMapel = ss.getSheetByName(CONFIG.SHEET_MAPEL);
    if (sheetMapel) {
      var dm = sheetMapel.getDataRange().getValues();
      for (var i = 1; i < dm.length; i++) {
        if (!dm[i][0] || dm[i][0].toString().trim() === '') continue;
        mapelList.push({
          rowIndex     : i + 1,
          mapel        : dm[i][0].toString().trim(),
          kelas        : dm[i][1] ? dm[i][1].toString().trim() : '',  // = tingkat
          honorOffline : (dm[i][2] && !isNaN(dm[i][2])) ? Number(dm[i][2]) : 0,
          honorOnline  : (dm[i][3] && !isNaN(dm[i][3])) ? Number(dm[i][3]) : 0
        });
      }
    }
  } catch(e) { Logger.log('Error mapel: ' + e); }

  // Data Ruangan
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

  // Data Kelas (Jenjang | Tingkat | Rombel)
  try {
    var sheetKelas = ensureDataKelasSheet_(ss);
    var dk = sheetKelas.getDataRange().getValues();
    for (var k = 1; k < dk.length; k++) {
      if (!dk[k][2] || dk[k][2].toString().trim() === '') continue;  // rombel wajib
      kelasList.push({
        rowIndex: k + 1,
        jenjang : dk[k][0] ? dk[k][0].toString().trim() : '',
        tingkat : dk[k][1] ? dk[k][1].toString().trim() : '',
        rombel  : dk[k][2].toString().trim()
      });
    }
  } catch(e) { Logger.log('Error kelas: ' + e); }

  var result = { mapelList: mapelList, ruanganList: ruanganList, kelasList: kelasList };
  try { cache.put(CACHE_KEY_MASTER, JSON.stringify(result), 120); } catch(e) {}
  return result;
}

// ==========================================
// SIMPAN / EDIT MATA PELAJARAN (kolom B = TINGKAT)
// ==========================================
function simpanMapel(payload) {
  if (!payload.mapel || payload.mapel.toString().trim() === '')
    return '❌ Nama mata pelajaran tidak boleh kosong!';
  if (!payload.kelas || payload.kelas.toString().trim() === '')
    return '❌ Tingkat tidak boleh kosong!';

  var honorOffline = Number(payload.honorOffline) || 0;
  var honorOnline  = Number(payload.honorOnline)  || 0;

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_MAPEL);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_MAPEL);
    sheet.appendRow(['Mata Pelajaran','Tingkat','Honor Offline per Pertemuan','Honor Online per Pertemuan']);
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

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1] &&
        data[i][0].toString().trim().toLowerCase() === payload.mapel.toLowerCase() &&
        data[i][1].toString().trim().toLowerCase() === payload.kelas.toLowerCase()) {
      return '⚠️ ' + payload.mapel + ' untuk tingkat ' + payload.kelas + ' sudah ada!';
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

// ==========================================
// SIMPAN / EDIT / HAPUS KELAS (Rombel)
// Data Kelas: Jenjang | Tingkat | Rombel
// ==========================================
function simpanKelas(payload) {
  if (!payload.rombel || payload.rombel.toString().trim() === '')
    return '❌ Nama rombel/kelas tidak boleh kosong!';
  if (!payload.tingkat || payload.tingkat.toString().trim() === '')
    return '❌ Tingkat tidak boleh kosong!';
  if (!payload.jenjang || payload.jenjang.toString().trim() === '')
    return '❌ Jenjang tidak boleh kosong!';

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ensureDataKelasSheet_(ss);

  if (payload.rowIndex && parseInt(payload.rowIndex) > 1) {
    var row = parseInt(payload.rowIndex);
    sheet.getRange(row, 1).setValue(payload.jenjang);
    sheet.getRange(row, 2).setValue(payload.tingkat);
    sheet.getRange(row, 3).setValue(payload.rombel);
    invalidateMasterCache();
    return '✅ Kelas berhasil diperbarui!';
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][2] && data[i][2].toString().trim().toLowerCase() === payload.rombel.toLowerCase()) {
      return '⚠️ Rombel ' + payload.rombel + ' sudah ada!';
    }
  }
  sheet.appendRow([payload.jenjang, payload.tingkat, payload.rombel]);
  invalidateMasterCache();
  return '✅ Kelas/Rombel berhasil ditambahkan!';
}

function hapusKelas(rowIndex) {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_KELAS);
  var row   = parseInt(rowIndex);
  if (sheet && row > 1) {
    sheet.deleteRow(row);
    invalidateMasterCache();
    return '✅ Kelas berhasil dihapus!';
  }
  return '❌ Gagal menghapus.';
}