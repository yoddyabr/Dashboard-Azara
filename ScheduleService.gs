// ==========================================
// SCHEDULESERVICE.GS - CRUD Schedule + Rekap Honor
// Kolom Schedule: Tgl | JamMulai | JamBerakhir | Durasi | Ruangan | Kelas | Mapel | IDGuru | NamaGuru | Tipe Sesi
// ==========================================

var CACHE_KEY_SCHEDULE = 'azara_scheduleData';

function invalidateScheduleCache() {
  CacheService.getScriptCache().remove(CACHE_KEY_SCHEDULE);
}

// ==========================================
// 1. GET SCHEDULE DATA (with cache)
// ==========================================
function getScheduleData() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY_SCHEDULE);
  if (cached) return JSON.parse(cached);

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scheduleList = [];

  try {
    var sheetSchedule = ss.getSheetByName(CONFIG.SHEET_SCHEDULE);
    if (sheetSchedule) {
      var data = sheetSchedule.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        if (!data[i][0] || data[i][0].toString().trim() === "") continue;
        var rawTgl = data[i][0];
        var tglString = "";
        if (rawTgl instanceof Date) {
          // Simpan dalam format lokal YYYY-MM-DD untuk hindari timezone offset
          var y = rawTgl.getFullYear();
          var m = String(rawTgl.getMonth() + 1).padStart(2,"0");
          var d = String(rawTgl.getDate()).padStart(2,"0");
          tglString = y + "-" + m + "-" + d + "T00:00:00";
        } else if (rawTgl) {
          tglString = rawTgl.toString();
        }
        scheduleList.push({
          rowIndex   : i + 1,
          tgl        : tglString,
          jamMulai   : data[i][1] ? data[i][1].toString().trim() : "",
          jamBerakhir: data[i][2] ? data[i][2].toString().trim() : "",
          durasi     : data[i][3] ? data[i][3].toString().trim() : "",
          ruangan    : data[i][4] ? data[i][4].toString().trim() : "",
          kelas      : data[i][5] ? data[i][5].toString().trim() : "",
          mapel      : data[i][6] ? data[i][6].toString().trim() : "",
          idGuru     : data[i][7] ? data[i][7].toString().trim() : "",
          namaGuru   : data[i][8] ? data[i][8].toString().trim() : "",
          tipeSesi   : data[i][9] ? data[i][9].toString().trim() : "Offline"
        });
      }
    }
  } catch(e) { Logger.log("Error sheet schedule: " + e.toString()); }

  try { cache.put(CACHE_KEY_SCHEDULE, JSON.stringify(scheduleList), 120); } catch(e) {}
  return scheduleList;
}

// ==========================================
// 2. SIMPAN / EDIT JADWAL
// ==========================================
function simpanSchedule(payload) {
  if (!payload.tgl)   return "❌ Gagal: Tanggal tidak boleh kosong!";
  if (!payload.kelas) return "❌ Gagal: Kelas tidak boleh kosong!";
  if (!payload.mapel) return "❌ Gagal: Mata pelajaran tidak boleh kosong!";
  if (!payload.idGuru)return "❌ Gagal: Guru tidak boleh kosong!";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSchedule = ss.getSheetByName(CONFIG.SHEET_SCHEDULE);
  if (!sheetSchedule) {
    sheetSchedule = ss.insertSheet(CONFIG.SHEET_SCHEDULE);
    sheetSchedule.appendRow(["Tanggal","Jam Mulai","Jam Berakhir","Durasi","Ruangan","Kelas","Mata Pelajaran","ID Guru","Nama Guru","Tipe Sesi"]);
  }

  // Parse tanggal dengan aman (tgl format YYYY-MM-DD)
  var tglObj;
  if (payload.tgl.indexOf("T") < 0) {
    // YYYY-MM-DD murni -> tambahkan jam 12 supaya tidak shift timezone
    tglObj = new Date(payload.tgl + "T12:00:00");
  } else {
    tglObj = new Date(payload.tgl);
  }

  var tipeSesi = payload.tipeSesi || "Offline";

  if (payload.rowIndex && parseInt(payload.rowIndex) > 1) {
    var row = parseInt(payload.rowIndex);
    sheetSchedule.getRange(row, 1).setValue(tglObj);
    sheetSchedule.getRange(row, 2).setValue(payload.jamMulai    || "");
    sheetSchedule.getRange(row, 3).setValue(payload.jamBerakhir || "");
    sheetSchedule.getRange(row, 4).setValue(payload.durasi      || "");
    sheetSchedule.getRange(row, 5).setValue(payload.ruangan     || "");
    sheetSchedule.getRange(row, 6).setValue(payload.kelas);
    sheetSchedule.getRange(row, 7).setValue(payload.mapel);
    sheetSchedule.getRange(row, 8).setValue(payload.idGuru);
    sheetSchedule.getRange(row, 9).setValue(payload.namaGuru    || "");
    sheetSchedule.getRange(row,10).setValue(tipeSesi);
    invalidateScheduleCache();
    return "✅ Jadwal berhasil diperbarui!";
  } else {
    sheetSchedule.appendRow([
      tglObj,
      payload.jamMulai    || "",
      payload.jamBerakhir || "",
      payload.durasi      || "",
      payload.ruangan     || "",
      payload.kelas,
      payload.mapel,
      payload.idGuru,
      payload.namaGuru    || "",
      tipeSesi
    ]);
    invalidateScheduleCache();
    return "✅ Jadwal berhasil ditambahkan!";
  }
}

// ==========================================
// 3. HAPUS JADWAL
// ==========================================
function hapusSchedule(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSchedule = ss.getSheetByName(CONFIG.SHEET_SCHEDULE);
  var row = parseInt(rowIndex);
  if (sheetSchedule && row > 1) {
    var cellCheck = sheetSchedule.getRange(row, 1).getValue();
    if (!cellCheck) return "⚠️ Baris tidak ditemukan. Silakan refresh halaman.";
    sheetSchedule.deleteRow(row);
    invalidateScheduleCache();
    return "✅ Jadwal berhasil dihapus!";
  }
  return "❌ Gagal menghapus jadwal.";
}

// ==========================================
// 4. REKAP HONOR BULANAN per GURU
// Honor diambil dari Data Mata Pelajaran berdasarkan mapel+kelas+tipeSesi
// ==========================================
function getRekapHonorBulanan(bulan, tahun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Build honor map: key = "mapel||kelas" → {offline, online}
  var honorMap = {};
  var sheetMapel = ss.getSheetByName(CONFIG.SHEET_MAPEL);
  if (sheetMapel) {
    var dm = sheetMapel.getDataRange().getValues();
    for (var h = 1; h < dm.length; h++) {
      if (!dm[h][0]) continue;
      var key = dm[h][0].toString().trim().toLowerCase()
              + "||"
              + (dm[h][1] ? dm[h][1].toString().trim().toLowerCase() : "");
      honorMap[key] = {
        offline: Number(dm[h][2]) || 0,
        online : Number(dm[h][3]) || 0
      };
    }
  }

  var sheetSchedule = ss.getSheetByName(CONFIG.SHEET_SCHEDULE);
  if (!sheetSchedule) return [];

  var dataSchedule = sheetSchedule.getDataRange().getValues();
  var rekapMap = {};

  for (var i = 1; i < dataSchedule.length; i++) {
    var row = dataSchedule[i];
    if (!row[0]) continue;

    var tgl = row[0] instanceof Date ? row[0] : new Date(row[0]);
    if (isNaN(tgl.getTime())) continue;
    if ((tgl.getMonth() + 1) !== parseInt(bulan) || tgl.getFullYear() !== parseInt(tahun)) continue;

    var idGuru   = row[7] ? row[7].toString().trim() : "";
    var namaGuru = row[8] ? row[8].toString().trim() : "";
    var mapel    = row[6] ? row[6].toString().trim() : "";
    var kelas    = row[5] ? row[5].toString().trim() : "";
    var tipeSesi = row[9] ? row[9].toString().trim().toLowerCase() : "offline";
    if (!idGuru) continue;

    var honorKey = mapel.toLowerCase() + "||" + kelas.toLowerCase();
    var honorEntry = honorMap[honorKey] || { offline: 0, online: 0 };
    var honorPerSesi = (tipeSesi === "online") ? honorEntry.online : honorEntry.offline;

    if (!rekapMap[idGuru]) {
      rekapMap[idGuru] = { id: idGuru, nama: namaGuru, totalSesi: 0, totalHonor: 0, detail: [] };
    }
    rekapMap[idGuru].totalSesi++;
    rekapMap[idGuru].totalHonor += honorPerSesi;
    rekapMap[idGuru].detail.push({
      tgl         : tgl.toISOString(),
      mapel       : mapel,
      kelas       : kelas,
      tipeSesi    : tipeSesi,
      honorPerSesi: honorPerSesi
    });
  }

  var result = Object.values(rekapMap);
  result.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return result;
}