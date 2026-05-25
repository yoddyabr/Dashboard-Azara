// ==========================================
// SCHEDULESERVICE.GS - CRUD Data Schedule + Rekap Honor Bulanan
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
        if (rawTgl instanceof Date) tglString = rawTgl.toISOString();
        else if (rawTgl) {
          var parsed = new Date(rawTgl);
          tglString = !isNaN(parsed.getTime()) ? parsed.toISOString() : "";
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
          namaGuru   : data[i][8] ? data[i][8].toString().trim() : ""
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
    sheetSchedule.appendRow(["Tanggal","Jam Mulai","Jam Berakhir","Durasi","Ruangan","Kelas","Mata Pelajaran","ID Guru","Nama Guru"]);
  }

  var tglObj = new Date(payload.tgl);

  if (payload.rowIndex && parseInt(payload.rowIndex) > 1) {
    // MODE EDIT
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
      payload.namaGuru    || ""
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
// Hitung: jumlah sesi × honor per sesi dari Data Honor Guru
// ==========================================
function getRekapHonorBulanan(bulan, tahun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Ambil honor rate map: key = "mapel||kelas" → honor
  var honorMap = {};
  var sheetHonor = ss.getSheetByName(CONFIG.SHEET_HONOR);
  if (sheetHonor) {
    var dataHonor = sheetHonor.getDataRange().getValues();
    for (var h = 1; h < dataHonor.length; h++) {
      if (!dataHonor[h][0]) continue;
      var key = dataHonor[h][0].toString().trim().toLowerCase()
              + "||"
              + (dataHonor[h][1] ? dataHonor[h][1].toString().trim().toLowerCase() : "");
      honorMap[key] = Number(dataHonor[h][2]) || 0;
    }
  }

  // Baca schedule, filter bulan & tahun
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
    if (!idGuru) continue;

    var honorKey     = mapel.toLowerCase() + "||" + kelas.toLowerCase();
    var honorPerSesi = honorMap[honorKey] || 0;

    if (!rekapMap[idGuru]) {
      rekapMap[idGuru] = { id: idGuru, nama: namaGuru, totalSesi: 0, totalHonor: 0, detail: [] };
    }
    rekapMap[idGuru].totalSesi++;
    rekapMap[idGuru].totalHonor += honorPerSesi;
    rekapMap[idGuru].detail.push({
      tgl         : tgl.toISOString(),
      mapel       : mapel,
      kelas       : kelas,
      honorPerSesi: honorPerSesi
    });
  }

  var result = Object.values(rekapMap);
  result.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return result;
}