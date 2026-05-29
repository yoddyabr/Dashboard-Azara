// ==========================================
// SCHEDULESERVICE.GS - CRUD Schedule + Rekap Honor
// Kolom Schedule: Tgl | JamMulai | JamBerakhir | Durasi | Ruangan | Kelas | Mapel | IDGuru | NamaGuru | Tipe Sesi
// ==========================================

var CACHE_KEY_SCHEDULE = 'azara_scheduleData';

function invalidateScheduleCache() {
  CacheService.getScriptCache().remove(CACHE_KEY_SCHEDULE);
}

// ── Helper: normalize jam (string "HH:MM" atau Date object) ──
function formatJamString(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  }
  var s = val.toString().trim();
  // Sudah HH:MM? langsung
  if (/^\d{1,2}:\d{2}$/.test(s)) {
    var p = s.split(':');
    return p[0].padStart(2, '0') + ':' + p[1];
  }
  // Coba parse sebagai date string
  var d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "HH:mm");
  }
  return s;
}

// ── Helper: jam string "HH:MM" → menit total (untuk komparasi) ──
function jamToMinutes(jamStr) {
  if (!jamStr) return null;
  var s = formatJamString(jamStr);
  var m = s.match(/(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return parseInt(m[1])*60 + parseInt(m[2]);
}

// ── Helper: konflik di sheet (untuk dipakai sebelum simpan) ──
function findKonflikInSheet(sheetSchedule, ruangan, tglObj, jamMulai, jamBerakhir, excludeRow) {
  if (!ruangan || !tglObj) return [];
  var data = sheetSchedule.getDataRange().getValues();
  var targetDateStr = Utilities.formatDate(tglObj, Session.getScriptTimeZone(), "yyyy-MM-dd");
  var newS = jamToMinutes(jamMulai);
  var newE = jamToMinutes(jamBerakhir);
  if (newS == null || newE == null) return [];

  var konflik = [];
  for (var i = 1; i < data.length; i++) {
    var rowNum = i + 1;
    if (excludeRow && rowNum === excludeRow) continue;
    var row = data[i];
    if (!row[0]) continue;

    // Date check (col A)
    var rowDate = row[0] instanceof Date ? row[0] : new Date(row[0]);
    if (isNaN(rowDate.getTime())) continue;
    var rowDateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (rowDateStr !== targetDateStr) continue;

    // Ruangan check (col E)
    var rowRuangan = row[4] ? row[4].toString().trim().toLowerCase() : "";
    if (rowRuangan !== ruangan.toString().trim().toLowerCase()) continue;

    // Time overlap check
    var rowMulai = jamToMinutes(row[1]);
    var rowAkhir = jamToMinutes(row[2]);
    if (rowMulai == null || rowAkhir == null) continue;
    if (rowMulai < newE && newS < rowAkhir) {
      konflik.push({
        rowIndex   : rowNum,
        mapel      : row[6] ? row[6].toString().trim() : "",
        kelas      : row[5] ? row[5].toString().trim() : "",
        namaGuru   : row[8] ? row[8].toString().trim() : "",
        jamMulai   : formatJamString(row[1]),
        jamBerakhir: formatJamString(row[2])
      });
    }
  }
  return konflik;
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
          jamMulai   : formatJamString(data[i][1]),   // NORMALIZE ke HH:MM
          jamBerakhir: formatJamString(data[i][2]),   // NORMALIZE ke HH:MM
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
// 2. SIMPAN / EDIT JADWAL (dengan KONFLIK CHECK di backend)
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

  // Parse tanggal dengan aman
  var tglObj;
  if (payload.tgl.indexOf("T") < 0) {
    tglObj = new Date(payload.tgl + "T12:00:00");
  } else {
    tglObj = new Date(payload.tgl);
  }

  var excludeRow = payload.rowIndex ? parseInt(payload.rowIndex) : null;

  // ── CEK KONFLIK SEBELUM SIMPAN (kecuali user sudah explicit confirm via forceOverwrite) ──
  if (payload.ruangan && payload.jamMulai && payload.jamBerakhir && !payload.forceOverwrite) {
    var konflik = findKonflikInSheet(sheetSchedule, payload.ruangan, tglObj, payload.jamMulai, payload.jamBerakhir, excludeRow);
    if (konflik.length > 0) {
      // Format pesan khusus yang akan di-detect oleh frontend
      var info = konflik.map(function(k){
        return k.mapel + " " + k.kelas + " (" + k.jamMulai + "-" + k.jamBerakhir + ") oleh " + k.namaGuru;
      }).join("; ");
      // Return string khusus dengan kode KONFLIK
      return "⚠️KONFLIK⚠️ Ruangan " + payload.ruangan + " sudah dipakai: " + info;
    }
  }

  // ── Kalau forceOverwrite, hapus konflik dulu lalu lanjut simpan ──
  if (payload.forceOverwrite && payload.ruangan && payload.jamMulai && payload.jamBerakhir) {
    var konflikForce = findKonflikInSheet(sheetSchedule, payload.ruangan, tglObj, payload.jamMulai, payload.jamBerakhir, excludeRow);
    // Hapus dari row tertinggi ke terendah (supaya rowIndex tidak shift)
    konflikForce.sort(function(a,b){ return b.rowIndex - a.rowIndex; });
    konflikForce.forEach(function(k) {
      try { sheetSchedule.deleteRow(k.rowIndex); } catch(e) {}
    });
  }

  var tipeSesi = payload.tipeSesi || "Offline";

  if (excludeRow && excludeRow > 1) {
    var row = excludeRow;
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

  // Build map rombel → tingkat dari Data Kelas
  var rombelToTingkat = {};
  var sheetKelas = ss.getSheetByName(CONFIG.SHEET_KELAS);
  if (sheetKelas) {
    var dkls = sheetKelas.getDataRange().getValues();
    for (var rk = 1; rk < dkls.length; rk++) {
      if (!dkls[rk][2]) continue;
      rombelToTingkat[dkls[rk][2].toString().trim().toLowerCase()] =
        dkls[rk][1] ? dkls[rk][1].toString().trim().toLowerCase() : '';
    }
  }

  // Build honor map: key = "mapel||tingkat" → {offline, online}
  var honorMap = {};
  var sheetMapel = ss.getSheetByName(CONFIG.SHEET_MAPEL);
  if (sheetMapel) {
    var dm = sheetMapel.getDataRange().getValues();
    for (var h = 1; h < dm.length; h++) {
      if (!dm[h][0]) continue;
      var key = dm[h][0].toString().trim().toLowerCase()
              + "||"
              + (dm[h][1] ? dm[h][1].toString().trim().toLowerCase() : "");  // tingkat
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
    var rombel   = row[5] ? row[5].toString().trim() : "";
    var tipeSesi = row[9] ? row[9].toString().trim().toLowerCase() : "offline";
    if (!idGuru) continue;

    // Rombel → tingkat (fallback ke rombel itu sendiri kalau tidak ada di map)
    var tingkat = rombelToTingkat[rombel.toLowerCase()] || rombel.toLowerCase();
    var honorKey = mapel.toLowerCase() + "||" + tingkat;
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
      kelas       : rombel,
      tipeSesi    : tipeSesi,
      honorPerSesi: honorPerSesi
    });
  }

  var result = Object.values(rekapMap);
  result.sort(function(a, b) { return a.nama.localeCompare(b.nama); });
  return result;
}