// ==========================================
// KELASSERVICE.GS - Proses Naik Kelas Tahunan
// ==========================================

// Peta kenaikan kelas — SD6 -> SMP7, SMP9 -> SMA10 otomatis lintas jenjang
var PETA_NAIK_KELAS = {
  "SD 1":  "SD 2",
  "SD 2":  "SD 3",
  "SD 3":  "SD 4",
  "SD 4":  "SD 5",
  "SD 5":  "SD 6",
  "SD 6":  "SMP 7",
  "SMP 7": "SMP 8",
  "SMP 8": "SMP 9",
  "SMP 9": "SMA 10",
  "SMA 10":"SMA 11",
  "SMA 11":"SMA 12",
  "SMA 12":"LULUS"   // SMA 12 -> status Lulus, kelas tetap SMA 12
};

// ==========================================
// 1. PREVIEW — kembalikan daftar perubahan tanpa menyimpan
// ==========================================
function previewNaikKelas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSiswa = ss.getSheetByName(CONFIG.SHEET_SISWA) || ss.getSheets()[0];
  var dataSiswa = sheetSiswa.getDataRange().getValues();
  var preview = [];

  for (var i = 1; i < dataSiswa.length; i++) {
    var row = dataSiswa[i];
    if (!row[0] || row[0].toString().trim() === "") continue;

    var id     = row[0].toString().trim();
    var nama   = row[1] ? row[1].toString().trim() : "";
    var kelas  = row[2] ? row[2].toString().trim() : "";
    var status = row[3] ? row[3].toString().trim().toLowerCase() : "";

    // Proses semua siswa tanpa filter status

    var kelasNext = PETA_NAIK_KELAS[kelas];
    if (!kelasNext) continue; // kelas tidak dikenal dalam peta, skip

    preview.push({
      rowIndex : i + 1,
      id       : id,
      nama     : nama,
      kelasLama: kelas,
      kelasBaru: kelasNext === "LULUS" ? kelas : kelasNext, // kelas tetap jika lulus
      lulus    : kelasNext === "LULUS"
    });
  }

  return preview;
}

// ==========================================
// 2. EKSEKUSI — simpan perubahan ke Sheets
// ==========================================
function eksekusiNaikKelas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSiswa = ss.getSheetByName(CONFIG.SHEET_SISWA) || ss.getSheets()[0];
  var dataSiswa = sheetSiswa.getDataRange().getValues();
  var countNaik = 0;
  var countLulus = 0;

  for (var i = 1; i < dataSiswa.length; i++) {
    var row    = dataSiswa[i];
    if (!row[0] || row[0].toString().trim() === "") continue;

    var kelas  = row[2] ? row[2].toString().trim() : "";
    var status = row[3] ? row[3].toString().trim().toLowerCase() : "";

    // Proses semua siswa tanpa filter status

    var kelasNext = PETA_NAIK_KELAS[kelas];
    if (!kelasNext) continue;

    if (kelasNext === "LULUS") {
      // Kelas tetap SMA 12, hanya status yang berubah jadi Lulus
      sheetSiswa.getRange(i + 1, 4).setValue("Lulus");
      countLulus++;
    } else {
      // Naik kelas, status tidak berubah
      sheetSiswa.getRange(i + 1, 3).setValue(kelasNext);
      countNaik++;
    }
  }

  invalidateCache();
  return { countNaik: countNaik, countLulus: countLulus };
}