// ==========================================
// DATASERVICE.GS - Ambil semua data (READ ONLY)
// ✅ OPT: CacheService — hindari baca Sheets setiap request
// ==========================================

var CACHE_KEY = 'azara_sistemData';
var CACHE_TTL = 120; // detik (2 menit)

// Hapus cache — dipanggil setiap kali ada operasi tulis
function invalidateCache() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

function getSistemData() {
  // 1. Coba ambil dari cache dulu
  var cache = CacheService.getScriptCache();
  var cached = cache.get(CACHE_KEY);
  if (cached) {
    return JSON.parse(cached);
  }

  // 2. Cache miss — baca dari Sheets
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var siswaList = [];
  var paymentList = [];

  try {
    var sheetSiswa = ss.getSheetByName(CONFIG.SHEET_SISWA) || ss.getSheets()[0];
    var dataSiswa = sheetSiswa.getDataRange().getValues();
    for (var i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][0] && dataSiswa[i][0].toString().trim() !== "") {
        var rawStatus = dataSiswa[i][3] ? dataSiswa[i][3].toString().trim().toLowerCase() : "tidak aktif";
        var statusFinal = "Tidak Aktif";
        if (rawStatus === "active" || rawStatus === "aktif") statusFinal = "Aktif";
        else if (rawStatus === "cuti") statusFinal = "Cuti";
        else if (rawStatus === "lulus") statusFinal = "Lulus";
        siswaList.push({
          id: dataSiswa[i][0].toString().trim(),
          nama: dataSiswa[i][1] ? dataSiswa[i][1].toString().trim() : "",
          kelas: dataSiswa[i][2] ? dataSiswa[i][2].toString().trim() : "Tanpa Kelas",
          status: statusFinal,
          tipe: dataSiswa[i][4] ? dataSiswa[i][4].toString().trim() : "Bulanan",
          biaya: (dataSiswa[i][5] && !isNaN(dataSiswa[i][5])) ? Number(dataSiswa[i][5]) : 0
        });
      }
    }
  } catch(e) { Logger.log("Error sheet siswa: " + e.toString()); }

  try {
    var sheetPay = ss.getSheetByName(CONFIG.SHEET_PAYMENT);
    if (sheetPay) {
      var dataPay = sheetPay.getDataRange().getValues();
      for (var j = 1; j < dataPay.length; j++) {
        if (dataPay[j][0] && dataPay[j][0].toString().trim() !== "") {
          var rawTgl = dataPay[j][4];
          var tglString = "";
          if (rawTgl instanceof Date) tglString = rawTgl.toISOString();
          else if (rawTgl && rawTgl.toString().trim() !== "") {
            tglString = new Date(rawTgl).toString() !== "Invalid Date" ? new Date(rawTgl).toISOString() : "";
          }
          paymentList.push({
            rowIndex: j + 1,
            id: dataPay[j][0].toString().trim(),
            nama: dataPay[j][1] ? dataPay[j][1].toString().trim() : "",
            nominal: (dataPay[j][2] && !isNaN(dataPay[j][2])) ? Number(dataPay[j][2]) : 0,
            periode: dataPay[j][3] ? dataPay[j][3].toString().trim() : "",
            tgl: tglString,
            transferKe: dataPay[j][5] ? dataPay[j][5].toString().trim() : "-",
            pengirim: dataPay[j][6] ? dataPay[j][6].toString().trim() : "-"
          });
        }
      }
    }
  } catch(e) { Logger.log("Error sheet payment: " + e.toString()); }

  var result = { siswaList: siswaList, paymentList: paymentList };

  // 3. Simpan ke cache
  try {
    cache.put(CACHE_KEY, JSON.stringify(result), CACHE_TTL);
  } catch(e) {
    Logger.log("Cache put error (data terlalu besar?): " + e.toString());
  }

  return result;
}