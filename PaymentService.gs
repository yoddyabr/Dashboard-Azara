// ==========================================
// PAYMENTSERVICE.GS - CRUD Data Pembayaran
// ✅ OPT: invalidateCache() setiap operasi tulis
// ==========================================

// 1. SIMPAN PEMBAYARAN BARU
function simpanPembayaran(payload) {
  if (!payload.idSiswa || payload.idSiswa.toString().trim() === "") return "❌ Gagal: ID Siswa tidak valid!";
  if (!payload.nominalTotal || isNaN(payload.nominalTotal) || Number(payload.nominalTotal) <= 0) return "❌ Gagal: Nominal pembayaran harus diisi dengan angka yang valid!";
  if (!payload.periodeList || payload.periodeList.length === 0) return "❌ Gagal: Minimal pilih satu periode bulan!";
  if (!payload.tglManual) return "❌ Gagal: Tanggal transaksi harus diisi!";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPay = ss.getSheetByName(CONFIG.SHEET_PAYMENT);
  if (!sheetPay) {
    sheetPay = ss.insertSheet("Historical Payment");
    sheetPay.appendRow(["ID Siswa", "Nama Siswa", "Nominal", "Periode Bulan", "Tanggal Transaksi", "Transfer Ke", "Pengirim"]);
  }

  var existingData = sheetPay.getDataRange().getValues();
  var idSiswaClean = payload.idSiswa.toString().trim().toLowerCase();
  var periodeYangSudahAda = [];

  for (var i = 1; i < existingData.length; i++) {
    var existingId = existingData[i][0] ? existingData[i][0].toString().trim().toLowerCase() : "";
    var existingPeriode = existingData[i][3] ? existingData[i][3].toString().trim().toLowerCase() : "";
    if (existingId === idSiswaClean) periodeYangSudahAda.push(existingPeriode);
  }

  var periodeDuplikat = [];
  payload.periodeList.forEach(function(periode) {
    if (periodeYangSudahAda.indexOf(periode.toString().trim().toLowerCase()) !== -1) periodeDuplikat.push(periode);
  });

  if (periodeDuplikat.length > 0) {
    return "⚠️ DUPLIKAT TERDETEKSI: Siswa " + payload.idSiswa + " sudah memiliki catatan pembayaran untuk periode: " + periodeDuplikat.join(", ") + ". Batalkan atau hapus data lama terlebih dahulu jika ingin menggantinya.";
  }

  var tglObj = new Date(payload.tglManual);
  var jumlahBulan = payload.periodeList.length;
  var nominalPerBulan = Math.round((Number(payload.nominalTotal) / jumlahBulan) * 100) / 100;

  payload.periodeList.forEach(function(periode) {
    var lastRow = sheetPay.getLastRow() + 1;
    sheetPay.appendRow([payload.idSiswa, payload.namaSiswa, nominalPerBulan, periode.toString().trim(), tglObj, payload.transferKe, payload.pengirim]);
    sheetPay.getRange(lastRow, 4).setNumberFormat("@");
  });

  invalidateCache();
  return "✅ Sukses mencatat " + jumlahBulan + " transaksi pembayaran!";
}

// 2. UPDATE PEMBAYARAN (EDIT MODE)
function updatePembayaranDariSheet(payload) {
  if (!payload.nominal || isNaN(payload.nominal) || Number(payload.nominal) <= 0) return "❌ Gagal: Nominal tidak valid!";
  if (!payload.periode || payload.periode.toString().trim() === "") return "❌ Gagal: Periode tidak boleh kosong!";
  if (!payload.tglManual) return "❌ Gagal: Tanggal transaksi harus diisi!";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPay = ss.getSheetByName(CONFIG.SHEET_PAYMENT);
  var row = parseInt(payload.rowIndex);

  if (sheetPay && row > 1) {
    var existingData = sheetPay.getDataRange().getValues();
    var targetRowData = existingData[row - 2];
    if (!targetRowData || !targetRowData[0]) {
      return "❌ Gagal: Baris data tidak ditemukan. Mungkin sudah terhapus. Silakan refresh halaman.";
    }
    sheetPay.getRange(row, 3).setValue(Number(payload.nominal));
    sheetPay.getRange(row, 4).setNumberFormat("@");
    sheetPay.getRange(row, 4).setValue(payload.periode.toString().trim());
    sheetPay.getRange(row, 5).setValue(new Date(payload.tglManual));
    sheetPay.getRange(row, 6).setValue(payload.transferKe);
    sheetPay.getRange(row, 7).setValue(payload.pengirim);
    invalidateCache();
    return "✅ Transaksi pada baris " + row + " berhasil diperbarui!";
  }
  return "❌ Gagal memperbarui transaksi.";
}

// 3. HAPUS SINGLE PEMBAYARAN
function hapusPembayaranDariSheet(rowIndex) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPay = ss.getSheetByName(CONFIG.SHEET_PAYMENT);
  var row = parseInt(rowIndex);

  if (sheetPay && row > 1) {
    var cellCheck = sheetPay.getRange(row, 1).getValue();
    if (!cellCheck || cellCheck.toString().trim() === "") {
      return "⚠️ Baris ini sudah kosong atau tidak ditemukan. Silakan refresh halaman.";
    }
    sheetPay.deleteRow(row);
    invalidateCache();
    return "✅ Log Transaksi berhasil dihapus secara permanen dari Sheet!";
  }
  return "❌ Gagal menghapus transaksi.";
}

// 4. BULK HAPUS PEMBAYARAN
function bulkHapusPembayaranDariSheet(rowIndicesList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetPay = ss.getSheetByName(CONFIG.SHEET_PAYMENT);
  if (!sheetPay) return "❌ Sheet tidak ditemukan.";

  var currentData = sheetPay.getDataRange().getValues();
  var totalRows = currentData.length;

  var validRows = rowIndicesList.filter(function(row) { return row > 1 && row <= totalRows; });
  if (validRows.length === 0) {
    return "⚠️ Tidak ada baris valid yang bisa dihapus. Mungkin data sudah berubah. Silakan refresh halaman.";
  }

  validRows.sort(function(a, b) { return b - a; });

  var countDeleted = 0;
  validRows.forEach(function(row) {
    var cellVal = sheetPay.getRange(row, 1).getValue();
    if (cellVal && cellVal.toString().trim() !== "") {
      sheetPay.deleteRow(row);
      countDeleted++;
    }
  });

  invalidateCache();
  return "🗑 Sukses menghapus " + countDeleted + " log transaksi dari spreadsheet!";
}