// ==========================================
// SISWASERVICE.GS - CRUD Data Siswa
// ✅ OPT: invalidateCache() setiap operasi tulis
// ==========================================

// 1. SIMPAN / EDIT DATA SISWA
function simpanSiswaBaru(payload) {
  if (!payload.nama || payload.nama.toString().trim() === "") return "❌ Gagal: Nama siswa tidak boleh kosong!";
  if (!payload.kelas || payload.kelas.toString().trim() === "") return "❌ Gagal: Kelas tidak boleh kosong!";
  if (!payload.biaya || isNaN(payload.biaya) || Number(payload.biaya) <= 0) return "❌ Gagal: Biaya harus diisi dengan angka yang valid!";

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSiswa = ss.getSheetByName(CONFIG.SHEET_SISWA) || ss.getSheets()[0];
  var dataSiswa = sheetSiswa.getDataRange().getValues();
  var idTarget = payload.id ? payload.id.toString().trim().toLowerCase() : "";
  var barisKetemu = -1;

  if (idTarget !== "") {
    for (var i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][0] && dataSiswa[i][0].toString().trim().toLowerCase() === idTarget) {
        barisKetemu = i + 1;
        break;
      }
    }
  }

  if (barisKetemu !== -1) {
    sheetSiswa.getRange(barisKetemu, 2).setValue(payload.nama);
    sheetSiswa.getRange(barisKetemu, 3).setValue(payload.kelas);
    sheetSiswa.getRange(barisKetemu, 4).setValue(payload.status);
    sheetSiswa.getRange(barisKetemu, 5).setValue(payload.tipeBayar);
    sheetSiswa.getRange(barisKetemu, 6).setValue(Number(payload.biaya));
    invalidateCache();
    return "✅ Data Siswa " + payload.id + " berhasil diperbarui!";
  } else {
    var namaBaru = payload.nama.toString().trim().toLowerCase();
    for (var k = 1; k < dataSiswa.length; k++) {
      if (dataSiswa[k][1] && dataSiswa[k][1].toString().trim().toLowerCase() === namaBaru) {
        return "⚠️ Peringatan: Siswa dengan nama \"" + payload.nama + "\" sudah terdaftar! Pastikan ini bukan duplikat.";
      }
    }
    var maxId = 0;
    for (var m = 1; m < dataSiswa.length; m++) {
      if (dataSiswa[m][0]) {
        var cleanId = parseInt(dataSiswa[m][0].toString().replace(/\D/g, ""));
        if (!isNaN(cleanId) && cleanId > maxId) maxId = cleanId;
      }
    }
    var nextId = "AZA-" + ("000" + (maxId + 1)).slice(-4);
    sheetSiswa.appendRow([nextId, payload.nama, payload.kelas, payload.status, payload.tipeBayar, Number(payload.biaya)]);
    invalidateCache();
    return "🎉 Siswa Baru Berhasil Didaftar dengan ID: " + nextId;
  }
}

// 2. HAPUS SINGLE SISWA
function hapusSiswaDariSheet(idSiswa) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSiswa = ss.getSheetByName(CONFIG.SHEET_SISWA) || ss.getSheets()[0];
  var dataSiswa = sheetSiswa.getDataRange().getValues();
  var idTarget = idSiswa ? idSiswa.toString().trim().toLowerCase() : "";
  var barisKetemu = -1;

  for (var i = 1; i < dataSiswa.length; i++) {
    if (dataSiswa[i][0] && dataSiswa[i][0].toString().trim().toLowerCase() === idTarget) {
      barisKetemu = i + 1;
      break;
    }
  }

  if (barisKetemu !== -1) {
    sheetSiswa.deleteRow(barisKetemu);
    invalidateCache();
    return "✅ Data Siswa dengan ID " + idSiswa + " telah berhasil dihapus permanen!";
  }
  return "❌ Gagal: Data siswa tidak ditemukan.";
}

// 3. BULK HAPUS SISWA
function bulkHapusSiswaDariSheet(idList) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetSiswa = ss.getSheetByName(CONFIG.SHEET_SISWA) || ss.getSheets()[0];
  var setTarget = new Set(idList.map(function(id) { return id.toString().trim().toLowerCase(); }));
  var dataSiswa = sheetSiswa.getDataRange().getValues();
  var countDeleted = 0;

  for (var i = dataSiswa.length - 1; i >= 1; i--) {
    if (dataSiswa[i][0]) {
      var currentId = dataSiswa[i][0].toString().trim().toLowerCase();
      if (setTarget.has(currentId)) {
        sheetSiswa.deleteRow(i + 1);
        countDeleted++;
      }
    }
  }
  invalidateCache();
  return "🗑 Sukses menghapus " + countDeleted + " data siswa dari spreadsheet!";
}

// ==========================================
// BULK PINDAH SISWA KE ROMBEL
// ==========================================
function pindahSiswaKeRombel(rombelTarget, idSiswaArr) {
  if (!rombelTarget || rombelTarget.toString().trim() === '')
    return '\u274c Rombel target tidak boleh kosong.';
  if (!idSiswaArr || idSiswaArr.length === 0)
    return '\u274c Tidak ada siswa yang dipilih.';

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_SISWA) || ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();

  var idSet = {};
  idSiswaArr.forEach(function(id){ idSet[id.toString().trim().toUpperCase()] = true; });

  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var rowId = data[i][0].toString().trim().toUpperCase();
    if (idSet[rowId]) {
      sheet.getRange(i + 1, 3).setValue(rombelTarget);  // kolom C (kelas/rombel)
      count++;
    }
  }
  invalidateCache();
  return '\u2705 ' + count + ' siswa berhasil dipindahkan ke rombel ' + rombelTarget + '.';
}