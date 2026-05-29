// ==========================================
// ABSENSISERVICE.GS - Absensi per sesi jadwal
// Sheet "Data Absensi": Tanggal | Rombel | Mapel | JamMulai
//   | IDGuru | NamaGuru | IDSiswa | NamaSiswa | Status | Catatan | Timestamp
// ==========================================

function ensureSheetAbsensi_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_ABSENSI);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_ABSENSI);
    sheet.appendRow(['Tanggal','Rombel','Mapel','JamMulai','IDGuru','NamaGuru','IDSiswa','NamaSiswa','Status','Catatan','Timestamp','Nilai','Materi']);
    return sheet;
  }
  // Migrasi: tambah kolom Nilai (12) + Materi (13) kalau belum ada
  var lastCol = sheet.getLastColumn();
  if (lastCol < 12) {
    sheet.getRange(1, 12).setValue('Nilai');
  }
  if (lastCol < 13) {
    sheet.getRange(1, 13).setValue('Materi');
  }
  return sheet;
}

// Normalisasi tanggal jadi string "YYYY-MM-DD" supaya konsisten saat compare
function normTgl_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var y = val.getFullYear(), m = val.getMonth()+1, d = val.getDate();
    return y + '-' + (m<10?'0':'') + m + '-' + (d<10?'0':'') + d;
  }
  var s = val.toString().trim();
  // sudah format YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0,10);
  var dt = new Date(s);
  if (!isNaN(dt.getTime())) return normTgl_(dt);
  return s;
}

function normJam_(val) {
  if (!val) return '';
  if (val instanceof Date) {
    var h = val.getHours(), m = val.getMinutes();
    return (h<10?'0':'') + h + ':' + (m<10?'0':'') + m;
  }
  return val.toString().trim();
}

// ==========================================
// AMBIL ABSENSI 1 SESI (untuk preload saat buka modal)
// Return: array of { idSiswa, status, catatan }
// ==========================================
// Return: { materi, adaNilai, entries: [{ idSiswa, status, catatan, nilai }] }
function getAbsensiSesi(tanggal, rombel, jamMulai) {
  var sheet = ensureSheetAbsensi_();
  var data = sheet.getDataRange().getValues();
  var entries = [];
  var materi = '';
  var adaNilai = false;
  var tglKey = normTgl_(tanggal);
  var jamKey = normJam_(jamMulai);
  var rombelKey = (rombel||'').toString().trim().toLowerCase();
  for (var i = 1; i < data.length; i++) {
    if (normTgl_(data[i][0]) === tglKey
        && (data[i][1]||'').toString().trim().toLowerCase() === rombelKey
        && normJam_(data[i][3]) === jamKey) {
      var nilai = data[i][11];
      var nilaiVal = (nilai !== '' && nilai !== null && nilai !== undefined && !isNaN(nilai)) ? Number(nilai) : '';
      if (nilaiVal !== '' && nilaiVal !== null) adaNilai = true;
      entries.push({
        idSiswa : data[i][6] ? data[i][6].toString().trim() : '',
        status  : data[i][8] ? data[i][8].toString().trim() : 'Hadir',
        catatan : data[i][9] ? data[i][9].toString().trim() : '',
        nilai   : nilaiVal
      });
      // Materi sama untuk semua baris satu sesi, ambil dari baris pertama yang ada
      if (!materi && data[i][12]) materi = data[i][12].toString();
    }
  }
  return { materi: materi, adaNilai: adaNilai, entries: entries };
}

// ==========================================
// SIMPAN ABSENSI 1 SESI (upsert: hapus lama, insert baru)
// payload = { tanggal, rombel, mapel, jamMulai, idGuru, namaGuru, entries: [{idSiswa, namaSiswa, status, catatan}] }
// ==========================================
function simpanAbsensi(payload) {
  if (!payload || !payload.tanggal || !payload.rombel || !payload.jamMulai) {
    return '\u274c Data sesi tidak lengkap (tanggal, rombel, jam wajib).';
  }
  if (!payload.entries || payload.entries.length === 0) {
    return '\u274c Tidak ada data siswa untuk disimpan.';
  }

  var sheet = ensureSheetAbsensi_();
  var data = sheet.getDataRange().getValues();
  var tglKey = normTgl_(payload.tanggal);
  var jamKey = normJam_(payload.jamMulai);
  var rombelKey = (payload.rombel||'').toString().trim().toLowerCase();

  // Hapus baris lama untuk sesi ini (dari bawah ke atas biar index aman)
  var hapusCount = 0;
  for (var i = data.length - 1; i >= 1; i--) {
    if (normTgl_(data[i][0]) === tglKey
        && (data[i][1]||'').toString().trim().toLowerCase() === rombelKey
        && normJam_(data[i][3]) === jamKey) {
      sheet.deleteRow(i + 1);
      hapusCount++;
    }
  }

  // Insert baris baru
  var now = new Date();
  var rows = [];
  var materi = payload.materi || '';
  payload.entries.forEach(function(e) {
    var nilai = (e.nilai !== '' && e.nilai !== null && e.nilai !== undefined && !isNaN(e.nilai)) ? Number(e.nilai) : '';
    rows.push([
      tglKey, payload.rombel, payload.mapel || '', jamKey,
      payload.idGuru || '', payload.namaGuru || '',
      e.idSiswa || '', e.namaSiswa || '',
      e.status || 'Hadir', e.catatan || '',
      now, nilai, materi
    ]);
  });
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }

  return '\u2705 Absensi tersimpan (' + rows.length + ' siswa).';
}


// ==========================================
// CEK STATUS ABSENSI BULK SESI (untuk badge di list jadwal guru)
// Return: map { "tglKey|rombelKey|jamKey" : { jumlahHadir, jumlahTotal, adaMateri } }
// ==========================================
function getStatusAbsensiBulkSesi(idGuru, bulan, tahun) {
  var sheet = ensureSheetAbsensi_();
  var data = sheet.getDataRange().getValues();
  var map = {};
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var tgl = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
    if (isNaN(tgl.getTime())) continue;
    if ((tgl.getMonth()+1) !== parseInt(bulan) || tgl.getFullYear() !== parseInt(tahun)) continue;
    if (idGuru && (data[i][4]||'').toString().trim() !== idGuru.toString().trim()) continue;
    var key = normTgl_(data[i][0]) + '|' + (data[i][1]||'').toString().trim().toLowerCase() + '|' + normJam_(data[i][3]);
    if (!map[key]) map[key] = { jumlahHadir: 0, jumlahTotal: 0, adaMateri: false };
    map[key].jumlahTotal++;
    if ((data[i][8]||'').toString().toLowerCase() === 'hadir') map[key].jumlahHadir++;
    if (data[i][12]) map[key].adaMateri = true;
  }
  return map;
}

// ==========================================
// REKAP ABSENSI per siswa (bulan/tahun, opsional filter rombel)
// Return: [{ idSiswa, namaSiswa, kelas, totalSesi, hadir, izin, sakit, alpha, persenHadir }]
// ==========================================
function getRekapAbsensi(bulan, tahun, rombel) {
  var sheet = ensureSheetAbsensi_();
  var data = sheet.getDataRange().getValues();
  var map = {};
  var filterRombel = rombel ? rombel.toString().trim().toLowerCase() : '';

  for (var i = 1; i < data.length; i++) {
    if (!data[i][6]) continue;  // ID Siswa wajib
    var tgl = data[i][0] instanceof Date ? data[i][0] : new Date(data[i][0]);
    if (isNaN(tgl.getTime())) continue;
    if ((tgl.getMonth() + 1) !== parseInt(bulan) || tgl.getFullYear() !== parseInt(tahun)) continue;

    var rombelRow = (data[i][1]||'').toString().trim();
    if (filterRombel && rombelRow.toLowerCase() !== filterRombel) continue;

    var idSiswa = data[i][6].toString().trim();
    var nama    = data[i][7] ? data[i][7].toString().trim() : '';
    var status  = (data[i][8]||'').toString().trim().toLowerCase();

    if (!map[idSiswa]) {
      map[idSiswa] = { idSiswa: idSiswa, namaSiswa: nama, kelas: rombelRow, totalSesi: 0, hadir: 0, izin: 0, sakit: 0, alpha: 0 };
    }
    map[idSiswa].totalSesi++;
    if (status === 'hadir') map[idSiswa].hadir++;
    else if (status === 'izin') map[idSiswa].izin++;
    else if (status === 'sakit') map[idSiswa].sakit++;
    else if (status === 'alpha' || status === 'alpa') map[idSiswa].alpha++;
  }

  var result = Object.values(map).map(function(r) {
    r.persenHadir = r.totalSesi > 0 ? Math.round((r.hadir / r.totalSesi) * 100) : 0;
    return r;
  });
  result.sort(function(a, b) { return a.namaSiswa.localeCompare(b.namaSiswa); });
  return result;
}