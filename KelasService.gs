// ==========================================
// KELASSERVICE.GS - Naik Kelas Tahunan (berbasis TINGKAT + Rombel)
// Saat naik kelas: tiap siswa naik 1 tingkat, rombel di-reset ke default
// (1 tingkat = 1 rombel, format "Jenjang Tingkat", mis. "SMA 11").
// Data Kelas juga di-reset jadi 1 rombel per tingkat.
// ==========================================

// Jenjang standar dari nomor tingkat (fallback kalau tidak ada di Data Kelas)
function jenjangDariTingkat_(t) {
  t = parseInt(t);
  if (t >= 1 && t <= 6)  return 'SD';
  if (t >= 7 && t <= 9)  return 'SMP';
  if (t >= 10 && t <= 12) return 'SMA';
  return '';
}

// Map tingkat → jenjang dari Data Kelas
function buildTingkatJenjangMap_(ss) {
  var map = {};
  var sheet = ss.getSheetByName(CONFIG.SHEET_KELAS);
  if (sheet) {
    var d = sheet.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (!d[i][1]) continue;
      var t = d[i][1].toString().trim();
      var j = d[i][0] ? d[i][0].toString().trim() : '';
      if (t && j && !map[t]) map[t] = j;
    }
  }
  return map;
}

// Map rombel(lower) → tingkat dari Data Kelas
function buildRombelTingkatMap_(ss) {
  var map = {};
  var sheet = ss.getSheetByName(CONFIG.SHEET_KELAS);
  if (sheet) {
    var d = sheet.getDataRange().getValues();
    for (var i = 1; i < d.length; i++) {
      if (!d[i][2]) continue;
      map[d[i][2].toString().trim().toLowerCase()] = d[i][1] ? d[i][1].toString().trim() : '';
    }
  }
  return map;
}

// Derive nomor tingkat dari rombel (lookup Data Kelas → fallback regex angka)
function tingkatDariRombel_(rombel, rombelMap) {
  if (!rombel) return null;
  var s = rombel.toString().trim();
  var mapped = rombelMap[s.toLowerCase()];
  if (mapped) { var n = parseInt(mapped); return isNaN(n) ? null : n; }
  var m = s.match(/\d+/);
  if (m) return parseInt(m[0]);
  return null;
}

// Default rombel untuk sebuah tingkat = "Jenjang Tingkat" (mis. "SMA 11")
function defaultRombel_(tingkat, tjMap) {
  var j = tjMap[tingkat.toString()] || jenjangDariTingkat_(tingkat);
  return (j ? j + ' ' : '') + tingkat;
}

// ==========================================
// 1. PREVIEW
// ==========================================
function previewNaikKelas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rombelMap = buildRombelTingkatMap_(ss);
  var tjMap     = buildTingkatJenjangMap_(ss);
  var sheetSiswa = ss.getSheetByName(CONFIG.SHEET_SISWA) || ss.getSheets()[0];
  var dataSiswa = sheetSiswa.getDataRange().getValues();
  var preview = [];

  for (var i = 1; i < dataSiswa.length; i++) {
    var row = dataSiswa[i];
    if (!row[0] || row[0].toString().trim() === "") continue;
    var id    = row[0].toString().trim();
    var nama  = row[1] ? row[1].toString().trim() : "";
    var kelas = row[2] ? row[2].toString().trim() : "";

    var t = tingkatDariRombel_(kelas, rombelMap);
    if (t === null) continue;

    var lulus = (t >= 12);
    var kelasBaru = lulus ? kelas : defaultRombel_(t + 1, tjMap);

    preview.push({
      rowIndex : i + 1,
      id       : id,
      nama     : nama,
      kelasLama: kelas,
      kelasBaru: kelasBaru,
      lulus    : lulus
    });
  }
  return preview;
}

// ==========================================
// 2. EKSEKUSI
// ==========================================
function eksekusiNaikKelas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var rombelMap = buildRombelTingkatMap_(ss);
  var tjMap     = buildTingkatJenjangMap_(ss);
  var sheetSiswa = ss.getSheetByName(CONFIG.SHEET_SISWA) || ss.getSheets()[0];
  // Paksa kolom Kelas siswa (C) jadi teks (safety dari auto-format)
  sheetSiswa.getRange(1, 3, sheetSiswa.getMaxRows(), 1).setNumberFormat('@');
  var dataSiswa = sheetSiswa.getDataRange().getValues();
  var countNaik = 0, countLulus = 0;

  for (var i = 1; i < dataSiswa.length; i++) {
    var row = dataSiswa[i];
    if (!row[0] || row[0].toString().trim() === "") continue;
    var kelas = row[2] ? row[2].toString().trim() : "";
    var t = tingkatDariRombel_(kelas, rombelMap);
    if (t === null) continue;

    if (t >= 12) {
      sheetSiswa.getRange(i + 1, 4).setValue("Lulus");
      countLulus++;
    } else {
      sheetSiswa.getRange(i + 1, 3).setValue(defaultRombel_(t + 1, tjMap));
      countNaik++;
    }
  }

  // Reset Data Kelas → 1 rombel default per tingkat
  resetDataKelasDefault_(ss);

  invalidateCache();
  try { invalidateMasterCache(); } catch(e) {}
  return { countNaik: countNaik, countLulus: countLulus };
}

// ==========================================
// Reset Data Kelas: tiap tingkat jadi 1 rombel default ("Jenjang Tingkat")
// ==========================================
function resetDataKelasDefault_(ss) {
  var sheet = ss.getSheetByName(CONFIG.SHEET_KELAS);
  if (!sheet) return;
  var d = sheet.getDataRange().getValues();
  var seen = {}, list = [];
  for (var i = 1; i < d.length; i++) {
    var j = d[i][0] ? d[i][0].toString().trim() : '';
    var t = d[i][1] ? d[i][1].toString().trim() : '';
    if (!t) continue;
    var key = j + '||' + t;
    if (!seen[key]) { seen[key] = true; list.push({ jenjang: j, tingkat: t }); }
  }
  list.sort(function(a, b) {
    var na = parseInt(a.tingkat), nb = parseInt(b.tingkat);
    if (!isNaN(na) && !isNaN(nb)) return na - nb;
    return a.tingkat.localeCompare(b.tingkat);
  });
  sheet.clearContents();
  sheet.getRange(1, 1, 1, 3).setValues([['Jenjang','Tingkat','Rombel']]);
  if (list.length > 0) {
    var rows = list.map(function(x) {
      return [x.jenjang, x.tingkat, (x.jenjang ? x.jenjang + ' ' : '') + x.tingkat];
    });
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}