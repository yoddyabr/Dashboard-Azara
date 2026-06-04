// ==========================================
// AKSESSERVICE.GS - Kontrol akses berbasis akun Google
// Sheet "Data Akses": Email | Nama | Role | IDGuru | Status
// Role: management | staff | guru
// ==========================================

// Daftar semua modul (key = nama view di switchView)
var SEMUA_MODUL = [
  'dashboard','data-siswa','data-pembayaran','data-guru','schedule',
  'absensi','rekap-honor','data-mapel','data-kelas','data-ruangan','manajemen-akses'
];

function getModulByRole_(role) {
  role = (role || '').toString().trim().toLowerCase();
  if (role === 'management') {
    return SEMUA_MODUL.slice();
  }
  if (role === 'guru') {
    return ['schedule','absensi'];
  }
  if (role === 'staff') {
    // Semua kecuali dashboard dan manajemen-akses
    return SEMUA_MODUL.filter(function(m){ return m !== 'dashboard' && m !== 'manajemen-akses'; });
  }
  return [];
}

function ensureSheetAkses_(seedEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_AKSES);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_AKSES);
    sheet.appendRow(['Email','Nama','Role','IDGuru','Status']);
    // Seed user pertama (yang setup) sebagai management
    if (seedEmail) {
      sheet.appendRow([seedEmail, 'Admin Utama', 'management', '', 'Aktif']);
    }
  }
  return sheet;
}

// ==========================================
// CEK AKSES USER YANG SEDANG LOGIN
// ==========================================
function getCurrentUserAkses() {
  var email = '';
  try { email = Session.getActiveUser().getEmail() || ''; } catch(e) {}
  if (!email) {
    try { email = Session.getEffectiveUser().getEmail() || ''; } catch(e) {}
  }

  var sheet = ensureSheetAkses_(email);  // auto-seed pakai email pertama

  if (!email) {
    return { email: '', role: 'none', nama: '', idGuru: '', status: '', modulList: [], pesan: 'no_email' };
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === email.toLowerCase()) {
      var role = (data[i][2] || '').toString().trim().toLowerCase();
      var status = (data[i][4] || '').toString().trim();
      var aktif = status.toLowerCase() === 'aktif';
      return {
        email    : email,
        nama     : data[i][1] ? data[i][1].toString() : '',
        role     : role,
        idGuru   : data[i][3] ? data[i][3].toString().trim() : '',
        status   : status,
        modulList: aktif ? getModulByRole_(role) : []
      };
    }
  }
  // Email tidak terdaftar
  return { email: email, role: 'none', nama: '', idGuru: '', status: '', modulList: [] };
}

// ==========================================
// CRUD DATA AKSES (management only — enforcement frontend)
// ==========================================
function getDaftarAkses() {
  var sheet = ensureSheetAkses_('');
  var data = sheet.getDataRange().getValues();
  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    list.push({
      rowIndex: i + 1,
      email   : data[i][0].toString().trim(),
      nama    : data[i][1] ? data[i][1].toString() : '',
      role    : (data[i][2] || '').toString().trim().toLowerCase(),
      idGuru  : data[i][3] ? data[i][3].toString().trim() : '',
      status  : data[i][4] ? data[i][4].toString().trim() : 'Aktif'
    });
  }
  return list;
}

function simpanAkses(payload) {
  if (!payload.email || payload.email.toString().trim() === '')
    return '\u274c Email tidak boleh kosong!';
  if (!payload.role || payload.role.toString().trim() === '')
    return '\u274c Role tidak boleh kosong!';

  var email = payload.email.toString().trim().toLowerCase();
  var sheet = ensureSheetAkses_('');

  if (payload.rowIndex && parseInt(payload.rowIndex) > 1) {
    var row = parseInt(payload.rowIndex);
    sheet.getRange(row, 1).setValue(email);
    sheet.getRange(row, 2).setValue(payload.nama || '');
    sheet.getRange(row, 3).setValue(payload.role);
    sheet.getRange(row, 4).setValue(payload.idGuru || '');
    sheet.getRange(row, 5).setValue(payload.status || 'Aktif');
    return '\u2705 Akses diperbarui.';
  }

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === email) {
      return '\u26a0\ufe0f Email ' + email + ' sudah terdaftar!';
    }
  }
  sheet.appendRow([email, payload.nama || '', payload.role, payload.idGuru || '', payload.status || 'Aktif']);
  return '\u2705 Akses ditambahkan.';
}

function hapusAkses(rowIndex) {
  var sheet = ensureSheetAkses_('');
  var row = parseInt(rowIndex);
  if (row > 1) {
    // Cegah hapus management terakhir
    var data = sheet.getDataRange().getValues();
    var jmlManagementAktif = 0;
    for (var i = 1; i < data.length; i++) {
      if ((data[i][2]||'').toString().trim().toLowerCase() === 'management'
          && (data[i][4]||'').toString().trim().toLowerCase() === 'aktif') {
        jmlManagementAktif++;
      }
    }
    var targetRole = (data[row-1][2]||'').toString().trim().toLowerCase();
    if (targetRole === 'management' && jmlManagementAktif <= 1) {
      return '\u26a0\ufe0f Tidak bisa hapus management terakhir! Harus ada minimal 1 management aktif.';
    }
    sheet.deleteRow(row);
    return '\u2705 Akses dihapus.';
  }
  return '\u274c Gagal menghapus.';
}