// =====================================================================
// DASHBOARDSERVICE.GS — "CEO Cockpit" Dashboard Overview Azara
// Satu fungsi getDashboardOverview() mengembalikan SEMUA metrik dashboard.
// Reuse parser yang sudah ada: getSistemData(), getScheduleData(), getGuruData().
// Absensi dibaca langsung (header: Tanggal,Rombel,Mapel,JamMulai,IDGuru,
// NamaGuru,IDSiswa,NamaSiswa,Status,Catatan,Timestamp,...).
// =====================================================================

// ---- Parameter yang bisa disetel (asumsi karena data belum tersimpan) ----
var DOV = {
  TZ: 'Asia/Jakarta',
  DUE_DAY: 10,              // tagihan bulanan jatuh tempo tiap tanggal ini
  TARGET_JAM_GURU: 80,      // target jam mengajar / guru / bulan (utk utilization)
  TARGET_SISWA_ROMBEL: 20,  // target siswa / rombel (utk occupancy; kapasitas blm tersimpan)
  ALPA_BERUNTUN: 3,         // ambang alpa beruntun utk "at risk"
  CACHE_KEY: 'azara_dashboard_overview',
  CACHE_TTL: 120
};

function getDashboardOverview() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(DOV.CACHE_KEY);
  if (hit) { try { return JSON.parse(hit); } catch(e){} }

  var now = new Date();
  var ym      = DOV_fmt_(now, 'yyyy-MM');
  var ymPrev  = DOV_fmt_(DOV_addMonth_(now, -1), 'yyyy-MM');
  var todayS  = DOV_fmt_(now, 'yyyy-MM-dd');
  var dayOfMonth = parseInt(DOV_fmt_(now, 'd'), 10);
  var daysInMonth = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();

  // ---- Ambil data (reuse parser yang ada) ----
  var sis   = (typeof getSistemData === 'function') ? getSistemData() : { siswaList:[], paymentList:[] };
  var siswa = sis.siswaList || [];
  var pay   = sis.paymentList || [];
  var sched = DOV_arr_(safeCall_(getScheduleData), ['scheduleList','data']);
  var guru  = DOV_arr_(safeCall_(getGuruData),     ['guruList','data']);
  var absen = DOV_readAbsensi_();

  // index siswa by id
  var siswaById = {};
  siswa.forEach(function(s){ siswaById[s.id] = s; });
  var aktif = siswa.filter(function(s){ return s.status === 'Aktif'; });

  // ================= SECTION 1: EXECUTIVE SUMMARY =================
  var revThis = DOV_sumPay_(pay, ym), revPrev = DOV_sumPay_(pay, ymPrev);
  var target  = aktif.reduce(function(a,s){ return a + (Number(s.biaya)||0); }, 0); // expektasi tagihan bulanan
  var achievement = target > 0 ? Math.round(revThis / target * 100) : 0;
  var revGrowth   = revPrev > 0 ? Math.round((revThis - revPrev) / revPrev * 100) : (revThis>0?100:0);

  // siswa yang sudah bayar bulan ini
  var paidSet = {};
  pay.forEach(function(p){ if (DOV_fmt_(DOV_date_(p.tgl), 'yyyy-MM') === ym) paidSet[p.id] = true; });
  var overdueSiswa = aktif.filter(function(s){ return !paidSet[s.id]; });
  var outstandingAmt = overdueSiswa.reduce(function(a,s){ return a + (Number(s.biaya)||0); }, 0);
  var cashToday = pay.reduce(function(a,p){ return a + (DOV_fmt_(DOV_date_(p.tgl),'yyyy-MM-dd')===todayS ? (Number(p.nominal)||0) : 0); }, 0);

  // new students this month (perlu tgl daftar — tidak tersimpan; pakai pembayaran pertama sbg proksi)
  var newThisMonth = DOV_newStudentsProxy_(pay, siswaById, ym);

  // attendance
  var attThis = DOV_attRate_(absen, ym), attPrev = DOV_attRate_(absen, ymPrev);

  // teacher hours bulan ini
  var hoursByGuru = {};
  sched.forEach(function(x){
    if (DOV_fmt_(DOV_date_(x.tgl),'yyyy-MM') !== ym) return;
    var h = DOV_jam_(x.durasi);
    var key = x.idGuru || x.namaGuru || '-';
    hoursByGuru[key] = (hoursByGuru[key]||0) + h;
    if (!hoursByGuru['_nama_'+key]) hoursByGuru['_nama_'+key] = x.namaGuru || x.idGuru;
  });
  var totalHours = 0; Object.keys(hoursByGuru).forEach(function(k){ if(k.indexOf('_nama_')!==0) totalHours += hoursByGuru[k]; });
  var guruCount = guru.length || Object.keys(hoursByGuru).filter(function(k){return k.indexOf('_nama_')!==0;}).length || 1;
  var utilization = Math.round(totalHours / (guruCount * DOV.TARGET_JAM_GURU) * 100);

  // ================= SECTION 2: BUSINESS HEALTH SCORE =================
  var collectionRate = aktif.length ? Math.round((aktif.length - overdueSiswa.length) / aktif.length * 100) : 0;
  var retention = siswa.length ? Math.round(aktif.length / siswa.length * 100) : 0; // proksi: aktif / total
  var occByRombel = DOV_countByClass_(aktif);
  var rombelKeys = Object.keys(occByRombel);
  var occupancy = rombelKeys.length
    ? Math.round(rombelKeys.reduce(function(a,k){ return a + Math.min(100, occByRombel[k]/DOV.TARGET_SISWA_ROMBEL*100); },0) / rombelKeys.length)
    : 0;
  var healthBreak = {
    revenue: Math.min(100, achievement),
    attendance: attThis.rate,
    collection: collectionRate,
    retention: retention,
    occupancy: occupancy
  };
  var score = Math.round((healthBreak.revenue + healthBreak.attendance + healthBreak.collection + healthBreak.retention + healthBreak.occupancy) / 5);
  var healthStatus = score >= 75 ? 'Healthy' : (score >= 50 ? 'Need Attention' : 'Critical');

  // ================= SECTION 3: REVENUE ANALYTICS =================
  var trend = DOV_revTrend12_(pay, now);
  var byLevel = DOV_revByLevel_(pay, siswaById, ym);
  var byClassMap = {};
  pay.forEach(function(p){
    if (DOV_fmt_(DOV_date_(p.tgl),'yyyy-MM') !== ym) return;
    var s = siswaById[p.id]; var k = s ? s.kelas : (p.kelas||'Lainnya');
    byClassMap[k] = (byClassMap[k]||0) + (Number(p.nominal)||0);
  });
  var byClass = DOV_topEntries_(byClassMap, 7);
  var forecast = dayOfMonth > 0 ? Math.round(revThis / dayOfMonth * daysInMonth) : revThis;

  // ================= SECTION 4: STUDENT ANALYTICS =================
  var atRisk = DOV_studentsAtRisk_(aktif, absen, paidSet, ym);

  // ================= SECTION 5: PAYMENT MONITORING =================
  var overdueTop = overdueSiswa.map(function(s){
    return { id:s.id, nama:s.nama, kelas:s.kelas, amount:(Number(s.biaya)||0), daysOverdue: DOV_daysOverdue_(now) };
  }).sort(function(a,b){ return b.amount - a.amount; }).slice(0, 10);
  var statusDist = { paid: aktif.length - overdueSiswa.length, partial: 0, overdue: overdueSiswa.length };
  // partial: bayar < biaya bulan ini
  aktif.forEach(function(s){
    if (!paidSet[s.id]) return;
    var bayar = pay.filter(function(p){ return p.id===s.id && DOV_fmt_(DOV_date_(p.tgl),'yyyy-MM')===ym; })
                   .reduce(function(a,p){ return a+(Number(p.nominal)||0); },0);
    if (s.biaya && bayar < s.biaya) { statusDist.partial++; statusDist.paid--; }
  });

  // ================= SECTION 6: TEACHER ANALYTICS =================
  var workload = guru.map(function(g){
    return { id:g.id, nama:g.nama, hours: Math.round((hoursByGuru[g.id]||0)*10)/10 };
  }).sort(function(a,b){ return b.hours - a.hours; });
  var topTeachers = workload.slice(0, 5); // proksi: berdasarkan jam mengajar (beban tertinggi)
  // guru tanpa jadwal minggu depan
  var next7 = {};
  sched.forEach(function(x){
    var d = DOV_date_(x.tgl); var diff = (d - now) / 86400000;
    if (diff >= 0 && diff <= 7) next7[x.idGuru || x.namaGuru] = true;
  });
  var noSchedule = guru.filter(function(g){ return !next7[g.id] && !next7[g.nama]; })
                       .map(function(g){ return { id:g.id, nama:g.nama, mapel:g.mapel }; });

  // ================= SECTION 7: CLASS PERFORMANCE =================
  var classCounts = DOV_topEntries_(occByRombel, 100);
  var popular = classCounts.slice(0, 5).map(function(e){ return { kelas:e.key, count:e.value }; });
  var underutilized = classCounts.filter(function(e){ return e.value < DOV.TARGET_SISWA_ROMBEL; })
                                 .map(function(e){ return { kelas:e.key, count:e.value, target:DOV.TARGET_SISWA_ROMBEL }; });

  // ================= SECTION 8: TODAY'S ACTION CENTER =================
  var classesToday = sched.filter(function(x){ return DOV_fmt_(DOV_date_(x.tgl),'yyyy-MM-dd')===todayS; })
    .map(function(x){ return { jam:x.jamMulai, mapel:x.mapel, kelas:x.kelas, guru:x.namaGuru, ruang:x.ruangan }; })
    .sort(function(a,b){ return (a.jam||'').localeCompare(b.jam||''); });
  var consecutiveAbsence = DOV_consecutiveAbsence_(absen, DOV.ALPA_BERUNTUN);
  var promotions = aktif.filter(function(s){ return DOV_tingkat_(s.kelas) >= 12; })
                        .map(function(s){ return { id:s.id, nama:s.nama, kelas:s.kelas }; }); // SMA 12 ~ lulus

  // ================= SECTION 10: NOTIFICATIONS =================
  var notifs = [];
  if (overdueSiswa.length) notifs.push({ level:'danger', icon:'fa-money-bill-wave', text: overdueSiswa.length + ' siswa belum bayar bulan ini (' + DOV_rp_(outstandingAmt) + ')' });
  var conflicts = DOV_scheduleConflicts_(sched, todayS);
  conflicts.forEach(function(c){ notifs.push({ level:'warning', icon:'fa-triangle-exclamation', text: c }); });
  var missing = DOV_missingAttendance_(sched, absen, now);
  if (missing > 0) notifs.push({ level:'warning', icon:'fa-clipboard-question', text: missing + ' sesi lampau belum diisi absensi' });
  underutilized.slice(0,3).forEach(function(u){ notifs.push({ level:'info', icon:'fa-users-slash', text: 'Kelas ' + u.kelas + ' di bawah target ('+u.count+'/'+u.target+' siswa)' }); });
  if (newThisMonth > 0) notifs.push({ level:'success', icon:'fa-user-plus', text: newThisMonth + ' siswa baru bulan ini (estimasi)' });

  var result = {
    generatedAt: DOV_fmt_(now, "EEEE, d MMM yyyy HH:mm"),
    month: ym,
    exec: {
      revenue: { current: revThis, target: target, achievement: achievement, prevMonth: revPrev, growth: revGrowth },
      students: { active: aktif.length, total: siswa.length, newThisMonth: newThisMonth },
      outstanding: { amount: outstandingAmt, overdueCount: overdueSiswa.length },
      attendance: { rate: attThis.rate, prevRate: attPrev.rate, samples: attThis.total },
      teacher: { totalHours: Math.round(totalHours*10)/10, utilization: utilization, count: guruCount },
      cashToday: cashToday
    },
    health: { score: score, status: healthStatus, breakdown: healthBreak },
    revenue: { trend: trend, byLevel: byLevel, byClass: byClass, forecast: forecast },
    students: { retentionRate: retention, atRisk: atRisk },
    payment: { collectionRate: collectionRate, overdueTop: overdueTop, statusDist: statusDist },
    teacher: { workload: workload, topTeachers: topTeachers, noSchedule: noSchedule },
    classPerf: { popular: popular, underutilized: underutilized, occupancyRate: occupancy },
    actions: { classesToday: classesToday, consecutiveAbsence: consecutiveAbsence, promotions: promotions, overdueCount: overdueSiswa.length },
    notifications: notifs,
    _assumptions: {
      dueDay: DOV.DUE_DAY, targetJamGuru: DOV.TARGET_JAM_GURU, targetSiswaRombel: DOV.TARGET_SISWA_ROMBEL
    }
  };

  try { cache.put(DOV.CACHE_KEY, JSON.stringify(result), DOV.CACHE_TTL); } catch(e){}
  return result;
}

// ===================== HELPERS =====================
function safeCall_(fn){ try { return (typeof fn==='function') ? fn() : null; } catch(e){ return null; } }
function DOV_arr_(v, keys){
  if (!v) return [];
  if (Array.isArray(v)) return v;
  for (var i=0;i<keys.length;i++){ if (Array.isArray(v[keys[i]])) return v[keys[i]]; }
  return [];
}
function DOV_fmt_(d, pat){ if(!(d instanceof Date)||isNaN(d)) return ''; return Utilities.formatDate(d, DOV.TZ, pat); }
function DOV_date_(v){
  if (v instanceof Date) return v;
  if (!v) return new Date('invalid');
  var d = new Date(v); return d;
}
function DOV_addMonth_(d, n){ return new Date(d.getFullYear(), d.getMonth()+n, 1); }
function DOV_sumPay_(pay, ym){
  return pay.reduce(function(a,p){ return a + (DOV_fmt_(DOV_date_(p.tgl),'yyyy-MM')===ym ? (Number(p.nominal)||0) : 0); }, 0);
}
function DOV_rp_(n){ return 'Rp ' + (Number(n)||0).toLocaleString('id-ID'); }
function DOV_jam_(durasi){
  var n = Number(durasi); if (isNaN(n)||n<=0) return 0;
  return n > 12 ? n/60 : n; // >12 dianggap menit, selainnya jam
}
function DOV_tingkat_(kelas){
  if (!kelas) return 0; var m = kelas.toString().match(/\d+/); return m ? parseInt(m[0],10) : 0;
}
function DOV_jenjang_(kelas){
  var s = (kelas||'').toString().toUpperCase();
  if (s.indexOf('SMA')>=0) return 'SMA'; if (s.indexOf('SMP')>=0) return 'SMP'; if (s.indexOf('SD')>=0) return 'SD';
  var t = DOV_tingkat_(kelas);
  if (t>=10) return 'SMA'; if (t>=7) return 'SMP'; if (t>=1) return 'SD'; return 'Lainnya';
}
function DOV_countByClass_(list){
  var m = {}; list.forEach(function(s){ var k=s.kelas||'Tanpa Kelas'; m[k]=(m[k]||0)+1; }); return m;
}
function DOV_topEntries_(map, n){
  return Object.keys(map).map(function(k){ return {key:k, value:map[k]}; })
              .sort(function(a,b){ return b.value-a.value; }).slice(0,n);
}
function DOV_daysOverdue_(now){
  var due = new Date(now.getFullYear(), now.getMonth(), DOV.DUE_DAY);
  var diff = Math.floor((now - due)/86400000);
  return diff > 0 ? diff : 0;
}
function DOV_readAbsensi_(){
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CONFIG.SHEET_ABSENSI);
    if (!sh || sh.getLastRow() < 2) return [];
    var d = sh.getDataRange().getValues();
    var out = [];
    for (var i=1;i<d.length;i++){
      if (!d[i][0] && !d[i][6]) continue;
      out.push({ tgl:d[i][0], rombel:d[i][1], mapel:d[i][2], idGuru:d[i][4], idSiswa:d[i][6], namaSiswa:d[i][7], status:(d[i][8]||'').toString().trim() });
    }
    return out;
  } catch(e){ return []; }
}
function DOV_attRate_(absen, ym){
  var hadir=0, total=0;
  absen.forEach(function(a){
    if (DOV_fmt_(DOV_date_(a.tgl),'yyyy-MM') !== ym) return;
    total++; if ((a.status||'').toLowerCase()==='hadir') hadir++;
  });
  return { rate: total ? Math.round(hadir/total*100) : 0, total: total };
}
function DOV_revTrend12_(pay, now){
  var labels=[], data=[];
  for (var i=11;i>=0;i--){
    var d = new Date(now.getFullYear(), now.getMonth()-i, 1);
    var ym = DOV_fmt_(d,'yyyy-MM');
    labels.push(DOV_fmt_(d,'MMM yy'));
    data.push(DOV_sumPay_(pay, ym));
  }
  return { labels:labels, data:data };
}
function DOV_revByLevel_(pay, siswaById, ym){
  var m = { SD:0, SMP:0, SMA:0, Lainnya:0 };
  pay.forEach(function(p){
    if (DOV_fmt_(DOV_date_(p.tgl),'yyyy-MM') !== ym) return;
    var s = siswaById[p.id]; var j = DOV_jenjang_(s ? s.kelas : '');
    m[j] = (m[j]||0) + (Number(p.nominal)||0);
  });
  return ['SD','SMP','SMA'].map(function(k){ return { level:k, amount:m[k]||0 }; });
}
function DOV_newStudentsProxy_(pay, siswaById, ym){
  // Estimasi: siswa yg pembayaran PERTAMA-nya jatuh di bulan ini
  var first = {};
  pay.forEach(function(p){
    var d = DOV_fmt_(DOV_date_(p.tgl),'yyyy-MM'); if (!d) return;
    if (!first[p.id] || d < first[p.id]) first[p.id] = d;
  });
  var c = 0; Object.keys(first).forEach(function(id){ if (first[id]===ym && siswaById[id]) c++; });
  return c;
}
function DOV_studentsAtRisk_(aktif, absen, paidSet, ym){
  // hitung rate kehadiran per siswa bulan ini
  var stat = {};
  absen.forEach(function(a){
    if (DOV_fmt_(DOV_date_(a.tgl),'yyyy-MM') !== ym) return;
    var id = a.idSiswa; if (!id) return;
    stat[id] = stat[id] || { hadir:0, total:0 };
    stat[id].total++; if ((a.status||'').toLowerCase()==='hadir') stat[id].hadir++;
  });
  var out = [];
  aktif.forEach(function(s){
    var reasons = [];
    var st = stat[s.id];
    if (st && st.total>=3 && (st.hadir/st.total)<0.6) reasons.push('Kehadiran rendah');
    if (!paidSet[s.id]) reasons.push('Belum bayar');
    if (!st || st.total===0) reasons.push('Tidak ada aktivitas');
    if (reasons.length >= 2) out.push({ id:s.id, nama:s.nama, kelas:s.kelas, reasons:reasons });
  });
  return out.slice(0, 15);
}
function DOV_consecutiveAbsence_(absen, threshold){
  // kelompokkan per siswa, urut tanggal, cari streak alpa terakhir
  var byStudent = {};
  absen.forEach(function(a){
    if (!a.idSiswa) return;
    byStudent[a.idSiswa] = byStudent[a.idSiswa] || { nama:a.namaSiswa, rows:[] };
    byStudent[a.idSiswa].rows.push({ d: DOV_fmt_(DOV_date_(a.tgl),'yyyy-MM-dd'), s:(a.status||'').toLowerCase() });
  });
  var out = [];
  Object.keys(byStudent).forEach(function(id){
    var rows = byStudent[id].rows.sort(function(a,b){ return a.d.localeCompare(b.d); });
    var streak = 0;
    for (var i=rows.length-1;i>=0;i--){
      if (rows[i].s==='alpa' || rows[i].s==='alpha') streak++; else break;
    }
    if (streak >= threshold) out.push({ id:id, nama:byStudent[id].nama, streak:streak });
  });
  return out.sort(function(a,b){ return b.streak-a.streak; }).slice(0,10);
}
function DOV_scheduleConflicts_(sched, todayS){
  var msgs = [], seen = {};
  var today = sched.filter(function(x){ return DOV_fmt_(DOV_date_(x.tgl),'yyyy-MM-dd')===todayS; });
  for (var i=0;i<today.length;i++){
    for (var j=i+1;j<today.length;j++){
      var a=today[i], b=today[j];
      var overlap = DOV_timeOverlap_(a.jamMulai,a.jamBerakhir,b.jamMulai,b.jamBerakhir);
      if (!overlap) continue;
      if (a.idGuru && a.idGuru===b.idGuru){ var k='g'+a.idGuru+a.jamMulai; if(!seen[k]){seen[k]=1; msgs.push('Bentrok guru '+(a.namaGuru||a.idGuru)+' jam '+a.jamMulai);} }
      if (a.ruangan && a.ruangan===b.ruangan){ var k2='r'+a.ruangan+a.jamMulai; if(!seen[k2]){seen[k2]=1; msgs.push('Bentrok ruang '+a.ruangan+' jam '+a.jamMulai);} }
    }
  }
  return msgs.slice(0,5);
}
function DOV_timeOverlap_(s1,e1,s2,e2){
  function m(t){ if(!t) return null; var p=t.toString().split(':'); return parseInt(p[0],10)*60+parseInt(p[1]||0,10); }
  var a1=m(s1),b1=m(e1),a2=m(s2),b2=m(e2);
  if (a1==null||a2==null) return false;
  if (b1==null) b1=a1+1; if (b2==null) b2=a2+1;
  return a1 < b2 && a2 < b1;
}
function DOV_missingAttendance_(sched, absen, now){
  var recorded = {};
  absen.forEach(function(a){ recorded[DOV_fmt_(DOV_date_(a.tgl),'yyyy-MM-dd')+'|'+a.rombel+'|'+a.mapel]=true; });
  var count = 0;
  sched.forEach(function(x){
    var d = DOV_date_(x.tgl); if (!(d<now)) return; // hanya yg sudah lewat
    var diff = (now - d)/86400000; if (diff>30) return; // batasi 30 hari ke belakang
    var key = DOV_fmt_(d,'yyyy-MM-dd')+'|'+x.kelas+'|'+x.mapel;
    if (!recorded[key]) count++;
  });
  return count;
}