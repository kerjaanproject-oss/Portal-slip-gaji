// ==========================================================================
// MODUL SURAT PERINTAH LEMBUR (SPL) & PERSETUJUAN BERJENJANG (2 TINGKAT)
// File: Kode_SPL.js
// ==========================================================================

/**
 * Setup / Buat Sheet PengajuanSPL di Google Sheets
 * Dapat dijalankan langsung dari menu dropdown editor Google Apps Script
 */
function setupDatabaseSPL() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Spreadsheet aktif tidak ditemukan. Pastikan Google Apps Script terhubung dengan file Google Sheets.');
  }
  const sheet = getOrCreatePengajuanSPLSheet(ss);
  SpreadsheetApp.flush();
  Logger.log('Berhasil: Sheet ' + sheet.getName() + ' telah siap.');
  return 'Setup Database SPL Berhasil! Sheet ' + sheet.getName() + ' telah siap digunakan.';
}

function getOrCreatePengajuanSPLSheet(ss) {
  if (!ss) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  if (!ss) {
    throw new Error('Spreadsheet aktif tidak ditemukan. Pastikan Google Apps Script ini terhubung dengan Spreadsheet Google Sheets Anda.');
  }

  let sheet = ss.getSheetByName('PengajuanSPL') ||
    ss.getSheetByName('Pengajuan_SPL') ||
    ss.getSheetByName('DataSPL') ||
    ss.getSheetByName('Data_SPL') ||
    ss.getSheetByName('SPL');

  const REQUIRED_HEADERS = [
    'ID SPL',                 // 0
    'Tanggal Pengajuan',      // 1
    'ID Karyawan',            // 2
    'Nama Karyawan',          // 3
    'Departemen',             // 4
    'Jabatan',                // 5
    'Tanggal Lembur',         // 6
    'Shift',                  // 7
    'Jam Mulai',              // 8
    'Jam Selesai',            // 9
    'Total Jam Lembur',       // 10
    'Uraian Pekerjaan',       // 11
    'Lokasi Kerja',           // 12
    'Status',                 // 13
    'Approver 1 Nama',        // 14 (Atasan Langsung)
    'Approver 1 Jabatan',     // 15
    'Approver 1 Tanggal',     // 16
    'Approver 1 Status',      // 17
    'Approver 1 Catatan',     // 18
    'Approver 2 Nama',        // 19 (HRD / Management)
    'Approver 2 Jabatan',     // 20
    'Approver 2 Tanggal',     // 21
    'Approver 2 Status',      // 22
    'Approver 2 Catatan',     // 23
    'Alasan Penolakan'        // 24
  ];

  if (!sheet) {
    sheet = ss.insertSheet('PengajuanSPL');
    sheet.appendRow(REQUIRED_HEADERS);
    sheet.getRange(1, 1, 1, REQUIRED_HEADERS.length).setFontWeight('bold').setBackground('#e2e8f0');
  } else {
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
        return String(h || '').trim().toLowerCase();
      });
      const missingHeaders = [];
      REQUIRED_HEADERS.forEach(function (reqH) {
        const found = existingHeaders.some(function (eh) {
          return eh.toLowerCase() === reqH.toLowerCase();
        });
        if (!found) missingHeaders.push(reqH);
      });
      if (missingHeaders.length > 0) {
        const startCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]).setFontWeight('bold').setBackground('#e2e8f0');
        SpreadsheetApp.flush();
      }
    }
  }
  return sheet;
}

/**
 * READ API: Mengambil daftar pengajuan SPL sesuai otorisasi role
 */
function getPengajuanSPLList(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splSheet = getOrCreatePengajuanSPLSheet(ss);
  const data = splSheet.getDataRange().getValues();

  const user = session.user;
  const currentUsername = String(user.username || '').trim().toLowerCase();
  const userRole = String(user.role || '').trim().toLowerCase();
  const userDept = String(user.dept || '').trim().toLowerCase();
  const userJab = String(user.jabatan || '').trim().toLowerCase();

  const isAdmin = (userRole === 'admin' || userRole === 'admin master' || currentUsername === 'admin');
  const isHRD = (isAdmin || userRole === 'hrd' || (typeof cleanDept === 'function' ? cleanDept(userDept) === 'hrd' : userDept.includes('hr')));
  const isSupervisor = (typeof isJobAtasan === 'function') ? isJobAtasan(userRole, userJab) : (userRole === 'supervisor' || userRole === 'atasan');

  if (!data || data.length <= 1) {
    return {
      success: true,
      list: [],
      pendingCount: 0,
      userRole: user.role,
      isHRD: isHRD,
      isSupervisor: isSupervisor
    };
  }

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id spl') || h === 'id'; });
  const tglPengajuanIdx = headers.findIndex(function (h) { return h.includes('tanggal pengajuan') || h.includes('tgl pengajuan'); });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h === 'username'; });
  const namaIdx = headers.findIndex(function (h) { return h.includes('nama karyawan') || h === 'nama'; });
  const deptIdx = headers.findIndex(function (h) { return h.includes('departemen') || h.includes('dept'); });
  const jabIdx = headers.findIndex(function (h) { return h.includes('jabatan'); });
  const tglLemburIdx = headers.findIndex(function (h) { return h.includes('tanggal lembur') || h.includes('tgl lembur'); });
  const shiftIdx = headers.findIndex(function (h) { return h === 'shift' || h.includes('shift'); });
  const jamMulaiIdx = headers.findIndex(function (h) { return h.includes('jam mulai') || h.includes('jam awal'); });
  const jamSelesaiIdx = headers.findIndex(function (h) { return h.includes('jam selesai') || h.includes('jam akhir'); });
  const totalJamIdx = headers.findIndex(function (h) { return h.includes('total jam'); });
  const uraianIdx = headers.findIndex(function (h) { return h.includes('uraian') || h.includes('pekerjaan') || h.includes('alasan'); });
  const lokasiIdx = headers.findIndex(function (h) { return h.includes('lokasi'); });
  const statusIdx = headers.findIndex(function (h) { return h.includes('status'); });

  const app1NameIdx = headers.findIndex(function (h) { return h.includes('approver 1 nama'); });
  const app1JabIdx = headers.findIndex(function (h) { return h.includes('approver 1 jabatan'); });
  const app1TglIdx = headers.findIndex(function (h) { return h.includes('approver 1 tanggal'); });
  const app1StatusIdx = headers.findIndex(function (h) { return h.includes('approver 1 status'); });
  const app1CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 1 catatan'); });

  const app2NameIdx = headers.findIndex(function (h) { return h.includes('approver 2 nama'); });
  const app2JabIdx = headers.findIndex(function (h) { return h.includes('approver 2 jabatan'); });
  const app2TglIdx = headers.findIndex(function (h) { return h.includes('approver 2 tanggal'); });
  const app2StatusIdx = headers.findIndex(function (h) { return h.includes('approver 2 status'); });
  const app2CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 2 catatan'); });

  const tolakAlasanIdx = headers.findIndex(function (h) { return h.includes('alasan penolakan') || h.includes('alasan tolak'); });

  const list = [];
  let pendingCount = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowId = idIdx !== -1 ? String(row[idIdx] || '').trim() : String(row[0] || '').trim();
    if (!rowId) continue;

    const rowEmpId = empIdIdx !== -1 ? String(row[empIdIdx] || '').trim() : String(row[2] || '').trim();
    const rowEmpLower = rowEmpId.toLowerCase();
    const rowDept = deptIdx !== -1 ? String(row[deptIdx] || '').trim() : String(row[4] || '').trim();
    const rowStatus = statusIdx !== -1 ? String(row[statusIdx] || '').trim() : 'Menunggu Persetujuan Atasan';

    // Filter hak akses:
    // - Admin / HRD: melihat semua SPL
    // - Supervisor/Atasan: melihat SPL milik departemennya ATAU SPL miliknya sendiri
    // - Karyawan biasa: HANYA melihat SPL miliknya sendiri
    let canView = false;
    if (isHRD) {
      canView = true;
    } else if (isSupervisor) {
      const matchDept = (typeof isSameDept === 'function') ? isSameDept(userDept, rowDept) : (userDept === rowDept.toLowerCase());
      if (matchDept || rowEmpLower === currentUsername) {
        canView = true;
      }
    } else {
      if (rowEmpLower === currentUsername) {
        canView = true;
      }
    }

    if (!canView) continue;

    // Hitung pending count untuk badge
    // - Untuk Atasan: pending jika status === 'Menunggu Persetujuan Atasan' & bukan diri sendiri
    // - Untuk HRD: pending jika status === 'Menunggu Persetujuan HRD' (tahap 2) atau 'Menunggu Persetujuan Atasan'
    if (isHRD) {
      if (rowStatus === 'Menunggu Persetujuan HRD' || rowStatus === 'Menunggu Persetujuan Atasan') {
        pendingCount++;
      }
    } else if (isSupervisor) {
      if (rowStatus === 'Menunggu Persetujuan Atasan' && rowEmpLower !== currentUsername) {
        pendingCount++;
      }
    }

    const tglPengajuanVal = tglPengajuanIdx !== -1 ? row[tglPengajuanIdx] : row[1];
    const tglLemburVal = tglLemburIdx !== -1 ? row[tglLemburIdx] : row[6];

    const formattedTglPengajuan = (tglPengajuanVal instanceof Date)
      ? Utilities.formatDate(tglPengajuanVal, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm')
      : String(tglPengajuanVal || '').trim();

    const formattedTglLembur = (tglLemburVal instanceof Date)
      ? Utilities.formatDate(tglLemburVal, 'Asia/Jakarta', 'yyyy-MM-dd')
      : String(tglLemburVal || '').trim();

    function formatTimeHHMM(timeVal) {
      if (!timeVal && timeVal !== 0) return '';
      if (timeVal instanceof Date) {
        return Utilities.formatDate(timeVal, 'Asia/Jakarta', 'HH:mm');
      }
      const s = String(timeVal || '').trim();
      const match = s.match(/(?:^|\s|T)(\d{1,2}):(\d{2})/);
      if (match) {
        return match[1].padStart(2, '0') + ':' + match[2];
      }
      return s;
    }

    const rawJamMulai = jamMulaiIdx !== -1 ? row[jamMulaiIdx] : row[7];
    const rawJamSelesai = jamSelesaiIdx !== -1 ? row[jamSelesaiIdx] : row[8];
    const formattedJamMulai = formatTimeHHMM(rawJamMulai);
    const formattedJamSelesai = formatTimeHHMM(rawJamSelesai);

    list.push({
      id: rowId,
      tanggalPengajuan: formattedTglPengajuan,
      empId: rowEmpId,
      namaKaryawan: namaIdx !== -1 ? String(row[namaIdx] || '').trim() : String(row[3] || '').trim(),
      dept: rowDept,
      jabatan: jabIdx !== -1 ? String(row[jabIdx] || '').trim() : String(row[5] || '').trim(),
      tanggalLembur: formattedTglLembur,
      shift: shiftIdx !== -1 ? String(row[shiftIdx] || 'Shift Pagi').trim() : 'Shift Pagi',
      jamMulai: formattedJamMulai,
      jamSelesai: formattedJamSelesai,
      totalJam: totalJamIdx !== -1 ? Number(row[totalJamIdx] || 0) : Number(row[9] || 0),
      uraianPekerjaan: uraianIdx !== -1 ? String(row[uraianIdx] || '').trim() : String(row[10] || '').trim(),
      lokasiKerja: lokasiIdx !== -1 ? String(row[lokasiIdx] || '').trim() : '',
      status: rowStatus,
      approver1: {
        nama: app1NameIdx !== -1 ? String(row[app1NameIdx] || '').trim() : '',
        jabatan: app1JabIdx !== -1 ? String(row[app1JabIdx] || '').trim() : '',
        tanggal: app1TglIdx !== -1 ? (row[app1TglIdx] instanceof Date ? Utilities.formatDate(row[app1TglIdx], 'Asia/Jakarta', 'yyyy-MM-dd') : String(row[app1TglIdx] || '').trim()) : '',
        status: app1StatusIdx !== -1 ? String(row[app1StatusIdx] || '').trim() : '',
        catatan: app1CatatanIdx !== -1 ? String(row[app1CatatanIdx] || '').trim() : ''
      },
      approver2: {
        nama: app2NameIdx !== -1 ? String(row[app2NameIdx] || '').trim() : '',
        jabatan: app2JabIdx !== -1 ? String(row[app2JabIdx] || '').trim() : '',
        tanggal: app2TglIdx !== -1 ? (row[app2TglIdx] instanceof Date ? Utilities.formatDate(row[app2TglIdx], 'Asia/Jakarta', 'yyyy-MM-dd') : String(row[app2TglIdx] || '').trim()) : '',
        status: app2StatusIdx !== -1 ? String(row[app2StatusIdx] || '').trim() : '',
        catatan: app2CatatanIdx !== -1 ? String(row[app2CatatanIdx] || '').trim() : ''
      },
      alasanPenolakan: tolakAlasanIdx !== -1 ? String(row[tolakAlasanIdx] || '').trim() : ''
    });
  }

  // Urutkan dari yang terbaru
  list.sort(function (a, b) {
    return String(b.tanggalPengajuan || b.id).localeCompare(String(a.tanggalPengajuan || a.id));
  });

  return {
    success: true,
    list: list,
    pendingCount: pendingCount,
    userRole: user.role,
    isHRD: isHRD,
    isSupervisor: isSupervisor
  };
}

/**
 * WRITE API: Simpan Pengajuan SPL Baru
 */
function simpanPengajuanSPL(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  payload = payload || {};
  const tanggalLembur = String(payload.tanggalLembur || '').trim();
  const shift = String(payload.shift || 'Shift Pagi').trim();
  const jamMulai = String(payload.jamMulai || '').trim();
  const jamSelesai = String(payload.jamSelesai || '').trim();
  const totalJam = Number(payload.totalJam || 0);
  const uraianPekerjaan = String(payload.uraianPekerjaan || '').trim();
  const lokasiKerja = String(payload.lokasiKerja || '').trim();

  if (!tanggalLembur) return { success: false, message: 'Tanggal pelaksanaan lembur wajib diisi.' };
  if (!jamMulai || !jamSelesai) return { success: false, message: 'Jam mulai dan jam selesai lembur wajib diisi.' };
  if (totalJam <= 0) return { success: false, message: 'Total jam lembur harus lebih besar dari 0 jam.' };
  if (!lokasiKerja) return { success: false, message: 'Lokasi kerja / jobsite wajib diisi.' };
  if (!uraianPekerjaan) return { success: false, message: 'Uraian pekerjaan lembur wajib diisi.' };

  const user = session.user;
  const empId = String(user.username || '').trim();
  const namaKaryawan = String(user.fullname || user.username || '').trim();
  const dept = String(user.dept || '-').trim();
  const jabatan = String(user.jabatan || 'Karyawan').trim();
  const userRole = String(user.role || '').trim().toLowerCase();

  const isSubmitterHRD = (userRole === 'admin' || userRole === 'hrd');
  const isSubmitterAtasan = (typeof isJobAtasan === 'function')
    ? isJobAtasan(userRole, jabatan)
    : (userRole === 'supervisor' || userRole === 'atasan');

  // Status awal pengajuan:
  // - Jika diajukan oleh Atasan: langsung ke tahap 'Menunggu Persetujuan HRD'
  // - Jika diajukan oleh HRD/Admin: bisa langsung 'Disetujui' atau 'Menunggu Persetujuan HRD'
  let initialStatus = 'Menunggu Persetujuan Atasan';
  if (isSubmitterHRD) {
    initialStatus = 'Disetujui';
  } else if (isSubmitterAtasan) {
    initialStatus = 'Menunggu Persetujuan HRD';
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splSheet = getOrCreatePengajuanSPLSheet(ss);

  // Generate ID SPL unik: SPL-YYYYMM-XXXX
  const now = new Date();
  const ymStr = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMM');
  const formattedNow = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  const data = splSheet.getDataRange().getValues();
  let maxSeq = 0;
  for (let i = 1; i < data.length; i++) {
    const curId = String(data[i][0] || '').trim();
    if (curId.indexOf('SPL-' + ymStr + '-') === 0) {
      const parts = curId.split('-');
      const num = parseInt(parts[2], 10);
      if (!isNaN(num) && num > maxSeq) maxSeq = num;
    }
  }
  const nextSeq = String(maxSeq + 1).padStart(4, '0');
  const splId = 'SPL-' + ymStr + '-' + nextSeq;

  let app1Nama = '', app1Jab = '', app1Tgl = '', app1Stat = '', app1Cat = '';
  let app2Nama = '', app2Jab = '', app2Tgl = '', app2Stat = '', app2Cat = '';

  if (isSubmitterHRD) {
    app1Nama = namaKaryawan;
    app1Jab = jabatan;
    app1Tgl = Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
    app1Stat = 'Disetujui';
    app1Cat = 'Auto-Approve HRD / Admin';

    app2Nama = namaKaryawan;
    app2Jab = jabatan;
    app2Tgl = app1Tgl;
    app2Stat = 'Disetujui';
    app2Cat = 'Auto-Approve HRD / Admin';
  } else if (isSubmitterAtasan) {
    app1Nama = namaKaryawan;
    app1Jab = jabatan;
    app1Tgl = Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
    app1Stat = 'Disetujui';
    app1Cat = 'Diajukan oleh Atasan Departemen';
  }

  const curHeaders = splSheet.getRange(1, 1, 1, splSheet.getLastColumn()).getValues()[0].map(function (h) {
    return String(h || '').trim().toLowerCase();
  });

  const rowData = new Array(curHeaders.length).fill('');
  function setCol(pred, val) {
    const idx = curHeaders.findIndex(pred);
    if (idx !== -1) rowData[idx] = val;
  }

  setCol(function (h) { return h.includes('id spl') || h === 'id'; }, splId);
  setCol(function (h) { return h.includes('tanggal pengajuan') || h.includes('tgl pengajuan'); }, formattedNow);
  setCol(function (h) { return h.includes('id karyawan') || h.includes('nik') || h === 'username'; }, empId);
  setCol(function (h) { return h.includes('nama karyawan') || h === 'nama'; }, namaKaryawan);
  setCol(function (h) { return h.includes('departemen') || h.includes('dept'); }, dept);
  setCol(function (h) { return h.includes('jabatan'); }, jabatan);
  setCol(function (h) { return h.includes('tanggal lembur') || h.includes('tgl lembur'); }, tanggalLembur);
  setCol(function (h) { return h === 'shift' || h.includes('shift'); }, shift);
  setCol(function (h) { return h.includes('jam mulai') || h.includes('jam awal'); }, jamMulai);
  setCol(function (h) { return h.includes('jam selesai') || h.includes('jam akhir'); }, jamSelesai);
  setCol(function (h) { return h.includes('total jam'); }, totalJam);
  setCol(function (h) { return h.includes('uraian') || h.includes('pekerjaan') || h.includes('alasan'); }, uraianPekerjaan);
  setCol(function (h) { return h.includes('lokasi'); }, lokasiKerja);
  setCol(function (h) { return h.includes('status'); }, initialStatus);
  setCol(function (h) { return h.includes('approver 1 nama'); }, app1Nama);
  setCol(function (h) { return h.includes('approver 1 jabatan'); }, app1Jab);
  setCol(function (h) { return h.includes('approver 1 tanggal'); }, app1Tgl);
  setCol(function (h) { return h.includes('approver 1 status'); }, app1Stat);
  setCol(function (h) { return h.includes('approver 1 catatan'); }, app1Cat);
  setCol(function (h) { return h.includes('approver 2 nama'); }, app2Nama);
  setCol(function (h) { return h.includes('approver 2 jabatan'); }, app2Jab);
  setCol(function (h) { return h.includes('approver 2 tanggal'); }, app2Tgl);
  setCol(function (h) { return h.includes('approver 2 status'); }, app2Stat);
  setCol(function (h) { return h.includes('approver 2 catatan'); }, app2Cat);
  setCol(function (h) { return h.includes('alasan penolakan') || h.includes('alasan tolak'); }, '');

  splSheet.appendRow(rowData);
  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Surat Perintah Lembur (SPL) berhasil dibuat dengan ID: ' + splId,
    splId: splId
  };
}

/**
 * WRITE API: Proses Persetujuan Berjenjang (Atasan Tahap 1 -> HRD Tahap 2) atau Penolakan
 */
function processApprovalSPL(token, splId, action, catatan) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  splId = String(splId || '').trim();
  action = String(action || '').trim(); // 'Setujui' or 'Tolak'
  catatan = String(catatan || '').trim();

  if (!splId) return { success: false, message: 'ID SPL tidak valid.' };
  if (action !== 'Setujui' && action !== 'Tolak') {
    return { success: false, message: 'Aksi approval tidak valid (harus Setujui atau Tolak).' };
  }
  if (action === 'Tolak' && !catatan) {
    return { success: false, message: 'Alasan penolakan wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splSheet = getOrCreatePengajuanSPLSheet(ss);
  const data = splSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { success: false, message: 'Data SPL tidak ditemukan.' };

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id spl') || h === 'id'; });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h === 'username'; });
  const deptIdx = headers.findIndex(function (h) { return h.includes('departemen') || h.includes('dept'); });
  const statusIdx = headers.findIndex(function (h) { return h.includes('status'); });

  const app1NameIdx = headers.findIndex(function (h) { return h.includes('approver 1 nama'); });
  const app1JabIdx = headers.findIndex(function (h) { return h.includes('approver 1 jabatan'); });
  const app1TglIdx = headers.findIndex(function (h) { return h.includes('approver 1 tanggal'); });
  const app1StatusIdx = headers.findIndex(function (h) { return h.includes('approver 1 status'); });
  const app1CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 1 catatan'); });

  const app2NameIdx = headers.findIndex(function (h) { return h.includes('approver 2 nama'); });
  const app2JabIdx = headers.findIndex(function (h) { return h.includes('approver 2 jabatan'); });
  const app2TglIdx = headers.findIndex(function (h) { return h.includes('approver 2 tanggal'); });
  const app2StatusIdx = headers.findIndex(function (h) { return h.includes('approver 2 status'); });
  const app2CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 2 catatan'); });

  const tolakAlasanIdx = headers.findIndex(function (h) { return h.includes('alasan penolakan') || h.includes('alasan tolak'); });

  const user = session.user;
  const currentUsername = String(user.username || '').trim().toLowerCase();
  const userRole = String(user.role || '').trim().toLowerCase();
  const userDept = String(user.dept || '').trim().toLowerCase();
  const userJab = String(user.jabatan || '').trim().toLowerCase();

  const isAdmin = (userRole === 'admin' || userRole === 'admin master' || currentUsername === 'admin');
  const isHRD = (isAdmin || userRole === 'hrd' || (typeof cleanDept === 'function' ? cleanDept(userDept) === 'hrd' : userDept.includes('hr')));
  const isSupervisor = (typeof isJobAtasan === 'function') ? isJobAtasan(userRole, userJab) : (userRole === 'supervisor' || userRole === 'atasan');

  if (!isHRD && !isSupervisor) {
    return { success: false, message: 'Akses ditolak: Anda tidak memiliki wewenang untuk melakukan approval SPL.' };
  }

  for (let i = 1; i < data.length; i++) {
    const rowId = idIdx !== -1 ? String(data[i][idIdx] || '').trim() : String(data[i][0] || '').trim();
    if (rowId === splId) {
      const rowEmpId = empIdIdx !== -1 ? String(data[i][empIdIdx] || '').trim().toLowerCase() : '';
      const rowDept = deptIdx !== -1 ? String(data[i][deptIdx] || '').trim() : '';
      const currentStatus = statusIdx !== -1 ? String(data[i][statusIdx] || '').trim() : 'Menunggu Persetujuan Atasan';

      // Cegah approval diri sendiri kecuali jika Admin
      if (rowEmpId === currentUsername && !isAdmin) {
        return { success: false, message: 'Anda tidak dapat menyetujui pengajuan SPL milik Anda sendiri.' };
      }

      const now = new Date();
      const nowStr = Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
      const rowIndex = i + 1;

      // === TAHAP 1: Approval Atasan ===
      if (currentStatus === 'Menunggu Persetujuan Atasan') {
        if (!isSupervisor && !isHRD) {
          return { success: false, message: 'Hanya Atasan/Supervisor atau HRD yang berhak memproses permohonan tahap ini.' };
        }

        // Atasan hanya bisa approve departemen yang sama (HRD/Admin bebas)
        if (!isHRD) {
          const matchDept = (typeof isSameDept === 'function') ? isSameDept(userDept, rowDept) : (userDept === rowDept.toLowerCase());
          if (!matchDept) {
            return {
              success: false,
              message: 'Akses ditolak: Anda hanya dapat memproses SPL karyawan di departemen Anda (' + (user.dept || '-') + '). Departemen pemohon adalah ' + (rowDept || '-') + '.'
            };
          }
        }

        if (action === 'Setujui') {
          if (app1NameIdx !== -1) splSheet.getRange(rowIndex, app1NameIdx + 1).setValue(user.fullname || user.username);
          if (app1JabIdx !== -1) splSheet.getRange(rowIndex, app1JabIdx + 1).setValue(user.jabatan || 'Atasan');
          if (app1TglIdx !== -1) splSheet.getRange(rowIndex, app1TglIdx + 1).setValue(nowStr);
          if (app1StatusIdx !== -1) splSheet.getRange(rowIndex, app1StatusIdx + 1).setValue('Disetujui');
          if (app1CatatanIdx !== -1) splSheet.getRange(rowIndex, app1CatatanIdx + 1).setValue(catatan || 'Disetujui Atasan Langsung');

          // Naik ke tahap 2 (HRD)
          if (statusIdx !== -1) splSheet.getRange(rowIndex, statusIdx + 1).setValue('Menunggu Persetujuan HRD');
          SpreadsheetApp.flush();

          return {
            success: true,
            message: 'SPL ' + splId + ' telah disetujui pada Tahap 1 (Atasan) dan diteruskan ke HRD untuk verifikasi akhir.'
          };
        } else {
          // Tolak
          if (app1NameIdx !== -1) splSheet.getRange(rowIndex, app1NameIdx + 1).setValue(user.fullname || user.username);
          if (app1JabIdx !== -1) splSheet.getRange(rowIndex, app1JabIdx + 1).setValue(user.jabatan || 'Atasan');
          if (app1TglIdx !== -1) splSheet.getRange(rowIndex, app1TglIdx + 1).setValue(nowStr);
          if (app1StatusIdx !== -1) splSheet.getRange(rowIndex, app1StatusIdx + 1).setValue('Ditolak');
          if (app1CatatanIdx !== -1) splSheet.getRange(rowIndex, app1CatatanIdx + 1).setValue(catatan);

          if (tolakAlasanIdx !== -1) splSheet.getRange(rowIndex, tolakAlasanIdx + 1).setValue('Ditolak Atasan: ' + catatan);
          if (statusIdx !== -1) splSheet.getRange(rowIndex, statusIdx + 1).setValue('Ditolak');
          SpreadsheetApp.flush();

          return {
            success: true,
            message: 'SPL ' + splId + ' telah ditolak oleh Atasan.'
          };
        }
      }

      // === TAHAP 2: Verifikasi Akhir HRD ===
      if (currentStatus === 'Menunggu Persetujuan HRD') {
        if (!isHRD) {
          return { success: false, message: 'Akses ditolak: Hanya HRD / Management yang berwenang memberikan persetujuan akhir pada tahap ini.' };
        }

        if (action === 'Setujui') {
          if (app2NameIdx !== -1) splSheet.getRange(rowIndex, app2NameIdx + 1).setValue(user.fullname || user.username);
          if (app2JabIdx !== -1) splSheet.getRange(rowIndex, app2JabIdx + 1).setValue(user.jabatan || 'HRD / Management');
          if (app2TglIdx !== -1) splSheet.getRange(rowIndex, app2TglIdx + 1).setValue(nowStr);
          if (app2StatusIdx !== -1) splSheet.getRange(rowIndex, app2StatusIdx + 1).setValue('Disetujui');
          if (app2CatatanIdx !== -1) splSheet.getRange(rowIndex, app2CatatanIdx + 1).setValue(catatan || 'Disetujui HRD & Management');

          // Final Approval
          if (statusIdx !== -1) splSheet.getRange(rowIndex, statusIdx + 1).setValue('Disetujui');
          SpreadsheetApp.flush();

          return {
            success: true,
            message: 'SPL ' + splId + ' berhasil disetujui penuh oleh HRD. Dokumen resmi Surat Perintah Lembur siap dicetak!'
          };
        } else {
          // Tolak
          if (app2NameIdx !== -1) splSheet.getRange(rowIndex, app2NameIdx + 1).setValue(user.fullname || user.username);
          if (app2JabIdx !== -1) splSheet.getRange(rowIndex, app2JabIdx + 1).setValue(user.jabatan || 'HRD / Management');
          if (app2TglIdx !== -1) splSheet.getRange(rowIndex, app2TglIdx + 1).setValue(nowStr);
          if (app2StatusIdx !== -1) splSheet.getRange(rowIndex, app2StatusIdx + 1).setValue('Ditolak');
          if (app2CatatanIdx !== -1) splSheet.getRange(rowIndex, app2CatatanIdx + 1).setValue(catatan);

          if (tolakAlasanIdx !== -1) splSheet.getRange(rowIndex, tolakAlasanIdx + 1).setValue('Ditolak HRD: ' + catatan);
          if (statusIdx !== -1) splSheet.getRange(rowIndex, statusIdx + 1).setValue('Ditolak');
          SpreadsheetApp.flush();

          return {
            success: true,
            message: 'SPL ' + splId + ' telah ditolak oleh HRD.'
          };
        }
      }

      return {
        success: false,
        message: 'Status SPL saat ini adalah "' + currentStatus + '", tidak dapat diproses approval lagi.'
      };
    }
  }

  return { success: false, message: 'SPL tidak ditemukan di database.' };
}

/**
 * WRITE API: Batalkan Pengajuan SPL oleh Karyawan Pemohon
 */
function batalPengajuanSPL(token, splId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  splId = String(splId || '').trim();
  if (!splId) return { success: false, message: 'ID SPL tidak valid.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splSheet = getOrCreatePengajuanSPLSheet(ss);
  const data = splSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { success: false, message: 'Data SPL tidak ditemukan.' };

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id spl') || h === 'id'; });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h === 'username'; });
  const statusIdx = headers.findIndex(function (h) { return h.includes('status'); });

  const currentUsername = String(session.user.username || '').trim().toLowerCase();
  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isAdmin = (userRole === 'admin' || userRole === 'admin master' || currentUsername === 'admin');

  for (let i = 1; i < data.length; i++) {
    const rowId = idIdx !== -1 ? String(data[i][idIdx] || '').trim() : String(data[i][0] || '').trim();
    if (rowId === splId) {
      const rowEmpId = empIdIdx !== -1 ? String(data[i][empIdIdx] || '').trim().toLowerCase() : '';
      const currentStatus = statusIdx !== -1 ? String(data[i][statusIdx] || '').trim() : '';

      if (!isAdmin && rowEmpId !== currentUsername) {
        return { success: false, message: 'Akses ditolak: Anda tidak memiliki wewenang untuk membatalkan pengajuan ini.' };
      }

      if (currentStatus === 'Disetujui') {
        return { success: false, message: 'SPL yang sudah disetujui penuh tidak dapat dibatalkan.' };
      }

      const rowIndex = i + 1;
      if (statusIdx !== -1) splSheet.getRange(rowIndex, statusIdx + 1).setValue('Dibatalkan');
      SpreadsheetApp.flush();

      return { success: true, message: 'Pengajuan SPL ' + splId + ' berhasil dibatalkan.' };
    }
  }

  return { success: false, message: 'Data SPL tidak ditemukan.' };
}

/**
 * WRITE API: Update / Edit Pengajuan SPL
 */
function updatePengajuanSPL(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  if (!payload || !payload.id) {
    return { success: false, message: 'ID SPL tidak valid.' };
  }

  const splId = String(payload.id).trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splSheet = getOrCreatePengajuanSPLSheet(ss);
  const data = splSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { success: false, message: 'Data SPL tidak ditemukan.' };

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id spl') || h === 'id'; });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h === 'username'; });
  const statusIdx = headers.findIndex(function (h) { return h.includes('status'); });
  const tglLemburIdx = headers.findIndex(function (h) { return h.includes('tanggal lembur') || h.includes('tgl lembur'); });
  const shiftIdx = headers.findIndex(function (h) { return h === 'shift' || h.includes('shift'); });
  const jamMulaiIdx = headers.findIndex(function (h) { return h.includes('jam mulai') || h.includes('jam awal'); });
  const jamSelesaiIdx = headers.findIndex(function (h) { return h.includes('jam selesai') || h.includes('jam akhir'); });
  const totalJamIdx = headers.findIndex(function (h) { return h.includes('total jam'); });
  const uraianIdx = headers.findIndex(function (h) { return h.includes('uraian') || h.includes('pekerjaan') || h.includes('alasan'); });
  const lokasiIdx = headers.findIndex(function (h) { return h.includes('lokasi'); });

  const currentUsername = String(session.user.username || '').trim().toLowerCase();
  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isAdmin = (userRole === 'admin' || userRole === 'admin master' || currentUsername === 'admin');

  for (let i = 1; i < data.length; i++) {
    const rowId = idIdx !== -1 ? String(data[i][idIdx] || '').trim() : String(data[i][0] || '').trim();
    if (rowId === splId) {
      const rowEmpId = empIdIdx !== -1 ? String(data[i][empIdIdx] || '').trim().toLowerCase() : '';
      const currentStatus = statusIdx !== -1 ? String(data[i][statusIdx] || '').trim() : '';

      if (!isAdmin && rowEmpId !== currentUsername) {
        return { success: false, message: 'Akses ditolak: Anda tidak memiliki wewenang untuk mengubah pengajuan SPL ini.' };
      }

      if (currentStatus === 'Disetujui') {
        return { success: false, message: 'SPL yang sudah disetujui penuh tidak dapat diedit lagi.' };
      }

      const rowIndex = i + 1;
      if (tglLemburIdx !== -1 && payload.tanggalLembur) splSheet.getRange(rowIndex, tglLemburIdx + 1).setValue(payload.tanggalLembur);
      if (shiftIdx !== -1 && payload.shift) splSheet.getRange(rowIndex, shiftIdx + 1).setValue(payload.shift);
      if (jamMulaiIdx !== -1 && payload.jamMulai) splSheet.getRange(rowIndex, jamMulaiIdx + 1).setValue(payload.jamMulai);
      if (jamSelesaiIdx !== -1 && payload.jamSelesai) splSheet.getRange(rowIndex, jamSelesaiIdx + 1).setValue(payload.jamSelesai);
      if (totalJamIdx !== -1 && payload.totalJam !== undefined) splSheet.getRange(rowIndex, totalJamIdx + 1).setValue(payload.totalJam);
      if (uraianIdx !== -1 && payload.uraianPekerjaan) splSheet.getRange(rowIndex, uraianIdx + 1).setValue(payload.uraianPekerjaan);
      if (lokasiIdx !== -1 && payload.lokasiKerja) splSheet.getRange(rowIndex, lokasiIdx + 1).setValue(payload.lokasiKerja);

      SpreadsheetApp.flush();

      return { success: true, message: 'Pengajuan SPL ' + splId + ' berhasil diperbarui.' };
    }
  }

  return { success: false, message: 'Data SPL tidak ditemukan.' };
}

/**
 * WRITE API: Hapus SPL dari Database (Admin / HRD atau Karyawan Pemilik sebelum Disetujui)
 */
function deletePengajuanSPL(token, splId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const currentUsername = String(session.user.username || '').trim().toLowerCase();
  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd' || currentUsername === 'admin');

  splId = String(splId || '').trim();
  if (!splId) return { success: false, message: 'ID SPL tidak valid.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splSheet = getOrCreatePengajuanSPLSheet(ss);
  const data = splSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { success: false, message: 'Data SPL tidak ditemukan.' };

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id spl') || h === 'id'; });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h === 'username'; });
  const statusIdx = headers.findIndex(function (h) { return h.includes('status'); });

  for (let i = 1; i < data.length; i++) {
    const rowId = idIdx !== -1 ? String(data[i][idIdx] || '').trim() : String(data[i][0] || '').trim();
    if (rowId === splId) {
      const rowEmpId = empIdIdx !== -1 ? String(data[i][empIdIdx] || '').trim().toLowerCase() : '';
      const currentStatus = statusIdx !== -1 ? String(data[i][statusIdx] || '').trim() : '';

      if (!isHRD && rowEmpId !== currentUsername) {
        return { success: false, message: 'Akses ditolak: Anda tidak memiliki wewenang untuk menghapus pengajuan ini.' };
      }

      if (currentStatus === 'Disetujui' && !isHRD) {
        return { success: false, message: 'SPL yang sudah disetujui penuh tidak dapat dihapus.' };
      }

      splSheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true, message: 'Data SPL ' + splId + ' berhasil dihapus dari database.' };
    }
  }

  return { success: false, message: 'Data SPL tidak ditemukan.' };
}
