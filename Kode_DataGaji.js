// ==========================================================================
// MODUL INPUT & PENGELOLAAN DATA GAJI KARYAWAN (SISTEM PERIODE + TUTUP BUKU)
// File: Kode_DataGaji.js
// 
// Solusi Database Cerdas:
// 1. Sheet 'PeriodePayroll': Mengelola daftar periode cut-off & status buka/tutup buku.
// 2. Sheet 'DataGaji': HANYA ADA 1 BARIS PER KARYAWAN PER PERIODE
//    (misal 37 karyawan = tetap 37 baris per periode).
// 3. Setiap baris menyimpan akumulasi total berjalan (Overtime, Ritase, Tunjangan, Total Gaji, Hari Kerja).
// 4. Rincian harian per tanggal disimpan dalam format JSON di kolom 'Rincian_Log_Harian'.
// 5. Saat tutup buku, data periode tersebut dikunci permanen (status: DITUTUP).
// 6. Saat buka periode baru, baris baru untuk seluruh karyawan dibuat dengan total Rp 0.
// ==========================================================================

const PERIODE_PAYROLL_HEADERS = [
  'ID Periode',
  'Nama Periode',
  'Tanggal Mulai',
  'Tanggal Selesai',
  'Status',
  'Total Pengeluaran Gaji',
  'Tanggal Ditutup',
  'Ditutup Oleh',
  'Keterangan',
  'Dibuat Pada'
];

const DATA_GAJI_HEADERS = [
  'ID Periode',
  'ID Karyawan',
  'Nama Karyawan',
  'Departemen & Jabatan',
  'Gaji Pokok',
  'Total Jam Overtime',
  'Total Nominal Overtime',
  'Total Ritase',
  'Total Nominal Ritase',
  'Total Tunjangan',
  'Total Gaji Berjalan',
  'Hari Kerja',
  'Rincian_Log_Harian',
  'Terakhir Diupdate'
];

/**
 * Utility: Bersihkan nilai numerik dari string atau format uang
 */
function cleanSalaryNum(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  const s = String(val).replace(/[^0-9.-]/g, '');
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

/**
 * Format header sheet DataGaji
 */
function formatDataGajiHeaders(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, DATA_GAJI_HEADERS.length);
  headerRange
    .setFontWeight('bold')
    .setBackground('#1e293b')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sheet.setRowHeight(1, 35);
  sheet.setFrozenRows(1);
}

/**
 * Format kolom numerik dan mata uang pada sheet DataGaji
 */
function formatDataGajiDataRange(sheet, rowCount) {
  if (!rowCount || rowCount <= 0) return;
  const startRow = 2;
  const lastRow = sheet.getLastRow();
  const count = lastRow >= 2 ? (lastRow - 1) : rowCount;
  if (count <= 0) return;

  // Kolom Rupiah: Gaji Pokok (5), Total Nominal Overtime (7), Total Nominal Ritase (9), Total Tunjangan (10), Total Gaji Berjalan (11)
  sheet.getRange(startRow, 5, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 7, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 9, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 10, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 11, count, 1).setNumberFormat('"Rp "#,##0');

  // Kolom Numerik/Jumlah: Total Jam Overtime (6), Total Ritase (8), Hari Kerja (12)
  sheet.getRange(startRow, 6, count, 1).setNumberFormat('#,##0.##');
  sheet.getRange(startRow, 8, count, 1).setNumberFormat('#,##0');
  sheet.getRange(startRow, 12, count, 1).setNumberFormat('#,##0');
}

/**
 * Inisialisasi atau ambil sheet PeriodePayroll
 */
function getOrCreatePeriodePayrollSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('PeriodePayroll') || ss.getSheetByName('Periode_Payroll') || ss.getSheetByName('Periode Gaji');
  if (!sheet) {
    sheet = ss.insertSheet('PeriodePayroll');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(PERIODE_PAYROLL_HEADERS);
    const headerRange = sheet.getRange(1, 1, 1, PERIODE_PAYROLL_HEADERS.length);
    headerRange
      .setFontWeight('bold')
      .setBackground('#1e293b')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    sheet.setRowHeight(1, 35);
    sheet.setFrozenRows(1);

    // Otomatis buat periode aktif pertama (Bulan berjalan saat ini)
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, '0');
    const BULAN_INDO = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
    const namaPeriode = BULAN_INDO[now.getMonth()] + ' ' + yyyy;
    const periodId = 'PRD-' + yyyy + mm;

    // Rentang cut-off default tgl 21 bulan lalu s/d tgl 20 bulan ini
    const prevMonthDate = new Date(yyyy, now.getMonth() - 1, 21);
    const thisMonthCutDate = new Date(yyyy, now.getMonth(), 20);
    const tglMulai = Utilities.formatDate(prevMonthDate, 'Asia/Jakarta', 'yyyy-MM-dd');
    const tglSelesai = Utilities.formatDate(thisMonthCutDate, 'Asia/Jakarta', 'yyyy-MM-dd');
    const nowStr = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      periodId,
      namaPeriode,
      tglMulai,
      tglSelesai,
      'TERBUKA',
      0,
      '',
      '',
      'Periode Inisial Otomatis',
      nowStr
    ]);
    SpreadsheetApp.flush();
  }

  return sheet;
}

/**
 * Mengambil periode yang sedang aktif (TERBUKA)
 */
function getActivePeriode(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = getOrCreatePeriodePayrollSheet(ss);
  const data = pSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return null;

  // Cari periode berstatus TERBUKA (paling akhir)
  for (let i = data.length - 1; i >= 1; i--) {
    const status = String(data[i][4] || '').trim().toUpperCase();
    if (status === 'TERBUKA') {
      return {
        id: String(data[i][0] || '').trim(),
        name: String(data[i][1] || '').trim(),
        nama: String(data[i][1] || '').trim(),
        startDate: data[i][2] ? (data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], 'Asia/Jakarta', 'yyyy-MM-dd') : String(data[i][2]).trim()) : '',
        endDate: data[i][3] ? (data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], 'Asia/Jakarta', 'yyyy-MM-dd') : String(data[i][3]).trim()) : '',
        status: 'TERBUKA',
        totalGaji: cleanSalaryNum(data[i][5]),
        closedAt: data[i][6] ? String(data[i][6]) : '',
        closedBy: String(data[i][7] || '').trim(),
        note: String(data[i][8] || '').trim(),
        rowIndex: i + 1
      };
    }
  }

  // Jika semua ditutup, kembalikan baris terakhir
  const lastIdx = data.length - 1;
  return {
    id: String(data[lastIdx][0] || '').trim(),
    name: String(data[lastIdx][1] || '').trim(),
    nama: String(data[lastIdx][1] || '').trim(),
    startDate: data[lastIdx][2] ? (data[lastIdx][2] instanceof Date ? Utilities.formatDate(data[lastIdx][2], 'Asia/Jakarta', 'yyyy-MM-dd') : String(data[lastIdx][2]).trim()) : '',
    endDate: data[lastIdx][3] ? (data[lastIdx][3] instanceof Date ? Utilities.formatDate(data[lastIdx][3], 'Asia/Jakarta', 'yyyy-MM-dd') : String(data[lastIdx][3]).trim()) : '',
    status: String(data[lastIdx][4] || '').trim().toUpperCase(),
    totalGaji: cleanSalaryNum(data[lastIdx][5]),
    closedAt: data[lastIdx][6] ? String(data[lastIdx][6]) : '',
    closedBy: String(data[lastIdx][7] || '').trim(),
    note: String(data[lastIdx][8] || '').trim(),
    rowIndex: lastIdx + 1
  };
}

/**
 * Sinkronisasi karyawan dari sheet 'Users' ke 'DataGaji' untuk periode tertentu
 */
function syncEmployeesToDataGajiForPeriod(ss, sheet, periodId) {
  if (!sheet || !periodId) return;
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return;
  const userData = userSheet.getDataRange().getValues();
  if (!userData || userData.length <= 1) return;

  const currentData = sheet.getDataRange().getValues();
  const existingIds = new Set();

  for (let i = 1; i < currentData.length; i++) {
    const rowPeriod = String(currentData[i][0] || '').trim();
    if (rowPeriod === periodId) {
      const id = String(currentData[i][1] || '').trim().toLowerCase();
      if (id) existingIds.add(id);
    }
  }

  const rowsToAppend = [];
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  for (let u = 1; u < userData.length; u++) {
    const uId = String(userData[u][0] || '').trim();
    const uName = String(userData[u][1] || '').trim();
    if (!uId && !uName) continue;

    if (existingIds.has(uId.toLowerCase())) continue;

    const uJabatan = String(userData[u][2] || '').trim();
    const uDept = String(userData[u][3] || '').trim();
    const uBasic = cleanSalaryNum(userData[u][11]);
    const deptJabatan = uDept ? (uDept + (uJabatan ? ' - ' + uJabatan : '')) : uJabatan;

    rowsToAppend.push([
      periodId,
      uId,
      uName,
      deptJabatan,
      uBasic,
      0,
      0,
      0,
      0,
      0,
      uBasic,
      0,
      '[]',
      nowStr
    ]);
  }

  if (rowsToAppend.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, DATA_GAJI_HEADERS.length).setValues(rowsToAppend);
    formatDataGajiDataRange(sheet, rowsToAppend.length);
    SpreadsheetApp.flush();
  }
}

/**
 * Otomatis migrasi data lama (format tanpa ID Periode) menjadi format periode
 */
function migrateOldDataGajiToSummaryFormat(sheet, ss, defaultPeriodId) {
  if (!sheet) return;
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  const allData = sheet.getDataRange().getValues();
  if (!allData || allData.length <= 1) {
    sheet.clear();
    sheet.appendRow(DATA_GAJI_HEADERS);
    formatDataGajiHeaders(sheet);
    syncEmployeesToDataGajiForPeriod(ss, sheet, defaultPeriodId);
    return;
  }

  sheet.clear();
  sheet.appendRow(DATA_GAJI_HEADERS);
  formatDataGajiHeaders(sheet);
  syncEmployeesToDataGajiForPeriod(ss, sheet, defaultPeriodId);
  SpreadsheetApp.flush();
}

/**
 * Inisialisasi atau ambil sheet DataGaji
 */
function getOrCreateDataGajiSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DataGaji') || ss.getSheetByName('Data Gaji') || ss.getSheetByName('Data_Gaji');
  if (!sheet) {
    sheet = ss.insertSheet('DataGaji');
  }

  const activePeriod = getActivePeriode(ss);
  const activePeriodId = activePeriod ? activePeriod.id : 'PRD-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMM');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(DATA_GAJI_HEADERS);
    formatDataGajiHeaders(sheet);
    syncEmployeesToDataGajiForPeriod(ss, sheet, activePeriodId);
  } else {
    // Periksa apakah format header masih belum ada kolom ID Periode
    const firstHeader = String(sheet.getRange(1, 1).getValue() || '').trim().toLowerCase();
    if (firstHeader === 'id karyawan') {
      // Migrasi in-place: sisipkan kolom ID Periode di posisi kolom 1
      sheet.insertColumnBefore(1);
      sheet.getRange(1, 1).setValue('ID Periode');
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const periodValues = [];
        for (let r = 2; r <= lastRow; r++) {
          periodValues.push([activePeriodId]);
        }
        sheet.getRange(2, 1, periodValues.length, 1).setValues(periodValues);
      }
      formatDataGajiHeaders(sheet);
      formatDataGajiDataRange(sheet, lastRow - 1);
    } else if (firstHeader !== 'id periode') {
      migrateOldDataGajiToSummaryFormat(sheet, ss, activePeriodId);
    }
  }

  return sheet;
}

// =========================================================================
// API PERIODE PAYROLL (BUKA BUKU, TUTUP BUKU, DAFTAR PERIODE)
// =========================================================================

/**
 * READ API: Mengambil daftar seluruh periode payroll & status aktifnya
 */
function getPeriodePayrollList(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = getOrCreatePeriodePayrollSheet(ss);
  const data = pSheet.getDataRange().getValues();

  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || !row[0]) continue;
    const pId = String(row[0] || '').trim();
    const pName = String(row[1] || '').trim();
    const tglMulai = row[2] ? (row[2] instanceof Date ? Utilities.formatDate(row[2], 'Asia/Jakarta', 'yyyy-MM-dd') : String(row[2]).trim()) : '';
    const tglSelesai = row[3] ? (row[3] instanceof Date ? Utilities.formatDate(row[3], 'Asia/Jakarta', 'yyyy-MM-dd') : String(row[3]).trim()) : '';
    const status = String(row[4] || 'TERBUKA').trim().toUpperCase();
    const totalGaji = cleanSalaryNum(row[5]);
    const closedAt = row[6] ? String(row[6]) : '';
    const closedBy = String(row[7] || '').trim();
    const note = String(row[8] || '').trim();

    list.push({
      id: pId,
      name: pName,
      nama: pName,
      startDate: tglMulai,
      endDate: tglSelesai,
      status: status,
      totalGaji: totalGaji,
      closedAt: closedAt,
      closedBy: closedBy,
      note: note
    });
  }

  list.reverse(); // Periode terbaru muncul paling atas

  const activePeriod = getActivePeriode(ss);

  return {
    success: true,
    list: list,
    activePeriod: activePeriod
  };
}

/**
 * WRITE API: Buka buku periode payroll baru dengan custom cut-off
 */
function openNewPeriodePayroll(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');
  if (!isHRD) {
    return { success: false, message: 'Akses ditolak: Hanya Admin/HRD yang dapat membuka periode payroll baru.' };
  }

  const pName = String((payload && (payload.name || payload.nama)) || '').trim();
  const tglMulai = String((payload && (payload.startDate || payload.tanggalMulai)) || '').trim();
  const tglSelesai = String((payload && (payload.endDate || payload.tanggalSelesai)) || '').trim();
  const note = String((payload && (payload.note || payload.catatan)) || '').trim();
  const autoClose = !!(payload && (payload.autoClosePrevious || payload.autoClosePrev));

  if (!payload || !pName || !tglMulai || !tglSelesai) {
    return { success: false, message: 'Mohon lengkapi Nama Periode, Tanggal Mulai Cut-Off, dan Tanggal Selesai Cut-Off.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = getOrCreatePeriodePayrollSheet(ss);
  const pData = pSheet.getDataRange().getValues();

  // Buat ID Periode unik
  const cleanNameForId = pName.replace(/[^a-zA-Z0-9]/g, '').slice(0, 10);
  const periodId = 'PRD-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMM') + '-' + cleanNameForId + '-' + Utilities.getUuid().slice(0, 4);

  // Jika ada periode sebelumnya yang masih TERBUKA, otomatis tutup atau beri status ditutup jika diminta
  if (autoClose) {
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    for (let i = 1; i < pData.length; i++) {
      if (String(pData[i][4] || '').trim().toUpperCase() === 'TERBUKA') {
        pSheet.getRange(i + 1, 5).setValue('DITUTUP');
        pSheet.getRange(i + 1, 7).setValue(nowStr);
        pSheet.getRange(i + 1, 8).setValue(session.user.fullname || session.user.username);
      }
    }
  }

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
  pSheet.appendRow([
    periodId,
    pName,
    tglMulai,
    tglSelesai,
    'TERBUKA',
    0,
    '',
    '',
    note,
    nowStr
  ]);

  // Siapkan baris karyawan di sheet DataGaji untuk periode baru ini (angka kembali 0)
  const dSheet = getOrCreateDataGajiSheet(ss);
  syncEmployeesToDataGajiForPeriod(ss, dSheet, periodId);

  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Periode payroll baru "' + pName + '" berhasil dibuka.',
    period: {
      id: periodId,
      name: pName,
      nama: pName,
      startDate: tglMulai,
      endDate: tglSelesai,
      status: 'TERBUKA'
    }
  };
}

/**
 * WRITE API: Tutup buku periode payroll (Kunci data & hitung rekapitulasi final)
 */
function closePeriodePayroll(token, periodId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');
  if (!isHRD) {
    return { success: false, message: 'Akses ditolak: Hanya Admin/HRD yang dapat menutup buku periode payroll.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = getOrCreatePeriodePayrollSheet(ss);
  const pData = pSheet.getDataRange().getValues();

  let foundRow = -1;
  let pName = '';
  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][0] || '').trim() === String(periodId).trim()) {
      foundRow = i + 1;
      pName = String(pData[i][1] || '').trim();
      break;
    }
  }

  if (foundRow === -1) {
    return { success: false, message: 'Periode payroll tidak ditemukan.' };
  }

  // Hitung total pengeluaran gaji periode ini dari sheet DataGaji
  const dSheet = getOrCreateDataGajiSheet(ss);
  const dData = dSheet.getDataRange().getValues();
  let totalPengeluaran = 0;
  let countEmp = 0;

  for (let d = 1; d < dData.length; d++) {
    if (String(dData[d][0] || '').trim() === String(periodId).trim()) {
      totalPengeluaran += cleanSalaryNum(dData[d][10]); // Total Gaji Berjalan
      countEmp++;
    }
  }

  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
  pSheet.getRange(foundRow, 5).setValue('DITUTUP');
  pSheet.getRange(foundRow, 6).setValue(totalPengeluaran);
  pSheet.getRange(foundRow, 7).setValue(nowStr);
  pSheet.getRange(foundRow, 8).setValue(session.user.fullname || session.user.username);

  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Tutup buku periode "' + pName + '" berhasil diselesaikan. Total ' + countEmp + ' karyawan terkunci aman.',
    totalGaji: totalPengeluaran,
    countEmp: countEmp
  };
}

/**
 * WRITE API: Hapus periode payroll yang sudah berstatus DITUTUP beserta data gajinya
 */
function deletePeriodePayroll(token, periodId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');
  if (!isHRD) {
    return { success: false, message: 'Akses ditolak: Hanya Admin/HRD yang dapat menghapus arsip periode payroll.' };
  }

  if (!periodId) {
    return { success: false, message: 'ID Periode tidak valid.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pSheet = getOrCreatePeriodePayrollSheet(ss);
  const pData = pSheet.getDataRange().getValues();

  let targetRow = -1;
  let periodName = '';
  let periodStatus = '';

  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][0] || '').trim() === String(periodId).trim()) {
      targetRow = i + 1;
      periodName = String(pData[i][1] || '').trim();
      periodStatus = String(pData[i][4] || '').trim().toUpperCase();
      break;
    }
  }

  if (targetRow === -1) {
    return { success: false, message: 'Periode payroll tidak ditemukan di database.' };
  }

  if (periodStatus === 'TERBUKA') {
    return { success: false, message: 'Periode yang sedang aktif (BUKU TERBUKA) tidak dapat dihapus. Silakan tutup buku periode terlebih dahulu sebelum menghapusnya.' };
  }

  // 1. Hapus baris di PeriodePayroll
  pSheet.deleteRow(targetRow);

  // 2. Hapus baris-baris terkait di DataGaji (dari bawah ke atas)
  const dSheet = getOrCreateDataGajiSheet(ss);
  const dData = dSheet.getDataRange().getValues();
  let deletedCount = 0;

  for (let d = dData.length - 1; d >= 1; d--) {
    if (String(dData[d][0] || '').trim() === String(periodId).trim()) {
      dSheet.deleteRow(d + 1);
      deletedCount++;
    }
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Periode "' + periodName + '" beserta ' + deletedCount + ' baris data gajinya berhasil dihapus dari database.',
    deletedCount: deletedCount
  };
}

// =========================================================================
// API DATA GAJI KARYAWAN (READ, SAVE, DELETE)
// =========================================================================

/**
 * READ API: Mengambil data gaji karyawan berdasarkan filter periode
 */
function getDataGaji(token, filterDept, filterTanggal, filterPeriodeId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');

  const currentUsername = String(session.user.username || '').trim().toLowerCase();
  const currentFullname = String(session.user.fullname || '').trim().toLowerCase();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();

  // Dapatkan periode yang ditargetkan
  const pSheet = getOrCreatePeriodePayrollSheet(ss);
  const allPeriodsRes = getPeriodePayrollList(token);
  const allPeriods = allPeriodsRes.success ? allPeriodsRes.list : [];
  const activePeriod = allPeriodsRes.success ? allPeriodsRes.activePeriod : null;

  let targetPeriodId = filterPeriodeId && filterPeriodeId !== 'SEMUA' && filterPeriodeId !== 'AKTIF'
    ? String(filterPeriodeId).trim()
    : (activePeriod ? activePeriod.id : '');

  // Cari info periode yang ditargetkan
  let targetPeriodInfo = allPeriods.find(function (p) { return p.id === targetPeriodId; }) || activePeriod;
  if (!targetPeriodId && targetPeriodInfo) {
    targetPeriodId = targetPeriodInfo.id;
  }

  const sheet = getOrCreateDataGajiSheet(ss);
  const data = sheet.getDataRange().getValues();

  if (!data || data.length <= 1) {
    return {
      success: true,
      list: [],
      employees: [],
      isHRD: isHRD,
      period: targetPeriodInfo,
      periods: allPeriods
    };
  }

  const flatList = [];
  const employeesSummary = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(function (cell) { return String(cell || '').trim() === ''; })) continue;

    const rowPeriodId = String(row[0] || '').trim();
    if (targetPeriodId && rowPeriodId !== targetPeriodId) {
      continue; // Lewati periode lain
    }

    const empId = String(row[1] || '').trim();
    const empName = String(row[2] || '').trim();

    // ISOLASI PRIVASI: Karyawan HANYA dapat membaca data diri sendiri
    if (!isHRD) {
      const rowIdLower = empId.toLowerCase();
      const rowNameLower = empName.toLowerCase();
      const cleanRowIdDigits = rowIdLower.replace(/[^0-9]/g, '');
      const cleanUserDigits = currentUsername.replace(/[^0-9]/g, '');

      const isMatch = (
        (rowIdLower && rowIdLower === currentUsername) ||
        (rowNameLower && rowNameLower === currentFullname) ||
        (cleanRowIdDigits && cleanUserDigits && cleanRowIdDigits === cleanUserDigits)
      );

      if (!isMatch) {
        continue;
      }
    }

    const deptJabatan = String(row[3] || '').trim();
    const basic = cleanSalaryNum(row[4]);
    const totalJamOt = cleanSalaryNum(row[5]);
    const totalNominalOt = cleanSalaryNum(row[6]);
    const totalRitase = cleanSalaryNum(row[7]);
    const totalNominalRit = cleanSalaryNum(row[8]);
    const totalTunjangan = cleanSalaryNum(row[9]);
    const totalGajiBerjalan = cleanSalaryNum(row[10]);
    const hariKerja = cleanSalaryNum(row[11]);
    const rawJson = String(row[12] || '').trim();
    const terakhirDiupdate = row[13] ? String(row[13]) : '';

    let dept = '';
    let jabatan = '';
    if (deptJabatan.indexOf(' - ') !== -1) {
      const parts = deptJabatan.split(' - ');
      dept = parts[0].trim();
      jabatan = parts.slice(1).join(' - ').trim();
    } else {
      dept = deptJabatan;
    }

    if (filterDept && filterDept !== 'SEMUA') {
      if (dept.toLowerCase() !== String(filterDept).trim().toLowerCase()) {
        continue;
      }
    }

    let logs = [];
    try {
      if (rawJson && rawJson.startsWith('[')) {
        logs = JSON.parse(rawJson);
      }
    } catch (e) {
      logs = [];
    }

    // Urutkan log dari tanggal terbaru
    logs.sort(function (a, b) {
      return String(b.tanggal || '').localeCompare(String(a.tanggal || ''));
    });

    logs.forEach(function (log) {
      if (!log) return;
      if (filterTanggal && filterTanggal !== 'SEMUA') {
        const logTgl = String(log.tanggal || '').trim();
        if (logTgl !== String(filterTanggal).trim()) return;
      }

      flatList.push({
        id: log.id || ('GJ-' + (log.tanggal ? log.tanggal.replace(/-/g, '') : '') + '-' + Utilities.getUuid().slice(0, 6)),
        periodId: rowPeriodId,
        tanggal: log.tanggal || '',
        dept: log.dept || dept,
        empId: log.empId || empId,
        empName: log.empName || empName,
        jabatan: log.jabatan || jabatan,
        basic: cleanSalaryNum(log.basic !== undefined ? log.basic : basic),
        jamOvertime: cleanSalaryNum(log.jamOvertime),
        tarifOvertime: cleanSalaryNum(log.tarifOvertime),
        totalOvertime: cleanSalaryNum(log.totalOvertime || (log.jamOvertime * log.tarifOvertime)),
        ritase: cleanSalaryNum(log.ritase),
        tarifRitase: cleanSalaryNum(log.tarifRitase),
        totalRitase: cleanSalaryNum(log.totalRitase || (log.ritase * log.tarifRitase)),
        tunjanganHarian: cleanSalaryNum(log.tunjanganHarian),
        totalGaji: cleanSalaryNum(log.totalGaji),
        catatan: String(log.catatan || '').trim(),
        author: String(log.author || '').trim()
      });
    });

    employeesSummary.push({
      periodId: rowPeriodId,
      id: empId,
      name: empName,
      dept: dept,
      jabatan: jabatan,
      deptJabatan: deptJabatan,
      basic: basic,
      totalJamOt: totalJamOt,
      totalNominalOt: totalNominalOt,
      totalRitase: totalRitase,
      totalNominalRit: totalNominalRit,
      totalTunjangan: totalTunjangan,
      grandTotal: totalGajiBerjalan,
      daysCount: hariKerja,
      logs: logs,
      terakhirDiupdate: terakhirDiupdate
    });
  }

  flatList.sort(function (a, b) {
    return String(b.tanggal || '').localeCompare(String(a.tanggal || ''));
  });

  return {
    success: true,
    list: flatList,
    employees: employeesSummary,
    isHRD: isHRD,
    period: targetPeriodInfo,
    periods: allPeriods
  };
}

/**
 * WRITE API: Simpan / Edit data gaji harian karyawan
 */
function saveDataGaji(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');
  if (!isHRD) {
    return { success: false, message: 'Akses ditolak: Hanya Admin/HRD yang dapat menginput data gaji.' };
  }

  if (!payload || typeof payload !== 'object') {
    return { success: false, message: 'Data gaji tidak valid.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Tentukan periode target
  let targetPeriodId = String(payload.periodId || '').trim();
  let targetPeriodInfo = null;

  if (targetPeriodId) {
    const allPeriodsRes = getPeriodePayrollList(token);
    if (allPeriodsRes.success && allPeriodsRes.list) {
      targetPeriodInfo = allPeriodsRes.list.find(function (p) { return p.id === targetPeriodId; });
    }
  } else {
    targetPeriodInfo = getActivePeriode(ss);
    if (targetPeriodInfo) targetPeriodId = targetPeriodInfo.id;
  }

  if (targetPeriodInfo && targetPeriodInfo.status === 'DITUTUP') {
    return {
      success: false,
      message: 'Periode "' + targetPeriodInfo.name + '" sudah ditutup buku. Data tidak dapat ditambah atau diubah.'
    };
  }

  const tanggal = String(payload.tanggal || '').trim() || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const dept = String(payload.dept || '').trim();
  const empId = String(payload.empId || '').trim();
  const empName = String(payload.empName || '').trim();
  const jabatan = String(payload.jabatan || '').trim();

  if (!tanggal || !empId || !empName) {
    return { success: false, message: 'Mohon lengkapi Tanggal Kerja dan Nama Karyawan.' };
  }

  // Basic salary
  let basic = Math.max(0, cleanSalaryNum(payload.basic));
  if (basic === 0) {
    try {
      const userSheet = ss.getSheetByName('Users');
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let u = 1; u < uData.length; u++) {
          if (String(uData[u][0] || '').trim().toLowerCase() === empId.toLowerCase() ||
              String(uData[u][1] || '').trim().toLowerCase() === empName.toLowerCase()) {
            basic = cleanSalaryNum(uData[u][11]);
            break;
          }
        }
      }
    } catch (ue) {
      console.warn('Lookup basic error:', ue);
    }
  }

  const jamOt = Math.max(0, cleanSalaryNum(payload.jamOvertime));
  const tarifOt = Math.max(0, cleanSalaryNum(payload.tarifOvertime));
  const totalOt = jamOt * tarifOt;

  const ritase = Math.max(0, cleanSalaryNum(payload.ritase));
  const tarifRit = Math.max(0, cleanSalaryNum(payload.tarifRitase));
  const totalRit = ritase * tarifRit;

  const tunjangan = Math.max(0, cleanSalaryNum(payload.tunjanganHarian));
  const catatan = String(payload.catatan || '').trim();
  const author = String(session.user.username || 'Admin').trim();
  const editId = String(payload.id || '').trim();

  const sheet = getOrCreateDataGajiSheet(ss);
  const data = sheet.getDataRange().getValues();

  // Cari baris karyawan di spreadsheet sesuai periodId
  let targetRowIndex = -1;
  let existingLogs = [];
  let existingBasic = basic;
  let existingDeptJabatan = dept ? (dept + (jabatan ? ' - ' + jabatan : '')) : jabatan;

  for (let i = 1; i < data.length; i++) {
    const rowPeriod = String(data[i][0] || '').trim();
    if (targetPeriodId && rowPeriod !== targetPeriodId) continue;

    const rowEmpId = String(data[i][1] || '').trim().toLowerCase();
    const rowEmpName = String(data[i][2] || '').trim().toLowerCase();
    if ((empId && rowEmpId === empId.toLowerCase()) || (empName && rowEmpName === empName.toLowerCase())) {
      targetRowIndex = i + 1; // 1-based index
      if (!existingDeptJabatan && data[i][3]) existingDeptJabatan = String(data[i][3]).trim();
      if (basic === 0 && data[i][4]) existingBasic = cleanSalaryNum(data[i][4]);
      const rawJson = String(data[i][12] || '').trim();
      try {
        if (rawJson && rawJson.startsWith('[')) {
          existingLogs = JSON.parse(rawJson);
        }
      } catch (e) {
        existingLogs = [];
      }
      break;
    }
  }

  if (basic === 0 && existingBasic > 0) basic = existingBasic;

  // Siapkan item catatan harian
  const logId = editId || ('GJ-' + tanggal.replace(/-/g, '') + '-' + Utilities.getUuid().slice(0, 6).toUpperCase());
  const logItem = {
    id: logId,
    periodId: targetPeriodId,
    tanggal: tanggal,
    dept: dept,
    empId: empId,
    empName: empName,
    jabatan: jabatan,
    basic: basic,
    jamOvertime: jamOt,
    tarifOvertime: tarifOt,
    totalOvertime: totalOt,
    ritase: ritase,
    tarifRitase: tarifRit,
    totalRitase: totalRit,
    tunjanganHarian: tunjangan,
    totalGaji: (basic + totalOt + totalRit + tunjangan),
    catatan: catatan,
    author: author,
    updatedAt: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')
  };

  if (editId) {
    const logIdx = existingLogs.findIndex(function (l) { return String(l.id || '').trim() === editId; });
    if (logIdx !== -1) {
      existingLogs[logIdx] = logItem;
    } else {
      existingLogs.push(logItem);
    }
  } else {
    existingLogs.push(logItem);
  }

  // Urutkan riwayat berdasarkan tanggal descending
  existingLogs.sort(function (a, b) {
    return String(b.tanggal || '').localeCompare(String(a.tanggal || ''));
  });

  // Hitung ulang akumulasi total karyawan
  let totalJamOt = 0;
  let totalNominalOt = 0;
  let totalRitase = 0;
  let totalNominalRit = 0;
  let totalTunjangan = 0;
  const uniqueDates = new Set();

  existingLogs.forEach(function (l) {
    totalJamOt += Number(l.jamOvertime || 0);
    totalNominalOt += Number(l.totalOvertime || 0);
    totalRitase += Number(l.ritase || 0);
    totalNominalRit += Number(l.totalRitase || 0);
    totalTunjangan += Number(l.tunjanganHarian || 0);
    if (l.tanggal) uniqueDates.add(String(l.tanggal).trim());
  });

  const hariKerja = uniqueDates.size || existingLogs.length;
  const totalGajiBerjalan = basic + totalNominalOt + totalNominalRit + totalTunjangan;
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
  const jsonStr = JSON.stringify(existingLogs);

  const rowValues = [
    targetPeriodId,
    empId,
    empName,
    existingDeptJabatan,
    basic,
    totalJamOt,
    totalNominalOt,
    totalRitase,
    totalNominalRit,
    totalTunjangan,
    totalGajiBerjalan,
    hariKerja,
    jsonStr,
    nowStr
  ];

  if (targetRowIndex > 0) {
    // UPDATE BARIS YANG SUDAH ADA
    sheet.getRange(targetRowIndex, 1, 1, DATA_GAJI_HEADERS.length).setValues([rowValues]);
    formatDataGajiDataRange(sheet, 1);
  } else {
    sheet.appendRow(rowValues);
    formatDataGajiDataRange(sheet, 1);
  }

  SpreadsheetApp.flush();

  return {
    success: true,
    message: editId ? 'Data gaji harian berhasil diperbarui.' : 'Data gaji harian berhasil disimpan.',
    id: logItem.id,
    item: logItem,
    totals: {
      totalJamOt: totalJamOt,
      totalNominalOt: totalNominalOt,
      totalRitase: totalRitase,
      totalNominalRit: totalNominalRit,
      totalTunjangan: totalTunjangan,
      totalGajiBerjalan: totalGajiBerjalan,
      hariKerja: hariKerja
    }
  };
}

/**
 * DELETE API: Hapus catatan log harian tertentu
 */
function deleteDataGaji(token, id) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');
  if (!isHRD) {
    return { success: false, message: 'Akses ditolak: Hanya Admin/HRD yang dapat menghapus data gaji.' };
  }

  const cleanId = String(id || '').trim();
  if (!cleanId) return { success: false, message: 'ID data gaji tidak valid.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreateDataGajiSheet(ss);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rawJson = String(data[i][12] || '').trim();
    if (!rawJson.includes(cleanId)) continue;

    // Cek apakah periode baris ini sudah ditutup
    const rowPeriodId = String(data[i][0] || '').trim();
    const pSheet = getOrCreatePeriodePayrollSheet(ss);
    const pData = pSheet.getDataRange().getValues();
    for (let p = 1; p < pData.length; p++) {
      if (String(pData[p][0] || '').trim() === rowPeriodId && String(pData[p][4] || '').trim().toUpperCase() === 'DITUTUP') {
        return { success: false, message: 'Periode payroll ini sudah ditutup buku. Data tidak dapat dihapus.' };
      }
    }

    let logs = [];
    try {
      if (rawJson && rawJson.startsWith('[')) {
        logs = JSON.parse(rawJson);
      }
    } catch (e) {
      logs = [];
    }

    const logIdx = logs.findIndex(function (l) { return String(l.id || '').trim() === cleanId; });
    if (logIdx !== -1) {
      logs.splice(logIdx, 1);

      // Hitung ulang akumulasi total
      let totalJamOt = 0;
      let totalNominalOt = 0;
      let totalRitase = 0;
      let totalNominalRit = 0;
      let totalTunjangan = 0;
      const uniqueDates = new Set();

      logs.forEach(function (l) {
        totalJamOt += Number(l.jamOvertime || 0);
        totalNominalOt += Number(l.totalOvertime || 0);
        totalRitase += Number(l.ritase || 0);
        totalNominalRit += Number(l.totalRitase || 0);
        totalTunjangan += Number(l.tunjanganHarian || 0);
        if (l.tanggal) uniqueDates.add(String(l.tanggal).trim());
      });

      const basic = cleanSalaryNum(data[i][4]);
      const hariKerja = uniqueDates.size || logs.length;
      const totalGajiBerjalan = basic + totalNominalOt + totalNominalRit + totalTunjangan;
      const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
      const jsonStr = JSON.stringify(logs);

      // Update kolom 6 s/d 14 di baris karyawan tersebut
      sheet.getRange(i + 1, 6, 1, 9).setValues([[
        totalJamOt,
        totalNominalOt,
        totalRitase,
        totalNominalRit,
        totalTunjangan,
        totalGajiBerjalan,
        hariKerja,
        jsonStr,
        nowStr
      ]]);
      formatDataGajiDataRange(sheet, 1);
      SpreadsheetApp.flush();

      return {
        success: true,
        message: 'Catatan harian berhasil dihapus.',
        totals: {
          totalJamOt: totalJamOt,
          totalNominalOt: totalNominalOt,
          totalRitase: totalRitase,
          totalNominalRit: totalNominalRit,
          totalTunjangan: totalTunjangan,
          totalGajiBerjalan: totalGajiBerjalan,
          hariKerja: hariKerja
        }
      };
    }
  }

  return { success: false, message: 'Catatan harian tidak ditemukan.' };
}
