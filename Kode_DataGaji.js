// ==========================================================================
// MODUL INPUT & PENGELOLAAN DATA GAJI KARYAWAN (1 BARIS PER KARYAWAN + JSON LOG)
// File: Kode_DataGaji.js
// 
// Solusi Database Cerdas:
// 1. Di database Google Sheets 'DataGaji', HANYA ADA 1 BARIS PER KARYAWAN
//    (misal 37 karyawan = tetap 37 baris, tidak akan bertambah ribuan baris).
// 2. Setiap baris menyimpan akumulasi total berjalan (Overtime, Ritase, Tunjangan, Total Gaji, Hari Kerja).
// 3. Rincian harian per tanggal disimpan dalam format JSON di kolom 'Rincian_Log_Harian'.
// 4. Saat input/edit/hapus harian, sistem memperbarui baris karyawan tersebut secara in-place.
// ==========================================================================

const DATA_GAJI_HEADERS = [
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

  // Kolom Rupiah: Gaji Pokok (4), Total Nominal Overtime (6), Total Nominal Ritase (8), Total Tunjangan (9), Total Gaji Berjalan (10)
  sheet.getRange(startRow, 4, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 6, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 8, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 9, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 10, count, 1).setNumberFormat('"Rp "#,##0');

  // Kolom Numerik/Jumlah: Total Jam Overtime (5), Total Ritase (7), Hari Kerja (11)
  sheet.getRange(startRow, 5, count, 1).setNumberFormat('#,##0.##');
  sheet.getRange(startRow, 7, count, 1).setNumberFormat('#,##0');
  sheet.getRange(startRow, 11, count, 1).setNumberFormat('#,##0');
}

/**
 * Sinkronisasi karyawan dari sheet 'Users' agar seluruh karyawan otomatis memiliki 1 baris
 */
function syncEmployeesToDataGaji(ss, sheet) {
  if (!sheet) return;
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return;
  const userData = userSheet.getDataRange().getValues();
  if (!userData || userData.length <= 1) return;

  const currentData = sheet.getDataRange().getValues();
  const existingIds = new Set();
  const existingNames = new Set();

  for (let i = 1; i < currentData.length; i++) {
    const id = String(currentData[i][0] || '').trim().toLowerCase();
    const name = String(currentData[i][1] || '').trim().toLowerCase();
    if (id) existingIds.add(id);
    if (name) existingNames.add(name);
  }

  const rowsToAppend = [];
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  for (let u = 1; u < userData.length; u++) {
    const uId = String(userData[u][0] || '').trim();
    const uName = String(userData[u][1] || '').trim();
    if (!uId && !uName) continue;

    if (existingIds.has(uId.toLowerCase()) || existingNames.has(uName.toLowerCase())) {
      continue;
    }

    const uJabatan = String(userData[u][2] || '').trim();
    const uDept = String(userData[u][3] || '').trim();
    const uBasic = cleanSalaryNum(userData[u][11]);
    const deptJabatan = uDept ? (uDept + (uJabatan ? ' - ' + uJabatan : '')) : uJabatan;

    rowsToAppend.push([
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
 * Otomatis migrasi data lama (format 1 baris per input harian) menjadi
 * format cerdas 1 baris per karyawan dengan kolom JSON log harian.
 */
function migrateOldDataGajiToSummaryFormat(sheet, ss) {
  if (!sheet) return;
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();

  const allData = sheet.getDataRange().getValues();
  if (!allData || allData.length <= 1) {
    sheet.clear();
    sheet.appendRow(DATA_GAJI_HEADERS);
    formatDataGajiHeaders(sheet);
    syncEmployeesToDataGaji(ss, sheet);
    return;
  }

  const oldHeaders = allData[0].map(function (h) {
    return String(h || '').trim().toLowerCase().replace(/[\s_\-]+/g, ' ');
  });

  function findCol(keywords, fallbackIndex) {
    const idx = oldHeaders.findIndex(function (h) {
      return keywords.some(function (k) { return h.includes(k) || h === k; });
    });
    return idx !== -1 ? idx : fallbackIndex;
  }

  const idIdx = findCol(['id gaji', 'id_gaji', 'kode gaji', 'id'], 0);
  const tglIdx = findCol(['tanggal', 'tgl', 'date'], 1);
  const deptIdx = findCol(['departemen', 'dept', 'divisi'], 2);
  const empIdIdx = findCol(['id karyawan', 'karyawan id', 'nik', 'nip', 'badge', 'id_karyawan'], 3);
  const nameIdx = findCol(['nama karyawan', 'nama lengkap', 'nama', 'name'], 4);
  const jabIdx = findCol(['jabatan', 'posisi', 'role'], 5);
  const basicIdx = findCol(['gaji pokok', 'basic', 'gapok'], 6);
  const jamOtIdx = findCol(['jam overtime', 'jam lembur', 'overtime jam', 'lembur jam'], 7);
  const tarifOtIdx = findCol(['tarif overtime', 'tarif lembur', 'harga lembur', 'rate overtime'], 8);
  const totalOtIdx = findCol(['total overtime', 'total lembur', 'subtotal overtime'], 9);
  const ritIdx = findCol(['jumlah ritase', 'ritase', 'total rit', 'rit'], 10);
  const tarifRitIdx = findCol(['tarif ritase', 'harga ritase', 'rate ritase'], 11);
  const totalRitIdx = findCol(['total ritase', 'subtotal ritase'], 12);
  const tunjIdx = findCol(['tunjangan harian', 'tunjangan', 'uang makan', 'uang lapangan'], 13);
  const totalGajiIdx = findCol(['total gaji', 'total upah', 'take home pay', 'penghasilan'], 14);
  const noteIdx = findCol(['catatan', 'keterangan', 'note'], 15);
  const authorIdx = findCol(['diinput', 'oleh', 'author', 'admin', 'user'], 16);

  // Kumpulkan data lama dan kelompokkan per karyawan
  const groupedByEmp = {};

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (!row || row.every(function (cell) { return String(cell || '').trim() === ''; })) continue;

    const rowEmpId = String(row[empIdIdx] !== undefined ? row[empIdIdx] : (row[3] || '')).trim();
    const rowEmpName = String(row[nameIdx] !== undefined ? row[nameIdx] : (row[4] || '')).trim();
    if (!rowEmpId && !rowEmpName) continue;

    const groupKey = (rowEmpId || rowEmpName).toLowerCase();
    if (!groupedByEmp[groupKey]) {
      groupedByEmp[groupKey] = {
        empId: rowEmpId,
        empName: rowEmpName,
        dept: String(row[deptIdx] !== undefined ? row[deptIdx] : (row[2] || '')).trim(),
        jabatan: String(row[jabIdx] !== undefined ? row[jabIdx] : (row[5] || '')).trim(),
        basic: cleanSalaryNum(row[basicIdx] !== undefined ? row[basicIdx] : row[6]),
        logs: []
      };
    }

    const rawTgl = row[tglIdx] !== undefined ? row[tglIdx] : row[1];
    let tglStr = '';
    if (rawTgl instanceof Date) {
      tglStr = Utilities.formatDate(rawTgl, 'Asia/Jakarta', 'yyyy-MM-dd');
    } else {
      tglStr = String(rawTgl || '').trim();
      if (tglStr.length > 10 && tglStr.includes('T')) {
        tglStr = tglStr.split('T')[0];
      }
    }

    const jamOt = cleanSalaryNum(row[jamOtIdx] !== undefined ? row[jamOtIdx] : row[7]);
    const tarifOt = cleanSalaryNum(row[tarifOtIdx] !== undefined ? row[tarifOtIdx] : row[8]);
    const totalOtRaw = cleanSalaryNum(row[totalOtIdx] !== undefined ? row[totalOtIdx] : row[9]);
    const totalOt = totalOtRaw > 0 ? totalOtRaw : (jamOt * tarifOt);

    const ritase = cleanSalaryNum(row[ritIdx] !== undefined ? row[ritIdx] : row[10]);
    const tarifRit = cleanSalaryNum(row[tarifRitIdx] !== undefined ? row[tarifRitIdx] : row[11]);
    const totalRitRaw = cleanSalaryNum(row[totalRitIdx] !== undefined ? row[totalRitIdx] : row[12]);
    const totalRit = totalRitRaw > 0 ? totalRitRaw : (ritase * tarifRit);

    const tunjangan = cleanSalaryNum(row[tunjIdx] !== undefined ? row[tunjIdx] : row[13]);
    const totalGajiRaw = cleanSalaryNum(row[totalGajiIdx] !== undefined ? row[totalGajiIdx] : row[14]);
    const totalGaji = totalGajiRaw > 0 ? totalGajiRaw : (groupedByEmp[groupKey].basic + totalOt + totalRit + tunjangan);
    const catatan = String(row[noteIdx] !== undefined ? row[noteIdx] : (row[15] || '')).trim();
    const author = String(row[authorIdx] !== undefined ? row[authorIdx] : (row[16] || '')).trim();
    const rowId = String(row[idIdx] !== undefined ? row[idIdx] : (row[0] || '')).trim() || ('GJ-' + (tglStr ? tglStr.replace(/-/g, '') : 'MIG') + '-' + i);

    groupedByEmp[groupKey].logs.push({
      id: rowId,
      tanggal: tglStr,
      dept: groupedByEmp[groupKey].dept,
      empId: rowEmpId,
      empName: rowEmpName,
      jabatan: groupedByEmp[groupKey].jabatan,
      basic: groupedByEmp[groupKey].basic,
      jamOvertime: jamOt,
      tarifOvertime: tarifOt,
      totalOvertime: totalOt,
      ritase: ritase,
      tarifRitase: tarifRit,
      totalRitase: totalRit,
      tunjanganHarian: tunjangan,
      totalGaji: totalGaji,
      catatan: catatan,
      author: author
    });
  }

  // Ambil master karyawan dari sheet Users
  const userSheet = ss.getSheetByName('Users');
  const userData = userSheet ? userSheet.getDataRange().getValues() : [];
  const summaryRows = [];
  const handledKeys = new Set();
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  for (let u = 1; u < userData.length; u++) {
    const uId = String(userData[u][0] || '').trim();
    const uName = String(userData[u][1] || '').trim();
    if (!uId && !uName) continue;

    const uJabatan = String(userData[u][2] || '').trim();
    const uDept = String(userData[u][3] || '').trim();
    const uBasic = cleanSalaryNum(userData[u][11]);

    const keyById = uId.toLowerCase();
    const keyByName = uName.toLowerCase();
    const group = groupedByEmp[keyById] || groupedByEmp[keyByName];

    let logs = [];
    if (group && Array.isArray(group.logs)) {
      logs = group.logs;
      handledKeys.add(keyById);
      handledKeys.add(keyByName);
    }

    logs.sort(function (a, b) {
      return String(b.tanggal || '').localeCompare(String(a.tanggal || ''));
    });

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

    const hariKerja = uniqueDates.size || logs.length;
    const totalGajiBerjalan = uBasic + totalNominalOt + totalNominalRit + totalTunjangan;
    const deptJabatan = uDept ? (uDept + (uJabatan ? ' - ' + uJabatan : '')) : uJabatan;

    summaryRows.push([
      uId,
      uName,
      deptJabatan,
      uBasic,
      totalJamOt,
      totalNominalOt,
      totalRitase,
      totalNominalRit,
      totalTunjangan,
      totalGajiBerjalan,
      hariKerja,
      JSON.stringify(logs),
      nowStr
    ]);
  }

  // Masukkan data log karyawan yang tidak ada di Users (jika ada)
  Object.keys(groupedByEmp).forEach(function (k) {
    if (handledKeys.has(k)) return;
    const g = groupedByEmp[k];
    const logs = g.logs || [];
    logs.sort(function (a, b) {
      return String(b.tanggal || '').localeCompare(String(a.tanggal || ''));
    });

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

    const hariKerja = uniqueDates.size || logs.length;
    const totalGajiBerjalan = g.basic + totalNominalOt + totalNominalRit + totalTunjangan;
    const deptJabatan = g.dept ? (g.dept + (g.jabatan ? ' - ' + g.jabatan : '')) : g.jabatan;

    summaryRows.push([
      g.empId,
      g.empName,
      deptJabatan,
      g.basic,
      totalJamOt,
      totalNominalOt,
      totalRitase,
      totalNominalRit,
      totalTunjangan,
      totalGajiBerjalan,
      hariKerja,
      JSON.stringify(logs),
      nowStr
    ]);
  });

  // Tulis ulang sheet dengan format 1 baris per karyawan
  sheet.clear();
  sheet.appendRow(DATA_GAJI_HEADERS);
  formatDataGajiHeaders(sheet);

  if (summaryRows.length > 0) {
    sheet.getRange(2, 1, summaryRows.length, DATA_GAJI_HEADERS.length).setValues(summaryRows);
    formatDataGajiDataRange(sheet, summaryRows.length);
  }

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

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(DATA_GAJI_HEADERS);
    formatDataGajiHeaders(sheet);
    syncEmployeesToDataGaji(ss, sheet);
  } else {
    // Periksa apakah format header masih format lama
    const firstHeader = String(sheet.getRange(1, 1).getValue() || '').trim().toLowerCase();
    if (firstHeader !== 'id karyawan') {
      migrateOldDataGajiToSummaryFormat(sheet, ss);
    }
  }

  return sheet;
}

/**
 * READ API: Mengambil data gaji karyawan
 * Mengembalikan:
 * - list: daftar semua log harian (diurai dari kolom JSON) untuk rincian riwayat
 * - employees: daftar ringkasan 1 baris per karyawan dengan total akumulasi
 */
function getDataGaji(token, filterDept, filterTanggal) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');

  const currentUsername = String(session.user.username || '').trim().toLowerCase();
  const currentFullname = String(session.user.fullname || '').trim().toLowerCase();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  const sheet = getOrCreateDataGajiSheet(ss);
  const data = sheet.getDataRange().getValues();

  if (!data || data.length <= 1) {
    return { success: true, list: [], employees: [], isHRD: isHRD };
  }

  const flatList = [];
  const employeesSummary = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(function (cell) { return String(cell || '').trim() === ''; })) continue;

    const empId = String(row[0] || '').trim();
    const empName = String(row[1] || '').trim();

    // ISOLASI PRIVASI KETAT: Jika bukan Admin/HRD, HANYA data diri sendiri yang diambil!
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
        continue; // Lewati data karyawan lain
      }
    }
    const deptJabatan = String(row[2] || '').trim();
    const basic = cleanSalaryNum(row[3]);
    const totalJamOt = cleanSalaryNum(row[4]);
    const totalNominalOt = cleanSalaryNum(row[5]);
    const totalRitase = cleanSalaryNum(row[6]);
    const totalNominalRit = cleanSalaryNum(row[7]);
    const totalTunjangan = cleanSalaryNum(row[8]);
    const totalGajiBerjalan = cleanSalaryNum(row[9]);
    const hariKerja = cleanSalaryNum(row[10]);
    const rawJson = String(row[11] || '').trim();
    const terakhirDiupdate = row[12] ? String(row[12]) : '';

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
      console.warn('Gagal parse JSON log di baris ' + (i + 1) + ':', e);
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

  // Urutkan flatList tanggal terbaru paling atas
  flatList.sort(function (a, b) {
    return String(b.tanggal || '').localeCompare(String(a.tanggal || ''));
  });

  return {
    success: true,
    list: flatList,
    employees: employeesSummary,
    isHRD: isHRD
  };
}

/**
 * WRITE API: Simpan / Edit data gaji harian karyawan
 * - Menemukan baris karyawan terkait di spreadsheet
 * - Meng-update akumulasi total di baris tersebut
 * - Menyimpan rincian harian ke dalam kolom JSON 'Rincian_Log_Harian'
 * - Baris baru di spreadsheet TIDAK BERTAMBAH!
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

  const tanggal = String(payload.tanggal || '').trim() || Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
  const dept = String(payload.dept || '').trim();
  const empId = String(payload.empId || '').trim();
  const empName = String(payload.empName || '').trim();
  const jabatan = String(payload.jabatan || '').trim();

  if (!tanggal || !empId || !empName) {
    return { success: false, message: 'Mohon lengkapi Tanggal Kerja dan Nama Karyawan.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

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

  // Cari baris karyawan di spreadsheet
  let targetRowIndex = -1;
  let existingLogs = [];
  let existingBasic = basic;
  let existingDeptJabatan = dept ? (dept + (jabatan ? ' - ' + jabatan : '')) : jabatan;

  for (let i = 1; i < data.length; i++) {
    const rowEmpId = String(data[i][0] || '').trim().toLowerCase();
    const rowEmpName = String(data[i][1] || '').trim().toLowerCase();
    if ((empId && rowEmpId === empId.toLowerCase()) || (empName && rowEmpName === empName.toLowerCase())) {
      targetRowIndex = i + 1; // 1-based index
      if (!existingDeptJabatan && data[i][2]) existingDeptJabatan = String(data[i][2]).trim();
      if (basic === 0 && data[i][3]) existingBasic = cleanSalaryNum(data[i][3]);
      const rawJson = String(data[i][11] || '').trim();
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
    // UPDATE BARIS YANG SUDAH ADA (BARIS TIDAK BERTAMBAH!)
    sheet.getRange(targetRowIndex, 1, 1, DATA_GAJI_HEADERS.length).setValues([rowValues]);
    formatDataGajiDataRange(sheet, 1);
  } else {
    // Baris baru hanya jika karyawan ini belum ada di sheet
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
 * - Mencari baris karyawan yang memuat ID catatan tersebut di kolom JSON
 * - Menghapus item dari array JSON
 * - Menghitung ulang akumulasi total
 * - Memperbarui baris tersebut in-place (baris karyawan tetap ada)
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
    const rawJson = String(data[i][11] || '').trim();
    if (!rawJson.includes(cleanId)) continue;

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

      const basic = cleanSalaryNum(data[i][3]);
      const hariKerja = uniqueDates.size || logs.length;
      const totalGajiBerjalan = basic + totalNominalOt + totalNominalRit + totalTunjangan;
      const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
      const jsonStr = JSON.stringify(logs);

      // Update kolom 5 s/d 13 di baris karyawan tersebut
      sheet.getRange(i + 1, 5, 1, 9).setValues([[
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

      SpreadsheetApp.flush();
      return { success: true, message: 'Catatan harian berhasil dihapus.' };
    }
  }

  return { success: false, message: 'Data gaji dengan ID tersebut tidak ditemukan.' };
}
