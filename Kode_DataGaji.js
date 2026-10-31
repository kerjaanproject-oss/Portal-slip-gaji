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
  'ID Periode',                // 0 (col 1)
  'ID Karyawan',               // 1 (col 2)
  'Nama Karyawan',             // 2 (col 3)
  'Departemen & Jabatan',      // 3 (col 4)
  'Gaji Pokok',                // 4 (col 5)
  'Total Jam Overtime',        // 5 (col 6)
  'Total Nominal Overtime',    // 6 (col 7)
  'Total Jam SPL',             // 7 (col 8)
  'Total Nominal SPL',         // 8 (col 9)
  'Total Ritase',              // 9 (col 10)
  'Total Nominal Ritase',      // 10 (col 11)
  'Total Tunjangan',           // 11 (col 12)
  'Total Gaji Berjalan',       // 12 (col 13)
  'Hari Kerja',                // 13 (col 14)
  'Rincian_Log_Harian',        // 14 (col 15)
  'Terakhir Diupdate',         // 15 (col 16)
  'Potongan BPJS Kesehatan',   // 16 (col 17)
  'Potongan BPJS TK',          // 17 (col 18)
  'Potongan PPh 21',           // 18 (col 19)
  'Pinjaman Kasbon',           // 19 (col 20)
  'Catatan Kasbon',            // 20 (col 21)
  'Potongan Lain',             // 21 (col 22)
  'Catatan Potongan Lain',     // 22 (col 23)
  'Total Potongan',            // 23 (col 24)
  'Take Home Pay',             // 24 (col 25)
  'Detail Potongan JSON'       // 25 (col 26)
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

  // Kolom Rupiah: Gaji Pokok (5), Total Nominal Overtime (7), Total Nominal SPL (9), Total Nominal Ritase (11), Total Tunjangan (12), Total Gaji Berjalan (13)
  sheet.getRange(startRow, 5, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 7, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 9, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 11, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 12, count, 1).setNumberFormat('"Rp "#,##0');
  sheet.getRange(startRow, 13, count, 1).setNumberFormat('"Rp "#,##0');

  // Kolom Numerik/Jumlah: Total Jam Overtime (6), Total Jam SPL (8), Total Ritase (10), Hari Kerja (14)
  sheet.getRange(startRow, 6, count, 1).setNumberFormat('#,##0.##');
  sheet.getRange(startRow, 8, count, 1).setNumberFormat('#,##0.##');
  sheet.getRange(startRow, 10, count, 1).setNumberFormat('#,##0');
  sheet.getRange(startRow, 14, count, 1).setNumberFormat('#,##0');

  // Kolom Rupiah Potongan & THP: BPJS Kes (17), BPJS TK (18), PPh 21 (19), Kasbon (20), Potongan Lain (22), Total Potongan (24), Take Home Pay (25)
  if (sheet.getLastColumn() >= 25) {
    sheet.getRange(startRow, 17, count, 1).setNumberFormat('"Rp "#,##0');
    sheet.getRange(startRow, 18, count, 1).setNumberFormat('"Rp "#,##0');
    sheet.getRange(startRow, 19, count, 1).setNumberFormat('"Rp "#,##0');
    sheet.getRange(startRow, 20, count, 1).setNumberFormat('"Rp "#,##0');
    sheet.getRange(startRow, 22, count, 1).setNumberFormat('"Rp "#,##0');
    sheet.getRange(startRow, 24, count, 1).setNumberFormat('"Rp "#,##0');
    sheet.getRange(startRow, 25, count, 1).setNumberFormat('"Rp "#,##0');
  }
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

    const bpjsKesSet = String(userData[u][12] || 'Tidak Ikut').trim();
    let bpjsKesVal = cleanSalaryNum(userData[u][13]);
    if (bpjsKesVal === 0) {
      if (bpjsKesSet === 'Kelas 1') bpjsKesVal = 150000;
      else if (bpjsKesSet === 'Kelas 2') bpjsKesVal = 100000;
      else if (bpjsKesSet === 'Kelas 3') bpjsKesVal = 35000;
      else if (bpjsKesSet === 'PPU 1%') bpjsKesVal = Math.round(uBasic * 0.01);
    }
    const bpjsTKSet = String(userData[u][14] || 'Tidak Ikut').trim();
    let bpjsTKVal = cleanSalaryNum(userData[u][15]);
    if (bpjsTKVal === 0 && bpjsTKSet.includes('3%')) {
      bpjsTKVal = Math.round(uBasic * 0.03);
    }
    const pph21Val = cleanSalaryNum(userData[u][16]);

    const totalPotongan = bpjsKesVal + bpjsTKVal + pph21Val;
    const takeHomePay = Math.max(0, uBasic - totalPotongan);
    const detailPotonganJson = JSON.stringify({
      bpjsKesSetting: bpjsKesSet,
      bpjsKesNominal: bpjsKesVal,
      bpjsTKSetting: bpjsTKSet,
      bpjsTKNominal: bpjsTKVal,
      pph21Nominal: pph21Val
    });

    rowsToAppend.push([
      periodId,
      uId,
      uName,
      deptJabatan,
      uBasic,
      0,
      0,
      0, // Total Jam SPL
      0, // Total Nominal SPL
      0, // Total Ritase
      0, // Total Nominal Ritase
      0, // Total Tunjangan
      uBasic, // Total Gaji Berjalan
      0, // Hari Kerja
      '[]',
      nowStr,
      bpjsKesVal,
      bpjsTKVal,
      pph21Val,
      0, // Pinjaman Kasbon
      '', // Catatan Kasbon
      0, // Potongan Lain
      '', // Catatan Potongan Lain
      totalPotongan,
      takeHomePay,
      detailPotonganJson
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

    // Periksa apakah kolom Total Jam SPL & Total Nominal SPL sudah ada
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
        return String(h || '').trim().toLowerCase();
      });
      if (!existingHeaders.includes('total jam spl')) {
        const nomOtIdx = existingHeaders.findIndex(function (h) {
          return h.includes('nominal overtime') || h.includes('total nominal overtime');
        });
        if (nomOtIdx !== -1) {
          sheet.insertColumnsAfter(nomOtIdx + 1, 2);
          sheet.getRange(1, nomOtIdx + 2, 1, 2).setValues([['Total Jam SPL', 'Total Nominal SPL']]);
          const lastR = sheet.getLastRow();
          if (lastR > 1) {
            sheet.getRange(2, nomOtIdx + 2, lastR - 1, 2).setValues(
              new Array(lastR - 1).fill([0, 0])
            );
          }
          formatDataGajiHeaders(sheet);
          formatDataGajiDataRange(sheet, lastR - 1);
        }
      }

      // Periksa apakah kolom Potongan sudah ada
      const existingHeadersNow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function (h) {
        return String(h || '').trim().toLowerCase();
      });
      if (!existingHeadersNow.includes('potongan bpjs kesehatan')) {
        const lastC = sheet.getLastColumn();
        const newCols = [
          'Potongan BPJS Kesehatan',
          'Potongan BPJS TK',
          'Potongan PPh 21',
          'Pinjaman Kasbon',
          'Catatan Kasbon',
          'Potongan Lain',
          'Catatan Potongan Lain',
          'Total Potongan',
          'Take Home Pay',
          'Detail Potongan JSON'
        ];
        sheet.insertColumnsAfter(lastC, newCols.length);
        sheet.getRange(1, lastC + 1, 1, newCols.length).setValues([newCols]);
        const lastR = sheet.getLastRow();
        if (lastR > 1) {
          const defaultVals = new Array(lastR - 1).fill([0, 0, 0, 0, '', 0, '', 0, 0, '{}']);
          sheet.getRange(2, lastC + 1, lastR - 1, newCols.length).setValues(defaultVals);
        }
        formatDataGajiHeaders(sheet);
        formatDataGajiDataRange(sheet, lastR - 1);
      }
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

    // Deteksi index kolom secara dinamis berdasarkan header baris pertama
    const hList = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
    const deptJabIdx = hList.findIndex(function (h) { return h.includes('departemen & jabatan') || h.includes('departemen'); });
    const basicIdx = hList.findIndex(function (h) { return h.includes('gaji pokok'); });
    const jamOtIdx = hList.findIndex(function (h) { return h === 'total jam overtime' || h.includes('jam overtime'); });
    const nomOtIdx = hList.findIndex(function (h) { return h === 'total nominal overtime' || h.includes('nominal overtime'); });
    const jamSplIdx = hList.findIndex(function (h) { return h === 'total jam spl' || h.includes('jam spl'); });
    const nomSplIdx = hList.findIndex(function (h) { return h === 'total nominal spl' || h.includes('nominal spl'); });
    const ritaseIdx = hList.findIndex(function (h) { return h === 'total ritase' || (h.includes('ritase') && !h.includes('nominal')); });
    const nomRitIdx = hList.findIndex(function (h) { return h === 'total nominal ritase' || h.includes('nominal ritase'); });
    const tunjIdx = hList.findIndex(function (h) { return h === 'total tunjangan' || h.includes('tunjangan'); });
    const grandTotalIdx = hList.findIndex(function (h) { return h === 'total gaji berjalan' || h.includes('total gaji'); });
    const hariKerjaIdx = hList.findIndex(function (h) { return h === 'hari kerja'; });
    const jsonIdx = hList.findIndex(function (h) { return h.includes('rincian_log_harian') || h.includes('log'); });
    const updatedIdx = hList.findIndex(function (h) { return h.includes('terakhir diupdate'); });

    const deptJabatan = deptJabIdx !== -1 ? String(row[deptJabIdx] || '').trim() : String(row[3] || '').trim();
    const basic = basicIdx !== -1 ? cleanSalaryNum(row[basicIdx]) : cleanSalaryNum(row[4]);
    const totalJamOt = jamOtIdx !== -1 ? cleanSalaryNum(row[jamOtIdx]) : cleanSalaryNum(row[5]);
    const totalNominalOt = nomOtIdx !== -1 ? cleanSalaryNum(row[nomOtIdx]) : cleanSalaryNum(row[6]);
    const totalJamSPL = jamSplIdx !== -1 ? cleanSalaryNum(row[jamSplIdx]) : 0;
    const totalNominalSPL = nomSplIdx !== -1 ? cleanSalaryNum(row[nomSplIdx]) : 0;
    const totalRitase = ritaseIdx !== -1 ? cleanSalaryNum(row[ritaseIdx]) : cleanSalaryNum(row[7]);
    const totalNominalRit = nomRitIdx !== -1 ? cleanSalaryNum(row[nomRitIdx]) : cleanSalaryNum(row[8]);
    const totalTunjangan = tunjIdx !== -1 ? cleanSalaryNum(row[tunjIdx]) : cleanSalaryNum(row[9]);
    const totalGajiBerjalan = grandTotalIdx !== -1 ? cleanSalaryNum(row[grandTotalIdx]) : cleanSalaryNum(row[10]);
    const hariKerja = hariKerjaIdx !== -1 ? cleanSalaryNum(row[hariKerjaIdx]) : cleanSalaryNum(row[11]);
    const rawJson = jsonIdx !== -1 ? String(row[jsonIdx] || '').trim() : String(row[12] || '').trim();
    const terakhirDiupdate = updatedIdx !== -1 ? String(row[updatedIdx] || '') : (row[13] ? String(row[13]) : '');

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
        jamSPL: cleanSalaryNum(log.jamSPL),
        tarifSPL: cleanSalaryNum(log.tarifSPL),
        totalSPL: cleanSalaryNum(log.totalSPL || (log.jamSPL * log.tarifSPL)),
        splId: String(log.splId || '').trim(),
        ritase: cleanSalaryNum(log.ritase),
        tarifRitase: cleanSalaryNum(log.tarifRitase),
        totalRitase: cleanSalaryNum(log.totalRitase || (log.ritase * log.tarifRitase)),
        tunjanganHarian: cleanSalaryNum(log.tunjanganHarian),
        totalGaji: cleanSalaryNum(log.totalGaji),
        catatan: String(log.catatan || '').trim(),
        author: String(log.author || '').trim()
      });
    });

    const potKesIdx = hList.findIndex(function (h) { return h.includes('bpjs kesehatan') || h.includes('bpjs kes'); });
    const potTkIdx = hList.findIndex(function (h) { return h.includes('bpjs ketenagakerjaan') || h.includes('bpjs tk'); });
    const potPphIdx = hList.findIndex(function (h) { return h.includes('pph 21') || h.includes('pph21'); });
    const kasbonIdx = hList.findIndex(function (h) { return h.includes('kasbon') || h.includes('pinjaman'); });
    const catKasbonIdx = hList.findIndex(function (h) { return h.includes('catatan kasbon'); });
    const potLainIdx = hList.findIndex(function (h) { return h.includes('potongan lain') && !h.includes('catatan'); });
    const catPotLainIdx = hList.findIndex(function (h) { return h.includes('catatan potongan lain'); });
    const totPotIdx = hList.findIndex(function (h) { return h.includes('total potongan'); });
    const thpIdx = hList.findIndex(function (h) { return h.includes('take home pay') || h.includes('thp'); });
    const detailPotIdx = hList.findIndex(function (h) { return h.includes('detail potongan'); });

    const potKes = potKesIdx !== -1 ? cleanSalaryNum(row[potKesIdx]) : 0;
    const potTk = potTkIdx !== -1 ? cleanSalaryNum(row[potTkIdx]) : 0;
    const potPph = potPphIdx !== -1 ? cleanSalaryNum(row[potPphIdx]) : 0;
    const potKasbon = kasbonIdx !== -1 ? cleanSalaryNum(row[kasbonIdx]) : 0;
    const catKasbon = catKasbonIdx !== -1 ? String(row[catKasbonIdx] || '').trim() : '';
    const potLain = potLainIdx !== -1 ? cleanSalaryNum(row[potLainIdx]) : 0;
    const catPotLain = catPotLainIdx !== -1 ? String(row[catPotLainIdx] || '').trim() : '';
    let totPot = totPotIdx !== -1 ? cleanSalaryNum(row[totPotIdx]) : (potKes + potTk + potPph + potKasbon + potLain);
    let thp = thpIdx !== -1 && row[thpIdx] !== '' ? cleanSalaryNum(row[thpIdx]) : Math.max(0, totalGajiBerjalan - totPot);
    let detailPot = {};
    if (detailPotIdx !== -1 && row[detailPotIdx]) {
      try { detailPot = JSON.parse(row[detailPotIdx]); } catch (e) { detailPot = {}; }
    }

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
      totalJamSPL: totalJamSPL,
      totalNominalSPL: totalNominalSPL,
      totalRitase: totalRitase,
      totalNominalRit: totalNominalRit,
      totalTunjangan: totalTunjangan,
      grandTotal: totalGajiBerjalan,
      daysCount: hariKerja,
      logs: logs,
      terakhirDiupdate: terakhirDiupdate,
      potonganBPJSKes: potKes,
      potonganBPJSTK: potTk,
      potonganPPh21: potPph,
      potonganKasbon: potKasbon,
      catatanKasbon: catKasbon,
      potonganLain: potLain,
      catatanPotonganLain: catPotLain,
      totalPotongan: totPot,
      takeHomePay: thp,
      detailPotongan: detailPot
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

  // 1. Overtime (Lembur biasa)
  const jamOt = Math.max(0, cleanSalaryNum(payload.jamOvertime));
  const tarifOt = Math.max(0, cleanSalaryNum(payload.tarifOvertime));
  const totalOt = jamOt * tarifOt;

  // 2. SPL (Surat Perintah Lembur resmi - data terpisah dari overtime)
  const jamSPL = Math.max(0, cleanSalaryNum(payload.jamSPL));
  const tarifSPL = Math.max(0, cleanSalaryNum(payload.tarifSPL));
  const totalSPL = jamSPL * tarifSPL;
  const splId = String(payload.splId || '').trim();

  // 3. Ritase
  const ritase = Math.max(0, cleanSalaryNum(payload.ritase));
  const tarifRit = Math.max(0, cleanSalaryNum(payload.tarifRitase));
  const totalRit = ritase * tarifRit;

  // 4. Tunjangan Harian
  const tunjangan = Math.max(0, cleanSalaryNum(payload.tunjanganHarian));
  const totalHariIni = basic + totalOt + totalSPL + totalRit + tunjangan;

  const catatan = String(payload.catatan || '').trim();
  const author = String(session.user.username || 'Admin').trim();
  const editId = String(payload.id || '').trim();

  const sheet = getOrCreateDataGajiSheet(ss);
  const data = sheet.getDataRange().getValues();
  const headerList = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const jsonColIdx = headerList.findIndex(function (h) { return h.includes('rincian_log_harian') || h.includes('log'); });

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
      const rawJson = jsonColIdx !== -1 ? String(data[i][jsonColIdx] || '').trim() : String(data[i][12] || '').trim();
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

  // Siapkan item catatan harian dengan Overtime dan SPL terpisah
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
    jamSPL: jamSPL,
    tarifSPL: tarifSPL,
    totalSPL: totalSPL,
    splId: splId,
    ritase: ritase,
    tarifRitase: tarifRit,
    totalRitase: totalRit,
    tunjanganHarian: tunjangan,
    totalGaji: totalHariIni,
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

  // Hitung ulang akumulasi total karyawan secara terpisah: Overtime & SPL
  let totalJamOt = 0;
  let totalNominalOt = 0;
  let totalJamSPL = 0;
  let totalNominalSPL = 0;
  let totalRitase = 0;
  let totalNominalRit = 0;
  let totalTunjangan = 0;
  const uniqueDates = new Set();

  existingLogs.forEach(function (l) {
    totalJamOt += Number(l.jamOvertime || 0);
    totalNominalOt += Number(l.totalOvertime || 0);
    totalJamSPL += Number(l.jamSPL || 0);
    totalNominalSPL += Number(l.totalSPL || 0);
    totalRitase += Number(l.ritase || 0);
    totalNominalRit += Number(l.totalRitase || 0);
    totalTunjangan += Number(l.tunjanganHarian || 0);
    if (l.tanggal) uniqueDates.add(String(l.tanggal).trim());
  });

  const hariKerja = uniqueDates.size || existingLogs.length;
  const totalGajiBerjalan = basic + totalNominalOt + totalNominalSPL + totalNominalRit + totalTunjangan;
  const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
  const jsonStr = JSON.stringify(existingLogs);

  let existingPotKes = 0, existingPotTk = 0, existingPotPph = 0;
  let existingKasbon = 0, existingCatKasbon = '';
  let existingPotLain = 0, existingCatPotLain = '';
  let existingDetailPotJson = '{}';
  if (targetRowIndex > 0) {
    const existingRow = data[targetRowIndex - 1];
    const pKesIdx = headerList.findIndex(function (h) { return h.includes('bpjs kesehatan') || h.includes('bpjs kes'); });
    const pTkIdx = headerList.findIndex(function (h) { return h.includes('bpjs ketenagakerjaan') || h.includes('bpjs tk'); });
    const pPphIdx = headerList.findIndex(function (h) { return h.includes('pph 21') || h.includes('pph21'); });
    const pKasIdx = headerList.findIndex(function (h) { return h.includes('kasbon') || h.includes('pinjaman'); });
    const cKasIdx = headerList.findIndex(function (h) { return h.includes('catatan kasbon'); });
    const pLainIdx = headerList.findIndex(function (h) { return h.includes('potongan lain') && !h.includes('catatan'); });
    const cLainIdx = headerList.findIndex(function (h) { return h.includes('catatan potongan lain'); });
    const dPotIdx = headerList.findIndex(function (h) { return h.includes('detail potongan'); });

    if (pKesIdx !== -1) existingPotKes = cleanSalaryNum(existingRow[pKesIdx]);
    if (pTkIdx !== -1) existingPotTk = cleanSalaryNum(existingRow[pTkIdx]);
    if (pPphIdx !== -1) existingPotPph = cleanSalaryNum(existingRow[pPphIdx]);
    if (pKasIdx !== -1) existingKasbon = cleanSalaryNum(existingRow[pKasIdx]);
    if (cKasIdx !== -1) existingCatKasbon = String(existingRow[cKasIdx] || '').trim();
    if (pLainIdx !== -1) existingPotLain = cleanSalaryNum(existingRow[pLainIdx]);
    if (cLainIdx !== -1) existingCatPotLain = String(existingRow[cLainIdx] || '').trim();
    if (dPotIdx !== -1) existingDetailPotJson = String(existingRow[dPotIdx] || '{}');
  }

  const totalPotonganSave = existingPotKes + existingPotTk + existingPotPph + existingKasbon + existingPotLain;
  const takeHomePaySave = Math.max(0, totalGajiBerjalan - totalPotonganSave);

  const rowValues = [
    targetPeriodId,
    empId,
    empName,
    existingDeptJabatan,
    basic,
    totalJamOt,
    totalNominalOt,
    totalJamSPL,
    totalNominalSPL,
    totalRitase,
    totalNominalRit,
    totalTunjangan,
    totalGajiBerjalan,
    hariKerja,
    jsonStr,
    nowStr,
    existingPotKes,
    existingPotTk,
    existingPotPph,
    existingKasbon,
    existingCatKasbon,
    existingPotLain,
    existingCatPotLain,
    totalPotonganSave,
    takeHomePaySave,
    existingDetailPotJson
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
      totalJamSPL: totalJamSPL,
      totalNominalSPL: totalNominalSPL,
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
      let totalJamSPL = 0;
      let totalNominalSPL = 0;
      let totalRitase = 0;
      let totalNominalRit = 0;
      let totalTunjangan = 0;
      const uniqueDates = new Set();

      logs.forEach(function (l) {
        totalJamOt += Number(l.jamOvertime || 0);
        totalNominalOt += Number(l.totalOvertime || 0);
        totalJamSPL += Number(l.jamSPL || 0);
        totalNominalSPL += Number(l.totalSPL || 0);
        totalRitase += Number(l.ritase || 0);
        totalNominalRit += Number(l.totalRitase || 0);
        totalTunjangan += Number(l.tunjanganHarian || 0);
        if (l.tanggal) uniqueDates.add(String(l.tanggal).trim());
      });

      const basic = cleanSalaryNum(data[i][4]);
      const hariKerja = uniqueDates.size || logs.length;
      const totalGajiBerjalan = basic + totalNominalOt + totalNominalSPL + totalNominalRit + totalTunjangan;
      const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
      const jsonStr = JSON.stringify(logs);

      // Update kolom 6 s/d 16 di baris karyawan tersebut
      sheet.getRange(i + 1, 6, 1, 11).setValues([[
        totalJamOt,
        totalNominalOt,
        totalJamSPL,
        totalNominalSPL,
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
          totalJamSPL: totalJamSPL,
          totalNominalSPL: totalNominalSPL,
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


/**
 * Helper: Normalisasi tanggal dari Date object atau String berbagai format ke string YYYY-MM-DD
 */
function normalizeDateYMD(val, ss) {
  if (!val) return '';
  if (val instanceof Date) {
    if (isNaN(val.getTime())) return '';
    let tz = 'Asia/Jakarta';
    try {
      if (ss && typeof ss.getSpreadsheetTimeZone === 'function') {
        tz = ss.getSpreadsheetTimeZone() || 'Asia/Jakarta';
      }
    } catch (e) {}
    try {
      return Utilities.formatDate(val, tz, 'yyyy-MM-dd');
    } catch (e) {
      const y = val.getFullYear();
      const m = String(val.getMonth() + 1).padStart(2, '0');
      const d = String(val.getDate()).padStart(2, '0');
      return y + '-' + m + '-' + d;
    }
  }

  const s = String(val).trim().split('T')[0].split(' ')[0];
  if (!s) return '';

  // YYYY-MM-DD
  if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) {
    const parts = s.split('-');
    return parts[0] + '-' + parts[1].padStart(2, '0') + '-' + parts[2].padStart(2, '0');
  }

  // DD/MM/YYYY atau MM/DD/YYYY
  if (s.indexOf('/') !== -1) {
    const parts = s.split('/');
    if (parts.length === 3) {
      if (parts[2].length === 4) {
        let p0 = parseInt(parts[0], 10);
        let p1 = parseInt(parts[1], 10);
        let day = p0;
        let month = p1;
        if (p0 > 12 && p1 <= 12) {
          day = p0; month = p1;
        } else if (p1 > 12 && p0 <= 12) {
          day = p1; month = p0;
        }
        return parts[2] + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
      } else if (parts[0].length === 4) {
        return parts[0] + '-' + parts[1].padStart(2, '0') + '-' + parts[2].padStart(2, '0');
      }
    }
  }

  // DD-MM-YYYY
  if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(s)) {
    const parts = s.split('-');
    let p0 = parseInt(parts[0], 10);
    let p1 = parseInt(parts[1], 10);
    let day = p0;
    let month = p1;
    if (p0 > 12 && p1 <= 12) {
      day = p0; month = p1;
    } else if (p1 > 12 && p0 <= 12) {
      day = p1; month = p0;
    }
    return parts[2] + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
  }

  return s;
}

/**
 * Helper: Cek apakah status SPL dianggap Disetujui
 */
function isSPLApprovedStatus(statusStr, app1StatusStr, app2StatusStr) {
  const s = String(statusStr || '').trim().toLowerCase();
  const a1 = String(app1StatusStr || '').trim().toLowerCase();
  const a2 = String(app2StatusStr || '').trim().toLowerCase();

  // Jika status utama adalah 'disetujui' atau 'approved'
  if (s === 'disetujui' || s === 'approved' || s.indexOf('disetujui') !== -1) {
    return true;
  }
  // Jika approval tahap 2 (HRD) disetujui
  if (a2 === 'disetujui' || a2.indexOf('disetujui') !== -1) {
    return true;
  }
  // Jika approver 1 sudah disetujui dan tidak ada kolom approver 2
  if (a1 === 'disetujui' && (!a2 || a2 === '-' || a2 === '')) {
    return true;
  }
  return false;
}

/**
 * Helper: Pencocokan fleksibel data karyawan antara SPL dan DataGaji
 */
function isSPLMatchEmployee(rowEmpId, rowEmpName, targetEmpId, targetEmpName, targetUsernames) {
  const rId = String(rowEmpId || '').trim().toLowerCase();
  const rName = String(rowEmpName || '').trim().toLowerCase();
  const tId = String(targetEmpId || '').trim().toLowerCase();
  const tName = String(targetEmpName || '').trim().toLowerCase();

  // 1. Cocok ID langsung
  if (rId && tId && rId === tId) return true;

  // 2. Cocok username / alias
  if (Array.isArray(targetUsernames) && rId && targetUsernames.includes(rId)) return true;

  // 3. Cocok Nama Lengkap langsung
  if (rName && tName && rName === tName) return true;

  // 4. Cocok nama dengan pembersihan spasi / karakter
  if (rName && tName) {
    const cleanR = rName.replace(/[^a-z0-9]/g, '');
    const cleanT = tName.replace(/[^a-z0-9]/g, '');
    if (cleanR && cleanT && cleanR === cleanT) return true;
    if (cleanR.length >= 4 && cleanT.length >= 4) {
      if (cleanR.includes(cleanT) || cleanT.includes(cleanR)) return true;
    }
  }

  // 5. Cocok angka NIK / ID (misal BBB-260620 dengan 260620)
  const rNum = rId.replace(/\D/g, '');
  const tNum = tId.replace(/\D/g, '');
  if (rNum && tNum && rNum === tNum) return true;

  // 6. Jika rId cocok dengan nama tName atau sebaliknya
  if (rId && tName && (rId === tName || tName.includes(rId))) return true;
  if (tId && rName && (tId === rName || rName.includes(tId))) return true;

  return false;
}

/**
 * READ API: Cek SPL yang sudah disetujui (Status: 'Disetujui') untuk Karyawan dan Tanggal tertentu.
 * Digunakan untuk mengisi otomatis Jam SPL di Form Input Gaji Harian.
 */
function getApprovedSPLForSalaryInput(token, empId, tanggalKerja, empName) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  if (!empId || !tanggalKerja) {
    return { success: false, message: 'ID Karyawan dan Tanggal Kerja wajib ditentukan.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const splSheet = getOrCreatePengajuanSPLSheet(ss);
  if (!splSheet) {
    return { success: true, hasSPL: false, totalJam: 0, list: [] };
  }

  const data = splSheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return { success: true, hasSPL: false, totalJam: 0, list: [] };
  }

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id spl') || h === 'id'; });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h === 'username'; });
  const namaIdx = headers.findIndex(function (h) { return h.includes('nama karyawan') || h === 'nama'; });
  const tglLemburIdx = headers.findIndex(function (h) { return h === 'tanggal lembur' || h.includes('tanggal lembur') || h.includes('tgl lembur') || h === 'tanggal'; });
  const shiftIdx = headers.findIndex(function (h) { return h === 'shift' || h.includes('shift'); });
  const jamMulaiIdx = headers.findIndex(function (h) { return h.includes('jam mulai') || h.includes('jam awal'); });
  const jamSelesaiIdx = headers.findIndex(function (h) { return h.includes('jam selesai') || h.includes('jam akhir'); });
  const totalJamIdx = headers.findIndex(function (h) { return h.includes('total jam'); });
  const uraianIdx = headers.findIndex(function (h) { return h.includes('uraian') || h.includes('pekerjaan') || h.includes('alasan'); });

  let statusIdx = headers.findIndex(function (h) { return h === 'status' || h === 'status spl' || h === 'status pengajuan'; });
  if (statusIdx === -1) {
    statusIdx = headers.findIndex(function (h) { return h.includes('status') && !h.includes('approver') && !h.includes('app'); });
  }
  if (statusIdx === -1) {
    statusIdx = headers.findIndex(function (h) { return h.includes('status'); });
  }
  const app1StatusIdx = headers.findIndex(function (h) { return h.includes('approver 1 status') || h.includes('app 1 status') || h.includes('status atasan'); });
  const app2StatusIdx = headers.findIndex(function (h) { return h.includes('approver 2 status') || h.includes('app 2 status') || h.includes('status hrd'); });

  const targetDateStr = normalizeDateYMD(tanggalKerja, ss);
  let targetEmpName = String(empName || '').trim();
  const targetUsernames = [String(empId || '').trim().toLowerCase()];

  // Lookup data karyawan di sheet Data_Karyawan jika nama belum diketahui
  try {
    const userSheet = ss.getSheetByName('Data_Karyawan') || ss.getSheetByName('Users') || ss.getSheetByName('Karyawan');
    if (userSheet) {
      const uData = userSheet.getDataRange().getValues();
      const targetEmpLower = String(empId).trim().toLowerCase();
      const targetEmpNum = targetEmpLower.replace(/\D/g, '');
      for (let u = 1; u < uData.length; u++) {
        const uId = String(uData[u][0] || '').trim().toLowerCase();
        const uName = String(uData[u][1] || '').trim();
        const uNum = uId.replace(/\D/g, '');
        if (uId === targetEmpLower || (targetEmpNum && uNum === targetEmpNum) || (targetEmpName && uName.toLowerCase() === targetEmpName.toLowerCase())) {
          if (!targetEmpName && uName) targetEmpName = uName;
          if (!targetUsernames.includes(uId)) targetUsernames.push(uId);
          break;
        }
      }
    }
  } catch (lookupErr) {
    console.warn('Lookup karyawan error:', lookupErr);
  }

  let sumHours = 0;
  const matchedSPL = [];
  const splIds = [];
  const shifts = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowStatus = statusIdx !== -1 ? String(row[statusIdx] || '').trim() : '';
    const rowApp1 = app1StatusIdx !== -1 ? String(row[app1StatusIdx] || '').trim() : '';
    const rowApp2 = app2StatusIdx !== -1 ? String(row[app2StatusIdx] || '').trim() : '';

    if (!isSPLApprovedStatus(rowStatus, rowApp1, rowApp2)) continue;

    // Cek kecocokan tanggal lembur
    const tglVal = tglLemburIdx !== -1 ? row[tglLemburIdx] : row[6];
    const rowDateStr = normalizeDateYMD(tglVal, ss);

    if (rowDateStr !== targetDateStr) continue;

    // Cek kecocokan Karyawan
    const rowEmpId = empIdIdx !== -1 ? String(row[empIdIdx] || '').trim() : '';
    const rowEmpName = namaIdx !== -1 ? String(row[namaIdx] || '').trim() : '';

    const isMatch = isSPLMatchEmployee(rowEmpId, rowEmpName, empId, targetEmpName, targetUsernames);
    if (!isMatch) continue;

    const sId = idIdx !== -1 ? String(row[idIdx] || '').trim() : ('SPL-' + i);

    // Ambil jam lembur
    let rawHours = totalJamIdx !== -1 ? row[totalJamIdx] : 0;
    let hours = 0;
    if (typeof rawHours === 'number') {
      hours = isNaN(rawHours) ? 0 : rawHours;
    } else {
      const cleanStr = String(rawHours || '').replace(',', '.').replace(/[^0-9.-]/g, '');
      hours = parseFloat(cleanStr) || 0;
    }

    const jMulai = jamMulaiIdx !== -1 ? String(row[jamMulaiIdx] || '').trim() : '';
    const jSelesai = jamSelesaiIdx !== -1 ? String(row[jamSelesaiIdx] || '').trim() : '';

    // Fallback jika hours 0 tapi ada jamMulai dan jamSelesai
    if (hours <= 0 && jMulai && jSelesai) {
      try {
        const mP = jMulai.split(':');
        const sP = jSelesai.split(':');
        if (mP.length >= 2 && sP.length >= 2) {
          const sMin = parseInt(mP[0], 10) * 60 + parseInt(mP[1], 10);
          let eMin = parseInt(sP[0], 10) * 60 + parseInt(sP[1], 10);
          if (eMin < sMin) eMin += 24 * 60;
          const diff = eMin - sMin;
          if (diff > 0) hours = Math.round((diff / 60) * 10) / 10;
        }
      } catch (e) {}
    }

    const shiftVal = shiftIdx !== -1 ? String(row[shiftIdx] || '').trim() : '-';
    const uraianVal = uraianIdx !== -1 ? String(row[uraianIdx] || '').trim() : '-';

    sumHours += hours;
    splIds.push(sId);
    if (shiftVal && shiftVal !== '-' && !shifts.includes(shiftVal)) shifts.push(shiftVal);

    matchedSPL.push({
      id: sId,
      totalJam: hours,
      shift: shiftVal,
      jamMulai: jMulai,
      jamSelesai: jSelesai,
      uraian: uraianVal
    });
  }

  return {
    success: true,
    hasSPL: matchedSPL.length > 0,
    totalJam: Math.round(sumHours * 100) / 100,
    splIds: splIds,
    shifts: shifts.join(', '),
    details: matchedSPL
  };
}

/**
 * WRITE API: Tarik / Sinkronkan semua SPL yang sudah Disetujui dalam 1 periode cut-off payroll
 * ke data log harian & total akumulasi SPL seluruh karyawan (terpisah dari Overtime).
 */
function syncApprovedSPLForPeriod(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');
  if (!isHRD) {
    return { success: false, message: 'Akses ditolak: Hanya Admin/HRD yang dapat melakukan sinkronisasi SPL ke Gaji.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let targetPeriodId = String(payload && payload.periodId || '').trim();
  let defaultTarif = Math.max(0, cleanSalaryNum(payload && (payload.defaultTarifSPL || payload.defaultTarifOvertime)));
  if (!defaultTarif) defaultTarif = 20000;

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

  if (!targetPeriodInfo) {
    return { success: false, message: 'Periode payroll aktif tidak ditemukan.' };
  }

  if (targetPeriodInfo.status === 'DITUTUP') {
    return { success: false, message: 'Periode "' + targetPeriodInfo.name + '" sudah ditutup buku.' };
  }

  const sDate = normalizeDateYMD(targetPeriodInfo.startDate, ss);
  const eDate = normalizeDateYMD(targetPeriodInfo.endDate, ss);

  const splSheet = getOrCreatePengajuanSPLSheet(ss);
  if (!splSheet) {
    return { success: false, message: 'Sheet PengajuanSPL tidak ditemukan.' };
  }

  const splData = splSheet.getDataRange().getValues();
  if (!splData || splData.length <= 1) {
    return { success: false, message: 'Belum ada data pengajuan SPL di database.' };
  }

  const splHeaders = splData[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = splHeaders.findIndex(function (h) { return h.includes('id spl') || h === 'id'; });
  const empIdIdx = splHeaders.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h === 'username'; });
  const namaIdx = splHeaders.findIndex(function (h) { return h.includes('nama karyawan') || h === 'nama'; });
  const deptIdx = splHeaders.findIndex(function (h) { return h.includes('departemen') || h.includes('dept'); });
  const jabIdx = splHeaders.findIndex(function (h) { return h.includes('jabatan'); });
  const tglLemburIdx = splHeaders.findIndex(function (h) { return h === 'tanggal lembur' || h.includes('tanggal lembur') || h.includes('tgl lembur') || h === 'tanggal'; });
  const totalJamIdx = splHeaders.findIndex(function (h) { return h.includes('total jam'); });
  const uraianIdx = splHeaders.findIndex(function (h) { return h.includes('uraian') || h.includes('pekerjaan'); });

  let statusIdx = splHeaders.findIndex(function (h) { return h === 'status' || h === 'status spl' || h === 'status pengajuan'; });
  if (statusIdx === -1) {
    statusIdx = splHeaders.findIndex(function (h) { return h.includes('status') && !h.includes('approver') && !h.includes('app'); });
  }
  if (statusIdx === -1) {
    statusIdx = splHeaders.findIndex(function (h) { return h.includes('status'); });
  }
  const app1StatusIdx = splHeaders.findIndex(function (h) { return h.includes('approver 1 status') || h.includes('app 1 status') || h.includes('status atasan'); });
  const app2StatusIdx = splHeaders.findIndex(function (h) { return h.includes('approver 2 status') || h.includes('app 2 status') || h.includes('status hrd'); });

  // Filter SPL yang disetujui dalam rentang tanggal periode
  const approvedSPLs = [];
  for (let i = 1; i < splData.length; i++) {
    const row = splData[i];
    const rowStatus = statusIdx !== -1 ? String(row[statusIdx] || '').trim() : '';
    const rowApp1 = app1StatusIdx !== -1 ? String(row[app1StatusIdx] || '').trim() : '';
    const rowApp2 = app2StatusIdx !== -1 ? String(row[app2StatusIdx] || '').trim() : '';

    if (!isSPLApprovedStatus(rowStatus, rowApp1, rowApp2)) continue;

    const tglVal = tglLemburIdx !== -1 ? row[tglLemburIdx] : row[6];
    const rowDateStr = normalizeDateYMD(tglVal, ss);

    if (sDate && rowDateStr < sDate) continue;
    if (eDate && rowDateStr > eDate) continue;

    const sId = idIdx !== -1 ? String(row[idIdx] || '').trim() : ('SPL-' + i);
    const eId = empIdIdx !== -1 ? String(row[empIdIdx] || '').trim() : '';
    const eName = namaIdx !== -1 ? String(row[namaIdx] || '').trim() : '';
    const dept = deptIdx !== -1 ? String(row[deptIdx] || '').trim() : '';
    const jab = jabIdx !== -1 ? String(row[jabIdx] || '').trim() : '';

    let rawHours = totalJamIdx !== -1 ? row[totalJamIdx] : 0;
    let hours = 0;
    if (typeof rawHours === 'number') {
      hours = isNaN(rawHours) ? 0 : rawHours;
    } else {
      const cleanStr = String(rawHours || '').replace(',', '.').replace(/[^0-9.-]/g, '');
      hours = parseFloat(cleanStr) || 0;
    }
    const uraian = uraianIdx !== -1 ? String(row[uraianIdx] || '').trim() : '';

    if (hours > 0 && (eId || eName)) {
      approvedSPLs.push({
        splId: sId,
        empId: eId,
        empName: eName,
        dept: dept,
        jabatan: jab,
        tanggal: rowDateStr,
        jamLembur: hours,
        uraian: uraian
      });
    }
  }

  if (approvedSPLs.length === 0) {
    return {
      success: true,
      count: 0,
      message: 'Tidak ditemukan data SPL dengan status "Disetujui" pada rentang periode ini (' + (sDate || '-') + ' s/d ' + (eDate || '-') + ').'
    };
  }

  // Buka Sheet DataGaji
  const gajiSheet = getOrCreateDataGajiSheet(ss);
  const gajiData = gajiSheet.getDataRange().getValues();
  const gajiHeaders = gajiData[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const rincianLogColIdx = gajiHeaders.findIndex(function (h) { return h.includes('rincian'); });
  const jsonColIdx = rincianLogColIdx !== -1 ? rincianLogColIdx : 14;

  let updatedEmployeesCount = 0;
  let totalSPLSynced = 0;
  let totalHoursSynced = 0;

  // Loop setiap baris karyawan di DataGaji untuk periode ini
  for (let i = 1; i < gajiData.length; i++) {
    const rowPeriod = String(gajiData[i][0] || '').trim();
    if (rowPeriod !== targetPeriodId) continue;

    const rowEmpId = String(gajiData[i][1] || '').trim();
    const rowEmpName = String(gajiData[i][2] || '').trim();

    // Cari semua SPL yang cocok untuk karyawan ini (menggunakan isSPLMatchEmployee)
    const empSPLs = approvedSPLs.filter(function (spl) {
      return isSPLMatchEmployee(spl.empId, spl.empName, rowEmpId, rowEmpName, [rowEmpId.toLowerCase()]);
    });

    if (!empSPLs || empSPLs.length === 0) continue;

    let existingLogs = [];
    const rawJson = String(gajiData[i][jsonColIdx] || '').trim();
    try {
      if (rawJson && rawJson.startsWith('[')) {
        existingLogs = JSON.parse(rawJson);
      }
    } catch (e) {
      existingLogs = [];
    }

    let basic = cleanSalaryNum(gajiData[i][4]);

    // Untuk setiap SPL karyawan ini, update atau tambahkan log harian
    empSPLs.forEach(function (spl) {
      totalSPLSynced++;
      totalHoursSynced += spl.jamLembur;

      let targetLog = existingLogs.find(function (l) {
        return normalizeDateYMD(l.tanggal, ss) === spl.tanggal;
      });

      if (targetLog) {
        targetLog.jamSPL = spl.jamLembur;
        if (!targetLog.tarifSPL || targetLog.tarifSPL === 0) {
          targetLog.tarifSPL = defaultTarif;
        }
        targetLog.totalSPL = targetLog.jamSPL * targetLog.tarifSPL;
        targetLog.splId = spl.splId;
        const bVal = targetLog.basic !== undefined ? targetLog.basic : basic;
        const otSub = Number(targetLog.totalOvertime || 0);
        const ritSub = Number(targetLog.totalRitase || 0);
        const tunjSub = Number(targetLog.tunjanganHarian || 0);
        targetLog.totalGaji = bVal + otSub + targetLog.totalSPL + ritSub + tunjSub;
        if (!String(targetLog.catatan || '').includes(spl.splId)) {
          targetLog.catatan = (targetLog.catatan ? targetLog.catatan + ' ' : '') + '[SPL ' + spl.splId + ']';
        }
      } else {
        const logId = 'GJ-' + spl.tanggal.replace(/-/g, '') + '-' + Utilities.getUuid().slice(0, 6).toUpperCase();
        const subtotalSPL = spl.jamLembur * defaultTarif;
        existingLogs.push({
          id: logId,
          periodId: targetPeriodId,
          tanggal: spl.tanggal,
          dept: spl.dept || String(gajiData[i][3] || ''),
          empId: rowEmpId,
          empName: rowEmpName,
          jabatan: spl.jabatan,
          basic: basic,
          jamOvertime: 0,
          tarifOvertime: 0,
          totalOvertime: 0,
          jamSPL: spl.jamLembur,
          tarifSPL: defaultTarif,
          totalSPL: subtotalSPL,
          splId: spl.splId,
          ritase: 0,
          tarifRitase: 0,
          totalRitase: 0,
          tunjanganHarian: 0,
          totalGaji: basic + subtotalSPL,
          catatan: 'Lembur SPL: ' + spl.splId + (spl.uraian ? ' (' + spl.uraian + ')' : ''),
          author: 'Sync SPL Otomatis',
          updatedAt: Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss')
        });
      }
    });

    // Urutkan riwayat berdasarkan tanggal descending
    existingLogs.sort(function (a, b) {
      return String(b.tanggal || '').localeCompare(String(a.tanggal || ''));
    });

    // Hitung ulang akumulasi total karyawan secara terpisah: Overtime & SPL
    let totalJamOt = 0;
    let totalNominalOt = 0;
    let totalJamSPL = 0;
    let totalNominalSPL = 0;
    let totalRitase = 0;
    let totalNominalRit = 0;
    let totalTunjangan = 0;
    const uniqueDates = new Set();

    existingLogs.forEach(function (l) {
      totalJamOt += Number(l.jamOvertime || 0);
      totalNominalOt += Number(l.totalOvertime || 0);
      totalJamSPL += Number(l.jamSPL || 0);
      totalNominalSPL += Number(l.totalSPL || 0);
      totalRitase += Number(l.ritase || 0);
      totalNominalRit += Number(l.totalRitase || 0);
      totalTunjangan += Number(l.tunjanganHarian || 0);
      if (l.tanggal) uniqueDates.add(String(l.tanggal).trim());
    });

    const hariKerja = uniqueDates.size || existingLogs.length;
    const totalGajiBerjalan = basic + totalNominalOt + totalNominalSPL + totalNominalRit + totalTunjangan;
    const nowStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    const jsonStr = JSON.stringify(existingLogs);

    // Update kolom 6 s/d 16 di baris karyawan tersebut
    gajiSheet.getRange(i + 1, 6, 1, 11).setValues([[
      totalJamOt,
      totalNominalOt,
      totalJamSPL,
      totalNominalSPL,
      totalRitase,
      totalNominalRit,
      totalTunjangan,
      totalGajiBerjalan,
      hariKerja,
      jsonStr,
      nowStr
    ]]);

    updatedEmployeesCount++;
  }

  formatDataGajiDataRange(gajiSheet, 1);
  SpreadsheetApp.flush();

  return {
    success: true,
    count: totalSPLSynced,
    hours: Math.round(totalHoursSynced * 100) / 100,
    employeesCount: updatedEmployeesCount,
    message: 'Berhasil menyinkronkan ' + totalSPLSynced + ' data SPL Disetujui (' + (Math.round(totalHoursSynced * 100) / 100) + ' Jam) ke ' + updatedEmployeesCount + ' karyawan pada periode ' + targetPeriodInfo.name + '.'
  };
}
/**
 * WRITE API: Simpan / Update Potongan Gaji Karyawan untuk Periode Tertentu
 */
function savePotonganGaji(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  if (userRole !== 'admin' && userRole !== 'hrd') {
    return { success: false, message: 'Akses ditolak. Hanya HRD/Admin yang dapat mengubah potongan gaji.' };
  }

  if (!payload || !payload.empId) {
    return { success: false, message: 'ID Karyawan wajib ditentukan.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allPeriodsRes = getPeriodePayrollList(token);
  const activePeriod = allPeriodsRes.success ? allPeriodsRes.activePeriod : null;
  const targetPeriodId = payload.periodId ? String(payload.periodId).trim() : (activePeriod ? activePeriod.id : '');

  // Cek apakah periode berstatus DITUTUP
  const pSheet = getOrCreatePeriodePayrollSheet(ss);
  const pData = pSheet.getDataRange().getValues();
  for (let p = 1; p < pData.length; p++) {
    if (String(pData[p][0] || '').trim() === targetPeriodId) {
      const pStatus = String(pData[p][4] || '').trim().toUpperCase();
      if (pStatus === 'DITUTUP') {
        return { success: false, message: 'Periode payroll ini sudah DITUTUP. Potongan gaji tidak dapat diubah.' };
      }
      break;
    }
  }

  const sheet = getOrCreateDataGajiSheet(ss);
  const data = sheet.getDataRange().getValues();
  const headerList = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });

  const empId = String(payload.empId).trim();
  const empName = String(payload.empName || '').trim();

  let targetRowIndex = -1;
  for (let i = 1; i < data.length; i++) {
    const rowPeriod = String(data[i][0] || '').trim();
    if (targetPeriodId && rowPeriod !== targetPeriodId) continue;

    const rowEmpId = String(data[i][1] || '').trim().toLowerCase();
    const rowEmpName = String(data[i][2] || '').trim().toLowerCase();
    if ((empId && rowEmpId === empId.toLowerCase()) || (empName && rowEmpName === empName.toLowerCase())) {
      targetRowIndex = i + 1; // 1-based index
      break;
    }
  }

  if (targetRowIndex === -1) {
    return { success: false, message: 'Data gaji karyawan untuk periode ini belum ditemukan.' };
  }

  const potKesIdx = headerList.findIndex(function (h) { return h.includes('bpjs kesehatan') || h.includes('bpjs kes'); });
  const potTkIdx = headerList.findIndex(function (h) { return h.includes('bpjs ketenagakerjaan') || h.includes('bpjs tk'); });
  const potPphIdx = headerList.findIndex(function (h) { return h.includes('pph 21') || h.includes('pph21'); });
  const kasbonIdx = headerList.findIndex(function (h) { return h.includes('kasbon') || h.includes('pinjaman'); });
  const catKasbonIdx = headerList.findIndex(function (h) { return h.includes('catatan kasbon'); });
  const potLainIdx = headerList.findIndex(function (h) { return h.includes('potongan lain') && !h.includes('catatan'); });
  const catPotLainIdx = headerList.findIndex(function (h) { return h.includes('catatan potongan lain'); });
  const totPotIdx = headerList.findIndex(function (h) { return h.includes('total potongan'); });
  const thpIdx = headerList.findIndex(function (h) { return h.includes('take home pay') || h.includes('thp'); });
  const detailPotIdx = headerList.findIndex(function (h) { return h.includes('detail potongan'); });
  const grandTotalIdx = headerList.findIndex(function (h) { return h.includes('total gaji berjalan') || h.includes('total gaji'); });

  const potKes = cleanSalaryNum(payload.potonganBPJSKes);
  const potTk = cleanSalaryNum(payload.potonganBPJSTK);
  const potPph = cleanSalaryNum(payload.potonganPPh21);
  const potKasbon = cleanSalaryNum(payload.potonganKasbon);
  const catKasbon = String(payload.catatanKasbon || '').trim();
  const potLain = cleanSalaryNum(payload.potonganLain);
  const catPotLain = String(payload.catatanPotonganLain || '').trim();

  const totalPotongan = potKes + potTk + potPph + potKasbon + potLain;
  const totalGajiBerjalan = grandTotalIdx !== -1 ? cleanSalaryNum(data[targetRowIndex - 1][grandTotalIdx]) : 0;
  const takeHomePay = Math.max(0, totalGajiBerjalan - totalPotongan);
  const detailJson = JSON.stringify(payload.detailPotongan || {});

  if (potKesIdx !== -1) sheet.getRange(targetRowIndex, potKesIdx + 1).setValue(potKes);
  if (potTkIdx !== -1) sheet.getRange(targetRowIndex, potTkIdx + 1).setValue(potTk);
  if (potPphIdx !== -1) sheet.getRange(targetRowIndex, potPphIdx + 1).setValue(potPph);
  if (kasbonIdx !== -1) sheet.getRange(targetRowIndex, kasbonIdx + 1).setValue(potKasbon);
  if (catKasbonIdx !== -1) sheet.getRange(targetRowIndex, catKasbonIdx + 1).setValue(catKasbon);
  if (potLainIdx !== -1) sheet.getRange(targetRowIndex, potLainIdx + 1).setValue(potLain);
  if (catPotLainIdx !== -1) sheet.getRange(targetRowIndex, catPotLainIdx + 1).setValue(catPotLain);
  if (totPotIdx !== -1) sheet.getRange(targetRowIndex, totPotIdx + 1).setValue(totalPotongan);
  if (thpIdx !== -1) sheet.getRange(targetRowIndex, thpIdx + 1).setValue(takeHomePay);
  if (detailPotIdx !== -1) sheet.getRange(targetRowIndex, detailPotIdx + 1).setValue(detailJson);

  // Update kolom Terakhir Diupdate
  const updatedIdx = headerList.findIndex(function (h) { return h.includes('terakhir diupdate'); });
  if (updatedIdx !== -1) {
    sheet.getRange(targetRowIndex, updatedIdx + 1).setValue(Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'));
  }

  formatDataGajiDataRange(sheet, 1);
  SpreadsheetApp.flush();

  return {
    success: true,
    message: 'Potongan gaji untuk ' + (empName || empId) + ' berhasil disimpan.',
    summary: {
      empId: empId,
      potonganBPJSKes: potKes,
      potonganBPJSTK: potTk,
      potonganPPh21: potPph,
      potonganKasbon: potKasbon,
      catatanKasbon: catKasbon,
      potonganLain: potLain,
      catatanPotonganLain: catPotLain,
      totalPotongan: totalPotongan,
      takeHomePay: takeHomePay
    }
  };
}
