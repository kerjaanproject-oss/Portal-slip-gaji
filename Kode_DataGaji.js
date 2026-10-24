// ==========================================================================
// MODUL INPUT & PENGELOLAAN DATA GAJI KARYAWAN (INPUT HARIAN & TOTAL GAJI)
// File: Kode_DataGaji.js
// Berisi fungsi inisialisasi sheet DataGaji, CRUD data gaji harian, dan kalkulasi
// ==========================================================================

function getOrCreateDataGajiSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('DataGaji') || ss.getSheetByName('Data Gaji') || ss.getSheetByName('Data_Gaji');
  if (!sheet) {
    sheet = ss.insertSheet('DataGaji');
  }
  if (sheet.getLastRow() === 0) {
    const headers = [
      'ID Gaji',
      'Tanggal Kerja',
      'Departemen',
      'ID Karyawan',
      'Nama Karyawan',
      'Jabatan',
      'Gaji Pokok',
      'Jam Overtime',
      'Tarif Overtime',
      'Total Overtime',
      'Jumlah Ritase',
      'Tarif Ritase',
      'Total Ritase',
      'Tunjangan Harian',
      'Total Gaji',
      'Catatan',
      'Diinput Oleh'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getDataGaji(token, filterDept, filterTanggal) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');
  if (!isHRD) {
    return { success: false, message: 'Akses ditolak: Hanya Admin/HRD yang dapat melihat data gaji.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.flush();
  const sheet = getOrCreateDataGajiSheet(ss);
  const data = sheet.getDataRange().getValues();
  const list = [];
  if (!data || data.length <= 1) {
    return { success: true, list: [] };
  }

  function cleanNum(val) {
    if (typeof val === 'number') return isNaN(val) ? 0 : val;
    if (!val) return 0;
    var s = String(val).replace(/[^0-9.-]/g, '');
    var n = parseFloat(s);
    return isNaN(n) ? 0 : n;
  }

  const headers = data[0].map(function (h) {
    return String(h || '').trim().toLowerCase().replace(/[\s_\-]+/g, ' ');
  });

  function findCol(keywords, fallbackIndex) {
    var idx = headers.findIndex(function (h) {
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

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(function (cell) { return String(cell || '').trim() === ''; })) continue;

    const rowDept = deptIdx !== -1 && row[deptIdx] ? String(row[deptIdx]).trim() : String(row[2] || '').trim();
    if (filterDept && filterDept !== 'SEMUA' && rowDept.toLowerCase() !== String(filterDept).trim().toLowerCase()) continue;

    const rawTgl = tglIdx !== -1 && row[tglIdx] ? row[tglIdx] : row[1];
    let tglStr = '';
    if (rawTgl instanceof Date) {
      tglStr = Utilities.formatDate(rawTgl, 'Asia/Jakarta', 'yyyy-MM-dd');
    } else {
      tglStr = String(rawTgl || '').trim();
      if (tglStr.length > 10 && tglStr.includes('T')) {
        tglStr = tglStr.split('T')[0];
      }
    }

    if (filterTanggal && filterTanggal !== 'SEMUA' && tglStr !== String(filterTanggal).trim()) continue;

    const rowEmpId = empIdIdx !== -1 && row[empIdIdx] ? String(row[empIdIdx]).trim() : String(row[3] || '').trim();
    const rowEmpName = nameIdx !== -1 && row[nameIdx] ? String(row[nameIdx]).trim() : String(row[4] || '').trim();
    const rowJabatan = jabIdx !== -1 && row[jabIdx] ? String(row[jabIdx]).trim() : String(row[5] || '').trim();

    const basic = cleanNum(basicIdx !== -1 && row[basicIdx] !== undefined ? row[basicIdx] : row[6]);
    const jamOt = cleanNum(jamOtIdx !== -1 && row[jamOtIdx] !== undefined ? row[jamOtIdx] : row[7]);
    const tarifOt = cleanNum(tarifOtIdx !== -1 && row[tarifOtIdx] !== undefined ? row[tarifOtIdx] : row[8]);
    const totalOtRaw = cleanNum(totalOtIdx !== -1 && row[totalOtIdx] !== undefined ? row[totalOtIdx] : row[9]);
    const totalOt = totalOtRaw > 0 ? totalOtRaw : (jamOt * tarifOt);

    const ritase = cleanNum(ritIdx !== -1 && row[ritIdx] !== undefined ? row[ritIdx] : row[10]);
    const tarifRit = cleanNum(tarifRitIdx !== -1 && row[tarifRitIdx] !== undefined ? row[tarifRitIdx] : row[11]);
    const totalRitRaw = cleanNum(totalRitIdx !== -1 && row[totalRitIdx] !== undefined ? row[totalRitIdx] : row[12]);
    const totalRit = totalRitRaw > 0 ? totalRitRaw : (ritase * tarifRit);

    const tunjangan = cleanNum(tunjIdx !== -1 && row[tunjIdx] !== undefined ? row[tunjIdx] : row[13]);
    const totalGajiRaw = cleanNum(totalGajiIdx !== -1 && row[totalGajiIdx] !== undefined ? row[totalGajiIdx] : row[14]);
    const totalGaji = totalGajiRaw > 0 ? totalGajiRaw : (basic + totalOt + totalRit + tunjangan);

    const catatan = noteIdx !== -1 && row[noteIdx] ? String(row[noteIdx]).trim() : String(row[15] || '').trim();
    const author = authorIdx !== -1 && row[authorIdx] ? String(row[authorIdx]).trim() : String(row[16] || '').trim();
    const rowId = idIdx !== -1 && row[idIdx] ? String(row[idIdx]).trim() : (String(row[0] || '').trim() || ('GJ-' + i));

    list.push({
      id: rowId,
      tanggal: tglStr,
      dept: rowDept,
      empId: rowEmpId,
      empName: rowEmpName,
      jabatan: rowJabatan,
      basic: basic,
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

  // Urutkan tanggal terbaru di paling atas
  list.sort(function (a, b) {
    if (a.tanggal && b.tanggal) {
      return b.tanggal.localeCompare(a.tanggal);
    }
    return 0;
  });

  return { success: true, list: list };
}

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

  if (!tanggal || !dept || !empId || !empName) {
    return { success: false, message: 'Mohon lengkapi Tanggal Kerja, Departemen, dan Nama Karyawan.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Basic salary diambil dari payload atau dicari dari master Data Karyawan (sheet Users col 12)
  let basic = Math.max(0, Number(payload.basic || 0));
  if (basic === 0) {
    try {
      const userSheet = ss.getSheetByName('Users');
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let u = 1; u < uData.length; u++) {
          if (String(uData[u][0] || '').trim().toLowerCase() === empId.toLowerCase()) {
            basic = Number(uData[u][11] || 0);
            break;
          }
        }
      }
    } catch (ue) {
      console.warn('Lookup basic error:', ue);
    }
  }

  // Komponen numerik harian (default 0 jika kosong)
  const jamOt = Math.max(0, Number(payload.jamOvertime || 0));
  const tarifOt = Math.max(0, Number(payload.tarifOvertime || 0));
  const totalOt = jamOt * tarifOt;

  const ritase = Math.max(0, Number(payload.ritase || 0));
  const tarifRit = Math.max(0, Number(payload.tarifRitase || 0));
  const totalRit = ritase * tarifRit;

  const tunjangan = Math.max(0, Number(payload.tunjanganHarian || 0));

  // TOTAL GAJI = Basic + Overtime + Ritase + Tunjangan
  const totalGaji = basic + totalOt + totalRit + tunjangan;
  const catatan = String(payload.catatan || '').trim();
  const author = String(session.user.username || 'Admin').trim();

  const sheet = getOrCreateDataGajiSheet(ss);
  const data = sheet.getDataRange().getValues();

  // Cek apakah mode Edit
  const editId = String(payload.id || '').trim();
  let existingRowIndex = -1;

  if (editId) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === editId) {
        existingRowIndex = i + 1; // 1-based index
        break;
      }
    }
  }

  const rowValues = [
    editId || '',
    tanggal,
    dept,
    empId,
    empName,
    jabatan,
    basic,
    jamOt,
    tarifOt,
    totalOt,
    ritase,
    tarifRit,
    totalRit,
    tunjangan,
    totalGaji,
    catatan,
    author
  ];

  const savedItem = {
    id: existingRowIndex > 0 ? editId : '',
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
    totalGaji: totalGaji,
    catatan: catatan,
    author: author
  };

  if (existingRowIndex > 0) {
    rowValues[0] = editId;
    sheet.getRange(existingRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    SpreadsheetApp.flush();
    return { success: true, message: 'Data gaji berhasil diperbarui.', id: editId, item: savedItem };
  } else {
    const newId = 'GJ-' + tanggal.replace(/-/g, '') + '-' + Utilities.getUuid().slice(0, 6).toUpperCase();
    rowValues[0] = newId;
    sheet.appendRow(rowValues);
    SpreadsheetApp.flush();
    savedItem.id = newId;
    return { success: true, message: 'Data gaji berhasil disimpan.', id: newId, item: savedItem };
  }
}

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
    if (String(data[i][0] || '').trim() === cleanId) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return { success: true, message: 'Data gaji berhasil dihapus.' };
    }
  }

  return { success: false, message: 'Data gaji dengan ID tersebut tidak ditemukan.' };
}
