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
  const sheet = getOrCreateDataGajiSheet(ss);
  const data = sheet.getDataRange().getValues();
  const list = [];
  if (!data || data.length <= 1) {
    return { success: true, list: [] };
  }

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id gaji') || h === 'id'; });
  const tglIdx = headers.findIndex(function (h) { return h.includes('tanggal') || h.includes('tgl'); });
  const deptIdx = headers.findIndex(function (h) { return h.includes('departemen') || h.includes('dept'); });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h.includes('nip'); });
  const nameIdx = headers.findIndex(function (h) { return h.includes('nama'); });
  const jabIdx = headers.findIndex(function (h) { return h.includes('jabatan'); });
  const basicIdx = headers.findIndex(function (h) { return h.includes('gaji pokok') || h.includes('basic'); });
  const jamOtIdx = headers.findIndex(function (h) { return h.includes('jam overtime') || h.includes('jam lembur'); });
  const tarifOtIdx = headers.findIndex(function (h) { return h.includes('tarif overtime') || h.includes('tarif lembur'); });
  const totalOtIdx = headers.findIndex(function (h) { return h.includes('total overtime') || h.includes('total lembur'); });
  const ritIdx = headers.findIndex(function (h) { return h.includes('jumlah ritase') || h.includes('ritase'); });
  const tarifRitIdx = headers.findIndex(function (h) { return h.includes('tarif ritase') || h.includes('harga ritase'); });
  const totalRitIdx = headers.findIndex(function (h) { return h.includes('total ritase'); });
  const tunjIdx = headers.findIndex(function (h) { return h.includes('tunjangan harian') || h.includes('tunjangan'); });
  const totalGajiIdx = headers.findIndex(function (h) { return h.includes('total gaji'); });
  const noteIdx = headers.findIndex(function (h) { return h.includes('catatan') || h.includes('keterangan'); });
  const authorIdx = headers.findIndex(function (h) { return h.includes('diinput') || h.includes('oleh'); });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(function (cell) { return String(cell || '').trim() === ''; })) continue;

    const rowDept = deptIdx !== -1 ? String(row[deptIdx] || '').trim() : String(row[2] || '').trim();
    if (filterDept && filterDept !== 'SEMUA' && rowDept.toLowerCase() !== String(filterDept).trim().toLowerCase()) continue;

    const rawTgl = tglIdx !== -1 ? row[tglIdx] : row[1];
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

    const basic = basicIdx !== -1 ? Number(row[basicIdx] || 0) : 0;
    const jamOt = jamOtIdx !== -1 ? Number(row[jamOtIdx] || 0) : 0;
    const tarifOt = tarifOtIdx !== -1 ? Number(row[tarifOtIdx] || 0) : 0;
    const totalOt = totalOtIdx !== -1 ? Number(row[totalOtIdx] || 0) : (jamOt * tarifOt);
    const ritase = ritIdx !== -1 ? Number(row[ritIdx] || 0) : 0;
    const tarifRit = tarifRitIdx !== -1 ? Number(row[tarifRitIdx] || 0) : 0;
    const totalRit = totalRitIdx !== -1 ? Number(row[totalRitIdx] || 0) : (ritase * tarifRit);
    const tunjangan = tunjIdx !== -1 ? Number(row[tunjIdx] || 0) : 0;

    // Total Gaji: Basic + Overtime + Ritase + Tunjangan Harian
    const totalGaji = basic + totalOt + totalRit + tunjangan;

    list.push({
      id: idIdx !== -1 && row[idIdx] ? String(row[idIdx]).trim() : ('GJ-' + i),
      tanggal: tglStr,
      dept: rowDept,
      empId: empIdIdx !== -1 ? String(row[empIdIdx] || '').trim() : '',
      empName: nameIdx !== -1 ? String(row[nameIdx] || '').trim() : '',
      jabatan: jabIdx !== -1 ? String(row[jabIdx] || '').trim() : '',
      basic: basic,
      jamOvertime: jamOt,
      tarifOvertime: tarifOt,
      totalOvertime: totalOt,
      ritase: ritase,
      tarifRitase: tarifRit,
      totalRitase: totalRit,
      tunjanganHarian: tunjangan,
      totalGaji: totalGaji,
      catatan: noteIdx !== -1 ? String(row[noteIdx] || '').trim() : '',
      author: authorIdx !== -1 ? String(row[authorIdx] || '').trim() : ''
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

  if (existingRowIndex > 0) {
    rowValues[0] = editId;
    sheet.getRange(existingRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    return { success: true, message: 'Data gaji berhasil diperbarui.', id: editId };
  } else {
    const newId = 'GJ-' + tanggal.replace(/-/g, '') + '-' + Utilities.getUuid().slice(0, 6).toUpperCase();
    rowValues[0] = newId;
    sheet.appendRow(rowValues);
    return { success: true, message: 'Data gaji berhasil disimpan.', id: newId };
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
      return { success: true, message: 'Data gaji berhasil dihapus.' };
    }
  }

  return { success: false, message: 'Data gaji dengan ID tersebut tidak ditemukan.' };
}
