// ==========================================================================
// MODUL INPUT & PENGELOLAAN DATA GAJI KARYAWAN (INPUT HARIAN & REKAPITULASI)
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
      'Bulan',
      'Tahun',
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
      'Total Harian',
      'Catatan',
      'Diinput Oleh'
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getDataGaji(token, filterBulan, filterTahun, filterDept) {
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
  const bulanIdx = headers.findIndex(function (h) { return h.includes('bulan'); });
  const tahunIdx = headers.findIndex(function (h) { return h.includes('tahun'); });
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
  const totalHarianIdx = headers.findIndex(function (h) { return h.includes('total harian') || h.includes('total gaji'); });
  const noteIdx = headers.findIndex(function (h) { return h.includes('catatan') || h.includes('keterangan'); });
  const authorIdx = headers.findIndex(function (h) { return h.includes('diinput') || h.includes('oleh'); });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(function (cell) { return String(cell || '').trim() === ''; })) continue;

    const rowBulan = bulanIdx !== -1 ? String(row[bulanIdx] || '').trim() : String(row[2] || '').trim();
    const rowTahun = tahunIdx !== -1 ? String(row[tahunIdx] || '').trim() : String(row[3] || '').trim();
    const rowDept = deptIdx !== -1 ? String(row[deptIdx] || '').trim() : String(row[4] || '').trim();

    if (filterBulan && filterBulan !== 'SEMUA' && rowBulan.toLowerCase() !== String(filterBulan).trim().toLowerCase()) continue;
    if (filterTahun && filterTahun !== 'SEMUA' && rowTahun !== String(filterTahun).trim()) continue;
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

    const basic = basicIdx !== -1 ? Number(row[basicIdx] || 0) : Number(row[8] || 0);
    const jamOt = jamOtIdx !== -1 ? Number(row[jamOtIdx] || 0) : Number(row[9] || 0);
    const tarifOt = tarifOtIdx !== -1 ? Number(row[tarifOtIdx] || 0) : Number(row[10] || 0);
    const totalOt = totalOtIdx !== -1 ? Number(row[totalOtIdx] || 0) : (jamOt * tarifOt);
    const ritase = ritIdx !== -1 ? Number(row[ritIdx] || 0) : Number(row[12] || 0);
    const tarifRit = tarifRitIdx !== -1 ? Number(row[tarifRitIdx] || 0) : Number(row[13] || 0);
    const totalRit = totalRitIdx !== -1 ? Number(row[totalRitIdx] || 0) : (ritase * tarifRit);
    const tunjangan = tunjIdx !== -1 ? Number(row[tunjIdx] || 0) : Number(row[15] || 0);
    const totalHarian = totalHarianIdx !== -1 && row[totalHarianIdx] !== '' ? Number(row[totalHarianIdx]) : (totalOt + totalRit + tunjangan);

    list.push({
      id: idIdx !== -1 && row[idIdx] ? String(row[idIdx]).trim() : ('GJ-' + i),
      tanggal: tglStr,
      bulan: rowBulan,
      tahun: rowTahun,
      dept: rowDept,
      empId: empIdIdx !== -1 ? String(row[empIdIdx] || '').trim() : String(row[5] || '').trim(),
      empName: nameIdx !== -1 ? String(row[nameIdx] || '').trim() : String(row[6] || '').trim(),
      jabatan: jabIdx !== -1 ? String(row[jabIdx] || '').trim() : String(row[7] || '').trim(),
      basic: basic,
      jamOvertime: jamOt,
      tarifOvertime: tarifOt,
      totalOvertime: totalOt,
      ritase: ritase,
      tarifRitase: tarifRit,
      totalRitase: totalRit,
      tunjanganHarian: tunjangan,
      totalHarian: totalHarian,
      totalGaji: (basic + totalHarian),
      catatan: noteIdx !== -1 ? String(row[noteIdx] || '').trim() : String(row[17] || '').trim(),
      author: authorIdx !== -1 ? String(row[authorIdx] || '').trim() : String(row[18] || '').trim()
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
  let bulan = String(payload.bulan || '').trim();
  let tahun = String(payload.tahun || '').trim();

  // Jika bulan/tahun belum terisi, ekstrak dari tanggal (yyyy-MM-dd)
  if (!bulan || !tahun) {
    const parts = tanggal.split('-');
    if (parts.length >= 2) {
      tahun = parts[0];
      const monthNames = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
      const mIdx = parseInt(parts[1], 10) - 1;
      if (mIdx >= 0 && mIdx < 12) bulan = monthNames[mIdx];
    }
  }

  const dept = String(payload.dept || '').trim();
  const empId = String(payload.empId || '').trim();
  const empName = String(payload.empName || '').trim();
  const jabatan = String(payload.jabatan || '').trim();

  if (!tanggal || !bulan || !tahun || !dept || !empId || !empName) {
    return { success: false, message: 'Mohon lengkapi Tanggal, Periode, Departemen, dan Nama Karyawan.' };
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

  // Komponen numerik harian bersifat fleksibel (default 0 jika tidak ada)
  const jamOt = Math.max(0, Number(payload.jamOvertime || 0));
  const tarifOt = Math.max(0, Number(payload.tarifOvertime || 0));
  const totalOt = jamOt * tarifOt;

  const ritase = Math.max(0, Number(payload.ritase || 0));
  const tarifRit = Math.max(0, Number(payload.tarifRitase || 0));
  const totalRit = ritase * tarifRit;

  const tunjangan = Math.max(0, Number(payload.tunjanganHarian || 0));
  const totalHarian = totalOt + totalRit + tunjangan;
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

  if (existingRowIndex > 0) {
    // Mode Update
    sheet.getRange(existingRowIndex, 1, 1, 19).setValues([[
      editId,
      tanggal,
      bulan,
      tahun,
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
      totalHarian,
      catatan,
      author
    ]]);
    return { success: true, message: 'Data gaji harian berhasil diperbarui.', id: editId };
  } else {
    // Mode Tambah Baru: Generate ID Unik Harian
    const now = new Date();
    const newId = 'GJ-' + tanggal.replace(/-/g, '') + '-' + Utilities.getUuid().slice(0, 6).toUpperCase();
    sheet.appendRow([
      newId,
      tanggal,
      bulan,
      tahun,
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
      totalHarian,
      catatan,
      author
    ]);
    return { success: true, message: 'Data gaji harian berhasil disimpan.', id: newId };
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
