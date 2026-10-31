// ==========================================================================
// MODUL MASTER DATA KARYAWAN & DEPARTEMEN
// File: Kode_Karyawan.js
// Berisi fungsi inisialisasi data aplikasi, CRUD karyawan, CRUD departemen,
// dan impor massal karyawan.
// ==========================================================================

// --- DATA READ APIs ---
function getInitialData(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');
  const currentUsername = String(session.user.username || '').trim().toLowerCase();

  // Get Karyawan
  const userSheet = ss.getSheetByName('Users');
  const userData = userSheet ? userSheet.getDataRange().getValues() : [];
  const karyawanList = [];
  for (let i = 1; i < userData.length; i++) {
    const rowId = String(userData[i][0] || '').trim().toLowerCase();
    if (isHRD || rowId === currentUsername) {
      const tglMulai = userData[i][9] instanceof Date
        ? Utilities.formatDate(userData[i][9], 'Asia/Jakarta', 'yyyy-MM-dd')
        : String(userData[i][9] || '').trim();
      karyawanList.push({
        id: String(userData[i][0]),
        name: String(userData[i][1]),
        jabatan: String(userData[i][2]),
        dept: String(userData[i][3]),
        role: String(userData[i][5]),
        email: String(userData[i][6] || ''),
        polaRoster: String(userData[i][8] || 'Non-Roster').trim() || 'Non-Roster',
        tglMulaiKerja: tglMulai,
        hasSlipPin: Boolean(String(userData[i][10] || '').trim()),
        gajiPokok: Number(userData[i][11]) || 0,
        bpjsKesSetting: String(userData[i][12] || 'Tidak Ikut').trim(),
        bpjsKesNominal: Number(userData[i][13]) || 0,
        bpjsTKSetting: String(userData[i][14] || 'Tidak Ikut').trim(),
        bpjsTKNominal: Number(userData[i][15]) || 0,
        pph21Nominal: Number(userData[i][16]) || 0
      });
    }
  }

  // Get Slips
  const slipSheet = ss.getSheetByName('SlipGaji');
  const slipData = slipSheet ? slipSheet.getDataRange().getValues() : [];
  const slipList = [];
  if (slipData && slipData.length > 1) {
    const sHeaders = slipData[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
    const sIdIdx = sHeaders.findIndex(function (h) { return h.includes('id slip') || h === 'id' || h.includes('kode'); });
    const sEmpIdIdx = sHeaders.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h.includes('nip') || h === 'id_karyawan'; });
    const sNameIdx = sHeaders.findIndex(function (h) { return h.includes('nama'); });
    const sBulanIdx = sHeaders.findIndex(function (h) { return h.includes('bulan'); });
    const sTahunIdx = sHeaders.findIndex(function (h) { return h.includes('tahun'); });
    const sFileIdx = sHeaders.findIndex(function (h) { return h.includes('nama file') || h.includes('file'); });
    const sCatIdx = sHeaders.findIndex(function (h) { return h.includes('catatan') || h.includes('keterangan'); });
    const sUrlIdx = sHeaders.findIndex(function (h) { return h.includes('url') || h.includes('link'); });
    const sTglIdx = sHeaders.findIndex(function (h) { return h.includes('tanggal') || h.includes('tgl'); });
    const sEmailIdx = sHeaders.findIndex(function (h) { return h.includes('email') || h.includes('status'); });
    const sResiNameIdx = sHeaders.findIndex(function (h) { return h.includes('file resi') || h.includes('nama resi'); });
    const sResiUrlIdx = sHeaders.findIndex(function (h) { return h.includes('url drive resi') || h.includes('url resi'); });
    const sResiIdIdx = sHeaders.findIndex(function (h) { return h.includes('file id resi') || h.includes('id resi'); });

    for (let j = 1; j < slipData.length; j++) {
      const row = slipData[j];
      if (!row || row.every(function (c) { return String(c || '').trim() === ''; })) continue;
      const empId = sEmpIdIdx !== -1 ? String(row[sEmpIdIdx] || '').trim().toLowerCase() : String(row[1] || '').trim().toLowerCase();
      if (isHRD || empId === currentUsername) {
        const resiId = sResiIdIdx !== -1 ? String(row[sResiIdIdx] || '').trim() : String(row[13] || '').trim();
        const resiName = sResiNameIdx !== -1 ? String(row[sResiNameIdx] || '').trim() : String(row[11] || '').trim();
        const uploadRaw = sTglIdx !== -1 ? row[sTglIdx] : row[9];
        const emailStatus = sEmailIdx !== -1 ? String(row[sEmailIdx] || 'Belum') : String(row[10] || 'Belum');
        slipList.push({
          id: sIdIdx !== -1 && row[sIdIdx] ? String(row[sIdIdx]).trim() : String(row[0] || ''),
          empId: sEmpIdIdx !== -1 && row[sEmpIdIdx] ? String(row[sEmpIdIdx]).trim() : String(row[1] || ''),
          empName: sNameIdx !== -1 && row[sNameIdx] ? String(row[sNameIdx]).trim() : String(row[2] || ''),
          bulan: sBulanIdx !== -1 ? String(row[sBulanIdx] || '').trim() : String(row[3] || '').trim(),
          tahun: sTahunIdx !== -1 ? String(row[sTahunIdx] || '').trim() : String(row[4] || '').trim(),
          fileName: sFileIdx !== -1 && row[sFileIdx] ? String(row[sFileIdx]).trim() : String(row[5] || ''),
          catatan: sCatIdx !== -1 && row[sCatIdx] ? String(row[sCatIdx]).trim() : String(row[6] || ''),
          driveUrl: sUrlIdx !== -1 && row[sUrlIdx] ? String(row[sUrlIdx]).trim() : String(row[7] || ''),
          uploadDate: uploadRaw instanceof Date ? Utilities.formatDate(uploadRaw, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd HH:mm') : (uploadRaw ? String(uploadRaw) : '-'),
          downloaded: emailStatus.startsWith('Terkirim') || emailStatus === 'Diakses',
          emailStatus: emailStatus,
          resiFileName: resiName,
          resiDriveUrl: sResiUrlIdx !== -1 ? String(row[sResiUrlIdx] || '').trim() : String(row[12] || ''),
          hasResi: !!(resiId || resiName)
        });
      }
    }

    const monthMap = {
      'januari': 1, 'februari': 2, 'maret': 3, 'april': 4,
      'mei': 5, 'juni': 6, 'juli': 7, 'agustus': 8,
      'september': 9, 'oktober': 10, 'november': 11, 'desember': 12
    };
    slipList.sort(function (a, b) {
      const yA = parseInt(a.tahun, 10) || 0;
      const yB = parseInt(b.tahun, 10) || 0;
      if (yA !== yB) return yB - yA;
      const mA = monthMap[String(a.bulan || '').toLowerCase()] || 0;
      const mB = monthMap[String(b.bulan || '').toLowerCase()] || 0;
      return mB - mA;
    });
  }

  // Get Departemen
  let deptSheet = ss.getSheetByName('Departemen');
  if (!deptSheet) {
    deptSheet = ss.insertSheet('Departemen');
    deptSheet.appendRow(['ID Departemen', 'Nama Departemen', 'Keterangan', 'Tanggal Dibuat']);
    deptSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e2e8f0');
  }

  const existingDeptNames = {};
  let currentDeptData = deptSheet.getDataRange().getValues();
  for (let ed = 1; ed < currentDeptData.length; ed++) {
    const deptName = String(currentDeptData[ed][1] || '').trim().toLowerCase();
    if (deptName) existingDeptNames[deptName] = true;
  }

  if (isHRD) {
    karyawanList.forEach(function (k) {
      const deptName = String(k.dept || '').trim();
      if (deptName && !existingDeptNames[deptName.toLowerCase()]) {
        deptSheet.appendRow(['DEPT-' + Utilities.getUuid().slice(0, 8).toUpperCase(), deptName, 'Diambil dari data karyawan', new Date()]);
        existingDeptNames[deptName.toLowerCase()] = true;
      }
    });
  }

  const deptData = deptSheet.getDataRange().getValues();
  const deptList = [];
  for (let d = 1; d < deptData.length; d++) {
    deptList.push({
      id: String(deptData[d][0]),
      name: String(deptData[d][1]),
      note: String(deptData[d][2] || ''),
      createdAt: deptData[d][3] ? Utilities.formatDate(new Date(deptData[d][3]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '-'
    });
  }

  // Get Master Roster
  const rosterSheet = getOrCreateMasterRosterSheet(ss);
  const rosterList = parseRosterDataFromSheet(rosterSheet);

  // Get Cuti
  const cutiSheet = getOrCreatePengajuanCutiSheet(ss);
  const cutiList = parseCutiDataFromSheet(cutiSheet, session.user);

  // Get Pending Pendaftar Count (HRD only)
  let pendingPendaftarCount = 0;
  if (isHRD) {
    try {
      const pSheet = ss.getSheetByName('PendaftarAkun');
      if (pSheet) {
        const pData = pSheet.getDataRange().getValues();
        for (let pi = 1; pi < pData.length; pi++) {
          if (String(pData[pi][8] || '').trim() === 'Menunggu Persetujuan') {
            pendingPendaftarCount++;
          }
        }
      }
    } catch (pe) {
      console.error('Count pending pendaftar note:', pe);
    }
  }

  // Get Pengumuman
  const pengumumanSheet = getOrCreatePengumumanSheet(ss);
  const annData = pengumumanSheet ? pengumumanSheet.getDataRange().getValues() : [];
  const annList = [];
  for (let a = 1; a < annData.length; a++) {
    const row = annData[a];
    const aId = String(row[0] || '').trim();
    if (!aId) continue;
    const aStatus = String(row[6] || 'Aktif').trim();
    if (!isHRD && aStatus !== 'Aktif') continue;
    annList.push({
      id: aId,
      judul: String(row[1] || '').trim(),
      isi: String(row[2] || '').trim(),
      tipe: String(row[3] || 'Informasi').trim(),
      tanggalMulai: formatDateString(row[4]),
      tanggalSelesai: formatDateString(row[5]),
      status: aStatus,
      tanggalDibuat: formatDateString(row[7]),
      dibuatOleh: String(row[8] || '').trim(),
      rowIndex: a + 1
    });
  }
  annList.reverse();

  return {
    success: true,
    karyawan: karyawanList,
    slips: slipList,
    departemen: deptList,
    roster: rosterList,
    cuti: cutiList,
    announcements: annList,
    pendingPendaftarCount: pendingPendaftarCount
  };
}

// --- DATA WRITE APIs ---
function saveDepartemen(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  const name = String(payload.name || '').trim();
  if (!name) return { success: false, message: 'Nama departemen wajib diisi.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'Sistem sedang sibuk. Coba lagi.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Departemen');
    if (!sheet) {
      sheet = ss.insertSheet('Departemen');
      sheet.appendRow(['ID Departemen', 'Nama Departemen', 'Keterangan', 'Tanggal Dibuat']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e2e8f0');
    }

    const data = sheet.getDataRange().getValues();
    const id = String(payload.id || '').trim();
    const note = String(payload.note || '').trim();

    for (let i = 1; i < data.length; i++) {
      const rowId = String(data[i][0]).trim();
      const rowName = String(data[i][1]).trim().toLowerCase();
      if (rowName === name.toLowerCase() && rowId !== id) {
        lock.releaseLock();
        return { success: false, message: 'Nama departemen sudah terdaftar.' };
      }
    }

    if (id) {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === id) {
          sheet.getRange(i + 1, 2).setValue(name);
          sheet.getRange(i + 1, 3).setValue(note);
          lock.releaseLock();
          return { success: true };
        }
      }
      lock.releaseLock();
      return { success: false, message: 'Departemen tidak ditemukan.' };
    }

    const nextId = 'DEPT-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([nextId, name, note, new Date()]);

    lock.releaseLock();
    return { success: true };
  } catch (err) {
    lock.releaseLock();
    return { success: false, message: err.toString() };
  }
}

function deleteDepartemen(token, deptId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'Sistem sedang sibuk.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Departemen');
    if (!sheet) {
      lock.releaseLock();
      return { success: true };
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(deptId).trim()) {
        const deptName = String(data[i][1]).trim();
        const userSheet = ss.getSheetByName('Users');
        if (userSheet) {
          const userData = userSheet.getDataRange().getValues();
          for (let u = 1; u < userData.length; u++) {
            if (String(userData[u][3] || '').trim().toLowerCase() === deptName.toLowerCase()) {
              lock.releaseLock();
              return { success: false, message: 'Departemen masih dipakai oleh data karyawan dan tidak bisa dihapus.' };
            }
          }
        }
        sheet.deleteRow(i + 1);
        break;
      }
    }

    lock.releaseLock();
    return { success: true };
  } catch (err) {
    lock.releaseLock();
    return { success: false, message: err.toString() };
  }
}

function saveKaryawan(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'Sistem sedang sibuk. Coba lagi.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    ensureUserSheetColumns(userSheet);
    const data = userSheet.getDataRange().getValues();

    if (payload.isNew) {
      // Check ID duplicate
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).toLowerCase() === String(payload.id).toLowerCase()) {
          lock.releaseLock();
          return { success: false, message: 'ID Karyawan sudah terdaftar.' };
        }
      }
      userSheet.appendRow([
        payload.id,
        payload.name,
        payload.jabatan,
        payload.dept,
        hashPassword(payload.pass, payload.id),
        payload.role || 'Karyawan',
        payload.email || '',
        new Date(),
        payload.polaRoster || 'Non-Roster',
        payload.tglMulaiKerja || '',
        '', // Col 11 PIN Slip
        Number(payload.gajiPokok) || 0, // Col 12 Gaji Pokok
        payload.bpjsKesSetting || 'Tidak Ikut', // Col 13 BPJS Kes Setting
        Number(payload.bpjsKesNominal) || 0, // Col 14 BPJS Kes Nominal
        payload.bpjsTKSetting || 'Tidak Ikut', // Col 15 BPJS TK Setting
        Number(payload.bpjsTKNominal) || 0, // Col 16 BPJS TK Nominal
        Number(payload.pph21Nominal) || 0 // Col 17 PPh 21 Nominal
      ]);
    } else {
      // Edit Karyawan (ID bisa diubah)
      const oldId = String(payload.oldId || payload.id || '').trim();
      let foundIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]).trim() === oldId) {
          foundIndex = i + 1;
          break;
        }
      }
      if (foundIndex === -1) {
        lock.releaseLock();
        return { success: false, message: 'Data karyawan tidak ditemukan.' };
      }

      // Check duplicate new ID (excluding current row)
      for (let i = 1; i < data.length; i++) {
        if ((i + 1) !== foundIndex && String(data[i][0]).trim().toLowerCase() === String(payload.id).trim().toLowerCase()) {
          lock.releaseLock();
          return { success: false, message: 'ID Karyawan sudah terdaftar.' };
        }
      }

      userSheet.getRange(foundIndex, 1).setValue(payload.id);
      userSheet.getRange(foundIndex, 2).setValue(payload.name);
      userSheet.getRange(foundIndex, 3).setValue(payload.jabatan);
      userSheet.getRange(foundIndex, 4).setValue(payload.dept);
      userSheet.getRange(foundIndex, 6).setValue(payload.role);
      userSheet.getRange(foundIndex, 7).setValue(payload.email || '');
      if (payload.pass) {
        userSheet.getRange(foundIndex, 5).setValue(hashPassword(payload.pass, payload.id));
      }
      userSheet.getRange(foundIndex, 9).setValue(payload.polaRoster || 'Non-Roster');
      userSheet.getRange(foundIndex, 10).setValue(payload.tglMulaiKerja || '');
      userSheet.getRange(foundIndex, 12).setValue(Number(payload.gajiPokok) || 0);
      userSheet.getRange(foundIndex, 13).setValue(payload.bpjsKesSetting || 'Tidak Ikut');
      userSheet.getRange(foundIndex, 14).setValue(Number(payload.bpjsKesNominal) || 0);
      userSheet.getRange(foundIndex, 15).setValue(payload.bpjsTKSetting || 'Tidak Ikut');
      userSheet.getRange(foundIndex, 16).setValue(Number(payload.bpjsTKNominal) || 0);
      userSheet.getRange(foundIndex, 17).setValue(Number(payload.pph21Nominal) || 0);

      // Sync SlipGaji records (update ID & Nama Karyawan)
      const slipSheet = ss.getSheetByName('SlipGaji');
      if (slipSheet) {
        const slipData = slipSheet.getDataRange().getValues();
        for (let j = 1; j < slipData.length; j++) {
          if (String(slipData[j][1]).trim() === oldId) {
            slipSheet.getRange(j + 1, 2).setValue(payload.id);
            slipSheet.getRange(j + 1, 3).setValue(payload.name);
          }
        }
      }

      // Sync ke sheet DataGaji untuk periode aktif/terbuka agar pengaturan BPJS & Pajak langsung terupdate
      try {
        const dataGajiSheet = ss.getSheetByName('DataGaji') || ss.getSheetByName('Data Gaji') || ss.getSheetByName('Data_Gaji');
        if (dataGajiSheet && dataGajiSheet.getLastRow() > 1) {
          const dgData = dataGajiSheet.getDataRange().getValues();
          const dgHeaders = dgData[0].map(function (h) { return String(h || '').trim().toLowerCase(); });

          const pSheet = ss.getSheetByName('PeriodePayroll');
          const openPeriodIds = new Set();
          if (pSheet && pSheet.getLastRow() > 1) {
            const pData = pSheet.getDataRange().getValues();
            for (let p = 1; p < pData.length; p++) {
              const pStatus = String(pData[p][4] || '').trim().toUpperCase();
              if (pStatus === 'TERBUKA') {
                openPeriodIds.add(String(pData[p][0] || '').trim());
              }
            }
          }

          const potKesIdx = dgHeaders.findIndex(function (h) { return h.includes('bpjs kesehatan') || h.includes('bpjs kes'); });
          const potTkIdx = dgHeaders.findIndex(function (h) { return h.includes('bpjs ketenagakerjaan') || h.includes('bpjs tk'); });
          const potPphIdx = dgHeaders.findIndex(function (h) { return h.includes('pph 21') || h.includes('pph21'); });
          const basicIdx = dgHeaders.findIndex(function (h) { return h === 'gaji pokok'; });
          const deptJabIdx = dgHeaders.findIndex(function (h) { return h.includes('departemen & jabatan') || h.includes('departemen'); });
          const empIdIdx = 1;   // Col 2
          const empNameIdx = 2; // Col 3
          const kasbonIdx = dgHeaders.findIndex(function (h) { return h.includes('kasbon') || h.includes('pinjaman'); });
          const potLainIdx = dgHeaders.findIndex(function (h) { return h.includes('potongan lain') && !h.includes('catatan'); });
          const totPotIdx = dgHeaders.findIndex(function (h) { return h.includes('total potongan'); });
          const thpIdx = dgHeaders.findIndex(function (h) { return h.includes('take home pay') || h.includes('thp'); });
          const detailPotIdx = dgHeaders.findIndex(function (h) { return h.includes('detail potongan'); });
          const grandTotalIdx = dgHeaders.findIndex(function (h) { return h.includes('total gaji berjalan') || h.includes('total gaji'); });

          const newBasic = Number(payload.gajiPokok) || 0;
          const newDeptJabatan = payload.dept ? (payload.dept + (payload.jabatan ? ' - ' + payload.jabatan : '')) : payload.jabatan;

          const kesSet = String(payload.bpjsKesSetting || 'Tidak Ikut').trim();
          let newPotKes = (kesSet === 'Tidak Ikut') ? 0 : (Number(payload.bpjsKesNominal) || 0);

          const tkSet = String(payload.bpjsTKSetting || 'Tidak Ikut').trim();
          let newPotTk = (tkSet === 'Tidak Ikut') ? 0 : (Number(payload.bpjsTKNominal) || 0);

          let newPotPph = Number(payload.pph21Nominal) || 0;

          const targetEmpIdLower = String(payload.id).trim().toLowerCase();
          const oldEmpIdLower = String(oldId).trim().toLowerCase();

          for (let r = 1; r < dgData.length; r++) {
            const rowPeriod = String(dgData[r][0] || '').trim();
            // Hanya perbarui periode yang masih terbuka (jika ada daftar periode terbuka)
            if (openPeriodIds.size > 0 && !openPeriodIds.has(rowPeriod)) continue;

            const rEmpId = String(dgData[r][1] || '').trim().toLowerCase();
            if (rEmpId === targetEmpIdLower || rEmpId === oldEmpIdLower) {
              const rowNum = r + 1; // 1-based index

              // Perbarui ID, Nama, Dept & Jabatan, Gaji Pokok
              if (empIdIdx !== -1) dataGajiSheet.getRange(rowNum, empIdIdx + 1).setValue(payload.id);
              if (empNameIdx !== -1) dataGajiSheet.getRange(rowNum, empNameIdx + 1).setValue(payload.name);
              if (deptJabIdx !== -1) dataGajiSheet.getRange(rowNum, deptJabIdx + 1).setValue(newDeptJabatan);
              if (basicIdx !== -1) dataGajiSheet.getRange(rowNum, basicIdx + 1).setValue(newBasic);

              // Perbarui Potongan BPJS & Pajak PPh 21
              if (potKesIdx !== -1) dataGajiSheet.getRange(rowNum, potKesIdx + 1).setValue(newPotKes);
              if (potTkIdx !== -1) dataGajiSheet.getRange(rowNum, potTkIdx + 1).setValue(newPotTk);
              if (potPphIdx !== -1) dataGajiSheet.getRange(rowNum, potPphIdx + 1).setValue(newPotPph);

              // Hitung ulang Total Potongan & THP
              const curKasbon = kasbonIdx !== -1 ? (Number(dgData[r][kasbonIdx]) || 0) : 0;
              const curPotLain = potLainIdx !== -1 ? (Number(dgData[r][potLainIdx]) || 0) : 0;
              const curTotalPot = newPotKes + newPotTk + newPotPph + curKasbon + curPotLain;
              if (totPotIdx !== -1) dataGajiSheet.getRange(rowNum, totPotIdx + 1).setValue(curTotalPot);

              const curGrandTotal = grandTotalIdx !== -1 ? (Number(dgData[r][grandTotalIdx]) || newBasic) : newBasic;
              const curTHP = Math.max(0, curGrandTotal - curTotalPot);
              if (thpIdx !== -1) dataGajiSheet.getRange(rowNum, thpIdx + 1).setValue(curTHP);

              // Perbarui Detail Potongan JSON
              let curDetail = {};
              if (detailPotIdx !== -1 && dgData[r][detailPotIdx]) {
                try { curDetail = JSON.parse(dgData[r][detailPotIdx]); } catch (e) { curDetail = {}; }
              }
              curDetail.bpjsKesSetting = kesSet;
              curDetail.bpjsKesNominal = newPotKes;
              curDetail.bpjsTKSetting = tkSet;
              curDetail.bpjsTKNominal = newPotTk;
              curDetail.pph21Nominal = newPotPph;
              if (detailPotIdx !== -1) dataGajiSheet.getRange(rowNum, detailPotIdx + 1).setValue(JSON.stringify(curDetail));
            }
          }
        }
      } catch (errDg) {
        console.warn('Sync to DataGaji failed in saveKaryawan:', errDg);
      }
    }

    lock.releaseLock();
    return { success: true };
  } catch (err) {
    lock.releaseLock();
    return { success: false, message: err.toString() };
  }
}

function deleteKaryawan(token, empId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'Sistem sedang sibuk.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    const data = userSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(empId)) {
        userSheet.deleteRow(i + 1);
        break;
      }
    }

    lock.releaseLock();
    return { success: true };
  } catch (err) {
    lock.releaseLock();
    return { success: false, message: err.toString() };
  }
}

function importKaryawanBatch(token, karyawanList) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak. Hanya Admin/HRD yang dapat mengimpor data.' };

  if (!Array.isArray(karyawanList) || karyawanList.length === 0) {
    return { success: false, message: 'Tidak ada data karyawan yang diimpor.' };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { success: false, message: 'Sistem sedang sibuk. Silakan coba lagi nanti.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    const existingData = userSheet.getDataRange().getValues();

    const existingMap = {};
    for (let i = 1; i < existingData.length; i++) {
      existingMap[String(existingData[i][0]).trim()] = i + 1;
    }

    let successCount = 0;

    karyawanList.forEach(function (k) {
      const empId = String(k.id || k.empId || k['ID Karyawan'] || k['ID'] || '').trim();
      const empName = String(k.name || k.empName || k['Nama Lengkap'] || k['Nama'] || '').trim();
      if (!empId || !empName) return;

      const jabatan = String(k.jabatan || k['Jabatan'] || '').trim();
      const dept = String(k.dept || k['Departemen'] || '').trim();
      const rawPass = String(k.pass || k.password || k['Password'] || '123456').trim();
      const hashedPass = hashPassword(rawPass, empId);
      const roleStr = String(k.role || k['Role'] || '').trim().toLowerCase();
      const role = (roleStr === 'admin' || roleStr === 'hrd') ? 'Admin' : 'Karyawan';
      const rawGaji = k.gajiPokok || k['Gaji Pokok'] || k['Basic'] || k['Gaji'] || 0;
      const gajiPokok = Number(String(rawGaji).replace(/[^\d]/g, '')) || 0;

      const existingRow = existingMap[empId];

      if (existingRow) {
        userSheet.getRange(existingRow, 2).setValue(empName);
        userSheet.getRange(existingRow, 3).setValue(jabatan);
        userSheet.getRange(existingRow, 4).setValue(dept);
        if (rawPass) {
          userSheet.getRange(existingRow, 5).setValue(hashedPass);
        }
        userSheet.getRange(existingRow, 6).setValue(role);
        userSheet.getRange(existingRow, 12).setValue(gajiPokok);
      } else {
        userSheet.appendRow([empId, empName, jabatan, dept, hashedPass, role, '', new Date(), 'Non-Roster', '', '', gajiPokok]);
        existingMap[empId] = userSheet.getLastRow();
      }
      successCount++;
    });

    lock.releaseLock();
    return { success: true, count: successCount, message: 'Berhasil mengimpor ' + successCount + ' data karyawan.' };
  } catch (err) {
    lock.releaseLock();
    return { success: false, message: err.toString() };
  }
}
