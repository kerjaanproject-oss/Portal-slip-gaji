// ==========================================================================
// MODUL PENGOLAHAN SLIP GAJI, FILE STORAGE & NOTIFIKASI EMAIL
// ==========================================================================

function getSlipKaryawan(token, filterBulan, filterTahun) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const currentEmpId = String(session.user.username || '').trim().toLowerCase();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const slipSheet = ss.getSheetByName('SlipGaji');
  if (!slipSheet) return { success: true, data: [], slips: [] };

  const slipData = slipSheet.getDataRange().getValues();
  const result = [];
  if (!slipData || slipData.length <= 1) return { success: true, data: [], slips: [] };

  const headers = slipData[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id slip') || h === 'id' || h.includes('kode'); });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h.includes('nip') || h === 'id_karyawan'; });
  const nameIdx = headers.findIndex(function (h) { return h.includes('nama'); });
  const bulanIdx = headers.findIndex(function (h) { return h.includes('bulan'); });
  const tahunIdx = headers.findIndex(function (h) { return h.includes('tahun'); });
  const fileIdx = headers.findIndex(function (h) { return h.includes('nama file') || h.includes('file'); });
  const catIdx = headers.findIndex(function (h) { return h.includes('catatan') || h.includes('keterangan'); });
  const urlIdx = headers.findIndex(function (h) { return h.includes('url') || h.includes('link'); });
  const fileIdIdx = headers.findIndex(function (h) { return h.includes('file id') && !h.includes('resi'); });
  const tglIdx = headers.findIndex(function (h) { return h.includes('tanggal') || h.includes('tgl'); });
  const resiNameIdx = headers.findIndex(function (h) { return h.includes('file resi') || h.includes('nama resi') || h.includes('nama file resi'); });
  const resiUrlIdx = headers.findIndex(function (h) { return h.includes('url drive resi') || h.includes('url resi'); });
  const resiIdIdx = headers.findIndex(function (h) { return h.includes('file id resi') || h.includes('id resi'); });

  for (let i = 1; i < slipData.length; i++) {
    const row = slipData[i];
    if (!row || row.every(function (cell) { return String(cell || '').trim() === ''; })) continue;

    const rowEmpId = empIdIdx !== -1 ? String(row[empIdIdx] || '').trim().toLowerCase() : String(row[1] || '').trim().toLowerCase();

    if (rowEmpId === currentEmpId) {
      const bulan = bulanIdx !== -1 ? String(row[bulanIdx] || '').trim() : String(row[3] || '').trim();
      const tahun = tahunIdx !== -1 ? String(row[tahunIdx] || '').trim() : String(row[4] || '').trim();

      const matchBulan = (!filterBulan || filterBulan === 'SEMUA' || bulan.toLowerCase() === String(filterBulan).trim().toLowerCase());
      const matchTahun = (!filterTahun || filterTahun === 'SEMUA' || tahun === String(filterTahun).trim());

      if (matchBulan && matchTahun) {
        const resiId = resiIdIdx !== -1 ? String(row[resiIdIdx] || '').trim() : String(row[13] || '').trim();
        const resiName = resiNameIdx !== -1 ? String(row[resiNameIdx] || '').trim() : String(row[11] || '').trim();
        const resiUrl = resiUrlIdx !== -1 ? String(row[resiUrlIdx] || '').trim() : String(row[12] || '').trim();
        const fileId = fileIdIdx !== -1 ? String(row[fileIdIdx] || '').trim() : String(row[8] || '').trim();
        const uploadRaw = tglIdx !== -1 ? row[tglIdx] : row[9];

        result.push({
          id: idIdx !== -1 && row[idIdx] ? String(row[idIdx]).trim() : String(row[0] || ''),
          empId: empIdIdx !== -1 && row[empIdIdx] ? String(row[empIdIdx]).trim() : String(row[1] || ''),
          empName: nameIdx !== -1 && row[nameIdx] ? String(row[nameIdx]).trim() : String(row[2] || ''),
          bulan: bulan,
          tahun: tahun,
          fileName: fileIdx !== -1 && row[fileIdx] ? String(row[fileIdx]).trim() : String(row[5] || ''),
          catatan: catIdx !== -1 && row[catIdx] ? String(row[catIdx]).trim() : String(row[6] || ''),
          driveUrl: urlIdx !== -1 && row[urlIdx] ? String(row[urlIdx]).trim() : String(row[7] || ''),
          fileId: fileId,
          uploadDate: uploadRaw instanceof Date ? Utilities.formatDate(uploadRaw, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm') : (uploadRaw ? String(uploadRaw) : '-'),
          resiFileName: resiName,
          resiDriveUrl: resiUrl,
          hasResi: !!(resiId || resiName)
        });
      }
    }
  }

  const monthMap = {
    'januari': 1, 'februari': 2, 'maret': 3, 'april': 4,
    'mei': 5, 'juni': 6, 'juli': 7, 'agustus': 8,
    'september': 9, 'oktober': 10, 'november': 11, 'desember': 12
  };
  result.sort(function (a, b) {
    const yA = parseInt(a.tahun, 10) || 0;
    const yB = parseInt(b.tahun, 10) || 0;
    if (yA !== yB) return yB - yA;
    const mA = monthMap[String(a.bulan || '').toLowerCase()] || 0;
    const mB = monthMap[String(b.bulan || '').toLowerCase()] || 0;
    return mB - mA;
  });

  return { success: true, data: result, slips: result };
}

const MAX_FILE_SIZE_BYTES = 1048576; // 1 MB
const MAX_RESI_SIZE_BYTES = 2097152; // 2 MB

function getOrCreateFolder(parent, name) {
  const safeName = String(name || '').replace(/[\\/:*?"<>|]/g, '_').trim();
  const folders = parent.getFoldersByName(safeName);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(safeName);
}

function getSlipStorageFolder(tahun, bulan, empName) {
  const rootName = 'Aplikasi_Slip_Gaji_Storage';
  let root;
  const roots = DriveApp.getFoldersByName(rootName);
  if (roots.hasNext()) {
    root = roots.next();
  } else {
    root = DriveApp.createFolder(rootName);
  }

  const yearFolder = getOrCreateFolder(root, tahun);
  const monthFolder = getOrCreateFolder(yearFolder, bulan);
  const empFolder = getOrCreateFolder(monthFolder, empName || 'Karyawan');
  return empFolder;
}

function getBase64FileSize(base64) {
  const length = String(base64 || '').length;
  if (length === 0) return 0;
  let padding = 0;
  if (base64.charAt(length - 1) === '=') padding++;
  if (base64.charAt(length - 2) === '=') padding++;
  return (length * 3) / 4 - padding;
}

function uploadSlipPdf(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: 'Proses pengunggahan sedang berjalan.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get Karyawan Name
    const userSheet = ss.getSheetByName('Users');
    const userData = userSheet.getDataRange().getValues();
    let empName = 'Karyawan';
    for (let i = 1; i < userData.length; i++) {
      if (String(userData[i][0]) === String(payload.empId)) {
        empName = userData[i][1];
        break;
      }
    }

    // Validate PDF file size (max 1 MB)
    if (!payload.fileData || getBase64FileSize(payload.fileData) > MAX_FILE_SIZE_BYTES) {
      lock.releaseLock();
      return { success: false, message: 'Ukuran file PDF melebihi batas maksimal 1 MB.' };
    }

    // Save PDF to Drive: Storage/Tahun/Bulan/Nama Karyawan
    const folder = getSlipStorageFolder(payload.tahun, payload.bulan, empName);

    const decoded = Utilities.base64Decode(payload.fileData);
    const blob = Utilities.newBlob(decoded, payload.mimeType || 'application/pdf', payload.fileName);
    const file = folder.createFile(blob);

    const slipId = 'SLIP-' + Date.now().toString().slice(-6);
    const fileUrl = file.getUrl();
    const fileId = file.getId();

    // Process Resi / Bukti Transfer if provided
    let resiFileName = '';
    let resiFileUrl = '';
    let resiFileId = '';

    if (payload.resiFileData && payload.resiFileName) {
      if (getBase64FileSize(payload.resiFileData) > MAX_RESI_SIZE_BYTES) {
        lock.releaseLock();
        return { success: false, message: 'Ukuran file resi melebihi batas maksimal 2 MB.' };
      }
      const resiDecoded = Utilities.base64Decode(payload.resiFileData);
      const resiBlob = Utilities.newBlob(resiDecoded, payload.resiMimeType || 'image/jpeg', payload.resiFileName);
      const resiFile = folder.createFile(resiBlob);
      resiFileName = payload.resiFileName;
      resiFileUrl = resiFile.getUrl();
      resiFileId = resiFile.getId();
    }

    const slipSheet = ss.getSheetByName('SlipGaji');
    // Ensure header columns if needed
    if (slipSheet.getLastRow() >= 1 && slipSheet.getLastColumn() < 14) {
      slipSheet.getRange(1, 12, 1, 3).setValues([['Nama File Resi', 'URL Drive Resi', 'File ID Resi']]);
    }

    slipSheet.appendRow([
      slipId,
      payload.empId,
      empName,
      payload.bulan,
      payload.tahun,
      payload.fileName,
      payload.catatan || '',
      fileUrl,
      fileId,
      new Date(),
      'Belum',
      resiFileName,
      resiFileUrl,
      resiFileId
    ]);

    lock.releaseLock();
    return { success: true, fileUrl: fileUrl, resiFileUrl: resiFileUrl };
  } catch (err) {
    lock.releaseLock();
    return { success: false, message: err.toString() };
  }
}

function updateSlipPdf(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { success: false, message: 'Proses pengubahan sedang berjalan.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get Karyawan Name
    const userSheet = ss.getSheetByName('Users');
    const userData = userSheet.getDataRange().getValues();
    let empName = 'Karyawan';
    for (let i = 1; i < userData.length; i++) {
      if (String(userData[i][0]) === String(payload.empId)) {
        empName = userData[i][1];
        break;
      }
    }

    const slipSheet = ss.getSheetByName('SlipGaji');
    const data = slipSheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(payload.slipId)) {
        rowIndex = i + 1;
        break;
      }
    }

    if (rowIndex === -1) {
      lock.releaseLock();
      return { success: false, message: 'Data slip tidak ditemukan.' };
    }

    const folder = getSlipStorageFolder(payload.tahun, payload.bulan, empName);

    // If new slip PDF file is uploaded
    if (payload.fileData && payload.fileName) {
      if (getBase64FileSize(payload.fileData) > MAX_FILE_SIZE_BYTES) {
        lock.releaseLock();
        return { success: false, message: 'Ukuran file PDF melebihi batas maksimal 1 MB.' };
      }

      // Trash old slip PDF file if exists
      const oldFileId = data[rowIndex - 1][8];
      if (oldFileId) {
        try {
          DriveApp.getFileById(oldFileId).setTrashed(true);
        } catch (e) {
          console.error('Drive file delete error during update:', e);
        }
      }

      const decoded = Utilities.base64Decode(payload.fileData);
      const blob = Utilities.newBlob(decoded, payload.mimeType || 'application/pdf', payload.fileName);
      const newFile = folder.createFile(blob);

      slipSheet.getRange(rowIndex, 6).setValue(payload.fileName); // Nama File
      slipSheet.getRange(rowIndex, 8).setValue(newFile.getUrl()); // URL Drive
      slipSheet.getRange(rowIndex, 9).setValue(newFile.getId()); // File ID Drive
    }

    // If new resi file is uploaded
    if (payload.resiFileData && payload.resiFileName) {
      if (getBase64FileSize(payload.resiFileData) > MAX_RESI_SIZE_BYTES) {
        lock.releaseLock();
        return { success: false, message: 'Ukuran file resi melebihi batas maksimal 2 MB.' };
      }

      // Trash old resi file if exists
      const oldResiId = data[rowIndex - 1][13];
      if (oldResiId) {
        try {
          DriveApp.getFileById(oldResiId).setTrashed(true);
        } catch (e) {
          console.error('Drive resi delete error during update:', e);
        }
      }

      const resiDecoded = Utilities.base64Decode(payload.resiFileData);
      const resiBlob = Utilities.newBlob(resiDecoded, payload.resiMimeType || 'image/jpeg', payload.resiFileName);
      const newResiFile = folder.createFile(resiBlob);

      slipSheet.getRange(rowIndex, 12).setValue(payload.resiFileName);
      slipSheet.getRange(rowIndex, 13).setValue(newResiFile.getUrl());
      slipSheet.getRange(rowIndex, 14).setValue(newResiFile.getId());
    } else if (payload.deleteResi === true) {
      // Admin chose to remove existing resi
      const oldResiId = data[rowIndex - 1][13];
      if (oldResiId) {
        try {
          DriveApp.getFileById(oldResiId).setTrashed(true);
        } catch (e) {
          console.error('Drive resi delete error during removal:', e);
        }
      }
      slipSheet.getRange(rowIndex, 12).setValue('');
      slipSheet.getRange(rowIndex, 13).setValue('');
      slipSheet.getRange(rowIndex, 14).setValue('');
    }

    // Update metadata fields
    slipSheet.getRange(rowIndex, 2).setValue(payload.empId);
    slipSheet.getRange(rowIndex, 3).setValue(empName);
    slipSheet.getRange(rowIndex, 4).setValue(payload.bulan);
    slipSheet.getRange(rowIndex, 5).setValue(payload.tahun);
    slipSheet.getRange(rowIndex, 7).setValue(payload.catatan || '');

    lock.releaseLock();
    return { success: true };
  } catch (err) {
    lock.releaseLock();
    return { success: false, message: err.toString() };
  }
}

function deleteSlipPdf(token, slipId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { success: false, message: 'Sistem sibuk.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const slipSheet = ss.getSheetByName('SlipGaji');
    const data = slipSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(slipId)) {
        const fileId = data[i][8];
        const resiFileId = data[i][13];
        if (fileId) {
          try {
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch (e) {
            console.error('Drive file delete error:', e);
          }
        }
        if (resiFileId) {
          try {
            DriveApp.getFileById(resiFileId).setTrashed(true);
          } catch (e) {
            console.error('Drive resi delete error:', e);
          }
        }
        slipSheet.deleteRow(i + 1);
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

function recordSlipAccess(token, slipId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const slipSheet = ss.getSheetByName('SlipGaji');
    const data = slipSheet.getDataRange().getValues();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(slipId)) {
        slipSheet.getRange(i + 1, 11).setValue('Diakses');
        break;
      }
    }
    return { success: true };
  } catch (e) {
    console.error('Record access error:', e);
    return { success: false, message: e.toString() };
  }
}

function getSlipPdfBase64(token, slipId, pin) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const slipSheet = ss.getSheetByName('SlipGaji');
    if (!slipSheet) return { success: false, message: 'Sheet SlipGaji tidak ditemukan.' };

    const data = slipSheet.getDataRange().getValues();
    let fileId = null;
    let empId = null;
    let empName = '';
    let bulan = '';
    let tahun = '';

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(slipId)) {
        empId = String(data[i][1]);
        empName = String(data[i][2]);
        bulan = String(data[i][3]);
        tahun = String(data[i][4]);
        fileId = data[i][8];
        break;
      }
    }

    if (!fileId) return { success: false, message: 'Data slip atau File ID tidak ditemukan.' };

    const currentUsername = String(session.user.username || '').trim().toLowerCase();
    const isOwner = currentUsername === String(empId).trim().toLowerCase();
    const isHRD = (session.user.role === 'Admin' || session.user.role === 'HRD');

    // Security check: Karyawan can only view their own slip unless Admin
    if (!isHRD && !isOwner) {
      return { success: false, message: 'Akses ditolak. Anda hanya dapat membaca slip gaji milik Anda sendiri.' };
    }

    // Jika yang mengakses adalah pemilik slip, verifikasi PIN jika akun sudah memiliki PIN
    if (isOwner) {
      const userSheet = ss.getSheetByName('Users');
      if (userSheet) {
        const uData = userSheet.getDataRange().getValues();
        for (let u = 1; u < uData.length; u++) {
          if (String(uData[u][0] || '').trim().toLowerCase() === currentUsername) {
            const userPinHash = String(uData[u][10] || '').trim();
            if (userPinHash) {
              pin = String(pin || '').trim();
              if (!pin) {
                return { success: false, needPin: true, message: 'PIN Keamanan diperlukan untuk melihat slip gaji.' };
              }
              const inputHash = hashSlipPin(pin, currentUsername);
              if (inputHash !== userPinHash) {
                recordPinFailure(currentUsername);
                return { success: false, needPin: true, message: 'PIN Keamanan salah.' };
              }
              clearPinFailures(currentUsername);
            }
            break;
          }
        }
      }
    }

    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    // Record access
    recordSlipAccess(token, slipId);

    return {
      success: true,
      base64: base64,
      fileName: file.getName(),
      empId: empId,
      empName: empName,
      bulan: bulan,
      tahun: tahun
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengambil file PDF dari Drive: ' + err.toString() };
  }
}

function getResiFileBase64(token, slipId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const slipSheet = ss.getSheetByName('SlipGaji');
    if (!slipSheet) return { success: false, message: 'Sheet SlipGaji tidak ditemukan.' };

    const data = slipSheet.getDataRange().getValues();
    let resiFileId = null;
    let empId = null;
    let empName = '';
    let bulan = '';
    let tahun = '';
    let resiFileName = '';

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(slipId)) {
        empId = String(data[i][1]);
        empName = String(data[i][2]);
        bulan = String(data[i][3]);
        tahun = String(data[i][4]);
        resiFileName = String(data[i][11] || '');
        resiFileId = data[i][13];
        break;
      }
    }

    if (!resiFileId) return { success: false, message: 'File bukti transfer / resi tidak ditemukan.' };

    // Security check: Karyawan can only view their own receipt unless Admin
    if (session.user.role !== 'Admin' && String(session.user.username).trim().toLowerCase() !== String(empId).trim().toLowerCase()) {
      return { success: false, message: 'Akses ditolak.' };
    }

    const file = DriveApp.getFileById(resiFileId);
    const blob = file.getBlob();
    const base64 = Utilities.base64Encode(blob.getBytes());

    return {
      success: true,
      base64: base64,
      mimeType: file.getMimeType(),
      fileName: file.getName() || resiFileName || 'Bukti_Transfer',
      empId: empId,
      empName: empName,
      bulan: bulan,
      tahun: tahun
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengambil file bukti transfer dari Drive: ' + err.toString() };
  }
}

function sendSlipEmailNotification(token, slipId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak. Hanya Admin/HRD yang dapat mengirimkan notifikasi email.' };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const slipSheet = ss.getSheetByName('SlipGaji');
    if (!slipSheet) return { success: false, message: 'Sheet SlipGaji tidak ditemukan.' };

    const slipData = slipSheet.getDataRange().getValues();
    let rowIndex = -1;
    let empId = '';
    let empName = '';
    let bulan = '';
    let tahun = '';
    let fileName = '';
    let fileId = '';
    let resiFileName = '';
    let resiFileId = '';

    for (let i = 1; i < slipData.length; i++) {
      if (String(slipData[i][0]).trim() === String(slipId).trim()) {
        rowIndex = i + 1;
        empId = String(slipData[i][1]).trim();
        empName = String(slipData[i][2]).trim();
        bulan = String(slipData[i][3]).trim();
        tahun = String(slipData[i][4]).trim();
        fileName = String(slipData[i][5]).trim();
        fileId = String(slipData[i][8]).trim();
        resiFileName = String(slipData[i][11] || '').trim();
        resiFileId = String(slipData[i][13] || '').trim();
        break;
      }
    }

    if (rowIndex === -1) return { success: false, message: 'Data slip gaji tidak ditemukan.' };

    // Get Employee Email from Users sheet
    const userSheet = ss.getSheetByName('Users');
    const userData = userSheet ? userSheet.getDataRange().getValues() : [];
    let recipientEmail = '';

    for (let u = 1; u < userData.length; u++) {
      if (String(userData[u][0]).trim().toLowerCase() === empId.toLowerCase()) {
        recipientEmail = String(userData[u][6] || '').trim();
        break;
      }
    }

    if (!recipientEmail || !recipientEmail.includes('@')) {
      return {
        success: false,
        message: 'Alamat email untuk karyawan ' + empName + ' (' + empId + ') tidak terdaftar atau tidak valid. Silakan perbarui data email karyawan terlebih dahulu.'
      };
    }

    if (!fileId) return { success: false, message: 'File PDF tidak ditemukan di Google Drive.' };

    // Get PDF file blob
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    if (fileName) blob.setName(fileName);

    const emailAttachments = [blob];
    if (resiFileId) {
      try {
        const resiFile = DriveApp.getFileById(resiFileId);
        const resiBlob = resiFile.getBlob();
        if (resiFileName) resiBlob.setName(resiFileName);
        emailAttachments.push(resiBlob);
      } catch (rErr) {
        console.error('Error attaching resi file to email:', rErr);
      }
    }

    const appTitle = 'e-Slip Gaji Online';
    const periodeStr = bulan + ' ' + tahun;
    // Clean subject without brackets or spammy prefixes
    const subject = 'Slip Gaji Periode ' + periodeStr + ' - ' + empName;

    // Plain text fallback (Crucial for spam filtering)
    const plainBody = 'Yth. ' + empName + ' (' + empId + '),\n\n' +
      'Slip gaji Anda untuk periode ' + periodeStr + ' telah diterbitkan.\n' +
      'Dokumen PDF slip gaji' + (resiFileId ? ' dan bukti transfer' : '') + ' terlampir langsung pada email ini.\n\n' +
      'Anda dapat menyimpan dokumen ini atau mengakses portal e-Slip Gaji Online kapan saja.\n\n' +
      'Hormat kami,\n' +
      'Tim HRD & Payroll';

    // Anti-Spam Clean HTML Body (Neutral colors, clean structure, no red boxes or emojis)
    const htmlBody = '<div style="font-family: \'Plus Jakarta Sans\', Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 8px; background-color: #ffffff;">' +
      '<div style="background-color: #10233f; padding: 20px 24px; text-align: center; border-top-left-radius: 6px; border-top-right-radius: 6px;">' +
      '<h2 style="color: #ffffff; margin: 0; font-size: 20px; font-weight: bold;">' + appTitle + '</h2>' +
      '<p style="color: #cbd5e1; margin: 4px 0 0 0; font-size: 13px;">Pemberitahuan Terbit Slip Gaji Digital</p>' +
      '</div>' +
      '<div style="padding: 24px; color: #334155;">' +
      '<p style="margin-top: 0; font-size: 14px;">Yth. Bapak/Ibu <strong>' + empName + '</strong>,</p>' +
      '<p style="line-height: 1.6; font-size: 14px;">Rincian slip gaji Anda untuk periode <strong>' + periodeStr + '</strong> telah resmi diterbitkan oleh Tim HRD & Payroll PT Berkah Batu Bara.</p>' +
      '<div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-left: 4px solid #2563eb; padding: 14px 18px; margin: 20px 0; border-radius: 6px;">' +
      '<table style="width: 100%; border-collapse: collapse; font-size: 13px;">' +
      '<tr><td style="padding: 4px 0; color: #64748b; width: 120px;">ID Karyawan</td><td style="font-weight: bold; color: #0f172a;">: ' + empId + '</td></tr>' +
      '<tr><td style="padding: 4px 0; color: #64748b;">Nama Lengkap</td><td style="font-weight: bold; color: #0f172a;">: ' + empName + '</td></tr>' +
      '<tr><td style="padding: 4px 0; color: #64748b;">Periode Gaji</td><td style="font-weight: bold; color: #2563eb;">: ' + periodeStr + '</td></tr>' +
      '<tr><td style="padding: 4px 0; color: #64748b;">Nama File</td><td style="color: #475569;">: ' + fileName + '</td></tr>' +
      (resiFileName ? '<tr><td style="padding: 4px 0; color: #64748b;">Bukti Transfer</td><td style="color: #ea580c; font-weight: 600;">: ' + resiFileName + ' (Terlampir)</td></tr>' : '') +
      '</table>' +
      '</div>' +
      '<p style="font-size: 13px; color: #475569; line-height: 1.5; background-color: #f1f5f9; border-left: 3px solid #94a3b8; padding: 12px; border-radius: 4px;">' +
      '<strong>Catatan Kerahasiaan:</strong> Dokumen PDF slip gaji' + (resiFileId ? ' dan bukti transfer' : '') + ' terlampir langsung pada email ini. Dokumen ini hanya ditujukan kepada pemilik gaji.' +
      '</p>' +
      '</div>' +
      '<div style="background-color: #f8fafc; padding: 14px; text-align: center; border-bottom-left-radius: 6px; border-bottom-right-radius: 6px; font-size: 12px; color: #64748b; border-top: 1px solid #f1f5f9;">' +
      '<p style="margin: 0;">&copy; ' + new Date().getFullYear() + ' Tim HRD & Payroll. Email otomatis e-Slip Gaji.</p>' +
      '</div>' +
      '</div>';

    // Get Active User Email for Reply-To
    let activeUserEmail = '';
    try {
      activeUserEmail = Session.getActiveUser().getEmail();
    } catch (e) { }

    const mailOptions = {
      htmlBody: htmlBody,
      name: 'HRD e-Slip Gaji',
      attachments: emailAttachments
    };
    if (activeUserEmail && activeUserEmail.includes('@')) {
      mailOptions.replyTo = activeUserEmail;
    }

    // Use GmailApp first for highest deliverability (SPF/DKIM inbox guarantee)
    try {
      GmailApp.sendEmail(recipientEmail, subject, plainBody, mailOptions);
    } catch (gErr) {
      // Fallback to MailApp if GmailApp scope is restricted
      MailApp.sendEmail({
        to: recipientEmail,
        subject: subject,
        body: plainBody,
        htmlBody: htmlBody,
        name: 'HRD e-Slip Gaji',
        replyTo: activeUserEmail || recipientEmail,
        attachments: emailAttachments
      });
    }

    // Update Email Status in Sheet (Col 11 / K)
    const sendTimeStr = 'Terkirim (' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yy HH:mm') + ')';
    slipSheet.getRange(rowIndex, 11).setValue(sendTimeStr);

    return { success: true, message: 'Email notifikasi berhasil dikirim ke ' + recipientEmail + ' (' + empName + ').' };

  } catch (err) {
    return { success: false, message: 'Gagal mengirim email: ' + err.toString() };
  }
}

function sendSlipEmailNotificationBatch(token, slipIds) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  if (!Array.isArray(slipIds) || slipIds.length === 0) {
    return { success: false, message: 'Pilih minimal satu slip gaji untuk dikirim.' };
  }

  let successCount = 0;
  let failedCount = 0;
  let errorMsgs = [];

  slipIds.forEach(function (id) {
    const res = sendSlipEmailNotification(token, id);
    if (res.success) {
      successCount++;
    } else {
      failedCount++;
      errorMsgs.push(res.message);
    }
  });

  return {
    success: true,
    successCount: successCount,
    failedCount: failedCount,
    total: slipIds.length,
    message: 'Proses selesai: ' + successCount + ' email terkirim' + (failedCount > 0 ? ', ' + failedCount + ' gagal' : '') + '.',
    errors: errorMsgs
  };
}