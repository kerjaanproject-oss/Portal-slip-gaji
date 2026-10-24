// ==========================================================================
// FITUR PENDAFTARAN MANDIRI & PERSETUJUAN AKUN (SELF-REGISTRATION)
// ==========================================================================

function getOrCreatePendaftarAkunSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let pendaftarSheet = ss.getSheetByName('PendaftarAkun');
  if (!pendaftarSheet) {
    pendaftarSheet = ss.insertSheet('PendaftarAkun');
    pendaftarSheet.appendRow([
      'ID Pendaftaran',
      'ID Karyawan',
      'Nama Lengkap',
      'Departemen',
      'Jabatan',
      'No WhatsApp / Email',
      'Password Hash',
      'Tanggal Daftar',
      'Status',
      'Tanggal Diproses',
      'Diproses Oleh',
      'Catatan HRD'
    ]);
    pendaftarSheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#e2e8f0');
  }
  return pendaftarSheet;
}

/**
 * Public Action: Mendapatkan daftar nama departemen dari Master Data Departemen untuk dropdown form pendaftaran
 */
function getPublicDepartments() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let deptSheet = ss.getSheetByName('Departemen');
    if (!deptSheet) {
      deptSheet = ss.insertSheet('Departemen');
      deptSheet.appendRow(['ID Departemen', 'Nama Departemen', 'Keterangan', 'Tanggal Dibuat']);
      deptSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e2e8f0');
    }

    const data = deptSheet.getDataRange().getValues();
    const depts = [];
    for (let i = 1; i < data.length; i++) {
      const dName = String(data[i][1] || '').trim();
      if (dName && depts.indexOf(dName) === -1) {
        depts.push(dName);
      }
    }
    // Urutkan abjad A-Z
    depts.sort(function (a, b) { return a.localeCompare(b); });

    return {
      success: true,
      departemen: depts,
      departments: depts
    };
  } catch (err) {
    return {
      success: false,
      message: 'Gagal mengambil master data departemen: ' + err.toString(),
      departemen: [],
      departments: []
    };
  }
}

/**
 * Public Action: Pendaftaran akun baru oleh karyawan secara mandiri
 */
function registerNewUser(payload) {
  if (!payload || typeof payload !== 'object') {
    return { success: false, message: 'Data pendaftaran tidak valid.' };
  }

  const username = String(payload.id || payload.username || '').trim();
  const fullname = String(payload.name || payload.fullname || '').trim();
  const dept = String(payload.dept || payload.departemen || '').trim();
  const jabatan = String(payload.jabatan || '').trim();
  const kontak = String(payload.kontak || payload.email || payload.whatsapp || payload.noHp || '').trim();
  const password = String(payload.password || '').trim();

  if (!username) return { success: false, message: 'ID Karyawan / NIK wajib diisi.' };
  if (!fullname) return { success: false, message: 'Nama Lengkap wajib diisi.' };
  if (!dept) return { success: false, message: 'Departemen wajib dipilih.' };
  if (!jabatan) return { success: false, message: 'Jabatan wajib diisi.' };
  if (!kontak) return { success: false, message: 'No. WhatsApp / Email wajib diisi.' };
  if (!password || password.length < 6) {
    return { success: false, message: 'Password minimal 6 karakter.' };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Server sedang sibuk. Silakan coba beberapa saat lagi.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Cek apakah ID Karyawan sudah aktif di Sheet Users
    const userSheet = ss.getSheetByName('Users');
    if (userSheet) {
      const userData = userSheet.getDataRange().getValues();
      for (let i = 1; i < userData.length; i++) {
        const existingId = String(userData[i][0] || '').trim().toLowerCase();
        if (existingId === username.toLowerCase()) {
          lock.releaseLock();
          return {
            success: false,
            message: 'ID Karyawan "' + username + '" sudah terdaftar sebagai akun aktif. Silakan langsung login atau gunakan fitur Lupa Password.'
          };
        }
      }
    }

    // 2. Cek apakah ID Karyawan sudah pernah mendaftar dan statusnya masih Menunggu Persetujuan
    const pendaftarSheet = getOrCreatePendaftarAkunSheet(ss);
    const pData = pendaftarSheet.getDataRange().getValues();
    for (let p = 1; p < pData.length; p++) {
      const regUsername = String(pData[p][1] || '').trim().toLowerCase();
      const regStatus = String(pData[p][8] || '').trim();
      if (regUsername === username.toLowerCase() && regStatus === 'Menunggu Persetujuan') {
        lock.releaseLock();
        return {
          success: false,
          message: 'Pendaftaran untuk ID Karyawan "' + username + '" sudah diajukan sebelumnya dan sedang menunggu verifikasi HRD.'
        };
      }
    }

    // 3. Simpan data pendaftaran
    const regId = 'REG-' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd-HHmmss') + '-' + Math.floor(100 + Math.random() * 900);
    const passHash = hashPassword(password, username);
    const now = new Date();

    pendaftarSheet.appendRow([
      regId,
      username,
      fullname,
      dept,
      jabatan,
      kontak,
      passHash,
      now,
      'Menunggu Persetujuan',
      '',
      '',
      ''
    ]);

    lock.releaseLock();
    return {
      success: true,
      message: 'Pendaftaran berhasil dikirim! Akun Anda sedang menunggu persetujuan dari HRD. Anda akan dapat login setelah disetujui.'
    };
  } catch (err) {
    lock.releaseLock();
    return { success: false, message: 'Gagal mengirim pendaftaran: ' + err.toString() };
  }
}

/**
 * Protected Action (HRD/Admin): Mengambil daftar pendaftaran akun
 */
function getPendaftarList(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  const userRole = String(session.user.role || '').trim().toLowerCase();
  if (userRole !== 'admin' && userRole !== 'hrd') {
    return { success: false, message: 'Akses ditolak: Hanya HRD yang dapat melihat daftar pendaftar.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreatePendaftarAkunSheet(ss);
    const data = sheet.getDataRange().getValues();
    const list = [];

    for (let i = 1; i < data.length; i++) {
      const regId = String(data[i][0] || '').trim();
      if (!regId) continue;

      let tglDaftar = '-';
      if (data[i][7]) {
        try {
          tglDaftar = Utilities.formatDate(new Date(data[i][7]), Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd/MM/yyyy');
        } catch (e) {
          tglDaftar = String(data[i][7]);
        }
      }

      let tglProses = '-';
      if (data[i][9]) {
        try {
          tglProses = Utilities.formatDate(new Date(data[i][9]), Session.getScriptTimeZone() || 'Asia/Jakarta', 'dd/MM/yyyy');
        } catch (e) {
          tglProses = String(data[i][9]);
        }
      }

      list.push({
        regId: regId,
        idKaryawan: String(data[i][1] || '').trim(),
        name: String(data[i][2] || '').trim(),
        dept: String(data[i][3] || '').trim(),
        jabatan: String(data[i][4] || '').trim(),
        kontak: String(data[i][5] || '').trim(),
        tanggalDaftar: tglDaftar,
        status: String(data[i][8] || 'Menunggu Persetujuan').trim(),
        tanggalDiproses: tglProses,
        diprosesOleh: String(data[i][10] || '-').trim(),
        catatan: String(data[i][11] || '-').trim()
      });
    }

    // Urutkan pendaftaran terbaru di posisi paling atas
    list.reverse();
    return { success: true, list: list };
  } catch (err) {
    return { success: false, message: 'Gagal memuat daftar pendaftar: ' + err.toString() };
  }
}

/**
 * Protected Action (HRD/Admin): Memproses persetujuan pendaftar akun (Setujui / Tolak)
 */
function processApprovalPendaftar(token, regId, action, catatan, polaRoster) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  const userRole = String(session.user.role || '').trim().toLowerCase();
  if (userRole !== 'admin' && userRole !== 'hrd') {
    return { success: false, message: 'Akses ditolak.' };
  }

  regId = String(regId || '').trim();
  action = String(action || '').trim();
  catatan = String(catatan || '').trim();
  polaRoster = String(polaRoster || 'Non-Roster').trim();

  if (!regId) return { success: false, message: 'ID Pendaftaran tidak valid.' };
  if (action !== 'Setujui' && action !== 'Tolak') {
    return { success: false, message: 'Aksi harus berupa Setujui atau Tolak.' };
  }
  if (action === 'Tolak' && !catatan) {
    return { success: false, message: 'Alasan penolakan wajib diisi.' };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Server sedang sibuk. Coba lagi sebentar.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreatePendaftarAkunSheet(ss);
    const data = sheet.getDataRange().getValues();

    let targetRow = -1;
    let pendaftarData = null;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === regId) {
        targetRow = i + 1;
        pendaftarData = {
          regId: data[i][0],
          username: String(data[i][1] || '').trim(),
          name: String(data[i][2] || '').trim(),
          dept: String(data[i][3] || '').trim(),
          jabatan: String(data[i][4] || '').trim(),
          kontak: String(data[i][5] || '').trim(),
          passHash: String(data[i][6] || '').trim(),
          status: String(data[i][8] || '').trim()
        };
        break;
      }
    }

    if (targetRow === -1 || !pendaftarData) {
      lock.releaseLock();
      return { success: false, message: 'Data pendaftaran tidak ditemukan.' };
    }

    if (pendaftarData.status !== 'Menunggu Persetujuan') {
      lock.releaseLock();
      return { success: false, message: 'Pendaftaran ini sudah diproses sebelumnya (Status: ' + pendaftarData.status + ').' };
    }

    const now = new Date();
    const processorName = session.user.fullname || session.user.username || 'HRD';

    if (action === 'Setujui') {
      // 1. Masukkan akun ke Sheet Users
      const userSheet = ss.getSheetByName('Users');
      ensureUserSheetColumns(userSheet);
      const uData = userSheet.getDataRange().getValues();

      let alreadyExists = false;
      for (let u = 1; u < uData.length; u++) {
        if (String(uData[u][0]).trim().toLowerCase() === pendaftarData.username.toLowerCase()) {
          alreadyExists = true;
          break;
        }
      }

      if (!alreadyExists) {
        userSheet.appendRow([
          pendaftarData.username,
          pendaftarData.name,
          pendaftarData.jabatan,
          pendaftarData.dept,
          pendaftarData.passHash,
          'Karyawan',
          pendaftarData.kontak,
          now,
          polaRoster,
          Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd')
        ]);
      }

      // 2. Update status di PendaftarAkun
      sheet.getRange(targetRow, 9).setValue('Disetujui');
      sheet.getRange(targetRow, 10).setValue(now);
      sheet.getRange(targetRow, 11).setValue(processorName);
      sheet.getRange(targetRow, 12).setValue(catatan || 'Disetujui oleh HRD');

      lock.releaseLock();
      return {
        success: true,
        message: 'Pendaftaran ' + pendaftarData.name + ' (' + pendaftarData.username + ') berhasil disetujui! Akun telah aktif dan siap digunakan untuk login.'
      };
    } else {
      // Tolak
      sheet.getRange(targetRow, 9).setValue('Ditolak');
      sheet.getRange(targetRow, 10).setValue(now);
      sheet.getRange(targetRow, 11).setValue(processorName);
      sheet.getRange(targetRow, 12).setValue(catatan);

      lock.releaseLock();
      return {
        success: true,
        message: 'Pendaftaran ' + pendaftarData.name + ' (' + pendaftarData.username + ') telah ditolak.'
      };
    }
  } catch (err) {
    lock.releaseLock();
    return { success: false, message: 'Gagal memproses persetujuan: ' + err.toString() };
  }
}
