// --- API FUNCTION WHITELISTS ---
const ALLOWED_PUBLIC_ACTIONS = [
  'loginUser',
  'requestPasswordReset',
  'verifyAndResetPassword'
];

const ALLOWED_PROTECTED_ACTIONS = [
  'destroySession',
  'changeOwnPassword',
  'getInitialData',
  'getSlipKaryawan',
  'saveDepartemen',
  'deleteDepartemen',
  'saveKaryawan',
  'deleteKaryawan',
  'importKaryawanBatch',
  'uploadSlipPdf',
  'updateSlipPdf',
  'deleteSlipPdf',
  'recordSlipAccess',
  'getSlipPdfBase64',
  'sendSlipEmailNotification',
  'sendSlipEmailNotificationBatch'
];

function doGet(e) {
  // Hanya layani UI HTML Web App (Cegah eksekusi aksi sensitif via GET URL parameter)
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('E-Slip Gaji')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function doPost(e) {
  try {
    let postData = {};
    if (e && e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    }
    const action = String(postData.action || '').trim();
    const args = Array.isArray(postData.args) ? postData.args : [];

    // Validasi Keamanan: Pastikan hanya fungsi dalam Whitelist yang dapat dieksekusi
    const isPublic = ALLOWED_PUBLIC_ACTIONS.indexOf(action) !== -1;
    const isProtected = ALLOWED_PROTECTED_ACTIONS.indexOf(action) !== -1;

    if (!isPublic && !isProtected) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false,
        message: 'Akses ditolak: aksi tidak diizinkan atau tidak terdaftar.'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (typeof this[action] === 'function') {
      let result = this[action].apply(this, args);
      if (result === undefined) {
        result = { success: true };
      }
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } else {
      return ContentService.createTextOutput(JSON.stringify({ success: false, message: 'Action not found: ' + action }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// --- INITIAL DATABASE SETUP ---
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet Users
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['ID Karyawan', 'Nama Lengkap', 'Jabatan', 'Departemen', 'Password Hash', 'Role', 'Email', 'Tanggal Dibuat']);
    userSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#e2e8f0');

    // Default Admin Login: admin / Admin@2026
    const defaultPassHash = hashPassword('Admin@2026', 'admin');
    userSheet.appendRow(['admin', 'HRD Administrator', 'HR Manager', 'Human Resource', defaultPassHash, 'Admin', 'admin@company.com', new Date()]);
  }

  // 2. Sheet SlipGaji
  let slipSheet = ss.getSheetByName('SlipGaji');
  if (!slipSheet) {
    slipSheet = ss.insertSheet('SlipGaji');
    slipSheet.appendRow(['ID Slip', 'ID Karyawan', 'Nama Karyawan', 'Bulan', 'Tahun', 'Nama File', 'Catatan', 'URL Drive', 'File ID Drive', 'Tanggal Upload', 'Status Akses']);
    slipSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#e2e8f0');
  }

  // 3. Sheet Departemen
  let deptSheet = ss.getSheetByName('Departemen');
  if (!deptSheet) {
    deptSheet = ss.insertSheet('Departemen');
    deptSheet.appendRow(['ID Departemen', 'Nama Departemen', 'Keterangan', 'Tanggal Dibuat']);
    deptSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e2e8f0');
    deptSheet.appendRow(['DEPT-001', 'Human Resource', 'Departemen awal sistem', new Date()]);
  }

  // 4. Sheet OTP_Reset
  let otpSheet = ss.getSheetByName('OTP_Reset');
  if (!otpSheet) {
    otpSheet = ss.insertSheet('OTP_Reset');
    otpSheet.appendRow(['ID Karyawan', 'Kode OTP', 'Expired At', 'Status', 'Tanggal Dibuat']);
    otpSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e2e8f0');
  }

  return 'Setup Database Berhasil! Sheet Users, SlipGaji, Departemen, dan OTP_Reset telah siap.';
}

// --- HASH & SECURITY HELPERS ---
function hashPassword(password, username) {
  const salt = username ? String(username).trim().toLowerCase() : '';
  const combined = 'eSlip_Salt_' + salt + '_' + String(password);
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined, Utilities.Charset.UTF_8);
  return rawHash.map(function (b) {
    const hex = (b < 0 ? b + 256 : b).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

function hashLegacyPassword(password) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  return rawHash.map(function (b) {
    const hex = (b < 0 ? b + 256 : b).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

// --- SECURE HMAC-SHA256 STATELESS SESSION TOKEN ---
const SESSION_SECRET_KEY = 'eSlip_Portal_Secret_Session_Key_2026_x89q!';

function generateSessionToken(username, role, fullname) {
  const payload = {
    u: String(username),
    r: String(role),
    f: String(fullname),
    exp: Date.now() + (8 * 3600 * 1000) // Berlaku 8 Jam
  };
  const payloadStr = JSON.stringify(payload);
  const payloadB64 = Utilities.base64EncodeWebSafe(payloadStr);
  const sigBytes = Utilities.computeHmacSha256Signature(payloadB64, SESSION_SECRET_KEY);
  const sigB64 = Utilities.base64EncodeWebSafe(sigBytes);
  return payloadB64 + '.' + sigB64;
}

function validateSession(token) {
  if (!token || typeof token !== 'string') {
    return { valid: false, message: 'Token tidak ditemukan.' };
  }

  const parts = token.split('.');
  if (parts.length !== 2) {
    return { valid: false, message: 'Format sesi tidak valid. Silakan login kembali.' };
  }

  const payloadB64 = parts[0];
  const sigB64 = parts[1];

  // 1. Verifikasi Keaslian Signature HMAC-SHA256
  try {
    const expectedSigBytes = Utilities.computeHmacSha256Signature(payloadB64, SESSION_SECRET_KEY);
    const expectedSigB64 = Utilities.base64EncodeWebSafe(expectedSigBytes);
    if (sigB64 !== expectedSigB64) {
      return { valid: false, message: 'Sesi tidak valid (tanda tangan digital salah).' };
    }
  } catch (sigErr) {
    return { valid: false, message: 'Verifikasi keamanan sesi gagal: ' + sigErr.toString() };
  }

  // 2. Decode Data & Cek Masa Berlaku
  try {
    const decodedBytes = Utilities.base64DecodeWebSafe(payloadB64);
    const jsonStr = Utilities.newBlob(decodedBytes).getDataAsString();
    const payload = JSON.parse(jsonStr);

    if (!payload.exp || Date.now() > payload.exp) {
      return { valid: false, message: 'Sesi telah berakhir. Silakan login kembali.' };
    }

    return {
      valid: true,
      user: {
        username: payload.u,
        role: payload.r,
        fullname: payload.f
      }
    };
  } catch (decErr) {
    return { valid: false, message: 'Gagal memproses data sesi: ' + decErr.toString() };
  }
}

function destroySession(token) {
  return { success: true };
}

function checkLoginRateLimit(username) {
  const cache = CacheService.getScriptCache();
  const attempts = parseInt(cache.get('rate_' + username) || '0', 10);
  return attempts >= 5;
}

function recordLoginFailure(username) {
  const cache = CacheService.getScriptCache();
  const key = 'rate_' + username;
  const attempts = parseInt(cache.get(key) || '0', 10) + 1;
  cache.put(key, String(attempts), 900); // Lock 15 Minutes after 5 failures
}

function clearLoginFailures(username) {
  const cache = CacheService.getScriptCache();
  cache.remove('rate_' + username);
}

// --- AUTHENTICATION API ---
function loginUser(username, password) {
  username = String(username || '').trim();
  password = String(password || '').trim();

  if (!username || !password) {
    return { success: false, message: 'ID Karyawan dan Password wajib diisi.' };
  }

  if (checkLoginRateLimit(username)) {
    return { success: false, message: 'Terlalu banyak percobaan gagal. Akun dikunci sementara 15 menit.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { success: false, message: 'Database belum disiapkan. Jalankan setupDatabase() terlebih dahulu.' };

  const data = userSheet.getDataRange().getValues();
  const inputHashSalted = hashPassword(password, username);
  const inputHashLegacy = hashLegacyPassword(password);

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0]).trim();
    const rowPassHash = String(data[i][4]).trim();

    if (rowId.toLowerCase() === username.toLowerCase()) {
      const isMatch = (rowPassHash === inputHashSalted || rowPassHash === inputHashLegacy);
      if (isMatch) {
        clearLoginFailures(username);
        // Otomatis tingkatkan hash legacy ke salted hash saat login berhasil
        if (rowPassHash !== inputHashSalted) {
          try {
            userSheet.getRange(i + 1, 5).setValue(inputHashSalted);
          } catch (upErr) {
            console.error('Password hash upgrade note:', upErr);
          }
        }
        const userObj = {
          username: data[i][0],
          fullname: data[i][1],
          jabatan: data[i][2],
          dept: data[i][3],
          role: data[i][5]
        };
        const token = generateSessionToken(userObj.username, userObj.role, userObj.fullname);
        return { success: true, user: userObj, token: token };
      }
    }
  }

  recordLoginFailure(username);
  return { success: false, message: 'ID Karyawan atau Password salah.' };
}

// --- OTP PASSWORD RESET API ---
function requestPasswordReset(usernameOrEmail) {
  usernameOrEmail = String(usernameOrEmail || '').trim().toLowerCase();
  if (!usernameOrEmail) {
    return { success: false, message: 'Silakan masukkan ID Karyawan atau Email Anda.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

  const data = userSheet.getDataRange().getValues();
  let userFound = null;

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0] || '').trim().toLowerCase();
    const rowEmail = String(data[i][6] || '').trim().toLowerCase();

    if (rowId === usernameOrEmail || (rowEmail && rowEmail === usernameOrEmail)) {
      userFound = {
        id: data[i][0],
        fullname: data[i][1],
        email: data[i][6] || ''
      };
      break;
    }
  }

  if (!userFound) {
    return { success: false, message: 'ID Karyawan atau Email tidak terdaftar di sistem.' };
  }

  if (!userFound.email) {
    return { success: false, message: 'Akun Anda belum memiliki email terdaftar. Silakan hubungi Admin HRD.' };
  }

  // Generate 6 Digit OTP
  const otpCode = String(Math.floor(100000 + Math.random() * 900000));
  const now = new Date();
  const expiresAt = new Date(now.getTime() + 15 * 60 * 1000); // 15 Minutes

  // Get or Create OTP_Reset sheet
  let otpSheet = ss.getSheetByName('OTP_Reset');
  if (!otpSheet) {
    otpSheet = ss.insertSheet('OTP_Reset');
    otpSheet.appendRow(['ID Karyawan', 'Kode OTP', 'Expired At', 'Status', 'Tanggal Dibuat']);
    otpSheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e2e8f0');
  }

  otpSheet.appendRow([userFound.id, otpCode, expiresAt, 'ACTIVE', now]);

  // Send Email via MailApp
  try {
    const maskedEmail = userFound.email.replace(/(.{2})(.*)(?=@)/, function (gp1, gp2, gp3) {
      return gp2 + "*".repeat(gp3.length);
    });

    MailApp.sendEmail({
      to: userFound.email,
      subject: '🔒 Kode OTP Reset Password - e-Slip Gaji',
      htmlBody: `
        <div style="font-family: Arial, sans-serif; max-width: 480px; margin: 0 auto; padding: 24px; border: 1px solid #e2e8f0; border-radius: 16px; background-color: #ffffff;">
          <div style="text-align: center; margin-bottom: 20px;">
            <h2 style="color: #2563eb; margin: 0; font-size: 20px;">Portal e-Slip Gaji Online</h2>
            <p style="color: #64748b; font-size: 14px; margin-top: 4px;">Permintaan Reset Password</p>
          </div>
          <p style="font-size: 15px; color: #334155;">Halo <b>${userFound.fullname}</b>,</p>
          <p style="font-size: 14px; color: #475569; line-height: 1.5;">Anda menerima email ini karena ada permintaan untuk mereset password akun e-Slip Gaji Anda. Masukkan kode OTP berikut di aplikasi:</p>
          <div style="text-align: center; margin: 24px 0;">
            <div style="display: inline-block; font-size: 32px; font-weight: 800; letter-spacing: 8px; color: #2563eb; background: #eff6ff; border: 1px dashed #3b82f6; padding: 12px 28px; border-radius: 12px;">
              ${otpCode}
            </div>
          </div>
          <p style="font-size: 13px; color: #ef4444; margin-bottom: 20px;">⚠️ Kode OTP ini berlaku selama <b>15 menit</b>. Jangan berikan kode ini kepada siapapun.</p>
          <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 20px 0;">
          <p style="font-size: 12px; color: #94a3b8; text-align: center; margin: 0;">Jika Anda tidak merasa melakukan permintaan ini, silakan abaikan email ini.</p>
        </div>
      `
    });

    return {
      success: true,
      message: 'Kode OTP 6-digit berhasil dikirim ke ' + maskedEmail,
      username: userFound.id
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengirim email: ' + err.toString() };
  }
}

function verifyAndResetPassword(username, otpCode, newPassword) {
  username = String(username || '').trim();
  otpCode = String(otpCode || '').trim();
  newPassword = String(newPassword || '').trim();

  if (!username || !otpCode || !newPassword) {
    return { success: false, message: 'Kode OTP dan Password Baru wajib diisi.' };
  }

  if (newPassword.length < 6) {
    return { success: false, message: 'Password minimal 6 karakter.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const otpSheet = ss.getSheetByName('OTP_Reset');
  const userSheet = ss.getSheetByName('Users');

  if (!otpSheet || !userSheet) {
    return { success: false, message: 'Sheet database tidak ditemukan.' };
  }

  const otpData = otpSheet.getDataRange().getValues();
  const now = new Date();
  let validOtpRowIndex = -1;

  for (let i = otpData.length - 1; i >= 1; i--) {
    const rowId = String(otpData[i][0]).trim();
    const rowOtp = String(otpData[i][1]).trim();
    const rowExpiry = new Date(otpData[i][2]);
    const rowStatus = String(otpData[i][3]).trim();

    if (rowId.toLowerCase() === username.toLowerCase() && rowOtp === otpCode && rowStatus === 'ACTIVE') {
      if (now <= rowExpiry) {
        validOtpRowIndex = i + 1;
        break;
      } else {
        otpSheet.getRange(i + 1, 4).setValue('EXPIRED');
        return { success: false, message: 'Kode OTP telah kedaluwarsa (lebih dari 15 menit). Silakan minta OTP baru.' };
      }
    }
  }

  if (validOtpRowIndex === -1) {
    return { success: false, message: 'Kode OTP tidak valid.' };
  }

  // Update Users Sheet Password Hash
  const userData = userSheet.getDataRange().getValues();
  let userUpdated = false;

  for (let u = 1; u < userData.length; u++) {
    const rowId = String(userData[u][0]).trim();
    if (rowId.toLowerCase() === username.toLowerCase()) {
      const newHash = hashPassword(newPassword, username);
      userSheet.getRange(u + 1, 5).setValue(newHash); // Column 5 = Password Hash
      userUpdated = true;
      break;
    }
  }

  if (!userUpdated) {
    return { success: false, message: 'Akun user tidak ditemukan di database.' };
  }

  // Mark OTP as USED
  otpSheet.getRange(validOtpRowIndex, 4).setValue('USED');
  clearLoginFailures(username);

  return { success: true, message: 'Password berhasil diubah. Silakan login dengan password baru Anda.' };
}

function changeOwnPassword(token, oldPassword, newPassword) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  oldPassword = String(oldPassword || '').trim();
  newPassword = String(newPassword || '').trim();

  if (!oldPassword || !newPassword) {
    return { success: false, message: 'Password Saat Ini dan Password Baru wajib diisi.' };
  }

  if (newPassword.length < 6) {
    return { success: false, message: 'Password baru minimal 6 karakter.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

  const data = userSheet.getDataRange().getValues();
  const username = session.user.username;
  const oldHashSalted = hashPassword(oldPassword, username);
  const oldHashLegacy = hashLegacyPassword(oldPassword);
  const newHash = hashPassword(newPassword, username);

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0] || '').trim();
    if (rowId.toLowerCase() === username.toLowerCase()) {
      const currentHash = String(data[i][4] || '').trim();
      if (currentHash !== oldHashSalted && currentHash !== oldHashLegacy) {
        return { success: false, message: 'Password saat ini (lama) yang Anda masukkan tidak sesuai.' };
      }
      userSheet.getRange(i + 1, 5).setValue(newHash);
      return { success: true, message: 'Password Anda berhasil diperbarui.' };
    }
  }

  return { success: false, message: 'Akun user tidak ditemukan di database.' };
}

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
      karyawanList.push({
        id: String(userData[i][0]),
        name: String(userData[i][1]),
        jabatan: String(userData[i][2]),
        dept: String(userData[i][3]),
        role: String(userData[i][5]),
        email: String(userData[i][6] || '')
      });
    }
  }

  // Get Slips
  const slipSheet = ss.getSheetByName('SlipGaji');
  const slipData = slipSheet ? slipSheet.getDataRange().getValues() : [];
  const slipList = [];
  for (let j = 1; j < slipData.length; j++) {
    const empId = String(slipData[j][1] || '').trim().toLowerCase();
    if (isHRD || empId === currentUsername) {
      slipList.push({
        id: String(slipData[j][0]),
        empId: String(slipData[j][1]),
        empName: String(slipData[j][2]),
        bulan: String(slipData[j][3]),
        tahun: String(slipData[j][4]),
        fileName: String(slipData[j][5]),
        catatan: String(slipData[j][6]),
        driveUrl: String(slipData[j][7]),
        uploadDate: slipData[j][9] ? Utilities.formatDate(new Date(slipData[j][9]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '-',
        downloaded: String(slipData[j][10]).startsWith('Terkirim') || String(slipData[j][10]) === 'Diakses',
        emailStatus: String(slipData[j][10] || 'Belum')
      });
    }
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

  return { success: true, karyawan: karyawanList, slips: slipList, departemen: deptList };
}

function getSlipKaryawan(token, filterBulan, filterTahun) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const currentEmpId = session.user.username;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const slipSheet = ss.getSheetByName('SlipGaji');
  if (!slipSheet) return { success: true, data: [] };

  const slipData = slipSheet.getDataRange().getValues();
  const result = [];

  for (let i = 1; i < slipData.length; i++) {
    const empId = String(slipData[i][1]);
    const bulan = String(slipData[i][3]);
    const tahun = String(slipData[i][4]);

    if (empId === currentEmpId) {
      if ((filterBulan === 'SEMUA' || bulan === filterBulan) && tahun === String(filterTahun)) {
        result.push({
          id: String(slipData[i][0]),
          empId: empId,
          empName: String(slipData[i][2]),
          bulan: bulan,
          tahun: tahun,
          fileName: String(slipData[i][5]),
          catatan: String(slipData[i][6]),
          driveUrl: String(slipData[i][7]),
          uploadDate: slipData[i][9] ? Utilities.formatDate(new Date(slipData[i][9]), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '-'
        });
      }
    }
  }

  return { success: true, data: result };
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
        new Date()
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

      const existingRow = existingMap[empId];

      if (existingRow) {
        userSheet.getRange(existingRow, 2).setValue(empName);
        userSheet.getRange(existingRow, 3).setValue(jabatan);
        userSheet.getRange(existingRow, 4).setValue(dept);
        if (rawPass) {
          userSheet.getRange(existingRow, 5).setValue(hashedPass);
        }
        userSheet.getRange(existingRow, 6).setValue(role);
      } else {
        userSheet.appendRow([empId, empName, jabatan, dept, hashedPass, role]);
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

const MAX_FILE_SIZE_BYTES = 1048576; // 1 MB

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
      return { success: false, message: 'Ukuran file melebihi batas maksimal 1 MB.' };
    }

    // Save PDF to Drive: Storage/Tahun/Bulan/Nama Karyawan
    const folder = getSlipStorageFolder(payload.tahun, payload.bulan, empName);

    const decoded = Utilities.base64Decode(payload.fileData);
    const blob = Utilities.newBlob(decoded, payload.mimeType || 'application/pdf', payload.fileName);
    const file = folder.createFile(blob);

    const slipId = 'SLIP-' + Date.now().toString().slice(-6);
    const fileUrl = file.getUrl();
    const fileId = file.getId();

    const slipSheet = ss.getSheetByName('SlipGaji');
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
      'Belum'
    ]);

    lock.releaseLock();
    return { success: true, fileUrl: fileUrl };
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

    // If new file is uploaded
    if (payload.fileData && payload.fileName) {
      // Validate PDF file size (max 1 MB)
      if (getBase64FileSize(payload.fileData) > MAX_FILE_SIZE_BYTES) {
        lock.releaseLock();
        return { success: false, message: 'Ukuran file melebihi batas maksimal 1 MB.' };
      }

      // Save PDF to Drive: Storage/Tahun/Bulan/Nama Karyawan
      const folder = getSlipStorageFolder(payload.tahun, payload.bulan, empName);

      // Trash old file if exists
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
        if (fileId) {
          try {
            DriveApp.getFileById(fileId).setTrashed(true);
          } catch (e) {
            console.error('Drive file delete error:', e);
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

function getSlipPdfBase64(token, slipId) {
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

    // Security check: Karyawan can only view their own slip unless Admin
    if (session.user.role !== 'Admin' && String(session.user.username).trim().toLowerCase() !== String(empId).trim().toLowerCase()) {
      return { success: false, message: 'Akses ditolak. Anda hanya dapat membaca slip gaji milik Anda sendiri.' };
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

    for (let i = 1; i < slipData.length; i++) {
      if (String(slipData[i][0]).trim() === String(slipId).trim()) {
        rowIndex = i + 1;
        empId = String(slipData[i][1]).trim();
        empName = String(slipData[i][2]).trim();
        bulan = String(slipData[i][3]).trim();
        tahun = String(slipData[i][4]).trim();
        fileName = String(slipData[i][5]).trim();
        fileId = String(slipData[i][8]).trim();
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

    const appTitle = 'e-Slip Gaji Online';
    const periodeStr = bulan + ' ' + tahun;
    // Clean subject without brackets or spammy prefixes
    const subject = 'Slip Gaji Periode ' + periodeStr + ' - ' + empName;

    // Plain text fallback (Crucial for spam filtering)
    const plainBody = 'Yth. ' + empName + ' (' + empId + '),\n\n' +
      'Slip gaji Anda untuk periode ' + periodeStr + ' telah diterbitkan.\n' +
      'Dokumen PDF slip gaji terlampir langsung pada email ini.\n\n' +
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
      '</table>' +
      '</div>' +
      '<p style="font-size: 13px; color: #475569; line-height: 1.5; background-color: #f1f5f9; border-left: 3px solid #94a3b8; padding: 12px; border-radius: 4px;">' +
      '<strong>Catatan Kerahasiaan:</strong> Dokumen PDF slip gaji terlampir langsung pada email ini. Dokumen ini hanya ditujukan kepada pemilik gaji.' +
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
      attachments: [blob]
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
        attachments: [blob]
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


