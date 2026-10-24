// --- API FUNCTION WHITELISTS ---
const ALLOWED_PUBLIC_ACTIONS = [
  'loginUser',
  'requestPasswordReset',
  'verifyAndResetPassword',
  'getActiveAnnouncement',
  'registerNewUser',
  'getPublicDepartments'
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
  'getResiFileBase64',
  'sendSlipEmailNotification',
  'sendSlipEmailNotificationBatch',
  'submitPengajuanCuti',
  'editPengajuanCuti',
  'getCutiKaryawan',
  'deletePengajuanCuti',
  'getLampiranCutiBase64',
  'saveRoster',
  'deleteRoster',
  'getRosterList',
  'processApprovalCuti',
  'getAnnouncements',
  'saveAnnouncement',
  'deleteAnnouncement',
  'toggleAnnouncementStatus',
  'getPendaftarList',
  'processApprovalPendaftar',
  'checkUserHasSlipPin',
  'setUserSlipPin',
  'verifyUserSlipPin',
  'requestResetSlipPinOtp',
  'verifyAndResetSlipPin',
  'adminResetSlipPin'
];

function doGet(e) {
  // Hanya layani UI HTML Web App (Cegah eksekusi aksi sensitif via GET URL parameter)
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('E-Slip Gaji')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
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

function ensureUserSheetColumns(userSheet) {
  if (!userSheet) return;
  const lastCol = userSheet.getLastColumn();
  if (lastCol < 9) {
    userSheet.getRange(1, 9).setValue('Pola Roster').setFontWeight('bold').setBackground('#e2e8f0');
  }
  if (lastCol < 10) {
    userSheet.getRange(1, 10).setValue('Tanggal Mulai Kerja').setFontWeight('bold').setBackground('#e2e8f0');
  }
  if (lastCol < 11) {
    userSheet.getRange(1, 11).setValue('PIN Slip Hash').setFontWeight('bold').setBackground('#e2e8f0');
  }
}

// --- INITIAL DATABASE SETUP ---
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet Users
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow(['ID Karyawan', 'Nama Lengkap', 'Jabatan', 'Departemen', 'Password Hash', 'Role', 'Email', 'Tanggal Dibuat', 'Pola Roster', 'Tanggal Mulai Kerja']);
    userSheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#e2e8f0');

    // Default Admin Login: admin / Admin@2026
    const defaultPassHash = hashPassword('Admin@2026', 'admin');
    userSheet.appendRow(['admin', 'HRD Administrator', 'HR Manager', 'Human Resource', defaultPassHash, 'Admin', 'admin@company.com', new Date(), 'Non-Roster', '']);
  } else {
    ensureUserSheetColumns(userSheet);
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

  // 5. Sheet PengajuanCuti
  getOrCreatePengajuanCutiSheet(ss);

  // 6. Sheet MasterRoster
  getOrCreateMasterRosterSheet(ss);

  // 7. Sheet Pengumuman
  getOrCreatePengumumanSheet(ss);

  // 8. Sheet PendaftarAkun
  getOrCreatePendaftarAkunSheet(ss);

  return 'Setup Database Berhasil! Sheet Users, SlipGaji, Departemen, OTP_Reset, PengajuanCuti, MasterRoster, Pengumuman, dan PendaftarAkun telah siap.';
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

function generateSessionToken(username, role, fullname, dept, jabatan) {
  const payload = {
    u: String(username || ''),
    r: String(role || ''),
    f: String(fullname || ''),
    d: String(dept || ''),
    j: String(jabatan || ''),
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

    let userDept = payload.d || '';
    let userJab = payload.j || '';
    let userRole = payload.r || '';
    let userFullName = payload.f || '';

    // Auto-fallback: ambil data user terkini dari Sheet Users jika token lama belum menyimpan dept/jabatan
    if (!userDept || !userJab) {
      try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const uSheet = ss.getSheetByName('Users');
        if (uSheet) {
          const uData = uSheet.getDataRange().getValues();
          const targetU = String(payload.u || '').trim().toLowerCase();
          for (let ui = 1; ui < uData.length; ui++) {
            if (String(uData[ui][0] || '').trim().toLowerCase() === targetU) {
              if (!userFullName) userFullName = String(uData[ui][1] || '').trim();
              if (!userJab) userJab = String(uData[ui][2] || '').trim();
              if (!userDept) userDept = String(uData[ui][3] || '').trim();
              if (!userRole) userRole = String(uData[ui][5] || '').trim();
              break;
            }
          }
        }
      } catch (fErr) {
        console.error('Session user details fallback note:', fErr);
      }
    }

    return {
      valid: true,
      user: {
        username: payload.u,
        role: userRole,
        fullname: userFullName,
        dept: userDept,
        jabatan: userJab
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
        const tglMulai = data[i][9] instanceof Date
          ? Utilities.formatDate(data[i][9], 'Asia/Jakarta', 'yyyy-MM-dd')
          : String(data[i][9] || '').trim();
        const userObj = {
          username: data[i][0],
          fullname: data[i][1],
          jabatan: data[i][2],
          dept: data[i][3],
          role: data[i][5],
          email: String(data[i][6] || '').trim(),
          polaRoster: String(data[i][8] || 'Non-Roster').trim() || 'Non-Roster',
          tglMulaiKerja: tglMulai,
          hasSlipPin: Boolean(String(data[i][10] || '').trim())
        };
        const token = generateSessionToken(userObj.username, userObj.role, userObj.fullname, userObj.dept, userObj.jabatan);
        return { success: true, user: userObj, token: token };
      }
    }
  }

  // Cek apakah akun terdaftar di PendaftarAkun dan masih Menunggu Persetujuan atau Ditolak
  try {
    const pendaftarSheet = ss.getSheetByName('PendaftarAkun');
    if (pendaftarSheet) {
      const pData = pendaftarSheet.getDataRange().getValues();
      for (let pi = pData.length - 1; pi >= 1; pi--) {
        const pId = String(pData[pi][1] || '').trim().toLowerCase();
        if (pId === username.toLowerCase()) {
          const pStatus = String(pData[pi][8] || '').trim();
          const pCatatan = String(pData[pi][11] || '').trim();
          if (pStatus === 'Menunggu Persetujuan') {
            return {
              success: false,
              message: 'Akun Anda sedang menunggu persetujuan dari HRD. Silakan hubungi bagian HRD untuk konfirmasi akun Anda.'
            };
          } else if (pStatus === 'Ditolak') {
            return {
              success: false,
              message: 'Pendaftaran akun Anda ditolak oleh HRD' + (pCatatan ? ': ' + pCatatan : '.') + ' Silakan hubungi HRD untuk informasi lebih lanjut.'
            };
          }
        }
      }
    }
  } catch (pCheckErr) {
    console.error('Check pendaftar status note:', pCheckErr);
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

// --- PIN KEAMANAN SLIP GAJI MODULE ---
function hashSlipPin(pin, username) {
  const salt = username ? String(username).trim().toLowerCase() : '';
  const combined = 'eSlip_PinSalt_' + salt + '_' + String(pin).trim();
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined, Utilities.Charset.UTF_8);
  return rawHash.map(function (b) {
    const hex = (b < 0 ? b + 256 : b).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

function checkPinRateLimit(username) {
  const cache = CacheService.getScriptCache();
  const attempts = parseInt(cache.get('pin_rate_' + username) || '0', 10);
  return attempts >= 5;
}

function recordPinFailure(username) {
  const cache = CacheService.getScriptCache();
  const key = 'pin_rate_' + username;
  const attempts = parseInt(cache.get(key) || '0', 10) + 1;
  cache.put(key, String(attempts), 300); // Kunci 5 menit setelah 5x gagal
  return Math.max(0, 5 - attempts);
}

function clearPinFailures(username) {
  const cache = CacheService.getScriptCache();
  cache.remove('pin_rate_' + username);
}

function checkUserHasSlipPin(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

  const data = userSheet.getDataRange().getValues();
  const username = String(session.user.username || '').trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0] || '').trim().toLowerCase();
    if (rowId === username) {
      const pinHash = String(data[i][10] || '').trim();
      return { success: true, hasPin: Boolean(pinHash) };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

function verifyUserSlipPin(token, pin) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  pin = String(pin || '').trim();
  if (!pin || !/^\d{6}$/.test(pin)) {
    return { success: false, message: 'PIN harus berupa 6 digit angka.' };
  }

  const username = String(session.user.username || '').trim().toLowerCase();
  if (checkPinRateLimit(username)) {
    return { success: false, locked: true, message: 'Terlalu banyak percobaan PIN salah. Akses slip gaji dikunci sementara 5 menit.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

  const data = userSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0] || '').trim().toLowerCase();
    if (rowId === username) {
      const pinHash = String(data[i][10] || '').trim();
      if (!pinHash) {
        return { success: false, needSetup: true, message: 'PIN Keamanan belum dibuat.' };
      }
      const inputHash = hashSlipPin(pin, username);
      if (inputHash === pinHash) {
        clearPinFailures(username);
        return { success: true, message: 'PIN valid.' };
      } else {
        const remaining = recordPinFailure(username);
        if (remaining <= 0) {
          return { success: false, locked: true, message: 'PIN salah 5x berturut-turut. Akses slip gaji dikunci sementara 5 menit demi keamanan.' };
        }
        return { success: false, message: 'PIN salah. Sisa kesempatan: ' + remaining + 'x.' };
      }
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

function setUserSlipPin(token, newPin, oldPin) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  newPin = String(newPin || '').trim();
  if (!newPin || !/^\d{6}$/.test(newPin)) {
    return { success: false, message: 'PIN baru harus berupa 6 digit angka.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

  ensureUserSheetColumns(userSheet);

  const data = userSheet.getDataRange().getValues();
  const username = String(session.user.username || '').trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0] || '').trim().toLowerCase();
    if (rowId === username) {
      const currentPinHash = String(data[i][10] || '').trim();

      // Jika sudah punya PIN sebelumnya, wajib verifikasi oldPin
      if (currentPinHash) {
        if (!oldPin) {
          return { success: false, message: 'PIN saat ini (lama) wajib dimasukkan.' };
        }
        const oldHash = hashSlipPin(oldPin, username);
        if (oldHash !== currentPinHash) {
          return { success: false, message: 'PIN saat ini (lama) yang Anda masukkan tidak cocok.' };
        }
      }

      const newHash = hashSlipPin(newPin, username);
      userSheet.getRange(i + 1, 11).setValue(newHash);
      clearPinFailures(username);
      return { success: true, message: 'PIN Keamanan Slip Gaji berhasil ' + (currentPinHash ? 'diperbarui' : 'disimpan') + '.' };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

function requestResetSlipPinOtp(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

  const data = userSheet.getDataRange().getValues();
  const username = String(session.user.username || '').trim().toLowerCase();
  let userEmail = '';
  let userFullname = '';

  for (let i = 1; i < data.length; i++) {
    const rowId = String(data[i][0] || '').trim().toLowerCase();
    if (rowId === username) {
      userFullname = String(data[i][1] || '').trim();
      userEmail = String(data[i][6] || '').trim();
      break;
    }
  }

  if (!userEmail) {
    return { success: false, message: 'Akun Anda belum memiliki email terdaftar. Hubungi Admin HRD untuk reset PIN.' };
  }

  const otpCode = String(Math.floor(100000 + Math.random() * 900000));
  const now = new Date();
  const expiresAt = new Date(now.getTime() + 15 * 60 * 1000); // 15 Menit

  let otpSheet = ss.getSheetByName('OTP_Reset');
  if (!otpSheet) {
    otpSheet = ss.insertSheet('OTP_Reset');
    otpSheet.appendRow(['ID Karyawan', 'Kode OTP', 'Expired At', 'Status', 'Tanggal Dibuat', 'Tipe']);
    otpSheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e2e8f0');
  }

  otpSheet.appendRow([session.user.username, otpCode, expiresAt, 'ACTIVE', now, 'RESET_SLIP_PIN']);

  try {
    const maskedEmail = userEmail.replace(/(.{2})(.*)(?=@)/, function (gp1, gp2, gp3) {
      return gp2 + "*".repeat(gp3.length);
    });

    MailApp.sendEmail({
      to: userEmail,
      subject: '🛡️ Kode OTP Reset PIN Slip Gaji - Portal E-Slip',
      htmlBody: `
        <div style="font-family: Arial, sans-serif; max-width: 480px; margin: 0 auto; padding: 24px; border: 1px solid #e2e8f0; border-radius: 16px; background-color: #ffffff;">
          <div style="text-align: center; margin-bottom: 20px;">
            <h2 style="color: #059669; margin: 0; font-size: 20px;">Portal E-Slip Gaji Online</h2>
            <p style="color: #64748b; font-size: 14px; margin-top: 4px;">Permintaan Reset PIN Slip Gaji</p>
          </div>
          <p style="color: #334155; font-size: 14px; line-height: 1.5;">Halo <b>` + (userFullname || session.user.username) + `</b>,</p>
          <p style="color: #334155; font-size: 14px; line-height: 1.5;">Kami menerima permintaan untuk mereset <b>PIN Keamanan Slip Gaji</b> Anda. Gunakan kode OTP 6 digit di bawah ini untuk melanjutkan:</p>
          <div style="text-align: center; margin: 24px 0;">
            <span style="display: inline-block; font-size: 32px; font-weight: bold; letter-spacing: 8px; color: #059669; background-color: #ecfdf5; padding: 14px 28px; border-radius: 12px; border: 1px dashed #10b981;">
              ` + otpCode + `
            </span>
          </div>
          <p style="color: #64748b; font-size: 12px; line-height: 1.5;">Kode OTP ini berlaku selama <b>15 menit</b>. Jangan berikan kode ini kepada siapa pun termasuk rekan kerja atau HRD.</p>
          <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 20px 0;">
          <p style="color: #94a3b8; font-size: 11px; text-align: center; margin: 0;">Jika Anda tidak merasa meminta reset PIN, segera amankan akun Anda atau hubungi HRD.</p>
        </div>
      `
    });

    return { success: true, message: 'Kode OTP telah dikirim ke email ' + maskedEmail };
  } catch (err) {
    return { success: false, message: 'Gagal mengirim email OTP: ' + err.message };
  }
}

function verifyAndResetSlipPin(token, otpCode, newPin) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  otpCode = String(otpCode || '').trim();
  newPin = String(newPin || '').trim();

  if (!otpCode || !/^\d{6}$/.test(otpCode)) {
    return { success: false, message: 'Kode OTP harus 6 digit angka.' };
  }
  if (!newPin || !/^\d{6}$/.test(newPin)) {
    return { success: false, message: 'PIN baru harus berupa 6 digit angka.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const otpSheet = ss.getSheetByName('OTP_Reset');
  if (!otpSheet) return { success: false, message: 'Data verifikasi OTP tidak ditemukan.' };

  const otpData = otpSheet.getDataRange().getValues();
  const username = String(session.user.username || '').trim().toLowerCase();
  const now = new Date();
  let validRow = -1;

  for (let i = otpData.length - 1; i >= 1; i--) {
    const rowUser = String(otpData[i][0] || '').trim().toLowerCase();
    const rowCode = String(otpData[i][1] || '').trim();
    const rowExp = new Date(otpData[i][2]);
    const rowStatus = String(otpData[i][3] || '').trim();

    if (rowUser === username && rowCode === otpCode) {
      if (rowStatus !== 'ACTIVE') {
        return { success: false, message: 'Kode OTP ini sudah pernah digunakan atau kedaluwarsa.' };
      }
      if (now > rowExp) {
        otpSheet.getRange(i + 1, 4).setValue('EXPIRED');
        return { success: false, message: 'Kode OTP telah kedaluwarsa (lebih dari 15 menit).' };
      }
      validRow = i + 1;
      break;
    }
  }

  if (validRow === -1) {
    return { success: false, message: 'Kode OTP salah. Periksa kembali email Anda.' };
  }

  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };
  ensureUserSheetColumns(userSheet);

  const uData = userSheet.getDataRange().getValues();
  for (let u = 1; u < uData.length; u++) {
    const rId = String(uData[u][0] || '').trim().toLowerCase();
    if (rId === username) {
      const newHash = hashSlipPin(newPin, username);
      userSheet.getRange(u + 1, 11).setValue(newHash);
      otpSheet.getRange(validRow, 4).setValue('USED');
      clearPinFailures(username);
      return { success: true, message: 'PIN Keamanan Slip Gaji berhasil direset.' };
    }
  }

  return { success: false, message: 'User tidak ditemukan.' };
}

function adminResetSlipPin(token, empId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const role = String(session.user.role || '').toLowerCase();
  if (role !== 'admin' && role !== 'hrd') {
    return { success: false, message: 'Akses ditolak: Hanya HRD / Admin yang dapat mereset PIN karyawan.' };
  }

  empId = String(empId || '').trim();
  if (!empId) return { success: false, message: 'ID Karyawan wajib ditentukan.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { success: false, message: 'Sheet Users tidak ditemukan.' };

  const uData = userSheet.getDataRange().getValues();
  for (let i = 1; i < uData.length; i++) {
    const rowId = String(uData[i][0] || '').trim().toLowerCase();
    if (rowId === empId.toLowerCase()) {
      userSheet.getRange(i + 1, 11).setValue('');
      clearPinFailures(empId);
      return { success: true, message: 'PIN Slip Gaji untuk karyawan ' + uData[i][1] + ' (' + empId + ') berhasil direset.' };
    }
  }

  return { success: false, message: 'Karyawan dengan ID ' + empId + ' tidak ditemukan.' };
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
        hasSlipPin: Boolean(String(userData[i][10] || '').trim())
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
        payload.tglMulaiKerja || ''
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

// --- MODUL PENGAJUAN CUTI KARYAWAN ---

function getOrCreatePengajuanCutiSheet(ss) {
  let sheet = ss.getSheetByName('PengajuanCuti') ||
    ss.getSheetByName('Pengajuan Cuti') ||
    ss.getSheetByName('Riwayat Cuti') ||
    ss.getSheetByName('Riwayat_Cuti') ||
    ss.getSheetByName('Data Cuti') ||
    ss.getSheetByName('Cuti');

  if (!sheet) {
    const sheets = ss.getSheets();
    for (let s = 0; s < sheets.length; s++) {
      const name = sheets[s].getName().toLowerCase();
      if (name.includes('cuti') && !name.includes('roster')) {
        sheet = sheets[s];
        break;
      }
    }
  }

  const REQUIRED_HEADERS = [
    'ID Pengajuan', 'Tanggal Pengajuan', 'ID Karyawan', 'Nama Karyawan',
    'Departemen', 'Jabatan', 'Jenis Cuti', 'Tanggal Mulai',
    'Tanggal Selesai', 'Jumlah Hari', 'Alasan Cuti', 'Alamat Selama Cuti',
    'Nomor Kontak', 'Nama File Lampiran', 'URL Drive', 'File ID Drive',
    'Catatan Tambahan', 'Status',
    'Approver 1 Nama', 'Approver 1 Jabatan', 'Approver 1 Tanggal', 'Approver 1 Status', 'Approver 1 Catatan',
    'Approver 2 Nama', 'Approver 2 Jabatan', 'Approver 2 Tanggal', 'Approver 2 Status', 'Approver 2 Catatan',
    'Approver 3 Nama', 'Approver 3 Jabatan', 'Approver 3 Tanggal', 'Approver 3 Status', 'Approver 3 Catatan',
    'Alasan Penolakan', 'Tanggal Awal Onsite', 'Total Hari Kerja Onsite',
    'Titik Drop Off', 'Moda Transportasi', 'Catatan Logistik Tiket',
    'Lokasi Kerja'
  ];

  if (!sheet) {
    sheet = ss.insertSheet('PengajuanCuti');
    sheet.appendRow(REQUIRED_HEADERS);
    sheet.getRange(1, 1, 1, REQUIRED_HEADERS.length).setFontWeight('bold').setBackground('#e2e8f0');
  } else {
    // Auto-migrasi header jika sheet lama belum memiliki kolom approval bertingkat atau logistik transportasi
    const lastCol = sheet.getLastColumn();
    if (lastCol > 0) {
      const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
        return String(h || '').trim().toLowerCase();
      });
      const missingHeaders = [];
      REQUIRED_HEADERS.forEach(function (reqH) {
        const found = existingHeaders.some(function (eh) {
          if (eh.toLowerCase() === reqH.toLowerCase()) return true;
          if (reqH.toLowerCase().includes('titik drop') && (eh.includes('point') || eh.includes('hire') || eh.includes('drop') || eh.includes('poh') || eh.includes('jemput'))) return true;
          if (reqH.toLowerCase().includes('moda transport') && (eh.includes('transport') || eh.includes('moda') || eh.includes('kendaraan'))) return true;
          if (reqH.toLowerCase().includes('catatan logistik') && (eh.includes('logistik') || eh.includes('tiket') || (eh.includes('catatan') && (eh.includes('transport') || eh.includes('khusus'))))) return true;
          if (reqH.toLowerCase().includes('lokasi kerja') && (eh.includes('lokasi') || eh.includes('jobsite') || eh === 'site')) return true;
          return false;
        });
        if (!found) {
          missingHeaders.push(reqH);
        }
      });
      if (missingHeaders.length > 0) {
        const startCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]).setFontWeight('bold').setBackground('#e2e8f0');
        SpreadsheetApp.flush();
      }
    }
  }
  return sheet;
}

function cleanDept(d) {
  if (!d) return '';
  let s = String(d).trim().toLowerCase();
  s = s.replace(/^(departemen|dept\.?|divisi|division|bagian|seksi|unit)\s+/g, '');
  s = s.replace(/[^a-z0-9]/g, '').trim();

  if (s === 'hr' || s === 'hrd' || s === 'humanresource' || s === 'humanresources' || s === 'personalia') return 'hrd';
  if (s === 'hrga' || s === 'hrdga' || s === 'hrandga' || s === 'hrdandga' || s === 'humanresourcegeneralaffair' || s === 'humanresourceandgeneralaffair') return 'hrga';
  if (s === 'it' || s === 'ti' || s === 'informationtechnology' || s === 'teknologiinformasi' || s === 'edp' || s === 'ict') return 'it';
  if (s === 'finance' || s === 'keuangan' || s === 'fa' || s === 'financeaccounting' || s === 'accounting') return 'finance';
  if (s === 'ga' || s === 'generalaffair' || s === 'generalaffairs' || s === 'umum') return 'ga';

  return s;
}

function isSameDept(dept1, dept2) {
  const c1 = cleanDept(dept1);
  const c2 = cleanDept(dept2);
  if (!c1 || !c2) return false;
  if (c1 === c2) return true;
  if ((c1 === 'hrga' && (c2 === 'hrd' || c2 === 'ga' || c2.startsWith('hr') || c2.endsWith('ga'))) ||
    (c2 === 'hrga' && (c1 === 'hrd' || c1 === 'ga' || c1.startsWith('hr') || c1.endsWith('ga')))) return true;
  if (c1 === 'it' && (c2.startsWith('it') || c2.endsWith('it'))) return true;
  if (c2 === 'it' && (c1.startsWith('it') || c1.endsWith('it'))) return true;
  if (c1.length >= 3 && c2.length >= 3) {
    if (c1.includes(c2) || c2.includes(c1)) return true;
  }
  return false;
}

function normalizeDept(d) {
  return cleanDept(d);
}

function isJobPJO(role, jabatan) {
  const r = String(role || '').trim().toLowerCase();
  const j = String(jabatan || '').trim().toLowerCase();
  if (r === 'pjo' || r === 'direktur' || r === 'director') return true;
  const keywords = [
    'pjo',
    'penanggung jawab operasional',
    'penanggungjawab operasional',
    'ktt',
    'kepala teknik tambang',
    'direktur',
    'director',
    'project manager',
    'general manager'
  ];
  for (let k = 0; k < keywords.length; k++) {
    if (j.includes(keywords[k])) return true;
  }
  return false;
}

function isJobAtasan(role, jabatan) {
  const r = String(role || '').trim().toLowerCase();
  const j = String(jabatan || '').trim().toLowerCase();
  if (r === 'supervisor' || r === 'atasan' || r === 'spv' || r === 'manager' || r === 'lead' || r === 'leader') return true;
  const keywords = [
    'supervisor', 'spv', 'atasan', 'manager', 'lead', 'head', 'kabag', 'kabid', 'kasie',
    'koordinator', 'direktur', 'director', 'superintendent', 'kepala', 'leader', 'tl',
    'team lead', 'pengawas', 'asmen', 'asisten manager'
  ];
  for (let k = 0; k < keywords.length; k++) {
    if (j.includes(keywords[k])) return true;
  }
  return false;
}

function parseCutiDataFromSheet(cutiSheet, sessionUser) {
  if (!cutiSheet) return [];
  const data = cutiSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Load user ground truth map dari sheet Users
  const userSheet = ss.getSheetByName('Users');
  const userSheetData = userSheet ? userSheet.getDataRange().getValues() : [];
  const userMap = {};
  for (let u = 1; u < userSheetData.length; u++) {
    const uRow = userSheetData[u];
    const uId = String(uRow[0] || '').trim().toLowerCase();
    if (uId) {
      userMap[uId] = {
        id: String(uRow[0] || '').trim(),
        fullname: String(uRow[1] || '').trim(),
        jabatan: String(uRow[2] || '').trim(),
        dept: String(uRow[3] || '').trim(),
        role: String(uRow[5] || '').trim()
      };
    }
  }

  const currentUsername = sessionUser ? String(sessionUser.username || '').trim().toLowerCase() : '';
  const currentDbUser = userMap[currentUsername] || null;

  // Pastikan role, dept, jabatan user yang sedang login selalu up-to-date
  const userRole = currentDbUser ? String(currentDbUser.role || '').trim() : (sessionUser ? String(sessionUser.role || '').trim() : '');
  const userDept = currentDbUser ? String(currentDbUser.dept || '').trim() : (sessionUser ? String(sessionUser.dept || '').trim() : '');
  const userJab = currentDbUser ? String(currentDbUser.jabatan || '').trim() : (sessionUser ? String(sessionUser.jabatan || '').trim() : '');

  const userDeptNorm = normalizeDept(userDept);
  const userRoleLower = userRole.toLowerCase();
  const isHRD = (userRoleLower === 'admin' || userRoleLower === 'hrd' || currentUsername === 'admin');
  const isSupervisor = isJobAtasan(userRole, userJab);
  const isPJO = isJobPJO(userRole, userJab);

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id pengajuan') || h === 'id cuti' || h === 'id' || h.includes('kode'); });
  const tglAjuIdx = headers.findIndex(function (h) { return h.includes('tanggal pengajuan') || h.includes('tgl aju') || h.includes('tgl pengajuan') || h === 'tanggal'; });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h.includes('nip') || h === 'username' || h === 'id_karyawan'; });
  const nameIdx = headers.findIndex(function (h) { return h.includes('nama'); });
  const deptIdx = headers.findIndex(function (h) { return h.includes('departemen') || h.includes('dept'); });
  const jabIdx = headers.findIndex(function (h) { return h.includes('jabatan'); });
  const jenisIdx = headers.findIndex(function (h) { return h.includes('jenis'); });
  const mulaiIdx = headers.findIndex(function (h) { return h.includes('mulai'); });
  const selesaiIdx = headers.findIndex(function (h) { return h.includes('selesai') || h.includes('akhir'); });
  const hariIdx = headers.findIndex(function (h) { return h.includes('hari') || h.includes('durasi') || h.includes('lama'); });
  const alasanIdx = headers.findIndex(function (h) { return h.includes('alasan') || h.includes('keterangan') || h.includes('keperluan'); });
  const alamatIdx = headers.findIndex(function (h) { return h.includes('alamat'); });
  const kontakIdx = headers.findIndex(function (h) { return h.includes('kontak') || h.includes('telepon') || h.includes('hp') || h.includes('no'); });
  const lampNameIdx = headers.findIndex(function (h) { return h.includes('lampiran') || h.includes('file'); });
  const driveUrlIdx = headers.findIndex(function (h) { return h.includes('url') || h.includes('link'); });
  const driveIdIdx = headers.findIndex(function (h) { return h.includes('file id') || h.includes('fileid'); });
  const catatanIdx = headers.findIndex(function (h) { return h.includes('catatan') && !h.includes('approver') && !h.includes('logistik') && !h.includes('transport') && !h.includes('tiket') && !h.includes('jemput'); });
  const statusIdx = headers.findIndex(function (h) { return h.includes('status'); });

  // Approver 1 (Atasan)
  const app1NameIdx = headers.findIndex(function (h) { return h.includes('approver 1 nama') || h.includes('approver1 nama') || h === 'approver 1'; });
  const app1JabIdx = headers.findIndex(function (h) { return h.includes('approver 1 jabatan') || h.includes('approver1 jabatan'); });
  const app1TglIdx = headers.findIndex(function (h) { return h.includes('approver 1 tanggal') || h.includes('approver 1 tgl'); });
  const app1StatusIdx = headers.findIndex(function (h) { return h.includes('approver 1 status'); });
  const app1CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 1 catatan'); });

  // Approver 2 (HRD)
  const app2NameIdx = headers.findIndex(function (h) { return h.includes('approver 2 nama') || h.includes('approver2 nama') || h === 'approver 2'; });
  const app2JabIdx = headers.findIndex(function (h) { return h.includes('approver 2 jabatan') || h.includes('approver2 jabatan'); });
  const app2TglIdx = headers.findIndex(function (h) { return h.includes('approver 2 tanggal') || h.includes('approver 2 tgl'); });
  const app2StatusIdx = headers.findIndex(function (h) { return h.includes('approver 2 status'); });
  const app2CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 2 catatan'); });

  // Approver 3 (PJO / Direktur)
  const app3NameIdx = headers.findIndex(function (h) { return h.includes('approver 3 nama') || h.includes('approver3 nama') || h === 'approver 3'; });
  const app3JabIdx = headers.findIndex(function (h) { return h.includes('approver 3 jabatan') || h.includes('approver3 jabatan'); });
  const app3TglIdx = headers.findIndex(function (h) { return h.includes('approver 3 tanggal') || h.includes('approver 3 tgl'); });
  const app3StatusIdx = headers.findIndex(function (h) { return h.includes('approver 3 status'); });
  const app3CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 3 catatan'); });

  const tolakAlasanIdx = headers.findIndex(function (h) { return h.includes('alasan penolakan') || h.includes('alasan tolak'); });
  const tglAwalOnsiteIdx = headers.findIndex(function (h) { return h.includes('tanggal awal onsite') || h.includes('awal onsite') || h.includes('tgl awal onsite'); });
  const totalHariOnsiteIdx = headers.findIndex(function (h) { return h.includes('total hari kerja onsite') || h.includes('hari kerja onsite') || h.includes('hari onsite') || h.includes('total hari onsite'); });
  const titikDropOffIdx = headers.findIndex(function (h) {
    return h.includes('titik drop') || h.includes('point of hire') || h.includes('point off hire') ||
           h.includes('drop') || h.includes('hire') || h.includes('point') || h.includes('poh') ||
           h.includes('jemput') || h.includes('antar') || h.includes('tujuan');
  });
  const modaTransportIdx = headers.findIndex(function (h) {
    return (h.includes('transport') || h.includes('moda') || h.includes('kendaraan') || h.includes('armada')) &&
           !h.includes('catatan') && !h.includes('note');
  });
  const catatanLogistikIdx = headers.findIndex(function (h) {
    return h.includes('logistik') || h.includes('tiket') || h.includes('penjemputan') ||
           (h.includes('catatan') && (h.includes('transport') || h.includes('khusus') || h.includes('tiket') || h.includes('jemput') || h.includes('drop')));
  });
  const lokasiKerjaIdx = headers.findIndex(function (h) {
    return h.includes('lokasi') || h.includes('jobsite') || h === 'site';
  });

  const list = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(function (cell) { return String(cell || '').trim() === ''; })) continue;

    const rowEmpId = empIdIdx !== -1 ? String(row[empIdIdx] || '').trim().toLowerCase() : String(row[2] || '').trim().toLowerCase();
    const applicantDb = userMap[rowEmpId] || null;
    const rawRowDept = deptIdx !== -1 ? String(row[deptIdx] || '').trim() : String(row[4] || '').trim();
    const rowDeptNorm = normalizeDept(applicantDb && applicantDb.dept ? applicantDb.dept : rawRowDept);

    // Hak Lihat Riwayat Cuti:
    // 1. HRD/Admin dan PJO/Direktur melihat semua pengajuan seluruh departemen.
    // 2. Supervisor melihat:
    //    - Pengajuan dari departemennya sendiri (userDeptNorm === rowDeptNorm)
    //    - Pengajuan miliknya sendiri (rowEmpId === currentUsername)
    // 3. Karyawan biasa hanya melihat pengajuannya sendiri.
    if (!isHRD && !isPJO) {
      if (isSupervisor) {
        const isOwnRow = (rowEmpId === currentUsername);
        const isDeptRow = isSameDept(userDept, rowDeptNorm) || isSameDept(userDept, rawRowDept) || (userDeptNorm !== '' && (userDeptNorm === rowDeptNorm || userDeptNorm === normalizeDept(rawRowDept)));
        if (!isOwnRow && !isDeptRow) continue;
      } else {
        if (rowEmpId !== currentUsername) continue;
      }
    }

    const rawTglAju = tglAjuIdx !== -1 ? row[tglAjuIdx] : row[1];
    const rawMulai = mulaiIdx !== -1 ? row[mulaiIdx] : row[7];
    const rawSelesai = selesaiIdx !== -1 ? row[selesaiIdx] : row[8];

    const tglAjuStr = rawTglAju instanceof Date
      ? Utilities.formatDate(rawTglAju, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm')
      : String(rawTglAju || '').trim();

    const mulaiStr = rawMulai instanceof Date
      ? Utilities.formatDate(rawMulai, 'Asia/Jakarta', 'yyyy-MM-dd')
      : String(rawMulai || '').trim();

    const selesaiStr = rawSelesai instanceof Date
      ? Utilities.formatDate(rawSelesai, 'Asia/Jakarta', 'yyyy-MM-dd')
      : String(rawSelesai || '').trim();

    const finalDept = applicantDb && applicantDb.dept ? applicantDb.dept : rawRowDept;
    const finalName = applicantDb && applicantDb.fullname ? applicantDb.fullname : (nameIdx !== -1 && row[nameIdx] ? String(row[nameIdx]).trim() : String(row[3] || ''));
    const finalJab = applicantDb && applicantDb.jabatan ? applicantDb.jabatan : (jabIdx !== -1 && row[jabIdx] ? String(row[jabIdx]).trim() : String(row[5] || ''));

    // Resolusi Titik Drop Off / Point of Hire
    let valTitik = '';
    if (titikDropOffIdx !== -1 && row[titikDropOffIdx] !== undefined && row[titikDropOffIdx] !== null) {
      valTitik = String(row[titikDropOffIdx]).trim();
    }
    if (!valTitik && row[36] !== undefined && row[36] !== null) {
      valTitik = String(row[36]).trim();
    }

    // Resolusi Moda Transportasi
    let valModa = '';
    if (modaTransportIdx !== -1 && row[modaTransportIdx] !== undefined && row[modaTransportIdx] !== null) {
      valModa = String(row[modaTransportIdx]).trim();
    }
    if (!valModa && row[37] !== undefined && row[37] !== null) {
      valModa = String(row[37]).trim();
    }

    // Resolusi Catatan Logistik / Tiket
    let valCatLog = '';
    if (catatanLogistikIdx !== -1 && row[catatanLogistikIdx] !== undefined && row[catatanLogistikIdx] !== null) {
      valCatLog = String(row[catatanLogistikIdx]).trim();
    }
    if (!valCatLog && row[38] !== undefined && row[38] !== null) {
      valCatLog = String(row[38]).trim();
    }

    // Heuristik jika valModa atau valTitik masih kosong (misal urutan kolom bergeser)
    if (!valModa || !valTitik) {
      const transportKeywords = ['travel', 'pesawat', 'speedboat', 'bus', 'elf', 'mandiri', 'kapal', 'mobil'];
      for (let c = Math.max(0, row.length - 8); c < row.length; c++) {
        const cellStr = String(row[c] || '').trim();
        if (!cellStr) continue;
        const cellLower = cellStr.toLowerCase();
        if (!valModa && transportKeywords.some(function (kw) { return cellLower.indexOf(kw) !== -1; })) {
          valModa = cellStr;
        } else if (!valTitik && c !== modaTransportIdx && c !== catatanLogistikIdx &&
                   !cellLower.includes('disetujui') && !cellLower.includes('ditolak') && !cellLower.includes('menunggu') &&
                   isNaN(Number(cellStr))) {
          if (cellLower.includes('samarinda') || cellLower.includes('balikpapan') || cellLower.includes('mess') ||
              cellLower.includes('site') || cellLower.includes('camp') || cellLower.includes('bandara') ||
              cellLower.includes('berau') || cellLower.includes('jakarta')) {
            valTitik = cellStr;
          }
        }
      }
    }

    list.push({
      id: idIdx !== -1 && row[idIdx] ? String(row[idIdx]).trim() : String(row[0] || ('CUTI-' + i)),
      tanggalPengajuan: tglAjuStr,
      empId: empIdIdx !== -1 && row[empIdIdx] ? String(row[empIdIdx] || '').trim() : String(row[2] || ''),
      namaKaryawan: finalName,
      applicantRole: applicantDb && applicantDb.role ? applicantDb.role : '',
      dept: finalDept,
      jabatan: finalJab,
      jenisCuti: jenisIdx !== -1 && row[jenisIdx] ? String(row[jenisIdx]).trim() : String(row[6] || ''),
      tanggalMulai: mulaiStr,
      tanggalSelesai: selesaiStr,
      jumlahHari: hariIdx !== -1 && Number(row[hariIdx]) > 0 ? Number(row[hariIdx]) : Number(row[9] || 0),
      alasan: alasanIdx !== -1 && row[alasanIdx] ? String(row[alasanIdx]).trim() : String(row[10] || ''),
      alamat: alamatIdx !== -1 && row[alamatIdx] ? String(row[alamatIdx]).trim() : String(row[11] || ''),
      kontak: kontakIdx !== -1 && row[kontakIdx] ? String(row[kontakIdx]).trim() : String(row[12] || ''),
      lampiranName: lampNameIdx !== -1 && row[lampNameIdx] ? String(row[lampNameIdx]).trim() : String(row[13] || ''),
      lampiranUrl: driveUrlIdx !== -1 && row[driveUrlIdx] ? String(row[driveUrlIdx]).trim() : String(row[14] || ''),
      lampiranFileId: driveIdIdx !== -1 && row[driveIdIdx] ? String(row[driveIdIdx]).trim() : String(row[15] || ''),
      catatanTambahan: catatanIdx !== -1 && row[catatanIdx] ? String(row[catatanIdx]).trim() : String(row[16] || ''),
      status: statusIdx !== -1 && row[statusIdx] ? String(row[statusIdx]).trim() : String(row[17] || 'Menunggu Persetujuan Atasan'),

      // Approver 1 (Atasan)
      approver1Nama: app1NameIdx !== -1 ? String(row[app1NameIdx] || '').trim() : '',
      approver1Jabatan: app1JabIdx !== -1 ? String(row[app1JabIdx] || '').trim() : '',
      approver1Tanggal: app1TglIdx !== -1 ? (row[app1TglIdx] instanceof Date ? Utilities.formatDate(row[app1TglIdx], 'Asia/Jakarta', 'dd/MM/yyyy HH:mm') : String(row[app1TglIdx] || '').trim()) : '',
      approver1Status: app1StatusIdx !== -1 ? String(row[app1StatusIdx] || '').trim() : '',
      approver1Catatan: app1CatatanIdx !== -1 ? String(row[app1CatatanIdx] || '').trim() : '',

      // Approver 2 (HRD)
      approver2Nama: app2NameIdx !== -1 ? String(row[app2NameIdx] || '').trim() : '',
      approver2Jabatan: app2JabIdx !== -1 ? String(row[app2JabIdx] || '').trim() : '',
      approver2Tanggal: app2TglIdx !== -1 ? (row[app2TglIdx] instanceof Date ? Utilities.formatDate(row[app2TglIdx], 'Asia/Jakarta', 'dd/MM/yyyy HH:mm') : String(row[app2TglIdx] || '').trim()) : '',
      approver2Status: app2StatusIdx !== -1 ? String(row[app2StatusIdx] || '').trim() : '',
      approver2Catatan: app2CatatanIdx !== -1 ? String(row[app2CatatanIdx] || '').trim() : '',

      // Approver 3 (PJO / Direktur)
      approver3Nama: app3NameIdx !== -1 ? String(row[app3NameIdx] || '').trim() : '',
      approver3Jabatan: app3JabIdx !== -1 ? String(row[app3JabIdx] || '').trim() : '',
      approver3Tanggal: app3TglIdx !== -1 ? (row[app3TglIdx] instanceof Date ? Utilities.formatDate(row[app3TglIdx], 'Asia/Jakarta', 'dd/MM/yyyy HH:mm') : String(row[app3TglIdx] || '').trim()) : '',
      approver3Status: app3StatusIdx !== -1 ? String(row[app3StatusIdx] || '').trim() : '',
      approver3Catatan: app3CatatanIdx !== -1 ? String(row[app3CatatanIdx] || '').trim() : '',

      alasanPenolakan: tolakAlasanIdx !== -1 ? String(row[tolakAlasanIdx] || '').trim() : '',
      tglAwalOnsite: tglAwalOnsiteIdx !== -1 && row[tglAwalOnsiteIdx] ? (row[tglAwalOnsiteIdx] instanceof Date ? Utilities.formatDate(row[tglAwalOnsiteIdx], 'Asia/Jakarta', 'yyyy-MM-dd') : String(row[tglAwalOnsiteIdx]).trim()) : '',
      totalHariKerjaOnsite: totalHariOnsiteIdx !== -1 && row[totalHariOnsiteIdx] !== '' && !isNaN(Number(row[totalHariOnsiteIdx])) ? Number(row[totalHariOnsiteIdx]) : null,
      titikDropOff: valTitik,
      pointOfHire: valTitik,
      pointOffHire: valTitik,
      dropOff: valTitik,
      modaTransportasi: valModa,
      transportasi: valModa,
      catatanLogistik: valCatLog,
      catatanTransportasi: valCatLog,
      lokasiKerja: lokasiKerjaIdx !== -1 && row[lokasiKerjaIdx] ? String(row[lokasiKerjaIdx]).trim() : (row[39] ? String(row[39]).trim() : 'TRIOP')
    });
  }

  list.reverse();
  return list;
}

function getLampiranCutiStorageFolder(tahun, empName) {
  const rootName = 'Aplikasi_Slip_Gaji_Storage';
  let root;
  const roots = DriveApp.getFoldersByName(rootName);
  if (roots.hasNext()) {
    root = roots.next();
  } else {
    root = DriveApp.createFolder(rootName);
  }

  const lampiranRoot = getOrCreateFolder(root, 'Lampiran_Cuti');
  const yearFolder = getOrCreateFolder(lampiranRoot, String(tahun || new Date().getFullYear()));
  const empFolder = getOrCreateFolder(yearFolder, empName || 'Karyawan');
  return empFolder;
}

function submitPengajuanCuti(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  payload = payload || {};
  const jenisCuti = String(payload.jenisCuti || '').trim();
  const tanggalMulai = String(payload.tanggalMulai || '').trim();
  const tanggalSelesai = String(payload.tanggalSelesai || '').trim();
  const alasanCuti = String(payload.alasanCuti || '').trim();
  const alamatCuti = String(payload.alamatCuti || '').trim();
  const nomorKontak = String(payload.nomorKontak || '').trim();
  const catatanTambahan = String(payload.catatanTambahan || '').trim();
  const tglAwalOnsite = String(payload.tglAwalOnsite || '').trim();
  const totalHariKerjaOnsite = (payload.totalHariKerjaOnsite !== undefined && payload.totalHariKerjaOnsite !== null && payload.totalHariKerjaOnsite !== '') ? Number(payload.totalHariKerjaOnsite) : '';

  if (!jenisCuti) return { success: false, message: 'Jenis cuti wajib dipilih.' };
  if (!tanggalMulai || !tanggalSelesai) return { success: false, message: 'Tanggal mulai dan tanggal selesai cuti wajib diisi.' };
  if (!alasanCuti) return { success: false, message: 'Alasan cuti wajib diisi.' };

  // Hitung jumlah hari kalender
  const dStart = new Date(tanggalMulai + 'T00:00:00');
  const dEnd = new Date(tanggalSelesai + 'T00:00:00');
  if (isNaN(dStart.getTime()) || isNaN(dEnd.getTime())) {
    return { success: false, message: 'Format tanggal cuti tidak valid.' };
  }
  if (dEnd < dStart) {
    return { success: false, message: 'Tanggal selesai tidak boleh lebih awal dari tanggal mulai.' };
  }
  const diffTime = Math.abs(dEnd.getTime() - dStart.getTime());
  const jumlahHari = Math.round(diffTime / (1000 * 60 * 60 * 24)) + 1;

  // Ambil data profil karyawan secara konsisten dari sheet Users
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  const userData = userSheet ? userSheet.getDataRange().getValues() : [];
  let userFullname = session.user.fullname || '';
  let userJabatan = session.user.jabatan || '';
  let userDept = session.user.dept || '';
  const empId = String(session.user.username || '').trim();

  for (let i = 1; i < userData.length; i++) {
    if (String(userData[i][0]).trim().toLowerCase() === empId.toLowerCase()) {
      userFullname = String(userData[i][1] || userFullname);
      userJabatan = String(userData[i][2] || userJabatan);
      userDept = String(userData[i][3] || userDept);
      break;
    }
  }

  // Simpan Lampiran / Bukti ke Google Drive jika ada
  let lampiranName = '';
  let lampiranUrl = '';
  let lampiranFileId = '';

  if (payload.lampiranBase64 && payload.lampiranName) {
    try {
      const b64Data = payload.lampiranBase64.replace(/^data:.*?;base64,/, '');
      const fileSize = getBase64FileSize(b64Data);
      if (fileSize > 5242880) { // Maksimal 5 MB
        return { success: false, message: 'Ukuran lampiran bukti maksimal 5 MB.' };
      }
      const fileBytes = Utilities.base64Decode(b64Data);
      const mimeType = payload.lampiranMimeType || 'application/octet-stream';
      const cleanFileName = String(payload.lampiranName).replace(/[^a-zA-Z0-9._-]/g, '_');
      const blob = Utilities.newBlob(fileBytes, mimeType, 'Bukti_Cuti_' + cleanFileName);

      const tahun = tanggalMulai.substring(0, 4) || String(new Date().getFullYear());
      const folder = getLampiranCutiStorageFolder(tahun, userFullname);
      const savedFile = folder.createFile(blob);
      lampiranName = savedFile.getName();
      lampiranUrl = savedFile.getUrl();
      lampiranFileId = savedFile.getId();
    } catch (fErr) {
      return { success: false, message: 'Gagal mengunggah berkas lampiran: ' + fErr.toString() };
    }
  }

  const cutiSheet = getOrCreatePengajuanCutiSheet(ss);
  const now = new Date();
  const yearMonth = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMM');
  const cutiId = 'CUTI-' + yearMonth + '-' + Math.floor(1000 + Math.random() * 9000);
  const formattedDate = Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');

  let jenisCutiFinal = jenisCuti;
  if (jenisCuti === 'Cuti Roster' && payload.polaRoster) {
    jenisCutiFinal = 'Cuti Roster (' + payload.polaRoster + ')';
  }

  const submitterRole = String(session.user.role || '').trim().toLowerCase();
  const submitterJab = String(session.user.jabatan || '').trim().toLowerCase();
  const isSubmitterPJO = isJobPJO(submitterRole, submitterJab);
  const isSubmitterHRD = (submitterRole === 'hrd' || submitterRole === 'admin');
  const isSubmitterAtasan = (
    submitterRole === 'supervisor' ||
    submitterRole === 'atasan' ||
    isJobAtasan(submitterRole, submitterJab)
  );

  let initialCutiStatus = 'Menunggu Persetujuan Atasan';
  if (isSubmitterPJO) {
    initialCutiStatus = 'Menunggu Persetujuan HRD';
  } else if (isSubmitterHRD) {
    initialCutiStatus = 'Menunggu Persetujuan PJO';
  } else if (isSubmitterAtasan) {
    initialCutiStatus = 'Menunggu Persetujuan HRD';
  }

  const titikDropOff = String(payload.titikDropOff || payload.pointOfHire || payload.pointOffHire || '').trim();
  const modaTransportasi = String(payload.modaTransportasi || payload.transportasi || '').trim();
  const catatanLogistik = String(payload.catatanLogistik || payload.catatanTransportasi || '').trim();
  const lokasiKerja = String(payload.lokasiKerja || payload.jobsite || 'TRIOP').trim() || 'TRIOP';

  // Dapatkan header sheet jika sudah ada kolom dinamis
  const lastCol = cutiSheet.getLastColumn();
  let headers = [];
  if (lastCol > 0) {
    headers = cutiSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
      return String(h || '').trim().toLowerCase();
    });
  }

  const dropIdx = headers.findIndex(function (h) {
    return h.includes('titik drop') || h.includes('point of hire') || h.includes('point off hire') ||
           h.includes('drop') || h.includes('hire') || h.includes('point') || h.includes('poh') ||
           h.includes('jemput') || h.includes('antar') || h.includes('tujuan');
  });
  const modaIdx = headers.findIndex(function (h) {
    return (h.includes('transport') || h.includes('moda') || h.includes('kendaraan') || h.includes('armada')) &&
           !h.includes('catatan') && !h.includes('note');
  });
  const logIdx = headers.findIndex(function (h) {
    return h.includes('logistik') || h.includes('tiket') || h.includes('penjemputan') ||
           (h.includes('catatan') && (h.includes('transport') || h.includes('khusus') || h.includes('tiket') || h.includes('jemput') || h.includes('drop')));
  });
  const lokasiIdx = headers.findIndex(function (h) {
    return h.includes('lokasi') || h.includes('jobsite') || h === 'site';
  });

  const rowData = [
    cutiId,
    formattedDate,
    empId,
    userFullname,
    userDept,
    userJabatan,
    jenisCutiFinal,
    tanggalMulai,
    tanggalSelesai,
    jumlahHari,
    alasanCuti,
    alamatCuti,
    nomorKontak,
    lampiranName,
    lampiranUrl,
    lampiranFileId,
    catatanTambahan,
    initialCutiStatus,
    '', '', '', '', '', // Approver 1
    '', '', '', '', '', // Approver 2
    '', '', '', '', '', // Approver 3
    '', // Alasan Penolakan
    tglAwalOnsite,
    totalHariKerjaOnsite,
    titikDropOff,
    modaTransportasi,
    catatanLogistik,
    lokasiKerja
  ];

  if (headers.length > rowData.length) {
    while (rowData.length < headers.length) {
      rowData.push('');
    }
  }

  if (dropIdx !== -1) rowData[dropIdx] = titikDropOff;
  if (modaIdx !== -1) rowData[modaIdx] = modaTransportasi;
  if (logIdx !== -1) rowData[logIdx] = catatanLogistik;
  if (lokasiIdx !== -1) rowData[lokasiIdx] = lokasiKerja;

  cutiSheet.appendRow(rowData);

  return {
    success: true,
    message: 'Pengajuan cuti berhasil dikirim! (' + jumlahHari + ' Hari)',
    cutiId: cutiId,
    jumlahHari: jumlahHari
  };
}

function editPengajuanCuti(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  payload = payload || {};
  const cutiId = String(payload.cutiId || '').trim();
  if (!cutiId) return { success: false, message: 'ID pengajuan cuti tidak valid.' };

  const jenisCuti = String(payload.jenisCuti || '').trim();
  const tanggalMulai = String(payload.tanggalMulai || '').trim();
  const tanggalSelesai = String(payload.tanggalSelesai || '').trim();
  const alasanCuti = String(payload.alasanCuti || '').trim();
  const alamatCuti = String(payload.alamatCuti || '').trim();
  const nomorKontak = String(payload.nomorKontak || '').trim();
  const catatanTambahan = String(payload.catatanTambahan || '').trim();
  const tglAwalOnsite = String(payload.tglAwalOnsite || '').trim();
  const totalHariKerjaOnsite = (payload.totalHariKerjaOnsite !== undefined && payload.totalHariKerjaOnsite !== null && payload.totalHariKerjaOnsite !== '') ? Number(payload.totalHariKerjaOnsite) : '';
  const titikDropOff = String(payload.titikDropOff || '').trim();
  const modaTransportasi = String(payload.modaTransportasi || '').trim();
  const catatanLogistik = String(payload.catatanLogistik || '').trim();

  if (!jenisCuti) return { success: false, message: 'Jenis cuti wajib dipilih.' };
  if (!tanggalMulai || !tanggalSelesai) return { success: false, message: 'Tanggal mulai dan tanggal selesai cuti wajib diisi.' };
  if (!alasanCuti) return { success: false, message: 'Alasan cuti wajib diisi.' };

  // Hitung jumlah hari kalender
  const dStart = new Date(tanggalMulai + 'T00:00:00');
  const dEnd = new Date(tanggalSelesai + 'T00:00:00');
  if (isNaN(dStart.getTime()) || isNaN(dEnd.getTime())) {
    return { success: false, message: 'Format tanggal cuti tidak valid.' };
  }
  if (dEnd < dStart) {
    return { success: false, message: 'Tanggal selesai tidak boleh lebih awal dari tanggal mulai.' };
  }
  const diffTime = Math.abs(dEnd.getTime() - dStart.getTime());
  const jumlahHari = Math.round(diffTime / (1000 * 60 * 60 * 24)) + 1;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cutiSheet = getOrCreatePengajuanCutiSheet(ss);
  const data = cutiSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { success: false, message: 'Data pengajuan cuti tidak ditemukan.' };

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id pengajuan') || h === 'id cuti' || h === 'id' || h.includes('kode'); });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h.includes('nip') || h === 'username' || h === 'id_karyawan'; });
  const statusIdx = headers.findIndex(function (h) { return h.includes('status'); });
  const jenisIdx = headers.findIndex(function (h) { return h.includes('jenis cuti') || h === 'jenis'; });
  const tglMulaiIdx = headers.findIndex(function (h) { return h.includes('tanggal mulai') || h.includes('tgl mulai'); });
  const tglSelesaiIdx = headers.findIndex(function (h) { return h.includes('tanggal selesai') || h.includes('tgl selesai'); });
  const jmlHariIdx = headers.findIndex(function (h) { return h.includes('jumlah hari') || h.includes('durasi'); });
  const alasanIdx = headers.findIndex(function (h) { return h.includes('alasan'); });
  const alamatIdx = headers.findIndex(function (h) { return h.includes('alamat'); });
  const kontakIdx = headers.findIndex(function (h) { return h.includes('kontak') || h.includes('telepon') || h.includes('hp'); });
  const lampiranNameIdx = headers.findIndex(function (h) { return h.includes('nama file lampiran') || h.includes('file lampiran'); });
  const lampiranUrlIdx = headers.findIndex(function (h) { return h.includes('url drive') || h.includes('url lampiran'); });
  const lampiranFileIdIdx = headers.findIndex(function (h) { return h.includes('file id') || h.includes('fileid'); });
  const catatanIdx = headers.findIndex(function (h) { return h.includes('catatan') && !h.includes('approver') && !h.includes('logistik') && !h.includes('transport') && !h.includes('tiket') && !h.includes('jemput'); });
  const app1StatusIdx = headers.findIndex(function (h) { return h.includes('approver 1 status'); });
  const tglAwalOnsiteIdx = headers.findIndex(function (h) { return h.includes('tanggal awal onsite') || h.includes('tgl awal onsite'); });
  const totalHariOnsiteIdx = headers.findIndex(function (h) { return h.includes('total hari kerja onsite') || h.includes('total hari onsite'); });
  const titikDropOffIdx = headers.findIndex(function (h) {
    return h.includes('titik drop') || h.includes('point of hire') || h.includes('point off hire') ||
           h.includes('drop') || h.includes('hire') || h.includes('point') || h.includes('poh') ||
           h.includes('jemput') || h.includes('antar') || h.includes('tujuan');
  });
  const modaTransportIdx = headers.findIndex(function (h) {
    return (h.includes('transport') || h.includes('moda') || h.includes('kendaraan') || h.includes('armada')) &&
           !h.includes('catatan') && !h.includes('note');
  });
  const catatanLogistikIdx = headers.findIndex(function (h) {
    return h.includes('logistik') || h.includes('tiket') || h.includes('penjemputan') ||
           (h.includes('catatan') && (h.includes('transport') || h.includes('khusus') || h.includes('tiket') || h.includes('jemput') || h.includes('drop')));
  });
  const lokasiKerjaIdx = headers.findIndex(function (h) {
    return h.includes('lokasi') || h.includes('jobsite') || h === 'site';
  });

  const currentUsername = String(session.user.username || '').trim().toLowerCase();
  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isAdmin = (userRole === 'admin' || userRole === 'admin master' || currentUsername === 'admin');

  for (let i = 1; i < data.length; i++) {
    const rowId = idIdx !== -1 ? String(data[i][idIdx] || '').trim() : String(data[i][0] || '').trim();
    if (rowId === cutiId) {
      const rowEmpId = empIdIdx !== -1 ? String(data[i][empIdIdx] || '').trim().toLowerCase() : '';
      const currentStatus = statusIdx !== -1 ? String(data[i][statusIdx] || '').trim() : '';
      const app1Status = app1StatusIdx !== -1 ? String(data[i][app1StatusIdx] || '').trim().toLowerCase() : '';

      // 1. Validasi Kepemilikan (Karyawan hanya bisa edit pengajuannya sendiri, kecuali admin)
      if (!isAdmin && rowEmpId !== currentUsername) {
        return { success: false, message: 'Akses ditolak: Anda tidak memiliki wewenang untuk mengubah pengajuan cuti ini.' };
      }

      // 2. Validasi Status Approval:
      // Karyawan hanya bisa edit SEBELUM diapprove oleh atasan.
      // Ketika sudah diapprove oleh atasan sampai PJO, karyawan TIDAK BISA edit.
      if (!isAdmin) {
        if (app1Status === 'disetujui' || app1Status === 'approved') {
          return { success: false, message: 'Pengajuan cuti ini sudah disetujui oleh atasan dan tidak dapat diubah lagi oleh karyawan.' };
        }
        if (currentStatus === 'Menunggu Persetujuan HRD' || currentStatus === 'Menunggu Persetujuan PJO' || currentStatus === 'Menunggu Persetujuan Direktur' || currentStatus === 'Disetujui') {
          return { success: false, message: 'Pengajuan cuti ini sudah diproses/disetujui tahap atasan ke atas dan tidak dapat diubah lagi.' };
        }
        if (currentStatus.toLowerCase().indexOf('tolak') !== -1) {
          return { success: false, message: 'Pengajuan cuti yang telah ditolak tidak dapat diubah. Silakan buat pengajuan cuti baru.' };
        }
      }

      const rowIndex = i + 1;

      // Format jenis cuti jika cuti roster
      let jenisCutiFinal = jenisCuti;
      if (jenisCuti === 'Cuti Roster' && payload.polaRoster) {
        jenisCutiFinal = 'Cuti Roster (' + payload.polaRoster + ')';
      }

      // Upload berkas lampiran baru jika ada
      if (payload.lampiranBase64 && payload.lampiranName) {
        try {
          const b64Data = payload.lampiranBase64.replace(/^data:.*?;base64,/, '');
          const fileSize = getBase64FileSize(b64Data);
          if (fileSize > 5242880) {
            return { success: false, message: 'Ukuran lampiran bukti maksimal 5 MB.' };
          }
          const fileBytes = Utilities.base64Decode(b64Data);
          const mimeType = payload.lampiranMimeType || 'application/octet-stream';
          const cleanFileName = String(payload.lampiranName).replace(/[^a-zA-Z0-9._-]/g, '_');
          const blob = Utilities.newBlob(fileBytes, mimeType, 'Bukti_Cuti_' + cleanFileName);

          const tahun = tanggalMulai.substring(0, 4) || String(new Date().getFullYear());
          const empName = String(data[i][3] || session.user.fullname || 'Karyawan');
          const folder = getLampiranCutiStorageFolder(tahun, empName);
          const savedFile = folder.createFile(blob);

          // Hapus file lampiran lama jika ada
          const oldFileId = lampiranFileIdIdx !== -1 ? String(data[i][lampiranFileIdIdx] || '').trim() : '';
          if (oldFileId) {
            try { DriveApp.getFileById(oldFileId).setTrashed(true); } catch (eIgnore) { }
          }

          if (lampiranNameIdx !== -1) cutiSheet.getRange(rowIndex, lampiranNameIdx + 1).setValue(savedFile.getName());
          if (lampiranUrlIdx !== -1) cutiSheet.getRange(rowIndex, lampiranUrlIdx + 1).setValue(savedFile.getUrl());
          if (lampiranFileIdIdx !== -1) cutiSheet.getRange(rowIndex, lampiranFileIdIdx + 1).setValue(savedFile.getId());
        } catch (fErr) {
          return { success: false, message: 'Gagal mengunggah berkas lampiran baru: ' + fErr.toString() };
        }
      }

      // Update nilai kolom pada Google Sheet
      if (jenisIdx !== -1) cutiSheet.getRange(rowIndex, jenisIdx + 1).setValue(jenisCutiFinal);
      if (tglMulaiIdx !== -1) cutiSheet.getRange(rowIndex, tglMulaiIdx + 1).setValue(tanggalMulai);
      if (tglSelesaiIdx !== -1) cutiSheet.getRange(rowIndex, tglSelesaiIdx + 1).setValue(tanggalSelesai);
      if (jmlHariIdx !== -1) cutiSheet.getRange(rowIndex, jmlHariIdx + 1).setValue(jumlahHari);
      if (alasanIdx !== -1) cutiSheet.getRange(rowIndex, alasanIdx + 1).setValue(alasanCuti);
      if (alamatIdx !== -1) cutiSheet.getRange(rowIndex, alamatIdx + 1).setValue(alamatCuti);
      if (kontakIdx !== -1) cutiSheet.getRange(rowIndex, kontakIdx + 1).setValue(nomorKontak);
      if (catatanIdx !== -1) cutiSheet.getRange(rowIndex, catatanIdx + 1).setValue(catatanTambahan);
      if (tglAwalOnsiteIdx !== -1) cutiSheet.getRange(rowIndex, tglAwalOnsiteIdx + 1).setValue(tglAwalOnsite);
      if (totalHariOnsiteIdx !== -1) cutiSheet.getRange(rowIndex, totalHariOnsiteIdx + 1).setValue(totalHariKerjaOnsite);
      const colTitik = titikDropOffIdx !== -1 ? (titikDropOffIdx + 1) : 37;
      const colModa = modaTransportIdx !== -1 ? (modaTransportIdx + 1) : 38;
      const colCatLog = catatanLogistikIdx !== -1 ? (catatanLogistikIdx + 1) : 39;
      const colLokasi = lokasiKerjaIdx !== -1 ? (lokasiKerjaIdx + 1) : 40;

      cutiSheet.getRange(rowIndex, colTitik).setValue(titikDropOff);
      cutiSheet.getRange(rowIndex, colModa).setValue(modaTransportasi);
      cutiSheet.getRange(rowIndex, colCatLog).setValue(catatanLogistik);
      cutiSheet.getRange(rowIndex, colLokasi).setValue(String(payload.lokasiKerja || payload.jobsite || 'TRIOP').trim() || 'TRIOP');

      return {
        success: true,
        message: 'Pengajuan cuti ' + cutiId + ' berhasil diperbarui (' + jumlahHari + ' Hari Kalender).',
        cutiId: cutiId,
        jumlahHari: jumlahHari
      };
    }
  }

  return { success: false, message: 'Data pengajuan cuti tidak ditemukan.' };
}

function getCutiKaryawan(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cutiSheet = getOrCreatePengajuanCutiSheet(ss);

  // Ambil profil data user terbaru dari sheet Users untuk menjamin role & dept selalu akurat
  const userSheet = ss.getSheetByName('Users');
  const uData = userSheet ? userSheet.getDataRange().getValues() : [];
  let freshUser = session.user;
  const targetU = String(session.user.username || '').trim().toLowerCase();
  for (let i = 1; i < uData.length; i++) {
    if (String(uData[i][0] || '').trim().toLowerCase() === targetU) {
      const tglMulai = uData[i][9] instanceof Date
        ? Utilities.formatDate(uData[i][9], 'Asia/Jakarta', 'yyyy-MM-dd')
        : String(uData[i][9] || '').trim();
      freshUser = {
        username: String(uData[i][0] || '').trim(),
        fullname: String(uData[i][1] || '').trim(),
        jabatan: String(uData[i][2] || '').trim(),
        dept: String(uData[i][3] || '').trim(),
        role: String(uData[i][5] || '').trim(),
        email: String(uData[i][6] || '').trim(),
        polaRoster: String(uData[i][8] || 'Non-Roster').trim() || 'Non-Roster',
        tglMulaiKerja: tglMulai
      };
      break;
    }
  }

  const list = parseCutiDataFromSheet(cutiSheet, freshUser);
  const freshRole = String(freshUser.role || '').trim().toLowerCase();
  const isHRD = (freshRole === 'admin' || freshRole === 'hrd' || cleanDept(freshUser.dept) === 'hrd');
  const isSupervisor = isJobAtasan(freshUser.role, freshUser.jabatan);
  let rosterData = [];
  try {
    const rRes = getRosterList(token);
    if (rRes && rRes.success && rRes.roster) rosterData = rRes.roster;
  } catch (e) { }

  return {
    success: true,
    list: list,
    cuti: list,
    roster: rosterData,
    currentUserProfile: freshUser,
    debugInfo: {
      username: targetU,
      role: freshUser.role,
      dept: freshUser.dept,
      jabatan: freshUser.jabatan,
      isSupervisor: isSupervisor,
      isHRD: isHRD,
      totalRowsInSheet: cutiSheet ? Math.max(0, cutiSheet.getLastRow() - 1) : 0,
      matchedRows: list.length
    }
  };
}

function deletePengajuanCuti(token, cutiId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  cutiId = String(cutiId || '').trim();
  if (!cutiId) return { success: false, message: 'ID pengajuan cuti tidak valid.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cutiSheet = getOrCreatePengajuanCutiSheet(ss);
  const data = cutiSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { success: false, message: 'Data pengajuan cuti tidak ditemukan.' };

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id pengajuan') || h === 'id cuti' || h === 'id' || h.includes('kode'); });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h.includes('nip') || h === 'username' || h === 'id_karyawan'; });
  const statusIdx = headers.findIndex(function (h) { return h.includes('status'); });
  const fileIdIdx = headers.findIndex(function (h) { return h.includes('file id') || h.includes('fileid'); });

  const currentUsername = String(session.user.username || '').trim().toLowerCase();

  // Ambil profil data user terbaru dari sheet Users untuk menjamin role akurat
  const userSheet = ss.getSheetByName('Users');
  const uData = userSheet ? userSheet.getDataRange().getValues() : [];
  let userRole = String(session.user.role || '').trim().toLowerCase();
  let userDept = String(session.user.dept || '').trim().toLowerCase();
  for (let u = 1; u < uData.length; u++) {
    if (String(uData[u][0] || '').trim().toLowerCase() === currentUsername) {
      userRole = String(uData[u][5] || userRole).trim().toLowerCase();
      userDept = String(uData[u][3] || userDept).trim().toLowerCase();
      break;
    }
  }

  const isHRD = (userRole === 'admin' || userRole === 'hrd' || currentUsername === 'admin' || cleanDept(userDept) === 'hrd');

  for (let i = 1; i < data.length; i++) {
    const rowId = idIdx !== -1 ? String(data[i][idIdx] || '').trim() : String(data[i][0] || '').trim();
    if (rowId === cutiId) {
      const rowEmpId = empIdIdx !== -1 ? String(data[i][empIdIdx] || '').trim().toLowerCase() : String(data[i][2] || '').trim().toLowerCase();
      const status = statusIdx !== -1 ? String(data[i][statusIdx] || '').trim() : String(data[i][17] || '').trim();
      const statusLower = status.toLowerCase();

      if (!isHRD && rowEmpId !== currentUsername) {
        return { success: false, message: 'Anda tidak memiliki hak untuk menghapus pengajuan ini.' };
      }

      const isPending = statusLower.includes('menunggu');
      const isRejected = statusLower.includes('tolak');
      const isCancelled = statusLower.includes('batal');

      if (!isHRD && !isPending && !isRejected && !isCancelled) {
        return { success: false, message: 'Pengajuan cuti yang sudah disetujui hanya dapat dihapus oleh HRD / Admin.' };
      }

      // Hapus file lampiran jika ada
      const fileId = fileIdIdx !== -1 ? String(data[i][fileIdIdx] || '').trim() : String(data[i][15] || '').trim();
      if (fileId) {
        try {
          DriveApp.getFileById(fileId).setTrashed(true);
        } catch (delErr) {
          console.error('Lampiran cuti delete note:', delErr);
        }
      }

      cutiSheet.deleteRow(i + 1);
      return { success: true, message: 'Data pengajuan cuti ' + cutiId + ' berhasil dihapus.' };
    }
  }

  return { success: false, message: 'Data pengajuan cuti tidak ditemukan.' };
}

function processApprovalCuti(token, cutiId, action, catatan) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  cutiId = String(cutiId || '').trim();
  action = String(action || '').trim(); // 'Setujui' or 'Tolak'
  catatan = String(catatan || '').trim();

  if (!cutiId) return { success: false, message: 'ID Pengajuan tidak valid.' };
  if (action !== 'Setujui' && action !== 'Tolak') {
    return { success: false, message: 'Aksi approval tidak valid (harus Setujui atau Tolak).' };
  }
  if (action === 'Tolak' && !catatan) {
    return { success: false, message: 'Alasan penolakan wajib diisi.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cutiSheet = getOrCreatePengajuanCutiSheet(ss);
  const data = cutiSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { success: false, message: 'Data pengajuan cuti tidak ditemukan.' };

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id pengajuan') || h === 'id cuti' || h === 'id' || h.includes('kode'); });
  const statusIdx = headers.findIndex(function (h) { return h.includes('status'); });
  const deptIdx = headers.findIndex(function (h) { return h.includes('departemen') || h.includes('dept'); });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h.includes('nip') || h === 'username' || h === 'id_karyawan'; });
  const jabIdx = headers.findIndex(function (h) { return h.includes('jabatan'); });

  // Indexes Approver 1 (Atasan)
  const app1NameIdx = headers.findIndex(function (h) { return h.includes('approver 1 nama') || h.includes('approver1 nama') || h === 'approver 1'; });
  const app1JabIdx = headers.findIndex(function (h) { return h.includes('approver 1 jabatan') || h.includes('approver1 jabatan'); });
  const app1TglIdx = headers.findIndex(function (h) { return h.includes('approver 1 tanggal') || h.includes('approver 1 tgl'); });
  const app1StatusIdx = headers.findIndex(function (h) { return h.includes('approver 1 status'); });
  const app1CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 1 catatan'); });

  // Indexes Approver 2 (HRD)
  const app2NameIdx = headers.findIndex(function (h) { return h.includes('approver 2 nama') || h.includes('approver2 nama') || h === 'approver 2'; });
  const app2JabIdx = headers.findIndex(function (h) { return h.includes('approver 2 jabatan') || h.includes('approver2 jabatan'); });
  const app2TglIdx = headers.findIndex(function (h) { return h.includes('approver 2 tanggal') || h.includes('approver 2 tgl'); });
  const app2StatusIdx = headers.findIndex(function (h) { return h.includes('approver 2 status'); });
  const app2CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 2 catatan'); });

  // Indexes Approver 3 (PJO / Direktur)
  const app3NameIdx = headers.findIndex(function (h) { return h.includes('approver 3 nama') || h.includes('approver3 nama') || h === 'approver 3'; });
  const app3JabIdx = headers.findIndex(function (h) { return h.includes('approver 3 jabatan') || h.includes('approver3 jabatan'); });
  const app3TglIdx = headers.findIndex(function (h) { return h.includes('approver 3 tanggal') || h.includes('approver 3 tgl'); });
  const app3StatusIdx = headers.findIndex(function (h) { return h.includes('approver 3 status'); });
  const app3CatatanIdx = headers.findIndex(function (h) { return h.includes('approver 3 catatan'); });

  const tolakAlasanIdx = headers.findIndex(function (h) { return h.includes('alasan penolakan') || h.includes('alasan tolak'); });

  const user = session.user;
  const userRole = String(user.role || '').trim().toLowerCase();
  const userDept = String(user.dept || '').trim().toLowerCase();
  const userDeptNorm = normalizeDept(user.dept);
  const userJab = String(user.jabatan || '').trim().toLowerCase();
  const currentUsername = String(user.username || '').trim().toLowerCase();

  const isAdmin = (userRole === 'admin' || userRole === 'admin master' || currentUsername === 'admin');
  const isHRD = (isAdmin || userRole === 'hrd' || userDept.includes('hr') || userDept.includes('human resource'));
  const isSupervisor = isJobAtasan(userRole, userJab);
  const isPJO = isJobPJO(userRole, userJab);
  const canApprovePJO = (isPJO || isAdmin);

  if (!isHRD && !isSupervisor && !canApprovePJO) {
    return { success: false, message: 'Akses ditolak: Anda tidak memiliki wewenang untuk melakukan approval cuti.' };
  }

  // Peta data karyawan dari sheet Users untuk memastikan role dan departemen pemohon
  const userSheet = ss.getSheetByName('Users');
  const userSheetData = userSheet ? userSheet.getDataRange().getValues() : [];
  const userMap = {};
  for (let u = 1; u < userSheetData.length; u++) {
    const uRow = userSheetData[u];
    const uId = String(uRow[0] || '').trim().toLowerCase();
    if (uId) {
      userMap[uId] = {
        id: String(uRow[0] || '').trim(),
        fullname: String(uRow[1] || '').trim(),
        jabatan: String(uRow[2] || '').trim(),
        dept: String(uRow[3] || '').trim(),
        role: String(uRow[5] || '').trim()
      };
    }
  }

  for (let i = 1; i < data.length; i++) {
    const rowId = idIdx !== -1 ? String(data[i][idIdx] || '').trim() : String(data[i][0] || '').trim();
    if (rowId === cutiId) {
      const rowEmpId = empIdIdx !== -1 ? String(data[i][empIdIdx] || '').trim().toLowerCase() : '';
      const rawRowDept = deptIdx !== -1 ? String(data[i][deptIdx] || '').trim() : '';
      const rawRowJab = jabIdx !== -1 ? String(data[i][jabIdx] || '').trim() : '';
      const currentStatus = statusIdx !== -1 ? String(data[i][statusIdx] || '').trim() : 'Menunggu Persetujuan Atasan';

      // Ground truth data pemohon dari Sheet Users
      const applicant = userMap[rowEmpId] || {
        id: rowEmpId,
        fullname: '',
        jabatan: rawRowJab,
        dept: rawRowDept,
        role: 'Karyawan'
      };

      const applicantRole = String(applicant.role || '').trim().toLowerCase();
      const applicantJab = String(applicant.jabatan || rawRowJab).trim().toLowerCase();
      const applicantDept = String(applicant.dept || rawRowDept).trim();

      // Cegah approval diri sendiri kecuali akun Admin
      if (rowEmpId === currentUsername && userRole !== 'admin') {
        return { success: false, message: 'Anda tidak dapat memproses persetujuan pengajuan cuti Anda sendiri.' };
      }

      const now = new Date();
      const nowStr = Utilities.formatDate(now, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
      const rowIndex = i + 1;

      // 1. TAHAP 1: Approval Atasan / Supervisor
      if (currentStatus === 'Menunggu Persetujuan Atasan' || currentStatus === 'Menunggu Persetujuan') {
        if (!isSupervisor && !isHRD) {
          return { success: false, message: 'Hanya Atasan/Supervisor (atau Admin/HRD) yang berhak memproses permohonan tahap ini.' };
        }

        // Atasan hanya bisa approve untuk karyawan yang sama departemen dengannya
        // HRD/Admin bisa approve semua departemen
        if (!isHRD) {
          const matchDept = isSameDept(userDept, applicantDept) || isSameDept(userDept, rawRowDept) || (userDeptNorm !== '' && (userDeptNorm === normalizeDept(applicantDept) || userDeptNorm === normalizeDept(rawRowDept)));
          if (!matchDept) {
            return {
              success: false,
              message: 'Akses ditolak: Anda hanya dapat memproses permohonan karyawan di departemen Anda (' + (user.dept || '-') + '). Departemen pemohon adalah ' + (applicantDept || rawRowDept || '-') + '.'
            };
          }
        }

        if (action === 'Setujui') {
          cutiSheet.getRange(rowIndex, statusIdx + 1).setValue('Menunggu Persetujuan HRD');
          if (app1NameIdx !== -1) cutiSheet.getRange(rowIndex, app1NameIdx + 1).setValue(user.fullname);
          if (app1JabIdx !== -1) cutiSheet.getRange(rowIndex, app1JabIdx + 1).setValue(user.jabatan);
          if (app1TglIdx !== -1) cutiSheet.getRange(rowIndex, app1TglIdx + 1).setValue(nowStr);
          if (app1StatusIdx !== -1) cutiSheet.getRange(rowIndex, app1StatusIdx + 1).setValue('Disetujui');
          if (app1CatatanIdx !== -1) cutiSheet.getRange(rowIndex, app1CatatanIdx + 1).setValue(catatan || 'Disetujui');
          return {
            success: true,
            message: 'Pengajuan cuti berhasil disetujui di tingkat Atasan dan diteruskan ke HRD.',
            newStatus: 'Menunggu Persetujuan HRD'
          };
        } else {
          cutiSheet.getRange(rowIndex, statusIdx + 1).setValue('Ditolak Atasan');
          if (app1NameIdx !== -1) cutiSheet.getRange(rowIndex, app1NameIdx + 1).setValue(user.fullname);
          if (app1JabIdx !== -1) cutiSheet.getRange(rowIndex, app1JabIdx + 1).setValue(user.jabatan);
          if (app1TglIdx !== -1) cutiSheet.getRange(rowIndex, app1TglIdx + 1).setValue(nowStr);
          if (app1StatusIdx !== -1) cutiSheet.getRange(rowIndex, app1StatusIdx + 1).setValue('Ditolak');
          if (app1CatatanIdx !== -1) cutiSheet.getRange(rowIndex, app1CatatanIdx + 1).setValue(catatan);
          if (tolakAlasanIdx !== -1) cutiSheet.getRange(rowIndex, tolakAlasanIdx + 1).setValue(catatan);
          return {
            success: true,
            message: 'Pengajuan cuti berhasil ditolak oleh Atasan.',
            newStatus: 'Ditolak Atasan'
          };
        }
      }

      // 2. TAHAP 2: Approval HRD
      else if (currentStatus === 'Menunggu Persetujuan HRD') {
        if (!isHRD) {
          return { success: false, message: 'Hanya HRD / Admin yang berhak memproses permohonan tahap verifikasi ini.' };
        }

        if (action === 'Setujui') {
          cutiSheet.getRange(rowIndex, statusIdx + 1).setValue('Menunggu Persetujuan PJO');
          if (app2NameIdx !== -1) cutiSheet.getRange(rowIndex, app2NameIdx + 1).setValue(user.fullname);
          if (app2JabIdx !== -1) cutiSheet.getRange(rowIndex, app2JabIdx + 1).setValue(user.jabatan);
          if (app2TglIdx !== -1) cutiSheet.getRange(rowIndex, app2TglIdx + 1).setValue(nowStr);
          if (app2StatusIdx !== -1) cutiSheet.getRange(rowIndex, app2StatusIdx + 1).setValue('Disetujui');
          if (app2CatatanIdx !== -1) cutiSheet.getRange(rowIndex, app2CatatanIdx + 1).setValue(catatan || 'Disetujui HRD');
          return {
            success: true,
            message: 'Pengajuan cuti berhasil diverifikasi oleh HRD dan diteruskan ke PJO / Direktur untuk pengesahan akhir.',
            newStatus: 'Menunggu Persetujuan PJO'
          };
        } else {
          cutiSheet.getRange(rowIndex, statusIdx + 1).setValue('Ditolak HRD');
          if (app2NameIdx !== -1) cutiSheet.getRange(rowIndex, app2NameIdx + 1).setValue(user.fullname);
          if (app2JabIdx !== -1) cutiSheet.getRange(rowIndex, app2JabIdx + 1).setValue(user.jabatan);
          if (app2TglIdx !== -1) cutiSheet.getRange(rowIndex, app2TglIdx + 1).setValue(nowStr);
          if (app2StatusIdx !== -1) cutiSheet.getRange(rowIndex, app2StatusIdx + 1).setValue('Ditolak');
          if (app2CatatanIdx !== -1) cutiSheet.getRange(rowIndex, app2CatatanIdx + 1).setValue(catatan);
          if (tolakAlasanIdx !== -1) cutiSheet.getRange(rowIndex, tolakAlasanIdx + 1).setValue(catatan);
          return {
            success: true,
            message: 'Pengajuan cuti telah ditolak oleh HRD.',
            newStatus: 'Ditolak HRD'
          };
        }
      }

      // 3. TAHAP 3: Approval PJO / Direktur (Final Sign-off)
      else if (currentStatus === 'Menunggu Persetujuan PJO' || currentStatus === 'Menunggu Persetujuan Direktur') {
        if (!canApprovePJO) {
          return { success: false, message: 'Akses ditolak: Hanya PJO / Direktur (atau Admin) yang berhak memproses persetujuan tahap akhir ini.' };
        }

        if (action === 'Setujui') {
          cutiSheet.getRange(rowIndex, statusIdx + 1).setValue('Disetujui');
          if (app3NameIdx !== -1) cutiSheet.getRange(rowIndex, app3NameIdx + 1).setValue(user.fullname);
          if (app3JabIdx !== -1) cutiSheet.getRange(rowIndex, app3JabIdx + 1).setValue(user.jabatan);
          if (app3TglIdx !== -1) cutiSheet.getRange(rowIndex, app3TglIdx + 1).setValue(nowStr);
          if (app3StatusIdx !== -1) cutiSheet.getRange(rowIndex, app3StatusIdx + 1).setValue('Disetujui');
          if (app3CatatanIdx !== -1) cutiSheet.getRange(rowIndex, app3CatatanIdx + 1).setValue(catatan || 'Disetujui Final PJO / Direktur');
          return {
            success: true,
            message: 'Pengajuan cuti telah disetujui dan disahkan secara resmi oleh PJO / Direktur.',
            newStatus: 'Disetujui'
          };
        } else {
          cutiSheet.getRange(rowIndex, statusIdx + 1).setValue('Ditolak PJO');
          if (app3NameIdx !== -1) cutiSheet.getRange(rowIndex, app3NameIdx + 1).setValue(user.fullname);
          if (app3JabIdx !== -1) cutiSheet.getRange(rowIndex, app3JabIdx + 1).setValue(user.jabatan);
          if (app3TglIdx !== -1) cutiSheet.getRange(rowIndex, app3TglIdx + 1).setValue(nowStr);
          if (app3StatusIdx !== -1) cutiSheet.getRange(rowIndex, app3StatusIdx + 1).setValue('Ditolak');
          if (app3CatatanIdx !== -1) cutiSheet.getRange(rowIndex, app3CatatanIdx + 1).setValue(catatan);
          if (tolakAlasanIdx !== -1) cutiSheet.getRange(rowIndex, tolakAlasanIdx + 1).setValue(catatan);
          return {
            success: true,
            message: 'Pengajuan cuti telah ditolak oleh PJO / Direktur.',
            newStatus: 'Ditolak PJO'
          };
        }
      }

      else {
        return { success: false, message: 'Pengajuan cuti ini sudah berstatus "' + currentStatus + '" dan tidak dapat diproses lagi.' };
      }
    }
  }

  return { success: false, message: 'Data pengajuan cuti tidak ditemukan.' };
}

function getLampiranCutiBase64(token, cutiId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  cutiId = String(cutiId || '').trim();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cutiSheet = getOrCreatePengajuanCutiSheet(ss);
  const data = cutiSheet.getDataRange().getValues();
  if (!data || data.length <= 1) return { success: false, message: 'Data tidak ditemukan.' };

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id pengajuan') || h === 'id cuti' || h === 'id' || h.includes('kode'); });
  const empIdIdx = headers.findIndex(function (h) { return h.includes('id karyawan') || h.includes('nik') || h.includes('nip') || h === 'username' || h === 'id_karyawan'; });
  const driveIdIdx = headers.findIndex(function (h) { return h.includes('file id') || h.includes('fileid'); });

  const currentUsername = String(session.user.username || '').trim().toLowerCase();
  const userRole = String(session.user.role || '').trim().toLowerCase();
  const isHRD = (userRole === 'admin' || userRole === 'hrd');

  for (let i = 1; i < data.length; i++) {
    const rowId = idIdx !== -1 ? String(data[i][idIdx] || '').trim() : String(data[i][0] || '').trim();
    if (rowId === cutiId) {
      const rowEmpId = empIdIdx !== -1 ? String(data[i][empIdIdx] || '').trim().toLowerCase() : String(data[i][2] || '').trim().toLowerCase();
      if (!isHRD && rowEmpId !== currentUsername) {
        return { success: false, message: 'Akses ditolak.' };
      }
      const fileId = driveIdIdx !== -1 ? String(data[i][driveIdIdx] || '').trim() : String(data[i][15] || '').trim();
      if (!fileId) return { success: false, message: 'Tidak ada lampiran pada pengajuan ini.' };

      try {
        const file = DriveApp.getFileById(fileId);
        const mimeType = file.getMimeType();
        const base64Data = Utilities.base64Encode(file.getBlob().getBytes());
        return {
          success: true,
          fileName: file.getName(),
          mimeType: mimeType,
          base64: 'data:' + mimeType + ';base64,' + base64Data
        };
      } catch (err) {
        return { success: false, message: 'Gagal memuat lampiran: ' + err.toString() };
      }
    }
  }

  return { success: false, message: 'Pengajuan cuti tidak ditemukan.' };
}

// --- MODUL MASTER ROSTER (POLA KERJA & CUTI) ---

function getOrCreateMasterRosterSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('MasterRoster') || ss.getSheetByName('Master Roster') || ss.getSheetByName('Roster') || ss.getSheetByName('Master_Roster');
  if (!sheet) {
    sheet = ss.insertSheet('MasterRoster');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['ID Roster', 'Nama Roster', 'Hari Kerja', 'Hari Cuti', 'Keterangan', 'Tanggal Dibuat']);
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#e2e8f0');
  }
  if (sheet.getLastRow() <= 1) {
    // Pola roster default operasional tambang/site
    sheet.appendRow(['ROSTER-001', '12 : 2', 90, 14, '12 Minggu Kerja, 2 Minggu Cuti On-Site', new Date()]);
    sheet.appendRow(['ROSTER-002', '8 : 2', 56, 14, '8 Minggu Kerja, 2 Minggu Cuti On-Site', new Date()]);
    sheet.appendRow(['ROSTER-003', '6 : 2', 42, 14, '6 Minggu Kerja, 2 Minggu Cuti On-Site', new Date()]);
    sheet.appendRow(['ROSTER-004', '4 : 2', 28, 14, '4 Minggu Kerja, 2 Minggu Cuti On-Site', new Date()]);
  }
  return sheet;
}

function parseRosterDataFromSheet(rosterSheet) {
  if (!rosterSheet) return [];
  const data = rosterSheet.getDataRange().getValues();
  const list = [];
  if (!data || data.length <= 1) return list;

  const headers = data[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const idIdx = headers.findIndex(function (h) { return h.includes('id'); });
  const nameIdx = headers.findIndex(function (h) { return h.includes('nama') || h.includes('pola') || (h.includes('roster') && !h.includes('id')); });
  const workIdx = headers.findIndex(function (h) { return h.includes('kerja'); });
  const leaveIdx = headers.findIndex(function (h) { return h.includes('cuti'); });
  const noteIdx = headers.findIndex(function (h) { return h.includes('ket') || h.includes('note') || h.includes('catatan'); });

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row || row.every(function (cell) { return String(cell || '').trim() === ''; })) continue;

    const id = idIdx !== -1 && row[idIdx] ? String(row[idIdx]).trim() : ('ROSTER-' + i);
    const name = nameIdx !== -1 && row[nameIdx] ? String(row[nameIdx]).trim() : String(row[1] || '').trim();
    const workDays = workIdx !== -1 ? Number(row[workIdx] || 0) : Number(row[2] || 0);
    const leaveDays = leaveIdx !== -1 ? Number(row[leaveIdx] || 0) : Number(row[3] || 0);
    const note = noteIdx !== -1 ? String(row[noteIdx] || '').trim() : String(row[4] || '').trim();

    if (name) {
      list.push({
        id: id,
        name: name,
        workDays: workDays,
        leaveDays: leaveDays,
        note: note,
        notes: note
      });
    }
  }
  return list;
}

function getRosterList(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = getOrCreateMasterRosterSheet(ss);
  const list = parseRosterDataFromSheet(rosterSheet);
  return { success: true, roster: list, list: list };
}

function saveRoster(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak. Hanya Admin yang dapat mengelola Master Roster.' };

  payload = payload || {};
  const name = String(payload.name || '').trim();
  const workDays = parseInt(payload.workDays, 10) || 0;
  const leaveDays = parseInt(payload.leaveDays, 10) || 0;
  const note = String(payload.note || payload.notes || '').trim();
  const rosterId = String(payload.id || '').trim();
  const isNew = !rosterId;

  if (!name) return { success: false, message: 'Nama Pola Roster wajib diisi (misal: 12 : 2).' };
  if (workDays <= 0) return { success: false, message: 'Jumlah hari kerja harus lebih dari 0.' };
  if (leaveDays <= 0) return { success: false, message: 'Jumlah hari cuti harus lebih dari 0.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = getOrCreateMasterRosterSheet(ss);
  const data = rosterSheet.getDataRange().getValues();

  if (isNew) {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toLowerCase() === name.toLowerCase()) {
        return { success: false, message: 'Nama pola roster "' + name + '" sudah terdaftar.' };
      }
    }
    const newId = 'ROSTER-' + Math.floor(100 + Math.random() * 900);
    rosterSheet.appendRow([newId, name, workDays, leaveDays, note, new Date()]);
    return { success: true, message: 'Pola Roster baru berhasil ditambahkan.' };
  } else {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === rosterId) {
        rosterSheet.getRange(i + 1, 2).setValue(name);
        rosterSheet.getRange(i + 1, 3).setValue(workDays);
        rosterSheet.getRange(i + 1, 4).setValue(leaveDays);
        rosterSheet.getRange(i + 1, 5).setValue(note);
        return { success: true, message: 'Data Pola Roster berhasil diperbarui.' };
      }
    }
    return { success: false, message: 'Data Pola Roster tidak ditemukan.' };
  }
}

function deleteRoster(token, rosterId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak. Hanya Admin yang dapat menghapus data Master Roster.' };

  rosterId = String(rosterId || '').trim();
  if (!rosterId) return { success: false, message: 'ID Roster tidak valid.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rosterSheet = getOrCreateMasterRosterSheet(ss);
  const data = rosterSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === rosterId) {
      rosterSheet.deleteRow(i + 1);
      return { success: true, message: 'Pola Roster berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Data Pola Roster tidak ditemukan.' };
}

// --- FITUR PENGUMUMAN (ANNOUNCEMENT) ---
// Catatan: Modul Pengumuman telah dipisahkan ke file 'Kode_Pengumuman.js'

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




