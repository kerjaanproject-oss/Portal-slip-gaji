// ==========================================================================
// MODUL AUTENTIKASI, SESI, KEAMANAN & PIN SLIP GAJI
// File: Kode_Auth.js
// Berisi fungsi setup database, hashing password & PIN, manajemen token sesi
// HMAC-SHA256, pembatasan rate limit login & PIN, OTP reset password & PIN.
// ==========================================================================

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
