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

// ==========================================================================
// MODUL AUTENTIKASI, SESI & PIN SLIP GAJI
// Catatan: Seluruh fungsi setup database, autentikasi, sesi, dan PIN slip gaji telah dipisahkan ke file 'Kode_Auth.js'
// ==========================================================================

// ==========================================================================
// MODUL MASTER DATA KARYAWAN & DEPARTEMEN
// Catatan: Seluruh fungsi data awal, karyawan & departemen telah dipisahkan ke file 'Kode_Karyawan.js'
// ==========================================================================

// ==========================================================================
// MODUL PENGOLAHAN SLIP GAJI, FILE STORAGE & NOTIFIKASI EMAIL
// Catatan: Seluruh fungsi slip gaji telah dipisahkan ke file 'Kode_SlipGaji.js'
// ==========================================================================

// --- MODUL PENGAJUAN & PERSETUJUAN CUTI KARYAWAN ---
// Catatan: Modul Pengajuan & Persetujuan Cuti telah dipisahkan ke file 'Kode_Cuti.js'

// --- MODUL MASTER ROSTER (POLA KERJA & CUTI) ---
// Catatan: Modul Master Roster telah dipisahkan ke file 'Kode_Roster.js'

// --- FITUR PENGUMUMAN (ANNOUNCEMENT) ---
// Catatan: Modul Pengumuman telah dipisahkan ke file 'Kode_Pengumuman.js'

// ==========================================================================
// FITUR PENDAFTARAN MANDIRI & PERSETUJUAN AKUN (SELF-REGISTRATION)
// Catatan: Modul Pendaftaran Akun telah dipisahkan ke file 'Kode_Pendaftaran.js'
// ==========================================================================
