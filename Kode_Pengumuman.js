// ==========================================================================
// MODUL PENGUMUMAN (ANNOUNCEMENTS)
// ==========================================================================

function getOrCreatePengumumanSheet(ss) {
  if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
  let annSheet = ss.getSheetByName('Pengumuman');
  if (!annSheet) {
    annSheet = ss.insertSheet('Pengumuman');
    annSheet.appendRow([
      'ID Pengumuman',
      'Judul',
      'Isi',
      'Tipe',
      'Tanggal Mulai',
      'Tanggal Selesai',
      'Status',
      'Tanggal Dibuat',
      'Dibuat Oleh'
    ]);
    annSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#e2e8f0');
  }
  return annSheet;
}

function formatDateString(d) {
  if (!d) return '';
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  const str = String(d).trim();
  if (!str) return '';
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
  }
  return str;
}

/**
 * Public Action: Mendapatkan pengumuman aktif untuk ditampilkan di modal pop-up login
 */
function getActiveAnnouncement() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreatePengumumanSheet(ss);
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { success: true, hasAnnouncement: false };
    }

    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyy-MM-dd');
    const activeList = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id = String(row[0] || '').trim();
      const judul = String(row[1] || '').trim();
      const isi = String(row[2] || '').trim();
      const tipe = String(row[3] || 'Informasi').trim();
      const startStr = formatDateString(row[4]);
      const endStr = formatDateString(row[5]);
      const status = String(row[6] || 'Aktif').trim();
      const createdStr = formatDateString(row[7]);
      const createdBy = String(row[8] || '').trim();

      if (!id || !judul || status.toLowerCase() !== 'aktif') {
        continue;
      }

      // Filter periode jika tanggal ditentukan
      if (startStr && startStr > todayStr) {
        continue; // Belum mulai
      }
      if (endStr && endStr < todayStr) {
        continue; // Sudah berakhir
      }

      // Prioritas: Mendesak (3) > Penting (2) > Informasi (1)
      let priority = 1;
      const lowerTipe = tipe.toLowerCase();
      if (lowerTipe === 'mendesak' || lowerTipe === 'urgent') priority = 3;
      else if (lowerTipe === 'penting' || lowerTipe === 'warning') priority = 2;

      activeList.push({
        id: id,
        judul: judul,
        isi: isi,
        tipe: tipe,
        tanggalMulai: startStr,
        tanggalSelesai: endStr,
        status: status,
        tanggalDibuat: createdStr,
        dibuatOleh: createdBy,
        priority: priority,
        rowIndex: i + 1
      });
    }

    if (activeList.length === 0) {
      return { success: true, hasAnnouncement: false };
    }

    // Urutkan prioritas tertinggi lalu yang terbaru
    activeList.sort(function(a, b) {
      if (b.priority !== a.priority) {
        return b.priority - a.priority;
      }
      return b.rowIndex - a.rowIndex;
    });

    const chosen = activeList[0];
    return {
      success: true,
      hasAnnouncement: true,
      announcement: chosen
    };
  } catch (err) {
    console.error('getActiveAnnouncement error:', err);
    return { success: false, hasAnnouncement: false, message: err.toString() };
  }
}

/**
 * Protected Action (HRD/Admin): Mengambil semua data pengumuman
 */
function getAnnouncements(token) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') {
    return { success: false, message: 'Akses ditolak. Fitur kelola pengumuman hanya untuk Admin/HRD.' };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreatePengumumanSheet(ss);
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const id = String(row[0] || '').trim();
    if (!id) continue;
    const status = String(row[6] || 'Aktif').trim();

    list.push({
      id: id,
      judul: String(row[1] || '').trim(),
      isi: String(row[2] || '').trim(),
      tipe: String(row[3] || 'Informasi').trim(),
      tanggalMulai: formatDateString(row[4]),
      tanggalSelesai: formatDateString(row[5]),
      status: status,
      tanggalDibuat: formatDateString(row[7]),
      dibuatOleh: String(row[8] || '').trim(),
      rowIndex: i + 1
    });
  }

  list.reverse();
  return { success: true, list: list };
}

/**
 * Protected Action (HRD/Admin): Menyimpan atau memperbarui pengumuman
 */
function saveAnnouncement(token, payload) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak. Hanya Admin/HRD yang dapat mengelola pengumuman.' };

  payload = payload || {};
  const annId = String(payload.id || '').trim();
  const judul = String(payload.judul || '').trim();
  const isi = String(payload.isi || '').trim();
  const tipe = String(payload.tipe || 'Informasi').trim();
  const tanggalMulai = String(payload.tanggalMulai || '').trim();
  const tanggalSelesai = String(payload.tanggalSelesai || '').trim();
  const status = String(payload.status || 'Aktif').trim();

  if (!judul) return { success: false, message: 'Judul pengumuman wajib diisi.' };
  if (!isi) return { success: false, message: 'Isi pengumuman wajib diisi.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreatePengumumanSheet(ss);
  const data = sheet.getDataRange().getValues();
  const isNew = !annId;

  if (isNew) {
    const dateCode = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Jakarta', 'yyyyMMdd');
    const newId = 'ANN-' + dateCode + '-' + Math.floor(100 + Math.random() * 900);
    const createdBy = session.user.fullname || session.user.username || 'HRD';
    sheet.appendRow([
      newId,
      judul,
      isi,
      tipe,
      tanggalMulai,
      tanggalSelesai,
      status,
      new Date(),
      createdBy
    ]);
    return { success: true, message: 'Pengumuman baru berhasil disimpan.', id: newId, annId: newId };
  } else {
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === annId) {
        sheet.getRange(i + 1, 2).setValue(judul);
        sheet.getRange(i + 1, 3).setValue(isi);
        sheet.getRange(i + 1, 4).setValue(tipe);
        sheet.getRange(i + 1, 5).setValue(tanggalMulai);
        sheet.getRange(i + 1, 6).setValue(tanggalSelesai);
        sheet.getRange(i + 1, 7).setValue(status);
        return { success: true, message: 'Pengumuman berhasil diperbarui.', id: annId, annId: annId };
      }
    }
    return { success: false, message: 'Pengumuman dengan ID ' + annId + ' tidak ditemukan.' };
  }
}

/**
 * Protected Action (HRD/Admin): Mengubah status Aktif/Nonaktif pengumuman
 */
function toggleAnnouncementStatus(token, annId, newStatus) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  annId = String(annId || '').trim();
  newStatus = String(newStatus || '').trim();
  if (!annId || !newStatus) return { success: false, message: 'Parameter tidak valid.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreatePengumumanSheet(ss);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === annId) {
      sheet.getRange(i + 1, 7).setValue(newStatus);
      return { success: true, message: 'Status pengumuman diubah menjadi ' + newStatus + '.' };
    }
  }
  return { success: false, message: 'Pengumuman tidak ditemukan.' };
}

/**
 * Protected Action (HRD/Admin): Menghapus pengumuman
 */
function deleteAnnouncement(token, annId) {
  const session = validateSession(token);
  if (!session.valid) return { success: false, message: session.message, sessionExpired: true };
  if (session.user.role !== 'Admin') return { success: false, message: 'Akses ditolak.' };

  annId = String(annId || '').trim();
  if (!annId) return { success: false, message: 'ID pengumuman tidak valid.' };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getOrCreatePengumumanSheet(ss);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === annId) {
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Pengumuman berhasil dihapus.' };
    }
  }
  return { success: false, message: 'Pengumuman tidak ditemukan.' };
}
