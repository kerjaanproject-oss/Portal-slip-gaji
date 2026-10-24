// ==========================================================================
// MODUL MASTER ROSTER (POLA KERJA & CUTI)
// ==========================================================================

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
