// ==========================================================================
// MODUL PENGAJUAN & PERSETUJUAN CUTI KARYAWAN
// ==========================================================================

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