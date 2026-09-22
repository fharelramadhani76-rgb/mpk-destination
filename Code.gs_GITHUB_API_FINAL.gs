/**
 * MPK DESTINATION — GOOGLE APPS SCRIPT BACKEND
 * SMAN 1 Kedamean
 *
 * FILE:
 *   Code.gs
 *   Index.html
 *
 * SETUP:
 * 1. Buat Google Spreadsheet baru.
 * 2. Extensions > Apps Script.
 * 3. Tempel Code.gs dan Index.html.
 * 4. Jalankan setupMPK() sekali dari editor Apps Script.
 * 5. Isi DRIVE_FOLDER_ID dengan ID folder Google Drive tujuan.
 *    Jika kosong, foto tetap dicatat tetapi tidak di-upload ke Drive.
 * 6. Deploy > New deployment > Web app.
 *    Execute as: Me
 *    Who has access: sesuai kebutuhan sekolah.
 */

const CONFIG = {
  SPREADSHEET_ID: '', // Kosong = Spreadsheet yang terikat ke project.
  DRIVE_FOLDER_ID: '', // Isi ID folder Drive untuk upload foto.
  APP_NAME: 'MPK Destination',
  SCHOOL_NAME: 'SMAN 1 Kedamean'
};

const SHEETS = {
  DATA: 'Data',
  PROGRAMS: 'Programs',
  MEMBERS: 'Members',
  PHOTOS: 'Photos',
  SESSIONS: 'Sessions'
};

function doGet(e) {
  const GITHUB_URL = 'https://fharelramadhani76-rgb.github.io/mpk-destination/';

  // Health check: /exec?api=health
  if (e && e.parameter && e.parameter.api === 'health') {
    return ContentService
      .createTextOutput(JSON.stringify({
        ok: true,
        service: CONFIG.APP_NAME + ' API',
        status: 'online',
        time: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // /exec is the entry link. The actual camera UI runs on GitHub Pages.
  const safeUrl = JSON.stringify(GITHUB_URL);
  return HtmlService.createHtmlOutput(`
<!doctype html>
<html>
<head>
  <base target="_top">
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta http-equiv="refresh" content="0;url=${GITHUB_URL}">
  <title>${CONFIG.APP_NAME}</title>
</head>
<body style="font-family:Arial,sans-serif;text-align:center;padding:50px">
  <h2>📸 Membuka ${CONFIG.APP_NAME}...</h2>
  <p>Jika tidak otomatis berpindah, <a href="${GITHUB_URL}">klik di sini</a>.</p>
  <script>
    window.location.replace(${safeUrl});
  </script>
</body>
</html>`);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSS_() {
  if (CONFIG.SPREADSHEET_ID) {
    return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Spreadsheet tidak ditemukan.');
  return ss;
}

function getOrCreateSheet_(name, headers) {
  const ss = getSS_();
  let sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#07111f')
      .setFontColor('#ffffff');
  } else {
    const current = sh.getRange(1, 1, 1, Math.max(headers.length, sh.getLastColumn())).getValues()[0];
    if (current.every(v => v === '')) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
    sh.setFrozenRows(1);
  }

  return sh;
}

function setupMPK() {
  const ss = getSS_();

  getOrCreateSheet_(
    SHEETS.DATA,
    ['ID', 'NAMA', 'UNIVERSITAS', 'ANGKATAN', 'JALUR MASUK', 'PROGRAM STUDI', 'FAKULTAS', 'STATUS', 'TIMESTAMP']
  );

  getOrCreateSheet_(
    SHEETS.PROGRAMS,
    ['ID', 'NAMA PROGRAM', 'GENERASI', 'DIVISI', 'PJ', 'TANGGAL', 'LOKASI', 'DESKRIPSI', 'STATUS', 'FOLDER DRIVE', 'TIMESTAMP']
  );

  getOrCreateSheet_(
    SHEETS.MEMBERS,
    ['ID', 'NAMA', 'NISN', 'KELAS', 'JABATAN', 'DIVISI', 'GENERASI', 'PERIODE', 'FOTO', 'STATUS', 'TIMESTAMP']
  );

  getOrCreateSheet_(
    SHEETS.PHOTOS,
    ['ID', 'SESSION', 'FILE NAME', 'DRIVE URL', 'DRIVE ID', 'LAYOUT', 'FILTER', 'FRAME', 'CAPTION', 'TIMESTAMP']
  );

  getOrCreateSheet_(
    SHEETS.SESSIONS,
    ['SESSION', 'STATUS', 'CREATED', 'LAST ACTIVE']
  );

  return {
    ok: true,
    message: 'Sheet MPK Destination siap digunakan.',
    spreadsheetUrl: ss.getUrl()
  };
}

function sheetToObjects_(sheetName) {
  const sh = getSS_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 1) return [];

  const values = sh.getDataRange().getDisplayValues();
  if (values.length < 2) return [];

  const headers = values[0];
  return values.slice(1)
    .filter(row => row.some(v => String(v).trim() !== ''))
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i] ?? '');
      return obj;
    });
}

function getAppData() {
  return {
    ok: true,
    data: sheetToObjects_(SHEETS.DATA),
    programs: sheetToObjects_(SHEETS.PROGRAMS),
    members: sheetToObjects_(SHEETS.MEMBERS),
    photos: sheetToObjects_(SHEETS.PHOTOS),
    school: CONFIG.SCHOOL_NAME,
    appName: CONFIG.APP_NAME
  };
}

function makeId_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function appendObject_(sheetName, headers, obj) {
  const sh = getOrCreateSheet_(sheetName, headers);
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  sh.appendRow(row);
  return obj;
}

function saveHallOfFame(item) {
  if (!item || !item.nama || !item.universitas) {
    throw new Error('Nama dan universitas wajib diisi.');
  }

  return appendObject_(
    SHEETS.DATA,
    ['ID', 'NAMA', 'UNIVERSITAS', 'ANGKATAN', 'JALUR MASUK', 'PROGRAM STUDI', 'FAKULTAS', 'STATUS', 'TIMESTAMP'],
    {
      ID: makeId_('PTN'),
      NAMA: String(item.nama).trim(),
      UNIVERSITAS: String(item.universitas).trim(),
      ANGKATAN: String(item.angkatan || '').trim(),
      'JALUR MASUK': String(item.jalur || '').trim(),
      'PROGRAM STUDI': String(item.prodi || '').trim(),
      FAKULTAS: String(item.fakultas || '').trim(),
      STATUS: String(item.status || 'Terverifikasi').trim(),
      TIMESTAMP: new Date()
    }
  );
}

function saveProgram(item) {
  if (!item || !item.namaProgram) {
    throw new Error('Nama program wajib diisi.');
  }

  return appendObject_(
    SHEETS.PROGRAMS,
    ['ID', 'NAMA PROGRAM', 'GENERASI', 'DIVISI', 'PJ', 'TANGGAL', 'LOKASI', 'DESKRIPSI', 'STATUS', 'FOLDER DRIVE', 'TIMESTAMP'],
    {
      ID: makeId_('PRG'),
      'NAMA PROGRAM': String(item.namaProgram).trim(),
      GENERASI: String(item.generasi || '').trim(),
      DIVISI: String(item.divisi || '').trim(),
      PJ: String(item.pj || '').trim(),
      TANGGAL: String(item.tanggal || '').trim(),
      LOKASI: String(item.lokasi || '').trim(),
      DESKRIPSI: String(item.deskripsi || '').trim(),
      STATUS: String(item.status || 'Terlaksana').trim(),
      'FOLDER DRIVE': String(item.folder || '').trim(),
      TIMESTAMP: new Date()
    }
  );
}

function saveMember(item) {
  if (!item || !item.nama) {
    throw new Error('Nama pengurus wajib diisi.');
  }

  return appendObject_(
    SHEETS.MEMBERS,
    ['ID', 'NAMA', 'NISN', 'KELAS', 'JABATAN', 'DIVISI', 'GENERASI', 'PERIODE', 'FOTO', 'STATUS', 'TIMESTAMP'],
    {
      ID: makeId_('MBR'),
      NAMA: String(item.nama).trim(),
      NISN: String(item.nisn || '').trim(),
      KELAS: String(item.kelas || '').trim(),
      JABATAN: String(item.jabatan || '').trim(),
      DIVISI: String(item.divisi || '').trim(),
      GENERASI: String(item.generasi || '').trim(),
      PERIODE: String(item.periode || '').trim(),
      FOTO: String(item.foto || '').trim(),
      STATUS: String(item.status || 'Aktif').trim(),
      TIMESTAMP: new Date()
    }
  );
}

/**
 * Menerima data URL gambar dari browser.
 * Proses gambar sudah dilakukan di browser sebelum dikirim.
 */
function randomSessionCode_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let code = '';
  for (let i = 0; i < 6; i++) code += chars[Math.floor(Math.random() * chars.length)];
  return code;
}

function createPhotoSession() {
  const sh = getOrCreateSheet_(SHEETS.SESSIONS, ['SESSION', 'STATUS', 'CREATED', 'LAST ACTIVE']);
  let code = randomSessionCode_();
  const existing = sheetToObjects_(SHEETS.SESSIONS).map(x => x.SESSION);
  while (existing.includes(code)) code = randomSessionCode_();

  const now = new Date();
  sh.appendRow([code, 'ACTIVE', now, now]);
  return { ok: true, session: code, created: now.toISOString() };
}

function joinPhotoSession(code) {
  code = String(code || '').trim().toUpperCase();
  if (!/^[A-Z0-9]{6}$/.test(code)) throw new Error('Kode sesi harus 6 karakter.');

  const sessions = sheetToObjects_(SHEETS.SESSIONS);
  const found = sessions.find(x => x.SESSION === code && x.STATUS === 'ACTIVE');
  if (!found) throw new Error('Sesi tidak ditemukan atau sudah ditutup.');

  touchPhotoSession(code);
  return { ok: true, session: code };
}

function touchPhotoSession(code) {
  const sh = getSS_().getSheetByName(SHEETS.SESSIONS);
  if (!sh || sh.getLastRow() < 2) return false;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === code) {
      sh.getRange(i + 2, 4).setValue(new Date());
      return true;
    }
  }
  return false;
}

function closePhotoSession(code) {
  const sh = getSS_().getSheetByName(SHEETS.SESSIONS);
  if (!sh || sh.getLastRow() < 2) return false;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] === code) {
      sh.getRange(i + 2, 2).setValue('CLOSED');
      sh.getRange(i + 2, 4).setValue(new Date());
      return true;
    }
  }
  return false;
}

function getSessionPhotos(code) {
  code = String(code || '').trim().toUpperCase();
  if (!code) return [];
  touchPhotoSession(code);

  const rows = sheetToObjects_(SHEETS.PHOTOS);
  return rows
    .filter(x => x.SESSION === code)
    .map(x => ({
      id: x.ID,
      session: x.SESSION,
      fileName: x['FILE NAME'],
      url: x['DRIVE URL'],
      driveId: x['DRIVE ID'],
      layout: x.LAYOUT,
      filter: x.FILTER,
      frame: x.FRAME,
      caption: x.CAPTION,
      timestamp: x.TIMESTAMP
    }));
}

/**
 * Menerima data URL gambar dari browser.
 * Data visual diproses di browser terlebih dahulu.
 * SESSION menghubungkan beberapa device ke satu ruang foto.
 */
function uploadPhoto(dataUrl, meta) {
  if (!dataUrl || dataUrl.indexOf('base64,') === -1) {
    throw new Error('Data gambar tidak valid.');
  }

  const session = String((meta && meta.session) || '').trim().toUpperCase();
  if (!session) throw new Error('Sesi Photo Booth belum dipilih.');

  const sessions = sheetToObjects_(SHEETS.SESSIONS);
  const valid = sessions.some(x => x.SESSION === session && x.STATUS === 'ACTIVE');
  if (!valid) throw new Error('Sesi tidak aktif. Buat atau gabung sesi baru.');

  const base64 = dataUrl.split('base64,')[1];
  const mime = (meta && meta.mime) || 'image/png';
  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(
    bytes,
    mime,
    (meta && meta.fileName) || ('MPK-Snap-' + Date.now() + '.png')
  );

  let fileId = '';
  let fileUrl = '';

  if (CONFIG.DRIVE_FOLDER_ID) {
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const file = folder.createFile(blob);
    fileId = file.getId();
    fileUrl = file.getUrl();

    // Supaya device lain dalam sesi dapat menampilkan foto.
    // Jika kebijakan Workspace melarang public link, URL tetap tercatat
    // dan dapat dibuka oleh akun yang memiliki akses Drive.
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {
      console.warn('Sharing link tidak dapat diubah: ' + sharingErr);
    }
  }

  const item = {
    ID: makeId_('PIC'),
    SESSION: session,
    'FILE NAME': blob.getName(),
    'DRIVE URL': fileUrl,
    'DRIVE ID': fileId,
    LAYOUT: String((meta && meta.layout) || ''),
    FILTER: String((meta && meta.filter) || ''),
    FRAME: String((meta && meta.frame) || ''),
    CAPTION: String((meta && meta.caption) || ''),
    TIMESTAMP: new Date()
  };

  appendObject_(
    SHEETS.PHOTOS,
    ['ID', 'SESSION', 'FILE NAME', 'DRIVE URL', 'DRIVE ID', 'LAYOUT', 'FILTER', 'FRAME', 'CAPTION', 'TIMESTAMP'],
    item
  );

  touchPhotoSession(session);

  return {
    ok: true,
    id: item.ID,
    session: session,
    fileName: item['FILE NAME'],
    url: fileUrl,
    driveId: fileId,
    message: CONFIG.DRIVE_FOLDER_ID
      ? 'Foto berhasil disimpan dan masuk ke sesi.'
      : 'Foto tercatat dalam sesi. DRIVE_FOLDER_ID belum dikonfigurasi.'
  };
}

function deleteById(sheetName, id) {
  const sh = getSS_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return false;

  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getDisplayValues();
  for (let i = values.length - 1; i >= 0; i--) {
    if (values[i][0] === id) {
      sh.deleteRow(i + 2);
      return true;
    }
  }
  return false;
}

function getStats() {
  const data = sheetToObjects_(SHEETS.DATA);
  const programs = sheetToObjects_(SHEETS.PROGRAMS);
  const members = sheetToObjects_(SHEETS.MEMBERS);

  const universities = [...new Set(data.map(x => x.UNIVERSITAS).filter(Boolean))];

  const counts = {};
  data.forEach(x => {
    if (x.UNIVERSITAS) counts[x.UNIVERSITAS] = (counts[x.UNIVERSITAS] || 0) + 1;
  });

  const topUniversity = Object.keys(counts).sort((a,b) => counts[b] - counts[a])[0] || '-';

  return {
    totalAlumni: data.length,
    totalUniversitas: universities.length,
    totalPrograms: programs.length,
    totalMembers: members.length,
    topUniversity
  };
}


/* =========================================================
   API UNTUK GITHUB PAGES
   Frontend GitHub memanggil Apps Script melalui POST /exec.
   Jangan hapus fungsi ini.
   ========================================================= */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return apiJson_({ok:false, error:'Request kosong.'});
    }

    let req;
    try {
      req = JSON.parse(e.postData.contents);
    } catch (err) {
      return apiJson_({ok:false, error:'Body request bukan JSON yang valid.'});
    }

    const action = String(req.action || '').trim();
    if (!action) {
      return apiJson_({ok:false, error:'Action API tidak ditemukan.'});
    }

    let result;

    switch (action) {
      case 'getAppData':
        result = getAppData();
        break;

      case 'getStats':
        result = getStats();
        break;

      case 'saveProgram':
        result = saveProgram(req.item || {});
        break;

      case 'saveMember':
        result = saveMember(req.item || {});
        break;

      case 'saveHallOfFame':
        result = saveHallOfFame(req.item || {});
        break;

      case 'createPhotoSession':
        result = createPhotoSession();
        break;

      case 'joinPhotoSession':
        result = joinPhotoSession(req.code || '');
        break;

      case 'getSessionPhotos':
        result = getSessionPhotos(req.code || '');
        break;

      case 'closePhotoSession':
        result = closePhotoSession(
          String(req.code || '').trim().toUpperCase()
        );
        break;

      case 'uploadPhoto':
        result = uploadPhoto(
          req.dataUrl || '',
          req.meta || {}
        );
        break;

      default:
        return apiJson_({
          ok:false,
          error:'Action tidak dikenal: ' + action
        });
    }

    return apiJson_(result);

  } catch (err) {
    console.error(err);
    return apiJson_({
      ok:false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function apiJson_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
