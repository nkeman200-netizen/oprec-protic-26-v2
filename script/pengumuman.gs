// ====== KONFIG ======
const SHEET_ID   = 'MASUKKAN_ID_SPREADSHEET_DISINI'; // <-- ganti
const DATA_SHEET = 'Hasil Akhir';
const LOG_SHEET  = 'log';

// ====== WEB API ======
function doGet(e) {
  const npm = (e.parameter.npm || '').trim();
  if (!npm) return jsonOutput({ ok: false, error: 'Parameter "npm" wajib diisi.' });

  try {
    const data = getDataByNPM(npm); // bisa null
    return jsonOutput({ ok: true, data });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function jsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ====== LOGIC ======
function getDataByNPM(NPM) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(DATA_SHEET);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(NPM).trim()) {
      logSearch(data[i][0], data[i][1], data[i][2], data[i][3], true);
      return {
        NPM:    data[i][0],
        Nama:   data[i][1],
        Divisi: data[i][2],
        Status: data[i][3]
      };
    }
  }
  // log juga kalau tidak ditemukan
  logSearch(NPM, '-', '-', '-', false);
  return null;
}

function logSearch(NPM, nama, status, divisi, ditemukan) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let logSheet = ss.getSheetByName(LOG_SHEET);
  if (!logSheet) {
    logSheet = ss.insertSheet(LOG_SHEET);
    logSheet.appendRow(['Timestamp', 'NPM', 'Nama', 'Status', 'Divisi', 'Ditemukan']);
  }
  logSheet.appendRow([new Date(), NPM, nama, status, divisi, ditemukan ? 'Ya' : 'Tidak']);
}

