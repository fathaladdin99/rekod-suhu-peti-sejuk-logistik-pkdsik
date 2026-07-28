/**
 * REKOD SUHU PETI SEJUK (FARMASI LOGISTIK) - WEB APP & WEBHOOK API BACKEND
 * Unlimited Years Support (2025 to 2100+)
 */

/* === SPREADSHEET IDs (Gantikan dengan ID Google Spreadsheet anda) === */
const SPREADSHEET_IDS = {
  'PETI_SEJUK_1': '1R55c5d_jQ4VSG2S5ATnIMYnCMwpxONs8gC2-ewncaPA',
  'PETI_SEJUK_2': '1t5ZB1rAceHNrK8C49gsAz3P--U2A1YrMsya8cqgY40c',
  'PETI_SEJUK_3': '1EhGj8QiOFxBWCySobklCwUUeqW6pfbeuJS1lCGQ62ek',
  'PETI_SEJUK_4': '1PmS1iB_5v4gl7rNb910UqH9XDfnO4WhI8su5WD-medE'
};

/* === Sheet Names === */
const SHEET_TAB_REKOD   = 'REKOD SUHU';
const SHEET_TAB_CATATAN = 'CATATAN SUHU LUAR JULAT';

/* === Bounds & Constraints === */
const BASE_YEAR = 2025;
const LAST_YEAR = 2100; // Support forever (up to 2100+)
const START_ROW = 2; // Row 1 is header (Tarikh, Waktu, Suhu Min, Suhu Semasa, Suhu Max, Perkara, Tandatangan)

/**
 * 1. Web App GET Handler
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'config') {
    return createJsonResponse(getConfig());
  }
  return HtmlService.createHtmlOutputFromFile('form')
    .setTitle('Rekod Suhu Peti Sejuk Farmasi')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * 2. Web App POST Handler (Web API Endpoint for Standalone Web App)
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch(err) {
        payload = e.parameter || {};
      }
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    if (typeof payload.incident === 'string') {
      try { payload.incident = JSON.parse(payload.incident); } catch(err) {}
    }

    if (payload.action === 'check') {
      const targetRow = rowFor_(payload.date, payload.slot);
      const ss = SpreadsheetApp.openById(idForLokasi_(payload.lokasi));
      const rekod = ss.getSheetByName(SHEET_TAB_REKOD);
      let exists = false;
      if (rekod) {
        const existing = rekod.getRange(targetRow, 3, 1, 5).getValues()[0];
        exists = existing.some(v => v !== '' && v != null);
      }
      return createJsonResponse({ ok: true, exists: exists, row: targetRow });
    }

    const result = handleSubmit(payload);
    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({ ok: false, error: err.message || String(err) });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getConfig() {
  return {
    lokasiList: [
      { value: 'PETI_SEJUK_1', label: 'Peti Sejuk 1' },
      { value: 'PETI_SEJUK_2', label: 'Peti Sejuk 2' },
      { value: 'PETI_SEJUK_3', label: 'Peti Sejuk 3' },
      { value: 'PETI_SEJUK_4', label: 'Peti Sejuk 4' }
    ],
    ranges: {
      PETI_SEJUK_1: { min: -4, max: 17, normalMin: 2, normalMax: 8 },
      PETI_SEJUK_2: { min: -4, max: 17, normalMin: 2, normalMax: 8 },
      PETI_SEJUK_3: { min: -4, max: 17, normalMin: 2, normalMax: 8 },
      PETI_SEJUK_4: { min: -4, max: 17, normalMin: 2, normalMax: 8 }
    },
    perkara: [
      { value: 'A', label: 'A - TIADA PRODUK/BAHAN RANGKAIAN SEJUK DISIMPAN' },
      { value: 'B', label: 'B - TIDAK CUKUP BEKALAN ELEKTRIK' },
      { value: 'C', label: 'C - PETI SEJUK TIDAK BERFUNGSI DENGAN BETUL' },
      { value: 'D', label: 'D - PEMBANTU TEKNIK DIPANGGIL UNTUK PENAMBAHBAIKAN' },
      { value: 'E', label: 'E - PETI SEJUK DALAM PEMBAIKAN' }
    ]
  };
}

function handleSubmit(payload) {
  if (!payload || !payload.lokasi) throw new Error('Sila pilih lokasi peti sejuk.');

  const ssId = idForLokasi_(payload.lokasi);
  const ss = SpreadsheetApp.openById(ssId);
  
  let rekod = ss.getSheetByName(SHEET_TAB_REKOD);
  if (!rekod) {
    rekod = initRekodSheet_(ss);
  }

  const targetRow = rowFor_(payload.date, payload.slot);

  const existing = rekod.getRange(targetRow, 3, 1, 5).getValues()[0];
  const hasData = existing.some(v => v !== '' && v != null);

  if (hasData && !payload.overwrite) {
    return { ok: false, already: true, row: targetRow };
  }

  const dateObj = new Date(payload.date);
  const dateFormatted = formatDateDDMMYYYY_(dateObj);

  const values = [[
    dateFormatted,
    String(payload.slot).toUpperCase(),
    payload.min === '' ? '' : Number(payload.min),
    payload.semasa === '' ? '' : Number(payload.semasa),
    payload.max === '' ? '' : Number(payload.max),
    fullPerkaraFromLetter_(payload.perkara || ''),
    payload.nama ? String(payload.nama).trim() : ''
  ]];

  rekod.getRange(targetRow, 1, 1, 7).setValues(values);

  let incidentSaved = false;
  if (payload.incident && (payload.incident.enabled || payload.incident.note || payload.incident.officer)) {
    incidentSaved = appendCatatan_(ss, {
      date: payload.incident.date || payload.date,
      time: payload.incident.time || '',
      note: payload.incident.note || '',
      officer: payload.incident.officer || payload.nama || ''
    });
  }

  return { ok: true, row: targetRow, incidentSaved: incidentSaved };
}

function idForLokasi_(lokasi) {
  const key = String(lokasi || '').toUpperCase();
  if (SPREADSHEET_IDS[key]) {
    return SPREADSHEET_IDS[key];
  }
  throw new Error('Lokasi peti sejuk tidak sah: ' + lokasi);
}

function rowFor_(isoDate, slot) {
  const d = new Date(isoDate);
  if (isNaN(d)) throw new Error('Tarikh tidak sah.');
  
  const y = d.getFullYear();
  if (y < BASE_YEAR || y > LAST_YEAR) {
    throw new Error(`Tarikh mesti antara tahun ${BASE_YEAR} hingga ${LAST_YEAR}.`);
  }
  
  d.setHours(0,0,0,0);
  const base = new Date(BASE_YEAR, 0, 1);
  const days = Math.floor((d - base) / (1000 * 60 * 60 * 24));
  const offset = (String(slot).toUpperCase() === 'PM') ? 1 : 0;
  
  return START_ROW + (days * 2) + offset;
}

function fullPerkaraFromLetter_(letter) {
  const val = String(letter || '').toUpperCase();
  const map = {
    A: 'A - TIADA PRODUK/BAHAN RANGKAIAN SEJUK DISIMPAN',
    B: 'B - TIDAK CUKUP BEKALAN ELEKTRIK',
    C: 'C - PETI SEJUK TIDAK BERFUNGSI DENGAN BETUL',
    D: 'D - PEMBANTU TEKNIK DIPANGGIL UNTUK PENAMBAHBAIKAN',
    E: 'E - PETI SEJUK DALAM PEMBAIKAN'
  };
  return map[val] || (letter ? String(letter).trim() : '');
}

function appendCatatan_(ss, payload) {
  let sheet = ss.getSheetByName(SHEET_TAB_CATATAN);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_TAB_CATATAN);
    const headers = [['Tarikh', 'Masa', 'Perkara/Penjelasan', 'Nama Pegawai']];
    sheet.getRange(1, 1, 1, 4).setValues(headers);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#FFF2CC');
  }
  
  const dateObj = new Date(payload.date);
  const dateFormatted = formatDateDDMMYYYY_(dateObj);

  sheet.appendRow([
    dateFormatted,
    payload.time || '',
    (payload.note || '').trim(),
    (payload.officer || '').trim()
  ]);
  return true;
}

function initRekodSheet_(ss) {
  let sheet = ss.insertSheet(SHEET_TAB_REKOD);
  const headers = [['Tarikh', 'Waktu', 'Suhu Minimum', 'Suhu Semasa', 'Suhu Maksimum', 'Perkara', 'Tandatangan Pencatat']];
  sheet.getRange(1, 1, 1, 7).setValues(headers);
  sheet.getRange(1, 1, 1, 7).setFontWeight('bold').setBackground('#FCE5CD');
  return sheet;
}

function formatDateDDMMYYYY_(d) {
  if (isNaN(d)) return '';
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  return `${day}/${month}/${year}`;
}
