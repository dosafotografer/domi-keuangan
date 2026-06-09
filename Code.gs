// ============================================================
// KONFIGURASI - Ganti SPREADSHEET_ID dengan ID Google Sheet Anda
// ============================================================
const SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET_ANDA';
const SHEET_NAME = 'Transaksi';
const CONFIG_SHEET_NAME = 'Konfigurasi';

function doGet(e) {
  return HtmlService.createHtmlOutput('Expense Tracker API aktif.');
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'addTransaction') {
      return addTransaction(data);
    } else if (action === 'getTransactions') {
      return getTransactions();
    } else if (action === 'getConfig') {
      return getConfig();
    } else if (action === 'setLimit') {
      return setLimit(data);
    } else if (action === 'deleteTransaction') {
      return deleteTransaction(data);
    }

    return respond({ success: false, message: 'Action tidak dikenal.' });
  } catch (err) {
    return respond({ success: false, message: err.toString() });
  }
}

// ---- Tambah Transaksi ----
function addTransaction(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_NAME);

  const now = new Date();
  const row = [
    Utilities.formatDate(now, 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss'),
    data.nominal,
    data.sumber,
    data.keterangan
  ];

  sheet.appendRow(row);
  return respond({ success: true, message: 'Transaksi berhasil disimpan.' });
}

// ---- Ambil Semua Transaksi ----
function getTransactions() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  const data = sheet.getDataRange().getValues();

  if (data.length <= 1) {
    return respond({ success: true, transactions: [] });
  }

  const rows = data.slice(1).map((row, index) => ({
    rowIndex: index + 2,
    tanggal: row[0],
    nominal: row[1],
    sumber: row[2],
    keterangan: row[3]
  }));

  // Sort descending by tanggal
  rows.sort((a, b) => new Date(b.tanggal) - new Date(a.tanggal));

  return respond({ success: true, transactions: rows });
}

// ---- Hapus Transaksi ----
function deleteTransaction(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateSheet(ss, SHEET_NAME);
  sheet.deleteRow(data.rowIndex);
  return respond({ success: true, message: 'Transaksi dihapus.' });
}

// ---- Ambil Konfigurasi Limit ----
function getConfig() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateConfigSheet(ss);
  const limit = sheet.getRange('B1').getValue();
  return respond({ success: true, limit: limit || 0 });
}

// ---- Set Limit Bulanan ----
function setLimit(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = getOrCreateConfigSheet(ss);
  sheet.getRange('A1').setValue('Limit Bulanan');
  sheet.getRange('B1').setValue(data.limit);
  return respond({ success: true, message: 'Limit berhasil disimpan.' });
}

// ---- Helper: Buat atau ambil sheet Transaksi ----
function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(['Tanggal', 'Nominal', 'Sumber', 'Keterangan']);
    sheet.getRange('1:1').setFontWeight('bold');
  }
  return sheet;
}

// ---- Helper: Buat atau ambil sheet Konfigurasi ----
function getOrCreateConfigSheet(ss) {
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.getRange('A1').setValue('Limit Bulanan');
    sheet.getRange('B1').setValue(0);
  }
  return sheet;
}

// ---- Helper: Format response JSON ----
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
