const SHEET_PRODUK = 'produk';
const SHEET_PENJUALAN = 'penjualan';
const SHEET_SETTINGS = 'settings';

// FIX: Tambah kolom 'sku' yang dipakai index.html
const HEADERS_PRODUK = ['id', 'nama', 'sku', 'kategori', 'harga_beli', 'harga_jual', 'stok', 'image_url'];

// FIX: Tambah kolom 'uang_diterima' yang dikirim index.html saat checkout
const HEADERS_PENJUALAN = ['id', 'tanggal', 'item_list', 'total_bayar', 'total_untung', 'metode_pembayaran', 'uang_diterima'];

const HEADERS_SETTINGS = ['id', 'value'];

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function rowToObject(headers, row) {
  const obj = {};
  headers.forEach((h, i) => obj[h] = row[i] !== undefined ? String(row[i]) : '');
  return obj;
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kasier - Kasir Warung')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    const payload = body.payload;
    let result;

    if (action === 'getData') {
      result = getData(payload.sheet);
    } else if (action === 'createProduct') {
      result = createProduct(payload);
    } else if (action === 'updateProduct') {
      result = updateProduct(payload.id, payload);
    } else if (action === 'deleteProduct') {
      result = deleteProduct(payload.id);
    } else if (action === 'processTransaction') {
      result = processTransaction(payload);
    } else if (action === 'saveSettings') {
      result = saveSettings(payload);
    } else {
      throw new Error('Aksi tidak dikenali.');
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Setup sheet Produk
  let sProduk = ss.getSheetByName(SHEET_PRODUK) || ss.insertSheet(SHEET_PRODUK);
  if (sProduk.getLastRow() === 0) sProduk.appendRow(HEADERS_PRODUK);

  // Setup sheet Penjualan
  let sPenj = ss.getSheetByName(SHEET_PENJUALAN) || ss.insertSheet(SHEET_PENJUALAN);
  if (sPenj.getLastRow() === 0) sPenj.appendRow(HEADERS_PENJUALAN);

  // Setup sheet Settings
  let sSet = ss.getSheetByName(SHEET_SETTINGS) || ss.insertSheet(SHEET_SETTINGS);
  if (sSet.getLastRow() === 0) {
    sSet.appendRow(HEADERS_SETTINGS);
    sSet.appendRow(['storeName', 'Warung Barokah']);
    sSet.appendRow(['storeAddress', 'Jl. Contoh No.123']);
    sSet.appendRow(['storePhone', '08123456789']);
    sSet.appendRow(['themeColor', '#10b981']);
    sSet.appendRow(['qrisImage', '']); // FIX: tambah key qrisImage
  }

  // Dummy data produk
  if (sProduk.getLastRow() <= 1) {
    sProduk.appendRow([Utilities.getUuid(), 'Kopi Hitam', 'KOP-001', 'Minuman', 3000, 4000, 50, '']);
    sProduk.appendRow([Utilities.getUuid(), 'Mie Instan Goreng', 'MIE-001', 'Makanan', 2800, 3500, 100, '']);
    sProduk.appendRow([Utilities.getUuid(), 'Telur Ayam 1kg', 'TELUR-01', 'Sembako', 26000, 28000, 10, '']);
    Logger.log('Setup selesai dengan data dummy.');
  } else {
    Logger.log('Sheet sudah siap.');
  }
}

// ─────────────────────────────────────────────
// GET DATA
// FIX: khusus sheet 'settings', return array of {id, value}
// sesuai cara index.html membacanya:
//   settingsData.forEach(s => settingsObj[s.id] = s.value)
// ─────────────────────────────────────────────
function getData(sheetName) {
  const sheet = getSheet(sheetName);
  if (!sheet) throw new Error('Sheet ' + sheetName + ' tidak ditemukan.');

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows[0];
  return rows.slice(1).map(row => rowToObject(headers, row));
}

// ─────────────────────────────────────────────
// CREATE PRODUCT
// FIX: include field 'sku' dalam mapping
// ─────────────────────────────────────────────
function createProduct(payload) {
  const sheet = getSheet(SHEET_PRODUK);
  const id = Utilities.getUuid();
  const row = HEADERS_PRODUK.map(h => {
    if (h === 'id') return id;
    return payload[h] !== undefined ? payload[h] : '';
  });
  sheet.appendRow(row);
  return rowToObject(HEADERS_PRODUK, row);
}

// ─────────────────────────────────────────────
// UPDATE PRODUCT
// ─────────────────────────────────────────────
function updateProduct(id, payload) {
  const sheet = getSheet(SHEET_PRODUK);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idCol = headers.indexOf('id');

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][idCol]) === String(id)) {
      headers.forEach((h, j) => {
        if (payload[h] !== undefined) {
          sheet.getRange(i + 1, j + 1).setValue(payload[h]);
        }
      });
      return true;
    }
  }
  throw new Error('Produk tidak ditemukan.');
}

// ─────────────────────────────────────────────
// DELETE PRODUCT
// ─────────────────────────────────────────────
function deleteProduct(id) {
  const sheet = getSheet(SHEET_PRODUK);
  const data = sheet.getDataRange().getValues();
  const idCol = data[0].indexOf('id');

  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][idCol]) === String(id)) {
      sheet.deleteRow(i + 1);
      return true;
    }
  }
  throw new Error('Produk tidak ditemukan.');
}

// ─────────────────────────────────────────────
// PROCESS TRANSACTION
// FIX: include 'uang_diterima' dalam mapping ke sheet
// ─────────────────────────────────────────────
function processTransaction(payload) {
  const sPenj = getSheet(SHEET_PENJUALAN);
  const sProd = getSheet(SHEET_PRODUK);

  const id = Utilities.getUuid();
  const now = new Date().toISOString();

  // Catat ke sheet Penjualan
  const rowPenj = HEADERS_PENJUALAN.map(h => {
    if (h === 'id') return id;
    if (h === 'tanggal') return now;
    return payload[h] !== undefined ? payload[h] : '';
  });
  sPenj.appendRow(rowPenj);

  // Kurangi stok produk
  let items = [];
  try {
    items = JSON.parse(payload.item_list);
  } catch (e) {
    throw new Error('Format item_list tidak valid.');
  }

  const prodData = sProd.getDataRange().getValues();
  const prodHeaders = prodData[0];
  const idColProd = prodHeaders.indexOf('id');
  const stokColProd = prodHeaders.indexOf('stok');

  items.forEach(item => {
    for (let i = 1; i < prodData.length; i++) {
      if (String(prodData[i][idColProd]) === String(item.id)) {
        const currentStok = parseInt(prodData[i][stokColProd]) || 0;
        const newStok = Math.max(0, currentStok - item.qty);
        sProd.getRange(i + 1, stokColProd + 1).setValue(newStok);
        break;
      }
    }
  });

  return rowToObject(HEADERS_PENJUALAN, rowPenj);
}

// ─────────────────────────────────────────────
// SAVE SETTINGS
// FIX: tambah support key 'qrisImage'
// Payload dari index.html: { storeName, storeAddress, storePhone, themeColor, qrisImage }
// ─────────────────────────────────────────────
function saveSettings(payload) {
  const sheet = getSheet(SHEET_SETTINGS);
  const data = sheet.getDataRange().getValues();

  Object.keys(payload).forEach(key => {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === key) {
        sheet.getRange(i + 1, 2).setValue(payload[key]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([key, payload[key]]);
    }
  });

  return true;
}
