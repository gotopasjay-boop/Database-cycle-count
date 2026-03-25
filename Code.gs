/**
 * BACKEND SCRIPT (Code.gs)
 * Mengelola penyimpanan hasil hitung dan pengambilan data rekapitulasi.
 * 
 * ALUR PENYIMPANAN:
 * 1. Jika SKU Ditemukan di Master -> Disimpan di sheet 'SO_HARIAN'.
 * 2. Jika SKU Tidak Ditemukan (Input Manual) -> Disimpan di sheet 'SO_MANUAL'.
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Sistem Cycle Count')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ============================================================
// FUNGSI SIMPAN DATA
// ============================================================

/**
 * Fungsi: saveCount
 * Kondisi: SKU ditemukan di data Master.
 * Aksi: Simpan ke Sheet 'SO_HARIAN'.
 */
function saveCount(sku, stokAktual, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = ss.getSheetByName('MASTER') || ss.getSheetByName('MasterData');
  
  // TARGET SHEET: SO_HARIAN
  const logSheet = ss.getSheetByName('SO_HARIAN');
  
  if (!masterSheet || !logSheet) {
    return "Error: Sheet MASTER atau SO_HARIAN tidak ditemukan! Jalankan setupSheet.";
  }

  const dataMaster = masterSheet.getDataRange().getValues();
  let itemInfo = null;

  // Cari data barang di Master
  for (let i = 1; i < dataMaster.length; i++) {
    if (dataMaster[i][0].toString().toLowerCase() === sku.toString().toLowerCase()) {
      itemInfo = {
        sku: dataMaster[i][0],
        nama: dataMaster[i][1],
        lokasi: dataMaster[i][2],
        stokSistem: Number(dataMaster[i][3] || 0)
      };
      break;
    }
  }

  // Jika tidak ketemu, perintahkan frontend untuk pakai mode manual
  if (!itemInfo) return "Error: SKU tidak ditemukan! Gunakan mode Input Manual.";

  const selisih = Number(stokAktual) - itemInfo.stokSistem;
  const timestamp = new Date();

  // Data untuk SO_HARIAN
  const rowData = [
    timestamp, 
    itemInfo.sku, 
    itemInfo.nama, 
    itemInfo.lokasi, 
    itemInfo.stokSistem, 
    stokAktual, 
    selisih, 
    user,
    "MASTER" // Keterangan
  ];

  logSheet.appendRow(rowData);
  SpreadsheetApp.flush();
  
  return "Berhasil disimpan ke SO_HARIAN (Master): " + itemInfo.nama;
}

/**
 * Fungsi: saveManualEntry
 * Kondisi: Input manual (SKU tidak ada di Master).
 * Aksi: Simpan ke Sheet 'SO_MANUAL'.
 */
function saveManualEntry(sku, nama, lokasi, stokSistem, stokAktual, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // TARGET SHEET: SO_MANUAL (Berbeda dengan saveCount)
  const logSheet = ss.getSheetByName('SO_MANUAL');
  
  if (!logSheet) {
    return "Error: Sheet SO_MANUAL tidak ditemukan! Jalankan setupSheet.";
  }

  if (!sku || !nama) return "Error: SKU dan Nama wajib diisi!";

  const sysStock = Number(stokSistem) || 0;
  const actStock = Number(stokAktual) || 0;
  const selisih = actStock - sysStock;
  const timestamp = new Date();

  // Data untuk SO_MANUAL
  // Header: TIMESTAMP, SKU, NAMA BARANG, LOKASI, STOCK SISTEM, STOCK AKTUAL, SELISIH, USER INPUT, KET.
  const rowData = [
    timestamp, 
    sku, 
    nama, 
    lokasi, 
    sysStock, 
    actStock, 
    selisih, 
    user,
    "MANUAL" // Keterangan
  ];

  logSheet.appendRow(rowData);
  SpreadsheetApp.flush();
  
  return "Berhasil disimpan ke SO_MANUAL: " + nama;
}

// ============================================================
// FUNGSI UTILITAS
// ============================================================

// Ambil data laporan (Default: SO_HARIAN untuk tampilan web)
function getLogData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SO_HARIAN');
  if (!sheet) return [];
  
  const values = sheet.getDataRange().getValues();
  const formattedData = values.map((row, index) => {
    if (index === 0) return row;
    const newRow = [...row];
    if (row[0] instanceof Date) {
      newRow[0] = row[0].toISOString();
    }
    return newRow;
  });

  return formattedData;
}

function checkSheetStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('SO_HARIAN');
  if (!sheet) return { status: "Error", message: "Sheet SO_HARIAN tidak ditemukan!" };
  return { status: "OK", message: "Database Terhubung (" + (sheet.getLastRow() - 1) + " data Master)" };
}

function getDownloadUrls() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ssId = ss.getId();
  const sheet = ss.getSheetByName('SO_HARIAN'); // Link download mengambil sheet SO_HARIAN
  const gid = sheet ? sheet.getSheetId() : 0;

  const baseUrl = "https://docs.google.com/spreadsheets/d/" + ssId + "/export?";
  return {
    pdf: baseUrl + "exportFormat=pdf&format=pdf&gid=" + gid + "&size=A4&portrait=true&fitw=true",
    excel: baseUrl + "exportFormat=xlsx&format=xlsx&gid=" + gid
  };
}

// Fungsi pencarian barang
function searchItem(sku) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('MASTER') || ss.getSheetByName('MasterData');
  if (!sheet) return { status: 'error' };
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === sku.toString().toLowerCase()) {
      return {
        status: 'found',
        sku: data[i][0],
        nama: data[i][1],
        lokasi: data[i][2],
        stokSistem: data[i][3] || 0
      };
    }
  }
  // Jika tidak ketemu, kembalikan not_found -> Frontend akan mengaktifkan mode Manual
  return { status: 'not_found' };
}

// ============================================================
// FUNGSI SETUP SHEET (Jalankan Sekali)
// ============================================================

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Buat Sheet MASTER
  let masterSheet = ss.getSheetByName('MASTER');
  if (!masterSheet) {
    masterSheet = ss.insertSheet('MASTER');
    masterSheet.appendRow(["SKU", "Nama", "Lokasi", "Stok"]);
    masterSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#d9d9d9");
  }

  // 2. Buat Sheet SO_HARIAN (Untuk data Master)
  let logSheetHarian = ss.getSheetByName('SO_HARIAN');
  if (!logSheetHarian) {
    logSheetHarian = ss.insertSheet('SO_HARIAN');
    logSheetHarian.appendRow(["Waktu", "SKU", "Nama", "Lokasi", "Sistem", "Aktual", "Selisih", "Petugas", "Keterangan"]);
    logSheetHarian.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#cfe2ff"); // Biru
  }
  
  // 3. Buat Sheet SO_MANUAL (Untuk input manual)
  let logSheetManual = ss.getSheetByName('SO_MANUAL');
  if (!logSheetManual) {
    logSheetManual = ss.insertSheet('SO_MANUAL');
    // Header sesuai permintaan
    logSheetManual.appendRow(["TIMESTAMP", "SKU", "NAMA BARANG", "LOKASI", "STOCK SISTEM", "STOCK AKTUAL", "SELISIH", "USER INPUT", "KET."]);
    logSheetManual.getRange(1, 1, 1, 9).setFontWeight("bold").setBackground("#fce5cd"); // Oranye
  }
  
  SpreadsheetApp.flush();
  
  SpreadsheetApp.getUi().alert("Setup Selesai!\n\n- MASTER\n- SO_HARIAN (Data dari Master)\n- SO_MANUAL (Input Manual)");
}
