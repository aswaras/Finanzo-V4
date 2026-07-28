/**
 * ====================================================================
 * FINANZO BACKEND v2 — Google Apps Script + Google Sheets (Tabel Normal)
 * ====================================================================
 *
 * PERUBAHAN UTAMA DARI VERSI SEBELUMNYA:
 * - Data TIDAK LAGI disimpan sebagai satu blok JSON raksasa di sel A1.
 * - Data sekarang disimpan sebagai TABEL SPREADSHEET NORMAL (baris & kolom),
 *   jadi bisa langsung dibuka/dibaca/di-filter di Google Sheets maupun
 *   di-download sebagai Excel (.xlsx) tanpa alat bantu apa pun.
 * - Sheet yang dipakai:
 *     1. "Transaksi" -> catatan pemasukan/pengeluaran (state.records)
 *     2. "Nota"       -> data header nota/invoice (state.notas, tanpa items)
 *     3. "Nota_Item"  -> rincian item per nota (relasi lewat kolom "ID Nota")
 *     4. "Profil"     -> data toko, tema, PIN, dsb dalam format Key-Value
 *     5. "Log"        -> riwayat setiap kali data disimpan (audit trail)
 * - Kontrak API (doGet / doPost) TIDAK BERUBAH sama sekali, jadi
 *   index.html tidak perlu diubah dan seluruh fitur lama tetap jalan
 *   persis seperti sebelumnya:
 *     - GET  -> { status: "success", data: {...state...} }
 *     - POST -> body JSON = seluruh objek state -> { status: "success" }
 * - Migrasi OTOMATIS: kalau Spreadsheet Anda masih pakai sheet "Database"
 *   lama (format JSON di A1), data itu otomatis dipindah ke tabel baru
 *   saat pertama kali dibuka/diakses. Sheet lama tidak dihapus, hanya
 *   diberi nama "Database_OLD_BACKUP" supaya data lama tetap aman.
 * - Tambahan: menu "Finanzo" di Spreadsheet dengan fungsi
 *   "Backup Sekarang (Export ke Excel)" -> membuat file .xlsx asli di
 *   Google Drive, folder "Finanzo Backup", lengkap dengan semua sheet.
 *
 * CARA PAKAI (SAMA SEPERTI SEBELUMNYA):
 * 1. Buka https://sheets.google.com -> buat Spreadsheet baru.
 * 2. Extensions/Ekstensi -> Apps Script -> hapus kode default -> tempel file ini.
 * 3. Simpan project.
 * 4. Deploy -> New deployment -> pilih tipe "Web app".
 *      - Execute as    : Me
 *      - Who has access: Anyone
 * 5. Copy "Web app URL" (...../exec) ke variabel API_URL di index.html.
 * 6. Setiap kali kode ini diedit ulang, WAJIB buat "New version" lagi
 *    lewat Deploy -> Manage deployments -> Edit -> New version -> Deploy.
 */

// ==========================================
// 0. KONFIGURASI NAMA SHEET & HEADER KOLOM
// ==========================================
var SHEET_TRANSAKSI = "Transaksi";
var SHEET_NOTA = "Nota";
var SHEET_NOTA_ITEM = "Nota_Item";
var SHEET_PROFIL = "Profil";
var SHEET_LOG = "Log";
var SHEET_OLD_DATABASE = "Database"; // nama sheet versi lama (format JSON)

var HEADER_TRANSAKSI = ["ID", "Tanggal", "Tipe", "Nominal", "Keterangan"];
var HEADER_NOTA = [
  "ID", "Pelanggan", "No HP", "Tanggal", "Deadline",
  "DP", "Subtotal", "Tipe Diskon", "Nilai Diskon", "Nominal Diskon", "Total"
];
var HEADER_NOTA_ITEM = ["ID Nota", "Nama Item", "Deskripsi", "Qty", "Harga", "Total"];
var HEADER_PROFIL = ["Key", "Value", "Keterangan"];

// Baris default sheet "Profil" (Key -> field state, Keterangan hanya info untuk pemilik sheet)
var PROFIL_KEYS = [
  ["nama_pemilik", "Pemilik", "Nama pemilik/pengguna aplikasi"],
  ["nama_toko", "Toko Finanzo", "Nama toko/usaha, tampil di nota"],
  ["alamat_toko", "Jl. Sukses No. 1", "Alamat toko, tampil di nota"],
  ["footer_nota", "Terima kasih.\nTransfer: BCA 1234567", "Catatan kaki di bawah nota"],
  ["warna_tema", "#2563eb", "Kode warna HEX aksen aplikasi"],
  ["mode_tampilan", "light", "light atau dark"],
  ["tampilan_aktif", "home", "Tab terakhir yang aktif (home/catat/nota/setup)"],
  ["pin_keamanan", "", "PIN 4 digit untuk kunci aplikasi (kosongkan = belum diset)"]
];

// ==========================================
// ENTRY POINT: GET (ambil / load data)
// ==========================================
function doGet(e) {
  try {
    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      migrateOldDatabaseIfNeeded_(ss);
      ensureAllSheets_(ss);
      var data = buildStateFromSheets_(ss);
      return jsonResponse_({ status: "success", data: data });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

// ==========================================
// ENTRY POINT: POST (simpan / save data)
// ==========================================
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Tidak ada data yang dikirim (body kosong).");
    }

    var body = e.postData.contents;
    var parsed = JSON.parse(body); // validasi format JSON dari frontend

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      ensureAllSheets_(ss);
      writeStateToSheets_(ss, parsed);
      logActivity_(ss, body.length);
      return jsonResponse_({ status: "success" });
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return jsonResponse_({ status: "error", message: err.toString() });
  }
}

// ==========================================
// 1. PEMBUATAN / PEMASTIAN STRUKTUR SHEET
// ==========================================
function ensureAllSheets_(ss) {
  ensureSheetWithHeader_(ss, SHEET_TRANSAKSI, HEADER_TRANSAKSI);
  ensureSheetWithHeader_(ss, SHEET_NOTA, HEADER_NOTA);
  ensureSheetWithHeader_(ss, SHEET_NOTA_ITEM, HEADER_NOTA_ITEM);
  ensureProfilSheet_(ss);
  ensureLogSheet_(ss);
}

function ensureSheetWithHeader_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  var isNew = false;
  if (!sheet) {
    sheet = ss.insertSheet(name);
    isNew = true;
  }
  if (isNew || sheet.getRange(1, 1, 1, headers.length).getValues()[0].join("|") !== headers.join("|")) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight("bold")
      .setBackground("#1e293b")
      .setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  // Format kolom sesuai jenis sheet supaya angka/tanggal tidak "diacak" otomatis oleh Sheets
  if (name === SHEET_TRANSAKSI) {
    sheet.getRange("A:A").setNumberFormat("@"); // ID -> teks
    sheet.getRange("B:B").setNumberFormat("@"); // Tanggal -> teks (format yyyy-mm-dd apa adanya)
    sheet.getRange("D:D").setNumberFormat("#,##0"); // Nominal -> angka ribuan
    sheet.setColumnWidths(1, 5, 130);
    sheet.setColumnWidth(5, 260);
  } else if (name === SHEET_NOTA) {
    sheet.getRange("A:A").setNumberFormat("@");
    sheet.getRange("D:E").setNumberFormat("@"); // Tanggal & Deadline -> teks
    sheet.getRange("F:F").setNumberFormat("#,##0"); // DP
    sheet.getRange("G:G").setNumberFormat("#,##0"); // Subtotal
    sheet.getRange("I:I").setNumberFormat("#,##0"); // Nilai Diskon (angka rupiah, tetap ditulis walau persen)
    sheet.getRange("J:J").setNumberFormat("#,##0"); // Nominal Diskon
    sheet.getRange("K:K").setNumberFormat("#,##0"); // Total
    sheet.setColumnWidths(1, 11, 130);
    sheet.setColumnWidth(2, 180);
  } else if (name === SHEET_NOTA_ITEM) {
    sheet.getRange("A:A").setNumberFormat("@");
    sheet.getRange("D:D").setNumberFormat("#,##0"); // Qty
    sheet.getRange("E:E").setNumberFormat("#,##0"); // Harga
    sheet.getRange("F:F").setNumberFormat("#,##0"); // Total
    sheet.setColumnWidths(1, 6, 150);
  }
}

function ensureProfilSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_PROFIL);
  var isNew = false;
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROFIL);
    isNew = true;
  }
  if (isNew) {
    sheet.getRange(1, 1, 1, HEADER_PROFIL.length).setValues([HEADER_PROFIL]);
    sheet.getRange(1, 1, 1, HEADER_PROFIL.length)
      .setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, PROFIL_KEYS.length, 3).setValues(PROFIL_KEYS);
    sheet.getRange("A:A").setNumberFormat("@");
    sheet.getRange("B:B").setNumberFormat("@");
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 300);
    sheet.setColumnWidth(3, 320);
  } else {
    // Pastikan semua key wajib ada (misal setelah update kode ini di sheet lama)
    var existing = sheet.getDataRange().getValues();
    var existingKeys = {};
    for (var i = 1; i < existing.length; i++) existingKeys[existing[i][0]] = true;
    var missing = PROFIL_KEYS.filter(function (row) { return !existingKeys[row[0]]; });
    if (missing.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, missing.length, 3).setValues(missing);
    }
  }
}

function ensureLogSheet_(ss) {
  var sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LOG);
    sheet.appendRow(["Waktu", "Ukuran Data (karakter)"]);
    sheet.getRange(1, 1, 1, 2).setFontWeight("bold").setBackground("#1e293b").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }
}

// ==========================================
// 2. MIGRASI OTOMATIS DARI FORMAT LAMA (JSON di A1)
// ==========================================
function migrateOldDatabaseIfNeeded_(ss) {
  var oldSheet = ss.getSheetByName(SHEET_OLD_DATABASE);
  if (!oldSheet) return; // tidak ada sheet lama, tidak perlu migrasi

  // Kalau sheet tabel baru sudah pernah terisi data, jangan timpa (anggap migrasi sudah pernah jalan)
  var trxSheet = ss.getSheetByName(SHEET_TRANSAKSI);
  if (trxSheet && trxSheet.getLastRow() > 1) {
    return;
  }

  try {
    var raw = oldSheet.getRange("A1").getValue();
    if (!raw) return;
    var oldState = JSON.parse(raw);

    ensureAllSheets_(ss);
    writeStateToSheets_(ss, oldState);

    // Amankan sheet lama (bukan dihapus, cuma diganti nama & disembunyikan)
    oldSheet.setName("Database_OLD_BACKUP");
    oldSheet.hideSheet();

    Logger.log("Migrasi dari format JSON lama ke tabel baru berhasil.");
  } catch (err) {
    Logger.log("Migrasi gagal (data lama mungkin sudah bukan JSON valid): " + err.toString());
  }
}

// ==========================================
// 3. BACA TABEL -> BENTUK OBJEK STATE (untuk doGet)
// ==========================================
function buildStateFromSheets_(ss) {
  var profil = readProfil_(ss);
  return {
    user: {
      name: profil.nama_pemilik,
      shop: profil.nama_toko,
      address: profil.alamat_toko,
      footer: profil.footer_nota,
      color: profil.warna_tema
    },
    theme: profil.mode_tampilan || "light",
    activeView: profil.tampilan_aktif || "home",
    records: readRecords_(ss),
    notas: readNotas_(ss),
    security: { pin: profil.pin_keamanan || "" }
  };
}

function readProfil_(ss) {
  var sheet = ss.getSheetByName(SHEET_PROFIL);
  var map = {};
  PROFIL_KEYS.forEach(function (row) { map[row[0]] = row[1]; }); // default dulu
  if (!sheet) return map;
  var values = sheet.getDataRange().getValues();
  for (var i = 1; i < values.length; i++) {
    var key = values[i][0];
    var val = values[i][1];
    if (key) map[key] = val;
  }
  return map;
}

function readRecords_(ss) {
  var sheet = ss.getSheetByName(SHEET_TRANSAKSI);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, HEADER_TRANSAKSI.length).getValues();
  var records = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue; // baris kosong dilewati
    records.push({
      id: String(row[0]),
      date: formatDateCell_(row[1]),
      type: String(row[2]),
      amount: Number(row[3]) || 0,
      note: row[4] ? String(row[4]) : ""
    });
  }
  return records;
}

function readNotas_(ss) {
  var notaSheet = ss.getSheetByName(SHEET_NOTA);
  var itemSheet = ss.getSheetByName(SHEET_NOTA_ITEM);

  var itemsByNotaId = {};
  if (itemSheet && itemSheet.getLastRow() > 1) {
    var itemValues = itemSheet.getRange(2, 1, itemSheet.getLastRow() - 1, HEADER_NOTA_ITEM.length).getValues();
    for (var j = 0; j < itemValues.length; j++) {
      var irow = itemValues[j];
      if (!irow[0]) continue;
      var notaId = String(irow[0]);
      if (!itemsByNotaId[notaId]) itemsByNotaId[notaId] = [];
      itemsByNotaId[notaId].push({
        name: String(irow[1] || ""),
        description: String(irow[2] || ""),
        qty: Number(irow[3]) || 0,
        price: Number(irow[4]) || 0
      });
    }
  }

  if (!notaSheet || notaSheet.getLastRow() < 2) return [];
  var values = notaSheet.getRange(2, 1, notaSheet.getLastRow() - 1, HEADER_NOTA.length).getValues();
  var notas = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    if (!row[0]) continue;
    var id = String(row[0]);
    notas.push({
      id: id,
      customer: String(row[1] || ""),
      phone: String(row[2] || ""),
      date: formatDateCell_(row[3]),
      deadline: formatDateCell_(row[4]),
      dp: Number(row[5]) || 0,
      subtotal: Number(row[6]) || 0,
      discountType: String(row[7] || "percent"),
      discountValue: Number(row[8]) || 0,
      discountAmount: Number(row[9]) || 0,
      total: Number(row[10]) || 0,
      items: itemsByNotaId[id] || []
    });
  }
  return notas;
}

// Sheets kadang otomatis mengubah teks tanggal jadi objek Date. Fungsi ini
// menjaga supaya yang dikembalikan ke frontend tetap string "yyyy-MM-dd".
function formatDateCell_(cell) {
  if (!cell) return "";
  if (Object.prototype.toString.call(cell) === "[object Date]") {
    return Utilities.formatDate(cell, Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd");
  }
  return String(cell);
}

// ==========================================
// 4. TULIS STATE -> TABEL (untuk doPost)
//    Strategi: replace-all per sheet (paling aman, karena frontend
//    selalu mengirim seluruh state, bukan perubahan sebagian).
// ==========================================
function writeStateToSheets_(ss, state) {
  if (!state || typeof state !== "object") throw new Error("Data yang dikirim bukan objek state yang valid.");

  writeProfil_(ss, state);
  writeRecords_(ss, state.records || []);
  writeNotas_(ss, state.notas || []);
}

function writeProfil_(ss, state) {
  var sheet = ss.getSheetByName(SHEET_PROFIL);
  var user = state.user || {};
  var security = state.security || {};

  var values = {
    nama_pemilik: user.name || "",
    nama_toko: user.shop || "",
    alamat_toko: user.address || "",
    footer_nota: user.footer || "",
    warna_tema: user.color || "#2563eb",
    mode_tampilan: state.theme || "light",
    tampilan_aktif: state.activeView || "home",
    pin_keamanan: security.pin || ""
  };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var key = data[i][0];
    if (Object.prototype.hasOwnProperty.call(values, key)) {
      sheet.getRange(i + 1, 2).setValue(values[key]);
    }
  }
}

function writeRecords_(ss, records) {
  var sheet = ss.getSheetByName(SHEET_TRANSAKSI);
  clearDataRows_(sheet);
  if (!records.length) return;

  var rows = records.map(function (r) {
    return [r.id || "", r.date || "", r.type || "", Number(r.amount) || 0, r.note || ""];
  });
  sheet.getRange(2, 1, rows.length, HEADER_TRANSAKSI.length).setValues(rows);
}

function writeNotas_(ss, notas) {
  var notaSheet = ss.getSheetByName(SHEET_NOTA);
  var itemSheet = ss.getSheetByName(SHEET_NOTA_ITEM);
  clearDataRows_(notaSheet);
  clearDataRows_(itemSheet);
  if (!notas.length) return;

  var notaRows = [];
  var itemRows = [];

  notas.forEach(function (n) {
    notaRows.push([
      n.id || "",
      n.customer || "",
      n.phone || "",
      n.date || "",
      n.deadline || "",
      Number(n.dp) || 0,
      Number(n.subtotal) || 0,
      n.discountType || "percent",
      Number(n.discountValue) || 0,
      Number(n.discountAmount) || 0,
      Number(n.total) || 0
    ]);

    (n.items || []).forEach(function (item) {
      itemRows.push([
        n.id || "",
        item.name || "",
        item.description || "",
        Number(item.qty) || 0,
        Number(item.price) || 0,
        (Number(item.qty) || 0) * (Number(item.price) || 0)
      ]);
    });
  });

  if (notaRows.length) notaSheet.getRange(2, 1, notaRows.length, HEADER_NOTA.length).setValues(notaRows);
  if (itemRows.length) itemSheet.getRange(2, 1, itemRows.length, HEADER_NOTA_ITEM.length).setValues(itemRows);
}

function clearDataRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

// ==========================================
// 5. LOG AKTIVITAS
// ==========================================
function logActivity_(ss, dataLength) {
  try {
    var logSheet = ss.getSheetByName(SHEET_LOG);
    if (!logSheet) return;
    logSheet.appendRow([new Date(), dataLength]);

    // Batasi log maksimal 200 baris terakhir supaya sheet tidak membengkak
    var maxRows = 201;
    var totalRows = logSheet.getLastRow();
    if (totalRows > maxRows) {
      logSheet.deleteRows(2, totalRows - maxRows);
    }
  } catch (err) {
    Logger.log("Gagal mencatat log: " + err.toString());
  }
}

// ==========================================
// 6. MENU KHUSUS DI SPREADSHEET + BACKUP EXCEL
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("⚙️ Finanzo")
    .addItem("🗂 Perbaiki / Inisialisasi Struktur Sheet", "menuInitSheets_")
    .addItem("💾 Backup Sekarang (Export ke Excel di Drive)", "menuBackupToExcel_")
    .addSeparator()
    .addItem("ℹ️ Tentang Struktur Data", "menuShowInfo_")
    .addToUi();
}

function menuInitSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  migrateOldDatabaseIfNeeded_(ss);
  ensureAllSheets_(ss);
  SpreadsheetApp.getUi().alert("Selesai. Struktur sheet Transaksi, Nota, Nota_Item, Profil, dan Log sudah siap/diperbaiki.");
}

/**
 * Backup ke Excel (.xlsx) ASLI — bukan JSON. Membuat file baru di Google
 * Drive folder "Finanzo Backup" berisi salinan penuh seluruh sheet
 * (Transaksi, Nota, Nota_Item, Profil, Log) sebagaimana terlihat di layar,
 * yang bisa langsung dibuka dengan Microsoft Excel / LibreOffice / dst.
 */
function menuBackupToExcel_() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureAllSheets_(ss);

    var url = "https://docs.google.com/spreadsheets/d/" + ss.getId() + "/export?format=xlsx";
    var response = UrlFetchApp.fetch(url, {
      headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error("Gagal export (kode " + response.getResponseCode() + ").");
    }

    var folder = getOrCreateBackupFolder_();
    var timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyyMMdd_HHmmss");
    var fileName = "Finanzo_Backup_" + timestamp + ".xlsx";
    var blob = response.getBlob().setName(fileName);
    var file = folder.createFile(blob);

    ui.alert(
      "Backup Berhasil!\n\n" +
      "File: " + fileName + "\n" +
      "Lokasi: Folder 'Finanzo Backup' di Google Drive\n" +
      "Link: " + file.getUrl()
    );
  } catch (err) {
    ui.alert("Backup gagal: " + err.toString());
  }
}

function getOrCreateBackupFolder_() {
  var folderName = "Finanzo Backup";
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(folderName);
}

function menuShowInfo_() {
  SpreadsheetApp.getUi().alert(
    "Struktur Data Finanzo\n\n" +
    "• Transaksi  -> catatan pemasukan/pengeluaran\n" +
    "• Nota       -> data header nota/invoice\n" +
    "• Nota_Item  -> rincian item tiap nota (dihubungkan lewat kolom 'ID Nota')\n" +
    "• Profil     -> data toko, tema, PIN (format Key-Value)\n" +
    "• Log        -> riwayat setiap kali aplikasi menyimpan data\n\n" +
    "Semua sheet ini AMAN dibuka & dibaca langsung. Untuk mengedit data, " +
    "tetap disarankan lewat aplikasi Finanzo supaya format & relasi antar " +
    "tabel tidak rusak."
  );
}

// ==========================================
// 7. HELPER: Bungkus response jadi JSON
// ==========================================
function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
