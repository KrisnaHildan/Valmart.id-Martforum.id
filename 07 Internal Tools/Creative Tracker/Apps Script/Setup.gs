/**
 * =========================================================================
 * SETUP.GS — Menu Kustom, Setup Aman (Idempotent), Format Database, & Reset
 * =========================================================================
 */

/**
 * Membangun menu kustom saat Spreadsheet dibuka.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(APP_CONSTANTS.MENU_NAME)
    .addItem("🛠️ Setup Awal (Aman)", "setupInitial")
    .addItem("🔄 Refresh Semua (Dashboard & Kalender)", "refreshAllViews")
    .addSeparator()
    .addItem("📥 Sinkronisasi Naskah dari GitHub", "promptAndSyncFromGithub")
    .addItem("🔑 Atur GitHub Token", "promptGithubToken")
    .addSeparator()
    .addItem("🔔 Aktifkan / Matikan Ringkasan Harian", "toggleDailyDigest")
    .addItem("⏰ Pasang Trigger Otomatis", "installAllTriggers")
    .addSeparator()
    .addItem("⚠️ Reset Total (Backup + Reset)", "resetTotalWithConfirmation")
    .addToUi();
}

/**
 * Setup awal yang aman (idempotent): tidak menghapus data yang sudah ada.
 */
function setupInitial() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    showAlertSafe("Sedang Diproses", "Setup sedang dijalankan oleh pengguna lain. Coba lagi sebentar.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Pastikan 4 tab ada dengan urutan baku: Dashboard, Kalender, Database, Config
    let configSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.CONFIG);
    const isNewConfig = !configSheet;
    if (!configSheet) configSheet = ss.insertSheet(APP_CONSTANTS.SHEETS.CONFIG);
    setupConfigSheet(configSheet, isNewConfig);

    let dbSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.DATABASE);
    const isNewDb = !dbSheet;
    if (!dbSheet) dbSheet = ss.insertSheet(APP_CONSTANTS.SHEETS.DATABASE);
    formatDatabaseSheet(dbSheet, isNewDb);

    let calSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.CALENDAR);
    if (!calSheet) calSheet = ss.insertSheet(APP_CONSTANTS.SHEETS.CALENDAR);

    let dashSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.DASHBOARD);
    if (!dashSheet) dashSheet = ss.insertSheet(APP_CONSTANTS.SHEETS.DASHBOARD, 0);

    // Urutkan posisi tab: 1. Dashboard, 2. Kalender, 3. Database, 4. Config
    const orderedNames = [
      APP_CONSTANTS.SHEETS.DASHBOARD,
      APP_CONSTANTS.SHEETS.CALENDAR,
      APP_CONSTANTS.SHEETS.DATABASE,
      APP_CONSTANTS.SHEETS.CONFIG
    ];
    orderedNames.forEach((name, idx) => {
      const sh = ss.getSheetByName(name);
      if (sh) {
        ss.setActiveSheet(sh);
        ss.moveActiveSheet(idx + 1);
      }
    });

    renderCalendar();
    renderDashboard();
    ss.setActiveSheet(ss.getSheetByName(APP_CONSTANTS.SHEETS.DASHBOARD));

    showToastSafe("Setup awal selesai. Data yang ada tetap aman.", "✅ Setup Berhasil", 5);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menyegarkan Dashboard dan Kalender Visual sekaligus.
 */
function refreshAllViews() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    showToastSafe("Proses lain sedang berjalan, coba sebentar lagi.", "⏳ Tunggu", 4);
    return;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dbSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.DATABASE);
    if (dbSheet) {
      applyDatabaseValidationsAndRules(dbSheet);
    }
    renderCalendar();
    renderDashboard();
    showToastSafe("Dashboard dan Kalender Visual sudah diperbarui.", "🔄 Refresh Selesai", 4);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Reset Total dengan konfirmasi ketik "RESET" + membuat backup spreadsheet otomatis terlebih dahulu.
 */
function resetTotalWithConfirmation() {
  const ui = SpreadsheetApp.getUi();
  const prompt = ui.prompt(
    "⚠️ KONFIRMASI RESET TOTAL",
    "Tindakan ini akan membuat salinan BACKUP spreadsheet di Google Drive, lalu mengosongkan dan mereset seluruh 4 tab.\n\n" +
    "Ketik kata RESET (huruf besar semua) untuk melanjutkan:",
    ui.ButtonSet.OK_CANCEL
  );

  if (prompt.getSelectedButton() !== ui.Button.OK) return;
  if (prompt.getResponseText().trim() !== "RESET") {
    ui.alert("Dibatalkan", "Kata konfirmasi tidak cocok. Tidak ada data yang diubah.", ui.ButtonSet.OK);
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    ui.alert("Sedang Diproses", "Proses lain sedang berjalan. Coba lagi sebentar.", ui.ButtonSet.OK);
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const stamp = Utilities.formatDate(new Date(), APP_CONSTANTS.TIMEZONE, "yyyy-MM-dd HH:mm");
    const backupFile = ss.copy(`[BACKUP ${stamp}] ${ss.getName()}`);

    const configSheet = ensureSheet(ss, APP_CONSTANTS.SHEETS.CONFIG);
    setupConfigSheet(configSheet, true);

    const dbSheet = ensureSheet(ss, APP_CONSTANTS.SHEETS.DATABASE);
    formatDatabaseSheet(dbSheet, true);

    const calSheet = ensureSheet(ss, APP_CONSTANTS.SHEETS.CALENDAR);
    calSheet.clear();
    renderCalendar();

    const dashSheet = ensureSheet(ss, APP_CONSTANTS.SHEETS.DASHBOARD, 0);
    dashSheet.clear();
    renderDashboard();

    ui.alert(
      "✅ Reset Total Selesai",
      `Backup berhasil disimpan di Google Drive:\n${backupFile.getName()}\nURL: ${backupFile.getUrl()}\n\nSeluruh 4 tab sudah direset bersih.`,
      ui.ButtonSet.OK
    );
  } catch (err) {
    ui.alert("❌ Gagal Reset", "Terjadi kendala saat membuat backup/reset: " + err.message, ui.ButtonSet.OK);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memastikan sheet ada; buat baru jika belum.
 */
function ensureSheet(ss, name, index) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = (typeof index === "number") ? ss.insertSheet(name, index) : ss.insertSheet(name);
  }
  return sh;
}

/**
 * Memformat tab 📝 Database Konten (header, lebar kolom, filter, banding, validasi 1000 baris).
 */
function formatDatabaseSheet(sheet, clearExisting) {
  if (clearExisting) {
    sheet.clear();
  }
  sheet.setTabColor(APP_CONSTANTS.COLORS.ACCENT);

  const totalCols = APP_CONSTANTS.DB_HEADERS.length; // 11 kolom (B..L)
  const startCol = APP_CONSTANTS.DB_COLS.ID;         // 2 (B)
  const endColLetter = columnToLetter(startCol + totalCols - 1); // L

  // Title Banner
  sheet.getRange(`B2:${endColLetter}2`).merge()
    .setValue("📝 MASTER DATABASE KONTEN — VALMART.ID & MARTFORUM.ID")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.HEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 38);

  sheet.getRange(`B3:${endColLetter}3`).merge()
    .setValue("💡 Tim cukup isi 'Link Draft Preview', ubah 'Status', dan tambah 'Catatan Review'. Kolom lainnya otomatis dari Sync GitHub.")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontColor(APP_CONSTANTS.COLORS.MUTED_TEXT)
    .setBackground(APP_CONSTANTS.COLORS.MUTED_BG)
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 22);

  // Header tabel baris 4
  sheet.getRange(4, startCol, 1, totalCols).setValues([APP_CONSTANTS.DB_HEADERS])
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.SUBHEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.SUBHEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(4, 34);

  // Lebar kolom
  const widths = [
    [1, 20],   // A Spacer
    [2, 130],  // B ID (disembunyikan)
    [3, 115],  // C Tanggal Tayang
    [4, 85],   // D Hari
    [5, 105],  // E Brand
    [6, 175],  // F PIC & Role
    [7, 290],  // G Topik & Hook
    [8, 135],  // H Link Naskah
    [9, 150],  // I Link Draft
    [10, 135], // J Status
    [11, 230], // K Catatan
    [12, 135]  // L Update Terakhir
  ];
  widths.forEach(([col, w]) => sheet.setColumnWidth(col, w));

  // Sembunyikan kolom ID agar tampilan bersih untuk tim
  sheet.hideColumns(APP_CONSTANTS.DB_COLS.ID);
  sheet.setFrozenRows(4);

  // Banded rows & Auto Filter
  const existingBandings = sheet.getBandings();
  existingBandings.forEach(b => b.remove());
  const tableRange = sheet.getRange(4, startCol, APP_CONSTANTS.MAX_DB_ROWS + 1, totalCols);
  tableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  tableRange.createFilter();

  applyDatabaseValidationsAndRules(sheet);
}

/**
 * Menerapkan data validation (1000 baris) & conditional formatting di Database Konten.
 */
function applyDatabaseValidationsAndRules(sheet) {
  const startRow = APP_CONSTANTS.DB_START_ROW;
  const maxRows = APP_CONSTANTS.MAX_DB_ROWS;
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONSTANTS.SHEETS.CONFIG);

  // Validasi Tanggal
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.DATE, maxRows, 1)
    .setDataValidation(dateRule)
    .setNumberFormat("yyyy-mm-dd");

  // Format timestamp Update Terakhir
  sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.UPDATED_AT, maxRows, 1)
    .setNumberFormat("yyyy-mm-dd hh:mm");

  // Dropdown dari tab ⚙️ Config (supaya dinamis tanpa edit kode)
  if (configSheet) {
    const brandRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(configSheet.getRange("I7:I20"), true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.BRAND, maxRows, 1).setDataValidation(brandRule);

    const picRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(configSheet.getRange("E7:E20"), true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.PIC, maxRows, 1).setDataValidation(picRule);

    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(configSheet.getRange("J7:J20"), true)
      .setAllowInvalid(false)
      .build();
    sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.STATUS, maxRows, 1).setDataValidation(statusRule);
  }

  // Conditional Formatting: Status + Baris TELAT
  const rules = [];
  const statusColRange = sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.STATUS, maxRows, 1);
  Object.keys(APP_CONSTANTS.COLORS.STATUS).forEach(statusText => {
    const style = APP_CONSTANTS.COLORS.STATUS[statusText];
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(statusText)
        .setBackground(style.bg)
        .setFontColor(style.text)
        .setBold(true)
        .setRanges([statusColRange])
        .build()
    );
  });

  // Highlight merah untuk baris TELAT (Tanggal < TODAY() & Tanggal tidak kosong & Status <> Published)
  const isIndo = isIndonesianLocale();
  const rowHighlightRange = sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.DATE, maxRows, 10); // C..L
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(localizeFormula(`=AND($C${startRow}<>"", $C${startRow}<TODAY(), $J${startRow}<>"🟣 Published")`, isIndo))
      .setBackground(APP_CONSTANTS.COLORS.LATE_BG)
      .setFontColor(APP_CONSTANTS.COLORS.LATE_TEXT)
      .setRanges([rowHighlightRange])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}
