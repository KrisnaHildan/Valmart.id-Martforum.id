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
function setupInitial(options) {
  const opts = options || {};
  const lock = LockService.getScriptLock();
  if (!opts.skipLock) {
    if (!lock.tryLock(15000)) {
      showAlertSafe("Sedang Diproses", "Setup sedang dijalankan oleh pengguna lain. Coba lagi sebentar.");
      return;
    }
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

    renderCalendar({ skipLock: true });
    renderDashboard();
    ss.setActiveSheet(ss.getSheetByName(APP_CONSTANTS.SHEETS.DASHBOARD));

    showToastSafe("Setup awal selesai. Data yang ada tetap aman.", "✅ Setup Berhasil", 5);
  } finally {
    if (!opts.skipLock) {
      lock.releaseLock();
    }
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
      applyDatabaseMonthYearFilter(dbSheet);
    }
    renderCalendar({ skipLock: true });
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
    renderCalendar({ skipLock: true });

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
 * Memformat tab 📝 Database Konten (bar filter Bulan/Tahun, header, auto-migrasi, kapasitas dinamis).
 */
function formatDatabaseSheet(sheet, clearExisting) {
  const ss = sheet.getParent();
  const isIndo = isIndonesianLocale(ss);
  const f = (formulaStr) => localizeFormula(formulaStr, isIndo);
  const now = new Date();

  let prevMonthRaw = "";
  let prevYearRaw = "";
  let migratedOldLayout = false;

  if (clearExisting) {
    sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
    sheet.clearConditionalFormatRules();
    sheet.clear();
  } else {
    // Deteksi layout lama (header masih di baris 4) agar data existing di baris 5+ tidak tertimpa
    const b4Val = String(sheet.getRange("B4").getValue() || "").trim();
    const c4Val = String(sheet.getRange("C4").getValue() || "").trim();
    const c5Val = String(sheet.getRange("C5").getValue() || "").trim();
    if (b4Val === "ID" && c4Val === "Tanggal Tayang" && c5Val !== "Tanggal Tayang") {
      const oldFilter = sheet.getFilter();
      if (oldFilter) oldFilter.remove();
      sheet.getBandings().forEach(b => b.remove());
      sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
      sheet.insertRowBefore(4);
      sheet.getRange(4, 1, 1, 12).clear();
      migratedOldLayout = true;
    } else if (c5Val === "Tanggal Tayang") {
      prevMonthRaw = sheet.getRange("D4").getValue();
      prevYearRaw = sheet.getRange("F4").getValue();
    }
  }

  const activeMonth = (clearExisting || migratedOldLayout || prevMonthRaw === "")
    ? (now.getMonth() + 1)
    : parseMonthFilterValue(prevMonthRaw, true, now.getMonth() + 1);
  const activeYear = (clearExisting || migratedOldLayout || prevYearRaw === "")
    ? now.getFullYear()
    : parseYearFilterValue(prevYearRaw, true, now.getFullYear());

  sheet.setTabColor(APP_CONSTANTS.COLORS.ACCENT);

  const totalCols = APP_CONSTANTS.DB_HEADERS.length; // 11 kolom (B..L)
  const startCol = APP_CONSTANTS.DB_COLS.ID;         // 2 (B)
  const endColLetter = columnToLetter(startCol + totalCols - 1); // L

  // Title Banner (C2:L2 agar tidak terpotong kolom B yang disembunyikan)
  sheet.getRange(`B2:${endColLetter}2`).breakApart();
  sheet.getRange(`C2:${endColLetter}2`).merge()
    .setValue("📝 MASTER DATABASE KONTEN - VALMART.ID & MARTFORUM.ID")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.HEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 38);

  sheet.getRange(`B3:${endColLetter}3`).breakApart();
  sheet.getRange(`C3:${endColLetter}3`).merge()
    .setValue("💡 Isi 'Link Draft Preview', 'Status', & 'Catatan Review'. Gunakan Filter Bulan/Tahun di baris 4 untuk fokus periode atau pilih 'Semua'.")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontColor(APP_CONSTANTS.COLORS.MUTED_TEXT)
    .setBackground(APP_CONSTANTS.COLORS.MUTED_BG)
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 22);

  // Baris 4: Filter Bulan & Tahun + Penghitung Konten Tampil
  sheet.getRange(`B4:${endColLetter}4`).breakApart();
  sheet.setRowHeight(4, 28);
  sheet.getRange("C4").setValue("🗓️ Filter Bulan:")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setFontWeight("bold")
    .setFontColor("#0F172A").setBackground("#FFFFFF")
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sheet.getRange("D4").setValue(activeMonth)
    .setNumberFormat(typeof activeMonth === "number" ? "0" : "@")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(10).setFontWeight("bold")
    .setFontColor("#1E1B4B").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground(APP_CONSTANTS.COLORS.ACCENT_SOFT);
  sheet.getRange("E4").setValue("Tahun:")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setFontWeight("bold")
    .setFontColor("#0F172A").setBackground("#FFFFFF")
    .setHorizontalAlignment("right").setVerticalAlignment("middle");
  sheet.getRange("F4").setValue(activeYear)
    .setNumberFormat(typeof activeYear === "number" ? "0" : "@")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(10).setFontWeight("bold")
    .setFontColor("#1E1B4B").setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground(APP_CONSTANTS.COLORS.ACCENT_SOFT);

  sheet.getRange("D4").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(["Semua", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"], true)
      .setAllowInvalid(false)
      .build()
  );
  sheet.getRange("F4").setDataValidation(
    SpreadsheetApp.newDataValidation()
      .requireValueInList(getSupportedYearsList(true), true)
      .setAllowInvalid(true)
      .build()
  );

  const dbStart = APP_CONSTANTS.DB_START_ROW;
  sheet.getRange(`G4:${endColLetter}4`).merge()
    .setFormula(f(`="Menampilkan " & SUBTOTAL(103, $G$${dbStart}:$G) & " dari " & COUNTA($G$${dbStart}:$G) & " total konten"`))
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9).setFontWeight("bold")
    .setFontColor("#4338CA").setBackground("#F8FAFC")
    .setHorizontalAlignment("right").setVerticalAlignment("middle");

  // Header tabel baris 5
  const headerRow = APP_CONSTANTS.DB_HEADER_ROW;
  sheet.getRange(headerRow, startCol, 1, totalCols).setValues([APP_CONSTANTS.DB_HEADERS])
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.SUBHEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.SUBHEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(headerRow, 34);

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
  sheet.setFrozenRows(headerRow);

  const capacityRows = getDatabaseCapacityRows(sheet, 0);
  ensureDatabasePhysicalRows(sheet, dbStart + capacityRows - 1);

  // Banded rows & Auto Filter
  sheet.getBandings().forEach(b => b.remove());
  const tableRange = sheet.getRange(headerRow, startCol, capacityRows + 1, totalCols);
  tableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  const existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  tableRange.createFilter();

  applyDatabaseValidationsAndRules(sheet, capacityRows);
  applyDatabaseProtection(sheet, capacityRows);
  applyDatabaseMonthYearFilter(sheet);
}

/**
 * Memastikan jumlah baris fisik sheet mencukupi target row terakhir.
 */
function ensureDatabasePhysicalRows(sheet, targetLastRow) {
  const maxRows = sheet.getMaxRows();
  if (maxRows < targetLastRow) {
    sheet.insertRowsAfter(maxRows, targetLastRow - maxRows);
  }
}

/**
 * Memperluas kapasitas format, filter, dan validasi secara otomatis jika data mendekati batas bawah.
 */
function ensureDatabaseCapacityIfNeeded(sheet, minNeededDataRows) {
  const headerRow = APP_CONSTANTS.DB_HEADER_ROW;
  const startRow = APP_CONSTANTS.DB_START_ROW;
  const startCol = APP_CONSTANTS.DB_COLS.ID;
  const totalCols = APP_CONSTANTS.DB_HEADERS.length;
  const neededEndRow = startRow + Math.max(minNeededDataRows || 0, 0) - 1;
  const existingFilter = sheet.getFilter();
  const filterEndRow = existingFilter ? existingFilter.getRange().getLastRow() : 0;

  if (existingFilter && filterEndRow >= neededEndRow && sheet.getMaxRows() >= neededEndRow) {
    return;
  }

  const capacityRows = getDatabaseCapacityRows(sheet, minNeededDataRows);
  ensureDatabasePhysicalRows(sheet, startRow + capacityRows - 1);

  sheet.getBandings().forEach(b => b.remove());
  const tableRange = sheet.getRange(headerRow, startCol, capacityRows + 1, totalCols);
  tableRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);

  if (existingFilter) existingFilter.remove();
  tableRange.createFilter();

  applyDatabaseValidationsAndRules(sheet, capacityRows);
  applyDatabaseProtection(sheet, capacityRows);
}

/**
 * Memasang proteksi Warning-Only untuk kolom otomatis (D4, F4, I..K tetap bebas diedit tanpa peringatan).
 */
function applyDatabaseProtection(sheet, capacityRows) {
  const startRow = APP_CONSTANTS.DB_START_ROW;
  const rows = capacityRows || getDatabaseCapacityRows(sheet, 0);
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  const protection = sheet.protect().setDescription("Proteksi Kolom Otomatis Database (Warning Only)");
  protection.setUnprotectedRanges([
    sheet.getRange("D4"),
    sheet.getRange("F4"),
    sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.DRAFT_LINK, rows, 3) // I..K (Link Draft, Status, Catatan)
  ]);
  protection.setWarningOnly(true);
}

/**
 * Menerapkan filter AutoFilter kolom Tanggal berdasarkan pilihan Bulan (D4) & Tahun (F4).
 */
function applyDatabaseMonthYearFilter(sheet) {
  if (!sheet) return;
  const ss = sheet.getParent();
  const isIndo = isIndonesianLocale(ss);
  const f = (formulaStr) => localizeFormula(formulaStr, isIndo);

  const rawMonth = sheet.getRange("D4").getValue();
  const rawYear = sheet.getRange("F4").getValue();
  const monthChoice = (rawMonth === "" || rawMonth === null) ? "Semua" : parseMonthFilterValue(rawMonth, true, "Semua");
  const yearChoice = (rawYear === "" || rawYear === null) ? "Semua" : parseYearFilterValue(rawYear, true, "Semua");

  // Rapikan format angka/teks pada sel D4 dan F4 sesuai pilihannya
  sheet.getRange("D4").setNumberFormat(typeof monthChoice === "number" ? "0" : "@");
  sheet.getRange("F4").setNumberFormat(typeof yearChoice === "number" ? "0" : "@");

  let filter = sheet.getFilter();
  if (!filter) {
    const capacityRows = getDatabaseCapacityRows(sheet, 0);
    const tableRange = sheet.getRange(APP_CONSTANTS.DB_HEADER_ROW, APP_CONSTANTS.DB_COLS.ID, capacityRows + 1, APP_CONSTANTS.DB_HEADERS.length);
    filter = tableRange.createFilter();
  }

  const dateCol = APP_CONSTANTS.DB_COLS.DATE;
  const startRow = APP_CONSTANTS.DB_START_ROW;

  if (monthChoice === "Semua" && yearChoice === "Semua") {
    filter.removeColumnFilterCriteria(dateCol);
    return;
  }

  const conds = [`ISNUMBER($C${startRow})`];
  if (monthChoice !== "Semua") conds.push(`MONTH($C${startRow})=${monthChoice}`);
  if (yearChoice !== "Semua") conds.push(`YEAR($C${startRow})=${yearChoice}`);

  const formula = f(`=AND(${conds.join(", ")})`);
  const criteria = SpreadsheetApp.newFilterCriteria()
    .whenFormulaSatisfied(formula)
    .build();
  filter.setColumnFilterCriteria(dateCol, criteria);
}

/**
 * Menerapkan data validation & conditional formatting di Database Konten sesuai kapasitas dinamis.
 */
function applyDatabaseValidationsAndRules(sheet, capacityRows) {
  const startRow = APP_CONSTANTS.DB_START_ROW;
  const maxRows = capacityRows || getDatabaseCapacityRows(sheet, 0);
  const configSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(APP_CONSTANTS.SHEETS.CONFIG);

  // Validasi Tanggal ketat agar formula Dashboard selalu sinkron dengan Calendar
  const dateRule = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(false)
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
