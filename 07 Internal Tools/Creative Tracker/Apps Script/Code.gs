/**
 * =========================================================================
 * VALMART.ID & MARTFORUM.ID — CREATIVE CONTENT ENGINE & TRACKER
 * =========================================================================
 * 3 Tab Terpadu:
 * 1. 📊 Dashboard: Ringkasan KPI, Pacing Talent (Sparkline), & Antrean Review
 * 2. 📅 Kalender Visual: Grid Kalender Bulanan (Senin–Minggu) + Card Status
 * 3. 📝 Database Konten: Master Data 10 Kolom dengan Link Naskah GitHub
 * =========================================================================
 */

const CONFIG = {
  SHEET_DASHBOARD: "📊 Dashboard",
  SHEET_CALENDAR: "📅 Kalender Visual",
  SHEET_DATABASE: "📝 Database Konten",
  REPO_BASE_URL: "https://github.com/KrisnaHildan/Valmart.id-Martforum.id/blob/main/",
  FONT_FAMILY: "Google Sans",
  COLORS: {
    HEADER_BG: "#0F172A",       // Deep Slate / Navy
    HEADER_TEXT: "#FFFFFF",
    DAY_HEADER_BG: "#1E293B",   // Slate 800
    DAY_HEADER_TEXT: "#F8FAFC",
    GRID_BORDER: "#CBD5E1",     // Slate 300
    TODAY_ACCENT: "#6366F1",    // Indigo 500
    TODAY_BG: "#EEF2FF",        // Indigo 50
    VALMART: "#4F46E5",         // Indigo
    MARTFORUM: "#0284C7",       // Sky Blue
    STATUS: {
      DRAFT: { bg: "#F1F5F9", text: "#475569" },
      PROGRESS: { bg: "#FEF3C7", text: "#92400E" },
      REVIEW: { bg: "#FFEDD5", text: "#9A3412" },
      REVISION: { bg: "#FEE2E2", text: "#991B1B" },
      APPROVED: { bg: "#DCFCE7", text: "#166534" },
      PUBLISHED: { bg: "#EDE9FE", text: "#5B21B6" }
    }
  },
  PIC_LIST: [
    "Helmy (Video)",
    "Audrey (VO)",
    "Dika / Jonathan (Editor & Desain)"
  ],
  BRAND_LIST: [
    "Valmart",
    "Martforum"
  ],
  STATUS_LIST: [
    "⚪ Draft",
    "🟡 In Progress",
    "🟠 In Review",
    "🔴 Revision",
    "🟢 Approved",
    "🟣 Published"
  ]
};

/**
 * Menu Kustom di Toolbar Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚡ Valmart & Martforum")
    .addItem("📊 Format Ulang Dashboard", "setupDashboard")
    .addItem("📅 Render / Refresh Kalender Visual", "renderCalendar")
    .addSeparator()
    .addItem("📥 Sinkronisasi Naskah dari GitHub (Nov 2026)", "syncNovemberSchedule")
    .addItem("🛠️ Setup Ulang Semua Tab (3 Tab Lengkap)", "setupTracker")
    .addToUi();
}

/**
 * Setup Lengkap 3 Tab: Dashboard, Kalender, Database
 */
function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Tab Database
  let dbSheet = ss.getSheetByName(CONFIG.SHEET_DATABASE);
  if (!dbSheet) {
    dbSheet = ss.insertSheet(CONFIG.SHEET_DATABASE);
  }
  formatDatabaseSheet(dbSheet);

  // 2. Setup Tab Kalender Visual
  let calSheet = ss.getSheetByName(CONFIG.SHEET_CALENDAR);
  if (!calSheet) {
    calSheet = ss.insertSheet(CONFIG.SHEET_CALENDAR);
  }
  renderCalendar();

  // 3. Setup Tab Dashboard (Diletakkan di Urutan Pertama)
  let dashSheet = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
  if (!dashSheet) {
    dashSheet = ss.insertSheet(CONFIG.SHEET_DASHBOARD, 0);
  }
  setupDashboard();

  SpreadsheetApp.getUi().alert(
    "✅ 3 Tab Berhasil Dibuat!",
    "Sistem Tracker & Dashboard berhasil disetup:\n" +
    "1. 📊 Dashboard (KPI & Antrean Review)\n" +
    "2. 📅 Kalender Visual (Grid Bulanan)\n" +
    "3. 📝 Database Konten (Master Data)\n\n" +
    "Klik menu [⚡ Valmart & Martforum] -> [📥 Sinkronisasi Naskah dari GitHub] untuk mengisi data.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * 1. PEMBUATAN TAB DASHBOARD (LEAN & REALTIME)
 */
function setupDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG.SHEET_DASHBOARD);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_DASHBOARD, 0);
  }

  const locale = (ss.getSpreadsheetLocale() || "").toLowerCase();
  const isIndo = locale.startsWith("in") || locale.startsWith("id");
  const setSafeFormula = (range, formulaEn, formulaId) => {
    if (isIndo && formulaId) {
      range.setFormulaLocal(formulaId);
    } else {
      range.setFormula(formulaEn);
    }
  };

  sheet.clear();
  sheet.setTabColor("#10B981"); // Emerald Tab Color

  // Column widths
  sheet.setColumnWidth(1, 20);  // A Spacing
  sheet.setColumnWidth(2, 130); // B
  sheet.setColumnWidth(3, 100); // C
  sheet.setColumnWidth(4, 110); // D
  sheet.setColumnWidth(5, 140); // E
  sheet.setColumnWidth(6, 25);  // F Spacer
  sheet.setColumnWidth(7, 100); // G
  sheet.setColumnWidth(8, 100); // H
  sheet.setColumnWidth(9, 130); // I
  sheet.setColumnWidth(10, 240);// J
  sheet.setColumnWidth(11, 120);// K

  // Header Banner
  sheet.getRange("B2:K2").merge()
    .setValue("📊 CREATIVE DASHBOARD — VALMART.ID & MARTFORUM.ID")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 38);

  sheet.getRange("B3:K3").merge()
    .setValue("💡 Dashboard Realtime: Metrik otomatis terupdate saat status diubah di tab '📝 Database Konten'.")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(9)
    .setFontColor("#64748B")
    .setBackground("#F8FAFC")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 22);

  // -------------------------------------------------------------------------
  // 4 KPI SCORECARDS (Row 5 - 7)
  // -------------------------------------------------------------------------
  sheet.setRowHeight(4, 10); // Spacer
  sheet.setRowHeight(5, 20); // Label Card
  sheet.setRowHeight(6, 40); // Big Number
  sheet.setRowHeight(7, 20); // Subtext / Badge

  // Card 1: Total Konten (B5:C7)
  createScoreCard(sheet, "B5:C5", "B6:C6", "B7:C7", 
    "🎯 TOTAL TERJADWAL", 
    "=COUNTIF('📝 Database Konten'!B5:B100, \">0\")", 
    "Bulan Berjalan", 
    "#F8FAFC", "#0F172A", "#CBD5E1", null, isIndo);

  // Card 2: Siap Tayang / Published (D5:E7)
  createScoreCard(sheet, "D5:E5", "D6:E6", "D7:E7", 
    "🚀 APPROVED / PUBLISHED", 
    "=COUNTIF('📝 Database Konten'!J5:J100, \"*Approved*\") + COUNTIF('📝 Database Konten'!J5:J100, \"*Published*\")", 
    "=IF(B6>0, TEXT((D6/B6),\"0%\") & \" Selesai\", \"0%\")", 
    "#F0FDF4", "#166534", "#86EFAC",
    "=IF(B6>0; TEXT((D6/B6);\"0%\") & \" Selesai\"; \"0%\")", isIndo);

  // Card 3: Butuh Review (G5:H7)
  createScoreCard(sheet, "G5:H5", "G6:H6", "G7:H7", 
    "⚠️ ANTREAN REVIEW", 
    "=COUNTIF('📝 Database Konten'!J5:J100, \"*In Review*\") + COUNTIF('📝 Database Konten'!J5:J100, \"*Revision*\")", 
    "Perlu Dicek Hari Ini", 
    "#FFF7ED", "#C2410C", "#FDBA74", null, isIndo);

  // Card 4: Dalam Pengerjaan (J5:K7)
  createScoreCard(sheet, "J5:K5", "J6:K6", "J7:K7", 
    "⏳ DALAM PENGERJAAN", 
    "=COUNTIF('📝 Database Konten'!J5:J100, \"*Draft*\") + COUNTIF('📝 Database Konten'!J5:J100, \"*In Progress*\")", 
    "On Track Talent", 
    "#FEFCE8", "#A16207", "#FDE047", null, isIndo);

  sheet.setRowHeight(8, 15); // Spacer

  // -------------------------------------------------------------------------
  // SECTION KIRI (B9:E19): WORKLOAD & PACING TALENT
  // -------------------------------------------------------------------------
  sheet.getRange("B9:E9").merge()
    .setValue("👥 TARGET MINGGUAN & PACING TALENT")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.DAY_HEADER_BG)
    .setFontColor(CONFIG.COLORS.DAY_HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(9, 28);

  const talentHeaders = ["Nama Talent / Role", "Target", "Selesai", "Visual Progress Bar"];
  sheet.getRange("B10:E10").setValues([talentHeaders])
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(9)
    .setFontWeight("bold")
    .setBackground("#F1F5F9")
    .setFontColor("#475569")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(10, 24);

  // Baris Helmy (Row 11)
  sheet.getRange("B11").setValue("🎬 Helmy (Valmart)").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setFontWeight("bold");
  sheet.getRange("C11").setValue("4 Video/Mgg").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setHorizontalAlignment("center");
  sheet.getRange("D11").setFormula("=COUNTIFS('📝 Database Konten'!F5:F100, \"*Helmy*\", '📝 Database Konten'!J5:J100, \"*Approved*\") + COUNTIFS('📝 Database Konten'!F5:F100, \"*Helmy*\", '📝 Database Konten'!J5:J100, \"*Published*\")").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
  setSafeFormula(sheet.getRange("E11"),
    "=SPARKLINE(D11, {\"charttype\", \"bar\"; \"max\", 8; \"color1\", \"#4F46E5\"; \"color2\", \"#E2E8F0\"})",
    "=SPARKLINE(D11; {\"charttype\"\\ \"bar\"; \"max\"\\ 8; \"color1\"\\ \"#4F46E5\"; \"color2\"\\ \"#E2E8F0\"})",
    isIndo);
  sheet.setRowHeight(11, 26);

  // Baris Audrey (Row 12)
  sheet.getRange("B12").setValue("🎙️ Audrey (Martforum)").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setFontWeight("bold");
  sheet.getRange("C12").setValue("3 VO/Mgg").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setHorizontalAlignment("center");
  sheet.getRange("D12").setFormula("=COUNTIFS('📝 Database Konten'!F5:F100, \"*Audrey*\", '📝 Database Konten'!J5:J100, \"*Approved*\") + COUNTIFS('📝 Database Konten'!F5:F100, \"*Audrey*\", '📝 Database Konten'!J5:J100, \"*Published*\")").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
  setSafeFormula(sheet.getRange("E12"),
    "=SPARKLINE(D12, {\"charttype\", \"bar\"; \"max\", 6; \"color1\", \"#0284C7\"; \"color2\", \"#E2E8F0\"})",
    "=SPARKLINE(D12; {\"charttype\"\\ \"bar\"; \"max\"\\ 6; \"color1\"\\ \"#0284C7\"; \"color2\"\\ \"#E2E8F0\"})",
    isIndo);
  sheet.setRowHeight(12, 26);

  // Baris Dika / Jonathan (Row 13)
  sheet.getRange("B13").setValue("🎨 Dika / Jonathan (Desain)").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setFontWeight("bold");
  sheet.getRange("C13").setValue("Support Aset").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setHorizontalAlignment("center");
  sheet.getRange("D13").setFormula("=COUNTIFS('📝 Database Konten'!F5:F100, \"*Dika*\", '📝 Database Konten'!J5:J100, \"*Approved*\") + COUNTIFS('📝 Database Konten'!F5:F100, \"*Dika*\", '📝 Database Konten'!J5:J100, \"*Published*\") + COUNTIFS('📝 Database Konten'!F5:F100, \"*Jonathan*\", '📝 Database Konten'!J5:J100, \"*Approved*\") + COUNTIFS('📝 Database Konten'!F5:F100, \"*Jonathan*\", '📝 Database Konten'!J5:J100, \"*Published*\")").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setFontWeight("bold").setHorizontalAlignment("center");
  setSafeFormula(sheet.getRange("E13"),
    "=SPARKLINE(D13, {\"charttype\", \"bar\"; \"max\", 4; \"color1\", \"#10B981\"; \"color2\", \"#E2E8F0\"})",
    "=SPARKLINE(D13; {\"charttype\"\\ \"bar\"; \"max\"\\ 4; \"color1\"\\ \"#10B981\"; \"color2\"\\ \"#E2E8F0\"})",
    isIndo);
  sheet.setRowHeight(13, 26);

  sheet.getRange("B11:E13").setBorder(true, true, true, true, true, true, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // Keseimbangan Brand (Row 15 - 18)
  sheet.getRange("B15:E15").merge()
    .setValue("⚖️ DISTRIBUSI KONTEN BRAND")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground("#334155")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(15, 26);

  sheet.getRange("B16:C16").merge().setValue("🟣 Valmart.id").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setFontWeight("bold");
  sheet.getRange("D16:E16").merge().setFormula("=COUNTIF('📝 Database Konten'!E5:E100, \"Valmart\") & \" Konten\"").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setHorizontalAlignment("center");
  sheet.setRowHeight(16, 24);

  sheet.getRange("B17:C17").merge().setValue("🔵 Martforum.id").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setFontWeight("bold");
  sheet.getRange("D17:E17").merge().setFormula("=COUNTIF('📝 Database Konten'!E5:E100, \"Martforum\") & \" Konten\"").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setHorizontalAlignment("center");
  sheet.setRowHeight(17, 24);

  setSafeFormula(sheet.getRange("B18:E18").merge(),
    "=SPARKLINE({COUNTIF('📝 Database Konten'!E5:E100, \"Valmart\"), COUNTIF('📝 Database Konten'!E5:E100, \"Martforum\")}, {\"charttype\", \"bar\"; \"color1\", \"#4F46E5\"; \"color2\", \"#0284C7\"})",
    "=SPARKLINE({COUNTIF('📝 Database Konten'!E5:E100; \"Valmart\")\\ COUNTIF('📝 Database Konten'!E5:E100; \"Martforum\")}; {\"charttype\"\\ \"bar\"; \"color1\"\\ \"#4F46E5\"; \"color2\"\\ \"#0284C7\"})",
    isIndo);
  sheet.setRowHeight(18, 20);
  sheet.getRange("B15:E18").setBorder(true, true, true, true, true, true, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // -------------------------------------------------------------------------
  // SECTION KANAN (G9:K18): ANTREAN REVIEW / ACTION NEEDED
  // -------------------------------------------------------------------------
  sheet.getRange("G9:K9").merge()
    .setValue("🚨 ANTREAN REVIEW (KLIK LINK DRAFT UNTUK APPROVAL)")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.DAY_HEADER_BG)
    .setFontColor(CONFIG.COLORS.DAY_HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  const reviewHeaders = ["Tanggal", "Brand", "PIC", "Topik Konten", "Status"];
  sheet.getRange("G10:K10").setValues([reviewHeaders])
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(9)
    .setFontWeight("bold")
    .setBackground("#F1F5F9")
    .setFontColor("#475569")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, null, null, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // Filter dinamis otomatis menampilkan baris yang "In Review" atau "Revision"
  setSafeFormula(sheet.getRange("G11"),
    "=IFERROR(FILTER({'📝 Database Konten'!C5:C100, '📝 Database Konten'!E5:E100, '📝 Database Konten'!F5:F100, '📝 Database Konten'!G5:G100, '📝 Database Konten'!J5:J100}, ('📝 Database Konten'!J5:J100 = \"🟠 In Review\") + ('📝 Database Konten'!J5:J100 = \"🔴 Revision\")), {\"-\", \"-\", \"-\", \"✅ Tidak ada antrean review saat ini! Semuanya aman.\", \"-\"})",
    "=IFERROR(FILTER({'📝 Database Konten'!C5:C100\\ '📝 Database Konten'!E5:E100\\ '📝 Database Konten'!F5:F100\\ '📝 Database Konten'!G5:G100\\ '📝 Database Konten'!J5:J100}; ('📝 Database Konten'!J5:J100 = \"🟠 In Review\") + ('📝 Database Konten'!J5:J100 = \"🔴 Revision\")); {\"-\"\\ \"-\"\\ \"-\"\\ \"✅ Tidak ada antrean review saat ini! Semuanya aman.\"\\ \"-\"})",
    isIndo);

  sheet.getRange("G11:K18").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(9).setVerticalAlignment("middle");
  sheet.getRange("G11:K18").setBorder(true, true, true, true, true, true, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

function createScoreCard(sheet, labelRange, numRange, tagRange, label, formula, tagFormula, bg, textColor, borderColor, tagFormulaId, isIndo) {
  sheet.getRange(labelRange).merge()
    .setValue(label)
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(8.5)
    .setFontWeight("bold")
    .setFontColor(textColor)
    .setBackground(bg)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("bottom");

  sheet.getRange(numRange).merge()
    .setFormula(formula)
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(20)
    .setFontWeight("bold")
    .setFontColor(textColor)
    .setBackground(bg)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  const tagCell = sheet.getRange(tagRange).merge();
  if (tagFormula && tagFormula.startsWith("=")) {
    if (isIndo && tagFormulaId) {
      tagCell.setFormulaLocal(tagFormulaId);
    } else {
      tagCell.setFormula(tagFormula);
    }
  } else {
    tagCell.setValue(tagFormula || "");
  }

  tagCell
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(8)
    .setFontColor("#64748B")
    .setBackground(bg)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("top");

  // Border outline kartu
  const startCol = sheet.getRange(labelRange).getColumn();
  const startRow = sheet.getRange(labelRange).getRow();
  sheet.getRange(startRow, startCol, 3, 2).setBorder(true, true, true, true, null, null, borderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

/**
 * 2. FORMAT DATABASE KONTEN
 */
function formatDatabaseSheet(sheet) {
  sheet.clear();
  sheet.setTabColor("#6366F1");

  // Title Banner
  sheet.getRange("B2:K2").merge()
    .setValue("📝 MASTER DATABASE KONTEN — VALMART.ID & MARTFORUM.ID")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.HEADER_BG)
    .setFontColor(CONFIG.COLORS.HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 40);

  sheet.getRange("B3:K3").merge()
    .setValue("💡 Info: Kolom 'Link Naskah GitHub' langsung mengarah ke naskah detail per detik di repository. Editor cukup drop link draft di kolom 'Link Draft Preview'.")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(9)
    .setFontColor("#64748B")
    .setBackground("#F8FAFC")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 24);

  const headers = [
    "No",
    "Tanggal Tayang",
    "Hari",
    "Brand",
    "PIC & Role",
    "Topik & Hook Konten",
    "📖 Link Naskah (GitHub)",
    "🎬 Link Draft Preview",
    "Status",
    "Catatan Review / Feedback"
  ];

  sheet.getRange("B4:K4").setValues([headers])
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.DAY_HEADER_BG)
    .setFontColor(CONFIG.COLORS.DAY_HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(4, 34);

  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 45);
  sheet.setColumnWidth(3, 115);
  sheet.setColumnWidth(4, 85);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 300);
  sheet.setColumnWidth(8, 200);
  sheet.setColumnWidth(9, 190);
  sheet.setColumnWidth(10, 130);
  sheet.setColumnWidth(11, 240);

  const maxRows = 100;

  const brandRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.BRAND_LIST, true).setAllowInvalid(false).build();
  sheet.getRange(5, 5, maxRows, 1).setDataValidation(brandRule);

  const picRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.PIC_LIST, true).setAllowInvalid(false).build();
  sheet.getRange(5, 6, maxRows, 1).setDataValidation(picRule);

  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.STATUS_LIST, true).setAllowInvalid(false).build();
  sheet.getRange(5, 10, maxRows, 1).setDataValidation(statusRule);

  sheet.getRange(5, 3, maxRows, 1).setNumberFormat("yyyy-mm-dd");

  const rules = [];
  const statusRange = sheet.getRange("J5:J100");
  const colorMap = [
    { text: "⚪ Draft", bg: CONFIG.COLORS.STATUS.DRAFT.bg, color: CONFIG.COLORS.STATUS.DRAFT.text },
    { text: "🟡 In Progress", bg: CONFIG.COLORS.STATUS.PROGRESS.bg, color: CONFIG.COLORS.STATUS.PROGRESS.text },
    { text: "🟠 In Review", bg: CONFIG.COLORS.STATUS.REVIEW.bg, color: CONFIG.COLORS.STATUS.REVIEW.text },
    { text: "🔴 Revision", bg: CONFIG.COLORS.STATUS.REVISION.bg, color: CONFIG.COLORS.STATUS.REVISION.text },
    { text: "🟢 Approved", bg: CONFIG.COLORS.STATUS.APPROVED.bg, color: CONFIG.COLORS.STATUS.APPROVED.text },
    { text: "🟣 Published", bg: CONFIG.COLORS.STATUS.PUBLISHED.bg, color: CONFIG.COLORS.STATUS.PUBLISHED.text }
  ];

  colorMap.forEach(item => {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(item.text)
        .setBackground(item.bg)
        .setFontColor(item.color)
        .setBold(true)
        .setRanges([statusRange])
        .build()
    );
  });
  sheet.setConditionalFormatRules(rules);
  sheet.setFrozenRows(4);
}

/**
 * 3. RENDER KALENDER VISUAL
 */
function renderCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calSheet = ss.getSheetByName(CONFIG.SHEET_CALENDAR);
  if (!calSheet) {
    calSheet = ss.insertSheet(CONFIG.SHEET_CALENDAR);
  }

  const dbSheet = ss.getSheetByName(CONFIG.SHEET_DATABASE);
  if (!dbSheet) return;

  let targetMonth = calSheet.getRange("C2").getValue();
  let targetYear = calSheet.getRange("E2").getValue();

  const now = new Date();
  if (!targetMonth || isNaN(targetMonth) || targetMonth < 1 || targetMonth > 12) {
    targetMonth = now.getMonth() + 1;
  }
  if (!targetYear || isNaN(targetYear) || targetYear < 2020) {
    targetYear = now.getFullYear();
  }

  // Clear area kalender lama dengan aman
  const maxRows = calSheet.getMaxRows();
  if (maxRows >= 5) {
    calSheet.getRange(5, 1, maxRows - 4, calSheet.getMaxColumns()).clear();
  }
  calSheet.setTabColor("#4F46E5");

  setupCalendarControls(calSheet, targetMonth, targetYear);
  const contentMap = loadContentFromDatabase(dbSheet, targetMonth, targetYear);
  drawCalendarGrid(calSheet, targetMonth, targetYear, contentMap);
}

function setupCalendarControls(sheet, month, year) {
  sheet.setColumnWidth(1, 20);

  for (let c = 2; c <= 8; c++) {
    sheet.setColumnWidth(c, 180);
  }

  sheet.getRange("B2").setValue("🗓️ BULAN:").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setFontColor("#1E293B");
  sheet.getRange("C2").setValue(month).setNumberFormat("0").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#EEF2FF");
  sheet.getRange("D2").setValue("TAHUN:").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setFontColor("#1E293B");
  sheet.getRange("E2").setValue(year).setNumberFormat("0").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#EEF2FF");

  const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  sheet.getRange("F2:H2").merge()
    .setValue("📅 " + monthNames[month - 1] + " " + year + " (Ganti C2/E2 lalu klik Refresh)")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setFontColor("#4338CA")
    .setHorizontalAlignment("right");
  sheet.setRowHeight(2, 32);

  const days = ["SENIN", "SELASA", "RABU", "KAMIS", "JUMAT", "SABTU", "MINGGU"];
  sheet.getRange("B4:H4").setValues([days])
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.DAY_HEADER_BG)
    .setFontColor(CONFIG.COLORS.DAY_HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(4, 30);
}

function loadContentFromDatabase(dbSheet, targetMonth, targetYear) {
  const lastRow = dbSheet.getLastRow();
  const map = {};
  if (lastRow < 5) return map;

  const data = dbSheet.getRange(5, 2, lastRow - 4, 10).getValues();

  for (let i = 0; i < data.length; i++) {
    const rawDate = data[i][1];
    if (!rawDate) continue;

    const d = new Date(rawDate);
    if (isNaN(d.getTime())) continue;

    const m = d.getMonth() + 1;
    const y = d.getFullYear();
    const dateNum = d.getDate();

    if (m === targetMonth && y === targetYear) {
      if (!map[dateNum]) {
        map[dateNum] = [];
      }
      map[dateNum].push({
        brand: data[i][3],
        pic: data[i][4],
        topik: data[i][5],
        scriptUrl: data[i][6],
        draftUrl: data[i][7],
        status: data[i][8],
        feedback: data[i][9]
      });
    }
  }
  return map;
}

function drawCalendarGrid(sheet, month, year, contentMap) {
  const firstDay = new Date(year, month - 1, 1);
  const lastDate = new Date(year, month, 0).getDate();
  const today = new Date();
  const isCurrentMonth = (today.getFullYear() === year && (today.getMonth() + 1) === month);
  const todayDate = today.getDate();

  let startCol = firstDay.getDay(); 
  startCol = (startCol === 0) ? 6 : startCol - 1;

  let currentDay = 1;
  let currentRow = 5;

  while (currentDay <= lastDate) {
    const dateHeaderRow = currentRow;
    const contentRow = currentRow + 1;

    sheet.setRowHeight(dateHeaderRow, 24);
    sheet.setRowHeight(contentRow, 115);

    for (let c = 0; c < 7; c++) {
      const colIndex = c + 2;
      const cellDateHeader = sheet.getRange(dateHeaderRow, colIndex);
      const cellContent = sheet.getRange(contentRow, colIndex);

      if ((currentRow === 5 && c < startCol) || currentDay > lastDate) {
        cellDateHeader.setBackground("#F8FAFC").setBorder(true, true, true, true, null, null, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);
        cellContent.setBackground("#F8FAFC").setBorder(true, true, true, true, null, null, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      } else {
        const isToday = isCurrentMonth && (currentDay === todayDate);

        cellDateHeader
          .setValue(currentDay + (isToday ? "  📍 HARI INI" : ""))
          .setFontFamily(CONFIG.FONT_FAMILY)
          .setFontSize(9)
          .setFontWeight("bold")
          .setHorizontalAlignment("left")
          .setVerticalAlignment("middle")
          .setBackground(isToday ? CONFIG.COLORS.TODAY_ACCENT : "#F1F5F9")
          .setFontColor(isToday ? "#FFFFFF" : "#334155")
          .setBorder(true, true, true, true, null, null, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);

        const items = contentMap[currentDay] || [];
        if (items.length > 0) {
          let cardText = "";
          items.forEach((item, idx) => {
            const brandTag = (item.brand === "Valmart") ? "[VAL]" : "[MART]";
            const statusTag = item.status ? item.status : "⚪ Draft";
            const picShort = (item.pic.includes("Dika") || item.pic.includes("Jonathan")) ? "Dika/Jhona" : item.pic.split(" ")[0];
            cardText += (idx > 0 ? "\n───\n" : "") +
              brandTag + " " + picShort + "\n" +
              "📌 " + item.topik + "\n" +
              statusTag;
          });

          cellContent
            .setValue(cardText)
            .setFontFamily(CONFIG.FONT_FAMILY)
            .setFontSize(8.5)
            .setWrap(true)
            .setVerticalAlignment("top")
            .setHorizontalAlignment("left")
            .setBackground(isToday ? CONFIG.COLORS.TODAY_BG : "#FFFFFF")
            .setBorder(true, true, true, true, null, null, isToday ? CONFIG.COLORS.TODAY_ACCENT : CONFIG.COLORS.GRID_BORDER, isToday ? SpreadsheetApp.BorderStyle.SOLID_MEDIUM : SpreadsheetApp.BorderStyle.SOLID);
        } else {
          cellContent
            .setValue("—")
            .setFontFamily(CONFIG.FONT_FAMILY)
            .setFontSize(8.5)
            .setFontColor("#94A3B8")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setBackground(isToday ? CONFIG.COLORS.TODAY_BG : "#FFFFFF")
            .setBorder(true, true, true, true, null, null, isToday ? CONFIG.COLORS.TODAY_ACCENT : CONFIG.COLORS.GRID_BORDER, isToday ? SpreadsheetApp.BorderStyle.SOLID_MEDIUM : SpreadsheetApp.BorderStyle.SOLID);
        }
        currentDay++;
      }
    }
    currentRow += 2;
  }
}

/**
 * 4. SINKRONISASI JADWAL NOVEMBER 2026 DARI GITHUB
 */
function syncNovemberSchedule() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dbSheet = ss.getSheetByName(CONFIG.SHEET_DATABASE);
  if (!dbSheet) {
    setupTracker();
    dbSheet = ss.getSheetByName(CONFIG.SHEET_DATABASE);
  }

  const baseW3Val = CONFIG.REPO_BASE_URL + "04%20Strategi%20Konten%20dan%20Kreatif/November%202026/Week%203%20(16%20-%2022%20Nov%202026)/BRIEF_VALMART_HELMY.md";
  const baseW3Mart = CONFIG.REPO_BASE_URL + "04%20Strategi%20Konten%20dan%20Kreatif/November%202026/Week%203%20(16%20-%2022%20Nov%202026)/BRIEF_MARTFORUM_JONATHAN.md";
  const baseW4Val = CONFIG.REPO_BASE_URL + "04%20Strategi%20Konten%20dan%20Kreatif/November%202026/Week%204%20(23%20-%2030%20Nov%202026)/BRIEF_VALMART_HELMY.md";
  const baseW4Mart = CONFIG.REPO_BASE_URL + "04%20Strategi%20Konten%20dan%20Kreatif/November%202026/Week%204%20(23%20-%2030%20Nov%202026)/BRIEF_MARTFORUM_JONATHAN.md";

  const schedule = [
    // Week 3
    [1, "2026-11-16", "Senin", "Valmart", "Helmy (Video)", "Mid-Month ACE Bounty Recap (Reaksi Klip 200k)", baseW3Val, "", "🟣 Published", "Tayang Senin W3"],
    [2, "2026-11-16", "Senin", "Martforum", "Audrey (VO)", "Teaser Resmi Brand: Top Up QRIS 3 Detik", baseW3Mart, "", "🟣 Published", "Selesai"],
    [3, "2026-11-17", "Selasa", "Valmart", "Helmy (Video)", "Meme 10 Omen Chaos di Ascent", baseW3Val, "", "🟣 Published", "Tayang Selasa W3"],
    [4, "2026-11-18", "Rabu", "Valmart", "Helmy (Video)", "Sinergi Dua Brand: Sewa Valmart Diskon Martforum", baseW3Val, "", "🟣 Published", "Tayang Rabu W3"],
    [5, "2026-11-18", "Rabu", "Martforum", "Audrey (VO)", "Meme Temen Discord Pas Ada Bundle Baru", baseW3Mart, "", "🟣 Published", "Tayang Rabu W3"],
    [6, "2026-11-19", "Kamis", "Valmart", "Helmy (Video)", "Anti-Scam Awareness: Bedain Akun Resmi vs Fake", baseW3Val, "", "🟣 Published", "Tayang Kamis W3"],
    [7, "2026-11-20", "Jumat", "Valmart", "Helmy (Video)", "Meme Sekte Skin: Kuronami vs Reaver vs Primordium", baseW3Val, "", "🟣 Published", "Tayang Jumat W3"],
    [8, "2026-11-20", "Jumat", "Martforum", "Audrey (VO)", "Weekend Mabar: Anti-Phishing Cukup Riot ID", baseW3Mart, "", "🟢 Approved", "Naskah Baru Audrey"],
    [9, "2026-11-21", "Sabtu", "Valmart", "Dika / Jonathan (Editor & Desain)", "Infografis: Panduan Sewa Goceng 1 Menit Anti-Ribet", baseW3Mart, "", "🟢 Approved", "Desain Carousel Dika"],
    [10, "2026-11-22", "Minggu", "Valmart", "Helmy (Video)", "Reminder Sayembara #ValmartAceBounty (Tersisa 8 Hari)", baseW3Val, "", "🟢 Approved", "Story & Feed Push"],

    // Week 4 (Minggu Ini)
    [11, "2026-11-23", "Senin", "Martforum", "Audrey (VO)", "Derita Solo Queue vs Obatnya (Mabar Skin Wangi)", baseW4Mart, "", "🟠 In Review", "Jadwal Hari Ini (Senin)"],
    [12, "2026-11-24", "Selasa", "Valmart", "Helmy (Video)", "Smart Buyer: Ngetes Skin di Valmart Sebelum Boncos", baseW4Val, "", "🟡 In Progress", "Jadwal Selasa W4"],
    [13, "2026-11-25", "Rabu", "Martforum", "Audrey (VO)", "Meme Streamer Jago vs Realita Whiff Kita", baseW4Mart, "", "⚪ Draft", "Jadwal Rabu W4"],
    [14, "2026-11-25", "Rabu", "Valmart", "Helmy (Video)", "Meme Tanggal Tua Tetep Bisa Sultan (Modal Goceng)", baseW4Val, "", "⚪ Draft", "Jadwal Rabu W4"],
    [15, "2026-11-26", "Kamis", "Valmart", "Helmy (Video)", "Edukasi Anti-HB: Rahasia First Mail & Garansi 100%", baseW4Val, "", "⚪ Draft", "Jadwal Kamis W4"],
    [16, "2026-11-27", "Jumat", "Martforum", "Audrey (VO)", "PAYDAY FLASH SALE: Amankan VP Akhir Bulan QRIS 3 Detik", baseW4Mart, "", "⚪ Draft", "Jadwal Jumat Payday"],
    [17, "2026-11-27", "Jumat", "Valmart", "Helmy (Video)", "Showcase Katalog Akun Sultan Weekend Ready Sewa", baseW4Val, "", "⚪ Draft", "Jadwal Jumat W4"],
    [18, "2026-11-30", "Senin", "Valmart", "Dika / Jonathan (Editor & Desain)", "Poster Pemenang Sayembara ACE Bounty (Rp 200k Cash)", baseW4Mart, "", "⚪ Draft", "Closing November"]
  ];

  dbSheet.getRange(5, 2, schedule.length, 10).setValues(schedule);
  renderCalendar();
  setupDashboard(); // Otomatis refresh dashboard

  SpreadsheetApp.getUi().alert(
    "📥 Sinkronisasi Berhasil!",
    "Seluruh jadwal aktif November 2026 berhasil disinkronkan.\n" +
    "Dashboard dan Kalender Visual otomatis ter-update realtime.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
