/**
 * =========================================================================
 * CONFIG.GS — Konstanta Global & Pengelolaan Tab "⚙️ Config"
 * =========================================================================
 */

const APP_CONSTANTS = {
  APP_NAME: "Valmart & Martforum Creative Content Engine",
  MENU_NAME: "⚡ Valmart & Martforum",
  FONT_FAMILY: "Google Sans",
  TIMEZONE: "Asia/Jakarta",
  MIN_DB_ROWS: 1000,
  DB_BUFFER_ROWS: 250,
  DB_FILTER_ROW: 4,
  DB_HEADER_ROW: 5,
  DB_START_ROW: 6,
  YEAR_MIN: 2020,
  YEAR_MAX: 2040,
  CALENDAR_MAX_ITEMS_PER_DAY: 4,
  SCRIPT_PROP_GITHUB_TOKEN: "GITHUB_TOKEN",

  SHEETS: {
    DASHBOARD: "📊 Dashboard",
    CALENDAR: "📅 Kalender Visual",
    DATABASE: "📝 Database Konten",
    CONFIG: "⚙️ Config"
  },

  // Indeks kolom (1-based) di tab 📝 Database Konten (Kolom A = spacer)
  DB_COLS: {
    ID: 2,          // B
    DATE: 3,        // C
    DAY: 4,         // D
    BRAND: 5,       // E
    PIC: 6,         // F
    TOPIC: 7,       // G
    SCRIPT_LINK: 8, // H
    DRAFT_LINK: 9,  // I
    STATUS: 10,     // J
    NOTES: 11,      // K
    UPDATED_AT: 12  // L
  },

  DB_HEADERS: [
    "ID",
    "Tanggal Tayang",
    "Hari",
    "Brand",
    "PIC & Role",
    "Topik & Hook Konten",
    "📖 Link Naskah (GitHub)",
    "🎬 Link Draft Preview",
    "Status",
    "Catatan Review / Feedback",
    "Update Terakhir"
  ],

  COLORS: {
    HEADER_BG: "#0F172A",
    HEADER_TEXT: "#FFFFFF",
    SUBHEADER_BG: "#1E293B",
    SUBHEADER_TEXT: "#F8FAFC",
    MUTED_BG: "#F8FAFC",
    MUTED_TEXT: "#64748B",
    BORDER: "#CBD5E1",
    ACCENT: "#6366F1",
    ACCENT_SOFT: "#EEF2FF",
    WEEKEND_BG: "#F1F5F9",
    OUT_OF_MONTH_BG: "#F8FAFC",
    OUT_OF_MONTH_TEXT: "#94A3B8",
    LATE_BG: "#FEE2E2",
    LATE_TEXT: "#991B1B",
    LATE_BORDER: "#FCA5A5",
    VALMART: "#4F46E5",
    MARTFORUM: "#0284C7",
    BRAND_PALETTE: ["#4F46E5", "#0284C7", "#0D9488", "#D97706", "#DC2626", "#7C3AED"],
    STATUS: {
      "⚪ Draft": { bg: "#F1F5F9", text: "#475569" },
      "🟡 In Progress": { bg: "#FEF3C7", text: "#92400E" },
      "🟠 In Review": { bg: "#FFEDD5", text: "#9A3412" },
      "🔴 Revision": { bg: "#FEE2E2", text: "#991B1B" },
      "🟢 Approved": { bg: "#DCFCE7", text: "#166534" },
      "🟣 Published": { bg: "#EDE9FE", text: "#5B21B6" }
    }
  },

  DEFAULTS: {
    REPO_OWNER: "KrisnaHildan",
    REPO_NAME: "Valmart-x-Martforum",
    REPO_BRANCH: "main",
    BASE_PATH: "04 Strategi Konten dan Kreatif",
    SCHEDULE_FILE: "Jadwal Konten.md",
    EMAIL_ENABLED: "OFF",
    EMAIL_RECIPIENTS: "",
    PICS: [
      { name: "Helmy (Video)", role: "Video Valmart", target: 4 },
      { name: "Audrey (VO)", role: "VO Martforum (Ad Interim)", target: 3 },
      { name: "Dika (Editor & Desain)", role: "Editor & Desain (Pengganti Jonathan)", target: 2 },
      { name: "Jonathan (Editor & Desain)", role: "Cuti s/d Akhir Des (Digantikan Dika)", target: 0 }
    ],
    BRANDS: ["Valmart", "Martforum"],
    STATUSES: [
      "⚪ Draft",
      "🟡 In Progress",
      "🟠 In Review",
      "🔴 Revision",
      "🟢 Approved",
      "🟣 Published"
    ]
  }
};

/**
 * Mengisi / memformat tab ⚙️ Config.
 * Jika clearExisting = false dan sudah ada isinya, tidak menimpa pengaturan user.
 */
function setupConfigSheet(sheet, clearExisting) {
  if (!clearExisting && sheet.getLastRow() >= 10) {
    return;
  }

  sheet.clear();
  sheet.setTabColor("#475569");

  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 230);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 30);
  sheet.setColumnWidth(5, 220);
  sheet.setColumnWidth(6, 160);
  sheet.setColumnWidth(7, 110);
  sheet.setColumnWidth(8, 30);
  sheet.setColumnWidth(9, 140);
  sheet.setColumnWidth(10, 170);

  // Banner utama
  sheet.getRange("B2:J2").merge()
    .setValue("⚙️ PENGATURAN SISTEM - VALMART & MARTFORUM TRACKER")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.HEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 38);

  sheet.getRange("B3:J3").merge()
    .setValue("💡 Ubah pengaturan repo, target PIC, atau penerima email di sini tanpa perlu menyentuh kode Apps Script.")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontColor(APP_CONSTANTS.COLORS.MUTED_TEXT)
    .setBackground(APP_CONSTANTS.COLORS.MUTED_BG)
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 22);

  // Blok 1: Pengaturan GitHub & Email (B5:C13)
  sheet.getRange("B5:C5").merge()
    .setValue("🔗 REPOSITORY & NOTIFIKASI")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.SUBHEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.SUBHEADER_TEXT)
    .setHorizontalAlignment("center");
  sheet.setRowHeight(5, 28);

  const settingsRows = [
    ["Repo Owner", APP_CONSTANTS.DEFAULTS.REPO_OWNER],
    ["Repo Name", APP_CONSTANTS.DEFAULTS.REPO_NAME],
    ["Branch", APP_CONSTANTS.DEFAULTS.REPO_BRANCH],
    ["Base Path Naskah", APP_CONSTANTS.DEFAULTS.BASE_PATH],
    ["File Sumber Jadwal", APP_CONSTANTS.DEFAULTS.SCHEDULE_FILE],
    ["Ringkasan Harian (ON/OFF)", APP_CONSTANTS.DEFAULTS.EMAIL_ENABLED],
    ["Email Penerima (pisah koma)", APP_CONSTANTS.DEFAULTS.EMAIL_RECIPIENTS]
  ];

  sheet.getRange(6, 2, settingsRows.length, 2).setValues(settingsRows)
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9.5)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(6, 2, settingsRows.length, 1).setFontWeight("bold").setBackground("#F1F5F9");
  sheet.getRange(6, 3, settingsRows.length, 1).setBackground("#FFFFFF");

  const emailToggleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["ON", "OFF"], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange("C11").setDataValidation(emailToggleRule);

  // Blok 2: Daftar PIC & Target Mingguan (E5:G20)
  sheet.getRange("E5:G5").merge()
    .setValue("👥 DAFTAR PIC & TARGET MINGGUAN")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.SUBHEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.SUBHEADER_TEXT)
    .setHorizontalAlignment("center");

  sheet.getRange("E6:G6").setValues([["PIC & Role (Dropdown)", "Keterangan Role", "Target / Mgg"]])
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontWeight("bold")
    .setBackground("#F1F5F9")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  const picValues = APP_CONSTANTS.DEFAULTS.PICS.map(p => [p.name, p.role, p.target]);
  sheet.getRange(7, 5, picValues.length, 3).setValues(picValues)
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9.5)
    .setVerticalAlignment("middle");
  sheet.getRange("E7:G20")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange("G7:G20").setHorizontalAlignment("center");

  // Blok 3: Brand & Status (I5:J20)
  sheet.getRange("I5:J5").merge()
    .setValue("🏷️ BRAND & STATUS")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.SUBHEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.SUBHEADER_TEXT)
    .setHorizontalAlignment("center");

  sheet.getRange("I6:J6").setValues([["Brand", "Status Workflow"]])
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontWeight("bold")
    .setBackground("#F1F5F9")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  const brandRange = APP_CONSTANTS.DEFAULTS.BRANDS.map(b => [b]);
  sheet.getRange(7, 9, brandRange.length, 1).setValues(brandRange).setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5);

  const statusRange = APP_CONSTANTS.DEFAULTS.STATUSES.map(s => [s]);
  sheet.getRange(7, 10, statusRange.length, 1).setValues(statusRange).setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5);

  sheet.getRange("I7:J20")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
}

/**
 * Membaca seluruh konfigurasi aktif dari tab ⚙️ Config (fallback ke DEFAULTS jika kosong).
 */
function getAppConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.CONFIG);

  const config = {
    repoOwner: APP_CONSTANTS.DEFAULTS.REPO_OWNER,
    repoName: APP_CONSTANTS.DEFAULTS.REPO_NAME,
    branch: APP_CONSTANTS.DEFAULTS.REPO_BRANCH,
    basePath: APP_CONSTANTS.DEFAULTS.BASE_PATH,
    scheduleFile: APP_CONSTANTS.DEFAULTS.SCHEDULE_FILE,
    emailEnabled: APP_CONSTANTS.DEFAULTS.EMAIL_ENABLED === "ON",
    emailRecipients: [],
    pics: APP_CONSTANTS.DEFAULTS.PICS.map((p, idx) => ({
      name: p.name,
      role: p.role,
      target: p.target,
      row: 7 + idx
    })),
    brands: APP_CONSTANTS.DEFAULTS.BRANDS.slice(),
    statuses: APP_CONSTANTS.DEFAULTS.STATUSES.slice()
  };

  if (!sheet) return config;

  // Batch read settings C6:C12
  const settingVals = sheet.getRange("C6:C12").getValues().map(r => String(r[0] || "").trim());
  if (settingVals[0]) config.repoOwner = settingVals[0];
  if (settingVals[1]) config.repoName = settingVals[1];
  if (settingVals[2]) config.branch = settingVals[2];
  if (settingVals[3] !== undefined) config.basePath = settingVals[3].replace(/^\/+|\/+$/g, "");
  if (settingVals[4]) config.scheduleFile = settingVals[4];
  config.emailEnabled = settingVals[5].toUpperCase() === "ON";
  if (settingVals[6]) {
    config.emailRecipients = settingVals[6]
      .split(",")
      .map(e => e.trim())
      .filter(e => e.includes("@"));
  }

  // Batch read PICs E7:G20
  const picRows = sheet.getRange("E7:G20").getValues();
  const parsedPics = [];
  for (let i = 0; i < picRows.length; i++) {
    const name = String(picRows[i][0] || "").trim();
    if (!name) continue;
    const role = String(picRows[i][1] || "").trim();
    const targetNum = Number(picRows[i][2]);
    parsedPics.push({
      name: name,
      role: role,
      target: isNaN(targetNum) || targetNum < 0 ? 0 : targetNum,
      row: 7 + i
    });
  }
  if (parsedPics.length > 0) config.pics = parsedPics;

  // Batch read Brands I7:I20 & Statuses J7:J20
  const brandRows = sheet.getRange("I7:I20").getValues().map(r => String(r[0] || "").trim()).filter(Boolean);
  if (brandRows.length > 0) config.brands = brandRows;

  const statusRows = sheet.getRange("J7:J20").getValues().map(r => String(r[0] || "").trim()).filter(Boolean);
  if (statusRows.length > 0) config.statuses = statusRows;

  return config;
}

/**
 * Mengubah status Ringkasan Harian (ON/OFF) di tab ⚙️ Config.
 */
function setConfigEmailEnabled(enabled) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.CONFIG);
  if (!sheet) return;
  sheet.getRange("C11").setValue(enabled ? "ON" : "OFF");
}
