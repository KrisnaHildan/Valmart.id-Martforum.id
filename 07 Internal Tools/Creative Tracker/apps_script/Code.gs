/**
 * =========================================================================
 * VALMART.ID & MARTFORUM.ID — CREATIVE CONTENT ENGINE & TRACKER
 * =========================================================================
 * Integrasi Dua Arah:
 * 1. GitHub Repository: Sebagai Library Naskah & Dokumen Detail
 * 2. Google Sheets: Sebagai Dashboard Visual & Tracker Kalender Harian
 *
 * Dikembangkan untuk: Tim Kreatif Ko Jeremy (Hildan, Helmy, Audrey, Dika)
 * =========================================================================
 */

const CONFIG = {
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
    "Dika (Editor/Desain)"
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
    .addItem("📅 Render / Refresh Kalender Visual", "renderCalendar")
    .addSeparator()
    .addItem("📥 Sinkronisasi Naskah dari GitHub (Nov 2026)", "syncNovemberSchedule")
    .addItem("🛠️ Setup Ulang Lembar Kerja (Format Cantik)", "setupTracker")
    .addToUi();
}

/**
 * Setup Tab Database & Tab Kalender dengan Desain Modern
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
    calSheet = ss.insertSheet(CONFIG.SHEET_CALENDAR, 0);
  }
  
  renderCalendar();

  SpreadsheetApp.getUi().alert(
    "✅ Sistem Berhasil Dibuat!",
    "Format lembar kerja siap dipakai.\n\n" +
    "Klik menu [⚡ Valmart & Martforum] -> [📥 Sinkronisasi Naskah dari GitHub] untuk mengisi jadwal konten aktif.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Format Tab Database dengan Desain Tabel Modern
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

  // Subtitle info
  sheet.getRange("B3:K3").merge()
    .setValue("💡 Info: Kolom 'Link Naskah GitHub' langsung mengarah ke naskah detail per detik di repository. Editor cukup drop link draft di kolom 'Link Draft Preview'.")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(9)
    .setFontColor("#64748B")
    .setBackground("#F8FAFC")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 24);

  // 10 Kolom Rapi
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

  const headerRange = sheet.getRange("B4:K4");
  headerRange.setValues([headers])
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.DAY_HEADER_BG)
    .setFontColor(CONFIG.COLORS.DAY_HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(4, 34);

  // Column Widths
  sheet.setColumnWidth(1, 20);   // Spacing A
  sheet.setColumnWidth(2, 45);   // No (B)
  sheet.setColumnWidth(3, 115);  // Tanggal (C)
  sheet.setColumnWidth(4, 85);   // Hari (D)
  sheet.setColumnWidth(5, 110);  // Brand (E)
  sheet.setColumnWidth(6, 160);  // PIC (F)
  sheet.setColumnWidth(7, 300);  // Topik (G)
  sheet.setColumnWidth(8, 200);  // Link Naskah GitHub (H)
  sheet.setColumnWidth(9, 190);  // Link Draft (I)
  sheet.setColumnWidth(10, 130); // Status (J)
  sheet.setColumnWidth(11, 240); // Feedback (K)

  const maxRows = 100;

  // Dropdown Brand (E5:E100)
  const brandRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.BRAND_LIST, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(5, 5, maxRows, 1).setDataValidation(brandRule);

  // Dropdown PIC (F5:F100)
  const picRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.PIC_LIST, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(5, 6, maxRows, 1).setDataValidation(picRule);

  // Dropdown Status (J5:J100)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STATUS_LIST, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(5, 10, maxRows, 1).setDataValidation(statusRule);

  // Date format C5:C100
  sheet.getRange(5, 3, maxRows, 1).setNumberFormat("yyyy-mm-dd");

  // Conditional Formatting Status Colors
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
 * Render Grid Kalender Visual (Senin s/d Minggu)
 */
function renderCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calSheet = ss.getSheetByName(CONFIG.SHEET_CALENDAR);
  if (!calSheet) {
    calSheet = ss.insertSheet(CONFIG.SHEET_CALENDAR, 0);
  }

  const dbSheet = ss.getSheetByName(CONFIG.SHEET_DATABASE);
  if (!dbSheet) return;

  // Baca pilihan bulan & tahun dari sel C2 & E2
  let targetMonth = calSheet.getRange("C2").getValue();
  let targetYear = calSheet.getRange("E2").getValue();

  const now = new Date();
  if (!targetMonth || isNaN(targetMonth) || targetMonth < 1 || targetMonth > 12) {
    targetMonth = now.getMonth() + 1;
  }
  if (!targetYear || isNaN(targetYear) || targetYear < 2020) {
    targetYear = now.getFullYear();
  }

  // Clear area kalender lama dari baris 5 ke bawah
  const maxRows = calSheet.getMaxRows();
  if (maxRows >= 5) {
    calSheet.getRange(5, 1, maxRows - 4, calSheet.getMaxColumns()).clear();
  }
  calSheet.setTabColor("#4F46E5");

  // Setup Top Controls
  setupCalendarControls(calSheet, targetMonth, targetYear);

  // Ambil data konten dari tab Database
  const contentMap = loadContentFromDatabase(dbSheet, targetMonth, targetYear);

  // Render Grid
  drawCalendarGrid(calSheet, targetMonth, targetYear, contentMap);
}

function setupCalendarControls(sheet, month, year) {
  sheet.setColumnWidth(1, 20); // Spacing A

  for (let c = 2; c <= 8; c++) {
    sheet.setColumnWidth(c, 180);
  }

  // Baris Kontrol Bulan/Tahun
  sheet.getRange("B2").setValue("🗓️ BULAN:").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setFontColor("#1E293B");
  sheet.getRange("C2").setValue(month).setNumberFormat("0").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#EEF2FF");
  sheet.getRange("D2").setValue("TAHUN:").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setFontColor("#1E293B");
  sheet.getRange("E2").setValue(year).setNumberFormat("0").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#EEF2FF");

  const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  sheet.getRange("F2:H2").merge()
    .setValue("📅 " + monthNames[month - 1] + " " + year + " (Ganti sel C2/E2 lalu klik menu Refresh)")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setFontColor("#4338CA")
    .setHorizontalAlignment("right");
  sheet.setRowHeight(2, 32);

  // Nama Hari
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
    const rawDate = data[i][1]; // Tanggal (Col C)
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
        brand: data[i][3],       // Col E
        pic: data[i][4],         // Col F
        topik: data[i][5],       // Col G
        scriptUrl: data[i][6],   // Col H (Link GitHub)
        draftUrl: data[i][7],    // Col I (Link Draft)
        status: data[i][8],      // Col J (Status)
        feedback: data[i][9]     // Col K (Notes)
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
  startCol = (startCol === 0) ? 6 : startCol - 1; // Senin = 0

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

        // Header Tanggal
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

        // Konten Kartu
        const items = contentMap[currentDay] || [];
        if (items.length > 0) {
          let cardText = "";
          items.forEach((item, idx) => {
            const brandTag = (item.brand === "Valmart") ? "[VAL]" : "[MART]";
            const statusTag = item.status ? item.status : "⚪ Draft";
            cardText += (idx > 0 ? "\n───\n" : "") +
              brandTag + " " + item.pic.split(" ")[0] + "\n" +
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
 * Sinkronisasi Naskah dari GitHub: Memasukkan Jadwal Aktif dengan Link Langsung ke Repo
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
    [1, "2026-11-16", "=TEXT(C5,\"dddd\")", "Valmart", "Helmy (Video)", "Mid-Month ACE Bounty Recap (Reaksi Klip 200k)", baseW3Val, "", "🟣 Published", "Tayang Senin W3"],
    [2, "2026-11-16", "=TEXT(C6,\"dddd\")", "Martforum", "Audrey (VO)", "Teaser Resmi Brand: Top Up QRIS 3 Detik", baseW3Mart, "", "🟣 Published", "Selesai"],
    [3, "2026-11-17", "=TEXT(C7,\"dddd\")", "Valmart", "Helmy (Video)", "Meme 10 Omen Chaos di Ascent", baseW3Val, "", "🟣 Published", "Tayang Selasa W3"],
    [4, "2026-11-18", "=TEXT(C8,\"dddd\")", "Valmart", "Helmy (Video)", "Sinergi Dua Brand: Sewa Valmart Diskon Martforum", baseW3Val, "", "🟣 Published", "Tayang Rabu W3"],
    [5, "2026-11-18", "=TEXT(C9,\"dddd\")", "Martforum", "Audrey (VO)", "Meme Temen Discord Pas Ada Bundle Baru", baseW3Mart, "", "🟣 Published", "Tayang Rabu W3"],
    [6, "2026-11-19", "=TEXT(C10,\"dddd\")", "Valmart", "Helmy (Video)", "Anti-Scam Awareness: Bedain Akun Resmi vs Fake", baseW3Val, "", "🟣 Published", "Tayang Kamis W3"],
    [7, "2026-11-20", "=TEXT(C11,\"dddd\")", "Valmart", "Helmy (Video)", "Meme Sekte Skin: Kuronami vs Reaver vs Primordium", baseW3Val, "", "🟣 Published", "Tayang Jumat W3"],
    [8, "2026-11-20", "=TEXT(C12,\"dddd\")", "Martforum", "Audrey (VO)", "Weekend Mabar: Anti-Phishing Cukup Riot ID", baseW3Mart, "", "🟢 Approved", "Naskah Baru Audrey"],
    [9, "2026-11-21", "=TEXT(C13,\"dddd\")", "Valmart", "Dika (Editor/Desain)", "Infografis: Panduan Sewa Goceng 1 Menit Anti-Ribet", baseW3Mart, "", "🟢 Approved", "Desain Carousel Dika"],
    [10, "2026-11-22", "=TEXT(C14,\"dddd\")", "Valmart", "Helmy (Video)", "Reminder Sayembara #ValmartAceBounty (Tersisa 8 Hari)", baseW3Val, "", "🟢 Approved", "Story & Feed Push"],

    // Week 4 (Minggu Ini)
    [11, "2026-11-23", "=TEXT(C15,\"dddd\")", "Martforum", "Audrey (VO)", "Derita Solo Queue vs Obatnya (Mabar Skin Wangi)", baseW4Mart, "", "🟠 In Review", "Jadwal Hari Ini (Senin)"],
    [12, "2026-11-24", "=TEXT(C16,\"dddd\")", "Valmart", "Helmy (Video)", "Smart Buyer: Ngetes Skin di Valmart Sebelum Boncos", baseW4Val, "", "🟡 In Progress", "Jadwal Selasa W4"],
    [13, "2026-11-25", "=TEXT(C17,\"dddd\")", "Martforum", "Audrey (VO)", "Meme Streamer Jago vs Realita Whiff Kita", baseW4Mart, "", "⚪ Draft", "Jadwal Rabu W4"],
    [14, "2026-11-25", "=TEXT(C18,\"dddd\")", "Valmart", "Helmy (Video)", "Meme Tanggal Tua Tetep Bisa Sultan (Modal Goceng)", baseW4Val, "", "⚪ Draft", "Jadwal Rabu W4"],
    [15, "2026-11-26", "=TEXT(C19,\"dddd\")", "Valmart", "Helmy (Video)", "Edukasi Anti-HB: Rahasia First Mail & Garansi 100%", baseW4Val, "", "⚪ Draft", "Jadwal Kamis W4"],
    [16, "2026-11-27", "=TEXT(C20,\"dddd\")", "Martforum", "Audrey (VO)", "PAYDAY FLASH SALE: Amankan VP Akhir Bulan QRIS 3 Detik", baseW4Mart, "", "⚪ Draft", "Jadwal Jumat Payday"],
    [17, "2026-11-27", "=TEXT(C21,\"dddd\")", "Valmart", "Helmy (Video)", "Showcase Katalog Akun Sultan Weekend Ready Sewa", baseW4Val, "", "⚪ Draft", "Jadwal Jumat W4"],
    [18, "2026-11-30", "=TEXT(C22,\"dddd\")", "Valmart", "Dika (Editor/Desain)", "Poster Pemenang Sayembara ACE Bounty (Rp 200k Cash)", baseW4Mart, "", "⚪ Draft", "Closing November"]
  ];

  dbSheet.getRange(5, 2, schedule.length, 10).setValues(schedule);
  renderCalendar();

  SpreadsheetApp.getUi().alert(
    "📥 Sinkronisasi Berhasil!",
    "Seluruh jadwal aktif November 2026 berhasil disinkronkan langsung dengan tautan naskah di GitHub.\n" +
    "Kalender visual otomatis ter-refresh.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}
