/**
 * VALMART.ID & MARTFORUM.ID — CREATIVE CONTENT TRACKER
 * Lean, High-Aesthetic & Automated Content Pipeline + Calendar Grid
 *
 * Dikembangkan untuk: Tim Kreatif Ko Jeremy (Hildan, Helmy, Audrey, Dika)
 */

const CONFIG = {
  SHEET_CALENDAR: "📅 Kalender Visual",
  SHEET_DATABASE: "📝 Database Konten",
  SHEET_CONFIG: "⚙️ Konfigurasi",
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
    MARTFORUM: "#0284C7",       // Sky / Cyan
    STATUS: {
      DRAFT: { bg: "#F1F5F9", text: "#475569" },       // Slate
      PROGRESS: { bg: "#FEF3C7", text: "#92400E" },    // Amber
      REVIEW: { bg: "#FFEDD5", text: "#9A3412" },      // Orange
      REVISION: { bg: "#FEE2E2", text: "#991B1B" },    // Red
      APPROVED: { bg: "#DCFCE7", text: "#166534" },    // Green
      PUBLISHED: { bg: "#EDE9FE", text: "#5B21B6" }    // Purple
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
 * Trigger saat spreadsheet dibuka: Buat custom menu
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("⚡ Valmart & Martforum")
    .addItem("📅 Render / Refresh Kalender Visual", "renderCalendar")
    .addSeparator()
    .addItem("🛠️ Setup Lembar Kerja (Reset & Format Cantik)", "setupTracker")
    .addItem("🌱 Muat Jadwal Aktif November 2026", "seedNovemberData")
    .addToUi();
}

/**
 * 1. FUNGSI UTAMA SETUP TRACKER: Format tampilan, dropdown, dan tata letak
 */
function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Tab Database Konten
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
  
  // Render kalender untuk bulan saat ini
  renderCalendar();

  SpreadsheetApp.getUi().alert(
    "✅ Tracker Siap Digunakan!",
    "Lembar kerja berhasil diformat dengan tampilan modern & lean.\n\n" +
    "Gunakan menu [⚡ Valmart & Martforum] -> [🌱 Muat Jadwal Aktif November 2026] untuk mengisi jadwal saat ini.",
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Format Tab Database Konten dengan styling modern
 */
function formatDatabaseSheet(sheet) {
  sheet.clear();
  sheet.clearFormats();
  sheet.setTabColor("#6366F1");

  // Title Banner
  sheet.getRange("B2:J2").merge()
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
  sheet.getRange("B3:J3").merge()
    .setValue("💡 Tips: Isi baris jadwal di bawah ini. Buka tab '📅 Kalender Visual' untuk melihat jadwal kalender otomatis.")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(9)
    .setFontColor("#64748B")
    .setBackground("#F8FAFC")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 24);

  // Headers
  const headers = [
    "No",
    "Tanggal Tayang",
    "Hari",
    "Brand",
    "PIC & Role",
    "Topik & Hook Konten",
    "Link Draft Preview",
    "Status",
    "Catatan / Feedback Review"
  ];

  const headerRange = sheet.getRange("B4:J4");
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
  sheet.setColumnWidth(1, 25);   // Padding A
  sheet.setColumnWidth(2, 45);   // No (B)
  sheet.setColumnWidth(3, 115);  // Tanggal (C)
  sheet.setColumnWidth(4, 85);   // Hari (D)
  sheet.setColumnWidth(5, 110);  // Brand (E)
  sheet.setColumnWidth(6, 160);  // PIC (F)
  sheet.setColumnWidth(7, 310);  // Topik (G)
  sheet.setColumnWidth(8, 200);  // Link (H)
  sheet.setColumnWidth(9, 130);  // Status (I)
  sheet.setColumnWidth(10, 240); // Feedback (J)

  // Data Validation Rules
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

  // Dropdown Status (I5:I100)
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.STATUS_LIST, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(5, 9, maxRows, 1).setDataValidation(statusRule);

  // Date Format for C5:C100
  sheet.getRange(5, 3, maxRows, 1).setNumberFormat("yyyy-mm-dd");

  // Conditional Formatting Warna Status (I5:I100)
  const rules = [];
  const statusRange = sheet.getRange("I5:I100");

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

  // Alternating row styling & borders
  const dataRange = sheet.getRange(5, 2, maxRows, 9);
  dataRange.setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(9.5)
    .setVerticalAlignment("middle");

  sheet.setFrozenRows(4);
}

/**
 * 2. RENDER KALENDER VISUAL (GRID BULANAN SENIN - MINGGU)
 */
function renderCalendar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calSheet = ss.getSheetByName(CONFIG.SHEET_CALENDAR);
  if (!calSheet) {
    calSheet = ss.insertSheet(CONFIG.SHEET_CALENDAR, 0);
  }

  const dbSheet = ss.getSheetByName(CONFIG.SHEET_DATABASE);
  if (!dbSheet) {
    SpreadsheetApp.getUi().alert("Error", "Tab '" + CONFIG.SHEET_DATABASE + "' tidak ditemukan. Jalankan Setup dulu.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  // Ambil pilihan bulan & tahun dari kontrol (atau default bulan ini)
  let targetMonth = calSheet.getRange("C2").getValue();
  let targetYear = calSheet.getRange("E2").getValue();

  const now = new Date();
  if (!targetMonth || isNaN(targetMonth) || targetMonth < 1 || targetMonth > 12) {
    targetMonth = now.getMonth() + 1; // 1-12
  }
  if (!targetYear || isNaN(targetYear) || targetYear < 2020) {
    targetYear = now.getFullYear();
  }

  // Clear area kalender dari row 5 ke bawah
  calSheet.getRange("A5:H60").clear().clearFormats();
  calSheet.setTabColor("#4F46E5");

  // Setup Top Header / Controls
  setupCalendarControls(calSheet, targetMonth, targetYear);

  // Baca database konten untuk bulan terpilih
  const contentMap = loadContentFromDatabase(dbSheet, targetMonth, targetYear);

  // Gambar Grid Kalender
  drawCalendarGrid(calSheet, targetMonth, targetYear, contentMap);
}

/**
 * Kontrol Top Bar Kalender (Bulan, Tahun, Tombol Refresh)
 */
function setupCalendarControls(sheet, month, year) {
  sheet.setColumnWidth(1, 20); // Spacing A

  // 7 Kolom Hari: B sampai H (Senin s/d Minggu)
  for (let c = 2; c <= 8; c++) {
    sheet.setColumnWidth(c, 175);
  }

  // Header Banner
  sheet.getRange("B2:H2").clear();
  sheet.getRange("B2").setValue("🗓️ KALENDER KONTEN:").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setFontColor("#1E293B");
  sheet.getRange("C2").setValue(month).setNumberFormat("0").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#EEF2FF");
  sheet.getRange("D2").setValue("TAHUN:").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setFontColor("#1E293B");
  sheet.getRange("E2").setValue(year).setNumberFormat("0").setFontFamily(CONFIG.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setBackground("#EEF2FF");
  
  const monthNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  sheet.getRange("F2:H2").merge()
    .setValue("📅 " + monthNames[month - 1] + " " + year + " (Edit C2/E2 lalu klik Refresh)")
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setFontColor("#4338CA")
    .setHorizontalAlignment("right");

  sheet.setRowHeight(2, 32);

  // Nama Hari (Senin - Minggu)
  const days = ["SENIN", "SELASA", "RABU", "KAMIS", "JUMAT", "SABTU", "MINGGU"];
  const dayRow = sheet.getRange("B4:H4");
  dayRow.setValues([days])
    .setFontFamily(CONFIG.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(CONFIG.COLORS.DAY_HEADER_BG)
    .setFontColor(CONFIG.COLORS.DAY_HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(4, 30);
}

/**
 * Membaca tab Database dan mengelompokkan konten per tanggal
 */
function loadContentFromDatabase(dbSheet, targetMonth, targetYear) {
  const lastRow = dbSheet.getLastRow();
  const map = {};

  if (lastRow < 5) return map;

  const data = dbSheet.getRange(5, 2, lastRow - 4, 9).getValues();

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
        link: data[i][6],        // Col H
        status: data[i][7],      // Col I
        notes: data[i][8]        // Col J
      });
    }
  }

  return map;
}

/**
 * Menggambar grid kotak kalender bulanan
 */
function drawCalendarGrid(sheet, month, year, contentMap) {
  const firstDay = new Date(year, month - 1, 1);
  const lastDate = new Date(year, month, 0).getDate();
  const today = new Date();
  const isCurrentMonth = (today.getFullYear() === year && (today.getMonth() + 1) === month);
  const todayDate = today.getDate();

  // Hari pertama dalam minggu (0 = Minggu, 1 = Senin, ..., 6 = Sabtu)
  let startCol = firstDay.getDay(); 
  // Konversi ke format Senin = 0, Minggu = 6
  startCol = (startCol === 0) ? 6 : startCol - 1;

  let currentDay = 1;
  let currentRow = 5;

  while (currentDay <= lastDate) {
    // 1 Baris Angka Tanggal (Header kotak hari)
    // 1 Baris Isi Konten (Isi kartu)
    const dateHeaderRow = currentRow;
    const contentRow = currentRow + 1;

    sheet.setRowHeight(dateHeaderRow, 24);
    sheet.setRowHeight(contentRow, 110); // Tinggi kartu konten yang lega

    for (let c = 0; c < 7; c++) {
      const colIndex = c + 2; // Mulai dari Kolom B (index 2)
      const cellDateHeader = sheet.getRange(dateHeaderRow, colIndex);
      const cellContent = sheet.getRange(contentRow, colIndex);

      if ((currentRow === 5 && c < startCol) || currentDay > lastDate) {
        // Kotak kosong di luar bulan
        cellDateHeader.setBackground("#F8FAFC").setBorder(true, true, true, true, null, null, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);
        cellContent.setBackground("#F8FAFC").setBorder(true, true, true, true, null, null, CONFIG.COLORS.GRID_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      } else {
        const isToday = isCurrentMonth && (currentDay === todayDate);

        // Header Angka Tanggal
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

        // Isi Konten Tanggal Tersebut
        const items = contentMap[currentDay] || [];
        if (items.length > 0) {
          let cardText = "";
          items.forEach((item, idx) => {
            const brandShort = (item.brand === "Valmart") ? "[VAL]" : "[MART]";
            const statusIcon = getStatusBadge(item.status);
            cardText += (idx > 0 ? "\n───\n" : "") +
              brandShort + " " + item.pic.split(" ")[0] + "\n" +
              "📌 " + item.topik + "\n" +
              statusIcon;
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

function getStatusBadge(status) {
  if (!status) return "⚪ Draft";
  return status;
}

/**
 * 3. SEED DATA NYATA: Muat jadwal riil tim November 2026 yang aktif sekarang
 */
function seedNovemberData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(CONFIG.SHEET_DATABASE);

  if (!dbSheet) {
    setupTracker();
    return;
  }

  const sampleData = [
    [1, "2026-11-16", "=TEXT(C5,\"dddd\")", "Valmart", "Helmy (Video)", "Mid-Month ACE Bounty Recap (Reaksi Klip 200k)", "", "🟣 Published", "Tayang Senin"],
    [2, "2026-11-16", "=TEXT(C6,\"dddd\")", "Martforum", "Audrey (VO)", "Teaser Resmi Brand: Top Up QRIS 3 Detik", "", "🟣 Published", "Selesai"],
    [3, "2026-11-17", "=TEXT(C7,\"dddd\")", "Valmart", "Helmy (Video)", "Meme 10 Omen Chaos di Ascent", "", "🟣 Published", "Tayang Selasa"],
    [4, "2026-11-18", "=TEXT(C8,\"dddd\")", "Valmart", "Helmy (Video)", "Sinergi Dua Brand: Sewa Valmart Diskon Martforum", "", "🟣 Published", "Tayang Rabu"],
    [5, "2026-11-18", "=TEXT(C9,\"dddd\")", "Martforum", "Audrey (VO)", "Meme Temen Discord Pas Ada Bundle Baru", "", "🟣 Published", "Tayang Rabu"],
    [6, "2026-11-19", "=TEXT(C10,\"dddd\")", "Valmart", "Helmy (Video)", "Anti-Scam Awareness: Bedain Akun Resmi vs Fake", "", "🟣 Published", "Tayang Kamis"],
    [7, "2026-11-20", "=TEXT(C11,\"dddd\")", "Valmart", "Helmy (Video)", "Meme Sekte Skin: Kuronami vs Reaver vs Primordium", "", "🟣 Published", "Tayang Jumat"],
    [8, "2026-11-20", "=TEXT(C12,\"dddd\")", "Martforum", "Audrey (VO)", "Weekend Mabar: Anti-Phishing Cukup Riot ID", "", "🟢 Approved", "Naskah Baru Audrey"],
    [9, "2026-11-21", "=TEXT(C13,\"dddd\")", "Valmart", "Dika (Editor/Desain)", "Infografis: Panduan Sewa Goceng 1 Menit Anti-Ribet", "", "🟢 Approved", "Desain Carousel Dika"],
    [10, "2026-11-22", "=TEXT(C14,\"dddd\")", "Valmart", "Helmy (Video)", "Reminder Sayembara #ValmartAceBounty (Tersisa 8 Hari)", "", "🟡 In Progress", "Story & Feed Push"],
    [11, "2026-11-23", "=TEXT(C15,\"dddd\")", "Martforum", "Audrey (VO)", "Derita Solo Queue vs Obatnya (Mabar Skin Wangi)", "", "🟠 In Review", "Jadwal Senin Week 4"],
    [12, "2026-11-24", "=TEXT(C16,\"dddd\")", "Valmart", "Helmy (Video)", "Smart Buyer: Ngetes Skin di Valmart Sebelum Boncos", "", "🟡 In Progress", "Jadwal Selasa Week 4"],
    [13, "2026-11-25", "=TEXT(C17,\"dddd\")", "Martforum", "Audrey (VO)", "Meme Streamer Jago vs Realita Whiff Kita", "", "⚪ Draft", "Jadwal Rabu Week 4"],
    [14, "2026-11-25", "=TEXT(C18,\"dddd\")", "Valmart", "Helmy (Video)", "Meme Tanggal Tua Tetep Bisa Sultan (Modal Goceng)", "", "⚪ Draft", "Jadwal Rabu Week 4"],
    [15, "2026-11-26", "=TEXT(C19,\"dddd\")", "Valmart", "Helmy (Video)", "Edukasi Anti-HB: Rahasia First Mail & Garansi 100%", "", "⚪ Draft", "Jadwal Kamis Week 4"],
    [16, "2026-11-27", "=TEXT(C20,\"dddd\")", "Martforum", "Audrey (VO)", "PAYDAY FLASH SALE: Amankan VP Akhir Bulan QRIS 3 Detik", "", "⚪ Draft", "Jadwal Jumat Payday"],
    [17, "2026-11-27", "=TEXT(C21,\"dddd\")", "Valmart", "Helmy (Video)", "Showcase Katalog Akun Sultan Weekend Ready Sewa", "", "⚪ Draft", "Jadwal Jumat Week 4"],
    [18, "2026-11-30", "=TEXT(C22,\"dddd\")", "Valmart", "Dika (Editor/Desain)", "Poster Pemenang Sayembara ACE Bounty (Rp 200k Cash)", "", "⚪ Draft", "Closing November"]
  ];

  // Bersihkan data lama dari row 5
  dbSheet.getRange(5, 2, sampleData.length, 9).setValues(sampleData);

  // Render ulang kalender agar datanya langsung tampil
  renderCalendar();

  SpreadsheetApp.getUi().alert("🌱 Data Berhasil Dimuat!", "Jadwal aktif November 2026 (Week 3 & Week 4) sudah masuk ke Database dan Kalender Visual.", SpreadsheetApp.getUi().ButtonSet.OK);
}
