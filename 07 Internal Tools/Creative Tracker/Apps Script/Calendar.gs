/**
 * =========================================================================
 * CALENDAR.GS — Render Kalender Visual Bulanan (Batch & Auto-Refresh)
 * =========================================================================
 */

/**
 * Merender tab 📅 Kalender Visual berdasarkan bulan (C2) & tahun (E2).
 * Gunakan options = { skipLock: true } jika pemanggil sudah memegang scriptLock.
 */
function renderCalendar(options) {
  const opts = options || {};
  const lock = LockService.getScriptLock();
  if (!opts.skipLock) {
    if (!lock.tryLock(10000)) {
      return;
    }
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let calSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.CALENDAR);
    if (!calSheet) {
      calSheet = ss.insertSheet(APP_CONSTANTS.SHEETS.CALENDAR);
    }

    const dbSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.DATABASE);
    if (!dbSheet) return;

    const now = new Date();
    const targetMonth = parseMonthFilterValue(calSheet.getRange("C2").getValue(), false, now.getMonth() + 1);
    const targetYear = parseYearFilterValue(calSheet.getRange("E2").getValue(), false, now.getFullYear());

    const maxRows = calSheet.getMaxRows();
    if (maxRows >= 5) {
      calSheet.getRange(5, 1, maxRows - 4, calSheet.getMaxColumns()).clear();
    }
    calSheet.setTabColor(APP_CONSTANTS.COLORS.VALMART);

    setupCalendarControls(calSheet, targetMonth, targetYear);
    const contentMap = loadContentMapFromDatabase(dbSheet, targetMonth, targetYear);
    drawCalendarGridBatch(calSheet, targetMonth, targetYear, contentMap);
  } finally {
    if (!opts.skipLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Menyiapkan kontrol dropdown Bulan (C2) & Tahun (E2) serta header hari Senin-Minggu.
 */
function setupCalendarControls(sheet, month, year) {
  sheet.setColumnWidth(1, 20);
  for (let c = 2; c <= 8; c++) {
    sheet.setColumnWidth(c, 185);
  }

  sheet.getRange("B2").setValue("🗓️ BULAN:").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setFontColor("#1E293B");
  sheet.getRange("C2").setValue(month).setNumberFormat("0").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setBackground(APP_CONSTANTS.COLORS.ACCENT_SOFT);
  sheet.getRange("D2").setValue("TAHUN:").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setFontColor("#1E293B");
  sheet.getRange("E2").setValue(year).setNumberFormat("0").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(11).setFontWeight("bold").setHorizontalAlignment("center").setBackground(APP_CONSTANTS.COLORS.ACCENT_SOFT);

  sheet.getRange("C2").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["1","2","3","4","5","6","7","8","9","10","11","12"], true).setAllowInvalid(false).build()
  );
  sheet.getRange("E2").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(getSupportedYearsList(false), true).setAllowInvalid(true).build()
  );

  sheet.getRange("F2:H2").merge()
    .setValue(`📅 ${getIndonesianMonthName(month)} ${year} (Otomatis update saat C2/E2 diubah)`)
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setFontColor("#4338CA")
    .setHorizontalAlignment("right");
  sheet.setRowHeight(2, 32);

  const days = ["SENIN", "SELASA", "RABU", "KAMIS", "JUMAT", "SABTU", "MINGGU"];
  sheet.getRange("B4:H4").setValues([days])
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.SUBHEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.SUBHEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(4, 30);
}

/**
 * Membaca data konten untuk bulan & tahun tertentu dalam satu kali batch read.
 */
function loadContentMapFromDatabase(dbSheet, targetMonth, targetYear) {
  const lastRow = dbSheet.getLastRow();
  const map = {};
  if (lastRow < APP_CONSTANTS.DB_START_ROW) return map;

  const numRows = lastRow - APP_CONSTANTS.DB_START_ROW + 1;
  const data = dbSheet.getRange(APP_CONSTANTS.DB_START_ROW, APP_CONSTANTS.DB_COLS.ID, numRows, 11).getValues();

  for (let i = 0; i < data.length; i++) {
    const dt = parseDateSafe(data[i][1]); // Kolom C = Tanggal Tayang
    if (!dt) continue;

    if ((dt.getMonth() + 1) === targetMonth && dt.getFullYear() === targetYear) {
      const dateNum = dt.getDate();
      if (!map[dateNum]) map[dateNum] = [];
      map[dateNum].push({
        brand: String(data[i][3] || "").trim(),
        pic: String(data[i][4] || "").trim(),
        topik: String(data[i][5] || "").trim(),
        status: String(data[i][8] || "").trim()
      });
    }
  }
  return map;
}

/**
 * Menghitung matriks grid kalender (dapat diuji unit test tanpa akses UI).
 */
function buildCalendarMatrix(month, year, contentMap, todayRef) {
  const firstDay = new Date(year, month - 1, 1);
  const lastDate = new Date(year, month, 0).getDate();
  const today = todayRef || new Date();
  const isCurrentMonth = (today.getFullYear() === year && (today.getMonth() + 1) === month);
  const todayDate = today.getDate();

  let startCol = firstDay.getDay();
  startCol = (startCol === 0) ? 6 : startCol - 1; // Senin = 0, Minggu = 6

  const weeks = [];
  let currentDay = 1;
  let weekIdx = 0;

  while (currentDay <= lastDate) {
    const headerVals = [];
    const headerBgs = [];
    const headerColors = [];
    const contentVals = [];
    const contentBgs = [];
    const contentColors = [];
    let todayCol = -1;
    let maxCards = 0;

    for (let c = 0; c < 7; c++) {
      const isWeekend = (c >= 5);
      if ((weekIdx === 0 && c < startCol) || currentDay > lastDate) {
        headerVals.push("");
        headerBgs.push(APP_CONSTANTS.COLORS.OUT_OF_MONTH_BG);
        headerColors.push(APP_CONSTANTS.COLORS.OUT_OF_MONTH_TEXT);
        contentVals.push("");
        contentBgs.push(APP_CONSTANTS.COLORS.OUT_OF_MONTH_BG);
        contentColors.push(APP_CONSTANTS.COLORS.OUT_OF_MONTH_TEXT);
      } else {
        const isToday = isCurrentMonth && (currentDay === todayDate);
        if (isToday) todayCol = c;

        const cellDate = new Date(year, month - 1, currentDay);
        const items = contentMap[currentDay] || [];
        let hasLate = false;
        const visibleCount = Math.min(items.length, APP_CONSTANTS.CALENDAR_MAX_ITEMS_PER_DAY);
        const effectiveCount = items.length > APP_CONSTANTS.CALENDAR_MAX_ITEMS_PER_DAY ? visibleCount + 1 : visibleCount;
        if (effectiveCount > maxCards) maxCards = effectiveCount;

        if (items.length > 0) {
          const maxShow = APP_CONSTANTS.CALENDAR_MAX_ITEMS_PER_DAY;
          const visibleItems = items.slice(0, maxShow);
          let cardText = visibleItems.map(item => {
            const late = isTaskLate(cellDate, item.status, today);
            if (late) hasLate = true;
            const brandLower = item.brand.toLowerCase();
            const brandTag = brandLower.includes("valmart")
              ? "[VAL]"
              : (brandLower.includes("martforum") ? "[MART]" : `[${item.brand || "-"}]`);
            const statusTag = item.status || "⚪ Draft";
            const picShort = item.pic ? item.pic.split(" ")[0] : "-";
            const topikShort = item.topik.length > 42 ? item.topik.substring(0, 39) + "..." : item.topik;
            const latePrefix = late ? "🚨 TELAT | " : "";
            return `${latePrefix}${brandTag} ${picShort}\n📌 ${topikShort}\n${statusTag}`;
          }).join("\n───\n");

          if (items.length > maxShow) {
            if (items.slice(maxShow).some(it => isTaskLate(cellDate, it.status, today))) {
              hasLate = true;
            }
            cardText += `\n+${items.length - maxShow} lainnya`;
          }

          headerVals.push(currentDay + (isToday ? "  📍 HARI INI" : (hasLate ? "  🚨 TELAT" : "")));
          headerBgs.push(isToday ? APP_CONSTANTS.COLORS.ACCENT : (hasLate ? APP_CONSTANTS.COLORS.LATE_BG : (isWeekend ? "#E2E8F0" : "#F1F5F9")));
          headerColors.push(isToday ? "#FFFFFF" : (hasLate ? APP_CONSTANTS.COLORS.LATE_TEXT : "#334155"));

          contentVals.push(cardText);
          contentBgs.push(isToday ? APP_CONSTANTS.COLORS.ACCENT_SOFT : (hasLate ? "#FEF2F2" : (isWeekend ? APP_CONSTANTS.COLORS.WEEKEND_BG : "#FFFFFF")));
          contentColors.push("#0F172A");
        } else {
          headerVals.push(currentDay + (isToday ? "  📍 HARI INI" : ""));
          headerBgs.push(isToday ? APP_CONSTANTS.COLORS.ACCENT : (isWeekend ? "#E2E8F0" : "#F1F5F9"));
          headerColors.push(isToday ? "#FFFFFF" : "#334155");

          contentVals.push("-");
          contentBgs.push(isToday ? APP_CONSTANTS.COLORS.ACCENT_SOFT : (isWeekend ? APP_CONSTANTS.COLORS.WEEKEND_BG : "#FFFFFF"));
          contentColors.push("#94A3B8");
        }
        currentDay++;
      }
    }

    weeks.push({
      headerVals, headerBgs, headerColors,
      contentVals, contentBgs, contentColors,
      todayCol,
      maxCards
    });
    weekIdx++;
  }

  return weeks;
}

/**
 * Menggambar seluruh grid kalender menggunakan operasi batch per blok.
 */
function drawCalendarGridBatch(sheet, month, year, contentMap) {
  const weeks = buildCalendarMatrix(month, year, contentMap, new Date());
  if (weeks.length === 0) return;

  const startRow = 5;
  const totalGridRows = weeks.length * 2;
  const allVals = [];
  const allBgs = [];
  const allColors = [];
  const allWeights = [];
  const allSizes = [];
  const allVAlign = [];
  const allHAlign = [];

  let todayGridCoords = null;

  weeks.forEach((wk, idx) => {
    const hRow = startRow + idx * 2;
    const cRow = hRow + 1;
    const mc = wk.maxCards || 0;
    const contentHeight = mc <= 0 ? 85 : (mc === 1 ? 105 : (mc === 2 ? 140 : (mc === 3 ? 180 : 220)));
    sheet.setRowHeight(hRow, 24);
    sheet.setRowHeight(cRow, contentHeight);

    allVals.push(wk.headerVals, wk.contentVals);
    allBgs.push(wk.headerBgs, wk.contentBgs);
    allColors.push(wk.headerColors, wk.contentColors);
    allWeights.push(
      Array(7).fill("bold"),
      Array(7).fill("normal")
    );
    allSizes.push(
      Array(7).fill(9),
      Array(7).fill(8.5)
    );
    allVAlign.push(
      Array(7).fill("middle"),
      Array(7).fill("top")
    );
    allHAlign.push(
      Array(7).fill("left"),
      Array(7).fill("left")
    );

    if (wk.todayCol >= 0) {
      todayGridCoords = { row: hRow, col: wk.todayCol + 2 };
    }
  });

  const gridRange = sheet.getRange(startRow, 2, totalGridRows, 7);
  gridRange
    .setValues(allVals)
    .setBackgrounds(allBgs)
    .setFontColors(allColors)
    .setFontWeights(allWeights)
    .setFontSizes(allSizes)
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setVerticalAlignments(allVAlign)
    .setHorizontalAlignments(allHAlign)
    .setWrap(true)
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  if (todayGridCoords) {
    sheet.getRange(todayGridCoords.row, todayGridCoords.col, 2, 1)
      .setBorder(true, true, true, true, null, null, APP_CONSTANTS.COLORS.ACCENT, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
}
