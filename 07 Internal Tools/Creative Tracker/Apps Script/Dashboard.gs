/**
 * =========================================================================
 * DASHBOARD.GS — Executive Dashboard, KPI, Pacing Talent, Review & Telat
 * =========================================================================
 */

/**
 * Membangun / memperbarui tab 📊 Dashboard.
 */
function renderDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.DASHBOARD);
  if (!sheet) {
    sheet = ss.insertSheet(APP_CONSTANTS.SHEETS.DASHBOARD, 0);
  }

  // Pertahankan pilihan filter periode user jika sudah ada
  const prevMonth = sheet.getRange("C4").getValue();
  const prevYear = sheet.getRange("E4").getValue();
  const prevWeekOpt = sheet.getRange("I4").getValue();

  const now = new Date();
  const activeMonth = parseMonthFilterValue(prevMonth, false, now.getMonth() + 1);
  const activeYear = parseYearFilterValue(prevYear, false, now.getFullYear());
  const activeWeekOpt = prevWeekOpt ? String(prevWeekOpt) : "Minggu Ini";

  const config = getAppConfig();
  const isIndo = isIndonesianLocale(ss);
  const f = (formulaStr) => localizeFormula(formulaStr, isIndo);
  const dbName = `'${APP_CONSTANTS.SHEETS.DATABASE}'`;
  const cfgName = `'${APP_CONSTANTS.SHEETS.CONFIG}'`;
  const dbStart = APP_CONSTANTS.DB_START_ROW;

  // Pastikan minimal 26 kolom (A..Z) agar kolom bantu tersembunyi O..Z selalu tersedia
  if (sheet.getMaxColumns() < 26) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), 26 - sheet.getMaxColumns());
  }
  sheet.showColumns(1, 14);
  sheet.hideColumns(15, 12);

  // Lepas proteksi, merge lama, & conditional formatting sebelum clear, lalu pasang lagi di akhir
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sheet.clearConditionalFormatRules();
  sheet.getRange(1, 1, Math.max(sheet.getMaxRows(), 50), Math.max(sheet.getMaxColumns(), 26)).breakApart();
  sheet.clear();
  sheet.setTabColor("#10B981");

  // Lebar kolom A..M (L = Link Draft agar bisa diklik langsung dari Antrean Review)
  const colWidths = [
    [1, 20],  // A Spacer
    [2, 175], // B
    [3, 90],  // C
    [4, 90],  // D
    [5, 90],  // E
    [6, 90],  // F
    [7, 150], // G
    [8, 22],  // H Spacer
    [9, 95],  // I Tanggal
    [10, 145],// J Brand & PIC
    [11, 215],// K Topik & Status
    [12, 115],// L Link Draft
    [13, 210] // M Catatan Review
  ];
  colWidths.forEach(([c, w]) => sheet.setColumnWidth(c, w));

  // Header Banner
  sheet.getRange("B2:M2").merge()
    .setValue("📊 CREATIVE DASHBOARD - VALMART.ID & MARTFORUM.ID")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.HEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 38);

  sheet.getRange("B3:M3").merge()
    .setValue("💡 Pilih Bulan/Tahun & Minggu di baris 4 untuk metrik bulanan & pacing. Antrean Review dan Konten Telat memantau seluruh periode.")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontColor(APP_CONSTANTS.COLORS.MUTED_TEXT)
    .setBackground(APP_CONSTANTS.COLORS.MUTED_BG)
    .setHorizontalAlignment("center");
  sheet.setRowHeight(3, 22);

  // Baris 4: Filter Periode
  sheet.setRowHeight(4, 28);
  sheet.getRange("B4").setValue("🗓️ Filter Bulan:").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  sheet.getRange("C4").setValue(activeMonth).setNumberFormat("0").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground(APP_CONSTANTS.COLORS.ACCENT_SOFT);
  sheet.getRange("D4").setValue("Tahun:").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  sheet.getRange("E4").setValue(activeYear).setNumberFormat("0").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(10).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground(APP_CONSTANTS.COLORS.ACCENT_SOFT);

  sheet.getRange("G4").setValue("📆 Pacing Minggu:").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  sheet.getRange("I4:J4").merge().setValue(activeWeekOpt).setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setFontWeight("bold").setHorizontalAlignment("center").setVerticalAlignment("middle").setBackground(APP_CONSTANTS.COLORS.ACCENT_SOFT);

  // K4 & L4: Tanggal Senin & Minggu terpilih (formula otomatis)
  sheet.getRange("K4").setFormula(f('=IF(I4="Minggu Lalu", TODAY()-WEEKDAY(TODAY(),2)+1-7, IF(I4="Minggu Depan", TODAY()-WEEKDAY(TODAY(),2)+1+7, TODAY()-WEEKDAY(TODAY(),2)+1))'))
    .setNumberFormat("dd mmm yyyy").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(8.5).setFontColor(APP_CONSTANTS.COLORS.MUTED_TEXT).setHorizontalAlignment("center").setVerticalAlignment("middle");
  sheet.getRange("L4").setFormula(f('="s/d " & TEXT(K4+6, "dd mmm yyyy")'))
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(8.5).setFontColor(APP_CONSTANTS.COLORS.MUTED_TEXT).setHorizontalAlignment("left").setVerticalAlignment("middle");

  // Dropdown filter bulan, tahun, dan minggu
  sheet.getRange("C4").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["1","2","3","4","5","6","7","8","9","10","11","12"], true).setAllowInvalid(false).build()
  );
  sheet.getRange("E4").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(getSupportedYearsList(false), true).setAllowInvalid(true).build()
  );
  sheet.getRange("I4").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Minggu Lalu", "Minggu Ini", "Minggu Depan"], true).setAllowInvalid(false).build()
  );

  // -------------------------------------------------------------------------
  // 5 KPI SCORECARDS (Baris 6 - 8) - Open-ended range $6:$C agar tembus > 1000 baris
  // -------------------------------------------------------------------------
  sheet.setRowHeight(5, 10);
  sheet.setRowHeight(6, 22);
  sheet.setRowHeight(7, 38);
  sheet.setRowHeight(8, 20);

  const dateCondStart = `${dbName}!$C$${dbStart}:$C, ">=" & DATE($E$4, $C$4, 1)`;
  const dateCondEnd = `${dbName}!$C$${dbStart}:$C, "<=" & EOMONTH(DATE($E$4, $C$4, 1), 0)`;
  const statusRange = `${dbName}!$J$${dbStart}:$J`;

  createScoreCard(sheet, "B6:C6", "B7:C7", "B8:C8",
    "🎯 TOTAL TERJADWAL",
    f(`=COUNTIFS(${dateCondStart}, ${dateCondEnd})`),
    f(`=TEXT(DATE($E$4, $C$4, 1), "mmmm yyyy")`),
    "#F8FAFC", "#0F172A", "#CBD5E1"
  );

  createScoreCard(sheet, "D6:E6", "D7:E7", "D8:E8",
    "🚀 APPROVED / PUBLISHED",
    f(`=COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${statusRange}, "🟢 Approved") + COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${statusRange}, "🟣 Published")`),
    f(`=IF(B7>0, TEXT(D7/B7, "0%") & " Selesai", "0% Selesai")`),
    "#F0FDF4", "#166534", "#86EFAC"
  );

  createScoreCard(sheet, "F6:G6", "F7:G7", "F8:G8",
    "⚠️ ANTREAN REVIEW",
    f(`=COUNTIFS(${statusRange}, "🟠 In Review") + COUNTIFS(${statusRange}, "🔴 Revision")`),
    "Semua Bulan (Dicek Jeje/Hildan)",
    "#FFF7ED", "#C2410C", "#FDBA74"
  );

  createScoreCard(sheet, "I6:K6", "I7:K7", "I8:K8",
    "⏳ DALAM PENGERJAAN",
    f(`=COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${statusRange}, "⚪ Draft") + COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${statusRange}, "🟡 In Progress")`),
    "On Track Talent (Bulan Ini)",
    "#FEFCE8", "#A16207", "#FDE047"
  );

  createScoreCard(sheet, "L6:M6", "L7:M7", "L8:M8",
    "🚨 KONTEN TELAT",
    f(`=COUNTIFS(${dbName}!$C$${dbStart}:$C, "<" & TODAY(), ${dbName}!$C$${dbStart}:$C, "<>", ${statusRange}, "<>🟣 Published")`),
    "Semua Bulan (Lewat & Blm Publish)",
    APP_CONSTANTS.COLORS.LATE_BG, APP_CONSTANTS.COLORS.LATE_TEXT, APP_CONSTANTS.COLORS.LATE_BORDER
  );

  sheet.setRowHeight(9, 14);

  // -------------------------------------------------------------------------
  // SECTION KIRI: TARGET MINGGUAN & PACING TALENT (B10:G...)
  // -------------------------------------------------------------------------
  sheet.getRange("B10:G10").merge()
    .setValue("👥 TARGET MINGGUAN & PACING TALENT (SENIN-MINGGU)")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.SUBHEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.SUBHEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(10, 28);

  sheet.getRange("B11:G11").setValues([["PIC & Role", "Target", "Terjadwal", "Selesai", "Sisa", "Progress Bar"]])
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontWeight("bold")
    .setBackground("#F1F5F9")
    .setFontColor("#475569")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(11, 24);

  // Sertakan PIC target 0 (misal Jonathan) otomatis kalau punya jadwal di bulan terpilih
  const dbSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.DATABASE);
  const picsWithTasksInPeriod = {};
  if (dbSheet && dbSheet.getLastRow() >= dbStart) {
    const numDbRows = dbSheet.getLastRow() - dbStart + 1;
    const dbQuick = dbSheet.getRange(dbStart, APP_CONSTANTS.DB_COLS.DATE, numDbRows, 4).getValues();
    for (let i = 0; i < dbQuick.length; i++) {
      const dt = parseDateSafe(dbQuick[i][0]);
      const pName = String(dbQuick[i][3] || "").trim();
      if (dt && pName && (dt.getMonth() + 1) === activeMonth && dt.getFullYear() === activeYear) {
        picsWithTasksInPeriod[pName] = true;
      }
    }
  }

  const picStartRow = 12;
  const activePics = config.pics.filter(p => p.target > 0 || picsWithTasksInPeriod[p.name]);
  const picsToDisplay = activePics.length > 0 ? activePics : config.pics;
  const picCount = picsToDisplay.length;
  const wkStartCond = `${dbName}!$C$${dbStart}:$C, ">=" & $K$4`;
  const wkEndCond = `${dbName}!$C$${dbStart}:$C, "<=" & ($K$4+6)`;
  const picRange = `${dbName}!$F$${dbStart}:$F`;

  const picRowsFormulas = [];
  for (let i = 0; i < picCount; i++) {
    const r = picStartRow + i;
    const cfgRow = picsToDisplay[i].row || (7 + i);
    picRowsFormulas.push([
      f(`=${cfgName}!E${cfgRow}`),
      f(`=${cfgName}!G${cfgRow}`),
      f(`=COUNTIFS(${wkStartCond}, ${wkEndCond}, ${picRange}, B${r})`),
      f(`=COUNTIFS(${wkStartCond}, ${wkEndCond}, ${picRange}, B${r}, ${statusRange}, "🟢 Approved") + COUNTIFS(${wkStartCond}, ${wkEndCond}, ${picRange}, B${r}, ${statusRange}, "🟣 Published")`),
      f(`=MAX(0, D${r}-E${r})`),
      f(`=SPARKLINE(E${r}, {"charttype", "bar"; "max", MAX(1, D${r}, E${r}); "color1", "${APP_CONSTANTS.COLORS.ACCENT}"; "color2", "#E2E8F0"})`)
    ]);
    sheet.setRowHeight(r, 32);
  }

  if (picCount > 0) {
    const picTableRange = sheet.getRange(picStartRow, 2, picCount, 6);
    picTableRange.setFormulas(picRowsFormulas)
      .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
      .setFontSize(9)
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(picStartRow, 2, picCount, 1).setFontWeight("bold");
    sheet.getRange(picStartRow, 3, picCount, 4).setHorizontalAlignment("center");
  }

  // Distribusi Brand dinamis membaca daftar brand dari tab Config
  const brandHeaderRow = picStartRow + picCount + 1;
  if (brandHeaderRow - 1 > 19) {
    sheet.setRowHeight(brandHeaderRow - 1, 12);
  }
  sheet.getRange(brandHeaderRow, 2, 1, 6).merge()
    .setValue("⚖️ DISTRIBUSI BRAND (BULAN TERPILIH)")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground("#334155")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(brandHeaderRow, 28);

  const brandRange = `${dbName}!$E$${dbStart}:$E`;
  const brandsList = (config.brands && config.brands.length > 0) ? config.brands : APP_CONSTANTS.DEFAULTS.BRANDS;
  const brandPalette = APP_CONSTANTS.COLORS.BRAND_PALETTE || [APP_CONSTANTS.COLORS.VALMART, APP_CONSTANTS.COLORS.MARTFORUM];
  const brandIcons = ["🟣", "🔵", "🟢", "🟠", "🔴", "🟣"];

  for (let bIdx = 0; bIdx < brandsList.length; bIdx++) {
    const bName = brandsList[bIdx];
    const bRow = brandHeaderRow + 1 + bIdx;
    const bEscaped = bName.replace(/"/g, '""');
    const displayLabel = `${brandIcons[bIdx % brandIcons.length]} ${bName}`;
    sheet.getRange(bRow, 2, 1, 3).merge().setValue(displayLabel).setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setFontWeight("bold");
    sheet.getRange(bRow, 5, 1, 3).merge().setFormula(f(`=COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${brandRange}, "${bEscaped}") & " Konten"`))
      .setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setHorizontalAlignment("center");
    sheet.setRowHeight(bRow, bRow <= 19 ? 34 : 26);
  }

  const bBarRow = brandHeaderRow + 1 + brandsList.length;
  let brandBarFormula = "";
  if (brandsList.length === 1) {
    const b0 = brandsList[0].replace(/"/g, '""');
    const c0 = `COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${brandRange}, "${b0}")`;
    brandBarFormula = f(`=IF(${c0}=0, "", SPARKLINE(${c0}, {"charttype", "bar"; "max", MAX(1, ${c0}); "color1", "${brandPalette[0]}"}))`);
  } else if (brandsList.length === 2) {
    const b0 = brandsList[0].replace(/"/g, '""');
    const b1 = brandsList[1].replace(/"/g, '""');
    const c0 = `COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${brandRange}, "${b0}")`;
    const c1 = `COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${brandRange}, "${b1}")`;
    brandBarFormula = f(`=IF(${c0} + ${c1} = 0, "", SPARKLINE({${c0}, ${c1}}, {"charttype", "bar"; "color1", "${brandPalette[0]}"; "color2", "${brandPalette[1]}"}))`);
  } else {
    const b0 = brandsList[0].replace(/"/g, '""');
    const c0 = `COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${brandRange}, "${b0}")`;
    const cAll = `COUNTIFS(${dateCondStart}, ${dateCondEnd})`;
    brandBarFormula = f(`=IF(${cAll}=0, "", SPARKLINE({${c0}, MAX(0, ${cAll}-${c0})}, {"charttype", "bar"; "color1", "${brandPalette[0]}"; "color2", "#94A3B8"}))`);
  }

  sheet.getRange(bBarRow, 2, 1, 6).merge()
    .setFormula(brandBarFormula)
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY);
  sheet.setRowHeight(bBarRow, bBarRow <= 19 ? 34 : 24);
  sheet.getRange(brandHeaderRow, 2, 2 + brandsList.length, 6)
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // -------------------------------------------------------------------------
  // SECTION KANAN: ANTREAN REVIEW OTOMATIS (I10:M20, helper di V12:Z)
  // -------------------------------------------------------------------------
  sheet.getRange("I10:M10").merge()
    .setValue("🚨 ANTREAN REVIEW - SEMUA BULAN (URUT TERDEKAT)")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.SUBHEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.SUBHEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange("I11:M11").setValues([["Tanggal", "Brand & PIC", "Topik & Status", "Link Draft", "Catatan Review"]])
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontWeight("bold")
    .setBackground("#F1F5F9")
    .setFontColor("#475569")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("V12").setFormula(f(
    `=IFERROR(SORT(FILTER({${dbName}!$C$${dbStart}:$C, ARRAYFORMULA(${dbName}!$E$${dbStart}:$E & CHAR(10) & ${dbName}!$F$${dbStart}:$F), ARRAYFORMULA(${dbName}!$G$${dbStart}:$G & " [" & ${dbName}!$J$${dbStart}:$J & "]"), ${dbName}!$I$${dbStart}:$I, ${dbName}!$K$${dbStart}:$K}, (${dbName}!$J$${dbStart}:$J="🟠 In Review") + (${dbName}!$J$${dbStart}:$J="🔴 Revision")), 1, TRUE), {"-", "-", "✅ Antrean kosong, aman!", "-", "-"})`
  ));

  const reviewDisplayRows = 8; // I12:M19
  const reviewFormulas = [];
  for (let i = 0; i < reviewDisplayRows; i++) {
    const r = 12 + i;
    if (r >= picStartRow + picCount && (r < brandHeaderRow || r > bBarRow)) {
      sheet.setRowHeight(r, 34);
    }
    reviewFormulas.push([
      f(`=IF($V${r}="", "", $V${r})`),
      f(`=IF($W${r}="", "", $W${r})`),
      f(`=IF($X${r}="", "", $X${r})`),
      f(`=IF($V${r}="", "", IF($Y${r}="-", "-", IF(REGEXMATCH(TO_TEXT($Y${r}), "^https?://"), HYPERLINK($Y${r}, "🎬 Buka Draft"), IF(TRIM(TO_TEXT($Y${r}))="", "⚪ Belum ada", $Y${r}))))`),
      f(`=IF($Z${r}="", "", $Z${r})`)
    ]);
  }
  sheet.getRange(12, 9, reviewDisplayRows, 5).setFormulas(reviewFormulas);
  sheet.getRange("I12:I19").setNumberFormat("yyyy-mm-dd").setHorizontalAlignment("center");
  sheet.getRange("L12:L19").setHorizontalAlignment("center");
  sheet.getRange("I12:M19")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(8.5)
    .setWrap(true)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("I20:M20").merge()
    .setFormula(f(`=IF($V20<>"", "⚠️ +" & COUNTA($V20:$V) & " antrean lagi - lihat tab '📝 Database Konten'", "")`))
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(8.5)
    .setFontWeight("bold")
    .setFontColor("#9A3412")
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(20, 22);

  // -------------------------------------------------------------------------
  // SECTION BAWAH: DAFTAR KONTEN TELAT (B22:M..., helper di O..T)
  // -------------------------------------------------------------------------
  const lateHeaderRow = Math.max(bBarRow + 2, 22);
  sheet.setRowHeight(lateHeaderRow - 1, 14);
  sheet.getRange(lateHeaderRow, 2, 1, 12).merge()
    .setValue("⏰ DAFTAR KONTEN TELAT - SEMUA BULAN (LEWAT TANGGAL & BELUM PUBLISHED)")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.LATE_TEXT)
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(lateHeaderRow, 28);

  const lateSubRow = lateHeaderRow + 1;
  sheet.getRange(lateSubRow, 2, 1, 12).setValues([
    ["Tanggal", "Brand", "PIC & Role", "", "Topik & Hook", "", "", "", "Status", "Catatan Review", "", ""]
  ]);
  sheet.getRange(lateSubRow, 4, 1, 2).merge(); // PIC
  sheet.getRange(lateSubRow, 6, 1, 4).merge(); // Topik
  sheet.getRange(lateSubRow, 11, 1, 3).merge(); // Catatan
  sheet.getRange(lateSubRow, 2, 1, 12)
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontWeight("bold")
    .setBackground("#FEF2F2")
    .setFontColor(APP_CONSTANTS.COLORS.LATE_TEXT)
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.LATE_BORDER, SpreadsheetApp.BorderStyle.SOLID);
  sheet.setRowHeight(lateSubRow, 24);

  const lateDataRow = lateSubRow + 1;
  const lateDisplayRows = 15;
  sheet.getRange(lateDataRow, 15).setFormula(f(
    `=IFERROR(SORT(FILTER({${dbName}!$C$${dbStart}:$C, ${dbName}!$E$${dbStart}:$E, ${dbName}!$F$${dbStart}:$F, ${dbName}!$G$${dbStart}:$G, ${dbName}!$J$${dbStart}:$J, ${dbName}!$K$${dbStart}:$K}, ${dbName}!$C$${dbStart}:$C<>"", ${dbName}!$C$${dbStart}:$C<TODAY(), ${dbName}!$J$${dbStart}:$J<>"🟣 Published"), 1, TRUE), {"-", "-", "-", "✅ Tidak ada konten telat!", "-", "-"})`
  ));

  const colBFormulas = [];
  const colCFormulas = [];
  const colDFormulas = [];
  const colFFormulas = [];
  const colJFormulas = [];
  const colKFormulas = [];

  for (let i = 0; i < lateDisplayRows; i++) {
    const r = lateDataRow + i;
    sheet.setRowHeight(r, 24);

    colBFormulas.push([f(`=IF($O${r}="", "", $O${r})`)]);
    colCFormulas.push([f(`=IF($P${r}="", "", $P${r})`)]);
    colDFormulas.push([f(`=IF($Q${r}="", "", $Q${r})`)]);
    colFFormulas.push([f(`=IF($R${r}="", "", $R${r})`)]);
    colJFormulas.push([f(`=IF($S${r}="", "", $S${r})`)]);
    colKFormulas.push([f(`=IF($T${r}="", "", $T${r})`)]);
  }

  sheet.getRange(lateDataRow, 2, lateDisplayRows, 1).setFormulas(colBFormulas).setNumberFormat("yyyy-mm-dd").setHorizontalAlignment("center");
  sheet.getRange(lateDataRow, 3, lateDisplayRows, 1).setFormulas(colCFormulas).setHorizontalAlignment("center");
  sheet.getRange(lateDataRow, 4, lateDisplayRows, 1).setFormulas(colDFormulas);
  sheet.getRange(lateDataRow, 6, lateDisplayRows, 1).setFormulas(colFFormulas);
  sheet.getRange(lateDataRow, 10, lateDisplayRows, 1).setFormulas(colJFormulas).setHorizontalAlignment("center");
  sheet.getRange(lateDataRow, 11, lateDisplayRows, 1).setFormulas(colKFormulas);

  for (let i = 0; i < lateDisplayRows; i++) {
    const r = lateDataRow + i;
    sheet.getRange(r, 4, 1, 2).merge();  // D:E PIC & Role
    sheet.getRange(r, 6, 1, 4).merge();  // F:I Topik & Hook
    sheet.getRange(r, 11, 1, 3).merge(); // K:M Catatan Review
  }

  const lateTableRange = sheet.getRange(lateDataRow, 2, lateDisplayRows, 12);
  lateTableRange
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(8.5)
    .setWrap(true)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#E2E8F0", SpreadsheetApp.BorderStyle.SOLID);

  const lateOverflowRow = lateDataRow + lateDisplayRows;
  sheet.getRange(lateOverflowRow, 2, 1, 12).merge()
    .setFormula(f(`=IF($O${lateOverflowRow}<>"", "⚠️ +" & COUNTA($O${lateOverflowRow}:$O) & " konten telat lainnya belum ditampilkan - buka tab '📝 Database Konten'.", "")`))
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(8.5)
    .setFontWeight("bold")
    .setFontColor(APP_CONSTANTS.COLORS.LATE_TEXT)
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(lateOverflowRow, 22);

  // Highlight merah hanya untuk baris telat yang terisi data (dan hijau jika kosong/aman)
  const cfRules = [
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(f(`=$F${lateDataRow}="✅ Tidak ada konten telat!"`))
      .setBackground("#F0FDF4")
      .setFontColor("#166534")
      .setBold(true)
      .setRanges([lateTableRange])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(f(`=AND($B${lateDataRow}<>"", $B${lateDataRow}<>"-")`))
      .setBackground("#FEF2F2")
      .setFontColor(APP_CONSTANTS.COLORS.LATE_TEXT)
      .setRanges([lateTableRange])
      .build()
  ];
  sheet.setConditionalFormatRules(cfRules);

  // Proteksi sheet (Warning Only), sisakan dropdown filter C4, E4, I4:J4 tetap bebas diubah tanpa warning
  const protection = sheet.protect().setDescription("Proteksi Dashboard (Warning Only)");
  protection.setUnprotectedRanges([
    sheet.getRange("C4"),
    sheet.getRange("E4"),
    sheet.getRange("I4:J4")
  ]);
  protection.setWarningOnly(true);
}

/**
 * Helper membuat KPI Scorecard 3 baris.
 */
function createScoreCard(sheet, labelRange, numRange, tagRange, label, formula, tagFormula, bg, textColor, borderColor) {
  sheet.getRange(labelRange).merge()
    .setValue(label)
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(8.5)
    .setFontWeight("bold")
    .setFontColor(textColor)
    .setBackground(bg)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("bottom");

  sheet.getRange(numRange).merge()
    .setFormula(formula)
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(20)
    .setFontWeight("bold")
    .setFontColor(textColor)
    .setBackground(bg)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  const tagCell = sheet.getRange(tagRange).merge();
  if (tagFormula && tagFormula.startsWith("=")) {
    tagCell.setFormula(tagFormula);
  } else {
    tagCell.setValue(tagFormula || "");
  }
  tagCell
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(8)
    .setFontColor(APP_CONSTANTS.COLORS.MUTED_TEXT)
    .setBackground(bg)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("top");

  const topLeft = sheet.getRange(labelRange);
  const numCols = topLeft.getNumColumns();
  sheet.getRange(topLeft.getRow(), topLeft.getColumn(), 3, numCols)
    .setBorder(true, true, true, true, null, null, borderColor, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}
