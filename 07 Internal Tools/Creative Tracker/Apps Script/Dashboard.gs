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
  const activeMonth = (prevMonth >= 1 && prevMonth <= 12) ? Number(prevMonth) : (now.getMonth() + 1);
  const activeYear = (prevYear >= 2024 && prevYear <= 2035) ? Number(prevYear) : now.getFullYear();
  const activeWeekOpt = prevWeekOpt ? String(prevWeekOpt) : "Minggu Ini";

  const config = getAppConfig();
  const isIndo = isIndonesianLocale(ss);
  const f = (formulaStr) => localizeFormula(formulaStr, isIndo);
  const dbName = `'${APP_CONSTANTS.SHEETS.DATABASE}'`;
  const cfgName = `'${APP_CONSTANTS.SHEETS.CONFIG}'`;
  const maxRow = APP_CONSTANTS.DB_START_ROW + APP_CONSTANTS.MAX_DB_ROWS - 1; // 1004

  // Lepas proteksi lama sebelum clear, lalu pasang lagi di akhir
  sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(p => p.remove());
  sheet.clear();
  sheet.setTabColor("#10B981");

  // Lebar kolom A..M
  const colWidths = [
    [1, 20],  // A Spacer
    [2, 175], // B
    [3, 90],  // C
    [4, 90],  // D
    [5, 90],  // E
    [6, 90],  // F
    [7, 150], // G
    [8, 22],  // H Spacer
    [9, 105], // I
    [10, 100],// J
    [11, 130],// K
    [12, 230],// L
    [13, 180] // M
  ];
  colWidths.forEach(([c, w]) => sheet.setColumnWidth(c, w));

  // Header Banner
  sheet.getRange("B2:M2").merge()
    .setValue("📊 CREATIVE DASHBOARD — VALMART.ID & MARTFORUM.ID")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(13)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.HEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.HEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet.setRowHeight(2, 38);

  sheet.getRange("B3:M3").merge()
    .setValue("💡 Pilih Bulan/Tahun & Minggu di baris 4 untuk memfilter metrik. Semua angka terhubung langsung ke '📝 Database Konten'.")
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
    SpreadsheetApp.newDataValidation().requireValueInList(["2025","2026","2027","2028"], true).setAllowInvalid(true).build()
  );
  sheet.getRange("I4").setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(["Minggu Lalu", "Minggu Ini", "Minggu Depan"], true).setAllowInvalid(false).build()
  );

  // -------------------------------------------------------------------------
  // 5 KPI SCORECARDS (Baris 6 - 8)
  // -------------------------------------------------------------------------
  sheet.setRowHeight(5, 10);
  sheet.setRowHeight(6, 22);
  sheet.setRowHeight(7, 38);
  sheet.setRowHeight(8, 20);

  const dateCondStart = `${dbName}!$C$5:$C$${maxRow}, ">=" & DATE($E$4, $C$4, 1)`;
  const dateCondEnd = `${dbName}!$C$5:$C$${maxRow}, "<=" & EOMONTH(DATE($E$4, $C$4, 1), 0)`;
  const statusRange = `${dbName}!$J$5:$J$${maxRow}`;

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
    f(`=COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${statusRange}, "🟠 In Review") + COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${statusRange}, "🔴 Revision")`),
    "Perlu Dicek Ko Jeje/Hildan",
    "#FFF7ED", "#C2410C", "#FDBA74"
  );

  createScoreCard(sheet, "I6:K6", "I7:K7", "I8:K8",
    "⏳ DALAM PENGERJAAN",
    f(`=COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${statusRange}, "⚪ Draft") + COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${statusRange}, "🟡 In Progress")`),
    "On Track Talent",
    "#FEFCE8", "#A16207", "#FDE047"
  );

  createScoreCard(sheet, "L6:M6", "L7:M7", "L8:M8",
    "🚨 KONTEN TELAT",
    f(`=COUNTIFS(${dbName}!$C$5:$C$${maxRow}, "<" & TODAY(), ${dbName}!$C$5:$C$${maxRow}, "<>", ${statusRange}, "<>🟣 Published")`),
    "Lewat Tanggal & Belum Publish",
    APP_CONSTANTS.COLORS.LATE_BG, APP_CONSTANTS.COLORS.LATE_TEXT, APP_CONSTANTS.COLORS.LATE_BORDER
  );

  sheet.setRowHeight(9, 14);

  // -------------------------------------------------------------------------
  // SECTION KIRI: TARGET MINGGUAN & PACING TALENT (B10:G...)
  // -------------------------------------------------------------------------
  sheet.getRange("B10:G10").merge()
    .setValue("👥 TARGET MINGGUAN & PACING TALENT (SENIN–MINGGU)")
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

  const picStartRow = 12;
  const picCount = config.pics.length;
  const wkStartCond = `${dbName}!$C$5:$C$${maxRow}, ">=" & $K$4`;
  const wkEndCond = `${dbName}!$C$5:$C$${maxRow}, "<=" & ($K$4+6)`;
  const picRange = `${dbName}!$F$5:$F$${maxRow}`;

  const picRowsFormulas = [];
  for (let i = 0; i < picCount; i++) {
    const r = picStartRow + i;
    const cfgRow = 7 + i; // Posisi baris di tab Config E7:G20
    picRowsFormulas.push([
      f(`=${cfgName}!E${cfgRow}`),
      f(`=${cfgName}!G${cfgRow}`),
      f(`=COUNTIFS(${wkStartCond}, ${wkEndCond}, ${picRange}, B${r})`),
      f(`=COUNTIFS(${wkStartCond}, ${wkEndCond}, ${picRange}, B${r}, ${statusRange}, "🟢 Approved") + COUNTIFS(${wkStartCond}, ${wkEndCond}, ${picRange}, B${r}, ${statusRange}, "🟣 Published")`),
      f(`=MAX(0, C${r}-D${r})`),
      f(`=SPARKLINE(D${r}, {"charttype", "bar"; "max", MAX(1, C${r}); "color1", "${APP_CONSTANTS.COLORS.ACCENT}"; "color2", "#E2E8F0"})`)
    ]);
    sheet.setRowHeight(r, 26);
  }

  if (picCount > 0) {
    const picTableRange = sheet.getRange(picStartRow, 2, picCount, 6);
    picTableRange.setFormulas(picRowsFormulas)
      .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
      .setFontSize(9)
      .setVerticalAlignment("middle")
      .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(picStartRow, 2, picCount, 1).setFontWeight("bold");
    sheet.getRange(picStartRow, 3, picCount, 3).setHorizontalAlignment("center");
  }

  // Distribusi Brand (di bawah tabel Talent)
  const brandHeaderRow = picStartRow + picCount + 1;
  sheet.setRowHeight(brandHeaderRow - 1, 12);
  sheet.getRange(brandHeaderRow, 2, 1, 6).merge()
    .setValue("⚖️ DISTRIBUSI BRAND (BULAN TERPILIH)")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground("#334155")
    .setFontColor("#FFFFFF")
    .setHorizontalAlignment("center");
  sheet.setRowHeight(brandHeaderRow, 26);

  const brandRange = `${dbName}!$E$5:$E$${maxRow}`;
  const bRow1 = brandHeaderRow + 1;
  const bRow2 = brandHeaderRow + 2;
  const bBarRow = brandHeaderRow + 3;

  sheet.getRange(bRow1, 2, 1, 3).merge().setValue("🟣 Valmart.id").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setFontWeight("bold");
  sheet.getRange(bRow1, 5, 1, 3).merge().setFormula(f(`=COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${brandRange}, "Valmart") & " Konten"`)).setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setHorizontalAlignment("center");
  sheet.setRowHeight(bRow1, 24);

  sheet.getRange(bRow2, 2, 1, 3).merge().setValue("🔵 Martforum.id").setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setFontWeight("bold");
  sheet.getRange(bRow2, 5, 1, 3).merge().setFormula(f(`=COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${brandRange}, "Martforum") & " Konten"`)).setFontFamily(APP_CONSTANTS.FONT_FAMILY).setFontSize(9.5).setHorizontalAlignment("center");
  sheet.setRowHeight(bRow2, 24);

  sheet.getRange(bBarRow, 2, 1, 6).merge()
    .setFormula(f(`=SPARKLINE({COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${brandRange}, "Valmart"), COUNTIFS(${dateCondStart}, ${dateCondEnd}, ${brandRange}, "Martforum")}, {"charttype", "bar"; "color1", "${APP_CONSTANTS.COLORS.VALMART}"; "color2", "${APP_CONSTANTS.COLORS.MARTFORUM}"})`))
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY);
  sheet.setRowHeight(bBarRow, 22);
  sheet.getRange(brandHeaderRow, 2, 4, 6).setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // -------------------------------------------------------------------------
  // SECTION KANAN: ANTREAN REVIEW OTOMATIS (I10:M18)
  // -------------------------------------------------------------------------
  sheet.getRange("I10:M10").merge()
    .setValue("🚨 ANTREAN REVIEW (URUT TANGGAL TERDEKAT)")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(10)
    .setFontWeight("bold")
    .setBackground(APP_CONSTANTS.COLORS.SUBHEADER_BG)
    .setFontColor(APP_CONSTANTS.COLORS.SUBHEADER_TEXT)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  sheet.getRange("I11:M11").setValues([["Tanggal", "Brand", "PIC", "Topik & Status", "Catatan Review"]])
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(9)
    .setFontWeight("bold")
    .setBackground("#F1F5F9")
    .setFontColor("#475569")
    .setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  sheet.getRange("I12").setFormula(f(
    `=IFERROR(SORT(FILTER({${dbName}!C5:C${maxRow}, ${dbName}!E5:E${maxRow}, ${dbName}!F5:F${maxRow}, ARRAYFORMULA(${dbName}!G5:G${maxRow} & " [" & ${dbName}!J5:J${maxRow} & "]"), ${dbName}!K5:K${maxRow}}, (${dbName}!J5:J${maxRow}="🟠 In Review") + (${dbName}!J5:J${maxRow}="🔴 Revision")), 1, TRUE), {"-", "-", "-", "✅ Antrean kosong, aman!", "-"})`
  ));
  sheet.getRange("I12:I19").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("I12:M19")
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(8.5)
    .setWrap(true)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // -------------------------------------------------------------------------
  // SECTION BAWAH: DAFTAR KONTEN TELAT (B22:M30)
  // -------------------------------------------------------------------------
  const lateHeaderRow = Math.max(bBarRow + 2, 21);
  sheet.setRowHeight(lateHeaderRow - 1, 14);
  sheet.getRange(lateHeaderRow, 2, 1, 12).merge()
    .setValue("⏰ DAFTAR KONTEN TELAT (LEWAT TANGGAL TAYANG & BELUM PUBLISHED)")
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

  const lateDataRow = lateSubRow + 1;
  sheet.getRange(lateDataRow, 2).setFormula(f(
    `=IFERROR(SORT(FILTER({${dbName}!C5:C${maxRow}, ${dbName}!E5:E${maxRow}, ${dbName}!F5:F${maxRow}, ${dbName}!G5:G${maxRow}, ${dbName}!J5:J${maxRow}, ${dbName}!K5:K${maxRow}}, ${dbName}!C5:C${maxRow}<>"", ${dbName}!C5:C${maxRow}<TODAY(), ${dbName}!J5:J${maxRow}<>"🟣 Published"), 1, TRUE), {"-", "-", "-", "✅ Tidak ada konten telat!", "-", "-"})`
  ));
  // Catatan: FILTER mengembalikan 6 kolom berurutan B..G di baris lateDataRow
  sheet.getRange(lateDataRow, 2, 8, 1).setNumberFormat("yyyy-mm-dd");
  sheet.getRange(lateDataRow, 2, 8, 6)
    .setFontFamily(APP_CONSTANTS.FONT_FAMILY)
    .setFontSize(8.5)
    .setWrap(true)
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, APP_CONSTANTS.COLORS.LATE_BORDER, SpreadsheetApp.BorderStyle.SOLID);

  // Proteksi sheet (Warning Only) agar formula tidak terhapus tidak sengaja
  const protection = sheet.protect().setDescription("Proteksi Dashboard (Warning Only)");
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
