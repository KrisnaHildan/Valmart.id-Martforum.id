/**
 * =========================================================================
 * TESTS.GS — Automated Unit & Integration Tests (testRunAll)
 * =========================================================================
 */

/**
 * Menjalankan seluruh rangkaian pengujian dan menampilkan ringkasan PASS/FAIL.
 * Jalankan dari Apps Script Editor: pilih testRunAll -> Run.
 */
function testRunAll() {
  const results = [];
  const assert = (name, condition, detail) => {
    results.push({ name, pass: Boolean(condition), detail: detail || "" });
    Logger.log(`${condition ? "[PASS]" : "[FAIL]"} ${name}${detail ? " - " + detail : ""}`);
  };

  // 1. Test Parser Markdown Toleran
  const sampleMd = [
    "# Jadwal Konten",
    "",
    "| ID | Tanggal | Brand | PIC | Topik & Hook | File Naskah |",
    "| :--- | :--- | :--- | :--- | :--- | :--- |",
    "| VAL-20261116-01 | 2026-11-16 | Valmart | Helmy | Mid-Month ACE Bounty | November 2026/Week 3 (16 - 22 Nov 2026)/Brief Valmart Helmy.md |",
    "MART-20261116-02 | 2026-11-16 | Martforum | Audrey | Teaser QRIS 3 Detik | November 2026/Week 3 (16 - 22 Nov 2026)/Brief Martforum Jonathan.md",
    "| BAD-ROW | tanggal-ngaco | Valmart | Helmy | Harus diskip | file.md |",
    ""
  ].join("\n");

  const parsed = parseMarkdownSchedule(sampleMd);
  assert("Parser Markdown: jumlah baris valid = 2", parsed.rows.length === 2, `Dapat ${parsed.rows.length}`);
  assert("Parser Markdown: baris cacat menghasilkan warning", parsed.warnings.length === 1, `Warnings ${parsed.warnings.length}`);
  assert("Parser Markdown: toleran tanpa pipe ujung", parsed.rows[1] && parsed.rows[1].id === "MART-20261116-02");

  // 2. Test Deteksi Telat
  const refToday = new Date(2026, 10, 25); // 25 Nov 2026
  assert("Deteksi Telat: tanggal kemarin + Draft = TELAT", isTaskLate("2026-11-24", "⚪ Draft", refToday) === true);
  assert("Deteksi Telat: tanggal kemarin + Published = AMAN", isTaskLate("2026-11-24", "🟣 Published", refToday) === false);
  assert("Deteksi Telat: hari ini + Draft = AMAN", isTaskLate("2026-11-25", "⚪ Draft", refToday) === false);

  const fEn = '=SPARKLINE({COUNTIFS(A, B), COUNTIFS(C, D)}, {"charttype", "bar"; "max", MAX(1, C12); "label", "Aman, oke"})';
  const fId = localizeFormula(fEn, true);
  assert("Locale Formula: konversi koma ke ; dan \\ termasuk nested () di dalam {}",
    fId === '=SPARKLINE({COUNTIFS(A; B)\\ COUNTIFS(C; D)}; {"charttype"\\ "bar"; "max"\\ MAX(1; C12); "label"\\ "Aman, oke"})',
    fId
  );

  const hlId = makeHyperlinkFormula("https://example.com/a", "⚠️ Link 404 (Cek Repo)", true);
  assert("Hyperlink Locale ID: pakai titik koma dan ekstrak label benar",
    hlId === '=HYPERLINK("https://example.com/a"; "⚠️ Link 404 (Cek Repo)")' &&
    extractLabelFromCell("", hlId) === "⚠️ Link 404 (Cek Repo)",
    hlId
  );

  const jonathanCfg = APP_CONSTANTS.DEFAULTS.PICS.find(p => p.name.includes("Jonathan"));
  const dikaCfg = APP_CONSTANTS.DEFAULTS.PICS.find(p => p.name.includes("Dika"));
  assert("Config PIC: Jonathan cuti (target 0) & Dika aktif menggantikan (target > 0)",
    jonathanCfg && jonathanCfg.target === 0 && dikaCfg && dikaCfg.target > 0
  );

  // 3. Test URL Builder & encodeURI untuk Title Case + spasi
  const dummyCfg = {
    repoOwner: "KrisnaHildan",
    repoName: "Valmart-x-Martforum",
    branch: "main",
    basePath: "04 Strategi Konten dan Kreatif",
    brands: APP_CONSTANTS.DEFAULTS.BRANDS,
    pics: APP_CONSTANTS.DEFAULTS.PICS
  };
  const builtUrl = buildGithubBlobUrl(dummyCfg, "November 2026/Week 4 (23 - 30 Nov 2026)/Brief Valmart Helmy.md");
  const winBackslashUrl = buildGithubBlobUrl(dummyCfg, "November 2026\\Week 4 (23 - 30 Nov 2026)\\Brief Valmart Helmy.md");
  assert("URL Builder: spasi ter-encode, backslash dinormalisasi, dan nama file Title Case benar",
    builtUrl.includes("Brief%20Valmart%20Helmy.md") && !builtUrl.includes("BRIEF_VALMART") && winBackslashUrl === builtUrl,
    builtUrl
  );

  // 3b. Test Parser toleran terhadap formatting markdown (`code`, **bold**), backslash, & escaped pipe
  const formattedMd = [
    "| ID | Tanggal | Brand | PIC | Topik & Hook | File Naskah |",
    "| :--- | :--- | :--- | :--- | :--- | :--- |",
    "| `VAL-20261201-01` | **2026-12-01** | **Valmart** | `Helmy` | Tes Format \\| Pipe | `Desember 2026\\Week 1\\Brief.md` |"
  ].join("\n");
  const parsedFormatted = parseMarkdownSchedule(formattedMd);
  assert("Parser Markdown: strip backtick/bold, normalisasi backslash, & escaped pipe",
    parsedFormatted.rows.length === 1 &&
    parsedFormatted.rows[0].id === "VAL-20261201-01" &&
    parsedFormatted.rows[0].tanggal === "2026-12-01" &&
    parsedFormatted.rows[0].topik === "Tes Format | Pipe" &&
    parsedFormatted.rows[0].fileNaskah === "Desember 2026/Week 1/Brief.md"
  );

  // 3c. Test normalisasi PIC kata utuh
  const picTestCfg = [
    { name: "Tim A", target: 1 },
    { name: "Jonathan (Editor & Desain)", target: 0 }
  ];
  assert("Normalisasi PIC: tidak salah cocok substring parsial",
    normalizePic("Fatimah", picTestCfg) === "Fatimah" &&
    normalizePic("Jonathan", picTestCfg) === "Jonathan (Editor & Desain)"
  );

  // 4. Test Render Kalender: bulan 28/30/31 hari, bulan mulai Minggu, & penanda TELAT
  const feb2026 = buildCalendarMatrix(2, 2026, {}, refToday); // 28 hari, mulai Minggu
  const apr2026 = buildCalendarMatrix(4, 2026, {}, refToday); // 30 hari
  const dec2026 = buildCalendarMatrix(12, 2026, {}, refToday); // 31 hari
  const nov2026 = buildCalendarMatrix(11, 2026, {
    24: [{ brand: "Valmart", pic: "Helmy", topik: "Konten Telat", status: "⚪ Draft" }],
    25: [{ brand: "Valmart", pic: "Helmy", topik: "Konten Hari Ini", status: "⚪ Draft" }]
  }, refToday); // 1 Nov 2026 = Minggu (kolom indeks 6)

  assert("Kalender: Feb 2026 (28 hari) ter-generate rapi", feb2026.length >= 4 && feb2026[0].headerVals[6] === "1");
  assert("Kalender: Apr 2026 (30 hari) ter-generate rapi", apr2026.length >= 5);
  assert("Kalender: Des 2026 (31 hari) ter-generate rapi", dec2026.length >= 5);
  assert("Kalender: Nov 2026 mulai hari Minggu (kolom ke-7)", nov2026[0].headerVals[6] === "1" && nov2026[0].headerVals[0] === "");
  const novTextFlat = nov2026.map(w => w.contentVals.join(" ")).join(" ");
  assert("Kalender: menandai item telat (24 Nov) sebagai TELAT", novTextFlat.includes("🚨 TELAT"));

  // 5. Test Integrasi Terisolasi: Setup Baru (row 6), Sync Upsert 2x, Chronological Sort, & Auto-Migrasi
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempName = "__TEST_DB_TEMP__";
  let tempSheet = ss.getSheetByName(tempName);
  if (tempSheet) ss.deleteSheet(tempSheet);
  tempSheet = ss.insertSheet(tempName);

  try {
    formatDatabaseSheet(tempSheet, true);
    const startRow = APP_CONSTANTS.DB_START_ROW; // 6

    // Sync pertama (masukkan ID duplikat & tanggal mundur untuk uji dedup + chronological sort)
    const rowsWithOutOrder = [
      parsed.rows[0], // 2026-11-16
      parsed.rows[1], // 2026-11-16
      { id: "VAL-20261116-01", tanggal: "2026-11-16", brand: "Valmart", pic: "Helmy", topik: "Mid-Month ACE Bounty", fileNaskah: parsed.rows[0].fileNaskah },
      { id: "VAL-20261102-01", tanggal: "2026-11-02", brand: "Valmart", pic: "Helmy", topik: "Harus Urut Atas", fileNaskah: "file.md" }
    ];
    const res1 = upsertScheduleToDatabase(tempSheet, rowsWithOutOrder, dummyCfg, { checkLinks: false, skipFilterRefresh: true });
    const firstIdSorted = tempSheet.getRange(startRow, APP_CONSTANTS.DB_COLS.ID).getValue();
    assert("Sync #1: dedup ID & sort kronologis berhasil (3 baris unik, terawal 2 Nov)",
      res1.added === 3 && firstIdSorted === "VAL-20261102-01",
      `added=${res1.added}, first=${firstIdSorted}`
    );

    // Simulasi edit manual tim pada baris ke-2 data (startRow + 1 = VAL-20261116-01)
    const rowMid = startRow + 1;
    const rowMart = startRow + 2;
    tempSheet.getRange(rowMid, APP_CONSTANTS.DB_COLS.DRAFT_LINK).setValue("https://drive.google.com/draft-helmy");
    tempSheet.getRange(rowMid, APP_CONSTANTS.DB_COLS.STATUS).setValue("🟠 In Review");
    tempSheet.getRange(rowMid, APP_CONSTANTS.DB_COLS.NOTES).setValue("Sudah oke, tinggal ACC");

    // Simulasi link 404 pada MART + tag [Tidak ada di GitHub] pada rowMid
    const row2Url = buildGithubBlobUrl(dummyCfg, parsed.rows[1].fileNaskah);
    const brokenFormula = makeHyperlinkFormula(row2Url, "⚠️ Link 404 (Cek Repo)", isIndonesianLocale(ss));
    tempSheet.getRange(rowMart, APP_CONSTANTS.DB_COLS.SCRIPT_LINK).setFormula(brokenFormula);
    tempSheet.getRange(rowMart, APP_CONSTANTS.DB_COLS.NOTES).setValue("[Tidak ada di GitHub] Catatan lama");

    const c6RawBeforeSync2 = tempSheet.getRange(startRow, APP_CONSTANTS.DB_COLS.DATE).getValue();
    const diagC6 = `type=${typeof c6RawBeforeSync2}, proto=${Object.prototype.toString.call(c6RawBeforeSync2)}, parsed=${formatDateIso(parseDateSafe(c6RawBeforeSync2))}, str=${String(c6RawBeforeSync2)}`;

    // Sync kedua (tanpa VAL-20261102-01 supaya teruji flag missing, topik rowMid berubah, link MART pulih)
    const updatedRows = JSON.parse(JSON.stringify(parsed.rows));
    updatedRows[0].topik = "Mid-Month ACE Bounty (Revisi Judul)";
    const res2 = upsertScheduleToDatabase(tempSheet, updatedRows, dummyCfg, { checkLinks: false, skipFilterRefresh: true });

    const savedStatus = tempSheet.getRange(rowMid, APP_CONSTANTS.DB_COLS.STATUS).getValue();
    const savedDraft = tempSheet.getRange(rowMid, APP_CONSTANTS.DB_COLS.DRAFT_LINK).getValue();
    const savedNotes = tempSheet.getRange(rowMid, APP_CONSTANTS.DB_COLS.NOTES).getValue();
    const savedTopic = tempSheet.getRange(rowMid, APP_CONSTANTS.DB_COLS.TOPIC).getValue();
    const recoveredLabel = tempSheet.getRange(rowMart, APP_CONSTANTS.DB_COLS.SCRIPT_LINK).getValue();
    const recoveredNote = tempSheet.getRange(rowMart, APP_CONSTANTS.DB_COLS.NOTES).getValue();
    const missingFlagged = tempSheet.getRange(startRow, APP_CONSTANTS.DB_COLS.NOTES).getValue();

    assert("Sync #2: tidak ada duplikat (0 add, 2 update, 1 missing)", res2.added === 0 && res2.updated === 2 && res2.missing === 1, diagC6);
    assert("Sync #2: Status tim tetap utuh", savedStatus === "🟠 In Review", `${savedStatus} | ${diagC6}`);
    assert("Sync #2: Link Draft tim tetap utuh", savedDraft === "https://drive.google.com/draft-helmy", `${savedDraft} | ${diagC6}`);
    assert("Sync #2: Catatan Review tim tetap utuh", savedNotes === "Sudah oke, tinggal ACC", `${savedNotes} | ${diagC6}`);
    assert("Sync #2: Topik dari GitHub terupdate", savedTopic === "Mid-Month ACE Bounty (Revisi Judul)", `${savedTopic} | ${diagC6}`);
    assert("Sync #2: Label Link 404 pulih & tag missing dibersihkan", recoveredLabel === "📖 Buka Naskah" && recoveredNote === "Catatan lama", `${recoveredNote} | ${diagC6}`);
    assert("Sync #2: Baris hilang dari GitHub diberi tanda non-destruktif", String(missingFlagged).includes("[Tidak ada di GitHub]"), `${missingFlagged} | ${diagC6}`);

    // Uji Filter Bulan/Tahun Database Konten
    tempSheet.getRange("D4").setValue(11);
    tempSheet.getRange("F4").setValue(2026);
    applyDatabaseMonthYearFilter(tempSheet);
    assert("Filter Database: kriteria filter tanggal aktif terpasang untuk 11/2026",
      Boolean(tempSheet.getFilter() && tempSheet.getFilter().getColumnFilterCriteria(APP_CONSTANTS.DB_COLS.DATE))
    );
    tempSheet.getRange("D4").setValue("Semua");
    tempSheet.getRange("F4").setValue("Semua");
    applyDatabaseMonthYearFilter(tempSheet);
    assert("Filter Database: memilih 'Semua' membersihkan kriteria filter tanggal",
      Boolean(tempSheet.getFilter() && !tempSheet.getFilter().getColumnFilterCriteria(APP_CONSTANTS.DB_COLS.DATE))
    );

    // Uji kapasitas dinamis > 1000 baris
    const capLarge = getDatabaseCapacityRows(tempSheet, 1250);
    assert("Kapasitas Dinamis: sanggup melebihi 1.000 baris (1250)", capLarge >= 1250, String(capLarge));

    // Uji Auto-Migrasi dari layout lama (header baris 4, data baris 5) ke layout baru (header baris 5, data baris 6)
    const migName = "__TEST_MIG_TEMP__";
    let migSheet = ss.getSheetByName(migName);
    if (migSheet) ss.deleteSheet(migSheet);
    migSheet = ss.insertSheet(migName);
    try {
      migSheet.getRange(4, 2, 1, 3).setValues([["ID", "Tanggal Tayang", "Hari"]]);
      migSheet.getRange(5, 2, 1, 3).setValues([["OLD-001", new Date(2026, 9, 20), "Selasa"]]);
      formatDatabaseSheet(migSheet, false);
      const migratedHeader = migSheet.getRange(5, 2).getValue();
      const migratedDataId = migSheet.getRange(6, 2).getValue();
      assert("Auto-Migrasi: header geser ke baris 5 dan data lama utuh di baris 6",
        migratedHeader === "ID" && migratedDataId === "OLD-001",
        `header=${migratedHeader}, data=${migratedDataId}`
      );
    } finally {
      ss.deleteSheet(migSheet);
    }
  } finally {
    ss.deleteSheet(tempSheet);
  }

  const passed = results.filter(r => r.pass).length;
  const failed = results.length - passed;
  const summary = `Total: ${results.length} | PASS: ${passed} | FAIL: ${failed}`;
  Logger.log("=== RINGKASAN TEST ===\n" + summary);
  showAlertSafe("🧪 Hasil testRunAll()", summary + "\n\nCek Execution Log untuk detail tiap kasus uji.");
  return { total: results.length, passed, failed, results };
}
