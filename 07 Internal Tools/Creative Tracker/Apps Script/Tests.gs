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
    Logger.log(`${condition ? "[PASS]" : "[FAIL]"} ${name}${detail ? " — " + detail : ""}`);
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
  assert("URL Builder: spasi ter-encode dan nama file Title Case benar",
    builtUrl.includes("Brief%20Valmart%20Helmy.md") && !builtUrl.includes("BRIEF_VALMART"),
    builtUrl
  );

  // 4. Test Render Kalender: bulan 28/30/31 hari & bulan mulai hari Minggu (Nov 2026)
  const feb2026 = buildCalendarMatrix(2, 2026, {}, refToday); // 28 hari, mulai Minggu
  const apr2026 = buildCalendarMatrix(4, 2026, {}, refToday); // 30 hari
  const dec2026 = buildCalendarMatrix(12, 2026, {}, refToday); // 31 hari
  const nov2026 = buildCalendarMatrix(11, 2026, {}, refToday); // 1 Nov 2026 = Minggu (kolom indeks 6)

  assert("Kalender: Feb 2026 (28 hari) ter-generate rapi", feb2026.length >= 4 && feb2026[0].headerVals[6] === "1");
  assert("Kalender: Apr 2026 (30 hari) ter-generate rapi", apr2026.length >= 5);
  assert("Kalender: Des 2026 (31 hari) ter-generate rapi", dec2026.length >= 5);
  assert("Kalender: Nov 2026 mulai hari Minggu (kolom ke-7)", nov2026[0].headerVals[6] === "1" && nov2026[0].headerVals[0] === "");

  // 5. Test Integrasi Terisolasi: Setup Idempotent & Sync Upsert 2x (kolom tim tidak tertimpa)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tempName = "__TEST_DB_TEMP__";
  let tempSheet = ss.getSheetByName(tempName);
  if (tempSheet) ss.deleteSheet(tempSheet);
  tempSheet = ss.insertSheet(tempName);

  try {
    formatDatabaseSheet(tempSheet, true);

    // Sync pertama
    const res1 = upsertScheduleToDatabase(tempSheet, parsed.rows, dummyCfg, { checkLinks: false });
    assert("Sync #1: menambah 2 baris baru", res1.added === 2 && res1.updated === 0);

    // Simulasi edit manual tim pada baris 5: ubah Link Draft, Status, dan Catatan
    tempSheet.getRange(5, APP_CONSTANTS.DB_COLS.DRAFT_LINK).setValue("https://drive.google.com/draft-helmy");
    tempSheet.getRange(5, APP_CONSTANTS.DB_COLS.STATUS).setValue("🟠 In Review");
    tempSheet.getRange(5, APP_CONSTANTS.DB_COLS.NOTES).setValue("Sudah oke, tinggal ACC");

    // Sync kedua dengan judul topik diupdate di GitHub
    const updatedRows = JSON.parse(JSON.stringify(parsed.rows));
    updatedRows[0].topik = "Mid-Month ACE Bounty (Revisi Judul)";
    const res2 = upsertScheduleToDatabase(tempSheet, updatedRows, dummyCfg, { checkLinks: false });

    const savedStatus = tempSheet.getRange(5, APP_CONSTANTS.DB_COLS.STATUS).getValue();
    const savedDraft = tempSheet.getRange(5, APP_CONSTANTS.DB_COLS.DRAFT_LINK).getValue();
    const savedNotes = tempSheet.getRange(5, APP_CONSTANTS.DB_COLS.NOTES).getValue();
    const savedTopic = tempSheet.getRange(5, APP_CONSTANTS.DB_COLS.TOPIC).getValue();

    assert("Sync #2: tidak ada duplikat (1 update, 1 unchanged)", res2.added === 0 && res2.updated === 1 && res2.unchanged === 1);
    assert("Sync #2: Status tim tetap utuh", savedStatus === "🟠 In Review", savedStatus);
    assert("Sync #2: Link Draft tim tetap utuh", savedDraft === "https://drive.google.com/draft-helmy", savedDraft);
    assert("Sync #2: Catatan Review tim tetap utuh", savedNotes === "Sudah oke, tinggal ACC", savedNotes);
    assert("Sync #2: Topik dari GitHub terupdate", savedTopic === "Mid-Month ACE Bounty (Revisi Judul)", savedTopic);
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
