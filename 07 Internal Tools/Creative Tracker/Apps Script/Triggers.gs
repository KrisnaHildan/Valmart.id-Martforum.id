/**
 * =========================================================================
 * TRIGGERS.GS — Auto-Refresh onEdit, Trigger Harian, & Ringkasan Email
 * =========================================================================
 */

/**
 * Simple Trigger onEdit: auto-fill Hari, Update Terakhir, & auto-refresh Kalender.
 */
function onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // 1. Jika user mengganti Bulan (C2) atau Tahun (E2) di Kalender Visual
  if (sheetName === APP_CONSTANTS.SHEETS.CALENDAR) {
    const a1 = e.range.getA1Notation();
    if (a1 === "C2" || a1 === "E2") {
      renderCalendar();
      showToastSafe("Kalender diperbarui.", "📅 Kalender", 2);
    }
    return;
  }

  // 2. Jika user mengedit Database Konten
  if (sheetName === APP_CONSTANTS.SHEETS.DATABASE) {
    const startRow = e.range.getRow();
    const numRows = e.range.getNumRows();
    const startCol = e.range.getColumn();
    const endCol = startCol + e.range.getNumColumns() - 1;

    if (startRow < APP_CONSTANTS.DB_START_ROW) return;

    // Jika Tanggal Tayang (Col C) diedit -> isi otomatis Hari (Col D)
    if (startCol <= APP_CONSTANTS.DB_COLS.DATE && endCol >= APP_CONSTANTS.DB_COLS.DATE) {
      const dateVals = sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.DATE, numRows, 1).getValues();
      const dayOut = dateVals.map(r => {
        const dt = parseDateSafe(r[0]);
        return [dt ? getIndonesianDayName(dt) : ""];
      });
      sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.DAY, numRows, 1).setValues(dayOut);
    }

    // Jika Status (Col J) atau Catatan (Col K) diedit -> isi Update Terakhir (Col L)
    const editedStatusOrNotes =
      (startCol <= APP_CONSTANTS.DB_COLS.STATUS && endCol >= APP_CONSTANTS.DB_COLS.STATUS) ||
      (startCol <= APP_CONSTANTS.DB_COLS.NOTES && endCol >= APP_CONSTANTS.DB_COLS.NOTES);

    if (editedStatusOrNotes) {
      const now = new Date();
      const stamps = Array.from({ length: numRows }, () => [now]);
      sheet.getRange(startRow, APP_CONSTANTS.DB_COLS.UPDATED_AT, numRows, 1)
        .setValues(stamps)
        .setNumberFormat("yyyy-mm-dd hh:mm");
    }

    // Refresh kalender jika kolom yang tampil di kalender berubah
    const affectsCalendar =
      (startCol <= APP_CONSTANTS.DB_COLS.PIC && endCol >= APP_CONSTANTS.DB_COLS.DATE) ||
      (startCol <= APP_CONSTANTS.DB_COLS.TOPIC && endCol >= APP_CONSTANTS.DB_COLS.TOPIC) ||
      (startCol <= APP_CONSTANTS.DB_COLS.STATUS && endCol >= APP_CONSTANTS.DB_COLS.STATUS);

    if (affectsCalendar) {
      renderCalendar();
    }
  }
}

/**
 * Trigger harian jam 07:00 WIB untuk menyegarkan Kalender & Dashboard.
 */
function scheduledDailyRefresh() {
  renderCalendar();
  renderDashboard();
}

/**
 * Trigger harian jam 08:00 WIB untuk mengirim email ringkasan jika aktif & ada item penting.
 */
function scheduledDailyEmailDigest() {
  const config = getAppConfig();
  if (!config.emailEnabled || config.emailRecipients.length === 0) {
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dbSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.DATABASE);
  if (!dbSheet || dbSheet.getLastRow() < APP_CONSTANTS.DB_START_ROW) return;

  const numRows = dbSheet.getLastRow() - APP_CONSTANTS.DB_START_ROW + 1;
  const rows = dbSheet.getRange(APP_CONSTANTS.DB_START_ROW, APP_CONSTANTS.DB_COLS.ID, numRows, 11).getValues();

  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const tomorrow = new Date(today.getTime() + 86400000);

  const reviewItems = [];
  const lateItems = [];
  const todayItems = [];
  const tomorrowItems = [];

  for (let i = 0; i < rows.length; i++) {
    const dt = parseDateSafe(rows[i][1]);
    if (!dt) continue;
    const brand = rows[i][3] || "-";
    const pic = rows[i][4] || "-";
    const topik = rows[i][5] || "-";
    const status = String(rows[i][8] || "").trim();
    const label = `[${brand}] ${topik} (${pic}) — ${status}`;

    if (status === "🟠 In Review" || status === "🔴 Revision") {
      reviewItems.push(`${formatDateIso(dt)} | ${label}`);
    }
    if (isTaskLate(dt, status, today)) {
      lateItems.push(`${formatDateIso(dt)} | ${label}`);
    }
    if (dt.getTime() === today.getTime()) {
      todayItems.push(label);
    } else if (dt.getTime() === tomorrow.getTime()) {
      tomorrowItems.push(label);
    }
  }

  // Kirim hanya kalau ada hal yang perlu perhatian
  if (reviewItems.length === 0 && lateItems.length === 0 && todayItems.length === 0 && tomorrowItems.length === 0) {
    return;
  }

  const bodyLines = [
    `Halo Ko Jeremy & Hildan,`,
    ``,
    `Berikut ringkasan harian Valmart & Martforum Creative Tracker (${formatDateIso(today)}):`,
    ``,
    `⚠️ Antrean Review / Revisi: ${reviewItems.length}`,
    ...(reviewItems.length > 0 ? reviewItems.map(x => `  - ${x}`) : ["  (Kosong)"]),
    ``,
    `🚨 Konten Telat: ${lateItems.length}`,
    ...(lateItems.length > 0 ? lateItems.map(x => `  - ${x}`) : ["  (Tidak ada)"]),
    ``,
    `📅 Tayang Hari Ini (${todayItems.length}):`,
    ...(todayItems.length > 0 ? todayItems.map(x => `  - ${x}`) : ["  (Tidak ada jadwal)"]),
    ``,
    `📆 Tayang Besok (${tomorrowItems.length}):`,
    ...(tomorrowItems.length > 0 ? tomorrowItems.map(x => `  - ${x}`) : ["  (Tidak ada jadwal)"]),
    ``,
    `Buka Tracker: ${ss.getUrl()}`
  ];

  MailApp.sendEmail({
    to: config.emailRecipients.join(","),
    subject: `[Creative Tracker] ${reviewItems.length} Review, ${lateItems.length} Telat — ${formatDateIso(today)}`,
    body: bodyLines.join("\n")
  });
}

/**
 * Mengaktifkan / mematikan Ringkasan Harian dari menu kustom.
 */
function toggleDailyDigest() {
  const config = getAppConfig();
  const nextState = !config.emailEnabled;

  if (nextState && config.emailRecipients.length === 0) {
    showAlertSafe(
      "⚠️ Email Penerima Kosong",
      "Isi dulu kolom 'Email Penerima' di tab ⚙️ Config (sel C12), lalu aktifkan lagi fitur ini."
    );
    return;
  }

  setConfigEmailEnabled(nextState);
  installAllTriggers();
  showAlertSafe(
    nextState ? "🔔 Ringkasan Harian AKTIF" : "🔕 Ringkasan Harian NONAKTIF",
    nextState
      ? `Email ringkasan akan dikirim tiap jam 08:00 WIB ke: ${config.emailRecipients.join(", ")}`
      : "Email ringkasan harian sudah dimatikan."
  );
}

/**
 * Memasang trigger harian jam 07:00 WIB (refresh) & 08:00 WIB (email jika aktif) secara idempotent.
 */
function installAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    const fn = t.getHandlerFunction();
    if (fn === "scheduledDailyRefresh" || fn === "scheduledDailyEmailDigest") {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger("scheduledDailyRefresh")
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .inTimezone(APP_CONSTANTS.TIMEZONE)
    .create();

  const config = getAppConfig();
  if (config.emailEnabled) {
    ScriptApp.newTrigger("scheduledDailyEmailDigest")
      .timeBased()
      .everyDays(1)
      .atHour(8)
      .inTimezone(APP_CONSTANTS.TIMEZONE)
      .create();
  }

  showToastSafe("Trigger harian 07:00 WIB (dan 08:00 WIB jika aktif) berhasil dipasang.", "⏰ Trigger Aktif", 4);
}
