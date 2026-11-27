/**
 * =========================================================================
 * SYNC.GS — Sinkronisasi Dinamis GitHub (API/Raw, Parser Markdown, Upsert)
 * =========================================================================
 */

/**
 * Dialog untuk menyimpan atau menghapus GitHub Personal Access Token di Script Properties.
 */
function promptGithubToken() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const hasToken = Boolean(props.getProperty(APP_CONSTANTS.SCRIPT_PROP_GITHUB_TOKEN));

  const res = ui.prompt(
    "🔑 Atur GitHub Personal Access Token",
    `Status saat ini: ${hasToken ? "✅ Token tersimpan" : "⚪ Belum ada token (mode repo public)"}\n\n` +
    "Masukkan GitHub Token (fine-grained read-only repo contents) jika repo bersifat PRIVATE.\n" +
    "Ketik HAPUS untuk menghapus token yang tersimpan:",
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;
  const input = res.getResponseText().trim();

  if (input.toUpperCase() === "HAPUS") {
    props.deleteProperty(APP_CONSTANTS.SCRIPT_PROP_GITHUB_TOKEN);
    ui.alert("✅ Token Dihapus", "Tracker sekarang berjalan dalam mode repo public.", ui.ButtonSet.OK);
    return;
  }

  if (!input) {
    ui.alert("Info", "Tidak ada perubahan pada token.", ui.ButtonSet.OK);
    return;
  }

  props.setProperty(APP_CONSTANTS.SCRIPT_PROP_GITHUB_TOKEN, input);
  ui.alert("✅ Token Tersimpan", "GitHub Token berhasil disimpan dengan aman di Script Properties.", ui.ButtonSet.OK);
}

/**
 * Dialog sederhana untuk memilih bulan/tahun sebelum sync dari GitHub.
 */
function promptAndSyncFromGithub() {
  const ui = SpreadsheetApp.getUi();
  const now = new Date();
  const defaultFilter = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;

  const res = ui.prompt(
    "📥 Sinkronisasi Naskah dari GitHub",
    "Masukkan periode Bulan & Tahun dengan format YYYY-MM (contoh: " + defaultFilter + ").\n" +
    "Kosongkan jika ingin sinkronisasi SEMUA bulan yang ada di Jadwal Konten.md:",
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) return;
  const raw = res.getResponseText().trim();

  let filterMonth = null;
  let filterYear = null;

  if (raw) {
    const match = raw.match(/^(\d{4})-(\d{1,2})$/);
    if (!match) {
      ui.alert("⚠️ Format Salah", "Gunakan format YYYY-MM (contoh: 2026-11) atau kosongkan untuk semua bulan.", ui.ButtonSet.OK);
      return;
    }
    filterYear = parseInt(match[1], 10);
    filterMonth = parseInt(match[2], 10);
    if (filterMonth < 1 || filterMonth > 12) {
      ui.alert("⚠️ Bulan Tidak Valid", "Bulan harus antara 01 sampai 12.", ui.ButtonSet.OK);
      return;
    }
  }

  syncFromGithub(filterMonth, filterYear);
}

/**
 * Menjalankan sinkronisasi non-destruktif (upsert) dari file sumber jadwal di GitHub.
 */
function syncFromGithub(filterMonth, filterYear) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    showAlertSafe("Sedang Diproses", "Sinkronisasi sedang dijalankan pengguna lain. Coba lagi sebentar.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = getAppConfig();
    let dbSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.DATABASE);
    if (!dbSheet) {
      setupInitial();
      dbSheet = ss.getSheetByName(APP_CONSTANTS.SHEETS.DATABASE);
    }

    showToastSafe("Mengambil jadwal dari GitHub...", "📥 Sinkronisasi", 5);
    const markdown = fetchGithubFileContent(config, config.scheduleFile);
    const parseResult = parseMarkdownSchedule(markdown);

    let rowsToSync = parseResult.rows;
    if (filterMonth && filterYear) {
      rowsToSync = rowsToSync.filter(r => {
        const dt = parseDateSafe(r.tanggal);
        return dt && (dt.getMonth() + 1) === filterMonth && dt.getFullYear() === filterYear;
      });
    }

    if (rowsToSync.length === 0) {
      showAlertSafe(
        "ℹ️ Tidak Ada Data",
        "Tidak ditemukan baris jadwal yang cocok di file sumber GitHub" +
        (filterMonth ? ` untuk periode ${filterYear}-${String(filterMonth).padStart(2, "0")}.` : ".")
      );
      return;
    }

    const summary = upsertScheduleToDatabase(dbSheet, rowsToSync, config, {
      filterMonth: filterMonth,
      filterYear: filterYear,
      checkLinks: true
    });

    renderCalendar();
    renderDashboard();

    const msg =
      `Sinkronisasi selesai!\n\n` +
      `• Baris baru: ${summary.added}\n` +
      `• Diupdate: ${summary.updated}\n` +
      `• Tidak berubah: ${summary.unchanged}\n` +
      `• Tidak ada di GitHub: ${summary.missing}\n` +
      (summary.linkWarnings > 0 ? `• ⚠️ Peringatan link 404: ${summary.linkWarnings} baris\n` : "") +
      (parseResult.warnings.length > 0 ? `\nCatatan parser: ${parseResult.warnings.length} baris dilewati karena format tidak valid.` : "");

    showAlertSafe("📥 Ringkasan Sinkronisasi", msg);
  } catch (err) {
    showAlertSafe("❌ Gagal Sinkronisasi", err.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mengambil isi file teks dari GitHub (mendukung repo private via token & public via raw URL).
 */
function fetchGithubFileContent(config, relativePath) {
  const token = PropertiesService.getScriptProperties().getProperty(APP_CONSTANTS.SCRIPT_PROP_GITHUB_TOKEN);

  if (token) {
    const apiUrl = buildGithubApiUrl(config, relativePath);
    const resp = UrlFetchApp.fetch(apiUrl, {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        "Authorization": "Bearer " + token,
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "Valmart-Martforum-Tracker"
      }
    });
    return handleGithubApiResponse(resp, relativePath);
  }

  // Fallback repo public: raw URL
  const rawUrl = buildGithubRawUrl(config, relativePath);
  const rawResp = UrlFetchApp.fetch(rawUrl, {
    method: "get",
    muteHttpExceptions: true,
    headers: { "User-Agent": "Valmart-Martforum-Tracker" }
  });
  const code = rawResp.getResponseCode();
  if (code === 200) {
    return rawResp.getContentText("UTF-8");
  }
  if (code === 404) {
    throw new Error(`File "${relativePath}" tidak ditemukan (404). Jika repo bersifat PRIVATE, atur GitHub Token lewat menu [🔑 Atur GitHub Token].`);
  }
  throw new Error(`Gagal mengakses GitHub (HTTP ${code}). Periksa pengaturan Repo di tab Config.`);
}

/**
 * Memproses respons GitHub Contents API termasuk Base64 decode & pesan error ramah.
 */
function handleGithubApiResponse(resp, relativePath) {
  const code = resp.getResponseCode();
  if (code === 200) {
    const json = JSON.parse(resp.getContentText("UTF-8"));
    if (!json.content) {
      throw new Error(`File "${relativePath}" ditemukan tetapi kontennya kosong.`);
    }
    const cleanB64 = json.content.replace(/\s+/g, "");
    const decodedBytes = Utilities.base64Decode(cleanB64);
    return Utilities.newBlob(decodedBytes).getDataAsString("UTF-8");
  }
  if (code === 401) {
    throw new Error("GitHub Token tidak valid atau kedaluwarsa (401). Silakan perbarui lewat menu [🔑 Atur GitHub Token].");
  }
  if (code === 403) {
    throw new Error("Akses GitHub ditolak atau terkena rate limit (403). Pastikan token punya izin membaca Contents.");
  }
  if (code === 404) {
    throw new Error(`File "${relativePath}" tidak ditemukan di repository (404). Periksa Base Path & nama file di tab Config.`);
  }
  throw new Error(`Terjadi kendala saat menghubungi GitHub API (HTTP ${code}).`);
}

/**
 * Cek ringan apakah file naskah ada di GitHub (hasil di-cache per path selama 1 kali sync).
 */
function checkGithubFileExists(config, relativePath, cache) {
  if (!relativePath) return false;
  if (cache && Object.prototype.hasOwnProperty.call(cache, relativePath)) {
    return cache[relativePath];
  }

  const token = PropertiesService.getScriptProperties().getProperty(APP_CONSTANTS.SCRIPT_PROP_GITHUB_TOKEN);
  let exists = false;
  try {
    if (token) {
      const apiUrl = buildGithubApiUrl(config, relativePath);
      const resp = UrlFetchApp.fetch(apiUrl, {
        method: "get",
        muteHttpExceptions: true,
        headers: {
          "Authorization": "Bearer " + token,
          "Accept": "application/vnd.github.v3+json",
          "User-Agent": "Valmart-Martforum-Tracker"
        }
      });
      exists = (resp.getResponseCode() === 200);
    } else {
      const rawUrl = buildGithubRawUrl(config, relativePath);
      const resp = UrlFetchApp.fetch(rawUrl, {
        method: "get",
        muteHttpExceptions: true,
        headers: { "User-Agent": "Valmart-Martforum-Tracker" }
      });
      exists = (resp.getResponseCode() === 200);
    }
  } catch (e) {
    exists = false;
  }

  if (cache) cache[relativePath] = exists;
  return exists;
}

/**
 * Parser tabel Markdown yang toleran terhadap spasi ekstra, pipe ujung hilang, & baris kosong.
 * Kolom standar: ID | Tanggal | Brand | PIC | Topik & Hook | File Naskah
 */
function parseMarkdownSchedule(markdownText) {
  const lines = String(markdownText || "").split(/\r?\n/);
  const rows = [];
  const warnings = [];
  let inTable = false;
  let colMap = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i].trim();
    if (!line || !line.includes("|")) {
      inTable = false;
      continue;
    }

    // Pecah kolom dan buang pipe kosong di awal/akhir baris
    let cells = line.split("|").map(c => c.trim());
    if (cells.length > 0 && cells[0] === "") cells.shift();
    if (cells.length > 0 && cells[cells.length - 1] === "") cells.pop();
    if (cells.length < 4) continue;

    // Deteksi baris separator (--- / :---:)
    if (cells.every(c => /^:?-{2,}:?$/.test(c))) {
      continue;
    }

    // Deteksi header tabel
    const lowerCells = cells.map(c => c.replace(/[*_`]/g, "").trim().toLowerCase());
    if (lowerCells.some(c => c === "id") && lowerCells.some(c => c.includes("tanggal"))) {
      colMap = {
        id: lowerCells.findIndex(c => c === "id"),
        tanggal: lowerCells.findIndex(c => c.includes("tanggal")),
        brand: lowerCells.findIndex(c => c.includes("brand")),
        pic: lowerCells.findIndex(c => c.includes("pic")),
        topik: lowerCells.findIndex(c => c.includes("topik") || c.includes("hook")),
        file: lowerCells.findIndex(c => c.includes("file") || c.includes("naskah"))
      };
      inTable = true;
      continue;
    }

    if (!inTable || !colMap) continue;

    const stripMd = (val) => String(val || "").trim().replace(/^[*_`]+|[*_`]+$/g, "").trim();
    const id = stripMd(colMap.id >= 0 ? cells[colMap.id] : "");
    const tanggalRaw = stripMd(colMap.tanggal >= 0 ? cells[colMap.tanggal] : "");
    const brand = stripMd(colMap.brand >= 0 ? cells[colMap.brand] : "");
    const pic = stripMd(colMap.pic >= 0 ? cells[colMap.pic] : "");
    const topik = (colMap.topik >= 0 ? cells[colMap.topik] : "") || "";
    const fileNaskah = stripMd(colMap.file >= 0 ? cells[colMap.file] : "").replace(/\\/g, "/");

    const dt = parseDateSafe(tanggalRaw);
    if (!id || !dt) {
      warnings.push(`Baris ${i + 1}: ID atau tanggal (${tanggalRaw}) tidak valid.`);
      continue;
    }

    rows.push({
      id: id,
      tanggal: formatDateIso(dt),
      brand: brand,
      pic: pic,
      topik: topik,
      fileNaskah: fileNaskah
    });
  }

  return { rows: rows, warnings: warnings };
}

/**
 * Upsert non-destruktif ke tab 📝 Database Konten.
 * Kolom milik tim (Link Draft, Status, Catatan Review) TIDAK PERNAH ditimpa.
 */
function upsertScheduleToDatabase(dbSheet, parsedRows, config, options) {
  const opts = options || {};
  const startRow = APP_CONSTANTS.DB_START_ROW;
  const lastRow = dbSheet.getLastRow();
  const numExisting = lastRow >= startRow ? (lastRow - startRow + 1) : 0;

  const existingVals = numExisting > 0 ? dbSheet.getRange(startRow, APP_CONSTANTS.DB_COLS.ID, numExisting, 11).getValues() : [];
  const existingFormulas = numExisting > 0 ? dbSheet.getRange(startRow, APP_CONSTANTS.DB_COLS.ID, numExisting, 11).getFormulas() : [];

  const existingById = {};
  const outputRows = [];
  let isIndo = false;
  try {
    isIndo = isIndonesianLocale(dbSheet.getParent());
  } catch (e) {
    isIndo = false;
  }

  for (let i = 0; i < existingVals.length; i++) {
    const rowVals = existingVals[i].slice();
    const rowForms = existingFormulas[i];
    const rowId = String(rowVals[0] || "").trim();

    // Lewati baris kosong total agar baris baru tidak terdorong ke bawah
    const isRowEmpty = rowVals.every(v => String(v === null || v === undefined ? "" : v).trim() === "");
    if (isRowEmpty) continue;

    // Pertahankan formula HYPERLINK di kolom Link Naskah (idx 6) & Link Draft (idx 7) jika ada
    if (rowForms[6]) rowVals[6] = rowForms[6];
    if (rowForms[7]) rowVals[7] = rowForms[7];

    const outIdx = outputRows.length;
    outputRows.push(rowVals);
    if (rowId) {
      existingById[rowId] = outIdx;
    }
  }

  const linkCache = {};
  const seenIds = {};
  const missingTag = "[Tidak ada di GitHub]";
  let added = 0;
  let updated = 0;
  let unchanged = 0;
  let missing = 0;
  let linkWarnings = 0;

  for (let i = 0; i < parsedRows.length; i++) {
    const item = parsedRows[i];
    seenIds[item.id] = true;

    const dt = parseDateSafe(item.tanggal);
    const dayName = getIndonesianDayName(dt);
    const normBrand = normalizeBrand(item.brand, config.brands);
    const normPic = normalizePic(item.pic, config.pics);
    const blobUrl = item.fileNaskah ? buildGithubBlobUrl(config, item.fileNaskah) : "";

    let linkFormula = "";
    let expectedLinkLabel = "";
    if (blobUrl) {
      let linkOk = true;
      if (opts.checkLinks) {
        linkOk = checkGithubFileExists(config, item.fileNaskah, linkCache);
      }
      if (!linkOk) {
        linkWarnings++;
        expectedLinkLabel = "⚠️ Link 404 (Cek Repo)";
      } else {
        expectedLinkLabel = "📖 Buka Naskah";
      }
      linkFormula = makeHyperlinkFormula(blobUrl, expectedLinkLabel, isIndo);
    }

    if (Object.prototype.hasOwnProperty.call(existingById, item.id)) {
      const idx = existingById[item.id];
      const cur = outputRows[idx];
      const curDateIso = formatDateIso(parseDateSafe(cur[1]));
      const curFormula = typeof cur[6] === "string" && cur[6].startsWith("=") ? cur[6] : "";
      const curLinkUrl = extractUrlFromCell(cur[6], curFormula);
      const curLinkLabel = extractLabelFromCell(cur[6], curFormula);
      const curNote = String(cur[9] || "").trim();

      const changed =
        curDateIso !== item.tanggal ||
        String(cur[2] || "") !== dayName ||
        String(cur[3] || "") !== normBrand ||
        String(cur[4] || "") !== normPic ||
        String(cur[5] || "") !== item.topik ||
        curLinkUrl !== blobUrl ||
        curLinkLabel !== expectedLinkLabel;

      if (changed) {
        cur[1] = dt;
        cur[2] = dayName;
        cur[3] = normBrand;
        cur[4] = normPic;
        cur[5] = item.topik;
        cur[6] = linkFormula;
        cur[10] = new Date();
        // cur[7] (Link Draft), cur[8] (Status), cur[9] (Catatan) TIDAK ditimpa
        updated++;
      } else {
        unchanged++;
      }

      if (curNote.includes(missingTag)) {
        cur[9] = curNote.replace(/\[Tidak ada di GitHub\]\s*/g, "").trim();
      }
    } else {
      outputRows.push([
        item.id,          // B: ID
        dt,               // C: Tanggal Tayang
        dayName,          // D: Hari
        normBrand,        // E: Brand
        normPic,          // F: PIC & Role
        item.topik,       // G: Topik & Hook
        linkFormula,      // H: Link Naskah
        "",               // I: Link Draft Preview (kosong untuk tim)
        "⚪ Draft",       // J: Status default
        "",               // K: Catatan Review
        new Date()        // L: Update Terakhir
      ]);
      added++;
    }
  }

  // Tandai baris yang hilang dari sumber GitHub (tanpa menghapus)
  for (let i = 0; i < outputRows.length; i++) {
    const rowId = String(outputRows[i][0] || "").trim();
    if (!rowId || seenIds[rowId]) continue;

    if (opts.filterMonth && opts.filterYear) {
      const rDate = parseDateSafe(outputRows[i][1]);
      if (!rDate || (rDate.getMonth() + 1) !== opts.filterMonth || rDate.getFullYear() !== opts.filterYear) {
        continue;
      }
    }

    const curNote = String(outputRows[i][9] || "").trim();
    if (!curNote.includes(missingTag)) {
      outputRows[i][9] = curNote ? `${missingTag} ${curNote}` : missingTag;
    }
    missing++;
  }

  if (outputRows.length > 0) {
    const targetRange = dbSheet.getRange(startRow, APP_CONSTANTS.DB_COLS.ID, outputRows.length, 11);
    targetRange.setValues(outputRows);
    dbSheet.getRange(startRow, APP_CONSTANTS.DB_COLS.DATE, outputRows.length, 1).setNumberFormat("yyyy-mm-dd");
    dbSheet.getRange(startRow, APP_CONSTANTS.DB_COLS.UPDATED_AT, outputRows.length, 1).setNumberFormat("yyyy-mm-dd hh:mm");
  }
  if (numExisting > outputRows.length) {
    dbSheet.getRange(startRow + outputRows.length, APP_CONSTANTS.DB_COLS.ID, numExisting - outputRows.length, 11).clearContent();
  }

  return { added, updated, unchanged, missing, linkWarnings };
}
