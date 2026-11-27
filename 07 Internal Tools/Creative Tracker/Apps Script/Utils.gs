/**
 * =========================================================================
 * UTILS.GS — Helper Tanggal, URL GitHub, Normalisasi, & UI Feedback
 * =========================================================================
 */

const INDONESIAN_DAYS = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
const INDONESIAN_MONTHS = [
  "Januari", "Februari", "Maret", "April", "Mei", "Juni",
  "Juli", "Agustus", "September", "Oktober", "November", "Desember"
];

/**
 * Mengubah nilai sel atau string tanggal (yyyy-mm-dd) menjadi objek Date lokal (jam 00:00:00).
 */
function parseDateSafe(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  const str = String(val).trim();
  if (!str) return null;

  // Format yyyy-mm-dd atau yyyy/mm/dd
  const isoMatch = str.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (isoMatch) {
    const y = parseInt(isoMatch[1], 10);
    const m = parseInt(isoMatch[2], 10) - 1;
    const d = parseInt(isoMatch[3], 10);
    const dt = new Date(y, m, d);
    if (dt.getFullYear() === y && dt.getMonth() === m && dt.getDate() === d) {
      return dt;
    }
  }

  // Format dd/mm/yyyy atau dd-mm-yyyy
  const dmyMatch = str.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (dmyMatch) {
    const d = parseInt(dmyMatch[1], 10);
    const m = parseInt(dmyMatch[2], 10) - 1;
    const y = parseInt(dmyMatch[3], 10);
    const dt = new Date(y, m, d);
    if (dt.getFullYear() === y && dt.getMonth() === m && dt.getDate() === d) {
      return dt;
    }
  }

  return null;
}

/**
 * Format Date menjadi string yyyy-mm-dd.
 */
function formatDateIso(date) {
  if (!date || isNaN(date.getTime())) return "";
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

/**
 * Mengembalikan nama hari bahasa Indonesia dari Date.
 */
function getIndonesianDayName(date) {
  if (!date || isNaN(date.getTime())) return "";
  return INDONESIAN_DAYS[date.getDay()];
}

/**
 * Mengembalikan nama bulan bahasa Indonesia (1 = Januari).
 */
function getIndonesianMonthName(monthNum) {
  return INDONESIAN_MONTHS[(monthNum - 1)] || "";
}

/**
 * Mengembalikan Date hari Senin di minggu yang sama dengan tanggal acuan.
 */
function getMondayOfWeek(refDate) {
  const d = new Date(refDate.getFullYear(), refDate.getMonth(), refDate.getDate());
  const day = d.getDay();
  const diff = (day === 0) ? -6 : (1 - day);
  d.setDate(d.getDate() + diff);
  return d;
}

/**
 * Menentukan apakah sebuah konten berstatus TELAT (tanggal < hari ini dan belum Published).
 */
function isTaskLate(dateVal, statusVal, todayRef) {
  const dt = parseDateSafe(dateVal);
  if (!dt) return false;
  const status = String(statusVal || "").trim();
  if (status === "🟣 Published") return false;

  const ref = todayRef || new Date();
  const todayMidnight = new Date(ref.getFullYear(), ref.getMonth(), ref.getDate());
  return dt.getTime() < todayMidnight.getTime();
}

/**
 * Menggabungkan basePath dan relativePath dengan rapi tanpa double slash.
 */
function joinRepoPath(basePath, relativePath) {
  const cleanBase = String(basePath || "").replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
  const cleanRel = String(relativePath || "").replace(/\\/g, "/").replace(/^\/+|\/+$/g, "");
  if (!cleanBase) return cleanRel;
  if (!cleanRel) return cleanBase;
  // Jika relativePath sudah diawali basePath, jangan digandakan
  if (cleanRel.toLowerCase().startsWith(cleanBase.toLowerCase() + "/")) {
    return cleanRel;
  }
  return `${cleanBase}/${cleanRel}`;
}

/**
 * Membangun URL GitHub web (blob) dengan encodeURI yang aman untuk spasi & Title Case.
 */
function buildGithubBlobUrl(config, relativePath) {
  const fullPath = joinRepoPath(config.basePath, relativePath);
  const encodedPath = fullPath.split("/").map(seg => encodeURIComponent(seg)).join("/");
  return `https://github.com/${config.repoOwner}/${config.repoName}/blob/${encodeURIComponent(config.branch)}/${encodedPath}`;
}

/**
 * Membangun URL GitHub Contents API.
 */
function buildGithubApiUrl(config, relativePath) {
  const fullPath = joinRepoPath(config.basePath, relativePath);
  const encodedPath = fullPath.split("/").map(seg => encodeURIComponent(seg)).join("/");
  return `https://api.github.com/repos/${config.repoOwner}/${config.repoName}/contents/${encodedPath}?ref=${encodeURIComponent(config.branch)}`;
}

/**
 * Membangun URL raw.githubusercontent.com untuk fallback repo public.
 */
function buildGithubRawUrl(config, relativePath) {
  const fullPath = joinRepoPath(config.basePath, relativePath);
  const encodedPath = fullPath.split("/").map(seg => encodeURIComponent(seg)).join("/");
  return `https://raw.githubusercontent.com/${config.repoOwner}/${config.repoName}/${encodeURIComponent(config.branch)}/${encodedPath}`;
}

/**
 * Membangun URL GitHub Git Trees API untuk mengambil seluruh daftar file dalam 1 request.
 */
function buildGithubTreeApiUrl(config) {
  return `https://api.github.com/repos/${config.repoOwner}/${config.repoName}/git/trees/${encodeURIComponent(config.branch)}?recursive=1`;
}

/**
 * Membuat formula =HYPERLINK("url", "label") yang aman & sesuai locale spreadsheet.
 */
function makeHyperlinkFormula(url, label, isIndo) {
  if (!url) return "";
  const safeUrl = String(url).replace(/"/g, "%22");
  const safeLabel = String(label || "Buka Link").replace(/"/g, "'");
  return localizeFormula(`=HYPERLINK("${safeUrl}", "${safeLabel}")`, Boolean(isIndo));
}

/**
 * Mengekstrak URL dari formula =HYPERLINK("...", "...") atau nilai sel biasa.
 */
function extractUrlFromCell(value, formula) {
  if (formula && typeof formula === "string") {
    const match = formula.match(/=HYPERLINK\(\s*"([^"]+)"/i);
    if (match) return match[1];
  }
  const str = String(value || "").trim();
  if (/^https?:\/\//i.test(str)) return str;
  return "";
}

/**
 * Mengekstrak label tampilan dari formula =HYPERLINK("...", "...") atau nilai sel.
 */
function extractLabelFromCell(value, formula) {
  if (formula && typeof formula === "string") {
    const match = formula.match(/=HYPERLINK\(\s*"[^"]+"\s*[,;]\s*"([^"]*)"/i);
    if (match) return match[1];
  }
  return String(value || "").trim();
}

/**
 * Normalisasi nama brand terhadap daftar brand di Config (tahan typo ringan).
 */
function normalizeBrand(rawBrand, brandList) {
  const raw = String(rawBrand || "").trim().toLowerCase();
  const list = brandList || APP_CONSTANTS.DEFAULTS.BRANDS;
  for (let i = 0; i < list.length; i++) {
    if (raw === list[i].toLowerCase() || raw.includes(list[i].toLowerCase())) {
      return list[i];
    }
  }
  return String(rawBrand || "").trim();
}

/**
 * Normalisasi nama PIC terhadap daftar PIC di Config (prioritaskan exact match lalu kata utuh).
 */
function normalizePic(rawPic, picList) {
  const raw = String(rawPic || "").trim();
  if (!raw) return "";
  const rawLower = raw.toLowerCase();
  for (let i = 0; i < picList.length; i++) {
    if (rawLower === picList[i].name.toLowerCase()) {
      return picList[i].name;
    }
  }
  for (let i = 0; i < picList.length; i++) {
    const pName = picList[i].name;
    const firstWord = pName.split(" ")[0].toLowerCase();
    const escaped = firstWord.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const rx = new RegExp(`(?:^|\\s)${escaped}(?:\\s|\\(|$)`, "i");
    if (rx.test(rawLower)) {
      return pName;
    }
  }
  return raw;
}

/**
 * Daftar pilihan tahun standar untuk dropdown filter.
 */
function getSupportedYearsList(includeAll) {
  const currentYear = new Date().getFullYear();
  const maxYear = Math.max(2028, currentYear + 2);
  const years = [];
  if (includeAll) years.push("Semua");
  for (let y = 2025; y <= maxYear; y++) {
    years.push(String(y));
  }
  return years;
}

/**
 * Parsing nilai pilihan filter Bulan ("Semua" atau 1..12).
 */
function parseMonthFilterValue(rawVal, allowAll, fallbackMonth) {
  if (rawVal !== null && rawVal !== undefined) {
    const str = String(rawVal).trim();
    if (allowAll && str.toLowerCase() === "semua") {
      return "Semua";
    }
    if (typeof rawVal === "number" || /^\d+$/.test(str)) {
      const n = Number(rawVal);
      if (Number.isInteger(n) && n >= 1 && n <= 12) return n;
    }
  }
  return fallbackMonth !== undefined ? fallbackMonth : (new Date().getMonth() + 1);
}

/**
 * Parsing nilai pilihan filter Tahun ("Semua" atau YEAR_MIN..YEAR_MAX).
 */
function parseYearFilterValue(rawVal, allowAll, fallbackYear) {
  if (rawVal !== null && rawVal !== undefined) {
    const str = String(rawVal).trim();
    if (allowAll && str.toLowerCase() === "semua") {
      return "Semua";
    }
    if (typeof rawVal === "number" || /^\d+$/.test(str)) {
      const n = Number(rawVal);
      if (Number.isInteger(n) && n >= APP_CONSTANTS.YEAR_MIN && n <= APP_CONSTANTS.YEAR_MAX) {
        return n;
      }
    }
  }
  return fallbackYear !== undefined ? fallbackYear : (new Date().getFullYear());
}

/**
 * Menghitung kapasitas baris dinamis Database Konten agar tetap aman di atas 1.000 baris.
 */
function getDatabaseCapacityRows(dbSheet, minNeededDataRows) {
  const startRow = APP_CONSTANTS.DB_START_ROW;
  const lastRow = dbSheet ? dbSheet.getLastRow() : 0;
  const usedRows = lastRow >= startRow ? (lastRow - startRow + 1) : 0;
  return Math.max(
    APP_CONSTANTS.MIN_DB_ROWS,
    usedRows + APP_CONSTANTS.DB_BUFFER_ROWS,
    minNeededDataRows || 0
  );
}

/**
 * Konversi indeks kolom (1-based) menjadi huruf kolom A1 (misal 2 -> "B").
 */
function columnToLetter(colIndex) {
  let temp = "";
  let col = colIndex;
  while (col > 0) {
    const rem = (col - 1) % 26;
    temp = String.fromCharCode(rem + 65) + temp;
    col = Math.floor((col - rem - 1) / 26);
  }
  return temp;
}

/**
 * Toast aman yang tidak error saat dipanggil dari trigger background.
 */
function showToastSafe(message, title, seconds) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title || "Info", seconds || 4);
  } catch (e) {
    Logger.log(`[TOAST] ${title || "Info"}: ${message}`);
  }
}

/**
 * Alert aman yang fallback ke Logger jika berjalan di mode non-UI (trigger).
 */
function showAlertSafe(title, message) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(title, message, ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(`[ALERT] ${title}: ${message}`);
  }
}

/**
 * Mengecek apakah locale spreadsheet menggunakan format Indonesia (separator ;).
 */
function isIndonesianLocale(ss) {
  const target = ss || SpreadsheetApp.getActiveSpreadsheet();
  const locale = String(target.getSpreadsheetLocale() || "").toLowerCase();
  return locale.startsWith("id") || locale.startsWith("in");
}

/**
 * Mengonversi formula sintaks Inggris (koma) ke sintaks lokal Indonesia (titik koma & backslash)
 * tanpa mengubah koma di dalam tanda kutip.
 */
function localizeFormula(formula, isIndo) {
  if (!isIndo || !formula || typeof formula !== "string" || !formula.startsWith("=")) {
    return formula;
  }
  let out = "";
  let inQuotes = false;
  const stack = [];

  for (let i = 0; i < formula.length; i++) {
    const ch = formula[i];
    if (ch === "\"") {
      inQuotes = !inQuotes;
      out += ch;
      continue;
    }
    if (!inQuotes) {
      if (ch === "{" || ch === "(") {
        stack.push(ch);
      } else if ((ch === "}" || ch === ")") && stack.length > 0) {
        stack.pop();
      } else if (ch === ",") {
        const top = stack.length > 0 ? stack[stack.length - 1] : "";
        out += (top === "{") ? "\\" : ";";
        continue;
      }
    }
    out += ch;
  }
  return out;
}
