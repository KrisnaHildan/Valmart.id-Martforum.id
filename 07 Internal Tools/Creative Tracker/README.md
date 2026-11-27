# ⚡ Valmart & Martforum Creative Content Engine & Tracker
**Integrasi**: GitHub Repository (`Jadwal Konten.md` & Naskah Detail) $\longleftrightarrow$ Google Sheets (Dashboard, Kalender, Database, Config)  
**Tujuan**: Sistem pemantauan produksi konten real-time untuk **Ko Jeremy (Jeje)**, **Hildan**, **Helmy**, **Audrey**, dan **Dika** *(pengganti Jonathan yang sedang cuti)*.

---

## 🏗️ Arsitektur 4 Tab Utama

1. **`📊 Dashboard`**: 5 KPI Scorecards (termasuk **Konten Telat**), filter Bulan/Tahun & selector Minggu, Pacing Talent aktif (Helmy, Audrey, Dika — Jonathan cuti target 0), Distribusi Brand, Antrean Review otomatis, dan Daftar Konten Telat.
2. **`📅 Kalender Visual`**: Grid bulanan Senin–Minggu dengan dropdown Bulan (`C2`) & Tahun (`E2`) yang auto-refresh saat diubah, penanda `📍 HARI INI`, highlight akhir pekan, dan batas 4 kartu per tanggal (`+N lainnya`).
3. **`📝 Database Konten`**: Master data 11 kolom (`ID`, `Tanggal Tayang`, `Hari`, `Brand`, `PIC & Role`, `Topik & Hook`, `Link Naskah`, `Link Draft Preview`, `Status`, `Catatan Review`, `Update Terakhir`) dengan validasi 1.000 baris, auto-filter, banded rows, dan highlight merah otomatis untuk konten **TELAT**.
4. **`⚙️ Config`**: Pusat pengaturan repo GitHub, daftar PIC & target mingguan, brand, status, dan penerima email ringkasan harian (bisa diubah tanpa menyentuh kode).

---

## 📁 Struktur File Apps Script

Semua file berada di folder [Apps Script](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/appsscript.json):
- `Config.gs` — Konstanta global & pembaca/penulis tab `⚙️ Config`
- `Utils.gs` — Helper tanggal Indonesia, URL GitHub (`encodeURI`), & UI feedback aman
- `Setup.gs` — Menu kustom, `setupInitial()` (aman/idempotent), `resetTotalWithConfirmation()` (dengan backup Drive), & format database 1.000 baris
- `Dashboard.gs` — Render KPI, pacing mingguan, antrean review, & daftar telat
- `Calendar.gs` — Render kalender bulanan secara batch
- `Sync.gs` — Fetch GitHub (API + token untuk private, raw URL untuk public), parser `Jadwal Konten.md`, dan upsert non-destruktif
- `Triggers.gs` — `onEdit` auto-refresh, trigger harian 07:00 WIB, dan email ringkasan 08:00 WIB
- `Tests.gs` — `testRunAll()` untuk verifikasi parser, deteksi telat, kalender, & upsert 2x
- `appsscript.json` — Konfigurasi V8, `Asia/Jakarta`, dan OAuth scopes minimal

---

## 🚀 Cara Deploy ke Google Sheets

### Opsi A: Copy-Paste Manual
1. Buka Google Sheet tujuan $\rightarrow$ **Extensions (Ekstensi)** $\rightarrow$ **Apps Script**.
2. Aktifkan tampilan file manifest di **Project Settings** $\rightarrow$ centang *"Show 'appsscript.json' manifest file in editor"*.
3. Salin isi file berikut ke project Apps Script (buat file dengan nama persis sama):
   - [appsscript.json](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/appsscript.json)
   - [Config.gs](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/Config.gs)
   - [Utils.gs](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/Utils.gs)
   - [Setup.gs](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/Setup.gs)
   - [Dashboard.gs](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/Dashboard.gs)
   - [Calendar.gs](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/Calendar.gs)
   - [Sync.gs](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/Sync.gs)
   - [Triggers.gs](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/Triggers.gs)
   - [Tests.gs](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Apps%20Script/Tests.gs)
4. Simpan (`Ctrl + S`), lalu muat ulang tab Google Sheet.

### Opsi B: Via `clasp` (CLI)
```powershell
cd "All Projects\Valmart x Martforum\07 Internal Tools\Creative Tracker\Apps Script"
clasp login
clasp push
```

---

## 🔑 Setup Awal & GitHub Token
1. Di Google Sheet, klik menu **`⚡ Valmart & Martforum`** $\rightarrow$ **`🛠️ Setup Awal (Aman)`**.
2. *(Khusus Repo Private)* Klik **`🔑 Atur GitHub Token`** lalu masukkan Personal Access Token (fine-grained, read-only `Contents`). Token disimpan aman di `Script Properties`, bukan di kode atau sheet.
3. Klik **`📥 Sinkronisasi Naskah dari GitHub`**:
   - Isi periode `YYYY-MM` (misal `2026-11`) atau kosongkan untuk sync semua bulan dari [Jadwal Konten.md](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/04%20Strategi%20Konten%20dan%20Kreatif/Jadwal%20Konten.md).
   - Sync bersifat **non-destruktif (upsert)**: kolom **Link Draft Preview**, **Status**, dan **Catatan Review** milik tim **tidak akan pernah ditimpa**.
4. Klik **`⏰ Pasang Trigger Otomatis`** agar kalender auto-refresh tiap jam 07:00 WIB.
5. *(Opsional)* Isi email Ko Jeremy & Hildan di tab **`⚙️ Config`** sel `C12` (pisahkan dengan koma), lalu aktifkan ringkasan email jam 08:00 WIB lewat menu **`🔔 Aktifkan / Matikan Ringkasan Harian`**.

---

## 📚 Panduan Tim
Lihat panduan penggunaan harian untuk talent, editor, dan reviewer di [Panduan Tim.md](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Valmart%20x%20Martforum/07%20Internal%20Tools/Creative%20Tracker/Panduan%20Tim.md).
