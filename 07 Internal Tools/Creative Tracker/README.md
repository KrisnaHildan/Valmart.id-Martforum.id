# ⚡ Valmart & Martforum Creative Engine & Dashboard
**Integrasi**: GitHub Repository (Naskah Detail) $\longleftrightarrow$ Google Sheets (Dashboard & Kalender)  
**Tujuan**: Sistem pemantauan produksi konten realtime yang lean, estetik, dan memotivasi tim kreatif (Ko Jeremy, Hildan, Helmy, Audrey, Dika).

---

## 🏗️ 3 Tab Utama yang Terbentuk

```
┌────────────────────────────────────────────────────────┐
│              SISTEM CREATIVE TRACKER LENGKAP           │
├──────────────────┬──────────────────┬──────────────────┤
│ 1. 📊 DASHBOARD  │ 2. 📅 KALENDER   │ 3. 📝 DATABASE   │
│ (Executive View) │ (Visual Monthly) │ (Master Table)   │
└──────────────────┴──────────────────┴──────────────────┘
```

### 1. Tab `📊 Dashboard` (Tampilan Eksekutif & Antrean Review)
* **4 Modern KPI Scorecards**:
  * 🎯 **Total Terjadwal**: Total konten bulan berjalan.
  * 🚀 **Approved / Published**: Berapa % konten yang sudah siap tayang / live.
  * ⚠️ **Antrean Review**: Jumlah video/naskah yang statusnya `🟠 In Review` atau `🔴 Revision` (perlu tindakan Ko Jeje/Hildan).
  * ⏳ **Dalam Pengerjaan**: Konten yang sedang diproduksi oleh talent (`Draft` & `In Progress`).
* **👥 Target Mingguan & Pacing Talent**:
  * **Helmy (Valmart)**: Target 4 Video/Mgg + Visual Progress Bar (Sparkline).
  * **Audrey (Martforum)**: Target 3 VO/Mgg + Visual Progress Bar (Sparkline).
  * **Dika / Jonathan (Editor/Desain)**: Support Aset Desain + Visual Progress Bar (Sparkline).
* **⚖️ Keseimbangan Brand**:
  * Rasio volume konten Valmart vs Martforum (menjaga sinergi cross-selling).
* **🚨 Antrean Review Otomatis (Action Needed)**:
  * Menampilkan otomatis daftar video/konten yang menunggu review agar Ko Jeremy / Hildan bisa langsung klik link draft tanpa mencari di database.

### 2. Tab `📅 Kalender Visual` (Grid Bulanan)
* Format kalender dinding 7 kolom (Senin s/d Minggu).
* Penanda tanggal otomatis: **`📍 HARI INI`** dengan border dan highlight warna ungu indigo.
* Di dalam kotak tanggal, muncul kartu ringkas: `[VAL]` / `[MART]`, nama PIC, judul hook, dan status badge warna (`🟢 Approved`, `🟡 In Progress`, dll).

### 3. Tab `📝 Database Konten` (Master Data 10 Kolom)
* Kolom lengkap dari tanggal, brand, PIC, hook, link naskah GitHub, link draft preview editor, status, hingga catatan review.

---

## 🚀 Cara Pasang ke Google Sheets (Tinggal Copy-Paste)

1. Buka Google Sheet Anda.
2. Buka menu **Extensions (Ekstensi)** $\rightarrow$ **Apps Script**.
3. Replace semua kodenya dengan isi file terbaru [apps_script/Code.gs](file:///d:/ANTIGRAVITY/Hildan/All%20Projects/Ko%20Jeremy/07%20Internal%20Tools/Creative%20Tracker/apps_script/Code.gs).
4. Tekan **Save (`Ctrl + S`)**.
5. Pilih fungsi **`setupTracker`**, lalu klik **Run (Jalankan)**.
   *(Semua 3 tab langsung terbuat dan terformat rapi secara otomatis!)*
6. Di Google Sheet Anda, klik menu **`⚡ Valmart & Martforum`** $\rightarrow$ **`📥 Sinkronisasi Naskah dari GitHub`**.
   * Seluruh data aktif November 2026 langsung masuk, dan Dashboard serta Kalender Visual langsung hidup realtime!
