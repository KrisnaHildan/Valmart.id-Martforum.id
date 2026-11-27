# 📱 Panduan Cepat Tim - Creative Tracker Valmart & Martforum

Selamat datang! Sheet ini dibuat supaya kita semua gampang pantau jadwal konten tanpa ribet. Cukup ikuti panduan di bawah ini setiap hari.

---

## 1. Tugas Kamu Sehari-hari (Di Tab `📝 Database Konten`)

Gunakan **Filter Bulan (`D4`) & Tahun (`F4`)** di baris 4 kalau mau fokus ke bulan aktif saja, atau pilih **`Semua`** untuk melihat seluruh riwayat konten dari awal.

Saat mau mengerjakan atau setor konten, kamu **CUKUP EDIT 3 KOLOM INI SAJA**:

1. **`🎬 Link Draft Preview`** -> Tempel link Google Drive / TikTok / Frame.io hasil editan kamu.
2. **`Status`** -> Pilih status dari dropdown:
   - `⚪ Draft`: Belum mulai dikerjakan.
   - `🟡 In Progress`: Lagi proses syuting / VO / editing.
   - `🟠 In Review`: Sudah beres! Siap dicek Ko Jeremy / Hildan *(pastikan Link Draft sudah diisi)*.
   - `🔴 Revision`: Ada catatan revisi (cek kolom Catatan Review).
   - `🟢 Approved`: Sudah ACC, siap tayang sesuai tanggal.
   - `🟣 Published`: Sudah tayang di medsos!
3. **`Catatan Review / Feedback`** -> Tulis catatan kalau ada kendala atau sudah selesai revisi.

> 💡 **Tips:** Kolom `Hari` dan `Update Terakhir` akan terisi **otomatis** begitu kamu ubah tanggal atau status. Naskah lengkap bisa langsung diklik di kolom **`📖 Link Naskah (GitHub)`**.

---

## 2. Cara Baca `📊 Dashboard` & `📅 Kalender Visual`

- **`📊 Dashboard` (Untuk Ko Jeremy, Hildan, & Cek Target):**
  - Lihat **5 Kartu Atas** untuk cek progres bulan terpilih plus **Antrean Review** dan **Konten Telat** (kedua bagian merah/oranye ini memantau **semua bulan**).
  - Di tabel **Antrean Review**, **Link Draft** bisa diklik langsung tanpa harus bolak-balik cari di tab Database.
  - Lihat tabel **Target Mingguan** untuk cek sisa tugas terjadwal minggu ini per orang aktif (Helmy, Audrey, Dika menggantikan Jonathan yang cuti).
  - Gunakan dropdown **Filter Bulan/Tahun** atau **Pacing Minggu** di baris 4 kalau mau lihat periode lain.
- **`📅 Kalender Visual`:**
  - Kotak berwarna ungu (`📍 HARI INI`) adalah jadwal hari ini, dan kotak merah (`🚨 TELAT`) menandakan jadwal yang lewat tanggal tayang tapi belum `Published`.
  - Mau intip bulan depan? Tinggal ganti angka di sel **`C2` (Bulan)** atau **`E2` (Tahun)**, kalender langsung berubah otomatis!

---

## 3. ⛔ Apa yang JANGAN Diedit?

Supaya sistem tidak error dan data tidak hilang, **tolong jangan ubah**:

- ❌ **Semua sel di tab `📊 Dashboard`** (kecuali dropdown filter Bulan/Tahun/Minggu di baris 4).
- ❌ **Grid tanggal di tab `📅 Kalender Visual`** (cukup ubah dropdown `C2` & `E2`).
- ❌ **Kolom `ID` (tersembunyi), `Tanggal Tayang`, `Hari`, `Brand`, `PIC`, `Topik`, `Link Naskah`, dan `Update Terakhir`** di `📝 Database Konten` (akan muncul peringatan jika tersentuh; biarkan sistem Sync GitHub yang mengisi).
- ❌ **Header filter di `C4`, `E4`, `G4:L4` serta menu corong filter kolom Tanggal** di `📝 Database Konten` (selalu ganti bulan/tahun lewat dropdown `D4` dan `F4`).
- ❌ **Tab `⚙️ Config`** (khusus Ko Jeremy / Hildan untuk atur target dan email).

---

*Kalau ada link naskah tertulis `⚠️ Link 404` atau jadwal belum muncul, kabari Hildan untuk jalankan `📥 Sinkronisasi Naskah dari GitHub`.*
