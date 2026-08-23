# ELS Computer - Service Division Portal

Aplikasi manajemen administrasi servis (Request Parts & Serial Number Management) untuk ELS Computer, dibangun menggunakan **Google Apps Script (GAS)** dan **Google Sheets** sebagai database.

## 🚀 Fitur Utama

- **Role-Based Access Control (RBAC):** Memiliki hak akses terpisah antara `PUSAT` (admin utama) dan `CABANG`.
- **Manajemen Serial Number (SN):** *Generate* otomatis Serial Number berdasarkan kategori komponen, tanggal, dan nomor urut (contoh: `ADP-AC-260820001`). Dilengkapi dengan deteksi anti-duplikasi lintas sesi.
- **Manajemen Master Stok:** Sinkronisasi stok secara otomatis (*auto-deduct / auto-add*) ketika ada SN baru yang masuk, diedit, atau terhapus.
- **Request Cabang:** Sistem pesanan komponen/part secara internal antar cabang dengan pihak pusat.
- **Dashboard & Statistik:** Visualisasi untuk antrean order, jumlah SN masuk, dan *Top 15 Parts* yang paling sering di-restock.
- **UI Ultra Responsif (SWR Caching):** Implementasi *Stale-While-Revalidate* menggunakan *Local Storage* dan `CacheService` milik GAS, membuat transisi halaman instan tanpa *loading spinner* yang mengganggu.

## 🛠️ Tech Stack

- **Backend:** Google Apps Script (`Code.gs`)
- **Database:** Google Sheets
- **Frontend:** Vanilla JavaScript, jQuery, HTML5 (`Index.html`)
- **Styling:** Tailwind CSS (via CDN), FontAwesome
- **Libraries pendukung:**
  - **SweetAlert2:** Untuk notifikasi dan modal.
  - **DataTables:** Untuk fitur *search*, *sort*, dan *pagination* tabel data.
  - **SheetJS (xlsx):** Untuk fitur Export/Import data dari Excel.
  - **jsBarcode & QRCode.js:** Untuk fitur pencetakan Barcode/QR Code SN.

## 📁 Struktur File

Proyek ini sangat ringkas (arsitektur monolitik SPA):
- [`Code.gs`](Code.gs) : Memuat seluruh logika backend, *routing*, pembacaan ke Google Sheets, manajemen sesi, dan *hashing* MD5.
- [`Index.html`](Index.html) : Memuat seluruh antarmuka pengguna (UI) mulai dari Login, *Sidebar*, hingga form tabel interaktif.

## ⚙️ Cara Instalasi & Deployment

Karena ini adalah proyek Google Apps Script, tidak membutuhkan server Node.js.

1. **Siapkan Database (Google Sheets):**
   - Buat satu file Google Spreadsheet.
   - Buat sheet/tab (jika belum ada) dengan nama persis seperti berikut: `users`, `master`, `sn_masuk`, `request_cabang`.
   - Ambil **ID Spreadsheet** (berada di dalam URL antara `/d/` dan `/edit`).
2. **Setup Google Apps Script:**
   - Di Google Drive, klik New > Google Apps Script.
   - Hapus fungsi bawaan.
   - Tempel (*paste*) isi `Code.gs` ke editor. 
   - Ubah nilai konstanta `DB_ID` pada baris ke-6 di `Code.gs` dengan ID Spreadsheet Anda.
   - Tambahkan file HTML baru, beri nama `Index.html`, lalu tempel kode dari `Index.html` Anda.
3. **Deploy:**
   - Klik **Deploy** (di pojok kanan atas) > **New deployment**.
   - Pilih *Select type* > **Web app**.
   - *Execute as*: "Me" (akun Google Anda).
   - *Who has access*: "Anyone".
   - Selesai! Simpan Web App URL yang dihasilkan.

## 🔒 Catatan Keamanan & Skalabilitas

- **Race Condition Protection:** Operasi penting seperti penyimpanan SN dan update Stok memanfaatkan `LockService` (`lock.waitLock()`). Ini memastikan data aman jika beberapa admin menekan tombol "Simpan" di detik yang bersamaan.
- **XSS Protection:** Pengamanan *Cross-Site Scripting* dilakukan dengan fungsi utilitas `escapeHtml()` sebelum data di-*render* ke DOM.
- **Password Hashing:** Sandi pengguna disimpan di spreadsheet dalam bentuk *hash* MD5 (melalui `Utilities.computeDigest`).
