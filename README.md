# ExcelQuiz Pro — Excel Practice Exam Add-in

Add-in Microsoft Excel untuk ujian praktik interaktif dengan sistem penilaian otomatis.

## File yang Disertakan
- manifest.xml      — File manifest untuk registrasi add-in
- taskpane.html     — UI utama add-in (panel soal + scoring)
- commands.html     — Handler perintah ribbon
- README.md         — Panduan ini

## Cara Instalasi

### Langkah 1 — Jalankan web server lokal (HTTPS wajib)
Office Add-in membutuhkan HTTPS (untuk Office.js). Project ini sudah menyediakan server HTTPS di `server.js` dan konfigurasi lewat `.env`.

1) Install dependencies:

```bash
npm install
```

2) Buat sertifikat self-signed (contoh pakai OpenSSL), simpan di folder `certs/`:

```bash
mkdir certs
openssl req -x509 -newkey rsa:2048 -nodes -keyout certs/server.key -out certs/server.crt -days 365 -subj "/CN=localhost"
```

3) Buat file `.env` dari contoh:

```bash
copy .env.example .env
```

4) Jalankan server:

```bash
npm run dev
```

Server akan jalan di `https://localhost:3000` (bisa diubah via `.env`).

### Langkah 2 — Sideload ke Excel Desktop (Windows)
1. File → Options → Trust Center → Trust Center Settings
2. Trusted Add-in Catalogs → masukkan path folder
3. Restart Excel → Insert → My Add-ins → Shared Folder → ExcelQuiz Pro

### Excel Desktop (Mac)
Copy manifest.xml ke: ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
Restart Excel → Insert → My Add-ins

### Excel Online
Insert → Office Add-ins → Upload My Add-in → pilih manifest.xml

## Publish (Hosting + Manifest)
Untuk “publish”, yang penting adalah:
- File add-in (HTML/JS/CSS/asset) harus di-host di domain **HTTPS public** (bukan localhost).
- `manifest.xml` harus diubah supaya semua URL `https://localhost:3000/...` diganti ke domain hosting Anda (misal `https://addin.domain-anda.com/...`).

Langkah ringkas:
1) Deploy file project ini ke hosting HTTPS (misal VPS + Nginx, Azure App Service, Render, dsb).
2) Pastikan URL berikut bisa diakses publik:
   - `/taskpane.html`
   - `/commands.html`
   - `/manifest.xml` (opsional untuk download, tapi berguna)
   - `/assets/*`
3) Duplikasi `manifest.xml` jadi misalnya `manifest.prod.xml`, lalu ganti semua `https://localhost:3000` menjadi URL produksi Anda.
4) Sideload `manifest.prod.xml` untuk test di Excel.
5) Jika ingin rilis resmi (AppSource/tenant deploy), ikuti proses submission Microsoft dan pastikan Add-in memenuhi policy.

## Paket Ujian
- Dasar      : Formula, format, header (15 menit, 6 soal, 100 poin)
- Menengah   : IF, VLOOKUP, COUNTIF, Data Validation (20 menit, 6 soal, 100 poin)
- Lanjutan   : PivotTable, Chart, INDEX-MATCH, Proteksi (25 menit, 6 soal, 100 poin)

## Sistem Nilai
A: 90-100 | B: 80-89 | C: 70-79 (LULUS) | D: 50-69 | E: 0-49
Passing grade: 70%
