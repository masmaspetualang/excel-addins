/**
 * ExcelQuiz Pro — Seed Questions to Supabase (Indonesian Schema)
 * Run: node seed-questions.js
 * Synchronized with the latest 10 Excel UMY LSI requirements
 */
require('dotenv').config();
const { createClient } = require('@supabase/supabase-js');

const SUPABASE_URL = process.env.SUPABASE_URL || '';
const SUPABASE_SERVICE_KEY = process.env.SUPABASE_SERVICE_KEY || '';

if (!SUPABASE_URL || !SUPABASE_SERVICE_KEY) {
  console.error('ERROR: Set SUPABASE_URL and SUPABASE_SERVICE_KEY in .env');
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_KEY);

const WORD_QUESTIONS = [
  { nomor_urut: 1, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Judul Utama', instruksi_tugas: 'Buka dokumen "Penyulingan Minyak Atsiri" yang tersedia. Lakukan pemformatan pada judul utama dokumen tersebut agar menggunakan jenis font Tahoma dengan ukuran 14 pt, berformat Tebal (Bold), dan diatur menjadi Rata Tengah (Center).', langkah_verifikasi: ['Ubah font ke Tahoma', 'Ubah ukuran ke 14 pt', 'Tebalkan judul (Bold)', 'Atur perataan Rata Tengah (Center)'], petunjuk_bantuan: 'Gunakan tab Home → group Font & Paragraph.' },
  { nomor_urut: 2, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Koreksi Kesalahan Ejaan', instruksi_tugas: 'Terdapat kesalahan penulisan istilah ilmiah pada dokumen. Gantilah seluruh kata "arsiri" menjadi "atsiri" menggunakan fitur Find and Replace secara otomatis.', langkah_verifikasi: ['Buka dialog Find & Replace (Ctrl+H)', 'Cari kata "arsiri" dan ganti dengan "atsiri"', 'Klik Replace All'], petunjuk_bantuan: 'Gunakan tab Home → group Editing → Replace.' },
  { nomor_urut: 3, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Kutipan Teks', instruksi_tugas: 'Format teks "A New Dictionary of Chemistry" yang berada di paragraf kedua dengan menggunakan jenis font Trebuchet MS berukuran 11 pt.', langkah_verifikasi: ['Temukan teks "A New Dictionary of Chemistry" di paragraf kedua', 'Ubah jenis font menjadi Trebuchet MS dan ukuran 11 pt'], petunjuk_bantuan: 'Gunakan panel Font di tab Home.' },
  { nomor_urut: 4, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Tulisan Ilmiah (Superscript/Subscript)', instruksi_tugas: 'Format dua teks ilmiah berikut: (1) Cari teks "mm2" dalam dokumen, lalu seleksi HANYA angka "2"-nya dan terapkan efek Superscript sehingga tampil sebagai mm². (2) Cari istilah "Woil", lalu seleksi HANYA kata "oil"-nya dan terapkan efek Subscript sehingga tampil sebagai W_oil (indeks bawah).', langkah_verifikasi: ['Cari teks "mm2" → seleksi HANYA karakter "2" → klik tombol X² (Superscript) di tab Home → Font', 'Cari istilah "Woil" → seleksi HANYA tiga huruf "oil" → klik tombol X₂ (Subscript) di tab Home → Font', 'Pastikan hasilnya: mm² dan W_oil (huruf oil berada di bawah garis teks)'], petunjuk_bantuan: 'Tombol Superscript (X²) dan Subscript (X₂) ada di tab Home, grup Font. Seleksi karakter secara presisi — jangan sampai ikut menyeleksi huruf lain.' },
  { nomor_urut: 5, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Kerapian Paragraf (Indentasi & Spasi)', instruksi_tugas: 'Lakukan dua pengaturan paragraf berikut secara berurutan: PERTAMA, klik di paragraf tubuh pertama (paragraf isi, bukan judul), buka pengaturan Paragraph (klik ikon panah kecil di sudut grup Paragraph), atur "Special: First Line" sebesar 1 inci (2.54 cm). KEDUA, seleksi seluruh dokumen dengan Ctrl+A, terapkan perataan Justify (Ctrl+J), lalu atur Line Spacing menjadi 1.0 (Single).', langkah_verifikasi: ['Klik di dalam paragraf isi pertama (bukan judul)', 'Buka dialog Paragraph Settings → atur Special: First Line → By: 1 inci (2.54 cm)', 'Tekan Ctrl+A untuk seleksi seluruh dokumen', 'Tekan Ctrl+J untuk perataan Justify (Rata Kiri-Kanan)', 'Di dialog Paragraph Settings → Line Spacing: Single (1.0)'], petunjuk_bantuan: 'Jika ukuran First Line Indent dalam cm, masukkan nilai "2.54 cm". Jika dalam inci, masukkan "1"". Line Spacing "Single" setara dengan 1.0.' },
  { nomor_urut: 6, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Subjudul', instruksi_tugas: 'Temukan teks subjudul "B. Metode Umum Penyulingan" dalam dokumen. Lakukan pemformatan pada subjudul tersebut agar menggunakan warna huruf Biru (Blue, pilih warna "Blue" standar dari palet warna) dan berformat Miring (Italic).', langkah_verifikasi: ['Temukan teks subjudul "B. Metode Umum Penyulingan"', 'Seleksi teks tersebut secara presisi', 'Ubah warna huruf menjadi Biru (Blue)', 'Klik tombol Italic (I) untuk memiringkan teks'], petunjuk_bantuan: 'Tombol Font Color (A) dan Italic (I) berada di tab Home grup Font.' },
  { nomor_urut: 7, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Pemberian Highlight Judul Utama', instruksi_tugas: 'Terapkan efek highlight (warna penanda teks) berwarna Kuning (Yellow) pada judul utama dokumen "Penyulingan Minyak Atsiri" di bagian atas.', langkah_verifikasi: ['Seleksi teks judul utama "Penyulingan Minyak Atsiri" di awal dokumen', 'Di tab Home → grup Font → klik tombol Text Highlight Color (ikon pena stabilo) → pilih warna Kuning (Yellow)'], petunjuk_bantuan: 'Pastikan warna highlight yang dipilih adalah Kuning standar.' },
  { nomor_urut: 8, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Penerapan Hyperlink', instruksi_tugas: 'Buatlah sebuah tautan interaktif (Hyperlink) pada teks "Minyak Atsiri" di paragraf pertama agar mengarah ke alamat website: "https://id.wikipedia.org/wiki/Minyak_atsiri".', langkah_verifikasi: ['Temukan teks "Minyak Atsiri" di paragraf pertama dan seleksi teks tersebut', 'Klik kanan → Link / Hyperlink (atau Ctrl+K)', 'Masukkan alamat: https://id.wikipedia.org/wiki/Minyak_atsiri pada kolom Address'], petunjuk_bantuan: 'Gunakan menu Insert → Link atau shortcut Ctrl+K untuk membuat hyperlink.' },
  { nomor_urut: 9, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Pembuatan Tabel Data', instruksi_tugas: 'Buatlah sebuah tabel baru di dokumen yang terdiri dari 3 kolom dengan judul kolom masing-masing: "No", "Nama", dan "Nilai".', langkah_verifikasi: ['Sisipkan tabel baru berukuran minimal 3 kolom x 2 baris', 'Ketik judul kolom: "No", "Nama", "Nilai"'], petunjuk_bantuan: 'Gunakan tab Insert → Table.' },
  { nomor_urut: 10, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Penyisipan Footer Dokumen', instruksi_tugas: 'Sisipkan catatan kaki halaman (Footer) tipe Primary yang berisi nama lengkap Anda sebagai tanda verifikasi identitas.', langkah_verifikasi: ['Buka menu Insert → Footer', 'Ketikkan nama lengkap Anda di area Footer'], petunjuk_bantuan: 'Gunakan tab Insert → Footer → Edit Footer.' }
];

const EXCEL_QUESTIONS = [
  { nomor_urut: 1, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Adventure Works', instruksi_tugas: "Tebalkan 'Adventure Works', ganti font Arial 16.", langkah_verifikasi: ["Cari teks Adventure Works", "Font: Arial", "Size: 16", "Bold"], petunjuk_bantuan: "Cek di sekitar baris 4 atau 5." },
  { nomor_urut: 2, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Alignment Header', instruksi_tugas: "Ubah posisi teks pada cells A4 sampai F4 menjadi rata kiri.", langkah_verifikasi: ["Seleksi A4:F4", "Klik Align Left"], petunjuk_bantuan: "Tab Home → Alignment." },
  { nomor_urut: 3, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Border Tabel', instruksi_tugas: "Buat border tabel tersebut dengan All Border.", langkah_verifikasi: ["Seleksi seluruh tabel", "Klik Borders → All Borders"], petunjuk_bantuan: "Pastikan semua data tercover." },
  { nomor_urut: 4, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Bersihkan Background', instruksi_tugas: "Hilangkan warna background pada kolom F.", langkah_verifikasi: ["Seleksi kolom F", "Fill Color → No Fill"], petunjuk_bantuan: "Ember cat → No Fill." },
  { nomor_urut: 5, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Tambah Kolom No', instruksi_tugas: "Tambah kolom di kiri Item, beri judul 'No' dan isi nomor 1-selesai.", langkah_verifikasi: ["Insert kolom di kiri Item", "Ketik 'No'", "Isi angka urut"], petunjuk_bantuan: "Gunakan AutoFill." },
  { nomor_urut: 6, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Price', instruksi_tugas: "Lebar kolom Price = 20, Format Custom ($).", langkah_verifikasi: ["Column Width: 20", "Format Cells → Custom → ($)"], petunjuk_bantuan: "Klik kanan kolom → Column Width." },
  { nomor_urut: 7, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Rename Sheets', instruksi_tugas: "Sheet1 → 'Database', Sheet2 → 'Januari'.", langkah_verifikasi: ["Double click tab sheet", "Ketik nama baru"], petunjuk_bantuan: "Pastikan ejaan benar." },
  { nomor_urut: 8, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Waktu', instruksi_tugas: "Isi cell A2 dengan waktu (misal 09:00) dan ganti formatnya menjadi Time.", langkah_verifikasi: ["Ketik waktu di A2", "Ganti format ke Time"], petunjuk_bantuan: "Gunakan format jam 24 jam." },
  { nomor_urut: 9, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Total Penjualan', instruksi_tugas: "Isi kolom G (Value) dengan rumus perkalian: =D6*F6.", langkah_verifikasi: ["Klik sel G6", "Ketik =D6*F6", "Copy ke bawah hingga baris 35"], petunjuk_bantuan: "Total = Quantity Ordered * Price." },
  { nomor_urut: 10, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Konfirmasi Selesai', instruksi_tugas: "Ketikkan teks 'LSI UMY' (tanpa tanda kutip) di sel A40 pada sheet pertama (Database/Sheet1) sebagai tanda verifikasi penyelesaian ujian.", langkah_verifikasi: ["Klik sel A40 di sheet pertama (Database/Sheet1)", "Ketik teks LSI UMY"], petunjuk_bantuan: "Gunakan huruf kapital 'LSI UMY' secara presisi." }
];

const PPT_QUESTIONS = [
  { nomor_urut: 1, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Judul Slide Utama', instruksi_tugas: 'Buka file presentasi "Computer organisasi soal" dari direktori soal. Ubah teks judul utama pada Slide 1 dari "COMPUTER ORGANISATION" menjadi "Organisasi Komputer" dengan format font Arial berukuran 44 pt dan Tebal (Bold).', langkah_verifikasi: ['Ubah teks judul Slide 1 menjadi "Organisasi Komputer"', 'Ubah font menjadi Arial', 'Ubah ukuran menjadi 44 pt', 'Terapkan format Tebal (Bold)'], petunjuk_bantuan: 'Seleksi bingkai teks judul untuk mengedit font.' },
  { nomor_urut: 2, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Pembuatan Slide Baru (Title Only)', instruksi_tugas: 'Tambahkan slide baru tepat setelah Slide 1 dengan layout "Title Only" (hanya ada kotak judul, tanpa kotak konten di bawahnya). Caranya: klik Slide 1 di panel kiri → tab Home → klik panah bawah di tombol "New Slide" → pilih layout "Title Only". Kemudian ketikkan teks judul slide baru tersebut: "Siklus Instruksi".', langkah_verifikasi: ['Klik Slide 1 di panel navigasi slide (sebelah kiri)', 'Tab Home → klik tanda panah (▼) di bawah tombol "New Slide"', 'Pilih layout "Title Only" dari daftar (bukan "Title and Content")', 'Ketikkan teks judul pada kotak judul slide baru: Siklus Instruksi'], petunjuk_bantuan: 'Layout "Title Only" hanya memiliki satu kotak teks di bagian atas slide. Pastikan tidak salah memilih "Title and Content" yang memiliki dua kotak teks.' },
  { nomor_urut: 3, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Penyisipan Gambar Ilustrasi', instruksi_tugas: 'Pada slide "Siklus Instruksi" (Slide 2) yang baru saja dibuat, sisipkan gambar ilustrasi bernama "cpu" dari komputer Anda.', langkah_verifikasi: ['Pilih Slide 2', 'Klik menu Insert → Picture', 'Pilih gambar "cpu" dan sisipkan'], petunjuk_bantuan: 'Pastikan kursor aktif pada Slide 2 sebelum menyisipkan gambar.' },
  { nomor_urut: 4, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Skala & Posisi Gambar', instruksi_tugas: 'Pilih gambar "cpu" pada Slide 2. Atur dimensinya menjadi Tinggi 9.5 cm dan Lebar 9 cm. Posisikan gambar tersebut pada koordinat: Horizontal 10 cm dan Vertical 7 cm.', langkah_verifikasi: ['Klik kanan gambar → Format Picture → Size & Properties', 'Atur Height ke 9.5 cm dan Width ke 9 cm', 'Atur Posisi Horizontal ke 10 cm dan Vertical ke 7 cm'], petunjuk_bantuan: 'Gunakan panel Format Picture.' },
  { nomor_urut: 5, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Penekanan Konten (Bold & Warna)', instruksi_tugas: 'Pada slide "RAM" (Slide 5), lakukan dua pemformatan: (1) Cari dan seleksi teks "ALU + CU +REG" (seleksi keseluruhan termasuk spasi dan tanda +), lalu terapkan format Tebal (Bold). (2) Cari dan seleksi kata "ROM" pada slide yang sama, lalu ubah warna hurufnya menjadi Merah (Red) melalui Font Color.', langkah_verifikasi: ['Klik Slide 5 di panel navigasi kiri', 'Klik area teks yang mengandung "ALU + CU +REG" → seleksi tepat teks "ALU + CU +REG" (termasuk semua spasi dan tanda +)', 'Tekan Ctrl+B atau klik tombol Bold (B) di tab Home untuk menebalkan', 'Cari teks "ROM" di slide yang sama → seleksi tepat kata "ROM"', 'Di tab Home → grup Font → klik panah di Font Color → pilih "Red" (Merah standar)'], petunjuk_bantuan: 'Seleksi harus tepat: untuk "ALU + CU +REG", seleksi dari huruf A hingga G terakhir. Warna Red ada di baris Standard Colors di palet warna.' },
  { nomor_urut: 6, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Penekanan Teks Pipelining', instruksi_tugas: 'Pada slide "Performance & Performance measurement" (Slide 4), temukan kata "Pipelining" dan ubah format hurufnya menjadi Tebal (Bold) dan berwarna Merah (Red).', langkah_verifikasi: ['Pilih Slide 4', 'Temukan kata "Pipelining"', 'Tebalkan kata tersebut (Bold)', 'Ubah warna hurufnya menjadi Merah (Red)'], petunjuk_bantuan: 'Gunakan tombol Bold dan Font Color di tab Home.' },
  { nomor_urut: 7, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Pembuatan Tabel Data', instruksi_tugas: 'Pada slide "FUNCTIONAL UNITS OF COMPUTER" (Slide 6), buatlah sebuah tabel baru berukuran 3 kolom x 5 baris.', langkah_verifikasi: ['Buka Slide 6', 'Klik menu Insert → Table', 'Tentukan jumlah kolom = 3 dan baris = 5'], petunjuk_bantuan: 'Gunakan tab Insert → Table.' },
  { nomor_urut: 8, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Organisasi Urutan Slide', instruksi_tugas: 'Pindahkan slide "INPUT UNIT:" (Slide 7) agar bergeser posisi menjadi slide terakhir (setelah slide "CPU").', langkah_verifikasi: ['Pada panel navigasi slide sebelah kiri, pilih slide "INPUT UNIT:"', 'Tarik (drag) slide tersebut ke posisi paling bawah (setelah slide "CPU")'], petunjuk_bantuan: 'Urutan slide akhir: ... → FUNCTIONAL UNITS → CPU → INPUT UNIT.' },
  { nomor_urut: 9, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Judul Slide RAM', instruksi_tugas: 'Pada slide "RAM" (Slide 5), ubah teks judul utama "RAM" menjadi "RAM & ROM".', langkah_verifikasi: ['Klik Slide 5 di panel navigasi kiri', 'Ubah teks judul utama "RAM" menjadi "RAM & ROM"'], petunjuk_bantuan: 'Klik pada kotak judul slide 5 dan ketikkan "& ROM" setelah kata RAM.' },
  { nomor_urut: 10, bobot_nilai: 10, metode_penilaian: 'api', judul_tugas: 'Format Huruf Miring Konten Slide', instruksi_tugas: 'Pada slide "Performance & Performance measurement" (Slide 4), temukan kata "superscalar" (atau "superscalar operation") dan ubah format hurufnya menjadi Miring (Italic).', langkah_verifikasi: ['Pilih Slide 4', 'Temukan kata "superscalar"', 'Seleksi kata tersebut secara presisi', 'Klik tombol Italic (I) di tab Home (atau tekan Ctrl+I)'], petunjuk_bantuan: 'Gunakan tombol Italic di tab Home grup Font.' }
];

async function seed() {
  console.log('🌱 Seeding questions to Supabase (butir_soal table)...\n');

  // Clear existing in butir_soal
  const { error: delErr } = await supabase.from('butir_soal').delete().neq('id_soal', 0);
  if (delErr) console.warn('Delete warning:', delErr.message);

  let total = 0;

  // Word
  for (const q of WORD_QUESTIONS) {
    const { error } = await supabase.from('butir_soal').insert({ ...q, jenis_aplikasi: 'word', kategori_ujian: 'praktik' });
    if (error) console.error('Word Q' + q.nomor_urut + ':', error.message);
    else { process.stdout.write('📝'); total++; }
  }

  // Excel
  for (const q of EXCEL_QUESTIONS) {
    const { error } = await supabase.from('butir_soal').insert({ ...q, jenis_aplikasi: 'excel', kategori_ujian: 'praktik' });
    if (error) console.error('Excel Q' + q.nomor_urut + ':', error.message);
    else { process.stdout.write('📊'); total++; }
  }

  // PowerPoint
  for (const q of PPT_QUESTIONS) {
    const { error } = await supabase.from('butir_soal').insert({ ...q, jenis_aplikasi: 'ppt', kategori_ujian: 'praktik' });
    if (error) console.error('PPT Q' + q.nomor_urut + ':', error.message);
    else { process.stdout.write('📽'); total++; }
  }

  console.log('\n\n✅ Done! ' + total + ' questions seeded successfully into butir_soal.\n');
}

seed().catch(console.error);
