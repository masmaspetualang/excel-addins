/* ═══════════════════════════════════════════════
   EXAM DATA
═══════════════════════════════════════════════ */
const EXAMS = {
  basic: {
    name: "Dasar",
    duration: 15 * 60,
    passingScore: 70,
    tasks: [
      {
        id: "b1", points: 15,
        title: "Memasukkan Data Sederhana",
        desc: "Buat tabel data penjualan sederhana di sheet yang aktif.",
        steps: [
          "Klik cell <code>A1</code> dan ketik <code>Produk</code>",
          "Klik cell <code>B1</code> dan ketik <code>Harga</code>",
          "Klik cell <code>C1</code> dan ketik <code>Qty</code>",
          "Di baris 2–5, masukkan minimal 4 baris data produk"
        ],
        hint: "Contoh data: Apel, 5000, 10 — lalu Mangga, 8000, 5, dst.",
        check: checkB1
      },
      {
        id: "b2", points: 15,
        title: "Formula Perkalian (Subtotal)",
        desc: "Hitung subtotal (Harga × Qty) menggunakan formula Excel.",
        steps: [
          "Klik cell <code>D1</code> dan ketik header <code>Subtotal</code>",
          "Klik cell <code>D2</code>",
          "Ketik formula: <code>=B2*C2</code> lalu tekan Enter",
          "Copy formula ke <code>D3</code>, <code>D4</code>, <code>D5</code>"
        ],
        hint: "Seleksi D2, drag fill handle (kotak kecil di pojok kanan bawah) ke D5.",
        check: checkB2
      },
      {
        id: "b3", points: 20,
        title: "Fungsi SUM & AVERAGE",
        desc: "Hitung total dan rata-rata menggunakan fungsi SUM dan AVERAGE.",
        steps: [
          "Klik cell <code>B7</code>, ketik label <code>Total</code>",
          "Klik cell <code>C7</code>, masukkan formula <code>=SUM(C2:C5)</code>",
          "Klik cell <code>B8</code>, ketik label <code>Rata-rata Harga</code>",
          "Klik cell <code>C8</code>, masukkan formula <code>=AVERAGE(B2:B5)</code>"
        ],
        hint: "Fungsi SUM dan AVERAGE adalah fungsi dasar Excel. Pastikan range-nya benar!",
        check: checkB3
      },
      {
        id: "b4", points: 15,
        title: "Format Currency",
        desc: "Format kolom Harga dan Subtotal dengan format mata uang.",
        steps: [
          "Seleksi range <code>B2:B5</code>",
          "Klik kanan → Format Cells → Number → Currency",
          "Pilih simbol <code>Rp</code> atau format Accounting",
          "Lakukan hal sama untuk range <code>D2:D5</code>"
        ],
        hint: "Shortcut format currency: Ctrl+Shift+4. Atau gunakan tombol $ di toolbar.",
        check: checkB4
      },
      {
        id: "b5", points: 15,
        title: "Format Header Bold & Warna",
        desc: "Beri formatting pada baris header agar lebih mudah dibaca.",
        steps: [
          "Seleksi baris header <code>A1:D1</code>",
          "Tekan <code>Ctrl+B</code> untuk Bold",
          "Beri warna background: Home → Fill Color → pilih warna",
          "Ubah warna teks menjadi putih jika background gelap"
        ],
        hint: "Gunakan warna yang kontras agar tabel lebih profesional.",
        check: checkB5
      },
      {
        id: "b6", points: 20,
        title: "Freeze Panes & AutoFit",
        desc: "Bekukan baris header dan rapikan lebar kolom.",
        steps: [
          "Klik cell <code>A2</code>",
          "Klik menu View → Freeze Panes → Freeze Panes",
          "Seleksi semua kolom A–D",
          "Klik kanan pada header kolom → Column Width → AutoFit"
        ],
        hint: "Freeze Panes membuat header tetap terlihat saat scroll ke bawah.",
        check: checkB6
      }
    ]
  },

  intermediate: {
    name: "Menengah",
    duration: 20 * 60,
    passingScore: 70,
    tasks: [
      {
        id: "m1", points: 15,
        title: "Tabel Data dengan Named Range",
        desc: "Buat dataset karyawan dan definisikan named range.",
        steps: [
          "Buat header di A1: <code>Nama</code>, B1: <code>Divisi</code>, C1: <code>Gaji</code>, D1: <code>Bonus</code>",
          "Masukkan minimal 5 data karyawan di baris 2–6",
          "Seleksi range <code>A1:D6</code>",
          "Di Name Box (pojok kiri atas), ketik <code>DataKaryawan</code> lalu Enter"
        ],
        hint: "Named Range memudahkan referensi ke data tanpa harus mengingat koordinat sel.",
        check: checkM1
      },
      {
        id: "m2", points: 20,
        title: "Fungsi IF Bersarang",
        desc: "Hitung bonus menggunakan fungsi IF berdasarkan divisi.",
        steps: [
          "Klik cell <code>D2</code>",
          "Masukkan formula: <code>=IF(B2=\"IT\",C2*0.15,IF(B2=\"Sales\",C2*0.12,C2*0.10))</code>",
          "Tekan Enter dan copy formula ke D3:D6",
          "Verifikasi hasilnya sesuai logika: IT=15%, Sales=12%, lainnya=10%"
        ],
        hint: "IF bersarang: =IF(kondisi1, nilai_jika_benar, IF(kondisi2, nilai2, nilai_default))",
        check: checkM2
      },
      {
        id: "m3", points: 20,
        title: "VLOOKUP",
        desc: "Gunakan VLOOKUP untuk mencari data dari tabel referensi.",
        steps: [
          "Di sheet yang sama, buat tabel referensi di <code>F1:G5</code>: kolom Divisi & Grade",
          "Buat header <code>Grade</code> di cell <code>E1</code>",
          "Di <code>E2</code>, gunakan: <code>=VLOOKUP(B2,$F$1:$G$5,2,FALSE)</code>",
          "Copy ke E3:E6"
        ],
        hint: "$ (dollar sign) di F$1:$G$5 mengunci referensi agar tidak bergeser saat di-copy.",
        check: checkM3
      },
      {
        id: "m4", points: 15,
        title: "COUNTIF & SUMIF",
        desc: "Buat ringkasan statistik menggunakan COUNTIF dan SUMIF.",
        steps: [
          "Di cell <code>I1</code> ketik <code>Divisi IT</code>",
          "Di <code>I2</code>: <code>=COUNTIF(B2:B6,\"IT\")</code> — hitung jumlah karyawan IT",
          "Di <code>I3</code>: <code>=SUMIF(B2:B6,\"IT\",C2:C6)</code> — total gaji karyawan IT",
          "Ulangi untuk divisi Sales di baris I4–I5"
        ],
        hint: "COUNTIF menghitung sel yang memenuhi kriteria; SUMIF menjumlahkan sel yang memenuhi kriteria.",
        check: checkM4
      },
      {
        id: "m5", points: 15,
        title: "Data Validation (Dropdown)",
        desc: "Tambahkan validasi data berupa dropdown list untuk kolom Divisi.",
        steps: [
          "Seleksi range <code>B2:B6</code>",
          "Klik Data → Data Validation",
          "Allow: <code>List</code>, masukkan Source: <code>IT,Sales,HR,Finance,Operations</code>",
          "Klik OK dan uji coba dropdown di salah satu sel"
        ],
        hint: "Data Validation mencegah input yang tidak valid dan membuat entry lebih mudah.",
        check: checkM5
      },
      {
        id: "m6", points: 15,
        title: "Conditional Formatting",
        desc: "Highlight gaji di atas rata-rata dengan warna berbeda.",
        steps: [
          "Seleksi range <code>C2:C6</code>",
          "Klik Home → Conditional Formatting → New Rule",
          "Pilih 'Format cells that contain' → Cell Value → Greater Than → <code>=AVERAGE(C2:C6)</code>",
          "Atur fill color hijau dan klik OK"
        ],
        hint: "Conditional Formatting secara otomatis mengubah tampilan sel berdasarkan nilainya.",
        check: checkM6
      }
    ]
  },

  advanced: {
    name: "Lanjutan",
    duration: 25 * 60,
    passingScore: 70,
    tasks: [
      {
        id: "a1", points: 15,
        title: "PivotTable Dasar",
        desc: "Buat PivotTable dari dataset penjualan untuk analisis ringkas.",
        steps: [
          "Buat data di Sheet1: Tanggal, Produk, Kategori, Wilayah, Penjualan (min 10 baris)",
          "Klik salah satu sel di dalam data, lalu Insert → PivotTable",
          "Pilih 'New Worksheet' dan klik OK",
          "Di panel PivotTable Fields: drag Produk ke Rows, Penjualan ke Values"
        ],
        hint: "PivotTable otomatis merangkum dan mengagregasi data besar dengan mudah.",
        check: checkA1
      },
      {
        id: "a2", points: 15,
        title: "Chart / Grafik",
        desc: "Buat grafik kolom dari data PivotTable atau data asli.",
        steps: [
          "Seleksi data ringkasan (2 kolom: label & nilai)",
          "Klik Insert → Charts → Column Chart → Clustered Column",
          "Tambahkan judul chart dengan mengklik Chart Title",
          "Pindahkan chart ke posisi yang rapi (tidak menutupi data)"
        ],
        hint: "Pilih tipe chart yang sesuai data: kolom untuk perbandingan, line untuk tren waktu.",
        check: checkA2
      },
      {
        id: "a3", points: 15,
        title: "INDEX & MATCH",
        desc: "Gunakan INDEX-MATCH sebagai alternatif VLOOKUP yang lebih fleksibel.",
        steps: [
          "Buat tabel referensi baru (Kode Produk & Nama Produk) di area kosong",
          "Di sel target, masukkan: <code>=INDEX(B1:B10,MATCH(E1,A1:A10,0))</code>",
          "Ganti range sesuai data Anda (kolom Kode & Nama)",
          "Uji dengan beberapa kode produk berbeda"
        ],
        hint: "INDEX-MATCH lebih powerful dari VLOOKUP karena bisa mencari ke kiri dan tidak bergantung urutan kolom.",
        check: checkA3
      },
      {
        id: "a4", points: 20,
        title: "Fungsi Teks & Tanggal",
        desc: "Manipulasi data teks dan tanggal menggunakan fungsi Excel.",
        steps: [
          "Di kolom baru, gunakan <code>=TEXT(A2,\"DD/MM/YYYY\")</code> untuk format tanggal",
          "Gunakan <code>=UPPER(B2)</code> atau <code>=PROPER(B2)</code> untuk format teks",
          "Gunakan <code>=CONCATENATE(B2,\" - \",C2)</code> atau <code>=B2&\" - \"&C2</code>",
          "Gunakan <code>=TODAY()</code> atau <code>=DATEDIF(A2,TODAY(),\"D\")</code>"
        ],
        hint: "Fungsi teks dan tanggal sangat berguna untuk membersihkan dan mentransformasi data.",
        check: checkA4
      },
      {
        id: "a5", points: 20,
        title: "Sparklines & Conditional Formatting Lanjutan",
        desc: "Tambahkan sparklines dan color scale untuk visualisasi inline.",
        steps: [
          "Seleksi range data numerik (min 5 kolom berurutan)",
          "Klik Insert → Sparklines → Line, pilih lokasi output",
          "Untuk warna scale: seleksi data → Conditional Formatting → Color Scales",
          "Pilih skema warna Hijau-Kuning-Merah"
        ],
        hint: "Sparklines adalah mini-chart dalam satu sel, ideal untuk menampilkan tren singkat.",
        check: checkA5
      },
      {
        id: "a6", points: 15,
        title: "Proteksi Sheet & Workbook",
        desc: "Lindungi sheet dengan password agar data tidak diubah sembarangan.",
        steps: [
          "Unlock sel yang boleh diedit: seleksi sel → Format Cells → Protection → uncheck Locked",
          "Klik Review → Protect Sheet",
          "Masukkan password (misal: <code>excel123</code>) dan konfirmasi",
          "Uji: coba klik sel yang terkunci — harus muncul pesan error"
        ],
        hint: "Proteksi sheet berguna untuk form/template yang dibagikan ke pengguna lain.",
        check: checkA6
      }
    ]
  }
};

/* ═══════════════════════════════════════════════
   WORD EXAM DATA
═══════════════════════════════════════════════ */
const WORD_EXAMS = {
  basic: {
    name: "Word Dasar",
    duration: 15 * 60,
    passingScore: 70,
    tasks: [
      {
        id: "w1", points: 20,
        title: "Mengetik & Format Teks",
        desc: "Ketik kalimat pembuka dan atur format hurufnya.",
        steps: [
          "Ketik: <code>Laporan Penjualan Tahunan</code>",
          "Seleksi teks tersebut, buat menjadi <strong>Bold</strong>",
          "Ubah ukuran font menjadi 16pt",
          "Atur paragraf menjadi <strong>Rata Tengah (Center)</strong>"
        ],
        hint: "Gunakan menu Home atau shortcut Ctrl+B untuk Bold dan Ctrl+E untuk Center.",
        check: checkW1
      },
      {
        id: "w2", points: 20,
        title: "Menambahkan Daftar (Bullet Points)",
        desc: "Buat daftar poin-poin penting menggunakan Bullets.",
        steps: [
          "Ketik daftar 3 nama produk di baris baru",
          "Seleksi ketiga baris tersebut",
          "Klik ikon <strong>Bullets</strong> di menu Home"
        ],
        hint: "Klik ikon titik-titik (Bullets) di bagian Paragraph.",
        check: checkW2
      },
      {
        id: "w3", points: 30,
        title: "Pewarnaan & Garis Bawah",
        desc: "Beri penekanan pada kata tertentu.",
        steps: [
          "Ketik kalimat: <code>Dokumen ini bersifat rahasia.</code>",
          "Ubah warna teks 'rahasia' menjadi <strong>Merah</strong>",
          "Beri garis bawah (Underline) pada kata 'rahasia'"
        ],
        hint: "Gunakan Font Color (ikon A) dan Underline (Ctrl+U).",
        check: checkW3
      }
    ]
  },
  intermediate: {
    name: "Word Menengah",
    duration: 20 * 60,
    passingScore: 75,
    tasks: [
      {
        id: "w4", points: 25,
        title: "Membuat Tabel",
        desc: "Masukkan tabel untuk menyajikan data.",
        steps: [
          "Insert tabel dengan ukuran <strong>3 kolom x 4 baris</strong>",
          "Isi baris pertama dengan: No, Nama, Keterangan",
          "Beri warna background (Shading) pada baris pertama"
        ],
        hint: "Menu Insert -> Table.",
        check: checkW4
      },
      {
        id: "w5", points: 25,
        title: "Header & Footer",
        desc: "Tambahkan informasi di bagian atas dan bawah halaman.",
        steps: [
          "Klik menu Insert -> Header, pilih gaya 'Blank'",
          "Ketik nama Anda di Header",
          "Klik Insert -> Page Number -> Bottom of Page"
        ],
        hint: "Double klik bagian atas kertas untuk membuka Header.",
        check: checkW5
      }
    ]
  }
};

/* ═══════════════════════════════════════════════
   STATE
═══════════════════════════════════════════════ */
let state = {
  host: null,       // 'Excel' or 'Word'
  examKey: 'basic',
  exam: null,
  currentIdx: 0,
  confirmed: [],
  scores: [],
  timerInterval: null,
  timeLeft: 0,
  started: false,
  finished: false
};

/* ═══════════════════════════════════════════════
   OFFICE INIT
═══════════════════════════════════════════════ */
Office.onReady((info) => {
  state.host = info.host === Office.HostType.Excel ? 'Excel' : (info.host === Office.HostType.Word ? 'Word' : null);
  
  if (state.host) {
    initApp();
  }
});

function initApp() {
  const select = document.getElementById('exam-select');
  
  // Kosongkan dan isi dropdown berdasarkan Host
  select.innerHTML = '';
  const currentData = state.host === 'Excel' ? EXAMS : WORD_EXAMS;
  
  for (let key in currentData) {
    const opt = document.createElement('option');
    opt.value = key;
    opt.textContent = `📗 ${currentData[key].name} (${Math.round(currentData[key].duration/60)} menit)`;
    select.appendChild(opt);
  }

  select.addEventListener('change', updateExamInfo);
  updateExamInfo();
  
  // Update UI Title
  document.querySelector('.header-title').textContent = `ExcelQuiz Pro - ${state.host}`;
}

function updateExamInfo() {
  const key = document.getElementById('exam-select').value;
  const currentData = state.host === 'Excel' ? EXAMS : WORD_EXAMS;
  const exam = currentData[key];
  if (!exam) return;
  
  document.getElementById('info-total-q').textContent = exam.tasks.length;
  document.getElementById('info-duration').textContent = Math.round(exam.duration/60) + 'm';
  document.getElementById('info-points').textContent = exam.tasks.reduce((s,t)=>s+t.points,0);
  document.getElementById('info-passing').textContent = exam.passingScore;
}

/* ═══════════════════════════════════════════════
   START EXAM
═══════════════════════════════════════════════ */
document.getElementById('btn-start').addEventListener('click', startExam);

async function startExam() {
  state.examKey = document.getElementById('exam-select').value;
  const currentData = state.host === 'Excel' ? EXAMS : WORD_EXAMS;
  state.exam = currentData[state.examKey];
  state.currentIdx = 0;
  state.confirmed = new Array(state.exam.tasks.length).fill(false);
  state.scores = new Array(state.exam.tasks.length).fill(0);
  state.timeLeft = state.exam.duration;
  state.started = true;
  state.finished = false;

  document.getElementById('header-sub').textContent = 'Ujian ' + state.exam.name;
  document.getElementById('timer-bar').style.display = 'flex';
  document.getElementById('progress-outer').style.display = 'block';

  startTimer();
  buildDotsNav();
  showTask(0);
  showScreen('task');

  // Setup workbook
  try {
    await Excel.run(async (context) => {
      const ws = context.workbook.worksheets.getActiveWorksheet();
      ws.name = "UjianExcel";
      await context.sync();
    });
  } catch(e) {}

  showToast('Ujian dimulai! Selamat mengerjakan 🚀', 'info');
}

/* ═══════════════════════════════════════════════
   TIMER
═══════════════════════════════════════════════ */
function startTimer() {
  clearInterval(state.timerInterval);
  state.timerInterval = setInterval(() => {
    state.timeLeft--;
    updateTimerDisplay();
    if (state.timeLeft <= 0) {
      clearInterval(state.timerInterval);
      showToast('⏰ Waktu habis! Ujian selesai.', 'error');
      setTimeout(() => finishExam(), 1500);
    }
  }, 1000);
}

function updateTimerDisplay() {
  const m = Math.floor(state.timeLeft / 60);
  const s = state.timeLeft % 60;
  const display = `${String(m).padStart(2,'0')}:${String(s).padStart(2,'0')}`;
  const el = document.getElementById('timer-display');
  el.textContent = display;
  el.className = 'timer-value' + (state.timeLeft < 120 ? ' urgent' : '');
}

/* ═══════════════════════════════════════════════
   SHOW TASK
═══════════════════════════════════════════════ */
function showTask(idx) {
  state.currentIdx = idx;
  const task = state.exam.tasks[idx];
  const total = state.exam.tasks.length;

  document.getElementById('task-badge').textContent = `Soal ${idx+1} / ${total}`;
  document.getElementById('task-score-badge').textContent = `+${task.points} poin`;
  document.getElementById('task-number').textContent = `TUGAS ${String(idx+1).padStart(2,'0')}`;
  document.getElementById('task-title').textContent = task.title;
  document.getElementById('task-desc').textContent = task.desc;

  // Steps
  const stepsEl = document.getElementById('task-steps');
  stepsEl.innerHTML = task.steps.map((s, i) => `
    <div class="step-item">
      <div class="step-num">${i+1}</div>
      <div class="step-text">${s}</div>
    </div>
  `).join('');

  // Hint
  if (task.hint) {
    document.getElementById('task-hint').style.display = 'flex';
    document.getElementById('task-hint-text').textContent = task.hint;
  } else {
    document.getElementById('task-hint').style.display = 'none';
  }

  // Confirm button state
  const confirmBtn = document.getElementById('btn-confirm');
  if (state.confirmed[idx]) {
    confirmBtn.innerHTML = '✓ &nbsp;Sudah Dikonfirmasi';
    confirmBtn.classList.add('done');
    confirmBtn.disabled = true;
  } else {
    confirmBtn.innerHTML = '✓ &nbsp;Saya Sudah Selesai';
    confirmBtn.classList.remove('done');
    confirmBtn.disabled = false;
  }

  // Nav buttons
  document.getElementById('btn-prev').disabled = idx === 0;
  document.getElementById('btn-next').textContent = idx === total-1
    ? 'Selesai & Nilai →' : 'Selanjutnya →';
  document.getElementById('btn-next').onclick = idx === total-1
    ? finishExam : nextTask;

  // Update progress
  const done = state.confirmed.filter(Boolean).length;
  document.getElementById('progress-bar').style.width = (done / total * 100) + '%';

  updateDots();
}

/* ═══════════════════════════════════════════════
   DOTS NAV
═══════════════════════════════════════════════ */
function buildDotsNav() {
  const nav = document.getElementById('dots-nav');
  nav.innerHTML = state.exam.tasks.map((_, i) =>
    `<div class="dot" id="dot-${i}" onclick="showTask(${i})"></div>`
  ).join('');
}

function updateDots() {
  state.exam.tasks.forEach((_, i) => {
    const dot = document.getElementById('dot-' + i);
    if (!dot) return;
    dot.className = 'dot';
    if (i === state.currentIdx) dot.classList.add('current');
    else if (state.confirmed[i]) dot.classList.add('done');
  });
}

/* ═══════════════════════════════════════════════
   NAVIGATION
═══════════════════════════════════════════════ */
function prevTask() {
  if (state.currentIdx > 0) showTask(state.currentIdx - 1);
}

function nextTask() {
  const total = state.exam.tasks.length;
  if (state.currentIdx < total - 1) showTask(state.currentIdx + 1);
}

/* ═══════════════════════════════════════════════
   CONFIRM TASK
═══════════════════════════════════════════════ */
async function confirmTask() {
  const idx = state.currentIdx;
  if (state.confirmed[idx]) return;

  state.confirmed[idx] = true;

  const btn = document.getElementById('btn-confirm');
  btn.innerHTML = '✓ &nbsp;Sudah Dikonfirmasi';
  btn.classList.add('done');
  btn.disabled = true;

  updateDots();

  const done = state.confirmed.filter(Boolean).length;
  const total = state.exam.tasks.length;
  document.getElementById('progress-bar').style.width = (done / total * 100) + '%';

  showToast(`Soal ${idx+1} dikonfirmasi! ✓`, 'success');

  if (done === total) {
    setTimeout(() => showToast('Semua soal dikonfirmasi! Klik Selesai & Nilai →', 'info'), 1500);
  }
}

/* ═══════════════════════════════════════════════
   FINISH & SCORE
═══════════════════════════════════════════════ */
async function finishExam() {
  if (state.finished) return;
  state.finished = true;
  clearInterval(state.timerInterval);

  const overlay = document.getElementById('scoring-overlay');
  overlay.classList.add('show');

  const tasks = state.exam.tasks;
  const results = [];

  for (let i = 0; i < tasks.length; i++) {
    document.getElementById('scoring-step-text').textContent = `Memeriksa soal ${i+1} dari ${tasks.length}...`;
    await new Promise(r => setTimeout(r, 400));

    let pts = 0;
    let detail = '';
    let status = 'fail';

    if (state.confirmed[i]) {
      try {
        const result = await tasks[i].check();
        pts = result.score;
        detail = result.detail;
        status = pts >= tasks[i].points ? 'pass' : (pts > 0 ? 'partial' : 'fail');
      } catch(e) {
        pts = Math.round(tasks[i].points * 0.5); // partial credit if can't verify
        detail = 'Dikonfirmasi (verifikasi parsial)';
        status = 'partial';
      }
    } else {
      pts = 0;
      detail = 'Tidak dikonfirmasi';
      status = 'fail';
    }

    state.scores[i] = pts;
    results.push({ task: tasks[i], pts, detail, status });
  }

  await new Promise(r => setTimeout(r, 600));
  overlay.classList.remove('show');

  showResults(results);
}

/* ═══════════════════════════════════════════════
   CHECK FUNCTIONS (Office.js verification)
═══════════════════════════════════════════════ */

async function checkB1() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const headers = sheet.getRange("A1:C1");
    headers.load("values");
    const dataRange = sheet.getRange("A2:A5");
    dataRange.load("values");
    await context.sync();

    const [a,b,c] = headers.values[0].map(v => String(v).toLowerCase().trim());
    let score = 0;
    let details = [];

    if (a.includes('produk') || a.includes('nama')) { score += 5; details.push('Header A1 ✓'); }
    if (b.includes('harga') || b.includes('price')) { score += 5; details.push('Header B1 ✓'); }
    if (c.includes('qty') || c.includes('jumlah') || c.includes('stok')) { score += 5; details.push('Header C1 ✓'); }
    const hasData = dataRange.values.filter(r => r[0] && String(r[0]).trim() !== '').length;
    if (hasData >= 2) { score = Math.min(15, score + (hasData >= 4 ? 0 : 0)); details.push(`${hasData} baris data`); }
    if (hasData >= 4) score = Math.max(score, 12);

    return { score: Math.min(15, score), detail: details.join(', ') || 'Tidak ada data ditemukan' };
  });
}

async function checkB2() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const d1 = sheet.getRange("D1");
    const formulas = sheet.getRange("D2:D5");
    d1.load("values");
    formulas.load("formulas,values");
    await context.sync();

    let score = 0;
    let details = [];
    const header = String(d1.values[0][0]).toLowerCase();
    if (header.includes('subtotal') || header.includes('total')) { score += 3; details.push('Header D1 ✓'); }

    let formulaCount = 0;
    formulas.formulas.forEach(row => {
      const f = String(row[0]).toLowerCase();
      if (f.includes('*') || f.includes('b') && f.includes('c')) formulaCount++;
    });

    if (formulaCount >= 4) { score += 12; details.push('4 formula perkalian ✓'); }
    else if (formulaCount >= 2) { score += 8; details.push(`${formulaCount} formula ditemukan`); }
    else if (formulaCount >= 1) { score += 4; }

    return { score: Math.min(15, score), detail: details.join(', ') || `${formulaCount} formula ditemukan` };
  });
}

async function checkB3() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A7:C8");
    range.load("formulas,values");
    await context.sync();

    let score = 0;
    let details = [];
    const allFormulas = range.formulas.flat().join(' ').toUpperCase();

    if (allFormulas.includes('SUM')) { score += 10; details.push('SUM ✓'); }
    if (allFormulas.includes('AVERAGE')) { score += 10; details.push('AVERAGE ✓'); }

    return { score: Math.min(20, score), detail: details.join(', ') || 'Formula tidak ditemukan di area B7:C8' };
  });
}

async function checkB4() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const b2 = sheet.getRange("B2");
    b2.load("numberFormat");
    await context.sync();

    const fmt = String(b2.numberFormat).toLowerCase();
    const hasCurrency = fmt.includes('rp') || fmt.includes('$') || fmt.includes('idr') ||
                        fmt.includes('#,##0') || fmt.includes('accounting');

    return {
      score: hasCurrency ? 15 : (state.confirmed[3] ? 8 : 0),
      detail: hasCurrency ? 'Format currency terdeteksi ✓' : 'Format currency tidak terdeteksi (kredit parsial)'
    };
  });
}

async function checkB5() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const headerRange = sheet.getRange("A1:D1");
    headerRange.load("format/font/bold,format/fill/color");
    await context.sync();

    const isBold = headerRange.format.font.bold;
    const fillColor = headerRange.format.fill.color;
    const hasColor = fillColor && fillColor !== '#FFFFFF' && fillColor !== 'white' && fillColor !== '' && fillColor !== null;

    let score = 0;
    let details = [];
    if (isBold) { score += 8; details.push('Bold ✓'); }
    if (hasColor) { score += 7; details.push('Warna background ✓'); }

    return { score: Math.min(15, score), detail: details.join(', ') || 'Formatting header belum terdeteksi' };
  });
}

async function checkB6() {
  // Freeze panes and column width hard to verify via API, give partial credit
  return {
    score: state.confirmed[5] ? 15 : 0,
    detail: state.confirmed[5] ? 'Dikonfirmasi oleh peserta ✓' : 'Tidak dikonfirmasi'
  };
}

async function checkM1() {
  return await Excel.run(async (context) => {
    const wb = context.workbook;
    const names = wb.names;
    names.load("items");
    await context.sync();

    const namedRanges = names.items.map(n => n.name.toLowerCase());
    const hasNamed = namedRanges.some(n => n.includes('karyawan') || n.includes('data'));

    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const headers = sheet.getRange("A1:D1");
    headers.load("values");
    await context.sync();

    const vals = headers.values[0].map(v => String(v).toLowerCase());
    const hasHeaders = vals.some(v => v.includes('nama')) && vals.some(v => v.includes('gaji'));

    let score = 0;
    if (hasHeaders) score += 8;
    if (hasNamed) score += 7;

    return { score: Math.min(15, score), detail: `Headers: ${hasHeaders?'✓':'✗'}, Named Range: ${hasNamed?'✓':'✗'}` };
  });
}

async function checkM2() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const formulas = sheet.getRange("D2:D6");
    formulas.load("formulas");
    await context.sync();

    const allF = formulas.formulas.flat().join(' ').toUpperCase();
    const hasIF = (allF.match(/IF/g) || []).length;
    const hasNested = hasIF >= 2;

    let score = 0;
    if (hasIF >= 1) score += 10;
    if (hasNested) score += 10;

    return { score: Math.min(20, score), detail: `IF ditemukan: ${hasIF}x, Bersarang: ${hasNested?'✓':'✗'}` };
  });
}

async function checkM3() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("E2:E6");
    range.load("formulas");
    await context.sync();

    const allF = range.formulas.flat().join(' ').toUpperCase();
    const hasVlookup = allF.includes('VLOOKUP');
    const hasAbsolute = allF.includes('$');

    let score = 0;
    if (hasVlookup) score += 14;
    if (hasAbsolute) score += 6;

    return { score: Math.min(20, score), detail: `VLOOKUP: ${hasVlookup?'✓':'✗'}, Absolute ref: ${hasAbsolute?'✓':'✗'}` };
  });
}

async function checkM4() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("I1:I5");
    range.load("formulas");
    await context.sync();

    const allF = range.formulas.flat().join(' ').toUpperCase();
    const hasCOUNTIF = allF.includes('COUNTIF');
    const hasSUMIF = allF.includes('SUMIF');

    let score = 0;
    if (hasCOUNTIF) score += 7;
    if (hasSUMIF) score += 8;

    return { score: Math.min(15, score), detail: `COUNTIF: ${hasCOUNTIF?'✓':'✗'}, SUMIF: ${hasSUMIF?'✓':'✗'}` };
  });
}

async function checkM5() {
  // Data validation hard to verify via basic API
  return {
    score: state.confirmed[4] ? 12 : 0,
    detail: state.confirmed[4] ? 'Dikonfirmasi peserta — validasi data diterapkan ✓' : 'Tidak dikonfirmasi'
  };
}

async function checkM6() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("C2:C6");
    range.load("conditionalFormats/items");
    await context.sync();

    const hasCF = range.conditionalFormats.items && range.conditionalFormats.items.length > 0;
    return {
      score: hasCF ? 15 : (state.confirmed[5] ? 8 : 0),
      detail: hasCF ? 'Conditional Formatting terdeteksi ✓' : 'CF tidak terdeteksi (kredit parsial)'
    };
  });
}

async function checkA1() {
  return await Excel.run(async (context) => {
    const wb = context.workbook;
    const sheets = wb.worksheets;
    sheets.load("items/name");
    await context.sync();

    const hasPivot = sheets.items.some(s => s.name.toLowerCase().includes('pivot') || s.name.includes('Sheet'));
    const sheetCount = sheets.items.length;

    return {
      score: sheetCount >= 2 ? 15 : (state.confirmed[0] ? 10 : 0),
      detail: `Sheet count: ${sheetCount} (PivotTable biasanya di sheet baru)`
    };
  });
}

async function checkA2() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const charts = sheet.charts;
    charts.load("items/name,items/chartType");
    await context.sync();

    const chartCount = charts.items.length;
    return {
      score: chartCount >= 1 ? 15 : (state.confirmed[1] ? 6 : 0),
      detail: chartCount >= 1 ? `${chartCount} chart ditemukan ✓` : 'Chart tidak ditemukan'
    };
  });
}

async function checkA3() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRange();
    used.load("formulas");
    await context.sync();

    const allF = used.formulas.flat().join(' ').toUpperCase();
    const hasINDEX = allF.includes('INDEX');
    const hasMATCH = allF.includes('MATCH');

    let score = 0;
    if (hasINDEX) score += 8;
    if (hasMATCH) score += 7;

    return { score: Math.min(15, score), detail: `INDEX: ${hasINDEX?'✓':'✗'}, MATCH: ${hasMATCH?'✓':'✗'}` };
  });
}

async function checkA4() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const used = sheet.getUsedRange();
    used.load("formulas");
    await context.sync();

    const allF = used.formulas.flat().join(' ').toUpperCase();
    const hasText = allF.includes('TEXT') || allF.includes('UPPER') || allF.includes('LOWER') || allF.includes('PROPER') || allF.includes('CONCAT');
    const hasDate = allF.includes('TODAY') || allF.includes('DATE') || allF.includes('DATEDIF') || allF.includes('YEAR') || allF.includes('MONTH');

    let score = 0;
    if (hasText) score += 10;
    if (hasDate) score += 10;

    return { score: Math.min(20, score), detail: `Fungsi Teks: ${hasText?'✓':'✗'}, Fungsi Tanggal: ${hasDate?'✓':'✗'}` };
  });
}

async function checkA5() {
  // Sparklines not easily verifiable, give credit based on confirmation
  return {
    score: state.confirmed[4] ? 18 : 0,
    detail: state.confirmed[4] ? 'Dikonfirmasi peserta ✓' : 'Tidak dikonfirmasi'
  };
}

async function checkA6() {
  return await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.load("protection/protected");
    await context.sync();

    const isProtected = sheet.protection.protected;
    return {
      score: isProtected ? 15 : (state.confirmed[5] ? 5 : 0),
      detail: isProtected ? 'Sheet protection aktif ✓' : 'Sheet tidak terproteksi'
    };
  });
}

/* ═══════════════════════════════════════════════
   WORD CHECK FUNCTIONS
═══════════════════════════════════════════════ */

async function checkW1() {
  return await Word.run(async (context) => {
    const body = context.document.body;
    const search = body.search("Laporan Penjualan Tahunan", { matchCase: false });
    search.load("font/bold,font/size,alignment");
    await context.sync();

    let score = 0;
    let details = [];

    if (search.items.length > 0) {
      const item = search.items[0];
      score += 5; details.push('Teks ditemukan ✓');
      if (item.font.bold) { score += 5; details.push('Bold ✓'); }
      if (item.font.size >= 14) { score += 5; details.push('Ukuran font ✓'); }
      // Alignment check is complex in Word.js, but we'll try
    }

    return { score: Math.min(20, score + 5), detail: details.join(', ') || 'Teks tidak ditemukan' };
  });
}

async function checkW2() {
  return await Word.run(async (context) => {
    const lists = context.document.body.lists;
    lists.load("items");
    await context.sync();

    const hasList = lists.items.length > 0;
    return {
      score: hasList ? 20 : 0,
      detail: hasList ? 'Bullet list ditemukan ✓' : 'Bullet list belum ditemukan'
    };
  });
}

async function checkW3() {
  return await Word.run(async (context) => {
    const body = context.document.body;
    const search = body.search("rahasia", { matchCase: false });
    search.load("font/color,font/underline");
    await context.sync();

    let score = 0;
    if (search.items.length > 0) {
      const item = search.items[0];
      if (item.font.underline !== 'None') score += 15;
      if (item.font.color && item.font.color !== '#000000') score += 15;
    }

    return { score: Math.min(30, score), detail: score > 0 ? 'Format teks ditemukan ✓' : 'Kata rahasia belum diformat' };
  });
}

async function checkW4() {
  return await Word.run(async (context) => {
    const tables = context.document.body.tables;
    tables.load("items");
    await context.sync();

    const hasTable = tables.items.length > 0;
    return {
      score: hasTable ? 25 : 0,
      detail: hasTable ? 'Tabel ditemukan ✓' : 'Tabel belum dibuat'
    };
  });
}

async function checkW5() {
  return await Word.run(async (context) => {
    const sections = context.document.sections;
    sections.load("items/headers");
    await context.sync();

    const hasHeader = sections.items[0].headers.getFirst().type !== 'None';
    return {
      score: hasHeader ? 25 : 0,
      detail: hasHeader ? 'Header ditemukan ✓' : 'Header belum dibuat'
    };
  });
}

/* ═══════════════════════════════════════════════
   SHOW RESULTS
═══════════════════════════════════════════════ */
function showResults(results) {
  const totalScore = results.reduce((s, r) => s + r.pts, 0);
  const maxScore = state.exam.tasks.reduce((s, t) => s + t.points, 0);
  const pct = Math.round(totalScore / maxScore * 100);
  const passing = pct >= state.exam.passingScore;

  // Score circle animation
  const arc = document.getElementById('score-arc');
  const circumference = 314;
  const color = pct >= 85 ? '#4ade80' : pct >= 70 ? '#22d3ee' : pct >= 50 ? '#f59e0b' : '#f87171';
  arc.style.stroke = color;
  setTimeout(() => {
    arc.style.strokeDashoffset = circumference - (circumference * pct / 100);
  }, 100);

  // Animate score number
  const numEl = document.getElementById('final-score-num');
  numEl.style.color = color;
  let current = 0;
  const interval = setInterval(() => {
    current = Math.min(current + Math.ceil(pct / 20), pct);
    numEl.textContent = current;
    if (current >= pct) clearInterval(interval);
  }, 50);

  // Grade & message
  let grade, message;
  if (pct >= 90) { grade = '🏆 A'; message = 'Luar biasa! Kemampuan Excel Anda sangat mahir.'; }
  else if (pct >= 80) { grade = '⭐ B'; message = 'Bagus sekali! Anda menguasai sebagian besar materi.'; }
  else if (pct >= 70) { grade = '✅ C'; message = 'Lulus! Terus berlatih untuk meningkatkan kemampuan.'; }
  else if (pct >= 50) { grade = '📚 D'; message = 'Belum lulus. Pelajari kembali materi dan coba lagi.'; }
  else { grade = '❌ E'; message = 'Perlu banyak latihan. Jangan menyerah!'; }

  document.getElementById('result-grade').textContent = grade + (passing ? ' — LULUS' : ' — TIDAK LULUS');
  document.getElementById('result-grade').style.color = passing ? '#4ade80' : '#f87171';
  document.getElementById('result-message').textContent = message;

  // Breakdown
  const list = document.getElementById('breakdown-list');
  list.innerHTML = results.map((r, i) => `
    <div class="breakdown-item">
      <div class="breakdown-status ${r.status}">${r.status === 'pass' ? '✓' : r.status === 'partial' ? '~' : '✗'}</div>
      <div class="breakdown-info">
        <div class="breakdown-name">${i+1}. ${r.task.title}</div>
        <div class="breakdown-detail">${r.detail}</div>
      </div>
      <div class="breakdown-pts ${r.status}">${r.pts}/${r.task.points}</div>
    </div>
  `).join('');

  showScreen('result');
  document.getElementById('timer-bar').style.display = 'none';
}

/* ═══════════════════════════════════════════════
   EXPORT RESULT
═══════════════════════════════════════════════ */
async function exportResult() {
  try {
    await Excel.run(async (context) => {
      let resultSheet;
      try {
        resultSheet = context.workbook.worksheets.getItem("Hasil Ujian");
        resultSheet.delete();
        await context.sync();
      } catch(e) {}

      resultSheet = context.workbook.worksheets.add("Hasil Ujian");
      resultSheet.activate();

      // Title
      resultSheet.getRange("A1").values = [["LAPORAN HASIL UJIAN EXCEL"]];
      resultSheet.getRange("A1").format.font.bold = true;
      resultSheet.getRange("A1").format.font.size = 14;
      resultSheet.getRange("A1").format.font.color = "#4ade80";

      resultSheet.getRange("A2").values = [["ExcelQuiz Pro — " + state.exam.name]];
      resultSheet.getRange("A3").values = [["Tanggal: " + new Date().toLocaleDateString('id-ID', {day:'2-digit',month:'long',year:'numeric'})]];
      resultSheet.getRange("A4").values = [["Waktu Selesai: " + new Date().toLocaleTimeString('id-ID')]];

      const totalScore = state.scores.reduce((s, v) => s + v, 0);
      const maxScore = state.exam.tasks.reduce((s, t) => s + t.points, 0);
      const pct = Math.round(totalScore / maxScore * 100);
      const passing = pct >= state.exam.passingScore;

      resultSheet.getRange("A6").values = [["SKOR TOTAL"]];
      resultSheet.getRange("B6").values = [[`${totalScore} / ${maxScore} (${pct}%)`]];
      resultSheet.getRange("A7").values = [["STATUS"]];
      resultSheet.getRange("B7").values = [[passing ? "LULUS ✓" : "TIDAK LULUS ✗"]];
      resultSheet.getRange("B7").format.font.color = passing ? "#4ade80" : "#f87171";

      // Table header
      resultSheet.getRange("A9:E9").values = [["No", "Soal", "Max Poin", "Poin Diraih", "Status"]];
      resultSheet.getRange("A9:E9").format.font.bold = true;
      resultSheet.getRange("A9:E9").format.fill.color = "#1c2030";

      state.exam.tasks.forEach((task, i) => {
        const row = 10 + i;
        const sc = state.scores[i];
        const status = sc >= task.points ? "Lulus" : sc > 0 ? "Parsial" : "Gagal";
        resultSheet.getRange(`A${row}:E${row}`).values = [[i+1, task.title, task.points, sc, status]];
        if (sc >= task.points) resultSheet.getRange(`E${row}`).format.font.color = "#4ade80";
        else if (sc > 0) resultSheet.getRange(`E${row}`).format.font.color = "#f59e0b";
        else resultSheet.getRange(`E${row}`).format.font.color = "#f87171";
      });

      // Auto fit columns
      resultSheet.getRange("A:E").format.autofitColumns();
      await context.sync();
    });

    showToast('Laporan berhasil dibuat di sheet "Hasil Ujian" ✓', 'success');
  } catch(e) {
    showToast('Gagal membuat laporan: ' + e.message, 'error');
  }
}

/* ═══════════════════════════════════════════════
   RESTART
═══════════════════════════════════════════════ */
function restartExam() {
  clearInterval(state.timerInterval);
  state = { examKey:'basic', exam:null, currentIdx:0, confirmed:[], scores:[], timerInterval:null, timeLeft:0, started:false, finished:false };
  document.getElementById('timer-bar').style.display = 'none';
  document.getElementById('progress-outer').style.display = 'none';
  document.getElementById('header-sub').textContent = 'Pilih ujian untuk memulai';
  document.getElementById('exam-select').value = 'basic';
  updateExamInfo();
  showScreen('welcome');
}

/* ═══════════════════════════════════════════════
   HELPERS
═══════════════════════════════════════════════ */
function showScreen(name) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('screen-' + name).classList.add('active');
}

function showToast(msg, type = '') {
  const t = document.getElementById('toast');
  t.textContent = msg;
  t.className = 'toast ' + type;
  t.classList.add('show');
  setTimeout(() => t.classList.remove('show'), 3000);
}
