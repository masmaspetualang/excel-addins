/* ═══════════════════════════════════════════════════════
   Excel Practice Quiz — Taskpane Logic
   ═══════════════════════════════════════════════════════ */

/* ── Quiz Definitions ── */
const QUIZZES = [
  {
    id: "basic-formula",
    name: "Rumus Dasar Excel",
    icon: "🧮",
    difficulty: "Mudah",
    diffTag: "tag-green",
    desc: "SUM, AVERAGE, COUNT dan rumus dasar lainnya",
    time: 600,
    setupSheet: setupBasicFormulaSheet,
    tasks: [
      {
        id: "sum",
        instruction: "Di sel <strong>B10</strong>, buat rumus <code>=SUM()</code> untuk menjumlahkan semua nilai di kolom B (B2 hingga B9).",
        hint: "Klik sel B10, ketik =SUM(B2:B9) lalu tekan Enter",
        check: checkSumTask,
        maxScore: 20,
      },
      {
        id: "average",
        instruction: "Di sel <strong>C10</strong>, buat rumus <code>=AVERAGE()</code> untuk menghitung rata-rata nilai di kolom C (C2 hingga C9).",
        hint: "Klik sel C10, ketik =AVERAGE(C2:C9) lalu tekan Enter",
        check: checkAverageTask,
        maxScore: 20,
      },
      {
        id: "max-min",
        instruction: "Di sel <strong>D10</strong> buat rumus <code>=MAX(D2:D9)</code>, dan di sel <strong>D11</strong> buat rumus <code>=MIN(D2:D9)</code>.",
        hint: "Gunakan fungsi MAX untuk nilai tertinggi dan MIN untuk terendah",
        check: checkMaxMinTask,
        maxScore: 20,
      },
      {
        id: "count",
        instruction: "Di sel <strong>E10</strong>, gunakan rumus <code>=COUNT()</code> untuk menghitung berapa banyak angka ada di kolom E (E2:E9).",
        hint: "COUNT menghitung sel yang berisi angka. Ketik =COUNT(E2:E9)",
        check: checkCountTask,
        maxScore: 20,
      },
      {
        id: "header",
        instruction: "Tambahkan teks <strong>\"TOTAL\"</strong> di sel <strong>A10</strong> dan buat teks tersebut <strong>Bold</strong> (tebal).",
        hint: "Klik A10, ketik TOTAL, lalu tekan Ctrl+B untuk bold",
        check: checkHeaderTask,
        maxScore: 20,
      },
    ]
  },
  {
    id: "data-format",
    name: "Format & Tampilan Data",
    icon: "🎨",
    difficulty: "Menengah",
    diffTag: "tag-yellow",
    desc: "Format sel, warna, border, dan penataan data",
    time: 480,
    setupSheet: setupFormatSheet,
    tasks: [
      {
        id: "bold-header",
        instruction: "Pilih baris pertama (<strong>A1:E1</strong>) dan buat semua teks header menjadi <strong>Bold</strong> (tebal).",
        hint: "Klik A1, tahan Shift klik E1, lalu tekan Ctrl+B",
        check: checkBoldHeader,
        maxScore: 25,
      },
      {
        id: "bg-color",
        instruction: "Beri warna background <strong>kuning (Yellow)</strong> pada sel <strong>A1:E1</strong> (baris header).",
        hint: "Pilih A1:E1 → Home → Fill Color → pilih Yellow",
        check: checkYellowBg,
        maxScore: 25,
      },
      {
        id: "center-align",
        instruction: "Buat semua isi sel di range <strong>A1:E1</strong> menjadi <strong>rata tengah (Center)</strong>.",
        hint: "Pilih A1:E1 → tekan Ctrl+E atau klik tombol Center di Home",
        check: checkCenterAlign,
        maxScore: 25,
      },
      {
        id: "number-format",
        instruction: "Format sel <strong>B2:B9</strong> agar angka ditampilkan dalam format <strong>mata uang (Currency)</strong> atau format angka dengan 2 desimal.",
        hint: "Pilih B2:B9 → Ctrl+1 → Number → Currency atau klik $ di toolbar",
        check: checkNumberFormat,
        maxScore: 25,
      },
    ]
  },
  {
    id: "if-logic",
    name: "Logika IF & Kondisi",
    icon: "🔀",
    difficulty: "Menengah",
    diffTag: "tag-yellow",
    desc: "Fungsi IF, AND, OR untuk logika bersyarat",
    time: 540,
    setupSheet: setupIfSheet,
    tasks: [
      {
        id: "if-basic",
        instruction: "Di sel <strong>C2</strong>, buat rumus IF: jika nilai di B2 ≥ 75, tampilkan <strong>\"LULUS\"</strong>, jika tidak tampilkan <strong>\"GAGAL\"</strong>.",
        hint: "=IF(B2>=75,\"LULUS\",\"GAGAL\")",
        check: checkIfBasic,
        maxScore: 25,
      },
      {
        id: "if-fill",
        instruction: "Salin rumus dari <strong>C2</strong> ke sel <strong>C3:C9</strong> (isi ke bawah untuk semua siswa).",
        hint: "Klik C2, lalu seret fill handle (kotak kecil di pojok kanan bawah) ke C9",
        check: checkIfFill,
        maxScore: 25,
      },
      {
        id: "if-grade",
        instruction: "Di sel <strong>D2</strong>, buat rumus IF bertingkat: ≥90 → <strong>\"A\"</strong>, ≥75 → <strong>\"B\"</strong>, ≥60 → <strong>\"C\"</strong>, sisanya → <strong>\"D\"</strong>.",
        hint: "=IF(B2>=90,\"A\",IF(B2>=75,\"B\",IF(B2>=60,\"C\",\"D\")))",
        check: checkIfGrade,
        maxScore: 25,
      },
      {
        id: "countif",
        instruction: "Di sel <strong>F2</strong>, gunakan <code>=COUNTIF()</code> untuk menghitung berapa siswa yang LULUS (dari kolom C).",
        hint: "=COUNTIF(C2:C9,\"LULUS\")",
        check: checkCountIf,
        maxScore: 25,
      },
    ]
  },
  {
    id: "vlookup",
    name: "VLOOKUP & Referensi",
    icon: "🔍",
    difficulty: "Sulit",
    diffTag: "tag-red",
    desc: "VLOOKUP, referensi absolut, dan pencarian data",
    time: 600,
    setupSheet: setupVlookupSheet,
    tasks: [
      {
        id: "abs-ref",
        instruction: "Di sel <strong>C2</strong>, buat rumus dengan <strong>referensi absolut</strong>: kalikan nilai B2 dengan tarif pajak di sel <strong>$E$2</strong>.",
        hint: "=$E$2*B2 — tanda $ mengunci referensi sel agar tidak bergeser saat disalin",
        check: checkAbsRef,
        maxScore: 25,
      },
      {
        id: "vlookup-basic",
        instruction: "Di sel <strong>G2</strong>, gunakan <code>=VLOOKUP()</code> untuk mencari nama produk berdasarkan kode di F2, menggunakan tabel referensi di <strong>A2:B8</strong>.",
        hint: "=VLOOKUP(F2,A2:B8,2,FALSE) — kolom 2 = kolom nama produk",
        check: checkVlookup,
        maxScore: 25,
      },
      {
        id: "vlookup-fill",
        instruction: "Salin rumus VLOOKUP dari <strong>G2</strong> ke <strong>G3:G6</strong> (gunakan referensi absolut pada tabel: <strong>$A$2:$B$8</strong>).",
        hint: "Ubah dulu G2 jadi =VLOOKUP(F2,$A$2:$B$8,2,FALSE) lalu salin ke G3:G6",
        check: checkVlookupFill,
        maxScore: 25,
      },
      {
        id: "sumif",
        instruction: "Di sel <strong>I2</strong>, gunakan <code>=SUMIF()</code> untuk menjumlahkan total harga hanya untuk produk kategori <strong>\"Elektronik\"</strong> (kolom A=kategori, kolom C=harga).",
        hint: "=SUMIF(A2:A8,\"Elektronik\",C2:C8)",
        check: checkSumIf,
        maxScore: 25,
      },
    ]
  },
];

/* ── State ── */
let state = {
  currentQuiz: null,
  currentTaskIdx: 0,
  scores: [],
  timerInterval: null,
  elapsedSeconds: 0,
  taskConfirmed: false,
};

/* ═══════════════════════════════════
   OFFICE.JS INIT
═══════════════════════════════════ */
Office.onReady(() => {
  setTimeout(() => {
    const loading = document.getElementById("loading-screen");
    loading.style.opacity = "0";
    setTimeout(() => loading.style.display = "none", 500);
  }, 1800);
  renderQuizList();
});

/* ═══════════════════════════════════
   NAVIGATION
═══════════════════════════════════ */
function showView(id) {
  document.querySelectorAll(".view").forEach(v => v.classList.remove("active"));
  document.getElementById(id).classList.add("active");
}

function goHome() {
  stopTimer();
  showView("view-home");
  document.getElementById("badge-mode").textContent = "READY";
}

/* ═══════════════════════════════════
   RENDER HOME
═══════════════════════════════════ */
function renderQuizList() {
  const list = document.getElementById("quiz-list");
  list.innerHTML = QUIZZES.map(q => `
    <div class="quiz-card" onclick="startQuiz('${q.id}')">
      <div class="quiz-card-icon" style="background:${getIconBg(q.difficulty)}">${q.icon}</div>
      <div class="quiz-card-info">
        <div class="quiz-card-name">${q.name}</div>
        <div class="quiz-card-desc">${q.desc}</div>
        <div class="quiz-card-meta">
          <span class="tag ${q.diffTag}">${q.difficulty}</span>
          <span class="tag tag-blue">${q.tasks.length} Soal</span>
          <span class="tag tag-blue">⏱ ${Math.floor(q.time/60)} Menit</span>
        </div>
      </div>
      <div class="quiz-card-arrow">›</div>
    </div>
  `).join("");
}

function getIconBg(diff) {
  if (diff === "Mudah") return "rgba(0,255,136,0.15)";
  if (diff === "Menengah") return "rgba(255,211,42,0.15)";
  return "rgba(255,71,87,0.15)";
}

/* ═══════════════════════════════════
   START QUIZ
═══════════════════════════════════ */
async function startQuiz(quizId) {
  const quiz = QUIZZES.find(q => q.id === quizId);
  if (!quiz) return;

  state.currentQuiz = quiz;
  state.currentTaskIdx = 0;
  state.scores = [];
  state.elapsedSeconds = 0;
  state.taskConfirmed = false;

  document.getElementById("quiz-name-display").textContent = quiz.name;
  document.getElementById("badge-mode").textContent = "UJIAN";
  showView("view-quiz");

  // Setup spreadsheet
  document.getElementById("setup-notice").style.display = "block";
  try {
    await quiz.setupSheet();
  } catch(e) {
    console.warn("Sheet setup error:", e);
  }
  setTimeout(() => {
    document.getElementById("setup-notice").style.display = "none";
  }, 1500);

  startTimer();
  renderTask();
}

/* ═══════════════════════════════════
   TIMER
═══════════════════════════════════ */
function startTimer() {
  stopTimer();
  state.elapsedSeconds = 0;
  state.timerInterval = setInterval(() => {
    state.elapsedSeconds++;
    const m = String(Math.floor(state.elapsedSeconds/60)).padStart(2,"0");
    const s = String(state.elapsedSeconds%60).padStart(2,"0");
    document.getElementById("timer-display").textContent = `${m}:${s}`;
  }, 1000);
}

function stopTimer() {
  if (state.timerInterval) {
    clearInterval(state.timerInterval);
    state.timerInterval = null;
  }
}

/* ═══════════════════════════════════
   RENDER TASK
═══════════════════════════════════ */
function renderTask() {
  const quiz = state.currentQuiz;
  const task = quiz.tasks[state.currentTaskIdx];
  const total = quiz.tasks.length;
  const idx = state.currentTaskIdx;

  // Progress
  const pct = (idx / total) * 100;
  document.getElementById("progress-fill").style.width = pct + "%";

  // Step
  document.getElementById("quiz-step-display").textContent = `SOAL ${idx+1} / ${total}`;

  // Task number
  document.getElementById("task-number").textContent = `TUGAS ${String(idx+1).padStart(2,"0")}`;

  // Instruction
  document.getElementById("task-instruction").innerHTML = task.instruction;

  // Hint
  document.getElementById("hint-text").textContent = task.hint;

  // Status
  setStatus("pending", "Kerjakan tugas di spreadsheet, lalu klik Konfirmasi.");

  // Buttons
  document.getElementById("btn-confirm").classList.remove("hidden");
  document.getElementById("btn-confirm").disabled = false;
  document.getElementById("btn-next").classList.add("hidden");

  // Clear result
  const result = document.getElementById("check-result");
  result.className = "check-result";

  state.taskConfirmed = false;
}

function setStatus(type, msg) {
  const dot = document.getElementById("status-dot");
  dot.className = "status-dot " + type;
  document.getElementById("status-text").textContent = msg;
}

/* ═══════════════════════════════════
   CONFIRM TASK
═══════════════════════════════════ */
async function confirmTask() {
  const btn = document.getElementById("btn-confirm");
  btn.disabled = true;
  setStatus("checking", "Memeriksa pekerjaan kamu...");

  const quiz = state.currentQuiz;
  const task = quiz.tasks[state.currentTaskIdx];

  try {
    const result = await task.check();
    showCheckResult(result, task.maxScore);
  } catch(e) {
    showCheckResult({ score: 0, correct: false, title: "Error", detail: "Tidak bisa membaca sel. Pastikan kamu sudah mengisi sel yang diminta." }, task.maxScore);
  }
}

function showCheckResult(result, maxScore) {
  const score = Math.round((result.score / 100) * maxScore);
  state.scores.push({ score, maxScore, title: result.title, correct: result.correct, partial: result.partial });

  const el = document.getElementById("check-result");
  const titleEl = document.getElementById("result-title");
  const detailEl = document.getElementById("result-detail");

  if (result.correct) {
    el.className = "check-result show correct";
    titleEl.innerHTML = "✅ " + result.title;
    setStatus("done", "Jawaban benar! Lanjutkan ke soal berikutnya.");
  } else if (result.partial) {
    el.className = "check-result show partial";
    titleEl.innerHTML = "⚠️ " + result.title;
    setStatus("done", "Sebagian benar. Kamu bisa lanjut.");
  } else {
    el.className = "check-result show incorrect";
    titleEl.innerHTML = "❌ " + result.title;
    setStatus("error", "Jawaban kurang tepat. Coba periksa kembali.");
  }
  detailEl.textContent = result.detail;

  document.getElementById("btn-confirm").classList.add("hidden");

  const isLast = state.currentTaskIdx >= state.currentQuiz.tasks.length - 1;
  const btnNext = document.getElementById("btn-next");
  btnNext.classList.remove("hidden");
  btnNext.textContent = isLast ? "Lihat Hasil Akhir 🏆" : "Lanjut ke Soal Berikutnya →";
}

/* ═══════════════════════════════════
   NEXT TASK
═══════════════════════════════════ */
function nextTask() {
  state.currentTaskIdx++;
  if (state.currentTaskIdx >= state.currentQuiz.tasks.length) {
    stopTimer();
    showResult();
  } else {
    renderTask();
  }
}

/* ═══════════════════════════════════
   SHOW RESULT
═══════════════════════════════════ */
function showResult() {
  const total = state.scores.reduce((a,b) => a + b.maxScore, 0);
  const earned = state.scores.reduce((a,b) => a + b.score, 0);
  const pct = Math.round((earned / total) * 100);

  // Animate score circle
  document.getElementById("final-score").textContent = pct;
  setTimeout(() => {
    const circumference = 377;
    const offset = circumference - (pct / 100) * circumference;
    document.getElementById("score-arc").style.transition = "stroke-dashoffset 1.2s ease";
    document.getElementById("score-arc").style.strokeDashoffset = offset;
  }, 200);

  // Grade
  let grade, gradeText;
  if (pct >= 90) { grade = "A"; gradeText = "Luar Biasa! Kamu Sangat Menguasai Materi 🎉"; }
  else if (pct >= 75) { grade = "B"; gradeText = "Bagus! Hasil yang Memuaskan 👍"; }
  else if (pct >= 60) { grade = "C"; gradeText = "Cukup. Masih Ada Ruang untuk Berkembang 📚"; }
  else { grade = "D"; gradeText = "Perlu Belajar Lebih Banyak. Jangan Menyerah! 💪"; }

  document.getElementById("grade-letter").textContent = grade;
  document.getElementById("grade-text").textContent = gradeText;
  const banner = document.getElementById("grade-banner");
  banner.className = "grade-banner " + grade;

  // Breakdown
  const breakdown = document.getElementById("breakdown-list");
  breakdown.innerHTML = state.scores.map((s, i) => {
    const cls = s.score === s.maxScore ? "ok" : s.score > 0 ? "partial" : "fail";
    const icon = s.score === s.maxScore ? "✅" : s.score > 0 ? "⚠️" : "❌";
    return `
      <div class="breakdown-item">
        <div class="bd-icon">${icon}</div>
        <div class="bd-name">Soal ${i+1}: ${state.currentQuiz.tasks[i].id}</div>
        <div class="bd-score ${cls}">${s.score}/${s.maxScore}</div>
      </div>
    `;
  }).join("") + `
    <div class="breakdown-item" style="background:rgba(0,229,255,0.04)">
      <div class="bd-icon">🏆</div>
      <div class="bd-name" style="color:#e8eaf6;font-weight:600">Total Skor</div>
      <div class="bd-score" style="color:#00e5ff;font-weight:700">${earned}/${total}</div>
    </div>
  `;

  document.getElementById("badge-mode").textContent = "SELESAI";
  showView("view-result");
}


/* ═══════════════════════════════════════════════════════════
   SHEET SETUP FUNCTIONS
═══════════════════════════════════════════════════════════ */

async function setupBasicFormulaSheet() {
  await Excel.run(async (ctx) => {
    let sheet;
    try {
      sheet = ctx.workbook.worksheets.getItem("QuizBasic");
      sheet.delete();
      await ctx.sync();
    } catch(e) {}

    sheet = ctx.workbook.worksheets.add("QuizBasic");
    sheet.activate();

    const headers = ["Nama", "Nilai Ujian", "Nilai Tugas", "Nilai Akhir", "Hadir"];
    const data = [
      ["Andi", 80, 75, 77.5, 1],
      ["Budi", 65, 70, 67.5, 1],
      ["Citra", 90, 88, 89, 1],
      ["Doni", 72, 68, 70, 1],
      ["Eka", 85, 92, 88.5, 1],
      ["Fira", 55, 60, 57.5, 0],
      ["Gani", 78, 75, 76.5, 1],
      ["Hana", 91, 95, 93, 1],
    ];

    // Headers
    headers.forEach((h, i) => {
      sheet.getCell(0, i).values = [[h]];
    });

    // Data rows
    data.forEach((row, r) => {
      row.forEach((val, c) => {
        sheet.getCell(r+1, c).values = [[val]];
      });
    });

    // Label row 10 col A (row 9 in 0-index)
    sheet.getCell(9, 0).values = [[""]];

    sheet.getUsedRange().format.autofitColumns();
    await ctx.sync();
  });
}

async function setupFormatSheet() {
  await Excel.run(async (ctx) => {
    let sheet;
    try {
      sheet = ctx.workbook.worksheets.getItem("QuizFormat");
      sheet.delete();
      await ctx.sync();
    } catch(e) {}

    sheet = ctx.workbook.worksheets.add("QuizFormat");
    sheet.activate();

    const headers = ["Produk", "Harga", "Stok", "Terjual", "Pendapatan"];
    const data = [
      ["Laptop", 8500000, 20, 5, 42500000],
      ["Mouse", 150000, 100, 45, 6750000],
      ["Keyboard", 350000, 60, 30, 10500000],
      ["Monitor", 3200000, 15, 8, 25600000],
      ["Headset", 450000, 40, 25, 11250000],
      ["Webcam", 700000, 25, 12, 8400000],
      ["USB Hub", 120000, 80, 55, 6600000],
      ["Charger", 95000, 120, 90, 8550000],
    ];

    headers.forEach((h, i) => sheet.getCell(0, i).values = [[h]]);
    data.forEach((row, r) => row.forEach((val, c) => sheet.getCell(r+1, c).values = [[val]]));

    sheet.getUsedRange().format.autofitColumns();
    await ctx.sync();
  });
}

async function setupIfSheet() {
  await Excel.run(async (ctx) => {
    let sheet;
    try {
      sheet = ctx.workbook.worksheets.getItem("QuizIF");
      sheet.delete();
      await ctx.sync();
    } catch(e) {}

    sheet = ctx.workbook.worksheets.add("QuizIF");
    sheet.activate();

    const headers = ["Nama Siswa", "Nilai", "Status", "Grade", "", "Keterangan", "Jumlah"];
    headers.forEach((h, i) => sheet.getCell(0, i).values = [[h]]);

    const students = [
      ["Ahmad", 85],["Bella", 72],["Candra", 91],["Diana", 58],
      ["Evan", 76],["Fani", 63],["Gita", 88],["Hendra", 45],
    ];
    students.forEach(([name, val], r) => {
      sheet.getCell(r+1, 0).values = [[name]];
      sheet.getCell(r+1, 1).values = [[val]];
    });

    sheet.getCell(1, 5).values = [["Jumlah Lulus:"]];
    sheet.getUsedRange().format.autofitColumns();
    await ctx.sync();
  });
}

async function setupVlookupSheet() {
  await Excel.run(async (ctx) => {
    let sheet;
    try {
      sheet = ctx.workbook.worksheets.getItem("QuizVlookup");
      sheet.delete();
      await ctx.sync();
    } catch(e) {}

    sheet = ctx.workbook.worksheets.add("QuizVlookup");
    sheet.activate();

    // Product table A1:C8
    const prodHeaders = ["Kode", "Nama Produk", "Harga"];
    prodHeaders.forEach((h,i) => sheet.getCell(0,i).values = [[h]]);

    const products = [
      ["P001","Laptop Asus",8500000],
      ["P002","Mouse Logitech",150000],
      ["P003","Keyboard Mechanical",450000],
      ["P004","Monitor LG",3200000],
      ["P005","Headset Sony",650000],
      ["P006","Webcam HD",700000],
      ["P007","USB Hub",120000],
    ];
    products.forEach(([kode,nama,harga],r) => {
      sheet.getCell(r+1,0).values=[[kode]];
      sheet.getCell(r+1,1).values=[[nama]];
      sheet.getCell(r+1,2).values=[[harga]];
    });

    // Tax rate at E2
    sheet.getCell(0,4).values=[["Tarif Pajak"]];
    sheet.getCell(1,4).values=[[0.11]];
    sheet.getCell(1,4).numberFormat=[["0%"]];

    // Tax column header
    sheet.getCell(0,2).values=[["Harga"]];
    sheet.getCell(0,3).values=[["Pajak (C*E2)"]];

    // Lookup table headers F-G
    sheet.getCell(0,5).values=[["Cari Kode"]];
    sheet.getCell(0,6).values=[["Nama Produk"]];
    const lookupCodes = ["P003","P001","P005","P007","P002","P006"];
    lookupCodes.forEach((c,r) => sheet.getCell(r+1,5).values=[[c]]);

    // SUMIF headers
    sheet.getCell(0,8).values=[["Kategori"]];
    sheet.getCell(1,8).values=[["Elektronik"]];

    sheet.getUsedRange().format.autofitColumns();
    await ctx.sync();
  });
}


/* ═══════════════════════════════════════════════════════════
   TASK CHECK FUNCTIONS
═══════════════════════════════════════════════════════════ */

// ── Basic Formula ──

async function checkSumTask() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizBasic");
      const cell = sheet.getRange("B10");
      cell.load("formulas,values");
      await ctx.sync();

      const formula = String(cell.formulas[0][0]).toUpperCase();
      const value = cell.values[0][0];

      if (formula.includes("SUM") && formula.includes("B2") && (formula.includes("B9") || formula.includes("B:B"))) {
        const expected = 80+65+90+72+85+55+78+91;
        if (Math.abs(value - expected) < 0.01) {
          return { score: 100, correct: true, title: "Benar Sempurna!", detail: `Rumus SUM kamu menghasilkan nilai ${value} dengan tepat.` };
        }
        return { score: 60, partial: true, title: "Rumus Benar, Hasil Berbeda", detail: `Rumus SUM ditemukan tapi hasilnya ${value}, seharusnya ${expected}.` };
      }
      if (formula.includes("SUM")) {
        return { score: 40, partial: true, title: "Fungsi SUM Ditemukan", detail: "Kamu pakai SUM, tapi range-nya salah. Pastikan B2:B9." };
      }
      return { score: 0, correct: false, title: "Belum Menggunakan SUM", detail: "Sel B10 belum berisi rumus =SUM(B2:B9)." };
    } catch(e) {
      return { score: 0, correct: false, title: "Sheet Tidak Ditemukan", detail: "Pastikan sheet QuizBasic ada dan sel B10 sudah diisi." };
    }
  });
}

async function checkAverageTask() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizBasic");
      const cell = sheet.getRange("C10");
      cell.load("formulas,values");
      await ctx.sync();

      const formula = String(cell.formulas[0][0]).toUpperCase();
      const value = cell.values[0][0];

      if (formula.includes("AVERAGE") && formula.includes("C2")) {
        const expected = (75+70+88+68+92+60+75+95)/8;
        if (Math.abs(value - expected) < 0.1) {
          return { score: 100, correct: true, title: "Benar!", detail: `AVERAGE menghasilkan ${value.toFixed(2)}.` };
        }
        return { score: 50, partial: true, title: "Fungsi AVERAGE Ada", detail: `Hasilnya ${value}, periksa range-nya (C2:C9).` };
      }
      return { score: 0, correct: false, title: "Belum Ada AVERAGE", detail: "Isi sel C10 dengan =AVERAGE(C2:C9)." };
    } catch(e) {
      return { score: 0, correct: false, title: "Error", detail: e.message };
    }
  });
}

async function checkMaxMinTask() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizBasic");
      const d10 = sheet.getRange("D10"); d10.load("formulas,values");
      const d11 = sheet.getRange("D11"); d11.load("formulas,values");
      await ctx.sync();

      const f10 = String(d10.formulas[0][0]).toUpperCase();
      const f11 = String(d11.formulas[0][0]).toUpperCase();
      const hasMax = f10.includes("MAX");
      const hasMin = f11.includes("MIN");

      if (hasMax && hasMin) return { score: 100, correct: true, title: "MAX & MIN Benar!", detail: `D10=${d10.values[0][0]} (MAX), D11=${d11.values[0][0]} (MIN).` };
      if (hasMax || hasMin) return { score: 50, partial: true, title: "Setengah Benar", detail: `${hasMax?"MAX ada":"MAX belum ada"}, ${hasMin?"MIN ada":"MIN belum ada"}.` };
      return { score: 0, correct: false, title: "Belum Ada MAX/MIN", detail: "Isi D10 dengan =MAX(D2:D9) dan D11 dengan =MIN(D2:D9)." };
    } catch(e) {
      return { score: 0, correct: false, title: "Error", detail: e.message };
    }
  });
}

async function checkCountTask() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizBasic");
      const cell = sheet.getRange("E10"); cell.load("formulas,values");
      await ctx.sync();
      const formula = String(cell.formulas[0][0]).toUpperCase();
      if (formula.includes("COUNT") && formula.includes("E2")) {
        return { score: 100, correct: true, title: "COUNT Benar!", detail: `Hasil COUNT = ${cell.values[0][0]}.` };
      }
      return { score: 0, correct: false, title: "Belum Ada COUNT", detail: "Isi E10 dengan =COUNT(E2:E9)." };
    } catch(e) {
      return { score: 0, correct: false, title: "Error", detail: e.message };
    }
  });
}

async function checkHeaderTask() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizBasic");
      const cell = sheet.getRange("A10");
      cell.load("values");
      cell.format.font.load("bold");
      await ctx.sync();

      const val = String(cell.values[0][0]).toUpperCase().trim();
      const isBold = cell.format.font.bold;

      if (val === "TOTAL" && isBold) return { score: 100, correct: true, title: "Sempurna!", detail: 'Teks "TOTAL" sudah ada dan ditebalkan (bold).' };
      if (val === "TOTAL") return { score: 60, partial: true, title: "Teks Ada, Belum Bold", detail: 'Teks "TOTAL" sudah ada tapi belum ditebalkan. Pilih A10 lalu Ctrl+B.' };
      if (isBold) return { score: 30, partial: true, title: "Bold Ada, Teks Salah", detail: `Sudah bold tapi isinya "${cell.values[0][0]}", seharusnya "TOTAL".` };
      return { score: 0, correct: false, title: "Belum Dikerjakan", detail: 'Ketik "TOTAL" di sel A10 lalu tekan Ctrl+B.' };
    } catch(e) {
      return { score: 0, correct: false, title: "Error", detail: e.message };
    }
  });
}

// ── Format Tasks ──

async function checkBoldHeader() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizFormat");
      const range = sheet.getRange("A1:E1");
      range.format.font.load("bold");
      await ctx.sync();
      if (range.format.font.bold) return { score: 100, correct: true, title: "Header Bold!", detail: "Semua header baris 1 sudah ditebalkan." };
      return { score: 0, correct: false, title: "Belum Bold", detail: "Pilih A1:E1 lalu tekan Ctrl+B." };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

async function checkYellowBg() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizFormat");
      const range = sheet.getRange("A1");
      range.format.fill.load("color");
      await ctx.sync();
      const color = (range.format.fill.color || "").toLowerCase();
      const isYellow = color.includes("ffff00") || color.includes("ffff") || color === "#ffff00";
      if (isYellow) return { score: 100, correct: true, title: "Warna Kuning Benar!", detail: `Background color: ${color}` };
      if (color && color !== "none" && color !== "") return { score: 40, partial: true, title: "Ada Warna, Bukan Kuning", detail: `Warnanya ${color}, seharusnya kuning (#FFFF00).` };
      return { score: 0, correct: false, title: "Belum Ada Warna", detail: "Pilih A1:E1 → Home → Fill Color → Yellow." };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

async function checkCenterAlign() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizFormat");
      const range = sheet.getRange("A1");
      range.format.load("horizontalAlignment");
      await ctx.sync();
      const align = (range.format.horizontalAlignment || "").toLowerCase();
      if (align === "center") return { score: 100, correct: true, title: "Rata Tengah Benar!", detail: "Teks header sudah rata tengah." };
      return { score: 0, correct: false, title: "Belum Center", detail: `Alignment saat ini: ${align}. Pilih A1:E1 lalu tekan Ctrl+E.` };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

async function checkNumberFormat() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizFormat");
      const range = sheet.getRange("B2");
      range.load("numberFormat");
      await ctx.sync();
      const fmt = String(range.numberFormat[0][0] || "").toLowerCase();
      const isCurrency = fmt.includes("$") || fmt.includes("rp") || fmt.includes("#,##0") || fmt.includes("accounting") || fmt.includes("currency");
      if (isCurrency) return { score: 100, correct: true, title: "Format Mata Uang Benar!", detail: `Format: ${fmt}` };
      if (fmt.includes("0.00") || fmt.includes("#,##")) return { score: 70, partial: true, title: "Format Angka Ada", detail: "Format angka sudah ada tapi bukan currency. Coba format Currency." };
      return { score: 0, correct: false, title: "Format Default", detail: "Pilih B2:B9 → Ctrl+1 → Number → Currency." };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

// ── IF Tasks ──

async function checkIfBasic() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizIF");
      const cell = sheet.getRange("C2"); cell.load("formulas,values");
      await ctx.sync();
      const f = String(cell.formulas[0][0]).toUpperCase();
      const v = String(cell.values[0][0]).toUpperCase();
      if (f.includes("IF") && f.includes("B2") && f.includes("LULUS") && f.includes("GAGAL")) {
        if (v === "LULUS") return { score: 100, correct: true, title: "IF Benar!", detail: `Untuk nilai 85, hasilnya "${cell.values[0][0]}" — tepat!` };
        return { score: 60, partial: true, title: "Rumus IF Ada", detail: `Hasilnya "${cell.values[0][0]}", cek kondisi >=75.` };
      }
      if (f.includes("IF")) return { score: 30, partial: true, title: "Fungsi IF Ada", detail: "Rumus IF ada tapi perlu LULUS dan GAGAL." };
      return { score: 0, correct: false, title: "Belum Ada IF", detail: 'Isi C2 dengan =IF(B2>=75,"LULUS","GAGAL").' };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

async function checkIfFill() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizIF");
      const range = sheet.getRange("C2:C9"); range.load("values,formulas");
      await ctx.sync();
      let filled = 0;
      for (let i = 0; i < 8; i++) {
        const f = String(range.formulas[i][0]).toUpperCase();
        if (f.includes("IF") && f.includes("LULUS")) filled++;
      }
      if (filled === 8) return { score: 100, correct: true, title: "Semua Terisi!", detail: "Rumus IF berhasil disalin ke semua 8 baris." };
      if (filled > 1) return { score: Math.round((filled/8)*100), partial: true, title: `${filled}/8 Baris Terisi`, detail: `Salin hingga C9. Sekarang baru ${filled} baris.` };
      return { score: 0, correct: false, title: "Belum Disalin", detail: "Klik C2, lalu seret fill handle ke C9." };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

async function checkIfGrade() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizIF");
      const cell = sheet.getRange("D2"); cell.load("formulas,values");
      await ctx.sync();
      const f = String(cell.formulas[0][0]).toUpperCase();
      const hasNested = f.includes("IF") && (f.match(/IF/g) || []).length >= 3;
      const hasGrades = f.includes('"A"') || f.includes('"B"') || f.includes("\"A\"");
      if (hasNested && (f.includes("90") || f.includes("75"))) {
        return { score: 100, correct: true, title: "IF Bertingkat Benar!", detail: `Untuk nilai 85, grade = "${cell.values[0][0]}".` };
      }
      if (f.includes("IF") && hasGrades) {
        return { score: 50, partial: true, title: "IF Ada", detail: "Coba lengkapi dengan 3 kondisi: >=90, >=75, >=60." };
      }
      return { score: 0, correct: false, title: "Belum Ada", detail: 'Isi D2 dengan =IF(B2>=90,"A",IF(B2>=75,"B",IF(B2>=60,"C","D"))).' };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

async function checkCountIf() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizIF");
      const cell = sheet.getRange("F2"); cell.load("formulas,values");
      await ctx.sync();
      const f = String(cell.formulas[0][0]).toUpperCase();
      if (f.includes("COUNTIF") && f.includes("LULUS")) {
        return { score: 100, correct: true, title: "COUNTIF Benar!", detail: `Jumlah siswa lulus = ${cell.values[0][0]}.` };
      }
      if (f.includes("COUNTIF")) return { score: 50, partial: true, title: "COUNTIF Ada", detail: 'Pastikan kriteria "LULUS" dan range C2:C9.' };
      return { score: 0, correct: false, title: "Belum COUNTIF", detail: 'Isi F2 dengan =COUNTIF(C2:C9,"LULUS").' };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

// ── VLOOKUP Tasks ──

async function checkAbsRef() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizVlookup");
      const cell = sheet.getRange("C2"); cell.load("formulas,values");
      await ctx.sync();
      const f = String(cell.formulas[0][0]).toUpperCase();
      if (f.includes("$E$2") && f.includes("B2")) {
        return { score: 100, correct: true, title: "Referensi Absolut Benar!", detail: `Rumus: ${cell.formulas[0][0]}, hasil: ${cell.values[0][0]}` };
      }
      if (f.includes("E2") && f.includes("B2")) {
        return { score: 50, partial: true, title: "Hampir Benar", detail: "Gunakan $E$2 bukan E2 agar tidak bergeser saat disalin." };
      }
      return { score: 0, correct: false, title: "Belum Ada", detail: "Isi C2 dengan =$E$2*B2" };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

async function checkVlookup() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizVlookup");
      const cell = sheet.getRange("G2"); cell.load("formulas,values");
      await ctx.sync();
      const f = String(cell.formulas[0][0]).toUpperCase();
      const v = String(cell.values[0][0]);
      if (f.includes("VLOOKUP") && f.includes("F2") && f.includes("A2")) {
        if (v.toLowerCase().includes("keyboard") || v.length > 3) {
          return { score: 100, correct: true, title: "VLOOKUP Benar!", detail: `Hasil pencarian: "${cell.values[0][0]}"` };
        }
        return { score: 60, partial: true, title: "VLOOKUP Ada", detail: `Hasilnya "${v}", cek kolom index (harus 2).` };
      }
      if (f.includes("VLOOKUP")) return { score: 40, partial: true, title: "VLOOKUP Ada", detail: "Periksa parameter: =VLOOKUP(F2,A2:B8,2,FALSE)" };
      return { score: 0, correct: false, title: "Belum VLOOKUP", detail: "Isi G2 dengan =VLOOKUP(F2,A2:B8,2,FALSE)" };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

async function checkVlookupFill() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizVlookup");
      const range = sheet.getRange("G2:G6"); range.load("values,formulas");
      await ctx.sync();
      let filled = 0;
      let hasAbs = false;
      for (let i = 0; i < 5; i++) {
        const f = String(range.formulas[i][0]).toUpperCase();
        const v = range.values[i][0];
        if (f.includes("VLOOKUP") && v && v !== 0) filled++;
        if (f.includes("$A$") || f.includes("$B$")) hasAbs = true;
      }
      if (filled >= 5 && hasAbs) return { score: 100, correct: true, title: "VLOOKUP Fill Benar!", detail: "Semua 5 baris terisi dengan referensi absolut yang tepat." };
      if (filled >= 3) return { score: Math.round((filled/5)*80), partial: true, title: `${filled}/5 Terisi`, detail: hasAbs ? "Bagus, pastikan semua baris terisi." : "Ingat pakai $A$2:$B$8 agar tidak error saat disalin." };
      return { score: 0, correct: false, title: "Belum Disalin", detail: "Ubah G2 pakai $A$2:$B$8 lalu salin ke G3:G6." };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}

async function checkSumIf() {
  return await Excel.run(async (ctx) => {
    try {
      const sheet = ctx.workbook.worksheets.getItem("QuizVlookup");
      const cell = sheet.getRange("I2"); cell.load("formulas,values");
      await ctx.sync();
      const f = String(cell.formulas[0][0]).toUpperCase();
      if (f.includes("SUMIF") && f.includes("ELEKTRONIK")) {
        return { score: 100, correct: true, title: "SUMIF Benar!", detail: `Total kategori Elektronik = ${cell.values[0][0]}` };
      }
      if (f.includes("SUMIF")) return { score: 50, partial: true, title: "SUMIF Ada", detail: 'Periksa kriteria "Elektronik" dan range sum.' };
      return { score: 0, correct: false, title: "Belum SUMIF", detail: 'Isi I2 dengan =SUMIF(A2:A8,"Elektronik",C2:C8).' };
    } catch(e) { return { score: 0, correct: false, title: "Error", detail: e.message }; }
  });
}
