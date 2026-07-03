/**
 * ExamQuiz — Taskpane Main Logic
 * Refactored Version
 */

let currentUser = null;
let currentProfile = null;
let examSessionId = null;
let _lastResults = null; // for report download

let state = {
  host: null,
  examKey: 'praktik',
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
   OFFICE & AUTH INIT
═══════════════════════════════════════════════ */
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Excel) state.host = 'Excel';
  else if (info.host === Office.HostType.Word) state.host = 'Word';
  else if (info.host === Office.HostType.PowerPoint) state.host = 'PowerPoint';

  try {
    await ExamsLoader.loadExamsData();
  } catch (err) {
    console.error('[ExamQuiz] Gagal memuat soal:', err);
    document.getElementById('header-sub').textContent = 'Gagal memuat data soal';
  }

  // Force Logout on fresh start to ensure "Login First" requirement
  await SupabaseClient.signOut();
  showScreen('login');
});

function onAuthReady() {
  const displayName = (currentProfile && currentProfile.full_name) || (currentUser && currentUser.email) || 'User';
  document.getElementById('header-user').textContent = displayName.split(' ')[0];
  document.getElementById('header-actions').style.display = 'flex';

  if (currentProfile && currentProfile.role === 'admin') {
    window.location.href = '/admin';
    return;
  }

  if (!state.host) {
    showScreen('web-chooser');
    document.getElementById('header-sub').textContent = 'Pilih & Download Soal';
    loadWebChooser(); // Load dynamic files
    return;
  }

  initApp();
}

async function loadWebChooser() {
  const grid = document.getElementById('web-chooser-grid');
  try {
    const files = await SupabaseClient.getExamFiles();
    const iconMap = { word: '📝', excel: '📊', ppt: '📽' };

    const allowed = (currentProfile && currentProfile.allowed_exams)
      ? currentProfile.allowed_exams.split(',').map(s => s.trim().toLowerCase())
      : ['word', 'excel', 'ppt'];

    const filteredFiles = files.filter(f => {
      const type = f.exam_type.toLowerCase() === 'ppt' ? 'ppt' : f.exam_type.toLowerCase();
      return allowed.includes(type);
    });

    if (filteredFiles.length === 0) {
      grid.innerHTML = `<div class="cms-card warning" style="text-align: center; padding: 20px; color: var(--text-dim); background: rgba(224, 160, 64, 0.1); border: 1px solid var(--warning); border-radius: 8px;">Anda tidak memiliki akses ke kuis apa pun saat ini.</div>`;
      return;
    }

    grid.innerHTML = filteredFiles.map(f => `
      <div class="chooser-card">
        <div class="chooser-card-icon">${iconMap[f.exam_type] || '📄'}</div>
        <h3>${f.display_name}</h3>
        <p>${f.is_available ? 'Soal sudah tersedia dan dapat dikerjakan.' : 'Maaf, soal belum diunggah oleh admin.'}</p>
        ${f.is_available
        ? `<a href="${f.file_path}" download class="btn btn-primary">⬇ Download Soal</a>`
        : `<button class="btn btn-primary" disabled>⌛ Belum Tersedia</button>`
      }
      </div>
    `).join('');
  } catch (err) {
    grid.innerHTML = `<div class="cms-card error">Gagal memuat daftar soal: ${err.message}</div>`;
  }
}

function initApp() {
  const hostLower = state.host.toLowerCase() === 'powerpoint' ? 'ppt' : state.host.toLowerCase();
  const allowed = (currentProfile && currentProfile.allowed_exams)
    ? currentProfile.allowed_exams.split(',').map(s => s.trim().toLowerCase())
    : ['word', 'excel', 'ppt'];

  const accessDeniedEl = document.getElementById('welcome-access-denied');
  const examSelectorEl = document.querySelector('.exam-selector');
  const infoGridEl = document.getElementById('exam-info-grid');
  const btnStartEl = document.getElementById('btn-start');

  if (!allowed.includes(hostLower)) {
    if (accessDeniedEl) {
      accessDeniedEl.classList.remove('d-none');
      accessDeniedEl.textContent = `Maaf, Anda tidak memiliki akses untuk ujian Microsoft ${state.host}.`;
    }
    if (examSelectorEl) examSelectorEl.style.display = 'none';
    if (infoGridEl) infoGridEl.style.display = 'none';
    if (btnStartEl) btnStartEl.style.display = 'none';

    document.getElementById('welcome-title').textContent = `Ujian Microsoft ${state.host} Terbatas`;
    document.getElementById('welcome-icon').textContent = '⚠️';
    document.getElementById('header-title').textContent = `ExamQuiz — Terbatas`;
    showScreen('welcome');
    return;
  } else {
    if (accessDeniedEl) accessDeniedEl.classList.add('d-none');
    if (examSelectorEl) examSelectorEl.style.display = 'block';
    if (infoGridEl) infoGridEl.style.display = 'grid';
    if (btnStartEl) btnStartEl.style.display = 'inline-flex';
  }

  const select = document.getElementById('exam-select');
  select.innerHTML = '';

  let dataMap, icon;
  if (state.host === 'Excel') { dataMap = EXAMS; icon = '📊'; }
  else if (state.host === 'Word') { dataMap = WORD_EXAMS; icon = '📝'; }
  else if (state.host === 'PowerPoint') { dataMap = POWERPOINT_EXAMS; icon = '📊'; }

  for (let key in dataMap) {
    const opt = document.createElement('option');
    opt.value = key;
    opt.textContent = `${icon} ${dataMap[key].name} (${Math.round(dataMap[key].duration / 60)}m)`;
    select.appendChild(opt);
  }

  select.addEventListener('change', updateExamInfo);
  updateExamInfo();

  document.getElementById('welcome-icon').textContent = icon || '📋';
  document.getElementById('welcome-title').textContent = `Ujian Microsoft ${state.host}`;
  document.getElementById('header-title').textContent = `ExamQuiz — ${state.host}`;

  showScreen('welcome');
}

function updateExamInfo() {
  const key = document.getElementById('exam-select').value;
  const data = state.host === 'Excel' ? EXAMS : (state.host === 'Word' ? WORD_EXAMS : POWERPOINT_EXAMS);
  const exam = data[key];
  if (!exam) return;

  document.getElementById('info-total-q').textContent = exam.tasks.length;
  document.getElementById('info-duration').textContent = Math.round(exam.duration / 60) + 'm';
  document.getElementById('info-points').textContent = exam.tasks.reduce((s, t) => s + t.points, 0);
}

/* ═══════════════════════════════════════════════
   EXAM FLOW
═══════════════════════════════════════════════ */
document.getElementById('btn-start').addEventListener('click', startExam);

async function startExam() {
  const key = document.getElementById('exam-select').value;
  const data = state.host === 'Excel' ? EXAMS : (state.host === 'Word' ? WORD_EXAMS : POWERPOINT_EXAMS);
  state.exam = data[key];
  state.examKey = key;
  state.currentIdx = 0;
  state.confirmed = new Array(state.exam.tasks.length).fill(false);
  state.scores = new Array(state.exam.tasks.length).fill(0);
  state.timeLeft = state.exam.duration;
  state.started = true;

  document.getElementById('timer-bar').style.display = 'flex';
  document.getElementById('progress-outer').style.display = 'block';

  // DB Session
  const examType = state.host.toLowerCase() === 'powerpoint' ? 'ppt' : state.host.toLowerCase();
  const session = await SupabaseClient.createExamSession(currentUser.id, examType, key, 100);
  if (session) examSessionId = session.id;

  startTimer();
  buildDotsNav();
  showTask(0);
  showScreen('task');
  showToast('Ujian dimulai! Semangat 🚀', 'info');
}

function startTimer() {
  clearInterval(state.timerInterval);
  state.timerInterval = setInterval(() => {
    state.timeLeft--;
    const m = Math.floor(state.timeLeft / 60);
    const s = state.timeLeft % 60;
    document.getElementById('timer-display').textContent = `${String(m).padStart(2, '0')}:${String(s).padStart(2, '0')}`;
    if (state.timeLeft <= 0) finishExam();
  }, 1000);
}

// Event listener untuk klik badge nomor soal (pop up grid lompat soal)
document.getElementById('task-badge').addEventListener('click', () => {
  if (state.started && !state.finished) {
    toggleQuestionPopup(true);
  }
});

window.toggleQuestionPopup = function (show) {
  const popup = document.getElementById('question-popup');
  if (!popup) return;
  if (show) {
    const grid = document.getElementById('question-popup-grid');
    grid.innerHTML = state.exam.tasks.map((t, i) => {
      let classes = 'question-grid-item';
      if (i === state.currentIdx) classes += ' active';
      else if (state.confirmed[i]) classes += ' done';
      return `<div class="${classes}" onclick="jumpToQuestion(${i})">${i + 1}</div>`;
    }).join('');
    popup.style.display = 'flex';
  } else {
    popup.style.display = 'none';
  }
};

window.jumpToQuestion = function (idx) {
  showTask(idx);
  toggleQuestionPopup(false);
};

function showTask(idx) {
  state.currentIdx = idx;
  const task = state.exam.tasks[idx];
  document.getElementById('task-badge').innerHTML = `Soal ${idx + 1} / ${state.exam.tasks.length} <span style="font-size: 8px; opacity: 0.6; margin-left: 2px;">▼</span>`;
  document.getElementById('task-title').textContent = task.title;
  document.getElementById('task-desc').textContent = task.desc;
  document.getElementById('task-steps').innerHTML = task.steps.map((s, i) => `<div class="step-item"><div class="step-num">${i + 1}</div><div>${s}</div></div>`).join('');

  // Update Nav Buttons
  document.getElementById('btn-prev').disabled = (idx === 0);
  document.getElementById('btn-next').disabled = (idx === state.exam.tasks.length - 1);

  // Show / Hide Confirm Section
  const confirmSection = document.getElementById('confirm-section');
  if (confirmSection) {
    if (task.isConfirm) {
      confirmSection.classList.remove('d-none');
      updateConfirmBtn(idx);
    } else {
      confirmSection.classList.add('d-none');
    }
  }

  updateDots();
}

function updateConfirmBtn(idx) {
  const btn = document.getElementById('btn-confirm-task');
  if (!btn) return;
  if (state.confirmed[idx]) {
    btn.classList.add('done');
    btn.innerHTML = `<span id="confirm-icon">✅</span> &nbsp;Tugas Telah Dikonfirmasi ✓`;
  } else {
    btn.classList.remove('done');
    btn.innerHTML = `<span id="confirm-icon">⬜</span> &nbsp;Saya Sudah Mengerjakan Tugas Ini`;
  }
}

window.toggleTaskConfirm = function () {
  const idx = state.currentIdx;
  state.confirmed[idx] = !state.confirmed[idx];
  updateConfirmBtn(idx);
  updateDots();
};

window.prevTask = function () {
  if (state.currentIdx > 0) showTask(state.currentIdx - 1);
};

window.nextTask = function () {
  if (state.currentIdx < state.exam.tasks.length - 1) showTask(state.currentIdx + 1);
};

function updateDots() {
  document.querySelectorAll('.dot').forEach((d, i) => {
    d.className = 'dot' + (i === state.currentIdx ? ' active current' : '');
  });
}

window.submitExam = async function () {
  const result = await Swal.fire({
    title: 'Submit Ujian?',
    text: 'Anda yakin ingin mengakhiri ujian dan melihat hasil sekarang?',
    icon: 'question',
    showCancelButton: true,
    confirmButtonColor: '#4ade80',
    cancelButtonColor: '#1c2030',
    confirmButtonText: 'Ya, Submit!',
    cancelButtonText: 'Batal',
    background: '#151820',
    color: '#e8edf5'
  });

  if (result.isConfirmed) {
    finishExam();
  }
};

async function finishExam() {
  if (state.finished) return;
  state.finished = true;
  clearInterval(state.timerInterval);

  const overlay = document.getElementById('scoring-overlay');
  overlay.classList.add('show');

  const tasks = state.exam.tasks;
  const results = [];

  for (let i = 0; i < tasks.length; i++) {
    document.getElementById('scoring-step-text').textContent = `Menilai soal ${i + 1}...`;
    let res = { score: 0, detail: 'Tidak dikerjakan' };

    // Check all tasks automatically on submit
    try {
      if (tasks[i].isConfirm) {
        res = state.confirmed[i] ? { score: 100, detail: 'Dikonfirmasi ✓' } : { score: 0, detail: 'Belum dikonfirmasi' };
      } else if (typeof tasks[i].check === 'function') {
        res = await tasks[i].check();
      } else {
        // Fallback for tasks with no check function but confirmed
        res = state.confirmed[i] ? { score: 100, detail: 'Dikonfirmasi ✓' } : { score: 0, detail: 'Belum dikonfirmasi' };
      }
    } catch (e) {
      console.error(`Error verifying task ${i + 1}:`, e);
      let msg = e.message || 'Verifikasi gagal';
      if (msg.includes('cell-editing mode')) {
        msg = 'Tekan ENTER di Excel sebelum Submit!';
      }
      res = { score: 0, detail: `Error: ${msg}` };
    }

    results.push({
      task: tasks[i],
      pts: Math.round((res.score / 100) * tasks[i].points),
      detail: res.detail,
      status: res.score >= 70 ? 'pass' : (res.score > 0 ? 'partial' : 'fail')
    });
    await new Promise(r => setTimeout(r, 200));
  }

  // Save Results
  const total = results.reduce((s, r) => s + r.pts, 0);
  await SupabaseClient.saveExamResults(examSessionId, total, 100, results.map(r => ({ questionId: r.task.id, title: r.task.title, score: r.pts, status: r.status, detail: r.detail })));

  overlay.classList.remove('show');
  showResults(results, total);
}

function showResults(results, total) {
  _lastResults = results;
  const maxScore = 100;
  const passed = total >= 70;

  // Summary bar
  document.getElementById('final-score-num').textContent = total;
  const gradeEl = document.getElementById('result-grade');
  gradeEl.textContent = passed ? 'LULUS' : 'TIDAK LULUS';
  gradeEl.style.color = passed ? '#16a34a' : '#dc2626';

  // Build clean, sidebar-friendly breakdown list
  const reportArea = document.getElementById('result-report-area');
  if (reportArea) {
    // Style the container to be transparent and fit the taskpane list design
    reportArea.style.background = 'transparent';
    reportArea.style.border = 'none';
    reportArea.style.boxShadow = 'none';
    reportArea.style.maxHeight = 'none';
    reportArea.style.overflow = 'visible';
    reportArea.style.padding = '0';

    const examType = state.host ? state.host.toLowerCase() : 'excel';
    const answers = results.map((r, i) => ({
      no: i + 1,
      title: r.task.title,
      score: r.pts,
      max: r.task.points,
      detail: r.detail
    }));

    const categories = window.ReportGenerator.groupByCategory(answers, examType);

    const breakdownHTML = `
      <div class="result-breakdown" style="margin-bottom: 0;">
        <div class="breakdown-title" style="font-size: 11px; font-weight: 700; color: var(--text-dim); margin-bottom: 12px; font-family: var(--mono); text-transform: uppercase; letter-spacing: 0.08em;">
          Rincian Hasil Pengerjaan
        </div>
        ${categories.map((cat, i) => {
          const catPct = cat.max > 0 ? Math.round((cat.score / cat.max) * 100) : 0;
          let statusClass = 'fail';
          let icon = '❌';
          if (catPct >= 70) {
            statusClass = 'pass';
            icon = '✓';
          } else if (cat.score > 0) {
            statusClass = 'partial';
            icon = '⚠';
          }

          const subItemsHTML = cat.items.map(item => {
            let itemColor = 'var(--danger)';
            let itemIcon = '❌';
            if (item.score >= item.max) {
              itemColor = 'var(--success)';
              itemIcon = '✓';
            } else if (item.score > 0) {
              itemColor = 'var(--warning)';
              itemIcon = '⚠';
            }
            return `
              <div style="margin-top: 4px; padding-left: 8px; border-left: 2px solid ${itemColor === 'var(--success)' ? 'rgba(46,158,107,0.3)' : (itemColor === 'var(--warning)' ? 'rgba(224,160,64,0.3)' : 'rgba(224,82,82,0.3)')};">
                <div style="display: flex; justify-content: space-between; font-size: 11px;">
                  <span style="font-weight: 500; color: var(--text-dim);">${item.title}</span>
                  <span style="font-family: var(--mono); font-weight: bold; color: ${itemColor};">${item.score}/${item.max}</span>
                </div>
                <div style="font-size: 9px; color: var(--text-faint); margin-top: 1px; font-family: var(--mono);">${item.detail || 'Tidak dikerjakan'}</div>
              </div>
            `;
          }).join('');

          return `
            <div class="breakdown-item" style="margin-bottom: 8px; display: block; padding: 10px 12px;">
              <div style="display: flex; align-items: center; gap: 10px; margin-bottom: 6px;">
                <div class="breakdown-status ${statusClass}">${icon}</div>
                <div class="breakdown-info" style="flex: 1; min-width: 0;">
                  <div class="breakdown-name" style="font-size: 12px; font-weight: 700; color: var(--text); line-height: 1.2;">${cat.name}</div>
                </div>
                <div class="breakdown-pts ${statusClass}" style="font-size: 12px; font-family: var(--mono); font-weight: 700;">${cat.score} / ${cat.max}</div>
              </div>
              <div class="breakdown-subitems" style="display: flex; flex-direction: column; gap: 4px; border-top: 1px solid var(--border-light); padding-top: 6px; margin-top: 6px;">
                ${subItemsHTML}
              </div>
            </div>
          `;
        }).join('')}
      </div>
    `;
    reportArea.innerHTML = breakdownHTML;
  }

  showScreen('result');
  if (window._reinitIcons) window._reinitIcons();
}

/* ═══════════════════════════════════════════════
   UI HELPERS
═══════════════════════════════════════════════ */
function showScreen(name) {
  document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
  document.getElementById('screen-' + name).classList.add('active');
}

function showToast(msg, type) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast ' + type + ' show';
  setTimeout(() => t.classList.remove('show'), 3000);
}
function updateDots() {
  document.querySelectorAll('.dot').forEach((d, i) => {
    d.className = 'dot' + (i === state.currentIdx ? ' active current' : '') + (state.confirmed[i] ? ' done' : '');
  });
}
function buildDotsNav() {
  document.getElementById('dots-nav').innerHTML = state.exam.tasks.map((_, i) => `<div class="dot" onclick="showTask(${i})"></div>`).join('');
}
function updateProgress() {
  // Progress is now just current index / total
  document.getElementById('progress-bar').style.width = ((state.currentIdx + 1) / state.exam.tasks.length * 100) + '%';
}

/* ═══════════════════════════════════════════════
   AUTH UI HANDLERS (for taskpane login screen)
═══════════════════════════════════════════════ */
/* lpSwitchTab and handleLpRegister removed — register form no longer exists */

window.handleLpLogin = async function () {
  const email = document.getElementById('lp-email').value;
  const pw = document.getElementById('lp-password').value;
  const btn = document.getElementById('btn-lp-login');

  try {
    if (btn) btn.disabled = true;
    const { user, session } = await SupabaseClient.signIn(email, pw);
    currentUser = user;
    currentProfile = await SupabaseClient.getUserProfile(user.id);

    Swal.fire({ icon: 'success', title: 'Berhasil', text: 'Login berhasil!', timer: 1000, showConfirmButton: false });
    setTimeout(() => onAuthReady(), 1000);
  } catch (e) {
    let msg = e.message || 'Login gagal';
    if (msg.includes('Invalid login credentials') || msg.includes('Invalid login') || msg.includes('invalid_credentials')) {
      msg = 'Email atau password salah. Silakan periksa kembali data Anda.';
    } else if (msg.includes('Email not confirmed')) {
      msg = 'Email Anda belum dikonfirmasi. Silakan verifikasi email Anda terlebih dahulu.';
    } else if (msg.includes('rate limit') || msg.includes('Rate limit exceeded')) {
      msg = 'Terlalu banyak percobaan login. Silakan tunggu beberapa menit sebelum mencoba lagi.';
    } else if (msg.includes('network') || msg.includes('Failed to fetch')) {
      msg = 'Gagal terhubung ke server/jaringan. Silakan periksa koneksi internet Anda.';
    }
    Swal.fire({ icon: 'error', title: 'Login Gagal', text: msg });
    if (btn) btn.disabled = false;
  }
};

window.handleSignOut = function () {
  SupabaseClient.signOut().then(() => {
    window.location.href = '/login';
  });
};

window.showTask = showTask;
window.restartExam = () => window.location.reload();
window.exportResult = () => Swal.fire('Info', 'Fitur ini akan segera hadir!', 'info');

window.downloadResultReport = function () {
  if (!_lastResults || !window.ReportGenerator) {
    Swal.fire('Gagal', 'Data laporan belum tersedia.', 'error');
    return;
  }
  const total = _lastResults.reduce((s, r) => s + r.pts, 0);
  const answers = _lastResults.map((r, i) => ({
    no: i + 1,
    title: r.task.title,
    score: r.pts,
    max: r.task.points
  }));
  ReportGenerator.printReport({
    sessionId: examSessionId,
    candidateName: currentProfile ? currentProfile.full_name : (currentUser ? currentUser.email : '—'),
    candidateNim: currentProfile ? (currentProfile.nim || '—') : '—',
    totalScore: total,
    maxScore: 100,
    examType: state.host ? state.host.toLowerCase() : 'excel',
    level: state.examKey || 'basic',
    date: new Date().toISOString(),
    answers
  });
};
