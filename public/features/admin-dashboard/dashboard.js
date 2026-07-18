/**
 * ExamQuiz — Dashboard Logic
 */

let allResults = [];
let filtered = [];
let currentPage = 1;
const PAGE_SIZE = 5;
let sortKey = 'started_at';
let sortDir = 'desc';

let allParticipants = [];
let filteredParticipants = [];
let activeTab = 'exams';


async function init() {
  showLoading(true);
  const session = await SupabaseClient.getSession();

  // Jika tidak ada sesi, lempar ke login khusus admin
  if (!session) {
    window.location.href = '/admin/login';
    return;
  }

  const profile = await SupabaseClient.getUserProfile(session.user.id);

  // Proteksi: Jika bukan admin, tendang keluar
  if (!profile || profile.role !== 'admin') {
    console.warn('Akses ditolak: User bukan admin');
    await SupabaseClient.signOut();
    window.location.href = '/admin/login?error=unauthorized';
    return;
  }

  document.getElementById('header-username').textContent = profile.full_name || session.user.email;
  await loadResults();
  await loadCMS(); // Load CMS data
  showLoading(false);
}

// ─── CMS MANAGEMENT ───
async function loadCMS() {
  const grid = document.getElementById('cms-grid');
  try {
    const files = await SupabaseClient.getExamFiles();
    const iconMap = { word: '📝', excel: '📊', ppt: '📽' };

    grid.innerHTML = files.map(f => `
      <div class="cms-card">
        <div class="cms-card-top">
          <div class="cms-card-icon">${iconMap[f.exam_type] || '📄'}</div>
          <div class="cms-card-info">
            <div class="cms-card-name">${f.display_name}</div>
            <div class="cms-card-status ${f.is_available ? 'status-available' : 'status-missing'}">
              ${f.is_available ? '✓ Tersedia' : '⚠ Belum Ada'}
            </div>
          </div>
          ${f.is_available ? `
            <button class="btn-cms-delete" onclick="confirmDelete('${f.exam_type}', '${f.file_path}')" title="Hapus Soal">
              🗑️
            </button>
          ` : ''}
        </div>
        <div class="cms-upload-area">
          <input type="file" id="file-${f.exam_type}" class="input-file-hidden" 
                 onchange="onFileSelected('${f.exam_type}')" />
          <label for="file-${f.exam_type}" class="btn-upload-trigger" id="label-${f.exam_type}">
            📁 ${f.is_available ? 'Ganti File...' : 'Pilih File...'}
          </label>
          <button class="btn-cms-submit" id="btn-up-${f.exam_type}" disabled 
                  onclick="handleUpload('${f.exam_type}')">
            🚀 ${f.is_available ? 'Update Soal' : 'Unggah Soal'}
          </button>
        </div>
      </div>
    `).join('');
  } catch (err) {
    grid.innerHTML = `<div class="cms-card error">Gagal memuat CMS: ${err.message}</div>`;
  }
}

async function confirmDelete(type, path) {
  const result = await Swal.fire({
    title: 'Hapus Soal?',
    text: `Apakah Anda yakin ingin menghapus file soal ${type}? Peserta tidak akan bisa mendownloadnya lagi.`,
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#f87171',
    cancelButtonColor: '#4a5568',
    confirmButtonText: 'Ya, Hapus!',
    cancelButtonText: 'Batal'
  });

  if (result.isConfirmed) {
    try {
      await SupabaseClient.deleteExamFile(type, path);
      showToast('Soal berhasil dihapus', 'success');
      loadCMS();
    } catch (err) {
      showToast('Gagal menghapus: ' + err.message, 'error');
    }
  }
}

function onFileSelected(type) {
  const input = document.getElementById(`file-${type}`);
  const label = document.getElementById(`label-${type}`);
  const btn = document.getElementById(`btn-up-${type}`);

  if (input.files && input.files[0]) {
    label.textContent = '📄 ' + input.files[0].name;
    label.classList.add('has-file');
    btn.disabled = false;
  }
}

async function handleUpload(type) {
  const input = document.getElementById(`file-${type}`);
  const btn = document.getElementById(`btn-up-${type}`);
  if (!input.files || !input.files[0]) return;

  btn.disabled = true;
  btn.innerHTML = '⌛ Mengunggah...';

  try {
    await SupabaseClient.uploadExamFile(type, input.files[0]);
    showToast(`Berhasil mengunggah soal ${type}!`, 'success');
    await loadCMS();
  } catch (err) {
    console.error(err);
    showToast('Gagal unggah: ' + err.message, 'error');
    btn.disabled = false;
    btn.innerHTML = '🚀 Unggah Soal';
  }
}

async function loadResults() {
  try {
    const data = await SupabaseClient.getAllResults();
    
    allResults = (data || []).map(r => {
      const prof = r.profiles || {};
      return {
        id: r.id,
        name: prof.full_name || r.user_id,
        nim: prof.nim || '—',
        exam_type: r.exam_type,
        level: r.level,
        total_score: r.total_score || 0,
        max_score: r.max_score || 100,
        pct: r.max_score > 0 ? Math.round((r.total_score / r.max_score) * 100) : 0,
        status: r.status || 'in_progress',
        started_at: r.started_at,
        finished_at: r.finished_at
      };
    });

    updateStats();
    populateGelombangDropdown();
    filterAndRender();
  } catch (err) {
    console.error(err);
    showEmptyState('Gagal memuat: ' + err.message);
  }
}

function updateStats() {
  const finished = allResults.filter(r => r.status !== 'in_progress');
  const avg = finished.length ? Math.round(finished.reduce((s, r) => s + r.pct, 0) / finished.length) : 0;
  
  const getExamType = (type) => type === 'powerpoint' ? 'ppt' : type;

  document.getElementById('stat-total').textContent = allResults.length;
  document.getElementById('stat-avg').textContent = avg + '%';
  document.getElementById('stat-pass').textContent = allResults.filter(r => r.status === 'lulus').length;
  document.getElementById('stat-fail').textContent = allResults.filter(r => r.status === 'tidak_lulus').length;
  document.getElementById('stat-word').textContent = allResults.filter(r => getExamType(r.exam_type) === 'word').length;
  document.getElementById('stat-excel').textContent = allResults.filter(r => getExamType(r.exam_type) === 'excel').length;
  document.getElementById('stat-ppt').textContent = allResults.filter(r => getExamType(r.exam_type) === 'ppt').length;
}

function getLocalDateString(dateStr) {
  if (!dateStr) return '';
  const d = new Date(dateStr);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function populateGelombangDropdown() {
  const select = document.getElementById('filter-date');
  if (!select) return;
  
  const prevVal = select.value;
  select.innerHTML = '<option value="">Semua Gelombang</option>';
  
  const dates = new Set();
  allResults.forEach(r => {
    if (r.started_at) {
      const dStr = getLocalDateString(r.started_at);
      if (dStr) dates.add(dStr);
    }
  });
  
  const sortedDates = Array.from(dates).sort((a, b) => a.localeCompare(b));
  
  sortedDates.forEach((dStr, index) => {
    const formatted = new Date(dStr).toLocaleDateString('id-ID', { day: '2-digit', month: 'short', year: 'numeric' });
    const opt = document.createElement('option');
    opt.value = dStr;
    opt.textContent = `Gelombang ${index + 1} (${formatted})`;
    select.appendChild(opt);
  });
  
  if (sortedDates.includes(prevVal)) {
    select.value = prevVal;
  } else {
    select.value = '';
  }
}

function filterAndRender() {
  const search = document.getElementById('search-input').value.toLowerCase();
  const fType = document.getElementById('filter-type').value;
  const fStat = document.getElementById('filter-status').value;
  const fDate = document.getElementById('filter-date').value;

  filtered = allResults.filter(r => {
    const examType = r.exam_type === 'powerpoint' ? 'ppt' : r.exam_type;
    
    let dateMatch = true;
    if (fDate && r.started_at) {
      dateMatch = (getLocalDateString(r.started_at) === fDate);
    }

    return (r.name.toLowerCase().includes(search) || r.nim.toLowerCase().includes(search)) &&
      (!fType || examType === fType) &&
      (!fStat || r.status === fStat) &&
      dateMatch;
  });

  filtered.sort((a, b) => {
    let va = a[sortKey], vb = b[sortKey];
    if (typeof va === 'string') va = va.toLowerCase();
    if (typeof vb === 'string') vb = vb.toLowerCase();
    if (va < vb) return sortDir === 'asc' ? -1 : 1;
    if (va > vb) return sortDir === 'asc' ? 1 : -1;
    return 0;
  });

  currentPage = 1;
  renderTable();
  renderPagination();
}

function sortBy(key) {
  sortDir = (sortKey === key && sortDir === 'asc') ? 'desc' : 'asc';
  sortKey = key;
  filterAndRender();
}

function renderTable() {
  const tbody = document.getElementById('table-body');
  const start = (currentPage - 1) * PAGE_SIZE;
  const page = filtered.slice(start, start + PAGE_SIZE);

  if (!filtered.length) {
    tbody.innerHTML = '<tr><td colspan="11"><div class="empty-state"><div class="empty-state-icon">🔍</div><div class="empty-state-text">Tidak ada data</div></div></td></tr>';
    return;
  }

  const typeMap = {
    word: { class: 'badge-word', label: 'Word' },
    excel: { class: 'badge-excel', label: 'Excel' },
    ppt: { class: 'badge-ppt', label: 'PowerPoint' },
    powerpoint: { class: 'badge-ppt', label: 'PowerPoint' }
  };

  const statusMap = {
    lulus: { class: 'status-lulus', label: 'Lulus' },
    tidak_lulus: { class: 'status-tidak', label: 'Tidak Lulus' },
    in_progress: { class: 'status-in_progress', label: 'Berjalan' }
  };

  tbody.innerHTML = page.map((r, i) => {
    const typeInfo = typeMap[r.exam_type] || { class: '', label: '—' };
    const statusInfo = statusMap[r.status] || { class: '', label: '—' };
    const col = r.pct >= 70 ? 'var(--accent)' : r.pct >= 50 ? 'var(--accent3)' : 'var(--danger)';
    
    const startDt = r.started_at ? new Date(r.started_at).toLocaleString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
    const endDt = r.finished_at ? new Date(r.finished_at).toLocaleString('id-ID', { day: '2-digit', month: 'short', year: 'numeric', hour: '2-digit', minute: '2-digit' }) : '—';
    
    // Action column: report and delete
    const hasReport = r.status === 'lulus' || r.status === 'tidak_lulus';
    const reportBtn = hasReport 
      ? `<button class="btn-action btn-action-report" onclick="viewReport('${r.id}')" title="Lihat Laporan">
          <i data-lucide="eye" style="width:12px;height:12px;"></i> Laporan
         </button>` 
      : `<span style="color:var(--text-faint);font-size:11px;font-style:italic;margin-right:8px;">Belum Selesai</span>`;

    const deleteBtn = `<button class="btn-action btn-action-delete" onclick="deleteExamConfirm('${r.id}', '${x(r.name)}', '${x(typeInfo.label)}')" title="Hapus Ujian">
        <i data-lucide="trash-2" style="width:12px;height:12px;"></i> Hapus
       </button>`;

    return `<tr>
      <td style="color:var(--text-faint);font-family:var(--mono)">${start + i + 1}</td>
      <td><strong>${x(r.name)}</strong></td>
      <td style="font-family:var(--mono);font-size:12px;color:var(--text-dim)">${x(r.nim)}</td>
      <td><span class="badge ${typeInfo.class}">${typeInfo.label}</span></td>
      <td style="font-size:12px;color:var(--text-dim);text-transform:capitalize">${x(r.level)}</td>
      <td class="score-cell" style="color:${col}">${r.total_score}<span style="color:var(--text-faint);font-size:11px"> / ${r.max_score}</span></td>
      <td style="font-family:var(--mono);font-weight:700;color:${col}">${r.pct}%</td>
      <td><span class="status-badge ${statusInfo.class}">${statusInfo.label}</span></td>
      <td style="font-size:11px;color:var(--text-dim);font-family:var(--mono)">${startDt}</td>
      <td style="font-size:11px;color:var(--text-dim);font-family:var(--mono)">${endDt}</td>
      <td>
        <div class="action-cell">
          ${reportBtn}
          ${deleteBtn}
        </div>
      </td>
    </tr>`;
  }).join('');
  if (window._reinitIcons) window._reinitIcons();
}

function renderPagination() {
  const total = Math.ceil(filtered.length / PAGE_SIZE);
  const pag = document.getElementById('pagination');
  const start = filtered.length > 0 ? (currentPage - 1) * PAGE_SIZE + 1 : 0;
  const end = Math.min(currentPage * PAGE_SIZE, filtered.length);
  const infoText = `Menampilkan ${start} - ${end} dari ${filtered.length} data`;

  if (total <= 1) {
    pag.innerHTML = `<span class="page-info">${infoText}</span>`;
    return;
  }
  let html = `<span class="page-info">${infoText}</span>
    <button class="page-btn" onclick="changePage(${currentPage - 1})" ${currentPage === 1 ? 'disabled' : ''} title="Sebelumnya">
      <i data-lucide="chevron-left" style="width:14px;height:14px;vertical-align:middle;"></i>
    </button>`;
  for (let i = 1; i <= total; i++) {
    html += `<button class="page-btn ${i === currentPage ? 'active' : ''}" onclick="changePage(${i})">${i}</button>`;
  }
  html += `<button class="page-btn" onclick="changePage(${currentPage + 1})" ${currentPage === total ? 'disabled' : ''} title="Berikutnya">
      <i data-lucide="chevron-right" style="width:14px;height:14px;vertical-align:middle;"></i>
    </button>`;
  pag.innerHTML = html;
  if (window._reinitIcons) window._reinitIcons();
}

function changePage(p) {
  const total = Math.ceil(filtered.length / PAGE_SIZE);
  if (p < 1 || p > total) return;
  currentPage = p;
  renderTable();
  renderPagination();
}

function exportCSV() {
  const headers = ['No', 'Nama', 'NIM', 'Jenis Ujian', 'Level', 'Skor', 'Max Skor', 'Persentase', 'Status', 'Waktu Mulai', 'Waktu Selesai'];
  const rows = filtered.map((r, i) => [
    i + 1, r.name, r.nim, r.exam_type.toUpperCase(), r.level,
    r.total_score, r.max_score, r.pct + '%',
    r.status === 'lulus' ? 'Lulus' : r.status === 'tidak_lulus' ? 'Tidak Lulus' : 'Berjalan',
    r.started_at ? new Date(r.started_at).toLocaleString('id-ID') : '',
    r.finished_at ? new Date(r.finished_at).toLocaleString('id-ID') : ''
  ]);
  const csv = [headers, ...rows].map(r => r.map(v => `"${String(v).replace(/"/g, '""')}"`).join(',')).join('\n');
  const blob = new Blob(['\ufeff' + csv], { type: 'text/csv;charset=utf-8' });
  const a = Object.assign(document.createElement('a'), { href: URL.createObjectURL(blob), download: `hasil_ujian_${new Date().toISOString().slice(0, 10)}.csv` });
  a.click();
  showToast('CSV berhasil diunduh ✓', 'success');
}

async function handleSignOut() {
  await SupabaseClient.signOut();
  window.location.href = '/admin/login';
}

function showLoading(v) { document.getElementById('loading-overlay').classList.toggle('show', v); }
function showEmptyState(msg) {
  document.getElementById('table-body').innerHTML = `<tr><td colspan="11"><div class="empty-state"><div class="empty-state-icon">⚠️</div><div class="empty-state-text">${x(msg)}</div></div></td></tr>`;
}
function showToast(msg, type) {
  const t = document.getElementById('toast');
  t.textContent = msg; t.className = 'toast ' + (type || '') + ' show';
  setTimeout(() => t.className = 'toast', 3000);
}
function x(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

async function showAddParticipantModal() {
  const { value: formValues } = await Swal.fire({
    title: '➕ Tambah Peserta Baru',
    html: `
      <div style="text-align: left; font-family: 'DM Sans', sans-serif;">
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 11px; font-weight: 600; color: #8892a4; text-transform: uppercase; margin-bottom: 6px; font-family: 'Space Mono', monospace;">Nama Lengkap</label>
          <input id="swal-name" class="swal2-input" placeholder="Nama Lengkap Peserta" style="margin: 0; width: 100%; box-sizing: border-box; background: #1a1f2e; color: #e8edf5; border: 1px solid #252d3d; border-radius: 8px; padding: 10px; font-size: 14px;">
        </div>
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 11px; font-weight: 600; color: #8892a4; text-transform: uppercase; margin-bottom: 6px; font-family: 'Space Mono', monospace;">NIM</label>
          <input id="swal-nim" class="swal2-input" placeholder="Nomor Induk Mahasiswa" style="margin: 0; width: 100%; box-sizing: border-box; background: #1a1f2e; color: #e8edf5; border: 1px solid #252d3d; border-radius: 8px; padding: 10px; font-size: 14px;">
        </div>
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 11px; font-weight: 600; color: #8892a4; text-transform: uppercase; margin-bottom: 6px; font-family: 'Space Mono', monospace;">Email</label>
          <input id="swal-email" type="email" class="swal2-input" placeholder="nama@email.com" style="margin: 0; width: 100%; box-sizing: border-box; background: #1a1f2e; color: #e8edf5; border: 1px solid #252d3d; border-radius: 8px; padding: 10px; font-size: 14px;">
        </div>
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 11px; font-weight: 600; color: #8892a4; text-transform: uppercase; margin-bottom: 6px; font-family: 'Space Mono', monospace;">Password</label>
          <div style="display: flex; gap: 8px;">
            <input id="swal-password" type="text" class="swal2-input" placeholder="Min. 6 karakter" style="margin: 0; flex: 1; box-sizing: border-box; background: #1a1f2e; color: #e8edf5; border: 1px solid #252d3d; border-radius: 8px; padding: 10px; font-size: 14px;">
            <button type="button" style="background: #252d3d; border: 1px solid #4a5568; border-radius: 8px; color: #e8edf5; padding: 0 12px; cursor: pointer; font-size: 13px;" onclick="document.getElementById('swal-password').value = Math.random().toString(36).slice(-8)">🎲 Acak</button>
          </div>
        </div>
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 11px; font-weight: 600; color: #8892a4; text-transform: uppercase; margin-bottom: 6px; font-family: 'Space Mono', monospace;">Akses Ujian</label>
          <div style="display: flex; gap: 16px; margin-top: 6px; color: #e8edf5;">
            <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
              <input type="checkbox" id="add-swal-exam-word" value="word" checked style="cursor: pointer; width: 16px; height: 16px;"> Word
            </label>
            <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
              <input type="checkbox" id="add-swal-exam-excel" value="excel" checked style="cursor: pointer; width: 16px; height: 16px;"> Excel
            </label>
            <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
              <input type="checkbox" id="add-swal-exam-ppt" value="ppt" checked style="cursor: pointer; width: 16px; height: 16px;"> PPT
            </label>
          </div>
        </div>
      </div>
    `,
    background: '#131720',
    color: '#e8edf5',
    showCancelButton: true,
    confirmButtonText: '✓ Daftarkan',
    cancelButtonText: 'Batal',
    confirmButtonColor: '#4ade80',
    cancelButtonColor: '#4a5568',
    focusConfirm: false,
    preConfirm: () => {
      const name = document.getElementById('swal-name').value.trim();
      const nim = document.getElementById('swal-nim').value.trim();
      const email = document.getElementById('swal-email').value.trim();
      const password = document.getElementById('swal-password').value;

      const allowed = [];
      if (document.getElementById('add-swal-exam-word').checked) allowed.push('word');
      if (document.getElementById('add-swal-exam-excel').checked) allowed.push('excel');
      if (document.getElementById('add-swal-exam-ppt').checked) allowed.push('ppt');

      if (!name || !email || !password) {
        Swal.showValidationMessage('Nama Lengkap, Email, dan Password wajib diisi!');
        return false;
      }
      if (password.length < 6) {
        Swal.showValidationMessage('Password minimal 6 karakter!');
        return false;
      }
      if (allowed.length === 0) {
        Swal.showValidationMessage('Pilih minimal satu akses ujian!');
        return false;
      }
      return { name, nim, email, password, allowedExams: allowed.join(',') };
    }
  });

  if (formValues) {
    showLoading(true);
    try {
      await SupabaseClient.registerParticipantByAdmin(
        formValues.email,
        formValues.password,
        formValues.name,
        formValues.nim,
        formValues.allowedExams
      );

      Swal.fire({
        icon: 'success',
        title: 'Berhasil Terdaftar',
        text: `Peserta "${formValues.name}" berhasil ditambahkan!`,
        background: '#131720',
        color: '#e8edf5',
        confirmButtonColor: '#4ade80'
      });

      // Reload both lists
      await loadResults();
      await loadParticipants();
    } catch (err) {
      console.error('[DashboardUI] Register Error:', err);
      let errMsg = err.message || 'Gagal menambahkan peserta.';
      if (errMsg.includes('already registered')) {
        errMsg = 'Email sudah terdaftar di sistem.';
      }
      Swal.fire({
        icon: 'error',
        title: 'Pendaftaran Gagal',
        text: errMsg,
        background: '#131720',
        color: '#e8edf5',
        confirmButtonColor: '#f87171'
      });
    } finally {
      showLoading(false);
    }
  }
}

// ─── DASHBOARD TAB SWITCHING ───
function switchDashboardTab(tab) {
  activeTab = tab;
  
  // Update tab buttons state
  document.querySelectorAll('.dash-tab').forEach(b => b.classList.remove('active'));
  document.getElementById('tab-' + tab).classList.add('active');
  
  // Update section containers display
  document.querySelectorAll('.tab-content').forEach(c => {
    c.classList.remove('active');
    c.style.display = 'none';
  });
  
  const targetSec = document.getElementById('section-' + tab);
  targetSec.classList.add('active');
  targetSec.style.display = 'block';
  
  if (tab === 'participants') {
    loadParticipants();
  }
}

// ─── PARTICIPANT CRUD LOGIC ───
async function loadParticipants() {
  showLoading(true);
  try {
    const data = await SupabaseClient.getAllParticipants();
    allParticipants = data || [];
    filterAndRenderParticipants();
  } catch (err) {
    console.error('[DashboardUI] Load Participants Error:', err);
    showToast('Gagal memuat peserta: ' + err.message, 'error');
  } finally {
    showLoading(false);
  }
}

function filterAndRenderParticipants() {
  const searchInput = document.getElementById('search-participant');
  const search = searchInput ? searchInput.value.toLowerCase() : '';
  
  filteredParticipants = allParticipants.filter(p =>
    (p.full_name || '').toLowerCase().includes(search) ||
    (p.nim || '').toLowerCase().includes(search)
  );
  
  renderParticipantsTable();
}

function renderParticipantsTable() {
  const tbody = document.getElementById('participant-table-body');
  if (!tbody) return;
  
  if (!filteredParticipants.length) {
    tbody.innerHTML = '<tr><td colspan="6"><div class="empty-state"><div class="empty-state-icon">🔍</div><div class="empty-state-text">Tidak ada data peserta</div></div></td></tr>';
    return;
  }
  
  const typeMap = {
    word: { class: 'badge-word', label: 'Word' },
    excel: { class: 'badge-excel', label: 'Excel' },
    ppt: { class: 'badge-ppt', label: 'PowerPoint' },
    powerpoint: { class: 'badge-ppt', label: 'PowerPoint' }
  };

  tbody.innerHTML = filteredParticipants.map((p, i) => {
    const allowed = p.allowed_exams ? p.allowed_exams.split(',').map(s => s.trim().toLowerCase()) : ['word', 'excel', 'ppt'];
    const badgeHtml = allowed.map(type => {
      const typeInfo = typeMap[type] || { class: '', label: type.toUpperCase() };
      return `<span class="badge ${typeInfo.class}" style="margin-right: 4px; font-size: 10px; padding: 2px 6px;">${typeInfo.label}</span>`;
    }).join('');

    return `
      <tr>
        <td style="color:var(--text-faint);font-family:var(--mono)">${i + 1}</td>
        <td><strong>${x(p.full_name)}</strong></td>
        <td style="font-family:var(--mono);font-size:12px;color:var(--text-dim)">${x(p.nim)}</td>
        <td style="font-family:var(--mono);font-size:12px;color:var(--text-dim)">${x(p.email || '—')}</td>
        <td>${badgeHtml}</td>
        <td>
          <div class="action-cell">
            <button class="btn-action btn-action-edit" onclick="editParticipant('${p.id}', '${x(p.full_name)}', '${x(p.nim)}', '${x(p.allowed_exams || 'word,excel,ppt')}')">✏️ Edit</button>
            <button class="btn-action btn-action-delete" onclick="deleteParticipantConfirm('${p.id}', '${x(p.full_name)}')">🗑️ Hapus</button>
          </div>
        </td>
      </tr>
    `;
  }).join('');
}

let _activeReportSession = null;

async function viewReport(sessionId) {
  showLoading(true);
  try {
    const reportData = await SupabaseClient.getSessionReport(sessionId);
    if (!reportData) {
      throw new Error('Data laporan tidak ditemukan atau kosong');
    }
    
    _activeReportSession = reportData;
    
    const modalBody = document.getElementById('report-modal-body');
    if (window.ReportGenerator) {
      modalBody.innerHTML = ReportGenerator.buildHTML({
        candidateName: reportData.candidate ? reportData.candidate.name : '—',
        candidateNim: reportData.candidate ? reportData.candidate.nim : '—',
        totalScore: reportData.total_score,
        maxScore: reportData.max_score,
        examType: reportData.exam_type,
        level: reportData.level,
        date: reportData.started_at,
        answers: reportData.answers
      });
    } else {
      modalBody.innerHTML = '<div style="padding:20px;color:red;">Error: ReportGenerator utility not found!</div>';
    }
    
    const modal = document.getElementById('report-modal');
    modal.classList.add('show');
    if (window._reinitIcons) window._reinitIcons();
  } catch (err) {
    console.error('[DashboardUI] viewReport Error:', err);
    Swal.fire({
      icon: 'error',
      title: 'Gagal Memuat Laporan',
      text: err.message || 'Terjadi kesalahan saat memuat rincian laporan.'
    });
  } finally {
    showLoading(false);
  }
}

function closeReportModal() {
  const modal = document.getElementById('report-modal');
  modal.classList.remove('show');
}

function printModalReport() {
  if (!_activeReportSession || !window.ReportGenerator) return;
  const mapped = {
    sessionId: _activeReportSession.id,
    candidateName: _activeReportSession.candidate ? _activeReportSession.candidate.name : '—',
    candidateNim: _activeReportSession.candidate ? _activeReportSession.candidate.nim : '—',
    totalScore: _activeReportSession.total_score,
    maxScore: _activeReportSession.max_score,
    examType: _activeReportSession.exam_type,
    level: _activeReportSession.level,
    date: _activeReportSession.started_at,
    answers: _activeReportSession.answers
  };
  ReportGenerator.printReport(mapped);
}

// Bind functions to window context
window.viewReport = viewReport;
window.closeReportModal = closeReportModal;
window.printModalReport = printModalReport;

async function editParticipant(id, oldName, oldNim, oldAllowedExams) {
  const allowedList = oldAllowedExams ? oldAllowedExams.split(',').map(s => s.trim().toLowerCase()) : ['word', 'excel', 'ppt'];

  const { value: formValues } = await Swal.fire({
    title: '✏️ Edit Data Peserta',
    html: `
      <div style="text-align: left; font-family: 'DM Sans', sans-serif;">
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 11px; font-weight: 600; color: #8892a4; text-transform: uppercase; margin-bottom: 6px; font-family: 'Space Mono', monospace;">Nama Lengkap</label>
          <input id="edit-swal-name" class="swal2-input" value="${oldName}" placeholder="Nama Lengkap Peserta" style="margin: 0; width: 100%; box-sizing: border-box; background: #1a1f2e; color: #e8edf5; border: 1px solid #252d3d; border-radius: 8px; padding: 10px; font-size: 14px;">
        </div>
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 11px; font-weight: 600; color: #8892a4; text-transform: uppercase; margin-bottom: 6px; font-family: 'Space Mono', monospace;">NIM</label>
          <input id="edit-swal-nim" class="swal2-input" value="${oldNim}" placeholder="Nomor Induk Mahasiswa" style="margin: 0; width: 100%; box-sizing: border-box; background: #1a1f2e; color: #e8edf5; border: 1px solid #252d3d; border-radius: 8px; padding: 10px; font-size: 14px;">
        </div>
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 11px; font-weight: 600; color: #8892a4; text-transform: uppercase; margin-bottom: 6px; font-family: 'Space Mono', monospace;">Password Baru (Kosongkan jika tidak diubah)</label>
          <input id="edit-swal-password" type="password" class="swal2-input" placeholder="Masukkan password baru (min. 6 karakter)" style="margin: 0; width: 100%; box-sizing: border-box; background: #1a1f2e; color: #e8edf5; border: 1px solid #252d3d; border-radius: 8px; padding: 10px; font-size: 14px;">
        </div>
        <div style="margin-bottom: 14px;">
          <label style="display: block; font-size: 11px; font-weight: 600; color: #8892a4; text-transform: uppercase; margin-bottom: 6px; font-family: 'Space Mono', monospace;">Akses Ujian</label>
          <div style="display: flex; gap: 16px; margin-top: 6px; color: #e8edf5;">
            <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
              <input type="checkbox" id="edit-swal-exam-word" value="word" ${allowedList.includes('word') ? 'checked' : ''} style="cursor: pointer; width: 16px; height: 16px;"> Word
            </label>
            <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
              <input type="checkbox" id="edit-swal-exam-excel" value="excel" ${allowedList.includes('excel') ? 'checked' : ''} style="cursor: pointer; width: 16px; height: 16px;"> Excel
            </label>
            <label style="display: flex; align-items: center; gap: 6px; cursor: pointer;">
              <input type="checkbox" id="edit-swal-exam-ppt" value="ppt" ${allowedList.includes('ppt') ? 'checked' : ''} style="cursor: pointer; width: 16px; height: 16px;"> PPT
            </label>
          </div>
        </div>
      </div>
    `,
    background: '#131720',
    color: '#e8edf5',
    showCancelButton: true,
    confirmButtonText: '💾 Simpan',
    cancelButtonText: 'Batal',
    confirmButtonColor: '#22d3ee',
    cancelButtonColor: '#4a5568',
    focusConfirm: false,
    preConfirm: () => {
      const name = document.getElementById('edit-swal-name').value.trim();
      const nim = document.getElementById('edit-swal-nim').value.trim();
      const password = document.getElementById('edit-swal-password').value;

      const allowed = [];
      if (document.getElementById('edit-swal-exam-word').checked) allowed.push('word');
      if (document.getElementById('edit-swal-exam-excel').checked) allowed.push('excel');
      if (document.getElementById('edit-swal-exam-ppt').checked) allowed.push('ppt');

      if (!name) {
        Swal.showValidationMessage('Nama Lengkap wajib diisi!');
        return false;
      }
      if (password && password.length < 6) {
        Swal.showValidationMessage('Password minimal harus 6 karakter!');
        return false;
      }
      if (allowed.length === 0) {
        Swal.showValidationMessage('Pilih minimal satu akses ujian!');
        return false;
      }
      return { name, nim, password, allowedExams: allowed.join(',') };
    }
  });

  if (formValues) {
    showLoading(true);
    try {
      await SupabaseClient.updateParticipant(id, formValues.name, formValues.nim, formValues.allowedExams);
      if (formValues.password) {
        await SupabaseClient.updateParticipantPassword(id, formValues.password);
      }
      
      Swal.fire({
        icon: 'success',
        title: 'Berhasil Diperbarui',
        text: `Data peserta "${formValues.name}" berhasil diubah!`,
        background: '#131720',
        color: '#e8edf5',
        confirmButtonColor: '#4ade80'
      });

      // Reload both results and participants lists
      await loadResults();
      await loadParticipants();
    } catch (err) {
      console.error('[DashboardUI] Edit Participant Error:', err);
      Swal.fire({
        icon: 'error',
        title: 'Gagal Memperbarui',
        text: err.message || 'Terjadi kesalahan saat memperbarui data.',
        background: '#131720',
        color: '#e8edf5',
        confirmButtonColor: '#f87171'
      });
    } finally {
      showLoading(false);
    }
  }
}

async function deleteParticipantConfirm(id, name) {
  const result = await Swal.fire({
    title: '🗑️ Hapus Peserta?',
    text: `Apakah Anda yakin ingin menghapus peserta "${name}"? Seluruh data profil dan riwayat nilai kuis peserta ini akan dihapus secara permanen dari sistem. Tindakan ini tidak dapat dibatalkan!`,
    icon: 'warning',
    background: '#131720',
    color: '#e8edf5',
    showCancelButton: true,
    confirmButtonColor: '#f87171',
    cancelButtonColor: '#4a5568',
    confirmButtonText: 'Ya, Hapus Permanen!',
    cancelButtonText: 'Batal'
  });

  if (result.isConfirmed) {
    showLoading(true);
    try {
      await SupabaseClient.deleteParticipant(id);
      
      Swal.fire({
        icon: 'success',
        title: 'Berhasil Dihapus',
        text: `Peserta "${name}" dan seluruh riwayat nilainya telah dihapus.`,
        background: '#131720',
        color: '#e8edf5',
        confirmButtonColor: '#4ade80'
      });

      // Reload both tables
      await loadResults();
      await loadParticipants();
    } catch (err) {
      console.error('[DashboardUI] Delete Participant Error:', err);
      Swal.fire({
        icon: 'error',
        title: 'Gagal Menghapus',
        text: err.message || 'Terjadi kesalahan saat menghapus peserta.',
        background: '#131720',
        color: '#e8edf5',
        confirmButtonColor: '#f87171'
      });
    } finally {
      showLoading(false);
    }
  }
}

async function deleteExamConfirm(sessionId, studentName, examType) {
  const result = await Swal.fire({
    title: '🗑️ Hapus Hasil Ujian?',
    text: `Apakah Anda yakin ingin menghapus hasil ujian ${examType} untuk peserta "${studentName}"? Tindakan ini akan menghapus riwayat ujian secara permanen dan tidak dapat dibatalkan!`,
    icon: 'warning',
    background: '#131720',
    color: '#e8edf5',
    showCancelButton: true,
    confirmButtonColor: '#f87171',
    cancelButtonColor: '#4a5568',
    confirmButtonText: 'Ya, Hapus!',
    cancelButtonText: 'Batal'
  });

  if (result.isConfirmed) {
    showLoading(true);
    try {
      await SupabaseClient.deleteExamSession(sessionId);

      Swal.fire({
        icon: 'success',
        title: 'Berhasil Dihapus',
        text: `Hasil ujian ${examType} untuk "${studentName}" telah berhasil dihapus.`,
        background: '#131720',
        color: '#e8edf5',
        confirmButtonColor: '#4ade80'
      });

      // Reload results
      await loadResults();
    } catch (err) {
      console.error('[DashboardUI] Delete Exam Error:', err);
      Swal.fire({
        icon: 'error',
        title: 'Gagal Menghapus',
        text: err.message || 'Terjadi kesalahan saat menghapus hasil ujian.',
        background: '#131720',
        color: '#e8edf5',
        confirmButtonColor: '#f87171'
      });
    } finally {
      showLoading(false);
    }
  }
}

async function deleteAllExamsConfirm() {
  const result = await Swal.fire({
    title: '⚠️ Hapus Semua Ujian?',
    text: 'Apakah Anda yakin ingin menghapus SELURUH riwayat hasil ujian semua peserta? Tindakan ini akan menghapus semua data kuis secara permanen dari sistem dan tidak dapat dibatalkan!',
    icon: 'warning',
    background: '#131720',
    color: '#e8edf5',
    showCancelButton: true,
    confirmButtonColor: '#f87171',
    cancelButtonColor: '#4a5568',
    confirmButtonText: 'Ya, Hapus Semua!',
    cancelButtonText: 'Batal'
  });

  if (result.isConfirmed) {
    const confirmInput = await Swal.fire({
      title: 'Konfirmasi Ulang',
      text: 'Silakan ketik kata "HAPUS" untuk mengonfirmasi tindakan ini:',
      input: 'text',
      inputPlaceholder: 'HAPUS',
      background: '#131720',
      color: '#e8edf5',
      showCancelButton: true,
      confirmButtonColor: '#f87171',
      cancelButtonColor: '#4a5568',
      confirmButtonText: 'Konfirmasi',
      cancelButtonText: 'Batal',
      preConfirm: (value) => {
        if (value.trim().toUpperCase() !== 'HAPUS') {
          Swal.showValidationMessage('Kata konfirmasi tidak cocok!');
          return false;
        }
        return true;
      }
    });

    if (confirmInput.isConfirmed) {
      showLoading(true);
      try {
        await SupabaseClient.deleteAllExamSessions();

        Swal.fire({
          icon: 'success',
          title: 'Berhasil Dihapus',
          text: 'Seluruh riwayat hasil ujian peserta telah dihapus dari sistem.',
          background: '#131720',
          color: '#e8edf5',
          confirmButtonColor: '#4ade80'
        });

        // Reload results
        await loadResults();
      } catch (err) {
        console.error('[DashboardUI] Delete All Exams Error:', err);
        Swal.fire({
          icon: 'error',
          title: 'Gagal Menghapus',
          text: err.message || 'Terjadi kesalahan saat menghapus seluruh hasil ujian.',
          background: '#131720',
          color: '#e8edf5',
          confirmButtonColor: '#f87171'
        });
      } finally {
        showLoading(false);
      }
    }
  }
}

// Bind delete functions to window context
window.deleteExamConfirm = deleteExamConfirm;
window.deleteAllExamsConfirm = deleteAllExamsConfirm;

init();
