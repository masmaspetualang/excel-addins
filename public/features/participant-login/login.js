/**
 * ExamQuiz — Login Page Logic
 */

// Redirect if already logged in
(async () => {
  const session = await SupabaseClient.getSession();
  if (session) {
    const redirect = new URLSearchParams(window.location.search).get('redirect') || '/app';
    window.location.href = redirect;
  }
})();


// ─── ALERT ───────────────────────────────────────
function showAlert(msg, type) {
  const el = document.getElementById('alert-msg');
  el.textContent = msg;
  el.className = 'alert ' + type + ' show';
}
function clearAlert() {
  document.getElementById('alert-msg').className = 'alert';
}

// ─── PASSWORD TOGGLE ─────────────────────────────
function togglePw(inputId, btnId) {
  const inp = document.getElementById(inputId);
  const btn = document.getElementById(btnId);
  if (inp.type === 'password') {
    inp.type = 'text';
    btn.textContent = '🙈';
  } else {
    inp.type = 'password';
    btn.textContent = '👁';
  }
}

// ─── LOADING STATE ───────────────────────────────
function setLoading(btnId, loading) {
  const btn = document.getElementById(btnId);
  if (loading) {
    btn.classList.add('loading');
    btn.disabled = true;
  } else {
    btn.classList.remove('loading');
    btn.disabled = false;
  }
}

// ─── LOGIN ───────────────────────────────────────
async function handleLogin(e) {
  e.preventDefault();
  clearAlert();

  const email = document.getElementById('login-email').value.trim();
  const password = document.getElementById('login-password').value;

  console.log('[LoginUI] Login clicked for:', email);

  setLoading('btn-login', true);
  try {
    const { user } = await SupabaseClient.signIn(email, password);
    const profile = await SupabaseClient.getUserProfile(user.id);

    console.log('[LoginUI] Login success, user profile:', profile);

    Swal.fire({
      icon: 'success',
      title: 'Login Berhasil',
      text: 'Membuka halaman ujian...',
      timer: 1500,
      showConfirmButton: false
    });

    setTimeout(() => {
      const redirect = new URLSearchParams(window.location.search).get('redirect') || '/app';
      window.location.href = redirect;
    }, 1500);
  } catch (err) {
    console.error('[LoginUI] Login Error:', err);
    let msg = err.message || 'Login gagal';
    
    if (msg.includes('Invalid login credentials') || msg.includes('Invalid login') || msg.includes('invalid_credentials')) {
      msg = 'Email atau password salah. Silakan periksa kembali data Anda.';
    } else if (msg.includes('Email not confirmed')) {
      msg = 'Email Anda belum dikonfirmasi. Silakan verifikasi email Anda terlebih dahulu.';
    } else if (msg.includes('rate limit') || msg.includes('Rate limit exceeded')) {
      msg = 'Terlalu banyak percobaan login. Silakan tunggu beberapa menit sebelum mencoba lagi.';
    } else if (msg.includes('network') || msg.includes('Failed to fetch')) {
      msg = 'Gagal terhubung ke server/jaringan. Silakan periksa koneksi internet Anda.';
    }

    Swal.fire({
      icon: 'error',
      title: 'Login Gagal',
      text: msg
    });
  } finally {
    setLoading('btn-login', false);
  }
}

