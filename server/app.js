/**
 * server/app.js — Express app (development lokal)
 */
const express = require('express');
const path = require('path');
const corsMiddleware = require('./middleware/cors');
const rateLimitMiddleware = require('./middleware/rate-limit');
const { notFoundHandler, serverErrorHandler } = require('./middleware/error');
const logger = require('./utils/logger');
const MIME_TYPES = require('./config/mime-types.json');
const { FILES, URLS, LEGACY_REDIRECTS } = require('./config/routes');

const app = express();
const publicDir = path.join(__dirname, '../public');

if (express.static.mime && typeof express.static.mime.define === 'function') {
  express.static.mime.define(MIME_TYPES, true);
}

app.use((req, res, next) => {
  res.on('finish', () => logger.request(req, res.statusCode));
  next();
});

app.use(corsMiddleware);
app.use('/api', rateLimitMiddleware);

function servePage(relativePath) {
  return (_req, res) => res.sendFile(path.join(publicDir, relativePath));
}

// Root → app utama
app.get('/', (_req, res) => res.redirect(302, URLS.app));

// URL profesional
app.get(URLS.app, servePage(FILES.app));
app.get(URLS.login, servePage(FILES.login));
app.get(URLS.admin, servePage(FILES.admin));
app.get(URLS.adminLogin, servePage(FILES.adminLogin));
app.get(URLS.adminCommands, servePage(FILES.adminCommands));

// Config client dari .env
const { toJavaScript } = require('./config/client-config');
app.get('/js/config/app.config.js', (_req, res) => {
  res.setHeader('Content-Type', 'application/javascript; charset=utf-8');
  res.setHeader('Cache-Control', 'no-store');
  res.send(toJavaScript());
});
app.get('/js/config/supabase.config.js', (_req, res) => {
  res.redirect(301, '/js/config/app.config.js');
});

// ─── SERVER API ─────────────────────────────────────────
// API: Participants with Email (menggunakan service key dengan fallback aman)
app.get('/api/participants-with-email', async (_req, res) => {
  try {
    res.setHeader('Cache-Control', 'no-store');
    let supabase;
    let users = [];

    try {
      const { createSupabaseAdmin } = require('./lib/supabase-admin');
      supabase = createSupabaseAdmin();

      const { data, error: authErr } = await supabase.auth.admin.listUsers({ perPage: 1000 });
      if (authErr) {
        console.warn('[API] Warning: Failed to list auth users via admin client:', authErr.message);
      } else {
        users = data.users || [];
      }
    } catch (adminErr) {
      console.warn('[API] Warning: Cannot initialize admin client or fetch auth users. Falling back to anon client. Error:', adminErr.message);
      // Fallback: use normal supabase client with anon key
      const { createClient } = require('@supabase/supabase-js');
      const env = require('./config/env');
      supabase = createClient(env.supabaseUrl, env.supabaseAnonKey);
    }

    // Fetch all participants from pengguna table
    const { data: participants, error: dbErr } = await supabase
      .from('pengguna')
      .select('*')
      .eq('peran', 'participant')
      .order('nama_lengkap', { ascending: true });
    if (dbErr) throw dbErr;

    // Merge email from auth.users
    const emailMap = {};
    (users || []).forEach(u => { emailMap[u.id] = u.email; });

    const result = (participants || []).map(p => ({
      id: p.id_pengguna,
      full_name: p.nama_lengkap,
      nim: p.nim,
      // Check database column p.email first, then merge from auth.users, then generate fallback from NIM
      email: p.email || emailMap[p.id_pengguna] || (p.nim ? `${p.nim}@student.umy.ac.id` : '—'),
      role: p.peran,
      allowed_exams: p.allowed_exams || 'word,excel,ppt'
    }));

    res.json({ data: result });
  } catch (err) {
    console.error('[API] participants-with-email error:', err);
    res.status(500).json({ error: err.message });
  }
});

// API: Session answers for a specific session (for report generation)
app.get('/api/session-report/:sessionId', async (req, res) => {
  try {
    const { createSupabaseAdmin } = require('./lib/supabase-admin');
    const supabase = createSupabaseAdmin();
    const { sessionId } = req.params;

    // Get session info
    const { data: session, error: sErr } = await supabase
      .from('sesi_ujian')
      .select('*, pengguna(nama_lengkap, nim)')
      .eq('id_sesi', sessionId)
      .single();
    if (sErr) throw sErr;

    // Get session answers
    const { data: answers, error: aErr } = await supabase
      .from('evaluasi_jawaban')
      .select('*, butir_soal(judul_tugas, bobot_nilai)')
      .eq('id_sesi', sessionId)
      .order('id_evaluasi', { ascending: true });
    if (aErr) throw aErr;

    res.json({
      data: {
        id: session.id_sesi,
        exam_type: session.jenis_aplikasi,
        level: session.kategori_ujian,
        total_score: session.skor_diperoleh,
        max_score: session.skor_maksimum,
        status: session.status_kelulusan,
        started_at: session.waktu_mulai,
        finished_at: session.waktu_selesai,
        candidate: session.pengguna ? {
          name: session.pengguna.nama_lengkap,
          nim: session.pengguna.nim
        } : null,
        answers: (answers || []).map((a, i) => {
          // Parse structured catatan_sistem: "TITLE::xxx|DETAIL::yyy"
          let titleFromNote = null;
          if (a.catatan_sistem && a.catatan_sistem.startsWith('TITLE::')) {
            const parts = a.catatan_sistem.split('|DETAIL::');
            titleFromNote = parts[0].replace('TITLE::', '').trim();
          }
          return {
            no: i + 1,
            title: (a.butir_soal && a.butir_soal.judul_tugas) || titleFromNote || `Soal ${i + 1}`,
            score: a.skor_diperoleh,
            max: (a.butir_soal && a.butir_soal.bobot_nilai) || 10,
            detail: a.catatan_sistem
          };
        })
      }
    });
  } catch (err) {
    console.error('[API] session-report error:', err);
    res.status(500).json({ error: err.message });
  }
});

// API: Update Participant Password (menggunakan service key secara aman di server)
app.post('/api/update-participant-password', express.json(), async (req, res) => {
  try {
    const { userId, newPassword } = req.body;
    if (!userId || !newPassword) {
      return res.status(400).json({ error: 'User ID dan password baru wajib diisi.' });
    }
    if (newPassword.length < 6) {
      return res.status(400).json({ error: 'Password minimal harus 6 karakter.' });
    }

    const { createSupabaseAdmin } = require('./lib/supabase-admin');
    const supabase = createSupabaseAdmin();

    const { error } = await supabase.auth.admin.updateUserById(userId, {
      password: newPassword
    });

    if (error) throw error;

    res.json({ success: true, message: 'Password berhasil diperbarui.' });
  } catch (err) {
    console.error('[API] update-participant-password error:', err);
    res.status(500).json({ error: err.message });
  }
});

// ────────────────────────────────────────────────────────

// Redirect path lama → URL baru
for (const [from, to] of Object.entries(LEGACY_REDIRECTS)) {
  app.get(from, (req, res) => {
    const qs = req.url.includes('?') ? req.url.slice(req.url.indexOf('?')) : '';
    res.redirect(301, to + qs);
  });
}

app.use(express.static(publicDir));
app.use(notFoundHandler);
app.use(serverErrorHandler);

module.exports = app;
