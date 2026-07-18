/**
 * public/js/core/supabase.js
 * ──────────────────────────
 * Menginisialisasi Klien Supabase tunggal (Singleton).
 */
(function () {
  'use strict';

  const cfg = window.APP_CONFIG || {};
  const SUPABASE_URL = cfg.SUPABASE_URL || '';
  const SUPABASE_ANON_KEY = cfg.SUPABASE_ANON_KEY || '';

  let _client = null;

  function getClient() {
    if (_client) return _client;
    if (!window.supabase || !SUPABASE_URL || SUPABASE_URL.includes('YOUR_PROJECT')) {
      console.warn('[ExamQuiz] Supabase belum dikonfigurasi. Isi SUPABASE_URL dan SUPABASE_ANON_KEY di .env');
      return null;
    }
    _client = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
      auth: { persistSession: true, storageKey: 'examquiz-auth' }
    });
    return _client;
  }

  // Daftarkan ke global namespace
  window.SupabaseClient = window.SupabaseClient || {};
  window.SupabaseClient.getClient = getClient;
})();
