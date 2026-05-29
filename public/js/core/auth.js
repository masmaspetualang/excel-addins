/**
 * public/js/core/auth.js
 * ───────────────────────
 * Menangani otentikasi pengguna, session management, dan profil pengguna.
 */
(function () {
  'use strict';

  const Supabase = window.SupabaseClient;

  async function getSession() {
    const sb = Supabase.getClient();
    if (!sb) return null;
    const { data: { session } } = await sb.auth.getSession();
    return session;
  }

  async function getCurrentUser() {
    const session = await getSession();
    return session ? session.user : null;
  }

  async function getUserProfile(userId) {
    const sb = Supabase.getClient();
    if (!sb) return null;
    const { data } = await sb.from('pengguna').select('*').eq('id_pengguna', userId).single();
    if (!data) return null;
    return {
      id: data.id_pengguna,
      full_name: data.nama_lengkap,
      nim: data.nim,
      role: data.peran
    };
  }

  async function requireAuth(redirectUrl) {
    redirectUrl = redirectUrl || '/login';
    const session = await getSession();
    if (!session) {
      window.location.href = redirectUrl;
      return null;
    }
    return session;
  }

  async function signIn(email, password) {
    console.log('[Auth] Attempting signIn for:', email);
    const sb = Supabase.getClient();
    if (!sb) throw new Error('Supabase belum dikonfigurasi');
    const { data, error } = await sb.auth.signInWithPassword({ email, password });
    if (error) {
      console.error('[Auth] signIn Error:', error);
      throw error;
    }

    // Sync profile on login if it's missing
    if (data.user) {
      const { data: profile } = await sb.from('pengguna').select('id_pengguna').eq('id_pengguna', data.user.id).single();
      if (!profile) {
        console.log('[Auth] Profile missing, syncing from metadata...');
        const meta = data.user.user_metadata || {};
        await sb.from('pengguna').upsert({
          id_pengguna: data.user.id,
          nama_lengkap: meta.full_name || 'User',
          nim: meta.nim || '—',
          peran: 'participant'
        });
      }
    }

    console.log('[Auth] signIn Success:', data.user.id);
    return data;
  }

  async function signUp(email, password, fullName, nim, role) {
    console.log('[Auth] Attempting signUp for:', email);
    const sb = Supabase.getClient();
    if (!sb) throw new Error('Supabase belum dikonfigurasi');
    const { data, error } = await sb.auth.signUp({
      email, password,
      options: { data: { full_name: fullName, nim: nim, role: role || 'participant' } }
    });
    if (error) {
      console.error('[Auth] signUp Error:', error);
      throw error;
    }
    // Insert profile
    if (data.user) {
      console.log('[Auth] Auth signUp success, inserting profile...');
      const { error: pError } = await sb.from('pengguna').upsert({
        id_pengguna: data.user.id,
        nama_lengkap: fullName,
        nim: nim,
        peran: role || 'participant'
      });
      if (pError) {
        console.error('[Auth] Profile insert error:', pError);
      } else {
        console.log('[Auth] Profile created successfully');
      }
    }
    return data;
  }

  async function registerParticipantByAdmin(email, password, fullName, nim) {
    console.log('[Auth] Admin registering participant:', email);
    const cfg = window.APP_CONFIG || {};
    if (!window.supabase || !cfg.SUPABASE_URL) throw new Error('Supabase belum dikonfigurasi');

    // Create a temporary, non-persisted client to avoid replacing the active Admin session
    const tempSb = window.supabase.createClient(cfg.SUPABASE_URL, cfg.SUPABASE_ANON_KEY, {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
        detectSessionInUrl: false
      }
    });

    const { data, error } = await tempSb.auth.signUp({
      email,
      password,
      options: { data: { full_name: fullName, nim: nim, role: 'participant' } }
    });

    if (error) {
      console.error('[Auth] Admin register error:', error);
      throw error;
    }

    if (data.user) {
      console.log('[Auth] Auth signUp success, inserting profile using admin session...');
      const mainSb = Supabase.getClient();
      if (!mainSb) throw new Error('Klien utama Supabase tidak siap');
      
      const { error: pError } = await mainSb.from('pengguna').upsert({
        id_pengguna: data.user.id,
        nama_lengkap: fullName,
        nim: nim,
        peran: 'participant'
      });
      
      if (pError) {
        console.error('[Auth] Profile insert error:', pError);
        throw pError;
      }
      console.log('[Auth] Profile created successfully by admin');
    }
    return data;
  }

  async function signOut() {
    const sb = Supabase.getClient();
    if (sb) await sb.auth.signOut();
  }

  // Extend window.SupabaseClient
  window.SupabaseClient = window.SupabaseClient || {};
  Object.assign(window.SupabaseClient, {
    getSession,
    getCurrentUser,
    getUserProfile,
    requireAuth,
    signIn,
    signUp,
    registerParticipantByAdmin,
    signOut
  });
})();
