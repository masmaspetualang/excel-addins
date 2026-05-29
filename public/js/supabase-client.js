/**
 * ExcelQuiz Pro — Supabase Client & Auth Helpers
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
      console.warn('[ExcelQuiz] Supabase belum dikonfigurasi. Edit public/config.js');
      return null;
    }
    _client = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
      auth: { persistSession: true, storageKey: 'excelquiz-auth' }
    });
    return _client;
  }

  async function getSession() {
    const sb = getClient();
    if (!sb) return null;
    const { data: { session } } = await sb.auth.getSession();
    return session;
  }

  async function getCurrentUser() {
    const session = await getSession();
    return session ? session.user : null;
  }

  async function getUserProfile(userId) {
    const sb = getClient();
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
    redirectUrl = redirectUrl || '/login.html';
    const session = await getSession();
    if (!session) {
      window.location.href = redirectUrl;
      return null;
    }
    return session;
  }

  async function signIn(email, password) {
    console.log('[SupabaseClient] Attempting signIn for:', email);
    const sb = getClient();
    if (!sb) throw new Error('Supabase belum dikonfigurasi');
    const { data, error } = await sb.auth.signInWithPassword({ email, password });
    if (error) {
      console.error('[SupabaseClient] signIn Error:', error);
      throw error;
    }

    // Sync profile on login if it's missing
    if (data.user) {
      const { data: profile } = await sb.from('pengguna').select('id_pengguna').eq('id_pengguna', data.user.id).single();
      if (!profile) {
        console.log('[SupabaseClient] Profile missing, syncing from metadata...');
        const meta = data.user.user_metadata || {};
        await sb.from('pengguna').upsert({
          id_pengguna: data.user.id,
          nama_lengkap: meta.full_name || 'User',
          nim: meta.nim || '—',
          peran: 'participant'
        });
      }
    }

    console.log('[SupabaseClient] signIn Success:', data.user.id);
    return data;
  }

  async function signUp(email, password, fullName, nim, role) {
    console.log('[SupabaseClient] Attempting signUp for:', email);
    const sb = getClient();
    if (!sb) throw new Error('Supabase belum dikonfigurasi');
    const { data, error } = await sb.auth.signUp({
      email, password,
      options: { data: { full_name: fullName, nim: nim, role: role || 'participant' } }
    });
    if (error) {
      console.error('[SupabaseClient] signUp Error:', error);
      throw error;
    }
    // Insert profile
    if (data.user) {
      console.log('[SupabaseClient] Auth signUp success, inserting profile...');
      const { error: pError } = await sb.from('pengguna').upsert({
        id_pengguna: data.user.id,
        nama_lengkap: fullName,
        nim: nim,
        peran: role || 'participant'
      });
      if (pError) {
        console.error('[SupabaseClient] Profile insert error:', pError);
        // We don't throw here to let the user know they registered but profile failed
      } else {
        console.log('[SupabaseClient] Profile created successfully');
      }
    }
    return data;
  }

  async function registerParticipantByAdmin(email, password, fullName, nim) {
    console.log('[SupabaseClient] Admin registering participant:', email);
    if (!window.supabase || !SUPABASE_URL) throw new Error('Supabase belum dikonfigurasi');

    // Create a temporary, non-persisted client to avoid replacing the active Admin session
    const tempSb = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
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
      console.error('[SupabaseClient] Admin register error:', error);
      throw error;
    }

    if (data.user) {
      console.log('[SupabaseClient] Auth signUp success, inserting profile using admin session...');
      const mainSb = getClient();
      if (!mainSb) throw new Error('Klien utama Supabase tidak siap');
      
      const { error: pError } = await mainSb.from('pengguna').upsert({
        id_pengguna: data.user.id,
        nama_lengkap: fullName,
        nim: nim,
        peran: 'participant'
      });
      
      if (pError) {
        console.error('[SupabaseClient] Profile insert error:', pError);
        throw pError;
      }
      console.log('[SupabaseClient] Profile created successfully by admin');
    }
    return data;
  }

  async function getAllParticipants() {
    console.log('[SupabaseClient] Fetching all participants...');
    const sb = getClient();
    if (!sb) return [];
    
    const { data, error } = await sb
      .from('pengguna')
      .select('*')
      .eq('peran', 'participant')
      .order('nama_lengkap', { ascending: true });
      
    if (error) {
      console.error('[SupabaseClient] Error fetching participants:', error);
      throw error;
    }
    return (data || []).map(p => ({
      id: p.id_pengguna,
      full_name: p.nama_lengkap,
      nim: p.nim,
      role: p.peran
    }));
  }

  async function updateParticipant(id, fullName, nim) {
    console.log('[SupabaseClient] Updating participant:', id);
    const sb = getClient();
    if (!sb) throw new Error('Klien utama Supabase tidak siap');
    
    const { data, error } = await sb
      .from('pengguna')
      .update({ nama_lengkap: fullName, nim: nim })
      .eq('id_pengguna', id);
      
    if (error) {
      console.error('[SupabaseClient] Error updating participant:', error);
      throw error;
    }
    return true;
  }

  async function deleteParticipant(userId) {
    console.log('[SupabaseClient] Deleting participant cascades for:', userId);
    const sb = getClient();
    if (!sb) throw new Error('Klien utama Supabase tidak siap');

    // 1. Get all sessions for this user
    const { data: sessions, error: sErr } = await sb
      .from('sesi_ujian')
      .select('id_sesi')
      .eq('id_pengguna', userId);
      
    if (sErr) {
      console.error('[SupabaseClient] Error fetching sessions for deletion:', sErr);
      throw sErr;
    }
    
    const sessionIds = (sessions || []).map(s => s.id_sesi);

    // 2. Delete evaluasi_jawaban for these sessions
    if (sessionIds.length > 0) {
      const { error: aErr } = await sb
        .from('evaluasi_jawaban')
        .delete()
        .in('id_sesi', sessionIds);
        
      if (aErr) {
        console.error('[SupabaseClient] Error deleting exam answers:', aErr);
        throw aErr;
      }
    }

    // 3. Delete sesi_ujian for this user
    const { error: usErr } = await sb
      .from('sesi_ujian')
      .delete()
      .eq('id_pengguna', userId);
      
    if (usErr) {
      console.error('[SupabaseClient] Error deleting exam sessions:', usErr);
      throw usErr;
    }

    // 4. Delete pengguna record
    const { error: pErr } = await sb
      .from('pengguna')
      .delete()
      .eq('id_pengguna', userId);
      
    if (pErr) {
      console.error('[SupabaseClient] Error deleting pengguna profile:', pErr);
      throw pErr;
    }

    return true;
  }

  async function signOut() {
    const sb = getClient();
    if (sb) await sb.auth.signOut();
  }

  async function loadQuestions(examType, level) {
    const sb = getClient();
    if (!sb) return null;
    const { data, error } = await sb
      .from('butir_soal')
      .select('*')
      .eq('jenis_aplikasi', examType)
      .eq('tingkat_kesulitan', level)
      .order('urutan_soal', { ascending: true });
    if (error) { console.error('Load questions error:', error); return null; }
    return (data || []).map(q => ({
      id: q.id_soal,
      exam_type: q.jenis_aplikasi,
      level: q.tingkat_kesulitan,
      question_order: q.urutan_soal,
      title: q.judul_soal,
      description: q.deskripsi_soal,
      points: q.poin,
      expected: q.jawaban_diharapkan,
      checker_type: q.tipe_pemeriksaan,
      params: q.parameter_pemeriksaan,
      hint: q.petunjuk
    }));
  }

  async function createExamSession(userId, examType, level, maxScore) {
    const sb = getClient();
    if (!sb) return null;
    const { data, error } = await sb.from('sesi_ujian').insert({
      id_pengguna: userId,
      jenis_aplikasi: examType,
      kategori_ujian: level,
      skor_maksimum: maxScore,
      status_kelulusan: 'in_progress'
    }).select().single();
    if (error) { console.error('Create session error:', error); return null; }
    return {
      id: data.id_sesi,
      user_id: data.id_pengguna,
      exam_type: data.jenis_aplikasi,
      level: data.kategori_ujian,
      max_score: data.skor_maksimum,
      status: data.status_kelulusan,
      started_at: data.waktu_mulai
    };
  }

  async function saveExamResults(sessionId, totalScore, maxScore, answers) {
    const sb = getClient();
    if (!sb) return false;

    try {
      // 1. Dapatkan jenis_aplikasi dari sesi_ujian
      const { data: session, error: sErr } = await sb
        .from('sesi_ujian')
        .select('jenis_aplikasi, kategori_ujian')
        .eq('id_sesi', sessionId)
        .single();

      if (sErr || !session) {
        console.error('Failed to get exam session:', sErr);
        return false;
      }

      // 2. Update sesi_ujian
      const { error: uErr } = await sb.from('sesi_ujian').update({
        skor_diperoleh: totalScore,
        skor_maksimum: maxScore,
        waktu_selesai: new Date().toISOString(),
        status_kelulusan: totalScore / maxScore >= 0.7 ? 'lulus' : 'tidak_lulus'
      }).eq('id_sesi', sessionId);

      if (uErr) {
        console.error('Failed to update exam session:', uErr);
        return false;
      }

      // 3. Ambil butir_soal terkait dari DB untuk mendapatkan id_soal yang valid
      const { data: dbQuestions, error: qErr } = await sb
        .from('butir_soal')
        .select('id_soal, nomor_urut')
        .eq('jenis_aplikasi', session.jenis_aplikasi)
        .eq('kategori_ujian', session.kategori_ujian)
        .order('nomor_urut', { ascending: true });

      if (qErr || !dbQuestions) {
        console.error('Failed to load questions for mapping:', qErr);
        return false;
      }

      // 4. Masukkan jawaban per soal (evaluasi_jawaban)
      if (answers && answers.length > 0) {
        const insertPayload = answers.map((a, index) => {
          // Cari id_soal yang sesuai berdasarkan indeks urutan
          const dbQ = dbQuestions[index];
          const dbQuestionId = dbQ ? dbQ.id_soal : null;

          return {
            id_sesi: sessionId,
            id_soal: dbQuestionId,
            skor_diperoleh: a.score,
            catatan_sistem: a.detail || 'Pemeriksaan otomatis selesai',
            status_jawaban: a.score > 0 ? 'benar' : 'salah'
          };
        }).filter(item => item.id_soal !== null); // Hanya masukkan jika id_soal ditemukan

        if (insertPayload.length > 0) {
          const { error: insErr } = await sb.from('evaluasi_jawaban').insert(insertPayload);
          if (insErr) {
            console.error('Failed to insert answer breakdown:', insErr);
            return false;
          }
        }
      }
      return true;
    } catch (err) {
      console.error('saveExamResults unexpected error:', err);
      return false;
    }
  }

  async function getAllResults() {
    const sb = getClient();
    if (!sb) return [];
    const { data } = await sb
      .from('sesi_ujian')
      .select(`*, pengguna(nama_lengkap, nim)`)
      .order('waktu_mulai', { ascending: false });
    
    // Map to original English format for frontend dashboard
    return (data || []).map(row => ({
      id: row.id_sesi,
      user_id: row.id_pengguna,
      exam_type: row.jenis_aplikasi,
      level: row.kategori_ujian,
      total_score: row.skor_diperoleh,
      max_score: row.skor_maksimum,
      status: row.status_kelulusan,
      started_at: row.waktu_mulai,
      finished_at: row.waktu_selesai,
      profiles: row.pengguna ? {
        full_name: row.pengguna.nama_lengkap,
        nim: row.pengguna.nim
      } : null
    }));
  }

  // ─── EXAM FILE MANAGEMENT (CMS) ─────────────────
  async function getExamFiles() {
    const sb = getClient();
    if (!sb) return [];
    const { data, error } = await sb.from('berkas_template').select('*').order('jenis_aplikasi');
    if (error) throw error;
    return (data || []).map(f => ({
      id: f.id_berkas,
      exam_type: f.jenis_aplikasi,
      display_name: f.nama_tampilan,
      file_path: f.tautan_berkas,
      is_available: f.status_aktif,
      updated_at: f.waktu_pembaruan
    }));
  }

  async function uploadExamFile(examType, file) {
    const sb = getClient();
    if (!sb) throw new Error('Supabase client not ready');

    const fileExt = file.name.split('.').pop();
    const fileName = `${examType}_soal_${Date.now()}.${fileExt}`;
    const filePath = `uploads/${fileName}`;

    // 1. Upload ke Storage (Bucket: soal-ujian)
    const { data: uploadData, error: uploadError } = await sb.storage
      .from('soal-ujian')
      .upload(filePath, file, { upsert: true });

    if (uploadError) throw uploadError;

    // 2. Dapatkan Public URL
    const { data: urlData } = sb.storage.from('soal-ujian').getPublicUrl(filePath);

    // 3. Update Tabel berkas_template
    const { error: dbError } = await sb
      .from('berkas_template')
      .update({
        tautan_berkas: urlData.publicUrl,
        status_aktif: true,
        waktu_pembaruan: new Date().toISOString()
      })
      .eq('jenis_aplikasi', examType);

    if (dbError) throw dbError;
    return urlData.publicUrl;
  }

  async function deleteExamFile(examType, filePath) {
    const sb = getClient();
    if (!sb) throw new Error('Supabase client not ready');

    // 1. Jika ada path, hapus dari Storage
    if (filePath) {
      // Ambil nama file dari URL atau path
      const pathParts = filePath.split('/soal-ujian/');
      if (pathParts.length > 1) {
        const internalPath = pathParts[1];
        await sb.storage.from('soal-ujian').remove([internalPath]);
      }
    }

    // 2. Update Tabel berkas_template menjadi tidak tersedia
    const { error } = await sb
      .from('berkas_template')
      .update({
        tautan_berkas: null,
        status_aktif: false,
        waktu_pembaruan: new Date().toISOString()
      })
      .eq('jenis_aplikasi', examType);

    if (error) throw error;
    return true;
  }

  // Expose globally
  window.SupabaseClient = {
    getClient,
    getSession,
    getCurrentUser,
    getUserProfile,
    requireAuth,
    signIn,
    signUp,
    registerParticipantByAdmin,
    getAllParticipants,
    updateParticipant,
    deleteParticipant,
    signOut,
    loadQuestions,
    createExamSession,
    saveExamResults,
    getAllResults,
    getExamFiles,
    uploadExamFile,
    deleteExamFile
  };
})();
