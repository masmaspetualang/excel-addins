/**
 * public/js/modules/dashboard/dashboard-api.js
 * ───────────────────────────────────────────
 * Handler komunikasi API khusus modul Dashboard Admin (CMS & Rekap Nilai).
 */
(function () {
  'use strict';

  const Supabase = window.SupabaseClient;

  async function getAllParticipants() {
    console.log('[DashboardAPI] Fetching all participants...');
    const sb = Supabase.getClient();
    if (!sb) return [];
    
    const { data, error } = await sb
      .from('pengguna')
      .select('*')
      .eq('peran', 'participant')
      .order('nama_lengkap', { ascending: true });
      
    if (error) {
      console.error('[DashboardAPI] Error fetching participants:', error);
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
    console.log('[DashboardAPI] Updating participant:', id);
    const sb = Supabase.getClient();
    if (!sb) throw new Error('Klien utama Supabase tidak siap');
    
    const { data, error } = await sb
      .from('pengguna')
      .update({ nama_lengkap: fullName, nim: nim })
      .eq('id_pengguna', id);
      
    if (error) {
      console.error('[DashboardAPI] Error updating participant:', error);
      throw error;
    }
    return true;
  }

  async function deleteParticipant(userId) {
    console.log('[DashboardAPI] Deleting participant cascades for:', userId);
    const sb = Supabase.getClient();
    if (!sb) throw new Error('Klien utama Supabase tidak siap');

    // 1. Get all sessions for this user
    const { data: sessions, error: sErr } = await sb
      .from('sesi_ujian')
      .select('id_sesi')
      .eq('id_pengguna', userId);
      
    if (sErr) {
      console.error('[DashboardAPI] Error fetching sessions for deletion:', sErr);
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
        console.error('[DashboardAPI] Error deleting exam answers:', aErr);
        throw aErr;
      }
    }

    // 3. Delete sesi_ujian for this user
    const { error: usErr } = await sb
      .from('sesi_ujian')
      .delete()
      .eq('id_pengguna', userId);
      
    if (usErr) {
      console.error('[DashboardAPI] Error deleting exam sessions:', usErr);
      throw usErr;
    }

    // 4. Delete pengguna record
    const { error: pErr } = await sb
      .from('pengguna')
      .delete()
      .eq('id_pengguna', userId);
      
    if (pErr) {
      console.error('[DashboardAPI] Error deleting pengguna profile:', pErr);
      throw pErr;
    }

    return true;
  }

  async function getAllResults() {
    const sb = Supabase.getClient();
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
    const sb = Supabase.getClient();
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
    const sb = Supabase.getClient();
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
    const sb = Supabase.getClient();
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

  // Extend window.SupabaseClient
  window.SupabaseClient = window.SupabaseClient || {};
  Object.assign(window.SupabaseClient, {
    getAllParticipants,
    updateParticipant,
    deleteParticipant,
    getAllResults,
    getExamFiles,
    uploadExamFile,
    deleteExamFile
  });
})();
