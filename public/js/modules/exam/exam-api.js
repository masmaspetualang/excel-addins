/**
 * public/js/modules/exam/exam-api.js
 * ─────────────────────────────────
 * Handler komunikasi API khusus modul Ujian/Exam.
 */
(function () {
  'use strict';

  const Supabase = window.SupabaseClient;

  async function loadQuestions(examType, level) {
    const sb = Supabase.getClient();
    if (!sb) return null;
    const { data, error } = await sb
      .from('butir_soal')
      .select('*')
      .eq('jenis_aplikasi', examType)
      .eq('tingkat_kesulitan', level)
      .order('urutan_soal', { ascending: true });
    if (error) {
      console.error('Load questions error:', error);
      return null;
    }
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
    const sb = Supabase.getClient();
    if (!sb) return null;
    const { data, error } = await sb.from('sesi_ujian').insert({
      id_pengguna: userId,
      jenis_aplikasi: examType,
      kategori_ujian: level,
      skor_maksimum: maxScore,
      status_kelulusan: 'in_progress'
    }).select().single();
    if (error) {
      console.error('Create session error:', error);
      return null;
    }
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
    const sb = Supabase.getClient();
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

  // Extend window.SupabaseClient
  window.SupabaseClient = window.SupabaseClient || {};
  Object.assign(window.SupabaseClient, {
    loadQuestions,
    createExamSession,
    saveExamResults
  });
})();
