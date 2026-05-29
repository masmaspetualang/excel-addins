/**
 * Seed butir_soal dari public/js/modules/exam/exams.json (satu sumber data soal).
 * Run: node scripts/seed-questions.js
 */
const fs = require('fs');
const path = require('path');
const { createSupabaseAdmin } = require('../server/lib/supabase-admin');

const EXAMS_PATH = path.join(__dirname, '../public/js/modules/exam/exams.json');

const APP_TYPES = [
  { key: 'EXAMS', jenis: 'excel' },
  { key: 'WORD_EXAMS', jenis: 'word' },
  { key: 'POWERPOINT_EXAMS', jenis: 'ppt' },
];

function taskToRow(task, nomorUrut, jenisAplikasi, kategoriUjian) {
  return {
    nomor_urut: nomorUrut,
    bobot_nilai: task.points,
    metode_penilaian: 'api',
    judul_tugas: task.title,
    instruksi_tugas: task.desc,
    langkah_verifikasi: task.steps,
    petunjuk_bantuan: task.hint || '',
    jenis_aplikasi: jenisAplikasi,
    kategori_ujian: kategoriUjian,
  };
}

function loadQuestionsFromJson() {
  const raw = fs.readFileSync(EXAMS_PATH, 'utf8');
  const data = JSON.parse(raw);
  const rows = [];

  for (const { key, jenis } of APP_TYPES) {
    const examMap = data[key] || {};
    for (const kategori of Object.keys(examMap)) {
      const exam = examMap[kategori];
      if (!exam?.tasks?.length) continue;
      exam.tasks.forEach((task, i) => {
        rows.push(taskToRow(task, i + 1, jenis, kategori));
      });
    }
  }

  return rows;
}

async function seed() {
  const supabase = createSupabaseAdmin();
  const rows = loadQuestionsFromJson();

  console.log(`🌱 Seeding ${rows.length} questions from exams.json...\n`);

  const { error: delErr } = await supabase.from('butir_soal').delete().neq('id_soal', 0);
  if (delErr) console.warn('Delete warning:', delErr.message);

  let total = 0;
  for (const row of rows) {
    const { error } = await supabase.from('butir_soal').insert(row);
    if (error) {
      console.error(`[${row.jenis_aplikasi}] Q${row.nomor_urut}:`, error.message);
    } else {
      process.stdout.write(row.jenis_aplikasi === 'excel' ? '📊' : row.jenis_aplikasi === 'word' ? '📝' : '📽');
      total++;
    }
  }

  console.log(`\n\n✅ Done! ${total} questions seeded into butir_soal.\n`);
}

seed().catch((err) => {
  console.error(err.message || err);
  process.exit(1);
});
