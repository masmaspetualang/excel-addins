/**
 * Memuat soal ujian dari exams.json dan mengikat fungsi checker dari OfficeCheckers.
 */
(function () {
  'use strict';

  const EXAMS_JSON_URL = '/js/modules/exam/exams.json';

  function bindCheckerFunctions(examMap) {
    if (!examMap || !window.OfficeCheckers) return;
    for (const examKey of Object.keys(examMap)) {
      const exam = examMap[examKey];
      if (!exam?.tasks) continue;
      exam.tasks.forEach((task) => {
        if (typeof task.check === 'string') {
          const fn = window.OfficeCheckers[task.check];
          if (fn) task.check = fn;
        }
      });
    }
  }

  async function loadExamsData() {
    const res = await fetch(EXAMS_JSON_URL, { cache: 'no-store' });
    if (!res.ok) {
      throw new Error(`Gagal memuat exams.json (${res.status})`);
    }
    const data = await res.json();

    window.EXAMS = data.EXAMS || {};
    window.WORD_EXAMS = data.WORD_EXAMS || {};
    window.POWERPOINT_EXAMS = data.POWERPOINT_EXAMS || {};

    bindCheckerFunctions(window.EXAMS);
    bindCheckerFunctions(window.WORD_EXAMS);
    bindCheckerFunctions(window.POWERPOINT_EXAMS);

    return data;
  }

  window.ExamsLoader = { loadExamsData };
})();
