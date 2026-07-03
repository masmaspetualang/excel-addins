/**
 * report-generator.js
 * Generates the UMY-style Examination Score Report HTML.
 * Used by both participant result screen and admin dashboard.
 */

window.ReportGenerator = (function () {

  const PASSING_SCORE = 70;
  const LOGO_PATH = '/assets/logo/LogoUMY.svg';

  /**
   * Formats date to "DD Bulan YYYY" Indonesian format.
   */
  function formatDate(dateStr) {
    if (!dateStr) return '—';
    const d = new Date(dateStr);
    const months = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
  }

  /**
   * Formats date to "Bulan YYYY" Indonesian format (for signature area).
   */
  function formatMonthYear(dateStr) {
    if (!dateStr) return new Date().getFullYear();
    const d = new Date(dateStr);
    const months = [
      'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
      'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
    ];
    return `${months[d.getMonth()]} ${d.getFullYear()}`;
  }

  /**
   * Calculates percentage from score / max score
   */
  function pct(score, max) {
    if (!max) return '0%';
    return Math.round((score / max) * 100) + '%';
  }

  /**
   * Returns a label for exam type
   */
  function examTypeLabel(type) {
    const map = { excel: 'Microsoft Excel', word: 'Microsoft Word', ppt: 'Microsoft PowerPoint' };
    return map[(type || '').toLowerCase()] || type || 'Komputer';
  }

  /**
   * Returns a level label
   */
  function levelLabel(lvl) {
    const map = { basic: 'Dasar', intermediate: 'Menengah', advanced: 'Mahir' };
    return map[(lvl || '').toLowerCase()] || lvl || '—';
  }

  // ─────────────────────────────────────────────────────────────
  //  5 KATEGORI KOMPETENSI — mapping per jenis ujian
  //  Format: setiap kategori berisi array keyword judul soal yang
  //  akan dicocokkan (case-insensitive, partial match) ke answers.
  // ─────────────────────────────────────────────────────────────
  const CATEGORY_MAP = {
    excel: [
      {
        name: 'Pemformatan Teks & Gaya Sel',
        titles: ['Format Adventure Works', 'Bersihkan Background']
      },
      {
        name: 'Tata Letak & Penyelarasan',
        titles: ['Alignment Header', 'Border Tabel']
      },
      {
        name: 'Pengaturan Kolom & Format Data',
        titles: ['Tambah Kolom No', 'Format Price']
      },
      {
        name: 'Formula & Analisis Perhitungan',
        titles: ['Total Penjualan']
      },
      {
        name: 'Manajemen Sheet & Input Data',
        titles: ['Rename Sheets', 'Format Waktu', 'Konfirmasi Selesai']
      }
    ],
    word: [
      {
        name: 'Pemformatan Teks & Gaya Font',
        titles: ['Format Judul Utama', 'Format Kutipan Teks', 'Format Subjudul']
      },
      {
        name: 'Tata Letak Paragraf & Halaman',
        titles: ['Kerapian Paragraf', 'Penyisipan Footer']
      },
      {
        name: 'Penyuntingan & Efek Tulisan',
        titles: ['Koreksi Kesalahan Ejaan', 'Format Tulisan Ilmiah']
      },
      {
        name: 'Penyisipan Tabel & Tata Dokumen',
        titles: ['Pembuatan Tabel Data']
      },
      {
        name: 'Tautan & Objek Interaktif',
        titles: ['Pemberian Highlight', 'Penerapan Hyperlink']
      }
    ],
    ppt: [
      {
        name: 'Pemformatan Judul & Teks Utama',
        titles: ['Format Judul Slide Utama', 'Format Judul Slide RAM']
      },
      {
        name: 'Manajemen & Tata Urutan Slide',
        titles: ['Pembuatan Slide Baru', 'Organisasi Urutan Slide']
      },
      {
        name: 'Penyisipan & Format Gambar',
        titles: ['Penyisipan Gambar Ilustrasi', 'Skala & Posisi Gambar']
      },
      {
        name: 'Penataan Teks & Penekanan Konten',
        titles: ['Format Penekanan Konten', 'Format Penekanan Teks Pipelining', 'Format Huruf Miring']
      },
      {
        name: 'Pembuatan Tabel & Objek Data',
        titles: ['Pembuatan Tabel Data']
      }
    ]
  };

  /**
   * Match an answer title to category keywords (case-insensitive, partial).
   */
  function matchesTitle(answerTitle, keywords) {
    const t = (answerTitle || '').toLowerCase();
    return keywords.some(k => t.includes(k.toLowerCase()));
  }

  /**
   * Group answers array into 5 competency categories for the given exam type.
   * Returns array of { name, score, max, items[] }.
   */
  function groupByCategory(answers, examType) {
    const typeKey = (examType || '').toLowerCase().replace('powerpoint', 'ppt');
    const categories = CATEGORY_MAP[typeKey] || CATEGORY_MAP['excel'];

    // Build category buckets
    const buckets = categories.map(cat => ({
      name: cat.name,
      keywords: cat.titles,
      score: 0,
      max: 0,
      items: []
    }));

    const unmatched = [];

    answers.forEach(a => {
      let matched = false;
      for (const bucket of buckets) {
        if (matchesTitle(a.title, bucket.keywords)) {
          bucket.score += (a.score || 0);
          bucket.max   += (a.max   || 0);
          bucket.items.push(a);
          matched = true;
          break;
        }
      }
      if (!matched) {
        unmatched.push(a);
      }
    });

    // Distribute unmatched answers across buckets by index
    unmatched.forEach((a, idx) => {
      const target = buckets[idx % buckets.length];
      target.score += (a.score || 0);
      target.max   += (a.max   || 0);
      target.items.push(a);
    });

    // Fallback max for empty buckets (10 pts per expected soal)
    buckets.forEach((b, i) => {
      if (b.max === 0) {
        b.max = (categories[i]?.titles.length || 1) * 10;
      }
    });

    return buckets;
  }

  /**
   * Build the full report HTML string.
   * @param {Object} opts
   * @param {string}  opts.candidateName
   * @param {string}  opts.candidateNim
   * @param {number}  opts.totalScore
   * @param {number}  opts.maxScore  (default 100)
   * @param {string}  opts.examType  (excel|word|ppt)
   * @param {string}  opts.level     (basic|intermediate|advanced)
   * @param {string}  opts.date      ISO date string
   * @param {Array}   opts.answers   [{no, title, score, max}]
   * @param {boolean} opts.forPrint  adds print-specific wrapping
   */
  function buildHTML(opts) {
    const candidateName = opts.candidateName || (opts.candidate ? opts.candidate.name : '—');
    const candidateNim = opts.candidateNim || (opts.candidate ? opts.candidate.nim : '—');
    const totalScore = opts.totalScore !== undefined ? opts.totalScore : (opts.total_score !== undefined ? opts.total_score : 0);
    const maxScore = opts.maxScore !== undefined ? opts.maxScore : (opts.max_score !== undefined ? opts.max_score : 100);
    const examType = opts.examType || opts.exam_type || 'excel';
    const level = opts.level || 'basic';
    const date = opts.date || opts.started_at;
    const answers = opts.answers || [];

    const passed = totalScore >= PASSING_SCORE;
    const statusLabel = passed ? 'LULUS' : 'TIDAK LULUS';
    const statusColor = passed ? '#16a34a' : '#dc2626';
    const totalPct = Math.round((totalScore / maxScore) * 100);

    // ── Group answers into 5 competency categories ──
    const categories = groupByCategory(answers, examType);

    const rowsHTML = categories.map((cat, i) => {
      const catPct = cat.max > 0 ? Math.round((cat.score / cat.max) * 100) : 0;
      const isKompeten = catPct >= PASSING_SCORE;

      // Sub-items bullet list
      const subItems = cat.items.map(a =>
        `<div style="font-size:9.5px;color:#666;padding-left:8px;margin-top:1px;line-height:1.2;">
          • ${a.title || '—'}&nbsp;<span style="font-weight:600;color:${(a.score||0)>0?'#15803d':'#b91c1c'};">(${a.score||0}/${a.max||0} poin)</span>
        </div>`
      ).join('');

      return `
        <tr style="background:${i % 2 === 0 ? '#fff' : '#f9fafb'};">
          <td style="padding:5px 6px;text-align:center;font-weight:700;color:#555;vertical-align:top;">${i + 1}</td>
          <td style="padding:5px 6px;vertical-align:top;">
            <div style="font-weight:700;color:#1a1a1a;font-size:11.5px;">${cat.name}</div>
            ${subItems}
          </td>
          <td style="padding:5px 6px;text-align:center;font-weight:700;font-size:12px;vertical-align:top;color:${isKompeten ? '#15803d' : '#dc2626'};">${cat.score}</td>
          <td style="padding:5px 6px;text-align:center;font-size:12px;vertical-align:top;color:#555;">${cat.max}</td>
          <td style="padding:5px 6px;text-align:center;font-weight:700;font-size:12px;vertical-align:top;color:${isKompeten ? '#15803d' : '#dc2626'};">${catPct}%</td>
          <td style="padding:5px 6px;text-align:center;vertical-align:top;">
            <span style="
              display:inline-block;padding:2px 8px;border-radius:999px;font-size:10px;font-weight:700;
              background:${isKompeten ? '#dcfce7' : '#fee2e2'};
              color:${isKompeten ? '#15803d' : '#b91c1c'};
              border:1px solid ${isKompeten ? '#86efac' : '#fca5a5'};
            ">${isKompeten ? '✓ Kompeten' : '✗ Belum Kompeten'}</span>
          </td>
        </tr>`;
    }).join('');

    return `
      <div class="report-wrap" style="
        font-family: 'Times New Roman', Times, serif;
        max-width: 800px;
        margin: 0 auto;
        background: #fff;
        padding: 10px 20px;
        color: #1a1a1a;
      ">
        <!-- HEADER -->
        <div style="display:flex;align-items:center;border-bottom:3px double #c8102e;padding-bottom:8px;margin-bottom:10px;gap:20px;">
          <img src="${LOGO_PATH}" alt="Logo UMY" style="height:60px;width:auto;flex-shrink:0;" onerror="this.style.display='none'" />
          <div style="text-align:center;flex:1;">
            <div style="font-size:13px;font-weight:700;letter-spacing:0.05em;text-transform:uppercase;color:#1a1a1a;">
              Universitas Muhammadiyah Yogyakarta
            </div>
            <div style="font-size:11px;color:#444;margin-top:2px;">
              Jl. Brawijaya, Tamantirto, Kasihan, Bantul, Yogyakarta 55183
            </div>
            <div style="margin-top:4px;font-size:14px;font-weight:800;color:#c8102e;letter-spacing:0.03em;text-transform:uppercase;">
              Laporan Hasil Ujian Kompetensi Komputer
            </div>
            <div style="font-size:11px;color:#555;margin-top:2px;">ExamQuiz — ${examTypeLabel(examType)} • Level ${levelLabel(level)}</div>
          </div>
        </div>

        <!-- CANDIDATE INFO -->
        <table style="width:100%;border-collapse:collapse;margin-bottom:10px;font-size:12.5px;">
          <tr>
            <td style="width:160px;padding:2px 0;color:#555;">Nama Peserta</td>
            <td style="padding:2px 0;">: <strong>${candidateName}</strong></td>
            <td style="width:160px;padding:2px 0;color:#555;">Tanggal Ujian</td>
            <td style="padding:2px 0;">: ${formatDate(date)}</td>
          </tr>
          <tr>
            <td style="padding:2px 0;color:#555;">NIM</td>
            <td style="padding:2px 0;">: <strong>${candidateNim}</strong></td>
            <td style="padding:2px 0;color:#555;">Jenis Ujian</td>
            <td style="padding:2px 0;">: ${examTypeLabel(examType)}</td>
          </tr>
          <tr>
            <td style="padding:2px 0;color:#555;">Nilai Kelulusan</td>
            <td style="padding:2px 0;">: <strong>≥ ${PASSING_SCORE}</strong></td>
            <td style="padding:2px 0;color:#555;"></td>
            <td style="padding:2px 0;"></td>
          </tr>
        </table>

        <!-- SCORE SECTION -->
        <div style="
          display:flex;align-items:center;justify-content:space-between;
          background:#f8f9fa;border:1px solid #e0e0e0;border-radius:8px;
          padding:8px 16px;margin-bottom:10px;
        ">
          <div style="text-align:center;">
            <div style="font-size:10px;color:#666;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:2px;">Skor Total</div>
            <div style="font-size:30px;font-weight:900;color:#1a1a1a;line-height:1;">${totalScore}</div>
            <div style="font-size:11px;color:#888;">dari ${maxScore} (${totalPct}%)</div>
          </div>
          <div style="width:1px;height:45px;background:#ddd;"></div>
          <div style="text-align:center;">
            <div style="font-size:10px;color:#666;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:2px;">Status</div>
            <div style="
              font-size:18px;font-weight:900;letter-spacing:0.08em;
              color:${statusColor};
              border:2px solid ${statusColor};
              border-radius:6px;padding:2px 14px;
            ">${statusLabel}</div>
          </div>
          <div style="width:1px;height:45px;background:#ddd;"></div>
          <div style="text-align:center;">
            <div style="font-size:10px;color:#666;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:2px;">Nilai Minimum Lulus</div>
            <div style="font-size:24px;font-weight:900;color:#c8102e;line-height:1;">${PASSING_SCORE}</div>
            <div style="font-size:11px;color:#888;">poin</div>
          </div>
        </div>

        <!-- BREAKDOWN TABLE — 5 Kategori Kompetensi -->
        <div style="margin-bottom:10px;">
          <div style="font-size:11.5px;font-weight:700;text-transform:uppercase;letter-spacing:0.06em;color:#555;margin-bottom:6px;border-bottom:1px solid #ddd;padding-bottom:4px;">
            Rincian Analisis Per Kategori Kompetensi
          </div>
          <table style="width:100%;border-collapse:collapse;font-size:11.5px;border:1px solid #e0e0e0;">
            <thead>
              <tr style="background:#1a1a1a;color:#fff;">
                <th style="padding:5px 6px;text-align:center;width:36px;">No</th>
                <th style="padding:5px 6px;text-align:left;">Kategori Kompetensi</th>
                <th style="padding:5px 6px;text-align:center;width:60px;">Skor</th>
                <th style="padding:5px 6px;text-align:center;width:60px;">Maks</th>
                <th style="padding:5px 6px;text-align:center;width:70px;">% Benar</th>
                <th style="padding:5px 6px;text-align:center;width:130px;">Keterangan</th>
              </tr>
            </thead>
            <tbody>
              ${rowsHTML || `<tr><td colspan="6" style="text-align:center;padding:12px;color:#888;">Tidak ada rincian jawaban</td></tr>`}
            </tbody>
            <tfoot>
              <tr style="background:#1a1a1a;color:#fff;font-weight:700;">
                <td colspan="2" style="padding:6px 6px;text-align:right;letter-spacing:0.04em;">TOTAL</td>
                <td style="padding:6px 6px;text-align:center;font-size:13px;">${totalScore}</td>
                <td style="padding:6px 6px;text-align:center;font-size:13px;">${maxScore}</td>
                <td style="padding:6px 6px;text-align:center;font-size:13px;">${totalPct}%</td>
                <td style="padding:6px 6px;text-align:center;color:${statusColor};font-size:12.5px;font-weight:900;">${statusLabel}</td>
              </tr>
            </tfoot>
          </table>
          <div style="margin-top:6px;font-size:9.5px;color:#888;font-style:italic;">
            * Kompetensi dinyatakan tercapai apabila nilai per kategori ≥ ${PASSING_SCORE}% dari skor maksimum kategori tersebut.
          </div>
        </div>

        <!-- SIGNATURE BLOCK -->
        <div style="margin-top:12px;border-top:1px solid #ddd;padding-top:8px;display:flex;justify-content:flex-end;">
          <div style="text-align:center;min-width:220px;font-family:'Times New Roman',Times,serif;font-size:12.5px;color:#1a1a1a;">
            <div style="text-decoration:underline;font-weight:normal;">
              Yogyakarta,&nbsp; ${formatMonthYear(date)}
            </div>
            <div style="margin-top:4px;margin-bottom:48px;">&nbsp;</div>
            <div style="font-weight:700;border-top:1px solid #1a1a1a;padding-top:4px;display:inline-block;min-width:180px;">
              Cahya Damarjati, S.T., M.Eng., Ph.D.
            </div>
          </div>
        </div>
      </div>`;
  }

  /**
   * Open the report in a new print window and trigger browser print dialog.
   */
  function printReport(opts) {
    const html = buildHTML(opts);

    // Save report data to localStorage so it can be read by print-report.html fallback
    try {
      localStorage.setItem('print-report-data', JSON.stringify(opts));
    } catch (e) {
      console.warn('Gagal menyimpan print-report-data ke localStorage:', e);
    }

    const sessionId = opts.sessionId || opts.id;
    const win = window.open('', '_blank', 'width=900,height=700');
    if (win) {
      win.document.write(`<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="UTF-8">
  <title>Laporan Ujian — ${opts.candidateName || 'Peserta'}</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body { background: #fff; }
    @page { size: A4; margin: 8mm 12mm; }
    @media print {
      body { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .no-print { display: none !important; }
    }
    .btn-print {
      display: block;
      margin: 16px auto;
      padding: 10px 32px;
      background: #c8102e;
      color: #fff;
      border: none;
      border-radius: 6px;
      font-size: 14px;
      font-weight: 700;
      cursor: pointer;
      letter-spacing: 0.04em;
    }
    .btn-print:hover { background: #a00d24; }
  </style>
</head>
<body>
  <button class="btn-print no-print" onclick="window.print()">🖨️ Cetak / Download PDF</button>
  ${html}
  <script>
    // Auto-close after print
    window.onafterprint = function() { window.close(); };
  </script>
</body>
</html>`);
      win.document.close();
    } else {
      // Fallback: we are likely in a sandboxed Office Add-in taskpane
      if (typeof Office !== 'undefined' && Office.context && Office.context.ui) {
        const dialogUrl = window.location.origin + '/print-report.html' + (sessionId ? `?sessionId=${sessionId}` : '');
        Office.context.ui.displayDialogAsync(dialogUrl, { height: 90, width: 90, displayInIframe: false }, function (asyncResult) {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.error('Gagal membuka dialog cetak:', asyncResult.error.message);
            if (typeof Swal !== 'undefined') {
              Swal.fire('Gagal membuka cetak laporan', asyncResult.error.message, 'error');
            } else {
              alert('Gagal membuka dialog cetak: ' + asyncResult.error.message);
            }
          }
        });
      } else {
        // Fallback for browsers that block popup
        const fallbackUrl = '/print-report.html' + (sessionId ? `?sessionId=${sessionId}` : '');
        if (typeof Swal !== 'undefined') {
          Swal.fire({
            icon: 'warning',
            title: 'Popup Diblokir',
            html: `Browser Anda memblokir popup. Silakan izinkan popup untuk website ini atau klik link berikut untuk membuka laporan:<br><br><a href="${fallbackUrl}" target="_blank" style="color:#c8102e;font-weight:bold;">Buka Laporan Ujian</a>`
          });
        } else {
          alert('Popup diblokir browser. Silakan buka ' + fallbackUrl + ' untuk mencetak.');
        }
      }
    }
  }

  return { buildHTML, printReport, formatDate, examTypeLabel, levelLabel, groupByCategory };
})();
