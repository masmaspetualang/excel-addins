/**
 * word.js
 * Word specific checkers.
 */

const WordCheckers = {
  checkWJudul: async () => {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const search = body.search("Penyulingan Minyak Atsiri", { matchCase: false });
      search.load("font/bold,alignment");
      await context.sync();
      if (search.items.length > 0) {
        const item = search.items[0];
        const ok = item.font.bold;
        return { score: ok ? 5 : 2, detail: ok ? "Judul Bold ✓" : "Judul ditemukan tapi tidak Bold" };
      }
      return { score: 0, detail: "Judul tidak ditemukan" };
    });
  },

  checkWFont: async () => {
    return await Word.run(async (context) => {
      const body = context.document.body;
      const search = body.search("Penyulingan Minyak Atsiri", { matchCase: false });
      search.load("font/name,font/size");
      await context.sync();
      if (search.items.length > 0) {
        const f = search.items[0].font;
        const nameOk = f.name.toLowerCase().includes('tahoma');
        const sizeOk = f.size === 14;
        return { score: (nameOk ? 3 : 0) + (sizeOk ? 2 : 0), detail: `Font: ${f.name}, Size: ${f.size}` };
      }
      return { score: 0, detail: "Judul tidak ditemukan" };
    });
  },

  checkWReplace: async () => {
    return await Word.run(async (context) => {
      const search = context.document.body.search("arsiri", { matchCase: false });
      search.load("items");
      await context.sync();
      const ok = search.items.length === 0;
      return { score: ok ? 5 : 0, detail: ok ? "Semua 'arsiri' telah diganti ✓" : "Masih ada kata 'arsiri'" };
    });
  },

  checkWDictFont: async () => {
    return await Word.run(async (context) => {
      const search = context.document.body.search("A New Dictionary of Chemistry", { matchCase: false });
      search.load("font/name,font/size");
      await context.sync();
      if (search.items.length > 0) {
        const f = search.items[0].font;
        const ok = f.name.toLowerCase().includes('trebuchet') && f.size === 11;
        return { score: ok ? 5 : 2, detail: `Font: ${f.name}, Size: ${f.size}` };
      }
      return { score: 0, detail: "Teks tidak ditemukan" };
    });
  },

  checkWSuperscript: async () => {
    return await Word.run(async (context) => {
      return { score: 5, detail: "Dikonfirmasi ✓" };
    });
  },

  checkWHighlight: async () => {
    return await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/font/highlightColor");
      await context.sync();
      const color = paragraphs.items[0].font.highlightColor;
      const ok = color && color !== "#000000" && color !== "";
      return { score: ok ? 5 : 0, detail: ok ? `Highlight: ${color} ✓` : "Tidak ada highlight" };
    });
  },

  checkWJustify: async () => {
    return await Word.run(async (context) => {
      const body = context.document.body;
      body.load("alignment");
      await context.sync();
      const p = body.paragraphs.getFirst();
      p.load("alignment");
      await context.sync();
      const ok = p.alignment === 'Justified';
      return { score: ok ? 5 : 0, detail: `Alignment: ${p.alignment}` };
    });
  },

  checkWTable: async () => {
    return await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();
      const ok = tables.items.length > 0;
      return { score: ok ? 10 : 0, detail: ok ? `${tables.items.length} tabel ditemukan ✓` : "Tabel belum dibuat" };
    });
  },

  checkWFooter: async () => {
    return await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load("items/footers");
      await context.sync();
      const footer = sections.items[0].footers.getFirst();
      footer.load("type");
      await context.sync();
      const ok = footer.type !== 'None';
      return { score: ok ? 5 : 0, detail: ok ? "Footer aktif ✓" : "Footer tidak ditemukan" };
    });
  },

  checkW1_New: async () => {
    return await Word.run(async (context) => {
      let search = context.document.body.search("Penyulingan Minyak Atsiri", { matchCase: false });
      search.load("font/bold,font/name,font/size");
      await context.sync();
      
      if (search.items.length === 0) {
        search = context.document.body.search("Penyulingan Minyak Arsiri", { matchCase: false });
        search.load("font/bold,font/name,font/size");
        await context.sync();
      }
      
      if (search.items.length > 0) {
        const item = search.items[0];
        const p = item.paragraphs.getFirst();
        p.load("alignment");
        await context.sync();
        
        const boldOk = item.font && item.font.bold === true;
        const alignOk = p.alignment ? p.alignment.toLowerCase().includes("center") : false;
        const fontOk = (item.font && item.font.name) ? item.font.name.toLowerCase().includes("tahoma") : false;
        const sizeOk = (item.font && item.font.size) ? (item.font.size >= 13 && item.font.size <= 15) : false;
        const ok = boldOk && alignOk && fontOk && sizeOk;
        
        return { 
          score: ok ? 100 : 0, 
          detail: `Bold: ${boldOk ? '✓' : '✗'}, Rata Tengah: ${alignOk ? '✓' : '✗'}, Font Tahoma: ${fontOk ? '✓' : '✗'}, Ukuran 14pt: ${sizeOk ? '✓' : '✗'}`
        };
      }
      return { score: 0, detail: "Judul tidak ditemukan" };
    });
  },

  checkW2_New: async () => {
    return await Word.run(async (context) => {
      const searchArsiri = context.document.body.search("arsiri", { matchCase: false });
      const searchAtsiri = context.document.body.search("atsiri", { matchCase: false });
      searchArsiri.load("items");
      searchAtsiri.load("items");
      await context.sync();
      
      const hasNoArsiri = searchArsiri.items.length === 0;
      const hasAtsiri = searchAtsiri.items.length > 0;
      const ok = hasNoArsiri && hasAtsiri;
      
      return { 
        score: ok ? 100 : 0, 
        detail: ok ? "Semua ejaan 'arsiri' telah diganti menjadi 'atsiri' ✓" : (hasAtsiri ? "Masih ada kata 'arsiri'" : "Silakan buka file soal terlebih dahulu")
      };
    });
  },

  checkW3_New: async () => {
    return await Word.run(async (context) => {
      const search = context.document.body.search("A New Dictionary of Chemistry", { matchCase: false });
      search.load("font/name,font/size");
      await context.sync();
      if (search.items.length > 0) {
        const f = search.items[0].font;
        const fontName = f && f.name ? f.name : "";
        const fontSize = f && f.size ? f.size : 0;
        const nameOk = fontName.toLowerCase().includes("trebuchet");
        const sizeOk = fontSize === 11;
        const ok = nameOk && sizeOk;
        return { 
          score: ok ? 100 : 0, 
          detail: `Font: ${fontName} ${nameOk?'✓':'✗'}, Ukuran: ${fontSize}pt ${sizeOk?'✓':'✗'}` 
        };
      }
      return { score: 0, detail: "Teks tidak ditemukan" };
    });
  },

  checkW4_New: async () => {
    return await Word.run(async (context) => {
      let superOk = false;
      const searchMm = context.document.body.search("mm2", { matchCase: false });
      searchMm.load("items");
      await context.sync();
      
      if (searchMm.items.length > 0) {
        const mmRange = searchMm.items[0];
        const searchChar2 = mmRange.search("2");
        searchChar2.load("items");
        await context.sync();
        if (searchChar2.items.length > 0) {
          const char2 = searchChar2.items[0];
          char2.load("font/superscript");
          await context.sync();
          if (char2.font.superscript === true) {
            superOk = true;
          }
        }
      }
      
      let subOk = false;
      const searchWoil = context.document.body.search("Woil", { matchCase: false });
      searchWoil.load("items");
      await context.sync();
      
      if (searchWoil.items.length > 0) {
        const woilRange = searchWoil.items[0];
        const searchOil = woilRange.search("oil");
        searchOil.load("items");
        await context.sync();
        if (searchOil.items.length > 0) {
          const oilRange = searchOil.items[0];
          oilRange.load("font/subscript");
          await context.sync();
          if (oilRange.font.subscript === true) {
            subOk = true;
          }
        }
      }
      
      const ok = superOk && subOk;
      return {
        score: ok ? 100 : 0,
        detail: `Superscript mm²: ${superOk ? '✓' : '✗'}, Subscript Woil: ${subOk ? '✓' : '✗'}`
      };
    });
  },

  checkW5_New: async () => {
    return await Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items/alignment,items/lineSpacing,items/firstLineIndent");
      await context.sync();
      
      let justifyOk = false;
      let indentOk = false;
      let spacingOk = false;
      
      for (let i = 1; i < paragraphs.items.length; i++) {
        const p = paragraphs.items[i];
        if (p.alignment === "Justified") justifyOk = true;
        if (p.firstLineIndent && Math.abs(p.firstLineIndent - 72) < 6) indentOk = true;
        if (p.lineSpacing && p.lineSpacing >= 10 && p.lineSpacing <= 14) spacingOk = true;
      }
      
      const ok = justifyOk && indentOk && spacingOk;
      return {
        score: ok ? 100 : 0,
        detail: `Rata Kiri-Kanan: ${justifyOk ? '✓' : '✗'}, Indentasi 1" (72pt): ${indentOk ? '✓' : '✗'}, Spasi Single (1.0): ${spacingOk ? '✓' : '✗'}`
      };
    });
  },

  checkW6_New: async () => {
    return await Word.run(async (context) => {
      try {
        const search = context.document.body.search("Metode Umum Penyulingan", { matchCase: false });
        search.load("items/font/color,items/font/italic");
        await context.sync();
        
        if (search.items.length === 0) {
          return { score: 0, detail: "Teks 'Metode Umum Penyulingan' tidak ditemukan" };
        }
        
        const item = search.items[0];
        const lastColor = (item.font.color || "").toLowerCase();
        const allowedBlues = [
          "blue", "#0000ff", "0000ff", 
          "#0070c0", "0070c0", 
          "#4472c4", "4472c4", 
          "#1f4e78", "1f4e78", 
          "#2f5597", "2f5597", 
          "#00b0f0", "00b0f0", 
          "#002060", "002060",
          "#4f81bd", "4f81bd",
          "#5b9bd5", "5b9bd5",
          "#418ab3", "418ab3"
        ];
        const blueFound = allowedBlues.some(b => lastColor.includes(b));
        const italicFound = item.font.italic === true;
        const ok = blueFound && italicFound;
        
        return {
          score: ok ? 100 : 0,
          detail: `Warna Biru pada subjudul: ${blueFound ? '✓' : `✗ (terbaca: "${lastColor}")`}, Format Miring (Italic): ${italicFound ? '✓' : '✗'}`
        };
      } catch (err) {
        console.error("Error in checkW6_New:", err);
        return {
          score: 0,
          detail: `Error membaca warna/format: ${err.message || err}`
        };
      }
    });
  },

  checkW7_New: async () => {
    return await Word.run(async (context) => {
      try {
        const section = context.document.sections.getFirst();
        const body = section.body;
        const paragraphs = body.paragraphs;
        paragraphs.load("items");
        await context.sync();
        
        if (paragraphs.items.length === 0) {
          return { score: 0, detail: "Dokumen kosong" };
        }
        
        const titleP = paragraphs.items[0];
        titleP.load("font/highlightColor");
        await context.sync();
        
        const hlColor = (titleP.font.highlightColor || "").toLowerCase();
        const hasYellowHighlight = hlColor === "#ffff00" || hlColor === "yellow";
        
        return {
          score: hasYellowHighlight ? 100 : 0,
          detail: `Highlight Kuning pada judul: ${hasYellowHighlight ? '✓' : `✗ (terbaca: "${hlColor}")`}`
        };
      } catch (err) {
        console.error("Error in checkW7_New:", err);
        return {
          score: 0,
          detail: `Gagal membaca highlight: ${err.message || err}`
        };
      }
    });
  },

  checkW8_New: async () => {
    return await Word.run(async (context) => {
      let hasHyperlink = false;
      let hlAddress = "";
      try {
        const search = context.document.body.search("Minyak Atsiri", { matchCase: false });
        search.load("items/hyperlink");
        await context.sync();
        
        if (search.items && search.items.length > 0) {
          for (let item of search.items) {
            if (item.hyperlink) {
              hlAddress = item.hyperlink;
              if (hlAddress.toLowerCase().includes("wikipedia.org/wiki/minyak_atsiri")) {
                hasHyperlink = true;
                break;
              }
            }
          }
        }
      } catch (e) {
        console.warn("Error loading hyperlinks:", e);
        hlAddress = `Error: ${e.message || e}`;
      }
      
      const ok = hasHyperlink;
      return {
        score: ok ? 100 : 0,
        detail: hasHyperlink ? "Tautan Hyperlink Wiki ditemukan ✓" : `Tautan Hyperlink Wiki belum sesuai (terbaca: "${hlAddress || 'tidak ada'}")`
      };
    });
  },

  checkW9_New: async () => {
    return await Word.run(async (context) => {
      const tables = context.document.body.tables;
      tables.load("items");
      await context.sync();
      
      if (tables.items.length > 0) {
        const t = tables.items[0];
        t.load("values");
        await context.sync();
        
        const firstRow = t.values[0] || [];
        const colCount = firstRow.length;
        const colOk = colCount >= 3;
        const headers = firstRow.map(v => (v || "").toString().toLowerCase().trim());
        const hasNo    = headers.some(h => h.includes("no"));
        const hasNama  = headers.some(h => h.includes("nama"));
        const hasNilai = headers.some(h => h.includes("nilai"));
        const headerOk = hasNo && hasNama && hasNilai;
        
        const ok = colOk && headerOk;
        return {
          score: ok ? 100 : 0,
          detail: `Kolom: ${colCount} (min 3) ${colOk?'✓':'✗'}, Header: [${headers.join(", ")}] ${headerOk?'✓':'— harus ada: No, Nama, Nilai'}`
        };
      }
      return { score: 0, detail: "Tabel tidak ditemukan" };
    });
  },

  checkW10_New: async () => {
    return await Word.run(async (context) => {
      const sections = context.document.sections;
      sections.load("items");
      await context.sync();
      
      const footer = sections.items[0].getFooter(Word.HeaderFooterType.primary);
      footer.load("text");
      await context.sync();
      
      const hasText = footer.text && footer.text.trim().length > 3;
      return {
        score: hasText ? 100 : 0,
        detail: hasText ? `Footer ditemukan: "${footer.text.trim()}" ✓` : "Footer kosong atau belum dibuat"
      };
    });
  },

  checkWConfirm: (id) => async () => ({ score: 100, detail: "Dikonfirmasi ✓" })
};

window.WordCheckers = WordCheckers;
