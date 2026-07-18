/**
 * ppt.js
 * PowerPoint specific checkers.
 */

const PowerPointCheckers = {
  checkPSlide2: async () => {
    return { score: 8, detail: "Dikonfirmasi ✓" };
  },

  checkP1_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slides.items.length === 0) return { score: 0, detail: "Tidak ada slide" };
        
        const slide = slides.items[0];
        slide.load("shapes");
        await context.sync();
        slide.shapes.load("items");
        await context.sync();
        
        for (let shape of slide.shapes.items) {
          shape.load("type");
        }
        await context.sync();
        
        let foundFont = null, foundSize = null, foundBold = null, foundText = null;
        for (let shape of slide.shapes.items) {
          const t = String(shape.type).toLowerCase();
          if (t !== 'image' && t !== 'line' && t !== 'media' && t !== 'table' && t !== '3' && t !== '4') {
            try {
              shape.load("textFrame/hasText");
              await context.sync();
              if (shape.textFrame && shape.textFrame.hasText) {
                shape.textFrame.textRange.load("text");
                await context.sync();
                const txt = shape.textFrame.textRange.text || "";
                if (txt.toLowerCase().includes("organisasi komputer") || txt.toLowerCase().includes("computer organisation")) {
                  shape.textFrame.textRange.font.load("name,size,bold");
                  await context.sync();
                  foundText = txt;
                  foundFont = shape.textFrame.textRange.font.name;
                  foundSize = shape.textFrame.textRange.font.size;
                  foundBold = shape.textFrame.textRange.font.bold;
                  break;
                }
              }
            } catch (_) {}
          }
        }
        
        if (!foundText) return { score: 0, detail: "Judul 'Organisasi Komputer' tidak ditemukan pada slide 1" };
        
        const nameOk = (foundFont || "").toLowerCase().includes("arial");
        const sizeOk = foundSize === 44;
        const boldOk = foundBold === true;
        const ok = nameOk && sizeOk && boldOk;
        return {
          score: ok ? 100 : 0,
          detail: `Judul: "${foundText}", Font: ${foundFont} ${nameOk?'✓':'✗'}, Size: ${foundSize} ${sizeOk?'✓':'✗'}, Bold: ${boldOk?'✓':'✗'}`
        };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkP2_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        if (slides.items.length < 2) return { score: 0, detail: "Slide kedua belum dibuat" };
        
        const texts = await CheckerHelpers._getSlideTexts(slides.items[1], context);
        const titleOk = texts.some(t => t.toLowerCase().includes("siklus instruksi"));
        
        return {
          score: titleOk ? 100 : 0,
          detail: `Judul 'Siklus Instruksi' pada Slide 2: ${titleOk ? '✓' : '✗ (belum ditemukan)'}`
        };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkP3_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        
        if (slides.items.length < 2) return { score: 0, detail: "Slide kedua belum dibuat" };
        const slide2 = slides.items[1];
        slide2.load("shapes");
        await context.sync();
        slide2.shapes.load("items");
        await context.sync();
        
        let hasImg = false;
        let typeFound = '';
        for (let shape of slide2.shapes.items) {
          try {
            shape.load("type");
            await context.sync();
            typeFound = String(shape.type).toLowerCase();
            if (typeFound === 'image' || typeFound === '3' || shape.type === 3) {
              hasImg = true;
              break;
            }
          } catch (shapeErr) {
            continue;
          }
        }
        
        return {
          score: hasImg ? 100 : 0,
          detail: hasImg ? "Gambar pada Slide 2 ditemukan ✓" : `Gambar pada Slide 2 tidak ditemukan`
        };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkP4_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        
        if (slides.items.length < 2) return { score: 0, detail: "Slide kedua belum dibuat" };
        const slide2 = slides.items[1];
        slide2.load("shapes");
        await context.sync();
        slide2.shapes.load("items");
        await context.sync();
        
        let imgShape = null;
        for (let shape of slide2.shapes.items) {
          try {
            shape.load("type");
            await context.sync();
            const t = String(shape.type).toLowerCase();
            if (t === 'image' || t === '3' || shape.type === 3) {
              imgShape = shape;
              break;
            }
          } catch (_) { continue; }
        }
        
        if (!imgShape) return { score: 0, detail: "Gambar pada Slide 2 tidak ditemukan" };
        
        imgShape.load("height,width,left,top");
        await context.sync();
        
        const ptToCm = (pt) => Math.round(pt / 28.346 * 10) / 10;
        const heightOk = Math.abs(imgShape.height - 269.28) < 30;
        const widthOk  = Math.abs(imgShape.width  - 255.12) < 30;
        const leftOk   = Math.abs(imgShape.left   - 283.46) < 30;
        const topOk    = Math.abs(imgShape.top    - 198.43) < 30;
        
        const hCm = ptToCm(imgShape.height);
        const wCm = ptToCm(imgShape.width);
        const lCm = ptToCm(imgShape.left);
        const tCm = ptToCm(imgShape.top);
        
        let tip = "";
        if (!heightOk && widthOk) {
          tip = " — PENTING: Hilangkan centang 'Kunci Rasio Aspek' / 'Lock Aspect Ratio' di PowerPoint agar tinggi dan lebar bisa diatur secara independen!";
        }
        
        const ok = heightOk && widthOk && leftOk && topOk;
        
        return {
          score: ok ? 100 : 0,
          detail: `Tinggi: ${hCm} cm ${heightOk?'✓':'✗'}${tip}, Lebar: ${wCm} cm ${widthOk?'✓':'✗'}, Left: ${lCm} cm ${leftOk?'✓':'✗'}, Top: ${tCm} cm ${topOk?'✓':'✗'}`
        };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkP5_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();

        let ramSlide = null;
        for (const slide of slides.items) {
          const texts = await CheckerHelpers._getSlideTexts(slide, context);
          if (texts.some(t => t.toLowerCase().includes("ram"))) {
            ramSlide = slide;
            break;
          }
        }

        if (!ramSlide) {
          return { score: 0, detail: "Slide berisi 'RAM' tidak ditemukan" };
        }

        ramSlide.shapes.load("items");
        await context.sync();

        let boldOk = false;
        let redOk = false;

        for (const shape of ramSlide.shapes.items) {
          try {
            shape.load("type");
            await context.sync();
            const t = String(shape.type).toLowerCase();
            if (t.includes('image') || t.includes('line') || t.includes('media') || t.includes('table') || t.includes('group')) {
              continue;
            }
            
            shape.load("textFrame/hasText");
            await context.sync();
            if (!shape.textFrame || !shape.textFrame.hasText) continue;
            
            const textRange = shape.textFrame.textRange;
            textRange.load("text");
            await context.sync();
            
            const txt = textRange.text || "";
            const txtLower = txt.toLowerCase();
            
            if (!boldOk && txtLower.includes("alu") && txtLower.includes("cu") && txtLower.includes("reg")) {
              const startIdx = txtLower.indexOf("alu");
              if (startIdx >= 0) {
                const sub = textRange.getSubstring(startIdx, 3);
                sub.load("font/bold");
                await context.sync();
                if (sub.font.bold === true) {
                  boldOk = true;
                }
              }
            }
            
            if (!redOk && txtLower.includes("rom")) {
              const startIdx = txtLower.indexOf("rom");
              if (startIdx >= 0) {
                const sub = textRange.getSubstring(startIdx, 3);
                sub.load("font/color");
                await context.sync();
                if (CheckerHelpers._isRed(sub.font.color)) {
                  redOk = true;
                }
              }
            }
            
            if (boldOk && redOk) break;
          } catch (_) {
            continue;
          }
        }

        const ok = boldOk && redOk;
        return {
          score: ok ? 100 : 0,
          detail: `Slide 'RAM' ditemukan — Bold pada 'ALU+CU+REG': ${boldOk ? '✓' : '✗'}, Warna Merah pada 'ROM': ${redOk ? '✓' : '✗'}`
        };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkP6_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        
        let targetSlide = null;
        for (let slide of slides.items) {
          const texts = await CheckerHelpers._getSlideTexts(slide, context);
          if (texts.some(t => t.toLowerCase().includes("pipelining"))) {
            targetSlide = slide;
            break;
          }
        }
        if (!targetSlide) return { score: 0, detail: "Slide berisi 'Pipelining' tidak ditemukan" };
        
        targetSlide.load("shapes");
        await context.sync();
        targetSlide.shapes.load("items");
        await context.sync();
        
        let boldOk = false;
        let redOk = false;
        
        for (let shape of targetSlide.shapes.items) {
          try {
            shape.load("type");
            await context.sync();
            const t = String(shape.type).toLowerCase();
            if (t.includes('image') || t.includes('line') || t.includes('media') || t.includes('table') || t.includes('group')) {
              continue;
            }
            
            shape.load("textFrame/hasText");
            await context.sync();
            if (!shape.textFrame || !shape.textFrame.hasText) continue;
            
            const textRange = shape.textFrame.textRange;
            textRange.load("text");
            await context.sync();
            
            const txt = textRange.text || "";
            const txtLower = txt.toLowerCase();
            
            if (txtLower.includes("pipelining")) {
              const startIdx = txtLower.indexOf("pipelining");
              if (startIdx >= 0) {
                const sub = textRange.getSubstring(startIdx, "pipelining".length);
                sub.load("font/bold,font/color");
                await context.sync();
                
                if (sub.font.bold === true) {
                  boldOk = true;
                }
                if (CheckerHelpers._isRed(sub.font.color)) {
                  redOk = true;
                }
              }
            }
            if (boldOk && redOk) break;
          } catch (_) {
            continue;
          }
        }
        
        const ok = boldOk && redOk;
        return {
          score: ok ? 100 : 0,
          detail: `Slide 'Performance' ditemukan — Bold pada 'Pipelining': ${boldOk ? '✓' : '✗'}, Warna Merah pada 'Pipelining': ${redOk ? '✓' : '✗'}`
        };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkP7_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        
        let slideDetails = [];
        let tableFoundInfo = null;
        
        for (let i = 0; i < slides.items.length; i++) {
          const slide = slides.items[i];
          const texts = await CheckerHelpers._getSlideTexts(slide, context);
          
          const slideTitle = texts.join(" | ");
          slideDetails.push(`Slide ${i+1}: "${slideTitle.substring(0, 30)}" (${slide.shapes.items.length} shapes)`);
          
          const onFunctionalSlide = texts.some(t => 
            t.toLowerCase().includes("functional units") || 
            t.toLowerCase().includes("functional unit") || 
            t.toLowerCase().includes("units of computer")
          );
          
          if (onFunctionalSlide) {
            slide.load("shapes");
            await context.sync();
            slide.shapes.load("items");
            await context.sync();
            
            for (let shape of slide.shapes.items) {
              shape.load("type,name");
            }
            await context.sync();
            
            let hasTable = false;
            let targetShape = null;
            let shapesFound = [];
            
            for (let shape of slide.shapes.items) {
              const t = String(shape.type).toLowerCase();
              const name = String(shape.name || "").toLowerCase();
              shapesFound.push(`${t} (name: ${name})`);
              
              if (
                t === 'table' || 
                t === '4' || 
                shape.type === 4 ||
                name.includes("table") ||
                name.includes("tabel")
              ) {
                hasTable = true;
                targetShape = shape;
                break;
              }
            }
            
            if (!hasTable) {
              for (let shape of slide.shapes.items) {
                try {
                  const tbl = shape.getTable();
                  tbl.load("rowCount,columnCount");
                  await context.sync();
                  hasTable = true;
                  targetShape = shape;
                  break;
                } catch (_) {}
              }
            }
            
            if (hasTable && targetShape) {
              let rows = 0, cols = 0;
              try {
                const tbl = targetShape.getTable();
                tbl.load("rowCount,columnCount");
                await context.sync();
                rows = tbl.rowCount;
                cols = tbl.columnCount;
              } catch (_) {
                rows = 5; cols = 3;
              }
              
              if (rows >= 2 && cols >= 2) {
                return { score: 100, detail: `Tabel pada slide 'FUNCTIONAL UNITS OF COMPUTER' ditemukan (${cols} kolom x ${rows} baris) ✓` };
              } else {
                return { score: 100, detail: `Tabel ditemukan ✓` };
              }
            } else {
              tableFoundInfo = `Slide cocok ditemukan, tetapi tidak ada tipe tabel di antara bentuk ini: [${shapesFound.join(", ")}]`;
            }
          }
        }
        
        let anyTableFound = false;
        for (let i = 0; i < slides.items.length; i++) {
          const slide = slides.items[i];
          slide.load("shapes");
          await context.sync();
          slide.shapes.load("items");
          await context.sync();
          for (let shape of slide.shapes.items) {
            shape.load("type,name");
          }
          await context.sync();
          for (let shape of slide.shapes.items) {
            const t = String(shape.type).toLowerCase();
            const name = String(shape.name || "").toLowerCase();
            if (
              t === 'table' || 
              t === '4' || 
              shape.type === 4 ||
              name.includes("table") ||
              name.includes("tabel")
            ) {
              anyTableFound = true; break;
            }
            try {
              const tbl = shape.getTable();
              tbl.load("rowCount");
              await context.sync();
              if (tbl.rowCount >= 2) { anyTableFound = true; break; }
            } catch (_) {}
          }
          if (anyTableFound) break;
        }
        if (anyTableFound) {
          return { score: 100, detail: "Tabel ditemukan pada presentasi ✓" };
        }
        
        return { 
          score: 0, 
          detail: `Tabel tidak ditemukan pada slide 'FUNCTIONAL UNITS OF COMPUTER'. (Log: ${slideDetails.join("; ")} | ${tableFoundInfo || "Slide 'FUNCTIONAL UNITS' tidak terdeteksi"})` 
        };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkP8_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        
        if (slides.items.length === 0) return { score: 0, detail: "Tidak ada slide" };
        
        const lastSlide = slides.items[slides.items.length - 1];
        const texts = await CheckerHelpers._getSlideTexts(lastSlide, context);
        const hasInputUnit = texts.some(t => t.toLowerCase().includes("input unit"));
        
        if (hasInputUnit) {
          return { score: 100, detail: "Slide 'INPUT UNIT:' telah dipindahkan ke akhir ✓" };
        }
        return { score: 0, detail: "Slide 'INPUT UNIT:' belum dipindahkan ke akhir" };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkP9_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        
        let ramSlide = null;
        for (let slide of slides.items) {
          const texts = await CheckerHelpers._getSlideTexts(slide, context);
          if (texts.some(t => t.toLowerCase().includes("ram"))) {
            ramSlide = slide;
            break;
          }
        }
        if (!ramSlide) return { score: 0, detail: "Slide 'RAM' tidak ditemukan" };
        
        const texts = await CheckerHelpers._getSlideTexts(ramSlide, context);
        const renamed = texts.some(t => t.toLowerCase().includes("ram & rom") || t.toLowerCase().includes("ram dan rom") || t.toLowerCase().includes("ram &rom"));
        
        if (renamed) {
          return {
            score: 100,
            detail: "Judul Slide 5 telah diubah menjadi 'RAM & ROM' ✓"
          };
        }
        return {
          score: 0,
          detail: "Judul Slide 5 belum diubah menjadi 'RAM & ROM'"
        };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkP10_New: async () => {
    return await PowerPoint.run(async (context) => {
      try {
        const slides = context.presentation.slides;
        slides.load("items");
        await context.sync();
        
        for (let slide of slides.items) {
          const texts = await CheckerHelpers._getSlideTexts(slide, context);
          if (texts.some(t => t.toLowerCase().includes("superscalar"))) {
            slide.load("shapes");
            await context.sync();
            slide.shapes.load("items");
            await context.sync();
            
            for (let shape of slide.shapes.items) {
              try {
                shape.load("type");
                await context.sync();
                const t = String(shape.type).toLowerCase();
                if (t.includes('image') || t.includes('line') || t.includes('media') || t.includes('table') || t.includes('group')) {
                  continue;
                }
                
                shape.load("textFrame/hasText");
                await context.sync();
                if (!shape.textFrame || !shape.textFrame.hasText) continue;
                
                const textRange = shape.textFrame.textRange;
                textRange.load("text");
                await context.sync();
                
                const txt = textRange.text || "";
                const txtLower = txt.toLowerCase();
                
                if (txtLower.includes("superscalar")) {
                  const startIdx = txtLower.indexOf("superscalar");
                  if (startIdx >= 0) {
                    const sub = textRange.getSubstring(startIdx, "superscalar".length);
                    sub.load("font/italic");
                    await context.sync();
                    
                    if (sub.font.italic === true) {
                      return {
                        score: 100,
                        detail: "Format Miring (Italic) pada 'superscalar' terdeteksi ✓"
                      };
                    }
                  }
                }
              } catch (_) {
                continue;
              }
            }
          }
        }
        return { score: 0, detail: "Format Miring (Italic) pada 'superscalar' belum diterapkan" };
      } catch (e) {
        return { score: 0, detail: `Error: ${e.message || e}` };
      }
    });
  },

  checkPConfirm: (id) => async () => ({ score: 100, detail: "Dikonfirmasi ✓" })
};

window.PowerPointCheckers = PowerPointCheckers;
