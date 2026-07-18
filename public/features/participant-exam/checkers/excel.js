/**
 * excel.js
 * Excel specific checkers.
 */

const ExcelCheckers = {
  checkE1: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1:O1");
      range.load("format/horizontalAlignment,format/font/bold,mergeCells");
      await context.sync();
      const isMerged = range.mergeCells === true;
      const isBold = range.format.font.bold === true;
      const isCenter = range.format.horizontalAlignment === "Center";
      let score = (isMerged ? 4 : 0) + (isBold ? 3 : 0) + (isCenter ? 3 : 0);
      return { score: score * 10, detail: `Merged: ${isMerged ? '✓' : '✗'}, Bold: ${isBold ? '✓' : '✗'}, Center: ${isCenter ? '✓' : '✗'}` };
    });
  },

  checkE2: async () => {
    return await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("C3:M3");
      range.load("numberFormat");
      await context.sync();
      const fmt = String(range.numberFormat[0][0]).toLowerCase();
      const isTime = fmt.includes('h:mm') || fmt.includes('am/pm') || fmt.includes(':');
      return { score: isTime ? 10 : 0, detail: isTime ? "Format Waktu ✓" : "Format belum sesuai" };
    });
  },

  checkE3: async () => {
    return await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("N4:N34");
      range.load("formulas");
      await context.sync();
      const allF = range.formulas.flat().join(' ').toUpperCase();
      const hasSum = allF.includes('SUM');
      return { score: hasSum ? 100 : 0, detail: hasSum ? "Rumus SUM ditemukan ✓" : "Gunakan rumus SUM di kolom N" };
    });
  },

  checkE4: async () => {
    return await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("C36:M36");
      range.load("formulas");
      await context.sync();
      const hasSum = range.formulas[0].some(f => String(f).toUpperCase().includes('SUM'));
      return { score: hasSum ? 10 : 0, detail: hasSum ? "Rumus SUM per jam ✓" : "Gunakan rumus SUM" };
    });
  },

  checkE5: async () => {
    return await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("O4:O34");
      range.load("formulas");
      await context.sync();
      const allF = range.formulas.flat().join(' ').toUpperCase();
      const ok = allF.includes('-') && allF.includes('AVERAGE');
      return { score: ok ? 100 : 0, detail: ok ? "Rumus Selisih & Average ✓" : "Gunakan formula =N4-AVERAGE(C4:M4) di kolom O" };
    });
  },

  checkE6: async () => {
    return await Excel.run(async (context) => {
      return { score: 10, detail: "Dikonfirmasi ✓" };
    });
  },

  checkE7: async () => {
    return await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("A3:O3");
      const border = range.format.borders.getItem("EdgeBottom");
      border.load("style");
      await context.sync();
      const hasBorders = border.style !== "None";
      return { score: hasBorders ? 10 : 0, detail: hasBorders ? "Border ditemukan ✓" : "Berikan All Borders" };
    });
  },

  checkE8: async () => { return { score: 10, detail: "Dikonfirmasi ✓" }; },

  checkE9: async () => {
    return await Excel.run(async (context) => {
      const range = context.workbook.worksheets.getActiveWorksheet().getRange("N38:N39");
      range.load("formulas");
      await context.sync();
      const f = range.formulas.flat().join(' ').toUpperCase();
      const hasMax = f.includes('MAX');
      const hasAvg = f.includes('AVERAGE');
      return { score: (hasMax ? 5 : 0) + (hasAvg ? 5 : 0), detail: `MAX: ${hasMax ? '✓' : '✗'}, AVG: ${hasAvg ? '✓' : '✗'}` };
    });
  },

  checkE10: async () => { return { score: 10, detail: "Dikonfirmasi ✓" }; },

  checkE2_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("B2:D2");
      range.load("values,mergeCells");
      await context.sync();
      const val = (range.values[0][0] || "").toString().toLowerCase().trim();
      const isTextOk = val.includes("database") || val.includes("supply");
      const isMerged = range.mergeCells === true;
      const ok = isTextOk && isMerged;
      return {
        score: ok ? 100 : 0,
        detail: ok ? "B2:D2 Merged & Teks Sesuai ✓" : `Merge: ${range.mergeCells ? 'Ya' : 'Tidak'}, Teks: '${val}'. Pastikan Merge B2:D2.`
      };
    });
  },

  checkE3_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("B2");
      range.load("format/fill/color,format/font/color,format/horizontalAlignment");
      await context.sync();
      const fill = (range.format.fill.color || "").toLowerCase();
      const font = (range.format.font.color || "").toLowerCase();
      const isGreen = fill !== "#ffffff" && fill !== "" && fill !== "none" && fill !== "#000000";
      const isRed = font !== "#000000" && font !== "" && font !== "#ffffff";
      const isCenter = range.format.horizontalAlignment === "Center";
      const ok = isGreen && isRed && isCenter;
      return {
        score: ok ? 100 : 0,
        detail: `Bg: ${fill} ${isGreen ? '✓' : '(harus berwarna)'}, Font: ${font} ${isRed ? '✓' : '(harus berwarna selain hitam)'}, Align: ${range.format.horizontalAlignment}`
      };
    });
  },

  checkE4_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const usedRange = sheet.getUsedRangeOrNullObject();
      await context.sync();
      if (usedRange.isNullObject) return { score: 0, detail: "Sheet kosong" };
      let search = usedRange.findOrNullObject("Adventure Works", { completeMatch: false, matchCase: false });
      await context.sync();
      if (search.isNullObject) {
        search = usedRange.findOrNullObject("Adventure", { completeMatch: false, matchCase: false });
        await context.sync();
      }
      if (search.isNullObject) return { score: 0, detail: "Teks 'Adventure Works' tidak ditemukan di sheet ini" };
      search.load("format/font/name,format/font/size,format/font/bold,address");
      await context.sync();
      const fontOk = search.format.font.name.toLowerCase().includes("arial");
      const sizeOk = search.format.font.size >= 14;
      const boldOk = search.format.font.bold === true;
      const ok = fontOk && sizeOk && boldOk;
      return { score: ok ? 100 : 0, detail: `Ditemukan di ${search.address}. Font: ${search.format.font.name}, Size: ${search.format.font.size}, Bold: ${boldOk}` };
    });
  },

  checkE5_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const itemCell = sheet.getRange("A4:B4");
      itemCell.load("values");
      await context.sync();
      const hasItemText = itemCell.values.flat().some(v => String(v || "").toLowerCase().trim() === "item");
      if (!hasItemText) {
        return { score: 0, detail: "Silakan buka/download file soal terlebih dahulu" };
      }
      const range = sheet.getRange("A4:F4");
      range.load("format/horizontalAlignment");
      await context.sync();
      const ok = range.format.horizontalAlignment === "Left";
      return { score: ok ? 100 : 0, detail: ok ? "Rata Kiri ✓" : "A4:F4 harus diatur rata kiri (saat ini: " + range.format.horizontalAlignment + ")" };
    });
  },

  checkE6_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRangeOrNullObject();
      await context.sync();
      if (range.isNullObject) return { score: 0, detail: "Tabel tidak ditemukan" };
      const border = range.format.borders.getItem("EdgeBottom");
      border.load("style");
      await context.sync();
      const ok = border.style !== "None";
      return { score: ok ? 100 : 0, detail: ok ? "All Border ✓" : "Gunakan All Borders pada tabel" };
    });
  },

  checkE7_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const itemCell = sheet.getRange("A4:B4");
      itemCell.load("values");
      await context.sync();
      const hasItemText = itemCell.values.flat().some(v => String(v || "").toLowerCase().trim() === "item");
      if (!hasItemText) {
        return { score: 0, detail: "Silakan buka/download file soal terlebih dahulu" };
      }
      const range = sheet.getRange("F:F");
      range.load("format/fill/color");
      await context.sync();
      const color = range.format.fill.color;
      const ok = color === "#FFFFFF" || color === "" || color === "none";
      return { score: ok ? 100 : 0, detail: ok ? "No Fill pada kolom F ✓" : "Kolom F masih memiliki warna fill" };
    });
  },

  checkE8_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A4");
      range.load("values");
      await context.sync();
      const val = (range.values[0][0] || "").toString().toLowerCase().trim();
      const ok = val === "no";
      return { score: ok ? 100 : 0, detail: ok ? "Kolom No ditemukan ✓" : "A4 harus berisi 'No'" };
    });
  },

  checkE9_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getUsedRangeOrNullObject();
      await context.sync();
      if (range.isNullObject) return { score: 100, detail: "Data sudah kosong" };
      range.load("values");
      await context.sync();
      const hasTen = range.values.some(row => row.includes(10) || row.includes("10"));
      return { score: !hasTen ? 100 : 0, detail: !hasTen ? "Data No 10 terhapus ✓" : "Data No 10 masih ditemukan" };
    });
  },

  checkE10_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const headerRange = sheet.getRange("A4:Z4");
      headerRange.load("values");
      await context.sync();
      const colIdx = headerRange.values[0].findIndex(v => v && v.toString().toLowerCase().includes("price"));
      if (colIdx === -1) return { score: 0, detail: "Kolom 'Price' tidak ditemukan di baris 4" };
      const dataCell = sheet.getRangeByIndexes(5, colIdx, 1, 1);
      const entireCol = dataCell.getEntireColumn();
      entireCol.load("columnWidth,format/columnWidth");
      dataCell.load("numberFormat");
      await context.sync();
      const width = entireCol.columnWidth || entireCol.format.columnWidth || 0;
      const widthOk = width >= 60;
      const fmt = (dataCell.numberFormat[0][0] || "").toString();
      const fmtOk = fmt.includes('$') || fmt.includes('Rp') || fmt.includes('"$"');
      return {
        score: (widthOk ? 50 : 0) + (fmtOk ? 50 : 0),
        detail: `Lebar: ${Math.round(width)}, Format: ${fmt}`
      };
    });
  },

  checkE11_New: async () => {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const names = sheets.items.map(s => s.name.toLowerCase().trim());
      const ok = names.includes("database") && names.includes("januari");
      return { score: ok ? 100 : 0, detail: `Sheets: ${names.join(', ')}` };
    });
  },

  checkE12_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A2");
      range.load("numberFormat,values,valueTypes");
      await context.sync();
      const fmt = (range.numberFormat[0][0] || "").toString().toLowerCase();
      const isTimeFmt = fmt.includes('h:mm') || fmt.includes(':') || fmt.includes('am/pm');
      const val = range.values[0][0];
      const hasValue = val !== null && val !== "" && val !== undefined;
      const ok = isTimeFmt && hasValue;
      let detail = `Format: ${fmt}`;
      if (!hasValue) detail += " (Isi cell A2 dengan waktu)";
      return { score: ok ? 100 : 0, detail: detail };
    });
  },

  checkE13_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("G6");
      range.load("formulas");
      await context.sync();
      const formula = (range.formulas[0][0] || "").toString().toUpperCase();
      const isProduct = formula.includes("D6") && formula.includes("F6") && formula.includes("*");
      return {
        score: isProduct ? 100 : 0,
        detail: isProduct ? "Rumus Perkalian G6 ditemukan ✓" : "Gunakan rumus =D6*F6 di sel G6"
      };
    });
  },

  checkE14_New: async () => {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      let sheet = sheets.items[0];
      if (!sheet) sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A40");
      range.load("values");
      await context.sync();
      const val = (range.values[0][0] || "").toString().trim().toUpperCase();
      const isOk = val === "LSI UMY";
      return { 
        score: isOk ? 100 : 0, 
        detail: isOk ? "Teks 'LSI UMY' di sel A40 ditemukan ✓" : `Ketik 'LSI UMY' di sel A40 pada sheet pertama (Database) (Terbaca: "${val}")` 
      };
    });
  },

  checkE15_New: async () => {
    return await Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      let sheet = sheets.items.find(s => s.name.toLowerCase().includes("januari"));
      if (!sheet) sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("F36");
      range.load("formulas,values");
      await context.sync();
      const formula = (range.formulas[0][0] || "").toString().toUpperCase();
      const value = range.values[0][0];
      const hasSum = formula.includes("SUM") && (formula.includes("F6") || formula.includes("F35"));
      let detail = "";
      if (hasSum) {
        detail = "Rumus SUM di F36 ditemukan ✓";
      } else {
        detail = `Gunakan rumus =SUM(F6:F35) di sel F36. Terbaca rumus: "${formula}", nilai: "${value}"`;
      }
      return { 
        score: hasSum ? 100 : 0, 
        detail: detail
      };
    });
  },

  checkE16_New: async () => {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const b39 = sheet.getRange("B39");
      const avgRange = sheet.getRange("C39:N39");
      b39.load("values");
      avgRange.load("formulas,values");
      await context.sync();
      const label = (b39.values[0][0] || "").toString().trim();
      const textOk = label.toLowerCase().includes("average");
      const formulasList = avgRange.formulas[0].filter(f => f && f.toString().trim() !== "");
      const formulaOk = avgRange.formulas[0].some(f => f && f.toString().toUpperCase().includes("AVERAGE"));
      let detail = "";
      if (textOk && formulaOk) {
        detail = "Label & Rumus Average ditemukan ✓";
      } else {
        detail = `Label B39: "${label}", Terbaca rumus di C39:N39: [${formulasList.slice(0, 3).join(', ')}...]`;
      }
      return { score: (textOk?50:0)+(formulaOk?50:0), detail: detail };
    });
  },

  checkEConfirm: (id) => async () => ({ score: 100, detail: "Dikonfirmasi ✓" })
};

window.ExcelCheckers = ExcelCheckers;
