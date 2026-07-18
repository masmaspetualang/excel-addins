/**
 * helpers.js
 * Shared helpers for OfficeCheckers.
 */

const CheckerHelpers = {
  confirm: (id) => async () => {
    return { score: 100, detail: "Dikonfirmasi oleh peserta ✓" };
  },

  _isRed: (colorStr) => {
    const c = String(colorStr || "").toLowerCase().replace("#", "").trim();
    if (c === "red" || c === "ff0000" || c === "ffff0000") return true;
    if (c.length === 6) {
      const r = parseInt(c.substring(0, 2), 16);
      const g = parseInt(c.substring(2, 4), 16);
      const b = parseInt(c.substring(4, 6), 16);
      return r > 200 && g < 100 && b < 100;
    }
    if (c.length === 8) {
      const r = parseInt(c.substring(2, 4), 16);
      const g = parseInt(c.substring(4, 6), 16);
      const b = parseInt(c.substring(6, 8), 16);
      return r > 200 && g < 100 && b < 100;
    }
    return false;
  },

  _getSlideTexts: async (slide, context) => {
    try {
      slide.load("shapes");
      await context.sync();
      
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();
      
      if (!shapes.items || shapes.items.length === 0) return [];
      
      for (let shape of shapes.items) {
        shape.load("type");
      }
      await context.sync();
      
      const textShapes = [];
      for (let shape of shapes.items) {
        const t = String(shape.type).toLowerCase();
        if (t.includes('image') || t.includes('line') || t.includes('media') || t.includes('table') || t.includes('group') || t.includes('unsupported') || t === '3' || t === '4') {
          continue;
        }
        textShapes.push(shape);
      }
      
      if (textShapes.length === 0) return [];
      
      try {
        for (let shape of textShapes) {
          shape.load("textFrame/hasText");
        }
        await context.sync();
        
        const textRanges = [];
        for (let shape of textShapes) {
          if (shape.textFrame && shape.textFrame.hasText) {
            shape.textFrame.textRange.load("text");
            textRanges.push(shape.textFrame.textRange);
          }
        }
        if (textRanges.length > 0) {
          await context.sync();
          return textRanges.map(tr => tr.text || "");
        }
        return [];
      } catch (batchErr) {
        console.warn("Batch text load failed, falling back to shape-by-shape:", batchErr);
        const texts = [];
        for (let shape of textShapes) {
          try {
            shape.load("textFrame/hasText");
            await context.sync();
            if (shape.textFrame && shape.textFrame.hasText) {
              shape.textFrame.textRange.load("text");
              await context.sync();
              if (shape.textFrame.textRange.text) {
                texts.push(shape.textFrame.textRange.text);
              }
            }
          } catch (singleErr) {
            console.warn("Failed loading text for single shape:", singleErr);
          }
        }
        return texts;
      }
    } catch (e) {
      console.warn("Failed safely fetching slide texts:", e);
      return [];
    }
  }
};

window.CheckerHelpers = CheckerHelpers;
