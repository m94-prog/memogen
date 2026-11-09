/* MemoPDF: draws a MemoFormat layout with pdf-lib. Same positions as the preview. For a digitally signed memo it also adds what
 * AR 25-50 appendix F has you build by hand in Adobe: a date text box across from the office symbol and a digital signature box
 * above the signer's name. A receipt block gets its own pair. All of them have to exist before the first signature (app F-2h). */
(function (global) {
  const { layout, FONTS, PAGE_W, PAGE_H, BOX, BOX_GLYPH } = MemoFormat;
  const NBSP = / /g;

  async function exportPDF(memo, office, fmt) {
    const { PDFDocument, PDFName, PDFString, TextAlignment, rgb } = PDFLib;
    const L = layout(memo, office, fmt);
    const doc = await PDFDocument.create();
    doc.setTitle(MemoFormat.plain(memo.subject) || 'Memorandum');
    doc.setCreator('MemoGen'); doc.setProducer('MemoGen (pdf-lib)');
    doc.setCreationDate(new Date()); doc.setModificationDate(new Date());

    const embedded = new Map();
    const fontFor = async (key, fr) => {
      const F = FONTS[key] || FONTS.arial, n = F[fr.b && fr.i ? 'bi' : fr.b ? 'b' : fr.i ? 'i' : 'r'];
      if (!embedded.has(n)) embedded.set(n, await doc.embedFont(n));
      return { pdf: embedded.get(n), F };
    };
    let seal = null;
    const sealFor = async (src) => {
      if (seal) return seal;
      const T = src.startsWith('data:') ? null : await MemoDocx.loadTemplate();                         // 'seal.png': the template's own seal
      const b64 = T ? T.parts[T.sealPart] : src.split(',')[1];
      seal = await doc.embedPng(Uint8Array.from(atob(b64), (c) => c.charCodeAt(0)));                    // uploads are converted to PNG in Setup
      return seal;
    };
    const boxes = [], choices = [], blanks = [];
    // The square of Word's checkbox glyph, outline inside its edges. A digitally signed memo gets a real checkbox over it, and a
    // text field over each blank of underscores in the body, so the recipient can tick or initial in Adobe before signing; on paper
    // both are marked by pen.
    const checkbox = (page, x, base, s) => {
      const bw = Math.max(0.5, s * 0.05), x0 = x + BOX_GLYPH.x0 * s, side = (BOX_GLYPH.x1 - BOX_GLYPH.x0) * s, h = BOX_GLYPH.top * s;
      page.drawRectangle({ x: x0 + bw / 2, y: base + bw / 2, width: side - bw, height: h - bw, borderWidth: bw, borderColor: rgb(0, 0, 0) });
      if (memo.digital) choices.push({ page, x: x0 + bw, y: base + bw, w: side - 2 * bw, h: h - 2 * bw });
    };

    for (const pg of L.pages) {
      const p = doc.addPage([PAGE_W, PAGE_H]);
      for (const it of pg.items) {
        if (it.k === 't') {
          let x = it.x;
          const base = PAGE_H - (it.y + (it.lh - it.s * 1.117) / 2 + it.s * 0.905);
          for (const fr of it.fr) {
            const { pdf, F } = await fontFor(it.f, fr);
            for (const t of fr.t.replace(NBSP, ' ').split(/(☐)/)) {   // a checkbox is drawn, not typed: the standard fonts have no box
              if (t === BOX) checkbox(p, x, base, it.s);
              else if (t.trim()) {
                const lead = t.length - t.trimStart().length;   // after a box: start at the letter, as Word does
                p.drawText(t.slice(lead), { x: x + MemoFormat.widthOf(F, it.s, t.slice(0, lead), fr), y: base, size: it.s, font: pdf, color: rgb(0, 0, 0) });
                if (it.body && memo.digital) for (const m of t.matchAll(/_{3,}/g)) {   // a blank in the text, such as one for initials
                  blanks.push({ page: p, x: x + MemoFormat.widthOf(F, it.s, t.slice(0, m.index), fr), y: base - 0.07 * it.s, w: MemoFormat.widthOf(F, it.s, m[0], fr), h: it.lh });
                }
              }
              x += MemoFormat.widthOf(F, it.s, t, fr);
            }
          }
        } else if (it.k === 'img') {
          const img = await sealFor(it.src);
          p.drawImage(img, { x: it.x, y: PAGE_H - it.y - it.h, width: it.w, height: it.h });
        } else if (it.k === 'rule') {   // roster borders, 0.5 point (Word's default table border)
          p.drawLine({ start: { x: it.x1, y: PAGE_H - it.y1 }, end: { x: it.x2, y: PAGE_H - it.y2 }, thickness: 0.5, color: rgb(0, 0, 0) });
        } else if (it.k === 'sig' || it.k === 'date') boxes.push({ ...it, page: p });
      }
    }

    if (boxes.length || choices.length || blanks.length) {
      const form = doc.getForm(), helv = await doc.embedFont('Helvetica');
      const names = choices.map((c, i) => {
        const name = `Choice ${i + 1}`, f = form.createCheckBox(name);
        f.addToPage(c.page, { x: c.x, y: c.y, width: c.w, height: c.h, borderWidth: 0, borderColor: undefined, backgroundColor: undefined, textColor: rgb(0, 0, 0) });
        return name;
      });
      blanks.forEach((b, i) => {
        const name = `Blank ${i + 1}`, f = form.createTextField(name);
        f.setAlignment(TextAlignment.Center);
        f.addToPage(b.page, { x: b.x, y: b.y, width: b.w, height: b.h, borderWidth: 0, borderColor: undefined, backgroundColor: undefined, font: helv, textColor: rgb(0, 0, 0) });
        f.setFontSize(L.size);
        names.push(name);
      });
      // The recipient's signature locks the boxes they marked and the blanks they filled, along with their date.
      boxes.forEach((b) => { if (b.name === 'ReceiptSignature' && names.length) b.lock = [...(b.lock || []), ...names]; });
      for (const b of boxes) {
        if (b.k === 'date') {   // app F-2e: single line, filled in by the signer before signing
          const f = form.createTextField(b.name);
          f.setAlignment(b.align === 'center' ? TextAlignment.Center : TextAlignment.Right);
          f.addToPage(b.page, { x: b.x, y: PAGE_H - b.y - b.h, width: b.w, height: b.h, borderWidth: 0, font: helv, textColor: rgb(0, 0, 0), backgroundColor: undefined });
          f.setFontSize(L.size);
        } else {                // pdf-lib has no signature field API; this is the plain /Sig widget Adobe creates
          const ctx = doc.context, x = b.x, y1 = PAGE_H - b.y - b.h;
          const dict = ctx.obj({ Type: 'Annot', Subtype: 'Widget', FT: 'Sig', F: 4, Rect: [x, y1, x + b.w, y1 + b.h] });
          dict.set(PDFName.of('T'), PDFString.of(b.name));
          dict.set(PDFName.of('P'), b.page.ref);
          if (b.lock?.length) { // app F-2f "Mark as read-only: just these fields": signing locks this signer's dates and leaves the other signatures open
            const lock = ctx.obj({ Type: 'SigFieldLock', Action: 'Include' });
            lock.set(PDFName.of('Fields'), ctx.obj(b.lock.map((n) => PDFString.of(n))));
            dict.set(PDFName.of('Lock'), ctx.register(lock));
          }
          const ref = ctx.register(dict);
          b.page.node.addAnnot(ref);
          form.acroForm.addField(ref);
        }
      }
    }
    return { bytes: await doc.save(), pages: L.pages.length };
  }

  global.MemoPDF = { exportPDF };
})(window);
