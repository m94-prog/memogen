/* MemoDocx: fills the real template (Docs/MEMO TEMPLATE.docx, packed into memo-template-data.js by tools/make_template.py).
 * Kept from the template: its styles, the ArmyMemo_25-50 multilevel list, the letterhead with the seal, the bound Office Symbol,
 * Subject and Name fields on page 1, and everything else in the package. Where the template and AR 25-50 differ, AR 25-50 wins:
 *   - font size follows Setup (12 point recommended, 1-19; the template is 11)
 *   - signature on the fifth line below the text (2-4c(2)(a); the template leaves one blank line too many)
 *   - continuation header 1 inch from the top (2-5a) and page number about 1 inch from the bottom (2-5d)
 *   - 1/4 inch after the parenthesis of (1) and (a) (1-39b(10); the template uses two spaces)
 *   - a fixed date (2-4a(3); the template's DATE field changes every time the file is opened)
 *   - MFR on plain paper unless the memo says otherwise (fig 2-17)
 * Line counts from the layout engine set Word's keep-together flags, so Word obeys para 2-5c if it reflows. */
(function (global) {
  const F = () => global.MemoFormat;
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const tw = (pt) => Math.round(pt * 20);
  const CORE_STORE = '{6C3C8BC8-F283-45AE-878A-BAB7291924A1}';
  const CORE_MAP = "xmlns:ns0='http://purl.org/dc/elements/1.1/' xmlns:ns1='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' ";
  const XP = { symbol: '/ns1:coreProperties[1]/ns1:contentStatus[1]', subject: '/ns1:coreProperties[1]/ns0:subject[1]', name: '/ns1:coreProperties[1]/ns0:description[1]' };
  const LIST_ID = 33;   // the template's ArmyMemo_25-50 list
  const SYMBOL_STYLE = 'MemoOfficeSymbol', SYMBOL_STYLE_NAME = 'Memo Office Symbol', SUBJECT_STYLE = 'MemoSubjectLine', SUBJECT_STYLE_NAME = 'Memo Subject Line';

  async function loadTemplate() {
    if (global.MemoTemplate) return global.MemoTemplate;
    await new Promise((ok, no) => {
      const s = document.createElement('script'); s.src = 'memo-template-data.js'; s.onload = ok;
      s.onerror = () => no(new Error('memo-template-data.js is missing. Run tools/make_template.py.'));
      document.head.append(s);
    });
    return global.MemoTemplate;
  }
  const unb64 = (b) => Uint8Array.from(atob(b), (c) => c.charCodeAt(0));
  const dec = new TextDecoder(), enc = new TextEncoder();

  /* ---- a stored (uncompressed) ZIP: all a .docx needs, and the seal is the only large part and is already compressed ---- */
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = (b) => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  function zip(files) {
    const chunks = [], central = []; let off = 0;
    const d = new Date(), time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1), date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    for (const [name, data] of files) {
      const nm = enc.encode(name), crc = crc32(data), h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 0, true);
      h.setUint16(10, time, true); h.setUint16(12, date, true); h.setUint32(14, crc, true); h.setUint32(18, data.length, true); h.setUint32(22, data.length, true); h.setUint16(26, nm.length, true);
      chunks.push(new Uint8Array(h.buffer), nm, data);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true);
      c.setUint16(12, time, true); c.setUint16(14, date, true); c.setUint32(16, crc, true); c.setUint32(20, data.length, true); c.setUint32(24, data.length, true); c.setUint16(28, nm.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), nm);
      off += 30 + nm.length + data.length;
    }
    const cdSize = central.reduce((a, x) => a + x.length, 0), e = new DataView(new ArrayBuffer(22));
    e.setUint32(0, 0x06054b50, true); e.setUint16(8, files.length, true); e.setUint16(10, files.length, true); e.setUint32(12, cdSize, true); e.setUint32(16, off, true);
    const all = [...chunks, ...central, new Uint8Array(e.buffer)], out = new Uint8Array(all.reduce((a, x) => a + x.length, 0));
    let p = 0; for (const x of all) { out.set(x, p); p += x.length; }
    return out;
  }

  async function exportDocx(memo, office, fmtIn) {
    const MF = F(), T = await loadTemplate();
    const fmt = { font: 'arial', size: 12, spaces: 2, ...(fmtIn || {}) };
    const L = MF.layout(memo, office, fmt), LH = L.lh, S = L.size;
    const fontDef = MF.FONTS[fmt.font] || MF.FONTS.arial, font = fontDef.label, hp = Math.round(S * 2), lineTw = tw(LH);
    const part = (n) => dec.decode(unb64(T.parts[n]));
    const parts = new Map(Object.entries(T.parts).map(([k, v]) => [k, null]));
    const set = (n, s) => parts.set(n, enc.encode(s));

    /* building blocks; element order inside pPr and rPr follows the schema */
    const rp = (o = {}) => `<w:rPr>${o.rs ? `<w:rStyle w:val="${o.rs}"/>` : ''}<w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:cs="${font}"/>${o.b ? '<w:b/><w:bCs/>' : ''}${o.i ? '<w:i/><w:iCs/>' : ''}<w:sz w:val="${o.sz || hp}"/><w:szCs w:val="${o.sz || hp}"/></w:rPr>`;
    const textRun = (t, o) => `<w:r>${rp(o)}<w:t xml:space="preserve">${esc(t)}</w:t></w:r>`;
    let sdtId = 1000;
    // "[ ]" in the text: Word's own checkbox control, drawn in Wingdings (168 empty, 253 marked with an X), which the layout measures.
    const wing = (o = {}) => `<w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:cs="Wingdings"/>${o.b ? '<w:b/><w:bCs/>' : ''}<w:sz w:val="${o.sz || hp}"/><w:szCs w:val="${o.sz || hp}"/></w:rPr>`;
    const checkbox = (o) => `<w:sdt><w:sdtPr>${wing(o)}<w:alias w:val="Choice"/><w:id w:val="${(sdtId += 7)}"/><w14:checkbox><w14:checked w14:val="0"/><w14:checkedState w14:val="F0FD" w14:font="Wingdings"/><w14:uncheckedState w14:val="F0A8" w14:font="Wingdings"/></w14:checkbox></w:sdtPr><w:sdtContent><w:r>${wing(o)}<w:t>&#xF0A8;</w:t></w:r></w:sdtContent></w:sdt>`;
    const run = (t, o) => String(t).split(/(☐)/).map((s) => (s === MF.BOX ? checkbox(o) : s ? textRun(s, o) : '')).join('');
    const frags = (fr) => fr.map((f) => run(f.t, f)).join('');
    const tabRun = `<w:r>${rp()}<w:tab/></w:r>`;
    const ppr = (o = {}) => `<w:pPr>${o.style ? `<w:pStyle w:val="${o.style}"/>` : ''}${o.keepNext ? '<w:keepNext/>' : ''}${o.keepLines ? '<w:keepLines/>' : ''}${o.num ? `<w:numPr><w:ilvl w:val="${o.num[0]}"/><w:numId w:val="${o.num[1]}"/></w:numPr>` : ''}${o.tabs ? `<w:tabs>${o.tabs.map(([v, p]) => `<w:tab w:val="${v}" w:pos="${p}"/>`).join('')}</w:tabs>` : ''}${o.spacing || ''}${o.ind || ''}${o.jc ? `<w:jc w:val="${o.jc}"/>` : ''}${rp()}</w:pPr>`;
    const para = (inner = '', o) => `<w:p>${ppr(o)}${inner}</w:p>`;
    const blank = (o) => para('', o);
    const exact = (pt) => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${Math.max(tw(pt), 14)}" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr></w:pPr></w:p>`;
    const bound = (alias, xpath, text, o) => `<w:sdt><w:sdtPr>${rp(o)}<w:alias w:val="${alias}"/><w:tag w:val=""/><w:id w:val="${(sdtId += 7)}"/><w:dataBinding w:prefixMappings="${CORE_MAP}" w:xpath="${xpath}" w:storeItemID="${CORE_STORE}"/><w:text/></w:sdtPr><w:sdtContent>${run(text, o)}</w:sdtContent></w:sdt>`;
    const field = (instr, text) => `<w:r>${rp()}<w:fldChar w:fldCharType="begin"/></w:r><w:r>${rp()}<w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r><w:r>${rp()}<w:fldChar w:fldCharType="separate"/></w:r>${run(text)}<w:r>${rp()}<w:fldChar w:fldCharType="end"/></w:r>`;

    /* the values the bound fields show (Word refreshes them from the document properties when it opens the file) */
    const symbol = MF.clean(memo.officeSymbol || '').trim();
    const recordTail = memo.record ? ` (${MF.clean(memo.record).trim()})` : '';
    const subject = MF.plain(MF.printable(memo.subject, fmt)), name = MF.clean(memo.subjectName || '').trim();
    const symbolRuns = () => `${symbol ? bound('Office Symbol', XP.symbol, symbol, { rs: SYMBOL_STYLE }) : ''}${recordTail ? run(recordTail, { rs: SYMBOL_STYLE }) : ''}`;
    const smark = MF.subjectMark(memo);   // a CUI portion mark, when portion marking is on
    const subjectRuns = () => `${run('SUBJECT: ' + smark)}${subject ? bound('Subject', XP.subject, subject) : ''}${name ? run(' ') + bound('Name', XP.name, name) : ''}`;

    /* CUI banner (DoDI 5200.48 3.4a): a framed paragraph placed on the page itself, 1/4 inch from the edge, so it moves nothing else */
    const cui = L.cui;
    const banner = (top) => (cui ? `<w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:xAlign="center" w:y="${tw(top ? cui.edge : MF.PAGE_H - cui.edge - LH)}"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="center"/>${rp({ b: true })}</w:pPr>${run(cui.banner, { b: true })}</w:p>` : '');

    /* heading */
    const body = [];
    const dateText = memo.date ? MF.clean(MF.fmtDate(memo.date, memo.dateStyle)) : '';
    body.push(para(`${symbolRuns()}${dateText ? tabRun + run(dateText) : ''}`, { tabs: [['right', 9360]] }));   // 2-4a(3)(b): same line, flush right
    body.push(blank(), blank());
    const inlines = (s) => MF.parseInline(MF.printable(s, fmt)).map((r) => run(r.t, r)).join('');
    const hanging = (list) => list.forEach((a) => body.push(para(inlines(a), { ind: '<w:ind w:left="360" w:hanging="360"/>' })));
    const to = MF.splitLines(memo.to), thru = MF.splitLines(memo.thru);
    if (memo.kind === 'mfr') body.push(para(run('MEMORANDUM FOR RECORD')));
    else if (memo.kind === 'dist') body.push(para(run('MEMORANDUM FOR SEE DISTRIBUTION')));
    else if (memo.kind === 'exclusive') body.push(para(inlines(MF.exclusiveLine(memo))));
    else if (memo.kind === 'multi') { body.push(para(run('MEMORANDUM FOR')), blank()); hanging(to); }
    else if (memo.kind === 'thru') {
      if (thru.length > 1) { body.push(para(run('MEMORANDUM THRU')), blank()); hanging(thru); } else body.push(para(inlines(`MEMORANDUM THRU ${thru[0] || ''}`)));
      body.push(blank(), para(inlines(`FOR ${to.join(' ')}`)));
    } else body.push(para(inlines(`MEMORANDUM FOR ${to.join(' ')}`)));
    body.push(blank(), para(subjectRuns(), { style: SUBJECT_STYLE }), blank(), blank());

    /* body: the template's list, so Word keeps numbering it the way the instructions describe */
    const list = MF.markedBody(memo), byId = new Map(list.map((p) => [p.id, p]));
    const spacing = `<w:spacing w:before="0" w:after="${lineTw}"/>`, afterTable = `<w:spacing w:before="${lineTw}" w:after="${lineTw}"/>`;
    const lineCount = new Map(L.paraLines.map((p) => [p.id, p.lines]));
    const tables = new Map(L.rosters.map((t) => [t.id, t]));

    /* a roster as a real Word table: the layout's column widths, rows that never split, a heading row that repeats, and keep-with-next
     * on the same rows the layout keeps together, so Word breaks the pages where the PDF does */
    const BORDER = '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>';
    function table(t, isLast) {
      const n = t.cells.length - 1;
      const keep = (i) => (i === 0 && n >= 1) || (n >= 2 && (i === 1 || i === n - 1)) || (i === n && isLast);
      const cellP = (s, b, k) => `<w:p><w:pPr>${k ? '<w:keepNext/>' : ''}<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/>${rp({ b })}</w:pPr>${s ? run(s, { b }) : ''}</w:p>`;
      const rows = t.cells.map((cells, i) => `<w:tr><w:trPr><w:cantSplit/>${i === 0 ? '<w:tblHeader/>' : ''}</w:trPr>${cells.map((s, c) => `<w:tc><w:tcPr><w:tcW w:w="${tw(t.widths[c])}" w:type="dxa"/></w:tcPr>${cellP(s, i === 0, keep(i))}</w:tc>`).join('')}</w:tr>`).join('');
      const total = t.widths.reduce((a, x) => a + tw(x), 0);
      return `<w:tbl><w:tblPr><w:tblW w:w="${total}" w:type="dxa"/><w:tblInd w:w="0" w:type="dxa"/><w:tblBorders>${BORDER}</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="20" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="20" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000" w:firstRow="0" w:lastRow="0" w:firstColumn="0" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr><w:tblGrid>${t.widths.map((x) => `<w:gridCol w:w="${tw(x)}"/>`).join('')}</w:tblGrid>${rows}</w:tbl>`;
    }

    L.body.forEach((b, i) => {
      const last = i === L.body.length - 1, prev = L.body[i - 1], next = L.body[i + 1];
      if (b.k === 'tab') {
        if (prev?.k === 'tab') body.push(blank({ keepNext: true }));   // Word joins two tables that touch; the layout counts this line
        body.push(table(tables.get(b.id), last));
        return;
      }
      const p = byId.get(b.id);
      // A paragraph after a roster takes its blank line before it. Word drops that space at the top of a page, as the layout does.
      const o = { style: 'MemoBodyStyle', spacing: prev?.k === 'tab' ? afterTable : spacing, keepNext: last || next?.k === 'tab', keepLines: (lineCount.get(p.id) || 9) <= 3 };   // 2-5c(1), 2-5c(4)
      if (p.label) o.num = [p.level, LIST_ID];
      else { o.num = [0, 0]; o.ind = '<w:ind w:left="0" w:firstLine="0"/>'; }
      const inner = MF.printable(p.text, fmt).split('\n').map((seg) => MF.parseInline(seg).map((r) => run(r.t, r)).join('')).join(`<w:r>${rp()}<w:br/></w:r>`);
      body.push(para(inner, o));
    });

    /* closing: the layout's grid, row by row, each cell on its own tab stop. Row 1 is the last paragraph's spacing, or a blank line
     * when a roster ends the text. */
    if (L.body[L.body.length - 1]?.k === 'tab') body.push(blank({ keepNext: true }));
    const { grid, lastRow } = L.closing;
    for (let r = 2; r <= lastRow; r++) {
      const tabs = [], inner = [];
      [...(grid[r] || [])].sort((a, b) => a.x - b.x).forEach((c) => {
        if (c.x > 72.01) { tabs.push(['left', tw(c.x - 72)]); inner.push(tabRun); }
        inner.push(frags(c.fr));
      });
      body.push(para(inner.join(''), { tabs: tabs.length ? tabs : undefined, keepNext: r < lastRow }));
    }
    if (memo.kind === 'dist') {
      body.push(blank(), para(run('DISTRIBUTION:'), { keepNext: true }));
      String(memo.distribution ?? '').split(/\r?\n/).filter((x) => x.trim()).forEach((raw) => {
        const lead = raw.match(/^\s*/)[0].length, w = MF.widthOf(fontDef, S, ' ') * lead;
        body.push(para(inlines(raw), lead ? { ind: `<w:ind w:left="${tw(w)}"/>` } : undefined));
      });
    }
    const cf = MF.splitLines(memo.cf);
    if (cf.length) {
      body.push(blank(), para(run(memo.cfNoEncl ? 'CF: (w/o encls)' : 'CF:'), { keepNext: true }));
      cf.forEach((c) => body.push(para(inlines(c))));
    }

    /* word/document.xml: the template's own wrapper and section, a new body */
    const doc = part('word/document.xml');
    const head = doc.slice(0, doc.indexOf('<w:body>') + 8);
    let sect = doc.slice(doc.lastIndexOf('<w:sectPr'), doc.lastIndexOf('</w:sectPr>') + 11).replace('w:footer="720"', 'w:footer="1440"');

    // New header and footer parts (CUI page 1, enclosures) are registered here and written to the package at the end.
    const hdrOpen = part('word/header1.xml'), ftrOpen = part('word/footer1.xml');
    const hOpen = hdrOpen.slice(0, hdrOpen.indexOf('>', hdrOpen.indexOf('<w:hdr')) + 1), fOpen = ftrOpen.slice(0, ftrOpen.indexOf('>', ftrOpen.indexOf('<w:ftr')) + 1);
    const hair = (inner = '') => `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="14" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/><w:szCs w:val="2"/></w:rPr>${inner}</w:pPr></w:p>`;
    const rels = [], types = [];
    const addPart = (name, kind, xml) => {
      const id = `rIdMg${rels.length + 1}`;
      set(`word/${name}`, xml);
      rels.push(`<Relationship Id="${id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${name}"/>`);
      types.push(`<Override PartName="/word/${name}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/>`);
      return id;
    };

    /* CUI page 1: its own footer holds the designation indicator (DoDI 5200.48 fig 2), lower right under a blank line, and a cover
     * memo's statement at the left margin. Each line is the layout's line, so Word cannot wrap them differently. */
    if (cui) {
      const dsz = cui.size * 2;
      const dl = (t, ind, first) => `<w:p><w:pPr><w:spacing w:before="${first ? lineTw : 0}" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="${ind}" w:right="0" w:firstLine="0"/>${rp({ sz: dsz })}</w:pPr>${run(t, { sz: dsz })}</w:p>`;
      const lines = [...cui.cover.map((t) => [t, 0]), ...cui.di.map((t) => [t, tw(cui.diX - 72)])];
      const id = addPart('footer-cui-first.xml', 'footer', `${fOpen}${lines.map(([t, ind], i) => dl(t, ind, i === 0)).join('')}${banner(false)}</w:ftr>`);
      sect = sect.replace('<w:type ', `<w:footerReference w:type="first" r:id="${id}"/><w:type `);
    }

    /* enclosure rosters: each is its own section on plain paper, so it gets no letterhead, memo header or memo page number. Its
     * first-page footer carries "Encl" and the number in the lower right corner (1-34, 4-2d(1)), where the page number sits on
     * the memo. Every header and footer is named, since a section that names none inherits the memo's. */
    const encls = L.rosters.filter((t) => t.place === 'encl');
    if (encls.length) {
      const hDefault = addPart('header-encl-a.xml', 'header', `${hOpen}${hair()}${banner(true)}</w:hdr>`);
      const hFirst = addPart('header-encl-b.xml', 'header', `${hOpen}${hair()}${banner(true)}</w:hdr>`);
      const fDefault = addPart('footer-encl.xml', 'footer', `${fOpen}${hair()}${banner(false)}</w:ftr>`);
      const enclSect = (fFirst) => `<w:sectPr><w:headerReference w:type="default" r:id="${hDefault}"/><w:footerReference w:type="default" r:id="${fDefault}"/><w:headerReference w:type="first" r:id="${hFirst}"/><w:footerReference w:type="first" r:id="${fFirst}"/><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="1440" w:gutter="0"/><w:cols w:space="720"/><w:titlePg/><w:docGrid w:linePitch="360"/></w:sectPr>`;
      // The memo's own section ends on its last paragraph, so no extra line is added to the memo.
      body[body.length - 1] = body[body.length - 1].replace('</w:pPr>', `${sect}</w:pPr>`);
      encls.forEach((t, k) => {
        const fFirst = addPart(`footer-encl-${k + 1}.xml`, 'footer', `${fOpen}<w:p>${ppr({ style: 'Footer', spacing: `<w:spacing w:before="${lineTw}"/>`, jc: 'right' })}${run(t.mark)}</w:p>${banner(false)}</w:ftr>`);
        body.push(para(run(t.title, { b: true }), { keepNext: true, spacing: `<w:spacing w:before="0" w:after="${lineTw}"/>`, jc: 'center' }));
        body.push(table(t, false));
        if (k < encls.length - 1) body.push(hair(enclSect(fFirst)));
        else { body.push(hair()); sect = enclSect(fFirst); }
      });
    }
    if (rels.length) {
      set('word/_rels/document.xml.rels', part('word/_rels/document.xml.rels').replace('</Relationships>', `${rels.join('')}</Relationships>`));
      set('[Content_Types].xml', part('[Content_Types].xml').replace('</Types>', `${types.join('')}</Types>`));
    }
    set('word/document.xml', `${head}${body.join('')}${sect}</w:body></w:document>`);

    /* continuation header: office symbol 1 inch from the top, subject on the next line, text on the third line below (2-5).
     * STYLEREF fields, not the template's content controls: any content control in this header makes Word fail when it exports a
     * two-page memo to PDF (the untouched template does it too), and that export is step one of a digital signature (app F-2b).
     * STYLEREF still copies the text from page 1, so a subject changed in Word follows onto the later pages. */
    const hdr1 = part('word/header1.xml'), open1 = hdr1.slice(0, hdr1.indexOf('>', hdr1.indexOf('<w:hdr')) + 1);
    const symbolText = `${symbol}${recordTail}`, subjectText = `SUBJECT: ${smark}${[subject, name].filter(Boolean).join(' ')}`;
    set('word/header1.xml', `${open1}${exact(36)}${para(symbolText ? field(`STYLEREF "${SYMBOL_STYLE_NAME}"`, symbolText) : '')}${para(field(`STYLEREF "${SUBJECT_STYLE_NAME}"`, subjectText))}${blank()}${blank()}${banner(true)}</w:hdr>`);

    /* first-page header: the template's letterhead paragraph with the seal, then this office's lines */
    const hdr2 = part('word/header2.xml'), open2 = hdr2.slice(0, hdr2.indexOf('>', hdr2.indexOf('<w:hdr')) + 1);
    const lh = MF.LETTERHEAD, top = L.officeTop;
    const suspense = memo.suspense ? MF.clean(`S: ${MF.fmtDate(memo.suspense, memo.dateStyle)}`) : '';
    let first = '', used = lh.top;
    if (MF.usesLetterhead(memo)) {
      const p0 = hdr2.slice(hdr2.indexOf('<w:p '), hdr2.indexOf('</w:p>') + 6);
      const headline = String(office?.headline ?? 'DEPARTMENT OF THE ARMY');
      // The template's 13 point setting sits on the seal's run, so its headline falls back to 8 point; the regulation's figures show it larger.
      first = p0.replace(/<w:r(?: [^>]*)?><w:t>Department of the Army<\/w:t><\/w:r>/, `<w:r><w:rPr><w:sz w:val="${lh.headSize * 2}"/><w:szCs w:val="${lh.headSize * 2}"/></w:rPr><w:t>${esc(MF.clean(headline))}</w:t></w:r>`);
      used += headline.trim() ? lh.headLH : 0;
      MF.splitLines(office?.lines).forEach((t) => { first += `<w:p><w:pPr><w:pStyle w:val="ALLCAPS"/></w:pPr><w:r><w:t xml:space="preserve">${esc(MF.clean(t))}</w:t></w:r></w:p>`; used += lh.lineLH; });
    }
    if (suspense) {   // 2-4a(4): two lines above the date, flush right, bold
      const sTop = top - 2 * LH;
      first += `${exact(Math.max(sTop - used, 1))}<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="${lineTw}" w:lineRule="exact"/><w:jc w:val="right"/>${rp({ b: true })}</w:pPr>${run(suspense, { b: true })}</w:p>${exact(LH)}`;
    } else first += exact(Math.max(top - used, 1));
    set('word/header2.xml', `${open2}${first}${banner(true)}</w:hdr>`);

    /* footer: PAGE field, about 1 inch from the bottom with a blank line above it */
    const ftr = part('word/footer1.xml'), open3 = ftr.slice(0, ftr.indexOf('>', ftr.indexOf('<w:ftr')) + 1);
    const fld = (x) => `<w:r>${rp()}${x}</w:r>`;
    set('word/footer1.xml', `${open3}<w:p>${ppr({ style: 'Footer', spacing: `<w:spacing w:before="${lineTw}"/>`, jc: 'center' })}${fld('<w:fldChar w:fldCharType="begin"/>')}${fld('<w:instrText xml:space="preserve"> PAGE </w:instrText>')}${fld('<w:fldChar w:fldCharType="separate"/>')}${fld('<w:t>2</w:t>')}${fld('<w:fldChar w:fldCharType="end"/>')}</w:p>${banner(false)}</w:ftr>`);

    /* styles: the chosen size, so text typed in Word later matches */
    let st = part('word/styles.xml');
    for (const id of ['MemoBodyStyle', 'MemoBodyStyleChar', 'MEMOLINE', 'MEMOLINEChar']) {
      st = st.replace(new RegExp(`(w:styleId="${id}">[\\s\\S]*?</w:style>)`), (b) => b.replace(/<w:sz w:val="\d+"\/><w:szCs w:val="\d+"\/>/g, `<w:sz w:val="${hp}"/><w:szCs w:val="${hp}"/>`));
    }
    // Two plain styles for the header's STYLEREF fields to find. They add no formatting of their own.
    st = st.replace('</w:styles>', `<w:style w:type="character" w:customStyle="1" w:styleId="${SYMBOL_STYLE}"><w:name w:val="${SYMBOL_STYLE_NAME}"/><w:uiPriority w:val="99"/></w:style><w:style w:type="paragraph" w:customStyle="1" w:styleId="${SUBJECT_STYLE}"><w:name w:val="${SUBJECT_STYLE_NAME}"/><w:basedOn w:val="Normal"/><w:uiPriority w:val="99"/></w:style></w:styles>`);
    set('word/styles.xml', st);

    /* list levels (1) and (a): 1-39b(10) puts 1/4 inch after the parenthesis, so a tab stop replaces the template's two spaces */
    const parenTab = 720 + tw(MF.widthOf(fontDef, S, '(1)') + 18);
    const num = part('word/numbering.xml');
    const a0 = num.indexOf('w:abstractNumId="3"'), a1 = num.indexOf('</w:abstractNum>', a0);
    const blk = num.slice(a0, a1).replace(/<w:suff w:val="space"\/><w:lvlText w:val="\(%([34])\) "\/><w:lvlJc w:val="left"\/><w:pPr><w:ind w:left="0" w:firstLine="720"\/><\/w:pPr>/g,
      (m, n) => `<w:suff w:val="tab"/><w:lvlText w:val="(%${n})"/><w:lvlJc w:val="left"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="${parenTab}"/></w:tabs><w:ind w:left="0" w:firstLine="720"/></w:pPr>`);
    set('word/numbering.xml', num.slice(0, a0) + blk + num.slice(a1));

    /* document properties feed the bound fields */
    const setProp = (xml, tag, v) => xml.replace(new RegExp(`<${tag}>[\\s\\S]*?</${tag}>|<${tag}/>`), `<${tag}>${esc(v)}</${tag}>`);
    let core = part('docProps/core.xml');
    const title = MF.plain(MF.subjectText(memo)) || 'Memorandum';
    core = setProp(core, 'dc:title', title); core = setProp(core, 'dc:subject', subject); core = setProp(core, 'dc:description', name); core = setProp(core, 'cp:contentStatus', symbol);
    core = core.replace(/(<dcterms:modified[^>]*>)[^<]*/, `$1${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}`);
    set('docProps/core.xml', core);
    set('docProps/app.xml', part('docProps/app.xml').replace(/<vt:lpstr>NGCA-INF-OP<\/vt:lpstr>/, `<vt:lpstr>${esc(title)}</vt:lpstr>`));

    /* the template points at a .dotx in one person's Downloads folder; a memo should not carry that path */
    set('word/settings.xml', part('word/settings.xml').replace(/<w:attachedTemplate [^>]*\/>/, ''));
    parts.delete('word/_rels/settings.xml.rels');

    /* a seal uploaded in Setup replaces the template's image */
    if (office?.seal && office.seal.startsWith('data:image/png')) parts.set('word/media/image1.png', unb64(office.seal.split(',')[1]));

    const files = [...parts].map(([n, v]) => [n, v || unb64(T.parts[n])]);
    files.sort((a, b) => (a[0] === '[Content_Types].xml' ? -1 : b[0] === '[Content_Types].xml' ? 1 : 0));
    return { bytes: zip(files), pages: L.pages.length };
  }

  global.MemoDocx = { exportDocx, loadTemplate };
})(window);
