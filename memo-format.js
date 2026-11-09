/* MemoFormat: lays a memorandum out to AR 25-50 (10 October 2020, administrative revision 4 October 2024).
 * layout() returns positioned lines, in points, origin at the top left. The on-screen preview and the PDF both draw that one
 * result, so they cannot disagree. Widths come from the PDF standard fonts: Helvetica has the same widths as Arial and Times the
 * same as Times New Roman, so a line that breaks here breaks the same way in the printed Arial.
 * Paragraph numbers in comments are AR 25-50 paragraphs. */
(function (global) {
  const PAGE_W = 612, PAGE_H = 792, IN = 72, QTR = 18;
  const LEFT = IN, RIGHT = PAGE_W - IN, TEXT_W = RIGHT - LEFT;
  const BOTTOM = PAGE_H - IN;          // 2-3c: 1-inch bottom margin
  const SIG_X = PAGE_W / 2;            // 2-4c(2)(a): the signature block begins in the center of the page
  const SEAL = { x: 33.6, y: 36, size: 78.4 };   // where Docs/MEMO TEMPLATE.docx places it
  const LETTERHEAD = { top: 36, headSize: 12, headLH: 13.8, lineSize: 8, lineLH: 9.2 };   // figs 2-1 to 2-14: the headline is larger than the unit lines
  const SEAL_BOTTOM = SEAL.y + SEAL.size;
  // Roster cells: Word's default 0.075 inch at the sides, 1 point above and below the text. Word makes each row the text plus
  // those margins plus its 0.5 point top border, and the table ends with one more border; measured in Word, not guessed.
  const PAD = 5.4, CELL_V = 1, BORDER = 0.5;
  // A checkbox, typed as "[ ]" in the text. Word draws it with its checkbox control in Wingdings (character 168, a square that is
  // 0.891 em wide and no taller than an Arial line), so the layout gives it that width; the PDF and preview draw the same square
  // (the glyph's outline: 0.0845 to 0.8071 em across, baseline to 0.7227 em up). Read from wingding.ttf, not guessed.
  const BOX = '☐', BOX_ADV = 891.1, BOX_GLYPH = { x0: 0.0845, x1: 0.8071, top: 0.7227 };
  const NBSP = ' ';

  const FONTS = {
    arial: { label: 'Arial', css: 'Arial, Helvetica, sans-serif', r: 'Helvetica', b: 'Helvetica-Bold', i: 'Helvetica-Oblique', bi: 'Helvetica-BoldOblique' },
    times: { label: 'Times New Roman', css: '"Times New Roman", Times, serif', r: 'Times-Roman', b: 'Times-Bold', i: 'Times-Italic', bi: 'Times-BoldItalic' },
    courier: { label: 'Courier New', css: '"Courier New", Courier, monospace', r: 'Courier', b: 'Courier-Bold', i: 'Courier-Oblique', bi: 'Courier-BoldOblique' },
  };
  const KINDS = {
    single: { label: 'Single address', hint: 'One addressee (para 2-4a(5)(a))' },
    multi: { label: 'Multiple address', hint: 'Two to five addressees (para 2-4a(5)(b))' },
    dist: { label: 'SEE DISTRIBUTION', hint: 'More than five addressees (para 2-4a(5)(c))' },
    thru: { label: 'MEMORANDUM THRU', hint: 'Routed through others first (para 2-4a(5)(d))' },
    exclusive: { label: 'Exclusive For', hint: 'A sensitive or privileged matter for one named person (para 1-12)' },
    mfr: { label: 'Memorandum for record', hint: 'Plain paper, no authority line (para 2-7, fig 2-17)' },
  };

  /* ---- glyph widths (unkerned: the PDF draws without kerning, and the preview turns kerning off) ---- */
  const embedders = {}, widths = {};
  const embedder = (n) => (embedders[n] ||= PDFLib.StandardFontEmbedder.for(n));
  function glyph(name, ch) {
    if (ch === BOX) return BOX_ADV;
    const t = (widths[name] ||= new Map());
    let w = t.get(ch);
    if (w === undefined) { try { w = embedder(name).widthOfTextAtSize(ch, 1000); } catch { w = null; } t.set(ch, w); }
    return w;
  }
  const SUBS = { '‐': '-', '‑': '-', '‒': '-', '−': '-', ' ': ' ', ' ': ' ', ' ': ' ', ' ': ' ', '​': '', '﻿': '', '\t': ' ' };
  // Only what the standard fonts can draw gets through; anything else shows as "?" in both the preview and the PDF.
  function clean(s) {
    let out = '';
    for (const c0 of String(s ?? '')) {
      const c = SUBS[c0] ?? c0;
      if (c === '' || c === NBSP || c === '\n') { out += c; continue; }
      out += glyph('Helvetica', c) === null ? '?' : c;
    }
    return out;
  }
  const fontName = (F, fr) => F[fr?.b && fr?.i ? 'bi' : fr?.b ? 'b' : fr?.i ? 'i' : 'r'];
  function widthOf(F, size, text, fr) {
    const n = fontName(F, fr); let w = 0;
    for (const c of text) w += glyph(n, c === NBSP ? ' ' : c) ?? 0;
    return w * size / 1000;
  }

  /* ---- text preparation ---- */
  // 1-39b(9): two spaces after periods and question marks; one after commas, colons and semicolons.
  const NOT_SENTENCE_END = /(?:^|[("'“‘])(?:Mr|Mrs|Ms|Dr|Sr|Jr|St|Mt|Ft|Rd|Ave|Blvd|No|Nos|para|paras|fig|figs|vol|vs|e\.g|i\.e|a\.m|p\.m|U\.S|Inc|Ltd|Co|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec|[A-Z])$/;
  function spaceSentences(s, n) {
    const gap = ' '.repeat(n);
    return s.replace(/ {2,}/g, ' ').replace(/(\S*?)([.?!])(["”’')\]*]*) (?=["“‘'(\[*]*[A-Z0-9☐]|_{3})/g, (m, word, p, close) => {
      if (p === '.' && NOT_SENTENCE_END.test(word)) return m;
      return word + p + close + gap;
    });
  }
  // 1-25c: do not separate the day, month and year of a date.
  const MONTHS = 'Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?';
  const DATE_RE = new RegExp(`\\b(\\d{1,2}) (${MONTHS}) (\\d{2,4})\\b`, 'g');
  const keepDates = (s) => s.replace(DATE_RE, `$1${NBSP}$2${NBSP}$3`);
  // "[ ]" becomes a checkbox. It and a blank of three or more underscores (for initials) stay with the word after them, so a line
  // never ends on a box or a blank.
  const boxes = (s) => String(s ?? '').replace(/\[ \]/g, BOX);
  function prep(text, fmt) { return keepDates(spaceSentences(clean(boxes(text)).trim(), fmt.spaces === 1 ? 1 : 2)).replace(/(☐|_{3,}) /g, `$1${NBSP}`); }

  // Inline emphasis (1-32): **bold**, *italic*, ***both***. Anything unmatched stays as typed.
  function parseInline(s) {
    const runs = [], re = /(\*{1,3})(?=\S)(.+?)(?<=\S)\1/g; let last = 0, m;
    while ((m = re.exec(s))) {
      if (m.index > last) runs.push({ t: s.slice(last, m.index) });
      runs.push({ t: m[2], b: m[1].length >= 2, i: m[1].length !== 2 });
      last = re.lastIndex;
    }
    if (last < s.length) runs.push({ t: s.slice(last) });
    return runs;
  }
  const plain = (s) => parseInline(String(s ?? '')).map((r) => r.t).join('');

  function wordsOf(runs) {
    const words = []; let cur = null, gap = 0;
    for (const r of runs) {
      for (const p of r.t.split(/( +)/)) {
        if (!p) continue;
        if (p[0] === ' ') { if (cur) { words.push(cur); cur = null; } gap += p.length; continue; }
        if (!cur) { cur = { fr: [], gap }; gap = 0; }
        cur.fr.push({ t: p, b: !!r.b, i: !!r.i });
      }
    }
    if (cur) words.push(cur);
    return words;
  }
  // Greedy line breaking at spaces only: words are never hyphenated (2-5c(3)).
  function wrap(F, size, runs, firstW, restW) {
    const lines = []; let cur = { fr: [], w: 0 }, avail = firstW;
    const flush = () => { lines.push(cur); cur = { fr: [], w: 0 }; avail = restW; };
    const spaceW = widthOf(F, size, ' ');
    for (const wd of wordsOf(runs)) {
      const ww = wd.fr.reduce((a, f) => a + widthOf(F, size, f.t, f), 0);
      const gapW = cur.fr.length ? wd.gap * spaceW : 0;
      if (cur.fr.length && cur.w + gapW + ww > avail + 0.01) flush();
      if (!cur.fr.length && ww > avail + 0.01) {   // one word wider than the line: the only case that breaks inside a word
        for (const f of wd.fr) for (const c of f.t) {
          const cw = widthOf(F, size, c, f);
          if (cur.fr.length && cur.w + cw > avail + 0.01) flush();
          const lastFr = cur.fr[cur.fr.length - 1];
          if (lastFr && lastFr.b === f.b && lastFr.i === f.i) lastFr.t += c; else cur.fr.push({ t: c, b: f.b, i: f.i });
          cur.w += cw;
        }
        continue;
      }
      const lead = cur.fr.length ? ' '.repeat(wd.gap) : '';
      wd.fr.forEach((f, i) => cur.fr.push({ t: (i === 0 ? lead : '') + f.t, b: f.b, i: f.i }));
      cur.w += (cur.fr.length > wd.fr.length ? gapW : 0) + ww;
    }
    if (cur.fr.length || !lines.length) lines.push(cur);
    return lines;
  }

  /* ---- paragraph numbering (fig 2-1): 1. / a. / (1) / (a), and no number on a one-paragraph memorandum (2-4b(4)(a)) ---- */
  const alpha = (n) => { let s = ''; for (n -= 1; n >= 0; n = Math.floor(n / 26) - 1) s = String.fromCharCode(97 + (n % 26)) + s; return s; };
  function numbered(paras) {
    const list = (paras || []).filter((p) => String(p.text ?? '').trim());
    const tops = list.filter((p) => !p.level).length;
    const n = [0, 0, 0, 0]; let prev = -1;
    return list.map((p) => {
      const level = Math.max(0, Math.min(3, p.level || 0, prev + 1)); prev = level;
      n[level] += 1; n.fill(0, level + 1);
      const v = n[level];
      const label = level === 0 ? (tops > 1 ? `${v}.` : '') : level === 1 ? `${alpha(v)}.` : level === 2 ? `(${v})` : `(${alpha(v)})`;
      return { id: p.id, level, label, text: p.text };
    });
  }

  /* ---- dates (1-25a, 2-4a(3)(c)): 13 March 2020 or 13 Mar 20 ---- */
  const MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  function fmtDate(iso, style) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso ?? '').trim());
    if (!m) return String(iso ?? '').trim();
    const d = +m[3], mon = MONTH_NAMES[+m[2] - 1] || '';
    return style === 'short' ? `${d} ${mon.slice(0, 3)} ${m[1].slice(2)}` : `${d} ${mon} ${m[1]}`;
  }

  const splitLines = (s) => String(s ?? '').split(/\r?\n/).map((x) => x.replace(/\s+$/, '')).filter((x) => x.trim());
  const usesLetterhead = (memo) => memo.kind !== 'mfr' || !!memo.letterhead;   // fig 2-17: the MFR goes on plain white paper
  // 1-12b(1): name, title and mailing address, in the regulation's own capitalization.
  const exclusiveLine = (memo) => `Memorandum Exclusive For ${splitLines(memo.to).map((x) => x.trim().replace(/,$/, '')).join(', ')}`;
  const subjectText = (memo) => [memo.subject, memo.subjectName].map((x) => String(x ?? '').trim()).filter(Boolean).join(' ');
  // 1-23a, 2-4b(1)(e): the point of contact closes the body. The sentence is the one in Docs/MEMO TEMPLATE.docx.
  function pocText(poc) {
    const name = String(poc?.name ?? '').trim();
    const reach = [poc?.phone, poc?.email].map((x) => String(x ?? '').trim()).filter(Boolean).join(' or ');
    if (!name && !reach) return '';
    return `The point of contact for this memorandum is ${[name, reach].filter(Boolean).join(' at ')}.`;
  }
  // AR 25-50 has no receipt format. The sentence goes in the body because text after the signature would be a postscript (1-29).
  function receiptText(memo) {
    const r = memo.receipt; if (!r?.on) return '';
    const own = String(r.text ?? '').trim(); if (own) return own;
    const by = memo.suspense ? ` and returning this memorandum by ${fmtDate(memo.suspense, memo.dateStyle)}` : '';
    return `Acknowledge receipt by signing below${by}. Your signature confirms receipt only, not agreement.`;
  }
  // The POC stays last (2-4b(1)(e)), so the receipt sentence comes just before it.
  function bodyParas(memo) {
    const r = receiptText(memo), t = pocText(memo.poc);
    return [...(memo.paras || []), ...(r ? [{ id: 'receipt', level: 0, text: r }] : []), ...(t ? [{ id: 'poc', level: 0, text: t }] : [])];
  }
  const printable = (s, fmt) => prep(s, fmt || {});

  /* ---- rosters: a table in the body, or on its own pages as an enclosure. AR 25-50 has no table format, so where it is silent
   * the layout keeps the memo's own font and spacing and saves space (see layout()). Rows are memo text: never saved. ---- */
  const cellText = (s) => clean(s).replace(/\s+/g, ' ').trim();
  function rosterTable(r) {
    if (!r) return null;
    const cols = (r.cols || []).map(cellText);
    if (!cols.some(Boolean)) return null;
    const rows = (r.rows || []).map((row) => cols.map((_, i) => cellText(row?.[i] ?? ''))).filter((row) => row.some(Boolean));
    const numbered = r.numbered !== false;
    return {
      id: r.id, place: r.place === 'encl' ? 'encl' : 'body', after: r.after, title: rosterTitle(r), numbered,
      head: numbered ? ['No.', ...cols] : cols, rows: numbered ? rows.map((row, i) => [String(i + 1), ...row]) : rows,
    };
  }
  const rosterTitle = (r) => cellText(r?.title || '') || 'Roster';
  // 1-34: enclosures in the order the text mentions them. Typed ones come first, then enclosure rosters in their order.
  const enclosureRosters = (memo) => (memo.rosters || []).filter((r) => r.place === 'encl' && rosterTable(r));
  const enclosureItems = (memo) => [...splitLines(memo.encl?.items), ...enclosureRosters(memo).map(rosterTitle)];

  /* ---- CUI marking. AR 25-50 para 1-22d sends CUI to DoDI 5200.48; these follow it and the DoD CUI Marking aid (December 2024):
   *   - "CUI" at the top and bottom of every page, nothing else in the banner (DoDI 5200.48 3.4a, 3.4f)
   *   - the designation indicator on the first page, lower right (fig 2; dodcui.mil): Controlled by (DoD Component, left off on
   *     letterhead), Controlled by (office), CUI Category, the dissemination control, POC
   *   - portion marks optional but recommended; if used, every subject, title, paragraph and table gets (CUI) or (U) (3.4b), never
   *     the signature block or the designation indicator
   * Where the sources are silent, MemoGen saves space: the banner sits 1/4 inch from the edge, outside the AR 25-50 positions, and
   * the designation indicator is 10 point. Abbreviations are the DoD CUI Registry's (dodcui.mil, CUI Categories and Abbreviations). */
  const CUI_CATEGORIES = { ADJ: 'Status Adjustment', ADPO: 'Administrative Proceedings', AG: 'Agriculture', AIV: 'Accident Investigation', APP: 'Patent Applications', ARCHR: 'Archaeological Resources', ASYL: 'Asylee', BARG: 'Collective Bargaining', BATT: 'Battered Spouse or Child', BUDG: 'Budget', CEII: 'Critical Energy Infrastructure Information', CENS: 'U.S. Census', CHLD: 'Child Pornography', CHRI: 'Criminal History Records Information', CMPRS: 'Committed Person', COMPT: 'Comptroller General', CONREG: 'Entity Registration Information', CONTRACT: 'Contract Use', CONV: 'Tax Convention', CRIT: 'General Critical Infrastructure Information', CRITAN: 'Ammonium Nitrate', CTI: 'Controlled Technical Information', CVI: 'Chemical-Terrorism Vulnerability Information', CVIC: 'Child Victim/Witness', DCNI: 'Unclassified Controlled Nuclear Information - Defense', DCRIT: 'DoD Critical Infrastructure Security Information', DREC: 'Death Records', EMGT: 'Emergency Management', EXPT: 'Export Controlled', EXPTR: 'Export Controlled Research', FINT: 'International Financial Institutions', FISA: 'Foreign Intelligence Surveillance Act', FISAB: 'FISA Business Records', FNC: 'General Financial Information', FSEC: 'Bank Secrecy', FSI: 'Financial Supervision Information', FUND: 'Campaign Funds', GENETIC: 'Genetic Information', GEO: 'Geodetic Product Information', HISTP: 'Historic Properties', HLTH: 'Health Information', ID: 'Internal Data', IFNC: 'Intelligence Financial Records', INF: 'Informant', INTL: 'International Agreement Information', INTEL: 'General Intelligence', INV: 'Investigation', INVENT: 'Inventions', ISVI: 'Information Systems Vulnerability Information', IVIC: 'Victims of Human Trafficking', JURY: 'Federal Grand Jury', JUV: 'Juvenile', LCOMM: 'Communications', LDNA: 'DNA', LEI: 'General Law Enforcement', LFNC: 'Law Enforcement Financial Records', LMI: 'Legislative Materials', LNSL: 'National Security Letter', LPROT: 'Protective Order', LSCRN: 'Terrorist Screening', LVIC: 'Victim', MERG: 'Mergers', MFC: 'Proprietary Manufacturer', MIL: 'Military Personnel Records', NETW: 'Net Worth', NNPI: 'Naval Nuclear Propulsion Information', NPSR: 'National Park System Resources', NUC: 'General Nuclear', OCCMTO: 'Ocean Common Carrier and Marine Terminal Operator Agreements', OPSEC: 'Operations Security', PCII: 'Protected Critical Infrastructure Information', PERS: 'Personnel Records', PEST: 'Pesticide Producer Survey', PHYS: 'Physical Security', POST: 'Proprietary Postal', PRE: 'Presentence Report', PRIIG: 'Inspector General Protected', PRIOR: 'Prior Arrest', PRIVILEGE: 'Legal Privilege', PROCURE: 'General Procurement and Acquisition', PROPIN: 'General Proprietary Business Information', PROT: 'Temporary Protected Status', PRVCY: 'General Privacy', PSEC: 'Secrecy Orders', PSI: 'Privileged Safety Information', RAIL: 'Railroad Safety Analysis Records', RECOMM: 'Nuclear Recommendation Material', RESD: 'Permanent Resident Status', RTR: 'Retirement', RWRD: 'Reward', SAFE: 'SAFETY Act Information', SBIZ: 'Small Business Research and Technology', SCV: 'Sex Crime Victim', SERV: 'Ocean Common Carrier Service Contracts', SGI: 'Safeguards Information', SRI: 'Nuclear Security-Related Information', SSEL: 'Source Selection', SSI: 'Sensitive Security Information', STAT: 'Statistical Information', STUD: 'Student Records', SUB: 'Controlled Substances', SURV: 'Investment Survey', TAI: 'Taxpayer Advocate Information', TAX: 'Federal Taxpayer Information', TRACE: 'Pen Register and Trap & Trace', TSCA: 'Toxic Substances', UCNI: 'Unclassified Controlled Nuclear Information - Energy', VISA: 'Visas', WATER: 'Water Assessments', WDT: 'Written Determinations', WHISTL: 'Whistleblower Identity', WIT: 'Witness Protection', XFER: 'Electronic Funds Transfer' };
  // Limited dissemination controls (dodcui.mil). None means anyone with a lawful government purpose may have it, not public release.
  const CUI_LDCS = { '': 'None', 'FED ONLY': 'Federal employees only', FEDCON: 'Federal employees and contractors only', NOCON: 'No dissemination to contractors', 'DL ONLY': 'Dissemination list controlled', NOFORN: 'No foreign dissemination', 'REL TO': 'Authorized for release to certain foreign nationals only', 'DISPLAY ONLY': 'Display only', RELIDO: 'Releasable by information disclosure official', 'ATTORNEY-CLIENT': 'Attorney client', 'ATTORNEY-WP': 'Attorney work product' };
  const CUI_BANNER = 'CUI', BANNER_EDGE = 18, DI_SIZE = 10;
  const cuiOn = (memo) => !!memo?.cui?.on;
  const portionOn = (memo) => cuiOn(memo) && !!memo.cui.portion;
  const portionMark = (m) => `(${m === 'CUI' ? 'CUI' : 'U'}) `;
  // The body as it prints: numbered, then portion marked. The POC and receipt sentences are unclassified.
  const markedBody = (memo) => numbered(bodyParas(memo)).map((p) => (portionOn(memo) ? { ...p, text: portionMark(memo.paras?.find((q) => q.id === p.id)?.mark) + p.text } : p));
  const subjectMark = (memo) => (portionOn(memo) ? portionMark(memo.cui.subject) : '');
  // The designation indicator's lines (DoDI 5200.48 fig 2), and the cover-memo statement when only enclosures hold CUI.
  function cuiPoc(memo) {
    const c = memo.cui || {}, own = String(c.poc ?? '').trim(); if (own) return own;
    const p = memo.poc || {}, name = String(p.name ?? '').trim(), reach = String(p.phone ?? '').trim() || String(p.email ?? '').trim();
    return [name, reach].filter(Boolean).join(', ');
  }
  function cuiIndicator(memo) {
    if (!cuiOn(memo)) return [];
    const c = memo.cui, lines = [];
    const comp = String(c.component ?? '').trim(), office = String(c.office ?? '').trim();
    if (!usesLetterhead(memo) && comp) lines.push(`Controlled by: ${comp}`);   // left off when the letterhead names the Component
    lines.push(`Controlled by: ${office}`);
    lines.push(`CUI Category: ${[...new Set(c.cats || [])].sort().join(', ')}`);
    const ldc = String(c.ldc ?? '');
    if (ldc) lines.push(`LDC: ${ldc === 'REL TO' ? `REL TO USA${String(c.rel ?? '').trim() ? ', ' + String(c.rel).trim() : ''}` : ldc}`);
    lines.push(`POC: ${cuiPoc(memo)}`);
    return lines.map(clean);
  }
  // DoD's published example is for email ("This email is unclassified when CUI Document is Removed"); this is that sentence for a memo.
  const cuiCoverText = (memo) => (cuiOn(memo) && memo.cui.cover ? `This memorandum is unclassified when the CUI ${enclosureItems(memo).length > 1 ? 'enclosures are' : 'enclosure is'} removed.` : '');

  /* ---- the layout itself ---- */
  function layout(memo, office, fmtIn) {
    const fmt = { font: 'arial', size: 12, spaces: 2, ...(fmtIn || {}) };
    const F = FONTS[fmt.font] || FONTS.arial, S = +fmt.size || 12, LH = S * 1.15;
    const FL = FONTS.arial;   // letterhead keeps the template's typeface whatever the body font is
    const w = (t, fr) => widthOf(F, S, t, fr);
    const pages = []; let page = null, y = 0, fresh = false;

    const text = (x, top, fr, o = {}) => page.items.push({ k: 't', x, y: top, s: o.s || S, lh: o.lh || LH, f: o.f || fmt.font, fr: typeof fr === 'string' ? [{ t: fr }] : fr, ...(o.body ? { body: true } : {}) });
    const wrapPlain = (s, firstW = TEXT_W, restW = TEXT_W) => wrap(F, S, parseInline(prep(s, fmt)), firstW, restW);
    /* CUI: the designation indicator sits at the bottom of page 1, lower right, with a blank line above it; a cover memo's statement
     * goes above it at the left margin. Both take room from page 1's text, as Word's first-page footer does. */
    const DLH = DI_SIZE * 1.15, diW = RIGHT - SIG_X;
    const di = cuiIndicator(memo).flatMap((t) => wrap(F, DI_SIZE, [{ t }], diW, diW));
    const diX = di.length ? Math.max(SIG_X, Math.floor((RIGHT - Math.max(...di.map((l) => l.w)) - 1) * 20) / 20) : RIGHT;
    const cover = cuiCoverText(memo) ? wrap(F, DI_SIZE, [{ t: cuiCoverText(memo) }], TEXT_W, TEXT_W) : [];
    const firstFoot = cuiOn(memo) ? (cover.length + di.length) * DLH + LH : 0;
    // Continuation pages keep room for the page number, and an enclosure's first page for its "Encl" mark. Enclosure pages leave
    // a point or two for the hairline paragraph that ends each Word section.
    const LIMITS = { first: () => BOTTOM - firstFoot, cont: () => BOTTOM - 2 * LH, encl: () => BOTTOM - 2 * LH - 1, enclCont: () => BOTTOM - 2 };
    const limit = () => LIMITS[page.kind]();
    const room = () => Math.floor((limit() - y + 0.01) / LH);

    const symbolLine = clean(`${memo.officeSymbol || ''}${memo.record ? ` (${memo.record})` : ''}`.trim());
    const subjectLines = wrapPlain(`SUBJECT: ${subjectMark(memo)}${subjectText(memo)}`);

    // kind: 'first' (page 1), 'cont' (a memo continuation page), 'encl' and 'enclCont' (an enclosure roster's pages, plain paper)
    function newPage(kind = 'cont') {
      page = { n: pages.length + 1, kind, items: [] }; pages.push(page);
      if (kind === 'first') return;
      y = IN; fresh = true;                          // 2-5a: office symbol 1 inch from the top edge
      if (kind !== 'cont') return;                   // enclosures carry no memo header and no memo page number
      if (symbolLine) text(LEFT, y, symbolLine);
      y += LH;                                       // 2-5b: subject on the line below
      subjectLines.forEach((ln) => { text(LEFT, y, ln.fr); y += LH; });
      y += 2 * LH;                                   // 2-5c: text on the third line below the subject
      fresh = true;
    }

    /* page 1: letterhead (1-16), or plain paper with the office symbol 1 inch down */
    newPage('first');
    if (usesLetterhead(memo)) {
      if (office?.seal) page.items.push({ k: 'img', x: SEAL.x, y: SEAL.y, w: SEAL.size, h: SEAL.size, src: office.seal });
      else page.items.push({ k: 'seal', x: SEAL.x, y: SEAL.y, w: SEAL.size, h: SEAL.size });
      let ly = LETTERHEAD.top;
      const head = clean(office?.headline ?? 'DEPARTMENT OF THE ARMY').toUpperCase();
      const center = (t, size, lh) => {
        const fr = [{ t, b: true }], lw = widthOf(FL, size, t, fr[0]);
        text((PAGE_W - lw) / 2, ly, fr, { s: size, lh, f: 'arial' }); ly += lh;
      };
      if (head) center(head, LETTERHEAD.headSize, LETTERHEAD.headLH);
      splitLines(office?.lines).forEach((t) => center(clean(t).toUpperCase(), LETTERHEAD.lineSize, LETTERHEAD.lineLH));
      y = SEAL_BOTTOM + LH;                          // 2-4a(1): office symbol on the second line below the seal
    } else y = IN;

    const dateText = memo.date ? clean(fmtDate(memo.date, memo.dateStyle)) : '';
    if (symbolLine) text(LEFT, y, symbolLine);
    if (dateText) text(RIGHT - w(dateText), y, dateText);
    else if (memo.digital) page.items.push({ k: 'date', name: 'Date', x: RIGHT - 130, y, w: 130, h: LH });
    if (memo.suspense) {                             // 2-4a(4): two lines above the date, right margin, bold
      const st = clean(`S: ${fmtDate(memo.suspense, memo.dateStyle)}`), fr = [{ t: st, b: true }];
      text(RIGHT - w(st, fr[0]), y - 2 * LH, fr);
    }

    /* address block (2-4a(5)) */
    y += 3 * LH;
    const hanging = (addrs) => addrs.forEach((a) => wrapPlain(a, TEXT_W, TEXT_W - QTR).forEach((ln, i) => { text(LEFT + (i ? QTR : 0), y, ln.fr); y += LH; }));
    const flush = (s) => wrapPlain(s).forEach((ln) => { text(LEFT, y, ln.fr); y += LH; });
    const to = splitLines(memo.to), thru = splitLines(memo.thru);
    if (memo.kind === 'mfr') flush('MEMORANDUM FOR RECORD');
    else if (memo.kind === 'dist') flush('MEMORANDUM FOR SEE DISTRIBUTION');
    else if (memo.kind === 'exclusive') flush(exclusiveLine(memo));
    else if (memo.kind === 'multi') { flush('MEMORANDUM FOR'); y += LH; hanging(to); }
    else if (memo.kind === 'thru') {
      if (thru.length > 1) { flush('MEMORANDUM THRU'); y += LH; hanging(thru); } else flush(`MEMORANDUM THRU ${thru[0] || ''}`);
      y += LH; flush(`FOR ${to.join(' ')}`);
    } else flush(`MEMORANDUM FOR ${to.join(' ')}`);

    y += LH;                                         // 2-4a(6): subject on the second line below the address
    subjectLines.forEach((ln) => { text(LEFT, y, ln.fr); y += LH; });
    y += 2 * LH;                                     // 2-4b(1): text on the third line below the subject
    fresh = true;

    /* body */
    const paras = markedBody(memo).map((p) => {
      const x0 = LEFT + [0, QTR, 2 * QTR, 2 * QTR][p.level];                         // fig 2-1: indent 1/4, then 1/2 inch; no further
      const lw = p.label ? w(p.label) + (p.level >= 2 ? QTR : w('  ')) : 0;          // 1-39b(10): 1/4 inch after a parenthesis
      const lines = [];
      prep(p.text, fmt).split('\n').forEach((seg, si) => {
        const first = si === 0 ? TEXT_W - (x0 - LEFT) - lw : TEXT_W;
        wrap(F, S, parseInline(seg), first, TEXT_W).forEach((ln, li) => {
          const parts = [];
          if (si === 0 && li === 0) {
            if (p.label) parts.push({ x: x0, fr: [{ t: p.label }] });
            parts.push({ x: x0 + lw, fr: ln.fr });
          } else parts.push({ x: LEFT, fr: ln.fr });                                  // block style: wrapped lines return to the margin
          lines.push({ parts });
        });
      });
      return { id: p.id, lines };
    });

    const auth = memo.kind === 'mfr' ? '' : clean(memo.authority || '').trim().toUpperCase();
    const encl = enclosureLines({ ...memo.encl, items: enclosureItems(memo).join('\n') });
    const close = closingPlan();
    const closeNeed = close.lastRow;   // line steps from the last line of text to the last line of the closing

    const lineText = (ln) => ln.parts.map((p) => p.fr.map((f) => f.t).join('')).join(' ');
    const SENT_END = /[.?!]["”’')\]]*(?:\s+|$)/;
    // 2-5c(2): at least two words on each page of a sentence that crosses the page break.
    function badSentenceBreak(lines, k) {
      const a = lineText(lines[k - 1]).trim(), b = lineText(lines[k]).trim();
      if (/[.?!]["”’')\]]*$/.test(a)) return false;
      const words = (s) => s.trim().split(/\s+/).filter(Boolean).length;
      return words(a.split(SENT_END).pop()) < 2 || words(b.split(SENT_END)[0]) < 2;
    }
    const put = (lines) => {
      lines.forEach((ln) => { ln.parts.forEach((p) => text(p.x, y, p.fr, { body: true })); y += LH; });
      fresh = false;
    };
    // keep: points that must follow the paragraph on the same page (the closing, or the start of a roster)
    function flowPara(P, keep) {
      let lines = P.lines;
      for (;;) {
        const n = lines.length, fit = room();
        if (y + n * LH + keep <= limit() + 0.01) { put(lines); return; }
        // 2-5c(1): keep a paragraph of three lines or fewer whole; two lines at least on each side of a break.
        // 2-5c(4): the closing never starts a page without at least two lines of the last paragraph (or a one-line paragraph whole).
        let k = n <= 3 ? 0 : Math.min(fit, n - 2);
        if (k < 2) k = 0;
        while (k >= 2 && badSentenceBreak(lines, k)) k -= 1;
        if (k < 2) k = 0;
        if (!k && fresh) {
          if (n <= fit) { put(lines); return; }   // the paragraph and the closing together exceed a page: the closing follows on the next page
          k = fit;                                // a paragraph longer than a whole page has to break somewhere
        }
        if (k) { put(lines.slice(0, k)); lines = lines.slice(k); }
        newPage();
      }
    }
    /* rosters in the body. Where AR 25-50 is silent: the full text width from the left margin, the memo's font and size, a blank
     * line before and after (the paragraph spacing of 2-4b(2)), a bold heading row that repeats on each page, and rows that never
     * split. 2-5c(1) carried over to rows: three rows or fewer stay whole, otherwise at least two rows on each side of a break. The
     * paragraph that introduces a roster stays with its heading and first rows, and a roster that ends the text keeps its last two
     * rows with the closing (2-5c(4)). */
    function buildTable(r) {
      const T = rosterTable(r); if (!T) return null;
      if (portionOn(memo)) {   // a table is a portion (3.4b): the mark leads its heading row, and an enclosure's title
        const m = portionMark(r.mark === 'U' ? 'U' : 'CUI');
        T.head = [m + T.head[0], ...T.head.slice(1)]; T.title = m + T.title;
      }
      const bold = { b: true };
      const nat = T.head.map((h, i) => Math.max(w(h, bold), ...T.rows.map((row) => w(row[i]))) + 2 * PAD + 1);   // a point spare, so Word never wraps what fits here
      const widths = fitWidths(nat, T.numbered ? 0 : -1);
      const cell = (t, i, b) => wrap(F, S, [{ t, b }], widths[i] - 2 * PAD, widths[i] - 2 * PAD).map((ln) => ln.fr);
      const grid = [T.head.map((h, i) => cell(h, i, true)), ...T.rows.map((row) => row.map((c, i) => cell(c, i, false)))];
      const heights = grid.map((cells) => Math.max(...cells.map((c) => c.length)) * LH + 2 * CELL_V + BORDER);
      return { ...T, widths, grid, heights };
    }
    // Natural widths, stretched to the full text width (not the No. column) or, when too wide, the widest columns capped so
    // their text wraps. Whole twips, so Word gets exactly the same columns.
    function fitWidths(nat, fixed) {
      const tot = nat.reduce((a, b) => a + b, 0);
      let out = nat.slice();
      if (tot <= TEXT_W) {
        const flex = nat.map((v, i) => (i === fixed ? 0 : v)), fs = flex.reduce((a, b) => a + b, 0) || 1;
        out = nat.map((v, i) => v + ((TEXT_W - tot) * flex[i]) / fs);
      } else {
        const order = nat.map((_, i) => i).sort((a, b) => nat[a] - nat[b]);
        let left = TEXT_W;
        for (let k = 0; k < order.length; k++) {
          const share = left / (order.length - k);
          if (nat[order[k]] <= share) { left -= nat[order[k]]; continue; }
          for (let m = k; m < order.length; m++) out[order[m]] = share;
          break;
        }
      }
      const tws = out.map((v) => Math.floor(v * 20));
      tws[tws.length - 1] += Math.round(TEXT_W * 20) - tws.reduce((a, b) => a + b, 0);
      return tws.map((t) => t / 20);
    }
    // Row i keeps with row i + 1: the heading with row 1, row 1 with row 2, the second-to-last row with the last (Word's keep with next).
    const keepsNext = (T, i) => { const n = T.grid.length - 1; return (i === 0 && n >= 1) || (n >= 2 && (i === 1 || i === n - 1)); };
    function chainFrom(T, i, isLast) {
      const n = T.grid.length - 1; let j = i, need = T.heights[i] + BORDER;   // the border under the last row on the page
      while (j < n && keepsNext(T, j)) { j += 1; need += T.heights[j]; }
      return { j, need: need + (j === n && isLast ? closeNeed * LH : 0) };
    }
    const rosters = [];
    function flowTable(T, isLast, breakKind) {
      const n = T.grid.length - 1, xs = T.widths.reduce((a, wd) => [...a, a[a.length - 1] + wd], [LEFT]);
      const segs = [], placed = []; let seg = null, count = 0;
      const open = () => { seg = { page, top: y, ys: [] }; segs.push(seg); count = 0; };
      const place = (r) => {
        T.grid[r].forEach((lines, c) => lines.forEach((fr, li) => { if (fr.length) text(xs[c] + PAD, y + BORDER + CELL_V + li * LH, fr); }));
        placed.push({ r, page: page.n, y, h: T.heights[r] });
        seg.ys.push(y); y += T.heights[r]; if (r) count += 1; fresh = false;
      };
      const shut = () => {   // borders centered in the half point Word gives them
        seg.bottom = y + BORDER;
        const rule = (x1, y1, x2, y2) => seg.page.items.push({ k: 'rule', x1, y1, x2, y2 });
        [...seg.ys, y].forEach((yy) => rule(LEFT, yy + BORDER / 2, xs[xs.length - 1], yy + BORDER / 2));
        xs.forEach((xx) => rule(xx, seg.top, xx, y + BORDER));
      };
      if (y + chainFrom(T, 0, isLast).need > limit() + 0.01 && !fresh) newPage(breakKind);
      open();
      for (let i = 0; i <= n;) {
        const { j, need } = chainFrom(T, i, isLast);
        if (i && y + need > limit() + 0.01 && count) { shut(); newPage(breakKind); open(); place(0); }   // the heading repeats
        for (let r = i; r <= j; r++) {
          if (r && y + T.heights[r] + BORDER > limit() + 0.01 && count) { shut(); newPage(breakKind); open(); place(0); }
          place(r);
        }
        i = j + 1;
      }
      shut(); y += BORDER;
      rosters.push({
        id: T.id, place: T.place, title: T.title, widths: T.widths, heights: T.heights, cells: [T.head, ...T.rows], placed,
        segs: segs.map((s) => ({ page: s.page.n, top: s.top, bottom: s.bottom, limit: LIMITS[s.page.kind]() })),
      });
    }

    /* the body as blocks: paragraphs, each followed by the rosters placed after it. A roster whose paragraph is gone or empty
     * goes after the last paragraph, so nothing is dropped. */
    const own = new Set((memo.paras || []).map((p) => p.id));
    const ownIds = paras.filter((P) => own.has(P.id)).map((P) => P.id);
    const bodyTables = (memo.rosters || []).filter((r) => r.place !== 'encl').map((r) => ({ r, T: buildTable(r) })).filter((x) => x.T);
    const anchorOf = (r) => (ownIds.includes(r.after) ? r.after : ownIds[ownIds.length - 1] ?? '');
    const blocks = bodyTables.filter((x) => anchorOf(x.r) === '').map((x) => ({ k: 'tab', T: x.T }));
    paras.forEach((P) => { blocks.push({ k: 'p', P }); bodyTables.filter((x) => anchorOf(x.r) === P.id).forEach((x) => blocks.push({ k: 'tab', T: x.T })); });
    blocks.forEach((b, i) => {
      const last = i === blocks.length - 1, next = blocks[i + 1];
      if (b.k === 'p') {
        if (!fresh) y += LH;                         // 2-4b(2): double spacing between paragraphs and subparagraphs
        flowPara(b.P, last ? closeNeed * LH : next?.k === 'tab' ? LH + chainFrom(next.T, 0, i + 1 === blocks.length - 1).need : 0);
      } else {
        // Two rosters in a row need a blank paragraph between them (Word joins adjacent tables); it stays with the second one.
        if (blocks[i - 1]?.k === 'tab') flowPara({ lines: [{ parts: [] }] }, chainFrom(b.T, 0, last).need);
        else if (!fresh) y += LH;
        flowTable(b.T, last, 'cont');
      }
    });

    /* closing: authority line, signature blocks, enclosures stay together */
    if (y + closeNeed * LH > limit() + 0.01 && !fresh) newPage();
    const top = y - LH;                              // top of the last line of text
    Object.entries(close.grid).forEach(([row, cells]) => cells.forEach((c) => text(c.x, top + row * LH, c.fr)));
    close.boxes.forEach(({ row, ...b }) => page.items.push({ ...b, y: top + row * LH }));
    y = top + (close.lastRow + 1) * LH; fresh = false;

    /* distribution, then copies furnished (2-4a(5)(c), 2-4c(5)) */
    const flowLines = (items) => items.forEach((it, i) => {
      const left = items.length - i;
      if (room() < 1 || (left > 1 && room() < 2)) {
        if (room() >= 1) { text(LEFT + 4, y, '(CONT)'); y += LH; }   // fig 2-8
        newPage(); fresh = false;
      }
      text(LEFT + (it.indent || 0), y, it.fr); y += LH;
    });
    if (memo.kind === 'dist') {
      y += LH;
      const list = String(memo.distribution ?? '').split(/\r?\n/).filter((x) => x.trim());
      const items = [{ fr: [{ t: 'DISTRIBUTION:' }] }];
      list.forEach((raw) => {
        const lead = raw.match(/^\s*/)[0].length;
        wrapPlain(raw, TEXT_W - lead * w(' ')).forEach((ln) => items.push({ fr: ln.fr, indent: lead * w(' ') }));
      });
      if (room() < 3) newPage();
      flowLines(items);
    }
    const cf = splitLines(memo.cf);
    if (cf.length) {
      y += LH;
      if (room() < 2) { newPage(); fresh = false; }
      flowLines([{ fr: [{ t: memo.cfNoEncl ? 'CF: (w/o encls)' : 'CF:' }] }, ...cf.flatMap((c) => wrapPlain(c).map((ln) => ({ fr: ln.fr })))]);
    }

    /* enclosure rosters, after everything else, each from a new page on plain paper: no letterhead, memo header or memo page
     * number, since an enclosure is its own document. 1-34 and 4-2d(1): "Encl" and its number in the lower right corner of the
     * first page. The centered bold title is MemoGen's choice; AR 25-50 is silent. */
    const typed = splitLines(memo.encl?.items).length, enclTotal = enclosureItems(memo).length;
    enclosureRosters(memo).forEach((r, k) => {
      const T = buildTable(r), num = typed + k + 1, mark = enclTotal > 1 ? `Encl ${num}` : 'Encl';
      newPage('encl'); page.encl = num;
      text(RIGHT - w(mark), BOTTOM - LH, mark);
      wrap(F, S, [{ t: T.title, b: true }], TEXT_W, TEXT_W).forEach((ln) => { text((PAGE_W - ln.w) / 2, y, ln.fr); y += LH; });
      y += LH;
      fresh = true;   // the title always shares its page with the table
      flowTable(T, false, 'enclCont');
      Object.assign(rosters[rosters.length - 1], { encl: num, mark });
    });

    /* CUI: banner at the top and bottom of every page, enclosures included; the designation indicator on page 1 */
    if (cuiOn(memo)) {
      const bf = [{ t: CUI_BANNER, b: true }], bx = (PAGE_W - widthOf(F, S, CUI_BANNER, bf[0])) / 2;
      pages.forEach((p) => [BANNER_EDGE, PAGE_H - BANNER_EDGE - LH].forEach((by) => p.items.push({ k: 't', x: bx, y: by, s: S, lh: LH, f: fmt.font, fr: bf })));
      let fy = BOTTOM - (cover.length + di.length) * DLH;
      cover.forEach((ln) => { pages[0].items.push({ k: 't', x: LEFT, y: fy, s: DI_SIZE, lh: DLH, f: fmt.font, fr: ln.fr }); fy += DLH; });
      di.forEach((ln) => { pages[0].items.push({ k: 't', x: diX, y: fy, s: DI_SIZE, lh: DLH, f: fmt.font, fr: ln.fr }); fy += DLH; });
    }

    /* 2-5d: page number centered about 1 inch from the bottom, on continuation pages */
    pages.filter((p) => p.kind === 'cont').forEach((p) => {
      const t = String(p.n);
      p.items.push({ k: 't', x: (PAGE_W - widthOf(F, S, t)) / 2, y: BOTTOM - LH, s: S, lh: LH, f: fmt.font, fr: [{ t }] });
    });

    return {
      pages, lh: LH, size: S, font: fmt.font, officeTop: usesLetterhead(memo) ? SEAL_BOTTOM + LH : IN, subjectLines: subjectLines.length,
      closing: { grid: close.grid, lastRow: close.lastRow }, paraLines: paras.map((p) => ({ id: p.id, lines: p.lines.length })),
      body: blocks.map((b) => (b.k === 'p' ? { k: 'p', id: b.P.id } : { k: 'tab', id: b.T.id })), rosters,
      cui: cuiOn(memo) ? { banner: CUI_BANNER, edge: BANNER_EDGE, size: DI_SIZE, lh: DLH, diX, di: di.map((l) => l.fr.map((f) => f.t).join('')), cover: cover.map((l) => l.fr.map((f) => f.t).join('')), coverText: cuiCoverText(memo), diText: cuiIndicator(memo) } : null,
    };

    // The closing as rows counted down from the last line of text (row 1 is the line just below it). The preview, the PDF and the
    // Word export all draw from this one grid, so the blocks land in the same place in each.
    function closingPlan() {
      const grid = {}, boxes = [];
      const block = (row, x, lines) => lines.forEach((l, i) => (grid[row + i] ||= []).push({ x: x + (l.indent || 0), fr: l.fr }));
      const sigBox = (name, label, nameRow, x, lock) => { if (memo.digital) boxes.push({ k: 'sig', name, label, lock, x, row: nameRow - 3.6, w: 198, h: 3.2 * LH }); };   // app F-2f: directly above and left-aligned with the name
      const rule = '_'.repeat(Math.round(180 / w('_'))), date = '(Date)', ruleW = w(rule);
      const leftW = SIG_X - LEFT - 12, rightW = RIGHT - SIG_X;
      const blockLines = (s, width) => signatureLines(s, width).map((l) => ({ fr: l.fr, indent: l.indent ? QTR : 0 }));
      // figs 2-15, 2-16: overscored, with a date under the block. Returns the block's last row.
      const ruled = (row, x, width, s, name, label, dateName, lock) => {
        const lines = [{ fr: [{ t: rule }] }, ...blockLines(s, width), { fr: [{ t: rule }] }, { fr: [{ t: date }], indent: (ruleW - w(date)) / 2 }];
        block(row, x, lines);
        sigBox(name, label, row + 1, x, lock);
        if (memo.digital) boxes.push({ k: 'date', name: dateName, align: 'center', x, row: row + lines.length - 2, w: ruleW, h: LH });
        return row + lines.length - 1;
      };
      const senior = memo.cosign?.on ? memo.cosign.signer : null, recipient = memo.receipt?.on ? memo.receipt.signer : null;
      const topDate = !dateText && memo.digital ? ['Date'] : [];   // the date box at the top goes with the last issuing signature (app F-2f, F-2h)
      const RECEIPT = ['ReceiptSignature', 'Recipient signature box', 'ReceiptDate', ['ReceiptDate']];
      const ISSUER = ['Signature1', 'Digital signature box', 'Signer1Date', ['Signer1Date']];
      const SENIOR = ['Signature2', 'Senior signer signature box', 'Signer2Date', [...topDate, 'Signer2Date']];
      const first = auth ? 7 : 5;   // the fifth line below the authority line or the text
      if (auth) block(2, LEFT, [{ fr: [{ t: auth }] }]);
      let last;
      if (senior && recipient) {    // 2-6c(5)(d): highest ranking centered at the bottom, next highest on the right, junior on the left
        const upper = Math.max(ruled(first, LEFT, leftW, recipient, ...RECEIPT), ruled(first, SIG_X, rightW, memo.signer, ...ISSUER));
        last = ruled(upper + 5, (PAGE_W - ruleW) / 2, rightW, senior, ...SENIOR);
      } else if (senior) {          // 2-6c(5)(d): senior official on the right, which keeps a block at the center of the page (2-4c(2)(a))
        last = Math.max(ruled(first, LEFT, leftW, memo.signer, ...ISSUER), ruled(first, SIG_X, rightW, senior, ...SENIOR));
      } else {
        const sig = blockLines(memo.signer, rightW);
        block(first, SIG_X, sig);
        sigBox('Signature1', 'Digital signature box', first, SIG_X, topDate);
        if (recipient) last = Math.max(first + sig.length - 1, ruled(first - 1, LEFT, leftW, recipient, ...RECEIPT));
        else { block(first, LEFT, encl); last = first + Math.max(sig.length, encl.length, 1) - 1; }
      }
      if ((senior || recipient) && encl.length) { block(last + 2, LEFT, encl); last += 1 + encl.length; }   // fig 2-14: two lines below the lowest block
      return { grid, boxes, lastRow: last };
    }
    function signatureLines(s = {}, width = RIGHT - SIG_X) {
      const name = clean(s.name || '').trim().toUpperCase();
      const lines = [];
      if (name) lines.push({ fr: [{ t: name }] });
      if (!s.civilian) {   // 6-4c: name, then rank and branch, then title. 6-5c(11): a Guard designation such as CAARNG follows the branch
        const rb = clean([s.rank, s.branch, s.component].map((x) => String(x ?? '').trim()).filter(Boolean).join(', '));
        if (rb) lines.push({ fr: [{ t: rb }] });
      }
      const addWrapped = (t) => wrap(F, S, [{ t: clean(t).trim() }], width, width - QTR).forEach((ln, i) => lines.push({ fr: ln.fr, indent: i > 0 }));   // a long title continues 1/4 inch in
      if (String(s.title ?? '').trim()) addWrapped(s.title);
      if (String(s.org ?? '').trim()) addWrapped(s.org);   // 6-5d: organization is the final element
      return lines;
    }
    function enclosureLines(e = {}) {
      const items = splitLines(e.items), n = items.length;
      if (!n) return [];
      const width = SIG_X - LEFT - 12;
      const head = n === 1 ? 'Encl' : e.inBody ? 'Encls' : `${n} Encls`;   // 2-4c(3), tables 4-1 to 4-5
      const lines = [{ fr: [{ t: head }] }];
      if (e.inBody) return lines;                                          // table 4-4: named in the text, so just "Encl" or "Encls"
      if (n === 1) { wrap(F, S, [{ t: clean(items[0]) }], width, width).forEach((ln) => lines.push({ fr: ln.fr })); return lines; }
      items.forEach((t, i) => {
        const label = `${i + 1}.  `, lw = w(label);
        wrap(F, S, [{ t: clean(t) }], width - lw, width - lw).forEach((ln, j) => lines.push(j ? { fr: ln.fr, indent: lw } : { fr: [{ t: label }, ...ln.fr] }));
      });
      return lines;
    }
  }

  /* ---- fill-ins: anything in [brackets] that still has to be replaced, such as [DATE] or [RANK Last, First MI]. "[ ]" is a checkbox
   * and underscores are blanks for the recipient, so neither counts. Only fields that print are looked at, in the order the form
   * shows them. Returns { key, where, text, at }: key "m:<field>" (its data-m path), "p:<paragraph id>" or "r:<roster id>". ---- */
  const FILL_RE = /\[(?! \])[^\[\]\n]*[A-Za-z0-9][^\[\]\n]*\]/g;
  function fillIns(memo) {
    const out = [], k = memo.kind;
    const scan = (key, where, v) => { for (const m of String(v ?? '').matchAll(FILL_RE)) out.push({ key, where, text: m[0], at: m.index }); };
    const person = (path, who, s) => {
      if (!s) return;
      scan(`m:${path}.name`, `${who}: Name`, s.name);
      if (!s.civilian) [['rank', 'Rank'], ['branch', 'Branch'], ['component', 'Component']].forEach(([f, l]) => scan(`m:${path}.${f}`, `${who}: ${l}`, s[f]));
      scan(`m:${path}.title`, `${who}: Position`, s.title);
      scan(`m:${path}.org`, `${who}: Organization`, s.org);
    };
    scan('m:officeSymbol', 'Office symbol', memo.officeSymbol);
    scan('m:record', 'ARIMS record number', memo.record);
    if (cuiOn(memo)) {
      const c = memo.cui;
      if (!usesLetterhead(memo)) scan('m:cui.component', 'CUI: DoD Component', c.component);
      scan('m:cui.office', 'CUI: office', c.office);
      if (c.ldc === 'REL TO') scan('m:cui.rel', 'CUI: REL TO', c.rel);
      scan('m:cui.poc', 'CUI: POC', c.poc);
    }
    if (k === 'thru') scan('m:thru', 'THRU', memo.thru);
    if (['single', 'multi', 'thru', 'exclusive'].includes(k)) scan('m:to', k === 'thru' ? 'FOR' : k === 'exclusive' ? 'Memorandum Exclusive For' : 'MEMORANDUM FOR', memo.to);
    if (k === 'dist') scan('m:distribution', 'Distribution', memo.distribution);
    scan('m:subject', 'Subject', memo.subject);
    scan('m:subjectName', 'Name after the subject', memo.subjectName);
    // "paragraph 1b(2)": each label joined to its parents', the way AR 25-50 refers to paragraphs
    const path = [], ref = new Map();
    numbered(bodyParas(memo)).forEach((p) => { path[p.level] = p.label.replace(/\.$/, ''); path.length = p.level + 1; ref.set(p.id, path.join('')); });
    (memo.paras || []).forEach((p) => scan(`p:${p.id}`, ref.get(p.id) ? `paragraph ${ref.get(p.id)}` : 'the paragraph', p.text));
    (memo.rosters || []).forEach((r, i) => { if (r.place === 'encl') scan(`r:${r.id}`, `Roster ${i + 1} title`, r.title); });
    scan('m:poc.name', 'POC: Rank and name', memo.poc?.name);
    scan('m:poc.phone', 'POC: Phone', memo.poc?.phone);
    scan('m:poc.email', 'POC: Email', memo.poc?.email);
    if (k !== 'mfr') scan('m:authority', 'Authority line', memo.authority);
    person('signer', 'Signature', memo.signer);
    if (memo.cosign?.on) person('cosign.signer', 'Second signer', memo.cosign.signer);
    if (memo.receipt?.on) { person('receipt.signer', 'Receipt', memo.receipt.signer); scan('m:receipt.text', 'Receipt wording', memo.receipt.text); }
    scan('m:encl.items', 'Enclosures', memo.encl?.items);
    scan('m:cf', 'Copies furnished', memo.cf);
    return out;
  }

  global.MemoFormat = { BOX, BOX_GLYPH, FILL_RE, fillIns, layout, rosterTable,rosterTitle, enclosureItems, markedBody, subjectMark, cuiOn, portionOn, cuiIndicator, cuiPoc, cuiCoverText, CUI_CATEGORIES, CUI_LDCS, numbered, bodyParas, pocText, receiptText, exclusiveLine, subjectText, printable, parseInline, fmtDate, clean, plain, widthOf, spaceSentences, usesLetterhead, splitLines, FONTS, KINDS, LETTERHEAD, PAGE_W, PAGE_H, MONTH_NAMES };
})(window);
