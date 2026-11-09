/* MemoChecks: reviews a memo against AR 25-50. Things the layout does by itself (spacing, indents, numbering, page breaks) are not
 * checks; these are the things only the writer can settle. Each result cites its paragraph.
 * level: 'fix' (breaks a rule), 'warn' (likely wrong, or a "should"), 'tip' (style guidance from section IV). */
(function (global) {
  const { numbered, bodyParas, subjectText, plain, splitLines, MONTH_NAMES, fmtDate, usesLetterhead, rosterTable, enclosureItems, cuiOn, cuiPoc, CUI_CATEGORIES, fillIns, FILL_RE } = MemoFormat;
  const PRIVACY_CATS = ['PRVCY', 'PERS', 'MIL', 'HLTH', 'CONTRACT', 'DREC', 'GENETIC', 'PRIIG', 'STUD'];   // the registry's Privacy index group
  // Roster columns with personal information (1-22d sends CUI marking to DoDI 5200.48). The heading is one clue; the numbers are the
  // other, since an ID column can hold a DoD ID, an SSN, a truncated SSN or another identifying number whatever it is called.
  const PII_COL = /\b(dod ?id|dodid|edipi|id|id (no|number)|identification|identifier|service (no|number)|passport|driver'?s? licen[cs]e|employee id|emplid|dob|date of birth|birth ?date|home address|personal (phone|email))\b/i;
  const SSN_COL = /\b(ssn|social security|social)\b/i;
  const LAST4_COL = /\b(last ?4|last four|l4)\b/i;   // a truncated SSN still falls under 6-5b
  const DODID_COL = /\b(dod ?id|dodid|edipi)\b/i;     // a DoD ID number (EDIPI) is 10 digits
  const NOT_ID_COL = /\b(phone|tel|dsn|fax|date|time|dtg)\b/i;   // long numbers that are not identifiers
  const SSN = /\b\d{3}-\d{2}-\d{4}\b/;
  const SSN_SHAPE = /^\d{3}[\s-]?\d{2}[\s-]?\d{4}$/;   // nine digits, with or without the dashes
  const GENERAL_ABBR = /^(GEN|LTG|MG|BG)$/i;
  const SENSITIVE = /\b(arrest(ed)?|continuous vetting|security clearance|derogatory|reprimand|misconduct)\b/i;

  // Common misspellings, including Army words. A short list, not a dictionary: the browser's red underlines cover the rest.
  const MISSPELLED = new Map(Object.entries({
    thier: 'their', recieve: 'receive', recieved: 'received', recieving: 'receiving', reciept: 'receipt', seperate: 'separate', seperately: 'separately', seperated: 'separated',
    occured: 'occurred', occuring: 'occurring', occurence: 'occurrence', accomodate: 'accommodate', accomodations: 'accommodations', definately: 'definitely', definatly: 'definitely',
    goverment: 'government', govenment: 'government', enviroment: 'environment', managment: 'management', maintainance: 'maintenance', maintenence: 'maintenance',
    equiptment: 'equipment', equipement: 'equipment', neccessary: 'necessary', necesary: 'necessary', neccesary: 'necessary', recomend: 'recommend', reccomend: 'recommend',
    recomendation: 'recommendation', reccomendation: 'recommendation', responsiblity: 'responsibility', responsibilty: 'responsibility', accountablity: 'accountability',
    accountibility: 'accountability', persue: 'pursue', untill: 'until', wich: 'which', teh: 'the', adn: 'and', alot: 'a lot', beleive: 'believe', belive: 'believe',
    calender: 'calendar', commited: 'committed', commitee: 'committee', comittee: 'committee', curent: 'current', existance: 'existence', foriegn: 'foreign', futher: 'further',
    gaurd: 'guard', gaurantee: 'guarantee', guarentee: 'guarantee', harrass: 'harass', harrassment: 'harassment', immediatly: 'immediately', immediatley: 'immediately',
    independant: 'independent', intelligance: 'intelligence', knowlege: 'knowledge', liason: 'liaison', liasion: 'liaison', lisence: 'license', millitary: 'military', militray: 'military',
    noticable: 'noticeable', occassion: 'occasion', offical: 'official', oppurtunity: 'opportunity', orginal: 'original', orginization: 'organization', organiztion: 'organization',
    paralell: 'parallel', perfomance: 'performance', performence: 'performance', permenant: 'permanent', permanant: 'permanent', personel: 'personnel', personell: 'personnel',
    posession: 'possession', possesion: 'possession', prefered: 'preferred', priviledge: 'privilege', privelege: 'privilege', proceedure: 'procedure', procedue: 'procedure',
    publically: 'publicly', questionaire: 'questionnaire', refered: 'referred', refering: 'referring', relevent: 'relevant', reponse: 'response', responce: 'response',
    resonsible: 'responsible', responisble: 'responsible', schedual: 'schedule', secratary: 'secretary', secuirty: 'security', sercurity: 'security', secrurity: 'security',
    sieze: 'seize', similiar: 'similar', sincerly: 'sincerely', succesful: 'successful', successfull: 'successful', supress: 'suppress', tommorow: 'tomorrow', tommorrow: 'tomorrow',
    transfered: 'transferred', transfering: 'transferring', truely: 'truly', usally: 'usually', vehical: 'vehicle', vehicule: 'vehicle', wierd: 'weird', writting: 'writing',
    batallion: 'battalion', battalian: 'battalion', batalion: 'battalion', battallion: 'battalion', sargeant: 'sergeant', sergent: 'sergeant', seargent: 'sergeant', sargent: 'sergeant',
    lieutentant: 'lieutenant', leiutenant: 'lieutenant', lieutanant: 'lieutenant', liuetenant: 'lieutenant', colonal: 'colonel', coronel: 'colonel', commader: 'commander',
    comander: 'commander', commmander: 'commander', commandor: 'commander', brigrade: 'brigade', regiement: 'regiment', regimant: 'regiment', infantary: 'infantry', infanty: 'infantry',
    artillary: 'artillery', calvary: 'cavalry', soldeir: 'Soldier', solider: 'Soldier', memorandom: 'memorandum', memorandun: 'memorandum', enclousure: 'enclosure', authorty: 'authority',
    adjudicaton: 'adjudication', counterintellegence: 'counterintelligence', investagation: 'investigation', investigaton: 'investigation', investigater: 'investigator',
    credentails: 'credentials', clearence: 'clearance',
  }));

  // Table 6-1, lowest to highest. SPC and CPL share a pay grade, as do 1SG/MSG and CSM/SGM.
  const RANKS = [['PV1'], ['PV2'], ['PFC'], ['SPC', 'CPL'], ['SGT'], ['SSG'], ['SFC'], ['MSG', '1SG'], ['SGM', 'CSM'], ['WO1'], ['CW2'], ['CW3'], ['CW4'], ['CW5'], ['2LT'], ['1LT'], ['CPT'], ['MAJ'], ['LTC'], ['COL'], ['BG', 'BRIGADIER GENERAL'], ['MG', 'MAJOR GENERAL'], ['LTG', 'LIEUTENANT GENERAL'], ['GEN', 'GENERAL']];
  const rankOrder = (r) => RANKS.findIndex((g) => g.includes(String(r ?? '').trim().toUpperCase()));

  // 1-38d: a form of "to be" plus a past participle. Words below end the same way but read as plain adjectives here.
  const PASSIVE = /\b(?:am|is|are|was|were|be|been|being)\s+(?:[a-z]+ly\s+)?([a-z]+ed|written|taken|given|driven|chosen|broken|stolen|spoken|shown|seen|known|forgotten|hidden|built|done|held|kept|lost|made|paid|sent|sold|told|found|brought|bought|caught|taught|sought)\b/gi;
  const NOT_PASSIVE = new Set(['indeed', 'red', 'bed', 'shed', 'speed', 'need', 'feed', 'seed', 'exceed', 'proceed', 'succeed', 'hundred', 'sacred', 'tired', 'pleased', 'interested', 'concerned', 'located', 'married', 'qualified', 'supposed', 'dated', 'based', 'excited', 'assigned']);

  function check(memo, office, L) {
    const out = [];
    const add = (level, ref, msg, target) => out.push({ level, ref, msg, target });
    const paras = numbered(bodyParas(memo));
    const body = paras.map((p) => plain(p.text)).join('\n');
    const subject = plain(subjectText(memo)).trim();
    const to = splitLines(memo.to), thru = splitLines(memo.thru);
    const s = memo.signer || {};

    /* fill-ins left from a template: one line per field, which takes you to it */
    const fills = new Map();
    fillIns(memo).forEach((f) => { if (!fills.has(f.key)) fills.set(f.key, { where: f.where, texts: [] }); const g = fills.get(f.key); if (!g.texts.includes(f.text)) g.texts.push(f.text); });
    fills.forEach((g, key) => add('fix', 'MemoGen', `Fill in ${g.texts.length > 1 ? `${g.texts.slice(0, -1).join(', ')} and ${g.texts[g.texts.length - 1]}` : g.texts[0]} (${g.where}).`, key));

    /* heading */
    if (!String(memo.officeSymbol ?? '').trim()) add('fix', '2-4a(1)', 'Add the office symbol.', 'officeSymbol');
    if (!String(memo.record ?? '').trim()) add('warn', '2-4a(2)', 'Add the ARIMS record number (it goes in parentheses after the office symbol).', 'record');
    if (!memo.date) add('tip', '2-4a(3)', memo.digital ? 'Date left open: the signer fills in the date box before signing digitally.' : 'Date left open: type or stamp it after the memo is signed.', 'date');
    if (memo.kind === 'mfr' && memo.letterhead) add('warn', 'fig 2-17', 'An MFR is typed on plain white paper. This one is set to use letterhead.', 'kind');
    if (usesLetterhead(memo)) {
      if (!office?.seal) add('warn', '1-16b(1)', 'Letterhead must bear the DoD seal. Add it on the Saved tab.', 'setup');
      const lines = splitLines(office?.lines);
      if (!lines.length) add('fix', '1-16a', 'Fill in the letterhead on the Saved tab: organization and full mailing address.', 'setup');
      else if (!/\b\d{5}-\d{4}\b/.test(lines.join(' '))) add('warn', '1-35', 'Letterhead should end with the nine-digit ZIP code (12345-1234).', 'setup');
    }

    /* address */
    if (['single', 'multi', 'thru', 'exclusive'].includes(memo.kind) && !to.length) add('fix', '2-4a(5)', 'Add the addressee.', 'to');
    if (memo.kind === 'exclusive') {
      if (to.length && to.join(', ').split(',').filter((x) => x.trim()).length < 3) add('warn', '1-12b', 'Give the full name, title and mailing address, separated by commas.', 'to');
      add('tip', '1-12c', 'Place it in a sealed envelope, with "Exclusive For" printed and underlined on the envelope.', 'to');
    }
    if (memo.kind === 'single' && to.length > 1) add('warn', '2-4a(5)(b)', 'More than one addressee: switch to Multiple address.', 'kind');
    if (memo.kind === 'multi' && to.length > 5) add('fix', '2-4a(5)(c)', `${to.length} addressees: more than five must use SEE DISTRIBUTION.`, 'kind');
    if (memo.kind === 'multi' && to.length === 1) add('warn', '2-4a(5)(a)', 'Only one addressee: switch to Single address.', 'kind');
    if (memo.kind === 'dist' && !splitLines(memo.distribution).length) add('fix', '2-4a(5)(c)', 'List the recipients under Distribution.', 'distribution');
    if (memo.kind === 'thru') {
      if (!thru.length) add('fix', '2-4a(5)(d)', 'Add the THRU addressee.', 'thru');
      if (thru.length > 2) add('warn', 'fig 2-12', 'Avoid more than two THRU addressees unless absolutely necessary.', 'thru');
    }

    /* subject */
    if (!subject) add('fix', '2-4a(6)', 'Add a subject.', 'subject');
    else {
      const n = subject.split(/\s+/).length;
      if (n > 10) add('warn', '2-4a(6)', `Subject is ${n} words. Keep it to 10 or fewer if you can.`, 'subject');
      const acr = [...new Set(subject.replace(FILL_RE, '').match(/\b[A-Z]{2,}\b/g) || [])].filter((a) => !['DA', 'FY', 'HQDA', 'DOD'].includes(a));
      if (acr.length) add('tip', '1-15c(2)', `Avoid abbreviations in the subject line (${acr.join(', ')}).`, 'subject');
    }

    /* body structure */
    if (!paras.length) add('fix', '2-4b', 'Write the body of the memo.', 'body');
    const groups = new Map();   // subdivisions of the same parent: each needs at least two members
    const stack = [];
    paras.forEach((p) => {
      stack.length = p.level; stack[p.level] = p;
      if (!p.level) return;
      const key = `${p.level}:${stack[p.level - 1]?.id ?? ''}`;
      const g = groups.get(key) || { n: 0, p }; g.n += 1; groups.set(key, g);
    });
    groups.forEach((g) => { if (g.n === 1) add('fix', 'fig 2-1', `Subparagraph ${g.p.label} stands alone. A subdivided paragraph needs at least two subparagraphs.`, g.p.id); });

    const longPara = (L?.paraLines || []).filter((p) => p.lines > 10);
    longPara.forEach((p) => { const q = paras.find((x) => x.id === p.id); add('tip', '1-39b(3)', `Paragraph ${q?.label || ''} runs ${p.lines} lines. Aim for 10 or fewer.`, p.id); });

    if (paras.length > 1 && memo.kind !== 'mfr') {
      const last = plain(paras[paras.length - 1].text);
      if (!/point of contact|\bPOC\b/i.test(last) && !/@|\(\d{3}\)|\b\d{3}-\d{3,4}\b/.test(last)) add('warn', '2-4b(1)(e)', 'Put the point of contact in the last paragraph.', 'body');
    }

    /* body wording */
    const hit = (re) => paras.find((p) => re.test(plain(p.text)));
    let p;
    if (/\b\d{3}-\d{2}-\d{4}\b/.test(`${body}\n${subject}`)) add('fix', '6-5b', 'That looks like a Social Security number. Do not use SSNs in correspondence unless a statute requires it.', 'body');
    if ((p = hit(/(^|\n)\s*P\.?\s?S\.?[:\s]/))) add('fix', '1-29', 'Do not use postscripts.', p.id);
    if ((p = hit(/\bsoldiers?\b/))) add('warn', '1-14a', `Paragraph ${p.label || '1'}: capitalize "Soldier" when it means a U.S. Army Soldier.`, p.id);
    if ((p = hit(/\b\d{4}\s+hours\b/i))) add('warn', '1-26', `Paragraph ${p.label || '1'}: drop the word "hours" after military time.`, p.id);
    if ((p = hit(/\b\d{1,2}(:\d{2})?\s?[ap]\.?m\b/i))) add('warn', '1-26', `Paragraph ${p.label || '1'}: use military time in memorandums.`, p.id);
    if ((p = hit(new RegExp(`\\b(${MONTH_NAMES.join('|')}) \\d{1,2}, \\d{4}\\b`)))) add('warn', '1-25a', `Paragraph ${p.label || '1'}: write dates day first (13 March 2020 or 13 Mar 20).`, p.id);
    if ((p = hit(/(^|[.?!]\s+)(It is|There is|There are)\b/))) add('tip', '1-39b(8)', `Paragraph ${p.label || '1'}: avoid opening a sentence with "It is," "There is" or "There are."`, p.id);
    const para_ = (q) => (q.label || '1').replace(/\.$/, '');

    /* spelling (1-39b(5)) */
    const spell = (where, text, target) => {
      const seen = new Set((String(text ?? '').match(/[A-Za-z]+/g) || []).map((w) => w.toLowerCase()));
      seen.forEach((w) => { if (MISSPELLED.has(w)) add('fix', '1-39b(5)', `${where}: "${w}" looks misspelled. Try "${MISSPELLED.get(w)}".`, target); });
    };
    spell('Subject', subject, 'subject');
    spell('Address', `${memo.to ?? ''} ${memo.thru ?? ''}`, 'to');
    paras.forEach((q) => spell(`Paragraph ${para_(q)}`, plain(q.text), q.id));
    spell('Signature block', `${s.title ?? ''} ${s.org ?? ''}`, 'signer');
    if (memo.receipt?.on) spell('Receipt block', `${memo.receipt.signer?.title ?? ''} ${memo.receipt.signer?.org ?? ''}`, 'receipt');
    if (usesLetterhead(memo)) spell('Letterhead', `${office?.headline ?? ''} ${office?.lines ?? ''}`, 'setup');

    /* an abbreviation set up in parentheses has to be used again, or the term should just be written out (1-15c(2)) */
    const defined = new Set();
    for (const m of body.matchAll(/\(([^()]{2,40})\)/g)) {
      m[1].split(/\s*(?:,|\/|\bor\b|\band\b)\s*/).forEach((t) => { if (/^[A-Z][A-Za-z0-9]*[A-Z0-9]$/.test(t) && (t.match(/[A-Z]/g) || []).length >= 2) defined.add(t); });
    }
    defined.forEach((a) => {
      if ((body.match(new RegExp(`\\b${a}s?\\b`, 'g')) || []).length > 1) return;
      const q = paras.find((x) => plain(x.text).includes(a));
      add('warn', '1-15c(2)', `"${a}" is set up in paragraph ${q ? para_(q) : '1'} but never used again. Write the term out and drop the abbreviation.`, q?.id || 'body');
    });

    /* passive voice (1-38b): a pattern match, so it is offered as a tip and can be wrong */
    const passive = [];
    paras.filter((q) => q.id !== 'poc' && q.id !== 'receipt').forEach((q) => {
      for (const m of plain(q.text).matchAll(PASSIVE)) if (!NOT_PASSIVE.has(m[1].toLowerCase())) passive.push({ q, phrase: m[0] });
    });
    if (passive.length) {
      const shown = passive.slice(0, 3).map((x) => `"${x.phrase}" (para ${para_(x.q)})`).join(', ');
      add('tip', '1-38b', `Possible passive voice: ${shown}${passive.length > 3 ? ` and ${passive.length - 3} more` : ''}. Say who does the action.`, passive[0].q.id);
    }

    /* sensitive matters */
    if (SENSITIVE.test(`${subject}\n${body}`)) {
      if (['single', 'multi', 'thru', 'dist'].includes(memo.kind)) add('tip', '1-12a', 'This looks sensitive. If it is directed to one person, consider the Exclusive For memo type.', 'kind');
      if (!cuiOn(memo)) add('tip', '1-22d', 'Confirm whether this memo needs a CUI marking. AR 25-50 refers to DoDI 5200.48.', 'cui');
    }

    const sentences = body.split(/[.?!]+(?:\s+|$)/).map((x) => x.trim()).filter(Boolean);
    if (sentences.length >= 3) {
      const avg = sentences.reduce((a, x) => a + x.split(/\s+/).length, 0) / sentences.length;
      if (avg > 20) add('tip', '1-39b(2)', `Sentences average ${Math.round(avg)} words. Aim for about 15.`, 'body');
    }

    /* suspense */
    if (memo.suspense) {
      const forms = [fmtDate(memo.suspense, 'long'), fmtDate(memo.suspense, 'short')];
      if (!forms.some((f) => body.includes(f))) add('warn', '1-27a', 'A suspense date also goes in the body of the memo.', 'body');
    }

    /* closing */
    if (memo.kind === 'mfr' && String(memo.authority ?? '').trim()) add('warn', 'fig 2-17', 'An MFR does not use an authority line. It is left off this memo.', 'authority');
    if (memo.kind !== 'mfr' && String(memo.authority ?? '').trim() && !/:\s*$/.test(memo.authority)) add('warn', 'fig 2-1', 'End the authority line with a colon.', 'authority');
    if (!String(s.name ?? '').trim()) add('fix', '6-4a(1)', 'Add the signer\'s name.', 'signer');
    if (!String(s.title ?? '').trim()) add('warn', '6-4c', 'Add the signer\'s title.', 'signer');
    if (!s.civilian) {
      if (!String(s.rank ?? '').trim()) add('fix', '6-4c', 'Add the signer\'s rank.', 'signer');
      if (!String(s.branch ?? '').trim() && !String(s.component ?? '').trim()) add('warn', '6-4f(2)', 'Use the branch abbreviation in the signature block.', 'signer');
      if (GENERAL_ABBR.test(String(s.rank ?? '').trim())) add('fix', '6-4f(3)', 'General officers use their full grade (Major General, not MG).', 'signer');
    }
    if (/\(P\)/.test(`${s.rank ?? ''} ${s.name ?? ''}`)) add('tip', '6-5c(2)', 'Leave "(P)" out of the signature block unless it benefits the Army\'s image.', 'signer');

    /* second signature */
    if (memo.cosign?.on) {
      const c = memo.cosign.signer || {};
      if (!String(c.name ?? '').trim()) add('fix', '6-4a(1)', 'Add the name of the senior signer.', 'cosign');
      if (!String(c.title ?? '').trim()) add('warn', '6-4c', 'Add the senior signer\'s title.', 'cosign');
      if (!c.civilian) {
        if (!String(c.rank ?? '').trim()) add('fix', '6-4c', 'Add the senior signer\'s rank.', 'cosign');
        if (GENERAL_ABBR.test(String(c.rank ?? '').trim())) add('fix', '6-4f(3)', 'General officers use their full grade (Major General, not MG).', 'cosign');
      }
      const a = rankOrder(s.rank), b = rankOrder(c.rank);
      if (!s.civilian && !c.civilian && a >= 0 && b >= 0 && b < a) add('warn', '2-6c(5)(d)', 'The senior official goes on the right, but this signer ranks below you. Swap the two blocks.', 'cosign');
      if (/\bcommander\b/i.test(String(c.title ?? '')) && String(memo.authority ?? '').trim()) add('warn', '6-2b', 'Omit the authority line when the commander signs personally.', 'authority');
    }

    /* receipt block */
    if (memo.receipt?.on) {
      const r = memo.receipt.signer || {};
      if (!String(r.name ?? '').trim()) add('fix', '6-4a(1)', 'Add the name of the person acknowledging receipt.', 'receipt');
      if (!r.civilian && !String(r.rank ?? '').trim()) add('warn', '6-4c', 'Add the rank of the person acknowledging receipt.', 'receipt');
      if (!memo.suspense) add('tip', '2-4a(4)', 'No suspense date set. Add one only if there is a compelling reason to need it back by a certain date.', 'suspense');
    }

    /* rosters. AR 25-50 has no table format; these checks cover what it does say about the content and the enclosures. */
    const typedEncl = splitLines(memo.encl?.items).length, enclTotal = enclosureItems(memo).length;
    let enclK = 0;
    (memo.rosters || []).forEach((r, i) => {
      const T = rosterTable(r), target = `roster:${r.id}`, name = (memo.rosters.length > 1 ? `Roster ${i + 1}` : 'The roster');
      if (!T) { add('fix', 'MemoGen', `${name} has no column headings, so it is left off the memo.`, target); return; }
      const cols = (r.cols || []).map((c) => String(c ?? '').trim());
      cols.forEach((c, k) => { if (!c) add('fix', 'MemoGen', `${name}: column ${k + 1} has no heading.`, target); });
      if (!T.rows.length) add('fix', 'MemoGen', `${name} has no rows yet.`, target);
      // T.rows skip empty rows and may lead with a No. column, so map each heading to its place in T.rows.
      const shift = T.numbered ? 1 : 0;
      const rowsOf = (xs) => (xs.length <= 4 ? xs.map((x) => x.n).join(', ') : `${xs.slice(0, 3).map((x) => x.n).join(', ')} and ${xs.length - 3} more`);
      const rowWord = (xs) => `row${xs.length === 1 ? '' : 's'} ${rowsOf(xs)}`;
      const digits = (v) => v.replace(/[\s-]/g, '');
      let cui = '';
      cols.forEach((c, k) => {
        const label = c ? `the ${c} column` : `column ${k + 1}`;
        const vals = T.rows.map((row, n) => ({ n: n + 1, v: row[k + shift] || '' })).filter((x) => x.v);
        // 6-5b: SSNs only when a statute requires them, and then marked and protected like any Privacy Act document, truncated ones too.
        const ssnRows = vals.filter((x) => SSN_SHAPE.test(x.v) || SSN.test(x.v));
        const ssnHead = SSN_COL.test(c) || LAST4_COL.test(c);
        if (ssnHead || ssnRows.length) {
          const what = ssnHead ? `${label} holds Social Security numbers${LAST4_COL.test(c) ? ' (truncated)' : ''}` : `${label}, ${rowWord(ssnRows)}: ${ssnRows.length === 1 ? 'a 9-digit number looks' : '9-digit numbers look'} like Social Security numbers`;
          add('warn', '6-5b', `${name}: ${what}. Use SSNs only when a statute requires them; a memo that has them, even truncated, is marked and protected like any Privacy Act document.`, target);
        }
        // An identifying number by its heading, or by its values: mostly 7 or more digits, in a column that is not a phone or a date.
        const idLike = !NOT_ID_COL.test(c) && vals.length && vals.filter((x) => /^[A-Za-z0-9\s-]+$/.test(x.v) && digits(x.v).replace(/\D/g, '').length >= 7).length * 2 >= vals.length;
        if (!cui && (PII_COL.test(c) || ssnHead || ssnRows.length || idLike)) cui = label;
        if (DODID_COL.test(c)) {
          const bad = vals.filter((x) => !SSN_SHAPE.test(x.v) && !SSN.test(x.v) && !/^\d{10}$/.test(digits(x.v)));   // SSN-shaped ones are called out above
          if (bad.length) add('warn', 'MemoGen', `${name}: ${bad.length === 1 ? 'one DoD ID number is' : `${bad.length} DoD ID numbers are`} not 10 digits (${rowWord(bad)}). A DoD ID number has 10 digits.`, target);
        }
      });
      if (cui && !cuiOn(memo)) add('warn', '1-22d', `${name}: ${cui} holds personal identifying information. That makes this memo CUI: turn on CUI marking.`, 'cui');
      else if (cui && !(memo.cui.cats || []).some((k) => PRIVACY_CATS.includes(k))) add('warn', 'DoDI 5200.48', `${name}: ${cui} holds personal identifying information, but no Privacy category is listed. Add PRVCY (or PERS or MIL for personnel records).`, 'cui');
      if (T.place === 'encl') {
        enclK += 1;
        const n = typedEncl + enclK, said = enclTotal > 1 ? new RegExp(`\\b(enclosure|encl)\\s*${n}\\b`, 'i') : /\b(enclosure|encl)\b/i;
        if (!said.test(body)) add('tip', '1-34', `${T.title} is ${enclTotal > 1 ? `enclosure ${n}` : 'the enclosure'}. Mention it in the text${enclTotal > 1 ? ` as "enclosure ${n}"` : ''}.`, target);
      } else if (T.rows.length > 20) add('tip', '1-39b(7)', `${name} has ${T.rows.length} rows. A list that long reads better as an enclosure.`, target);
    });

    /* CUI marking: AR 25-50 para 1-22d points to DoDI 5200.48 (3.4, fig 2) */
    if (cuiOn(memo)) {
      const c = memo.cui, cats = c.cats || [];
      if (!usesLetterhead(memo) && !String(c.component ?? '').trim()) add('fix', 'DoDI 5200.48', 'On plain paper the designation indicator names your DoD Component first ("Controlled by: Department of the Army").', 'cui');
      if (!String(c.office ?? '').trim()) add('fix', 'DoDI 5200.48', 'Add the office that controls this CUI ("Controlled by").', 'cui');
      if (!cats.length) add('fix', 'DoDI 5200.48', 'List the CUI category, such as PRVCY.', 'cui');
      cats.filter((k) => !CUI_CATEGORIES[k]).forEach((k) => add('warn', 'DoDI 5200.48', `"${k}" is not in the DoD CUI Registry list. Use the registry's abbreviation.`, 'cui'));
      if (!cuiPoc(memo)) add('fix', 'DoDI 5200.48', 'Add a POC name and phone number or office email for the designation indicator.', 'cui');
      if (c.ldc === 'REL TO' && !String(c.rel ?? '').trim()) add('fix', 'DoDI 5200.48', 'List who it is releasable to after REL TO USA, for example FVEY.', 'cui');
      if (c.ldc === 'DL ONLY') add('tip', 'DoDI 5200.48', 'DL ONLY: send the dissemination list with the memo.', 'cui');
      if (c.cover && !enclosureItems(memo).length) add('warn', 'DoDI 5200.48', 'Marked as a cover memo, but no enclosure is listed.', 'cui');
      if (!c.portion) add('tip', 'DoDI 5200.48', 'Portion marking is optional, but DoD recommends it: (CUI) or (U) on the subject, each paragraph and each roster.', 'cui');
      else if (!c.cover) {
        const any = c.subject === 'CUI' || (memo.paras || []).some((q) => q.mark === 'CUI' && String(q.text ?? '').trim()) || (memo.rosters || []).some((r) => rosterTable(r) && r.mark !== 'U');
        if (!any) add('warn', 'DoDI 5200.48', 'Portion marking is on, but nothing is marked (CUI). Mark the paragraphs that hold CUI.', 'cui');
      }
    }

    const encl = enclosureItems(memo);
    const mentions = /\benclosures?\b|\bencls?\b/i.test(body);
    if (mentions && !encl.length) add('warn', '2-4c(3)', 'The text mentions an enclosure, but none is listed.', 'encl');
    if (memo.encl?.inBody && encl.length && !mentions) add('warn', '4-2c(4)', 'Marked as named in the text, but the text does not mention an enclosure.', 'encl');

    /* length */
    const pages = (L?.pages || []).filter((p) => p.kind === 'first' || p.kind === 'cont').length || 1;   // enclosures do not count
    if (pages > 1) add('tip', '1-39b(7)', `This memo runs ${pages} pages. Most fit on one; consider moving detail to an enclosure.`, 'body');
    return out;
  }

  /* The rules this app applies, for the AR 25-50 tab. Own wording; the regulation governs. how: auto = done for you, check = flagged, you = your call. */
  const RULES = [
    ['Page setup', [
      ['8 1/2 by 11 inch paper; letterhead on page 1, plain white after', '1-18, 2-3a', 'auto'],
      ['1-inch left, right and bottom margins; right margin not justified', '2-3c', 'auto'],
      ['Font and size are set by your senior leader; 12 point recommended, no script styles', '1-19', 'you'],
      ['Black ink, DoD seal, nine-digit ZIP on letterhead', '1-16, 1-20, 1-35', 'check'],
    ]],
    ['Heading', [
      ['Office symbol on the second line below the seal, ARIMS number in parentheses after it', '2-4a(1)-(2)', 'auto'],
      ['Date on the office symbol line, flush right, as 13 March 2020 or 13 Mar 20', '2-4a(3)', 'auto'],
      ['Suspense date flush right, bold, two lines above the date, preceded by S:', '2-4a(4)', 'auto'],
      ['MEMORANDUM FOR on the third line below the office symbol', '2-4a(5)', 'auto'],
      ['Up to five addressees as multiple-address; more than five use SEE DISTRIBUTION', '2-4a(5)(b)-(c)', 'check'],
      ['Exclusive For: sensitive matters to one person, by name, title and mailing address, in a sealed envelope', '1-12', 'check'],
      ['SUBJECT on the second line below the address; 10 words or fewer, one subject', '2-4a(6)', 'check'],
    ]],
    ['Body', [
      ['Text begins on the third line below the subject', '2-4b(1)', 'auto'],
      ['Single spaced, double spaced between paragraphs and subparagraphs', '2-4b(2)', 'auto'],
      ['1. then a. then (1) then (a); indent 1/4 inch, then 1/2 inch, no further; wrapped lines return to the margin', 'fig 2-1', 'auto'],
      ['One-paragraph memo is not numbered', '2-4b(4)(a)', 'auto'],
      ['If there is an "a" there must be a "b"', 'fig 2-1', 'check'],
      ['Two spaces after periods and question marks; one after commas, colons and semicolons', '1-39b(9)', 'auto'],
      ['1/4 inch after the parenthesis in numbered subparagraphs', '1-39b(10)', 'auto'],
      ['Keep day, month and year of a date together', '1-25c', 'auto'],
      ['Point of contact goes in the last paragraph', '1-23, 2-4b(1)(e)', 'check'],
      ['Correct spelling (common errors are flagged; the browser underlines the rest)', '1-39b(5)', 'check'],
      ['Set up an abbreviation only if you use it again', '1-15c(2)', 'check'],
      ['Short sentences (about 15 words), paragraphs of 10 lines or fewer, active voice', '1-38, 1-39b', 'check'],
      ['Every [bracketed] fill-in from a template is replaced before export (Alt+N goes to the next)', 'MemoGen', 'check'],
    ]],
    ['Closing', [
      ['Authority line at the left margin, uppercase, second line below the text', '2-4c(1)', 'auto'],
      ['Signature block at page center, fifth line below the authority line or text', '2-4c(2), 6-4b', 'auto'],
      ['Name in capitals; rank and branch; title; long titles continue 1/4 inch in', '6-4a, 6-4c', 'auto'],
      ['General officers use their full grade', '6-4f(3)', 'check'],
      ['Enclosures at the left margin on the signature line; Encl, or 2 Encls with a numbered list', '2-4c(3), 4-2', 'auto'],
      ['DISTRIBUTION and CF on the second line below whichever block is lower', '2-4a(5)(c), 2-4c(5)', 'auto'],
    ]],
    ['More than one page', [
      ['Continuation pages: office symbol 1 inch from the top, subject on the next line, text on the third line below', '2-5a-c', 'auto'],
      ['Never split a paragraph of three lines or fewer; at least two lines on each side of a break', '2-5c(1)', 'auto'],
      ['At least two words of a divided sentence on each page; no word hyphenated across pages', '2-5c(2)-(3)', 'auto'],
      ['Authority line and signature block never alone on a page: two lines of the last paragraph go with them', '2-5c(4)', 'auto'],
      ['Page number centered about 1 inch from the bottom', '2-5d', 'auto'],
    ]],
    ['Second signature', [
      ['More than one signature block is allowed; follow figures 2-15 and 2-16', '2-7b(3), 2-8b(3), app F-2h', 'auto'],
      ['Senior official on the right; with three, highest centered at the bottom, junior on the left', '2-6c(5)(d)', 'auto'],
      ['All blocks overscored, with the date each official signs', '2-6c(5)(b)-(c)', 'auto'],
      ['Signature boxes for all signers; the date at the top goes with the last signature', 'app F-2h', 'auto'],
      ['For a chain endorsing an action in turn, the regulation\'s route is MEMORANDUM THRU', '2-4a(5)(d)', 'you'],
      ['Omit the authority line when the commander signs personally', '6-2b', 'check'],
    ]],
    ['Receipt acknowledgment (AR 25-50 has no format for this; these rules are the closest it offers)', [
      ['Second signature block on the same line, junior signer on the left, overscored, with a date under it', '2-6c(5), 2-7b(3), figs 2-15, 2-16', 'auto'],
      ['Acknowledgment sentence in the body, never after the signature', '1-29', 'auto'],
      ['POC stays in the last paragraph, so the sentence goes just before it', '2-4b(1)(e)', 'auto'],
      ['Encl moves two lines below the lowest signature block', 'fig 2-14', 'auto'],
      ['A return deadline is a suspense date, shown at the top and in the body', '2-4a(4), 1-27a', 'auto'],
      ['Signature boxes for all signers are placed before anyone signs', 'app F-2h', 'auto'],
    ]],
    ['Rosters (AR 25-50 has no table format; where it is silent, MemoGen keeps the memo\'s own font and spacing)', [
      ['In the body: full text width, a blank line before and after, bold heading row repeated on each page', '2-4b(2)', 'auto'],
      ['Rows never split; three rows or fewer stay together, otherwise at least two rows on each side of a page break', '2-5c(1)', 'auto'],
      ['As an enclosure: own pages after the memo, numbered in the order the text mentions it', '1-34, 4-2', 'auto'],
      ['Encl and its number in the lower right corner of the enclosure\'s first page', '1-34, 4-2d(1)', 'auto'],
      ['Identifying numbers (DoD ID, SSN, service number or any long ID number, found by heading or by value) make the memo CUI; mark it as DoDI 5200.48 requires', '1-22d', 'check'],
      ['SSNs, even truncated, only when a statute requires them, and then marked and protected as Privacy Act information', '6-5b', 'check'],
      ['Templates keep a roster\'s columns, never its rows', 'MemoGen', 'auto'],
    ]],
    ['CUI marking (para 1-22d points to DoDI 5200.48; details from the DoD CUI Marking aid, December 2024)', [
      ['"CUI" at the top and bottom of every page; no category, no UNCLASSIFIED in the banner', 'DoDI 5200.48 3.4a', 'auto'],
      ['Designation indicator on the first page, lower right: Controlled by (Component, left off on letterhead), Controlled by (office), CUI Category, dissemination control, POC', 'DoDI 5200.48 3.4f, fig 2', 'auto'],
      ['Category abbreviations from the DoD CUI Registry; a dissemination control only when one applies', 'DoDI 5200.48 fig 2', 'check'],
      ['Portion marks optional but recommended; if used, every subject, paragraph and table gets (CUI) or (U)', 'DoDI 5200.48 3.4b', 'you'],
      ['No portion mark on the signature block or the designation indicator', 'DoD CUI Marking aid', 'auto'],
      ['Banner 1/4 inch from the edge and a 10 point designation indicator: MemoGen\'s choice where the sources are silent', 'MemoGen', 'auto'],
    ]],
    ['Digital signature', [
      ['Internal correspondence should be digitally signed with a CAC', '1-17, 6-4e', 'auto'],
      ['Date text box across from the office symbol; signature box above the name', 'app F-2e, F-2f', 'auto'],
    ]],
  ];

  global.MemoChecks = { check, RULES };
})(window);
