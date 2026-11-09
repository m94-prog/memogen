/* MemoGen: Army memorandums to AR 25-50, filled into the unit's own Word template.
 *
 * What is kept and what is not. The memo on the form lives only in this page's memory: typing writes nothing to the browser, to a
 * file or anywhere else, and closing the page clears it. The only things stored are the saved items (lib below): templates, POCs,
 * signature blocks, letterhead and format, and each is written only when the user clicks a Save button. */
(() => {
  const LIB_KEY = 'memogen.library.v1';
  const { layout, numbered, bodyParas, pocText, receiptText, subjectText, plain, usesLetterhead, rosterTable, splitLines, portionOn, cuiPoc, CUI_CATEGORIES, CUI_LDCS, FONTS, KINDS, fillIns, FILL_RE } = MemoFormat;
  const { check, RULES } = MemoChecks;
  const $ = (s, r = document) => r.querySelector(s);
  const $$ = (s, r = document) => [...r.querySelectorAll(s)];
  const uid = () => Math.random().toString(36).slice(2, 9) + Date.now().toString(36).slice(-3);
  const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const ico = (n) => `<svg class="ic" aria-hidden="true" focusable="false"><use href="#i-${n}"></use></svg>`;
  const isTyping = (el) => !!el && (el.tagName === 'TEXTAREA' || el.tagName === 'SELECT' || (el.tagName === 'INPUT' && !['checkbox', 'radio', 'button'].includes(el.type)));
  const download = (blob, name) => { const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.append(a); a.click(); a.remove(); setTimeout(() => URL.revokeObjectURL(a.href), 30000); };

  /* ---------------- saved items ---------------- */
  // What Docs/MEMO TEMPLATE.docx carries, as the starting letterhead. "BATALLION" in the template is corrected here.
  const TEMPLATE_OFFICE = {
    name: 'Generic letterhead', headline: 'DEPARTMENT OF THE ARMY', symbol: '[OFFICE SYMBOL]', seal: 'seal.png',
    lines: '[ORGANIZATION NAME]\n[STREET ADDRESS]\n[CITY, STATE ZIP]',
  };
  const DEFAULTS = { officeId: '', font: 'arial', size: 12, spaces: 2, kind: 'mfr', digital: true };
  const blankSigner = () => ({ name: '', rank: '', branch: '', component: '', title: '', org: '', civilian: false });
  const blankPoc = () => ({ name: '', phone: '', email: '' });
  const para = (level = 0, text = '') => ({ id: uid(), level, text });
  function freshLib() {
    const office = { id: uid(), ...TEMPLATE_OFFICE };
    return { version: 1, templates: [], pocs: [], signers: [], offices: [office], defaults: { ...DEFAULTS, officeId: office.id } };
  }
  // Ids are written into HTML attributes unescaped, so ids from a restored backup or the file keep only letters, digits, _ and -.
  // The same cleanup on every copy keeps references between items (para "after", officeId) matching.
  const ID_KEYS = new Set(['id', 'after', 'officeId']);
  function cleanIds(v) {
    if (Array.isArray(v)) v.forEach(cleanIds);
    else if (v && typeof v === 'object') for (const k of Object.keys(v)) {
      if (ID_KEYS.has(k) && typeof v[k] === 'string' && /[^\w-]/.test(v[k])) v[k] = v[k].replace(/[^\w-]/g, '') || uid();
      else cleanIds(v[k]);
    }
    return v;
  }
  function repairLib(s) {
    cleanIds(s);
    for (const k of ['templates', 'pocs', 'signers']) s[k] = Array.isArray(s[k]) ? s[k] : [];
    if (!Array.isArray(s.offices) || !s.offices.length) s.offices = [{ id: uid(), ...TEMPLATE_OFFICE }];
    s.defaults = { ...DEFAULTS, officeId: s.offices[0].id, ...(s.defaults || {}) };
    if (!s.offices.some((o) => o.id === s.defaults.officeId)) s.defaults.officeId = s.offices[0].id;
    s.version = 1;
    return s;
  }
  let lib = (() => { try { const s = JSON.parse(localStorage.getItem(LIB_KEY)); if (s && Array.isArray(s.offices)) return repairLib(s); } catch { /* start clean */ } return freshLib(); })();
  const fmt = () => ({ font: lib.defaults.font, size: +lib.defaults.size, spaces: +lib.defaults.spaces });
  // Templates that come with the page (builtin-templates.js, filled only in the public copy). They are never saved or changed:
  // Save as template makes your own copy, and a saved template with the same name takes the built-in one's place in the list.
  const BUILTINS = cleanIds((Array.isArray(window.MemoGenBuiltins) ? window.MemoGenBuiltins : []).filter((t) => t && t.memo && typeof t.name === 'string')
    .map((t) => ({ id: `builtin-${t.id}`, name: t.name, memo: t.memo })));
  const builtinsShown = () => BUILTINS.filter((b) => !lib.templates.some((t) => t.name.toLowerCase() === b.name.toLowerCase()));
  const tplById = (id) => (id && (lib.templates.find((t) => t.id === id) || builtinsShown().find((t) => t.id === id))) || null;

  /* ---------------- the memo on the form: memory only ---------------- */
  function newMemo(o = {}) {
    const d = lib.defaults, office = lib.offices.find((x) => x.id === d.officeId) || lib.offices[0];
    return {
      kind: d.kind, letterhead: false, officeId: office.id, officeSymbol: office.symbol || '', record: '', date: '', dateStyle: 'long', suspense: '',
      to: '', thru: '', distribution: '', subject: '', subjectName: '', paras: [para()], poc: blankPoc(),
      authority: '', signer: blankSigner(), cosign: { on: false, signer: blankSigner() }, receipt: { on: false, text: '', signer: blankSigner() }, digital: d.digital, encl: { items: '', inBody: false }, cf: '', cfNoEncl: false, rosters: [],
      cui: { on: false, component: 'Department of the Army', office: '', cats: [], ldc: '', rel: '', poc: '', portion: false, subject: 'U', cover: false }, ...o,
    };
  }
  let memo = newMemo();
  let activeTpl = '';          // the saved template this form was started from, if any
  let exported = true;         // nothing on the form that has not been exported
  const officeOf = (m) => lib.offices.find((o) => o.id === m.officeId) || lib.offices[0];
  const memoHasText = (m) => !!(String(m.subject ?? '').trim() || (m.paras || []).some((p) => String(p.text ?? '').trim()) || (m.rosters || []).some((r) => (r.rows || []).some((row) => row.some((c) => String(c ?? '').trim()))));
  const unexported = () => !exported && memoHasText(memo);
  window.addEventListener('beforeunload', (e) => { if (unexported()) { e.preventDefault(); e.returnValue = ''; } });

  /* ---------------- storing saved items ----------------
   * Only saveLib() writes, and only the Save buttons call it. Saved items go to this browser and to memogen-data.json through the
   * local helper (memogen-server.ps1, 127.0.0.1 only). When the two disagree the newer one wins and the other is kept aside. */
  const HELPER = 'http://127.0.0.1:47618';
  const SERVER_KEY = 'memogen.server';            // 'off' = browser only
  const CONFLICT_KEY = 'memogen.library.conflict'; // + '.<ms>': the copy that lost a disagreement (newest 5 are kept)
  // The helper is reached only from a page on this PC. A hosted copy (for example on github.io) keeps saved items in its browser
  // only and never calls 127.0.0.1: the helper would refuse it, and the browser would ask about local network access.
  const LOCAL_PAGE = location.origin === HELPER || /^(127\.0\.0\.1|localhost|\[::1\])$/.test(location.hostname) || location.protocol === 'file:';
  const SYNC = { base: (() => { let o = null; try { o = localStorage.getItem(SERVER_KEY); } catch { /* blocked */ } if (o === 'off' || !LOCAL_PAGE) return 'off'; return location.origin === HELPER ? '' : HELPER; })(), kind: 'idle', pending: false, damaged: false, at: 0, checked: 0 };
  const hasItems = (s) => !!(s.templates?.length || s.pocs?.length || s.signers?.length || s.savedAt);
  const validFile = (d) => !!(d && typeof d === 'object' && Array.isArray(d.templates) && Array.isArray(d.offices));
  const canon = (v) => Array.isArray(v) ? '[' + v.map(canon).join(',') + ']' : v && typeof v === 'object' ? '{' + Object.keys(v).sort().filter((k) => k !== 'savedAt').map((k) => JSON.stringify(k) + ':' + canon(v[k])).join(',') + '}' : JSON.stringify(v);
  const when = (ms) => (ms ? new Date(ms).toLocaleString([], { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }) : 'an earlier save');
  async function fetchJ(method, path, body, ms = 2500) {
    const ac = new AbortController(), to = setTimeout(() => ac.abort(), ms);
    try { return await fetch(SYNC.base + path, { method, cache: 'no-store', signal: ac.signal, headers: body != null ? { 'Content-Type': 'application/json' } : undefined, body }); }
    finally { clearTimeout(to); }
  }
  const PILL = {
    idle: { text: 'Memo text is never saved', icon: 'eye-off', cls: 'idle' },
    file: { text: 'Saved items in file', icon: 'hard-drive', cls: 'ok' },
    browser: { text: 'Saved items in browser only', icon: 'save', cls: 'warn' },
    local: { text: 'Saved items in this browser', icon: 'save', cls: 'ok' },
    bad: { text: 'Could not save', icon: 'triangle-alert', cls: 'bad' },
  };
  const pillTip = (k) => ({
    idle: 'Nothing you type is stored. Only what you save with a Save button is kept.',
    file: 'Your templates, POCs, signature blocks, letterhead and format are in memogen-data.json and this browser. Memo text is never saved.',
    browser: SYNC.damaged ? 'memogen-data.json looks damaged, so it was left alone. Your saved items are in this browser.' : 'The save helper is not running, so your saved items are in this browser only. Open MemoGen with launch.cmd and they are copied to memogen-data.json. Memo text is never saved.',
    local: 'Your templates, POCs, signature blocks, letterhead and format are kept in this browser only, never sent anywhere. To use them on another computer, choose Export, Back up saved items, then Restore saved items there. Memo text is never saved.',
    bad: SYNC.base === 'off' ? 'The last save did not work: this browser could not store it. Try again, or download a backup from Export.' : 'The last save did not work: this browser could not store it and the file was not reached. Try again, or download a backup from Export.',
  }[k]);
  function renderPill(k = SYNC.kind) {
    SYNC.kind = k;
    const el = $('#saveState'), p = PILL[k], tip = pillTip(k);
    el.className = 'syncbtn savepill ' + p.cls; el.dataset.state = k;
    $('.stext', el).textContent = p.text; $('use', el).setAttribute('href', '#i-' + p.icon);
    el.title = tip; el.setAttribute('aria-label', p.text + '. ' + tip);
    $('#saveLive').textContent = p.text;
  }
  async function pushFile(force) {
    if (SYNC.base === 'off') return false;
    try {
      const r = await fetchJ('PUT', '/api/data' + (force ? '?force=1&backup=1' : ''), JSON.stringify(lib));
      if (r.ok) { SYNC.at = Date.now(); SYNC.damaged = false; SYNC.pending = false; return true; }
      if (r.status === 409) { const cur = await r.json().catch(() => null); if (validFile(cur)) { await reconcile(cur); return true; } }
    } catch { /* helper not running */ }
    return false;
  }
  async function saveLib() {
    lib.savedAt = Math.max(Date.now(), (lib.savedAt || 0) + 1);
    let local = true;
    try { localStorage.setItem(LIB_KEY, JSON.stringify(lib)); } catch (e) { console.warn('MemoGen save failed', e); local = false; }
    SYNC.pending = true;
    const file = await pushFile();
    renderPill(file ? 'file' : !local ? 'bad' : SYNC.base === 'off' ? 'local' : 'browser');
    return file || local;
  }
  function keepAside(s, why) {
    const key = CONFLICT_KEY + '.' + Date.now();
    try {
      localStorage.setItem(key, JSON.stringify({ why, at: Date.now(), savedAt: s.savedAt || 0, data: s }));
      const keys = Object.keys(localStorage).filter((k) => k.startsWith(CONFLICT_KEY + '.')).sort();
      keys.slice(0, Math.max(0, keys.length - 5)).forEach((k) => localStorage.removeItem(k));
      return key;
    } catch { return null; }
  }
  function adopt(file) {
    lib = repairLib(structuredClone(file));
    try { localStorage.setItem(LIB_KEY, JSON.stringify(lib)); } catch { /* the file has it */ }
    SYNC.pending = false; SYNC.at = Date.now();
    renderLibrary();
  }
  let noticeAlt = null;
  function notice(msg, alt) {
    $('#syncMsg').textContent = msg;
    $('#syncAlt').hidden = !alt; noticeAlt = alt ? alt.fn : null; if (alt) $('#syncAltLabel').textContent = alt.label;
    $('#syncBanner').hidden = false;
  }
  async function useAside(key) {
    let e = null; try { e = JSON.parse(localStorage.getItem(key)); } catch { /* gone */ }
    if (!e?.data) { toast('That copy is no longer there.', true); return; }
    keepAside(lib, 'swapped'); lib = repairLib(e.data); renderLibrary(); await saveLib();
    toast('Switched to the other copy.');
  }
  async function reconcile(file) {
    const localAt = lib.savedAt || 0, fileAt = file.savedAt || 0;
    SYNC.damaged = false;
    if (canon(lib) === canon(repairLib(structuredClone(file)))) { lib.savedAt = lib.savedAt || fileAt; SYNC.pending = false; return; }
    if (!hasItems(lib)) { adopt(file); return; }                       // a new or cleared browser: the file has the saved items
    if (localAt > fileAt) { await pushFile(true); return; }            // this browser saved last (the helper backs the old file up first)
    const k = keepAside(lib, 'older'); adopt(file);
    notice(`memogen-data.json had newer saved items, so those were loaded. This browser's copy (from ${when(localAt)}) is kept aside.`, k ? { label: 'Use my browser copy instead', fn: () => useAside(k) } : null);
  }
  async function initSync() {
    SYNC.checked = Date.now();
    if (SYNC.base === 'off') { renderPill(hasItems(lib) ? 'local' : 'idle'); return; }
    let r;
    try { r = await fetchJ('GET', '/api/data'); } catch { renderPill(hasItems(lib) ? 'browser' : 'idle'); return; }
    if (r.status === 404) { const ok = hasItems(lib) && await pushFile(); renderPill(ok ? 'file' : hasItems(lib) ? 'browser' : 'idle'); return; }
    if (!r.ok) { renderPill(hasItems(lib) ? 'browser' : 'idle'); return; }
    const d = await r.json().catch(() => null);
    if (!validFile(d)) {
      SYNC.damaged = true;
      notice('memogen-data.json is there but unreadable, so it was left alone. Your saved items are safe in this browser. The next Save replaces the file, and the damaged one is backed up first.');
      renderPill('browser'); return;
    }
    await reconcile(d); renderPill('file');
  }
  $('#saveState').addEventListener('click', async () => {
    const k = SYNC.kind, retry = SYNC.base !== 'off' && (k === 'browser' || k === 'bad');
    const go = await ask({ title: PILL[k].text, message: pillTip(k) + (k === 'file' || retry ? ' The last 10 versions of the file are kept in data-backups.' : ''), okLabel: retry ? 'Try the file again' : 'OK' });
    if (!go || !retry) return;
    await initSync();
    if (SYNC.kind !== 'file' && hasItems(lib)) { const ok = await pushFile(SYNC.damaged); if (ok) renderPill('file'); }
    toast(SYNC.kind === 'file' ? 'Saved items are in memogen-data.json.' : 'Could not reach the save helper. Open MemoGen with launch.cmd.', SYNC.kind !== 'file');
  });
  $('#syncAlt').addEventListener('click', () => { const f = noticeAlt; $('#syncBanner').hidden = true; if (f) f(); });
  $('#syncDismiss').addEventListener('click', () => { $('#syncBanner').hidden = true; });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'visible' && Date.now() - SYNC.checked > 3000 && !document.querySelector('dialog[open]')) initSync(); });
  setInterval(() => { if (document.visibilityState === 'visible' && SYNC.pending) initSync(); }, 15000);   // a save made while the helper was down

  /* ---------------- toast + undo ---------------- */
  let toastT, toastHide;
  function hideToast() { const t = $('#toast'); t.classList.remove('in'); clearTimeout(toastHide); toastHide = setTimeout(() => { t.hidden = true; }, 200); }
  function toast(msg, o = {}) {
    if (o === true) o = { err: true };
    const t = $('#toast'); $('#toastMsg').textContent = msg; $('#toastUndo').hidden = !o.undo;
    t.classList.toggle('err', !!o.err); t.setAttribute('role', o.err ? 'alert' : 'status');
    clearTimeout(toastHide); t.hidden = false; requestAnimationFrame(() => t.classList.add('in'));
    clearTimeout(toastT); toastT = setTimeout(hideToast, o.undo ? 8000 : o.err ? 7000 : 3500);
  }
  $('#toastX').addEventListener('click', hideToast);
  const undoStack = [];
  function pushUndo(label, fn) { undoStack.push({ label, fn }); if (undoStack.length > 30) undoStack.shift(); toast(label, { undo: true }); }
  async function undo() {
    const u = undoStack.pop();
    if (!u) { toast('Nothing to undo.'); return; }
    try { await u.fn(); } catch (e) { console.warn(e); toast('Could not undo that.', true); return; }
    renderLibrary(); renderCompose(); toast(`Undone: ${u.label.replace(/^(Deleted|Removed) /, 'delete of ')}`);
  }
  $('#toastUndo').addEventListener('click', undo);

  /* ---------------- in-app confirm / prompt (native ones are silently dismissed in embedded panes) ---------------- */
  function ask({ title, message = '', input = null, okLabel = 'OK', cancelLabel = 'Cancel', danger = false }) {
    return new Promise((resolve) => {
      const d = $('#askDialog'), inp = $('#askInput'), ok = $('#askOk');
      $('#askTitle').textContent = title; $('#askMsg').textContent = message; $('#askMsg').hidden = !message;
      inp.hidden = input === null; inp.value = input ?? '';
      ok.textContent = okLabel; ok.classList.toggle('danger', danger); $('#askCancel').textContent = cancelLabel;
      const onClose = () => { d.removeEventListener('close', onClose); const yes = d.returnValue === 'ok'; resolve(input === null ? yes : (yes ? inp.value : null)); };
      d.addEventListener('close', onClose);
      d.returnValue = 'cancel'; d.showModal();
      if (input !== null) { inp.focus(); inp.select(); }
    });
  }
  document.addEventListener('click', (e) => { const c = e.target.closest('[data-cancel]'); if (c) c.closest('dialog')?.close('cancel'); });

  /* ---------------- views / tabs ---------------- */
  function moveInd() {
    const list = $('#tabs'), ind = list && $('.tabind', list), t = list && $('[role=tab][aria-selected=true]', list);
    if (!ind || !t || !list) return;
    const lr = list.getBoundingClientRect(), tr = t.getBoundingClientRect();
    ind.style.transform = 'none';
    ind.style.left = `${tr.left - lr.left + list.scrollLeft}px`;
    ind.style.width = `${tr.width}px`;
  }
  function showView(name) {
    $$('[role=tab]').forEach((t) => { const on = t.dataset.view === name; t.setAttribute('aria-selected', on); t.tabIndex = on ? 0 : -1; });
    $$('.view').forEach((v) => { const on = v.id === 'view-' + name; v.classList.toggle('active', on); if (on) { v.classList.remove('enter'); void v.offsetWidth; v.classList.add('enter'); } });
    moveInd();
    if (name === 'saved') renderSaved();
    if (name === 'compose') sizeSheets();
    window.scrollTo(0, 0);
  }
  $$('[role=tab]').forEach((t) => t.addEventListener('click', () => showView(t.dataset.view)));
  $('#tabs').addEventListener('keydown', (e) => {
    const tabs = $$('[role=tab]'), i = tabs.indexOf(document.activeElement); if (i < 0) return;
    const to = { ArrowRight: (i + 1) % tabs.length, ArrowLeft: (i + tabs.length - 1) % tabs.length, Home: 0, End: tabs.length - 1 }[e.key];
    if (to === undefined) return; e.preventDefault(); tabs[to].focus(); showView(tabs[to].dataset.view);
  });
  window.addEventListener('resize', moveInd);
  const setStick = () => document.documentElement.style.setProperty('--stick-top', ($('.topwrap')?.getBoundingClientRect().height || 0) + 'px');
  new ResizeObserver(setStick).observe($('.topwrap'));

  /* menus */
  function closeMenus(except) { $$('.menu').forEach((m) => { if (m === except) return; m.classList.remove('open'); $('.menu-body', m).hidden = true; $('[aria-haspopup]', m).setAttribute('aria-expanded', 'false'); }); }
  function openMenu(m, focusFirst) { closeMenus(m); m.classList.add('open'); $('.menu-body', m).hidden = false; $('[aria-haspopup]', m).setAttribute('aria-expanded', 'true'); if (focusFirst) $('[role=menuitem]', m).focus(); }
  $$('.menu').forEach((m) => {
    const trg = $('[aria-haspopup]', m), body = $('.menu-body', m);
    trg.addEventListener('click', (e) => { e.stopPropagation(); if (m.classList.contains('open')) closeMenus(); else openMenu(m, false); });
    trg.addEventListener('keydown', (e) => { if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); openMenu(m, true); } });
    body.addEventListener('keydown', (e) => {
      const items = $$('[role=menuitem]', body), i = items.indexOf(document.activeElement);
      if (e.key === 'ArrowDown') { e.preventDefault(); items[(i + 1) % items.length].focus(); }
      else if (e.key === 'ArrowUp') { e.preventDefault(); items[(i + items.length - 1) % items.length].focus(); }
      else if (e.key === 'Escape') { e.preventDefault(); closeMenus(); trg.focus(); }
      else if (e.key === 'Tab') closeMenus();
    });
    body.addEventListener('click', () => closeMenus());
  });
  document.addEventListener('click', () => closeMenus());

  /* ---------------- compose: fields ---------------- */
  const getPath = (o, p) => p.split('.').reduce((a, k) => a?.[k], o);
  const setPath = (o, p, v) => { const ks = p.split('.'), last = ks.pop(); ks.reduce((a, k) => (a[k] ||= {}), o)[last] = v; };
  const fieldVal = (el) => (el.type === 'checkbox' ? el.checked : el.value);
  function fillField(el, v) { if (el.type === 'checkbox') el.checked = !!v; else if (el.value !== String(v ?? '')) el.value = v ?? ''; }

  function renderCompose() {
    $('#m_kind').innerHTML = Object.entries(KINDS).map(([k, v]) => `<option value="${k}">${esc(v.label)}</option>`).join('');
    $('#m_office').innerHTML = lib.offices.map((o) => `<option value="${o.id}">${esc(o.name || 'Letterhead')}</option>`).join('');
    if (!lib.offices.some((o) => o.id === memo.officeId)) memo.officeId = officeOf(memo).id;
    $$('[data-m]').forEach((el) => fillField(el, getPath(memo, el.dataset.m)));
    renderTemplatePick(); renderQuick(); renderParas(); renderRosters(); renderCuiCats(); applyKind(); renderPreview();
  }
  /* CUI marking card: the lists come from the DoD CUI Registry and the LDC table on dodcui.mil (memo-format.js) */
  const COMMON_CUI = ['PRVCY', 'PERS', 'MIL', 'HLTH', 'OPSEC', 'PHYS', 'LEI', 'INV'];
  $('#m_cuiLdc').innerHTML = Object.entries(CUI_LDCS).map(([k, v]) => `<option value="${esc(k)}">${esc(k ? `${k}: ${v}` : v)}</option>`).join('');
  $('#cuiCatList').innerHTML = Object.entries(CUI_CATEGORIES).map(([k, v]) => `<option value="${esc(k)}">${esc(v)}</option>`).join('');
  $('#cuiCommon').innerHTML = `<span>Common:</span>${COMMON_CUI.map((k) => `<button type="button" class="b" data-cat="${k}" title="${esc(CUI_CATEGORIES[k])}">${k}</button>`).join('')}`;
  function renderCuiCats() {
    $('#cuiCats').innerHTML = (memo.cui.cats || []).map((k) => `<button type="button" class="qchip" data-cat-del="${esc(k)}" aria-pressed="true" title="Remove ${esc(k)}">${ico('x')}<span>${esc(k)}${CUI_CATEGORIES[k] ? ` · ${esc(CUI_CATEGORIES[k])}` : ''}</span></button>`).join('');
  }
  function addCuiCat(raw) {
    const t = String(raw ?? '').trim(); if (!t) return;
    // Accept the abbreviation or the category's name; keep the registry's spelling.
    const k = Object.keys(CUI_CATEGORIES).find((a) => a.toLowerCase() === t.toLowerCase() || CUI_CATEGORIES[a].toLowerCase() === t.toLowerCase()) || t.toUpperCase();
    if (!memo.cui.cats.includes(k)) memo.cui.cats.push(k);
    $('#m_cuiCat').value = ''; renderCuiCats(); touched();
    if (!CUI_CATEGORIES[k]) toast(`${k} is not in the DoD CUI Registry list. Check the abbreviation.`, true);
  }
  $('#cuiCatAdd').addEventListener('click', () => addCuiCat($('#m_cuiCat').value));
  $('#m_cuiCat').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addCuiCat(e.target.value); } });
  $('#m_cuiCat').addEventListener('change', (e) => { if (CUI_CATEGORIES[e.target.value]) addCuiCat(e.target.value); });   // picked from the list
  $('#cuiCommon').addEventListener('click', (e) => { const b = e.target.closest('[data-cat]'); if (b) addCuiCat(b.dataset.cat); });
  $('#cuiCats').addEventListener('click', (e) => {
    const b = e.target.closest('[data-cat-del]'); if (!b) return;
    memo.cui.cats = memo.cui.cats.filter((k) => k !== b.dataset.catDel); renderCuiCats(); touched();
  });
  function applyKind() {
    $('#kindHint').textContent = KINDS[memo.kind]?.hint || '';
    $$('.only').forEach((el) => { el.hidden = !el.dataset.kinds.split(' ').includes(memo.kind); });
    $$('.lh-only').forEach((el) => { el.hidden = !usesLetterhead(memo); });
    $('#toLabel').textContent = memo.kind === 'thru' ? 'FOR' : memo.kind === 'exclusive' ? 'Memorandum Exclusive For' : 'MEMORANDUM FOR';
    $('#toHint').textContent = memo.kind === 'multi' ? 'One address per line, up to five.' : memo.kind === 'exclusive' ? 'Full name, title, mailing address. Use "Commander of" and the unit name to address a commander.' : 'Office or person expected to act, with the address if needed.';
    $$('.field.mil').forEach((el) => { el.hidden = !!memo.signer.civilian; });
    $$('.field.rmil').forEach((el) => { el.hidden = !!memo.receipt.signer.civilian; });
    $('#receiptFields').hidden = !memo.receipt.on; $('#saveReceipt').hidden = !memo.receipt.on;
    $$('.field.cmil').forEach((el) => { el.hidden = !!memo.cosign.signer.civilian; });
    $('#cosignFields').hidden = !memo.cosign.on; $('#saveCosign').hidden = !memo.cosign.on;
    const c = memo.cui;
    $('#cuiFields').hidden = !c.on;
    $('#cuiCompField').hidden = usesLetterhead(memo);
    $('#cuiRelField').hidden = c.ldc !== 'REL TO';
    $('#cuiSubjField').hidden = !c.portion;
    const p = cuiPoc({ ...memo, cui: { ...c, poc: '' } });
    $('#cuiPocHint').textContent = p ? `Leave blank to use the point of contact: ${p}.` : 'Name, and a phone number or office email (DoDI 5200.48 fig 2).';
  }
  function subjectHint() {
    const words = plain(subjectText(memo)).trim().split(/\s+/).filter(Boolean).length;
    const el = $('#subjectCount'); el.textContent = `${words} of 10 words`; el.classList.toggle('over', words > 10);
  }
  function pocHint() { $('#pocPreview').textContent = pocText(memo.poc) || 'Added as the last paragraph.'; $('#receiptPreview').textContent = receiptText(memo); }

  let rt = 0;
  const soon = () => { clearTimeout(rt); rt = setTimeout(renderPreview, 90); };
  function touched() { exported = false; soon(); }   // memory only: nothing is written anywhere
  $('#memoForm').addEventListener('submit', (e) => e.preventDefault());
  $('#memoForm').addEventListener('input', (e) => {
    const el = e.target.closest('[data-m]'); if (!el) return;
    const key = el.dataset.m;
    setPath(memo, key, fieldVal(el));
    if (key === 'officeId') { const o = officeOf(memo); if (!memo.officeSymbol.trim()) { memo.officeSymbol = o.symbol || ''; $('#m_symbol').value = memo.officeSymbol; } }
    if (['kind', 'letterhead', 'signer.civilian', 'receipt.on', 'receipt.signer.civilian', 'cosign.on', 'cosign.signer.civilian'].includes(key) || key.startsWith('cui.') || key.startsWith('poc.')) applyKind();
    // A fill-in, not the letterhead's name: that name is only shown in this app ("Generic letterhead" on the public copy) and is
    // seldom the office that decided this is CUI, so the counter and export guard catch it until it is filled in.
    if (key === 'cui.on' && memo.cui.on && !String(memo.cui.office).trim()) { memo.cui.office = '[Office]'; $('#m_cuiOffice').value = memo.cui.office; }
    if (key === 'cui.on' || key === 'cui.portion') { renderParas(); renderRosters(); }   // show or hide the portion marks
    if (key.startsWith('poc.') || key.startsWith('signer.') || key.startsWith('receipt.signer.') || key.startsWith('cosign.signer.')) renderQuick();
    touched();
  });
  $('#todayBtn').addEventListener('click', () => {
    const d = new Date(), iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    memo.date = iso; $('#m_date').value = iso; touched();
  });

  /* ---------------- compose: body outline ---------------- */
  const labels = () => new Map(numbered(bodyParas(memo)).map((p) => [p.id, p]));
  function renderParas(focusId, caret) {
    const lab = labels();
    $('#paraList').innerHTML = memo.paras.map((p) => {
      const l = lab.get(p.id), level = l ? l.level : Math.min(3, p.level || 0);
      const cuiP = p.mark === 'CUI';
      const mark = portionOn(memo) ? `<button type="button" class="pmark" data-p="mark" aria-pressed="${cuiP}" aria-label="Portion mark ${cuiP ? 'CUI' : 'unclassified'}; click to change" title="Portion mark: click to switch between (U) and (CUI)">${cuiP ? '(CUI)' : '(U)'}</button>` : '';
      return `<div class="prow" data-id="${p.id}" style="--lvl:${[0, 1, 2, 2][level]}">
        <span class="plead"><span class="pnum" aria-hidden="true">${esc(l ? l.label : '·')}</span>${mark}</span>
        <textarea rows="1" autocomplete="off" aria-label="Paragraph ${esc(l?.label || 'text')}">${esc(p.text)}</textarea>
        <span class="pact">
          <button type="button" class="iconb bare" data-p="out" aria-label="Move up a level" title="Up a level (Shift+Tab)">${ico('chevron-left')}</button>
          <button type="button" class="iconb bare" data-p="in" aria-label="Move down a level" title="Down a level (Tab)">${ico('chevron-right')}</button>
          <button type="button" class="iconb bare" data-p="up" aria-label="Move paragraph up" title="Move up (Alt+Up)">${ico('arrow-up')}</button>
          <button type="button" class="iconb bare" data-p="down" aria-label="Move paragraph down" title="Move down (Alt+Down)">${ico('arrow-down')}</button>
          <button type="button" class="iconb bare danger" data-p="del" aria-label="Delete paragraph" title="Delete">${ico('trash-2')}</button>
        </span></div>`;
    }).join('');
    $$('#paraList textarea').forEach(grow);
    rosterAnchors();
    if (focusId) { const t = $(`.prow[data-id="${focusId}"] textarea`); if (t) { t.focus(); const c = caret ?? t.value.length; t.setSelectionRange(c, c); } }
  }
  const grow = (t) => { t.style.height = 'auto'; t.style.height = t.scrollHeight + 2 + 'px'; };
  function refreshLabels() {
    const lab = labels();
    $$('.prow').forEach((r) => { const l = lab.get(r.dataset.id); $('.pnum', r).textContent = l ? l.label : '·'; if (l) r.style.setProperty('--lvl', [0, 1, 2, 2][l.level]); });
    rosterAnchors();
  }
  const paraIndex = (id) => memo.paras.findIndex((p) => p.id === id);
  const rowOf = (el) => el.closest('.prow')?.dataset.id;
  function setLevel(id, d) { const p = memo.paras[paraIndex(id)]; p.level = Math.max(0, Math.min(3, (p.level || 0) + d)); touched(); renderParas(id); }
  function movePara(id, d) {
    const a = memo.paras, i = paraIndex(id), j = i + d; if (j < 0 || j >= a.length) return;
    [a[i], a[j]] = [a[j], a[i]]; touched(); renderParas(id);
  }
  function delPara(id, focusPrev) {
    const i = paraIndex(id);
    if (memo.paras.length === 1) { memo.paras[0].text = ''; touched(); renderParas(memo.paras[0].id); return; }
    const [p] = memo.paras.splice(i, 1), owner = memo;
    touched(); renderParas(focusPrev ? memo.paras[Math.max(0, i - 1)].id : undefined);
    if (String(p.text).trim()) pushUndo('Deleted a paragraph', () => { owner.paras.splice(i, 0, p); });
  }
  function addPara(afterId, level, text = '') {
    const i = afterId ? paraIndex(afterId) : memo.paras.length - 1, p = para(level ?? memo.paras[i]?.level ?? 0, text);
    memo.paras.splice(i + 1, 0, p); touched(); renderParas(p.id, 0);
  }
  $('#addPara').addEventListener('click', () => addPara(null));
  $('#paraList').addEventListener('input', (e) => {
    if (e.target.tagName !== 'TEXTAREA') return;
    memo.paras[paraIndex(rowOf(e.target))].text = e.target.value; grow(e.target); refreshLabels(); touched();
  });
  $('#paraList').addEventListener('click', (e) => {
    const go = e.target.closest('[data-goto]');
    if (go) { const el = $(`.roster[data-r="${go.dataset.goto}"] tbody input`); if (el) { el.scrollIntoView({ block: 'center', behavior: 'smooth' }); el.focus({ preventScroll: true }); } return; }
    const b = e.target.closest('[data-p]'); if (!b) return;
    const id = rowOf(b), a = b.dataset.p;
    if (a === 'mark') {
      const p = memo.paras[paraIndex(id)]; p.mark = p.mark === 'CUI' ? 'U' : 'CUI';
      b.textContent = p.mark === 'CUI' ? '(CUI)' : '(U)'; b.setAttribute('aria-pressed', p.mark === 'CUI');
      b.setAttribute('aria-label', `Portion mark ${p.mark === 'CUI' ? 'CUI' : 'unclassified'}; click to change`);
      touched(); return;
    }
    if (a === 'in') setLevel(id, 1); else if (a === 'out') setLevel(id, -1); else if (a === 'up') movePara(id, -1); else if (a === 'down') movePara(id, 1); else delPara(id);
  });
  function wrapSel(t, mark) {
    const s = t.selectionStart, e = t.selectionEnd, v = t.value, sel = v.slice(s, e);
    t.value = v.slice(0, s) + mark + sel + mark + v.slice(e);
    t.setSelectionRange(s + mark.length, e + mark.length);
    t.dispatchEvent(new Event('input', { bubbles: true }));
  }
  let lastTa = null;
  $('#paraList').addEventListener('focusin', (e) => { if (e.target.tagName === 'TEXTAREA') lastTa = e.target; });
  const emphasis = (mark) => { if (lastTa?.isConnected) { lastTa.focus(); wrapSel(lastTa, mark); } else toast('Click into a paragraph and select the words first.'); };
  $('#boldBtn').addEventListener('mousedown', (e) => e.preventDefault());
  $('#italicBtn').addEventListener('mousedown', (e) => e.preventDefault());
  $('#boldBtn').addEventListener('click', () => emphasis('**'));
  $('#italicBtn').addEventListener('click', () => emphasis('*'));
  $('#paraList').addEventListener('keydown', (e) => {
    const t = e.target; if (t.tagName !== 'TEXTAREA') return;
    const id = rowOf(t), mod = e.ctrlKey || e.metaKey;
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      const p = memo.paras[paraIndex(id)], c = t.selectionStart, rest = p.text.slice(t.selectionEnd);
      p.text = p.text.slice(0, c).replace(/\s+$/, ''); addPara(id, p.level, rest.replace(/^\s+/, ''));
    } else if (e.key === 'Tab') { e.preventDefault(); setLevel(id, e.shiftKey ? -1 : 1); }
    else if (e.key === 'Backspace' && t.selectionStart === 0 && t.selectionEnd === 0) {
      const p = memo.paras[paraIndex(id)];
      if (!t.value && memo.paras.length > 1) { e.preventDefault(); delPara(id, true); }
      else if (p.level > 0) { e.preventDefault(); setLevel(id, -1); }
    } else if (e.altKey && (e.key === 'ArrowUp' || e.key === 'ArrowDown')) { e.preventDefault(); movePara(id, e.key === 'ArrowUp' ? -1 : 1); }
    else if (mod && e.key.toLowerCase() === 'b') { e.preventDefault(); wrapSel(t, '**'); }
    else if (mod && e.key.toLowerCase() === 'i') { e.preventDefault(); wrapSel(t, '*'); }
  });
  // Pasting several lines: numbered lines keep their level; an unnumbered line after a numbered one is that paragraph wrapping.
  function parseOutline(text) {
    const lines = text.replace(/\r/g, '').split('\n'), MARK = /^\s*(?:(\d{1,3})\.|([a-z]{1,2})\.|\((\d{1,3})\)|\(([a-z]{1,2})\))\s+(.*)$/;
    const hasMarks = lines.some((l) => MARK.test(l)), out = []; let gap = false;
    for (const raw of lines) {
      const l = raw.trim();
      if (!l) { gap = true; continue; }
      const m = raw.match(MARK);
      if (m) out.push({ level: m[1] ? 0 : m[2] ? 1 : m[3] ? 2 : 3, text: m[5].trim() });
      else if (hasMarks && out.length && !gap) out[out.length - 1].text += ' ' + l;
      else out.push({ level: 0, text: l });
      gap = false;
    }
    return out;
  }
  $('#paraList').addEventListener('paste', (e) => {
    const t = e.target; if (t.tagName !== 'TEXTAREA') return;
    const text = e.clipboardData.getData('text/plain'); if (!/\n/.test(text.trim())) return;
    const items = parseOutline(text); if (items.length < 2) return;
    e.preventDefault();
    const i = paraIndex(rowOf(t)), blankRow = !memo.paras[i].text.trim();
    memo.paras.splice(blankRow ? i : i + 1, blankRow ? 1 : 0, ...items.map((x) => para(x.level, x.text)));
    touched(); renderParas(memo.paras[(blankRow ? i : i + 1) + items.length - 1].id);
    toast(`Added ${items.length} paragraphs.`);
  });

  /* ---------------- compose: rosters (memory only, like the rest of the form) ---------------- */
  const PRESETS = [['Rank', 'Name', 'DODID'], ['Rank', 'Name', 'Unit'], ['Rank', 'Name', 'Position']];
  const rosterOf = (el) => memo.rosters.find((r) => r.id === el.closest('.roster')?.dataset.r);
  const lastTextPara = () => (lastTa?.isConnected && rowOf(lastTa)) || [...memo.paras].reverse().find((p) => String(p.text).trim())?.id || memo.paras[memo.paras.length - 1]?.id || '';
  // Enclosure numbers as the layout gives them: typed enclosures first, then enclosure rosters in order (1-34).
  function enclNumber(r) {
    const list = memo.rosters.filter((x) => x.place === 'encl' && rosterTable(x));
    const k = list.indexOf(r); if (k < 0) return 0;
    return splitLines(memo.encl.items).length + k + 1;
  }
  function anchorOptions(r) {
    const lab = labels(), ids = memo.paras.map((p) => p.id);
    if (!ids.includes(r.after)) r.after = ids[ids.length - 1] || '';
    return memo.paras.map((p) => {
      const l = lab.get(p.id), t = plain(p.text).trim();
      const name = l ? `${l.label || 'Paragraph'} ${t.length > 48 ? t.slice(0, 48) + '…' : t}` : '(empty paragraph: the roster goes after the last paragraph)';
      return `<option value="${p.id}"${p.id === r.after ? ' selected' : ''}>${esc(name)}</option>`;
    }).join('');
  }
  function rosterAnchors() { $$('.roster').forEach((sec) => { const r = rosterOf(sec), s = $('[data-rk="after"]', sec); if (r && s) s.innerHTML = anchorOptions(r); }); rosterMarkers(); }
  // Shows in the Body list where each body roster prints, placed the way the layout places it: after its paragraph, or after the
  // last paragraph with text when its own is empty or gone. Anything after that paragraph prints below the roster.
  function rosterMarkers() {
    const list = $('#paraList'); if (!list) return;
    const own = numbered(bodyParas(memo)).filter((p) => memo.paras.some((q) => q.id === p.id)).map((p) => p.id);
    const at = memo.rosters.map((r, i) => ({ r, i, T: r.place !== 'encl' && rosterTable(r) })).filter((x) => x.T)
      .map(({ r, i, T }) => ({ i, id: r.id, n: T.rows.length, anchor: own.includes(r.after) ? r.after : own[own.length - 1] || '' }));
    const sig = JSON.stringify(at);
    if (list.dataset.marks === sig && $$('.rmark', list).length === at.length) return;   // unchanged: leave the DOM alone while typing
    $$('.rmark', list).forEach((m) => m.remove());
    const groups = new Map();
    at.forEach((x) => groups.set(x.anchor, (groups.get(x.anchor) || '') + `<div class="rmark" role="note">${ico('layers')}<span>Roster ${x.i + 1} prints here, ${x.n} ${x.n === 1 ? 'row' : 'rows'}. Paragraphs below it print under the roster.</span><button type="button" class="link" data-goto="${x.id}">Edit roster</button></div>`));
    groups.forEach((html, anchor) => {
      const row = anchor && $(`.prow[data-id="${anchor}"]`, list);
      if (row) row.insertAdjacentHTML('afterend', html); else list.insertAdjacentHTML('afterbegin', html);
    });
    list.dataset.marks = sig;
  }
  function rosterHtml(r, i) {
    const encl = r.place === 'encl', n = encl ? enclNumber(r) : 0;
    const tag = encl ? (n ? `Enclosure ${n}` : 'Enclosure') : 'In the body';
    const head = `${r.numbered ? '<th class="rno" scope="col">No.</th>' : ''}${r.cols.map((c, k) => `<th scope="col"><span class="rcol"><input data-col="${k}" value="${esc(c)}" autocomplete="off" aria-label="Column ${k + 1} heading" placeholder="Heading"><button type="button" class="iconb bare" data-ra="delcol" data-k="${k}" aria-label="Remove column ${esc(c || k + 1)}" title="Remove column">${ico('x')}</button></span></th>`).join('')}<th class="ract"><button type="button" class="iconb bare" data-ra="addcol" aria-label="Add a column" title="Add a column">${ico('plus')}</button></th>`;
    const rows = r.rows.map((row, ri) => `<tr data-row="${ri}">${r.numbered ? `<td class="rno">${ri + 1}</td>` : ''}${r.cols.map((c, k) => `<td><input data-cell="${k}" value="${esc(row[k] ?? '')}" autocomplete="off" aria-label="Row ${ri + 1}, ${esc(c || `column ${k + 1}`)}"></td>`).join('')}<td class="ract"><button type="button" class="iconb bare" data-ra="up" aria-label="Move row ${ri + 1} up" title="Move up">${ico('arrow-up')}</button><button type="button" class="iconb bare" data-ra="down" aria-label="Move row ${ri + 1} down" title="Move down">${ico('arrow-down')}</button><button type="button" class="iconb bare danger" data-ra="delrow" aria-label="Delete row ${ri + 1}" title="Delete row">${ico('trash-2')}</button></td></tr>`).join('');
    return `<section class="roster" data-r="${r.id}" aria-label="Roster ${i + 1}">
      <div class="rhead"><span><b>Roster ${i + 1}</b><span class="tag">${esc(tag)}</span></span><button type="button" class="iconb bare danger" data-ra="delroster" aria-label="Delete roster ${i + 1}" title="Delete roster">${ico('trash-2')}</button></div>
      <div class="fields">
        <div class="field"><label for="rp_${r.id}">Placement</label><select id="rp_${r.id}" data-rk="place"><option value="body"${encl ? '' : ' selected'}>In the body</option><option value="encl"${encl ? ' selected' : ''}>As an enclosure</option></select></div>
        ${encl ? `<div class="field"><label for="rt_${r.id}">Title</label><input id="rt_${r.id}" data-rk="title" value="${esc(r.title)}" autocomplete="off" placeholder="Training Roster, 3 November 2026"></div>`
              : `<div class="field"><label for="ra_${r.id}">After paragraph</label><select id="ra_${r.id}" data-rk="after">${anchorOptions(r)}</select></div>`}
        ${portionOn(memo) ? `<div class="field"><label for="rm_${r.id}">Portion mark</label><select id="rm_${r.id}" data-rk="mark"><option value="CUI"${r.mark === 'U' ? '' : ' selected'}>(CUI)</option><option value="U"${r.mark === 'U' ? ' selected' : ''}>(U)</option></select></div>` : ''}
        <div class="field wide"><label class="check"><input type="checkbox" data-rk="numbered"${r.numbered ? ' checked' : ''}> Number the rows (No. column)</label></div>
      </div>
      <p class="hint after">${encl ? 'Prints on its own pages after the memo, listed under Encl, with "Encl" and its number in the lower right corner of its first page (para 1-34). Mention it in the text.' : 'Prints right under the paragraph you pick. Introduce it in that paragraph, for example "The following Soldiers completed the training:".'}</p>
      <div class="presets"><span>Columns:</span>${PRESETS.map((p, k) => `<button type="button" class="b" data-ra="preset" data-k="${k}">${esc(p.join(', '))}</button>`).join('')}</div>
      <div class="rgrid-wrap"><table class="rgrid"><thead><tr>${head}</tr></thead><tbody>${rows}</tbody></table></div>
      <div class="radd"><button type="button" class="b" data-ra="addrow">${ico('plus')}Add row</button><span class="hint">Enter moves down a row. Paste rows copied from Excel into any cell.</span></div>
    </section>`;
  }
  function renderRosters(focus) {
    $('#rosterList').innerHTML = memo.rosters.map(rosterHtml).join('');
    rosterMarkers();
    if (focus) { const el = $(focus); if (el) { el.focus(); if (el.select) el.select(); } }
  }
  const cellSel = (r, row, col) => `.roster[data-r="${r.id}"] tr[data-row="${row}"] [data-cell="${col}"]`;
  $('#addRoster').addEventListener('click', () => {
    const r = { id: uid(), place: 'body', after: lastTextPara(), title: '', cols: [...PRESETS[0]], numbered: true, rows: [PRESETS[0].map(() => '')] };
    memo.rosters.push(r); touched(); renderRosters(cellSel(r, 0, 0));
  });
  $('#rosterList').addEventListener('input', (e) => {
    const el = e.target, r = rosterOf(el); if (!r) return;
    if (el.dataset.rk) {
      const k = el.dataset.rk;
      r[k] = el.type === 'checkbox' ? el.checked : el.value;
      if (k === 'place' || k === 'numbered') renderRosters(`.roster[data-r="${r.id}"] [data-rk="${k}"]`);
    } else if (el.dataset.col !== undefined) r.cols[+el.dataset.col] = el.value;
    else if (el.dataset.cell !== undefined) r.rows[+el.closest('tr').dataset.row][+el.dataset.cell] = el.value;
    rosterMarkers(); touched();
  });
  $('#rosterList').addEventListener('click', (e) => {
    const b = e.target.closest('[data-ra]'); if (!b) return;
    const r = rosterOf(b), a = b.dataset.ra, ri = +b.closest('tr')?.dataset.row, k = +b.dataset.k;
    if (!r) return;
    if (a === 'delroster') {
      const i = memo.rosters.indexOf(r), owner = memo;
      memo.rosters.splice(i, 1); touched(); renderRosters();
      if (memoHasText({ rosters: [r] })) pushUndo('Deleted a roster', () => { owner.rosters.splice(i, 0, r); });
      return;
    }
    if (a === 'addrow') { r.rows.push(r.cols.map(() => '')); touched(); renderRosters(cellSel(r, r.rows.length - 1, 0)); return; }
    if (a === 'delrow') {
      r.rows.splice(ri, 1); if (!r.rows.length) r.rows.push(r.cols.map(() => ''));
      touched(); renderRosters(cellSel(r, Math.min(ri, r.rows.length - 1), 0)); return;
    }
    if (a === 'up' || a === 'down') {
      const j = ri + (a === 'up' ? -1 : 1); if (j < 0 || j >= r.rows.length) return;
      [r.rows[ri], r.rows[j]] = [r.rows[j], r.rows[ri]]; touched();
      renderRosters(`.roster[data-r="${r.id}"] tr[data-row="${j}"] [data-ra="${a}"]`); return;
    }
    if (a === 'addcol') { r.cols.push(''); r.rows.forEach((row) => row.push('')); touched(); renderRosters(`.roster[data-r="${r.id}"] [data-col="${r.cols.length - 1}"]`); return; }
    if (a === 'delcol') {
      if (r.cols.length === 1) { toast('A roster needs at least one column.', true); return; }
      r.cols.splice(k, 1); r.rows.forEach((row) => row.splice(k, 1)); touched(); renderRosters(); return;
    }
    if (a === 'preset') {   // headings only: the rows keep their cells in the same order
      const p = PRESETS[k];
      r.cols = [...p]; r.rows = r.rows.map((row) => p.map((_, i) => row[i] ?? ''));
      touched(); renderRosters(`.roster[data-r="${r.id}"] [data-ra="preset"][data-k="${k}"]`);
    }
  });
  $('#rosterList').addEventListener('keydown', (e) => {
    const el = e.target; if (e.key !== 'Enter' || el.dataset.cell === undefined) return;
    e.preventDefault();
    const r = rosterOf(el), ri = +el.closest('tr').dataset.row, col = +el.dataset.cell;
    if (ri === r.rows.length - 1) { r.rows.push(r.cols.map(() => '')); touched(); renderRosters(cellSel(r, ri + 1, col)); }
    else $(cellSel(r, ri + 1, col))?.focus();
  });
  // Paste from Excel or Word: tab-separated cells, one row per line. A heading line that matches the columns is skipped; pasted
  // into a heading cell (or onto a roster with no headings) it becomes the headings. A leading No. column is dropped.
  $('#rosterList').addEventListener('paste', (e) => {
    const el = e.target, r = rosterOf(el); if (!r || (el.dataset.cell === undefined && el.dataset.col === undefined)) return;
    const text = e.clipboardData.getData('text/plain').replace(/\r/g, '').replace(/\n+$/, '');
    if (!/[\t\n]/.test(text)) return;
    e.preventDefault();
    let lines = text.split('\n').map((l) => l.split('\t').map((c) => c.trim()));
    const norm = (s) => String(s ?? '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    const isNo = (s) => ['no', '', 'number', 'nr'].includes(norm(s));
    if (lines.length > 1 && isNo(lines[0][0]) && lines.slice(1).every((l) => /^\d*\.?$/.test(l[0] ?? ''))) lines = lines.map((l) => l.slice(1));
    let col = el.dataset.cell !== undefined ? +el.dataset.cell : +el.dataset.col;
    const intoHead = el.dataset.col !== undefined || r.cols.every((c) => !c.trim());
    const matchesHead = lines[0].every((c, i) => norm(c) === norm(r.cols[col + i]));
    if (intoHead || matchesHead) {
      const head = lines.shift();
      if (intoHead) head.forEach((c, i) => { while (r.cols.length <= col + i) { r.cols.push(''); r.rows.forEach((row) => row.push('')); } r.cols[col + i] = c; });
    }
    let start = el.dataset.cell !== undefined ? +el.closest('tr').dataset.row : r.rows.findIndex((row) => row.every((c) => !String(c).trim()));
    if (start < 0) start = r.rows.length;
    lines.forEach((cells, j) => {
      while (r.rows.length <= start + j) r.rows.push(r.cols.map(() => ''));
      cells.forEach((c, i) => {
        while (r.cols.length <= col + i) { r.cols.push(''); r.rows.forEach((row) => row.push('')); }
        r.rows[start + j][col + i] = c;
      });
    });
    touched(); renderRosters(lines.length ? cellSel(r, start + lines.length - 1, col) : `.roster[data-r="${r.id}"] [data-col="${col}"]`);
    toast(lines.length ? `Pasted ${lines.length} ${lines.length === 1 ? 'row' : 'rows'}.` : 'Pasted the headings.');
  });

  /* ---------------- quick switching: saved POCs and signature blocks ---------------- */
  const signerLabel = (s) => [s.rank, s.name].filter(Boolean).join(' ') || 'Unnamed';
  const same = (a, b, keys) => keys.every((k) => String(a?.[k] ?? '').trim() === String(b?.[k] ?? '').trim());
  const POC_KEYS = ['name', 'phone', 'email'], SIG_KEYS = ['name', 'rank', 'branch', 'component', 'title', 'org'];
  function renderQuick() {
    const chip = (id, text, on) => `<button type="button" class="qchip" data-id="${id}" aria-pressed="${on}">${on ? ico('check') : ''}<span>${esc(text)}</span></button>`;
    $('#pocQuick').innerHTML = lib.pocs.map((p) => chip(p.id, p.name || p.email || p.phone, same(p, memo.poc, POC_KEYS))).join('');
    $('#sigQuick').innerHTML = lib.signers.map((s) => chip(s.id, signerLabel(s), same(s, memo.signer, SIG_KEYS))).join('');
    $('#rcvQuick').innerHTML = lib.signers.map((s) => chip(s.id, signerLabel(s), same(s, memo.receipt.signer, SIG_KEYS))).join('');
    $('#coQuick').innerHTML = lib.signers.map((s) => chip(s.id, signerLabel(s), same(s, memo.cosign.signer, SIG_KEYS))).join('');
    $('#pocQuick').hidden = !lib.pocs.length; $('#sigQuick').hidden = !lib.signers.length; $('#rcvQuick').hidden = !lib.signers.length; $('#coQuick').hidden = !lib.signers.length;
  }
  const fillGroup = (prefix) => { $$(`[data-m^="${prefix}"]`).forEach((el) => fillField(el, getPath(memo, el.dataset.m))); };
  $('#pocQuick').addEventListener('click', (e) => {
    const b = e.target.closest('.qchip'); if (!b) return;
    const p = lib.pocs.find((x) => x.id === b.dataset.id); if (!p) return;
    const off = b.getAttribute('aria-pressed') === 'true';           // pressing the active one clears it
    memo.poc = off ? blankPoc() : { name: p.name, phone: p.phone, email: p.email };
    fillGroup('poc.'); renderQuick(); touched();
  });
  $('#sigQuick').addEventListener('click', (e) => {
    const b = e.target.closest('.qchip'); if (!b) return;
    const s = lib.signers.find((x) => x.id === b.dataset.id); if (!s) return;
    const off = b.getAttribute('aria-pressed') === 'true';
    const { id, authority, ...rest } = s;
    memo.signer = off ? blankSigner() : { ...blankSigner(), ...rest };
    if (memo.kind !== 'mfr') memo.authority = off ? '' : authority || '';
    fillGroup('signer.'); fillField($('#m_auth'), memo.authority); applyKind(); renderQuick(); touched();
  });
  const chipFills = (host, part) => $(host).addEventListener('click', (e) => {
    const b = e.target.closest('.qchip'); if (!b) return;
    const s = lib.signers.find((x) => x.id === b.dataset.id); if (!s) return;
    const { id, authority, ...rest } = s;
    memo[part].signer = b.getAttribute('aria-pressed') === 'true' ? blankSigner() : { ...blankSigner(), ...rest };
    fillGroup(`${part}.signer.`); applyKind(); renderQuick(); touched();
  });
  chipFills('#rcvQuick', 'receipt'); chipFills('#coQuick', 'cosign');
  // Saving under a name that already exists updates that entry.
  $('#savePoc').addEventListener('click', async () => {
    const p = memo.poc; if (!POC_KEYS.some((k) => String(p[k] ?? '').trim())) { toast('Fill in the POC first.', true); return; }
    const found = lib.pocs.find((x) => x.name.trim().toLowerCase() === p.name.trim().toLowerCase());
    if (found) Object.assign(found, p); else lib.pocs.push({ id: uid(), ...p });
    const ok = await saveLib(); renderQuick();
    toast(ok ? (found ? 'POC updated.' : 'POC saved.') : 'Could not save.', !ok);
  });
  const sigKey = (x) => `${x.name} ${x.civilian ? '' : x.rank}`.trim().toLowerCase();
  const pocKey = (x) => String(x.name ?? '').trim().toLowerCase();
  async function saveSignature(s, authority) {
    if (!s.name.trim()) { toast('Fill in the name first.', true); return; }
    // Same name and same rank is the same block. A different rank stays a separate block, so saving one can never replace another person's.
    const found = lib.signers.find((x) => sigKey(x) === sigKey(s));
    if (found) Object.assign(found, s, authority === undefined ? {} : { authority }); else lib.signers.push({ id: uid(), ...s, authority: authority || '' });
    const ok = await saveLib(); renderQuick();
    toast(ok ? (found ? 'Signature block updated.' : 'Signature block saved.') : 'Could not save.', !ok);
  }
  $('#saveSigner').addEventListener('click', () => saveSignature(memo.signer, memo.authority || ''));
  $('#saveReceipt').addEventListener('click', () => saveSignature(memo.receipt.signer));
  $('#saveCosign').addEventListener('click', () => saveSignature(memo.cosign.signer));

  /* ---------------- templates ---------------- */
  // A template is the form as it stands, minus the dates, which belong to one memo.
  // A roster keeps its columns, title and placement, never its rows: rows are names and DoD ID numbers.
  const toTemplate = (m) => { const { date, suspense, ...rest } = structuredClone(m); rest.rosters = (rest.rosters || []).map((r) => ({ ...r, rows: [] })); return rest; };
  const fromTemplate = (t) => {
    const m = newMemo(structuredClone(t.memo)), ids = new Map();
    m.paras = m.paras.map((p) => { const id = uid(); ids.set(p.id, id); return { ...p, id }; });
    m.rosters = (m.rosters || []).map((r) => ({ ...r, id: uid(), after: ids.get(r.after) || '', rows: [(r.cols || []).map(() => '')] }));
    m.date = ''; m.suspense = ''; return m;
  };
  function renderTemplatePick() {
    const opts = (list) => list.map((t) => `<option value="${t.id}">${esc(t.name)}</option>`).join(''), built = builtinsShown();
    $('#tplPick').innerHTML = `<option value="">Blank memo</option>`
      + (built.length && lib.templates.length ? `<optgroup label="Your templates">${opts(lib.templates)}</optgroup>` : opts(lib.templates))
      + (built.length ? `<optgroup label="Built-in templates">${opts(built)}</optgroup>` : '');
    if (!tplById(activeTpl)) activeTpl = '';
    $('#tplPick').value = activeTpl;
    const t = lib.templates.find((x) => x.id === activeTpl), b = $('#updateTpl');
    b.hidden = !t; if (t) $('#updateTplName').textContent = t.name;
  }
  async function confirmClear(what) {
    if (!unexported()) return true;
    return ask({ title: 'Replace what is on the form?', message: `${what} clears what you typed. It has not been exported, and memo text is never saved.`, okLabel: 'Replace', danger: true });
  }
  $('#tplPick').addEventListener('change', async (e) => {
    const id = e.target.value;
    if (!(await confirmClear('Switching templates'))) { e.target.value = activeTpl; return; }
    const t = tplById(id);
    memo = t ? fromTemplate(t) : newMemo(); activeTpl = t ? id : ''; exported = true;
    renderCompose();
    const n = fillIns(memo).length;   // straight to the first thing to fill in
    if (n) { gotoFill(1, null); toast(`${n} to fill in. Alt+N goes to the next.`); }
  });
  $('#saveTpl').addEventListener('click', async () => {
    const name = (await ask({ title: 'Save as template', message: 'Saves this form, except the dates, so you can start from it again.', input: plain(memo.subject).trim(), okLabel: 'Save' }))?.trim();
    if (!name) return;
    const found = lib.templates.find((t) => t.name.toLowerCase() === name.toLowerCase());
    if (found) found.memo = toTemplate(memo); else lib.templates.push({ id: uid(), name, memo: toTemplate(memo) });
    activeTpl = (found || lib.templates[lib.templates.length - 1]).id;
    const ok = await saveLib(); renderTemplatePick();
    toast(ok ? `Template "${name}" saved.` : 'Could not save.', !ok);
  });
  $('#updateTpl').addEventListener('click', async () => {
    const t = lib.templates.find((x) => x.id === activeTpl); if (!t) return;
    t.memo = toTemplate(memo); const ok = await saveLib();
    toast(ok ? `Template "${t.name}" updated.` : 'Could not save.', !ok);
  });
  async function startMemo() {
    if (!(await confirmClear('Starting a new memo'))) return;
    const t = tplById(activeTpl);
    memo = t ? fromTemplate(t) : newMemo(); exported = true;
    renderCompose(); showView('compose'); $('#m_subject').focus();
  }
  $('#newBtn').addEventListener('click', startMemo);

  /* ---------------- preview ---------------- */
  function itemHtml(it) {
    if (it.k === 't') {
      // Fill-ins are highlighted across the whole line: the layout gives each word its own fragment. A fill-in that wraps is
      // highlighted on both lines (an open "[" at the end of one, a closing "]" at the start of the next).
      const line = it.fr.map((f) => f.t).join(''), R = [...line.matchAll(FILL_RE)].map((m) => [m.index, m.index + m[0].length]);
      const o = line.indexOf('['), c = line.indexOf(']'), lo = line.lastIndexOf('[');
      if (c >= 0 && (o < 0 || o > c)) R.push([0, c + 1]);
      if (lo >= 0 && line.indexOf(']', lo) < 0 && /[A-Za-z0-9]/.test(line.slice(lo))) R.push([lo, line.length]);
      let at = 0;
      const spans = it.fr.map((f) => {
        let html = '', i = 0;
        while (i < f.t.length) {
          const g = at + i, r = R.find(([s, e]) => g >= s && g < e);
          const next = r ? Math.min(r[1] - at, f.t.length) : Math.min(f.t.length, ...R.map(([s]) => s - at).filter((s) => s > i));
          const part = esc(f.t.slice(i, next)).replace(/☐/g, '<span class="cbox" role="img" aria-label="checkbox"></span>');
          html += r ? `<mark class="fill">${part}</mark>` : part; i = next;
        }
        at += f.t.length;
        return `<span class="${f.b ? 'b' : ''}${f.i ? ' i' : ''}">${html}</span>`;
      }).join('');
      return `<div class="ln f-${it.f}" style="left:${it.x}pt;top:${it.y}pt;font-size:${it.s}pt;line-height:${it.lh}pt">${spans}</div>`;
    }
    if (it.k === 'rule') {   // a roster border: 0.5 point, centered on the line like the PDF's
      const h = it.y1 === it.y2;
      return `<div class="rl" style="left:${(h ? it.x1 : it.x1 - 0.25).toFixed(2)}pt;top:${(h ? it.y1 - 0.25 : it.y1).toFixed(2)}pt;width:${h ? (it.x2 - it.x1).toFixed(2) : 0.5}pt;height:${h ? 0.5 : (it.y2 - it.y1).toFixed(2)}pt"></div>`;
    }
    const box = `left:${it.x}pt;top:${it.y}pt;width:${it.w}pt;height:${it.h}pt`;
    if (it.k === 'img') return `<img class="seal" alt="" src="${esc(it.src)}" style="${box}">`;
    if (it.k === 'seal') return `<div class="sbox seal-ph" style="${box}">Seal</div>`;
    if (it.k === 'sig') return `<div class="sbox" style="${box}">${esc(it.label)}</div>`;
    if (it.k === 'date') return `<div class="sbox" style="${box}">Date box</div>`;
    return '';
  }
  function sizeSheets() { const w = $('#pages').clientWidth; if (w) $('#pages').style.setProperty('--s', (w / 816).toFixed(4)); }
  new ResizeObserver(sizeSheets).observe($('#pages'));
  function renderPreview() {
    const o = officeOf(memo), L = layout(memo, o, fmt());
    $('#pages').innerHTML = L.pages.map((p) => `<div class="sheetwrap"><div class="sheet" role="img" aria-label="Page ${p.n} of ${L.pages.length}">${p.items.map(itemHtml).join('')}</div></div>`).join('');
    sizeSheets(); renderChecks(o, L); renderFills(); subjectHint(); pocHint();
    const e = L.pages.filter((p) => p.kind === 'encl' || p.kind === 'enclCont').length, n = L.pages.length - e;
    $('#subline').textContent = `AR 25-50 memorandums · ${n} ${n === 1 ? 'page' : 'pages'}${e ? ` + ${e} enclosure ${e === 1 ? 'page' : 'pages'}` : ''}`;
  }
  const LEVEL = { fix: 'circle-x', warn: 'triangle-alert', tip: 'info' };
  function renderChecks(o, L) {
    const res = check(memo, o, L), order = { fix: 0, warn: 1, tip: 2 };
    res.sort((a, b) => order[a.level] - order[b.level]);
    const n = (k) => res.filter((r) => r.level === k).length, fix = n('fix'), warn = n('warn');
    const summary = fix ? `${fix} to fix${warn ? `, ${warn} to check` : ''}` : warn ? `${warn} to check` : 'Meets the rules checked';
    if ($('#checkSum').textContent !== summary) $('#checkSum').textContent = summary;
    $('#checkSum').className = 'chk-sum ' + (fix ? 'bad' : warn ? 'warn' : 'ok');
    $('#reviewCount').textContent = fix + warn || 'Ready';
    $('#reviewCount').classList.toggle('ready', !fix && !warn);
    $('#reviewBtn').setAttribute('aria-label', 'Preview and checks. ' + $('#checkSum').textContent);
    $('#checkList').innerHTML = res.map((r) => `<li class="chk ${r.level}"><button type="button" class="chk-go" data-target="${esc(r.target)}">${ico(LEVEL[r.level])}<span>${esc(r.msg)}</span><em>${esc(/^(fig|app|MemoGen|DoDI)/.test(r.ref) ? r.ref : 'para ' + r.ref)}</em></button></li>`).join('');
  }
  const TARGETS = { cui: '#m_cuiOn', officeSymbol: '#m_symbol', record: '#m_record', date: '#m_date', suspense: '#m_susp', receipt: '#m_rcvName', cosign: '#m_coName', poc: '#m_pocname', to: '#m_to', thru: '#m_thru', distribution: '#m_dist', kind: '#m_kind', subject: '#m_subject', authority: '#m_auth', signer: '#m_sigName', encl: '#m_encl' };
  $('#checkList').addEventListener('click', (e) => {
    const b = e.target.closest('.chk-go'); if (!b) return;
    const t = b.dataset.target;
    if (t === 'setup') { showView('saved'); return; }
    if (/^[mpr]:/.test(t)) { const f = fillIns(memo).find((x) => x.key === t); if (f && selectFill(f)) return; }   // a fill-in: select it
    if (t.startsWith('roster:')) {
      const sec = $(`.roster[data-r="${t.slice(7)}"]`), el = sec && ($('tbody input', sec) || $('input', sec));
      if (el) { el.scrollIntoView({ block: 'center', behavior: 'smooth' }); el.focus({ preventScroll: true }); }
      return;
    }
    const el = t === 'body' ? ($$('#paraList textarea').pop() || $('#addPara')) : $(TARGETS[t] || `.prow[data-id="${t}"] textarea`);
    if (el) { el.scrollIntoView({ block: 'center', behavior: 'smooth' }); el.focus({ preventScroll: true }); }
  });

  /* ---------------- fill-ins: [bracketed] text a template leaves to replace (memo-format.js fillIns) ---------------- */
  const fillEl = (key) => (key.startsWith('m:') ? $(`#memoForm [data-m="${key.slice(2)}"]`) : key.startsWith('p:') ? $(`.prow[data-id="${key.slice(2)}"] textarea`) : key.startsWith('r:') ? $(`#rt_${key.slice(2)}`) : null);
  const fillKeyOf = (el) => (el?.dataset?.m ? `m:${el.dataset.m}` : el?.closest?.('.prow') ? `p:${el.closest('.prow').dataset.id}` : el?.dataset?.rk === 'title' ? `r:${el.closest('.roster').dataset.r}` : '');
  let lastField = null;   // where the cursor was, so the counter button continues from there
  $('#memoForm').addEventListener('focusin', (e) => { lastField = e.target; });
  // Scrolls to the field and selects the bracketed text, so typing replaces it.
  function selectFill(f) {
    const el = fillEl(f.key); if (!el) return false;
    el.scrollIntoView({ block: 'center', behavior: 'smooth' });
    el.focus({ preventScroll: true });
    try { el.setSelectionRange(f.at, f.at + f.text.length); } catch { /* not a text field */ }
    return true;
  }
  // The next (dir 1) or previous (-1) fill-in from the cursor, wrapping around. from: an element, or null to start at the top.
  function gotoFill(dir = 1, from) {
    const list = fillIns(memo).map((f) => ({ ...f, el: fillEl(f.key) })).filter((f) => f.el);
    if (!list.length) { toast('Nothing left to fill in.'); return; }
    const inForm = (el) => (el && el.isConnected && $('#memoForm').contains(el) ? el : null);
    const cur = from === undefined ? inForm(document.activeElement) || inForm(lastField) : inForm(from);
    const after = (f) => (f.el === cur ? f.at >= (cur.selectionEnd ?? 0) : !!(cur.compareDocumentPosition(f.el) & Node.DOCUMENT_POSITION_FOLLOWING));
    const before = (f) => (f.el === cur ? f.at + f.text.length <= (cur.selectionStart ?? 0) : !!(cur.compareDocumentPosition(f.el) & Node.DOCUMENT_POSITION_PRECEDING));
    const f = !cur ? list[dir > 0 ? 0 : list.length - 1] : dir > 0 ? list.find(after) || list[0] : [...list].reverse().find(before) || list[list.length - 1];
    selectFill(f);
  }
  // The counter in the top bar and an amber outline on each field that still has one.
  function renderFills() {
    const list = fillIns(memo), keys = new Set(list.map((f) => f.key)), n = list.length;
    $$('#memoForm .needs').forEach((el) => { if (!keys.has(fillKeyOf(el))) el.classList.remove('needs'); });
    keys.forEach((k) => fillEl(k)?.classList.add('needs'));
    const b = $('#fillBtn'); b.hidden = !n;
    $('#fillCount').textContent = `${n} to fill in`;
    b.setAttribute('aria-label', `${n} ${n === 1 ? 'fill-in' : 'fill-ins'} left. Go to the next one.`);
  }
  $('#fillBtn').addEventListener('click', () => gotoFill(1));
  document.addEventListener('keydown', (e) => {
    if (!e.altKey || e.ctrlKey || e.metaKey || e.code !== 'KeyN' || document.querySelector('dialog[open]') || !$('#view-compose').classList.contains('active')) return;
    e.preventDefault(); gotoFill(e.shiftKey ? -1 : 1);
  });

  /* ---------------- saved view ---------------- */
  let pickOffice = '', pendingSeal = null;
  const curOffice = () => lib.offices.find((o) => o.id === pickOffice) || lib.offices[0];
  const row = (id, title, detail, verb = 'Rename') => `<li data-id="${id}"><span class="s-main"><b>${esc(title)}</b>${detail ? `<small>${esc(detail)}</small>` : ''}</span><span class="row nowrap">${verb ? `<button type="button" class="iconb bare" data-a="edit" aria-label="${verb} ${esc(title)}" title="${verb}">${ico('pencil')}</button>` : ''}<button type="button" class="iconb bare danger" data-a="del" aria-label="Delete ${esc(title)}" title="Delete">${ico('trash-2')}</button></span></li>`;
  const none = (msg) => `<li class="s-none">${esc(msg)}</li>`;
  function renderSaved() {
    $('#tplList').innerHTML = lib.templates.map((t) => row(t.id, t.name, KINDS[t.memo?.kind]?.label || '')).join('') || none('None yet. Fill in a memo, then choose Save as template.');
    $('#pocList').innerHTML = lib.pocs.map((p) => row(p.id, p.name || 'Unnamed', [p.phone, p.email].filter(Boolean).join(' · '), 'Edit')).join('') || none('None yet. Choose Add, or fill in the POC on the form and choose Save POC.');
    $('#sigList').innerHTML = lib.signers.map((s) => row(s.id, signerLabel(s), s.title || '', 'Edit')).join('') || none('None yet. Choose Add, or fill in a signature block on the form and choose Save signature block.');
    const o = curOffice(); pickOffice = o.id; pendingSeal = null;
    $('#officePick').innerHTML = lib.offices.map((x) => `<option value="${x.id}">${esc(x.name || 'Letterhead')}</option>`).join(''); $('#officePick').value = o.id;
    $$('[data-o]').forEach((el) => fillField(el, o[el.dataset.o]));
    showSeal(o.seal);
    $('#delOffice').disabled = lib.offices.length < 2;
    $('#f_font').innerHTML = Object.entries(FONTS).map(([k, v]) => `<option value="${k}">${esc(v.label)}</option>`).join('');
    $('#f_kind').innerHTML = Object.entries(KINDS).map(([k, v]) => `<option value="${k}">${esc(v.label)}</option>`).join('');
    $$('[data-f]').forEach((el) => fillField(el, lib.defaults[el.dataset.f]));
  }
  function showSeal(src) { const img = $('#sealImg'); img.hidden = !src; if (src) img.src = src; else img.removeAttribute('src'); }
  function renderLibrary() { renderTemplatePick(); renderQuick(); if ($('#view-saved').classList.contains('active')) renderSaved(); $('#m_office').innerHTML = lib.offices.map((o) => `<option value="${o.id}">${esc(o.name || 'Letterhead')}</option>`).join(''); $('#m_office').value = officeOf(memo).id; renderPreview(); }
  function listDelete(list, id, label) {
    const i = list.findIndex((x) => x.id === id); if (i < 0) return;
    const [x] = list.splice(i, 1);
    saveLib(); renderLibrary();
    pushUndo(`Deleted ${label(x)}`, async () => { list.splice(i, 0, x); await saveLib(); });
  }
  $('#tplList').addEventListener('click', async (e) => {
    const b = e.target.closest('[data-a]'); if (!b) return;
    const id = b.closest('li').dataset.id, t = lib.templates.find((x) => x.id === id);
    if (b.dataset.a === 'del') { listDelete(lib.templates, id, (x) => x.name); return; }
    const name = (await ask({ title: 'Rename template', input: t.name, okLabel: 'Save' }))?.trim();
    if (name) { t.name = name; await saveLib(); renderLibrary(); }
  });
  $('#pocList').addEventListener('click', (e) => {
    const b = e.target.closest('[data-a]'); if (!b) return;
    const id = b.closest('li').dataset.id;
    if (b.dataset.a === 'del') listDelete(lib.pocs, id, (x) => x.name || 'POC'); else openEntry('poc', id);
  });
  $('#sigList').addEventListener('click', (e) => {
    const b = e.target.closest('[data-a]'); if (!b) return;
    const id = b.closest('li').dataset.id;
    if (b.dataset.a === 'del') listDelete(lib.signers, id, signerLabel); else openEntry('sig', id);
  });

  /* Add or edit a POC or signature block straight from the Saved tab. The dialog's fields are not memo fields:
     nothing is kept until Save, and Cancel drops what was typed. */
  let entry = null;
  const entryBox = () => $(entry.type === 'poc' ? '#entryPoc' : '#entrySig');
  const entryCiv = () => { $$('#entrySig .emil').forEach((el) => { el.hidden = $('#e_sigCiv').checked; }); };
  function openEntry(type, id) {
    const list = type === 'poc' ? lib.pocs : lib.signers, cur = id ? list.find((x) => x.id === id) : null;
    if (id && !cur) return;
    entry = { type, id: cur ? cur.id : '' };
    $('#entryTitle').textContent = `${cur ? 'Edit' : 'Add'} ${type === 'poc' ? 'point of contact' : 'signature block'}`;
    $('#entryPoc').hidden = type !== 'poc'; $('#entrySig').hidden = type === 'poc';
    $('#entryErr').hidden = true;
    $$('[data-e]', entryBox()).forEach((el) => fillField(el, cur?.[el.dataset.e] ?? (el.type === 'checkbox' ? false : '')));
    entryCiv();
    $('#entryDialog').showModal();
    $('input', entryBox()).focus();
  }
  $('#e_sigCiv').addEventListener('change', entryCiv);
  $('#entryForm').addEventListener('input', () => { $('#entryErr').hidden = true; });
  $('#entryForm').addEventListener('submit', async (e) => {
    e.preventDefault(); if (!entry) return;
    const v = {};
    $$('[data-e]', entryBox()).forEach((el) => { v[el.dataset.e] = el.type === 'checkbox' ? el.checked : el.value.trim(); });
    const poc = entry.type === 'poc', list = poc ? lib.pocs : lib.signers, key = poc ? pocKey : sigKey;
    const err = (msg) => { $('#entryErr').textContent = msg; $('#entryErr').hidden = false; };
    if (!v.name) { err('Fill in the name.'); $('input', entryBox()).focus(); return; }
    if (!poc && v.civilian) { v.rank = ''; v.branch = ''; v.component = ''; }
    const clash = list.find((x) => x.id !== entry.id && key(x) === key(v));
    if (clash) { err(`${poc ? clash.name : signerLabel(clash)} is already saved. Edit that one instead.`); return; }
    const cur = list.find((x) => x.id === entry.id);
    if (cur) Object.assign(cur, v); else list.push({ id: uid(), ...v });
    const ok = await saveLib();
    $('#entryDialog').close(); renderLibrary();
    toast(ok ? `${poc ? 'POC' : 'Signature block'} ${cur ? 'updated' : 'saved'}.` : 'Could not save.', !ok);
  });
  $('#entryDialog').addEventListener('close', () => { entry = null; $$('[data-e]').forEach((el) => fillField(el, el.type === 'checkbox' ? false : '')); });
  $('#addPoc').addEventListener('click', () => openEntry('poc'));
  $('#addSig').addEventListener('click', () => openEntry('sig'));
  $('#officePick').addEventListener('change', (e) => { pickOffice = e.target.value; renderSaved(); });
  $('#officeForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const o = curOffice(); $$('[data-o]').forEach((el) => { o[el.dataset.o] = fieldVal(el); });
    if (pendingSeal) o.seal = pendingSeal;
    pendingSeal = null;
    const ok = await saveLib(); renderLibrary(); toast(ok ? 'Letterhead saved.' : 'Could not save.', !ok);
  });
  $('#fmtForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    $$('[data-f]').forEach((el) => { const k = el.dataset.f; lib.defaults[k] = el.type === 'checkbox' ? el.checked : ['size', 'spaces'].includes(k) ? +el.value : el.value; });
    const ok = await saveLib(); renderPreview(); toast(ok ? 'Format saved.' : 'Could not save.', !ok);
  });
  $('#newOffice').addEventListener('click', async () => {
    const name = (await ask({ title: 'New letterhead', message: 'A short name, only shown in this app.', input: '', okLabel: 'Add' }))?.trim(); if (!name) return;
    const o = { id: uid(), name, headline: 'DEPARTMENT OF THE ARMY', lines: '', symbol: '', seal: 'seal.png' };
    lib.offices.push(o); pickOffice = o.id; await saveLib(); renderLibrary();
  });
  $('#delOffice').addEventListener('click', () => {
    if (lib.offices.length < 2) return;
    const o = curOffice(), i = lib.offices.indexOf(o), wasDefault = lib.defaults.officeId === o.id;
    lib.offices.splice(i, 1);
    if (wasDefault) lib.defaults.officeId = lib.offices[0].id;
    pickOffice = lib.offices[0].id; saveLib(); renderLibrary();
    pushUndo(`Deleted ${o.name || 'letterhead'}`, async () => { lib.offices.splice(i, 0, o); if (wasDefault) lib.defaults.officeId = o.id; await saveLib(); });
  });
  // Any uploaded seal becomes an 800 px PNG, the one format both the PDF and the Word template take as is. It is kept when you choose Save letterhead.
  $('#sealBtn').addEventListener('click', () => $('#sealInput').click());
  $('#sealInput').addEventListener('change', async (e) => {
    const f = e.target.files[0]; e.target.value = ''; if (!f) return;
    try {
      const bmp = await createImageBitmap(f), c = document.createElement('canvas'); c.width = c.height = 800;
      const k = Math.min(800 / bmp.width, 800 / bmp.height), w = bmp.width * k, h = bmp.height * k;
      c.getContext('2d').drawImage(bmp, (800 - w) / 2, (800 - h) / 2, w, h);
      pendingSeal = c.toDataURL('image/png'); showSeal(pendingSeal); toast('Choose Save letterhead to keep this seal.');
    } catch { toast('That image could not be read. Use a PNG or JPG.', true); }
  });
  $('#sealReset').addEventListener('click', () => { pendingSeal = 'seal.png'; showSeal('seal.png'); toast('Choose Save letterhead to keep this seal.'); });

  /* ---------------- AR 25-50 tab ---------------- */
  const HOW = { auto: ['Automatic', 'ok'], check: ['Checked', 'warn'], you: ['Your call', ''] };
  $('#ruleTable').innerHTML = RULES.map(([title, rows]) => `<table class="rules"><thead><tr><th>${esc(title)}</th><th>Paragraph</th><th>In MemoGen</th></tr></thead><tbody>${rows.map(([t, ref, how]) => `<tr><td>${esc(t)}</td><td class="nw">${esc(ref)}</td><td class="nw"><span class="how ${HOW[how][1]}">${HOW[how][0]}</span></td></tr>`).join('')}</tbody></table>`).join('');

  /* ---------------- export ---------------- */
  function fileBase(m) {
    const words = plain(subjectText(m)).replace(/[^\w\- ]+/g, '').trim().split(/\s+/).slice(0, 8).join('_') || 'Memo';
    return `${m.date ? m.date.replace(/-/g, '') + '_' : ''}${words}`;
  }
  async function exportMemo(kind) {
    const o = officeOf(memo);
    if (['docx', 'pdf', 'print'].includes(kind)) {   // nothing leaves with a [fill-in] still in it unless you say so
      const left = fillIns(memo);
      if (left.length) {
        const shown = left.slice(0, 5).map((f) => `${f.text} (${f.where})`).join(', ');
        const go = await ask({ title: `${left.length} ${left.length === 1 ? 'fill-in' : 'fill-ins'} left`, message: `${shown}${left.length > 5 ? `, and ${left.length - 5} more` : ''}.`, okLabel: 'Export anyway', cancelLabel: 'Go to first fill-in' });
        if (!go) { gotoFill(1, null); return; }
      }
    }
    try {
      if (kind === 'docx') {
        const r = await MemoDocx.exportDocx(memo, o, fmt());
        download(new Blob([r.bytes], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), fileBase(memo) + '.docx');
        exported = true; toast('Word document downloaded, built from your template.');
      } else if (kind === 'pdf') {
        const r = await MemoPDF.exportPDF(memo, o, fmt());
        download(new Blob([r.bytes], { type: 'application/pdf' }), fileBase(memo) + '.pdf');
        exported = true; toast(memo.digital ? 'PDF downloaded with signature and date boxes.' : 'PDF downloaded.');
      } else if (kind === 'print') { showView('compose'); window.print(); }
      else if (kind === 'json') { download(new Blob([JSON.stringify(lib, null, 2)], { type: 'application/json' }), `memogen-saved-items-${new Date().toISOString().slice(0, 10)}.json`); toast('Backup downloaded. It holds your saved items, not memo text.'); }
      else if (kind === 'import') $('#fileInput').click();
    } catch (e) { console.error(e); toast(e.message || 'Export failed.', true); }
  }
  $('#exportMenu').addEventListener('click', (e) => { const k = e.target.closest('[data-export]')?.dataset.export; if (k) exportMemo(k); });
  $('#fileInput').addEventListener('change', async (e) => {
    const f = e.target.files[0]; e.target.value = ''; if (!f) return;
    try {
      const d = JSON.parse(await f.text());
      if (!validFile(d)) throw new Error('That is not a MemoGen backup.');
      const go = await ask({ title: 'Restore saved items?', message: `"${f.name}" replaces your templates, POCs, signature blocks, letterhead and format. The current ones are kept aside first.`, okLabel: 'Replace', danger: true });
      if (!go) return;
      keepAside(lib, 'restore'); lib = repairLib(d); await saveLib(); renderLibrary(); toast('Saved items restored.');
    } catch (err) { toast(err.message || 'Could not read that file.', true); }
  });

  /* ---------------- theme ---------------- */
  function setTheme(t, keep) {
    document.documentElement.setAttribute('data-theme', t);
    $('#themeBtn').setAttribute('aria-checked', String(t === 'dark')); $('#themeState').textContent = t === 'dark' ? 'on' : 'off';
    if (keep) { try { localStorage.setItem('memogen.theme', t); } catch { /* not remembered */ } }
  }
  $('#themeBtn').addEventListener('click', () => setTheme(document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark', true));
  setTheme(document.documentElement.getAttribute('data-theme') === 'light' ? 'light' : 'dark', false);

  /* ---------------- shortcuts ---------------- */
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z' && !isTyping(document.activeElement)) { e.preventDefault(); undo(); return; }
    if (e.ctrlKey || e.metaKey || e.altKey || isTyping(document.activeElement) || document.querySelector('dialog[open]')) return;
    const views = { 1: 'compose', 2: 'saved', 3: 'rules' };
    if (views[e.key]) showView(views[e.key]);
  });

  /* Workflow navigation changes focus only; it never writes memo or library data. */
  const sectionCommands = [["section-template", "Template"], ["section-heading", "Heading"], ["cuiCard", "CUI marking"], ["section-address", "Address"], ["section-subject", "Subject"], ["section-body", "Body"], ["section-rosters", "Rosters"], ["section-point-of-contact", "Point of contact"], ["section-signature-block", "Signature block"], ["section-second-signature", "Second signature"], ["section-receipt-acknowledgment", "Receipt acknowledgment"], ["section-enclosures-and-copies", "Enclosures and copies"]];
  const jumpToSection = (id) => {
    showView('compose');
    const section = document.getElementById(id);
    if (!section || !section.getClientRects().length) return;
    const heading = section.querySelector('h3');
    if (heading) { heading.tabIndex = -1; heading.focus({ preventScroll: true }); }
    section.scrollIntoView({ block: 'start', behavior: 'instant' });
  };
  const reviewMemo = () => {
    showView('compose');
    $('.side').scrollTop = 0;
    $('#reviewHeading').focus({ preventScroll: true });
    $('.side').scrollIntoView({ block: 'start', behavior: 'instant' });
  };
  const commands = () => [
    ...sectionCommands.filter(([id]) => { const el = document.getElementById(id); return el && getComputedStyle(el).display !== 'none'; }).map(([id, name]) => ({ name, detail: 'Compose', run: () => jumpToSection(id) })),
    { name: 'Preview and checks', detail: 'Alt P', run: reviewMemo },
    { name: 'Next fill-in', detail: 'Alt N', run: () => { showView('compose'); gotoFill(1); } },
    { name: 'Saved items', detail: '2', run: () => { showView('saved'); $('#tab-saved').focus(); } },
    { name: 'AR 25-50 reference', detail: '3', run: () => { showView('rules'); $('#tab-rules').focus(); } },
    { name: 'Export options', detail: 'Word · PDF · Print', run: () => { openMenu($('#exportBtn').closest('.menu'), true); } },
    { name: 'Keyboard shortcuts', detail: '?', run: () => $('#shortcutDialog').showModal() },
  ];
  let matchingCommands = [];
  function renderJump() {
    const q = $('#jumpSearch').value.trim().toLowerCase();
    matchingCommands = commands().filter((c) => (c.name + ' ' + c.detail).toLowerCase().includes(q));
    $('#jumpResults').innerHTML = matchingCommands.length ? matchingCommands.map((c, i) => `<button type="button" class="jump-result" data-command="${i}"><span>${esc(c.name)}</span><small>${esc(c.detail)}</small></button>`).join('') : '<p class="pal-empty" role="status">No matching actions. Try Body, signature or export.</p>';
  }
  const openJump = () => { closeMenus(); $('#jumpSearch').value = ''; renderJump(); $('#jumpDialog').showModal(); $('#jumpSearch').focus(); };
  $('#jumpBtn').addEventListener('click', openJump);
  $('#jumpSearch').addEventListener('input', renderJump);
  $('#jumpDialog').addEventListener('keydown', (e) => { if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); $('#jumpDialog').close(); } });
  $('#jumpSearch').addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown' && matchingCommands.length) { e.preventDefault(); $('#jumpResults button').focus(); }
    if (e.key === 'Enter' && matchingCommands.length) { e.preventDefault(); $('#jumpDialog').close(); matchingCommands[0].run(); }
  });
  $('#jumpResults').addEventListener('click', (e) => {
    const button = e.target.closest('[data-command]'); if (!button) return;
    const command = matchingCommands[+button.dataset.command]; $('#jumpDialog').close(); command.run();
  });
  $('#jumpResults').addEventListener('keydown', (e) => {
    if (!['ArrowDown', 'ArrowUp'].includes(e.key)) return;
    const buttons = $$('button', $('#jumpResults')), at = buttons.indexOf(document.activeElement);
    if (at < 0) return; e.preventDefault();
    if (e.key === 'ArrowUp' && !at) $('#jumpSearch').focus();
    else buttons[(at + (e.key === 'ArrowDown' ? 1 : -1) + buttons.length) % buttons.length].focus();
  });
  $('#reviewBtn').addEventListener('click', reviewMemo);
  $('#helpBtn').addEventListener('click', () => $('#shortcutDialog').showModal());
  document.addEventListener('keydown', (e) => {
    if (document.querySelector('dialog[open]')) return;
    if ((e.ctrlKey || e.metaKey) && !e.altKey && e.key.toLowerCase() === 'k') { e.preventDefault(); openJump(); }
    else if (e.altKey && !e.ctrlKey && !e.metaKey && e.code === 'KeyP') { e.preventDefault(); reviewMemo(); }
    else if (e.key === '?' && !e.ctrlKey && !e.metaKey && !e.altKey && !isTyping(document.activeElement)) { e.preventDefault(); $('#shortcutDialog').showModal(); }
  });

  /* ---------------- boot ---------------- */
  renderCompose(); renderPill(); moveInd(); setStick();
  document.documentElement.classList.add('booted');
  initSync();
})();
