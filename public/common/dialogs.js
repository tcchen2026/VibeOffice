/* VibeOffice — the dialogs the three apps share, as Office has them: one Symbol, Insert Hyperlink, Spelling and
 * Properties dialog, each laid out once here. An app passes an adapter: what inserting does, where its links can
 * point, how its text is walked; the dialog itself, its labels and keys are the same in every app (docs/suite.md). */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const ui = L.ui;
  const h = (...a) => L.h(...a);
  const N = (s) => s;   // UI text shown later through L.t (common/i18n.js)
  const hex4 = (ch) => ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0');

  /* ================= Symbol ================= */
  /* Word's Special Characters; the keys are Word's own (o.keys) */
  const SPECIALS = [['—', N('Em Dash'), 'Alt+Ctrl+Num -'], ['–', N('En Dash'), 'Ctrl+Num -'], ['‑', N('Nonbreaking Hyphen'), 'Ctrl+_'], ['­', N('Optional Hyphen'), 'Ctrl+-'],
    [' ', N('Em Space'), ''], [' ', N('En Space'), ''], [' ', N('1/4 Em Space'), ''], [' ', N('Nonbreaking Space'), 'Ctrl+Shift+Space'], ['©', N('Copyright'), 'Alt+Ctrl+C'],
    ['®', N('Registered'), 'Alt+Ctrl+R'], ['™', N('Trademark'), 'Alt+Ctrl+T'], ['§', N('Section'), ''], ['¶', N('Paragraph'), ''], ['…', N('Ellipsis'), 'Alt+Ctrl+.'],
    ['‘', N('Single Opening Quote'), 'Ctrl+`,`'], ['’', N('Single Closing Quote'), "Ctrl+','"], ['“', N('Double Opening Quote'), 'Ctrl+`,"'], ['”', N('Double Closing Quote'), "Ctrl+',\""],
    ['​', N('No-Width Optional Break'), ''], ['⁠', N('No-Width Non Break'), '']];
  const SUBSETS = [[N('Basic Latin'), 0x20, 0x7e], [N('Latin-1 Supplement'), 0xa0, 0xff], [N('Latin Extended-A'), 0x100, 0x17f], [N('Greek and Coptic'), 0x370, 0x3ff],
    [N('Cyrillic'), 0x400, 0x4ff], [N('General Punctuation'), 0x2010, 0x205e], [N('Currency Symbols'), 0x20a0, 0x20bf], [N('Letterlike Symbols'), 0x2100, 0x214f],
    [N('Number Forms'), 0x2150, 0x218b], [N('Arrows'), 0x2190, 0x21ff], [N('Mathematical Operators'), 0x2200, 0x22ff], [N('Miscellaneous Technical'), 0x2300, 0x23ff],
    [N('Box Drawing'), 0x2500, 0x257f], [N('Block Elements'), 0x2580, 0x259f], [N('Geometric Shapes'), 0x25a0, 0x25ff], [N('Miscellaneous Symbols'), 0x2600, 0x26ff], [N('Dingbats'), 0x2700, 0x27bf]];
  const SYMBOL_FONTS = ['Symbol', 'Wingdings', 'Webdings'];
  const RECENT = ['€', '£', '¥', '©', '®', '™', '±', '≠', '≤', '≥', '÷', '×', '∞', 'µ', 'α', 'β'];
  /**
   * Insert ▸ Symbol. o: { fonts (the app's font names), keys (show Word's shortcut keys), insert(c) where c is
   * { ch, font, raw } (raw: the code in a symbol font, Symbol, Wingdings or Webdings; ch: its Unicode look-alike),
   * pick(c) instead of insert: a picker (a bullet, a note mark), closed by OK }. The dialog stays open after Insert,
   * as Office's does, and its Cancel becomes Close.
   */
  ui.symbolDialog = function (o) {
    o = o || {};
    let cur = null;
    const recent = L.store.get('recentSymbols', RECENT).slice(0, 16);
    const fonts = [['', L.t('(normal text)')]].concat(SYMBOL_FONTS.map((f) => [f, f]), (o.fonts || []).filter((f) => !SYMBOL_FONTS.includes(f)).map((f) => [f, f]));
    const font = ui.select(fonts, '', () => fill(), { id: 'sy-font', translate: 'no' });
    const subset = ui.select(SUBSETS.map((s, i) => [i, s[0]]), 1, () => fill(), { id: 'sy-sub' });
    const grid = h('div', { class: 'sym-grid', tabindex: '0', role: 'grid', 'aria-label': 'Symbols', translate: 'no' });
    const codeIn = h('input', { type: 'text', id: 'sy-code', style: 'width:70px', autocomplete: 'off', spellcheck: 'false' });
    const name = h('div', { class: 'tp-note sy-name' });
    const choose = (c, btn) => {
      cur = c;
      L.$$('button.on', grid).forEach((x) => x.classList.remove('on'));
      if (btn) { btn.classList.add('on'); btn.scrollIntoView({ block: 'nearest' }); }
      codeIn.value = c ? hex4(c.raw || c.ch) : '';
      name.textContent = c ? (c.raw ? L.t('{0}: {1}', c.font, hex4(c.raw)) : L.t('Unicode character {0}', hex4(c.ch))) : '';
    };
    const fill = (keep) => {
      L.clear(grid);
      const f = font.value, sym = SYMBOL_FONTS.includes(f);
      subset.disabled = sym;
      const chars = [];
      if (sym) for (let c = 0x21; c <= 0xff; c++) { if (c >= 0x7f && c < 0xa1) continue; const raw = String.fromCharCode(c), u = L.mapSymbolChar(raw, f); if (u) chars.push({ ch: u, raw, font: f }); }
      else { const [, a, b] = SUBSETS[+subset.value]; for (let c = a; c <= b; c++) chars.push({ ch: String.fromCodePoint(c), font: f }); }
      for (const c of chars) {
        const b = h('button', { type: 'button', tabindex: '-1', text: c.ch, 'aria-label': 'U+' + hex4(c.ch), style: c.font && !sym ? `font-family:${L.fontStack(c.font)}` : null });
        b._c = c;
        b.addEventListener('click', () => { choose(c, b); grid.focus(); });
        b.addEventListener('dblclick', () => { choose(c, b); insertBtn.click(); });
        grid.appendChild(b);
      }
      const again = keep && cur && [...grid.children].find((b) => b._c.ch === cur.ch);
      choose(again ? again._c : null, again);
    };
    /* the arrow keys move through the grid, 16 to a row */
    grid.addEventListener('keydown', (e) => {
      const bs = [...grid.children];
      if (!bs.length || e.altKey || e.ctrlKey || e.metaKey) return;
      let i = bs.findIndex((b) => b.classList.contains('on'));
      const step = { ArrowRight: 1, ArrowLeft: -1, ArrowDown: 16, ArrowUp: -16, Home: -Infinity, End: Infinity }[e.key];
      if (step == null) return;
      e.preventDefault(); e.stopPropagation();
      i = i < 0 ? 0 : Math.max(0, Math.min(bs.length - 1, step === Infinity ? bs.length - 1 : step === -Infinity ? 0 : i + step));
      choose(bs[i]._c, bs[i]);
    });
    fill();
    /* a code typed: that character, in the subset it belongs to */
    codeIn.addEventListener('change', () => {
      const n = parseInt(codeIn.value, 16);
      if (isNaN(n) || n < 0x20 || n > 0x10ffff) return;
      if (SYMBOL_FONTS.includes(font.value)) { font.value = ''; }
      const at = SUBSETS.findIndex(([, a, b]) => n >= a && n <= b);
      if (at >= 0 && String(at) !== subset.value) subset.value = String(at);
      fill();
      const b = [...grid.children].find((x) => x._c.ch.codePointAt(0) === n);
      choose(b ? b._c : { ch: String.fromCodePoint(n), font: font.value }, b);
    });
    const rec = h('div', { class: 'sym-recent', translate: 'no' });
    for (const ch of recent) {
      const b = h('button', { type: 'button', tabindex: '-1', text: ch, 'aria-label': 'U+' + hex4(ch) });
      b.addEventListener('click', () => choose({ ch, font: '' }));
      b.addEventListener('dblclick', () => { choose({ ch, font: '' }); insertBtn.click(); });
      rec.appendChild(b);
    }
    const spec = ui.listBox(SPECIALS.map(([c, n, k]) => ({ value: c, label: n, note: o.keys && k ? ui.keyText(k) : '' })), SPECIALS[0][0], { height: 226, onDbl: () => insertBtn.click() });
    let tab = 0;
    const tabs = ui.tabs([
      { label: '&Symbols', onShow: () => { tab = 0; }, body: h('div', { class: 'col' },
        h('div', { class: 'row', style: 'flex-wrap:nowrap' }, ui.field('&Font:', font), ui.field('Su&bset:', subset)),
        grid, h('div', { class: 'cd-cap', text: 'Recently used symbols:' }), rec,
        h('div', { class: 'row', style: 'align-items:center;flex-wrap:nowrap' }, name, ui.field('&Character code:', codeIn))) },
      { label: 'S&pecial Characters', onShow: () => { tab = 1; }, body: spec },
    ]);
    let inserted = false;
    const dlg = ui.dialog({
      title: 'Symbol', width: 520, body: tabs, focus: grid,
      buttons: [{ label: o.pick ? 'OK' : '&Insert', primary: true, onClick: () => {
        const c = tab === 1 ? { ch: spec.get(), font: '' } : cur;
        if (!c || !c.ch) return false;
        L.store.set('recentSymbols', [c.ch].concat(recent.filter((x) => x !== c.ch)).slice(0, 16));
        if (o.pick) { o.pick(c); return; }
        o.insert(c);
        /* as Office: the dialog stays for the next symbol, and Cancel says Close */
        if (!inserted) { inserted = true; dlg.buttons[1].innerHTML = ui.ampHTML('Close'); }
        return false;
      } }, { label: 'Cancel', cancel: true }],
    });
    const insertBtn = dlg.buttons[0];
    return dlg;
  };

  /* ================= Insert Hyperlink ================= */
  /** an address as typed, completed as Office does: www.… and example.com get http://; a document's name stays a
   *  file name, relative to the file */
  const DOC_EXT = /\.(docx?|docm|dotx?|xlsx?|xlsm|xltx?|pptx?|pptm|ppsx?|potx?|pdf|txt|csv|rtf|html?|xml|od[tsp]|png|jpe?g|gif|bmp|svg|zip)$/i;
  ui.completeAddress = (u) => {
    u = String(u || '').trim();
    if (!u || /^[a-z][a-z0-9+.-]*:/i.test(u) || /^[\\/.#]/.test(u) || DOC_EXT.test(u)) return u;
    return /^www\./i.test(u) || /^[^\s/]+\.[a-z]{2,}(\/|$)/i.test(u) ? 'http://' + u : u;
  };
  const KINDS = [['web', N('E&xisting File or Web Page'), 'webPage'], ['place', N('Place in This &Document'), 'bookmark'], ['mail', N('E-&mail Address'), 'mail']];
  /**
   * Insert ▸ Hyperlink (Ctrl+K), Office's dialog: Text to display and ScreenTip... at the top, the Link to bar on the
   * left, the kind's fields on the right, Remove Link when a link is being edited. o: {
   *   link: the link being edited, { url (a web address, a file, mailto:…?subject=…), place, tip }, or null;
   *   text: the text to display; textLocked: it can't change here (a shape's link: shown as <<Selection in Document>>);
   *   places: [{ value, label } | { group }] for Place in This Document, defaultPlace: the one chosen for a new link;
   *   ref: { label, value } a box for a reference
   *   typed above the list (a cell reference); tipKinds: the kinds that keep a ScreenTip (default all);
   *   apply(r) on OK, r = { kind, url, place, ref, tip, text } (url is the completed address or the mailto: link),
   *   remove() for Remove Link }
   */
  ui.hyperlinkDialog = function (o) {
    const cur = o.link || null;
    const mailto = cur && /^mailto:/i.test(cur.url || '') ? cur.url : null;
    let kind = cur ? (cur.place != null && !cur.url ? 'place' : mailto ? 'mail' : 'web') : 'web';
    let tip = (cur && cur.tip) || '';
    const disp = h('input', { type: 'text', id: 'hl-disp', autocomplete: 'off', value: o.textLocked ? L.t('<<Selection in Document>>') : o.text || '' });
    if (o.textLocked) disp.disabled = true;
    const tipBtn = ui.button('ScreenTi&p...', () => ui.prompt('ScreenTip text:', tip, 'Set Hyperlink ScreenTip').then((v) => { if (v != null) tip = v; }), { class: 'btn small' });
    /* the kinds' fields */
    const addr = h('input', { type: 'text', id: 'hl-addr', autocomplete: 'off', spellcheck: 'false', translate: 'no', value: cur && cur.url && !mailto ? cur.url : '' });
    const refIn = o.ref ? h('input', { type: 'text', id: 'hl-ref', autocomplete: 'off', spellcheck: 'false', translate: 'no', value: o.ref.value || '' }) : null;
    const first = (o.places || []).find((x) => !x.group && x.value != null);
    const places = ui.listBox(o.places || [], cur && cur.place != null ? cur.place : o.defaultPlace != null ? o.defaultPlace : first ? first.value : null, { id: 'hl-place', height: o.ref ? 150 : 190, translate: 'no' });
    const [mailAddr, query] = mailto ? mailto.replace(/^mailto:/i, '').split('?') : ['', ''];
    const subjMatch = /(?:^|&)subject=([^&]*)/i.exec(query || '');
    const mail = h('input', { type: 'text', id: 'hl-mail', autocomplete: 'off', spellcheck: 'false', translate: 'no', value: decodeURIComponent(mailAddr || '') });
    const subj = h('input', { type: 'text', id: 'hl-subj', autocomplete: 'off', value: subjMatch ? decodeURIComponent(subjMatch[1].replace(/\+/g, ' ')) : '' });
    const field = (label, ctl) => h('div', { class: 'col hl-f' }, h('label', { for: ctl.id || (ctl.input && ctl.input.id), html: ui.ampHTML(label) }), ctl);
    const panes = {
      web: h('div', { class: 'col' }, field('Addr&ess:', addr), h('div', { class: 'tp-note', text: 'Type a web address (www.example.com) or the name of a file.' })),
      place: h('div', { class: 'col' }, refIn ? field(o.ref.label, refIn) : null, field(o.ref ? 'Or select a place in this d&ocument:' : 'Select a place in this d&ocument:', places)),
      mail: h('div', { class: 'col' }, field('E-mail addre&ss:', mail), field('Su&bject:', subj)),
    };
    const host = h('div', { class: 'hl-pane' });
    const bar = h('div', { class: 'hl-bar', role: 'tablist', 'aria-label': 'Link to' });
    const kindBtns = KINDS.map(([k, label, icon]) => {
      const b = h('button', { type: 'button', class: 'hl-kind', role: 'tab', html: L.icons.get(icon, 24) + `<span>${ui.ampHTML(label)}</span>` });
      b.addEventListener('click', () => show(k, true));
      bar.appendChild(b);
      return [k, b];
    });
    function show(k, focus) {
      kind = k;
      for (const [x, b] of kindBtns) { b.classList.toggle('on', x === k); b.setAttribute('aria-selected', String(x === k)); }
      L.clear(host); host.appendChild(panes[k]);
      if (focus) { const f = k === 'web' ? addr : k === 'mail' ? mail : refIn || places.list; f.focus(); }
    }
    show(kind);
    const body = h('div', { class: 'col hl-dlg' },
      h('div', { class: 'row', style: 'align-items:center;flex-wrap:nowrap' }, h('label', { for: 'hl-disp', html: ui.ampHTML('&Text to display:') }), disp, tipBtn),
      h('div', { class: 'row', style: 'flex-wrap:nowrap;align-items:stretch' }, h('div', { class: 'col', style: 'gap:2px' }, h('div', { class: 'hl-cap', text: 'Link to:' }), bar), host));
    const buttons = [];
    if (cur && o.remove) buttons.push({ label: '&Remove Link', cls: 'left', onClick: () => { o.remove(); } });
    buttons.push({ label: 'OK', primary: true, onClick: () => {
      const r = { kind, text: o.textLocked ? null : disp.value, tip: (o.tipKinds || ['web', 'place', 'mail']).includes(kind) && tip ? tip : undefined };
      if (kind === 'web') { r.url = ui.completeAddress(addr.value); if (!r.url) { addr.focus(); return false; } }
      else if (kind === 'mail') { const m = mail.value.trim().replace(/^mailto:/i, ''); if (!m) { mail.focus(); return false; } r.url = 'mailto:' + m + (subj.value ? '?subject=' + encodeURIComponent(subj.value) : ''); }
      else { r.place = places.get(); r.ref = refIn ? refIn.value.trim() : undefined; if (r.place == null && !r.ref) { places.list.focus(); return false; } }
      return o.apply(r);
    } }, { label: 'Cancel' });
    return ui.dialog({ title: cur ? 'Edit Hyperlink' : 'Insert Hyperlink', width: 560, body, buttons, focus: o.textLocked ? (kind === 'web' ? addr : undefined) : disp });
  };

  /* ================= Spelling ================= */
  /**
   * Tools ▸ Spelling (F7): one dialog for the three apps, as Office's: what was found in its sentence, the
   * suggestions under a box for your own correction, the buttons on the right. The app walks its text; s is its
   * session: {
   *   title(lang) → the dialog's title; language() → the dictionary's name; languages: [[id, name]] and
   *   setLanguage(id) for a choice of dictionary; grammar: { on, set(v) } (Check grammar);
   *   next() → the next finding, selected in the document, or null at the end: { kind: 'spelling' | 'repeat' |
   *     'grammar', caption (a grammar rule's), before, word, after (its sentence), suggestions: ui.listBox items };
   *   ignoreOnce(), ignoreAll(), add(), change(text), changeAll(text), autocorrect(text) (optional; these three
   *   return false to stay on the finding), nextSentence(), setLanguage(id) re-checks from the current finding,
   *   explain() → a grammar rule's explanation, undo() (optional: false when there is nothing to undo), options()
   *   (optional), end: the message when the check is complete }
   * Each button acts on the current finding and moves to the next one.
   */
  ui.spellingDialog = function (s) {
    let cur = null;
    const caption = h('div', { class: 'cd-cap' });
    const ctxBox = h('div', { class: 'sg-context', translate: 'no' });
    const sugg = ui.listBox([], null, { input: true, id: 'sg-sugg', height: 110, translate: 'no', onDbl: () => btnChange.click() });
    const btn = (label, fn) => { const b = ui.button(label, () => { if (cur) fn(); }); b.style.width = '100%'; return b; };
    const relabel = (b, label) => { b.innerHTML = ui.ampHTML(label); };
    const text = () => String(sugg.get() ?? '');
    const go = () => {
      cur = s.next();
      if (!cur) { dlg.close(0); ui.msg(s.end || 'The spelling check is complete.', { icon: 'info', title: L.t(s.title(s.language())) }); return; }
      const sp = cur.kind === 'spelling';
      caption.textContent = L.t(sp ? 'Not in Dictionary:' : cur.kind === 'repeat' ? 'Repeated Word:' : cur.caption || 'Grammar:');
      ctxBox.replaceChildren(document.createTextNode(cur.before || ''), h('span', { class: sp || cur.kind === 'repeat' ? 'sg-bad' : 'sg-gram', translate: 'no', text: cur.word }), document.createTextNode(cur.after || ''));
      const items = cur.suggestions && cur.suggestions.length ? cur.suggestions : [{ label: sp ? 'No Spelling Suggestions' : 'No Suggestions', dim: true }];
      sugg.setItems(items);
      const first = items.find((x) => !(x && x.dim));
      sugg.set(first == null ? '' : typeof first === 'object' ? first.value : first);
      relabel(btnIgnoreAll, sp || cur.kind === 'repeat' ? 'I&gnore All' : 'Ignore R&ule');
      relabel(btnAdd, cur.kind === 'grammar' ? '&Next Sentence' : '&Add to Dictionary');
      btnAdd.disabled = cur.kind === 'repeat';
      btnChangeAll.disabled = !sp;
      btnAuto.disabled = !sp || !s.autocorrect;
      btnExplain.hidden = cur.kind !== 'grammar';
      btnUndo.disabled = false;
    };
    const btnIgnore = btn('&Ignore Once', () => { s.ignoreOnce(); go(); });
    const btnIgnoreAll = btn('I&gnore All', () => { s.ignoreAll(); go(); });
    const btnAdd = btn('&Add to Dictionary', () => { if (cur.kind === 'grammar') s.nextSentence(); else s.add(); go(); });
    /* a change the session refuses (nothing to change to) leaves the finding where it is */
    const btnChange = btn('&Change', () => { if (s.change(text()) !== false) go(); });
    const btnChangeAll = btn('Change A&ll', () => { if (s.changeAll(text()) !== false) go(); });
    const btnAuto = btn('AutoCo&rrect', () => { if (text() && s.autocorrect(text()) !== false) go(); });
    const btnExplain = btn('E&xplain...', () => ui.msg(s.explain(), { icon: 'info', title: 'Grammar' }));
    const btnUndo = ui.button('&Undo', () => { if (s.undo && s.undo() !== false) go(); });
    btnUndo.style.width = '100%';
    if (!s.undo) btnUndo.hidden = true;
    const lang = s.languages && s.setLanguage
      ? ui.select(s.languages, s.languages.find((x) => x[1] === s.language()) ? s.languages.find((x) => x[1] === s.language())[0] : s.languages[0][0], async (v) => { if (await s.setLanguage(v) === false) return; dlg.el.querySelector('.dlg-ttl').textContent = L.t(s.title(s.language())); go(); }, { id: 'sg-lang' })
      : h('span', { class: 'sg-lang', text: s.language() });
    const grammar = s.grammar ? ui.check('Check &grammar', s.grammar.on, (v) => s.grammar.set(v)) : null;
    const left = h('div', { class: 'col', style: 'flex:1;min-width:0' }, caption, ctxBox,
      h('label', { for: 'sg-sugg', html: ui.ampHTML('Suggestio&ns:') }), sugg,
      h('div', { class: 'row', style: 'align-items:center;flex-wrap:nowrap' }, h('label', { for: lang.id || null, html: ui.ampHTML(lang.id ? 'Dictionary &language:' : 'Dictionary language:') }), lang), grammar);
    const side = h('div', { class: 'col sg-side' }, btnIgnore, btnIgnoreAll, btnAdd, h('div', { class: 'sg-gap' }), btnChange, btnChangeAll, btnAuto, btnExplain);
    const buttons = [];
    if (s.options) buttons.push({ label: '&Options...', cls: 'left', onClick: () => { s.options(); return false; } });
    buttons.push({ label: 'Cancel', cancel: true });
    const dlg = ui.dialog({ title: s.title(s.language()), width: 560, body: h('div', { class: 'row', style: 'flex-wrap:nowrap' }, left, side), buttons, onClose: s.onClose });
    if (s.undo) dlg.el.querySelector('.dlg-foot').insertBefore(btnUndo, dlg.buttons[dlg.buttons.length - 1]);
    btnUndo.style.width = '';
    go();
    return dlg;
  };
})(typeof window !== 'undefined' ? window : globalThis);
