/* Quire — Find, Replace and Go To: Word's special characters (^p, ^t, ^?, ^# …), wildcards
 * (?, *, [ ], {n,m}, @, < >, groups and \1 back-references), match case / whole words / sounds like /
 * all word forms, formatting criteria, Highlight All, and the Go To tab.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout;
  const { h } = L;
  const ui = L.ui;
  const G = (L.dlg = L.dlg || {});
  const FD = (L.find = {});
  const doc = () => D.doc;
  const A = () => L.app;

  /* ---------- story text with paragraph marks ---------- */
  function storyText(story) {
    D.reindex(doc());
    const paras = story.paras || [];
    const starts = [];
    let s = '';
    for (const p of paras) { starts.push(s.length); s += D.ptext(p) + '\r'; }
    return { text: s, paras, starts };
  }
  function toPos(st, off) {
    let lo = 0, hi = st.starts.length - 1;
    while (lo < hi) { const mid = (lo + hi + 1) >> 1; if (st.starts[mid] <= off) lo = mid; else hi = mid - 1; }
    const p = st.paras[lo];
    return D.pos(p, Math.min(off - st.starts[lo], D.plen(p)));
  }
  function fromPos(st, pos) { const i = st.paras.indexOf(pos.p); return i < 0 ? 0 : st.starts[i] + pos.o; }

  /* ---------- pattern compilation ---------- */
  const esc = (s) => s.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&');
  const SPECIAL = { p: '\\r', t: '\\t', l: '\\n', m: '\\f', n: '\\v', '?': '[^\\r]', '#': '\\d', $: '\\p{L}', w: '[ \\t\\u00a0]+', s: '\\u00a0', '~': '\\u2011', '-': '\\u00ad', '+': '\\u2014', '=': '\\u2013', '^': '\\^', g: '\\uFFFC', f: '\\uFFFC', e: '\\uFFFC', a: '(?:)', d: '(?:)', b: '\\r', '%': '§', v: '¶', '&': '' };
  function specials(src, wildcards) {
    let out = '';
    for (let i = 0; i < src.length; i++) {
      const c = src[i];
      if (c === '^' && i + 1 < src.length) {
        const n = src[i + 1];
        if (n === 'u' && /^\d+/.test(src.slice(i + 2))) { const m = /^\d+/.exec(src.slice(i + 2))[0]; out += esc(String.fromCharCode(+m)); i += 1 + m.length; continue; }
        if (n === '0' && /^0\d+/.test(src.slice(i + 1))) { const m = /^\d+/.exec(src.slice(i + 1))[0]; out += esc(String.fromCharCode(+m)); i += m.length; continue; }
        const lower = n.toLowerCase();
        if (SPECIAL[lower] != null || SPECIAL[n] != null) { out += SPECIAL[n] != null ? SPECIAL[n] : SPECIAL[lower]; i++; continue; }
      }
      out += wildcards ? c : esc(c);
    }
    return out;
  }
  /** translate Word wildcard syntax to a JS regexp source */
  function wildcard(src) {
    let out = '';
    for (let i = 0; i < src.length; i++) {
      const c = src[i];
      if (c === '\\' && i + 1 < src.length) { out += esc(src[i + 1]); i++; continue; }
      if (c === '^') { const m = specials(src.slice(i, i + 2), false); out += m; i++; continue; }
      if (c === '?') { out += '[^\\r]'; continue; }
      if (c === '*') { out += '[^\\r]*?'; continue; }
      if (c === '@') { out += '+'; continue; }
      if (c === '<') { out += '(?<![\\p{L}\\p{N}_])(?=[\\p{L}\\p{N}_])'; continue; }
      if (c === '>') { out += '(?<=[\\p{L}\\p{N}_])(?![\\p{L}\\p{N}_])'; continue; }
      if (c === '[') {
        const j = src.indexOf(']', i + 1);
        if (j < 0) throw new Error('The Find What text contains a Pattern Match expression which is not valid.');
        let body = src.slice(i + 1, j);
        const neg = body[0] === '!';
        if (neg) body = body.slice(1);
        out += '[' + (neg ? '^' : '') + body.replace(/\\/g, '\\\\').replace(/\]/g, '\\]') + ']';
        i = j;
        continue;
      }
      if (c === '{') { const j = src.indexOf('}', i); if (j < 0) throw new Error('Invalid {} expression.'); out += src.slice(i, j + 1).replace(/;/g, ','); i = j; continue; }
      if (c === '(' || c === ')' || c === '|') { out += c; continue; }
      out += esc(c);
    }
    return out;
  }
  const FORMS = ['s', 'es', 'ed', 'd', 'ing', 'er', 'ers', 'ly', "'s"];
  function soundex(w) { const a = w.toUpperCase().replace(/[^A-Z]/g, ''); if (!a) return ''; const map = { B: 1, F: 1, P: 1, V: 1, C: 2, G: 2, J: 2, K: 2, Q: 2, S: 2, X: 2, Z: 2, D: 3, T: 3, L: 4, M: 5, N: 5, R: 6 }; let out = a[0], last = map[a[0]] || 0; for (const ch of a.slice(1)) { const c = map[ch] || 0; if (c && c !== last) out += c; if (ch !== 'H' && ch !== 'W') last = c; } return (out + '000').slice(0, 4); }
  /** compile options into a matcher: returns RegExp (global, unicode) and a post-filter */
  FD.compile = function (o) {
    let src;
    if (!o.text && !o.format) throw new Error('empty');
    if (!o.text) src = '[^\\r]+';
    else if (o.wildcards) src = wildcard(o.text);
    else if (o.allForms) { const w = o.text.trim(); const stem = w.replace(/(ing|ed|es|s)$/i, ''); src = '(?<![\\p{L}\\p{N}_])' + esc(stem.length > 2 ? stem : w) + '(?:' + FORMS.map(esc).join('|') + ')?(?![\\p{L}\\p{N}_])'; }
    else if (o.soundsLike) src = '[\\p{L}]+';
    else src = specials(o.text, false);
    if (o.wholeWord && !o.wildcards && !o.allForms && !o.soundsLike && /^[\p{L}\p{N}_]/u.test(o.text) && /[\p{L}\p{N}_]$/u.test(o.text)) src = '(?<![\\p{L}\\p{N}_])' + src + '(?![\\p{L}\\p{N}_])';
    const flags = 'gu' + (o.matchCase || o.wildcards ? '' : 'i');
    let re;
    try { re = new RegExp(src, flags); } catch (e) { throw new Error('The Find What text contains a Pattern Match expression which is not valid.'); }
    const sx = o.soundsLike ? soundex(o.text) : null;
    return { re, filter: (m) => (!sx || soundex(m[0]) === sx) };
  };
  /** does the matched range satisfy formatting criteria */
  function formatOK(o, a, b) {
    if (!o.format) return true;
    const f = o.format;
    const d = doc();
    if (f.style) { const st = d.styles[f.style]; if (st && st.type === 'paragraph' && (a.p.pPr.style || 'Normal') !== f.style) return false; }
    if (f.highlight != null) { const it = D.itemAfter(a.p, a.o); if (!!(it && it.rPr && it.rPr.hl) !== f.highlight) return false; }
    if (f.rPr) {
      for (const p of D.parasBetween(d, a.p, b.p)) {
        const s = p === a.p ? a.o : 0, e = p === b.p ? b.o : D.plen(p);
        let o2 = 0;
        for (const it of p.runs) { const n = D.ilen(it); if (n && o2 < e && o2 + n > s) { const r = D.rProps(d, p, it, LY.tblCtx(p)); for (const k in f.rPr) if (JSON.stringify(r[k] || undefined) !== JSON.stringify(f.rPr[k] || undefined) && !(f.rPr[k] === false && !r[k])) return false; } o2 += n; }
      }
    }
    return true;
  }
  /** all matches in a story: [{a, b, m}] */
  FD.all = function (o, story) {
    const st = storyText(story || doc().main);
    const { re, filter } = FD.compile(o);
    const out = [];
    re.lastIndex = 0;
    let m;
    while ((m = re.exec(st.text))) {
      if (m[0].length === 0) { re.lastIndex++; continue; }
      if (!filter(m)) continue;
      const a = toPos(st, m.index), b = toPos(st, m.index + m[0].length);
      if (!formatOK(o, a, b)) continue;
      out.push({ a, b, m, s: m.index, e: m.index + m[0].length });
      if (out.length > 20000) break;
    }
    return { matches: out, st };
  };
  const storiesFor = (o) => (o.scope === 'all' ? D.stories(doc()).filter((s) => s.kind !== 'tb') : [o.scope === 'current' && E.sel ? E.story() : doc().main]);
  /** next match after (or before) the selection, wrapping around */
  FD.next = function (o) {
    const d = doc();
    const stories = storiesFor(o);
    const curStory = E.sel ? E.story() : d.main;
    const up = o.dir === 'up';
    const order = stories.includes(curStory) ? stories.slice(stories.indexOf(curStory)).concat(stories.slice(0, stories.indexOf(curStory))) : stories;
    for (let pass = 0; pass < order.length + 1; pass++) {
      const story = order[pass % order.length];
      const { matches, st } = FD.all(o, story);
      if (!matches.length) continue;
      if (pass === 0 && E.sel && story === curStory) {
        const [a, b] = E.range();
        const from = fromPos(st, up ? a : b), cur = fromPos(st, a);
        const hit = up ? matches.slice().reverse().find((x) => x.e <= from && x.s < cur) : matches.find((x) => x.s >= from && !(x.s === cur && x.e === from && !E.collapsed() && false));
        if (hit) return hit;
        if (o.dir !== 'all' && pass === 0 && order.length === 1) { FD.wrapped = true; return up ? matches[matches.length - 1] : matches[0]; }
        continue;
      }
      return up ? matches[matches.length - 1] : matches[0];
    }
    return null;
  };
  /** replacement text for a match (wildcard groups, ^&, ^c, specials) */
  function replacementText(o, hit) {
    let r = o.replace || '';
    let out = '';
    for (let i = 0; i < r.length; i++) {
      const c = r[i];
      if (c === '\\' && o.wildcards && /\d/.test(r[i + 1] || '')) { out += hit.m[+r[i + 1]] || ''; i++; continue; }
      if (c === '^' && i + 1 < r.length) {
        const n = r[i + 1];
        i++;
        if (n === '&') { out += hit.m[0]; continue; }
        if (n === 'c' || n === 'C') { out += E.clip ? E.clip.text : ''; continue; }
        const map = { p: '\n', t: '\t', l: '\u000b', m: '\f', n: '\u000c', s: ' ', '~': '‑', '-': '­', '+': '—', '=': '–', '^': '^' };
        if (map[n] != null) { out += map[n]; continue; }
        out += '^' + n;
        continue;
      }
      out += c;
    }
    return out;
  }
  /** replace one match (inside a transaction); returns the position after the inserted text */
  function replaceHit(o, hit) {
    const text = replacementText(o, hit);
    const keepR = (() => { const it = D.itemAfter(hit.a.p, hit.a.o); return it && it.rPr ? L.clone(it.rPr) : O.inheritRPr(hit.a); })();
    const pos = O.deleteRange(hit.a, hit.b);
    if (!text && !o.replaceFormat) return pos;
    let end = pos;
    if (text) {
      /* manual breaks: \u000b line break, \f page break; \n paragraph */
      const parts = text.split(/(\u000b|\f|\u000c)/);
      for (const part of parts) {
        if (part === '\u000b') end = O.insertItem(end, D.item('br', { type: 'line' }, L.clone(keepR)));
        else if (part === '\f') end = O.insertItem(end, D.item('br', { type: 'page' }, L.clone(keepR)));
        else if (part === '\u000c') end = O.insertItem(end, D.item('br', { type: 'column' }, L.clone(keepR)));
        else if (part) end = O.insertPlain(end, part, L.clone(keepR));
      }
    }
    if (o.replaceFormat) {
      const f = o.replaceFormat;
      if (f.rPr && D.cmp(doc(), pos, end) < 0) O.setRunProps(pos, end, f.rPr);
      if (f.style && doc().styles[f.style]) O.applyStyle(pos, D.cmp(doc(), pos, end) < 0 ? end : pos, f.style);
      if (f.highlight != null && D.cmp(doc(), pos, end) < 0) O.setRunProps(pos, end, { hl: f.highlight ? A().lastHighlight || 'yellow' : undefined });
      if (!text) return pos;
    }
    return end;
  }
  FD.replaceAll = function (o) {
    let n = 0;
    E.edit('Replace All', () => {
      for (const story of storiesFor(o)) {
        const { matches } = FD.all(o, story);
        for (let i = matches.length - 1; i >= 0; i--) { replaceHit(o, matches[i]); n++; }
        O.repairFields(story);
      }
      return E.sel && D.info(doc(), E.sel.f.p) ? E.sel : D.pos(D.firstPara(doc().main.blocks), 0);
    });
    return n;
  };
  FD.replaceOne = function (o) {
    /* if the selection is exactly a match, replace it; then find the next */
    if (E.sel && !E.collapsed()) {
      const story = E.story();
      const { matches, st } = FD.all(o, story);
      const [a, b] = E.range();
      const s = fromPos(st, a), e = fromPos(st, b);
      const hit = matches.find((x) => x.s === s && x.e === e);
      if (hit) E.edit('Replace', () => replaceHit(o, hit));
    }
    return FD.next(Object.assign({}, o, { dir: o.dir === 'up' ? 'up' : 'down' }));
  };
  /* ---------- Highlight All (CSS Custom Highlight API; no document change) ---------- */
  FD.highlightAll = function (o) {
    if (!(window.CSS && CSS.highlights && window.Highlight)) return -1;
    const ranges = [];
    let n = 0;
    for (const story of storiesFor(o)) {
      const { matches } = FD.all(o, story);
      n += matches.length;
      for (const m of matches.slice(0, 3000)) {
        const a = LY.posToDom(m.a), b = LY.posToDom(m.b);
        if (!a || !b) continue;
        try { const r = new Range(); r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset); ranges.push(r); } catch (e) { /* skip */ }
      }
    }
    CSS.highlights.set('quire-find', new Highlight(...ranges));
    FD.hlOpts = o;
    return n;
  };
  FD.clearHighlight = () => { if (window.CSS && CSS.highlights) CSS.highlights.delete('quire-find'); FD.hlOpts = null; };
  L.bus.on('layout-done', () => { if (FD.hlOpts) setTimeout(() => { try { FD.highlightAll(FD.hlOpts); } catch (e) { FD.clearHighlight(); } }, 0); });

  /* ---------- Go To ---------- */
  const GOTO = [['page', 'Page'], ['section', 'Section'], ['line', 'Line'], ['bookmark', 'Bookmark'], ['comment', 'Comment'], ['footnote', 'Footnote'], ['endnote', 'Endnote'], ['field', 'Field'], ['table', 'Table'], ['graphic', 'Graphic'], ['equation', 'Equation'], ['heading', 'Heading']];
  FD.goto = function (kind, val, dir) {
    const d = doc();
    const rel = /^[+-]/.test(val);
    const n = parseInt(val, 10);
    if (kind === 'page') {
      if (!LY.pages.length) { A().setView('print'); }
      const cur = Math.max(0, LY.pageOf(E.sel.f));
      let i = !val ? cur + (dir || 1) : rel ? cur + n : (() => { const k = LY.pages.findIndex((p) => String(D.fmtNum(p.numVal, p.numFmt)) === String(val).trim()); return k >= 0 ? k : n - 1; })();
      i = L.clamp(i, 0, LY.pages.length - 1);
      const pg = LY.pages[i];
      const f = pg && pg.body.querySelector('.p');
      if (f) { const p = D.byId(d, +f.dataset.pid); A().gotoPos(D.pos(p, +f.dataset.from || 0)); pg.el.scrollIntoView({ block: 'start' }); }
      return true;
    }
    if (kind === 'line') {
      /* line numbers counted in the main story */
      const target = !val ? null : rel ? null : n;
      if (target) {
        let count = 0;
        for (const pg of LY.pages) for (const f of pg.el.querySelectorAll('.pg-body .p')) {
          const lines = LY.lines(f, f.getBoundingClientRect().top, LY.zoom);
          if (count + lines.length >= target) { const ln = lines[target - count - 1]; const p = D.byId(d, +f.dataset.pid); A().gotoPos(D.pos(p, ln ? ln.o : 0)); return true; }
          count += lines.length;
        }
      }
      return false;
    }
    if (kind === 'bookmark') { if (val) return A().gotoBookmark(val); return false; }
    const map = { section: 'section', comment: 'comment', footnote: 'footnote', endnote: 'endnote', field: 'field', table: 'table', graphic: 'graphic', heading: 'heading', equation: 'field' };
    if (val && !rel && n > 0 && kind !== 'section') {
      /* nth item */
      E.setSel(D.pos(D.firstPara(d.main.blocks), 0));
      for (let k = 0; k < n; k++) L.panes.browseTo(map[kind], 1);
      return true;
    }
    if (kind === 'section' && val) { const secs = D.sections(d); const s = secs[L.clamp((rel ? 0 : n) - 1, 0, secs.length - 1)]; const b0 = d.main.blocks[s.from]; const p = b0.t === 'p' ? b0 : D.firstPara([b0]); A().gotoPos(D.pos(p, 0)); return true; }
    L.panes.browseTo(map[kind], dir || (rel && n < 0 ? -1 : 1));
    return true;
  };

  /* ================= the dialog ================= */
  let open = null;
  const state = { text: '', replace: '', matchCase: false, wholeWord: false, wildcards: false, soundsLike: false, allForms: false, dir: 'all', scope: 'main', format: null, replaceFormat: null, highlight: false };
  const hist = { find: [], repl: [] };
  G.find = function (mode) {
    if (open) { open.select(mode); return; }
    if (E.sel && !E.collapsed()) { const [a, b] = E.range(); if (a.p === b.p && b.o - a.o < 255) state.text = O.textRange(a, b, '\n'); }
    const findIn = h('input', { type: 'text', id: 'fr-find', list: 'fr-fhist', value: state.text, style: 'flex:1;min-width:0' });
    const replIn = h('input', { type: 'text', id: 'fr-repl', list: 'fr-rhist', value: state.replace, style: 'flex:1;min-width:0' });
    const fhist = h('datalist', { id: 'fr-fhist' }), rhist = h('datalist', { id: 'fr-rhist' });
    const fillHist = () => { L.clear(fhist); L.clear(rhist); for (const x of hist.find) fhist.appendChild(h('option', { value: x })); for (const x of hist.repl) rhist.appendChild(h('option', { value: x })); };
    fillHist();
    const fmtNote = h('div', { class: 'tp-note fr-fmt' });
    const rfmtNote = h('div', { class: 'tp-note fr-fmt' });
    const showFmt = () => { fmtNote.textContent = state.format ? 'Format: ' + describe(state.format) : ''; rfmtNote.textContent = state.replaceFormat ? 'Format: ' + describe(state.replaceFormat) : ''; };
    const hlChk = ui.check('&Highlight all items found in:', state.highlight, (v) => { state.highlight = v; });
    const scope = ui.select([['main', 'Main Document'], ['all', 'All stories (headers, notes, comments)']], state.scope, (v) => { state.scope = v; }, { id: 'fr-scope' });
    const opt = (k, l) => ui.check(l, state[k], (v) => { state[k] = v; if (k === 'wildcards' && v) { soundsChk.input.checked = state.soundsLike = false; formsChk.input.checked = state.allForms = false; } });
    const caseChk = opt('matchCase', 'Matc&h case'), wholeChk = opt('wholeWord', 'Find whole words onl&y'), wildChk = opt('wildcards', 'Use wildcar&ds'), soundsChk = opt('soundsLike', 'Sounds li&ke (English)'), formsChk = opt('allForms', 'Find all &word forms (English)');
    const dir = ui.select([['all', 'All'], ['down', 'Down'], ['up', 'Up']], state.dir, (v) => { state.dir = v; }, { id: 'fr-dir' });
    const more = h('div', { class: 'col', hidden: !state.more }, ui.group('Search Options', G.f('Search:', dir), caseChk, wholeChk, wildChk, soundsChk, formsChk), h('div', { class: 'tp-btns' }, ui.button('N&o Formatting', () => { if (curTab === 1 && document.activeElement === replIn) state.replaceFormat = null; else { state.format = null; state.replaceFormat = null; } showFmt(); }, { class: 'btn small' }), ui.button('F&ormat ▾', (e) => formatMenu(e.currentTarget, document.activeElement === replIn ? 'replaceFormat' : curTab === 1 && state.lastFocus === 'repl' ? 'replaceFormat' : 'format'), { class: 'btn small' }), ui.button('Sp&ecial ▾', (e) => specialMenu(e.currentTarget, state.lastFocus === 'repl' ? replIn : findIn), { class: 'btn small' })));
    findIn.addEventListener('focus', () => { state.lastFocus = 'find'; });
    replIn.addEventListener('focus', () => { state.lastFocus = 'repl'; });
    const moreBtn = ui.button(state.more ? '<< &Less' : '&More >>', () => { state.more = !state.more; more.hidden = !state.more; moreBtn.innerHTML = state.more ? '&lt;&lt; <u>L</u>ess' : '<u>M</u>ore &gt;&gt;'; });
    /* Go To */
    const gotoKind = G.listBox(GOTO.map(([v, l]) => ({ value: v, label: l })), 'page', { height: 150, onChange: (v) => { gotoLbl.textContent = 'Enter ' + GOTO.find((x) => x[0] === v)[1].toLowerCase() + (v === 'bookmark' ? ' name:' : ' number:'); gotoBtn.innerHTML = gotoIn.value ? 'Go <u>T</u>o' : 'Nex<u>t</u>'; } });
    const gotoLbl = h('label', { for: 'fr-goto', text: 'Enter page number:' });
    const gotoIn = h('input', { type: 'text', id: 'fr-goto', style: 'width:100%' });
    const gotoBtn = ui.button('Nex&t', () => FD.goto(gotoKind.get(), gotoIn.value.trim(), 1));
    gotoIn.addEventListener('input', () => { gotoBtn.innerHTML = gotoIn.value ? 'Go <u>T</u>o' : 'Nex<u>t</u>'; });
    const prevBtn = ui.button('Pre&vious', () => FD.goto(gotoKind.get(), '', -1));
    const findBody = h('div', { class: 'col' }, h('div', { class: 'row', style: 'align-items:center;flex-wrap:nowrap' }, h('label', { for: 'fr-find', html: 'Fi<u>n</u>d what:', style: 'width:80px' }), findIn), fmtNote);
    const findTabBody = h('div', { class: 'col' });
    const replTabBody = h('div', { class: 'col' });
    const gotoTabBody = h('div', { class: 'row' }, h('div', { class: 'col', style: 'width:140px' }, h('label', { text: 'G&o to what:' }), gotoKind), h('div', { class: 'col', style: 'flex:1;min-width:150px' }, gotoLbl, gotoIn, h('div', { class: 'tp-note', text: 'Enter + and − to move relative to the current location. Example: +4 will move forward four items.' })));
    const replRow = h('div', { class: 'col' }, h('div', { class: 'row', style: 'align-items:center;flex-wrap:nowrap' }, h('label', { for: 'fr-repl', html: 'Re<u>p</u>lace with:', style: 'width:80px' }), replIn), rfmtNote);
    let curTab = mode === 'replace' ? 1 : mode === 'goto' ? 2 : 0;
    let dlg = null;
    const foot = h('div', { class: 'tp-btns', style: 'justify-content:flex-end;margin-top:8px' });
    const bFind = ui.button('&Find Next', () => findNext(), { class: 'btn primary' });
    const bRepl = ui.button('&Replace', () => replace());
    const bAll = ui.button('Replace &All', () => replaceAll());
    const bClose = ui.button('Cancel', () => dlg && dlg.close(null));
    function setButtons(t) {
      L.clear(foot);
      if (t === 2) foot.append(prevBtn, gotoBtn, ui.button('Close', () => dlg && dlg.close(null)));
      else { if (t === 1) foot.append(moreBtn, bRepl, bAll, bFind, bClose); else foot.append(moreBtn, bFind, bClose); }
      more.hidden = t === 2 || !state.more;
    }
    const tabs = ui.tabs([
      { label: 'Fin&d', body: findTabBody, onShow: () => { findTabBody.append(findBody, h('div', { class: 'row', style: 'align-items:center' }, hlChk, scope)); setButtons(0); curTab = 0; setTimeout(() => findIn.focus(), 0); } },
      { label: 'Re&place', body: replTabBody, onShow: () => { replTabBody.append(findBody, replRow); setButtons(1); curTab = 1; setTimeout(() => findIn.focus(), 0); } },
      { label: '&Go To', body: gotoTabBody, onShow: () => { setButtons(2); curTab = 2; setTimeout(() => gotoIn.focus(), 0); } },
    ], mode === 'replace' ? 1 : mode === 'goto' ? 2 : 0);
    const opts = () => Object.assign({}, state, { text: findIn.value, replace: replIn.value });
    const remember = () => { const t = findIn.value; if (t && !hist.find.includes(t)) hist.find.unshift(t); hist.find.length = Math.min(hist.find.length, 12); const r = replIn.value; if (r && !hist.repl.includes(r)) hist.repl.unshift(r); fillHist(); state.text = findIn.value; state.replace = replIn.value; };
    const findNext = () => {
      remember();
      try {
        if (state.highlight && curTab === 0) { const n = FD.highlightAll(opts()); A().status(n < 0 ? 'Highlighting is not supported in this browser.' : `Quire found ${n} item${n === 1 ? '' : 's'} matching this criteria.`); if (n > 0) { const hit = FD.next(opts()); if (hit) A().gotoPos(hit.a, hit.b); } return; }
        FD.wrapped = false;
        const hit = FD.next(opts());
        if (!hit) { ui.msg(`Quire has finished searching the document. The search item was not found.`, { icon: 'info' }); return; }
        A().gotoPos(hit.a, hit.b);
        if (FD.wrapped) A().status('Quire reached the end of the document and continued from the beginning.');
      } catch (e) { if (e.message !== 'empty') ui.msg(e.message, { icon: 'warn' }); }
    };
    const replace = () => { remember(); try { const hit = FD.replaceOne(opts()); if (hit) A().gotoPos(hit.a, hit.b); else ui.msg('Quire has finished searching the document.', { icon: 'info' }); } catch (e) { if (e.message !== 'empty') ui.msg(e.message, { icon: 'warn' }); } };
    const replaceAll = () => { remember(); try { const n = FD.replaceAll(opts()); ui.msg(n ? `Quire has completed its search of the document and has made ${n} replacement${n === 1 ? '' : 's'}.` : 'Quire has finished searching the document. The search item was not found.', { icon: 'info' }); } catch (e) { if (e.message !== 'empty') ui.msg(e.message, { icon: 'warn' }); } };
    const body = h('div', { class: 'col' }, tabs, more, foot);
    dlg = ui.dialog({ title: 'Find and Replace', width: 470, body, modeless: true, buttons: [], focus: mode === 'goto' ? gotoIn : findIn, onClose: () => { open = null; FD.clearHighlight(); } });
    dlg.el.querySelector('.dlg-foot').remove();
    dlg.el.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.target.tagName === 'INPUT') { e.preventDefault(); if (curTab === 2) gotoBtn.click(); else bFind.click(); }
    });
    findIn.addEventListener('input', () => { if (FD.hlOpts) FD.clearHighlight(); });
    showFmt();
    open = { select: (m) => { tabs.select(m === 'replace' ? 1 : m === 'goto' ? 2 : 0); }, dlg };
    function describe(f) {
      const parts = [];
      if (f.rPr) for (const [k, v] of Object.entries(f.rPr)) parts.push(k === 'font' ? 'Font: ' + v : k === 'sz' ? v + ' pt' : k === 'b' ? (v ? 'Bold' : 'Not Bold') : k === 'i' ? (v ? 'Italic' : 'Not Italic') : k === 'u' ? (v && v !== 'none' ? 'Underline' : 'No underline') : k === 'color' ? 'Font color: ' + v : k);
      if (f.style) parts.push('Style: ' + D.styleDisplayName(doc().styles[f.style]));
      if (f.highlight != null) parts.push(f.highlight ? 'Highlight' : 'Not Highlight');
      return parts.join(', ');
    }
    function formatMenu(anchor, key) {
      const cur = state[key] || {};
      ui.openMenu([
        { label: '&Font...', run: () => G.font({ target: Object.assign({}, cur.rPr || {}), sample: 'Sample', onOK: (r) => { const rr = Object.assign({}, cur.rPr || {}, r); for (const k in rr) if (rr[k] === undefined) delete rr[k]; state[key] = Object.assign({}, cur, { rPr: rr }); showFmt(); } }) },
        { label: '&Style...', run: () => { const ids = Object.keys(doc().styles).filter((k) => ['paragraph', 'character'].includes(doc().styles[k].type)).sort(); const lb = G.listBox(ids.map((k) => ({ value: k, label: D.styleDisplayName(doc().styles[k]) })), cur.style || 'Normal', { height: 180 }); ui.dialog({ title: key === 'format' ? 'Find Style' : 'Replace Style', width: 280, body: lb, buttons: [{ label: 'OK', primary: true, onClick: () => { state[key] = Object.assign({}, cur, { style: lb.get() }); showFmt(); } }, { label: 'Cancel' }] }); } },
        { label: '&Highlight', run: () => { state[key] = Object.assign({}, cur, { highlight: cur.highlight === true ? false : true }); showFmt(); } },
      ], anchor.getBoundingClientRect());
    }
    function specialMenu(anchor, input) {
      const ins = (code) => { const s = input.selectionStart ?? input.value.length, e = input.selectionEnd ?? s; input.value = input.value.slice(0, s) + code + input.value.slice(e); input.focus(); input.setSelectionRange(s + code.length, s + code.length); };
      const isRepl = input === replIn;
      const wild = state.wildcards;
      const items = isRepl ? [['^p', 'Paragraph Mark'], ['^t', 'Tab Character'], ['^^', 'Caret Character'], ['^&', 'Find What Text'], ['^c', 'Clipboard Contents'], ['^+', 'Em Dash'], ['^=', 'En Dash'], ['^l', 'Manual Line Break'], ['^m', 'Manual Page Break'], ['^n', 'Column Break'], ['^~', 'Nonbreaking Hyphen'], ['^s', 'Nonbreaking Space'], ['^-', 'Optional Hyphen']].concat(wild ? [['\\1', 'Find What Expression']] : [])
        : wild ? [['?', 'Any Character'], ['[-]', 'Character in Range'], ['<', 'Beginning of Word'], ['>', 'End of Word'], ['()', 'Expression'], ['[!]', 'Not'], ['{,}', 'Num Occurrences'], ['@', 'Previous 1 or More'], ['*', '0 or More Characters'], ['^t', 'Tab Character'], ['^p', 'Paragraph Mark'], ['^l', 'Manual Line Break'], ['^m', 'Manual Page Break']]
          : [['^p', 'Paragraph Mark'], ['^t', 'Tab Character'], ['^?', 'Any Character'], ['^#', 'Any Digit'], ['^$', 'Any Letter'], ['^^', 'Caret Character'], ['^+', 'Em Dash'], ['^=', 'En Dash'], ['^g', 'Graphic'], ['^l', 'Manual Line Break'], ['^m', 'Manual Page Break'], ['^n', 'Column Break'], ['^~', 'Nonbreaking Hyphen'], ['^s', 'Nonbreaking Space'], ['^-', 'Optional Hyphen'], ['^w', 'White Space']];
      ui.openMenu(items.map(([c, l]) => ({ label: l, run: () => ins(c) })), anchor.getBoundingClientRect());
    }
  };
})();
