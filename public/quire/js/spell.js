/* Quire — offline Spelling and Grammar.
 * Words are checked against Hunspell's English dictionaries (SCOWL-derived word lists shipped in dict/),
 * suggestions come from edit distance ranked by how common a word is, and grammar is a set of
 * rules in the spirit of Word 2003's checker. Misspellings get red wavy underlines and grammar
 * problems green ones (CSS Custom Highlight API); right-click one for suggestions.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout;
  const { h } = L;
  const ui = L.ui;
  const SP = (L.spell = {});
  const doc = () => D.doc;
  const A = () => L.app;
  const opt = (k, def) => { const o = A() && A().opts; return o && o[k] !== undefined ? o[k] !== false : def !== false; };

  /* ================= lexicon ================= */
  const LEX = (L.lex = { lists: {}, loading: {}, failed: {} });
  const LEVEL = '0123456789';
  /** load a front-coded list: "<prefix-length char><rest><level digit>" per line */
  LEX.load = function (name) {
    if (LEX.lists[name]) return Promise.resolve(LEX.lists[name]);
    if (LEX.loading[name]) return LEX.loading[name];
    LEX.loading[name] = fetch('../common/dict/' + name + '.words').then((r) => { if (!r.ok) throw new Error('HTTP ' + r.status); return r.text(); }).then((txt) => {
      const map = new Map();
      let prev = '';
      const P = '0123456789abcdefghijklmnopqrstuvwxyz';
      for (const ln of txt.split('\n')) {
        if (!ln) continue;
        const k = P.indexOf(ln[0]);
        const w = prev.slice(0, k) + ln.slice(1, -1);
        map.set(w, LEVEL.indexOf(ln[ln.length - 1]));
        prev = w;
      }
      /* lower-case index so "Paris" at sentence start and "PARIS" both resolve */
      const lower = new Map();
      for (const [w, lv] of map) { const l = w.toLowerCase(); if (!lower.has(l) || lower.get(l).lv > lv) lower.set(l, { w, lv }); }
      LEX.lists[name] = { map, lower };
      L.bus.emit('lexicon-loaded', name);
      return LEX.lists[name];
    }).catch((e) => { LEX.failed[name] = String(e && e.message || e); LEX.loading[name] = null; throw e; });
    return LEX.loading[name];
  };
  /** which word list a language tag uses (null: no English proofing tools for it) */
  LEX.listFor = function (lang) {
    lang = String(lang || 'en-US').toLowerCase();
    if (!/^en\b/.test(lang)) return null;
    return /^en-(gb|au|nz|ie|za|in|sg|hk|jm|bz|tt|zw|ph|my)/.test(lang) ? 'en_GB' : 'en_US';
  };
  LEX.ready = (name) => !!LEX.lists[name];

  /* ================= checking words ================= */
  const custom = () => { if (!SP._custom) SP._custom = new Set((L.store.get('customDict', []) || []).map((x) => String(x).toLowerCase())); return SP._custom; };
  SP.reloadCustom = () => { SP._custom = null; SP.refreshSoon(); };
  const ignoredAll = new Set();
  const cache = new Map();
  const isUpper = (w) => w === w.toUpperCase() && w !== w.toLowerCase();
  const isCap = (w) => w[0] === w[0].toUpperCase() && w.slice(1) === w.slice(1).toLowerCase();
  /** true when the word is spelled correctly (or should not be checked) */
  SP.okWord = function (word, list) {
    const lex = LEX.lists[list];
    if (!lex) return true;
    let w = word.replace(/’/g, "'");
    if (w.length < 2) return true;
    if (/\d/.test(w)) return opt('ignoreNum', true) || /^\d/.test(w);
    if (isUpper(w) && opt('ignoreUpper', true)) return true;
    const key = list + '\u0000' + w;
    if (cache.has(key)) return cache.get(key);
    const r = checkWord(w, lex);
    cache.set(key, r);
    return r;
  };
  function checkWord(w, lex) {
    const lw = w.toLowerCase();
    if (custom().has(lw) || ignoredAll.has(lw)) return true;
    if (lex.map.has(w)) return true;
    /* sentence-initial capitals and ALL CAPS accept the lower-case entry; a capitalised entry ("Paris") needs its capital */
    if ((isCap(w) || isUpper(w)) && lex.map.has(lw)) return true;
    if (isUpper(w)) { const e = lex.lower.get(lw); if (e) return true; }
    if (isCap(w)) { const e = lex.lower.get(lw); if (e && isCap(e.w)) return true; }
    /* possessive of a known word (James's, CEO's) */
    const m = /^(.+?)'s?$/i.exec(w);
    if (m && m[1].length > 1 && (lex.map.has(m[1]) || lex.lower.has(m[1].toLowerCase()))) return true;
    /* hyphenated compounds are fine when every part is */
    if (w.includes('-')) return w.split('-').every((part) => !part || part.length < 2 || checkWord(part, lex));
    /* accented spellings of listed words: café, naïve, résumé, façade */
    const plain = w.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
    if (plain !== w) return checkWord(plain, lex);
    /* AutoCorrect replacements the user taught are accepted words too */
    return false;
  }

  /* ================= suggestions ================= */
  const ALPHA = "abcdefghijklmnopqrstuvwxyz'";
  function edits1(w) {
    const out = new Set();
    for (let i = 0; i <= w.length; i++) {
      const a = w.slice(0, i), b = w.slice(i);
      if (b) out.add(a + b.slice(1));
      if (b.length > 1) out.add(a + b[1] + b[0] + b.slice(2));
      for (const c of ALPHA) { if (b) out.add(a + c + b.slice(1)); out.add(a + c + b); }
    }
    out.delete(w);
    return out;
  }
  const KEYS = ['qwertyuiop', 'asdfghjkl', 'zxcvbnm'];
  const keyPos = {};
  KEYS.forEach((row, r) => { for (let c = 0; c < row.length; c++) keyPos[row[c]] = [r, c + r * 0.5]; });
  const near = (a, b) => { const p = keyPos[a], q = keyPos[b]; return p && q && Math.abs(p[0] - q[0]) <= 1 && Math.abs(p[1] - q[1]) <= 1.5; };
  /** Damerau-Levenshtein distance with cheaper adjacent-key and doubled-letter mistakes */
  function dist(a, b) {
    const n = a.length, m = b.length;
    const d = [];
    for (let i = 0; i <= n; i++) { d[i] = [i]; }
    for (let j = 0; j <= m; j++) d[0][j] = j;
    for (let i = 1; i <= n; i++) for (let j = 1; j <= m; j++) {
      const sub = a[i - 1] === b[j - 1] ? 0 : near(a[i - 1], b[j - 1]) ? 0.8 : 1;
      let v = Math.min(d[i - 1][j] + (i > 1 && a[i - 1] === a[i - 2] ? 0.6 : 1), d[i][j - 1] + (j > 1 && b[j - 1] === b[j - 2] ? 0.6 : 1), d[i - 1][j - 1] + sub);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) v = Math.min(v, d[i - 2][j - 2] + 0.7);
      d[i][j] = v;
    }
    return d[n][m];
  }
  const restoreCase = (sugg, orig) => (isUpper(orig) && orig.length > 1 ? sugg.toUpperCase() : /^\p{Lu}/u.test(orig) && !/^\p{Lu}/u.test(sugg) ? sugg[0].toUpperCase() + sugg.slice(1) : sugg);
  SP.suggest = function (word, list, max) {
    max = max || 8;
    const lex = LEX.lists[list];
    if (!lex) return [];
    const w = word.replace(/’/g, "'");
    const lw = w.toLowerCase();
    const found = new Map();
    const consider = (cand, base) => {
      const e = lex.lower.get(cand);
      if (!e) return;
      const real = e.w;
      if (real === w) return;
      const score = dist(lw, cand) * 10 + e.lv * 1.6 + base + (real.length < 3 ? 4 : 0);
      if (!found.has(real) || found.get(real) > score) found.set(real, score);
    };
    /* AutoCorrect knows many typical slips */
    const ac = L.autocorrect ? L.autocorrect.list() : null;
    if (ac && ac.has(lw)) { const r = ac.get(lw); if (!/\s/.test(r)) found.set(r, -100); else found.set(r, -50); }
    const e1 = edits1(lw);
    for (const c of e1) consider(c, 0);
    if (found.size < 6 && lw.length <= 14) {
      for (const c of e1) { if (lex.lower.has(c)) continue; for (const c2 of edits1(c)) if (c2.length > 1) consider(c2, 6); }
    }
    /* run-together words: "alot" → "a lot", "infact" → "in fact" */
    for (let i = 1; i < lw.length; i++) {
      const a = lw.slice(0, i), b = lw.slice(i);
      if ((a.length > 1 || a === 'a' || a === 'i') && b.length > 1 && lex.lower.has(a) && lex.lower.has(b) && lex.lower.get(a).lv <= 4 && lex.lower.get(b).lv <= 5) {
        const score = 5 + (lex.lower.get(a).lv + lex.lower.get(b).lv) * 1.2;
        const s = (a === 'i' ? 'I' : lex.lower.get(a).w) + ' ' + lex.lower.get(b).w;
        if (!found.has(s) || found.get(s) > score) found.set(s, score);
      }
    }
    return Array.from(found.entries()).sort((x, y) => x[1] - y[1]).slice(0, max).map(([s]) => (/^\p{Lu}/u.test(s) && !isUpper(w) ? s : restoreCase(s, w)));
  };

  /* ================= tokenising ================= */
  /** words of a text: [{w, a, b}] (letters with inner apostrophes and hyphens) */
  function words(t) {
    const out = [];
    const re = /[\p{L}\p{M}\p{N}]+(?:['’\-][\p{L}\p{M}\p{N}]+)*/gu;
    let m;
    while ((m = re.exec(t))) out.push({ w: m[0], a: m.index, b: m.index + m[0].length });
    return out;
  }
  /** spans that look like web addresses, e-mail addresses or file paths */
  function skipZones(t) {
    const z = [];
    const re = /\b(?:https?:\/\/|ftp:\/\/|www\.)\S+|\b[\w.+-]+@[\w-]+\.[\w.-]+|(?:[A-Za-z]:\\|\\\\)[^\s]+|\b\S+\.(?:docx?|xlsx?|pdf|txt|html?|png|jpe?g|gif|zip)\b/gi;
    let m;
    while ((m = re.exec(t))) z.push([m.index, m.index + m[0].length]);
    return z;
  }

  /* ================= grammar rules ================= */
  const AN_EXCEPT_VOWEL = /^(uni(?!n)|use|usu|uti|ure|uro|eu|ewe|one|once|ouija|u\b|ubiq|ufo|uk\b|us\b|usa\b)/i;
  const AN_EXCEPT_CONS = /^(hour|honest|honor|honour|heir|herb(?!al)|hors d)/i;
  const ACRONYM_VOWEL_SOUND = /^[AEFHILMNORSX](?:[A-Z]{1,4})?$/;
  function needsAn(next) {
    if (/^\d/.test(next)) return /^(8|11|18|80|800)\b/.test(next);
    if (/^[A-Z]{2,5}$/.test(next)) return ACRONYM_VOWEL_SOUND.test(next);
    if (AN_EXCEPT_CONS.test(next)) return true;
    if (AN_EXCEPT_VOWEL.test(next)) return false;
    return /^[aeiou]/i.test(next);
  }
  const PLURAL_SUBJ = /^(we|they|you)$/i, SING_SUBJ = /^(he|she|it)$/i;
  const CONFUSED = [
    [/\b(could|should|would|must|might)\s+of\b/gi, (m) => m[1] + ' have', 'Commonly Confused Words', 'Use "have" after could, should, would, might and must.'],
    [/\b(more|less|better|worse|rather|other|greater|larger|smaller|higher|lower|faster|slower|older|younger|bigger)\s+then\b/gi, (m) => m[1] + ' than', 'Commonly Confused Words', 'Use "than" in comparisons.'],
    [/\byour\s+(welcome|right|wrong|not|going|being)\b/gi, (m) => "you're " + m[1], 'Possessives and Plurals', '"Your" is possessive; "you\'re" means "you are".'],
    [/\b(its)\s+(a|an|the|been|not|going|very|so|too|time)\b/gi, (m) => "it's " + m[2], 'Possessives and Plurals', '"Its" is possessive; "it\'s" means "it is" or "it has".'],
    [/\btheir\s+(is|are|was|were)\b/gi, (m) => 'there ' + m[1], 'Commonly Confused Words', '"Their" is possessive; use "there" with is, are, was and were.'],
    [/\b(irregardless)\b/gi, () => 'regardless', 'Commonly Confused Words', '"Irregardless" is nonstandard.'],
  ];
  /** grammar issues within one paragraph text: [{a, b, rule, suggestions, explain}] */
  function grammarIssues(t, o) {
    o = o || {};
    const out = [];
    if (!opt('grammar', true)) return out;
    /* sentence rules only make sense in prose: not in table cells or one-line labels without a full stop */
    const prose = !o.inCell && /[.!?]["'”’)]?(\s|$)/.test(t) && t.trim().split(/\s+/).length > 3;
    const ws = words(t);
    /* repeated words */
    for (let i = 1; i < ws.length; i++) {
      const p = ws[i - 1], c = ws[i];
      if (p.w.toLowerCase() === c.w.toLowerCase() && /^[ \t]+$/.test(t.slice(p.b, c.a)) && /\p{L}/u.test(c.w) && !/^(that|had|is|bye|so|no|ha|very|bora|can|in|walla|chop|pom|tut|aye|do|knock|well|now)$/i.test(c.w)) out.push({ a: c.a, b: c.b, delA: p.b, rule: 'Repeated Word', kind: 'repeat', suggestions: ['Delete Repeated Word'], explain: 'The word appears twice in a row.' });
    }
    /* a / an */
    for (let i = 0; i + 1 < ws.length; i++) {
      const art = ws[i].w, nx = ws[i + 1];
      if (!/^(a|an)$/i.test(art) || !/^\s+$/.test(t.slice(ws[i].b, nx.a))) continue;
      if (/^[A-Z]$/.test(art) && i > 0 && !/[.!?]\s*$/.test(t.slice(0, ws[i].a))) continue; /* "Exhibit A" */
      const an = needsAn(nx.w);
      if (art.toLowerCase() === 'a' && an) out.push({ a: ws[i].a, b: ws[i].b, rule: 'Use of "a" or "an"', suggestions: [art === 'A' ? 'An' : 'an'], explain: 'Use "an" before a word that starts with a vowel sound.' });
      else if (art.toLowerCase() === 'an' && !an && /^[a-z]/i.test(nx.w)) out.push({ a: ws[i].a, b: ws[i].b, rule: 'Use of "a" or "an"', suggestions: [art === 'An' ? 'A' : 'a'], explain: 'Use "a" before a word that starts with a consonant sound.' });
    }
    /* subject–verb agreement for pronouns */
    for (let i = 0; i + 1 < ws.length; i++) {
      const s = ws[i].w, v = ws[i + 1];
      if (!/^\s+$/.test(t.slice(ws[i].b, v.a))) continue;
      const lv = v.w.toLowerCase();
      let fix = null;
      if (/^i$/i.test(s) && /^(is|are|has)$/.test(lv)) fix = { is: 'am', are: 'am', has: 'have' }[lv];
      else if (PLURAL_SUBJ.test(s) && /^(is|was|has|does|doesn't)$/.test(lv) && !(s.toLowerCase() === 'you' && lv === 'was')) fix = { is: 'are', was: 'were', has: 'have', does: 'do', "doesn't": "don't" }[lv];
      else if (SING_SUBJ.test(s) && /^(are|have|do|don't)$/.test(lv)) fix = { are: 'is', have: 'has', do: 'does', "don't": "doesn't" }[lv];
      if (fix) {
        /* questions ("Do they", "Have we") and "it do" in odd constructions are left alone */
        const before = t.slice(Math.max(0, ws[i].a - 30), ws[i].a);
        if (/\b(do|does|did|will|would|can|could|should|let|make|made|help|see|saw|watch|hear|heard|have|had)\s+$/i.test(before)) continue;
        out.push({ a: v.a, b: v.b, rule: 'Subject-Verb Agreement', suggestions: [fix], explain: 'The verb does not agree with its subject.' });
      }
    }
    /* lower-case pronoun i */
    for (const w of ws) if (w.w === 'i' && !/[.\\/]$/.test(t.slice(0, w.a)) && !/^[.\\/]/.test(t.slice(w.b))) out.push({ a: w.a, b: w.b, rule: 'Capitalization', suggestions: ['I'], explain: 'The pronoun "I" is always capitalized.' });
    /* first letter of a sentence */
    if (prose) {
      const re = /(^|[.!?]["'”’)\]]*\s+)(\p{Ll}[\p{L}'’]*)/gu;
      let m;
      while ((m = re.exec(t))) {
        const at = m.index + m[1].length;
        const prevWord = (/(\S+)\s+$/.exec(t.slice(0, at)) || [])[1] || '';
        if (m[1] && (/^(e\.g|i\.e|etc|vs|cf|al|approx|no|fig|p|pp|vol|ch|sec)\.$/i.test(prevWord) || /^\S\.$/.test(prevWord) || /\.\S+\.$/.test(prevWord))) continue;
        if (!m[1] && /^[\s(“"'‘]*$/.test(t.slice(0, at)) === false) continue;
        if (/^(e\.g|i\.e|www|http|iphone|ipad|ipod|ebay)/i.test(m[2])) continue;
        out.push({ a: at, b: at + m[2].length, rule: 'Capitalization', suggestions: [m[2][0].toUpperCase() + m[2].slice(1)], explain: 'Start a sentence with a capital letter.' });
      }
    }
    /* spacing and punctuation */
    let m;
    const sp = /([\p{L}\p{N},;:.!?)]) {2,}(?=[\p{L}\p{N}(])/gu;
    while ((m = sp.exec(t))) { if (/[.!?]/.test(m[1])) continue; out.push({ a: m.index + 1, b: m.index + m[0].length, rule: 'Spacing', suggestions: [' '], explain: 'There is more than one space between these words.' }); }
    const before = /([\p{L}\p{N}]) +([,;:.!?])(?=\s|$)/gu;
    while ((m = before.exec(t))) { if (m[2] === '.' && /\.\.$/.test(t.slice(0, m.index + m[0].length + 2))) continue; out.push({ a: m.index + 1, b: m.index + m[0].length, rule: 'Punctuation', suggestions: [m[2]], explain: 'Punctuation follows the word directly.' }); }
    const after = /([\p{Ll}]{2,})([,;])(?=\p{L})/gu;
    while ((m = after.exec(t))) out.push({ a: m.index + m[1].length, b: m.index + m[0].length, rule: 'Punctuation', suggestions: [m[2] + ' '], explain: 'Put a space after a comma or semicolon.' });
    const dbl = /([,;:])\1+|(?<!\.)\.\.(?!\.)/g;
    while ((m = dbl.exec(t))) out.push({ a: m.index, b: m.index + m[0].length, rule: 'Punctuation', suggestions: [m[0][0]], explain: 'This punctuation mark is repeated.' });
    for (const [re, fix, rule, explain] of CONFUSED) {
      if (!fix) continue;
      re.lastIndex = 0;
      while ((m = re.exec(t))) out.push({ a: m.index, b: m.index + m[0].length, rule, suggestions: [restoreCase(fix(m), m[0])], explain });
    }
    /* overlapping findings: keep the first */
    out.sort((x, y) => x.a - y.a || y.b - x.b);
    const keep = [];
    for (const g of out) if (!keep.some((k) => g.a < k.b && k.a < g.b)) keep.push(g);
    return keep;
  }
  SP.grammarIssues = grammarIssues;

  /* ================= paragraph language ================= */
  function paraLang(p) {
    const d = doc();
    const it = p.runs.find((r) => r.t === 'text' && r.text.trim());
    const rp = it ? D.rProps(d, p, it) : D.rProps(d, p, { t: 'text', text: '', rPr: p.rPr || {} });
    return rp.lang || (d.defaults && d.defaults.rPr && d.defaults.rPr.lang) || 'en-US';
  }
  SP.paraList = (p) => LEX.listFor(paraLang(p));

  /** all issues in a paragraph (model text) */
  function paraIssues(p, list, withGrammar) {
    const t = D.ptext(p);
    const out = [];
    if (!list) return out;
    const zones = opt('ignoreUrl', true) ? skipZones(t) : [];
    const noproof = new Set();
    let o = 0;
    for (const it of p.runs) { const n = D.ilen(it); if (it.rPr && (it.rPr.noProof || it.rPr.del || it.rPr.hidden)) for (let k = o; k < o + n; k++) noproof.add(k); o += n; }
    for (const w of words(t)) {
      if (zones.some(([a, b]) => w.a < b && a < w.b)) continue;
      if (noproof.has(w.a)) continue;
      if (!SP.okWord(w.w, list)) out.push({ p, a: w.a, b: w.b, kind: 'spelling', word: w.w, rule: 'Not in Dictionary' });
    }
    if (withGrammar) for (const g of grammarIssues(t, { inCell: !!D.cellOf(doc(), p) })) if (!noproof.has(g.a) && !out.some((s) => s.a < g.b && g.a < s.b)) out.push(Object.assign({ p, kind: g.kind || 'grammar', word: t.slice(g.a, g.b) }, g));
    return out.sort((x, y) => x.a - y.a);
  }

  /* ================= underlines while typing ================= */
  const HL = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight === 'function';
  let spellHL = null, gramHL = null;
  if (HL) { spellHL = new Highlight(); gramHL = new Highlight(); CSS.highlights.set('quire-spell', spellHL); CSS.highlights.set('quire-grammar', gramHL); }
  SP.ignoredOnce = new Set(); /* "p.id:a:word" */
  let timer = 0;
  SP.refreshSoon = function (ms) { clearTimeout(timer); timer = setTimeout(SP.refresh, ms == null ? 350 : ms); };
  /** text of one rendered paragraph fragment with a map back to text nodes */
  function fragText(frag) {
    const nodes = [];
    let text = '';
    const walk = (n) => {
      for (let c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3) { nodes.push({ node: c, at: text.length }); text += c.data; }
        else if (c.nodeType === 1) {
          if (c.contentEditable === 'false' || c.classList.contains('lbl') || c.classList.contains('fcode') || c.getAttribute('spellcheck') === 'false') { text += c.nodeName === 'BR' ? '\n' : '\u0001'; continue; }
          if (c.nodeName === 'BR') { text += '\n'; continue; }
          walk(c);
        }
      }
    };
    walk(frag);
    return { text, nodes };
  }
  function rangeFor(map, a, b) {
    const pos = (o, end) => {
      for (let i = map.nodes.length - 1; i >= 0; i--) { const n = map.nodes[i]; if (o > n.at || (o === n.at && !end) || (o === n.at && i === 0)) { const k = o - n.at; if (k <= n.node.data.length) return [n.node, k]; } }
      return null;
    };
    const s = pos(a, false), e = pos(b, true);
    if (!s || !e) return null;
    const r = document.createRange();
    try { r.setStart(s[0], s[1]); r.setEnd(e[0], e[1]); } catch (err) { return null; }
    return r;
  }
  SP.refresh = function () {
    if (!HL || !LY.root) return;
    spellHL.clear(); gramHL.clear();
    if (!opt('spell', true) || SP.hidden) return;
    const d = doc();
    if (!d) return;
    const sc = LY.scroller.getBoundingClientRect();
    const sel = window.getSelection();
    const caretNode = sel && sel.isCollapsed ? sel.focusNode : null, caretOff = sel ? sel.focusOffset : -1;
    const wantGrammar = opt('grammar', true) && opt('grammarAsType', true);
    const frags = LY.root.querySelectorAll('.p[data-pid]');
    const need = new Set();
    for (const f of frags) {
      const r = f.getBoundingClientRect();
      if (r.bottom < sc.top - 200 || r.top > sc.bottom + 200) continue;
      const p = D.byId(d, +f.dataset.pid);
      if (!p || p.t !== 'p') continue;
      const list = SP.paraList(p);
      if (!list) continue;
      if (!LEX.ready(list)) { need.add(list); continue; }
      const pc = f.querySelector('.pc') || f;
      const map = fragText(pc);
      const t = map.text;
      const zones = opt('ignoreUrl', true) ? skipZones(t) : [];
      for (const w of words(t)) {
        if (zones.some(([a, b]) => w.a < b && a < w.b)) continue;
        if (SP.okWord(w.w, list)) continue;
        const rg = rangeFor(map, w.a, w.b);
        if (!rg) continue;
        if (caretNode && rg.endContainer === caretNode && rg.endOffset === caretOff) continue; /* still being typed */
        if (SP.ignoredOnce.has(p.id + ':' + w.w)) continue;
        spellHL.add(rg);
      }
      const bad = [];
      for (const r of spellHL) if (pc.contains(r.startContainer)) bad.push(r);
      if (wantGrammar) for (const g of grammarIssues(t.replace(/\u0001/g, ' '), { inCell: !!f.closest('td') })) {
        if (SP.ignoredOnce.has(p.id + ':g:' + t.slice(g.a, g.b))) continue;
        const rg = rangeFor(map, g.a, g.b);
        /* a misspelling already underlined in red is not underlined again in green */
        if (rg && !bad.some((r) => r.compareBoundaryPoints(Range.START_TO_END, rg) > 0 && r.compareBoundaryPoints(Range.END_TO_START, rg) < 0)) gramHL.add(rg);
      }
    }
    for (const list of need) if (!LEX.failed[list]) LEX.load(list).then(() => SP.refreshSoon(0)).catch(() => {});
  };
  L.bus.on('layout-done', () => SP.refreshSoon());
  L.bus.on('doc-loaded', () => { SP.ignoredOnce.clear(); SP.refreshSoon(100); });
  L.bus.on('app-ready', () => {
    LY.scroller.addEventListener('scroll', () => SP.refreshSoon(150), { passive: true });
    if (LY.root) LY.root.spellcheck = false;
    /* warm the default dictionary so the first check is instant */
    setTimeout(() => LEX.load('en_US').then(() => SP.refreshSoon(0)).catch(() => {}), 400);
  });

  /* ================= right-click suggestions ================= */
  /** the misspelled word or grammar finding under the mouse, as context menu items (or null) */
  SP.suggestionsAt = function (e) {
    if (!opt('spell', true) || !e) return null;
    let node = null, off = 0;
    if (document.caretRangeFromPoint) { const r = document.caretRangeFromPoint(e.clientX, e.clientY); if (r) { node = r.startContainer; off = r.startOffset; } }
    else if (document.caretPositionFromPoint) { const r = document.caretPositionFromPoint(e.clientX, e.clientY); if (r) { node = r.offsetNode; off = r.offset; } }
    if (!node) return null;
    const pos = LY.domToPos(node, off);
    if (!pos) return null;
    const p = pos.p;
    const list = SP.paraList(p);
    if (!list || !LEX.ready(list)) return null;
    const t = D.ptext(p);
    const hit = paraIssues(p, list, opt('grammar', true)).find((x) => x.a <= pos.o && pos.o <= x.b && !SP.ignoredOnce.has(p.id + ':' + x.word) && !SP.ignoredOnce.has(p.id + ':g:' + x.word));
    if (!hit) return null;
    const replace = (rep) => {
      E.edit(hit.kind === 'spelling' ? 'Spelling Change' : 'Grammar Change', () => {
        if (hit.kind === 'repeat') { O.deleteRange(D.pos(p, hit.delA), D.pos(p, hit.b)); return D.pos(p, hit.delA); }
        const it = D.itemAfter(p, hit.a);
        O.deleteRange(D.pos(p, hit.a), D.pos(p, hit.b));
        return O.insertText(D.pos(p, hit.a), rep, it && it.rPr ? L.clone(it.rPr) : {});
      });
      SP.refreshSoon(0);
    };
    if (hit.kind === 'spelling') {
      const sugg = SP.suggest(hit.word, list, 5);
      const items = sugg.length ? sugg.map((s) => ({ label: s.replace(/&/g, '&&'), bold: true, run: () => replace(s) })) : [{ label: '(no spelling suggestions)', disabled: true }];
      items.push('-',
        { label: 'I&gnore All', run: () => { ignoredAll.add(hit.word.toLowerCase()); cache.clear(); SP.refreshSoon(0); } },
        { label: '&Add to Dictionary', run: () => SP.addWord(hit.word) },
        { label: 'AutoCorr&ect', disabled: !sugg.length, sub: sugg.map((s) => ({ label: s.replace(/&/g, '&&'), run: () => { if (L.autocorrect) L.autocorrect.setEntry(hit.word, s); replace(s); } })).concat(['-', { label: '&AutoCorrect Options...', run: () => L.dlg.autocorrect() }]) },
        { label: '&Language', sub: [{ label: '&Set Language...', run: () => L.dlg.language && L.dlg.language() }] },
        { label: '&Spelling...', run: () => { E.setSel(D.pos(p, hit.a)); SP.check(); } },
        { label: '&Look Up...', run: () => { if (L.panes) { L.panes.researchQuery = hit.word; L.panes.researchAuto = true; L.app.toggleTask(true); L.panes.task.show('research'); } } });
      return items;
    }
    const items = (hit.suggestions || []).map((s) => ({ label: (s === ' ' ? '(single space)' : s).replace(/&/g, '&&'), bold: true, run: () => replace(s === 'Delete Repeated Word' ? '' : s) }));
    items.push('-', { label: '&Ignore Once', run: () => { SP.ignoredOnce.add(p.id + ':g:' + hit.word); SP.refreshSoon(0); } }, { label: '&Grammar...', run: () => { E.setSel(D.pos(p, hit.a)); SP.check(); } }, { label: '&About This Sentence', run: () => ui.msg(hit.rule + '\n\n' + (hit.explain || ''), { icon: 'info', title: 'Grammar' }) });
    return items;
  };
  SP.addWord = function (w) {
    const s = custom();
    s.add(String(w).toLowerCase());
    L.store.set('customDict', Array.from(s));
    cache.clear();
    SP.refreshSoon(0);
  };

  /* ================= the Spelling and Grammar dialog ================= */
  SP.check = async function () {
    const d = doc();
    if (!E.sel) return;
    D.reindex(d);
    const st = E.story && E.story() && E.story().paras ? E.story() : d.main;
    const paras = (st.paras || d.main.paras).slice();
    const startP = E.sel.f.p;
    const from = Math.max(0, paras.indexOf(startP));
    const ordered = paras.slice(from).concat(paras.slice(0, from));
    const lists = new Set(ordered.map(SP.paraList).filter(Boolean));
    if (!lists.size) { ui.msg('The text is not marked as English, so it was not checked. Use Tools ▸ Language ▸ Set Language to mark it.', { icon: 'info', title: 'Spelling and Grammar' }); return; }
    L.ui.busy(true, 'Loading the dictionary...');
    try { for (const l of lists) await LEX.load(l); }
    catch (e) { L.ui.busy(false); ui.msg('The dictionary could not be loaded (' + (e.message || e) + '). Quire needs to be served from a web server, not opened as a local file, for spelling to work.', { icon: 'warn', title: 'Spelling and Grammar' }); return; }
    L.ui.busy(false);
    const withGrammar = opt('grammar', true);
    const issues = [];
    for (const p of ordered) {
      const list = SP.paraList(p);
      for (const x of paraIssues(p, list, withGrammar)) {
        if (p === startP && from > 0 && x.b <= E.sel.f.o && ordered[0] === startP) { x.wrap = true; }
        issues.push(x);
      }
    }
    /* issues before the caret in the starting paragraph come last */
    const head = issues.filter((x) => !x.wrap), tail = issues.filter((x) => x.wrap);
    const all = head.concat(tail);
    if (!all.length) { ui.msg('The spelling and grammar check is complete.', { icon: 'info', title: 'Spelling and Grammar' }); return; }
    run(all, withGrammar);
  };
  function run(issues, withGrammar) {
    let idx = -1;
    let cur = null;
    const ctxBox = h('div', { class: 'sg-context', contenteditable: 'false' });
    const label = h('div', { class: 'cd-cap' });
    const sugg = h('select', { size: 6, id: 'sg-sugg', style: 'width:100%' });
    const changes = new Map();
    const ignoreRules = new Set();
    let gram = withGrammar;
    const next = () => {
      for (;;) {
        idx++;
        if (idx >= issues.length) { dlg.close(0); SP.refreshSoon(0); ui.msg('The spelling and grammar check is complete.', { icon: 'info', title: 'Spelling and Grammar' }); return; }
        cur = issues[idx];
        if (!D.info(doc(), cur.p)) continue;
        const t = D.ptext(cur.p);
        if (t.slice(cur.a, cur.b) !== cur.word) { const i = t.indexOf(cur.word); if (i < 0) continue; cur.a = i; cur.b = i + cur.word.length; }
        if (cur.kind === 'spelling') {
          const list = SP.paraList(cur.p);
          if (SP.okWord(cur.word, list)) continue;
          if (changes.has(cur.word)) { apply(changes.get(cur.word)); continue; }
          cur.suggestions = cur.suggestions || SP.suggest(cur.word, list, 8);
        } else {
          if (!gram || ignoreRules.has(cur.rule)) continue;
        }
        break;
      }
      show();
    };
    const show = () => {
      const t = D.ptext(cur.p);
      label.textContent = cur.kind === 'spelling' ? 'Not in Dictionary:' : cur.kind === 'repeat' ? 'Repeated Word:' : cur.rule + ':';
      L.clear(ctxBox);
      const s0 = Math.max(0, t.lastIndexOf('.', cur.a - 2) + 1), e0 = Math.min(t.length, (t.indexOf('.', cur.b) + 1) || t.length);
      ctxBox.append(document.createTextNode(t.slice(s0, cur.a).replace(/^\s+/, '')), h('span', { class: cur.kind === 'spelling' ? 'sg-bad' : 'sg-gram', text: t.slice(cur.a, cur.b) }), document.createTextNode(t.slice(cur.b, e0)));
      L.clear(sugg);
      const list = cur.suggestions || [];
      if (!list.length) sugg.appendChild(h('option', { text: cur.kind === 'spelling' ? '(No Spelling Suggestions)' : '(No Suggestions)', disabled: true }));
      for (const s of list) sugg.appendChild(h('option', { value: s, text: s === ' ' ? '(single space)' : s }));
      sugg.selectedIndex = list.length ? 0 : -1;
      L.app.gotoPos(D.pos(cur.p, cur.a), D.pos(cur.p, cur.b));
      const isSp = cur.kind === 'spelling';
      setLabel(btnIgnoreAll, isSp ? 'I&gnore All' : 'Ignore Rule');
      setLabel(btnAdd, isSp ? '&Add to Dictionary' : '&Next Sentence');
      btnChangeAll.disabled = !isSp;
      btnAuto.disabled = !isSp || !list.length;
      btnExplain.hidden = isSp;
    };
    const setLabel = (b, s) => { b.innerHTML = ''; b.appendChild(h('span', { html: s.replace(/&(.)/, '<u>$1</u>') })); b._label = s; };
    const apply = (rep) => {
      const p = cur.p, a = cur.a, b = cur.b;
      if (cur.kind === 'repeat') {
        /* the finding covers the space and the second word */
        const s0 = cur.delA != null ? cur.delA : a;
        E.edit('Grammar', () => { O.deleteRange(D.pos(p, s0), D.pos(p, b)); return D.pos(p, s0); });
        for (const x of issues.slice(idx + 1)) if (x.p === p && x.a >= b) { x.a -= b - s0; x.b -= b - s0; }
        return;
      }
      E.edit(cur.kind === 'spelling' ? 'Spelling' : 'Grammar', () => { const it = D.itemAfter(p, a); O.deleteRange(D.pos(p, a), D.pos(p, b)); return O.insertText(D.pos(p, a), rep, it && it.rPr ? L.clone(it.rPr) : {}); });
      for (const x of issues.slice(idx + 1)) if (x.p === p && x.a >= b) { x.a += rep.length - (b - a); x.b += rep.length - (b - a); }
    };
    const pick = () => (sugg.selectedIndex >= 0 && sugg.value ? sugg.value : null);
    const btnIgnore = ui.button('&Ignore Once', () => { SP.ignoredOnce.add(cur.p.id + ':' + (cur.kind === 'spelling' ? '' : 'g:') + cur.word); next(); });
    const btnIgnoreAll = ui.button('I&gnore All', () => { if (cur.kind === 'spelling') { ignoredAll.add(cur.word.toLowerCase()); cache.clear(); } else ignoreRules.add(cur.rule); next(); });
    const btnAdd = ui.button('&Add to Dictionary', () => { if (cur.kind === 'spelling') SP.addWord(cur.word); else { const p = cur.p; while (idx + 1 < issues.length && issues[idx + 1].p === p) idx++; } next(); });
    const btnChange = ui.button('&Change', () => { const v = pick(); if (v == null && cur.kind !== 'repeat') return; apply(v === 'Delete Repeated Word' ? '' : v); next(); });
    const btnChangeAll = ui.button('Change A&ll', () => { const v = pick(); if (v == null) return; changes.set(cur.word, v); apply(v); next(); });
    const btnAuto = ui.button('AutoCo&rrect', () => { const v = pick(); if (v == null) return; if (L.autocorrect) L.autocorrect.setEntry(cur.word, v); apply(v); next(); });
    const btnExplain = ui.button('E&xplain...', () => ui.msg(cur.rule + '\n\n' + (cur.explain || ''), { icon: 'info', title: 'Grammar' }));
    sugg.addEventListener('dblclick', () => btnChange.click());
    const side = h('div', { class: 'col', style: 'width:140px' }, btnIgnore, btnIgnoreAll, btnAdd, h('div', { style: 'height:16px' }), btnChange, btnChangeAll, btnAuto, btnExplain);
    for (const b of side.querySelectorAll('.btn')) b.style.width = '100%';
    const langName = (l) => (l === 'en_GB' ? 'English (U.K.)' : 'English (U.S.)');
    const langLbl = h('span', { text: langName(SP.paraList(issues[0].p)) });
    const chkG = ui.check('Check &grammar', gram, (v) => { gram = v; });
    const left = h('div', { class: 'col', style: 'flex:1;min-width:0' }, label, ctxBox, h('label', { for: 'sg-sugg', text: 'Suggestions:' }), sugg, h('div', { class: 'field' }, h('label', { text: 'Dictionary language:' }), langLbl), chkG);
    const dlg = ui.dialog({ title: 'Spelling and Grammar: ' + langName(SP.paraList(issues[0].p)), width: 540, body: h('div', { class: 'row', style: 'flex-wrap:nowrap' }, left, side), buttons: [{ label: '&Options...', onClick: () => { L.dlg.options('spelling'); return false; } }, { label: '&Undo', onClick: () => { if (D.canUndo()) { D.undo(); idx = Math.max(-1, idx - 2); next(); } return false; } }, { label: 'Cancel', cancel: true }], onClose: () => SP.refreshSoon(0) });
    dlg.buttons[0].classList.add('left');
    next();
  }

  /* ================= readability statistics (shown after a full check when asked) ================= */
  SP.readability = function () {
    const d = doc();
    D.reindex(d);
    const text = d.main.paras.map((p) => D.plainText(p)).join('\n');
    const sentences = text.split(/[.!?]+(?:\s|$)/).filter((s) => /\p{L}/u.test(s));
    const ws = words(text).map((x) => x.w);
    const syl = (w) => { w = w.toLowerCase().replace(/[^a-z]/g, ''); if (w.length <= 3) return 1; w = w.replace(/(?:[^laeiouy]es|ed|[^laeiouy]e)$/, '').replace(/^y/, ''); const m = w.match(/[aeiouy]{1,2}/g); return m ? m.length : 1; };
    const sylls = ws.reduce((a, w) => a + syl(w), 0);
    const W = Math.max(1, ws.length), S = Math.max(1, sentences.length);
    const paras = d.main.paras.filter((p) => D.plainText(p).trim()).length;
    /* passive voice: a form of "to be" (optionally followed by an adverb) and a past participle */
    const PART = /^(?:\p{L}+ed|been|born|borne|beaten|begun|bent|bitten|blown|broken|brought|built|bought|caught|chosen|come|cut|dealt|done|drawn|driven|eaten|fallen|felt|fought|found|forgotten|forgiven|frozen|given|gone|grown|heard|held|hidden|hit|hurt|kept|known|laid|led|left|lent|let|lost|made|meant|met|paid|put|read|ridden|run|said|seen|sent|set|shaken|shown|shot|shut|sold|sought|spent|spoken|spread|stolen|struck|sung|sunk|taken|taught|thrown|told|thought|torn|understood|won|worn|woven|written)$/u;
    const BE = /^(?:am|is|are|was|were|be|been|being|'s|'re)$/i;
    const passive = sentences.filter((s0) => {
      const t = s0.toLowerCase().match(/[\p{L}']+/gu) || [];
      for (let i = 0; i < t.length - 1; i++) {
        if (!BE.test(t[i])) continue;
        const nx = /ly$/.test(t[i + 1]) && t[i + 2] ? t[i + 2] : t[i + 1];
        if (PART.test(nx) && !/^(?:need|red|bed|shed|seed|feed|bred|sled|wed|fled|speed|hundred|indeed|naked|wicked|sacred|kindred|ragged|rugged|beloved|aged|learned|tired|interested|excited|bored|pleased|concerned|used|supposed|married|based|related|located|involved)$/.test(nx)) return true;
      }
      return false;
    }).length;
    const ease = 206.835 - 1.015 * (W / S) - 84.6 * (sylls / W);
    const grade = 0.39 * (W / S) + 11.8 * (sylls / W) - 15.59;
    return { words: ws.length, chars: text.replace(/\s/g, '').length, paras, sentences: sentences.length, sPerPara: L.round(sentences.length / Math.max(1, paras), 1), wPerS: L.round(W / S, 1), cPerW: L.round(text.replace(/[^\p{L}]/gu, '').length / W, 1), passive: Math.round(100 * passive / S), ease: L.round(Math.max(0, Math.min(100, ease)), 1), grade: L.round(Math.max(0, grade), 1) };
  };
})();
