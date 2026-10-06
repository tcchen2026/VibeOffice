/* VibeOffice — the spelling engine shared by Quire and Ledger: the offline word lists (common/dict/,
 * Hunspell/SCOWL-derived, ranked by frequency), checking, suggestions by edit distance, the custom dictionary
 * and Ignore All. Each app has its own spelling UI on top (Quire: underlines, grammar and its dialog;
 * Ledger: Tools ▸ Spelling); an app that draws underlines sets SP.refreshSoon to redraw them.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const SP = (L.spell = L.spell || {});
  const opt = (k, def) => { const o = L.app && L.app.opts; return o && o[k] !== undefined ? o[k] !== false : def !== false; };
  SP.refreshSoon = SP.refreshSoon || (() => {});

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
  const ignoredAll = new Set();
  const cache = new Map();
  /** after the custom dictionary changed elsewhere (Tools ▸ Options): forget the cached answers */
  SP.reloadCustom = () => { SP._custom = null; cache.clear(); SP.refreshSoon(); };
  /** Add to Dictionary: the word is spelled correctly from now on, in every app (localStorage per app) */
  SP.addWord = function (w) {
    const list = L.store.get('customDict', []) || [];
    if (!list.some((x) => String(x).toLowerCase() === String(w).toLowerCase())) list.push(String(w));
    L.store.set('customDict', list);
    SP.reloadCustom();
  };
  /** Ignore All: accept the word for the rest of the session */
  SP.ignoreAll = (w) => { ignoredAll.add(String(w).toLowerCase()); cache.clear(); };
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
        const score = 11 + (lex.lower.get(a).lv + lex.lower.get(b).lv) * 1.2;
        const s = (a === 'i' ? 'I' : lex.lower.get(a).w) + ' ' + lex.lower.get(b).w;
        if (!found.has(s) || found.get(s) > score) found.set(s, score);
      }
    }
    return Array.from(found.entries()).sort((x, y) => x[1] - y[1]).slice(0, max).map(([s]) => (/^\p{Lu}/u.test(s) && !isUpper(w) ? s : restoreCase(s, w)));
  };
})(typeof window !== 'undefined' ? window : globalThis);
