/* Ledger — spelling (Tools ▸ Spelling, F7) and AutoCorrect for cell entries.
 * The words come from the shared engine (common/spell.js); this file adds the dialog and AutoCorrect.
 */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, O = L.ops, F = L.formula;
  const { h } = L;
  const ui = L.ui;
  const SP = L.spell;   // the engine: common/spell.js
  const opt = (k, def) => { const o = L.app && L.app.opts; return o && o[k] !== undefined ? o[k] !== false : def !== false; };

  const LEX = L.lex;

  /* ================= AutoCorrect ================= */
  const AC = (L.autocorrect = {});
  const EXCEPTIONS = ['a.', 'abbr.', 'abs.', 'acad.', 'acct.', 'adj.', 'adv.', 'al.', 'apr.', 'approx.', 'assn.', 'aug.', 'ave.', 'b.a.', 'c.', 'ca.', 'cf.', 'co.', 'corp.', 'dec.', 'dept.', 'dr.', 'e.g.', 'ed.', 'esp.', 'est.', 'etc.', 'ex.', 'feb.', 'fig.', 'figs.', 'fri.', 'gen.', 'i.e.', 'inc.', 'jan.', 'jr.', 'jul.', 'jun.', 'ltd.', 'mar.', 'max.', 'min.', 'misc.', 'mon.', 'mr.', 'mrs.', 'ms.', 'mt.', 'no.', 'nov.', 'oct.', 'p.', 'pp.', 'prof.', 'ref.', 'rev.', 'sat.', 'sep.', 'sept.', 'sr.', 'st.', 'sun.', 'tel.', 'thu.', 'tue.', 'u.s.', 'vol.', 'vs.', 'wed.'];
  /* the replacement list: common/spell.js's, less what Excel's leaves out, with the user's entries */
  AC.list = () => L.spell.autocorrectList(['i', ':)', ':-)', ':(', ':-(', '==>', '<==', '<=>']);
  let LIST = null;
  AC.reload = () => { LIST = AC.list(); };
  AC.setEntry = (from, to) => { L.spell.setAutocorrectEntry(from, to); AC.reload(); };
  AC.exceptions = () => L.store.get('acExceptions', EXCEPTIONS);
  const DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
  /**
   * AutoCorrect a typed text entry (not formulas): replacement list, TWo INitial CApitals,
   * capitalised day names and first letters of sentences, accidental cAPS LOCK.
   */
  AC.applyCell = function (text) {
    if (!text || /^[=+\-@']/.test(text) || !opt('acReplace', true) && !opt('acTwoCaps', true)) return text;
    if (!LIST) AC.reload();
    let out = text;
    if (opt('acReplace', true)) {
      out = out.replace(/(^|[\s(])([^\s()]+?)(?=$|[\s,.;:!?)])/g, (m, pre, w) => {
        if (LIST.has(w)) return pre + LIST.get(w);
        const lw = w.toLowerCase();
        if (LIST.has(lw) && /^[a-z]/.test(LIST.get(lw))) { const r = LIST.get(lw); return pre + (w[0] === w[0].toUpperCase() && w[0] !== w[0].toLowerCase() ? r[0].toUpperCase() + r.slice(1) : r); }
        return m;
      });
      for (const k of ['(c)', '(r)', '(tm)', '...', '-->', '<--', '==>', '<==', '<=>']) if (out.indexOf(k) >= 0 && LIST.has(k)) out = out.split(k).join(LIST.get(k));
    }
    if (opt('acTwoCaps', true)) out = out.replace(/\b([A-Z])([A-Z])([a-z]{2,})\b/g, (m, a, b, rest) => a + b.toLowerCase() + rest);
    if (opt('acDays', true)) out = out.replace(/\b(monday|tuesday|wednesday|thursday|friday|saturday|sunday)\b/g, (m) => m[0].toUpperCase() + m.slice(1));
    if (opt('acCapsLock', true)) out = out.replace(/\b([a-z])([A-Z]{2,})\b/g, (m, a, rest) => a.toUpperCase() + rest.toLowerCase());
    if (opt('acSentence', false)) out = out.replace(/(^|[.!?]\s+)([a-z])/g, (m, pre, ch) => pre + ch.toUpperCase());
    void DAYS;
    return out;
  };

  /* ================= Spelling dialog (cells and comments) ================= */
  const WORD = /[\p{L}][\p{L}\p{M}'’\-]*[\p{L}\p{M}]|[\p{L}]/gu;
  /** misspellings in a text: [{word, index}] */
  function findErrors(text, list) {
    const out = [];
    WORD.lastIndex = 0;
    let m;
    while ((m = WORD.exec(text))) {
      const w = m[0];
      if (/^(https?|www|mailto)/i.test(w) || /\.\w+\//.test(text.slice(m.index - 1, m.index + w.length + 2))) continue;
      if (opt('ignoreInternet', true) && /[@\/\\]/.test(text.slice(Math.max(0, m.index - 1), m.index + w.length + 1))) continue;
      if (!SP.okWord(w, list)) out.push({ word: w, index: m.index });
    }
    return out;
  }
  SP.findErrors = findErrors;
  /** Tools ▸ Spelling (F7): the common Spelling dialog (common/dialogs.js) over the sheet's text cells */
  SP.check = async function () {
    const G = L.grid, A = L.app;
    const sh = G.sheet(), wb = G.wb;
    let list = (A.opts && A.opts.dictLang) === 'en_GB' ? 'en_GB' : 'en_US';
    ui.busy(true, 'Loading the dictionary...');
    try { await LEX.load(list); } catch (e) { ui.busy(false); ui.msg(L.t('The spelling dictionary could not be loaded ({0}).', e.message || e), { icon: 'warn' }); return; }
    ui.busy(false);
    /* cells to check: the selection when it is more than one cell, otherwise the sheet from the active cell, wrapping */
    const rg0 = G.range();
    const single = G.ranges().length === 1 && rg0.r1 === rg0.r2 && rg0.c1 === rg0.c2;
    const s0 = G.sel();
    const cells = [];
    sh.rows.forEach((row, r) => { if (!row) return; row.cells.forEach((cell, c) => { if (!cell || typeof cell.v !== 'string' || cell.f != null || cell.am) return; if (!single && !G.ranges().some((x) => M.rangeContains(x, r, c))) return; cells.push({ r, c }); }); });
    if (single) { const k = cells.findIndex((x) => x.r > s0.r || (x.r === s0.r && x.c >= s0.c)); if (k > 0) cells.push(...cells.splice(0, k)); }
    let ci = 0, hits = [], hi = -1, cur = null;
    const changeAll = new Map();
    const scan = (it) => findErrors(sh.val(it.r, it.c), list).map((x) => Object.assign(x, { r: it.r, c: it.c }));
    /** the word at its place in the cell, else its first appearance there */
    const replaceIn = (x, to) => {
      const v = sh.val(x.r, x.c);
      if (typeof v !== 'string') return;
      let nv;
      if (v.substr(x.index, x.word.length) === x.word) nv = v.slice(0, x.index) + to + v.slice(x.index + x.word.length);
      else { const re = new RegExp('(^|[^\\p{L}])' + x.word.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?=$|[^\\p{L}])', 'u'); nv = v.replace(re, (m, pre) => pre + to); }
      if (nv !== v) O.tx(wb, L.N('Spelling'), () => O.setValue(sh, x.r, x.c, nv));
    };
    const advance = () => {
      for (;;) {
        if (++hi < hits.length) { cur = hits[hi]; return true; }
        if (ci >= cells.length) return false;
        const it = cells[ci++];
        if (typeof sh.val(it.r, it.c) !== 'string') { hits = []; hi = -1; continue; }
        /* Change All replacements apply silently */
        hits = scan(it).filter((x) => { if (changeAll.has(x.word)) { replaceIn(x, changeAll.get(x.word)); return false; } return true; });
        hi = -1;
      }
    };
    /* after a change, the rest of the cell is checked again */
    const rescan = (from) => { hits = scan(cur).filter((x) => x.index >= from); hi = -1; };
    const LANGS = [['en_US', 'English (U.S.)'], ['en_GB', 'English (U.K.)']];
    const langName = (k) => (LANGS.find((x) => x[0] === k) || LANGS[0])[1];
    if (!advance()) { ui.msg('The spelling check is complete for the entire sheet.', { icon: 'info' }); return; }
    hi--;   // the first finding again, for the dialog's first next()
    ui.spellingDialog({
      title: (lang) => L.t('Spelling: {0}', L.t(lang)), language: () => langName(list), languages: LANGS,
      end: 'The spelling check is complete for the entire sheet.',
      next: () => {
        if (!advance()) { G.paint(); return null; }
        G.select(cur.r, cur.c);
        const v = sh.val(cur.r, cur.c);
        return { kind: 'spelling', before: v.slice(Math.max(0, cur.index - 40), cur.index), word: cur.word, after: v.slice(cur.index + cur.word.length, cur.index + cur.word.length + 40), suggestions: SP.suggest(cur.word, list, 8) };
      },
      ignoreOnce: () => {},
      ignoreAll: () => { SP.ignoreAll(cur.word); hits = hits.filter((x, k) => k <= hi || x.word.toLowerCase() !== cur.word.toLowerCase()); },
      add: () => SP.addWord(cur.word),
      change: (to) => { if (!to) return false; replaceIn(cur, to); rescan(cur.index + to.length); },
      changeAll: (to) => { if (!to) return false; changeAll.set(cur.word, to); replaceIn(cur, to); rescan(cur.index + to.length); },
      autocorrect: (to) => { L.autocorrect.setEntry(cur.word, to); replaceIn(cur, to); rescan(cur.index + to.length); },
      /* another dictionary: the current cell is checked again with it */
      setLanguage: async (v) => {
        ui.busy(true, 'Loading the dictionary...');
        try { await LEX.load(v); } catch (e) { ui.busy(false); ui.msg(L.t('The spelling dictionary could not be loaded ({0}).', e.message || e), { icon: 'warn' }); return false; }
        ui.busy(false);
        list = v;
        if (A.opts) { A.opts.dictLang = v; if (A.saveOpts) A.saveOpts(); }
        rescan(0);
      },
      options: () => L.dlg.options('Spelling'),
      rect: () => { const q = G.cellRect(cur.r, cur.c), w = L.$('#gridwrap').getBoundingClientRect(); return { left: w.left + q.x, top: w.top + q.y, right: w.left + q.x + q.w, bottom: w.top + q.y + q.h }; },
      onClose: () => G.paint(),
    });
  };
})(typeof window !== 'undefined' ? window : globalThis);
