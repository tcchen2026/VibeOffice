/* Quire — AutoCorrect, AutoFormat As You Type, AutoComplete and AutoText. */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout;
  const AC = (L.autocorrect = {});
  const doc = () => D.doc;
  const opt = (k) => { const a = L.app && L.app.opts; return a ? a[k] !== false : true; };

  const EXCEPTIONS = ['a.', 'abbr.', 'abs.', 'acad.', 'acct.', 'adj.', 'adv.', 'al.', 'apr.', 'approx.', 'assn.', 'aug.', 'ave.', 'b.a.', 'c.', 'ca.', 'cf.', 'co.', 'corp.', 'dec.', 'dept.', 'dr.', 'e.g.', 'ed.', 'esp.', 'est.', 'etc.', 'ex.', 'feb.', 'fig.', 'figs.', 'fri.', 'gen.', 'i.e.', 'inc.', 'jan.', 'jr.', 'jul.', 'jun.', 'ltd.', 'mar.', 'max.', 'min.', 'misc.', 'mon.', 'mr.', 'mrs.', 'ms.', 'mt.', 'no.', 'nov.', 'oct.', 'p.', 'pp.', 'prof.', 'ref.', 'rev.', 'sat.', 'sep.', 'sept.', 'sr.', 'st.', 'sun.', 'tel.', 'thu.', 'tue.', 'u.s.', 'vol.', 'vs.', 'wed.'];
  /* the replacement list: common/spell.js's, with the user's entries */
  AC.list = () => L.spell.autocorrectList();
  let LIST = null;
  AC.reload = () => { LIST = AC.list(); };
  AC.setEntry = (from, to) => { L.spell.setAutocorrectEntry(from, to); AC.reload(); };
  AC.exceptions = () => L.store.get('acExceptions', EXCEPTIONS);
  AC.setExceptions = (list) => { L.store.set('acExceptions', list); };
  const DAYS = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday'];
  const MONTHS = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];

  /* ---------- smart quotes ---------- */
  AC.smartChar = function (ch) {
    if ((ch !== '"' && ch !== "'") || !opt('smartQuotes') || !E.sel) return ch;
    const pos = E.collapsed() ? E.sel.f : E.range()[0];
    const prev = pos.o > 0 ? D.ptext(pos.p, pos.o - 1, pos.o) : '';
    const opening = !prev || /[\s([{—–\-\/“‘"'«„‚]/.test(prev) && !(ch === "'" && /[\p{L}\p{N}]/u.test(prev));
    /* the quotes of the paragraph's language: « » in French, „ “ in German (common/locales.js) */
    const q = L.locale && L.spell && L.spell.paraLang ? L.locale.quotes(L.spell.paraLang(pos.p)) : ['“', '”', '‘', '’'];
    if (ch === '"') return opening ? q[0] : q[1];
    return opening ? q[2] : q[3];
  };

  /* ---------- after typing a character ---------- */
  let busy = false;
  const TERM = /[\s.,;:!?)\]}"'”’]/;
  AC.afterType = function (ch, virtual) {
    if (busy || !E.sel || !E.collapsed()) return;
    const pos = E.sel.f;
    const p = pos.p;
    if (/[\p{L}\p{N}]/u.test(ch)) { tipCheck(p, pos.o); return; }
    hideTip();
    if (!TERM.test(ch) && ch !== '-' && ch !== '/' && ch !== '*' && ch !== '_') return;
    const text = D.ptext(p, 0, pos.o);
    const end = pos.o - (virtual ? 0 : ch.length); /* offset where the terminator starts */
    busy = true;
    try {
      /* replacements that include their own punctuation, e.g. "(c)" or "..." */
      if (opt('autocorrect')) {
        if (!LIST) AC.reload();
        if (ch === ')' || ch === '.' || ch === '>' || ch === '-') {
          for (const k of ['(c)', '(r)', '(tm)', '...', ':)', ':-)', ':(', ':-(', '-->', '<--', '==>', '<==', '<=>']) {
            if (text.endsWith(k) && LIST.has(k) && (k !== '...' || ch === '.')) { replace(p, pos.o - k.length, pos.o, LIST.get(k), 'AutoCorrect'); return; }
          }
        }
      }
      if (!/[\s.,;:!?)\]}"'”’]/.test(ch)) { if (ch === '*' || ch === '_') emphasis(p, pos.o, ch); return; }
      const m = /(\S+)$/.exec(D.ptext(p, 0, end));
      if (!m) { dashFix(p, end); return; }
      let word = m[1];
      const ws = end - word.length;
      /* trailing punctuation inside the token stays outside the word */
      const lead = /^[("“‘'\[]+/.exec(word);
      const w0 = ws + (lead ? lead[0].length : 0);
      word = word.slice(lead ? lead[0].length : 0);
      if (!word) return;
      /* list entries are tried with their exact case, then lowercase with the case pattern restored */
      if (opt('autocorrect')) {
        if (!LIST) AC.reload();
        let rep = LIST.get(word);
        if (rep == null && LIST.has(word.toLowerCase())) { rep = LIST.get(word.toLowerCase()); if (/^\p{Lu}/u.test(word)) rep = rep.charAt(0).toUpperCase() + rep.slice(1); if (word === word.toUpperCase() && word.length > 1) rep = rep.toUpperCase(); }
        /* multi-word entries ("could of been") */
        if (rep == null) { for (const [k, v] of LIST) if (k.includes(' ') && D.ptext(p, 0, end).toLowerCase().endsWith(k)) { replace(p, end - k.length, end, v, 'AutoCorrect'); return; } }
        if (rep != null && rep !== word) { replace(p, w0, end, rep, 'AutoCorrect'); word = rep; }
      }
      const wordNow = D.ptext(p, w0, end);
      /* TWo INitial CApitals */
      if (opt('twoCaps') && /^\p{Lu}{2}\p{Ll}{2,}$/u.test(wordNow) && !/^(IDs|PCs|CDs|URLs|TVs)$/.test(wordNow)) { replace(p, w0, end, wordNow.charAt(0) + wordNow.slice(1).toLowerCase(), 'AutoCorrect'); }
      /* days of the week */
      if (opt('capDays') && DAYS.includes(wordNow.toLowerCase()) && /^\p{Ll}/u.test(wordNow)) replace(p, w0, w0 + 1, wordNow.charAt(0).toUpperCase(), 'AutoCorrect');
      /* first letter of sentences */
      if (opt('capSentence') && /^\p{Ll}/u.test(D.ptext(p, w0, w0 + 1))) {
        const before = D.ptext(p, 0, w0);
        const prevWord = (/(\S+)\s+$/.exec(before) || [])[1] || '';
        const startOfPara = !before.trim() && !(LY.labels && false);
        const afterEnd = /[.!?]["'”’)\]]*\s+$/.test(before) && !AC.exceptions().includes(prevWord.toLowerCase()) && !/^\S\.$/.test(prevWord) && !/\.\S+\.$/.test(prevWord);
        const inCell = D.cellOf(doc(), p);
        if ((startOfPara && (!inCell || opt('capCells'))) || afterEnd) {
          const wNow = D.ptext(p, w0, end);
          if (!/^(e\.g|i\.e|www|http|iPod|iPhone|eBay)/i.test(wNow) && !/[@\/\\]/.test(wNow)) replace(p, w0, w0 + 1, wNow.charAt(0).toUpperCase(), 'AutoCorrect');
        }
      }
      /* "i" alone */
      /* ordinals 1st → 1ˢᵗ */
      if (opt('ordinals')) { const om = /^(\d+)(st|nd|rd|th)$/.exec(D.ptext(p, w0, end)); if (om && D.fmtNum(+om[1], 'ordinal').endsWith(om[2])) superscript(p, w0 + om[1].length, end); }
      /* fractions */
      if (opt('fractions')) { const fr = { '1/2': '½', '1/4': '¼', '3/4': '¾' }[D.ptext(p, w0, end)]; if (fr) replace(p, w0, end, fr, 'AutoFormat'); }
      /* internet and network paths → hyperlinks */
      if (opt('autoUrl')) {
        const tok = D.ptext(p, w0, end).replace(/[.,;:!?]+$/, '');
        if (/^(https?:\/\/|ftp:\/\/|www\.)[^\s]+\.[^\s]+/i.test(tok) || /^[\w.+-]+@[\w-]+\.[\w.-]+$/.test(tok) || /^\\\\\w+/.test(tok)) {
          const it = D.itemAfter(p, w0);
          if (!(it && it.rPr && it.rPr.link)) {
            const url = /@/.test(tok) && !/^https?:/.test(tok) ? 'mailto:' + tok : /^www\./i.test(tok) ? 'http://' + tok : tok;
            link(p, w0, w0 + tok.length, url);
          }
        }
      }
      dashFix(p, w0);
      if (ch === '*' || ch === '_') emphasis(p, pos.o, ch);
    } finally { busy = false; }
  };
  /** "word--word" → em dash; "word - word" → en dash (applied when the following word is finished) */
  function dashFix(p, w0) {
    if (!opt('dashes')) return;
    const t = D.ptext(p, 0, w0 + 1);
    let m = /([\p{L}\p{N}])--$/u.exec(D.ptext(p, 0, w0));
    if (m) { replace(p, w0 - 2, w0, '—', 'AutoFormat'); return; }
    m = /\S -{1,2} $/.exec(D.ptext(p, 0, w0));
    if (m) { const k = m[0].length - 3; void k; const s = w0 - (m[0].endsWith('-- ') ? 3 : 2); replace(p, s, w0 - 1, '–', 'AutoFormat'); }
    void t;
  }
  function emphasis(p, o, ch) {
    if (!opt('autoFormat')) return;
    const t = D.ptext(p, 0, o);
    const re = ch === '*' ? /(^|[\s(])\*([^\s*][^*]*?[^\s*]|[^\s*])\*$/ : /(^|[\s(])_([^\s_][^_]*?[^\s_]|[^\s_])_$/;
    const m = re.exec(t);
    if (!m) return;
    const s = o - m[2].length - 2;
    const prop = ch === '*' ? { b: true } : { i: true };
    E.edit('AutoFormat', () => {
      O.deleteRange(D.pos(p, o - 1), D.pos(p, o));
      O.deleteRange(D.pos(p, s), D.pos(p, s + 1));
      O.setRunProps(D.pos(p, s), D.pos(p, s + m[2].length), prop);
      return D.pos(p, s + m[2].length);
    });
    E.pending = null;
  }
  function replace(p, a, b, text, label) {
    const keepTail = E.sel.f.o - b;
    E.edit(label || 'AutoCorrect', () => {
      const r = O.rPrAt(D.pos(p, a + 1 > b ? a : a + 1));
      void r;
      const it = D.itemAfter(p, a);
      const rp = it && it.rPr ? L.clone(it.rPr) : {};
      O.deleteRange(D.pos(p, a), D.pos(p, b));
      const end = O.insertText(D.pos(p, a), text, rp);
      return D.pos(p, end.o + keepTail);
    });
    D.breakMerge();
  }
  function superscript(p, a, b) {
    const keep = E.sel.f;
    E.edit('AutoFormat', () => { O.setRunProps(D.pos(p, a), D.pos(p, b), { vert: 'superscript' }); return keep; });
    E.pending = { vert: undefined };
  }
  function link(p, a, b, url) {
    const keep = E.sel.f;
    const d = doc();
    E.edit('AutoFormat', () => {
      const style = D.useStyle(d, 'Hyperlink');
      O.formatRange(D.pos(p, a), D.pos(p, b), (r) => Object.assign(r, { link: { url }, style }));
      return keep;
    });
    E.pending = { link: undefined, style: undefined };
  }

  /* ---------- after Enter: AutoFormat on the paragraph just finished ---------- */
  AC.afterEnter = function () {
    hideTip();
    if (!E.sel || busy) return;
    const d = doc();
    const cur = E.sel.f.p;
    const prev = D.prevPara(d, cur);
    if (!prev || O.containerOf(prev) !== O.containerOf(cur)) return;
    const t = D.ptext(prev);
    busy = true;
    try {
      /* finish the last word of the paragraph (AutoCorrect on Enter) */
      if (t && opt('autocorrect')) {
        const m = /(\S+)$/.exec(t);
        if (m) {
          if (!LIST) AC.reload();
          const w = m[1].replace(/[.,;:!?]+$/, '');
          let rep = LIST.get(w);
          if (rep == null && LIST.has(w.toLowerCase())) { rep = LIST.get(w.toLowerCase()); if (/^\p{Lu}/u.test(w)) rep = rep.charAt(0).toUpperCase() + rep.slice(1); }
          const s = t.length - m[1].length;
          if (rep != null && rep !== w) { const keep = E.sel; E.edit('AutoCorrect', () => { const it = D.itemAfter(prev, s); O.deleteRange(D.pos(prev, s), D.pos(prev, s + w.length)); O.insertText(D.pos(prev, s), rep, it && it.rPr ? L.clone(it.rPr) : {}); return keep; }); }
          if (opt('autoUrl') && /^(https?:\/\/|www\.)\S+\.\S+/i.test(w)) link(prev, s, s + w.length, /^www\./i.test(w) ? 'http://' + w : w);
        }
      }
      /* border lines */
      if (opt('autoBorders')) {
        const bm = /^(-{3,}|_{3,}|={3,}|\*{3,}|~{3,}|#{3,})$/.exec(t.trim());
        if (bm && !prev.runs.some((it) => it.t !== 'text')) {
          const kind = bm[1][0];
          const border = { '-': { val: 'single', sz: 0.75 }, _: { val: 'single', sz: 1.5 }, '=': { val: 'double', sz: 0.75 }, '*': { val: 'dotted', sz: 2.25 }, '~': { val: 'wave', sz: 0.75 }, '#': { val: 'thinThickSmallGap', sz: 2.25 } }[kind];
          E.edit('AutoFormat', () => { D.touch(prev); prev.runs = []; prev.pPr.borders = { bottom: Object.assign({ color: 'auto', space: 1 }, border) }; return E.sel; });
          return;
        }
      }
      /* "+----+-----+" → table */
      if (opt('autoTables') && /^\+(-+\+)+$/.test(t.trim())) {
        const segs = t.trim().slice(1, -1).split('+').map((s) => s.length);
        const total = segs.reduce((a, b) => a + b, 0);
        const width = L.tables.availWidth(D.pos(prev, 0));
        E.edit('AutoFormat', () => {
          const tbl = D.simpleTable(1, segs.length, width);
          tbl.tblPr.style = D.useStyle(doc(), tbl.tblPr.style);
          tbl.grid = segs.map((n) => L.round((n / total) * width, 2));
          L.tables.fixWidths(tbl);
          const cont = D.touchList(D.info(d, prev).cont);
          cont.blocks.splice(cont.blocks.indexOf(prev), 1, tbl);
          d._idxDirty = true;
          return D.pos(tbl.rows[0].cells[0].blocks[0], 0);
        });
        return;
      }
      /* automatic bulleted and numbered lists */
      if (!(L.lists && L.lists.info(prev))) {
        const nm = /^((\d{1,3})|([a-zA-Z])|([ivxIVX]{1,5}))([.)\->]|:)( +|\t)\S/.exec(t);
        const bm = /^([*\-•>o]|-->|=>|->|>)( +|\t)\S/.exec(t);
        if (nm && opt('autoNumbers')) {
          const lbl = nm[1];
          const fmt = nm[2] ? 'decimal' : nm[4] && /^[ivx]+$/.test(lbl) && (lbl === 'i' || lbl.length > 1) ? 'lowerRoman' : nm[4] && /^[IVX]+$/.test(lbl) && (lbl === 'I' || lbl.length > 1) ? 'upperRoman' : /[a-z]/.test(lbl) ? 'lowerLetter' : 'upperLetter';
          const start = nm[2] ? +nm[2] : fmt === 'lowerRoman' || fmt === 'upperRoman' ? 1 : lbl.toLowerCase().charCodeAt(0) - 96;
          if (start === 1 || nm[2] && start < 100) makeList(prev, cur, nm[0].length - 1, D.makeNumberAbs(fmt, '%1' + nm[5]), start);
          return;
        }
        if (bm && opt('autoBullets')) {
          const sym = bm[1];
          const ch = sym === '*' ? ['•', 'Symbol', ''] : sym === '-' ? ['–', 'Arial', '–'] : sym === 'o' ? ['o', 'Courier New', 'o'] : sym === '>' ? ['➢', 'Wingdings', ''] : sym.length > 1 ? ['➔', 'Wingdings', ''] : ['•', 'Symbol', ''];
          makeList(prev, cur, bm[0].length - 1, D.makeBulletAbs(ch[0], ch[1], ch[2]), 1);
          return;
        }
      }
      /* headings: a short line without ending punctuation followed by two Enters */
      if (opt('autoHeadings') && false) return;
    } finally { busy = false; }
  };
  function makeList(prev, cur, cut, abs, start) {
    const d = doc();
    E.edit('AutoFormat', () => {
      D.touchKey(d, 'numbering');
      const nid = D.addNum(d, abs);
      if (start > 1) d.numbering.nums[nid].ov = { 0: { start } };
      O.deleteRange(D.pos(prev, 0), D.pos(prev, cut));
      for (const p of [prev, cur]) { D.touch(p); p.pPr.num = { id: nid, lvl: 0 }; delete p.pPr.ind; }
      return D.pos(cur, 0);
    });
  }

  /* ---------- AutoComplete tips ---------- */
  let tip = null, tipState = null;
  const ATEXT = () => AC.autoTextEntries();
  function tipCheck(p, o) {
    if (!opt('autoComplete')) return hideTip();
    const before = D.ptext(p, 0, o);
    const m = /([\p{L}][\p{L}' ]*)$/u.exec(before);
    if (!m) return hideTip();
    const frag = m[1];
    const word = /(\p{L}+)$/u.exec(frag)[1];
    if (word.length < 4) return hideTip();
    const now = new Date();
    const cands = MONTHS.map((x) => x[0].toUpperCase() + x.slice(1)).concat(DAYS.map((x) => x[0].toUpperCase() + x.slice(1)), [`${MONTHS[now.getMonth()][0].toUpperCase() + MONTHS[now.getMonth()].slice(1)} ${now.getDate()}, ${now.getFullYear()}`]);
    for (const e of ATEXT()) if (e.text.length >= 5 && !/\n/.test(e.text)) cands.push(e.text);
    const hit = cands.find((c) => c.length > word.length && c.toLowerCase().startsWith(word.toLowerCase()) && c.toLowerCase() !== word.toLowerCase());
    if (!hit) return hideTip();
    showTip(hit, p, o, word.length);
  }
  function showTip(text, p, o, typed) {
    if (!tip) { tip = L.h('div', { class: 'tooltip ac-tip' }); document.body.appendChild(tip); }
    tip.textContent = L.t('{0} (Press ENTER to Insert)', text);
    tip.hidden = false;
    const r = E.caretRect();
    if (r) { tip.style.left = Math.max(4, r.left - 10) + 'px'; tip.style.top = Math.max(4, r.top - 24) + 'px'; }
    tipState = { text, p, o, typed };
  }
  function hideTip() { if (tip) tip.hidden = true; tipState = null; }
  AC.hideTip = hideTip;
  AC.acceptTip = function () {
    if (!tipState || !E.sel || !E.collapsed()) return false;
    const st = tipState;
    if (E.sel.f.p !== st.p || E.sel.f.o !== st.o) { hideTip(); return false; }
    hideTip();
    const p = st.p;
    E.edit(L.N('AutoComplete'), () => {
      const a = st.o - st.typed;
      const it = D.itemBefore(p, st.o);
      O.deleteRange(D.pos(p, a), D.pos(p, st.o));
      return O.insertText(D.pos(p, a), st.text, it && it.rPr ? L.clone(it.rPr) : {});
    });
    return true;
  };
  L.bus.on('sel', () => { if (tipState && E.sel && (E.sel.f.p !== tipState.p || E.sel.f.o !== tipState.o)) hideTip(); });

  /* ---------- AutoText ---------- */
  const userName = () => L.store.get('userName', 'Quire User');
  const userInitials = () => L.store.get('userInitials', '') || O.initials();
  AC.AUTOTEXT = {
    'Attention Line': ['ATTENTION:', 'Attention:'],
    Closing: ['Best regards,', 'Best wishes,', 'Cordially,', 'Love,', 'Regards,', 'Respectfully yours,', 'Respectfully,', 'Sincerely yours,', 'Sincerely,', 'Take care,', 'Thank you,', 'Thanks,', 'Yours truly,'],
    'Header/Footer': [{ name: L.N('- PAGE -'), fields: ['- ', { f: 'PAGE' }, ' -'] }, { name: L.N('Author, Page #, Date'), fields: [{ f: 'AUTHOR' }, ', Page ', { f: 'PAGE' }, ', ', { f: 'DATE \\@ "M/d/yyyy"' }] }, { name: L.N('Confidential, Page #, Date'), fields: ['Confidential, Page ', { f: 'PAGE' }, ', ', { f: 'DATE \\@ "M/d/yyyy"' }] }, { name: L.N('Created by'), fields: ['Created by ', { f: 'AUTHOR' }] }, { name: L.N('Created on'), fields: ['Created on ', { f: 'CREATEDATE \\@ "M/d/yyyy h:mm:ss am/pm"' }] }, { name: L.N('Filename'), fields: [{ f: 'FILENAME' }] }, { name: L.N('Filename and path'), fields: [{ f: 'FILENAME \\p' }] }, { name: L.N('Last printed'), fields: ['Last printed ', { f: 'PRINTDATE \\@ "M/d/yyyy h:mm:ss am/pm"' }] }, { name: L.N('Last saved by'), fields: ['Last saved by ', { f: 'LASTSAVEDBY' }] }, { name: L.N('Page X of Y'), fields: ['Page ', { f: 'PAGE' }, ' of ', { f: 'NUMPAGES' }] }],
    'Mailing Instructions': ['CERTIFIED MAIL', 'CONFIDENTIAL', 'PERSONAL', 'REGISTERED MAIL', 'SPECIAL DELIVERY', 'VIA AIRMAIL', 'VIA FACSIMILE', 'VIA OVERNIGHT MAIL'],
    'Reference Initials': [() => userInitials() + '/' + userInitials().toLowerCase()],
    'Reference Line': ['In reply to:', 'RE:', 'Reference:'],
    Salutation: ['Dear Mother and Father,', 'Dear Sir or Madam:', 'Ladies and Gentlemen:', 'To Whom It May Concern:'],
    Signature: [() => userName()],
    'Subject Line': ['Subject:'],
  };
  /* the categories as the AutoText menu shows them */
  const CAT_LABEL = { 'Attention Line': L.N('&Attention Line'), Closing: L.N('&Closing'), 'Header/Footer': L.N('&Header/Footer'), 'Mailing Instructions': L.N('&Mailing Instructions'), 'Reference Initials': L.N('&Reference Initials'),
    'Reference Line': L.N('&Reference Line'), Salutation: L.N('&Salutation'), Signature: L.N('&Signature'), 'Subject Line': L.N('&Subject Line') };
  AC.autoTextEntries = function () {
    const out = [];
    for (const cat in AC.AUTOTEXT) for (const e of AC.AUTOTEXT[cat]) { if (typeof e === 'string') out.push({ cat, text: e, name: e }); else if (typeof e === 'function') out.push({ cat, text: e(), name: e() }); }
    for (const u of L.store.get('autoText', [])) out.push({ cat: 'Normal', text: u.text, name: u.name, user: true });
    return out;
  };
  function insertEntry(e) {
    if (typeof e === 'string' || typeof e === 'function') { E.insertTextAt(typeof e === 'function' ? e() : e, L.N('AutoText')); return; }
    if (e.text) { E.insertTextAt(e.text, L.N('AutoText')); return; }
    E.edit(L.N('AutoText'), () => {
      let pos = E.deleteSelection();
      for (const part of e.fields) {
        if (typeof part === 'string') pos = O.insertText(pos, part, O.inheritRPr(pos));
        else { const res = L.fields ? L.fields.evaluate(part.f, pos) : ''; pos = O.insertField(pos, part.f, res == null ? '' : String(res), O.inheritRPr(pos)); }
      }
      return pos;
    });
    LY.refreshDynamic();
  }
  AC.insertEntry = insertEntry;
  AC.autoTextMenu = function () {
    const items = [{ label: 'Auto&Text...', run: () => L.dlg && L.dlg.autocorrect && L.dlg.autocorrect('autotext') }, '-'];
    for (const cat in AC.AUTOTEXT) items.push({ label: CAT_LABEL[cat] || cat.replace(/^(\w)/, '&$1'), sub: AC.AUTOTEXT[cat].map((e) => ({ label: (typeof e === 'string' ? e : typeof e === 'function' ? e() : e.name).replace(/&/g, '&&'), run: () => insertEntry(e) })) });
    const user = L.store.get('autoText', []);
    if (user.length) items.push({ label: '&Normal', sub: user.map((u) => ({ label: u.name, run: () => insertEntry({ text: u.text }) })) });
    return items;
  };
  AC.hfAutoText = () => AC.AUTOTEXT['Header/Footer'].map((e) => ({ label: e.name, run: () => insertEntry(e) }));
  AC.addAutoText = function (name, text) { const list = L.store.get('autoText', []).filter((x) => x.name !== name); list.push({ name, text }); L.store.set('autoText', list); };
  AC.deleteAutoText = function (name) { L.store.set('autoText', L.store.get('autoText', []).filter((x) => x.name !== name)); };
  /* F3 inserts the AutoText entry whose name precedes the caret */
  AC.expandF3 = function () {
    if (!E.sel || !E.collapsed()) return false;
    const p = E.sel.f.p, o = E.sel.f.o;
    const m = /(\S+)$/.exec(D.ptext(p, 0, o));
    if (!m) return false;
    const e = AC.autoTextEntries().find((x) => x.name.toLowerCase() === m[1].toLowerCase() || x.name.toLowerCase().startsWith(m[1].toLowerCase()) && m[1].length >= 4);
    if (!e) return false;
    E.edit(L.N('AutoText'), () => { O.deleteRange(D.pos(p, o - m[1].length), D.pos(p, o)); return O.insertText(D.pos(p, o - m[1].length), e.text, O.inheritRPr(D.pos(p, o - m[1].length))); });
    return true;
  };
})();
