/* Quire — document model.
 * A document is a tree of containers ({blocks:[...]}) holding paragraphs and tables.
 * Lengths are in points, font sizes in points, colours as 'RRGGBB' or 'auto'.
 *
 *   doc.main            body story            {kind:'main', blocks}
 *   doc.sect            final section properties (earlier sections end at paragraphs carrying p.sect)
 *   doc.hf[id]          header/footer stories {kind:'hdr'|'ftr', id, blocks}
 *   doc.fn[id]/doc.en[id] footnotes / endnotes {kind:'fn'|'en', id, blocks}
 *   doc.comments[id]    {kind:'cmt', id, author, initials, date, blocks, done}
 *   paragraph           {t:'p', id, pPr, rPr (paragraph mark), runs:[item], sect?}
 *   table               {t:'tbl', id, tblPr, grid:[pt], rows:[{id, trPr, cells:[{id, tcPr, blocks}]}]}
 *   items               text | tab | br | img | shape | fn | en | sym | cs/ce (comment range) | bs/be (bookmark)
 *                       | fb/fs/fe (field begin/separate/end) | raw (preserved XML) | ptab | sep
 */
(function () {
  'use strict';
  const L = window.L;
  const D = (L.D = {});
  let NID = 1;
  D.nid = () => NID++;

  /* ================= constructors ================= */
  D.para = (runs, pPr, rPr) => ({ t: 'p', id: D.nid(), pPr: pPr || {}, rPr: rPr || {}, runs: runs || [] });
  D.text = (text, rPr) => ({ t: 'text', text: String(text), rPr: rPr || {} });
  D.item = (t, extra, rPr) => Object.assign({ t, rPr: rPr || {} }, extra || {});
  D.cell = (blocks, tcPr) => ({ id: D.nid(), tcPr: tcPr || {}, blocks: blocks && blocks.length ? blocks : [D.para()] });
  D.row = (cells, trPr) => ({ id: D.nid(), trPr: trPr || {}, cells });
  D.table = (rows, grid, tblPr) => ({ t: 'tbl', id: D.nid(), tblPr: tblPr || {}, grid: grid || [], rows: rows || [] });
  D.simpleTable = function (nr, nc, width, opts) {
    opts = opts || {};
    const cw = width / nc;
    const grid = Array.from({ length: nc }, () => cw);
    const rows = [];
    for (let r = 0; r < nr; r++) rows.push(D.row(Array.from({ length: nc }, () => D.cell([D.para([], opts.pPr ? L.clone(opts.pPr) : {})], { w: cw }))));
    const tblPr = Object.assign({ style: 'TableGrid', w: { type: 'auto', v: 0 }, look: { firstRow: true, firstCol: true, noVBand: true } }, opts.tblPr || {});
    return D.table(rows, grid, tblPr);
  };
  /** item length in character positions */
  D.ilen = (it) => {
    if (it.zero) return 0;
    switch (it.t) {
      case 'text': return it.text.length;
      case 'cs': case 'ce': case 'bs': case 'be': case 'fb': case 'fs': case 'fe': case 'perm': case 'sdts': case 'sdte': return 0;
      default: return 1;
    }
  };
  D.isMarker = (it) => D.ilen(it) === 0;
  D.plen = (p) => { let n = 0; for (const it of p.runs) n += D.ilen(it); return n; };
  /** paragraph text with one placeholder char per object (￼), tabs as \t, breaks as \n / \f */
  D.ptext = (p, from, to) => {
    let s = '';
    for (const it of p.runs) {
      if (it.t === 'text') s += it.text;
      else if (it.t === 'tab' || it.t === 'ptab') s += '\t';
      else if (it.t === 'br') s += it.type === 'page' ? '\f' : it.type === 'column' ? '\v' : '\n';
      else if (it.t === 'sym') s += it.char || '￼';
      else if (it.t === 'ruby') s += '￼';
      else if (D.ilen(it)) s += '￼';
    }
    return from != null ? s.slice(from, to == null ? undefined : to) : s;
  };
  /** visible text (for word count, search): objects removed, deleted revisions skipped */
  D.plainText = (p, opts) => {
    let s = '';
    for (const it of p.runs) {
      if (opts && opts.skipDeleted && it.rPr && it.rPr.del) continue;
      if (it.t === 'text') s += it.text;
      else if (it.t === 'tab') s += '\t';
      else if (it.t === 'br') s += it.type === 'line' || !it.type ? '\n' : '';
      else if (it.t === 'sym') s += it.char || '';
      else if (it.t === 'ruby') s += it.base || '';
    }
    return s;
  };

  /* ================= units & numbers ================= */
  D.PX = 96 / 72; /* px per pt */
  D.fmtLen = (pt, unit) => {
    unit = unit || D.unit || 'in';
    if (unit === 'cm') return L.round(pt / 28.3465, 2) + ' cm';
    if (unit === 'mm') return L.round(pt / 2.83465, 1) + ' mm';
    if (unit === 'pt') return L.round(pt, 1) + ' pt';
    if (unit === 'pi') return L.round(pt / 12, 2) + ' pi';
    return L.round(pt / 72, 2) + '"';
  };
  D.parseLen = (s, unit) => {
    const m = /(-?\d*\.?\d+)\s*(cm|mm|pt|pi|in|"|li|line|lines)?/i.exec(String(s));
    if (!m) return null;
    const v = parseFloat(m[1]);
    const u = (m[2] || unit || D.unit || 'in').toLowerCase();
    return { cm: v * 28.3465, mm: v * 2.83465, pt: v, pi: v * 12, in: v * 72, '"': v * 72, li: v * 12, line: v * 12, lines: v * 12 }[u];
  };
  D.unit = 'in';

  /** number formatting for lists, page numbers, fields */
  D.fmtNum = function (n, fmt) {
    switch (fmt) {
      case 'upperRoman': return L.romanize(n);
      case 'lowerRoman': return L.romanize(n).toLowerCase();
      case 'upperLetter': return n > 0 ? String.fromCharCode(65 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1) : '';
      case 'lowerLetter': return n > 0 ? String.fromCharCode(97 + ((n - 1) % 26)).repeat(Math.floor((n - 1) / 26) + 1) : '';
      case 'decimalZero': return n < 10 ? '0' + n : String(n);
      case 'ordinal': { const s = ['th', 'st', 'nd', 'rd'], v = n % 100; return n + (s[(v - 20) % 10] || s[v] || s[0]); }
      case 'cardinalText': return D.numWords(n, false);
      case 'ordinalText': return D.numWords(n, true);
      case 'chicago': { const sy = ['*', '†', '‡', '§']; return sy[(n - 1) % 4].repeat(Math.floor((n - 1) / 4) + 1); }
      case 'decimalEnclosedCircle': return n >= 1 && n <= 20 ? String.fromCharCode(0x2460 + n - 1) : String(n);
      case 'numberInDash': return `- ${n} -`;
      case 'hex': return n.toString(16).toUpperCase();
      case 'none': return '';
      case 'bullet': return '';
      default: return String(n);
    }
  };
  const ONES = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
  const TENS = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];
  const ORD = { one: 'first', two: 'second', three: 'third', five: 'fifth', eight: 'eighth', nine: 'ninth', twelve: 'twelfth' };
  D.numWords = function (n, ordinal) {
    const w = (x) => {
      if (x < 20) return ONES[x];
      if (x < 100) return TENS[Math.floor(x / 10)] + (x % 10 ? '-' + ONES[x % 10] : '');
      if (x < 1000) return ONES[Math.floor(x / 100)] + ' hundred' + (x % 100 ? ' ' + w(x % 100) : '');
      if (x < 1e6) return w(Math.floor(x / 1000)) + ' thousand' + (x % 1000 ? ' ' + w(x % 1000) : '');
      return String(x);
    };
    let s = w(n);
    if (ordinal) {
      const m = /([a-z]+)$/.exec(s);
      const last = m[1];
      const o = ORD[last] || (last.endsWith('y') ? last.slice(0, -1) + 'ieth' : last + 'th');
      s = s.slice(0, -last.length) + o;
    }
    return s.charAt(0).toUpperCase() + s.slice(1);
  };

  /* ================= sections ================= */
  D.defaultSect = () => ({
    pgW: 612, pgH: 792, orient: 'portrait', mt: 72, mb: 72, ml: 90, mr: 90, gutter: 0, hdr: 36, ftr: 36,
    cols: { n: 1, space: 36, sep: false, eq: true, w: [] }, type: 'nextPage', titlePg: false, vAlign: 'top',
    pgNum: { fmt: 'decimal', start: null }, refs: { hdr: {}, ftr: {} }, borders: null, lnNum: null, paperSrc: null,
  });
  D.PAPER = [
    ['Letter', 612, 792], ['Legal', 612, 1008], ['Executive', 522, 756], ['A4', 595.3, 841.9], ['A5', 419.55, 595.3], ['A3', 841.9, 1190.55],
    ['B5 (JIS)', 515.9, 728.5], ['Tabloid', 792, 1224], ['Statement', 396, 612], ['Envelope #10', 684, 297], ['Envelope DL', 623.6, 311.8], ['Envelope C5', 649.15, 459.2], ['Envelope Monarch', 540, 279],
  ];
  D.paperName = (s) => {
    const w = Math.min(s.pgW, s.pgH), hh = Math.max(s.pgW, s.pgH);
    const hit = D.PAPER.find((p) => Math.abs(Math.min(p[1], p[2]) - w) < 2 && Math.abs(Math.max(p[1], p[2]) - hh) < 2);
    return hit ? hit[0] : 'Custom size';
  };

  /* ================= built-in styles (Word 2003 defaults) ================= */
  const S = (id, name, type, o) => Object.assign({ id, name, type, basedOn: type === 'paragraph' && id !== 'Normal' ? 'Normal' : null, next: null, pPr: {}, rPr: {}, builtin: true }, o || {});
  D.builtinStyles = function () {
    const st = [
      S('Normal', 'Normal', 'paragraph', { basedOn: null, q: true }),
      S('Heading1', 'heading 1', 'paragraph', { next: 'Normal', q: true, pPr: { keepNext: true, sp: { b: 12, a: 3 }, outline: 0 }, rPr: { font: 'Arial', b: true, sz: 16, kern: 16 } }),
      S('Heading2', 'heading 2', 'paragraph', { next: 'Normal', q: true, pPr: { keepNext: true, sp: { b: 12, a: 3 }, outline: 1 }, rPr: { font: 'Arial', b: true, i: true, sz: 14 } }),
      S('Heading3', 'heading 3', 'paragraph', { next: 'Normal', q: true, pPr: { keepNext: true, sp: { b: 12, a: 3 }, outline: 2 }, rPr: { font: 'Arial', b: true, sz: 13 } }),
      S('Heading4', 'heading 4', 'paragraph', { next: 'Normal', pPr: { keepNext: true, sp: { b: 12, a: 3 }, outline: 3 }, rPr: { b: true, sz: 14 } }),
      S('Heading5', 'heading 5', 'paragraph', { next: 'Normal', pPr: { sp: { b: 12, a: 3 }, outline: 4 }, rPr: { b: true, i: true, sz: 13 } }),
      S('Heading6', 'heading 6', 'paragraph', { next: 'Normal', pPr: { sp: { b: 12, a: 3 }, outline: 5 }, rPr: { b: true, sz: 11 } }),
      S('Heading7', 'heading 7', 'paragraph', { next: 'Normal', pPr: { sp: { b: 12, a: 3 }, outline: 6 } }),
      S('Heading8', 'heading 8', 'paragraph', { next: 'Normal', pPr: { sp: { b: 12, a: 3 }, outline: 7 }, rPr: { i: true } }),
      S('Heading9', 'heading 9', 'paragraph', { next: 'Normal', pPr: { sp: { b: 12, a: 3 }, outline: 8 }, rPr: { font: 'Arial', sz: 11 } }),
      S('Title', 'Title', 'paragraph', { next: 'Normal', q: true, pPr: { jc: 'center', sp: { b: 12, a: 3 }, outline: 0 }, rPr: { font: 'Arial', b: true, sz: 16, kern: 14 } }),
      S('Subtitle', 'Subtitle', 'paragraph', { next: 'Normal', q: true, pPr: { jc: 'center', sp: { a: 3 }, outline: 1 }, rPr: { font: 'Arial' } }),
      S('BodyText', 'Body Text', 'paragraph', { pPr: { sp: { a: 6 } } }),
      S('BodyText2', 'Body Text 2', 'paragraph', { pPr: { sp: { a: 6, line: 2, rule: 'auto' } } }),
      S('BodyText3', 'Body Text 3', 'paragraph', { pPr: { sp: { a: 6 } }, rPr: { sz: 8 } }),
      S('BodyTextIndent', 'Body Text Indent', 'paragraph', { pPr: { ind: { l: 18 }, sp: { a: 6 } } }),
      S('BodyTextFirstIndent', 'Body Text First Indent', 'paragraph', { basedOn: 'BodyText', pPr: { ind: { fl: 10.5 } } }),
      S('NormalIndent', 'Normal Indent', 'paragraph', { pPr: { ind: { l: 36 } } }),
      S('NormalWeb', 'Normal (Web)', 'paragraph', { pPr: { sp: { b: 14, a: 14, bAuto: true, aAuto: true } } }),
      S('BlockText', 'Block Text', 'paragraph', { pPr: { ind: { l: 72, r: 72 }, sp: { a: 120 / 20 } } }),
      S('Quote', 'Quote', 'paragraph', { pPr: { ind: { l: 36, r: 36 }, sp: { a: 6 } }, rPr: { i: true } }),
      S('PlainText', 'Plain Text', 'paragraph', { rPr: { font: 'Courier New', sz: 10 } }),
      S('Caption', 'caption', 'paragraph', { next: 'Normal', q: true, pPr: { sp: { b: 6, a: 6 } }, rPr: { b: true, sz: 10 } }),
      S('Header', 'header', 'paragraph', { pPr: { tabs: [{ pos: 216, al: 'center' }, { pos: 432, al: 'right' }] } }),
      S('Footer', 'footer', 'paragraph', { pPr: { tabs: [{ pos: 216, al: 'center' }, { pos: 432, al: 'right' }] } }),
      S('FootnoteText', 'footnote text', 'paragraph', { rPr: { sz: 10 } }),
      S('EndnoteText', 'endnote text', 'paragraph', { rPr: { sz: 10 } }),
      S('CommentText', 'annotation text', 'paragraph', { rPr: { sz: 10 } }),
      S('BalloonText', 'Balloon Text', 'paragraph', { rPr: { font: 'Tahoma', sz: 8 } }),
      S('ListParagraph', 'List Paragraph', 'paragraph', { pPr: { ind: { l: 36 } } }),
      S('ListBullet', 'List Bullet', 'paragraph', { pPr: { num: { id: '__bul', lvl: 0 } } }),
      S('ListNumber', 'List Number', 'paragraph', { pPr: { num: { id: '__num', lvl: 0 } } }),
      S('List', 'List', 'paragraph', { pPr: { ind: { l: 18, fl: -18 } } }),
      S('Salutation', 'Salutation', 'paragraph', { next: 'Normal' }),
      S('Closing', 'Closing', 'paragraph', { pPr: { ind: { l: 216 } } }),
      S('Signature', 'Signature', 'paragraph', { pPr: { ind: { l: 216 } } }),
      S('Date', 'Date', 'paragraph', { next: 'Normal' }),
      S('EnvelopeAddress', 'envelope address', 'paragraph', { pPr: { frame: { w: 316.8, h: 158.4, hAnchor: 'page', hRule: 'exact', x: 'center', y: 'bottom' }, ind: { l: 158.4 } }, rPr: { font: 'Arial' } }),
      S('EnvelopeReturn', 'envelope return', 'paragraph', { rPr: { font: 'Arial', sz: 10 } }),
      S('TOCHeading', 'TOC Heading', 'paragraph', { basedOn: 'Heading1', next: 'Normal', pPr: { outline: 9 } }),
      S('TableofFigures', 'table of figures', 'paragraph', { next: 'Normal', pPr: { ind: { l: 24, fl: -24 } } }),
      S('IndexHeading', 'index heading', 'paragraph', { next: 'Index1', rPr: { font: 'Arial', b: true } }),
      S('DocumentMap', 'Document Map', 'paragraph', { pPr: { shd: { fill: '000080', val: 'clear' } }, rPr: { font: 'Tahoma' } }),
      S('MacroText', 'macro', 'paragraph', { pPr: { tabs: [0, 1, 2, 3, 4, 5, 6, 7].map((i) => ({ pos: 24 + i * 24, al: 'left' })) }, rPr: { font: 'Courier New', sz: 10 } }),
      S('HTMLPreformatted', 'HTML Preformatted', 'paragraph', { rPr: { font: 'Courier New', sz: 10 } }),
      S('DefaultParagraphFont', 'Default Paragraph Font', 'character', { basedOn: null, isDefault: true }),
      S('Hyperlink', 'Hyperlink', 'character', { basedOn: null, rPr: { color: '0000FF', u: 'single' } }),
      S('FollowedHyperlink', 'FollowedHyperlink', 'character', { basedOn: null, rPr: { color: '800080', u: 'single' } }),
      S('Strong', 'Strong', 'character', { basedOn: null, q: true, rPr: { b: true } }),
      S('Emphasis', 'Emphasis', 'character', { basedOn: null, q: true, rPr: { i: true } }),
      S('FootnoteReference', 'footnote reference', 'character', { basedOn: null, rPr: { vert: 'superscript' } }),
      S('EndnoteReference', 'endnote reference', 'character', { basedOn: null, rPr: { vert: 'superscript' } }),
      S('CommentReference', 'annotation reference', 'character', { basedOn: null, rPr: { sz: 8 } }),
      S('PageNumber', 'page number', 'character', { basedOn: null }),
      S('LineNumber', 'line number', 'character', { basedOn: null }),
      S('TableNormal', 'Normal Table', 'table', { basedOn: null, isDefault: true, tblPr: { cellMar: { l: 5.4, r: 5.4, t: 0, b: 0 } } }),
      S('TableGrid', 'Table Grid', 'table', { basedOn: 'TableNormal', tblPr: { borders: ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].reduce((o, k) => { o[k] = { val: 'single', sz: 0.5, color: 'auto' }; return o; }, {}), cellMar: { l: 5.4, r: 5.4, t: 0, b: 0 } } }),
      S('NoList', 'No List', 'numbering', { basedOn: null, isDefault: true }),
    ];
    for (let i = 1; i <= 9; i++) st.push(S('TOC' + i, 'toc ' + i, 'paragraph', { next: 'Normal', pPr: { ind: i > 1 ? { l: (i - 1) * 12 } : {} } }));
    for (let i = 1; i <= 3; i++) st.push(S('Index' + i, 'index ' + i, 'paragraph', { next: 'Normal', pPr: { ind: { l: i * 12, fl: -12 } } }));
    const o = {};
    for (const s of st) o[s.id] = s;
    return o;
  };
  D.styleDisplayName = (s) => {
    if (!s) return '';
    const n = s.name || s.id;
    if (s.builtin || /^(heading|toc|index|caption|header|footer|footnote|endnote|annotation|envelope|table of|page number|line number|macro)/.test(n)) return n.charAt(0).toUpperCase() + n.slice(1);
    return n;
  };

  /* ================= list (numbering) definitions ================= */
  D.BULLET_CHARS = [['•', 'Symbol', ''], ['o', 'Courier New', 'o'], ['▪', 'Wingdings', ''], ['❖', 'Wingdings', ''], ['➢', 'Wingdings', ''], ['✓', 'Wingdings', ''], ['◆', 'Wingdings', ''], ['■', 'Wingdings', '']];
  /** A bullet level. ch is the displayed Unicode glyph; font/raw are what Word stores. */
  D.bulletLevel = (lvl, ch, font, raw) => ({ start: 1, fmt: 'bullet', text: raw || ch, glyph: ch, jc: 'left', ind: { l: 36 + lvl * 36, fl: -18 }, rPr: font ? { font } : {}, suff: 'tab' });
  D.numLevel = (lvl, fmt, text, ind) => ({ start: 1, fmt, text, jc: 'left', ind: ind || { l: 36 + lvl * 36, fl: -18 }, rPr: {}, suff: 'tab' });
  D.makeBulletAbs = function (ch, font, raw) {
    const cyc = [[ch, font, raw], ['o', 'Courier New', 'o'], ['▪', 'Wingdings', '']];
    const levels = [];
    for (let i = 0; i < 9; i++) { const c = i === 0 ? cyc[0] : cyc[i % 3]; levels.push(D.bulletLevel(i, c[0], c[1], c[2])); }
    return { levels, multi: 'hybridMultilevel' };
  };
  D.makeNumberAbs = function (fmt, text) {
    const fm = [fmt || 'decimal', 'lowerLetter', 'lowerRoman'];
    const levels = [];
    for (let i = 0; i < 9; i++) {
      const f = i === 0 ? fm[0] : fm[i % 3];
      const t = i === 0 ? (text || '%1.').replace(/%1/g, '%1') : `%${i + 1}.`;
      const lv = D.numLevel(i, f, t);
      if (f === 'lowerRoman' || f === 'upperRoman') lv.jc = 'right', lv.ind = { l: 36 + i * 36, fl: -9 };
      levels.push(lv);
    }
    return { levels, multi: 'hybridMultilevel' };
  };
  /** outline-numbered gallery presets (Bullets and Numbering ▸ Outline Numbered) */
  D.OUTLINE_PRESETS = [
    { name: '1) a) i)', lv: (i) => [['decimal', '%1)'], ['lowerLetter', '%2)'], ['lowerRoman', '%3)'], ['decimal', '(%4)'], ['lowerLetter', '(%5)'], ['lowerRoman', '(%6)'], ['decimal', '%7.'], ['lowerLetter', '%8.'], ['lowerRoman', '%9.']][i] },
    { name: '1. 1.1. 1.1.1.', lv: (i) => ['decimal', Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.') + '.'], legal: true },
    { name: 'Article / Section', lv: (i) => [['upperRoman', 'Article %1.'], ['decimalZero', 'Section %1.%2'], ['lowerLetter', '(%3)'], ['lowerRoman', '(%4)'], ['decimal', '%5)'], ['lowerLetter', '%6)'], ['lowerRoman', '%7)'], ['lowerLetter', '%8.'], ['lowerRoman', '%9.']][i] },
    { name: 'I. A. 1. a)', lv: (i) => [['upperRoman', '%1.'], ['upperLetter', '%2.'], ['decimal', '%3.'], ['lowerLetter', '%4)'], ['decimal', '(%5)'], ['lowerLetter', '(%6)'], ['lowerRoman', '(%7)'], ['lowerLetter', '(%8)'], ['lowerRoman', '(%9)']][i] },
    { name: 'Chapter 1 Heading', lv: (i) => (i === 0 ? ['decimal', 'Chapter %1'] : ['none', '']), headings: true },
    { name: '1 Heading 1.1 Heading', lv: (i) => ['decimal', Array.from({ length: i + 1 }, (_, k) => `%${k + 1}`).join('.')], headings: true },
  ];
  D.makeOutlineAbs = function (preset) {
    const levels = [];
    for (let i = 0; i < 9; i++) {
      const [fmt, text] = preset.lv(i);
      const lv = D.numLevel(i, fmt, text, preset.legal ? { l: 18 * (i + 1) + i * 7, fl: -(18 + i * 7) } : preset.headings ? { l: 0, fl: 0 } : { l: 18 * (i + 1), fl: -18 });
      if (preset.headings) { lv.pStyle = 'Heading' + (i + 1); lv.suff = i === 0 && /Chapter/.test(text) ? 'space' : 'tab'; }
      levels.push(lv);
    }
    return { levels, multi: 'multilevel' };
  };
  D.addNum = function (doc, abs) {
    const nb = doc.numbering;
    // Source definitions need an independent fragment identity. Fresh templates
    // retain this API's in-place id assignment (also used by Markdown import).
    if (L.opc && abs.x) abs = L.opc.duplicate(abs);
    let aid = 0; while (nb.abs[aid] != null) aid++;
    abs.id = aid;
    nb.abs[aid] = abs;
    let nid = 1; while (nb.nums[nid] != null) nid++;
    nb.nums[nid] = { abs: aid, ov: {} };
    return String(nid);
  };
  D.addNumInstance = function (doc, absId, restart) {
    const nb = doc.numbering;
    let nid = 1; while (nb.nums[nid] != null) nid++;
    nb.nums[nid] = { abs: absId, ov: restart ? { 0: { start: 1 } } : {} };
    return String(nid);
  };
  D.numDef = (doc, numId) => { const n = doc.numbering.nums[numId]; return n ? { num: n, abs: doc.numbering.abs[n.abs] } : null; };
  D.numLevelDef = function (doc, numId, lvl) {
    const d = D.numDef(doc, numId);
    if (!d || !d.abs) return null;
    const ov = d.num.ov && d.num.ov[lvl];
    const base = d.abs.levels[lvl] || d.abs.levels[d.abs.levels.length - 1];
    if (ov && ov.lvl) return ov.lvl;
    return base;
  };

  /* ================= new document ================= */
  D.newDoc = function (opts) {
    opts = opts || {};
    const doc = {
      main: { kind: 'main', blocks: [] },
      sect: D.defaultSect(),
      styles: D.builtinStyles(),
      defaults: { rPr: { font: 'Times New Roman', sz: 12, lang: 'en-US' }, pPr: {} },
      numbering: { abs: {}, nums: {} },
      hf: {}, fn: {}, en: {}, comments: {},
      settings: { defTab: 36, evenOdd: false, track: false, protect: null, fnPr: { fmt: 'decimal', pos: 'pageBottom', start: 1, restart: 'continuous' }, enPr: { fmt: 'lowerRoman', pos: 'docEnd', start: 1 }, mirror: false, fieldShade: 'selected', autoHyphen: false },
      props: { title: '', subject: '', creator: L.store.get('userName', 'Quire User'), keywords: '', description: '', lastModifiedBy: '', category: '', company: '', manager: '', created: new Date().toISOString(), modified: new Date().toISOString(), revision: 1 },
      custom: {},
      theme: { major: 'Cambria', minor: 'Calibri' },
      bg: null, watermark: null, rsid: L.uid('r'),
    };
    /* standard list definitions used by the toolbar buttons */
    doc.numbering.abs[0] = Object.assign(D.makeBulletAbs('•', 'Symbol', ''), { id: 0 });
    doc.numbering.abs[1] = Object.assign(D.makeNumberAbs('decimal', '%1.'), { id: 1 });
    doc.numbering.nums[1] = { abs: 0, ov: {} };
    doc.numbering.nums[2] = { abs: 1, ov: {} };
    doc.styles.ListBullet.pPr.num = { id: '1', lvl: 0 };
    doc.styles.ListNumber.pPr.num = { id: '2', lvl: 0 };
    if (opts.modern) {
      doc.defaults.rPr = { font: 'Calibri', sz: 11, lang: 'en-US' };
      doc.defaults.pPr = { sp: { a: 8, line: 1.08, rule: 'auto' } };
      doc.sect.ml = doc.sect.mr = 72;
    }
    const s = L.store.get('defaultSect', null);
    if (s && !opts.modern) Object.assign(doc.sect, s);
    doc.main.blocks.push(D.para());
    return doc;
  };

  /* ================= property resolution ================= */
  const NESTED_P = ['ind', 'sp', 'borders', 'shd', 'num', 'frame', 'dropCap'];
  D.mergePPr = function (base, over) {
    if (!over) return base;
    const o = Object.assign({}, base);
    for (const k in over) {
      const v = over[k];
      if (v === undefined) continue;
      if (k === 'tabs') {
        const t = (o.tabs || []).slice();
        for (const tb of v) {
          const i = t.findIndex((x) => Math.abs(x.pos - tb.pos) < 0.5);
          if (tb.al === 'clear') { if (i >= 0) t.splice(i, 1); continue; }
          if (i >= 0) t[i] = tb; else t.push(tb);
        }
        o.tabs = t.sort((a, b) => a.pos - b.pos);
      } else if (NESTED_P.includes(k) && v && typeof v === 'object' && o[k] && typeof o[k] === 'object') o[k] = Object.assign({}, o[k], v);
      else o[k] = v;
    }
    return o;
  };
  D.mergeRPr = function (base, over) {
    if (!over) return base;
    const o = Object.assign({}, base);
    for (const k in over) if (over[k] !== undefined) o[k] = over[k];
    return o;
  };
  let styleCache = new Map();
  D.stylesChanged = () => { styleCache = new Map(); D.styleVer = (D.styleVer || 0) + 1; };
  D.styleVer = 0;
  D.styleChain = function (doc, id) {
    const chain = [];
    const seen = new Set();
    let s = doc.styles[id];
    while (s && !seen.has(s.id)) { chain.unshift(s); seen.add(s.id); s = s.basedOn ? doc.styles[s.basedOn] : null; }
    return chain;
  };
  const defIdCache = new WeakMap();
  D.defaultStyleId = function (doc, type) {
    /* hot path (called per paragraph / table cell): cache per styles object + style version */
    let c = defIdCache.get(doc.styles);
    if (!c || c.ver !== D.styleVer) { c = { ver: D.styleVer }; defIdCache.set(doc.styles, c); }
    if (type in c) return c[type];
    let id = null;
    for (const k in doc.styles) { const s = doc.styles[k]; if (s.type === type && s.isDefault) { id = s.id; break; } }
    if (!id) id = type === 'paragraph' ? 'Normal' : type === 'table' ? 'TableNormal' : null;
    return (c[type] = id);
  };
  /** resolved {pPr, rPr} of a paragraph style (cached) */
  D.styleProps = function (doc, id) {
    const key = 'p:' + id;
    if (styleCache.has(key)) return styleCache.get(key);
    let pPr = Object.assign({}, doc.defaults.pPr), rPr = Object.assign({}, doc.defaults.rPr);
    const sid = doc.styles[id] ? id : D.defaultStyleId(doc, 'paragraph');
    for (const s of D.styleChain(doc, sid)) { pPr = D.mergePPr(pPr, s.pPr); rPr = D.mergeRPr(rPr, s.rPr); }
    const out = { pPr, rPr };
    styleCache.set(key, out);
    return out;
  };
  D.charStyleProps = function (doc, id) {
    const key = 'c:' + id;
    if (styleCache.has(key)) return styleCache.get(key);
    let rPr = {};
    for (const s of D.styleChain(doc, id)) rPr = D.mergeRPr(rPr, s.rPr);
    styleCache.set(key, rPr);
    return rPr;
  };
  D.tableStyleProps = function (doc, id) {
    const key = 't:' + id;
    if (styleCache.has(key)) return styleCache.get(key);
    let tblPr = {}, pPr = {}, rPr = {}, tcPr = {}, cond = {};
    const sid = doc.styles[id] ? id : D.defaultStyleId(doc, 'table');
    for (const s of D.styleChain(doc, sid)) {
      tblPr = Object.assign({}, tblPr, s.tblPr || {});
      if (s.tblPr && s.tblPr.borders) tblPr.borders = Object.assign({}, (tblPr.borders || {}), s.tblPr.borders);
      pPr = D.mergePPr(pPr, s.pPr); rPr = D.mergeRPr(rPr, s.rPr);
      if (s.tcPr) tcPr = Object.assign({}, tcPr, s.tcPr);
      if (s.cond) for (const k in s.cond) cond[k] = Object.assign({}, cond[k] || {}, s.cond[k]);
    }
    const out = { tblPr, pPr, rPr, tcPr, cond };
    styleCache.set(key, out);
    return out;
  };
  D.isDefaultParaStyle = (doc, sid) => !doc.styles[sid] || sid === D.defaultStyleId(doc, 'paragraph');
  /** the default paragraph style's own settings, and those it inherits (and does not override) from parents */
  D.defaultStyleSplit = function (doc, sid) {
    const key = 'ds:' + sid;
    if (styleCache.has(key)) return styleCache.get(key);
    const chain = D.styleChain(doc, doc.styles[sid] ? sid : D.defaultStyleId(doc, 'paragraph'));
    const own = (chain.length && chain[chain.length - 1].pPr) || {};
    let anc = {};
    for (const st of chain.slice(0, -1)) anc = D.mergePPr(anc, st.pPr);
    const inherited = {};
    for (const k in anc) {
      const v = anc[k], o = own[k];
      if (v === undefined) continue;
      if (o === undefined) inherited[k] = v;
      else if (NESTED_P.includes(k) && v && o && typeof v === 'object' && typeof o === 'object' && !Array.isArray(v)) {
        const sub = {};
        for (const kk in v) if (o[kk] === undefined) sub[kk] = v[kk];
        if (Object.keys(sub).length) inherited[k] = sub;
      }
    }
    const out = { own, inherited };
    styleCache.set(key, out);
    return out;
  };
  /** paragraph properties after style + numbering + direct formatting */
  D.pProps = function (doc, p, ctx) {
    const sid = p.pPr.style || 'Normal';
    let pPr;
    /* precedence: document defaults < table style < paragraph style < numbering < direct */
    if (ctx && ctx.tblPPr) {
      if (D.isDefaultParaStyle(doc, sid)) {
        /* inside a table, settings made in the default paragraph style itself count like document defaults:
           the table style overrides them (what Word and LibreOffice do); settings inherited from a parent
           style still win. overrideTableStyleFontSizeAndJustification gives the style's justification back. */
        const sp = D.defaultStyleSplit(doc, sid);
        pPr = D.mergePPr(D.mergePPr(D.mergePPr(doc.defaults.pPr, sp.own), ctx.tblPPr), sp.inherited);
        if (sp.own.jc && doc.settings && doc.settings.tblFsJcOverride) pPr = Object.assign({}, pPr, { jc: sp.own.jc });
      } else pPr = D.mergePPr(D.mergePPr(doc.defaults.pPr, ctx.tblPPr), D.stylePPrOnly(doc, sid));
    } else pPr = D.styleProps(doc, sid).pPr;
    const num = (p.pPr.num !== undefined ? p.pPr.num : pPr.num);
    if (num && num.id && num.id !== '0') {
      const lv = D.numLevelDef(doc, num.id, num.lvl || 0);
      if (lv && lv.ind) {
        pPr = D.mergePPr(pPr, { ind: lv.ind });
        /* numbering inherited from the paragraph style: the style's own indents still win (Word) */
        if (p.pPr.num === undefined) {
          /* only indents set in the style that brings the numbering, or in styles derived from it */
          const ck = 'ni:' + sid;
          let sInd = styleCache.get(ck);
          if (sInd === undefined) {
            const chain = D.styleChain(doc, sid);
            let k = -1;
            chain.forEach((st, i) => { if (st.pPr && st.pPr.num !== undefined) k = i; });
            sInd = null;
            if (k >= 0) for (let i = k; i < chain.length; i++) if (chain[i].pPr && chain[i].pPr.ind) sInd = Object.assign({}, sInd || {}, chain[i].pPr.ind);
            if (sInd && sInd.l == null && sInd.fl == null) sInd = null;
            styleCache.set(ck, sInd);
          }
          if (sInd) pPr = D.mergePPr(pPr, { ind: sInd });
        }
      }
      if (lv && lv.tabPos != null) pPr = D.mergePPr(pPr, { tabs: [{ pos: lv.tabPos, al: 'num' }] });
    }
    pPr = D.mergePPr(pPr, p.pPr);
    /* a before/after value set on the paragraph itself overrides the style's HTML auto spacing
       (styles that set a value still inherit auto spacing from their base) */
    const dsp = p.pPr.sp;
    if (dsp && pPr.sp && ((dsp.b != null && dsp.bAuto == null && pPr.sp.bAuto) || (dsp.a != null && dsp.aAuto == null && pPr.sp.aAuto))) {
      const m = Object.assign({}, pPr.sp);
      if (dsp.b != null && dsp.bAuto == null) delete m.bAuto;
      if (dsp.a != null && dsp.aAuto == null) delete m.aAuto;
      pPr = Object.assign({}, pPr, { sp: m });
    }
    if (num !== pPr.num) pPr = Object.assign({}, pPr, { num });
    return pPr;
  };
  /** run properties after style chain, table style, character style, direct formatting */
  D.rProps = function (doc, p, it, ctx) {
    const sp = D.styleProps(doc, p.pPr.style || 'Normal');
    let r = sp.rPr;
    if (ctx && ctx.tblRPr) {
      r = D.mergeRPr(D.mergeRPr(doc.defaults.rPr, ctx.tblRPr), D.styleRPrOnly(doc, p.pPr.style || 'Normal'));
      if (ctx.tblRPr.sz && !(doc.settings && doc.settings.tblFsJcOverride) && D.isDefaultParaStyle(doc, p.pPr.style || 'Normal')) r = Object.assign({}, r, { sz: ctx.tblRPr.sz });
    }
    const rp = it ? it.rPr : p.rPr;
    if (rp && rp.style) r = D.mergeRPr(r, D.charStyleProps(doc, rp.style));
    return D.mergeRPr(r, rp);
  };
  D.stylePPrOnly = function (doc, id) {
    const key = 'pp:' + id;
    if (styleCache.has(key)) return styleCache.get(key);
    let pPr = {};
    const sid = doc.styles[id] ? id : D.defaultStyleId(doc, 'paragraph');
    for (const s of D.styleChain(doc, sid)) pPr = D.mergePPr(pPr, s.pPr);
    styleCache.set(key, pPr);
    return pPr;
  };
  /** run properties contributed by the paragraph style chain only (no doc defaults) — used so table styles sit between defaults and paragraph styles */
  D.styleRPrOnly = function (doc, id) {
    const key = 'po:' + id;
    if (styleCache.has(key)) return styleCache.get(key);
    let rPr = {};
    const sid = doc.styles[id] ? id : D.defaultStyleId(doc, 'paragraph');
    /* the whole chain counts, Normal included when the style derives from it; a style that is not
       based on Normal (no w:basedOn) does not pick up Normal's run formatting */
    for (const s of D.styleChain(doc, sid)) rPr = D.mergeRPr(rPr, s.rPr);
    styleCache.set(key, rPr);
    return rPr;
  };

  /* ================= list numbering ================= */
  /**
   * Walk the main story in order and compute list labels. Returns Map(pid → {text, lv, numId, lvl}).
   * Counters are kept per abstract definition so separate num instances of one list continue,
   * unless the instance carries a startOverride (restart).
   */
  D.computeLists = function (doc, stories) {
    const labels = new Map();
    const counters = new Map(); /* absId → int[9] */
    const seenNum = new Set();
    const lastLvl = new Map();
    const visit = (p) => {
      const pPr = D.pProps(doc, p);
      const num = pPr.num;
      if (!num || !num.id || num.id === '0') return;
      /* Word neither shows nor counts a number on an empty paragraph that only carries a section break */
      if (p.sect && !D.plen(p)) return;
      const def = D.numDef(doc, num.id);
      if (!def || !def.abs) return;
      const lvl = L.clamp(num.lvl || 0, 0, 8);
      const absId = def.num.abs;
      let c = counters.get(absId);
      if (!c) { c = new Array(9).fill(null); counters.set(absId, c); }
      if (!seenNum.has(num.id)) {
        seenNum.add(num.id);
        const ov = def.num.ov || {};
        for (const k in ov) if (ov[k] && ov[k].start != null) { c[+k] = null; for (let j = +k + 1; j < 9; j++) c[j] = null; }
        if (Object.keys(ov).some((k) => ov[k] && ov[k].start != null)) def._restart = true;
      }
      const lvDef = D.numLevelDef(doc, num.id, lvl) || {};
      const startOf = (l) => { const ov = def.num.ov && def.num.ov[l]; if (ov && ov.start != null) return ov.start; const d = D.numLevelDef(doc, num.id, l); return d && d.start != null ? d.start : 1; };
      for (let j = 0; j < lvl; j++) if (c[j] == null) c[j] = startOf(j);
      c[lvl] = c[lvl] == null ? startOf(lvl) : c[lvl] + 1;
      for (let j = lvl + 1; j < 9; j++) {
        const d = D.numLevelDef(doc, num.id, j);
        if (!d || d.restart == null || d.restart > lvl) c[j] = null;
      }
      lastLvl.set(absId, lvl);
      let text;
      if (lvDef.fmt === 'bullet') text = lvDef.glyph || L.mapSymbolChar(lvDef.text || '•', lvDef.rPr && lvDef.rPr.font);
      else {
        text = String(lvDef.text == null ? '%' + (lvl + 1) + '.' : lvDef.text).replace(/%(\d)/g, (m, d) => {
          const k = +d - 1;
          if (k > lvl) return '';
          const ld = k === lvl ? lvDef : D.numLevelDef(doc, num.id, k) || {};
          const fmt = lvDef.isLgl && k < lvl ? 'decimal' : ld.fmt || 'decimal';
          return D.fmtNum(c[k] == null ? startOf(k) : c[k], fmt);
        });
      }
      labels.set(p.id, { text, lv: lvDef, numId: num.id, lvl, value: c[lvl] });
    };
    const walk = (blocks) => {
      for (const b of blocks) {
        if (b.t === 'p') visit(b);
        else if (b.t === 'tbl') for (const r of b.rows) for (const c of r.cells) walk(c.blocks);
      }
    };
    walk(doc.main.blocks);
    if (stories) for (const s of stories) walk(s.blocks);
    return labels;
  };

  /* ================= stories, containers, index ================= */
  D.stories = function (doc) {
    const out = [doc.main];
    for (const k in doc.hf) out.push(doc.hf[k]);
    for (const k in doc.fn) out.push(doc.fn[k]);
    for (const k in doc.en) out.push(doc.en[k]);
    for (const k in doc.comments) out.push(doc.comments[k]);
    return out;
  };
  /** visit every block depth-first: fn(block, container, index, story, parents) */
  D.walk = function (container, fn, story, parents) {
    story = story || container;
    parents = parents || [];
    const bl = container.blocks;
    for (let i = 0; i < bl.length; i++) {
      const b = bl[i];
      if (fn(b, container, i, story, parents) === false) return false;
      if (b.t === 'tbl') {
        for (const r of b.rows) for (const c of r.cells) {
          if (D.walk(c, fn, story, parents.concat([{ tbl: b, row: r, cell: c }])) === false) return false;
        }
      } else if (b.t === 'p') {
        for (const it of b.runs) if ((it.t === 'shape' || it.t === 'group') && it.tb) {
          const tbs = it.t === 'group' ? D.groupTextboxes(it) : [it.tb];
          for (const tb of tbs) if (D.walk(tb, fn, tb, parents.concat([{ shape: it, para: b }])) === false) return false;
        }
      }
    }
    return true;
  };
  D.groupTextboxes = (g) => { const out = []; const rec = (k) => { for (const c of k.kids || []) { if (c.tb) out.push(c.tb); if (c.kids) rec(c); } }; rec(g); return out; };

  /**
   * Index of every paragraph in every story: id → {p, cont, i, story, parents}.
   * story.paras lists paragraphs in reading order (tables row-major); p._o is the order index.
   */
  D.reindex = function (doc) {
    const map = new Map();
    const tbls = new Map();
    const index = (story) => {
      const paras = [];
      story.kind = story.kind || 'tb';
      D.walk(story, (b, cont, i, st, parents) => {
        if (b.t === 'p') {
          if (st !== story) return; /* text boxes are indexed as their own stories */
          b._o = paras.length;
          paras.push(b);
          map.set(b.id, { p: b, cont, i, story, parents });
        } else if (b.t === 'tbl') tbls.set(b.id, { tbl: b, cont, i, story, parents });
      });
      story.paras = paras;
    };
    const allStories = D.stories(doc);
    const tbStories = [];
    for (const st of allStories) {
      D.walk(st, (b) => {
        if (b.t === 'p') for (const it of b.runs) if ((it.t === 'shape' || it.t === 'group') && it.tb) {
          for (const tb of it.t === 'group' ? D.groupTextboxes(it) : [it.tb]) {
            tb.kind = 'tb';
            // Runtime back-reference, rebuilt on indexing. Keep it out of JSON
            // snapshots and clipboard copies, where it would make the shape cyclic.
            Object.defineProperty(tb, 'owner', { value: it, writable: true, configurable: true, enumerable: false });
            tbStories.push(tb);
          }
        }
      });
    }
    for (const st of allStories.concat(tbStories)) index(st);
    doc._idx = map;
    doc._tbls = tbls;
    doc._tbStories = tbStories;
    doc._idxDirty = false;
    return map;
  };
  D.info = function (doc, p) {
    if (!p) return null;
    if (doc._idxDirty || !doc._idx) D.reindex(doc);
    const id = typeof p === 'object' ? p.id : p;
    let r = doc._idx.get(id);
    if (r && typeof p === 'object' && r.p !== p) { D.reindex(doc); r = doc._idx.get(id); }
    if (!r) { D.reindex(doc); r = doc._idx.get(id); }
    return r || null;
  };
  D.byId = (doc, id) => { const r = D.info(doc, +id); return r ? r.p : null; };
  D.tblInfo = (doc, t) => { if (doc._idxDirty || !doc._idx) D.reindex(doc); let r = doc._tbls.get(t.id); if (!r || r.tbl !== t) { D.reindex(doc); r = doc._tbls.get(t.id); } return r; };
  D.storyOf = (doc, p) => { const r = D.info(doc, p); return r ? r.story : doc.main; };
  D.invalidate = (doc) => { doc._idxDirty = true; };
  /** innermost table cell containing paragraph p: {tbl,row,cell,ri,ci} or null */
  D.cellOf = function (doc, p) {
    const r = D.info(doc, p);
    if (!r || !r.parents.length) return null;
    const last = r.parents[r.parents.length - 1];
    if (!last.tbl) return null;
    const ri = last.tbl.rows.indexOf(last.row);
    return { tbl: last.tbl, row: last.row, cell: last.cell, ri, ci: last.row.cells.indexOf(last.cell) };
  };
  /** top-level block of the story that contains paragraph p */
  D.topBlock = function (doc, p) {
    const r = D.info(doc, p);
    if (!r) return null;
    if (!r.parents.length) return { block: p, i: r.i, cont: r.cont };
    const t = r.parents[0].tbl || r.parents[0].para;
    const ti = r.story.blocks.indexOf(t);
    return { block: t, i: ti, cont: r.story };
  };
  D.firstPara = (blocks) => { for (const b of blocks) { if (b.t === 'p') return b; if (b.t === 'tbl') { const p = D.firstPara(b.rows[0].cells[0].blocks); if (p) return p; } } return null; };
  D.lastPara = (blocks) => { for (let i = blocks.length - 1; i >= 0; i--) { const b = blocks[i]; if (b.t === 'p') return b; if (b.t === 'tbl') { const r = b.rows[b.rows.length - 1]; const p = D.lastPara(r.cells[r.cells.length - 1].blocks); if (p) return p; } } return null; };
  D.allParas = function (doc, opts) {
    const out = [];
    const stories = opts && opts.all ? D.stories(doc) : [doc.main];
    for (const s of stories) D.walk(s, (b) => { if (b.t === 'p') out.push(b); });
    return out;
  };

  /* ================= positions ================= */
  D.pos = (p, o) => ({ p, o });
  D.cmp = function (doc, a, b) {
    if (a.p === b.p) return a.o - b.o;
    const ia = D.info(doc, a.p), ib = D.info(doc, b.p);
    if (!ia || !ib) return 0;
    if (ia.story !== ib.story) return 0;
    return a.p._o - b.p._o;
  };
  D.order = (doc, a, b) => (D.cmp(doc, a, b) <= 0 ? [a, b] : [b, a]);
  D.eqPos = (a, b) => a && b && a.p === b.p && a.o === b.o;
  /** paragraphs from a to b inclusive, in story order */
  D.parasBetween = function (doc, a, b) {
    const ia = D.info(doc, a), ib = D.info(doc, b);
    if (!ia || !ib || ia.story !== ib.story) return [a];
    const s = ia.story.paras;
    const i0 = Math.min(a._o, b._o), i1 = Math.max(a._o, b._o);
    return s.slice(i0, i1 + 1);
  };
  D.nextPara = (doc, p) => { const r = D.info(doc, p); return r ? r.story.paras[p._o + 1] || null : null; };
  D.prevPara = (doc, p) => { const r = D.info(doc, p); return r ? r.story.paras[p._o - 1] || null : null; };

  /** locate item at offset: {i (item index), off (offset inside item), start (item start offset)}; side: 'before' prefers the item ending at o */
  D.locate = function (p, o, side) {
    let pos = 0;
    const runs = p.runs;
    for (let i = 0; i < runs.length; i++) {
      const n = D.ilen(runs[i]);
      if (side === 'before' ? o <= pos + n && o > pos : o < pos + n) return { i, off: o - pos, start: pos };
      if (n === 0 && o === pos && side === 'marker') return { i, off: 0, start: pos };
      pos += n;
    }
    return { i: runs.length, off: 0, start: pos };
  };
  /** the item covering the character just before offset o (used for "current formatting") */
  D.itemBefore = function (p, o) {
    let pos = 0, last = null;
    for (const it of p.runs) {
      const n = D.ilen(it);
      if (n && pos < o && o <= pos + n) return it;
      if (n && pos + n <= o) last = it;
      if (pos >= o && n) break;
      pos += n;
    }
    return last;
  };
  D.itemAfter = function (p, o) {
    let pos = 0;
    for (const it of p.runs) { const n = D.ilen(it); if (n && pos >= o) return it; if (n && pos <= o && o < pos + n) return it; pos += n; }
    return null;
  };
  /** split runs so that a run boundary exists at offset o; returns the index of the first item at/after o */
  /** Field/comment/bookmark ends precede new content; content-control starts
   *  precede it too, so typing at a control's boundary edits its value. */
  D.insertIndex = function (p, o) {
    let i = D.splitAt(p, o);
    while (i < p.runs.length && (p.runs[i].t === 'fe' || p.runs[i].t === 'ce' || p.runs[i].t === 'be' || p.runs[i].t === 'sdts')) i++;
    return i;
  };
  D.splitAt = function (p, o) {
    let pos = 0;
    for (let i = 0; i < p.runs.length; i++) {
      const it = p.runs[i];
      const n = D.ilen(it);
      if (o === pos) return i;
      if (o < pos + n) {
        if (it.t !== 'text') return i;
        const k = o - pos;
        const a = Object.assign({}, it, { text: it.text.slice(0, k) });
        const b = Object.assign({}, it, { text: it.text.slice(k), rPr: L.clone(it.rPr) });
        p.runs.splice(i, 1, a, b);
        return i + 1;
      }
      pos += n;
    }
    return p.runs.length;
  };
  const sameR = (a, b) => {
    if (a === b) return true;
    const ka = Object.keys(a || {}), kb = Object.keys(b || {});
    if (ka.length !== kb.length) return false;
    for (const k of ka) {
      const x = a[k], y = b[k];
      if (x === y) continue;
      if (x && y && typeof x === 'object' && typeof y === 'object' && JSON.stringify(x) === JSON.stringify(y)) continue;
      return false;
    }
    return true;
  };
  D.sameRPr = sameR;
  /** merge adjacent text runs with equal formatting, drop empty text runs */
  D.normalize = function (p) {
    const out = [];
    for (const it of p.runs) {
      if (it.t === 'text' && !it.text) continue;
      const last = out[out.length - 1];
      if (last && it.t === 'text' && last.t === 'text' && !it.keep?.ac && !last.keep?.ac && sameR(last.rPr, it.rPr)) { out[out.length - 1] = Object.assign({}, last, { text: last.text + it.text }); continue; }
      out.push(it);
    }
    p.runs = out;
    return p;
  };
  /** deep copy of items in [a,b) */
  D.sliceRuns = function (p, a, b) {
    const out = [];
    let pos = 0;
    for (const it of p.runs) {
      const n = D.ilen(it);
      const s = pos, e = pos + n;
      pos = e;
      if (n === 0) { const end = ['ce', 'be', 'fe', 'sdte'].includes(it.t) || it.t === 'perm' && it.end; if (s >= a && s <= b && !(s === b && b !== a && !end) && !(s === a && a !== b && end)) out.push(L.clone(it)); continue; }
      if (e <= a || s >= b) continue;
      if (it.t === 'text') out.push(Object.assign(L.clone(it), { text: it.text.slice(Math.max(0, a - s), Math.min(n, b - s)) }));
      else out.push(L.clone(it));
    }
    return out;
  };

  /* ================= cloning ================= */
  /** deep clone of blocks with fresh ids */
  D.cloneBlocks = function (blocks) {
    const c = L.preserve ? L.preserve.duplicateControls(L.clone(blocks)) : L.clone(blocks);
    const fresh = (bl) => {
      for (const b of bl) {
        b.id = D.nid();
        delete b._o;
        if (b.t === 'tbl') for (const r of b.rows) { r.id = D.nid(); for (const ce of r.cells) { ce.id = D.nid(); fresh(ce.blocks); } }
        if (b.t === 'p') for (const it of b.runs) if (it.tb) fresh(it.tb.blocks);
      }
    };
    fresh(c);
    return c;
  };
  const cloneP = (p) => ({ runs: L.clone(p.runs), pPr: L.clone(p.pPr), rPr: L.clone(p.rPr), sect: p.sect ? L.clone(p.sect) : undefined, mark: p.mark ? L.clone(p.mark) : undefined });
  const restoreP = (p, s) => { p.runs = L.clone(s.runs); p.pPr = L.clone(s.pPr); p.rPr = L.clone(s.rPr); if (s.sect) p.sect = L.clone(s.sect); else delete p.sect; if (s.mark) p.mark = L.clone(s.mark); else delete p.mark; };
  const cloneTbl = (t) => ({ tblPr: L.clone(t.tblPr), grid: t.grid.slice(), rows: t.rows.map((r) => ({ r, trPr: L.clone(r.trPr), cells: r.cells.map((c) => ({ c, tcPr: L.clone(c.tcPr), blocks: c.blocks.slice() })) })) });
  const restoreTbl = (t, s) => {
    t.tblPr = L.clone(s.tblPr); t.grid = s.grid.slice();
    t.rows = s.rows.map((rs) => { rs.r.trPr = L.clone(rs.trPr); rs.r.cells = rs.cells.map((cs) => { cs.c.tcPr = L.clone(cs.tcPr); cs.c.blocks = cs.blocks.slice(); return cs.c; }); return rs.r; });
  };

  /* ================= history (undo / redo) ================= */
  /**
   * Transactions record "before" snapshots of every object they touch:
   *   D.touch(p)            paragraph content & properties
   *   D.touchList(cont)     a container's block list (shallow)
   *   D.touchTbl(t)         a table's structure (rows, cells, properties; cell block lists shallow)
   *   D.touchKey(obj, key)  any property of any object (deep), e.g. (doc,'styles'), (sect,'cols')
   * Commands run inside D.tx(label, fn). Typing merges into one step.
   */
  const H = (D.hist = { undo: [], redo: [], cur: null, depth: 0, limit: 300, seq: 0, savedSeq: 0 });
  D.doc = null;
  D.touch = function (p) {
    const t = H.cur;
    if (!t) return p;
    if (!t.objs.has(p)) t.objs.set(p, { k: 'p', o: p, before: cloneP(p) });
    t.paras.add(p);
    return p;
  };
  D.touchList = function (cont) {
    const t = H.cur;
    if (!t) return cont;
    if (!t.objs.has(cont)) t.objs.set(cont, { k: 'list', o: cont, before: cont.blocks.slice() });
    t.struct = true;
    if (D.doc) D.doc._idxDirty = true;
    return cont;
  };
  D.touchTbl = function (tb) {
    const t = H.cur;
    if (!t) return tb;
    if (!t.objs.has(tb)) t.objs.set(tb, { k: 'tbl', o: tb, before: cloneTbl(tb) });
    t.tbls.add(tb);
    t.struct = true;
    if (D.doc) D.doc._idxDirty = true;
    return tb;
  };
  D.touchKey = function (obj, key) {
    const t = H.cur;
    if (!t) return obj;
    let m = t.keys.get(obj);
    if (!m) { m = new Map(); t.keys.set(obj, m); }
    if (!m.has(key)) m.set(key, L.clone(obj[key]));
    t.global = true;
    if (key === 'styles' || key === 'numbering' || key === 'defaults') D.stylesChanged();
    return obj;
  };
  /** whole-document snapshot for operations too broad to track piecemeal */
  D.touchAll = function () {
    const t = H.cur;
    if (!t || t.full) return;
    t.full = { snap: D.snapshot(D.doc) };
    t.global = true;
  };
  const DOC_KEYS = ['main', 'sect', 'styles', 'defaults', 'numbering', 'hf', 'fn', 'en', 'comments', 'settings', 'props', 'custom', 'theme', 'bg', 'watermark', 'keep', 'losses', 'ooxmlFormat'];
  D.snapshot = (doc) => { const o = {}; for (const k of DOC_KEYS) o[k] = doc[k]; return L.clone(o); };
  D.restoreSnapshot = (doc, s) => { const c = L.clone(s); for (const k of DOC_KEYS) doc[k] = c[k]; doc._idxDirty = true; D.stylesChanged(); };

  D.tx = function (label, fn, opts) {
    opts = opts || {};
    if (H.cur) { H.depth++; try { return fn(); } finally { H.depth--; } }
    const t = { label, objs: new Map(), keys: new Map(), paras: new Set(), tbls: new Set(), struct: false, global: false, full: null, selBefore: D.getSel ? D.getSel() : null, merge: opts.merge || null, time: Date.now() };
    H.cur = t;
    let res, err = null;
    try { const kept = L.preserve?.beforeEdit(D.doc); res = fn(); L.preserve?.afterEdit(D.doc, kept); } catch (e) { err = e; }
    H.cur = null;
    if (err) {
      /* roll back partial changes so the document stays consistent */
      try { undoEntry(t); } catch (e2) { console.error(e2); }
      D.doc._idxDirty = true;
      L.bus.emit('doc-changed', { full: true });
      if (err.code === 'locked-control') { L.bus.emit('history', { sel: t.selBefore }); L.ui?.toast(err.message); return false; }
      throw err;
    }
    if (!t.objs.size && !t.keys.size && !t.full) return res;
    t.selAfter = D.getSel ? D.getSel() : null;
    for (const p of t.paras) p._v = (p._v || 0) + 1;
    for (const tb of t.tbls) tb._v = (tb._v || 0) + 1;
    if (t.struct) D.doc._idxDirty = true;
    /* merge consecutive typing into one undo step */
    const last = H.undo[H.undo.length - 1];
    if (t.merge && last && last.merge === t.merge && !H.breakMerge && t.time - last.time < 4000 && !t.struct && !last.full && !t.full && !t.global && !last.global && t.objs.size === 1 && last.objs.has(t.objs.keys().next().value)) {
      last.time = t.time;
      last.selAfter = t.selAfter;
    } else {
      H.undo.push(t);
      if (H.undo.length > H.limit) H.undo.shift();
    }
    H.breakMerge = false;
    H.redo.length = 0;
    H.seq++;
    D.doc.dirty = true;
    L.bus.emit('doc-changed', { tx: t, paras: t.paras, struct: t.struct, global: t.global || !!t.full });
    return res;
  };
  D.breakMerge = () => { H.breakMerge = true; };
  function snapshotAfter(t) {
    const after = new Map();
    for (const [o, e] of t.objs) after.set(o, e.k === 'p' ? cloneP(o) : e.k === 'list' ? o.blocks.slice() : cloneTbl(o));
    const keysAfter = new Map();
    for (const [o, m] of t.keys) { const n = new Map(); for (const k of m.keys()) n.set(k, L.clone(o[k])); keysAfter.set(o, n); }
    return { after, keysAfter, fullAfter: t.full ? D.snapshot(D.doc) : null };
  }
  function undoEntry(t) {
    if (t.full) { D.restoreSnapshot(D.doc, t.full.snap); return; }
    /* restore in reverse dependency order: lists & tables first, then paragraphs, then keys */
    for (const [o, e] of t.objs) if (e.k === 'list') o.blocks = e.before.slice();
    for (const [o, e] of t.objs) if (e.k === 'tbl') restoreTbl(o, e.before);
    for (const [o, e] of t.objs) if (e.k === 'p') restoreP(o, e.before);
    for (const [o, m] of t.keys) for (const [k, v] of m) o[k] = L.clone(v);
  }
  function redoEntry(t, snap) {
    if (t.full) { D.restoreSnapshot(D.doc, snap.fullAfter); return; }
    for (const [o, e] of t.objs) if (e.k === 'list') o.blocks = snap.after.get(o).slice();
    for (const [o, e] of t.objs) if (e.k === 'tbl') restoreTbl(o, snap.after.get(o));
    for (const [o, e] of t.objs) if (e.k === 'p') restoreP(o, snap.after.get(o));
    for (const [o, m] of snap.keysAfter) for (const [k, v] of m) o[k] = L.clone(v);
  }
  D.canUndo = () => H.undo.length > 0;
  D.canRedo = () => H.redo.length > 0;
  D.undoLabel = () => (H.undo.length ? H.undo[H.undo.length - 1].label : '');
  D.redoLabel = () => (H.redo.length ? H.redo[H.redo.length - 1].label : '');
  D.undo = function () {
    const t = H.undo.pop();
    if (!t) return false;
    const snap = snapshotAfter(t);
    undoEntry(t);
    t._snap = snap;
    H.redo.push(t);
    finishHistory(t, t.selBefore);
    return true;
  };
  D.redo = function () {
    const t = H.redo.pop();
    if (!t) return false;
    redoEntry(t, t._snap);
    H.undo.push(t);
    finishHistory(t, t.selAfter);
    return true;
  };
  function finishHistory(t, sel) {
    for (const [o, e] of t.objs) { if (e.k === 'p') o._v = (o._v || 0) + 1; if (e.k === 'tbl') o._v = (o._v || 0) + 1; }
    D.doc._idxDirty = true;
    D.stylesChanged();
    H.seq++;
    D.doc.dirty = true;
    L.bus.emit('doc-changed', { history: true, full: true });
    L.bus.emit('history', { sel });
  }
  D.resetHistory = () => { H.undo.length = 0; H.redo.length = 0; H.cur = null; H.seq = 0; };
  /** each open window keeps its own undo history */
  D.histState = () => ({ undo: H.undo.slice(), redo: H.redo.slice(), seq: H.seq, savedSeq: H.savedSeq });
  D.setHistState = (st) => { H.undo.length = 0; H.redo.length = 0; H.cur = null; H.depth = 0; if (!st) { H.seq = 0; return; } H.undo.push(...st.undo); H.redo.push(...st.redo); H.seq = st.seq || 0; H.savedSeq = st.savedSeq || 0; };
  D.undoList = () => H.undo.map((t) => t.label).reverse();
  D.redoList = () => H.redo.map((t) => t.label).reverse();

  /* ================= misc helpers ================= */
  /** all sections in order: [{sect, from (block index), to (exclusive), endPara}] for the main story */
  D.sections = function (doc) {
    const out = [];
    let from = 0;
    const bl = doc.main.blocks;
    for (let i = 0; i < bl.length; i++) {
      const b = bl[i];
      if (b.t === 'p' && b.sect) { out.push({ sect: b.sect, from, to: i + 1, endPara: b }); from = i + 1; }
    }
    out.push({ sect: doc.sect, from, to: bl.length, endPara: null });
    return out;
  };
  D.sectOfBlock = function (doc, idx) {
    for (const s of D.sections(doc)) if (idx < s.to || s.endPara === null) return s;
    return { sect: doc.sect };
  };
  /** section properties governing paragraph p (main story) */
  D.sectFor = function (doc, p) {
    const tb = D.topBlock(doc, p);
    if (!tb || tb.cont !== doc.main) return doc.sect;
    return D.sectOfBlock(doc, tb.i).sect;
  };
  D.headingLevel = function (doc, p) {
    const pp = D.pProps(doc, p);
    return pp.outline != null && pp.outline < 9 ? pp.outline + 1 : 0;
  };
  /** words / chars statistics over paragraphs */
  D.stats = function (paras) {
    let words = 0, chars = 0, charsNoSp = 0, paraCount = 0;
    for (const p of paras) {
      const t = D.plainText(p, { skipDeleted: true });
      if (t.trim()) paraCount++;
      const w = t.match(/[^\s    ]+/g);
      words += w ? w.length : 0;
      chars += t.replace(/[\n\t]/g, '').length;
      charsNoSp += t.replace(/[\s ]/g, '').length;
    }
    return { words, chars, charsNoSp, paras: paraCount };
  };
  /** collect all field begin items with their paragraph */
  D.fields = function (doc, stories) {
    const out = [];
    for (const st of stories || D.stories(doc)) D.walk(st, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.t === 'fb') out.push({ p: b, it }); });
    return out;
  };
  D.fieldType = (instr) => (/^\s*([A-Za-z]+)/.exec(instr || '') || [, ''])[1].toUpperCase();
  D.newFid = () => 'f' + D.nid();
  /** used style ids (for the Style box "in use" list) */
  D.usedStyles = function (doc) {
    const s = new Set(['Normal', 'Heading1', 'Heading2', 'Heading3']);
    D.walk(doc.main, (b) => { if (b.t === 'p') { if (b.pPr.style) s.add(b.pPr.style); for (const it of b.runs) if (it.rPr && it.rPr.style) s.add(it.rPr.style); } });
    return s;
  };
  D.findStyleByName = (doc, name) => {
    const n = String(name).toLowerCase();
    for (const k in doc.styles) { const s = doc.styles[k]; if (s.id.toLowerCase() === n || String(s.name).toLowerCase() === n || D.styleDisplayName(s).toLowerCase() === n) return s; }
    return null;
  };
  D.uniqueStyleId = (doc, name) => { let base = String(name).replace(/[^A-Za-z0-9]/g, '') || 'Style'; let id = base, i = 1; while (doc.styles[id]) id = base + (++i); return id; };
})();
