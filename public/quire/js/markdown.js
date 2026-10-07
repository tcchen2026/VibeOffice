/* Quire — Markdown reader: CommonMark 0.31 plus what GitHub adds (tables, task lists, strikethrough,
 * extended autolinks, footnotes, alerts, math spans kept as source). parse() gives a node tree with
 * source line numbers for every block; toHTML() renders it the way the reference implementations do,
 * which is how the parser is checked against the spec examples (tools/quire/test/mdspec.js).
 * Loads in Node as well as in the page.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const M = (L.md = {});

  /* ================= nodes ================= */
  class Node {
    constructor(type, line) {
      this.type = type;
      this.parent = this.first = this.last = this.prev = this.next = null;
      this.sl = line || 0; /* first and last source line (1-based) of a block */
      this.el = line || 0;
      this.open = true;
      this.str = '';
      this.depth = 0;
    }
    append(c) {
      c.unlink();
      c.parent = this;
      if (this.last) { this.last.next = c; c.prev = this.last; this.last = c; } else this.first = this.last = c;
      return c;
    }
    unlink() {
      if (this.prev) this.prev.next = this.next; else if (this.parent) this.parent.first = this.next;
      if (this.next) this.next.prev = this.prev; else if (this.parent) this.parent.last = this.prev;
      this.parent = this.next = this.prev = null;
    }
    insertAfter(s) {
      s.unlink();
      s.next = this.next;
      if (s.next) s.next.prev = s;
      s.prev = this;
      this.next = s;
      s.parent = this.parent;
      if (!s.next && s.parent) s.parent.last = s;
    }
    get kids() { const a = []; for (let c = this.first; c; c = c.next) a.push(c); return a; }
  }
  M.Node = Node;
  const text = (s) => { const n = new Node('text'); n.literal = s; return n; };

  /* ================= characters, entities, URLs ================= */
  const ESC_CHARS = '!"#$%&\'()*+,./:;<=>?@[\\]^_`{|}~-';
  const isEscapable = (c) => !!c && ESC_CHARS.includes(c);
  const reEntityOrEscaped = /\\[!"#$%&'()*+,./:;<=>?@[\\\]^_`{|}~-]|&(?:#x[a-f0-9]{1,6}|#[0-9]{1,7}|[a-z][a-z0-9]{1,31});/gi;
  const reEntityHere = /&(?:#x[a-f0-9]{1,6}|#[0-9]{1,7}|[a-z][a-z0-9]{1,31});/iy;
  const NAMED = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ', copy: '©', reg: '®', trade: '™', hellip: '…', mdash: '—', ndash: '–', laquo: '«', raquo: '»', lsquo: '‘', rsquo: '’', ldquo: '“', rdquo: '”', bull: '•', middot: '·', times: '×', divide: '÷', deg: '°', plusmn: '±', para: '¶', sect: '§', euro: '€', pound: '£', yen: '¥', cent: '¢', larr: '←', rarr: '→', uarr: '↑', darr: '↓', harr: '↔', check: '✓', ensp: ' ', emsp: ' ', thinsp: ' ', zwj: '‍', zwnj: '‌' };
  /** the full HTML5 table, name (without & and ;) → text; tests load it, the page asks the browser */
  M.entities = null;
  let decoder = null;
  function decodeEntity(s) {
    const body = s.slice(1, -1);
    if (body[0] === '#') {
      let n = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
      if (!n || n > 0x10ffff || (n >= 0xd800 && n <= 0xdfff)) n = 0xfffd;
      return String.fromCodePoint(n);
    }
    if (M.entities) return M.entities[body] == null ? null : M.entities[body];
    if (NAMED[body]) return NAMED[body];
    if (typeof document === 'undefined') return null;
    decoder = decoder || document.createElement('textarea');
    decoder.innerHTML = s;
    const v = decoder.value;
    /* a known name decodes to one or two characters; "&notit;" decodes its "&not" prefix and keeps the rest */
    return v !== s && v.length <= 2 && !v.endsWith(';') ? v : null;
  }
  const unescape = (s) => (/[\\&]/.test(s) ? s.replace(reEntityOrEscaped, (m) => (m[0] === '\\' ? m[1] : decodeEntity(m) ?? m)) : s);
  M.unescape = unescape;
  /** percent-encode what may not stand in a URL, keeping existing %XX escapes (as the reference implementations do) */
  function normalizeURI(s) {
    let out = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i], c = s.charCodeAt(i);
      if (ch === '%' && /^[0-9a-fA-F]{2}$/.test(s.slice(i + 1, i + 3))) { out += s.slice(i, i + 3); i += 2; continue; }
      if (c < 128) { out += /[A-Za-z0-9;/?:@&=+$,\-_.!~*'()#]/.test(ch) ? ch : '%' + c.toString(16).toUpperCase().padStart(2, '0'); continue; }
      if (c >= 0xd800 && c <= 0xdbff && i + 1 < s.length && s.charCodeAt(i + 1) >= 0xdc00 && s.charCodeAt(i + 1) <= 0xdfff) { out += encodeURIComponent(ch + s[i + 1]); i++; continue; }
      if (c >= 0xd800 && c <= 0xdfff) { out += '%EF%BF%BD'; continue; }
      out += encodeURIComponent(ch);
    }
    return out;
  }
  const reWhitespaceChar = /^[ \t\n\r\f\p{Zs}]/u;
  const rePunct = /^[\p{P}\p{S}]/u;
  /** link labels match case-insensitively with collapsed whitespace (Unicode case fold via lower → upper) */
  const normLabel = (s) => s.slice(1, -1).trim().replace(/[ \t\r\n]+/g, ' ').toLowerCase().toUpperCase();

  /* ================= HTML patterns ================= */
  const TAGNAME = '[A-Za-z][A-Za-z0-9-]*';
  const ATTRNAME = '[a-zA-Z_:][a-zA-Z0-9:._-]*';
  const ATTRVALUE = '(?:[^"\'=<>`\\x00-\\x20]+|\'[^\']*\'|"[^"]*")';
  const ATTR = '(?:[ \\t\\n\\r\\f]+' + ATTRNAME + '(?:[ \\t\\n\\r\\f]*=[ \\t\\n\\r\\f]*' + ATTRVALUE + ')?)';
  const OPENTAG = '<' + TAGNAME + ATTR + '*[ \\t\\n\\r\\f]*/?>';
  const CLOSETAG = '</' + TAGNAME + '[ \\t\\n\\r\\f]*>';
  const HTMLTAG = '(?:' + OPENTAG + '|' + CLOSETAG + '|<!-->|<!--->|<!--[\\s\\S]*?-->|[<][?][\\s\\S]*?[?][>]|<![A-Za-z][^>]*>|<!\\[CDATA\\[[\\s\\S]*?\\]\\]>)';
  const reHtmlTag = new RegExp(HTMLTAG, 'iy');
  const reHtmlBlockOpen = [
    null,
    /^<(?:script|pre|textarea|style)(?:\s|>|$)/i,
    /^<!--/,
    /^<[?]/,
    /^<![A-Za-z]/,
    /^<!\[CDATA\[/,
    /^<[/]?(?:address|article|aside|base|basefont|blockquote|body|caption|center|col|colgroup|dd|details|dialog|dir|div|dl|dt|fieldset|figcaption|figure|footer|form|frame|frameset|h[123456]|head|header|hr|html|iframe|legend|li|link|main|menu|menuitem|nav|noframes|ol|optgroup|option|p|param|search|section|summary|table|tbody|td|tfoot|th|thead|title|tr|track|ul)(?:\s|[/]?[>]|$)/i,
    new RegExp('^(?:' + OPENTAG + '|' + CLOSETAG + ')\\s*$', 'i'),
  ];
  const reHtmlBlockClose = [null, /<\/(?:script|pre|textarea|style)>/i, /-->/, /\?>/, />/, /\]\]>/];

  /* ================= block patterns ================= */
  const reThematicBreak = /^(?:\*[ \t]*){3,}$|^(?:_[ \t]*){3,}$|^(?:-[ \t]*){3,}$/;
  const reATX = /^#{1,6}(?:[ \t]+|$)/;
  const reCodeFence = /^`{3,}(?!.*`)|^~{3,}/;
  const reClosingFence = /^(?:`{3,}|~{3,})(?=[ \t]*$)/;
  const reSetext = /^(?:=+|-+)[ \t]*$/;
  const reNonSpace = /[^ \t\f\v\r\n]/;
  const reTableDelim = /^\|?[ \t]*:?-+:?[ \t]*(?:\|[ \t]*:?-+:?[ \t]*)*\|?[ \t]*$/;
  const reFootnoteDef = /^\[\^([^\]\s]+)\]:/;
  const reTask = /^\[([ xX])\](?=[ \t\n])[ \t]?/;
  const reAlert = /^\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\][ \t]*(?:\n|$)/i;
  const isBlank = (s) => !reNonSpace.test(s);
  /** quotes and lists nest at most this deep (deeper markers are text), which keeps every walk shallow */
  const MAX_DEPTH = 100;
  const spaceOrTab = (c) => c === ' ' || c === '\t';

  /** a table row's cells: the outer pipes are optional, \| is a pipe inside a cell */
  function splitRow(line) {
    let s = line.trim();
    if (s[0] === '|') s = s.slice(1);
    if (s.endsWith('|') && !/(^|[^\\])(\\\\)*\\\|$/.test(s)) s = s.slice(0, -1);
    const cells = [];
    let cur = '';
    for (let i = 0; i < s.length; i++) {
      const ch = s[i];
      if (ch === '\\' && s[i + 1] === '|') { cur += '|'; i++; continue; }
      if (ch === '\\' && i + 1 < s.length) { cur += ch + s[i + 1]; i++; continue; }
      if (ch === '|') { cells.push(cur); cur = ''; continue; }
      cur += ch;
    }
    cells.push(cur);
    return cells.map((c) => c.trim());
  }

  /* ================= block parser ================= */
  /** opts: { gfm: true (tables, task lists, strikethrough, autolinks, footnotes, alerts), math: true } */
  M.parse = function (src, opts) {
    opts = Object.assign({ gfm: true, math: true }, opts || {});
    const doc = new Node('document', 1);
    const refmap = Object.create(null);
    const footnotes = Object.create(null);
    let tip = doc, oldtip = doc, lastMatched = doc, allClosed = true;
    let line = '', lineNo = 0, offset = 0, column = 0, nextNonspace = 0, nextNonspaceColumn = 0, indent = 0, indented = false, blank = false, partialTab = false, lastLineLength = 0;

    function findNextNonspace() {
      let i = offset, cols = column, c;
      while ((c = line[i]) !== undefined) {
        if (c === ' ') { i++; cols++; } else if (c === '\t') { i++; cols += 4 - (cols % 4); } else break;
      }
      blank = c === undefined || c === '\n' || c === '\r';
      nextNonspace = i;
      nextNonspaceColumn = cols;
      indent = nextNonspaceColumn - column;
      indented = indent >= 4;
    }
    function advanceOffset(count, columns) {
      let c;
      while (count > 0 && (c = line[offset]) !== undefined) {
        if (c === '\t') {
          const toTab = 4 - (column % 4);
          if (columns) {
            partialTab = toTab > count;
            const adv = toTab > count ? count : toTab;
            column += adv;
            offset += partialTab ? 0 : 1;
            count -= adv;
          } else { partialTab = false; column += toTab; offset += 1; count -= 1; }
        } else { partialTab = false; offset += 1; column += 1; count -= 1; }
      }
    }
    function advanceNextNonspace() { offset = nextNonspace; column = nextNonspaceColumn; partialTab = false; }
    function addLine() {
      if (partialTab) { offset += 1; tip.str += ' '.repeat(4 - (column % 4)); }
      tip.str += line.slice(offset) + '\n';
    }
    function finalize(b, ln) {
      const above = b.parent;
      b.open = false;
      b.el = ln;
      B[b.type].finalize(b);
      tip = above;
    }
    function addChild(type) {
      while (!B[tip.type].canContain(type)) finalize(tip, lineNo - 1);
      const n = new Node(type, lineNo);
      n.depth = tip.depth + 1;
      tip.append(n);
      tip = n;
      return n;
    }
    function closeUnmatched() {
      if (allClosed) return;
      while (oldtip !== lastMatched) { const parent = oldtip.parent; finalize(oldtip, lineNo - 1); oldtip = parent; }
      allClosed = true;
    }
    /** link reference definitions at the start of a paragraph go into refmap */
    function takeRefs(b) {
      let pos, any = false;
      const before = b.str;
      while (b.str[0] === '[' && (pos = parseReference(b.str, refmap))) { b.str = b.str.slice(pos); any = true; }
      /* how many source lines the definitions took (the block's own text starts after them) */
      if (any) b.refLines = (b.refLines || 0) + (before.slice(0, before.length - b.str.length).match(/\n/g) || []).length;
      return any;
    }

    /* ---------- block kinds: continue() → 0 matched, 1 not matched, 2 line consumed ---------- */
    const B = {
      document: { continue: () => 0, finalize() {}, canContain: (t) => t !== 'item', lines: false },
      list: {
        continue: () => 0,
        finalize(b) {
          /* loose when any item, or any block inside an item, is followed by a blank line */
          const gap = (n) => n.next && n.el !== n.next.sl - 1;
          for (let item = b.first; item; item = item.next) {
            if (gap(item)) { b.tight = false; break; }
            let found = false;
            for (let sub = item.first; sub; sub = sub.next) if (gap(sub)) { found = true; break; }
            if (found) { b.tight = false; break; }
          }
          b.el = b.last.el;
        },
        canContain: (t) => t === 'item',
        lines: false,
      },
      item: {
        continue(c) {
          if (blank) {
            if (!c.first) return 1;
            advanceNextNonspace();
          } else if (indent >= c.data.markerOffset + c.data.padding) advanceOffset(c.data.markerOffset + c.data.padding, true);
          else return 1;
          return 0;
        },
        finalize(b) { b.el = b.last ? b.last.el : b.sl; },
        canContain: (t) => t !== 'item',
        lines: false,
      },
      footnote_def: {
        continue() {
          if (blank) { advanceNextNonspace(); return 0; }
          if (indent >= 4) { advanceOffset(4, true); return 0; }
          return 1;
        },
        finalize(b) { b.el = b.last ? b.last.el : b.sl; },
        canContain: (t) => t !== 'item',
        lines: false,
      },
      block_quote: {
        continue() {
          if (!indented && line[nextNonspace] === '>') {
            advanceNextNonspace();
            advanceOffset(1, false);
            if (spaceOrTab(line[offset])) advanceOffset(1, true);
            return 0;
          }
          return 1;
        },
        finalize() {},
        canContain: (t) => t !== 'item',
        lines: false,
      },
      heading: { continue: () => 1, finalize() {}, canContain: () => false, lines: false },
      thematic_break: { continue: () => 1, finalize() {}, canContain: () => false, lines: false },
      table: { continue: () => (blank ? 1 : 0), finalize() {}, canContain: () => false, lines: false },
      code_block: {
        continue(c) {
          if (c.fenced) {
            const m = indent <= 3 && line[nextNonspace] === c.fenceChar && reClosingFence.exec(line.slice(nextNonspace));
            if (m && m[0].length >= c.fenceLength) {
              lastLineLength = offset + indent + m[0].length;
              finalize(c, lineNo);
              return 2;
            }
            let i = c.fenceOffset;
            while (i > 0 && spaceOrTab(line[offset])) { advanceOffset(1, true); i--; }
          } else if (indent >= 4) advanceOffset(4, true);
          else if (blank) advanceNextNonspace();
          else return 1;
          return 0;
        },
        finalize(b) {
          if (b.fenced) {
            const nl = b.str.indexOf('\n');
            b.info = unescape(b.str.slice(0, nl).trim());
            b.literal = b.str.slice(nl + 1);
          } else {
            const lines = b.str.split('\n');
            while (lines.length && /^[ \t]*$/.test(lines[lines.length - 1])) lines.pop();
            b.literal = lines.join('\n') + '\n';
            b.el = b.sl + lines.length - 1;
          }
          b.str = '';
        },
        canContain: () => false,
        lines: true,
      },
      html_block: {
        continue: (c) => (blank && (c.htmlType === 6 || c.htmlType === 7) ? 1 : 0),
        finalize(b) { b.literal = b.str.replace(/\n+$/, ''); b.str = ''; },
        canContain: () => false,
        lines: true,
      },
      paragraph: {
        continue: () => (blank ? 1 : 0),
        finalize(b) { if (takeRefs(b) && isBlank(b.str)) { b.refsOnly = true; b.unlink(); } },
        canContain: () => false,
        lines: true,
      },
    };

    /* ---------- block starts: 0 no match, 1 container, 2 leaf ---------- */
    function parseListMarker(container) {
      const rest = line.slice(nextNonspace);
      const data = { type: null, tight: true, bulletChar: null, start: null, delimiter: null, padding: null, markerOffset: indent };
      if (indent >= 4) return null;
      let m;
      if ((m = /^[*+-]/.exec(rest))) { data.type = 'bullet'; data.bulletChar = m[0]; }
      else if ((m = /^(\d{1,9})([.)])/.exec(rest)) && (container.type !== 'paragraph' || m[1] === '1')) { data.type = 'ordered'; data.start = parseInt(m[1], 10); data.delimiter = m[2]; }
      else return null;
      const nextc = line[nextNonspace + m[0].length];
      if (!(nextc === undefined || nextc === '\t' || nextc === ' ')) return null;
      if (container.type === 'paragraph' && !reNonSpace.test(line.slice(nextNonspace + m[0].length))) return null;
      advanceNextNonspace();
      advanceOffset(m[0].length, true);
      const spacesStartCol = column, spacesStartOffset = offset;
      do { advanceOffset(1, true); } while (column - spacesStartCol < 5 && spaceOrTab(line[offset]));
      const blankItem = line[offset] === undefined;
      const after = column - spacesStartCol;
      if (after >= 5 || after < 1 || blankItem) {
        data.padding = m[0].length + 1;
        column = spacesStartCol;
        offset = spacesStartOffset;
        if (spaceOrTab(line[offset])) advanceOffset(1, true);
      } else data.padding = m[0].length + after;
      return data;
    }
    const starts = [
      /* block quote */
      (container) => {
        if (indented || line[nextNonspace] !== '>' || container.depth >= MAX_DEPTH) return 0;
        advanceNextNonspace();
        advanceOffset(1, false);
        if (spaceOrTab(line[offset])) advanceOffset(1, true);
        closeUnmatched();
        addChild('block_quote');
        return 1;
      },
      /* ATX heading */
      () => {
        let m;
        if (indented || !(m = reATX.exec(line.slice(nextNonspace)))) return 0;
        advanceNextNonspace();
        advanceOffset(m[0].length, false);
        closeUnmatched();
        const h = addChild('heading');
        h.level = m[0].trim().length;
        h.str = line.slice(offset).replace(/^[ \t]*#+[ \t]*$/, '').replace(/[ \t]+#+[ \t]*$/, '');
        advanceOffset(line.length - offset);
        return 2;
      },
      /* fenced code */
      () => {
        let m;
        if (indented || !(m = reCodeFence.exec(line.slice(nextNonspace)))) return 0;
        closeUnmatched();
        const c = addChild('code_block');
        c.fenced = true;
        c.fenceLength = m[0].length;
        c.fenceChar = m[0][0];
        c.fenceOffset = indent;
        advanceNextNonspace();
        advanceOffset(m[0].length, false);
        return 2;
      },
      /* HTML block */
      (container) => {
        if (indented || line[nextNonspace] !== '<') return 0;
        const s = line.slice(nextNonspace);
        for (let t = 1; t <= 7; t++) {
          if (reHtmlBlockOpen[t].test(s) && (t < 7 || (container.type !== 'paragraph' && !(!allClosed && !blank && tip.type === 'paragraph')))) {
            closeUnmatched();
            const b = addChild('html_block');
            b.htmlType = t;
            return 2;
          }
        }
        return 0;
      },
      /* setext heading */
      (container) => {
        let m;
        if (indented || container.type !== 'paragraph' || !(m = reSetext.exec(line.slice(nextNonspace)))) return 0;
        closeUnmatched();
        takeRefs(container);
        if (!container.str.length) return 0;
        const h = new Node('heading', container.sl);
        h.level = m[0][0] === '=' ? 1 : 2;
        h.str = container.str;
        h.refLines = container.refLines;
        container.insertAfter(h);
        container.unlink();
        tip = h;
        advanceOffset(line.length - offset, false);
        return 2;
      },
      /* table: the paragraph's last line is the header row, this line the delimiter row */
      (container) => {
        if (!opts.gfm || indented || container.type !== 'paragraph') return 0;
        const s = line.slice(nextNonspace);
        if (!reTableDelim.test(s)) return 0;
        const lines = container.str.replace(/\n$/, '').split('\n');
        const headLine = lines[lines.length - 1];
        const delim = splitRow(s), head = splitRow(headLine);
        if (head.length !== delim.length || (!s.includes('|') && !/(^|[^\\])\|/.test(headLine))) return 0;
        closeUnmatched();
        const t = new Node('table', container.sl + lines.length - 1);
        t.align = delim.map((c) => (/^:-+:$/.test(c) ? 'center' : /^:-+$/.test(c) ? 'left' : /^-+:$/.test(c) ? 'right' : ''));
        t.rows = [head];
        const parent = container.parent;
        if (lines.length > 1) {
          /* the lines above the header row stay a paragraph */
          container.str = lines.slice(0, -1).join('\n') + '\n';
          finalize(container, t.sl - 1);
          parent.append(t);
        } else { container.insertAfter(t); container.unlink(); }
        tip = t;
        advanceOffset(line.length - offset, false);
        return 2;
      },
      /* thematic break */
      () => {
        if (indented || !reThematicBreak.test(line.slice(nextNonspace))) return 0;
        closeUnmatched();
        addChild('thematic_break');
        advanceOffset(line.length - offset, false);
        return 2;
      },
      /* footnote definition */
      (container) => {
        let m;
        if (!opts.gfm || indented || container.type === 'paragraph' || !(m = reFootnoteDef.exec(line.slice(nextNonspace)))) return 0;
        advanceNextNonspace();
        advanceOffset(m[0].length, false);
        findNextNonspace();
        advanceNextNonspace();
        closeUnmatched();
        const f = addChild('footnote_def');
        f.label = m[1];
        const key = m[1].toUpperCase();
        if (!footnotes[key]) footnotes[key] = f;
        return 1;
      },
      /* list item */
      (container) => {
        let data;
        if (container.depth >= MAX_DEPTH || !((!indented || container.type === 'list') && (data = parseListMarker(container)))) return 0;
        closeUnmatched();
        if (tip.type !== 'list' || !(tip.data.type === data.type && tip.data.delimiter === data.delimiter && tip.data.bulletChar === data.bulletChar)) {
          const l = addChild('list');
          l.data = data;
          l.tight = true;
        }
        const it = addChild('item');
        it.data = data;
        return 1;
      },
      /* indented code */
      () => {
        if (!indented || tip.type === 'paragraph' || blank) return 0;
        advanceOffset(4, true);
        closeUnmatched();
        addChild('code_block');
        return 2;
      },
    ];

    function incorporateLine(ln) {
      let allMatched = true;
      let container = doc;
      oldtip = tip;
      offset = 0; column = 0; blank = false; partialTab = false;
      lineNo++;
      line = ln.includes('\0') ? ln.replace(/\0/g, '�') : ln;
      let last;
      while ((last = container.last) && last.open) {
        container = last;
        findNextNonspace();
        const r = B[container.type].continue(container);
        if (r === 2) return;
        if (r === 1) { allMatched = false; container = container.parent; break; }
      }
      void allMatched;
      allClosed = container === oldtip;
      lastMatched = container;
      let matchedLeaf = container.type !== 'paragraph' && B[container.type].lines;
      while (!matchedLeaf) {
        findNextNonspace();
        let i = 0, r = 0;
        for (; i < starts.length; i++) {
          r = starts[i](container);
          if (r) break;
        }
        if (!r) { advanceNextNonspace(); break; }
        container = tip;
        if (r === 2) matchedLeaf = true;
      }
      if (!allClosed && !blank && tip.type === 'paragraph') addLine(); /* lazy continuation */
      else {
        closeUnmatched();
        const t = container.type;
        if (B[t].lines) {
          addLine();
          if (t === 'html_block' && container.htmlType <= 5 && reHtmlBlockClose[container.htmlType].test(line.slice(offset))) {
            lastLineLength = ln.length;
            finalize(container, lineNo);
          }
        } else if (t === 'table') {
          if (!blank && offset < line.length) { container.rows.push(splitRow(line.slice(offset))); container.el = lineNo; }
        } else if (offset < line.length && !blank) {
          addChild('paragraph');
          advanceNextNonspace();
          addLine();
        }
      }
      lastLineLength = ln.length;
      void lastLineLength;
    }

    const lines = String(src).split(/\r\n|\n|\r/);
    if (lines.length && lines[lines.length - 1] === '') lines.pop();
    for (const ln of lines) incorporateLine(ln);
    while (tip) finalize(tip, lines.length);
    doc.el = lines.length;

    /* ---------- inlines ---------- */
    const P = new InlineParser(refmap, footnotes, opts);
    const walkBlocks = (n) => {
      for (let c = n.first; c; c = c.next) {
        if (c.type === 'paragraph' || c.type === 'heading') {
          if (c.type === 'paragraph' && opts.gfm && c.parent.type === 'item' && c.parent.first === c) {
            const m = reTask.exec(c.str);
            if (m) { c.parent.task = { checked: m[1] !== ' ' }; c.str = c.str.slice(m[0].length); }
          }
          if (c.type === 'paragraph' && opts.gfm && c.parent.type === 'block_quote' && c.parent.first === c && !c.parent.alert) {
            const m = reAlert.exec(c.str);
            if (m) { c.parent.alert = m[1].toLowerCase(); c.str = c.str.slice(m[0].length); }
          }
          P.parse(c, c.str);
          c.str = '';
        } else if (c.type === 'table') {
          const n = c.align.length;
          c.rows.forEach((cells, ri) => {
            const row = c.append(new Node('table_row', c.sl + (ri ? ri + 1 : 0)));
            row.head = ri === 0;
            for (let k = 0; k < n; k++) {
              const cell = row.append(new Node('table_cell'));
              cell.align = c.align[k];
              P.parse(cell, cells[k] || '');
            }
          });
          delete c.rows;
        } else walkBlocks(c);
      }
    };
    walkBlocks(doc);
    /* drop paragraphs emptied by an alert or task marker? no: keep them, they have no inlines */
    doc.refmap = refmap;
    doc.footnotes = footnotes;
    /* footnotes are numbered in the order they are first referenced */
    doc.footnoteOrder = P.fnOrder;
    return doc;
  };

  /* ================= link reference definitions ================= */
  function parseReference(s, refmap) {
    const p = new InlineParser(refmap, {}, {});
    p.subject = s;
    p.pos = 0;
    const startpos = 0;
    const n = p.parseLinkLabel();
    if (n === 0) return 0;
    const rawlabel = s.slice(0, n);
    if (p.peek() !== ':') return 0;
    p.pos++;
    p.spnl();
    const dest = p.parseLinkDestination();
    if (dest === null) return 0;
    const beforetitle = p.pos;
    p.spnl();
    let title = null;
    if (p.pos !== beforetitle) title = p.parseLinkTitle();
    if (title === null) { title = ''; p.pos = beforetitle; }
    let atLineEnd = true;
    if (p.match(/[ \t]*(?:\n|$)/y) === null) {
      if (title === '') atLineEnd = false;
      else {
        title = '';
        p.pos = beforetitle;
        atLineEnd = p.match(/[ \t]*(?:\n|$)/y) !== null;
      }
    }
    if (!atLineEnd) return 0;
    const norm = normLabel(rawlabel);
    if (norm === '') return 0;
    if (!refmap[norm]) refmap[norm] = { destination: dest, title };
    return p.pos - startpos;
  }

  /* ================= inline parser ================= */
  function InlineParser(refmap, footnotes, opts) {
    this.refmap = refmap;
    this.footnotes = footnotes;
    this.opts = opts;
    this.fnOrder = [];
    this.subject = '';
    this.pos = 0;
    this.delimiters = null;
    this.brackets = null;
  }
  const IP = InlineParser.prototype;
  IP.peek = function () { return this.subject[this.pos]; };
  /** match a sticky (y) or global (g) regex at pos; advances pos past the match */
  IP.match = function (re) {
    re.lastIndex = this.pos;
    const m = re.exec(this.subject);
    if (!m) return null;
    this.pos = m.index + m[0].length;
    return m[0];
  };
  IP.spnl = function () { this.match(/[ \t]*(?:\n[ \t]*)?/y); return true; };
  IP.parse = function (block, s) {
    this.subject = s.replace(/^[ \t\n]+/, '').replace(/[ \t\n]+$/, '');
    this.pos = 0;
    this.delimiters = null;
    this.brackets = null;
    while (this.parseInline(block));
    this.processEmphasis(null);
    mergeText(block);
    if (this.opts.gfm) autolinks(block);
  };
  IP.parseInline = function (block) {
    const c = this.peek();
    if (c === undefined) return false;
    let res = false;
    switch (c) {
      case '\n': res = this.parseNewline(block); break;
      case '\\': res = this.parseBackslash(block); break;
      case '`': res = this.parseBackticks(block); break;
      case '*': case '_': res = this.handleDelim(c, block); break;
      case '~': res = this.opts.gfm && this.handleDelim(c, block); break;
      case '$': res = this.opts.math && this.parseMath(block); break;
      case '[': res = this.parseOpenBracket(block); break;
      case '!': res = this.parseBang(block); break;
      case ']': res = this.parseCloseBracket(block); break;
      case '<': res = this.parseAutolink(block) || this.parseHtmlTag(block); break;
      case '&': res = this.parseEntity(block); break;
      default: res = this.parseString(block);
    }
    if (!res) { this.pos++; block.append(text(c)); }
    return true;
  };
  IP.parseString = function (block) {
    const m = this.match(this.opts.math ? /[^\n`[\]\\!<&*_~$]+/y : /[^\n`[\]\\!<&*_~]+/y);
    if (m === null) return false;
    block.append(text(m));
    return true;
  };
  IP.parseNewline = function (block) {
    this.pos++;
    const last = block.last;
    if (last && last.type === 'text' && last.literal.endsWith(' ')) {
      const hard = last.literal.endsWith('  ');
      last.literal = last.literal.replace(/ +$/, '');
      block.append(new Node(hard ? 'linebreak' : 'softbreak'));
    } else block.append(new Node('softbreak'));
    this.match(/[ \t]*/y);
    return true;
  };
  IP.parseBackslash = function (block) {
    this.pos++;
    const c = this.peek();
    if (c === '\n') { this.pos++; block.append(new Node('linebreak')); this.match(/[ \t]*/y); }
    else if (isEscapable(c)) { block.append(text(c)); this.pos++; }
    else block.append(text('\\'));
    return true;
  };
  IP.parseBackticks = function (block) {
    const ticks = this.match(/`+/y);
    if (ticks === null) return false;
    const after = this.pos;
    let m;
    while ((m = this.match(/`+/g)) !== null) {
      if (m === ticks) {
        const n = new Node('code');
        let s = this.subject.slice(after, this.pos - ticks.length).replace(/\n/g, ' ');
        if (s.length > 0 && /[^ ]/.test(s) && s[0] === ' ' && s[s.length - 1] === ' ') s = s.slice(1, -1);
        n.literal = s;
        block.append(n);
        return true;
      }
    }
    this.pos = after;
    block.append(text(ticks));
    return true;
  };
  /** GitHub math: $…$ (not next to spaces, the closing $ not followed by a digit), $`…`$ and $$…$$; kept as source */
  IP.parseMath = function (block) {
    const s = this.subject, start = this.pos;
    let m;
    if (s[start + 1] === '`') {
      const end = s.indexOf('`$', start + 2);
      if (end < 0) return false;
      m = s.slice(start, end + 2);
    } else if (s[start + 1] === '$') {
      const end = s.indexOf('$$', start + 2);
      if (end < 0 || end === start + 2) return false;
      m = s.slice(start, end + 2);
    } else {
      if (start > 0 && s[start - 1] === '$') return false;
      const c = s[start + 1];
      if (c === undefined || /\s/.test(c)) return false;
      let i = start + 1;
      for (; i < s.length; i++) {
        if (s[i] === '\\') { i++; continue; }
        if (s[i] === '$') break;
      }
      if (i >= s.length || /\s/.test(s[i - 1]) || /[0-9]/.test(s[i + 1] || '')) return false;
      m = s.slice(start, i + 1);
    }
    const n = new Node('math');
    n.literal = m;
    block.append(n);
    this.pos = start + m.length;
    return true;
  };

  /* ---------- emphasis ---------- */
  IP.scanDelims = function (c) {
    const start = this.pos;
    let n = 0;
    while (this.peek() === c) { n++; this.pos++; }
    if (n === 0) return null;
    const before = start === 0 ? '\n' : String.fromCodePoint(this.subject.codePointAt(start - 1 - (start > 1 && /[\udc00-\udfff]/.test(this.subject[start - 1]) ? 1 : 0)));
    const aft = this.peek() === undefined ? '\n' : String.fromCodePoint(this.subject.codePointAt(this.pos));
    const afterWs = reWhitespaceChar.test(aft), afterPunct = rePunct.test(aft);
    const beforeWs = reWhitespaceChar.test(before), beforePunct = rePunct.test(before);
    const left = !afterWs && (!afterPunct || beforeWs || beforePunct);
    const right = !beforeWs && (!beforePunct || afterWs || afterPunct);
    let open, close;
    if (c === '_') { open = left && (!right || beforePunct); close = right && (!left || afterPunct); }
    else { open = left; close = right; }
    this.pos = start;
    return { n, open, close };
  };
  IP.handleDelim = function (c, block) {
    const r = this.scanDelims(c);
    if (!r) return false;
    const start = this.pos;
    this.pos += r.n;
    const node = text(this.subject.slice(start, this.pos));
    block.append(node);
    if ((r.open || r.close) && !(c === '~' && r.n > 2)) {
      this.delimiters = { c, n: r.n, orig: r.n, node, previous: this.delimiters, next: null, open: r.open, close: r.close };
      if (this.delimiters.previous) this.delimiters.previous.next = this.delimiters;
    }
    return true;
  };
  IP.removeDelimiter = function (d) {
    if (d.previous) d.previous.next = d.next;
    if (d.next) d.next.previous = d.previous;
    else this.delimiters = d.previous;
  };
  const removeBetween = (bottom, top) => { if (bottom.next !== top) { bottom.next = top; top.previous = bottom; } };
  IP.processEmphasis = function (bottom) {
    const openersBottom = {};
    let closer = this.delimiters;
    while (closer && closer.previous !== bottom) closer = closer.previous;
    while (closer) {
      if (!closer.close) { closer = closer.next; continue; }
      const key = closer.c === '~' ? '~' + closer.n : closer.c + (closer.open ? 3 : 0) + (closer.orig % 3);
      let opener = closer.previous, found = false;
      const floor = key in openersBottom ? openersBottom[key] : bottom;
      while (opener && opener !== bottom && opener !== floor) {
        if (closer.c === '~') { if (opener.c === '~' && opener.open && opener.n === closer.n) { found = true; break; } }
        else {
          const odd = (closer.open || opener.close) && closer.orig % 3 !== 0 && (opener.orig + closer.orig) % 3 === 0;
          if (opener.c === closer.c && opener.open && !odd) { found = true; break; }
        }
        opener = opener.previous;
      }
      const oldCloser = closer;
      if (found) {
        const use = closer.c === '~' ? closer.n : closer.n >= 2 && opener.n >= 2 ? 2 : 1;
        const oi = opener.node, ci = closer.node;
        opener.n -= use; closer.n -= use;
        oi.literal = oi.literal.slice(0, oi.literal.length - use);
        ci.literal = ci.literal.slice(0, ci.literal.length - use);
        const em = new Node(closer.c === '~' ? 'del' : use === 1 ? 'emph' : 'strong');
        let t = oi.next;
        while (t && t !== ci) { const nx = t.next; em.append(t); t = nx; }
        oi.insertAfter(em);
        removeBetween(opener, closer);
        if (opener.n === 0) { oi.unlink(); this.removeDelimiter(opener); }
        if (closer.n === 0) { ci.unlink(); const nx = closer.next; this.removeDelimiter(closer); closer = nx; }
      } else {
        closer = closer.next;
        openersBottom[key] = oldCloser.previous;
        if (!oldCloser.open) this.removeDelimiter(oldCloser);
      }
    }
    while (this.delimiters && this.delimiters !== bottom) this.removeDelimiter(this.delimiters);
  };

  /* ---------- links and images ---------- */
  IP.parseLinkTitle = function () {
    const m = this.match(/"(?:\\[\s\S]|[^\\"\x00])*"|'(?:\\[\s\S]|[^\\'\x00])*'|\((?:\\[\s\S]|[^\\()\x00])*\)/y);
    return m === null ? null : unescape(m.slice(1, -1));
  };
  IP.parseLinkDestination = function () {
    const m = this.match(/<(?:[^<>\n\\\x00]|\\.)*>/y);
    if (m !== null) return normalizeURI(unescape(m.slice(1, -1)));
    if (this.peek() === '<') return null;
    const save = this.pos;
    let parens = 0, c;
    while ((c = this.peek()) !== undefined) {
      if (c === '\\' && isEscapable(this.subject[this.pos + 1])) { this.pos += 2; continue; }
      if (c === '(') { this.pos++; parens++; }
      else if (c === ')') { if (parens < 1) break; this.pos++; parens--; }
      else if (/[\x00-\x20\x7f]/.test(c)) break;
      else this.pos++;
    }
    if (this.pos === save && c !== ')') return null;
    if (parens !== 0) return null;
    return normalizeURI(unescape(this.subject.slice(save, this.pos)));
  };
  IP.parseLinkLabel = function () {
    const m = this.match(/\[(?:[^\\[\]]|\\[\s\S]){0,1000}\]/y);
    return m === null || m.length > 1001 ? 0 : m.length;
  };
  IP.addBracket = function (node, index, image) {
    if (this.brackets) this.brackets.bracketAfter = true;
    this.brackets = { node, previous: this.brackets, previousDelimiter: this.delimiters, index, image, active: true };
  };
  IP.parseOpenBracket = function (block) {
    const start = this.pos;
    this.pos++;
    const n = text('[');
    block.append(n);
    this.addBracket(n, start, false);
    return true;
  };
  IP.parseBang = function (block) {
    const start = this.pos;
    this.pos++;
    if (this.peek() === '[') {
      this.pos++;
      const n = text('![');
      block.append(n);
      this.addBracket(n, start + 1, true);
    } else block.append(text('!'));
    return true;
  };
  IP.parseCloseBracket = function (block) {
    this.pos++;
    const start = this.pos;
    let opener = this.brackets;
    if (!opener) { block.append(text(']')); return true; }
    if (!opener.active) { block.append(text(']')); this.brackets = opener.previous; return true; }
    const isImage = opener.image;
    let dest, title, matched = false, reflabel;
    const save = this.pos;
    if (this.peek() === '(') {
      this.pos++;
      if (this.spnl() && (dest = this.parseLinkDestination()) !== null && this.spnl() &&
        ((reWhitespaceChar.test(this.subject[this.pos - 1]) && (title = this.parseLinkTitle())) || true) &&
        this.spnl() && this.peek() === ')') { this.pos++; matched = true; }
      else this.pos = save;
    }
    if (!matched) {
      const before = this.pos;
      const n = this.parseLinkLabel();
      if (n > 2) reflabel = this.subject.slice(before, before + n);
      else if (!opener.bracketAfter) reflabel = this.subject.slice(opener.index, start);
      if (n === 0) this.pos = save;
      /* GitHub footnote reference [^label] */
      const inner = this.subject.slice(opener.index, start);
      if (this.opts.gfm && /^\[\^[^\]\s]+\]$/.test(inner) && n === 0) {
        const key = inner.slice(2, -1).toUpperCase();
        const def = this.footnotes[key];
        if (def) {
          const fr = new Node('footnote_ref');
          fr.label = def.label;
          if (!this.fnOrder.includes(def)) this.fnOrder.push(def);
          fr.def = def;
          let t = opener.node.next;
          while (t) { const nx = t.next; t.unlink(); t = nx; }
          if (isImage) opener.node.literal = '!'; else opener.node.unlink();
          block.append(fr);
          this.brackets = opener.previous;
          return true;
        }
      }
      if (reflabel) {
        const link = this.refmap[normLabel(reflabel)];
        if (link) { dest = link.destination; title = link.title; matched = true; }
      }
    }
    if (matched) {
      const node = new Node(isImage ? 'image' : 'link');
      node.dest = dest;
      node.title = title || '';
      if (reflabel && !this.subject.slice(save, save + 1).match(/\(/)) node.ref = true;
      let t = opener.node.next;
      while (t) { const nx = t.next; node.append(t); t = nx; }
      block.append(node);
      this.processEmphasis(opener.previousDelimiter);
      this.brackets = opener.previous;
      opener.node.unlink();
      if (!isImage) for (let o = this.brackets; o; o = o.previous) if (!o.image) o.active = false;
      return true;
    }
    this.brackets = opener.previous;
    this.pos = start;
    block.append(text(']'));
    return true;
  };
  IP.parseAutolink = function (block) {
    let m;
    if ((m = this.match(/<([a-zA-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*)>/y))) {
      const d = m.slice(1, -1);
      const n = new Node('link');
      n.dest = normalizeURI('mailto:' + d);
      n.title = '';
      n.auto = true;
      n.append(text(d));
      block.append(n);
      return true;
    }
    if ((m = this.match(/<[A-Za-z][A-Za-z0-9.+-]{1,31}:[^<>\x00-\x20]*>/y))) {
      const d = m.slice(1, -1);
      const n = new Node('link');
      n.dest = normalizeURI(d);
      n.title = '';
      n.auto = true;
      n.append(text(d));
      block.append(n);
      return true;
    }
    return false;
  };
  IP.parseHtmlTag = function (block) {
    const m = this.match(reHtmlTag);
    if (m === null) return false;
    const n = new Node('html_inline');
    n.literal = m;
    block.append(n);
    return true;
  };
  IP.parseEntity = function (block) {
    const m = this.match(reEntityHere);
    if (m === null) return false;
    const v = decodeEntity(m);
    block.append(text(v == null ? m : v));
    return true;
  };

  /* ---------- after parsing ---------- */
  function mergeText(n) {
    for (let c = n.first; c; c = c.next) {
      if (c.type === 'text') { while (c.next && c.next.type === 'text') { c.literal += c.next.literal; c.next.unlink(); } if (!c.literal && c.prev) { /* keep */ } }
      else if (c.first) mergeText(c);
    }
  }
  /** GitHub's extended autolinks: www., http(s)://, e-mail addresses and mailto:/xmpp: in plain text */
  function autolinks(n) {
    for (let c = n.first; c; c = c.next) {
      if (c.type === 'link' || c.type === 'image') continue;
      if (c.type !== 'text') { if (c.first) autolinks(c); continue; }
      const s = c.literal;
      const hit = findAutolink(s);
      if (!hit) continue;
      const before = s.slice(0, hit.start), after = s.slice(hit.end);
      const link = new Node('link');
      link.dest = normalizeURI(hit.url);
      link.title = '';
      link.auto = true;
      link.append(text(s.slice(hit.start, hit.end)));
      c.literal = before;
      c.insertAfter(link);
      if (after) link.insertAfter(text(after));
      if (!before) { const nx = c.next; c.unlink(); c = nx; } else c = link;
    }
  }
  const validDomain = (d, www) => {
    const parts = d.split('.');
    if ((www && parts.length < 2) || parts.some((p) => !p)) return false;
    return !parts.slice(-2).some((p) => p.includes('_'));
  };
  function trimTrail(s) {
    for (;;) {
      const last = s[s.length - 1];
      if (/[?!.,:*_~'"]/.test(last)) { s = s.slice(0, -1); continue; }
      if (last === ')') {
        const open = (s.match(/\(/g) || []).length, close = (s.match(/\)/g) || []).length;
        if (close > open) { s = s.slice(0, -1); continue; }
      }
      if (last === ';') { const m = /&[A-Za-z0-9]+;$/.exec(s); if (m) { s = s.slice(0, m.index); continue; } }
      return s;
    }
  }
  function findAutolink(s) {
    const re = /(?:https?:\/\/|ftp:\/\/|www\.)|@|(?:mailto|xmpp):/gi;
    let m;
    while ((m = re.exec(s))) {
      const i = m.index;
      if (m[0] === '@') {
        let a = i;
        while (a > 0 && /[A-Za-z0-9._+-]/.test(s[a - 1])) a--;
        let b = i + 1;
        while (b < s.length && /[A-Za-z0-9._-]/.test(s[b])) b++;
        let domain = s.slice(i + 1, b);
        if (a === i || !domain) continue;
        let lead = a;
        const pre = s.slice(0, a);
        let pm = /(mailto|xmpp):$/i.exec(pre);
        if (pm) lead = a - pm[0].length;
        if (pm && lead > 0 && /[A-Za-z0-9]/.test(s[lead - 1])) { lead = a; pm = null; }
        if (lead > 0 && /[A-Za-z0-9]/.test(s[lead - 1])) continue;
        while (domain.endsWith('.')) { domain = domain.slice(0, -1); b--; }
        if (!domain.includes('.') || /[-_]$/.test(domain)) continue;
        if (pm && /^xmpp$/i.test(pm[1]) && s[b] === '/') {
          let k = b + 1;
          while (k < s.length && /[A-Za-z0-9@.]/.test(s[k])) k++;
          const res = s.slice(b + 1, k).replace(/\.+$/, '');
          if (res) b += 1 + res.length;
        }
        const txt = s.slice(lead, b);
        return { start: lead, end: b, url: pm ? txt : 'mailto:' + txt };
      }
      if (/^(mailto|xmpp):$/i.test(m[0])) continue; /* handled at the @ */
      if (i > 0 && !/[\s*_~("']/.test(s[i - 1])) continue;
      const www = /^www\./i.test(m[0]);
      let j = i + m[0].length;
      let k = j;
      while (k < s.length && /[^\s<>/?#:]/.test(s[k])) k++;
      if (!validDomain((www ? s.slice(i, k) : s.slice(j, k)).replace(/\.+$/, ''), www)) continue;
      while (k < s.length && !/[\s<]/.test(s[k])) k++;
      const linkText = trimTrail(s.slice(i, k));
      if (!linkText || linkText.length <= m[0].length) continue;
      return { start: i, end: i + linkText.length, url: (www ? 'http://' : '') + linkText };
    }
    return null;
  }

  /* ================= HTML output ================= */
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  M.escapeHTML = esc;
  const TAGFILTER = /<(?=\/?(?:title|textarea|style|xmp|iframe|noembed|noframes|script|plaintext)(?:[\s/>]|$))/gi;
  /** plain text of a node's inlines (image alt text, heading anchors) */
  M.plain = function plain(n) {
    let s = '';
    for (let c = n.first; c; c = c.next) {
      if (c.type === 'text' || c.type === 'code' || c.type === 'math') s += c.literal;
      else if (c.type === 'softbreak' || c.type === 'linebreak') s += '\n';
      else if (c.first) s += plain(c);
    }
    return s;
  };
  /** opts: { tagfilter: true } */
  M.toHTML = function (doc, opts) {
    opts = Object.assign({ tagfilter: true }, opts || {});
    const out = [];
    let lastc = '\n';
    const put = (str) => { if (str) { out.push(str); lastc = str[str.length - 1]; } };
    const cr = () => { if (lastc !== '\n') put('\n'); };
    const raw = (s) => (opts.tagfilter ? s.replace(TAGFILTER, '&lt;') : s);
    const fnIndex = new Map((doc.footnoteOrder || []).map((d, i) => [d, i + 1]));
    const inl = (n) => { for (let c = n.first; c; c = c.next) inline(c); };
    function inline(n) {
      switch (n.type) {
        case 'text': put(esc(n.literal)); break;
        case 'softbreak': put('\n'); break;
        case 'linebreak': put('<br />\n'); break;
        case 'emph': put('<em>'); inl(n); put('</em>'); break;
        case 'strong': put('<strong>'); inl(n); put('</strong>'); break;
        case 'del': put('<del>'); inl(n); put('</del>'); break;
        case 'code': put('<code>' + esc(n.literal) + '</code>'); break;
        case 'math': put('<span class="math">' + esc(n.literal) + '</span>'); break;
        case 'html_inline': put(raw(n.literal)); break;
        case 'link': put(`<a href="${esc(n.dest)}"${n.title ? ` title="${esc(n.title)}"` : ''}>`); inl(n); put('</a>'); break;
        case 'image': put(`<img src="${esc(n.dest)}" alt="${esc(M.plain(n))}"${n.title ? ` title="${esc(n.title)}"` : ''} />`); break;
        case 'footnote_ref': { const i = fnIndex.get(n.def); put(`<sup class="footnote-ref"><a href="#fn-${esc(n.label)}" id="fnref-${esc(n.label)}" data-footnote-ref>${i}</a></sup>`); break; }
        default: inl(n);
      }
    }
    function block(n, tight) {
      switch (n.type) {
        case 'document': for (let c = n.first; c; c = c.next) block(c, false); break;
        case 'paragraph':
          if (tight) { inl(n); break; }
          cr(); put('<p>'); inl(n); put('</p>'); cr();
          break;
        case 'heading': cr(); put(`<h${n.level}>`); inl(n); put(`</h${n.level}>`); cr(); break;
        case 'code_block': {
          const lang = n.info ? n.info.split(/[ \t]+/)[0] : '';
          cr(); put(`<pre><code${lang ? ` class="language-${esc(lang)}"` : ''}>${esc(n.literal)}</code></pre>`); cr();
          break;
        }
        case 'html_block': cr(); put(raw(n.literal)); cr(); break;
        case 'thematic_break': cr(); put('<hr />'); cr(); break;
        case 'block_quote':
          cr(); put(n.alert ? `<div class="markdown-alert markdown-alert-${n.alert}">` : '<blockquote>'); cr();
          for (let c = n.first; c; c = c.next) block(c, false);
          cr(); put(n.alert ? '</div>' : '</blockquote>'); cr();
          break;
        case 'list': {
          const ol = n.data.type === 'ordered';
          cr(); put(ol ? (n.data.start !== 1 ? `<ol start="${n.data.start}">` : '<ol>') : '<ul>'); cr();
          for (let c = n.first; c; c = c.next) block(c, n.tight);
          cr(); put(ol ? '</ol>' : '</ul>'); cr();
          break;
        }
        case 'item':
          put('<li>');
          if (n.task) put(n.task.checked ? '<input checked="" disabled="" type="checkbox"> ' : '<input disabled="" type="checkbox"> ');
          for (let c = n.first; c; c = c.next) block(c, tight);
          put('</li>'); cr();
          break;
        case 'table': {
          cr(); put('<table>\n<thead>\n');
          let body = false;
          for (let r = n.first; r; r = r.next) {
            if (!r.head && !body) { put('<tbody>\n'); body = true; }
            put('<tr>\n');
            for (let c = r.first; c; c = c.next) { const tag = r.head ? 'th' : 'td'; put(`<${tag}${c.align ? ` align="${c.align}"` : ''}>`); inl(c); put(`</${tag}>\n`); }
            put('</tr>\n');
            if (r.head) put('</thead>\n');
          }
          if (body) put('</tbody>\n');
          put('</table>\n');
          break;
        }
        case 'footnote_def': break;
        default: for (let c = n.first; c; c = c.next) block(c, tight);
      }
    }
    block(doc, false);
    const defs = doc.footnoteOrder || [];
    if (defs.length) {
      cr(); put('<section class="footnotes" data-footnotes>\n<ol>\n');
      defs.forEach((d) => {
        put(`<li id="fn-${esc(d.label)}">\n`);
        for (let c = d.first; c; c = c.next) block(c, false);
        cr(); put('</li>\n');
      });
      put('</ol>\n</section>\n');
    }
    return out.join('');
  };

  /** GitHub's heading anchors: lower case, punctuation dropped, spaces to hyphens, repeats numbered */
  M.slug = function (s, used) {
    let base = String(s).trim().toLowerCase().replace(/[^\p{L}\p{M}\p{N}\p{Pc} -]/gu, '').replace(/ /g, '-');
    if (!used) return base;
    let id = base, i = 0;
    while (used.has(id)) id = base + '-' + ++i;
    used.add(id);
    return id;
  };
})(typeof window !== 'undefined' ? window : globalThis);
