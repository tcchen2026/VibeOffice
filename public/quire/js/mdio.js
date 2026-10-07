/* Quire — Markdown files: opening .md (or a .zip holding a .md and its pictures) as a document, and saving
 * a document as GitHub-flavoured Markdown. The reader (markdown.js) gives a block tree with source lines;
 * every top-level block remembers its lines, and when the document is saved, blocks that were not edited
 * are written back exactly as they were in the file. Edited and new blocks are written in one style:
 * ATX headings, "-" bullets, ** and *, fenced code, pipe tables (HTML tables when cells are merged).
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, MD = L.md;
  const X = (L.mdio = {});

  /* ================= the look of a Markdown document ================= */
  const BORDER = (color, sz, space) => ({ val: 'single', sz, color, space });
  const ALERTS = { note: ['Note', '0969DA'], tip: ['Tip', '1A7F37'], important: ['Important', '8250DF'], warning: ['Warning', '9A6700'], caution: ['Caution', 'CF222E'] };
  const QUOTE_STEP = 14;
  const CODE_STYLES = new Set(['HTMLPreformatted', 'PlainText', 'MacroText']);
  const QUOTE_STYLES = new Set(['Quote', 'IntenseQuote', 'BlockText']);
  const levelInd = (lvl) => 36 + lvl * 36;

  function setupStyles(d) {
    const st = d.styles, b = D.builtinStyles();
    const font = 'Calibri';
    [[20, true], [16, true], [13], [11], [10.5], [10]].forEach(([sz, rule], i) => {
      const id = 'Heading' + (i + 1);
      const s = (st[id] = st[id] || b[id]);
      s.rPr = Object.assign({ font, b: true, sz }, i === 5 ? { color: '59636E' } : {});
      s.pPr = Object.assign({ keepNext: true, sp: { b: 18, a: 8 }, outline: i }, rule ? { borders: { bottom: BORDER('D8DEE4', 0.75, 4) } } : {});
    });
    st.Quote.pPr = { ind: { l: QUOTE_STEP }, borders: { left: BORDER('D0D7DE', 3, 10) } };
    st.Quote.rPr = { color: '59636E' };
    st.HTMLPreformatted.pPr = { sp: { b: 8, a: 8, line: 1, rule: 'auto' }, contextual: true, shd: { val: 'clear', fill: 'F6F8FA', color: 'auto' } };
    st.HTMLPreformatted.rPr = { font: 'Courier New', sz: 10 };
    st.ListParagraph.pPr = { ind: { l: 36 }, contextual: true };
    st.Hyperlink.rPr = { color: '0969DA', u: 'single' };
    const ch = (id, name, rPr, o) => (st[id] = Object.assign({ id, name, type: 'character', basedOn: null, next: null, pPr: {}, rPr, builtin: true }, o || {}));
    ch('HTMLCode', 'HTML Code', { font: 'Courier New', sz: 10, shd: { val: 'clear', fill: 'EFF1F3', color: 'auto' } });
    ch('HTMLKeyboard', 'HTML Keyboard', { font: 'Courier New', sz: 10, shd: { val: 'clear', fill: 'F6F8FA', color: 'auto' } });
    ch('MarkdownSource', 'Markdown Source', { font: 'Courier New', sz: 10, color: '6E7781' }, { builtin: false, custom: true });
    const line = BORDER('D0D7DE', 0.5, 0);
    st.MarkdownTable = { id: 'MarkdownTable', name: 'Markdown Table', type: 'table', basedOn: 'TableNormal', next: null, custom: true, pPr: {}, rPr: {},
      tblPr: { borders: { top: line, left: line, bottom: line, right: line, insideH: line, insideV: line }, cellMar: { l: 9, r: 9, t: 4, b: 4 } },
      cond: { firstRow: { rPr: { b: true } }, band2Horz: { tcPr: { shd: { val: 'clear', fill: 'F6F8FA', color: 'auto' } } } } };
    D.stylesChanged();
  }

  /* ================= reading ================= */
  /** a picture that could not be loaded: a box with its alt text (the file keeps the reference) */
  function placeholder(label) {
    const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;');
    const w = Math.min(420, Math.max(120, label.length * 7 + 40)), h = 40;
    const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><rect x="0.5" y="0.5" width="${w - 1}" height="${h - 1}" fill="#F6F8FA" stroke="#D0D7DE"/><text x="${w / 2}" y="25" font-family="Arial,sans-serif" font-size="12" fill="#57606A" text-anchor="middle">${esc(label)}</text></svg>`;
    return new Blob([svg], { type: 'image/svg+xml' });
  }
  const isFrontMatter = (s) => s.split('\n').slice(1, -1).every((ln) => !ln.trim() || /^(#|\s|-\s|[\w"'.-]+\s*:)/.test(ln));

  /** Markdown text → document. opts.resolve(path) → Blob|null finds pictures (inside a .zip) */
  X.read = async function (src, opts) {
    opts = opts || {};
    let text = String(src);
    const bom = text.charCodeAt(0) === 0xfeff;
    if (bom) text = text.slice(1);
    const eol = /\r\n/.test(text) ? '\r\n' : '\n';
    text = text.replace(/\r\n?/g, '\n');
    const finalNewline = text.endsWith('\n') || !text;
    let front = null;
    const fm = /^---[ \t]*\n[\s\S]*?\n(?:---|\.\.\.)[ \t]*(?:\n|$)/.exec(text);
    if (fm && isFrontMatter(fm[0].replace(/\n$/, ''))) { front = fm[0].replace(/\n$/, '').split('\n'); text = text.slice(fm[0].length); }
    const ast = MD.parse(text);
    const lines = text.split('\n');
    if (text.endsWith('\n')) lines.pop();

    const d = D.newDoc({ modern: true });
    d.main.blocks = [];
    setupStyles(d);
    const prevDoc = D.doc;
    D.doc = d;
    try {
      const cx = newContext(d, opts);
      const groups = [], lead = [];
      let cur = null, prevEnd = 0;
      const pieces = () => (cur ? cur.after : lead);
      for (let n = ast.first; n; n = n.next) {
        if (n.sl > prevEnd + 1) pieces().push({ raw: lines.slice(prevEnd, n.sl - 1) });
        /* link reference definitions in front of a paragraph stay with the text before it */
        const sl = n.sl + (n.refLines || 0);
        if (sl > n.sl) pieces().push({ raw: lines.slice(n.sl - 1, sl - 1) });
        const own = lines.slice(sl - 1, n.el);
        if (n.type === 'footnote_def') pieces().push({ fn: n, raw: own });
        else {
          const blocks = await convert(n, cx, ROOT);
          if (!blocks.length) pieces().push({ raw: own });
          else {
            cur = { i: groups.length, body: own, after: [], n: blocks.length };
            groups.push(cur);
            for (const b of blocks) { b.md = { g: cur.i }; d.main.blocks.push(b); }
          }
        }
        prevEnd = Math.max(prevEnd, n.el);
      }
      if (lines.length > prevEnd) pieces().push({ raw: lines.slice(prevEnd) });
      await buildFootnotes(cx);
      /* Word needs a paragraph after a table, and between two tables */
      const bl = d.main.blocks;
      for (let i = bl.length - 1; i >= 0; i--) if (bl[i].t === 'tbl' && (!bl[i + 1] || bl[i + 1].t === 'tbl')) bl.splice(i + 1, 0, D.para());
      if (!bl.length) bl.push(D.para());
      for (const b of bl) decorate(b);
      for (const id in d.fn) for (const b of d.fn[id].blocks) decorate(b);
      d.md = { front, lead, groups, eol, bom, finalNewline };
      /* what each group and footnote looks like when written: unchanged blocks are written as they were */
      const w = writerFor(d, { check: true });
      for (const g of groups) g.norm = w.blocks(bl.filter((b) => b.md && b.md.g === g.i)).join('\n');
      for (const g of [{ after: lead }, ...groups]) for (const pc of g.after) if (pc.fn && cx.fnIds.has(pc.fn)) { pc.id = cx.fnIds.get(pc.fn); pc.norm = w.footnote(pc.id).join('\n'); }
    } finally { D.doc = prevDoc; D.stylesChanged(); }
    return d;
  };

  const ROOT = { q: 0, alert: null, lvl: 0, cont: 0, tight: true };
  function newContext(d, opts) {
    const bulletAbs = D.makeBulletAbs('•', 'Symbol', '');
    D.addNum(d, bulletAbs);
    const ordAbs = D.makeNumberAbs('decimal', '%1.');
    ordAbs.levels.forEach((lv, i) => { lv.fmt = ['decimal', 'lowerRoman', 'lowerLetter'][i % 3]; lv.text = `%${i + 1}.`; if (lv.fmt === 'lowerRoman') { lv.jc = 'left'; lv.ind = { l: levelInd(i), fl: -18 }; } });
    D.addNum(d, ordAbs);
    return { d, opts, bulletAbs: bulletAbs.id, ordAbs: ordAbs.id, fnIds: new Map(), fnQueue: [], fnNext: 1, media: new Map(), bm: 0 };
  }
  function newNum(cx, ordered, lvl, start) {
    const nb = cx.d.numbering;
    let nid = 1;
    while (nb.nums[nid] != null) nid++;
    nb.nums[nid] = { abs: ordered ? cx.ordAbs : cx.bulletAbs, ov: ordered ? { [lvl]: { start: start == null ? 1 : start } } : {} };
    return String(nid);
  }
  /** paragraph properties for a block at this nesting */
  function pPrFor(c, extra) {
    const p = Object.assign({}, extra || {});
    if (c.q) { p.mdq = c.q; if (c.alert) p.mdAlert = c.alert; if (c.qIn) p.mdqIn = c.qIn; }
    if (c.cont) { p.mdli = c.cont; if (!p.style) p.style = 'ListParagraph'; if (c.tight === false) p.contextual = false; }
    else if (c.q && !p.style) p.style = 'Quote';
    return p;
  }
  /** indents and quote bars from the nesting recorded in mdq / mdli / num */
  function decorate(b) {
    if (b.t === 'tbl') { for (const r of b.rows) for (const c of r.cells) for (const x of c.blocks) decorate(x); return; }
    const p = b.pPr, q = p.mdq || 0;
    const l = p.num ? levelInd(p.num.lvl) : p.mdli ? levelInd(p.mdli - 1) : null;
    if (p.mdli && !p.num) p.ind = { l };
    if (!q) return;
    if (p.style !== 'Quote' || q > 1 || l != null || p.mdAlert) {
      p.ind = Object.assign({}, p.num ? { fl: -18 } : {}, { l: (l || 0) + QUOTE_STEP * q });
      p.borders = { left: BORDER(p.mdAlert ? ALERTS[p.mdAlert][1] : 'D0D7DE', 3, 10) };
    }
  }

  async function convertKids(n, cx, c) {
    const out = [];
    for (let k = n.first; k; k = k.next) out.push(...(await convert(k, cx, c)));
    return out;
  }
  async function convert(n, cx, c) {
    switch (n.type) {
      case 'paragraph': return [D.para(await inlines(n, cx), pPrFor(c))];
      case 'heading': return [D.para(await inlines(n, cx), pPrFor(c, { style: 'Heading' + n.level }))];
      case 'thematic_break': return [D.para([], pPrFor(c, { borders: { bottom: BORDER('D0D7DE', 1.5, 1) } }))];
      case 'code_block': {
        const src = n.literal.replace(/\n$/, '');
        return src.split('\n').map((ln) => {
          const runs = [];
          ln.split('\t').forEach((t, i) => { if (i) runs.push(D.item('tab')); if (t) runs.push(D.text(t)); });
          return D.para(runs, pPrFor(c, { style: 'HTMLPreformatted', mdInfo: n.info || '' }));
        });
      }
      case 'html_block': {
        if (/^\s*(<!--[\s\S]*?-->\s*)+$/.test(n.literal) || !L.htmlio) return [];
        const blocks = await L.htmlio.htmlToBlocks(await resolveHtmlImages(n.literal, cx), { clipboard: false });
        const out = blocks.filter((b) => !(b.t === 'p' && !b.runs.length && blocks.length === 1));
        for (const b of out) { delete b.partial; if (b.t === 'p') b.pPr = Object.assign(pPrFor(c), b.pPr); }
        return out;
      }
      case 'block_quote': {
        const inner = Object.assign({}, c, { q: c.q + 1, alert: n.alert || c.alert, qIn: c.q ? c.qIn : c.cont });
        const out = [];
        if (n.alert) { const [title, color] = ALERTS[n.alert]; out.push(D.para([D.text(title, { b: true, color })], pPrFor(inner, { mdAlertTitle: true }))); }
        out.push(...(await convertKids(n, cx, inner)));
        if (!out.length) out.push(D.para([], pPrFor(inner)));
        return out;
      }
      case 'list': {
        const ordered = n.data.type === 'ordered';
        const lvl = Math.min(8, c.lvl);
        const numId = newNum(cx, ordered, lvl, n.data.start);
        const inner = Object.assign({}, c, { lvl: c.lvl + 1, cont: c.lvl + 1, tight: n.tight });
        const out = [];
        for (let it = n.first; it; it = it.next) {
          const kids = await convertKids(it, cx, inner);
          let first = kids[0];
          if (!first || first.t !== 'p' || first.pPr.num || CODE_STYLES.has(first.pPr.style)) { first = D.para([], pPrFor(inner)); kids.unshift(first); }
          first.pPr.num = { id: numId, lvl };
          delete first.pPr.mdli;
          if (!first.pPr.style || first.pPr.style === 'Quote') first.pPr.style = 'ListParagraph';
          if (it.task) first.runs.unshift(...checkbox(it.task.checked), D.text(' '));
          if (!n.tight) first.pPr.mdLoose = true;
          out.push(...kids);
        }
        /* space after the list (its paragraphs have none between them) */
        const last = out[out.length - 1];
        if (c.lvl === 0 && last && last.t === 'p' && last.pPr.style === 'ListParagraph') last.pPr.contextual = false;
        return out;
      }
      case 'table': {
        const rows = n.kids, ncol = n.align.length;
        const sect = cx.d.sect;
        const avail = sect.pgW - sect.ml - sect.mr - QUOTE_STEP * c.q - (c.cont ? levelInd(c.cont - 1) : 0);
        const wts = new Array(ncol).fill(3);
        for (const r of rows) r.kids.forEach((cell, k) => { wts[k] = Math.max(wts[k], Math.min(40, MD.plain(cell).length)); });
        const sum = wts.reduce((a, b) => a + b, 0);
        const grid = wts.map((x) => L.round((avail * x) / sum, 2));
        const tbl = D.table([], grid, { style: 'MarkdownTable', w: { type: 'auto', v: 0 }, look: { firstRow: true, noVBand: true } });
        if (c.q) tbl.tblPr.mdq = c.q;
        if (c.cont) tbl.tblPr.mdli = c.cont;
        for (const r of rows) {
          const cells = [];
          let k = 0;
          for (const cell of r.kids) {
            const jc = { left: 'left', center: 'center', right: 'right' }[cell.align];
            cells.push(D.cell([D.para(await inlines(cell, cx), Object.assign({ sp: { b: 0, a: 0 } }, jc ? { jc } : {}))], { w: grid[k++] }));
          }
          tbl.rows.push(D.row(cells, r.head ? { header: true } : {}));
        }
        return [tbl];
      }
      default: return [];
    }
  }
  const checkbox = (on) => { const fid = D.newFid(); return [D.item('fb', { fid, instr: ' FORMCHECKBOX ', ff: { type: 'checkbox', name: '', def: false, checked: !!on } }), D.item('fe', { fid })]; };

  /** inline nodes → paragraph items */
  async function inlines(n, cx) {
    const runs = [];
    const html = { b: 0, i: 0, s: 0, u: 0, sup: 0, sub: 0, kbd: 0, code: 0, link: null };
    const rp = (f) => {
      const r = {};
      if (f.b || html.b) r.b = true;
      if (f.i || html.i) r.i = true;
      if (f.s || html.s) r.strike = true;
      if (html.u) r.u = 'single';
      if (html.sup) r.vert = 'superscript'; else if (html.sub) r.vert = 'subscript';
      const link = f.link || html.link;
      if (link) { r.link = link; r.style = 'Hyperlink'; }
      if (f.style) r.style = f.style;
      else if (html.code) r.style = 'HTMLCode';
      else if (html.kbd) r.style = 'HTMLKeyboard';
      return r;
    };
    const linkOf = (dest, title) => Object.assign(dest[0] === '#' ? { anchor: decodeURIComponentSafe(dest.slice(1)) } : { url: dest }, title ? { tip: title } : {});
    async function walk(node, f) {
      for (let k = node.first; k; k = k.next) {
        switch (k.type) {
          case 'text': if (k.literal) runs.push(D.text(k.literal, rp(f))); break;
          case 'softbreak': runs.push(D.text(' ', rp(f))); break;
          case 'linebreak': runs.push(D.item('br', { type: 'line' }, rp(f))); break;
          case 'emph': await walk(k, Object.assign({}, f, { i: true })); break;
          case 'strong': await walk(k, Object.assign({}, f, { b: true })); break;
          case 'del': await walk(k, Object.assign({}, f, { s: true })); break;
          case 'code': runs.push(D.text(k.literal, rp(Object.assign({}, f, { style: 'HTMLCode' })))); break;
          case 'math': runs.push(D.text(k.literal, rp(Object.assign({}, f, { style: 'MarkdownSource' })))); break;
          case 'link': await walk(k, Object.assign({}, f, { link: linkOf(k.dest, k.title) })); break;
          case 'image': runs.push(await picture(cx, k.dest, MD.plain(k), k.title, null, null, rp(f))); break;
          case 'footnote_ref': runs.push(footRef(cx, k)); break;
          case 'html_inline': await htmlInline(k.literal, f); break;
          default: if (k.first) await walk(k, f);
        }
      }
    }
    async function htmlInline(s, f) {
      const m = /^<(\/?)([A-Za-z][A-Za-z0-9-]*)([^>]*)>$/.exec(s);
      const tag = m && m[2].toLowerCase(), close = m && m[1] === '/';
      const attrs = m ? parseAttrs(m[3]) : {};
      const key = { b: 'b', strong: 'b', i: 'i', em: 'i', s: 's', del: 's', strike: 's', u: 'u', ins: 'u', sup: 'sup', sub: 'sub', kbd: 'kbd', code: 'code', tt: 'code' }[tag];
      if (key) { html[key] = Math.max(0, html[key] + (close ? -1 : 1)); return; }
      if (tag === 'br' && !close) { runs.push(D.item('br', { type: 'line' }, rp(f))); return; }
      if (tag === 'img' && !close && attrs.src) { runs.push(await picture(cx, attrs.src, attrs.alt || '', attrs.title || '', parseFloat(attrs.width) || null, parseFloat(attrs.height) || null, rp(f), true)); return; }
      if (tag === 'a' && attrs.href && !close) { html.link = linkOf(attrs.href, attrs.title); return; }
      if (tag === 'a' && close && html.link) { html.link = null; return; }
      /* <a name="x"></a> / <a id="x"></a>: a bookmark, so links to #x work */
      if (tag === 'a' && !close && (attrs.name || attrs.id) && !attrs.href) { const id = 'md' + ++cx.bm; runs.push(D.item('bs', { id, name: attrs.name || attrs.id, mdAttr: attrs.name ? 'name' : 'id' }), D.item('be', { id })); html.anchor = (html.anchor || 0) + 1; return; }
      if (tag === 'a' && close && html.anchor) { html.anchor--; return; }
      runs.push(D.text(s, Object.assign(rp(f), { style: 'MarkdownSource' })));
    }
    await walk(n, {});
    return runs;
  }
  function decodeURIComponentSafe(s) { try { return decodeURIComponent(s); } catch (e) { return s; } }
  function parseAttrs(s) {
    const o = {};
    for (const m of s.matchAll(/([a-zA-Z_:][\w:.-]*)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g)) o[m[1].toLowerCase()] = MD.unescape(m[2] ?? m[3] ?? m[4] ?? '');
    return o;
  }
  /** a picture: data: URLs and (in a .zip) relative paths load, anything else shows a placeholder */
  async function picture(cx, src, alt, title, wPx, hPx, rPr, fromHtml) {
    let blob = null;
    try {
      if (/^data:/i.test(src)) blob = await (await fetch(src)).blob();
      else if (!/^[a-z][a-z0-9+.-]*:|^\/\//i.test(src) && cx.opts.resolve) blob = await cx.opts.resolve(decodeURIComponentSafe(src.replace(/[?#].*$/, '')));
    } catch (e) { blob = null; }
    const missing = !blob;
    if (!blob) blob = placeholder(alt || src.split('/').pop() || 'picture');
    const name = src.split(/[/?#]/).filter(Boolean).pop() || 'image.png';
    const id = L.media.add(blob, /^data:/i.test(src) ? 'image.' + L.mimeToExt(blob.type) : name);
    let nw = 96, nh = 96;
    try { const im = await L.loadImage(L.media.url(id)); nw = im.naturalWidth || nw; nh = im.naturalHeight || nh; } catch (e) { /* keep the default size */ }
    let w = (wPx || (hPx ? (nw * hPx) / nh : nw)) * 0.75, h = (hPx || (wPx ? (nh * wPx) / nw : nh)) * 0.75;
    const sect = cx.d.sect, max = sect.pgW - sect.ml - sect.mr;
    if (w > max) { h *= max / w; w = max; }
    const it = D.item('img', { media: id, w: L.round(w, 2), h: L.round(h, 2), alt: alt || '' }, Object.assign({}, rPr, { style: undefined }));
    if (!it.rPr.style) delete it.rPr.style;
    it.md = { src, title: title || '', media: id, w: it.w, h: it.h, html: !!fromHtml && !!(wPx || hPx), missing };
    return it;
  }
  async function resolveHtmlImages(html, cx) {
    if (!cx.opts.resolve) return html;
    const subs = [];
    for (const m of html.matchAll(/(<img\b[^>]*?\bsrc\s*=\s*)("([^"]*)"|'([^']*)')/gi)) {
      const u = m[3] ?? m[4];
      if (/^[a-z][a-z0-9+.-]*:|^\/\//i.test(u)) continue;
      const blob = await cx.opts.resolve(decodeURIComponentSafe(u.replace(/[?#].*$/, '')));
      if (blob) subs.push([m[0], m[1] + '"' + URL.createObjectURL(blob) + '"']);
    }
    for (const [a, b] of subs) html = html.split(a).join(b);
    return html;
  }
  /** footnotes become real footnotes; a second reference to the same note stays as source */
  function footRef(cx, k) {
    if (cx.fnIds.has(k.def)) return D.text('[^' + k.label + ']', { style: 'MarkdownSource' });
    const id = String(cx.fnNext++);
    cx.fnIds.set(k.def, id);
    cx.fnQueue.push(k.def);
    return D.item('fn', { id, mdLabel: k.def.label }, { style: 'FootnoteReference' });
  }
  async function buildFootnotes(cx) {
    while (cx.fnQueue.length) {
      const def = cx.fnQueue.shift();
      const id = cx.fnIds.get(def);
      const blocks = await convertKids(def, cx, ROOT);
      let first = blocks[0];
      if (!first || first.t !== 'p') { first = D.para(); blocks.unshift(first); }
      first.runs.unshift(D.item('fn', { id, self: true }, { style: 'FootnoteReference' }), D.text(' '));
      for (const b of blocks) if (b.t === 'p' && !b.pPr.style) b.pPr.style = 'FootnoteText';
      cx.d.fn[id] = { kind: 'fn', id, blocks };
    }
  }

  /* ================= writing ================= */
  /** document → { text, images: [{ path, media }], dropped } (dropped: pictures a plain .md cannot hold) */
  X.write = function (d, opts) {
    opts = opts || {};
    const cacheReset = d !== D.doc;
    if (cacheReset) D.stylesChanged();
    const prevDoc = D.doc;
    D.doc = d;
    try { return writeDoc(d, opts); } finally { D.doc = prevDoc; if (cacheReset) D.stylesChanged(); }
  };
  function writeDoc(d, opts) {
    const w = writerFor(d, opts), wc = writerFor(d, { check: true });
    const md = d.md || { front: null, lead: [], groups: [], eol: '\n', bom: false, finalNewline: true };
    const blocks = d.main.blocks;
    const empty = (b) => b.t === 'p' && wc.info(b).kind === 'empty';

    /* runs of blocks from one source block; a run is kept as it was when it is complete, unchanged,
       and nothing next to it would merge with it (a new list item, a new line of code) */
    const runs = [];
    for (let i = 0; i < blocks.length;) {
      const g = blocks[i].md ? blocks[i].md.g : null;
      let j = i + 1;
      if (g != null && md.groups[g]) while (j < blocks.length && blocks[j].md && blocks[j].md.g === g) j++;
      runs.push({ g: g != null && md.groups[g] ? g : null, from: i, to: j });
      i = j;
    }
    const count = new Map();
    for (const r of runs) if (r.g != null) count.set(r.g, (count.get(r.g) || 0) + 1);
    for (const r of runs) {
      if (r.g == null) continue;
      const grp = md.groups[r.g];
      r.keep = count.get(r.g) === 1 && r.to - r.from === grp.n && wc.blocks(blocks.slice(r.from, r.to)).join('\n') === grp.norm;
    }
    const neighbour = (i, step) => { for (let k = i; k >= 0 && k < blocks.length; k += step) if (!empty(blocks[k])) return blocks[k]; return null; };
    const isNew = (b) => b && !(b.md && md.groups[b.md.g]);
    for (const r of runs) if (r.keep) {
      let before = neighbour(r.from - 1, -1), after = neighbour(r.to, 1);
      if (!isNew(before)) before = null;
      if (!isNew(after)) after = null;
      if ((before && wc.joins(before, blocks[r.from])) || (after && wc.joins(blocks[r.to - 1], after))) r.keep = false;
    }

    const out = [];
    const isBlank = (s) => !s || !s.trim();
    const ensureGap = (next) => { if (out.length && !isBlank(out[out.length - 1]) && next.length && !isBlank(next[0])) out.push(''); };
    const emitted = new Set();
    const pieceLines = (pc) => {
      if (!pc.fn) return pc.raw;
      if (pc.id == null) return pc.raw;
      if (!d.fn[pc.id] || !w.referenced(pc.id) || emitted.has(pc.id)) return [];
      emitted.add(pc.id);
      return wc.footnote(pc.id).join('\n') === pc.norm ? pc.raw : w.footnote(pc.id);
    };
    const emitPieces = (list) => { for (const pc of list) out.push(...pieceLines(pc)); };
    if (md.front) out.push(...md.front);
    emitPieces(md.lead);
    let lastKept = -1; /* group index of the previous unit when it was written as it was (the lead is always) */
    let stretch = [], stretchGroups = [];
    const flush = () => {
      if (!stretch.length) return;
      const lines = w.blocks(stretch);
      if (lines.length) { ensureGap(lines); out.push(...lines); }
      for (const g of stretchGroups) for (const pc of md.groups[g].after) {
        const ls = pieceLines(pc).slice();
        while (ls.length && isBlank(ls[0])) ls.shift();
        while (ls.length && isBlank(ls[ls.length - 1])) ls.pop();
        if (ls.length) { ensureGap(ls); out.push(...ls); }
      }
      stretch = []; stretchGroups = [];
      if (lines.length) lastKept = null;
    };
    for (const r of runs) {
      if (r.keep) {
        flush();
        const grp = md.groups[r.g];
        if (lastKept !== r.g - 1) ensureGap(grp.body);
        out.push(...grp.body);
        emitPieces(grp.after);
        lastKept = r.g;
      } else {
        stretch.push(...blocks.slice(r.from, r.to));
        if (r.g != null && !stretchGroups.includes(r.g)) stretchGroups.push(r.g);
      }
    }
    flush();
    /* footnotes written in Quire, or whose definition is not in the file any more */
    for (const id of w.noteOrder()) if (!emitted.has(id)) { const ls = w.footnote(id); ensureGap(ls); out.push(...ls); }
    while (out.length && isBlank(out[out.length - 1]) && !(d.md && md.groups.length)) out.pop();
    let text = out.join(md.eol);
    if (md.finalNewline && text) text += md.eol;
    if (md.bom) text = '﻿' + text;
    return { text, images: w.images, dropped: w.dropped };
  }

  /* ---------- the block and inline writer ---------- */
  function writerFor(d, opts) {
    const labels = D.computeLists(d, Object.values(d.fn));
    const images = [], usedPaths = new Set(opts.paths || []);
    let dropped = 0;
    const charProps = (id) => (id && d.styles[id] && d.styles[id].type === 'character' ? D.charStyleProps(d, id) : {});
    const styleOf = (p) => p.pPr.style || 'Normal';
    const baseStyle = (id) => { const s = d.styles[id]; return s && s.basedOn ? s.basedOn : null; };
    const isCode = (p) => CODE_STYLES.has(styleOf(p)) || CODE_STYLES.has(baseStyle(styleOf(p)));
    const quoteDepth = (p) => p.pPr.mdq || (QUOTE_STYLES.has(styleOf(p)) ? 1 : 0);

    function info(b) {
      if (b.t === 'tbl') return { kind: 'table', q: b.tblPr.mdq || 0, cont: b.tblPr.mdli || 0 };
      const o = { q: quoteDepth(b), alert: b.pPr.mdAlert || null, cont: b.pPr.mdli || 0, qIn: b.pPr.mdqIn || 0 };
      const lb = labels.get(b.id);
      const head = D.headingLevel(d, b);
      if (b.pPr.mdAlertTitle && o.q) return Object.assign(o, { kind: 'alertTitle' });
      if (isCode(b)) return Object.assign(o, { kind: 'code', lang: b.pPr.mdInfo || '' });
      const ordered = lb && lb.lv.fmt !== 'bullet' && lb.lv.fmt !== 'none';
      /* a heading numbered by a list of its own (not outline numbering) is a list item that holds a heading */
      if (head && lb && b.pPr.num && (b.pPr.mdLoose != null || !ordered || /ListParagraph/.test(b.pPr.style || ''))) return Object.assign(o, { kind: 'item', li: lb, ordered, heading: Math.min(6, head) });
      if (head) return Object.assign(o, { kind: 'heading', level: Math.min(6, head), label: lb ? lb.text : '' });
      if (lb) return Object.assign(o, { kind: 'item', li: lb, ordered });
      const hasText = b.runs.some((it) => (it.t === 'text' && it.text.trim()) || ['img', 'sym', 'fn', 'en', 'br', 'tab', 'fb'].includes(it.t));
      if (!hasText) {
        const pp = D.pProps(d, b);
        if (pp.borders && pp.borders.bottom && pp.borders.bottom.val && pp.borders.bottom.val !== 'none' && pp.borders.bottom.val !== 'nil') return Object.assign(o, { kind: 'hr' });
        return Object.assign(o, { kind: 'empty' });
      }
      return Object.assign(o, { kind: 'para' });
    }
    /** do two neighbouring blocks belong to one Markdown block (so one cannot be kept without the other)? */
    function joins(a, b) {
      if (a.t !== 'p' || b.t !== 'p') return false;
      const x = info(a), y = info(b);
      if (x.kind === 'item' && y.kind === 'item') return x.li.numId === y.li.numId || x.li.lvl !== y.li.lvl;
      if ((x.kind === 'item' || x.cont) && (y.kind === 'item' || y.cont)) return true;
      if (x.kind === 'code' && y.kind === 'code') return true;
      return x.q > 0 && y.q > 0;
    }

    /* inline content: tokens with their marks, then Markdown with * ** ~~ and HTML where those cannot open or close */
    function tokens(p) {
      const toks = [];
      const fstack = [];
      for (const it of p.runs) {
        const r = it.rPr || {};
        if (r.del) continue;
        if (it.t === 'fb') {
          const cb = it.ff && it.ff.type === 'checkbox';
          fstack.push(cb ? 'cb' : 'code');
          if (cb) toks.push({ t: 'atom', md: it.ff.checked ? '☒' : '☐', html: it.ff.checked ? '☒' : '☐', m: {}, cb: it.ff.checked });
          continue;
        }
        if (it.t === 'fs') { if (fstack.length) fstack[fstack.length - 1] = 'result'; continue; }
        if (it.t === 'fe') { fstack.pop(); continue; }
        if (fstack.length && fstack[fstack.length - 1] === 'code') continue;
        if (r.hidden) continue;
        const m = marks(r);
        switch (it.t) {
          case 'text': toks.push({ t: 'text', s: it.text, m }); break;
          case 'tab': toks.push({ t: 'text', s: '\t', m }); break;
          case 'br': if (!it.type || it.type === 'line' || it.type === 'textWrapping') toks.push({ t: 'br', m }); break;
          case 'sym': toks.push({ t: 'text', s: L.mapSymbolChar(it.char, it.font) || it.char || '', m }); break;
          case 'ruby': toks.push({ t: 'text', s: it.base || '', m }); break;
          case 'bs': if (it.mdAttr) toks.push({ t: 'atom', md: `<a ${it.mdAttr}="${MD.escapeHTML(it.name)}"></a>`, html: `<a ${it.mdAttr}="${MD.escapeHTML(it.name)}"></a>`, m: {} }); break;
          case 'img': { const s = picture(it); if (s) toks.push({ t: 'atom', md: s.md, html: s.html, m: m.link ? { link: m.link } : {} }); break; }
          case 'fn': case 'en': if (!it.self) { const lb = noteLabel(it); toks.push({ t: 'atom', md: `[^${lb}]`, html: `<sup>${MD.escapeHTML(lb)}</sup>`, m: {} }); } break;
          default: break;
        }
      }
      return toks;
    }
    /** a link whose text is its address is written as an autolink (www.… as plain text, which GitHub links) */
    function autolinkTokens(toks) {
      for (let i = 0; i < toks.length; i++) {
        const l = toks[i].m.link;
        if (!l) continue;
        let j = i, t = '';
        while (j < toks.length && toks[j].m.link === l) { if (toks[j].t !== 'text' || Object.keys(toks[j].m).length > 1) { t = null; break; } t += toks[j].s; j++; }
        if (t == null) { i = j; continue; }
        const [href] = JSON.parse(l);
        let md = null;
        if (href === t && /^(https?|ftp):\/\/[^\s<>]+$/i.test(t)) md = `<${t}>`;
        else if (href === 'mailto:' + t && /^[^\s<>@]+@[^\s<>@]+$/.test(t)) md = `<${t}>`;
        else if (href === 'http://' + t && /^www\.[^\s<>*_~\\[\]`]+$/i.test(t) && !/[.,:;!?)]$/.test(t)) md = t;
        if (md) toks.splice(i, j - i, { t: 'atom', md, html: `<a href="${MD.escapeHTML(href)}">${MD.escapeHTML(t)}</a>`, m: {} });
      }
      return toks;
    }
    function marks(r) {
      const cs = charProps(r.style);
      const v = (k) => (r[k] != null ? r[k] : cs[k]);
      const m = {};
      if (r.style === 'MarkdownSource') { m.raw = 1; return m; }
      if (v('b')) m.b = 1;
      if (v('i')) m.i = 1;
      if (v('strike') || v('dstrike')) m.s = 1;
      const u = v('u');
      if (u && u !== 'none' && !r.link && r.style !== 'Hyperlink') m.u = 1;
      const vert = v('vert');
      if (vert === 'superscript' && r.style !== 'FootnoteReference') m.sup = 1;
      else if (vert === 'subscript') m.sub = 1;
      if (r.style === 'HTMLKeyboard') m.kbd = 1;
      else if (r.style === 'HTMLCode' || r.style === 'HTMLTypewriter' || r.style === 'HTMLSample' || /^(courier|consolas|cousine|menlo|monaco|lucida console|source code|dejavu sans mono|liberation mono)/i.test(r.font || '')) m.code = 1;
      if (r.link) {
        const l = r.link;
        const href = typeof l === 'string' ? l : (l.url || '') + (l.anchor ? '#' + l.anchor : '');
        if (href) m.link = JSON.stringify([href, (l && l.tip) || '']);
      }
      return m;
    }
    function noteLabel(it) {
      if (it.mdLabel) return it.mdLabel;
      if (it.t === 'en') return 'en-' + it.id;
      return String(it.id);
    }
    function picture(it) {
      const m = it.md;
      let src;
      if (m && m.media === it.media && m.src) src = m.src;
      else if (opts.check) src = 'media:' + it.media;
      else if (opts.zip) {
        const med = L.media.get(it.media);
        const ext = (med && L.mimeToExt(med.type)) || 'png';
        let base = String((med && med.name) || 'image').replace(/\.[^.]*$/, '').replace(/[^\w.-]+/g, '-').replace(/^-+|-+$/g, '') || 'image';
        if (base.length > 40) base = base.slice(0, 40);
        let path = `images/${base}.${ext}`, k = 1;
        while (usedPaths.has(path)) path = `images/${base}-${++k}.${ext}`;
        usedPaths.add(path);
        images.push({ path, media: it.media });
        src = path;
      } else { dropped++; return null; }
      const alt = it.alt || '';
      const title = (m && m.title) || '';
      const resized = m && (Math.abs(m.w - it.w) > 0.5 || Math.abs(m.h - it.h) > 0.5);
      const wPx = Math.round(it.w / 0.75), hPx = Math.round(it.h / 0.75);
      const html = `<img src="${MD.escapeHTML(src)}" alt="${MD.escapeHTML(alt)}"${title ? ` title="${MD.escapeHTML(title)}"` : ''} width="${wPx}"${resized ? ` height="${hPx}"` : ''}>`;
      if ((m && m.html) || resized) return { md: html, html };
      return { md: `![${escLabel(alt)}](${dest(src)}${title ? ' ' + titleMd(title) : ''})`, html };
    }
    const escLabel = (s) => s.replace(/[\\[\]]/g, '\\$&').replace(/\n/g, ' ');
    const dest = (s) => (/[\s<>]|^$/.test(s) || unbalanced(s) ? '<' + s.replace(/[<>\n]/g, (c) => encodeURIComponent(c)) + '>' : s.replace(/\\(?=[!-/:-@[-`{-~])/g, '\\\\'));
    const unbalanced = (s) => { let n = 0; for (const c of s) { if (c === '(') n++; else if (c === ')' && --n < 0) return true; } return n !== 0; };
    const titleMd = (t) => '"' + t.replace(/\\/g, '\\\\').replace(/"/g, '\\"') + '"';

    const ORDER = ['link', 'b', 'i', 's', 'u', 'sup', 'sub', 'kbd', 'code'];
    const MDDELIM = { b: '**', i: '*', s: '~~' };
    const HTMLTAG = { b: 'strong', i: 'em', s: 'del', u: 'ins', sup: 'sup', sub: 'sub', kbd: 'kbd', code: 'code' };
    /** Markdown for a paragraph's content; cell: inside a pipe table */
    function inline(p, ctx) {
      ctx = ctx || {};
      const toks = autolinkTokens(tokens(p));
      while (toks.length && (toks[toks.length - 1].t === 'br' || (toks[toks.length - 1].t === 'text' && !toks[toks.length - 1].m.raw && !/\S/.test(toks[toks.length - 1].s)))) toks.pop();
      if (ctx.skipCheckbox) { const i = toks.findIndex((t) => t.t === 'atom' && t.cb != null); if (i >= 0) { toks.splice(i, 1); const nx = toks[i]; if (nx && nx.t === 'text') nx.s = nx.s.replace(/^ /, ''); } }
      const dollars = toks.reduce((n, t) => n + (t.t === 'text' && !t.m.raw ? (t.s.match(/\$/g) || []).length : 0), 0);
      /* phase 1: parts with open/close markers; a mark that lasts longer opens first (outside) */
      const parts = [], stack = [];
      let ws = '';
      const keyOf = (k, m) => (k === 'link' ? 'link' + m.link : k);
      const extent = (k, i) => { let n = 0; const key = keyOf(k, toks[i].m); for (let j = i; j < toks.length && toks[j].m[k] && keyOf(k, toks[j].m) === key; j++) n++; return n; };
      toks.forEach((tk, i) => {
        const openLink = stack.find((x) => x.k === 'link');
        if (tk.t === 'text' && !tk.m.code && !tk.m.raw && /^[ \t]*$/.test(tk.s) && (tk.m.link || null) === (openLink ? openLink.link : null)) { ws += tk.s; return; }
        if (tk.t === 'br') { ws = ''; parts.push({ br: true }); return; }
        const want = new Set(ORDER.filter((k) => tk.m[k]).map((k) => keyOf(k, tk.m)));
        let keep = 0;
        while (keep < stack.length && want.has(stack[keep].key)) keep++;
        while (stack.length > keep) parts.push({ close: stack.pop() });
        if (ws) { parts.push({ text: ws, raw: true }); ws = ''; }
        const have = new Set(stack.map((s) => s.key));
        const add = ORDER.filter((k) => tk.m[k] && !have.has(keyOf(k, tk.m))).sort((a, b) => extent(b, i) - extent(a, i) || ORDER.indexOf(a) - ORDER.indexOf(b));
        for (const k of add) { const s = { k, key: keyOf(k, tk.m), link: tk.m.link }; stack.push(s); parts.push({ open: s }); }
        if (tk.t === 'atom') parts.push({ text: tk.md, raw: true });
        else if (tk.m.raw || tk.m.code) parts.push({ text: tk.s, raw: true, code: !!tk.m.code });
        else {
          let s = tk.s;
          if (/^[ \t]/.test(s) && !stack.length) { const m = /^[ \t]+/.exec(s)[0]; parts.push({ text: m, raw: true }); s = s.slice(m.length); }
          const m2 = /[ \t]+$/.exec(s);
          if (m2) { ws = m2[0]; s = s.slice(0, -ws.length); }
          parts.push({ text: escText(s, { cell: ctx.cell, dollars }), raw: true });
        }
      });
      while (stack.length) parts.push({ close: stack.pop() });
      /* phase 2: a delimiter run that would not open or close becomes an HTML tag */
      const flat = [];
      for (const pt of parts) {
        if (pt.open && MDDELIM[pt.open.k]) pt.open.md = true;
        flat.push(pt);
      }
      const charBefore = (i) => { for (let j = i - 1; j >= 0; j--) { const q = flat[j]; if (q.br) return '\n'; if (q.text) return q.text[q.text.length - 1]; if (q.open || q.close) return '*'; } return '\n'; };
      const charAfter = (i) => { for (let j = i + 1; j < flat.length; j++) { const q = flat[j]; if (q.br) return '\n'; if (q.text) return q.text[0]; if (q.open || q.close) return '*'; } return '\n'; };
      const ws1 = (c) => /[\s]/.test(c);
      const pun = (c) => /[\p{P}\p{S}]/u.test(c);
      flat.forEach((pt, i) => {
        const s = pt.open || pt.close;
        if (!s || !s.md) return;
        const b = charBefore(i), a = charAfter(i);
        if (pt.open) { const left = !ws1(a) && (!pun(a) || ws1(b) || pun(b)); if (!left) s.md = false; }
        else { const right = !ws1(b) && (!pun(b) || ws1(a) || pun(a)); if (!right) s.md = false; }
      });
      /* phase 3: text */
      let out = '';
      let code = null;
      for (const pt of flat) {
        if (code !== null && !(pt.close && pt.close.k === 'code')) { if (pt.text != null) code += pt.text; continue; }
        if (pt.br) { out += ctx.cell ? '<br>' : '\\\n'; continue; }
        if (pt.text != null) { out += pt.text; continue; }
        const s = pt.open || pt.close;
        if (s.k === 'code') {
          if (pt.open) { code = ''; continue; }
          out += codeSpan(code);
          code = null;
          continue;
        }
        if (s.k === 'link') {
          const [href, tip] = JSON.parse(s.link);
          out += pt.open ? '[' : `](${dest(href)}${tip ? ' ' + titleMd(tip) : ''})`;
          continue;
        }
        if (s.md) out += MDDELIM[s.k];
        else out += pt.open ? `<${HTMLTAG[s.k]}>` : `</${HTMLTAG[s.k]}>`;
      }
      if (code !== null) out += codeSpan(code);
      return out.split('\n').map((ln) => escLineStart(ln.replace(/[ \t]+$/, ''))).join('\n');
    }
    function codeSpan(s) {
      let n = 1;
      for (const m of s.match(/`+/g) || []) n = Math.max(n, m.length + 1);
      const f = '`'.repeat(n);
      const pad = /^`|`$/.test(s) || (/^ .*[^ ].* $/.test(s)) ? ' ' : '';
      return f + pad + s + pad + f;
    }
    function escText(s, o) {
      return s.replace(/[\\`*_[\]<&~|$\n\r!]/g, (c, i, str) => {
        const prev = str[i - 1] || ' ', next = str[i + 1] || ' ';
        switch (c) {
          case '\\': return /[!-/:-@[-`{-~]/.test(next) || i === str.length - 1 ? '\\\\' : c;
          case '_': return /[\p{L}\p{N}]/u.test(prev) && /[\p{L}\p{N}]/u.test(next) ? c : '\\_';
          case '<': return /[A-Za-z/!?]/.test(next) ? '\\<' : c;
          case '&': return /^&#?[A-Za-z0-9]+;/.test(str.slice(i)) ? '\\&' : c;
          case '|': return o.cell ? '\\|' : c;
          case '$': return o.dollars > 1 ? '\\$' : c;
          case '\n': return '&#10;';
          case '\r': return '&#13;';
          case '!': return i === str.length - 1 ? '\\!' : c;
          default: return '\\' + c;
        }
      });
    }
    function escLineStart(s) {
      s = s.replace(/^[ \t]+/, '');
      if (/^(#{1,6}|>|[-+*])([ \t]|$)/.test(s) || /^#{1,6}$/.test(s) || /^(=+|-+)[ \t]*$/.test(s)) return '\\' + s;
      const m = /^(\d{1,9})([.)])([ \t]|$)/.exec(s);
      if (m) return m[1] + '\\' + s.slice(m[1].length);
      return s;
    }
    /** HTML for a paragraph's content (cells of HTML tables) */
    function inlineHTML(p) {
      const toks = tokens(p);
      let out = '';
      const stack = [];
      for (const tk of toks) {
        const want = ORDER.filter((k) => tk.m[k]);
        while (stack.length && !want.includes(stack[stack.length - 1].k)) out += close(stack.pop());
        for (const k of want) if (!stack.some((s) => s.k === k)) { const s = { k, link: tk.m.link }; stack.push(s); out += open(s); }
        if (tk.t === 'br') out += '<br>';
        else if (tk.t === 'atom') out += tk.html;
        else out += MD.escapeHTML(tk.s);
      }
      while (stack.length) out += close(stack.pop());
      return out;
      function open(s) { if (s.k === 'link') { const [href, tip] = JSON.parse(s.link); return `<a href="${MD.escapeHTML(href)}"${tip ? ` title="${MD.escapeHTML(tip)}"` : ''}>`; } return `<${HTMLTAG[s.k]}>`; }
      function close(s) { return s.k === 'link' ? '</a>' : `</${HTMLTAG[s.k]}>`; }
    }

    /* ---------- blocks ---------- */
    const gap = (a, b) => {
      const pa = D.pProps(d, a), pb = D.pProps(d, b);
      const same = styleOf(a) === styleOf(b);
      const after = pa.contextual && same ? 0 : (pa.sp && pa.sp.a) || 0;
      const before = pb.contextual && same ? 0 : (pb.sp && pb.sp.b) || 0;
      return Math.max(after, before);
    };
    function table(t, ctx) {
      const map = L.R.tblMap(t);
      const simple = map.every((row) => row.every((m) => m.span === 1 && m.rowspan === 1 && !m.hidden && m.cell.blocks.length === 1 && m.cell.blocks[0].t === 'p' && !labels.get(m.cell.blocks[0].id)));
      const cellParas = (cell) => cell.blocks.filter((b) => b.t === 'p');
      if (simple && map.length) {
        const n = Math.max(...map.map((r) => r.length));
        const row = (cells) => '| ' + Array.from({ length: n }, (_, k) => (cells[k] ? inline(cells[k].cell.blocks[0], Object.assign({}, ctx, { cell: true })) : '')).join(' | ') + ' |';
        const alignRow = map.length > 1 ? map[1] : map[0];
        const delim = '| ' + Array.from({ length: n }, (_, k) => {
          const c = alignRow[k];
          const jc = c ? D.pProps(d, c.cell.blocks[0]).jc : null;
          return jc === 'center' ? ':-:' : jc === 'right' ? '--:' : c && c.cell.blocks[0].pPr.jc === 'left' ? ':--' : '---';
        }).join(' | ') + ' |';
        return [row(map[0]), delim, ...map.slice(1).map(row)];
      }
      const lines = ['<table>'];
      map.forEach((row, ri) => {
        let s = '<tr>';
        for (const m of row) {
          if (m.hidden) continue;
          const tag = ri === 0 && t.rows[0].trPr.header !== false && t.rows.length > 1 ? 'th' : 'td';
          const jc = cellParas(m.cell)[0] ? D.pProps(d, cellParas(m.cell)[0]).jc : null;
          const attrs = (m.span > 1 ? ` colspan="${m.span}"` : '') + (m.rowspan > 1 ? ` rowspan="${m.rowspan}"` : '') + (jc === 'center' || jc === 'right' ? ` align="${jc}"` : '');
          s += `<${tag}${attrs}>${cellParas(m.cell).map(inlineHTML).filter(Boolean).join('<br>')}</${tag}>`;
        }
        lines.push(s + '</tr>');
      });
      lines.push('</table>');
      return lines;
    }
    /** blocks → lines; keep: the run is being compared with what it was (no side effects) */
    function blocks(list) {
      const out = [];
      const stack = []; /* open list levels: { numId, ordered, indent, content, marker } */
      const loose = new Map();
      const lastMarker = []; /* per level: { numId, marker } of the list that ended just before */
      /* loose lists: items with space between them, or a plain paragraph inside an item */
      const prevItem = new Map(), itemAt = [];
      for (const b of list) {
        if (b.t !== 'p') continue;
        const x = info(b);
        if (x.kind === 'item') {
          const pv = prevItem.get(x.li.numId);
          if ((pv && gap(pv, b) > 0) || b.pPr.mdLoose) loose.set(x.li.numId, true);
          prevItem.set(x.li.numId, b);
          itemAt[x.li.lvl] = x.li.numId;
          itemAt.length = x.li.lvl + 1;
        } else if (x.cont && x.kind === 'para' && !x.q && itemAt[x.cont - 1]) loose.set(itemAt[x.cont - 1], true);
        else if (!x.cont && x.kind !== 'empty') itemAt.length = 0;
      }
      let prev = null;
      const listAt = (depth) => stack[depth - 1] || null;
      const quotePrefix = (q) => '> '.repeat(q);
      for (let i = 0; i < list.length; i++) {
        const b = list[i];
        const x = info(b);
        if (x.kind === 'empty') continue;
        /* list nesting for this block */
        let depth = 0;
        if (x.kind === 'item') {
          const lvl = x.li.lvl;
          while (stack.length > lvl + 1) stack.pop();
          if (stack.length === lvl + 1 && stack[lvl].numId !== x.li.numId) { lastMarker[lvl] = { numId: stack[lvl].numId, marker: stack[lvl].marker, ordered: stack[lvl].ordered }; stack.pop(); }
          if (stack.length < lvl + 1) {
            const parent = stack[stack.length - 1];
            const lm = lastMarker[stack.length];
            const adj = lm && lm.ordered === x.ordered && lm.numId !== x.li.numId && prev && prev.listy;
            const marker = x.ordered ? (adj && lm.marker === '.' ? ')' : '.') : adj && lm.marker === '-' ? '*' : '-';
            stack.push({ numId: x.li.numId, ordered: x.ordered, indent: parent ? parent.content : 0, content: 0, marker });
          }
          depth = stack.length;
          const lv = stack[depth - 1];
          const mk = x.ordered ? `${x.li.value}${lv.marker}` : lv.marker;
          lv.content = lv.indent + mk.length + 1;
          lv.mk = mk;
        } else if (x.cont) {
          while (stack.length > x.cont) stack.pop();
          depth = Math.min(stack.length, x.cont);
        } else {
          for (let k = 0; k < stack.length; k++) lastMarker[k] = { numId: stack[k].numId, marker: stack[k].marker, ordered: stack[k].ordered };
          stack.length = 0;
        }
        /* content lines */
        let lines;
        let codeEnd = i;
        switch (x.kind) {
          case 'heading': {
            const t = inline(b).replace(/\\\n/g, ' ').replace(/(#+)$/, '\\$1');
            lines = ['#'.repeat(x.level) + ' ' + (x.label ? escText(x.label, {}) + ' ' : '') + t];
            break;
          }
          case 'hr': lines = ['---']; break;
          case 'alertTitle': lines = [`[!${b.pPr.mdAlert.toUpperCase()}]`]; break;
          case 'code': {
            const body = [];
            let j = i;
            for (; j < list.length; j++) {
              const c = list[j];
              if (c.t !== 'p') break;
              const y = info(c);
              if (y.kind !== 'code' || y.lang !== x.lang || y.q !== x.q || y.cont !== x.cont || (j > i && c.pPr.num)) break;
              body.push(...codeText(c).split('\n'));
            }
            codeEnd = j - 1;
            let n = 3, ch = '`';
            for (const ln of body) { const m = /^\s*(`{3,}|~{3,})/.exec(ln); if (m && m[1][0] === '`') n = Math.max(n, m[1].length + 1); }
            if (/`/.test(x.lang)) ch = '~';
            const fence = ch.repeat(n);
            lines = [fence + x.lang, ...body, fence];
            break;
          }
          case 'table': lines = table(b, {}); break;
          case 'item': {
            const t = inline(b, { skipCheckbox: true });
            const cb = tokens(b).find((tk) => tk.t === 'atom' && tk.cb != null);
            const first = (x.heading ? '#'.repeat(x.heading) + ' ' : '') + (cb ? (cb.cb ? '[x] ' : '[ ] ') : '') + (x.heading ? t.replace(/\\\n/g, ' ').replace(/(#+)$/, '\\$1') : t);
            lines = first.split('\n');
            if (!t && !cb) lines = [''];
            break;
          }
          case 'empty': lines = ['']; break;
          default: lines = inline(b).split('\n');
        }
        /* list markers and indents */
        const lv = depth ? listAt(depth) : null;
        const qFirst = x.q && x.qIn && depth >= x.qIn;
        if (qFirst) lines = lines.map((ln) => (quotePrefix(x.q) + ln).replace(/ +$/, ''));
        if (x.kind === 'item') lines = lines.map((ln, k) => (k === 0 ? ' '.repeat(lv.indent) + lv.mk + (ln ? ' ' + ln : '') : ln ? ' '.repeat(lv.content) + ln : ln));
        else if (lv) lines = lines.map((ln) => (ln ? ' '.repeat(lv.content) + ln : ln));
        /* quote prefix */
        if (x.q && !qFirst) lines = lines.map((ln) => (quotePrefix(x.q) + ln).replace(/ +$/, ''));
        /* separator */
        if (prev) {
          let sep = true;
          /* inside a tight list nothing separates items, nested lists and code; a second paragraph needs a blank line */
          if (depth && prev.depth) {
            const owner = listAt(Math.min(prev.depth, depth));
            /* a paragraph after a paragraph needs a blank line; a quote opening after an item's text does not */
            const para = x.kind === 'para' && (!x.q || (prev.q === x.q && prev.kind === 'para'));
            if (owner && !loose.get(owner.numId) && !para) sep = false;
          }
          if (x.kind === 'item' && depth === 1 && prev.listy && prev.top !== x.li.numId) sep = true;
          if (prev.kind === 'alertTitle' && x.q === prev.q && x.alert === prev.alert) sep = false;
          if (prev.kind === 'item' && prev.empty && x.cont && x.cont >= depth) sep = false;
          if (sep) {
            let ln = prev.q && x.q && x.kind !== 'alertTitle' && prev.alert === x.alert ? '>' + ' >'.repeat(Math.min(prev.q, x.q) - 1) : '';
            if (ln && qFirst && lv) ln = ' '.repeat(lv.content) + ln;
            out.push(ln);
          }
        }
        out.push(...lines);
        prev = { kind: x.kind, empty: x.kind === 'item' && lines.length === 1 && /^\s*(\d+[.)]|[-*+])$/.test(lines[0]), q: x.q, alert: x.alert, depth, listy: x.kind === 'item' || !!x.cont, top: stack[0] && stack[0].numId };
        i = codeEnd;
      }
      return out;
    }
    function codeText(p) {
      let s = '';
      for (const it of p.runs) {
        if (it.rPr && it.rPr.del) continue;
        if (it.t === 'text') s += it.text;
        else if (it.t === 'tab') s += '\t';
        else if (it.t === 'br') s += '\n';
        else if (it.t === 'sym') s += L.mapSymbolChar(it.char, it.font) || it.char || '';
      }
      return s;
    }
    /* ---------- footnotes ---------- */
    const refs = new Map(); /* note id → label, in reading order */
    for (const p of D.allParas ? D.allParas(d) : []) for (const it of p.runs) if (it.t === 'fn' && !it.self && !(it.rPr && it.rPr.del) && !refs.has(it.id)) refs.set(it.id, noteLabel(it));
    function footnote(id) {
      const st = d.fn[id];
      if (!st) return [];
      const bl = st.blocks.map((b) => (b.t === 'p' ? Object.assign({}, b, { runs: dropSelf(b.runs) }) : b));
      const lines = blocks(bl);
      const label = refs.get(id) || String(id);
      const head = `[^${label}]:`;
      if (!lines.length) return [head];
      return lines.map((ln, k) => (k === 0 ? head + ' ' + ln : ln ? '    ' + ln : ln));
    }
    const dropSelf = (runs) => {
      const i = runs.findIndex((it) => it.t === 'fn' && it.self);
      if (i < 0) return runs;
      const out = runs.slice(0, i).concat(runs.slice(i + 1));
      if (out[i] && out[i].t === 'text') out[i] = Object.assign({}, out[i], { text: out[i].text.replace(/^ /, '') });
      return out;
    };
    return {
      info, joins, blocks, footnote, images,
      get dropped() { return dropped; },
      referenced: (id) => refs.has(id),
      noteOrder: () => [...refs.keys()].filter((id) => d.fn[id]),
    };
  }

  /* ================= .zip packages: a .md file and its pictures ================= */
  const MD_EXT = /\.(md|markdown|mdown|mkd)$/i;
  /** pick the Markdown file of a package: the only one, else README / index nearest the top */
  function mainMarkdown(names) {
    const mds = names.filter((n) => MD_EXT.test(n) && !/(^|\/)(__MACOSX|\.)/.test(n));
    if (mds.length <= 1) return mds[0] || null;
    const depth = (n) => n.split('/').length;
    const rank = (n) => (/(^|\/)readme\.[^/]+$/i.test(n) ? 0 : /(^|\/)index\.[^/]+$/i.test(n) ? 1 : 2);
    return mds.sort((a, b) => depth(a) - depth(b) || rank(a) - rank(b) || a.localeCompare(b))[0];
  }
  function joinPath(dir, rel) {
    const parts = (rel.startsWith('/') ? [] : dir.split('/').filter(Boolean));
    for (const seg of rel.split('/')) {
      if (!seg || seg === '.') continue;
      if (seg === '..') parts.pop(); else parts.push(seg);
    }
    return parts.join('/');
  }
  X.isPackage = (name) => /\.zip$/i.test(name);
  X.isMarkdown = (name) => MD_EXT.test(name);
  /** a .zip → document (d.md.zip keeps the package so saving writes the other files back unchanged) */
  X.readZip = async function (buf) {
    const files = await L.zip.read(buf);
    const names = [...files.keys()];
    const main = mainMarkdown(names);
    if (!main) { const e = new Error('There is no Markdown (.md) file in this .zip file.'); e.code = 'nomd'; throw e; }
    const dir = main.includes('/') ? main.slice(0, main.lastIndexOf('/')) : '';
    const lower = new Map(names.map((n) => [n.toLowerCase(), n]));
    const resolve = async (rel) => {
      const p = joinPath(dir, rel);
      const name = files.has(p) ? p : lower.get(p.toLowerCase());
      if (!name) return null;
      const bytes = await files.get(name).bytes();
      return new Blob([bytes], { type: L.extToMime(name.split('.').pop()) });
    };
    const text = await files.get(main).text();
    const d = await X.read(text, { resolve });
    d.md.zip = { files, main };
    return d;
  };
  /** document → .zip Blob: the Markdown file, the pictures it uses, and the other files of the package it came from */
  X.writeZip = async function (d, name) {
    const z = d.md && d.md.zip;
    const main = z ? z.main : (name || 'Document') + '.md';
    const dir = main.includes('/') ? main.slice(0, main.lastIndexOf('/') + 1) : '';
    const res = X.write(d, { zip: true, paths: z ? [...z.files.keys()].map((n) => (n.startsWith(dir) ? n.slice(dir.length) : '')) : [] });
    const out = [];
    if (z) for (const [n, e] of z.files) if (n !== main && !n.endsWith('/')) out.push({ name: n, data: await e.bytes() });
    out.push({ name: main, data: res.text });
    for (const im of res.images) {
      const m = L.media.get(im.media);
      if (m) out.push({ name: dir + im.path, data: new Uint8Array(await m.blob.arrayBuffer()), store: /^image\/(png|jpeg|gif|webp)$/.test(m.type) });
    }
    return L.zip.write(out, 'application/zip');
  };
  /** pictures of an opened .md that could not be loaded (they show as boxes) */
  X.missingPictures = function (d) {
    let n = 0;
    D.walk(d.main, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.t === 'img' && it.md && it.md.missing && it.md.media === it.media) n++; });
    return n;
  };
  /** pictures the document has that a plain .md file cannot carry */
  X.newPictures = function (d) {
    let n = 0;
    D.walk(d.main, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.t === 'img' && !(it.md && it.md.media === it.media && it.md.src)) n++; });
    return n;
  };
})();
