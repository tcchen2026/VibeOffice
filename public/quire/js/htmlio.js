/* Quire — HTML in and out: clipboard HTML (including Word's and Google Docs' flavours), opening .htm
 * files, Save as Web Page (single self-contained file) and plain-text export.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, R = L.R;
  const H = (L.htmlio = {});
  const doc = () => D.doc;
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const attr = (s) => esc(s).replace(/"/g, '&quot;');

  /* ================= export ================= */
  function runStyle(r, base) {
    const st = [];
    if (r.font && (!base || r.font !== base.font)) st.push(`font-family:${cssFont(r.font)}`);
    if (r.sz && (!base || r.sz !== base.sz)) st.push(`font-size:${L.round(r.sz, 1)}pt`);
    if (r.color && r.color !== 'auto') st.push(`color:#${r.color}`);
    if (r.hl && r.hl !== 'none') st.push(`background:${R.HIGHLIGHT[r.hl] || r.hl}`);
    else if (r.shd && r.shd.fill && r.shd.fill !== 'auto') st.push(`background:#${r.shd.fill}`);
    if (r.caps) st.push('text-transform:uppercase');
    if (r.smallCaps) st.push('font-variant:small-caps');
    if (r.spacing) st.push(`letter-spacing:${L.round(r.spacing, 2)}pt`);
    if (r.hidden) st.push('display:none');
    return st.join(';');
  }
  const cssFont = (f) => (/\s/.test(f) ? `'${f}'` : f) + (/courier|mono|consol/i.test(f) ? ',monospace' : /arial|helvetica|verdana|tahoma|calibri|segoe|sans/i.test(f) ? ',sans-serif' : ',serif');
  function wrapTags(r, inner) {
    let s = inner;
    if (r.vert === 'superscript') s = `<sup>${s}</sup>`;
    else if (r.vert === 'subscript') s = `<sub>${s}</sub>`;
    if (r.strike || r.dstrike) s = `<s>${s}</s>`;
    if (r.u && r.u !== 'none') s = `<u>${s}</u>`;
    if (r.i) s = `<i>${s}</i>`;
    if (r.b) s = `<b>${s}</b>`;
    return s;
  }
  function linkHref(l) { if (!l) return null; if (typeof l === 'string') return l; return (l.url || '') + (l.anchor ? '#' + l.anchor : ''); }
  function parasHTML(p, d, opts, labels) {
    const ctx = opts.tblCtx || null;
    const pp = D.pProps(d, p, ctx);
    const markR = D.rProps(d, p, null, ctx);
    let inner = '';
    const lb = labels && labels.get(p.id);
    const inList = opts.inList && lb;
    if (lb && !inList) inner += `<span style="${runStyle(D.mergeRPr(markR, lb.lv.rPr || {}), null)}">${esc(lb.text)}</span>&nbsp;&nbsp;`;
    let linkOpen = null;
    for (const it of p.runs) {
      const r = D.rProps(d, p, it, ctx);
      if (it.rPr && it.rPr.del) continue;
      const href = linkHref(it.rPr && it.rPr.link);
      if (href !== linkOpen) { if (linkOpen != null) inner += '</a>'; if (href) inner += `<a href="${attr(href)}">`; linkOpen = href || null; }
      let piece = '';
      if (it.t === 'text') piece = esc(it.text).replace(/ {2}/g, ' &nbsp;');
      else if (it.t === 'tab') piece = '&emsp;';
      else if (it.t === 'br') piece = it.type === 'page' && opts.page ? '<br style="page-break-before:always">' : '<br>';
      else if (it.t === 'sym') piece = esc(L.mapSymbolChar(it.char, it.font) || it.char || '');
      else if (it.t === 'ruby') piece = `<ruby>${esc(it.base || '')}<rt>${esc(it.rt || '')}</rt></ruby>`;
      else if (it.t === 'raw' && it.math) { const m = L.R.mathFor(it); piece = m ? m.math.outerHTML : esc(it.text || ''); }
      else if (it.t === 'img' && it.media) { const src = opts.media ? opts.media.get(it.media) || L.media.url(it.media) : L.media.url(it.media); piece = `<img src="${attr(src)}" width="${Math.round((it.w || 72) * D.PX)}" height="${Math.round((it.h || 72) * D.PX)}" alt="${attr(it.alt || '')}"${it.float ? ` style="float:${it.float.posH && it.float.posH.align === 'right' ? 'right' : 'left'};margin:4px 9px"` : ''}>`; }
      else if (it.t === 'shape' && it.tb) piece = `<span style="display:inline-block;border:1px solid #000;padding:3px 7px">${it.tb.blocks.filter((b) => b.t === 'p').map((q) => esc(D.plainText(q))).join('<br>')}</span>`;
      else if ((it.t === 'fn' || it.t === 'en') && !it.self) { const n = L.layout.noteNum ? L.layout.noteNum(it) : '*'; piece = opts.notes ? `<sup><a href="#_${it.t}${it.id}" name="_${it.t}ref${it.id}">[${esc(n)}]</a></sup>` : `<sup>${esc(n)}</sup>`; if (opts.notes) opts.notes.push(it); }
      else continue;
      const css = runStyle(r, opts.baseR);
      piece = wrapTags(r, piece);
      inner += css ? `<span style="${css}">${piece}</span>` : piece;
    }
    if (linkOpen != null) inner += '</a>';
    if (!inner) inner = '&nbsp;';
    const st = [];
    const sp = pp.sp || {};
    st.push(`margin:${L.round(sp.b || 0, 1)}pt ${L.round((pp.ind && pp.ind.r) || 0, 1)}pt ${L.round(sp.a || 0, 1)}pt ${L.round(inList ? 0 : (pp.ind && pp.ind.l) || 0, 1)}pt`);
    if (pp.ind && pp.ind.fl && !inList) st.push(`text-indent:${L.round(pp.ind.fl, 1)}pt`);
    if (pp.jc && pp.jc !== 'left') st.push(`text-align:${{ both: 'justify', distribute: 'justify', center: 'center', right: 'right' }[pp.jc] || 'left'}`);
    if (sp.line && sp.rule === 'auto' && sp.line !== 1) st.push(`line-height:${L.round(sp.line * 1.15, 2)}`);
    else if (sp.line && (sp.rule === 'exact' || sp.rule === 'atLeast')) st.push(`line-height:${L.round(sp.line, 1)}pt`);
    if (pp.shd && pp.shd.fill && pp.shd.fill !== 'auto') st.push(`background:#${pp.shd.fill}`);
    if (pp.borders) for (const k of ['top', 'left', 'bottom', 'right']) if (R.hasBorder(pp.borders[k])) st.push(`border-${k}:${R.borderCSS(pp.borders[k])};padding-${k}:${L.round(pp.borders[k].space || 1, 1)}pt`);
    if (pp.pageBreakBefore && opts.page) st.push('page-break-before:always');
    const fstyle = runStyle(markR, null);
    const lvl = D.headingLevel(d, p);
    const tag = inList ? 'li' : lvl && lvl <= 6 ? 'h' + lvl : 'p';
    return `<${tag} style="${st.join(';')};${fstyle}">${inner}</${tag}>`;
  }
  function tableHTML(t, d, opts, labels) {
    const { tp, sp } = R.tblProps(d, t);
    const map = R.tblMap(t);
    const total = t.grid.reduce((a, b) => a + b, 0);
    let s = `<table style="border-collapse:collapse;width:${L.round(total, 1)}pt${tp.jc === 'center' ? ';margin:0 auto' : ''}" cellspacing="0" cellpadding="0">`;
    const nR = t.rows.length, nC = t.grid.length;
    map.forEach((row, ri) => {
      s += '<tr' + (t.rows[ri].trPr.h ? ` style="height:${L.round(t.rows[ri].trPr.h, 1)}pt"` : '') + '>';
      for (const m of row) {
        if (m.hidden) continue;
        const tcPr = m.cell.tcPr || {};
        const cs = R.cellStyle(d, t, sp, ri, m.c0, nR, nC);
        const st = [`width:${L.round(t.grid.slice(m.c0, m.c0 + m.span).reduce((a, b) => a + b, 0), 1)}pt`, `padding:${L.round(tp.cellMar.t || 0, 1)}pt ${L.round(tp.cellMar.r || 0, 1)}pt ${L.round(tp.cellMar.b || 0, 1)}pt ${L.round(tp.cellMar.l || 0, 1)}pt`, 'vertical-align:' + ({ center: 'middle', bottom: 'bottom' }[tcPr.vAlign || cs.tcPr.vAlign] || 'top')];
        const tb = tp.borders || {};
        const own = Object.assign({}, cs.borders || {}, tcPr.borders || {});
        const side = (k, outer, inner, isOuter) => own[k] || (isOuter ? tb[outer] : tb[inner]);
        const bds = { top: side('top', 'top', 'insideH', ri === 0), bottom: side('bottom', 'bottom', 'insideH', ri + m.rowspan >= nR), left: side('left', 'left', 'insideV', m.c0 === 0), right: side('right', 'right', 'insideV', m.c0 + m.span >= nC) };
        for (const k in bds) st.push(`border-${k}:${R.hasBorder(bds[k]) ? R.borderCSS(bds[k]) : 'none'}`);
        const shd = tcPr.shd || cs.tcPr.shd;
        if (shd) { const c = R.shadeCSS(shd); if (c) st.push(`background:${c}`); }
        const cellOpts = Object.assign({}, opts, { tblCtx: { tblRPr: D.mergeRPr(sp.rPr, cs.rPr), tblPPr: D.mergePPr(sp.pPr, cs.pPr) } });
        s += `<td${m.span > 1 ? ` colspan="${m.span}"` : ''}${m.rowspan > 1 ? ` rowspan="${m.rowspan}"` : ''} style="${st.join(';')}">${blocksHTML(m.cell.blocks, d, cellOpts, labels)}</td>`;
      }
      s += '</tr>';
    });
    return s + '</table>';
  }
  function blocksHTML(blocks, d, opts, labels) {
    let s = '';
    const stack = []; /* open lists: {numId, lvl, tag} */
    const closeTo = (n) => { while (stack.length > n) s += `</${stack.pop().tag}>`; };
    for (const b of blocks) {
      if (b.t === 'tbl') { closeTo(0); s += tableHTML(b, d, opts, labels); continue; }
      if (b.t !== 'p') continue;
      const lb = labels && labels.get(b.id);
      if (lb && opts.lists !== false) {
        const lvl = lb.lvl || 0;
        const tag = lb.lv.fmt === 'bullet' ? 'ul' : 'ol';
        while (stack.length && (stack.length > lvl + 1 || (stack.length === lvl + 1 && (stack[lvl].numId !== lb.numId || stack[lvl].tag !== tag)))) s += `</${stack.pop().tag}>`;
        while (stack.length < lvl + 1) {
          const lt = lb.lv.fmt === 'upperRoman' ? 'I' : lb.lv.fmt === 'lowerRoman' ? 'i' : lb.lv.fmt === 'upperLetter' ? 'A' : lb.lv.fmt === 'lowerLetter' ? 'a' : '1';
          const start = stack.length === lvl && tag === 'ol' && lb.value > 1 ? ` start="${lb.value}"` : '';
          s += tag === 'ol' ? `<ol type="${lt}"${start} style="margin-top:0;margin-bottom:0">` : '<ul style="margin-top:0;margin-bottom:0">';
          stack.push({ numId: lb.numId, lvl: stack.length, tag });
        }
        s += parasHTML(b, d, Object.assign({}, opts, { inList: true }), labels);
        continue;
      }
      closeTo(0);
      s += parasHTML(b, d, opts, labels);
    }
    closeTo(0);
    return s;
  }
  H.blocksToHTML = function (blocks, opts) {
    opts = opts || {};
    const d = doc();
    const labels = D.computeLists(d, [{ blocks }]);
    const html = blocksHTML(blocks, d, Object.assign({ baseR: null }, opts), labels);
    return opts.clipboard ? `<meta charset="utf-8"><div>${html}</div>` : html;
  };
  /** Save as Web Page: one self-contained .htm with embedded pictures */
  H.webPage = async function (d, name) {
    d = d || doc();
    const media = new Map();
    for (const [id, m] of L.media.all()) { try { media.set(id, await L.blobToDataURL(m.view || m.blob)); } catch (e) { /* skip */ } }
    const labels = D.computeLists(d);
    const notes = [];
    const baseR = D.styleProps(d, 'Normal').rPr;
    const body = blocksHTML(d.main.blocks, d, { media, notes, page: true, baseR }, labels);
    let foot = '';
    if (notes.length) {
      foot = '<hr style="width:33%;text-align:left;margin-left:0">';
      for (const it of notes) {
        const st = (it.t === 'fn' ? d.fn : d.en)[it.id];
        if (!st) continue;
        const n = L.layout.noteNum ? L.layout.noteNum(it) : '*';
        foot += `<div id="_${it.t}${it.id}"><a href="#_${it.t}ref${it.id}">[${esc(n)}]</a> ${blocksHTML(st.blocks.map((b) => (b.t === 'p' ? Object.assign({}, b, { runs: b.runs.filter((x) => !x.self) }) : b)), d, { media, baseR }, labels).replace(/^<p([^>]*)>/, '<span$1>').replace(/<\/p>$/, '</span>')}</div>`;
      }
    }
    const title = d.props.title || name || 'Document';
    const bg = d.bg ? `background:#${d.bg};` : '';
    const html = `<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="utf-8">\n<meta name="generator" content="Quire 2003">\n<title>${esc(title)}</title>\n` +
      `${d.props.creator ? `<meta name="author" content="${attr(d.props.creator)}">\n` : ''}${d.props.description ? `<meta name="description" content="${attr(d.props.description)}">\n` : ''}` +
      `<style>body{${bg}font-family:${cssFont(baseR.font || 'Times New Roman')};font-size:${baseR.sz || 12}pt;margin:1in;max-width:${L.round((d.sect.pgW - d.sect.ml - d.sect.mr) / 72, 2)}in}p,h1,h2,h3,h4,h5,h6,li{margin:0}table{margin:4pt 0}a{color:#0000ff}</style>\n</head>\n<body>\n${body}\n${foot}\n</body>\n</html>\n`;
    return new Blob([html], { type: 'text/html;charset=utf-8' });
  };
  /** plain text (Save as .txt): list numbers included, tables tab-separated, line breaks as CRLF */
  H.plainText = function (d) {
    d = d || doc();
    const labels = D.computeLists(d);
    const out = [];
    const walk = (blocks) => {
      for (const b of blocks) {
        if (b.t === 'p') { const lb = labels.get(b.id); out.push((lb ? lb.text + '\t' : '') + D.plainText(b, { skipDeleted: true }).replace(/\n/g, '\r\n')); }
        else if (b.t === 'tbl') for (const r of b.rows) out.push(r.cells.filter((c) => c.tcPr.vMerge !== 'continue').map((c) => c.blocks.filter((x) => x.t === 'p').map((x) => D.plainText(x, { skipDeleted: true })).join(' ')).join('\t'));
      }
    };
    walk(d.main.blocks);
    return out.join('\r\n') + '\r\n';
  };

  /* ================= import ================= */
  const BLOCK = new Set(['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'LI', 'BLOCKQUOTE', 'PRE', 'ADDRESS', 'CENTER', 'DT', 'DD', 'FIGCAPTION', 'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'MAIN', 'ASIDE', 'NAV', 'FIGURE', 'FORM', 'FIELDSET', 'LEGEND', 'TITLE']);
  const SKIP = new Set(['SCRIPT', 'STYLE', 'HEAD', 'META', 'LINK', 'NOSCRIPT', 'TEMPLATE', 'IFRAME', 'OBJECT', 'EMBED', 'BUTTON', 'SELECT', 'INPUT', 'TEXTAREA', 'SVG', 'CANVAS', 'VIDEO', 'AUDIO']);
  const pt = (v, base) => {
    if (v == null || v === '') return null;
    const m = /^(-?[\d.]+)(pt|px|in|cm|mm|em|rem|%|pc)?$/i.exec(String(v).trim());
    if (!m) return null;
    const n = parseFloat(m[1]);
    switch ((m[2] || 'px').toLowerCase()) { case 'pt': return n; case 'px': return n * 0.75; case 'in': return n * 72; case 'cm': return n * 28.3465; case 'mm': return n * 2.83465; case 'pc': return n * 12; case 'em': case 'rem': return n * (base || 12); case '%': return (n / 100) * (base || 12); default: return n; }
  };
  const SIZE_KW = { 'xx-small': 7.5, 'x-small': 7.5, small: 10, medium: 12, large: 13.5, 'x-large': 18, 'xx-large': 24, smaller: 10, larger: 14 };
  const FONT_SIZE_ATTR = [0, 7.5, 10, 12, 13.5, 18, 24, 36];
  function parseColor(c) {
    if (!c) return null;
    c = String(c).trim().toLowerCase();
    if (c === 'transparent' || c === 'inherit' || c === 'initial' || c === 'windowtext' || c === 'currentcolor') return null;
    let m = /^#([0-9a-f]{3})$/.exec(c);
    if (m) return m[1].split('').map((x) => x + x).join('').toUpperCase();
    m = /^#([0-9a-f]{6})/.exec(c);
    if (m) return m[1].toUpperCase();
    m = /^rgba?\(\s*(\d+)[\s,]+(\d+)[\s,]+(\d+)(?:[\s,/]+([\d.]+%?))?/.exec(c);
    if (m) { if (m[4] != null && parseFloat(m[4]) === 0) return null; return [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase(); }
    if (L.color && L.color.PRESET && L.color.PRESET[c]) return String(L.color.PRESET[c]).replace('#', '').toUpperCase();
    const named = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080', grey: '808080', silver: 'C0C0C0', maroon: '800000', navy: '000080', purple: '800080', teal: '008080', olive: '808000', lime: '00FF00', aqua: '00FFFF', fuchsia: 'FF00FF', orange: 'FFA500' };
    return named[c] || null;
  }
  const HL_NEAR = [['yellow', 'FFFF00'], ['green', '00FF00'], ['cyan', '00FFFF'], ['magenta', 'FF00FF'], ['blue', '0000FF'], ['red', 'FF0000'], ['darkBlue', '000080'], ['darkCyan', '008080'], ['darkGreen', '008000'], ['darkMagenta', '800080'], ['darkRed', '800000'], ['darkYellow', '808000'], ['darkGray', '808080'], ['lightGray', 'C0C0C0'], ['black', '000000']];
  function styleMap(el) {
    const out = {};
    const s = el.getAttribute && el.getAttribute('style');
    if (!s) return out;
    for (const part of s.split(';')) { const i = part.indexOf(':'); if (i < 0) continue; out[part.slice(0, i).trim().toLowerCase()] = part.slice(i + 1).trim().replace(/\s*!important$/, ''); }
    return out;
  }
  /** run formatting contributed by an element */
  function elRPr(el, inherited) {
    const r = Object.assign({}, inherited);
    const tag = el.tagName;
    if (tag === 'B' || tag === 'STRONG') r.b = true;
    if (tag === 'I' || tag === 'EM' || tag === 'CITE' || tag === 'DFN' || tag === 'VAR') r.i = true;
    if (tag === 'U' || tag === 'INS') r.u = 'single';
    if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') r.strike = true;
    if (tag === 'SUP') r.vert = 'superscript';
    if (tag === 'SUB') r.vert = 'subscript';
    if (tag === 'CODE' || tag === 'TT' || tag === 'KBD' || tag === 'SAMP' || tag === 'PRE') r.font = 'Courier New';
    if (tag === 'MARK') r.hl = 'yellow';
    if (tag === 'SMALL') r.sz = Math.max(1, (r.sz || 12) * 0.83);
    if (tag === 'BIG') r.sz = (r.sz || 12) * 1.2;
    if (tag === 'FONT') {
      if (el.getAttribute('face')) r.font = el.getAttribute('face').split(',')[0].replace(/["']/g, '').trim();
      if (el.getAttribute('size')) { const n = parseInt(el.getAttribute('size'), 10); if (n >= 1 && n <= 7) r.sz = FONT_SIZE_ATTR[n]; }
      if (el.getAttribute('color')) { const c = parseColor(el.getAttribute('color')); if (c) r.color = c; }
    }
    if (/^H[1-6]$/.test(tag)) r.b = true;
    const st = styleMap(el);
    if (st['font-family']) { const f = st['font-family'].split(',')[0].replace(/["']/g, '').trim(); if (f && !/^(inherit|initial|serif|sans-serif|monospace)$/i.test(f)) r.font = f.replace(/^&quot;|&quot;$/g, ''); }
    if (st['font-size']) { const v = SIZE_KW[st['font-size']] || pt(st['font-size'], inherited.sz || 12); if (v) r.sz = L.round(Math.max(1, v) * 2, 0) / 2; }
    if (st['font-weight']) { const w = st['font-weight']; r.b = w === 'bold' || w === 'bolder' || +w >= 600; }
    if (st['font-style']) r.i = /italic|oblique/.test(st['font-style']);
    const td = st['text-decoration'] || st['text-decoration-line'];
    if (td) { if (/underline/.test(td)) r.u = 'single'; else if (/none/.test(td)) delete r.u; if (/line-through/.test(td)) r.strike = true; }
    if (st.color) { const c = parseColor(st.color); if (c) r.color = c === '000000' ? r.color && r.color !== '000000' ? 'auto' : undefined : c; if (r.color === undefined) delete r.color; }
    const bg = st['background-color'] || st.background;
    if (bg && tag !== 'P' && tag !== 'DIV' && tag !== 'TD' && tag !== 'TH') {
      const c = parseColor(bg.split(' ')[0]);
      if (c && c !== 'FFFFFF') { const hl = HL_NEAR.find((x) => x[1] === c); if (hl) r.hl = hl[0]; else r.shd = { val: 'clear', fill: c, color: 'auto' }; }
    }
    if (st['vertical-align']) { const v = st['vertical-align']; if (v === 'super') r.vert = 'superscript'; else if (v === 'sub') r.vert = 'subscript'; else if (v === 'baseline') delete r.vert; }
    if (st['text-transform'] === 'uppercase') r.caps = true;
    if (st['font-variant'] === 'small-caps') r.smallCaps = true;
    if (st['letter-spacing'] && st['letter-spacing'] !== 'normal') { const v = pt(st['letter-spacing']); if (v) r.spacing = L.round(v, 2); }
    if (st.display === 'none' || st['mso-hide'] === 'all') r.hidden = true;
    return r;
  }
  function elPPr(el, base) {
    const p = Object.assign({}, base);
    const st = styleMap(el);
    const align = st['text-align'] || el.getAttribute && el.getAttribute('align');
    if (align) { const a = align.toLowerCase(); p.jc = a === 'center' ? 'center' : a === 'right' || a === 'end' ? 'right' : a === 'justify' ? 'both' : 'left'; if (p.jc === 'left') delete p.jc; }
    if (el.tagName === 'CENTER') p.jc = 'center';
    const ml = pt(st['margin-left'] || st['padding-left'] && null);
    if (ml) p.ind = Object.assign({}, p.ind, { l: L.round(Math.max(0, ml + ((p.ind && p.ind.l) || 0)), 2) });
    const ti = pt(st['text-indent']);
    if (ti) p.ind = Object.assign({}, p.ind, { fl: L.round(ti, 2) });
    const mt = pt(st['margin-top']), mb = pt(st['margin-bottom']);
    if (mt != null || mb != null) p.sp = Object.assign({}, p.sp, mt != null ? { b: L.round(Math.max(0, mt), 2) } : {}, mb != null ? { a: L.round(Math.max(0, mb), 2) } : {});
    const lh = st['line-height'];
    if (lh && lh !== 'normal') { if (/^[\d.]+$/.test(lh)) p.sp = Object.assign({}, p.sp, { line: L.round(+lh / 1.15, 2), rule: 'auto' }); else if (/%$/.test(lh)) p.sp = Object.assign({}, p.sp, { line: L.round(parseFloat(lh) / 115, 2), rule: 'auto' }); else { const v = pt(lh); if (v) p.sp = Object.assign({}, p.sp, { line: L.round(v, 2), rule: 'atLeast' }); } }
    const bg = st['background-color'] || st.background;
    if (bg && el.tagName !== 'TD' && el.tagName !== 'TH') { const c = parseColor(bg.split(' ')[0]); if (c && c !== 'FFFFFF') p.shd = { val: 'clear', fill: c, color: 'auto' }; }
    if (st['page-break-before'] === 'always' || st['break-before'] === 'page') p.pageBreakBefore = true;
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const b = st['border-' + side] || st.border;
      if (b && !/none|hidden|^0/.test(b)) { const m = /([\d.]+(?:px|pt)?)\s+(solid|dotted|dashed|double)\s+(.+)/.exec(b); if (m) { p.borders = Object.assign({}, p.borders); p.borders[side] = { val: m[2] === 'solid' ? 'single' : m[2], sz: L.round(pt(m[1]) || 0.5, 2), color: parseColor(m[3]) || 'auto', space: 1 }; } }
    }
    return p;
  }
  /** convert HTML to blocks. Images given as data: URLs become embedded pictures. */
  H.htmlToBlocks = async function (html, opts) {
    opts = opts || {};
    const d = doc();
    const dp = new DOMParser();
    const dom = dp.parseFromString(String(html).replace(/<!--\[if[\s\S]*?<!\[endif\]-->/g, ''), 'text/html');
    const root = dom.body || dom.documentElement;
    const blocks = [];
    const pending = [];
    const listDefs = new Map();
    let cur = null;
    const flush = () => { cur = null; };
    const curPara = (pPr, rPr) => { if (!cur) { cur = D.para([], L.clone(pPr || {}), L.clone(rPr || {})); blocks.push(cur); } return cur; };
    const target = (arr) => ({ push: (b) => arr.push(b) });
    void target;
    const listNum = (type, kind) => {
      const key = kind + ':' + type;
      if (listDefs.has(key)) return listDefs.get(key);
      const abs = kind === 'ul' ? D.makeBulletAbs('•', 'Symbol', '') : D.makeNumberAbs(type === 'a' ? 'lowerLetter' : type === 'A' ? 'upperLetter' : type === 'i' ? 'lowerRoman' : type === 'I' ? 'upperRoman' : 'decimal', '%1.');
      const id = D.addNum(d, abs);
      listDefs.set(key, id);
      return id;
    };
    async function addImage(el, rPr, para) {
      const src = el.getAttribute('src') || '';
      let w = parseFloat(el.getAttribute('width')) || 0, hh = parseFloat(el.getAttribute('height')) || 0;
      const st = styleMap(el);
      if (st.width) w = (pt(st.width) || w * 0.75) / 0.75;
      if (st.height) hh = (pt(st.height) || hh * 0.75) / 0.75;
      if (!/^data:|^blob:/.test(src)) { if (el.getAttribute('alt')) para.runs.push(D.text('[' + el.getAttribute('alt') + ']', rPr)); return; }
      try {
        const blob = await (await fetch(src)).blob();
        const id = L.media.add(blob, el.getAttribute('alt') || 'image.' + (L.mimeToExt(blob.type) || 'png'));
        if (!w || !hh) { const im = await L.loadImage(L.media.url(id)); w = w || im.naturalWidth; hh = hh || im.naturalHeight * (w / im.naturalWidth); }
        para.runs.push(D.item('img', { media: id, w: L.round(w * 0.75, 2), h: L.round(hh * 0.75, 2), alt: el.getAttribute('alt') || '' }, {}));
      } catch (e) { /* unreadable image */ }
    }
    async function inline(node, rPr, pPr, ctxList) {
      for (const ch of Array.from(node.childNodes)) {
        if (ch.nodeType === 3) {
          let t = ch.data;
          if (!opts.pre) t = t.replace(/[\r\n\t ]+/g, ' ');
          if (!t) continue;
          const p = curPara(pPr, rPr);
          if (!D.plen(p) && !opts.pre) t = t.replace(/^ +/, '');
          if (!t) continue;
          if (opts.pre) { const lines = t.split('\n'); lines.forEach((ln, i) => { if (i) { flush(); curPara(pPr, rPr); } if (ln) curPara(pPr, rPr).runs.push(D.text(ln.replace(/\t/g, '    '), L.clone(rPr))); }); continue; }
          p.runs.push(D.text(t.replace(/ /g, ' '), L.clone(rPr)));
          continue;
        }
        if (ch.nodeType !== 1) continue;
        const tag = ch.tagName;
        if (SKIP.has(tag)) continue;
        const st = styleMap(ch);
        if (st['mso-list'] === 'Ignore' || /mso-list:\s*Ignore/i.test(ch.getAttribute('style') || '')) continue;
        if (tag === 'BR') { const p = curPara(pPr, rPr); p.runs.push(D.item('br', { type: 'line' }, L.clone(rPr))); continue; }
        if (tag === 'IMG') { await addImage(ch, rPr, curPara(pPr, rPr)); continue; }
        if (tag === 'HR') { flush(); blocks.push(D.para([], { borders: { bottom: { val: 'single', sz: 1.5, color: 'auto', space: 1 } } })); flush(); continue; }
        if (tag === 'TABLE') { flush(); const t = await table(ch, rPr); if (t) { blocks.push(t); } flush(); continue; }
        if (tag === 'UL' || tag === 'OL') { flush(); await list(ch, rPr, pPr, ctxList ? ctxList.lvl + 1 : 0); flush(); continue; }
        if (BLOCK.has(tag)) { flush(); await block(ch, rPr, pPr, ctxList); flush(); continue; }
        const r2 = elRPr(ch, rPr);
        if (tag === 'A' && ch.getAttribute('href')) {
          const href = ch.getAttribute('href');
          r2.link = href.startsWith('#') ? { anchor: href.slice(1) } : { url: href };
          r2.style = 'Hyperlink';
          if (!d.styles.Hyperlink) d.styles.Hyperlink = D.builtinStyles().Hyperlink;
        }
        await inline(ch, r2, pPr, ctxList);
      }
    }
    async function block(el, rPr, pPr, ctxList) {
      const tag = el.tagName;
      let p2 = elPPr(el, pPr);
      const r2 = elRPr(el, rPr);
      if (/^H[1-6]$/.test(tag)) { p2.style = 'Heading' + tag[1]; delete r2.b; if (+tag[1] > 3 && !d.styles[p2.style]) d.styles[p2.style] = D.builtinStyles()[p2.style]; }
      if (tag === 'BLOCKQUOTE') p2.ind = Object.assign({}, p2.ind, { l: ((p2.ind && p2.ind.l) || 0) + 36, r: 36 });
      if (tag === 'PRE') { r2.font = 'Courier New'; r2.sz = 10; }
      /* Word's own list paragraphs: class MsoListParagraph + mso-list:l0 level1 lfo1 */
      const st = styleMap(el);
      if (st['mso-list'] && /level(\d)/.test(st['mso-list'])) {
        const lvl = +/level(\d)/.exec(st['mso-list'])[1] - 1;
        const ign = el.querySelector('[style*="mso-list:Ignore"],[style*="mso-list: Ignore"]');
        const lbl = ign ? ign.textContent.trim() : '•';
        const kind = /^[\dA-Za-z]{1,4}[.)]$/.test(lbl) ? 'ol' : 'ul';
        const typ = /^[ivx]+[.)]$/.test(lbl) ? 'i' : /^[IVX]+[.)]$/.test(lbl) ? 'I' : /^[a-z][.)]$/.test(lbl) ? 'a' : /^[A-Z][.)]$/.test(lbl) ? 'A' : '1';
        const lid = /l(\d+)/.exec(st['mso-list'])[1];
        p2.num = { id: listNum(typ + lid, kind), lvl };
        delete p2.ind;
      }
      if (/MsoTitle/.test(el.className || '')) p2.style = 'Title';
      else if (/MsoSubtitle/.test(el.className || '')) p2.style = 'Subtitle';
      else if (/MsoQuote/.test(el.className || '')) p2.style = 'Quote';
      if (ctxList && tag === 'LI') p2.num = { id: ctxList.numId, lvl: ctxList.lvl };
      /* a div that only wraps other blocks creates no paragraph of its own */
      const hasBlockKids = Array.from(el.children).some((c) => BLOCK.has(c.tagName) || c.tagName === 'TABLE' || c.tagName === 'UL' || c.tagName === 'OL');
      if (hasBlockKids && (tag === 'TD' || tag === 'TH' || tag === 'BODY' || tag === 'DIV' || tag === 'SECTION' || tag === 'ARTICLE' || tag === 'MAIN' || tag === 'BLOCKQUOTE' || tag === 'LI' || tag === 'FIGURE' || tag === 'FORM' || tag === 'HEADER' || tag === 'FOOTER' || tag === 'ASIDE' || tag === 'NAV' || tag === 'CENTER')) {
        await inline(el, r2, p2, ctxList);
        return;
      }
      const prevPre = opts.pre;
      if (tag === 'PRE') opts.pre = true;
      curPara(p2, r2);
      await inline(el, r2, p2, ctxList);
      opts.pre = prevPre;
      if (cur && cur.runs.length) { const last = cur.runs[cur.runs.length - 1]; if (last.t === 'text') last.text = last.text.replace(/ +$/, ''); if (last.t === 'br' && cur.runs.length > 1) cur.runs.pop(); }
    }
    async function list(el, rPr, pPr, lvl) {
      const kind = el.tagName === 'UL' ? 'ul' : 'ol';
      const numId = listNum(el.getAttribute('type') || '1', kind);
      for (const li of Array.from(el.children)) {
        if (li.tagName === 'LI') { flush(); await block(li, rPr, Object.assign({}, pPr, { num: { id: numId, lvl: Math.min(8, lvl) } }), { numId, lvl: Math.min(8, lvl) }); flush(); }
        else if (li.tagName === 'UL' || li.tagName === 'OL') await list(li, rPr, pPr, lvl + 1);
      }
    }
    async function table(el, rPr) {
      const rowsEl = Array.from(el.querySelectorAll(':scope > tr, :scope > tbody > tr, :scope > thead > tr, :scope > tfoot > tr'));
      if (!rowsEl.length) return null;
      /* grid positions with rowspans */
      const occupied = [];
      const rows = [];
      let nCols = 0;
      const colW = [];
      for (let ri = 0; ri < rowsEl.length; ri++) {
        const tr = rowsEl[ri];
        occupied[ri] = occupied[ri] || [];
        const cells = [];
        let c = 0;
        for (const td of Array.from(tr.children).filter((x) => x.tagName === 'TD' || x.tagName === 'TH')) {
          while (occupied[ri][c]) {
            /* continuation of a rowspan above */
            const above = occupied[ri][c];
            cells.push({ cont: true, span: above.span, c0: c });
            c += above.span;
          }
          const span = Math.max(1, parseInt(td.getAttribute('colspan') || '1', 10) || 1);
          const rs = Math.max(1, parseInt(td.getAttribute('rowspan') || '1', 10) || 1);
          const saveBlocks = blocks.length;
          const tst = styleMap(td);
          flush();
          await block(td, elRPr(td, td.tagName === 'TH' ? Object.assign({}, rPr, { b: true }) : rPr), {}, null);
          flush();
          const inner = blocks.splice(saveBlocks);
          const tc = {};
          if (span > 1) tc.span = span;
          if (rs > 1) tc.vMerge = 'restart';
          const bg = parseColor((tst['background-color'] || tst.background || td.getAttribute('bgcolor') || '').split(' ')[0]);
          if (bg && bg !== 'FFFFFF') tc.shd = { val: 'clear', fill: bg, color: 'auto' };
          const va = tst['vertical-align'] || td.getAttribute('valign');
          if (va === 'middle' || va === 'center') tc.vAlign = 'center'; else if (va === 'bottom') tc.vAlign = 'bottom';
          const w = pt(tst.width) || (td.getAttribute('width') && !/%/.test(td.getAttribute('width')) ? parseFloat(td.getAttribute('width')) * 0.75 : null);
          if (w && span === 1) colW[c] = Math.max(colW[c] || 0, w);
          cells.push({ blocks: inner.length ? inner : [D.para()], tcPr: tc, c0: c, span });
          for (let k = 1; k < rs; k++) { occupied[ri + k] = occupied[ri + k] || []; occupied[ri + k][c] = { span }; }
          c += span;
        }
        while (occupied[ri][c]) { cells.push({ cont: true, span: occupied[ri][c].span, c0: c }); c += occupied[ri][c].span; }
        nCols = Math.max(nCols, c);
        rows.push({ cells, header: tr.parentNode && tr.parentNode.tagName === 'THEAD' });
      }
      if (!nCols) return null;
      const avail = L.tables ? L.tables.availWidth(L.ed.sel ? L.ed.sel.f : D.pos(D.firstPara(d.main.blocks), 0)) : 432;
      let grid = [];
      const known = colW.filter(Boolean).reduce((a, b) => a + b, 0), unknown = Array.from({ length: nCols }, (_, i) => !colW[i]).filter(Boolean).length;
      for (let i = 0; i < nCols; i++) grid.push(colW[i] || Math.max(24, (avail - known) / Math.max(1, unknown)));
      const sum = grid.reduce((a, b) => a + b, 0);
      if (sum > avail) grid = grid.map((g) => (g * avail) / sum);
      grid = grid.map((g) => L.round(g, 2));
      const tbl = D.table([], grid, { style: 'TableGrid', w: { type: 'auto', v: 0 }, look: { firstRow: true, firstCol: true, noVBand: true } });
      const tst = styleMap(el);
      const noBorder = el.getAttribute('border') === '0' || /none/.test(tst.border || '') || (!el.getAttribute('border') && !tst.border && !el.querySelector('td[style*="border"],th[style*="border"]'));
      if (noBorder) tbl.tblPr.style = 'TableNormal';
      for (const r of rows) {
        const cells = r.cells.map((cc) => (cc.cont ? D.cell([D.para()], Object.assign({ vMerge: 'continue' }, cc.span > 1 ? { span: cc.span } : {})) : D.cell(cc.blocks, cc.tcPr)));
        const row = D.row(cells, r.header ? { header: true } : {});
        tbl.rows.push(row);
      }
      if (L.tables) L.tables.fixWidths(tbl);
      return tbl;
    }
    await block(root, opts.baseR || {}, {}, null);
    flush();
    for (const x of pending) await x;
    /* drop empty leading/trailing paragraphs that come from wrapper markup */
    while (blocks.length > 1 && blocks[blocks.length - 1].t === 'p' && !blocks[blocks.length - 1].runs.length) blocks.pop();
    while (blocks.length > 1 && blocks[0].t === 'p' && !blocks[0].runs.length) blocks.shift();
    for (const b of blocks) if (b.t === 'p') D.normalize(b);
    /* a single paragraph pasted is partial (merges into the target paragraph) */
    if (blocks.length === 1 && blocks[0].t === 'p') blocks[0].partial = true;
    if (blocks.length > 1 && blocks[blocks.length - 1].t === 'p' && opts.clipboard !== false) blocks[blocks.length - 1].partial = true;
    if (blocks.length && blocks[blocks.length - 1].t === 'tbl') blocks.push(D.para());
    return blocks;
  };
  /** open an .htm(l) file as a new document */
  H.docFromHTML = async function (html) {
    const d = D.newDoc();
    const prev = D.doc;
    D.doc = d;
    try {
      const dp = new DOMParser();
      const dom = dp.parseFromString(html, 'text/html');
      const t = dom.querySelector('title');
      if (t) d.props.title = t.textContent.trim();
      const blocks = await H.htmlToBlocks(html, { clipboard: false });
      for (const b of blocks) delete b.partial;
      d.main.blocks = blocks.length ? blocks : [D.para()];
      const bg = dom.body && (dom.body.getAttribute('bgcolor') || styleMap(dom.body)['background-color']);
      if (bg) { const c = parseColor(bg); if (c && c !== 'FFFFFF') d.bg = c; }
    } finally { D.doc = prev; }
    return d;
  };
})();
