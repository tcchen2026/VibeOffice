/* Quire — rendering of paragraphs, runs and tables into DOM.
 * Every paragraph fragment is <div class="p" data-pid data-from data-to><div class="pc">…spans…</div></div>.
 * Spans that map to model content carry data-o (start offset) and data-n (length).
 */
(function () {
  'use strict';
  const L = window.L, D = L.D;
  const R = (L.R = {});
  const PX = D.PX;
  R.px = (pt) => Math.round(pt * PX * 100) / 100;

  /* ---------- font metrics (line height of "single" spacing as a multiple of font size) ---------- */
  const LH = {
    'calibri': 1.2207, 'calibri light': 1.2207, 'cambria': 1.1724, 'times new roman': 1.149, 'arial': 1.149, 'courier new': 1.1328, 'verdana': 1.2153, 'georgia': 1.1362,
    'tahoma': 1.2069, 'trebuchet ms': 1.1611, 'segoe ui': 1.3301, 'garamond': 1.1235, 'comic sans ms': 1.3936, 'book antiqua': 1.1704, 'bookman old style': 1.1636, 'century gothic': 1.2251,
    'palatino linotype': 1.1353, 'lucida console': 1.0, 'impact': 1.2197, 'arial black': 1.41, 'arial narrow': 1.149, 'franklin gothic medium': 1.1338, 'symbol': 1.2275, 'wingdings': 1.1, 'aptos': 1.2, 'helvetica': 1.149,
  };
  R.lineFactor = (font) => LH[String(font || '').toLowerCase()] || 1.17;
  const SPACE_EM = { 'times new roman': 0.25, arial: 0.2778, calibri: 0.2261, 'calibri light': 0.2261, cambria: 0.2197, georgia: 0.2412, verdana: 0.3516, tahoma: 0.3125, 'courier new': 0.6, garamond: 0.25, 'book antiqua': 0.25, 'palatino linotype': 0.25, 'segoe ui': 0.2744, 'trebuchet ms': 0.3013, 'century gothic': 0.2773, helvetica: 0.2778, aptos: 0.2, 'arial narrow': 0.2278 };
  R.spaceEm = (font) => SPACE_EM[String(font || '').toLowerCase()] || 0.25;
  R.mathSolo = (p) => {
    let maths = 0, other = '';
    for (const it of p.runs) {
      if (it.t === 'raw' && it.math) maths++;
      else if (it.t === 'text') other += it.text;
      else if (it.t === 'tab' || it.t === 'ptab' || D.ilen(it) === 0 || it.t === 'fb' || it.t === 'fs' || it.t === 'fe') continue;
      else return false;
    }
    return maths === 1 && other.replace(/[\s()[\]0-9.,:;\-–a-zA-Z]/g, '').length === 0 && other.trim().length <= 14;
  };
  /* equations: MathML built once per source XML */
  const mathCache = new Map();
  R.mathFor = (it) => {
    if (!L.omml) return null;
    let m = mathCache.get(it.xml);
    if (m === undefined) {
      try { m = L.omml.toMathML(it.xml); } catch (e) { m = null; }
      if (mathCache.size > 500) mathCache.clear();
      mathCache.set(it.xml, m);
    }
    return m;
  };

  /* ---------- colours ---------- */
  R.HIGHLIGHT = { yellow: '#FFFF00', green: '#00FF00', cyan: '#00FFFF', magenta: '#FF00FF', blue: '#0000FF', red: '#FF0000', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#C0C0C0', black: '#000000', white: '#FFFFFF' };
  R.HIGHLIGHT_NAMES = Object.keys(R.HIGHLIGHT);
  R.col = (c, dflt) => (!c || c === 'auto' ? dflt || '' : '#' + String(c).replace(/^#/, ''));
  R.AUTHOR_COLORS = ['#C00000', '#1F4FBF', '#2E8B37', '#9B3FB0', '#C06000', '#008080', '#7F6000', '#B03060'];
  const authors = [];
  R.authorColor = (name) => { let i = authors.indexOf(name || ''); if (i < 0) { authors.push(name || ''); i = authors.length - 1; } return R.AUTHOR_COLORS[i % R.AUTHOR_COLORS.length]; };
  R.authors = () => authors.slice();
  /* automatic text colour turns white on dark shading (same luminance threshold LibreOffice uses for Word compatibility) */
  R.isDark = (hex) => { if (!hex || hex === 'auto') return false; const [r, g, b] = L.color.hexToRgb('#' + String(hex).replace('#', '')); return (r * 299 + g * 587 + b * 114) / 1000 <= 156; };
  /** 'pct25' style shading over a colour */
  R.shadeCSS = function (shd) {
    if (!shd) return '';
    const fill = shd.fill && shd.fill !== 'auto' ? '#' + shd.fill : null;
    const v = shd.val || 'clear';
    if (v === 'clear' || v === 'nil') return fill || '';
    const fg = shd.color && shd.color !== 'auto' ? '#' + shd.color : '#000000';
    if (v === 'solid') return fg;
    const m = /^pct(\d+)$/.exec(v);
    if (m) return L.color.mix(fg, fill || '#FFFFFF', +m[1] / 100);
    /* patterns approximated by their density */
    const dens = { horzStripe: 0.5, vertStripe: 0.5, reverseDiagStripe: 0.5, diagStripe: 0.5, horzCross: 0.6, diagCross: 0.6, thinHorzStripe: 0.25, thinVertStripe: 0.25, thinReverseDiagStripe: 0.25, thinDiagStripe: 0.25, thinHorzCross: 0.4, thinDiagCross: 0.4 }[v];
    return dens ? L.color.mix(fg, fill || '#FFFFFF', dens) : fill || '';
  };

  /* ---------- borders ---------- */
  R.borderCSS = function (b) {
    if (!b || !b.val || b.val === 'nil' || b.val === 'none') return 'none';
    const w = Math.max(0.75, (b.sz || 0.5) * PX);
    const c = R.col(b.color, '#000000');
    const v = b.val;
    let style = 'solid', width = w;
    if (v === 'double' || /^thinThick|^thickThin|triple/.test(v)) { style = 'double'; width = Math.max(3, w * 3); }
    else if (v === 'dotted') style = 'dotted';
    else if (/dash/.test(v) && !/dot/i.test(v)) style = 'dashed';
    else if (/dotDash|dotDotDash/.test(v)) style = 'dashed';
    else if (v === 'thick') width = Math.max(1.5, w);
    else if (v === 'threeDEmboss' || v === 'outset') style = 'outset';
    else if (v === 'threeDEngrave' || v === 'inset') style = 'inset';
    else if (v === 'wave' || v === 'doubleWave') style = 'solid';
    return `${L.round(width, 2)}px ${style} ${c || '#000'}`;
  };
  R.hasBorder = (b) => !!(b && b.val && b.val !== 'nil' && b.val !== 'none');
  /* Browsers snap border widths to whole device pixels (a 0.5pt rule becomes 1px at 100%), so each table
     row grows by a fraction of a pixel; over a long table that is enough to push rows to another page.
     snapPx(w): the width the browser really draws for a w px border (measured once per width and DPR). */
  const snapCache = new Map();
  R.snapPx = function (w) {
    if (!(w > 0)) return 0;
    const k = Math.round(w * 1000) + '@' + (window.devicePixelRatio || 1);
    let v = snapCache.get(k);
    if (v === undefined) {
      const dv = document.createElement('div');
      dv.style.cssText = `position:absolute;visibility:hidden;left:0;top:0;width:10px;height:0;border-top:${w}px solid #000`;
      document.body.appendChild(dv);
      v = dv.getBoundingClientRect().height;
      dv.remove();
      snapCache.set(k, v);
    }
    return v;
  };
  /** how much taller (px, may be negative) a drawn border is than the border Word lays out */
  R.borderExcess = function (b) {
    if (!R.hasBorder(b)) return 0;
    const v = b.val;
    if (v !== 'single' && v !== 'dotted' && v !== 'dashed' && !/^dash|^dot/.test(v)) return 0;
    const exact = (b.sz || 0.5) * PX;
    const drawn = Math.max(0.75, exact);
    return R.snapPx(L.round(drawn, 2)) - exact;
  };

  /* ---------- run CSS ---------- */
  const rcssCache = new Map();
  R.clearCaches = () => rcssCache.clear();
  /** CSS for resolved run properties. lineMult: line spacing multiple (auto rule) or null for exact */
  R.runCSS = function (r, lineMult, opts) {
    const key = JSON.stringify(r) + '|' + lineMult + '|' + (opts ? opts.key : '');
    let css = rcssCache.get(key);
    if (css != null) return css;
    const s = [];
    const sz = r.sz || 12;
    const vert = r.vert;
    const fsz = vert === 'superscript' || vert === 'subscript' ? sz * 0.65 : sz;
    s.push(`font-family:${L.fontStack(r.font || 'Times New Roman').replace(/"/g, "'")}`);
    s.push(`font-size:${R.px(fsz)}px`);
    /* superscript/subscript text is raised or lowered inside the line Word has already sized: it must not
       stretch the line (a footnote reference would otherwise push its line down a few pixels) */
    if (vert === 'superscript' || vert === 'subscript') s.push('line-height:0');
    else if (lineMult != null) s.push(`line-height:${L.round(R.lineFactor(r.font) * lineMult * (fsz < sz ? sz / fsz : 1), 4)}`);
    if (r.b) s.push('font-weight:bold');
    if (r.i) s.push('font-style:italic');
    const deco = [];
    let decoStyle = '';
    if (r.u && r.u !== 'none') {
      deco.push('underline');
      decoStyle = { double: 'double', dotted: 'dotted', dottedHeavy: 'dotted', dash: 'dashed', dashLong: 'dashed', dashedHeavy: 'dashed', dashLongHeavy: 'dashed', dotDash: 'dashed', dotDotDash: 'dashed', wave: 'wavy', wavyHeavy: 'wavy', wavyDouble: 'wavy' }[r.u] || '';
      if (/thick|Heavy/.test(r.u)) s.push('text-decoration-thickness:2px');
      if (r.uColor && r.uColor !== 'auto') s.push(`text-decoration-color:#${r.uColor}`);
    }
    if (r.strike) deco.push('line-through');
    if (r.dstrike) { deco.push('line-through'); if (!decoStyle) decoStyle = 'double'; }
    if (deco.length) s.push(`text-decoration-line:${deco.join(' ')}`, 'text-decoration-skip-ink:none', 'text-underline-offset:0.12em');
    if (decoStyle) s.push(`text-decoration-style:${decoStyle}`);
    let color = r.color && r.color !== 'auto' ? '#' + r.color : null;
    const bg = r.hl && r.hl !== 'none' ? R.HIGHLIGHT[r.hl] : r.shd ? R.shadeCSS(r.shd) : null;
    if (!color && bg && R.isDark(bg.replace('#', ''))) color = '#FFFFFF';
    if (!color && opts && opts.autoWhite) color = '#FFFFFF';
    if (color) s.push(`color:${color}`);
    if (bg) s.push(`background-color:${bg}`);
    if (r.caps) s.push('text-transform:uppercase');
    if (r.smallCaps) s.push('font-variant:small-caps');
    if (r.spacing) s.push(`letter-spacing:${L.round(r.spacing * PX, 2)}px`);
    if (vert === 'superscript') s.push('vertical-align:super');
    else if (vert === 'subscript') s.push('vertical-align:sub');
    else if (r.pos) s.push(`vertical-align:${L.round(r.pos * PX, 2)}px`);
    if (r.outline) s.push(`-webkit-text-stroke:0.8px ${color || '#000'}`, 'color:#FFFFFF');
    if (r.shadow) s.push('text-shadow:1px 1px 0 #A0A0A0');
    if (r.emboss) s.push('text-shadow:-1px -1px 0 #FFFFFF,1px 1px 0 #808080', 'color:#D0D0D0');
    if (r.imprint) s.push('text-shadow:1px 1px 0 #FFFFFF,-1px -1px 0 #808080', 'color:#D0D0D0');
    if (r.border && R.hasBorder(r.border)) s.push(`border:${R.borderCSS(r.border)}`, 'padding:0 1px');
    if (r.kern === 0) s.push('font-kerning:none');
    if (r.w && r.w !== 100) s.push(`font-stretch:${L.clamp(r.w, 50, 200)}%`);
    css = s.join(';');
    rcssCache.set(key, css);
    return css;
  };

  /* ---------- paragraph helpers ---------- */
  /** line spacing info: {mult} for auto rule or {px, rule} for exact/atLeast */
  R.lineSpec = function (pPr) {
    const sp = pPr.sp || {};
    if (sp.rule === 'exact' && sp.line) return { exact: R.px(sp.line) };
    if (sp.rule === 'atLeast' && sp.line) return { atLeast: R.px(sp.line), mult: 1 };
    return { mult: sp.line && sp.rule !== 'exact' && sp.rule !== 'atLeast' ? sp.line : 1 };
  };
  const ALIGN = { left: 'left', start: 'left', center: 'center', right: 'right', end: 'right', both: 'justify', distribute: 'justify', lowKashida: 'justify', mediumKashida: 'justify', highKashida: 'justify', thaiDistribute: 'justify' };

  /**
   * Render a fragment [from,to) of paragraph p.
   * ctx: {doc, labels, story, first, last, cont (continuation), showMarks, showHidden, view, fieldCodes,
   *       tbl (table context {tblRPr,tblPPr}), prev, next (neighbour paragraphs for border/contextual grouping),
   *       atTop (suppress space before), pageNum, numPages, sectPages, openCmts (Set), markup}
   */
  R.para = function (p, from, to, ctx) {
    const d = ctx.doc;
    const tctx = ctx.tbl || null;
    const pPr = D.pProps(d, p, tctx);
    const n = D.plen(p);
    if (to == null) to = n;
    const first = from === 0, last = to >= n;
    const outer = document.createElement('div');
    outer.className = 'p';
    outer.dataset.pid = p.id;
    outer.dataset.from = from;
    outer.dataset.to = to;
    const pc = document.createElement('div');
    pc.className = 'pc';
    outer.appendChild(pc);
    const ls = R.lineSpec(pPr);
    const markR = D.rProps(d, p, null, tctx);
    /* paragraph box */
    const st = [];
    const ind = pPr.ind || {};
    const l = ind.l || 0, r = ind.r || 0, fl = ind.fl || 0;
    const sp = pPr.sp || {};
    let before = sp.bAuto ? 14 : sp.b || 0, after = sp.aAuto ? 14 : sp.a || 0;
    if (pPr.contextual) {
      const sid = p.pPr.style || 'Normal';
      if (ctx.prev && (ctx.prev.pPr.style || 'Normal') === sid) before = 0;
      if (ctx.next && (ctx.next.pPr.style || 'Normal') === sid) after = 0;
    }
    if (ctx.inCell && ctx.firstInCell && sp.bAuto) before = 0;
    if (ctx.inCell && ctx.lastInCell && sp.aAuto) after = 0;
    if (!first || ctx.atTop) before = 0;
    /* Word (since 2000) doesn't add a paragraph's space before to the previous paragraph's space after:
       the larger of the two separates them — unless the document asks for Word 97's cumulative spacing */
    if (before && ctx.prev && !ctx.prevTbl && !ctx.frameEmpty && ctx.prev.t === 'p' && !R.cumulativeSpacing(d)) {
      const pp = D.pProps(d, ctx.prev, tctx);
      const ps = pp.sp || {};
      let pa = ps.aAuto ? 14 : ps.a || 0;
      if (pp.contextual && (ctx.prev.pPr.style || 'Normal') === (p.pPr.style || 'Normal')) pa = 0;
      before = Math.max(0, before - pa);
    }
    if (!last) after = 0;
    if (before) outer.style.paddingTop = R.px(before) + 'px';
    if (after) outer.style.paddingBottom = R.px(after) + 'px';
    const mirrorInd = false;
    void mirrorInd;
    if (l) st.push(`margin-left:${R.px(l)}px`);
    if (r) st.push(`margin-right:${R.px(r)}px`);
    if (fl && first) st.push(`text-indent:${R.px(fl)}px`);
    const jc = ALIGN[pPr.jc] || 'left';
    if (jc !== 'left') st.push(`text-align:${jc}`);
    if (jc === 'justify' && (!last || pPr.jc === 'distribute')) st.push('text-align-last:' + (pPr.jc === 'distribute' ? 'justify' : 'justify'));
    if (jc === 'justify' && !last) outer.classList.add('jl');
    /* document grid (lines): a line takes whole grid pitches, times the spacing multiple */
    const gridPx = ctx.grid && !pPr.noSnap && !ls.exact && !ls.atLeast ? ctx.grid : 0;
    const gridLine = (rr) => { const nat = R.px((rr.sz || 12) * R.lineFactor(rr.font)); return L.round(gridPx * Math.max(1, Math.ceil(nat / gridPx - 0.02)) * (ls.mult || 1), 3); };
    /* line height strut from the first text run (Word sizes lines by their content, not the mark) */
    const firstText = p.runs.find((it) => it.t === 'text' && it.text && !(it.rPr && it.rPr.hidden && !ctx.showHidden) && !(it.rPr && it.rPr.del && ctx.markup === 'final'));
    const strutR = firstText ? D.rProps(d, p, firstText, tctx) : markR;
    st.push(`font-family:${L.fontStack(strutR.font).replace(/"/g, "'")}`, `font-size:${R.px(strutR.sz || 12)}px`);
    /* Word 2013+ (compatibility mode 15) fits justified lines by shrinking spaces up to 20% */
    if (pPr.jc === 'both' && d.settings && d.settings.compat >= 15) st.push(`word-spacing:${L.round(-0.2 * R.spaceEm(strutR.font) * R.px(strutR.sz || 12), 3)}px`);
    if (ls.exact) st.push(`line-height:${ls.exact}px`);
    else if (ls.atLeast) st.push(`line-height:${Math.max(ls.atLeast, R.px((strutR.sz || 12) * R.lineFactor(strutR.font)))}px`);
    else if (gridPx) st.push(`line-height:${gridLine(strutR)}px`);
    else st.push(`line-height:${L.round(R.lineFactor(strutR.font) * ls.mult, 4)}`);
    /* borders & shading, grouped with identical neighbours */
    const bd = pPr.borders;
    const shd = pPr.shd ? R.shadeCSS(pPr.shd) : '';
    const sameBox = (q) => { if (!q) return false; const qp = D.pProps(d, q, tctx); return JSON.stringify(qp.borders || null) === JSON.stringify(bd || null) && (qp.ind || {}).l === l && (qp.ind || {}).r === r; };
    if (bd && Object.keys(bd).some((k) => R.hasBorder(bd[k]))) {
      const groupPrev = first && sameBox(ctx.prev), groupNext = last && sameBox(ctx.next);
      const side = (k) => (R.hasBorder(bd[k]) ? R.borderCSS(bd[k]) : 'none');
      const spc = (k) => R.px(bd[k] && bd[k].space != null ? bd[k].space : k === 'left' || k === 'right' ? 4 : 1);
      if (!groupPrev && first) { st.push(`border-top:${side('top')}`); if (R.hasBorder(bd.top)) st.push(`padding-top:${spc('top')}px`); }
      else if (groupPrev && R.hasBorder(bd.between)) { st.push(`border-top:${R.borderCSS(bd.between)}`, `padding-top:${spc('between')}px`); }
      if (!groupNext && last) { st.push(`border-bottom:${side('bottom')}`); if (R.hasBorder(bd.bottom)) st.push(`padding-bottom:${spc('bottom')}px`); }
      /* the left border sits at the leftmost of the left and first-line indents */
      if (R.hasBorder(bd.left)) { const lo = Math.min(l, l + Math.min(0, fl)); st.push(`border-left:${side('left')}`, `padding-left:${spc('left') + R.px(l - lo)}px`, `margin-left:${R.px(lo) - spc('left') - 1}px`); }
      if (R.hasBorder(bd.right)) st.push(`border-right:${side('right')}`, `padding-right:${spc('right')}px`, `margin-right:${R.px(r) - spc('right') - 1}px`);
      /* with grouped borders, the spacing between the boxed paragraphs sits inside the box */
      if (groupPrev && before) { outer.style.paddingTop = '0'; st.push(`padding-top:${R.px(before)}px`); }
      if (groupNext && after) { outer.style.paddingBottom = '0'; st.push(`padding-bottom:${R.px(after)}px`); }
    }
    if (shd) { st.push(`background-color:${shd}`); if (R.isDark(shd.replace('#', ''))) outer.classList.add('dk'); }
    if (pPr.bidi) st.push('direction:rtl');
    if (pPr.textDir === 'tbRl' || pPr.textDir === 'btLr') { /* vertical text only meaningful in cells */ }
    pc.style.cssText = st.join(';');
    /* drop cap paragraph: floats beside the following paragraph */
    if (pPr.dropCap && pPr.dropCap.type && pPr.dropCap.type !== 'none' && ctx.view !== 'outline') {
      outer.classList.add('dropcap');
      const lines = pPr.dropCap.lines || 3;
      const nextR = ctx.next ? D.rProps(d, ctx.next, ctx.next.runs.find((x) => x.t === 'text'), tctx) : markR;
      const lhPx = R.px((nextR.sz || 12) * R.lineFactor(nextR.font) * (R.lineSpec(ctx.next ? D.pProps(d, ctx.next, tctx) : pPr).mult || 1));
      outer.style.cssText += `;float:left;height:${lhPx * lines}px;margin-right:${R.px(pPr.dropCap.dist || 0) + 2}px;padding:0;${pPr.dropCap.type === 'margin' ? `margin-left:-${lhPx * lines * 0.7}px;` : ''}`;
      pc.style.lineHeight = lhPx * lines + 'px';
      pc.dataset.dropLines = lines;
    }
    /* list label */
    if (first && ctx.labels) {
      const lb = ctx.labels.get(p.id);
      if (lb && ctx.view !== 'outline-plain') {
        const lv = lb.lv || {};
        const lr = D.mergeRPr(D.mergeRPr(markR, lv.rPr ? Object.assign({}, lv.rPr) : {}), { u: undefined, hl: undefined });
        delete lr.u; delete lr.hl; delete lr.link; delete lr.style; delete lr.ins; delete lr.del; delete lr.border;
        const span = document.createElement('span');
        span.className = 'lbl';
        span.contentEditable = 'false';
        let text = lb.text;
        if (lv.fmt === 'bullet') { text = lb.text; if (lv.picture && L.media.has(lv.picture)) text = ''; }
        span.textContent = text;
        if (lv.fmt === 'bullet' && lv.rPr && /symbol|wingdings/i.test(lv.rPr.font || '')) lr.font = 'Arial';
        span.style.cssText = R.runCSS(lr, ls.mult != null && !ls.exact ? ls.mult : null);
        if (lv.picture && L.media.has(lv.picture)) { const im = document.createElement('img'); im.src = L.media.url(lv.picture); im.style.cssText = 'height:0.75em;vertical-align:baseline'; span.appendChild(im); }
        pc.appendChild(span);
        if (lv.suff !== 'nothing') {
          const t = document.createElement('span');
          t.className = lv.suff === 'space' ? 'lsp' : 'tb lbt';
          t.contentEditable = 'false';
          if (lv.suff === 'space') t.textContent = ' ';
          else t.dataset.al = lv.jc || 'left';
          pc.appendChild(t);
        }
        if (lv.jc === 'right' || lv.jc === 'center') span.dataset.jc = lv.jc;
      }
    }
    R.fillRuns(pc, p, from, to, ctx, pPr, ls, tctx);
    /* empty fragment or a trailing line break needs a filler so the line has height */
    const lastEl = pc.lastChild;
    if (!pc.querySelector('[data-o]:not(.zw)') && !pc.querySelector('.lbl')) {
      const br = document.createElement('br');
      br.className = 'fill';
      pc.appendChild(br);
      if (!firstText) { pc.style.fontFamily = L.fontStack(markR.font).replace(/"/g, "'"); pc.style.fontSize = R.px(markR.sz || 12) + 'px'; if (!ls.exact && !ls.atLeast) pc.style.lineHeight = gridPx ? gridLine(markR) + 'px' : L.round(R.lineFactor(markR.font) * ls.mult, 4); }
    } else if (lastEl && lastEl.nodeName === 'BR' && !lastEl.classList.contains('fill')) {
      const br = document.createElement('br');
      br.className = 'fill';
      pc.appendChild(br);
    }
    if (last) {
      outer.classList.add('last');
      if (p.sect && ctx.view !== 'web') outer.classList.add('has-sect');
      if (p.mark && p.mark.del && ctx.markup !== 'original') outer.classList.add('mdel');
      if (p.mark && p.mark.ins) outer.classList.add('mins');
    }
    if (first) outer.classList.add('first');
    /* paragraph the reader had to add after a trailing table: takes no visible space until used */
    /* (Word also collapses the empty end-of-cell paragraph that follows a nested table) */
    if (!n && (p.synth || (ctx.inCell && ctx.lastInCell && ctx.prevTbl))) { outer.classList.add('synth'); outer.style.padding = '0'; }
    if (pPr.keepNext) outer.dataset.kn = '1';
    if (ctx.markup && ctx.markup !== 'final' && ctx.markup !== 'original' && (p.mark || p.runs.some((it) => it.rPr && (it.rPr.ins || it.rPr.del || it.rPr.chg)) || p.pPr.chg)) outer.classList.add('chg');
    if (pPr.outline != null && pPr.outline < 9) outer.dataset.lvl = pPr.outline + 1;
    if (p.sect && last && ctx.showMarks) {
      const sb = document.createElement('span');
      sb.className = 'sectmark';
      sb.contentEditable = 'false';
      sb.dataset.label = 'Section Break (' + ({ nextPage: 'Next Page', continuous: 'Continuous', evenPage: 'Even Page', oddPage: 'Odd Page', nextColumn: 'Next Column' }[p.sect.type || 'nextPage'] || 'Next Page') + ')';
      pc.appendChild(sb);
    }
    return outer;
  };

  /** fill a paragraph content element with spans for items in [from,to) */
  R.fillRuns = function (pc, p, from, to, ctx, pPr, ls, tctx) {
    const d = ctx.doc;
    const mult = ls.exact ? null : ls.atLeast ? 1 : ls.mult;
    let pos = 0;
    let maxTextSz; /* largest size of runs with visible text (computed when first needed) */
    const open = ctx.openCmts ? new Set(ctx.openCmts) : new Set();
    let fieldDepth = ctx.fieldStack ? ctx.fieldStack.slice() : [];
    const markup = ctx.markup || 'markup';
    const items = p.runs;
    const controls = (p.pPr.sdts || []).slice();
    const controlSpan = el => { const c = controls.at(-1); if (c && !c.converted && c.type === 'checkbox') { el.dataset.control = c.key; el.style.cursor = 'pointer'; el.title = c.name; } };
    let lastSpan = null, lastCss = null;
    for (let idx = 0; idx < items.length; idx++) {
      const it = items[idx];
      const n = D.ilen(it);
      const s = pos, e = pos + n;
      pos = e;
      /* zero-length markers */
      if (n === 0) {
        if (it.t === 'sdts') { controls.push(it.control); lastSpan = null; }
        if (it.t === 'sdte') { const i = controls.findIndex(c => c.key === it.key); if (i >= 0) controls.splice(i, 1); lastSpan = null; }
        if (s < from || s > to) { trackMarker(it, open, fieldDepth); continue; }
        if (s === to && to !== from && to < (ctx._plen != null ? ctx._plen : D.plen(p)) && it.t !== 'ce' && it.t !== 'be' && it.t !== 'fe') { continue; }
        trackMarker(it, open, fieldDepth);
        if (it.t === 'cs' || it.t === 'ce') {
          if (ctx.showComments !== false && d.comments[it.id]) {
            const m = document.createElement('span');
            m.className = 'cmk ' + (it.t === 'cs' ? 'cmk-s' : 'cmk-e');
            m.contentEditable = 'false';
            m.dataset.cid = it.id;
            m.dataset.z = s;
            m.style.color = R.authorColor(d.comments[it.id].author);
            pc.appendChild(m);
            lastSpan = null;
          }
        } else if ((it.t === 'bs' || it.t === 'be') && ctx.showBookmarks && it.name && !/^_/.test(it.name || '')) {
          const m = document.createElement('span');
          m.className = 'bmk ' + (it.t === 'bs' ? 'bmk-s' : 'bmk-e');
          m.contentEditable = 'false';
          m.dataset.z = s;
          pc.appendChild(m);
          lastSpan = null;
        } else if (it.t === 'fb' && ctx.fieldCodes) {
          const m = document.createElement('span');
          m.className = 'fcode';
          m.contentEditable = 'false';
          m.dataset.z = s;
          m.dataset.fid = it.fid;
          m.textContent = '{' + (it.instr || '') + '}';
          pc.appendChild(m);
          lastSpan = null;
        } else if (it.t === 'fb' && it.ff && it.ff.type === 'checkbox') {
          const m = document.createElement('span');
          m.className = 'ffcb';
          m.contentEditable = 'false';
          m.dataset.z = s;
          m.dataset.fid = it.fid;
          const rr = D.rProps(d, p, it, tctx);
          m.style.cssText = R.runCSS(Object.assign({}, rr, { font: 'Segoe UI Symbol', sz: it.ff.size || rr.sz }), mult);
          m.textContent = it.ff.checked ? '\u2612' : '\u2610';
          pc.appendChild(m);
          lastSpan = null;
        } else if (it.t === 'fb') {
          /* dynamic fields rendered from layout state */
          const ty = D.fieldType(it.instr);
          if (ty === 'PAGE' || ty === 'NUMPAGES' || ty === 'SECTIONPAGES' || ty === 'SECTION' || (ty === 'STYLEREF' && ctx.hf)) {
            const f = findFieldEnd(items, idx, it.fid);
            if (f) {
              const resLen = f.resLen;
              const rr = D.rProps(d, p, f.firstRes || it, tctx);
              const m = document.createElement('span');
              m.className = 'fdyn';
              m.contentEditable = 'false';
              m.dataset.o = f.resStart;
              m.dataset.n = resLen;
              m.dataset.fty = ty;
              m.dataset.fid = it.fid;
              m.dataset.fmt = (/\\\*\s*(\w+)/.exec(it.instr || '') || [, ''])[1];
              if (ty === 'STYLEREF') m.dataset.instr = it.instr || '';
              m.textContent = ctx.dyn ? ctx.dyn(ty, m.dataset.fmt, m.dataset.instr) : '#';
              m.style.cssText = R.runCSS(rr, mult);
              if (f.resStart >= from && f.resStart < to || resLen === 0 && f.resStart >= from && f.resStart <= to) pc.appendChild(m);
              /* skip the result items */
              const skipTo = f.endIdx;
              for (let k = idx + 1; k < skipTo; k++) pos += D.ilen(items[k]);
              idx = skipTo - 1;
              lastSpan = null;
              continue;
            }
          }
        }
        continue;
      }
      if (e <= from || s >= to) { continue; }
      const a = Math.max(s, from), b = Math.min(e, to);
      const rp = it.rPr || {};
      /* revisions display */
      if (rp.del && (markup === 'final')) continue;
      if (rp.ins && (markup === 'original')) continue;
      if (fieldDepth.length && ctx.fieldCodes) continue;
      let r = D.rProps(d, p, it, tctx);
      if (r.hidden && !ctx.showHidden) continue;
      if (rp.style === 'CommentReference') continue;
      const cls = [];
      let extra = '';
      if (rp.ins && markup !== 'final' && markup !== 'original') { const c = R.authorColor(rp.ins.author); extra += `;color:${c};text-decoration-line:underline;text-decoration-color:${c}`; cls.push('ins'); }
      if (rp.del && markup !== 'final' && markup !== 'original') { const c = R.authorColor(rp.del.author); extra += `;color:${c};text-decoration-line:line-through;text-decoration-color:${c}`; cls.push('del'); }
      if (rp.chg && markup !== 'final' && markup !== 'original' && ctx.showFmtChanges) cls.push('fchg');
      if (open.size) { cls.push('cmr'); }
      if (r.hidden) cls.push('hid');
      if (fieldDepth.length) cls.push('fr');
      if (r.link || rp.link) cls.push('lnk');
      let el;
      if (it.t === 'text') {
        const text = it.text.slice(a - s, b - s);
        /* Word does not let a run of spaces in a larger font raise the line: draw it at the line's
           text size and keep its width with letter spacing */
        if (/^ +$/.test(text) && r.sz && !r.u && !r.highlight && !r.shd) {
          if (maxTextSz === undefined) { maxTextSz = 0; for (const t2 of p.runs) if (t2.t === 'text' && /\S/.test(t2.text)) maxTextSz = Math.max(maxTextSz, D.rProps(d, p, t2, tctx).sz || 0); }
          if (maxTextSz && r.sz > maxTextSz + 0.5) {
            extra += `;letter-spacing:${L.round(R.px((r.sz - maxTextSz) * R.spaceEm(r.font)), 2)}px`;
            r = Object.assign({}, r, { sz: maxTextSz });
          }
        }
        const css = R.runCSS(r, mult, ctx.dkPara ? { autoWhite: true, key: 'w' } : null) + extra;
        const clsS = cls.join(' ');
        /* merge with the previous span when formatting is identical */
        if (lastSpan && lastCss === css + '|' + clsS && +lastSpan.dataset.o + +lastSpan.dataset.n === a && !ctx.showMarks) {
          lastSpan.firstChild.appendData(text);
          lastSpan.dataset.n = +lastSpan.dataset.n + text.length;
          continue;
        }
        if (ctx.showMarks && text.includes(' ')) {
          /* spaces shown as centred dots: one span per space so widths stay exact */
          let k = 0;
          const parts = text.split(/( +)/);
          for (const part of parts) {
            if (!part) continue;
            const sp = document.createElement('span');
            sp.dataset.o = a + k;
            sp.dataset.n = part.length;
            sp.style.cssText = css;
            if (part[0] === ' ') sp.className = (clsS + ' spc').trim(); else if (clsS) sp.className = clsS;
            sp.appendChild(document.createTextNode(part));
            if (r.link || rp.link) sp.dataset.link = linkStr(r.link || rp.link);
            if (r.noProof) sp.spellcheck = false;
            controlSpan(sp);
            pc.appendChild(sp);
            k += part.length;
          }
          lastSpan = null;
          continue;
        }
        el = document.createElement('span');
        el.appendChild(document.createTextNode(text));
        el.style.cssText = css;
        if (clsS) el.className = clsS;
        el.dataset.o = a;
        el.dataset.n = b - a;
        if (r.link || rp.link) el.dataset.link = linkStr(r.link || rp.link);
        if (r.noProof) el.spellcheck = false;
        if (r.lang && !/^en/i.test(r.lang)) el.lang = r.lang;
        controlSpan(el);
        pc.appendChild(el);
        lastSpan = el; lastCss = css + '|' + clsS;
        continue;
      }
      lastSpan = null;
      if (it.t === 'tab' || it.t === 'ptab') {
        el = document.createElement('span');
        el.className = 'tb' + (cls.length ? ' ' + cls.join(' ') : '');
        el.contentEditable = 'false';
        el.style.cssText = R.runCSS(r, mult) + extra;
        if (it.t === 'ptab') { el.dataset.al = it.al || 'left'; el.dataset.ptab = it.rel || 'margin'; el.dataset.leader = it.leader || 'none'; }
      } else if (it.t === 'br') {
        if (it.type === 'page' || it.type === 'column') {
          el = document.createElement('span');
          el.className = 'hbrk ' + (it.type === 'page' ? 'pbrk' : 'cbrk');
          el.contentEditable = 'false';
          el.dataset.label = it.type === 'page' ? 'Page Break' : 'Column Break';
        } else {
          el = document.createElement('br');
          if (ctx.showMarks) { const m = document.createElement('span'); m.className = 'lbmark'; m.contentEditable = 'false'; m.dataset.z = a; pc.appendChild(m); }
        }
      } else if ((it.t === 'img' || it.t === 'shape' || it.t === 'group' || it.t === 'chart') && it.float) {
        if (!L.drawing) continue;
        el = L.drawing.renderInline(it, r, ctx, p);
      } else if (it.t === 'img') {
        el = R.image(it, r, ctx);
        if (!el) continue;
      } else if (it.t === 'shape' || it.t === 'group' || it.t === 'chart') {
        el = L.drawing && L.drawing.renderInline ? L.drawing.renderInline(it, r, ctx, p) : null;
        if (!el) { el = document.createElement('span'); el.className = 'obj-missing'; el.contentEditable = 'false'; }
      } else if (it.t === 'fn' || it.t === 'en') {
        el = document.createElement('span');
        el.className = 'nref';
        el.contentEditable = 'false';
        const rr = D.mergeRPr(r, { vert: r.vert || 'superscript' });
        el.style.cssText = R.runCSS(rr, mult) + extra;
        el.textContent = it.custom || (ctx.noteNum ? ctx.noteNum(it) : '1');
        el.dataset.note = it.t + ':' + it.id;
      } else if (it.t === 'ruby') {
        /* phonetic guide: annotation set small above the base text */
        el = document.createElement('ruby');
        el.contentEditable = 'false';
        el.className = 'ruby';
        const br = D.mergeRPr(r, it.baseRPr || {});
        el.style.cssText = R.runCSS(br, mult) + extra;
        el.appendChild(document.createTextNode(it.base || ''));
        const rt = document.createElement('rt');
        const hps = it.pr && it.pr.hps ? +it.pr.hps / 2 : null;
        rt.style.cssText = R.runCSS(D.mergeRPr(r, Object.assign({}, it.rtRPr || {}, hps ? { sz: hps } : {})), mult);
        rt.textContent = it.rt || '';
        el.appendChild(rt);
      } else if (it.t === 'sym') {
        el = document.createElement('span');
        const rr = Object.assign({}, r, { font: it.font || r.font });
        el.style.cssText = R.runCSS(rr, mult) + extra;
        el.textContent = L.mapSymbolChar(it.char, it.font) || it.char;
        el.contentEditable = 'false';
        el.className = 'sym';
      } else if (it.t === 'sep' || it.t === 'csep') {
        el = document.createElement('span');
        el.className = it.t === 'sep' ? 'fnsep' : 'fnsep cont';
        el.contentEditable = 'false';
      } else if (it.t === 'raw') {
        el = document.createElement('span');
        el.className = 'raw' + (it.math ? ' math' : '');
        el.contentEditable = 'false';
        const mres = it.math && it.xml ? R.mathFor(it) : null;
        if (mres) {
          /* Office Math rendered as MathML */
          el.style.cssText = R.runCSS(r, mult) + ';font-style:normal;font-weight:normal;text-decoration:none';
          el.classList.add('mml');
          if (mres.display) { el.classList.add('mpara'); el.style.textAlign = mres.jc === 'left' ? 'left' : mres.jc === 'right' ? 'right' : 'center'; }
          const mm = mres.math.cloneNode(true);
          /* an equation alone on its line (perhaps with a tab and a number) is laid out in display style, as Word does */
          if (!mres.display && R.mathSolo(p)) mm.setAttribute('displaystyle', 'true');
          el.appendChild(mm);
        } else {
          el.style.cssText = R.runCSS(r, mult) + (it.math ? ';font-family:"Cambria Math",Caladea,serif;font-style:italic' : '');
          el.textContent = it.text || it.label || '';
          if (it.opaquePlaceholder) { el.style.opacity = '0.65'; el.style.border = '1px dotted currentColor'; el.style.padding = '2px 5px'; }
          if (!it.text && !it.label) el.classList.add('empty');
        }
      } else {
        el = document.createElement('span');
        el.contentEditable = 'false';
      }
      el.dataset.o = a;
      el.dataset.n = 1;
      controlSpan(el);
      if (cls.length && el.classList) for (const c of cls) el.classList.add(c);
      pc.appendChild(el);
    }
    /* trailing zero-width filler so the caret can sit after a trailing non-editable object */
    const lastEl = pc.lastChild;
    if (lastEl && lastEl.nodeType === 1 && (lastEl.contentEditable === 'false' || lastEl.nodeName === 'IMG') && !lastEl.classList.contains('sectmark')) {
      const z = document.createElement('span');
      z.className = 'zw';
      z.dataset.z = to;
      z.appendChild(document.createTextNode('​'));
      pc.appendChild(z);
    }
  };
  function linkStr(l) { return typeof l === 'string' ? l : l.url ? l.url + (l.anchor ? '#' + l.anchor : '') : '#' + (l.anchor || ''); }
  function trackMarker(it, open, fieldDepth) {
    if (it.t === 'cs') open.add(it.id);
    else if (it.t === 'ce') open.delete(it.id);
    else if (it.t === 'fb') fieldDepth.push(it.fid);
    else if (it.t === 'fs') { /* result follows */ const i = fieldDepth.lastIndexOf(it.fid); if (i >= 0) fieldDepth.splice(i, 1); }
    else if (it.t === 'fe') { const i = fieldDepth.lastIndexOf(it.fid); if (i >= 0) fieldDepth.splice(i, 1); }
  }
  /** locate a field's separate/end within one paragraph */
  function findFieldEnd(items, bIdx, fid) {
    let pos = 0;
    for (let k = 0; k <= bIdx; k++) pos += D.ilen(items[k]);
    let sepIdx = -1, resStart = pos, resLen = 0, firstRes = null;
    for (let k = bIdx + 1; k < items.length; k++) {
      const x = items[k];
      if (x.fid === fid && x.t === 'fs') { sepIdx = k; resStart = pos; continue; }
      if (x.fid === fid && x.t === 'fe') return { endIdx: k + 1, resStart: sepIdx >= 0 ? resStart : pos, resLen: sepIdx >= 0 ? resLen : 0, firstRes };
      if (sepIdx >= 0) { resLen += D.ilen(x); if (!firstRes && x.t === 'text') firstRes = x; }
      pos += D.ilen(x);
    }
    return null;
  }

  /* ---------- images ---------- */
  R.image = function (it, r, ctx) {
    const wrap = document.createElement('span');
    wrap.className = 'img';
    wrap.contentEditable = 'false';
    const w = R.px(it.w || 72), hh = R.px(it.h || 72);
    wrap.style.width = w + 'px';
    wrap.style.height = hh + 'px';
    if (it.float && ctx.view !== 'normal' && ctx.view !== 'outline') return null;
    const url = it.media ? L.media.url(it.media) : '';
    if (url) {
      const im = document.createElement('img');
      im.src = url;
      im.alt = it.alt || '';
      im.draggable = false;
      const c = it.crop;
      if (c && (c.l || c.t || c.r || c.b)) {
        const fw = 1 - (c.l || 0) - (c.r || 0), fh = 1 - (c.t || 0) - (c.b || 0);
        im.style.cssText = `position:absolute;width:${w / fw}px;height:${hh / fh}px;left:${(-w * (c.l || 0)) / fw}px;top:${(-hh * (c.t || 0)) / fh}px;max-width:none`;
      } else im.style.cssText = 'width:100%;height:100%';
      const f = [];
      if (it.gray) f.push('grayscale(1)');
      if (it.bw) f.push('grayscale(1) contrast(30)');
      if (it.washout) f.push('brightness(1.7) contrast(0.25)');
      if (it.bright) f.push(`brightness(${1 + it.bright})`);
      if (it.contrast) f.push(`contrast(${1 + it.contrast})`);
      if (f.length) im.style.filter = f.join(' ');
      wrap.appendChild(im);
    } else {
      wrap.classList.add('missing');
      wrap.textContent = it.alt || 'Picture';
    }
    if (it.rot) wrap.style.transform = `rotate(${it.rot}deg)`;
    if (it.flipH || it.flipV) wrap.style.transform = (wrap.style.transform || '') + ` scale(${it.flipH ? -1 : 1},${it.flipV ? -1 : 1})`;
    if (it.border && R.hasBorder(it.border)) wrap.style.outline = R.borderCSS(it.border);
    if (it.link) wrap.dataset.link = linkStr(it.link);
    return wrap;
  };

  /* ---------- tab stops (post-render pass) ---------- */
  /**
   * Resolve widths of tab spans in a rendered paragraph fragment.
   * colLeft: x (px, client coords) of the column's left edge; scale: current zoom.
   */
  /** documents saved with Word 97 spacing rules add space after and space before together */
  R.cumulativeSpacing = (d) => { const c = d.settings && d.settings.compatXML; return !!c && /doNotUseHTMLParagraphAutoSpacing(?![^>]*w:val="(0|false)")/.test(c); };
  R.layoutTabs = function (outer, p, ctx, colLeftClient, scale, colWidthPx) {
    const tabs = outer.querySelectorAll('.tb');
    if (!tabs.length) return;
    const d = ctx.doc;
    const pPr = D.pProps(d, p, ctx.tbl || null);
    const custom = (pPr.tabs || []).filter((t) => t.al !== 'clear' && t.al !== 'bar' && t.al !== 'num');
    const numTabs = (pPr.tabs || []).filter((t) => t.al === 'num');
    /* right/centre-aligned list numbers hang to the left of the number position */
    const lbl = outer.querySelector('.lbl[data-jc]');
    if (lbl) { lbl.style.marginLeft = '0px'; const lw = lbl.getBoundingClientRect().width / scale; lbl.style.marginLeft = -(lbl.dataset.jc === 'right' ? lw : lw / 2) + 'px'; }
    const defTab = R.px(d.settings.defTab || 36);
    const ind = pPr.ind || {};
    const lIndPx = R.px(ind.l || 0), hangPx = R.px(ind.fl || 0) < 0 ? lIndPx : null;
    const rightLimit = colWidthPx - R.px(ind.r || 0);
    const pc = outer.firstChild;
    const anyDecimal = custom.some((t) => t.al === 'decimal');
    for (const tb of tabs) {
      tb.style.width = '0px';
      tb.style.marginRight = '';
      tb.classList.remove('lead-dot', 'lead-hyphen', 'lead-underscore', 'lead-heavy', 'lead-middleDot');
    }
    for (const e of outer.querySelectorAll('.tbsh')) { e.classList.remove('tbsh'); e.style.left = ''; }
    /** width for tab i standing at x (layout px from the column's left); seg(i, decimal) measures the text after it */
    const widthFor = (i, x, seg) => {
      const tb = tabs[i];
      if (tb.dataset.ptab) {
        const al = tb.dataset.al;
        const target = al === 'center' ? colWidthPx / 2 : al === 'right' ? colWidthPx : 0;
        const sw = seg(i, false) / scale;
        const w = al === 'right' ? target - x - sw : al === 'center' ? target - x - sw / 2 : target - x;
        return { w: Math.max(0, w), leader: tb.dataset.leader };
      }
      /* next stop after x */
      let stop = null;
      if (tb.classList.contains('lbt')) for (const t of numTabs) { const px = R.px(t.pos); if (px > x + 0.5) { stop = { px, al: 'left' }; break; } }
      if (!stop) for (const t of custom) { const px = R.px(t.pos); if (px > x + 0.5) { stop = { px, al: t.al, leader: t.leader }; break; } }
      if (hangPx != null && x + 0.5 < hangPx && (!stop || hangPx < stop.px)) stop = { px: hangPx, al: 'left' };
      if (!stop) {
        const lastCustom = custom.length ? R.px(custom[custom.length - 1].pos) : 0;
        const k = Math.floor(Math.max(x, lastCustom) / defTab) + 1;
        let px = k * defTab;
        while (px <= x + 0.5) px += defTab;
        stop = { px, al: 'left' };
      }
      let w, over = 0;
      if (stop.al === 'left' || stop.al === 'start' || !stop.al) w = stop.px - x;
      else {
        const sw = seg(i, stop.al === 'decimal') / scale;
        if (stop.al === 'center') w = stop.px - x - sw / 2;
        else w = stop.px - x - sw;
        if (w < 0) w = 0;
        /* a right/centre/decimal stop set beyond the right indent: Word wraps the text at the indent but
           lets the tabbed text run out to the stop, into the margin (TOC page numbers set this way) */
        if (stop.px > rightLimit + 0.5) {
          const fullW = stop.al === 'decimal' ? seg(i, false) / scale : sw;
          over = Math.max(0, x + w + fullW - rightLimit);
          if (over > stop.px - rightLimit + 0.5) over = 0;
        }
      }
      /* label tab never collapses below a space */
      if (tb.classList.contains('lbt') && w < 1) w = Math.max(defTab - (x % defTab), 4);
      if (x + w > rightLimit + 1 && stop.al === 'left' && !tb.classList.contains('lbt')) w = Math.max(0, Math.min(w, rightLimit - x));
      return { w: Math.max(0, w), leader: stop.leader, over };
    };
    const apply = (i, r) => {
      const tb = tabs[i];
      tb.style.width = L.round(r.w, 2) + 'px';
      if (r.leader && r.leader !== 'none') tb.classList.add('lead-' + r.leader);
      if (r.over > 0.5) {
        /* the line only counts the part that fits; what follows is drawn shifted out to the stop */
        tb.style.marginRight = -L.round(r.over + 0.05, 2) + 'px';
        for (let e = tb.nextSibling; e && !(e.classList && e.classList.contains('tb')); e = e.nextSibling) {
          if (e.nodeType !== 1) continue;
          e.classList.add('tbsh');
          e.style.left = L.round(r.over, 2) + 'px';
        }
      }
    };
    const live = (i, dec) => segWidth(tabs[i], tabs[i + 1], pc, dec);
    /* fast path: read every tab's position once with all tabs collapsed, work out the widths line by line,
       then verify with a single read; only a tab whose line wrapped differently is redone one by one */
    const n = tabs.length;
    let from = 0;
    if (n > 2) {
      const r0 = Array.from(tabs, (tb) => tb.getBoundingClientRect());
      const segs = Array.from(tabs, (_, i) => live(i, false));
      const decs = anyDecimal ? Array.from(tabs, (_, i) => live(i, true)) : null;
      const expect = new Array(n), res = new Array(n);
      let lineTop = null, shift = 0, lastRight = -Infinity;
      for (let i = 0; i < n; i++) {
        const r = r0[i];
        /* a new visual line starts when the tab sits lower, or left of where the previous one ended */
        if (lineTop === null || r.top > lineTop + Math.max(2, r.height * 0.5) || r.left < lastRight - 1) { lineTop = r.top; shift = 0; }
        const x = (r.left - colLeftClient) / scale + shift;
        expect[i] = x;
        res[i] = widthFor(i, x, (k, dec) => (dec ? decs[k] : segs[k]));
        shift += res[i].w - (res[i].over > 0.5 ? res[i].over + 0.05 : 0);
        lastRight = r.left;
      }
      for (let i = 0; i < n; i++) apply(i, res[i]);
      from = n;
      for (let i = 0; i < n; i++) {
        const x = (tabs[i].getBoundingClientRect().left - colLeftClient) / scale;
        if (Math.abs(x - expect[i]) > 0.75) { from = i; break; }
      }
      if (from < n) for (let i = from; i < n; i++) { tabs[i].style.width = '0px'; tabs[i].classList.remove('lead-dot', 'lead-hyphen', 'lead-underscore', 'lead-heavy', 'lead-middleDot'); }
    }
    for (let i = from; i < n; i++) {
      const x = (tabs[i].getBoundingClientRect().left - colLeftClient) / scale;
      apply(i, widthFor(i, x, live));
    }
  };
  /** width of the content that follows a tab, up to the next tab or line end (decimal: up to the decimal point) */
  function segWidth(tb, nextTab, pc, decimal) {
    const range = document.createRange();
    range.setStartAfter(tb);
    if (nextTab) range.setEndBefore(nextTab); else range.setEnd(pc, pc.childNodes.length);
    if (decimal) {
      /* shrink the range to the first '.' */
      const walker = document.createTreeWalker(pc, NodeFilter.SHOW_TEXT);
      let n, started = false;
      while ((n = walker.nextNode())) {
        if (!started) { if (tb.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_FOLLOWING) started = true; else continue; }
        if (nextTab && (nextTab.compareDocumentPosition(n) & Node.DOCUMENT_POSITION_FOLLOWING)) break;
        const k = n.data.indexOf('.');
        if (k >= 0) { range.setEnd(n, k); break; }
      }
    }
    const rects = range.getClientRects();
    if (!rects.length) return 0;
    const top = rects[0].top;
    let x0 = Infinity, x1 = -Infinity;
    for (const r of rects) { if (Math.abs(r.top - top) > r.height * 0.6 && x1 > -Infinity) break; x0 = Math.min(x0, r.left); x1 = Math.max(x1, r.right); }
    return x1 > x0 ? x1 - x0 : 0;
  }

  /* ---------- tables ---------- */
  /** grid map: for each row, list of {cell, c0, span}; vertical merge resolution → rowspan */
  R.tblMap = function (t) {
    const rows = t.rows.map((r) => {
      let c0 = (r.trPr.gridBefore || 0);
      return r.cells.map((cell) => { const span = cell.tcPr.span || 1; const o = { cell, c0, span }; c0 += span; return o; });
    });
    /* rowspans */
    for (let ri = 0; ri < rows.length; ri++) {
      for (const m of rows[ri]) {
        m.rowspan = 1;
        if (m.cell.tcPr.vMerge === 'continue') { m.hidden = true; continue; }
        if (m.cell.tcPr.vMerge === 'restart') {
          let k = ri + 1;
          while (k < rows.length) {
            const below = rows[k].find((x) => x.c0 === m.c0);
            if (below && below.cell.tcPr.vMerge === 'continue') { m.rowspan++; k++; } else break;
          }
        }
      }
    }
    return rows;
  };
  /** table conditional-format types that apply to a cell */
  R.cellConds = function (t, ri, ci, nRows, nCols, look) {
    look = look || {};
    const out = ['wholeTable'];
    /* Word gives the table style's header-row formatting to every repeating header row at the top
       (its own cnfStyle marks them all firstRow), not only to row 0 */
    let hdrRows = 0;
    if (look.firstRow) { hdrRows = 1; const rows = t.rows || []; while (hdrRows < rows.length && rows[hdrRows - 1].trPr && rows[hdrRows - 1].trPr.header && rows[hdrRows].trPr && rows[hdrRows].trPr.header) hdrRows++; }
    const bodyR = ri - hdrRows;
    const bandSizeR = t.tblPr.rowBand || 1, bandSizeC = t.tblPr.colBand || 1;
    if (!look.noVBand) { const cc = ci - (look.firstCol ? 1 : 0); if (cc >= 0 && !(look.lastCol && ci === nCols - 1)) out.push(Math.floor(cc / bandSizeC) % 2 ? 'band2Vert' : 'band1Vert'); }
    if (!look.noHBand && bodyR >= 0 && !(look.lastRow && ri === nRows - 1)) out.push(Math.floor(bodyR / bandSizeR) % 2 ? 'band2Horz' : 'band1Horz');
    if (look.firstCol && ci === 0) out.push('firstCol');
    if (look.lastCol && ci === nCols - 1) out.push('lastCol');
    if (look.firstRow && ri < hdrRows) out.push('firstRow');
    if (look.lastRow && ri === nRows - 1) out.push('lastRow');
    if (look.firstRow && look.firstCol && ri < hdrRows && ci === 0) out.push('nwCell');
    if (look.firstRow && look.lastCol && ri < hdrRows && ci === nCols - 1) out.push('neCell');
    if (look.lastRow && look.firstCol && ri === nRows - 1 && ci === 0) out.push('swCell');
    if (look.lastRow && look.lastCol && ri === nRows - 1 && ci === nCols - 1) out.push('seCell');
    return out;
  };
  /** resolved table-level properties: style + direct */
  R.tblProps = function (d, t) {
    const sp = D.tableStyleProps(d, t.tblPr.style || D.defaultStyleId(d, 'table'));
    const tp = Object.assign({}, sp.tblPr, t.tblPr);
    tp.borders = Object.assign({}, sp.tblPr.borders || {}, t.tblPr.borders || {});
    tp.cellMar = Object.assign({ l: 5.4, r: 5.4, t: 0, b: 0 }, sp.tblPr.cellMar || {}, t.tblPr.cellMar || {});
    return { tp, sp };
  };
  /** per-cell resolved formatting from the table style's conditional formats */
  R.cellStyle = function (d, t, sp, ri, ci, nRows, nCols) {
    const look = t.tblPr.look || { firstRow: true, firstCol: true, noVBand: true };
    const conds = R.cellConds(t, ri, ci, nRows, nCols, look);
    let tcPr = Object.assign({}, sp.tcPr || {}), rPr = Object.assign({}, sp.rPr || {}), pPr = Object.assign({}, sp.pPr || {});
    let borders = {};
    for (const c of conds) {
      const cf = sp.cond && sp.cond[c];
      if (!cf) continue;
      if (cf.tcPr) { tcPr = Object.assign({}, tcPr, cf.tcPr); if (cf.tcPr.borders) borders = Object.assign({}, borders, cf.tcPr.borders); }
      if (cf.rPr) rPr = Object.assign({}, rPr, cf.rPr);
      if (cf.pPr) pPr = D.mergePPr(pPr, cf.pPr);
      if (cf.tblPr && cf.tblPr.borders) borders = Object.assign({}, borders, cf.tblPr.borders);
    }
    return { tcPr, rPr, pPr, borders, conds };
  };
})();
