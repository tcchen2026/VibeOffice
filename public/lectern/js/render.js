/* Lectern — slide renderer.
 * Builds DOM for a slide in a 1pt = 1px coordinate space; callers scale it
 * with CSS transforms (editor zoom, thumbnails, slide show).
 */
(function () {
  'use strict';
  const L = window.L;
  const { h, s } = L;
  const R = (L.render = {});
  let gid = 0;
  const nid = (p) => (p || 'g') + (++gid);

  /* ---------- patterns (8×8 bitmaps, OOXML prst names) ---------- */
  const PATT = {
    pct5: [0x80, 0, 0, 0, 0x08, 0, 0, 0], pct10: [0x80, 0, 0x08, 0, 0x80, 0, 0x08, 0], pct20: [0x88, 0, 0x22, 0, 0x88, 0, 0x22, 0],
    pct25: [0x88, 0x22, 0x88, 0x22, 0x88, 0x22, 0x88, 0x22], pct30: [0xAA, 0x44, 0xAA, 0x11, 0xAA, 0x44, 0xAA, 0x11], pct40: [0xAA, 0x55, 0xAA, 0x51, 0xAA, 0x55, 0xAA, 0x15],
    pct50: [0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55], pct60: [0xEE, 0x55, 0xBB, 0x55, 0xEE, 0x55, 0xBB, 0x55], pct70: [0xEE, 0xBB, 0xEE, 0xBB, 0xEE, 0xBB, 0xEE, 0xBB],
    pct75: [0xEE, 0xFF, 0xBB, 0xFF, 0xEE, 0xFF, 0xBB, 0xFF], pct80: [0xF7, 0xFF, 0x7F, 0xFF, 0xF7, 0xFF, 0x7F, 0xFF], pct90: [0xFF, 0xF7, 0xFF, 0xFF, 0xFF, 0x7F, 0xFF, 0xFF],
    horz: [0xFF, 0, 0, 0, 0xFF, 0, 0, 0], vert: [0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88], ltHorz: [0xFF, 0, 0, 0, 0, 0, 0, 0], ltVert: [0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
    dkHorz: [0xFF, 0xFF, 0, 0, 0xFF, 0xFF, 0, 0], dkVert: [0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC], narHorz: [0xFF, 0, 0xFF, 0, 0xFF, 0, 0xFF, 0], narVert: [0xAA, 0xAA, 0xAA, 0xAA, 0xAA, 0xAA, 0xAA, 0xAA],
    dashHorz: [0xF0, 0, 0, 0, 0x0F, 0, 0, 0], dashVert: [0x80, 0x80, 0x80, 0x80, 0x08, 0x08, 0x08, 0x08],
    dnDiag: [0x88, 0x44, 0x22, 0x11, 0x88, 0x44, 0x22, 0x11], upDiag: [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88],
    ltDnDiag: [0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01], ltUpDiag: [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80],
    dkDnDiag: [0xCC, 0x66, 0x33, 0x99, 0xCC, 0x66, 0x33, 0x99], dkUpDiag: [0x33, 0x66, 0xCC, 0x99, 0x33, 0x66, 0xCC, 0x99],
    wdDnDiag: [0xC1, 0xE0, 0x70, 0x38, 0x1C, 0x0E, 0x07, 0x83], wdUpDiag: [0x83, 0x07, 0x0E, 0x1C, 0x38, 0x70, 0xE0, 0xC1],
    dashDnDiag: [0x88, 0x44, 0x22, 0x11, 0, 0, 0, 0], dashUpDiag: [0x11, 0x22, 0x44, 0x88, 0, 0, 0, 0],
    cross: [0xFF, 0x88, 0x88, 0x88, 0xFF, 0x88, 0x88, 0x88], smGrid: [0xFF, 0x88, 0x88, 0x88, 0xFF, 0x88, 0x88, 0x88], lgGrid: [0xFF, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
    diagCross: [0x81, 0x42, 0x24, 0x18, 0x18, 0x24, 0x42, 0x81], smCheck: [0xCC, 0xCC, 0x33, 0x33, 0xCC, 0xCC, 0x33, 0x33], lgCheck: [0xF0, 0xF0, 0xF0, 0xF0, 0x0F, 0x0F, 0x0F, 0x0F],
    dotGrid: [0xAA, 0, 0x80, 0, 0x80, 0, 0x80, 0], smConfetti: [0x80, 0x08, 0x40, 0x02, 0x10, 0x01, 0x20, 0x04], lgConfetti: [0xB1, 0x30, 0x03, 0x1B, 0xD8, 0xC0, 0x0C, 0x8D],
    horzBrick: [0xFF, 0x80, 0x80, 0x80, 0xFF, 0x08, 0x08, 0x08], diagBrick: [0x80, 0x40, 0x20, 0x10, 0x18, 0x24, 0x42, 0x81],
    solidDmnd: [0x10, 0x38, 0x7C, 0xFE, 0x7C, 0x38, 0x10, 0x00], openDmnd: [0x80, 0x41, 0x22, 0x14, 0x08, 0x14, 0x22, 0x41], dotDmnd: [0x80, 0x00, 0x22, 0x00, 0x08, 0x00, 0x22, 0x00],
    plaid: [0xAA, 0x55, 0xAA, 0x55, 0xF0, 0xF0, 0xF0, 0xF0], sphere: [0x77, 0x98, 0xF8, 0xF8, 0x77, 0x89, 0x8F, 0x8F], weave: [0x88, 0x54, 0x22, 0x45, 0x88, 0x14, 0x22, 0x51],
    divot: [0x00, 0x08, 0x04, 0x08, 0x00, 0x80, 0x40, 0x80], shingle: [0x03, 0x84, 0x48, 0x30, 0x0C, 0x02, 0x01, 0x01], wave: [0x00, 0x18, 0xA4, 0x03, 0x00, 0x18, 0xA4, 0x03],
    trellis: [0xFF, 0x66, 0xFF, 0x99, 0xFF, 0x66, 0xFF, 0x99], zigZag: [0x81, 0x42, 0x24, 0x18, 0x81, 0x42, 0x24, 0x18],
  };
  R.PATTERNS = PATT;
  const pattPath = (bits) => {
    let d = '';
    bits.forEach((row, y) => { for (let x = 0; x < 8; x++) if (row & (0x80 >> x)) d += `M${x},${y}h1v1h-1z`; });
    return d;
  };
  R.patternSVG = (prst, fg, bg, tile) => {
    tile = tile || 6;
    const bits = PATT[prst] || PATT.pct50;
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${tile}" height="${tile}" viewBox="0 0 8 8" shape-rendering="crispEdges"><rect width="8" height="8" fill="${bg}"/><path d="${pattPath(bits)}" fill="${fg}"/></svg>`;
  };
  R.patternURL = (prst, fg, bg, tile) => 'url("data:image/svg+xml;utf8,' + encodeURIComponent(R.patternSVG(prst, fg, bg, tile)) + '")';

  /* ---------- colours & fills ---------- */
  R.col = (c, design, fb) => L.model.resolveColor(c, design, fb);
  /** CSS background declaration for a Fill (backgrounds, table cells, pane swatches) */
  /** tile geometry of a tiled picture fill inside a box {w, h}: the picture at its own size (pixels at its
   *  resolution) scaled by sx/sy, placed from the alignment corner plus the tx/ty offset */
  R.tileGeom = function (fill, box) {
    const o = fill.tileOpts || { sx: 1, sy: 1, tx: 0, ty: 0, algn: 'tl', flip: 'none' };
    const m = L.media.get(fill.media), z = m && m.size;
    const iw = z ? (z.w * 72) / (z.dpiX || 96) : 64, ih = z ? (z.h * 72) / (z.dpiY || 96) : 64;
    const w = Math.max(0.5, iw * (o.sx || 1)), h = Math.max(0.5, ih * (o.sy || 1));
    const a = o.algn || 'tl', bw = (box && box.w) || 0, bh = (box && box.h) || 0;
    let x = o.tx || 0, y = o.ty || 0;
    if (/r$/.test(a)) x += bw - w; else if (!/l$/.test(a)) x += (bw - w) / 2;
    if (/^b/.test(a)) y += bh - h; else if (!/^t/.test(a)) y += (bh - h) / 2;
    return { x, y, w, h, flip: o.flip || 'none' };
  };
  R.fillCSS = function (fill, design, box) {
    if (fill && fill.t === 'bg') fill = R.curBg || { t: 'solid', c: 'bg1' };
    if (!fill || fill.t === 'none') return 'transparent';
    if (fill.t === 'solid') return L.color.rgba(R.col(fill.c, design), fill.a);
    if (fill.t === 'grad') {
      const stops = (fill.stops || []).slice().sort((a, b) => a.p - b.p).map((st) => `${L.color.rgba(R.col(st.c, design), st.a)} ${L.round(st.p * 100, 1)}%`).join(', ');
      const at = fill.focus ? `${L.round(fill.focus[0] * 100, 2)}% ${L.round(fill.focus[1] * 100, 2)}%` : '50% 50%';
      if (fill.path === 'circle' || fill.path === 'shape') return `radial-gradient(circle farthest-side at ${at}, ${stops})`;
      if (fill.path === 'rect') return `radial-gradient(ellipse farthest-corner at ${at}, ${stops})`;
      return `linear-gradient(${L.round((fill.ang || 0) + 90, 2)}deg, ${stops})`;
    }
    if (fill.t === 'img') {
      const u = L.media.url(fill.media);
      if (fill.tile) { const g = R.tileGeom(fill, box); return `url("${u}") ${L.round(g.x, 2)}px ${L.round(g.y, 2)}px / ${L.round(g.w, 2)}px ${L.round(g.h, 2)}px repeat`; }
      if (fill.crop) { const c = fill.crop, kw = 1 - c.l - c.r || 1, kh = 1 - c.t - c.b || 1; return `url("${u}") ${L.round((c.l / (c.l + c.r || 1)) * 100, 3)}% ${L.round((c.t / (c.t + c.b || 1)) * 100, 3)}% / ${L.round(100 / kw, 3)}% ${L.round(100 / kh, 3)}% no-repeat`; }
      if (fill.fillRect) { const c = fill.fillRect, kw = 1 - c.l - c.r || 1, kh = 1 - c.t - c.b || 1; return `url("${u}") ${L.round(c.l + c.r ? (c.l / (c.l + c.r)) * 100 : 0, 3)}% ${L.round(c.t + c.b ? (c.t / (c.t + c.b)) * 100 : 0, 3)}% / ${L.round(kw * 100, 3)}% ${L.round(kh * 100, 3)}% no-repeat`; }
      return `url("${u}") center / 100% 100% no-repeat`;
    }
    if (fill.t === 'patt') return R.patternURL(fill.prst, R.col(fill.fg, design, '#000000'), R.col(fill.bg, design, '#FFFFFF'));
    return 'transparent';
  };
  /** SVG paint for a Fill; adds defs to the given <defs>. Returns {paint, opacity} */
  R.svgPaint = function (fill, design, defs, box) {
    if (fill && fill.t === 'bg') fill = R.curBg || { t: 'solid', c: 'bg1' };
    if (!fill || fill.t === 'none') return { paint: 'none', opacity: 1 };
    if (fill.t === 'solid') return { paint: R.col(fill.c, design), opacity: fill.a == null ? 1 : fill.a };
    if (fill.t === 'grad') {
      const id = nid('gr');
      let g;
      if (fill.path === 'circle' || fill.path === 'rect' || fill.path === 'shape') {
        const [fx, fy] = fill.focus || [0.5, 0.5];
        const far = Math.max(fx, 1 - fx, fy, 1 - fy);
        g = s('radialGradient', { id, cx: L.round(fx * 100, 2) + '%', cy: L.round(fy * 100, 2) + '%', r: L.round((fill.path === 'rect' ? Math.hypot(Math.max(fx, 1 - fx), Math.max(fy, 1 - fy)) : far) * 100, 2) + '%' });
      }
      else {
        /* gradient vector in bounding-box space spanning the whole box (PowerPoint 'scaled' gradients) */
        const a = ((fill.ang || 0) * Math.PI) / 180, c = Math.cos(a), si = Math.sin(a);
        const k = Math.abs(c) + Math.abs(si);
        const dx = c * k, dy = si * k;
        g = s('linearGradient', { id, gradientUnits: 'objectBoundingBox', x1: L.round(0.5 - dx / 2, 4), y1: L.round(0.5 - dy / 2, 4), x2: L.round(0.5 + dx / 2, 4), y2: L.round(0.5 + dy / 2, 4) });
        void box;
      }
      for (const st of (fill.stops || []).slice().sort((a, b) => a.p - b.p)) g.appendChild(s('stop', { offset: L.round(st.p, 4), 'stop-color': R.col(st.c, design), 'stop-opacity': st.a == null ? 1 : st.a }));
      defs.appendChild(g);
      return { paint: `url(#${id})`, opacity: 1 };
    }
    if (fill.t === 'img') {
      const id = nid('im');
      const href = L.media.url(fill.media);
      let p;
      if (fill.tile) {
        const g = R.tileGeom(fill, box);
        const fx = /x/.test(g.flip), fy = /y/.test(g.flip);
        const pw = g.w * (fx ? 2 : 1), ph = g.h * (fy ? 2 : 1);
        p = s('pattern', { id, patternUnits: 'userSpaceOnUse', x: L.round(g.x, 3), y: L.round(g.y, 3), width: L.round(pw, 3), height: L.round(ph, 3) });
        for (const [mx, my] of [[0, 0], [1, 0], [0, 1], [1, 1]]) {
          if ((mx && !fx) || (my && !fy)) continue;
          /* flipped tiles mirror their neighbours */
          const tr = `translate(${mx ? 2 * g.w : 0} ${my ? 2 * g.h : 0}) scale(${mx ? -1 : 1} ${my ? -1 : 1})`;
          p.appendChild(s('image', { href, width: L.round(g.w, 3), height: L.round(g.h, 3), preserveAspectRatio: 'none', transform: tr }));
        }
      } else if (fill.crop) {
        const c = fill.crop, kw = 1 - c.l - c.r || 1, kh = 1 - c.t - c.b || 1;
        p = s('pattern', { id, patternContentUnits: 'objectBoundingBox', width: 1, height: 1 }, s('image', { href, x: L.round(-c.l / kw, 5), y: L.round(-c.t / kh, 5), width: L.round(1 / kw, 5), height: L.round(1 / kh, 5), preserveAspectRatio: 'none' }));
      } else if (fill.fillRect) {
        const c = fill.fillRect;
        p = s('pattern', { id, patternContentUnits: 'objectBoundingBox', width: 1, height: 1 }, s('image', { href, x: L.round(c.l, 5), y: L.round(c.t, 5), width: L.round(1 - c.l - c.r || 1, 5), height: L.round(1 - c.t - c.b || 1, 5), preserveAspectRatio: 'none' }));
      } else p = s('pattern', { id, patternContentUnits: 'objectBoundingBox', width: 1, height: 1 }, s('image', { href, width: 1, height: 1, preserveAspectRatio: 'none' }));
      defs.appendChild(p);
      return { paint: `url(#${id})`, opacity: fill.a == null ? 1 : fill.a };
    }
    if (fill.t === 'patt') {
      const id = nid('pt');
      const p = s('pattern', { id, patternUnits: 'userSpaceOnUse', width: 6, height: 6 },
        s('rect', { width: 6, height: 6, fill: R.col(fill.bg, design, '#FFFFFF') }),
        s('path', { d: pattPath(PATT[fill.prst] || PATT.pct50), fill: R.col(fill.fg, design, '#000000'), transform: 'scale(0.75)', 'shape-rendering': 'crispEdges' }));
      defs.appendChild(p);
      return { paint: `url(#${id})`, opacity: 1 };
    }
    return { paint: 'none', opacity: 1 };
  };
  R.DASH = {
    solid: null, sysDot: [1, 1], sysDash: [3, 1], dot: [1, 2], dash: [4, 3], dashDot: [4, 3, 1, 3], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1],
  };
  R.dashArray = (dash, w) => { const d = R.DASH[dash]; return d ? d.map((v) => L.round(v * Math.max(w, 0.75), 2)).join(' ') : null; };
  function marker(defs, end, color, lw, isHead) {
    if (!end || !end.type || end.type === 'none') return null;
    const id = nid('mk');
    const sz = { sm: 2, med: 3, lg: 5 };
    /* hairlines still get legible heads (PowerPoint sizes heads as if the line were at least ~2pt) */
    const k = Math.max(lw || 0.75, 2) / Math.max(lw || 0.75, 0.1);
    const mw = (sz[end.w || 'med'] || 3) * k, ml = (sz[end.len || 'med'] || 3) * k;
    const m = s('marker', { id, viewBox: '0 0 10 10', refX: end.type === 'arrow' ? 9 : end.type === 'oval' || end.type === 'diamond' ? 5 : 8, refY: 5, markerWidth: ml, markerHeight: mw, orient: isHead ? 'auto-start-reverse' : 'auto', markerUnits: 'strokeWidth' });
    let shape;
    switch (end.type) {
      case 'stealth': shape = s('path', { d: 'M0,0 L10,5 L0,10 L3,5 Z', fill: color }); break;
      case 'diamond': shape = s('path', { d: 'M0,5 L5,0 L10,5 L5,10 Z', fill: color }); break;
      case 'oval': shape = s('circle', { cx: 5, cy: 5, r: 4.5, fill: color }); break;
      case 'arrow': shape = s('path', { d: 'M1,1 L9,5 L1,9', fill: 'none', stroke: color, 'stroke-width': 1.6, 'stroke-linejoin': 'miter' }); break;
      default: shape = s('path', { d: 'M0,0 L10,5 L0,10 Z', fill: color });
    }
    m.appendChild(shape);
    defs.appendChild(m);
    void lw;
    return `url(#${id})`;
  }
  R.shadowCSS = (sh, design) => {
    if (!sh) return '';
    const c = L.color.rgba(R.col(sh.c || '#000000', design), sh.a == null ? 0.5 : sh.a);
    return `drop-shadow(${L.round(sh.dx == null ? 3 : sh.dx, 2)}px ${L.round(sh.dy == null ? 3 : sh.dy, 2)}px ${L.round(sh.blur || 0, 2)}px ${c})`;
  };

  /* ---------- background ---------- */
  R.background = function (fill, design, W, H) {
    const el = h('div', { class: 'sl-bg', style: `position:absolute;left:0;top:0;width:${W}px;height:${H}px` });
    el.style.background = R.fillCSS(fill || { t: 'solid', c: 'bg1' }, design, { w: W, h: H });
    if (fill && fill.t === 'patt') el.style.backgroundRepeat = 'repeat';
    return el;
  };

  /* ---------- text ---------- */
  const ALIGN = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify', justLow: 'justify', thaiDist: 'justify' };
  R.ALIGN = ALIGN;
  function numberText(scheme, n) {
    switch (scheme) {
      case 'arabicParenR': return n + ')';
      case 'arabicParenBoth': return '(' + n + ')';
      case 'arabicPlain': return String(n);
      case 'romanUcPeriod': return L.romanize(n) + '.';
      case 'romanLcPeriod': return L.romanize(n).toLowerCase() + '.';
      case 'romanUcParenR': return L.romanize(n) + ')';
      case 'romanLcParenR': return L.romanize(n).toLowerCase() + ')';
      case 'alphaUcPeriod': return L.alpha(n) + '.';
      case 'alphaLcPeriod': return L.alpha(n).toLowerCase() + '.';
      case 'alphaUcParenR': return L.alpha(n) + ')';
      case 'alphaLcParenR': return L.alpha(n).toLowerCase() + ')';
      case 'alphaLcParenBoth': return '(' + L.alpha(n).toLowerCase() + ')';
      case 'circleNumDbPlain': return String.fromCharCode(0x245F + Math.min(n, 20));
      default: return n + '.';
    }
  }
  R.numberText = numberText;
  function lineHeight(ps, scale) {
    const ls = ps.lnSpc || { pct: 100 };
    if (ls.pts != null) return { css: `${L.round(ls.pts * scale, 2)}px`, factor: null };
    const f = Math.max(0.3, ((ls.pct == null ? 100 : ls.pct) / 100) * 1.2);
    return { css: String(L.round(f, 3)), factor: f };
  }
  function spacing(sp, sz) {
    if (!sp) return 0;
    if (sp.pts != null) return sp.pts;
    return ((sp.pct || 0) / 100) * sz * 1.2;
  }
  /** Apply run formatting to a span; returns effective size */
  R.runStyle = function (span, rp, design, ctx, scale) {
    const sz = (rp.sz || 18) * scale;
    const st = span.style;
    st.fontFamily = L.fontStack(L.style.font(rp.font, design));
    st.fontSize = L.round(sz, 2) + 'px';
    st.fontWeight = rp.b ? '700' : '400';
    st.fontStyle = rp.i ? 'italic' : 'normal';
    const deco = [];
    if (rp.u && rp.u !== 'none') deco.push('underline');
    if (rp.strike && rp.strike !== 'noStrike') deco.push('line-through');
    let color = R.col(rp.color || 'tx1', design);
    if (rp.link) { color = R.col('hlink', design); if (!deco.includes('underline')) deco.push('underline'); }
    st.color = color;
    if (deco.length) { st.textDecoration = deco.join(' '); if (rp.u === 'dbl') st.textDecorationStyle = 'double'; else if (rp.u === 'dotted' || rp.u === 'dottedHeavy') st.textDecorationStyle = 'dotted'; else if (rp.u === 'dash' || rp.u === 'dashHeavy') st.textDecorationStyle = 'dashed'; else if (rp.u === 'wavy') st.textDecorationStyle = 'wavy'; }
    if (rp.base) { st.verticalAlign = rp.base > 0 ? 'super' : 'sub'; st.fontSize = L.round(sz * 0.7, 2) + 'px'; }
    if (rp.spc) st.letterSpacing = L.round(rp.spc * scale, 2) + 'px';
    if (rp.cap === 'all') st.textTransform = 'uppercase';
    else if (rp.cap === 'small') st.fontVariant = 'small-caps';
    const shadows = [];
    if (rp.shd) {
      /* the file's own shadow when there is one; a shadow fades with semi-transparent text, as in PowerPoint */
      const x = rp.shdX, ta = rp.fill && rp.fill.t === 'solid' && rp.fill.a != null ? rp.fill.a : 1;
      if (x) shadows.push(`${L.round((x.dx || 0) * scale, 2)}px ${L.round((x.dy || 0) * scale, 2)}px ${L.round((x.blur || 0) * scale * 0.5, 2)}px ${L.color.rgba(R.col(x.c || '#000000', design), (x.a == null ? 0.5 : x.a) * ta)}`);
      else shadows.push(`${L.round(sz * 0.06, 2)}px ${L.round(sz * 0.06, 2)}px 0 rgba(0,0,0,${L.round(0.35 * ta, 3)})`);
    }
    if (rp.emb) shadows.push(`-1px -1px 0 rgba(255,255,255,0.7), 1px 1px 0 rgba(0,0,0,0.4)`);
    if (shadows.length) st.textShadow = shadows.join(',');
    if (rp.hl) st.backgroundColor = R.col(rp.hl, design);
    if (rp.fill && rp.fill.t === 'grad') {
      st.backgroundImage = R.fillCSS(rp.fill, design);
      st.webkitBackgroundClip = 'text'; st.backgroundClip = 'text'; st.color = 'transparent';
    } else if (rp.fill && rp.fill.t === 'solid') st.color = L.color.rgba(R.col(rp.fill.c, design), rp.fill.a);
    else if (rp.fill && rp.fill.t === 'none') st.color = 'transparent';
    if (rp.ln && rp.ln.t !== 'none' && rp.ln.c) st.webkitTextStroke = `${L.round(Math.max(0.5, (rp.ln.w || 0.75) * scale), 2)}px ${R.col(rp.ln.c, design)}`;
    return sz;
  };
  function fieldText(r, ctx) {
    if (r.fld === 'slidenum') return String(ctx.num != null ? ctx.num : '‹#›');
    /* the standard date formats update; a custom one (think-cell writes datetime'2''0''2''1' to pin literal
       text) keeps the text saved with the file */
    if (r.fld && /^datetime([1-9]|1[0-3])?$/.test(r.fld)) return L.fmtDate(new Date(), r.fld === 'datetime' ? 'datetime1' : r.fld);
    return r.t || '';
  }
  function appendText(el, text) {
    const parts = String(text).split('\n');
    parts.forEach((t, i) => { if (i) el.appendChild(h('br')); if (t) el.appendChild(document.createTextNode(t)); });
  }
  /**
   * Render paragraphs into a container. Returns the container.
   * opts: {scale, prompt(string), editMode}
   */
  R.paragraphs = function (sh, tx, design, ctx, container, opts) {
    opts = opts || {};
    const scale = (tx.fontScale || 1) * (opts.scale || 1);
    const lnRed = tx.lnSpcRed || 0;
    const counters = [];
    const prevNum = [];
    const paras = opts.promptParas || tx.ps;
    paras.forEach((p, pi) => {
      const ps = L.style.para(sh, p, design);
      const el = h('p', { class: 'pa' });
      el.dataset.lvl = p.lvl || 0;
      el.dataset.pi = pi;
      if (p.pp && Object.keys(p.pp).length) el.dataset.pp = JSON.stringify(p.pp);
      const text = L.txt.paraText(p);
      const firstRun = p.rs[0] ? L.style.run(ps, p.rs[0]) : L.style.run(ps, p.end || {});
      const fsz = (firstRun.sz || 18) * scale;
      el.style.textAlign = ALIGN[ps.algn] || 'left';
      el.style.paddingLeft = L.round((ps.marL || 0) * scale, 2) + 'px';
      el.style.textIndent = L.round((ps.indent || 0) * scale, 2) + 'px';
      el.style.fontSize = L.round(fsz, 2) + 'px';
      el.style.fontFamily = L.fontStack(L.style.font(firstRun.font, design));
      const lh = lineHeight(ps, scale);
      el.style.lineHeight = lh.factor ? String(L.round(lh.factor * (1 - lnRed), 3)) : lh.css;
      if (pi > 0) el.style.marginTop = L.round(spacing(ps.spcBef, fsz / scale) * scale, 2) + 'px';
      const after = spacing(ps.spcAft, fsz / scale) * scale;
      if (after) el.style.marginBottom = L.round(after, 2) + 'px';
      /* bullets & numbering */
      const lvl = p.lvl || 0;
      counters.length = Math.max(counters.length, lvl + 1);
      for (let k = lvl + 1; k < counters.length; k++) { counters[k] = 0; prevNum[k] = null; }
      const bu = ps.bu || { t: 'none' };
      let bullet = null;
      if (bu.t === 'char' && bu.ch) bullet = L.mapSymbolChar(bu.ch, bu.font);
      if (bu.t === 'num') {
        const sch = bu.scheme || 'arabicPeriod';
        counters[lvl] = prevNum[lvl] === sch && counters[lvl] ? counters[lvl] + 1 : bu.start || 1;
        prevNum[lvl] = sch;
        bullet = numberText(sch, counters[lvl]);
      } else if (text.length || opts.promptParas) { counters[lvl] = 0; prevNum[lvl] = null; }
      if (bullet) {
        el.dataset.bu = bullet;
        const bc = bu.c ? R.col(bu.c, design) : R.col(firstRun.color || 'tx1', design);
        el.style.setProperty('--bu-c', bc);
        el.style.setProperty('--bu-s', L.round(fsz * (bu.sz || 1), 2) + 'px');
        el.style.setProperty('--bu-w', L.round(Math.max(-(ps.indent || 0) * scale, fsz * 0.55), 2) + 'px');
        if (bu.font && !/wingdings|symbol/i.test(bu.font)) el.style.setProperty('--bu-f', L.fontStack(bu.font));
      }
      if (!text.length) el.classList.add('empty');
      /* runs */
      if (!p.rs.length) {
        el.appendChild(h('br'));
      } else {
        for (const r of p.rs) {
          const rp = L.style.run(ps, r);
          const sp = h('span', { class: 'r' });
          const own = L.txt.runProps(r);
          if (Object.keys(own).length) sp.dataset.r = JSON.stringify(own);
          if (r.fld) { sp.dataset.fld = r.fld; sp.contentEditable = 'false'; }
          R.runStyle(sp, rp, design, ctx, scale);
          if (r.link) sp.dataset.link = JSON.stringify(r.link);
          appendText(sp, r.fld ? fieldText(r, ctx) : r.t);
          el.appendChild(sp);
        }
        if (/\n$/.test(L.txt.paraText(p))) el.appendChild(h('br'));
      }
      container.appendChild(el);
    });
    return container;
  };

  /** Text body positioned inside the shape's text rectangle */
  R.textBody = function (sh, tx, design, ctx, rect, opts) {
    opts = opts || {};
    const ins = tx.ins || [7.2, 3.6, 7.2, 3.6];
    let [x0, y0, x1, y1] = rect || [0, 0, sh.w, sh.h];
    let bw = Math.max(0, x1 - x0 - ins[0] - ins[2]), bh = Math.max(0, y1 - y0 - ins[1] - ins[3]);
    const outer = h('div', { class: 'tx' });
    const st = outer.style;
    st.left = L.round(x0 + ins[0], 2) + 'px';
    st.top = L.round(y0 + ins[1], 2) + 'px';
    st.width = L.round(bw, 2) + 'px';
    st.height = L.round(bh, 2) + 'px';
    st.justifyContent = tx.anchor === 'ctr' ? 'center' : tx.anchor === 'b' ? 'flex-end' : 'flex-start';
    if (tx.vert === 'vert' || tx.vert === 'vert270' || tx.vert === 'eaVert') {
      /* rotate the whole text frame like PowerPoint's rotated text */
      st.left = L.round(x0 + ins[0] + (bw - bh) / 2, 2) + 'px';
      st.top = L.round(y0 + ins[1] + (bh - bw) / 2, 2) + 'px';
      st.width = L.round(bh, 2) + 'px';
      st.height = L.round(bw, 2) + 'px';
      st.transform = `rotate(${tx.vert === 'vert270' ? -90 : 90}deg)`;
    }
    if (tx.rot) st.transform = (st.transform || '') + ` rotate(${tx.rot}deg)`;
    const inner = h('div', { class: 'txi' });
    if (tx.wrap === false) inner.classList.add('nowrap');
    if (tx.anchorCtr) inner.style.alignSelf = 'center';
    if (tx.cols > 1) { inner.style.columnCount = tx.cols; inner.style.columnGap = L.round(tx.colGap || 0, 2) + 'px'; }
    R.paragraphs(sh, tx, design, ctx, inner, opts);
    outer.appendChild(inner);
    return outer;
  };

  /* ---------- shapes ---------- */
  function geomSVG(sh, design, ctx, opts) {
    const w = Math.max(sh.w, 0.01), hh = Math.max(sh.h, 0.01);
    /* a zero-height box is never painted, so the SVG viewport is at least 1pt (unscaled; content overflows visibly) */
    const vw = L.round(Math.max(w, 1), 3), vh = L.round(Math.max(hh, 1), 3);
    const svg = s('svg', { class: 'geo', width: vw, height: vh, viewBox: `0 0 ${vw} ${vh}`, overflow: 'visible' });
    const defs = s('defs');
    svg.appendChild(defs);
    const parts = L.geom.parts(sh, w, hh);
    const paint = R.svgPaint(sh.fill, design, defs, { w, h: hh });
    const ln = sh.line && sh.line.t !== 'none' && sh.line.c ? sh.line : null;
    const lc = ln ? L.color.rgba(R.col(ln.c, design), ln.a) : 'none';
    const lw = ln ? Math.max(ln.w == null ? 0.75 : ln.w, 0.1) : 0;
    const g = s('g');
    if (sh.flipH || sh.flipV) g.setAttribute('transform', `translate(${sh.flipH ? w : 0},${sh.flipV ? hh : 0}) scale(${sh.flipH ? -1 : 1},${sh.flipV ? -1 : 1})`);
    const hasFill = paint.paint !== 'none';
    for (const part of parts) {
      const filled = part.fill !== 'none';
      const p = s('path', { d: part.d });
      if (part.rule) p.setAttribute('fill-rule', part.rule);
      p.setAttribute('fill', filled ? paint.paint : 'none');
      if (filled && paint.opacity < 1) p.setAttribute('fill-opacity', paint.opacity);
      if (part.stroke && ln) {
        p.setAttribute('stroke', lc);
        p.setAttribute('stroke-width', lw);
        const da = R.dashArray(ln.dash, lw);
        if (da) p.setAttribute('stroke-dasharray', da);
        p.setAttribute('stroke-linejoin', ln.join === 'round' ? 'round' : ln.join === 'bevel' ? 'bevel' : 'miter');
        if (ln.cap === 'rnd' || ln.dash === 'sysDot' || ln.dash === 'dot') p.setAttribute('stroke-linecap', 'round');
        if (part.open) {
          const mh = marker(defs, ln.head, lc, lw, true), mt = marker(defs, ln.tail, lc, lw, false);
          if (mh) p.setAttribute('marker-start', mh);
          if (mt) p.setAttribute('marker-end', mt);
        }
      } else p.setAttribute('stroke', 'none');
      if (opts.interactive) p.setAttribute('pointer-events', filled && (hasFill || sh.tx) ? 'visiblePainted' : 'stroke');
      g.appendChild(p);
      if (filled && hasFill && part.fill !== 'norm') {
        const o = { darken: ['#000', 0.4], darkenLess: ['#000', 0.2], lighten: ['#fff', 0.4], lightenLess: ['#fff', 0.2] }[part.fill];
        if (o) g.appendChild(s('path', { d: part.d, fill: o[0], 'fill-opacity': o[1] * (paint.opacity || 1), stroke: 'none', 'pointer-events': 'none' }));
      }
      if (opts.interactive && part.stroke && (!filled || !hasFill)) {
        g.appendChild(s('path', { d: part.d, fill: 'none', stroke: 'transparent', 'stroke-width': Math.max(8, lw + 6), class: 'hitpath', 'pointer-events': 'stroke' }));
      }
    }
    svg.appendChild(g);
    if (sh.shadow) svg.style.filter = R.shadowCSS(sh.shadow, design);
    return svg;
  }

  function lineSVG(sh, design, ctx, opts) {
    const w = Math.max(sh.w, 0.01), hh = Math.max(sh.h, 0.01);
    /* a zero-height box is never painted, so the SVG viewport is at least 1pt (unscaled; content overflows visibly) */
    const vw = L.round(Math.max(w, 1), 3), vh = L.round(Math.max(hh, 1), 3);
    const svg = s('svg', { class: 'geo', width: vw, height: vh, viewBox: `0 0 ${vw} ${vh}`, overflow: 'visible' });
    const defs = s('defs');
    svg.appendChild(defs);
    const ln = sh.line && sh.line.t !== 'none' ? sh.line : { c: 'tx1', w: 0.75 };
    const lc = L.color.rgba(R.col(ln.c || 'tx1', design), ln.a);
    const lw = Math.max(ln.w == null ? 0.75 : ln.w, 0.1);
    const d = L.geom.linePath(sh);
    const p = s('path', { d, fill: 'none', stroke: sh.line && sh.line.t === 'none' ? 'none' : lc, 'stroke-width': lw });
    const da = R.dashArray(ln.dash, lw);
    if (da) p.setAttribute('stroke-dasharray', da);
    if (ln.cap === 'rnd') p.setAttribute('stroke-linecap', 'round');
    const mh = marker(defs, ln.head, lc, lw, true), mt = marker(defs, ln.tail, lc, lw, false);
    if (mh) p.setAttribute('marker-start', mh);
    if (mt) p.setAttribute('marker-end', mt);
    p.setAttribute('pointer-events', 'none');
    svg.appendChild(p);
    if (opts.interactive) svg.appendChild(s('path', { d, fill: 'none', stroke: 'transparent', 'stroke-width': Math.max(9, lw + 6), 'pointer-events': 'stroke', class: 'hitpath' }));
    if (sh.shadow) svg.style.filter = R.shadowCSS(sh.shadow, design);
    return svg;
  }

  function pictureEl(sh, design, opts) {
    const wrap = h('div', { class: 'pic' });
    const c = sh.crop || { l: 0, t: 0, r: 0, b: 0 };
    const iw = sh.w / Math.max(0.01, 1 - c.l - c.r), ih = sh.h / Math.max(0.01, 1 - c.t - c.b);
    const img = h('img', { src: L.media.url(sh.img && sh.img.view && L.media.has(sh.img.view) ? sh.img.view : sh.media), alt: sh.alt || '', draggable: 'false' });
    img.style.cssText = `position:absolute;left:${L.round(-c.l * iw, 2)}px;top:${L.round(-c.t * ih, 2)}px;width:${L.round(iw, 2)}px;height:${L.round(ih, 2)}px;max-width:none`;
    const f = [];
    const im = sh.img || {};
    if (im.mode === 'gray') f.push('grayscale(1)');
    if (im.mode === 'bw') f.push('grayscale(1) contrast(12)');
    if (im.mode === 'wash') f.push('brightness(1.55) contrast(0.32)');
    if (im.bright) f.push(`brightness(${L.round(1 + im.bright, 3)})`);
    if (im.contrast) f.push(`contrast(${L.round(1 + im.contrast, 3)})`);
    if (f.length) img.style.filter = f.join(' ');
    if (im.alpha != null && im.alpha < 1) img.style.opacity = im.alpha;
    if (sh.flipH || sh.flipV) img.style.transform = `scale(${sh.flipH ? -1 : 1},${sh.flipV ? -1 : 1})`;
    if (!L.media.has(sh.media)) {
      wrap.classList.add('missing');
      wrap.appendChild(h('div', { class: 'pic-missing', text: sh.missingLabel || 'Picture' }));
    } else {
      /* formats browsers can't decode (TIFF, HD Photo, …) degrade to a labelled frame; the original bytes still save */
      img.addEventListener('error', () => {
        if (!wrap.isConnected && !wrap.parentNode) return;
        const m = L.media.get(sh.media);
        const ext = m ? L.mimeToExt(m.type) : '';
        img.remove();
        wrap.classList.add('missing');
        wrap.appendChild(h('div', { class: 'pic-missing', text: ext && ext !== 'png' ? `Picture (.${ext})` : 'Picture' }));
      }, { once: true });
      wrap.appendChild(img);
    }
    if (sh.geom && sh.geom !== 'rect') {
      const parts = L.geom.parts(sh, sh.w, sh.h).filter((p) => p.fill !== 'none');
      if (parts.length) wrap.style.clipPath = `path('${parts.map((p) => p.d).join(' ')}')`;
    }
    void design; void opts;
    return wrap;
  }

  /* table */
  function tableEl(sh, design, ctx, opts) {
    const t = sh.tbl;
    const tbl = h('table', { class: 'tbl' });
    tbl.style.width = L.round(t.cols.reduce((a, b) => a + b, 0), 2) + 'px';
    const cg = h('colgroup');
    for (const w of t.cols) cg.appendChild(h('col', { style: `width:${L.round(w, 2)}px` }));
    tbl.appendChild(cg);
    const bdCSS = (b) => {
      if (!b || b.t === 'none' || !b.c) return 'none';
      const w = Math.max(0.5, b.w == null ? 1 : b.w);
      const style = b.dash && b.dash !== 'solid' ? (/dot/i.test(b.dash) && !/dash/i.test(b.dash) ? 'dotted' : 'dashed') : 'solid';
      return `${L.round(w, 2)}px ${style} ${L.color.rgba(R.col(b.c, design), b.a)}`;
    };
    t.rows.forEach((row, ri) => {
      const tr = h('tr');
      tr.style.height = L.round(row.h, 2) + 'px';
      row.cells.forEach((cell, ci) => {
        if (cell.hm || cell.vm) return;
        const td = h('td');
        td.dataset.r = ri; td.dataset.c = ci;
        if (cell.gs > 1) td.colSpan = cell.gs;
        if (cell.rs > 1) td.rowSpan = cell.rs;
        const ins = (cell.tx && cell.tx.ins) || [7.2, 3.6, 7.2, 3.6];
        td.style.padding = `${L.round(ins[1], 2)}px ${L.round(ins[2], 2)}px ${L.round(ins[3], 2)}px ${L.round(ins[0], 2)}px`;
        td.style.verticalAlign = cell.tx && cell.tx.anchor === 'ctr' ? 'middle' : cell.tx && cell.tx.anchor === 'b' ? 'bottom' : 'top';
        td.style.background = R.fillCSS(cell.fill, design);
        const bd = cell.bd || {};
        td.style.borderLeft = bdCSS(bd.l); td.style.borderRight = bdCSS(bd.r); td.style.borderTop = bdCSS(bd.t); td.style.borderBottom = bdCSS(bd.b);
        const inner = h('div', { class: 'txi' });
        R.paragraphs(sh, cell.tx || L.txt.body(), design, ctx, inner, {});
        td.appendChild(inner);
        tr.appendChild(td);
      });
      tbl.appendChild(tr);
    });
    void opts;
    return tbl;
  }

  /* WordArt (2003-style stretched text, optional warp) */
  R.WARPS = [
    ['textPlain', 'Plain Text'], ['textArchUp', 'Arch Up'], ['textArchDown', 'Arch Down'], ['textCircle', 'Circle'], ['textWave1', 'Wave'],
    ['textSlantUp', 'Slant Up'], ['textSlantDown', 'Slant Down'], ['textTriangle', 'Triangle Up'], ['textChevron', 'Chevron Up'], ['textCanUp', 'Can Up'],
  ];
  function warpPoints(warp, w, hh, fs) {
    const pts = [];
    const N = 64;
    const m = fs * 0.15;
    for (let i = 0; i <= N; i++) {
      const t = i / N;
      switch (warp) {
        case 'textArchUp': { const a = Math.PI - t * Math.PI; pts.push([w / 2 + Math.cos(a) * (w / 2 - m), hh - m - Math.sin(a) * (hh - fs - m)]); break; }
        case 'textArchDown': { const a = Math.PI + t * Math.PI; pts.push([w / 2 + Math.cos(a) * (w / 2 - m), fs + m - Math.sin(a) * (hh - fs - m * 2)]); break; }
        case 'textCircle': { const a = Math.PI + t * Math.PI * 2 * 0.98; pts.push([w / 2 + Math.cos(a) * (w / 2 - fs), hh / 2 + Math.sin(a) * (hh / 2 - fs)]); break; }
        case 'textWave1': pts.push([m + t * (w - 2 * m), hh / 2 + fs / 3 + Math.sin(t * Math.PI * 2) * (hh / 2 - fs * 0.8)]); break;
        case 'textSlantUp': pts.push([m + t * (w - 2 * m), hh - m - t * (hh - fs - m)]); break;
        case 'textSlantDown': pts.push([m + t * (w - 2 * m), fs + t * (hh - fs - m)]); break;
        case 'textTriangle': pts.push([m + t * (w - 2 * m), t < 0.5 ? hh - m - t * 2 * (hh - fs - m) : fs + (t - 0.5) * 2 * (hh - fs - m)]); break;
        case 'textChevron': pts.push([m + t * (w - 2 * m), t < 0.5 ? hh * 0.75 - t * 2 * hh * 0.35 : hh * 0.4 + (t - 0.5) * 2 * hh * 0.35]); break;
        case 'textCanUp': { pts.push([m + t * (w - 2 * m), hh * 0.85 - Math.sin(t * Math.PI) * hh * 0.25]); break; }
        default: pts.push([t * w, hh / 2]);
      }
    }
    return pts;
  }
  R.wordart = function (sh, design, ctx) {
    const wa = sh.wa || { text: 'LettersArt' };
    const w = Math.max(sh.w, 1), hh = Math.max(sh.h, 1);
    const svg = s('svg', { class: 'wa', width: L.round(w, 2), height: L.round(hh, 2), viewBox: `0 0 ${L.round(w, 3)} ${L.round(hh, 3)}`, preserveAspectRatio: 'none', overflow: 'visible' });
    const defs = s('defs');
    svg.appendChild(defs);
    const fill = R.svgPaint(wa.fill || { t: 'solid', c: 'accent1' }, design, defs, { w, h: hh });
    const ln = wa.line && wa.line.t !== 'none' && wa.line.c ? wa.line : null;
    const lines = String(wa.text || '').split('\n');
    const family = L.fontStack(wa.font || 'Arial Black');
    const common = { 'font-family': family, 'font-weight': wa.b ? 700 : 400, 'font-style': wa.i ? 'italic' : 'normal', fill: fill.paint, 'fill-opacity': fill.opacity };
    if (ln) Object.assign(common, { stroke: R.col(ln.c, design), 'stroke-width': Math.max(0.25, ln.w || 0.75), 'paint-order': 'stroke', 'stroke-linejoin': 'round' });
    if (wa.shadow) svg.style.filter = R.shadowCSS(wa.shadow, design);
    if (wa.cap) svg.style.textTransform = 'uppercase';
    if (wa.spc) common['letter-spacing'] = wa.spc;
    const warp = wa.warp || 'textPlain';
    if (warp === 'textPlain' || !lines.length || lines.length > 1) {
      const lh = hh / lines.length;
      lines.forEach((t, i) => {
        const fs = lh * 0.86;
        const txt = s('text', Object.assign({ x: 0, y: L.round(i * lh + lh * 0.8, 2), 'font-size': L.round(fs, 2), textLength: L.round(w, 2), lengthAdjust: 'spacingAndGlyphs' }, common));
        txt.textContent = t || ' ';
        if (wa.vert) { txt.setAttribute('writing-mode', 'tb'); }
        svg.appendChild(txt);
      });
      return svg;
    }
    const fs = Math.min(hh * (warp === 'textCircle' ? 0.2 : 0.42), (w / Math.max(lines[0].length, 1)) * 1.4);
    const pts = warpPoints(warp, w, hh, fs);
    let len = 0;
    for (let i = 1; i < pts.length; i++) len += Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]);
    const id = nid('wp');
    defs.appendChild(s('path', { id, d: 'M' + pts.map((p) => L.round(p[0], 2) + ',' + L.round(p[1], 2)).join(' L') }));
    const txt = s('text', Object.assign({ 'font-size': L.round(fs, 2) }, common));
    const tp = s('textPath', { href: '#' + id, textLength: L.round(len * 0.98, 2), lengthAdjust: 'spacingAndGlyphs', startOffset: L.round(len * 0.01, 2) });
    tp.textContent = lines[0];
    txt.appendChild(tp);
    svg.appendChild(txt);
    svg.setAttribute('preserveAspectRatio', 'xMidYMid meet');
    return svg;
  };

  /* content-placeholder icon palette (2003) */
  const CONTENT_ICONS = [['table', 'Insert Table'], ['chart', 'Insert Chart'], ['clipart', 'Insert Clip Art'], ['picture', 'Insert Picture'], ['diagram', 'Insert Diagram or Organization Chart'], ['media', 'Insert Media Clip']];

  /**
   * Render a shape. ctx = {pres, slide, design, mode, num}
   * opts = {deco, interactive, inGroup}
   */
  R.shape = function (sh, ctx, opts) {
    opts = opts || {};
    const design = ctx.design;
    const interactive = ctx.mode === 'edit' && !opts.deco;
    const el = h('div', { class: 'shp t-' + sh.type });
    el.dataset.sid = sh.id;
    const st = el.style;
    const ox = opts.ox || 0, oy = opts.oy || 0;
    st.left = L.round(sh.x - ox, 3) + 'px';
    st.top = L.round(sh.y - oy, 3) + 'px';
    st.width = L.round(Math.max(sh.w, 0), 3) + 'px';
    st.height = L.round(Math.max(sh.h, 0), 3) + 'px';
    if (sh.rot) st.transform = `rotate(${L.round(sh.rot, 3)}deg)`;
    if (sh.hidden) st.visibility = 'hidden';
    if (opts.deco) el.classList.add('deco');
    if (interactive) el.classList.add('ia');
    if (sh.link && ctx.mode === 'show') { el.dataset.link = JSON.stringify(sh.link); el.classList.add('has-link'); }

    switch (sh.type) {
      case 'line':
        el.appendChild(lineSVG(sh, design, ctx, { interactive }));
        break;
      case 'image':
        el.appendChild(pictureEl(sh, design, opts));
        if (sh.line && sh.line.t !== 'none' && sh.line.c) el.appendChild(geomSVG(Object.assign({}, sh, { fill: { t: 'none' } }), design, ctx, {}));
        if (sh.shadow) el.style.filter = R.shadowCSS(sh.shadow, design);
        break;
      case 'table':
        el.appendChild(tableEl(sh, design, ctx, opts));
        break;
      case 'chart':
        if (L.chart) el.appendChild(L.chart.render(sh.chart, sh.w, sh.h, design));
        break;
      case 'wordart':
        el.appendChild(R.wordart(sh, design, ctx));
        break;
      case 'group': {
        for (const k of sh.kids || []) el.appendChild(R.shape(k, ctx, { ox: sh.x, oy: sh.y, inGroup: true, deco: opts.deco }));
        break;
      }
      default: {
        const hasGeom = (sh.fill && sh.fill.t !== 'none') || (sh.line && sh.line.t !== 'none' && sh.line.c) || (sh.geom && sh.geom !== 'rect' && sh.type === 'shape');
        /* an empty placeholder is only an editing aid: thumbnails, the slide show and printouts leave it out */
        const emptyPh = sh.ph && (!sh.tx || L.txt.isEmpty(sh.tx)) && ctx.mode !== 'edit' && !opts.deco;
        if (hasGeom && !emptyPh) el.appendChild(geomSVG(sh, design, ctx, { interactive }));
        if (sh.tx) {
          const empty = L.txt.isEmpty(sh.tx);
          const rect = L.geom.textRect(sh);
          if (!(empty && sh.ph && ctx.mode !== 'edit')) {
            const tb = R.textBody(sh, sh.tx, design, ctx, rect, {});
            if (sh.shadow && !hasGeom) tb.style.filter = R.shadowCSS(sh.shadow, design);
            el.appendChild(tb);
          }
          if (empty && (sh.ph || (sh.sa && sh.sa.role === 'node')) && ctx.mode === 'edit' && ctx.prompts !== false) {
            const ptype = sh.ph ? sh.ph.type : null;
            const prompt = !sh.ph ? '[Text]' : L.model.PROMPTS[ptype] != null ? L.model.PROMPTS[ptype] : 'Click to add text';
            if (prompt) {
              const ptx = Object.assign({}, sh.tx, { ps: [L.txt.para(prompt, sh.tx.ps[0] && sh.tx.ps[0].pp, sh.tx.ps[0] && (sh.tx.ps[0].end || (sh.tx.ps[0].rs[0] && L.txt.runProps(sh.tx.ps[0].rs[0]))), sh.tx.ps[0] ? sh.tx.ps[0].lvl : 0)] });
              const pe = R.textBody(sh, ptx, design, ctx, rect, {});
              pe.classList.add('prompt');
              el.appendChild(pe);
            }
            if (ptype === 'obj' || ptype === 'tbl' || ptype === 'chart') {
              const pal = h('div', { class: 'ph-content' });
              const icons = ptype === 'obj' ? CONTENT_ICONS : CONTENT_ICONS.filter((c) => c[0] === (ptype === 'tbl' ? 'table' : 'chart'));
              for (const [k, tip] of icons) pal.appendChild(h('button', { class: 'ph-ico', 'data-content-act': k, 'data-tip': tip, html: L.icons ? L.icons.get('ct_' + k) : '' }));
              el.appendChild(pal);
            }
            if (sh.ph) el.classList.add('ph-empty');
          }
        }
      }
    }
    if (sh.ph && ctx.mode === 'edit' && !opts.deco) el.classList.add('is-ph');
    return el;
  };

  /* header / footer (from pres.hf, overridable per slide) */
  function footers(ctx, root) {
    const { pres, slide, design } = ctx;
    const hf = Object.assign({}, pres.hf || {}, slide.hf || {});
    if (!hf.dt && !hf.num && !hf.ftr) return;
    if (hf.notOnTitle && slide.layout === 'title') return;
    const has = (t) => slide.shapes.some((x) => x.ph && x.ph.type === t);
    const mk = (type, run) => {
      const r = design.ph[type];
      if (!r) return;
      const sh = { id: 'hf-' + type, type: 'text', ph: { type }, x: r.x, y: r.y, w: r.w, h: r.h, tx: L.txt.body([{ lvl: 0, pp: {}, rs: [run] }], { anchor: 't' }) };
      const el = R.shape(sh, Object.assign({}, ctx, { mode: ctx.mode === 'edit' ? 'thumb' : ctx.mode }), { deco: true });
      el.classList.add('hf');
      root.appendChild(el);
    };
    if (hf.dt && !has('dt')) mk('dt', hf.dtAuto ? { t: '', fld: hf.dtFmt || 'datetime1' } : { t: hf.dtText || '' });
    if (hf.ftr && !has('ftr') && hf.ftrText) mk('ftr', { t: hf.ftrText });
    if (hf.num && !has('sldNum')) mk('sldNum', { t: '', fld: 'slidenum' });
  }

  /** Render a complete slide. opts: {mode:'edit'|'thumb'|'show'|'print', index} */
  R.slide = function (pres, slide, opts) {
    opts = opts || {};
    const mode = opts.mode || 'thumb';
    const design = L.model.design(pres, slide);
    const W = pres.W, H = pres.H;
    const root = h('div', { class: 'sl sl-' + mode });
    root.style.width = W + 'px';
    root.style.height = H + 'px';
    root.dataset.slide = slide.id;
    const index = opts.index != null ? opts.index : pres.slides.indexOf(slide);
    const ctx = { pres, slide, design, mode, num: index + (pres.firstNum || 1), prompts: opts.prompts };
    const isTitle = slide.layout === 'title';
    const lkDeco = slide.lkey && design.layoutDecos ? design.layoutDecos[slide.lkey] : null;
    const lkBg = slide.lkey && design.layoutBgs ? design.layoutBgs[slide.lkey] : null;
    const bgFill = slide.bg || lkBg || (isTitle && design.titleBg) || design.bg;
    /* shapes filled with "Background" take this slide's background */
    R.curBg = bgFill && bgFill.t !== 'bg' ? bgFill : null;
    root.appendChild(R.background(bgFill, design, W, H));
    if (!slide.hideMaster) {
      const layer = h('div', { class: 'sl-master' });
      const deco = isTitle && design.titleDeco ? design.titleDeco : lkDeco ? (design.layoutShowMaster?.[slide.lkey] ? (design.deco || []).concat(lkDeco) : lkDeco) : design.deco || [];
      for (const d of deco) layer.appendChild(R.shape(d, ctx, { deco: true }));
      root.appendChild(layer);
    }
    footers(ctx, root);
    for (const sh of slide.shapes) {
      const el = R.shape(sh, ctx, {});
      if (mode === 'show') {
        const aw = h('div', { class: 'aw' });
        aw.dataset.sid = sh.id;
        aw.appendChild(el);
        root.appendChild(aw);
      } else root.appendChild(el);
    }
    return root;
  };

  /** Scaled thumbnail element of fixed pixel width */
  R.thumb = function (pres, slide, width, opts) {
    const k = width / pres.W;
    const box = h('div', { class: 'thumb-box' });
    box.style.width = Math.round(width) + 'px';
    box.style.height = Math.round(pres.H * k) + 'px';
    const el = R.slide(pres, slide, Object.assign({ mode: 'thumb' }, opts || {}));
    el.style.transform = `scale(${k})`;
    el.style.transformOrigin = '0 0';
    box.appendChild(el);
    return box;
  };

  /* small swatch background for a fill (UI) */
  R.swatch = (fill, design) => R.fillCSS(fill, design);
})();
