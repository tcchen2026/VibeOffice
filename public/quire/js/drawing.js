/* Quire — drawing layer: inline and floating pictures, AutoShapes, text boxes, groups, charts, WordArt,
 * watermarks; positioning with Word's wrapping styles; object selection, moving and resizing.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, R = L.R;
  const { s } = L;
  const DR = (L.drawing = {});
  const PX = D.PX;
  const px = (pt) => R.px(pt);
  let gid = 0;
  const nid = (p) => (p || 'g') + (++gid);
  const el = (tag, cls, css) => { const e = document.createElement(tag); if (cls) e.className = cls; if (css) e.style.cssText = css; return e; };

  /* ---------- colour / slide-model shims for the shared chart module ---------- */
  const themeCtx = () => ({ theme: D.doc && D.doc.theme });
  L.model = L.model || {
    resolveColor: (c, design, fb) => (!c ? fb || '#000000' : c[0] === '#' ? c : L.dml.schemeHex(c, themeCtx())),
    design: () => null,
  };
  const col = (c, fb) => (!c ? fb || '#000000' : c[0] === '#' ? c : /^[0-9A-F]{6}$/i.test(c) ? '#' + c : L.dml.schemeHex(c, themeCtx()));

  /* ---------- patterns ---------- */
  const PATT = {
    pct5: [0x80, 0, 0, 0, 0x08, 0, 0, 0], pct10: [0x80, 0, 0x08, 0, 0x80, 0, 0x08, 0], pct20: [0x88, 0, 0x22, 0, 0x88, 0, 0x22, 0],
    pct25: [0x88, 0x22, 0x88, 0x22, 0x88, 0x22, 0x88, 0x22], pct30: [0xAA, 0x44, 0xAA, 0x11, 0xAA, 0x44, 0xAA, 0x11], pct40: [0xAA, 0x55, 0xAA, 0x51, 0xAA, 0x55, 0xAA, 0x15],
    pct50: [0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55], pct60: [0xEE, 0x55, 0xBB, 0x55, 0xEE, 0x55, 0xBB, 0x55], pct70: [0xEE, 0xBB, 0xEE, 0xBB, 0xEE, 0xBB, 0xEE, 0xBB],
    pct75: [0xEE, 0xFF, 0xBB, 0xFF, 0xEE, 0xFF, 0xBB, 0xFF], pct80: [0xF7, 0xFF, 0x7F, 0xFF, 0xF7, 0xFF, 0x7F, 0xFF], pct90: [0xFF, 0xF7, 0xFF, 0xFF, 0xFF, 0x7F, 0xFF, 0xFF],
    horz: [0xFF, 0, 0, 0, 0xFF, 0, 0, 0], vert: [0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88], ltHorz: [0xFF, 0, 0, 0, 0, 0, 0, 0], ltVert: [0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
    dkHorz: [0xFF, 0xFF, 0, 0, 0xFF, 0xFF, 0, 0], dkVert: [0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC], narHorz: [0xFF, 0, 0xFF, 0, 0xFF, 0, 0xFF, 0], narVert: [0xAA, 0xAA, 0xAA, 0xAA, 0xAA, 0xAA, 0xAA, 0xAA],
    dnDiag: [0x88, 0x44, 0x22, 0x11, 0x88, 0x44, 0x22, 0x11], upDiag: [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88],
    ltDnDiag: [0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01], ltUpDiag: [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80],
    cross: [0xFF, 0x88, 0x88, 0x88, 0xFF, 0x88, 0x88, 0x88], diagCross: [0x81, 0x42, 0x24, 0x18, 0x18, 0x24, 0x42, 0x81], smCheck: [0xCC, 0xCC, 0x33, 0x33, 0xCC, 0xCC, 0x33, 0x33],
    lgCheck: [0xF0, 0xF0, 0xF0, 0xF0, 0x0F, 0x0F, 0x0F, 0x0F], horzBrick: [0xFF, 0x80, 0x80, 0x80, 0xFF, 0x08, 0x08, 0x08], weave: [0x88, 0x54, 0x22, 0x45, 0x88, 0x14, 0x22, 0x51],
    smConfetti: [0x80, 0x08, 0x40, 0x02, 0x10, 0x01, 0x20, 0x04], zigZag: [0x81, 0x42, 0x24, 0x18, 0x81, 0x42, 0x24, 0x18], dotGrid: [0xAA, 0, 0x80, 0, 0x80, 0, 0x80, 0],
  };
  DR.PATTERNS = PATT;
  const pattPath = (bits) => { let d = ''; bits.forEach((row, y) => { for (let x = 0; x < 8; x++) if (row & (0x80 >> x)) d += `M${x},${y}h1v1h-1z`; }); return d; };
  DR.patternSVG = (prst, fg, bg, tile) => `<svg xmlns="http://www.w3.org/2000/svg" width="${tile || 6}" height="${tile || 6}" viewBox="0 0 8 8" shape-rendering="crispEdges"><rect width="8" height="8" fill="${bg}"/><path d="${pattPath(PATT[prst] || PATT.pct50)}" fill="${fg}"/></svg>`;
  DR.fillCSS = function (fill) {
    if (!fill || fill.t === 'none') return 'transparent';
    if (fill.t === 'solid') return L.color.rgba(col(fill.c), fill.a);
    if (fill.t === 'grad') {
      const stops = (fill.stops || []).slice().sort((a, b) => a.p - b.p).map((st) => `${L.color.rgba(col(st.c), st.a)} ${L.round(st.p * 100, 1)}%`).join(', ');
      if (fill.path === 'circle' || fill.path === 'shape') return `radial-gradient(circle farthest-side at 50% 50%, ${stops})`;
      if (fill.path === 'rect') return `radial-gradient(ellipse farthest-corner at 50% 50%, ${stops})`;
      return `linear-gradient(${L.round((fill.ang || 0) + 90, 2)}deg, ${stops})`;
    }
    if (fill.t === 'img') { const u = L.media.url(fill.media); return fill.tile ? `url("${u}") repeat` : `url("${u}") center / 100% 100% no-repeat`; }
    if (fill.t === 'patt') return 'url("data:image/svg+xml;utf8,' + encodeURIComponent(DR.patternSVG(fill.prst, col(fill.fg), col(fill.bg, '#FFFFFF'))) + '")';
    return 'transparent';
  };
  L.render = L.render || { fillCSS: DR.fillCSS };
  function svgPaint(fill, defs) {
    if (!fill || fill.t === 'none') return { paint: 'none', opacity: 1 };
    if (fill.t === 'solid') return { paint: col(fill.c), opacity: fill.a == null ? 1 : fill.a };
    if (fill.t === 'grad') {
      const id = nid('gr');
      let g;
      if (fill.path === 'circle' || fill.path === 'rect' || fill.path === 'shape') g = s('radialGradient', { id, cx: '50%', cy: '50%', r: fill.path === 'rect' ? '71%' : '50%' });
      else {
        const a = ((fill.ang || 0) * Math.PI) / 180, c = Math.cos(a), si = Math.sin(a), k = Math.abs(c) + Math.abs(si);
        g = s('linearGradient', { id, gradientUnits: 'objectBoundingBox', x1: L.round(0.5 - (c * k) / 2, 4), y1: L.round(0.5 - (si * k) / 2, 4), x2: L.round(0.5 + (c * k) / 2, 4), y2: L.round(0.5 + (si * k) / 2, 4) });
      }
      for (const st of (fill.stops || []).slice().sort((a, b) => a.p - b.p)) g.appendChild(s('stop', { offset: L.round(st.p, 4), 'stop-color': col(st.c), 'stop-opacity': st.a == null ? 1 : st.a }));
      defs.appendChild(g);
      return { paint: `url(#${id})`, opacity: 1 };
    }
    if (fill.t === 'img') {
      const id = nid('im');
      defs.appendChild(fill.tile ? s('pattern', { id, patternUnits: 'userSpaceOnUse', width: 64, height: 64 }, s('image', { href: L.media.url(fill.media), width: 64, height: 64, preserveAspectRatio: 'none' }))
        : s('pattern', { id, patternContentUnits: 'objectBoundingBox', width: 1, height: 1 }, s('image', { href: L.media.url(fill.media), width: 1, height: 1, preserveAspectRatio: 'none' })));
      return { paint: `url(#${id})`, opacity: fill.a == null ? 1 : fill.a };
    }
    if (fill.t === 'patt') {
      const id = nid('pt');
      defs.appendChild(s('pattern', { id, patternUnits: 'userSpaceOnUse', width: 6, height: 6 }, s('rect', { width: 6, height: 6, fill: col(fill.bg, '#FFFFFF') }), s('path', { d: pattPath(PATT[fill.prst] || PATT.pct50), fill: col(fill.fg), transform: 'scale(0.75)', 'shape-rendering': 'crispEdges' })));
      return { paint: `url(#${id})`, opacity: 1 };
    }
    return { paint: 'none', opacity: 1 };
  }
  DR.DASH = { solid: null, sysDot: [1, 1], sysDash: [3, 1], dot: [1, 2], dash: [4, 3], dashDot: [4, 3, 1, 3], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1] };
  const dashArray = (dash, w) => { const d = DR.DASH[dash]; return d ? d.map((v) => L.round(v * Math.max(w, 0.75), 2)).join(' ') : null; };
  function marker(defs, end, color, lw, isHead) {
    if (!end || !end.type || end.type === 'none') return null;
    const id = nid('mk');
    const sz = { sm: 2, med: 3, lg: 5 };
    const k = Math.max(lw || 0.75, 2) / Math.max(lw || 0.75, 0.1);
    const m = s('marker', { id, viewBox: '0 0 10 10', refX: end.type === 'arrow' ? 9 : end.type === 'oval' || end.type === 'diamond' ? 5 : 8, refY: 5, markerWidth: (sz[end.len || 'med'] || 3) * k, markerHeight: (sz[end.w || 'med'] || 3) * k, orient: isHead ? 'auto-start-reverse' : 'auto', markerUnits: 'strokeWidth' });
    let shape;
    switch (end.type) {
      case 'stealth': shape = s('path', { d: 'M0,0 L10,5 L0,10 L3,5 Z', fill: color }); break;
      case 'diamond': shape = s('path', { d: 'M0,5 L5,0 L10,5 L5,10 Z', fill: color }); break;
      case 'oval': shape = s('circle', { cx: 5, cy: 5, r: 4.5, fill: color }); break;
      case 'arrow': shape = s('path', { d: 'M1,1 L9,5 L1,9', fill: 'none', stroke: color, 'stroke-width': 1.6 }); break;
      default: shape = s('path', { d: 'M0,0 L10,5 L0,10 Z', fill: color });
    }
    m.appendChild(shape);
    defs.appendChild(m);
    return `url(#${id})`;
  }
  const shadowCSS = (sh) => (sh ? `drop-shadow(${L.round(sh.dx == null ? 3 : sh.dx, 2)}px ${L.round(sh.dy == null ? 3 : sh.dy, 2)}px ${L.round(sh.blur || 0, 2)}px ${L.color.rgba(col(sh.c || '#000000'), sh.a == null ? 0.5 : sh.a)})` : '');

  /* ---------- shape geometry → SVG (in points, scaled by the caller) ---------- */
  function geomSVG(sh) {
    const w = Math.max(sh.w, 0.01), hh = Math.max(sh.h, 0.01);
    const vw = L.round(Math.max(w, 1), 3), vh = L.round(Math.max(hh, 1), 3);
    const svg = s('svg', { class: 'geo', width: L.round(vw * PX, 2), height: L.round(vh * PX, 2), viewBox: `0 0 ${vw} ${vh}`, overflow: 'visible' });
    const defs = s('defs');
    svg.appendChild(defs);
    const isLine = sh.geom === 'line' || L.geom.isLine(sh.geom);
    const ln = sh.line && sh.line.t !== 'none' ? sh.line : isLine ? { c: '#000000', w: 0.75 } : null;
    const lc = ln ? L.color.rgba(col(ln.c || '#000000'), ln.a) : 'none';
    const lw = ln ? Math.max(ln.w == null ? 0.75 : ln.w, 0.1) : 0;
    if (isLine) {
      const p = s('path', { d: L.geom.linePath(sh), fill: 'none', stroke: lc, 'stroke-width': lw });
      const da = dashArray(ln.dash, lw); if (da) p.setAttribute('stroke-dasharray', da);
      const mh = marker(defs, ln.head, lc, lw, true), mt = marker(defs, ln.tail, lc, lw, false);
      if (mh) p.setAttribute('marker-start', mh);
      if (mt) p.setAttribute('marker-end', mt);
      svg.appendChild(p);
      if (sh.shadow) svg.style.filter = shadowCSS(sh.shadow);
      return svg;
    }
    const parts = L.geom.parts(sh, w, hh);
    const paint = svgPaint(sh.fill, defs);
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
        const da = dashArray(ln.dash, lw); if (da) p.setAttribute('stroke-dasharray', da);
        if (part.open) { const mh = marker(defs, ln.head, lc, lw, true), mt = marker(defs, ln.tail, lc, lw, false); if (mh) p.setAttribute('marker-start', mh); if (mt) p.setAttribute('marker-end', mt); }
      } else p.setAttribute('stroke', 'none');
      g.appendChild(p);
      if (filled && hasFill && part.fill && part.fill !== 'norm') {
        const o = { darken: ['#000', 0.4], darkenLess: ['#000', 0.2], lighten: ['#fff', 0.4], lightenLess: ['#fff', 0.2] }[part.fill];
        if (o) g.appendChild(s('path', { d: part.d, fill: o[0], 'fill-opacity': o[1] * (paint.opacity || 1), stroke: 'none' }));
      }
    }
    svg.appendChild(g);
    if (sh.shadow) svg.style.filter = shadowCSS(sh.shadow);
    return svg;
  }

  /* ---------- WordArt ---------- */
  DR.WARPS = [['textPlain', 'Plain Text'], ['textArchUp', 'Arch Up'], ['textArchDown', 'Arch Down'], ['textCircle', 'Circle'], ['textWave1', 'Wave'], ['textSlantUp', 'Slant Up'], ['textSlantDown', 'Slant Down'], ['textTriangle', 'Triangle Up'], ['textChevron', 'Chevron Up'], ['textCanUp', 'Can Up']];
  function warpPoints(warp, w, hh, fs) {
    const pts = [], N = 64, m = fs * 0.15;
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
        case 'textCanUp': pts.push([m + t * (w - 2 * m), hh * 0.85 - Math.sin(t * Math.PI) * hh * 0.25]); break;
        default: pts.push([t * w, hh / 2]);
      }
    }
    return pts;
  }
  DR.wordartSVG = function (sh) {
    const wa = sh.wordart || { text: 'LettersArt' };
    const w = Math.max(sh.w, 1), hh = Math.max(sh.h, 1);
    const svg = s('svg', { class: 'wa', width: L.round(w * PX, 2), height: L.round(hh * PX, 2), viewBox: `0 0 ${L.round(w, 3)} ${L.round(hh, 3)}`, preserveAspectRatio: 'none', overflow: 'visible' });
    const defs = s('defs');
    svg.appendChild(defs);
    const fill = svgPaint(sh.fill && sh.fill.t !== 'none' ? sh.fill : { t: 'solid', c: '#3366CC' }, defs);
    const ln = sh.line && sh.line.t !== 'none' && sh.line.c ? sh.line : null;
    const lines = String(wa.text || '').split('\n');
    const common = { 'font-family': L.fontStack(wa.font || 'Arial Black'), 'font-weight': wa.b ? 700 : 400, 'font-style': wa.i ? 'italic' : 'normal', fill: fill.paint, 'fill-opacity': fill.opacity };
    if (ln) Object.assign(common, { stroke: col(ln.c), 'stroke-width': Math.max(0.25, ln.w || 0.75), 'paint-order': 'stroke', 'stroke-linejoin': 'round' });
    if (sh.shadow) svg.style.filter = shadowCSS(sh.shadow);
    const warp = sh.warp || wa.warp || 'textPlain';
    if (warp === 'textPlain' || lines.length > 1) {
      const lh = hh / lines.length;
      lines.forEach((t, i) => {
        const txt = s('text', Object.assign({ x: 0, y: L.round(i * lh + lh * 0.8, 2), 'font-size': L.round(lh * 0.86, 2), textLength: L.round(w, 2), lengthAdjust: 'spacingAndGlyphs' }, common));
        txt.textContent = t || ' ';
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

  /* ---------- object element (shared by inline and floating placement) ---------- */
  /** build the visual for an item: picture, shape (+text box), group, chart */
  DR.objectEl = function (it, ctx, opts) {
    opts = opts || {};
    const w = px(it.w || 72), hh = px(it.h || 72);
    if (it.t === 'img') {
      const box = R.image(Object.assign({}, it, { float: null }), it.rPr || {}, ctx);
      if (it.alpha != null && it.alpha < 1 && box.firstChild) box.firstChild.style.opacity = it.alpha;
      if (it.geom && it.geom.geom && it.geom.geom !== 'rect') { const parts = L.geom.parts(Object.assign({ w: w, h: hh }, it.geom), w, hh).filter((p) => p.fill !== 'none'); if (parts.length) box.style.clipPath = `path('${parts.map((p) => p.d).join(' ')}')`; }
      return box;
    }
    const box = el('span', 'shp', `display:inline-block;position:relative;width:${w}px;height:${hh}px;text-indent:0;line-height:normal;vertical-align:baseline`);
    box.contentEditable = 'false';
    if (it.t === 'chart') {
      if (L.chart && it.chart) {
        const svg = L.chart.render(it.chart, it.w, it.h, null);
        svg.setAttribute('width', w); svg.setAttribute('height', hh);
        svg.style.cssText = 'position:absolute;left:0;top:0';
        box.appendChild(svg);
      }
      box.classList.add('chart');
      return box;
    }
    if (it.t === 'group') {
      if (it.fill && it.fill.t !== 'none') box.style.background = DR.fillCSS(it.fill);
      for (const k of it.kids || []) {
        const kEl = DR.objectEl(k, ctx, { inGroup: true });
        kEl.style.position = 'absolute';
        kEl.style.left = px(k.x || 0) + 'px';
        kEl.style.top = px(k.y || 0) + 'px';
        if (k.rot) kEl.style.transform = (kEl.style.transform || '') + ` rotate(${k.rot}deg)`;
        box.appendChild(kEl);
      }
      box.classList.add('sgrp');
      return box;
    }
    /* AutoShape / text box / WordArt */
    if (it.wordart) box.appendChild(DR.wordartSVG(it));
    else box.appendChild(geomSVG(it));
    if (it.tb || it.tbLink) {
      const tr = it.txRect || (L.geom.textRect && it.geom !== 'rect' ? (() => { const r = L.geom.textRect({ geom: it.geom, adj: it.adj, w: it.w, h: it.h }); return { x: r[0], y: r[1], w: r[2] - r[0], h: r[3] - r[1] }; })() : { x: 0, y: 0, w: it.w, h: it.h });
      const ins = it.ins || { l: 7.2, t: 3.6, r: 7.2, b: 3.6 };
      const tbx = el('div', 'tbx', `left:${px(tr.x + ins.l)}px;top:${px(tr.y + ins.t)}px;width:${Math.max(4, px(tr.w - ins.l - ins.r))}px;height:${Math.max(4, px(tr.h - ins.t - ins.b))}px;display:flex;flex-direction:column;justify-content:${it.anchor === 'ctr' ? 'center' : it.anchor === 'b' ? 'flex-end' : 'flex-start'}`);
      if (it.overflowText) tbx.style.overflow = 'visible';
      if (it.noWrap) tbx.classList.add('nowrap');
      if (it.vert === 'vert' || it.vert === 'eaVert') tbx.style.writingMode = 'vertical-rl';
      else if (it.vert === 'vert270') { tbx.style.writingMode = 'vertical-rl'; tbx.style.transform = 'rotate(180deg)'; }
      tbx.contentEditable = opts.readOnly ? 'false' : 'true';
      tbx.classList.add('ed-tb');
      const inner = el('div', 'tbi', 'position:relative;width:100%');
      tbx.appendChild(inner);
      box.appendChild(tbx);
      /* paragraphs are laid out once the box is in the document (tab stops need measuring) */
      box.dataset.pending = '1';
      box._fill = () => {
        delete box.dataset.pending;
        inner.textContent = '';
        /* linked text boxes: the story starts in the first box and continues where the previous one stopped */
        const chains = L.layout.tbChains || (L.layout.tbChains = new Map());
        const link = it.tbLink ? chains.get(it.tbLink.id) : null;
        if (it.tbLink && (!link || link.seq !== it.tbLink.seq - 1)) {
          /* the box before this one in the chain is not laid out yet (it may come later on the page) */
          const w = chains.waiters || (chains.waiters = []);
          w.push({ id: it.tbLink.id, seq: it.tbLink.seq, fill: box._fill, el: box });
          return;
        }
        const story = it.tb || (link && link.story);
        if (!story || (it.tbLink && !link.st)) return;
        const chained = !!(it.tbId || it.tbLink) && !it.autofit && !opts.inGroup && !opts.readOnly;
        const tctx = Object.assign({}, ctx, { story, inTextBox: true, openCmts: null, tbl: null, inCell: false });
        if (it.fontColor) tctx.dkPara = R.isDark(it.fontColor.replace('#', '')) ? false : true;
        const fr = { el: inner, limit: chained ? Math.max(4, px(tr.h - ins.t - ins.b)) : Infinity, top: chained ? inner.getBoundingClientRect().top : 0, empty: true, kind: 'tb', width: Math.max(4, px(tr.w - ins.l - ins.r)) };
        const res = L.layout.flow(fr, story.blocks, it.tbLink ? link.st : { bi: 0, sub: null }, story.blocks.length, tctx);
        if (chained) {
          const cid = it.tbLink ? it.tbLink.id : it.tbId, seq = it.tbLink ? it.tbLink.seq : 0;
          chains.set(cid, { story, st: res.stop === 'done' ? null : res.st, seq });
          const w = chains.waiters || [];
          const next = w.find((x) => x.id === cid && x.seq === seq + 1 && x.el.isConnected);
          if (next) { chains.waiters = w.filter((x) => x !== next); next.fill(); }
        }
        if (it.autofit && !opts.inGroup) {
          /* resize-shape-to-fit-text: the box grows with its text */
          const need = inner.offsetHeight + px(ins.t + ins.b);
          if (need > hh + 1) { box.style.height = need + 'px'; tbx.style.height = inner.offsetHeight + 'px'; const svg = box.querySelector('svg.geo'); if (svg) { const nh = need / PX; svg.setAttribute('height', need); svg.setAttribute('viewBox', `0 0 ${L.round(Math.max(it.w, 1), 3)} ${L.round(nh, 3)}`); const parts = L.geom.parts(Object.assign({}, it, { h: nh }), it.w, nh); const ps = svg.querySelectorAll('g > path'); parts.forEach((pp, i) => { if (ps[i]) ps[i].setAttribute('d', pp.d); }); } }
        }
      };
    }
    return box;
  };
  /** run deferred text-box layouts below an element */
  DR.fillPending = function (root) {
    for (const b of root.querySelectorAll('[data-pending]')) if (b._fill) b._fill();
  };

  /* ---------- inline objects (In Line With Text) ---------- */
  DR.renderInline = function (it, r, ctx, p) {
    void p;
    if (it.float) {
      /* floating objects leave an anchor in the text; the object itself is positioned after layout */
      const a = el('span', 'anc');
      a.contentEditable = 'false';
      a._item = it;
      if (ctx.view === 'normal' || ctx.view === 'outline') a.classList.add('hidden-float');
      return a;
    }
    const box = DR.objectEl(it, ctx);
    if (it.rot) box.style.transform = `rotate(${it.rot}deg)`;
    return box;
  };

  /* ---------- floating placement ---------- */
  function pageRec(frag, ctx, frame) {
    if (frame && frame.page) return frame.page;
    if (ctx && ctx.page) return ctx.page;
    const pg = frag.closest('.pg');
    if (!pg) return null;
    return L.layout.pages[+pg.dataset.pg] || null;
  }
  function layer(page, behind) {
    const cls = behind ? 'pg-float back' : 'pg-float front';
    let lay = page.el.querySelector(':scope > .pg-float.' + (behind ? 'back' : 'front'));
    if (!lay) {
      lay = el('div', cls);
      lay.contentEditable = 'false';
      if (behind) page.el.insertBefore(lay, page.el.firstChild); else page.el.appendChild(lay);
    }
    return lay;
  }
  /** x, y (px, page coordinates) for a float */
  function floatPos(it, page, frag, anc, frame) {
    const f = it.float;
    const z = L.layout.scale();
    const pgR = page.el.getBoundingClientRect();
    const rel = (r) => ({ x: (r.left - pgR.left) / z, y: (r.top - pgR.top) / z, w: r.width / z, h: r.height / z });
    const sect = page.sect;
    const ow = px(it.w || 72), oh = px(it.h || 72);
    const even = page.numVal % 2 === 0;
    const colEl = frame && frame.el && frame.kind === 'col' ? frame.el : frag.closest('.pg-col, .cf, .pg-hdr, .pg-ftr') || page.body;
    const colR = rel(colEl.getBoundingClientRect());
    let x0, aw;
    switch (f.posH.rel) {
      case 'page': x0 = 0; aw = page.W; break;
      case 'margin': x0 = page.bodyLeft; aw = page.bodyW; break;
      case 'leftMargin': x0 = 0; aw = page.bodyLeft; break;
      case 'rightMargin': x0 = page.bodyLeft + page.bodyW; aw = page.W - x0; break;
      case 'insideMargin': x0 = even ? page.bodyLeft + page.bodyW : 0; aw = even ? page.W - x0 : page.bodyLeft; break;
      case 'outsideMargin': x0 = even ? 0 : page.bodyLeft + page.bodyW; aw = even ? page.bodyLeft : page.W - x0; break;
      case 'character': { const ar = rel(anc.getBoundingClientRect()); x0 = ar.x; aw = 0; break; }
      default: x0 = colR.x; aw = colR.w; break;
    }
    let x;
    const ah = f.posH.align;
    if (ah) x = ah === 'center' ? x0 + (aw - ow) / 2 : ah === 'right' || (ah === 'outside' && !even) || (ah === 'inside' && even) ? x0 + aw - ow : x0;
    else if (f.posH.pct != null) x = x0 + aw * f.posH.pct / 100;
    else x = x0 + px(f.posH.off || 0);
    const fragR = rel(frag.getBoundingClientRect());
    const mt = px(sect.mt), mb = px(sect.mb);
    let y0, ahh;
    switch (f.posV.rel) {
      case 'page': y0 = 0; ahh = page.H; break;
      case 'margin': y0 = mt; ahh = page.H - mt - mb; break;
      case 'topMargin': y0 = 0; ahh = mt; break;
      case 'bottomMargin': y0 = page.H - mb; ahh = mb; break;
      case 'insideMargin': case 'outsideMargin': y0 = 0; ahh = mt; break;
      case 'line': { const ar = rel(anc.getBoundingClientRect()); y0 = ar.y; ahh = 0; break; }
      /* "relative to paragraph" measures from the top of the paragraph, space before included (as Word does) */
      default: y0 = fragR.y; ahh = 0; break;
    }
    let y;
    const av = f.posV.align;
    if (av) y = av === 'center' ? y0 + (ahh - oh) / 2 : av === 'bottom' || av === 'outside' ? y0 + ahh - oh : y0;
    else if (f.posV.pct != null) y = y0 + ahh * f.posV.pct / 100;
    else y = y0 + px(f.posV.off || 0);
    return { x, y, ow, oh, colR, fragR, baseX: x - (ah ? 0 : px(f.posH.off || 0)), baseY: y - (av ? 0 : px(f.posV.off || 0)) };
  }
  /**
   * Called after a paragraph fragment is attached: places floating objects anchored in it
   * and lays out text boxes of inline shapes.
   */
  DR.anchorFloats = function (frag, p, ctx, frame) {
    DR.fillPending(frag);
    const ancs = frag.querySelectorAll('.anc');
    if (!ancs.length) return;
    if (ctx.view === 'normal' || ctx.view === 'outline') return;
    const page = ctx.view === 'web' ? null : pageRec(frag, ctx, frame);
    const wraps = [];
    for (const anc of ancs) {
      const it = anc._item;
      if (!it || !it.float || it.hidden) continue;
      const f = it.float;
      const obj = DR.objectEl(it, ctx);
      obj.classList.add('flt-obj');
      const wrap = el('div', 'flt');
      wrap.contentEditable = 'false';
      wrap.appendChild(obj);
      wrap.dataset.pid = p.id;
      wrap.dataset.o = anc.dataset.o;
      wrap._item = it;
      if (it.rot) obj.style.transform = (obj.style.transform || '') + ` rotate(${it.rot}deg)`;
      const textWrap = (f.wrap === 'square' || f.wrap === 'tight' || f.wrap === 'through' || f.wrap === 'topBottom') && !ctx.hf;
      if (!page || ctx.view === 'web') {
        /* web layout: float to the side of the paragraph */
        wrap.style.cssText = `float:${f.posH.align === 'right' ? 'right' : 'left'};margin:${px(f.dist ? f.dist.t : 0)}px 9px 4px 9px;position:relative`;
        frag.firstChild.insertBefore(wrap, frag.firstChild.firstChild);
        DR.fillPending(wrap);
        continue;
      }
      const pos = floatPos(it, page, frag, anc, frame);
      if (!textWrap || ctx.inCell && !f.layoutInCell) {
        const lay = layer(page, f.wrap === 'behind' || f.behind);
        wrap.style.cssText = `position:absolute;left:${pos.x}px;top:${pos.y}px;z-index:${Math.max(0, Math.min(1000, Math.round((f.z || 0) / 1e6)))}`;
        wrap._pos = pos;
        lay.appendChild(wrap);
        (frag._floats || (frag._floats = [])).push(wrap);
        DR.fillPending(wrap);
        continue;
      }
      /* text wraps around: CSS floats inside the anchor paragraph (placed below, top to bottom) */
      wraps.push({ wrap, obj, pos, f });
    }
    if (!wraps.length) return;
    const pc = frag.firstChild;
    const pcR = pc.getBoundingClientRect(), z = L.layout.scale(), pgR = page.el.getBoundingClientRect();
    const pcX = (pcR.left - pgR.left) / z, pcY = (pcR.top - pgR.top) / z, pcW = pcR.width / z;
    /* CSS never places a float above an earlier one, so insert them in vertical order; with several,
       each starts below the previous ones (clear) at its own offset */
    wraps.sort((a, b) => a.pos.y - b.pos.y);
    const many = wraps.length > 1;
    let cursor = 0;
    let ref = pc.firstChild;
    for (const { wrap, obj, pos, f } of wraps) {
      const dist = f.dist || { t: 0, b: 0, l: 9, r: 9 };
      const t = px(dist.t || 0), b = px(dist.b || 0);
      const dy = Math.max(0, pos.y - pcY - t);
      const leftSide = pos.x + pos.ow / 2 < pcX + pcW / 2 || f.side === 'right';
      const side = f.wrap === 'topBottom' || leftSide ? 'left' : 'right';
      const sp = el('span', 'fl-sp', `float:${side};width:0;height:${Math.max(0, many ? dy - cursor : dy)}px${many ? ';clear:both' : ''}`);
      sp.contentEditable = 'false';
      if (f.wrap === 'topBottom') {
        wrap.style.cssText = `float:left;clear:${many ? 'both' : 'left'};width:100%;height:${pos.oh + b}px;position:relative;padding-top:${t}px`;
        obj.style.position = 'absolute';
        obj.style.left = pos.x - pcX + 'px';
        obj.style.top = t + 'px';
      } else if (leftSide) {
        wrap.style.cssText = `float:left;clear:${many ? 'both' : 'left'};margin-left:${pos.x - pcX}px;margin-right:${px(dist.r || 9)}px;margin-bottom:${b}px;margin-top:${t}px;position:relative`;
      } else {
        wrap.style.cssText = `float:right;clear:${many ? 'both' : 'right'};margin-right:${pcX + pcW - pos.x - pos.ow}px;margin-left:${px(dist.l || 9)}px;margin-bottom:${b}px;margin-top:${t}px;position:relative`;
      }
      wrap._pos = pos;
      pc.insertBefore(sp, ref);
      pc.insertBefore(wrap, ref);
      DR.fillPending(wrap);
      /* a text box that grows to fit its text takes its real height in the wrap */
      let oh = pos.oh;
      if (obj.querySelector('.tbx')) { const real = obj.offsetHeight; if (real > oh + 0.5) { oh = real; if (f.wrap === 'topBottom') wrap.style.height = oh + b + 'px'; } }
      cursor = Math.max(cursor, dy) + t + oh + b;
    }
  };

  /* ---------- watermark ---------- */
  DR.watermark = function (page, wm) {
    if (!wm) return;
    const box = el('div', 'pg-wm');
    box.contentEditable = 'false';
    if (wm.type === 'text' && wm.text) {
      const diag = wm.layout !== 'horizontal';
      const W = page.W, H = page.H;
      const len = diag ? Math.hypot(W, H) * 0.75 : W * 0.8;
      const fs = wm.size && wm.size > 1 ? px(wm.size) : Math.min(len / Math.max(4, wm.text.length) * 1.6, 220);
      const t = el('div', null, `font-family:${L.fontStack(wm.font || 'Times New Roman').replace(/"/g, "'")};font-size:${fs}px;color:#${wm.color || 'C0C0C0'};white-space:nowrap;transform:rotate(${diag ? -45 : 0}deg);opacity:${wm.semi === false ? 1 : 0.5};line-height:1`);
      t.textContent = wm.text;
      box.appendChild(t);
    } else if (wm.type === 'picture' && wm.media) {
      const im = el('img', null, `max-width:${page.bodyW}px;max-height:${page.H * 0.7}px;${wm.washout !== false ? 'filter:brightness(1.7) contrast(0.25);' : ''}`);
      im.src = L.media.url(wm.media);
      if (wm.scale) { im.style.width = px((wm.w || 200) * wm.scale) + 'px'; }
      box.appendChild(im);
    }
    page.el.insertBefore(box, page.el.firstChild);
  };

  /* ---------- selection & manipulation of floating objects ---------- */
  let handles = null, selWrap = null;
  DR.clearSelection = function () { if (handles) { handles.remove(); handles = null; } if (selWrap) selWrap.classList.remove('sel-obj'); selWrap = null; };
  function findAnchor(wrap) {
    const p = D.byId(D.doc, +wrap.dataset.pid);
    if (!p) return null;
    const o = +wrap.dataset.o;
    return { p, o };
  }
  /** mousedown on an object element; returns true if handled */
  DR.onObjectDown = function (e, objEl) {
    const wrap = objEl.closest('.flt');
    /* presses anywhere in a text box's text area (on the text or the space around it) edit the text: the
       caret goes there and a drag selects; the box's border and margins move it */
    if (e.target.closest && e.target.closest('.tbx')) { DR.clearSelection(); return false; }
    if (!wrap) {
      /* inline shape or chart: select as an inline object */
      if (objEl.classList.contains('shp') && objEl.closest('.p') && !objEl.closest('.flt')) {
        e.preventDefault();
        const pc = objEl.parentNode;
        const pos = L.layout.domToPos(pc, Array.prototype.indexOf.call(pc.childNodes, objEl));
        if (pos) L.ed.selectInlineObject(pos, objEl);
        return true;
      }
      return false;
    }
    e.preventDefault();
    const a = findAnchor(wrap);
    if (!a) return true;
    DR.select(wrap, a);
    startMove(e, wrap, a);
    return true;
  };
  DR.select = function (wrap, a) {
    DR.clearSelection();
    L.ed.hideHandles();
    selWrap = wrap;
    wrap.classList.add('sel-obj');
    const it = wrap._item;
    L.ed.objSel = { p: a.p, o: a.o, it, el: wrap, float: true };
    L.ed.sel = { a: D.pos(a.p, a.o), f: D.pos(a.p, a.o + 1) };
    showHandles(wrap);
    L.bus.emit('objsel', L.ed.objSel);
    L.bus.emit('sel', L.ed.sel);
  };
  function showHandles(target) {
    const sc = L.layout.scroller;
    const box = el('div', 'obj-handles float');
    const place = () => {
      const r = target.firstChild.getBoundingClientRect(), sr = sc.getBoundingClientRect();
      box.style.left = r.left - sr.left + sc.scrollLeft + 'px';
      box.style.top = r.top - sr.top + sc.scrollTop + 'px';
      box.style.width = r.width + 'px';
      box.style.height = r.height + 'px';
    };
    place();
    for (const hn of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
      const hd = el('div', 'oh oh-' + hn);
      hd.addEventListener('pointerdown', (ev) => startResize(ev, hn, target));
      box.appendChild(hd);
    }
    const it = target._item;
    if (it && it.t === 'shape') { const rh = el('div', 'oh oh-rot'); rh.addEventListener('pointerdown', (ev) => startRotate(ev, target)); box.appendChild(rh); }
    sc.appendChild(box);
    handles = box;
    box._place = place;
  }
  DR.placeHandles = () => { if (handles && handles._place && selWrap && selWrap.isConnected) handles._place(); else if (handles && selWrap && !selWrap.isConnected) reselect(); };
  function reselect() {
    const os = L.ed.objSel;
    if (!os || !os.float) { DR.clearSelection(); return; }
    const w = Array.from(L.layout.root.querySelectorAll('.flt')).find((x) => +x.dataset.pid === os.p.id && +x.dataset.o === os.o);
    if (w) DR.select(w, { p: os.p, o: os.o }); else DR.clearSelection();
  }
  L.bus.on('layout-done', () => setTimeout(() => { if (L.ed.objSel && L.ed.objSel.float) reselect(); else if (handles) DR.clearSelection(); }, 0));
  /** convert the float's position so that it moves by (dx,dy) points */
  function moveFloat(it, dxPt, dyPt, wrap) {
    const f = it.float;
    const pos = wrap && wrap._pos;
    if (f.posH.align) { const cur = pos ? (pos.x - pos.baseX) / PX : 0; delete f.posH.align; f.posH.off = L.round(cur + dxPt, 2); if (pos) { /* baseX already includes alignment origin */ } }
    else f.posH.off = L.round((f.posH.off || 0) + dxPt, 2);
    if (f.posV.align) { const cur = pos ? (pos.y - pos.baseY) / PX : 0; delete f.posV.align; f.posV.off = L.round(cur + dyPt, 2); }
    else f.posV.off = L.round((f.posV.off || 0) + dyPt, 2);
    delete f.posH.pct; delete f.posV.pct;
  }
  function startMove(e, wrap, a) {
    const z = L.layout.zoom;
    const sx = e.clientX, sy = e.clientY;
    let moved = false;
    const ghost = el('div', 'resize-ghost');
    const mv = (ev) => {
      const dx = ev.clientX - sx, dy = ev.clientY - sy;
      if (!moved && Math.hypot(dx, dy) < 3) return;
      if (!moved) { moved = true; handles && handles.appendChild(ghost); ghost.style.width = '100%'; ghost.style.height = '100%'; }
      ghost.style.transform = `translate(${dx}px,${dy}px)`;
    };
    const up = (ev) => {
      window.removeEventListener('pointermove', mv);
      window.removeEventListener('pointerup', up);
      ghost.remove();
      if (!moved) return;
      const dxPt = (ev.clientX - sx) / z / PX, dyPt = (ev.clientY - sy) / z / PX;
      L.ed.edit('Move Object', () => {
        D.touch(a.p);
        const it = D.itemAfter(a.p, a.o);
        if (it && it.float) moveFloat(it, dxPt, dyPt, wrap);
        return { a: D.pos(a.p, a.o), f: D.pos(a.p, a.o + 1) };
      });
    };
    window.addEventListener('pointermove', mv);
    window.addEventListener('pointerup', up);
  }
  DR.nudge = function (key, pt) {
    const os = L.ed.objSel;
    if (!os || !os.float) return;
    const dx = key === 'ArrowLeft' ? -pt : key === 'ArrowRight' ? pt : 0, dy = key === 'ArrowUp' ? -pt : key === 'ArrowDown' ? pt : 0;
    L.ed.edit('Nudge', () => { D.touch(os.p); const it = D.itemAfter(os.p, os.o); if (it && it.float) moveFloat(it, dx, dy, os.el); return { a: D.pos(os.p, os.o), f: D.pos(os.p, os.o + 1) }; }, { merge: 'nudge' });
  };
  function startResize(ev, hn, target) {
    ev.preventDefault(); ev.stopPropagation();
    const os = L.ed.objSel;
    if (!os) return;
    const z = L.layout.zoom;
    const it = os.it;
    const w0 = it.w || 72, h0 = it.h || 72;
    const sx = ev.clientX, sy = ev.clientY;
    const corner = hn.length === 2;
    const ghost = el('div', 'resize-ghost');
    handles.appendChild(ghost);
    let nw = w0, nh = h0, dxl = 0, dyt = 0;
    const lockAspect = it.t === 'img' || it.t === 'chart';
    const mv = (e2) => {
      const dx = (e2.clientX - sx) / z / PX, dy = (e2.clientY - sy) / z / PX;
      nw = w0 + (hn.includes('e') ? dx : hn.includes('w') ? -dx : 0);
      nh = h0 + (hn.includes('s') ? dy : hn.includes('n') ? -dy : 0);
      if (corner && (lockAspect ? !e2.shiftKey : e2.shiftKey)) { const k = Math.max(nw / w0, nh / h0); nw = w0 * k; nh = h0 * k; }
      nw = Math.max(2, nw); nh = Math.max(2, nh);
      dxl = hn.includes('w') ? w0 - nw : 0; dyt = hn.includes('n') ? h0 - nh : 0;
      ghost.style.cssText = `left:${dxl * PX * z}px;top:${dyt * PX * z}px;width:${nw * PX * z}px;height:${nh * PX * z}px`;
    };
    const up = () => {
      window.removeEventListener('pointermove', mv);
      window.removeEventListener('pointerup', up);
      ghost.remove();
      if (Math.abs(nw - w0) < 0.3 && Math.abs(nh - h0) < 0.3) return;
      L.ed.edit('Resize Object', () => {
        D.touch(os.p);
        const item = D.itemAfter(os.p, os.o);
        if (item) {
          if (item.t === 'group') scaleGroup(item, nw / (item.w || 1), nh / (item.h || 1));
          item.w = L.round(nw, 2); item.h = L.round(nh, 2);
          if (item.float && (dxl || dyt)) moveFloat(item, dxl, dyt, target);
        }
        return { a: D.pos(os.p, os.o), f: D.pos(os.p, os.o + 1) };
      });
    };
    window.addEventListener('pointermove', mv);
    window.addEventListener('pointerup', up);
  }
  function scaleGroup(g, kx, ky) { for (const k of g.kids || []) { k.x = (k.x || 0) * kx; k.y = (k.y || 0) * ky; k.w *= kx; k.h *= ky; if (k.t === 'group') scaleGroup(k, kx, ky); } }
  function startRotate(ev, target) {
    ev.preventDefault(); ev.stopPropagation();
    const os = L.ed.objSel;
    if (!os) return;
    const r = target.firstChild.getBoundingClientRect();
    const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
    let ang = os.it.rot || 0;
    const mv = (e2) => { ang = (Math.atan2(e2.clientY - cy, e2.clientX - cx) * 180) / Math.PI + 90; if (!e2.shiftKey) ang = Math.round(ang / 15) * 15; target.firstChild.style.transform = `rotate(${ang}deg)`; };
    const up = () => {
      window.removeEventListener('pointermove', mv);
      window.removeEventListener('pointerup', up);
      L.ed.edit('Rotate', () => { D.touch(os.p); const it = D.itemAfter(os.p, os.o); if (it) it.rot = L.round(((ang % 360) + 360) % 360, 1) || undefined; return { a: D.pos(os.p, os.o), f: D.pos(os.p, os.o + 1) }; });
    };
    window.addEventListener('pointermove', mv);
    window.addEventListener('pointerup', up);
  }
  /** modify the selected object (inline or floating) */
  DR.modifySelected = function (label, fn) {
    const os = L.ed.objSel;
    if (!os) return false;
    L.ed.edit(label, () => {
      D.touch(os.p);
      const it = D.itemAfter(os.p, os.o);
      if (it) fn(it);
      return { a: D.pos(os.p, os.o), f: D.pos(os.p, os.o + 1) };
    });
    return true;
  };
  DR.selectedItem = () => (L.ed.objSel ? D.itemAfter(L.ed.objSel.p, L.ed.objSel.o) : null);

  /* ---------- wrapping styles (Format ▸ Picture ▸ Layout, Draw ▸ Text Wrapping) ---------- */
  DR.WRAPS = [['inline', 'In Line with Text'], ['square', 'Square'], ['tight', 'Tight'], ['behind', 'Behind Text'], ['none', 'In Front of Text'], ['topBottom', 'Top and Bottom'], ['through', 'Through']];
  DR.setWrap = function (it, wrap, anchorInfo) {
    if (wrap === 'inline') { delete it.float; return; }
    if (!it.float) {
      it.float = { posH: { rel: 'column', off: 0 }, posV: { rel: 'paragraph', off: 0 }, dist: { t: 0, b: 0, l: 9, r: 9 }, z: 251659264, allowOverlap: true, layoutInCell: true };
      if (anchorInfo && anchorInfo.x != null) { it.float.posH.off = anchorInfo.x; it.float.posV.off = anchorInfo.y || 0; }
    }
    it.float.wrap = wrap;
    it.float.behind = wrap === 'behind';
  };
  /** default AutoShape / text box item */
  DR.newShape = function (geom, w, h, opts) {
    opts = opts || {};
    const isLine = geom === 'line' || L.geom.isLine(geom);
    const it = { t: 'shape', geom, w, h, fill: isLine ? { t: 'none' } : { t: 'solid', c: '#FFFFFF', a: 1 }, line: { c: '#000000', w: 0.75, dash: 'solid' } };
    if (opts.arrow) it.line.tail = { type: 'triangle', w: 'med', len: 'med' };
    if (opts.dblArrow) { it.line.tail = { type: 'triangle', w: 'med', len: 'med' }; it.line.head = { type: 'triangle', w: 'med', len: 'med' }; }
    if (opts.textbox) { it.tb = { kind: 'tb', blocks: [D.para()] }; it.ins = { l: 7.2, t: 3.6, r: 7.2, b: 3.6 }; }
    return it;
  };
})();
