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

  /* ---------- patterns (8×8 bitmaps, OOXML prst names; common/geometry.js) ---------- */
  const PATT = L.geom.PATTERNS;
  DR.PATTERNS = PATT;
  const pattPath = L.geom.patternPath;
  DR.patternSVG = (prst, fg, bg, tile) => `<svg xmlns="http://www.w3.org/2000/svg" width="${tile || 6}" height="${tile || 6}" viewBox="0 0 8 8" shape-rendering="crispEdges"><rect width="8" height="8" fill="${bg}"/><path d="${pattPath(PATT[prst] || PATT.pct50)}" fill="${fg}"/></svg>`;
  /** a gradient's stops in order, in their colours, blended as Office blends them (L.color.lightStops) */
  const gradStops = (fill) => L.color.lightStops((fill.stops || []).slice().sort((a, b) => a.p - b.p).map((st) => ({ p: st.p, c: col(st.c), a: st.a })));
  DR.fillCSS = function (fill) {
    if (!fill || fill.t === 'none') return 'transparent';
    if (fill.t === 'solid') return L.color.rgba(col(fill.c), fill.a);
    if (fill.t === 'grad') {
      const stops = gradStops(fill).map((st) => `${L.color.rgba(st.c, st.a)} ${L.round(st.p * 100, 2)}%`).join(', ');
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
      for (const st of gradStops(fill)) g.appendChild(s('stop', { offset: L.round(st.p, 4), 'stop-color': st.c, 'stop-opacity': st.a == null ? 1 : L.round(st.a, 3) }));
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
  const marker = (defs, end, color, lw, isHead) => L.geom.marker(defs, end, color, lw, isHead);   // common/geometry.js
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
  DR.WARPS = [['textPlain', L.N('Plain Text')], ['textArchUp', L.N('Arch Up')], ['textArchDown', L.N('Arch Down')], ['textCircle', L.N('Circle')], ['textWave1', L.N('Wave')], ['textSlantUp', L.N('Slant Up')], ['textSlantDown', L.N('Slant Down')], ['textTriangle', L.N('Triangle Up')], ['textChevron', L.N('Chevron Up')], ['textCanUp', L.N('Can Up')]];
  /** LettersArt: its lines at the size they are saved with (docx-write.js: centred), bent by its warp
      (L.geom.warpSVG, as PowerPoint and Word bend text) */
  DR.wordartSVG = function (sh) {
    const wa = sh.wordart || { text: 'LettersArt' };
    const w = Math.max(sh.w, 1), hh = Math.max(sh.h, 1), sz = wa.sz || 36;
    const run = { size: sz, family: L.fontStack(wa.font || 'Arial Black'), bold: !!wa.b, italic: !!wa.i, paint: 0 };
    const lines = String(wa.text || '').split('\n').map((t) => ({ runs: t ? [Object.assign({ text: t }, run)] : [], algn: 'ctr', size: sz, pitch: sz * 1.2 }));
    const fill = sh.fill && sh.fill.t !== 'none' ? sh.fill : { t: 'solid', c: '#3366CC' }, ln = sh.line && sh.line.t !== 'none' && sh.line.c ? sh.line : null;
    const svg = L.geom.warpSVG(sh.warp || wa.warp || 'textPlain', sh.warpAdj, w, hh, lines, (key, defs) => {
      const out = { fill: null, stroke: null };
      if (fill.t === 'solid') Object.assign(out, { fill: col(fill.c), fillOpacity: fill.a });
      else { const p = svgPaint(fill, defs); Object.assign(out, { fill: p.paint, fillOpacity: p.opacity }); }
      if (ln) Object.assign(out, { stroke: col(ln.c), strokeWidth: Math.max(0.25, ln.w || 0.75), strokeOpacity: ln.a });
      return out;
    }) || s('svg', { width: 1, height: 1 });
    svg.setAttribute('class', 'wa');
    svg.setAttribute('width', L.round(w * PX, 2)); svg.setAttribute('height', L.round(hh * PX, 2));
    if (sh.shadow) svg.style.filter = shadowCSS(sh.shadow);
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
      /* as long as the lines (the box's height in vertical text), so the anchor places it across them */
      const inner = el('div', 'tbi', 'position:relative;inline-size:100%');
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
        if (it.warp && !it.vert) DR.warpBox(box, tbx, it);
        if (it.autofit && !opts.inGroup) {
          /* resize-shape-to-fit-text: the box is as tall as its text, grown or shrunk from the size saved (Word fits
             it again when it lays the page out) */
          const need = inner.offsetHeight + px(ins.t + ins.b);
          if (Math.abs(need - hh) > 1) { box.style.height = need + 'px'; tbx.style.height = inner.offsetHeight + 'px'; const svg = box.querySelector('svg.geo'); if (svg) { const nh = need / PX; svg.setAttribute('height', need); svg.setAttribute('viewBox', `0 0 ${L.round(Math.max(it.w, 1), 3)} ${L.round(nh, 3)}`); const parts = L.geom.parts(Object.assign({}, it, { h: nh }), it.w, nh); const ps = svg.querySelectorAll('g > path'); parts.forEach((pp, i) => { if (ps[i]) ps[i].setAttribute('d', pp.d); }); } }
        }
      };
    }
    return box;
  };
  /** a text box's text bent by its warp (bodyPr's prstTxWarp, L.geom.warpSVG) over the whole shape, from its lines
      as laid out: the characters on each line with their fonts and colours as drawn. The laid-out text stays under it
      for editing, shown in its place while the caret is in the box. */
  DR.warpBox = function (box, tbx, it) {
    const old = box.querySelector(':scope > svg.warp');
    if (old) old.remove();
    const lines = [], walk = document.createTreeWalker(tbx, NodeFilter.SHOW_TEXT);
    let line = null, n = 0;
    for (let t; (t = walk.nextNode()) && n < 4000;) {
      const pe = t.parentElement, para = (pe && pe.closest('.p')) || tbx, cs = pe && getComputedStyle(pe);
      if (!cs || cs.display === 'none' || cs.visibility === 'hidden') continue;
      const style = { family: cs.fontFamily, size: parseFloat(cs.fontSize), bold: +cs.fontWeight >= 600, italic: cs.fontStyle !== 'normal',
        paint: { color: /rgba\([^)]*,\s*0\)|transparent/.test(cs.color) && /gradient/.test(cs.backgroundImage) ? (cs.backgroundImage.match(/rgba?\([^)]*\)/) || [cs.color])[0] : cs.color, stroke: parseFloat(cs.webkitTextStrokeWidth) > 0 ? [cs.webkitTextStrokeColor, parseFloat(cs.webkitTextStrokeWidth)] : null } };
      const r = document.createRange();
      let i = 0;
      for (const ch of Array.from(t.data)) {
        r.setStart(t, i); r.setEnd(t, i + ch.length); i += ch.length; n++;
        const q = r.getClientRects()[0];
        if (!q) continue;
        /* a new line where the text moves down a line, or a new paragraph starts */
        if (!line || line.para !== para || q.top > line.top + line.size * 0.6) {
          line = { para, top: q.top, size: style.size, runs: [], algn: { center: 'ctr', right: 'r', end: 'r' }[getComputedStyle(pe.closest('.pc') || para).textAlign] || 'l' };
          lines.push(line);
        }
        const last = line.runs[line.runs.length - 1];
        if (last && JSON.stringify([last.family, last.size, last.bold, last.italic, last.paint]) === JSON.stringify([style.family, style.size, style.bold, style.italic, style.paint])) last.text += ch;
        else line.runs.push(Object.assign({ text: ch }, style));
        line.size = Math.max(line.size, style.size);
      }
    }
    lines.forEach((ln, k) => { ln.runs = ln.runs.filter((x, j, all) => j || x.text.trim() || all.length > 1); ln.pitch = k < lines.length - 1 ? Math.max(1, lines[k + 1].top - ln.top) : ln.size * 1.2; });
    const w = Math.max(1, px(it.w)), hh = Math.max(1, px(it.h));
    const svg = lines.length && L.geom.warpSVG(it.warp, it.warpAdj, w, hh, lines, ({ color, stroke }) => ({ fill: color, stroke: stroke && stroke[0], strokeWidth: stroke && stroke[1] }));
    if (!svg) return;
    svg.style.cssText = 'position:absolute;left:0;top:0';
    box.classList.add('warped');
    box.appendChild(svg);
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
    R.effectRoom(box, it);
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
        /* stacked by relativeHeight inside the page's layer, every step kept (Word's VML z-index counts by one) */
        wrap.style.cssText = `position:absolute;left:${pos.x}px;top:${pos.y}px;z-index:${Math.max(-2147483647, Math.min(2147483647, Math.round((f.z || 0) - 251658240)))}`;
        wrap._pos = pos;
        wrap._frag = frag;
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
      /* the room for the object's effects (wp:effectExtent) adds to its distance from the text, as Word keeps it */
      const e = wrap._item.eff || { l: 0, t: 0, r: 0, b: 0 };
      const t = px((dist.t || 0) + e.t), b = px((dist.b || 0) + e.b);
      const dy = Math.max(0, pos.y - pcY - t);
      const leftSide = pos.x + pos.ow / 2 < pcX + pcW / 2 || f.side === 'right';
      const side = f.wrap === 'topBottom' || leftSide ? 'left' : 'right';
      const sp = el('span', 'fl-sp', `float:${side};width:0;height:${Math.max(0, many ? dy - cursor : dy)}px${many ? ';clear:both' : ''}`);
      sp.contentEditable = 'false';
      if (f.wrap === 'topBottom') {
        wrap.style.cssText = `float:left;clear:${many ? 'both' : 'left'};width:100%;height:${t + pos.oh + b}px;position:relative;padding-top:${t}px`;
        obj.style.position = 'absolute';
        obj.style.left = pos.x - pcX + 'px';
        obj.style.top = t + 'px';
      } else if (leftSide) {
        wrap.style.cssText = `float:left;clear:${many ? 'both' : 'left'};margin-left:${pos.x - pcX}px;margin-right:${px((dist.r != null ? dist.r : 9) + e.r)}px;margin-bottom:${b}px;margin-top:${t}px;position:relative`;
      } else {
        wrap.style.cssText = `float:right;clear:${many ? 'both' : 'right'};margin-right:${pcX + pcW - pos.x - pos.ow}px;margin-left:${px((dist.l != null ? dist.l : 9) + e.l)}px;margin-bottom:${b}px;margin-top:${t}px;position:relative`;
      }
      wrap._pos = pos;
      pc.insertBefore(sp, ref);
      pc.insertBefore(wrap, ref);
      DR.fillPending(wrap);
      /* a text box that grows to fit its text takes its real height in the wrap */
      let oh = pos.oh;
      if (obj.querySelector('.tbx')) { const real = obj.offsetHeight; if (Math.abs(real - oh) > 0.5) { oh = real; if (f.wrap === 'topBottom') wrap.style.height = t + oh + b + 'px'; } }
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
      const t = el('div', null, `font-family:${L.fontStack(wm.font || 'Times New Roman').replace(/\x22/g, "'")};font-size:${fs}px;color:#${wm.color || 'C0C0C0'};white-space:nowrap;transform:rotate(${diag ? -45 : 0}deg);opacity:${wm.semi === false ? 1 : 0.5};line-height:1`);
      t.textContent = wm.text;
      box.appendChild(t);
    } else if (wm.type === 'picture' && wm.media) {
      const im = el('img', null, `max-width:${page.bodyW}px;max-height:${page.H * 0.7}px`);
      im.src = L.media.url(wm.media);
      if (wm.scale) { im.style.width = px((wm.w || 200) * wm.scale) + 'px'; }
      const lum = wm.washout !== false && L.lumFilter(0.7, -0.7);   // Word's Washout
      if (lum) { im.style.filter = lum.css; box.appendChild(lum.defs); }
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
    const tbx = e.target.closest && e.target.closest('.tbx');
    if (tbx) {
      DR.clearSelection();
      /* in the box's empty space (above or below its lines) the browser doesn't put the caret in the box: its
         nearest line takes it */
      if (e.button === 0 && !e.shiftKey && !e.target.closest('.p') && document.caretRangeFromPoint) {
        const b = (tbx.firstElementChild || tbx).getBoundingClientRect();
        const r = document.caretRangeFromPoint(L.clamp(e.clientX, b.left + 1, b.right - 1), L.clamp(e.clientY, b.top + 1, b.bottom - 1));
        if (r && tbx.contains(r.startContainer)) {
          e.preventDefault();
          L.layout.root.focus({ preventScroll: true });
          window.getSelection().setBaseAndExtent(r.startContainer, r.startOffset, r.startContainer, r.startOffset);
          L.ed.syncSel();
        }
      }
      return false;
    }
    if (!wrap) {
      /* inline shape or chart: select as an inline object */
      if (objEl.classList.contains('shp') && objEl.closest('.p') && !objEl.closest('.flt')) {
        e.preventDefault();
        if (e.button !== 0 && L.ed.objSel && L.ed.objSel.el === objEl) return true;
        const pc = objEl.parentNode;
        const pos = L.layout.domToPos(pc, Array.prototype.indexOf.call(pc.childNodes, objEl));
        if (pos) { L.ed.selectInlineObject(pos, objEl); if (e.button === 0) L.ed.startObjectMove(e, pos); }
        return true;
      }
      return false;
    }
    e.preventDefault();
    const a = findAnchor(wrap);
    if (!a) return true;
    if (e.button !== 0 && L.ed.objSel && L.ed.objSel.el === wrap) return true;   /* a right-click keeps it selected */
    DR.select(wrap, a);
    if (e.button === 0) startMove(e, wrap, a);
    return true;
  };
  DR.select = function (wrap, a) {
    DR.clearSelection();
    L.ed.selectObject({ p: a.p, o: a.o, it: wrap._item, el: wrap, float: true });
    selWrap = wrap;
    wrap.classList.add('sel-obj');
    showHandles(wrap);
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
    /* a right-click on a handle: the object's menu, as anywhere else on it */
    box.addEventListener('contextmenu', (ev) => { ev.preventDefault(); ev.stopPropagation(); if (L.app && L.app.contextMenu) L.app.contextMenu(ev); });
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
  L.bus.on('objsel', (os) => { if (!os || !os.float) DR.clearSelection(); });
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
  /** a press on a floating object: dragging it moves it (an outline follows the pointer; the object moves on release) */
  function startMove(e, wrap, a) {
    const z = L.layout.zoom;
    const ghost = el('div', 'resize-ghost');
    L.ui.drag(e, {
      move(ev, d) {
        if (!ghost.parentNode) { if (!handles) return; handles.appendChild(ghost); ghost.style.width = '100%'; ghost.style.height = '100%'; }
        ghost.style.transform = `translate(${d.dx}px,${d.dy}px)`;
      },
      up(ev, d) {
        ghost.remove();
        if (!d.moved) return;
        const dxPt = (ev.clientX - d.x0) / z / PX, dyPt = (ev.clientY - d.y0) / z / PX;
        L.ed.edit('Move Object', () => {
          D.touch(a.p);
          const it = D.itemAfter(a.p, a.o);
          if (it && it.float) moveFloat(it, dxPt, dyPt, wrap);
          return { a: D.pos(a.p, a.o), f: D.pos(a.p, a.o + 1) };
        });
      },
      cancel: () => ghost.remove(),
    });
  }
  DR.nudge = function (key, pt) {
    const os = L.ed.objSel;
    if (!os || !os.float) return;
    const dx = key === 'ArrowLeft' ? -pt : key === 'ArrowRight' ? pt : 0, dy = key === 'ArrowUp' ? -pt : key === 'ArrowDown' ? pt : 0;
    L.ed.edit(L.N('Nudge'), () => { D.touch(os.p); const it = D.itemAfter(os.p, os.o); if (it && it.float) moveFloat(it, dx, dy, os.el); return { a: D.pos(os.p, os.o), f: D.pos(os.p, os.o + 1) }; }, { merge: 'nudge' });
  };
  function startResize(ev, hn, target) {
    ev.preventDefault(); ev.stopPropagation();
    const os = L.ed.objSel;
    if (ev.button !== 0 || !os) return;
    const z = L.layout.zoom;
    const it = os.it;
    const w0 = it.w || 72, h0 = it.h || 72;
    const corner = hn.length === 2;
    const ghost = el('div', 'resize-ghost');
    let nw = w0, nh = h0, dxl = 0, dyt = 0;
    const lockAspect = it.t === 'img' || it.t === 'chart';
    L.ui.drag(ev, {
      move(e2, d) {
        if (!ghost.parentNode && handles) handles.appendChild(ghost);
        const dx = d.dx / z / PX, dy = d.dy / z / PX;
        nw = w0 + (hn.includes('e') ? dx : hn.includes('w') ? -dx : 0);
        nh = h0 + (hn.includes('s') ? dy : hn.includes('n') ? -dy : 0);
        if (corner && (lockAspect ? !e2.shiftKey : e2.shiftKey)) { const k = Math.max(nw / w0, nh / h0); nw = w0 * k; nh = h0 * k; }
        nw = Math.max(2, nw); nh = Math.max(2, nh);
        dxl = hn.includes('w') ? w0 - nw : 0; dyt = hn.includes('n') ? h0 - nh : 0;
        ghost.style.cssText = `left:${dxl * PX * z}px;top:${dyt * PX * z}px;width:${nw * PX * z}px;height:${nh * PX * z}px`;
      },
      up(e2, d) {
        ghost.remove();
        if (!d.moved || (Math.abs(nw - w0) < 0.3 && Math.abs(nh - h0) < 0.3)) return;
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
      },
      cancel: () => ghost.remove(),
      capture: L.layout.scroller,
    });
  }
  function scaleGroup(g, kx, ky) { for (const k of g.kids || []) { k.x = (k.x || 0) * kx; k.y = (k.y || 0) * ky; k.w *= kx; k.h *= ky; if (k.t === 'group') scaleGroup(k, kx, ky); } }
  function startRotate(ev, target) {
    ev.preventDefault(); ev.stopPropagation();
    const os = L.ed.objSel;
    if (ev.button !== 0 || !os) return;
    const shape = target.firstChild;
    const r = shape.getBoundingClientRect();
    const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
    const transform0 = shape.style.transform;
    let ang = os.it.rot || 0;
    L.ui.drag(ev, {
      move(e2) { ang = (Math.atan2(e2.clientY - cy, e2.clientX - cx) * 180) / Math.PI + 90; if (!e2.shiftKey) ang = Math.round(ang / 15) * 15; shape.style.transform = `rotate(${ang}deg)`; },
      /* a click on the handle that never moved changes nothing */
      up(e2, d) {
        if (!d.moved) return;
        L.ed.edit(L.N('Rotate'), () => { D.touch(os.p); const it = D.itemAfter(os.p, os.o); if (it) it.rot = L.round(((ang % 360) + 360) % 360, 1) || undefined; return { a: D.pos(os.p, os.o), f: D.pos(os.p, os.o + 1) }; });
      },
      cancel: () => { shape.style.transform = transform0; },
      capture: L.layout.scroller,
    });
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
  DR.WRAPS = [['inline', L.N('In Line with Text')], ['square', L.N('Square')], ['tight', L.N('Tight')], ['behind', L.N('Behind Text')], ['none', L.N('In Front of Text')], ['topBottom', L.N('Top and Bottom')], ['through', L.N('Through')]];
  DR.setWrap = function (it, wrap, anchorInfo) {
    if (wrap === 'inline') { delete it.float; return; }
    if (!it.float) {
      it.float = { posH: { rel: 'column', off: 0 }, posV: { rel: 'paragraph', off: 0 }, dist: { t: 0, b: 0, l: 9, r: 9 }, z: 251659264, allowOverlap: true, layoutInCell: true };
      if (anchorInfo && anchorInfo.x != null) { it.float.posH.off = anchorInfo.x; it.float.posV.off = anchorInfo.y || 0; }
    }
    it.float.wrap = wrap;
    it.float.behind = wrap === 'behind';
  };
  /** a shape that can hold text: AutoShapes and text boxes (not lines, LettersArt, the boxes of a linked story) */
  DR.takesText = (it) => !!it && it.t === 'shape' && !it.wordart && !it.tbLink && it.geom !== 'line' && !L.geom.isLine(it.geom);
  /** Word's Add Text / Edit Text on the selected shape: a caret at the end of its text, typing `text` there if given;
      a shape without text first gets a paragraph centred in it, as Word starts one (an undo step of its own) */
  DR.editText = function (text) {
    const E = L.ed, os = E.objSel, it = DR.selectedItem();
    if (!os || !DR.takesText(it)) return false;
    const end = () => { const p = D.lastPara(it.tb.blocks); return D.pos(p, D.plen(p)); };
    let pos = null;
    if (!it.tb || text) {
      E.edit(it.tb ? L.N('Typing') : L.N('Add Text'), () => {
        if (!it.tb) {
          D.touch(os.p);
          it.tb = { kind: 'tb', blocks: [D.para([], { jc: 'center' })] };
          it.txBox = false;
          if (!it.anchor) it.anchor = 'ctr';
          if (!it.ins) it.ins = { l: 7.2, t: 3.6, r: 7.2, b: 3.6 };
          D.reindex(D.doc);
        }
        pos = text ? L.O.insertText(end(), text) : end();
        return pos;
      }, it.tb ? { merge: 'type' } : undefined);
    }
    /* the caret in its text ends the object selection (so the next keys type there) */
    E.setSel(pos || end());
    E.focus();
    return true;
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
