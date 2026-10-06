/* Lectern — drawing of charts read from PowerPoint 2007-and-later files (chart model v2, see chartModel in
 * pptx-read.js). Follows Office's conventions: automatic axis scaling, gap width and overlap, markers,
 * data labels, secondary axes, reversed and logarithmic axes, manual layout of the plot area and legend,
 * pie explosion, doughnut holes, radar webs and bubbles. Charts made in Lectern itself (model v1) keep the
 * PowerPoint 2003 look drawn by charts.js. */
(function () {
  'use strict';
  const L = window.L;
  const CH = L.chart;
  const s = L.s;
  const R = () => L.render;

  /* ---------- text ---------- */
  let mc = null;
  const measure = (t, st) => {
    if (!mc) mc = document.createElement('canvas').getContext('2d');
    mc.font = `${st.i ? 'italic ' : ''}${st.b ? 'bold ' : ''}${st.sz}px ${st.stack}`;
    let w = 0;
    for (const line of String(t).split('\n')) w = Math.max(w, mc.measureText(line).width);
    return w;
  };
  const lineH = (st) => st.sz * 1.2;
  function style(base, o, design) {
    o = o || {};
    const font = o.font || base.font;
    return { sz: o.sz || base.sz, b: o.b != null ? !!o.b : !!base.b, i: o.i != null ? !!o.i : !!base.i, color: R().col(o.color || base.color, design, '#000000'), stack: L.fontStack(L.style.font(font, design)), rot: o.rot };
  }
  /** wrap text into lines no wider than w */
  function wrap(t, st, w) {
    const out = [];
    for (const para of String(t).split('\n')) {
      const words = para.split(/(\s+)/).filter((x) => x !== '');
      let cur = '';
      for (const wd of words) {
        const next = cur + wd;
        if (cur.trim() && measure(next.trimEnd(), st) > w && /\S/.test(wd)) { out.push(cur.trimEnd()); cur = wd.trimStart(); } else cur = next;
      }
      out.push(cur.trimEnd());
    }
    return out;
  }
  /** text centred (anchor middle/start/end) on x; y is the vertical centre of the block */
  function text(parent, lines, x, y, st, anchor, rot) {
    if (!Array.isArray(lines)) lines = String(lines).split('\n');
    const lh = lineH(st);
    const y0 = y - ((lines.length - 1) * lh) / 2 + st.sz * 0.35;
    const t = s('text', { x: L.round(x, 2), y: L.round(y0, 2), 'text-anchor': anchor || 'middle', 'font-size': L.round(st.sz, 2), fill: st.color, 'font-family': st.stack });
    if (st.b) t.setAttribute('font-weight', 700);
    if (st.i) t.setAttribute('font-style', 'italic');
    if (rot) t.setAttribute('transform', `rotate(${L.round(rot, 2)} ${L.round(x, 2)} ${L.round(y, 2)})`);
    lines.forEach((ln, k) => {
      const sp = s('tspan', { x: L.round(x, 2), dy: k ? L.round(lh, 2) : 0 });
      sp.textContent = ln;
      t.appendChild(sp);
    });
    parent.appendChild(t);
    return t;
  }

  /* ---------- numbers ---------- */
  const general = (v) => (L.numfmt ? L.numfmt.general(v, 11) : String(L.round(v, 9)));
  const fmtVal = (v, code) => {
    if (v == null || !isFinite(v)) return '';
    if (code && code !== 'General' && L.numfmt) { try { return L.numfmt.text(code, v); } catch (e) { /* fall through */ } }
    return general(v);
  };
  function niceStep(raw) {
    if (!(raw > 0)) return 1;
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const n = raw / mag;
    return (n <= 1.0000001 ? 1 : n <= 2.0000001 ? 2 : n <= 5.0000001 ? 5 : 10) * mag;
  }
  /**
   * Office's automatic value-axis scale: zero is included unless the data sit far from it (minimum above 5/6
   * of the maximum), the top gets about 5 % headroom, and the major unit is the smallest 1-2-5 step that
   * keeps the number of intervals readable for the axis length.
   */
  function scale(lo, hi, ax, len, labelSz, opts) {
    ax = ax || {};
    opts = opts || {};
    if (ax.log) {
      const b = ax.log || 10;
      let mn = ax.min != null && ax.min > 0 ? ax.min : Math.pow(b, Math.floor(Math.log(Math.max(lo, 1e-12)) / Math.log(b)));
      let mx = ax.max != null ? ax.max : Math.pow(b, Math.ceil(Math.log(Math.max(hi, mn * b)) / Math.log(b)));
      if (!(mx > mn)) mx = mn * b;
      const ticks = [];
      for (let v = mn; v <= mx * 1.0001; v *= b) ticks.push(v);
      return { min: mn, max: mx, ticks, log: b };
    }
    if (!isFinite(lo) || !isFinite(hi)) { lo = 0; hi = 1; }
    if (lo > hi) [lo, hi] = [hi, lo];
    const fixMin = ax.min != null, fixMax = ax.max != null;
    /* zero is on the axis unless the data keep well away from it; then the far end gets half the data range */
    if (!opts.noZero) {
      if (lo >= 0 && hi > 0) { if (!fixMin) { if (lo > hi * 5 / 6) lo = Math.max(0, lo - (hi - lo) / 2); else lo = 0; } }
      else if (hi <= 0 && lo < 0) { if (!fixMax) { if (hi < lo * 5 / 6) hi = Math.min(0, hi + (hi - lo) / 2); else hi = 0; } }
    }
    if (fixMin && !fixMax && hi <= ax.min) hi = ax.min + Math.abs(ax.min || 1);
    if (fixMax && !fixMin && lo >= ax.max) lo = ax.max - Math.abs(ax.max || 1);
    if (lo === hi) { if (lo === 0) hi = 1; else if (lo > 0) lo = 0; else hi = 0; }
    /* Office allows at most ten major intervals, fewer on a short axis */
    const maxN = Math.max(2, Math.min(10, Math.floor(len / Math.max(labelSz * 1.6, 10))));
    const lo0 = fixMin ? ax.min : lo, hi0 = fixMax ? ax.max : hi;
    const head = opts.noHead ? 0 : (hi0 - lo0) * 0.05;
    const hiH = fixMax ? hi0 : hi0 > 0 ? hi0 + head : hi0;
    const loH = fixMin ? lo0 : lo0 < 0 ? lo0 - head : lo0;
    let step = ax.major > 0 ? ax.major : niceStep((hiH - loH) / maxN || 1);
    const mx = fixMax ? ax.max : Math.ceil(hiH / step - 1e-9) * step;
    const mn = fixMin ? ax.min : Math.floor(loH / step + 1e-9) * step;
    let maxV = mx, minV = mn;
    if (!(maxV > minV)) maxV = minV + (step || 1);
    if (!(step > 0) || (maxV - minV) / step > 200) step = niceStep((maxV - minV) / maxN);
    const ticks = [];
    const first = Math.ceil(minV / step - 1e-9) * step;
    for (let v = first; v <= maxV + step * 1e-6; v += step) ticks.push(Math.abs(v) < step * 1e-9 ? 0 : parseFloat(v.toPrecision(12)));
    return { min: minV, max: maxV, ticks, step };
  }

  /* ---------- paint ---------- */
  const AUTO_LINE = '#868686';
  function strokeOf(line, design, def) {
    /* def: {c, w} automatic line; line: the explicit one ({t:'none'}, or with c / w left out when automatic) */
    if (line && line.t === 'none') return null;
    if (!line && !def) return null;
    const l = Object.assign({}, def || {}, line || {});
    if (!l.c) return null;
    const w = l.w != null ? l.w : 0.75;
    return { color: R().col(l.c, design, '#000000'), a: l.a == null ? 1 : l.a, w: Math.max(0.25, w), dash: l.dash && l.dash !== 'solid' ? R().dashArray(l.dash, w) : null, cap: l.cap === 'rnd' ? 'round' : null };
  }
  function applyStroke(el, st) {
    if (!st) { el.setAttribute('stroke', 'none'); return el; }
    el.setAttribute('stroke', st.color);
    el.setAttribute('stroke-width', L.round(st.w, 2));
    if (st.a < 1) el.setAttribute('stroke-opacity', L.round(st.a, 3));
    if (st.dash) el.setAttribute('stroke-dasharray', st.dash);
    if (st.cap) el.setAttribute('stroke-linecap', st.cap);
    return el;
  }
  function applyFill(el, fill, design, defs) {
    if (!fill || fill.t === 'none') { el.setAttribute('fill', 'none'); return el; }
    const p = R().svgPaint(fill, design, defs);
    el.setAttribute('fill', p.paint);
    if (p.opacity < 1) el.setAttribute('fill-opacity', L.round(p.opacity, 3));
    return el;
  }
  const solid = (c) => ({ t: 'solid', c, a: 1 });

  /* ---------- markers ---------- */
  const AUTO_SYMBOLS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus'];
  function marker(parent, sym, x, y, size, fill, stroke, design, defs) {
    const r = size / 2;
    let el;
    switch (sym) {
      case 'none': return;
      case 'circle': el = s('circle', { cx: L.round(x, 2), cy: L.round(y, 2), r: L.round(r, 2) }); break;
      case 'square': el = s('rect', { x: L.round(x - r, 2), y: L.round(y - r, 2), width: L.round(size, 2), height: L.round(size, 2) }); break;
      case 'diamond': el = s('path', { d: `M${x},${y - r}L${x + r},${y}L${x},${y + r}L${x - r},${y}Z` }); break;
      case 'triangle': el = s('path', { d: `M${x},${y - r}L${x + r},${y + r}L${x - r},${y + r}Z` }); break;
      case 'x': el = s('path', { d: `M${x - r},${y - r}L${x + r},${y + r}M${x + r},${y - r}L${x - r},${y + r}` }); fill = null; break;
      case 'star': el = s('path', { d: `M${x - r},${y - r}L${x + r},${y + r}M${x + r},${y - r}L${x - r},${y + r}M${x},${y - r}L${x},${y + r}` }); fill = null; break;
      case 'plus': el = s('path', { d: `M${x - r},${y}L${x + r},${y}M${x},${y - r}L${x},${y + r}` }); fill = null; break;
      case 'dot': el = s('circle', { cx: x, cy: y, r: Math.max(0.75, r / 2.5) }); break;
      case 'dash': el = s('rect', { x: x - r, y: y - Math.max(0.5, r / 5), width: size, height: Math.max(1, r / 2.5) }); break;
      default: el = s('rect', { x: x - r, y: y - r, width: size, height: size });
    }
    applyFill(el, fill, design, defs);
    if (!fill && stroke) applyStroke(el, Object.assign({}, stroke, { w: Math.max(stroke.w, 0.75) }));
    else applyStroke(el, stroke);
    parent.appendChild(el);
  }

  /* ---------- curves ---------- */
  function pathOf(pts, smooth) {
    const segs = [];
    let cur = [];
    for (const p of pts) { if (p) cur.push(p); else if (cur.length) { segs.push(cur); cur = []; } }
    if (cur.length) segs.push(cur);
    let d = '';
    for (const sg of segs) {
      d += `M${L.round(sg[0][0], 2)},${L.round(sg[0][1], 2)}`;
      for (let i = 1; i < sg.length; i++) {
        if (smooth && sg.length > 2) {
          const p0 = sg[i - 2] || sg[i - 1], p1 = sg[i - 1], p2 = sg[i], p3 = sg[i + 1] || p2;
          const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6], c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
          d += `C${L.round(c1[0], 2)},${L.round(c1[1], 2)} ${L.round(c2[0], 2)},${L.round(c2[1], 2)} ${L.round(p2[0], 2)},${L.round(p2[1], 2)}`;
        } else d += `L${L.round(sg[i][0], 2)},${L.round(sg[i][1], 2)}`;
      }
    }
    return d;
  }

  /* ================= the renderer ================= */
  CH.renderV2 = function (c, w, hh, design) {
    w = Math.max(w, 20); hh = Math.max(hh, 20);
    const svg = s('svg', { class: 'chart', width: L.round(w, 2), height: L.round(hh, 2), viewBox: `0 0 ${L.round(w, 2)} ${L.round(hh, 2)}`, overflow: 'visible' });
    const defs = s('defs');
    svg.appendChild(defs);
    const base = { sz: c.fsz || 10, b: !!c.bold, i: false, color: c.textColor || 'tx1', font: c.font };
    const TS = (o) => style(base, o, design);
    const series = c.series || [];
    const groups = c.groups && c.groups.length ? c.groups : [{ type: 'bar', dir: 'col', grouping: 'clustered', gap: 150, overlap: 0, si: series.map((_, i) => i) }];
    const cats = c.cats || [];
    const ax = c.ax || {};
    /* chart area */
    if (c.bg || c.border) {
      const st = strokeOf(c.border, design, null);
      const inset = st ? st.w / 2 : 0;
      const r = s('rect', { x: inset, y: inset, width: Math.max(0, w - 2 * inset), height: Math.max(0, hh - 2 * inset) });
      applyFill(r, c.bg, design, defs); applyStroke(r, st);
      svg.appendChild(r);
    }
    if (c.empty || !series.length) return svg;
    const pad = Math.max(4, Math.min(w, hh) * 0.025);
    let top = pad, left = pad, right = w - pad, bottom = hh - pad;

    /* ----- title ----- */
    if (c.title) {
      const st = TS(Object.assign({ sz: base.sz * 1.2, b: true }, c.titleTx || {}));
      const lines = wrap(c.title, st, (w - 2 * pad) * 0.8);
      const th = lines.length * lineH(st);
      const tw = Math.max(...lines.map((l) => measure(l, st)));
      const cy = top + th / 2 + 2;
      if (c.titleFill || c.titleLine) {
        const r = s('rect', { x: w / 2 - tw / 2 - 4, y: cy - th / 2 - 2, width: tw + 8, height: th + 4 });
        applyFill(r, c.titleFill, design, defs); applyStroke(r, strokeOf(c.titleLine, design, null));
        svg.appendChild(r);
      }
      text(svg, lines, w / 2, cy, st, 'middle');
      if (!c.titleOverlay) top += th + st.sz * 0.6;
    }

    /* ----- legend ----- */
    const kindOf = (sr) => groups[sr.g || 0] || groups[0];
    const isPieMain = groups[0].type === 'pie' || groups[0].type === 'doughnut';
    const lineLike = (g, sr) => g.type === 'line' || g.type === 'stock' || (g.type === 'scatter') || (g.type === 'radar' && g.radarStyle !== 'filled');
    const serFill = (sr) => sr.fill || solid(sr.color || sr.auto || '#4472C4');
    const serStroke = (sr, g) => {
      if (lineLike(g, sr)) return strokeOf(sr.line, design, { c: sr.color || sr.auto, w: 2.25 });
      return strokeOf(sr.line, design, null);
    };
    const markerSym = (sr, g, i) => {
      const m = sr.marker || {};
      if (m.sym && m.sym !== 'auto') return m.sym;
      if (g.type === 'line' && g.marker === false) return 'none';
      if (g.type === 'scatter' && /^(line|smooth)$/.test(g.scatterStyle || '')) return 'none';
      if (g.type === 'radar' && g.radarStyle !== 'marker') return 'none';
      if (g.type === 'stock') return 'none';
      if (g.type !== 'line' && g.type !== 'scatter' && g.type !== 'radar') return 'none';
      return AUTO_SYMBOLS[(sr.idx || i || 0) % AUTO_SYMBOLS.length];
    };
    const markerSize = (sr) => (sr.marker && sr.marker.size) || 7;
    const markerPaint = (sr, pt) => {
      const m = Object.assign({}, sr.marker || {}, (pt && pt.marker) || {});
      return { fill: m.fill || solid(sr.color || sr.auto), stroke: strokeOf(m.line, design, { c: sr.color || sr.auto, w: 0.75 }) };
    };
    const legItems = [];
    if (c.legend && c.legend !== 'none') {
      if (isPieMain) {
        const s0 = series.find((x) => (x.g || 0) === 0) || series[0];
        const g0 = groups[0];
        cats.forEach((cat, i) => {
          if (c.legendDel && c.legendDel.includes(i)) return;
          const pt = s0.pts && s0.pts[i];
          const f = (pt && pt.fill) || (g0.vary ? solid((c.pieColors && c.pieColors[i]) || s0.color) : serFill(s0));
          legItems.push({ label: String(cat), fill: f, stroke: strokeOf((pt && pt.line) || s0.line, design, null) });
        });
      } else {
        let order = series.map((sr, i) => i);
        const g0 = groups[0];
        const vertical = c.legend === 'r' || c.legend === 'l' || c.legend === 'tr';
        if (vertical && ((g0.type === 'bar' && (g0.dir === 'bar' || /stacked/i.test(g0.grouping))) || (g0.type === 'area' && /stacked/i.test(g0.grouping)))) order = order.reverse();
        /* single-series charts with varied colours list the categories */
        if (series.length === 1 && g0.vary && g0.type === 'bar') {
          cats.forEach((cat, i) => { if (!(c.legendDel && c.legendDel.includes(i))) legItems.push({ label: String(cat), fill: (series[0].pts && series[0].pts[i] && series[0].pts[i].fill) || solid(CH.autoColor ? CH.autoColor(i, cats.length, c.style, design) : series[0].color) }); });
        } else for (const i of order) {
          const sr = series[i], g = kindOf(sr);
          if (c.legendDel && c.legendDel.includes(sr.idx != null ? sr.idx : i)) continue;
          if (lineLike(g, sr)) legItems.push({ label: sr.name, line: serStroke(sr, g), sym: markerSym(sr, g, i), mk: markerPaint(sr), mkSize: Math.min(markerSize(sr), base.sz) });
          else legItems.push({ label: sr.name, fill: serFill(sr), stroke: serStroke(sr, g) });
        }
      }
    }
    if (legItems.length) {
      const st = TS(c.legendTx);
      const key = st.sz * 0.65, lh = lineH(st) * 1.05;
      const itemW = legItems.map((it) => (it.line || it.sym ? key * 2.6 : key) + st.sz * 0.4 + measure(it.label, st));
      const pos = c.legend;
      let box;
      const manual = c.legendBox && c.legendBox.edge ? { x: c.legendBox.x * w, y: c.legendBox.y * hh, w: c.legendBox.w * w, h: c.legendBox.h * hh } : null;
      const rowsFor = (avail) => {
        const rows = [];
        let row = [], rw = 0;
        itemW.forEach((iw, i) => { const ww = iw + st.sz; if (row.length && rw + ww > avail) { rows.push({ items: row, w: rw }); row = []; rw = 0; } row.push(i); rw += ww; });
        if (row.length) rows.push({ items: row, w: rw });
        return rows;
      };
      if (pos === 't' || pos === 'b') {
        const rows = rowsFor(right - left);
        const bh = rows.length * lh + st.sz * 0.4;
        const bw = Math.max(...rows.map((r) => r.w));
        box = { x: (w - bw) / 2, y: pos === 't' ? top : bottom - bh, w: bw, h: bh, rows };
        if (!c.legendOverlay) { if (pos === 't') top += bh + st.sz * 0.3; else bottom -= bh + st.sz * 0.3; }
      } else {
        const bw = Math.min((right - left) * 0.45, Math.max(...itemW) + st.sz * 0.8);
        const bh = Math.min(bottom - top, legItems.length * lh + st.sz * 0.4);
        const x = pos === 'l' ? left : right - bw;
        const y = pos === 'tr' ? top : top + (bottom - top - bh) / 2;
        box = { x, y, w: bw, h: bh };
        if (!c.legendOverlay) { if (pos === 'l') left += bw + st.sz * 0.5; else right -= bw + st.sz * 0.5; }
      }
      /* a legend moved by hand keeps its place; the plot area is laid out as if it had not moved */
      if (manual) {
        box = manual;
        if (manual.w > manual.h * 1.5 && legItems.length > 1) { const rows = rowsFor(manual.w); if (rows.length * lh <= manual.h + lh * 0.5) box.rows = rows; }
      }
      const g = s('g', { class: 'legend' });
      if (c.legendFill || c.legendLine) { const r = s('rect', { x: box.x, y: box.y, width: box.w, height: box.h }); applyFill(r, c.legendFill, design, defs); applyStroke(r, strokeOf(c.legendLine, design, null)); g.appendChild(r); }
      const drawItem = (it, x, y) => {
        if (it.line || (it.sym && it.sym !== 'none')) {
          if (it.line) g.appendChild(applyStroke(s('line', { x1: x, y1: y, x2: x + key * 2.6, y2: y }), it.line));
          if (it.sym && it.sym !== 'none') marker(g, it.sym, x + key * 1.3, y, it.mkSize || key, it.mk.fill, it.mk.stroke, design, defs);
          x += key * 2.6;
        } else {
          const r = s('rect', { x: L.round(x, 2), y: L.round(y - key / 2, 2), width: L.round(key, 2), height: L.round(key, 2) });
          applyFill(r, it.fill, design, defs); applyStroke(r, it.stroke);
          g.appendChild(r);
          x += key;
        }
        text(g, [it.label], x + st.sz * 0.35, y, st, 'start');
      };
      if (box.rows) box.rows.forEach((r, ri) => { let x = box.x + (box.w - r.w) / 2 + st.sz / 2; const y = box.y + (box.h - box.rows.length * lh) / 2 + lh * (ri + 0.5); for (const i of r.items) { drawItem(legItems[i], x, y); x += itemW[i] + st.sz; } });
      else {
        const n = legItems.length, avail = box.h - st.sz * 0.4;
        const step = Math.min(lh, avail / Math.max(n, 1));
        const y0 = box.y + (box.h - step * n) / 2 + step / 2;
        legItems.forEach((it, i) => { if (step * (i + 1) <= avail + 0.5 || i === 0) drawItem(it, box.x + st.sz * 0.4, y0 + i * step); });
      }
      svg.appendChild(g);
    }

    /* ----- plot area ----- */
    let P = { x: left + pad * 0.5, y: top + pad * 0.5, w: Math.max(4, right - left - pad), h: Math.max(4, bottom - top - pad) };
    let manualInner = false;
    if (c.plotBox && c.plotBox.edge) {
      P = { x: c.plotBox.x * w, y: c.plotBox.y * hh, w: Math.max(4, c.plotBox.w * w), h: Math.max(4, c.plotBox.h * hh) };
      manualInner = c.plotBox.inner;
    }
    const g0 = groups[0];
    if (isPieMain) { drawPie(svg, defs, c, P, design, TS, series, groups); return svg; }
    if (g0.type === 'radar') { drawRadar(svg, defs, c, P, design, TS, series, groups, cats); return svg; }

    const horiz = g0.type === 'bar' && g0.dir === 'bar';
    const xy = g0.type === 'scatter' || g0.type === 'bubble';
    /* value extents per axis (primary / secondary) */
    const ext = [{ lo: Infinity, hi: -Infinity }, { lo: Infinity, hi: -Infinity }];
    const pctAxis = [false, false];
    const nCat = xy ? 0 : Math.max(cats.length, ...series.map((sr) => sr.vals.length));
    for (const g of groups) {
      const e = ext[g.sec ? 1 : 0];
      const sis = (g.si || []).filter((i) => series[i]);
      if (/percentStacked/.test(g.grouping)) { pctAxis[g.sec ? 1 : 0] = true; e.lo = Math.min(e.lo, 0); e.hi = Math.max(e.hi, 1); continue; }
      if (/stacked/i.test(g.grouping) && (g.type === 'bar' || g.type === 'area' || g.type === 'line')) {
        for (let k = 0; k < nCat; k++) {
          let pos = 0, neg = 0;
          for (const i of sis) { const v = series[i].vals[k]; if (v == null) continue; if (v >= 0) pos += v; else neg += v; }
          e.hi = Math.max(e.hi, pos); e.lo = Math.min(e.lo, neg);
        }
      } else for (const i of sis) for (const v of series[i].vals) if (v != null && isFinite(v)) { e.hi = Math.max(e.hi, v); e.lo = Math.min(e.lo, v); }
      if (g.type === 'stock' && g.upDown) void 0;
    }
    const vAx = ax.v || {}, v2Ax = ax.v2 || null, cAx = ax.c || {};
    const unit = (a) => (a && a.unit) || 1;
    /* tick label styles */
    const vSt = TS(vAx.tx), cSt = TS(cAx.tx), v2St = v2Ax ? TS(v2Ax.tx) : null;
    const showLbl = (a) => a && !a.del && a.lblPos !== 'none';
    const valFmt = (a, sec) => {
      if (pctAxis[sec ? 1 : 0] && (!a.fmt || a.linked || a.fmt === 'General')) return '0%';
      if (a.fmt && !a.linked) return a.fmt;
      const sr = series.find((x) => !!groups[x.g || 0].sec === !!sec && x.fmt);
      return (a.linked && sr && sr.fmt) || (a.fmt && a.fmt !== 'General' ? a.fmt : null);
    };
    /* the axis length is only known after the labels are measured: two passes */
    const vLen0 = horiz ? P.w : P.h;
    let sc = scale(ext[0].lo / unit(vAx), ext[0].hi / unit(vAx), vAx, vLen0, vSt.sz, { noZero: xy && false });
    let sc2 = v2Ax ? scale(ext[1].lo / unit(v2Ax), ext[1].hi / unit(v2Ax), v2Ax, vLen0, v2St.sz) : null;
    if (pctAxis[0] && vAx.max == null) sc = scale(0, 1, Object.assign({}, vAx, { max: Math.min(ext[0].hi, 1) > 0 ? 1 : 0, min: vAx.min != null ? vAx.min : ext[0].lo < 0 ? -1 : 0 }), vLen0, vSt.sz, { noHead: true });
    /* x axis of XY charts */
    let xsc = null, xExt = { lo: Infinity, hi: -Infinity };
    if (xy) {
      for (const sr of series) (sr.xs || sr.vals.map((_, k) => k + 1)).forEach((x) => { if (x != null && isFinite(x)) { xExt.lo = Math.min(xExt.lo, x); xExt.hi = Math.max(xExt.hi, x); } });
      xsc = scale(xExt.lo, xExt.hi, cAx, P.w, cSt.sz, {});
    }
    const lblW = (scl, a, st, sec) => (showLbl(a) ? Math.max(...scl.ticks.map((t) => measure(fmtVal(t, valFmt(a, sec)), st))) : 0);
    /* sides */
    const vSide = horiz ? ((vAx.crosses === 'max') !== !!cAx.rev ? 't' : 'b') : ((vAx.crosses === 'max') !== !!cAx.rev ? 'r' : 'l');
    const cSide = horiz ? ((cAx.crosses === 'max') !== !!vAx.rev ? 'r' : 'l') : ((cAx.crosses === 'max') !== !!vAx.rev ? 't' : 'b');
    /* category labels: straight, wrapped, or turned */
    const catLabels = xy ? [] : cats.map((x) => (x == null ? '' : String(x)));
    const catGap = cSt.sz * 0.4;
    let catMode = { rot: 0, lines: catLabels.map((t) => [t]), skip: cAx.skip || 1 };
    const fitCats = (len) => {
      if (horiz) {
        const maxW = Math.max(P.w * 0.35, 10);
        const lines = catLabels.map((t) => (measure(t, cSt) > maxW ? wrap(t, cSt, maxW) : [t]));
        const band = len / Math.max(nCat, 1);
        const skip = cAx.skip || Math.max(1, Math.ceil(lineH(cSt) / Math.max(band, 0.1)));
        return { rot: 0, lines, skip, size: Math.max(0, ...lines.map((ls) => Math.max(...ls.map((t) => measure(t, cSt))))) };
      }
      const between0 = vAx.between != null ? vAx.between : g0.type !== 'area';
      const band = len / Math.max(between0 ? nCat : nCat - 1, 1);
      const widths = catLabels.map((t) => measure(t, cSt));
      const maxW = Math.max(0, ...widths);
      const explicit = cAx.tx && cAx.tx.rot != null ? cAx.tx.rot : null;
      if (explicit != null && Math.abs(explicit) > 0.5) {
        const a = (Math.abs(explicit) * Math.PI) / 180;
        const skip = cAx.skip || Math.max(1, Math.ceil(lineH(cSt) / Math.max(band * Math.sin(a), 0.1)));
        return { rot: explicit, lines: catLabels.map((t) => [t]), skip, size: maxW * Math.sin(a) + lineH(cSt) * Math.cos(a) };
      }
      if (maxW <= band * 0.95 || explicit === 0) return { rot: 0, lines: catLabels.map((t) => [t]), skip: cAx.skip || 1, size: lineH(cSt) };
      /* wrap at spaces into at most three lines when that fits the band */
      const wrapped = catLabels.map((t) => wrap(t, cSt, band * 0.95));
      if (wrapped.every((ls) => ls.length <= 3 && ls.every((x) => measure(x, cSt) <= band * 0.98))) return { rot: 0, lines: wrapped, skip: cAx.skip || 1, size: Math.max(...wrapped.map((ls) => ls.length)) * lineH(cSt) };
      const rot = band >= lineH(cSt) * 0.9 ? -45 : -90;
      const a = (Math.abs(rot) * Math.PI) / 180;
      const skip = cAx.skip || Math.max(1, Math.ceil(lineH(cSt) / Math.max(band * Math.sin(a), 0.1)));
      return { rot, lines: catLabels.map((t) => [t]), skip, size: Math.min(hh * 0.4, maxW * Math.sin(a) + lineH(cSt) * Math.cos(a)) };
    };
    const outerLevels = !xy && c.catOuter && c.catOuter.length ? c.catOuter : null;
    const d3 = !!(c.view3D && g0.d3 && g0.type === 'bar' && (c.view3D.rotX || c.view3D.rotY));
    let dep = { dx: 0, dy: 0 };
    for (let pass = 0; pass < 2; pass++) {
      const vW = lblW(sc, vAx, vSt, false), v2W = v2Ax ? lblW(sc2, v2Ax, v2St, true) : 0;
      const xW = xy && showLbl(cAx) ? Math.max(...xsc.ticks.map((t) => measure(fmtVal(t, cAx.fmt && !cAx.linked ? cAx.fmt : (series[0] && series[0].xfmt) || null), cSt))) : 0;
      if (!xy) catMode = fitCats(horiz ? P.h : P.w);
      if (manualInner) break;
      const base0 = c.plotBox && c.plotBox.edge ? { x: c.plotBox.x * w, y: c.plotBox.y * hh, w: c.plotBox.w * w, h: c.plotBox.h * hh } : { x: left + pad * 0.5, y: top + pad * 0.5, w: right - left - pad, h: bottom - top - pad };
      let l = 0, r = 0, t = 0, b = 0;
      const vTitle = (a, st0) => (a && a.title ? lineH(TS(Object.assign({ b: true }, a.titleTx))) * a.title.split('\n').length + st0.sz * 0.3 : 0);
      const outerH = outerLevels ? outerLevels.length * (lineH(cSt) + catGap) : 0;
      if (horiz) {
        const vh = showLbl(vAx) ? lineH(vSt) + vSt.sz * 0.3 : 0;
        if (vSide === 't') t += vh; else b += vh;
        b += vTitle(vAx, vSt);
        const cw = showLbl(cAx) ? catMode.size + catGap : 0;
        if (cSide === 'r') r += cw; else l += cw;
        l += vTitle(cAx, cSt) + outerH * 2;
        if (v2Ax && showLbl(v2Ax)) { if (vSide === 't') b += lineH(v2St) + v2St.sz * 0.3; else t += lineH(v2St) + v2St.sz * 0.3; }
      } else {
        const vw = showLbl(vAx) ? vW + vSt.sz * 0.4 : 0;
        if (vSide === 'r') r += vw; else l += vw;
        l += vTitle(vAx, vSt);
        if (v2Ax) { if (showLbl(v2Ax)) { if (vSide === 'r') l += v2W + v2St.sz * 0.4; else r += v2W + v2St.sz * 0.4; } r += vTitle(v2Ax, v2St); }
        const ch = xy ? (showLbl(cAx) ? lineH(cSt) + catGap : 0) : showLbl(cAx) ? catMode.size + catGap + outerH : 0;
        if (cSide === 't') t += ch; else b += ch;
        b += vTitle(cAx, cSt);
        void xW;
      }
      if (!xy && !horiz && catMode.rot < 0 && showLbl(cAx)) {
        /* turned labels hang left of their category: keep the first one inside the chart */
        const first = catLabels[0] ? measure(catLabels[0], cSt) * Math.cos((Math.abs(catMode.rot) * Math.PI) / 180) : 0;
        l = Math.max(l, first - (base0.w - l - r) / Math.max(nCat, 1) / 2);
      }
      if (xy && showLbl(cAx)) { r = Math.max(r, xW / 2); l = Math.max(l, xW / 2); }
      P = { x: base0.x + l, y: base0.y + t, w: Math.max(4, base0.w - l - r), h: Math.max(4, base0.h - t - b) };
      /* 3-D columns and bars: room for the depth, seen obliquely from the view angles */
      if (d3) {
        const nDeep = g0.deep ? Math.max(1, (g0.si || []).length) : 1;
        const D = Math.min((horiz ? P.h : P.w) / Math.max(nCat, 1) * 0.45 * nDeep, P.h * (g0.deep ? 0.32 : 0.18), P.w * (g0.deep ? 0.32 : 0.18)) * Math.max(0.2, Math.min(2, (c.view3D.depth || 100) / 100));
        const vx = Math.sin((c.view3D.rotY * Math.PI) / 180), vy = Math.sin((c.view3D.rotX * Math.PI) / 180), n2 = Math.hypot(vx, vy) || 1;
        dep = { dx: (D * vx) / n2, dy: (-D * vy) / n2 };
        if (dep.dx > 0) P.w -= dep.dx; else { P.x -= dep.dx; P.w += dep.dx; }
        if (dep.dy < 0) { P.y -= dep.dy; P.h += dep.dy; } else P.h -= dep.dy;
        P.w = Math.max(4, P.w); P.h = Math.max(4, P.h);
      }
      const len = horiz ? P.w : P.h;
      sc = pctAxis[0] && vAx.max == null ? scale(0, 1, Object.assign({}, vAx, { max: 1, min: vAx.min != null ? vAx.min : ext[0].lo < 0 ? -1 : 0 }), len, vSt.sz, { noHead: true }) : scale(ext[0].lo / unit(vAx), ext[0].hi / unit(vAx), vAx, len, vSt.sz);
      if (v2Ax) sc2 = scale(ext[1].lo / unit(v2Ax), ext[1].hi / unit(v2Ax), v2Ax, len, v2St.sz);
      if (xy) xsc = scale(xExt.lo, xExt.hi, cAx, P.w, cSt.sz, {});
    }
    if (!xy) catMode = fitCats(horiz ? P.h : P.w);

    /* ----- mapping ----- */
    const mapV = (scl, a) => (v) => {
      let t;
      if (scl.log) t = (Math.log(Math.max(v, 1e-300)) - Math.log(scl.min)) / (Math.log(scl.max) - Math.log(scl.min));
      else t = (v - scl.min) / (scl.max - scl.min || 1);
      if (a && a.rev) t = 1 - t;
      return horiz ? P.x + t * P.w : P.y + P.h - t * P.h;
    };
    const vm = mapV(sc, vAx), vm2 = sc2 ? mapV(sc2, v2Ax) : vm;
    const between = vAx.between != null ? vAx.between : g0.type !== 'area';
    const catLen = horiz ? P.h : P.w;
    const band = catLen / Math.max(between ? nCat : Math.max(nCat - 1, 1), 1);
    const catPos = (k, rev) => {
      /* position along the category axis of the centre of category k */
      let t = between ? (k + 0.5) * band : k * band;
      if (rev) t = catLen - t;
      return horiz ? P.y + P.h - t : P.x + t;
    };
    const catEdge = (k) => { let t = k * (catLen / Math.max(nCat, 1)); if (cAx.rev) t = catLen - t; return horiz ? P.y + P.h - t : P.x + t; };
    const xm = xy ? (x) => { let t = xsc.log ? (Math.log(Math.max(x, 1e-300)) - Math.log(xsc.min)) / (Math.log(xsc.max) - Math.log(xsc.min)) : (x - xsc.min) / (xsc.max - xsc.min || 1); if (cAx.rev) t = 1 - t; return P.x + t * P.w; } : null;
    const clampV = (scl, v) => Math.max(Math.min(v, Math.max(scl.min, scl.max)), Math.min(scl.min, scl.max));

    /* ----- plot background, gridlines ----- */
    if (c.plotBg || c.plotBorder) {
      const r = s('rect', { x: L.round(P.x, 2), y: L.round(P.y, 2), width: L.round(P.w, 2), height: L.round(P.h, 2) });
      applyFill(r, c.plotBg, design, defs); applyStroke(r, strokeOf(c.plotBorder, design, null));
      svg.appendChild(r);
    }
    const grid = s('g', { class: 'grid' });
    const gridLine = (pos, st, alongCat) => {
      if (!st) return;
      const l2 = alongCat !== horiz ? s('line', { x1: L.round(pos, 2), y1: P.y, x2: L.round(pos, 2), y2: P.y + P.h }) : s('line', { x1: P.x, y1: L.round(pos, 2), x2: P.x + P.w, y2: L.round(pos, 2) });
      grid.appendChild(applyStroke(l2, st));
    };
    if (vAx.minorGrid && sc.step && !sc.log) { const st = strokeOf(vAx.minorGrid, design, { c: '#D9D9D9', w: 0.5 }); const ms = sc.step / 5; for (let v = Math.ceil(sc.min / ms) * ms; v <= sc.max + ms * 1e-6; v += ms) gridLine(vm(v), st, false); }
    if (d3) {
      /* back wall, side wall and floor, with the value gridlines on the walls */
      const { dx, dy } = dep;
      const poly = (pts, part, fallbackLine) => { const el = s('polygon', { points: pts.map((q) => L.round(q[0], 2) + ',' + L.round(q[1], 2)).join(' ') }); applyFill(el, part && part.fill, design, defs); applyStroke(el, strokeOf(part && part.line, design, fallbackLine)); grid.appendChild(el); };
      const W3 = c.walls || {};
      const sideX = dx >= 0 ? P.x : P.x + P.w;
      if (horiz) {
        poly([[P.x + dx, P.y + dy], [P.x + P.w + dx, P.y + dy], [P.x + P.w + dx, P.y + P.h + dy], [P.x + dx, P.y + P.h + dy]], W3.back, null);
        poly([[P.x, P.y + P.h], [P.x + P.w, P.y + P.h], [P.x + P.w + dx, P.y + P.h + dy], [P.x + dx, P.y + P.h + dy]], W3.floor, null);
        poly([[P.x, P.y], [P.x + dx, P.y + dy], [P.x + dx, P.y + P.h + dy], [P.x, P.y + P.h]], W3.side, null);
        if (vAx.grid) { const st = strokeOf(vAx.grid, design, { c: AUTO_LINE, w: 0.75 }); for (const t of sc.ticks) { const x = vm(t); grid.appendChild(applyStroke(s('line', { x1: x + dx, y1: P.y + dy, x2: x + dx, y2: P.y + P.h + dy }), st)); grid.appendChild(applyStroke(s('line', { x1: x, y1: P.y + P.h, x2: x + dx, y2: P.y + P.h + dy }), st)); } }
      } else {
        poly([[P.x + dx, P.y + dy], [P.x + P.w + dx, P.y + dy], [P.x + P.w + dx, P.y + P.h + dy], [P.x + dx, P.y + P.h + dy]], W3.back, null);
        poly([[sideX, P.y], [sideX + dx, P.y + dy], [sideX + dx, P.y + P.h + dy], [sideX, P.y + P.h]], W3.side, null);
        poly([[P.x, P.y + P.h], [P.x + P.w, P.y + P.h], [P.x + P.w + dx, P.y + P.h + dy], [P.x + dx, P.y + P.h + dy]], W3.floor, null);
        if (vAx.grid) { const st = strokeOf(vAx.grid, design, { c: AUTO_LINE, w: 0.75 }); for (const t of sc.ticks) { const y = vm(t); grid.appendChild(applyStroke(s('line', { x1: P.x + dx, y1: y + dy, x2: P.x + P.w + dx, y2: y + dy }), st)); grid.appendChild(applyStroke(s('line', { x1: sideX, y1: y, x2: sideX + dx, y2: y + dy }), st)); } }
      }
    } else if (vAx.grid) { const st = strokeOf(vAx.grid, design, { c: AUTO_LINE, w: 0.75 }); for (const t of sc.ticks) gridLine(vm(t), st, false); }
    if (cAx.grid) {
      const st = strokeOf(cAx.grid, design, { c: AUTO_LINE, w: 0.75 });
      if (xy) for (const t of xsc.ticks) grid.appendChild(applyStroke(s('line', { x1: L.round(xm(t), 2), y1: P.y, x2: L.round(xm(t), 2), y2: P.y + P.h }), st));
      else for (let k = 0; k <= nCat; k++) { const p = between ? catEdge(k) : catPos(k, cAx.rev); if (!between && k === nCat) break; gridLine(p, st, true); }
    }
    svg.appendChild(grid);

    /* ----- series ----- */
    const plotG = s('g', { class: 'plot' });
    const front = s('g', { class: 'plot-front' });
    const labels = s('g', { class: 'labels' });
    const lblText = (sr, k, v, pctV, lb) => {
      const p = (lb.pts && lb.pts[k]) || null;
      if (p && p.del) return null;
      const o = p ? Object.assign({}, lb, p) : lb;
      if (o.text) return { text: fields(o.text, sr, k, v, pctV, lb), o };
      const parts = [];
      if (o.showSerName) parts.push(sr.name);
      if (o.showCatName) parts.push(xy ? fmtVal(sr.xs ? sr.xs[k] : k + 1, null) : catLabels[k] || '');
      if (o.showVal) parts.push(fmtVal(v, o.fmt || sr.fmt));
      if (o.showPercent && pctV != null) parts.push(fmtVal(pctV, o.fmt && /%/.test(o.fmt) ? o.fmt : '0%'));
      if (o.showBubbleSize && sr.sizes) parts.push(fmtVal(sr.sizes[k], null));
      if (o.showRange && lb.range && lb.range[k]) parts.unshift(lb.range[k]);
      if (!parts.length) return null;
      return { text: parts.join(o.sep != null ? o.sep : ', '), o };
    };
    function fields(t, sr, k, v, pctV, lb) {
      return String(t).replace(/\u0001(\w+)\u0001/g, (m, f) => {
        switch (f) {
          case 'VALUE': case 'YVALUE': return fmtVal(v, lb.fmt || sr.fmt);
          case 'SERIESNAME': return sr.name;
          case 'CATEGORYNAME': return xy ? fmtVal(sr.xs ? sr.xs[k] : k + 1, null) : catLabels[k] || '';
          case 'XVALUE': return fmtVal(sr.xs ? sr.xs[k] : k + 1, null);
          case 'PERCENTAGE': return pctV != null ? fmtVal(pctV, '0%') : '';
          case 'BUBBLESIZE': return sr.sizes ? fmtVal(sr.sizes[k], null) : '';
          case 'CELLRANGE': return (lb.range && lb.range[k]) || '';
          default: return '';
        }
      });
    }
    const putLabel = (info, x, y, anchor) => {
      const st = TS(info.o.tx);
      const lines = String(info.text).split('\n');
      if (info.o.fill || info.o.line) {
        const tw = Math.max(...lines.map((l) => measure(l, st))), th = lines.length * lineH(st);
        const bx = anchor === 'start' ? x - 2 : anchor === 'end' ? x - tw - 2 : x - tw / 2 - 2;
        const r = s('rect', { x: L.round(bx, 2), y: L.round(y - th / 2 - 1, 2), width: L.round(tw + 4, 2), height: L.round(th + 2, 2) });
        applyFill(r, info.o.fill, design, defs); applyStroke(r, strokeOf(info.o.line, design, null));
        labels.appendChild(r);
      }
      text(labels, lines, x, y, st, anchor || 'middle');
    };

    for (const g of groups) {
      const sis = (g.si || []).filter((i) => series[i]);
      if (!sis.length) continue;
      const vmap = g.sec ? vm2 : vm;
      const scl = g.sec && sc2 ? sc2 : sc;
      const pct = /percentStacked/.test(g.grouping);
      const stacked = /stacked/i.test(g.grouping);
      const totals = [];
      if (pct) for (let k = 0; k < nCat; k++) totals[k] = sis.reduce((a, i) => a + Math.abs(series[i].vals[k] || 0), 0) || 1;
      const valueAt = (i, k) => { const v = series[i].vals[k]; if (v == null) return null; return (pct ? v / totals[k] : v) / (pct ? 1 : unit(g.sec ? v2Ax : vAx)); };
      if (g.type === 'bar') {
        const deep = d3 && g.deep;
        const n = stacked || deep ? 1 : sis.length;
        const ov = stacked ? 1 : Math.max(-1, Math.min(1, (g.overlap || 0) / 100));
        const gap = Math.max(0, (g.gap != null ? g.gap : 150) / 100);
        const bandW = catLen / Math.max(nCat, 1);
        const bw = bandW / (n - (n - 1) * ov + gap);
        const posAcc = new Array(nCat).fill(0), negAcc = new Array(nCat).fill(0);
        const zero = clampV(scl, scl.log ? scl.min : 0);
        /* rows in depth are drawn from the back */
        (deep ? sis.map((i, j) => [i, j]).reverse() : sis.map((i, j) => [i, j])).forEach(([i, j]) => {
          const sr = series[i];
          const stroke = serStroke(sr, g);
          const zf = deep ? j / sis.length : 0, zd = deep ? 0.75 / sis.length : 1;
          for (let k = 0; k < nCat; k++) {
            const v = valueAt(i, k);
            if (v == null) continue;
            let a = zero, b = v;
            if (stacked) { if (v >= 0) { a = posAcc[k]; b = a + v; posAcc[k] = b; } else { a = negAcc[k]; b = a + v; negAcc[k] = b; } }
            const p1 = vmap(clampV(scl, a)), p2 = vmap(clampV(scl, b));
            const off = gap * bw / 2 + (stacked || deep ? 0 : j * bw * (1 - ov));
            const e0 = catEdge(k), e1 = catEdge(k + 1);
            const lo = Math.min(e0, e1);
            const pt = sr.pts && sr.pts[k];
            let fill = (pt && pt.fill) || (g.vary && sis.length === 1 && !sr.fill && CH.autoColor ? solid(CH.autoColor(k, nCat, c.style, design)) : serFill(sr));
            if (sr.invNeg && v < 0 && fill.t === 'solid' && !(pt && pt.fill)) fill = { t: 'solid', c: '#FFFFFF', a: 1 };
            /* bars of horizontal charts run bottom-up within a category */
            const cPos = horiz ? (cAx.rev ? lo + off : lo + bandW - off - bw) : (cAx.rev ? lo + bandW - off - bw : lo + off);
            const rect = horiz
              ? s('rect', { x: L.round(Math.min(p1, p2), 2), y: L.round(cPos, 2), width: L.round(Math.abs(p2 - p1), 2), height: L.round(bw, 2) })
              : s('rect', { x: L.round(cPos, 2), y: L.round(Math.min(p1, p2), 2), width: L.round(bw, 2), height: L.round(Math.abs(p2 - p1), 2) });
            applyFill(rect, fill, design, defs);
            applyStroke(rect, pt && pt.line ? strokeOf(pt.line, design, null) : stroke);
            if (d3) {
              /* a box: the front face, then the top and the side in lighter and darker shades */
              const bx = +rect.getAttribute('x') + dep.dx * zf, by = +rect.getAttribute('y') + dep.dy * zf, bW = +rect.getAttribute('width'), bH = +rect.getAttribute('height');
              if (zf) { rect.setAttribute('x', L.round(bx, 2)); rect.setAttribute('y', L.round(by, 2)); }
              const dx = dep.dx * zd, dy = dep.dy * zd;
              const base = R().col(fill.t === 'solid' ? fill.c : fill.t === 'grad' && fill.stops && fill.stops[0] ? fill.stops[0].c : fill.t === 'patt' ? fill.fg : sr.color || '#4472C4', design, '#4472C4');
              const fa = fill.a == null ? 1 : fill.a;
              const face = (pts, col) => { const el = s('polygon', { points: pts.map((q) => L.round(q[0], 2) + ',' + L.round(q[1], 2)).join(' '), fill: col }); if (fa < 1) el.setAttribute('fill-opacity', L.round(fa, 3)); applyStroke(el, pt && pt.line ? strokeOf(pt.line, design, null) : stroke); plotG.appendChild(el); };
              plotG.appendChild(rect);
              face([[bx, by], [bx + bW, by], [bx + bW + dx, by + dy], [bx + dx, by + dy]], L.color.lighten(base, 0.25));
              const sx = dx >= 0 ? bx + bW : bx;
              face([[sx, by], [sx + dx, by + dy], [sx + dx, by + bH + dy], [sx, by + bH]], L.color.darken(base, 0.25));
            } else plotG.appendChild(rect);
            if (sr.lbl) {
              const info = lblText(sr, k, series[i].vals[k], pct ? v : null, sr.lbl);
              if (info) {
                const posn = info.o.pos || (stacked ? 'ctr' : 'outEnd');
                const st = TS(info.o.tx);
                const neg = (b < a) !== !!vAx.rev;
                const far = Math.max(p1, p2), near = Math.min(p1, p2);
                const mid = cPos + bw / 2;
                if (horiz) {
                  const endX = neg ? near : far, baseX = neg ? far : near;
                  const tw0 = measure(info.text.split('\n')[0], st);
                  const x = posn === 'ctr' ? (p1 + p2) / 2 : posn === 'inEnd' ? endX + (neg ? 1 : -1) * (tw0 / 2 + 3) : posn === 'inBase' ? baseX + (neg ? -1 : 1) * (tw0 / 2 + 3) : endX + (neg ? -1 : 1) * (tw0 / 2 + 3);
                  putLabel(info, x, mid, 'middle');
                } else {
                  const endY = neg ? far : near, baseY = neg ? near : far;
                  const hh2 = lineH(st) * info.text.split('\n').length / 2 + 2;
                  const y = posn === 'ctr' ? (p1 + p2) / 2 : posn === 'inEnd' ? endY + (neg ? -1 : 1) * hh2 : posn === 'inBase' ? baseY + (neg ? 1 : -1) * hh2 : endY + (neg ? 1 : -1) * hh2;
                  putLabel(info, mid, y, 'middle');
                }
              }
            }
          }
        });
      } else if (g.type === 'area') {
        const acc = new Array(nCat).fill(0);
        const zero = clampV(scl, 0);
        /* overlapping (standard) areas: the first series is drawn in front, as Office does */
        (stacked ? sis : sis.slice().reverse()).forEach((i) => {
          const sr = series[i];
          const topV = [], baseV = [];
          for (let k = 0; k < nCat; k++) {
            const v = valueAt(i, k) || 0;
            baseV.push(stacked ? acc[k] : zero);
            topV.push(stacked ? acc[k] + v : v);
            if (stacked) acc[k] += v;
          }
          const pts = topV.map((v, k) => (horiz ? [vmap(clampV(scl, v)), catPos(k, cAx.rev)] : [catPos(k, cAx.rev), vmap(clampV(scl, v))]));
          const bpts = baseV.map((v, k) => (horiz ? [vmap(clampV(scl, v)), catPos(k, cAx.rev)] : [catPos(k, cAx.rev), vmap(clampV(scl, v))])).reverse();
          const poly = s('polygon', { points: pts.concat(bpts).map((p) => L.round(p[0], 2) + ',' + L.round(p[1], 2)).join(' ') });
          applyFill(poly, serFill(sr), design, defs); applyStroke(poly, serStroke(sr, g));
          plotG.appendChild(poly);
          if (sr.lbl) pts.forEach((p, k) => { const info = lblText(sr, k, sr.vals[k], pct ? valueAt(i, k) : null, sr.lbl); if (info) putLabel(info, p[0], (p[1] + (bpts[nCat - 1 - k] || p)[1]) / 2, 'middle'); });
        });
      } else if (g.type === 'line' || g.type === 'stock') {
        const acc = new Array(nCat).fill(0);
        const allPts = [];
        sis.forEach((i, j) => {
          const sr = series[i];
          const pts = [];
          for (let k = 0; k < nCat; k++) {
            let v = valueAt(i, k);
            if (v == null) { pts.push(c.blanks === 'zero' ? null : null); continue; }
            if (stacked) { acc[k] += v; v = acc[k]; }
            pts.push(horiz ? [vmap(clampV(scl, v)), catPos(k, cAx.rev)] : [catPos(k, cAx.rev), vmap(clampV(scl, v))]);
          }
          allPts.push(pts);
          const st = g.type === 'stock' && !sr.line ? null : serStroke(sr, g);
          if (st) plotG.appendChild(applyStroke(s('path', { d: pathOf(c.blanks === 'span' ? pts.filter(Boolean) : pts, sr.smooth), fill: 'none', 'stroke-linejoin': 'round' }), st));
          const sym = markerSym(sr, g, j);
          if (sym !== 'none') pts.forEach((p, k) => { if (!p) return; const pt = sr.pts && sr.pts[k]; const mp = markerPaint(sr, pt); marker(front, (pt && pt.marker && pt.marker.sym) || sym, p[0], p[1], markerSize(sr), mp.fill, mp.stroke, design, defs); });
          if (sr.lbl) pts.forEach((p, k) => {
            if (!p) return;
            const info = lblText(sr, k, sr.vals[k], pct ? valueAt(i, k) : null, sr.lbl);
            if (!info) return;
            const st = TS(info.o.tx), d = markerSize(sr) / 2 + 3;
            const posn = info.o.pos || 'r';
            if (posn === 't') putLabel(info, p[0], p[1] - d - st.sz * 0.5, 'middle');
            else if (posn === 'b') putLabel(info, p[0], p[1] + d + st.sz * 0.5, 'middle');
            else if (posn === 'l') putLabel(info, p[0] - d, p[1], 'end');
            else if (posn === 'ctr') putLabel(info, p[0], p[1], 'middle');
            else putLabel(info, p[0] + d, p[1], 'start');
          });
        });
        if (g.hiLow) {
          const st = strokeOf(g.hiLow, design, { c: '#000000', w: 0.75 });
          for (let k = 0; k < nCat; k++) {
            const ys = allPts.map((pts) => pts[k]).filter(Boolean);
            if (ys.length < 2) continue;
            const a = Math.min(...ys.map((p) => p[1])), b = Math.max(...ys.map((p) => p[1]));
            plotG.appendChild(applyStroke(s('line', { x1: ys[0][0], y1: a, x2: ys[0][0], y2: b }), st));
          }
        }
        if (g.drop) {
          const st = strokeOf(g.drop, design, { c: '#000000', w: 0.75 });
          allPts.forEach((pts) => pts.forEach((p) => { if (p) plotG.appendChild(applyStroke(s('line', { x1: p[0], y1: p[1], x2: p[0], y2: P.y + P.h }), st)); }));
        }
      } else if (g.type === 'scatter' || g.type === 'bubble') {
        const sizes = g.type === 'bubble' ? series.filter((sr) => sr.sizes).flatMap((sr) => sr.sizes.filter((x) => x != null && x > 0)) : [];
        const maxSize = Math.max(1e-12, ...sizes);
        const dMax = Math.min(P.w, P.h) * 0.25 * ((g.bubbleScale || 100) / 100);
        sis.forEach((i, j) => {
          const sr = series[i];
          const xs = sr.xs || sr.vals.map((_, k) => k + 1);
          /* an XY series in a category chart (combination charts) sits on the categories */
          const xpos = xy ? (x) => xm(x) : (x, k) => catPos(k, cAx.rev);
          const pts = sr.vals.map((v, k) => {
            if (v == null || xs[k] == null) return null;
            const a = xpos(xs[k], k), b = vmap(clampV(scl, v / unit(g.sec ? v2Ax : vAx)));
            return horiz && !xy ? [b, a] : [a, b];
          });
          if (g.type === 'bubble') {
            const order = pts.map((p, k) => k).filter((k) => pts[k]).sort((a, b) => (sr.sizes[b] || 0) - (sr.sizes[a] || 0));
            for (const k of order) {
              const z = Math.max(0, sr.sizes[k] || 0);
              const d = g.sizeArea === false ? dMax * (z / maxSize) : dMax * Math.sqrt(z / maxSize);
              const pt = sr.pts && sr.pts[k];
              const ci = s('circle', { cx: L.round(pts[k][0], 2), cy: L.round(pts[k][1], 2), r: L.round(Math.max(d / 2, 0.5), 2) });
              applyFill(ci, (pt && pt.fill) || (g.vary && series.length === 1 && CH.autoColor ? solid(CH.autoColor(k, sr.vals.length, c.style, design)) : serFill(sr)), design, defs); applyStroke(ci, (pt && pt.line && strokeOf(pt.line, design, null)) || strokeOf(sr.line, design, null));
              plotG.appendChild(ci);
            }
          } else {
            const drawLine = sr.line ? sr.line.t !== 'none' : g.scatterStyle !== 'marker' && g.scatterStyle !== 'none';
            const smooth = sr.smooth != null ? sr.smooth : /smooth/i.test(g.scatterStyle || '');
            if (drawLine) { const st = serStroke(sr, g); if (st) plotG.appendChild(applyStroke(s('path', { d: pathOf(pts, smooth), fill: 'none', 'stroke-linejoin': 'round' }), st)); }
            const sym = markerSym(sr, g, j);
            if (sym !== 'none') pts.forEach((p, k) => { if (!p) return; const pt = sr.pts && sr.pts[k]; const mp = markerPaint(sr, pt); marker(front, (pt && pt.marker && pt.marker.sym) || sym, p[0], p[1], markerSize(sr), mp.fill, mp.stroke, design, defs); });
          }
          if (sr.lbl) pts.forEach((p, k) => {
            if (!p) return;
            const info = lblText(sr, k, sr.vals[k], null, sr.lbl);
            if (!info) return;
            const posn = info.o.pos || (g.type === 'bubble' ? 'ctr' : 'r');
            const d = markerSize(sr) / 2 + 3, st = TS(info.o.tx);
            if (posn === 'ctr') putLabel(info, p[0], p[1], 'middle');
            else if (posn === 't') putLabel(info, p[0], p[1] - d - st.sz * 0.5, 'middle');
            else if (posn === 'b') putLabel(info, p[0], p[1] + d + st.sz * 0.5, 'middle');
            else if (posn === 'l') putLabel(info, p[0] - d, p[1], 'end');
            else putLabel(info, p[0] + d, p[1], 'start');
          });
        });
      }
    }
    svg.appendChild(plotG);

    /* ----- axes ----- */
    const axG = s('g', { class: 'axes' });
    const tickLen = Math.max(3, vSt.sz * 0.3);
    const drawValAxis = (a, scl, st, side, mapF, sec) => {
      if (!a || a.del) return;
      const lineSt = strokeOf(a.line, design, { c: AUTO_LINE, w: 0.75 });
      /* where the value axis crosses the category axis */
      const at = horiz ? (side === 't' ? P.y : P.y + P.h) : side === 'r' ? P.x + P.w : P.x;
      if (lineSt) axG.appendChild(applyStroke(horiz ? s('line', { x1: P.x, y1: at, x2: P.x + P.w, y2: at }) : s('line', { x1: at, y1: P.y, x2: at, y2: P.y + P.h }), lineSt));
      const out = side === 'l' || side === 't' ? -1 : 1;
      const fmt = valFmt(a, sec);
      for (const t of scl.ticks) {
        const p = mapF(t);
        if (lineSt && a.tick !== 'none') {
          const t0 = a.tick === 'in' ? 0 : a.tick === 'cross' ? -tickLen / 2 * out : 0, t1 = a.tick === 'in' ? -tickLen * out : a.tick === 'cross' ? tickLen / 2 * out : tickLen * out;
          axG.appendChild(applyStroke(horiz ? s('line', { x1: p, y1: at + t0, x2: p, y2: at + t1 }) : s('line', { x1: at + t0, y1: p, x2: at + t1, y2: p }), lineSt));
        }
        if (a.lblPos === 'none') continue;
        const label = fmtVal(t, fmt);
        if (horiz) text(axG, [label], p, side === 't' ? at - st.sz * 0.3 - lineH(st) / 2 : at + st.sz * 0.3 + lineH(st) / 2, st, 'middle');
        else text(axG, [label], side === 'r' ? at + st.sz * 0.4 : at - st.sz * 0.4, p, st, side === 'r' ? 'start' : 'end');
      }
      if (a.title) {
        const tst = TS(Object.assign({ b: true }, a.titleTx));
        const lw = Math.max(...scl.ticks.map((t) => measure(fmtVal(t, fmt), st)));
        if (horiz) text(axG, a.title, P.x + P.w / 2, side === 't' ? at - lineH(st) - st.sz * 0.6 - lineH(tst) / 2 : at + lineH(st) + st.sz * 0.6 + lineH(tst) / 2, tst, 'middle');
        else {
          const rot = tst.rot != null ? tst.rot : side === 'r' ? 90 : -90;
          const x = side === 'r' ? at + lw + st.sz * 0.6 + lineH(tst) / 2 : at - lw - st.sz * 0.6 - lineH(tst) / 2;
          text(axG, a.title.split('\n'), x, P.y + P.h / 2, tst, 'middle', rot);
        }
      }
    };
    drawValAxis(vAx, sc, vSt, vSide, vm, false);
    if (v2Ax) drawValAxis(v2Ax, sc2, v2St, horiz ? (vSide === 't' ? 'b' : 't') : vSide === 'r' ? 'l' : 'r', vm2, true);
    /* category / x axis */
    if (!cAx.del) {
      const lineSt = strokeOf(cAx.line, design, { c: AUTO_LINE, w: 0.75 });
      const cross = cAx.crosses === 'max' ? sc.max : cAx.crosses === 'min' ? sc.min : cAx.crosses === 'at' ? clampV(sc, cAx.crossesAt) : clampV(sc, sc.log ? sc.min : 0);
      const at = xy ? vm(cross) : vm(cross);
      const lowAt = horiz ? (cSide === 'r' ? P.x + P.w : P.x) : cSide === 't' ? P.y : P.y + P.h;
      const lblAt = cAx.lblPos === 'low' ? lowAt : cAx.lblPos === 'high' ? (horiz ? (cSide === 'r' ? P.x : P.x + P.w) : cSide === 't' ? P.y + P.h : P.y) : (xy || !horiz ? (cSide === 't' ? Math.min(at, P.y + P.h) : Math.max(at, P.y)) : at);
      const labelSide = horiz ? (cSide === 'r' ? 1 : -1) : cSide === 't' ? -1 : 1;
      if (lineSt) axG.appendChild(applyStroke(horiz ? s('line', { x1: at, y1: P.y, x2: at, y2: P.y + P.h }) : s('line', { x1: P.x, y1: at, x2: P.x + P.w, y2: at }), lineSt));
      if (xy) {
        const fmt = cAx.fmt && !cAx.linked && cAx.fmt !== 'General' ? cAx.fmt : null;
        for (const t of xsc.ticks) {
          const p = xm(t);
          if (lineSt && cAx.tick !== 'none') axG.appendChild(applyStroke(s('line', { x1: p, y1: at, x2: p, y2: at + (cAx.tick === 'in' ? -tickLen : tickLen) * labelSide }), lineSt));
          if (cAx.lblPos !== 'none') text(axG, [fmtVal(t, fmt)], p, (cAx.lblPos === 'low' ? P.y + P.h : at) + (cSt.sz * 0.3 + lineH(cSt) / 2) * labelSide, cSt, 'middle');
        }
      } else {
        if (lineSt && cAx.tick !== 'none') for (let k = 0; k <= nCat; k++) {
          const p = between ? catEdge(k) : catPos(k, cAx.rev);
          if (!between && k === nCat) break;
          const t1 = (cAx.tick === 'in' ? -tickLen : tickLen) * labelSide, t0 = cAx.tick === 'cross' ? -t1 : 0;
          axG.appendChild(applyStroke(horiz ? s('line', { x1: at + t0, y1: p, x2: at + t1, y2: p }) : s('line', { x1: p, y1: at + t0, x2: p, y2: at + t1 }), lineSt));
        }
        if (cAx.lblPos !== 'none') {
          for (let k = 0; k < nCat; k += catMode.skip) {
            const p = catPos(k, cAx.rev);
            const lines = catMode.lines[k] || [''];
            if (horiz) text(axG, lines, lblAt + catGap * labelSide, p, cSt, labelSide < 0 ? 'end' : 'start');
            else if (catMode.rot) {
              const y = lblAt + catGap * labelSide;
              const t = text(axG, lines, p, y + cSt.sz * 0.35 * labelSide, cSt, catMode.rot < 0 ? 'end' : 'start', catMode.rot);
              void t;
            } else text(axG, lines, p, lblAt + (catGap + lines.length * lineH(cSt) / 2) * labelSide, cSt, 'middle');
          }
          /* outer levels of a multi-level category axis */
          if (outerLevels && !horiz && !catMode.rot) {
            let y = lblAt + (catGap + Math.max(...catMode.lines.map((ls) => ls.length)) * lineH(cSt)) * labelSide;
            for (const lvl of outerLevels) {
              let k0 = 0;
              for (let k = 1; k <= nCat; k++) {
                if (k < nCat && (lvl[k] == null || lvl[k] === '')) continue;
                const a = catEdge(k0), b = catEdge(k);
                if (lvl[k0]) text(axG, [lvl[k0]], (a + b) / 2, y + (catGap + lineH(cSt) / 2) * labelSide, cSt, 'middle');
                if (lineSt) axG.appendChild(applyStroke(s('line', { x1: a, y1: y, x2: a, y2: y + (lineH(cSt) + catGap) * labelSide }), lineSt));
                k0 = k;
              }
              if (lineSt) axG.appendChild(applyStroke(s('line', { x1: catEdge(nCat), y1: y, x2: catEdge(nCat), y2: y + (lineH(cSt) + catGap) * labelSide }), lineSt));
              y += (lineH(cSt) + catGap) * labelSide;
            }
          }
        }
      }
      if (cAx.title) {
        const tst = TS(Object.assign({ b: true }, cAx.titleTx));
        if (horiz) {
          const rot = tst.rot != null ? tst.rot : -90;
          text(axG, cAx.title.split('\n'), P.x - catMode.size - catGap * 2 - lineH(tst) / 2, P.y + P.h / 2, tst, 'middle', rot);
        } else {
          const lblH = xy ? lineH(cSt) + catGap : catMode.size + catGap + (outerLevels ? outerLevels.length * (lineH(cSt) + catGap) : 0);
          text(axG, cAx.title.split('\n'), P.x + P.w / 2, cSide === 't' ? P.y - lblH - lineH(tst) / 2 - 2 : P.y + P.h + lblH + lineH(tst) / 2 + 2, tst, 'middle');
        }
      }
    }
    svg.appendChild(axG);
    svg.appendChild(front);
    svg.appendChild(labels);
    return svg;
  };

  /* ---------- 3-D pie: a tilted disc with its rim ---------- */
  function drawPie3D(svg, defs, c, P, design, TS, series, groups) {
    const g0 = groups[0];
    const si = (g0.si || []).find((i) => series[i]);
    if (si == null) return;
    const sr = series[si];
    const vals = sr.vals.map((v) => (v == null ? 0 : Math.abs(v)));
    const tot = vals.reduce((a, b) => a + b, 0);
    if (!(tot > 0)) return;
    const tilt = Math.max(0.2, Math.min(1, Math.sin(((c.view3D && c.view3D.rotX) || 30) * Math.PI / 180)));
    const maxExpl = Math.max(0, sr.expl || 0, ...Object.values(sr.pts || {}).map((p) => p.expl || 0));
    const anyLbl = !!sr.lbl;
    let rx = Math.min(P.w / 2, (P.h / (2 * tilt + 0.3))) * (anyLbl ? 0.8 : 0.95) / (1 + maxExpl / 100);
    const ry = rx * tilt, th = rx * 0.25 * Math.max(0.3, ((c.view3D && c.view3D.depth) || 100) / 100);
    const cx = P.x + P.w / 2, cy = P.y + (P.h - th) / 2;
    const slices = [];
    let a0 = ((g0.firstAng || 0) - 90) * Math.PI / 180;
    vals.forEach((v, k) => {
      const frac = v / tot, a1 = a0 + frac * Math.PI * 2;
      const pt = sr.pts && sr.pts[k];
      const fill = (pt && pt.fill) || (g0.vary ? solid((c.pieColors && c.pieColors[k]) || (CH.autoColor ? CH.autoColor(k, vals.length, c.style, design) : sr.color)) : sr.fill || solid(sr.color));
      const ex = ((pt && pt.expl != null ? pt.expl : sr.expl) || 0) / 100 * rx;
      const am = (a0 + a1) / 2;
      slices.push({ k, frac, a0, a1, am, fill, line: (pt && pt.line) || sr.line, ox: Math.cos(am) * ex, oy: Math.sin(am) * ex * tilt });
      a0 = a1;
    });
    const E = (sl, a, dy) => [cx + sl.ox + rx * Math.cos(a), cy + sl.oy + ry * Math.sin(a) + (dy || 0)];
    const f2 = (q) => L.round(q[0], 2) + ',' + L.round(q[1], 2);
    const colOf = (f) => R().col(f.t === 'solid' ? f.c : f.t === 'grad' && f.stops && f.stops[0] ? f.stops[0].c : f.t === 'patt' ? f.fg : '#4472C4', design, '#4472C4');
    const g = s('g', { class: 'plot' });
    /* rims: only the front half of the disc shows its edge; farther pieces first */
    const rims = [];
    for (const sl of slices) {
      if (sl.frac <= 0) continue;
      for (let turn = -1; turn <= 1; turn++) {
        const lo = Math.max(sl.a0, turn * 2 * Math.PI), hi = Math.min(sl.a1, turn * 2 * Math.PI + Math.PI);
        if (hi > lo) rims.push({ sl, lo, hi });
      }
    }
    rims.sort((p, q) => Math.sin((q.lo + q.hi) / 2) - Math.sin((p.lo + p.hi) / 2)).reverse();
    for (const r of rims) {
      const sl = r.sl, large = r.hi - r.lo > Math.PI ? 1 : 0;
      const d = `M${f2(E(sl, r.lo))}A${L.round(rx, 2)},${L.round(ry, 2)} 0 ${large} 1 ${f2(E(sl, r.hi))}L${f2(E(sl, r.hi, th))}A${L.round(rx, 2)},${L.round(ry, 2)} 0 ${large} 0 ${f2(E(sl, r.lo, th))}Z`;
      const el = s('path', { d, fill: L.color.darken(colOf(sl.fill), 0.3) });
      applyStroke(el, strokeOf(sl.line, design, null));
      g.appendChild(el);
    }
    /* the cut faces at the slice edges that face the viewer */
    if (maxExpl > 0) for (const sl of slices) {
      if (sl.frac <= 0 || sl.frac >= 0.9999) continue;
      for (const a of [sl.a0, sl.a1]) {
        /* a cut face shows when it turns towards the viewer: the start edge on the left, the end edge on the right */
        if (a === sl.a0 ? Math.cos(a) >= 0 : Math.cos(a) <= 0) continue;
        const ctr = [cx + sl.ox, cy + sl.oy];
        const el = s('path', { d: `M${f2(ctr)}L${f2(E(sl, a))}L${f2(E(sl, a, th))}L${f2([ctr[0], ctr[1] + th])}Z`, fill: L.color.darken(colOf(sl.fill), 0.2) });
        g.appendChild(el);
      }
    }
    for (const sl of slices) {
      if (sl.frac <= 0) continue;
      const ctr = [cx + sl.ox, cy + sl.oy];
      const d = sl.frac >= 0.9999 ? `M${f2(E(sl, 0))}A${rx},${ry} 0 1 1 ${f2(E(sl, Math.PI))}A${rx},${ry} 0 1 1 ${f2(E(sl, 0))}Z` : `M${f2(ctr)}L${f2(E(sl, sl.a0))}A${L.round(rx, 2)},${L.round(ry, 2)} 0 ${sl.a1 - sl.a0 > Math.PI ? 1 : 0} 1 ${f2(E(sl, sl.a1))}Z`;
      const el = s('path', { d });
      applyFill(el, sl.fill, design, defs); applyStroke(el, strokeOf(sl.line, design, null));
      g.appendChild(el);
    }
    svg.appendChild(g);
    if (sr.lbl) {
      const labels = s('g', { class: 'labels' });
      for (const sl of slices) {
        const lb = sr.lbl, p = lb.pts && lb.pts[sl.k];
        if (p && p.del) continue;
        const o = p ? Object.assign({}, lb, p) : lb;
        const parts = [];
        if (o.showSerName) parts.push(sr.name);
        if (o.showCatName) parts.push((c.cats || [])[sl.k] || '');
        if (o.showVal) parts.push(fmtVal(sr.vals[sl.k], o.fmt || sr.fmt));
        if (o.showPercent) parts.push(fmtVal(sl.frac, o.fmt && /%/.test(o.fmt) ? o.fmt : '0%'));
        const txt = o.text ? String(o.text).replace(/\u0001\w+\u0001/g, '') : parts.join(o.sep != null ? o.sep : ', ');
        if (!txt) continue;
        const st = TS(o.tx), out = (o.pos || 'bestFit') === 'outEnd';
        const k = out ? 1.15 : 0.62, q = E(sl, sl.am);
        const x = cx + sl.ox + (q[0] - cx - sl.ox) * k, y = cy + sl.oy + (q[1] - cy - sl.oy) * k;
        text(labels, String(txt).split('\n'), x, y, st, out ? (Math.cos(sl.am) > 0.15 ? 'start' : Math.cos(sl.am) < -0.15 ? 'end' : 'middle') : 'middle');
      }
      svg.appendChild(labels);
    }
  }

  /* ---------- pie & doughnut ---------- */
  function drawPie(svg, defs, c, P, design, TS, series, groups) {
    const g0 = groups[0];
    if (g0.d3 && g0.type === 'pie') return drawPie3D(svg, defs, c, P, design, TS, series, groups);
    const sis = (g0.si || []).filter((i) => series[i]);
    if (!sis.length) return;
    const doughnut = g0.type === 'doughnut';
    const rings = doughnut ? sis : sis.slice(0, 1);
    const cx = P.x + P.w / 2, cy = P.y + P.h / 2;
    /* labels outside the pie need room */
    const anyOut = rings.some((i) => series[i].lbl && /outEnd|bestFit/.test(series[i].lbl.pos || 'bestFit') && !doughnut);
    const outerLbl = rings.some((i) => series[i].lbl && (series[i].lbl.pos === 'outEnd'));
    let R0 = Math.max(4, Math.min(P.w, P.h) / 2);
    if (outerLbl) R0 *= 0.78; else if (anyOut) R0 *= 0.95;
    const maxExpl = Math.max(0, ...rings.map((i) => { const sr = series[i]; return Math.max(sr.expl || 0, ...Object.values(sr.pts || {}).map((p) => p.expl || 0)); }));
    const r = R0 / (1 + maxExpl / 100);
    const hole = doughnut ? Math.max(0.1, Math.min(0.9, (g0.hole || 50) / 100)) : 0;
    const ringW = (r - r * hole) / rings.length;
    const plotG = s('g', { class: 'plot' }), labels = s('g', { class: 'labels' });
    rings.forEach((si, ri) => {
      const sr = series[si];
      const vals = sr.vals.map((v) => (v == null ? 0 : Math.abs(v)));
      const tot = vals.reduce((a, b) => a + b, 0);
      if (!(tot > 0)) return;
      const rOut = r - ri * ringW, rIn = doughnut ? rOut - ringW : 0;
      let a0 = ((g0.firstAng || 0) - 90) * Math.PI / 180;
      vals.forEach((v, k) => {
        const frac = v / tot;
        const a1 = a0 + frac * Math.PI * 2;
        if (frac <= 0) { a0 = a1; return; }
        const pt = sr.pts && sr.pts[k];
        const fill = (pt && pt.fill) || (g0.vary ? solid((c.pieColors && c.pieColors[k]) || (CH.autoColor ? CH.autoColor(k, vals.length, c.style, design) : sr.color)) : sr.fill || solid(sr.color));
        const ex = ((pt && pt.expl != null ? pt.expl : sr.expl) || 0) / 100 * r;
        const am = (a0 + a1) / 2;
        const ox = cx + Math.cos(am) * ex, oy = cy + Math.sin(am) * ex;
        let d;
        const P2 = (rad, a) => `${L.round(ox + rad * Math.cos(a), 2)},${L.round(oy + rad * Math.sin(a), 2)}`;
        const large = a1 - a0 > Math.PI ? 1 : 0;
        if (frac >= 0.99999) {
          d = `M${P2(rOut, a0)}A${rOut},${rOut} 0 1 1 ${P2(rOut, a0 + Math.PI)}A${rOut},${rOut} 0 1 1 ${P2(rOut, a0)}Z`;
          if (rIn > 0) d += `M${P2(rIn, a0)}A${rIn},${rIn} 0 1 0 ${P2(rIn, a0 + Math.PI)}A${rIn},${rIn} 0 1 0 ${P2(rIn, a0)}Z`;
        } else if (rIn > 0) d = `M${P2(rOut, a0)}A${rOut},${rOut} 0 ${large} 1 ${P2(rOut, a1)}L${P2(rIn, a1)}A${rIn},${rIn} 0 ${large} 0 ${P2(rIn, a0)}Z`;
        else d = `M${L.round(ox, 2)},${L.round(oy, 2)}L${P2(rOut, a0)}A${rOut},${rOut} 0 ${large} 1 ${P2(rOut, a1)}Z`;
        const el = s('path', { d, 'fill-rule': 'evenodd' });
        applyFill(el, fill, design, defs);
        applyStroke(el, strokeOf((pt && pt.line) || sr.line, design, null));
        plotG.appendChild(el);
        if (sr.lbl) {
          const lb = sr.lbl;
          const p = lb.pts && lb.pts[k];
          if (!(p && p.del)) {
            const o = p ? Object.assign({}, lb, p) : lb;
            const parts = [];
            if (o.showSerName) parts.push(sr.name);
            if (o.showCatName) parts.push((c.cats || [])[k] || '');
            if (o.showVal) parts.push(fmtVal(sr.vals[k], o.fmt || sr.fmt));
            if (o.showPercent) parts.push(fmtVal(frac, o.fmt && /%/.test(o.fmt) ? o.fmt : '0%'));
            if (o.showRange && lb.range && lb.range[k]) parts.unshift(lb.range[k]);
            const F = { VALUE: fmtVal(sr.vals[k], o.fmt || sr.fmt), SERIESNAME: sr.name, CATEGORYNAME: (c.cats || [])[k] || '', PERCENTAGE: fmtVal(frac, '0%'), CELLRANGE: (lb.range && lb.range[k]) || '' };
            const txt = o.text ? o.text.replace(/\u0001(\w+)\u0001/g, (m, f) => F[f] || '') : parts.join(o.sep != null ? o.sep : ', ');
            if (txt) {
              const st = TS(o.tx);
              const pos = o.pos || 'bestFit';
              let rad, anchor = 'middle';
              const lines = String(txt).split('\n');
              const tw = Math.max(...lines.map((l) => measure(l, st)));
              if (doughnut) rad = (rOut + rIn) / 2;
              else if (pos === 'ctr') rad = rOut * 0.5;
              else if (pos === 'inEnd') rad = rOut * 0.75;
              else if (pos === 'outEnd') { rad = rOut + st.sz * 0.6; anchor = Math.cos(am) > 0.15 ? 'start' : Math.cos(am) < -0.15 ? 'end' : 'middle'; }
              else {
                /* best fit: inside when the label fits the slice, else just outside */
                const arcLen = frac * Math.PI * 2 * rOut * 0.65;
                if (tw < arcLen * 1.1 && tw < rOut * 0.9) rad = rOut * 0.65; else { rad = rOut + st.sz * 0.5; anchor = Math.cos(am) > 0.15 ? 'start' : Math.cos(am) < -0.15 ? 'end' : 'middle'; }
              }
              const x = ox + rad * Math.cos(am), y = oy + rad * Math.sin(am) + (anchor === 'middle' && rad > rOut ? Math.sign(Math.sin(am)) * lineH(st) * lines.length / 2 : 0);
              if (o.fill || o.line) {
                const th = lines.length * lineH(st);
                const bx = anchor === 'start' ? x - 2 : anchor === 'end' ? x - tw - 2 : x - tw / 2 - 2;
                const rr = s('rect', { x: bx, y: y - th / 2 - 1, width: tw + 4, height: th + 2 });
                applyFill(rr, o.fill, design, defs); applyStroke(rr, strokeOf(o.line, design, null));
                labels.appendChild(rr);
              }
              text(labels, lines, x, y, st, anchor);
            }
          }
        }
        a0 = a1;
      });
    });
    svg.appendChild(plotG);
    svg.appendChild(labels);
  }

  /* ---------- radar ---------- */
  function drawRadar(svg, defs, c, P, design, TS, series, groups, cats) {
    const g0 = groups[0];
    const sis = (g0.si || []).filter((i) => series[i]);
    const n = Math.max(cats.length, ...sis.map((i) => series[i].vals.length));
    if (n < 1) return;
    const ax = c.ax || {};
    const vAx = ax.v || {}, cAx = ax.c || {};
    const cSt = TS(cAx.tx), vSt = TS(vAx.tx);
    const maxLbl = Math.max(0, ...cats.map((t) => measure(String(t), cSt)));
    const r = Math.max(4, Math.min(P.w / 2 - maxLbl - 6, P.h / 2 - lineH(cSt) - 4));
    const cx = P.x + P.w / 2, cy = P.y + P.h / 2;
    let lo = Infinity, hi = -Infinity;
    for (const i of sis) for (const v of series[i].vals) if (v != null) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
    const sc = scale(lo, hi, vAx, r, vSt.sz);
    const ang = (k) => -Math.PI / 2 + (k / n) * Math.PI * 2;
    const rad = (v) => ((clamp(v, sc.min, sc.max) - sc.min) / (sc.max - sc.min || 1)) * r;
    const clamp = (v, a, b) => Math.max(a, Math.min(b, v));
    const pt = (k, v) => [cx + rad(v) * Math.cos(ang(k)), cy + rad(v) * Math.sin(ang(k))];
    const g = s('g', { class: 'grid' });
    const gridSt = vAx.grid ? strokeOf(vAx.grid, design, { c: AUTO_LINE, w: 0.75 }) : null;
    if (gridSt) for (const t of sc.ticks) { if (t === sc.min) continue; g.appendChild(applyStroke(s('polygon', { points: Array.from({ length: n }, (_, k) => pt(k, t).map((x) => L.round(x, 2)).join(',')).join(' '), fill: 'none' }), gridSt)); }
    const spoke = strokeOf(cAx.line || (cAx.grid ? cAx.grid : null), design, { c: AUTO_LINE, w: 0.75 });
    if (spoke && !cAx.del) for (let k = 0; k < n; k++) { const p = pt(k, sc.max); g.appendChild(applyStroke(s('line', { x1: cx, y1: cy, x2: p[0], y2: p[1] }), spoke)); }
    svg.appendChild(g);
    const plotG = s('g', { class: 'plot' });
    sis.forEach((i, j) => {
      const sr = series[i];
      const pts = Array.from({ length: n }, (_, k) => (sr.vals[k] == null ? null : pt(k, sr.vals[k])));
      const d = pathOf(pts.concat(pts[0] ? [pts[0]] : []), false);
      const el = s('path', { d, 'stroke-linejoin': 'round' });
      if (g0.radarStyle === 'filled') { applyFill(el, sr.fill || solid(sr.color), design, defs); applyStroke(el, strokeOf(sr.line, design, null)); }
      else { el.setAttribute('fill', 'none'); applyStroke(el, strokeOf(sr.line, design, { c: sr.color, w: 2.25 })); }
      plotG.appendChild(el);
      if (g0.radarStyle === 'marker' || (sr.marker && sr.marker.sym && sr.marker.sym !== 'none')) {
        const sym = (sr.marker && sr.marker.sym && sr.marker.sym !== 'auto') ? sr.marker.sym : AUTO_SYMBOLS[(sr.idx || j) % AUTO_SYMBOLS.length];
        const m = sr.marker || {};
        pts.forEach((p) => { if (p) marker(plotG, sym, p[0], p[1], m.size || 7, m.fill || solid(sr.color), strokeOf(m.line, design, { c: sr.color, w: 0.75 }), design, defs); });
      }
    });
    svg.appendChild(plotG);
    const axG = s('g', { class: 'axes' });
    if (!cAx.del && cAx.lblPos !== 'none') cats.forEach((t, k) => {
      const a = ang(k), x = cx + (r + 6) * Math.cos(a), y = cy + (r + 6 + cSt.sz * 0.3) * Math.sin(a);
      text(axG, [String(t)], x, y, cSt, Math.cos(a) > 0.2 ? 'start' : Math.cos(a) < -0.2 ? 'end' : 'middle');
    });
    if (!vAx.del && vAx.lblPos !== 'none') {
      const fmt = vAx.fmt && !vAx.linked ? vAx.fmt : (series[sis[0]] && series[sis[0]].fmt) || null;
      for (const t of sc.ticks) text(axG, [fmtVal(t, fmt)], cx - 4, cy - rad(t), vSt, 'end');
    }
    svg.appendChild(axG);
  }

  CH._scale = scale;
})();
