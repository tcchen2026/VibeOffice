/* VibeOffice — charts shared by Quire, Ledger and Lectern: the SVG renderer (Insert ▸ Chart, the Chart Wizard
 * preview, slide charts) and the chart data dialog.
 * Model: {kind, title, legend, cats[], series:[{name, vals[], color}], gridY, labels, fsz, axTitleX, axTitleY}.
 * Lectern's chart-draw.js adds CH.renderV2 for charts read from files ({v: 2, …}); the hooks for it are inert elsewhere.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const { h, s } = L;
  const CH = (L.chart = {});
  CH.PALETTE = ['#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00', '#00FFFF'];
  CH.KINDS = [
    ['col', 'Clustered Column'], ['colStacked', 'Stacked Column'], ['colPct', '100% Stacked Column'],
    ['bar', 'Clustered Bar'], ['barStacked', 'Stacked Bar'],
    ['line', 'Line'], ['lineMarkers', 'Line with Markers'],
    ['area', 'Area'], ['areaStacked', 'Stacked Area'],
    ['pie', 'Pie'], ['doughnut', 'Doughnut'], ['scatter', 'XY (Scatter)'],
  ];
  /* kinds only the v2 renderer (Lectern's chart-draw.js) draws; offered when it is loaded */
  const KINDS_V2 = CH.KINDS.slice();
  KINDS_V2.splice(5, 0, ['barPct', '100% Stacked Bar']);
  KINDS_V2.splice(10, 0, ['areaPct', '100% Stacked Area']);
  KINDS_V2.push(['bubble', 'Bubble'], ['radar', 'Radar'], ['radarFilled', 'Filled Radar']);
  CH.kinds = () => (CH.renderV2 ? KINDS_V2 : CH.KINDS);
  /* the design (theme) of what is being edited, for theme colours: Lectern's slide; none in Quire and Ledger */
  const curDesign = () => (L.ui && L.ui.currentDesign ? L.ui.currentDesign() : null);

  /* charts read from a file keep their original part (with the embedded workbook and style parts) until they
     are edited here, so that saving does not lose what Lectern does not model */
  CH.src = new Map();
  CH.keep = (o) => { const id = L.uid('cs'); CH.src.set(id, o); return id; };

  /** the series colour Office picks when a series has no fill of its own (c:style 1–48: grey, colourful, or one accent) */
  CH.autoColor = function (i, n, style, design) {
    const acc = (k) => L.model.resolveColor('accent' + k, design);
    const col = ((style || 2) - 1) % 8;
    const VAR = [[], [['lumMod', 60000]], [['lumMod', 80000], ['lumOff', 20000]], [['lumMod', 80000]], [['lumMod', 60000], ['lumOff', 40000]], [['lumMod', 50000]], [['lumMod', 70000], ['lumOff', 30000]], [['lumMod', 70000]], [['lumMod', 50000], ['lumOff', 50000]]];
    if (col === 1) return L.color.applyMods(acc((i % 6) + 1), VAR[Math.floor(i / 6) % VAR.length].map(([name, val]) => ({ name, val })));
    const t = n > 1 ? i / (n - 1) : 0.5;
    if (col === 0) return L.color.mix('#3F3F3F', '#D9D9D9', t);
    const base = acc(col - 1);
    return t < 0.5 ? L.color.darken(base, (0.5 - t) * 0.7) : L.color.lighten(base, (t - 0.5) * 0.9);
  };

  CH.sample = () => ({
    kind: 'col', title: '', legend: 'r', gridY: true, labels: false, fsz: 12,
    cats: ['1st Qtr', '2nd Qtr', '3rd Qtr', '4th Qtr'],
    series: [
      { name: 'East', vals: [20.4, 27.4, 90, 20.4], color: CH.PALETTE[0] },
      { name: 'West', vals: [30.6, 38.6, 34.6, 31.6], color: CH.PALETTE[1] },
      { name: 'North', vals: [45.9, 46.9, 45, 43.9], color: CH.PALETTE[2] },
    ],
  });

  function niceScale(min, max, ticks) {
    if (min === max) { max = min + 1; }
    const span = max - min;
    const raw = span / (ticks || 5);
    const mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const norm = raw / mag;
    const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
    const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step;
    const out = [];
    for (let v = lo; v <= hi + step / 2; v += step) out.push(L.round(v, 6));
    return { lo, hi, step, ticks: out };
  }
  /** Excel-style number format (the common subset: 0, 0.00, #,##0, %, literal prefixes/suffixes, sections) */
  const fmtCode = (v, code) => {
    let sec = String(code).split(';');
    let f = sec[0];
    if (v < 0 && sec[1]) { f = sec[1]; v = -v; }
    else if (v === 0 && sec[2]) f = sec[2];
    f = f.replace(/\[[^\]]*\]/g, '').replace(/\\(.)/g, '$1').replace(/"([^"]*)"/g, '$1').replace(/_./g, ' ').replace(/\*./g, '');
    const m = /([#0,]*\.?[0#]*)(%?)/.exec(f.replace(/[^#0,.%]/g, (ch) => ch)) ;
    const core = /[#0][#0,]*(\.[0#]+)?%?/.exec(f);
    if (!core) return f.replace(/@/g, String(v));
    const spec = core[0];
    const pct = /%/.test(spec);
    const dec = (/\.([0#]+)/.exec(spec) || [, ''])[1].length;
    const comma = /,/.test(spec.split('.')[0]);
    let x = pct ? v * 100 : v;
    let str = x.toFixed(dec);
    if (comma) { const [i, d] = str.split('.'); str = i.replace(/\B(?=(\d{3})+(?!\d))/g, ',') + (d ? '.' + d : ''); }
    void m;
    return f.slice(0, core.index) + str + (pct ? '%' : '') + f.slice(core.index + spec.length).replace(/%/g, '');
  };
  const fmtNum = (v, code) => {
    if (code) { try { return fmtCode(v, code); } catch (e) { /* fall through */ } }
    if (Math.abs(v) >= 1e6) return L.round(v / 1e6, 2) + 'M';
    if (Math.abs(v) >= 1e4) return L.round(v / 1e3, 1) + 'K';
    return String(L.round(v, 2));
  };
  CH.fmtNum = fmtNum;

  /** Render chart into an <svg> of w×h user units */
  CH.render = function (c, w, hh, design) {
    c = c || CH.sample();
    if (c.v === 2 && CH.renderV2) return CH.renderV2(c, w, hh, design);
    w = Math.max(w, 20); hh = Math.max(hh, 20);
    const fsz = c.fsz || 12;
    const textCol = design ? L.model.resolveColor('tx1', design) : '#000';
    const svg = s('svg', { class: 'chart', width: L.round(w, 2), height: L.round(hh, 2), viewBox: `0 0 ${L.round(w, 2)} ${L.round(hh, 2)}`, overflow: 'visible', 'font-family': L.fontStack(c.font || 'Arial'), 'font-size': fsz, fill: textCol });
    if (c.bg && c.bg.t !== 'none') svg.appendChild(s('rect', { width: w, height: hh, fill: L.render.fillCSS(c.bg, design) }));
    const series = (c.series || []).filter((x) => x);
    const cats = c.cats || [];
    const isPie = c.kind === 'pie' || c.kind === 'doughnut';
    const color = (i) => (series[i] && series[i].color) || CH.PALETTE[i % CH.PALETTE.length];
    const pad = fsz * 0.6;
    let top = pad, left = pad, right = w - pad, bottom = hh - pad;
    if (c.title) {
      /* wrap long titles to the chart width */
      const tf = fsz * 1.4, perLine = Math.max(8, Math.floor((w - pad * 2) / (tf * 0.55)));
      const lines = [];
      for (const para of String(c.title).split('\n')) {
        let cur = '';
        for (const word of para.split(/\s+/)) { if ((cur + ' ' + word).trim().length > perLine && cur) { lines.push(cur); cur = word; } else cur = (cur + ' ' + word).trim(); }
        lines.push(cur);
      }
      lines.slice(0, 4).forEach((ln, i) => {
        const t = s('text', { x: w / 2, y: top + tf * 0.86 + i * tf * 1.15, 'text-anchor': 'middle', 'font-size': tf, 'font-weight': 700 });
        t.textContent = ln;
        svg.appendChild(t);
      });
      top += tf * 1.15 * Math.min(4, lines.length) + fsz * 0.4;
    }
    /* legend */
    const legItems = isPie ? cats.map((n, i) => [n, (c.pieColors && c.pieColors[i]) || CH.PALETTE[i % CH.PALETTE.length]]) : series.map((sr, i) => [sr.name, color(i)]);
    if (c.legend && c.legend !== 'none' && legItems.length) {
      const g = s('g', { class: 'legend' });
      const sw = fsz * 0.8;
      if (c.legend === 'b' || c.legend === 't') {
        const itemW = legItems.map(([n]) => sw * 1.6 + String(n).length * fsz * 0.6 + fsz);
        const total = itemW.reduce((a, b) => a + b, 0);
        let x = Math.max(pad, (w - total) / 2);
        const y = c.legend === 'b' ? bottom - fsz : top + fsz * 0.2;
        legItems.forEach(([n, col], i) => {
          g.appendChild(s('rect', { x, y: y - sw * 0.1, width: sw, height: sw, fill: col, stroke: '#000', 'stroke-width': 0.5 }));
          const t = s('text', { x: x + sw * 1.4, y: y + sw * 0.8 }); t.textContent = n; g.appendChild(t);
          x += itemW[i];
        });
        if (c.legend === 'b') bottom -= fsz * 2; else top += fsz * 1.8;
      } else {
        const maxLen = Math.max(...legItems.map(([n]) => String(n).length));
        const lw = Math.min(w * 0.35, sw * 2 + maxLen * fsz * 0.58 + fsz);
        const lh = legItems.length * fsz * 1.5 + fsz;
        const x0 = c.legend === 'l' ? left : right - lw;
        const y0 = top + (bottom - top - lh) / 2;
        g.appendChild(s('rect', { x: x0, y: y0, width: lw, height: lh, fill: 'none', stroke: '#000', 'stroke-width': 0.5 }));
        legItems.forEach(([n, col], i) => {
          const y = y0 + fsz * 0.7 + i * fsz * 1.5;
          g.appendChild(s('rect', { x: x0 + fsz * 0.5, y, width: sw, height: sw, fill: col, stroke: '#000', 'stroke-width': 0.5 }));
          const t = s('text', { x: x0 + fsz * 0.5 + sw * 1.5, y: y + sw * 0.85 }); t.textContent = n; g.appendChild(t);
        });
        if (c.legend === 'l') left += lw + pad; else right -= lw + pad;
      }
      svg.appendChild(g);
    }

    if (isPie) {
      const sr = series[0] || { vals: [] };
      const vals = sr.vals.map((v) => Math.max(0, +v || 0));
      const tot = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = (left + right) / 2, cy = (top + bottom) / 2, r = Math.max(4, Math.min(right - left, bottom - top) / 2 - pad);
      let a0 = -Math.PI / 2;
      vals.forEach((v, i) => {
        const a1 = a0 + (v / tot) * Math.PI * 2;
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const p0 = [cx + r * Math.cos(a0), cy + r * Math.sin(a0)], p1 = [cx + r * Math.cos(a1), cy + r * Math.sin(a1)];
        const col = (c.pieColors && c.pieColors[i]) || CH.PALETTE[i % CH.PALETTE.length];
        let d;
        if (v / tot >= 0.9999) d = `M${cx - r},${cy} A${r},${r} 0 1 1 ${cx + r},${cy} A${r},${r} 0 1 1 ${cx - r},${cy} Z`;
        else d = `M${cx},${cy} L${L.round(p0[0], 2)},${L.round(p0[1], 2)} A${r},${r} 0 ${large} 1 ${L.round(p1[0], 2)},${L.round(p1[1], 2)} Z`;
        svg.appendChild(s('path', { d, fill: col, stroke: '#000', 'stroke-width': 0.6 }));
        if (c.labels) {
          const am = (a0 + a1) / 2;
          const t = s('text', { x: cx + r * 0.68 * Math.cos(am), y: cy + r * 0.68 * Math.sin(am) + fsz * 0.35, 'text-anchor': 'middle', 'font-size': fsz * 0.9 });
          t.textContent = Math.round((v / tot) * 100) + '%';
          svg.appendChild(t);
        }
        a0 = a1;
      });
      if (c.kind === 'doughnut') svg.appendChild(s('circle', { cx, cy, r: r * 0.5, fill: design ? L.model.resolveColor('bg1', design) : '#fff', stroke: '#000', 'stroke-width': 0.6 }));
      return svg;
    }

    /* axis titles take a band beside / below the plot area */
    if (c.axTitleY && c.axTitleYHoriz) {
      /* a horizontal title beside the vertical axis, wrapped to a fifth of the width */
      const maxCh = Math.max(6, Math.floor((w * 0.2) / (fsz * 0.55)));
      const lines = [];
      let cur = '';
      for (const word of String(c.axTitleY).split(/\s+/)) { if ((cur + ' ' + word).trim().length > maxCh && cur) { lines.push(cur); cur = word; } else cur = (cur + ' ' + word).trim(); }
      lines.push(cur);
      const bw = Math.max(...lines.map((l) => l.length)) * fsz * 0.55;
      const y0 = (top + bottom) / 2 - (lines.length - 1) * fsz * 0.6;
      lines.forEach((ln, i) => { const t = s('text', { x: left + bw / 2, y: y0 + i * fsz * 1.2, 'text-anchor': 'middle', 'font-weight': 700 }); t.textContent = ln; svg.appendChild(t); });
      left += bw + fsz * 0.6;
    } else if (c.axTitleY) {
      const x = left + fsz * 0.9, y = (top + bottom) / 2;
      const t = s('text', { x, y, 'text-anchor': 'middle', 'font-weight': 700, transform: `rotate(-90 ${L.round(x, 2)} ${L.round(y, 2)})` });
      t.textContent = c.axTitleY;
      svg.appendChild(t);
      left += fsz * 1.5;
    }
    if (c.axTitleX) {
      const t = s('text', { x: (left + right) / 2, y: bottom - fsz * 0.2, 'text-anchor': 'middle', 'font-weight': 700 });
      t.textContent = c.axTitleX;
      svg.appendChild(t);
      bottom -= fsz * 1.5;
    }
    const horiz = c.kind === 'bar' || c.kind === 'barStacked';
    const stacked = /Stacked|Pct/.test(c.kind);
    const pct = c.kind === 'colPct';
    const scatter = c.kind === 'scatter';
    let lo = 0, hi = 0;
    const nC = scatter ? 0 : cats.length;
    if (stacked) {
      for (let i = 0; i < nC; i++) {
        let pos = 0, neg = 0;
        for (const sr of series) { const v = +sr.vals[i] || 0; if (v >= 0) pos += v; else neg += v; }
        hi = Math.max(hi, pos); lo = Math.min(lo, neg);
      }
      if (pct) { lo = 0; hi = 100; }
    } else for (const sr of series) for (const v of sr.vals) { const n = +v || 0; hi = Math.max(hi, n); lo = Math.min(lo, n); }
    const sc = niceScale(lo, hi === lo ? lo + 1 : hi, 5);
    const labelW = Math.max(...sc.ticks.map((t) => (c.axisFmt ? fmtNum(t, c.axisFmt) : fmtNum(t)).length)) * fsz * 0.6 + fsz * 0.5;
    /* diagonal category labels need room below the axis */
    const estBand = (w - labelW) / Math.max(nC, 1);
    const longCat = Math.max(1, ...cats.map((x) => String(x).length)) * fsz * 0.55;
    const willRotate = !horiz && !scatter && (c.catRot ? Math.abs(c.catRot) > 1 : longCat > estBand * 0.98);
    const catLabelH = willRotate ? Math.min(hh * 0.35, longCat * (Math.abs(c.catRot || 45) >= 80 ? 1 : 0.72) + fsz) : fsz * 1.6;
    /* plot area */
    let px0, py0, px1, py1;
    if (horiz) {
      const maxCat = Math.max(1, ...cats.map((x) => String(x).length));
      px0 = left + Math.min(w * 0.3, maxCat * fsz * 0.6 + fsz * 0.5); px1 = right; py0 = top; py1 = bottom - catLabelH;
    } else { px0 = left + labelW; px1 = right; py0 = top + fsz * 0.5; py1 = bottom - catLabelH; }
    const pw = Math.max(px1 - px0, 1), ph = Math.max(py1 - py0, 1);
    if (c.plotBg && c.plotBg.t !== 'none') svg.appendChild(s('rect', { x: px0, y: py0, width: pw, height: ph, fill: L.render.fillCSS(c.plotBg, design) }));
    const vmap = (v) => (horiz ? px0 + ((v - sc.lo) / (sc.hi - sc.lo)) * pw : py1 - ((v - sc.lo) / (sc.hi - sc.lo)) * ph);
    const axis = s('g', { class: 'axis', stroke: '#000', 'stroke-width': 0.75 });
    /* gridlines + value labels */
    for (const t of sc.ticks) {
      const p = vmap(t);
      if (c.gridY !== false) axis.appendChild(horiz ? s('line', { x1: p, y1: py0, x2: p, y2: py1, stroke: '#888', 'stroke-width': 0.5 }) : s('line', { x1: px0, y1: p, x2: px1, y2: p, stroke: '#888', 'stroke-width': 0.5 }));
      const tx = horiz ? s('text', { x: p, y: py1 + fsz * 1.2, 'text-anchor': 'middle', stroke: 'none' }) : s('text', { x: px0 - fsz * 0.4, y: p + fsz * 0.35, 'text-anchor': 'end', stroke: 'none' });
      tx.textContent = c.axisFmt ? fmtNum(t, c.axisFmt) : fmtNum(t) + (pct ? '%' : '');
      axis.appendChild(tx);
    }
    axis.appendChild(s('rect', { x: px0, y: py0, width: pw, height: ph, fill: 'none', stroke: '#808080', 'stroke-width': 0.5 }));
    svg.appendChild(axis);
    const g = s('g', { class: 'plot' });
    if (scatter) {
      const xs = cats.map((x) => +x || 0);
      const xsc = niceScale(Math.min(0, ...xs), Math.max(...xs, 1), 5);
      const xm = (v) => px0 + ((v - xsc.lo) / (xsc.hi - xsc.lo)) * pw;
      for (const t of xsc.ticks) { const tx = s('text', { x: xm(t), y: py1 + fsz * 1.2, 'text-anchor': 'middle' }); tx.textContent = fmtNum(t); svg.appendChild(tx); }
      series.forEach((sr, si) => sr.vals.forEach((v, i) => g.appendChild(s('rect', { x: xm(xs[i]) - fsz * 0.3, y: vmap(+v || 0) - fsz * 0.3, width: fsz * 0.6, height: fsz * 0.6, fill: color(si), stroke: '#000', 'stroke-width': 0.5 }))));
      svg.appendChild(g);
      return svg;
    }
    const band = (horiz ? ph : pw) / Math.max(nC, 1);
    /* category labels */
    /* crowded category labels turn diagonal, as in Office */
    const longest = Math.max(1, ...cats.map((x) => String(x).length)) * fsz * 0.55;
    const rot = !horiz ? (c.catRot != null ? c.catRot : longest > band * 0.98 ? -45 : 0) : 0;
    cats.forEach((cat, i) => {
      let t;
      if (horiz) t = s('text', { x: px0 - fsz * 0.4, y: py0 + band * (i + 0.5) + fsz * 0.35, 'text-anchor': 'end' });
      else if (rot) { const x = px0 + band * (i + 0.5), y = py1 + fsz * 0.9; t = s('text', { x, y, 'text-anchor': 'end', transform: `rotate(${rot} ${L.round(x, 2)} ${L.round(y, 2)})`, 'font-size': fsz * 0.9 }); }
      else t = s('text', { x: px0 + band * (i + 0.5), y: py1 + fsz * 1.25, 'text-anchor': 'middle' });
      t.textContent = cat;
      svg.appendChild(t);
    });
    const zero = vmap(Math.max(sc.lo, 0));
    if (c.kind === 'line' || c.kind === 'lineMarkers') {
      series.forEach((sr, si) => {
        const pts = sr.vals.map((v, i) => [px0 + band * (i + 0.5), vmap(+v || 0)]);
        g.appendChild(s('polyline', { points: pts.map((p) => L.round(p[0], 2) + ',' + L.round(p[1], 2)).join(' '), fill: 'none', stroke: color(si), 'stroke-width': Math.max(1.5, fsz * 0.18) }));
        if (c.kind === 'lineMarkers') pts.forEach((p) => g.appendChild(s('rect', { x: p[0] - fsz * 0.25, y: p[1] - fsz * 0.25, width: fsz * 0.5, height: fsz * 0.5, fill: color(si), stroke: '#000', 'stroke-width': 0.4 })));
        if (c.labels) pts.forEach((p, i) => { const t = s('text', { x: p[0], y: p[1] - fsz * 0.5, 'text-anchor': 'middle', 'font-size': fsz * 0.85 }); t.textContent = fmtNum(+sr.vals[i] || 0, c.numFmt); g.appendChild(t); });
      });
    } else if (c.kind === 'area' || c.kind === 'areaStacked') {
      const acc = new Array(nC).fill(0);
      const order = c.kind === 'area' ? series.map((x, i) => i).reverse() : series.map((x, i) => i);
      for (const si of order) {
        const sr = series[si];
        const base = acc.slice();
        const top2 = sr.vals.map((v, i) => (c.kind === 'areaStacked' ? base[i] + (+v || 0) : +v || 0));
        const pts = top2.map((v, i) => [px0 + band * (i + 0.5), vmap(v)]);
        const basePts = base.map((v, i) => [px0 + band * (i + 0.5), c.kind === 'areaStacked' ? vmap(v) : zero]).reverse();
        g.appendChild(s('polygon', { points: pts.concat(basePts).map((p) => L.round(p[0], 2) + ',' + L.round(p[1], 2)).join(' '), fill: color(si), stroke: '#000', 'stroke-width': 0.5 }));
        if (c.kind === 'areaStacked') for (let i = 0; i < nC; i++) acc[i] = top2[i];
      }
    } else {
      const bars = series.map((sr, si) => [sr, si]).filter(([sr]) => !sr.overlay);
      const ns = bars.length || 1;
      const groupW = band * 0.62;
      const bw = stacked ? groupW : groupW / ns;
      for (let i = 0; i < nC; i++) {
        let pos = 0, neg = 0;
        const tot = pct ? bars.reduce((a, [sr]) => a + Math.abs(+sr.vals[i] || 0), 0) || 1 : 1;
        bars.forEach(([sr, si], bi) => {
          let v = +sr.vals[i] || 0;
          if (pct) v = (v / tot) * 100;
          let a, b;
          if (stacked) { if (v >= 0) { a = pos; b = pos + v; pos = b; } else { a = neg; b = neg + v; neg = b; } } else { a = 0; b = v; }
          const off = (band - groupW) / 2 + (stacked ? 0 : bi * bw);
          const p1 = vmap(Math.max(sc.lo, Math.min(a, b))), p2 = vmap(Math.max(a, b));
          const rect = horiz
            ? s('rect', { x: L.round(Math.min(p1, p2), 2), y: L.round(py0 + band * i + off, 2), width: L.round(Math.abs(p2 - p1), 2), height: L.round(bw, 2) })
            : s('rect', { x: L.round(px0 + band * i + off, 2), y: L.round(Math.min(p1, p2), 2), width: L.round(bw, 2), height: L.round(Math.abs(p2 - p1), 2) });
          rect.setAttribute('fill', color(si)); rect.setAttribute('stroke', '#000'); rect.setAttribute('stroke-width', 0.5);
          g.appendChild(rect);
          if (c.labels) {
            const t = horiz ? s('text', { x: Math.max(p1, p2) + fsz * 0.3, y: py0 + band * i + off + bw / 2 + fsz * 0.35, 'font-size': fsz * 0.85 })
              : s('text', { x: px0 + band * i + off + bw / 2, y: Math.min(p1, p2) - fsz * 0.3, 'text-anchor': 'middle', 'font-size': fsz * 0.85 });
            t.textContent = fmtNum(+sr.vals[i] || 0, c.numFmt);
            g.appendChild(t);
          }
        });
      }
      /* overlay series of a combo chart */
      series.forEach((sr, si) => {
        if (!sr.overlay || horiz) return;
        const pts = sr.vals.map((v, i) => [px0 + band * (i + 0.5), vmap(+v || 0)]);
        if (sr.overlay !== 'markers') g.appendChild(s('polyline', { points: pts.map((p) => L.round(p[0], 2) + ',' + L.round(p[1], 2)).join(' '), fill: 'none', stroke: color(si), 'stroke-width': Math.max(1.5, fsz * 0.18) }));
        if (sr.overlay !== 'line') pts.forEach((p) => g.appendChild(s('rect', { x: p[0] - fsz * 0.25, y: p[1] - fsz * 0.25, width: fsz * 0.5, height: fsz * 0.5, fill: color(si), stroke: 'none' })));
      });
    }
    svg.appendChild(g);
    return svg;
  };

  /* ---------- datasheet & chart options dialog ---------- */
  /** chart type → the chart group of a v2 model */
  const groupOf = (kind, old) => {
    const g = { si: [] };
    if (/^(col|bar)/.test(kind)) Object.assign(g, { type: 'bar', dir: /^bar/.test(kind) ? 'bar' : 'col', grouping: /Pct$/.test(kind) ? 'percentStacked' : /Stacked$/.test(kind) ? 'stacked' : 'clustered', gap: old && old.type === 'bar' ? old.gap : 150, overlap: /Pct$|Stacked$/.test(kind) ? 100 : old && old.type === 'bar' && !/stacked/i.test(old.grouping) ? old.overlap : 0 });
    else if (/^line/.test(kind)) Object.assign(g, { type: 'line', grouping: 'standard', marker: kind === 'lineMarkers' });
    else if (/^area/.test(kind)) Object.assign(g, { type: 'area', grouping: kind === 'areaPct' ? 'percentStacked' : kind === 'areaStacked' ? 'stacked' : 'standard' });
    else if (kind === 'pie' || kind === 'doughnut') Object.assign(g, { type: kind, vary: true, firstAng: old && old.firstAng || 0, hole: kind === 'doughnut' ? (old && old.hole) || 50 : 0 });
    else if (/^radar/.test(kind)) Object.assign(g, { type: 'radar', radarStyle: kind === 'radarFilled' ? 'filled' : 'marker' });
    else if (kind === 'bubble') Object.assign(g, { type: 'bubble', bubbleScale: 100 });
    else Object.assign(g, { type: 'scatter', scatterStyle: 'lineMarker' });
    return g;
  };
  /** after the datasheet or options change, bring the v2 drawing model in line with the simple fields */
  CH.syncV2 = function (c, prevKind) {
    if (c.v !== 2) return c;
    const n = c.series.length;
    if (prevKind !== undefined && prevKind !== c.kind) {
      const g = groupOf(c.kind, c.groups && c.groups[0]);
      g.si = c.series.map((_, i) => i);
      c.groups = [g];
      c.series.forEach((sr, i) => { sr.g = 0; delete sr.overlay; if (c.kind === 'line') { sr.marker = Object.assign({}, sr.marker || {}, { sym: 'none' }); } else if (c.kind === 'lineMarkers' && sr.marker && sr.marker.sym === 'none') delete sr.marker.sym; void i; });
      if (c.ax && c.ax.v2) delete c.ax.v2;
      if (c.kind === 'pie' || c.kind === 'doughnut') c.pieColors = c.cats.map((_, k) => CH.autoColor(k, c.cats.length, c.style, curDesign()));
    }
    /* series added or removed in the datasheet */
    for (const g of c.groups) g.si = g.si.filter((i) => i < n);
    const known = new Set(c.groups.flatMap((g) => g.si));
    for (let i = 0; i < n; i++) if (!known.has(i)) { c.series[i].g = 0; c.groups[0].si.push(i); if (c.series[i].idx == null) c.series[i].idx = i; }
    if ((c.kind === 'scatter' || c.kind === 'bubble')) {
      const xs = c.cats.map((x) => parseFloat(x));
      if (xs.every((x) => isFinite(x))) for (const sr of c.series) { sr.xs = xs.slice(); if (c.kind === 'bubble' && !sr.sizes) sr.sizes = sr.vals.map(() => 1); }
    }
    for (const sr of c.series) if (sr.sizes && sr.sizes.length < sr.vals.length) while (sr.sizes.length < sr.vals.length) sr.sizes.push(1);
    c.ax = c.ax || {};
    if (c.gridY === false) { if (c.ax.v) delete c.ax.v.grid; } else if (c.ax.v && !c.ax.v.grid) c.ax.v.grid = {};
    else if (!c.ax.v && c.gridY) c.ax.v = { grid: {} };
    const pie = c.kind === 'pie' || c.kind === 'doughnut';
    if (!c.labels) for (const sr of c.series) delete sr.lbl;
    else for (const sr of c.series) if (!sr.lbl) sr.lbl = pie ? { showPercent: true } : { showVal: true };
    if (!c.title) delete c.titleTx;
    return c;
  };
  const isLineKind = (k) => /^(line|scatter|radar$)/.test(k);
  CH.setSeriesColor = function (c, sr, col) {
    sr.color = col;
    if (c.v !== 2) return;
    const solidC = { t: 'solid', c: col, a: 1 };
    if (isLineKind(c.kind)) { sr.line = Object.assign({}, sr.line && sr.line.t !== 'none' ? sr.line : {}, { c: col }); delete sr.line.t; if (sr.marker) { sr.marker.fill = solidC; sr.marker.line = { c: col }; } }
    else sr.fill = solidC;
  };

  CH.edit = function (shape, onDone) {
    const c = L.clone(shape.chart || CH.sample());
    const ui = L.ui;
    const sheet = h('div', { class: 'datasheet' });
    const preview = h('div', { class: 'chart-preview' });
    const kind = ui.select(CH.kinds().map(([k, n]) => [k, n]), c.kind, (v) => { const prev = c.kind; c.kind = v; CH.syncV2(c, prev); draw(); });
    const title = h('input', { type: 'text', id: 'chart-title', value: c.title || '', oninput: (e) => { c.title = e.target.value; draw(); } });
    const legend = ui.select([['r', 'Right'], ['b', 'Bottom'], ['t', 'Top'], ['l', 'Left'], ['none', 'None']], c.legend || 'r', (v) => { c.legend = v; draw(); });
    const grid = ui.check('Major gridlines', c.gridY !== false, (v) => { c.gridY = v; draw(); });
    const labels = ui.check('Show values', !!c.labels, (v) => { c.labels = v; draw(); });
    function draw() {
      L.clear(preview);
      const design = curDesign();
      if (c.v === 2) {
        /* imported charts are drawn at their real size and scaled into the preview */
        CH.syncV2(c);
        const W = (shape && shape.w) || 300, H = (shape && shape.h) || 190, k = Math.min(300 / W, 190 / H);
        const svg = CH.render(c, W, H, design);
        svg.setAttribute('width', L.round(W * k, 2)); svg.setAttribute('height', L.round(H * k, 2));
        preview.appendChild(svg);
      } else preview.appendChild(CH.render(c, 300, 190, design));
    }
    function buildSheet() {
      L.clear(sheet);
      const t = h('table');
      const head = h('tr', null, h('th', { class: 'corner' }), ...c.cats.map((cat, ci) => h('th', null, h('input', { value: cat, oninput: (e) => { c.cats[ci] = e.target.value; draw(); } }))));
      t.appendChild(head);
      c.series.forEach((sr, si) => {
        const sw = h('span', { class: 'sw', style: `background:${sr.color || CH.PALETTE[si % 12]}` });
        sw.addEventListener('click', () => ui.colorMenu(sw, { mode: 'plain' }, (col) => { CH.setSeriesColor(c, sr, col); sw.style.background = col; draw(); }));
        t.appendChild(h('tr', null,
          h('th', null, sw, h('input', { value: sr.name, oninput: (e) => { sr.name = e.target.value; draw(); } })),
          ...c.cats.map((_, ci) => h('td', null, h('input', { value: sr.vals[ci] == null ? '' : sr.vals[ci], inputmode: 'decimal', oninput: (e) => { const v = parseFloat(e.target.value); sr.vals[ci] = isNaN(v) ? 0 : v; draw(); } })))));
      });
      sheet.appendChild(t);
    }
    const tools = h('div', { class: 'ds-tools' },
      ui.button('Add Column', () => { c.cats.push('Col ' + (c.cats.length + 1)); c.series.forEach((sr) => sr.vals.push(0)); buildSheet(); draw(); }),
      ui.button('Delete Column', () => { if (c.cats.length > 1) { c.cats.pop(); c.series.forEach((sr) => sr.vals.pop()); buildSheet(); draw(); } }),
      ui.button('Add Row', () => { c.series.push({ name: 'Series ' + (c.series.length + 1), vals: c.cats.map(() => 0), color: CH.PALETTE[c.series.length % 12] }); buildSheet(); draw(); }),
      ui.button('Delete Row', () => { if (c.series.length > 1) { c.series.pop(); buildSheet(); draw(); } }));
    const body = h('div', { class: 'chart-dlg' },
      h('div', { class: 'cd-left' },
        h('div', { class: 'cd-row' }, h('label', { text: 'Chart type:' }), kind),
        h('div', { class: 'cd-row' }, h('label', { for: 'chart-title', text: 'Chart title:' }), title),
        h('div', { class: 'cd-row' }, h('label', { text: 'Legend:' }), legend),
        h('div', { class: 'cd-row' }, grid, labels),
        h('fieldset', null, h('legend', { text: 'Datasheet' }), sheet, tools)),
      h('div', { class: 'cd-right' }, h('div', { class: 'cd-cap', text: 'Preview' }), preview));
    buildSheet(); draw();
    ui.dialog({ title: 'Chart', body, width: 720, buttons: [{ label: 'OK', primary: true, onClick: () => { CH.syncV2(c); c.edited = true; onDone(c); } }, { label: 'Cancel' }] });
  };
})(typeof window !== 'undefined' ? window : globalThis);
