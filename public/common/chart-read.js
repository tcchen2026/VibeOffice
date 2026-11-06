/* VibeOffice — the chart model of charts read from files (c:chartSpace → model v2, drawn by chart-draw.js): chart
 * groups and their series, primary and secondary axes with their scaling, markers, data labels, blanks, legend, title,
 * manual layout and 3-D views, with the simple fields (kind, cats, series colours) that charts.js and its datasheet use.
 * Shared by Lectern, Ledger and Quire. CH.read(cx, X, ctx): X supplies the DrawingML readers (kid, kids, at, num, path,
 * desc, fillIn, line, runProps: Lectern's own, or common/dml.js); ctx is theirs, with ctx.design and ctx.accent(k)
 * passed to CH.autoColor for colours left to Office, ctx.chartArea {bg, border} for a chart area without formatting of
 * its own (Word's and Excel's white with a grey border; PowerPoint leaves it empty). Moved from Lectern's pptx-read.js. */
(function (root) {
  'use strict';
  const L = root.L;
  const CH = L.chart;
  /* a chart area without formatting of its own, as Word and Excel 2016 draw it (measured: #868686, 0.75 pt, square) */
  CH.OFFICE_AREA = { bg: { t: 'solid', c: '#FFFFFF', a: 1 }, border: { c: '#868686', w: 0.75 } };
  const CHART_TYPES = /^(bar|bar3D|line|line3D|pie|pie3D|ofPie|doughnut|area|area3D|scatter|radar|bubble|stock|surface|surface3D)Chart$/;
  CH.read = function (cx, X, ctx) {
    const { kid, kids, at, num, path, desc, fillIn, line, runProps } = X;
    const autoSeriesColor = (i, n, style, design) => CH.autoColor(i, n, style, design, ctx.accent);
    const chart = kid(cx, 'chart');
    const plot = kid(chart, 'plotArea');
    if (!plot) return null;
    const cctx = Object.assign({}, ctx, { keepScheme: false, media: () => null, link: () => null, phClr: null });
    const design = ctx.design;
    const style = num(kid(cx, 'style'), 'val', 2);
    const yes = (el, d) => { if (!el) return d; const v = at(el, 'val'); return v == null ? true : v === '1' || v === 'true'; };
    const val = (el, n) => at(kid(el, n), 'val');
    const nval = (el, n, d) => { const v = val(el, n); return v == null || v === '' || isNaN(+v) ? d : +v; };
    const fillOf = (sp) => (sp ? fillIn(sp, cctx) : undefined);
    /* a line without its own colour or width keeps Office's automatic one (c and w left out) */
    const lineOf = (sp) => {
      const ln = sp && kid(sp, 'ln');
      if (!ln) return undefined;
      const l = line(ln, cctx, {});
      if (l && l.t !== 'none') { if (!fillIn(ln, cctx)) delete l.c; if (at(ln, 'w') == null) delete l.w; }
      return l;
    };
    const textOf = (tx) => {
      if (!tx) return null;
      const o = {};
      const d = path(tx, 'p', 'pPr', 'defRPr') || (kid(tx, 'p') && desc(kid(tx, 'p'), 'rPr'));
      if (d) { const r = runProps(d, cctx); if (r.sz) o.sz = r.sz; if (r.b != null) o.b = r.b; if (r.i != null) o.i = r.i; if (r.color) o.color = r.color; if (r.font) o.font = r.font; }
      const bp = kid(tx, 'bodyPr');
      if (bp && at(bp, 'rot') != null && +at(bp, 'rot') !== -60000000) o.rot = +at(bp, 'rot') / 60000;
      if (bp && /^(vert|vert270|wordArtVert|eaVert)$/.test(at(bp, 'vert') || '')) o.rot = at(bp, 'vert') === 'vert270' ? -90 : 90;
      return o;
    };
    const richText = (el) => {
      const rich = el && path(el, 'tx', 'rich');
      if (!rich) return null;
      const lines = kids(rich, 'p').map((p) => kids(p).filter((r) => r.localName === 'r' || r.localName === 'fld').map((r) => (kid(r, 't') || { textContent: '' }).textContent).join(''));
      const t = textOf(rich) || {};
      const r1 = desc(rich, 'r') && kid(desc(rich, 'r'), 'rPr');
      if (r1) { const r = runProps(r1, cctx); if (r.sz) t.sz = r.sz; if (r.b != null) t.b = r.b; if (r.i != null) t.i = r.i; if (r.color) t.color = r.color; if (r.font) t.font = r.font; }
      return { text: lines.join('\n'), tx: t };
    };
    /* cached values of a reference or literal: strings, numbers (null when absent) and the format code */
    const cacheOf = (el) => {
      if (!el) return null;
      const ml = desc(el, 'multiLvlStrCache');
      const cache = ml || desc(el, 'numCache') || desc(el, 'strCache') || desc(el, 'numLit') || desc(el, 'strLit');
      if (!cache) { const v = kid(el, 'v') || desc(el, 'v'); return v ? { str: [v.textContent], num: [+v.textContent], fmt: null, isNum: false } : null; }
      const isNum = /^num/.test(cache.localName);
      const lvl = ml ? kid(cache, 'lvl') || cache : cache;
      const pts = kids(lvl, 'pt');
      let n = num(kid(cache, 'ptCount'), 'val', 0);
      for (const p of pts) n = Math.max(n, num(p, 'idx', 0) + 1);
      const str = new Array(n).fill(''), nums = new Array(n).fill(null);
      const fmt = (kid(cache, 'formatCode') || { textContent: '' }).textContent || null;
      for (const p of pts) {
        const i = num(p, 'idx', 0), t = (kid(p, 'v') || { textContent: '' }).textContent;
        str[i] = t;
        const x = parseFloat(t);
        nums[i] = t !== '' && isFinite(x) ? x : null;
        if (isNum && at(p, 'formatCode')) (str.fmts = str.fmts || [])[i] = at(p, 'formatCode');
      }
      /* the outer levels of a multi-level category axis, innermost first */
      let outer = null;
      if (ml) outer = kids(cache, 'lvl').slice(1).map((lv) => { const a = new Array(n).fill(null); for (const p of kids(lv, 'pt')) a[num(p, 'idx', 0)] = (kid(p, 'v') || { textContent: '' }).textContent; return a; });
      return { str, num: nums, fmt: fmt && fmt !== 'General' ? fmt : null, isNum, outer };
    };
    const fmtCat = (c) => {
      if (!c) return [];
      if (c.isNum && c.fmt && L.numfmt) return c.num.map((v, i) => (v == null ? c.str[i] : L.numfmt.text((c.str.fmts && c.str.fmts[i]) || c.fmt, v)));
      return c.str.slice();
    };
    const lblsOf = (el, inherit) => {
      if (!el) return inherit || null;
      if (yes(kid(el, 'delete'), false)) return { del: true };
      const o = Object.assign({}, inherit && !inherit.del ? inherit : {});
      delete o.pts;
      for (const k of ['showVal', 'showPercent', 'showCatName', 'showSerName', 'showLegendKey', 'showBubbleSize']) { const e = kid(el, k); if (e) o[k] = yes(e, false); }
      if (val(el, 'dLblPos')) o.pos = val(el, 'dLblPos');
      const nf = kid(el, 'numFmt');
      if (nf && at(nf, 'formatCode') && at(nf, 'sourceLinked') !== '1') o.fmt = at(nf, 'formatCode');
      if (kid(el, 'separator')) o.sep = kid(el, 'separator').textContent;
      const rng = kids(path(el, 'extLst')).map((x) => kid(x, 'showDataLabelsRange')).find(Boolean);
      if (rng) o.showRange = yes(rng, false);
      const tx = textOf(kid(el, 'txPr')); if (tx) o.tx = tx;
      const sp = kid(el, 'spPr');
      if (sp) { const f = fillOf(sp), l = lineOf(sp); if (f && f.t !== 'none') o.fill = f; if (l && l.t !== 'none') o.line = l; }
      for (const d of kids(el, 'dLbl')) {
        const i = nval(d, 'idx', -1);
        if (i < 0) continue;
        o.pts = o.pts || {};
        if (yes(kid(d, 'delete'), false)) { o.pts[i] = { del: true }; continue; }
        const one = lblsOf(d, o);
        /* custom label text; fields ([VALUE], [SERIES NAME], [CELLRANGE]…) are filled in when drawn */
        const rich = path(d, 'tx', 'rich');
        if (rich) {
          const runs = [];
          kids(rich, 'p').forEach((p, pi) => {
            if (pi) runs.push('\n');
            for (const r of kids(p)) {
              if (r.localName === 'r') runs.push((kid(r, 't') || { textContent: '' }).textContent);
              else if (r.localName === 'fld') { const ty = String(at(r, 'type') || '').toUpperCase(); runs.push(/^(VALUE|SERIESNAME|CATEGORYNAME|PERCENTAGE|XVALUE|YVALUE|BUBBLESIZE|CELLRANGE)$/.test(ty) ? '\u0001' + ty + '\u0001' : (kid(r, 't') || { textContent: '' }).textContent); }
            }
          });
          one.text = runs.join('');
          const r1 = desc(rich, 'rPr') || path(rich, 'p', 'pPr', 'defRPr');
          if (r1) { const rp = runProps(r1, cctx); one.tx = Object.assign({}, one.tx || {}); if (rp.sz) one.tx.sz = rp.sz; if (rp.b != null) one.tx.b = rp.b; if (rp.i != null) one.tx.i = rp.i; if (rp.color) one.tx.color = rp.color; }
        }
        o.pts[i] = one;
      }
      return o;
    };
    const markerOf = (m) => {
      if (!m) return null;
      const o = {};
      if (val(m, 'symbol')) o.sym = val(m, 'symbol');
      if (val(m, 'size')) o.size = +val(m, 'size');
      const sp = kid(m, 'spPr');
      if (sp) { const f = fillOf(sp), l = lineOf(sp); if (f) o.fill = f; if (l) o.line = l; }
      return o;
    };
    /* axes by id */
    const axes = new Map();
    for (const a of kids(plot)) {
      if (!/^(valAx|catAx|dateAx|serAx)$/.test(a.localName)) continue;
      const o = { kind: a.localName === 'valAx' ? 'val' : a.localName === 'serAx' ? 'ser' : 'cat', id: val(a, 'axId'), cross: val(a, 'crossAx'), date: a.localName === 'dateAx' };
      o.del = yes(kid(a, 'delete'), false);
      o.pos = val(a, 'axPos') || 'b';
      const scl = kid(a, 'scaling');
      if (scl) {
        if (val(scl, 'orientation') === 'maxMin') o.rev = true;
        if (val(scl, 'min') != null) o.min = nval(scl, 'min');
        if (val(scl, 'max') != null) o.max = nval(scl, 'max');
        if (val(scl, 'logBase') != null) o.log = nval(scl, 'logBase', 10);
      }
      if (val(a, 'majorUnit') != null) o.major = nval(a, 'majorUnit');
      const nf = kid(a, 'numFmt');
      if (nf && at(nf, 'formatCode')) { o.fmt = at(nf, 'formatCode'); o.linked = at(nf, 'sourceLinked') === '1'; }
      const gl = kid(a, 'majorGridlines');
      if (gl) { const l = lineOf(kid(gl, 'spPr')); o.grid = l || {}; }
      const gm = kid(a, 'minorGridlines');
      if (gm) { const l = lineOf(kid(gm, 'spPr')); o.minorGrid = l || {}; }
      o.tick = val(a, 'majorTickMark') || 'cross';
      o.lblPos = val(a, 'tickLblPos') || 'nextTo';
      o.crosses = val(a, 'crosses') || (kid(a, 'crossesAt') ? 'at' : 'autoZero');
      if (kid(a, 'crossesAt')) o.crossesAt = nval(a, 'crossesAt', 0);
      if (val(a, 'crossBetween')) o.between = val(a, 'crossBetween') !== 'midCat';
      const sp = kid(a, 'spPr');
      const l = lineOf(sp); if (l) o.line = l;
      const tx = textOf(kid(a, 'txPr')); if (tx) o.tx = tx;
      if (val(a, 'tickLblSkip')) o.skip = nval(a, 'tickLblSkip', 1);
      const t = kid(a, 'title');
      if (t) { const rt = richText(t); o.title = rt ? rt.text : a.localName === 'valAx' ? 'Axis Title' : 'Axis Title'; o.titleTx = (rt && rt.tx) || textOf(kid(t, 'txPr')) || {}; }
      const du = kid(a, 'dispUnits');
      if (du) { const bu = val(du, 'builtInUnit'); const U = { hundreds: 1e2, thousands: 1e3, tenThousands: 1e4, hundredThousands: 1e5, millions: 1e6, tenMillions: 1e7, hundredMillions: 1e8, billions: 1e9, trillions: 1e12 }; o.unit = U[bu] || nval(du, 'custUnit', 1); }
      axes.set(o.id, o);
    }
    /* chart groups */
    const groups = [];
    const series = [];
    let catCache = null;
    const catCaches = [];   // [group, categories] of each series
    for (const ct of kids(plot).filter((c) => CHART_TYPES.test(c.localName))) {
      const ln = ct.localName.replace(/3D|Chart$/g, '').replace('Chart', '');
      let type = { bar: 'bar', line: 'line', pie: 'pie', ofPie: 'pie', doughnut: 'doughnut', area: 'area', scatter: 'scatter', radar: 'radar', bubble: 'bubble', stock: 'stock', surface: 'area' }[ln] || 'bar';
      const g = { type, d3: /3D/.test(ct.localName) };
      g.grouping = val(ct, 'grouping') || (type === 'line' || type === 'area' ? 'standard' : 'clustered');
      if (type === 'bar') {
        g.dir = val(ct, 'barDir') || 'col';
        /* 3-D "standard" columns stand in rows one behind the other */
        if (g.grouping === 'standard') { g.grouping = 'clustered'; if (g.d3) g.deep = true; }
        g.gap = nval(ct, 'gapWidth', 150);
        g.overlap = /stacked/i.test(g.grouping) ? 100 : nval(ct, 'overlap', 0);
      }
      g.vary = yes(kid(ct, 'varyColors'), type === 'pie' || type === 'doughnut');
      if (type === 'pie' || type === 'doughnut') { g.firstAng = nval(ct, 'firstSliceAng', 0); g.hole = type === 'doughnut' ? nval(ct, 'holeSize', 50) : 0; }
      if (type === 'scatter') g.scatterStyle = val(ct, 'scatterStyle') || 'marker';
      if (type === 'radar') g.radarStyle = val(ct, 'radarStyle') || 'standard';
      if (type === 'bubble') { g.bubbleScale = nval(ct, 'bubbleScale', 100); g.sizeArea = val(ct, 'sizeRepresents') !== 'w'; }
      if (type === 'line' || type === 'stock') { g.marker = yes(kid(ct, 'marker'), true); if (kid(ct, 'hiLowLines')) g.hiLow = lineOf(kid(kid(ct, 'hiLowLines'), 'spPr')) || {}; if (kid(ct, 'dropLines')) g.drop = lineOf(kid(kid(ct, 'dropLines'), 'spPr')) || {}; if (kid(ct, 'upDownBars')) g.upDown = true; }
      const ax = kids(ct, 'axId').map((e) => at(e, 'val'));
      g.axIds = ax;
      const gl = lblsOf(kid(ct, 'dLbls'), null);
      g.si = [];
      for (const s of kids(ct, 'ser')) {
        const i = series.length;
        const sr = { g: groups.length, idx: nval(s, 'idx', i), order: nval(s, 'order', i) };
        const tx = kid(s, 'tx');
        const nc = cacheOf(tx);
        sr.name = (nc && nc.str[0]) || (tx && kid(tx, 'v') ? kid(tx, 'v').textContent : '') || 'Series ' + (sr.idx + 1);
        if (tx) sr.named = true;
        const cEl = kid(s, 'cat') || kid(s, 'xVal');
        const cc = cacheOf(cEl);
        if (cc) catCaches.push([groups.length, cc]);
        if (kid(s, 'xVal')) sr.xs = cc ? (cc.isNum ? cc.num.slice() : cc.str.map((_, k) => k + 1)) : null;
        const vc = cacheOf(kid(s, 'val') || kid(s, 'yVal'));
        sr.vals = vc ? vc.num.slice() : [];
        if (vc && vc.fmt) sr.fmt = vc.fmt;
        if (type === 'bubble') { const bc = cacheOf(kid(s, 'bubbleSize')); sr.sizes = bc ? bc.num.slice() : sr.vals.map(() => 1); }
        const sp = kid(s, 'spPr');
        const f = fillOf(sp), l = lineOf(sp);
        if (f) sr.fill = f;
        if (l) sr.line = l;
        const mk = markerOf(kid(s, 'marker')); if (mk) sr.marker = mk;
        if (kid(s, 'smooth')) sr.smooth = yes(kid(s, 'smooth'), false);
        if (kid(s, 'explosion')) sr.expl = nval(s, 'explosion', 0);
        if (yes(kid(s, 'invertIfNegative'), false)) sr.invNeg = true;
        for (const dp of kids(s, 'dPt')) {
          const k = nval(dp, 'idx', -1);
          if (k < 0) continue;
          const o = {};
          const dsp = kid(dp, 'spPr');
          const df = fillOf(dsp), dl = lineOf(dsp);
          if (df) o.fill = df; if (dl) o.line = dl;
          if (kid(dp, 'explosion')) o.expl = nval(dp, 'explosion', 0);
          const dm = markerOf(kid(dp, 'marker')); if (dm) o.marker = dm;
          (sr.pts = sr.pts || {})[k] = o;
        }
        const sl = lblsOf(kid(s, 'dLbls'), gl);
        /* labels taken from a cell range (Office 2013 "Value From Cells") */
        const dr = desc(s, 'dlblRangeCache');
        if (sl && dr) { const rc = cacheOf(dr.parentNode); if (rc) sl.range = rc.str; }
        if (sl && !sl.del && (sl.showVal || sl.showPercent || sl.showCatName || sl.showSerName || sl.showBubbleSize || sl.pts)) sr.lbl = sl;
        /* the colour that stands for the series: legend, datasheet, editing */
        const solid = (x) => (x && x.t === 'solid' ? x.c : x && x.t === 'grad' && x.stops[0] ? x.stops[0].c : x && x.t === 'patt' ? x.fg : null);
        const lineType = type === 'line' || type === 'scatter' || type === 'radar' && g.radarStyle !== 'filled' || type === 'stock';
        sr.color = (lineType ? (l && l.t !== 'none' && l.c && l.c[0] === '#' ? l.c : null) || solid(mk && mk.fill) : solid(f)) || null;
        g.si.push(i);
        series.push(sr);
      }
      groups.push(g);
    }
    if (!groups.length) return null;
    /* secondary axes: a group whose value axis is not the first group's */
    /* XY charts list the horizontal value axis first */
    const isXY = (g) => g.type === 'scatter' || g.type === 'bubble';
    const valAxOf = (g) => (isXY(g) ? axes.get(g.axIds[1]) || null : g.axIds.map((id) => axes.get(id)).find((a) => a && a.kind === 'val') || null);
    const catAxOf = (g) => (isXY(g) ? axes.get(g.axIds[0]) || null : g.axIds.map((id) => axes.get(id)).find((a) => a && a.kind === 'cat') || null);
    const v1 = valAxOf(groups[0]), c1 = catAxOf(groups[0]);
    let v2 = null, c2 = null;
    for (const g of groups) {
      const va = valAxOf(g);
      /* (a deleted secondary value axis leaves its series on the primary scale, as Office plots them) */
      if (va && v1 && va !== v1 && !va.del) { g.sec = true; v2 = va; c2 = catAxOf(g); }
    }
    /* the category axis shows the longest categories of the series on it (another range a group with an axis of its
       own refers to stays on that axis, shown or not) */
    for (const [gi, cc] of catCaches) if (catAxOf(groups[gi]) === c1 && (!catCache || cc.str.length > catCache.str.length)) catCache = cc;
    if (!catCache) for (const [, cc] of catCaches) if (!catCache || cc.str.length > catCache.str.length) catCache = cc;
    const nSer = series.length;
    /* colours left to Office's automatic choice */
    for (const sr of series) {
      const auto = autoSeriesColor(sr.idx, Math.max(nSer, sr.idx + 1), style, design);
      if (!sr.color) sr.color = auto;
      sr.auto = auto;
    }
    const cats = catCache ? fmtCat(catCache) : (series[0] ? series[0].vals.map((_, i) => String(i + 1)) : []);
    const main = groups[0];
    const m = { v: 2, style, groups, series, cats };
    if (catCache && catCache.outer && catCache.outer.length) m.catOuter = catCache.outer;
    if (catCache && catCache.isNum) m.catNums = catCache.num.slice();
    /* legacy fields: chart type, datasheet and editing */
    const k = main.type;
    if (k === 'bar') m.kind = main.dir === 'bar' ? (main.grouping === 'percentStacked' ? 'barPct' : main.grouping === 'stacked' ? 'barStacked' : 'bar') : main.grouping === 'percentStacked' ? 'colPct' : main.grouping === 'stacked' ? 'colStacked' : 'col';
    else if (k === 'line' || k === 'stock') m.kind = series.some((s) => s.g === 0 && (!s.marker || s.marker.sym !== 'none')) && main.marker !== false ? 'lineMarkers' : 'line';
    else if (k === 'area') m.kind = main.grouping === 'percentStacked' ? 'areaPct' : main.grouping === 'stacked' ? 'areaStacked' : 'area';
    else if (k === 'pie') m.kind = 'pie';
    else if (k === 'doughnut') m.kind = 'doughnut';
    else if (k === 'radar') m.kind = main.radarStyle === 'filled' ? 'radarFilled' : 'radar';
    else if (k === 'bubble') m.kind = 'bubble';
    else m.kind = 'scatter';
    for (const sr of series) if (sr.g !== 0) sr.overlay = groups[sr.g].type === 'line' || groups[sr.g].type === 'scatter' ? (sr.line && sr.line.t === 'none' ? 'markers' : sr.marker && sr.marker.sym === 'none' ? 'line' : 'lineMarkers') : groups[sr.g].type;
    m.ax = {};
    if (c1) m.ax.c = c1; if (v1) m.ax.v = v1; if (v2) m.ax.v2 = v2; if (c2 && c2 !== c1) m.ax.c2 = c2;
    m.gridY = !!(v1 && v1.grid);
    if (c1 && c1.grid) m.gridX = true;
    /* text: chart-wide default, title, legend */
    const tx0 = textOf(kid(cx, 'txPr')) || {};
    m.fsz = tx0.sz || 10; if (tx0.font) m.font = tx0.font; if (tx0.color) m.textColor = tx0.color; if (tx0.b) m.bold = true;
    const title = kid(chart, 'title');
    const autoDel = yes(kid(chart, 'autoTitleDeleted'), false);
    m.title = '';
    if (title) {
      const rt = richText(title);
      /* (a title without text of its own: the one series' name when the file names it, as Office shows it; Excel 2016) */
      m.title = rt ? rt.text : series.length === 1 && series[0].named ? series[0].name : 'Chart Title';
      m.titleTx = (rt && rt.tx) || textOf(kid(title, 'txPr')) || {};
      if (yes(kid(title, 'overlay'), false)) m.titleOverlay = true;
      const tsp = kid(title, 'spPr'); if (tsp) { const f = fillOf(tsp), l = lineOf(tsp); if (f && f.t !== 'none') m.titleFill = f; if (l && l.t !== 'none') m.titleLine = l; }
    } else if (!autoDel && series.length === 1) { m.title = series[0].name; m.titleTx = {}; } /* Office's automatic title for a single series */
    const leg = kid(chart, 'legend');
    m.legend = 'none';
    if (leg) {
      m.legend = val(leg, 'legendPos') || 'r';
      const lt = textOf(kid(leg, 'txPr')); if (lt) m.legendTx = lt;
      if (yes(kid(leg, 'overlay'), false)) m.legendOverlay = true;
      const lsp = kid(leg, 'spPr'); if (lsp) { const f = fillOf(lsp), l = lineOf(lsp); if (f && f.t !== 'none') m.legendFill = f; if (l && l.t !== 'none') m.legendLine = l; }
      const del = kids(leg, 'legendEntry').filter((e) => yes(kid(e, 'delete'), false)).map((e) => nval(e, 'idx', -1));
      if (del.length) m.legendDel = del;
      const ml = path(leg, 'layout', 'manualLayout');
      if (ml && val(ml, 'x') != null && val(ml, 'w') != null) m.legendBox = { x: nval(ml, 'x', 0), y: nval(ml, 'y', 0), w: nval(ml, 'w', 0), h: nval(ml, 'h', 0), edge: (val(ml, 'xMode') || 'factor') === 'edge' };
    }
    /* chart area and plot area */
    const csp = kid(cx, 'spPr');
    if (csp) { const f = fillOf(csp), l = lineOf(csp); if (f && f.t !== 'none') m.bg = f; if (l && l.t !== 'none') m.border = l; }
    /* (without formatting of its own the chart area is the host's: Word and Excel draw it white with a grey border) */
    else if (ctx.chartArea) { m.bg = Object.assign({}, ctx.chartArea.bg); m.border = Object.assign({}, ctx.chartArea.border); }
    const psp = kid(plot, 'spPr');
    if (psp) { const f = fillOf(psp), l = lineOf(psp); if (f && f.t !== 'none') m.plotBg = f; if (l && l.t !== 'none') m.plotBorder = l; }
    const pml = path(plot, 'layout', 'manualLayout');
    if (pml && val(pml, 'w') != null && val(pml, 'h') != null) m.plotBox = { x: nval(pml, 'x', 0), y: nval(pml, 'y', 0), w: nval(pml, 'w', 1), h: nval(pml, 'h', 1), inner: val(pml, 'layoutTarget') === 'inner', edge: (val(pml, 'xMode') || 'factor') === 'edge' };
    if (val(chart, 'dispBlanksAs')) m.blanks = val(chart, 'dispBlanksAs');
    /* 3-D charts: the view angles and the walls */
    if (groups.some((g) => g.d3)) {
      const v3 = kid(chart, 'view3D');
      m.view3D = { rotX: nval(v3, 'rotX', groups[0].type === 'pie' ? 30 : 15), rotY: nval(v3, 'rotY', groups[0].type === 'pie' ? 0 : 20), depth: nval(v3, 'depthPercent', 100), rAng: yes(kid(v3, 'rAngAx'), false) };
      const wall = (n) => { const w = kid(chart, n); const sp = w && kid(w, 'spPr'); if (!sp) return null; const f = fillOf(sp), l = lineOf(sp); return { fill: f, line: l }; };
      m.walls = { back: wall('backWall'), side: wall('sideWall'), floor: wall('floor') };
    }
    /* legacy summary fields */
    const anyLbl = series.find((s) => s.lbl);
    m.labels = !!anyLbl;
    const fmt1 = series[0] && series[0].fmt;
    if (anyLbl && anyLbl.lbl.fmt) m.numFmt = anyLbl.lbl.fmt; else if (fmt1) m.numFmt = fmt1;
    if (v1 && v1.fmt && (!v1.linked || !fmt1) && v1.fmt !== 'General') m.axisFmt = v1.fmt; else if (v1 && v1.linked && fmt1) m.axisFmt = fmt1;
    if (c1 && c1.tx && c1.tx.rot != null) m.catRot = c1.tx.rot;
    /* axis titles by where the axis stands (a horizontal one's is X), for the editors and writers of the simple fields */
    for (const a of [c1, v1]) if (a && a.title && !a.del) { if (/[bt]/.test(a.pos)) m.axTitleX = m.axTitleX || a.title; else m.axTitleY = m.axTitleY || a.title; }
    if (k === 'pie' || k === 'doughnut') {
      const s0 = series.find((s) => s.g === 0);
      if (s0) m.pieColors = cats.map((_, i) => { const p = s0.pts && s0.pts[i]; const f = p && p.fill; return f && f.t === 'solid' ? f.c : main.vary ? autoSeriesColor(i, cats.length, style, design) : s0.color; });
    }
    return m;
  };
})(typeof window !== 'undefined' ? window : globalThis);
