/* VibeOffice — AutoShape geometry, shared by Quire, Ledger and Lectern.
 * Preset names follow the OOXML prstGeom vocabulary so shapes round-trip unchanged. Each preset is drawn from
 * DrawingML's own definition (common/presets.js) as SVG path parts in a w×h box, with its text rectangle and handles.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const n = (v) => (Math.abs(v) < 1e-9 ? 0 : Math.round(v * 100) / 100);
  /* ---------- preset table: each preset's name as Office shows it, what kind of shape it is, its adjustments for a new
     one; the shapes themselves are DrawingML's own definitions (common/presets.js), evaluated below ---------- */
  const G = {};
  let markers = 0;
  const def = (name, label, opts) => { G[name] = Object.assign({ name, label }, opts || {}); };
  /* basic shapes */
  def('rect', 'Rectangle');
  def('roundRect', 'Rounded Rectangle', { adj: { adj: 16667 } });
  def('ellipse', 'Oval');
  def('triangle', 'Isosceles Triangle', { adj: { adj: 50000 } });
  def('rtTriangle', 'Right Triangle');
  def('parallelogram', 'Parallelogram', { adj: { adj: 25000 } });
  def('trapezoid', 'Trapezoid', { adj: { adj: 25000 } });
  def('diamond', 'Diamond');
  def('pentagon', 'Regular Pentagon');
  def('hexagon', 'Hexagon', { adj: { adj: 25000 } });
  def('octagon', 'Octagon', { adj: { adj: 29289 } });
  def('plus', 'Cross', { adj: { adj: 25000 } });
  def('can', 'Can', { adj: { adj: 25000 } });
  def('cube', 'Cube', { adj: { adj: 25000 } });
  def('bevel', 'Bevel', { adj: { adj: 12500 } });
  def('foldedCorner', 'Folded Corner', { adj: { adj: 16667 } });
  def('frame', 'Frame', { adj: { adj1: 12500 } });
  def('donut', 'Donut', { adj: { adj: 25000 } });
  def('noSmoking', '"No" Symbol', { adj: { adj: 18750 } });
  def('smileyFace', 'Smiley Face', { adj: { adj: 4653 } });
  def('heart', 'Heart');
  def('lightningBolt', 'Lightning Bolt');
  def('sun', 'Sun', { adj: { adj: 25000 } });
  def('moon', 'Moon', { adj: { adj: 50000 } });
  def('arc', 'Arc', { noFill: true, adj: { adj1: 16200000, adj2: 0 } });
  def('wave', 'Wave', { adj: { adj1: 12500 } });
  /* block arrows */
  def('rightArrow', 'Right Arrow', { adj: { adj1: 50000, adj2: 50000 } });
  def('leftArrow', 'Left Arrow', { adj: { adj1: 50000, adj2: 50000 } });
  def('upArrow', 'Up Arrow', { adj: { adj1: 50000, adj2: 50000 } });
  def('downArrow', 'Down Arrow', { adj: { adj1: 50000, adj2: 50000 } });
  def('leftRightArrow', 'Left-Right Arrow', { adj: { adj1: 50000, adj2: 50000 } });
  def('upDownArrow', 'Up-Down Arrow', { adj: { adj1: 50000, adj2: 50000 } });
  def('quadArrow', 'Quad Arrow', { adj: { adj1: 22500, adj2: 22500, adj3: 22500 } });
  def('notchedRightArrow', 'Notched Right Arrow', { adj: { adj1: 50000, adj2: 50000 } });
  def('stripedRightArrow', 'Striped Right Arrow', { adj: { adj1: 50000, adj2: 50000 } });
  def('homePlate', 'Pentagon', { adj: { adj: 50000 } });
  def('chevron', 'Chevron', { adj: { adj: 50000 } });
  def('bentArrow', 'Bent Arrow', { adj: { adj1: 25000, adj2: 25000, adj3: 25000, adj4: 43750 } });
  def('uturnArrow', 'U-Turn Arrow', { adj: { adj1: 25000, adj2: 25000, adj3: 25000, adj4: 43750, adj5: 75000 } });
  def('circularArrow', 'Circular Arrow', { adj: { adj1: 12500, adj2: 1142319, adj3: 20457681, adj4: 10800000, adj5: 12500 } });
  /* stars and banners */
  def('star4', '4-Point Star', { adj: { adj: 12500 } });
  def('star5', '5-Point Star', { adj: { adj: 19098 } });
  def('star6', '6-Point Star', { adj: { adj: 28868 } });
  def('star8', '8-Point Star', { adj: { adj: 37500 } });
  def('star16', '16-Point Star', { adj: { adj: 37500 } });
  def('star24', '24-Point Star', { adj: { adj: 37500 } });
  def('star32', '32-Point Star', { adj: { adj: 37500 } });
  def('irregularSeal1', 'Explosion 1');
  def('irregularSeal2', 'Explosion 2');
  def('ribbon2', 'Up Ribbon');
  def('verticalScroll', 'Vertical Scroll');
  def('horizontalScroll', 'Horizontal Scroll');
  /* flowchart */
  def('flowChartProcess', 'Flowchart: Process');
  def('flowChartAlternateProcess', 'Flowchart: Alternate Process');
  def('flowChartDecision', 'Flowchart: Decision');
  def('flowChartInputOutput', 'Flowchart: Data');
  def('flowChartPredefinedProcess', 'Flowchart: Predefined Process');
  def('flowChartInternalStorage', 'Flowchart: Internal Storage');
  def('flowChartDocument', 'Flowchart: Document');
  def('flowChartMultidocument', 'Flowchart: Multidocument');
  def('flowChartTerminator', 'Flowchart: Terminator');
  def('flowChartPreparation', 'Flowchart: Preparation');
  def('flowChartManualInput', 'Flowchart: Manual Input');
  def('flowChartManualOperation', 'Flowchart: Manual Operation');
  def('flowChartConnector', 'Flowchart: Connector');
  def('flowChartOffpageConnector', 'Flowchart: Off-page Connector');
  def('flowChartPunchedCard', 'Flowchart: Card');
  def('flowChartPunchedTape', 'Flowchart: Punched Tape');
  def('flowChartSummingJunction', 'Flowchart: Summing Junction');
  def('flowChartOr', 'Flowchart: Or');
  def('flowChartCollate', 'Flowchart: Collate');
  def('flowChartSort', 'Flowchart: Sort');
  def('flowChartExtract', 'Flowchart: Extract');
  def('flowChartMerge', 'Flowchart: Merge');
  def('flowChartDelay', 'Flowchart: Delay');
  def('flowChartMagneticDisk', 'Flowchart: Magnetic Disk');
  def('flowChartDisplay', 'Flowchart: Display');
  /* callouts */
  def('wedgeRectCallout', 'Rectangular Callout', { adj: { adj1: -20833, adj2: 62500 } });
  def('wedgeRoundRectCallout', 'Rounded Rectangular Callout', { adj: { adj1: -20833, adj2: 62500 } });
  def('wedgeEllipseCallout', 'Oval Callout', { adj: { adj1: -20833, adj2: 62500 } });
  def('cloudCallout', 'Cloud Callout', { adj: { adj1: -20833, adj2: 62500 } });
  def('cloud', 'Cloud');
  /* brackets and braces: open shapes, drawn without a fill */
  def('leftBracket', 'Left Bracket', { line: true, adj: { adj: 8333 } });
  def('rightBracket', 'Right Bracket', { line: true, adj: { adj: 8333 } });
  def('leftBrace', 'Left Brace', { line: true, adj: { adj1: 8333, adj2: 50000 } });
  def('rightBrace', 'Right Brace', { line: true, adj: { adj1: 8333, adj2: 50000 } });
  def('bracketPair', 'Double Bracket', { line: true, adj: { adj: 16667 } });
  def('bracePair', 'Double Brace', { line: true, adj: { adj: 8333 } });
  /* lines and connectors, drawn from (0,0) to (w,h) honoring flips (linePath) */
  def('line', 'Line', { isLine: true });
  def('straightConnector1', 'Straight Connector', { isLine: true });
  def('bentConnector3', 'Elbow Connector', { isLine: true });
  def('bentConnector2', 'Elbow Connector', { isLine: true });
  def('bentConnector4', 'Elbow Connector', { isLine: true });
  def('curvedConnector2', 'Curved Connector', { isLine: true });
  def('curvedConnector3', 'Curved Connector', { isLine: true });
  /* Office 2007's additions */
  def('pie', 'Pie', { adj: { adj1: 0, adj2: 16200000 } });
  def('pieWedge', 'Pie Wedge');
  def('chord', 'Chord', { adj: { adj1: 2700000, adj2: 16200000 } });
  def('blockArc', 'Block Arc', { adj: { adj1: 10800000, adj2: 0, adj3: 25000 } });
  def('teardrop', 'Teardrop', { adj: { adj: 100000 } });
  def('corner', 'L-Shape', { adj: { adj1: 50000, adj2: 50000 } });
  def('halfFrame', 'Half Frame', { adj: { adj1: 33333, adj2: 33333 } });
  def('diagStripe', 'Diagonal Stripe', { adj: { adj: 50000 } });
  def('plaque', 'Plaque', { adj: { adj: 16667 } });
  def('bentUpArrow', 'Bent-Up Arrow', { adj: { adj1: 25000, adj2: 25000, adj3: 25000 } });
  def('leftUpArrow', 'Left-Up Arrow', { adj: { adj1: 25000, adj2: 25000, adj3: 25000 } });
  def('snip1Rect', 'Snip Single Corner Rectangle', { adj: { adj: 16667 } });
  def('snip2SameRect', 'Snip Same Side Corner Rectangle', { adj: { adj1: 16667, adj2: 0 } });
  def('snip2DiagRect', 'Snip Diagonal Corner Rectangle', { adj: { adj1: 0, adj2: 16667 } });
  def('snipRoundRect', 'Snip and Round Single Corner Rectangle', { adj: { adj1: 16667, adj2: 16667 } });
  def('round1Rect', 'Round Single Corner Rectangle', { adj: { adj: 16667 } });
  def('round2SameRect', 'Round Same Side Corner Rectangle', { adj: { adj1: 16667, adj2: 0 } });
  def('round2DiagRect', 'Round Diagonal Corner Rectangle', { adj: { adj1: 16667, adj2: 0 } });
  def('heptagon', 'Heptagon');
  def('decagon', 'Decagon');
  def('dodecagon', 'Dodecagon');
  def('star7', '7-Point Star', { adj: { adj: 34601 } });
  def('star10', '10-Point Star', { adj: { adj: 42533 } });
  def('star12', '12-Point Star', { adj: { adj: 37500 } });
  def('flowChartOnlineStorage', 'Flowchart: Stored Data');
  def('flowChartOfflineStorage', 'Flowchart: Offline Storage');
  def('flowChartMagneticTape', 'Flowchart: Sequential Access Storage');
  def('mathPlus', 'Plus', { adj: { adj1: 23520 } });
  def('mathMinus', 'Minus', { adj: { adj1: 23520 } });
  def('mathEqual', 'Equal', { adj: { adj1: 23520, adj2: 11760 } });
  def('mathMultiply', 'Multiply', { adj: { adj1: 23520 } });
  def('mathDivide', 'Division', { adj: { adj1: 23520, adj2: 5880, adj3: 11760 } });
  /* action buttons: a new one links where its kind goes */
  def('actionButtonBlank', 'Action Button: Custom', { action: 'blank' });
  def('actionButtonHome', 'Action Button: Home', { action: 'home' });
  def('actionButtonHelp', 'Action Button: Help', { action: 'help' });
  def('actionButtonInformation', 'Action Button: Information', { action: 'info' });
  def('actionButtonBackPrevious', 'Action Button: Back or Previous', { action: 'back' });
  def('actionButtonForwardNext', 'Action Button: Forward or Next', { action: 'next' });
  def('actionButtonBeginning', 'Action Button: Beginning', { action: 'begin' });
  def('actionButtonEnd', 'Action Button: End', { action: 'end' });
  def('actionButtonReturn', 'Action Button: Return', { action: 'return' });
  def('actionButtonDocument', 'Action Button: Document', { action: 'doc' });
  def('actionButtonSound', 'Action Button: Sound', { action: 'sound' });
  def('actionButtonMovie', 'Action Button: Movie', { action: 'movie' });

  /* ---------- DrawingML's own preset definitions (common/presets.js), evaluated as Office does ---------- */
  const ANG = Math.PI / 10800000;   // DrawingML angles are 60000ths of a degree
  const FMLA = {
    '*/': (x, y, z) => (x * y) / z, '+-': (x, y, z) => x + y - z, '+/': (x, y, z) => (x + y) / z, '?:': (x, y, z) => (x > 0 ? y : z),
    abs: (x) => Math.abs(x), at2: (x, y) => Math.atan2(y, x) / ANG, cat2: (x, y, z) => x * Math.cos(Math.atan2(z, y)),
    cos: (x, y) => x * Math.cos(y * ANG), max: (x, y) => Math.max(x, y), min: (x, y) => Math.min(x, y),
    mod: (x, y, z) => Math.sqrt(x * x + y * y + z * z), pin: (x, y, z) => (y < x ? x : y > z ? z : y),
    sat2: (x, y, z) => x * Math.sin(Math.atan2(z, y)), sin: (x, y) => x * Math.sin(y * ANG), sqrt: (x) => Math.sqrt(Math.max(0, x)),
    tan: (x, y) => x * Math.tan(y * ANG), val: (x) => x,
  };
  const spec = (name) => (L.presetShapes && Object.prototype.hasOwnProperty.call(L.presetShapes, name) ? L.presetShapes[name] : null);
  /** the guides of a preset in a w×h box: the shape's own adjust values (adj), else the preset's, then its formulas */
  function guideValues(def, w, h, adj) {
    const v = { w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, ss: Math.min(w, h), ls: Math.max(w, h),
      cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000 };
    const get = (a) => {
      if (typeof a === 'number') return a;
      if (a in v) return v[a];
      const m = /^(w|h|ss)d(\d+)$/.exec(a);
      return m ? v[m[1]] / +m[2] : 0;
    };
    const run = (g) => { v[g[0]] = FMLA[g[1]](get(g[2]), get(g[3]), get(g[4])); };
    for (const g of def.av || []) { if (adj && adj[g[0]] != null && isFinite(adj[g[0]])) v[g[0]] = +adj[g[0]]; else run(g); }
    for (const g of def.gd || []) run(g);
    v.get = get;
    return v;
  }
  /** an arc as DrawingML's arcTo draws it: from (x, y), on the wR×hR ellipse through that point, from visual angle
      st sweeping sw (degrees, clockwise), scaled by sx, sy; SVG commands and the end point */
  function arcTo(x, y, wR, hR, st, sw, sx, sy) {
    const par = (deg) => { const t = deg * Math.PI / 180; return Math.atan2(wR * Math.sin(t), hR * Math.cos(t)); };
    const p0 = par(st), cx = x - wR * Math.cos(p0), cy = y - hR * Math.sin(p0);
    if (!(wR > 0 && hR > 0) || !sw) {
      const e = par(st + sw);
      return { d: `L${n((cx + wR * Math.cos(e)) * sx)},${n((cy + hR * Math.sin(e)) * sy)} `, x: cx + wR * Math.cos(e), y: cy + hR * Math.sin(e) };
    }
    /* in pieces of at most 90° so that whole turns draw too */
    const k = Math.max(1, Math.ceil(Math.abs(sw) / 90 - 1e-9));
    let d = '', ex = x, ey = y;
    for (let i = 1; i <= k; i++) {
      const e = par(st + (sw * i) / k);
      ex = cx + wR * Math.cos(e); ey = cy + hR * Math.sin(e);
      d += `A${n(wR * sx)},${n(hR * sy)} 0 0 ${sw > 0 ? 1 : 0} ${n(ex * sx)},${n(ey * sy)} `;
    }
    return { d, x: ex, y: ey };
  }
  /** a preset's paths in a w×h box, as drawing parts */
  function specParts(def, w, h, adj) {
    const v = guideValues(def, w, h, adj), get = v.get, out = [];
    for (const p of def.path || []) {
      const sx = p.w ? w / p.w : 1, sy = p.h ? h / p.h : 1;
      let d = '', x = 0, y = 0, x0 = 0, y0 = 0, subs = 0;
      const pt = (i, c) => `${n(get(c[i]) * sx)},${n(get(c[i + 1]) * sy)}`;
      for (const c of p.c) {
        switch (c[0]) {
          case 'M': x = x0 = get(c[1]); y = y0 = get(c[2]); d += `M${pt(1, c)} `; subs++; break;
          case 'L': x = get(c[1]); y = get(c[2]); d += `L${pt(1, c)} `; break;
          case 'Q': d += `Q${pt(1, c)} ${pt(3, c)} `; x = get(c[3]); y = get(c[4]); break;
          case 'C': d += `C${pt(1, c)} ${pt(3, c)} ${pt(5, c)} `; x = get(c[5]); y = get(c[6]); break;
          case 'A': { const a = arcTo(x, y, get(c[1]), get(c[2]), get(c[3]) / 60000, get(c[4]) / 60000, sx, sy); d += a.d; x = a.x; y = a.y; break; }
          case 'Z': d += 'Z '; x = x0; y = y0; break;
          default: break;
        }
      }
      const part = { d: d.trim(), fill: p.fill || 'norm', stroke: p.stroke !== 0 };
      if (subs > 1 && part.fill !== 'none') part.rule = 'evenodd';
      if (!/Z$/.test(part.d)) part.open = true;
      out.push(part);
    }
    return out;
  }
  function specText(def, w, h, adj) {
    if (!def.rect) return [0, 0, w, h];
    const v = guideValues(def, w, h, adj);
    return def.rect.map(v.get);
  }
  /** the value in [lo, hi] (from cur) where cost is least: sampled, then narrowed; ranges too wide to sample (callout
      tails reach anywhere) are solved from the slope, the handles being linear in them */
  function least(cost, lo, hi, cur, signed) {
    if (hi < lo) [lo, hi] = [hi, lo];
    if (hi - lo > 1e7 && signed) {
      let p0 = L.clamp(cur, lo, hi), f0 = signed(p0);
      for (let i = 0; i < 6 && Math.abs(f0) > 0.01; i++) {
        const p1 = p0 + 1000, slope = (signed(p1) - f0) / 1000;
        if (!slope) break;
        p0 = L.clamp(p0 - f0 / slope, lo, hi); f0 = signed(p0);
      }
      return p0;
    }
    let best = L.clamp(cur, lo, hi), bc = cost(best);
    const k = 72;
    for (let i = 0; i <= k; i++) { const p = lo + ((hi - lo) * i) / k, c = cost(p); if (c < bc) { bc = c; best = p; } }
    let a = Math.max(lo, best - (hi - lo) / k), b = Math.min(hi, best + (hi - lo) / k);
    for (let i = 0; i < 40; i++) {
      const m1 = a + (b - a) / 3, m2 = b - (b - a) / 3;
      if (cost(m1) < cost(m2)) b = m2; else a = m1;
    }
    const p = (a + b) / 2;
    return cost(p) < bc ? p : best;
  }
  /** a preset's adjust handles (ahXY, ahPolar): where each one is, and the adjust values that bring it under the pointer */
  function specHandles(def) {
    return (def.ah || []).map((ah) => {
      const [kind, ref1, min1, max1, ref2, min2, max2, px, py] = ah;
      const where = (w, h, adj) => { const v = guideValues(def, w, h, adj); return [v.get(px), v.get(py)]; };
      return {
        pos: where,
        set(x, y, w, h, adj) {
          const a = Object.assign({}, adj), v = guideValues(def, w, h, a), out = {};
          const range = (lo, hi) => [v.get(lo), v.get(hi)];
          const at = (ref, val) => where(w, h, Object.assign({}, a, { [ref]: val }));
          const dist = (ref) => (val) => { const p = at(ref, val); return Math.hypot(p[0] - x, p[1] - y); };
          if (kind === 'xy') {
            if (ref1) { const [lo, hi] = range(min1, max1); a[ref1] = out[ref1] = least((val) => Math.abs(at(ref1, val)[0] - x), lo, hi, +v[ref1] || 0, (val) => at(ref1, val)[0] - x); }
            if (ref2) { const [lo, hi] = range(min2, max2); a[ref2] = out[ref2] = least((val) => Math.abs(at(ref2, val)[1] - y), lo, hi, +v[ref2] || 0, (val) => at(ref2, val)[1] - y); }
          } else {
            /* polar: the angle first, then the radius, twice over */
            for (let round = 0; round < 2; round++) {
              if (ref2) { const [lo, hi] = range(min2, max2); a[ref2] = out[ref2] = least(dist(ref2), lo, hi, +a[ref2] || +v[ref2] || 0); }
              if (ref1) { const [lo, hi] = range(min1, max1); a[ref1] = out[ref1] = least(dist(ref1), lo, hi, +a[ref1] || +v[ref1] || 0); }
            }
          }
          return out;
        },
      };
    });
  }
  const handleCache = new Map();

  /* ---------- text warps (bodyPr's prstTxWarp: L.presetWarps, presetTextWarpDefinitions.xml) ----------
   * As PowerPoint 2016 bends text, measured in its PDF of tools/shapes/warp-deck.py's 228 cells (2026-11-07):
   * - a warp fills the whole shape: the text box's insets and anchoring don't count;
   * - its paths come in pairs, the top and bottom of a band (Inflate, Can Up, Plain…), or alone (Arch Up, Circle, and
   *   Button's three: an odd number are all alone); lines are dealt out among the bands or paths in order, as evenly
   *   as they go, the first ones taking one more;
   * - a band's lines are stretched to fill it: their ink from the start of each path to its end, by the distance
   *   along it, and from the one path down to the other;
   * - a path alone carries its lines at their size, placed along it as they are aligned (shrunk to its length only
   *   when longer); the path runs 0.19 em above the baseline of the line on the inside of its curve, and the others
   *   are stacked outward from it. */
  const warpDef = (name) => (L.presetWarps && Object.prototype.hasOwnProperty.call(L.presetWarps, name) ? L.presetWarps[name] : null);
  /** a warp's paths in a w×h box, as polylines with the distance along them: [{ pts: [[x, y]…], at: [0…], len }] */
  function warpPaths(def, w, h, adj) {
    const get = guideValues(def, w, h, adj).get;
    return (def.path || []).map((p) => {
      const pts = [];
      let x = 0, y = 0;
      const add = (k, f) => { for (let i = 1; i <= k; i++) pts.push(f(i / k)); };
      for (const c of p.c) {
        const g = (i) => get(c[i]);
        switch (c[0]) {
          case 'M': case 'L': x = g(1); y = g(2); pts.push([x, y]); break;
          case 'Q': { const x0 = x, y0 = y, x1 = g(1), y1 = g(2); x = g(3); y = g(4); const x2 = x, y2 = y; add(48, (t) => { const s = 1 - t; return [s * s * x0 + 2 * s * t * x1 + t * t * x2, s * s * y0 + 2 * s * t * y1 + t * t * y2]; }); break; }
          case 'C': { const x0 = x, y0 = y, x1 = g(1), y1 = g(2), x2 = g(3), y2 = g(4); x = g(5); y = g(6); const x3 = x, y3 = y; add(64, (t) => { const s = 1 - t; return [s * s * s * x0 + 3 * s * s * t * x1 + 3 * s * t * t * x2 + t * t * t * x3, s * s * s * y0 + 3 * s * s * t * y1 + 3 * s * t * t * y2 + t * t * t * y3]; }); break; }
          case 'A': {
            /* as arcTo: from the current point, on the wR×hR ellipse through it, by visual angles */
            const wR = g(1), hR = g(2), st = g(3) / 60000, sw = g(4) / 60000;
            const par = (deg) => { const t = (deg * Math.PI) / 180; return Math.atan2(wR * Math.sin(t), hR * Math.cos(t)); };
            const p0 = par(st), cx = x - wR * Math.cos(p0), cy = y - hR * Math.sin(p0);
            add(Math.max(8, Math.ceil(Math.abs(sw) / 2)), (t) => { const e = par(st + sw * t); return [cx + wR * Math.cos(e), cy + hR * Math.sin(e)]; });
            [x, y] = pts[pts.length - 1];
            break;
          }
          default: break;
        }
      }
      const at = [0];
      for (let i = 1; i < pts.length; i++) at.push(at[i - 1] + Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]));
      return { pts, at, len: at[at.length - 1] };
    });
  }
  /** the point at distance d along a path (beyond its ends, on along its first or last piece) */
  function along(p, d) {
    const { pts, at } = p;
    if (pts.length < 2) return pts[0] || [0, 0];
    let lo = 1, hi = at.length - 1;
    while (lo < hi) { const m = (lo + hi) >> 1; if (at[m] < d) lo = m + 1; else hi = m; }
    while (lo > 1 && at[lo] === at[lo - 1]) lo--;
    const a = pts[lo - 1], b = pts[lo], t = at[lo] > at[lo - 1] ? (d - at[lo - 1]) / (at[lo] - at[lo - 1]) : 0;
    return [a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t];
  }
  /** the unit tangent at distance d along a path */
  function tangent(p, d) {
    const e = Math.max(p.len / 400, 0.01), a = along(p, d - e), b = along(p, d + e), l = Math.hypot(b[0] - a[0], b[1] - a[1]) || 1;
    return [(b[0] - a[0]) / l, (b[1] - a[1]) / l];
  }
  /** which way a path turns overall: > 0 clockwise on the page (y down), < 0 anticlockwise */
  function turning(p) {
    let s = 0;
    for (let i = 2; i < p.pts.length; i++) {
      const [a, b, c] = [p.pts[i - 2], p.pts[i - 1], p.pts[i]];
      s += (b[0] - a[0]) * (c[1] - b[1]) - (b[1] - a[1]) * (c[0] - b[0]);
    }
    return s;
  }
  /** a path moved off along its normal (positive: to the right of its way, down for a path going right) */
  function parallel(p, off) {
    const pts = p.pts.map((pt, i) => {
      const a = p.pts[Math.max(0, i - 1)], b = p.pts[Math.min(p.pts.length - 1, i + 1)], l = Math.hypot(b[0] - a[0], b[1] - a[1]) || 1;
      return [pt[0] - ((b[1] - a[1]) / l) * off, pt[1] + ((b[0] - a[0]) / l) * off];
    });
    const at = [0];
    for (let i = 1; i < pts.length; i++) at.push(at[i - 1] + Math.hypot(pts[i][0] - pts[i - 1][0], pts[i][1] - pts[i - 1][1]));
    return { pts, at, len: at[at.length - 1] };
  }
  /** the SVG matrix that takes a glyph from its own coordinates (its baseline start at 0, 0) to where the map f puts
      the text space point (x, y) → [X, Y], made affine about the glyph's centre (cx, cy) in text space */
  function glyphMatrix(f, gx, by, cx, cy, sx, sy) {
    const p = f(cx, cy), px = f(cx + sx, cy), mx = f(cx - sx, cy), py = f(cx, cy + sy), my = f(cx, cy - sy);
    const a = (px[0] - mx[0]) / (2 * sx), b = (px[1] - mx[1]) / (2 * sx), c = (py[0] - my[0]) / (2 * sy), d = (py[1] - my[1]) / (2 * sy);
    return [a, b, c, d, p[0] + a * (gx - cx) + c * (by - cy), p[1] + b * (gx - cx) + d * (by - cy)];
  }
  let warpCanvas = null;
  const COMPLEX = /[\u0590-\u08ff\u0900-\u0dff\u0e00-\u0eff\u0f00-\u0fff\u1000-\u109f\u1780-\u17ff\ufb1d-\ufdff\ufe70-\ufeff]/;
  const SLICE = 6;   // the widest a slice of a glyph in a band is drawn
  /** a line of runs [{ text, family (a CSS font stack), size (px), bold, italic }] measured for L.geom.warp: its
      glyphs (the characters drawn, or a whole run of a right-to-left or joined script, with the run each is in and its
      advance), width, ink box about its baseline and font size (measured at a large size and scaled: the ink of a
      small size is rounded to pixels, and warps stretch it many times); pending: a promise of the fonts it needs that
      aren't loaded yet (measured with a stand-in till then) */
  function warpLine(runs) {
    if (!warpCanvas) warpCanvas = document.createElement('canvas').getContext('2d');
    const ctx = warpCanvas, glyphs = [], waits = [], M = 240;
    let x = 0, ink = null, size = 0;
    runs.forEach((r, ri) => {
      const style = `${r.italic ? 'italic ' : ''}${r.bold ? 'bold ' : ''}`, k = (r.size || 0) / M, text = r.text || '';
      ctx.font = `${style}${M}px ${r.family}`;
      if (document.fonts && !document.fonts.check(`${style}${r.size}px ${r.family}`, text || ' ')) waits.push(document.fonts.load(`${style}${r.size}px ${r.family}`, text || ' '));
      const m = ctx.measureText(text);
      /* right-to-left and joined scripts are bent a run at a time, in one piece, so the browser orders and shapes them */
      if (COMPLEX.test(text)) { if (/\S/.test(text)) glyphs.push({ ch: text, run: ri, x, w: m.width * k }); }
      else {
        let pre = '';
        for (const ch of Array.from(text)) {
          const x0 = ctx.measureText(pre).width;
          pre += ch;
          if (/\S/.test(ch)) glyphs.push({ ch, run: ri, x: x + x0 * k, w: (ctx.measureText(pre).width - x0) * k });
        }
      }
      if (/\S/.test(text)) {
        const box = [x - m.actualBoundingBoxLeft * k, -m.actualBoundingBoxAscent * k, x + m.actualBoundingBoxRight * k, m.actualBoundingBoxDescent * k];
        ink = ink ? [Math.min(ink[0], box[0]), Math.min(ink[1], box[1]), Math.max(ink[2], box[2]), Math.max(ink[3], box[3])] : box;
      }
      x += m.width * k;
      size = Math.max(size, r.size || 0);
    });
    return { glyphs, width: x, ink: ink || [0, 0, 0, 0], size, pending: waits.length ? Promise.all(waits) : null };
  }
  /** text bent by a warp in a w×h shape. lines: warpLine's, each with pitch (to the next baseline, px) and
      algn ('l', 'ctr', 'r'). Returns each line's glyphs, each as pieces { m, clip }: m the SVG matrix [a, b, c, d, e,
      f] from the glyph's own coordinates (its baseline start at 0, 0) to the shape's, clip the part of the glyph it
      takes, [x0, x1] across it (null: all of it). [] for a line the warp has no room for; null for an unknown warp */
  function warp(name, adj, w, h, lines) {
    const def = warpDef(name);
    if (!def || !def.path || !def.path.length) return null;
    const paths = warpPaths(def, w, h, adj), pairs = paths.length % 2 === 0;
    const slots = pairs ? paths.length / 2 : paths.length, out = lines.map(() => []);
    /* lines dealt out: as evenly as they go, the first slots one more */
    const per = Array.from({ length: slots }, (_, i) => Math.floor(lines.length / slots) + (i < lines.length % slots ? 1 : 0));
    /* in bands, the lines are aligned within the widest of them all, and their ink from side to side spans each band */
    const width = Math.max(0, ...lines.map((ln) => ln.width));
    const dxOf = (ln) => (ln.algn === 'ctr' ? (width - ln.width) / 2 : ln.algn === 'r' ? width - ln.width : 0);
    const inked = lines.filter((ln) => ln.glyphs.length);
    const x0 = Math.min(...inked.map((ln) => dxOf(ln) + ln.ink[0])), x1 = Math.max(...inked.map((ln) => dxOf(ln) + ln.ink[2]));
    let next = 0;
    per.forEach((count, slot) => {
      const idx = Array.from({ length: count }, (_, k) => next + k);
      next += count;
      if (!count) return;
      if (pairs) {
        const top = paths[slot * 2], bot = paths[slot * 2 + 1];
        /* the band's lines, baselines a pitch apart: their ink from top to bottom fills it */
        let base = 0, y0 = Infinity, y1 = -Infinity;
        const bases = idx.map((i, k) => {
          const ln = lines[i];
          if (k) base += ln.pitch;
          if (ln.glyphs.length) { y0 = Math.min(y0, base + ln.ink[1]); y1 = Math.max(y1, base + ln.ink[3]); }
          return base;
        });
        if (!(x1 > x0 && y1 > y0)) return;
        const f = (x, y) => {
          const u = (x - x0) / (x1 - x0), v = (y - y0) / (y1 - y0), p = along(top, u * top.len), q = along(bot, u * bot.len);
          return [p[0] + (q[0] - p[0]) * v, p[1] + (q[1] - p[1]) * v];
        };
        idx.forEach((i, k) => {
          const ln = lines[i], dx = dxOf(ln), b = bases[k], cy = b + (ln.ink[1] + ln.ink[3]) / 2, sy = Math.max(1, (ln.ink[3] - ln.ink[1]) / 2);
          /* each glyph in upright slices, each bent about its own middle, so that it follows the curves */
          out[i] = ln.glyphs.map((gl) => {
            const whole = glyphMatrix(f, dx + gl.x, b, dx + gl.x + gl.w / 2, cy, Math.max(0.5, gl.w / 2), sy);
            const n = Math.min(6, Math.max(1, Math.ceil((Math.hypot(whole[0], whole[1]) * gl.w) / SLICE)));
            return Array.from({ length: n }, (_, k) => ({
              m: glyphMatrix(f, dx + gl.x, b, dx + gl.x + (gl.w * (k + 0.5)) / n, cy, Math.max(0.5, gl.w / n / 2), sy),
              clip: n > 1 ? [k ? (gl.w * k) / n : -1e4, k < n - 1 ? (gl.w * (k + 1)) / n : 1e4] : null,
            }));
          });
        });
      } else {
        const p = paths[slot], outward = turning(p) > 0;
        /* the line on the inside of the curve sits on the path; the others are stacked outward, a pitch apart, each
           along its own curve parallel to the path (so their letters keep their spacing) */
        const order = outward ? idx.slice().reverse() : idx;
        let off = 0;
        order.forEach((i, k) => {
          const ln = lines[i];
          if (k) off += outward ? -lines[order[k - 1]].pitch : ln.pitch;
          const q = off ? parallel(p, off) : p, s = ln.width > q.len && ln.width > 0 ? q.len / ln.width : 1;
          const base = 0.19 * ln.size * s, start = ln.algn === 'ctr' ? (q.len - ln.width * s) / 2 : ln.algn === 'r' ? q.len - ln.width * s : 0;
          out[i] = ln.glyphs.map((gl) => {
            const d = start + (gl.x + gl.w / 2) * s, pt = along(q, d), [tx, ty] = tangent(q, d), nx = -ty, ny = tx;
            /* the glyph turned with the path at its middle, its baseline `base` below the path */
            const a = tx * s, b = ty * s, c = nx * s, dd = ny * s, ox = pt[0] + nx * base, oy = pt[1] + ny * base;
            return [{ m: [a, b, c, dd, ox - a * (gl.w / 2), oy - b * (gl.w / 2)], clip: null }];
          });
        });
      }
    });
    return out;
  }

  let warpIds = 0;
  const svgEl = (tag, a) => { const e = document.createElementNS('http://www.w3.org/2000/svg', tag); for (const k in a) e.setAttribute(k, a[k]); return e; };
  const r4 = (v) => Math.round(v * 1e4) / 1e4;
  /** text bent by a warp, as SVG over a w×h shape: lines [{ runs: [{ text, family (a CSS font stack), size, bold,
      italic, paint }], algn, pitch, size (an empty line's) }]. paint(run.paint, defs) gives a run's { fill (a colour, a
      url(#…) painted across the whole shape through its glyphs, or null), fillOpacity, stroke (a colour or null),
      strokeWidth, strokeOpacity }; runs with the same paint are drawn together, the fill then the outline over it,
      each as opaque as it is (the slices of a glyph overlap a little). An outline keeps its width however its glyphs
      are stretched. Drawn again in place once fonts it waits for have loaded. null for an unknown warp. */
  function warpSVG(name, adj, w, h, lines, paint) {
    const measured = lines.map((ln) => Object.assign(warpLine(ln.runs), { algn: ln.algn, pitch: ln.pitch }));
    measured.forEach((m, i) => { if (!m.size) m.size = lines[i].size || 0; });
    const mats = warp(name, adj, w, h, measured);
    if (!mats) return null;
    const svg = svgEl('svg', { class: 'warp', width: r4(w), height: r4(h), viewBox: `0 0 ${r4(w)} ${r4(h)}`, overflow: 'visible', 'aria-hidden': 'true' });
    const defs = svgEl('defs', {}), groups = new Map(), id = () => 'warp' + ++warpIds;
    svg.appendChild(defs);
    measured.forEach((ln, li) => ln.glyphs.forEach((gl, gi) => {
      const run = lines[li].runs[gl.run], key = JSON.stringify(run.paint);
      if (!groups.has(key)) groups.set(key, { paint: run.paint, glyphs: [] });
      for (const { m, clip } of mats[li][gi] || []) {
        const t = svgEl('text', { transform: `matrix(${m.map(r4).join(' ')})`, 'font-family': run.family, 'font-size': run.size, 'font-weight': run.bold ? 700 : 400, 'font-style': run.italic ? 'italic' : 'normal' });
        t.textContent = gl.ch;
        if (clip) {
          /* a slice of the glyph, in its own coordinates */
          const cid = id(), cp = svgEl('clipPath', { id: cid }), o = 0.4 / Math.max(0.01, Math.hypot(m[0], m[1]));
          cp.appendChild(svgEl('rect', { x: r4(clip[0] - o), y: -1e4, width: r4(clip[1] - clip[0] + 2 * o), height: 2e4 }));
          defs.appendChild(cp);
          t.setAttribute('clip-path', `url(#${cid})`);
        }
        groups.get(key).glyphs.push({ t, m });
      }
    }));
    const layer = (opacity) => { const g = svgEl('g', {}); if (opacity != null && opacity < 1) g.setAttribute('opacity', r4(opacity)); svg.appendChild(g); return g; };
    for (const { paint: key, glyphs } of groups.values()) {
      const p = paint(key, defs);
      if (p.fill && /^url\(/.test(p.fill)) {
        const cid = id(), clip = svgEl('clipPath', { id: cid });
        for (const { t } of glyphs) clip.appendChild(t.cloneNode(true));
        defs.appendChild(clip);
        layer(p.fillOpacity).appendChild(svgEl('rect', { width: r4(w), height: r4(h), fill: p.fill, 'clip-path': `url(#${cid})` }));
      } else if (p.fill) {
        const g = layer(p.fillOpacity);
        for (const { t } of glyphs) { t.setAttribute('fill', p.fill); g.appendChild(t); }
      }
      if (!p.stroke) continue;
      const g = layer(p.strokeOpacity);
      for (const { t, m } of glyphs) {
        const o = t.cloneNode(true);
        o.setAttribute('fill', 'none'); o.setAttribute('stroke', p.stroke); o.setAttribute('stroke-linejoin', 'round');
        o.setAttribute('stroke-width', r4(p.strokeWidth / (Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1)));
        g.appendChild(o);
      }
    }
    const pending = measured.map((m) => m.pending).filter(Boolean);
    if (pending.length) Promise.all(pending).then(() => {
      const again = warpSVG(name, adj, w, h, lines, paint);
      if (!again || !svg.parentNode) return;
      for (const k of ['class', 'style', 'width', 'height']) if (svg.hasAttribute(k)) again.setAttribute(k, svg.getAttribute(k));   // as the app set them
      svg.replaceWith(again);
    }, () => {});
    return svg;
  }

  /* ---------- custom geometry (freeforms, imported custGeom) ----------
   * path = { paths: [{ w, h, cmds: [['M',x,y], ['L',x,y], ['C',...6], ['Q',...4], ['A', wR, hR, stDeg, swDeg], ['Z']], fill, stroke }] }
   */
  function customParts(path, W, H) {
    const out = [];
    for (const p of path.paths || []) {
      const sx = p.w ? W / p.w : 1, sy = p.h ? H / p.h : 1;
      let d = '', cx = 0, cy = 0, sxp = 0, syp = 0;
      for (const c of p.cmds) {
        switch (c[0]) {
          case 'M': cx = c[1]; cy = c[2]; sxp = cx; syp = cy; d += `M${n(cx * sx)},${n(cy * sy)} `; break;
          case 'L': cx = c[1]; cy = c[2]; d += `L${n(cx * sx)},${n(cy * sy)} `; break;
          case 'C': d += `C${n(c[1] * sx)},${n(c[2] * sy)} ${n(c[3] * sx)},${n(c[4] * sy)} ${n(c[5] * sx)},${n(c[6] * sy)} `; cx = c[5]; cy = c[6]; break;
          case 'Q': d += `Q${n(c[1] * sx)},${n(c[2] * sy)} ${n(c[3] * sx)},${n(c[4] * sy)} `; cx = c[3]; cy = c[4]; break;
          case 'A': { const a = arcTo(cx, cy, c[1], c[2], c[3], c[4], sx, sy); d += a.d; cx = a.x; cy = a.y; break; }
          case 'Z': d += 'Z '; cx = sxp; cy = syp; break;
          default: break;
        }
      }
      const part = { d: d.trim(), fill: p.fill === 'none' ? 'none' : p.fill || 'norm', stroke: p.stroke !== false, open: !/Z\s*$/.test(d.trim()) };
      if ((d.match(/M/g) || []).length > 1 && part.fill !== 'none') part.rule = 'evenodd';   // holes, as the presets' (specParts)
      out.push(part);
    }
    return out;
  }

  /* ---------- public API ---------- */
  /* Preset pattern fills (a:pattFill prst, ECMA-376 ST_PresetPatternVal): 8×8 bitmaps, rows top to bottom, high bit
     left, anchored at the shape's top-left corner. Taken from what PowerPoint draws: its PDF of a deck with each preset
     in a rectangle embeds every tile as a 64×64 image (tools/patterns/). */
  const PATTERNS = {
    pct5: [0x80, 0, 0, 0, 0x08, 0, 0, 0], pct10: [0x80, 0, 0x08, 0, 0x80, 0, 0x08, 0], pct20: [0x88, 0, 0x22, 0, 0x88, 0, 0x22, 0],
    pct25: [0x88, 0x22, 0x88, 0x22, 0x88, 0x22, 0x88, 0x22], pct30: [0xAA, 0x44, 0xAA, 0x11, 0xAA, 0x44, 0xAA, 0x11],
    pct40: [0xAA, 0x55, 0xAA, 0x51, 0xAA, 0x55, 0xAA, 0x15], pct50: [0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55, 0xAA, 0x55],
    pct60: [0xEE, 0x55, 0xBB, 0x55, 0xEE, 0x55, 0xBB, 0x55], pct70: [0x77, 0xDD, 0x77, 0xDD, 0x77, 0xDD, 0x77, 0xDD],
    pct75: [0x77, 0xFF, 0xDD, 0xFF, 0x77, 0xFF, 0xDD, 0xFF], pct80: [0xEF, 0xFF, 0xFE, 0xFF, 0xEF, 0xFF, 0xFE, 0xFF],
    pct90: [0xFF, 0xFF, 0xFF, 0xF7, 0xFF, 0xFF, 0xFF, 0x7F], horz: [0xFF, 0, 0, 0, 0, 0, 0, 0], vert: [0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
    ltHorz: [0xFF, 0, 0, 0, 0xFF, 0, 0, 0], ltVert: [0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88, 0x88], dkHorz: [0xFF, 0xFF, 0, 0, 0xFF, 0xFF, 0, 0],
    dkVert: [0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC, 0xCC], narHorz: [0xFF, 0, 0xFF, 0, 0xFF, 0, 0xFF, 0],
    narVert: [0x55, 0x55, 0x55, 0x55, 0x55, 0x55, 0x55, 0x55], dashHorz: [0xF0, 0, 0, 0, 0x0F, 0, 0, 0],
    dashVert: [0x80, 0x80, 0x80, 0x80, 0x08, 0x08, 0x08, 0x08], cross: [0xFF, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80],
    dnDiag: [0x80, 0x40, 0x20, 0x10, 0x08, 0x04, 0x02, 0x01], upDiag: [0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40, 0x80],
    ltDnDiag: [0x88, 0x44, 0x22, 0x11, 0x88, 0x44, 0x22, 0x11], ltUpDiag: [0x11, 0x22, 0x44, 0x88, 0x11, 0x22, 0x44, 0x88],
    dkDnDiag: [0xCC, 0x66, 0x33, 0x99, 0xCC, 0x66, 0x33, 0x99], dkUpDiag: [0x33, 0x66, 0xCC, 0x99, 0x33, 0x66, 0xCC, 0x99],
    wdDnDiag: [0xC1, 0xE0, 0x70, 0x38, 0x1C, 0x0E, 0x07, 0x83], wdUpDiag: [0x83, 0x07, 0x0E, 0x1C, 0x38, 0x70, 0xE0, 0xC1],
    dashDnDiag: [0, 0, 0x88, 0x44, 0x22, 0x11, 0, 0], dashUpDiag: [0, 0, 0x11, 0x22, 0x44, 0x88, 0, 0],
    diagCross: [0x81, 0x42, 0x24, 0x18, 0x18, 0x24, 0x42, 0x81], smCheck: [0x99, 0x66, 0x66, 0x99, 0x99, 0x66, 0x66, 0x99],
    lgCheck: [0xF0, 0xF0, 0xF0, 0xF0, 0x0F, 0x0F, 0x0F, 0x0F], smGrid: [0xFF, 0x88, 0x88, 0x88, 0xFF, 0x88, 0x88, 0x88],
    lgGrid: [0xFF, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80, 0x80], dotGrid: [0xAA, 0, 0x80, 0, 0x80, 0, 0x80, 0],
    smConfetti: [0x80, 0x08, 0x40, 0x02, 0x10, 0x01, 0x20, 0x04], lgConfetti: [0xB1, 0x30, 0x03, 0x1B, 0xD8, 0xC0, 0x0C, 0x8D],
    horzBrick: [0xFF, 0x80, 0x80, 0x80, 0xFF, 0x08, 0x08, 0x08], diagBrick: [0x01, 0x02, 0x04, 0x08, 0x18, 0x24, 0x42, 0x81],
    solidDmnd: [0x10, 0x38, 0x7C, 0xFE, 0x7C, 0x38, 0x10, 0], openDmnd: [0x82, 0x44, 0x28, 0x10, 0x28, 0x44, 0x82, 0x01],
    dotDmnd: [0x80, 0, 0x22, 0, 0x08, 0, 0x22, 0], plaid: [0xAA, 0x55, 0xAA, 0x55, 0xF0, 0xF0, 0xF0, 0xF0],
    sphere: [0x77, 0x89, 0x8F, 0x8F, 0x77, 0x98, 0xF8, 0xF8], weave: [0x88, 0x54, 0x22, 0x45, 0x88, 0x14, 0x22, 0x51],
    divot: [0, 0x10, 0x08, 0x10, 0, 0x80, 0x01, 0x80], shingle: [0x03, 0x84, 0x48, 0x30, 0x0C, 0x02, 0x01, 0x01],
    wave: [0, 0x18, 0x25, 0xC0, 0, 0x18, 0x25, 0xC0], trellis: [0xFF, 0x66, 0xFF, 0x99, 0xFF, 0x66, 0xFF, 0x99],
    zigZag: [0x81, 0x42, 0x24, 0x18, 0x81, 0x42, 0x24, 0x18],
  };
  /* each pattern's name, as Office's pattern pickers show it (UI text, translated where it is shown) */
  const N = (x) => x;
  const PATTERN_NAMES = { pct5: '5%', pct10: '10%', pct20: '20%', pct25: '25%', pct30: '30%', pct40: '40%', pct50: '50%', pct60: '60%', pct70: '70%', pct75: '75%', pct80: '80%', pct90: '90%',
    horz: N('Horizontal'), vert: N('Vertical'), ltHorz: N('Light horizontal'), ltVert: N('Light vertical'), dkHorz: N('Dark horizontal'), dkVert: N('Dark vertical'),
    narHorz: N('Narrow horizontal'), narVert: N('Narrow vertical'), dashHorz: N('Dashed horizontal'), dashVert: N('Dashed vertical'), cross: N('Cross'),
    dnDiag: N('Downward diagonal'), upDiag: N('Upward diagonal'), ltDnDiag: N('Light downward diagonal'), ltUpDiag: N('Light upward diagonal'),
    dkDnDiag: N('Dark downward diagonal'), dkUpDiag: N('Dark upward diagonal'), wdDnDiag: N('Wide downward diagonal'), wdUpDiag: N('Wide upward diagonal'),
    dashDnDiag: N('Dashed downward diagonal'), dashUpDiag: N('Dashed upward diagonal'), diagCross: N('Diagonal cross'), smCheck: N('Small checker board'),
    lgCheck: N('Large checker board'), smGrid: N('Small grid'), lgGrid: N('Large grid'), dotGrid: N('Dotted grid'), smConfetti: N('Small confetti'),
    lgConfetti: N('Large confetti'), horzBrick: N('Horizontal brick'), diagBrick: N('Diagonal brick'), solidDmnd: N('Solid diamond'), openDmnd: N('Outlined diamond'),
    dotDmnd: N('Dotted diamond'), plaid: N('Plaid'), sphere: N('Sphere'), weave: N('Weave'), divot: N('Divot'), shingle: N('Shingle'), wave: N('Wave'),
    trellis: N('Trellis'), zigZag: N('Zig zag') };
  /** SVG path of a pattern's set bits, one 1×1 square each, in an 8×8 box */
  const patternPath = (bits) => {
    let d = '';
    bits.forEach((row, y) => { for (let x = 0; x < 8; x++) if (row & (0x80 >> x)) d += `M${x},${y}h1v1h-1z`; });
    return d;
  };

  L.geom = {
    PATTERNS,
    PATTERN_NAMES,
    patternPath,
    presets: G,
    get: (name) => G[name] || G.rect,
    isLine: (name) => !!(G[name] && G[name].isLine),
    adj(sh) {
      const sp = spec(sh.geom);
      return Object.assign(sp ? Object.fromEntries((sp.av || []).map((a) => [a[0], a[2]])) : {}, sh.adj || {});
    },
    parts(sh, w, h) {
      w = Math.max(w, 0.01); h = Math.max(h, 0.01);
      if (sh.geom === 'custom' && sh.path) return customParts(sh.path, w, h);
      if (G[sh.geom] && G[sh.geom].isLine) return [];
      const sp = spec(sh.geom) || spec('rect');
      return sp ? specParts(sp, w, h, sh.adj) : [{ d: `M0,0 H${n(w)} V${n(h)} H0 Z`, fill: 'norm', stroke: true }];
    },
    textRect(sh) {
      if (sh.txRect) return sh.txRect;   // its own text area, [l, t, r, b] in the shape (SmartArt's txXfrm)
      const sp = spec(sh.geom);
      return sp ? specText(sp, sh.w, sh.h, sh.adj) : [0, 0, sh.w, sh.h];
    },
    warpLine,
    warp,
    warpSVG,
    handles(sh) {
      const sp = !(G[sh.geom] && G[sh.geom].isLine) && spec(sh.geom);
      if (!sp) return [];
      if (!handleCache.has(sh.geom)) handleCache.set(sh.geom, specHandles(sp));
      return handleCache.get(sh.geom);
    },
    linePath(sh) {
      const w = sh.w, h = sh.h;
      const x1 = sh.flipH ? w : 0, y1 = sh.flipV ? h : 0, x2 = sh.flipH ? 0 : w, y2 = sh.flipV ? 0 : h;
      if (sh.geom === 'bentConnector2') return `M${n(x1)},${n(y1)} H${n(x2)} V${n(y2)}`;
      if (sh.geom === 'bentConnector3' || sh.geom === 'bentConnector4') { const a = sh.adj && isFinite(sh.adj.adj1) ? sh.adj.adj1 / 100000 : 0.5; const mx = x1 + (x2 - x1) * a; return `M${n(x1)},${n(y1)} H${n(mx)} V${n(y2)} H${n(x2)}`; }
      if (sh.geom === 'curvedConnector2') return `M${n(x1)},${n(y1)} Q${n(x2)},${n(y1)} ${n(x2)},${n(y2)}`;
      if (sh.geom === 'curvedConnector3') { const mx = (x1 + x2) / 2; return `M${n(x1)},${n(y1)} C${n(mx)},${n(y1)} ${n(mx)},${n(y2)} ${n(x2)},${n(y2)}`; }
      return `M${n(x1)},${n(y1)} L${n(x2)},${n(y2)}`;
    },
    /** an SVG marker in `defs` for a line end ({ type, w, len }, DrawingML's headEnd / tailEnd), in the line's colour;
        lw is the line's width (heads stay legible on hairlines, as Office sizes them for at least ~2 pt); its url(#id) */
    marker(defs, end, color, lw, isHead) {
      if (!end || !end.type || end.type === 'none') return null;
      const svg = (tag, a) => { const e = document.createElementNS('http://www.w3.org/2000/svg', tag); for (const k in a) e.setAttribute(k, a[k]); return e; };
      const id = 'lnEnd' + (++markers), sz = { sm: 2, med: 3, lg: 5 };
      const k = Math.max(lw || 0.75, 2) / Math.max(lw || 0.75, 0.1);
      const m = svg('marker', { id, viewBox: '0 0 10 10', refX: end.type === 'arrow' ? 9 : end.type === 'oval' || end.type === 'diamond' ? 5 : 8, refY: 5, markerWidth: (sz[end.len || 'med'] || 3) * k, markerHeight: (sz[end.w || 'med'] || 3) * k, orient: isHead ? 'auto-start-reverse' : 'auto', markerUnits: 'strokeWidth' });
      switch (end.type) {
        case 'stealth': m.appendChild(svg('path', { d: 'M0,0 L10,5 L0,10 L3,5 Z', fill: color })); break;
        case 'diamond': m.appendChild(svg('path', { d: 'M0,5 L5,0 L10,5 L5,10 Z', fill: color })); break;
        case 'oval': m.appendChild(svg('circle', { cx: 5, cy: 5, r: 4.5, fill: color })); break;
        case 'arrow': m.appendChild(svg('path', { d: 'M1,1 L9,5 L1,9', fill: 'none', stroke: color, 'stroke-width': 1.6, 'stroke-linejoin': 'miter' })); break;
        default: m.appendChild(svg('path', { d: 'M0,0 L10,5 L0,10 Z', fill: color }));
      }
      defs.appendChild(m);
      return `url(#${id})`;
    },
    lineEnds(sh) {
      return [[sh.x + (sh.flipH ? sh.w : 0), sh.y + (sh.flipV ? sh.h : 0)], [sh.x + (sh.flipH ? 0 : sh.w), sh.y + (sh.flipV ? 0 : sh.h)]];
    },
    setLineEnds(sh, p1, p2) {
      sh.x = Math.min(p1[0], p2[0]); sh.y = Math.min(p1[1], p2[1]);
      sh.w = Math.abs(p2[0] - p1[0]); sh.h = Math.abs(p2[1] - p1[1]);
      sh.flipH = p1[0] > p2[0]; sh.flipV = p1[1] > p2[1];
    },
    /* AutoShapes menu groups (2003 order), by their English names (UI text, translated where shown) */
    MENU: {
      [N('Basic Shapes')]: ['rect', 'parallelogram', 'trapezoid', 'diamond', 'roundRect', 'octagon', 'triangle', 'rtTriangle', 'ellipse', 'hexagon', 'plus', 'pentagon', 'can', 'cube', 'bevel', 'foldedCorner', 'smileyFace', 'donut', 'noSmoking', 'heart', 'lightningBolt', 'sun', 'moon', 'arc', 'bracketPair', 'bracePair', 'leftBracket', 'rightBracket', 'leftBrace', 'rightBrace', 'frame', 'cloud'],
      [N('Block Arrows')]: ['rightArrow', 'leftArrow', 'upArrow', 'downArrow', 'leftRightArrow', 'upDownArrow', 'quadArrow', 'bentArrow', 'uturnArrow', 'circularArrow', 'stripedRightArrow', 'notchedRightArrow', 'homePlate', 'chevron'],
      [N('Flowchart')]: ['flowChartProcess', 'flowChartAlternateProcess', 'flowChartDecision', 'flowChartInputOutput', 'flowChartPredefinedProcess', 'flowChartInternalStorage', 'flowChartDocument', 'flowChartMultidocument', 'flowChartTerminator', 'flowChartPreparation', 'flowChartManualInput', 'flowChartManualOperation', 'flowChartConnector', 'flowChartOffpageConnector', 'flowChartPunchedCard', 'flowChartPunchedTape', 'flowChartSummingJunction', 'flowChartOr', 'flowChartCollate', 'flowChartSort', 'flowChartExtract', 'flowChartMerge', 'flowChartDelay', 'flowChartMagneticDisk', 'flowChartDisplay'],
      [N('Stars and Banners')]: ['irregularSeal1', 'irregularSeal2', 'star4', 'star5', 'star6', 'star8', 'star16', 'star24', 'star32', 'ribbon2', 'verticalScroll', 'horizontalScroll', 'wave'],
      [N('Callouts')]: ['wedgeRectCallout', 'wedgeRoundRectCallout', 'wedgeEllipseCallout', 'cloudCallout'],
      [N('Action Buttons')]: ['actionButtonBlank', 'actionButtonHome', 'actionButtonHelp', 'actionButtonInformation', 'actionButtonBackPrevious', 'actionButtonForwardNext', 'actionButtonBeginning', 'actionButtonEnd', 'actionButtonReturn', 'actionButtonDocument', 'actionButtonSound', 'actionButtonMovie'],
    },
    /* small SVG preview of a preset (for menus) */
    preview(name, size) {
      size = size || 18;
      const s = 14, o = (size - s) / 2;
      if (G[name] && G[name].isLine) {
        const d = name === 'bentConnector3' ? `M${o},${o} H${size / 2} V${o + s} H${o + s}` : name === 'curvedConnector3' ? `M${o},${o} C${size / 2},${o} ${size / 2},${o + s} ${o + s},${o + s}` : `M${o},${o + s} L${o + s},${o}`;
        return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}"><path d="${d}" fill="none" stroke="#1f3f7a" stroke-width="1.2"/></svg>`;
      }
      const parts = L.geom.parts({ geom: name }, s, s);
      const body = parts.map((p) => `<path d="${p.d}" transform="translate(${o},${o})" fill="${p.fill === 'none' ? 'none' : p.fill.startsWith('darken') ? '#5b7fc0' : p.fill.startsWith('lighten') ? '#dfe9fb' : '#a9c3ef'}" ${p.rule ? `fill-rule="${p.rule}"` : ''} stroke="${p.stroke ? '#1f3f7a' : 'none'}" stroke-width="0.8"/>`).join('');
      return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}">${body}</svg>`;
    },
  };
})(typeof window !== 'undefined' ? window : globalThis);
