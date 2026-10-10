/* VibeOffice — the SmartArt layout engine: a diagram's items (a tree of text) laid out as shapes by one of the
 * built-in layouts. Pure functions in points, with no DOM beyond an optional canvas for measuring text, so
 * the apps draw the result as their own shapes and tools/smartart.test.mjs runs it in Node.
 *
 *   items  [{ id, text, kids: [items] }]                    the outline the user types in the text pane
 *   SA.layout(id, items, box, opts) → { shapes, layout }   box { x, y, w, h } in points
 *   opts   { measure(text, sizePt) → widthPt, maxFont, minFont }
 *   shape  { role, geom, adj, x, y, w, h, rot, itemId, level, styleLbl, colorIdx, colorCnt,
 *            text: { lines: [{ t, lvl }], box: { x, y, w, h }, anchor, align }, fontSize, path }
 *
 * role: 'node' (an item's shape), 'child' (its sub-items as bullets), 'arrow', 'conn' (a line, path [[x, y]…]).
 * geom: an OOXML preset name (roundRect, rightArrow, chevron, ellipse, triangle, trapezoid, rect), 'path' (a line
 * through path [[x, y]…]) or 'poly' (a filled outline through poly [[x, y]…]).
 * adj: the preset's adjust values in OOXML units (adj 50000 = half), only where the default is not wanted.
 * styleLbl: the SmartArt style label (node1, sibTrans2D1, parChTrans1D2, vennNode1, …) that colours and
 * styles map from. The layouts and their names are our own reading of what Office's everyday layouts show;
 * nothing here is taken from Office's layout definitions.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const SA = (L.smartart = {});

  /* ---------------------------------------------------------------- text */
  let canvas = null;
  /** width in points of text at a size in points, in the theme's body font */
  SA.measureWith = (font) => (text, size) => {
    if (typeof document === 'undefined') return text.length * size * 0.5;   // Node: an average glyph
    canvas = canvas || document.createElement('canvas').getContext('2d');
    canvas.font = `${size}px ${font || 'Calibri, Carlito, sans-serif'}`;
    return canvas.measureText(text).width;
  };
  const LINE = 1.2;   // line height, × font size

  /** text broken into lines that fit width at size; null when a single word cannot fit */
  function wrap(lines, width, size, measure) {
    const out = [];
    for (const ln of lines) {
      const indent = ln.lvl ? size * 0.9 : 0;
      const words = String(ln.t || '').split(/\s+/).filter(Boolean);
      if (!words.length) { out.push({ t: '', lvl: ln.lvl }); continue; }
      let cur = '';
      for (const w of words) {
        if (measure(w, size) > width - indent) return null;
        const next = cur ? cur + ' ' + w : w;
        if (measure(next, size) <= width - indent) cur = next;
        else { out.push({ t: cur, lvl: ln.lvl }); cur = w; }
      }
      out.push({ t: cur, lvl: ln.lvl });
    }
    return out;
  }
  function fits(sh, size, measure) {
    const b = sh.text.box, inset = size * 0.25;
    const wrapped = wrap(sh.text.lines, b.w - 2 * inset, size, measure);
    return !!wrapped && wrapped.length * size * LINE <= b.h - 2 * inset + 0.01;
  }
  /* SmartArt's rule: every shape of a role shows its text at one size, the largest at which all fit */
  function fitText(shapes, opts) {
    const measure = opts.measure || SA.measureWith();
    const groups = new Map();
    for (const sh of shapes) if (sh.text) {
      const k = sh.role === 'child' ? 'child' : 'node';
      if (!groups.has(k)) groups.set(k, []);
      groups.get(k).push(sh);
    }
    for (const [k, list] of groups) {
      const max = k === 'child' ? (opts.maxChildFont || 24) : (opts.maxFont || 40), min = opts.minFont || 6;
      let lo = min, hi = max;
      if (list.every((sh) => fits(sh, hi, measure))) lo = hi;
      else for (let i = 0; i < 20; i++) {
        const mid = (lo + hi) / 2;
        if (list.every((sh) => fits(sh, mid, measure))) lo = mid; else hi = mid;
      }
      const size = Math.max(min, Math.floor(lo * 2) / 2);
      for (const sh of list) sh.fontSize = size;
    }
    /* sub-item bullets are never bigger than their items' text */
    const node = shapes.find((s) => s.text && s.role !== 'child');
    if (node) for (const sh of shapes) if (sh.role === 'child' && sh.fontSize > node.fontSize * 0.9) sh.fontSize = Math.floor(node.fontSize * 0.9 * 2) / 2;
  }

  /* ---------------------------------------------------------------- helpers */
  const linesOf = (it, withKids) => [{ t: it.text || '', lvl: 0 }].concat(withKids ? (it.kids || []).map((k) => ({ t: k.text || '', lvl: 1 })) : []);
  const kidLines = (it) => (it.kids || []).map((k) => ({ t: k.text || '', lvl: 1 }));
  const tb = (x, y, w, h) => ({ x, y, w, h });
  function node(it, i, n, geom, x, y, w, h, extra) {
    return Object.assign({ role: 'node', geom, x, y, w, h, rot: 0, itemId: it.id, level: 0, styleLbl: 'node1', colorIdx: i, colorCnt: n,
      text: { lines: linesOf(it, false), box: tb(x, y, w, h), anchor: 'ctr', align: 'ctr' } }, extra || {});
  }
  /** the largest w for items of aspect a (h = a·w) laid in a row of n with gaps g·w, inside box */
  const rowUnit = (box, n, a, g) => Math.min(box.w / (n + (n - 1) * g), box.h / a);
  const deg = (r) => r * 180 / Math.PI;

  /* ---------------------------------------------------------------- the layouts */
  const LAYOUTS = [];
  const def = (id, name, cat, desc, make, extra) => LAYOUTS.push(Object.assign({ id, name, cat, desc, make }, extra || {}));

  def('blockList', 'Basic Block List', 'list', 'Blocks of equal size, in rows; sub-items become bullets inside each block.', (items, box) => {
    const n = Math.max(1, items.length), a = 0.6, g = 0.1;
    let best = null;
    for (let c = 1; c <= n; c++) {
      const r = Math.ceil(n / c);
      const w = Math.min(box.w / (c + (c - 1) * g), box.h / (r * a + (r - 1) * g * a));
      if (!best || w > best.w + 0.01) best = { c, r, w };
    }
    const { c, r, w } = best, h = a * w, gx = g * w, gy = g * h;
    const top = box.y + (box.h - (r * h + (r - 1) * gy)) / 2;
    return items.map((it, i) => {
      const row = Math.floor(i / c), inRow = row === r - 1 ? n - row * c : c;
      const left = box.x + (box.w - (inRow * w + (inRow - 1) * gx)) / 2;
      const x = left + (i - row * c) * (w + gx), y = top + row * (h + gy);
      return node(it, i, n, 'rect', x, y, w, h, { text: { lines: linesOf(it, true), box: tb(x, y, w, h), anchor: 'ctr', align: 'ctr' } });
    });
  });

  def('verticalBullet', 'Vertical Bullet List', 'list', 'Items stacked in bars, each with its sub-items as a bulleted list below.', (items, box) => {
    const n = Math.max(1, items.length), kidH = 0.5, g = 0.12;
    const weight = items.reduce((s, it) => s + 1 + (it.kids && it.kids.length ? kidH * it.kids.length + 0.1 : 0), 0) + (n - 1) * g;
    const u = Math.min(box.h / weight, box.w * 0.3);
    let y = box.y + (box.h - u * weight) / 2;
    const out = [];
    items.forEach((it, i) => {
      out.push(node(it, i, n, 'roundRect', box.x, y, box.w, u, { text: { lines: linesOf(it, false), box: tb(box.x + u * 0.15, y, box.w - u * 0.3, u), anchor: 'ctr', align: 'l' } }));
      y += u;
      if (it.kids && it.kids.length) {
        const h = kidH * u * it.kids.length + 0.1 * u;
        out.push({ role: 'child', geom: 'rect', noFill: true, x: box.x, y, w: box.w, h, rot: 0, itemId: it.id, level: 1, styleLbl: 'revTx', colorIdx: i, colorCnt: n,
          text: { lines: kidLines(it), box: tb(box.x + u * 0.4, y, box.w - u * 0.5, h), anchor: 't', align: 'l' } });
        y += h;
      }
      y += g * u;
    });
    return out;
  }, { levels: 2 });

  def('process', 'Basic Process', 'process', 'Steps in a row, joined by arrows.', (items, box) => {
    const n = Math.max(1, items.length), a = 0.6, g = 0.4;
    const w = rowUnit(box, n, a, g), h = a * w;
    const left = box.x + (box.w - (n * w + (n - 1) * g * w)) / 2, y = box.y + (box.h - h) / 2;
    const out = [];
    items.forEach((it, i) => {
      const x = left + i * w * (1 + g);
      out.push(node(it, i, n, 'roundRect', x, y, w, h, { adj: [10000] }));
      if (i < n - 1) {
        const aw = w * 0.22, ah = w * 0.26;
        out.push({ role: 'arrow', geom: 'rightArrow', adj: [60000, 50000], x: x + w + (g * w - aw) / 2, y: y + (h - ah) / 2, w: aw, h: ah, rot: 0, itemId: it.id, level: 0, styleLbl: 'sibTrans2D1', colorIdx: i, colorCnt: n - 1 });
      }
    });
    return out;
  });

  def('chevron', 'Basic Chevron Process', 'process', 'Steps as chevrons pointing onward.', (items, box) => {
    const n = Math.max(1, items.length), a = 0.4, step = 0.84;
    const w = Math.min(box.w / (1 + (n - 1) * step), box.h / a), h = a * w;
    const left = box.x + (box.w - (w + (n - 1) * step * w)) / 2, y = box.y + (box.h - h) / 2;
    return items.map((it, i) => {
      const x = left + i * step * w;
      return node(it, i, n, 'chevron', x, y, w, h, { text: { lines: linesOf(it, false), box: tb(x + h * 0.5, y, w - h, h), anchor: 'ctr', align: 'ctr' } });
    });
  });

  def('cycle', 'Basic Cycle', 'cycle', 'Stages that repeat, in a circle with arrows between them.', (items, box) => {
    const n = Math.max(1, items.length), S = Math.min(box.w, box.h), cx = box.x + box.w / 2, cy = box.y + box.h / 2;
    if (n === 1) { const d = S; return [node(items[0], 0, 1, 'ellipse', cx - d / 2, cy - d / 2, d, d, { text: { lines: linesOf(items[0], false), box: tb(cx - d * 0.35, cy - d * 0.35, d * 0.7, d * 0.7), anchor: 'ctr', align: 'ctr' } })]; }
    const k = Math.min(1.05, 1.55 * Math.sin(Math.PI / n));   // node diameter ÷ circle radius
    const R = S / (2 + k), d = k * R;
    const out = [];
    const at = (i) => -Math.PI / 2 + (2 * Math.PI * i) / n;
    items.forEach((it, i) => {
      const t = at(i), x = cx + R * Math.cos(t) - d / 2, y = cy + R * Math.sin(t) - d / 2;
      out.push(node(it, i, n, 'ellipse', x, y, d, d, { text: { lines: linesOf(it, false), box: tb(x + d * 0.12, y + d * 0.15, d * 0.76, d * 0.7), anchor: 'ctr', align: 'ctr' } }));
    });
    items.forEach((it, i) => {
      const t = (at(i) + at(i + 1)) / 2, chord = 2 * R * Math.sin(Math.PI / n);
      const aw = Math.max(d * 0.15, (chord - d) * 0.6), ah = d * 0.22;
      const mx = cx + R * Math.cos(t), my = cy + R * Math.sin(t);
      out.push({ role: 'arrow', geom: 'rightArrow', adj: [60000, 50000], x: mx - aw / 2, y: my - ah / 2, w: aw, h: ah, rot: deg(t) + 90, itemId: it.id, level: 0, styleLbl: 'sibTrans2D1', colorIdx: i, colorCnt: n });
    });
    return out;
  });

  def('radial', 'Basic Radial', 'cycle', 'A central idea with the items that relate to it around it.', (items, box) => {
    if (!items.length) return [];
    const centre = items[0], around = centre.kids && centre.kids.length ? centre.kids : items.slice(1);
    const n = around.length, S = Math.min(box.w, box.h), cx = box.x + box.w / 2, cy = box.y + box.h / 2;
    if (!n) { const d = S * 0.5; return [node(centre, 0, 1, 'ellipse', cx - d / 2, cy - d / 2, d, d, { text: { lines: linesOf(centre, false), box: tb(cx - d * 0.35, cy - d * 0.35, d * 0.7, d * 0.7), anchor: 'ctr', align: 'ctr' } })]; }
    /* centre D, satellites d = 0.6·D at radius R = D/2 + gap + d/2, shrunk when they would touch */
    let D = S / 2.9, d = 0.8 * D, R = D / 2 + 0.2 * D + d / 2;
    const room = 2 * R * Math.sin(Math.PI / Math.max(n, 2)) * 0.85;
    if (d > room) { d = room; R = D / 2 + 0.2 * D + d / 2; }
    const scale = S / (2 * R + d);
    D *= scale; d *= scale; R *= scale;
    const out = [node(centre, 0, n + 1, 'ellipse', cx - D / 2, cy - D / 2, D, D, { text: { lines: linesOf(centre, false), box: tb(cx - D * 0.35, cy - D * 0.35, D * 0.7, D * 0.7), anchor: 'ctr', align: 'ctr' } })];
    around.forEach((it, i) => {
      const t = -Math.PI / 2 + (2 * Math.PI * i) / n, sx = cx + R * Math.cos(t), sy = cy + R * Math.sin(t);
      const p0 = [cx + (D / 2) * Math.cos(t), cy + (D / 2) * Math.sin(t)], p1 = [sx - (d / 2) * Math.cos(t), sy - (d / 2) * Math.sin(t)];
      out.push({ role: 'conn', geom: 'path', path: [p0, p1], x: Math.min(p0[0], p1[0]), y: Math.min(p0[1], p1[1]), w: Math.abs(p1[0] - p0[0]), h: Math.abs(p1[1] - p0[1]), rot: 0, itemId: it.id, level: 1, styleLbl: 'parChTrans1D2', colorIdx: i, colorCnt: n });
      out.push(node(it, i + 1, n + 1, 'ellipse', sx - d / 2, sy - d / 2, d, d, { level: 1, text: { lines: linesOf(it, false), box: tb(sx - d * 0.35, sy - d * 0.35, d * 0.7, d * 0.7), anchor: 'ctr', align: 'ctr' } }));
    });
    /* in a wide space the circles become ellipses: everything stretched across (lines stay joined) */
    const k = Math.min(2, box.w / (2 * R + d));
    if (k > 1) for (const sh of out) {
      const sx = (x) => cx + (x - cx) * k;
      sh.x = sx(sh.x); sh.w *= k;
      if (sh.path) sh.path = sh.path.map(([x, y]) => [sx(x), y]);
      if (sh.text) { sh.text.box.x = sx(sh.text.box.x); sh.text.box.w *= k; }
    }
    return out;
  }, { levels: 2 });

  /* a tree: each item centred over (or, across, beside) its sub-items; connectors bent or straight */
  function tree(items, box, o) {
    if (!items.length) return [];
    const H = !!o.across, gapX = 0.25, a = o.aspect || 0.55, gapY = o.gapY || 0.4;
    const width = (it) => (it.kids && it.kids.length ? Math.max(1, it.kids.reduce((s, k) => s + width(k), 0) + (it.kids.length - 1) * gapX) : 1);
    const depth = (it) => 1 + (it.kids && it.kids.length ? Math.max(...it.kids.map(depth)) : 0);
    const Wu = items.reduce((s, it) => s + width(it), 0) + (items.length - 1) * gapX, D = Math.max(...items.map(depth));
    /* along: the spread of siblings; down: the levels. Across trees spread down the page, levels go right */
    const spread = H ? box.h : box.w, levels = H ? box.w : box.h;
    const u = H ? Math.min(spread / (Wu * a), levels / (D + (D - 1) * gapY)) : Math.min(spread / Wu, levels / (D * a + (D - 1) * gapY));
    const nw = u, nh = a * u;                       // node size, in its own orientation (w along levels when across)
    const sib = H ? nh : nw, lev = H ? nw : nh;     // a node's extent along siblings and along levels
    const step = sib * (1 + gapX) - sib * gapX + sib * gapX;   // = sib·(1+gapX)
    const gap = (H ? nw : nh) * 0 + gapY * (H ? nw : u);
    const start0 = (H ? box.y : box.x) + (spread - (Wu * sib + 0)) / 2, top = (H ? box.x : box.y) + (levels - (D * lev + (D - 1) * gap)) / 2;
    const out = [];
    let count = 0;
    const at = (along, down, w, h) => (H ? { x: down, y: along, w, h } : { x: along, y: down, w, h });
    const pt = (along, down) => (H ? [down, along] : [along, down]);
    const place = (it, left, lvl) => {
      const wu = width(it), a0 = left + (wu * sib - sib) / 2, d0 = top + lvl * (lev + gap);
      const r = at(a0, d0, H ? nw : nw, H ? nh : nh);
      out.push(node(it, count++, 0, o.geom || 'rect', r.x, r.y, r.w, r.h, { level: lvl, adj: o.adj }));
      if (!it.kids || !it.kids.length) return;
      let cl = left + (wu - (it.kids.reduce((s, k) => s + width(k), 0) + (it.kids.length - 1) * gapX)) * sib / 2;
      const pa = a0 + sib / 2, pd = d0 + lev, mid = pd + gap / 2;
      for (const k of it.kids) {
        const ka = cl + (width(k) * sib - sib) / 2 + sib / 2;
        const path = o.straight ? [pt(pa, pd), pt(ka, pd + gap)] : [pt(pa, pd), pt(pa, mid), pt(ka, mid), pt(ka, pd + gap)];
        const xs = path.map((q) => q[0]), ys = path.map((q) => q[1]);
        out.push({ role: 'conn', geom: 'path', path, x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys), rot: 0, itemId: k.id, level: lvl + 1, styleLbl: 'parChTrans1D2', colorIdx: 0, colorCnt: 1 });
        place(k, cl, lvl + 1);
        cl += (width(k) + gapX) * sib;
      }
    };
    void step;
    let left = start0;
    for (const it of items) { place(it, left, 0); left += (width(it) + gapX) * sib; }
    for (const sh of out) if (sh.role === 'node') { sh.colorCnt = count; if (o.byLevel) { sh.colorIdx = sh.level; sh.colorCnt = D; } }
    return out;
  }
  def('orgChart', 'Organization Chart', 'hierarchy', 'People or units and who reports to whom, top down.', (items, box) => tree(items, box, {}), { levels: 9, tree: true });
  def('hierarchy', 'Hierarchy', 'hierarchy', 'Levels from the top down, each joined straight to the one above.', (items, box) => tree(items, box, { geom: 'roundRect', adj: [10000], straight: true, byLevel: true, aspect: 0.6 }), { levels: 9, tree: true });
  def('hHierarchy', 'Horizontal Hierarchy', 'hierarchy', 'A tree that grows from left to right.', (items, box) => tree(items, box, { across: true, geom: 'roundRect', adj: [10000], aspect: 0.5, gapY: 0.35 }), { levels: 9, tree: true });

  def('pyramid', 'Basic Pyramid', 'pyramid', 'Levels that build on one another, broadest at the bottom.', (items, box) => {
    const n = Math.max(1, items.length), a = 0.87;   // height ÷ base
    const base = Math.min(box.w, box.h / a), H = base * a, gap = n > 1 ? H * 0.015 : 0;
    /* the apex is taller than the other layers, so its text need not be much smaller than theirs */
    const wt = (i) => (i === 0 && n > 1 ? 1.7 : 1), unit = (H - (n - 1) * gap) / items.reduce((s, _, i) => s + wt(i), 0) || H;
    const x0 = box.x + (box.w - base) / 2, y0 = box.y + (box.h - H) / 2;
    let yAt = y0;
    return items.map((it, i) => {
      const lh = unit * wt(i), yTop = yAt, wTop = base * (yTop - y0) / H, wBot = base * (yTop + lh - y0) / H;
      yAt += lh + gap;
      const x = x0 + (base - wBot) / 2;
      const inset = (wBot - wTop) / 2;
      const geom = i === 0 ? 'triangle' : 'trapezoid';
      const adj = i === 0 ? [50000] : [Math.round(inset / Math.min(wBot, lh) * 100000)];
      /* text sits in the layer's middle band; the apex uses the triangle's own text area (the middle half of its
         lower half), which applications that ignore the stored text box use too */
      const tw = i === 0 ? wBot * 0.5 : wTop + inset * 0.8, th = i === 0 ? lh * 0.5 : lh;
      return node(it, i, n, geom, x, yTop, wBot, lh, { adj, text: { lines: linesOf(it, false), box: tb(x + (wBot - tw) / 2, yTop + lh - th, tw, th), anchor: 'ctr', align: 'ctr' } });
    });
  });

  def('venn', 'Basic Venn', 'relationship', 'Overlapping circles for ideas that share something.', (items, box) => {
    const n = Math.max(1, items.length), S = Math.min(box.w, box.h), cx = box.x + box.w / 2, cy = box.y + box.h / 2;
    const k = n === 1 ? 0 : n === 2 ? 0.35 : 0.375 / Math.sin(Math.PI / n);   // centre distance ÷ diameter
    const d = S / (1 + 2 * k), r = k * d;
    return items.map((it, i) => {
      const t = n === 2 ? (i ? 0 : Math.PI) : -Math.PI / 2 + (2 * Math.PI * i) / n;
      const ex = cx + r * Math.cos(t), ey = cy + r * Math.sin(t), x = ex - d / 2, y = ey - d / 2;
      const off = n === 1 ? 0 : d * 0.18;
      return node(it, i, n, 'ellipse', x, y, d, d, { styleLbl: 'vennNode1', alpha: 0.5,
        text: { lines: linesOf(it, false), box: tb(ex + off * Math.cos(t) - d * 0.3, ey + off * Math.sin(t) - d * 0.22, d * 0.6, d * 0.44), anchor: 'ctr', align: 'ctr' } });
    });
  });

  def('target', 'Basic Target', 'relationship', 'Rings inside rings, from the whole to the core.', (items, box) => {
    const n = Math.max(1, items.length), S = Math.min(box.w, box.h), cx = box.x + box.w / 2, cy = box.y + box.h / 2;
    return items.map((it, i) => {
      const d = S * (n - i) / n, inner = S * (n - i - 1) / n;
      const x = cx - d / 2, y = cy - d / 2;
      const band = i === n - 1 ? tb(cx - d * 0.35, cy - d * 0.25, d * 0.7, d * 0.5) : tb(cx - d * 0.3, y + d * 0.02, d * 0.6, (d - inner) / 2 * 0.96);
      return node(it, i, n, 'ellipse', x, y, d, d, { text: { lines: linesOf(it, false), box: band, anchor: 'ctr', align: 'ctr' } });
    });
  });

  /* ---------------------------------------------------------------- more lists */
  const child = (it, i, n, x, y, w, h, inset, extra) => Object.assign({ role: 'child', geom: 'rect', tint: true, x, y, w, h, rot: 0, itemId: it.id, level: 1, styleLbl: 'alignAcc1', colorIdx: i, colorCnt: n,
    text: { lines: kidLines(it), box: tb(x + inset, y, w - inset * 1.2, h), anchor: 't', align: 'l' } }, extra || {});
  const hasKids = (it) => !!(it.kids && it.kids.length);
  def('stackedList', 'Stacked List', 'list', 'Each item in a circle, its sub-items on a card beside it.', (items, box) => {
    const n = Math.max(1, items.length), g = 0.15, rh = Math.min(box.h / (n + (n - 1) * g), box.w * 0.3);
    let y = box.y + (box.h - (n * rh + (n - 1) * g * rh)) / 2;
    const out = [];
    items.forEach((it, i) => {
      const d = rh, cx = box.x;
      out.push(child(it, i, n, cx + d * 0.5, y + rh * 0.12, box.w - d * 0.5, rh * 0.76, d * 0.6));
      out.push(node(it, i, n, 'ellipse', cx, y, d, d, { text: { lines: linesOf(it, false), box: tb(cx + d * 0.146, y + d * 0.146, d * 0.708, d * 0.708), anchor: 'ctr', align: 'ctr' } }));
      y += rh * (1 + g);
    });
    return out;
  }, { levels: 2 });
  def('hBullet', 'Horizontal Bullet List', 'list', 'Columns, each with a heading and its sub-items as bullets.', (items, box) => {
    const n = Math.max(1, items.length), g = 0.06, cw = box.w / (n + (n - 1) * g), hh = Math.min(box.h * 0.3, cw * 0.6);
    const out = [];
    items.forEach((it, i) => {
      const x = box.x + i * cw * (1 + g);
      out.push(child(it, i, n, x, box.y + hh, cw, box.h - hh, cw * 0.08, { text: { lines: kidLines(it), box: tb(x + cw * 0.08, box.y + hh + cw * 0.04, cw * 0.86, box.h - hh - cw * 0.04), anchor: 't', align: 'l' } }));
      out.push(node(it, i, n, 'rect', x, box.y, cw, hh));
    });
    return out;
  }, { levels: 2 });
  def('linedList', 'Lined List', 'list', 'Items under rules, with their sub-items to the right.', (items, box) => {
    const n = Math.max(1, items.length), rh = box.h / n, out = [];
    items.forEach((it, i) => {
      const y = box.y + i * rh;
      out.push({ role: 'conn', geom: 'path', path: [[box.x, y], [box.x + box.w, y]], x: box.x, y, w: box.w, h: 0, rot: 0, itemId: it.id, level: 0, styleLbl: 'parChTrans1D2', colorIdx: i, colorCnt: n });
      out.push(node(it, i, n, 'rect', box.x, y, box.w * 0.3, rh, { noFill: true, text: { lines: linesOf(it, false), box: tb(box.x, y, box.w * 0.3, rh), anchor: 't', align: 'l' } }));
      if (hasKids(it)) out.push(child(it, i, n, box.x + box.w * 0.32, y, box.w * 0.68, rh, 0, { tint: false, noFill: true }));
    });
    return out;
  }, { levels: 2 });
  def('vBox', 'Vertical Box List', 'list', 'A labelled box for each item, holding its sub-items.', (items, box) => {
    const n = Math.max(1, items.length), tabH = 0.38, g = 0.1;
    const u = box.h / (n * (1 + tabH * 0.5) + (n - 1) * g);
    let y = box.y + (box.h - u * (n * (1 + tabH * 0.5) + (n - 1) * g)) / 2;
    const out = [];
    items.forEach((it, i) => {
      const th = u * tabH, by = y + th * 0.5;
      out.push(child(it, i, n, box.x, by, box.w, u, box.w * 0.05, { line: true, text: { lines: kidLines(it), box: tb(box.x + box.w * 0.05, by + th * 0.55, box.w * 0.9, u - th * 0.55), anchor: 't', align: 'l' } }));
      out.push(node(it, i, n, 'roundRect', box.x + box.w * 0.04, y, box.w * 0.62, th, { adj: [16667], text: { lines: linesOf(it, false), box: tb(box.x + box.w * 0.06, y, box.w * 0.58, th), anchor: 'ctr', align: 'l' } }));
      y += u * (1 + tabH * 0.5 + g);
    });
    return out;
  }, { levels: 2 });

  /* ---------------------------------------------------------------- more processes */
  const arrowAt = (it, i, n, cx, cy, len, thick, rot) => ({ role: 'arrow', geom: 'rightArrow', adj: [60000, 50000], x: cx - len / 2, y: cy - thick / 2, w: len, h: thick, rot: rot || 0, itemId: it.id, level: 0, styleLbl: 'sibTrans2D1', colorIdx: i, colorCnt: n });
  def('stepUp', 'Step Up Process', 'process', 'Steps that climb, each a little higher than the last.', (items, box) => {
    const n = Math.max(1, items.length), g = 0.1, w = box.w / (n + (n - 1) * g), rise = 0.55;
    const h = Math.min(box.h / (1 + (n - 1) * rise), w * 0.9);
    const y0 = box.y + (box.h - h * (1 + (n - 1) * rise)) / 2 + h * (n - 1) * rise;
    return items.map((it, i) => node(it, i, n, 'roundRect', box.x + i * w * (1 + g), y0 - i * h * rise, w, h, { adj: [10000] }));
  });
  def('continuousArrow', 'Continuous Arrow Process', 'process', 'Steps along one arrow, for a flow that keeps going.', (items, box) => {
    const n = Math.max(1, items.length), ah = Math.min(box.h, box.w * 0.32), ay = box.y + (box.h - ah) / 2;
    const out = [{ role: 'arrow', big: true, geom: 'rightArrow', adj: [70000, 50000], x: box.x, y: ay, w: box.w, h: ah, rot: 0, itemId: items[0] ? items[0].id : null, level: 0, styleLbl: 'sibTrans2D1', colorIdx: 0, colorCnt: 1 }];
    const span = box.w - ah * 0.6, g = 0.12, w = span / (n + (n - 1) * g), h = ah * 0.5;
    items.forEach((it, i) => out.push(node(it, i, n, 'roundRect', box.x + i * w * (1 + g), ay + (ah - h) / 2, w, h, { adj: [16667] })));
    return out;
  });
  def('vProcess', 'Vertical Process', 'process', 'Steps from top to bottom, joined by arrows.', (items, box) => {
    const n = Math.max(1, items.length), g = 0.5, h = box.h / (n + (n - 1) * g), w = Math.min(box.w, h * 7), x = box.x + (box.w - w) / 2;
    const out = [];
    items.forEach((it, i) => {
      const y = box.y + i * h * (1 + g);
      out.push(node(it, i, n, 'roundRect', x, y, w, h, { adj: [10000] }));
      if (i < n - 1) out.push(arrowAt(it, i, n - 1, box.x + box.w / 2, y + h + g * h / 2, g * h * 0.8, Math.min(h * 0.45, g * h * 0.8 * 1.2), 90));
    });
    return out;
  });
  def('timeline', 'Basic Timeline', 'process', 'Events along a line, in order.', (items, box) => {
    const n = Math.max(1, items.length), cy = box.y + box.h / 2, lineH = Math.min(box.h * 0.12, 24);
    const out = [{ role: 'arrow', big: true, geom: 'rightArrow', adj: [50000, 60000], x: box.x, y: cy - lineH / 2, w: box.w, h: lineH, rot: 0, itemId: items[0] ? items[0].id : null, level: 0, styleLbl: 'sibTrans2D1', colorIdx: 0, colorCnt: 1 }];
    const step = (box.w - lineH * 1.5) / n, d = Math.min(lineH * 1.1, step * 0.4), lh = box.h / 2 - d;
    items.forEach((it, i) => {
      const mx = box.x + step * (i + 0.5);
      out.push({ role: 'marker', geom: 'ellipse', x: mx - d / 2, y: cy - d / 2, w: d, h: d, rot: 0, itemId: it.id, level: 0, styleLbl: 'node1', colorIdx: i, colorCnt: n });
      const up = i % 2 === 0, y = up ? box.y : cy + d / 2;
      const lw = Math.min(step * 1.2, box.w), lx = Math.max(box.x, Math.min(mx - lw / 2, box.x + box.w - lw));
      out.push(node(it, i, n, 'rect', lx, y, lw, lh, { noFill: true, text: { lines: linesOf(it, false), box: tb(lx, y, lw, lh), anchor: up ? 'b' : 't', align: 'ctr' } }));
    });
    return out;
  });

  /* ---------------------------------------------------------------- more cycles and relationships */
  def('continuousCycle', 'Continuous Cycle', 'cycle', 'Stages around a ring, for something that never stops.', (items, box) => {
    const n = Math.max(1, items.length), S = Math.min(box.w, box.h), cx = box.x + box.w / 2, cy = box.y + box.h / 2;
    const k = Math.min(0.9, 1.6 * Math.sin(Math.PI / Math.max(n, 2)));
    const nw = (S / (2 + k * 0.6)) * k, nh = nw * 0.6, R = (S - nh) / 2;
    const ring = Array.from({ length: 73 }, (_, j) => { const t = (2 * Math.PI * j) / 72; return [cx + R * Math.cos(t), cy + R * Math.sin(t)]; });
    const out = [{ role: 'ring', geom: 'path', path: ring, x: cx - R, y: cy - R, w: 2 * R, h: 2 * R, rot: 0, itemId: items[0] ? items[0].id : null, level: 0, styleLbl: 'sibTrans2D1', colorIdx: 0, colorCnt: 1 }];
    items.forEach((it, i) => {
      const t = -Math.PI / 2 + (2 * Math.PI * i) / n, x = cx + R * Math.cos(t) - nw / 2, y = cy + R * Math.sin(t) - nh / 2;
      out.push(node(it, i, n, 'roundRect', Math.max(box.x, Math.min(x, box.x + box.w - nw)), y, nw, nh, { adj: [16667] }));
    });
    return out;
  });
  def('linearVenn', 'Linear Venn', 'relationship', 'Overlapping circles in a row, for ideas that follow on.', (items, box) => {
    const n = Math.max(1, items.length), ov = 0.75, d = Math.min(box.h, box.w / (1 + (n - 1) * ov));
    const x0 = box.x + (box.w - d * (1 + (n - 1) * ov)) / 2, y = box.y + (box.h - d) / 2;
    return items.map((it, i) => {
      const x = x0 + i * d * ov, mid = i === 0 ? 0.1 : i === n - 1 ? 0.3 : 0.2;
      return node(it, i, n, 'ellipse', x, y, d, d, { styleLbl: 'vennNode1', alpha: 0.5, text: { lines: linesOf(it, false), box: tb(x + d * mid, y + d * 0.3, d * 0.6, d * 0.4), anchor: 'ctr', align: 'ctr' } });
    });
  });
  def('funnel', 'Funnel', 'relationship', 'Ideas that narrow down through layers to one result at the bottom.', (items, box) => {
    if (!items.length) return [];
    const n = items.length, layers = n > 1 ? items.slice(0, -1) : [], last = items[n - 1];
    const W = Math.min(box.w, box.h * 1.15), x0 = box.x + (box.w - W) / 2, d = W * 0.3;
    const fh = box.h - d * 1.05, lh = layers.length ? fh / layers.length : 0, spout = W * 0.22;
    const out = layers.map((it, i) => {
      const wTop = W - (W - spout) * i / layers.length, wBot = W - (W - spout) * (i + 1) / layers.length;
      const x = box.x + (box.w - wTop) / 2, y = box.y + i * lh, inset = (wTop - wBot) / 2;
      const hh = lh * 0.94, g = poly([[x, y], [x + wTop, y], [x + wTop - inset, y + hh], [x + inset, y + hh]]);
      return node(it, i, n, 'poly', g.x, g.y, g.w, g.h, { poly: g.poly, text: { lines: linesOf(it, false), box: tb(x + inset * 0.5, y, wBot + inset, hh), anchor: 'ctr', align: 'ctr' } });
    });
    const cx = x0 + W / 2, y = box.y + box.h - d;
    out.push(node(last, n - 1, n, 'ellipse', cx - d / 2, y, d, d, { text: { lines: linesOf(last, false), box: tb(cx - d * 0.354, y + d * 0.146, d * 0.708, d * 0.708), anchor: 'ctr', align: 'ctr' } }));
    return out;
  });
  def('matrix', 'Basic Matrix', 'matrix', 'Items in a grid of equal quarters, for comparing them side by side.', (items, box) => {
    const n = Math.max(1, items.length), c = Math.ceil(Math.sqrt(n)), r = Math.ceil(n / c), g = 0.05;
    const s = Math.min(box.w / (c + (c - 1) * g), box.h / (r + (r - 1) * g)) , w = Math.min(box.w / (c + (c - 1) * g), s * 1.6);
    const x0 = box.x + (box.w - (c * w + (c - 1) * g * s)) / 2, y0 = box.y + (box.h - (r * s + (r - 1) * g * s)) / 2;
    return items.map((it, i) => node(it, i, n, 'roundRect', x0 + (i % c) * (w + g * s), y0 + Math.floor(i / c) * s * (1 + g), w, s, { adj: [8000] }));
  });

  /* ---------------------------------------------------------------- more pyramids */
  /* drawn as outlines (geom 'poly', points in the slide), so their text stays upright everywhere */
  const poly = (pts) => ({ geom: 'poly', poly: pts, x: Math.min(...pts.map((q) => q[0])), y: Math.min(...pts.map((q) => q[1])), w: Math.max(...pts.map((q) => q[0])) - Math.min(...pts.map((q) => q[0])), h: Math.max(...pts.map((q) => q[1])) - Math.min(...pts.map((q) => q[1])) });
  def('invertedPyramid', 'Inverted Pyramid', 'pyramid', 'Levels from the broadest at the top to the narrowest at the bottom.', (items, box) => {
    const n = Math.max(1, items.length), a = 0.87;
    const base = Math.min(box.w, box.h / a), H = base * a, gap = n > 1 ? H * 0.015 : 0;
    const wt = (i) => (i === n - 1 && n > 1 ? 1.7 : 1), unit = (H - (n - 1) * gap) / items.reduce((s, _, i) => s + wt(i), 0) || H;
    const cx = box.x + box.w / 2, y0 = box.y + (box.h - H) / 2;
    const wAt = (y) => base * (1 - (y - y0) / H);
    let yAt = y0;
    return items.map((it, i) => {
      const lh = unit * wt(i), yTop = yAt, yBot = yTop + lh, wT = wAt(yTop), wB = wAt(yBot);
      yAt += lh + gap;
      const last = i === n - 1;
      const g = poly(last ? [[cx - wT / 2, yTop], [cx + wT / 2, yTop], [cx, yBot]] : [[cx - wT / 2, yTop], [cx + wT / 2, yTop], [cx + wB / 2, yBot], [cx - wB / 2, yBot]]);
      const tw = last ? wT * 0.5 : wB + (wT - wB) * 0.4, th = last ? lh * 0.5 : lh;
      return node(it, i, n, 'poly', g.x, g.y, g.w, g.h, { poly: g.poly, text: { lines: linesOf(it, false), box: tb(cx - tw / 2, yTop, tw, th), anchor: 'ctr', align: 'ctr' } });
    });
  });
  def('pyramidList', 'Pyramid List', 'pyramid', 'Items on cards laid over a pyramid, from the top down.', (items, box) => {
    const n = Math.max(1, items.length), a = 0.87, base = Math.min(box.w, box.h / a), H = base * a, x0 = box.x + (box.w - base) / 2, y0 = box.y + (box.h - H) / 2;
    const out = [{ role: 'bg', geom: 'triangle', adj: [50000], x: x0, y: y0, w: base, h: H, rot: 0, itemId: items[0] ? items[0].id : null, level: 0, styleLbl: 'bgShp', colorIdx: 0, colorCnt: 1 }];
    const top = y0 + H * 0.28, avail = H * 0.7, g = 0.15, h = avail / (n + (n - 1) * g), w = base * 0.5;
    items.forEach((it, i) => out.push(node(it, i, n, 'roundRect', x0 + base / 2 - w / 2, top + i * h * (1 + g), w, h, { adj: [16667] })));
    return out;
  });

  /* gallery order: by category, then as defined */
  const CAT_ORDER = ['list', 'process', 'cycle', 'hierarchy', 'relationship', 'matrix', 'pyramid'];
  LAYOUTS.sort((p, q) => CAT_ORDER.indexOf(p.cat) - CAT_ORDER.indexOf(q.cat));
  SA.LAYOUTS = LAYOUTS;
  SA.CATEGORIES = [['list', 'List'], ['process', 'Process'], ['cycle', 'Cycle'], ['hierarchy', 'Hierarchy'], ['relationship', 'Relationship'], ['matrix', 'Matrix'], ['pyramid', 'Pyramid']];
  SA.get = (id) => LAYOUTS.find((l) => l.id === id) || null;

  /** the shapes of a diagram: layout id, items, box in points */
  SA.layout = (id, items, box, opts) => {
    opts = opts || {};
    const lay = SA.get(id) || LAYOUTS[0];
    const shapes = lay.make(items || [], box);
    fitText(shapes, opts);
    return { layout: lay.id, shapes };
  };

  /** new items for a fresh diagram of a layout: empty text, shown as [Text] while editing */
  let seq = 0;
  SA.newId = () => 'i' + Date.now().toString(36) + (seq++).toString(36);
  SA.blank = (id, n) => {
    const it = () => ({ id: SA.newId(), text: '', kids: [] });
    if ((SA.get(id) || {}).tree) { const top = it(); top.kids = [it(), it(), it()]; return [top]; }
    if (id === 'radial') { const c = it(); c.kids = [it(), it(), it(), it()]; return [c]; }
    return Array.from({ length: n || (id === 'venn' || id === 'target' ? 3 : id === 'cycle' ? 5 : 3) }, it);
  };
  /** flat outline (text pane order): [{ item, level, parent }] */
  SA.outline = (items) => {
    const out = [];
    const walk = (list, level, parent) => list.forEach((it) => { out.push({ item: it, level, parent }); walk(it.kids || [], level + 1, it); });
    walk(items, 0, null);
    return out;
  };
})(typeof window !== 'undefined' ? window : globalThis);
