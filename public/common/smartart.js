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
 * geom: an OOXML preset name (roundRect, rightArrow, chevron, ellipse, triangle, trapezoid, rect) or 'path'.
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
    let D = S / 2.9, d = 0.6 * D, R = D / 2 + 0.25 * D + d / 2;
    const room = 2 * R * Math.sin(Math.PI / Math.max(n, 2)) * 0.85;
    if (d > room) { d = room; R = D / 2 + 0.25 * D + d / 2; }
    const scale = S / (2 * R + d);
    D *= scale; d *= scale; R *= scale;
    const out = [node(centre, 0, n + 1, 'ellipse', cx - D / 2, cy - D / 2, D, D, { text: { lines: linesOf(centre, false), box: tb(cx - D * 0.35, cy - D * 0.35, D * 0.7, D * 0.7), anchor: 'ctr', align: 'ctr' } })];
    around.forEach((it, i) => {
      const t = -Math.PI / 2 + (2 * Math.PI * i) / n, sx = cx + R * Math.cos(t), sy = cy + R * Math.sin(t);
      const p0 = [cx + (D / 2) * Math.cos(t), cy + (D / 2) * Math.sin(t)], p1 = [sx - (d / 2) * Math.cos(t), sy - (d / 2) * Math.sin(t)];
      out.push({ role: 'conn', geom: 'path', path: [p0, p1], x: Math.min(p0[0], p1[0]), y: Math.min(p0[1], p1[1]), w: Math.abs(p1[0] - p0[0]), h: Math.abs(p1[1] - p0[1]), rot: 0, itemId: it.id, level: 1, styleLbl: 'parChTrans1D2', colorIdx: i, colorCnt: n });
      out.push(node(it, i + 1, n + 1, 'ellipse', sx - d / 2, sy - d / 2, d, d, { level: 1, text: { lines: linesOf(it, false), box: tb(sx - d * 0.35, sy - d * 0.35, d * 0.7, d * 0.7), anchor: 'ctr', align: 'ctr' } }));
    });
    return out;
  }, { levels: 2 });

  def('orgChart', 'Organization Chart', 'hierarchy', 'People or units and who reports to whom, top down.', (items, box) => {
    if (!items.length) return [];
    const gapX = 0.25, a = 0.55, gapY = 0.4;
    const width = (it) => (it.kids && it.kids.length ? Math.max(1, it.kids.reduce((s, k) => s + width(k), 0) + (it.kids.length - 1) * gapX) : 1);
    const depth = (it) => 1 + (it.kids && it.kids.length ? Math.max(...it.kids.map(depth)) : 0);
    const Wu = items.reduce((s, it) => s + width(it), 0) + (items.length - 1) * gapX, D = Math.max(...items.map(depth));
    const u = Math.min(box.w / Wu, box.h / (D * a + (D - 1) * gapY)), h = a * u;
    const left0 = box.x + (box.w - Wu * u) / 2, top = box.y + (box.h - (D * h + (D - 1) * gapY * u)) / 2;
    const out = [];
    let count = 0;
    const place = (it, left, lvl) => {
      const wu = width(it), x = left + (wu * u - u) / 2, y = top + lvl * (h + gapY * u);
      out.push(node(it, count++, 0, 'rect', x, y, u, h, { level: lvl }));
      if (!it.kids || !it.kids.length) return;
      let cl = left + (wu - (it.kids.reduce((s, k) => s + width(k), 0) + (it.kids.length - 1) * gapX)) * u / 2;
      const midY = y + h + gapY * u / 2;
      for (const k of it.kids) {
        const kx = cl + (width(k) * u - u) / 2 + u / 2, px = x + u / 2;
        const path = [[px, y + h], [px, midY], [kx, midY], [kx, midY + gapY * u / 2]];
        out.push({ role: 'conn', geom: 'path', path, x: Math.min(px, kx), y: y + h, w: Math.abs(kx - px), h: gapY * u, rot: 0, itemId: k.id, level: lvl + 1, styleLbl: 'parChTrans1D2', colorIdx: 0, colorCnt: 1 });
        place(k, cl, lvl + 1);
        cl += (width(k) + gapX) * u;
      }
    };
    let left = left0;
    for (const it of items) { place(it, left, 0); left += (width(it) + gapX) * u; }
    for (const sh of out) if (sh.role === 'node') sh.colorCnt = count;
    return out;
  }, { levels: 9 });

  def('pyramid', 'Basic Pyramid', 'pyramid', 'Levels that build on one another, broadest at the bottom.', (items, box) => {
    const n = Math.max(1, items.length), a = 0.87;   // height ÷ base
    const base = Math.min(box.w, box.h / a), H = base * a, gap = n > 1 ? H * 0.015 : 0;
    const lh = (H - (n - 1) * gap) / n, x0 = box.x + (box.w - base) / 2, y0 = box.y + (box.h - H) / 2;
    return items.map((it, i) => {
      const yTop = y0 + i * (lh + gap), wTop = base * (yTop - y0) / H, wBot = base * (yTop + lh - y0) / H;
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
    if (id === 'orgChart') { const top = it(); top.kids = [it(), it(), it()]; return [top]; }
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
