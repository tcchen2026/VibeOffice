/* Lectern — SmartArt diagrams. A diagram is a group with an `sa` record:
 *   sa = { layout, colors, style, items: [{ id, text, kids }] }     (common/smartart.js lays the items out)
 * Its kids are the laid-out shapes, regenerated from the items on every change, each marked
 * kid.sa = { item, role } so a click, an on-slide text edit or Delete finds its item. Because it stays a group,
 * moving, copying, undo, drafts and drawing work as for any group; resizing lays it out again
 * (editor.js) and SmartArt Tools replace Drawing Tools while it is selected (app.js A.inContext).
 * Items are plain text; colours and styles name theme slots, so a new design recolours the diagram.
 */
(function () {
  'use strict';
  const L = window.L, ui = L.ui, { h } = L;
  const SA = L.smartart, M = L.model, T = L.txt;
  const E = () => L.ed;
  const DG = (L.diagram = {});

  /* ---------------------------------------------------------------- colours and styles */
  DG.COLORS = [['accent1', 'Colored Fill - Accent 1'], ['accent2', 'Colored Fill - Accent 2'], ['accent3', 'Colored Fill - Accent 3'], ['accent4', 'Colored Fill - Accent 4'],
    ['accent5', 'Colored Fill - Accent 5'], ['accent6', 'Colored Fill - Accent 6'], ['colorful', 'Colorful - Accent Colors'], ['range', 'Gradient Range - Accent 1']];
  DG.STYLES = [['simple', 'Simple Fill'], ['white', 'White Outline'], ['subtle', 'Subtle Effect'], ['moderate', 'Moderate Effect'], ['intense', 'Intense Effect']];
  const COLORFUL = ['accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

  /** the theme's usable colours: accents too close to the background are skipped (some themes have a white
      accent), and text is dark on light fills; the saved colours part follows the same choice (smartart-io.js) */
  DG.palette = (d) => {
    const lum = (slot) => L.color.luma(M.resolveColor(slot, d));
    const bg = lum('lt1'), ok = (slot) => Math.abs(lum(slot) - bg) > 0.15;
    const first = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].find(ok) || 'accent1';
    const colorful = COLORFUL.filter(ok);
    return { slot: (slot) => (ok(slot) ? slot : first), colorful: colorful.length ? colorful : [first], text: (slot) => (lum(slot) > 0.62 ? 'dk1' : 'lt1') };
  };
  /** fill, line, shadow and text colour of one laid-out shape */
  function paint(s, sa, pal) {
    const scheme = sa.colors || 'accent1', style = sa.style || 'simple';
    let c = /^accent[1-6]$/.test(scheme) ? pal.slot(scheme) : scheme === 'colorful' ? pal.colorful[(s.colorIdx || 0) % pal.colorful.length] : pal.slot('accent1');
    let a = s.alpha == null ? 1 : s.alpha;
    if (scheme === 'range' && s.colorCnt > 1) a *= 1 - 0.6 * (s.colorIdx || 0) / (s.colorCnt - 1);
    if (scheme === 'colorful' && s.role === 'conn') c = pal.slot('accent1');
    /* sub-item cards: a pale tint of the item's colour, with an outline when the layout asks for one */
    if (s.role === 'child' && s.tint && !s.noFill) return { fill: { t: 'solid', c, a: 0.18 }, line: s.line ? { c, w: 1, dash: 'solid' } : { t: 'none' }, shadow: null, text: 'tx1' };
    if (s.role === 'child' || s.noFill) return { fill: { t: 'none' }, line: { t: 'none' }, shadow: null, text: 'tx1' };
    if (s.role === 'conn') return { fill: { t: 'none' }, line: { c, w: 1.5, dash: 'solid' }, shadow: null };
    if (s.role === 'ring') return { fill: { t: 'none' }, line: { c, a: 0.4, w: 8, dash: 'solid' }, shadow: null };
    if (s.role === 'bg') return { fill: { t: 'solid', c, a: 0.25 }, line: { t: 'none' }, shadow: null };
    if (s.role === 'marker') return { fill: { t: 'solid', c, a: 1 }, line: { c: 'lt1', w: 1.5, dash: 'solid' }, shadow: null };
    if (s.role === 'arrow') return { fill: { t: 'solid', c, a: a * (s.big ? 0.3 : 0.45) }, line: { t: 'none' }, shadow: null };
    const solid = { t: 'solid', c, a };
    const grad = (top) => ({ t: 'grad', stops: [{ p: 0, c, a: a * top }, { p: 1, c, a }], ang: 90, path: 'lin' });
    const shadow = (o, dy, blur) => ({ c: '#000000', a: o, dx: 0, dy, blur });
    const text = s.styleLbl === 'vennNode1' ? 'tx1' : pal.text(c);
    switch (style) {
      case 'white': return { fill: solid, line: { c: 'lt1', w: 2.5, dash: 'solid' }, shadow: null, text };
      case 'subtle': return { fill: solid, line: { c: 'lt1', w: 1, dash: 'solid' }, shadow: shadow(0.3, 1.5, 3), text };
      case 'moderate': return { fill: grad(0.72), line: { t: 'none' }, shadow: shadow(0.35, 2, 4), text };
      case 'intense': return { fill: grad(0.55), line: { t: 'none' }, shadow: shadow(0.45, 3, 7), text };
      default: return { fill: solid, line: { c: 'lt1', w: 1, dash: 'solid' }, shadow: null, text };
    }
  }

  /* ---------------------------------------------------------------- items → shapes */
  const design = () => M.design(L.pres, E().slide());
  /* 15% to spare, for applications that draw the text in a wider font or with wider margins (LibreOffice) */
  const measurer = (d) => { const m = SA.measureWith(L.fontStack(L.style.font('+mn', d))); return (t, sz) => m(t, sz) * 1.15; };
  DG.isDiagram = (sh) => !!(sh && sh.type === 'group' && sh.sa);
  /** the diagram a shape is, or is part of */
  DG.of = (sh) => { if (!sh) return null; if (DG.isDiagram(sh)) return sh; const top = E().topOf(sh.id); return DG.isDiagram(top) ? top : null; };
  DG.current = () => DG.of(E().primary());

  /* formatting the user gave one shape (fill, outline, font colour, bold, italic) is kept on its item, as
     item.fmt[role], so it survives laying the diagram out again; Reset Graphic clears it */
  const FMT_ROLES = ['node', 'child'];
  const fmtOf = (sa, id, role) => { const r = id && FMT_ROLES.includes(role) && locate(sa.items, id); return (r && r.item.fmt && r.item.fmt[role]) || null; };
  function kidOf(s, sa, keep, pal) {
    const p = Object.assign({}, paint(s, sa, pal)), f = fmtOf(sa, s.itemId, s.role) || {};
    if (f.fill) p.fill = L.clone(f.fill);
    if (f.line) p.line = L.clone(f.line);
    if (f.color) p.text = f.color;
    const key = `${s.itemId}:${s.role}:${keep.n[s.itemId + s.role] = (keep.n[s.itemId + s.role] || 0) + 1}`;
    const base = { id: keep.ids[key] || L.uid('s'), type: 'shape', name: (SA.get(sa.layout) || {}).name + ' ' + s.role, x: s.x, y: s.y, w: Math.max(s.w, 0.01), h: Math.max(s.h, 0.01), rot: s.rot || 0, fill: p.fill, line: p.line, shadow: p.shadow, sa: { item: s.itemId, role: s.role } };
    if (s.flipV) base.flipV = true;
    if (FMT_ROLES.includes(s.role)) base.sa.gen = { fill: JSON.stringify(base.fill), line: JSON.stringify(base.line), color: p.text, b: !!f.b, i: !!f.i };
    if (s.geom === 'path') {
      const cmds = s.path.map(([x, y], i) => [i ? 'L' : 'M', x - s.x, y - s.y]);
      return Object.assign(base, { geom: 'custom', path: { paths: [{ w: base.w, h: base.h, cmds, fill: 'none', stroke: true }] } });
    }
    if (s.geom === 'poly') {
      const cmds = s.poly.map(([x, y], i) => [i ? 'L' : 'M', x - s.x, y - s.y]).concat([['Z']]);
      Object.assign(base, { geom: 'custom', path: { paths: [{ w: base.w, h: base.h, cmds, fill: 'norm', stroke: true }] } });
    } else Object.assign(base, { geom: s.geom });
    const defAdj = (L.geom.get(s.geom) || {}).adj;
    if (defAdj) {
      base.adj = L.clone(defAdj);
      const names = Object.keys(defAdj);
      (s.adj || []).forEach((v, i) => { if (names[i]) base.adj[names[i]] = v; });
    }
    if (s.text) {
      const sz = s.fontSize, b = s.text.box, inset = sz * 0.25;
      const rp = Object.assign({ sz, color: p.text }, f.b ? { b: true } : {}, f.i ? { i: true } : {});
      const ps = s.text.lines.map((ln) => (ln.lvl
        ? T.para(ln.t, { algn: 'l', marL: sz * 0.9, indent: -sz * 0.9, bu: { t: 'char', ch: '•' } }, rp, 1)
        : T.para(ln.t, { algn: s.text.align === 'l' ? 'l' : 'ctr', bu: { t: 'none' } }, rp)));
      base.tx = T.body(ps, { anchor: s.text.anchor === 't' || s.text.anchor === 'b' ? s.text.anchor : 'ctr', wrap: true, autofit: 'none', ins: [inset, inset, inset, inset] });
      base.txRect = [b.x - s.x, b.y - s.y, b.x + b.w - s.x, b.y + b.h - s.y].map((v) => L.round(v, 3));
    }
    return base;
  }
  /** lay the diagram out again in its box; shapes keep their ids where their item and role are the same */
  /** what the user changed on the shapes since they were laid out, recorded on their items */
  DG.capture = (g) => {
    for (const k of g.kids || []) {
      const gen = k.sa && k.sa.gen, r = gen && locate(g.sa.items, k.sa.item);
      if (!r) continue;
      const f = {};
      if (JSON.stringify(k.fill) !== gen.fill) f.fill = L.clone(k.fill);
      if (JSON.stringify(k.line) !== gen.line) f.line = L.clone(k.line);
      const run = k.tx && k.tx.ps[0] && (k.tx.ps[0].rs[0] || k.tx.ps[0].end);
      if (run) {
        if (run.color && run.color !== gen.color) f.color = run.color;
        if (!!run.b !== gen.b) f.b = !!run.b;
        if (!!run.i !== gen.i) f.i = !!run.i;
      }
      if (!Object.keys(f).length) continue;
      const fmt = (r.item.fmt = r.item.fmt || {});
      fmt[k.sa.role] = Object.assign(fmt[k.sa.role] || {}, f);
      for (const key of Object.keys(fmt[k.sa.role])) if (fmt[k.sa.role][key] === false) delete fmt[k.sa.role][key];
    }
  };
  DG.relayout = (g, d) => {
    d = d || design();
    DG.capture(g);
    const keep = { ids: {}, n: {} }, seen = {};
    for (const k of g.kids || []) if (k.sa) { const kk = k.sa.item + k.sa.role; seen[kk] = (seen[kk] || 0) + 1; keep.ids[`${k.sa.item}:${k.sa.role}:${seen[kk]}`] = k.id; }
    const res = SA.layout(g.sa.layout, g.sa.items, { x: g.x, y: g.y, w: g.w, h: g.h }, { measure: measurer(d) });
    const pal = DG.palette(d);
    g.kids = res.shapes.map((s) => kidOf(s, g.sa, keep, pal));
    return g;
  };
  /** after a change: undo point, lay out, draw, keep the text pane in step */
  function change(g, label, fn, coalesce) {
    L.hist.push(label, coalesce);
    fn();
    DG.relayout(g);
    E().renderShape(g.id);
    E().touched();
    refreshPane();
  }

  /* ---------------------------------------------------------------- inserting */
  DG.insert = (layout, phId) => {
    const slide = E().slide();
    if (!slide) return;
    /* like PowerPoint, an empty content placeholder on the slide takes the diagram */
    const ph = phId ? E().shape(phId) : slide.shapes.find((s) => s.ph && ['obj', 'body'].includes(s.ph.type || 'body') && T.isEmpty(s.tx));
    const box = ph ? { x: ph.x, y: ph.y, w: ph.w, h: ph.h } : { x: L.pres.W * 0.15, y: L.pres.H * 0.22, w: L.pres.W * 0.7, h: L.pres.H * 0.66 };
    L.hist.push('Insert SmartArt');
    const lay = SA.get(layout) || SA.LAYOUTS[0];
    const g = { id: L.uid('s'), type: 'group', name: lay.name + ' ' + (slide.shapes.length + 1), rot: 0, x: box.x, y: box.y, w: box.w, h: box.h,
      sa: { layout: lay.id, colors: 'accent1', style: 'simple', items: SA.blank(lay.id) }, kids: [] };
    DG.relayout(g);
    if (ph) slide.shapes.splice(slide.shapes.indexOf(ph), 1, g); else slide.shapes.push(g);
    E().select([g.id]);
    E().render(); E().touched();
    DG.showPane(true);
    /* with the ribbon, SmartArt Tools ▸ Design opens, as in PowerPoint */
    if (L.ribbonUI && L.ribbonUI.built && L.ribbonUI.built()) { ui.refresh(); setTimeout(() => L.ribbonUI.select('smartartDesign'), 60); }
  };

  /** items from outline lines: '\t' before a line makes it a sub-item of the line above */
  DG.itemsFrom = (lines) => {
    const out = [], stack = [];
    for (const ln of lines) {
      const lvl = /^\t*/.exec(ln)[0].length, it = { id: SA.newId(), text: ln.slice(lvl), kids: [] };
      stack.length = Math.min(stack.length, lvl);
      (stack.length ? stack[stack.length - 1].kids : out).push(it);
      stack.push(it);
    }
    return out;
  };
  /** a finished diagram for a deck being built (templates, the sample): in place of a placeholder, or in box */
  DG.make = (pres, slide, at, layout, lines, opts) => {
    const box = { x: at.x, y: at.y, w: at.w, h: at.h };
    const lay = SA.get(layout) || SA.LAYOUTS[0];
    const g = Object.assign({ id: L.uid('s'), type: 'group', name: lay.name + ' ' + (slide.shapes.length + 1), rot: 0 }, box,
      { sa: Object.assign({ layout: lay.id, colors: 'accent1', style: 'simple', items: DG.itemsFrom(lines) }, opts || {}), kids: [] });
    DG.relayout(g, M.design(pres, slide));
    const i = slide.shapes.indexOf(at);
    if (i >= 0) slide.shapes.splice(i, 1, g); else slide.shapes.push(g);
    return g;
  };

  /* ---------------------------------------------------------------- the gallery */
  /** a small SVG picture of a layout with sample items, in the design's accent colours */
  DG.thumb = (layout, w, hh, opts) => {
    opts = opts || {};
    const d = opts.design || design();
    const items = opts.items || sampleItems(layout);
    const { shapes } = SA.layout(layout, items, { x: 2, y: 2, w: w - 4, h: hh - 4 }, { measure: (t, sz) => t.length * sz * 0.5, maxFont: 9, minFont: 4 });
    const sa = { colors: opts.colors || 'accent1', style: 'simple' };
    const col = (c) => M.resolveColor(c, d), pal = DG.palette(d);
    let out = '';
    for (const s of shapes) {
      const p = paint(s, sa, pal);
      const fill = p.fill.t === 'solid' ? `fill="${col(p.fill.c)}" fill-opacity="${L.round(p.fill.a, 2)}"` : p.fill.t === 'grad' ? `fill="${col(p.fill.stops[1].c)}"` : 'fill="none"';
      const stroke = p.line.t === 'none' ? '' : ` stroke="${col(p.line.c)}" stroke-width="${s.role === 'conn' ? 1 : 0.5}"`;
      const tr = s.rot ? ` transform="rotate(${L.round(s.rot, 1)} ${L.round(s.x + s.w / 2, 1)} ${L.round(s.y + s.h / 2, 1)})"` : '';
      if (s.geom === 'path') out += `<polyline points="${s.path.map((q) => q.map((v) => L.round(v, 1)).join(',')).join(' ')}" fill="none"${stroke}/>`;
      else if (s.noFill) out += s.text ? s.text.lines.map((ln, i) => `<rect x="${L.round(s.text.box.x + 3, 1)}" y="${L.round(s.text.box.y + 2 + i * 4, 1)}" width="${L.round(s.text.box.w * 0.5, 1)}" height="1.6" fill="#999"/>`).join('') : '';
      else out += `<path d="${shapePath(s)}" ${fill}${stroke}${s.flipV ? ` transform="translate(0 ${L.round(2 * s.y + s.h, 1)}) scale(1 -1)"` : tr}/>`;
    }
    return `<svg width="${w}" height="${hh}" viewBox="0 0 ${w} ${hh}" aria-hidden="true">${out}</svg>`;
  };
  function shapePath(s) {
    const { x, y, w, h: hh } = s, a = (s.adj || [])[0], r = (v) => L.round(v, 1);
    const P = (pts) => 'M' + pts.map((q) => q.map(r).join(',')).join('L') + 'Z';
    switch (s.geom) {
      case 'poly': return P(s.poly);
      case 'ellipse': return `M${r(x)},${r(y + hh / 2)}a${r(w / 2)},${r(hh / 2)} 0 1 0 ${r(w)},0a${r(w / 2)},${r(hh / 2)} 0 1 0 ${r(-w)},0Z`;
      case 'triangle': return P([[x + w / 2, y], [x + w, y + hh], [x, y + hh]]);
      case 'trapezoid': { const i = ((a == null ? 25000 : a) / 100000) * Math.min(w, hh); return P([[x + i, y], [x + w - i, y], [x + w, y + hh], [x, y + hh]]); }
      case 'chevron': { const dd = hh / 2; return P([[x, y], [x + w - dd, y], [x + w, y + hh / 2], [x + w - dd, y + hh], [x, y + hh], [x + dd, y + hh / 2]]); }
      case 'rightArrow': { const hw = w - hh * 0.5; return P([[x, y + hh * 0.2], [x + hw, y + hh * 0.2], [x + hw, y], [x + w, y + hh / 2], [x + hw, y + hh], [x + hw, y + hh * 0.8], [x, y + hh * 0.8]]); }
      default: return P([[x, y], [x + w, y], [x + w, y + hh], [x, y + hh]]);
    }
  }
  function sampleItems(layout) {
    const it = (t, kids) => ({ id: t, text: '', kids: kids || [] });
    const lay = SA.get(layout) || {};
    if (lay.tree) return [it('a', [it('b', layout === 'orgChart' ? [] : [it('b1'), it('b2')]), it('c'), it('d')])];
    if (layout === 'radial') return [it('a', [it('b'), it('c'), it('d'), it('e')])];
    if (lay.levels === 2) return [it('a', [it('a1'), it('a2')]), it('b', [it('b1')]), it('c', [it('c1')])].slice(0, layout === 'verticalBullet' ? 2 : 3);
    return SA.blank(layout).map((x, i) => it('n' + i));
  }

  /** Insert ▸ SmartArt (and Change Layout): categories, pictures of the layouts, a description */
  DG.gallery = (cb, current) => {
    let pick = current || SA.LAYOUTS[0].id, cat = 'all';
    const cats = h('div', { class: 'sa-cats list-box', role: 'listbox', 'aria-label': 'Categories' });
    const grid = h('div', { class: 'sa-grid', role: 'listbox', 'aria-label': 'Layouts' });
    const big = h('div', { class: 'sa-big' }), title = h('b'), desc = h('div', { class: 'sa-desc' });
    const show = () => { const lay = SA.get(pick); big.innerHTML = DG.thumb(pick, 190, 150); title.textContent = lay.name; desc.textContent = lay.desc; L.$$('button', grid).forEach((b) => b.classList.toggle('on', b.dataset.lay === pick)); };
    const fill = () => {
      L.clear(grid);
      const shown = SA.LAYOUTS.filter((l) => cat === 'all' || l.cat === cat);
      if (shown.length && !shown.some((l) => l.id === pick)) pick = shown[0].id;
      for (const lay of shown) {
        const b = h('button', { type: 'button', class: 'sa-cell', 'data-lay': lay.id, 'aria-label': lay.name, 'data-tip': lay.name, html: DG.thumb(lay.id, 74, 58) });
        b.addEventListener('click', () => { pick = lay.id; show(); });
        b.addEventListener('dblclick', () => { pick = lay.id; d.buttons[0].click(); });
        grid.append(b);
      }
      show();
    };
    const used = new Set(SA.LAYOUTS.map((l) => l.cat));
    for (const [k, n] of [['all', 'All']].concat(SA.CATEGORIES.filter(([k]) => used.has(k)))) {
      const b = h('button', { type: 'button', class: 'sa-cat' + (k === cat ? ' on' : ''), text: n });
      b.addEventListener('click', () => { cat = k; L.$$('.sa-cat', cats).forEach((x) => x.classList.toggle('on', x === b)); fill(); });
      cats.append(b);
    }
    const d = ui.dialog({ title: current ? 'Change SmartArt Layout' : 'Choose a SmartArt Graphic', width: 720,
      body: h('div', { class: 'sa-gallery' }, cats, grid, h('div', { class: 'sa-side' }, big, title, desc)),
      buttons: [{ label: 'OK', primary: true, onClick: () => cb(pick) }, { label: 'Cancel' }] });
    fill();
  };

  /* ---------------------------------------------------------------- editing the items */
  function locate(items, id, parent) {
    for (let i = 0; i < items.length; i++) {
      if (items[i].id === id) return { list: items, i, item: items[i], parent: parent || null };
      const r = locate(items[i].kids || [], id, items[i]);
      if (r) return r;
    }
    return null;
  }
  DG.locate = (g, id) => locate(g.sa.items, id);
  const newItem = (text) => ({ id: SA.newId(), text: text || '', kids: [] });
  /** the item the user is on: the text pane's row, or the selected shape's item */
  let paneItem = null;
  function focusItem(g) {
    if (paneItem && DG.locate(g, paneItem)) return paneItem;
    const p = E().primary();
    if (p && p.sa && p.sa.item && DG.locate(g, p.sa.item)) return p.sa.item;
    const out = SA.outline(g.sa.items);
    return out.length ? out[out.length - 1].item.id : null;
  }
  const ops = {
    after(g, id) { const r = DG.locate(g, id); const it = newItem(); if (r) r.list.splice(r.i + 1, 0, it); else g.sa.items.push(it); return it.id; },
    before(g, id) { const r = DG.locate(g, id); const it = newItem(); if (r) r.list.splice(r.i, 0, it); else g.sa.items.unshift(it); return it.id; },
    above(g, id) { const r = DG.locate(g, id); if (!r) return ops.after(g, id); const it = newItem(); it.kids = [r.item]; r.list.splice(r.i, 1, it); return it.id; },
    below(g, id) { const r = DG.locate(g, id); if (!r) return ops.after(g, id); const it = newItem(); (r.item.kids = r.item.kids || []).push(it); return it.id; },
    promote(g, id) {
      const r = DG.locate(g, id); if (!r || !r.parent) return id;
      const up = DG.locate(g, r.parent.id);
      r.list.splice(r.i, 1);
      up.list.splice(up.i + 1, 0, r.item);
      return id;
    },
    demote(g, id) {
      const r = DG.locate(g, id); if (!r || r.i === 0) return id;
      r.list.splice(r.i, 1);
      (r.list[r.i - 1].kids = r.list[r.i - 1].kids || []).push(r.item);
      return id;
    },
    up(g, id) { const r = DG.locate(g, id); if (r && r.i > 0) [r.list[r.i - 1], r.list[r.i]] = [r.list[r.i], r.list[r.i - 1]]; return id; },
    down(g, id) { const r = DG.locate(g, id); if (r && r.i < r.list.length - 1) [r.list[r.i + 1], r.list[r.i]] = [r.list[r.i], r.list[r.i + 1]]; return id; },
    remove(g, id) {
      const r = DG.locate(g, id); if (!r) return null;
      r.list.splice(r.i, 1, ...(r.item.kids || []));
      const out = SA.outline(g.sa.items);
      return out.length ? out[Math.min(out.length - 1, Math.max(0, r.i - 1))].item.id : null;
    },
  };
  const LABELS = { after: 'Add Shape', before: 'Add Shape', above: 'Add Shape', below: 'Add Shape', promote: 'Promote', demote: 'Demote', up: 'Move Up', down: 'Move Down', remove: 'Delete Shape' };
  /** run an item operation on the current diagram; the pane follows the item it leaves you on */
  DG.op = (name, g, id) => {
    g = g || DG.current();
    if (!g) return;
    id = id || focusItem(g);
    let next = id;
    change(g, LABELS[name], () => { next = ops[name](g, id); });
    paneItem = next;
    selectItem(g, next);
    refreshPane(true);
  };
  function selectItem(g, id) {
    const k = g.kids.find((x) => x.sa && x.sa.item === id && x.sa.role === 'node');
    if (k && E().sel[0] !== g.id) E().select([k.id]);
  }
  DG.setText = (g, id, text) => {
    const r = DG.locate(g, id);
    if (!r || r.item.text === text) return;
    change(g, 'Typing', () => { r.item.text = text; }, 'sa-type-' + id);
  };
  DG.set = (key, val, g) => {
    g = g || DG.current();
    if (!g || g.sa[key] === val) return;
    const label = { layout: 'Change Layout', colors: 'Change Colors', style: 'SmartArt Style' }[key];
    change(g, label, () => { g.sa[key] = val; if (key === 'layout') g.name = (SA.get(val) || {}).name + ' ' + g.name.split(' ').pop(); });
  };
  DG.reset = (g) => {
    g = g || DG.current();
    if (!g) return;
    change(g, 'Reset Graphic', () => {
      g.sa.colors = 'accent1'; g.sa.style = 'simple';
      for (const k of g.kids) if (k.sa) delete k.sa.gen;   // nothing to capture: the shapes start again
      for (const o of SA.outline(g.sa.items)) delete o.item.fmt;
    });
  };
  /** Convert to Shapes: the same shapes, as an ordinary group */
  DG.toShapes = (g) => {
    g = g || DG.current();
    if (!g) return;
    E().commit('Convert to Shapes', () => { delete g.sa; for (const k of g.kids) { delete k.sa; } g.name = 'Group ' + g.name.split(' ').pop(); });
    refreshPane();
  };
  /** Delete on shapes of one diagram removes their items (their sub-items move up); true when handled */
  DG.deleteKids = (shapes) => {
    const g = shapes.length && DG.of(shapes[0]);
    if (!g || DG.isDiagram(shapes[0]) || !shapes.every((s) => s.sa && s.sa.item && DG.of(s) === g)) return false;
    const ids = [...new Set(shapes.map((s) => s.sa.item))];
    change(g, 'Delete Shape', () => { for (const id of ids) ops.remove(g, id); });
    E().select([g.id]);
    return true;
  };
  /* an edit on the slide goes back to the item: its first paragraph is the item, the rest its sub-items */
  let editing = null;
  L.bus.on('edit-start', () => { editing = L.te && L.te.state ? L.te.state.id : null; });
  L.bus.on('edit-end', () => {
    const id = editing;
    editing = null;
    const k = id && E().shape(id), g = k && k.sa && DG.of(k);
    if (!g) return;
    const r = DG.locate(g, k.sa.item);
    if (!r) return;
    const paras = (k.tx ? k.tx.ps : []).map((p) => ({ t: T.paraText(p), lvl: p.lvl || 0 }));
    const top = paras.filter((p) => !p.lvl).map((p) => p.t);
    const subs = paras.filter((p) => p.lvl).map((p) => p.t);
    L.hist.push('Typing');
    if (k.sa.role === 'child') r.item.kids = subs.concat(top).map((t, i) => Object.assign(r.item.kids[i] || newItem(), { text: t }));
    else {
      r.item.text = top.join(' ');
      if (subs.length || ((SA.get(g.sa.layout) || {}).id === 'blockList' && r.item.kids.length)) r.item.kids = subs.map((t, i) => Object.assign(r.item.kids[i] || newItem(), { text: t }));
    }
    DG.relayout(g);
    E().renderShape(g.id);
    E().touched();
    refreshPane();
  });

  /* a diagram read from a file shows the shapes stored in it; the first time it is selected they are laid out
     here, so that each shape knows its item (clicks, on-slide typing, Delete) */
  L.bus.on('selection', () => {
    const g = DG.current();
    if (g && g.kids.some((k) => !k.sa)) { DG.relayout(g); E().renderShape(g.id); }
  });

  /* ---------------------------------------------------------------- the text pane */
  let paneG = null;
  function rowsOf(g) { return SA.outline(g.sa.items); }
  function renderPane(b) {
    L.clear(b);
    const g = DG.current();
    paneG = g ? g.id : null;
    if (!g) { b.append(h('div', { class: 'tp-note', text: 'Select a SmartArt graphic to type its text here.' }), h('div', { class: 'tp-row' }, ui.button('Insert SmartArt...', () => ui.exec('insertDiagram'), { class: 'btn' }))); return; }
    b.append(h('div', { class: 'tp-h', text: 'Type your text here' }));
    const list = h('div', { class: 'sa-pane', role: 'tree', 'aria-label': 'SmartArt text' });
    for (const row of rowsOf(g)) {
      const inp = h('input', { type: 'text', class: 'sa-row', value: row.item.text, placeholder: '[Text]', 'aria-label': 'Item, level ' + (row.level + 1), 'data-item': row.item.id, style: `margin-left:${row.level * 16}px` });
      inp.addEventListener('focus', () => { paneItem = row.item.id; selectItem(g, row.item.id); L.hist.breakCoalesce(); });
      inp.addEventListener('input', () => DG.setText(g, row.item.id, inp.value));
      inp.addEventListener('keydown', (e) => paneKey(e, g, row, inp));
      list.append(h('div', { class: 'sa-line', role: 'treeitem' }, h('span', { class: 'sa-bul', text: '•', style: `margin-left:${row.level * 16}px` }), inp));
    }
    b.append(list);
    b.append(h('div', { class: 'tp-note', text: 'Enter adds an item, Tab and Shift+Tab move it in and out a level, Alt+Shift+↑/↓ move it.' }));
    b.append(h('div', { class: 'tp-row' }, ui.button('Change Layout...', () => ui.exec('diagramLayout'), { class: 'btn small' })));
    const want = paneItem && L.$(`.sa-row[data-item="${paneItem}"]`, list);
    if (want && paneFocus) { want.focus(); want.setSelectionRange(want.value.length, want.value.length); }
  }
  let paneFocus = false;
  function paneKey(e, g, row, inp) {
    e.stopPropagation();
    const go = (name) => { e.preventDefault(); paneFocus = true; DG.op(name, g, row.item.id); paneFocus = false; };
    if (e.key === 'Enter') { e.preventDefault(); paneFocus = true; DG.op('after', g, row.item.id); paneFocus = false; }
    else if (e.key === 'Tab') go(e.shiftKey ? 'promote' : 'demote');
    else if (e.altKey && e.shiftKey && e.key === 'ArrowUp') go('up');
    else if (e.altKey && e.shiftKey && e.key === 'ArrowDown') go('down');
    else if (e.key === 'Backspace' && !inp.value && rowsOf(g).length > 1) go('remove');
    else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      e.preventDefault();
      const all = L.$$('.sa-row', inp.closest('.sa-pane'));
      const n = all[all.indexOf(inp) + (e.key === 'ArrowUp' ? -1 : 1)];
      if (n) n.focus();
    } else if (e.key === 'Escape') { e.preventDefault(); inp.blur(); E().select([g.id]); }
  }
  function refreshPane(force) {
    const TP = L.panes && L.panes.task;
    if (!TP || TP.current !== 'diagramText' || !TP.el || TP.el.hidden) return;
    const b = TP.body;
    const focused = b.contains(document.activeElement);
    if (focused && !force) {
      /* typing: only the values change, the rows stay */
      const g = DG.current();
      if (g && g.id === paneG) return;
    }
    paneFocus = paneFocus || focused;
    renderPane(b);
    paneFocus = false;
  }
  DG.showPane = (on) => { const TP = L.panes && L.panes.task; if (!TP) return; if (on === false) { if (TP.current === 'diagramText') L.app.toggleTask(false); return; } TP.show('diagramText'); };
  if (L.panes && L.panes.task) L.panes.task.register({ id: 'diagramText', title: 'SmartArt Text', render(b) { renderPane(b); }, refresh(b, reason) { if (reason === 'selection') { const g = DG.current(); if (g && g.id === paneG && b.contains(document.activeElement)) return; } renderPane(b); } });

  /* ---------------------------------------------------------------- menus for the toolbar and the ribbon */
  DG.layoutMenu = (r) => {
    const g = DG.current();
    ui.openMenu([{ custom: (close) => {
      const grid = h('div', { class: 'sa-menu-grid' });
      for (const lay of SA.LAYOUTS) {
        const b = h('button', { type: 'button', class: 'sa-cell' + (g && g.sa.layout === lay.id ? ' on' : ''), 'aria-label': lay.name, 'data-tip': lay.name, html: DG.thumb(lay.id, 64, 50) });
        b.addEventListener('click', () => { close(); DG.set('layout', lay.id); });
        grid.append(b);
      }
      return grid;
    } }, '-', 'diagramLayout'], r);
  };
  DG.colorsMenu = (r) => {
    const g = DG.current();
    ui.openMenu(DG.COLORS.map(([k, n]) => ({ label: n, html: DG.thumb('process', 22, 12, { colors: k, items: [1, 2, 3].map((i) => ({ id: 'c' + i, text: '', kids: [] })) }), checked: () => !!g && g.sa.colors === k, run: () => DG.set('colors', k) })), r);
  };
  DG.styleMenu = (r) => {
    const g = DG.current();
    ui.openMenu(DG.STYLES.map(([k, n]) => ({ label: n, checked: () => !!g && g.sa.style === k, run: () => DG.set('style', k) })), r);
  };

  /* ---------------------------------------------------------------- commands */
  const C = (id, def) => ui.cmd(id, def);
  const inDiagram = () => !!DG.current() && E().view !== 'master';
  C('diagramAddShape', { label: 'Add Shape', icon: 'addShape', tip: 'Add Shape After', enabled: inDiagram, run: () => DG.op('after') });
  C('diagramAddAfter', { label: 'Add Shape &After', icon: 'addShape', enabled: inDiagram, run: () => DG.op('after') });
  C('diagramAddBefore', { label: 'Add Shape &Before', enabled: inDiagram, run: () => DG.op('before') });
  C('diagramAddAbove', { label: 'Add Shape Ab&ove', enabled: inDiagram, run: () => DG.op('above') });
  C('diagramAddBelow', { label: 'Add Shape Be&low', enabled: inDiagram, run: () => DG.op('below') });
  C('diagramPromote', { label: 'Promote', icon: 'promote', enabled: inDiagram, run: () => DG.op('promote') });
  C('diagramDemote', { label: 'Demote', icon: 'demote', enabled: inDiagram, run: () => DG.op('demote') });
  C('diagramMoveUp', { label: 'Move Up', icon: 'moveUp', enabled: inDiagram, run: () => DG.op('up') });
  C('diagramMoveDown', { label: 'Move Down', icon: 'moveDown', enabled: inDiagram, run: () => DG.op('down') });
  C('diagramTextPane', { label: 'Text Pane', icon: 'taskpane', enabled: inDiagram, checked: () => !!(L.panes && L.panes.task.current === 'diagramText' && L.panes.task.el && !L.panes.task.el.hidden), run: () => { const on = ui.checked('diagramTextPane'); DG.showPane(!on); } });
  C('diagramLayout', { label: 'Change &Layout...', icon: 'layout', enabled: inDiagram, run: () => { const g = DG.current(); DG.gallery((k) => DG.set('layout', k, g), g.sa.layout); } });
  C('diagramColors', { label: 'Change Colors', icon: 'colorPic', enabled: inDiagram, run: () => {} });
  C('diagramStyle', { label: 'SmartArt Styles', icon: 'themes', enabled: inDiagram, run: () => {} });
  C('diagramReset', { label: 'Reset Graphic', icon: 'resetGraphic', enabled: inDiagram, run: () => DG.reset() });
  C('diagramToShapes', { label: 'Convert to Shapes', icon: 'toShapes', enabled: inDiagram, run: () => DG.toShapes() });
})();
