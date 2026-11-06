/* Lectern — slide editor: selection, handles, move/resize/rotate, drawing tools. */
(function () {
  'use strict';
  const L = window.L;
  const { h } = L;
  const M = L.model;
  const E = (L.ed = {
    idx: 0, sel: [], zoom: 'fit', scale: 1, tool: null,
    grid: { show: false, snap: true, size: 6, guides: false, gx: null, gy: null },
    view: 'normal', masterKind: 'slide', masterSlide: null,
  });

  E.slide = () => (E.view === 'master' ? E.masterSlide : L.pres.slides[E.idx]);
  E.design = () => M.design(L.pres, E.slide());
  E.find = (id) => { const s = E.slide(); return s ? M.find(s, id) : null; };
  E.shape = (id) => { const f = E.find(id); return f ? f.shape : null; };
  E.selected = () => E.sel.map(E.shape).filter(Boolean);
  E.primary = () => E.shape(E.sel[0]);
  E.hasSel = () => E.selected().length > 0;
  E.editing = () => L.te && L.te.active();

  /* ---------- DOM ---------- */
  E.mount = function (container) {
    E.wrap = container;
    E.scroller = h('div', { class: 'ed-scroll', tabindex: '-1', 'aria-label': 'Slide editing area' });
    E.canvas = h('div', { class: 'ed-canvas' });
    E.host = h('div', { class: 'ed-host' });
    E.gridEl = h('div', { class: 'ed-grid', hidden: true });
    E.overlay = h('div', { class: 'ed-overlay' });
    E.rulerH = h('div', { class: 'ruler ruler-h', hidden: true });
    E.rulerV = h('div', { class: 'ruler ruler-v', hidden: true });
    E.empty = h('div', { class: 'ed-empty', hidden: true, text: 'Click to add first slide' });
    E.canvas.append(E.host, E.gridEl, E.overlay);
    E.scroller.append(E.canvas, E.empty);
    container.append(E.rulerH, E.rulerV, E.scroller);
    E.empty.addEventListener('click', () => L.ui.exec('newSlide'));
    E.scroller.addEventListener('pointerdown', (e) => {
      if (e.button === 0) pressStart();
      try { onDown(e); } finally { if (!L.ui.dragging()) pressEnd(true); }
    });
    E.scroller.addEventListener('dblclick', onDbl);
    E.scroller.addEventListener('contextmenu', onContext);
    E.scroller.addEventListener('wheel', (e) => {
      if (e.ctrlKey) { e.preventDefault(); const z = E.scale / L.PX_PER_PT; E.setZoom(L.clamp(z * (e.deltaY < 0 ? 1.1 : 0.9), 0.1, 4)); return; }
      /* in fit mode, the wheel pages through slides like PowerPoint */
      if (E.zoom === 'fit' && E.view === 'normal' && !E.editing() && Math.abs(e.deltaY) > 20) {
        const now = Date.now();
        if (now - (E._wheelT || 0) < 350) return;
        E._wheelT = now;
        E.goto(E.idx + (e.deltaY > 0 ? 1 : -1));
      }
    }, { passive: false });
    if (window.ResizeObserver) new ResizeObserver(() => { if (E.zoom === 'fit') E.layout(); else E.layout(); }).observe(E.scroller);
    else window.addEventListener('resize', () => E.layout());
  };

  E.refocus = () => {
    if (E.editing()) { L.te.focus(); return; }
    if (L.app && L.app.focusArea === 'slides') { const b = L.panes && L.panes.left && L.panes.left.body; if (b && b.isConnected && document.activeElement !== b && !b.contains(document.activeElement)) { try { b.focus({ preventScroll: true }); } catch (e) { /* ignore */ } } return; }
    try { E.scroller.focus({ preventScroll: true }); } catch (e) { /* ignore */ }
  };

  E.layout = function () {
    if (!E.scroller || !E.scroller.isConnected) return;
    const W = L.pres.W, H = L.pres.H;
    const cw = E.scroller.clientWidth, ch = E.scroller.clientHeight;
    if (!cw || !ch) return;
    const margin = Math.max(16, Math.min(40, cw * 0.04));
    if (E.zoom === 'fit') E.scale = Math.max(0.05, Math.min((cw - margin * 2) / W, (ch - margin * 2) / H));
    else E.scale = E.zoom * L.PX_PER_PT;
    const sw = W * E.scale, sh = H * E.scale;
    const canW = Math.max(cw, sw + margin * 2), canH = Math.max(ch, sh + margin * 2);
    E.canvas.style.width = canW + 'px';
    E.canvas.style.height = canH + 'px';
    E.ox = Math.round((canW - sw) / 2); E.oy = Math.round((canH - sh) / 2);
    for (const el of [E.host, E.overlay, E.gridEl]) {
      el.style.left = E.ox + 'px'; el.style.top = E.oy + 'px';
      el.style.width = sw + 'px'; el.style.height = sh + 'px';
    }
    const sl = E.host.firstChild;
    if (sl) sl.style.transform = `scale(${E.scale})`;
    drawGrid();
    E.drawRulers();
    E.drawOverlay();
    L.bus.emit('zoom');
  };
  E.setZoom = function (z) {
    if (z === 'fit') E.zoom = 'fit';
    else E.zoom = L.clamp(z, 0.1, 4);
    E.layout();
    L.ui.refresh();
  };
  E.zoomPct = () => Math.round((E.scale / L.PX_PER_PT) * 100);

  function drawGrid() {
    const g = E.gridEl;
    g.hidden = !E.grid.show && !E.grid.guides;
    if (g.hidden) return;
    L.clear(g);
    if (E.grid.show) {
      const step = E.grid.size * E.scale;
      const big = step * Math.max(1, Math.round(72 / E.grid.size / 2));
      g.style.backgroundImage = `radial-gradient(circle, rgba(0,0,0,.45) 0.7px, transparent 0.9px)`;
      g.style.backgroundSize = `${big}px ${big}px`;
      g.style.backgroundPosition = `${-big / 2}px ${-big / 2}px`;
    } else g.style.backgroundImage = 'none';
    if (E.grid.guides) {
      const W = L.pres.W, H = L.pres.H;
      if (E.grid.gx == null) E.grid.gx = W / 2;
      if (E.grid.gy == null) E.grid.gy = H / 2;
      const gv = h('div', { class: 'guide gv', style: `left:${E.grid.gx * E.scale}px` });
      const gh = h('div', { class: 'guide gh', style: `top:${E.grid.gy * E.scale}px` });
      gv.addEventListener('pointerdown', (e) => dragGuide(e, 'x'));
      gh.addEventListener('pointerdown', (e) => dragGuide(e, 'y'));
      g.append(gv, gh);
    }
  }
  function dragGuide(e, axis) {
    if (e.button !== 0) return;
    e.stopPropagation(); e.preventDefault();
    const el = e.currentTarget;
    const tip = h('div', { class: 'guide-tip' });
    const was = axis === 'x' ? E.grid.gx : E.grid.gy;
    const put = (v) => { if (axis === 'x') { E.grid.gx = v; el.style.left = v * E.scale + 'px'; } else { E.grid.gy = v; el.style.top = v * E.scale + 'px'; } };
    L.ui.drag(e, {
      capture: E.scroller,
      move(ev) {
        const p = E.toSlide(ev.clientX, ev.clientY);
        const v = L.clamp(axis === 'x' ? p.x : p.y, 0, axis === 'x' ? L.pres.W : L.pres.H);
        if (!tip.isConnected) E.gridEl.appendChild(tip);
        put(v);
        if (axis === 'x') { tip.style.cssText = `left:${v * E.scale + 6}px;top:6px`; tip.textContent = L.round(Math.abs(v - L.pres.W / 2) / 72, 2) + '"'; }
        else { tip.style.cssText = `left:6px;top:${v * E.scale + 6}px`; tip.textContent = L.round(Math.abs(v - L.pres.H / 2) / 72, 2) + '"'; }
      },
      up() { tip.remove(); },
      cancel() { tip.remove(); put(was); },
    });
  }

  L.onLook(() => { if (E.rulerH) E.drawRulers(); });   /* the canvas follows the look */
  E.drawRulers = function () {
    const show = !!(L.app && L.app.opts.ruler) && E.view !== 'sorter';
    E.rulerH.hidden = E.rulerV.hidden = !show;
    E.wrap.classList.toggle('with-rulers', show);
    if (!show) return;
    const W = L.pres.W, H = L.pres.H, s = E.scale;
    const mk = (len, horiz) => {
      const c = h('canvas');
      const dpr = window.devicePixelRatio || 1;
      const size = horiz ? E.scroller.clientWidth : E.scroller.clientHeight;
      c.width = Math.max(1, size * dpr); c.height = 18 * dpr;
      c.style.width = size + 'px'; c.style.height = '18px';
      const g = c.getContext('2d');
      g.scale(dpr, dpr);
      g.fillStyle = L.lookColor('ruler', '#fff'); g.fillRect(0, 0, size, 18);
      const off = (horiz ? E.ox - E.scroller.scrollLeft : E.oy - E.scroller.scrollTop);
      g.fillStyle = L.lookColor('ruler-off-2', '#c5d6f3');
      g.fillRect(0, 0, off, 18); g.fillRect(off + len * s, 0, size, 18);
      g.strokeStyle = L.lookColor('ruler-tick-2', '#4b5d7e'); g.fillStyle = L.lookColor('ruler-ink-2', '#334'); g.font = '9px Tahoma, Verdana, sans-serif'; g.textAlign = 'center';
      const mid = len / 2;
      for (let pt = 0; pt <= len + 0.1; pt += 9) {
        const x = Math.round(off + pt * s) + 0.5;
        const rel = Math.abs(pt - mid) / 72;
        const isIn = Math.abs(rel - Math.round(rel)) < 0.01;
        const half = Math.abs(rel * 2 - Math.round(rel * 2)) < 0.01;
        const tl = isIn ? 0 : half ? 6 : 3;
        if (isIn) { g.fillText(String(Math.round(rel)), x, 12); continue; }
        g.beginPath(); g.moveTo(x, 18); g.lineTo(x, 18 - tl); g.stroke();
      }
      if (!horiz) { c.style.transform = `rotate(90deg) translate(0,-18px)`; c.style.transformOrigin = '0 0'; }
      return c;
    };
    L.clear(E.rulerH).appendChild(mk(W, true));
    L.clear(E.rulerV).appendChild(mk(H, false));
  };

  /* ---------- render ---------- */
  E.render = function (opts) {
    opts = opts || {};
    const sl = E.slide();
    const wasEditing = E.editing() ? L.te.snapshot() : null;
    if (wasEditing && !opts.keepEdit) L.te.detach();
    L.clear(E.host);
    E.empty.hidden = !!sl;
    E.canvas.hidden = !sl;
    if (!sl) { E.sel = []; E.drawOverlay(); return; }
    const el = L.render.slide(L.pres, sl, { mode: 'edit', index: E.view === 'master' ? 0 : E.idx });
    el.style.transform = `scale(${E.scale})`;
    el.style.transformOrigin = '0 0';
    E.host.appendChild(el);
    E.host.classList.toggle('gray', !!(L.app && L.app.opts.gray));
    E.host.classList.toggle('bw', !!(L.app && L.app.opts.bw));
    E.sel = E.sel.filter((id) => E.shape(id));
    if (wasEditing && opts.keepEdit !== false && E.shape(wasEditing.id)) L.te.resume(wasEditing);
    E.layout();
    L.bus.emit('slide-rendered');
  };
  /** Replace one shape's element in place (fast path) */
  E.renderShape = function (id) {
    const f = E.find(id);
    const old = E.host.querySelector(`.shp[data-sid="${id}"]`);
    if (!f || !old) { E.render(); return; }
    const top = f.parent ? E.topOf(id) : f.shape;
    if (top !== f.shape) { E.render(); return; }
    const snap = E.editing() && L.te.target() && L.te.target().sh.id === id ? L.te.snapshot() : null;
    if (snap) L.te.detach();
    const ctx = { pres: L.pres, slide: E.slide(), design: E.design(), mode: 'edit', num: E.idx + (L.pres.firstNum || 1) };
    const el = L.render.shape(f.shape, ctx, {});
    old.replaceWith(el);
    if (snap) L.te.resume(snap);
    E.drawOverlay();
  };
  E.topOf = function (id) {
    const sl = E.slide();
    for (const s of sl.shapes) {
      if (s.id === id) return s;
      if (s.type === 'group') { let hit = false; M.walk(s.kids, (k) => { if (k.id === id) { hit = true; return false; } return true; }); if (hit) return s; }
    }
    return null;
  };
  E.elOf = (id) => E.host.querySelector(`.shp[data-sid="${id}"]`);

  /* ---------- selection ---------- */
  E.select = function (ids, opts) {
    opts = opts || {};
    if (E.editing() && !opts.keepEdit) L.te.end();
    E.sel = (Array.isArray(ids) ? ids : [ids]).filter(Boolean);
    E.cellSel = null;
    E.drawOverlay();
    L.bus.emit('selection');
    L.ui.refresh();
  };
  E.clearSel = () => E.select([]);
  E.toggleSel = (id) => { const i = E.sel.indexOf(id); if (i >= 0) E.sel.splice(i, 1); else E.sel.push(id); E.select(E.sel.slice()); };
  E.selectAll = () => { const sl = E.slide(); if (sl) E.select(sl.shapes.map((s) => s.id)); };
  E.saveSel = () => ({ idx: E.idx, view: E.view, sel: E.sel.slice(), edit: E.editing() ? L.te.snapshot() : null });
  E.restoreSel = function (st) {
    E.endInteraction();   /* the model was replaced (undo, redo): a drag in progress holds shapes no longer there */
    if (!st) { E.render(); return; }
    if (E.view === 'master') E.buildMaster();
    E.idx = L.clamp(st.idx || 0, 0, Math.max(0, L.pres.slides.length - 1));
    E.sel = (st.sel || []).filter((id) => E.shape(id));
    /* the model was just replaced: never sync the (stale) editing DOM back into it */
    if (L.te.active()) L.te.detach(true);
    E.render({ keepEdit: false });
    if (st.edit && E.shape(st.edit.id)) L.te.resume(st.edit);
    L.bus.emit('slide-changed');
    L.ui.refresh();
  };
  E.goto = function (i, opts) {
    const n = L.pres.slides.length;
    if (!n) { E.idx = 0; E.render(); return; }
    i = L.clamp(i, 0, n - 1);
    E.endInteraction();   /* another slide: what was under way on this one ends */
    if (E.editing()) L.te.end();
    if (E.view === 'master') { E.view = 'normal'; L.bus.emit('view'); }
    const changed = i !== E.idx;
    E.idx = i;
    E.sel = [];
    E.render();
    if (changed || (opts && opts.force)) L.bus.emit('slide-changed');
    L.ui.refresh();
  };

  /* ---------- overlay (selection handles) ---------- */
  const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
  const HPOS = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5] };
  E.drawOverlay = function () {
    if (!E.overlay) return;
    L.clear(E.overlay);
    const s = E.scale;
    const editId = E.editing() ? L.te.target().sh.id : null;
    for (const id of E.sel) {
      const f = E.find(id);
      if (!f) continue;
      const sh = f.shape;
      const vis = E.visualBox(sh, f.parent);
      if (f.parent) {
        const gb = f.parent;
        const gfr = h('div', { class: 'sel-frame grp-frame', style: `left:${gb.x * s}px;top:${gb.y * s}px;width:${gb.w * s}px;height:${gb.h * s}px;${gb.rot ? `transform:rotate(${gb.rot}deg)` : ''}` });
        E.overlay.appendChild(gfr);
      }
      if (sh.type === 'line' || (sh.type === 'shape' && L.geom.isLine(sh.geom))) {
        const [p1, p2] = L.geom.lineEnds(sh);
        for (const [k, p] of [['p1', p1], ['p2', p2]]) {
          E.overlay.appendChild(h('div', { class: 'hd hd-end', 'data-h': k, 'data-sid': id, style: `left:${p[0] * s}px;top:${p[1] * s}px` }));
        }
        continue;
      }
      const fr = h('div', { class: 'sel-frame' + (editId === id ? ' editing' : '') + (sh.type === 'group' ? ' is-group' : ''), 'data-sid': id });
      fr.style.cssText = `left:${vis.x * s}px;top:${vis.y * s}px;width:${Math.max(vis.w * s, 1)}px;height:${Math.max(vis.h * s, 1)}px;${vis.rot ? `transform:rotate(${vis.rot}deg)` : ''}`;
      if (editId === id) {
        fr.appendChild(h('div', { class: 'edit-ring', 'data-h': 'move', 'data-sid': id }));
      } else if (sh.type === 'group') fr.appendChild(h('div', { class: 'grp-ring', 'data-h': 'move', 'data-sid': id }));
      const locked = !!f.parent && !!f.parent.rot;
      if (!locked) {
        if (E.cropMode && sh.type === 'image' && E.sel.length === 1) {
          for (const k of HANDLES) fr.appendChild(h('div', { class: 'hd crop-' + k, 'data-h': 'crop-' + k, 'data-sid': id, style: `left:${HPOS[k][0] * 100}%;top:${HPOS[k][1] * 100}%` }));
        } else {
          for (const k of HANDLES) {
            if (sh.type === 'table' && (k === 'n' || k === 's')) continue;
            fr.appendChild(h('div', { class: 'hd', 'data-h': k, 'data-sid': id, style: `left:${HPOS[k][0] * 100}%;top:${HPOS[k][1] * 100}%` }));
          }
          if (E.selected().every(M.canRotate)) {
            fr.appendChild(h('div', { class: 'hd-rotline' }));
            fr.appendChild(h('div', { class: 'hd hd-rot', 'data-h': 'rot', 'data-sid': id }));
          }
          if (E.sel.length === 1 && (sh.type === 'shape' || sh.type === 'text') && !f.parent) {
            L.geom.handles(sh).forEach((hd, i) => {
              let [x, y] = hd.pos(sh.w, sh.h, L.geom.adj(sh));
              if (sh.flipH) x = sh.w - x;
              if (sh.flipV) y = sh.h - y;
              fr.appendChild(h('div', { class: 'hd hd-adj', 'data-h': 'adj' + i, 'data-sid': id, style: `left:${x * s}px;top:${y * s}px` }));
            });
          }
        }
      }
      E.overlay.appendChild(fr);
    }
    if (E.cellSel && E.editing()) {
      const t = L.te.target();
      if (t && t.sh.type === 'table') {
        const { r0, c0, r1, c1 } = E.cellSel;
        const tbl = t.sh.tbl;
        let x = t.sh.x, y = t.sh.y;
        for (let c = 0; c < Math.min(c0, c1); c++) x += tbl.cols[c];
        for (let r = 0; r < Math.min(r0, r1); r++) y += tbl.rows[r].h;
        let w = 0, hh = 0;
        for (let c = Math.min(c0, c1); c <= Math.max(c0, c1); c++) w += tbl.cols[c];
        for (let r = Math.min(r0, r1); r <= Math.max(r0, r1); r++) hh += tbl.rows[r].h;
        if (r0 !== r1 || c0 !== c1) E.overlay.appendChild(h('div', { class: 'cell-sel', style: `left:${x * s}px;top:${y * s}px;width:${w * s}px;height:${hh * s}px` }));
      }
    }
    E.drawRulers();
  };
  /** Visual box of a shape (accounts for parent group rotation) */
  E.visualBox = function (sh, parent) {
    if (!parent || !parent.rot) return { x: sh.x, y: sh.y, w: sh.w, h: sh.h, rot: sh.rot || 0 };
    const gcx = parent.x + parent.w / 2, gcy = parent.y + parent.h / 2;
    const [cx, cy] = L.rotPt(sh.x + sh.w / 2, sh.y + sh.h / 2, gcx, gcy, parent.rot);
    return { x: cx - sh.w / 2, y: cy - sh.h / 2, w: sh.w, h: sh.h, rot: (sh.rot || 0) + parent.rot };
  };

  E.toSlide = (cx, cy) => {
    const r = E.host.getBoundingClientRect();
    return { x: (cx - r.left) / E.scale, y: (cy - r.top) / E.scale };
  };
  const snapV = (v, ev) => (E.grid.snap && !(ev && ev.altKey) ? Math.round(v / E.grid.size) * E.grid.size : v);

  /* ---------- commit helpers ---------- */
  E.commit = function (label, fn, opts) {
    L.hist.push(label);
    fn();
    if (E.view === 'master') E.syncMaster();
    if (opts && opts.shapeOnly && opts.shapeOnly.length === 1) E.renderShape(opts.shapeOnly[0]);
    else E.render();
    E.touched();
  };
  E.touched = L.debounce(() => { L.bus.emit('slide-modified', E.idx); }, 120);
  E.touchNow = () => L.bus.emit('slide-modified', E.idx);

  /* ---------- pointer interaction ---------- */
  /* Every press-drag-release on the slide is a ui.drag (common/ui.js): it moves only past its threshold, ends with up on
     the release (d.moved false: it was a click), and with cancel instead on Esc, a second button, a release the page
     never saw, the window losing focus or a tool switch; cancel puts back what the drag moved and records nothing.
     Only the browser's own text selection inside the text being edited is not one. The editing area holds the
     pointer, so the slide re-rendering under a drag doesn't end it. */
  function slideDrag(e, o) {
    const then = (fn, end) => (...a) => { try { if (fn) fn(...a); } finally { L.ui.refresh(); end(...a); } };
    return L.ui.drag(e, { threshold: o.threshold, capture: o.capture || E.scroller, move: o.move, up: then(o.up, (ev, d) => pressEnd(!d.moved)), cancel: then(o.cancel, () => pressEnd(false)) });
  }
  /* The toolbars that come and go with the selection (Picture, Tables and Borders, …) change the slide area's size, and
     so where the slide is. While a press on the slide is under way, and for a double-click's time after a click, that
     waits (E.holdLayout), so nothing moves under the pointer: a drag from a table's cell stays on the cells it
     started on, and a double-click's second press lands where the first did. The press ends with its drag (whichever
     way it ends) or, when it started none, at once. */
  let pressing = false, laterTimer = 0, later = null;
  function pressStart() { pressing = true; clearTimeout(laterTimer); laterTimer = 0; }
  function pressEnd(clicked) {
    if (!pressing) return;
    pressing = false;
    if (later) { if (clicked) laterTimer = setTimeout(flushLater, 450); else flushLater(); }
  }
  function flushLater() { clearTimeout(laterTimer); laterTimer = 0; const fn = later; later = null; if (fn) fn(); }
  /** true when fn (a change of the slide area's size) must wait for the press; it then runs when the press is over */
  E.holdLayout = (fn) => { if (!pressing && !laterTimer) return false; later = fn; return true; };
  /** shapes a drag changes as it goes: put back on cancel; on commit, one undo step for what changed */
  function liveEdit(shapes) {
    const before = shapes.map((s) => [s, L.clone(s)]);
    const put = (s, o) => {
      for (const k of new Set(Object.keys(s).concat(Object.keys(o)))) {
        if (!(k in o)) delete s[k];
        else if (JSON.stringify(s[k]) !== JSON.stringify(o[k])) s[k] = L.clone(o[k]);
      }
    };
    return {
      revert() { for (const [s, o] of before) put(s, o); },
      commit(label) { const after = before.map(([s]) => L.clone(s)); this.revert(); L.hist.push(label); before.forEach(([s], i) => put(s, after[i])); },
    };
  }
  /** with the Format Painter on, what a press marks is painted when the mouse is let go: the text selected (id unset)
      or the object id; null when the painter is off */
  const painterUp = (id) => (L.app.painter.on() ? () => L.app.painter.release(id) : null);
  function shapeHit(target) {
    const el = target.closest && target.closest('.shp.ia');
    if (!el || !E.host.contains(el)) return null;
    let top = el;
    for (let p = el.parentElement; p && p !== E.host; p = p.parentElement) if (p.classList && p.classList.contains('shp') && p.classList.contains('ia')) top = p;
    return { top: top.dataset.sid, inner: el.dataset.sid };
  }
  function onDown(e) {
    if (e.button === 1) return;
    L.ui.closeMenus();
    /* the prompt of an empty presentation is a button of its own: its click adds the first slide (P15) */
    if (E.empty.contains(e.target)) return;
    /* the slide area's scroll bars only scroll: no press on the slide's background, selection and editing stay (P12) */
    if (e.target === E.scroller && (e.offsetX >= E.scroller.clientWidth || e.offsetY >= E.scroller.clientHeight)) return;
    /* a press on the slide takes the keyboard from the Slides pane: Delete, the arrows and the clipboard keys work on
       the slide again (the presses below prevent the browser's own focus change) */
    if (L.app && L.app.focusArea === 'slides') { L.app.focusArea = null; try { E.scroller.focus({ preventScroll: true }); } catch (err) { /* ignore */ } }
    const tgt = e.target;
    /* content placeholder icons */
    const act = tgt.closest && tgt.closest('[data-content-act]');
    if (act && e.button === 0) {
      e.preventDefault();
      const sid = act.closest('.shp').dataset.sid;
      const map = { table: 'insertTable', chart: 'insertChart', clipart: 'insertClipArt', picture: 'insertPicture', diagram: 'insertDiagram', media: 'insertMedia' };
      E.select([sid]);
      L.ui.exec(map[act.dataset.contentAct], { ph: sid });
      return;
    }
    if (E.tool && e.button === 0) { startTool(e); return; }
    const hd = tgt.dataset && tgt.dataset.h;
    if (hd && e.button === 0) { startHandle(e, hd, tgt.dataset.sid); return; }
    /* inside the active text editor → native caret handling */
    if (E.editing() && L.te.contains(tgt)) {
      if (e.button === 0 && L.te.target().sh.type === 'table') startCellSelect(e, painterUp());
      /* the browser selects the text; the painter only needs to know when the mouse is let go */
      else if (e.button === 0 && L.app.painter.on()) slideDrag(e, { threshold: 0, capture: e.target, up: painterUp() });   /* held by the text: see startCellSelect */
      return;
    }
    /* elsewhere inside the box being edited (its margins, the space under the text): place the caret and
       drag to select text, as PowerPoint does; the hatched ring around it moves the box */
    const st = L.te.state;
    if (st && !st.cell && e.button === 0 && st.el.contains(tgt)) { startTextSelect(e, st, painterUp()); return; }
    const hit = shapeHit(tgt);
    if (e.button === 2) {
      if (hit) { const id = pickId(hit); if (!E.sel.includes(id)) E.select([id]); }
      return;
    }
    if (!hit) {
      /* no browser selection on the slide's background (a double-click there would select its line break, and the next
         press would drag that selection instead of a rubber band) */
      if (e.target !== E.scroller) e.preventDefault();
      L.app.painter.stop(); /* a click beside the objects puts the brush down, kept on or not (as Impress) */
      if (E.editing()) L.te.end();
      if (!e.shiftKey && !e.ctrlKey) E.clearSel();
      E.refocus();
      startMarquee(e);
      return;
    }
    const id = pickId(hit);
    /* a press in a table's cell edits that cell at once, and a drag from it selects cells, as PowerPoint does; the
       table moves by its border (P8) */
    if (e.button === 0 && !e.shiftKey && !e.ctrlKey && !L.app.painter.on() && E.shape(id) && E.shape(id).type === 'table' && tgt.closest('td')) {
      L.te.begin(id, { point: { x: e.clientX, y: e.clientY } });
      if (L.te.state && L.te.contains(tgt)) { startCellSelect(e, null); return; }
    }
    e.preventDefault();
    if (L.app.painter.on()) { paintPress(e, id); return; }
    const wasSelected = E.sel.includes(id) && E.sel.length === 1;
    /* Shift or Ctrl adds to the selection and never edits text (P6): a click on a selected object takes it out
       (Ctrl only when it doesn't drag, which copies) */
    const multi = e.shiftKey || e.ctrlKey, wasIn = E.sel.includes(id);
    if (multi) {
      if (E.editing()) L.te.end();
      if (wasIn && !e.ctrlKey) { E.toggleSel(id); return; }
      if (!wasIn) { E.sel.push(id); E.select(E.sel.slice()); }
    } else if (!wasIn) E.select([id]);
    else if (E.editing()) L.te.end();
    E.refocus();
    startMove(e, id, wasSelected, multi ? { multi: true, wasIn } : null);
  }
  /** caret position for a point in or around the text being edited: above it → start, below → end, beside → that line */
  function caretNear(root, x, y) {
    const r = root.getBoundingClientRect();
    const edge = (end) => { const rg = document.createRange(); rg.selectNodeContents(root); rg.collapse(!end); return rg; };
    if (y < r.top) return edge(false);
    if (y > r.bottom) return edge(true);
    const rg = L.te.caretFromPoint(L.clamp(x, r.left + 1, r.right - 1), y);
    return rg && root.contains(rg.startContainer) ? rg : edge(x > r.left + r.width / 2);
  }
  /** select text from the press to the pointer (then: run when the mouse is let go); cancelled, the selection is the
      press's own (the caret where it went down, or Shift's extension to there) */
  function startTextSelect(e, st, then) {
    e.preventDefault();
    const root = st.root;
    L.te.focus();
    const sel = window.getSelection();
    const at = caretNear(root, e.clientX, e.clientY);
    let anchor = { node: at.startContainer, off: at.startOffset };
    if (e.shiftKey && sel.rangeCount && root.contains(sel.anchorNode)) anchor = { node: sel.anchorNode, off: sel.anchorOffset };
    const to = (f) => { try { sel.setBaseAndExtent(anchor.node, anchor.off, f.startContainer, f.startOffset); } catch (err) { /* outside the text */ } };
    to(at);
    slideDrag(e, {
      threshold: 0,
      move: (ev) => to(caretNear(root, ev.clientX, ev.clientY)),
      up() { L.te.lastSel = L.te.getSel(); if (then) then(); },
      cancel() { if (root.isConnected) to(at); L.te.lastSel = L.te.getSel(); },
    });
  }
  /** a press on another object with the Format Painter on: on its text, when the painter carries text
      formatting, it edits that text and the drag marks what to paint; elsewhere the whole object is painted */
  function paintPress(e, id) {
    const sh = E.shape(id), fmt = L.app.painter.value;
    if (fmt.run && textCapable(sh) && !L.txt.isEmpty(sh.tx) && e.target.closest && e.target.closest('.txi')) {
      L.te.begin(id, { point: { x: e.clientX, y: e.clientY } });
      if (L.te.state && !L.te.state.cell) { startTextSelect(e, L.te.state, painterUp()); return; }
    }
    slideDrag(e, { up: painterUp(id) });
  }
  function pickId(hit) {
    if (hit.top === hit.inner) return hit.top;
    /* group: second click drills into the child */
    const inGroupSel = E.sel.includes(hit.top) || E.sel.some((s) => { const f = E.find(s); return f && f.parent && E.topOf(s) && E.topOf(s).id === hit.top; });
    return inGroupSel ? hit.inner : hit.top;
  }
  const textCapable = (sh) => sh && (sh.type === 'text' || sh.type === 'shape') && !L.geom.isLine(sh.geom) && sh.tx;

  /** opts: { multi, wasIn } a press with Shift or Ctrl (it selects, never edits); { ring } a press on the hatched
      ring of the box being edited (a click there selects the box as an object, P11) */
  function startMove(e, id, wasSelected, opts) {
    opts = opts || {};
    const p0 = E.toSlide(e.clientX, e.clientY);
    const ids = E.sel.slice();
    const shapes = ids.map((x) => E.find(x)).filter((f) => f && !(f.parent && f.parent.rot));
    const orig = shapes.map((f) => ({ f, x: f.shape.x, y: f.shape.y }));
    const prim = E.shape(id);
    const els = shapes.map((f) => E.elOf(f.shape.id));
    let mdx = 0, mdy = 0;
    const unmove = () => { els.forEach((el) => { if (el) el.style.translate = ''; }); L.$$('.sel-frame,.hd-end', E.overlay).forEach((el) => { el.style.translate = ''; }); };
    slideDrag(e, {
      move(ev) {
        const p = E.toSlide(ev.clientX, ev.clientY);
        let dx = p.x - p0.x, dy = p.y - p0.y;
        if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
        if (prim && E.grid.snap && !ev.altKey) {
          const o = orig.find((q) => q.f.shape.id === prim.id) || orig[0];
          if (o) { dx = snapV(o.x + dx, ev) - o.x; dy = snapV(o.y + dy, ev) - o.y; }
        }
        mdx = dx; mdy = dy;
        els.forEach((el) => { if (el) el.style.translate = `${dx}px ${dy}px`; });
        L.$$('.sel-frame,.hd-end', E.overlay).forEach((el) => { el.style.translate = `${dx * E.scale}px ${dy * E.scale}px`; });
        L.app && L.app.status && L.app.status(L.t('Position: {0} from left, {1} from top', L.fmtIn(o0().x + dx), L.fmtIn(o0().y + dy)));
      },
      up(ev, d) {
        if (!d.moved) {
          if (opts.multi) { if (ev.ctrlKey && opts.wasIn && E.sel.length > 1) E.toggleSel(id); return; }
          if (opts.ring) { if (E.editing()) L.te.end(); E.select([id]); E.refocus(); return; }
          if (prim && textCapable(prim) && (wasSelected || prim.type === 'text' || (prim.ph && L.txt.isEmpty(prim.tx)))) {
            if (!(prim.type === 'shape' && !wasSelected)) { L.te.begin(prim.id, { point: { x: ev.clientX, y: ev.clientY } }); clickEdit = { id: prim.id, t: Date.now() }; }
          } else if (prim && prim.type === 'table') L.te.begin(prim.id, { point: { x: ev.clientX, y: ev.clientY } });
          return;
        }
        const dx = mdx, dy = mdy;
        unmove();
        if (ev.ctrlKey && !dx && !dy) return;
        if (ev.ctrlKey) {
          L.hist.push(L.N('Duplicate'));
          const sl = E.slide();
          const copies = shapes.map((f) => { const c = M.dup(f.shape); M.translate(c, dx, dy); f.list.push(c); return c.id; });
          if (E.view === 'master') E.syncMaster();
          E.sel = copies;
          E.render();
          E.touched();
          void sl;
          return;
        }
        E.commit('Move Object', () => { for (const o of orig) { M.translate(o.f.shape, o.x + dx - o.f.shape.x, o.y + dy - o.f.shape.y); if (o.f.parent) E.fitGroup(o.f.parent); } });
      },
      cancel() { unmove(); if (L.app && L.app.status) L.app.status(''); },
    });
    function o0() { return orig.find((q) => prim && q.f.shape.id === prim.id) || orig[0] || { x: 0, y: 0 }; }
  }
  E.fitGroup = function (g) {
    if (!g || g.type !== 'group' || g.rot) return;
    const b = M.groupBounds(g);
    g.x = b.x; g.y = b.y; g.w = b.w; g.h = b.h;
  };

  function startHandle(e, hd, sid) {
    e.preventDefault(); e.stopPropagation();
    const f = E.find(sid);
    if (!f) return;
    const sh = f.shape;
    if (hd === 'move') { startMove(e, sid, true, { ring: e.target.classList.contains('edit-ring') }); return; }
    if (hd === 'rot' && !E.selected().every(M.canRotate)) return;
    const orig = L.clone(sh);
    const others = E.sel.filter((x) => x !== sid).map((x) => E.find(x)).filter(Boolean).map((of) => ({ f: of, o: L.clone(of.shape) }));
    const live = () => { E.renderShape(E.topOf(sid).id); for (const o of others) E.renderShape(E.topOf(o.f.shape.id).id); };
    /* the shapes change as the pointer goes; the release records them as one step (label), a cancel puts them back */
    const edit = liveEdit([sh].concat(others.map((o) => o.f.shape)));
    const handle = (label, move, after) => slideDrag(e, {
      move,
      up(ev, d) {
        if (!d.moved) return;
        edit.commit(label);
        if (after) after();
        if (E.view === 'master') E.syncMaster();
        E.render(); E.touched();
      },
      cancel(d) { if (d.moved) { edit.revert(); E.render(); } },
    });
    if (hd === 'p1' || hd === 'p2') {
      handle(L.N('Resize'), (ev) => {
        let p = E.toSlide(ev.clientX, ev.clientY);
        const [a, b] = L.geom.lineEnds(orig);
        const fixed = hd === 'p1' ? b : a;
        if (ev.shiftKey) {
          const ang = Math.atan2(p.y - fixed[1], p.x - fixed[0]);
          const sn = (Math.round(ang / (Math.PI / 12)) * Math.PI) / 12;
          const len = Math.hypot(p.x - fixed[0], p.y - fixed[1]);
          p = { x: fixed[0] + Math.cos(sn) * len, y: fixed[1] + Math.sin(sn) * len };
        } else p = { x: snapV(p.x, ev), y: snapV(p.y, ev) };
        L.geom.setLineEnds(sh, hd === 'p1' ? [p.x, p.y] : a, hd === 'p2' ? [p.x, p.y] : b);
        live();
      }, () => { if (f.parent) E.fitGroup(f.parent); });
      return;
    }
    if (hd === 'rot') {
      const cx = orig.x + orig.w / 2, cy = orig.y + orig.h / 2;
      handle(L.N('Rotate'), (ev) => {
        const p = E.toSlide(ev.clientX, ev.clientY);
        let ang = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
        if (ev.shiftKey) ang = Math.round(ang / 15) * 15;
        ang = ((ang % 360) + 360) % 360;
        const turn = ang - (orig.rot || 0);
        sh.rot = L.round(ang, 2);
        for (const o of others) o.f.shape.rot = L.round((((o.o.rot || 0) + turn) % 360 + 360) % 360, 2);
        live();
        L.app && L.app.status && L.app.status(L.t('Rotation: {0}°', Math.round(ang)));
      });
      return;
    }
    if (hd.startsWith('adj')) {
      const hdef = L.geom.handles(sh)[+hd.slice(3)];
      handle(L.N('Adjust'), (ev) => {
        const p = E.toSlide(ev.clientX, ev.clientY);
        const c = { x: orig.x + orig.w / 2, y: orig.y + orig.h / 2 };
        let [lx, ly] = L.rotPt(p.x, p.y, c.x, c.y, -(orig.rot || 0));
        lx -= orig.x; ly -= orig.y;
        if (sh.flipH) lx = orig.w - lx;
        if (sh.flipV) ly = orig.h - ly;
        sh.adj = Object.assign({}, L.geom.adj(orig), hdef.set(lx, ly, orig.w, orig.h, L.geom.adj(orig)));
        for (const k in sh.adj) sh.adj[k] = Math.round(sh.adj[k]);
        live();
      });
      return;
    }
    if (hd.startsWith('crop-')) { startCrop(e, hd.slice(5), sh, orig); return; }
    /* resize */
    const rot = orig.rot || 0;
    const c0 = { x: orig.x + orig.w / 2, y: orig.y + orig.h / 2 };
    const ar = orig.h ? orig.w / orig.h : 1;
    const isCorner = hd.length === 2;
    handle(L.N('Resize'), (ev) => {
      const p = E.toSlide(ev.clientX, ev.clientY);
      let [lx, ly] = L.rotPt(p.x, p.y, c0.x, c0.y, -rot);
      lx -= c0.x; ly -= c0.y;
      let l = -orig.w / 2, t = -orig.h / 2, r = orig.w / 2, b = orig.h / 2;
      if (!rot && E.grid.snap && !ev.altKey) { lx = snapV(lx + c0.x, ev) - c0.x; ly = snapV(ly + c0.y, ev) - c0.y; }
      if (hd.includes('w')) l = lx;
      if (hd.includes('e')) r = lx;
      if (hd.includes('n')) t = ly;
      if (hd.includes('s')) b = ly;
      if (ev.ctrlKey) {
        if (hd.includes('w')) r = -l; if (hd.includes('e')) l = -r;
        if (hd.includes('n')) b = -t; if (hd.includes('s')) t = -b;
      }
      let w = r - l, hh = b - t;
      const keep = (isCorner && (ev.shiftKey || sh.lockAspect)) || (sh.type === 'chart' && ev.shiftKey);
      if (keep && ar) {
        const sw = Math.abs(w) / orig.w, sH = Math.abs(hh) / orig.h, k = Math.max(sw, sH);
        const nw = orig.w * k * Math.sign(w || 1), nh = orig.h * k * Math.sign(hh || 1);
        if (hd.includes('w')) l = r - nw; else r = l + nw;
        if (hd.includes('n')) t = b - nh; else b = t + nh;
        if (ev.ctrlKey) { l = -Math.abs(nw) / 2; r = Math.abs(nw) / 2; t = -Math.abs(nh) / 2; b = Math.abs(nh) / 2; }
        w = r - l; hh = b - t;
      }
      let flipX = false, flipY = false;
      if (w < 0) { flipX = true; [l, r] = [r, l]; w = -w; }
      if (hh < 0) { flipY = true; [t, b] = [b, t]; hh = -hh; }
      w = Math.max(w, 1); hh = Math.max(hh, 1);
      const [ncx, ncy] = L.rotPt(c0.x + (l + r) / 2, c0.y + (t + b) / 2, c0.x, c0.y, rot);
      const ob = { x: sh.x, y: sh.y, w: sh.w, h: sh.h };
      sh.x = L.round(ncx - w / 2, 3); sh.y = L.round(ncy - hh / 2, 3); sh.w = L.round(w, 3); sh.h = L.round(hh, 3);
      if (sh.type !== 'table' && sh.type !== 'group') { sh.flipH = flipX ? !orig.flipH : !!orig.flipH; sh.flipV = flipY ? !orig.flipV : !!orig.flipV; }
      if (sh.txRect) sh.txRect = sh.txRect.map((v, i) => v * (i % 2 ? sh.h / ob.h : sh.w / ob.w));
      if (sh.type === 'group') { if (sh.sa) L.diagram.relayout(sh); else M.scaleGroup(sh, ob, sh); }
      if (sh.type === 'table') E.scaleTable(sh, orig);
      const sx = sh.w / orig.w, sy = sh.h / orig.h;
      for (const o of others) {
        const s2 = o.f.shape, o2 = o.o;
        const ob2 = { x: s2.x, y: s2.y, w: s2.w, h: s2.h };
        s2.w = Math.max(1, o2.w * sx); s2.h = Math.max(1, o2.h * sy);
        s2.x = hd.includes('w') ? o2.x + o2.w - s2.w : o2.x;
        s2.y = hd.includes('n') ? o2.y + o2.h - s2.h : o2.y;
        if (s2.type === 'group') { if (s2.sa) L.diagram.relayout(s2); else M.scaleGroup(s2, ob2, s2); }
        if (s2.type === 'table') E.scaleTable(s2, o2);
      }
      live();
      L.app && L.app.status && L.app.status(L.t('Size: {0} × {1}', L.fmtIn(sh.w), L.fmtIn(sh.h)));
    }, () => { if (f.parent) E.fitGroup(f.parent); });
  }
  E.scaleTable = function (sh, orig) {
    const sx = sh.w / orig.w, sy = sh.h / orig.h;
    sh.tbl.cols = orig.tbl.cols.map((w) => L.round(w * sx, 2));
    sh.tbl.rows.forEach((r, i) => { r.h = L.round(Math.max(8, orig.tbl.rows[i].h * sy), 2); });
  };
  function startCrop(e, hd, sh, orig) {
    const c = Object.assign({ l: 0, t: 0, r: 0, b: 0 }, orig.crop);
    const fullW = orig.w / (1 - c.l - c.r), fullH = orig.h / (1 - c.t - c.b);
    const edit = liveEdit([sh]);
    let changed = false;
    slideDrag(e, {
      move(ev) {
        if (orig.rot) return;
        const p = E.toSlide(ev.clientX, ev.clientY);
        changed = true;
        const nc = Object.assign({}, c);
        const imgX = orig.x - c.l * fullW, imgY = orig.y - c.t * fullH;
        if (hd.includes('w')) nc.l = L.clamp((p.x - imgX) / fullW, -2, 1 - nc.r - 0.02);
        if (hd.includes('e')) nc.r = L.clamp((imgX + fullW - p.x) / fullW, -2, 1 - nc.l - 0.02);
        if (hd.includes('n')) nc.t = L.clamp((p.y - imgY) / fullH, -2, 1 - nc.b - 0.02);
        if (hd.includes('s')) nc.b = L.clamp((imgY + fullH - p.y) / fullH, -2, 1 - nc.t - 0.02);
        sh.crop = nc;
        sh.x = imgX + nc.l * fullW; sh.y = imgY + nc.t * fullH;
        sh.w = fullW * (1 - nc.l - nc.r); sh.h = fullH * (1 - nc.t - nc.b);
        E.renderShape(sh.id);
      },
      up() { if (changed) { edit.commit('Crop Picture'); E.render(); E.touched(); } },
      cancel() { if (changed) { edit.revert(); E.render(); } },
    });
  }

  function startMarquee(e) {
    const p0 = E.toSlide(e.clientX, e.clientY);
    const box = h('div', { class: 'marquee' });
    const base = e.shiftKey || e.ctrlKey ? E.sel.slice() : [];
    let r = null;
    slideDrag(e, {
      move(ev) {
        const p = E.toSlide(ev.clientX, ev.clientY);
        if (!box.isConnected) E.overlay.appendChild(box);
        const x = Math.min(p.x, p0.x), y = Math.min(p.y, p0.y), w = Math.abs(p.x - p0.x), hh = Math.abs(p.y - p0.y);
        box.style.cssText = `left:${x * E.scale}px;top:${y * E.scale}px;width:${w * E.scale}px;height:${hh * E.scale}px`;
        r = { x, y, w, h: hh };
      },
      up() {
        box.remove();
        if (!r) return;
        const ids = E.slide().shapes.filter((sh) => { const b = sh.type === 'group' ? M.groupBounds(sh) : L.rotBounds(sh); return b.x >= r.x - 0.5 && b.y >= r.y - 0.5 && b.x + b.w <= r.x + r.w + 0.5 && b.y + b.h <= r.y + r.h + 0.5; }).map((s) => s.id);
        E.select(Array.from(new Set(base.concat(ids))));
      },
      cancel() { box.remove(); },
    });
  }
  /** select table cells from the press to the pointer (the browser selects text within the one cell); then: run when the
      mouse is let go */
  function startCellSelect(e, then) {
    const td0 = e.target.closest('td');
    if (!td0) { if (then) slideDrag(e, { threshold: 0, capture: e.target, up: then }); return; }
    const r0 = +td0.dataset.r, c0 = +td0.dataset.c;
    if (E.cellSel) { E.cellSel = null; E.drawOverlay(); }   /* a press in one cell ends a range, and its highlight (P9) */
    slideDrag(e, {
      threshold: 0,
      capture: e.target,   /* held by the editing area (scaled), the browser would select the cell's text short of the pointer */
      move(ev) {
        const el = document.elementFromPoint(ev.clientX, ev.clientY);
        const td = el && el.closest && el.closest('td');
        if (!td || !td0.closest('table').contains(td)) return;
        const r1 = +td.dataset.r, c1 = +td.dataset.c;
        if (r1 === r0 && c1 === c0) { if (E.cellSel) { E.cellSel = null; E.drawOverlay(); } return; }
        E.cellSel = { r0, c0, r1, c1 };
        window.getSelection().removeAllRanges();
        E.drawOverlay();
      },
      up() { if (then) then(); },
      cancel() { if (E.cellSel) { E.cellSel = null; E.drawOverlay(); } },
    });
  }

  let clickEdit = null;   /* { id, t }: the text edit a click just started */
  function onDbl(e) {
    if (E.tool && ['freeform', 'curve'].includes(E.tool.kind)) return;
    /* a press's drag holds the pointer in the editing area, so the double-click is addressed to it: the shape is the
       one under the pointer */
    const hit = shapeHit(e.target === E.scroller ? document.elementFromPoint(e.clientX, e.clientY) || e.target : e.target);
    if (!hit) return;
    const id = pickId(hit);
    const sh = E.shape(id);
    if (!sh) return;
    if (sh.type === 'image') { L.ui.exec('formatObject'); return; }
    if (sh.type === 'chart') { L.ui.exec('editChart'); return; }
    if (sh.type === 'wordart') { L.ui.exec('wordartEdit'); return; }
    if (sh.type === 'line') { L.ui.exec('formatObject'); return; }
    if (textCapable(sh)) {
      if (!E.editing() || L.te.target().sh.id !== id) L.te.begin(id, { point: { x: e.clientX, y: e.clientY }, word: true });
      /* the double-click's second press started the edit (the first selected the shape): the browser placed no word
         selection, since the press was the slide's, so select the word here (P13) */
      else if (clickEdit && clickEdit.id === id && Date.now() - clickEdit.t < 1000) L.te.selectWordAt(e.clientX, e.clientY);
    }
    clickEdit = null;
  }
  function onContext(e) {
    if (L.show && L.show.active) return;
    if (E.editing() && L.te.contains(e.target)) { if (e.shiftKey) return; e.preventDefault(); L.app.contextText(e); return; }
    const hit = shapeHit(e.target);
    e.preventDefault();
    if (hit) { const id = pickId(hit); if (!E.sel.includes(id)) E.select([id]); L.app.contextShape(e); }
    else { E.clearSel(); L.app.contextSlide(e); }
  }

  /* ---------- drawing tools ---------- */
  /** the one way what is under way on the slide ends unfinished: the drag in progress (put back, nothing recorded), an
      unfinished freeform or curve, a range of table cells and its highlight. A tool switch, another slide and a model
      replaced (undo, redo) call it; Esc finishes a freeform first (app.js) */
  E.endInteraction = function () {
    L.ui.dragCancel();
    if (ff) { ff.el.remove(); ff = null; }
    if (E.cellSel) { E.cellSel = null; E.drawOverlay(); }
  };
  E.setTool = function (tool) {
    E.endInteraction();
    if (E.editing()) L.te.end();
    if (tool) L.ui.brushOn()?.stop();
    E.tool = tool;
    E.scroller.classList.toggle('drawing', !!tool);
    L.ui.refresh();
    if (tool) L.app && L.app.status && L.app.status(tool.kind === 'freeform' || tool.kind === 'curve' ? L.N('Click to add points; double-click to finish') : L.N('Drag to draw; hold Shift to constrain'));
  };
  function startTool(e) {
    e.preventDefault();
    const t = E.tool;
    const p0raw = E.toSlide(e.clientX, e.clientY);
    const p0 = { x: snapV(p0raw.x, e), y: snapV(p0raw.y, e) };
    if (t.kind === 'scribble') return startScribble(e, p0raw);
    if (t.kind === 'freeform' || t.kind === 'curve') return pointFreeform(e, p0raw);
    const ghost = h('div', { class: 'draw-ghost' + (t.kind === 'line' ? ' line' : '') });
    E.overlay.appendChild(ghost);
    let svgLine = null;
    if (t.kind === 'line') { ghost.innerHTML = '<svg width="100%" height="100%" style="overflow:visible;position:absolute;left:0;top:0"><line stroke="#000" stroke-width="1" stroke-dasharray="3 2"/></svg>'; svgLine = ghost.querySelector('line'); ghost.style.cssText = 'left:0;top:0;width:100%;height:100%'; }
    let p1 = p0;
    slideDrag(e, {
      move(ev) {
        const raw = E.toSlide(ev.clientX, ev.clientY);
        let x = snapV(raw.x, ev), y = snapV(raw.y, ev);
        if (t.kind === 'line') {
          if (ev.shiftKey) {
            const ang = Math.atan2(y - p0.y, x - p0.x), sn = (Math.round(ang / (Math.PI / 12)) * Math.PI) / 12, len = Math.hypot(x - p0.x, y - p0.y);
            x = p0.x + Math.cos(sn) * len; y = p0.y + Math.sin(sn) * len;
          }
          p1 = { x, y };
          svgLine.setAttribute('x1', p0.x * E.scale); svgLine.setAttribute('y1', p0.y * E.scale);
          svgLine.setAttribute('x2', x * E.scale); svgLine.setAttribute('y2', y * E.scale);
          return;
        }
        let w = x - p0.x, hh = y - p0.y;
        if (ev.shiftKey) { const m = Math.max(Math.abs(w), Math.abs(hh)); w = m * Math.sign(w || 1); hh = m * Math.sign(hh || 1); }
        let x0 = p0.x, y0 = p0.y;
        if (ev.ctrlKey) { x0 = p0.x - w; y0 = p0.y - hh; w *= 2; hh *= 2; }
        p1 = { x: x0 + w, y: y0 + hh, x0, y0 };
        const rx = Math.min(x0, x0 + w), ry = Math.min(y0, y0 + hh);
        ghost.style.cssText = `left:${rx * E.scale}px;top:${ry * E.scale}px;width:${Math.abs(w) * E.scale}px;height:${Math.abs(hh) * E.scale}px`;
      },
      up() {
        ghost.remove();
        const sl = E.slide();
        if (!sl) return;
        const x0 = p1.x0 != null ? p1.x0 : p0.x, y0 = p1.y0 != null ? p1.y0 : p0.y;
        const x = Math.min(x0, p1.x), y = Math.min(y0, p1.y), w = Math.abs(p1.x - x0), hh = Math.abs(p1.y - y0);
        const tiny = w < 3 && hh < 3;
        let sh;
        L.hist.push('Insert Object');
        if (t.kind === 'line') {
          if (tiny) { L.hist.undo.pop(); E.setTool(t.sticky ? t : null); return; }
          sh = M.newLine(L.pres, sl, t.geom || 'line', [p0.x, p0.y], [p1.x, p1.y], t.arrows);
        } else if (t.kind === 'text') {
          if (tiny) { sh = M.newTextBox(L.pres, sl, p0.x, p0.y, 72, 28.8, false); sh.tx.wrap = false; }
          else sh = M.newTextBox(L.pres, sl, x, y, w, Math.max(hh, 28.8), true);
          if (t.vert) sh.tx.vert = 'vert';
        } else {
          sh = M.newShape(L.pres, sl, t.geom, tiny ? p0.x : x, tiny ? p0.y : y, tiny ? 72 : w, tiny ? 72 : hh);
        }
        sl.shapes.push(sh);
        if (E.view === 'master') E.syncMaster();
        if (!t.sticky) E.setTool(null);
        E.sel = [sh.id];
        E.render();
        E.touched();
        if (t.kind === 'text') L.te.begin(sh.id, { atEnd: true });
      },
      cancel(d, why) { ghost.remove(); if (why === 'esc') E.setTool(null); },   /* Esc ends the tool; other cancels keep it */
    });
  }
  function startScribble(e, p0) {
    const pts = [[p0.x, p0.y]];
    const svg = h('div', { class: 'draw-ghost line', style: 'left:0;top:0;width:100%;height:100%' });
    svg.innerHTML = '<svg width="100%" height="100%" style="overflow:visible;position:absolute;left:0;top:0"><polyline fill="none" stroke="#000" stroke-width="1"/></svg>';
    E.overlay.appendChild(svg);
    const pl = svg.querySelector('polyline');
    slideDrag(e, {
      move(ev) {
        const p = E.toSlide(ev.clientX, ev.clientY);
        const last = pts[pts.length - 1];
        if (Math.hypot(p.x - last[0], p.y - last[1]) * E.scale < 2) return;
        pts.push([p.x, p.y]);
        pl.setAttribute('points', pts.map((q) => q[0] * E.scale + ',' + q[1] * E.scale).join(' '));
      },
      up() { svg.remove(); if (pts.length > 2) E.addFreeform(pts, false, true); if (!(E.tool && E.tool.sticky)) E.setTool(null); },
      cancel(d, why) { svg.remove(); if (why === 'esc') E.setTool(null); },
    });
  }
  let ff = null;
  function pointFreeform(e, p) {
    const now = Date.now();
    const had = ff ? ff.pts.length : 0;
    if (!ff) {
      ff = { pts: [[p.x, p.y]], el: h('div', { class: 'draw-ghost line', style: 'left:0;top:0;width:100%;height:100%' }) };
      ff.el.innerHTML = '<svg width="100%" height="100%" style="overflow:visible;position:absolute;left:0;top:0"><path fill="none" stroke="#000" stroke-width="1"/></svg>';
      E.overlay.appendChild(ff.el);
    } else {
      const first = ff.pts[0];
      const closeHit = Math.hypot(p.x - first[0], p.y - first[1]) * E.scale < 8 && ff.pts.length > 2;
      const last = ff.pts[ff.pts.length - 1];
      const dbl = now - ff.t < 400 && Math.hypot(p.x - last[0], p.y - last[1]) * E.scale < 6;
      if (dbl || closeHit) { finishFreeform(closeHit); return; }
      ff.pts.push([p.x, p.y]);
    }
    ff.t = now;
    const path = ff.el.querySelector('path');
    const draw = (extra) => {
      const pts = extra ? ff.pts.concat([extra]) : ff.pts;
      path.setAttribute('d', E.tool && E.tool.kind === 'curve' ? smoothD(pts.map((q) => [q[0] * E.scale, q[1] * E.scale])) : 'M' + pts.map((q) => q[0] * E.scale + ',' + q[1] * E.scale).join(' L'));
    };
    draw();
    /* while the button is held in freeform mode, drawing is freehand; the release keeps collecting points, a cancel
       takes back what this press added (the whole shape if it started it) */
    const mine = ff;
    slideDrag(e, {
      move(ev) {
        if (ff !== mine) return;
        const q = E.toSlide(ev.clientX, ev.clientY);
        if (E.tool && E.tool.kind === 'freeform') { const last = ff.pts[ff.pts.length - 1]; if (Math.hypot(q.x - last[0], q.y - last[1]) * E.scale > 3) ff.pts.push([q.x, q.y]); }
        draw([q.x, q.y]);
      },
      cancel(d, why) {
        if (why === 'esc') { E.setTool(null); return; }   /* Esc ends the tool and drops the shape, as when no button is down */
        if (ff !== mine) return;
        if (!had) { ff.el.remove(); ff = null; return; }
        ff.pts.length = had;
        draw();
      },
    });
    if (!E._ffMove) {
      E._ffMove = (ev) => { if (!ff) return; const q = E.toSlide(ev.clientX, ev.clientY); draw([q.x, q.y]); };
      E.scroller.addEventListener('pointermove', E._ffMove);
    }
  }
  function finishFreeform(closed) {
    if (!ff) return;
    const pts = ff.pts;
    ff.el.remove();
    const kind = E.tool && E.tool.kind;
    ff = null;
    if (pts.length > 1) E.addFreeform(pts, closed, kind === 'curve' ? 'curve' : false);
    if (!(E.tool && E.tool.sticky)) E.setTool(null);
  }
  E.cancelFreeform = (finish) => { if (ff) { if (finish) finishFreeform(false); else { ff.el.remove(); ff = null; } } };
  function smoothD(pts) {
    if (pts.length < 3) return 'M' + pts.map((q) => q.join(',')).join(' L');
    let d = `M${pts[0][0]},${pts[0][1]}`;
    for (let i = 0; i < pts.length - 1; i++) {
      const p0 = pts[i - 1] || pts[i], p1 = pts[i], p2 = pts[i + 1], p3 = pts[i + 2] || p2;
      const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6], c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
      d += ` C${c1[0]},${c1[1]} ${c2[0]},${c2[1]} ${p2[0]},${p2[1]}`;
    }
    return d;
  }
  /** Create a freeform (custom geometry) shape from slide points */
  E.addFreeform = function (pts, closed, smooth, opts) {
    const sl = E.slide();
    if (!sl) return null;
    const xs = pts.map((p) => p[0]), ys = pts.map((p) => p[1]);
    const x = Math.min(...xs), y = Math.min(...ys);
    const w = Math.max(Math.max(...xs) - x, 1), hh = Math.max(Math.max(...ys) - y, 1);
    const loc = pts.map((p) => [L.round(p[0] - x, 2), L.round(p[1] - y, 2)]);
    const cmds = [['M', loc[0][0], loc[0][1]]];
    if (smooth && loc.length > 2) {
      for (let i = 0; i < loc.length - 1; i++) {
        const p0 = loc[i - 1] || loc[i], p1 = loc[i], p2 = loc[i + 1], p3 = loc[i + 2] || p2;
        cmds.push(['C', L.round(p1[0] + (p2[0] - p0[0]) / 6, 2), L.round(p1[1] + (p2[1] - p0[1]) / 6, 2), L.round(p2[0] - (p3[0] - p1[0]) / 6, 2), L.round(p2[1] - (p3[1] - p1[1]) / 6, 2), p2[0], p2[1]]);
      }
    } else for (let i = 1; i < loc.length; i++) cmds.push(['L', loc[i][0], loc[i][1]]);
    if (closed) cmds.push(['Z']);
    const sh = M.newShape(L.pres, sl, 'rect', x, y, w, hh);
    sh.geom = 'custom';
    sh.name = (closed ? 'Freeform ' : 'Freeform ') + (sl.shapes.length + 1);
    sh.path = { paths: [{ w, h: hh, cmds, fill: closed ? 'norm' : 'none', stroke: true }] };
    if (!closed) { sh.fill = { t: 'none' }; delete sh.tx; }
    if (opts && opts.line) sh.line = opts.line;
    if (!(opts && opts.noHistory)) L.hist.push(L.N('Freeform'));
    sl.shapes.push(sh);
    if (opts && opts.noRender) return sh;
    E.sel = [sh.id];
    E.render();
    E.touched();
    return sh;
  };

  /* ---------- format painter ---------- */
  /* As Impress does (sd FuFormatPaintBrush): from the text being edited the painter carries the text's
     formatting only (and the paragraph's when the selection is a caret or a whole paragraph); from a selected
     object, its fill, line and shadow and its text's. It carries the text's look: the run's own formatting over
     its paragraph's level style. Painted, that replaces a run's own formatting; links, fields and kept XML stay. */
  const painted = () => L.txt.RUN_KEYS.filter((k) => !['link', 'fld', 'keep'].includes(k));
  function lookOf(sh, p, own) {
    const look = L.style.run(L.style.para(sh, p, M.design(L.pres, E.slide())), own);
    const out = {};
    for (const k of painted()) if (look[k] !== undefined) out[k] = L.clone(look[k]);
    return out;
  }
  const paintRun = (r, run) => { const o = Object.assign({}, r); for (const k of painted()) { if (run[k] !== undefined) o[k] = L.clone(run[k]); else delete o[k]; } return o; };
  const wholePara = (tx, s) => s.p0 === s.p1 && s.o0 === 0 && s.o1 >= L.txt.len(tx.ps[s.p0]);
  E.pickFormat = function () {
    const sh = E.primary();
    if (!sh) return null;
    if (E.editing()) {
      const t = L.te.target(), s = L.txt.norm(L.te.currentSel()), p = t.tx.ps[s.p0];
      const fmt = { run: lookOf(t.sh, p, L.te.runPropsAtSel()) };
      if (L.txt.collapsed(s) || wholePara(t.tx, s)) fmt.para = L.clone(p.pp || {});
      return fmt;
    }
    const fmt = { fill: L.clone(sh.fill), line: L.clone(sh.line), shadow: L.clone(sh.shadow) };
    if (sh.tx && sh.tx.ps[0]) { const p = sh.tx.ps[0]; fmt.run = lookOf(sh, p, L.txt.runProps(p.rs[0] || p.end || {})); fmt.para = L.clone(p.pp || {}); }
    if (sh.type === 'wordart') fmt.wa = { fill: L.clone(sh.wa.fill), line: L.clone(sh.wa.line), shadow: L.clone(sh.wa.shadow), font: sh.wa.font, b: sh.wa.b, i: sh.wa.i };
    return fmt;
  };
  /** paint the text selected in the text being edited, or the word a click landed in; false: nothing to paint */
  E.paintText = function (fmt) {
    const t = E.editing() && L.te.target();
    let s = t && fmt.run && L.te.getSel();
    if (!s) return false;
    s = L.txt.norm(s);
    if (L.txt.collapsed(s) && !(s = L.txt.wordAt(t.tx, { p: s.p0, o: s.o0 }))) return false;
    L.te.applyRun((r) => paintRun(r, fmt.run), 'Format Painter', s);
    if (fmt.para && wholePara(t.tx, s)) { const p = t.tx.ps[s.p0]; p.pp = Object.assign({}, p.pp, fmt.para); L.te.rerender(s); }
    return true;
  };
  /** paint a whole object: its fill, line and shadow, and all its text */
  E.paintShape = function (id, fmt) {
    const sh = E.shape(id);
    if (!sh) return false;
    if (E.editing()) L.te.end();
    E.commit('Format Painter', () => {
      if (fmt.fill && sh.type !== 'line' && sh.type !== 'image' && (sh.type !== 'text' || fmt.fill.t !== 'none')) sh.fill = L.clone(fmt.fill);
      if (fmt.line) sh.line = L.clone(fmt.line);
      if ('shadow' in fmt) sh.shadow = L.clone(fmt.shadow);
      if (fmt.run && sh.tx) L.txt.mapRange(sh.tx, L.txt.allRange(sh.tx), (r) => paintRun(r, fmt.run));
      if (fmt.para && sh.tx) for (const p of sh.tx.ps) p.pp = Object.assign({}, p.pp, fmt.para);
      if (fmt.wa && sh.type === 'wordart') Object.assign(sh.wa, L.clone(fmt.wa));
    });
    E.select([id]);
    return true;
  };

  /* ---------- master view ---------- */
  E.buildMaster = function () {
    const sl = L.pres.slides[E.idx] || L.pres.slides[0];
    const d = M.design(L.pres, sl);
    const titleKind = E.masterKind === 'title' && d.titleDeco;
    const deco = titleKind ? d.titleDeco : d.deco;
    const ph = d.ph;
    const mk = (type, r, paras) => {
      const s = M.makePlaceholder(Object.assign({ type }, r), d);
      s.id = 'mph-' + type;
      s.tx.ps = paras;
      s.tx.autofit = 'none';
      return s;
    };
    const shapes = deco.slice();
    if (titleKind) {
      shapes.push(mk('ctrTitle', ph.ctrTitle, [L.txt.para(L.t('Click to edit Master title style'))]));
      shapes.push(mk('subTitle', ph.subTitle, [L.txt.para(L.t('Click to edit Master subtitle style'))]));
    } else {
      shapes.push(mk('title', ph.title, [L.txt.para(L.t('Click to edit Master title style'))]));
      shapes.push(mk('body', ph.body, ['Click to edit Master text styles', 'Second level', 'Third level', 'Fourth level', 'Fifth level'].map((t, i) => L.txt.para(L.t(t), null, null, i))));
    }
    const ft = (type, text) => { const s = mk(type, ph[type], [L.txt.para(text)]); return s; };
    shapes.push(ft('dt', L.t('<date/time>')), ft('ftr', L.t('<footer>')), ft('sldNum', '<#>'));
    E.masterSlide = { id: 'master', layout: titleKind ? 'title' : 'text', design: d.id, shapes, hideMaster: true, bg: titleKind && d.titleBg ? d.titleBg : d.bg, anims: [], trans: {}, notes: '', isMaster: true, titleKind: !!titleKind };
  };
  /** fold master-view edits back into the design */
  E.syncMaster = function () {
    const ms = E.masterSlide;
    if (!ms) return;
    const d = L.pres.designs[ms.design];
    if (!d) return;
    const deco = ms.shapes.filter((s) => !String(s.id).startsWith('mph-'));
    if (ms.titleKind) d.titleDeco = deco; else d.deco = deco;
    for (const s of ms.shapes) {
      if (!String(s.id).startsWith('mph-')) continue;
      const type = s.ph.type;
      d.ph[type] = { x: L.round(s.x, 2), y: L.round(s.y, 2), w: L.round(s.w, 2), h: L.round(s.h, 2) };
      const cls = L.style.cls(s);
      if (type === 'dt' || type === 'ftr' || type === 'sldNum') {
        const p = s.tx.ps[0];
        if (p && p.rs[0]) { const own = L.txt.runProps(p.rs[0]); if (Object.keys(own).length) { d.footer.rPr = Object.assign({}, d.footer.rPr, own); p.rs.forEach((r) => { for (const k in own) delete r[k]; }); } }
        continue;
      }
      if (type === 'subTitle') { const p = s.tx.ps[0]; if (p && p.pp && p.pp.algn) d.subTitleAlgn = p.pp.algn; continue; }
      s.tx.ps.forEach((p, i) => {
        const lvl = cls === 'title' ? 0 : i;
        if (!d.tx[cls][lvl]) return;
        const own = p.rs[0] ? L.txt.runProps(p.rs[0]) : {};
        const lv = d.tx[cls][lvl];
        if (Object.keys(own).length) { lv.rPr = Object.assign({}, lv.rPr, own); delete lv.rPr.link; p.rs.forEach((r) => { for (const k in own) delete r[k]; }); }
        if (p.pp && Object.keys(p.pp).length) { Object.assign(lv, L.clone(p.pp)); if (type === 'ctrTitle' && p.pp.algn) d.ctrTitleAlgn = p.pp.algn; p.pp = {}; }
      });
    }
    if (ms.bg) { if (ms.titleKind) d.titleBg = ms.bg; else d.bg = ms.bg; }
  };
})();
