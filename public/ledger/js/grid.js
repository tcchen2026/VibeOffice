/* Ledger — grid controller: viewport, scrolling, selection, mouse & keyboard, objects layer. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, LY = L.layout, RD = L.render, O = L.ops, C = L.calc, F = L.formula;
  const { h } = L;
  const G = (L.grid = {});
  const MAXR = M.MAXR, MAXC = M.MAXC;

  /* ------------------------------------------------------------ state */
  G.wb = null;
  G.zoomOf = (sh) => (sh.view.zoom || 100) / 100;
  G.state = { marquee: null, hl: null, fillTo: null, dragTo: null, pages: null };
  let wrap, canvas, layer, vbar, hbar, inp;
  let dirty = true, raf = 0;

  G.sheet = () => G.wb.sheets[G.wb.active];
  /** per-sheet runtime view state (scroll / selection), kept on the sheet object */
  G.vs = (sh) => {
    sh = sh || G.sheet();
    if (!sh._vs) {
      const top = sh.view.freeze && sh.view.freeze.top ? sh.view.freeze.top : sh.view.top || { r: 0, c: 0 };
      const s = sh.view.sel;
      sh._vs = {
        scrollR: top.r || 0, scrollC: top.c || 0,
        sel: s ? { r: s.r, c: s.c, ranges: s.ranges.map((x) => Object.assign({}, x)), active: 0, anchor: { r: s.r, c: s.c } } : { r: 0, c: 0, ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], active: 0, anchor: { r: 0, c: 0 } },
      };
      const g = sh._vs.sel;
      g.active = Math.max(0, g.ranges.findIndex((x) => M.rangeContains(x, g.r, g.c)));
    }
    return sh._vs;
  };
  G.sel = () => G.vs().sel;
  /** the active range */
  G.range = () => { const s = G.sel(); return s.ranges[s.active] || s.ranges[0]; };
  G.ranges = () => G.sel().ranges;

  /* ------------------------------------------------------------ setup */
  G.init = function (wb) {
    wrap = L.$('#gridwrap');
    layer = L.$('#glayer');
    canvas = h('canvas', { class: 'gcv', 'aria-hidden': 'true' });
    wrap.insertBefore(canvas, layer);
    inp = h('textarea', { class: 'celledit', id: 'celledit', 'aria-label': 'Cell editor', spellcheck: 'false', autocomplete: 'off', autocapitalize: 'off', wrap: 'off', rows: '1' });
    inp.style.opacity = '0';
    inp.style.pointerEvents = 'none';
    wrap.appendChild(inp);
    G.input = inp;
    vbar = new Scrollbar(L.$('#vscroll'), 'v', (pos, how) => scrollTo('v', pos, how));
    hbar = new Scrollbar(L.$('#hscroll'), 'h', (pos, how) => scrollTo('h', pos, how));
    G.setWorkbook(wb);
    wrap.addEventListener('pointerdown', onPointerDown);
    wrap.addEventListener('pointermove', onHover);
    wrap.addEventListener('dblclick', onDblClick);
    wrap.addEventListener('wheel', onWheel, { passive: false });
    wrap.addEventListener('contextmenu', onContext);
    wrap.addEventListener('focus', () => G.focus());
    inp.addEventListener('keydown', onKey);
    inp.addEventListener('input', onInput);
    inp.addEventListener('compositionstart', () => { if (!L.editor.active) L.editor.begin('enter', ''); });
    inp.addEventListener('copy', (e) => { if (L.editor.active) return; e.preventDefault(); L.clip.copy(e, false); });
    inp.addEventListener('cut', (e) => { if (L.editor.active) return; e.preventDefault(); L.clip.copy(e, true); });
    inp.addEventListener('paste', (e) => { if (L.editor.active) return; e.preventDefault(); L.clip.paste(e); });
    new ResizeObserver(() => G.paint(true)).observe(wrap);
    G.focus();
  };
  G.setWorkbook = function (wb) {
    G.wb = wb;
    G.state.marquee = null;
    if (!wb.graph) C.rebuild(wb);
    G.paint(true);
    G.syncObjects(true);
  };
  G.focus = () => { if (inp && document.activeElement !== inp && !L.ui.dialogOpen()) { try { inp.focus({ preventScroll: true }); } catch (e) { inp.focus(); } } };
  G.refocus = G.focus;

  /* ------------------------------------------------------------ painting */
  G.paint = function (now) {
    dirty = true;
    if (now === true) { draw(); return; }
    if (!raf) raf = requestAnimationFrame(() => { raf = 0; if (dirty) draw(); });
  };
  G.view = function () {
    const sh = G.sheet(), vs = G.vs(sh);
    const W = Math.max(50, wrap.clientWidth - 17), H = Math.max(50, wrap.clientHeight);
    return { sh, zoom: G.zoomOf(sh), scrollR: vs.scrollR, scrollC: vs.scrollC, w: W, h: H, dpr: window.devicePixelRatio || 1 };
  };
  let lastFrame = null;
  G.frame = () => lastFrame || (draw(), lastFrame);
  function draw() {
    dirty = false;
    if (!G.wb) return;
    const sh = G.sheet();
    const v = G.view();
    const st = G.state;
    const state = { sel: G.sel(), editing: L.editor && L.editor.active ? { sh: L.editor.sheet, r: L.editor.r, c: L.editor.c } : null, marquee: st.marquee, hl: st.hl, fillTo: st.fillTo, dragTo: st.dragTo, pages: sh._pages || null, hideSel: st.hideSel };
    try { lastFrame = RD.draw(canvas, v, state); } catch (e) { console.error(e); }
    updateScrollbars(v);
    G.syncObjects();
    if (L.editor && L.editor.active) L.editor.place();
    else placeInput();
    L.bus.emit('grid-paint', lastFrame);
  }
  function updateScrollbars(v) {
    const sh = v.sh, fr = lastFrame, g = fr.g;
    const rowsVis = Math.max(1, fr.panes.length ? fr.panes[fr.panes.length - 1].rEnd - fr.sR : 1);
    const colsVis = Math.max(1, fr.panes.length ? fr.panes[fr.panes.length - 1].cEnd - fr.sC : 1);
    const vs = G.vs(sh);
    const totR = Math.min(MAXR, Math.max(sh.maxR + 1 + 2, vs.scrollR + rowsVis * 2, 100));
    const totC = Math.min(MAXC, Math.max(sh.maxC + 1 + 2, vs.scrollC + colsVis * 2, 26));
    vbar.set(vs.scrollR - fr.sR + fr.sR, rowsVis, totR, fr.sR);
    hbar.set(vs.scrollC, colsVis, totC, fr.sC);
    void g;
  }
  function scrollTo(axis, pos, how) {
    const sh = G.sheet(), vs = G.vs(sh), fr = lastFrame;
    const min = axis === 'v' ? fr.sR : fr.sC;
    const g = LY.geo(sh);
    if (axis === 'v') { let r = Math.max(min, Math.min(MAXR - 1, Math.round(pos))); r = g.rows.visible(r, how === 'up' ? -1 : 1); vs.scrollR = r; }
    else { let c = Math.max(min, Math.min(MAXC - 1, Math.round(pos))); c = g.cols.visible(c, how === 'up' ? -1 : 1); vs.scrollC = c; }
    G.paint();
    L.bus.emit('scrolled');
  }
  G.scrollBy = function (dr, dc) {
    const vs = G.vs(), fr = G.frame();
    const g = LY.geo(G.sheet());
    if (dr) { let r = vs.scrollR, k = Math.abs(dr); const d = dr > 0 ? 1 : -1; while (k > 0) { r += d; if (r < fr.sR) { r = fr.sR; break; } if (r >= MAXR) { r = MAXR - 1; break; } if (g.rows.size(r) > 0) k--; } vs.scrollR = Math.max(fr.sR, r); }
    if (dc) { let c = vs.scrollC, k = Math.abs(dc); const d = dc > 0 ? 1 : -1; while (k > 0) { c += d; if (c < fr.sC) { c = fr.sC; break; } if (c >= MAXC) { c = MAXC - 1; break; } if (g.cols.size(c) > 0) k--; } vs.scrollC = Math.max(fr.sC, c); }
    G.paint();
    L.bus.emit('scrolled');
  };
  /** make cell (r, c) visible in the scrolling pane */
  G.reveal = function (r, c) {
    const sh = G.sheet(), vs = G.vs(sh);
    draw();
    const fr = lastFrame, g = fr.g, z = fr.z;
    const main = fr.panes[fr.panes.length - 1];
    if (!main) return;
    if (r >= fr.sR) {
      if (r < vs.scrollR) vs.scrollR = r;
      else {
        const bottom = g.rows.pos(vs.scrollR) + main.h / z;
        if (g.rows.end(r) > bottom) {
          /* scroll so r is the last fully visible row */
          let top = r, acc = g.rows.size(r);
          while (top > fr.sR && acc + g.rows.size(top - 1) <= main.h / z) { top--; acc += g.rows.size(top); }
          vs.scrollR = Math.max(fr.sR, top);
        }
      }
    }
    if (c >= fr.sC) {
      if (c < vs.scrollC) vs.scrollC = c;
      else {
        const right = g.cols.pos(vs.scrollC) + main.w / z;
        if (g.cols.end(c) > right) {
          let left = c, acc = g.cols.size(c);
          while (left > fr.sC && acc + g.cols.size(left - 1) <= main.w / z) { left--; acc += g.cols.size(left); }
          vs.scrollC = Math.max(fr.sC, left);
        }
      }
    }
    G.paint();
  };

  /* ------------------------------------------------------------ scrollbars */
  class Scrollbar {
    constructor(el, orient, onScroll) {
      this.el = el; this.v = orient === 'v'; this.onScroll = onScroll;
      this.pos = 0; this.page = 10; this.total = 100; this.min = 0;
      const arrow = (d) => `<svg width="7" height="7" viewBox="0 0 7 7"><path d="${{ up: 'M0.5 5.5L3.5 1.5 6.5 5.5', dn: 'M0.5 1.5L3.5 5.5 6.5 1.5', lt: 'M5.5 0.5L1.5 3.5 5.5 6.5', rt: 'M1.5 0.5L5.5 3.5 1.5 6.5' }[d]}" fill="none" stroke="#4d6185" stroke-width="1.6"/></svg>`;
      this.b1 = h('button', { class: 'sb-btn ' + (this.v ? 'sb-up' : 'sb-lt'), type: 'button', tabindex: '-1', 'aria-label': this.v ? 'Scroll up' : 'Scroll left', html: arrow(this.v ? 'up' : 'lt') });
      this.b2 = h('button', { class: 'sb-btn ' + (this.v ? 'sb-dn' : 'sb-rt'), type: 'button', tabindex: '-1', 'aria-label': this.v ? 'Scroll down' : 'Scroll right', html: arrow(this.v ? 'dn' : 'rt') });
      this.thumb = h('div', { class: 'sb-thumb' }, h('i'));
      el.append(this.b1, this.thumb, this.b2);
      const repeat = (fn) => (e) => { e.preventDefault(); fn(); let t = setTimeout(function tick() { fn(); t = setTimeout(tick, 50); }, 350); const stop = () => { clearTimeout(t); window.removeEventListener('pointerup', stop); }; window.addEventListener('pointerup', stop); G.focus(); };
      this.b1.addEventListener('pointerdown', repeat(() => this.onScroll(this.pos - 1, 'up')));
      this.b2.addEventListener('pointerdown', repeat(() => this.onScroll(this.pos + 1, 'down')));
      el.addEventListener('pointerdown', (e) => {
        if (e.target !== el) return;
        e.preventDefault();
        const r = this.thumb.getBoundingClientRect();
        const before = this.v ? e.clientY < r.top : e.clientX < r.left;
        repeat(() => this.onScroll(this.pos + (before ? -1 : 1) * Math.max(1, this.page - 1), before ? 'up' : 'down'))(e);
      });
      this.thumb.addEventListener('pointerdown', (e) => {
        e.preventDefault(); e.stopPropagation();
        this.thumb.setPointerCapture(e.pointerId);
        const start = this.v ? e.clientY : e.clientX, p0 = this.pos;
        const track = this.trackLen(), tl = this.thumbLen();
        const span = Math.max(1, this.total - this.min);
        const mv = (ev) => { const d = (this.v ? ev.clientY : ev.clientX) - start; this.onScroll(p0 + (d / Math.max(1, track - tl)) * Math.max(1, span - this.page), 'drag'); if (L.ui.tipFor) L.ui.tipFor(this.thumb, this.v ? 'Row: ' + (this.pos + 1) : 'Column: ' + L.formula.colName(this.pos)); };
        const up = () => { this.thumb.removeEventListener('pointermove', mv); this.thumb.removeEventListener('pointerup', up); G.focus(); };
        this.thumb.addEventListener('pointermove', mv);
        this.thumb.addEventListener('pointerup', up);
      });
    }
    trackLen() { const r = this.el.getBoundingClientRect(); return (this.v ? r.height : r.width) - 34; }
    thumbLen() { const span = Math.max(1, this.total - this.min); return Math.max(9, Math.min(this.trackLen(), (this.trackLen() * this.page) / span)); }
    set(pos, page, total, min) {
      this.pos = pos; this.page = page; this.total = Math.max(total, pos + page); this.min = min || 0;
      const track = this.trackLen(), tl = this.thumbLen();
      const span = Math.max(1, this.total - this.min - this.page);
      const off = 17 + (track - tl) * Math.min(1, Math.max(0, (pos - this.min) / span));
      if (this.v) { this.thumb.style.top = off + 'px'; this.thumb.style.height = tl + 'px'; }
      else { this.thumb.style.left = off + 'px'; this.thumb.style.width = tl + 'px'; }
    }
  }

  /* ------------------------------------------------------------ selection */
  const norm = (a) => ({ r1: Math.min(a.r1, a.r2), c1: Math.min(a.c1, a.c2), r2: Math.max(a.r1, a.r2), c2: Math.max(a.c1, a.c2) });
  /** grow a range so merged areas are fully inside */
  G.expandMerges = function (sh, rg) {
    if (!sh.merges.length) return rg;
    let changed = true, x = Object.assign({}, rg), n = 0;
    while (changed && n++ < 50) {
      changed = false;
      for (const m of sh.merges) {
        if (M.rangesOverlap(m, x) && !(m.r1 >= x.r1 && m.r2 <= x.r2 && m.c1 >= x.c1 && m.c2 <= x.c2)) { x = M.rangeUnion(x, m); changed = true; }
      }
    }
    return x;
  };
  /** select a single cell (or the merged area containing it) */
  G.select = function (r, c, opts) {
    opts = opts || {};
    const sh = G.sheet(), s = G.sel();
    r = Math.max(0, Math.min(MAXR - 1, r)); c = Math.max(0, Math.min(MAXC - 1, c));
    const m = LY.mergeAt(sh, r, c);
    if (m) { r = m.r1; c = m.c1; }
    const rg = m ? { r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 } : { r1: r, c1: c, r2: r, c2: c };
    if (opts.add) { s.ranges.push(rg); s.active = s.ranges.length - 1; }
    else { s.ranges = [rg]; s.active = 0; }
    s.r = r; s.c = c; s.anchor = { r, c };
    if (!opts.noScroll) G.reveal(opts.revealR != null ? opts.revealR : r, c);
    G.changed();
  };
  /** extend the active range from the anchor to (r, c) */
  G.extendTo = function (r, c, opts) {
    const sh = G.sheet(), s = G.sel();
    r = Math.max(0, Math.min(MAXR - 1, r)); c = Math.max(0, Math.min(MAXC - 1, c));
    const a = s.anchor || { r: s.r, c: s.c };
    let rg = norm({ r1: a.r, c1: a.c, r2: r, c2: c });
    rg = G.expandMerges(sh, rg);
    s.ranges[s.active] = rg;
    s.extent = { r, c };
    if (!opts || !opts.noScroll) G.reveal(r, c);
    G.changed();
  };
  G.selectRange = function (rg, active, opts) {
    const s = G.sel();
    rg = norm(rg);
    const sh = G.sheet();
    rg = G.expandMerges(sh, rg);
    s.ranges = [rg]; s.active = 0;
    s.r = active ? active.r : rg.r1; s.c = active ? active.c : rg.c1;
    s.anchor = { r: s.r, c: s.c };
    if (!opts || !opts.noScroll) G.reveal(s.r, s.c);
    G.changed();
  };
  G.selectAll = function () {
    const s = G.sel();
    s.ranges = [{ r1: 0, c1: 0, r2: MAXR - 1, c2: MAXC - 1 }]; s.active = 0;
    s.anchor = { r: s.r, c: s.c };
    G.changed();
  };
  G.changed = function () { G.paint(); L.bus.emit('selection'); };
  G.isWholeCols = (rg) => rg.r1 === 0 && rg.r2 >= MAXR - 1;
  G.isWholeRows = (rg) => rg.c1 === 0 && rg.c2 >= MAXC - 1;

  /* data edges for Ctrl+Arrow */
  const filled = (sh, r, c) => { const cl = sh.get(r, c); return !!(cl && ((cl.v != null && cl.v !== '') || cl.f != null)); };
  G.edge = function (r, c, dr, dc) {
    const sh = G.sheet();
    const g = LY.geo(sh);
    const step = (x, y) => [x + dr, y + dc];
    const inside = (x, y) => x >= 0 && y >= 0 && x < MAXR && y < MAXC;
    const visible = (x, y) => g.rows.size(x) > 0 && g.cols.size(y) > 0;
    let [nr, nc] = step(r, c);
    if (!inside(nr, nc)) return [r, c];
    const limR = dr > 0 ? Math.max(sh.maxR, r) + 1 : -1, limC = dc > 0 ? Math.max(sh.maxC, c) + 1 : -1;
    const beyond = (x, y) => (dr > 0 && x > limR) || (dc > 0 && y > limC);
    if (filled(sh, r, c) && filled(sh, nr, nc)) {
      /* run to the last filled cell */
      let pr = nr, pc = nc;
      while (true) { const [xr, xc] = step(pr, pc); if (!inside(xr, xc) || !filled(sh, xr, xc)) break; pr = xr; pc = xc; }
      return [pr, pc];
    }
    /* skip blanks to the next filled cell (or the sheet edge) */
    let pr = nr, pc = nc;
    while (inside(pr, pc) && !(filled(sh, pr, pc) && visible(pr, pc))) {
      if (beyond(pr, pc)) return [dr > 0 ? MAXR - 1 : pr, dc > 0 ? MAXC - 1 : pc];
      const [xr, xc] = step(pr, pc);
      if (!inside(xr, xc)) return [pr, pc];
      pr = xr; pc = xc;
    }
    return [pr, pc];
  };
  /** move the active cell by (dr, dc), skipping hidden rows / columns and merged areas */
  G.move = function (dr, dc, opts) {
    opts = opts || {};
    const sh = G.sheet(), s = G.sel(), g = LY.geo(sh);
    let base = opts.extend ? s.extent || { r: s.r, c: s.c } : { r: s.r, c: s.c };
    if (!opts.extend) { const m = LY.mergeAt(sh, base.r, base.c); if (m) base = { r: dr > 0 ? m.r2 : m.r1, c: dc > 0 ? m.c2 : m.c1 }; }
    let r = base.r, c = base.c;
    if (opts.ctrl) [r, c] = G.edge(r, c, Math.sign(dr), Math.sign(dc));
    else {
      if (dr) { let k = Math.abs(dr); while (k > 0) { r += Math.sign(dr); if (r < 0) { r = 0; break; } if (r >= MAXR) { r = MAXR - 1; break; } if (g.rows.size(r) > 0) k--; } }
      if (dc) { let k = Math.abs(dc); while (k > 0) { c += Math.sign(dc); if (c < 0) { c = 0; break; } if (c >= MAXC) { c = MAXC - 1; break; } if (g.cols.size(c) > 0) k--; } }
    }
    if (opts.extend) G.extendTo(r, c);
    else G.select(r, c);
  };
  /** Enter / Tab inside a multi-cell selection cycles through it */
  G.moveInSelection = function (dr, dc) {
    const s = G.sel();
    const rg = s.ranges[s.active];
    const multi = s.ranges.length > 1 || rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
    const sh = G.sheet();
    if (!multi || (LY.mergeAt(sh, s.r, s.c) && s.ranges.length === 1 && (() => { const m = LY.mergeAt(sh, s.r, s.c); return m.r1 === rg.r1 && m.r2 === rg.r2 && m.c1 === rg.c1 && m.c2 === rg.c2; })())) { G.move(dr, dc); return; }
    let r = s.r, c = s.c;
    const m = LY.mergeAt(sh, r, c);
    for (let guard = 0; guard < 100000; guard++) {
      if (dr) { r += dr; if (r > rg.r2) { r = rg.r1; c++; if (c > rg.c2) { c = rg.c1; nextArea(1); } } else if (r < rg.r1) { r = rg.r2; c--; if (c < rg.c1) { c = rg.c2; nextArea(-1); } } }
      else { c += dc; if (c > rg.c2) { c = rg.c1; r++; if (r > rg.r2) { r = rg.r1; nextArea(1); } } else if (c < rg.c1) { c = rg.c2; r--; if (r < rg.r1) { r = rg.r2; nextArea(-1); } } }
      const m2 = LY.mergeAt(sh, r, c);
      if (m2 && (m2.r1 !== r || m2.c1 !== c)) continue;
      if (m && m2 === m) continue;
      break;
    }
    function nextArea(d) { if (s.ranges.length > 1) { s.active = (s.active + d + s.ranges.length) % s.ranges.length; const a = s.ranges[s.active]; r = d > 0 ? a.r1 : a.r2; c = d > 0 ? a.c1 : a.c2; } }
    s.r = r; s.c = c;
    G.reveal(r, c);
    G.changed();
  };

  /* ------------------------------------------------------------ hidden input placement */
  function placeInput() {
    const fr = lastFrame;
    if (!fr) return;
    const s = G.sel();
    const p = RD.paneOf(fr, s.r, s.c) || fr.panes[fr.panes.length - 1];
    if (!p) return;
    const x = RD.colX(fr, p, s.c), y = RD.rowY(fr, p, s.r);
    inp.style.left = Math.max(0, Math.min(fr.panes.length ? x : 0, wrap.clientWidth - 40)) + 'px';
    inp.style.top = Math.max(0, Math.min(y, wrap.clientHeight - 20)) + 'px';
    inp.style.width = '20px'; inp.style.height = '18px';
  }
  G.cellRect = function (r, c, sh) {
    const fr = G.frame();
    sh = sh || G.sheet();
    const m = LY.mergeAt(sh, r, c);
    const rg = m || { r1: r, c1: c, r2: r, c2: c };
    const p = RD.paneOf(fr, rg.r1, rg.c1) || fr.panes[fr.panes.length - 1];
    const x = RD.colX(fr, p, rg.c1), y = RD.rowY(fr, p, rg.r1);
    return { x, y, w: RD.colX(fr, p, rg.c2 + 1) - x, h: RD.rowY(fr, p, rg.r2 + 1) - y, pane: p, fr };
  };

  /* ------------------------------------------------------------ mouse */
  let drag = null;
  const pt = (e) => { const r = wrap.getBoundingClientRect(); return { x: e.clientX - r.left, y: e.clientY - r.top }; };
  /** what is under the pointer (adds resize / fill / border zones) */
  function zone(p) {
    const fr = G.frame();
    const hit = RD.hit(fr, p.x, p.y);
    const sh = G.sheet(), g = fr.g, z = fr.z;
    if (hit.area === 'colhdr') {
      /* near a column boundary → resize (the column to the left of the boundary) */
      const pane = hit.pane;
      const x1 = RD.colX(fr, pane, hit.c), x2 = RD.colX(fr, pane, hit.c + 1);
      if (Math.abs(p.x - x2) <= 3) { let c = hit.c; return { kind: 'colres', c, x: x2 }; }
      if (Math.abs(p.x - x1) <= 3 && hit.c > 0) { let c = hit.c - 1; while (c > 0 && g.cols.size(c) === 0 && Math.abs(p.x - x1) <= 2) { if (p.x > x1) break; c--; } return { kind: 'colres', c, x: x1 }; }
      return { kind: 'colhdr', c: hit.c };
    }
    if (hit.area === 'rowhdr') {
      const pane = hit.pane;
      const y1 = RD.rowY(fr, pane, hit.r), y2 = RD.rowY(fr, pane, hit.r + 1);
      if (Math.abs(p.y - y2) <= 2) return { kind: 'rowres', r: hit.r, y: y2 };
      if (Math.abs(p.y - y1) <= 2 && hit.r > 0) return { kind: 'rowres', r: hit.r - 1, y: y1 };
      return { kind: 'rowhdr', r: hit.r };
    }
    if (hit.area === 'corner') return { kind: 'corner' };
    if (hit.area !== 'cell') return { kind: 'none' };
    /* Page Break Preview: the blue break lines can be dragged */
    if (sh.view.pageBreakPreview && sh._pages && !G.protectedSheet()) {
      for (const br of sh._pages.rows || []) { const y = RD.rowY(fr, hit.pane, br.r); if (br.r > 0 && Math.abs(p.y - y) <= 3) return { kind: 'pbreak', axis: 'r', at: br.r, r: hit.r, c: hit.c }; }
      for (const br of sh._pages.cols || []) { const x = RD.colX(fr, hit.pane, br.c); if (br.c > 0 && Math.abs(p.x - x) <= 3) return { kind: 'pbreak', axis: 'c', at: br.c, r: hit.r, c: hit.c }; }
    }
    /* fill handle and selection border */
    const s = G.sel();
    if (s.ranges.length === 1 && !(L.editor && L.editor.active)) {
      const rg = s.ranges[0];
      const q = RD.rangeRect(fr, rg, RD.paneOf(fr, Math.min(Math.max(rg.r2, hit.pane.r0), hit.pane.rEnd), Math.min(Math.max(rg.c2, hit.pane.c0), hit.pane.cEnd)) || hit.pane);
      const fx = q.x + q.w - 1, fy = q.y + q.h - 1;
      if (Math.abs(p.x - fx) <= 4 && Math.abs(p.y - fy) <= 4 && !G.protectedSheet()) return { kind: 'fill', r: hit.r, c: hit.c };
      const onV = (Math.abs(p.x - q.x) <= 2 || Math.abs(p.x - (q.x + q.w - 1)) <= 2) && p.y >= q.y - 2 && p.y <= q.y + q.h + 1;
      const onH = (Math.abs(p.y - q.y) <= 2 || Math.abs(p.y - (q.y + q.h - 1)) <= 2) && p.x >= q.x - 2 && p.x <= q.x + q.w + 1;
      if ((onV || onH) && !G.isWholeCols(rg) && !G.isWholeRows(rg)) return { kind: 'border', r: hit.r, c: hit.c };
    }
    void sh; void z;
    return { kind: 'cell', r: hit.r, c: hit.c, pane: hit.pane };
  }
  function onHover(e) {
    if (drag || e.pointerType === 'touch') return;
    if (e.target !== canvas && e.target !== wrap && e.target !== layer) { setCursor(null); return; }
    const z = zone(pt(e));
    const map = { colres: 'cur-colres', rowres: 'cur-rowres', colhdr: 'cur-down', rowhdr: 'cur-right', fill: 'cur-fill', border: 'cur-move', corner: 'cur-arrow', none: 'cur-arrow' };
    let cls = map[z.kind] || 'cur-cell';
    if (z.kind === 'pbreak') cls = z.axis === 'r' ? 'cur-rowres' : 'cur-colres';
    if (z.kind === 'cell') {
      const sh = G.sheet();
      const link = sh.links.find((l) => M.rangeContains(l.ref, z.r, z.c));
      if (link && !(L.editor && L.editor.active)) cls = 'cur-link';
      showCommentHover(z.r, z.c);
    } else showCommentHover(-1, -1);
    if (G.painting) cls = 'painting';
    setCursor(cls);
  }
  let curCls = 'cur-cell';
  function setCursor(cls) {
    cls = cls || 'cur-arrow';
    if (cls === curCls) return;
    wrap.classList.remove(curCls); wrap.classList.add(cls); curCls = cls;
  }
  function onPointerDown(e) {
    if (e.target.closest('.gobj, .af-btn, .dv-btn, .cmt-box, .outline-btn, .sb, .celledit.on')) return;
    if (e.button === 2) {
      /* right click keeps a selection that contains the cell */
      const z = zone(pt(e));
      const s = G.sel();
      if (z.kind === 'cell' && !s.ranges.some((rg) => M.rangeContains(rg, z.r, z.c))) G.select(z.r, z.c, { noScroll: true });
      if (z.kind === 'colhdr' && !s.ranges.some((rg) => G.isWholeCols(rg) && z.c >= rg.c1 && z.c <= rg.c2)) G.selectRange({ r1: 0, c1: z.c, r2: MAXR - 1, c2: z.c }, { r: G.vs().scrollR, c: z.c }, { noScroll: true });
      if (z.kind === 'rowhdr' && !s.ranges.some((rg) => G.isWholeRows(rg) && z.r >= rg.r1 && z.r <= rg.r2)) G.selectRange({ r1: z.r, c1: 0, r2: z.r, c2: MAXC - 1 }, { r: z.r, c: G.vs().scrollC }, { noScroll: true });
      return;
    }
    if (e.button !== 0) return;
    e.preventDefault();
    L.ui.closeMenus();
    G.deselectObject();
    const p = pt(e);
    if (e.pointerType === 'touch') { startTouch(e, p); return; }
    const z = zone(p);
    const ed = L.editor;
    /* formula point mode: clicking a cell inserts a reference */
    if (ed && ed.active && ed.canPoint() && (z.kind === 'cell' || z.kind === 'colhdr' || z.kind === 'rowhdr' || z.kind === 'border' || z.kind === 'fill')) {
      const r = z.r != null ? z.r : 0, c = z.c != null ? z.c : 0;
      const rg = z.kind === 'colhdr' ? { r1: 0, c1: z.c, r2: MAXR - 1, c2: z.c } : z.kind === 'rowhdr' ? { r1: z.r, c1: 0, r2: z.r, c2: MAXC - 1 } : { r1: r, c1: c, r2: r, c2: c };
      ed.pointStart(rg, e.ctrlKey || e.metaKey);
      beginDrag(e, { kind: 'point', anchor: rg, whole: z.kind === 'colhdr' ? 'c' : z.kind === 'rowhdr' ? 'r' : null });
      return;
    }
    if (ed && ed.active) { if (!ed.commit()) return; }
    G.focus();
    const s = G.sel();
    const sh = G.sheet();
    switch (z.kind) {
      case 'corner': G.selectAll(); return;
      case 'colres': beginDrag(e, { kind: 'colres', c: z.c, x0: p.x, w0: LY.geo(sh).cols.size(z.c) * G.zoomOf(sh) }); return;
      case 'rowres': beginDrag(e, { kind: 'rowres', r: z.r, y0: p.y, h0: LY.geo(sh).rows.size(z.r) * G.zoomOf(sh) }); return;
      case 'colhdr': {
        const top = G.vs().scrollR;
        if (e.shiftKey) { const a = s.anchor; s.ranges[s.active] = { r1: 0, c1: Math.min(a.c, z.c), r2: MAXR - 1, c2: Math.max(a.c, z.c) }; }
        else if (e.ctrlKey || e.metaKey) { s.ranges.push({ r1: 0, c1: z.c, r2: MAXR - 1, c2: z.c }); s.active = s.ranges.length - 1; s.r = top; s.c = z.c; s.anchor = { r: top, c: z.c }; }
        else { s.ranges = [{ r1: 0, c1: z.c, r2: MAXR - 1, c2: z.c }]; s.active = 0; s.r = top; s.c = z.c; s.anchor = { r: top, c: z.c }; }
        G.changed();
        beginDrag(e, { kind: 'cols' });
        return;
      }
      case 'rowhdr': {
        const left = G.vs().scrollC;
        if (e.shiftKey) { const a = s.anchor; s.ranges[s.active] = { r1: Math.min(a.r, z.r), c1: 0, r2: Math.max(a.r, z.r), c2: MAXC - 1 }; }
        else if (e.ctrlKey || e.metaKey) { s.ranges.push({ r1: z.r, c1: 0, r2: z.r, c2: MAXC - 1 }); s.active = s.ranges.length - 1; s.r = z.r; s.c = left; s.anchor = { r: z.r, c: left }; }
        else { s.ranges = [{ r1: z.r, c1: 0, r2: z.r, c2: MAXC - 1 }]; s.active = 0; s.r = z.r; s.c = left; s.anchor = { r: z.r, c: left }; }
        G.changed();
        beginDrag(e, { kind: 'rows' });
        return;
      }
      case 'pbreak': beginDrag(e, { kind: 'pbreak', axis: z.axis, at: z.at }); return;
      case 'fill': beginDrag(e, { kind: 'fill', src: Object.assign({}, s.ranges[0]) }); return;
      case 'border': beginDrag(e, { kind: 'move', src: Object.assign({}, s.ranges[0]), grab: { r: z.r, c: z.c } }); return;
      case 'cell': {
        if (G.painting) { beginDrag(e, { kind: 'paint' }); G.select(z.r, z.c, { noScroll: true }); return; }
        if (e.shiftKey) G.extendTo(z.r, z.c, { noScroll: true });
        else G.select(z.r, z.c, { add: e.ctrlKey || e.metaKey, noScroll: true });
        const link = !e.shiftKey && !e.ctrlKey && sh.links.find((l) => M.rangeContains(l.ref, z.r, z.c));
        beginDrag(e, { kind: 'cells', link, start: z });
        /* data validation drop-down button */
        G.syncObjects();
        return;
      }
      default:
    }
  }
  function beginDrag(e, d) {
    drag = d;
    drag.pid = e.pointerId;
    try { wrap.setPointerCapture(e.pointerId); } catch (err) { /* ignore */ }
    drag.moved = false;
    drag.x0 = drag.x0 != null ? drag.x0 : pt(e).x; drag.y0 = drag.y0 != null ? drag.y0 : pt(e).y;
    wrap.addEventListener('pointermove', onDragMove);
    wrap.addEventListener('pointerup', onDragEnd);
    wrap.addEventListener('pointercancel', onDragEnd);
  }
  let autoTimer = null;
  function autoScroll(p, fn) {
    clearInterval(autoTimer);
    const fr = G.frame();
    const main = fr.panes[fr.panes.length - 1];
    let dr = 0, dc = 0;
    if (p.y > fr.hdrH + (main ? main.y + main.h - fr.hdrH : 0) - 2) dr = 1; else if (p.y < (main ? main.y : fr.hdrH) && G.vs().scrollR > fr.sR) dr = -1;
    if (p.x > (main ? main.x + main.w : 0) - 2) dc = 1; else if (p.x < (main ? main.x : fr.hdrW) && G.vs().scrollC > fr.sC) dc = -1;
    if (!dr && !dc) return;
    autoTimer = setInterval(() => { G.scrollBy(dr, dc); fn(); }, 60);
  }
  function cellFromPoint(p) {
    const fr = G.frame();
    const x = Math.max(fr.hdrW + 1, Math.min(p.x, fr.hdrW + (G.view().w - fr.hdrW) - 1));
    const y = Math.max(fr.hdrH + 1, Math.min(p.y, G.view().h - 1));
    const hit = RD.hit(fr, x, y);
    return { r: hit.r != null ? hit.r : 0, c: hit.c != null ? hit.c : 0 };
  }
  function onDragMove(e) {
    if (!drag) return;
    const p = pt(e);
    if (!drag.moved && Math.abs(p.x - drag.x0) < 3 && Math.abs(p.y - drag.y0) < 3) return;
    drag.moved = true;
    const sh = G.sheet(), s = G.sel(), fr = G.frame();
    const step = () => {
      const q = cellFromPoint(drag.last || p);
      switch (drag.kind) {
        case 'cells': case 'paint': G.extendTo(q.r, q.c, { noScroll: true }); break;
        case 'cols': { const a = s.anchor; const hit = RD.hit(fr, Math.max(fr.hdrW + 1, (drag.last || p).x), 2); const c = hit.c != null ? hit.c : q.c; s.ranges[s.active] = { r1: 0, c1: Math.min(a.c, c), r2: MAXR - 1, c2: Math.max(a.c, c) }; G.changed(); break; }
        case 'rows': { const a = s.anchor; const hit = RD.hit(fr, 2, Math.max(fr.hdrH + 1, (drag.last || p).y)); const r = hit.r != null ? hit.r : q.r; s.ranges[s.active] = { r1: Math.min(a.r, r), c1: 0, r2: Math.max(a.r, r), c2: MAXC - 1 }; G.changed(); break; }
        case 'point': {
          const a = drag.anchor;
          let rg = drag.whole === 'c' ? { r1: 0, c1: Math.min(a.c1, q.c), r2: MAXR - 1, c2: Math.max(a.c2, q.c) } : drag.whole === 'r' ? { r1: Math.min(a.r1, q.r), c1: 0, r2: Math.max(a.r2, q.r), c2: MAXC - 1 } : norm({ r1: a.r1, c1: a.c1, r2: q.r, c2: q.c });
          L.editor.pointMove(rg);
          break;
        }
        case 'fill': {
          const src = drag.src;
          /* fill extends along the axis with the larger movement */
          const dr = q.r > src.r2 ? q.r - src.r2 : q.r < src.r1 ? q.r - src.r1 : 0;
          const dc = q.c > src.c2 ? q.c - src.c2 : q.c < src.c1 ? q.c - src.c1 : 0;
          let to;
          if (Math.abs(dr) >= Math.abs(dc) && dr) to = dr > 0 ? { r1: src.r1, c1: src.c1, r2: q.r, c2: src.c2 } : { r1: q.r, c1: src.c1, r2: src.r2, c2: src.c2 };
          else if (dc) to = dc > 0 ? { r1: src.r1, c1: src.c1, r2: src.r2, c2: q.c } : { r1: src.r1, c1: q.c, r2: src.r2, c2: src.c2 };
          else if (q.r < src.r2 && q.r >= src.r1 && src.r2 > src.r1) to = { r1: src.r1, c1: src.c1, r2: q.r, c2: src.c2, shrink: true };
          else if (q.c < src.c2 && q.c >= src.c1 && src.c2 > src.c1) to = { r1: src.r1, c1: src.c1, r2: src.r2, c2: q.c, shrink: true };
          else to = null;
          G.state.fillTo = to;
          drag.to = to;
          if (to) setStatusTip(fillTip(sh, src, to, e.ctrlKey));
          G.paint();
          break;
        }
        case 'move': {
          const src = drag.src;
          const dr = q.r - drag.grab.r, dc = q.c - drag.grab.c;
          const to = { r1: Math.max(0, src.r1 + dr), c1: Math.max(0, src.c1 + dc), r2: Math.max(0, src.r1 + dr) + src.r2 - src.r1, c2: Math.max(0, src.c1 + dc) + src.c2 - src.c1 };
          G.state.dragTo = to; drag.to = to;
          setStatusTip(F.rangeName(to));
          G.paint();
          break;
        }
        default:
      }
    };
    drag.last = p;
    if (drag.kind === 'colres') {
      const w = Math.max(0, drag.w0 + p.x - drag.x0);
      drag.w = w;
      const mdw = M.mdw(sh.wb);
      setStatusTip('Width: ' + M.pxToChars(w / G.zoomOf(sh), mdw).toFixed(2) + ' (' + Math.round(w / G.zoomOf(sh)) + ' pixels)', p);
      showGuide('v', p.x);
      return;
    }
    if (drag.kind === 'pbreak') {
      /* snap to the nearest row / column boundary */
      const hit = RD.hit(fr, Math.max(fr.hdrW + 1, p.x), Math.max(fr.hdrH + 1, p.y));
      if (hit.area !== 'cell') return;
      if (drag.axis === 'r') { const y1 = RD.rowY(fr, hit.pane, hit.r), y2 = RD.rowY(fr, hit.pane, hit.r + 1); drag.to = p.y - y1 < y2 - p.y ? hit.r : hit.r + 1; showGuide('h', RD.rowY(fr, hit.pane, drag.to)); setStatusTip('Page break before row ' + (drag.to + 1), p); }
      else { const x1 = RD.colX(fr, hit.pane, hit.c), x2 = RD.colX(fr, hit.pane, hit.c + 1); drag.to = p.x - x1 < x2 - p.x ? hit.c : hit.c + 1; showGuide('v', RD.colX(fr, hit.pane, drag.to)); setStatusTip('Page break before column ' + F.colName(drag.to), p); }
      return;
    }
    if (drag.kind === 'rowres') {
      const hpx = Math.max(0, drag.h0 + p.y - drag.y0);
      drag.h = hpx;
      setStatusTip('Height: ' + (Math.round((hpx / G.zoomOf(sh)) * 0.75 * 4) / 4).toFixed(2) + ' (' + Math.round(hpx / G.zoomOf(sh)) + ' pixels)', p);
      showGuide('h', p.y);
      return;
    }
    step();
    autoScroll(p, step);
  }
  function fillTip(sh, src, to, ctrl) {
    /* the value the last filled cell will get */
    try {
      const dest = { r1: Math.min(src.r1, to.r1), c1: Math.min(src.c1, to.c1), r2: Math.max(src.r2, to.r2), c2: Math.max(src.c2, to.c2) };
      const r = to.r2 > src.r2 ? to.r2 : to.r1 < src.r1 ? to.r1 : src.r1, c = to.c2 > src.c2 ? to.c2 : to.c1 < src.c1 ? to.c1 : src.c1;
      void dest; void ctrl;
      return F.cellName(r, c);
    } catch (e) { return ''; }
  }
  let guide = null, tipEl = null;
  function showGuide(axis, pos) {
    if (!guide) { guide = h('div', { style: 'position:absolute;background:#000;z-index:30;pointer-events:none' }); wrap.appendChild(guide); }
    guide.hidden = false;
    if (axis === 'v') Object.assign(guide.style, { left: pos + 'px', top: '0', width: '1px', height: '100%' });
    else Object.assign(guide.style, { top: pos + 'px', left: '0', height: '1px', width: '100%' });
  }
  function setStatusTip(t, p) {
    if (!tipEl) { tipEl = h('div', { class: 'tooltip', style: 'position:absolute' }); wrap.appendChild(tipEl); }
    tipEl.hidden = !t;
    tipEl.textContent = t;
    const q = p || drag.last || { x: 0, y: 0 };
    tipEl.style.left = (q.x + 14) + 'px'; tipEl.style.top = (q.y + 18) + 'px';
  }
  function onDragEnd(e) {
    clearInterval(autoTimer);
    wrap.removeEventListener('pointermove', onDragMove);
    wrap.removeEventListener('pointerup', onDragEnd);
    wrap.removeEventListener('pointercancel', onDragEnd);
    if (guide) guide.hidden = true;
    if (tipEl) tipEl.hidden = true;
    const d = drag;
    drag = null;
    if (!d) return;
    const sh = G.sheet();
    try {
      switch (d.kind) {
        case 'pbreak': {
          if (!d.moved || d.to == null || d.to === d.at) break;
          /* the dragged break becomes a manual break at its new place (dragging it to the edge removes it) */
          const pr = JSON.parse(JSON.stringify(sh.print));
          const key = d.axis === 'r' ? 'rowBreaks' : 'colBreaks';
          const list = (pr[key] || []).filter((b) => b !== d.at);
          if (d.to > 0) list.push(d.to);
          pr[key] = Array.from(new Set(list)).sort((a, b) => a - b);
          O.tx(sh.wb, 'Move Page Break', () => O.setPrint(sh, pr));
          if (L.print) L.print.computeBreaks(sh);
          G.paint();
          break;
        }
        case 'colres': {
          if (!d.moved || d.w == null) break;
          const s = G.sel();
          const cols = colsForResize(s, d.c);
          const w = M.pxToWidth(Math.round(d.w / G.zoomOf(sh)), M.mdw(sh.wb));
          if (G.guardProtect('formatColumns')) break;
          O.setColWidth(sh, cols, Math.round(d.w) <= 1 ? 0 : M.pxToChars(Math.round(d.w / G.zoomOf(sh)), M.mdw(sh.wb)) > 0 ? Math.round(w * 256) / 256 : 0);
          break;
        }
        case 'rowres': {
          if (!d.moved || d.h == null) break;
          const s = G.sel();
          const rows = rowsForResize(s, d.r);
          if (G.guardProtect('formatRows')) break;
          const pt = Math.round((d.h / G.zoomOf(sh)) * 0.75 * 4) / 4;
          O.setRowHeight(sh, rows, pt <= 0.5 ? 0 : Math.min(409, pt));
          break;
        }
        case 'fill': {
          G.state.fillTo = null;
          if (!d.moved || !d.to) { G.paint(); break; }
          if (G.guardProtect()) break;
          const src = d.src, to = d.to;
          if (to.shrink) { O.clear(sh, [to.r2 < src.r2 ? { r1: to.r2 + 1, c1: src.c1, r2: src.r2, c2: src.c2 } : { r1: src.r1, c1: to.c2 + 1, r2: src.r2, c2: src.c2 }], 'contents'); G.selectRange({ r1: src.r1, c1: src.c1, r2: to.r2, c2: to.c2 }); break; }
          const dest = { r1: Math.min(src.r1, to.r1), c1: Math.min(src.c1, to.c1), r2: Math.max(src.r2, to.r2), c2: Math.max(src.c2, to.c2) };
          const single = src.r1 === src.r2 && src.c1 === src.c2;
          const mode = (e.ctrlKey || e.metaKey) ? (single ? 'linear' : 'copy') : 'auto';
          O.autoFill(sh, src, dest, mode);
          G.selectRange(dest, { r: src.r1, c: src.c1 }, { noScroll: true });
          G.lastFill = { src, dest };
          L.bus.emit('autofill', { src, dest });
          break;
        }
        case 'move': {
          G.state.dragTo = null;
          if (!d.moved || !d.to) { G.paint(); break; }
          if (G.guardProtect()) break;
          const to = d.to;
          if (to.r1 === d.src.r1 && to.c1 === d.src.c1) { G.paint(); break; }
          L.clip.moveRange(sh, d.src, to, e.ctrlKey || e.metaKey);
          G.selectRange(to);
          break;
        }
        case 'point': L.editor.pointEnd(); break;
        case 'paint': if (G.painting) L.app.finishPaint(); break;
        case 'cells': {
          if (!d.moved && d.link && !e.shiftKey) {
            /* a plain click on a hyperlink follows it */
            L.app.followLink(d.link);
          }
          break;
        }
        default:
      }
    } catch (err) { L.app.error(err); }
    G.paint();
    G.focus();
  }
  function colsForResize(s, c) {
    /* resizing a selected column resizes every selected column */
    for (const rg of s.ranges) if (G.isWholeCols(rg) && c >= rg.c1 && c <= rg.c2) { const out = []; for (const x of s.ranges) if (G.isWholeCols(x)) for (let k = x.c1; k <= x.c2; k++) out.push(k); return out; }
    return [c];
  }
  function rowsForResize(s, r) {
    for (const rg of s.ranges) if (G.isWholeRows(rg) && r >= rg.r1 && r <= rg.r2) { const out = []; for (const x of s.ranges) if (G.isWholeRows(x)) for (let k = x.r1; k <= x.r2; k++) out.push(k); return out; }
    return [r];
  }
  function onDblClick(e) {
    if (e.target.closest('.gobj, .af-btn, .dv-btn, .cmt-box, .sb')) return;
    const z = zone(pt(e));
    const sh = G.sheet();
    try {
      if (z.kind === 'colres') { if (G.guardProtect('formatColumns')) return; const s = G.sel(); O.autofitCols(sh, colsForResize(s, z.c)); return; }
      if (z.kind === 'rowres') { if (G.guardProtect('formatRows')) return; const s = G.sel(); O.autofitRows(sh, rowsForResize(s, z.r)); return; }
      if (z.kind === 'fill') { L.app.fillDownToData(); return; }
      if (z.kind === 'border') {
        /* jump to the edge of the data in that direction */
        const s = G.sel(), rg = s.ranges[0];
        const q = G.cellRect(rg.r1, rg.c1);
        const p = pt(e);
        const dir = Math.abs(p.y - q.y) <= 3 ? [-1, 0] : p.y > q.y + q.h - 4 ? [1, 0] : Math.abs(p.x - q.x) <= 3 ? [0, -1] : [0, 1];
        G.move(dir[0], dir[1], { ctrl: true });
        return;
      }
      if (z.kind === 'cell') {
        if (sh.comments.get(M.key(z.r, z.c)) && L.app.editComment && e.altKey) { L.app.editComment(z.r, z.c); return; }
        L.editor.begin('edit', null, { caretFromPoint: pt(e) });
      }
    } catch (err) { L.app.error(err); }
  }
  function onWheel(e) {
    if (e.target.closest('.dlg, .menu, .gobj.chart.editing')) return;
    e.preventDefault();
    if (e.ctrlKey || e.metaKey) {
      const sh = G.sheet();
      const z = Math.max(10, Math.min(400, (sh.view.zoom || 100) + (e.deltaY < 0 ? 15 : -15)));
      G.setZoom(z);
      return;
    }
    let dy = e.deltaY, dx = e.deltaX;
    if (e.shiftKey && !dx) { dx = dy; dy = 0; }
    const unit = e.deltaMode === 1 ? 1 : e.deltaMode === 2 ? 20 : 1 / 33;
    G._wy = (G._wy || 0) + dy * unit; G._wx = (G._wx || 0) + dx * unit;
    const rows = Math.trunc(G._wy), cols = Math.trunc(G._wx);
    if (rows) { G._wy -= rows; G.scrollBy(rows * (e.deltaMode === 1 ? 1 : 1), 0); }
    if (cols) { G._wx -= cols; G.scrollBy(0, cols); }
  }
  G.setZoom = function (z) {
    const sh = G.sheet();
    sh.view.zoom = Math.max(10, Math.min(400, Math.round(z)));
    G.paint();
    L.bus.emit('zoom');
  };
  /* touch: pan to scroll, tap to select, long press for the context menu */
  function startTouch(e, p) {
    const sh = G.sheet();
    const z0 = zone(p);
    let last = p, acc = { x: 0, y: 0 }, moved = false;
    const fr = G.frame();
    const g = LY.geo(sh);
    const lp = setTimeout(() => { if (!moved) { if (z0.kind === 'cell') G.select(z0.r, z0.c, { noScroll: true }); L.app.contextMenu({ clientX: e.clientX, clientY: e.clientY, preventDefault() {} }, z0); } }, 600);
    const mv = (ev) => {
      const q = pt(ev);
      if (Math.abs(q.x - p.x) + Math.abs(q.y - p.y) > 6) { moved = true; clearTimeout(lp); }
      if (!moved) return;
      acc.y += last.y - q.y; acc.x += last.x - q.x;
      last = q;
      const rh = g.rows.size(G.vs().scrollR) * fr.z || 20, cw = g.cols.size(G.vs().scrollC) * fr.z || 64;
      const dr = Math.trunc(acc.y / rh), dc = Math.trunc(acc.x / cw);
      if (dr) { acc.y -= dr * rh; G.scrollBy(dr, 0); }
      if (dc) { acc.x -= dc * cw; G.scrollBy(0, dc); }
    };
    const up = () => {
      clearTimeout(lp);
      wrap.removeEventListener('pointermove', mv); wrap.removeEventListener('pointerup', up); wrap.removeEventListener('pointercancel', up);
      if (moved) return;
      const ed = L.editor;
      if (ed && ed.active) { if (ed.canPoint() && z0.kind === 'cell') { ed.pointStart({ r1: z0.r, c1: z0.c, r2: z0.r, c2: z0.c }); ed.pointEnd(); return; } if (!ed.commit()) return; }
      if (z0.kind === 'cell') {
        const s = G.sel();
        if (s.ranges.length === 1 && s.r === z0.r && s.c === z0.c && s.ranges[0].r1 === s.ranges[0].r2) { L.editor.begin('edit'); return; }
        G.select(z0.r, z0.c, { noScroll: true });
      } else if (z0.kind === 'colhdr') G.selectRange({ r1: 0, c1: z0.c, r2: MAXR - 1, c2: z0.c }, { r: G.vs().scrollR, c: z0.c }, { noScroll: true });
      else if (z0.kind === 'rowhdr') G.selectRange({ r1: z0.r, c1: 0, r2: z0.r, c2: MAXC - 1 }, { r: z0.r, c: G.vs().scrollC }, { noScroll: true });
      else if (z0.kind === 'corner') G.selectAll();
    };
    try { wrap.setPointerCapture(e.pointerId); } catch (err) { /* ignore */ }
    wrap.addEventListener('pointermove', mv); wrap.addEventListener('pointerup', up); wrap.addEventListener('pointercancel', up);
  }
  function onContext(e) {
    if (e.target.closest('.gobj')) return;
    e.preventDefault();
    const z = zone(pt(e));
    if (L.editor.active) return;
    L.app.contextMenu(e, z);
  }

  /* ------------------------------------------------------------ keyboard */
  let endMode = false;
  G.endMode = () => endMode;
  function onKey(e) {
    if (L.ui.menuKey(e) || L.ui.dialogKey(e)) { e.preventDefault(); return; }
    const ed = L.editor;
    if (ed.active) { ed.key(e); return; }
    if (e.isComposing) return;
    if (G.selectedObject() && L.drawing && L.drawing.key && L.drawing.key(e, G.selectedObject())) { e.preventDefault(); return; }
    const k = e.key, ctrl = e.ctrlKey || e.metaKey, shift = e.shiftKey, alt = e.altKey;
    const s = G.sel();
    /* app-level shortcuts first */
    if (L.app.shortcut(e)) { e.preventDefault(); return; }
    const sh = G.sheet();
    const nav = (dr, dc) => {
      e.preventDefault();
      const useCtrl = ctrl || endMode;
      if (endMode) { endMode = false; L.bus.emit('mode'); }
      if (G.scrollLock) { G.scrollBy(dr, dc); return; }
      G.move(dr, dc, { ctrl: useCtrl, extend: shift });
    };
    switch (k) {
      case 'ArrowUp': nav(-1, 0); return;
      case 'ArrowDown': if (alt) { e.preventDefault(); L.app.pickList(); return; } nav(1, 0); return;
      case 'ArrowLeft': nav(0, -1); return;
      case 'ArrowRight': nav(0, 1); return;
      case 'Enter': e.preventDefault(); if (s.ranges.length === 1 && s.ranges[0].r1 === s.ranges[0].r2 && s.ranges[0].c1 === s.ranges[0].c2 && !L.app.opts.moveAfterEnter) return; { const d = L.app.opts.enterDir || 'down'; const m = { down: [1, 0], up: [-1, 0], right: [0, 1], left: [0, -1] }[d]; G.moveInSelection(shift ? -m[0] : m[0], shift ? -m[1] : m[1]); } return;
      case 'Tab': e.preventDefault(); G.moveInSelection(0, shift ? -1 : 1); return;
      case 'Home': {
        e.preventDefault();
        if (endMode) { endMode = false; L.bus.emit('mode'); const r = sh.maxR, c = sh.maxC; shift ? G.extendTo(r, c) : G.select(r, c); return; }
        const fr = G.frame();
        if (ctrl) { const r = fr.fr ? fr.sR : 0, c = fr.fc ? fr.sC : 0; shift ? G.extendTo(r, c) : G.select(r, c); }
        else { const c = fr.fc ? fr.sC : 0; shift ? G.extendTo(s.extent && shift ? s.extent.r : s.r, c) : G.select(s.r, c); }
        return;
      }
      case 'End':
        e.preventDefault();
        if (ctrl) { const r = Math.max(0, sh.maxR), c = Math.max(0, sh.maxC); shift ? G.extendTo(r, c) : G.select(r, c); return; }
        endMode = !endMode; L.bus.emit('mode');
        return;
      case 'PageDown': case 'PageUp': {
        e.preventDefault();
        if (ctrl) { L.app.nextSheet(k === 'PageDown' ? 1 : -1); return; }
        const fr = G.frame();
        const main = fr.panes[fr.panes.length - 1];
        if (alt) { const n = Math.max(1, main.cEnd - main.c0); G.scrollBy(0, k === 'PageDown' ? n : -n); const c = Math.max(0, Math.min(MAXC - 1, s.c + (k === 'PageDown' ? n : -n))); shift ? G.extendTo(s.r, c, { noScroll: true }) : G.select(s.r, c, { noScroll: true }); return; }
        const n = Math.max(1, main.rEnd - main.r0);
        G.scrollBy(k === 'PageDown' ? n : -n, 0);
        const r = Math.max(0, Math.min(MAXR - 1, (shift && s.extent ? s.extent.r : s.r) + (k === 'PageDown' ? n : -n)));
        shift ? G.extendTo(r, s.c, { noScroll: true }) : G.select(r, s.c, { noScroll: true });
        return;
      }
      case 'F2': e.preventDefault(); if (shift) { L.app.editComment(); return; } if (!G.guardProtect(null, true)) ed.begin('edit'); return;
      case 'Delete': e.preventDefault(); L.ui.exec('clearContents'); return;
      case 'Backspace': e.preventDefault(); if (!G.guardProtect(null, true)) ed.begin('enter', ''); return;
      case 'Escape': e.preventDefault(); if (G.state.marquee) { G.state.marquee = null; L.clip.clearMarquee(); G.paint(); } if (G.painting) L.app.finishPaint(true); return;
      case ' ':
        if (ctrl && shift) { e.preventDefault(); G.selectAll(); return; }
        if (ctrl) { e.preventDefault(); const rg = s.ranges[s.active]; s.ranges[s.active] = { r1: 0, c1: rg.c1, r2: MAXR - 1, c2: rg.c2 }; G.changed(); return; }
        if (shift) { e.preventDefault(); const rg = s.ranges[s.active]; s.ranges[s.active] = { r1: rg.r1, c1: 0, r2: rg.r2, c2: MAXC - 1 }; G.changed(); return; }
        break;
      default:
    }
    if (ctrl && !alt) return; /* other Ctrl shortcuts are handled by the app or the browser (copy / paste) */
    if (k === 'F4' || k === 'F5' || k === 'F9' || k === 'F11' || k === 'F7' || k === 'F1' || k === 'F12') return;
    /* typing starts editing: the character lands in the input; onInput opens the editor */
  }
  function onInput() {
    const ed = L.editor;
    if (ed.active) { ed.onInput(); return; }
    const v = inp.value;
    if (!v) return;
    if (G.guardProtect(null, true)) { inp.value = ''; return; }
    ed.begin('enter', v);
  }
  G.scrollLock = false;

  /* ------------------------------------------------------------ protection */
  G.protectedSheet = () => { const sh = G.sheet(); return !!sh.protection; };
  /** true (and a message) when the sheet is protected against an action; cell edits check locked cells */
  G.guardProtect = function (allowKey, cellEdit) {
    const sh = G.sheet();
    if (!sh.protection) return false;
    const at = sh.protection.attrs || {};
    if (allowKey && at[allowKey] === '0') return false;
    if (cellEdit) {
      const s = G.sel();
      let locked = false;
      for (const rg of s.ranges) {
        const rr = O.clip(sh, rg);
        for (let r = rr.r1; r <= Math.min(rr.r2, rr.r1 + 500); r++) for (let c = rr.c1; c <= Math.min(rr.c2, rr.c1 + 100); c++) { const st = LY.styleOf(sh, r, c, sh.get(r, c)); if (!(st.prot && st.prot.locked === false)) { locked = true; break; } }
      }
      if (!locked) return false;
    }
    L.ui.msg('The cell or chart you are trying to change is protected and therefore read-only.\n\nTo modify a protected cell or chart, first remove protection using the Unprotect Sheet command (Tools menu, Protection submenu). You may be prompted for a password.', { icon: 'warn' });
    return true;
  };

  /* ------------------------------------------------------------ objects layer */
  let selObj = null;
  G.selectedObject = () => selObj;
  G.deselectObject = () => { if (selObj) { selObj = null; G.syncObjects(); L.bus.emit('selection'); } };
  G.selectObject = (d) => { selObj = d; G.syncObjects(); L.bus.emit('selection'); };
  /** pixel rect (canvas coords) of a drawing's anchor */
  G.anchorRect = function (sh, a) {
    const fr = G.frame();
    const g = LY.geo(sh), z = fr.z;
    const main = fr.panes[fr.panes.length - 1];
    if (!main) return null;
    const pane = (r, c) => RD.paneOf(fr, r, c) || main;
    const px = (p) => { const pn = pane(p.r, p.c); return { x: RD.colX(fr, pn, p.c) + Math.min(p.cOff || 0, 1e6) * (96 / 72) * z, y: RD.rowY(fr, pn, p.r) + (p.rOff || 0) * (96 / 72) * z }; };
    if (a.type === 'abs') { const x = main.x + (a.x * 96 / 72 - g.cols.pos(main.c0)) * z, y = main.y + (a.y * 96 / 72 - g.rows.pos(main.r0)) * z; return { x, y, w: a.w * 96 / 72 * z, h: a.h * 96 / 72 * z }; }
    const p1 = px(a.from);
    if (a.type === 'one' || !a.to) return { x: p1.x, y: p1.y, w: (a.w || 0) * 96 / 72 * z, h: (a.h || 0) * 96 / 72 * z };
    const p2 = px(a.to);
    return { x: p1.x, y: p1.y, w: Math.max(0, p2.x - p1.x), h: Math.max(0, p2.y - p1.y) };
  };
  /** inverse: a canvas rect → two-cell anchor */
  G.rectToAnchor = function (sh, x, y, w, h, keep) {
    const fr = G.frame(), g = LY.geo(sh), z = fr.z;
    const main = fr.panes[fr.panes.length - 1];
    const toCell = (px, py) => {
      const ux = g.cols.pos(main.c0) + (px - main.x) / z, uy = g.rows.pos(main.r0) + (py - main.y) / z;
      const c = g.cols.at(Math.max(0, ux)), r = g.rows.at(Math.max(0, uy));
      return { c, cOff: Math.max(0, (ux - g.cols.pos(c)) * 72 / 96), r, rOff: Math.max(0, (uy - g.rows.pos(r)) * 72 / 96) };
    };
    const a = { type: 'two', from: toCell(x, y), to: toCell(x + w, y + h), editAs: keep && keep.editAs || 'twoCell' };
    return a;
  };
  let objEls = new Map();
  G.syncObjects = function (rebuild) {
    if (!layer || !G.wb) return;
    const sh = G.sheet();
    const fr = lastFrame;
    if (!fr) return;
    if (rebuild || layer._sheet !== sh) { layer.textContent = ''; objEls = new Map(); layer._sheet = sh; }
    const seen = new Set();
    const clip = { x: fr.hdrW, y: fr.hdrH, w: G.view().w - fr.hdrW, h: G.view().h - fr.hdrH };
    for (const d of sh.drawings) {
      if (d.hidden) continue;
      const rc = G.anchorRect(sh, d.anchor);
      if (!rc) continue;
      let el = objEls.get(d);
      if (!el) { el = L.drawing ? L.drawing.element(d, sh) : h('div', { class: 'gobj' }); objEls.set(d, el); layer.appendChild(el); }
      seen.add(d);
      const visible = rc.x + rc.w > clip.x && rc.y + rc.h > clip.y && rc.x < clip.x + clip.w && rc.y < clip.y + clip.h;
      el.hidden = !visible;
      if (!visible) continue;
      Object.assign(el.style, { left: rc.x + 'px', top: rc.y + 'px', width: Math.max(1, rc.w) + 'px', height: Math.max(1, rc.h) + 'px' });
      el.classList.toggle('sel', selObj === d);
      if (L.drawing && L.drawing.update) L.drawing.update(el, d, sh, rc, selObj === d);
    }
    for (const [d, el] of objEls) if (!seen.has(d) || !sh.drawings.includes(d)) { el.remove(); objEls.delete(d); }
    /* clip the layer to the cell area so objects slide under the headers */
    layer.style.clipPath = `inset(${fr.hdrH}px 0 0 ${fr.hdrW}px)`;
    syncWidgets(sh, fr);
  };
  /* AutoFilter buttons, visible comments, validation drop-down, outline buttons */
  let widgets = [];
  function syncWidgets(sh, fr) {
    for (const w of widgets) w.remove();
    widgets = [];
    const add = (el) => { layer.appendChild(el); widgets.push(el); return el; };
    const af = sh.autoFilter && sh.autoFilter.ref;
    const filters = [];
    if (af) filters.push({ ref: af, f: sh.autoFilter, table: null });
    for (const t of G.wb.tables) if (t.sheet === sh && t.autoFilter && t.header) filters.push({ ref: t.ref, f: t.filter || { cols: [] }, table: t });
    for (const x of filters) {
      const r = x.ref.r1;
      for (let c = x.ref.c1; c <= x.ref.c2; c++) {
        const fc = x.f.cols.find((q) => q.col === c - x.ref.c1);
        if (fc && fc.hiddenButton) continue;
        const m = LY.mergeAt(sh, r, c);
        if (m && m.c2 !== c) continue;
        const q = G.cellRect(r, c);
        if (!q.w || !q.h || q.x + q.w < fr.hdrW || q.y + q.h < fr.hdrH) continue;
        const s = Math.min(16, q.h - 1);
        const on = !!(fc && (fc.values || fc.custom || fc.top10 || fc.dynamic || fc.color || fc.blank || fc.dates));
        const b = add(h('button', { class: 'af-btn' + (on ? ' on' : ''), type: 'button', tabindex: '-1', 'aria-label': 'AutoFilter', html: '<span class="dd-arrow"></span>' }));
        Object.assign(b.style, { left: (q.x + q.w - s - 1) + 'px', top: (q.y + q.h - s - 1) + 'px', width: s + 'px', height: s + 'px' });
        b.addEventListener('pointerdown', (e) => { e.preventDefault(); e.stopPropagation(); });
        b.addEventListener('click', (e) => { e.stopPropagation(); L.filter.open(b, x, c); });
      }
    }
    /* comments shown permanently (or by View ▸ Comments) */
    const showAll = G.wb._showComments;
    for (const cm of sh.comments.values()) {
      if (!(cm.visible || showAll || (hoverCm && hoverCm === cm))) continue;
      const q = G.cellRect(cm.r, cm.c);
      if (q.x + q.w < fr.hdrW - 200 || q.y < fr.hdrH - 200) continue;
      const box = add(h('div', { class: 'cmt-box' }));
      const w = (cm.w || 108) * 96 / 72 * fr.z, hh = (cm.h || 59) * 96 / 72 * fr.z;
      const bx = q.x + q.w + 15, by = Math.max(fr.hdrH + 2, q.y - 8);
      Object.assign(box.style, { left: bx + 'px', top: by + 'px', width: w + 'px', height: hh + 'px', fontSize: (8 * fr.z) + 'pt', background: cm.fill && /^#[0-9a-f]{6}$/i.test(cm.fill) ? cm.fill : '' });
      if (cm.author && !/^\s*$/.test(cm.author) && !(cm.text || '').startsWith(cm.author)) box.appendChild(h('b', { text: cm.author + ':' })), box.appendChild(document.createTextNode('\n'));
      if (cm.runs) for (const run of cm.runs) { const sp = h('span', { text: run.t }); if (run.font && run.font.b) sp.style.fontWeight = 'bold'; if (run.font && run.font.i) sp.style.fontStyle = 'italic'; box.appendChild(sp); }
      else box.appendChild(document.createTextNode(cm.text || ''));
      const line = add(h('div', { class: 'cmt-line' }));
      const x1 = q.x + q.w - 1, y1 = q.y + 2, x2 = bx, y2 = by + 6;
      const len = Math.hypot(x2 - x1, y2 - y1), ang = Math.atan2(y2 - y1, x2 - x1);
      Object.assign(line.style, { left: x1 + 'px', top: y1 + 'px', width: len + 'px', transform: `rotate(${ang}rad)` });
      box.addEventListener('pointerdown', (e) => { e.stopPropagation(); });
      box.addEventListener('dblclick', (e) => { e.stopPropagation(); L.app.editComment(cm.r, cm.c); });
    }
    /* data validation drop-down arrow beside the active cell */
    const s = G.sel();
    const dv = L.cf.dvAt(sh, s.r, s.c);
    if (dv && dv.type === 'list' && dv.showDrop !== false && !(L.editor && L.editor.active)) {
      const q = G.cellRect(s.r, s.c);
      if (q.w && q.h) {
        const sz = Math.min(16, q.h);
        const b = add(h('button', { class: 'dv-btn', type: 'button', tabindex: '-1', 'aria-label': 'Choose from list', html: '<span class="dd-arrow"></span>' }));
        Object.assign(b.style, { left: (q.x + q.w + 1) + 'px', top: (q.y + q.h - sz) + 'px', height: sz + 'px' });
        b.addEventListener('pointerdown', (e) => { e.preventDefault(); e.stopPropagation(); });
        b.addEventListener('click', (e) => { e.stopPropagation(); L.app.pickList(); });
      }
    }
    if (dv && dv.showInput && (dv.prompt || dv.promptTitle) && !(L.editor && L.editor.active)) {
      const q = G.cellRect(s.r, s.c);
      const tip = add(h('div', { class: 'dv-tip' }));
      if (dv.promptTitle) tip.appendChild(h('b', { text: dv.promptTitle }));
      tip.appendChild(document.createTextNode(dv.prompt || ''));
      Object.assign(tip.style, { left: (q.x + Math.min(q.w, 40)) + 'px', top: (q.y + q.h + 4) + 'px' });
    }
  }
  let hoverCm = null, hoverTimer = null;
  function showCommentHover(r, c) {
    const sh = G.sheet();
    const cm = r >= 0 ? sh.comments.get(M.key(r, c)) || (() => { const m = LY.mergeAt(sh, r, c); return m ? sh.comments.get(M.key(m.r1, m.c1)) : null; })() : null;
    if (cm === hoverCm) return;
    clearTimeout(hoverTimer);
    if (!cm) { if (hoverCm) { hoverCm = null; G.syncObjects(); } return; }
    hoverTimer = setTimeout(() => { hoverCm = cm; G.syncObjects(); }, 250);
  }
  wrapLeave();
  function wrapLeave() { document.addEventListener('pointerleave', () => { if (hoverCm) { hoverCm = null; G.syncObjects(); } }); }
})(typeof window !== 'undefined' ? window : globalThis);
