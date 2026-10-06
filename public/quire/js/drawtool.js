/* Quire — drawing new AutoShapes, lines and text boxes onto the page (Drawing toolbar). */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout;
  const DT = (L.drawtool = { active: null });
  const doc = () => D.doc;
  const A = () => L.app;

  DT.start = function (opts) {
    if (A().view !== 'print') A().setView('print');
    DT.active = opts;
    const sc = L.$('#scroller');
    sc.classList.add('drawing-mode');
    A().status(opts.textbox ? 'Drag to draw a text box. Press Esc to cancel.' : 'Drag to draw. Hold Shift for squares, circles and 15° lines. Press Esc to cancel.', 8000);
  };
  DT.stop = function () { DT.active = null; L.$('#scroller').classList.remove('drawing-mode'); if (L.app && L.app.status) L.app.status(''); };
  function pageAt(x, y) {
    for (const pg of LY.pages) { const r = pg.el.getBoundingClientRect(); if (x >= r.left && x <= r.right && y >= r.top && y <= r.bottom) return { pg, r }; }
    return null;
  }
  /** first paragraph fragment on the page at or above y (client px) */
  function anchorOn(pg, y) {
    const frags = Array.from(pg.el.querySelectorAll('.pg-body .p')).filter((f) => !f.closest('table') && !f.closest('.flt') && !f.closest('.tbx'));
    let best = frags[0] || null;
    for (const f of frags) { if (f.getBoundingClientRect().top <= y) best = f; else break; }
    if (!best) { const any = pg.el.querySelector('.pg-body .p'); best = any; }
    if (!best) return null;
    const p = D.byId(doc(), +best.dataset.pid);
    return p ? { p, o: +best.dataset.from || 0, frag: best } : null;
  }
  function onDown(e) {
    const opts = DT.active;
    if (!opts || e.button !== 0) return;
    const hit = pageAt(e.clientX, e.clientY);
    if (!hit) return;
    e.preventDefault(); e.stopPropagation();
    const z = LY.zoom, k = 1 / (z * D.PX);
    const sc = LY.scroller;
    const sr = sc.getBoundingClientRect();
    const ghost = L.h('div', { class: 'draw-ghost' + (opts.line ? ' line' : '') });
    sc.appendChild(ghost);
    const x0 = e.clientX, y0 = e.clientY;
    let x1 = x0, y1 = y0;
    const upd = (ev) => {
      x1 = ev.clientX; y1 = ev.clientY;
      if (ev.shiftKey) {
        const dx = x1 - x0, dy = y1 - y0;
        if (opts.line) { const ang = Math.round(Math.atan2(dy, dx) / (Math.PI / 12)) * (Math.PI / 12); const len = Math.hypot(dx, dy); x1 = x0 + Math.cos(ang) * len; y1 = y0 + Math.sin(ang) * len; }
        else { const m = Math.max(Math.abs(dx), Math.abs(dy)); x1 = x0 + Math.sign(dx || 1) * m; y1 = y0 + Math.sign(dy || 1) * m; }
      }
      if (opts.line) {
        const len = Math.hypot(x1 - x0, y1 - y0), ang = Math.atan2(y1 - y0, x1 - x0);
        ghost.style.cssText = `left:${x0 - sr.left + sc.scrollLeft}px;top:${y0 - sr.top + sc.scrollTop}px;width:${len}px;height:0;border:0;border-top:1px solid #000;transform-origin:0 0;transform:rotate(${ang}rad)`;
      } else ghost.style.cssText = `left:${Math.min(x0, x1) - sr.left + sc.scrollLeft}px;top:${Math.min(y0, y1) - sr.top + sc.scrollTop}px;width:${Math.abs(x1 - x0)}px;height:${Math.abs(y1 - y0)}px;${opts.geom === 'ellipse' ? 'border-radius:50%' : ''}`;
    };
    upd(e);
    const mv = (ev) => upd(ev);
    const up = (ev) => {
      window.removeEventListener('pointermove', mv, true);
      window.removeEventListener('pointerup', up, true);
      ghost.remove();
      upd(ev);
      let w = Math.abs(x1 - x0) * k, hh = Math.abs(y1 - y0) * k;
      const small = w < 3 && hh < 3;
      if (small) { if (opts.line) { w = 72; hh = 0; x1 = x0 + 72 / k; y1 = y0; } else { w = opts.textbox ? 144 : 72; hh = opts.textbox ? 72 : 72; } }
      const left = (Math.min(x0, small && !opts.line ? x0 : x1) - hit.r.left) * k, top = (Math.min(y0, small && !opts.line ? y0 : y1) - hit.r.top) * k;
      const anc = anchorOn(hit.pg, y0);
      DT.stop();
      if (!anc) return;
      const it = L.drawing.newShape(opts.geom || 'rect', L.round(Math.max(opts.line ? 0 : 4, w), 2), L.round(Math.max(opts.line ? 0 : 4, hh), 2), { arrow: opts.arrow, dblArrow: opts.dblArrow, textbox: opts.textbox });
      if (opts.line) { if (x1 < x0) it.flipH = true; if (y1 < y0) it.flipV = true; }
      if (A().shapeDefaults && !opts.line && !opts.textbox) Object.assign(it, L.clone(A().shapeDefaults));
      if (opts.textbox) { it.geom = 'rect'; it.line = { c: '#000000', w: 0.75, dash: 'solid' }; it.fill = { t: 'solid', c: '#FFFFFF', a: 1 }; }
      L.drawing.setWrap(it, 'none');
      it.float.posH = { rel: 'page', off: L.round(left, 2) };
      it.float.posV = { rel: 'page', off: L.round(top, 2) };
      it.float.z = A().maxZ() + 1024;
      it.name = (opts.textbox ? 'Text Box ' : 'AutoShape ') + D.nid();
      let tbPara = null;
      E.edit(opts.textbox ? 'Insert Text Box' : 'Insert AutoShape', () => {
        O.insertItem(D.pos(anc.p, anc.o), it);
        if (it.tb) tbPara = it.tb.blocks[0];
        return D.pos(anc.p, anc.o);
      });
      setTimeout(() => {
        if (tbPara) { D.reindex(doc()); E.setSel(D.pos(tbPara, 0)); E.focus(); return; }
        const wrap = Array.from(LY.root.querySelectorAll('.flt')).find((x) => x._item === it || (+x.dataset.pid === anc.p.id && +x.dataset.o === anc.o));
        if (wrap) L.drawing.select(wrap, { p: anc.p, o: anc.o });
      }, 30);
    };
    window.addEventListener('pointermove', mv, true);
    window.addEventListener('pointerup', up, true);
  }
  L.bus.on('app-ready', () => {
    LY.scroller.addEventListener('pointerdown', onDown, true);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && DT.active) { DT.stop(); A().status(''); } });
  });
})();
