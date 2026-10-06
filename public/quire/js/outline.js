/* Quire — Outline view: heading levels, promote/demote, move up/down, expand/collapse, show level. */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout;
  const OL = (L.outline = { showLevel: 0, firstLine: false, plain: false });
  const doc = () => D.doc;
  const collapsed = new Set();

  /** outline level of a top-level block: 1..9 for headings, 0 for body text */
  const levelOf = (b) => (b.t === 'p' ? D.headingLevel(doc(), b) : 0);
  OL.levelOf = levelOf;
  OL.render = function (box, ctx) {
    const d = doc();
    const blocks = d.main.blocks;
    box.classList.toggle('firstline', !!OL.firstLine);
    box.classList.toggle('plain', !!OL.plain);
    const frame = { el: box, limit: Infinity, top: box.getBoundingClientRect().top, empty: true, kind: 'flow', width: parseFloat(box.style.width) * D.PX / D.PX };
    LY.flow(frame, blocks, { bi: 0, sub: null }, blocks.length, Object.assign({}, ctx, { view: 'outline' }));
    /* decorate */
    let curHead = 0;
    let hideUnder = 0; /* level of a collapsed heading whose body is hidden */
    const byId = new Map();
    for (const f of box.querySelectorAll(':scope > .p, :scope > table.tbl')) {
      if (f.classList.contains('p')) byId.set(+f.dataset.pid, f);
    }
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      const lv = levelOf(b);
      const el = b.t === 'p' ? byId.get(b.id) : box.querySelector(`:scope > table.tbl[data-tid="${b.id}"]`);
      if (lv) {
        if (hideUnder && lv <= hideUnder) hideUnder = 0;
        curHead = lv;
      }
      const effective = lv || 10;
      const hidden = (hideUnder && (!lv || lv > hideUnder)) || (OL.showLevel && effective > OL.showLevel);
      if (el) {
        el.style.setProperty('--olvl', lv ? lv - 1 : curHead);
        el.classList.add(lv ? 'olh' : 'olb');
        if (hidden) el.classList.add('olcoll');
        if (lv && collapsed.has(b.id)) el.classList.add('olhascoll');
        /* headings without subordinate text show a minus sign */
        if (lv) { const nxt = blocks[i + 1]; const nl = nxt ? levelOf(nxt) : -1; if (!nxt || (nl && nl <= lv)) el.classList.add('olempty'); }
        if (b.t === 'tbl') el.style.marginLeft = (curHead * 36 + 18) * D.PX + 'px';
      }
      if (lv && collapsed.has(b.id) && !hideUnder) hideUnder = lv;
    }
  };
  /* ---------- level changes ---------- */
  function selectedTop() {
    const d = doc();
    if (!E.sel) return [];
    const [a, b] = E.range();
    const ta = D.topBlock(d, a.p), tb = D.topBlock(d, b.p);
    if (!ta || !tb || ta.cont !== d.main || tb.cont !== d.main) return [];
    return d.main.blocks.slice(ta.i, tb.i + 1).filter((x) => x.t === 'p');
  }
  const ensureHeading = (n) => { const d = doc(); const id = 'Heading' + n; if (!d.styles[id]) { D.touchKey(d, 'styles'); d.styles[id] = D.builtinStyles()[id]; D.stylesChanged(); } return id; };
  OL.setLevel = function (n) {
    const paras = selectedTop();
    if (!paras.length) return;
    E.edit(n ? 'Promote' : 'Demote to Body Text', () => {
      if (n) ensureHeading(n);
      O.setParaProps(paras, (pPr) => { if (n) pPr.style = 'Heading' + n; else { delete pPr.style; delete pPr.outline; } });
      return E.sel;
    });
  };
  OL.shift = function (delta) {
    const d = doc();
    const paras = selectedTop();
    if (!paras.length) return;
    /* list items in other views: Alt+Shift+arrows change list level */
    if (L.app.view !== 'outline' && paras.every((p) => !levelOf(p)) && L.lists && L.lists.info(paras[0])) { L.lists.shift(delta); return; }
    E.edit(delta < 0 ? 'Promote' : 'Demote', () => {
      for (const p of paras) {
        let lv = levelOf(p);
        if (!lv) {
          if (delta > 0) continue;
          /* body text promoted: becomes a heading one level below the heading above */
          let q = D.prevPara(d, p), above = 0;
          while (q) { const ql = levelOf(q); if (ql) { above = ql; break; } q = D.prevPara(d, q); }
          lv = Math.min(9, above || 1);
        } else lv = L.clamp(lv + delta, 1, 9);
        ensureHeading(lv);
        D.touch(p);
        p.pPr.style = 'Heading' + lv;
        delete p.pPr.outline;
      }
      return E.sel;
    });
  };
  /** move the selected paragraphs (with collapsed subtext) past the neighbouring block */
  OL.move = function (dir) {
    const d = doc();
    if (!E.sel) return;
    const [a, b] = E.range();
    const ta = D.topBlock(d, a.p), tb = D.topBlock(d, b.p);
    if (!ta || !tb || ta.cont !== tb.cont) return;
    const cont = ta.cont;
    let i0 = ta.i, i1 = tb.i;
    const bl = cont.blocks;
    /* a collapsed heading carries its subordinate content */
    const extent = (i) => { const lv = levelOf(bl[i]); if (!lv || !collapsed.has(bl[i].id) || L.app.view !== 'outline') return i; let k = i + 1; while (k < bl.length && (!levelOf(bl[k]) || levelOf(bl[k]) > lv)) k++; return k - 1; };
    i1 = Math.max(i1, extent(i1));
    if (dir < 0 && i0 === 0) return;
    if (dir > 0 && i1 >= bl.length - 1) return;
    /* never move past the final paragraph carrying the last section */
    E.edit(dir < 0 ? 'Move Up' : 'Move Down', () => {
      D.touchList(cont);
      const chunk = bl.slice(i0, i1 + 1);
      if (dir < 0) {
        /* the block above, or the collapsed group it belongs to */
        let j = i0 - 1;
        if (L.app.view === 'outline') { const hidden = (k) => { for (let m = k - 1; m >= 0; m--) { const lv = levelOf(bl[m]); if (lv && collapsed.has(bl[m].id)) { const own = levelOf(bl[k]); return !own || own > lv ? m : -1; } if (lv && lv <= (levelOf(bl[k]) || 10)) break; } return -1; }; const hh = hidden(j); if (hh >= 0) j = hh; }
        bl.splice(i0, chunk.length);
        bl.splice(j, 0, ...chunk);
      } else {
        const j = extent(i1 + 1);
        const after = bl.slice(i1 + 1, j + 1);
        bl.splice(i0, chunk.length + after.length, ...after, ...chunk);
      }
      d._idxDirty = true;
      return E.sel;
    });
  };
  OL.expand = function (on) {
    const d = doc();
    if (!E.sel) return;
    const paras = selectedTop();
    const heads = paras.filter((p) => levelOf(p));
    if (!heads.length) { let q = E.sel.f.p; while (q && !levelOf(q)) q = D.prevPara(d, q); if (q) heads.push(q); }
    for (const p of heads) { if (on) collapsed.delete(p.id); else collapsed.add(p.id); }
    if (on && !heads.length) collapsed.clear();
    LY.render();
    E.restoreDom(true);
  };
  OL.toggleHeading = function (p) { if (collapsed.has(p.id)) collapsed.delete(p.id); else collapsed.add(p.id); LY.render(); E.restoreDom(true); };
  OL.show = function (n) { OL.showLevel = n || 0; collapsed.clear(); if (L.app.view === 'outline') { LY.render(); E.restoreDom(true); } L.ui.refresh(); };
  OL.toggle = function (k) { OL[k] = !OL[k]; if (L.app.view === 'outline') { LY.render(); E.restoreDom(true); } L.ui.refresh(); };
  /* clicking an outline symbol selects the heading with its subtext; double-click expands/collapses */
  L.bus.on('app-ready', () => {
    LY.root.addEventListener('mousedown', (e) => {
      if (L.app.view !== 'outline') return;
      const f = e.target.closest && e.target.closest('.flow.outline > .p');
      if (!f) return;
      const r = f.getBoundingClientRect();
      const lvl = parseFloat(getComputedStyle(f).getPropertyValue('--olvl')) || 0;
      const symX = r.left + lvl * 36 * LY.zoom;
      if (e.clientX < symX - 2 || e.clientX > symX + 14 * LY.zoom) return;
      e.preventDefault();
      const p = D.byId(doc(), +f.dataset.pid);
      if (!p) return;
      if (e.detail >= 2 && levelOf(p)) { OL.toggleHeading(p); return; }
      /* select heading and its subordinate text */
      const d = doc();
      const bl = d.main.blocks;
      const i = bl.indexOf(p);
      const lv = levelOf(p) || 10;
      let k = i + 1;
      while (k < bl.length && (bl[k].t !== 'p' || !levelOf(bl[k]) || levelOf(bl[k]) > lv) && levelOf(p)) k++;
      const last = bl[Math.max(i, k - 1)];
      const lp = last.t === 'p' ? last : D.lastPara([last]);
      LY.root.focus({ preventScroll: true });
      E.setSel(D.pos(p, 0), D.pos(lp, D.plen(lp)), { noScroll: true });
    });
  });
})();
