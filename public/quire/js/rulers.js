/* Quire — horizontal and vertical rulers: margins, indents, tab stops, column and table markers,
 * the tab selector, and dragging all of them.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, E = L.ed, LY = L.layout, R = L.R;
  const RU = (L.rulers = {});
  const doc = () => D.doc;
  const { h } = L;
  let H = null, V = null, corner = null;
  let tabKind = 0;
  const TAB_KINDS = [['left', 'Left Tab'], ['center', 'Center Tab'], ['right', 'Right Tab'], ['decimal', 'Decimal Tab'], ['bar', 'Bar Tab'], ['firstLine', 'First Line Indent'], ['hanging', 'Hanging Indent']];
  const SVG = {
    left: '<svg width="9" height="9"><path d="M2 0v7h6" stroke="#000" stroke-width="1.6" fill="none"/></svg>',
    center: '<svg width="9" height="9"><path d="M4.5 0v7M0 7h9" stroke="#000" stroke-width="1.6" fill="none"/></svg>',
    right: '<svg width="9" height="9"><path d="M7 0v7H1" stroke="#000" stroke-width="1.6" fill="none"/></svg>',
    decimal: '<svg width="9" height="9"><path d="M4.5 0v7M0 7h9" stroke="#000" stroke-width="1.6" fill="none"/><circle cx="7.3" cy="4" r="1" fill="#000"/></svg>',
    bar: '<svg width="9" height="9"><path d="M4.5 0v9" stroke="#000" stroke-width="1.6"/></svg>',
    firstLine: '<svg width="9" height="9"><path d="M0.5 0.5h8l-4 5z" fill="#fff" stroke="#000"/></svg>',
    hanging: '<svg width="9" height="9"><path d="M4.5 1l4 5h-8z" fill="#fff" stroke="#000"/></svg>',
  };
  const MARK = {
    first: '<svg width="9" height="7"><path d="M0.5 0.5h8v2l-4 4-4-4z" style="fill:var(--indent-fill);stroke:var(--indent-line)"/></svg>',
    hang: '<svg width="9" height="7"><path d="M4.5 0.5l4 4v2h-8v-2z" style="fill:var(--indent-fill);stroke:var(--indent-line)"/></svg>',
    left: '<svg width="9" height="5"><rect x="0.5" y="0.5" width="8" height="4" style="fill:var(--indent-fill);stroke:var(--indent-line)"/></svg>',
    right: '<svg width="9" height="7"><path d="M4.5 0.5l4 4v2h-8v-2z" style="fill:var(--indent-fill);stroke:var(--indent-line)"/></svg>',
  };
  RU.mount = function (hEl, vEl, cEl) {
    H = hEl; V = vEl; corner = cEl;
    H.appendChild(h('canvas', { class: 'r-canvas' }));
    V.appendChild(h('canvas', { class: 'r-canvas' }));
    const sel = h('button', { class: 'tabsel', type: 'button', 'aria-label': 'Tab type', 'data-tip': TAB_KINDS[0][1], html: SVG.left });
    sel.addEventListener('pointerdown', (e) => e.preventDefault());
    sel.addEventListener('click', () => { tabKind = (tabKind + 1) % TAB_KINDS.length; sel.innerHTML = SVG[TAB_KINDS[tabKind][0]]; sel.setAttribute('data-tip', TAB_KINDS[tabKind][1]); sel.setAttribute('aria-label', TAB_KINDS[tabKind][1]); });
    corner.appendChild(sel);
    H.addEventListener('pointerdown', onRulerDown);
    H.addEventListener('dblclick', (e) => { if (e.target.closest('.rmark.tabstop')) { L.dlg && L.dlg.tabs(); return; } L.dlg && (state && e.clientX > state.x0 && e.clientX < state.x0 + state.w ? L.dlg.tabs() : L.dlg.pageSetup()); });
    V.addEventListener('pointerdown', onVDown);
    V.addEventListener('dblclick', () => L.dlg && L.dlg.pageSetup());
    L.onLook(() => RU.draw());   /* the canvas follows the look */
  };
  RU.drawSoon = L.rafThrottle(() => RU.draw());
  let state = null;
  /** geometry for the caret's paragraph: x0 (client px of the text column's left), w (px), z, page rect */
  function geometry() {
    if (!E.sel || !LY.root) return null;
    const z = LY.zoom;
    const p = E.sel.f.p;
    const frag = LY.fragFor(p, E.sel.f.o) || LY.frags(p.id)[0];
    let colEl = frag ? frag.closest('.cf, .pg-col, .pg-hdr, .pg-ftr, .fnote, .flow, .bb, .tbi') : null;
    const pgEl = frag ? frag.closest('.pg') : LY.root.querySelector('.pg');
    if (!colEl) colEl = pgEl ? pgEl.querySelector('.pg-col') || pgEl.querySelector('.pg-body') : LY.root.querySelector('.flow');
    if (!colEl) return null;
    const cr = colEl.getBoundingClientRect();
    const pr = pgEl ? pgEl.getBoundingClientRect() : cr;
    const pg = pgEl ? LY.pages[+pgEl.dataset.pg] : null;
    const sect = pg ? pg.sect : D.sectFor(doc(), p);
    let x0 = cr.left, w = cr.width;
    if (colEl.classList.contains('flow')) { const cs = getComputedStyle(colEl); x0 += parseFloat(cs.paddingLeft) * z; w -= (parseFloat(cs.paddingLeft) + parseFloat(cs.paddingRight)) * z; }
    const ci = D.cellOf(doc(), p);
    return { z, x0, w, pgEl, pg, pr, sect, frag, colEl, ci, p };
  }
  RU.draw = function () {
    if (!H || H.hidden) return;
    const g = geometry();
    const hr = H.getBoundingClientRect();
    const cv = H.querySelector('canvas');
    const dpr = window.devicePixelRatio || 1;
    cv.width = Math.max(1, hr.width * dpr); cv.height = Math.max(1, hr.height * dpr);
    cv.style.width = hr.width + 'px'; cv.style.height = hr.height + 'px';
    const c = cv.getContext('2d');
    c.setTransform(dpr, 0, 0, dpr, 0, 0);
    c.clearRect(0, 0, hr.width, hr.height);
    for (const m of H.querySelectorAll('.rmark')) m.remove();
    state = null;
    if (!g) return;
    const z = g.z;
    const pxPt = D.PX * z;
    /* page extent and margins */
    let pageL, pageR, textL, textR;
    if (g.pgEl) {
      pageL = g.pr.left - hr.left; pageR = g.pr.right - hr.left;
      const pgB = g.pg;
      textL = pageL + pgB.bodyLeft * z; textR = textL + pgB.bodyW * z;
    } else { pageL = g.x0 - hr.left - 20; pageR = g.x0 - hr.left + g.w + 20; textL = g.x0 - hr.left; textR = textL + g.w; }
    c.fillStyle = L.lookColor('ruler-off', '#c9d6ee'); c.fillRect(Math.max(0, pageL), 4, Math.max(0, pageR - pageL), hr.height - 8);
    /* text area (or the cell / column) is white */
    const colL = g.x0 - hr.left, colR = colL + g.w;
    c.fillStyle = L.lookColor('ruler', '#fff');
    if (g.ci || (g.sect.cols && g.sect.cols.n > 1)) {
      /* white for each column or table cell boundary */
      if (g.ci) {
        const tblEl = g.frag.closest('table.tbl');
        const tr = g.frag.closest('tr');
        if (tr) for (const td of tr.children) { const r = td.getBoundingClientRect(); c.fillRect(r.left - hr.left + 3, 4, Math.max(0, r.width - 6), hr.height - 8); }
        void tblEl;
      } else {
        for (const col of g.pgEl ? g.pgEl.querySelectorAll('.pg-reg:last-child .pg-col, .pg-col') : []) { const r = col.getBoundingClientRect(); c.fillRect(r.left - hr.left, 4, r.width, hr.height - 8); }
      }
    } else c.fillRect(textL, 4, Math.max(0, textR - textL), hr.height - 8);
    /* ticks from the text start (Word measures from the left margin) */
    const origin = g.ci ? colL : textL;
    const unit = D.unit === 'cm' || D.unit === 'mm' ? 28.3465 : 72;
    const sub = unit === 72 ? 8 : 4;
    c.strokeStyle = L.lookColor('ruler-tick', '#444'); c.fillStyle = L.lookColor('ruler-ink', '#222');
    c.font = '9px Tahoma, Arial, sans-serif';
    c.textAlign = 'center';
    c.lineWidth = 1;
    const step = (unit / sub) * pxPt;
    if (step > 1.5) {
      const kmin = Math.ceil((Math.max(0, pageL) - origin) / step), kmax = Math.floor((Math.min(hr.width, pageR) - origin) / step);
      for (let k = kmin; k <= kmax; k++) {
        const x = Math.round(origin + k * step) + 0.5;
        if (k % sub === 0) { if (k !== 0) { const n = Math.abs(k / sub); if (z >= 0.5 || n % 2 === 0) c.fillText(String(n), x, hr.height / 2 + 3); } }
        else { const len = k % (sub / 2) === 0 ? 5 : 2; c.beginPath(); c.moveTo(x, hr.height / 2 - len / 2); c.lineTo(x, hr.height / 2 + len / 2); c.stroke(); }
      }
    }
    state = { g, hr, origin, pxPt, textL, textR, colL, colR, pageL, pageR, x0: g.x0, w: g.w };
    if (E.readOnly || L.app.view === 'reading') return;
    /* indent markers */
    const pp = E.curPara();
    const ind = pp.ind || {};
    const l = (ind.l || 0) * pxPt, fl = (ind.fl || 0) * pxPt, r = (ind.r || 0) * pxPt;
    const base = colL;
    addMark('first', base + l + fl, 1, 'First Line Indent');
    addMark('hang', base + l, hr.height - 12, 'Hanging Indent');
    addMark('left', base + l, hr.height - 6, 'Left Indent');
    addMark('right', colR - r, hr.height - 9, 'Right Indent');
    /* tab stops (custom ones) */
    for (const t of pp.tabs || []) {
      if (t.al === 'clear' || t.al === 'num') continue;
      const m = h('div', { class: 'rmark tabstop', 'data-tip': ({ left: 'Left Tab', center: 'Center Tab', right: 'Right Tab', decimal: 'Decimal Tab', bar: 'Bar Tab', start: 'Left Tab', end: 'Right Tab' }[t.al] || 'Tab'), html: SVG[{ start: 'left', end: 'right' }[t.al] || t.al] || SVG.left });
      m.style.left = base + t.pos * pxPt - 1 + 'px';
      m.dataset.kind = 'tab';
      m.dataset.pos = t.pos;
      H.appendChild(m);
    }
    /* default tab ticks after the last custom tab */
    c.strokeStyle = L.lookColor('ruler-deftab', '#777');
    const defTab = (doc().settings.defTab || 36) * pxPt;
    const lastTab = Math.max(l, ...(pp.tabs || []).map((t) => t.pos * pxPt));
    if (defTab > 4) for (let x = Math.ceil((lastTab + 1) / defTab) * defTab; base + x < colR; x += defTab) { c.beginPath(); c.moveTo(base + x + 0.5, hr.height - 3); c.lineTo(base + x + 0.5, hr.height - 1); c.stroke(); }
    /* column / cell boundary markers */
    if (g.ci) {
      const tr = g.frag.closest('tr');
      if (tr) { const tds = Array.from(tr.children).filter((x) => x.dataset.cid); tds.forEach((td, i) => { if (i === 0) return; const rr = td.getBoundingClientRect(); const m = h('div', { class: 'rmark colmark', 'data-tip': 'Move Table Column' }); m.style.left = rr.left - hr.left + 'px'; m.dataset.kind = 'cell'; m.dataset.cid = tds[i - 1].dataset.cid; H.appendChild(m); }); const last = tds[tds.length - 1]; if (last) { const rr = last.getBoundingClientRect(); const m = h('div', { class: 'rmark colmark', 'data-tip': 'Move Table Column' }); m.style.left = rr.right - hr.left + 'px'; m.dataset.kind = 'cell'; m.dataset.cid = last.dataset.cid; H.appendChild(m); } }
    } else if (g.pgEl && g.sect.cols && g.sect.cols.n > 1) {
      const cols = Array.from(g.colEl.parentNode.querySelectorAll('.pg-col'));
      cols.slice(1).forEach((col) => { const rr = col.getBoundingClientRect(); const m = h('div', { class: 'rmark colmark', 'data-tip': 'Move Column' }); m.style.left = rr.left - hr.left - 4 + 'px'; m.style.width = '8px'; m.dataset.kind = 'col'; H.appendChild(m); });
    }
    drawV();
  };
  function addMark(kind, x, y, tip) {
    const m = h('div', { class: 'rmark ind-' + kind, 'data-tip': tip, html: MARK[kind] });
    m.style.left = x + 'px';
    m.style.top = y + 'px';
    m.dataset.kind = kind;
    H.appendChild(m);
    return m;
  }
  function drawV() {
    if (!V || V.hidden || !state || !state.g.pgEl) { if (V) { const cv = V.querySelector('canvas'); if (cv) cv.getContext('2d').clearRect(0, 0, cv.width, cv.height); } return; }
    const g = state.g;
    const vr = V.getBoundingClientRect();
    const cv = V.querySelector('canvas');
    const dpr = window.devicePixelRatio || 1;
    cv.width = Math.max(1, vr.width * dpr); cv.height = Math.max(1, vr.height * dpr);
    cv.style.width = vr.width + 'px'; cv.style.height = vr.height + 'px';
    const c = cv.getContext('2d');
    c.setTransform(dpr, 0, 0, dpr, 0, 0);
    c.clearRect(0, 0, vr.width, vr.height);
    const z = g.z, pxPt = D.PX * z;
    const top = g.pr.top - vr.top, bottom = g.pr.bottom - vr.top;
    const pg = g.pg;
    c.fillStyle = L.lookColor('ruler-off', '#c9d6ee'); c.fillRect(4, top, vr.width - 8, bottom - top);
    const tT = top + pg.bodyTop * z, tB = bottom - pg.bodyBottomGap * z;
    c.fillStyle = L.lookColor('ruler', '#fff'); c.fillRect(4, tT, vr.width - 8, Math.max(0, tB - tT));
    const unit = D.unit === 'cm' || D.unit === 'mm' ? 28.3465 : 72;
    const sub = unit === 72 ? 8 : 4;
    const step = (unit / sub) * pxPt;
    c.strokeStyle = L.lookColor('ruler-tick', '#444'); c.fillStyle = L.lookColor('ruler-ink', '#222'); c.font = '9px Tahoma, Arial, sans-serif'; c.textAlign = 'center';
    if (step > 1.5) {
      const kmin = Math.ceil((Math.max(0, top) - tT) / step), kmax = Math.floor((Math.min(vr.height, bottom) - tT) / step);
      for (let k = kmin; k <= kmax; k++) {
        const y = Math.round(tT + k * step) + 0.5;
        if (k % sub === 0) { if (k) { c.save(); c.translate(vr.width / 2 + 3, y); c.rotate(-Math.PI / 2); c.fillText(String(Math.abs(k / sub)), 0, 0); c.restore(); } }
        else { const len = k % (sub / 2) === 0 ? 5 : 2; c.beginPath(); c.moveTo(vr.width / 2 - len / 2, y); c.lineTo(vr.width / 2 + len / 2, y); c.stroke(); }
      }
    }
    /* table row boundaries */
    if (g.ci) { const tblEl = g.frag.closest('table.tbl'); if (tblEl) { c.fillStyle = L.lookColor('ruler-row', '#7f95bd'); for (const tr of tblEl.querySelectorAll(':scope > tbody > tr')) { const rr = tr.getBoundingClientRect(); c.fillRect(4, rr.bottom - vr.top - 1, vr.width - 8, 2); } } }
    state.v = { vr, tT, tB, top, bottom };
  }

  /* ---------- dragging ---------- */
  const guide = () => { const sc = LY.scroller; const gl = h('div', { class: 'r-guide' }); sc.appendChild(gl); return gl; };
  function placeGuide(gl, clientX) { const sc = LY.scroller; const sr = sc.getBoundingClientRect(); gl.style.left = clientX - sr.left + sc.scrollLeft + 'px'; gl.style.top = sc.scrollTop + 'px'; gl.style.height = sc.clientHeight + 'px'; }
  const snap = (pt, alt) => (alt ? L.round(pt, 2) : Math.round(pt / ((D.unit === 'cm' || D.unit === 'mm') ? 2.83465 * 2.5 : 4.5)) * ((D.unit === 'cm' || D.unit === 'mm') ? 2.83465 * 2.5 : 4.5));
  function onRulerDown(e) {
    if (!state || e.button !== 0 || E.readOnly) return;
    e.preventDefault();
    const st = state;
    const mk = e.target.closest('.rmark');
    const xRel = e.clientX - st.hr.left;
    const toPt = (cx) => (cx - st.hr.left - st.colL) / st.pxPt;
    const gl = guide();
    placeGuide(gl, e.clientX);
    let kind = mk ? mk.dataset.kind : null;
    /* margins: drag the white/blue boundary */
    if (!mk && !st.g.ci && st.g.pgEl && (Math.abs(xRel - st.textL) < 4 || Math.abs(xRel - st.textR) < 4)) kind = Math.abs(xRel - st.textL) < 4 ? 'marginL' : 'marginR';
    /* click in the text area adds a tab stop (or sets an indent if the selector shows one) */
    if (!kind) {
      if (xRel < st.colL || xRel > st.colR) { gl.remove(); return; }
      const k = TAB_KINDS[tabKind][0];
      if (k === 'firstLine' || k === 'hanging') kind = k === 'firstLine' ? 'first' : 'hang';
      else kind = 'newtab';
    }
    const startPos = mk ? +mk.dataset.pos : null;
    const mv = (ev) => { placeGuide(gl, ev.clientX); if (mk) mk.style.left = ev.clientX - st.hr.left - (kind === 'tab' ? 1 : 0) + 'px'; };
    const up = (ev) => {
      window.removeEventListener('pointermove', mv, true);
      window.removeEventListener('pointerup', up, true);
      gl.remove();
      const pt = snap(toPt(ev.clientX), ev.altKey);
      const off = ev.clientY > st.hr.bottom + 18 || ev.clientY < st.hr.top - 18;
      switch (kind) {
        case 'newtab': {
          const k = TAB_KINDS[tabKind][0];
          if (pt <= 0) return;
          E.formatPara((pPr) => { pPr.tabs = (pPr.tabs || []).filter((t) => Math.abs(t.pos - pt) > 0.5).concat([{ pos: pt, al: k }]).sort((a, b) => a.pos - b.pos); }, 'Set Tab');
          break;
        }
        case 'tab':
          E.formatPara((pPr, p) => {
            const eff = D.pProps(doc(), p).tabs || [];
            let tabs = (pPr.tabs || []).filter((t) => Math.abs(t.pos - startPos) > 0.5);
            const fromStyle = eff.find((t) => Math.abs(t.pos - startPos) < 0.5) && !(pPr.tabs || []).find((t) => Math.abs(t.pos - startPos) < 0.5);
            if (fromStyle) tabs.push({ pos: startPos, al: 'clear' });
            if (!off) { const orig = eff.find((t) => Math.abs(t.pos - startPos) < 0.5) || { al: 'left' }; tabs = tabs.filter((t) => Math.abs(t.pos - pt) > 0.5); tabs.push(Object.assign({}, orig, { pos: pt })); }
            pPr.tabs = tabs.sort((a, b) => a.pos - b.pos);
          }, off ? 'Clear Tab' : 'Move Tab');
          break;
        case 'first': E.formatPara((pPr, p) => { const cur = D.pProps(doc(), p).ind || {}; pPr.ind = Object.assign({}, pPr.ind, { fl: L.round(pt - (cur.l || 0), 2) }); }, 'First Line Indent'); break;
        case 'hang': E.formatPara((pPr, p) => { const cur = D.pProps(doc(), p).ind || {}; const firstAbs = (cur.l || 0) + (cur.fl || 0); pPr.ind = Object.assign({}, pPr.ind, { l: L.round(Math.max(-72, pt), 2), fl: L.round(firstAbs - pt, 2) }); }, 'Hanging Indent'); break;
        case 'left': E.formatPara((pPr, p) => { const cur = D.pProps(doc(), p).ind || {}; pPr.ind = Object.assign({}, pPr.ind, { l: L.round(Math.max(-72, pt), 2), fl: cur.fl || 0 }); }, 'Left Indent'); break;
        case 'right': { const fromRight = (st.colR - (ev.clientX - st.hr.left)) / st.pxPt; E.formatPara((pPr) => { pPr.ind = Object.assign({}, pPr.ind, { r: L.round(snap(fromRight, ev.altKey), 2) }); }, 'Right Indent'); break; }
        case 'marginL': case 'marginR': {
          const s = L.app.curSect();
          const dx = (ev.clientX - e.clientX) / st.pxPt;
          E.edit('Page Setup', () => { D.touchKey(s.holder, s.key); const sect = s.holder === doc() ? doc().sect : s.holder.sect; if (kind === 'marginL') sect.ml = L.round(Math.max(0, sect.ml + dx), 2); else sect.mr = L.round(Math.max(0, sect.mr - dx), 2); return E.sel; });
          break;
        }
        case 'cell': {
          const tblEl = st.g.frag && st.g.frag.closest('table.tbl');
          const tbl = tblEl ? L.tables.findTable(+tblEl.dataset.tid) : null;
          if (!tbl) break;
          const map = R.tblMap(tbl);
          let mm = null;
          for (const row of map) for (const m of row) if (m.cell.id === +mk.dataset.cid) mm = m;
          if (!mm) break;
          const k = mm.c0 + mm.span - 1;
          const dx = (ev.clientX - e.clientX) / st.pxPt;
          E.edit('Column Width', () => { D.touchTbl(tbl); if (k + 1 < tbl.grid.length && !ev.shiftKey) { const dd = L.clamp(dx, -(tbl.grid[k] - 6), tbl.grid[k + 1] - 6); tbl.grid[k] += dd; tbl.grid[k + 1] -= dd; } else tbl.grid[k] = Math.max(6, tbl.grid[k] + dx); L.tables.fixWidths(tbl); return E.sel; });
          break;
        }
        case 'col': {
          const s = L.app.curSect();
          const dx = (ev.clientX - e.clientX) / st.pxPt;
          E.edit('Columns', () => { D.touchKey(s.holder, s.key); const sect = s.holder === doc() ? doc().sect : s.holder.sect; sect.cols = Object.assign({}, sect.cols, { space: L.round(Math.max(0, (sect.cols.space || 36) - dx * 2 / Math.max(1, sect.cols.n - 1)), 2), eq: true }); return E.sel; });
          break;
        }
        default: break;
      }
      RU.draw();
    };
    window.addEventListener('pointermove', mv, true);
    window.addEventListener('pointerup', up, true);
  }
  function onVDown(e) {
    if (!state || !state.v || e.button !== 0 || E.readOnly) return;
    const v = state.v;
    const y = e.clientY - v.vr.top;
    const which = Math.abs(y - v.tT) < 4 ? 'mt' : Math.abs(y - v.tB) < 4 ? 'mb' : null;
    if (!which) return;
    e.preventDefault();
    const y0 = e.clientY;
    const sc = LY.scroller;
    const gl = h('div', { class: 'r-guide h' });
    sc.appendChild(gl);
    const place = (cy) => { const sr = sc.getBoundingClientRect(); gl.style.cssText = `top:${cy - sr.top + sc.scrollTop}px;left:${sc.scrollLeft}px;width:${sc.clientWidth}px;height:0;border-left:0;border-top:1px dotted #000`; };
    place(e.clientY);
    const mv = (ev) => place(ev.clientY);
    const up = (ev) => {
      window.removeEventListener('pointermove', mv, true);
      window.removeEventListener('pointerup', up, true);
      gl.remove();
      const dy = (ev.clientY - y0) / state.pxPt;
      if (Math.abs(dy) < 0.5) return;
      const s = L.app.curSect();
      E.edit('Page Setup', () => { D.touchKey(s.holder, s.key); const sect = s.holder === doc() ? doc().sect : s.holder.sect; if (which === 'mt') sect.mt = L.round(Math.max(0, sect.mt + dy), 2); else sect.mb = L.round(Math.max(0, sect.mb - dy), 2); return E.sel; });
    };
    window.addEventListener('pointermove', mv, true);
    window.addEventListener('pointerup', up, true);
  }
  RU.tabKind = () => TAB_KINDS[tabKind][0];
})();
