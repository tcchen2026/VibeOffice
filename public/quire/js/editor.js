/* Quire — editing surface. The pages live in one contentEditable host; the browser draws the caret and
 * selection natively, but every change is intercepted (beforeinput / keydown / clipboard) and applied to
 * the model, after which the affected pages are re-laid out and the selection restored.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, LY = L.layout;
  const E = (L.ed = {});
  E.sel = null; /* {a:{p,o}, f:{p,o}} */
  E.pending = null; /* formatting for the next typed text at a collapsed caret */
  E.overtype = false;
  E.extend = false;
  E.composing = false;
  E.objSel = null; /* {p, o, it} selected inline/floating object */
  E.readOnly = false;
  const doc = () => D.doc;

  /* ---------- selection model ---------- */
  E.collapsed = () => !E.sel || D.eqPos(E.sel.a, E.sel.f);
  E.range = () => (E.sel ? D.order(doc(), E.sel.a, E.sel.f) : null);
  E.caret = () => (E.sel ? E.sel.f : null);
  E.story = () => (E.sel ? D.storyOf(doc(), E.sel.f.p) : doc().main);
  E.setSel = function (a, f, opts) {
    if (!a) return;
    a = clampPos(a);
    f = clampPos(f || a);
    E.sel = { a, f };
    if (!opts || !opts.keepObj) E.clearObjSel();   /* a text selection ends an object selection */
    if (!opts || !opts.keepPending) E.pending = null;
    if (!opts || !opts.noDom) E.restoreDom(opts && opts.noScroll);
    E.selChanged();
  };
  /** the one way into an object selection, os = { p, o, it, el, float }: the text selection becomes the object's
      anchor, and the page's own selection covers the object (inline) or goes (floating), so no text stays
      highlighted. Inline handles are drawn here, floating ones by drawing.js. */
  E.selectObject = function (os) {
    if (E.objSel && E.objSel.el !== os.el) E.clearObjSel();
    E.objSel = os;
    if (LY.root) LY.root.classList.toggle('objsel', !!os.float);
    E.setSel(D.pos(os.p, os.o), D.pos(os.p, os.o + 1), { noScroll: true, keepObj: true });
    if (!os.float) showHandles(os.el);
    L.bus.emit('objsel', os);
  };
  /** the one way an object selection ends: its handles go with it (the inline ones here, the floating
      ones in drawing.js, which follows 'objsel') */
  E.clearObjSel = function () {
    hideHandles();
    if (!E.objSel) return;
    E.objSel = null;
    if (LY.root) LY.root.classList.remove('objsel');
    L.bus.emit('objsel', null);
  };
  function clampPos(pos) { return { p: pos.p, o: L.clamp(pos.o, 0, D.plen(pos.p)) }; }
  E.selectedParas = function () {
    if (!E.sel) return [];
    const [a, b] = E.range();
    return D.parasBetween(doc(), a.p, b.p);
  };
  /** serialised selection for undo history */
  D.getSel = function () {
    if (!E.sel) return null;
    return { a: { pid: E.sel.a.p.id, o: E.sel.a.o }, f: { pid: E.sel.f.p.id, o: E.sel.f.o } };
  };
  E.applySerialized = function (s) {
    if (!s) return false;
    const d = doc();
    const a = D.byId(d, s.a.pid), f = D.byId(d, s.f.pid);
    if (!a || !f) return false;
    E.sel = { a: clampPos({ p: a, o: s.a.o }), f: clampPos({ p: f, o: s.f.o }) };
    return true;
  };

  /* ---------- DOM selection sync ---------- */
  let settingDom = 0;
  let lastDom = null;
  const domIsOurs = () => { const s = window.getSelection(); return !!lastDom && s.anchorNode === lastDom.an && s.anchorOffset === lastDom.ao && s.focusNode === lastDom.fn && s.focusOffset === lastDom.fo; };
  E.restoreDom = function (noScroll) {
    if (!E.sel || !LY.root) return;
    E.drawSel(false);
    const s = window.getSelection();
    /* a selected floating object: the page holds a (hidden) caret at its anchor, so no text shows selected (Chrome
       puts a caret back into a focused page left with no selection at all) */
    const a = LY.posToDom(E.sel.a), f = E.objSel && E.objSel.float ? a : LY.posToDom(E.sel.f);
    if (!a || !f) return;
    settingDom++;
    try {
      if (s.anchorNode !== a.node || s.anchorOffset !== a.offset || s.focusNode !== f.node || s.focusOffset !== f.offset) s.setBaseAndExtent(a.node, a.offset, f.node, f.offset);
    } catch (e) { /* ignore */ }
    /* remember exactly what we put there, so a later read can tell our own selection from the user's */
    lastDom = { an: s.anchorNode, ao: s.anchorOffset, fn: s.focusNode, fo: s.focusOffset };
    setTimeout(() => { settingDom = Math.max(0, settingDom - 1); }, 0);
    if (!noScroll) E.scrollToCaret();
    E.markCells();
  };
  E.readDom = function () {
    const s = window.getSelection();
    if (!s.rangeCount || !LY.root) return null;
    if (!isEditableNode(s.anchorNode) || !isEditableNode(s.focusNode)) return null;
    const a = LY.domToPos(s.anchorNode, s.anchorOffset);
    const f = LY.domToPos(s.focusNode, s.focusOffset);
    if (!a || !f) return null;
    const d = doc();
    if (D.storyOf(d, a.p) !== D.storyOf(d, f.p)) return { a, f: a };
    return { a, f };
  };
  function isEditableNode(n) {
    if (!n) return false;
    const e = n.nodeType === 3 ? n.parentNode : n;
    return !!(e && e.closest && (e.closest('#pages') || e.closest('.ed-host')));
  }
  /** the part of the page not being edited: the body while a header or footer is, a header or footer while the
      body is; a selection there is put back (keys would edit text the user can't see being edited) */
  const offLimits = (n) => { const e = n && (n.nodeType === 3 ? n.parentNode : n); return !!(e && e.closest && e.closest(LY.hfEdit ? '.pg-body' : '.pg-hdr, .pg-ftr') && !e.closest('.tbx')); };
  /** while a header or footer is edited, the page it is edited on follows the selection */
  const hfPage = (n) => { const e = n && (n.nodeType === 3 ? n.parentNode : n), pg = e && e.closest && e.closest('.pg'); if (LY.hfEdit && pg && pg.dataset.pg != null) LY.hfEdit.page = +pg.dataset.pg; };
  E.onSelectionChange = function () {
    if (E.composing) return;
    if (E.sel && domIsOurs()) return;   /* our own (the caret held at a floating object's anchor, a selection-bar drag) */
    const s = window.getSelection();
    if (!s.rangeCount || !isEditableNode(s.focusNode)) return;
    const m = E.readDom();
    /* caret landed somewhere unmapped (page gap, margin), or a drag reached the empty part of a page (the
       browser puts it at the page's first position) — snap to the nearest paragraph */
    if (!m || (pointer && !settingDom && !overText(pointer))) {
      if (!settingDom) snapCaret();
      return;
    }
    const changed = !E.sel || !D.eqPos(m.a, E.sel.a) || !D.eqPos(m.f, E.sel.f);
    if (!changed) return;
    if (domIsOurs() && E.sel) return; /* our own restore landed on an equivalent DOM point */
    if (offLimits(s.focusNode)) { if (E.sel) E.restoreDom(true); return; }
    hfPage(s.focusNode);
    E.pending = null; D.breakMerge();
    E.sel = m;
    E.clearObjSel();
    E.drawSel(false);
    E.selChanged();
    E.markCells();
  };
  /* The model's selection drawn while the keyboard is elsewhere (Find and Replace keeps it): putting it into the
     page would take the focus with it, as Chrome focuses the editable page a selection lands in. Drawn again after
     each layout, gone when the page's own selection is set. */
  let selDrawn = false;
  E.drawSel = function (on) {
    const can = typeof CSS !== 'undefined' && CSS.highlights && typeof Highlight === 'function';
    selDrawn = !!on && can;
    if (!can) return;
    const [a, b] = on && E.sel ? E.range().map((p) => LY.posToDom(p)) : [];
    if (!a || !b) { CSS.highlights.delete('quire-sel'); return; }
    const r = document.createRange();
    try { r.setStart(a.node, a.offset); r.setEnd(b.node, b.offset); } catch (e) { CSS.highlights.delete('quire-sel'); return; }
    CSS.highlights.set('quire-sel', new Highlight(r));
  };
  L.bus.on('layout-done', () => { if (selDrawn) E.drawSel(true); });
  /** pull a pending DOM selection into the model before acting on a key (selectionchange is asynchronous) */
  E.syncSel = function () {
    if (E.composing || !LY.root || domIsOurs()) return;
    const s = window.getSelection();
    if (!s.rangeCount || !isEditableNode(s.focusNode)) return;
    const m = E.readDom();
    if (!m) { snapCaret(); return; }
    if (E.sel && D.eqPos(m.a, E.sel.a) && D.eqPos(m.f, E.sel.f)) return;
    if (offLimits(s.focusNode)) { if (E.sel) E.restoreDom(true); return; }
    hfPage(s.focusNode);
    E.pending = null; D.breakMerge();
    E.sel = m; E.clearObjSel(); E.selChanged();
  };
  /* A press in the document is a ui.drag (common/ui.js) that watches the browser's own text selection: it holds no
     pointer and draws nothing. Its release ends the selection ('sel-done'); a release the page never sees (let go in
     another window, cancelled), Esc or a second button end it without one, so nothing acts on a selection the user
     didn't finish; a drag of the page's own (an object, the selection bar, a tool) takes over from it. While it
     lasts, `pointer` is where the pointer is: a selection dragged out of the text follows it. */
  let pointer = null, textDrag = false, snapQueued = false, press = null;
  document.addEventListener('pointerdown', (e) => {
    const t = e.target;
    if (e.button !== 0 || !isEditableNode(t)) return;
    textDrag = !!(t.closest && t.closest('.p') && LY.root && LY.root.contains(t));
    pointer = { x: e.clientX, y: e.clientY };
    const done = () => { if (press === d) { press = null; pointer = null; textDrag = false; } };
    const d = L.ui.drag(e, {
      capture: false, threshold: 0, cursor: false,
      move(ev) {
        pointer = { x: ev.clientX, y: ev.clientY };
        /* outside the page (grey area, other panes) the browser stops extending the selection: follow the pointer */
        if (textDrag && !overText(pointer) && !snapQueued) { snapQueued = true; requestAnimationFrame(() => { snapQueued = false; if (pointer && textDrag) snapCaret(); }); }
      },
      up() { done(); setTimeout(selDone, 0); },
      cancel: done,
    });
    press = d;
  }, true);
  /** a mouse button is down in the document (a selection may still be dragged out) */
  E.pressing = () => !!press;
  /** the user let go of the mouse in the document: 'sel' fires all through a drag, 'sel-done' once, when
      the text selection (or caret) is made; the Format Painter and the highlighter act on it */
  function selDone() {
    if (E.objSel) return; /* a press on a shape or picture selected it, not text */
    E.syncSel();
    if (E.sel && !E.objSel) L.bus.emit('sel-done', E.sel);
  }
  /** the selection's moving end landed where there is no text (page gap, margins): put it at the nearest
      paragraph — its end when the point is below it, its start when above — and, when a selection is being
      dragged out, keep where it started */
  const overText = (pt) => { const e = document.elementFromPoint(pt.x, pt.y); return !!(e && e.closest && e.closest('.p')); };
  function snapCaret() {
    const s = window.getSelection();
    const n = s.focusNode;
    if (!n) return;
    let pt = pointer;
    if (!pt) {
      let r = null;
      try { const rg = document.createRange(); rg.setStart(n, s.focusOffset); rg.collapse(true); r = rg.getClientRects()[0] || rg.getBoundingClientRect(); } catch (err) { r = null; }
      if (!r || (!r.top && !r.height)) { const e = n.nodeType === 3 ? n.parentNode : n; r = e.getBoundingClientRect ? e.getBoundingClientRect() : null; }
      if (!r) return;
      pt = { x: r.left, y: r.top + r.height / 2 };
    }
    /* a selection being dragged stays in the story it started in (body, a header, a text box…) */
    const anchor = !s.isCollapsed && isEditableNode(s.anchorNode) ? LY.domToPos(s.anchorNode, s.anchorOffset) : null;
    const story = anchor ? D.storyOf(doc(), anchor.p) : null;
    /* in a text box's (or a shape's) text area, away from its lines: the nearest of its own */
    const under = !story && document.elementFromPoint(pt.x, pt.y), tbx = under && under.closest && under.closest('.tbx');
    let frag = null, bd = Infinity, below = false, level = false, box = null;
    for (const f of (tbx || LY.root).querySelectorAll(story || tbx ? '.p' : '.pg-body .p, .pg-hdr .p, .pg-ftr .p')) {
      if (story) { const q = D.byId(doc(), +f.dataset.pid); if (!q || D.storyOf(doc(), q) !== story) continue; }
      const fr = f.getBoundingClientRect();
      if (!fr.height) continue;
      const dy = pt.y < fr.top ? fr.top - pt.y : pt.y > fr.bottom ? pt.y - fr.bottom : 0;
      const dx = pt.x < fr.left ? fr.left - pt.x : pt.x > fr.right ? pt.x - fr.right : 0;
      const dd = dy * 4 + dx;
      if (dd < bd) { bd = dd; frag = f; below = pt.y > fr.bottom; level = !dy; box = fr; }
    }
    if (!frag) frag = LY.root.querySelector('.p');
    if (!frag) return;
    const p = D.byId(doc(), +frag.dataset.pid);
    if (!p) return;
    const to = parseFloat(frag.dataset.to);
    let at = { p, o: below && isFinite(to) ? to : +frag.dataset.from || 0 };
    /* beside a paragraph (in the margin): the start or end of the line at that height, as Word does */
    if (level && box && document.caretRangeFromPoint) {
      const pc = frag.querySelector('.pc') || frag, cb = pc.getBoundingClientRect();
      const rg = document.caretRangeFromPoint(L.clamp(pt.x, cb.left + 1, cb.right - 1), L.clamp(pt.y, cb.top + 1, cb.bottom - 1));
      const q = rg && frag.contains(rg.startContainer) ? LY.domToPos(rg.startContainer, rg.startOffset) : null;
      if (q) at = q;
    }
    if (anchor && D.storyOf(doc(), anchor.p) === D.storyOf(doc(), at.p)) E.setSel(anchor, at, { noScroll: true });
    else E.setSel(at, null, pointer ? { noScroll: true } : undefined);
  }
  /** notify UI (debounced through rAF) */
  E.selChanged = L.rafThrottle(() => { L.bus.emit('sel', E.sel); });

  /** bring the caret into view: the page's selection, or (fromModel) the model's, when the page's isn't moved */
  E.scrollToCaret = function (fromModel) {
    const sc = LY.scroller;
    if (!sc || !E.sel) return;
    const s = window.getSelection();
    const at = fromModel ? LY.posToDom(E.sel.f) : s.rangeCount ? { node: s.focusNode, offset: s.focusOffset } : null;
    if (!at) return;
    const rng = document.createRange();
    try { rng.setStart(at.node, at.offset); rng.collapse(true); } catch (e) { return; }
    let r = rng.getClientRects()[0] || rng.getBoundingClientRect();
    if (!r || (!r.height && !r.top)) {
      const n = at.node && (at.node.nodeType === 3 ? at.node.parentNode : at.node);
      if (n && n.getBoundingClientRect) r = n.getBoundingClientRect();
    }
    if (!r) return;
    const sr = sc.getBoundingClientRect();
    const m = 30;
    if (r.top < sr.top + 4) sc.scrollTop -= sr.top - r.top + m;
    else if (r.bottom > sr.bottom - 4) sc.scrollTop += r.bottom - sr.bottom + m;
    if (r.left < sr.left) sc.scrollLeft -= sr.left - r.left + m;
    else if (r.right > sr.right) sc.scrollLeft += r.right - sr.right + m;
  };
  E.caretRect = function () {
    const s = window.getSelection();
    if (!s.rangeCount) return null;
    const rng = document.createRange();
    try { rng.setStart(s.focusNode, s.focusOffset); rng.collapse(true); } catch (e) { return null; }
    return rng.getClientRects()[0] || null;
  };

  /* ---------- editing primitives ---------- */
  /** run fn inside a transaction; fn returns the new caret position or selection */
  E.edit = function (label, fn, opts) {
    if (E.guard()) return false;
    let res;
    const committed = D.tx(label, () => {
      res = fn();
      if (res && res.p) E.sel = { a: res, f: res };
      else if (res && res.a) E.sel = res;
    }, opts);
    return committed !== false;
  };
  /** refuse edits in protected documents */
  E.guard = function () {
    if (E.readOnly) { L.ui.toast('This document is read-only.'); return true; }
    const pr = doc().settings.protect;
    if (pr && pr.on) {
      if (pr.kind === 'readOnly' || pr.kind === 'forms') { L.ui.toast('This modification is not allowed because the selection is locked.'); return true; }
      if (pr.kind === 'comments' && E.story().kind !== 'cmt') { L.ui.toast('Only comments can be added to this protected document.'); return true; }
    }
    return false;
  };
  /** delete the current selection (if any) and return the caret position; o.wholeTable (Cut, Backspace): a whole
   *  table selected goes, as in Word, where Delete empties its cells */
  E.deleteSelection = function (o) {
    if (E.collapsed()) return E.sel.f;
    const [a, b] = E.range();
    if (E.cellSel && o && o.wholeTable && !O.tracking()) { const r = removeTable(); if (r) return r; }
    if (E.cellSel) { const r = clearCells(); if (r) return r; }
    if (E.rowSel && !O.tracking()) return deleteRowSel(a, b);   /* tracked, the text is marked deleted as any (rows can't be) */
    const pos = O.deleteRange(a, b);
    O.repairFields(D.storyOf(doc(), pos.p));
    return pos;
  };
  /** a selection from inside a table out of it (E.rowSel): the text outside the table goes, then the rows */
  function deleteRowSel(a, b) {
    const rs = E.rowSel, d = doc();
    E.rowSel = null;
    const i = rs.cont.blocks.indexOf(rs.tbl);
    let pos;
    if (rs.down) { const q = D.firstPara(rs.cont.blocks.slice(i + 1)); pos = q ? O.deleteRange(D.pos(q, 0), b) : b; }
    else { const q = D.lastPara(rs.cont.blocks.slice(0, i)); pos = q ? O.deleteRange(a, D.pos(q, D.plen(q))) : a; }
    if (rs.r0 === 0 && rs.r1 >= rs.tbl.rows.length - 1) O.removeBlock(rs.tbl);
    else if (L.tables && L.tables.removeRows) L.tables.removeRows(rs.tbl, rs.r0, rs.r1);
    d._idxDirty = true;
    O.repairFields(D.storyOf(d, pos.p));
    return pos;
  }
  /* every cell of a table selected: the table itself goes (tracked, its cells are emptied as before) */
  function removeTable() {
    const t = E.cellSel && E.cellSel.tbl;
    if (!t || !L.tables.wholeSelected()) return null;
    E.cellSel = null;
    return O.removeBlock(t);
  }
  function clearCells() {
    const cs = E.cellSel;
    if (!cs) return null;
    let first = null;
    for (const c of cs.cells) {
      D.touchList(c);
      const p = c.blocks.find((b) => b.t === 'p') || D.para();
      D.touch(p);
      p.runs = [];
      c.blocks = [p];
      if (!first) first = D.pos(p, 0);
    }
    doc()._idxDirty = true;
    E.cellSel = null;
    return first;
  }
  E.typeText = function (text) {
    if (!E.sel) return;
    const d = doc();
    if (E.objSel && E.objSel.float) return;
    E.edit(L.N('Typing'), () => {
      let pos = E.deleteSelection();
      if (E.overtype && E.collapsed()) {
        const n = D.plen(pos.p);
        const k = Math.min(text.length, n - pos.o);
        if (k > 0) O.deleteRange(pos, D.pos(pos.p, pos.o + k));
      }
      const r = E.pending ? D.mergeRPr(O.inheritRPr(pos), E.pending) : null;
      const st = d.settings;
      void st;
      pos = O.insertText(pos, text, r);
      return pos;
    }, { merge: 'type' });
    const keep = E.pending;
    if (keep && E.sel) E.pending = keep;
    if (L.autocorrect && text.length === 1) L.autocorrect.afterType(text);
  };
  E.enter = function (kind) {
    if (!E.sel) return;
    if (!kind && L.autocorrect && L.autocorrect.acceptTip && L.autocorrect.acceptTip()) return;
    /* finishing a paragraph finishes its last word: run the as-you-type corrections on it */
    if (L.autocorrect && E.collapsed() && E.sel.f.o > 0 && /\S/.test(D.ptext(E.sel.f.p, E.sel.f.o - 1, E.sel.f.o))) L.autocorrect.afterType('\n', true);
    const d = doc();
    if (kind === 'line') return E.insertItem(D.item('br', { type: 'line' }), 'Line Break');
    if (kind === 'page') return E.insertItem(D.item('br', { type: 'page' }), 'Page Break');
    if (kind === 'column') return E.insertItem(D.item('br', { type: 'column' }), 'Column Break');
    const p = E.sel.f.p;
    /* Enter on an empty list item ends the list */
    if (E.collapsed() && !D.plen(p) && LY.labels && LY.labels.get(p.id)) {
      const pp = D.pProps(d, p);
      const lvl = pp.num ? pp.num.lvl || 0 : 0;
      E.edit(L.N('Numbering'), () => {
        D.touch(p);
        if (lvl > 0) p.pPr.num = { id: pp.num.id, lvl: lvl - 1 };
        else { p.pPr.num = { id: '0', lvl: 0 }; delete p.pPr.ind; if (D.styleProps(d, p.pPr.style || 'Normal').pPr.num) p.pPr.style = undefined; }
        return D.pos(p, 0);
      });
      return;
    }
    E.edit(L.N('Typing'), () => {
      const pos = E.deleteSelection();
      const np = O.splitPara(pos);
      /* AutoFormat: "---" + Enter → border line */
      return np;
    });
    if (L.autocorrect) L.autocorrect.afterEnter();
  };
  E.insertItem = function (item, label) {
    E.edit(label || L.N('Insert'), () => O.insertItem(E.deleteSelection(), item));
  };
  E.insertTextAt = (text, label) => E.edit(label || L.N('Insert'), () => O.insertPlain(E.deleteSelection(), text, E.pending ? D.mergeRPr(O.inheritRPr(E.sel.f), E.pending) : null));
  E.backspace = function (word) {
    if (!E.sel) return;
    const d = doc();
    if (!E.collapsed()) { E.edit(L.N('Delete'), () => E.deleteSelection({ wholeTable: true })); return; }
    const pos = E.sel.f;
    const p = pos.p;
    if (pos.o === 0) {
      /* list number / indent removal first */
      const lb = LY.labels && LY.labels.get(p.id);
      const pp = D.pProps(d, p);
      if (lb) {
        E.edit(L.N('Numbering'), () => {
          D.touch(p);
          const ind = pp.ind || {};
          p.pPr.num = { id: '0', lvl: 0 };
          p.pPr.ind = { l: ind.l || 0, fl: 0, r: ind.r || 0 };
          return D.pos(p, 0);
        });
        return;
      }
      if (p.pPr.ind && (p.pPr.ind.fl > 0 || p.pPr.ind.l > 0) && L.store.get('acIndentBackspace', true)) {
        E.edit(L.N('Indent'), () => { D.touch(p); if (p.pPr.ind.fl > 0) p.pPr.ind.fl = 0; else p.pPr.ind.l = Math.max(0, (p.pPr.ind.l || 0) - 36); return D.pos(p, 0); });
        return;
      }
      const info = D.info(d, p);
      const idx = info.cont.blocks.indexOf(p);
      const prevB = info.cont.blocks[idx - 1];
      if (!prevB) return; /* start of story / cell */
      if (prevB.t === 'tbl') {
        /* empty paragraph after a table is removed; otherwise move into the table */
        if (!D.plen(p) && info.cont.blocks[idx + 1]) { E.edit(L.N('Delete'), () => { const r = O.removeBlock(p); void r; const lp = D.lastPara([prevB]); return D.pos(lp, D.plen(lp)); }); return; }
        const lp = D.lastPara([prevB]);
        E.setSel(D.pos(lp, D.plen(lp)));
        return;
      }
      if (O.tracking()) { E.edit(L.N('Delete'), () => O.deleteRange(D.pos(prevB, D.plen(prevB)), pos)); return; }
      E.edit(L.N('Delete'), () => O.joinNext(prevB) || pos);
      return;
    }
    let a = word ? wordStartBefore(p, pos.o) : prevCharOffset(p, pos.o);
    E.edit(L.N('Delete'), () => {
      const r = O.deleteRange(D.pos(p, a), pos);
      O.repairFields(D.storyOf(d, p));
      return O.tracking() ? D.pos(p, a) : r;
    }, { merge: word ? null : 'del' });
  };
  E.del = function (word) {
    if (!E.sel) return;
    const d = doc();
    if (!E.collapsed()) { E.edit(L.N('Delete'), () => E.deleteSelection()); return; }
    const pos = E.sel.f;
    const p = pos.p;
    const n = D.plen(p);
    if (pos.o >= n) {
      const info = D.info(d, p);
      const idx = info.cont.blocks.indexOf(p);
      const nb = info.cont.blocks[idx + 1];
      if (!nb) return;
      if (nb.t === 'tbl') return;
      if (O.tracking()) { E.edit(L.N('Delete'), () => { O.deleteRange(pos, D.pos(nb, 0)); return pos; }); return; }
      E.edit(L.N('Delete'), () => O.joinNext(p) || pos);
      return;
    }
    const b = word ? wordEndAfter(p, pos.o) : nextCharOffset(p, pos.o);
    E.edit(L.N('Delete'), () => { O.deleteRange(pos, D.pos(p, b)); O.repairFields(D.storyOf(d, p)); return pos; }, { merge: word ? null : 'fdel' });
  };
  function prevCharOffset(p, o) {
    /* skip over surrogate pairs and zero-length markers; skip deleted (tracked) text */
    const t = D.ptext(p);
    let k = o - 1;
    if (k > 0 && /[\uDC00-\uDFFF]/.test(t[k]) && /[\uD800-\uDBFF]/.test(t[k - 1])) k--;
    return Math.max(0, k);
  }
  function nextCharOffset(p, o) {
    const t = D.ptext(p);
    let k = o + 1;
    if (k < t.length && /[\uDC00-\uDFFF]/.test(t[k]) && /[\uD800-\uDBFF]/.test(t[k - 1])) k++;
    return Math.min(t.length, k);
  }
  function wordStartBefore(p, o) {
    const t = D.ptext(p);
    let k = o;
    while (k > 0 && /\s/.test(t[k - 1])) k--;
    const w = /[\p{L}\p{N}_]/u;
    if (k > 0 && w.test(t[k - 1])) while (k > 0 && w.test(t[k - 1])) k--;
    else if (k > 0) k--;
    return k;
  }
  function wordEndAfter(p, o) {
    const t = D.ptext(p);
    let k = o;
    const w = /[\p{L}\p{N}_]/u;
    if (k < t.length && w.test(t[k])) while (k < t.length && w.test(t[k])) k++;
    else if (k < t.length && !/\s/.test(t[k])) k++;
    while (k < t.length && t[k] === ' ') k++;
    return k;
  }

  /* ---------- tab key ---------- */
  E.tab = function (shift) {
    if (!E.sel) return;
    const d = doc();
    const p = E.sel.f.p;
    const ci = D.cellOf(d, p);
    if (ci) {
      L.tables.moveCell(shift ? -1 : 1);
      return;
    }
    const pp = D.pProps(d, p);
    const lb = LY.labels && LY.labels.get(p.id);
    const sel = E.selectedParas();
    const atStart = E.collapsed() ? E.sel.f.o === 0 : sel.length > 1;
    if (lb && atStart) {
      E.edit(L.N('Numbering'), () => {
        for (const q of sel.length ? sel : [p]) {
          const qp = D.pProps(d, q);
          if (!qp.num || !qp.num.id || qp.num.id === '0') continue;
          D.touch(q);
          q.pPr.num = { id: qp.num.id, lvl: L.clamp((qp.num.lvl || 0) + (shift ? -1 : 1), 0, 8) };
          if (q.pPr.ind) delete q.pPr.ind;
        }
        return E.sel;
      });
      return;
    }
    if (shift) return;
    if (!lb && atStart && E.collapsed() && D.plen(p) > 0 && L.store.get('acIndentTab', true) && !(pp.ind && pp.ind.fl > 0)) {
      /* Tab at the start of a paragraph sets the first-line indent (AutoFormat As You Type) */
      E.edit(L.N('Indent'), () => { D.touch(p); p.pPr.ind = Object.assign({}, p.pPr.ind || {}, { fl: 36 }); return E.sel.f; });
      return;
    }
    E.insertItem(D.item('tab'), L.N('Typing'));
  };

  /* ---------- formatting ---------- */
  /** apply run properties to the selection, or set pending formatting at a collapsed caret */
  E.formatRun = function (props, label) {
    if (!E.sel) return;
    if (E.collapsed()) {
      /* a caret inside a word formats the whole word (Word behaviour) */
      const w = O.wordAt(E.sel.f);
      if (w && w[0].o < E.sel.f.o && E.sel.f.o < w[1].o) {
        const keep = E.sel;
        E.edit(label || L.N('Format'), () => { O.setRunProps(w[0], w[1], props); return keep; });
        return;
      }
      E.pending = Object.assign({}, E.pending || {}, props);
      for (const k in props) if (props[k] === undefined) E.pending[k] = undefined;
      L.bus.emit('sel', E.sel);
      return;
    }
    const [a, b] = E.range();
    const keep = E.sel;
    E.edit(label || L.N('Format'), () => { O.setRunProps(a, b, props); return keep; });
  };
  /** toggle a boolean run property based on the first character's state */
  E.toggleRun = function (key, val, label) {
    const cur = E.curRun();
    const on = val === undefined ? !cur[key] : cur[key] !== val;
    const props = {};
    props[key] = on ? (val === undefined ? true : val) : undefined;
    if (key === 'vert' && on) { /* exclusive */ }
    if (key === 'caps' && on) props.smallCaps = undefined;
    if (key === 'smallCaps' && on) props.caps = undefined;
    if (!on && val === undefined) props[key] = false;
    E.formatRun(props, label);
  };
  /** formatting the toolbar shows: run props at the selection start (with pending) */
  E.curRun = function () {
    if (!E.sel) return {};
    const [a, b] = E.range();
    let pos = a;
    if (!E.collapsed() && a.o >= D.plen(a.p) && a.p !== b.p) { const np = D.nextPara(doc(), a.p); if (np) pos = D.pos(np, 0); }
    let r = E.collapsed() ? O.rPrAt(pos) : runAfter(pos);
    if (E.pending) r = D.mergeRPr(r, E.pending);
    return r;
  };
  function runAfter(pos) {
    const it = D.itemAfter(pos.p, pos.o);
    const ctx = LY.tblCtx(pos.p);
    return D.rProps(doc(), pos.p, it && !D.isMarker(it) ? it : null, ctx);
  }
  /** is a run property uniform over the selection? returns its value, or `mixed` (default undefined) when it varies */
  E.uniformRun = function (key, mixed) {
    if (!E.sel || E.collapsed()) return E.curRun()[key];
    const [a, b] = E.range();
    const d = doc();
    let val, first = true;
    for (const p of D.parasBetween(d, a.p, b.p)) {
      const s = p === a.p ? a.o : 0, e = p === b.p ? b.o : D.plen(p);
      let pos = 0;
      const ctx = LY.tblCtx(p);
      for (const it of p.runs) {
        const n = D.ilen(it);
        const s0 = pos; pos += n;
        if (!n || pos <= s || s0 >= e) continue;
        const v = D.rProps(d, p, it, ctx)[key];
        if (first) { val = v; first = false; } else if (JSON.stringify(v) !== JSON.stringify(val)) return mixed;
      }
    }
    return first ? E.curRun()[key] : val;
  };
  E.formatPara = function (fn, label) {
    if (!E.sel) return;
    const paras = E.selectedParas();
    const keep = E.sel;
    E.edit(label || 'Paragraph Formatting', () => { O.setParaProps(paras, fn); return keep; });
  };
  E.curPara = () => (E.sel ? D.pProps(doc(), E.sel.f.p, LY.tblCtx(E.sel.f.p)) : {});

  /* ---------- select helpers ---------- */
  E.selectAll = function () {
    const st = E.story();
    D.info(doc(), E.sel ? E.sel.f.p : st.paras && st.paras[0]);
    const paras = st.paras || [];
    if (!paras.length) return;
    const last = paras[paras.length - 1];
    E.setSel(D.pos(paras[0], 0), D.pos(last, D.plen(last)), { noScroll: true });
  };
  E.selectPara = function (p) {
    const nx = D.nextPara(doc(), p);
    if (nx && O.containerOf(nx) === O.containerOf(p)) E.setSel(D.pos(p, 0), D.pos(nx, 0));
    else E.setSel(D.pos(p, 0), D.pos(p, D.plen(p)));
  };
  E.selectWord = function () { if (!E.sel) return; const w = O.wordExtent(E.sel.f); E.setSel(w[0], w[1]); };
  E.focus = function () {
    if (!LY.root) return;
    const host = LY.root;
    if (document.activeElement !== host) { try { host.focus({ preventScroll: true }); } catch (e) { host.focus(); } }
    E.restoreDom(true);
  };
  E.refocus = () => setTimeout(() => E.focus(), 0);
  E.textSave = () => {};

  /* ---------- table cell selection highlighting ---------- */
  E.cellSel = null;
  E.markCells = function () {
    if (!LY.root) return;
    for (const td of LY.root.querySelectorAll('td.csel')) td.classList.remove('csel');
    E.cellSel = null; E.rowSel = null;
    if (!E.sel || E.collapsed()) return;
    const d = doc();
    const ca = D.cellOf(d, E.sel.a.p), cb = D.cellOf(d, E.sel.f.p);
    if (!ca && !cb) return;
    /* selection that spans cells of one table selects whole cells */
    if (ca && cb && ca.tbl === cb.tbl && ca.cell !== cb.cell) {
      const map = L.R.tblMap(ca.tbl);
      const pos = (ci) => { const m = map[ci.ri].find((x) => x.cell === ci.cell); return { r: ci.ri, c0: m.c0, c1: m.c0 + m.span - 1 }; };
      const pa = pos(ca), pb = pos(cb);
      const r0 = Math.min(pa.r, pb.r), r1 = Math.max(pa.r, pb.r), c0 = Math.min(pa.c0, pb.c0), c1 = Math.max(pa.c1, pb.c1);
      const cells = [];
      for (let r = r0; r <= r1; r++) for (const m of map[r]) if (!m.hidden && m.c0 + m.span - 1 >= c0 && m.c0 <= c1) cells.push(m.cell);
      E.cellSel = { tbl: ca.tbl, r0, r1, c0, c1, cells };
      for (const c of cells) for (const td of LY.root.querySelectorAll(`td[data-cid="${c.id}"]`)) td.classList.add('csel');
    } else {
      const [a, b] = E.range();
      /* from inside a table out of it (into the text above or below): whole rows, from the row the selection starts
         in to the table's end (or from its start), as Word does */
      const ia = D.info(d, a.p), ib = D.info(d, b.p);
      const ta = ia && ia.parents[0] && ia.parents[0].tbl ? ia.parents[0] : null, tb = ib && ib.parents[0] && ib.parents[0].tbl ? ib.parents[0] : null;
      const rs = ta && ib && !ib.parents.length && ia.story === ib.story ? { tbl: ta.tbl, cont: ia.story, r0: ta.tbl.rows.indexOf(ta.row), r1: ta.tbl.rows.length - 1, down: true }
        : tb && ia && !ia.parents.length && ia.story === ib.story ? { tbl: tb.tbl, cont: ib.story, r0: 0, r1: tb.tbl.rows.indexOf(tb.row), down: false } : null;
      if (rs) {
        E.rowSel = rs;
        for (let r = rs.r0; r <= rs.r1; r++) for (const c of rs.tbl.rows[r].cells) for (const td of LY.root.querySelectorAll(`td[data-cid="${c.id}"]`)) td.classList.add('csel');
        return;
      }
      /* whole cells covered when selection runs from inside a table to outside */
      for (const p of D.parasBetween(d, a.p, b.p)) { const ci = D.cellOf(d, p); if (ci && (p !== a.p || a.o === 0)) for (const td of LY.root.querySelectorAll(`td[data-cid="${ci.cell.id}"]`)) td.classList.add('csel'); }
    }
  };

  /* ---------- event wiring ---------- */
  E.init = function () {
    const root = LY.root;
    try { document.execCommand('enableObjectResizing', false, false); document.execCommand('enableInlineTableEditing', false, false); } catch (e) { /* firefox only */ }
    document.addEventListener('selectionchange', E.onSelectionChange);
    root.addEventListener('beforeinput', onBeforeInput);
    root.addEventListener('keydown', onKeyDown);
    root.addEventListener('compositionstart', onCompStart);
    root.addEventListener('compositionend', onCompEnd);
    root.addEventListener('paste', onPaste);
    root.addEventListener('copy', (e) => onCopy(e, false));
    root.addEventListener('cut', (e) => onCopy(e, true));
    root.addEventListener('drop', onDrop);
    root.addEventListener('dragstart', onDragStart);
    root.addEventListener('dragover', (e) => { if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) e.preventDefault(); });
    root.addEventListener('mousedown', onMouseDown);
    root.addEventListener('click', onClick);
    root.addEventListener('dblclick', onDblClick);
    root.addEventListener('contextmenu', (e) => { if (L.app && L.app.contextMenu) L.app.contextMenu(e); });
    root.addEventListener('focus', () => { root.classList.add('focused'); });
    root.addEventListener('blur', () => { root.classList.remove('focused'); });
    L.bus.on('doc-changed', (info) => {
      if (E.composing) { E.needLayout = info; return; }
      LY.update(info);
      if (info && info.history) return; /* history handler restores selection */
      E.restoreDom();
      E.selChanged();
    });
    L.bus.on('history', (h) => {
      if (E.applySerialized(h.sel)) E.restoreDom();
      E.selChanged();
    });
  };

  function onBeforeInput(e) {
    const t = e.inputType;
    if (E.composing || t === 'insertCompositionText' || t === 'deleteCompositionText' || t === 'insertFromComposition') return;
    e.preventDefault();
    if (!E.objSel) E.syncSel();
    if (!E.sel) { const m = E.readDom(); if (m) E.sel = m; else return; }
    switch (t) {
      case 'insertText': case 'insertReplacementText': {
        let text = e.data;
        if (text == null && e.dataTransfer) text = e.dataTransfer.getData('text/plain');
        if (t === 'insertReplacementText' && e.getTargetRanges) {
          const rs = e.getTargetRanges();
          if (rs.length) {
            const a = LY.domToPos(rs[0].startContainer, rs[0].startOffset), b = LY.domToPos(rs[0].endContainer, rs[0].endOffset);
            if (a && b) { E.sel = { a, f: b }; }
          }
        }
        if (text) {
          if (text.length === 1 && L.autocorrect && L.autocorrect.smartChar) text = L.autocorrect.smartChar(text);
          E.typeText(text);
        }
        break;
      }
      case 'insertParagraph': E.enter(); break;
      case 'insertLineBreak': E.enter('line'); break;
      case 'deleteContentBackward': E.backspace(false); break;
      case 'deleteContentForward': E.del(false); break;
      case 'deleteWordBackward': E.backspace(true); break;
      case 'deleteWordForward': E.del(true); break;
      case 'deleteSoftLineBackward': case 'deleteHardLineBackward': case 'deleteSoftLineForward': case 'deleteHardLineForward': case 'deleteContent': case 'deleteByCut':
        if (!E.collapsed()) E.edit(L.N('Delete'), () => E.deleteSelection());
        break;
      case 'historyUndo': L.ui.exec('undo'); break;
      case 'historyRedo': L.ui.exec('redo'); break;
      case 'formatBold': L.ui.exec('bold'); break;
      case 'formatItalic': L.ui.exec('italic'); break;
      case 'formatUnderline': L.ui.exec('underline'); break;
      case 'insertFromPaste': case 'insertFromDrop': break; /* handled by paste/drop events */
      default: break;
    }
  }
  /* --- IME composition: let the browser draw the composition, then commit it to the model --- */
  let compStart = null;
  function onCompStart() {
    if (!E.sel) { const m = E.readDom(); if (m) E.sel = m; }
    if (E.sel && !E.collapsed()) {
      /* replace the selection first so the composition starts at a clean caret */
      E.edit(L.N('Typing'), () => E.deleteSelection());
    }
    E.composing = true;
    compStart = E.sel ? { p: E.sel.f.p, o: E.sel.f.o } : null;
  }
  function onCompEnd(e) {
    E.composing = false;
    const text = e.data || '';
    if (!compStart) return;
    E.sel = { a: compStart, f: compStart };
    /* discard the browser's DOM changes by re-rendering from the model, then insert */
    if (text) E.typeText(text);
    else { LY.update({ paras: new Set([compStart.p]) }); E.restoreDom(); }
    if (E.needLayout) { E.needLayout = null; }
    compStart = null;
  }

  /* ---------- keyboard ---------- */
  function onKeyDown(e) {
    if (E.composing || e.isComposing) return;
    if (L.ui.menuOpen() || L.ui.dialogOpen()) return;
    const k = e.key;
    const ctrl = e.ctrlKey || e.metaKey, shift = e.shiftKey, alt = e.altKey;
    if (!E.objSel && !/^(Shift|Control|Alt|Meta)$/.test(k)) E.syncSel();
    if (L.app && L.app.shortcut && L.app.shortcut(e)) { e.preventDefault(); return; }
    /* a selected floating object: the selection is its anchor, so Enter and Tab would type over it and delete it
       (Delete and Backspace below do delete it, as they should) */
    if (E.objSel && E.objSel.float && (k === 'Enter' || k === 'Tab')) { e.preventDefault(); return; }
    if (k === 'Enter' && !alt) {
      e.preventDefault();
      if (ctrl && shift) E.enter('column');
      else if (ctrl) E.enter('page');
      else if (shift) E.enter('line');
      else E.enter();
      return;
    }
    if (k === 'Backspace') { e.preventDefault(); if (alt) L.ui.exec('undo'); else E.backspace(ctrl); return; }
    if (k === 'Delete') { e.preventDefault(); if (shift && !ctrl) L.ui.exec('cut'); else E.del(ctrl); return; }
    if (k === 'Tab' && !alt) {
      e.preventDefault();
      if (ctrl) E.insertItem(D.item('tab'), L.N('Typing')); else E.tab(shift);
      return;
    }
    if (k === 'Insert' && !ctrl && !shift) { e.preventDefault(); E.overtype = !E.overtype; L.bus.emit('sel', E.sel); return; }
    if (k === ' ' && ctrl && shift) { e.preventDefault(); E.typeText(' '); return; }
    if (k === 'Escape') { E.clearObjSel(); return; }
    /* paragraph navigation */
    if (ctrl && !alt && (k === 'ArrowUp' || k === 'ArrowDown')) {
      e.preventDefault();
      const f = E.sel.f;
      let target;
      if (k === 'ArrowUp') target = f.o > 0 ? D.pos(f.p, 0) : (D.prevPara(doc(), f.p) ? D.pos(D.prevPara(doc(), f.p), 0) : f);
      else { const nx = D.nextPara(doc(), f.p); target = nx ? D.pos(nx, 0) : D.pos(f.p, D.plen(f.p)); }
      if (shift) E.setSel(E.sel.a, target); else E.setSel(target);
      return;
    }
    /* word movement the Word way: Ctrl+→ goes to the start of the next word, so Ctrl+Shift+→ takes the trailing space too */
    if (ctrl && !alt && (k === 'ArrowLeft' || k === 'ArrowRight') && E.sel && !E.objSel) {
      e.preventDefault();
      const isW = (c) => /[\p{L}\p{N}_'’]/u.test(c);
      const isS = (c) => /[    ]/.test(c);
      let p = E.sel.f.p, o = E.sel.f.o;
      const t = D.ptext(p);
      if (k === 'ArrowRight') {
        if (o >= t.length) { const nx = D.nextPara(doc(), p); if (nx) { p = nx; o = 0; } }
        else { if (isW(t[o])) while (o < t.length && isW(t[o])) o++; else if (!isS(t[o])) o++; while (o < t.length && isS(t[o])) o++; }
      } else if (o === 0) { const pv = D.prevPara(doc(), p); if (pv) { p = pv; o = D.plen(pv); } }
      else { while (o > 0 && isS(t[o - 1])) o--; if (o > 0 && isW(t[o - 1])) { while (o > 0 && isW(t[o - 1])) o--; } else if (o > 0) o--; }
      const target = D.pos(p, o);
      if (shift) E.setSel(E.sel.a, target); else E.setSel(target);
      return;
    }
    /* document start / end (the browser would put the caret on the page container) */
    if (ctrl && !alt && (k === 'Home' || k === 'End') && E.sel) {
      e.preventDefault();
      const st = E.story();
      const p = k === 'Home' ? D.firstPara(st.blocks) : D.lastPara(st.blocks);
      if (!p) return;
      const target = D.pos(p, k === 'Home' ? 0 : D.plen(p));
      if (shift) E.setSel(E.sel.a, target); else E.setSel(target);
      E.pending = null;
      return;
    }
    if (alt && shift && (k === 'ArrowUp' || k === 'ArrowDown')) { e.preventDefault(); L.ui.exec(k === 'ArrowUp' ? 'moveUp' : 'moveDown'); return; }
    if (alt && shift && (k === 'ArrowLeft' || k === 'ArrowRight')) { e.preventDefault(); L.ui.exec(k === 'ArrowLeft' ? 'promote' : 'demote'); return; }
    /* objects: arrow keys nudge floating objects */
    if (E.objSel && E.objSel.float && /^Arrow/.test(k) && L.drawing) { e.preventDefault(); L.drawing.nudge(k, ctrl ? 1 : 7.2); return; }
    /* typing on a selected floating shape goes into its text (Word); over another floating object it does nothing */
    if (k.length === 1 && !ctrl && !alt && E.objSel && E.objSel.float) { e.preventDefault(); const c = L.ui.cmds.shapeText; if (c && c.enabled()) L.drawing.editText(k); return; }
  }

  /* ---------- clipboard ---------- */
  E.clip = null; /* internal clipboard: {blocks, text} */
  function onCopy(e, cut) {
    if (!E.sel || E.collapsed()) { if (!E.objSel) return; }
    e.preventDefault();
    const data = E.copyData();
    if (!data) return;
    try {
      e.clipboardData.setData('text/plain', data.text);
      e.clipboardData.setData('text/html', data.html);
      e.clipboardData.setData('application/x-quire', JSON.stringify({ blocks: data.blocks, media: data.media, preserved: data.preserved }));
    } catch (err) { /* ignore */ }
    E.clip = data;
    if (cut) E.edit(L.N('Cut'), () => E.deleteSelection({ wholeTable: true }));
  }
  E.copyData = function () {
    if (!E.sel) return null;
    const [a, b] = E.range();
    if (D.eqPos(a, b)) return null;
    const blocks = O.copyRange(a, b);
    const text = O.textRange(a, b);
    const html = L.htmlio ? L.htmlio.blocksToHTML(blocks, { clipboard: true }) : L.esc(text);
    const media = {};
    return { blocks, text, html, media, preserved: L.preserve.clipboard(blocks) };
  };
  async function onPaste(e) {
    e.preventDefault();
    if (!E.sel) return;
    const cd = e.clipboardData;
    let q = null;
    try { q = cd.getData('application/x-quire'); } catch (err) { /* ignore */ }
    if (q) { try { const o = JSON.parse(q); const blocks = o.preserved ? await L.preserve.hydrateClipboard(L.opc.remapSources(o.blocks, L.opc.import(o.preserved))) : o.blocks; E.pasteBlocks(blocks); return; } catch (err) { /* fall through */ } }
    const files = Array.from(cd.files || []).filter((f) => /^image\//.test(f.type));
    if (files.length && L.app && L.app.insertPictureFiles) { L.app.insertPictureFiles(files); return; }
    const html = cd.getData('text/html');
    const text = cd.getData('text/plain');
    if (html && L.htmlio) {
      const blocks = await L.htmlio.htmlToBlocks(html);
      if (blocks && blocks.length) { E.pasteBlocks(blocks); return; }
    }
    if (text) E.pasteText(text);
  }
  E.pasteBlocks = function (blocks, opts) {
    if (!blocks || !blocks.length) return;
    E.edit(L.N('Paste'), () => {
      const pos = E.deleteSelection();
      return O.insertBlocks(pos, blocks, opts);
    });
    E.lastPaste = Date.now();
  };
  E.pasteText = function (text) {
    E.edit(L.N('Paste'), () => O.insertPlain(E.deleteSelection(), text.replace(/\r\n?/g, '\n').replace(/\n$/, '')));
  };
  /* internal drag & drop of text: dragData is the text dragged from this page, until the drag ends (dropped
     anywhere, or cancelled); a drop carries its mark, so a drop from elsewhere never moves it */
  let dragData = null;
  const DRAG_MARK = 'application/x-quire-drag';
  const dndOff = () => !!(L.app && L.app.opts && L.app.opts.dragDrop === false);
  function onDragStart(e) {
    dragData = null;
    if (!E.sel || E.collapsed()) return;
    if (dndOff()) { e.preventDefault(); return; }
    dragData = { range: E.range(), data: E.copyData() };
    if (!dragData.data) { dragData = null; return; }
    try { e.dataTransfer.setData('text/plain', dragData.data.text); e.dataTransfer.setData('text/html', dragData.data.html); e.dataTransfer.setData(DRAG_MARK, '1'); e.dataTransfer.effectAllowed = 'copyMove'; } catch (err) { /* ignore */ }
  }
  document.addEventListener('dragend', () => { dragData = null; }, true);
  /** move (Ctrl: copy) the text a…b to target, one undo step; the text moved is selected after, as in Word */
  function moveRange(a, b, target, copy) {
    const d = doc();
    if (D.cmp(d, target, a) > 0 && D.cmp(d, target, b) < 0) return false;
    const blocks = O.copyRange(a, b);
    const one = D.info(d, a.p).cont === D.info(d, b.p).cont;
    E.clearObjSel();
    return E.edit(copy ? 'Drag and Drop' : L.N('Move'), () => {
      let at = target;
      if (!copy) {
        if (D.cmp(d, target, b) >= 0 && !one) {
          /* across a table: insert first, so the positions stay valid */
          const end = O.insertBlocks(target, blocks);
          O.deleteRange(a, b);
          return end;
        }
        /* a drop later in the source's last paragraph is counted from that paragraph's end, which the removal doesn't
           touch: the removal may shorten the paragraph and join it to the source's first, or (tracked) leave it whole */
        const tail = D.cmp(d, target, b) >= 0 && target.p === b.p ? D.plen(b.p) - target.o : null;
        O.deleteRange(a, b);
        if (tail != null) { const p = D.info(d, b.p) ? b.p : a.p; at = D.pos(p, D.plen(p) - tail); }
      }
      return { a: at, f: O.insertBlocks(at, blocks) };
    });
  }
  async function onDrop(e) {
    const files = Array.from((e.dataTransfer && e.dataTransfer.files) || []);
    e.preventDefault();
    let target = null;
    const cp = document.caretPositionFromPoint ? document.caretPositionFromPoint(e.clientX, e.clientY) : null;
    if (cp) target = LY.domToPos(cp.offsetNode, cp.offset);
    else if (document.caretRangeFromPoint) { const r = document.caretRangeFromPoint(e.clientX, e.clientY); if (r) target = LY.domToPos(r.startContainer, r.startOffset); }
    if (files.length) {
      if (target) E.setSel(target);
      if (files.some((f) => /\.(docx|docm|dotx|txt|html?|rtf|doc)$/i.test(f.name))) { if (L.app) L.app.openFiles(files); return; }
      if (L.app) L.app.insertPictureFiles(files.filter((f) => /^image\//.test(f.type) || /\.(emf|wmf)$/i.test(f.name)));
      return;
    }
    if (!target) return;
    if (dragData && !Array.from(e.dataTransfer.types || []).includes(DRAG_MARK)) dragData = null;
    if (dragData) {
      const [a, b] = dragData.range;
      dragData = null;
      moveRange(a, b, target, e.ctrlKey || e.altKey);
      return;
    }
    const html = e.dataTransfer.getData('text/html'), text = e.dataTransfer.getData('text/plain');
    E.setSel(target);
    if (html && L.htmlio) { const bl = await L.htmlio.htmlToBlocks(html); if (bl.length) { E.pasteBlocks(bl); return; } }
    if (text) E.pasteText(text);
  }

  /* ---------- mouse ---------- */
  function onMouseDown(e) {
    if (e.button !== 0 && e.button !== 2) return;
    const t = e.target;
    /* the part of the page not being edited (the body while a header is, a header while the body is) takes no
       caret; a double-click there switches (onDblClick) */
    if (offLimits(t)) { e.preventDefault(); return; }
    /* floating objects & inline images; a right-click selects one too (a selected one stays so), so the menu is its own */
    const obj = t.closest && t.closest('.img, .shp, .flt');
    if (obj && L.drawing && L.drawing.onObjectDown && L.drawing.onObjectDown(e, obj)) return;
    if (obj && obj.classList.contains('img') && obj.closest('.p')) {
      e.preventDefault();
      if (e.button === 2 && E.objSel && E.objSel.el === obj) return;
      const pos = LY.domToPos(obj.parentNode, Array.prototype.indexOf.call(obj.parentNode.childNodes, obj));
      if (pos) { selectInlineObject(pos, obj); if (e.button === 0) startObjectMove(e, pos); }
      return;
    }
    if (e.button !== 0) return;
    E.clearObjSel();
    /* Drag-and-drop text editing off (Tools ▸ Options): a press inside the selection starts a new one there */
    if (dndOff() && !e.shiftKey && E.sel && !E.collapsed() && document.caretRangeFromPoint) {
      const r = document.caretRangeFromPoint(e.clientX, e.clientY);
      if (r && LY.root.contains(r.startContainer)) window.getSelection().setBaseAndExtent(r.startContainer, r.startOffset, r.startContainer, r.startOffset);
    }
    /* selection bar (left margin) */
    if (LY.view === 'print' || LY.view === 'normal') {
      const pg = t.closest && t.closest('.pg, .flow');
      if (pg && (t === pg || t.classList.contains('crop')) && !LY.hfEdit) {
        const body = pg.querySelector('.pg-body') || pg;
        const br = body.getBoundingClientRect();
        if (e.clientX < br.left - 2) {
          e.preventDefault();
          LY.root.focus({ preventScroll: true });
          const line = lineAt(e.clientY);
          if (!line) return;
          if (e.ctrlKey || e.detail >= 3) { E.selectAll(); return; }
          if (e.detail === 2) { E.selectPara(line[0].p); return; }
          selectBar(e, line);
          return;
        }
      }
    }
    /* Ctrl+click selects a sentence */
    if (e.ctrlKey && !e.shiftKey && !(t.closest && t.closest('[data-link]'))) {
      setTimeout(() => { E.syncSel(); if (E.sel) { const s = O.sentenceExtent(E.sel.f); E.setSel(s[0], s[1]); } }, 0);
    }
  }
  function posAtPoint(x, y) {
    let node = null, off = 0;
    if (document.caretPositionFromPoint) { const cp = document.caretPositionFromPoint(x, y); if (cp) { node = cp.offsetNode; off = cp.offset; } }
    else if (document.caretRangeFromPoint) { const r = document.caretRangeFromPoint(x, y); if (r) { node = r.startContainer; off = r.startOffset; } }
    return node ? LY.domToPos(node, off) : null;
  }
  E.posAtPoint = posAtPoint;
  /** [start, end] of the body's line at height y (the selection bar beside it), as the page lays it out */
  function lineAt(y) {
    const pg = LY.pages.find((x) => { const r = x.el.getBoundingClientRect(); return y >= r.top && y < r.bottom; }) || null;
    const body = pg ? pg.el.querySelector('.pg-body') : LY.root.querySelector('.flow');
    if (!body) return null;
    const br = body.getBoundingClientRect();
    const r = document.caretRangeFromPoint && document.caretRangeFromPoint(br.left + 2, L.clamp(y, br.top + 1, br.bottom - 1));
    if (!r || !body.contains(r.startContainer)) return null;
    const s = window.getSelection();
    settingDom++;
    try {
      s.setBaseAndExtent(r.startContainer, r.startOffset, r.startContainer, r.startOffset);
      s.modify('move', 'backward', 'lineboundary');
      const a = LY.domToPos(s.focusNode, s.focusOffset);
      s.modify('move', 'forward', 'lineboundary');
      const b = LY.domToPos(s.focusNode, s.focusOffset);
      return a && b && a.p === b.p ? [a, b] : null;
    } catch (err) { return null; } finally { setTimeout(() => { settingDom = Math.max(0, settingDom - 1); }, 0); }
  }
  /** a press in the selection bar: whole lines, from the line pressed (with Shift, from the selection's start) to
      the line at the pointer as it is dragged */
  function selectBar(e, line) {
    const from = e.shiftKey && E.sel ? E.sel.a : null;
    const put = (cur) => {
      if (from) E.setSel(from, D.cmp(doc(), cur[0], from) < 0 ? cur[0] : cur[1], { noScroll: true });
      else if (D.cmp(doc(), cur[0], line[0]) < 0) E.setSel(line[1], cur[0], { noScroll: true });
      else E.setSel(line[0], cur[1], { noScroll: true });
    };
    put(line);
    L.ui.drag(e, { threshold: 0, move(ev) { const cur = lineAt(ev.clientY); if (cur) put(cur); }, up() { setTimeout(selDone, 0); } });
  }
  /** a hyperlink: a bookmark in this document, or (L.ui.linkDialog) its address to copy, never opened from here */
  E.followLink = function (url) {
    if (url.startsWith('#')) { if (L.app && L.app.gotoBookmark) L.app.gotoBookmark(url.slice(1)); }
    else L.ui.linkDialog(url);
  };
  function onClick(e) {
    const t = e.target;
    const control = t.closest && t.closest('[data-control]');
    if (control && !E.readOnly) {
      const record = L.preserve.controls(D.doc).records.get(+control.dataset.control);
      if (record) { e.preventDefault(); L.preserve.controlDialog(record); return; }
    }
    const ln = t.closest && t.closest('[data-link]');
    if (ln && (e.ctrlKey || e.metaKey || LY.view === 'reading' || E.readOnly)) {
      e.preventDefault();
      E.followLink(ln.dataset.link);
      return;
    }
    const nr = t.closest && t.closest('.nref');
    if (nr && L.app && L.app.gotoNote && e.detail === 2) L.app.gotoNote(nr.dataset.note);
    /* a check box (task lists, form fields) ticks on click */
    const cb = t.closest && t.closest('.ffcb');
    if (cb && !E.readOnly) {
      const f = D.fields(D.doc).find((x) => x.it.fid === cb.dataset.fid && x.it.ff);
      if (f) E.edit('Check Box', () => { D.touch(f.p); f.it.ff = Object.assign({}, f.it.ff, { checked: !f.it.ff.checked }); return E.sel; });
    }
  }
  function onDblClick(e) {
    const t = e.target;
    const hf = t.closest && t.closest('.pg-hdr, .pg-ftr');
    if (hf && !LY.hfEdit && L.app && L.app.editHeaderFooter) { e.preventDefault(); L.app.editHeaderFooter(hf.dataset.kind, +hf.closest('.pg').dataset.pg, e); return; }
    if (LY.hfEdit && t.closest && t.closest('.pg-body') && L.app && L.app.closeHeaderFooter) { e.preventDefault(); L.app.closeHeaderFooter(false, e); return; }
    /* the top or bottom margin of a page (an empty header or footer has no height): open it */
    const pg = t.closest && t.closest('.pg');
    if (pg && !LY.hfEdit && LY.view === 'print' && (t === pg || t.classList.contains('crop')) && L.app && L.app.editHeaderFooter) {
      const b = (pg.querySelector('.pg-body') || pg).getBoundingClientRect();
      if (e.clientY < b.top || e.clientY > b.bottom) { e.preventDefault(); L.app.editHeaderFooter(e.clientY < b.top ? 'hdr' : 'ftr', +pg.dataset.pg, e); return; }
    }
    const nr = t.closest && t.closest('.nref');
    if (nr && L.app && L.app.gotoNote) { e.preventDefault(); L.app.gotoNote(nr.dataset.note); return; }
    const img = t.closest && t.closest('.img, .shp');
    if (img && L.app && L.app.formatObject) { e.preventDefault(); L.app.formatObject(); return; }
  }
  /** an inline picture (or shape, chart) pressed and dragged: it moves into the text where it is let go (Ctrl: a
      copy), a caret showing where; Esc puts it down where it was */
  function startObjectMove(e, pos) {
    let mark = null;
    const caretAt = (ev) => { const r = document.caretRangeFromPoint && document.caretRangeFromPoint(ev.clientX, ev.clientY); return r && LY.root.contains(r.startContainer) && !(r.startContainer.nodeType === 1 && r.startContainer.closest('.pg-hdr, .pg-ftr')) ? r : null; };
    const drop = () => { if (mark) { mark.remove(); mark = null; } };
    L.ui.drag(e, {
      move(ev) {
        const r = caretAt(ev), box = r && (r.getClientRects()[0] || r.getBoundingClientRect());
        if (!mark) { mark = document.createElement('div'); mark.className = 'drop-caret'; LY.scroller.appendChild(mark); }
        if (!box || !box.height) { mark.style.display = 'none'; return; }
        const sr = LY.scroller.getBoundingClientRect();
        Object.assign(mark.style, { display: '', left: box.left - sr.left + LY.scroller.scrollLeft + 'px', top: box.top - sr.top + LY.scroller.scrollTop + 'px', height: box.height + 'px' });
      },
      up(ev, d) {
        drop();
        if (!d.moved) return;
        const r = caretAt(ev), target = r && LY.domToPos(r.startContainer, r.startOffset);
        if (!target || (target.p === pos.p && (target.o === pos.o || target.o === pos.o + 1))) return;
        if (moveRange(D.pos(pos.p, pos.o), D.pos(pos.p, pos.o + 1), target, ev.ctrlKey)) {
          const at = E.sel.a;
          setTimeout(() => { if (L.app && L.app.selectItemAt) L.app.selectItemAt(at.p, at.o); }, 0);
        }
      },
      cancel: drop,
    });
  }
  E.startObjectMove = startObjectMove;
  /* inline object selection with resize handles */
  let handles = null;
  function selectInlineObject(pos, elObj) {
    E.selectObject({ p: pos.p, o: pos.o, it: D.itemAfter(pos.p, pos.o), el: elObj, float: false });
  }
  function hideHandles() { if (handles) { handles.remove(); handles = null; } }
  E.hideHandles = hideHandles;
  function showHandles(target) {
    hideHandles();
    const sc = LY.scroller;
    const box = document.createElement('div');
    box.className = 'obj-handles';
    /* a right-click on a handle: the object's menu, as anywhere else on it */
    box.addEventListener('contextmenu', (ev) => { ev.preventDefault(); ev.stopPropagation(); if (L.app && L.app.contextMenu) L.app.contextMenu(ev); });
    const place = () => {
      const r = target.getBoundingClientRect(), sr = sc.getBoundingClientRect();
      box.style.left = r.left - sr.left + sc.scrollLeft + 'px';
      box.style.top = r.top - sr.top + sc.scrollTop + 'px';
      box.style.width = r.width + 'px';
      box.style.height = r.height + 'px';
    };
    place();
    for (const hname of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
      const hd = document.createElement('div');
      hd.className = 'oh oh-' + hname;
      hd.dataset.h = hname;
      hd.addEventListener('pointerdown', (ev) => startResize(ev, hname, target, box));
      box.appendChild(hd);
    }
    sc.appendChild(box);
    handles = box;
    box._place = place;
  }
  E.placeHandles = () => { if (handles && handles._place) handles._place(); };
  function startResize(ev, hname, target, box) {
    ev.preventDefault();
    ev.stopPropagation();
    const os = E.objSel;
    if (ev.button !== 0 || !os || !os.it) return;
    const z = LY.zoom;
    const it = os.it;
    const w0 = it.w || 72, h0 = it.h || 72;
    const corner = hname.length === 2;
    const ghost = document.createElement('div');
    ghost.className = 'resize-ghost';
    let nw = w0, nh = h0;
    L.ui.drag(ev, {
      move(e2, d) {
        if (!ghost.parentNode) box.appendChild(ghost);
        const dx = d.dx / z / D.PX, dy = d.dy / z / D.PX;
        nw = w0 + (hname.includes('e') ? dx : hname.includes('w') ? -dx : 0);
        nh = h0 + (hname.includes('s') ? dy : hname.includes('n') ? -dy : 0);
        if (corner && !e2.shiftKey) { const k = Math.max(nw / w0, nh / h0); nw = w0 * k; nh = h0 * k; }
        nw = Math.max(4, nw); nh = Math.max(4, nh);
        ghost.style.width = nw * D.PX * z + 'px';
        ghost.style.height = nh * D.PX * z + 'px';
        if (hname.includes('w')) ghost.style.left = (w0 - nw) * D.PX * z + 'px';
        if (hname.includes('n')) ghost.style.top = (h0 - nh) * D.PX * z + 'px';
      },
      up(e2, d) {
        ghost.remove();
        if (!d.moved || (Math.abs(nw - w0) < 0.5 && Math.abs(nh - h0) < 0.5)) return;
        const pos = { p: os.p, o: os.o };
        E.edit('Resize Picture', () => {
          D.touch(pos.p);
          const item = D.itemAfter(pos.p, pos.o);
          if (item) { item.w = L.round(nw, 2); item.h = L.round(nh, 2); }
          return { a: D.pos(pos.p, pos.o), f: D.pos(pos.p, pos.o + 1) };
        });
        setTimeout(() => {
          const f = LY.fragFor(pos.p, pos.o);
          const elx = f && f.querySelector(`[data-o="${pos.o}"]`);
          if (elx) selectInlineObject(pos, elx);
        }, 0);
      },
      cancel: () => ghost.remove(),
      capture: LY.scroller,   /* the handles are drawn again with the selection */
    });
  }
  E.selectInlineObject = selectInlineObject;
})();
