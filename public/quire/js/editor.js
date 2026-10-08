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
    if (!opts || !opts.keepPending) E.pending = null;
    if (!opts || !opts.noDom) E.restoreDom(opts && opts.noScroll);
    E.selChanged();
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
    const s = window.getSelection();
    const a = LY.posToDom(E.sel.a), f = LY.posToDom(E.sel.f);
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
  E.onSelectionChange = function () {
    if (E.composing) return;
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
    E.pending = null; D.breakMerge();
    E.sel = m;
    E.objSel = null;
    E.selChanged();
    E.markCells();
  };
  /** pull a pending DOM selection into the model before acting on a key (selectionchange is asynchronous) */
  E.syncSel = function () {
    if (E.composing || !LY.root || domIsOurs()) return;
    const s = window.getSelection();
    if (!s.rangeCount || !isEditableNode(s.focusNode)) return;
    const m = E.readDom();
    if (!m) { snapCaret(); return; }
    if (E.sel && D.eqPos(m.a, E.sel.a) && D.eqPos(m.f, E.sel.f)) return;
    E.pending = null; D.breakMerge();
    E.sel = m; E.objSel = null; E.selChanged();
  };
  /* where the pointer is while a button is held: a drag that leaves the text is placed by the pointer */
  let pointer = null, textDrag = false, snapQueued = false;
  document.addEventListener('pointerdown', (e) => { const t = e.target; textDrag = e.button === 0 && !!(t.closest && t.closest('.p') && LY.root && LY.root.contains(t)); }, true);
  document.addEventListener('pointermove', (e) => {
    pointer = e.buttons & 1 ? { x: e.clientX, y: e.clientY } : null;
    /* outside the page (grey area, other panes) the browser stops extending the selection: follow the pointer */
    if (pointer && textDrag && !overText(pointer) && !snapQueued) { snapQueued = true; requestAnimationFrame(() => { snapQueued = false; if (pointer && textDrag) snapCaret(); }); }
  }, true);
  document.addEventListener('pointerup', () => { pointer = null; textDrag = false; }, true);
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
    let frag = null, bd = Infinity, below = false, level = false, box = null;
    for (const f of LY.root.querySelectorAll(story ? '.p' : '.pg-body .p, .pg-hdr .p, .pg-ftr .p')) {
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

  E.scrollToCaret = function () {
    const sc = LY.scroller;
    if (!sc || !E.sel) return;
    const s = window.getSelection();
    if (!s.rangeCount) return;
    const rng = document.createRange();
    try { rng.setStart(s.focusNode, s.focusOffset); rng.collapse(true); } catch (e) { return; }
    let r = rng.getClientRects()[0] || rng.getBoundingClientRect();
    if (!r || (!r.height && !r.top)) {
      const n = s.focusNode && (s.focusNode.nodeType === 3 ? s.focusNode.parentNode : s.focusNode);
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
  /** delete the current selection (if any) and return the caret position */
  E.deleteSelection = function () {
    if (E.collapsed()) return E.sel.f;
    const [a, b] = E.range();
    if (E.cellSel) { const r = clearCells(); if (r) return r; }
    const pos = O.deleteRange(a, b);
    O.repairFields(D.storyOf(doc(), pos.p));
    return pos;
  };
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
    E.edit('Typing', () => {
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
      E.edit('Numbering', () => {
        D.touch(p);
        if (lvl > 0) p.pPr.num = { id: pp.num.id, lvl: lvl - 1 };
        else { p.pPr.num = { id: '0', lvl: 0 }; delete p.pPr.ind; if (D.styleProps(d, p.pPr.style || 'Normal').pPr.num) p.pPr.style = undefined; }
        return D.pos(p, 0);
      });
      return;
    }
    E.edit('Typing', () => {
      const pos = E.deleteSelection();
      const np = O.splitPara(pos);
      /* AutoFormat: "---" + Enter → border line */
      return np;
    });
    if (L.autocorrect) L.autocorrect.afterEnter();
  };
  E.insertItem = function (item, label) {
    E.edit(label || 'Insert', () => O.insertItem(E.deleteSelection(), item));
  };
  E.insertTextAt = (text, label) => E.edit(label || 'Insert', () => O.insertPlain(E.deleteSelection(), text, E.pending ? D.mergeRPr(O.inheritRPr(E.sel.f), E.pending) : null));
  E.backspace = function (word) {
    if (!E.sel) return;
    const d = doc();
    if (!E.collapsed()) { E.edit('Delete', () => E.deleteSelection()); return; }
    const pos = E.sel.f;
    const p = pos.p;
    if (pos.o === 0) {
      /* list number / indent removal first */
      const lb = LY.labels && LY.labels.get(p.id);
      const pp = D.pProps(d, p);
      if (lb) {
        E.edit('Numbering', () => {
          D.touch(p);
          const ind = pp.ind || {};
          p.pPr.num = { id: '0', lvl: 0 };
          p.pPr.ind = { l: ind.l || 0, fl: 0, r: ind.r || 0 };
          return D.pos(p, 0);
        });
        return;
      }
      if (p.pPr.ind && (p.pPr.ind.fl > 0 || p.pPr.ind.l > 0) && L.store.get('acIndentBackspace', true)) {
        E.edit('Indent', () => { D.touch(p); if (p.pPr.ind.fl > 0) p.pPr.ind.fl = 0; else p.pPr.ind.l = Math.max(0, (p.pPr.ind.l || 0) - 36); return D.pos(p, 0); });
        return;
      }
      const info = D.info(d, p);
      const idx = info.cont.blocks.indexOf(p);
      const prevB = info.cont.blocks[idx - 1];
      if (!prevB) return; /* start of story / cell */
      if (prevB.t === 'tbl') {
        /* empty paragraph after a table is removed; otherwise move into the table */
        if (!D.plen(p) && info.cont.blocks[idx + 1]) { E.edit('Delete', () => { const r = O.removeBlock(p); void r; const lp = D.lastPara([prevB]); return D.pos(lp, D.plen(lp)); }); return; }
        const lp = D.lastPara([prevB]);
        E.setSel(D.pos(lp, D.plen(lp)));
        return;
      }
      if (O.tracking()) { E.edit('Delete', () => O.deleteRange(D.pos(prevB, D.plen(prevB)), pos)); return; }
      E.edit('Delete', () => O.joinNext(prevB) || pos);
      return;
    }
    let a = word ? wordStartBefore(p, pos.o) : prevCharOffset(p, pos.o);
    E.edit('Delete', () => {
      const r = O.deleteRange(D.pos(p, a), pos);
      O.repairFields(D.storyOf(d, p));
      return O.tracking() ? D.pos(p, a) : r;
    }, { merge: word ? null : 'del' });
  };
  E.del = function (word) {
    if (!E.sel) return;
    const d = doc();
    if (!E.collapsed()) { E.edit('Delete', () => E.deleteSelection()); return; }
    const pos = E.sel.f;
    const p = pos.p;
    const n = D.plen(p);
    if (pos.o >= n) {
      const info = D.info(d, p);
      const idx = info.cont.blocks.indexOf(p);
      const nb = info.cont.blocks[idx + 1];
      if (!nb) return;
      if (nb.t === 'tbl') return;
      if (O.tracking()) { E.edit('Delete', () => { O.deleteRange(pos, D.pos(nb, 0)); return pos; }); return; }
      E.edit('Delete', () => O.joinNext(p) || pos);
      return;
    }
    const b = word ? wordEndAfter(p, pos.o) : nextCharOffset(p, pos.o);
    E.edit('Delete', () => { O.deleteRange(pos, D.pos(p, b)); O.repairFields(D.storyOf(d, p)); return pos; }, { merge: word ? null : 'fdel' });
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
      E.edit('Numbering', () => {
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
      E.edit('Indent', () => { D.touch(p); p.pPr.ind = Object.assign({}, p.pPr.ind || {}, { fl: 36 }); return E.sel.f; });
      return;
    }
    E.insertItem(D.item('tab'), 'Typing');
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
        E.edit(label || 'Format', () => { O.setRunProps(w[0], w[1], props); return keep; });
        return;
      }
      E.pending = Object.assign({}, E.pending || {}, props);
      for (const k in props) if (props[k] === undefined) E.pending[k] = undefined;
      L.bus.emit('sel', E.sel);
      return;
    }
    const [a, b] = E.range();
    const keep = E.sel;
    E.edit(label || 'Format', () => { O.setRunProps(a, b, props); return keep; });
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
  /** is a run property uniform over the selection? returns value or undefined when mixed */
  E.uniformRun = function (key) {
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
        if (first) { val = v; first = false; } else if (JSON.stringify(v) !== JSON.stringify(val)) return undefined;
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
    E.cellSel = null;
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
      /* whole cells covered when selection runs from inside a table to outside */
      const [a, b] = E.range();
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
        if (!E.collapsed()) E.edit('Delete', () => E.deleteSelection());
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
      E.edit('Typing', () => E.deleteSelection());
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
      if (ctrl) E.insertItem(D.item('tab'), 'Typing'); else E.tab(shift);
      return;
    }
    if (k === 'Insert' && !ctrl && !shift) { e.preventDefault(); E.overtype = !E.overtype; L.bus.emit('sel', E.sel); return; }
    if (k === ' ' && ctrl && shift) { e.preventDefault(); E.typeText(' '); return; }
    if (k === 'Escape') { if (E.objSel) { E.objSel = null; L.bus.emit('objsel', null); } return; }
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
    /* typing over a selected floating object does nothing */
    if (k.length === 1 && !ctrl && !alt && E.objSel && E.objSel.float) { e.preventDefault(); return; }
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
    if (cut) E.edit('Cut', () => E.deleteSelection());
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
    E.edit('Paste', () => {
      const pos = E.deleteSelection();
      return O.insertBlocks(pos, blocks, opts);
    });
    E.lastPaste = Date.now();
  };
  E.pasteText = function (text) {
    E.edit('Paste', () => O.insertPlain(E.deleteSelection(), text.replace(/\r\n?/g, '\n').replace(/\n$/, '')));
  };
  /* internal drag & drop of text */
  let dragData = null;
  function onDragStart(e) {
    if (!E.sel || E.collapsed()) return;
    dragData = { range: E.range(), data: E.copyData() };
    try { e.dataTransfer.setData('text/plain', dragData.data.text); e.dataTransfer.setData('text/html', dragData.data.html); e.dataTransfer.effectAllowed = 'copyMove'; } catch (err) { /* ignore */ }
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
    if (dragData) {
      const [a, b] = dragData.range;
      const d = doc();
      const inside = D.cmp(d, target, a) > 0 && D.cmp(d, target, b) < 0;
      if (inside) { dragData = null; return; }
      const copy = e.ctrlKey || e.altKey;
      const blocks = dragData.data.blocks;
      E.edit(copy ? 'Drag and Drop' : 'Move', () => {
        /* insert first when dropping after the source, so positions stay valid */
        const after = D.cmp(d, target, b) >= 0;
        let end;
        if (after) {
          end = O.insertBlocks(target, blocks);
          if (!copy) O.deleteRange(a, b);
        } else {
          if (!copy) O.deleteRange(a, b);
          end = O.insertBlocks(target, blocks);
        }
        return end;
      });
      dragData = null;
      return;
    }
    const html = e.dataTransfer.getData('text/html'), text = e.dataTransfer.getData('text/plain');
    E.setSel(target);
    if (html && L.htmlio) { const bl = await L.htmlio.htmlToBlocks(html); if (bl.length) { E.pasteBlocks(bl); return; } }
    if (text) E.pasteText(text);
  }

  /* ---------- mouse ---------- */
  function onMouseDown(e) {
    hideHandles();
    if (e.button !== 0) return;
    const t = e.target;
    /* floating objects & inline images */
    const obj = t.closest && t.closest('.img, .shp, .flt');
    if (obj && L.drawing && L.drawing.onObjectDown && L.drawing.onObjectDown(e, obj)) return;
    if (obj && obj.classList.contains('img') && obj.closest('.p')) {
      e.preventDefault();
      const pos = LY.domToPos(obj.parentNode, Array.prototype.indexOf.call(obj.parentNode.childNodes, obj));
      if (pos) selectInlineObject(pos, obj);
      return;
    }
    /* header/footer double-click is handled in dblclick; a single click on a header in body mode just moves the caret */
    E.objSel = null;
    L.bus.emit('objsel', null);
    /* selection bar (left margin) */
    if (LY.view === 'print' || LY.view === 'normal') {
      const pg = t.closest && t.closest('.pg, .flow');
      if (pg && (t === pg || t.classList.contains('crop')) && !LY.hfEdit) {
        const body = pg.querySelector('.pg-body') || pg;
        const br = body.getBoundingClientRect();
        if (e.clientX < br.left - 2) {
          e.preventDefault();
          LY.root.focus({ preventScroll: true });
          const pos = posAtPoint(br.left + 2, e.clientY);
          if (!pos) return;
          if (e.ctrlKey) { E.selectAll(); return; }
          if (e.detail >= 2) { E.selectPara(pos.p); return; }
          selectLine(pos, e.shiftKey);
          return;
        }
      }
    }
    /* Ctrl+click selects a sentence */
    if (e.ctrlKey && !e.shiftKey && !(t.closest && t.closest('[data-link]'))) {
      setTimeout(() => { if (E.sel) { const s = O.sentenceExtent(E.sel.f); E.setSel(s[0], s[1]); } }, 0);
    }
  }
  function posAtPoint(x, y) {
    let node = null, off = 0;
    if (document.caretPositionFromPoint) { const cp = document.caretPositionFromPoint(x, y); if (cp) { node = cp.offsetNode; off = cp.offset; } }
    else if (document.caretRangeFromPoint) { const r = document.caretRangeFromPoint(x, y); if (r) { node = r.startContainer; off = r.startOffset; } }
    return node ? LY.domToPos(node, off) : null;
  }
  E.posAtPoint = posAtPoint;
  function selectLine(pos, extend) {
    const caret = pos;
    E.setSel(caret, caret);
    const s = window.getSelection();
    try {
      s.modify('move', 'backward', 'lineboundary');
      s.modify('extend', 'forward', 'lineboundary');
      const m = E.readDom();
      if (m) {
        if (extend && E.sel) E.setSel(E.sel.a, m.f); else { E.sel = m; E.selChanged(); }
      }
    } catch (err) { /* modify not supported */ }
  }
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
      const url = ln.dataset.link;
      if (url.startsWith('#')) { if (L.app && L.app.gotoBookmark) L.app.gotoBookmark(url.slice(1)); }
      else { const a = document.createElement('a'); a.href = url; a.target = '_blank'; a.rel = 'noopener'; a.click(); }
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
    if (LY.hfEdit && t.closest && t.closest('.pg-body') && L.app && L.app.closeHeaderFooter) { e.preventDefault(); L.app.closeHeaderFooter(); return; }
    const nr = t.closest && t.closest('.nref');
    if (nr && L.app && L.app.gotoNote) { e.preventDefault(); L.app.gotoNote(nr.dataset.note); return; }
    const img = t.closest && t.closest('.img, .shp');
    if (img && L.app && L.app.formatObject) { e.preventDefault(); L.app.formatObject(); return; }
  }
  /* inline object selection with resize handles */
  let handles = null;
  function selectInlineObject(pos, elObj) {
    const p = pos.p;
    const it = D.itemAfter(p, pos.o);
    E.setSel(D.pos(p, pos.o), D.pos(p, pos.o + 1), { noScroll: true });
    E.objSel = { p, o: pos.o, it, el: elObj, float: false };
    showHandles(elObj);
    L.bus.emit('objsel', E.objSel);
  }
  function hideHandles() { if (handles) { handles.remove(); handles = null; } }
  E.hideHandles = hideHandles;
  function showHandles(target) {
    hideHandles();
    const sc = LY.scroller;
    const box = document.createElement('div');
    box.className = 'obj-handles';
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
    if (!os || !os.it) return;
    const z = LY.zoom;
    const it = os.it;
    const w0 = it.w || 72, h0 = it.h || 72;
    const sx = ev.clientX, sy = ev.clientY;
    const corner = hname.length === 2;
    const ghost = document.createElement('div');
    ghost.className = 'resize-ghost';
    box.appendChild(ghost);
    let nw = w0, nh = h0;
    const mv = (e2) => {
      const dx = (e2.clientX - sx) / z / D.PX, dy = (e2.clientY - sy) / z / D.PX;
      nw = w0 + (hname.includes('e') ? dx : hname.includes('w') ? -dx : 0);
      nh = h0 + (hname.includes('s') ? dy : hname.includes('n') ? -dy : 0);
      if (corner && !e2.shiftKey) { const k = Math.max(nw / w0, nh / h0); nw = w0 * k; nh = h0 * k; }
      nw = Math.max(4, nw); nh = Math.max(4, nh);
      ghost.style.width = nw * D.PX * z + 'px';
      ghost.style.height = nh * D.PX * z + 'px';
      if (hname.includes('w')) ghost.style.left = (w0 - nw) * D.PX * z + 'px';
      if (hname.includes('n')) ghost.style.top = (h0 - nh) * D.PX * z + 'px';
    };
    const up = () => {
      window.removeEventListener('pointermove', mv);
      window.removeEventListener('pointerup', up);
      ghost.remove();
      if (Math.abs(nw - w0) < 0.5 && Math.abs(nh - h0) < 0.5) return;
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
    };
    window.addEventListener('pointermove', mv);
    window.addEventListener('pointerup', up);
  }
  E.selectInlineObject = selectInlineObject;
})();
