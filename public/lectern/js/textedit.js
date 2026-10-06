/* Lectern — in-place rich text editing.
 * The DOM is authoritative while typing; formatting commands go through the
 * model (split runs → apply → re-render) so markup never drifts.
 */
(function () {
  'use strict';
  const L = window.L;
  const T = L.txt;
  const E = L.ed;
  const TE = (L.te = { state: null, lastSel: null, pending: null });
  const ZW = /​/g;

  /* ---------- model text operations ---------- */
  T.len = (p) => T.paraText(p).length;
  T.cmp = (a, b) => (a.p - b.p) || (a.o - b.o);
  T.norm = (s) => (T.cmp({ p: s.p0, o: s.o0 }, { p: s.p1, o: s.o1 }) <= 0 ? s : { p0: s.p1, o0: s.o1, p1: s.p0, o1: s.o0 });
  T.collapsed = (s) => s.p0 === s.p1 && s.o0 === s.o1;
  /** run props in effect at a caret position (character before caret wins) */
  T.propsAt = (tx, pos) => {
    const p = tx.ps[pos.p];
    if (!p) return {};
    if (!p.rs.length) return Object.assign({}, p.end || {});
    let acc = 0;
    for (let i = 0; i < p.rs.length; i++) {
      const r = p.rs[i];
      acc += r.t.length;
      if (pos.o <= acc && (pos.o > acc - r.t.length || i === 0)) { const pr = T.runProps(r); delete pr.link; return pr; }
    }
    return T.runProps(p.rs[p.rs.length - 1]);
  };
  T.deleteRange = (tx, sel) => {
    sel = T.norm(sel);
    if (T.collapsed(sel)) return { p: sel.p0, o: sel.o0 };
    const a = tx.ps[sel.p0], b = tx.ps[sel.p1];
    if (sel.p0 === sel.p1) {
      const ia = T.splitAt(a, sel.o0), ib = T.splitAt(a, sel.o1);
      const keepProps = a.rs[ia] ? T.runProps(a.rs[ia]) : null;
      a.rs.splice(ia, ib - ia);
      if (!a.rs.length && keepProps) a.end = keepProps;
      T.normalize(a);
    } else {
      const ia = T.splitAt(a, sel.o0);
      const firstProps = a.rs[ia] ? T.runProps(a.rs[ia]) : null;
      a.rs.splice(ia);
      const ib = T.splitAt(b, sel.o1);
      a.rs = a.rs.concat(b.rs.slice(ib));
      if (!a.rs.length && firstProps) a.end = firstProps;
      tx.ps.splice(sel.p0 + 1, sel.p1 - sel.p0);
      T.normalize(a);
    }
    return { p: sel.p0, o: sel.o0 };
  };
  T.insertAt = (tx, pos, text, props) => {
    const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
    let p = tx.ps[pos.p];
    const base = props || T.propsAt(tx, pos);
    const insertRun = (para, off, str) => {
      if (!str) return;
      const i = T.splitAt(para, off);
      para.rs.splice(i, 0, Object.assign({ t: str }, base));
      T.normalize(para);
    };
    insertRun(p, pos.o, lines[0]);
    if (lines.length === 1) return { p: pos.p, o: pos.o + lines[0].length };
    /* split after the first line, insert middle lines as paragraphs */
    const at = pos.o + lines[0].length;
    const i = T.splitAt(p, at);
    const tail = p.rs.splice(i);
    let cur = pos.p;
    for (let k = 1; k < lines.length; k++) {
      const np = { lvl: p.lvl, pp: L.clone(p.pp || {}), rs: [], end: L.clone(base) };
      if (lines[k]) np.rs.push(Object.assign({ t: lines[k] }, base));
      tx.ps.splice(++cur, 0, np);
    }
    const last = tx.ps[cur];
    const endO = T.len(last);
    last.rs = last.rs.concat(tail);
    T.normalize(last);
    T.normalize(p);
    return { p: cur, o: endO };
  };
  T.splitPara = (tx, pos) => {
    const p = tx.ps[pos.p];
    const props = T.propsAt(tx, pos);
    const i = T.splitAt(p, pos.o);
    const tail = p.rs.splice(i);
    const np = { lvl: p.lvl, pp: L.clone(p.pp || {}), rs: tail, end: L.clone(props) };
    if (!p.rs.length) p.end = L.clone(props);
    tx.ps.splice(pos.p + 1, 0, np);
    return { p: pos.p + 1, o: 0 };
  };
  T.wordAt = (tx, pos) => {
    const s = T.paraText(tx.ps[pos.p] || { rs: [] });
    const isW = (c) => /[\p{L}\p{N}_'’-]/u.test(c);
    if (!(isW(s[pos.o - 1] || ' ') && isW(s[pos.o] || ' '))) return null;
    let a = pos.o, b = pos.o;
    while (a > 0 && isW(s[a - 1])) a--;
    while (b < s.length && isW(s[b])) b++;
    return { p0: pos.p, o0: a, p1: pos.p, o1: b };
  };

  /* ---------- DOM <-> model ---------- */
  const isPara = (n) => n.nodeType === 1 && (n.tagName === 'P' || n.tagName === 'DIV');
  function readJSON(s) { try { return s ? JSON.parse(s) : {}; } catch (e) { return {}; } }
  function propsOf(node, root, fallback) {
    let props = null;
    const extra = {};
    for (let n = node.nodeType === 1 ? node : node.parentNode; n && n !== root && !isPara(n); n = n.parentNode) {
      if (n.nodeType !== 1) continue;
      if (!props && n.dataset && n.dataset.r != null) props = readJSON(n.dataset.r);
      if (n.classList && n.classList.contains('r') && !props) props = {};
      const tg = n.tagName;
      if (tg === 'B' || tg === 'STRONG') extra.b = true;
      if (tg === 'I' || tg === 'EM') extra.i = true;
      if (tg === 'U') extra.u = 'sng';
      if (tg === 'STRIKE' || tg === 'S') extra.strike = 'sngStrike';
      if (tg === 'SUP') extra.base = 30000;
      if (tg === 'SUB') extra.base = -25000;
    }
    return Object.assign({}, props || fallback || {}, extra);
  }
  /** Serialize an editing root into paragraphs */
  TE.serialize = function (root, fallbackParas) {
    const fb = fallbackParas || [];
    const out = [];
    let cur = null;
    let lastProps = fb[0] && fb[0].rs[0] ? T.runProps(fb[0].rs[0]) : (fb[0] && fb[0].end) || {};
    const newPara = (el) => {
      const lvl = el && el.dataset ? +el.dataset.lvl || 0 : (out[out.length - 1] || fb[0] || { lvl: 0 }).lvl || 0;
      const pp = el && el.dataset && el.dataset.pp ? readJSON(el.dataset.pp) : el ? (out[out.length - 1] ? L.clone(out[out.length - 1].pp) : L.clone((fb[0] && fb[0].pp) || {})) : L.clone((fb[0] && fb[0].pp) || {});
      cur = { lvl, pp, rs: [], end: el && el.dataset && el.dataset.end ? readJSON(el.dataset.end) : null };
      out.push(cur);
      return cur;
    };
    const addText = (txt, props, fld) => {
      if (!txt && !fld) return;
      const last = cur.rs[cur.rs.length - 1];
      if (!fld && last && !last.fld && L.equal(T.runProps(last), props)) last.t += txt;
      else cur.rs.push(Object.assign({ t: txt }, props, fld ? { fld } : {}));
      lastProps = props;
    };
    const walk = (node, para) => {
      for (let n = node.firstChild; n; n = n.nextSibling) {
        if (n.nodeType === 3) { const t = n.data.replace(ZW, ''); if (t) addText(t, propsOf(n, para, lastProps)); continue; }
        if (n.nodeType !== 1) continue;
        if (n.tagName === 'BR') {
          /* trailing <br> is a placeholder, not a line break */
          if (isTrailing(n, para)) continue;
          addText('\n', propsOf(n, para, lastProps));
          continue;
        }
        if (n.dataset && n.dataset.fld) { addText(n.textContent, propsOf(n, para, lastProps), n.dataset.fld); continue; }
        if (isPara(n)) { newPara(n); walk(n, n); continue; }
        walk(n, para);
      }
    };
    function isTrailing(br, para) {
      for (let n = br; n && n !== para; n = n.parentNode) {
        let s = n.nextSibling;
        while (s && ((s.nodeType === 3 && !s.data.replace(ZW, '')) || (s.nodeType === 1 && s.tagName !== 'BR' && !s.textContent.replace(ZW, '') && !s.querySelector('br')))) s = s.nextSibling;
        if (s) return false;
      }
      return true;
    }
    /* loose inline nodes directly under root form an implicit paragraph */
    for (let n = root.firstChild; n; n = n.nextSibling) {
      if (isPara(n)) { newPara(n); walk(n, n); }
      else {
        if (!cur || cur._implicit !== true) { newPara(null); cur._implicit = true; }
        const tmp = document.createElement('span');
        if (n.nodeType === 3) { const t = n.data.replace(ZW, ''); if (t) addText(t, lastProps); }
        else if (n.tagName === 'BR') { /* ignore */ }
        else { tmp.appendChild(n.cloneNode(true)); walk(tmp, tmp); }
      }
    }
    for (const p of out) {
      delete p._implicit;
      if (!p.rs.length && !p.end) p.end = L.clone(lastProps);
      if (!p.end) delete p.end;
      if (p.rs.length) delete p.end;
    }
    if (!out.length) out.push({ lvl: (fb[0] && fb[0].lvl) || 0, pp: L.clone((fb[0] && fb[0].pp) || {}), rs: [], end: L.clone(lastProps) });
    return out;
  };
  function paraEls(root) { return Array.from(root.children).filter(isPara); }
  function paraLen(pEl) {
    const r = document.createRange();
    r.selectNodeContents(pEl);
    return measure(r, pEl);
  }
  function measure(range, pEl) {
    const frag = range.cloneContents();
    let n = frag.textContent.replace(ZW, '').length;
    const brs = frag.querySelectorAll('br').length;
    n += brs;
    /* do not count the trailing placeholder <br> */
    const all = pEl.querySelectorAll('br');
    if (all.length) {
      const lastBr = all[all.length - 1];
      const after = document.createRange();
      after.setStartAfter(lastBr);
      after.setEnd(pEl, pEl.childNodes.length);
      const tail = after.toString().replace(ZW, '');
      if (!tail && range.comparePoint(lastBr, 0) <= 0 && brs) {
        const end = document.createRange(); end.setStartBefore(lastBr); end.setEndAfter(lastBr);
        if (range.compareBoundaryPoints(Range.END_TO_END, end) >= 0) n -= 1;
      }
    }
    return n;
  }
  function toOffset(root, node, off) {
    const ps = paraEls(root);
    if (node === root) {
      const i = Math.min(off, root.childNodes.length);
      let pi = 0;
      for (let k = 0; k < i; k++) if (isPara(root.childNodes[k])) pi++;
      if (pi >= ps.length) return { p: Math.max(0, ps.length - 1), o: ps.length ? paraLen(ps[ps.length - 1]) : 0 };
      return { p: pi, o: 0 };
    }
    let pEl = node.nodeType === 1 && isPara(node) && node.parentNode === root ? node : null;
    for (let n = node; !pEl && n && n !== root; n = n.parentNode) if (n.parentNode === root && isPara(n)) pEl = n;
    if (!pEl) return { p: 0, o: 0 };
    const r = document.createRange();
    r.setStart(pEl, 0);
    try { r.setEnd(node, off); } catch (e) { return { p: ps.indexOf(pEl), o: 0 }; }
    return { p: ps.indexOf(pEl), o: Math.min(measure(r, pEl), paraLen(pEl)) };
  }
  function toPoint(root, pos) {
    const ps = paraEls(root);
    const pEl = ps[L.clamp(pos.p, 0, ps.length - 1)];
    if (!pEl) return [root, 0];
    let remaining = pos.o;
    const w = document.createTreeWalker(pEl, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
    let last = null;
    for (let n = w.nextNode(); n; n = w.nextNode()) {
      if (n.nodeType === 1 && n.dataset && n.dataset.fld) {
        const L2 = n.textContent.length;
        if (remaining <= 0) return [n.parentNode, Array.prototype.indexOf.call(n.parentNode.childNodes, n)];
        if (remaining <= L2) return [n.parentNode, Array.prototype.indexOf.call(n.parentNode.childNodes, n) + 1];
        remaining -= L2;
        /* skip field children */
        let skip = n; while (skip.lastChild) skip = skip.lastChild; w.currentNode = skip;
        last = [n.parentNode, Array.prototype.indexOf.call(n.parentNode.childNodes, n) + 1];
        continue;
      }
      if (n.nodeType === 3) {
        const t = n.data.replace(ZW, '');
        if (remaining <= t.length) return [n, remaining];
        remaining -= t.length;
        last = [n, n.data.length];
      } else if (n.tagName === 'BR') {
        if (remaining === 0) return [n.parentNode, Array.prototype.indexOf.call(n.parentNode.childNodes, n)];
        remaining -= 1;
        last = [n.parentNode, Array.prototype.indexOf.call(n.parentNode.childNodes, n) + 1];
      }
    }
    if (last) return last;
    /* empty paragraph: caret inside the first span if any, so typing inherits its style */
    const sp = pEl.querySelector('span.r');
    if (sp) return [sp, 0];
    return [pEl, 0];
  }
  TE.getSel = function (root) {
    root = root || (TE.state && TE.state.root);
    if (!root) return null;
    const s = window.getSelection();
    if (!s || !s.rangeCount) return null;
    const r = s.getRangeAt(0);
    if (!root.contains(r.startContainer) || !root.contains(r.endContainer)) return null;
    const a = toOffset(root, r.startContainer, r.startOffset);
    const b = toOffset(root, r.endContainer, r.endOffset);
    return { p0: a.p, o0: a.o, p1: b.p, o1: b.o };
  };
  TE.setSel = function (sel, root) {
    root = root || (TE.state && TE.state.root);
    if (!root || !sel) return;
    const [an, ao] = toPoint(root, { p: sel.p0, o: sel.o0 });
    const [bn, bo] = toPoint(root, { p: sel.p1, o: sel.o1 });
    const s = window.getSelection();
    const r = document.createRange();
    try { r.setStart(an, ao); r.setEnd(bn, bo); } catch (e) { r.selectNodeContents(root); r.collapse(false); }
    s.removeAllRanges();
    s.addRange(r);
    TE.lastSel = sel;
  };
  TE.currentSel = () => TE.getSel() || TE.lastSel || (TE.target() ? T.allRange(TE.target().tx) : null);

  /* ---------- session ---------- */
  TE.active = () => !!TE.state;
  TE.target = function () {
    const st = TE.state;
    if (!st) return null;
    const sh = E.shape(st.id);
    if (!sh) return null;
    let tx = sh.tx;
    if (st.cell) { const c = sh.tbl.rows[st.cell[0]] && sh.tbl.rows[st.cell[0]].cells[st.cell[1]]; tx = c ? c.tx : null; }
    if (!tx) return null;
    return { sh, tx, root: st.root, cell: st.cell, el: st.el };
  };
  TE.contains = (node) => {
    const st = TE.state;
    if (!st || !node) return false;
    if (st.cell) return !!(node.closest && node.closest('td') && st.el.contains(node));
    return st.root.contains(node);
  };
  TE.focus = () => { if (TE.state) { const r = TE.state.root; if (document.activeElement !== r) { try { r.focus({ preventScroll: true }); } catch (e) { r.focus(); } if (TE.lastSel) TE.setSel(TE.lastSel); } } };
  TE.snapshot = () => (TE.state ? { id: TE.state.id, cell: TE.state.cell ? TE.state.cell.slice() : null, sel: TE.getSel() || TE.lastSel } : null);

  function cellRoot(el, r, c) {
    const td = el.querySelector(`td[data-r="${r}"][data-c="${c}"]`);
    return td ? td.querySelector('.txi') : null;
  }
  /**
   * Start editing. opts: {point:{x,y}, sel, cell:[r,c], atEnd, selectAll, word}
   */
  TE.begin = function (id, opts) {
    opts = opts || {};
    if (TE.state) TE.end();
    let sh = E.shape(id);
    if (!sh) return;
    if (sh.type === 'shape' && !sh.tx) {
      L.hist.push('Add Text');
      sh.tx = T.body([T.para('', { algn: 'ctr' })], { anchor: 'ctr' });
      E.renderShape(E.topOf(id).id);
    }
    const top = E.topOf(id);
    if (!E.sel.includes(id) || E.sel.length !== 1) { E.sel = [id]; L.bus.emit('selection'); }
    const el = E.elOf(id);
    if (!el) return;
    let root, cell = null;
    if (sh.type === 'table') {
      cell = opts.cell || [0, 0];
      if (opts.point) {
        const hit = document.elementFromPoint(opts.point.x, opts.point.y);
        const td = hit && hit.closest && hit.closest('td');
        if (td && el.contains(td)) cell = [+td.dataset.r, +td.dataset.c];
      }
      L.$$('td .txi', el).forEach((x) => { x.contentEditable = 'true'; x.spellcheck = !!(L.app && L.app.opts.spell); });
      root = cellRoot(el, cell[0], cell[1]);
    } else {
      const tb = el.querySelector(':scope > .tx:not(.prompt)');
      if (!tb) return;
      root = tb.querySelector('.txi');
      root.contentEditable = 'true';
      root.spellcheck = !!(L.app && L.app.opts.spell);
    }
    if (!root) return;
    el.classList.add('editing');
    TE.state = { id, root, el, cell, top: top ? top.id : id };
    attach(root);
    try { root.focus({ preventScroll: true }); } catch (e) { root.focus(); }
    const tx = TE.target().tx;
    if (opts.sel) TE.setSel(opts.sel);
    else if (opts.point) {
      const rg = caretFromPoint(opts.point.x, opts.point.y);
      if (rg && root.contains(rg.startContainer)) { const s = window.getSelection(); s.removeAllRanges(); s.addRange(rg); if (opts.word && s.modify) { s.modify('move', 'backward', 'word'); s.modify('extend', 'forward', 'word'); } }
      else TE.setSel(endSel(tx));
    } else if (opts.selectAll) TE.setSel(T.allRange(tx));
    else TE.setSel(endSel(tx));
    TE.lastSel = TE.getSel();
    E.drawOverlay();
    L.bus.emit('edit-start');
    L.ui.refresh();
  };
  const endSel = (tx) => { const p = tx.ps.length - 1; const o = T.len(tx.ps[p]); return { p0: p, o0: o, p1: p, o1: o }; };
  function caretFromPoint(x, y) {
    if (document.caretRangeFromPoint) return document.caretRangeFromPoint(x, y);
    if (document.caretPositionFromPoint) { const p = document.caretPositionFromPoint(x, y); if (!p) return null; const r = document.createRange(); r.setStart(p.offsetNode, p.offset); r.collapse(true); return r; }
    return null;
  }
  TE.resume = function (snap) {
    if (!snap) return;
    TE.begin(snap.id, { cell: snap.cell, sel: snap.sel || undefined });
  };
  /** Stop editing without re-rendering. */
  TE.detach = function (noSync) {
    const st = TE.state;
    if (!st) return;
    if (!noSync) TE.sync();
    detachListeners(st.root);
    if (st.cell) L.$$('td .txi', st.el).forEach((x) => { x.contentEditable = 'false'; x.removeAttribute('contenteditable'); });
    else { st.root.removeAttribute('contenteditable'); }
    st.el.classList.remove('editing');
    TE.state = null;
    TE.pending = null;
    E.cellSel = null;
  };
  /** Finish editing, apply autofit and re-render. */
  TE.end = function () {
    const st = TE.state;
    if (!st) return;
    const t = TE.target();
    TE.sync();
    TE.detach(true);
    if (t) {
      const sh = t.sh;
      if (sh.type !== 'table' && sh.tx && sh.tx.autofit === 'norm') shrinkToFit(sh);
      if (E.view === 'master') E.syncMaster();
      const top = E.topOf(sh.id);
      if (top) E.renderShape(top.id);
      if (sh.type === 'table') E.render();
    }
    E.drawOverlay();
    L.bus.emit('edit-end');
    E.touchNow();
    L.ui.refresh();
  };
  /** PowerPoint-style "shrink text on overflow" for placeholders */
  function shrinkToFit(sh) {
    const el = E.elOf(sh.id);
    if (!el) return;
    const tb = el.querySelector(':scope > .tx:not(.prompt)');
    if (!tb) return;
    const inner = tb.querySelector('.txi');
    const avail = tb.offsetHeight;
    const cur = sh.tx.fontScale || 1;
    const need = inner.offsetHeight;
    if (need <= avail + 1 && cur === 1) return;
    let scale = cur;
    if (need > avail + 1) scale = Math.max(0.25, Math.floor(cur * (avail / need) * 20) / 20);
    else if (cur < 1) scale = Math.min(1, Math.floor(cur * (avail / need) * 20) / 20);
    if (Math.abs(scale - cur) < 0.01) return;
    sh.tx.fontScale = scale >= 0.99 ? undefined : scale;
    if (!sh.tx.fontScale) delete sh.tx.fontScale;
  }
  TE.shrinkToFit = shrinkToFit;

  /** Read DOM into the model (no re-render). */
  TE.sync = function () {
    const t = TE.target();
    if (!t) return null;
    const paras = TE.serialize(t.root, t.tx.ps);
    paras.forEach(T.normalize);
    t.tx.ps = paras;
    return paras;
  };
  /** Re-render the paragraphs inside the editing root and restore selection */
  TE.rerender = function (sel) {
    const t = TE.target();
    if (!t) return;
    const design = E.design();
    const ctx = { pres: L.pres, slide: E.slide(), design, mode: 'edit', num: E.idx + (L.pres.firstNum || 1) };
    L.clear(t.root);
    L.render.paragraphs(t.sh, t.tx, design, ctx, t.root, {});
    if (sel) TE.setSel(sel);
    afterChange(t);
  };
  function afterChange(t) {
    t = t || TE.target();
    if (!t) return;
    const sh = t.sh;
    if (sh.type === 'table') fitTableRows(sh, t.el);
    else if (sh.tx && sh.tx.autofit === 'shape') autoSize(sh, t);
    E.drawOverlay();
    E.touched();
  }
  function autoSize(sh, t) {
    const tb = t.root.parentElement;
    const ins = sh.tx.ins || [7.2, 3.6, 7.2, 3.6];
    let changed = false;
    if (sh.tx.wrap === false && !sh.tx.vert) {
      const prevW = t.root.style.width;
      t.root.style.width = 'max-content';
      const w = Math.max(t.root.offsetWidth + ins[0] + ins[2] + 1, 18);
      t.root.style.width = prevW;
      if (Math.abs(w - sh.w) > 0.5) { sh.w = L.round(w, 2); changed = true; }
    }
    const hgt = (sh.tx.vert ? t.root.offsetWidth : t.root.offsetHeight) + ins[1] + ins[3];
    if (Math.abs(hgt - sh.h) > 0.5) { sh.h = L.round(Math.max(hgt, 10), 2); changed = true; }
    if (changed) {
      t.el.style.width = sh.w + 'px'; t.el.style.height = sh.h + 'px';
      tb.style.width = Math.max(0, sh.w - ins[0] - ins[2]) + 'px';
      tb.style.height = Math.max(0, sh.h - ins[1] - ins[3]) + 'px';
      const geo = t.el.querySelector(':scope > svg.geo');
      if (geo) { const top = E.topOf(sh.id); if (top && top.id === sh.id) { const snap = TE.snapshot(); TE.detach(true); E.renderShape(sh.id); TE.resume(snap); } }
      if (E.find(sh.id) && E.find(sh.id).parent) E.fitGroup(E.find(sh.id).parent);
    }
  }
  function fitTableRows(sh, el) {
    const trs = el.querySelectorAll('tr');
    let total = 0;
    trs.forEach((tr, i) => {
      const hh = tr.offsetHeight;
      if (sh.tbl.rows[i]) { if (hh > sh.tbl.rows[i].h + 0.5) sh.tbl.rows[i].h = L.round(hh, 2); total += Math.max(hh, sh.tbl.rows[i].h); }
    });
    const tbl = el.querySelector('table');
    const realH = tbl ? tbl.offsetHeight : total;
    if (Math.abs(realH - sh.h) > 0.5) { sh.h = L.round(realH, 2); el.style.height = sh.h + 'px'; }
  }
  TE.fitTableRows = fitTableRows;

  /* ---------- listeners ---------- */
  function attach(root) {
    root.addEventListener('beforeinput', onBeforeInput);
    root.addEventListener('input', onInput);
    root.addEventListener('paste', onPaste);
    root.addEventListener('focusin', onFocusIn, true);
    if (TE.state && TE.state.cell) L.$$('td .txi', TE.state.el).forEach((x) => { if (x !== root) { x.addEventListener('beforeinput', onBeforeInput); x.addEventListener('input', onInput); x.addEventListener('paste', onPaste); x.addEventListener('focusin', onFocusIn, true); } });
  }
  function detachListeners(root) {
    const all = TE.state && TE.state.cell ? L.$$('td .txi', TE.state.el) : [root];
    for (const x of all) { x.removeEventListener('beforeinput', onBeforeInput); x.removeEventListener('input', onInput); x.removeEventListener('paste', onPaste); x.removeEventListener('focusin', onFocusIn, true); }
  }
  function onFocusIn(e) {
    const st = TE.state;
    if (!st || !st.cell) return;
    const td = e.target.closest && e.target.closest('td');
    if (!td) return;
    const r = +td.dataset.r, c = +td.dataset.c;
    if (r !== st.cell[0] || c !== st.cell[1]) { TE.sync(); st.cell = [r, c]; st.root = td.querySelector('.txi'); TE.lastSel = null; L.ui.refresh(); }
  }
  document.addEventListener('selectionchange', () => {
    if (!TE.state) return;
    const s = TE.getSel();
    if (s) {
      if (TE.pending && (s.p0 !== TE.pending.at.p || s.o0 !== TE.pending.at.o || !T.collapsed(s))) TE.pending = null;
      TE.lastSel = s;
      if (E.tool && E.tool.kind === 'painter' && !T.collapsed(s) && !TE._painting) {
        TE._painting = true;
        const fmt = E.tool.fmt;
        const done = () => { document.removeEventListener('pointerup', done); TE._painting = false; const s2 = TE.getSel(); if (s2 && !T.collapsed(s2) && fmt.run) { TE.applyRun((r) => Object.assign({}, r, fmt.run), 'Format Painter', s2); if (!E.tool.sticky) E.setTool(null); } };
        document.addEventListener('pointerup', done);
      }
      L.ui.refresh();
    }
  });

  const SMART = { '"': ['“', '”'], "'": ['‘', '’'] };
  function onBeforeInput(e) {
    const st = TE.state;
    if (!st) return;
    const t = e.inputType;
    if (t === 'historyUndo' || t === 'historyRedo') { e.preventDefault(); L.ui.exec(t === 'historyUndo' ? 'undo' : 'redo'); return; }
    L.hist.push('Typing', 'type:' + st.id + ':' + (st.cell ? st.cell.join(',') : ''));
    if (t === 'insertParagraph') { e.preventDefault(); TE.newParagraph(); return; }
    if (t === 'insertLineBreak') { e.preventDefault(); TE.insertText('\n', { asBreak: true }); return; }
    if (t === 'insertText' && e.data) {
      let data = e.data;
      if (L.app && L.app.opts.smartQuotes && SMART[data]) {
        const s = TE.getSel();
        const prev = s ? T.paraText(TE.target().tx.ps[s.p0] || { rs: [] }).slice(0, s.o0) : '';
        const pc = prev.slice(-1);
        data = !pc || /[\s([{“‘—–-]/.test(pc) ? SMART[data][0] : SMART[data][1];
      }
      const s = TE.getSel();
      if ((TE.pending && s && T.collapsed(s) && s.p0 === TE.pending.at.p && s.o0 === TE.pending.at.o) || data !== e.data) {
        e.preventDefault();
        TE.sync();
        const tx = TE.target().tx;
        const base = Object.assign(T.propsAt(tx, { p: s.p0, o: s.o0 }), TE.pending ? TE.pending.props : {});
        const c = T.deleteRange(tx, s);
        const np = T.insertAt(tx, c, data, base);
        TE.pending = null;
        TE.rerender({ p0: np.p, o0: np.o, p1: np.p, o1: np.o });
      }
    }
  }
  function onInput(e) {
    const t = TE.target();
    if (!t) return;
    /* self-heal: anything that isn't one of our paragraphs gets normalized */
    const bad = Array.from(t.root.childNodes).some((n) => !(n.nodeType === 1 && n.tagName === 'P'));
    if (bad) {
      const sel = TE.getSel();
      TE.sync();
      TE.rerender(sel || endSel(t.tx));
    } else TE.sync();
    if (e && e.inputType === 'insertText' && e.data && /[\s.,;:!?)]/.test(e.data)) autocorrect(e.data);
    afterChange(t);
    if (L.app && L.app.outlineDirty) L.app.outlineDirty();
  }
  function onPaste(e) {
    if (!TE.state) return;
    const dt = e.clipboardData;
    if (!dt) return;
    const files = Array.from(dt.files || []).filter((f) => /^image\//.test(f.type));
    e.preventDefault();
    if (files.length && L.app) { TE.end(); L.app.insertPictureFiles(files); return; }
    const text = dt.getData('text/plain');
    if (!text) return;
    L.hist.push('Paste');
    TE.insertText(text.replace(/\r\n?/g, '\n').replace(/\t/g, '\t'));
  }

  /* ---------- AutoCorrect ---------- */
  const AC = {
    '(c)': '©', '(r)': '®', '(tm)': '™', '(e)': '€', '...': '…', '-->': '→', '<--': '←', '==>': '⇒', '<==': '⇐', '<=>': '⇔', ':)': '☺', ':(': '☹',
    teh: 'the', adn: 'and', taht: 'that', recieve: 'receive', seperate: 'separate', occured: 'occurred', definately: 'definitely', wich: 'which', thier: 'their',
    becuase: 'because', accomodate: 'accommodate', acheive: 'achieve', beleive: 'believe', calender: 'calendar', goverment: 'government', untill: 'until', tommorow: 'tomorrow',
    i: 'I', "dont": "don't", "doesnt": "doesn't", "cant": "can't", "wont": "won't", "isnt": "isn't", "didnt": "didn't",
  };
  TE.AUTOCORRECT = AC;
  function autocorrect(typed) {
    if (!(L.app && L.app.opts.autocorrect)) return;
    const s = window.getSelection();
    if (!s.rangeCount || !s.isCollapsed) return;
    const node = s.anchorNode;
    if (!node || node.nodeType !== 3) return;
    const off = s.anchorOffset;
    const before = node.data.slice(0, off - typed.length);
    const m = /(\S+)$/.exec(before);
    if (!m) return;
    const word = m[1];
    let rep = null;
    const lower = word.toLowerCase();
    if (AC[word] != null) rep = AC[word];
    else if (AC[lower] != null && /^[a-z]/.test(word)) rep = AC[lower];
    else if (/^[A-Z]{2}[a-z]{2,}$/.test(word)) rep = word[0] + word.slice(1).toLowerCase();
    else if (L.app.opts.capSentence && /^[a-z]/.test(word)) {
      const t = TE.target();
      const sel = TE.getSel();
      if (t && sel) {
        const ptext = T.paraText(t.tx.ps[sel.p0]).slice(0, sel.o0 - typed.length - word.length);
        if (!ptext.trim() || /[.!?]\s+$/.test(ptext)) rep = word[0].toUpperCase() + word.slice(1);
      }
    }
    if (rep == null || rep === word) return;
    const start = off - typed.length - word.length;
    node.data = node.data.slice(0, start) + rep + node.data.slice(off - typed.length);
    const r = document.createRange();
    r.setStart(node, start + rep.length + typed.length);
    r.collapse(true);
    s.removeAllRanges();
    s.addRange(r);
    TE.sync();
  }

  /* ---------- editing operations ---------- */
  TE.newParagraph = function () {
    const t = TE.target();
    if (!t) return;
    TE.sync();
    const s = TE.currentSel();
    const c = T.deleteRange(t.tx, s);
    const np = T.splitPara(t.tx, c);
    TE.pending = null;
    TE.rerender({ p0: np.p, o0: 0, p1: np.p, o1: 0 });
    if (L.app && L.app.outlineDirty) L.app.outlineDirty();
  };
  TE.insertText = function (str, opts) {
    const t = TE.target();
    if (!t) return;
    TE.sync();
    const s = TE.currentSel();
    const c = T.deleteRange(t.tx, s);
    let np;
    if (opts && opts.asBreak) {
      /* soft line break: a '\n' inside a run */
      const p = t.tx.ps[c.p];
      const props = T.propsAt(t.tx, c);
      const i = T.splitAt(p, c.o);
      p.rs.splice(i, 0, Object.assign({ t: '\n' }, props));
      T.normalize(p);
      np = { p: c.p, o: c.o + 1 };
    } else np = T.insertAt(t.tx, c, str, TE.pending ? Object.assign(T.propsAt(t.tx, c), TE.pending.props) : null);
    TE.pending = null;
    TE.rerender({ p0: np.p, o0: np.o, p1: np.p, o1: np.o });
  };
  TE.insertField = function (fld) {
    const t = TE.target();
    if (!t) return;
    TE.sync();
    const s = TE.currentSel();
    const c = T.deleteRange(t.tx, s);
    const p = t.tx.ps[c.p];
    const props = T.propsAt(t.tx, c);
    const text = fld === 'slidenum' ? String(E.idx + (L.pres.firstNum || 1)) : L.fmtDate(new Date(), fld);
    const i = T.splitAt(p, c.o);
    p.rs.splice(i, 0, Object.assign({ t: text, fld }, props));
    TE.rerender({ p0: c.p, o0: c.o + text.length, p1: c.p, o1: c.o + text.length });
  };
  /** Apply a run transform to the selection (editing) */
  TE.applyRun = function (fn, label, selOverride) {
    const t = TE.target();
    if (!t) return;
    TE.sync();
    let s = selOverride || TE.currentSel();
    if (!s) return;
    s = T.norm(s);
    if (T.collapsed(s)) {
      const w = T.wordAt(t.tx, { p: s.p0, o: s.o0 });
      if (w) {
        L.hist.push(label || 'Format');
        T.mapRange(t.tx, w, fn);
        TE.rerender(s);
      } else {
        const base = T.propsAt(t.tx, { p: s.p0, o: s.o0 });
        const p = t.tx.ps[s.p0];
        if (!T.len(p)) { L.hist.push(label || 'Format'); p.end = fn(Object.assign({}, p.end || base)); delete p.end.t; TE.rerender(s); }
        else TE.pending = { at: { p: s.p0, o: s.o0 }, props: (() => { const r = fn(Object.assign({}, base, TE.pending ? TE.pending.props : {})); delete r.t; return r; })() };
      }
    } else {
      L.hist.push(label || 'Format');
      T.mapRange(t.tx, s, fn);
      TE.rerender(s);
    }
    if (E.view === 'master') E.syncMaster();
    L.ui.refresh();
  };
  TE.applyPara = function (fn, label) {
    const t = TE.target();
    if (!t) return;
    TE.sync();
    const s = T.norm(TE.currentSel() || T.allRange(t.tx));
    L.hist.push(label || 'Format Paragraph');
    for (let i = s.p0; i <= s.p1; i++) if (t.tx.ps[i]) fn(t.tx.ps[i], i, t);
    TE.rerender(s);
    if (E.view === 'master') E.syncMaster();
    L.ui.refresh();
  };
  TE.runPropsAtSel = function () {
    const t = TE.target();
    if (!t) return {};
    const s = TE.currentSel();
    if (!s) return {};
    const n = T.norm(s);
    const pos = T.collapsed(n) ? { p: n.p0, o: n.o0 } : { p: n.p0, o: Math.min(n.o0 + 1, T.len(t.tx.ps[n.p0])) };
    return Object.assign(T.propsAt(t.tx, pos), TE.pending ? TE.pending.props : {});
  };
  TE.paraPropsAtSel = function () {
    const t = TE.target();
    if (!t) return {};
    const s = TE.currentSel();
    const p = t.tx.ps[s ? T.norm(s).p0 : 0];
    return p ? L.clone(p.pp || {}) : {};
  };
  TE.selectedText = function () {
    const t = TE.target();
    if (!t) return '';
    const s = TE.getSel();
    if (!s || T.collapsed(s)) return '';
    return String(window.getSelection());
  };
  TE.atParaStart = function () {
    const s = TE.getSel();
    return !!s && T.collapsed(s) && s.o0 === 0;
  };
  TE.multiPara = function () { const s = TE.getSel(); return !!s && s.p0 !== s.p1; };

  /** Key handling while editing; return true if consumed */
  TE.onKey = function (e) {
    const t = TE.target();
    if (!t) return false;
    const k = e.key;
    if (k === 'Escape') { e.preventDefault(); const id = t.sh.id; TE.end(); E.select([id]); E.refocus(); return true; }
    if (k === 'F2') { e.preventDefault(); const id = t.sh.id; TE.end(); E.select([id]); E.refocus(); return true; }
    if (k === 'Enter' && !e.ctrlKey && !e.altKey && !e.metaKey) {
      e.preventDefault();
      L.hist.push('Typing', 'type:' + TE.state.id + ':' + (TE.state.cell ? TE.state.cell.join(',') : ''));
      if (e.shiftKey) TE.insertText('\n', { asBreak: true }); else TE.newParagraph();
      return true;
    }
    if (k === 'Tab' && !e.ctrlKey && !e.altKey) {
      e.preventDefault();
      if (t.sh.type === 'table') { TE.moveCell(e.shiftKey ? -1 : 1); return true; }
      if (TE.atParaStart() || TE.multiPara() || e.shiftKey) L.ui.exec(e.shiftKey ? 'promote' : 'demote');
      else { L.hist.push('Typing', 'type:' + TE.state.id + ':' + (TE.state.cell ? TE.state.cell.join(',') : '')); TE.insertText('\t'); }
      return true;
    }
    if (k === 'Backspace' && !e.ctrlKey && TE.atParaStart()) {
      const s = TE.getSel();
      const p = t.tx.ps[s.p0];
      if (p && p.lvl > 0 && t.sh.type !== 'table') { e.preventDefault(); L.ui.exec('promote'); return true; }
    }
    return false;
  };
  TE.moveCell = function (dir) {
    const st = TE.state;
    const sh = TE.target().sh;
    TE.sync();
    const rows = sh.tbl.rows.length, cols = sh.tbl.cols.length;
    let [r, c] = st.cell;
    do {
      c += dir;
      if (c >= cols) { c = 0; r++; }
      if (c < 0) { c = cols - 1; r--; }
    } while (r >= 0 && r < rows && sh.tbl.rows[r].cells[c] && (sh.tbl.rows[r].cells[c].hm || sh.tbl.rows[r].cells[c].vm));
    if (r >= rows) {
      L.hist.push('Insert Row');
      L.app.tableAddRow(sh, rows - 1, true);
      const id = sh.id;
      TE.detach(true);
      E.render();
      TE.begin(id, { cell: [rows, 0] });
      return;
    }
    if (r < 0) return;
    const root = cellRoot(st.el, r, c);
    if (!root) return;
    st.cell = [r, c]; st.root = root;
    root.focus();
    TE.setSel(T.allRange(TE.target().tx));
    L.ui.refresh();
  };

  /* ---------- formatting API (editing or whole shapes) ---------- */
  const F = (L.fmt = {});
  /** shapes whose text is affected when not editing */
  F.textShapes = () => {
    const out = [];
    for (const sh of E.selected()) M_walk(sh, (s) => { if (s.tx || s.type === 'table' || s.type === 'wordart') out.push(s); });
    return out;
  };
  function M_walk(sh, fn) { fn(sh); if (sh.type === 'group') for (const k of sh.kids) M_walk(k, fn); }
  function forEachTx(sh, fn) {
    if (sh.type === 'table') {
      const cs = E.cellSel;
      sh.tbl.rows.forEach((r, ri) => r.cells.forEach((c, ci) => { if (!cs || (ri >= Math.min(cs.r0, cs.r1) && ri <= Math.max(cs.r0, cs.r1) && ci >= Math.min(cs.c0, cs.c1) && ci <= Math.max(cs.c0, cs.c1))) fn(c.tx); }));
    } else if (sh.tx) fn(sh.tx);
  }
  F.forEachTx = forEachTx;
  /** Apply run transform. fn(run) -> run */
  F.run = function (fn, label, waFn) {
    if (TE.active() && !E.cellSel) { TE.applyRun(fn, label); return; }
    const shapes = TE.active() ? [TE.target().sh] : F.textShapes();
    if (!shapes.length) return;
    const editSnap = TE.active() ? TE.snapshot() : null;
    if (editSnap) TE.detach();
    L.hist.push(label || 'Format');
    for (const sh of shapes) {
      if (sh.type === 'wordart') { if (waFn) waFn(sh.wa); continue; }
      forEachTx(sh, (tx) => { T.mapRange(tx, T.allRange(tx), fn); for (const p of tx.ps) if (!p.rs.length) { p.end = fn(Object.assign({}, p.end || {})); delete p.end.t; } });
    }
    if (E.view === 'master') E.syncMaster();
    E.render({ keepEdit: false });
    if (editSnap) TE.resume(editSnap);
    E.touched();
  };
  F.para = function (fn, label) {
    if (TE.active() && !E.cellSel) { TE.applyPara(fn, label); return; }
    const shapes = TE.active() ? [TE.target().sh] : F.textShapes().filter((s) => s.type !== 'wordart');
    if (!shapes.length) return;
    const editSnap = TE.active() ? TE.snapshot() : null;
    if (editSnap) TE.detach();
    L.hist.push(label || 'Format Paragraph');
    for (const sh of shapes) forEachTx(sh, (tx) => tx.ps.forEach((p, i) => fn(p, i, { sh, tx })));
    if (E.view === 'master') E.syncMaster();
    E.render({ keepEdit: false });
    if (editSnap) TE.resume(editSnap);
    E.touched();
  };
  /** Effective run & paragraph state for toolbar display */
  F.state = function () {
    const design = E.design();
    let sh, tx, p, run;
    if (TE.active()) {
      const t = TE.target();
      if (!t) return null;
      sh = t.sh; tx = t.tx;
      const s = T.norm(TE.currentSel() || T.allRange(tx));
      p = tx.ps[s.p0] || tx.ps[0];
      run = TE.runPropsAtSel();
    } else {
      const shapes = F.textShapes();
      if (!shapes.length) return null;
      sh = shapes[0];
      if (sh.type === 'wordart') return { font: sh.wa.font || 'Arial Black', sz: null, b: !!sh.wa.b, i: !!sh.wa.i, u: false, shd: !!sh.wa.shadow, algn: sh.wa.algn || 'ctr', bu: 'none', color: null, wa: true, ps: {}, r: {}, lvl: 0 };
      if (sh.type === 'table') { const c = sh.tbl.rows[0].cells[0]; tx = c.tx; } else tx = sh.tx;
      if (!tx) return null;
      p = tx.ps[0];
      run = p.rs[0] ? T.runProps(p.rs[0]) : p.end || {};
    }
    if (!p) return null;
    const ps = L.style.para(sh, p, design);
    const r = Object.assign({}, ps.rPr || {}, run || {});
    return {
      font: L.style.font(r.font, design), sz: r.sz ? L.round(r.sz * (tx.fontScale || 1), 1) : 18, b: !!r.b, i: !!r.i, u: !!(r.u && r.u !== 'none'),
      shd: !!r.shd, algn: ps.algn || 'l', bu: ps.bu ? ps.bu.t : 'none', color: r.color, lvl: p.lvl || 0, ps, r,
    };
  };
})();
