/* Quire — model editing operations. All functions assume they run inside D.tx(). */
(function () {
  'use strict';
  const L = window.L, D = L.D;
  const O = (L.O = {});
  const doc = () => D.doc;

  /* ---------- revision marks ---------- */
  O.author = () => L.store.get('userName', 'Quire User');
  O.initials = () => L.store.get('userInitials', (O.author().match(/\b\w/g) || ['Q']).join('').slice(0, 3).toUpperCase());
  O.revStamp = () => ({ author: O.author(), date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), id: D.nid() });
  O.tracking = () => !!doc().settings.track;

  /* ---------- inserting text ---------- */
  /** formatting new text should get at pos when no explicit pending format exists */
  O.inheritRPr = function (pos) {
    const p = pos.p;
    let src = D.itemBefore(p, pos.o);
    let fromBefore = !!src;
    if (src && (src.t === 'fn' || src.t === 'en' || src.t === 'img' || src.t === 'shape' || src.t === 'group' || src.t === 'chart')) src = null;
    if (!src) { src = D.itemAfter(p, pos.o); fromBefore = false; }
    if (src && (src.t === 'img' || src.t === 'shape' || src.t === 'group' || src.t === 'fn' || src.t === 'en' || src.t === 'chart')) src = null;
    const r = L.clone(src ? src.rPr : p.rPr) || {};
    delete r.ins; delete r.del; delete r.chg;
    /* typing after a footnote/endnote reference or hyperlink end should not continue the reference style */
    if (r.style === 'FootnoteReference' || r.style === 'EndnoteReference' || r.style === 'CommentReference') { delete r.style; delete r.vert; }
    if (fromBefore && r.link) {
      /* only stay inside the link when there is more link text after the caret */
      const nx = D.itemAfter(p, pos.o);
      if (!nx || !nx.rPr || nx.rPr.link !== src.rPr.link && JSON.stringify(nx.rPr.link) !== JSON.stringify(src.rPr.link)) { delete r.link; if (r.style === 'Hyperlink') delete r.style; }
    }
    return r;
  };
  O.insertText = function (pos, text, rPr) {
    if (!text) return pos;
    const p = D.touch(pos.p);
    const r = rPr ? L.clone(rPr) : O.inheritRPr(pos);
    if (O.tracking()) r.ins = O.revStamp(); else delete r.ins;
    delete r.del;
    const i = D.insertIndex(p, pos.o);
    /* append into the previous text run when formatting matches (keeps runs compact) */
    const prev = p.runs[i - 1];
    if (prev && prev.t === 'text' && D.sameRPr(prev.rPr, r)) prev.text += text;
    else {
      const next = p.runs[i];
      if (next && next.t === 'text' && D.sameRPr(next.rPr, r)) next.text = text + next.text;
      else p.runs.splice(i, 0, D.text(text, r));
    }
    return D.pos(p, pos.o + text.length);
  };
  /** insert an inline item (tab, break, image, field markers...) */
  O.insertItem = function (pos, item) {
    const p = D.touch(pos.p);
    if (!item.rPr) item.rPr = O.inheritRPr(pos);
    if (O.tracking() && D.ilen(item)) item.rPr.ins = O.revStamp();
    let i;
    if (D.isMarker(item)) { i = D.splitAt(p, pos.o); while (i < p.runs.length && D.isMarker(p.runs[i])) i++; /* keep marker order when several land on one offset */ }
    else i = D.insertIndex(p, pos.o);
    p.runs.splice(i, 0, item);
    return D.pos(p, pos.o + D.ilen(item));
  };
  /** insert text that may contain \n (new paragraph), \t, \v (line break), \f (page break) */
  O.insertPlain = function (pos, str, rPr) {
    const parts = String(str).replace(/\r\n?/g, '\n').split('\n');
    parts.forEach((part, k) => {
      if (k > 0) pos = O.splitPara(pos);
      let buf = '';
      const flush = () => { if (buf) { pos = O.insertText(pos, buf, rPr); buf = ''; } };
      for (const ch of part) {
        if (ch === '\t') { flush(); pos = O.insertItem(pos, D.item('tab', null, rPr ? L.clone(rPr) : null)); }
        else if (ch === '\v' || ch === ' ') { flush(); pos = O.insertItem(pos, D.item('br', { type: 'line' }, rPr ? L.clone(rPr) : null)); }
        else if (ch === '\f') { flush(); pos = O.insertItem(pos, D.item('br', { type: 'page' }, rPr ? L.clone(rPr) : null)); }
        else buf += ch;
      }
      flush();
    });
    return pos;
  };

  /* ---------- paragraphs ---------- */
  O.containerOf = (p) => { const r = D.info(doc(), p); return r ? r.cont : null; };
  /** Split paragraph at pos; returns position at the start of the new (second) paragraph. */
  O.splitPara = function (pos, opts) {
    const d = doc();
    const p = D.touch(pos.p);
    const info = D.info(d, p);
    const cont = D.touchList(info.cont);
    const n = D.plen(p);
    /* end markers at the split point (field, comment, bookmark ends) stay with the first half */
    const i = D.insertIndex(p, pos.o);
    const tail = p.runs.splice(i);
    const np = D.para(tail, L.clone(p.pPr), L.clone(p.rPr));
    /* the paragraph mark (and with it a section break) belongs to the second half */
    if (p.sect) { np.sect = p.sect; delete p.sect; }
    if (p.mark) { np.mark = p.mark; delete p.mark; }
    const st = d.styles[p.pPr.style || 'Normal'];
    if (pos.o >= n && st && st.next && st.next !== st.id && !(opts && opts.keepStyle)) {
      np.pPr = { style: st.next };
      if (st.next === 'Normal') delete np.pPr.style;
      /* the new mark takes the mark formatting minus style-specific direct formatting */
      np.rPr = {};
    }
    if (pos.o === 0 && n > 0) {
      /* splitting at the start: the empty first part keeps the formatting, but page-break-before moves along */
      if (p.pPr.pageBreakBefore) { np.pPr.pageBreakBefore = true; delete p.pPr.pageBreakBefore; }
    }
    delete np.pPr.dropCap;
    if (O.tracking()) p.mark = { ins: O.revStamp() };
    cont.blocks.splice(cont.blocks.indexOf(p) + 1, 0, np);
    d._idxDirty = true;
    return D.pos(np, 0);
  };
  /** Join paragraph p with the paragraph that follows it in the same container. Returns the join position. */
  O.joinNext = function (p) {
    const d = doc();
    const info = D.info(d, p);
    const cont = info.cont;
    const idx = cont.blocks.indexOf(p);
    const q = cont.blocks[idx + 1];
    if (!q || q.t !== 'p') return null;
    D.touch(p); D.touch(q); D.touchList(cont);
    const o = D.plen(p);
    const pEmpty = o === 0 && !p.runs.length;
    if (pEmpty) { p.pPr = L.clone(q.pPr); p.rPr = L.clone(q.rPr); }
    p.runs = p.runs.concat(q.runs);
    if (q.sect) p.sect = q.sect; else delete p.sect;
    if (q.mark) p.mark = q.mark; else delete p.mark;
    q.runs = [];
    cont.blocks.splice(idx + 1, 1);
    D.normalize(p);
    d._idxDirty = true;
    return D.pos(p, o);
  };

  /* ---------- deleting ---------- */
  /** remove items in [a,b) of one paragraph (no tracking) */
  function cutRange(p, a, b) {
    if (a >= b) return;
    D.touch(p);
    const i0 = D.splitAt(p, a), i1 = D.splitAt(p, b);
    /* keep zero-length end markers (comment/bookmark/field ends) whose start is outside the range */
    const removed = p.runs.splice(i0, i1 - i0);
    const keep = removed.filter((it) => D.isMarker(it) && it.t !== 'fb' && it.t !== 'fs' && it.t !== 'fe');
    p.runs.splice(i0, 0, ...keep);
    D.normalize(p);
  }
  /** mark [a,b) as deleted (tracked). Inserted text by anyone is removed outright when deleted (as Word does for your own insertions). */
  function trackDelete(p, a, b) {
    if (a >= b) return;
    D.touch(p);
    const i0 = D.splitAt(p, a), i1 = D.splitAt(p, b);
    const stamp = O.revStamp();
    const out = [];
    for (let i = i0; i < i1; i++) {
      const it = p.runs[i];
      if (D.isMarker(it)) { out.push(it); continue; }
      if (it.rPr && it.rPr.ins && it.rPr.ins.author === stamp.author) continue; /* deleting own insertion */
      if (it.rPr && it.rPr.del) { out.push(it); continue; }
      const c = Object.assign({}, it, { rPr: Object.assign({}, it.rPr, { del: stamp }) });
      out.push(c);
    }
    p.runs.splice(i0, i1 - i0, ...out);
    D.normalize(p);
  }
  /**
   * Delete the range [a,b) (positions in one story). Returns the resulting caret position.
   * Paragraph marks inside the range are removed (paragraphs merge), whole tables inside are removed,
   * partially covered tables only have their covered cell contents cleared.
   */
  O.deleteRange = function (a, b) {
    const d = doc();
    [a, b] = D.order(d, a, b);
    if (D.eqPos(a, b)) return a;
    if (O.tracking()) return trackDeleteRange(a, b);
    if (a.p === b.p) { cutRange(a.p, a.o, b.o); return D.pos(a.p, a.o); }
    const ia = D.info(d, a.p), ib = D.info(d, b.p);
    if (ia.cont === ib.cont) {
      const cont = D.touchList(ia.cont);
      const i0 = cont.blocks.indexOf(a.p), i1 = cont.blocks.indexOf(b.p);
      cutRange(a.p, a.o, D.plen(a.p));
      cutRange(b.p, 0, b.o);
      /* blocks strictly between are removed entirely */
      const between = cont.blocks.slice(i0 + 1, i1);
      /* a section break in a removed paragraph is dropped (Word asks; we follow its default of merging sections) */
      cont.blocks.splice(i0 + 1, i1 - i0 - 1);
      void between;
      d._idxDirty = true;
      const jp = O.joinNext(a.p);
      return jp || D.pos(a.p, a.o);
    }
    /* cross-container (tables involved) */
    const paras = D.parasBetween(d, a.p, b.p);
    const top = D.topBlock(d, a.p).cont;
    const firstTop = D.topBlock(d, a.p), lastTop = D.topBlock(d, b.p);
    /* whole tables strictly inside the range at the story's top level are removed */
    const removeTbls = [];
    if (firstTop.cont === lastTop.cont) {
      for (let i = firstTop.i + 1; i < lastTop.i; i++) if (top.blocks[i].t === 'tbl') removeTbls.push(top.blocks[i]);
    }
    for (const p of paras) {
      const inRemoved = removeTbls.some((t) => D.info(d, p).parents.some((x) => x.tbl === t));
      if (inRemoved) continue;
      const s = p === a.p ? a.o : 0, e = p === b.p ? b.o : D.plen(p);
      cutRange(p, s, e);
    }
    /* paragraphs fully inside the range at the top level are removed, except the endpoints */
    if (firstTop.cont === lastTop.cont) {
      D.touchList(top);
      const keep = new Set([firstTop.block, lastTop.block]);
      top.blocks = top.blocks.filter((bl, i) => i <= firstTop.i || i >= lastTop.i || keep.has(bl) || (bl.t === 'tbl' && !removeTbls.includes(bl)));
      d._idxDirty = true;
      /* if both ends are paragraphs in the top container and became adjacent, join them */
      if (firstTop.block === a.p && lastTop.block === b.p) {
        const j = O.joinNext(a.p);
        if (j) return j;
      }
    }
    d._idxDirty = true;
    return D.pos(a.p, a.o);
  };
  function trackDeleteRange(a, b) {
    const d = doc();
    const paras = D.parasBetween(d, a.p, b.p);
    for (const p of paras) {
      const s = p === a.p ? a.o : 0, e = p === b.p ? b.o : D.plen(p);
      trackDelete(p, s, e);
      if (p !== b.p) {
        /* deleting the paragraph mark: tracked as a deleted mark unless it was inserted by the author */
        D.touch(p);
        if (p.mark && p.mark.ins && p.mark.ins.author === O.author()) { const q = D.nextPara(d, p); if (q && O.containerOf(q) === O.containerOf(p)) { delete p.mark; O.joinNext(p); } }
        else p.mark = { del: O.revStamp() };
      }
    }
    /* caret goes to the start (Word leaves it before the struck-through text when deleting forward, after when backspacing) */
    return D.pos(a.p, Math.min(a.o, D.plen(a.p)));
  }

  /* ---------- formatting ---------- */
  /** apply fn(rPr) → rPr to every item in [a,b) */
  O.formatRange = function (a, b, fn) {
    const d = doc();
    [a, b] = D.order(d, a, b);
    const paras = D.parasBetween(d, a.p, b.p);
    for (const p of paras) {
      const s = p === a.p ? a.o : 0, e = p === b.p ? b.o : D.plen(p);
      D.touch(p);
      if (s < e) {
        const i0 = D.splitAt(p, s), i1 = D.splitAt(p, e);
        for (let i = i0; i < i1; i++) {
          const it = p.runs[i];
          if (D.isMarker(it)) continue;
          const old = it.rPr || {};
          const nr = fn(Object.assign({}, old), it) || {};
          if (O.tracking() && !old.chg && !old.ins && !D.sameRPr(old, nr)) nr.chg = Object.assign(O.revStamp(), { old: L.clone(old) });
          it.rPr = nr;
        }
        D.normalize(p);
      }
      /* when a whole paragraph is covered (or the selection crosses its mark), the paragraph mark gets the formatting too */
      if ((p !== b.p) || (s === 0 && e === D.plen(p) && e > 0)) p.rPr = fn(Object.assign({}, p.rPr), null) || {};
    }
  };
  /** set or clear run properties. props: {key: value|undefined} */
  O.setRunProps = (a, b, props) => O.formatRange(a, b, (r) => { for (const k in props) { if (props[k] === undefined || props[k] === null) delete r[k]; else r[k] = props[k]; } return r; });
  O.setParaProps = function (paras, fn) {
    for (const p of paras) {
      D.touch(p);
      const before = L.clone(p.pPr);
      const r = fn(p.pPr, p);
      if (r) p.pPr = r;
      if (O.tracking() && !p.pPr.chg && JSON.stringify(before) !== JSON.stringify(p.pPr)) p.pPr.chg = Object.assign(O.revStamp(), { old: before });
    }
  };
  /** current resolved run formatting at a position (what the toolbar shows) */
  O.rPrAt = function (pos) {
    const d = doc();
    const it = D.itemBefore(pos.p, pos.o) || D.itemAfter(pos.p, pos.o);
    const ctx = L.layout && L.layout.tblCtx ? L.layout.tblCtx(pos.p) : null;
    return D.rProps(d, pos.p, it && !D.isMarker(it) ? it : null, ctx);
  };

  /* ---------- copy / structure ---------- */
  /** deep copy of the range as blocks (first/last paragraph partial) */
  O.copyRange = function (a, b) {
    const d = doc();
    [a, b] = D.order(d, a, b);
    if (a.p === b.p) {
      const p = D.para(D.sliceRuns(a.p, a.o, b.o), L.clone(a.p.pPr), L.clone(a.p.rPr));
      p.partial = true;
      return [p];
    }
    const ia = D.info(d, a.p), ib = D.info(d, b.p);
    const out = [];
    if (ia.cont === ib.cont) {
      const bl = ia.cont.blocks;
      const i0 = bl.indexOf(a.p), i1 = bl.indexOf(b.p);
      for (let i = i0; i <= i1; i++) {
        const blk = bl[i];
        if (blk === a.p) out.push(D.para(D.sliceRuns(blk, a.o, D.plen(blk)), L.clone(blk.pPr), L.clone(blk.rPr)));
        else if (blk === b.p) { const q = D.para(D.sliceRuns(blk, 0, b.o), L.clone(blk.pPr), L.clone(blk.rPr)); q.partial = true; out.push(q); }
        else out.push(D.cloneBlocks([blk])[0]);
        if (blk.t === 'p' && blk.sect && blk !== b.p) out[out.length - 1].sect = L.clone(blk.sect);
      }
      return out.map((x) => { if (x.t === 'p') x.id = D.nid(); return x; });
    }
    /* across containers: copy top-level blocks spanning the selection */
    const fa = D.topBlock(d, a.p), fb = D.topBlock(d, b.p);
    if (fa.cont === fb.cont) {
      for (let i = fa.i; i <= fb.i; i++) {
        const blk = fa.cont.blocks[i];
        if (blk === a.p) out.push(D.para(D.sliceRuns(blk, a.o, D.plen(blk)), L.clone(blk.pPr), L.clone(blk.rPr)));
        else if (blk === b.p) { const q = D.para(D.sliceRuns(blk, 0, b.o), L.clone(blk.pPr), L.clone(blk.rPr)); q.partial = true; out.push(q); }
        else out.push(D.cloneBlocks([blk])[0]);
      }
      return out;
    }
    for (const p of D.parasBetween(d, a.p, b.p)) out.push(D.para(D.sliceRuns(p, p === a.p ? a.o : 0, p === b.p ? b.o : D.plen(p)), L.clone(p.pPr), L.clone(p.rPr)));
    return out;
  };
  /** plain text of a range (paragraphs separated by \r\n for clipboard) */
  O.textRange = function (a, b, sep) {
    const d = doc();
    [a, b] = D.order(d, a, b);
    const paras = D.parasBetween(d, a.p, b.p);
    const out = [];
    for (const p of paras) {
      const s = p === a.p ? a.o : 0, e = p === b.p ? b.o : D.plen(p);
      let t = '';
      let pos = 0;
      for (const it of p.runs) {
        const n = D.ilen(it);
        const s0 = pos, e0 = pos + n;
        pos = e0;
        if (!n || e0 <= s || s0 >= e) continue;
        if (it.rPr && it.rPr.del) continue;
        if (it.t === 'text') t += it.text.slice(Math.max(0, s - s0), Math.min(n, e - s0));
        else if (it.t === 'tab') t += '\t';
        else if (it.t === 'br') t += it.type === 'page' ? '\f' : '\v';
        else if (it.t === 'sym') t += it.char || '';
        else if (it.t === 'fn' || it.t === 'en') t += it.mark || '';
      }
      out.push(t.replace(/\v/g, sep === '\n' ? '\n' : '\v'));
    }
    return out.join(sep == null ? '\r\n' : sep);
  };
  /**
   * Insert blocks at pos. The first block (if a paragraph) merges into the paragraph at pos;
   * the last one (if paragraph and "partial") merges with the remainder. Returns the end position.
   */
  O.insertBlocks = function (pos, blocks, opts) {
    const d = doc();
    if (!blocks.length) return pos;
    blocks = D.cloneBlocks(blocks);
    const trk = O.tracking();
    if (trk) {
      const st = O.revStamp();
      const mark = (bl) => { for (const b of bl) { if (b.t === 'p') { for (const it of b.runs) if (D.ilen(it)) it.rPr = Object.assign({}, it.rPr, { ins: st }); } else if (b.t === 'tbl') for (const r of b.rows) for (const c of r.cells) mark(c.blocks); } };
      mark(blocks);
    }
    const keepSrcFmt = !opts || opts.keepFormat !== false;
    if (blocks.length === 1 && blocks[0].t === 'p') {
      const src = blocks[0];
      const p = D.touch(pos.p);
      const i = D.splitAt(p, pos.o);
      const items = src.runs.map((it) => Object.assign({}, it));
      if (!keepSrcFmt) for (const it of items) it.rPr = Object.assign(O.inheritRPr(pos), it.rPr && it.rPr.ins ? { ins: it.rPr.ins } : {});
      p.runs.splice(i, 0, ...items);
      const n = items.reduce((s, it) => s + D.ilen(it), 0);
      D.normalize(p);
      /* a complete pasted paragraph (not partial) brings its formatting when the target is empty */
      if (!src.partial && !D.plen(p) - n && keepSrcFmt) p.pPr = Object.assign({}, src.pPr);
      return D.pos(p, pos.o + n);
    }
    /* multi-block: split target, insert between */
    const startP = pos.p;
    const wasEmpty = D.plen(startP) === 0;
    const after = O.splitPara(pos, { keepStyle: true });
    const info = D.info(d, startP);
    const cont = D.touchList(info.cont);
    let insertAt = cont.blocks.indexOf(startP) + 1;
    let first = 0;
    if (blocks[0].t === 'p') {
      const p = D.touch(startP);
      p.runs = p.runs.concat(blocks[0].runs);
      if (wasEmpty || pos.o === 0) { p.pPr = L.clone(blocks[0].pPr); p.rPr = L.clone(blocks[0].rPr); }
      if (blocks[0].sect) p.sect = blocks[0].sect;
      D.normalize(p);
      first = 1;
    }
    let lastIdx = blocks.length;
    let endPos = null;
    const last = blocks[blocks.length - 1];
    if (blocks.length > 1 && last.t === 'p' && last.partial) {
      const q = D.touch(after.p);
      const n = last.runs.reduce((s, it) => s + D.ilen(it), 0);
      q.runs = last.runs.concat(q.runs);
      if (!D.plen(q) - n) { /* keep target paragraph's formatting otherwise */ }
      D.normalize(q);
      endPos = D.pos(q, n);
      lastIdx = blocks.length - 1;
    }
    const mid = blocks.slice(first, lastIdx);
    for (const b of mid) delete b.partial;
    cont.blocks.splice(insertAt, 0, ...mid);
    d._idxDirty = true;
    if (!endPos) endPos = D.pos(after.p, 0);
    /* an empty trailing paragraph created by the split is removed when the paste ended with a full paragraph */
    if (!endPos.o && !D.plen(after.p) && mid.length && mid[mid.length - 1].t === 'p' && after.p !== startP && !after.p.sect) {
      const ai = cont.blocks.indexOf(after.p);
      const prevB = cont.blocks[ai - 1];
      if (ai > 0 && prevB.t === 'p' && cont.blocks.length > 1 && cont.blocks[ai + 1]) {
        cont.blocks.splice(ai, 1);
        d._idxDirty = true;
        return D.pos(prevB, D.plen(prevB));
      }
    }
    return endPos;
  };
  /** insert a block (table, etc.) at pos: splits the paragraph if needed. Returns first paragraph inside the inserted block. */
  O.insertBlockAt = function (pos, block) {
    const d = doc();
    const p = pos.p;
    const n = D.plen(p);
    const info = D.info(d, p);
    const cont = D.touchList(info.cont);
    let idx;
    if (pos.o === 0 && n > 0) idx = cont.blocks.indexOf(p);
    else if (pos.o === 0 && n === 0) idx = cont.blocks.indexOf(p);
    else if (pos.o >= n) idx = cont.blocks.indexOf(p) + 1;
    else { const np = O.splitPara(pos, { keepStyle: true }); idx = cont.blocks.indexOf(np.p); }
    cont.blocks.splice(idx, 0, block);
    /* a table must be followed by a paragraph */
    if (!cont.blocks[idx + 1]) cont.blocks.push(D.para([], {}, {}));
    d._idxDirty = true;
    return block;
  };
  /** remove a block from its container, keeping at least one paragraph */
  O.removeBlock = function (block) {
    const d = doc();
    let cont = null;
    if (block.t === 'p') cont = D.info(d, block).cont;
    else { const ti = D.tblInfo(d, block); cont = ti && ti.cont; }
    if (!cont) return null;
    D.touchList(cont);
    const i = cont.blocks.indexOf(block);
    cont.blocks.splice(i, 1);
    if (!cont.blocks.length) cont.blocks.push(D.para());
    d._idxDirty = true;
    const nb = cont.blocks[Math.min(i, cont.blocks.length - 1)];
    const p = nb.t === 'p' ? nb : D.firstPara([nb]);
    return D.pos(p, 0);
  };

  /* ---------- fields ---------- */
  /** insert a field: begin marker, result text, end marker */
  O.insertField = function (pos, instr, resultText, rPr) {
    const fid = D.newFid();
    const base = rPr ? L.clone(rPr) : O.inheritRPr(pos);
    delete base.ins; delete base.del;
    const p = D.touch(pos.p);
    const mk = (t, extra) => D.item(t, Object.assign({ fid }, extra || {}), L.clone(base));
    const items = [mk('fb', { instr: ' ' + String(instr).trim() + ' ' }), mk('fs')];
    if (resultText) { const r = L.clone(base); if (O.tracking()) r.ins = O.revStamp(); items.push(D.text(String(resultText), r)); }
    items.push(mk('fe'));
    /* the markers go in as one ordered run (inserting them one by one at the same offset would reverse them) */
    p.runs.splice(D.insertIndex(p, pos.o), 0, ...items);
    return D.pos(p, pos.o + (resultText ? String(resultText).length : 0));
  };
  /** locate field parts within a story: returns {begin:{p,i,o}, sep, end} for fid */
  O.findField = function (story, fid) {
    const out = {};
    for (const p of story.paras || []) {
      let o = 0;
      for (let i = 0; i < p.runs.length; i++) {
        const it = p.runs[i];
        if (it.fid === fid) {
          if (it.t === 'fb') out.begin = { p, i, o, it };
          else if (it.t === 'fs') out.sep = { p, i, o, it };
          else if (it.t === 'fe') out.end = { p, i, o, it };
        }
        o += D.ilen(it);
      }
    }
    return out;
  };
  /** replace a field's result with new blocks or text */
  O.setFieldResult = function (story, fid, result) {
    const d = doc();
    const f = O.findField(story, fid);
    if (!f.begin || !f.end) return;
    if (!f.sep) {
      /* field without separator: add one right before the end */
      D.touch(f.end.p);
      f.end.p.runs.splice(f.end.i, 0, D.item('fs', { fid }, L.clone(f.end.it.rPr)));
      return O.setFieldResult(story, fid, result);
    }
    const a = D.pos(f.sep.p, f.sep.o), b = D.pos(f.end.p, f.end.o);
    /* remove old result */
    const trk = d.settings.track; d.settings.track = false;
    try {
      if (f.sep.p === f.end.p) {
        /* single-paragraph result: drop the items between the markers, leaving both markers intact */
        if (f.end.i > f.sep.i + 1) { D.touch(f.sep.p); f.sep.p.runs.splice(f.sep.i + 1, f.end.i - f.sep.i - 1); }
      } else if (!D.eqPos(a, b)) O.deleteRange(a, b);
      let g = O.findField(D.storyOf(d, f.sep.p) || story, fid);
      if (g.begin && g.end && !g.sep) { D.touch(g.end.p); g.end.p.runs.splice(g.end.i, 0, D.item('fs', { fid }, L.clone(g.end.it.rPr))); g = O.findField(D.storyOf(d, g.end.p) || story, fid); }
      if (!g.sep) return;
      const at = D.pos(g.sep.p, g.sep.o);
      if (typeof result === 'string') {
        const rp = L.clone(g.begin.it.rPr || {});
        const ins = D.splitAt(at.p, at.o);
        void ins;
        if (result) { D.touch(at.p); const i = at.p.runs.indexOf(g.sep.it) + 1; at.p.runs.splice(i, 0, D.text(result, rp)); D.normalize(at.p); }
      } else if (Array.isArray(result)) {
        O.insertBlocks(at, result, { keepFormat: true });
      }
    } finally { d.settings.track = trk; }
  };
  /** remove field markers whose partners are missing (after deletions) */
  O.repairFields = function (story) {
    const parts = new Map();
    for (const p of story.paras || []) for (const it of p.runs) if (it.fid && (it.t === 'fb' || it.t === 'fs' || it.t === 'fe')) { const s = parts.get(it.fid) || {}; s[it.t] = true; parts.set(it.fid, s); }
    const bad = new Set();
    for (const [fid, s] of parts) if (!s.fb || !s.fe) bad.add(fid);
    if (!bad.size) return;
    for (const p of story.paras || []) if (p.runs.some((it) => bad.has(it.fid))) { D.touch(p); p.runs = p.runs.filter((it) => !bad.has(it.fid)); }
  };

  /* ---------- case ---------- */
  O.changeCase = function (a, b, mode) {
    const d = doc();
    [a, b] = D.order(d, a, b);
    const all = O.textRange(a, b, '\n');
    let atSentence = true;
    void all;
    for (const p of D.parasBetween(d, a.p, b.p)) {
      const s = p === a.p ? a.o : 0, e = p === b.p ? b.o : D.plen(p);
      if (s >= e) continue;
      D.touch(p);
      const i0 = D.splitAt(p, s), i1 = D.splitAt(p, e);
      let prevCh = s > 0 ? D.ptext(p, s - 1, s) : ' ';
      for (let i = i0; i < i1; i++) {
        const it = p.runs[i];
        if (it.t !== 'text') { prevCh = ' '; continue; }
        let out = '';
        for (const ch of it.text) {
          let c = ch;
          if (mode === 'lower') c = ch.toLowerCase();
          else if (mode === 'upper') c = ch.toUpperCase();
          else if (mode === 'toggle') c = ch === ch.toUpperCase() ? ch.toLowerCase() : ch.toUpperCase();
          else if (mode === 'title') c = /[\s\-("“'‘\/]/.test(prevCh) || prevCh === '' ? ch.toUpperCase() : ch.toLowerCase();
          else if (mode === 'sentence') { c = atSentence && /\p{L}/u.test(ch) ? ch.toUpperCase() : ch.toLowerCase(); if (/\p{L}/u.test(ch)) atSentence = false; if (/[.!?]/.test(ch)) atSentence = true; }
          out += c;
          prevCh = ch;
        }
        it.text = out;
      }
      if (mode === 'sentence') atSentence = true;
    }
  };

  /* ---------- styles ---------- */
  O.applyStyle = function (a, b, styleId) {
    const d = doc();
    const st = d.styles[styleId];
    if (!st) return;
    [a, b] = D.order(d, a, b);
    if (st.type === 'character') {
      if (D.eqPos(a, b)) {
        /* whole word under caret */
        const w = O.wordAt(a);
        if (w) { a = w[0]; b = w[1]; } else return;
      }
      O.formatRange(a, b, (r) => { if (styleId === 'DefaultParagraphFont') delete r.style; else r.style = styleId; return r; });
      return;
    }
    if (st.type === 'paragraph') {
      const paras = D.parasBetween(d, a.p, b.p);
      O.setParaProps(paras, (pPr) => {
        const keep = {};
        /* applying a paragraph style clears direct numbering if the style defines its own */
        for (const k of ['sect', 'frame']) if (pPr[k]) keep[k] = pPr[k];
        const np = Object.assign({}, keep);
        if (styleId !== 'Normal') np.style = styleId;
        /* direct list membership survives unless the new style defines numbering or is a heading */
        if (pPr.num && !(D.styleProps(d, styleId).pPr.num)) np.num = pPr.num;
        return np;
      });
      /* Word removes direct character formatting that covers the whole paragraph when applying a style */
      for (const p of paras) {
        D.touch(p);
        const items = p.runs.filter((it) => D.ilen(it));
        if (items.length) {
          for (const k of ['b', 'i', 'font', 'sz', 'color', 'u', 'caps', 'smallCaps']) {
            const v0 = items[0].rPr ? JSON.stringify(items[0].rPr[k]) : undefined;
            if (v0 !== undefined && items.every((it) => it.rPr && JSON.stringify(it.rPr[k]) === v0)) for (const it of items) delete it.rPr[k];
          }
          D.normalize(p);
        }
      }
    }
  };
  /** word boundaries around pos: [start,end] positions or null */
  O.wordAt = function (pos) {
    const t = D.ptext(pos.p);
    const isW = (c) => /[\p{L}\p{N}_'’]/u.test(c);
    let s = pos.o, e = pos.o;
    while (s > 0 && isW(t[s - 1])) s--;
    while (e < t.length && isW(t[e])) e++;
    if (s === e) return null;
    return [D.pos(pos.p, s), D.pos(pos.p, e)];
  };
  /** word selection extent as Word does on double-click (includes trailing spaces) */
  O.wordExtent = function (pos) {
    const t = D.ptext(pos.p);
    if (!t.length) return [pos, pos];
    const cls = (c) => (c == null ? -1 : /[\p{L}\p{N}_'’]/u.test(c) ? 1 : /\s/.test(c) ? 0 : 2);
    let o = Math.min(pos.o, t.length - 1);
    if (o > 0 && cls(t[o]) === 0 && cls(t[o - 1]) === 1) o--;
    const k = cls(t[o]);
    let s = o, e = o + 1;
    while (s > 0 && cls(t[s - 1]) === k && k !== 2) s--;
    while (e < t.length && cls(t[e]) === k && k !== 2) e++;
    while (e < t.length && t[e] === ' ') e++;
    return [D.pos(pos.p, s), D.pos(pos.p, e)];
  };
  O.sentenceExtent = function (pos) {
    const t = D.ptext(pos.p);
    let s = pos.o, e = pos.o;
    while (s > 0 && !/[.!?]\s/.test(t.slice(s - 2, s))) s--;
    while (e < t.length && !/[.!?]/.test(t[e])) e++;
    if (e < t.length) e++;
    while (e < t.length && /\s/.test(t[e])) e++;
    return [D.pos(pos.p, s), D.pos(pos.p, e)];
  };
})();
