/* Quire — layout: pagination (Print Layout), flow views (Normal, Web, Outline), headers/footers,
 * columns, footnotes/endnotes, table row splitting, widow/orphan and keep rules, incremental relayout,
 * and the mapping between DOM positions and model positions.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, R = L.R;
  const LY = (L.layout = {});
  const PX = D.PX;
  const px = R.px;
  LY.view = 'print';
  LY.zoom = 1;
  LY.opts = { showMarks: false, showHidden: false, showBookmarks: false, fieldCodes: false, markup: 'markup', showComments: true, textBoundaries: true, showFmtChanges: false, balloons: true };
  LY.pages = [];
  LY.root = null; /* contentEditable host */
  LY.sizer = null;
  LY.scroller = null;
  LY.hfEdit = null; /* {story, pageIdx} while editing headers/footers */

  const el = (tag, cls, css) => { const e = document.createElement(tag); if (cls) e.className = cls; if (css) e.style.cssText = css; return e; };
  const doc = () => D.doc;

  LY.mount = function (scroller) {
    LY.scroller = scroller;
    LY.sizer = el('div', 'sizer');
    LY.root = el('div', 'pages');
    LY.root.id = 'pages';
    LY.root.contentEditable = 'true';
    LY.root.spellcheck = false;
    LY.root.setAttribute('role', 'textbox');
    LY.root.setAttribute('aria-multiline', 'true');
    LY.root.setAttribute('aria-label', 'Document');
    LY.sizer.appendChild(LY.root);
    scroller.appendChild(LY.sizer);
  };

  /* ================= shared render context ================= */
  LY.ctx = function (extra) {
    const d = doc();
    const c = {
      doc: d, labels: LY.labels, view: LY.view, showMarks: LY.opts.showMarks, showHidden: LY.opts.showHidden || LY.opts.showMarks,
      showBookmarks: LY.opts.showBookmarks, fieldCodes: LY.opts.fieldCodes, markup: LY.opts.markup, showComments: LY.opts.showComments && LY.opts.markup !== 'final' && LY.opts.markup !== 'original',
      showFmtChanges: LY.opts.showFmtChanges, noteNum: LY.noteNum, openCmts: null,
    };
    return Object.assign(c, extra || {});
  };
  /** table style context for a paragraph (rPr/pPr contributed by its table's style) */
  LY.tblCtx = function (p) {
    const d = doc();
    const ci = D.cellOf(d, p);
    if (!ci) return null;
    const { tp, sp } = R.tblProps(d, ci.tbl);
    void tp;
    const nCols = ci.tbl.grid.length || 1;
    const map = R.tblMap(ci.tbl);
    const m = map[ci.ri].find((x) => x.cell === ci.cell);
    const cs = R.cellStyle(d, ci.tbl, sp, ci.ri, m ? m.c0 : ci.ci, ci.tbl.rows.length, nCols);
    return { tblRPr: D.mergeRPr(sp.rPr, cs.rPr), tblPPr: D.mergePPr(sp.pPr, cs.pPr) };
  };

  /* ================= notes numbering ================= */
  LY.computeNotes = function () {
    const d = doc();
    const nums = new Map();
    let fn = (d.settings.fnPr && d.settings.fnPr.start) || 1, en = (d.settings.enPr && d.settings.enPr.start) || 1;
    D.walk(d.main, (b) => {
      if (b.t !== 'p') return;
      for (const it of b.runs) {
        if (it.t === 'fn' && !it.custom && !nums.has('fn:' + it.id)) nums.set('fn:' + it.id, fn++);
        if (it.t === 'en' && !it.custom && !nums.has('en:' + it.id)) nums.set('en:' + it.id, en++);
      }
    });
    LY.noteNums = nums;
  };
  LY.noteNum = function (it) {
    const d = doc();
    const k = it.t + ':' + it.id;
    const n = LY.noteNums ? LY.noteNums.get(k) : 1;
    if (it.custom) return it.custom;
    const fmt = it.t === 'fn' ? (d.settings.fnPr && d.settings.fnPr.fmt) || 'decimal' : (d.settings.enPr && d.settings.enPr.fmt) || 'lowerRoman';
    return D.fmtNum(n || 1, fmt);
  };
  /** open comment ranges at the start of every paragraph (for highlighting) */
  function computeOpenComments(story) {
    const map = new Map();
    const open = new Set();
    const fields = [];
    for (const p of story.paras || []) {
      if (open.size) map.set(p.id, new Set(open));
      if (fields.length) map.set('f' + p.id, fields.slice());
      for (const it of p.runs) {
        if (it.t === 'cs') open.add(it.id); else if (it.t === 'ce') open.delete(it.id);
        else if (it.t === 'fb') fields.push(it.fid); else if (it.t === 'fs' || it.t === 'fe') { const i = fields.lastIndexOf(it.fid); if (i >= 0) fields.splice(i, 1); }
      }
    }
    return map;
  }
  LY.openFor = (p) => (LY.openMap ? LY.openMap.get(p.id) || null : null);

  /* ================= measurement helpers ================= */
  LY.scale = () => LY.zoom || 1;
  /** rendered lines of a paragraph fragment: [{o, top, bottom}] relative to frameTop (client px), in layout px */
  LY.lines = function (frag, frameTop, scale) {
    const pc = frag.firstChild;
    const lines = [];
    let prev = null; /* previous rect */
    let forceNew = false;
    const lhCache = new Map();
    const lhOf = (span) => {
      let v = lhCache.get(span);
      if (v == null) { const cs = getComputedStyle(span); v = parseFloat(cs.lineHeight); if (isNaN(v)) v = parseFloat(cs.fontSize) * 1.2; lhCache.set(span, v); }
      return v;
    };
    const add = (o, rect, lh) => {
      const top = (rect.top - frameTop) / scale, bottom = (rect.bottom - frameTop) / scale;
      const half = lh != null ? Math.max(0, (lh - rect.height / scale) / 2) : 0;
      const isNew = !prev || forceNew || (rect.left < prev.left - 1 && rect.top > prev.top + 0.5) || rect.top >= prev.bottom - 0.5 && rect.left <= prev.left;
      if (isNew) lines.push({ o, top: top - half, bottom: bottom + half });
      else { const ln = lines[lines.length - 1]; ln.bottom = Math.max(ln.bottom, bottom + half); ln.top = Math.min(ln.top, top - half); }
      prev = rect;
      forceNew = false;
    };
    const range = document.createRange();
    const charRect = (tn, i) => { range.setStart(tn, i); range.setEnd(tn, i + 1); const rs = range.getClientRects(); return rs.length ? rs[0] : null; };
    for (const node of pc.childNodes) {
      if (node.nodeType !== 1) continue;
      if (node.nodeName === 'BR') { forceNew = true; continue; }
      if (node.classList.contains('lbl') || node.classList.contains('lbt') || node.classList.contains('lsp')) {
        const r = node.getBoundingClientRect();
        if (r.width || r.height) add(+frag.dataset.from, r, null);
        continue;
      }
      if (node.dataset.o == null) continue;
      const o0 = +node.dataset.o;
      const tn = node.firstChild;
      if (tn && tn.nodeType === 3 && node.contentEditable !== 'false') {
        const lh = lhOf(node);
        const text = tn.data;
        /* sample at break opportunities, refine with binary search between samples */
        let lastI = -1, lastR = null;
        const visit = (i) => {
          const r = charRect(tn, i);
          if (!r || (!r.width && !r.height)) return;
          if (lastR && prev && (r.left < prev.left - 1 && r.top > prev.top + 0.5) && i - lastI > 1) {
            /* a line break happened between lastI and i: find the first char on the new line */
            let lo = lastI + 1, hi = i;
            while (lo < hi) {
              const mid = (lo + hi) >> 1;
              const mr = charRect(tn, mid);
              if (mr && mr.top > lastR.top + 0.5 && mr.left < lastR.left + 1) hi = mid; else lo = mid + 1;
            }
            if (lo < i) { const lr = charRect(tn, lo); if (lr) add(o0 + lo, lr, lh); }
          }
          add(o0 + i, r, lh);
          lastI = i; lastR = r;
        };
        for (let i = 0; i < text.length; i++) {
          const c = text.charCodeAt(i);
          const pc0 = i ? text.charCodeAt(i - 1) : 32;
          if (i === 0 || pc0 === 32 || pc0 === 45 || pc0 === 47 || pc0 === 0x2014 || pc0 === 0x2013 || c === 32 && false || (i - lastI) >= 24 || i === text.length - 1 || c > 0x2E80) visit(i);
        }
      } else {
        const r = node.getBoundingClientRect();
        if (r.width || r.height) add(o0, r, null);
      }
    }
    return lines;
  };
  /** remove a laid-out fragment together with the page-level floats it anchored */
  const dropFrag = (f) => { if (f._floats) { for (const w of f._floats) w.remove(); f._floats = null; } f.remove(); };
  /** lowest edge of the text-wrapped floats inside a fragment (they hang outside its box) */
  const floatBottom = (frag, frameTop, scale) => {
    let b = 0;
    for (const w of frag.querySelectorAll(':scope > .pc > .flt')) b = Math.max(b, (w.getBoundingClientRect().bottom - frameTop) / scale);
    return b;
  };
  /* a break that ends its paragraph keeps the paragraph mark with it (Word): nothing carries over */
  const afterBreak = (st, segEnd, n) => (segEnd >= n ? { bi: st.bi + 1, sub: null } : { bi: st.bi, sub: segEnd });
  /** bottom (layout px, relative to frameTop) of an element */
  const bottomOf = (e, frameTop, scale) => (e.getBoundingClientRect().bottom - frameTop) / scale;

  /* ================= flowing blocks into a frame ================= */
  /**
   * frame: {el, limit (px), top (client px of el), empty, natural (started by overflow), page, notes}
   * st: {bi, sub}; blocks: array; end: exclusive index
   * returns {st, stop: 'done'|'full'|'page'|'column'}
   */
  LY.flow = function (frame, blocks, st, end, ctx) {
    const scale = LY.scale();
    const d = ctx.doc;
    frame.chain = frame.chain || [];
    while (st.bi < end) {
      const b = blocks[st.bi];
      if (b.t === 'tbl') {
        /* "page break before" on the first paragraph of the table's first cell breaks before the table (Word) */
        if (!st.sub && !frame.empty && frame.kind === 'col' && !ctx.inCell) {
          const c0 = b.rows[0] && b.rows[0].cells[0];
          const p0 = c0 && c0.blocks && c0.blocks[0];
          if (p0 && p0.t === 'p') {
            const { sp } = R.tblProps(d, b);
            const cs = R.cellStyle(d, b, sp, 0, 0, b.rows.length, b.grid.length || 1);
            if (D.pProps(d, p0, { tblRPr: D.mergeRPr(sp.rPr, cs.rPr), tblPPr: D.mergePPr(sp.pPr, cs.pPr) }).pageBreakBefore) return { st, stop: 'page', pbb: true };
          }
        }
        const res = placeTable(frame, b, st.sub, ctx);
        if (res.placed) { frame.empty = false; frame.chain = []; }
        if (res.done) { st = { bi: st.bi + 1, sub: null }; continue; }
        if (!res.placed && !frame.empty) return keepChainBreak(frame, { bi: st.bi, sub: res.sub }, ctx);
        return { st: { bi: st.bi, sub: res.sub }, stop: 'full' };
      }
      if (b.t !== 'p') { st = { bi: st.bi + 1, sub: null }; continue; }
      const p = b;
      const n = D.plen(p);
      const o = st.sub || 0;
      const pPr = D.pProps(d, p, ctx.tbl || null);
      /* (Word ignores "page break before" on an empty paragraph that only carries a section break) */
      if (o === 0 && pPr.pageBreakBefore && !frame.empty && frame.kind === 'col' && !ctx.inCell && !(p.sect && !n)) return { st, stop: 'page', pbb: true };
      /* frames: consecutive paragraphs sharing frame properties form one positioned box */
      if (o === 0 && pPr.frame && !pPr.dropCap && frame.kind !== 'flow' && !ctx.inTextBox) {
        const res = placeFrame(frame, blocks, st, end, ctx, pPr.frame);
        if (res.placed) { frame.empty = false; frame.chain = []; st = { bi: res.next, sub: null }; continue; }
        return { st, stop: 'full' };
      }
      /* hard break inside paragraph? */
      let brk = -1, brkType = null;
      if (frame.kind === 'col' || frame.kind === 'flow') {
        let pos = 0;
        for (const it of p.runs) {
          const ln = D.ilen(it);
          if (it.t === 'br' && (it.type === 'page' || it.type === 'column') && pos >= o && !(it.rPr && it.rPr.del && LY.opts.markup === 'final')) { brk = pos; brkType = it.type; break; }
          pos += ln;
        }
      }
      if (frame.kind === 'flow') brk = -1;
      const segEnd = brk >= 0 ? brk + 1 : n;
      const prevP = st.bi > 0 && blocks[st.bi - 1].t === 'p' ? blocks[st.bi - 1] : null;
      const nextP = st.bi + 1 < blocks.length && blocks[st.bi + 1].t === 'p' ? blocks[st.bi + 1] : null;
      const fctx = Object.assign({}, ctx, { prev: prevP, next: nextP, atTop: frame.empty && frame.natural && o === 0, frameEmpty: frame.empty && !ctx.inCell, openCmts: LY.openFor(p), fieldStack: LY.openMap ? LY.openMap.get('f' + p.id) : null, firstInCell: st.bi === 0, lastInCell: st.bi === blocks.length - 1, prevTbl: st.bi > 0 && blocks[st.bi - 1].t === 'tbl' });
      let frag = R.para(p, o, segEnd, fctx);
      frame.el.appendChild(frag);
      postPara(frag, p, fctx, frame);
      /* an unlimited frame (a table cell being filled) needs no measuring: skip the forced reflow */
      let bottom = frame.limit === Infinity && !frame.notes ? 0 : bottomOf(frag, frame.top, scale);
      /* footnotes referenced in this fragment */
      let noteH = 0;
      if (frame.notes) noteH = frame.notes.measure(frag);
      const limit = frame.limit - noteH;
      /* a wrapped picture anchored here that would run off the page takes its paragraph to the next page (Word) */
      if (frame.kind === 'col' && !frame.empty && !ctx.inCell && frame.limit !== Infinity && frag.querySelector(':scope > .pc > .flt') && floatBottom(frag, frame.top, scale) > limit + 0.5) {
        dropFrag(frag);
        if (frame.notes) frame.notes.rollback();
        return keepChainBreak(frame, st, ctx);
      }
      if (bottom <= limit + 0.5 || frame.limit === Infinity) {
        if (frame.notes) frame.notes.commit(frag);
        frame.empty = false;
        if (segEnd < n || (brk >= 0)) {
          frame.chain = [];
          return { st: afterBreak(st, segEnd, n), stop: brkType === 'column' ? 'column' : 'page' };
        }
        if (pPr.keepNext && frame.kind === 'col' && !ctx.inCell) frame.chain.push({ bi: st.bi, el: frag, sub: o }); else frame.chain = [];
        st = { bi: st.bi + 1, sub: null };
        if (p.sect && frame.kind === 'col' && st.bi < end) { /* section break mid-flow is handled by the caller */ }
        continue;
      }
      /* does not fit: split by lines */
      const lines = LY.lines(frag, frame.top, scale);
      /* Word's "multiple" line spacing: the last line on a page only needs its single-spaced height to
         fit; the extra leading may run into the bottom margin (same capacity as Word and LibreOffice) */
      const lsx = R.lineSpec(pPr);
      const multX = !ctx.inCell && frame.kind === 'col' && lsx.mult > 1 && !lsx.atLeast && !lsx.exact ? lsx.mult : 1;
      const slackOf = (ln) => (multX > 1 && ln ? (ln.bottom - ln.top) * (1 - 1 / multX) : 0);
      /* (an empty paragraph has no measured lines: its text bottom is the box minus the space after) */
      let textBottom;
      if (lines.length) textBottom = lines[lines.length - 1].bottom - slackOf(lines[lines.length - 1]);
      else {
        const pcR = frag.firstChild.getBoundingClientRect();
        const pcTop = (pcR.top - frame.top) / scale, pcBot = (pcR.bottom - frame.top) / scale;
        textBottom = pcBot - slackOf({ top: pcTop, bottom: pcBot });
      }
      if (textBottom <= limit + 0.5 && (lines.length || !n)) {
        /* only the space after overflows: keep it on this page */
        if (frame.notes) frame.notes.commit(frag);
        frame.empty = false;
        if (segEnd < n || brk >= 0) { frame.chain = []; return { st: afterBreak(st, segEnd, n), stop: brkType === 'column' ? 'column' : 'page' }; }
        if (pPr.keepNext && frame.kind === 'col') frame.chain.push({ bi: st.bi, el: frag, sub: o }); else frame.chain = [];
        st = { bi: st.bi + 1, sub: null };
        continue;
      }
      dropFrag(frag);
      if (frame.notes) frame.notes.rollback();
      /* first guess against the space left above the notes already on the page: notes referenced further
         down the paragraph may move to the next page with their lines (checked line by line below) */
      const kLimit = frame.notes ? frame.limit - frame.notes.height() : limit;
      let k = 0;
      while (k < lines.length && lines[k].bottom - slackOf(lines[k]) <= kLimit + 0.5) k++;
      const widow = pPr.widow !== false;
      const nl = lines.length;
      if (pPr.keepLines && !frame.empty && !ctx.inCell) k = 0;
      if (widow && nl > 1) {
        if (k === 1 && !frame.empty) k = 0;
        if (nl - k === 1 && k >= 2) k--;
        if (k === 1 && nl > 2 && !frame.empty) k = 0;
      }
      if (k === 0) {
        if (frame.empty || frame.limit - (frame.usedTop || 0) < 4) {
          if (!frame.empty) return keepChainBreak(frame, st, ctx);
          k = Math.min(1, nl);
          if (!nl || (nl === 1)) {
            /* a single line taller than the frame: place it anyway */
            frag = R.para(p, o, segEnd, fctx);
            frame.el.appendChild(frag);
            postPara(frag, p, fctx, frame);
            if (frame.notes) { frame.notes.measure(frag); frame.notes.commit(frag); }
            frame.empty = false;
            frame.chain = [];
            if (segEnd < n) return { st: { bi: st.bi, sub: segEnd }, stop: brkType === 'column' ? 'column' : 'page' };
            st = { bi: st.bi + 1, sub: null };
            if (ctx.inCell) continue;
            return { st, stop: 'full' };
          }
        } else return keepChainBreak(frame, st, ctx);
      }
      /* place lines [0,k) */
      let placed = false;
      while (k > 0) {
        const cut = lines[k] ? lines[k].o : segEnd;
        if (cut <= o) { k--; continue; }
        frag = R.para(p, o, cut, fctx);
        frame.el.appendChild(frag);
        postPara(frag, p, fctx, frame);
        const nh = frame.notes ? frame.notes.measure(frag) : 0;
        const bt = bottomOf(frag, frame.top, scale);
        const ls = bt - (parseFloat(frag.style.paddingBottom) || 0);
        if (ls - slackOf(lines[k - 1]) <= frame.limit - nh + 0.5 || (k === 1 && frame.empty)) {
          if (frame.notes) frame.notes.commit(frag);
          frame.empty = false;
          frame.chain = [];
          placed = true;
          st = cut >= n ? { bi: st.bi + 1, sub: null } : { bi: st.bi, sub: cut };
          break;
        }
        dropFrag(frag);
        if (frame.notes) frame.notes.rollback();
        k--;
      }
      if (!placed) return frame.empty ? { st, stop: 'full' } : keepChainBreak(frame, st, ctx);
      return { st, stop: 'full' };
    }
    return { st, stop: 'done' };
  };
  /** Word frames (framePr): positioned boxes of one or more paragraphs; text wraps around them */
  function placeFrame(frame, blocks, st, end, ctx, fp) {
    const d = ctx.doc;
    const key = JSON.stringify(fp);
    let j = st.bi;
    while (j < end && blocks[j].t === 'p' && !blocks[j].sect && JSON.stringify(D.pProps(d, blocks[j], ctx.tbl || null).frame || null) === key) j++;
    if (j === st.bi) j = st.bi + 1;
    /* frames placed relative to the page or margin (the default when no anchor is given), and frames
       with wrap="none" (text runs through them: page numbers in footers), sit outside the text flow */
    /* (with no anchor the position counts from the page, but a frame with no vertical position at all stays in the text) */
    const vPage = fp.vAnchor === 'page' || fp.vAnchor === 'margin' || (!fp.vAnchor && (typeof fp.y === 'number' || (fp.yAlign && fp.yAlign !== 'inline')));
    if ((fp.wrap === 'none' || vPage) && frame.kind !== 'cell' && !ctx.inTextBox) return placeFrameAbs(frame, blocks, st, j, ctx, fp, fp.wrap !== 'none' && fp.wrap !== 'through' && frame.kind === 'col');
    const box = el('div', 'frm');
    const css = [];
    const around = !fp.wrap || fp.wrap === 'around' || fp.wrap === 'tight' || fp.wrap === 'through';
    const right = fp.xAlign === 'right' || fp.xAlign === 'outside';
    if (fp.w) css.push(`width:${px(fp.w)}px`);
    if (fp.h && fp.hRule === 'exact') css.push(`height:${px(fp.h)}px`, 'overflow:hidden');
    else if (fp.h) css.push(`min-height:${px(fp.h)}px`);
    if (around) {
      css.push(`float:${right ? 'right' : 'left'}`, `clear:${right ? 'right' : 'left'}`);
      css.push(right ? `margin-left:${px(fp.hSpace || 0)}px` : `margin-right:${px(fp.hSpace || 0)}px`);
    } else css.push(fp.w ? '' : 'display:table');
    if (fp.xAlign === 'center') css.push('margin-left:auto', 'margin-right:auto', around ? 'float:none' : '', 'display:table');
    else if (typeof fp.x === 'number' && !fp.xAlign && fp.x > 0) {
      /* x counts from the page edge unless the frame is anchored to the margin or the text column */
      const pg = ctx.page || frame.page;
      const ml = !fp.hAnchor || fp.hAnchor === 'page' ? px(fp.x) - (pg ? pg.bodyLeft : 0) : px(fp.x);
      if (ml > 0) css.push(`margin-left:${ml}px`);
    }
    if (typeof fp.y === 'number' && fp.y > 0 && fp.vAnchor !== 'page' && fp.vAnchor !== 'margin') css.push(`margin-top:${px(fp.y)}px`);
    if (fp.vSpace) css.push(`margin-bottom:${px(fp.vSpace)}px`);
    box.style.cssText = css.filter(Boolean).join(';');
    frame.el.appendChild(box);
    for (let k = st.bi; k < j; k++) {
      const p = blocks[k];
      if (p.t !== 'p') continue;
      const fctx = Object.assign({}, ctx, { prev: k > st.bi ? blocks[k - 1] : null, next: k + 1 < j ? blocks[k + 1] : null, openCmts: LY.openFor(p), inFrame: true });
      const fr = R.para(p, 0, null, fctx);
      box.appendChild(fr);
      postPara(fr, p, fctx, frame);
    }
    const bottom = bottomOf(box, frame.top, LY.scale());
    if (bottom > frame.limit + 0.5 && !frame.empty && frame.limit !== Infinity) { for (const f of box.querySelectorAll('.p')) dropFrag(f); box.remove(); return { placed: false }; }
    return { placed: true, next: j };
  }
  function placeFrameAbs(frame, blocks, st, j, ctx, fp, exclude) {
    const wrap = el('div', 'frm-anchor');
    wrap.style.cssText = 'position:relative;height:0;margin:0;padding:0;overflow:visible';
    const box = el('div', 'frm frm-abs');
    const css = ['position:absolute', 'z-index:2', fp.w ? `width:${px(fp.w)}px` : 'width:max-content'];
    if (fp.h && fp.hRule === 'exact') css.push(`height:${px(fp.h)}px`, 'overflow:hidden');
    else if (fp.h) css.push(`min-height:${px(fp.h)}px`);
    box.style.cssText = css.join(';');
    wrap.appendChild(box);
    frame.el.appendChild(wrap);
    for (let k = st.bi; k < j; k++) {
      const p = blocks[k];
      if (p.t !== 'p') continue;
      const fctx = Object.assign({}, ctx, { prev: k > st.bi ? blocks[k - 1] : null, next: k + 1 < j ? blocks[k + 1] : null, openCmts: LY.openFor(p), inFrame: true });
      const fr = R.para(p, 0, null, fctx);
      box.appendChild(fr);
      postPara(fr, p, fctx, frame);
    }
    wrap._fp = fp;
    wrap._page = ctx.page || frame.page || null;
    LY.posAbsFrame(wrap);
    if (exclude) {
      /* keep the text clear of the frame: an empty float of its size, pushed down to its top */
      const y = parseFloat(box.style.top) || 0, x = parseFloat(box.style.left) || 0;
      const bw = box.offsetWidth, bh = box.offsetHeight, colW = frame.el.clientWidth || frame.width || 0;
      const vs = px(fp.vSpace || 0), hs = px(fp.hSpace || 0);
      const top = y - vs, h = bh + vs * 2;
      const band = Math.max(0, top >= 0 ? h : h + top);
      if (band > 0) {
        const full = fp.wrap === 'notBeside' || bw >= colW * 0.66;
        const left = x + bw / 2 < colW / 2;
        const side = full || left ? 'left' : 'right';
        const sp = el('div', 'fl-sp', `float:${side};clear:both;width:0;height:${Math.max(0, top)}px`);
        const ph = el('div', 'frm-excl', full ? `float:left;clear:both;width:100%;height:${band}px` : left
          ? `float:left;clear:left;width:${Math.max(0, x) + bw + hs}px;height:${band}px`
          : `float:right;clear:right;width:${Math.max(0, colW - x) + hs}px;height:${band}px`);
        wrap.append(sp, ph);
      }
    }
    return { placed: true, next: j };
  }
  /** position an out-of-flow frame from its anchors (re-run when its container shifts) */
  LY.posAbsFrame = function (wrap) {
    const fp = wrap._fp, box = wrap.firstChild, pg = wrap._page;
    if (!fp || !box) return;
    const scale = LY.scale();
    const wr = wrap.getBoundingClientRect();
    const host = wrap.parentElement;
    const hr = host.getBoundingClientRect();
    const pr = pg && pg.el && pg.el.isConnected ? pg.el.getBoundingClientRect() : null;
    const bw = box.offsetWidth, bh = box.offsetHeight;
    let ax0 = (hr.left - wr.left) / scale, aw = host.clientWidth || (pg ? pg.bodyW : 0);
    if (pr && (fp.hAnchor === 'page' || (!fp.hAnchor && (typeof fp.x === 'number' || fp.xAlign)))) { ax0 = (pr.left - wr.left) / scale; aw = pg.W; }
    else if (pr && fp.hAnchor === 'margin') { ax0 = (pr.left - wr.left) / scale + pg.bodyLeft; aw = pg.bodyW; }
    const xa = { left: 0, inside: 0, center: (aw - bw) / 2, right: aw - bw, outside: aw - bw };
    const x = ax0 + (fp.xAlign && xa[fp.xAlign] != null ? xa[fp.xAlign] : fp.x ? px(fp.x) : 0);
    let y;
    if (pr && (fp.vAnchor === 'page' || fp.vAnchor === 'margin' || (!fp.vAnchor && (typeof fp.y === 'number' || (fp.yAlign && fp.yAlign !== 'inline'))))) {
      const s = pg.sect || {};
      let ay0 = (pr.top - wr.top) / scale, ah = pg.H;
      if (fp.vAnchor === 'margin') { const mt = px(Math.abs(s.mt || 0)), mb = px(Math.abs(s.mb || 0)); ay0 += mt; ah = pg.H - mt - mb; }
      const ya = { top: 0, inside: 0, center: (ah - bh) / 2, bottom: ah - bh, outside: ah - bh };
      y = ay0 + (fp.yAlign && ya[fp.yAlign] != null ? ya[fp.yAlign] : fp.y ? px(fp.y) : 0);
    } else y = fp.y ? px(fp.y) : 0;
    box.style.left = L.round(x, 2) + 'px';
    box.style.top = L.round(y, 2) + 'px';
  };
  const fixAbsFrames = (root) => { for (const w of root.querySelectorAll('.frm-anchor')) LY.posAbsFrame(w); };
  /** move trailing keep-with-next paragraphs to the next frame along with the block that did not fit */
  function keepChainBreak(frame, st, ctx) {
    const ch = frame.chain || [];
    if (ch.length && !ctx.inCell) {
      const firstEl = ch[0].el;
      /* never empty a frame entirely */
      if (firstEl.previousElementSibling && ch[0].sub === 0) {
        for (const c of ch) dropFrag(c.el);
        if (frame.notes) frame.notes.removeFrom(firstEl);
        frame.chain = [];
        return { st: { bi: ch[0].bi, sub: 0 }, stop: 'full' };
      }
    }
    return { st, stop: 'full' };
  }
  /** after inserting a paragraph fragment: tabs, floats */
  function postPara(frag, p, ctx, frame) {
    if (frag.querySelector('.tb')) {
      const colRect = frame.el.getBoundingClientRect();
      R.layoutTabs(frag, p, ctx, colRect.left + (frame.leftPad || 0) * LY.scale(), LY.scale(), frame.width || frame.el.offsetWidth);
    }
    if (L.drawing && (frag.querySelector('.anc') || frag.querySelector('[data-pending]'))) L.drawing.anchorFloats(frag, p, ctx, frame);
  }

  /* ================= tables ================= */
  function tableShell(t, ctx) {
    const d = ctx.doc;
    const { tp, sp } = R.tblProps(d, t);
    const tbl = el('table', 'tbl');
    tbl.dataset.tid = t.id;
    tbl.contentEditable = 'inherit';
    const cg = el('colgroup');
    let total = 0;
    const grid = t.grid && t.grid.length ? t.grid : [ctx.width ? ctx.width / PX : 432];
    for (const w of grid) { const c = el('col'); c.style.width = px(w) + 'px'; cg.appendChild(c); total += w; }
    tbl.appendChild(cg);
    tbl.style.width = px(total) + 'px';
    const compat = (d.settings.compat || 11) < 15;
    const ind = (tp.ind || 0) - (compat ? tp.cellMar.l || 0 : 0);
    if (tp.jc === 'center') tbl.style.marginLeft = tbl.style.marginRight = 'auto';
    else if (tp.jc === 'right' || tp.jc === 'end') tbl.style.marginLeft = 'auto';
    else if (ind) tbl.style.marginLeft = px(ind) + 'px';
    if (tp.shd) tbl.style.backgroundColor = R.shadeCSS(tp.shd);
    const tb = el('tbody');
    tbl.appendChild(tb);
    return { tbl, tb, tp, sp, grid };
  }
  function cellBorders(tp, cs, tcPr, ri, rowspan, c0, span, nRows, nCols, contTop) {
    const tb = tp.borders || {};
    const own = tcPr.borders || {};
    const cb = cs.borders || {};
    const pick = (k, outer, inner, isOuter) => own[k] || cb[k] || (isOuter ? tb[outer] : tb[inner]);
    return {
      top: pick('top', 'top', 'insideH', ri === 0 || contTop),
      bottom: pick('bottom', 'bottom', 'insideH', ri + rowspan >= nRows),
      left: pick('left', 'left', 'insideV', c0 === 0) || own.start,
      right: pick('right', 'right', 'insideV', c0 + span >= nCols) || own.end,
    };
  }
  /** render one row; cellStates: array (per cell index) of {bi,sub}|'done'|null; limit: content height limit for split rows */
  function renderRow(t, ri, mapRow, shell, ctx, cellStates, limit, contTop, rowspanOverride) {
    const d = ctx.doc;
    const row = t.rows[ri];
    const tr = el('tr');
    tr.dataset.ri = ri;
    const nRows = t.rows.length, nCols = shell.grid.length;
    const trPr = row.trPr || {};
    if (trPr.h && trPr.hRule !== 'auto') tr.style.height = px(trPr.h) + 'px';
    if (trPr.hidden) tr.style.display = 'none';
    const rem = [];
    let anyRemaining = false, anyContent = false;
    if (trPr.gridBefore) { const td = el('td', 'gb'); td.colSpan = trPr.gridBefore; td.contentEditable = 'false'; tr.appendChild(td); }
    mapRow.forEach((m, ci) => {
      if (m.hidden) { rem.push('done'); return; }
      const cell = m.cell;
      const tcPr = cell.tcPr || {};
      const cs = R.cellStyle(d, t, shell.sp, ri, m.c0, nRows, nCols);
      const td = el('td');
      td.dataset.cid = cell.id;
      td.dataset.ri = ri;
      td.dataset.ci = ci;
      if (m.span > 1) td.colSpan = m.span;
      const rs = rowspanOverride ? rowspanOverride(m) : m.rowspan;
      if (rs > 1) td.rowSpan = rs;
      const mar = Object.assign({}, shell.tp.cellMar, tcPr.mar || {});
      const bd = cellBorders(shell.tp, cs, tcPr, ri, m.rowspan, m.c0, m.span, nRows, nCols, contTop);
      const st = [`padding:${px(mar.t || 0)}px ${px(mar.r || 0)}px ${px(mar.b || 0)}px ${px(mar.l || 0)}px`,
        `border-top:${R.borderCSS(bd.top)}`, `border-bottom:${R.borderCSS(bd.bottom)}`, `border-left:${R.borderCSS(bd.left)}`, `border-right:${R.borderCSS(bd.right)}`];
      const shd = tcPr.shd || cs.tcPr.shd;
      const bg = shd ? R.shadeCSS(shd) : '';
      if (bg) st.push(`background-color:${bg}`);
      /* diagonal borders: a hairline gradient whose middle runs corner to corner */
      const diag = [];
      for (const [k, dir] of [['tl2br', 'to top right'], ['tr2bl', 'to bottom right']]) {
        const b = (tcPr.borders && tcPr.borders[k]) || (cs.borders && cs.borders[k]);
        if (!R.hasBorder(b)) continue;
        const w = Math.max(1, px(b.sz || 0.5));
        const col = !b.color || b.color === 'auto' ? '#000' : '#' + String(b.color).replace('#', '');
        diag.push(`linear-gradient(${dir}, transparent calc(50% - ${w / 2 + 0.3}px), ${col} calc(50% - ${w / 2}px), ${col} calc(50% + ${w / 2}px), transparent calc(50% + ${w / 2 + 0.3}px))`);
      }
      if (diag.length) st.push(`background-image:${diag.join(',')}`);
      const va = tcPr.vAlign || cs.tcPr.vAlign;
      st.push(`vertical-align:${va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top'}`);
      if (tcPr.textDir === 'btLr') st.push('writing-mode:vertical-rl;transform:rotate(180deg)');
      else if (tcPr.textDir === 'tbRl') st.push('writing-mode:vertical-rl');
      td.style.cssText = st.join(';');
      const cf = el('div', 'cf');
      /* rows take the borders' laid-out width (Word), not the device-pixel-snapped width the browser draws */
      const bex = (R.borderExcess(bd.top) + R.borderExcess(bd.bottom)) / 2;
      if (Math.abs(bex) > 0.01) cf.style.marginBottom = L.round(-bex, 3) + 'px';
      /* while the cell is being filled its contents must not make the browser lay out the whole table again
         after every paragraph: size/layout containment isolates it until fillRow releases it */
      cf.style.contain = 'size layout';
      if (trPr.hRule === 'exact' && trPr.h) { cf.style.maxHeight = Math.max(0, px(trPr.h) - px((mar.t || 0) + (mar.b || 0))) + 'px'; cf.style.overflow = 'hidden'; }
      td.appendChild(cf);
      tr.appendChild(td);
      const sst = cellStates ? cellStates[ci] : null;
      if (sst === 'done') { rem.push('done'); return; }
      const start = sst || { bi: 0, sub: null };
      const tctx = { tblRPr: D.mergeRPr(shell.sp.rPr, cs.rPr), tblPPr: D.mergePPr(shell.sp.pPr, cs.pPr) };
      const cctx = Object.assign({}, ctx, { tbl: tctx, inCell: true, dkPara: bg && R.isDark(bg.replace('#', '')) });
      const frame = { el: cf, limit: limit == null ? Infinity : Math.max(0, limit - px((mar.t || 0) + (mar.b || 0))), top: 0, empty: true, kind: 'cell', width: null };
      if (limit != null) { frame.top = cf.getBoundingClientRect().top; }
      frame.width = Math.max(10, shell.grid.slice(m.c0, m.c0 + m.span).reduce((a, b) => a + b, 0) * PX - px((mar.l || 0) + (mar.r || 0)));
      if (limit != null) {
        /* the cell must be in the document to measure: caller already appended tr */
      }
      cell._frame = { frame, start, cctx };
      rem.push(start);
    });
    return { tr, rem, anyRemaining, anyContent };
  }
  /** fill rendered row cells (after the row is attached to the DOM) */
  function fillRow(tr, t, ri, limit) {
    const rem = [];
    let anyRemaining = false, anyContent = false;
    const row = t.rows[ri];
    for (const cell of row.cells) {
      const f = cell._frame;
      if (!f) { rem.push('done'); continue; }
      delete cell._frame;
      if (f.start === 'done') f.frame.el.style.contain = '';
      if (limit != null) f.frame.top = f.frame.el.getBoundingClientRect().top;
      if (f.start === 'done') { rem.push('done'); continue; }
      const res = LY.flow(f.frame, cell.blocks, f.start, cell.blocks.length, f.cctx);
      f.frame.el.style.contain = '';
      if (!f.frame.empty) anyContent = true;
      if (res.stop === 'done') rem.push('done');
      else { rem.push(res.st); anyRemaining = true; }
    }
    void tr;
    return { rem, anyRemaining, anyContent };
  }
  /**
   * place a table (from state sub = {r, cells}) into frame.
   * returns {done, placed, sub}
   */
  function placeTable(frame, t, sub, ctx) {
    const scale = LY.scale();
    const shell = tableShell(t, ctx);
    const map = R.tblMap(t);
    const startRow = sub ? sub.r : 0;
    let cellStates = sub ? sub.cells : null;
    frame.el.appendChild(shell.tbl);
    const nRows = t.rows.length;
    /* repeated header rows */
    let hdrCount = 0;
    while (hdrCount < nRows && t.rows[hdrCount].trPr.header) hdrCount++;
    if (startRow > 0 && hdrCount && startRow >= hdrCount && frame.kind === 'col') {
      const reps = [];
      for (let hr = 0; hr < hdrCount; hr++) {
        const rr = renderRow(t, hr, map[hr], shell, ctx, null, null, hr === 0, (m) => Math.min(m.rowspan, hdrCount - hr));
        rr.tr.classList.add('rep');
        shell.tb.appendChild(rr.tr);
        fillRow(rr.tr, t, hr, null);
        reps.push(rr.tr);
      }
      /* heading rows that leave no room for the table's own rows are not repeated (Word) */
      if (frame.limit !== Infinity && reps.length && bottomOf(reps[reps.length - 1], frame.top, scale) + 20 > frame.limit) for (const tr of reps) tr.remove();
    }
    let placedRows = 0;
    let ri = startRow;
    const lim = frame.limit;
    for (; ri < nRows; ri++) {
      /* rows inside a vertical merge must stay with their group: compute rows remaining in the merge started above */
      const firstOnFrame = placedRows === 0;
      const rowspanCut = (m) => Math.min(m.rowspan, nRows - ri);
      const rr = renderRow(t, ri, map[ri], shell, ctx, firstOnFrame ? cellStates : null, null, firstOnFrame && ri > 0 && frame.kind === 'col' && sub, rowspanCut);
      shell.tb.appendChild(rr.tr);
      fillRow(rr.tr, t, ri, null);
      if (lim === Infinity) { placedRows++; continue; }
      const bt = bottomOf(rr.tr, frame.top, scale);
      /* rowspans extending below: measure the bottom of the whole group */
      if (bt <= lim + 0.5) { placedRows++; continue; }
      /* row does not fit */
      const trPr = t.rows[ri].trPr || {};
      const canSplit = !trPr.cantSplit && !trPr.header;
      const trTop = (rr.tr.getBoundingClientRect().top - frame.top) / scale;
      rr.tr.remove();
      if (canSplit && map[ri].every((m) => m.rowspan === 1)) {
        const avail = lim - trTop - 1;
        if (avail > 8) {
          const r2 = renderRow(t, ri, map[ri], shell, ctx, firstOnFrame ? cellStates : null, avail, firstOnFrame && ri > 0, rowspanCut);
          shell.tb.appendChild(r2.tr);
          const f2 = fillRow(r2.tr, t, ri, avail);
          const bt2 = bottomOf(r2.tr, frame.top, scale);
          if (f2.anyContent && (bt2 <= lim + 2 || frame.empty && placedRows === 0)) {
            if (!f2.anyRemaining) { placedRows++; continue; }
            r2.tr.classList.add('split');
            return { done: false, placed: true, sub: { r: ri, cells: f2.rem } };
          }
          r2.tr.remove();
        }
      }
      if (placedRows === 0 && frame.empty) {
        /* nothing fits even on an empty frame: place the row anyway (clipped) */
        const r3 = renderRow(t, ri, map[ri], shell, ctx, cellStates, null, ri > 0 && !!sub, rowspanCut);
        shell.tb.appendChild(r3.tr);
        fillRow(r3.tr, t, ri, null);
        placedRows++;
        if (ri + 1 < nRows) return { done: false, placed: true, sub: { r: ri + 1, cells: null } };
        continue;
      }
      /* move the remaining rows to the next frame; never split inside a vertical merge */
      let back = ri;
      while (back > startRow && map[back].some((m) => m.cell.tcPr.vMerge === 'continue')) back--;
      if (back !== ri && back > startRow) {
        const trs = Array.from(shell.tb.children).filter((x) => +x.dataset.ri >= back && !x.classList.contains('rep'));
        for (const x of trs) x.remove();
        placedRows -= ri - back;
        ri = back;
      }
      if (placedRows <= 0) { shell.tbl.remove(); return { done: false, placed: false, sub: sub || null }; }
      return { done: false, placed: true, sub: { r: ri, cells: null } };
    }
    return { done: true, placed: true, sub: null };
  }

  /* ================= footnote areas ================= */
  /** per-page footnote collector: measures/commits notes referenced by fragments */
  function noteArea(page, ctx) {
    const area = el('div', 'pg-fn');
    area.style.cssText = `left:${page.bodyLeft}px;width:${page.bodyW}px;bottom:${page.bodyBottomGap}px`;
    page.el.appendChild(area);
    let committed = [];
    let pending = [];
    let committedH = 0;
    let pendingH = 0;
    const cache = new Map();
    const renderNote = (id) => {
      const note = doc().fn[id];
      const box = el('div', 'fnote');
      box.dataset.fn = id;
      if (!note) return box;
      const fr = { el: box, limit: Infinity, top: 0, empty: true, kind: 'note' };
      LY.flow(fr, note.blocks, { bi: 0, sub: null }, note.blocks.length, Object.assign({}, ctx, { story: note, noteOwner: id, grid: 0 }));
      return box;
    };
    const api = {
      area,
      /** height needed (px) if fragment's notes were added */
      measure(frag) {
        pending = [];
        pendingH = 0;
        const refs = frag.querySelectorAll('.nref[data-note^="fn:"]');
        if (!refs.length) return committedH;
        for (const r of refs) {
          const id = r.dataset.note.slice(3);
          if (committed.some((c) => c.id === id) || pending.some((c) => c.id === id)) continue;
          let box = cache.get(id);
          if (!box) { box = renderNote(id); cache.set(id, box); }
          if (!area.childNodes.length) { const sep = el('div', 'fnsep-line'); sep.contentEditable = 'false'; area.appendChild(sep); }
          area.appendChild(box);
          pending.push({ id, box });
        }
        pendingH = area.offsetHeight + 6;
        return pendingH;
      },
      commit() { committed = committed.concat(pending); pending = []; committedH = committed.length ? area.offsetHeight + 6 : 0; },
      rollback() { for (const pnd of pending) pnd.box.remove(); pending = []; if (!committed.length) area.textContent = ''; },
      removeFrom(fragEl) {
        /* drop notes referenced by fragments being moved off this page */
        const ids = new Set();
        let e = fragEl;
        while (e) { for (const r of e.querySelectorAll ? e.querySelectorAll('.nref[data-note^="fn:"]') : []) ids.add(r.dataset.note.slice(3)); e = e.nextElementSibling; }
        committed = committed.filter((c) => { if (ids.has(c.id)) { c.box.remove(); return false; } return true; });
        if (!committed.length) area.textContent = '';
        committedH = committed.length ? area.offsetHeight + 6 : 0;
      },
      height: () => committedH,
    };
    return api;
  }

  /* ================= pages ================= */
  /** header/footer story for a page */
  LY.hfFor = function (sections, si, kind, pageType) {
    for (let k = si; k >= 0; k--) {
      const refs = sections[k].sect.refs && sections[k].sect.refs[kind];
      if (!refs) continue;
      const id = refs[pageType];
      if (id && doc().hf[id]) return doc().hf[id];
    }
    if (pageType !== 'default') {
      if (pageType === 'first') return null;
      if (pageType === 'even') return null;
    }
    return null;
  };
  LY.hfRefFor = function (sections, si, kind, pageType) {
    for (let k = si; k >= 0; k--) {
      const refs = sections[k].sect.refs && sections[k].sect.refs[kind];
      if (refs && refs[pageType]) return { si: k, id: refs[pageType] };
    }
    return null;
  };
  function pageType(sect, firstOfSection, pageNo) {
    if (sect.titlePg && firstOfSection) return 'first';
    if (doc().settings.evenOdd && pageNo % 2 === 0) return 'even';
    return 'default';
  }
  /** render a header/footer story into a container */
  function renderHF(container, story, ctx, page) {
    if (!story) return 0;
    D.info(doc(), story.paras && story.paras[0]);
    const fr = { el: container, limit: Infinity, top: 0, empty: true, kind: 'hf', width: page.bodyW };
    const hctx = Object.assign({}, ctx, { story, dyn: (ty, fmt, instr) => LY.dynField(ty, fmt, page, instr), openCmts: null, page, hf: true, grid: 0 });
    LY.flow(fr, story.blocks, { bi: 0, sub: null }, story.blocks.length, hctx);
    fixAbsFrames(container);
    return container.offsetHeight;
  }
  LY.dynField = function (ty, fmt, page, instr) {
    if (!page) return '1';
    if (ty === 'STYLEREF') return page.body && page.body.querySelector('.p') ? LY.styleRefOnPage(page, instr || '') : '';
    const f = { Arabic: 'decimal', ROMAN: 'upperRoman', roman: 'lowerRoman', ALPHABETIC: 'upperLetter', alphabetic: 'lowerLetter', Ordinal: 'ordinal', CardText: 'cardinalText', OrdText: 'ordinalText' }[fmt];
    if (ty === 'PAGE') return D.fmtNum(page.numVal, f || page.numFmt || 'decimal');
    if (ty === 'NUMPAGES') return D.fmtNum(LY.pages.length || 1, f || 'decimal');
    if (ty === 'SECTIONPAGES') return D.fmtNum(LY.pages.filter((p) => p.si === page.si).length || 1, f || 'decimal');
    if (ty === 'SECTION') return String(page.si + 1);
    return '';
  };
  /**
   * STYLEREF in a header or footer: the first (\l: last) text on this page in the given paragraph or
   * character style; failing that, the nearest such text before the page, then after it (as Word does).
   */
  LY.styleRefOnPage = function (page, instr) {
    const d = doc();
    const f = L.fields.parse(instr);
    const name = String(f.args[0] || '');
    const last = !!f.sw['\\l'];
    const level = /^\d$/.test(name) ? +name : 0;
    const st = level ? null : D.findStyleByName(d, name.toLowerCase()) || d.styles[name];
    if (!level && !st) return 'Error! No text of specified style in document.';
    const isChar = !!st && st.type === 'character';
    const paraHit = (q) => (level ? D.headingLevel(d, q) === level : (q.pPr.style || 'Normal') === st.id);
    /* text of character-styled runs in [from, to): first or last contiguous stretch */
    const charHit = (q, from, to, fromEnd) => {
      const spans = [];
      let pos = 0, cur = null;
      for (const it of q.runs) {
        const n = D.ilen(it);
        const inside = pos + n > from && pos < to;
        if (it.t === 'text' && inside && it.rPr && it.rPr.style === st.id) {
          const t = it.text.slice(Math.max(0, from - pos), Math.min(n, to - pos));
          if (cur) cur.push(t); else { cur = [t]; spans.push(cur); }
        } else if (n > 0) cur = null;
        pos += n;
      }
      if (!spans.length) return null;
      return (fromEnd ? spans[spans.length - 1] : spans[0]).join('').trim() || null;
    };
    const hit = (q, from, to, fromEnd) => (isChar ? charHit(q, from, to == null ? D.plen(q) : to, fromEnd) : paraHit(q) ? D.plainText(q).trim() : null);
    let frags = Array.from(page.body.querySelectorAll('.p[data-pid]')).filter((e) => !e.closest('.shp, .frm-abs'));
    if (last) frags = frags.reverse();
    for (const e of frags) {
      const q = D.byId(d, +e.dataset.pid);
      if (!q) continue;
      const tt = parseFloat(e.dataset.to);
      const r = hit(q, +e.dataset.from || 0, isFinite(tt) ? tt : null, last);
      if (r) return r;
    }
    /* not on this page: search back from its start, then forward from its end */
    D.reindex(d);
    const firstE = frags.length ? (last ? frags[frags.length - 1] : frags[0]) : null;
    const lastE = frags.length ? (last ? frags[0] : frags[frags.length - 1]) : null;
    const firstQ = firstE && D.byId(d, +firstE.dataset.pid), lastQ = lastE && D.byId(d, +lastE.dataset.pid);
    /* the part of the page's first paragraph that lies on earlier pages comes first */
    if (firstQ && isChar && +firstE.dataset.from > 0) { const r = charHit(firstQ, 0, +firstE.dataset.from, true); if (r) return r; }
    for (let q = firstQ && D.prevPara(d, firstQ); q; q = D.prevPara(d, q)) { const r = hit(q, 0, null, true); if (r) return r; }
    if (lastQ && isChar) { const tt = parseFloat(lastE.dataset.to); if (isFinite(tt) && tt < D.plen(lastQ)) { const r = charHit(lastQ, tt, D.plen(lastQ), false); if (r) return r; } }
    for (let q = lastQ && D.nextPara(d, lastQ); q; q = D.nextPara(d, q)) { const r = hit(q, 0, null, false); if (r) return r; }
    return 'Error! No text of specified style in document.';
  };
  /** create a page DOM for section index si */
  function newPage(sections, si, pageIdx, numVal, firstOfSection, ctx) {
    const sect = sections[si].sect;
    const W = px(sect.pgW), H = px(sect.pgH);
    const pg = el('div', 'pg');
    pg.style.width = W + 'px';
    pg.style.height = H + 'px';
    pg.dataset.pg = pageIdx;
    const d = doc();
    if (d.bg && LY.opts.printBg !== false) pg.style.background = '#' + d.bg;
    const mirror = d.settings.mirror && numVal % 2 === 0;
    const ml = px(mirror ? sect.mr : sect.ml) + (mirror ? 0 : px(sect.gutter || 0)), mr = px(mirror ? sect.ml : sect.mr) + (mirror ? px(sect.gutter || 0) : 0);
    const page = { el: pg, si, sect, idx: pageIdx, numVal, numFmt: (sect.pgNum && sect.pgNum.fmt) || 'decimal', W, H, bodyLeft: ml, bodyW: W - ml - mr, firstOfSection };
    page.type = pageType(sect, firstOfSection, numVal);
    /* header & footer */
    const hdrStory = LY.hfFor(sections, si, 'hdr', page.type) || (page.type === 'even' ? null : null);
    const ftrStory = LY.hfFor(sections, si, 'ftr', page.type);
    page.hdrStory = hdrStory; page.ftrStory = ftrStory;
    const hdr = el('div', 'pg-hdr');
    hdr.contentEditable = LY.hfEdit ? 'true' : 'false';
    hdr.style.cssText = `left:${ml}px;width:${page.bodyW}px;top:${px(sect.hdr)}px`;
    hdr.dataset.kind = 'hdr';
    const ftr = el('div', 'pg-ftr');
    ftr.contentEditable = LY.hfEdit ? 'true' : 'false';
    ftr.style.cssText = `left:${ml}px;width:${page.bodyW}px;bottom:${px(sect.ftr)}px`;
    ftr.dataset.kind = 'ftr';
    pg.appendChild(hdr);
    pg.appendChild(ftr);
    if (hdrStory) hdr.dataset.sid = hdrStory.id;
    if (ftrStory) ftr.dataset.sid = ftrStory.id;
    /* during an incremental re-layout new pages go in front of the old ones still waiting to be reused */
    LY.root.insertBefore(pg, LY.insertAnchor && LY.insertAnchor.isConnected ? LY.insertAnchor : null);
    const hh = renderHF(hdr, hdrStory, ctx, page);
    const fh = renderHF(ftr, ftrStory, ctx, page);
    let top = px(sect.mt), bottomGap = px(sect.mb);
    if (sect.mt >= 0 && hh && px(sect.hdr) + hh > top) top = px(sect.hdr) + hh;
    if (sect.mb >= 0 && fh && px(sect.ftr) + fh > bottomGap) bottomGap = px(sect.ftr) + fh;
    top = Math.abs(top); bottomGap = Math.abs(bottomGap);
    page.bodyTop = top;
    page.bodyBottomGap = bottomGap;
    page.bodyH = H - top - bottomGap;
    const body = el('div', 'pg-body');
    body.style.cssText = `left:${ml}px;top:${top}px;width:${page.bodyW}px;height:${page.bodyH}px`;
    if (LY.hfEdit) body.contentEditable = 'false';
    pg.appendChild(body);
    page.body = body;
    page.hdr = hdr; page.ftr = ftr;
    page.regions = [];
    page.notes = null;
    /* text boundaries (corner marks) */
    if (LY.opts.textBoundaries) {
      const cm = el('div', 'crop');
      cm.contentEditable = 'false';
      cm.style.cssText = `left:${ml}px;top:${px(sect.mt)}px;width:${page.bodyW}px;height:${H - px(sect.mt) - px(sect.mb)}px`;
      pg.appendChild(cm);
    }
    /* page borders */
    if (sect.borders && L.drawing && L.drawing.pageBorders) L.drawing.pageBorders(page, sect.borders);
    else if (sect.borders) pageBorders(page, sect.borders);
    /* watermark */
    if (d.watermark && L.drawing && L.drawing.watermark) L.drawing.watermark(page, d.watermark);
    return page;
  }
  function pageBorders(page, b) {
    const pb = el('div', 'pgb');
    pb.contentEditable = 'false';
    const fromText = b.offsetFrom === 'text';
    const sect = page.sect;
    const g = (k) => (b[k] ? px(b[k].space || (fromText ? 4 : 24)) : 0);
    const inset = fromText ? { t: page.bodyTop - g('top'), l: page.bodyLeft - g('left'), r: page.W - page.bodyLeft - page.bodyW - g('right'), b: page.bodyBottomGap - g('bottom') } : { t: g('top'), l: g('left'), r: g('right'), b: g('bottom') };
    void sect;
    pb.style.cssText = `left:${inset.l}px;top:${inset.t}px;right:${inset.r}px;bottom:${inset.b}px;border-top:${R.borderCSS(b.top)};border-left:${R.borderCSS(b.left)};border-right:${R.borderCSS(b.right)};border-bottom:${R.borderCSS(b.bottom)}`;
    if (b.display === 'firstPage' && !page.firstOfSection) return;
    if (b.display === 'notFirstPage' && page.firstOfSection) return;
    page.el.appendChild(pb);
  }
  /** create a column region at y within the page body */
  function newRegion(page, sect, y) {
    const reg = el('div', 'pg-reg');
    reg.style.top = y + 'px';
    reg.style.width = page.bodyW + 'px';
    page.body.appendChild(reg);
    const cols = sect.cols || { n: 1 };
    const n = Math.max(1, cols.n || 1);
    const frames = [];
    let x = 0;
    for (let i = 0; i < n; i++) {
      let w, gap;
      if (cols.eq !== false || !cols.w || !cols.w[i]) { gap = px(cols.space != null ? cols.space : 36); w = (page.bodyW - gap * (n - 1)) / n; }
      else { w = px(cols.w[i].w); gap = px(cols.w[i].space || 0); }
      const c = el('div', 'pg-col');
      c.style.cssText = `left:${x}px;width:${w}px`;
      reg.appendChild(c);
      if (cols.sep && i > 0) { const s = el('div', 'colsep'); s.contentEditable = 'false'; s.style.left = (x - gap / 2) + 'px'; reg.appendChild(s); }
      frames.push({ el: c, x, w });
      x += w + gap;
    }
    const region = { el: reg, y, frames, ci: 0, sect };
    page.regions.push(region);
    return region;
  }
  function frameFor(page, region, ci, ctx, natural) {
    const f = region.frames[ci];
    const sc = LY.scale();
    const fr = { el: f.el, limit: page.bodyH - region.y, top: f.el.getBoundingClientRect().top, empty: !f.el.firstChild, natural, kind: 'col', page, width: f.w, notes: page.notes };
    if (!fr.empty) fr.top = f.el.getBoundingClientRect().top;
    void sc;
    void ctx;
    return fr;
  }
  function regionBottom(region) {
    let m = 0;
    for (const f of region.frames) { const last = f.el.lastElementChild; if (last) m = Math.max(m, last.offsetTop + last.offsetHeight); }
    return region.y + m;
  }
  /** balance a multi-column region (before a continuous section break) */
  function balance(page, region, blocks, startSt, endIdx, ctx) {
    const n = region.frames.length;
    if (n < 2) return null;
    const total = region.frames.reduce((s, f) => s + (f.el.lastElementChild ? f.el.lastElementChild.offsetTop + f.el.lastElementChild.offsetHeight : 0), 0);
    if (!total) return null;
    let h = Math.ceil(total / n);
    const maxH = page.bodyH - region.y;
    for (let attempt = 0; attempt < 12 && h <= maxH; attempt++) {
      for (const f of region.frames) f.el.textContent = '';
      let st = startSt;
      let ok = true;
      for (let ci = 0; ci < n; ci++) {
        const fr = frameFor(page, region, ci, ctx, ci > 0);
        fr.limit = ci === n - 1 ? maxH : h;
        const res = LY.flow(fr, blocks, st, endIdx, ctx);
        st = res.st;
        if (res.stop === 'done') break;
        if (ci === n - 1) ok = res.stop === 'done';
      }
      if (ok && regionBottom(region) - region.y <= h + 2) return true;
      h += Math.max(6, Math.round(h * 0.06));
    }
    /* fall back to an unbalanced fill */
    for (const f of region.frames) f.el.textContent = '';
    let st = startSt;
    for (let ci = 0; ci < n; ci++) { const fr = frameFor(page, region, ci, ctx, ci > 0); const res = LY.flow(fr, blocks, st, endIdx, ctx); st = res.st; if (res.stop === 'done') break; }
    return false;
  }

  /* ================= print layout ================= */
  /**
   * Lay out the main story into pages. fromPage: index of the first page to rebuild (incremental).
   */
  LY.paginate = function (fromPage) {
    const d = doc();
    const t0 = performance.now();
    D.reindex(d);
    LY.labels = D.computeLists(d);
    LY.computeNotes();
    LY.openMap = computeOpenComments(d.main);
    const ctx = LY.ctx({ story: d.main });
    const sections = D.sections(d);
    const blocks = d.main.blocks;
    let pageIdx = 0, si = 0, st = { bi: 0, sub: null }, numVal = 1, firstOfSection = true;
    const oldPages = LY.pages;
    if (fromPage && fromPage > 0 && fromPage < oldPages.length && oldPages[fromPage].start) {
      const ps = oldPages[fromPage].start;
      const bi = blocks.indexOf(ps.block);
      if (bi >= 0) {
        pageIdx = fromPage; si = ps.si; st = { bi, sub: ps.sub }; numVal = ps.numVal; firstOfSection = ps.firstOfSection;
      } else fromPage = 0;
    } else fromPage = 0;
    /* pages from fromPage are laid out again; the old ones stay in the DOM (behind the new ones) until we know
       whether the layout re-joins them, so an edit early in a long document doesn't detach every later page */
    const keep = oldPages.slice(0, fromPage);
    const resyncCandidates = oldPages.slice(fromPage);
    if (!fromPage) { LY.root.textContent = ''; LY.insertAnchor = null; LY.tbChains = new Map(); }
    else LY.insertAnchor = resyncCandidates.length ? resyncCandidates[0].el : null;
    LY.pages = keep;
    let page = null, region = null;
    const startPage = (natural) => {
      page = newPage(sections, si, pageIdx, numVal, firstOfSection, ctx);
      page.start = { block: blocks[st.bi], sub: st.sub, si, numVal, firstOfSection };
      page.notes = d.fn && Object.keys(d.fn).length ? noteArea(page, ctx) : null;
      LY.pages.push(page);
      region = newRegion(page, sections[si].sect, 0);
      region.natural = natural;
      firstOfSection = false;
      pageIdx++;
      numVal++;
    };
    const sectStartNum = (k) => { const s = sections[k].sect; return s.pgNum && s.pgNum.start != null ? s.pgNum.start : null; };
    if (!fromPage) { const sn = sectStartNum(0); if (sn != null) numVal = sn; }
    startPage(false);
    let guard = 0;
    let resynced = false;
    while (si < sections.length) {
      if (++guard > 20000) { console.warn('layout guard'); break; }
      const sec = sections[si];
      const fr = frameFor(page, region, region.ci, ctx, region.natural || region.ci > 0);
      /* East Asian document grid: lines snap to the section's line pitch */
      const dg = sec.sect.docGrid;
      ctx.grid = dg && /^(lines|linesAndChars|snapToChars)$/.test(dg.type || '') && dg.linePitch > 0 ? px(dg.linePitch / 20) : 0;
      const res = LY.flow(fr, blocks, st, sec.to, ctx);
      st = res.st;
      /* Word 2013+ (compat 15): a page (or last-column) break that ends the section's last paragraph keeps
         the paragraph mark on its own page, so no empty page is made when the document ends or the next
         section starts on a new page anyway (older modes carry the mark over to a new page) */
      if ((d.settings && d.settings.compat) >= 15 && (res.stop === 'page' || (res.stop === 'column' && region.ci + 1 >= region.frames.length)) && !res.pbb && st.bi >= sec.to) {
        const nt = si + 1 < sections.length ? (sections[si + 1].sect.type || 'nextPage') : null;
        if (!nt || (nt !== 'continuous' && nt !== 'nextColumn')) res.stop = 'done';
      }
      if (res.stop === 'done') {
        si++;
        if (si >= sections.length) break;
        const nsect = sections[si].sect;
        const type = nsect.type || 'nextPage';
        const sameGeom = Math.abs(nsect.pgW - sec.sect.pgW) < 1 && Math.abs(nsect.pgH - sec.sect.pgH) < 1;
        const sn = sectStartNum(si);
        /* a "next column" break in a one-column layout starts a new page */
        const colLeft = region.ci + 1 < region.frames.length || region.frames.length > 1;
        if ((type === 'continuous' || (type === 'nextColumn' && colLeft)) && sameGeom) {
          if (type === 'nextColumn' && region.ci + 1 < region.frames.length && (nsect.cols || {}).n === (sec.sect.cols || {}).n) { region.ci++; continue; }
          /* balance columns of the region that just ended */
          if (region.frames.length > 1) balance(page, region, blocks, regionStartState(region, page, blocks), sec.to, ctx);
          const y = regionBottom(region) + 1;
          if (y >= page.bodyH - 12) { if (sn != null) numVal = sn; firstOfSection = true; startPage(false); continue; }
          region = newRegion(page, nsect, y);
          region.natural = false;
          if (sn != null) { /* restart numbering applies from the next page */ }
          continue;
        }
        if (sn != null) numVal = sn;
        firstOfSection = true;
        if ((type === 'evenPage' && numVal % 2 === 1) || (type === 'oddPage' && numVal % 2 === 0)) {
          /* blank page to reach the right parity */
          startPage(false);
          page.el.classList.add('blank');
          firstOfSection = true;
        }
        /* Word 2013+ also drops it at the top of a page that starts a new section */
        startPage((d.settings && d.settings.compat) >= 15);
        continue;
      }
      if (res.stop === 'column' && region.ci + 1 < region.frames.length) { region.ci++; continue; }
      if (res.stop === 'full' && region.ci + 1 < region.frames.length) { region.ci++; continue; }
      /* new page; try to resync with the previous layout */
      const nextStart = { block: blocks[st.bi], sub: st.sub };
      if (fromPage && LY.resyncOK) {
        const old = resyncCandidates.find((op) => op.idx === pageIdx && op.start && op.start.block === nextStart.block && JSON.stringify(op.start.sub) === JSON.stringify(nextStart.sub) && op.start.si === si && op.start.numVal === numVal);
        if (old && LY.resyncOK(old, blocks.indexOf(old.start.block))) {
          for (const op of resyncCandidates) { if (op.idx >= old.idx) LY.pages.push(op); else op.el.remove(); }
          resynced = true;
          break;
        }
      }
      /* space before is dropped at the top of a page reached by overflow or by "page break before";
         Word 2013+ documents (compat 15) drop it after a manual page break as well */
      startPage(res.stop === 'full' || !!res.pbb || (d.settings && d.settings.compat) >= 15);
    }
    if (!resynced) for (const op of resyncCandidates) op.el.remove();
    LY.insertAnchor = null;
    /* endnotes after the last block */
    if (!resynced && d.en && Object.keys(d.en).length) {
      si = sections.length - 1;
      const ens = Array.from(LY.noteNums.keys()).filter((k) => k.startsWith('en:')).map((k) => k.slice(3));
      if (ens.length) {
        const sep = D.para([D.item('sep')]);
        const eb = [sep];
        for (const id of ens) if (d.en[id]) for (const b of d.en[id].blocks) eb.push(b);
        let est = { bi: 0, sub: null };
        let g2 = 0;
        while (est.bi < eb.length && ++g2 < 500) {
          const fr = frameFor(page, region, region.ci, ctx, true);
          const res = LY.flow(fr, eb, est, eb.length, Object.assign({}, ctx, { endnotes: true }));
          est = res.st;
          if (res.stop === 'done') break;
          if (region.ci + 1 < region.frames.length) { region.ci++; continue; }
          startPage(true);
        }
      }
    }
    /* finish: page numbers that depend on the total */
    LY.refreshDynamic();
    for (const p of LY.pages) p.el.classList.toggle('vcenter', false);
    applyVAlign();
    LY.layoutMs = performance.now() - t0;
    LY.fitSizer();
    L.bus.emit('layout-done', { pages: LY.pages.length });
  };
  function regionStartState(region, page, blocks) {
    /* first fragment in the region */
    const first = region.frames[0].el.querySelector('.p, .tbl');
    if (!first) return { bi: 0, sub: null };
    const top = first.classList.contains('tbl') ? first : first;
    if (top.classList.contains('p')) {
      const p = D.byId(doc(), +top.dataset.pid);
      const bi = blocks.indexOf(p);
      if (bi >= 0) return { bi, sub: +top.dataset.from || null };
    }
    const tid = +top.dataset.tid;
    const bi = blocks.findIndex((b) => b.id === tid);
    return { bi: Math.max(0, bi), sub: null };
  }
  function applyVAlign() {
    for (const p of LY.pages) {
      const va = p.sect.vAlign;
      if (!va || va === 'top') continue;
      const reg = p.regions[p.regions.length - 1];
      if (!reg) continue;
      const used = regionBottom(reg);
      const free = p.bodyH - used;
      if (free <= 0) continue;
      const off = va === 'center' ? free / 2 : va === 'bottom' ? free : 0;
      if (va === 'both') continue;
      for (const r of p.regions) { r.el.style.top = r.y + off + 'px'; fixAbsFrames(r.el); }
    }
  }
  /** update PAGE/NUMPAGES fields after layout */
  LY.refreshDynamic = function () {
    for (const p of LY.pages) {
      for (const f of p.el.querySelectorAll('.fdyn')) f.textContent = LY.dynField(f.dataset.fty, f.dataset.fmt, p, f.dataset.instr);
    }
  };

  /* ================= flow views (Normal / Web / Outline) ================= */
  LY.flowView = function () {
    const d = doc();
    D.reindex(d);
    LY.labels = D.computeLists(d);
    LY.computeNotes();
    LY.openMap = computeOpenComments(d.main);
    LY.root.textContent = '';
    LY.pages = [];
    const sect = d.sect;
    const box = el('div', 'flow ' + LY.view);
    const w = LY.view === 'web' ? Math.max(200, (LY.scroller.clientWidth - 40) / LY.zoom) : px(sect.pgW - sect.ml - sect.mr - (sect.gutter || 0));
    box.style.width = w + 'px';
    if (LY.view === 'web' && d.bg) box.style.background = '#' + d.bg;
    LY.root.appendChild(box);
    const ctx = LY.ctx({ story: d.main, view: LY.view });
    const blocks = d.main.blocks;
    const frame = { el: box, limit: Infinity, top: box.getBoundingClientRect().top, empty: true, kind: 'flow', width: w };
    if (LY.view === 'outline' && L.outline) { L.outline.render(box, ctx); }
    else LY.flow(frame, blocks, { bi: 0, sub: null }, blocks.length, ctx);
    /* section / page break markers are rendered via CSS on .has-sect and .pbrk */
    LY.fitSizer();
    L.bus.emit('layout-done', { pages: 0 });
  };

  /* ================= top-level relayout ================= */
  LY.lastBlocks = [];
  LY.render = function (opts) {
    opts = opts || {};
    if (!LY.root) return;
    R.clearCaches();
    if (LY.view === 'print' || LY.view === 'reading' || LY.view === 'preview') LY.paginate(opts.fromPage || 0);
    else LY.flowView();
    LY.lastBlocks = doc().main.blocks.slice();
    LY.lastLabels = LY.labels;
    LY.lastNotes = LY.noteNums;
    LY.lastStyleVer = D.styleVer;
  };
  /**
   * Incremental relayout after a transaction: rebuild from the first affected page and stop as soon as
   * pagination matches the previous run again.
   */
  LY.update = function (info) {
    const d = doc();
    if (!LY.root) return;
    if (LY.view !== 'print' && LY.view !== 'reading' && LY.view !== 'preview' || !info || info.full || info.global || D.styleVer !== LY.lastStyleVer || !LY.pages.length) return LY.render();
    const blocks = d.main.blocks;
    /* first top-level index that changed */
    let first = Infinity, lastDirty = -1;
    const old = LY.lastBlocks;
    if (info.struct) {
      let i = 0;
      while (i < blocks.length && i < old.length && blocks[i] === old[i]) i++;
      first = Math.min(first, i);
      let j = 0;
      while (j < blocks.length && j < old.length && blocks[blocks.length - 1 - j] === old[old.length - 1 - j]) j++;
      lastDirty = Math.max(lastDirty, blocks.length - 1 - j);
    }
    D.reindex(d);
    const touchesOther = info.paras && Array.from(info.paras).some((p) => { const r = D.info(d, p); return !r || r.story !== d.main; });
    if (touchesOther) return LY.render();
    if (info.paras) for (const p of info.paras) {
      const tb = D.topBlock(d, p);
      if (!tb) continue;
      first = Math.min(first, tb.i);
      lastDirty = Math.max(lastDirty, tb.i);
    }
    if (info.tx && info.tx.tbls) for (const t of info.tx.tbls) { const ti = D.tblInfo(d, t); if (ti) { const tb = ti.parents.length ? blocks.indexOf(ti.parents[0].tbl) : blocks.indexOf(t); if (tb >= 0) { first = Math.min(first, tb); lastDirty = Math.max(lastDirty, tb); } } }
    if (first === Infinity) return LY.render();
    /* list labels / note numbers changing after the dirty range forbid resync */
    const newLabels = D.computeLists(d);
    let labelsChangedAfter = -1;
    const oldL = LY.lastLabels || new Map();
    for (let i = 0; i < blocks.length; i++) {
      const b = blocks[i];
      if (b.t !== 'p') continue;
      const a = newLabels.get(b.id), o = oldL.get(b.id);
      if ((a && a.text) !== (o && o.text)) { labelsChangedAfter = i; first = Math.min(first, i); }
    }
    lastDirty = Math.max(lastDirty, labelsChangedAfter);
    /* find the page holding the block before `first` (keep-with-next and widow rules may pull it) */
    const startBi = Math.max(0, first - 1);
    const target = blocks[startBi];
    let fromPage = 0;
    for (let k = LY.pages.length - 1; k >= 0; k--) {
      const ps = LY.pages[k].start;
      if (!ps) continue;
      const bi = blocks.indexOf(ps.block);
      if (bi >= 0 && bi <= startBi && (bi < startBi || !ps.sub)) { fromPage = k; break; }
      if (ps.block === target) { fromPage = k; break; }
    }
    /* step back one more page if the start block is mid-paragraph */
    if (fromPage > 0 && LY.pages[fromPage].start.sub) fromPage--;
    const ld = lastDirty;
    LY.resyncOK = (oldPage, bi) => bi > ld;
    R.clearCaches();
    LY.paginate(fromPage);
    LY.resyncOK = null;
    LY.lastBlocks = blocks.slice();
    LY.lastLabels = LY.labels;
    LY.lastNotes = LY.noteNums;
  };

  /* ================= sizing & zoom ================= */
  LY.fitSizer = function () {
    if (!LY.root) return;
    const z = LY.zoom;
    const r = LY.root;
    r.style.transform = z === 1 ? '' : `scale(${z})`;
    const w = r.offsetWidth, hh = r.offsetHeight;
    const sw = Math.max(LY.scroller.clientWidth, w * z + 20);
    LY.sizer.style.width = sw + 'px';
    LY.sizer.style.height = hh * z + 24 + 'px';
    r.style.left = Math.max(10, (sw - w * z) / 2) + 'px';
  };
  LY.setZoom = function (z) {
    LY.zoom = L.clamp(z, 0.1, 5);
    if (LY.view === 'web') return LY.render();
    LY.fitSizer();
    L.bus.emit('zoom', LY.zoom);
  };
  /** zoom presets: 'pageWidth', 'textWidth', 'wholePage', 'twoPages' */
  LY.zoomFor = function (kind) {
    const d = doc();
    const s = d.sect;
    const avail = LY.scroller.clientWidth - 36, availH = LY.scroller.clientHeight - 24;
    const W = px(s.pgW), H = px(s.pgH);
    if (kind === 'pageWidth') return avail / W;
    if (kind === 'textWidth') return avail / px(s.pgW - s.ml - s.mr + 24);
    if (kind === 'wholePage') return Math.min(avail / W, availH / H);
    if (kind === 'twoPages') return Math.min((avail - 20) / (2 * W), availH / H);
    return 1;
  };
  LY.setMultiPage = function (n) {
    LY.multi = n || 0;
    LY.root.classList.toggle('multi', !!n);
    LY.root.style.width = n ? (px(doc().sect.pgW) + 16) * n + 'px' : '';
    LY.fitSizer();
  };

  /* ================= DOM ↔ model mapping ================= */
  /** fragments for paragraph id */
  LY.frags = (pid, scope) => Array.from((scope || LY.root).querySelectorAll(`.p[data-pid="${pid}"]`));
  /** map a DOM point to a model position {p, o} */
  LY.domToPos = function (node, off) {
    if (!node) return null;
    const d = doc();
    let n = node.nodeType === 3 ? node.parentNode : node;
    const frag = n.closest ? n.closest('.p[data-pid]') : null;
    if (!frag) return null;
    const p = D.byId(d, +frag.dataset.pid);
    if (!p) return null;
    const from = +frag.dataset.from, to = +frag.dataset.to;
    const clampO = (o) => Math.max(from, Math.min(to, o));
    /* text node inside mapped span */
    if (node.nodeType === 3) {
      const span = node.parentNode;
      if (span.dataset && span.dataset.o != null && span.contentEditable !== 'false') {
        let o = +span.dataset.o;
        for (const c of span.childNodes) { if (c === node) break; if (c.nodeType === 3) o += c.length; }
        return { p, o: clampO(o + Math.min(off, +span.dataset.n)) };
      }
      const z = span.closest ? span.closest('[data-z]') : null;
      if (z) return { p, o: clampO(+z.dataset.z) };
      const m = span.closest ? span.closest('[data-o]') : null;
      if (m) return { p, o: clampO(+m.dataset.o + (off > 0 && node.length && off >= node.length ? +m.dataset.n : 0)) };
      return { p, o: from };
    }
    /* element point: look at neighbouring children */
    const elNode = node;
    if (elNode.dataset && elNode.dataset.o != null && elNode !== frag) {
      if (elNode.contentEditable === 'false' || elNode.nodeName === 'IMG') return { p, o: clampO(+elNode.dataset.o + (off > 0 ? +elNode.dataset.n : 0)) };
      let o = +elNode.dataset.o;
      for (let i = 0; i < off && i < elNode.childNodes.length; i++) { const c = elNode.childNodes[i]; if (c.nodeType === 3) o += c.length; }
      return { p, o: clampO(o) };
    }
    const kids = elNode.childNodes;
    for (let i = off; i < kids.length; i++) {
      const k = kids[i];
      if (k.nodeType !== 1) continue;
      if (k.dataset.o != null) return { p, o: clampO(+k.dataset.o) };
      if (k.dataset.z != null) return { p, o: clampO(+k.dataset.z) };
      const inner = k.querySelector && k.querySelector('[data-o]');
      if (inner) return { p, o: clampO(+inner.dataset.o) };
    }
    for (let i = Math.min(off, kids.length) - 1; i >= 0; i--) {
      const k = kids[i];
      if (k.nodeType !== 1) continue;
      if (k.dataset.o != null) return { p, o: clampO(+k.dataset.o + +k.dataset.n) };
      if (k.dataset.z != null) return { p, o: clampO(+k.dataset.z) };
      const all = k.querySelectorAll ? k.querySelectorAll('[data-o]') : [];
      if (all.length) { const lst = all[all.length - 1]; return { p, o: clampO(+lst.dataset.o + +lst.dataset.n) }; }
    }
    return { p, o: elNode === frag && off > 0 ? to : from };
  };
  /** choose the fragment for (p,o): prefer one where from <= o < to, else the one ending at o */
  LY.fragFor = function (p, o, scope) {
    const fr = LY.frags(p.id, scope);
    if (!fr.length) return null;
    let best = null;
    for (const f of fr) {
      const a = +f.dataset.from, b = +f.dataset.to;
      if (o >= a && o < b) return f;
      if (o === b) best = f;
      if (!best && o < a) best = f;
    }
    return best || fr[fr.length - 1];
  };
  /** model position → DOM point {node, offset} */
  LY.posToDom = function (pos, scope) {
    if (!pos || !pos.p) return null;
    const frag = LY.fragFor(pos.p, pos.o, scope);
    if (!frag) return null;
    const pc = frag.firstChild;
    const o = pos.o;
    let candidate = null;
    for (const k of pc.childNodes) {
      if (k.nodeType !== 1 || k.dataset.o == null) continue;
      const a = +k.dataset.o, n = +k.dataset.n;
      const isText = k.contentEditable !== 'false' && k.nodeName !== 'IMG' && k.nodeName !== 'BR' && k.firstChild && k.firstChild.nodeType === 3;
      if (isText) {
        if (o > a && o < a + n) return { node: k.firstChild, offset: o - a };
        if (o === a + n) candidate = { node: k.firstChild, offset: n };
        if (o === a) { if (candidate && candidate.node.parentNode !== k) return candidate; return { node: k.firstChild, offset: 0 }; }
      } else {
        if (o === a) { if (candidate) return candidate; const i = Array.prototype.indexOf.call(pc.childNodes, k); return { node: pc, offset: i }; }
        if (o === a + n) { const i = Array.prototype.indexOf.call(pc.childNodes, k); candidate = { node: pc, offset: i + 1 }; }
      }
    }
    if (candidate) return candidate;
    /* empty paragraph or no mapped content: place before the filler <br> (after a list label) */
    const br = pc.querySelector('br.fill');
    if (br) return { node: pc, offset: Array.prototype.indexOf.call(pc.childNodes, br) };
    const z = pc.querySelector('.zw');
    if (z) return { node: z.firstChild, offset: 0 };
    return { node: pc, offset: pc.childNodes.length };
  };
  /** page index containing (p,o) */
  LY.pageOf = function (pos) {
    const f = pos && LY.fragFor(pos.p, pos.o);
    if (!f) return -1;
    const pg = f.closest('.pg');
    return pg ? +pg.dataset.pg : -1;
  };
  LY.pageEl = (i) => (LY.pages[i] ? LY.pages[i].el : null);
})();
