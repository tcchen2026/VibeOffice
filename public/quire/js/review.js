/* Quire — reviewing: track changes (accept / reject / navigate), comments with balloons and the
 * reviewing pane, reviewer colours, and Compare and Merge Documents.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const RV = (L.review = {});
  const doc = () => D.doc;
  const app = () => L.app;
  const { h } = L;

  /* ================= tracking ================= */
  RV.toggleTracking = function () {
    const d = doc();
    d.settings.track = !d.settings.track;
    d.dirty = true;
    app().status(d.settings.track ? 'Track Changes is on.' : 'Track Changes is off.');
    if (d.settings.track && app().opts.toolbars.reviewing === 'auto') app().updateToolbars();
    app().updateStatus();
    L.ui.refresh();
  };

  /* ================= revisions ================= */
  const revOf = (it) => { const r = it.rPr || {}; return r.ins ? ['ins', r.ins] : r.del ? ['del', r.del] : r.chg ? ['chg', r.chg] : null; };
  /** every revision in story order: [{type, stamp, p, a, b, text}] (adjacent runs of the same revision merge) */
  RV.list = function (story) {
    const d = doc();
    D.reindex(d);
    const st = story || d.main;
    const out = [];
    for (const p of st.paras || []) {
      let o = 0;
      let cur = null;
      for (const it of p.runs) {
        const n = D.ilen(it);
        const rv = n ? revOf(it) : null;
        if (rv && cur && cur.type === rv[0] && cur.stamp.author === rv[1].author && cur.b === o) { cur.b = o + n; cur.text += it.t === 'text' ? it.text : ' '; }
        else if (rv) { cur = { type: rv[0], stamp: rv[1], p, a: o, b: o + n, text: it.t === 'text' ? it.text : it.t === 'tab' ? '\t' : '[object]' }; out.push(cur); }
        else if (n) cur = null;
        o += n;
      }
      if (p.pPr.chg) out.push({ type: 'pchg', stamp: p.pPr.chg, p, a: 0, b: D.plen(p), text: 'Paragraph formatting' });
      if (p.mark && (p.mark.ins || p.mark.del)) out.push({ type: p.mark.ins ? 'mins' : 'mdel', stamp: p.mark.ins || p.mark.del, p, a: D.plen(p), b: D.plen(p), text: '¶' });
    }
    return out;
  };
  RV.authors = function () {
    const d = doc();
    const set = new Set();
    for (const st of D.stories(d)) for (const r of RV.list(st)) set.add(r.stamp.author || 'Unknown');
    for (const k in d.comments) set.add(d.comments[k].author || 'Unknown');
    return Array.from(set);
  };
  RV.revisionAt = function (pos) {
    pos = pos || (E.sel && E.sel.f);
    if (!pos) return null;
    const st = D.storyOf(doc(), pos.p);
    return RV.list(st).find((r) => r.p === pos.p && ((r.a <= pos.o && pos.o < r.b) || (r.a < pos.o && pos.o <= r.b) || (r.type === 'mins' || r.type === 'mdel') && pos.o === r.a)) || null;
  };
  /** apply accept/reject to one revision record (inside a transaction) */
  function resolve(r, accept) {
    const p = r.p;
    D.touch(p);
    if (r.type === 'pchg') { if (accept) delete p.pPr.chg; else p.pPr = Object.assign({}, r.stamp.old || {}); return; }
    if (r.type === 'mins' || r.type === 'mdel') {
      const joining = (r.type === 'mins' && !accept) || (r.type === 'mdel' && accept);
      delete p.mark.ins; delete p.mark.del;
      if (!Object.keys(p.mark).length) delete p.mark;
      if (joining) joinNoTrack(p);
      return;
    }
    const i0 = D.splitAt(p, r.a), i1 = D.splitAt(p, r.b);
    const keep = [];
    for (let i = 0; i < p.runs.length; i++) {
      const it = p.runs[i];
      if (i < i0 || i >= i1 || !D.ilen(it)) { keep.push(it); continue; }
      const rp = it.rPr || {};
      if (r.type === 'ins') { if (accept) { delete rp.ins; keep.push(it); } /* rejected insertion disappears */ }
      else if (r.type === 'del') { if (!accept) { delete rp.del; keep.push(it); } }
      else if (r.type === 'chg') { if (accept) delete rp.chg; else { const old = rp.chg && rp.chg.old ? L.clone(rp.chg.old) : {}; for (const k of ['ins', 'del', 'link']) if (rp[k]) old[k] = rp[k]; it.rPr = old; } keep.push(it); }
    }
    p.runs = keep;
    D.normalize(p);
  }
  function joinNoTrack(p) {
    const d = doc();
    const trk = d.settings.track;
    d.settings.track = false;
    try { O.joinNext(p); } finally { d.settings.track = trk; }
  }
  RV.acceptReject = function (accept, scope) {
    const d = doc();
    if (!E.sel) return;
    /* a comment under the caret: Reject deletes it */
    if (!accept && scope === 'selection') { const c = RV.commentAt(); const rv0 = RV.revisionAt(); if (c && !rv0) { RV.deleteComment(c); return; } }
    let targets;
    if (scope === 'all') { targets = []; for (const st of D.stories(d)) targets.push(...RV.list(st)); }
    else if (!E.collapsed()) {
      const [a, b] = E.range();
      const st = D.storyOf(d, a.p);
      targets = RV.list(st).filter((r) => D.cmp(d, D.pos(r.p, r.b), a) >= 0 && D.cmp(d, D.pos(r.p, r.a), b) <= 0 && !(r.p === b.p && r.a >= b.o && r.a !== r.b) && !(r.p === a.p && r.b <= a.o && r.a !== r.b));
    } else { const r = RV.revisionAt(); targets = r ? [r] : []; }
    if (!targets.length) {
      if (scope === 'selection') { RV.step(1); app().status(accept ? 'There are no tracked changes at the selection.' : 'There are no tracked changes or comments at the selection.'); }
      return;
    }
    const trk = d.settings.track;
    E.edit(accept ? 'Accept Change' : 'Reject Change', () => {
      d.settings.track = false;
      try {
        /* work from the end so offsets stay valid */
        const order = targets.slice().sort((x, y) => (x.p === y.p ? y.a - x.a : (y.p._o || 0) - (x.p._o || 0)));
        const marks = order.filter((r) => r.type === 'mins' || r.type === 'mdel');
        for (const r of order) if (!(r.type === 'mins' || r.type === 'mdel')) resolve(r, accept);
        for (const r of marks) if (D.info(d, r.p)) resolve(r, accept);
      } finally { d.settings.track = trk; }
      const p = targets[0].p;
      return D.info(d, p) ? D.pos(p, Math.min(targets[0].a, D.plen(p))) : E.sel;
    });
    if (scope === 'selection' && E.collapsed()) RV.step(1, true);
  };
  /** move to the next/previous revision or comment */
  RV.step = function (dir, quiet) {
    const d = doc();
    if (!E.sel) return;
    const st = E.story().kind === 'main' ? E.story() : d.main;
    const items = RV.list(st).map((r) => ({ p: r.p, a: r.a, b: r.b }));
    D.reindex(d);
    for (const p of st.paras || []) { let o = 0; for (const it of p.runs) { if (it.t === 'cs' && d.comments[it.id]) items.push({ p, a: o, b: o, cmt: it.id }); o += D.ilen(it); } }
    items.sort((x, y) => (x.p === y.p ? x.a - y.a : x.p._o - y.p._o));
    const cur = E.sel.f;
    const at = (x) => D.cmp(d, D.pos(x.p, x.a), cur);
    const hit = dir > 0 ? items.find((x) => at(x) > 0 || (E.collapsed() && at(x) === 0 && false)) : items.slice().reverse().find((x) => D.cmp(d, D.pos(x.p, x.b), E.range()[0]) < 0);
    if (!hit) { if (!quiet) app().status(dir > 0 ? 'Word reached the end of the document.' : 'Word reached the beginning of the document.'); return; }
    E.setSel(D.pos(hit.p, hit.a), D.pos(hit.p, hit.b));
    if (hit.cmt) highlightBalloon(hit.cmt);
  };

  /* ================= comments ================= */
  RV.insertComment = function () {
    const d = doc();
    if (!E.sel) return;
    if (D.storyOf(d, E.sel.f.p).kind === 'cmt') return;
    let [a, b] = E.range();
    if (D.eqPos(a, b)) { const w = O.wordAt(a); if (w) [a, b] = w; }
    let id = 0;
    while (d.comments[id] != null) id++;
    id = String(id);
    let target = null;
    E.edit('Insert Comment', () => {
      D.touchKey(d, 'comments');
      if (!d.styles.CommentText) { D.touchKey(d, 'styles'); d.styles.CommentText = D.builtinStyles().CommentText; d.styles.CommentReference = d.styles.CommentReference || D.builtinStyles().CommentReference; D.stylesChanged(); }
      const p = D.para([], { style: 'CommentText' });
      d.comments[id] = { kind: 'cmt', id, author: O.author(), initials: O.initials(), date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), blocks: [p] };
      D.touch(b.p);
      const ib = D.splitAt(b.p, b.o);
      b.p.runs.splice(ib, 0, D.item('ce', { id }));
      D.touch(a.p);
      const ia = D.splitAt(a.p, a.o);
      a.p.runs.splice(ia, 0, D.item('cs', { id }));
      target = p;
      d._idxDirty = true;
      return E.sel;
    });
    LY.opts.showComments = true;
    if (LY.view !== 'print' || LY.opts.balloons === false) { RV.togglePane(true); }
    setTimeout(() => { D.reindex(d); if (target) { E.setSel(D.pos(target, 0)); E.focus(); highlightBalloon(id); } if (LY.view !== 'print' || LY.opts.balloons === false) RV.editComment(id); }, 0);
  };
  RV.commentAt = function (pos) {
    pos = pos || (E.sel && E.sel.f);
    if (!pos) return null;
    const d = doc();
    const st = D.storyOf(d, pos.p);
    if (st.kind === 'cmt') return st.id;
    const open = new Set();
    for (const p of st.paras || []) {
      let o = 0;
      for (const it of p.runs) {
        if (p === pos.p && o >= pos.o && !(it.t === 'ce' && o === pos.o)) return open.size ? Array.from(open).pop() : null;
        if (it.t === 'cs') open.add(it.id); else if (it.t === 'ce') open.delete(it.id);
        o += D.ilen(it);
      }
      if (p === pos.p) return open.size ? Array.from(open).pop() : null;
    }
    return null;
  };
  RV.deleteComment = function (id) {
    const d = doc();
    E.edit('Delete Comment', () => {
      D.touchKey(d, 'comments');
      const kids = Object.keys(d.comments).filter((k) => d.comments[k].parent === id);
      for (const k of [id].concat(kids)) delete d.comments[k];
      D.walk(d.main, (b) => { if (b.t === 'p' && b.runs.some((it) => (it.t === 'cs' || it.t === 'ce') && (it.id === id || kids.includes(it.id)))) { D.touch(b); b.runs = b.runs.filter((it) => !((it.t === 'cs' || it.t === 'ce') && (it.id === id || kids.includes(it.id)))); } });
      d._idxDirty = true;
      return E.sel && D.storyOf(d, E.sel.f.p) && D.storyOf(d, E.sel.f.p).kind === 'cmt' ? D.pos(D.firstPara(d.main.blocks), 0) : E.sel;
    });
  };
  RV.deleteAllComments = function () {
    const d = doc();
    E.edit('Delete All Comments', () => {
      D.touchKey(d, 'comments');
      d.comments = {};
      D.walk(d.main, (b) => { if (b.t === 'p' && b.runs.some((it) => it.t === 'cs' || it.t === 'ce')) { D.touch(b); b.runs = b.runs.filter((it) => it.t !== 'cs' && it.t !== 'ce'); } });
      d._idxDirty = true;
      return E.sel && D.storyOf(d, E.sel.f.p).kind === 'cmt' ? D.pos(D.firstPara(d.main.blocks), 0) : E.sel;
    });
  };
  /** edit a comment: in balloons the text is edited in place; otherwise a small editor dialog */
  RV.editComment = function (id) {
    const d = doc();
    const c = d.comments[id];
    if (!c) return;
    if (LY.view === 'print' && LY.opts.balloons !== false && LY.root.querySelector(`.balloon[data-cid="${id}"] .p`)) {
      D.reindex(d);
      const p = D.lastPara(c.blocks);
      E.setSel(D.pos(p, D.plen(p)));
      E.focus();
      highlightBalloon(id);
      return;
    }
    const ta = h('textarea', { rows: 6, style: 'width:100%' });
    ta.value = c.blocks.filter((b) => b.t === 'p').map((p) => D.plainText(p)).join('\n');
    L.ui.dialog({
      title: `Comment [${c.initials || ''}${+id + 1}] — ${c.author}`, width: 380, body: h('div', { class: 'col' }, ta), focus: ta,
      buttons: [{ label: 'OK', primary: true, onClick: () => {
        E.edit('Edit Comment', () => {
          D.touchKey(d, 'comments');
          const cc = d.comments[id];
          const style = (D.firstPara(cc.blocks) || { pPr: { style: 'CommentText' } }).pPr.style || 'CommentText';
          cc.blocks = ta.value.split('\n').map((line) => D.para(line ? [D.text(line)] : [], { style }));
          d._idxDirty = true;
          return E.sel && D.storyOf(d, E.sel.f.p).kind === 'cmt' ? D.pos(D.firstPara(d.main.blocks), 0) : E.sel;
        });
      } }, { label: 'Cancel' }],
    });
  };
  RV.replyComment = function (id) {
    const d = doc();
    const par = d.comments[id];
    if (!par) return;
    let nid = 0;
    while (d.comments[nid] != null) nid++;
    nid = String(nid);
    E.edit('Reply to Comment', () => {
      D.touchKey(d, 'comments');
      d.comments[nid] = { kind: 'cmt', id: nid, author: O.author(), initials: O.initials(), date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), blocks: [D.para([], { style: 'CommentText' })], parent: id };
      /* the reply shares the parent's range */
      D.walk(d.main, (b) => { if (b.t !== 'p') return; let ch = false; const out = []; for (const it of b.runs) { out.push(it); if (it.t === 'cs' && it.id === id) { out.push(D.item('cs', { id: nid })); ch = true; } if (it.t === 'ce' && it.id === id) { out.push(D.item('ce', { id: nid })); ch = true; } } if (ch) { D.touch(b); b.runs = out; } });
      d._idxDirty = true;
      return E.sel;
    });
    setTimeout(() => RV.editComment(nid), 0);
  };

  /* ================= balloons ================= */
  const fmtDate = (s) => { if (!s) return ''; const dt = new Date(s); return isNaN(dt) ? '' : dt.toLocaleDateString() + ' ' + dt.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }); };
  RV.commentNumber = function (id) {
    const d = doc();
    const ids = [];
    D.walk(d.main, (b) => { if (b.t === 'p') for (const it of b.runs) if (it.t === 'cs' && d.comments[it.id] && !ids.includes(it.id)) ids.push(it.id); });
    const i = ids.indexOf(id);
    return i < 0 ? +id + 1 : i + 1;
  };
  function highlightBalloon(id) {
    for (const b of LY.root.querySelectorAll('.balloon.on')) b.classList.remove('on');
    const b = LY.root.querySelector(`.balloon[data-cid="${id}"]`);
    if (b) b.classList.add('on');
    for (const r of document.querySelectorAll('#revpane .rp-item.on')) r.classList.remove('on');
  }
  RV.highlightBalloon = highlightBalloon;
  function renderBalloons() {
    for (const b of LY.root.querySelectorAll(':scope > .balloons')) b.remove();
    const d = doc();
    const cids = Object.keys(d.comments);
    const wantPad = LY.view === 'print' && LY.opts.balloons !== false && LY.opts.showComments !== false && LY.opts.markup !== 'final' && LY.opts.markup !== 'original' && cids.length > 0;
    LY.root.classList.toggle('with-balloons', wantPad);
    /* tooltips on markers for every view */
    for (const m of LY.root.querySelectorAll('.cmk-s')) {
      const c = d.comments[m.dataset.cid];
      if (c) m.setAttribute('data-tip', `${c.author}, ${fmtDate(c.date)}:\n${c.blocks.filter((b) => b.t === 'p').map((p) => D.plainText(p)).join(' ').slice(0, 200)}`);
    }
    if (!wantPad) { LY.fitSizer(); return; }
    const z = LY.zoom;
    const ctxBase = LY.ctx({});
    const placed = new Set();
    for (const pg of LY.pages) {
      const marks = Array.from(pg.el.querySelectorAll('.cmk-s')).filter((m) => d.comments[m.dataset.cid] && !placed.has(m.dataset.cid));
      if (!marks.length) continue;
      const pr = pg.el.getBoundingClientRect();
      const box = h('div', { class: 'balloons' });
      box.style.left = pg.el.offsetLeft + pg.el.offsetWidth + 12 + 'px';
      box.style.top = pg.el.offsetTop + 'px';
      box.style.height = pg.el.offsetHeight + 'px';
      LY.root.appendChild(box);
      let bottom = 0;
      const ordered = [];
      for (const m of marks) {
        const id = m.dataset.cid;
        placed.add(id);
        const em = pg.el.querySelector(`.cmk-e[data-cid="${id}"]`) || LY.root.querySelector(`.cmk-e[data-cid="${id}"]`);
        const er = (em || m).getBoundingClientRect();
        ordered.push({ id, y: (m.getBoundingClientRect().top - pr.top) / z, x: (m.getBoundingClientRect().left - pr.left) / z, lx: (er.right - pr.left) / z, ly: (er.bottom - pr.top) / z });
        for (const k of cids) if (d.comments[k].parent === id && !placed.has(k)) { placed.add(k); ordered.push({ id: k, y: ordered[ordered.length - 1].y, x: ordered[ordered.length - 1].x, reply: true }); }
      }
      for (const o of ordered) {
        const c = d.comments[o.id];
        const color = R.authorColor(c.author);
        const bl = h('div', { class: 'balloon' + (o.reply ? ' reply' : ''), 'data-cid': o.id });
        bl.style.borderColor = color;
        bl.style.background = L.color.mix ? L.color.mix(color, '#ffffff', 0.85) : '#fff8dc';
        const head = h('div', { class: 'bh', text: `Comment [${c.initials || ''}${RV.commentNumber(o.id)}]${o.reply ? ' (reply)' : ''}: ` });
        head.contentEditable = 'false';
        head.setAttribute('data-tip', `${c.author}, ${fmtDate(c.date)}`);
        const body = h('div', { class: 'bb' });
        bl.append(head, body);
        box.appendChild(bl);
        const fr = { el: body, limit: Infinity, top: 0, empty: true, kind: 'cmt', width: 190 };
        LY.flow(fr, c.blocks, { bi: 0, sub: null }, c.blocks.length, Object.assign({}, ctxBase, { story: c, openCmts: null, view: 'print', showMarks: false }));
        const y = Math.max(o.y - 2, bottom + 4);
        bl.style.top = y + 'px';
        bottom = y + bl.offsetHeight;
        /* dashed connector from the comment mark to the balloon */
        const line = h('div', { class: 'balloon-line' });
        line.contentEditable = 'false';
        const lx = o.lx != null ? o.lx : o.x, ly = o.ly != null ? o.ly : o.y + 14;
        if (Math.abs(ly - (y + 8)) > 2) {
          /* text line → right edge at the line's foot, then a short riser to the balloon */
          const rise = h('div', { class: 'balloon-line' });
          rise.contentEditable = 'false';
          const top = Math.min(ly, y + 8), hgt = Math.abs(ly - (y + 8));
          rise.style.cssText = `left:${pg.el.offsetLeft + pg.el.offsetWidth + 6}px;top:${pg.el.offsetTop + top}px;height:${hgt}px;width:0;border-top:0;border-left:1px dashed ${color}`;
          LY.root.appendChild(rise);
          rise.classList.add('bl-' + o.id);
        }
        line.style.cssText = `left:${pg.el.offsetLeft + lx}px;top:${pg.el.offsetTop + ly}px;width:${Math.max(4, pg.el.offsetWidth - lx + 6)}px;border-top-color:${color}`;
        const stub = h('div', { class: 'balloon-line' });
        stub.contentEditable = 'false';
        stub.style.cssText = `left:${pg.el.offsetLeft + pg.el.offsetWidth + 6}px;top:${pg.el.offsetTop + y + 8}px;width:6px;border-top-color:${color}`;
        LY.root.appendChild(stub);
        stub.classList.add('bl-' + o.id);
        LY.root.appendChild(line);
        line.classList.add('bl-' + o.id);
        bl.addEventListener('pointerdown', () => highlightBalloon(o.id));
      }
    }
    for (const l of LY.root.querySelectorAll(':scope > .balloon-line')) if (!l.isConnected) l.remove();
    LY.fitSizer();
  }
  RV.renderBalloons = renderBalloons;
  L.bus.on('layout-done', () => {
    for (const l of LY.root.querySelectorAll(':scope > .balloon-line')) l.remove();
    renderBalloons();
    if (!L.$('#revpane').hidden) RV.renderPane();
  });

  /* ================= reviewing pane ================= */
  RV.togglePane = function (on) {
    const pane = L.$('#revpane');
    const show = on == null ? pane.hidden : on;
    pane.hidden = !show;
    if (show) RV.renderPane();
    LY.fitSizer();
    L.ui.refresh();
  };
  const TYPE_LABEL = { ins: 'Inserted', del: 'Deleted', chg: 'Formatted', pchg: 'Formatted', mins: 'Inserted', mdel: 'Deleted' };
  RV.renderPane = function () {
    const pane = L.$('#revpane');
    if (pane.hidden) return;
    const d = doc();
    pane.textContent = '';
    const close = h('button', { class: 'tp-nav', type: 'button', 'aria-label': 'Close Reviewing Pane', html: '<svg width="8" height="8"><path d="M1 1l6 6M7 1L1 7" stroke="#000" stroke-width="1.4"/></svg>', onclick: () => RV.togglePane(false) });
    pane.appendChild(h('div', { class: 'rp-head' }, h('span', { text: 'Main document changes and comments' }), close));
    const items = [];
    for (const r of RV.list(d.main)) items.push({ kind: 'rev', r, p: r.p, o: r.a });
    D.reindex(d);
    for (const p of d.main.paras || []) { let o = 0; for (const it of p.runs) { if (it.t === 'cs' && d.comments[it.id]) items.push({ kind: 'cmt', id: it.id, p, o }); o += D.ilen(it); } }
    items.sort((x, y) => (x.p === y.p ? x.o - y.o : x.p._o - y.p._o));
    if (!items.length) { pane.appendChild(h('div', { class: 'rp-item tp-note', text: 'There are no tracked changes or comments in this document.' })); return; }
    for (const it of items) {
      let head, body, color;
      if (it.kind === 'rev') {
        const r = it.r;
        color = R.authorColor(r.stamp.author);
        head = `${TYPE_LABEL[r.type]} ${r.stamp.author || ''} ${fmtDate(r.stamp.date)}`;
        body = r.type === 'chg' ? describeFmt(r) : r.text;
      } else {
        const c = d.comments[it.id];
        color = R.authorColor(c.author);
        head = `Comment [${c.initials || ''}${RV.commentNumber(it.id)}] ${c.author} ${fmtDate(c.date)}`;
        body = c.blocks.filter((b) => b.t === 'p').map((p) => D.plainText(p)).join('\n');
      }
      const el = h('div', { class: 'rp-item' }, h('div', { class: 'rp-h', text: head, style: `color:${color}` }), h('div', { class: 'rp-b', text: body }));
      el.addEventListener('click', () => { const a = D.pos(it.p, it.kind === 'rev' ? it.r.a : it.o); const b = D.pos(it.p, it.kind === 'rev' ? it.r.b : it.o); app().gotoPos(a, b); if (it.kind === 'cmt') highlightBalloon(it.id); });
      if (it.kind === 'cmt') el.addEventListener('dblclick', () => RV.editComment(it.id));
      pane.appendChild(el);
    }
  };
  function describeFmt(r) {
    const old = (r.stamp && r.stamp.old) || {};
    const now = (D.itemAfter(r.p, r.a) || {}).rPr || {};
    const out = [];
    const keys = ['b', 'i', 'u', 'strike', 'font', 'sz', 'color', 'hl', 'caps', 'smallCaps', 'vert'];
    const name = { b: 'Bold', i: 'Italic', u: 'Underline', strike: 'Strikethrough', font: 'Font', sz: 'Font size', color: 'Font color', hl: 'Highlight', caps: 'All caps', smallCaps: 'Small caps', vert: 'Position' };
    for (const k of keys) if (JSON.stringify(old[k]) !== JSON.stringify(now[k])) { const v = now[k]; out.push(v === true ? name[k] : v === false || v == null || v === 'none' ? 'Not ' + name[k] : `${name[k]}: ${k === 'sz' ? v + ' pt' : v}`); }
    return 'Formatted: ' + (out.join(', ') || 'Font');
  }
  RV.reviewersDialog = function () {
    const list = RV.authors();
    const body = h('div', { class: 'col' }, h('div', { text: 'Reviewers in this document:' }));
    if (!list.length) body.appendChild(h('div', { class: 'tp-note', text: 'No one has tracked changes or comments yet.' }));
    for (const a of list) body.appendChild(h('div', { style: 'display:flex;align-items:center;gap:6px' }, h('span', { style: `width:12px;height:12px;background:${R.authorColor(a)};border:1px solid #555;display:inline-block` }), h('span', { text: a })));
    L.ui.dialog({ title: 'Reviewers', body, width: 300, buttons: [{ label: 'Close', primary: true }] });
  };

  /* ================= Compare and Merge Documents ================= */
  RV.compareDialog = async function () {
    const files = await L.pickFiles('.docx,.docm,.dotx');
    const f = files[0];
    if (!f) return;
    L.ui.busy(true, 'Comparing documents...');
    try {
      const res = await L.docx.read(await L.readAsArrayBuffer(f));
      const other = res.doc;
      const n = RV.compareWith(other, other.props.lastModifiedBy || other.props.creator || f.name.replace(/\.\w+$/, ''));
      L.ui.busy(false);
      app().status(n ? `Marked ${n} difference${n === 1 ? '' : 's'} as tracked changes.` : 'The documents are identical.');
      if (n) { app().opts.toolbars.reviewing = true; app().updateToolbars(); }
    } catch (e) { L.ui.busy(false); console.error(e); L.ui.msg('The documents could not be compared. ' + (e.message || ''), { icon: 'error' }); }
  };
  const tokenize = (s) => s.match(/[\p{L}\p{N}_'’]+|\s+|[^\s\p{L}\p{N}_'’]/gu) || [];
  /** LCS diff of two arrays with an equality key; returns ops [{op:'eq'|'del'|'ins', a, b}] */
  function diff(A, B, key) {
    key = key || ((x) => x);
    const n = A.length, m = B.length;
    /* trim common prefix/suffix to keep the table small */
    let s = 0;
    while (s < n && s < m && key(A[s]) === key(B[s])) s++;
    let e = 0;
    while (e < n - s && e < m - s && key(A[n - 1 - e]) === key(B[m - 1 - e])) e++;
    const a = A.slice(s, n - e), b = B.slice(s, m - e);
    const ops = [];
    for (let i = 0; i < s; i++) ops.push({ op: 'eq', a: i, b: i });
    if (a.length * b.length > 4e6) {
      /* too big: treat as replace */
      for (let i = 0; i < a.length; i++) ops.push({ op: 'del', a: s + i });
      for (let j = 0; j < b.length; j++) ops.push({ op: 'ins', b: s + j });
    } else {
      const dp = Array.from({ length: a.length + 1 }, () => new Uint32Array(b.length + 1));
      for (let i = a.length - 1; i >= 0; i--) for (let j = b.length - 1; j >= 0; j--) dp[i][j] = key(a[i]) === key(b[j]) ? dp[i + 1][j + 1] + 1 : Math.max(dp[i + 1][j], dp[i][j + 1]);
      let i = 0, j = 0;
      while (i < a.length && j < b.length) {
        if (key(a[i]) === key(b[j])) { ops.push({ op: 'eq', a: s + i, b: s + j }); i++; j++; }
        else if (dp[i + 1][j] >= dp[i][j + 1]) { ops.push({ op: 'del', a: s + i }); i++; }
        else { ops.push({ op: 'ins', b: s + j }); j++; }
      }
      while (i < a.length) { ops.push({ op: 'del', a: s + i }); i++; }
      while (j < b.length) { ops.push({ op: 'ins', b: s + j }); j++; }
    }
    for (let k = 0; k < e; k++) ops.push({ op: 'eq', a: n - e + k, b: m - e + k });
    return ops;
  }
  RV.diff = diff;
  const blockKey = (b) => (b.t === 'p' ? 'p:' + D.ptext(b) : 't:' + JSON.stringify(b.rows.map((r) => r.cells.map((c) => c.blocks.map((x) => (x.t === 'p' ? D.ptext(x) : '#')).join('\n')))));
  /** mark the differences between the current document and `other` as tracked changes; returns the count */
  RV.compareWith = function (other, author) {
    const d = doc();
    const stamp = () => ({ author: author || 'Author', date: new Date().toISOString().replace(/\.\d+Z$/, 'Z'), id: D.nid() });
    let count = 0;
    E.edit('Compare Documents', () => {
      const trk = d.settings.track;
      d.settings.track = false;
      try {
        app().mergeNumbering(other, other.main.blocks);
        for (const k in other.styles) if (!d.styles[k]) { D.touchKey(d, 'styles'); d.styles[k] = other.styles[k]; }
        const A = d.main.blocks, B = other.main.blocks;
        const ops = diff(A, B, blockKey);
        /* group replace regions: consecutive del+ins runs pair up paragraph-by-paragraph */
        const cont = D.touchList(d.main);
        const out = [];
        let k = 0;
        while (k < ops.length) {
          const o = ops[k];
          if (o.op === 'eq') { out.push(A[o.a]); k++; continue; }
          const dels = [], inss = [];
          while (k < ops.length && ops[k].op !== 'eq') { if (ops[k].op === 'del') dels.push(A[ops[k].a]); else inss.push(B[ops[k].b]); k++; }
          const pairs = Math.min(dels.length, inss.length);
          for (let i = 0; i < Math.max(dels.length, inss.length); i++) {
            const x = dels[i], y = inss[i];
            if (i < pairs && x.t === 'p' && y.t === 'p') { count += diffPara(x, y, stamp()); out.push(x); continue; }
            if (i < pairs && x.t === 'tbl' && y.t === 'tbl' && sameShape(x, y)) { count += diffTable(x, y, stamp); out.push(x); continue; }
            if (x) { markDeleted(x, stamp()); out.push(x); count++; }
            if (y) { const c = D.cloneBlocks([y])[0]; markInserted(c, stamp()); out.push(c); count++; }
          }
        }
        cont.blocks = out;
        d._idxDirty = true;
      } finally { d.settings.track = trk; }
      return D.pos(D.firstPara(d.main.blocks), 0);
    });
    return count;
  };
  function sameShape(x, y) { return x.rows.length === y.rows.length && x.rows.every((r, i) => r.cells.length === y.rows[i].cells.length); }
  function diffTable(x, y, stamp) {
    let n = 0;
    D.touchTbl(x);
    x.rows.forEach((r, ri) => r.cells.forEach((c, ci) => {
      const yc = y.rows[ri].cells[ci];
      const ops = diff(c.blocks, yc.blocks, blockKey);
      const out = [];
      let k = 0;
      while (k < ops.length) {
        if (ops[k].op === 'eq') { out.push(c.blocks[ops[k].a]); k++; continue; }
        const dels = [], inss = [];
        while (k < ops.length && ops[k].op !== 'eq') { if (ops[k].op === 'del') dels.push(c.blocks[ops[k].a]); else inss.push(yc.blocks[ops[k].b]); k++; }
        for (let i = 0; i < Math.max(dels.length, inss.length); i++) {
          const a = dels[i], b = inss[i];
          if (a && b && a.t === 'p' && b.t === 'p') { n += diffPara(a, b, stamp()); out.push(a); continue; }
          if (a) { markDeleted(a, stamp()); out.push(a); n++; }
          if (b) { const cb = D.cloneBlocks([b])[0]; markInserted(cb, stamp()); out.push(cb); n++; }
        }
      }
      c.blocks = out;
    }));
    return n;
  }
  function markDeleted(b, st) {
    D.walk({ blocks: [b] }, (x) => { if (x.t === 'p') { D.touch(x); for (const it of x.runs) if (D.ilen(it)) it.rPr = Object.assign({}, it.rPr, { del: st }); x.mark = Object.assign({}, x.mark, { del: st }); } });
  }
  function markInserted(b, st) {
    D.walk({ blocks: [b] }, (x) => { if (x.t === 'p') { for (const it of x.runs) if (D.ilen(it)) it.rPr = Object.assign({}, it.rPr, { ins: st }); x.mark = Object.assign({}, x.mark, { ins: st }); } });
  }
  /** word-level diff of two paragraphs; edits x in place */
  function diffPara(x, y, st) {
    D.touch(x);
    const ta = tokenize(D.ptext(x)), tb = tokenize(D.ptext(y));
    const ops = diff(ta, tb);
    /* offsets of tokens */
    const offA = [], offB = [];
    let o = 0; for (const t of ta) { offA.push(o); o += t.length; } offA.push(o);
    o = 0; for (const t of tb) { offB.push(o); o += t.length; } offB.push(o);
    /* collect edits as {at (offset in x), delLen, insFrom, insTo (offsets in y)} */
    const edits = [];
    let k = 0;
    while (k < ops.length) {
      if (ops[k].op === 'eq') { k++; continue; }
      let a0 = null, a1 = null, b0 = null, b1 = null;
      const anchor = k > 0 ? offA[ops[k - 1].a + 1] : 0;
      while (k < ops.length && ops[k].op !== 'eq') {
        const q = ops[k];
        if (q.op === 'del') { if (a0 == null) a0 = offA[q.a]; a1 = offA[q.a + 1]; }
        else { if (b0 == null) b0 = offB[q.b]; b1 = offB[q.b + 1]; }
        k++;
      }
      edits.push({ at: a0 != null ? a0 : anchor, delEnd: a1, b0, b1 });
    }
    if (!edits.length) {
      /* same text: compare formatting of runs is out of scope; paragraph props may differ */
      if (JSON.stringify(x.pPr) !== JSON.stringify(y.pPr)) { x.pPr = Object.assign(L.clone(y.pPr), { chg: Object.assign({}, st, { old: L.clone(x.pPr) }) }); return 1; }
      return 0;
    }
    for (let i = edits.length - 1; i >= 0; i--) {
      const ed = edits[i];
      const insAt = ed.delEnd != null ? ed.delEnd : ed.at;
      if (ed.b0 != null) {
        /* inserted items copied from y keep their formatting */
        const items = D.sliceRuns(y, ed.b0, ed.b1).map((it) => Object.assign(it, { rPr: Object.assign({}, it.rPr, { ins: st }) }));
        const idx = D.splitAt(x, insAt);
        x.runs.splice(idx, 0, ...items);
      }
      if (ed.delEnd != null) {
        const i0 = D.splitAt(x, ed.at), i1 = D.splitAt(x, ed.delEnd);
        for (let j = i0; j < i1; j++) if (D.ilen(x.runs[j])) x.runs[j].rPr = Object.assign({}, x.runs[j].rPr, { del: st });
      }
    }
    D.normalize(x);
    return edits.length;
  }
})();
