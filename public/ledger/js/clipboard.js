/* Ledger — clipboard: copy / cut / paste, Paste Special, drag-and-drop moves, exchange with other programs. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, LY = L.layout, O = L.ops, F = L.formula, C = L.calc, NF = L.numfmt;
  const CL = (L.clip = {});
  const G = () => L.grid;
  let clip = null, marqueeTimer = null;
  CL.get = () => clip;
  CL.clearMarquee = () => { clip = clip && !clip.cut ? Object.assign(clip, { live: false }) : null; G().state.marquee = null; clearInterval(marqueeTimer); G().paint(); };

  /* ------------------------------------------------------------ copy */
  /** snapshot of a range (values, formulas, styles, merges, comments) */
  function snapshot(sh, rg, visibleOnly) {
    const rows = [];
    const w = rg.c2 - rg.c1 + 1, sourceRows = [], rowIndex = new Map();
    for (let r = rg.r1; r <= rg.r2; r++) {
      if (visibleOnly && sh.rows[r]?.filtered) continue;
      rowIndex.set(r, rows.length);
      sourceRows.push(r);
      const row = [];
      for (let j = 0; j < w; j++) {
        const c = rg.c1 + j;
        const cell = sh.get(r, c);
        const st = LY.styleIdOf(sh, r, c, cell);
        let copy = cell ? Object.assign({}, cell) : null;
        if (copy) { delete copy._gnode; delete copy.ast; delete copy._fsrc; delete copy.computing; if (copy.dirty) copy.v = C.cellValue(sh, r, c); delete copy.dirty; }
        if (!copy && st) copy = { v: null, s: st };
        else if (copy && copy.s == null && st) copy.s = st;
        if (copy && copy.am) {
          /* members of an array copy as values */
          const v = copy.v; copy = { v, s: copy.s };
        }
        row.push(copy);
      }
      rows.push(row);
    }
    function project(range) {
      const hit = O.intersection(range, rg);
      if (!hit) return null;
      while (hit.r1 <= hit.r2 && !rowIndex.has(hit.r1)) hit.r1++;
      while (hit.r2 >= hit.r1 && !rowIndex.has(hit.r2)) hit.r2--;
      if (hit.r1 > hit.r2) return null;
      return { r1: rowIndex.get(hit.r1), c1: hit.c1 - rg.c1, r2: rowIndex.get(hit.r2), c2: hit.c2 - rg.c1 };
    }
    return {
      rows, sourceRows, h: rows.length, w,
      merges: sh.merges.filter(m => rowIndex.has(m.r1) && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2)
        .map(project).filter(m => m && (m.r1 !== m.r2 || m.c1 !== m.c2)),
      comments: Array.from(sh.comments.values()).filter(cm => rowIndex.has(cm.r) && cm.c >= rg.c1 && cm.c <= rg.c2)
        .map(cm => ({ ...cm, dr: rowIndex.get(cm.r), dc: cm.c - rg.c1 })),
      dv: sh.dv.flatMap(d => {
        const ranges = d.ranges.map(project).filter(Boolean);
        return ranges.length ? [{ ...JSON.parse(JSON.stringify(d)), ranges }] : [];
      }),
      links: sh.links.flatMap(link => { const ref = project(link.ref); return ref ? [{ ...link, ref }] : []; }),
      colW: Array.from({ length: Math.min(w, 256) }, (_, j) => (sh.cols[rg.c1 + j] ? Object.assign({}, sh.cols[rg.c1 + j]) : null)),
      rowH: sourceRows.map(r => sh.rows[r]?.ht),
    };
  }
  /** shown text of every cell (for other programs) */
  function asText(sh, rg, cut) {
    const out = [];
    const clipped = O.clip(sh, rg);
    for (let r = rg.r1; r <= clipped.r2; r++) {
      if (cut ? sh.rows[r]?.hidden && rg.r1 !== rg.r2 : sh.rows[r]?.filtered) continue;
      const row = [];
      for (let c = rg.c1; c <= clipped.c2; c++) {
        const cell = sh.get(r, c);
        if (!cell || (cell.v == null && !cell.dirty)) { row.push(''); continue; }
        const v = cell.dirty ? C.cellValue(sh, r, c) : cell.v;
        const st = LY.styleOf(sh, r, c, cell);
        let t = v == null ? '' : NF.text(st.nf || 'General', v, { date1904: sh.wb.date1904 });
        if (/[\t\n"]/.test(t)) t = '"' + t.replace(/"/g, '""') + '"';
        row.push(t);
      }
      out.push(row.join('\t'));
    }
    return out.join('\r\n') + '\r\n';
  }
  function asHTML(sh, rg, id, snap) {
    const esc = L.esc;
    const clipped = O.clip(sh, rg);
    const g = LY.geo(sh);
    const spans = new Map(snap.merges.map(m => [M.key(snap.sourceRows[m.r1], rg.c1 + m.c1), m]));
    let html = `<html><head><meta charset="utf-8"><meta name="ledger-clip" content="${id}"></head><body><table style="border-collapse:collapse">`;
    for (let i = 0; i < snap.h; i++) {
      const r = snap.sourceRows[i];
      if (r > clipped.r2) break;
      html += `<tr style="height:${sh.rows[r]?.hidden ? M.ptToPx(sh.rows[r].ht ?? M.defaultRowPt(sh)) : g.rows.size(r)}px">`;
      for (let c = rg.c1; c <= clipped.c2; c++) {
        const m = LY.mergeAt(sh, r, c);
        const span = m && spans.get(M.key(m.r1, m.c1));
        if (span && (span.r1 !== i || span.c1 !== c - rg.c1)) continue;
        const cell = sh.get(r, c);
        const st = LY.styleOf(sh, r, c, cell);
        const v = cell ? (cell.dirty ? C.cellValue(sh, r, c) : cell.v) : null;
        const t = v == null ? '' : NF.text(st.nf || 'General', v, { date1904: sh.wb.date1904 });
        const css = [];
        const f = st.font || {};
        if (f.b) css.push('font-weight:bold');
        if (f.i) css.push('font-style:italic');
        if (f.u) css.push('text-decoration:underline');
        if (f.color) css.push('color:' + M.colorHex(sh.wb, f.color, '#000000'));
        if (f.name) css.push(`font-family:${L.cssString(f.name)}`);
        if (f.sz) css.push('font-size:' + f.sz + 'pt');
        if (st.fill && st.fill.pattern === 'solid') css.push('background:' + M.colorHex(sh.wb, st.fill.fg || st.fill.bg, '#FFFFFF'));
        if (st.align && st.align.h) css.push('text-align:' + ({ general: 'left', centerContinuous: 'center' }[st.align.h] || st.align.h));
        else if (typeof v === 'number') css.push('text-align:right');
        const b = st.border || {};
        for (const [k, side] of [['t', 'top'], ['b', 'bottom'], ['l', 'left'], ['r', 'right']]) if (b[k]) css.push(`border-${side}:${L.render.borderWidth(b[k])}px ${b[k].style === 'double' ? 'double' : /dash|dot|hair/.test(b[k].style) ? 'dashed' : 'solid'} ${M.colorHex(sh.wb, b[k].color, '#000000')}`);
        const mergeAttrs = span ? ` rowspan="${span.r2 - span.r1 + 1}" colspan="${span.c2 - span.c1 + 1}"` : '';
        const num = typeof v === 'number' ? ` x:num="${v}"` : '';
        html += `<td${mergeAttrs}${num} style="${esc(css.join(';'))};width:${g.cols.size(c)}px">${esc(t).replace(/\n/g, '<br>')}</td>`;
      }
      html += '</tr>';
    }
    return html + '</table></body></html>';
  }
  /** copy (or cut) the selection; e: the ClipboardEvent when available */
  CL.copy = function (e, cut) {
    const g = G(), sh = g.sheet(), s = g.sel();
    if (g.selectedObject() && L.drawing) { L.drawing.copy(g.selectedObject(), cut); return; }
    if (L.drawing) L.drawing.clipboard = null;
    if (s.ranges.length > 1) {
      /* several areas only copy when they line up */
      const rs = s.ranges;
      const sameCols = rs.every((x) => x.c1 === rs[0].c1 && x.c2 === rs[0].c2), sameRows = rs.every((x) => x.r1 === rs[0].r1 && x.r2 === rs[0].r2);
      if (cut || !(sameCols || sameRows)) { L.ui.msg('That command cannot be used on multiple selections.', { icon: 'warn' }); return; }
    }
    const rg0 = s.ranges[s.active];
    const rg = { r1: rg0.r1, c1: rg0.c1, r2: rg0.r2 >= M.MAXR - 1 ? Math.max(rg0.r1, sh.maxR) : rg0.r2, c2: rg0.c2 >= M.MAXC - 1 ? Math.max(rg0.c1, sh.maxC) : rg0.c2 };
    if (cut && g.guardProtect(null, true)) return;
    const id = 'L' + Date.now().toString(36);
    clip = { id, wb: sh.wb, sh, range: rg, full: rg0, cut: !!cut, snap: snapshot(sh, rg, !cut), live: true };
    clip.text = asText(sh, rg, cut);
    if (e && e.clipboardData) {
      e.clipboardData.setData('text/plain', clip.text);
      try { e.clipboardData.setData('text/html', asHTML(sh, rg, id, clip.snap)); } catch (err) { /* ignore */ }
    } else if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(clip.text).catch(() => {});
    g.state.marquee = { sh, ranges: [rg], phase: 0 };
    clearInterval(marqueeTimer);
    marqueeTimer = setInterval(() => { if (!g.state.marquee) { clearInterval(marqueeTimer); return; } g.state.marquee.phase = (g.state.marquee.phase + 2) % 8; g.paint(); }, 120);
    L.bus.emit('clipboard', clip);
    if (L.app.addOfficeClip) L.app.addOfficeClip(clip);
    L.app.status(L.t('Select destination and press ENTER or choose Paste'));
    g.paint();
  };

  /* ------------------------------------------------------------ paste */
  /** paste from a ClipboardEvent (keyboard) */
  CL.paste = function (e) {
    if (L.drawing && L.drawing.clipboard) { L.drawing.pasteObject(); return; }
    const dt = e.clipboardData;
    const html = dt ? dt.getData('text/html') : '';
    const text = dt ? dt.getData('text/plain') : '';
    const m = /<meta name="ledger-clip" content="([^"]+)"/.exec(html || '');
    if (clip && ((m && m[1] === clip.id) || (!html && text && text.replace(/\r\n/g, '\n') === clip.text.replace(/\r\n/g, '\n')) || (!text && !html))) { CL.pasteInternal({}); return; }
    /* images */
    if (dt) for (const it of dt.items || []) if (it.kind === 'file' && /^image\//.test(it.type)) { const f = it.getAsFile(); if (f && L.app.insertPictureFile) { L.app.insertPictureFile(f); return; } }
    if (html && /<table/i.test(html)) { pasteHTML(html); return; }
    if (text) { pasteText(text); return; }
    if (clip) CL.pasteInternal({});
  };
  /** Edit ▸ Paste from the menu: internal clipboard, else the system clipboard if readable */
  CL.pasteCommand = async function (opts) {
    if (L.drawing && L.drawing.clipboard) { L.drawing.pasteObject(); return; }
    if (clip) { CL.pasteInternal(opts || {}); return; }
    try {
      if (navigator.clipboard && navigator.clipboard.readText) { const t = await navigator.clipboard.readText(); if (t) { pasteText(t); return; } }
    } catch (err) { /* blocked */ }
    L.ui.msg('Use Ctrl+V to paste from other programs.', { icon: 'info' });
  };
  /** paste destination range for a block of h×w at the selection (tiles when the selection is a multiple) */
  function destFor(h, w) {
    const g = G(), s = g.sel();
    const rg = s.ranges[s.active];
    const sh = g.sheet();
    const selH = rg.r2 - rg.r1 + 1, selW = rg.c2 - rg.c1 + 1;
    if ((selH > 1 || selW > 1) && selH % h === 0 && selW % w === 0 && !(rg.r2 >= M.MAXR - 1 && h < M.MAXR) && !(rg.c2 >= M.MAXC - 1 && w < M.MAXC)) return { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, tile: true };
    void sh;
    return { r1: rg.r1, c1: rg.c1, r2: Math.min(M.MAXR - 1, rg.r1 + h - 1), c2: Math.min(M.MAXC - 1, rg.c1 + w - 1) };
  }
  CL.set = (c) => { clip = c; };
  CL.canPaste = () => !!clip || !!(L.drawing && L.drawing.clipboard);
  /**
   * Paste the internal clipboard. opts: {what:'all'|'formulas'|'values'|'formats'|'comments'|'validation'|'noBorders'|'colWidths'|'formulasNF'|'valuesNF',
   *   op:'none'|'add'|'sub'|'mul'|'div', skipBlanks, transpose, link}
   */
  CL.pasteInternal = function (opts) {
    opts = opts || {};
    const g = G(), sh = g.sheet();
    if (!clip) return;
    if (g.guardProtect(null, true)) return;
    const what = opts.what || 'all';
    if (clip.cut && what === 'all' && !opts.transpose && !opts.op) {
      const s = g.sel();
      const rg = s.ranges[s.active];
      const to = { r1: rg.r1, c1: rg.c1, r2: rg.r1 + clip.range.r2 - clip.range.r1, c2: rg.c1 + clip.range.c2 - clip.range.c1 };
      CL.moveRange(clip.sh, clip.range, to, false, sh);
      clip = null; g.state.marquee = null;
      g.selectRange(to);
      return;
    }
    const snap = clip.snap;
    if (!snap.h) return;
    let H = opts.transpose ? snap.w : snap.h, W = opts.transpose ? snap.h : snap.w;
    const dest = destFor(H, W);
    if (O.cutsArray(sh, [dest])) { L.ui.msg('You cannot change part of an array.', { icon: 'warn' }); return; }
    if (what === 'all' && !dest.tile && sh.merges.some((m) => M.rangesOverlap(m, dest) && !(m.r1 >= dest.r1 && m.r2 <= dest.r2 && m.c1 >= dest.c1 && m.c2 <= dest.c2))) { L.ui.msg('Cannot change part of a merged cell.', { icon: 'warn' }); return; }
    const srcWb = clip.wb;
    O.tx(sh.wb, opts.what || opts.op || opts.transpose ? 'Paste Special' : L.N('Paste'), () => {
      if (what === 'colWidths') {
        for (let j = 0; j < W && j < snap.colW.length; j++) { const o = snap.colW[opts.transpose ? 0 : j]; if (o && o.w != null) O.setColWidth(sh, [dest.c1 + j], o.w); }
        return;
      }
      for (let r = dest.r1; r <= dest.r2; r++) for (let c = dest.c1; c <= dest.c2; c++) {
        const i = (r - dest.r1) % H, j = (c - dest.c1) % W;
        const si = opts.transpose ? j : i, sj = opts.transpose ? i : j;
        const src = snap.rows[si] && snap.rows[si][sj];
        const srcR = snap.sourceRows[si], srcC = clip.range.c1 + sj;
        if (opts.skipBlanks && (!src || (src.v == null && src.f == null))) continue;
        const old = sh.get(r, c);
        let cell = old ? Object.assign({}, old) : { v: null };
        delete cell._gnode; delete cell.ast; delete cell._fsrc;
        const styleId = src && src.s != null ? (srcWb === sh.wb ? src.s : sh.wb.styles.add(JSON.parse(JSON.stringify(srcWb.styles.get(src.s))))) : 0;
        if (opts.link) {
          const ref = (clip.sh !== sh ? F.quoteSheet(clip.sh.name) + '!' : '') + F.cellName(srcR, srcC, true, true);
          cell = Object.assign(cell, { f: ref, v: null, edited: true });
          delete cell.af; delete cell.am;
          O.put(sh, r, c, cell);
          continue;
        }
        const takeFormula = what === 'all' || what === 'formulas' || what === 'noBorders' || what === 'formulasNF';
        const takeValue = what === 'values' || what === 'valuesNF';
        const takeFormat = what === 'all' || what === 'formats' || what === 'noBorders';
        const takeNF = what === 'formulasNF' || what === 'valuesNF';
        if (takeFormula || takeValue) {
          let v = src ? src.v : null, f = src && takeFormula ? src.f : null;
          if (f != null && f !== undefined) {
            f = F.translate(f, r - srcR, c - srcC);
            if (srcWb !== sh.wb) f = f; /* references stay relative */
          }
          if (opts.op && opts.op !== 'none') {
            /* arithmetic onto existing values */
            const cur = old ? (old.dirty ? C.cellValue(sh, r, c) : old.v) : null;
            const sym = { add: '+', sub: '-', mul: '*', div: '/' }[opts.op];
            if (f != null) { f = old && old.f != null ? '(' + old.f + ')' + sym + '(' + f + ')' : (cur == null ? 0 : typeof cur === 'number' ? cur : JSON.stringify(cur)) + sym + '(' + f + ')'; }
            else if (typeof v === 'number' || v == null) {
              if (old && old.f != null) f = '(' + old.f + ')' + sym + (v || 0);
              else if (typeof cur === 'number' || cur == null) { const a = cur || 0, b = v || 0; v = opts.op === 'add' ? a + b : opts.op === 'sub' ? a - b : opts.op === 'mul' ? a * b : b === 0 ? M.ERR.DIV0 : a / b; }
              else if (v == null) v = cur;
            }
          }
          delete cell.f; delete cell.af; delete cell.am; delete cell.rt; delete cell.dt;
          cell.v = v == null ? null : v;
          if (f != null && f !== undefined) { cell.f = f; cell.edited = true; }
          if (src && src.rt && f == null && what === 'all') cell.rt = src.rt;
          if (takeNF && src && src.s != null) { const nf = srcWb.styles.get(src.s).nf; cell.s = sh.wb.styles.derive(cell.s || 0, { nf }); }
        }
        if (takeFormat) {
          if (what === 'noBorders') { const st = sh.wb.styles.get(styleId); const keepB = old && old.s ? sh.wb.styles.get(old.s).border : null; cell.s = sh.wb.styles.add(Object.assign({}, st, { border: keepB })); }
          else cell.s = styleId;
          if (!cell.s) delete cell.s;
        }
        if (what === 'comments' || what === 'validation') continue;
        const empty = cell.v == null && cell.f == null && !cell.s;
        O.put(sh, r, c, empty ? null : cell);
      }
      function place(ref, r, c) {
        const block = opts.transpose ? { r1: ref.c1, c1: ref.r1, r2: ref.c2, c2: ref.r2 } : ref;
        return O.intersection(O.offsetRange(block, r, c), dest);
      }
      const merges = [], validations = [], links = [];
      for (let r = dest.r1; r <= dest.r2; r += H) for (let c = dest.c1; c <= dest.c2; c += W) {
        if (what === 'all' || what === 'formats') {
          for (const ref of snap.merges) { const m = place(ref, r, c); if (m) merges.push(m); }
        }
        if (what === 'all' || what === 'comments') for (const cm of snap.comments) {
          const rr = r + (opts.transpose ? cm.dc : cm.dr), cc = c + (opts.transpose ? cm.dr : cm.dc);
          if (M.rangeContains(dest, rr, cc)) O.setComment(sh, rr, cc, { ...L.threads.copy(cm, sh.wb), r: rr, c: cc });
        }
        if (what === 'all' || what === 'validation') for (const d of snap.dv) {
          const ranges = d.ranges.map(x => place(x, r, c)).filter(Boolean);
          if (ranges.length) validations.push({ ...L.sheetExtensions.copy(d), ranges });
        }
        if (what === 'all') for (const link of snap.links) {
          const ref = place(link.ref, r, c);
          if (ref) links.push({ ...link, ref });
        }
        if (what === 'all' && !opts.transpose && clip.full?.c2 === M.MAXC - 1 && clip.full.c1 === 0 && c === dest.c1) {
          snap.rowH.forEach((ht, i) => { if (ht != null && r + i <= dest.r2) O.setRowHeight(sh, [r + i], ht); });
        }
      }
      if (merges.length || what === 'all') {
        const keep = sh.merges.filter(m => what === 'all' ? !M.rangesOverlap(m, dest) : !merges.some(a => M.rangesOverlap(m, a)));
        if (merges.length || keep.length !== sh.merges.length) O.setMerges(sh, keep.concat(merges));
      }
      if (what === 'all' || what === 'validation') {
        if (!opts.skipBlanks) O.removeRangeRules(sh, dest, 'dv');
        if (validations.length) O.setDV(sh, sh.dv.concat(validations));
      }
      if (what === 'all' && !opts.link) {
        const keep = sh.links.flatMap(link => O.subtractRange(link.ref, dest).map(ref => ({ ...link, ref })));
        if (links.length || sh.links.some(link => M.rangesOverlap(link.ref, dest))) O.setLinks(sh, keep.concat(links));
      }
    });
    if (dest.r2 >= dest.r1) g.selectRange({ r1: dest.r1, c1: dest.c1, r2: dest.r2, c2: dest.c2 }, null, { noScroll: true });
  };
  /* ------------------------------------------------------------ foreign data */
  function parseTSV(text) {
    const rows = [];
    let row = [], cur = '', q = false;
    text = text.replace(/\r\n?/g, '\n');
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) { if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch; continue; }
      if (ch === '"' && cur === '') { q = true; continue; }
      if (ch === '\t') { row.push(cur); cur = ''; continue; }
      if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; continue; }
      cur += ch;
    }
    if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
    return rows;
  }
  CL.parseTSV = parseTSV;
  function pasteText(text) {
    const g = G(), sh = g.sheet();
    const rows = parseTSV(text);
    if (!rows.length) return;
    const W = Math.max(...rows.map((r) => r.length));
    const dest = destFor(rows.length, W);
    if (O.cutsArray(sh, [dest])) { L.ui.msg('You cannot change part of an array.', { icon: 'warn' }); return; }
    O.tx(sh.wb, L.N('Paste'), () => {
      for (let r = dest.r1; r <= dest.r2; r++) for (let c = dest.c1; c <= dest.c2; c++) {
        const t = (rows[(r - dest.r1) % rows.length] || [])[(c - dest.c1) % W];
        if (t == null) continue;
        try { O.enter(sh, r, c, t); } catch (e) { O.setValue(sh, r, c, t); }
      }
    });
    g.selectRange({ r1: dest.r1, c1: dest.c1, r2: dest.r2, c2: dest.c2 }, null, { noScroll: true });
  }
  CL.pasteText = pasteText;
  CL.pasteHTML = (html) => pasteHTML(html);
  /** HTML tables from Excel, browsers, Word: values plus basic formatting */
  function pasteHTML(html) {
    const g = G(), sh = g.sheet(), wb = sh.wb;
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const table = doc.querySelector('table');
    if (!table) return;
    const grid = [];
    const spans = [];
    const rules = {};
    for (const st of doc.querySelectorAll('style')) {
      const re = /\.([\w-]+)\s*\{([^}]*)\}/g; let m;
      while ((m = re.exec(st.textContent))) rules[m[1]] = m[2];
    }
    Array.from(table.rows).forEach((tr, i) => {
      grid[i] = grid[i] || [];
      let j = 0;
      for (const td of tr.cells) {
        while (grid[i][j] !== undefined) j++;
        const rs = +td.rowSpan || 1, cs = +td.colSpan || 1;
        const css = (td.className && rules[td.className] ? rules[td.className] + ';' : '') + (td.getAttribute('style') || '');
        const num = td.getAttribute('x:num');
        const text = (td.innerText || td.textContent || '').replace(/ /g, ' ').replace(/\n$/, '');
        grid[i][j] = { text, num: num != null && num !== '' ? +num : null, css, td };
        for (let a = 0; a < rs; a++) for (let b = 0; b < cs; b++) if (a || b) { grid[i + a] = grid[i + a] || []; grid[i + a][j + b] = null; }
        if (rs > 1 || cs > 1) spans.push({ r1: i, c1: j, r2: i + rs - 1, c2: j + cs - 1 });
        j += cs;
      }
    });
    const H = grid.length, W = Math.max(...grid.map((r) => r.length));
    const dest = destFor(H, W);
    const cssStyle = (css) => {
      const d = {};
      const get = (k) => { const m = new RegExp('(?:^|;)\\s*' + k + '\\s*:\\s*([^;]+)', 'i').exec(css); return m ? m[1].trim() : null; };
      const font = {};
      const fw = get('font-weight'); if (fw && (fw === 'bold' || +fw >= 600)) font.b = true;
      if (/italic/i.test(get('font-style') || '')) font.i = true;
      if (/underline/i.test(get('text-decoration') || '')) font.u = 'single';
      const col = cssColor(get('color')); if (col && col !== '#000000') font.color = M.rgb(col);
      const ff = get('font-family'); if (ff) font.name = ff.split(',')[0].replace(/["']/g, '').trim();
      const fs = get('font-size'); if (fs) { const m = /([\d.]+)\s*(pt|px)/.exec(fs); if (m) font.sz = m[2] === 'px' ? Math.round(+m[1] * 0.75 * 2) / 2 : +m[1]; }
      if (Object.keys(font).length) d.font = font;
      const bg = cssColor(get('background') || get('background-color'));
      if (bg && bg !== '#FFFFFF') d.fill = { pattern: 'solid', fg: M.rgb(bg) };
      const ta = get('text-align'); if (ta && /left|right|center|justify/.test(ta)) d.align = { h: ta.replace('justify', 'justify') };
      const nf = get('mso-number-format'); if (nf) d.nf = nf.replace(/^["']|["']$/g, '').replace(/\\(.)/g, '$1');
      return Object.keys(d).length ? d : null;
    };
    O.tx(wb, L.N('Paste'), () => {
      for (let r = dest.r1; r <= dest.r2; r++) for (let c = dest.c1; c <= dest.c2; c++) {
        const it = (grid[(r - dest.r1) % H] || [])[(c - dest.c1) % W];
        if (it === null) continue;
        if (!it) { O.put(sh, r, c, null); continue; }
        try { if (it.num != null && !isNaN(it.num)) O.setValue(sh, r, c, it.num); else O.enter(sh, r, c, it.text); } catch (e) { O.setValue(sh, r, c, it.text); }
        const d = cssStyle(it.css);
        if (d) { const cell = Object.assign({}, sh.get(r, c) || { v: null }); cell.s = wb.styles.derive(cell.s || 0, d); O.put(sh, r, c, cell); }
      }
      if (spans.length) O.setMerges(sh, sh.merges.concat(spans.map((m) => ({ r1: dest.r1 + m.r1, c1: dest.c1 + m.c1, r2: dest.r1 + m.r2, c2: dest.c1 + m.c2 }))));
    });
    g.selectRange({ r1: dest.r1, c1: dest.c1, r2: dest.r2, c2: dest.c2 }, null, { noScroll: true });
  }
  function cssColor(v) {
    if (!v) return null;
    v = v.trim();
    let m = /^#([0-9a-f]{3}|[0-9a-f]{6})\b/i.exec(v);
    if (m) { let x = m[1]; if (x.length === 3) x = x.split('').map((ch) => ch + ch).join(''); return '#' + x.toUpperCase(); }
    m = /rgba?\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/i.exec(v);
    if (m) return '#' + [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase();
    const named = { black: '#000000', white: '#FFFFFF', red: '#FF0000', green: '#008000', blue: '#0000FF', yellow: '#FFFF00', gray: '#808080', grey: '#808080', silver: '#C0C0C0', navy: '#000080', maroon: '#800000', purple: '#800080', teal: '#008080', olive: '#808000', orange: '#FFA500' };
    return named[v.toLowerCase()] || null;
  }

  /* ------------------------------------------------------------ moving */
  /**
   * Move (cut + paste, or drag) src on sheet `sh` to `to` (on `toSheet`, default same). copy: Ctrl-drag copies instead.
   */
  /** whether cells of range `to` outside `except` hold a value or a formula (formatting alone doesn't count) */
  CL.holdsData = function (sh, to, except) {
    for (let r = to.r1; r <= Math.min(to.r2, sh.maxR); r++) {
      const row = sh.rows[r];
      if (row) for (let c = to.c1; c <= Math.min(to.c2, row.cells.length - 1); c++) {
        const cl = row.cells[c];
        if (cl && (cl.v != null || cl.f != null) && !(except && M.rangeContains(except, r, c))) return true;
      }
    }
    return false;
  };
  CL.moveRange = function (sh, src, to, copy, toSheet) {
    toSheet = toSheet || sh;
    const wb = sh.wb;
    const dr = to.r1 - src.r1, dc = to.c1 - src.c1;
    if (copy) {
      clip = { id: 'drag', wb, sh, range: src, cut: false, snap: snapshot(sh, src, true), text: '' };
      const g = G();
      const keep = g.sel();
      const s = g.sel(); s.ranges = [Object.assign({}, to)]; s.active = 0; s.r = to.r1; s.c = to.c1;
      CL.pasteInternal({});
      clip = null;
      void keep;
      return;
    }
    if (O.cutsArray(sh, [src]) || O.cutsArray(toSheet, [to])) { L.ui.msg('You cannot change part of an array.', { icon: 'warn' }); return; }
    O.tx(wb, L.N('Move'), () => {
      const snap = snapshot(sh, src);
      const sameSheet = toSheet === sh;
      const before = O.snapSheet(sh), before2 = sameSheet ? null : O.snapSheet(toSheet);
      /* formulas anywhere that point at the overwritten destination become #REF!, those pointing at src follow */
      const adjust = O.moveReferences(sh, src, dr, dc, toSheet, to);
      for (const s of wb.sheets) s.rows.forEach((row, r) => { if (!row) return; row.cells.forEach((cell, c) => {
        if (!cell || cell.f == null || (s === sh && M.rangeContains(src, r, c))) return;
        const ast = C.astOf(cell); if (!ast) return;
        const res = adjust(ast, s.name);
        if (res.changed) { if (s !== sh && s !== toSheet) wb.undo.note(s, r, c); cell.f = F.toText(res.ast, { store: true }); cell.ast = res.ast; cell._fsrc = cell.f; cell.dirty = true; }
      }); });
      O.moveCellState(sh, src, dr, dc, toSheet, to, adjust);
      /* clear source, write destination */
      for (let r = src.r1; r <= Math.min(src.r2, sh.maxR); r++) { const row = sh.rows[r]; if (row) for (let c = src.c1; c <= Math.min(src.c2, row.cells.length - 1); c++) delete row.cells[c]; }
      for (let i = 0; i < snap.h; i++) for (let j = 0; j < snap.w; j++) {
        const cell = snap.rows[i][j];
        const r = to.r1 + i, c = to.c1 + j;
        if (r >= M.MAXR || c >= M.MAXC) continue;
        const row = toSheet.rowObj(r);
        if (!cell) { delete row.cells[c]; continue; }
        const n = Object.assign({}, cell);
        if (n.f != null) {
          /* references inside the moved block move with it; others keep pointing at the same cells */
          let ast; try { ast = F.parse(n.f); } catch (e) { ast = null; }
          if (ast) {
            const inner = F.adjustMove(ast, Object.assign({ sheet: sh.name }, src), dr, dc, sh.name, toSheet.name);
            let a = inner.ast;
            if (!sameSheet) a = F.map(a, (x) => ((x.t === 'ref' || x.t === 'area') && x.sheet == null && !(inner.changed && false) ? Object.assign({}, x, { sheet: sh.name }) : null));
            n.f = F.toText(a, { store: true });
          }
          n.dirty = true;
        }
        if (n.af) n.af = { r1: n.af.r1 + dr, c1: n.af.c1 + dc, r2: n.af.r2 + dr, c2: n.af.c2 + dc };
        if (n.am) n.am = { r: n.am.r + dr, c: n.am.c + dc };
        row.cells[c] = n;
        if (r > toSheet.maxR) toSheet.maxR = r;
        if (c > toSheet.maxC) toSheet.maxC = c;
      }
      sh.merges = sh.merges.filter((m) => !(m.r1 >= src.r1 && m.r2 <= src.r2 && m.c1 >= src.c1 && m.c2 <= src.c2));
      toSheet.merges = toSheet.merges.filter((m) => !M.rangesOverlap(m, to)).concat(snap.merges.map((m) => ({ r1: to.r1 + m.r1, c1: to.c1 + m.c1, r2: to.r1 + m.r2, c2: to.c1 + m.c2 })));
      for (const cm of snap.comments) sh.comments.delete(M.key(cm.r, cm.c));
      for (const [key, cm] of toSheet.comments) if (M.rangeContains(to, cm.r, cm.c)) toSheet.comments.delete(key);
      for (const cm of snap.comments) { const r = to.r1 + cm.dr, c = to.c1 + cm.dc; toSheet.comments.set(M.key(r, c), Object.assign({}, cm, { r, c })); }
      sh.recalcBounds();
      LY.invalidate(sh); LY.touchMerges(sh); LY.invalidate(toSheet); LY.touchMerges(toSheet);
      const after = O.snapSheet(sh), after2 = sameSheet ? null : O.snapSheet(toSheet);
      wb.undo.op(() => { O.restoreSheet(sh, before); if (!sameSheet) O.restoreSheet(toSheet, before2); }, () => { O.restoreSheet(sh, after); if (!sameSheet) O.restoreSheet(toSheet, after2); });
      O.markStructural();
    });
  };
})(typeof window !== 'undefined' ? window : globalThis);
