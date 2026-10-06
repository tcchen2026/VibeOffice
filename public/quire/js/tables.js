/* Quire — tables: insertion, structure editing (rows, columns, merge, split), borders & shading, AutoFit,
 * sorting, formulas, conversion, Table AutoFormat styles, the Draw Table pencil and eraser, and
 * mouse resizing of columns and rows.
 */
(function () {
  'use strict';
  const L = window.L, D = L.D, O = L.O, E = L.ed, LY = L.layout, R = L.R;
  const T = (L.tables = {});
  const doc = () => D.doc;
  const A = () => L.app;

  /* ================= queries ================= */
  T.inTable = () => !!(E.sel && D.cellOf(doc(), E.sel.f.p));
  T.cur = () => (E.sel ? D.cellOf(doc(), E.sel.f.p) : null);
  const xsOf = (grid) => { const xs = [0]; for (const g of grid) xs.push(xs[xs.length - 1] + g); return xs; };
  /** the rectangle of selected cells: {tbl, r0, r1, c0, c1, cells:[cell], map} */
  T.range = function () {
    const d = doc();
    if (E.cellSel) { const cs = E.cellSel; return Object.assign({ map: R.tblMap(cs.tbl) }, cs); }
    const ci = T.cur();
    if (!ci) return null;
    const map = R.tblMap(ci.tbl);
    const m = map[ci.ri].find((x) => x.cell === ci.cell);
    /* a selection inside one table running over several cells of the same column: all rows in range */
    if (E.sel && !E.collapsed()) {
      const ca = D.cellOf(d, E.sel.a.p);
      if (ca && ca.tbl === ci.tbl && ca.cell !== ci.cell) {
        const ma = map[ca.ri].find((x) => x.cell === ca.cell);
        const r0 = Math.min(ca.ri, ci.ri), r1 = Math.max(ca.ri, ci.ri), c0 = Math.min(ma.c0, m.c0), c1 = Math.max(ma.c0 + ma.span, m.c0 + m.span) - 1;
        return cellsIn(ci.tbl, map, r0, r1, c0, c1);
      }
    }
    return { tbl: ci.tbl, map, r0: ci.ri, r1: ci.ri + (m.rowspan || 1) - 1, c0: m.c0, c1: m.c0 + m.span - 1, cells: [ci.cell] };
  };
  function cellsIn(tbl, map, r0, r1, c0, c1) {
    const cells = [];
    for (let r = r0; r <= r1; r++) for (const m of map[r]) if (!m.hidden && m.c0 + m.span - 1 >= c0 && m.c0 <= c1) cells.push(m.cell);
    return { tbl, map, r0, r1, c0, c1, cells };
  }
  T.wholeCellsSelected = function () {
    if (E.cellSel) return true;
    if (!E.sel || E.collapsed()) return false;
    const ci = T.cur();
    if (!ci) return false;
    const [a, b] = E.range();
    const f = D.firstPara(ci.cell.blocks), l = D.lastPara(ci.cell.blocks);
    return a.p === f && a.o === 0 && b.p === l && b.o >= D.plen(l);
  };
  T.isHeaderRow = () => { const c = T.cur(); return !!(c && c.row.trPr.header); };
  T.cellVAlign = () => { const c = T.cur(); return c ? c.cell.tcPr.vAlign || 'top' : 'top'; };
  /** width available for a new table at the caret (pt) */
  function availWidth(pos) {
    const d = doc();
    const ci = D.cellOf(d, pos.p);
    if (ci) {
      const map = R.tblMap(ci.tbl);
      const m = map[ci.ri].find((x) => x.cell === ci.cell);
      const { tp } = R.tblProps(d, ci.tbl);
      const w = ci.tbl.grid.slice(m.c0, m.c0 + m.span).reduce((a, b) => a + b, 0);
      return Math.max(36, w - (tp.cellMar.l || 0) - (tp.cellMar.r || 0));
    }
    const s = D.sectFor(d, pos.p);
    const cols = s.cols || { n: 1 };
    const body = s.pgW - s.ml - s.mr - (s.gutter || 0);
    if ((cols.n || 1) > 1) return (body - (cols.space || 36) * (cols.n - 1)) / cols.n;
    return body;
  }
  T.availWidth = availWidth;
  function fixWidths(t) {
    const map = R.tblMap(t);
    const xs = xsOf(t.grid);
    map.forEach((row) => { for (const m of row) m.cell.tcPr.w = L.round(xs[Math.min(xs.length - 1, m.c0 + m.span)] - xs[m.c0], 2); });
    if (t.tblPr.w && t.tblPr.w.type === 'dxa') t.tblPr.w = { type: 'dxa', v: L.round(xs[xs.length - 1], 2) };
  }
  T.fixWidths = fixWidths;
  const firstPos = (cell) => { const p = D.firstPara(cell.blocks); return D.pos(p, 0); };
  const emptyLike = (cell) => { const p = D.firstPara(cell.blocks) || D.para(); return D.para([], L.clone(p.pPr), L.clone(p.rPr)); };
  const newCellLike = (cell, span) => { const tc = L.clone(cell.tcPr); delete tc.vMerge; delete tc.span; if (span > 1) tc.span = span; return D.cell([emptyLike(cell)], tc); };

  /* ================= insert ================= */
  T.insert = function (nr, nc, opts) {
    opts = opts || {};
    if (!E.sel) return;
    nr = L.clamp(nr | 0, 1, 32767); nc = L.clamp(nc | 0, 1, 63);
    let firstCell = null;
    E.edit('Insert Table', () => {
      let pos = E.collapsed() ? E.sel.f : E.range()[0];
      const width = opts.width || availWidth(pos);
      const cur = pos.p;
      const t = D.simpleTable(nr, nc, opts.colWidth ? opts.colWidth * nc : width, { pPr: cur.pPr.style && !/^Heading/.test(cur.pPr.style) ? { style: cur.pPr.style } : {} });
      if (opts.style) { T.ensureStyle(opts.style); t.tblPr.style = opts.style; }
      if (opts.look) t.tblPr.look = Object.assign({}, opts.look);
      if (opts.autofit === 'window') t.tblPr.w = { type: 'pct', v: 5000 };
      if (opts.autofit === 'fixed' || opts.colWidth) { t.tblPr.w = { type: 'dxa', v: L.round(t.grid.reduce((a, b) => a + b, 0), 2) }; t.tblPr.layout = 'fixed'; }
      if (opts.worksheet) { t.rows[0].trPr.header = true; for (const c of t.rows[0].cells) c.tcPr.shd = { val: 'clear', fill: 'D9D9D9' }; }
      O.insertBlockAt(pos, t);
      firstCell = t.rows[0].cells[0];
      return firstPos(firstCell);
    });
    if (opts.autofit === 'contents') T.autofit('contents');
  };
  /* ================= cell navigation ================= */
  T.moveCell = function (dir) {
    const ci = T.cur();
    if (!ci) return;
    const t = ci.tbl;
    const order = [];
    const map = R.tblMap(t);
    map.forEach((row) => row.forEach((m) => { if (!m.hidden) order.push(m.cell); }));
    let i = order.indexOf(ci.cell) + dir;
    if (i >= order.length) {
      /* Tab in the last cell adds a row */
      T.insertRows(false, { count: 1, from: ci.ri });
      const nc = T.cur();
      if (nc) { const r = nc.tbl.rows[nc.tbl.rows.length - 1]; E.setSel(firstPos(r.cells[0])); }
      return;
    }
    if (i < 0) return;
    const cell = order[i];
    const f = D.firstPara(cell.blocks), l = D.lastPara(cell.blocks);
    E.setSel(D.pos(f, 0), D.pos(l, D.plen(l)));
  };
  T.select = function (kind) {
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl, map = rg.map;
    let r0 = rg.r0, r1 = rg.r1, c0 = rg.c0, c1 = rg.c1;
    if (kind === 'table') { r0 = 0; r1 = t.rows.length - 1; c0 = 0; c1 = t.grid.length - 1; }
    else if (kind === 'row') { c0 = 0; c1 = t.grid.length - 1; }
    else if (kind === 'col') { r0 = 0; r1 = t.rows.length - 1; }
    const sel = cellsIn(t, map, r0, r1, c0, c1);
    const first = sel.cells[0], last = sel.cells[sel.cells.length - 1];
    if (!first) return;
    const fp = D.firstPara(first.blocks), lp = D.lastPara(last.blocks);
    E.setSel(D.pos(fp, 0), D.pos(lp, D.plen(lp)), { noScroll: true });
    E.cellSel = Object.assign({}, sel);
    for (const td of LY.root.querySelectorAll('td.csel')) td.classList.remove('csel');
    for (const c of sel.cells) for (const td of LY.root.querySelectorAll(`td[data-cid="${c.id}"]`)) td.classList.add('csel');
    L.bus.emit('sel', E.sel);
  };

  /* ================= rows ================= */
  T.insertRows = function (above, opts) {
    opts = opts || {};
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl;
    const n = opts.count || rg.r1 - rg.r0 + 1;
    const refIdx = opts.from != null ? opts.from : above ? rg.r0 : rg.r1;
    let target = null;
    E.edit(above ? 'Insert Rows Above' : 'Insert Rows Below', () => {
      D.touchTbl(t);
      const map = R.tblMap(t);
      const ref = t.rows[refIdx];
      const at = above ? refIdx : refIdx + 1;
      const news = [];
      for (let k = 0; k < n; k++) {
        const cells = [];
        for (const m of map[refIdx]) {
          const nc = newCellLike(m.cell, m.span);
          /* inside a vertical merge the new row joins it */
          const belowM = map[refIdx + 1] && map[refIdx + 1].find((x) => x.c0 === m.c0);
          const inMerge = above ? m.cell.tcPr.vMerge === 'continue' : (m.cell.tcPr.vMerge === 'restart' || m.cell.tcPr.vMerge === 'continue') && belowM && belowM.cell.tcPr.vMerge === 'continue';
          if (inMerge) nc.tcPr.vMerge = 'continue';
          cells.push(nc);
        }
        const trPr = L.clone(ref.trPr);
        if (!above || refIdx > 0) { if (!(t.rows[at] && t.rows[at].trPr.header && trPr.header)) delete trPr.header; }
        news.push(D.row(cells, trPr));
      }
      t.rows.splice(at, 0, ...news);
      target = news[0].cells[0];
      return firstPos(target);
    });
  };
  T.deleteRows = function () {
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl;
    if (rg.r0 === 0 && rg.r1 >= t.rows.length - 1) { T.deleteTable(); return; }
    E.edit('Delete Rows', () => {
      D.touchTbl(t);
      const map = R.tblMap(t);
      /* a merge that starts in a deleted row and continues below restarts below */
      const below = map[rg.r1 + 1];
      if (below) for (const m of below) if (m.cell.tcPr.vMerge === 'continue') {
        let k = rg.r0 - 1, starts = false;
        while (k >= 0) { const x = map[k].find((y) => y.c0 === m.c0); if (!x) break; if (x.cell.tcPr.vMerge === 'restart') { starts = true; break; } if (x.cell.tcPr.vMerge !== 'continue') break; k--; }
        if (!starts || k < 0) m.cell.tcPr.vMerge = 'restart';
        const restart = map.slice(rg.r0, rg.r1 + 1).some((row) => row.some((y) => y.c0 === m.c0 && y.cell.tcPr.vMerge === 'restart'));
        if (restart) { m.cell.tcPr.vMerge = 'restart'; if (m.cell.blocks.every((b) => b.t === 'p' && !D.plen(b))) { const src = map[rg.r0].find((y) => y.c0 === m.c0); if (src && src.cell.tcPr.vMerge === 'restart') m.cell.blocks = src.cell.blocks; } }
      }
      t.rows.splice(rg.r0, rg.r1 - rg.r0 + 1);
      const r = t.rows[Math.min(rg.r0, t.rows.length - 1)];
      cleanMerges(t);
      return firstPos(r.cells.find((c) => c.tcPr.vMerge !== 'continue') || r.cells[0]);
    });
  };
  /** a lone 'continue' with no restart above becomes a restart; a restart with nothing below loses its merge */
  function cleanMerges(t) {
    const map = R.tblMap(t);
    for (let r = 0; r < map.length; r++) for (const m of map[r]) {
      const v = m.cell.tcPr.vMerge;
      if (v === 'continue') { const up = r > 0 && map[r - 1].find((x) => x.c0 === m.c0 && x.span === m.span); if (!up || !up.cell.tcPr.vMerge) m.cell.tcPr.vMerge = 'restart'; }
    }
    const map2 = R.tblMap(t);
    for (let r = 0; r < map2.length; r++) for (const m of map2[r]) if (m.cell.tcPr.vMerge === 'restart') { const dn = map2[r + 1] && map2[r + 1].find((x) => x.c0 === m.c0); if (!dn || dn.cell.tcPr.vMerge !== 'continue') delete m.cell.tcPr.vMerge; }
  }

  /* ================= columns ================= */
  T.insertCols = function (left, opts) {
    opts = opts || {};
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl;
    const n = opts.count || rg.c1 - rg.c0 + 1;
    const gi = left ? rg.c0 : rg.c1 + 1;
    let target = null;
    E.edit(left ? 'Insert Columns to the Left' : 'Insert Columns to the Right', () => {
      D.touchTbl(t);
      const map = R.tblMap(t);
      const widths = t.grid.slice(rg.c0, rg.c1 + 1);
      const newW = [];
      for (let k = 0; k < n; k++) newW.push(widths[k % widths.length] || 72);
      /* keep the table inside the text column when it already fills it */
      const total = t.grid.reduce((a, b) => a + b, 0);
      const avail = availWidth(E.sel.f);
      const grow = newW.reduce((a, b) => a + b, 0);
      if (total + grow > avail + 1 && total <= avail + 1) {
        const k = avail / (total + grow);
        t.grid = t.grid.map((g) => g * k);
        for (let i = 0; i < newW.length; i++) newW[i] *= k;
      }
      t.grid.splice(gi, 0, ...newW);
      map.forEach((row, ri) => {
        const r = t.rows[ri];
        const before = r.trPr.gridBefore || 0;
        if (gi < before) { r.trPr.gridBefore = before + n; return; }
        const spanning = row.find((m) => m.c0 < gi && m.c0 + m.span > gi);
        if (spanning) { spanning.cell.tcPr.span = spanning.span + n; return; }
        const ref = row.find((m) => (left ? m.c0 === gi : m.c0 + m.span === gi)) || row[left ? 0 : row.length - 1];
        const idx = row.findIndex((m) => m.c0 >= gi);
        const cells = [];
        for (let k = 0; k < n; k++) { const c = newCellLike(ref.cell, 1); if (ref.cell.tcPr.vMerge) c.tcPr.vMerge = ref.cell.tcPr.vMerge; cells.push(c); }
        r.cells.splice(idx < 0 ? r.cells.length : idx, 0, ...cells);
        if (ri === (left ? rg.r0 : rg.r0)) target = cells[0];
      });
      fixWidths(t);
      return target ? firstPos(target) : E.sel;
    });
  };
  T.deleteCols = function () {
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl;
    if (rg.c0 === 0 && rg.c1 >= t.grid.length - 1) { T.deleteTable(); return; }
    E.edit('Delete Columns', () => {
      D.touchTbl(t);
      const map = R.tblMap(t);
      const n = rg.c1 - rg.c0 + 1;
      map.forEach((row, ri) => {
        const r = t.rows[ri];
        for (const m of row) {
          const a = Math.max(m.c0, rg.c0), b = Math.min(m.c0 + m.span - 1, rg.c1);
          if (a > b) continue;
          const overlap = b - a + 1;
          if (overlap >= m.span) r.cells.splice(r.cells.indexOf(m.cell), 1);
          else m.cell.tcPr.span = m.span - overlap;
        }
        if ((r.trPr.gridBefore || 0) > rg.c0) r.trPr.gridBefore = Math.max(0, (r.trPr.gridBefore || 0) - n);
      });
      t.grid.splice(rg.c0, n);
      t.rows = t.rows.filter((r) => r.cells.length);
      fixWidths(t);
      cleanMerges(t);
      const ri = Math.min(rg.r0, t.rows.length - 1);
      const row = R.tblMap(t)[ri];
      const m = row.find((x) => x.c0 + x.span > rg.c0) || row[row.length - 1];
      return firstPos(m.cell);
    });
  };
  T.deleteTable = function () {
    const ci = T.cur();
    if (!ci) return;
    E.edit('Delete Table', () => O.removeBlock(ci.tbl));
  };
  /** Insert Cells dialog result: 'right' | 'down' | 'row' | 'col' */
  T.insertCells = function (mode) {
    if (mode === 'row') return T.insertRows(true);
    if (mode === 'col') return T.insertCols(true);
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl;
    E.edit('Insert Cells', () => {
      D.touchTbl(t);
      if (mode === 'right') {
        /* new cells shift existing ones right in each selected row */
        for (let r = rg.r0; r <= rg.r1; r++) {
          const row = t.rows[r];
          const map = R.tblMap(t)[r];
          const idx = map.findIndex((m) => m.c0 >= rg.c0);
          const ref = map[Math.max(0, idx)].cell;
          for (let k = 0; k <= rg.c1 - rg.c0; k++) row.cells.splice(Math.max(0, idx), 0, newCellLike(ref, 1));
        }
        const maxCols = Math.max(...t.rows.map((r) => (r.trPr.gridBefore || 0) + r.cells.reduce((s, c) => s + (c.tcPr.span || 1), 0)));
        while (t.grid.length < maxCols) t.grid.push(t.grid[t.grid.length - 1] || 72);
      } else {
        /* shift cells down: append a row at the end, then move contents down in the selected columns */
        const last = t.rows[t.rows.length - 1];
        const nrow = D.row(last.cells.map((c) => newCellLike(c, c.tcPr.span || 1)), L.clone(last.trPr));
        delete nrow.trPr.header;
        const nAdd = rg.r1 - rg.r0 + 1;
        for (let k = 0; k < nAdd; k++) t.rows.push(k ? D.row(last.cells.map((c) => newCellLike(c, c.tcPr.span || 1)), L.clone(nrow.trPr)) : nrow);
        const map = R.tblMap(t);
        for (let c = rg.c0; c <= rg.c1; c++) {
          for (let r = t.rows.length - 1; r >= rg.r0 + nAdd; r--) {
            const dst = map[r].find((m) => m.c0 === c), src = map[r - nAdd].find((m) => m.c0 === c);
            if (dst && src) dst.cell.blocks = src.cell.blocks;
          }
          for (let r = rg.r0; r < rg.r0 + nAdd; r++) { const m = map[r].find((x) => x.c0 === c); if (m) m.cell.blocks = [emptyLike(m.cell)]; }
        }
      }
      fixWidths(t);
      const m = R.tblMap(t)[rg.r0].find((x) => x.c0 >= rg.c0) || R.tblMap(t)[rg.r0][0];
      return firstPos(m.cell);
    });
  };
  T.deleteCells = function (mode) {
    if (mode === 'row') return T.deleteRows();
    if (mode === 'col') return T.deleteCols();
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl;
    E.edit('Delete Cells', () => {
      D.touchTbl(t);
      if (mode === 'left') {
        for (let r = rg.r0; r <= rg.r1; r++) {
          const row = t.rows[r];
          const map = R.tblMap(t)[r];
          const del = map.filter((m) => m.c0 >= rg.c0 && m.c0 + m.span - 1 <= rg.c1).map((m) => m.cell);
          row.cells = row.cells.filter((c) => !del.includes(c));
          if (!row.cells.length) row.cells.push(D.cell([D.para()], {}));
        }
      } else {
        const map = R.tblMap(t);
        const n = rg.r1 - rg.r0 + 1;
        for (let c = rg.c0; c <= rg.c1; c++) {
          for (let r = rg.r0; r < t.rows.length; r++) {
            const dst = map[r].find((m) => m.c0 === c), src = map[r + n] && map[r + n].find((m) => m.c0 === c);
            if (dst) dst.cell.blocks = src ? src.cell.blocks : [emptyLike(dst.cell)];
          }
        }
      }
      fixWidths(t);
      const row = R.tblMap(t)[Math.min(rg.r0, t.rows.length - 1)];
      return firstPos((row.find((m) => m.c0 >= rg.c0) || row[0]).cell);
    });
  };

  /* ================= merge & split ================= */
  T.merge = function () {
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl;
    let { r0, r1, c0, c1 } = rg;
    /* grow the rectangle until no cell crosses its edge */
    const map = R.tblMap(t);
    for (let guard = 0; guard < 20; guard++) {
      let changed = false;
      for (let r = r0; r <= r1; r++) for (const m of map[r]) {
        if (m.c0 + m.span - 1 < c0 || m.c0 > c1) continue;
        if (m.c0 < c0) { c0 = m.c0; changed = true; }
        if (m.c0 + m.span - 1 > c1) { c1 = m.c0 + m.span - 1; changed = true; }
        if (!m.hidden && m.rowspan > 1 && r + m.rowspan - 1 > r1) { r1 = r + m.rowspan - 1; changed = true; }
        if (m.hidden) { let k = r; while (k > 0 && map[k][map[k].indexOf(m)] && map[k].find((x) => x.c0 === m.c0 && x.cell.tcPr.vMerge === 'continue')) k--; if (k < r0) { r0 = k; changed = true; } }
      }
      if (!changed) break;
    }
    if (r0 === r1 && c0 === c1) return;
    E.edit('Merge Cells', () => {
      D.touchTbl(t);
      const mp = R.tblMap(t);
      const top = mp[r0].find((m) => m.c0 === c0);
      if (!top) return E.sel;
      const blocks = [];
      for (let r = r0; r <= r1; r++) for (const m of mp[r]) {
        if (m.c0 < c0 || m.c0 > c1 || m.hidden && m.cell !== top.cell) { if (!(m.c0 >= c0 && m.c0 <= c1)) continue; }
        if (m.c0 < c0 || m.c0 > c1) continue;
        const content = m.cell.blocks.filter((b) => b.t !== 'p' || D.plen(b));
        if (content.length) blocks.push(...content);
      }
      for (let r = r0; r <= r1; r++) {
        const row = t.rows[r];
        const inRange = mp[r].filter((m) => m.c0 >= c0 && m.c0 <= c1);
        const keep = inRange[0];
        if (!keep) continue;
        for (const m of inRange.slice(1)) row.cells.splice(row.cells.indexOf(m.cell), 1);
        keep.cell.tcPr.span = c1 - c0 + 1;
        if (keep.cell.tcPr.span === 1) delete keep.cell.tcPr.span;
        if (r1 > r0) keep.cell.tcPr.vMerge = r === r0 ? 'restart' : 'continue';
        else delete keep.cell.tcPr.vMerge;
        if (r > r0) keep.cell.blocks = [emptyLike(keep.cell)];
      }
      top.cell.blocks = blocks.length ? blocks : [emptyLike(top.cell)];
      if (top.cell.blocks[top.cell.blocks.length - 1].t !== 'p') top.cell.blocks.push(emptyLike(top.cell));
      fixWidths(t);
      E.cellSel = null;
      return firstPos(top.cell);
    });
  };
  /** split the current cell (or each selected cell) into rows × cols */
  T.split = function (nRows, nCols, mergeFirst) {
    nRows = Math.max(1, nRows | 0); nCols = Math.max(1, nCols | 0);
    if (mergeFirst && E.cellSel) T.merge();
    const ci = T.cur();
    if (!ci) return;
    const t = ci.tbl;
    const targets = E.cellSel && !mergeFirst ? E.cellSel.cells.slice() : [ci.cell];
    E.edit('Split Cells', () => {
      D.touchTbl(t);
      for (const cell of targets) {
        if (nCols > 1) splitCols(t, cell, nCols);
        if (nRows > 1) splitRows(t, cell, nRows);
      }
      fixWidths(t);
      E.cellSel = null;
      return firstPos(targets[0]);
    });
  };
  function splitCols(t, cell, n) {
    const map = R.tblMap(t);
    let ri = -1, m = null;
    for (let r = 0; r < map.length && !m; r++) { const x = map[r].find((y) => y.cell === cell); if (x) { ri = r; m = x; } }
    if (!m) return;
    const xs = xsOf(t.grid);
    const x0 = xs[m.c0], x1 = xs[m.c0 + m.span];
    const cuts = [];
    for (let k = 1; k < n; k++) cuts.push(L.round(x0 + ((x1 - x0) * k) / n, 2));
    regrid(t, cuts);
    /* now the cell spans several grid columns; replace it by n cells with matching spans */
    const map2 = R.tblMap(t);
    const xs2 = xsOf(t.grid);
    const m2 = map2[ri].find((y) => y.cell === cell);
    const edges = [x0].concat(cuts, [x1]);
    const idx = (x) => xs2.findIndex((v) => Math.abs(v - x) < 0.02);
    const news = [];
    for (let k = 0; k < n; k++) {
      const c = k === 0 ? cell : newCellLike(cell, 1);
      const span = idx(edges[k + 1]) - idx(edges[k]);
      if (span > 1) c.tcPr.span = span; else delete c.tcPr.span;
      if (k) delete c.tcPr.vMerge;
      news.push(c);
    }
    /* like Word, a cell holding several paragraphs hands them out one per new cell (the rest stay in the last) */
    const paras = cell.blocks;
    if (paras.length > 1 && paras.every((b) => b.t === 'p')) {
      const nonEmpty = paras.filter((b) => D.plen(b) > 0);
      if (nonEmpty.length > 1) {
        D.touchList(cell);
        const k = Math.min(n, nonEmpty.length);
        for (let j = 0; j < k; j++) news[j].blocks = j < k - 1 ? [nonEmpty[j]] : nonEmpty.slice(j);
      }
    }
    const row = t.rows[ri];
    row.cells.splice(row.cells.indexOf(cell), 1, ...news);
    /* rows below that continue a vertical merge of this cell get matching cells */
    for (let r = ri + 1; r < map2.length; r++) {
      const below = map2[r].find((y) => y.c0 === m2.c0 && y.cell.tcPr.vMerge === 'continue');
      if (!below) break;
      const bs = news.map((nc, k) => { const c = k === 0 ? below.cell : newCellLike(below.cell, 1); if (nc.tcPr.span) c.tcPr.span = nc.tcPr.span; else delete c.tcPr.span; c.tcPr.vMerge = 'continue'; return c; });
      t.rows[r].cells.splice(t.rows[r].cells.indexOf(below.cell), 1, ...bs);
      for (const nc of news.slice(1)) nc.tcPr.vMerge = 'restart';
    }
  }
  function splitRows(t, cell, n) {
    const map = R.tblMap(t);
    let ri = -1, m = null;
    for (let r = 0; r < map.length && !m; r++) { const x = map[r].find((y) => y.cell === cell); if (x) { ri = r; m = x; } }
    if (!m) return;
    if (m.rowspan >= n && m.rowspan > 1) {
      /* a vertically merged cell is divided among the rows it already spans */
      const per = Math.floor(m.rowspan / n);
      let r = ri;
      for (let k = 0; k < n; k++) {
        for (let j = 0; j < per; j++) {
          const x = R.tblMap(t)[r + j].find((y) => y.c0 === m.c0);
          if (!x) continue;
          if (j === 0) { x.cell.tcPr.vMerge = per > 1 ? 'restart' : undefined; if (!x.cell.tcPr.vMerge) delete x.cell.tcPr.vMerge; if (k) x.cell.blocks = [emptyLike(cell)]; }
          else x.cell.tcPr.vMerge = 'continue';
        }
        r += per;
      }
      return;
    }
    const row = t.rows[ri];
    const h = row.trPr.h;
    const newRows = [];
    for (let k = 1; k < n; k++) {
      const cells = map[ri].map((x) => {
        if (x.cell === cell) return newCellLike(cell, x.span);
        const c = newCellLike(x.cell, x.span);
        c.tcPr.vMerge = 'continue';
        return c;
      });
      const trPr = L.clone(row.trPr);
      delete trPr.header;
      if (h) trPr.h = L.round(h / n, 2);
      newRows.push(D.row(cells, trPr));
    }
    for (const x of map[ri]) if (x.cell !== cell) { if (!x.cell.tcPr.vMerge) x.cell.tcPr.vMerge = 'restart'; }
    /* cells that were merged down from this row keep continuing below the new rows */
    if (h) row.trPr.h = L.round(h / n, 2);
    t.rows.splice(ri + 1, 0, ...newRows);
  }
  /** add grid boundaries (pt from the table's left edge) and recompute every cell's span */
  function regrid(t, extra) {
    const map = R.tblMap(t);
    const xs = xsOf(t.grid);
    const cellEdges = map.map((row, ri) => { const before = t.rows[ri].trPr.gridBefore || 0; return { b: xs[before], cells: row.map((m) => [xs[m.c0], xs[Math.min(xs.length - 1, m.c0 + m.span)]]) }; });
    const all = xs.concat(extra || []).map((v) => L.round(v, 2)).sort((a, b) => a - b);
    const bounds = [];
    for (const v of all) if (!bounds.length || v - bounds[bounds.length - 1] > 0.05) bounds.push(v);
    t.grid = [];
    for (let i = 1; i < bounds.length; i++) t.grid.push(L.round(bounds[i] - bounds[i - 1], 2));
    const idx = (x) => { let best = 0, bd = Infinity; bounds.forEach((v, i) => { const dd = Math.abs(v - x); if (dd < bd) { bd = dd; best = i; } }); return best; };
    map.forEach((row, ri) => {
      const before = idx(cellEdges[ri].b);
      if (before) t.rows[ri].trPr.gridBefore = before; else delete t.rows[ri].trPr.gridBefore;
      row.forEach((m, k) => { const [a, b] = cellEdges[ri].cells[k]; const span = Math.max(1, idx(b) - idx(a)); if (span > 1) m.cell.tcPr.span = span; else delete m.cell.tcPr.span; });
    });
  }
  T.regrid = regrid;
  T.splitTable = function () {
    const ci = T.cur();
    if (!ci) return;
    const t = ci.tbl;
    E.edit('Split Table', () => {
      const tb = D.tblInfo(doc(), t);
      const cont = D.touchList(tb.cont);
      const i = cont.blocks.indexOf(t);
      const np = D.para();
      if (ci.ri === 0) { cont.blocks.splice(i, 0, np); return D.pos(np, 0); }
      D.touchTbl(t);
      const rest = t.rows.splice(ci.ri);
      const t2 = D.table(rest, t.grid.slice(), L.clone(t.tblPr));
      cleanMerges(t); cleanMerges(t2);
      cont.blocks.splice(i + 1, 0, np, t2);
      doc()._idxDirty = true;
      return D.pos(np, 0);
    });
  };

  /* ================= properties ================= */
  function forCells(label, fn) {
    const rg = T.range();
    if (!rg) return;
    E.edit(label, () => { D.touchTbl(rg.tbl); for (const c of rg.cells) fn(c, rg); return E.sel; });
  }
  T.forCells = forCells;
  T.setVAlign = (v) => forCells('Cell Alignment', (c) => { if (v === 'top') delete c.tcPr.vAlign; else c.tcPr.vAlign = v; });
  T.cycleTextDirection = () => forCells('Text Direction', (c) => { const v = c.tcPr.textDir; c.tcPr.textDir = !v ? 'tbRl' : v === 'tbRl' ? 'btLr' : undefined; if (!c.tcPr.textDir) delete c.tcPr.textDir; });
  T.toggleHeaderRow = function () {
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl;
    const on = !t.rows[rg.r0].trPr.header;
    E.edit('Heading Rows Repeat', () => { D.touchTbl(t); for (let r = 0; r <= rg.r1; r++) { if (on) t.rows[r].trPr.header = true; else delete t.rows[r].trPr.header; } return E.sel; });
  };
  T.shade = function (fill) {
    if (!T.inTable()) { E.formatPara((pPr) => { pPr.shd = fill ? { val: 'clear', fill, color: 'auto' } : undefined; }, 'Shading'); return; }
    forCells('Shading', (c) => { if (fill) c.tcPr.shd = { val: 'clear', fill, color: 'auto' }; else delete c.tcPr.shd; });
  };
  /** border buttons applied to the selected cells */
  T.applyBorder = function (k) {
    const rg = T.range();
    if (!rg) return;
    const pen = Object.assign({}, A().borderPen);
    const b = pen.val === 'nil' ? { val: 'nil' } : { val: pen.val, sz: pen.sz, color: pen.color, space: 0 };
    const nil = { val: 'nil' };
    const t = rg.tbl;
    const map = rg.map;
    const pos = new Map();
    for (let r = rg.r0; r <= rg.r1; r++) for (const m of map[r]) if (rg.cells.includes(m.cell)) pos.set(m.cell, { r, c0: m.c0, c1: m.c0 + m.span - 1, r1: r + (m.rowspan || 1) - 1 });
    const edgesFor = (p) => {
      const outer = { top: p.r === rg.r0, bottom: p.r1 >= rg.r1, left: p.c0 === rg.c0, right: p.c1 >= rg.c1 };
      switch (k) {
        case 'all': return { top: 1, bottom: 1, left: 1, right: 1 };
        case 'outside': return { top: outer.top, bottom: outer.bottom, left: outer.left, right: outer.right };
        case 'inside': return { top: !outer.top, bottom: !outer.bottom, left: !outer.left, right: !outer.right };
        case 'insideH': return { top: !outer.top, bottom: !outer.bottom };
        case 'insideV': return { left: !outer.left, right: !outer.right };
        case 'top': return { top: outer.top };
        case 'bottom': return { bottom: outer.bottom };
        case 'left': return { left: outer.left };
        case 'right': return { right: outer.right };
        case 'diagDown': return { tl2br: 1 };
        case 'diagUp': return { tr2bl: 1 };
        case 'none': return { top: 1, bottom: 1, left: 1, right: 1, tl2br: 1, tr2bl: 1 };
        default: return {};
      }
    };
    const { tp } = R.tblProps(doc(), t);
    const effective = (c, side, p) => { const own = c.tcPr.borders && c.tcPr.borders[side]; if (own) return own; const tb = tp.borders || {}; if (side === 'top') return p.r === 0 ? tb.top : tb.insideH; if (side === 'bottom') return p.r1 >= t.rows.length - 1 ? tb.bottom : tb.insideH; if (side === 'left') return p.c0 === 0 ? tb.left : tb.insideV; if (side === 'right') return p.c1 >= t.grid.length - 1 ? tb.right : tb.insideV; return null; };
    /* toggle: if every target edge already shows a border, the button removes them */
    let allOn = k !== 'none';
    for (const [c, p] of pos) { const e = edgesFor(p); for (const s in e) if (e[s] && !R.hasBorder(effective(c, s, p))) allOn = false; }
    E.edit('Borders', () => {
      D.touchTbl(t);
      for (const [c, p] of pos) {
        const e = edgesFor(p);
        const bd = Object.assign({}, c.tcPr.borders || {});
        for (const s in e) if (e[s]) bd[s] = k === 'none' || allOn ? (s === 'tl2br' || s === 'tr2bl' ? undefined : nil) : Object.assign({}, b);
        for (const s in bd) if (!bd[s]) delete bd[s];
        c.tcPr.borders = Object.keys(bd).length ? bd : undefined;
        if (!c.tcPr.borders) delete c.tcPr.borders;
      }
      /* neighbours share edges: clear the facing side so the new border shows */
      if (k !== 'none') for (const [c, p] of pos) {
        const e = edgesFor(p);
        if (e.right && p.c1 + 1 < t.grid.length) { const nb = map[p.r].find((m) => m.c0 === p.c1 + 1); if (nb && !pos.has(nb.cell) && nb.cell.tcPr.borders && nb.cell.tcPr.borders.left) delete nb.cell.tcPr.borders.left; }
        if (e.bottom && map[p.r1 + 1]) { const nb = map[p.r1 + 1].find((m) => m.c0 === p.c0); if (nb && !pos.has(nb.cell) && nb.cell.tcPr.borders && nb.cell.tcPr.borders.top) delete nb.cell.tcPr.borders.top; }
        void c;
      }
      return E.sel;
    });
  };

  /* ================= AutoFit & distribute ================= */
  let measureCtx = null;
  function textWidth(text, r) {
    if (!measureCtx) measureCtx = document.createElement('canvas').getContext('2d');
    measureCtx.font = `${r.i ? 'italic ' : ''}${r.b ? 'bold ' : ''}${(r.sz || 12)}px ${L.fontStack(r.font)}`;
    return measureCtx.measureText(text).width; /* in pt because font size is given in "px" = pt units */
  }
  function cellContentWidths(t, cell) {
    const d = doc();
    let min = 0, max = 0;
    D.walk({ blocks: cell.blocks }, (b) => {
      if (b.t !== 'p') return;
      const pp = D.pProps(d, b);
      const ind = (pp.ind ? (pp.ind.l || 0) + (pp.ind.r || 0) + Math.max(0, pp.ind.fl || 0) : 0);
      let line = 0, lineMax = 0;
      for (const it of b.runs) {
        if (it.t === 'text') {
          const r = D.rProps(d, b, it);
          const w = textWidth(it.text, r);
          line += w;
          for (const word of it.text.split(/\s+/)) if (word) min = Math.max(min, textWidth(word, r) + ind);
        } else if (it.t === 'tab') line += 36;
        else if (it.t === 'br') { lineMax = Math.max(lineMax, line); line = 0; }
        else if ((it.t === 'img' || it.t === 'shape' || it.t === 'chart' || it.t === 'group') && !it.float) { line += it.w || 0; min = Math.max(min, (it.w || 0) + ind); }
      }
      lineMax = Math.max(lineMax, line);
      const lb = LY.labels && LY.labels.get(b.id);
      if (lb) lineMax += 18;
      max = Math.max(max, lineMax + ind);
    });
    return { min, max };
  }
  T.autofit = function (mode) {
    const ci = T.cur();
    if (!ci) return;
    const t = ci.tbl;
    const d = doc();
    const avail = availWidth(D.pos(D.firstPara(t.rows[0].cells[0].blocks) || E.sel.f.p, 0)) ;
    const outer = (() => { const tb = D.tblInfo(d, t); if (!tb || !tb.parents.length) { const s = D.sectFor(d, E.sel.f.p); return s.pgW - s.ml - s.mr - (s.gutter || 0); } return avail; })();
    E.edit('AutoFit', () => {
      D.touchTbl(t);
      const { tp } = R.tblProps(d, t);
      const mar = (tp.cellMar.l || 0) + (tp.cellMar.r || 0);
      if (mode === 'window') {
        const total = t.grid.reduce((a, b) => a + b, 0) || 1;
        t.grid = t.grid.map((g) => L.round((g * outer) / total, 2));
        t.tblPr.w = { type: 'pct', v: 5000 };
        t.tblPr.layout = 'autofit';
      } else if (mode === 'fixed') {
        t.tblPr.layout = 'fixed';
        t.tblPr.w = { type: 'dxa', v: L.round(t.grid.reduce((a, b) => a + b, 0), 2) };
      } else {
        const n = t.grid.length;
        const mins = new Array(n).fill(mar + 6), maxs = new Array(n).fill(mar + 6);
        const map = R.tblMap(t);
        /* single-span cells first, then spread spanning cells */
        for (const row of map) for (const m of row) {
          if (m.hidden) continue;
          const w = cellContentWidths(t, m.cell);
          if (m.span === 1) { mins[m.c0] = Math.max(mins[m.c0], w.min + mar); maxs[m.c0] = Math.max(maxs[m.c0], w.max + mar); }
        }
        for (const row of map) for (const m of row) {
          if (m.hidden || m.span === 1) continue;
          const w = cellContentWidths(t, m.cell);
          const curMax = maxs.slice(m.c0, m.c0 + m.span).reduce((a, b) => a + b, 0);
          if (w.max + mar > curMax) { const extra = (w.max + mar - curMax) / m.span; for (let k = m.c0; k < m.c0 + m.span; k++) maxs[k] += extra; }
        }
        const sumMax = maxs.reduce((a, b) => a + b, 0), sumMin = mins.reduce((a, b) => a + b, 0);
        let g;
        if (sumMax <= outer) g = maxs;
        else if (sumMin >= outer) g = mins.map((v) => (v * outer) / sumMin);
        else { const room = outer - sumMin, want = sumMax - sumMin; g = mins.map((v, i) => v + ((maxs[i] - v) * room) / want); }
        t.grid = g.map((v) => L.round(Math.max(v, 12), 2));
        t.tblPr.w = { type: 'auto', v: 0 };
        t.tblPr.layout = 'autofit';
      }
      fixWidths(t);
      return E.sel;
    });
  };
  T.distribute = function (what) {
    const rg = T.range();
    if (!rg) return;
    const t = rg.tbl;
    const whole = !E.cellSel;
    E.edit(what === 'rows' ? 'Distribute Rows Evenly' : 'Distribute Columns Evenly', () => {
      D.touchTbl(t);
      if (what === 'cols') {
        const c0 = whole ? 0 : rg.c0, c1 = whole ? t.grid.length - 1 : rg.c1;
        const sum = t.grid.slice(c0, c1 + 1).reduce((a, b) => a + b, 0);
        for (let i = c0; i <= c1; i++) t.grid[i] = L.round(sum / (c1 - c0 + 1), 2);
        fixWidths(t);
      } else {
        const r0 = whole ? 0 : rg.r0, r1 = whole ? t.rows.length - 1 : rg.r1;
        let total = 0, n = 0;
        for (let r = r0; r <= r1; r++) {
          const tr = LY.root.querySelector(`table[data-tid="${t.id}"] tr[data-ri="${r}"]`);
          if (tr) { total += tr.getBoundingClientRect().height / LY.zoom / D.PX; n++; }
        }
        const hh = n ? L.round(total / n, 2) : 18;
        for (let r = r0; r <= r1; r++) { t.rows[r].trPr.h = hh; t.rows[r].trPr.hRule = 'atLeast'; }
      }
      return E.sel;
    });
  };
  /** set row height / column width (Table Properties) */
  T.setColWidth = function (c0, c1, w) {
    const ci = T.cur();
    if (!ci) return;
    const t = ci.tbl;
    for (let i = c0; i <= c1; i++) t.grid[i] = Math.max(6, w);
    fixWidths(t);
  };

  /* ================= sorting ================= */
  const parseNum = (s) => { const m = /-?\(?[$€£¥]?\s*-?[\d,]*\.?\d+%?\)?/.exec(String(s).replace(/\s+/g, ' ')); if (!m) return NaN; let t = m[0]; const neg = /^\(.*\)$/.test(t) || /^-/.test(t); t = t.replace(/[()$€£¥,%\s-]/g, ''); const v = parseFloat(t); return isNaN(v) ? NaN : neg ? -v : v; };
  T.parseNum = parseNum;
  const cmpBy = (type, desc, caseSens) => (a, b) => {
    let r;
    if (type === 'number') { const x = parseNum(a), y = parseNum(b); r = isNaN(x) && isNaN(y) ? 0 : isNaN(x) ? 1 : isNaN(y) ? -1 : x - y; }
    else if (type === 'date') { const x = Date.parse(a), y = Date.parse(b); r = isNaN(x) && isNaN(y) ? 0 : isNaN(x) ? 1 : isNaN(y) ? -1 : x - y; }
    else r = caseSens ? (a < b ? -1 : a > b ? 1 : 0) : a.localeCompare(b, undefined, { sensitivity: 'base', numeric: true });
    return desc ? -r : r;
  };
  const cellText = (cell) => cell.blocks.filter((b) => b.t === 'p').map((p) => D.plainText(p, { skipDeleted: true })).join(' ').trim();
  T.cellText = cellText;
  /** keys: [{col (grid index or field index), type, desc}] */
  T.sortTable = function (keys, opts) {
    opts = opts || {};
    const ci = T.cur();
    if (!ci) return;
    const t = ci.tbl;
    const map = R.tblMap(t);
    if (map.some((row) => row.some((m) => m.rowspan > 1 || m.hidden))) { L.ui.msg('Tables with vertically merged cells cannot be sorted.', { icon: 'warn' }); return; }
    let hdr = 0;
    if (opts.header) hdr = 1;
    else while (hdr < t.rows.length && t.rows[hdr].trPr.header) hdr++;
    E.edit('Sort', () => {
      D.touchTbl(t);
      const body = t.rows.slice(hdr);
      const textAt = (row, col) => { const m = R.tblMap({ rows: [row], tblPr: t.tblPr })[0].find((x) => x.c0 <= col && col < x.c0 + x.span); return m ? cellText(m.cell) : ''; };
      body.sort((a, b) => { for (const k of keys) { const r = cmpBy(k.type, k.desc, opts.caseSensitive)(textAt(a, k.col), textAt(b, k.col)); if (r) return r; } return 0; });
      t.rows = t.rows.slice(0, hdr).concat(body);
      return E.sel;
    });
  };
  T.sortParagraphs = function (keys, opts) {
    opts = opts || {};
    const paras = E.selectedParas();
    if (paras.length < 2) return;
    const d = doc();
    const info = D.info(d, paras[0]);
    if (paras.some((p) => D.info(d, p).cont !== info.cont)) return;
    E.edit('Sort Text', () => {
      const cont = D.touchList(info.cont);
      const idx = paras.map((p) => cont.blocks.indexOf(p));
      const i0 = Math.min(...idx);
      const sep = opts.sep || '\t';
      const field = (p, k) => { const t = D.plainText(p).trim(); if (k.field == null || k.field === 0) return t; const parts = t.split(sep); return (parts[k.field - 1] || '').trim(); };
      let list = paras.slice();
      const head = opts.header ? list.shift() : null;
      list = list.filter((p) => true || p);
      list.sort((a, b) => { for (const k of keys) { const r = cmpBy(k.type, k.desc, opts.caseSensitive)(field(a, k), field(b, k)); if (r) return r; } return 0; });
      /* empty paragraphs sort to the top in Word */
      const empties = list.filter((p) => !D.plainText(p).trim()), full = list.filter((p) => D.plainText(p).trim());
      const ordered = (head ? [head] : []).concat(empties, full);
      cont.blocks.splice(i0, paras.length, ...ordered);
      const last = ordered[ordered.length - 1];
      return { a: D.pos(ordered[0], 0), f: D.pos(last, D.plen(last)) };
    });
  };
  /** Word guesses whether the first row is a heading: flagged as a header, set in a different type
   *  (text above numbers), or formatted differently (bold over plain) */
  function looksLikeHeader(t) {
    if (t.rows.length < 2) return false;
    if (t.rows[0].trPr.header) return true;
    const rowCells = (r) => R.tblMap({ rows: [r], tblPr: {} })[0].map((m) => m.cell);
    const first = rowCells(t.rows[0]), rest = t.rows.slice(1).map(rowCells);
    const isNum = (s) => s !== '' && !isNaN(parseNum(s));
    for (let c = 0; c < first.length; c++) {
      const head = cellText(first[c]);
      const col = rest.map((r) => (r[c] ? cellText(r[c]) : '')).filter(Boolean);
      if (head && !isNum(head) && col.length && col.every(isNum)) return true;
    }
    const bold = (cell) => { const p = cell.blocks.find((b) => b.t === 'p'); if (!p) return false; const runs = p.runs.filter((r) => r.t === 'text' && r.text.trim()); return runs.length > 0 && runs.every((r) => !!D.rProps(doc(), p, r).b); };
    if (first.every(bold) && rest.length && !rest[0].every(bold)) return true;
    return false;
  }
  T.quickSort = function (desc) {
    if (T.inTable()) {
      const rg = T.range();
      const col = rg.c0;
      const sample = rg.tbl.rows.slice(1).map((r) => { const m = R.tblMap({ rows: [r], tblPr: {} })[0].find((x) => x.c0 <= col && col < x.c0 + x.span); return m ? cellText(m.cell) : ''; });
      const numeric = sample.filter(Boolean).length && sample.filter(Boolean).every((s) => !isNaN(parseNum(s)));
      T.sortTable([{ col, type: numeric ? 'number' : 'text', desc }], { header: looksLikeHeader(rg.tbl) });
    } else T.sortParagraphs([{ field: 0, type: 'text', desc }]);
  };

  /* ================= formulas ================= */
  const colName = (i) => { let s = ''; i++; while (i > 0) { const k = (i - 1) % 26; s = String.fromCharCode(65 + k) + s; i = Math.floor((i - 1) / 26); } return s; };
  T.colName = colName;
  function refToRC(ref) { const m = /^([A-Z]+)(\d+)$/i.exec(ref); if (!m) return null; let c = 0; for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64); return { r: +m[2] - 1, c: c - 1 }; }
  /** evaluate a formula like "SUM(ABOVE)" for the cell containing paragraph p; bookmarks resolve to numbers too */
  T.evalFormula = function (expr, p) {
    const d = doc();
    const ci = p ? D.cellOf(d, p) : null;
    const t = ci ? ci.tbl : null;
    const map = t ? R.tblMap(t) : null;
    const myC = ci ? map[ci.ri].find((m) => m.cell === ci.cell).c0 : 0;
    const valAt = (r, c) => { if (!t || !map[r]) return NaN; const m = map[r].find((x) => x.c0 <= c && c < x.c0 + x.span); return m && !m.hidden ? parseNum(cellText(m.cell)) : NaN; };
    const rangeVals = (arg) => {
      arg = arg.trim().toUpperCase();
      const out = [];
      if (!t) return out;
      if (arg === 'ABOVE') { for (let r = ci.ri - 1; r >= 0; r--) { const v = valAt(r, myC); if (isNaN(v)) { if (out.length) break; continue; } out.push(v); } return out; }
      if (arg === 'BELOW') { for (let r = ci.ri + 1; r < map.length; r++) { const v = valAt(r, myC); if (isNaN(v)) { if (out.length) break; continue; } out.push(v); } return out; }
      if (arg === 'LEFT') { for (let c = myC - 1; c >= 0; c--) { const v = valAt(ci.ri, c); if (isNaN(v)) { if (out.length) break; continue; } out.push(v); } return out; }
      if (arg === 'RIGHT') { for (let c = myC + 1; c < t.grid.length; c++) { const v = valAt(ci.ri, c); if (isNaN(v)) { if (out.length) break; continue; } out.push(v); } return out; }
      const rm = /^([A-Z]+\d+):([A-Z]+\d+)$/.exec(arg);
      if (rm) { const a = refToRC(rm[1]), b = refToRC(rm[2]); for (let r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r++) for (let c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c++) { const v = valAt(r, c); if (!isNaN(v)) out.push(v); } return out; }
      const one = refToRC(arg);
      if (one) { const v = valAt(one.r, one.c); if (!isNaN(v)) out.push(v); return out; }
      const n = parseFloat(arg);
      if (!isNaN(n)) out.push(n);
      return out;
    };
    const FN = {
      SUM: (v) => v.reduce((a, b) => a + b, 0), AVERAGE: (v) => (v.length ? v.reduce((a, b) => a + b, 0) / v.length : 0), PRODUCT: (v) => v.reduce((a, b) => a * b, 1),
      MIN: (v) => Math.min(...v), MAX: (v) => Math.max(...v), COUNT: (v) => v.length, ABS: (v) => Math.abs(v[0]), INT: (v) => Math.trunc(v[0]), ROUND: (v) => { const k = Math.pow(10, v[1] || 0); return Math.round(v[0] * k) / k; },
      MOD: (v) => v[0] % v[1], SIGN: (v) => Math.sign(v[0]), AND: (v) => (v.every(Boolean) ? 1 : 0), OR: (v) => (v.some(Boolean) ? 1 : 0), NOT: (v) => (v[0] ? 0 : 1), TRUE: () => 1, FALSE: () => 0, IF: (v) => (v[0] ? v[1] : v[2]), DEFINED: (v) => (v.length ? 1 : 0),
    };
    /* bookmark values */
    const bm = (name) => { const r = L.fields && L.fields.bookmarkText ? L.fields.bookmarkText(name) : null; return r == null ? NaN : parseNum(r); };
    let src = String(expr).replace(/^\s*=\s*/, '').replace(/\\#.*$/, '').replace(/\\\*.*$/, '').trim();
    /* tokenizer + recursive descent */
    const toks = [];
    const re = /\s*(<=|>=|<>|[-+*/^%(),:<>=]|\d*\.?\d+(?:[eE][-+]?\d+)?|[A-Za-z_][A-Za-z0-9_]*(?::[A-Za-z]+\d+)?)/y;
    let m;
    while (re.lastIndex < src.length && (m = re.exec(src))) toks.push(m[1]);
    let i = 0;
    const peek = () => toks[i], next = () => toks[i++];
    function args() {
      const out = [];
      if (peek() === ')') { next(); return out; }
      for (;;) {
        /* range-like arguments */
        const tk = peek();
        if (tk && /^(ABOVE|BELOW|LEFT|RIGHT)$/i.test(tk) || tk && /^[A-Z]+\d+:[A-Z]+\d+$/i.test(tk)) { next(); out.push(...rangeVals(tk)); }
        else if (tk && /^[A-Z]+\d+$/i.test(tk) && toks[i + 1] === ':' && /^[A-Z]+\d+$/i.test(toks[i + 2] || '')) { next(); next(); const b = next(); out.push(...rangeVals(tk + ':' + b)); }
        else out.push(cmp());
        const s = next();
        if (s === ')') break;
        if (s !== ',') throw new Error('Syntax Error');
      }
      return out;
    }
    function prim() {
      const tk = next();
      if (tk == null) throw new Error('Syntax Error');
      if (tk === '(') { const v = cmp(); if (next() !== ')') throw new Error('Syntax Error'); return v; }
      if (tk === '-') return -prim();
      if (tk === '+') return prim();
      if (/^\d*\.?\d/.test(tk)) return parseFloat(tk);
      const up = tk.toUpperCase();
      if (peek() === '(') { next(); const f = FN[up]; if (!f) throw new Error('Syntax Error, ' + tk); return f(args()); }
      if (/^[A-Z]+\d+$/.test(up)) { const v = rangeVals(up); return v.length ? v[0] : 0; }
      if (FN[up]) return FN[up]([]);
      const b = bm(tk);
      if (!isNaN(b)) return b;
      throw new Error('Undefined Bookmark, ' + tk);
    }
    function pow() { let v = prim(); while (peek() === '^') { next(); v = Math.pow(v, prim()); } if (peek() === '%') { next(); v = v / 100; } return v; }
    function mul() { let v = pow(); while (peek() === '*' || peek() === '/') { const o = next(); const w = pow(); v = o === '*' ? v * w : v / w; } return v; }
    function add() { let v = mul(); while (peek() === '+' || peek() === '-') { const o = next(); const w = mul(); v = o === '+' ? v + w : v - w; } return v; }
    function cmp() { let v = add(); while (['=', '<', '>', '<=', '>=', '<>'].includes(peek())) { const o = next(); const w = add(); v = o === '=' ? +(v === w) : o === '<' ? +(v < w) : o === '>' ? +(v > w) : o === '<=' ? +(v <= w) : o === '>=' ? +(v >= w) : +(v !== w); } return v; }
    const v = cmp();
    if (i < toks.length) throw new Error('Syntax Error');
    return v;
  };
  /** Word numeric picture: "#,##0.00", "$#,##0.00;($#,##0.00)", "0%", "0.00" */
  T.formatNumber = function (v, pic) {
    if (!pic) return String(L.round(v, 10)).replace(/\.0+$/, '');
    pic = pic.replace(/^"|"$/g, '');
    const parts = pic.split(';');
    let p = parts[0];
    if (v < 0 && parts[1]) { p = parts[1]; v = -v; } else if (v === 0 && parts[2]) p = parts[2];
    const pct = p.includes('%');
    if (pct) v *= 100;
    const m = /[#0][#0,]*(\.[#0]+)?/.exec(p);
    if (!m) return p;
    const dec = m[1] ? m[1].length - 1 : 0;
    const minDec = m[1] ? (m[1].match(/0/g) || []).length : 0;
    const grouping = m[0].includes(',');
    const neg = v < 0;
    let s = Math.abs(v).toFixed(dec);
    if (dec > minDec) s = s.replace(new RegExp(`(\\.\\d{${minDec}}\\d*?)0+$`), '$1').replace(/\.$/, '');
    let [ip, fp] = s.split('.');
    const minInt = (m[0].split('.')[0].match(/0/g) || []).length;
    while (ip.length < minInt) ip = '0' + ip;
    if (grouping) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
    const num = ip + (fp ? '.' + fp : '');
    return p.slice(0, m.index) + (neg && !/[(-]/.test(p) ? '-' : '') + num + p.slice(m.index + m[0].length).replace('%', pct ? '%' : '');
  };
  T.autoSum = function () {
    const ci = T.cur();
    if (!ci) return;
    let dir = 'ABOVE';
    try { const v = T.evalFormula('SUM(ABOVE)', E.sel.f.p); const left = T.evalFormula('SUM(LEFT)', E.sel.f.p); if (!v && left) dir = 'LEFT'; } catch (e) { /* default above */ }
    A().insertFieldNow('=SUM(' + dir + ')', 'AutoSum');
  };

  /* ================= conversion ================= */
  T.textToTable = function (opts) {
    opts = opts || {};
    const paras = E.selectedParas();
    if (!paras.length) return;
    const d = doc();
    const info = D.info(d, paras[0]);
    if (paras.some((p) => D.info(d, p).cont !== info.cont)) return;
    const sep = opts.sep == null ? '\t' : opts.sep;
    const rowsText = [];
    if (sep === 'para') {
      const n = Math.max(1, opts.cols || 1);
      for (let i = 0; i < paras.length; i += n) rowsText.push(paras.slice(i, i + n).map((p) => ({ src: p, runs: p.runs })));
    } else {
      for (const p of paras) {
        /* split runs at the separator character, keeping formatting */
        const cells = [[]];
        for (const it of p.runs) {
          if ((sep === '\t' && it.t === 'tab')) { cells.push([]); continue; }
          if (it.t === 'text' && sep !== '\t' && it.text.includes(sep)) {
            const parts = it.text.split(sep);
            parts.forEach((t, k) => { if (k) cells.push([]); if (t) cells[cells.length - 1].push(Object.assign({}, it, { text: t, rPr: L.clone(it.rPr) })); });
            continue;
          }
          cells[cells.length - 1].push(L.clone(it));
        }
        rowsText.push(cells.map((runs) => ({ src: p, runs })));
      }
    }
    const nc = Math.max(opts.cols || 0, ...rowsText.map((r) => r.length));
    const width = availWidth(D.pos(paras[0], 0));
    E.edit('Convert Text to Table', () => {
      const t = D.simpleTable(rowsText.length, nc, opts.colWidth ? opts.colWidth * nc : width);
      rowsText.forEach((r, ri) => r.forEach((c, cix) => { const p = t.rows[ri].cells[cix].blocks[0]; p.runs = c.runs.map((x) => L.clone(x)); p.pPr = L.clone(c.src.pPr); delete p.pPr.num; D.normalize(p); }));
      const cont = D.touchList(info.cont);
      const i0 = cont.blocks.indexOf(paras[0]);
      cont.blocks.splice(i0, paras.length, t);
      if (!cont.blocks[i0 + 1]) cont.blocks.push(D.para());
      d._idxDirty = true;
      return firstPos(t.rows[0].cells[0]);
    });
    if (opts.autofit === 'contents') T.autofit('contents');
    else if (opts.autofit === 'window') T.autofit('window');
  };
  T.tableToText = function (sep) {
    const ci = T.cur();
    if (!ci) return;
    sep = sep == null ? '\t' : sep;
    const t = ci.tbl;
    E.edit('Convert Table to Text', () => {
      const out = [];
      const flat = (blocks) => { const res = []; for (const b of blocks) { if (b.t === 'p') res.push(b); else if (b.t === 'tbl') for (const r of b.rows) for (const c of r.cells) res.push(...flat(c.blocks)); } return res; };
      for (const row of t.rows) {
        const cells = row.cells.filter((c) => c.tcPr.vMerge !== 'continue');
        if (sep === 'para') { for (const c of cells) out.push(...flat(c.blocks).map((p) => D.para(L.clone(p.runs), L.clone(p.pPr), L.clone(p.rPr)))); continue; }
        const p = D.para([], L.clone((D.firstPara(cells[0].blocks) || D.para()).pPr));
        cells.forEach((c, k) => {
          if (k) p.runs.push(sep === '\t' ? D.item('tab') : D.text(sep));
          flat(c.blocks).forEach((q, j) => { if (j) p.runs.push(D.text(' ')); p.runs.push(...L.clone(q.runs)); });
        });
        D.normalize(p);
        out.push(p);
      }
      const tb = D.tblInfo(doc(), t);
      const cont = D.touchList(tb.cont);
      const i = cont.blocks.indexOf(t);
      cont.blocks.splice(i, 1, ...out);
      doc()._idxDirty = true;
      return D.pos(out[0], 0);
    });
  };

  /* ================= Table AutoFormat styles (Word 2003 set) ================= */
  const BD = (val, sz, color) => ({ val, sz: sz || 0.5, color: color || 'auto' });
  const all = (b) => ({ top: b, left: b, bottom: b, right: b, insideH: b, insideV: b });
  const SH = (fill, val, color) => ({ val: val || 'clear', fill, color: color || 'auto' });
  T.STYLES = {
    'Table Simple 1': { tblPr: { borders: { top: BD('single', 1.5, '008000'), bottom: BD('single', 1.5, '008000') } }, cond: { firstRow: { tcPr: { borders: { bottom: BD('single', 0.75, '008000') } } }, lastRow: { tcPr: { borders: { top: BD('single', 0.75, '008000') } } } } },
    'Table Simple 2': { cond: { firstRow: { rPr: { b: true }, tcPr: { borders: { bottom: BD('single', 1.5, '000000') } } }, lastRow: { rPr: { b: true }, tcPr: { borders: { top: BD('single', 1.5, '000000') } } }, lastCol: { rPr: { b: true }, tcPr: { borders: { left: BD('single', 1.5, '000000') } } }, band1Vert: { tcPr: { borders: { right: BD('single', 0.5) } } } } },
    'Table Simple 3': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5), left: BD('single', 1.5), right: BD('single', 1.5) } }, cond: { firstRow: { rPr: { b: true, color: 'FFFFFF' }, tcPr: { shd: SH('000000', 'solid', '000000') } } } },
    'Table Classic 1': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5) } }, cond: { firstRow: { rPr: { i: true }, tcPr: { borders: { bottom: BD('single', 0.75) } } }, lastRow: { tcPr: { borders: { top: BD('single', 0.75) } } }, firstCol: { tcPr: { borders: { right: BD('single', 0.75) } } } } },
    'Table Classic 2': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5) } }, cond: { firstRow: { rPr: { b: true, color: 'FFFFFF' }, tcPr: { shd: SH('800080', 'solid', '800080') } }, firstCol: { rPr: { b: true }, tcPr: { borders: { right: BD('single', 0.75) } } }, nwCell: { rPr: { color: 'FFFFFF' } } } },
    'Table Classic 3': { tblPr: { borders: all(BD('single', 1.5, '000080')) }, tcPr: { shd: SH('C0C0C0', 'solid', 'C0C0C0') }, rPr: { color: '000080' }, cond: { firstRow: { rPr: { b: true, i: true, color: 'FFFFFF' }, tcPr: { shd: SH('000080', 'solid', '000080') } }, lastRow: { rPr: { color: '000080' }, tcPr: { shd: SH('FFFFFF', 'solid', 'FFFFFF') } } } },
    'Table Classic 4': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5), left: BD('single', 1.5), right: BD('single', 1.5) } }, cond: { firstRow: { rPr: { b: true, color: 'FFFFFF' }, tcPr: { shd: SH('000080', 'solid', '000080'), borders: { bottom: BD('single', 1.5) } } }, firstCol: { tcPr: { borders: { right: BD('single', 1.5) } } } } },
    'Table Colorful 1': { tblPr: { borders: all(BD('single', 1.5)) }, tcPr: { shd: SH('008080', 'solid', '008080') }, rPr: { b: true, color: 'FFFFFF' }, cond: { firstRow: { rPr: { i: true }, tcPr: { shd: SH('000000', 'solid', '000000'), borders: { bottom: BD('single', 4.5) } } }, firstCol: { rPr: { b: true } } } },
    'Table Colorful 2': { tblPr: { borders: { bottom: BD('single', 1.5) } }, tcPr: { shd: SH('FFFFC0', 'clear') }, cond: { firstRow: { rPr: { b: true, i: true, color: 'FFFFFF' }, tcPr: { shd: SH('800000', 'solid', '800000'), borders: { bottom: BD('single', 1.5) } } }, lastRow: { rPr: { b: true } } } },
    'Table Colorful 3': { tblPr: { borders: { top: BD('single', 1.5, '008080'), left: BD('single', 1.5, '008080'), bottom: BD('single', 1.5, '008080'), right: BD('single', 1.5, '008080') } }, cond: { firstRow: { rPr: { b: true }, tcPr: { shd: SH('C0C0C0', 'clear'), borders: { bottom: BD('single', 1.5, '008080') } } }, firstCol: { tcPr: { shd: SH('C0C0C0', 'clear'), borders: { right: BD('single', 1.5, '008080') } } } } },
    'Table Columns 1': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5) } }, cond: { firstRow: { rPr: { b: true }, tcPr: { borders: { bottom: BD('single', 0.75) } } }, band1Vert: { tcPr: { shd: SH('C0C0C0', 'clear') } }, lastRow: { rPr: { b: true } } } },
    'Table Columns 2': { cond: { firstRow: { rPr: { b: true, color: 'FFFFFF' }, tcPr: { shd: SH('808080', 'solid', '808080') } }, band1Vert: { tcPr: { shd: SH('C0C0C0', 'clear') } }, band2Vert: { tcPr: { shd: SH('E0E0E0', 'clear') } } } },
    'Table Columns 3': { tblPr: { borders: all(BD('single', 0.75, '000080')) }, cond: { firstRow: { rPr: { b: true, color: '000080' }, tcPr: { shd: SH('C0C0C0', 'clear') } }, firstCol: { rPr: { b: true }, tcPr: { shd: SH('C0C0C0', 'clear') } } } },
    'Table Columns 4': { tblPr: { borders: { insideV: BD('single', 0.75) } }, cond: { firstRow: { rPr: { b: true }, tcPr: { borders: { bottom: BD('single', 1.5) } } }, band1Vert: { tcPr: { shd: SH('FFFF99', 'clear') } } } },
    'Table Columns 5': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5), insideV: BD('single', 0.75) } }, cond: { firstRow: { rPr: { b: true, i: true }, tcPr: { borders: { bottom: BD('double', 0.5) } } }, firstCol: { rPr: { b: true } } } },
    'Table Grid 1': { tblPr: { borders: all(BD('single', 0.75)) }, cond: { firstRow: { rPr: { b: true } }, lastCol: { rPr: { b: true } } } },
    'Table Grid 2': { tblPr: { borders: { insideH: BD('single', 0.5), insideV: BD('single', 0.5) } }, cond: { firstRow: { rPr: { b: true }, tcPr: { borders: { bottom: BD('single', 1.5) } } }, lastRow: { tcPr: { borders: { top: BD('single', 1.5) } } } } },
    'Table Grid 3': { tblPr: { borders: all(BD('single', 0.5)) }, cond: { firstRow: { rPr: { b: true }, tcPr: { borders: { bottom: BD('single', 1.5) } } }, firstCol: { tcPr: { borders: { right: BD('single', 1.5) } } } } },
    'Table Grid 4': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5), left: BD('single', 1.5), right: BD('single', 1.5), insideH: BD('single', 0.5), insideV: BD('single', 0.5) } }, cond: { firstRow: { rPr: { b: true } } } },
    'Table Grid 5': { tblPr: { borders: all(BD('single', 1.5)) }, cond: { firstRow: { rPr: { b: true } } } },
    'Table Grid 6': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5), left: BD('single', 1.5), right: BD('single', 1.5), insideH: BD('single', 0.5), insideV: BD('single', 0.5) } }, cond: { firstRow: { rPr: { b: true }, tcPr: { borders: { bottom: BD('single', 1.5) } } }, firstCol: { rPr: { b: true } } } },
    'Table Grid 7': { tblPr: { borders: all(BD('single', 1.5)) }, rPr: { b: true }, cond: { firstRow: { rPr: { b: true, i: true }, tcPr: { borders: { bottom: BD('double', 0.75) } } } } },
    'Table Grid 8': { tblPr: { borders: all(BD('single', 0.75, '000080')) }, cond: { firstRow: { rPr: { b: true, color: 'FFFFFF' }, tcPr: { shd: SH('000080', 'solid', '000080') } }, lastRow: { rPr: { b: true } } } },
    'Table List 1': { tblPr: { borders: { top: BD('single', 1.5, '008000'), bottom: BD('single', 1.5, '008000') } }, cond: { firstRow: { rPr: { b: true, color: 'FFFFFF' }, tcPr: { shd: SH('008000', 'solid', '008000') } }, band2Horz: { tcPr: { shd: SH('FFFFC0', 'clear') } } } },
    'Table List 2': { tblPr: { borders: { bottom: BD('single', 1.5, '808080') } }, cond: { firstRow: { rPr: { b: true }, tcPr: { borders: { bottom: BD('single', 1.5, '808080') } } }, band1Horz: { tcPr: { shd: SH('C0C0C0', 'clear') } }, band2Horz: { tcPr: { shd: SH('E0FFE0', 'clear') } } } },
    'Table List 3': { tblPr: { borders: { top: BD('single', 1.5, '000080'), bottom: BD('single', 1.5, '000080') } }, cond: { firstRow: { rPr: { b: true, color: '000080' }, tcPr: { borders: { bottom: BD('single', 1.5, '000080') } } }, lastRow: { tcPr: { borders: { top: BD('single', 0.75, '000080') } } } } },
    'Table List 4': { tblPr: { borders: all(BD('single', 0.75)) }, cond: { firstRow: { rPr: { b: true, color: 'FFFFFF' }, tcPr: { shd: SH('000000', 'solid', '000000') } }, band1Horz: { tcPr: { shd: SH('C0C0C0', 'clear') } } } },
    'Table List 5': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5), left: BD('single', 1.5), right: BD('single', 1.5), insideV: BD('single', 0.75) } }, cond: { firstRow: { rPr: { b: true, i: true }, tcPr: { borders: { bottom: BD('single', 1.5) } } }, firstCol: { rPr: { b: true } } } },
    'Table List 6': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5) } }, cond: { firstRow: { rPr: { b: true }, tcPr: { borders: { bottom: BD('single', 1.5) } } }, band1Horz: { tcPr: { shd: SH('E6E6E6', 'clear') } }, firstCol: { rPr: { b: true } } } },
    'Table List 7': { tblPr: { borders: { top: BD('single', 1.5, '008000'), bottom: BD('single', 1.5, '008000') } }, cond: { firstRow: { rPr: { b: true, color: '008000' }, tcPr: { borders: { bottom: BD('single', 0.75, '008000') } } }, band1Horz: { tcPr: { shd: SH('C0FFC0', 'clear') } } } },
    'Table List 8': { tblPr: { borders: { top: BD('single', 1.5), bottom: BD('single', 1.5) } }, cond: { firstRow: { rPr: { b: true, color: 'FFFFFF' }, tcPr: { shd: SH('000080', 'solid', '000080') } }, band1Horz: { tcPr: { shd: SH('C0C0FF', 'clear') } }, band2Horz: { tcPr: { shd: SH('FFFFC0', 'clear') } } } },
    'Table 3D effects 1': { tcPr: { shd: SH('C0C0C0', 'clear') }, tblPr: { borders: { top: BD('threeDEmboss', 1.5), left: BD('threeDEmboss', 1.5), bottom: BD('threeDEngrave', 1.5), right: BD('threeDEngrave', 1.5), insideH: BD('single', 0.75, '808080'), insideV: BD('single', 0.75, '808080') } }, cond: { firstRow: { rPr: { b: true } } } },
    'Table 3D effects 2': { tcPr: { shd: SH('C0C0C0', 'clear') }, tblPr: { borders: all(BD('threeDEmboss', 1.5)) }, cond: { firstRow: { rPr: { b: true, color: '000080' } }, firstCol: { rPr: { b: true } } } },
    'Table 3D effects 3': { tblPr: { borders: { top: BD('threeDEmboss', 3), left: BD('threeDEmboss', 3), bottom: BD('threeDEngrave', 3), right: BD('threeDEngrave', 3) } }, cond: { firstRow: { rPr: { b: true }, tcPr: { shd: SH('C0C0C0', 'clear') } }, band1Horz: { tcPr: { shd: SH('E6E6E6', 'clear') } } } },
    'Table Contemporary': { tblPr: { borders: { insideH: BD('single', 3, 'FFFFFF'), insideV: BD('single', 3, 'FFFFFF') } }, cond: { firstRow: { rPr: { b: true }, tcPr: { shd: SH('D9D9D9', 'clear') } }, band1Horz: { tcPr: { shd: SH('F2F2F2', 'clear') } }, band2Horz: { tcPr: { shd: SH('E6E6E6', 'clear') } } } },
    'Table Elegant': { tblPr: { borders: { top: BD('double', 0.5), bottom: BD('double', 0.5), left: BD('double', 0.5), right: BD('double', 0.5), insideV: BD('single', 0.5) } }, cond: { firstRow: { rPr: { caps: true }, tcPr: { borders: { bottom: BD('single', 0.5) } } } } },
    'Table Professional': { tblPr: { borders: all(BD('single', 0.5)) }, cond: { firstRow: { rPr: { b: true, color: 'FFFFFF' }, tcPr: { shd: SH('000000', 'solid', '000000') } } } },
    'Table Subtle 1': { tblPr: { borders: { top: BD('single', 1.5, '008000'), bottom: BD('single', 1.5, '008000') } }, cond: { firstRow: { tcPr: { borders: { bottom: BD('single', 0.75, '008000') } } }, band1Horz: { tcPr: { shd: SH('E0FFE0', 'clear') } }, firstCol: { tcPr: { shd: SH('C0FFC0', 'clear') } } } },
    'Table Subtle 2': { tblPr: { borders: { left: BD('single', 1.5), right: BD('single', 1.5) } }, cond: { firstRow: { tcPr: { borders: { bottom: BD('single', 0.75) } } }, lastCol: { tcPr: { shd: SH('C0FFC0', 'clear') } }, band1Horz: { tcPr: { shd: SH('E6E6E6', 'clear') } } } },
    'Table Web 1': { tblPr: { cellSpacing: 1.5, borders: { top: BD('outset', 0.75), left: BD('outset', 0.75), bottom: BD('outset', 0.75), right: BD('outset', 0.75), insideH: BD('inset', 0.75), insideV: BD('inset', 0.75) } }, cond: {} },
    'Table Web 2': { tblPr: { cellSpacing: 1.5, borders: all(BD('inset', 0.75)) }, cond: {} },
    'Table Web 3': { tblPr: { cellSpacing: 1.5, borders: all(BD('outset', 2.25)) }, cond: {} },
    'Table Grid': null,
  };
  T.styleId = (name) => (name === 'Table Grid' ? 'TableGrid' : name.replace(/[^A-Za-z0-9]/g, ''));
  T.ensureStyle = function (idOrName) {
    const d = doc();
    if (d.styles[idOrName]) return idOrName;
    const name = Object.keys(T.STYLES).find((n) => T.styleId(n) === idOrName || n === idOrName);
    if (!name) return null;
    const id = T.styleId(name);
    if (d.styles[id]) return id;
    const def = T.STYLES[name];
    if (!def) return 'TableGrid';
    D.touchKey(d, 'styles');
    d.styles[id] = { id, name, type: 'table', basedOn: 'TableNormal', builtin: true, tblPr: Object.assign({ cellMar: { l: 5.4, r: 5.4, t: 0, b: 0 } }, L.clone(def.tblPr || {})), pPr: {}, rPr: L.clone(def.rPr || {}), tcPr: L.clone(def.tcPr || {}), cond: L.clone(def.cond || {}) };
    D.stylesChanged();
    return id;
  };
  T.applyStyle = function (name, look, opts) {
    const ci = T.cur();
    if (!ci) return;
    const t = ci.tbl;
    E.edit('Table AutoFormat', () => {
      const id = name === 'Table Normal' ? 'TableNormal' : T.ensureStyle(T.styleId(name)) || T.ensureStyle(name);
      D.touchTbl(t);
      t.tblPr.style = id;
      if (look) t.tblPr.look = Object.assign({}, look);
      if (opts && opts.clearDirect) { delete t.tblPr.borders; for (const r of t.rows) for (const c of r.cells) { delete c.tcPr.borders; delete c.tcPr.shd; } }
      return E.sel;
    });
  };

  /* ================= Draw Table pencil, eraser, and border dragging ================= */
  T.tool = null;
  T.setTool = function (kind) {
    T.tool = kind;
    const sc = L.$('#scroller');
    sc.classList.toggle('tool-pen', kind === 'draw');
    sc.classList.toggle('tool-eraser', kind === 'erase');
    if (kind && A().tbShown && !A().tbShown('tables')) { A().opts.toolbars.tables = true; A().updateToolbars(); }
    L.ui.refresh();
  };
  const ptPerPx = () => 1 / (LY.zoom * D.PX);
  let drag = null;
  function hitBorder(e) {
    const td = e.target.closest && e.target.closest('td[data-cid]');
    if (!td || td.closest('.fcode')) return null;
    const r = td.getBoundingClientRect();
    const tol = 4;
    const tblEl = td.closest('table.tbl');
    if (Math.abs(e.clientX - r.right) <= tol) return { kind: 'col', td, side: 'right', tblEl };
    if (Math.abs(e.clientX - r.left) <= tol && +td.dataset.ci > 0) {
      const prev = td.previousElementSibling;
      if (prev && prev.dataset.cid) return { kind: 'col', td: prev, side: 'right', tblEl };
    }
    if (Math.abs(e.clientY - r.bottom) <= tol) return { kind: 'row', td, tblEl };
    return null;
  }
  function tableOf(tblEl) { return tblEl ? findTable(+tblEl.dataset.tid) : null; }
  function findTable(id) {
    const d = doc();
    if (d._idxDirty || !d._tbls) D.reindex(d);
    const r = d._tbls.get(id);
    return r ? r.tbl : null;
  }
  T.findTable = findTable;
  function cellById(t, id) { for (let ri = 0; ri < t.rows.length; ri++) { const ci = t.rows[ri].cells.findIndex((c) => c.id === id); if (ci >= 0) return { ri, ci, cell: t.rows[ri].cells[ci] }; } return null; }
  function onMove(e) {
    if (drag || T.tool || !LY.root || E.readOnly || (L.drawtool && L.drawtool.active)) return;
    const hb = hitBorder(e);
    const sc = L.$('#scroller');
    sc.classList.toggle('col-resize', !!(hb && hb.kind === 'col'));
    sc.classList.toggle('row-resize', !!(hb && hb.kind === 'row'));
  }
  function onDown(e) {
    if (e.button !== 0 || !LY.root || E.readOnly || (L.drawtool && L.drawtool.active)) return;
    if (T.tool) { toolDown(e); return; }
    const hb = hitBorder(e);
    if (!hb) return;
    const t = findTable(+hb.tblEl.dataset.tid);
    if (!t) return;
    e.preventDefault(); e.stopPropagation();
    const c = cellById(t, +hb.td.dataset.cid);
    if (!c) return;
    const map = R.tblMap(t);
    const m = map[c.ri].find((x) => x.cell === c.cell);
    const sc = L.$('#scroller');
    const guide = L.h('div', { class: 'r-guide' + (hb.kind === 'row' ? ' h' : '') });
    sc.appendChild(guide);
    const sr = sc.getBoundingClientRect();
    const place = (x, y) => { if (hb.kind === 'col') { guide.style.cssText = `left:${x - sr.left + sc.scrollLeft}px;top:${sc.scrollTop}px;height:${sc.clientHeight}px`; } else guide.style.cssText = `top:${y - sr.top + sc.scrollTop}px;left:${sc.scrollLeft}px;width:${sc.clientWidth}px;border-left:0;border-top:1px dotted #000;height:0`; };
    place(e.clientX, e.clientY);
    const x0 = e.clientX, y0 = e.clientY;
    drag = { hb };
    const mv = (ev) => place(ev.clientX, ev.clientY);
    const up = (ev) => {
      window.removeEventListener('pointermove', mv, true);
      window.removeEventListener('pointerup', up, true);
      guide.remove();
      drag = null;
      const dxPt = (ev.clientX - x0) * ptPerPx(), dyPt = (ev.clientY - y0) * ptPerPx();
      if (hb.kind === 'col' && Math.abs(dxPt) > 0.5) {
        const k = m.c0 + m.span - 1;
        E.edit('Column Width', () => {
          D.touchTbl(t);
          const shift = !ev.shiftKey;
          if (k + 1 < t.grid.length && shift) {
            const d = L.clamp(dxPt, -(t.grid[k] - 6), t.grid[k + 1] - 6);
            t.grid[k] = L.round(t.grid[k] + d, 2); t.grid[k + 1] = L.round(t.grid[k + 1] - d, 2);
          } else t.grid[k] = L.round(Math.max(6, t.grid[k] + dxPt), 2);
          if (t.tblPr.w && t.tblPr.w.type === 'pct') t.tblPr.w = { type: 'auto', v: 0 };
          fixWidths(t);
          return E.sel;
        });
      } else if (hb.kind === 'row' && Math.abs(dyPt) > 0.5) {
        const tr = hb.td.closest('tr');
        const cur = tr.getBoundingClientRect().height * ptPerPx();
        const ri = c.ri + (m.rowspan || 1) - 1;
        E.edit('Row Height', () => { D.touchTbl(t); t.rows[ri].trPr.h = L.round(Math.max(6, cur + dyPt), 2); if (t.rows[ri].trPr.hRule !== 'exact') t.rows[ri].trPr.hRule = 'atLeast'; return E.sel; });
      }
    };
    window.addEventListener('pointermove', mv, true);
    window.addEventListener('pointerup', up, true);
  }
  /* pencil & eraser */
  function toolDown(e) {
    e.preventDefault(); e.stopPropagation();
    const sc = L.$('#scroller');
    const sr = sc.getBoundingClientRect();
    const x0 = e.clientX, y0 = e.clientY;
    if (T.tool === 'erase') {
      const hb = hitBorder(e) || (() => { const td = e.target.closest && e.target.closest('td[data-cid]'); if (!td) return null; const r = td.getBoundingClientRect(); const dl = e.clientX - r.left, dr = r.right - e.clientX, dt = e.clientY - r.top, db = r.bottom - e.clientY; const mn = Math.min(dl, dr, dt, db); if (mn > 10) return null; return { td, tblEl: td.closest('table.tbl'), kind: mn === dr || mn === dl ? 'col' : 'row', side: mn === dl ? 'left' : mn === dr ? 'right' : mn === dt ? 'top' : 'bottom' }; })();
      if (!hb) return;
      const t = findTable(+hb.tblEl.dataset.tid);
      if (!t) return;
      const c = cellById(t, +hb.td.dataset.cid);
      const map = R.tblMap(t);
      const m = map[c.ri].find((x) => x.cell === c.cell);
      /* erase the border between this cell and its neighbour: merge the pair */
      let other = null;
      if (hb.kind === 'col') { const c2 = hb.side === 'left' ? m.c0 - 1 : m.c0 + m.span; other = map[c.ri].find((x) => x.c0 <= c2 && c2 < x.c0 + x.span); }
      else { const r2 = hb.side === 'top' ? c.ri - 1 : c.ri + (m.rowspan || 1); other = map[r2] && map[r2].find((x) => x.c0 === m.c0 && x.span === m.span); }
      if (!other || other.hidden) return;
      const fa = D.firstPara(c.cell.blocks), lb = D.lastPara(other.cell.blocks);
      E.setSel(D.pos(fa, 0), D.pos(lb, D.plen(lb)));
      E.markCells();
      T.merge();
      return;
    }
    /* pencil: draw a rectangle (new table) or a line inside a cell (split) */
    const ghost = L.h('div', { class: 'draw-ghost' });
    sc.appendChild(ghost);
    const startTd = e.target.closest && e.target.closest('td[data-cid]');
    const upd = (ev) => {
      const l = Math.min(x0, ev.clientX), t0 = Math.min(y0, ev.clientY);
      ghost.style.cssText = `left:${l - sr.left + sc.scrollLeft}px;top:${t0 - sr.top + sc.scrollTop}px;width:${Math.abs(ev.clientX - x0)}px;height:${Math.abs(ev.clientY - y0)}px`;
    };
    upd(e);
    const mv = (ev) => upd(ev);
    const up = (ev) => {
      window.removeEventListener('pointermove', mv, true);
      window.removeEventListener('pointerup', up, true);
      ghost.remove();
      const w = Math.abs(ev.clientX - x0) * ptPerPx(), hh = Math.abs(ev.clientY - y0) * ptPerPx();
      if (w < 4 && hh < 4) return;
      if (startTd) {
        const t = findTable(+startTd.closest('table.tbl').dataset.tid);
        const c = t && cellById(t, +startTd.dataset.cid);
        if (!c) return;
        const fp = D.firstPara(c.cell.blocks);
        E.setSel(D.pos(fp, 0));
        if (w > hh) { E.edit('Draw Table', () => { D.touchTbl(t); splitRows(t, c.cell, 2); fixWidths(t); return E.sel; }); }
        else {
          /* vertical line: split at the drawn x */
          const r = startTd.getBoundingClientRect();
          const frac = L.clamp((Math.min(x0, ev.clientX) + Math.abs(ev.clientX - x0) / 2 - r.left) / r.width, 0.1, 0.9);
          E.edit('Draw Table', () => {
            D.touchTbl(t);
            const map = R.tblMap(t);
            const m = map[c.ri].find((x) => x.cell === c.cell);
            const xs = xsOf(t.grid);
            const cut = L.round(xs[m.c0] + (xs[m.c0 + m.span] - xs[m.c0]) * frac, 2);
            regrid(t, [cut]);
            const map2 = R.tblMap(t), xs2 = xsOf(t.grid);
            const m2 = map2[c.ri].find((x) => x.cell === c.cell);
            const ci2 = xs2.findIndex((v) => Math.abs(v - cut) < 0.05);
            const leftSpan = ci2 - m2.c0, rightSpan = m2.span - leftSpan;
            const nc = newCellLike(c.cell, rightSpan);
            if (leftSpan > 1) c.cell.tcPr.span = leftSpan; else delete c.cell.tcPr.span;
            t.rows[c.ri].cells.splice(t.rows[c.ri].cells.indexOf(c.cell) + 1, 0, nc);
            fixWidths(t);
            return E.sel;
          });
        }
        return;
      }
      /* outside tables: a new one-cell table at the paragraph under the drag start */
      const pos = E.posAtPoint(x0, y0) || E.sel && E.sel.f;
      if (!pos) return;
      E.setSel(pos);
      T.insert(1, 1, { width: Math.max(18, w), colWidth: Math.max(18, w) });
      const ci = T.cur();
      if (ci && hh > 14) E.edit('Draw Table', () => { D.touchTbl(ci.tbl); ci.tbl.rows[0].trPr.h = L.round(hh, 2); ci.tbl.rows[0].trPr.hRule = 'atLeast'; return E.sel; });
    };
    window.addEventListener('pointermove', mv, true);
    window.addEventListener('pointerup', up, true);
  }
  void tableOf;
  T.mount = function () {
    const sc = L.$('#scroller');
    sc.addEventListener('pointermove', onMove);
    sc.addEventListener('pointerdown', onDown, true);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && T.tool) T.setTool(null); });
  };
  L.bus.on('app-ready', () => T.mount());
})();
