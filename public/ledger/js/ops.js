/* Ledger — model operations with undo: cell input, formatting, rows/columns, sheets, sort, fill, merge.
 * Everything here works on the model only (no DOM); the grid controller repaints afterwards. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, C = L.calc, F = L.formula, NF = L.numfmt, LY = L.layout;
  const O = (L.ops = {});
  const MAXR = M.MAXR, MAXC = M.MAXC;

  /* ------------------------------------------------------------ transactions */
  let changes = null, structural = false;
  /** run fn inside one undoable step; recalculates what changed afterwards */
  O.tx = function (wb, label, fn) {
    const outer = changes == null;
    const pivots = outer && L.pivots?.begin(wb);
    if (outer) { changes = []; structural = false; }
    wb.undo.begin(label);
    let res;
    try { res = fn(); }
    finally {
      if (outer) L.pivots?.end(wb, pivots);
      wb.undo.end();
      if (outer) {
        const list = changes, st = structural;
        changes = null; structural = false;
        O.recalc(wb, list, st);
        if (L.bus) L.bus.emit('changed', { wb, label, structural: st });
      }
    }
    return res;
  };
  O.markStructural = () => { structural = true; };
  O.recalc = function (wb, list, full) {
    if (wb.calcPr.mode === 'manual' && !full) { wb.needsCalc = true; return; }
    try {
      if (full || !wb.graph) { C.rebuild(wb); markAllDirty(wb); C.evaluate(Array.from(wb.graph.byCell.values()).filter((n) => n.cell.dirty)); }
      else if (list && list.length) C.changed(wb, list);
    } catch (e) { if (typeof console !== 'undefined') console.error('recalc', e); }
    if (L.cf) L.cf.reset();
  };
  function markAllDirty(wb) { for (const n of wb.graph.byCell.values()) if (n.cell.f != null) n.cell.dirty = true; }
  const touch = (sh, r, c) => { if (changes) changes.push({ sh, r, c }); };
  const touchArea = (sh, rg) => { if (changes) changes.push({ sh, r: rg.r1, c: rg.c1, area: rg }); };
  O.touch = touch;

  /* ------------------------------------------------------------ cells */
  /** replace a cell (recorded for undo); null deletes */
  O.put = function (sh, r, c, cell) {
    sh.wb.undo.note(sh, r, c);
    const old = sh.get(r, c);
    L.pivots?.cell(sh, r, c, old, cell);
    if (old && old._gnode && sh.wb.graph) sh.wb.graph.remove(old);
    if (cell && cell.f == null) { delete cell.ast; delete cell._fsrc; delete cell._gnode; }
    sh.put(r, c, cell);
    if (cell && cell.f != null && sh.wb.graph) { cell.dirty = true; sh.wb.graph.add(sh, r, c, cell); }
    touch(sh, r, c);
  };
  /** clone a cell for modification */
  const dup = (cell) => (cell ? Object.assign({}, cell) : { v: null });
  O.dup = dup;
  /** a cell keeps its style even when emptied */
  function emptyKeepStyle(cell) {
    if (!cell || !cell.s) return null;
    return { v: null, s: cell.s };
  }
  /**
   * Parse what the user typed. Returns {kind:'formula', f} | {kind:'value', v, fmt}
   */
  O.parseInput = function (wb, text, opts) {
    opts = opts || {};
    if (text == null) return { kind: 'value', v: null };
    const s = String(text);
    if (s === '') return { kind: 'value', v: null };
    if (s[0] === "'") return { kind: 'value', v: s.slice(1), quoted: true };
    if (opts.textFormat) return { kind: 'value', v: s };
    if (s[0] === '=' && s.length > 1) return { kind: 'formula', f: s.slice(1) };
    if ((s[0] === '+' || s[0] === '-') && s.length > 1) {
      const p = NF.parseInput(s, { date1904: wb.date1904 });
      if (p && typeof p.v === 'number') return { kind: 'value', v: p.v, fmt: p.fmt };
      return { kind: 'formula', f: s };
    }
    if (s[0] === '@' && /^@[A-Za-z]/.test(s)) return { kind: 'formula', f: s.slice(1) };
    const up = s.trim().toUpperCase();
    if (up === 'TRUE' || up === 'FALSE') return { kind: 'value', v: up === 'TRUE' };
    if (F.ERRORS.includes(up)) return { kind: 'value', v: M.err(up) };
    const p = NF.parseInput(s, { date1904: wb.date1904 });
    if (p && typeof p.v === 'number') return { kind: 'value', v: p.v, fmt: p.fmt };
    return { kind: 'value', v: s };
  };
  /** check and normalise a typed formula; returns {f (store form)} or {error, pos, fixed?} */
  O.checkFormula = function (text) {
    let t = String(text);
    try { F.parse(F.toStore(t)); return { f: F.toStore(t) }; } catch (e) {
      /* the classic correction: close open parentheses */
      let depth = 0, q = false;
      for (const ch of t) { if (ch === '"') q = !q; else if (!q && ch === '(') depth++; else if (!q && ch === ')') depth--; }
      if (depth > 0 && !q) { const fixed = t + ')'.repeat(depth); try { F.parse(F.toStore(fixed)); return { error: e.message, pos: e.pos, fixed }; } catch (e2) { /* no */ } }
      if (q) { const fixed = t + '"' + (depth > 0 ? ')'.repeat(depth) : ''); try { F.parse(F.toStore(fixed)); return { error: e.message, pos: e.pos, fixed }; } catch (e3) { /* no */ } }
      return { error: e.message, pos: e.pos };
    }
  };
  /**
   * Enter text into a cell as if typed. opts: {array: range} for Ctrl+Shift+Enter, {keepFormat}
   */
  O.enter = function (sh, r, c, text, opts) {
    opts = opts || {};
    const wb = sh.wb;
    const old = sh.get(r, c);
    const st = LY.styleOf(sh, r, c, old);
    const textFmt = st.nf === '@';
    const p = O.parseInput(wb, text, { textFormat: textFmt });
    if (p.kind === 'formula') {
      const chk = O.checkFormula(p.f);
      if (chk.error) throw Object.assign(new Error(chk.error), { pos: chk.pos, fixed: chk.fixed, code: 'formula' });
      if (opts.array) return O.setArrayFormula(sh, opts.array, chk.f);
      const cell = dup(old);
      delete cell.rt; delete cell.af; delete cell.am; delete cell.dt;
      cell.f = chk.f; cell.v = cell.v === undefined ? null : cell.v; cell.edited = true;
      /* a formula whose first referenced cell has a number format adopts it when the cell is General */
      if ((!st.nf || st.nf === 'General') && !opts.keepFormat) {
        const nf = formatFromFormula(sh, r, c, chk.f);
        if (nf) cell.s = wb.styles.derive(cell.s || LY.styleIdOf(sh, r, c, old), { nf });
      }
      O.put(sh, r, c, cell);
      return cell;
    }
    let cell = dup(old);
    delete cell.f; delete cell.rt; delete cell.af; delete cell.am; delete cell.dt; delete cell.ast; delete cell._fsrc; delete cell.edited;
    cell.v = p.v;
    if (p.v == null) { O.put(sh, r, c, emptyKeepStyle(old) || (old && old.s ? { v: null, s: old.s } : null)); return null; }
    if (p.fmt && (!st.nf || st.nf === 'General' || (NF.isDate(p.fmt) !== NF.isDate(st.nf) && NF.category(st.nf) !== 'custom'))) {
      cell.s = wb.styles.derive(cell.s != null ? cell.s : LY.styleIdOf(sh, r, c, old), { nf: p.fmt });
    }
    if (typeof p.v === 'string' && /\n/.test(p.v)) cell.s = wb.styles.derive(cell.s != null ? cell.s : LY.styleIdOf(sh, r, c, old), { align: { wrap: true } });
    if (cell.s == null) { const sid = LY.styleIdOf(sh, r, c, old); if (sid) cell.s = sid; }
    O.put(sh, r, c, cell);
    return cell;
  };
  function formatFromFormula(sh, r, c, f) {
    let ast; try { ast = F.parse(f); } catch (e) { return null; }
    /* =A1+1 with A1 a date shows a date; =NOW() shows date and time; =TODAY() a date */
    if (ast.t === 'fn' && ast.name === 'NOW') return 'm/d/yyyy h:mm';
    if (ast.t === 'fn' && (ast.name === 'TODAY' || ast.name === 'DATE' || ast.name === 'EDATE' || ast.name === 'EOMONTH' || ast.name === 'WORKDAY' || ast.name === 'DATEVALUE')) return ast.name === 'EOMONTH' || ast.name === 'EDATE' ? null : 'm/d/yyyy';
    if (ast.t === 'fn' && (ast.name === 'TIME' || ast.name === 'TIMEVALUE')) return 'h:mm AM/PM';
    let first = null;
    F.walk(ast, (n) => { if (!first && (n.t === 'ref' || n.t === 'area') && n.sheet == null) first = n; });
    if (!first) return null;
    if (ast.t === 'fn' && !/^(SUM|AVERAGE|MIN|MAX|ROUND|ABS)$/.test(ast.name)) return null;
    const rr = first.t === 'ref' ? first.r : first.r1, cc = first.t === 'ref' ? first.c : first.c1;
    if (rr === r && cc === c) return null;
    const st = LY.styleOf(sh, rr, cc, sh.get(rr, cc));
    if (st.nf && st.nf !== 'General' && (NF.isDate(st.nf) || /[$%€£¥]/.test(st.nf) || /#,##0/.test(st.nf))) return st.nf;
    return null;
  }
  /** Ctrl+Shift+Enter */
  O.setArrayFormula = function (sh, rg, f) {
    /* an existing array that overlaps must be the same block */
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      const cl = sh.get(r, c);
      if (cl && cl.am && !(cl.am.r === rg.r1 && cl.am.c === rg.c1)) {
        const m = sh.get(cl.am.r, cl.am.c);
        if (m && m.af && !(m.af.r1 === rg.r1 && m.af.c1 === rg.c1 && m.af.r2 === rg.r2 && m.af.c2 === rg.c2)) throw Object.assign(new Error('You cannot change part of an array.'), { code: 'array' });
      }
    }
    for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
      const cl = dup(sh.get(r, c));
      delete cl.f; delete cl.rt; delete cl.af; delete cl.am; delete cl.dt;
      cl.v = null;
      if (r === rg.r1 && c === rg.c1) { cl.f = f; cl.af = { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }; cl.edited = true; }
      else cl.am = { r: rg.r1, c: rg.c1 };
      O.put(sh, r, c, cl);
    }
    return sh.get(rg.r1, rg.c1);
  };
  /** the array block containing (r, c), or null */
  O.arrayAt = function (sh, r, c) {
    const cl = sh.get(r, c);
    if (!cl) return null;
    if (cl.af) return cl.af;
    if (cl.am) { const m = sh.get(cl.am.r, cl.am.c); return m && m.af ? m.af : null; }
    return null;
  };
  /** would changing these ranges cut through an array formula? */
  O.cutsArray = function (sh, ranges) {
    for (const rg of ranges) {
      let hit = null;
      sh.each(rg.r1, rg.c1, rg.r2, rg.c2, (cl, r, c) => {
        const a = O.arrayAt(sh, r, c);
        if (a && !(a.r1 >= rg.r1 && a.r2 <= rg.r2 && a.c1 >= rg.c1 && a.c2 <= rg.c2)) { hit = a; return false; }
      });
      if (hit) return hit;
    }
    return null;
  };
  /** set a plain value (no parsing) */
  O.setValue = function (sh, r, c, v, keep) {
    const old = sh.get(r, c);
    if (v == null && !keep) { O.put(sh, r, c, emptyKeepStyle(old)); return; }
    const cell = dup(old);
    delete cell.f; delete cell.rt; delete cell.af; delete cell.am; delete cell.dt; delete cell.edited;
    cell.v = v;
    O.put(sh, r, c, cell);
  };
  O.setFormula = function (sh, r, c, f) {
    const cell = dup(sh.get(r, c));
    delete cell.rt; delete cell.af; delete cell.am; delete cell.dt;
    cell.f = f; cell.edited = true;
    O.put(sh, r, c, cell);
  };

  /* ------------------------------------------------------------ iteration helpers */
  /** every cell position in ranges that has a stored cell; fn(cell, r, c) */
  O.eachStored = function (sh, ranges, fn) { for (const rg of ranges) sh.each(rg.r1, rg.c1, Math.min(rg.r2, sh.maxR), Math.min(rg.c2, sh.maxC), fn); };
  /** used extent clipped range (whole rows/columns are limited to the used area) */
  O.clip = function (sh, rg) {
    return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(rg.r1, sh.maxR)), c2: Math.min(rg.c2, Math.max(rg.c1, sh.maxC)) };
  };

  /* ------------------------------------------------------------ clear */
  O.clear = function (sh, ranges, what) {
    what = what || 'contents';
    const wb = sh.wb;
    if (what !== 'formats' && what !== 'comments') { const a = O.cutsArray(sh, ranges); if (a) throw Object.assign(new Error('You cannot change part of an array.'), { code: 'array' }); }
    O.tx(wb, what === 'all' ? 'Clear' : what === 'formats' ? 'Clear Formats' : what === 'comments' ? 'Clear Comments' : 'Clear Contents', () => {
      for (const rg0 of ranges) {
        const whole = rg0.c1 === 0 && rg0.c2 >= MAXC - 1 || rg0.r1 === 0 && rg0.r2 >= MAXR - 1;
        const rg = O.clip(sh, rg0);
        if (what === 'contents' || what === 'all') {
          sh.each(rg.r1, rg.c1, rg.r2, rg.c2, (cell, r, c) => {
            if (cell.v == null && cell.f == null && !cell.am) { if (what === 'all') O.put(sh, r, c, null); return; }
            O.put(sh, r, c, what === 'all' ? null : emptyKeepStyle(cell));
          });
        }
        if (what === 'formats' || what === 'all') {
          sh.each(rg.r1, rg.c1, rg.r2, rg.c2, (cell, r, c) => {
            if (!cell.s) return;
            const n = dup(cell); delete n.s;
            O.put(sh, r, c, n.v == null && n.f == null && !n.am ? null : n);
          });
          /* row / column styles of whole rows / columns */
          if (whole) O.clearLineStyles(sh, rg0);
          /* merged cells and conditional formats go with "formats" */
          const kept = sh.merges.filter((m) => !(m.r1 >= rg.r1 && m.r2 <= rg.r2 && m.c1 >= rg.c1 && m.c2 <= rg.c2));
          if (kept.length !== sh.merges.length) O.setMerges(sh, kept);
          O.removeRangeRules(sh, rg0, 'cf');
        }
        if (what === 'comments' || what === 'all') {
          for (const cm of Array.from(sh.comments.values())) if (M.rangeContains(rg0, cm.r, cm.c)) O.setComment(sh, cm.r, cm.c, null);
        }
        if (what === 'all' || what === 'hyperlinks') {
          const keep = sh.links.filter((l) => !M.rangesOverlap(l.ref, rg0));
          if (keep.length !== sh.links.length) O.setLinks(sh, keep);
        }
      }
    });
  };
  O.clearLineStyles = function (sh, rg) {
    const before = { rows: sh.rows.map((r) => r && r.s), cols: sh.cols.map((c) => c && c.s) };
    const apply = () => {
      if (rg.c1 === 0 && rg.c2 >= MAXC - 1) for (let r = rg.r1; r <= Math.min(rg.r2, sh.rows.length - 1); r++) if (sh.rows[r]) delete sh.rows[r].s;
      if (rg.r1 === 0 && rg.r2 >= MAXR - 1) for (let c = rg.c1; c <= Math.min(rg.c2, sh.cols.length - 1); c++) if (sh.cols[c]) delete sh.cols[c].s;
    };
    apply();
    sh.wb.undo.op(() => { before.rows.forEach((s, r) => { if (sh.rows[r]) { if (s) sh.rows[r].s = s; else delete sh.rows[r].s; } }); before.cols.forEach((s, c) => { if (sh.cols[c]) { if (s) sh.cols[c].s = s; else delete sh.cols[c].s; } }); }, apply);
  };

  /* ------------------------------------------------------------ formatting */
  /**
   * Apply a style change to ranges. delta: partial style ({font:{b:true}}, {nf:'0.00'}, {fill:...}, {align:{h:'center'}})
   * or a function (style) → delta. Whole rows / columns also set row / column styles.
   */
  O.format = function (sh, ranges, delta, label) {
    const wb = sh.wb;
    const cache = new Map();
    const conv = (sid) => {
      let o = cache.get(sid);
      if (o == null) {
        const d = typeof delta === 'function' ? delta(wb.styles.get(sid), sid) : delta;
        o = d ? wb.styles.derive(sid, d) : sid;
        cache.set(sid, o);
      }
      return o;
    };
    O.tx(wb, label || 'Format Cells', () => {
      for (const rg of ranges) {
        const allCols = rg.c1 === 0 && rg.c2 >= MAXC - 1, allRows = rg.r1 === 0 && rg.r2 >= MAXR - 1;
        if (allCols || allRows) {
          /* line styles for whole rows / columns, plus every stored cell */
          lineStyles(sh, rg, conv, allCols, allRows);
          sh.each(rg.r1, rg.c1, rg.r2, rg.c2, (cell, r, c) => { const n = dup(cell); n.s = conv(cell.s || 0); O.put(sh, r, c, n); });
          continue;
        }
        for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
          const cell = sh.get(r, c);
          const sid = cell && cell.s != null ? cell.s : LY.styleIdOf(sh, r, c, cell);
          const ns = conv(sid);
          if (cell) { if (ns === cell.s) continue; const n = dup(cell); n.s = ns; O.put(sh, r, c, n); }
          else if (ns !== sid || ns) { O.put(sh, r, c, { v: null, s: ns }); }
        }
      }
    });
  };
  function lineStyles(sh, rg, conv, allCols, allRows) {
    const snap = () => ({ rows: allCols ? rangeSnap(sh.rows, rg.r1, rg.r2, 'row') : null, cols: allRows ? rangeSnap(sh.cols, rg.c1, rg.c2, 'col') : null, tail: sh.colTail ? JSON.parse(JSON.stringify(sh.colTail)) : null });
    const before = snap();
    if (allCols) for (let r = rg.r1; r <= Math.min(rg.r2, allRows ? Math.max(sh.maxR, rg.r1) : rg.r2); r++) { const row = sh.rowObj(r); row.s = conv(row.s || 0); }
    if (allRows) {
      const last = Math.min(rg.c2, 16383);
      for (let c = rg.c1; c <= last; c++) { if (c > 1024 && !sh.cols[c]) continue; const o = sh.colObj(c); o.s = conv(o.s || 0); }
      if (rg.c2 >= MAXC - 1) sh.colTail = { from: 1025, o: { s: conv(sh.colTail && sh.colTail.o.s || 0) } };
    }
    LY.invalidate(sh);
    const after = snap();
    sh.wb.undo.op(() => restoreSnap(sh, before, rg), () => restoreSnap(sh, after, rg));
  }
  function rangeSnap(arr, a, b, kind) { const out = []; for (let i = a; i <= Math.min(b, arr.length - 1); i++) out.push(arr[i] ? Object.assign({}, arr[i], kind === 'row' ? { cells: undefined } : {}) : null); return { a, list: out }; }
  function restoreSnap(sh, s, rg) {
    if (s.rows) s.rows.list.forEach((o, k) => { const r = s.rows.a + k; const row = sh.rows[r]; if (!o) { if (row) { delete row.s; } return; } const cells = row ? row.cells : []; sh.rows[r] = Object.assign({}, o, { cells }); });
    if (s.cols) s.cols.list.forEach((o, k) => { sh.cols[s.cols.a + k] = o ? Object.assign({}, o) : undefined; });
    sh.colTail = s.tail ? JSON.parse(JSON.stringify(s.tail)) : null;
    LY.invalidate(sh);
    void rg;
  }
  /** font toggles: bold / italic / underline follow the active cell */
  O.toggleFont = function (sh, ranges, key, activeStyle) {
    const on = !(activeStyle.font && activeStyle.font[key]);
    const val = key === 'u' ? (on ? 'single' : undefined) : on || undefined;
    O.format(sh, ranges, (st) => ({ font: { [key]: val } }), { b: 'Bold', i: 'Italic', u: 'Underline', strike: 'Strikethrough' }[key]);
  };
  /**
   * Borders across ranges. spec: {outline, inside, top, bottom, left, right, insideH, insideV, all, none, diagUp, diagDown, line:{style,color}}
   */
  O.borders = function (sh, ranges, spec, label) {
    const line = spec.line || { style: 'thin' };
    const wb = sh.wb;
    O.tx(wb, label || 'Borders', () => {
      for (const rg0 of ranges) {
        const rg = O.clip(sh, rg0);
        if (rg0.r2 >= MAXR - 1 || rg0.c2 >= MAXC - 1) { if (rg.r2 - rg.r1 > 2000 || rg.c2 - rg.c1 > 200) continue; }
        for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
          const cell = sh.get(r, c);
          const sid = cell && cell.s != null ? cell.s : LY.styleIdOf(sh, r, c, cell);
          const st = wb.styles.get(sid);
          const b = Object.assign({}, st.border || {});
          const set = (k, v) => { if (v === null) delete b[k]; else b[k] = Object.assign({}, v); };
          if (spec.none) { for (const k of ['t', 'b', 'l', 'r', 'd', 'du', 'dd']) delete b[k]; }
          const top = r === rg.r1, bottom = r === rg.r2, left = c === rg.c1, right = c === rg.c2;
          if (spec.all) { set('t', line); set('b', line); set('l', line); set('r', line); }
          if (spec.outline) { if (top) set('t', line); if (bottom) set('b', line); if (left) set('l', line); if (right) set('r', line); }
          if (spec.top !== undefined && top) set('t', spec.top);
          if (spec.bottom !== undefined && bottom) set('b', spec.bottom);
          if (spec.left !== undefined && left) set('l', spec.left);
          if (spec.right !== undefined && right) set('r', spec.right);
          if (spec.insideH !== undefined) { if (!bottom) set('b', spec.insideH); if (!top) set('t', spec.insideH); }
          if (spec.insideV !== undefined) { if (!right) set('r', spec.insideV); if (!left) set('l', spec.insideV); }
          if (spec.diagUp !== undefined || spec.diagDown !== undefined) {
            const d = spec.diagUp || spec.diagDown || b.d;
            if (spec.diagUp !== undefined) b.du = !!spec.diagUp;
            if (spec.diagDown !== undefined) b.dd = !!spec.diagDown;
            if (b.du || b.dd) b.d = Object.assign({}, d); else { delete b.d; delete b.du; delete b.dd; }
          }
          const ns = wb.styles.add(Object.assign({}, st, { border: Object.keys(b).length ? b : null }));
          if (ns === sid && cell) continue;
          if (cell) { const n = dup(cell); n.s = ns; O.put(sh, r, c, n); } else if (ns !== sid) O.put(sh, r, c, { v: null, s: ns });
          /* neighbours share edges: clear the opposite side when removing */
        }
        if (spec.none || spec.outline || spec.all) {
          /* the cells just outside the range lose the shared edge (so "No Border" really clears) */
          const clearEdge = (r, c, k) => { const cell = sh.get(r, c); if (!cell || !cell.s) return; const st = wb.styles.get(cell.s); if (!st.border || !st.border[k]) return; if (!spec.none) return; const b = Object.assign({}, st.border); delete b[k]; const n = dup(cell); n.s = wb.styles.add(Object.assign({}, st, { border: Object.keys(b).length ? b : null })); O.put(sh, r, c, n); };
          if (rg.r1 > 0) for (let c = rg.c1; c <= rg.c2; c++) clearEdge(rg.r1 - 1, c, 'b');
          if (rg.r2 < MAXR - 1) for (let c = rg.c1; c <= rg.c2; c++) clearEdge(rg.r2 + 1, c, 't');
          if (rg.c1 > 0) for (let r = rg.r1; r <= rg.r2; r++) clearEdge(r, rg.c1 - 1, 'r');
          if (rg.c2 < MAXC - 1) for (let r = rg.r1; r <= rg.r2; r++) clearEdge(r, rg.c2 + 1, 'l');
        }
      }
    });
  };
  /** the style used for new cells in a range when copying formats (Format Painter) */
  O.paintFormats = function (sh, srcRange, srcSheet, dest) {
    const wb = sh.wb;
    const h = srcRange.r2 - srcRange.r1 + 1, w = srcRange.c2 - srcRange.c1 + 1;
    O.tx(wb, 'Format Painter', () => {
      const rg = dest;
      const R = Math.max(rg.r2 - rg.r1 + 1, h), Cc = Math.max(rg.c2 - rg.c1 + 1, w);
      for (let i = 0; i < R; i++) for (let j = 0; j < Cc; j++) {
        const sr = srcRange.r1 + (i % h), scl = srcRange.c1 + (j % w);
        const s = LY.styleIdOf(srcSheet, sr, scl, srcSheet.get(sr, scl));
        const r = rg.r1 + i, c = rg.c1 + j;
        const cell = sh.get(r, c);
        const sid = srcSheet.wb === wb ? s : wb.styles.add(JSON.parse(JSON.stringify(srcSheet.wb.styles.get(s))));
        if (cell) { if (cell.s === sid) continue; const n = dup(cell); n.s = sid; O.put(sh, r, c, n); } else if (sid) O.put(sh, r, c, { v: null, s: sid });
      }
      /* merged areas copy too */
      const ms = srcSheet.merges.filter((m) => m.r1 >= srcRange.r1 && m.r2 <= srcRange.r2 && m.c1 >= srcRange.c1 && m.c2 <= srcRange.c2);
      if (ms.length) {
        const add = ms.map((m) => ({ r1: m.r1 - srcRange.r1 + dest.r1, c1: m.c1 - srcRange.c1 + dest.c1, r2: m.r2 - srcRange.r1 + dest.r1, c2: m.c2 - srcRange.c1 + dest.c1 }));
        O.setMerges(sh, sh.merges.filter((m) => !add.some((a) => M.rangesOverlap(a, m))).concat(add));
      }
    });
  };

  /* ------------------------------------------------------------ simple sheet properties with undo */
  const prop = (label, get, set) => (sh, v) => { const old = get(sh); set(sh, v); sh.wb.undo.op(() => set(sh, old), () => set(sh, v)); };
  O.setMerges = prop('Merge', (sh) => sh.merges.slice(), (sh, v) => { sh.merges = v.slice(); LY.touchMerges(sh); });
  O.setLinks = prop('Hyperlink', (sh) => sh.links.slice(), (sh, v) => { sh.links = v.slice(); });
  O.setCF = prop('Conditional Formatting', (sh) => sh.cf.slice(), (sh, v) => { sh.cf = v.slice(); sh._cfIdx = null; if (L.cf) L.cf.reset(); });
  O.setDV = prop('Validation', (sh) => sh.dv.slice(), (sh, v) => { sh.dv = v.slice(); });
  /** rectangle a minus rectangle b → up to four rectangles */
  O.subtractRange = function (a, b) {
    if (!M.rangesOverlap(a, b)) return [a];
    const out = [];
    if (a.r1 < b.r1) out.push({ r1: a.r1, c1: a.c1, r2: b.r1 - 1, c2: a.c2 });
    if (a.r2 > b.r2) out.push({ r1: b.r2 + 1, c1: a.c1, r2: a.r2, c2: a.c2 });
    const r1 = Math.max(a.r1, b.r1), r2 = Math.min(a.r2, b.r2);
    if (a.c1 < b.c1) out.push({ r1, c1: a.c1, r2, c2: b.c1 - 1 });
    if (a.c2 > b.c2) out.push({ r1, c1: b.c2 + 1, r2, c2: a.c2 });
    return out;
  };
  /** take a range out of the conditional formats ('cf') or validations ('dv'); pred(rule) limits which CF rules go */
  O.removeRangeRules = function (sh, rg, kind, pred) {
    const list = kind === 'dv' ? sh.dv : sh.cf;
    let changed = false;
    const out = [];
    for (const x of list) {
      if (kind === 'cf' && pred && !x.rules.some(pred)) { out.push(x); continue; }
      const ranges = [];
      let hit = false;
      for (const r of x.ranges) { if (M.rangesOverlap(r, rg)) { hit = true; ranges.push(...O.subtractRange(r, rg)); } else ranges.push(r); }
      if (!hit) { out.push(x); continue; }
      changed = true;
      if (kind === 'cf' && pred && !x.rules.every(pred)) {
        /* the range keeps the rules that are not being replaced */
        const keep = x.rules.filter((q) => !pred(q));
        out.push(Object.assign({}, x, { rules: keep }));
        if (ranges.length) out.push(Object.assign({}, x, { ranges, rules: x.rules.filter(pred) }));
        continue;
      }
      if (ranges.length) out.push(Object.assign({}, x, { ranges }));
    }
    if (!changed) return false;
    if (kind === 'dv') O.setDV(sh, out); else O.setCF(sh, out);
    return true;
  };
  O.setAutoFilter = prop('AutoFilter', (sh) => sh.autoFilter, (sh, v) => { sh.autoFilter = v; });
  O.setDrawings = prop('Drawing', (sh) => sh.drawings.slice(), (sh, v) => { sh.drawings = v.slice(); });
  O.setPrint = prop('Page Setup', (sh) => JSON.parse(JSON.stringify(sh.print)), (sh, v) => { sh.print = JSON.parse(JSON.stringify(v)); });
  O.setProtection = prop('Protection', (sh) => sh.protection, (sh, v) => { sh.protection = v; });
  O.setFreeze = prop('Freeze Panes', (sh) => ({ f: sh.view.freeze, t: sh.view.top }), (sh, v) => { sh.view.freeze = v && v.f; sh.view.top = v && v.t ? v.t : { r: 0, c: 0 }; });
  O.setNames = function (wb, names) { const old = wb.names.slice(); wb.names = names.slice(); wb.names.forEach((n) => { n._ast = undefined; }); wb.undo.op(() => { wb.names = old; }, () => { wb.names = names.slice(); }); O.markStructural(); };
  O.setComment = function (sh, r, c, cm) {
    const k = M.key(r, c);
    const old = sh.comments.get(k) || null;
    cm = L.threads.editNote(cm);
    const apply = (v) => { if (v) sh.comments.set(k, Object.assign({}, v, { r, c })); else sh.comments.delete(k); };
    apply(cm);
    sh.wb.undo.op(() => apply(old), () => apply(cm));
  };

  /* ------------------------------------------------------------ merge */
  O.merge = function (sh, ranges, opts) {
    opts = opts || {};
    const wb = sh.wb;
    O.tx(wb, opts.center ? 'Merge and Center' : 'Merge Cells', () => {
      let merges = sh.merges.slice();
      for (const rg0 of ranges) {
        const rgs = opts.across ? Array.from({ length: rg0.r2 - rg0.r1 + 1 }, (_, i) => ({ r1: rg0.r1 + i, c1: rg0.c1, r2: rg0.r1 + i, c2: rg0.c2 })) : [rg0];
        for (const rg of rgs) {
          if (rg.r1 === rg.r2 && rg.c1 === rg.c2) continue;
          merges = merges.filter((m) => !M.rangesOverlap(m, rg));
          /* only the upper-left value survives */
          let keep = null;
          for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) {
            const cell = sh.get(r, c);
            if (!cell) continue;
            if (keep == null && (cell.v != null || cell.f != null) && !(r === rg.r1 && c === rg.c1)) keep = { r, c };
            if (!(r === rg.r1 && c === rg.c1) && (cell.v != null || cell.f != null)) O.put(sh, r, c, emptyKeepStyle(cell));
          }
          merges.push({ r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 });
          /* the merged area takes the top-left cell's format */
          const tl = sh.get(rg.r1, rg.c1);
          const sid = LY.styleIdOf(sh, rg.r1, rg.c1, tl);
          for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 2000); r++) for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 200); c++) {
            if (r === rg.r1 && c === rg.c1) continue;
            const cell = sh.get(r, c);
            if ((cell && cell.s) !== sid) O.put(sh, r, c, sid ? Object.assign(dup(cell), { s: sid, v: null }) : null);
          }
        }
      }
      O.setMerges(sh, merges);
      if (opts.center) O.format(sh, ranges, { align: { h: 'center' } });
    });
  };
  O.unmerge = function (sh, ranges) {
    O.tx(sh.wb, 'Unmerge Cells', () => { O.setMerges(sh, sh.merges.filter((m) => !ranges.some((rg) => M.rangesOverlap(m, rg)))); });
  };

  /* ------------------------------------------------------------ rows & columns sizes */
  O.setColWidth = function (sh, cols, w, label) {
    const wb = sh.wb;
    O.tx(wb, label || 'Column Width', () => {
      const before = cols.map((c) => [c, sh.cols[c] ? Object.assign({}, sh.cols[c]) : null]);
      for (const c of cols) { const o = sh.colObj(c); if (w === 0) o.hidden = true; else { o.w = w; o.custom = true; delete o.hidden; delete o.bestFit; } }
      const after = cols.map((c) => [c, sh.cols[c] ? Object.assign({}, sh.cols[c]) : null]);
      LY.invalidate(sh);
      wb.undo.op(() => { for (const [c, o] of before) sh.cols[c] = o ? Object.assign({}, o) : undefined; LY.invalidate(sh); }, () => { for (const [c, o] of after) sh.cols[c] = o ? Object.assign({}, o) : undefined; LY.invalidate(sh); });
    });
  };
  O.setRowHeight = function (sh, rows, pt, label) {
    const wb = sh.wb;
    O.tx(wb, label || 'Row Height', () => {
      const snap = () => rows.map((r) => [r, sh.rows[r] ? Object.assign({}, sh.rows[r]) : null]);
      const before = snap();
      for (const r of rows) { const o = sh.rowObj(r); if (pt === 0) o.hidden = true; else { o.ht = pt; o.customHeight = true; delete o.hidden; delete o.auto; } }
      const after = snap();
      LY.invalidate(sh);
      const restore = (s) => { for (const [r, o] of s) { const cells = sh.rows[r] ? sh.rows[r].cells : []; sh.rows[r] = o ? Object.assign({}, o, { cells }) : (cells.length ? { cells } : undefined); } LY.invalidate(sh); };
      wb.undo.op(() => restore(before), () => restore(after));
    });
  };
  O.hideLines = function (sh, axis, list, hide) {
    const wb = sh.wb;
    O.tx(wb, hide ? 'Hide' : 'Unhide', () => {
      const arr = axis === 'r' ? sh.rows : sh.cols;
      const snap = () => list.map((i) => [i, arr[i] ? Object.assign({}, arr[i]) : null]);
      const before = snap();
      for (const i of list) {
        const o = axis === 'r' ? sh.rowObj(i) : sh.colObj(i);
        if (hide) o.hidden = true;
        else { delete o.hidden; if (axis === 'c' && o.w === 0) o.w = sh.defColW || 8.43; if (axis === 'r' && o.ht === 0) delete o.ht; delete o.filtered; }
      }
      const after = snap();
      LY.invalidate(sh);
      const restore = (s) => { for (const [i, o] of s) { if (axis === 'r') { const cells = sh.rows[i] ? sh.rows[i].cells : []; sh.rows[i] = o ? Object.assign({}, o, { cells }) : (cells.length ? { cells } : undefined); } else sh.cols[i] = o ? Object.assign({}, o) : undefined; } LY.invalidate(sh); };
      wb.undo.op(() => restore(before), () => restore(after));
    });
  };
  /** best width (file units) for columns: widest shown text */
  O.bestWidth = function (sh, c, rows) {
    const wb = sh.wb;
    const mdw = M.mdw(wb);
    let px = 0;
    const r1 = rows ? rows.r1 : 0, r2 = rows ? rows.r2 : sh.maxR;
    sh.each(r1, c, r2, c, (cell, r) => {
      if (cell.v == null && !cell.dirty) return;
      if (LY.mergeAt(sh, r, c)) return;
      const st = LY.styleOf(sh, r, c, cell);
      if (st.align && (st.align.wrap || st.align.rot)) return;
      const f = Object.assign({}, st.font, { name: LY.fontName(wb, st.font || {}) });
      const font = LY.cssFont(f, 1);
      const d = LY.display(sh, cell, st, null, font, { noHash: true });
      if (!d) return;
      const w = d.segs ? LY.segWidth(d.segs, font) : LY.measure(font, d.text);
      const ind = st.align && st.align.indent ? st.align.indent * 9 : 0;
      px = Math.max(px, Math.ceil(w + ind + 7));
    });
    if (!px) return null;
    return Math.round(M.pxToWidth(Math.max(px, 5), mdw) * 256) / 256;
  };
  O.autofitCols = function (sh, cols, rows) {
    const wb = sh.wb;
    O.tx(wb, 'AutoFit', () => {
      for (const c of cols) { const w = O.bestWidth(sh, c, rows); if (w != null) O.setColWidth(sh, [c], w); }
    });
  };
  /** best height (pt) for a row from its tallest content */
  O.bestHeight = function (sh, r) {
    const wb = sh.wb;
    let pt = M.defaultRowPt(sh);
    const row = sh.rows[r];
    if (!row) return pt;
    const g = LY.geo(sh);
    for (const k in row.cells) {
      const c = +k, cell = row.cells[c];
      if (!cell) continue;
      const st = LY.styleOf(sh, r, c, cell);
      const f = st.font || {};
      const fpt = M.fontRowPt({ name: LY.fontName(wb, f), sz: f.sz || 11 });
      if (cell.v == null) { if (f.sz && f.sz > 11) pt = Math.max(pt, fpt); continue; }
      const m = LY.mergeAt(sh, r, c);
      if (m && m.r1 !== m.r2) continue;
      let lines = 1;
      if (st.align && st.align.wrap && typeof cell.v === 'string') {
        const w = m ? g.cols.pos(m.c2 + 1) - g.cols.pos(m.c1) : g.cols.size(c);
        const font = LY.cssFont(Object.assign({}, f, { name: LY.fontName(wb, f) }), 1);
        lines = L.render.wrapLines(font, cell.v, Math.max(4, w - 5)).length;
      } else if (typeof cell.v === 'string' && cell.v.indexOf('\n') >= 0) lines = cell.v.split('\n').length;
      if (st.align && st.align.rot && st.align.rot !== 255) {
        const font = LY.cssFont(Object.assign({}, f, { name: LY.fontName(wb, f) }), 1);
        const tw = LY.measure(font, String(cell.v));
        const a = (Math.min(90, st.align.rot > 90 ? st.align.rot - 90 : st.align.rot) * Math.PI) / 180;
        pt = Math.max(pt, (tw * Math.sin(a) + 8) * 0.75);
        continue;
      }
      pt = Math.max(pt, lines === 1 ? fpt : Math.ceil(((f.sz || 11) * 1.2 * lines + 3) * 4) / 4);
    }
    return Math.min(409, pt);
  };
  O.autofitRows = function (sh, rows, label) {
    const wb = sh.wb;
    O.tx(wb, label || 'AutoFit', () => {
      const snap = () => rows.map((r) => [r, sh.rows[r] ? Object.assign({}, sh.rows[r], { cells: undefined }) : null]);
      const before = snap();
      for (const r of rows) {
        const row = sh.rows[r];
        if (!row) continue;
        delete row.customHeight; delete row.ht;
        const pt = O.bestHeight(sh, r);
        if (Math.abs(pt - M.defaultRowPt(sh)) < 0.01) delete row.auto; else row.auto = pt;
      }
      const after = snap();
      LY.invalidate(sh);
      const restore = (s) => { for (const [r, o] of s) { const cells = sh.rows[r] ? sh.rows[r].cells : []; if (o) sh.rows[r] = Object.assign({}, o, { cells }); } LY.invalidate(sh); };
      wb.undo.op(() => restore(before), () => restore(after));
    });
  };
  /** after typing: rows without a custom height grow to fit wrapped or large text */
  O.autoRow = function (sh, r) {
    const row = sh.rows[r];
    if (!row || row.customHeight) return;
    const pt = O.bestHeight(sh, r);
    const cur = M.rowPt(sh, r);
    if (Math.abs(pt - cur) < 0.01 && row.ht == null) return;
    const old = { ht: row.ht, auto: row.auto };
    delete row.ht;
    if (Math.abs(pt - M.defaultRowPt(sh)) < 0.01) delete row.auto; else row.auto = pt;
    const now = { ht: row.ht, auto: row.auto };
    LY.invalidate(sh);
    const set = (o) => { const rr = sh.rows[r]; if (!rr) return; if (o.ht == null) delete rr.ht; else rr.ht = o.ht; if (o.auto == null) delete rr.auto; else rr.auto = o.auto; LY.invalidate(sh); };
    sh.wb.undo.op(() => set(old), () => set(now));
  };

  /* ------------------------------------------------------------ sheet snapshots (structural undo) */
  function snapSheet(sh) {
    return {
      rows: sh.rows.map((row) => (row ? Object.assign({}, row, { cells: row.cells.map((c) => (c ? Object.assign({}, c) : c)) }) : row)),
      cols: sh.cols.map((c) => (c ? Object.assign({}, c) : c)),
      colTail: sh.colTail ? JSON.parse(JSON.stringify(sh.colTail)) : null,
      merges: sh.merges.map((m) => Object.assign({}, m)),
      cf: JSON.parse(JSON.stringify(sh.cf)), dv: JSON.parse(JSON.stringify(sh.dv)), links: JSON.parse(JSON.stringify(sh.links)),
      sparklines: sh.sparklines ? sh.sparklines.map((g) => Object.assign({}, g, { items: g.items.map((i) => ({ f: i.f, sqref: i.sqref })) })) : null,
      comments: new Map(Array.from(sh.comments.entries()).map(([k, v]) => [k, Object.assign({}, v)])),
      drawings: sh.drawings.map((d) => Object.assign({}, d, { anchor: JSON.parse(JSON.stringify(d.anchor || {})) })),
      autoFilter: sh.autoFilter ? JSON.parse(JSON.stringify(sh.autoFilter)) : null,
      print: JSON.parse(JSON.stringify(sh.print)), maxR: sh.maxR, maxC: sh.maxC,
      view: { freeze: sh.view.freeze ? Object.assign({}, sh.view.freeze) : null, top: Object.assign({}, sh.view.top) },
      tableSheet: sh.id, tableOrder: sh.wb.tables.slice(),
      tables: sh.wb.tables.filter((t) => t.sheet === sh).map((t) => ({ t, ref: Object.assign({}, t.ref), columns: t.columns.map((c) => Object.assign({}, c)), filter: t.filter ? JSON.parse(JSON.stringify(t.filter)) : null, queryOps: t.queryOps?.slice() })),
      names: sh.wb.names.map((n) => Object.assign({}, n, { _ast: undefined })),
      pivotSheet: sh.id, pivots: L.pivots?.snapshot(sh.wb),
    };
  }
  function restoreSheet(sh, s) {
    sh.rows = s.rows.map((row) => (row ? Object.assign({}, row, { cells: row.cells.map((c) => (c ? Object.assign({}, c, { _gnode: null }) : c)) }) : row));
    sh.cols = s.cols.map((c) => (c ? Object.assign({}, c) : c));
    sh.colTail = s.colTail ? JSON.parse(JSON.stringify(s.colTail)) : null;
    sh.merges = s.merges.map((m) => Object.assign({}, m));
    sh.cf = JSON.parse(JSON.stringify(s.cf)); sh.dv = JSON.parse(JSON.stringify(s.dv)); sh.links = JSON.parse(JSON.stringify(s.links));
    sh.sparklines = s.sparklines ? s.sparklines.map((g) => Object.assign({}, g, { items: g.items.map((i) => ({ f: i.f, sqref: i.sqref })) })) : sh.sparklines;
    sh.comments = new Map(Array.from(s.comments.entries()).map(([k, v]) => [k, Object.assign({}, v)]));
    sh.drawings = s.drawings.map((d) => Object.assign({}, d, { anchor: JSON.parse(JSON.stringify(d.anchor || {})) }));
    sh.autoFilter = s.autoFilter ? JSON.parse(JSON.stringify(s.autoFilter)) : null;
    sh.print = JSON.parse(JSON.stringify(s.print));
    sh.maxR = s.maxR; sh.maxC = s.maxC;
    sh.view.freeze = s.view.freeze ? Object.assign({}, s.view.freeze) : null; sh.view.top = Object.assign({}, s.view.top);
    if (s.tableSheet === sh.id) {
      sh.wb.tables = s.tableOrder.slice();
      for (const x of s.tables) { x.t.ref = Object.assign({}, x.ref); x.t.columns = x.columns.map((c) => Object.assign({}, c)); x.t.filter = x.filter ? JSON.parse(JSON.stringify(x.filter)) : null; x.t.queryOps = x.queryOps?.slice(); }
    }
    sh.wb.names = s.names.map((n) => Object.assign({}, n));
    if (s.pivotSheet === sh.id) L.pivots?.restore(sh.wb, s.pivots);
    LY.invalidate(sh); LY.touchMerges(sh); sh._cfIdx = null;
  }
  O.snapSheet = snapSheet; O.restoreSheet = restoreSheet;
  /** run a structural change on one sheet with snapshot undo */
  const inStruct = new Set();
  O.structural = function (sh, label, fn) {
    const wb = sh.wb;
    /* nested structural edits of the same sheet share the outer snapshot */
    if (inStruct.has(sh)) { const r = fn(); O.markStructural(); return r; }
    return O.tx(wb, label, () => {
      const before = snapSheet(sh);
      const pivots = L.pivots?.structural(sh);
      let res;
      inStruct.add(sh);
      try { res = fn(); } finally { inStruct.delete(sh); }
      L.pivots?.afterStructural(sh, pivots);
      const after = snapSheet(sh);
      wb.undo.op(() => restoreSheet(sh, before), () => restoreSheet(sh, after));
      O.markStructural();
      return res;
    });
  };

  /* ------------------------------------------------------------ insert / delete rows & columns */
  /** move references in every formula of the workbook (op: {sheet, axis, at, n}) */
  function adjustFormulas(wb, op) {
    for (const s of wb.sheets) {
      s.rows.forEach((row, r) => {
        if (!row) return;
        row.cells.forEach((cell, c) => {
          if (!cell || cell.f == null) return;
          let ast = C.astOf(cell);
          if (!ast) return;
          const res = F.adjust(ast, op, s.name);
          if (res.changed) {
            if (s.name !== op.sheet) wb.undo.note(s, r, c);
            cell.f = F.toText(res.ast, { store: true });
            cell.ast = res.ast; cell._fsrc = cell.f; cell.dirty = true;
          }
        });
      });
    }
    for (const n of wb.names) {
      let ast; try { ast = F.parse(String(n.ref).replace(/^=/, '')); } catch (e) { continue; }
      const res = F.adjust(ast, op, n.scope != null && wb.sheets[n.scope] ? wb.sheets[n.scope].name : null);
      if (res.changed) { n.ref = F.toText(res.ast, { store: true }); n._ast = undefined; }
    }
    /* chart series and validation / CF formulas on other sheets */
    for (const s of wb.sheets) {
      for (const d of s.drawings) if (d.chart && L.xchart && L.xchart.adjust) L.xchart.adjust(d.chart, (ast) => F.adjust(ast, op, s.name));
      if (s.name === op.sheet) continue;
      const beforeCF = JSON.stringify(s.cf), beforeDV = JSON.stringify(s.dv);
      for (const v of s.dv) for (const k of ['f1', 'f2']) if (v[k]) { const t = adjText(v[k], op, s.name); if (t !== v[k]) v[k] = t; }
      for (const cf of s.cf) for (const rule of cf.rules) adjustRule(rule, op, s.name);
      const afterCF = JSON.stringify(s.cf), afterDV = JSON.stringify(s.dv);
      if (beforeCF !== afterCF || beforeDV !== afterDV) {
        const put = (cf, dv) => { s.cf = JSON.parse(cf); s.dv = JSON.parse(dv); s._cfIdx = null; };
        wb.undo.op(() => put(beforeCF, beforeDV), () => put(afterCF, afterDV));
      }
    }
    /* sparkline data ranges (the edited sheet's own groups are restored by its snapshot) */
    for (const s of wb.sheets) {
      if (!s.sparklines || !s.sparklines.length) continue;
      const before = s.sparklines.map((g) => g.items.map((i) => i.f));
      let changed = false;
      for (const g of s.sparklines) for (const it of g.items) { const t = adjText(it.f, op, s.name); if (t !== it.f) { it.f = t; changed = true; } }
      if (changed && s.name !== op.sheet) {
        const after = s.sparklines.map((g) => g.items.map((i) => i.f));
        const put = (v) => s.sparklines.forEach((g, gi) => g.items.forEach((it, ii) => { if (v[gi]) it.f = v[gi][ii]; }));
        wb.undo.op(() => put(before), () => put(after));
      }
    }
  }
  function adjText(text, op, home) {
    try { const a = F.parse(String(text).replace(/^=/, '')); const r = F.adjust(a, op, home); return r.changed ? F.toText(r.ast, { store: true }) : text; } catch (e) { return text; }
  }
  function adjustRule(rule, op, home) {
    if (rule.f) rule.f = rule.f.map(x => adjText(x, op, home));
    for (const property of ['bar', 'icons', 'scale']) for (const value of rule[property]?.cfvo || []) {
      if (value.type === 'formula' && value.val != null) value.val = adjText(value.val, op, home);
    }
  }
  /** shift a range for an insert / delete; null when it disappears */
  function shiftRange(rg, axis, at, n) {
    const k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
    const lim = axis === 'r' ? MAXR : MAXC;
    let a = rg[k1], b = rg[k2];
    if (n > 0) { if (a >= at) a += n; if (b >= at) b += n; if (a >= lim) return null; b = Math.min(b, lim - 1); }
    else {
      const cnt = -n, d0 = at, d1 = at + cnt - 1;
      if (b < d0) { /* before */ }
      else if (a > d1) { a -= cnt; b -= cnt; }
      else if (a >= d0 && b <= d1) return null;
      else { const na = a < d0 ? a : d0; const nb = b > d1 ? b - cnt : d0 - 1; a = na; b = nb; }
    }
    return Object.assign({}, rg, { [k1]: a, [k2]: b });
  }
  O.shiftRange = shiftRange;
  /**
   * Insert (n > 0) or delete (n < 0) whole rows / columns at index `at`.
   */
  O.insertLines = function (sh, axis, at, n) {
    const wb = sh.wb;
    const label = (n > 0 ? 'Insert ' : 'Delete ') + (axis === 'r' ? (Math.abs(n) > 1 ? 'Rows' : 'Row') : Math.abs(n) > 1 ? 'Columns' : 'Column');
    /* refuse to split arrays / push data off the sheet */
    if (n > 0) {
      const lastUsed = axis === 'r' ? sh.maxR : sh.maxC;
      const lim = axis === 'r' ? MAXR : MAXC;
      if (lastUsed + n >= lim && lastUsed >= at) {
        let blocked = false;
        if (axis === 'r') { for (let r = lim - n; r <= sh.maxR; r++) if (sh.rows[r] && sh.rows[r].cells.some((c) => c && (c.v != null || c.f != null))) blocked = true; }
        else sh.rows.forEach((row) => { if (row) for (let c = lim - n; c < row.cells.length; c++) if (row.cells[c] && (row.cells[c].v != null || row.cells[c].f != null)) blocked = true; });
        if (blocked) throw Object.assign(new Error('To prevent possible loss of data, Ledger cannot shift nonblank cells off of the worksheet.'), { code: 'offsheet' });
      }
    }
    for (const m of sh.merges) {
      const a = axis === 'r' ? m.r1 : m.c1, b = axis === 'r' ? m.r2 : m.c2;
      if (n < 0 && a < at && b >= at && b > at - n - 1 && false) { /* partially deleted merges shrink */ }
      void a; void b;
    }
    return O.structural(sh, label, () => {
      const op = { sheet: sh.name, axis, at, n };
      const fixedSizes = new Map(sh.drawings.filter(d => d.anchor?.editAs === 'oneCell' && d.anchor.to)
        .map(d => [d, anchorSpan(sh, d.anchor, axis)]));
      L.pivots?.lines(sh, axis, at, n);
      adjustFormulas(wb, op);
      /* move cells */
      if (axis === 'r') {
        if (n > 0) { if (sh.rows.length > at) { sh.rows.splice(at, 0, ...new Array(n)); for (let k = 0; k < n; k++) delete sh.rows[at + k]; } if (sh.rows.length > MAXR) sh.rows.length = MAXR; }
        else sh.rows.splice(at, -n);
        /* inserted rows take the format of the row above */
        if (n > 0 && at > 0) {
          const above = sh.rows[at - 1];
          if (above) for (let k = 0; k < n; k++) {
            const row = { cells: [] };
            if (above.s) row.s = above.s;
            if (above.ht != null && above.customHeight) { row.ht = above.ht; row.customHeight = true; }
            above.cells.forEach((cl, c) => { if (cl && cl.s) row.cells[c] = { v: null, s: cl.s }; });
            if (row.s || row.cells.length || row.ht) sh.rows[at + k] = row;
          }
        }
      } else {
        sh.rows.forEach((row) => {
          if (!row) return;
          if (n > 0) { if (row.cells.length > at) { row.cells.splice(at, 0, ...new Array(n)); for (let k = 0; k < n; k++) delete row.cells[at + k]; if (row.cells.length > MAXC) row.cells.length = MAXC; } }
          else if (row.cells.length > at) row.cells.splice(at, -n);
          if (n > 0 && at > 0) { const left = row.cells[at - 1]; if (left && left.s) for (let k = 0; k < n; k++) row.cells[at + k] = { v: null, s: left.s }; }
        });
        if (n > 0) { if (sh.cols.length > at) { sh.cols.splice(at, 0, ...new Array(n)); for (let k = 0; k < n; k++) delete sh.cols[at + k]; } if (at > 0 && sh.cols[at - 1]) for (let k = 0; k < n; k++) sh.cols[at + k] = Object.assign({}, sh.cols[at - 1], { hidden: undefined, level: undefined, collapsed: undefined }); }
        else if (sh.cols.length > at) sh.cols.splice(at, -n);
      }
      /* array masters move: fix master pointers */
      sh.rows.forEach((row, r) => { if (!row) return; row.cells.forEach((cl, c) => { if (cl && cl.af) { const s = shiftRange(cl.af, axis, at, n); cl.af = s || { r1: r, c1: c, r2: r, c2: c }; } }); });
      sh.rows.forEach((row) => { if (!row) return; row.cells.forEach((cl) => { if (cl && cl.am) { const p = shiftRange({ r1: cl.am.r, c1: cl.am.c, r2: cl.am.r, c2: cl.am.c }, axis, at, n); if (p) cl.am = { r: p.r1, c: p.c1 }; } }); });
      /* ranges hanging off the sheet */
      sh.merges = sh.merges.map((m) => shiftRange(m, axis, at, n)).filter((m) => m && (m.r1 !== m.r2 || m.c1 !== m.c2));
      for (const cf of sh.cf) { cf.ranges = cf.ranges.map((g) => shiftRange(g, axis, at, n)).filter(Boolean); for (const rule of cf.rules) adjustRule(rule, op, sh.name); }
      sh.cf = sh.cf.filter((cf) => cf.ranges.length);
      for (const v of sh.dv) { v.ranges = v.ranges.map((g) => shiftRange(g, axis, at, n)).filter(Boolean); for (const k of ['f1', 'f2']) if (v[k]) v[k] = adjText(v[k], op, sh.name); }
      sh.dv = sh.dv.filter((v) => v.ranges.length);
      if (sh.sparklines) {
        for (const g of sh.sparklines) g.items = g.items.map((it) => {
          const p = F.parseCell(String(it.sqref).replace(/\$/g, ''));
          if (!p) return it;
          const q = shiftRange({ r1: p.r, c1: p.c, r2: p.r, c2: p.c }, axis, at, n);
          return q ? { f: it.f, sqref: F.cellName(q.r1, q.c1) } : null;
        }).filter(Boolean);
        sh.sparklines = sh.sparklines.filter((g) => g.items.length);
      }
      sh.links = sh.links.map((l) => { const g = shiftRange(l.ref, axis, at, n); return g ? Object.assign({}, l, { ref: g }) : null; }).filter(Boolean);
      const cm = new Map();
      for (const v of sh.comments.values()) { const g = shiftRange({ r1: v.r, c1: v.c, r2: v.r, c2: v.c }, axis, at, n); if (g) cm.set(M.key(g.r1, g.c1), Object.assign({}, v, { r: g.r1, c: g.c1 })); }
      sh.comments = cm;
      if (sh.autoFilter && sh.autoFilter.ref) { const g = shiftRange(sh.autoFilter.ref, axis, at, n); sh.autoFilter = g ? Object.assign({}, sh.autoFilter, { ref: g }) : null; }
      L.tableKeep.shift(sh, axis, at, n, shiftRange);
      for (const d of sh.drawings) moveAnchor(sh, d, axis, at, n, fixedSizes.get(d));
      if (sh.print.rowBreaks && axis === 'r') sh.print.rowBreaks = sh.print.rowBreaks.map((b) => (b >= at ? b + n : b)).filter((b) => b > 0);
      if (sh.print.colBreaks && axis === 'c') sh.print.colBreaks = sh.print.colBreaks.map((b) => (b >= at ? b + n : b)).filter((b) => b > 0);
      if (axis === 'r') sh.maxR = Math.max(-1, sh.maxR + n); else sh.maxC = Math.max(-1, sh.maxC + n);
      sh.recalcBounds();
      LY.invalidate(sh); LY.touchMerges(sh);
    });
  };
  function anchorSpan(sh, a, axis) {
    let size = (a.to[axis + 'Off'] || 0) - (a.from[axis + 'Off'] || 0);
    for (let i = a.from[axis]; i < a.to[axis]; i++) size += axis === 'r' ? M.rowPt(sh, i) : M.colPx(sh, i) * 0.75;
    return Math.max(0, size);
  }
  function moveAnchor(sh, d, axis, at, n, fixedSize) {
    const a = d.anchor;
    if (!a || a.type === 'abs') return;
    const k = axis === 'r' ? 'r' : 'c';
    const mv = (p) => { if (!p) return; if (n > 0) { if (p[k] >= at) p[k] += n; } else { const cnt = -n; if (p[k] >= at + cnt) p[k] -= cnt; else if (p[k] >= at) { p[k] = at; p[k + 'Off'] = 0; } } };
    if (a.editAs === 'absolute') return;
    mv(a.from);
    if (fixedSize != null && a.to) {
      // A two-cell anchor with oneCell behavior moves but keeps its physical size.
      // Recompute its end against the rows/columns after the structural edit.
      let i = a.from[k], off = (a.from[k + 'Off'] || 0) + fixedSize;
      const limit = k === 'r' ? MAXR : MAXC;
      for (; i < limit - 1; i++) {
        const size = k === 'r' ? M.rowPt(sh, i) : M.colPx(sh, i) * 0.75;
        if (off < size) break;
        off -= size;
      }
      a.to[k] = i; a.to[k + 'Off'] = off;
    } else if (a.editAs !== 'oneCell') mv(a.to);
  }
  /** insert / delete cells inside a range, shifting neighbours */
  O.shiftCells = function (sh, rg, dir, insert) {
    const wb = sh.wb;
    const vertical = dir === 'down' || dir === 'up';
    /* implemented as a cut / paste of the block beyond the range */
    return O.structural(sh, insert ? 'Insert Cells' : 'Delete Cells', () => {
      const h = rg.r2 - rg.r1 + 1, w = rg.c2 - rg.c1 + 1;
      if (vertical) {
        const last = sh.maxR;
        const src = { r1: insert ? rg.r1 : rg.r2 + 1, c1: rg.c1, r2: last, c2: rg.c2 };
        const dr = insert ? h : -h;
        if (src.r2 >= src.r1) moveBlock(sh, src, dr, 0, insert ? null : rg);
        else clearBlock(sh, rg);
        if (insert) clearBlock(sh, rg);
        else clearBlock(sh, { r1: Math.max(rg.r1, last - h + 1), c1: rg.c1, r2: last, c2: rg.c2 });
      } else {
        const last = sh.maxC;
        const src = { r1: rg.r1, c1: insert ? rg.c1 : rg.c2 + 1, r2: rg.r2, c2: last };
        const dc = insert ? w : -w;
        if (src.c2 >= src.c1) moveBlock(sh, src, 0, dc, insert ? null : rg);
        else clearBlock(sh, rg);
        if (insert) clearBlock(sh, rg);
        else clearBlock(sh, { r1: rg.r1, c1: Math.max(rg.c1, last - w + 1), r2: rg.r2, c2: last });
      }
      sh.recalcBounds();
      LY.invalidate(sh); LY.touchMerges(sh);
    });
  };
  function clearBlock(sh, rg) { for (let r = rg.r1; r <= Math.min(rg.r2, sh.maxR); r++) { const row = sh.rows[r]; if (!row) continue; for (let c = rg.c1; c <= Math.min(rg.c2, row.cells.length - 1); c++) delete row.cells[c]; } }
  /** move a block of cells by (dr, dc) within the sheet, adjusting references to it (deleted: refs to this become #REF!) */
  function moveBlock(sh, src, dr, dc, deleted) {
    const wb = sh.wb;
    const buf = [];
    for (let r = src.r1; r <= Math.min(src.r2, sh.maxR); r++) { const row = sh.rows[r]; if (!row) continue; for (let c = src.c1; c <= Math.min(src.c2, row.cells.length - 1); c++) if (row.cells[c]) { buf.push([r, c, row.cells[c]]); delete row.cells[c]; } }
    for (const [r, c, cell] of buf) { const nr = r + dr, nc = c + dc; if (nr < 0 || nc < 0 || nr >= MAXR || nc >= MAXC) continue; const row = sh.rowObj(nr); row.cells[nc] = cell; if (nr > sh.maxR) sh.maxR = nr; if (nc > sh.maxC) sh.maxC = nc; }
    /* references into the moved block follow it; references into the deleted range become #REF! */
    const range = Object.assign({ sheet: sh.name }, src);
    for (const s of wb.sheets) s.rows.forEach((row, r) => { if (!row) return; row.cells.forEach((cell, c) => {
      if (!cell || cell.f == null) return;
      let ast = C.astOf(cell); if (!ast) return;
      let changed = false;
      if (deleted) {
        const out = F.map(ast, (n) => {
          if ((n.t !== 'ref' && n.t !== 'area') || !((n.sheet == null && s === sh) || (n.sheet && n.sheet.toLowerCase() === sh.name.toLowerCase()))) return null;
          const inDel = (rr, cc) => rr >= deleted.r1 && rr <= deleted.r2 && cc >= deleted.c1 && cc <= deleted.c2;
          if (n.t === 'ref' && inDel(n.r, n.c)) { changed = true; return { t: 'referr', sheet: n.sheet }; }
          return null;
        });
        if (changed) ast = out;
      }
      const res = F.adjustMove(ast, range, dr, dc, s.name, null);
      if (res.changed || changed) { if (s !== sh) wb.undo.note(s, r, c); cell.f = F.toText(res.ast, { store: true }); cell.ast = res.ast; cell._fsrc = cell.f; cell.dirty = true; }
    }); });
    for (const m of sh.merges) if (m.r1 >= src.r1 && m.r2 <= src.r2 && m.c1 >= src.c1 && m.c2 <= src.c2) { m.r1 += dr; m.r2 += dr; m.c1 += dc; m.c2 += dc; }
    const cm = new Map();
    for (const v of sh.comments.values()) { const inside = M.rangeContains(src, v.r, v.c); const nv = inside ? Object.assign({}, v, { r: v.r + dr, c: v.c + dc }) : v; cm.set(M.key(nv.r, nv.c), nv); }
    sh.comments = cm;
    for (const cl of buf) { const cell = cl[2]; if (cell.af) cell.af = { r1: cell.af.r1 + dr, c1: cell.af.c1 + dc, r2: cell.af.r2 + dr, c2: cell.af.c2 + dc }; if (cell.am) cell.am = { r: cell.am.r + dr, c: cell.am.c + dc }; }
  }
  O.moveBlock = moveBlock;

  /* ------------------------------------------------------------ sheets */
  O.addSheet = function (wb, name, at) {
    let sh;
    O.tx(wb, 'Insert Worksheet', () => {
      sh = wb.addSheet(name || wb.nextSheetName(), at);
      const idx = wb.sheets.indexOf(sh);
      wb.undo.op(() => { wb.sheets.splice(wb.sheets.indexOf(sh), 1); shiftScopes(wb, idx, -1); if (wb.active >= wb.sheets.length) wb.active = wb.sheets.length - 1; }, () => { wb.sheets.splice(idx, 0, sh); shiftScopes(wb, idx, 1); });
      shiftScopes(wb, idx, 1, sh);
      O.markStructural();
    });
    return sh;
  };
  /* sheet-scoped names keep pointing at the same sheet when indices move */
  function shiftScopes(wb, idx, d, except) {
    for (const n of wb.names) if (n.scope != null && n.scope >= idx && !(except && wb.sheets[n.scope] === except && d > 0 && n.scope === idx)) n.scope += d;
  }
  O.deleteSheet = function (wb, sh) {
    if (wb.sheets.filter((s) => s.state === 'visible').length <= 1 && sh.state === 'visible') throw Object.assign(new Error('A workbook must contain at least one visible worksheet.'), { code: 'lastsheet' });
    O.tx(wb, 'Delete Sheet', () => {
      const idx = wb.sheets.indexOf(sh);
      const oldNames = wb.names.map((n) => Object.assign({}, n));
      const oldTables = wb.tables.slice();
      /* formulas pointing at the sheet become #REF! */
      for (const s of wb.sheets) {
        if (s === sh) continue;
        s.rows.forEach((row, r) => { if (!row) return; row.cells.forEach((cell, c) => {
          if (!cell || cell.f == null) return;
          const ast = C.astOf(cell); if (!ast) return;
          const res = F.renameSheet(ast, sh.name, null);
          if (res.changed) { wb.undo.note(s, r, c); cell.f = F.toText(res.ast, { store: true }); cell.ast = res.ast; cell._fsrc = cell.f; cell.dirty = true; }
        }); });
      }
      wb.names = wb.names.filter((n) => n.scope !== idx).map((n) => { if (n.scope != null && n.scope > idx) n.scope--; return n; });
      wb.tables = wb.tables.filter((t) => t.sheet !== sh);
      wb.sheets.splice(idx, 1);
      const act = wb.active;
      if (wb.active >= idx && wb.active > 0) wb.active--;
      if (wb.sheets[wb.active].state !== 'visible') wb.active = wb.sheets.findIndex((s) => s.state === 'visible');
      const newNames = wb.names.slice(), newTables = wb.tables.slice(), newAct = wb.active;
      wb.undo.op(() => { wb.sheets.splice(idx, 0, sh); wb.names = oldNames; wb.tables = oldTables; wb.active = act; }, () => { wb.sheets.splice(idx, 1); wb.names = newNames; wb.tables = newTables; wb.active = newAct; });
      O.markStructural();
    });
  };
  O.validSheetName = function (wb, name, self) {
    if (!name || name.length > 31) return 'A sheet name must be 1 to 31 characters long.';
    if (/[:\\/?*[\]]/.test(name)) return 'A sheet name cannot contain any of these characters:  : \\ / ? * [ ]';
    if (/^'|'$/.test(name)) return 'A sheet name cannot begin or end with an apostrophe.';
    if (/^history$/i.test(name)) return 'History is a reserved name.';
    const other = wb.sheetByName(name);
    if (other && other !== self) return 'Cannot rename a sheet to the same name as another sheet, a referenced object library or a workbook referenced by Visual Basic.';
    return null;
  };
  O.renameSheet = function (wb, sh, name) {
    const err = O.validSheetName(wb, name, sh);
    if (err) throw Object.assign(new Error(err), { code: 'name' });
    const old = sh.name;
    if (old === name) return;
    O.tx(wb, 'Rename Sheet', () => {
      const fix = (from, to) => {
        for (const s of wb.sheets) s.rows.forEach((row, r) => { if (!row) return; row.cells.forEach((cell, c) => {
          if (!cell || cell.f == null) return;
          const ast = C.astOf(cell); if (!ast) return;
          const res = F.renameSheet(ast, from, to);
          if (res.changed) { cell.f = F.toText(res.ast, { store: true }); cell.ast = res.ast; cell._fsrc = cell.f; }
        }); });
        for (const n of wb.names) { try { const a = F.parse(String(n.ref).replace(/^=/, '')); const res = F.renameSheet(a, from, to); if (res.changed) { n.ref = F.toText(res.ast, { store: true }); n._ast = undefined; } } catch (e) { /* keep */ } }
        for (const s of wb.sheets) for (const d of s.drawings) if (d.chart && L.xchart && L.xchart.rename) L.xchart.rename(d.chart, from, to);
        for (const s of wb.sheets) for (const g of s.sparklines || []) for (const it of g.items) { try { const res = F.renameSheet(F.parse(it.f), from, to); if (res.changed) it.f = F.toText(res.ast, { store: true }); } catch (e) { /* keep */ } }
        sh.name = to;
      };
      fix(old, name);
      wb.undo.op(() => fix(name, old), () => fix(old, name));
      O.markStructural();
    });
  };
  O.moveSheet = function (wb, sh, to) {
    const from = wb.sheets.indexOf(sh);
    if (from === to || from + 1 === to) return;
    O.tx(wb, 'Move Sheet', () => {
      const order0 = wb.sheets.slice(), names0 = wb.names.map((n) => n.scope);
      const activeSheet = wb.sheets[wb.active];
      wb.sheets.splice(from, 1);
      wb.sheets.splice(to > from ? to - 1 : to, 0, sh);
      for (const n of wb.names) if (n.scope != null) n.scope = wb.sheets.indexOf(order0[n.scope]);
      wb.active = wb.sheets.indexOf(activeSheet);
      const order1 = wb.sheets.slice(), names1 = wb.names.map((n) => n.scope), a1 = wb.active;
      const a0 = order0.indexOf(activeSheet);
      wb.undo.op(() => { wb.sheets = order0.slice(); wb.names.forEach((n, i) => { n.scope = names0[i]; }); wb.active = a0; }, () => { wb.sheets = order1.slice(); wb.names.forEach((n, i) => { n.scope = names1[i]; }); wb.active = a1; });
      O.markStructural();
    });
  };
  /** copy a sheet (Move or Copy… with "Create a copy") */
  O.copySheet = function (wb, sh, to, name) {
    let copy;
    O.tx(wb, 'Copy Sheet', () => {
      const oldNames = wb.names.map(n => ({ ...n })), oldOrder = wb.sheets.slice();
      let base = sh.name.replace(/ \(\d+\)$/, ''), k = 2;
      name = name || (() => { let n; do { n = base + ' (' + k++ + ')'; } while (wb.sheetByName(n)); return n.length > 31 ? n.slice(0, 31) : n; })();
      copy = wb.addSheet(name, to);
      const snap = snapSheet(sh);
      restoreSheet(copy, snap);
      const extensions = L.sheetExtensions.copy({ cf: copy.cf, dv: copy.dv, sparklines: copy.sparklines, extensions: sh.extra.extensions });
      Object.assign(copy, { cf: extensions.cf, dv: extensions.dv, sparklines: extensions.sparklines });
      copy.comments = new Map(Array.from(copy.comments, ([key, cm]) => [key, L.threads.copy(cm, wb)]));
      copy.wb = wb;
      copy.view = JSON.parse(JSON.stringify(sh.view));
      copy.print = JSON.parse(JSON.stringify(sh.print));
      copy.tabColor = sh.tabColor; copy.defColW = sh.defColW; copy.baseColW = sh.baseColW; copy.defRowH = sh.defRowH; copy.outline = Object.assign({}, sh.outline);
      copy.fileDefRowH = sh.fileDefRowH; copy.zeroHeight = sh.zeroHeight;
      if (sh.extra.formatKeep) copy.extra.formatKeep = L.opc.duplicate(sh.extra.formatKeep);
      if (sh.extra.pageSetupKeep) copy.extra.pageSetupKeep = L.opc.duplicate(sh.extra.pageSetupKeep);
      copy.drawings = sh.drawings.map((d) => Object.assign({}, d, { anchor: JSON.parse(JSON.stringify(d.anchor)), chart: d.chart ? JSON.parse(JSON.stringify(d.chart)) : d.chart }));
      L.sheetObjects.copySheet(sh, copy);
      if (sh.extra.extensions) copy.extra.extensions = extensions.extensions;
      L.tableKeep.copySheet(sh, copy);
      wb.names = oldNames.map(n => ({ ...n, scope: n.scope == null ? n.scope : wb.sheets.indexOf(oldOrder[n.scope]) }));
      /* sheet-scoped names are duplicated for the copy */
      const idx = wb.sheets.indexOf(copy), srcIdx = wb.sheets.indexOf(sh);
      for (const n of wb.names.slice()) if (n.scope === srcIdx) {
        let ref = n.ref;
        try { const result = F.renameSheet(F.parse(String(ref).replace(/^=/, '')), sh.name, copy.name); if (result.changed) ref = F.toText(result.ast, { store: true }); } catch (_) { /* keep unparsed names */ }
        wb.names.push({ ...n, scope: idx, ref, _ast: undefined });
      }
      /* tables get new names */
      for (const t of wb.tables.filter((t) => t.sheet === sh)) {
        let tn = t.name.replace(/\d+$/, ''), i = 1; let nm; do { nm = tn + i++; } while (wb.findTable(nm));
        wb.tables.push(L.tableKeep.copy(Object.assign({}, t, { sheet: copy, name: nm, dname: nm, ooxmlCopy: true, id: Math.max(0, ...wb.tables.map(t => t.id || 0)) + 1, ref: Object.assign({}, t.ref), columns: t.columns.map((c) => Object.assign({}, c)), filter: t.filter ? JSON.parse(JSON.stringify(t.filter)) : t.filter }), wb));
      }
      const tablesAfter = wb.tables.slice();
      const namesAfter = wb.names.map(n => ({ ...n }));
      L.pivots?.copySheet(wb, sh, copy);
      wb.undo.op(() => { wb.sheets.splice(wb.sheets.indexOf(copy), 1); wb.tables = wb.tables.filter((t) => t.sheet !== copy); wb.names = oldNames.map(n => ({ ...n })); }, () => { wb.sheets.splice(idx, 0, copy); wb.tables = tablesAfter; wb.names = namesAfter.map(n => ({ ...n })); });
      O.markStructural();
    });
    return copy;
  };
  O.sheetProp = function (wb, sh, key, val, label) {
    const old = sh[key];
    O.tx(wb, label || 'Sheet', () => { sh[key] = val; wb.undo.op(() => { sh[key] = old; }, () => { sh[key] = val; }); });
  };

  /* ------------------------------------------------------------ sort */
  const typeRank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : M.isErr(v) ? 3 : 4);
  O.compareValues = function (a, b, opts) {
    opts = opts || {};
    const ea = a == null || a === '', eb = b == null || b === '';
    if (ea || eb) return ea && eb ? 0 : ea ? 1 : -1; /* blanks always last */
    const ta = typeRank(a), tb = typeRank(b);
    if (ta !== tb) return opts.desc ? tb - ta : ta - tb;
    let d = 0;
    if (ta === 0) d = a - b;
    else if (ta === 1) {
      if (opts.list) { const ia = opts.list.findIndex((x) => x.toLowerCase() === a.toLowerCase()), ib = opts.list.findIndex((x) => x.toLowerCase() === b.toLowerCase()); if (ia >= 0 || ib >= 0) d = (ia < 0 ? 1e9 : ia) - (ib < 0 ? 1e9 : ib); else d = C.strCmp(a, b); }
      else if (opts.caseSensitive) { d = C.strCmp(a, b); if (!d) d = a < b ? 1 : a > b ? -1 : 0; }
      else d = C.strCmp(a, b);
    } else if (ta === 2) d = (a ? 1 : 0) - (b ? 1 : 0);
    else d = 0;
    return opts.desc ? -d : d;
  };
  /**
   * Sort a range. keys: [{index (column / row offset within range), desc, list}], opts: {header, byRows (true = sort rows, default), caseSensitive}
   */
  O.sort = function (sh, rg, keys, opts) {
    opts = opts || {};
    const byRows = opts.byCols ? false : true;
    const a = O.cutsArray(sh, [rg]);
    if (a) throw Object.assign(new Error('You cannot change part of an array.'), { code: 'array' });
    if (sh.merges.some((m) => M.rangesOverlap(m, rg) && !(m.r1 === m.r2 && byRows) && !(m.c1 === m.c2 && !byRows))) throw Object.assign(new Error('This operation requires the merged cells to be identically sized.'), { code: 'merge' });
    return O.structural(sh, 'Sort', () => {
      const r1 = rg.r1 + (byRows && opts.header ? 1 : 0), c1 = rg.c1 + (!byRows && opts.header ? 1 : 0);
      const n = byRows ? rg.r2 - r1 + 1 : rg.c2 - c1 + 1;
      const lines = [];
      for (let i = 0; i < n; i++) {
        const cells = [];
        if (byRows) for (let c = rg.c1; c <= rg.c2; c++) cells.push(sh.get(r1 + i, c));
        else for (let r = rg.r1; r <= rg.r2; r++) cells.push(sh.get(r, c1 + i));
        const vals = keys.map((k) => { const cl = cells[k.index]; return cl ? (cl.dirty ? C.cellValue(sh, byRows ? r1 + i : rg.r1 + k.index, byRows ? rg.c1 + k.index : c1 + i) : cl.v) : null; });
        lines.push({ i, cells, vals, ht: byRows && sh.rows[r1 + i] ? { ht: sh.rows[r1 + i].ht, customHeight: sh.rows[r1 + i].customHeight, auto: sh.rows[r1 + i].auto } : null });
      }
      lines.sort((x, y) => { for (let k = 0; k < keys.length; k++) { const d = O.compareValues(x.vals[k], y.vals[k], { desc: keys[k].desc, list: keys[k].list, caseSensitive: opts.caseSensitive }); if (d) return d; } return x.i - y.i; });
      /* rewrite cells; relative references within each moved line follow it (like Excel, they are not adjusted) */
      lines.forEach((ln, k) => {
        const dst = byRows ? r1 + k : c1 + k;
        const off = k - ln.i;
        ln.cells.forEach((cl, j) => {
          const r = byRows ? dst : rg.r1 + j, c = byRows ? rg.c1 + j : dst;
          let cell = cl ? Object.assign({}, cl) : null;
          if (cell && cell.f != null && off) {
            /* references to cells in the same line move with it */
            const ast = C.astOf(cl);
            if (ast) { const sh2 = F.map(ast, (n) => { if (n.t === 'ref' && n.sheet == null && !(byRows ? n.ra : n.ca) && (byRows ? n.r === ln.i + r1 : n.c === ln.i + c1)) return Object.assign({}, n, byRows ? { r: n.r + off } : { c: n.c + off }); return null; }); cell.f = F.toText(sh2, { store: true }); cell.ast = sh2; cell._fsrc = cell.f; }
            cell.dirty = true;
          }
          if (cell) delete cell._gnode;
          if (byRows) { const row = sh.rowObj(r); if (cell) row.cells[c] = cell; else delete row.cells[c]; }
          else { const row = sh.rowObj(r); if (cell) row.cells[c] = cell; else delete row.cells[c]; }
        });
        if (byRows && ln.ht) { const row = sh.rowObj(dst); if (ln.ht.ht != null) row.ht = ln.ht.ht; else delete row.ht; if (ln.ht.customHeight) row.customHeight = true; else delete row.customHeight; if (ln.ht.auto != null) row.auto = ln.ht.auto; else delete row.auto; }
      });
      /* comments travel with their cells */
      const cm = new Map();
      const map = new Map(lines.map((ln, k) => [ln.i, k]));
      for (const v of sh.comments.values()) {
        const inside = v.r >= r1 && v.r <= rg.r2 && v.c >= c1 && v.c <= rg.c2;
        if (!inside) { cm.set(M.key(v.r, v.c), v); continue; }
        const idx = byRows ? v.r - r1 : v.c - c1;
        const k = map.get(idx);
        const nv = Object.assign({}, v, byRows ? { r: r1 + k } : { c: c1 + k });
        cm.set(M.key(nv.r, nv.c), nv);
      }
      sh.comments = cm;
      sh.sortState = { ref: Object.assign({}, rg), keys: keys.map((k) => ({ ref: byRows ? { r1, c1: rg.c1 + k.index, r2: rg.r2, c2: rg.c1 + k.index } : { r1: rg.r1 + k.index, c1, r2: rg.r1 + k.index, c2: rg.c2 }, desc: !!k.desc })), columnSort: !byRows };
      LY.invalidate(sh);
    });
  };
  /** the contiguous data region around a cell (Ctrl+* / sort / autofilter guess) */
  O.currentRegion = function (sh, r, c) {
    const has = (rr, cc) => { if (rr < 0 || cc < 0) return false; const cl = sh.get(rr, cc); return !!(cl && (cl.v != null && cl.v !== '' || cl.f != null)); };
    let r1 = r, r2 = r, c1 = c, c2 = c;
    let grew = true, guard = 0;
    while (grew && guard++ < 1000) {
      grew = false;
      const rowHas = (rr) => { for (let cc = c1 - 1; cc <= c2 + 1; cc++) if (has(rr, cc)) return true; return false; };
      const colHas = (cc) => { for (let rr = r1 - 1; rr <= r2 + 1; rr++) if (has(rr, cc)) return true; return false; };
      if (r1 > 0 && rowHas(r1 - 1)) { r1--; grew = true; }
      if (r2 < MAXR - 1 && rowHas(r2 + 1)) { r2++; grew = true; }
      if (c1 > 0 && colHas(c1 - 1)) { c1--; grew = true; }
      if (c2 < MAXC - 1 && colHas(c2 + 1)) { c2++; grew = true; }
    }
    return { r1, c1, r2, c2 };
  };
  /** does the first row of a range look like a header (text above numbers / different format)? */
  O.guessHeader = function (sh, rg) {
    if (rg.r2 <= rg.r1) return false;
    let txt = 0, cols = 0;
    for (let c = rg.c1; c <= rg.c2; c++) {
      const a = sh.val(rg.r1, c), b = sh.val(rg.r1 + 1, c);
      if (a == null) continue;
      cols++;
      if (typeof a === 'string' && (typeof b !== 'string' || b == null)) txt++;
      else if (typeof a === 'string') {
        const sa = LY.styleOf(sh, rg.r1, c, sh.get(rg.r1, c)), sb = LY.styleOf(sh, rg.r1 + 1, c, sh.get(rg.r1 + 1, c));
        if ((sa.font && sa.font.b) !== (sb.font && sb.font.b) || JSON.stringify(sa.fill) !== JSON.stringify(sb.fill)) txt++;
      }
    }
    return cols > 0 && txt / cols >= 0.5;
  };

  /* ------------------------------------------------------------ fill */
  const LISTS = [
    ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],
    ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
  ];
  O.customLists = LISTS;
  function listOf(s) {
    for (const l of O.customLists) { const i = l.findIndex((x) => x.toLowerCase() === s.toLowerCase()); if (i >= 0) return { l, i }; }
    return null;
  }
  const caseLike = (model, s) => (model === model.toUpperCase() && model !== model.toLowerCase() ? s.toUpperCase() : model === model.toLowerCase() ? s.toLowerCase() : s);
  /** linear least squares for a trend series */
  function trend(vals) {
    const n = vals.length;
    if (n === 1) return { a: vals[0], b: 1 };
    let sx = 0, sy = 0, sxx = 0, sxy = 0;
    vals.forEach((y, x) => { sx += x; sy += y; sxx += x * x; sxy += x * y; });
    const b = (n * sxy - sx * sy) / (n * sxx - sx * sx);
    const a = (sy - b * sx) / n;
    return { a, b };
  }
  /**
   * Fill a series along one line. src: source cells [{cell, r, c}] in order; returns a generator k → cell for target offsets k ≥ n (or negative for reverse)
   * mode: 'auto' (series), 'copy', 'formats', 'values', 'days', 'weekdays', 'months', 'years'
   */
  function seriesFor(sh, src, mode) {
    const n = src.length;
    const vals = src.map((s) => (s.cell ? s.cell.v : null));
    const isDate = (s) => s.cell && typeof s.cell.v === 'number' && NF.isDate(LY.styleOf(sh, s.r, s.c, s.cell).nf || 'General');
    if (mode === 'copy' || mode === 'formats') return (k) => src[((k % n) + n) % n];
    /* numbers: one value copies (unless Ctrl), several form a linear trend; dates step by days (or the detected unit) */
    const allNum = vals.every((v) => typeof v === 'number') && src.every((s) => s.cell && s.cell.f == null);
    if (allNum && n >= 1) {
      if (src.every(isDate) && (mode === 'auto' || mode === 'days' || mode === 'weekdays' || mode === 'months' || mode === 'years')) {
        /* dates: detect month / year steps */
        const p = vals.map((v) => NF.serialToParts(v, sh.wb.date1904));
        let unit = mode === 'auto' ? 'd' : mode[0];
        let step = n > 1 ? vals[1] - vals[0] : 1;
        if (mode === 'auto' && n > 1 && p.every((x) => x.d === p[0].d) && p[0].d <= 28 || (mode === 'auto' && n > 1 && p.every((x, i) => i === 0 || (x.y * 12 + x.m) - (p[i - 1].y * 12 + p[i - 1].m) === (p[1].y * 12 + p[1].m) - (p[0].y * 12 + p[0].m)) && p.every((x) => x.d === p[0].d))) {
          const dm = (p[1].y * 12 + p[1].m) - (p[0].y * 12 + p[0].m);
          if (dm && dm % 12 === 0) { unit = 'y'; step = dm / 12; } else if (dm) { unit = 'm'; step = dm; }
        }
        if (unit === 'm' && n === 1) step = 1;
        if (unit === 'y' && n === 1) step = 1;
        if (unit === 'w' && n === 1) step = 1;
        const last = p[n - 1], lastV = vals[n - 1];
        return (k) => {
          const j = k - (n - 1);
          let v;
          if (unit === 'm' || unit === 'y') {
            const months = unit === 'm' ? step * j : step * j * 12;
            const t = last.y * 12 + (last.m - 1) + months;
            const y = Math.floor(t / 12), m = t - y * 12 + 1;
            const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
            v = NF.dateToSerial(y, m, Math.min(last.d, dim), sh.wb.date1904) + (lastV - Math.floor(lastV));
          } else if (unit === 'w') {
            /* weekdays: skip Saturdays and Sundays */
            v = Math.floor(lastV); let left = Math.abs(j); const dir = j >= 0 ? 1 : -1;
            while (left > 0) { v += dir; const wd = NF.serialToParts(v, sh.wb.date1904).wd; if (wd !== 0 && wd !== 6) left--; }
            v += lastV - Math.floor(lastV);
          } else v = lastV + step * j;
          return { cell: Object.assign({}, src[((k % n) + n) % n].cell, { v }), r: null, c: null, gen: true };
        };
      }
      if (n === 1 && mode === 'auto') return (k) => src[0];
      if (n === 1) return (k) => ({ cell: Object.assign({}, src[0].cell, { v: NF.r15(vals[0] + k) }), gen: true });
      const t = trend(vals);
      return (k) => ({ cell: Object.assign({}, src[((k % n) + n) % n].cell, { v: NF.r15(t.a + t.b * k) }), gen: true });
    }
    /* text with a trailing (or leading) number: "Item 1" → "Item 2"; lists: Mon → Tue */
    const parsed = vals.map((v) => {
      if (typeof v !== 'string') return null;
      const lst = listOf(v);
      if (lst) return { kind: 'list', l: lst.l, i: lst.i, model: v };
      const m = /^(.*?)(\d+)(\D*)$/.exec(v);
      if (m && !/^\d+$/.test(v)) return { kind: 'num', pre: m[1], n: +m[2], post: m[3], width: m[2].length };
      return null;
    });
    if ((mode === 'auto' || mode === 'linear') && parsed.every((x) => x && x.kind === 'list') && new Set(parsed.map((x) => x.l)).size === 1) {
      const l = parsed[0].l;
      const step = n > 1 ? parsed[1].i - parsed[0].i : 1;
      return (k) => { const j = k - (n - 1); const i = (((parsed[n - 1].i + step * j) % l.length) + l.length) % l.length; return { cell: Object.assign({}, src[((k % n) + n) % n].cell, { v: caseLike(parsed[n - 1].model, l[i]) }), gen: true }; };
    }
    if (mode === 'auto' || mode === 'linear') {
      const pad = (num, w) => { const t = String(Math.abs(num)); return (num < 0 ? '-' : '') + (t.length < w ? t.padStart(w, '0') : t); };
      const same = parsed.every((x) => x && x.kind === 'num' && x.pre === parsed[0].pre && x.post === parsed[0].post);
      if (same && n > 1) {
        /* "Item 1", "Item 3" → step 2, continuing from the last */
        const steps = parsed.slice(1).map((x, i) => x.n - parsed[i].n);
        const step = steps.every((d) => d === steps[0]) ? steps[0] : 1;
        return (k) => { const j = k - (n - 1); const pi = parsed[n - 1]; const s0 = src[((k % n) + n) % n]; return { cell: Object.assign({}, s0.cell, { v: pi.pre + pad(pi.n + step * j, pi.width) + pi.post }), gen: true }; };
      }
      /* each text-with-number cell steps by one per repetition of the block; other cells repeat */
      return (k) => {
        const i = ((k % n) + n) % n, cyc = Math.floor(k / n);
        const s0 = src[i], pi = parsed[i];
        if (s0.cell && s0.cell.f == null && pi && pi.kind === 'num') return { cell: Object.assign({}, s0.cell, { v: pi.pre + pad(pi.n + cyc, pi.width) + pi.post }), gen: true };
        return s0;
      };
    }
    return (k) => src[((k % n) + n) % n];
  }
  /**
   * AutoFill from src range into dest range (dest contains src; fill direction by shape).
   * mode: 'auto' | 'copy' | 'formats' | 'values' | 'days' | 'weekdays' | 'months' | 'years'
   */
  O.autoFill = function (sh, src, dest, mode) {
    mode = mode || 'auto';
    const wb = sh.wb;
    const down = dest.r2 > src.r2, up = dest.r1 < src.r1, right = dest.c2 > src.c2, left = dest.c1 < src.c1;
    const vertical = down || up;
    if (O.cutsArray(sh, [dest])) throw Object.assign(new Error('You cannot change part of an array.'), { code: 'array' });
    O.tx(wb, 'AutoFill', () => {
      const lines = vertical ? src.c2 - src.c1 + 1 : src.r2 - src.r1 + 1;
      const n = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
      for (let li = 0; li < lines; li++) {
        const srcCells = [];
        for (let k = 0; k < n; k++) { const r = vertical ? src.r1 + k : src.r1 + li, c = vertical ? src.c1 + li : src.c1 + k; srcCells.push({ cell: sh.get(r, c), r, c }); }
        const gen = seriesFor(sh, srcCells, mode);
        const targets = [];
        if (down) for (let r = src.r2 + 1; r <= dest.r2; r++) targets.push({ r, c: src.c1 + li, k: r - src.r1 });
        if (up) for (let r = src.r1 - 1; r >= dest.r1; r--) targets.push({ r, c: src.c1 + li, k: r - src.r1 });
        if (right) for (let c = src.c2 + 1; c <= dest.c2; c++) targets.push({ r: src.r1 + li, c, k: c - src.c1 });
        if (left) for (let c = src.c1 - 1; c >= dest.c1; c--) targets.push({ r: src.r1 + li, c, k: c - src.c1 });
        for (const t of targets) {
          const g = gen(t.k);
          const from = g.r != null && g.r !== undefined ? g : srcCells[((t.k % n) + n) % n];
          const old = sh.get(t.r, t.c);
          let cell;
          if (mode === 'formats') { cell = dup(old); if (from.cell && from.cell.s) cell.s = from.cell.s; else delete cell.s; if (cell.v == null && cell.f == null && !cell.s) cell = null; }
          else if (!from.cell) cell = null;
          else {
            cell = Object.assign({}, g.cell || from.cell);
            delete cell._gnode; delete cell.ast; delete cell._fsrc; delete cell.am; delete cell.af; delete cell.dt;
            if (cell.f != null) {
              const sr = from.r != null ? from.r : srcCells[((t.k % n) + n) % n].r, scl = from.c != null ? from.c : srcCells[((t.k % n) + n) % n].c;
              cell.f = F.translate(cell.f, t.r - sr, t.c - scl);
              cell.edited = true;
            }
            if (mode === 'values' && old) cell.s = old.s;
          }
          O.put(sh, t.r, t.c, cell);
        }
      }
    });
  };
  /** Edit ▸ Fill ▸ Series… (linear / growth / date) */
  O.fillSeries = function (sh, rg, o) {
    const wb = sh.wb;
    const byRows = o.rows; /* series along rows (left → right) */
    O.tx(wb, 'Series', () => {
      const lines = byRows ? rg.r2 - rg.r1 + 1 : rg.c2 - rg.c1 + 1;
      const len = byRows ? rg.c2 - rg.c1 + 1 : rg.r2 - rg.r1 + 1;
      for (let li = 0; li < lines; li++) {
        const r0 = byRows ? rg.r1 + li : rg.r1, c0 = byRows ? rg.c1 : rg.c1 + li;
        const first = sh.get(r0, c0);
        let v = first && typeof first.v === 'number' ? first.v : null;
        if (v == null) continue;
        const vals = [];
        if (o.trend) {
          const known = [];
          for (let k = 0; k < len; k++) { const cl = sh.get(byRows ? r0 : r0 + k, byRows ? c0 + k : c0); if (cl && typeof cl.v === 'number') known.push(cl.v); else break; }
          const t = trend(known);
          for (let k = 0; k < len; k++) vals.push(o.type === 'growth' ? known[0] * Math.pow(known.length > 1 ? known[1] / known[0] : 1, k) : t.a + t.b * k);
        } else {
          for (let k = 0; k < len; k++) {
            if (o.stop != null && (o.step >= 0 ? v > o.stop : v < o.stop)) break;
            vals.push(v);
            if (o.type === 'growth') v *= o.step;
            else if (o.type === 'date') {
              const p = NF.serialToParts(v, wb.date1904);
              if (o.unit === 'month' || o.unit === 'year') { const t = p.y * 12 + p.m - 1 + (o.unit === 'month' ? o.step : o.step * 12); const y = Math.floor(t / 12); v = NF.dateToSerial(y, t - y * 12 + 1, Math.min(p.d, new Date(Date.UTC(y, t - y * 12 + 1, 0)).getUTCDate()), wb.date1904); }
              else if (o.unit === 'weekday') { let left = Math.abs(o.step); const dir = o.step >= 0 ? 1 : -1; while (left > 0) { v += dir; const wd = NF.serialToParts(v, wb.date1904).wd; if (wd !== 0 && wd !== 6) left--; } }
              else v += o.step;
            } else v = NF.r15(v + o.step);
          }
        }
        vals.forEach((x, k) => { if (k === 0) return; const r = byRows ? r0 : r0 + k, c = byRows ? c0 + k : c0; const cl = dup(sh.get(r, c)); delete cl.f; cl.v = NF.r15(x); if (first.s && !cl.s) cl.s = first.s; O.put(sh, r, c, cl); });
      }
    });
  };

  /* ------------------------------------------------------------ find */
  /**
   * Find cells. o: {what, within:'sheet'|'workbook', by:'rows'|'cols', lookIn:'formulas'|'values'|'comments', matchCase, whole, from:{sh,r,c}, back}
   * returns array of {sh, r, c} (all matches, in search order)
   */
  O.findAll = function (wb, o) {
    const out = [];
    const pat = String(o.what);
    let re;
    try {
      const esc = pat.replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (m, e, w, sp) => (e ? '\\' + e : w === '*' ? '[\\s\\S]*' : w === '?' ? '[\\s\\S]' : '\\' + sp));
      re = new RegExp(o.whole ? '^' + esc + '$' : esc, o.matchCase ? '' : 'i');
    } catch (e) { return out; }
    const sheets = o.within === 'workbook' ? wb.sheets : [o.sheet || wb.activeSheet];
    for (const sh of sheets) {
      const hits = [];
      if (o.lookIn === 'comments') { for (const cm of sh.comments.values()) if (re.test(cm.text || '')) hits.push({ sh, r: cm.r, c: cm.c }); }
      else sh.rows.forEach((row, r) => { if (!row) return; row.cells.forEach((cell, c) => {
        if (!cell) return;
        if (o.range && !M.rangeContains(o.range, r, c)) return;
        let text;
        if (o.lookIn === 'values' || cell.f == null) {
          const v = cell.dirty ? C.cellValue(sh, r, c) : cell.v;
          if (v == null) return;
          if (o.lookIn === 'values') { const st = LY.styleOf(sh, r, c, cell); text = NF.text(st.nf || 'General', v, { date1904: wb.date1904 }); }
          else text = typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : M.isErr(v) ? v.e : String(v);
        } else text = '=' + F.display(cell.f);
        if (re.test(text)) hits.push({ sh, r, c });
      }); });
      if (o.by === 'cols') hits.sort((a, b) => a.c - b.c || a.r - b.r);
      out.push(...hits);
    }
    return out;
  };
  /** replace text in matching cells; returns count */
  O.replaceAll = function (wb, o, cells) {
    let n = 0;
    const flags = 'g' + (o.matchCase ? '' : 'i');
    const esc = String(o.what).replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (m, e, w, sp) => (e ? '\\' + e : w === '*' ? '[\\s\\S]*' : w === '?' ? '[\\s\\S]' : '\\' + sp));
    const re = new RegExp(o.whole ? '^' + esc + '$' : esc, flags);
    O.tx(wb, 'Replace', () => {
      for (const h of cells) {
        const cell = h.sh.get(h.r, h.c);
        if (!cell) continue;
        const src = cell.f != null ? '=' + F.display(cell.f) : cell.v == null ? '' : typeof cell.v === 'boolean' ? (cell.v ? 'TRUE' : 'FALSE') : M.isErr(cell.v) ? cell.v.e : String(cell.v);
        const out = src.replace(re, () => o.with);
        if (out === src) continue;
        try { O.enter(h.sh, h.r, h.c, typeof cell.v === 'string' && cell.f == null && !/^=/.test(out) && /^[\d.,$%+-]/.test(out) && cell.v === src ? out : out, { keepFormat: true }); n++; } catch (e) { /* invalid formula: skip */ }
      }
    });
    return n;
  };
})(typeof window !== 'undefined' ? window : globalThis);
