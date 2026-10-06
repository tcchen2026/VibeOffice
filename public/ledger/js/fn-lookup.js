/* Ledger — lookup & reference functions and dynamic-array helpers. */
(function (root) {
  'use strict';
  const L = root.L;
  const C = L.calc, M = L.model, F = L.formula;
  const E = M.ERR;
  const { def, isErr, MISSING, Arr, Ref, RefList } = C;
  const opt = C.opt;
  const trunc = Math.trunc;

  /** a 1-D view over a reference / array / scalar: {n, get(i)} (whole columns stop at the used range) */
  function vec(v) {
    if (v instanceof Ref) {
      const sh = v.sheet;
      if (v.cols === 1) {
        const n = v.rows, last = Math.min(v.r2, sh.maxR) - v.r1 + 1;
        return { n, used: Math.max(0, last), get: (i) => C.cellValue(sh, v.r1 + i, v.c1), dir: 'col' };
      }
      if (v.rows === 1) {
        const n = v.cols, last = Math.min(v.c2, sh.maxC) - v.c1 + 1;
        return { n, used: Math.max(0, last), get: (i) => C.cellValue(sh, v.r1, v.c1 + i), dir: 'row' };
      }
      return null;
    }
    if (v instanceof Arr) {
      if (v.cols === 1) return { n: v.rows, used: v.rows, get: (i) => v.d[i][0], dir: 'col' };
      if (v.rows === 1) return { n: v.cols, used: v.cols, get: (i) => v.d[0][i], dir: 'row' };
      return null;
    }
    return { n: 1, used: 1, get: () => v, dir: 'col' };
  }
  C.vec = vec;
  /** comparison for approximate lookups: blanks sort after everything; types rank numbers < text < logical */
  const lkRank = (v) => (v == null ? 9 : typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 8);
  function lkCmp(v, x) {
    const a = lkRank(v), b = lkRank(x);
    if (a !== b) return a < b ? -1 : 1;
    return C.compare(v, x);
  }
  /** largest index with value <= x (ascending) or >= x (descending) */
  function bsearch(get, n, x, desc) {
    let lo = 0, hi = n - 1, found = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      const v = get(mid);
      let c = isErr(v) ? 1 : lkCmp(v, x);
      if (desc) { if (v == null) c = 1; else c = -c; }
      if (c <= 0) { found = mid; lo = mid + 1; } else hi = mid - 1;
    }
    /* the hit must be the same type as the lookup value */
    if (found >= 0) { const v = get(found); if (lkRank(v) !== lkRank(x)) return -1; }
    return found;
  }
  /** exact match predicate (case-insensitive, optional wildcards for text) */
  function exactPred(x, wild) {
    if (typeof x === 'string') {
      if (wild && /[*?~]/.test(x)) { const re = C.wildRe(x); return (v) => typeof v === 'string' && re.test(v); }
      const k = x.toLowerCase();
      return (v) => typeof v === 'string' && v.toLowerCase() === k;
    }
    if (typeof x === 'number') return (v) => typeof v === 'number' && C.numEq(v, x);
    if (typeof x === 'boolean') return (v) => v === x;
    return (v) => v == null && x == null;
  }
  C.exactPred = exactPred;
  function linear(get, n, pred, from, step) {
    if (step < 0) { for (let i = n - 1; i >= 0; i--) if (pred(get(i))) return i; return -1; }
    for (let i = from || 0; i < n; i++) if (pred(get(i))) return i;
    return -1;
  }
  const lookupVal = (x, ctx) => { x = C.scalar(x, ctx); if (x === MISSING) x = null; return x; };

  /* VLOOKUP / HLOOKUP */
  function vhlookup(ctx, x, table, idx, approx, vertical) {
    if (isErr(x)) return x;
    if (x == null) return E.NA; /* a blank lookup value finds nothing (checked before the column index) */
    if (!(table instanceof Ref) && !(table instanceof Arr)) return isErr(table) ? table : E.VALUE;
    idx = trunc(idx);
    if (isErr(idx)) return idx;
    approx = approx === MISSING ? true : C.toBool(approx);
    if (isErr(approx)) return approx;
    const len = vertical ? table.rows : table.cols;
    const width = vertical ? table.cols : table.rows;
    const sh = table instanceof Ref ? table.sheet : null;
    const get = table instanceof Ref
      ? (vertical ? (i) => C.cellValue(sh, table.r1 + i, table.c1) : (i) => C.cellValue(sh, table.r1, table.c1 + i))
      : (vertical ? (i) => table.d[i][0] : (i) => table.d[0][i]);
    const used = table instanceof Ref ? Math.max(0, vertical ? Math.min(table.r2, sh.maxR) - table.r1 + 1 : Math.min(table.c2, sh.maxC) - table.c1 + 1) : len;
    let i;
    if (approx) i = bsearch(get, used, x, false);
    else i = linear(get, used, exactPred(x, true), 0, 1);
    if (i < 0) return E.NA;
    /* the column index is checked after the lookup: a missing key gives #N/A first */
    if (idx < 1) return E.VALUE;
    if (idx > width) return E.REF;
    if (table instanceof Ref) return vertical ? C.cellValue(sh, table.r1 + i, table.c1 + idx - 1) : C.cellValue(sh, table.r1 + idx - 1, table.c1 + i);
    const v = vertical ? table.d[i][idx - 1] : table.d[idx - 1][i];
    return v == null ? 0 : v;
  }
  def('VLOOKUP', 3, 4, 'vrnv', (ctx, x, t, i, a) => { const v = vhlookup(ctx, lookupVal(x, ctx), t, i, a, true); return v == null ? 0 : v; }, { errOk: true });
  def('HLOOKUP', 3, 4, 'vrnv', (ctx, x, t, i, a) => { const v = vhlookup(ctx, lookupVal(x, ctx), t, i, a, false); return v == null ? 0 : v; }, { errOk: true });

  def('MATCH', 2, 3, 'van', (ctx, x, arr, type) => {
    x = lookupVal(x, ctx);
    if (isErr(x)) return x;
    type = opt(type, 1);
    if (isErr(type)) return type;
    const v = vec(arr);
    if (!v) return E.NA;
    if (x == null) return E.NA;
    let i;
    if (type === 0) i = linear(v.get, v.used, exactPred(x, true), 0, 1);
    else if (type > 0) i = bsearch(v.get, v.used, x, false);
    else i = bsearch(v.get, v.used, x, true);
    return i < 0 ? E.NA : i + 1;
  }, { errOk: true });
  /* the lookup array of MATCH must not be intersected: re-declare with a reference parameter */
  C.FN.MATCH.kinds = 'vrn';

  def('XMATCH', 2, 4, 'vrnn', (ctx, x, arr, mm, sm) => {
    x = lookupVal(x, ctx);
    if (isErr(x)) return x;
    const v = vec(arr);
    if (!v) return E.VALUE;
    const i = xfind(v, x, opt(mm, 0), opt(sm, 1));
    return isErr(i) ? i : i < 0 ? E.NA : i + 1;
  }, { errOk: true });
  /** XLOOKUP / XMATCH search */
  function xfind(v, x, mm, sm) {
    if (x == null) x = '';
    if (![0, -1, 1, 2].includes(mm) || ![1, -1, 2, -2].includes(sm)) return E.VALUE;
    if (sm === 2 || sm === -2) {
      const desc = sm === -2;
      if (mm === 0 || mm === 2) { const i = bsearch(v.get, v.used, x, desc); return i >= 0 && C.compare(v.get(i), x) === 0 ? i : -1; }
      let i = bsearch(v.get, v.used, x, desc);
      if (i >= 0 && C.compare(v.get(i), x) === 0) return i;
      if (mm === -1) return desc ? (i + 1 < v.used ? i + 1 : -1) : i;
      return desc ? i : i + 1 < v.used ? i + 1 : -1;
    }
    const pred = exactPred(x, mm === 2);
    const n = v.used;
    const order = sm === 1 ? (k) => k : (k) => n - 1 - k;
    let best = -1, bestV;
    for (let k = 0; k < n; k++) {
      const i = order(k);
      const val = v.get(i);
      if (pred(val)) return i;
      if (mm === -1 || mm === 1) {
        if (val == null || lkRank(val) !== lkRank(x)) continue;
        const c = C.compare(val, x);
        if (mm === -1 && c < 0 && (best < 0 || C.compare(val, bestV) > 0)) { best = i; bestV = val; }
        if (mm === 1 && c > 0 && (best < 0 || C.compare(val, bestV) < 0)) { best = i; bestV = val; }
      }
    }
    return best;
  }
  def('XLOOKUP', 3, 6, 'vrrlnn', null, { special: true, impl: (ctx, a) => {
    const actx = Object.assign({}, ctx, { array: true });
    let x = C.ev(a[0], ctx);
    const la = C.ev(a[1], ctx), ra = C.ev(a[2], ctx);
    const mm = a[4] && a[4].t !== 'missing' ? C.toNum(C.scalar(C.ev(a[4], ctx), ctx)) : 0;
    const sm = a[5] && a[5].t !== 'missing' ? C.toNum(C.scalar(C.ev(a[5], ctx), ctx)) : 1;
    if (isErr(mm)) return mm; if (isErr(sm)) return sm;
    const v = vec(la);
    if (!v) return E.VALUE;
    const one = (xv) => {
      if (isErr(xv)) return xv;
      const i = xfind(v, xv, mm, sm);
      if (isErr(i)) return i;
      if (i < 0) return a[3] && a[3].t !== 'missing' ? C.ev(a[3], ctx) : E.NA;
      /* return the matching row / column of the return array */
      if (ra instanceof Ref) {
        if (v.dir === 'col') { if (ra.rows !== v.n) return E.VALUE; return new Ref(ra.sheet, ra.r1 + i, ra.c1, ra.r1 + i, ra.c2); }
        if (ra.cols !== v.n) return E.VALUE;
        return new Ref(ra.sheet, ra.r1, ra.c1 + i, ra.r2, ra.c1 + i);
      }
      const R = C.toArr(ra);
      if (v.dir === 'col') { if (R.rows !== v.n) return E.VALUE; return R.cols === 1 ? R.d[i][0] : new Arr([R.d[i].slice()]); }
      if (R.cols !== v.n) return E.VALUE;
      return R.rows === 1 ? R.d[0][i] : new Arr(R.d.map((r) => [r[i]]));
    };
    if (x instanceof Ref && !x.single) x = ctx.array ? C.refToArr(x) : C.intersect(x, ctx);
    else x = lookupVal(x, ctx);
    if (x instanceof Arr) return new Arr(x.d.map((r) => r.map((xv) => C.scalarOut(one(xv), actx))));
    return one(x);
  } });

  def('LOOKUP', 2, 3, 'vrr', (ctx, x, arr, res) => {
    x = lookupVal(x, ctx);
    if (isErr(x)) return x;
    if (x == null) return E.NA;
    let v = vec(arr), resV;
    if (res === MISSING || res === undefined) {
      /* array form */
      const A = arr instanceof Ref ? arr : C.toArr(arr);
      const wide = A.cols > A.rows;
      const getA = (i, j) => (A instanceof Ref ? C.cellValue(A.sheet, A.r1 + i, A.c1 + j) : A.d[i][j]);
      const n = wide ? A.cols : A.rows;
      const used = A instanceof Ref ? (wide ? Math.min(A.c2, A.sheet.maxC) - A.c1 + 1 : Math.min(A.r2, A.sheet.maxR) - A.r1 + 1) : n;
      const get = wide ? (i) => getA(0, i) : (i) => getA(i, 0);
      const i = bsearch(get, Math.max(0, used), x, false);
      if (i < 0) return E.NA;
      const r = wide ? getA(A.rows - 1, i) : getA(i, A.cols - 1);
      return r == null ? 0 : r;
    }
    if (!v) return E.NA;
    const i = bsearch(v.get, v.used, x, false);
    if (i < 0) return E.NA;
    if (res instanceof Ref) {
      const r = res.cols === 1 || res.rows > 1 ? C.cellValue(res.sheet, res.r1 + i, res.c1) : C.cellValue(res.sheet, res.r1, res.c1 + i);
      return r == null ? 0 : r;
    }
    resV = vec(res);
    if (!resV) return E.NA;
    const r = i < resV.n ? resV.get(i) : E.NA;
    return r == null ? 0 : r;
  }, { errOk: true });

  /* INDEX */
  def('INDEX', 2, 4, 'rvvv', (ctx, a, r, c, area) => {
    if (isErr(a)) return a;
    let ref = a;
    if (a instanceof RefList) { const k = trunc(C.toNum(opt(area, 1))); if (isErr(k)) return k; if (k < 1 || k > a.list.length) return E.REF; ref = a.list[k - 1]; }
    else if (area !== MISSING && area !== undefined && trunc(C.toNum(area)) !== 1) return E.REF;
    const one = (rv, cv) => {
      let ri = rv === MISSING || rv == null ? 0 : C.toNum(rv), ci = cv === MISSING || cv === undefined || cv == null ? null : C.toNum(cv);
      if (isErr(ri)) return ri;
      if (isErr(ci)) return ci;
      ri = trunc(ri);
      if (ci != null) ci = trunc(ci);
      const rows = ref instanceof Ref || ref instanceof Arr ? ref.rows : 1, cols = ref instanceof Ref || ref instanceof Arr ? ref.cols : 1;
      /* a single row or column takes one index as the position along it */
      if (ci == null) { if (rows === 1 && cols > 1) { ci = ri; ri = 0; } else ci = 0; }
      if (ri < 0 || ci < 0 || ri > rows || ci > cols) return E.REF;
      if (ref instanceof Ref) {
        if (ri === 0 && ci === 0) return ref;
        if (ri === 0) return new Ref(ref.sheet, ref.r1, ref.c1 + ci - 1, ref.r2, ref.c1 + ci - 1);
        if (ci === 0) return new Ref(ref.sheet, ref.r1 + ri - 1, ref.c1, ref.r1 + ri - 1, ref.c2);
        return new Ref(ref.sheet, ref.r1 + ri - 1, ref.c1 + ci - 1);
      }
      const A = C.toArr(ref);
      if (ri === 0 && ci === 0) return A;
      if (ri === 0) return new Arr(A.d.map((row) => [row[ci - 1]]));
      if (ci === 0) return new Arr([A.d[ri - 1].slice()]);
      return A.d[ri - 1][ci - 1];
    };
    if (r instanceof Arr || c instanceof Arr) {
      return C.broadcast(r instanceof Arr ? r : r, c instanceof Arr ? c : c, (x, y) => C.scalarOut(one(x, y), ctx));
    }
    return one(r, c);
  }, { errOk: true });

  def('CHOOSE', 2, 255, 'l', null, { special: true, impl: (ctx, a) => {
    if (ctx.array && !ctx.cse) ctx = Object.assign({}, ctx, { array: false });
    let k = C.operand(C.ev(a[0], ctx), ctx);
    if (k instanceof Arr) {
      const actx = Object.assign({}, ctx, { array: true });
      const opts = a.slice(1).map((n) => C.operand(C.ev(n, actx), actx));
      return new Arr(k.d.map((row, i) => row.map((x, j) => {
        const n = trunc(C.toNum(x)); if (isErr(n)) return n;
        if (n < 1 || n > opts.length) return E.VALUE;
        const o = opts[n - 1];
        return o instanceof Arr ? (o.d[Math.min(i, o.rows - 1)] || [])[Math.min(j, o.cols - 1)] : o;
      })));
    }
    k = C.toNum(k);
    if (isErr(k)) return k;
    k = trunc(k);
    if (k < 1 || k >= a.length) return E.VALUE;
    const n = a[k];
    return n.t === 'missing' ? 0 : C.ev(n, ctx);
  } });

  def('OFFSET', 3, 5, 'rnnnn', (ctx, ref, dr, dc, hh, ww) => {
    if (!(ref instanceof Ref)) return isErr(ref) ? ref : E.VALUE;
    dr = trunc(dr); dc = trunc(dc);
    const h = trunc(opt(hh, ref.rows)), w = trunc(opt(ww, ref.cols));
    if (h === 0 || w === 0) return E.REF;
    let r1 = ref.r1 + dr, c1 = ref.c1 + dc;
    let r2 = r1 + Math.abs(h) - 1, c2 = c1 + Math.abs(w) - 1;
    if (h < 0) { r2 = r1; r1 = r1 + h + 1; }
    if (w < 0) { c2 = c1; c1 = c1 + w + 1; }
    if (r1 < 0 || c1 < 0 || r2 >= M.MAXR || c2 >= M.MAXC) return E.REF;
    return new Ref(ref.sheet, r1, c1, r2, c2);
  }, { volatile: true });

  def('INDIRECT', 1, 2, 'sb', (ctx, text, a1) => {
    a1 = opt(a1, true);
    let t = String(text).trim();
    if (!t) return E.REF;
    if (!a1) t = r1c1ToA1(t, ctx);
    if (t == null) return E.REF;
    let ast;
    try { ast = F.parse(t); } catch (e) { return E.REF; }
    if (!['ref', 'area', 'name', 'struct'].includes(ast.t) && !(ast.t === 'op' && ast.op === ':')) return E.REF;
    const v = C.ev(ast, ctx);
    return v instanceof Ref || v instanceof RefList ? v : E.REF;
  }, { volatile: true });
  function r1c1ToA1(t, ctx) {
    const m = /^(?:((?:'[^']+'|[^!]+))!)?(.*)$/.exec(t);
    const sheet = m[1] ? m[1] + '!' : '';
    const part = (s) => {
      const k = /^R(\[?-?\d*\]?)?C(\[?-?\d*\]?)?$/i.exec(s);
      if (!k) return null;
      const comp = (x, base, isRow) => {
        if (!x) return { v: base, abs: false };
        if (x[0] === '[') return { v: base + parseInt(x.slice(1, -1), 10), abs: false };
        return { v: parseInt(x, 10) - 1, abs: true };
      };
      const r = comp(k[1], ctx.r, true), c = comp(k[2], ctx.c, false);
      if (r.v < 0 || c.v < 0) return null;
      return (c.abs ? '$' : '') + F.colName(c.v) + (r.abs ? '$' : '') + (r.v + 1);
    };
    const bits = m[2].split(':').map(part);
    if (bits.some((b) => b == null)) return null;
    return sheet + bits.join(':');
  }

  def('ROW', 0, 1, 'r', (ctx, ref) => {
    if (ref === undefined || ref === MISSING) return ctx.r + 1;
    if (!(ref instanceof Ref)) return isErr(ref) ? ref : E.VALUE;
    if (ref.rows === 1) return ref.r1 + 1;
    return new Arr(Array.from({ length: ref.rows }, (_, i) => [ref.r1 + i + 1]));
  });
  def('COLUMN', 0, 1, 'r', (ctx, ref) => {
    if (ref === undefined || ref === MISSING) return ctx.c + 1;
    if (!(ref instanceof Ref)) return isErr(ref) ? ref : E.VALUE;
    if (ref.cols === 1) return ref.c1 + 1;
    return new Arr([Array.from({ length: ref.cols }, (_, j) => ref.c1 + j + 1)]);
  });
  def('ROWS', 1, 1, 'r', (ctx, a) => (a instanceof Ref ? a.rows : a instanceof Arr ? a.vrows : isErr(a) ? a : a instanceof RefList ? E.REF : 1), { errOk: true });
  def('COLUMNS', 1, 1, 'r', (ctx, a) => (a instanceof Ref || a instanceof Arr ? a.cols : isErr(a) ? a : a instanceof RefList ? E.REF : 1), { errOk: true });
  def('AREAS', 1, 1, 'r', (ctx, a) => (a instanceof RefList ? a.list.length : a instanceof Ref ? 1 : isErr(a) ? a : E.VALUE), { errOk: true });
  def('ADDRESS', 2, 5, 'nnnbs', (ctx, r, c, abs, a1, sheet) => {
    r = trunc(r); c = trunc(c); abs = trunc(opt(abs, 1)); a1 = opt(a1, true);
    if (r < 1 || c < 1 || r > M.MAXR || c > M.MAXC || abs < 1 || abs > 4) return E.VALUE;
    let s;
    if (a1) s = (abs === 1 || abs === 3 ? '$' : '') + F.colName(c - 1) + (abs === 1 || abs === 2 ? '$' : '') + r;
    else s = 'R' + (abs === 1 || abs === 2 ? r : '[' + r + ']') + 'C' + (abs === 1 || abs === 3 ? c : '[' + c + ']');
    if (sheet !== MISSING && sheet !== undefined) s = (sheet === '' ? '' : (/^\[[^\]]+\][A-Za-z_][\w.]*$/.test(sheet) ? sheet : F.quoteSheet(sheet)) + '!') + s;
    return s;
  });
  def('TRANSPOSE', 1, 1, 'a', (ctx, a) => {
    const A = C.toArr(a);
    if (!A.rows) return E.VALUE;
    return new Arr(A.d[0].map((_, j) => A.d.map((row) => (row[j] == null ? 0 : row[j]))));
  });
  def('HYPERLINK', 1, 2, 'vv', (ctx, link, name) => (name === MISSING || name === undefined ? link : name == null ? 0 : name));
  def('FORMULATEXT', 1, 1, 'r', (ctx, ref) => {
    if (!(ref instanceof Ref)) return isErr(ref) ? ref : E.NA;
    let c = ref.sheet.get(ref.r1, ref.c1);
    if (c && c.am) c = ref.sheet.get(c.am.r, c.am.c);
    if (!c || c.f == null) return E.NA;
    const t = '=' + F.display(c.f);
    return c.af && !c.dyn ? '{' + t + '}' : t;
  }, { errOk: true });
  /* A1# — the range a dynamic array formula spills into */
  def('ANCHORARRAY', 1, 1, 'r', (ctx, ref) => {
    if (!(ref instanceof Ref)) return isErr(ref) ? ref : E.REF;
    if (!ref.single) return E.REF;
    const c = ref.sheet.get(ref.r1, ref.c1);
    if (!c || !c.af) return E.REF;
    if (c.dirty) C.evalCell(ref.sheet, ref.r1, ref.c1, c);
    const a = c.af;
    return new Ref(ref.sheet, a.r1, a.c1, a.r2, a.c2);
  }, { errOk: true });
  def('SINGLE', 1, 1, 'r', (ctx, v) => (v instanceof Ref || v instanceof Arr ? C.scalar(v, ctx) : v), { errOk: true });
  /* functions that need data we do not have keep the value saved in the file */
  const KEEP = (C.KEEP = M.err('#KEEP'));
  for (const n of ['GETPIVOTDATA', 'CUBEMEMBER', 'CUBEVALUE', 'CUBESET', 'CUBESETCOUNT', 'CUBERANKEDMEMBER', 'CUBEMEMBERPROPERTY', 'CUBEKPIMEMBER', 'WEBSERVICE', 'FILTERXML', 'RTD', 'STOCKHISTORY', 'CALL', 'REGISTER.ID', 'EUROCONVERT', 'SQL.REQUEST', 'IMAGE', 'TRANSLATE', 'DETECTLANGUAGE', 'COPILOT', 'PY'])
    def(n, 0, 255, 'r', (ctx) => { ctx.flags.keep = true; return KEEP; }, { list: true });

  /* ------------------------------------------------------------ dynamic arrays */
  const arr = (v) => C.toArr(v);
  def('FILTER', 2, 3, 'aav', (ctx, a, inc, empty) => {
    const A = arr(a), I = arr(inc);
    let keepRows = null, keepCols = null;
    if (I.cols === 1 && I.rows === A.rows) keepRows = I.d.map((r) => r[0]);
    else if (I.rows === 1 && I.cols === A.cols) keepCols = I.d[0];
    else return E.VALUE;
    const truthy = (v) => { if (isErr(v)) return v; const b = typeof v === 'string' ? E.VALUE : C.toBool(v); return b; };
    let out;
    if (keepRows) {
      out = [];
      for (let i = 0; i < A.rows; i++) { const t = truthy(keepRows[i]); if (isErr(t)) return t; if (t) out.push(A.d[i]); }
    } else {
      const idx = [];
      for (let j = 0; j < A.cols; j++) { const t = truthy(keepCols[j]); if (isErr(t)) return t; if (t) idx.push(j); }
      out = idx.length ? A.d.map((r) => idx.map((j) => r[j])) : [];
    }
    if (!out.length || !out[0].length) return empty === MISSING || empty === undefined ? E.CALC : empty;
    return new Arr(out);
  });
  const sortKey = (a, b) => {
    const ra = a == null ? 4 : typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : typeof a === 'boolean' ? 2 : 3;
    const rb = b == null ? 4 : typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : typeof b === 'boolean' ? 2 : 3;
    if (ra !== rb) return ra - rb;
    if (ra === 0) return a - b;
    if (ra === 1) return C.strCmp(a, b);
    if (ra === 2) return (a ? 1 : 0) - (b ? 1 : 0);
    return 0;
  };
  C.sortKey = sortKey;
  def('SORT', 1, 4, 'annb', (ctx, a, idx, ord, byCol) => {
    const A = arr(a);
    idx = trunc(opt(idx, 1)); ord = opt(ord, 1); byCol = opt(byCol, false);
    if (ord !== 1 && ord !== -1) return E.VALUE;
    if (!byCol) {
      if (idx < 1 || idx > A.cols) return E.VALUE;
      const rows = A.d.map((r, i) => ({ r, i }));
      rows.sort((x, y) => sortKey(x.r[idx - 1], y.r[idx - 1]) * ord || x.i - y.i);
      return new Arr(rows.map((x) => x.r));
    }
    if (idx < 1 || idx > A.rows) return E.VALUE;
    const cols = A.d[0].map((_, j) => j);
    cols.sort((x, y) => sortKey(A.d[idx - 1][x], A.d[idx - 1][y]) * ord || x - y);
    return new Arr(A.d.map((r) => cols.map((j) => r[j])));
  });
  def('SORTBY', 2, 254, 'a', (ctx, args) => {
    const A = arr(args[0]);
    const keys = [];
    for (let i = 1; i < args.length; i += 2) {
      const K = arr(args[i]);
      const o = i + 1 < args.length ? C.toNum(C.scalar(args[i + 1], ctx)) : 1;
      if (o !== 1 && o !== -1) return E.VALUE;
      if (K.cols === 1 && K.rows === A.rows) keys.push({ v: K.d.map((r) => r[0]), o, rows: true });
      else if (K.rows === 1 && K.cols === A.cols) keys.push({ v: K.d[0], o, rows: false });
      else return E.VALUE;
    }
    const byRows = keys[0].rows;
    if (keys.some((k) => k.rows !== byRows)) return E.VALUE;
    const n = byRows ? A.rows : A.cols;
    const order = [...Array(n).keys()];
    order.sort((x, y) => { for (const k of keys) { const c = sortKey(k.v[x], k.v[y]) * k.o; if (c) return c; } return x - y; });
    return byRows ? new Arr(order.map((i) => A.d[i])) : new Arr(A.d.map((r) => order.map((j) => r[j])));
  }, { list: true });
  const keyOf = (v) => (typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + ':' + String(v));
  def('UNIQUE', 1, 3, 'abb', (ctx, a, byCol, once) => {
    const A = arr(a);
    byCol = opt(byCol, false); once = opt(once, false);
    const items = byCol ? A.d[0].map((_, j) => A.d.map((r) => r[j])) : A.d;
    const counts = new Map();
    for (const it of items) { const k = it.map(keyOf).join('\u0001'); counts.set(k, (counts.get(k) || 0) + 1); }
    const seen = new Set();
    const out = [];
    for (const it of items) {
      const k = it.map(keyOf).join('\u0001');
      if (once ? counts.get(k) !== 1 : seen.has(k)) continue;
      seen.add(k);
      out.push(it);
    }
    if (!out.length) return E.CALC;
    return byCol ? new Arr(A.d.map((_, i) => out.map((col) => col[i]))) : new Arr(out);
  });
  def('TAKE', 2, 3, 'ann', (ctx, a, r, c) => {
    const A = arr(a);
    r = r === MISSING ? A.rows : trunc(r); c = c === MISSING || c === undefined ? A.cols : trunc(c);
    if (r === 0 || c === 0) return E.CALC;
    const rows = r > 0 ? A.d.slice(0, r) : A.d.slice(Math.max(0, A.rows + r));
    return new Arr(rows.map((row) => (c > 0 ? row.slice(0, c) : row.slice(Math.max(0, A.cols + c)))));
  });
  def('DROP', 2, 3, 'ann', (ctx, a, r, c) => {
    const A = arr(a);
    r = r === MISSING ? 0 : trunc(r); c = c === MISSING || c === undefined ? 0 : trunc(c);
    const rows = r >= 0 ? A.d.slice(r) : A.d.slice(0, A.rows + r);
    const out = rows.map((row) => (c >= 0 ? row.slice(c) : row.slice(0, A.cols + c)));
    if (!out.length || !out[0].length) return E.CALC;
    return new Arr(out);
  });
  def('CHOOSEROWS', 2, 255, 'an', (ctx, a, ...idx) => {
    const A = arr(a);
    const out = [];
    for (let k of idx) { k = trunc(k); if (k === 0 || Math.abs(k) > A.rows) return E.VALUE; out.push(A.d[k > 0 ? k - 1 : A.rows + k]); }
    return new Arr(out);
  });
  def('CHOOSECOLS', 2, 255, 'an', (ctx, a, ...idx) => {
    const A = arr(a);
    const js = [];
    for (let k of idx) { k = trunc(k); if (k === 0 || Math.abs(k) > A.cols) return E.VALUE; js.push(k > 0 ? k - 1 : A.cols + k); }
    return new Arr(A.d.map((r) => js.map((j) => r[j])));
  });
  def('VSTACK', 1, 254, 'a', (ctx, args) => {
    const As = args.map(arr);
    const w = Math.max(...As.map((x) => x.cols));
    const out = [];
    for (const A of As) for (const r of A.d) { const row = r.slice(); while (row.length < w) row.push(E.NA); out.push(row); }
    return new Arr(out);
  }, { list: true });
  def('HSTACK', 1, 254, 'a', (ctx, args) => {
    const As = args.map(arr);
    const h = Math.max(...As.map((x) => x.rows));
    const out = Array.from({ length: h }, () => []);
    for (const A of As) for (let i = 0; i < h; i++) { const r = A.d[i]; for (let j = 0; j < A.cols; j++) out[i].push(r ? r[j] : E.NA); }
    return new Arr(out);
  }, { list: true });
  const flat = (A, ignore, byCol) => {
    const out = [];
    const take = (v) => { if (ignore === 1 || ignore === 3) { if (v == null) return; } if (ignore === 2 || ignore === 3) { if (isErr(v)) return; } out.push(v); };
    if (byCol) { for (let j = 0; j < A.cols; j++) for (let i = 0; i < A.rows; i++) take(A.d[i][j]); }
    else for (const r of A.d) for (const v of r) take(v);
    return out;
  };
  def('TOCOL', 1, 3, 'anb', (ctx, a, ig, byCol) => { const o = flat(arr(a), trunc(opt(ig, 0)), opt(byCol, false)); return o.length ? new Arr(o.map((v) => [v])) : E.CALC; });
  def('TOROW', 1, 3, 'anb', (ctx, a, ig, byCol) => { const o = flat(arr(a), trunc(opt(ig, 0)), opt(byCol, false)); return o.length ? new Arr([o]) : E.CALC; });
  const wrap = (byRows) => (ctx, a, n, pad) => {
    const A = arr(a);
    if (A.rows !== 1 && A.cols !== 1) return E.VALUE;
    n = trunc(n);
    if (n < 1) return E.NUM;
    pad = pad === MISSING || pad === undefined ? E.NA : pad;
    const v = A.d.flat();
    const k = Math.ceil(v.length / n);
    const out = [];
    if (byRows) for (let i = 0; i < k; i++) { const row = v.slice(i * n, i * n + n); while (row.length < n) row.push(pad); out.push(row); }
    else { for (let i = 0; i < n; i++) out.push([]); for (let j = 0; j < k; j++) for (let i = 0; i < n; i++) { const x = v[j * n + i]; out[i].push(x === undefined ? pad : x); } }
    return new Arr(out);
  };
  def('WRAPROWS', 2, 3, 'anv', wrap(true));
  def('WRAPCOLS', 2, 3, 'anv', wrap(false));
  def('EXPAND', 2, 4, 'annv', (ctx, a, r, c, pad) => {
    const A = arr(a);
    r = r === MISSING ? A.rows : trunc(r); c = c === MISSING || c === undefined ? A.cols : trunc(c);
    if (r < A.rows || c < A.cols) return E.VALUE;
    pad = pad === MISSING || pad === undefined ? E.NA : pad;
    return new Arr(Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => (i < A.rows && j < A.cols ? A.d[i][j] : pad))));
  });
  /* LAMBDA helpers */
  const lam = (v) => (v && v.lambda ? v : null);
  def('MAP', 2, 254, 'a', (ctx, args) => {
    const f = lam(args[args.length - 1]); if (!f) return E.VALUE;
    const As = args.slice(0, -1).map(arr);
    const A = As[0];
    return new Arr(A.d.map((r, i) => r.map((_, j) => C.scalarOut(C.callLambda(f, As.map((X) => (X.d[i] ? X.d[i][j] : E.NA)), ctx), ctx))));
  }, { list: true });
  def('REDUCE', 3, 3, 'vaa', (ctx, init, a, f) => { f = lam(f); if (!f) return E.VALUE; let acc = init === MISSING ? 0 : init; for (const v of arr(a).d.flat()) acc = C.scalarOut(C.callLambda(f, [acc, v], ctx), ctx); return acc; });
  def('SCAN', 3, 3, 'vaa', (ctx, init, a, f) => { f = lam(f); if (!f) return E.VALUE; let acc = init === MISSING ? 0 : init; const A = arr(a); return new Arr(A.d.map((r) => r.map((v) => (acc = C.scalarOut(C.callLambda(f, [acc, v], ctx), ctx))))); });
  def('BYROW', 2, 2, 'aa', (ctx, a, f) => { f = lam(f); if (!f) return E.VALUE; const A = arr(a); return new Arr(A.d.map((r) => [C.scalarOut(C.callLambda(f, [new Arr([r])], ctx), ctx)])); });
  def('BYCOL', 2, 2, 'aa', (ctx, a, f) => { f = lam(f); if (!f) return E.VALUE; const A = arr(a); return new Arr([A.d[0].map((_, j) => C.scalarOut(C.callLambda(f, [new Arr(A.d.map((r) => [r[j]]))], ctx), ctx))]); });
  def('MAKEARRAY', 3, 3, 'nna', (ctx, r, c, f) => { f = lam(f); if (!f) return E.VALUE; r = trunc(r); c = trunc(c); if (r < 1 || c < 1) return E.VALUE; return new Arr(Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => C.scalarOut(C.callLambda(f, [i + 1, j + 1], ctx), ctx)))); });
})(typeof window !== 'undefined' ? window : globalThis);
