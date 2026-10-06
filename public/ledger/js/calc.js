/* Ledger — calculation engine.
 * Values and coercion, references and arrays, the AST evaluator (implicit intersection, array formulas,
 * broadcasting), the function registry, criteria matching, and dependency-driven recalculation.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const M = L.model, F = L.formula, NF = L.numfmt;
  const C = (L.calc = {});
  const E = M.ERR;
  const isErr = (v) => v instanceof M.XErr;
  C.isErr = isErr;

  /* ------------------------------------------------------------ objects */
  /** a rectangular reference on one sheet */
  class Ref {
    constructor(sheet, r1, c1, r2, c2) { this.sheet = sheet; this.r1 = r1; this.c1 = c1; this.r2 = r2 == null ? r1 : r2; this.c2 = c2 == null ? c1 : c2; }
    get rows() { return this.r2 - this.r1 + 1; }
    get cols() { return this.c2 - this.c1 + 1; }
    get single() { return this.r1 === this.r2 && this.c1 === this.c2; }
  }
  /** several areas (union operator) */
  class RefList { constructor(list) { this.list = list; } }
  /** a 2-D array of scalar values */
  /* Arrays built from tall references keep only the used rows: vr is the full (virtual) row count and
     tail the value of every row past the stored ones (blank for references, f(blank) after arithmetic). */
  class Arr {
    constructor(d) { this.d = d; }
    get rows() { return this.d.length; }
    get vrows() { return this.vr || this.d.length; }
    get cols() { return this.d[0] ? this.d[0].length : 0; }
    get(i, j) { const row = this.d[i]; return row ? row[j] : undefined; }
    static fill(r, c, v) { return new Arr(Array.from({ length: r }, () => new Array(c).fill(v))); }
  }
  const MISSING = (C.MISSING = { missing: true });
  /** 3-D reference across sheets */
  class Ref3D { constructor(list) { this.list = list; } }
  C.Ref = Ref; C.RefList = RefList; C.Arr = Arr; C.Ref3D = Ref3D;

  /* ------------------------------------------------------------ helpers */
  const collator = new Intl.Collator('en-US', { sensitivity: 'accent', numeric: false });
  C.strCmp = (a, b) => { const x = collator.compare(a, b); return x < 0 ? -1 : x > 0 ? 1 : 0; };
  C.strEq = (a, b) => a.length === b.length ? a.toLowerCase() === b.toLowerCase() || collator.compare(a, b) === 0 && a.toLowerCase() === b.toLowerCase() : false;
  /** number → text the way Excel converts values in formulas (15 significant digits) */
  C.numToStr = function (x) {
    if (!isFinite(x)) return '#NUM!';
    if (x === 0) return '0';
    const a = Math.abs(x);
    let s;
    if (a >= 1e-10 && a < 1e15) {
      s = parseFloat(x.toPrecision(15)).toString();
      if (/e/.test(s)) {
        const p = parseFloat(x.toPrecision(15));
        s = p.toFixed(Math.min(20, Math.max(0, 14 - Math.floor(Math.log10(Math.abs(p)))))).replace(/\.?0+$/, '');
      }
      return s;
    }
    s = parseFloat(x.toPrecision(15)).toExponential();
    const m = /^(-?[\d.]+)e([+-])(\d+)$/.exec(s);
    return m[1] + 'E' + m[2] + (m[3].length < 2 ? '0' + m[3] : m[3]);
  };
  /** text → number for arithmetic ("1,234", "$5", "50%", dates, times); null when not numeric */
  C.parseNum = function (s, wb) {
    const t = String(s).trim();
    if (t === '') return null;
    const r = NF.parseInput(t, { date1904: wb && wb.date1904 });
    if (typeof r.v === 'number') return r.v;
    return null;
  };
  C.toNum = function (v, ctx) {
    if (typeof v === 'number') return v;
    if (v == null || v === MISSING) return 0;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') { const n = C.parseNum(v, ctx && ctx.wb); return n == null ? E.VALUE : n; }
    if (isErr(v)) return v;
    return E.VALUE;
  };
  C.toStr = function (v) {
    if (typeof v === 'string') return v;
    if (v == null || v === MISSING) return '';
    if (typeof v === 'number') return C.numToStr(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (isErr(v)) return v;
    return '';
  };
  C.toBool = function (v) {
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (v == null || v === MISSING) return false;
    if (typeof v === 'string') { const u = v.toUpperCase(); if (u === 'TRUE') return true; if (u === 'FALSE') return false; return E.VALUE; }
    if (isErr(v)) return v;
    return E.VALUE;
  };
  /** Excel's 15-digit tolerant equality for numbers */
  C.numEq = (a, b) => a === b || (Math.abs(a - b) <= Math.max(Math.abs(a), Math.abs(b)) * 1e-15);

  /* --------------------------------------------------------- dereference */
  /** value at a cell, recalculating it first when needed */
  C.cellValue = function (sh, r, c) {
    const row = sh.rows[r];
    if (!row) return null;
    const cell = row.cells[c];
    if (!cell) return null;
    if (cell.dirty) C.evalCell(sh, r, c, cell);
    else if (cell.am) { const m = sh.get(cell.am.r, cell.am.c); if (m && m.dirty) C.evalCell(sh, cell.am.r, cell.am.c, m); }
    return cell.v === undefined ? null : cell.v;
  };
  /** Ref → Arr of values */
  C.refToArr = function (ref) {
    let rows = ref.rows;
    const cols = ref.cols;
    const sh = ref.sheet;
    const rMax = Math.min(ref.r2, Math.max(sh.maxR, -1));
    let vr = 0;
    const used = Math.max(1, rMax - ref.r1 + 1);
    if (rows > 8192 && rows - used > 4096) { vr = rows; rows = used; }
    const d = new Array(rows);
    for (let i = 0; i < rows; i++) {
      const r = ref.r1 + i;
      const row = new Array(cols).fill(null);
      if (r <= rMax && sh.rows[r]) for (let j = 0; j < cols; j++) row[j] = C.cellValue(sh, r, ref.c1 + j);
      d[i] = row;
    }
    const a = new Arr(d);
    if (vr) { a.vr = vr; a.tail = null; }
    return a;
  };
  /** an array result with the virtual size of its inputs */
  const withTail = (out, vr, tail) => { if (vr) { out.vr = vr; out.tail = tail; } return out; };
  C.withTail = withTail;
  /** implicit intersection of a multi-cell reference with the formula cell */
  C.intersect = function (ref, ctx) {
    if (ref.single) return C.cellValue(ref.sheet, ref.r1, ref.c1);
    if (ref.c1 === ref.c2) { if (ctx.r >= ref.r1 && ctx.r <= ref.r2) return C.cellValue(ref.sheet, ctx.r, ref.c1); return E.VALUE; }
    if (ref.r1 === ref.r2) { if (ctx.c >= ref.c1 && ctx.c <= ref.c2) return C.cellValue(ref.sheet, ref.r1, ctx.c); return E.VALUE; }
    if (ctx.r >= ref.r1 && ctx.r <= ref.r2 && ctx.c >= ref.c1 && ctx.c <= ref.c2) return C.cellValue(ref.sheet, ctx.r, ctx.c);
    return E.VALUE;
  };
  /** any evaluation result → a scalar for a scalar position */
  C.scalar = function (v, ctx) {
    if (v instanceof Ref) return v.single ? C.cellValue(v.sheet, v.r1, v.c1) : C.intersect(v, ctx);
    if (v instanceof Arr) return v.rows ? v.d[0][0] : E.VALUE;
    if (v instanceof RefList || v instanceof Ref3D) return E.VALUE;
    return v;
  };
  /** any result → Arr (arrays and references), scalars become 1×1 */
  C.toArr = function (v) {
    if (v instanceof Arr) return v;
    if (v instanceof Ref) return C.refToArr(v);
    if (v instanceof RefList) return v.list.length === 1 ? C.refToArr(v.list[0]) : new Arr([[E.VALUE]]);
    return new Arr([[v === MISSING ? null : v]]);
  };

  /* ------------------------------------------------------------ iteration */
  /**
   * Visit every value in the arguments of an aggregate function.
   * fn(value, direct) — direct is true for values typed straight into the argument list.
   * Blank cells are skipped. Returns an error if visitor returns one (stops).
   */
  C.each = function (args, fn, opts) {
    for (const a of args) {
      if (a === MISSING) { if (opts && opts.missingAsZero) { const r = fn(0, true); if (r !== undefined) return r; } continue; }
      if (a instanceof Ref) { const r = C.eachRef(a, fn, opts); if (r !== undefined) return r; continue; }
      if (a instanceof RefList) { for (const x of a.list) { const r = C.eachRef(x, fn, opts); if (r !== undefined) return r; } continue; }
      if (a instanceof Ref3D) { for (const x of a.list) { const r = C.eachRef(x, fn, opts); if (r !== undefined) return r; } continue; }
      if (a instanceof Arr) {
        for (const row of a.d) for (const v of row) { if (v == null) continue; const r = fn(v, false); if (r !== undefined) return r; }
        continue;
      }
      const r = fn(a, true);
      if (r !== undefined) return r;
    }
    return undefined;
  };
  C.eachRef = function (ref, fn, opts) {
    const sh = ref.sheet;
    let out;
    const skipHidden = opts && opts.skipHidden; /* SUBTOTAL 101-111 */
    const skipSub = opts && opts.skipSubtotals;
    sh.each(ref.r1, ref.c1, ref.r2, ref.c2, (cell, r, c) => {
      if (skipHidden && sh.rows[r] && sh.rows[r].hidden) return;
      if (skipSub && cell.f && /(^|[^A-Za-z0-9_.])(?:_xlfn\.)?(SUBTOTAL|AGGREGATE)\s*\(/i.test(cell.f)) return;
      const v = cell.dirty || (cell.am && cell.am.dirty) ? C.cellValue(sh, r, c) : cell.v;
      if (v == null) return;
      const res = fn(v, false, r, c);
      if (res !== undefined) { out = res; return false; }
    });
    return out;
  };
  /** collect numbers like SUM does: direct text→number conversion, booleans counted when direct */
  C.nums = function (args, ctx, opts) {
    const out = [];
    const all = opts && opts.all; /* the *A functions: text counts as 0, booleans count */
    const err = C.each(args, (v, direct, r) => {
      if (typeof v === 'number') { out.push(v); return; }
      if (isErr(v)) return v;
      /* *A functions count logical values in references, not in array constants */
      if (typeof v === 'boolean') { if (direct || (all && r !== undefined)) out.push(v ? 1 : 0); return; }
      if (typeof v === 'string') {
        if (direct) { const n = C.parseNum(v, ctx.wb); if (n == null) return E.VALUE; out.push(n); return; }
        if (all) out.push(0);
      }
    }, opts);
    return err || out;
  };

  /* ------------------------------------------------------------ criteria */
  const wildRe = (s) => new RegExp('^' + s.replace(/~([*?~])|([*?])|([.+^${}()|[\]\\])/g, (m, esc, wild, sp) => (esc ? '\\' + esc : wild === '*' ? '[\\s\\S]*' : wild === '?' ? '[\\s\\S]' : '\\' + sp)) + '$', 'i');
  /** compile a COUNTIF-style criterion into a predicate(value) */
  C.criteria = function (crit, ctx) {
    if (crit instanceof Ref) crit = C.scalar(crit, ctx);
    if (crit == null) crit = '';
    if (typeof crit === 'number') return (v) => (typeof v === 'number' ? C.numEq(v, crit) : typeof v === 'string' ? (() => { const n = C.parseNum(v, ctx.wb); return n != null && C.numEq(n, crit); })() : false);
    if (typeof crit === 'boolean') return (v) => v === crit;
    if (isErr(crit)) return (v) => v === crit;
    const s = String(crit);
    const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(s);
    const op = m[1] || '=';
    const rhs = m[2];
    if (rhs === '') {
      if (op === '=') return (v) => v == null || v === '';
      if (op === '<>') return (v) => !(v == null || v === '');
      return () => false;
    }
    const num = C.parseNum(rhs, ctx.wb);
    const upper = rhs.toUpperCase();
    if (upper === 'TRUE' || upper === 'FALSE') {
      const b = upper === 'TRUE';
      if (op === '=') return (v) => v === b;
      if (op === '<>') return (v) => v !== b;
    }
    if (rhs[0] === '#') { const e = M.ERR[upper]; if (e) { if (op === '=') return (v) => v === e; if (op === '<>') return (v) => v !== e; } }
    if (num != null) {
      const cmp = (n) => {
        switch (op) {
          case '=': return C.numEq(n, num); case '<>': return !C.numEq(n, num);
          case '<': return n < num && !C.numEq(n, num); case '>': return n > num && !C.numEq(n, num);
          case '<=': return n <= num || C.numEq(n, num); case '>=': return n >= num || C.numEq(n, num);
          default: return false;
        }
      };
      return (v) => {
        if (typeof v === 'number') return cmp(v);
        if (op === '<>') return true;
        if (typeof v === 'string' && op === '=') { const n = C.parseNum(v, ctx.wb); return n != null && cmp(n); }
        return false;
      };
    }
    if (op === '=' || op === '<>') {
      const wild = /[*?~]/.test(rhs);
      const re = wild ? wildRe(rhs) : null;
      const low = rhs.toLowerCase();
      const eq = (v) => { if (typeof v !== 'string') { if (v == null) return false; v = C.toStr(v); if (typeof v !== 'string') return false; } return re ? re.test(v) : v.toLowerCase() === low; };
      return op === '=' ? eq : (v) => !eq(v);
    }
    return (v) => {
      if (typeof v !== 'string') return false;
      const k = C.strCmp(v, rhs);
      return op === '<' ? k < 0 : op === '>' ? k > 0 : op === '<=' ? k <= 0 : k >= 0;
    };
  };
  C.wildRe = wildRe;

  /* ------------------------------------------------------------ registry */
  /**
   * C.def(name, min, max, kinds, impl, flags)
   *  kinds: one letter per parameter ('v' value, 'n' number, 's' text, 'b' logical, 'r' raw reference/array,
   *         'a' array-evaluated raw, 'l' lazy AST) — the last letter repeats for extra arguments.
   *  impl(ctx, ...args) or impl(ctx, args) when flags.list
   *  flags: {volatile, list, refs (returns references), lift: false to disable element-wise lifting}
   */
  const FN = (C.FN = Object.create(null));
  C.def = function (name, min, max, kinds, impl, flags) {
    FN[name] = Object.assign({ name, min, max, kinds, impl }, flags || {});
  };
  C.alias = (name, target) => { FN[name] = Object.assign({}, FN[target], { name }); };
  const kindAt = (def, i) => def.kinds[Math.min(i, def.kinds.length - 1)] || 'v';

  /* ------------------------------------------------------------ evaluator */
  C.ctx = (wb, sh, r, c, array) => ({ wb, sh, r, c, array: !!array, cse: !!array, depth: 0, vars: null, flags: {} });

  /** evaluate an AST node; result may be a scalar, Ref, RefList, Arr */
  function ev(n, ctx) {
    switch (n.t) {
      case 'num': return n.v;
      case 'str': return n.v;
      case 'bool': return n.v;
      case 'err': return M.err(n.v);
      case 'missing': return MISSING;
      case 'ref': {
        const sh = n.sheet != null || n.book != null ? sheetOf(n, ctx) : ctx.sh;
        if (!sh || isErr(sh)) return sh || E.REF;
        if (n.sheet2) return ref3d(n, ctx, (s) => new Ref(s, n.r, n.c));
        return new Ref(sh, n.r, n.c);
      }
      case 'area': {
        const sh = n.sheet != null || n.book != null ? sheetOf(n, ctx) : ctx.sh;
        if (!sh || isErr(sh)) return sh || E.REF;
        const r1 = Math.min(n.r1, n.r2), r2 = Math.max(n.r1, n.r2), c1 = Math.min(n.c1, n.c2), c2 = Math.max(n.c1, n.c2);
        if (n.sheet2) return ref3d(n, ctx, (s) => new Ref(s, r1, c1, r2, c2));
        return new Ref(sh, r1, c1, r2, c2);
      }
      case 'referr': return E.REF;
      case 'name': return evName(n, ctx);
      case 'struct': return evStruct(n, ctx);
      case 'paren': return ev(n.a, ctx);
      case 'union': {
        const list = [];
        for (const it of n.items) {
          const v = ev(it, ctx);
          if (v instanceof Ref) list.push(v);
          else if (v instanceof RefList) list.push(...v.list);
          else return isErr(v) ? v : E.VALUE;
        }
        return new RefList(list);
      }
      case 'arr': return new Arr(n.rows.map((row) => row.map((x) => (x.t === 'err' ? M.err(x.v) : x.t === 'missing' ? null : x.v))));
      case 'fn': return call(n, ctx);
      case 'invoke': {
        const lam = ev(n.fn, ctx);
        if (isErr(lam)) return lam;
        if (!lam || !lam.lambda) return E.VALUE;
        return callLambda(lam, n.args, ctx);
      }
      case 'un': {
        if (n.op === '@') { const v = ev(n.a, ctx); return v instanceof Ref || v instanceof Arr ? C.scalar(v, ctx) : v; }
        const v = operand(ev(n.a, ctx), ctx);
        return lift1(v, (x) => { const k = C.toNum(x, ctx); if (isErr(k)) return k; return n.op === '-' ? -k : k; });
      }
      case 'pct': {
        const v = operand(ev(n.a, ctx), ctx);
        return lift1(v, (x) => { const k = C.toNum(x, ctx); return isErr(k) ? k : k / 100; });
      }
      case 'op': return binop(n, ctx);
      default: return E.VALUE;
    }
  }
  C.ev = ev;
  /** an external workbook by index ([1]) or by name / path */
  C.extBook = function (wb, key) {
    if (!wb.external) return null;
    key = String(key).replace(/^'|'$/g, '');
    if (wb.external[key]) return wb.external[key];
    const f = decodeURIComponent(key.split(/[\\/]/).pop() || '').toLowerCase();
    for (const k in wb.external) if (k.toLowerCase() === f) return wb.external[k];
    return null;
  };
  function sheetOf(n, ctx) {
    if (n.book != null) {
      const ext = C.extBook(ctx.wb, n.book);
      if (!ext) return E.REF;
      const sh = ext.sheets && ext.sheets[String(n.sheet).toLowerCase()];
      return sh || E.REF;
    }
    return ctx.wb.sheetByName(n.sheet) || E.REF;
  }
  function ref3d(n, ctx, mk) {
    const a = ctx.wb.sheetIndex(ctx.wb.sheetByName(n.sheet)), b = ctx.wb.sheetIndex(ctx.wb.sheetByName(n.sheet2));
    if (a < 0 || b < 0) return E.REF;
    const list = [];
    for (let i = Math.min(a, b); i <= Math.max(a, b); i++) list.push(mk(ctx.wb.sheets[i]));
    return new Ref3D(list);
  }
  /** operands of operators: references → scalar (intersection) or array (array context) */
  function operand(v, ctx) {
    if (v instanceof Ref) { if (v.single) return C.cellValue(v.sheet, v.r1, v.c1); return ctx.array ? C.refToArr(v) : C.intersect(v, ctx); }
    if (v instanceof RefList || v instanceof Ref3D) return E.VALUE;
    if (v === MISSING) return null;
    if (v instanceof Arr && !ctx.array && false) return v.d[0][0];
    return v;
  }
  C.operand = operand;
  function lift1(v, f) {
    if (v instanceof Arr) return withTail(new Arr(v.d.map((row) => row.map((x) => (isErr(x) ? x : f(x))))), v.vr, v.vr ? (isErr(v.tail) ? v.tail : f(v.tail)) : undefined);
    if (isErr(v)) return v;
    return f(v);
  }
  /** broadcast two arrays the way Excel does (1-row / 1-column stretch, #N/A outside) */
  function broadcast(a, b, f) {
    const A = a instanceof Arr ? a : null, B = b instanceof Arr ? b : null;
    if (!A && !B) return f(a, b);
    const ar = A ? A.rows : 1, ac = A ? A.cols : 1, br = B ? B.rows : 1, bc = B ? B.cols : 1;
    const R = Math.max(ar, br), Cc = Math.max(ac, bc);
    const get = (X, xr, xc, s, i, j) => {
      if (!X) return s;
      const ii = xr === 1 && !X.vr ? 0 : i, jj = xc === 1 ? 0 : j;
      if (jj >= xc) return E.NA;
      if (ii >= xr) return X.vr && ii < X.vr ? X.tail : E.NA;
      return X.d[ii][jj];
    };
    const d = new Array(R);
    for (let i = 0; i < R; i++) {
      const row = new Array(Cc);
      for (let j = 0; j < Cc; j++) row[j] = f(get(A, ar, ac, a, i, j), get(B, br, bc, b, i, j));
      d[i] = row;
    }
    const out = new Arr(d);
    /* virtual rows continue when every array operand has them (or stretches as a single row) */
    const va = A ? (A.vr || (ar === 1 ? -1 : 0)) : -1, vb = B ? (B.vr || (br === 1 ? -1 : 0)) : -1;
    if ((A && A.vr) || (B && B.vr)) {
      if (va !== 0 && vb !== 0) {
        const vr = Math.max(va, vb);
        const ta = !A ? a : A.vr ? A.tail : A.d[0][0], tb = !B ? b : B.vr ? B.tail : B.d[0][0];
        if (Cc === 1) withTail(out, vr, f(ta, tb));
      }
    }
    return out;
  }
  C.broadcast = broadcast;
  /** three-way broadcast (IF over arrays) */
  C.broadcast3 = function (a, b, c, f) {
    const xs = [a, b, c];
    let R = 1, Cc = 1;
    for (const x of xs) if (x instanceof Arr) { R = Math.max(R, x.rows); Cc = Math.max(Cc, x.cols); }
    const at = (x, i, j) => {
      if (!(x instanceof Arr)) return x;
      const ii = x.rows === 1 ? 0 : i, jj = x.cols === 1 ? 0 : j;
      if (ii >= x.rows || jj >= x.cols) return E.NA;
      return x.d[ii][jj];
    };
    const d = [];
    for (let i = 0; i < R; i++) { const row = []; for (let j = 0; j < Cc; j++) row.push(f(at(a, i, j), at(b, i, j), at(c, i, j))); d.push(row); }
    return new Arr(d);
  };
  const typeRank = (v) => (typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 3 : 0);
  /** Excel comparison: numbers < text < logical; text case-insensitive; blanks take the other side's type */
  C.compare = function (a, b) {
    if (a == null || a === MISSING) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
    if (b == null || b === MISSING) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
    const ta = typeRank(a), tb = typeRank(b);
    if (ta !== tb) return ta < tb ? -1 : 1;
    if (ta === 1) return C.numEq(a, b) ? 0 : a < b ? -1 : 1;
    if (ta === 2) { if (a.toLowerCase() === b.toLowerCase()) return 0; return C.strCmp(a, b) || (a.toLowerCase() < b.toLowerCase() ? -1 : 1); }
    if (ta === 3) return a === b ? 0 : a ? 1 : -1;
    return 0;
  };
  function binop(n, ctx) {
    const op = n.op;
    if (op === ':') {
      const a = ev(n.a, ctx), b = ev(n.b, ctx);
      if (isErr(a)) return a;
      if (isErr(b)) return b;
      /* the range operator spans the bounding box of every area on both sides: (A1,B2):C3 → A1:C3 */
      const list = [].concat(a instanceof RefList ? a.list : [a], b instanceof RefList ? b.list : [b]);
      if (!list.every((x) => x instanceof Ref) || list.some((x) => x.sheet !== list[0].sheet)) return E.VALUE;
      return new Ref(list[0].sheet, Math.min(...list.map((x) => x.r1)), Math.min(...list.map((x) => x.c1)), Math.max(...list.map((x) => x.r2)), Math.max(...list.map((x) => x.c2)));
    }
    if (op === ' ') {
      const a = ev(n.a, ctx), b = ev(n.b, ctx);
      if (isErr(a)) return a;
      if (isErr(b)) return b;
      const la = a instanceof RefList ? a.list : [a], lb = b instanceof RefList ? b.list : [b];
      const out = [];
      for (const x of la) for (const y of lb) {
        if (!(x instanceof Ref) || !(y instanceof Ref)) return E.VALUE;
        if (x.sheet !== y.sheet) continue;
        const r1 = Math.max(x.r1, y.r1), r2 = Math.min(x.r2, y.r2), c1 = Math.max(x.c1, y.c1), c2 = Math.min(x.c2, y.c2);
        if (r1 <= r2 && c1 <= c2) out.push(new Ref(x.sheet, r1, c1, r2, c2));
      }
      if (!out.length) return E.NULL;
      return out.length === 1 ? out[0] : new RefList(out);
    }
    const a = operand(ev(n.a, ctx), ctx);
    const b = operand(ev(n.b, ctx), ctx);
    const fn = OPS[op];
    if (!fn) return E.VALUE;
    return broadcast(a, b, (x, y) => fn(x, y, ctx));
  }
  const arith = (f) => (x, y, ctx) => {
    if (isErr(x)) return x;
    if (isErr(y)) return y;
    const a = C.toNum(x, ctx); if (isErr(a)) return a;
    const b = C.toNum(y, ctx); if (isErr(b)) return b;
    const r = f(a, b);
    if (typeof r === 'number' && !isFinite(r)) return isNaN(r) ? E.NUM : E.NUM;
    return r;
  };
  const cmpOp = (f) => (x, y) => { if (isErr(x)) return x; if (isErr(y)) return y; return f(C.compare(x, y)); };
  const OPS = {
    '+': arith((a, b) => a + b),
    '-': arith((a, b) => a - b),
    '*': arith((a, b) => a * b),
    '/': arith((a, b) => (b === 0 ? E.DIV0 : a / b)),
    '^': arith((a, b) => {
      if (a === 0 && b === 0) return E.NUM;
      if (a === 0 && b < 0) return E.DIV0;
      const r = Math.pow(a, b);
      if (isNaN(r)) return E.NUM;
      return r;
    }),
    '&': (x, y) => { if (isErr(x)) return x; if (isErr(y)) return y; return C.toStr(x) + C.toStr(y); },
    '=': cmpOp((k) => k === 0), '<>': cmpOp((k) => k !== 0), '<': cmpOp((k) => k < 0), '>': cmpOp((k) => k > 0), '<=': cmpOp((k) => k <= 0), '>=': cmpOp((k) => k >= 0),
  };
  C.OPS = OPS;

  /* names: workbook / sheet scope, LET variables */
  function evName(n, ctx) {
    const key = n.name.toUpperCase();
    if (ctx.vars && n.sheet == null) {
      for (let s = ctx.vars; s; s = s.parent) if (s.map.has(key)) return s.map.get(key);
    }
    let sheetIdx = ctx.wb.sheetIndex(ctx.sh);
    if (n.sheet != null) { const sh = ctx.wb.sheetByName(n.sheet); if (!sh) return E.REF; sheetIdx = ctx.wb.sheetIndex(sh); }
    if (n.book != null) {
      /* a name defined in another workbook: [1]!Name, using the values cached with this file */
      const ext = C.extBook(ctx.wb, n.book);
      const d = ext && ext.names[n.name.toLowerCase()];
      if (!d) return E.REF;
      let a;
      try { a = F.parse(d.ref); } catch (e) { return E.REF; }
      F.walk(a, (x) => { if ((x.t === 'ref' || x.t === 'area' || x.t === 'name') && x.book == null) x.book = n.book; });
      return ev(a, ctx);
    }
    const def = ctx.wb.findName(n.name, sheetIdx);
    if (!def) {
      const tbl = ctx.wb.findTable(n.name);
      if (tbl) return new Ref(tbl.sheet, tbl.ref.r1 + (tbl.header === false ? 0 : 1), tbl.ref.c1, tbl.ref.r2 - (tbl.totals ? 1 : 0), tbl.ref.c2);
      return E.NAME;
    }
    if (ctx.depth > 60) return E.REF;
    let ast = def._ast;
    if (ast === undefined || def._src !== def.ref) {
      try { ast = F.parse(String(def.ref).replace(/^=/, '')); } catch (e) { ast = null; }
      def._ast = ast; def._src = def.ref;
    }
    if (!ast) return E.NAME;
    /* names with relative references are relative to the cell that uses them */
    const home = def.scope != null ? ctx.wb.sheets[def.scope] : ctx.sh;
    const sub = Object.assign({}, ctx, { depth: ctx.depth + 1, sh: home || ctx.sh, vars: null });
    /* relative names are stored relative to A1 and wrap around the sheet edges ($G1048576 = the row above) */
    const shifted = hasRelative(ast) ? F.shift(ast, ctx.r, ctx.c, true) : ast;
    const v = ev(shifted, sub);
    /* a relative name evaluated without a sheet qualifier refers to the using sheet */
    if (v instanceof Ref && def.scope == null && !refHasSheet(ast) && home !== ctx.sh) return new Ref(ctx.sh, v.r1, v.c1, v.r2, v.c2);
    return v;
  }
  const hasRelative = (ast) => { let rel = false; F.walk(ast, (x) => { if ((x.t === 'ref' && (!x.ra || !x.ca)) || (x.t === 'area' && x.kind !== 'cols' && x.kind !== 'rows' && (!x.ra1 || !x.ca1 || !x.ra2 || !x.ca2))) rel = true; }); return rel; };
  const refHasSheet = (ast) => { let s = false; F.walk(ast, (x) => { if ((x.t === 'ref' || x.t === 'area') && x.sheet != null) s = true; }); return s; };

  /* structured references: Table[Col], Table[[#Headers],[A]:[B]], [@Col] */
  function evStruct(n, ctx) {
    let tbl = n.table ? ctx.wb.findTable(n.table) : ctx.wb.tableAt(ctx.sh, ctx.r, ctx.c);
    if (!tbl) return n.table ? E.REF : E.REF;
    const rg = C.structRange(tbl, n.spec, ctx.r);
    if (!rg) return E.REF;
    if (isErr(rg)) return rg;
    return new Ref(tbl.sheet, rg.r1, rg.c1, rg.r2, rg.c2);
  }
  /** resolve a structured reference spec "[...]" against a table → {r1,c1,r2,c2} */
  C.structRange = function (tbl, spec, row) {
    const s = spec.slice(1, -1);
    const t0 = s.trim();
    const unesc = (x) => x.replace(/'(.)/g, '$1');
    /* items: {name} column, {special} '#…' item, ':' column range */
    const items = [];
    if (!t0.startsWith('[')) {
      if (t0[0] === '@') { items.push({ special: '#this row' }); const rest = t0.slice(1).trim(); if (rest) items.push({ name: unesc(rest.replace(/^\[([\s\S]*)\]$/, '$1')) }); }
      else if (t0[0] === '#') items.push({ special: t0.toLowerCase() });
      else if (t0) items.push({ name: unesc(s) });
    } else {
      let depth = 0, cur = '', esc0 = false;
      for (let i = 0; i < s.length; i++) {
        const ch = s[i];
        if (depth > 0 && ch === "'") { if (!cur.length) esc0 = true; cur += s[i + 1] || ''; i++; continue; }
        if (ch === '[') { if (depth++ === 0) { cur = ''; esc0 = false; continue; } }
        if (ch === ']') {
          if (--depth === 0) {
            if (cur[0] === '@' && !esc0) { items.push({ special: '#this row' }); const rest = cur.slice(1).trim().replace(/^\[([\s\S]*)\]$/, '$1'); if (rest) items.push({ name: rest }); }
            else if (cur[0] === '#' && !esc0) items.push({ special: cur.trim().toLowerCase() });
            else items.push({ name: cur });
            cur = '';
            continue;
          }
        }
        if (depth > 0) cur += ch;
        else if (ch === ':') items.push(':');
      }
    }
    const parts = items;
    const hdr = tbl.header === false ? 0 : 1, tot = tbl.totals ? 1 : 0;
    const ref = tbl.ref;
    let r1 = ref.r1 + hdr, r2 = ref.r2 - tot;
    const specials = [], cols = [];
    for (const it of parts) {
      if (it === ':') continue;
      if (it.special) specials.push(it.special.replace(/\s+/g, ' '));
      else cols.push(it.name);
    }
    for (const sp of specials) {
      if (sp === '#all') { r1 = ref.r1; r2 = ref.r2; }
      else if (sp === '#headers') { if (!hdr) return M.ERR.REF; r1 = r2 = ref.r1; }
      else if (sp === '#totals') { if (!tot) return M.ERR.REF; r1 = r2 = ref.r2; }
      else if (sp === '#data') { /* default */ }
      else if (sp === '#this row') { if (row < ref.r1 + hdr || row > ref.r2 - tot) return M.ERR.VALUE; r1 = r2 = row; }
    }
    if (specials.includes('#headers') && specials.includes('#data')) { r1 = ref.r1; r2 = ref.r2 - tot; }
    if (specials.includes('#data') && specials.includes('#totals')) { r1 = ref.r1 + hdr; r2 = ref.r2; }
    let c1 = ref.c1, c2 = ref.c2;
    if (cols.length) {
      const idx = (name) => {
        const k = String(name).toLowerCase();
        let i = tbl.columns.findIndex((c) => String(c.name).toLowerCase() === k);
        if (i < 0) { const kt = k.trim().replace(/\s+/g, ' '); i = tbl.columns.findIndex((c) => String(c.name).toLowerCase().trim().replace(/\s+/g, ' ') === kt); }
        return i;
      };
      const a = idx(cols[0]), b = idx(cols[cols.length - 1]);
      if (a < 0 || b < 0) return M.ERR.REF;
      c1 = ref.c1 + Math.min(a, b); c2 = ref.c1 + Math.max(a, b);
    }
    return { r1, c1, r2, c2 };
  };

  /* function calls */
  function call(n, ctx) {
    const def = n.udf ? null : FN[n.name];
    if (!def) {
      /* LAMBDA held in a LET variable or a defined name */
      const lam = lookupLambda(n.name, ctx);
      if (lam) return callLambda(lam, n.args, ctx);
      /* an add-in or newer function we do not know: keep the result saved with the file */
      ctx.flags.keep = true;
      return E.NAME;
    }
    const argc = n.args.length;
    if (argc < def.min || argc > def.max) return E.VALUE;
    if (def.special) return def.impl(ctx, n.args, n);
    const args = new Array(argc);
    let liftAt = null;
    for (let i = 0; i < argc; i++) {
      const k = kindAt(def, i);
      const node = n.args[i];
      if (k === 'l') { args[i] = node; continue; }
      if (k === 'r' || k === 'a') {
        const sub = k === 'a' && !ctx.array ? Object.assign({}, ctx, { array: true }) : ctx;
        args[i] = node.t === 'missing' ? MISSING : ev(node, sub);
        continue;
      }
      let v = node.t === 'missing' ? MISSING : ev(node, ctx);
      /* scalar parameter: intersect references; arrays lift the function element-wise */
      if (v instanceof Ref) v = v.single ? C.cellValue(v.sheet, v.r1, v.c1) : ctx.array ? C.refToArr(v) : C.intersect(v, ctx);
      else if (v instanceof RefList || v instanceof Ref3D) v = E.VALUE;
      if (v instanceof Arr) { if (def.lift === false) v = v.d[0][0]; else (liftAt || (liftAt = [])).push(i); }
      args[i] = v;
    }
    if (liftAt) return liftCall(def, args, liftAt, ctx);
    return finish(def, args, ctx);
  }
  function convert(k, v, ctx) {
    if (v === MISSING) return MISSING;
    switch (k) {
      case 'n': return isErr(v) ? v : C.toNum(v, ctx);
      case 's': return isErr(v) ? v : C.toStr(v);
      case 'b': return isErr(v) ? v : C.toBool(v);
      default: return v;
    }
  }
  function finish(def, args, ctx) {
    if (!def.list && args.length < def.max && def.max <= 32) while (args.length < def.max) args.push(MISSING);
    for (let i = 0; i < args.length; i++) {
      const k = kindAt(def, i);
      if (k === 'n' || k === 's' || k === 'b') {
        const v = convert(k, args[i], ctx);
        if (isErr(v) && !def.errOk) return v;
        args[i] = v;
      } else if (k === 'v' && isErr(args[i]) && !def.errOk) return args[i];
    }
    const r = def.list ? def.impl(ctx, args) : def.impl(ctx, ...args);
    if (typeof r === 'number' && !isFinite(r)) return E.NUM;
    return r === undefined ? E.VALUE : r;
  }
  function liftCall(def, args, at, ctx) {
    let R = 1, Cc = 1;
    for (const i of at) { R = Math.max(R, args[i].rows); Cc = Math.max(Cc, args[i].cols); }
    const d = new Array(R);
    for (let i = 0; i < R; i++) {
      const row = new Array(Cc);
      for (let j = 0; j < Cc; j++) {
        const a = args.slice();
        let bad = null;
        for (const k of at) {
          const A = args[k];
          const ii = A.rows === 1 ? 0 : i, jj = A.cols === 1 ? 0 : j;
          if (ii >= A.rows || jj >= A.cols) { bad = E.NA; break; }
          a[k] = A.d[ii][jj];
        }
        row[j] = bad || C.scalarOut(finish(def, a, ctx), ctx);
      }
      d[i] = row;
    }
    const out = new Arr(d);
    const vrs = at.map((k) => args[k].vr || 0);
    if (Cc === 1 && vrs.some(Boolean) && at.every((k) => args[k].vr || args[k].rows === 1)) {
      const a = args.slice();
      for (const k of at) a[k] = args[k].vr ? args[k].tail : args[k].d[0][0];
      withTail(out, Math.max(...vrs), C.scalarOut(finish(def, a, ctx), ctx));
    }
    return out;
  }
  /** a function result used as an element must be a scalar */
  C.scalarOut = (v, ctx) => (v instanceof Ref || v instanceof Arr ? C.scalar(v, ctx) : v);

  function lookupLambda(name, ctx) {
    const key = name.toUpperCase();
    if (ctx.vars) for (let s = ctx.vars; s; s = s.parent) if (s.map.has(key)) { const v = s.map.get(key); if (v && v.lambda) return v; }
    const def = ctx.wb.findName(name, ctx.wb.sheetIndex(ctx.sh));
    if (def) { const v = evName({ t: 'name', name }, ctx); if (v && v.lambda) return v; }
    return null;
  }
  function callLambda(lam, argNodes, ctx) {
    if (argNodes.length > lam.params.length) return E.VALUE;
    const map = new Map();
    lam.params.forEach((p, i) => { map.set(p, i < argNodes.length ? ev(argNodes[i], ctx) : MISSING); });
    const sub = Object.assign({}, ctx, { vars: { map, parent: lam.scope } });
    return ev(lam.body, sub);
  }
  C.callLambda = (lam, values, ctx) => {
    const map = new Map();
    lam.params.forEach((p, i) => map.set(p, i < values.length ? values[i] : MISSING));
    return ev(lam.body, Object.assign({}, ctx, { vars: { map, parent: lam.scope } }));
  };

  /* ------------------------------------------------------------ cells */
  C.circular = new Set();
  let evalStack = 0;
  /** parse a cell's formula (cached on the cell) */
  C.astOf = function (cell) {
    if (cell.ast !== undefined && cell._fsrc === cell.f) return cell.ast;
    let ast = null;
    try { ast = F.parse(cell.f); } catch (e) { ast = null; }
    cell.ast = ast; cell._fsrc = cell.f;
    return ast;
  };
  /** compute one formula cell (and spread array results) */
  C.evalCell = function (sh, r, c, cell) {
    if (!cell || !cell.dirty) return;
    if (cell.computing) {
      /* circular reference: keep the previous value */
      C.circular.add(cell);
      return;
    }
    const ast = C.astOf(cell);
    cell.computing = true;
    evalStack++;
    let v;
    try {
      if (!ast) v = E.NAME;
      else if (evalStack > 2500) { v = cell.v; C.circular.add(cell); }
      else {
        const ctx = C.ctx(sh.wb, sh, r, c, !!cell.af);
        v = ev(ast, ctx);
        if (ctx.flags.keep && !cell.edited) {
          /* the formula needs data we cannot compute (pivot caches, cube or web functions): keep the saved result */
          cell.computing = false; cell.dirty = false; evalStack--; cell.kept = true;
          if (cell.af) { const a = cell.af; for (let rr = a.r1; rr <= a.r2; rr++) for (let cc = a.c1; cc <= a.c2; cc++) { const m = sh.get(rr, cc); if (m && m !== cell) m.dirty = false; } }
          if (cell.v === undefined || cell.v === null) cell.v = E.NAME;
          return;
        }
        if (cell.af) { spreadArray(sh, cell, v, ctx); cell.computing = false; cell.dirty = false; evalStack--; return; }
        v = finalValue(v, ctx, ast);
      }
    } catch (e) {
      if (typeof console !== 'undefined' && C.debug) console.error('calc error', sh.name, F.colName(c) + (r + 1), cell.f, e);
      v = E.VALUE;
    }
    evalStack--;
    cell.computing = false;
    cell.dirty = false;
    cell.v = v;
  };
  /** what a single cell shows: references dereferenced, arrays → top-left, near-zero sums cleaned */
  function finalValue(v, ctx, ast) {
    if (v instanceof Ref) v = v.single ? C.cellValue(v.sheet, v.r1, v.c1) : C.intersect(v, ctx);
    else if (v instanceof Arr) v = v.rows ? v.d[0][0] : E.VALUE;
    else if (v instanceof RefList || v instanceof Ref3D) v = E.VALUE;
    else if (v && v.lambda) v = E.CALC;
    if (v === MISSING) v = 0;
    if (v == null) v = 0; /* a reference to a blank cell shows 0 */
    if (typeof v === 'number') {
      if (!isFinite(v)) return E.NUM;
      if (v === 0) return 0; /* no -0 */
      /* Excel snaps a final addition/subtraction that cancels to within the last bits to zero */
      if (ast && ast.t === 'op' && (ast.op === '+' || ast.op === '-') && Math.abs(v) < 1e-15 * 64) {
        const a = Math.abs(numOf(ast.a, ctx)), b = Math.abs(numOf(ast.b, ctx));
        if (isFinite(a) && isFinite(b) && Math.abs(v) <= Math.max(a, b) * Math.pow(2, -49)) return 0;
      }
    }
    return v;
  }
  function numOf(node, ctx) { try { const v = operand(ev(node, ctx), ctx); return typeof v === 'number' ? v : NaN; } catch (e) { return NaN; } }
  function spreadArray(sh, master, v, ctx) {
    const a = master.af;
    let arr;
    if (v instanceof Ref) arr = C.refToArr(v);
    else if (v instanceof Arr) arr = v;
    else if (v instanceof RefList || v instanceof Ref3D) arr = new Arr([[E.VALUE]]);
    else arr = new Arr([[v == null || v === MISSING ? 0 : v]]);
    const R = arr.rows, Cc = arr.cols;
    for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) {
      const i = r - a.r1, j = c - a.c1;
      let x;
      if ((R === 1 && !arr.vr || i < R) && (Cc === 1 || j < Cc)) x = arr.d[R === 1 && !arr.vr ? 0 : i][Cc === 1 ? 0 : j];
      else if (arr.vr && i < arr.vr && (Cc === 1 || j < Cc)) x = arr.tail;
      else x = E.NA;
      if (x == null) x = 0;
      if (x && x.lambda) x = E.CALC;
      const cell = r === a.r1 && c === a.c1 ? master : sh.get(r, c);
      if (cell) { cell.v = x; if (cell !== master) cell.dirty = false; }
    }
    void ctx;
  }

  /* ------------------------------------------------------------ dependencies */
  /**
   * Static precedents of an AST evaluated at (sh, r, c): list of {sheet, r1,c1,r2,c2}; flags volatile/dynamic.
   */
  const VOLATILE = new Set(['NOW', 'TODAY', 'RAND', 'RANDBETWEEN', 'RANDARRAY', 'OFFSET', 'INDIRECT', 'CELL', 'INFO']);
  C.VOLATILE = VOLATILE;
  C.precedents = function (wb, sh, ast, r, c, out, seen) {
    out = out || { refs: [], volatile: false };
    seen = seen || new Set();
    F.walk(ast, (n) => {
      if (n.t === 'fn' && VOLATILE.has(n.name)) out.volatile = true;
      if (n.t === 'ref' || n.t === 'area') {
        if (n.book != null) return;
        const s = n.sheet != null ? wb.sheetByName(n.sheet) : sh;
        if (!s) return;
        const r1 = n.t === 'ref' ? n.r : Math.min(n.r1, n.r2), r2 = n.t === 'ref' ? n.r : Math.max(n.r1, n.r2);
        const c1 = n.t === 'ref' ? n.c : Math.min(n.c1, n.c2), c2 = n.t === 'ref' ? n.c : Math.max(n.c1, n.c2);
        if (n.sheet2) {
          const a = wb.sheetIndex(s), b = wb.sheetIndex(wb.sheetByName(n.sheet2));
          if (a < 0 || b < 0) return;
          for (let i = Math.min(a, b); i <= Math.max(a, b); i++) out.refs.push({ sheet: wb.sheets[i], r1, c1, r2, c2 });
        } else out.refs.push({ sheet: s, r1, c1, r2, c2 });
      } else if (n.t === 'name') {
        const def = wb.findName(n.name, wb.sheetIndex(sh));
        if (def && !seen.has(def)) {
          seen.add(def);
          try {
            let a = F.parse(String(def.ref).replace(/^=/, ''));
            if (hasRelative(a)) a = F.shift(a, r, c);
            C.precedents(wb, def.scope != null ? wb.sheets[def.scope] : sh, a, r, c, out, seen);
          } catch (e) { /* ignore */ }
        } else if (!def) {
          const tbl = wb.findTable(n.name);
          if (tbl) out.refs.push({ sheet: tbl.sheet, r1: tbl.ref.r1, c1: tbl.ref.c1, r2: tbl.ref.r2, c2: tbl.ref.c2 });
        }
      } else if (n.t === 'struct') {
        const tbl = n.table ? wb.findTable(n.table) : wb.tableAt(sh, r, c);
        if (tbl) out.refs.push({ sheet: tbl.sheet, r1: tbl.ref.r1, c1: tbl.ref.c1, r2: tbl.ref.r2, c2: tbl.ref.c2 });
      }
    });
    return out;
  };

  /**
   * Dependency graph for incremental recalculation.
   * Each formula cell registers its precedent areas in per-sheet column buckets.
   */
  class Graph {
    constructor(wb) { this.wb = wb; this.build(); }
    build() {
      this.idx = new Map(); /* sheet → {cols: Map<c, entries[]>, wide: entries[]} */
      this.volatile = new Set();
      this.byCell = new Map();
      for (const sh of this.wb.sheets) {
        sh.rows.forEach((row, r) => { if (row) row.cells.forEach((cell, c) => { if (cell && cell.f != null) this.add(sh, r, c, cell); }); });
      }
    }
    bucket(sh) { let b = this.idx.get(sh); if (!b) { b = { cols: new Map(), wide: [] }; this.idx.set(sh, b); } return b; }
    add(sh, r, c, cell) {
      this.remove(cell);
      const ast = C.astOf(cell);
      if (!ast) return;
      const p = C.precedents(this.wb, sh, ast, r, c);
      const node = { sh, r, c, cell, entries: [] };
      if (p.volatile) this.volatile.add(node);
      for (const a of p.refs) {
        const e = { a, node, dead: false };
        const b = this.bucket(a.sheet);
        const w = a.c2 - a.c1;
        if (w > 24) b.wide.push(e);
        else for (let k = a.c1; k <= a.c2; k++) { let l = b.cols.get(k); if (!l) { l = []; b.cols.set(k, l); } l.push(e); }
        node.entries.push(e);
      }
      cell._gnode = node;
      this.byCell.set(cell, node);
    }
    remove(cell) {
      const node = cell._gnode;
      if (!node) return;
      for (const e of node.entries) e.dead = true;
      this.volatile.delete(node);
      this.byCell.delete(cell);
      cell._gnode = null;
      this.garbage = (this.garbage || 0) + node.entries.length;
      if (this.garbage > 200000) this.compact();
    }
    compact() {
      for (const b of this.idx.values()) {
        for (const [k, l] of b.cols) { const f = l.filter((e) => !e.dead); if (f.length) b.cols.set(k, f); else b.cols.delete(k); }
        b.wide = b.wide.filter((e) => !e.dead);
      }
      this.garbage = 0;
    }
    /** formula nodes that depend directly on cell (sh, r, c) */
    dependents(sh, r, c, out) {
      const b = this.idx.get(sh);
      if (!b) return;
      const l = b.cols.get(c);
      if (l) for (const e of l) if (!e.dead && r >= e.a.r1 && r <= e.a.r2) out.push(e.node);
      for (const e of b.wide) if (!e.dead && r >= e.a.r1 && r <= e.a.r2 && c >= e.a.c1 && c <= e.a.c2) out.push(e.node);
    }
    /** mark everything that depends on the changed cells dirty; returns the dirty nodes */
    invalidate(changes) {
      const dirty = [];
      const seen = new Set();
      const queue = [];
      const push = (n) => { if (seen.has(n)) return; seen.add(n); dirty.push(n); queue.push(n); };
      for (const ch of changes) {
        const tmp = [];
        if (ch.area) { for (let r = ch.area.r1; r <= ch.area.r2; r++) for (let c = ch.area.c1; c <= ch.area.c2; c++) this.dependents(ch.sh, r, c, tmp); }
        else this.dependents(ch.sh, ch.r, ch.c, tmp);
        tmp.forEach(push);
        /* the changed cell itself, when it is a formula */
        const cell = ch.sh.get(ch.r, ch.c);
        if (cell && cell._gnode && !ch.area) push(cell._gnode);
      }
      for (const v of this.volatile) push(v);
      while (queue.length) {
        const n = queue.shift();
        const tmp = [];
        if (n.cell.af) {
          const a = n.cell.af;
          for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) this.dependents(n.sh, r, c, tmp);
        } else this.dependents(n.sh, n.r, n.c, tmp);
        tmp.forEach(push);
      }
      for (const n of dirty) n.cell.dirty = true;
      return dirty;
    }
  }
  C.Graph = Graph;

  /** evaluate dirty nodes in dependency order (precedents first) */
  C.evaluate = function (nodes) {
    C.circular = new Set();
    /* order: depth-first on precedents within the dirty set, iteratively (no recursion) */
    const order = C.order(nodes);
    for (const n of order) if (n.cell.dirty) C.evalCell(n.sh, n.r, n.c, n.cell);
    return C.circular.size;
  };
  /** topological order of formula nodes; cycles fall back to sheet order */
  C.order = function (nodes) {
    if (nodes.length < 2) return nodes;
    const set = new Set(nodes);
    /* index dirty nodes by sheet / column → sorted rows for range lookups */
    const index = new Map();
    for (const n of nodes) {
      let s = index.get(n.sh); if (!s) { s = new Map(); index.set(n.sh, s); }
      const cells = n.cell.af ? spanCols(n.cell.af) : [n.c];
      for (const c of cells) {
        let l = s.get(c); if (!l) { l = []; s.set(c, l); }
        if (n.cell.af) for (let r = n.cell.af.r1; r <= n.cell.af.r2; r++) l.push({ r, n }); else l.push({ r: n.r, n });
      }
    }
    for (const s of index.values()) for (const l of s.values()) l.sort((a, b) => a.r - b.r);
    const precOf = (n) => {
      const out = [];
      const p = n.entries || (n.cell._gnode ? n.cell._gnode.entries : []);
      for (const e of p) {
        const a = e.a;
        const s = index.get(a.sheet);
        if (!s) continue;
        if (a.c2 - a.c1 > 2048 && s.size < a.c2 - a.c1) {
          for (const [c, l] of s) if (c >= a.c1 && c <= a.c2) rowsIn(l, a.r1, a.r2, out);
        } else for (let c = a.c1; c <= a.c2; c++) { const l = s.get(c); if (l) rowsIn(l, a.r1, a.r2, out); }
      }
      return out;
    };
    const state = new Map(); /* 1 visiting, 2 done */
    const order = [];
    for (const start of nodes) {
      if (state.get(start) === 2) continue;
      const stack = [{ n: start, kids: null, i: 0 }];
      state.set(start, 1);
      while (stack.length) {
        const top = stack[stack.length - 1];
        if (!top.kids) top.kids = precOf(top.n).filter((k) => set.has(k) && k !== top.n);
        if (top.i < top.kids.length) {
          const k = top.kids[top.i++];
          const st = state.get(k);
          if (!st) { state.set(k, 1); stack.push({ n: k, kids: null, i: 0 }); }
        } else { state.set(top.n, 2); order.push(top.n); stack.pop(); }
      }
    }
    return order;
  };
  function spanCols(a) { const out = []; for (let c = a.c1; c <= a.c2; c++) out.push(c); return out; }
  function rowsIn(l, r1, r2, out) {
    /* binary search first row >= r1 */
    let lo = 0, hi = l.length;
    while (lo < hi) { const m = (lo + hi) >> 1; if (l[m].r < r1) lo = m + 1; else hi = m; }
    for (let i = lo; i < l.length && l[i].r <= r2; i++) out.push(l[i].n);
  }

  /** recalculate the whole workbook (F9 / Ctrl+Alt+F9) */
  C.recalcAll = function (wb) {
    if (!wb.graph) wb.graph = new Graph(wb);
    const nodes = Array.from(wb.graph.byCell.values());
    for (const n of nodes) n.cell.dirty = true;
    return C.evaluate(nodes);
  };
  /** after cells changed: invalidate dependents and recompute them */
  C.changed = function (wb, changes) {
    if (!wb.graph) wb.graph = new Graph(wb);
    const nodes = wb.graph.invalidate(changes);
    return C.evaluate(nodes);
  };
  /** register / unregister a cell's formula with the graph */
  C.setFormula = function (wb, sh, r, c, cell) {
    if (!wb.graph) return;
    if (cell && cell.f != null) wb.graph.add(sh, r, c, cell); else if (cell) wb.graph.remove(cell);
  };
  C.rebuild = function (wb) { wb.graph = new Graph(wb); };

  /** evaluate a formula string in a cell context (Name Manager, CF, validation, Goal Seek...) */
  C.evalText = function (wb, sh, r, c, text, array) {
    let ast;
    try { ast = F.parse(String(text).replace(/^=/, '')); } catch (e) { return E.NAME; }
    const ctx = C.ctx(wb, sh, r, c, array);
    return ev(ast, ctx);
  };
  C.evalScalar = function (wb, sh, r, c, text) { const v = C.evalText(wb, sh, r, c, text); const ctx = C.ctx(wb, sh, r, c); return finalValue(C.scalarOut(v, ctx), ctx); };
})(typeof window !== 'undefined' ? window : globalThis);
