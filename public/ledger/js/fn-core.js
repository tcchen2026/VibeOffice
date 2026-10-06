/* Ledger — worksheet functions: logical, math & trig, text, information, date & time. */
(function (root) {
  'use strict';
  const L = root.L;
  const C = L.calc, M = L.model, NF = L.numfmt, F = L.formula;
  const E = M.ERR;
  const { def, isErr, MISSING, Arr, Ref, RefList } = C;
  const opt = (v, d) => (v === MISSING || v === undefined ? d : v);
  const num = (v, ctx) => C.toNum(v, ctx);
  const trunc = Math.trunc;
  C.opt = opt;

  /* ---------------------------------------------------------------- logical */
  def('IF', 1, 3, 'l', null, { special: true, impl: (ctx, a) => {
    /* outside array formulas IF intersects ranges even inside array arguments (SUMPRODUCT(IF(…)) needs Ctrl+Shift+Enter) */
    if (ctx.array && !ctx.cse) ctx = Object.assign({}, ctx, { array: false });
    const cond = C.operand(C.ev(a[0], ctx), ctx);
    const branch = (i, dflt) => (a[i] && a[i].t !== 'missing' ? C.ev(a[i], ctx) : i === 1 && a.length > 1 ? 0 : a.length > i ? 0 : dflt);
    if (cond instanceof Arr) {
      const actx = Object.assign({}, ctx, { array: true });
      const br = (i, d) => (a.length > i ? (a[i].t === 'missing' ? 0 : C.operand(C.ev(a[i], actx), actx)) : d);
      const t = br(1, true), f = br(2, false);
      return C.broadcast3(cond, t, f, (c, x, y) => { if (isErr(c)) return c; const b = C.toBool(c); if (isErr(b)) return b; return b ? x : y; });
    }
    if (isErr(cond)) return cond;
    const b = C.toBool(cond);
    if (isErr(b)) return b;
    if (b) return a.length >= 2 ? (a[1].t === 'missing' ? 0 : C.ev(a[1], ctx)) : true;
    return a.length >= 3 ? (a[2].t === 'missing' ? 0 : C.ev(a[2], ctx)) : false;
  } });
  const logicAgg = (name, init, step, fin) => def(name, 1, 255, 'r', (ctx, args) => {
    let acc = init, n = 0;
    const err = C.each(args, (v, direct) => {
      if (isErr(v)) return v;
      let b;
      if (typeof v === 'boolean') b = v;
      else if (typeof v === 'number') b = v !== 0;
      else if (typeof v === 'string') { if (!direct) return; b = C.toBool(v); if (isErr(b)) return b; }
      else return;
      n++; acc = step(acc, b);
    });
    if (err) return err;
    if (!n) return E.VALUE;
    return fin ? fin(acc) : acc;
  }, { list: true });
  logicAgg('AND', true, (a, b) => a && b);
  logicAgg('OR', false, (a, b) => a || b);
  logicAgg('XOR', 0, (a, b) => a + (b ? 1 : 0), (x) => x % 2 === 1);
  def('NOT', 1, 1, 'b', (ctx, b) => !b);
  def('TRUE', 0, 0, '', () => true);
  def('FALSE', 0, 0, '', () => false);
  /* IFERROR / IFNA: element-wise over arrays; the alternative broadcasts against the value */
  function ifErrLike(ctx, a, test) {
    const v = C.ev(a[0], ctx);
    const x = v instanceof Ref ? (ctx.array && !v.single ? C.refToArr(v) : v) : v;
    const altOf = () => {
      if (a[1].t === 'missing') return 0;
      let alt = C.ev(a[1], ctx);
      if (alt instanceof Ref) alt = ctx.array && !alt.single ? C.refToArr(alt) : C.scalar(alt, ctx);
      if (alt === MISSING || alt == null) alt = 0;
      return alt;
    };
    if (x instanceof Arr) {
      if (!x.d.some((r) => r.some(test))) return x;
      const alt = altOf();
      if (alt instanceof Arr) return C.broadcast(x, alt, (e, f) => (test(e) ? (f == null ? 0 : f) : e));
      return new Arr(x.d.map((r) => r.map((e) => (test(e) ? alt : e))));
    }
    const s = x instanceof Ref ? C.scalar(x, ctx) : x;
    if (test(s)) { const alt = a[1].t === 'missing' ? 0 : C.ev(a[1], ctx); return alt === MISSING ? 0 : alt; }
    return v;
  }
  def('IFERROR', 2, 2, 'l', null, { special: true, impl: (ctx, a) => ifErrLike(ctx, a, isErr) });
  def('IFNA', 2, 2, 'l', null, { special: true, impl: (ctx, a) => ifErrLike(ctx, a, (e) => e === E.NA) });
  def('IFS', 2, 254, 'l', null, { special: true, impl: (ctx, a) => {
    if (a.length % 2) return E.VALUE;
    for (let i = 0; i < a.length; i += 2) {
      const c = C.scalar(C.ev(a[i], ctx), ctx);
      if (isErr(c)) return c;
      const b = C.toBool(c);
      if (isErr(b)) return b;
      if (b) return C.ev(a[i + 1], ctx);
    }
    return E.NA;
  } });
  def('SWITCH', 3, 254, 'l', null, { special: true, impl: (ctx, a) => {
    const x = C.scalar(C.ev(a[0], ctx), ctx);
    if (isErr(x)) return x;
    let i = 1;
    for (; i + 1 < a.length; i += 2) {
      const v = C.scalar(C.ev(a[i], ctx), ctx);
      if (isErr(v)) return v;
      if (C.compare(x, v) === 0 && typeof x === typeof v) return C.ev(a[i + 1], ctx);
    }
    return i < a.length ? C.ev(a[i], ctx) : E.NA;
  } });
  def('LET', 3, 253, 'l', null, { special: true, impl: (ctx, a) => {
    if (a.length % 2 === 0) return E.VALUE;
    const map = new Map();
    const sub = Object.assign({}, ctx, { vars: { map, parent: ctx.vars } });
    for (let i = 0; i + 1 < a.length; i += 2) {
      if (a[i].t !== 'name') return E.VALUE;
      map.set(a[i].name.toUpperCase(), C.ev(a[i + 1], sub));
    }
    return C.ev(a[a.length - 1], sub);
  } });
  def('LAMBDA', 1, 254, 'l', null, { special: true, impl: (ctx, a) => {
    const params = a.slice(0, -1).map((p) => (p.t === 'name' ? p.name.toUpperCase() : null));
    if (params.some((p) => !p)) return E.VALUE;
    return { lambda: true, params, body: a[a.length - 1], scope: ctx.vars };
  } });

  /* --------------------------------------------------------------- numbers */
  /** decimal rounding half away from zero after 15-digit normalisation */
  C.round = function (x, d) {
    if (!isFinite(x)) return x;
    d = trunc(d);
    if (d > 15) d = 15;
    if (d >= 0) {
      const s = NF.fixed(x, d);
      const v = parseFloat(s);
      return x < 0 ? -v : v;
    }
    const p = Math.pow(10, -d);
    const q = NF.r15(Math.abs(x) / p);
    const r = Math.floor(q + 0.5) * p;
    return x < 0 ? -r : r;
  };
  const rDir = (x, d, up) => {
    d = trunc(d);
    const p = Math.pow(10, d);
    const a = NF.r15(Math.abs(x) * p);
    const k = up ? Math.ceil(NF.r15(a)) : Math.floor(NF.r15(a));
    const r = NF.r15(k / p);
    return x < 0 ? -r : r;
  };
  def('ROUND', 2, 2, 'nn', (ctx, x, d) => C.round(x, d));
  def('ROUNDUP', 2, 2, 'nn', (ctx, x, d) => rDir(x, d, true));
  def('ROUNDDOWN', 2, 2, 'nn', (ctx, x, d) => rDir(x, d, false));
  def('TRUNC', 1, 2, 'nn', (ctx, x, d) => rDir(x, opt(d, 0), false));
  def('INT', 1, 1, 'n', (ctx, x) => Math.floor(NF.r15(x)));
  def('ABS', 1, 1, 'n', (ctx, x) => Math.abs(x));
  def('SIGN', 1, 1, 'n', (ctx, x) => (x > 0 ? 1 : x < 0 ? -1 : 0));
  def('MOD', 2, 2, 'nn', (ctx, n, d) => {
    if (d === 0) return E.DIV0;
    const q = n / d;
    if (Math.abs(q) > 1.3421773e8 * 8) return E.NUM;
    let r = n - d * Math.floor(NF.r15(q));
    if (Math.abs(r) < 1e-15 * Math.abs(d)) r = 0;
    if (r !== 0 && Math.sign(r) !== Math.sign(d)) r += d;
    return NF.r15(r) === 0 ? 0 : r;
  });
  def('QUOTIENT', 2, 2, 'nn', (ctx, n, d) => (d === 0 ? E.DIV0 : trunc(NF.r15(n / d))));
  def('PI', 0, 0, '', () => Math.PI);
  def('SQRT', 1, 1, 'n', (ctx, x) => (x < 0 ? E.NUM : Math.sqrt(x)));
  def('SQRTPI', 1, 1, 'n', (ctx, x) => (x < 0 ? E.NUM : Math.sqrt(x * Math.PI)));
  def('EXP', 1, 1, 'n', (ctx, x) => Math.exp(x));
  def('LN', 1, 1, 'n', (ctx, x) => (x <= 0 ? E.NUM : Math.log(x)));
  def('LOG10', 1, 1, 'n', (ctx, x) => (x <= 0 ? E.NUM : Math.log10(x)));
  def('LOG', 1, 2, 'nn', (ctx, x, b) => { b = opt(b, 10); if (x <= 0 || b <= 0) return E.NUM; if (b === 1) return E.DIV0; return b === 10 ? Math.log10(x) : Math.log(x) / Math.log(b); });
  def('POWER', 2, 2, 'nn', (ctx, a, b) => C.OPS['^'](a, b, ctx));
  const trig = (name, f, dom) => def(name, 1, 1, 'n', (ctx, x) => { if (dom && !dom(x)) return E.NUM; const r = f(x); return isFinite(r) ? r : E.NUM; });
  trig('SIN', Math.sin, (x) => Math.abs(x) < 134217728); trig('COS', Math.cos, (x) => Math.abs(x) < 134217728); trig('TAN', Math.tan, (x) => Math.abs(x) < 134217728);
  trig('ASIN', Math.asin, (x) => x >= -1 && x <= 1); trig('ACOS', Math.acos, (x) => x >= -1 && x <= 1); trig('ATAN', Math.atan);
  trig('SINH', Math.sinh); trig('COSH', Math.cosh); trig('TANH', Math.tanh); trig('ASINH', Math.asinh); trig('ACOSH', Math.acosh, (x) => x >= 1); trig('ATANH', Math.atanh, (x) => x > -1 && x < 1);
  def('COT', 1, 1, 'n', (ctx, x) => (x === 0 ? E.DIV0 : 1 / Math.tan(x)));
  def('COTH', 1, 1, 'n', (ctx, x) => (x === 0 ? E.DIV0 : 1 / Math.tanh(x)));
  def('CSC', 1, 1, 'n', (ctx, x) => (x === 0 ? E.DIV0 : 1 / Math.sin(x)));
  def('CSCH', 1, 1, 'n', (ctx, x) => (x === 0 ? E.DIV0 : 1 / Math.sinh(x)));
  def('SEC', 1, 1, 'n', (ctx, x) => 1 / Math.cos(x));
  def('SECH', 1, 1, 'n', (ctx, x) => 1 / Math.cosh(x));
  def('ACOT', 1, 1, 'n', (ctx, x) => Math.PI / 2 - Math.atan(x));
  def('ACOTH', 1, 1, 'n', (ctx, x) => (Math.abs(x) <= 1 ? E.NUM : 0.5 * Math.log((x + 1) / (x - 1))));
  def('ATAN2', 2, 2, 'nn', (ctx, x, y) => (x === 0 && y === 0 ? E.DIV0 : Math.atan2(y, x)));
  def('RADIANS', 1, 1, 'n', (ctx, x) => (x * Math.PI) / 180);
  def('DEGREES', 1, 1, 'n', (ctx, x) => (x * 180) / Math.PI);
  def('EVEN', 1, 1, 'n', (ctx, x) => { const a = Math.ceil(NF.r15(Math.abs(x)) / 2) * 2; return x < 0 ? -a : a; });
  def('ODD', 1, 1, 'n', (ctx, x) => { let a = Math.ceil(NF.r15(Math.abs(x))); if (a % 2 === 0) a += 1; return x < 0 ? -a : a; });
  const factTab = [1];
  for (let i = 1; i <= 170; i++) factTab[i] = factTab[i - 1] * i;
  C.fact = (n) => (n < 0 ? NaN : n <= 170 ? factTab[n] : Infinity);
  def('FACT', 1, 1, 'n', (ctx, n) => { n = trunc(n); if (n < 0) return E.NUM; return n > 170 ? E.NUM : factTab[n]; });
  def('FACTDOUBLE', 1, 1, 'n', (ctx, n) => { n = trunc(n); if (n < -1) return E.NUM; let r = 1; for (let k = n; k > 1; k -= 2) r *= k; return r; });
  C.combin = (n, k) => { if (k < 0 || n < k) return NaN; k = Math.min(k, n - k); let r = 1; for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i; return Math.round(r) === r ? r : r < 1e15 ? Math.round(r) : r; };
  def('COMBIN', 2, 2, 'nn', (ctx, n, k) => { n = trunc(n); k = trunc(k); if (n < 0 || k < 0 || n < k) return E.NUM; return C.combin(n, k); });
  def('COMBINA', 2, 2, 'nn', (ctx, n, k) => { n = trunc(n); k = trunc(k); if (n < 0 || k < 0 || (n === 0 && k > 0)) return E.NUM; return k === 0 ? 1 : C.combin(n + k - 1, k); });
  def('PERMUT', 2, 2, 'nn', (ctx, n, k) => { n = trunc(n); k = trunc(k); if (n <= 0 || k < 0 || n < k) return E.NUM; let r = 1; for (let i = 0; i < k; i++) r *= n - i; return r; });
  def('PERMUTATIONA', 2, 2, 'nn', (ctx, n, k) => { n = trunc(n); k = trunc(k); if (n < 0 || k < 0) return E.NUM; return Math.pow(n, k); });
  const gcd = (a, b) => { while (b) { [a, b] = [b, a % b]; } return a; };
  const intList = (args, ctx) => { const v = C.nums(args, ctx); if (isErr(v)) return v; for (const x of v) if (x < 0 || x >= 9.007199254740992e15) return E.NUM; return v.map(Math.trunc); };
  def('GCD', 1, 255, 'r', (ctx, args) => { const v = intList(args, ctx); if (isErr(v)) return v; return v.reduce((a, b) => gcd(a, b), 0); }, { list: true });
  def('LCM', 1, 255, 'r', (ctx, args) => { const v = intList(args, ctx); if (isErr(v)) return v; let r = 1; for (const x of v) { if (x === 0) return 0; r = (r * x) / gcd(r, x); } return r; }, { list: true });
  def('MULTINOMIAL', 1, 255, 'r', (ctx, args) => { const v = C.nums(args, ctx); if (isErr(v)) return v; let s = 0, d = 1; for (let x of v) { x = trunc(x); if (x < 0) return E.NUM; s += x; d *= C.fact(x); } return C.fact(s) / d; }, { list: true });
  /* CEILING / FLOOR families */
  const q15 = (x) => NF.r15(x);
  def('CEILING', 1, 2, 'nn', (ctx, x, s) => {
    s = opt(s, x >= 0 ? 1 : -1);
    if (s === 0 || x === 0) return 0;
    if (x > 0 && s < 0) return E.NUM;
    if (x < 0 && s > 0) return -Math.floor(q15(-x / s)) * s;
    return Math.ceil(q15(x / s)) * s;
  });
  def('FLOOR', 1, 2, 'nn', (ctx, x, s) => {
    s = opt(s, x >= 0 ? 1 : -1);
    if (x === 0) return 0;
    if (s === 0) return E.DIV0;
    if (x > 0 && s < 0) return E.NUM;
    if (x < 0 && s > 0) return -Math.ceil(q15(-x / s)) * s;
    return Math.floor(q15(x / s)) * s;
  });
  def('CEILING.MATH', 1, 3, 'nnn', (ctx, x, s, mode) => {
    s = Math.abs(opt(s, 1)); mode = opt(mode, 0);
    if (s === 0 || x === 0) return 0;
    if (x < 0 && mode) return -Math.ceil(q15(-x / s)) * s;
    return Math.ceil(q15(x / s)) * s;
  });
  def('FLOOR.MATH', 1, 3, 'nnn', (ctx, x, s, mode) => {
    s = Math.abs(opt(s, 1)); mode = opt(mode, 0);
    if (s === 0 || x === 0) return 0;
    if (x < 0 && mode) return -Math.floor(q15(-x / s)) * s;
    return Math.floor(q15(x / s)) * s;
  });
  def('CEILING.PRECISE', 1, 2, 'nn', (ctx, x, s) => { s = Math.abs(opt(s, 1)); if (s === 0 || x === 0) return 0; return Math.ceil(q15(x / s)) * s; });
  C.alias('ISO.CEILING', 'CEILING.PRECISE');
  C.alias('ECMA.CEILING', 'CEILING');
  def('FLOOR.PRECISE', 1, 2, 'nn', (ctx, x, s) => { s = Math.abs(opt(s, 1)); if (s === 0 || x === 0) return 0; return Math.floor(q15(x / s)) * s; });
  def('MROUND', 2, 2, 'nn', (ctx, n, m) => { if (m === 0) return 0; if (n * m < 0) return E.NUM; const q = q15(n / m); return Math.sign(q) * Math.floor(Math.abs(q) + 0.5) * m; });

  /* sums & products */
  def('SUM', 1, 255, 'r', (ctx, args) => { let s = 0; const e = C.each(args, (v, d) => { if (typeof v === 'number') s += v; else if (isErr(v)) return v; else if (d) { if (typeof v === 'boolean') s += v ? 1 : 0; else if (typeof v === 'string') { const n = C.parseNum(v, ctx.wb); if (n == null) return E.VALUE; s += n; } } }); return e || s; }, { list: true });
  def('SUMSQ', 1, 255, 'r', (ctx, args) => { const v = C.nums(args, ctx); if (isErr(v)) return v; return v.reduce((a, b) => a + b * b, 0); }, { list: true });
  def('PRODUCT', 1, 255, 'r', (ctx, args) => { const v = C.nums(args, ctx); if (isErr(v)) return v; return v.length ? v.reduce((a, b) => a * b, 1) : 0; }, { list: true });
  def('SUMPRODUCT', 1, 255, 'a', (ctx, args) => {
    const arrs = args.map((a) => (a === MISSING ? null : C.toArr(a)));
    if (arrs.some((a) => !a)) return E.VALUE;
    const R = arrs[0].rows, Cc = arrs[0].cols;
    for (const a of arrs) if (a.rows !== R || a.cols !== Cc) return E.VALUE;
    let s = 0;
    for (let i = 0; i < R; i++) for (let j = 0; j < Cc; j++) {
      let p = 1;
      for (const a of arrs) { const v = a.d[i][j]; if (isErr(v)) return v; p *= typeof v === 'number' ? v : 0; }
      s += p;
    }
    return s;
  }, { list: true });
  const pairs = (a, b, f) => {
    const A = C.toArr(a), B = C.toArr(b);
    if (A.rows * A.cols !== B.rows * B.cols) return E.NA;
    const fa = A.d.flat(), fb = B.d.flat();
    let s = 0;
    for (let i = 0; i < fa.length; i++) { if (isErr(fa[i])) return fa[i]; if (isErr(fb[i])) return fb[i]; if (typeof fa[i] === 'number' && typeof fb[i] === 'number') s += f(fa[i], fb[i]); }
    return s;
  };
  def('SUMX2MY2', 2, 2, 'aa', (ctx, a, b) => pairs(a, b, (x, y) => x * x - y * y));
  def('SUMX2PY2', 2, 2, 'aa', (ctx, a, b) => pairs(a, b, (x, y) => x * x + y * y));
  def('SUMXMY2', 2, 2, 'aa', (ctx, a, b) => pairs(a, b, (x, y) => (x - y) * (x - y)));
  def('SERIESSUM', 4, 4, 'nnna', (ctx, x, n, m, co) => { const a = C.toArr(co).d.flat(); let s = 0; for (let i = 0; i < a.length; i++) { if (typeof a[i] !== 'number') return E.VALUE; s += a[i] * Math.pow(x, n + i * m); } return s; });

  /* conditional aggregation */
  /** pairs (criteria range, criteria) → predicate over (row offset, col offset) */
  function condSet(ctx, list) {
    const conds = [];
    let R = -1, Cc = -1;
    for (let i = 0; i < list.length; i += 2) {
      const rg = list[i];
      if (!(rg instanceof Ref) && !(rg instanceof Arr)) return E.VALUE;
      const rows = rg.rows, cols = rg.cols;
      if (R < 0) { R = rows; Cc = cols; } else if (rows !== R || cols !== Cc) return E.VALUE;
      let crit = list[i + 1];
      if (crit instanceof Ref || crit instanceof Arr) crit = C.scalar(crit, ctx);
      conds.push({ rg, pred: C.criteria(crit === MISSING ? '' : crit, ctx) });
    }
    return { conds, R, Cc };
  }
  const valAt = (rg, i, j) => (rg instanceof Ref ? C.cellValue(rg.sheet, rg.r1 + i, rg.c1 + j) : rg.d[i][j]);
  /** iterate matching offsets; for large sparse references walk only stored cells of the first range */
  function eachMatch(cs, fn) {
    const { conds, R, Cc } = cs;
    const first = conds[0];
    const blankOK = conds.map((c) => c.pred(null));
    if (first.rg instanceof Ref && !blankOK[0]) {
      const ref = first.rg;
      C.eachRef(ref, (v, d, r, c) => {
        const i = r - ref.r1, j = c - ref.c1;
        if (!first.pred(v)) return;
        for (let k = 1; k < conds.length; k++) if (!conds[k].pred(valAt(conds[k].rg, i, j))) return;
        fn(i, j);
      });
      return;
    }
    const sh = first.rg instanceof Ref ? first.rg.sheet : null;
    const rMax = sh ? Math.min(R, Math.max(0, sh.maxR - first.rg.r1 + 1)) : R;
    for (let i = 0; i < R; i++) for (let j = 0; j < Cc; j++) {
      if (i >= rMax && !blankOK.every(Boolean)) break;
      let ok = true;
      for (const c of conds) if (!c.pred(i >= rMax && c.rg === first.rg ? null : valAt(c.rg, i, j))) { ok = false; break; }
      if (ok) fn(i, j);
    }
  }
  /** *IFS functions with array criteria return an array: one result per criterion */
  const critLift = (impl, start) => (ctx, args) => {
    let lifted = null;
    for (let i = start; i < args.length; i += 2) {
      let v = args[i];
      if (v instanceof Ref && !v.single && ctx.array) v = C.refToArr(v);
      if (v instanceof Arr && (v.rows > 1 || v.cols > 1)) { if (!lifted) { lifted = []; args = args.slice(); } args[i] = v; lifted.push(i); }
    }
    if (!lifted) return impl(ctx, args);
    let R = 1, Cc = 1;
    for (const i of lifted) { R = Math.max(R, args[i].rows); Cc = Math.max(Cc, args[i].cols); }
    const d = [];
    for (let r = 0; r < R; r++) {
      const row = [];
      for (let c = 0; c < Cc; c++) {
        const a = args.slice();
        let bad = null;
        for (const i of lifted) {
          const A = args[i];
          const rr = A.rows === 1 ? 0 : r, cc = A.cols === 1 ? 0 : c;
          if (rr >= A.rows || cc >= A.cols) { bad = E.NA; break; }
          a[i] = A.d[rr][cc];
        }
        row.push(bad || impl(ctx, a));
      }
      d.push(row);
    }
    return new Arr(d);
  };
  /** SUMIF's sum_range takes the shape of the criteria range from its own top-left cell */
  const reshape = (sum, rg) => (sum instanceof Ref ? new Ref(sum.sheet, sum.r1, sum.c1, sum.r1 + rg.rows - 1, sum.c1 + rg.cols - 1) : sum);
  def('SUMIF', 2, 3, 'rvr', (ctx, rg, crit, sum) => {
    if (!(rg instanceof Ref) && !(rg instanceof Arr)) return E.VALUE;
    const cs = condSet(ctx, [rg, crit]); if (isErr(cs)) return cs;
    const S = sum === MISSING ? rg : reshape(sum, rg);
    let s = 0, err = null;
    eachMatch(cs, (i, j) => { const v = valAt(S, i, j); if (typeof v === 'number') s += v; else if (isErr(v) && !err) err = v; });
    return err || s;
  }, { errOk: true });
  def('SUMIFS', 3, 255, 'r', critLift((ctx, args) => {
    if (args.length % 2 === 0) return E.VALUE;
    const S = args[0];
    const cs = condSet(ctx, args.slice(1)); if (isErr(cs)) return cs;
    if (!(S instanceof Ref || S instanceof Arr) || S.rows !== cs.R || S.cols !== cs.Cc) return E.VALUE;
    let s = 0, err = null;
    eachMatch(cs, (i, j) => { const v = valAt(S, i, j); if (typeof v === 'number') s += v; else if (isErr(v) && !err) err = v; });
    return err || s;
  }, 2), { list: true });
  def('COUNTIF', 2, 2, 'rv', (ctx, rg, crit) => { if (!(rg instanceof Ref) && !(rg instanceof Arr)) return E.VALUE; const cs = condSet(ctx, [rg, crit]); if (isErr(cs)) return cs; let n = 0; eachMatch(cs, () => n++); return n; }, { errOk: true });
  def('COUNTIFS', 2, 254, 'r', critLift((ctx, args) => { if (args.length % 2) return E.VALUE; const cs = condSet(ctx, args); if (isErr(cs)) return cs; let n = 0; eachMatch(cs, () => n++); return n; }, 1), { list: true });
  def('AVERAGEIF', 2, 3, 'rvr', (ctx, rg, crit, avg) => {
    if (!(rg instanceof Ref) && !(rg instanceof Arr)) return E.VALUE;
    const cs = condSet(ctx, [rg, crit]); if (isErr(cs)) return cs;
    const S = avg === MISSING ? rg : reshape(avg, rg);
    let s = 0, n = 0, err = null;
    eachMatch(cs, (i, j) => { const v = valAt(S, i, j); if (typeof v === 'number') { s += v; n++; } else if (isErr(v) && !err) err = v; });
    return err || (n ? s / n : E.DIV0);
  }, { errOk: true });
  def('AVERAGEIFS', 3, 255, 'r', critLift((ctx, args) => {
    if (args.length % 2 === 0) return E.VALUE;
    const S = args[0];
    const cs = condSet(ctx, args.slice(1)); if (isErr(cs)) return cs;
    if (!(S instanceof Ref || S instanceof Arr) || S.rows !== cs.R || S.cols !== cs.Cc) return E.VALUE;
    let s = 0, n = 0, err = null;
    eachMatch(cs, (i, j) => { const v = valAt(S, i, j); if (typeof v === 'number') { s += v; n++; } else if (isErr(v) && !err) err = v; });
    return err || (n ? s / n : E.DIV0);
  }, 2), { list: true });
  const minmaxIfs = (isMax) => (ctx, args) => {
    if (args.length % 2 === 0) return E.VALUE;
    const S = args[0];
    const cs = condSet(ctx, args.slice(1)); if (isErr(cs)) return cs;
    if (!(S instanceof Ref || S instanceof Arr) || S.rows !== cs.R || S.cols !== cs.Cc) return E.VALUE;
    let best = null, err = null;
    eachMatch(cs, (i, j) => { const v = valAt(S, i, j); if (typeof v === 'number') { if (best == null || (isMax ? v > best : v < best)) best = v; } else if (isErr(v) && !err) err = v; });
    return err || (best == null ? 0 : best);
  };
  def('MAXIFS', 3, 255, 'r', critLift(minmaxIfs(true), 2), { list: true });
  def('MINIFS', 3, 255, 'r', critLift(minmaxIfs(false), 2), { list: true });

  /* SUBTOTAL / AGGREGATE */
  const SUBFN = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };
  def('SUBTOTAL', 2, 255, 'nr', (ctx, args) => {
    const code = trunc(args[0]);
    if (isErr(args[0])) return args[0];
    const base = code > 100 ? code - 100 : code;
    const name = SUBFN[base];
    if (!name) return E.VALUE;
    const refs = args.slice(1);
    for (const r of refs) { if (isErr(r)) return r; if (!(r instanceof Ref) && !(r instanceof RefList) && !(r instanceof C.Ref3D)) return E.VALUE; }
    const vals = [];
    let err = null;
    for (const r of refs) {
      const list = r instanceof Ref ? [r] : r.list;
      for (const x of list) C.eachRef(x, (v, d, rr) => {
        const row = x.sheet.rows[rr];
        if (row && row.hidden && (code > 100 || row.filtered)) return;
        if (isErr(v)) { if (!err) err = v; return; }
        vals.push(v);
      }, { skipSubtotals: true });
    }
    if (err && base !== 2 && base !== 3) return err;
    return C.aggregate(name, vals, ctx);
  }, { list: true });
  /** apply a statistical function to plain values (from references: text/booleans ignored) */
  C.aggregate = function (name, vals, ctx) {
    const nums = vals.filter((v) => typeof v === 'number');
    switch (name) {
      case 'SUM': return nums.reduce((a, b) => a + b, 0);
      case 'COUNT': return nums.length;
      case 'COUNTA': return vals.filter((v) => v != null).length;
      case 'AVERAGE': return nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : E.DIV0;
      case 'MAX': return nums.length ? Math.max(...nums) : 0;
      case 'MIN': return nums.length ? Math.min(...nums) : 0;
      case 'PRODUCT': return nums.length ? nums.reduce((a, b) => a * b, 1) : 0;
      default: return C.FN[name] ? C.FN[name].impl(ctx, [new Arr([nums])]) : E.VALUE;
    }
  };
  const AGGFN = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S', 11: 'VAR.P', 12: 'MEDIAN', 13: 'MODE.SNGL', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC', 18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC' };
  def('AGGREGATE', 3, 255, 'nnr', (ctx, args) => {
    const code = trunc(args[0]), o = trunc(args[1]);
    if (isErr(args[0])) return args[0];
    if (isErr(args[1])) return args[1];
    const name = AGGFN[code];
    if (!name || o < 0 || o > 7) return E.VALUE;
    const ignHidden = o === 1 || o === 3 || o === 5 || o === 7, ignErr = o === 2 || o === 3 || o === 6 || o === 7, ignSub = o <= 3;
    const vals = [];
    let err = null;
    const take = (x) => {
      if (x instanceof Ref) {
        C.eachRef(x, (v, d, rr) => {
          const row = x.sheet.rows[rr];
          if (ignHidden && row && row.hidden) return;
          if (isErr(v)) { if (!ignErr && !err) err = v; return; }
          vals.push(v);
        }, { skipSubtotals: ignSub });
      } else if (x instanceof RefList) x.list.forEach(take);
      else if (x instanceof Arr) x.d.flat().forEach((v) => { if (isErr(v)) { if (!ignErr && !err) err = v; } else if (v != null) vals.push(v); });
      else if (x !== MISSING) { if (isErr(x)) { if (!ignErr && !err) err = x; } else vals.push(typeof x === 'string' ? (C.parseNum(x, ctx.wb) ?? x) : x); }
    };
    if (code >= 14) {
      take(args[2]);
      if (err) return err;
      const k = C.scalar(args[3], ctx);
      if (k === undefined) return E.VALUE;
      return C.FN[name].impl(ctx, new Arr([vals.filter((v) => typeof v === 'number')]), k);
    }
    args.slice(2).forEach(take);
    if (err) return err;
    return C.aggregate(name, vals, ctx);
  }, { list: true });

  /* random */
  def('RAND', 0, 0, '', () => Math.random(), { volatile: true });
  def('RANDBETWEEN', 2, 2, 'nn', (ctx, a, b) => { a = Math.ceil(a); b = Math.floor(b); if (a > b) return E.NUM; return a + Math.floor(Math.random() * (b - a + 1)); });
  def('RANDARRAY', 0, 5, 'nnnnb', (ctx, r, c, mn, mx, whole) => {
    r = trunc(opt(r, 1)); c = trunc(opt(c, 1)); mn = opt(mn, 0); mx = opt(mx, 1); whole = opt(whole, false);
    if (r < 1 || c < 1 || mn > mx) return E.VALUE;
    return new Arr(Array.from({ length: r }, () => Array.from({ length: c }, () => (whole ? Math.floor(mn + Math.random() * (mx - mn + 1)) : mn + Math.random() * (mx - mn)))));
  });
  def('SEQUENCE', 1, 4, 'nnnn', (ctx, r, c, start, step) => {
    r = trunc(r); c = trunc(opt(c, 1)); start = opt(start, 1); step = opt(step, 1);
    if (r < 1 || c < 1) return E.CALC;
    return new Arr(Array.from({ length: r }, (_, i) => Array.from({ length: c }, (_, j) => start + (i * c + j) * step)));
  });
  /* roman / bases */
  def('ROMAN', 1, 2, 'nn', (ctx, n, form) => {
    n = trunc(n); form = opt(form, 0);
    if (typeof form === 'boolean') form = form ? 0 : 4;
    if (n < 0 || n > 3999) return E.VALUE;
    if (n === 0) return '';
    const tables = [
      [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']],
      [[1000, 'M'], [950, 'LM'], [900, 'CM'], [500, 'D'], [450, 'LD'], [400, 'CD'], [100, 'C'], [95, 'VC'], [90, 'XC'], [50, 'L'], [45, 'VL'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']],
      [[1000, 'M'], [990, 'XM'], [950, 'LM'], [900, 'CM'], [500, 'D'], [490, 'XD'], [450, 'LD'], [400, 'CD'], [100, 'C'], [99, 'IC'], [95, 'VC'], [90, 'XC'], [50, 'L'], [49, 'IL'], [45, 'VL'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']],
      [[1000, 'M'], [995, 'VM'], [990, 'XM'], [950, 'LM'], [900, 'CM'], [500, 'D'], [495, 'VD'], [490, 'XD'], [450, 'LD'], [400, 'CD'], [100, 'C'], [99, 'IC'], [95, 'VC'], [90, 'XC'], [50, 'L'], [49, 'IL'], [45, 'VL'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']],
      [[1000, 'M'], [999, 'IM'], [995, 'VM'], [990, 'XM'], [950, 'LM'], [900, 'CM'], [500, 'D'], [499, 'ID'], [495, 'VD'], [490, 'XD'], [450, 'LD'], [400, 'CD'], [100, 'C'], [99, 'IC'], [95, 'VC'], [90, 'XC'], [50, 'L'], [49, 'IL'], [45, 'VL'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']],
    ];
    const t = tables[Math.max(0, Math.min(4, trunc(form)))];
    let s = '';
    for (const [v, r] of t) while (n >= v) { s += r; n -= v; }
    return s;
  });
  def('ARABIC', 1, 1, 's', (ctx, s) => {
    s = s.trim().toUpperCase();
    if (!s) return 0;
    let neg = false;
    if (s[0] === '-') { neg = true; s = s.slice(1); }
    const V = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
    let t = 0;
    for (let i = 0; i < s.length; i++) { const a = V[s[i]], b = V[s[i + 1]]; if (!a) return E.VALUE; t += b > a ? -a : a; }
    return neg ? -t : t;
  });
  def('BASE', 2, 3, 'nnn', (ctx, n, radix, len) => {
    n = trunc(n); radix = trunc(radix);
    if (n < 0 || n >= 9007199254740992 || radix < 2 || radix > 36) return E.NUM;
    let s = n.toString(radix).toUpperCase();
    len = opt(len, 0);
    if (len < 0 || len > 255) return E.NUM;
    while (s.length < len) s = '0' + s;
    return s;
  });
  def('DECIMAL', 2, 2, 'sn', (ctx, s, radix) => {
    radix = trunc(radix);
    if (radix < 2 || radix > 36) return E.NUM;
    s = s.trim();
    if (radix === 16) s = s.replace(/^0x/i, '');
    if (!s || s.length > 255) return s ? E.NUM : 0;
    let v = 0;
    for (const ch of s.toUpperCase()) { const d = parseInt(ch, 36); if (isNaN(d) || d >= radix) return E.NUM; v = v * radix + d; }
    return v;
  });
  /* matrices */
  const matrix = (a) => { const A = C.toArr(a); for (const r of A.d) for (const v of r) if (typeof v !== 'number') return isErr(v) ? v : E.VALUE; return A.d; };
  def('MMULT', 2, 2, 'aa', (ctx, a, b) => {
    const A = matrix(a), B = matrix(b);
    if (isErr(A)) return A; if (isErr(B)) return B;
    if (A[0].length !== B.length) return E.VALUE;
    const out = A.map((row) => B[0].map((_, j) => row.reduce((s, v, k) => s + v * B[k][j], 0)));
    return new Arr(out);
  });
  const lu = (A) => {
    const n = A.length;
    const a = A.map((r) => r.slice());
    const perm = [...Array(n).keys()];
    let sign = 1;
    for (let k = 0; k < n; k++) {
      let p = k;
      for (let i = k + 1; i < n; i++) if (Math.abs(a[i][k]) > Math.abs(a[p][k])) p = i;
      if (Math.abs(a[p][k]) < 1e-300) return { singular: true, a, perm, sign: 0 };
      if (p !== k) { [a[p], a[k]] = [a[k], a[p]]; [perm[p], perm[k]] = [perm[k], perm[p]]; sign = -sign; }
      for (let i = k + 1; i < n; i++) { a[i][k] /= a[k][k]; for (let j = k + 1; j < n; j++) a[i][j] -= a[i][k] * a[k][j]; }
    }
    return { a, perm, sign };
  };
  def('MDETERM', 1, 1, 'a', (ctx, a) => {
    const A = matrix(a); if (isErr(A)) return A;
    if (A.length !== A[0].length) return E.VALUE;
    const d = lu(A);
    if (d.singular) return 0;
    let det = d.sign;
    for (let i = 0; i < A.length; i++) det *= d.a[i][i];
    return det;
  });
  def('MINVERSE', 1, 1, 'a', (ctx, a) => {
    const A = matrix(a); if (isErr(A)) return A;
    const n = A.length;
    if (n !== A[0].length) return E.VALUE;
    const d = lu(A);
    if (d.singular) return E.NUM;
    const inv = [];
    for (let j = 0; j < n; j++) {
      const b = d.perm.map((p) => (p === j ? 1 : 0));
      for (let i = 0; i < n; i++) for (let k = 0; k < i; k++) b[i] -= d.a[i][k] * b[k];
      for (let i = n - 1; i >= 0; i--) { for (let k = i + 1; k < n; k++) b[i] -= d.a[i][k] * b[k]; b[i] /= d.a[i][i]; }
      for (let i = 0; i < n; i++) (inv[i] || (inv[i] = []))[j] = b[i];
    }
    return new Arr(inv);
  });
  def('MUNIT', 1, 1, 'n', (ctx, n) => { n = trunc(n); if (n < 1) return E.VALUE; return new Arr(Array.from({ length: n }, (_, i) => Array.from({ length: n }, (_, j) => (i === j ? 1 : 0)))); });

  /* ------------------------------------------------------------------ text */
  const W1252 = { 128: 0x20AC, 130: 0x201A, 131: 0x0192, 132: 0x201E, 133: 0x2026, 134: 0x2020, 135: 0x2021, 136: 0x02C6, 137: 0x2030, 138: 0x0160, 139: 0x2039, 140: 0x0152, 142: 0x017D, 145: 0x2018, 146: 0x2019, 147: 0x201C, 148: 0x201D, 149: 0x2022, 150: 0x2013, 151: 0x2014, 152: 0x02DC, 153: 0x2122, 154: 0x0161, 155: 0x203A, 156: 0x0153, 158: 0x017E, 159: 0x0178 };
  const W1252R = {};
  for (const k in W1252) W1252R[W1252[k]] = +k;
  def('CHAR', 1, 1, 'n', (ctx, n) => { n = trunc(n); if (n < 1 || n > 255) return E.VALUE; return String.fromCharCode(W1252[n] || n); });
  def('CODE', 1, 1, 's', (ctx, s) => { if (!s) return E.VALUE; const c = s.charCodeAt(0); return W1252R[c] || (c > 255 ? 63 : c); });
  def('UNICHAR', 1, 1, 'n', (ctx, n) => { n = trunc(n); if (n < 1 || n > 0x10FFFF || (n >= 0xD800 && n <= 0xDFFF)) return E.VALUE; return String.fromCodePoint(n); });
  def('UNICODE', 1, 1, 's', (ctx, s) => (s ? s.codePointAt(0) : E.VALUE));
  def('LEN', 1, 1, 's', (ctx, s) => s.length);
  C.alias('LENB', 'LEN');
  def('LEFT', 1, 2, 'sn', (ctx, s, n) => { n = trunc(opt(n, 1)); if (n < 0) return E.VALUE; return s.slice(0, n); });
  def('RIGHT', 1, 2, 'sn', (ctx, s, n) => { n = trunc(opt(n, 1)); if (n < 0) return E.VALUE; return n === 0 ? '' : s.slice(-n); });
  def('MID', 3, 3, 'snn', (ctx, s, a, n) => { a = trunc(a); n = trunc(n); if (a < 1 || n < 0) return E.VALUE; return s.substr(a - 1, n); });
  C.alias('LEFTB', 'LEFT'); C.alias('RIGHTB', 'RIGHT'); C.alias('MIDB', 'MID');
  def('FIND', 2, 3, 'ssn', (ctx, f, s, start) => { start = trunc(opt(start, 1)); if (start < 1 || start > s.length + 1) return E.VALUE; const i = s.indexOf(f, start - 1); return i < 0 ? E.VALUE : i + 1; });
  def('SEARCH', 2, 3, 'ssn', (ctx, f, s, start) => {
    start = trunc(opt(start, 1));
    if (start < 1 || start > s.length + 1) return E.VALUE;
    if (/[*?~]/.test(f)) {
      const re = new RegExp(C.wildRe(f).source.slice(1, -1), 'i');
      const m = re.exec(s.slice(start - 1));
      return m ? m.index + start : E.VALUE;
    }
    const i = s.toLowerCase().indexOf(f.toLowerCase(), start - 1);
    return i < 0 ? E.VALUE : i + 1;
  });
  C.alias('FINDB', 'FIND'); C.alias('SEARCHB', 'SEARCH');
  def('REPLACE', 4, 4, 'snns', (ctx, s, a, n, t) => { a = trunc(a); n = trunc(n); if (a < 1 || n < 0) return E.VALUE; return s.slice(0, a - 1) + t + s.slice(a - 1 + n); });
  C.alias('REPLACEB', 'REPLACE');
  def('SUBSTITUTE', 3, 4, 'sssn', (ctx, s, o, n, k) => {
    if (o === '') return s;
    if (k === MISSING || k === undefined) return s.split(o).join(n);
    k = trunc(k);
    if (k < 1) return E.VALUE;
    let i = -1;
    for (let c = 0; c < k; c++) { i = s.indexOf(o, i + 1); if (i < 0) return s; }
    return s.slice(0, i) + n + s.slice(i + o.length);
  });
  def('UPPER', 1, 1, 's', (ctx, s) => s.toUpperCase());
  def('LOWER', 1, 1, 's', (ctx, s) => s.toLowerCase());
  def('PROPER', 1, 1, 's', (ctx, s) => s.toLowerCase().replace(/(^|[^A-Za-zÀ-ɏ])([a-zà-ɏ])/g, (m, a, b) => a + b.toUpperCase()));
  def('TRIM', 1, 1, 's', (ctx, s) => s.replace(/^ +| +$/g, '').replace(/ {2,}/g, ' '));
  def('CLEAN', 1, 1, 's', (ctx, s) => s.replace(/[\u0000-\u001F]/g, ''));
  def('REPT', 2, 2, 'sn', (ctx, s, n) => { n = trunc(n); if (n < 0 || s.length * n > 32767) return E.VALUE; return s.repeat(n); });
  def('EXACT', 2, 2, 'ss', (ctx, a, b) => a === b);
  def('CONCATENATE', 1, 255, 's', (ctx, ...a) => { const s = a.map((x) => (x === MISSING ? '' : x)).join(''); return s.length > 32767 ? E.VALUE : s; });
  def('CONCAT', 1, 254, 'r', (ctx, args) => {
    let s = '';
    const e = C.each(args, (v) => { if (isErr(v)) return v; s += C.toStr(v); });
    return e || (s.length > 32767 ? E.VALUE : s);
  }, { list: true });
  def('TEXTJOIN', 3, 252, 'vbr', (ctx, args) => {
    let delim = args[0], ign = args[1];
    if (isErr(delim)) return delim;
    const ds = delim instanceof Arr ? delim.d.flat().map(C.toStr) : [C.toStr(delim === MISSING ? '' : delim)];
    ign = C.toBool(ign === MISSING ? true : ign);
    if (isErr(ign)) return ign;
    const parts = [];
    const e = C.each(args.slice(2), (v) => { if (isErr(v)) return v; const s = C.toStr(v); if (ign && s === '') return; parts.push(s); }, { keepEmpty: true });
    if (e) return e;
    if (!ign) {
      /* blanks inside references count when empty strings are kept */
    }
    let out = '';
    parts.forEach((p, i) => { out += (i ? ds[(i - 1) % ds.length] : '') + p; });
    return out.length > 32767 ? E.VALUE : out;
  }, { list: true });
  /* T and N read the first cell of a range (no implicit intersection) */
  const firstOf = (v, ctx) => (v instanceof Ref ? C.cellValue(v.sheet, v.r1, v.c1) : v instanceof Arr ? (ctx.array ? v : v.d[0][0]) : v);
  def('T', 1, 1, 'r', (ctx, v) => { v = firstOf(v, ctx); if (v instanceof Arr) return C.withTail(new Arr(v.d.map((r) => r.map((x) => (isErr(x) ? x : typeof x === 'string' ? x : '')))), v.vr, ''); if (isErr(v)) return v; return typeof v === 'string' ? v : ''; }, { errOk: true });
  def('N', 1, 1, 'r', (ctx, v) => { v = firstOf(v, ctx); if (v instanceof Arr) return new Arr(v.d.map((r) => r.map((x) => (isErr(x) ? x : typeof x === 'number' ? x : x === true ? 1 : 0)))); if (isErr(v)) return v; return typeof v === 'number' ? v : typeof v === 'boolean' ? (v ? 1 : 0) : 0; }, { errOk: true });
  def('VALUE', 1, 1, 'v', (ctx, v) => {
    if (typeof v === 'number') return v;
    if (v == null || v === MISSING) return 0;
    if (typeof v === 'boolean') return E.VALUE;
    const s = String(v).trim();
    if (s === '') return String(v) === '' ? E.VALUE : 0;
    const n = C.parseNum(s, ctx.wb);
    return n == null ? E.VALUE : n;
  });
  def('NUMBERVALUE', 1, 3, 'sss', (ctx, s, dec, grp) => {
    dec = opt(dec, '.'); grp = opt(grp, ',');
    if (!dec || !grp && grp !== '') return E.VALUE;
    dec = dec[0]; grp = grp ? grp[0] : '';
    if (dec === grp) return E.VALUE;
    let t = s.replace(/\s+/g, '');
    if (t === '') return 0;
    let pct = 0;
    while (t.endsWith('%')) { pct++; t = t.slice(0, -1); }
    const di = t.indexOf(dec);
    if (di >= 0 && t.indexOf(grp, di + 1) >= 0 && grp) return E.VALUE;
    if (grp) t = t.split(grp).join('');
    t = t.replace(dec, '.');
    if (!/^[+-]?(\d*\.?\d*)(e[+-]?\d+)?$/i.test(t) || !/\d/.test(t)) return E.VALUE;
    return parseFloat(t) / Math.pow(100, pct);
  });
  def('TEXT', 2, 2, 'vs', (ctx, v, fmt) => {
    if (fmt === '' || fmt === MISSING) return typeof v === 'string' ? v : '';
    if (v == null || v === MISSING) v = 0;
    if (typeof v === 'string') { const n = C.parseNum(v, ctx.wb); if (n != null) v = n; }
    return NF.text(fmt, v, { date1904: ctx.wb.date1904, fn: true });
  });
  def('FIXED', 1, 3, 'nnb', (ctx, n, d, noComma) => {
    d = trunc(opt(d, 2)); noComma = opt(noComma, false);
    if (d > 127) return E.VALUE;
    const x = C.round(n, d);
    const code = (noComma ? '0' : '#,##0') + (d > 0 ? '.' + '0'.repeat(d) : '');
    return NF.text(code, x);
  });
  def('DOLLAR', 1, 2, 'nn', (ctx, n, d) => {
    d = trunc(opt(d, 2));
    const x = C.round(n, d);
    const dp = d > 0 ? '.' + '0'.repeat(d) : '';
    return NF.text('$#,##0' + dp + '_);($#,##0' + dp + ')', x).trim();
  });
  def('TEXTBEFORE', 2, 6, 'svnnnv', (ctx, s, delim, inst, mode, matchEnd, nf) => textSplitPart(ctx, s, delim, inst, mode, matchEnd, nf, true));
  def('TEXTAFTER', 2, 6, 'svnnnv', (ctx, s, delim, inst, mode, matchEnd, nf) => textSplitPart(ctx, s, delim, inst, mode, matchEnd, nf, false));
  function textSplitPart(ctx, s, delim, inst, mode, matchEnd, nf, before) {
    inst = trunc(opt(inst, 1)); mode = opt(mode, 0); matchEnd = opt(matchEnd, 0);
    const ds = (delim instanceof Arr ? delim.d.flat() : [delim]).map(C.toStr);
    if (inst === 0 || Math.abs(inst) > s.length + 1) return E.VALUE;
    const ci = mode === 1;
    const hay = ci ? s.toLowerCase() : s;
    const positions = [];
    for (let i = 0; i <= hay.length; i++) {
      for (const d of ds) {
        const dd = ci ? d.toLowerCase() : d;
        if (dd === '' ? true : hay.startsWith(dd, i)) { positions.push([i, dd.length]); break; }
      }
    }
    let hit;
    if (inst > 0) { hit = positions[inst - 1]; if (!hit && matchEnd && inst === positions.length + 1) hit = [s.length, 0]; }
    else { hit = positions[positions.length + inst]; if (!hit && matchEnd && -inst === positions.length + 1) hit = [0, 0]; }
    if (!hit) return nf === MISSING || nf === undefined ? E.NA : nf;
    return before ? s.slice(0, hit[0]) : s.slice(hit[0] + hit[1]);
  }
  def('TEXTSPLIT', 2, 6, 'svvbnv', (ctx, s, cd, rd, ign, mode, pad) => {
    ign = opt(ign, false); mode = opt(mode, 0); pad = opt(pad, E.NA);
    const list = (x) => (x === MISSING || x == null ? [] : (x instanceof Arr ? x.d.flat() : [x]).map(C.toStr).filter((d) => d !== ''));
    const cds = list(cd), rds = list(rd);
    if (!cds.length && !rds.length) return E.VALUE;
    const esc = (d) => d.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const split = (str, ds) => (ds.length ? str.split(new RegExp(ds.map(esc).join('|'), mode === 1 ? 'i' : '')) : [str]);
    let rows = split(s, rds).map((r) => split(r, cds));
    if (ign) rows = rows.map((r) => r.filter((x) => x !== '')).filter((r) => r.length);
    if (!rows.length) return E.CALC;
    const w = Math.max(...rows.map((r) => r.length));
    return new Arr(rows.map((r) => { const o = r.slice(); while (o.length < w) o.push(pad); return o; }));
  });
  def('VALUETOTEXT', 1, 2, 'vn', (ctx, v, f) => { f = opt(f, 0); if (isErr(v)) return v.e; if (typeof v === 'string') return f === 1 ? '"' + v.replace(/"/g, '""') + '"' : v; return C.toStr(v); }, { errOk: true });
  def('ARRAYTOTEXT', 1, 2, 'an', (ctx, a, f) => {
    f = opt(f, 0);
    const A = C.toArr(a);
    const t = (v) => (isErr(v) ? v.e : typeof v === 'string' ? (f === 1 ? '"' + v.replace(/"/g, '""') + '"' : v) : C.toStr(v));
    if (f === 1) return '{' + A.d.map((r) => r.map(t).join(',')).join(';') + '}';
    return A.d.flat().map(t).join(', ');
  });
  def('ASC', 1, 1, 's', (ctx, s) => s.replace(/[！-～]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0)).replace(/　/g, ' '));
  def('DBCS', 1, 1, 's', (ctx, s) => s.replace(/[!-~]/g, (ch) => String.fromCharCode(ch.charCodeAt(0) + 0xFEE0)).replace(/ /g, '　'));
  C.alias('JIS', 'DBCS');
  def('PHONETIC', 1, 1, 'r', (ctx, v) => C.toStr(C.scalar(v, ctx)));
  def('BAHTTEXT', 1, 1, 'n', (ctx, n) => bahtText(n));
  function bahtText(n) {
    const digits = ['', 'หนึ่ง', 'สอง', 'สาม', 'สี่', 'ห้า', 'หก', 'เจ็ด', 'แปด', 'เก้า'];
    const units = ['', 'สิบ', 'ร้อย', 'พัน', 'หมื่น', 'แสน'];
    const say = (x) => {
      if (x === 0) return '';
      let s = '';
      const str = String(x);
      if (str.length > 6) return say(Math.floor(x / 1e6)) + 'ล้าน' + say(x % 1e6);
      for (let i = 0; i < str.length; i++) {
        const d = +str[i], pos = str.length - i - 1;
        if (!d) continue;
        if (pos === 1 && d === 1) s += 'สิบ';
        else if (pos === 1 && d === 2) s += 'ยี่สิบ';
        else if (pos === 0 && d === 1 && str.length > 1) s += 'เอ็ด';
        else s += digits[d] + units[pos];
      }
      return s;
    };
    const neg = n < 0;
    const v = C.round(Math.abs(n), 2);
    const baht = Math.floor(v), satang = Math.round((v - baht) * 100);
    let out = (baht ? say(baht) + 'บาท' : '') + (satang ? say(satang) + 'สตางค์' : 'ถ้วน');
    if (!baht && !satang) out = 'ศูนย์บาทถ้วน';
    return (neg ? 'ลบ' : '') + out;
  }

  /* ------------------------------------------------------------ information */
  const typeOf = (v) => (v == null ? 1 : typeof v === 'number' ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : isErr(v) ? 16 : v instanceof Arr ? 64 : 1);
  const raw1 = (ctx, v) => (v instanceof Ref ? (v.single ? C.cellValue(v.sheet, v.r1, v.c1) : ctx.array ? C.refToArr(v) : C.intersect(v, ctx)) : v);
  const isFn = (name, test) => def(name, 1, 1, 'r', (ctx, v) => {
    const x = raw1(ctx, v);
    if (x instanceof Arr) return new Arr(x.d.map((r) => r.map((e) => test(e))));
    return test(x === MISSING ? null : x);
  }, { errOk: true });
  isFn('ISBLANK', (v) => v == null);
  isFn('ISERR', (v) => isErr(v) && v !== E.NA);
  isFn('ISERROR', (v) => isErr(v));
  isFn('ISNA', (v) => v === E.NA);
  isFn('ISTEXT', (v) => typeof v === 'string');
  isFn('ISNONTEXT', (v) => typeof v !== 'string');
  isFn('ISNUMBER', (v) => typeof v === 'number');
  isFn('ISLOGICAL', (v) => typeof v === 'boolean');
  def('ISREF', 1, 1, 'r', (ctx, v) => v instanceof Ref || v instanceof RefList || v instanceof C.Ref3D, { errOk: true });
  def('ISEVEN', 1, 1, 'n', (ctx, n) => trunc(n) % 2 === 0);
  def('ISODD', 1, 1, 'n', (ctx, n) => Math.abs(trunc(n) % 2) === 1);
  def('ISFORMULA', 1, 1, 'r', (ctx, v) => { if (!(v instanceof Ref)) return isErr(v) ? v : E.VALUE; const c = v.sheet.get(v.r1, v.c1); if (!c) return false; if (c.f != null) return true; if (c.am) { const m = v.sheet.get(c.am.r, c.am.c); return !!(m && !m.dyn); } return false; }, { errOk: true });
  def('ISOMITTED', 1, 1, 'r', (ctx, v) => v === MISSING);
  def('NA', 0, 0, '', () => E.NA);
  def('TYPE', 1, 1, 'r', (ctx, v) => { if (v instanceof Arr || (v instanceof Ref && !v.single)) return 64; return typeOf(raw1(ctx, v)); }, { errOk: true });
  def('ERROR.TYPE', 1, 1, 'v', (ctx, v) => { const codes = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA']; if (!isErr(v)) return E.NA; const i = codes.indexOf(v.e); return i < 0 ? (v.e === '#SPILL!' ? 9 : v.e === '#CALC!' ? 14 : E.NA) : i + 1; }, { errOk: true });
  def('SHEET', 0, 1, 'r', (ctx, v) => {
    if (v === undefined || v === MISSING) return ctx.wb.sheetIndex(ctx.sh) + 1;
    if (v instanceof Ref) return ctx.wb.sheetIndex(v.sheet) + 1;
    if (typeof v === 'string') { const s = ctx.wb.sheetByName(v); return s ? ctx.wb.sheetIndex(s) + 1 : E.NA; }
    return E.VALUE;
  });
  def('SHEETS', 0, 1, 'r', (ctx, v) => { if (v === undefined || v === MISSING) return ctx.wb.sheets.length; if (v instanceof C.Ref3D) return v.list.length; if (v instanceof Ref) return 1; return E.VALUE; });
  def('CELL', 1, 2, 'sr', (ctx, what, ref) => {
    if (isErr(what)) return what;
    const r = ref === MISSING || ref === undefined ? new Ref(ctx.sh, ctx.r, ctx.c) : ref;
    if (!(r instanceof Ref)) return E.VALUE;
    const sh = r.sheet, cell = sh.get(r.r1, r.c1);
    const st = sh.wb.styles.get(sh.styleAt(r.r1, r.c1));
    switch (what.toLowerCase()) {
      case 'address': return '$' + F.colName(r.c1) + '$' + (r.r1 + 1) + (sh !== ctx.sh ? '' : '');
      case 'col': return r.c1 + 1;
      case 'row': return r.r1 + 1;
      case 'contents': { const v = C.cellValue(sh, r.r1, r.c1); return v == null ? 0 : v; }
      case 'filename': return sh.wb.fileName ? (sh.wb.filePath || '') + '[' + sh.wb.fileName + ']' + sh.name : '';
      case 'sheetname': return sh.name;
      case 'type': { const v = cell ? cell.v : null; return v == null ? 'b' : typeof v === 'string' ? 'l' : 'v'; }
      case 'width': return Math.round(M.pxToChars(M.colPx(sh, r.c1), M.mdw(sh.wb)));
      case 'protect': return st.prot && st.prot.locked === false ? 0 : 1;
      case 'prefix': { const a = st.align && st.align.h; const v = cell && cell.v; if (typeof v !== 'string') return ''; return a === 'right' ? '"' : a === 'center' ? '^' : a === 'fill' ? '\\' : "'"; }
      case 'color': return /;\[Red\]|;\[Color/i.test(st.nf || '') ? 1 : 0;
      case 'parentheses': return /^[^;]*\(/.test(st.nf || '') ? 1 : 0;
      case 'format': return cellFormatCode(st.nf || 'General');
      default: return E.VALUE;
    }
  }, { volatile: true });
  function cellFormatCode(code) {
    const cat = NF.category(code);
    if (code === 'General') return 'G';
    const neg = /;\[Red\]/i.test(code) ? '-' : '';
    const dec = NF.decimals(code);
    if (cat === 'Percentage') return 'P' + dec;
    if (cat === 'Scientific') return 'S' + dec;
    if (cat === 'Currency' || cat === 'Accounting') return 'C' + dec + neg;
    if (cat === 'Date' || cat === 'Time') {
      if (/^m+\/d+\/y+/i.test(code)) return 'D4';
      if (/^d+-m+-y+/i.test(code)) return 'D1';
      if (/^d+-m+$/i.test(code)) return 'D2';
      if (/^m+-y+$/i.test(code)) return 'D3';
      if (/h:mm:ss AM/i.test(code)) return 'D6';
      if (/h:mm AM/i.test(code)) return 'D7';
      if (/h:mm:ss/i.test(code)) return 'D8';
      if (/h:mm/i.test(code)) return 'D9';
      return 'D4';
    }
    if (/#,##0/.test(code)) return ',' + dec + neg;
    return 'F' + dec + neg;
  }
  def('INFO', 1, 1, 's', (ctx, t) => {
    switch (t.toLowerCase()) {
      case 'directory': return '/';
      case 'numfile': return ctx.wb.sheets.length;
      case 'origin': return '$A:$A$1';
      case 'osversion': return 'Web';
      case 'recalc': return 'Automatic';
      case 'release': return '11.0';
      case 'system': return 'pcdos';
      default: return E.VALUE;
    }
  }, { volatile: true });

  /* ------------------------------------------------------------------ dates */
  const parts = (v, ctx) => NF.serialToParts(v, ctx.wb.date1904, 0);
  const D2S = (y, m, d, ctx) => NF.dateToSerial(y, m, d, ctx.wb.date1904);
  const checkDate = (v) => v < 0 || v >= 2958466;
  def('DATE', 3, 3, 'nnn', (ctx, y, m, d) => {
    y = trunc(y); m = trunc(m); d = trunc(d);
    if (y < 0 || y >= 10000) return E.NUM;
    if (y < 1900) y += 1900;
    /* normalise month overflow */
    y += Math.floor((m - 1) / 12);
    m = ((((m - 1) % 12) + 12) % 12) + 1;
    const s = NF.dateToSerial(y, m, 1, ctx.wb.date1904) + d - 1;
    if (s < 0 || s >= 2958466) return E.NUM;
    return s;
  });
  def('TIME', 3, 3, 'nnn', (ctx, h, m, s) => {
    h = trunc(h); m = trunc(m); s = trunc(s);
    const t = h * 3600 + m * 60 + s;
    if (t < 0 || h > 32767 || m > 32767 || s > 32767) return E.NUM;
    return (t % 86400) / 86400;
  });
  const dpart = (name, f) => def(name, 1, 1, 'n', (ctx, v) => { if (checkDate(v)) return E.NUM; return f(parts(v, ctx)); });
  dpart('YEAR', (p) => p.y); dpart('MONTH', (p) => p.m); dpart('DAY', (p) => p.d);
  dpart('HOUR', (p) => p.H); dpart('MINUTE', (p) => p.M); dpart('SECOND', (p) => p.S);
  def('WEEKDAY', 1, 2, 'nn', (ctx, v, type) => {
    if (checkDate(v)) return E.NUM;
    type = trunc(opt(type, 1));
    const wd = parts(v, ctx).wd; /* 0 Sunday */
    switch (type) {
      case 1: case 17: return wd + 1;
      case 2: case 11: return ((wd + 6) % 7) + 1;
      case 3: return (wd + 6) % 7;
      case 12: return ((wd + 5) % 7) + 1;
      case 13: return ((wd + 4) % 7) + 1;
      case 14: return ((wd + 3) % 7) + 1;
      case 15: return ((wd + 2) % 7) + 1;
      case 16: return ((wd + 1) % 7) + 1;
      default: return E.NUM;
    }
  });
  /* ISO week in serial-number space, so the 1900 calendar quirk (1 Jan 1900 = Sunday) matches Excel */
  const isoWeek = (p, ctx) => {
    const s = Math.floor(D2S(p.y, p.m, p.d, ctx));
    const d1904 = ctx.wb.date1904;
    const wd = d1904 ? (s + 4) % 7 : ((s + 5) % 7 + 7) % 7; /* Monday 0 */
    const thu = s - wd + 3;
    const yearOf = (x) => (x < (d1904 ? 0 : 1) ? (d1904 ? 1903 : 1899) : NF.serialToParts(x, d1904).y);
    const jan1 = (y) => (y < (d1904 ? 1904 : 1900) ? (d1904 ? -365 : -364) : Math.floor(D2S(y, 1, 1, ctx)));
    return Math.floor((thu - jan1(yearOf(thu))) / 7) + 1;
  };
  def('ISOWEEKNUM', 1, 1, 'n', (ctx, v) => (checkDate(v) ? E.NUM : isoWeek(parts(v, ctx), ctx)));
  def('WEEKNUM', 1, 2, 'nn', (ctx, v, type) => {
    if (checkDate(v)) return E.NUM;
    type = trunc(opt(type, 1));
    const p = parts(v, ctx);
    if (type === 21) return isoWeek(p, ctx);
    const startMap = { 1: 0, 17: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6 };
    const start = startMap[type];
    if (start === undefined) return E.NUM;
    const jan1 = Math.floor(D2S(p.y, 1, 1, ctx));
    const jan1wd = NF.serialToParts(jan1, ctx.wb.date1904).wd;
    const offset = (jan1wd - start + 7) % 7;
    return Math.floor((Math.floor(v) - jan1 + offset) / 7) + 1;
  });
  def('DATEVALUE', 1, 1, 's', (ctx, s) => {
    const r = NF.parseInput(s, { date1904: ctx.wb.date1904 });
    if (typeof r.v !== 'number' || !r.fmt || !/[dmy]/i.test(r.fmt)) return E.VALUE;
    return Math.floor(r.v);
  });
  def('TIMEVALUE', 1, 1, 's', (ctx, s) => {
    const r = NF.parseInput(s, { date1904: ctx.wb.date1904 });
    if (typeof r.v !== 'number' || !r.fmt || !/[hs]/i.test(r.fmt)) return E.VALUE;
    return r.v - Math.floor(r.v);
  });
  def('TODAY', 0, 0, '', (ctx) => Math.floor(NF.jsDateToSerial(new Date(), ctx.wb.date1904)), { volatile: true });
  def('NOW', 0, 0, '', (ctx) => NF.jsDateToSerial(new Date(), ctx.wb.date1904), { volatile: true });
  const addMonths = (v, n, ctx, eom) => {
    const p = parts(v, ctx);
    let y = p.y, m = p.m + n;
    y += Math.floor((m - 1) / 12);
    m = ((((m - 1) % 12) + 12) % 12) + 1;
    const last = new Date(Date.UTC(y, m, 0)).getUTCDate();
    const d = eom ? last : Math.min(p.d, last);
    return D2S(y, m, d, ctx);
  };
  def('EDATE', 2, 2, 'nn', (ctx, v, n) => { if (checkDate(v)) return E.NUM; const r = addMonths(Math.floor(v), trunc(n), ctx, false); return checkDate(r) ? E.NUM : r; });
  def('EOMONTH', 2, 2, 'nn', (ctx, v, n) => { if (checkDate(v)) return E.NUM; const r = addMonths(Math.floor(v), trunc(n), ctx, true); return checkDate(r) ? E.NUM : r; });
  def('DAYS', 2, 2, 'vv', (ctx, e, s) => {
    const conv = (x) => (typeof x === 'string' ? C.parseNum(x, ctx.wb) ?? E.VALUE : C.toNum(x, ctx));
    e = conv(e); s = conv(s);
    if (isErr(e)) return e; if (isErr(s)) return s;
    return Math.floor(e) - Math.floor(s);
  });
  C.days360 = (s, e, method, ctx) => {
    const a = parts(s, ctx), b = parts(e, ctx);
    let d1 = a.d, d2 = b.d;
    if (method) { if (d1 === 31) d1 = 30; if (d2 === 31) d2 = 30; }
    else {
      const lastFeb = (p) => p.m === 2 && p.d === new Date(Date.UTC(p.y, 2, 0)).getUTCDate();
      if (d1 === 31 || lastFeb(a)) d1 = 30;
      if (d2 === 31 && d1 >= 30) d2 = 30;
    }
    return (b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1);
  };
  def('DAYS360', 2, 3, 'nnb', (ctx, s, e, m) => C.days360(Math.floor(s), Math.floor(e), opt(m, false), ctx));
  /** weekend mask: 7 flags Monday..Sunday */
  function weekendMask(w) {
    if (w === MISSING || w === undefined) return [0, 0, 0, 0, 0, 1, 1];
    if (typeof w === 'string') { if (!/^[01]{7}$/.test(w) || w === '1111111') return E.VALUE; return w.split('').map(Number); }
    const n = trunc(C.toNum(w));
    const map = { 1: [5, 6], 2: [6, 0], 3: [0, 1], 4: [1, 2], 5: [2, 3], 6: [3, 4], 7: [4, 5], 11: [6], 12: [0], 13: [1], 14: [2], 15: [3], 16: [4], 17: [5] };
    if (!map[n]) return E.NUM;
    const m = [0, 0, 0, 0, 0, 0, 0];
    for (const i of map[n]) m[i] = 1;
    return m;
  }
  const holidaySet = (h, ctx) => {
    const set = new Set();
    if (h === MISSING || h === undefined) return set;
    const e = C.each([h], (v) => { if (isErr(v)) return v; const n = typeof v === 'number' ? v : C.parseNum(String(v), ctx.wb); if (n == null) return E.VALUE; set.add(Math.floor(n)); });
    return e || set;
  };
  const isWork = (s, mask, hol, ctx) => { const wd = NF.serialToParts(s, ctx.wb.date1904).wd; return !mask[(wd + 6) % 7] && !hol.has(s); };
  function networkdays(ctx, s, e, mask, hol) {
    s = Math.floor(s); e = Math.floor(e);
    const sign = s <= e ? 1 : -1;
    if (sign < 0) [s, e] = [e, s];
    /* whole weeks quickly */
    const workPerWeek = 7 - mask.reduce((a, b) => a + b, 0);
    let n = 0;
    const span = e - s + 1;
    const weeks = Math.floor(span / 7);
    n += weeks * workPerWeek;
    for (let d = s + weeks * 7; d <= e; d++) if (isWork(d, mask, new Set(), ctx)) n++;
    for (const h of hol) if (h >= s && h <= e && !mask[(NF.serialToParts(h, ctx.wb.date1904).wd + 6) % 7]) n--;
    return n * sign;
  }
  def('NETWORKDAYS', 2, 3, 'nnr', (ctx, s, e, h) => { const hol = holidaySet(h, ctx); if (isErr(hol)) return hol; return networkdays(ctx, s, e, [0, 0, 0, 0, 0, 1, 1], hol); });
  def('NETWORKDAYS.INTL', 2, 4, 'nnvr', (ctx, s, e, w, h) => { const mask = weekendMask(w); if (isErr(mask)) return mask; const hol = holidaySet(h, ctx); if (isErr(hol)) return hol; return networkdays(ctx, s, e, mask, hol); });
  function workday(ctx, s, n, mask, hol) {
    s = Math.floor(s); n = trunc(n);
    if (mask.every((x) => x)) return E.VALUE;
    const step = n >= 0 ? 1 : -1;
    let d = s;
    let k = Math.abs(n);
    while (k > 0) { d += step; if (isWork(d, mask, hol, ctx)) k--; if (d < 0 || d > 2958465) return E.NUM; }
    return d;
  }
  def('WORKDAY', 2, 3, 'nnr', (ctx, s, n, h) => { const hol = holidaySet(h, ctx); if (isErr(hol)) return hol; return workday(ctx, s, n, [0, 0, 0, 0, 0, 1, 1], hol); });
  def('WORKDAY.INTL', 2, 4, 'nnvr', (ctx, s, n, w, h) => { const mask = weekendMask(w); if (isErr(mask)) return mask; const hol = holidaySet(h, ctx); if (isErr(hol)) return hol; return workday(ctx, s, n, mask, hol); });
  /** YEARFRAC basis 0..4 */
  C.yearfrac = function (s, e, basis, ctx) {
    s = Math.floor(s); e = Math.floor(e);
    if (s > e) [s, e] = [e, s];
    const a = parts(s, ctx), b = parts(e, ctx);
    const leap = (y) => (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
    switch (basis) {
      case 0: {
        let d1 = a.d, d2 = b.d;
        const lastFeb = (p) => p.m === 2 && p.d === new Date(Date.UTC(p.y, 2, 0)).getUTCDate();
        if (d1 === 31 && d2 === 31) { d1 = 30; d2 = 30; }
        else if (d1 === 31) d1 = 30;
        else if (d1 === 30 && d2 === 31) d2 = 30;
        else if (lastFeb(a) && lastFeb(b)) { d1 = 30; d2 = 30; }
        else if (lastFeb(a)) d1 = 30;
        return ((b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)) / 360;
      }
      case 1: {
        const diff = e - s;
        if (a.y === b.y || (b.y === a.y + 1 && (a.m > b.m || (a.m === b.m && a.d >= b.d)))) {
          let den;
          if (a.y === b.y) den = leap(a.y) ? 366 : 365;
          else {
            /* spans Feb 29? */
            const feb29 = (y) => leap(y) && D2S(y, 2, 29, ctx);
            let has = false;
            for (const y of [a.y, b.y]) { const f = feb29(y); if (f && f >= s && f <= e) has = true; }
            den = has ? 366 : 365;
          }
          return diff / den;
        }
        let total = 0;
        for (let y = a.y; y <= b.y; y++) total += leap(y) ? 366 : 365;
        return diff / (total / (b.y - a.y + 1));
      }
      case 2: return (e - s) / 360;
      case 3: return (e - s) / 365;
      case 4: { let d1 = Math.min(a.d, 30), d2 = Math.min(b.d, 30); return ((b.y - a.y) * 360 + (b.m - a.m) * 30 + (d2 - d1)) / 360; }
      default: return E.NUM;
    }
  };
  def('YEARFRAC', 2, 3, 'nnn', (ctx, s, e, b) => { b = trunc(opt(b, 0)); if (b < 0 || b > 4 || s < 0 || e < 0) return E.NUM; return C.yearfrac(s, e, b, ctx); });
  def('DATEDIF', 3, 3, 'nns', (ctx, s, e, u) => {
    s = Math.floor(s); e = Math.floor(e);
    if (s > e) return E.NUM;
    const a = parts(s, ctx), b = parts(e, ctx);
    const months = (b.y - a.y) * 12 + (b.m - a.m) - (b.d < a.d ? 1 : 0);
    switch (u.toUpperCase()) {
      case 'D': return e - s;
      case 'M': return months;
      case 'Y': return Math.floor(months / 12);
      case 'MD': { let d = b.d - a.d; if (d < 0) { const pm = new Date(Date.UTC(b.y, b.m - 1, 0)).getUTCDate(); d += pm; } return d; }
      case 'YM': return months % 12;
      case 'YD': {
        let y = b.y;
        let st = D2S(y, a.m, a.d, ctx);
        if (st > e) st = D2S(y - 1, a.m, a.d, ctx);
        return e - st;
      }
      default: return E.NUM;
    }
  });
})(typeof window !== 'undefined' ? window : globalThis);
