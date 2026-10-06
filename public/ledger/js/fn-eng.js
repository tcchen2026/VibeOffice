/* Ledger — engineering (bases, bits, complex numbers, Bessel, CONVERT) and database functions. */
(function (root) {
  'use strict';
  const L = root.L;
  const C = L.calc, M = L.model, F = L.formula;
  const E = M.ERR;
  const { def, isErr, MISSING, Arr, Ref } = C;
  const opt = C.opt;
  const trunc = Math.trunc;
  const S = C.special;

  /* ------------------------------------------------------------ number bases */
  const BITS = { 2: 10, 8: 30, 16: 40 };
  /** text in base `b` (10 chars max, two's complement) → number */
  function fromBase(s, b) {
    if (typeof s === 'number') s = String(trunc(s));
    s = String(s).trim().toUpperCase();
    if (s.length > 10) return E.NUM;
    if (!s) return 0;
    for (const ch of s) { const d = parseInt(ch, 16); if (isNaN(d) || d >= b) return E.NUM; }
    let v = parseInt(s, b);
    const bits = BITS[b];
    if (s.length === 10 && v >= Math.pow(2, bits - 1)) v -= Math.pow(2, bits);
    return v;
  }
  function toBase(n, b, places) {
    const bits = BITS[b];
    const lim = Math.pow(2, bits - 1);
    if (isErr(n)) return n;
    n = trunc(n);
    if (n < -lim || n >= lim) return E.NUM;
    if (n < 0) return (n + Math.pow(2, bits)).toString(b).toUpperCase();
    let s = n.toString(b).toUpperCase();
    if (places !== MISSING && places !== undefined) {
      places = trunc(places);
      if (places < s.length || places > 10 || places < 0) return E.NUM;
      s = s.padStart(places, '0');
    }
    return s;
  }
  const baseArg = (v) => (typeof v === 'boolean' ? E.VALUE : v);
  for (const [a, ba] of [['BIN', 2], ['OCT', 8], ['HEX', 16]]) {
    def(a + '2DEC', 1, 1, 'v', (ctx, s) => { s = baseArg(s); if (isErr(s)) return s; return fromBase(s, ba); });
    for (const [b, bb] of [['BIN', 2], ['OCT', 8], ['HEX', 16]]) {
      if (a === b) continue;
      def(a + '2' + b, 1, 2, 'vn', (ctx, s, p) => { s = baseArg(s); if (isErr(s)) return s; const v = fromBase(s, ba); if (isErr(v)) return v; return toBase(v, bb, p); });
    }
  }
  def('DEC2BIN', 1, 2, 'nn', (ctx, n, p) => toBase(n, 2, p));
  def('DEC2OCT', 1, 2, 'nn', (ctx, n, p) => toBase(n, 8, p));
  def('DEC2HEX', 1, 2, 'nn', (ctx, n, p) => toBase(n, 16, p));
  const bitArg = (n) => (n < 0 || n >= 281474976710656 || n !== Math.floor(n) ? null : n);
  const big = (n) => BigInt(n);
  def('BITAND', 2, 2, 'nn', (ctx, a, b) => { a = bitArg(a); b = bitArg(b); if (a == null || b == null) return E.NUM; return Number(big(a) & big(b)); });
  def('BITOR', 2, 2, 'nn', (ctx, a, b) => { a = bitArg(a); b = bitArg(b); if (a == null || b == null) return E.NUM; return Number(big(a) | big(b)); });
  def('BITXOR', 2, 2, 'nn', (ctx, a, b) => { a = bitArg(a); b = bitArg(b); if (a == null || b == null) return E.NUM; return Number(big(a) ^ big(b)); });
  def('BITLSHIFT', 2, 2, 'nn', (ctx, a, s) => { a = bitArg(a); s = trunc(s); if (a == null || Math.abs(s) > 53) return E.NUM; const r = s >= 0 ? big(a) << big(s) : big(a) >> big(-s); if (r >= 281474976710656n) return E.NUM; return Number(r); });
  def('BITRSHIFT', 2, 2, 'nn', (ctx, a, s) => { a = bitArg(a); s = trunc(s); if (a == null || Math.abs(s) > 53) return E.NUM; const r = s >= 0 ? big(a) >> big(s) : big(a) << big(-s); if (r >= 281474976710656n) return E.NUM; return Number(r); });
  def('DELTA', 1, 2, 'nn', (ctx, a, b) => (a === opt(b, 0) ? 1 : 0));
  def('GESTEP', 1, 2, 'nn', (ctx, a, b) => (a >= opt(b, 0) ? 1 : 0));
  def('ERF', 1, 2, 'nn', (ctx, a, b) => (b === MISSING || b === undefined ? S.erf(a) : S.erf(b) - S.erf(a)));
  def('ERF.PRECISE', 1, 1, 'n', (ctx, a) => S.erf(a));
  def('ERFC', 1, 1, 'n', (ctx, a) => S.erfc(a));
  C.alias('ERFC.PRECISE', 'ERFC');

  /* --------------------------------------------------------------- complex */
  function parseCx(v) {
    if (typeof v === 'number') return { re: v, im: 0, sfx: 'i' };
    if (typeof v === 'boolean') return E.VALUE;
    let s = String(v == null ? '' : v).trim();
    if (s === '') return { re: 0, im: 0, sfx: 'i' };
    const m = /^([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)?(?:([+-]?(?:\d+\.?\d*|\.\d+)?(?:[eE][+-]?\d+)?)([ij]))?$/.exec(s);
    if (!m || (!m[1] && !m[3])) return E.NUM;
    let re = 0, im = 0;
    if (m[3]) {
      re = m[1] ? parseFloat(m[1]) : 0;
      const t = m[2];
      im = t === '' || t === '+' ? 1 : t === '-' ? -1 : parseFloat(t);
      if (!m[1] && m[2] === undefined) im = 1;
    } else re = parseFloat(m[1]);
    if (isNaN(re) || isNaN(im)) return E.NUM;
    return { re, im, sfx: m[3] || 'i' };
  }
  const n15 = (x) => C.numToStr(L.numfmt.r15(x));
  function fmtCx(z, sfx) {
    const re = z.re, im = z.im;
    if (!isFinite(re) || !isFinite(im)) return E.NUM;
    sfx = sfx || 'i';
    if (im === 0) return n15(re);
    const ims = im === 1 ? '' : im === -1 ? '-' : n15(im);
    if (re === 0) return ims + sfx;
    return n15(re) + (im > 0 ? '+' : '') + ims + sfx;
  }
  const cx = {
    add: (a, b) => ({ re: a.re + b.re, im: a.im + b.im }),
    sub: (a, b) => ({ re: a.re - b.re, im: a.im - b.im }),
    mul: (a, b) => ({ re: a.re * b.re - a.im * b.im, im: a.re * b.im + a.im * b.re }),
    div: (a, b) => { const d = b.re * b.re + b.im * b.im; return { re: (a.re * b.re + a.im * b.im) / d, im: (a.im * b.re - a.re * b.im) / d }; },
    abs: (a) => Math.hypot(a.re, a.im),
    arg: (a) => Math.atan2(a.im, a.re),
    exp: (a) => ({ re: Math.exp(a.re) * Math.cos(a.im), im: Math.exp(a.re) * Math.sin(a.im) }),
    ln: (a) => ({ re: Math.log(Math.hypot(a.re, a.im)), im: Math.atan2(a.im, a.re) }),
    sin: (a) => ({ re: Math.sin(a.re) * Math.cosh(a.im), im: Math.cos(a.re) * Math.sinh(a.im) }),
    cos: (a) => ({ re: Math.cos(a.re) * Math.cosh(a.im), im: -Math.sin(a.re) * Math.sinh(a.im) }),
    sinh: (a) => ({ re: Math.sinh(a.re) * Math.cos(a.im), im: Math.cosh(a.re) * Math.sin(a.im) }),
    cosh: (a) => ({ re: Math.cosh(a.re) * Math.cos(a.im), im: Math.sinh(a.re) * Math.sin(a.im) }),
  };
  const one = { re: 1, im: 0 };
  def('COMPLEX', 2, 3, 'nns', (ctx, re, im, s) => { s = opt(s, 'i'); if (s !== 'i' && s !== 'j' && s !== '') return E.VALUE; return fmtCx({ re, im }, s || 'i'); });
  const cxFn = (name, f, scalar) => def(name, 1, 1, 'v', (ctx, v) => { const z = parseCx(v); if (isErr(z)) return z; const r = f(z); return scalar ? r : isErr(r) ? r : fmtCx(r, z.sfx); });
  cxFn('IMABS', (z) => cx.abs(z), true);
  cxFn('IMAGINARY', (z) => z.im, true);
  cxFn('IMREAL', (z) => z.re, true);
  cxFn('IMARGUMENT', (z) => (z.re === 0 && z.im === 0 ? E.DIV0 : cx.arg(z)), true);
  cxFn('IMCONJUGATE', (z) => ({ re: z.re, im: -z.im }));
  cxFn('IMEXP', cx.exp);
  cxFn('IMLN', (z) => (z.re === 0 && z.im === 0 ? E.NUM : cx.ln(z)));
  cxFn('IMLOG10', (z) => { if (z.re === 0 && z.im === 0) return E.NUM; const l = cx.ln(z); return { re: l.re / Math.LN10, im: l.im / Math.LN10 }; });
  cxFn('IMLOG2', (z) => { if (z.re === 0 && z.im === 0) return E.NUM; const l = cx.ln(z); return { re: l.re / Math.LN2, im: l.im / Math.LN2 }; });
  cxFn('IMSQRT', (z) => { const r = Math.sqrt(cx.abs(z)), t = cx.arg(z) / 2; return { re: r * Math.cos(t), im: r * Math.sin(t) }; });
  cxFn('IMSIN', cx.sin); cxFn('IMCOS', cx.cos); cxFn('IMSINH', cx.sinh); cxFn('IMCOSH', cx.cosh);
  cxFn('IMTAN', (z) => cx.div(cx.sin(z), cx.cos(z)));
  cxFn('IMCOT', (z) => cx.div(cx.cos(z), cx.sin(z)));
  cxFn('IMSEC', (z) => cx.div(one, cx.cos(z)));
  cxFn('IMCSC', (z) => cx.div(one, cx.sin(z)));
  cxFn('IMSECH', (z) => cx.div(one, cx.cosh(z)));
  cxFn('IMCSCH', (z) => cx.div(one, cx.sinh(z)));
  def('IMPOWER', 2, 2, 'vn', (ctx, v, n) => { const z = parseCx(v); if (isErr(z)) return z; const r = Math.pow(cx.abs(z), n), t = cx.arg(z) * n; if (z.re === 0 && z.im === 0 && n <= 0) return E.NUM; return fmtCx({ re: r * Math.cos(t), im: r * Math.sin(t) }, z.sfx); });
  def('IMDIV', 2, 2, 'vv', (ctx, a, b) => { const x = parseCx(a), y = parseCx(b); if (isErr(x)) return x; if (isErr(y)) return y; if (y.re === 0 && y.im === 0) return E.NUM; return fmtCx(cx.div(x, y), x.sfx); });
  def('IMSUB', 2, 2, 'vv', (ctx, a, b) => { const x = parseCx(a), y = parseCx(b); if (isErr(x)) return x; if (isErr(y)) return y; return fmtCx(cx.sub(x, y), x.sfx); });
  const cxAgg = (name, f, init) => def(name, 1, 255, 'r', (ctx, args) => {
    let acc = init, sfx = null;
    const e = C.each(args, (v) => { const z = parseCx(v); if (isErr(z)) return z; if (sfx && z.sfx !== sfx && (z.im !== 0)) return E.VALUE; if (z.im !== 0) sfx = z.sfx; acc = f(acc, z); });
    return e || fmtCx(acc, sfx || 'i');
  }, { list: true });
  cxAgg('IMSUM', cx.add, { re: 0, im: 0 });
  cxAgg('IMPRODUCT', cx.mul, { re: 1, im: 0 });

  /* --------------------------------------------------------------- Bessel */
  /* Integral representations evaluated to full double precision:
     periodic trapezoid rules (exponentially convergent) for J and I, composite Gauss–Legendre for Y,
     a doubly-exponential-decay trapezoid for K; power series where they are cancellation-free. */
  const GL = (function () {
    const n = 16, xs = [], ws = [];
    for (let i = 1; i <= n; i++) {
      let z = Math.cos((Math.PI * (i - 0.25)) / (n + 0.5)), pp = 0;
      for (let it = 0; it < 100; it++) {
        let p1 = 1, p2 = 0;
        for (let j = 1; j <= n; j++) { const p3 = p2; p2 = p1; p1 = ((2 * j - 1) * z * p2 - (j - 1) * p3) / j; }
        pp = (n * (z * p1 - p2)) / (z * z - 1);
        const z1 = z; z = z1 - p1 / pp;
        if (Math.abs(z - z1) < 1e-16) break;
      }
      xs.push(z); ws.push(2 / ((1 - z * z) * pp * pp));
    }
    return { xs, ws };
  })();
  function integrate(f, a, b, panels) {
    const h = (b - a) / panels;
    let sum = 0;
    for (let p = 0; p < panels; p++) {
      const m = a + (p + 0.5) * h, r = h / 2;
      let s = 0;
      for (let i = 0; i < GL.xs.length; i++) s += GL.ws[i] * f(m + r * GL.xs[i]);
      sum += s * r;
    }
    return sum;
  }
  function besselJ(x, n) {
    if (x === 0) return n === 0 ? 1 : 0;
    if ((x * x) / 4 <= 4 * (n + 1)) {
      /* power series, terms shrink from the start */
      let term = Math.pow(Math.abs(x) / 2, n) / C.fact(n), sum = 0;
      for (let k = 0; k < 300; k++) { sum += term; term *= -(x * x) / 4 / ((k + 1) * (k + 1 + n)); if (Math.abs(term) < 1e-17 * Math.abs(sum)) break; }
      return x < 0 && n % 2 ? -sum : sum;
    }
    const N = Math.min(4000000, 2 * Math.ceil(Math.abs(x) + n) + 96);
    let s = 0;
    for (let k = 0; k < N; k++) { const t = (2 * Math.PI * k) / N; s += Math.cos(n * t - x * Math.sin(t)); }
    return s / N;
  }
  function besselI(x, n) {
    if (x === 0) return n === 0 ? 1 : 0;
    const ax = Math.abs(x);
    let r;
    if (ax <= 60) {
      let term = Math.pow(ax / 2, n) / C.fact(n), sum = 0;
      for (let k = 0; k < 1000; k++) { sum += term; term *= (ax * ax) / 4 / ((k + 1) * (k + 1 + n)); if (term < 1e-17 * sum) break; }
      r = sum;
    } else {
      const N = Math.min(4000000, 2 * Math.ceil(ax + n) + 96);
      let s = 0;
      for (let k = 0; k < N; k++) { const t = (2 * Math.PI * k) / N; s += Math.exp(ax * (Math.cos(t) - 1)) * Math.cos(n * t); }
      r = (s / N) * Math.exp(ax);
    }
    return x < 0 && n % 2 ? -r : r;
  }
  function kInt(x, nu) {
    /* K_nu(x) = ∫0^∞ exp(-x cosh t) cosh(nu t) dt ; even integrand → trapezoid is spectrally accurate */
    const h = 0.05;
    let s = 0.5 * Math.exp(-x);
    for (let k = 1; k < 100000; k++) {
      const t = k * h;
      const v = Math.exp(-x * Math.cosh(t) + nu * t) * 0.5 * (1 + Math.exp(-2 * nu * t));
      s += v;
      if (v < 1e-18 * s && x * Math.cosh(t) > nu * t) break;
    }
    return s * h;
  }
  function besselK(x, n) {
    const k0 = kInt(x, 0), k1 = kInt(x, 1);
    if (n === 0) return k0;
    let bkm = k0, bk = k1;
    for (let j = 1; j < n; j++) { const bkp = bkm + ((2 * j) / x) * bk; bkm = bk; bk = bkp; }
    return bk;
  }
  function yN01(x, nu) {
    /* Y_nu(x) = 1/π ∫0^π sin(x sinθ − νθ) dθ − 1/π ∫0^∞ (e^{νt} + (−1)^ν e^{−νt}) e^{−x sinh t} dt */
    const a = integrate((th) => Math.sin(x * Math.sin(th) - nu * th), 0, Math.PI, Math.ceil(x / 2) + 8);
    const T = Math.asinh(45 / x) + 2;
    const b = integrate((t) => (Math.exp(nu * t) + (nu % 2 ? -1 : 1) * Math.exp(-nu * t)) * Math.exp(-x * Math.sinh(t)), 0, T, 48);
    return (a - b) / Math.PI;
  }
  function besselY(x, n) {
    const y0 = yN01(x, 0);
    if (n === 0) return y0;
    const y1 = yN01(x, 1);
    let bym = y0, by = y1;
    for (let j = 1; j < n; j++) { const byp = ((2 * j) / x) * by - bym; bym = by; by = byp; }
    return by;
  }
  def('BESSELJ', 2, 2, 'nn', (ctx, x, n) => { n = trunc(n); if (n < 0) return E.NUM; return besselJ(x, n); });
  def('BESSELY', 2, 2, 'nn', (ctx, x, n) => { n = trunc(n); if (n < 0 || x <= 0) return E.NUM; return besselY(x, n); });
  def('BESSELI', 2, 2, 'nn', (ctx, x, n) => { n = trunc(n); if (n < 0) return E.NUM; return besselI(x, n); });
  def('BESSELK', 2, 2, 'nn', (ctx, x, n) => { n = trunc(n); if (n < 0 || x <= 0) return E.NUM; return besselK(x, n); });

  /* --------------------------------------------------------------- CONVERT */
  /* units: [category, factor to base] — base units: g, m, s, Pa, N, J, W, T, K(special), m3, m2, bit, m/s */
  const U = {};
  const add = (cat, list, prefixable) => { for (const [k, f] of list) U[k] = { cat, f, prefixable: prefixable !== false && !(Array.isArray(prefixable) && !prefixable.includes(k)) }; };
  add('mass', [['g', 1], ['sg', 14593.90294], ['lbm', 453.59237], ['u', 1.66053906660e-24], ['ozm', 28.349523125], ['grain', 0.06479891], ['cwt', 45359.237], ['shweight', 45359.237], ['uk_cwt', 50802.34544], ['lcwt', 50802.34544], ['hweight', 50802.34544], ['stone', 6350.29318], ['ton', 907184.74], ['uk_ton', 1016046.9088], ['LTON', 1016046.9088], ['brton', 1016046.9088]]);
  add('dist', [['m', 1], ['mi', 1609.344], ['Nmi', 1852], ['in', 0.0254], ['ft', 0.3048], ['yd', 0.9144], ['ang', 1e-10], ['ell', 1.143], ['ly', 9460730472580800], ['parsec', 3.08567758149137e16], ['pc', 3.08567758149137e16], ['Pica', 0.0254 / 72], ['Picapt', 0.0254 / 72], ['pica', 0.0254 / 6], ['survey_mi', 1609.347218694437]]);
  add('time', [['yr', 31557600], ['day', 86400], ['d', 86400], ['hr', 3600], ['mn', 60], ['min', 60], ['sec', 1], ['s', 1]]);
  add('press', [['Pa', 1], ['p', 1], ['atm', 101325], ['at', 101325], ['mmHg', 133.322], ['psi', 6894.757293168361], ['Torr', 133.322368421]]);
  add('force', [['N', 1], ['dyn', 1e-5], ['dy', 1e-5], ['lbf', 4.4482216152605], ['pond', 0.00980665]]);
  add('energy', [['J', 1], ['e', 1e-7], ['c', 4.184], ['cal', 4.1868], ['eV', 1.602176634e-19], ['ev', 1.602176634e-19], ['HPh', 2684519.537696172792], ['hh', 2684519.537696172792], ['Wh', 3600], ['wh', 3600], ['flb', 1.3558179483314004], ['BTU', 1055.05585262], ['btu', 1055.05585262]]);
  add('power', [['HP', 745.69987158227022], ['h', 745.69987158227022], ['PS', 735.49875], ['W', 1], ['w', 1]]);
  add('mag', [['T', 1], ['ga', 1e-4]]);
  add('temp', [['C', 0], ['cel', 0], ['F', 0], ['fah', 0], ['K', 0], ['kel', 0], ['Rank', 0], ['Reau', 0]]);
  add('vol', [['tsp', 4.92892159375e-6], ['tspm', 5e-6], ['tbs', 1.478676478125e-5], ['oz', 2.95735295625e-5], ['cup', 2.365882365e-4], ['pt', 4.73176473e-4], ['us_pt', 4.73176473e-4], ['uk_pt', 5.6826125e-4], ['qt', 9.46352946e-4], ['uk_qt', 1.1365225e-3], ['gal', 3.785411784e-3], ['uk_gal', 4.54609e-3], ['l', 1e-3], ['L', 1e-3], ['lt', 1e-3], ['ang3', 1e-30], ['ang^3', 1e-30], ['barrel', 0.158987294928], ['bushel', 0.03523907016688], ['ft3', 0.028316846592], ['ft^3', 0.028316846592], ['in3', 1.6387064e-5], ['in^3', 1.6387064e-5], ['ly3', 8.46786664623715e47], ['ly^3', 8.46786664623715e47], ['m3', 1], ['m^3', 1], ['mi3', 4168181825.440579584], ['mi^3', 4168181825.440579584], ['yd3', 0.764554857984], ['yd^3', 0.764554857984], ['Nmi3', 6352182208], ['Nmi^3', 6352182208], ['Pica3', 4.39039566186557e-11], ['Pica^3', 4.39039566186557e-11], ['GRT', 2.8316846592], ['regton', 2.8316846592], ['MTON', 1.13267386368]]);
  add('area', [['uk_acre', 4046.8564224], ['us_acre', 4046.87260987425], ['ang2', 1e-20], ['ang^2', 1e-20], ['ar', 100], ['ft2', 0.09290304], ['ft^2', 0.09290304], ['ha', 10000], ['in2', 0.00064516], ['in^2', 0.00064516], ['ly2', 8.95054210748189e31], ['ly^2', 8.95054210748189e31], ['m2', 1], ['m^2', 1], ['Morgen', 2500], ['mi2', 2589988.110336], ['mi^2', 2589988.110336], ['Nmi2', 3429904], ['Nmi^2', 3429904], ['Pica2', 1.24452160493827e-7], ['Pica^2', 1.24452160493827e-7], ['yd2', 0.83612736], ['yd^2', 0.83612736]]);
  add('info', [['bit', 1], ['byte', 8]]);
  add('speed', [['admkn', 0.514773333333333], ['kn', 0.514444444444444], ['m/h', 1 / 3600], ['m/hr', 1 / 3600], ['m/s', 1], ['m/sec', 1], ['mph', 0.44704]]);
  const PFX = { Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 1e1, e: 1e1, d: 1e-1, c: 1e-2, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24 };
  const BPFX = { Yi: 2 ** 80, Zi: 2 ** 70, Ei: 2 ** 60, Pi: 2 ** 50, Ti: 2 ** 40, Gi: 2 ** 30, Mi: 2 ** 20, ki: 2 ** 10 };
  const NOPFX = new Set(['lbm', 'ozm', 'grain', 'cwt', 'shweight', 'uk_cwt', 'lcwt', 'hweight', 'stone', 'ton', 'uk_ton', 'LTON', 'brton', 'mi', 'Nmi', 'in', 'ft', 'yd', 'ell', 'Pica', 'Picapt', 'pica', 'survey_mi', 'yr', 'day', 'd', 'hr', 'mn', 'min', 'mmHg', 'psi', 'lbf', 'HPh', 'hh', 'flb', 'BTU', 'btu', 'HP', 'h', 'PS', 'F', 'fah', 'Rank', 'Reau', 'tsp', 'tspm', 'tbs', 'oz', 'cup', 'pt', 'us_pt', 'uk_pt', 'qt', 'uk_qt', 'gal', 'uk_gal', 'barrel', 'bushel', 'ft3', 'ft^3', 'in3', 'in^3', 'mi3', 'mi^3', 'yd3', 'yd^3', 'Nmi3', 'Nmi^3', 'uk_acre', 'us_acre', 'ft2', 'ft^2', 'ha', 'in2', 'in^2', 'Morgen', 'mi2', 'mi^2', 'Nmi2', 'Nmi^2', 'yd2', 'yd^2', 'admkn', 'kn', 'mph', 'GRT', 'regton', 'MTON']);
  function unit(s) {
    if (U[s]) return { u: U[s], k: s, mul: 1 };
    for (const [p, f] of Object.entries(BPFX)) if (s.startsWith(p) && U[s.slice(p.length)] && U[s.slice(p.length)].cat === 'info') return { u: U[s.slice(p.length)], k: s.slice(p.length), mul: f };
    for (const [p, f] of Object.entries(PFX).sort((a, b) => b[0].length - a[0].length)) {
      if (!s.startsWith(p)) continue;
      const base = s.slice(p.length);
      const u = U[base];
      if (!u || NOPFX.has(base)) continue;
      /* area/volume units with prefixes scale by the square/cube */
      const pw = /\^?2$/.test(base) && u.cat === 'area' ? 2 : /\^?3$/.test(base) && u.cat === 'vol' ? 3 : 1;
      return { u, k: base, mul: Math.pow(f, pw) };
    }
    return null;
  }
  const toK = (v, k) => ({ C: v + 273.15, cel: v + 273.15, F: (v - 32) * 5 / 9 + 273.15, fah: (v - 32) * 5 / 9 + 273.15, K: v, kel: v, Rank: (v * 5) / 9, Reau: v * 1.25 + 273.15 }[k]);
  const fromK = (v, k) => ({ C: v - 273.15, cel: v - 273.15, F: (v - 273.15) * 9 / 5 + 32, fah: (v - 273.15) * 9 / 5 + 32, K: v, kel: v, Rank: (v * 9) / 5, Reau: (v - 273.15) * 0.8 }[k]);
  def('CONVERT', 3, 3, 'nss', (ctx, v, from, to) => {
    const a = unit(from), b = unit(to);
    if (!a || !b || a.u.cat !== b.u.cat) return E.NA;
    if (a.u.cat === 'temp') return fromK(toK(v * a.mul, a.k), b.k) / b.mul;
    return (v * a.u.f * a.mul) / (b.u.f * b.mul);
  });

  /* --------------------------------------------------------------- database */
  /** records of a database range that satisfy a criteria range; returns {rows:[record row offsets], col} */
  function dbSelect(ctx, db, field, crit) {
    if (!(db instanceof Ref) || !(crit instanceof Ref)) return E.VALUE;
    const sh = db.sheet;
    const headers = [];
    for (let c = db.c1; c <= db.c2; c++) headers.push(C.toStr(C.cellValue(sh, db.r1, c)).toLowerCase());
    let col = null;
    if (field !== MISSING && field !== undefined) {
      if (typeof field === 'number') { col = trunc(field) - 1; if (col < 0 || col >= headers.length) return E.VALUE; }
      else { col = headers.indexOf(String(field).toLowerCase()); if (col < 0) return E.VALUE; }
    }
    /* criteria rows: OR across rows, AND across columns */
    const csh = crit.sheet;
    const cHead = [];
    for (let c = crit.c1; c <= crit.c2; c++) cHead.push(C.toStr(C.cellValue(csh, crit.r1, c)).toLowerCase());
    const rules = [];
    for (let r = crit.r1 + 1; r <= crit.r2; r++) {
      const row = [];
      for (let k = 0; k < cHead.length; k++) {
        const cc = crit.c1 + k;
        const cell = csh.get(r, cc);
        const v = C.cellValue(csh, r, cc);
        if (v == null || v === '') continue;
        const ci = headers.indexOf(cHead[k]);
        if (cell && cell.f != null && ci < 0) {
          /* computed criterion: a formula written for the first record */
          row.push({ formula: cell, at: { r, c: cc } });
          continue;
        }
        if (ci < 0) { row.push({ never: true }); continue; }
        let pred;
        if (typeof v === 'string' && !/^(<=|>=|<>|<|>|=)/.test(v)) pred = C.criteria(v + '*', ctx);
        else pred = C.criteria(v, ctx);
        row.push({ ci, pred });
      }
      rules.push(row);
    }
    const out = [];
    for (let r = db.r1 + 1; r <= db.r2; r++) {
      if (r > sh.maxR) break;
      const ok = !rules.length || rules.some((row) => row.every((t) => {
        if (t.never) return false;
        if (t.formula) {
          let ast;
          try { ast = F.shift(F.parse(t.formula.f), r - (db.r1 + 1), 0); } catch (e) { return false; }
          const v = C.scalarOut(C.ev(ast, C.ctx(ctx.wb, csh, t.at.r, t.at.c)), ctx);
          return C.toBool(v) === true;
        }
        return t.pred(C.cellValue(sh, r, db.c1 + t.ci));
      }));
      if (ok) out.push(r);
    }
    return { rows: out, col, sh, db };
  }
  const dbFn = (name, f, needField) => def(name, 3, 3, 'rvr', (ctx, db, field, crit) => {
    if (isErr(field)) return field;
    const s = dbSelect(ctx, db, field, crit);
    if (isErr(s)) return s;
    if (needField !== false && s.col == null) return E.VALUE;
    const vals = s.rows.map((r) => (s.col == null ? null : C.cellValue(s.sh, r, s.db.c1 + s.col)));
    return f(vals, s, ctx);
  }, { errOk: true });
  const nm = (v) => v.filter((x) => typeof x === 'number');
  dbFn('DSUM', (v) => nm(v).reduce((a, b) => a + b, 0));
  dbFn('DAVERAGE', (v) => { const n = nm(v); return n.length ? n.reduce((a, b) => a + b, 0) / n.length : E.DIV0; });
  dbFn('DCOUNT', (v, s) => (s.col == null ? 0 : nm(v).length), false);
  dbFn('DCOUNTA', (v, s) => (s.col == null ? s.rows.length : v.filter((x) => x != null && x !== '').length), false);
  dbFn('DMAX', (v) => { const n = nm(v); return n.length ? Math.max(...n) : 0; });
  dbFn('DMIN', (v) => { const n = nm(v); return n.length ? Math.min(...n) : 0; });
  dbFn('DPRODUCT', (v) => { const n = nm(v); return n.length ? n.reduce((a, b) => a * b, 1) : 0; });
  dbFn('DGET', (v) => (v.length === 0 ? E.VALUE : v.length > 1 ? E.NUM : v[0] == null ? 0 : v[0]));
  const dv = (pop, root) => (v) => { const n = nm(v); if (n.length < (pop ? 1 : 2)) return E.DIV0; const r = S.devsq(n) / (pop ? n.length : n.length - 1); return root ? Math.sqrt(r) : r; };
  dbFn('DSTDEV', dv(false, true)); dbFn('DSTDEVP', dv(true, true)); dbFn('DVAR', dv(false, false)); dbFn('DVARP', dv(true, false));
  C.dbSelect = dbSelect;
  void Arr;
})(typeof window !== 'undefined' ? window : globalThis);
