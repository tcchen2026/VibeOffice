/* Ledger — statistical functions and the special functions behind them
 * (erf, gamma, incomplete gamma/beta, normal / t / chi-square / F / beta / binomial distributions, regression).
 */
(function (root) {
  'use strict';
  const L = root.L;
  const C = L.calc, M = L.model;
  const E = M.ERR;
  const { def, isErr, MISSING, Arr, Ref } = C;
  const opt = C.opt;
  const trunc = Math.trunc;
  const S = (C.special = {});

  /* --------------------------------------------------------- special functions */
  /* W. J. Cody's rational approximations (SPECFUN CALERF) */
  const EA = [3.16112374387056560e00, 1.13864154151050156e02, 3.77485237685302021e02, 3.20937758913846947e03, 1.85777706184603153e-1];
  const EB = [2.36012909523441209e01, 2.44024637934444173e02, 1.28261652607737228e03, 2.84423683343917062e03];
  const EC = [5.64188496988670089e-1, 8.88314979438837594e00, 6.61191906371416295e01, 2.98635138197400131e02, 8.81952221241769090e02, 1.71204761263407058e03, 2.05107837782607147e03, 1.23033935479799725e03, 2.15311535474403846e-8];
  const ED = [1.57449261107098347e01, 1.17693950891312499e02, 5.37181101862009858e02, 1.62138957456669019e03, 3.29079923573345963e03, 4.36261909014324716e03, 3.43936767414372164e03, 1.23033935480374942e03];
  const EP = [3.05326634961232344e-1, 3.60344899949804439e-1, 1.25781726111229246e-1, 1.60837851487422766e-2, 6.58749161529837803e-4, 1.63153871373020978e-2];
  const EQ = [2.56852019228982242e00, 1.87295284992346725e00, 5.27905102951428412e-1, 6.05183413124413191e-2, 2.33520497626869185e-3];
  const SQRPI = 5.6418958354775628695e-1;
  /** jint 0: erf, 1: erfc */
  function calerf(x, jint) {
    const y = Math.abs(x);
    let result;
    if (y <= 0.46875) {
      let ysq = y > 1.11e-16 ? y * y : 0;
      let xnum = EA[4] * ysq, xden = ysq;
      for (let i = 0; i < 3; i++) { xnum = (xnum + EA[i]) * ysq; xden = (xden + EB[i]) * ysq; }
      result = (x * (xnum + EA[3])) / (xden + EB[3]);
      return jint ? 1 - result : result;
    }
    if (y <= 4) {
      let xnum = EC[8] * y, xden = y;
      for (let i = 0; i < 7; i++) { xnum = (xnum + EC[i]) * y; xden = (xden + ED[i]) * y; }
      result = (xnum + EC[7]) / (xden + ED[7]);
      const ysq = Math.trunc(y * 16) / 16;
      const del = (y - ysq) * (y + ysq);
      result = Math.exp(-ysq * ysq) * Math.exp(-del) * result;
    } else {
      if (y >= 26.543) result = 0;
      else {
        const ysq = 1 / (y * y);
        let xnum = EP[5] * ysq, xden = ysq;
        for (let i = 0; i < 4; i++) { xnum = (xnum + EP[i]) * ysq; xden = (xden + EQ[i]) * ysq; }
        result = (ysq * (xnum + EP[4])) / (xden + EQ[4]);
        result = (SQRPI - result) / y;
        const ysq2 = Math.trunc(y * 16) / 16;
        const del = (y - ysq2) * (y + ysq2);
        result = Math.exp(-ysq2 * ysq2) * Math.exp(-del) * result;
      }
    }
    /* result is erfc(|x|) here */
    if (jint === 0) { result = 0.5 - result + 0.5; if (x < 0) result = -result; }
    else if (x < 0) result = 2 - result;
    return result;
  }
  S.erf = (x) => calerf(x, 0);
  S.erfc = (x) => calerf(x, 1);
  S.normCdf = (z) => 0.5 * S.erfc(-z / Math.SQRT2);
  S.normPdf = (z) => Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
  /** Wichura AS241 */
  S.normInv = function (p) {
    if (p <= 0 || p >= 1) return NaN;
    const q = p - 0.5;
    if (Math.abs(q) <= 0.425) {
      const r = 0.180625 - q * q;
      return (q * (((((((2509.0809287301226727 * r + 33430.575583588128105) * r + 67265.770927008700853) * r + 45921.953931549871457) * r + 13731.693765509461125) * r + 1971.5909503065514427) * r + 133.14166789178437745) * r + 3.387132872796366608)) /
        (((((((5226.495278852545925 * r + 28729.085735721942674) * r + 39307.89580009271061) * r + 21213.794301586595867) * r + 5394.1960214247511077) * r + 687.1870074920579083) * r + 42.313330701600911252) * r + 1);
    }
    let r = q < 0 ? p : 1 - p;
    r = Math.sqrt(-Math.log(r));
    let val;
    if (r <= 5) {
      r -= 1.6;
      val = (((((((7.7454501427834140764e-4 * r + 0.0227238449892691845833) * r + 0.24178072517745061177) * r + 1.27045825245236838258) * r + 3.64784832476320460504) * r + 5.7694972214606914055) * r + 4.6303378461565452959) * r + 1.42343711074968357734) /
        (((((((1.05075007164441684324e-9 * r + 5.475938084995344946e-4) * r + 0.0151986665636164571966) * r + 0.14810397642748007459) * r + 0.68976733498510000455) * r + 1.6763848301838038494) * r + 2.05319162663775882187) * r + 1);
    } else {
      r -= 5;
      val = (((((((2.01033439929228813265e-7 * r + 2.71155556874348757815e-5) * r + 0.0012426609473880784386) * r + 0.026532189526576123093) * r + 0.29656057182850489123) * r + 1.7848265399172913358) * r + 5.4637849111641143699) * r + 6.6579046435011037772) /
        (((((((2.04426310338993978564e-15 * r + 1.4215117583164458887e-7) * r + 1.8463183175100546818e-5) * r + 7.868691311456132591e-4) * r + 0.0148753612908506148525) * r + 0.13692988092273580531) * r + 0.59983220655588793769) * r + 1);
    }
    return q < 0 ? -val : val;
  };
  /* Lanczos gamma (g = 7, n = 9) */
  const LG = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
  S.gamma = function (z) {
    if (z < 0.5) return Math.PI / (Math.sin(Math.PI * z) * S.gamma(1 - z));
    if (z > 171.62) return Infinity;
    if (Number.isInteger(z) && z <= 171) return C.fact(z - 1);
    z -= 1;
    let x = LG[0];
    for (let i = 1; i < 9; i++) x += LG[i] / (z + i);
    const t = z + 7.5;
    return Math.sqrt(2 * Math.PI) * Math.pow(t, z + 0.5) * Math.exp(-t) * x;
  };
  S.lnGamma = function (z) {
    if (z < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * z))) - S.lnGamma(1 - z);
    if (Number.isInteger(z) && z < 171) return Math.log(C.fact(z - 1));
    z -= 1;
    let x = LG[0];
    for (let i = 1; i < 9; i++) x += LG[i] / (z + i);
    const t = z + 7.5;
    return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(x);
  };
  /** regularised lower incomplete gamma P(a, x) */
  S.gammaP = function (a, x) {
    if (x <= 0) return 0;
    if (x < a + 1) {
      let sum = 1 / a, del = sum, ap = a;
      for (let n = 0; n < 1000; n++) { ap++; del *= x / ap; sum += del; if (Math.abs(del) < Math.abs(sum) * 1e-16) break; }
      return sum * Math.exp(-x + a * Math.log(x) - S.lnGamma(a));
    }
    return 1 - S.gammaQcf(a, x);
  };
  S.gammaQcf = function (a, x) {
    let b = x + 1 - a, c = 1 / 1e-300, d = 1 / b, h = d;
    for (let i = 1; i < 1000; i++) {
      const an = -i * (i - a);
      b += 2;
      d = an * d + b; if (Math.abs(d) < 1e-300) d = 1e-300;
      c = b + an / c; if (Math.abs(c) < 1e-300) c = 1e-300;
      d = 1 / d;
      const del = d * c;
      h *= del;
      if (Math.abs(del - 1) < 1e-16) break;
    }
    return Math.exp(-x + a * Math.log(x) - S.lnGamma(a)) * h;
  };
  S.gammaQ = (a, x) => (x < a + 1 ? 1 - S.gammaP(a, x) : S.gammaQcf(a, x));
  /** regularised incomplete beta I_x(a, b) */
  S.betaI = function (x, a, b) {
    if (x <= 0) return 0;
    if (x >= 1) return 1;
    const bt = Math.exp(S.lnGamma(a + b) - S.lnGamma(a) - S.lnGamma(b) + a * Math.log(x) + b * Math.log(1 - x));
    if (x < (a + 1) / (a + b + 2)) return (bt * betacf(x, a, b)) / a;
    return 1 - (bt * betacf(1 - x, b, a)) / b;
  };
  function betacf(x, a, b) {
    const qab = a + b, qap = a + 1, qam = a - 1;
    let c = 1, d = 1 - (qab * x) / qap;
    if (Math.abs(d) < 1e-300) d = 1e-300;
    d = 1 / d;
    let h = d;
    for (let m = 1; m <= 3000; m++) {
      const m2 = 2 * m;
      let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
      d = 1 + aa * d; if (Math.abs(d) < 1e-300) d = 1e-300;
      c = 1 + aa / c; if (Math.abs(c) < 1e-300) c = 1e-300;
      d = 1 / d; h *= d * c;
      aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
      d = 1 + aa * d; if (Math.abs(d) < 1e-300) d = 1e-300;
      c = 1 + aa / c; if (Math.abs(c) < 1e-300) c = 1e-300;
      d = 1 / d;
      const del = d * c;
      h *= del;
      if (Math.abs(del - 1) < 1e-16) break;
    }
    return h;
  }
  /** invert a monotone cdf on [lo, hi] (bisection + secant polish) */
  S.invert = function (cdf, p, lo, hi, increasing) {
    if (increasing === undefined) increasing = true;
    /* widen the bracket when needed */
    let flo = cdf(lo) - p, fhi = cdf(hi) - p;
    let guard = 0;
    while (flo * fhi > 0 && guard++ < 200) { hi = hi * 2 + 1; fhi = cdf(hi) - p; }
    if (flo * fhi > 0) return NaN;
    for (let i = 0; i < 300; i++) {
      const mid = (lo + hi) / 2;
      const f = cdf(mid) - p;
      if (Math.abs(f) < 1e-15 || (hi - lo) < 1e-15 * Math.max(1, Math.abs(mid))) return mid;
      if ((f < 0) === (flo < 0)) { lo = mid; flo = f; } else { hi = mid; }
    }
    return (lo + hi) / 2;
  };
  S.tCdf = (t, df) => { const x = df / (df + t * t); const tail = 0.5 * S.betaI(x, df / 2, 0.5); return t > 0 ? 1 - tail : tail; };
  S.tPdf = (t, df) => Math.exp(S.lnGamma((df + 1) / 2) - S.lnGamma(df / 2)) / Math.sqrt(df * Math.PI) * Math.pow(1 + (t * t) / df, -(df + 1) / 2);
  S.tInv = (p, df) => { if (p === 0.5) return 0; const r = S.invert((x) => S.tCdf(x, df), p, -1e3, 1e3); return r; };
  S.chiCdf = (x, k) => S.gammaP(k / 2, x / 2);
  S.fCdf = (x, d1, d2) => (x <= 0 ? 0 : S.betaI((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2));
  S.betaPdf = (x, a, b) => Math.exp((a - 1) * Math.log(x) + (b - 1) * Math.log(1 - x) - (S.lnGamma(a) + S.lnGamma(b) - S.lnGamma(a + b)));

  /* --------------------------------------------------------- collections */
  const nums = (args, ctx, all) => C.nums(Array.isArray(args) ? args : [args], ctx, all ? { all: true } : undefined);
  /** numbers from a range/array argument (refs: numbers only) */
  const numsOf = (a, ctx) => { const v = C.nums([a], ctx); return v; };
  const sum = (a) => { let s = 0; for (const x of a) s += x; return s; };
  const mean = (a) => sum(a) / a.length;
  const devsq = (a) => { const m = mean(a); let s = 0; for (const x of a) s += (x - m) * (x - m); return s; };
  S.mean = mean; S.devsq = devsq;

  def('COUNT', 1, 255, 'r', (ctx, args) => {
    let n = 0;
    C.each(args, (v, d) => { if (typeof v === 'number') n++; else if (d && (typeof v === 'boolean' || (typeof v === 'string' && C.parseNum(v, ctx.wb) != null))) n++; }, { missingAsZero: false });
    for (const a of args) if (a === MISSING) n++;
    return n;
  }, { list: true, errOk: true });
  def('COUNTA', 1, 255, 'r', (ctx, args) => { let n = 0; C.each(args, () => { n++; }); for (const a of args) if (a === MISSING) n++; return n; }, { list: true, errOk: true });
  def('COUNTBLANK', 1, 1, 'r', (ctx, rg) => {
    if (rg instanceof Ref) { let filled = 0; C.eachRef(rg, (v) => { if (v !== '') filled++; }); return rg.rows * rg.cols - filled; }
    if (rg instanceof Arr) return rg.d.flat().filter((v) => v == null || v === '').length;
    return E.VALUE;
  }, { errOk: true });
  def('AVERAGE', 1, 255, 'r', (ctx, args) => { const v = nums(args, ctx); if (isErr(v)) return v; return v.length ? mean(v) : E.DIV0; }, { list: true });
  def('AVERAGEA', 1, 255, 'r', (ctx, args) => { const v = nums(args, ctx, true); if (isErr(v)) return v; return v.length ? mean(v) : E.DIV0; }, { list: true });
  const mx = (all, isMax) => (ctx, args) => { const v = nums(args, ctx, all); if (isErr(v)) return v; if (!v.length) return 0; let m = v[0]; for (const x of v) if (isMax ? x > m : x < m) m = x; return m; };
  def('MAX', 1, 255, 'r', mx(false, true), { list: true });
  def('MIN', 1, 255, 'r', mx(false, false), { list: true });
  def('MAXA', 1, 255, 'r', mx(true, true), { list: true });
  def('MINA', 1, 255, 'r', mx(true, false), { list: true });
  const sorted = (v) => v.slice().sort((a, b) => a - b);
  def('MEDIAN', 1, 255, 'r', (ctx, args) => { const v = nums(args, ctx); if (isErr(v)) return v; if (!v.length) return E.NUM; const s = sorted(v); const n = s.length; return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; }, { list: true });
  function modes(v) {
    const counts = new Map();
    let best = 1;
    for (const x of v) { const k = counts.get(x) || 0; counts.set(x, k + 1); if (k + 1 > best) best = k + 1; }
    if (best < 2) return [];
    const out = [];
    const seen = new Set();
    for (const x of v) if (counts.get(x) === best && !seen.has(x)) { seen.add(x); out.push(x); }
    return out;
  }
  def('MODE', 1, 255, 'r', (ctx, args) => { if (args.some((a) => typeof a === 'string' || typeof a === 'boolean')) return E.VALUE; const v = nums(args, ctx); if (isErr(v)) return v; const m = modes(v); return m.length ? m[0] : E.NA; }, { list: true });
  C.alias('MODE.SNGL', 'MODE');
  def('MODE.MULT', 1, 255, 'r', (ctx, args) => { const v = nums(args, ctx); if (isErr(v)) return v; const m = modes(v); return m.length ? new Arr(m.map((x) => [x])) : E.NA; }, { list: true });
  def('GEOMEAN', 1, 255, 'r', (ctx, args) => { const v = nums(args, ctx); if (isErr(v)) return v; if (!v.length) return E.NUM; let s = 0; for (const x of v) { if (x <= 0) return E.NUM; s += Math.log(x); } return Math.exp(s / v.length); }, { list: true });
  def('HARMEAN', 1, 255, 'r', (ctx, args) => { const v = nums(args, ctx); if (isErr(v)) return v; if (!v.length) return E.NUM; let s = 0; for (const x of v) { if (x <= 0) return E.NUM; s += 1 / x; } return v.length / s; }, { list: true });
  def('AVEDEV', 1, 255, 'r', (ctx, args) => { const v = nums(args, ctx); if (isErr(v)) return v; if (!v.length) return E.NUM; const m = mean(v); return v.reduce((a, x) => a + Math.abs(x - m), 0) / v.length; }, { list: true });
  def('DEVSQ', 1, 255, 'r', (ctx, args) => { const v = nums(args, ctx); if (isErr(v)) return v; return v.length ? devsq(v) : E.NUM; }, { list: true });
  const variance = (all, pop) => (ctx, args) => { const v = nums(args, ctx, all); if (isErr(v)) return v; const n = v.length; if (n < (pop ? 1 : 2)) return E.DIV0; return devsq(v) / (pop ? n : n - 1); };
  def('VAR', 1, 255, 'r', variance(false, false), { list: true }); C.alias('VAR.S', 'VAR');
  def('VARP', 1, 255, 'r', variance(false, true), { list: true }); C.alias('VAR.P', 'VARP');
  def('VARA', 1, 255, 'r', variance(true, false), { list: true });
  def('VARPA', 1, 255, 'r', variance(true, true), { list: true });
  const sd = (all, pop) => (ctx, args) => { const v = variance(all, pop)(ctx, args); return isErr(v) ? v : Math.sqrt(v); };
  def('STDEV', 1, 255, 'r', sd(false, false), { list: true }); C.alias('STDEV.S', 'STDEV');
  def('STDEVP', 1, 255, 'r', sd(false, true), { list: true }); C.alias('STDEV.P', 'STDEVP');
  def('STDEVA', 1, 255, 'r', sd(true, false), { list: true });
  def('STDEVPA', 1, 255, 'r', sd(true, true), { list: true });
  def('KURT', 1, 255, 'r', (ctx, args) => {
    const v = nums(args, ctx); if (isErr(v)) return v;
    const n = v.length; if (n < 4) return E.DIV0;
    const m = mean(v), s = Math.sqrt(devsq(v) / (n - 1)); if (s === 0) return E.DIV0;
    let k = 0; for (const x of v) k += Math.pow((x - m) / s, 4);
    return (n * (n + 1) / ((n - 1) * (n - 2) * (n - 3))) * k - (3 * (n - 1) * (n - 1)) / ((n - 2) * (n - 3));
  }, { list: true });
  def('SKEW', 1, 255, 'r', (ctx, args) => {
    const v = nums(args, ctx); if (isErr(v)) return v;
    const n = v.length; if (n < 3) return E.DIV0;
    const m = mean(v), s = Math.sqrt(devsq(v) / (n - 1)); if (s === 0) return E.DIV0;
    let k = 0; for (const x of v) k += Math.pow((x - m) / s, 3);
    return (n / ((n - 1) * (n - 2))) * k;
  }, { list: true });
  def('SKEW.P', 1, 255, 'r', (ctx, args) => {
    const v = nums(args, ctx); if (isErr(v)) return v;
    const n = v.length; if (n < 1) return E.DIV0;
    const m = mean(v), s = Math.sqrt(devsq(v) / n); if (s === 0) return E.DIV0;
    let k = 0; for (const x of v) k += Math.pow((x - m) / s, 3);
    return k / n;
  }, { list: true });
  def('LARGE', 2, 2, 'rn', (ctx, a, k) => { const v = numsOf(a, ctx); if (isErr(v)) return v; k = Math.ceil(k); if (k < 1 || k > v.length) return E.NUM; return sorted(v)[v.length - k]; });
  def('SMALL', 2, 2, 'rn', (ctx, a, k) => { const v = numsOf(a, ctx); if (isErr(v)) return v; k = Math.ceil(k); if (k < 1 || k > v.length) return E.NUM; return sorted(v)[k - 1]; });
  C.FN.LARGE.lift = true; C.FN.SMALL.lift = true;
  const pctInc = (s, k) => { const n = s.length; const h = (n - 1) * k; const lo = Math.floor(h); return lo + 1 < n ? s[lo] + (h - lo) * (s[lo + 1] - s[lo]) : s[lo]; };
  const pctExc = (s, k) => { const n = s.length; const h = (n + 1) * k - 1; if (h < 0 || h > n - 1) return E.NUM; const lo = Math.floor(h); return lo + 1 < n ? s[lo] + (h - lo) * (s[lo + 1] - s[lo]) : s[lo]; };
  def('PERCENTILE', 2, 2, 'rn', (ctx, a, k) => { const v = numsOf(a, ctx); if (isErr(v)) return v; if (!v.length || k < 0 || k > 1) return E.NUM; return pctInc(sorted(v), k); });
  C.alias('PERCENTILE.INC', 'PERCENTILE');
  def('PERCENTILE.EXC', 2, 2, 'rn', (ctx, a, k) => { const v = numsOf(a, ctx); if (isErr(v)) return v; const n = v.length; if (!n || k <= 0 || k >= 1) return E.NUM; if (k < 1 / (n + 1) || k > n / (n + 1)) return E.NUM; return pctExc(sorted(v), k); });
  def('QUARTILE', 2, 2, 'rn', (ctx, a, q) => { const v = numsOf(a, ctx); if (isErr(v)) return v; q = trunc(q); if (!v.length || q < 0 || q > 4) return E.NUM; return pctInc(sorted(v), q / 4); });
  C.alias('QUARTILE.INC', 'QUARTILE');
  def('QUARTILE.EXC', 2, 2, 'rn', (ctx, a, q) => { const v = numsOf(a, ctx); if (isErr(v)) return v; q = trunc(q); if (!v.length || q <= 0 || q >= 4) return E.NUM; return pctExc(sorted(v), q / 4); });
  const truncSig = (x, sig) => { const p = Math.pow(10, sig); return Math.floor(L.numfmt.r15(x * p)) / p; };
  function percentRank(v, x, sig, exc) {
    const s = sorted(v), n = s.length;
    if (!n) return E.NUM;
    sig = sig === MISSING || sig === undefined ? 3 : trunc(sig);
    if (sig < 1) return E.NUM;
    if (x < s[0] || x > s[n - 1]) return E.NA;
    let below = 0;
    while (below < n && s[below] < x) below++;
    let r;
    if (below < n && s[below] === x) r = exc ? (below + 1) / (n + 1) : n === 1 ? 1 : below / (n - 1);
    else {
      const lo = s[below - 1], hi = s[below];
      const rl = exc ? below / (n + 1) : (below - 1) / (n - 1), rh = exc ? (below + 1) / (n + 1) : below / (n - 1);
      r = rl + ((x - lo) / (hi - lo)) * (rh - rl);
    }
    return truncSig(r, sig);
  }
  def('PERCENTRANK', 2, 3, 'rnn', (ctx, a, x, sig) => { const v = numsOf(a, ctx); if (isErr(v)) return v; return percentRank(v, x, sig, false); });
  C.alias('PERCENTRANK.INC', 'PERCENTRANK');
  def('PERCENTRANK.EXC', 2, 3, 'rnn', (ctx, a, x, sig) => { const v = numsOf(a, ctx); if (isErr(v)) return v; return percentRank(v, x, sig, true); });
  function rank(ctx, x, ref, order, avg) {
    if (!(ref instanceof Ref) && !(ref instanceof C.RefList) && !(ref instanceof Arr)) return E.VALUE;
    const v = numsOf(ref, ctx); if (isErr(v)) return v;
    order = opt(order, 0);
    let higher = 0, equal = 0;
    for (const y of v) { if (C.numEq(y, x)) equal++; else if (order ? y < x : y > x) higher++; }
    if (!equal) return E.NA;
    return avg ? higher + (equal + 1) / 2 : higher + 1;
  }
  def('RANK', 2, 3, 'nrn', (ctx, x, ref, o) => rank(ctx, x, ref, o, false));
  C.alias('RANK.EQ', 'RANK');
  def('RANK.AVG', 2, 3, 'nrn', (ctx, x, ref, o) => rank(ctx, x, ref, o, true));
  def('TRIMMEAN', 2, 2, 'rn', (ctx, a, p) => {
    const v = numsOf(a, ctx); if (isErr(v)) return v;
    if (p < 0 || p >= 1 || !v.length) return E.NUM;
    const s = sorted(v);
    const k = Math.floor((s.length * p) / 2);
    const t = s.slice(k, s.length - k);
    return mean(t);
  });
  def('FREQUENCY', 2, 2, 'aa', (ctx, data, bins) => {
    const scalarOk = (v) => v instanceof Ref || v instanceof Arr || v instanceof C.RefList || typeof v === 'number' || isErr(v);
    if (!scalarOk(data) || !scalarOk(bins)) return E.VALUE;
    const d = C.nums([data], ctx), b = C.nums([bins], ctx);
    if (isErr(d)) return d; if (isErr(b)) return b;
    const order = b.map((x, i) => ({ x, i })).sort((p, q) => p.x - q.x);
    const counts = new Array(b.length + 1).fill(0);
    for (const x of d) {
      let lo = 0, hi = order.length;
      while (lo < hi) { const m = (lo + hi) >> 1; if (x <= order[m].x) hi = m; else lo = m + 1; }
      counts[lo >= order.length ? b.length : order[lo].i]++;
    }
    return new Arr(counts.map((c) => [c]));
  });
  def('PROB', 3, 4, 'aann', (ctx, xr, pr, lo, hi) => {
    const X = C.toArr(xr).d.flat(), P = C.toArr(pr).d.flat();
    if (X.length !== P.length) return E.NA;
    hi = opt(hi, lo);
    let s = 0, tot = 0;
    for (let i = 0; i < X.length; i++) {
      if (typeof X[i] !== 'number' || typeof P[i] !== 'number') continue;
      if (P[i] < 0 || P[i] > 1) return E.NUM;
      tot += P[i];
      if (X[i] >= lo && X[i] <= hi) s += P[i];
    }
    if (Math.abs(tot - 1) > 1e-7) return E.NUM;
    return s;
  });
  def('STANDARDIZE', 3, 3, 'nnn', (ctx, x, m, s) => (s <= 0 ? E.NUM : (x - m) / s));
  def('FISHER', 1, 1, 'n', (ctx, x) => (x <= -1 || x >= 1 ? E.NUM : 0.5 * Math.log((1 + x) / (1 - x))));
  def('FISHERINV', 1, 1, 'n', (ctx, y) => { const e = Math.exp(2 * y); return (e - 1) / (e + 1); });

  /* ------------------------------------------------------------ pairs & regression */
  /** numeric pairs from two equally sized arguments; pairs with a non-number on either side are skipped */
  function paired(a, b) {
    const A = C.toArr(a).d.flat(), B = C.toArr(b).d.flat();
    if (A.length !== B.length) return E.NA;
    const x = [], y = [];
    for (let i = 0; i < A.length; i++) {
      if (isErr(A[i])) return A[i];
      if (isErr(B[i])) return B[i];
      if (typeof A[i] === 'number' && typeof B[i] === 'number') { x.push(A[i]); y.push(B[i]); }
    }
    return { x, y };
  }
  const cov = (x, y, pop) => { const mx2 = mean(x), my = mean(y); let s = 0; for (let i = 0; i < x.length; i++) s += (x[i] - mx2) * (y[i] - my); return s / (pop ? x.length : x.length - 1); };
  def('COVAR', 2, 2, 'rr', (ctx, a, b) => { const p = paired(a, b); if (isErr(p)) return p; if (!p.x.length) return E.DIV0; return cov(p.x, p.y, true); });
  C.alias('COVARIANCE.P', 'COVAR');
  def('COVARIANCE.S', 2, 2, 'rr', (ctx, a, b) => { const p = paired(a, b); if (isErr(p)) return p; if (p.x.length < 2) return E.DIV0; return cov(p.x, p.y, false); });
  const correl = (ctx, a, b) => { const p = paired(a, b); if (isErr(p)) return p; if (p.x.length < 2) return E.DIV0; const sx = devsq(p.x), sy = devsq(p.y); if (!sx || !sy) return E.DIV0; return (cov(p.x, p.y, true) * p.x.length) / Math.sqrt(sx * sy); };
  def('CORREL', 2, 2, 'rr', correl);
  def('PEARSON', 2, 2, 'rr', correl);
  def('RSQ', 2, 2, 'rr', (ctx, a, b) => { const r = correl(ctx, a, b); return isErr(r) ? r : r * r; });
  function slopeInt(ctx, ys, xs) {
    const p = paired(xs, ys); if (isErr(p)) return p;
    if (p.x.length < 2) return E.DIV0;
    const sx = devsq(p.x);
    if (!sx) return E.DIV0;
    const slope = (cov(p.x, p.y, true) * p.x.length) / sx;
    return { slope, icpt: mean(p.y) - slope * mean(p.x), p };
  }
  def('SLOPE', 2, 2, 'rr', (ctx, y, x) => { const r = slopeInt(ctx, y, x); return isErr(r) ? r : r.slope; });
  def('INTERCEPT', 2, 2, 'rr', (ctx, y, x) => { const r = slopeInt(ctx, y, x); return isErr(r) ? r : r.icpt; });
  def('STEYX', 2, 2, 'rr', (ctx, y, x) => {
    const r = slopeInt(ctx, y, x); if (isErr(r)) return r;
    const n = r.p.x.length; if (n < 3) return E.DIV0;
    const sx = devsq(r.p.x), sy = devsq(r.p.y), sxy = cov(r.p.x, r.p.y, true) * n;
    return Math.sqrt((sy - (sxy * sxy) / sx) / (n - 2));
  });
  def('FORECAST', 3, 3, 'nrr', (ctx, x, ys, xs) => { const r = slopeInt(ctx, ys, xs); return isErr(r) ? r : r.icpt + r.slope * x; });
  C.alias('FORECAST.LINEAR', 'FORECAST');
  /* least squares via Householder QR. X: n×p (column of ones prepended when const) */
  function lstsq(X, y) {
    const n = X.length, p = X[0].length;
    const A = X.map((r) => r.slice());
    const b = y.slice();
    const R = [];
    const keep = [];
    for (let k = 0; k < p; k++) {
      let norm = 0;
      for (let i = k; i < n; i++) norm += A[i][k] * A[i][k];
      norm = Math.sqrt(norm);
      if (norm < 1e-12) { keep.push(false); continue; }
      keep.push(true);
      const alpha = A[k][k] > 0 ? -norm : norm;
      const v = new Array(n).fill(0);
      for (let i = k; i < n; i++) v[i] = A[i][k];
      v[k] -= alpha;
      let vn = 0; for (let i = k; i < n; i++) vn += v[i] * v[i];
      if (vn === 0) continue;
      for (let j = k; j < p; j++) { let s = 0; for (let i = k; i < n; i++) s += v[i] * A[i][j]; s = (2 * s) / vn; for (let i = k; i < n; i++) A[i][j] -= s * v[i]; }
      let s = 0; for (let i = k; i < n; i++) s += v[i] * b[i]; s = (2 * s) / vn; for (let i = k; i < n; i++) b[i] -= s * v[i];
    }
    const beta = new Array(p).fill(0);
    for (let k = p - 1; k >= 0; k--) {
      if (!keep[k]) continue;
      let s = b[k];
      for (let j = k + 1; j < p; j++) s -= A[k][j] * beta[j];
      beta[k] = s / A[k][k];
    }
    void R;
    return { beta, keep };
  }
  /** LINEST core: returns {coef (b first), stats} */
  function linest(ys, xs, useConst, wantStats) {
    const n = ys.length;
    const p = xs[0] ? xs[0].length : 0;
    const X = xs.map((r) => (useConst ? [1].concat(r) : r.slice()));
    const { beta, keep } = lstsq(X, ys);
    const fit = X.map((r) => r.reduce((s, v, j) => s + v * beta[j], 0));
    const res = { beta, keep, n, p };
    if (wantStats) {
      const my = useConst ? mean(ys) : 0;
      let ssreg = 0, ssres = 0;
      for (let i = 0; i < n; i++) { ssres += (ys[i] - fit[i]) * (ys[i] - fit[i]); ssreg += (fit[i] - my) * (fit[i] - my); }
      const k = keep.filter(Boolean).length;
      const df = n - k;
      const sey = df > 0 ? Math.sqrt(ssres / df) : E.NUM;
      /* standard errors: sqrt(diag((X'X)^-1)) * sey */
      const XtX = [];
      for (let a = 0; a < X[0].length; a++) { XtX.push([]); for (let b2 = 0; b2 < X[0].length; b2++) { let s = 0; for (let i = 0; i < n; i++) s += X[i][a] * X[i][b2]; XtX[a].push(s); } }
      const inv = invert(XtX, keep);
      res.se = beta.map((_, j) => (keep[j] && inv && typeof sey === 'number' ? Math.sqrt(Math.max(0, inv[j][j])) * sey : 0));
      res.r2 = ssreg + ssres ? ssreg / (ssreg + ssres) : 1;
      res.sey = sey; res.F = df > 0 && ssres ? (ssreg / (k - (useConst ? 1 : 0))) / (ssres / df) : E.NUM; res.df = df; res.ssreg = ssreg; res.ssres = ssres;
    }
    return res;
  }
  function invert(A, keep) {
    const idx = keep.map((k, i) => (k ? i : -1)).filter((i) => i >= 0);
    const n = idx.length;
    const a = idx.map((i) => idx.map((j) => A[i][j]).concat(idx.map((j) => (i === j ? 1 : 0))));
    for (let c = 0; c < n; c++) {
      let p = c; for (let r = c + 1; r < n; r++) if (Math.abs(a[r][c]) > Math.abs(a[p][c])) p = r;
      if (Math.abs(a[p][c]) < 1e-300) return null;
      [a[p], a[c]] = [a[c], a[p]];
      const d = a[c][c];
      for (let j = 0; j < 2 * n; j++) a[c][j] /= d;
      for (let r = 0; r < n; r++) if (r !== c) { const f = a[r][c]; for (let j = 0; j < 2 * n; j++) a[r][j] -= f * a[c][j]; }
    }
    const out = A.map((r) => r.map(() => 0));
    idx.forEach((i, ii) => idx.forEach((j, jj) => { out[i][j] = a[ii][n + jj]; }));
    return out;
  }
  /** shape known_x's into rows of predictors matching known_y's */
  function design(yA, xA) {
    const Y = C.toArr(yA);
    const ys = Y.d.flat();
    for (const v of ys) if (typeof v !== 'number') return E.VALUE;
    const n = ys.length;
    let rows;
    if (xA === MISSING || xA === undefined) rows = ys.map((_, i) => [i + 1]);
    else {
      const X = C.toArr(xA);
      for (const r of X.d) for (const v of r) if (typeof v !== 'number') return E.VALUE;
      if (Y.cols === 1 && X.rows === n) rows = X.d.map((r) => r.slice());
      else if (Y.rows === 1 && X.cols === n) rows = ys.map((_, j) => X.d.map((r) => r[j]));
      else if (X.rows * X.cols === n) rows = X.d.flat().map((v) => [v]);
      else return E.REF;
    }
    return { ys, rows, yShape: Y };
  }
  function linestFn(ctx, y, x, cst, stats, log) {
    const d = design(y, x);
    if (isErr(d)) return d;
    let ys = d.ys;
    if (log) { if (ys.some((v) => v <= 0)) return E.NUM; ys = ys.map(Math.log); }
    const useConst = cst === MISSING || cst === undefined ? true : C.toBool(cst);
    const want = stats === MISSING || stats === undefined ? false : C.toBool(stats);
    const r = linest(ys, d.rows, useConst, want);
    const p = d.rows[0].length;
    const coef = [];
    for (let j = p; j >= 1; j--) coef.push(r.beta[useConst ? j : j - 1]);
    coef.push(useConst ? r.beta[0] : 0);
    const tx = log ? (v) => Math.exp(v) : (v) => v;
    if (!want) return new Arr([coef.map(tx)]);
    const se = [];
    for (let j = p; j >= 1; j--) se.push(r.se[useConst ? j : j - 1]);
    se.push(useConst ? r.se[0] : E.NA);
    const pad = (row) => { const o = row.slice(); while (o.length < p + 1) o.push(E.NA); return o; };
    return new Arr([coef.map(tx), se, pad([r.r2, r.sey]), pad([r.F, r.df]), pad([r.ssreg, r.ssres])]);
  }
  def('LINEST', 1, 4, 'aavv', (ctx, y, x, c, s) => linestFn(ctx, y, x, c, s, false));
  def('LOGEST', 1, 4, 'aavv', (ctx, y, x, c, s) => linestFn(ctx, y, x, c, s, true));
  function trend(ctx, y, x, nx, cst, log) {
    const d = design(y, x);
    if (isErr(d)) return d;
    let ys = d.ys;
    if (log) { if (ys.some((v) => v <= 0)) return E.NUM; ys = ys.map(Math.log); }
    const useConst = cst === MISSING || cst === undefined ? true : C.toBool(cst);
    const r = linest(ys, d.rows, useConst, false);
    let newRows, shape;
    if (nx === MISSING || nx === undefined) { newRows = d.rows; shape = d.yShape; }
    else {
      const N = C.toArr(nx);
      const p = d.rows[0].length;
      if (p === 1) newRows = N.d.flat().map((v) => [v]); else if (N.cols === p) newRows = N.d; else newRows = N.d[0].map((_, j) => N.d.map((rr) => rr[j]));
      shape = N;
    }
    const pred = newRows.map((row) => { let s = useConst ? r.beta[0] : 0; row.forEach((v, j) => { s += v * r.beta[useConst ? j + 1 : j]; }); return log ? Math.exp(s) : s; });
    if (shape.cols === 1 || newRows.length === shape.rows) return new Arr(pred.map((v) => [v]));
    return new Arr([pred]);
  }
  def('TREND', 1, 4, 'aaav', (ctx, y, x, n, c) => trend(ctx, y, x, n, c, false));
  def('GROWTH', 1, 4, 'aaav', (ctx, y, x, n, c) => trend(ctx, y, x, n, c, true));

  /* ------------------------------------------------------------ distributions */
  const bool = (v, d) => (v === MISSING || v === undefined ? d : C.toBool(v));
  def('NORMDIST', 4, 4, 'nnnb', (ctx, x, m, s, cum) => { if (s <= 0) return E.NUM; const z = (x - m) / s; return cum ? S.normCdf(z) : S.normPdf(z) / s; });
  C.alias('NORM.DIST', 'NORMDIST');
  def('NORMSDIST', 1, 1, 'n', (ctx, z) => S.normCdf(z));
  def('NORM.S.DIST', 2, 2, 'nb', (ctx, z, cum) => (cum ? S.normCdf(z) : S.normPdf(z)));
  def('NORMINV', 3, 3, 'nnn', (ctx, p, m, s) => { if (p <= 0 || p >= 1 || s <= 0) return E.NUM; return m + s * S.normInv(p); });
  C.alias('NORM.INV', 'NORMINV');
  def('NORMSINV', 1, 1, 'n', (ctx, p) => (p <= 0 || p >= 1 ? E.NUM : S.normInv(p)));
  C.alias('NORM.S.INV', 'NORMSINV');
  def('PHI', 1, 1, 'n', (ctx, x) => S.normPdf(x));
  def('GAUSS', 1, 1, 'n', (ctx, x) => S.normCdf(x) - 0.5);
  def('LOGNORMDIST', 3, 3, 'nnn', (ctx, x, m, s) => (x <= 0 || s <= 0 ? E.NUM : S.normCdf((Math.log(x) - m) / s)));
  def('LOGNORM.DIST', 4, 4, 'nnnb', (ctx, x, m, s, cum) => { if (x <= 0 || s <= 0) return E.NUM; const z = (Math.log(x) - m) / s; return cum ? S.normCdf(z) : S.normPdf(z) / (x * s); });
  def('LOGINV', 3, 3, 'nnn', (ctx, p, m, s) => (p <= 0 || p >= 1 || s <= 0 ? E.NUM : Math.exp(m + s * S.normInv(p))));
  C.alias('LOGNORM.INV', 'LOGINV');
  def('CONFIDENCE', 3, 3, 'nnn', (ctx, a, s, n) => { n = trunc(n); if (a <= 0 || a >= 1 || s <= 0 || n < 1) return E.NUM; return (-S.normInv(a / 2) * s) / Math.sqrt(n); });
  C.alias('CONFIDENCE.NORM', 'CONFIDENCE');
  def('CONFIDENCE.T', 3, 3, 'nnn', (ctx, a, s, n) => { n = trunc(n); if (a <= 0 || a >= 1 || s <= 0 || n < 1) return E.NUM; if (n === 1) return E.DIV0; return (-S.tInv(a / 2, n - 1) * s) / Math.sqrt(n); });
  /* t */
  def('TDIST', 3, 3, 'nnn', (ctx, x, df, tails) => { df = trunc(df); tails = trunc(tails); if (x < 0 || df < 1 || (tails !== 1 && tails !== 2)) return E.NUM; const p = 1 - S.tCdf(x, df); return tails === 2 ? 2 * p : p; });
  def('T.DIST', 3, 3, 'nnb', (ctx, x, df, cum) => { df = trunc(df); if (df < 1) return E.NUM; return cum ? S.tCdf(x, df) : S.tPdf(x, df); });
  def('T.DIST.2T', 2, 2, 'nn', (ctx, x, df) => { df = trunc(df); if (x < 0 || df < 1) return E.NUM; return 2 * (1 - S.tCdf(x, df)); });
  def('T.DIST.RT', 2, 2, 'nn', (ctx, x, df) => { df = trunc(df); if (df < 1) return E.NUM; return 1 - S.tCdf(x, df); });
  def('TINV', 2, 2, 'nn', (ctx, p, df) => { df = trunc(df); if (p <= 0 || p > 1 || df < 1) return E.NUM; return Math.abs(S.tInv(p / 2, df)); });
  C.alias('T.INV.2T', 'TINV');
  def('T.INV', 2, 2, 'nn', (ctx, p, df) => { df = trunc(df); if (p <= 0 || p >= 1 || df < 1) return E.NUM; return S.tInv(p, df); });
  /* chi-square */
  def('CHIDIST', 2, 2, 'nn', (ctx, x, df) => { df = trunc(df); if (x < 0 || df < 1 || df > 1e10) return E.NUM; return S.gammaQ(df / 2, x / 2); });
  C.alias('CHISQ.DIST.RT', 'CHIDIST');
  def('CHISQ.DIST', 3, 3, 'nnb', (ctx, x, df, cum) => { df = trunc(df); if (x < 0 || df < 1) return E.NUM; if (cum) return S.gammaP(df / 2, x / 2); if (x === 0) return df === 2 ? 0.5 : df === 1 ? E.DIV0 : 0; return Math.exp((df / 2 - 1) * Math.log(x) - x / 2 - (df / 2) * Math.LN2 - S.lnGamma(df / 2)); });
  def('CHIINV', 2, 2, 'nn', (ctx, p, df) => { df = trunc(df); if (p <= 0 || p > 1 || df < 1) return E.NUM; if (p === 1) return 0; return S.invert((x) => 1 - S.gammaP(df / 2, x / 2), p, 0, Math.max(10, df * 10), false); });
  C.alias('CHISQ.INV.RT', 'CHIINV');
  def('CHISQ.INV', 2, 2, 'nn', (ctx, p, df) => { df = trunc(df); if (p < 0 || p >= 1 || df < 1) return E.NUM; if (p === 0) return 0; return S.invert((x) => S.gammaP(df / 2, x / 2), p, 0, Math.max(10, df * 10)); });
  /* F */
  def('FDIST', 3, 3, 'nnn', (ctx, x, a, b) => { a = trunc(a); b = trunc(b); if (x < 0 || a < 1 || b < 1) return E.NUM; return 1 - S.fCdf(x, a, b); });
  C.alias('F.DIST.RT', 'FDIST');
  def('F.DIST', 4, 4, 'nnnb', (ctx, x, a, b, cum) => {
    a = trunc(a); b = trunc(b); if (x < 0 || a < 1 || b < 1) return E.NUM;
    if (cum) return S.fCdf(x, a, b);
    if (x === 0) return a === 2 ? 1 : a < 2 ? E.NUM : 0;
    return Math.exp(0.5 * (a * Math.log(a * x) + b * Math.log(b) - (a + b) * Math.log(a * x + b)) - Math.log(x) - (S.lnGamma(a / 2) + S.lnGamma(b / 2) - S.lnGamma((a + b) / 2)));
  });
  def('FINV', 3, 3, 'nnn', (ctx, p, a, b) => { a = trunc(a); b = trunc(b); if (p <= 0 || p > 1 || a < 1 || b < 1) return E.NUM; if (p === 1) return 0; return S.invert((x) => 1 - S.fCdf(x, a, b), p, 0, 1000, false); });
  C.alias('F.INV.RT', 'FINV');
  def('F.INV', 3, 3, 'nnn', (ctx, p, a, b) => { a = trunc(a); b = trunc(b); if (p < 0 || p >= 1 || a < 1 || b < 1) return E.NUM; if (p === 0) return 0; return S.invert((x) => S.fCdf(x, a, b), p, 0, 1000); });
  /* beta */
  def('BETADIST', 3, 5, 'nnnnn', (ctx, x, a, b, A, B) => { A = opt(A, 0); B = opt(B, 1); if (a <= 0 || b <= 0 || x < A || x > B || A === B) return E.NUM; return S.betaI((x - A) / (B - A), a, b); });
  def('BETA.DIST', 4, 6, 'nnnbnn', (ctx, x, a, b, cum, A, B) => {
    A = opt(A, 0); B = opt(B, 1);
    if (a <= 0 || b <= 0 || x < A || x > B || A === B) return E.NUM;
    const t = (x - A) / (B - A);
    return cum ? S.betaI(t, a, b) : S.betaPdf(t, a, b) / (B - A);
  });
  def('BETAINV', 3, 5, 'nnnnn', (ctx, p, a, b, A, B) => { A = opt(A, 0); B = opt(B, 1); if (p <= 0 || p > 1 || a <= 0 || b <= 0 || A >= B) return E.NUM; return A + (B - A) * S.invert((x) => S.betaI(x, a, b), p, 0, 1); });
  C.alias('BETA.INV', 'BETAINV');
  /* gamma */
  def('GAMMA', 1, 1, 'n', (ctx, x) => { if (x <= 0 && Number.isInteger(x)) return E.NUM; const g = S.gamma(x); return isFinite(g) ? g : E.NUM; });
  def('GAMMALN', 1, 1, 'n', (ctx, x) => (x <= 0 ? E.NUM : S.lnGamma(x)));
  C.alias('GAMMALN.PRECISE', 'GAMMALN');
  def('GAMMADIST', 4, 4, 'nnnb', (ctx, x, a, b, cum) => {
    if (x < 0 || a <= 0 || b <= 0) return E.NUM;
    if (cum) return S.gammaP(a, x / b);
    if (x === 0) return a < 1 ? E.NUM : a === 1 ? 1 / b : 0;
    return Math.exp((a - 1) * Math.log(x) - x / b - a * Math.log(b) - S.lnGamma(a));
  });
  C.alias('GAMMA.DIST', 'GAMMADIST');
  def('GAMMAINV', 3, 3, 'nnn', (ctx, p, a, b) => { if (p < 0 || p >= 1 || a <= 0 || b <= 0) return E.NUM; if (p === 0) return 0; return S.invert((x) => S.gammaP(a, x / b), p, 0, Math.max(1, a * b * 10)); });
  C.alias('GAMMA.INV', 'GAMMAINV');
  def('EXPONDIST', 3, 3, 'nnb', (ctx, x, l, cum) => { if (x < 0 || l <= 0) return E.NUM; return cum ? 1 - Math.exp(-l * x) : l * Math.exp(-l * x); });
  C.alias('EXPON.DIST', 'EXPONDIST');
  def('WEIBULL', 4, 4, 'nnnb', (ctx, x, a, b, cum) => { if (x < 0 || a <= 0 || b <= 0) return E.NUM; return cum ? 1 - Math.exp(-Math.pow(x / b, a)) : (a / Math.pow(b, a)) * Math.pow(x, a - 1) * Math.exp(-Math.pow(x / b, a)); });
  C.alias('WEIBULL.DIST', 'WEIBULL');
  /* discrete */
  const lnC = (n, k) => S.lnGamma(n + 1) - S.lnGamma(k + 1) - S.lnGamma(n - k + 1);
  const binPmf = (k, n, p) => (p === 0 ? (k === 0 ? 1 : 0) : p === 1 ? (k === n ? 1 : 0) : Math.exp(lnC(n, k) + k * Math.log(p) + (n - k) * Math.log(1 - p)));
  def('BINOMDIST', 4, 4, 'nnnb', (ctx, k, n, p, cum) => {
    k = trunc(k); n = trunc(n);
    if (k < 0 || k > n || p < 0 || p > 1) return E.NUM;
    if (!cum) return binPmf(k, n, p);
    let s = 0; for (let i = 0; i <= k; i++) s += binPmf(i, n, p);
    return Math.min(1, s);
  });
  C.alias('BINOM.DIST', 'BINOMDIST');
  def('BINOM.DIST.RANGE', 3, 4, 'nnnn', (ctx, n, p, a, b) => {
    n = trunc(n); a = trunc(a); b = trunc(opt(b, a));
    if (n < 0 || p < 0 || p > 1 || a < 0 || a > n || b < a || b > n) return E.NUM;
    let s = 0; for (let i = a; i <= b; i++) s += binPmf(i, n, p);
    return s;
  });
  def('CRITBINOM', 3, 3, 'nnn', (ctx, n, p, alpha) => {
    n = trunc(n);
    if (n < 0 || p < 0 || p > 1 || alpha < 0 || alpha > 1) return E.NUM;
    let s = 0;
    for (let k = 0; k <= n; k++) { s += binPmf(k, n, p); if (s >= alpha - 1e-12) return k; }
    return n;
  });
  C.alias('BINOM.INV', 'CRITBINOM');
  def('NEGBINOMDIST', 3, 3, 'nnn', (ctx, f, s, p) => { f = trunc(f); s = trunc(s); if (f < 0 || s < 1 || p < 0 || p > 1) return E.NUM; return Math.exp(lnC(f + s - 1, s - 1) + s * Math.log(p) + f * Math.log(1 - p)); });
  def('NEGBINOM.DIST', 4, 4, 'nnnb', (ctx, f, s, p, cum) => {
    f = trunc(f); s = trunc(s); if (f < 0 || s < 1 || p < 0 || p > 1) return E.NUM;
    const pmf = (k) => Math.exp(lnC(k + s - 1, s - 1) + s * Math.log(p) + k * Math.log(1 - p));
    if (!cum) return pmf(f);
    let t = 0; for (let k = 0; k <= f; k++) t += pmf(k);
    return t;
  });
  def('POISSON', 3, 3, 'nnb', (ctx, x, m, cum) => {
    x = trunc(x); if (x < 0 || m < 0) return E.NUM;
    const pmf = (k) => Math.exp(k * Math.log(m) - m - S.lnGamma(k + 1));
    if (m === 0) return x === 0 || cum ? 1 : 0;
    if (!cum) return pmf(x);
    return S.gammaQ(x + 1, m);
  });
  C.alias('POISSON.DIST', 'POISSON');
  const hyp = (s, n, M2, N) => Math.exp(lnC(M2, s) + lnC(N - M2, n - s) - lnC(N, n));
  def('HYPGEOMDIST', 4, 4, 'nnnn', (ctx, s, n, M2, N) => { s = trunc(s); n = trunc(n); M2 = trunc(M2); N = trunc(N); if (s < 0 || s > n || s > M2 || n > N || M2 > N || n - s > N - M2 || n <= 0 || M2 <= 0 || N <= 0) return E.NUM; return hyp(s, n, M2, N); });
  def('HYPGEOM.DIST', 5, 5, 'nnnnb', (ctx, s, n, M2, N, cum) => {
    s = trunc(s); n = trunc(n); M2 = trunc(M2); N = trunc(N);
    if (s < 0 || s > n || s > M2 || n > N || M2 > N || n <= 0 || M2 <= 0 || N <= 0) return E.NUM;
    if (!cum) return n - s > N - M2 ? E.NUM : hyp(s, n, M2, N);
    let t = 0; for (let k = Math.max(0, n - (N - M2)); k <= s; k++) t += hyp(k, n, M2, N);
    return t;
  });
  /* tests */
  def('ZTEST', 2, 3, 'rnn', (ctx, a, x, sigma) => {
    const v = numsOf(a, ctx); if (isErr(v)) return v;
    const n = v.length; if (n < 1) return E.NA;
    const s = sigma === MISSING || sigma === undefined ? Math.sqrt(devsq(v) / (n - 1)) : sigma;
    if (!s) return E.DIV0;
    return 1 - S.normCdf((mean(v) - x) / (s / Math.sqrt(n)));
  });
  C.alias('Z.TEST', 'ZTEST');
  def('TTEST', 4, 4, 'aann', (ctx, a, b, tails, type) => {
    tails = trunc(tails); type = trunc(type);
    if ((tails !== 1 && tails !== 2) || type < 1 || type > 3) return E.NUM;
    let t, df;
    if (type === 1) {
      const p = paired(a, b); if (isErr(p)) return p;
      const d = p.x.map((x, i) => x - p.y[i]);
      const n = d.length; if (n < 2) return E.DIV0;
      const sd2 = Math.sqrt(devsq(d) / (n - 1));
      if (!sd2) return E.DIV0;
      t = Math.abs(mean(d) / (sd2 / Math.sqrt(n))); df = n - 1;
    } else {
      const x = C.nums([a], ctx), y = C.nums([b], ctx);
      if (isErr(x)) return x; if (isErr(y)) return y;
      const n1 = x.length, n2 = y.length; if (n1 < 2 || n2 < 2) return E.DIV0;
      const v1 = devsq(x) / (n1 - 1), v2 = devsq(y) / (n2 - 1);
      if (type === 2) { const sp = ((n1 - 1) * v1 + (n2 - 1) * v2) / (n1 + n2 - 2); t = Math.abs(mean(x) - mean(y)) / Math.sqrt(sp * (1 / n1 + 1 / n2)); df = n1 + n2 - 2; }
      else { const se = v1 / n1 + v2 / n2; t = Math.abs(mean(x) - mean(y)) / Math.sqrt(se); df = (se * se) / ((v1 / n1) * (v1 / n1) / (n1 - 1) + (v2 / n2) * (v2 / n2) / (n2 - 1)); }
    }
    const p = 1 - S.tCdf(t, df);
    return tails === 2 ? 2 * p : p;
  });
  C.alias('T.TEST', 'TTEST');
  def('FTEST', 2, 2, 'aa', (ctx, a, b) => {
    const x = C.nums([a], ctx), y = C.nums([b], ctx);
    if (isErr(x)) return x; if (isErr(y)) return y;
    if (x.length < 2 || y.length < 2) return E.DIV0;
    const v1 = devsq(x) / (x.length - 1), v2 = devsq(y) / (y.length - 1);
    if (!v1 || !v2) return E.DIV0;
    const f = v1 / v2;
    const p = S.fCdf(f, x.length - 1, y.length - 1);
    return 2 * Math.min(p, 1 - p);
  });
  C.alias('F.TEST', 'FTEST');
  def('CHITEST', 2, 2, 'aa', (ctx, a, b) => {
    const A = C.toArr(a), B = C.toArr(b);
    if (A.rows !== B.rows || A.cols !== B.cols) return E.NA;
    let chi = 0;
    for (let i = 0; i < A.rows; i++) for (let j = 0; j < A.cols; j++) {
      const o = A.d[i][j], e = B.d[i][j];
      if (typeof o !== 'number' || typeof e !== 'number') continue;
      if (e === 0) return E.DIV0;
      chi += ((o - e) * (o - e)) / e;
    }
    const df = A.rows > 1 && A.cols > 1 ? (A.rows - 1) * (A.cols - 1) : A.rows * A.cols - 1;
    if (df < 1) return E.NA;
    return S.gammaQ(df / 2, chi / 2);
  });
  C.alias('CHISQ.TEST', 'CHITEST');
  for (const n of ['FORECAST.ETS', 'FORECAST.ETS.CONFINT', 'FORECAST.ETS.SEASONALITY', 'FORECAST.ETS.STAT']) def(n, 0, 255, 'r', (ctx) => { ctx.flags.keep = true; return C.KEEP; }, { list: true });
})(typeof window !== 'undefined' ? window : globalThis);
