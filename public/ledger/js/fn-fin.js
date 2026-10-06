/* Ledger — financial functions: annuities, cash flows, depreciation, bonds and securities. */
(function (root) {
  'use strict';
  const L = root.L;
  const C = L.calc, M = L.model, NF = L.numfmt;
  const E = M.ERR;
  const { def, isErr, MISSING } = C;
  const opt = C.opt;
  const trunc = Math.trunc;

  /* -------------------------------------------------------------- annuities */
  const pv = (r, n, p, f, t) => (r === 0 ? -(p * n + f) : -(f + p * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r) / Math.pow(1 + r, n));
  const fv = (r, n, p, v, t) => (r === 0 ? -(v + p * n) : -(v * Math.pow(1 + r, n) + p * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r));
  const pmt = (r, n, v, f, t) => (r === 0 ? -(v + f) / n : -(r * (f + v * Math.pow(1 + r, n))) / ((1 + r * t) * (Math.pow(1 + r, n) - 1)));
  C.fin = { pv, fv, pmt };
  const T = (t) => (opt(t, 0) ? 1 : 0);
  def('PV', 3, 5, 'nnnnn', (ctx, r, n, p, f, t) => pv(r, n, p, opt(f, 0), T(t)));
  def('FV', 3, 5, 'nnnnn', (ctx, r, n, p, v, t) => fv(r, n, p, opt(v, 0), T(t)));
  def('PMT', 3, 5, 'nnnnn', (ctx, r, n, v, f, t) => (n === 0 ? E.NUM : pmt(r, n, v, opt(f, 0), T(t))));
  def('NPER', 3, 5, 'nnnnn', (ctx, r, p, v, f, t) => {
    f = opt(f, 0); t = T(t);
    if (r === 0) return p === 0 ? E.NUM : -(v + f) / p;
    const a = p * (1 + r * t) - f * r, b = p * (1 + r * t) + v * r;
    if (a / b <= 0) return E.NUM;
    return Math.log(a / b) / Math.log(1 + r);
  });
  /** RATE: Newton iteration like Excel (20 steps), then a wider search */
  function rate(n, p, v, f, t, guess) {
    const F = (r) => (Math.abs(r) < 1e-12 ? v + p * n + f : v * Math.pow(1 + r, n) + p * (1 + r * t) * (Math.pow(1 + r, n) - 1) / r + f);
    let r = guess;
    for (let i = 0; i < 100; i++) {
      const y = F(r);
      const h = 1e-7 * Math.max(1, Math.abs(r));
      const d = (F(r + h) - F(r - h)) / (2 * h);
      if (!isFinite(d) || d === 0) break;
      const nr = r - y / d;
      if (!isFinite(nr)) break;
      if (Math.abs(nr - r) < 1e-12) return nr;
      r = nr;
      if (r <= -1) r = -0.999999;
    }
    if (Math.abs(F(r)) < 1e-7) return r;
    return NaN;
  }
  def('RATE', 3, 6, 'nnnnnn', (ctx, n, p, v, f, t, g) => { const r = rate(n, p, v, opt(f, 0), T(t), opt(g, 0.1)); return isNaN(r) ? E.NUM : r; });
  function ipmt(r, per, n, v, f, t) {
    const p = pmt(r, n, v, f, t);
    let i;
    if (per === 1) i = t === 1 ? 0 : -v;
    else i = t === 1 ? fv(r, per - 2, p, v, 1) - p : fv(r, per - 1, p, v, 0);
    return i * r;
  }
  def('IPMT', 4, 6, 'nnnnnn', (ctx, r, per, n, v, f, t) => { if (per < 1 || per > n) return E.NUM; return ipmt(r, per, n, v, opt(f, 0), T(t)); });
  def('PPMT', 4, 6, 'nnnnnn', (ctx, r, per, n, v, f, t) => { if (per < 1 || per > n) return E.NUM; f = opt(f, 0); t = T(t); return pmt(r, n, v, f, t) - ipmt(r, per, n, v, f, t); });
  const cum = (prin) => (ctx, r, n, v, s, e, t) => {
    s = trunc(s); e = trunc(e);
    if (r <= 0 || n <= 0 || v <= 0 || s < 1 || e < 1 || s > e || e > n || (t !== 0 && t !== 1)) return E.NUM;
    let tot = 0;
    for (let k = s; k <= e; k++) { const ip = ipmt(r, k, n, v, 0, t); tot += prin ? pmt(r, n, v, 0, t) - ip : ip; }
    return tot;
  };
  def('CUMIPMT', 6, 6, 'nnnnnn', cum(false));
  def('CUMPRINC', 6, 6, 'nnnnnn', cum(true));
  def('ISPMT', 4, 4, 'nnnn', (ctx, r, per, n, v) => (n === 0 ? E.DIV0 : v * r * (per / n - 1)));
  def('EFFECT', 2, 2, 'nn', (ctx, nr, n) => { n = trunc(n); if (nr <= 0 || n < 1) return E.NUM; return Math.pow(1 + nr / n, n) - 1; });
  def('NOMINAL', 2, 2, 'nn', (ctx, ef, n) => { n = trunc(n); if (ef <= 0 || n < 1) return E.NUM; return n * (Math.pow(1 + ef, 1 / n) - 1); });
  def('FVSCHEDULE', 2, 2, 'na', (ctx, p, s) => { let v = p; for (const r of C.toArr(s).d.flat()) { if (r == null) continue; const x = C.toNum(r); if (isErr(x)) return x; v *= 1 + x; } return v; });
  def('PDURATION', 3, 3, 'nnn', (ctx, r, a, b) => (r <= 0 || a <= 0 || b <= 0 ? E.NUM : (Math.log(b) - Math.log(a)) / Math.log(1 + r)));
  def('RRI', 3, 3, 'nnn', (ctx, n, a, b) => (n <= 0 || a === 0 ? E.NUM : Math.pow(b / a, 1 / n) - 1));
  def('DOLLARDE', 2, 2, 'nn', (ctx, d, f) => { f = trunc(f); if (f < 0) return E.NUM; if (f === 0) return E.DIV0; const i = trunc(d); const digits = Math.ceil(Math.log10(f)); return i + ((d - i) * Math.pow(10, digits)) / f; });
  def('DOLLARFR', 2, 2, 'nn', (ctx, d, f) => { f = trunc(f); if (f < 0) return E.NUM; if (f === 0) return E.DIV0; const i = trunc(d); const digits = Math.ceil(Math.log10(f)); return i + ((d - i) * f) / Math.pow(10, digits); });

  /* -------------------------------------------------------------- cash flows */
  def('NPV', 2, 255, 'nr', (ctx, args) => {
    const r = args[0];
    if (isErr(r)) return r;
    const v = C.nums(args.slice(1), ctx);
    if (isErr(v)) return v;
    if (r === -1) return E.DIV0;
    let s = 0;
    for (let i = 0; i < v.length; i++) s += v[i] / Math.pow(1 + r, i + 1);
    return s;
  }, { list: true });
  function cashflows(a, ctx) { const out = []; const e = C.each([a], (v) => { if (isErr(v)) return v; if (typeof v === 'number') out.push(v); }); return e || out; }
  function irr(v, guess) {
    const F = (r) => v.reduce((s, x, i) => s + x / Math.pow(1 + r, i), 0);
    const D = (r) => v.reduce((s, x, i) => s - (i * x) / Math.pow(1 + r, i + 1), 0);
    let r = guess;
    for (let i = 0; i < 100; i++) {
      const f = F(r), d = D(r);
      if (!isFinite(f) || !isFinite(d) || d === 0) break;
      const nr = r - f / d;
      if (Math.abs(nr - r) < 1e-12 * Math.max(1, Math.abs(r))) return nr;
      r = nr <= -1 ? (r - 1) / 2 : nr;
    }
    if (Math.abs(F(r)) < 1e-7) return r;
    /* fallback: bisection on (-0.99, 10) */
    let lo = -0.9999, hi = 10, flo = F(lo), fhi = F(hi);
    if (flo * fhi > 0) return NaN;
    for (let i = 0; i < 300; i++) { const m = (lo + hi) / 2, fm = F(m); if (Math.abs(fm) < 1e-12) return m; if ((fm < 0) === (flo < 0)) { lo = m; flo = fm; } else hi = m; }
    return (lo + hi) / 2;
  }
  def('IRR', 1, 2, 'rn', (ctx, a, g) => {
    const v = cashflows(a, ctx); if (isErr(v)) return v;
    if (!v.some((x) => x > 0) || !v.some((x) => x < 0)) return E.NUM;
    const r = irr(v, opt(g, 0.1));
    return isNaN(r) ? E.NUM : r;
  });
  def('MIRR', 3, 3, 'rnn', (ctx, a, fr, rr) => {
    const v = cashflows(a, ctx); if (isErr(v)) return v;
    const n = v.length;
    let pos = 0, neg = 0;
    v.forEach((x, i) => { if (x > 0) pos += x * Math.pow(1 + rr, n - 1 - i); else neg += x / Math.pow(1 + fr, i); });
    if (!pos || !neg) return E.DIV0;
    return Math.pow(-pos / neg, 1 / (n - 1)) - 1;
  });
  function xflows(vals, dates, ctx) {
    const V = C.toArr(vals).d.flat(), D = C.toArr(dates).d.flat();
    if (V.length !== D.length) return E.NUM;
    const v = [], d = [];
    for (let i = 0; i < V.length; i++) {
      if (isErr(V[i])) return V[i];
      if (isErr(D[i])) return D[i];
      const x = C.toNum(V[i]), y = C.toNum(D[i]);
      if (isErr(x) || isErr(y)) return E.VALUE;
      v.push(x); d.push(Math.floor(y));
    }
    if (d.some((x) => x < d[0])) return E.NUM;
    return { v, d };
  }
  const xnpv = (r, f) => f.v.reduce((s, x, i) => s + x / Math.pow(1 + r, (f.d[i] - f.d[0]) / 365), 0);
  def('XNPV', 3, 3, 'naa', (ctx, r, v, d) => { const f = xflows(v, d, ctx); if (isErr(f)) return f; if (r <= -1) return E.NUM; return xnpv(r, f); });
  def('XIRR', 2, 3, 'aan', (ctx, v, d, g) => {
    const f = xflows(v, d, ctx); if (isErr(f)) return f;
    if (!f.v.some((x) => x > 0) || !f.v.some((x) => x < 0)) return E.NUM;
    let r = opt(g, 0.1);
    const D = (x) => f.v.reduce((s, c, i) => { const t = (f.d[i] - f.d[0]) / 365; return s - (t * c) / Math.pow(1 + x, t + 1); }, 0);
    for (let i = 0; i < 100; i++) {
      const y = xnpv(r, f), dy = D(r);
      if (!isFinite(y) || !dy) break;
      let nr = r - y / dy;
      if (nr <= -1) nr = (r - 1) / 2;
      if (Math.abs(nr - r) < 1e-12) return nr;
      r = nr;
    }
    if (Math.abs(xnpv(r, f)) < 1e-6) return r;
    let lo = -0.9999, hi = 100, flo = xnpv(lo, f);
    if (flo * xnpv(hi, f) > 0) return E.NUM;
    for (let i = 0; i < 400; i++) { const m = (lo + hi) / 2, fm = xnpv(m, f); if (Math.abs(fm) < 1e-10) return m; if ((fm < 0) === (flo < 0)) { lo = m; flo = fm; } else hi = m; }
    return (lo + hi) / 2;
  });

  /* ------------------------------------------------------------- depreciation */
  def('SLN', 3, 3, 'nnn', (ctx, c, s, l) => (l === 0 ? E.DIV0 : (c - s) / l));
  def('SYD', 4, 4, 'nnnn', (ctx, c, s, l, p) => { if (l <= 0 || p <= 0 || p > l || s < 0) return E.NUM; return ((c - s) * (l - p + 1) * 2) / (l * (l + 1)); });
  def('DB', 4, 5, 'nnnnn', (ctx, c, s, l, p, m) => {
    m = trunc(opt(m, 12));
    if (c < 0 || s < 0 || l <= 0 || p <= 0 || m < 1 || m > 12 || p > l + 1 || (p > l && m === 12)) return E.NUM;
    if (c === 0) return 0;
    const r = Math.round((1 - Math.pow(s / c, 1 / l)) * 1000) / 1000;
    let total = (c * r * m) / 12;
    if (trunc(p) === 1) return total;
    let dep = 0;
    for (let i = 2; i <= trunc(p); i++) {
      dep = i === l + 1 ? ((c - total) * r * (12 - m)) / 12 : (c - total) * r;
      total += dep;
    }
    return dep;
  });
  function ddb(c, s, l, p, f) {
    let r = f / l, old;
    if (r >= 1) { r = 1; old = p === 1 ? c : 0; } else old = c * Math.pow(1 - r, p - 1);
    const nw = c * Math.pow(1 - r, p);
    let d = nw < s ? old - s : old - nw;
    return d < 0 ? 0 : d;
  }
  def('DDB', 4, 5, 'nnnnn', (ctx, c, s, l, p, f) => { f = opt(f, 2); if (c < 0 || s < 0 || l <= 0 || p <= 0 || f <= 0 || p > l) return E.NUM; return ddb(c, s, l, p, f); });
  /* VDB (port of the classic algorithm used by Excel-compatible engines) */
  function interVdb(c, s, l, l1, per, f) {
    let v = 0;
    const intEnd = Math.ceil(per - 1e-12);
    let term, sln = 0, salv = c - s, now = false;
    for (let i = 1; i <= intEnd; i++) {
      if (!now) {
        const d = gDdb(c, s, l, i, f);
        sln = salv / (l1 - (i - 1));
        if (sln > d) { term = sln; now = true; } else { term = d; salv -= d; }
      } else term = sln;
      if (i === intEnd) term *= per + 1 - intEnd;
      v += term;
    }
    return v;
  }
  function gDdb(c, s, l, p, f) {
    let rate = f / l, old;
    if (rate >= 1) { rate = 1; old = p === 1 ? c : 0; } else old = c * Math.pow(1 - rate, p - 1);
    const nw = c * Math.pow(1 - rate, p);
    const d = nw < s ? old - s : old - nw;
    return d < 0 ? 0 : d;
  }
  const near = (a, b) => Math.abs(a - b) < 1e-12;
  def('VDB', 5, 7, 'nnnnnnb', (ctx, c, s, l, a, b, f, ns) => {
    f = opt(f, 2); ns = opt(ns, false);
    if (c < 0 || s < 0 || l <= 0 || a < 0 || b < a || b > l || f <= 0) return E.NUM;
    const iS = Math.floor(a + 1e-12), iE = Math.ceil(b - 1e-12);
    let v = 0;
    if (ns) {
      for (let i = iS + 1; i <= iE; i++) {
        let term = gDdb(c, s, l, i, f);
        if (i === iS + 1) term *= Math.min(b, iS + 1) - a;
        else if (i === iE) term *= b + 1 - iE;
        v += term;
      }
      return v;
    }
    let part = 0;
    if (!near(a, iS)) { const tv = c - interVdb(c, s, l, l, iS, f); part += (a - iS) * interVdb(tv, s, l, l - iS, 1, f); }
    if (!near(b, iE)) { const ts = iE - 1; const tv = c - interVdb(c, s, l, l, ts, f); part += (iE - b) * interVdb(tv, s, l, l - ts, iE - ts, f); }
    const c2 = c - interVdb(c, s, l, l, iS, f);
    return interVdb(c2, s, l, l - iS, iE - iS, f) - part;
  });
  def('AMORLINC', 6, 7, 'nnnnnnn', (ctx, cost, d0, first, salv, per, rt, basis) => {
    basis = trunc(opt(basis, 0)); per = trunc(per);
    if (rt <= 0 || salv > cost || cost <= 0 || basis === 2 || basis < 0 || basis > 4) return E.NUM;
    const one = cost * rt, delta = cost - salv;
    const r0 = C.yearfrac(d0, first, basis, ctx) * rt * cost;
    const full = trunc((cost - salv - r0) / one);
    if (per === 0) return r0;
    if (per <= full) return one;
    if (per === full + 1) return delta - one * full - r0;
    return 0;
  });
  def('AMORDEGRC', 6, 7, 'nnnnnnn', (ctx, cost, d0, first, salv, per, rt, basis) => {
    basis = trunc(opt(basis, 0)); per = trunc(per);
    if (rt <= 0 || cost <= 0 || salv > cost || basis === 2 || basis < 0 || basis > 4) return E.NUM;
    const use = 1 / rt;
    const coeff = use < 3 ? 1 : use < 5 ? 1.5 : use <= 6 ? 2 : 2.5;
    const r = rt * coeff;
    let nr = Math.round(C.yearfrac(d0, first, basis, ctx) * r * cost);
    if (per === 0) return nr;
    cost -= nr;
    let rest = cost - salv;
    for (let n = 0; n < per; n++) {
      nr = Math.round(r * cost);
      rest -= nr;
      if (rest < 0) { switch (per - n) { case 0: case 1: return Math.round(cost * 0.5); default: return 0; } }
      cost -= nr;
    }
    return nr;
  });

  /* ------------------------------------------------------------------ bonds */
  const parts = (v, ctx) => NF.serialToParts(v, ctx.wb.date1904, 0);
  const D2S = (y, m, d, ctx) => NF.dateToSerial(y, m, d, ctx.wb.date1904);
  const lastDay = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
  /** coupon date stepping back from maturity by `months` (end-of-month rule) */
  function couponDate(mat, months, ctx) {
    const p = parts(mat, ctx);
    const eom = p.d === lastDay(p.y, p.m);
    let y = p.y, m = p.m + months;
    y += Math.floor((m - 1) / 12);
    m = ((((m - 1) % 12) + 12) % 12) + 1;
    const d = eom ? lastDay(y, m) : Math.min(p.d, lastDay(y, m));
    return D2S(y, m, d, ctx);
  }
  function pcd(set, mat, freq, ctx) {
    const step = 12 / freq;
    let k = 0, d = mat;
    /* jump close first */
    const years = Math.floor((mat - set) / 366);
    k = -years * freq;
    d = couponDate(mat, k * step, ctx);
    while (d <= set) { k++; d = couponDate(mat, k * step, ctx); }
    while (d > set) { k--; d = couponDate(mat, k * step, ctx); }
    return { pcd: d, ncd: couponDate(mat, (k + 1) * step, ctx), k };
  }
  function bondArgs(ctx, set, mat, freq, basis) {
    set = Math.floor(set); mat = Math.floor(mat); freq = trunc(freq); basis = trunc(opt(basis, 0));
    if (set >= mat || ![1, 2, 4].includes(freq) || basis < 0 || basis > 4) return E.NUM;
    return { set, mat, freq, basis };
  }
  const dayCount = (a, b, basis, ctx) => (basis === 0 ? C.days360(a, b, false, ctx) : basis === 4 ? C.days360(a, b, true, ctx) : b - a);
  function coupdays(b, ctx) {
    if (b.basis === 1) { const c = pcd(b.set, b.mat, b.freq, ctx); return c.ncd - c.pcd; }
    return (b.basis === 3 ? 365 : 360) / b.freq;
  }
  const coupdaybs = (b, ctx) => dayCount(pcd(b.set, b.mat, b.freq, ctx).pcd, b.set, b.basis, ctx);
  function coupdaysnc(b, ctx) {
    const c = pcd(b.set, b.mat, b.freq, ctx);
    if (b.basis === 0 || b.basis === 4) return coupdays(b, ctx) - coupdaybs(b, ctx);
    return c.ncd - b.set;
  }
  function coupnum(b, ctx) { const c = pcd(b.set, b.mat, b.freq, ctx); return -c.k; }
  const bondFn = (name, f) => def(name, 3, 4, 'nnnn', (ctx, s, m, fr, ba) => { const b = bondArgs(ctx, s, m, fr, ba); return isErr(b) ? b : f(b, ctx); });
  bondFn('COUPDAYS', coupdays);
  bondFn('COUPDAYBS', coupdaybs);
  bondFn('COUPDAYSNC', coupdaysnc);
  bondFn('COUPNCD', (b, ctx) => pcd(b.set, b.mat, b.freq, ctx).ncd);
  bondFn('COUPPCD', (b, ctx) => pcd(b.set, b.mat, b.freq, ctx).pcd);
  bondFn('COUPNUM', coupnum);
  function price(b, rate, yld, red, ctx) {
    const Ed = coupdays(b, ctx), A = coupdaybs(b, ctx), DSC = coupdaysnc(b, ctx), N = coupnum(b, ctx);
    const f = b.freq, cpn = (100 * rate) / f;
    if (N === 1) return (red + cpn) / (1 + (DSC / Ed) * (yld / f)) - cpn * (A / Ed);
    let p = red / Math.pow(1 + yld / f, N - 1 + DSC / Ed);
    for (let k = 1; k <= N; k++) p += cpn / Math.pow(1 + yld / f, k - 1 + DSC / Ed);
    return p - cpn * (A / Ed);
  }
  def('PRICE', 6, 7, 'nnnnnnn', (ctx, s, m, rate, yld, red, fr, ba) => {
    const b = bondArgs(ctx, s, m, fr, ba); if (isErr(b)) return b;
    if (rate < 0 || yld < 0 || red <= 0) return E.NUM;
    return price(b, rate, yld, red, ctx);
  });
  def('YIELD', 6, 7, 'nnnnnnn', (ctx, s, m, rate, pr, red, fr, ba) => {
    const b = bondArgs(ctx, s, m, fr, ba); if (isErr(b)) return b;
    if (rate < 0 || pr <= 0 || red <= 0) return E.NUM;
    const N = coupnum(b, ctx);
    if (N === 1) {
      const Ed = coupdays(b, ctx), A = coupdaybs(b, ctx), DSR = Ed - A, f = b.freq;
      const cpn = (100 * rate) / f;
      return ((red / 100 + rate / f - (pr / 100 + (A / Ed) * rate / f)) / (pr / 100 + (A / Ed) * rate / f)) * ((f * Ed) / DSR) * 1;
    }
    let lo = -0.99, hi = 1, y = rate || 0.05;
    for (let i = 0; i < 200; i++) {
      const pv0 = price(b, rate, y, red, ctx) - pr;
      const h = 1e-7;
      const d = (price(b, rate, y + h, red, ctx) - price(b, rate, y - h, red, ctx)) / (2 * h);
      const ny = y - pv0 / d;
      if (!isFinite(ny)) break;
      if (Math.abs(ny - y) < 1e-12) return ny;
      y = ny;
    }
    void lo; void hi;
    return Math.abs(price(b, rate, y, red, ctx) - pr) < 1e-7 ? y : E.NUM;
  });
  function duration(b, cpn, yld, ctx) {
    const Ed = coupdays(b, ctx), DSC = coupdaysnc(b, ctx), N = coupnum(b, ctx), f = b.freq;
    const c = (100 * cpn) / f, y = yld / f;
    let num = 0, den = 0;
    for (let k = 1; k <= N; k++) {
      const t = k - 1 + DSC / Ed;
      const cf = c + (k === N ? 100 : 0);
      const pvk = cf / Math.pow(1 + y, t);
      num += t * pvk; den += pvk;
    }
    return num / den / f;
  }
  def('DURATION', 5, 6, 'nnnnnn', (ctx, s, m, cpn, yld, fr, ba) => { const b = bondArgs(ctx, s, m, fr, ba); if (isErr(b)) return b; if (cpn < 0 || yld < 0) return E.NUM; return duration(b, cpn, yld, ctx); });
  def('MDURATION', 5, 6, 'nnnnnn', (ctx, s, m, cpn, yld, fr, ba) => { const b = bondArgs(ctx, s, m, fr, ba); if (isErr(b)) return b; if (cpn < 0 || yld < 0) return E.NUM; return duration(b, cpn, yld, ctx) / (1 + yld / b.freq); });
  const yf = (a, b, basis, ctx) => C.yearfrac(a, b, basis, ctx);
  const basisOK = (b) => b >= 0 && b <= 4;
  def('ACCRINT', 6, 8, 'nnnnnnnb', (ctx, iss, first, set, rate, par, fr, basis, calc) => {
    basis = trunc(opt(basis, 0)); par = opt(par, 1000); fr = trunc(fr);
    if (rate <= 0 || par <= 0 || ![1, 2, 4].includes(fr) || !basisOK(basis) || iss >= set) return E.NUM;
    return par * rate * yf(iss, set, basis, ctx);
  });
  def('ACCRINTM', 3, 5, 'nnnnn', (ctx, iss, set, rate, par, basis) => {
    basis = trunc(opt(basis, 0)); par = opt(par, 1000);
    if (rate <= 0 || par <= 0 || !basisOK(basis) || iss >= set) return E.NUM;
    return par * rate * yf(iss, set, basis, ctx);
  });
  const simple = (name, f) => def(name, 4, 5, 'nnnnn', (ctx, s, m, a, b, basis) => { basis = trunc(opt(basis, 0)); s = Math.floor(s); m = Math.floor(m); if (s >= m || !basisOK(basis)) return E.NUM; return f(s, m, a, b, basis, ctx); });
  simple('DISC', (s, m, pr, red, basis, ctx) => (pr <= 0 || red <= 0 ? E.NUM : (red - pr) / red / yf(s, m, basis, ctx)));
  simple('INTRATE', (s, m, inv, red, basis, ctx) => (inv <= 0 || red <= 0 ? E.NUM : (red - inv) / inv / yf(s, m, basis, ctx)));
  simple('RECEIVED', (s, m, inv, dsc, basis, ctx) => { if (inv <= 0 || dsc <= 0) return E.NUM; const r = 1 - dsc * yf(s, m, basis, ctx); return r <= 0 ? E.NUM : inv / r; });
  simple('PRICEDISC', (s, m, dsc, red, basis, ctx) => (dsc <= 0 || red <= 0 ? E.NUM : red * (1 - dsc * yf(s, m, basis, ctx))));
  simple('YIELDDISC', (s, m, pr, red, basis, ctx) => (pr <= 0 || red <= 0 ? E.NUM : (red - pr) / pr / yf(s, m, basis, ctx)));
  const yearDays = (y, basis) => (basis === 1 ? ((y % 4 === 0 && y % 100 !== 0) || y % 400 === 0 ? 366 : 365) : basis === 3 ? 365 : 360);
  def('PRICEMAT', 5, 6, 'nnnnnn', (ctx, s, m, iss, rate, yld, basis) => {
    basis = trunc(opt(basis, 0));
    if (s >= m || rate < 0 || yld < 0 || !basisOK(basis)) return E.NUM;
    const B = yearDays(parts(s, ctx).y, basis);
    const DIM = dayCount(iss, m, basis, ctx), DSM = dayCount(s, m, basis, ctx), A = dayCount(iss, s, basis, ctx);
    return (100 + (DIM / B) * rate * 100) / (1 + (DSM / B) * yld) - (A / B) * rate * 100;
  });
  def('YIELDMAT', 5, 6, 'nnnnnn', (ctx, s, m, iss, rate, pr, basis) => {
    basis = trunc(opt(basis, 0));
    if (s >= m || rate < 0 || pr <= 0 || !basisOK(basis)) return E.NUM;
    const B = yearDays(parts(s, ctx).y, basis);
    const DIM = dayCount(iss, m, basis, ctx), DSM = dayCount(s, m, basis, ctx), A = dayCount(iss, s, basis, ctx);
    const t = pr / 100 + (A / B) * rate;
    return ((1 + (DIM / B) * rate - t) / t) * (B / DSM);
  });
  const tb = (s, m) => { s = Math.floor(s); m = Math.floor(m); const d = m - s; return s >= m || d > 366 ? null : d; };
  def('TBILLPRICE', 3, 3, 'nnn', (ctx, s, m, dsc) => { const d = tb(s, m); if (d == null || dsc <= 0) return E.NUM; const p = 100 * (1 - (dsc * d) / 360); return p <= 0 ? E.NUM : p; });
  def('TBILLYIELD', 3, 3, 'nnn', (ctx, s, m, pr) => { const d = tb(s, m); if (d == null || pr <= 0) return E.NUM; return ((100 - pr) / pr) * (360 / d); });
  def('TBILLEQ', 3, 3, 'nnn', (ctx, s, m, dsc) => {
    const d = tb(s, m); if (d == null || dsc <= 0) return E.NUM;
    if (d <= 182) return (365 * dsc) / (360 - dsc * d);
    const pr = 100 * (1 - (dsc * d) / 360);
    const a = d / (2 * 365) - 0.25, b = d / 365, c = (pr - 100) / pr;
    return (-b + Math.sqrt(b * b - 4 * a * c)) / (2 * a);
  });
  for (const n of ['ODDFPRICE', 'ODDFYIELD', 'ODDLPRICE', 'ODDLYIELD']) def(n, 0, 255, 'r', (ctx) => { ctx.flags.keep = true; return C.KEEP; }, { list: true });
  void MISSING;
})(typeof window !== 'undefined' ? window : globalThis);
