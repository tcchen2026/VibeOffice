/* Ledger — conditional formatting evaluation and data validation helpers. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, C = L.calc, F = L.formula, NF = L.numfmt;
  const CF = (L.cf = {});
  const isErr = M.isErr;

  /* rule-level caches, rebuilt when the workbook changes */
  let cacheVer = -1, cache = new Map();
  const ruleCache = (wb, rule) => {
    if (cacheVer !== wb.version) { cache = new Map(); cacheVer = wb.version; }
    let o = cache.get(rule);
    if (!o) { o = {}; cache.set(rule, o); }
    return o;
  };
  CF.reset = () => { cacheVer = -1; cache = new Map(); };

  const anchorOf = (cf) => ({ r: cf.ranges[0].r1, c: cf.ranges[0].c1 });
  function parsed(o, key, text) {
    if (o[key] !== undefined) return o[key];
    let a = null;
    try { a = F.parse(String(text).replace(/^=/, '')); } catch (e) { a = null; }
    o[key] = a;
    return a;
  }
  /** evaluate a rule formula for cell (r, c): relative references move with the cell */
  function evalAt(sh, ast, anchor, r, c) {
    if (!ast) return M.ERR.NAME;
    const a = F.shift(ast, r - anchor.r, c - anchor.c);
    const ctx = C.ctx(sh.wb, sh, r, c);
    let v;
    try { v = C.ev(a, ctx); } catch (e) { return M.ERR.VALUE; }
    if (v instanceof C.Ref) v = v.single ? C.cellValue(v.sheet, v.r1, v.c1) : C.intersect(v, ctx);
    else if (v instanceof C.Arr) v = v.d[0][0];
    return v === C.MISSING ? null : v;
  }
  CF.evalAt = evalAt;
  function inRanges(ranges, r, c) { for (const g of ranges) if (r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2) return true; return false; }
  /** every value in the rule's ranges (numbers only unless all) */
  function rangeValues(sh, cf, o, all) {
    const k = all ? 'allVals' : 'nums';
    if (o[k]) return o[k];
    const out = [];
    for (const g of cf.ranges) {
      sh.each(g.r1, g.c1, Math.min(g.r2, sh.maxR), Math.min(g.c2, sh.maxC), (cell, r, c) => {
        const v = C.cellValue(sh, r, c);
        if (all) { if (v != null && v !== '') out.push(v); } else if (typeof v === 'number') out.push(v);
      });
    }
    o[k] = out;
    return out;
  }
  const cmp = (a, b) => {
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    if (typeof a === 'string' && typeof b === 'string') return C.strCmp(a, b);
    return C.compare ? C.compare(a, b) : 0;
  };
  function norm(v) { return v == null ? 0 : v; }
  /** threshold value of a cfvo (colour scale / data bar / icon set) */
  function cfvoValue(sh, cf, o, v, nums) {
    const sorted = o.sorted || (o.sorted = nums.slice().sort((a, b) => a - b));
    const min = sorted[0], max = sorted[sorted.length - 1];
    const num = (x) => { if (x == null) return 0; const n = typeof x === 'number' ? x : parseFloat(x); if (!isNaN(n)) return n; const a = parsed(o, 'cfvo:' + x, x); const r = evalAt(sh, a, anchorOf(cf), anchorOf(cf).r, anchorOf(cf).c); return typeof r === 'number' ? r : 0; };
    switch (v.type) {
      case 'min': return min;
      case 'max': return max;
      case 'num': case 'formula': return num(v.val);
      case 'percent': { const p = num(v.val); return min + ((max - min) * p) / 100; }
      case 'percentile': {
        const p = num(v.val) / 100;
        const n = sorted.length;
        if (!n) return 0;
        const x = p * (n - 1), i = Math.floor(x);
        return i + 1 < n ? sorted[i] + (x - i) * (sorted[i + 1] - sorted[i]) : sorted[n - 1];
      }
      default: return 0;
    }
  }
  const lerp = (a, b, t) => Math.round(a + (b - a) * t);
  function mix(c1, c2, t) {
    const a = parseInt(c1.slice(1), 16), b = parseInt(c2.slice(1), 16);
    const r = lerp((a >> 16) & 255, (b >> 16) & 255, t), g = lerp((a >> 8) & 255, (b >> 8) & 255, t), bl = lerp(a & 255, b & 255, t);
    return '#' + ((1 << 24) | (r << 16) | (g << 8) | bl).toString(16).slice(1).toUpperCase();
  }
  const TODAY = () => { const d = new Date(); return NF.dateToSerial(d.getFullYear(), d.getMonth() + 1, d.getDate(), false); };

  /** does a (non-graphical) rule apply to cell (r, c)? */
  function ruleTrue(sh, cf, rule, o, r, c, v) {
    const anchor = anchorOf(cf);
    switch (rule.type) {
      case 'cellIs': {
        if (v == null) v = 0;
        const a = evalAt(sh, parsed(o, 'f0', (rule.f || [])[0] || '0'), anchor, r, c);
        if (isErr(v) || isErr(a)) return false;
        const op = rule.op || 'equal';
        if (op === 'between' || op === 'notBetween') {
          const b = evalAt(sh, parsed(o, 'f1', (rule.f || [])[1] || '0'), anchor, r, c);
          if (isErr(b)) return false;
          const lo = cmp(a, b) <= 0 ? a : b, hi = lo === a ? b : a;
          const inside = cmp(v, norm(lo)) >= 0 && cmp(v, norm(hi)) <= 0;
          return op === 'between' ? inside : !inside;
        }
        const d = typeof v === typeof norm(a) ? cmp(v, norm(a)) : (C.compare ? C.compare(v, norm(a)) : NaN);
        switch (op) {
          case 'equal': return typeof v === 'string' && typeof a === 'string' ? C.strEq(v, a) : d === 0;
          case 'notEqual': return !(typeof v === 'string' && typeof a === 'string' ? C.strEq(v, a) : d === 0);
          case 'greaterThan': return d > 0;
          case 'lessThan': return d < 0;
          case 'greaterThanOrEqual': return d >= 0;
          case 'lessThanOrEqual': return d <= 0;
          default: return false;
        }
      }
      case 'expression': case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': case 'timePeriod': case 'containsBlanks': case 'notContainsBlanks': case 'containsErrors': case 'notContainsErrors': {
        if (rule.type === 'containsBlanks') return v == null || (typeof v === 'string' && v.trim() === '');
        if (rule.type === 'notContainsBlanks') return !(v == null || (typeof v === 'string' && v.trim() === ''));
        if (rule.type === 'containsErrors') return isErr(v);
        if (rule.type === 'notContainsErrors') return !isErr(v);
        if (rule.type !== 'expression' && rule.text != null && !(rule.f && rule.f.length)) {
          const s = v == null ? '' : String(isErr(v) ? v.e : v).toLowerCase(), t = String(rule.text).toLowerCase();
          if (rule.type === 'containsText') return s.indexOf(t) >= 0;
          if (rule.type === 'notContainsText') return s.indexOf(t) < 0;
          if (rule.type === 'beginsWith') return s.startsWith(t);
          if (rule.type === 'endsWith') return s.endsWith(t);
        }
        if (rule.type === 'timePeriod' && !(rule.f && rule.f.length)) return timePeriod(rule.period, v);
        const res = evalAt(sh, parsed(o, 'f0', (rule.f || [])[0] || 'FALSE'), anchor, r, c);
        if (isErr(res)) return false;
        if (typeof res === 'number') return res !== 0;
        if (typeof res === 'boolean') return res;
        if (typeof res === 'string') { const u = res.toUpperCase(); return u === 'TRUE'; }
        return false;
      }
      case 'duplicateValues': case 'uniqueValues': {
        if (!o.counts) {
          o.counts = new Map();
          for (const x of rangeValues(sh, cf, o, true)) { const k = typeof x === 'string' ? 's' + x.toLowerCase() : typeof x + String(isErr(x) ? x.e : x); o.counts.set(k, (o.counts.get(k) || 0) + 1); }
        }
        if (v == null || v === '') return false;
        const k = typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + String(isErr(v) ? v.e : v);
        const n = o.counts.get(k) || 0;
        return rule.type === 'duplicateValues' ? n > 1 : n === 1;
      }
      case 'top10': {
        if (typeof v !== 'number') return false;
        const nums = rangeValues(sh, cf, o, false);
        if (!nums.length) return false;
        if (o.cut == null) {
          const s = nums.slice().sort((a, b) => (rule.bottom ? a - b : b - a));
          let n = rule.rank == null ? 10 : rule.rank;
          if (rule.percent) n = Math.max(1, Math.floor((s.length * n) / 100));
          n = Math.max(1, Math.min(n, s.length));
          o.cut = s[n - 1];
        }
        return rule.bottom ? v <= o.cut : v >= o.cut;
      }
      case 'aboveAverage': {
        if (typeof v !== 'number') return false;
        const nums = rangeValues(sh, cf, o, false);
        if (!nums.length) return false;
        if (o.avg == null) { o.avg = nums.reduce((a, b) => a + b, 0) / nums.length; o.sd = Math.sqrt(nums.reduce((a, b) => a + (b - o.avg) * (b - o.avg), 0) / nums.length); }
        const lim = o.avg + (rule.stdDev ? (rule.below ? -1 : 1) * rule.stdDev * o.sd : 0);
        if (rule.below) return rule.equal ? v <= lim : v < lim;
        return rule.equal ? v >= lim : v > lim;
      }
      default: return false;
    }
  }
  function timePeriod(p, v) {
    if (typeof v !== 'number') return false;
    const t = TODAY(), d = Math.floor(v);
    const parts = (x) => NF.serialToParts(x, false);
    const tp = parts(t), dp = parts(d);
    const mIndex = (q) => q.y * 12 + q.m;
    const wkStart = t - ((tp.wd + 7) % 7);
    switch (p) {
      case 'today': return d === t;
      case 'yesterday': return d === t - 1;
      case 'tomorrow': return d === t + 1;
      case 'last7Days': return d <= t && d > t - 7;
      case 'thisMonth': return mIndex(dp) === mIndex(tp);
      case 'lastMonth': return mIndex(dp) === mIndex(tp) - 1;
      case 'nextMonth': return mIndex(dp) === mIndex(tp) + 1;
      case 'thisWeek': return d >= wkStart && d < wkStart + 7;
      case 'lastWeek': return d >= wkStart - 7 && d < wkStart;
      case 'nextWeek': return d >= wkStart + 7 && d < wkStart + 14;
      default: return false;
    }
  }

  /** index the sheet's rules by bounding boxes for quick lookup */
  function rulesFor(sh) {
    if (sh._cfIdx && sh._cfIdx.ver === sh.wb.version && sh._cfIdx.n === sh.cf.length) return sh._cfIdx.list;
    const list = [];
    for (const cf of sh.cf) for (const rule of cf.rules) list.push({ cf, rule });
    list.sort((a, b) => (a.rule.priority || 0) - (b.rule.priority || 0));
    sh._cfIdx = { ver: sh.wb.version, n: sh.cf.length, list };
    return list;
  }

  /**
   * Conditional formatting result for one cell, or null.
   * {dxf: style delta, scale: '#hex', bar: {...}, icon: {set, i, n, showValue}}
   */
  CF.at = function (sh, r, c, cell) {
    if (!sh.cf.length) return null;
    const rules = rulesFor(sh);
    let out = null;
    const v = cell ? (cell.dirty ? C.cellValue(sh, r, c) : cell.v) : null;
    for (const { cf, rule } of rules) {
      if (!inRanges(cf.ranges, r, c)) continue;
      const o = ruleCache(sh.wb, rule);
      if (rule.scale || rule.type === 'colorScale') {
        if (typeof v !== 'number' || !rule.scale) continue;
        const nums = rangeValues(sh, cf, o, false);
        if (!nums.length) continue;
        const cols = rule.scale.colors.map((x) => M.colorHex(sh.wb, x, '#FFFFFF'));
        const th = rule.scale.cfvo.map((x) => cfvoValue(sh, cf, o, x, nums));
        let hex;
        if (th.length === 2) { const t = th[1] === th[0] ? 0.5 : (v - th[0]) / (th[1] - th[0]); hex = mix(cols[0], cols[1], Math.max(0, Math.min(1, t))); }
        else { if (v <= th[1]) { const t = th[1] === th[0] ? 1 : (v - th[0]) / (th[1] - th[0]); hex = mix(cols[0], cols[1], Math.max(0, Math.min(1, t))); } else { const t = th[2] === th[1] ? 1 : (v - th[1]) / (th[2] - th[1]); hex = mix(cols[1], cols[2], Math.max(0, Math.min(1, t))); } }
        out = out || {};
        if (!out.scale) out.scale = hex;
        if (rule.stop) break;
        continue;
      }
      if (rule.bar || rule.type === 'dataBar') {
        if (typeof v !== 'number' || !rule.bar) continue;
        const nums = rangeValues(sh, cf, o, false);
        if (!nums.length) continue;
        const b = rule.bar;
        let lo = cfvoValue(sh, cf, o, b.cfvo[0] || { type: 'min' }, nums), hi = cfvoValue(sh, cf, o, b.cfvo[1] || { type: 'max' }, nums);
        if (lo > hi) [lo, hi] = [hi, lo];
        out = out || {};
        if (!out.bar) {
          const minL = (b.minLength == null ? 10 : b.minLength) / 100, maxL = (b.maxLength == null ? 90 : b.maxLength) / 100;
          let frac;
          let neg = false, axis = null;
          if (b.gradient === undefined && !b.negColor) {
            /* Excel 2007 bars: length from min to max */
            frac = hi === lo ? maxL : minL + (maxL - minL) * Math.max(0, Math.min(1, (v - lo) / (hi - lo)));
          } else {
            /* Excel 2010 bars: negative values run left of an axis */
            const allPos = lo >= 0, allNeg = hi <= 0;
            if (allPos) frac = hi === 0 ? 0 : Math.max(minL, Math.min(1, v / hi)) * maxL / maxL;
            else if (allNeg) { neg = true; frac = lo === 0 ? 0 : Math.min(1, v / lo); }
            else { axis = -lo / (hi - lo); neg = v < 0; frac = v < 0 ? (v / lo) * axis : (v / hi) * (1 - axis); }
            if (allPos) frac = hi === lo ? 1 : Math.max(0, Math.min(1, (v - Math.min(0, lo)) / (hi - Math.min(0, lo))));
          }
          out.bar = { frac, neg, axis, color: M.colorHex(sh.wb, b.color, '#638EC6'), negColor: b.negColor ? M.colorHex(sh.wb, b.negColor, '#FF0000') : '#FF0000', gradient: b.gradient !== false, border: b.border, borderColor: b.borderColor ? M.colorHex(sh.wb, b.borderColor) : null, showValue: b.showValue !== false, dir: b.direction };
        }
        if (rule.stop) break;
        continue;
      }
      if (rule.icons || rule.type === 'iconSet') {
        if (typeof v !== 'number' || !rule.icons) continue;
        const nums = rangeValues(sh, cf, o, false);
        if (!nums.length) continue;
        const ic = rule.icons;
        const th = ic.cfvo.map((x) => (x.type === 'percent' && ic.percent !== false ? cfvoValue(sh, cf, o, x, nums) : cfvoValue(sh, cf, o, x, nums)));
        let idx = 0;
        for (let k = th.length - 1; k >= 1; k--) {
          const gte = ic.cfvo[k].gte !== false;
          if (gte ? v >= th[k] : v > th[k]) { idx = k; break; }
        }
        const n = th.length;
        out = out || {};
        if (!out.icon) {
          const cu = ic.custom && ic.custom[idx];
          if (cu && cu.set === 'NoIcons') { if (ic.showValue === false) out.icon = { set: 'NoIcons', i: 0, n, showValue: false }; }
          else if (cu && cu.set) out.icon = { set: cu.set, i: cu.id, n, showValue: ic.showValue !== false };
          else out.icon = { set: ic.set, i: ic.reverse ? n - 1 - idx : idx, n, showValue: ic.showValue !== false };
        }
        if (rule.stop) break;
        continue;
      }
      if (!ruleTrue(sh, cf, rule, o, r, c, v)) continue;
      out = out || {};
      if (rule.dxf != null && sh.wb.dxfs[rule.dxf]) {
        const d = sh.wb.dxfs[rule.dxf];
        out.dxf = out.dxf ? mergeDxf(d, out.dxf) : d;
      }
      if (rule.stop) break;
    }
    return out;
  };
  /** combine differential formats: properties set by higher-priority rules (b) win */
  function mergeDxf(lower, higher) {
    const o = {};
    for (const k of ['nf', 'fill']) { if (higher[k] != null) o[k] = higher[k]; else if (lower[k] != null) o[k] = lower[k]; }
    if (lower.font || higher.font) o.font = Object.assign({}, lower.font || {}, higher.font || {});
    if (lower.border || higher.border) o.border = Object.assign({}, lower.border || {}, higher.border || {});
    return o;
  }

  /* ------------------------------------------------------------ data validation */
  CF.dvAt = function (sh, r, c) {
    for (let i = sh.dv.length - 1; i >= 0; i--) { const d = sh.dv[i]; if (inRanges(d.ranges, r, c)) return d; }
    return null;
  };
  /** list items of a list validation */
  CF.dvList = function (sh, dv, r, c) {
    const f = dv.f1 || '';
    if (/^".*"$/.test(f.trim())) return f.trim().slice(1, -1).split(',').map((s) => s.trim());
    const anchor = { r: dv.ranges[0].r1, c: dv.ranges[0].c1 };
    let a; try { a = F.parse(f.replace(/^=/, '')); } catch (e) { return []; }
    const sh2 = sh;
    const v = C.ev(F.shift(a, r - anchor.r, c - anchor.c), C.ctx(sh.wb, sh2, r, c, true));
    const arr = v instanceof C.Ref ? C.refToArr(v) : v instanceof C.Arr ? v : null;
    if (!arr) return v == null || isErr(v) ? [] : String(v).split(',');
    const out = [];
    for (const row of arr.d) for (const x of row) if (x != null && x !== '') out.push(x);
    return out;
  };
  /** check a value against a cell's validation; returns null when valid, else the rule */
  CF.dvCheck = function (sh, r, c, v) {
    const dv = CF.dvAt(sh, r, c);
    if (!dv || dv.type === 'none' || dv.type === 'any') return null;
    if (v == null || v === '') return dv.allowBlank ? null : null;
    const anchor = { r: dv.ranges[0].r1, c: dv.ranges[0].c1 };
    const ev = (f) => { if (f == null) return null; let a; try { a = F.parse(String(f).replace(/^=/, '')); } catch (e) { return null; } return evalAt(sh, a, anchor, r, c); };
    if (dv.type === 'list') {
      const items = CF.dvList(sh, dv, r, c);
      const ok = items.some((x) => (typeof x === 'number' && typeof v === 'number' ? x === v : String(x).toLowerCase() === String(v).toLowerCase()));
      return ok ? null : dv;
    }
    if (dv.type === 'custom') { const res = ev(dv.f1); return res === true || (typeof res === 'number' && res !== 0) ? null : dv; }
    let x = v;
    if (dv.type === 'textLength') x = String(v).length;
    else if (typeof v !== 'number') return dv;
    if (dv.type === 'whole' && Math.floor(x) !== x) return dv;
    const a = ev(dv.f1), b = ev(dv.f2);
    const A = typeof a === 'number' ? a : parseFloat(a), B = typeof b === 'number' ? b : parseFloat(b);
    const op = dv.op || 'between';
    let ok = true;
    switch (op) {
      case 'between': ok = x >= Math.min(A, B) && x <= Math.max(A, B); break;
      case 'notBetween': ok = !(x >= Math.min(A, B) && x <= Math.max(A, B)); break;
      case 'equal': ok = x === A; break;
      case 'notEqual': ok = x !== A; break;
      case 'greaterThan': ok = x > A; break;
      case 'lessThan': ok = x < A; break;
      case 'greaterThanOrEqual': ok = x >= A; break;
      case 'lessThanOrEqual': ok = x <= A; break;
      default: ok = true;
    }
    return ok ? null : dv;
  };
})(typeof window !== 'undefined' ? window : globalThis);
