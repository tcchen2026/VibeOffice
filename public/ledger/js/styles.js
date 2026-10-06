/* Ledger — list (table) styles: the built-in Light / Medium / Dark table styles drawn from the theme colours,
 * and custom table styles stored in the workbook. L.tstyle.at(sheet, r, c) → {fill, font, border} | null */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model;
  const TS = (L.tstyle = {});
  const solid = (hex) => ({ pattern: 'solid', fg: M.rgb(hex) });
  const line = (hex, style) => ({ style: style || 'thin', color: M.rgb(hex) });
  const tint = (hex, t) => M.applyTint(hex, t);
  const cache = new Map();
  /** theme colour for accent index k (0 = the dark text colour) */
  function accent(wb, k) {
    const th = (wb && wb.theme && wb.theme.colors) || M.DEFAULT_THEME.colors;
    return k === 0 ? th[1] || '#000000' : th[3 + k] || M.DEFAULT_THEME.colors[3 + k];
  }
  /** built-in style spec by name */
  function builtin(wb, name) {
    const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
    if (!m) return null;
    const kind = m[1], n = +m[2];
    const k = (n - 1) % 7;
    const a = accent(wb, k);
    const grayish = k === 0;
    const pale = grayish ? '#D9D9D9' : tint(a, 0.8), pale2 = grayish ? '#BFBFBF' : tint(a, 0.6), mid = grayish ? '#808080' : tint(a, 0.4);
    const W = '#FFFFFF';
    if (kind === 'Light') {
      if (n <= 7) return { whole: { border: { t: line(a), b: line(a) } }, header: { font: { b: true, color: grayish ? undefined : M.rgb(tint(a, -0.25)) }, border: { b: line(a) } }, total: { font: { b: true }, border: { t: line(a) } }, band1: { fill: solid(pale), font: { color: grayish ? undefined : M.rgb(tint(a, -0.25)) } }, band2: { font: { color: grayish ? undefined : M.rgb(tint(a, -0.25)) } }, firstCol: { font: { b: true } }, lastCol: { font: { b: true } } };
      if (n <= 14) return { whole: { border: { t: line(a), b: line(a), l: line(a), r: line(a) } }, header: { fill: solid(a), font: { b: true, color: M.rgb(W) } }, total: { font: { b: true }, border: { t: line(a, 'double') } }, band1: { border: { t: line(a), b: line(a) } }, firstCol: { font: { b: true } }, lastCol: { font: { b: true } } };
      return { whole: { border: { t: line(a), b: line(a), l: line(a), r: line(a) } }, inner: line(a), header: { font: { b: true }, border: { b: line(a, 'medium') } }, total: { font: { b: true }, border: { t: line(a, 'double') } }, band1: { fill: solid(pale) }, firstCol: { font: { b: true } }, lastCol: { font: { b: true } } };
    }
    if (kind === 'Medium') {
      if (n <= 7) return { whole: { border: { t: line(mid), b: line(mid), l: line(mid), r: line(mid) } }, innerH: line(mid), header: { fill: solid(a), font: { b: true, color: M.rgb(W) } }, total: { font: { b: true }, border: { t: line(a, 'double') } }, band1: { fill: solid(pale) }, firstCol: { font: { b: true } }, lastCol: { font: { b: true } } };
      if (n <= 14) return { whole: { fill: solid(pale), border: { innerH: true } }, innerH: line(W), innerV: line(W), header: { fill: solid(a), font: { b: true, color: M.rgb(W) }, border: { b: line(W, 'medium') } }, total: { fill: solid(a), font: { b: true, color: M.rgb(W) }, border: { t: line(W, 'medium') } }, band1: { fill: solid(pale2) }, firstCol: { fill: solid(a), font: { b: true, color: M.rgb(W) } }, lastCol: { fill: solid(a), font: { b: true, color: M.rgb(W) } } };
      if (n <= 21) return { whole: { border: { t: line('#000000', 'medium'), b: line('#000000', 'medium') } }, header: { fill: solid(a), font: { b: true, color: M.rgb(W) }, border: { b: line('#000000', 'medium') } }, total: { font: { b: true }, border: { t: line('#000000', 'double') } }, band1: { fill: solid(pale2) }, firstCol: { font: { b: true } }, lastCol: { font: { b: true } } };
      return { whole: { fill: solid(pale), border: { t: line(mid), b: line(mid), l: line(mid), r: line(mid) } }, inner: line(mid), header: { font: { b: true } }, total: { font: { b: true }, border: { t: line(a, 'double') } }, band1: { fill: solid(pale2) }, firstCol: { font: { b: true } }, lastCol: { font: { b: true } } };
    }
    /* Dark */
    if (n <= 7) { const dark = grayish ? '#404040' : tint(a, -0.25); return { whole: { fill: solid(a), font: { color: M.rgb(W) } }, header: { fill: solid('#000000'), font: { b: true, color: M.rgb(W) }, border: { b: line(W, 'medium') } }, total: { fill: solid(tint(a, -0.5)), font: { b: true, color: M.rgb(W) }, border: { t: line(W, 'double') } }, band1: { fill: solid(dark) }, firstCol: { fill: solid(dark), font: { b: true, color: M.rgb(W) } }, lastCol: { fill: solid(dark), font: { b: true, color: M.rgb(W) } } }; }
    const pairs = [[accent(wb, 0), accent(wb, 1)], [accent(wb, 1), accent(wb, 2)], [accent(wb, 3), accent(wb, 4)], [accent(wb, 5), accent(wb, 6)]];
    const [p1, p2] = pairs[(n - 8) % 4];
    return { whole: { fill: solid(tint(p2, 0.8)) }, header: { fill: solid(p1), font: { b: true, color: M.rgb(W) } }, total: { fill: solid(tint(p2, 0.8)), font: { b: true }, border: { t: line('#000000', 'double') } }, band1: { fill: solid(tint(p2, 0.6)) }, firstCol: { font: { b: true } }, lastCol: { font: { b: true } } };
  }
  /** custom table style from the workbook's tableStyles (dxf-based) */
  function custom(wb, name) {
    const ts = wb.tableStyles && wb.tableStyles.custom && wb.tableStyles.custom.find((t) => t.name === name);
    if (!ts) return null;
    const map = { wholeTable: 'whole', headerRow: 'header', totalRow: 'total', firstRowStripe: 'band1', secondRowStripe: 'band2', firstColumn: 'firstCol', lastColumn: 'lastCol', firstColumnStripe: 'colBand1', secondColumnStripe: 'colBand2' };
    const spec = {};
    for (const e of ts.elements) {
      const key = map[e.type];
      if (!key || e.dxf < 0) continue;
      const d = wb.dxfs[e.dxf];
      if (!d) continue;
      const part = {};
      if (d.fill && (d.fill.fg || d.fill.bg)) part.fill = { pattern: 'solid', fg: d.fill.fg || d.fill.bg };
      if (d.font) part.font = d.font;
      if (d.border) { part.border = {}; for (const k of ['t', 'b', 'l', 'r']) if (d.border[k]) part.border[k] = d.border[k]; if (key === 'whole') { if (d.border.h) spec.innerH = d.border.h; if (d.border.v) spec.innerV = d.border.v; } }
      spec[key] = part;
    }
    return spec;
  }
  function specFor(wb, name) {
    const k = wb.id + '|' + name;
    if (cache.has(k)) return cache.get(k);
    const s = custom(wb, name) || builtin(wb, name);
    cache.set(k, s);
    return s;
  }
  TS.reset = () => cache.clear();
  const merge = (out, part) => {
    if (!part) return;
    if (part.fill) out.fill = part.fill;
    if (part.font) out.font = Object.assign({}, out.font || {}, part.font);
    if (part.border) { out.border = Object.assign({}, out.border || {}); for (const k of ['t', 'b', 'l', 'r']) if (part.border[k]) out.border[k] = part.border[k]; }
  };
  TS.at = function (sh, r, c) {
    const wb = sh.wb;
    if (!wb.tables.length) return null;
    const t = wb.tableAt(sh, r, c);
    if (!t || !t.style || !t.style.name) return null;
    const spec = specFor(wb, t.style.name);
    if (!spec) return null;
    const ref = t.ref;
    const hasHead = t.header !== false;
    const header = hasHead && r === ref.r1;
    const total = !!t.totals && r === ref.r2;
    const out = {};
    /* whole-table outline and inner lines */
    if (spec.whole) {
      const w = spec.whole;
      if (w.fill) out.fill = w.fill;
      if (w.font) out.font = Object.assign({}, w.font);
      const b = {};
      if (w.border) { if (r === ref.r1 && w.border.t) b.t = w.border.t; if (r === ref.r2 && w.border.b) b.b = w.border.b; if (c === ref.c1 && w.border.l) b.l = w.border.l; if (c === ref.c2 && w.border.r) b.r = w.border.r; }
      const ih = spec.innerH || spec.inner, iv = spec.innerV || spec.inner;
      if (ih && r < ref.r2) b.b = b.b || ih;
      if (ih && r > ref.r1) b.t = b.t || ih;
      if (iv && c < ref.c2) b.r = b.r || iv;
      if (iv && c > ref.c1) b.l = b.l || iv;
      if (Object.keys(b).length) out.border = b;
    }
    if (!header && !total) {
      const st = t.style;
      const di = r - ref.r1 - (hasHead ? 1 : 0);
      if (st.rowStripes) merge(out, di % 2 === 0 ? spec.band1 : spec.band2);
      if (st.colStripes) merge(out, (c - ref.c1) % 2 === 0 ? spec.colBand1 : spec.colBand2);
      if (st.first && c === ref.c1) merge(out, spec.firstCol);
      if (st.last && c === ref.c2) merge(out, spec.lastCol);
    }
    if (header) merge(out, spec.header);
    if (total) merge(out, spec.total);
    return out.fill || out.font || out.border ? out : null;
  };
  TS.NAMES = [].concat(Array.from({ length: 21 }, (_, i) => 'TableStyleLight' + (i + 1)), Array.from({ length: 28 }, (_, i) => 'TableStyleMedium' + (i + 1)), Array.from({ length: 11 }, (_, i) => 'TableStyleDark' + (i + 1)));
})(typeof window !== 'undefined' ? window : globalThis);
