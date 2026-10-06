/* Ledger — sheet geometry: column / row positions, fonts and text measurement, cell display text. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, NF = L.numfmt;
  const LY = (L.layout = {});

  /* ------------------------------------------------------------ axes */
  /** Positions along one axis: a default size plus sparse overrides. Sizes are in unzoomed CSS pixels. */
  class Axis {
    constructor(def, max, overrides) {
      this.def = def;
      this.max = max;
      this.idx = [];
      this.sz = [];
      this.cum = [0]; /* cum[k] = Σ (sz - def) for overrides 0..k-1 */
      const keys = Array.from(overrides.keys()).sort((a, b) => a - b);
      for (const k of keys) { const s = overrides.get(k); if (s === def) continue; this.idx.push(k); this.sz.push(s); this.cum.push(this.cum[this.cum.length - 1] + (s - def)); }
    }
    /** index of the first override >= i */
    lb(i) { let lo = 0, hi = this.idx.length; while (lo < hi) { const m = (lo + hi) >> 1; if (this.idx[m] < i) lo = m + 1; else hi = m; } return lo; }
    size(i) { const k = this.lb(i); return k < this.idx.length && this.idx[k] === i ? this.sz[k] : this.def; }
    pos(i) { return i * this.def + this.cum[this.lb(i)]; }
    end(i) { return this.pos(i + 1); }
    get total() { return this.pos(this.max); }
    /** index containing offset p (clamped) */
    at(p) {
      if (p <= 0) return 0;
      let lo = 0, hi = this.max - 1;
      while (lo < hi) { const m = (lo + hi + 1) >> 1; if (this.pos(m) <= p) lo = m; else hi = m - 1; }
      return lo;
    }
    /** next visible (non-zero size) index from i in direction d */
    visible(i, d) {
      while (i >= 0 && i < this.max && this.size(i) === 0) i += d;
      return Math.max(0, Math.min(this.max - 1, i));
    }
  }
  LY.Axis = Axis;

  /** geometry cache per sheet; call LY.invalidate(sh) after width / height / hidden changes */
  LY.invalidate = (sh) => { if (sh) sh._geo = null; };
  LY.geo = function (sh) {
    if (sh._geo) return sh._geo;
    const dc = M.defaultColPx(sh);
    const co = new Map();
    sh.cols.forEach((o, c) => { if (o) co.set(c, M.colPx(sh, c)); });
    if (sh.colTail) { const px = M.colPx(sh, sh.colTail.from); if (px !== dc) for (let c = sh.colTail.from; c < M.MAXC; c++) if (!sh.cols[c]) co.set(c, px); }
    const dr = M.ptToPx(M.defaultRowPt(sh));
    const ro = new Map();
    sh.rows.forEach((row, r) => { if (row && (row.hidden || row.ht != null || row.auto != null)) ro.set(r, M.ptToPx(M.rowPt(sh, r))); });
    if (sh.zeroHeight) { /* rows without an explicit height are hidden */ }
    sh._geo = { cols: new Axis(dc, M.MAXC, co), rows: new Axis(sh.zeroHeight ? 0 : dr, M.MAXR, ro), defRow: dr, defCol: dc };
    return sh._geo;
  };

  /* ------------------------------------------------------------ fonts */
  /* metric-compatible stand-ins served from Google Fonts */
  const STACK = {
    calibri: 'Carlito, Calibri, "Segoe UI", Arial, sans-serif',
    'calibri light': 'Carlito, "Calibri Light", Calibri, sans-serif',
    arial: 'Arimo, Arial, "Liberation Sans", Helvetica, sans-serif',
    'arial narrow': '"Arial Narrow", Arimo, Arial, sans-serif',
    helvetica: 'Arimo, Helvetica, Arial, sans-serif',
    'times new roman': 'Tinos, "Times New Roman", "Liberation Serif", Times, serif',
    times: 'Tinos, Times, "Times New Roman", serif',
    'courier new': 'Cousine, "Courier New", "Liberation Mono", monospace',
    courier: 'Cousine, Courier, monospace',
    cambria: 'Caladea, Cambria, Georgia, serif',
    'liberation sans': 'Arimo, "Liberation Sans", Arial, sans-serif',
    'liberation serif': 'Tinos, "Liberation Serif", serif',
    'liberation mono': 'Cousine, "Liberation Mono", monospace',
    verdana: 'Verdana, "DejaVu Sans", Arimo, sans-serif',
    tahoma: 'Tahoma, Verdana, Arimo, sans-serif',
    'segoe ui': '"Segoe UI", Arimo, sans-serif',
    georgia: 'Georgia, Tinos, serif',
    'comic sans ms': '"Comic Sans MS", "Comic Neue", cursive',
    aptos: 'Aptos, Carlito, Calibri, sans-serif',
    'aptos narrow': '"Aptos Narrow", Carlito, Calibri, sans-serif',
  };
  LY.fontStack = (name) => {
    const k = String(name || 'Arial').toLowerCase();
    if (STACK[k]) return STACK[k];
    return '"' + String(name).replace(/"/g, '') + '", Arimo, Arial, sans-serif';
  };
  /** canvas font string for a style font at zoom z (sizes in points) */
  LY.cssFont = (f, z, scale) => {
    const px = ((f.sz || 11) * 96) / 72 * (z || 1) * (scale || 1);
    return (f.i ? 'italic ' : '') + (f.b ? 'bold ' : '') + px.toFixed(2) + 'px ' + LY.fontStack(f.name);
  };
  let mctx = null;
  const measureCache = new Map();
  LY.ctx = () => {
    if (!mctx) {
      if (typeof OffscreenCanvas !== 'undefined') mctx = new OffscreenCanvas(4, 4).getContext('2d');
      else if (typeof document !== 'undefined') mctx = document.createElement('canvas').getContext('2d');
    }
    return mctx;
  };
  /** text width in CSS px for a canvas font string */
  LY.measure = function (font, text) {
    const c = LY.ctx();
    if (!c) return String(text).length * 7;
    const key = font + '\u0001' + text;
    let w = measureCache.get(key);
    if (w == null) {
      if (measureCache.size > 50000) measureCache.clear();
      c.font = font;
      w = c.measureText(text).width;
      measureCache.set(key, w);
    }
    return w;
  };
  LY.clearMeasure = () => measureCache.clear();
  /** width as Excel's GDI text layout sees it: every glyph advance hinted to whole pixels (digits of Calibri 11
   *  are 7 px there, 7.4 px in a browser), so a value that fits a column in Excel fits here too */
  LY.hintedWidth = function (font, text) {
    let w = 0;
    for (const ch of String(text)) w += Math.floor(LY.measure(font, ch) + 0.05);
    return w;
  };
  /* maximum digit width of a font: the unit of column widths */
  L.measureDigit = function (f) {
    if (!LY.ctx()) return null;
    const font = LY.cssFont(f, 1);
    let w = 0;
    for (const d of '0123456789') w = Math.max(w, LY.measure(font, d));
    return Math.max(5, Math.round(w));
  };

  /* ------------------------------------------------------------ styles */
  /** effective style of a cell (cell style, else row style, else column style) */
  LY.styleOf = function (sh, r, c, cell) {
    const st = sh.wb.styles;
    if (cell && cell.s) return st.get(cell.s);
    if (cell && cell.s === 0) return st.get(0);
    const row = sh.rows[r];
    if (row && row.s) return st.get(row.s);
    const co = sh.cols[c] || (sh.colTail && c >= sh.colTail.from ? sh.colTail.o : null);
    if (co && co.s) return st.get(co.s);
    return st.get(0);
  };
  LY.styleIdOf = function (sh, r, c, cell) {
    if (cell && cell.s != null) return cell.s;
    const row = sh.rows[r];
    if (row && row.s) return row.s;
    const co = sh.cols[c] || (sh.colTail && c >= sh.colTail.from ? sh.colTail.o : null);
    if (co && co.s) return co.s;
    return 0;
  };
  /** theme-aware font name: scheme fonts follow the workbook theme */
  LY.fontName = (wb, f) => (f.scheme === 'minor' && wb.theme && wb.theme.minor ? wb.theme.minor : f.scheme === 'major' && wb.theme && wb.theme.major ? wb.theme.major : f.name || 'Arial');

  /* ------------------------------------------------------------ display text */
  /**
   * What a cell shows. Returns {text, segs, color, align ('l'|'r'|'c'), num (bool), hashes}
   * width: available pixels (for General numbers and ####), fontCss: the cell font
   */
  LY.display = function (sh, cell, st, width, fontCss, opts) {
    const v = cell.v;
    if (v == null || v === '') return null;
    const wb = sh.wb;
    if (sh.view.formulas && cell.f != null) return { text: '=' + L.formula.display(cell.f), segs: null, align: 'l', num: false };
    const nf = st.nf || 'General';
    const isNum = typeof v === 'number';
    if (isNum && /^General$/i.test(nf) && width != null) {
      /* General shows as many digits as the column width allows */
      const mdw = M.mdw(wb);
      let chars = Math.max(1, Math.min(11, Math.floor((width - 5) / mdw) + 1));
      let t = NF.general(v, chars);
      let tw = LY.hintedWidth(fontCss, t);
      while (tw > width && chars > 1) { chars--; t = NF.general(v, chars); if (/^#+$/.test(t)) break; tw = LY.hintedWidth(fontCss, t); }
      if (/^#+$/.test(t) || tw > width) return { text: hashes(width, fontCss), segs: null, align: 'r', num: true, hashes: true };
      return { text: t, segs: null, align: 'r', num: true, color: null };
    }
    const f = NF.format(nf, M.isErr(v) ? v : v, { date1904: wb.date1904 });
    const align = M.isErr(v) || typeof v === 'boolean' ? 'c' : isNum ? 'r' : 'l';
    if (f.hashes) return { text: hashes(width || 64, fontCss), segs: null, align: 'r', num: true, hashes: true, color: f.color };
    const out = { text: f.text, segs: f.segs && f.segs.some((x) => x.fill != null || x.pad != null) ? f.segs : null, align, num: isNum, color: f.color };
    if (isNum && width != null && !opts?.noHash) {
      const tw = out.segs ? segWidth(out.segs, fontCss) : LY.hintedWidth(fontCss, out.text);
      if (tw > width) return { text: hashes(width, fontCss), segs: null, align: 'r', num: true, hashes: true, color: f.color };
    }
    return out;
  };
  function hashes(width, font) {
    const w = LY.measure(font, '#') || 7;
    return '#'.repeat(Math.max(1, Math.floor((width - 3) / w)));
  }
  function segWidth(segs, font) {
    let w = 0;
    for (const s of segs) { if (s.s != null) w += LY.measure(font, s.s); else if (s.pad != null) w += LY.measure(font, s.pad); }
    return w;
  }
  LY.segWidth = segWidth;

  /* ------------------------------------------------------------ merges */
  /** quick merge lookup: index merges by top row */
  LY.mergeIndex = function (sh) {
    if (sh._mi && sh._mi.n === sh.merges.length && sh._mi.ver === sh._mergeVer) return sh._mi;
    const byCell = new Map();
    for (const m of sh.merges) {
      if ((m.r2 - m.r1 + 1) * (m.c2 - m.c1 + 1) > 200000) { byCell.set(M.key(m.r1, m.c1), m); continue; }
      for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) byCell.set(M.key(r, c), m);
    }
    sh._mi = { n: sh.merges.length, ver: sh._mergeVer, byCell };
    return sh._mi;
  };
  LY.mergeAt = (sh, r, c) => LY.mergeIndex(sh).byCell.get(M.key(r, c)) || null;
  LY.touchMerges = (sh) => { sh._mergeVer = (sh._mergeVer || 0) + 1; };
})(typeof window !== 'undefined' ? window : globalThis);
