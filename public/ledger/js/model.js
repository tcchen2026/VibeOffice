/* Ledger — workbook model.
 * Workbook → Sheets → sparse rows → sparse cells. Cells are small plain objects:
 *   {v: value, f?: formula text, s?: style index, rt?: rich-text runs, af?: array range (master), am?: array master cell}
 * Values: number | string | boolean | XErr | null. Styles are interned in a per-workbook table.
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const M = (L.model = {});

  /* ------------------------------------------------------------- errors */
  class XErr { constructor(e) { this.e = e; } toString() { return this.e; } }
  const ERR = {};
  for (const e of ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!', '#FIELD!', '#BLOCKED!', '#UNKNOWN!', '#CONNECT!', '#BUSY!', '#EXTERNAL!', '#PYTHON!']) ERR[e] = new XErr(e);
  M.XErr = XErr;
  M.ERR = ERR;
  M.err = (code) => ERR[code] || (ERR[code] = new XErr(code));
  M.isErr = (v) => v instanceof XErr;
  ERR.NULL = ERR['#NULL!']; ERR.DIV0 = ERR['#DIV/0!']; ERR.VALUE = ERR['#VALUE!']; ERR.REF = ERR['#REF!']; ERR.NAME = ERR['#NAME?']; ERR.NUM = ERR['#NUM!']; ERR.NA = ERR['#N/A']; ERR.CALC = ERR['#CALC!']; ERR.SPILL = ERR['#SPILL!'];

  M.MAXR = 1048576; M.MAXC = 16384;
  M.key = (r, c) => r * 16384 + c;

  /* ------------------------------------------------------------- colours */
  M.DEFAULT_THEME = {
    colors: ['#FFFFFF', '#000000', '#EEECE1', '#1F497D', '#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646', '#0000FF', '#800080'],
    major: 'Cambria', minor: 'Calibri', name: 'Office',
  };
  /* theme index order used by SpreadsheetML: lt1, dk1, lt2, dk2, accent1..6, hlink, folHlink */
  M.INDEXED = ['#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF', '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
    '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080', '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
    '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF', '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
    '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696', '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333'];
  const hex2rgb = (h) => { const n = parseInt(String(h).replace('#', '').slice(-6), 16) || 0; return [(n >> 16) & 255, (n >> 8) & 255, n & 255]; };
  const rgb2hex = (r, g, b) => '#' + [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase();
  function rgb2hsl(r, g, b) { r /= 255; g /= 255; b /= 255; const mx = Math.max(r, g, b), mn = Math.min(r, g, b); let h = 0, s = 0; const l = (mx + mn) / 2; if (mx !== mn) { const d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; } return [h, s, l]; }
  function hsl2rgb(h, s, l) { if (!s) return [l * 255, l * 255, l * 255]; const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; }; const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; return [f(p, q, h + 1 / 3) * 255, f(p, q, h) * 255, f(p, q, h - 1 / 3) * 255]; }
  /** SpreadsheetML tint: lighten (tint > 0) or darken (tint < 0) in HSL luminance */
  M.applyTint = function (hex, tint) {
    if (!tint) return hex;
    const [r, g, b] = hex2rgb(hex);
    const [h, s, l] = rgb2hsl(r, g, b);
    const l2 = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    const [r2, g2, b2] = hsl2rgb(h, s, Math.max(0, Math.min(1, l2)));
    return rgb2hex(r2, g2, b2);
  };
  /** colour object {rgb:'FFRRGGBB'|'#RRGGBB'} {theme:n, tint} {indexed:n} {auto:true} or '#hex' → '#RRGGBB' | null */
  M.colorHex = function (wb, c, auto) {
    if (c == null) return auto || null;
    if (typeof c === 'string') return c[0] === '#' ? c.toUpperCase() : '#' + c.slice(-6).toUpperCase();
    let hex = null;
    if (c.rgb) hex = '#' + String(c.rgb).replace('#', '').slice(-6).toUpperCase();
    else if (c.theme != null) {
      const th = (wb && wb.theme && wb.theme.colors) || M.DEFAULT_THEME.colors;
      /* SpreadsheetML swaps the first pairs: 0 lt1→dk1? Excel uses dk1/lt1 order as stored: 0=lt1(bg1),1=dk1(tx1),2=lt2,3=dk2 */
      hex = th[c.theme] || '#000000';
    } else if (c.indexed != null) {
      if (c.indexed === 64) return auto || '#000000'; /* system foreground */
      if (c.indexed === 65) return '#FFFFFF'; /* system background */
      const pal = (wb && wb.palette) || M.INDEXED;
      hex = pal[c.indexed] || '#000000';
    } else if (c.auto) return auto || '#000000';
    if (hex && c.tint) hex = M.applyTint(hex, c.tint);
    return hex;
  };
  M.rgb = (hex) => ({ rgb: 'FF' + String(hex).replace('#', '').toUpperCase() });

  /* --------------------------------------------------------------- styles */
  /**
   * A style: {font:{name,sz,b,i,u,strike,color,vert,family,scheme}, fill:{pattern,fg,bg}|{gradient}, border:{l,r,t,b,d,du,dd},
   *   align:{h,v,wrap,indent,rot,shrink,rtl}, nf: code, prot:{locked,hidden}, xs: named cell style}
   * Objects are interned: equal styles share one index.
   */
  class StyleTable {
    constructor(base) {
      this.list = [];
      this.map = new Map();
      this.add(base || M.defaultStyle());
    }
    add(st) {
      const k = JSON.stringify(st);
      let i = this.map.get(k);
      if (i != null) return i;
      i = this.list.length;
      this.list.push(Object.freeze(st));
      this.map.set(k, i);
      return i;
    }
    get(i) { return this.list[i || 0] || this.list[0]; }
    /** style i with fields replaced by delta (a partial style; nested objects merged one level) */
    derive(i, delta) {
      const base = this.get(i);
      const st = M.mergeStyle(base, delta);
      return this.add(st);
    }
  }
  M.StyleTable = StyleTable;
  M.defaultStyle = (font) => ({ font: Object.assign({ name: 'Arial', sz: 10, family: 2 }, font || {}), fill: null, border: null, align: null, nf: 'General', prot: null });
  const clean = (o) => { if (!o) return o; for (const k of Object.keys(o)) if (o[k] === undefined) delete o[k]; return o; };
  M.mergeStyle = function (base, delta) {
    const st = Object.assign({}, base);
    for (const k in delta) {
      const v = delta[k];
      if (v === undefined) continue;
      if (v === null) { st[k] = null; continue; }
      if (k === 'font' || k === 'align' || k === 'prot') st[k] = clean(Object.assign({}, base[k] || {}, v));
      else if (k === 'border') {
        const b = Object.assign({}, base.border || {});
        for (const side in v) { if (v[side] == null) delete b[side]; else b[side] = v[side]; }
        st.border = Object.keys(b).length ? b : null;
      } else st[k] = v;
    }
    if (st.align && !Object.keys(st.align).length) st.align = null;
    if (st.prot && !Object.keys(st.prot).length) st.prot = null;
    return st;
  };

  /* ------------------------------------------------------------- workbook */
  class Workbook {
    constructor(opts) {
      opts = opts || {};
      this.sheets = [];
      this.names = []; /* {name, ref, scope: sheetIndex|null, hidden, comment} */
      this.tables = [];
      this.theme = JSON.parse(JSON.stringify(M.DEFAULT_THEME));
      this.styles = new StyleTable(M.defaultStyle(opts.font));
      this.cellStyles = [{ name: 'Normal', builtinId: 0, style: 0 }];
      this.dxfs = []; /* differential formats for conditional formatting */
      this.props = { creator: opts.author || '', title: '', subject: '', keywords: '', description: '', lastModifiedBy: '', category: '', company: '', manager: '', created: null, modified: null };
      this.calcPr = { mode: 'auto', iterate: false, iterateCount: 100, iterateDelta: 0.001, fullPrecision: true };
      this.date1904 = false;
      this.active = 0;
      this.protection = null;
      this.palette = null;
      this.extra = {}; /* parts we keep but do not edit (vbaProject, customXml ...) */
      this.id = 'wb' + Math.random().toString(36).slice(2, 8);
      this.undo = new Undo(this);
      this.version = 0;
    }
    get defaultFont() { return this.styles.get(0).font; }
    sheet(i) { return this.sheets[i == null ? this.active : i]; }
    get activeSheet() { return this.sheets[this.active]; }
    sheetByName(name) {
      if (name == null) return null;
      const k = String(name).toLowerCase();
      return this.sheets.find((s) => s.name.toLowerCase() === k) || null;
    }
    sheetIndex(sh) { return this.sheets.indexOf(sh); }
    addSheet(name, at, kind) {
      const sh = new Sheet(this, name || this.nextSheetName(), kind);
      if (at == null || at >= this.sheets.length) this.sheets.push(sh); else this.sheets.splice(at, 0, sh);
      return sh;
    }
    nextSheetName(prefix) {
      prefix = prefix || 'Sheet';
      let i = 1;
      while (this.sheetByName(prefix + i)) i++;
      return prefix + i;
    }
    /** defined name lookup: sheet-scoped first */
    findName(name, sheetIdx) {
      const k = String(name).toLowerCase();
      let hit = null;
      for (const n of this.names) {
        if (n.name.toLowerCase() !== k) continue;
        if (n.scope != null && n.scope === sheetIdx) return n;
        if (n.scope == null) hit = hit || n;
      }
      return hit;
    }
    findTable(name) { const k = String(name).toLowerCase(); return this.tables.find((t) => t.name.toLowerCase() === k) || null; }
    tableAt(sheet, r, c) { return this.tables.find((t) => t.sheet === sheet && r >= t.ref.r1 && r <= t.ref.r2 && c >= t.ref.c1 && c <= t.ref.c2) || null; }
    touch() { this.version++; }
  }
  M.Workbook = Workbook;

  /* ---------------------------------------------------------------- sheet */
  class Sheet {
    constructor(wb, name, kind) {
      this.wb = wb;
      this.name = name;
      this.kind = kind || 'worksheet';
      this.rows = [];
      this.cols = [];
      this.merges = [];
      this.dv = [];
      this.cf = [];
      this.links = [];
      this.comments = new Map();
      this.drawings = [];
      this.autoFilter = null;
      this.state = 'visible';
      this.tabColor = null;
      this.defColW = null; /* in file units; null → computed from baseColW */
      this.baseColW = 8;
      this.defRowH = null; /* points; null → from default font */
      this.view = { freeze: null, zoom: 100, grid: true, headings: true, zeros: true, rtl: false, top: { r: 0, c: 0 }, sel: null, formulas: false, gridColor: null, pageBreakPreview: false };
      this.print = { margins: { l: 0.75, r: 0.75, t: 1, b: 1, header: 0.5, footer: 0.5 }, orientation: 'portrait', paper: 1, scale: 100, fitW: null, fitH: null, fit: false, hCenter: false, vCenter: false, gridLines: false, headings: false, header: '', footer: '', firstPage: null, pageOrder: 'downThenOver', rowBreaks: [], colBreaks: [], area: null, titleRows: null, titleCols: null, bw: false, draft: false, comments: 'none', errors: 'displayed' };
      this.protection = null;
      this.outline = { below: true, right: true };
      this.maxR = -1; this.maxC = -1; /* used range bounds (cells with content or style) */
      this.id = 's' + Math.random().toString(36).slice(2, 9);
      this.extra = {};
    }
    get index() { return this.wb.sheets.indexOf(this); }
    row(r) { return this.rows[r]; }
    rowObj(r) { let row = this.rows[r]; if (!row) { row = { cells: [] }; this.rows[r] = row; } return row; }
    get(r, c) { const row = this.rows[r]; return row ? row.cells[c] : undefined; }
    /** value of a cell (null when blank) */
    val(r, c) { const row = this.rows[r]; const cell = row && row.cells[c]; return cell ? cell.v : null; }
    /** create the cell object if needed (used by loaders; editing goes through Undo) */
    cell(r, c) {
      const row = this.rowObj(r);
      let cell = row.cells[c];
      if (!cell) { cell = { v: null }; row.cells[c] = cell; if (r > this.maxR) this.maxR = r; if (c > this.maxC) this.maxC = c; }
      return cell;
    }
    put(r, c, cell) {
      const row = this.rowObj(r);
      if (cell == null) { delete row.cells[c]; return; }
      row.cells[c] = cell;
      if (r > this.maxR) this.maxR = r;
      if (c > this.maxC) this.maxC = c;
    }
    /** iterate existing cells in a rectangle: fn(cell, r, c) ; return false to stop */
    each(r1, c1, r2, c2, fn) {
      r2 = Math.min(r2, this.maxR); c2 = Math.min(c2, this.maxC);
      const wide = c2 - c1 > 48;
      for (let r = r1; r <= r2; r++) {
        const row = this.rows[r];
        if (!row) continue;
        const cells = row.cells;
        if (wide) {
          for (const k in cells) { const c = +k; if (c >= c1 && c <= c2) { const cell = cells[c]; if (cell && fn(cell, r, c) === false) return; } }
        } else {
          for (let c = c1; c <= c2; c++) { const cell = cells[c]; if (cell && fn(cell, r, c) === false) return; }
        }
      }
    }
    /** recompute the used range, including retained source styles on blank cells */
    recalcBounds() {
      let mr = -1, mc = -1;
      this.rows.forEach((row, r) => {
        if (!row) return;
        let any = false;
        row.cells.forEach((cell, c) => { if (cell && (cell.v != null || cell.f != null || cell.s || cell.keep?.style != null)) { any = true; if (c > mc) mc = c; } });
        if (any && r > mr) mr = r;
      });
      this.maxR = mr; this.maxC = mc;
      return { r: mr, c: mc };
    }
    /** bottom-right of content (values / formulas only), like Ctrl+End */
    contentEnd() {
      let mr = 0, mc = 0;
      this.rows.forEach((row, r) => { if (row) row.cells.forEach((cell, c) => { if (cell && (cell.v != null || cell.f != null || cell.s)) { if (r > mr) mr = r; if (c > mc) mc = c; } }); });
      return { r: mr, c: mc };
    }
    colInfo(c) { return this.cols[c]; }
    colObj(c) { let o = this.cols[c]; if (!o) { o = {}; this.cols[c] = o; } return o; }
    mergeAt(r, c) { for (const m of this.merges) if (r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) return m; return null; }
    styleAt(r, c) {
      const cell = this.get(r, c);
      if (cell && cell.s != null) return cell.s;
      const row = this.rows[r];
      if (row && row.s != null) return row.s;
      const col = this.cols[c];
      if (col && col.s != null) return col.s;
      return 0;
    }
  }
  M.Sheet = Sheet;

  /* ---------------------------------------------------------- sizing */
  /* Column widths are kept in SpreadsheetML units (characters of the default font's widest digit, with padding). */
  M.mdw = function (wb) {
    if (wb && wb._mdw) return wb._mdw;
    const f = wb ? wb.defaultFont : { name: 'Arial', sz: 10 };
    let w = 7;
    const name = String(f.name || '').toLowerCase();
    const sz = f.sz || 11;
    if (L.measureDigit) w = L.measureDigit(f);
    else {
      const k = { calibri: 7 / 11, arial: 7 / 10, 'aptos narrow': 7 / 11, aptos: 7.5 / 11, cambria: 7 / 11, 'times new roman': 7 / 12, verdana: 8 / 10, tahoma: 7 / 10, 'courier new': 8 / 10, 'segoe ui': 7 / 10 }[name] || 7 / 11;
      w = Math.max(5, Math.round(k * sz));
    }
    if (wb) wb._mdw = w;
    return w;
  };
  M.widthToPx = (w, mdw) => Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw);
  M.pxToWidth = (px, mdw) => Math.trunc((px / mdw) * 256) / 256 + (px > 0 ? 0 : 0);
  /** characters shown in the Column Width box for a pixel width */
  M.pxToChars = (px, mdw) => Math.max(0, Math.trunc(((px - 5) / mdw) * 100 + 0.5) / 100);
  M.charsToPx = (ch, mdw) => (ch <= 0 ? 0 : Math.trunc(ch * mdw + 5 + 0.5));
  M.defaultColPx = function (sh) {
    const mdw = M.mdw(sh.wb);
    if (sh.defColW != null) return M.widthToPx(sh.defColW, mdw);
    return Math.ceil(((sh.baseColW == null ? 8 : sh.baseColW) * mdw + 5) / 8) * 8;
  };
  M.colPx = function (sh, c) {
    let o = sh.cols[c];
    if (!o && sh.colTail && c >= sh.colTail.from) o = sh.colTail.o;
    if (o) { if (o.hidden) return 0; if (o.w != null) return M.widthToPx(o.w, M.mdw(sh.wb)); }
    return M.defaultColPx(sh);
  };
  /** default row height in points for the default font */
  M.fontRowPt = function (font) {
    const name = String((font && font.name) || 'Arial').toLowerCase();
    const sz = (font && font.sz) || 11;
    const table = { arial: { 8: 11.25, 9: 12, 10: 12.75, 11: 14.25, 12: 15, 14: 18, 16: 20.25, 18: 23.25, 20: 25.5, 24: 30 }, calibri: { 8: 11.25, 9: 12, 10: 12.75, 11: 15, 12: 15.75, 14: 18.75, 16: 21, 18: 23.25, 20: 26.25, 24: 31.5 } };
    const t = table[name] || (/calibri|aptos|segoe/.test(name) ? table.calibri : null);
    if (t && t[sz]) return t[sz];
    return Math.round(sz * 1.275 * 4) / 4;
  };
  M.defaultRowPt = (sh) => sh.defRowH ?? (sh.extra.formatKeep?.legacyNamespace ? sh.fileDefRowH : null) ?? M.fontRowPt(sh.wb.defaultFont);
  M.rowPt = function (sh, r) {
    const row = sh.rows[r];
    if (row) { if (row.hidden) return 0; if (row.ht != null && (row.customHeight || row.auto == null)) return row.ht; if (row.auto != null) return row.auto; }
    return M.defaultRowPt(sh);
  };
  M.ptToPx = (pt) => Math.round((pt * 96) / 72);

  /* ---------------------------------------------------------- undo */
  /**
   * Transactions record cell patches (old/new cell objects) and custom undo/redo closures.
   *   wb.undo.begin('Typing'); ...set cells through U.setCell...; wb.undo.end();
   */
  class Undo {
    constructor(wb) { this.wb = wb; this.stack = []; this.redoStack = []; this.cur = null; this.depth = 0; this.limit = 100; this.onChange = null; }
    begin(label) {
      if (this.depth++ === 0) this.cur = { label: label || '', cells: new Map(), ops: [], sheet: this.wb.active };
      return this.cur;
    }
    end() {
      if (this.depth === 0) return;
      if (--this.depth > 0) return;
      const t = this.cur;
      this.cur = null;
      if (!t || (!t.cells.size && !t.ops.length)) return;
      this.stack.push(t);
      if (this.stack.length > this.limit) this.stack.shift();
      this.redoStack = [];
      this.wb.touch();
      if (this.onChange) this.onChange(t);
    }
    get active() { return !!this.cur; }
    /** remember a cell before it changes (first touch only) */
    note(sh, r, c) {
      if (!this.cur) return;
      const k = sh.id + ':' + r + ':' + c;
      if (this.cur.cells.has(k)) return;
      const old = sh.get(r, c);
      this.cur.cells.set(k, { sh, r, c, old: old ? Object.assign({}, old) : null });
    }
    /** custom step: undo() and redo() closures */
    op(undo, redo) { if (this.cur) this.cur.ops.push({ undo, redo }); }
    canUndo() { return this.stack.length > 0; }
    canRedo() { return this.redoStack.length > 0; }
    get undoLabel() { const t = this.stack[this.stack.length - 1]; return t ? t.label : ''; }
    get redoLabel() { const t = this.redoStack[this.redoStack.length - 1]; return t ? t.label : ''; }
    apply(t, dir) {
      /* capture the current state for the opposite direction */
      const touched = [];
      for (const e of t.cells.values()) {
        const now = e.sh.get(e.r, e.c);
        touched.push({ sh: e.sh, r: e.r, c: e.c, now: now ? Object.assign({}, now) : null });
      }
      if (dir < 0) for (let i = t.ops.length - 1; i >= 0; i--) t.ops[i].undo();
      for (const e of t.cells.values()) e.sh.put(e.r, e.c, e.old ? Object.assign({}, e.old) : null);
      if (dir > 0) for (const o of t.ops) o.redo();
      /* swap recorded states so the transaction can be replayed the other way */
      const next = { label: t.label, cells: new Map(), ops: t.ops, sheet: t.sheet };
      for (const x of touched) next.cells.set(x.sh.id + ':' + x.r + ':' + x.c, { sh: x.sh, r: x.r, c: x.c, old: x.now });
      return next;
    }
    undo() {
      const t = this.stack.pop();
      if (!t) return null;
      const back = this.apply(t, -1);
      this.redoStack.push(back);
      this.wb.touch();
      if (this.onChange) this.onChange(t, 'undo');
      return t;
    }
    redo() {
      const t = this.redoStack.pop();
      if (!t) return null;
      const fwd = this.apply(t, 1);
      this.stack.push(fwd);
      this.wb.touch();
      if (this.onChange) this.onChange(t, 'redo');
      return t;
    }
    clear() { this.stack = []; this.redoStack = []; }
  }
  M.Undo = Undo;

  /* ---------------------------------------------------- range helpers */
  M.rangeContains = (a, r, c) => r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2;
  M.rangesOverlap = (a, b) => a.r1 <= b.r2 && b.r1 <= a.r2 && a.c1 <= b.c2 && b.c1 <= a.c2;
  M.rangeUnion = (a, b) => ({ r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) });
  M.rangeSize = (a) => (a.r2 - a.r1 + 1) * (a.c2 - a.c1 + 1);
  M.normRange = (a) => ({ r1: Math.min(a.r1, a.r2), c1: Math.min(a.c1, a.c2), r2: Math.max(a.r1, a.r2), c2: Math.max(a.c1, a.c2) });
  /** "A1:B2 D4" (space separated, SpreadsheetML sqref) → ranges */
  M.parseSqref = (s) => String(s || '').trim().split(/\s+/).filter(Boolean).map((x) => L.formula.parseRange(x)).filter(Boolean);
  M.sqref = (list) => list.map((r) => L.formula.rangeName(r)).join(' ');
})(typeof window !== 'undefined' ? window : globalThis);
