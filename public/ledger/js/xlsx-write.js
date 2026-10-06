/* Ledger — .xlsx / .xlsm writer (SpreadsheetML transitional).
 * L.xlsxWrite.write(wb, {type:'xlsx'|'xlsm'|'xltx', password}) → Promise<Blob>
 * L.xlsxWrite.bytes(wb, opts) → Promise<Uint8Array>   (used by tests)
 */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, F = L.formula, NF = L.numfmt, XML = L.xml;
  const W = (L.xlsxWrite = {});
  const HDR = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';
  const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
  const RT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const CT = {
    wb: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
    wbm: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
    wbt: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
    wbtm: 'application/vnd.ms-excel.template.macroEnabled.main+xml',
    sheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml',
    chartsheet: 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml',
    styles: 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml',
    sst: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml',
    theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
    comments: 'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml',
    table: 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml',
    drawing: 'application/vnd.openxmlformats-officedocument.drawing+xml',
    chart: 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml',
    chartStyle: 'application/vnd.ms-office.chartstyle+xml',
    chartColors: 'application/vnd.ms-office.chartcolorstyle+xml',
    core: 'application/vnd.openxmlformats-package.core-properties+xml',
    app: 'application/vnd.openxmlformats-officedocument.extended-properties+xml',
    custom: 'application/vnd.openxmlformats-officedocument.custom-properties+xml',
    ext: 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml',
    meta: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml',
    vml: 'application/vnd.openxmlformats-officedocument.vmlDrawing',
    vba: 'application/vnd.ms-office.vbaProject',
  };
  const IMG_CT = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/x-emf', wmf: 'image/x-wmf', svg: 'image/svg+xml', webp: 'image/webp' };
  const EMU = 12700;

  const esc = XML.esc;
  /* lone surrogates cannot appear in XML: they are written as _xHHHH_ escapes */
  const escT = (s) => XML.escText(String(s)).replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g, (m) => '_x' + m.charCodeAt(0).toString(16).toUpperCase() + '_');
  /** a text node that keeps its leading / trailing spaces */
  const tEl = (s, tag) => { s = String(s); tag = tag || 't'; return /^\s|\s$|\n/.test(s) ? `<${tag} xml:space="preserve">${escT(s)}</${tag}>` : `<${tag}>${escT(s)}</${tag}>`; };
  const attrs = (o) => { let s = ''; for (const k in o) { const v = o[k]; if (v === undefined || v === null || v === false) continue; s += ' ' + k + '="' + esc(v === true ? '1' : v) + '"'; } return s; };
  const numStr = (v) => { if (!isFinite(v)) return null; if (Object.is(v, -0)) return '0'; return String(v); };

  /* ------------------------------------------------------------ package plumbing */
  class Rels {
    constructor() { this.list = []; }
    add(type, target, external) {
      const id = 'rId' + (this.list.length + 1);
      this.list.push({ id, type: /^https?:/.test(type) ? type : RT + type, target, external });
      return id;
    }
    xml() {
      return HDR + `<Relationships xmlns="${NS_PKG}">` + this.list.map((r) => `<Relationship Id="${r.id}" Type="${esc(r.type)}" Target="${esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('') + '</Relationships>';
    }
  }
  const relsPath = (p) => { const i = p.lastIndexOf('/'); return (i < 0 ? '' : p.slice(0, i + 1)) + '_rels/' + p.slice(i + 1) + '.rels'; };

  /* ------------------------------------------------------------ colours */
  function colorAttrs(c) {
    if (!c) return '';
    if (typeof c === 'string') return ` rgb="FF${c.replace('#', '').slice(-6).toUpperCase()}"`;
    if (c.auto) return ' auto="1"';
    let s = '';
    if (c.rgb) { let v = String(c.rgb).replace('#', '').toUpperCase(); if (v.length === 6) v = 'FF' + v; s = ` rgb="${v}"`; }
    else if (c.theme != null) s = ` theme="${c.theme}"`;
    else if (c.indexed != null) s = ` indexed="${c.indexed}"`;
    else return '';
    if (c.tint) s += ` tint="${c.tint}"`;
    return s;
  }
  const colorEl = (tag, c) => { const a = colorAttrs(c); return a ? `<${tag}${a}/>` : ''; };

  /* ------------------------------------------------------------ styles */
  function fontXml(f, rich) {
    if (!f) f = {};
    let s = '';
    if (rich && f.name) s += `<rFont val="${esc(f.name)}"/>`;
    if (f.charset != null) s += `<charset val="${f.charset}"/>`;
    if (rich && f.family != null) s += `<family val="${f.family}"/>`;
    if (f.b) s += '<b/>';
    if (f.i) s += '<i/>';
    if (f.strike) s += '<strike/>';
    if (f.outline) s += '<outline/>';
    if (f.shadow) s += '<shadow/>';
    if (rich) { if (f.color) s += colorEl('color', f.color); if (f.sz) s += `<sz val="${f.sz}"/>`; }
    if (f.u) s += f.u === 'single' || f.u === true ? '<u/>' : `<u val="${esc(f.u)}"/>`;
    if (f.vert) s += `<vertAlign val="${esc(f.vert)}"/>`;
    if (!rich) {
      if (f.sz) s += `<sz val="${f.sz}"/>`;
      if (f.color) s += colorEl('color', f.color);
      if (f.name) s += `<name val="${esc(f.name)}"/>`;
      if (f.family != null) s += `<family val="${f.family}"/>`;
    }
    if (f.scheme) s += `<scheme val="${esc(f.scheme)}"/>`;
    return s;
  }
  /* charset must precede the others in CT_Font order; rebuild in schema order for cell fonts */
  /** a cell font; name and size missing from it are those of the Normal style (what Ledger draws) */
  function cellFontXml(f, dflt) {
    f = f || {};
    dflt = dflt || {};
    let s = '';
    if (f.b) s += '<b/>';
    if (f.i) s += '<i/>';
    if (f.strike) s += '<strike/>';
    if (f.outline) s += '<outline/>';
    if (f.shadow) s += '<shadow/>';
    if (f.u) s += f.u === 'single' || f.u === true ? '<u/>' : `<u val="${esc(f.u)}"/>`;
    if (f.vert) s += `<vertAlign val="${esc(f.vert)}"/>`;
    s += `<sz val="${f.sz || dflt.sz || 11}"/>`;
    if (f.color) s += colorEl('color', f.color);
    s += `<name val="${esc(f.name || dflt.name || 'Calibri')}"/>`;
    if (f.family != null) s += `<family val="${f.family}"/>`;
    if (f.charset != null) s += `<charset val="${f.charset}"/>`;
    if (f.scheme) s += `<scheme val="${esc(f.scheme)}"/>`;
    return '<font>' + s + '</font>';
  }
  function fillXml(fl, dxf) {
    if (!fl) return dxf ? '' : '<fill><patternFill patternType="none"/></fill>';
    if (fl.gradient) {
      const g = fl.gradient;
      return `<fill><gradientFill${attrs({ type: g.type === 'path' ? 'path' : undefined, degree: g.type !== 'path' && g.degree ? g.degree : undefined, left: g.left, right: g.right, top: g.top, bottom: g.bottom })}>` +
        (g.stops || []).map((s) => `<stop position="${s.pos}">${colorEl('color', s.color) || '<color rgb="FF000000"/>'}</stop>`).join('') + '</gradientFill></fill>';
    }
    if (dxf) {
      if (!fl.pattern || fl.pattern === 'solid') return `<fill><patternFill>${colorEl('bgColor', fl.fg || fl.bg)}</patternFill></fill>`;
      return `<fill><patternFill patternType="${esc(fl.pattern)}">${colorEl('fgColor', fl.fg)}${colorEl('bgColor', fl.bg)}</patternFill></fill>`;
    }
    const pt = fl.pattern || 'solid';
    if (pt === 'none') return '<fill><patternFill patternType="none"/></fill>';
    return `<fill><patternFill patternType="${esc(pt)}">${colorEl('fgColor', fl.fg)}${colorEl('bgColor', fl.bg)}</patternFill></fill>`;
  }
  function borderXml(b) {
    b = b || {};
    const side = (tag, x) => (x && x.style ? (x.color ? `<${tag} style="${esc(x.style)}">${colorEl('color', x.color)}</${tag}>` : `<${tag} style="${esc(x.style)}"/>`) : `<${tag}/>`);
    return `<border${attrs({ diagonalUp: b.du ? '1' : undefined, diagonalDown: b.dd ? '1' : undefined })}>` + side('left', b.l) + side('right', b.r) + side('top', b.t) + side('bottom', b.b) + side('diagonal', b.d) + '</border>';
  }
  function alignXml(a) {
    if (!a) return '';
    const o = { horizontal: a.h, vertical: a.v, textRotation: a.rot || undefined, wrapText: a.wrap ? '1' : undefined, indent: a.indent || undefined, justifyLastLine: a.justLast ? '1' : undefined, shrinkToFit: a.shrink ? '1' : undefined, readingOrder: a.rtl ? '2' : a.ltr ? '1' : undefined };
    const s = attrs(o);
    return s ? `<alignment${s}/>` : '';
  }
  function protXml(p) {
    if (!p) return '';
    const s = attrs({ locked: p.locked === false ? '0' : undefined, hidden: p.hidden ? '1' : undefined });
    return s ? `<protection${s}/>` : '';
  }
  function buildStyles(wb) {
    const fonts = [], fontIdx = new Map();
    const fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], fillIdx = new Map([[fills[0], 0], [fills[1], 1]]);
    const borders = [], borderIdx = new Map();
    const numFmts = [], nfIdx = new Map();
    let nextNf = 164;
    const intern = (list, map, x) => { let i = map.get(x); if (i == null) { i = list.length; list.push(x); map.set(x, i); } return i; };
    const nfId = (code) => {
      code = code == null ? 'General' : String(code);
      if (code === 'General') return 0;
      const b = NF.builtinId(code);
      if (b >= 0) return b;
      let id = nfIdx.get(code);
      if (id == null) { id = nextNf++; nfIdx.set(code, id); numFmts.push(`<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`); }
      return id;
    };
    intern(borders, borderIdx, borderXml(null));
    const xfOf = (st, xfId) => {
      const fontId = intern(fonts, fontIdx, cellFontXml(st.font, wb.styles.get(0).font));
      const fx = fillXml(st.fill);
      const fillId = intern(fills, fillIdx, fx);
      const borderId = intern(borders, borderIdx, borderXml(st.border));
      const numFmtId = nfId(st.nf);
      const al = alignXml(st.align), pr = protXml(st.prot);
      const a = { numFmtId, fontId, fillId, borderId };
      if (xfId != null) { a.xfId = xfId; if (numFmtId) a.applyNumberFormat = '1'; if (fontId) a.applyFont = '1'; if (fillId) a.applyFill = '1'; if (borderId) a.applyBorder = '1'; if (al) a.applyAlignment = '1'; if (pr) a.applyProtection = '1'; }
      return `<xf${attrs(a)}` + (al || pr ? '>' + al + pr + '</xf>' : '/>');
    };
    /* named styles: Normal first */
    let named = (wb.cellStyles || []).filter((c) => c && c.name);
    if (!named.some((c) => c.builtinId === 0)) named.unshift({ name: 'Normal', builtinId: 0, style: 0 });
    named = named.slice().sort((a, b) => (a.builtinId === 0 ? -1 : b.builtinId === 0 ? 1 : 0));
    const seenNames = new Set();
    named = named.filter((c) => { const k = c.name.toLowerCase(); if (seenNames.has(k)) return false; seenNames.add(k); return true; });
    const csXfs = named.map((c, i) => xfOf(i === 0 ? wb.styles.get(0) : wb.styles.get(c.style || 0), null));
    const nameIdx = new Map(named.map((c, i) => [c.name, i]));
    /* cell formats: one per style-table entry, so a cell's style index is its xf index */
    const cellXfs = wb.styles.list.map((st) => xfOf(st, (st.xs && nameIdx.get(st.xs)) || 0));
    /* differential formats */
    const dxfs = (wb.dxfs || []).map((d) => {
      d = d || {};
      let s = '<dxf>';
      if (d.font) s += '<font>' + fontXml(d.font, true).replace(/<rFont [^>]*>/, '') + '</font>';
      if (d.nf) s += `<numFmt numFmtId="${nfId(d.nf)}" formatCode="${esc(d.nf)}"/>`;
      if (d.fill) s += fillXml(d.fill, true);
      if (d.align) s += alignXml(d.align);
      if (d.border) s += borderXml(d.border);
      return s + '</dxf>';
    });
    let x = HDR + `<styleSheet xmlns="${NS_MAIN}">`;
    if (numFmts.length) x += `<numFmts count="${numFmts.length}">${numFmts.join('')}</numFmts>`;
    x += `<fonts count="${fonts.length}">${fonts.join('')}</fonts>`;
    x += `<fills count="${fills.length}">${fills.join('')}</fills>`;
    x += `<borders count="${borders.length}">${borders.join('')}</borders>`;
    x += `<cellStyleXfs count="${csXfs.length}">${csXfs.join('')}</cellStyleXfs>`;
    x += `<cellXfs count="${cellXfs.length}">${cellXfs.join('')}</cellXfs>`;
    x += `<cellStyles count="${named.length}">` + named.map((c, i) => `<cellStyle${attrs({ name: c.name, xfId: i, builtinId: c.builtinId != null ? c.builtinId : undefined, hidden: c.hidden ? '1' : undefined, customBuiltin: c.customBuiltin ? '1' : undefined })}/>`).join('') + '</cellStyles>';
    x += `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`;
    const ts = wb.tableStyles;
    if (ts && ts.custom && ts.custom.length) {
      x += `<tableStyles count="${ts.custom.length}"${attrs({ defaultTableStyle: ts.def || 'TableStyleMedium2', defaultPivotStyle: ts.defPivot || 'PivotStyleLight16' })}>` +
        ts.custom.map((t) => `<tableStyle${attrs({ name: t.name, pivot: t.pivot ? undefined : '0', table: t.table ? undefined : '0', count: t.elements.length })}>` + t.elements.map((e) => `<tableStyleElement${attrs({ type: e.type, size: e.size > 1 ? e.size : undefined, dxfId: e.dxf >= 0 ? e.dxf : undefined })}/>`).join('') + '</tableStyle>').join('') + '</tableStyles>';
    } else x += `<tableStyles count="0" defaultTableStyle="${esc((ts && ts.def) || 'TableStyleMedium2')}" defaultPivotStyle="${esc((ts && ts.defPivot) || 'PivotStyleLight16')}"/>`;
    if (wb.palette || (wb.mruColors && wb.mruColors.length)) {
      x += '<colors>';
      if (wb.palette) x += '<indexedColors>' + wb.palette.slice(0, 64).map((c) => `<rgbColor rgb="FF${String(c).replace('#', '').toUpperCase()}"/>`).join('') + '</indexedColors>';
      if (wb.mruColors && wb.mruColors.length) x += '<mruColors>' + wb.mruColors.map((c) => colorEl('color', c)).join('') + '</mruColors>';
      x += '</colors>';
    }
    x += '</styleSheet>';
    return x;
  }

  /* ------------------------------------------------------------ shared strings */
  class SST {
    constructor() { this.list = []; this.map = new Map(); this.count = 0; }
    add(v, runs) {
      this.count++;
      let key = runs && runs.length ? '\u0001' + JSON.stringify(runs) : v;
      let i = this.map.get(key);
      if (i == null) { i = this.list.length; this.list.push(runs && runs.length ? { runs } : v); this.map.set(key, i); }
      return i;
    }
    xml() {
      const out = [HDR, `<sst xmlns="${NS_MAIN}" count="${this.count}" uniqueCount="${this.list.length}">`];
      for (const s of this.list) {
        if (typeof s === 'string') out.push('<si>' + tEl(s.length > 32767 ? s.slice(0, 32767) : s) + '</si>');
        else out.push('<si>' + s.runs.map((r) => '<r>' + (r.font ? '<rPr>' + fontXml(r.font, true) + '</rPr>' : '') + tEl(r.t || '') + '</r>').join('') + '</si>');
      }
      out.push('</sst>');
      return out.join('');
    }
  }

  /* ------------------------------------------------------------ worksheet */
  const cellName = (r, c) => F.cellName(r, c);
  const rangeName = (rg) => (rg.r1 === rg.r2 && rg.c1 === rg.c2 ? cellName(rg.r1, rg.c1) : cellName(rg.r1, rg.c1) + ':' + cellName(rg.r2, rg.c2));
  const sqref = (list) => list.map(rangeName).join(' ');

  function sheetViewXml(sh, wb, idx) {
    const v = sh.view;
    const vs = sh._vs;
    const a = { tabSelected: idx === wb.active ? '1' : undefined, showGridLines: v.grid === false ? '0' : undefined, showRowColHeaders: v.headings === false ? '0' : undefined, showZeros: v.zeros === false ? '0' : undefined, rightToLeft: v.rtl ? '1' : undefined, showFormulas: v.formulas ? '1' : undefined, showOutlineSymbols: v.outlineSymbols === false ? '0' : undefined };
    if (v.gridColor && v.gridColor.indexed != null) { a.defaultGridColor = '0'; a.colorId = v.gridColor.indexed; }
    if (v.pageBreakPreview) a.view = 'pageBreakPreview';
    else if (v.pageLayout) a.view = 'pageLayout';
    const top = v.freeze ? (v.top || { r: 0, c: 0 }) : vs ? { r: vs.scrollR, c: vs.scrollC } : v.top || { r: 0, c: 0 };
    if (top.r || top.c) a.topLeftCell = cellName(top.r, top.c);
    if ((v.zoom || 100) !== 100) { a.zoomScale = Math.round(v.zoom); if (!v.pageBreakPreview) a.zoomScaleNormal = Math.round(v.zoom); }
    if (v.pageBreakPreview) a.zoomScaleSheetLayoutView = Math.round(v.zoom || 100);
    a.workbookViewId = '0';
    let s = `<sheetView${attrs(a)}>`;
    const sel = vs ? vs.sel : v.sel;
    let pane = null;
    if (v.freeze && (v.freeze.r || v.freeze.c)) {
      const fr = v.freeze;
      const tl = vs ? { r: Math.max(vs.scrollR, top.r + fr.r), c: Math.max(vs.scrollC, top.c + fr.c) } : fr.top || { r: top.r + fr.r, c: top.c + fr.c };
      pane = fr.r && fr.c ? 'bottomRight' : fr.r ? 'bottomLeft' : 'topRight';
      s += `<pane${attrs({ xSplit: fr.c || undefined, ySplit: fr.r || undefined, topLeftCell: cellName(tl.r, tl.c), activePane: pane, state: 'frozen' })}/>`;
      if (fr.r && fr.c) s += '<selection pane="topRight"/><selection pane="bottomLeft"/>';
    } else if (v.split && (v.split.x || v.split.y)) {
      const sp = v.split;
      pane = sp.x && sp.y ? 'bottomRight' : sp.y ? 'bottomLeft' : 'topRight';
      s += `<pane${attrs({ xSplit: sp.x || undefined, ySplit: sp.y || undefined, topLeftCell: sp.top ? cellName(sp.top.r, sp.top.c) : undefined, activePane: pane })}/>`;
    }
    if (sel && sel.ranges && sel.ranges.length) {
      const rs = sel.ranges.map((r) => ({ r1: Math.min(r.r1, M.MAXR - 1), c1: Math.min(r.c1, M.MAXC - 1), r2: Math.min(r.r2, M.MAXR - 1), c2: Math.min(r.c2, M.MAXC - 1) }));
      s += `<selection${attrs({ pane: pane || undefined, activeCell: cellName(sel.r, sel.c), sqref: sqref(rs) })}/>`;
    }
    return '<sheetViews>' + s + '</sheetView></sheetViews>';
  }

  function colsXml(sh) {
    const out = [];
    const sig = (o) => (o ? JSON.stringify([o.w, o.custom, o.hidden, o.bestFit, o.level, o.collapsed, o.s]) : null);
    const n = Math.max(sh.cols.length, 0);
    let c = 0;
    const emit = (min, max, o) => {
      if (!o) return;
      const a = { min: min + 1, max: max + 1, width: o.w != null ? o.w : undefined, style: o.s ? o.s : undefined, hidden: o.hidden ? '1' : undefined, bestFit: o.bestFit ? '1' : undefined, customWidth: o.custom || (o.w != null && !o.bestFit) ? '1' : undefined, outlineLevel: o.level || undefined, collapsed: o.collapsed ? '1' : undefined };
      if (a.width == null) { if (!a.style && !a.hidden && !a.outlineLevel && !a.collapsed) return; a.width = sh.defColW != null ? sh.defColW : M.pxToWidth(M.defaultColPx(sh), M.mdw(sh.wb)); a.customWidth = undefined; }
      out.push(`<col${attrs(a)}/>`);
    };
    while (c < n) {
      const o = sh.cols[c];
      if (!o || !Object.keys(o).length) { c++; continue; }
      const k = sig(o);
      let e = c;
      while (e + 1 < n && sh.cols[e + 1] && sig(sh.cols[e + 1]) === k) e++;
      if (sh.colTail && e + 1 >= sh.colTail.from && sig(sh.colTail.o) === k && e + 1 >= n) { emit(c, M.MAXC - 1, o); return '<cols>' + out.join('') + '</cols>'; }
      emit(c, e, o);
      c = e + 1;
    }
    if (sh.colTail) emit(Math.max(n, sh.colTail.from), M.MAXC - 1, sh.colTail.o);
    return out.length ? '<cols>' + out.join('') + '</cols>' : '';
  }

  function cellXml(cell, r, c, ctx, headerName) {
    let s = '<c r="' + cellName(r, c) + '"';
    if (cell.s) s += ' s="' + cell.s + '"';
    let v = cell.v;
    const hasF = cell.f != null && cell.f !== '' && !cell.am;
    if (headerName != null && !hasF) v = headerName;
    let t = null, body = '';
    if (hasF || cell.dt) {
      if (cell.dt) {
        const d = cell.dt;
        body += `<f${attrs({ t: 'dataTable', ref: d.ref, dt2D: d.dt2D ? '1' : undefined, dtr: d.dtr ? '1' : undefined, r1: d.r1, r2: d.r2, del1: d.del1 ? '1' : undefined, del2: d.del2 ? '1' : undefined, ca: d.ca ? '1' : undefined })}/>`;
      } else if (cell.af) {
        body += `<f t="array" ref="${rangeName(cell.af)}"${cell.ca ? ' ca="1"' : ''}>${esc(cell.f)}</f>`;
        if (cell.dyn) { s += ' cm="1"'; ctx.dynamic = true; }
      } else body += `<f${cell.ca ? ' ca="1"' : ''}>${esc(cell.f)}</f>`;
      if (v == null) { /* never calculated: Excel computes it on load */ }
      else if (typeof v === 'number') { const n = numStr(v); if (n == null) { t = 'e'; body += '<v>#NUM!</v>'; } else body += '<v>' + n + '</v>'; }
      else if (typeof v === 'string') { t = 'str'; body += '<v>' + escT(v) + '</v>'; }
      else if (typeof v === 'boolean') { t = 'b'; body += '<v>' + (v ? 1 : 0) + '</v>'; }
      else if (M.isErr(v)) { t = 'e'; body += '<v>' + esc(v.e) + '</v>'; }
    } else {
      if (v == null) { /* styled blank */ }
      else if (typeof v === 'number') { const n = numStr(v); if (n == null) { t = 'e'; body = '<v>#NUM!</v>'; } else body = '<v>' + n + '</v>'; }
      else if (typeof v === 'string') { t = 's'; body = '<v>' + ctx.sst.add(v.length > 32767 ? v.slice(0, 32767) : v, headerName != null ? null : cell.rt) + '</v>'; }
      else if (typeof v === 'boolean') { t = 'b'; body = '<v>' + (v ? 1 : 0) + '</v>'; }
      else if (M.isErr(v)) { t = 'e'; body = '<v>' + esc(v.e) + '</v>'; }
    }
    if (t) s += ' t="' + t + '"';
    return body ? s + '>' + body + '</c>' : s + '/>';
  }

  function sheetDataXml(sh, ctx, headers) {
    const out = ['<sheetData>'];
    const def = M.defaultRowPt(sh);
    sh.rows.forEach((row, r) => {
      if (!row || r >= M.MAXR) return;
      const cells = row.cells;
      let first = -1, last = -1;
      let parts = [];
      let extra = null;
      if (headers && headers.size) for (const [k, name] of headers) { const rr = Math.floor(k / 16384), cc = k % 16384; if (rr === r && !cells[cc]) (extra || (extra = [])).push([cc, cellXml({ v: name }, rr, cc, ctx, name)]); }
      cells.forEach((cell, c) => {
        if (!cell || c >= M.MAXC) return;
        if (cell.v == null && cell.f == null && !cell.s && !cell.dt && !(headers && headers.has(r * 16384 + c))) return;
        parts.push(extra ? [c, cellXml(cell, r, c, ctx, headers.get(r * 16384 + c))] : cellXml(cell, r, c, ctx, headers ? headers.get(r * 16384 + c) : undefined));
        if (first < 0) first = c;
        last = c;
      });
      if (extra) { parts = parts.concat(extra).sort((a, b) => a[0] - b[0]); first = parts[0][0]; last = parts[parts.length - 1][0]; parts = parts.map((x) => x[1]); }
      const a = {};
      if (row.s) { a.s = row.s; a.customFormat = '1'; }
      if (row.ht != null && (row.customHeight || row.auto == null)) { a.ht = row.ht; if (row.customHeight) a.customHeight = '1'; }
      else if (row.auto != null && Math.abs(row.auto - def) > 0.01) a.ht = row.auto;
      if (row.hidden) a.hidden = '1';
      if (row.level) a.outlineLevel = row.level;
      if (row.collapsed) a.collapsed = '1';
      if (row.thickTop) a.thickTop = '1';
      if (row.thickBot) a.thickBot = '1';
      if (!parts.length && !Object.keys(a).length) return;
      if (a.hidden && a.ht == null) a.ht = def;
      out.push(`<row r="${r + 1}"${first >= 0 ? ` spans="${first + 1}:${last + 1}"` : ''}${attrs(a)}${parts.length ? '>' + parts.join('') + '</row>' : '/>'}`);
    });
    out.push('</sheetData>');
    return out.join('');
  }

  function autoFilterXml(f, ref) {
    if (!f) return '';
    ref = ref || f.ref;
    if (!ref) return '';
    let s = `<autoFilter ref="${rangeName(ref)}"`;
    const cols = (f.cols || []).filter((c) => c.values || c.custom || c.top10 || c.dynamic || c.color || c.icon || c.blank || c.dates || c.hiddenButton);
    if (!cols.length && !f.sort) return s + '/>';
    s += '>';
    for (const c of cols) {
      s += `<filterColumn colId="${c.col}"${c.hiddenButton ? ' hiddenButton="1"' : ''}>`;
      if (c.values || c.blank || c.dates) {
        s += `<filters${c.blank ? ' blank="1"' : ''}>` + (c.values || []).map((v) => `<filter val="${esc(v)}"/>`).join('') +
          (c.dates || []).map((d) => `<dateGroupItem${attrs({ year: d.y, month: d.g !== 'year' ? d.m || undefined : undefined, day: /day|hour|minute|second/.test(d.g) ? d.d || undefined : undefined, hour: /hour|minute|second/.test(d.g) ? d.H : undefined, minute: /minute|second/.test(d.g) ? d.M : undefined, second: d.g === 'second' ? d.S : undefined, dateTimeGrouping: d.g })}/>`).join('') + '</filters>';
      } else if (c.custom) {
        s += `<customFilters${c.custom.and ? ' and="1"' : ''}>` + c.custom.list.map((x) => `<customFilter${attrs({ operator: x.op && x.op !== 'equal' ? x.op : undefined, val: x.val })}/>`).join('') + '</customFilters>';
      } else if (c.top10) s += `<top10${attrs({ top: c.top10.top ? undefined : '0', percent: c.top10.percent ? '1' : undefined, val: c.top10.val, filterVal: c.top10.filterVal })}/>`;
      else if (c.dynamic) s += `<dynamicFilter${attrs({ type: c.dynamic.type, val: c.dynamic.val, maxVal: c.dynamic.maxVal })}/>`;
      else if (c.color) s += `<colorFilter${attrs({ dxfId: c.color.dxf >= 0 ? c.color.dxf : undefined, cellColor: c.color.cell ? undefined : '0' })}/>`;
      else if (c.icon) s += `<iconFilter${attrs({ iconSet: c.icon.set, iconId: c.icon.id })}/>`;
      s += '</filterColumn>';
    }
    if (f.sort) s += sortStateXml(f.sort);
    return s + '</autoFilter>';
  }
  function sortStateXml(ss) {
    if (!ss || !ss.ref) return '';
    return `<sortState${attrs({ columnSort: ss.columnSort ? '1' : undefined, caseSensitive: ss.caseSensitive ? '1' : undefined, ref: rangeName(ss.ref) })}>` +
      (ss.keys || []).filter((k) => k.ref).map((k) => `<sortCondition${attrs({ descending: k.desc ? '1' : undefined, sortBy: k.by && k.by !== 'value' ? k.by : undefined, ref: rangeName(k.ref), customList: k.list, dxfId: k.dxf })}/>`).join('') + '</sortState>';
  }
  const cfvoXml = (v) => `<cfvo${attrs({ type: v.type, val: v.val, gte: v.gte === false ? '0' : undefined })}/>`;
  const guid = () => '{' + [8, 4, 4, 4, 12].map((n) => Array.from({ length: n }, () => '0123456789ABCDEF'[Math.floor(Math.random() * 16)]).join('')).join('-') + '}';
  function dxfInner(d) { d = d || {}; let s = ''; if (d.font) s += '<font>' + fontXml(d.font, true).replace(/<rFont [^>]*>/, '') + '</font>'; if (d.fill) s += fillXml(d.fill, true); if (d.border) s += borderXml(d.border); return s; }
  /** icon sets the 2006 schema does not know (Excel 2010+), and icon sets with per-threshold custom icons */
  const X14_ICON = /^(3Triangles|3Stars|5Boxes|NoIcons)$/;
  const crossSheet = (r) => !r.scale && !r.bar && !r.icons && r.f && r.f.some((f) => /!/.test(String(f)) && !/^"/.test(String(f)));
  const iconExt = (r) => r.icons && (X14_ICON.test(r.icons.set || '') || (r.icons.custom && r.icons.custom.length) || r.icons.cfvo.some((v) => /!/.test(String(v.val || ''))));
  /** rules that must live in the Excel 2010 extension list */
  const extRule = (r) => crossSheet(r) || iconExt(r);
  function cfExtXml(sh, wb) {
    const items = [];
    const all = sh.cf.flatMap((x) => x.rules).slice().sort((a, b) => (a.priority || 0) - (b.priority || 0));
    for (const cf of sh.cf) for (const r of cf.rules) if (extRule(r) && cf.ranges.length) items.push({ cf, r });
    if (!items.length) return '';
    const body = (r) => {
      if (r.icons) {
        const ic = r.icons;
        return `<x14:iconSet${attrs({ iconSet: ic.set && ic.set !== '3TrafficLights1' ? ic.set : undefined, showValue: ic.showValue === false ? '0' : undefined, percent: ic.percent === false ? '0' : undefined, reverse: ic.reverse ? '1' : undefined, custom: ic.custom && ic.custom.length ? '1' : undefined })}>` +
          ic.cfvo.map((v) => `<x14:cfvo${attrs({ type: v.type, gte: v.gte === false ? '0' : undefined })}>${v.val != null ? `<xm:f>${esc(v.val)}</xm:f>` : ''}</x14:cfvo>`).join('') +
          (ic.custom || []).map((c) => `<x14:cfIcon iconSet="${esc(c.set || 'NoIcons')}" iconId="${c.id || 0}"/>`).join('') + '</x14:iconSet>';
      }
      return r.f.slice(0, 3).map((f) => `<xm:f>${esc(f)}</xm:f>`).join('') + (r.dxf != null && wb.dxfs[r.dxf] ? `<x14:dxf>${dxfInner(wb.dxfs[r.dxf])}</x14:dxf>` : '');
    };
    return '<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:conditionalFormattings>' +
      items.map(({ cf, r }) => `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule${attrs({ type: r.icons ? 'iconSet' : r.type, priority: all.indexOf(r) + 1, stopIfTrue: r.stop ? '1' : undefined, operator: r.icons ? undefined : r.op, text: r.icons ? undefined : r.text, id: guid() })}>` + body(r) + `</x14:cfRule><xm:sqref>${sqref(cf.ranges)}</xm:sqref></x14:conditionalFormatting>`).join('') +
      '</x14:conditionalFormattings></ext>';
  }
  function cfXml(sh) {
    let prio = 0;
    const out = [];
    const all = [];
    for (const cf of sh.cf) for (const r of cf.rules) all.push(r);
    const order = all.slice().sort((a, b) => (a.priority || 0) - (b.priority || 0));
    const pr = new Map(order.map((r, i) => [r, i + 1]));
    for (const cf of sh.cf) {
      const rules = cf.rules.filter((r) => !extRule(r));
      if (!rules.length || !cf.ranges.length) continue;
      let s = `<conditionalFormatting${cf.pivot ? ' pivot="1"' : ''} sqref="${sqref(cf.ranges)}">`;
      for (const r of rules) {
        prio = pr.get(r);
        const a = { type: r.type, dxfId: r.dxf != null && r.dxf >= 0 && !/colorScale|dataBar|iconSet/.test(r.type) ? r.dxf : undefined, priority: prio, stopIfTrue: r.stop ? '1' : undefined, aboveAverage: r.below ? '0' : undefined, percent: r.percent ? '1' : undefined, bottom: r.bottom ? '1' : undefined, operator: r.op, text: r.text, timePeriod: r.period, rank: r.rank, stdDev: r.stdDev || undefined, equalAverage: r.equal ? '1' : undefined };
        s += `<cfRule${attrs(a)}>`;
        if (r.scale) s += '<colorScale>' + r.scale.cfvo.map(cfvoXml).join('') + r.scale.colors.map((c) => colorEl('color', c) || '<color rgb="FF000000"/>').join('') + '</colorScale>';
        else if (r.bar) s += `<dataBar${attrs({ minLength: r.bar.minLength !== 10 ? r.bar.minLength : undefined, maxLength: r.bar.maxLength !== 90 ? r.bar.maxLength : undefined, showValue: r.bar.showValue === false ? '0' : undefined })}>` + r.bar.cfvo.slice(0, 2).map((v) => cfvoXml({ type: v.type === 'autoMin' ? 'min' : v.type === 'autoMax' ? 'max' : v.type, val: v.type === 'min' || v.type === 'max' ? undefined : v.val })).join('') + (colorEl('color', r.bar.color) || '<color rgb="FF638EC6"/>') + '</dataBar>';
        else if (r.icons) s += `<iconSet${attrs({ iconSet: r.icons.set && r.icons.set !== '3TrafficLights1' ? r.icons.set : undefined, showValue: r.icons.showValue === false ? '0' : undefined, percent: r.icons.percent === false ? '0' : undefined, reverse: r.icons.reverse ? '1' : undefined })}>` + r.icons.cfvo.map(cfvoXml).join('') + '</iconSet>';
        if (r.f && !r.scale && !r.bar && !r.icons) for (const f of r.f.slice(0, 3)) s += '<formula>' + esc(f) + '</formula>';
        s += '</cfRule>';
      }
      out.push(s + '</conditionalFormatting>');
    }
    return out.join('');
  }
  function dvXml(sh) {
    const list = sh.dv.filter((d) => d.ranges && d.ranges.length);
    if (!list.length) return '';
    return `<dataValidations count="${list.length}">` + list.map((d) => {
      const a = { type: d.type && d.type !== 'none' ? d.type : undefined, errorStyle: d.errorStyle && d.errorStyle !== 'stop' ? d.errorStyle : undefined, imeMode: d.ime, operator: d.op && d.op !== 'between' ? d.op : undefined, allowBlank: d.allowBlank ? '1' : undefined, showDropDown: d.showDrop === false ? '1' : undefined, showInputMessage: d.showInput ? '1' : undefined, showErrorMessage: d.showError ? '1' : undefined, errorTitle: d.errorTitle, error: d.error, promptTitle: d.promptTitle, prompt: d.prompt, sqref: sqref(d.ranges) };
      return `<dataValidation${attrs(a)}>` + (d.f1 != null ? '<formula1>' + esc(d.f1) + '</formula1>' : '') + (d.f2 != null ? '<formula2>' + esc(d.f2) + '</formula2>' : '') + '</dataValidation>';
    }).join('') + '</dataValidations>';
  }
  function sparklineXml(sh, wb) {
    const cfx = cfExtXml(sh, wb);
    if (!sh.sparklines || !sh.sparklines.length) return cfx ? '<extLst>' + cfx + '</extLst>' : '';
    const g = sh.sparklines.map((s) => `<x14:sparklineGroup${attrs({ manualMax: s.maxAxisType === 'custom' && s.manualMax != null ? s.manualMax : undefined, manualMin: s.minAxisType === 'custom' && s.manualMin != null ? s.manualMin : undefined, lineWeight: s.lineWeight !== 0.75 ? s.lineWeight : undefined, type: s.type !== 'line' ? s.type : undefined, dateAxis: s.dateAxis ? '1' : undefined, displayEmptyCellsAs: s.displayEmptyCellsAs || 'gap', markers: s.markers ? '1' : undefined, high: s.high ? '1' : undefined, low: s.low ? '1' : undefined, first: s.first ? '1' : undefined, last: s.last ? '1' : undefined, negative: s.negative ? '1' : undefined, displayXAxis: s.displayXAxis ? '1' : undefined, displayHidden: s.displayHidden ? '1' : undefined, minAxisType: s.minAxisType && s.minAxisType !== 'individual' ? s.minAxisType : undefined, maxAxisType: s.maxAxisType && s.maxAxisType !== 'individual' ? s.maxAxisType : undefined, rightToLeft: s.rightToLeft ? '1' : undefined })}>` +
      [['colorSeries', s.color], ['colorNegative', s.negColor], ['colorAxis', s.axisColor || { rgb: 'FF000000' }], ['colorMarkers', s.markerColor], ['colorFirst', s.firstColor], ['colorLast', s.lastColor], ['colorHigh', s.highColor], ['colorLow', s.lowColor]].map(([t, c]) => colorEl('x14:' + t, c || { rgb: 'FF376092' })).join('') +
      '<x14:sparklines>' + s.items.map((i) => `<x14:sparkline><xm:f>${esc(i.f)}</xm:f><xm:sqref>${esc(i.sqref)}</xm:sqref></x14:sparkline>`).join('') + '</x14:sparklines></x14:sparklineGroup>').join('');
    return `<extLst>${cfx}<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${g}</x14:sparklineGroups></ext></extLst>`;
  }
  function printXml(sh) {
    const p = sh.print;
    let s = '';
    if (p.gridLines || p.headings || p.hCenter || p.vCenter) s += `<printOptions${attrs({ horizontalCentered: p.hCenter ? '1' : undefined, verticalCentered: p.vCenter ? '1' : undefined, headings: p.headings ? '1' : undefined, gridLines: p.gridLines ? '1' : undefined })}/>`;
    const m = p.margins || {};
    s += `<pageMargins left="${m.l != null ? m.l : 0.75}" right="${m.r != null ? m.r : 0.75}" top="${m.t != null ? m.t : 1}" bottom="${m.b != null ? m.b : 1}" header="${m.header != null ? m.header : 0.5}" footer="${m.footer != null ? m.footer : 0.5}"/>`;
    const ps = { paperSize: p.paper && p.paper !== 1 ? p.paper : undefined, scale: p.scale && p.scale !== 100 ? Math.round(p.scale) : undefined, firstPageNumber: p.firstPage != null ? p.firstPage : undefined, fitToWidth: p.fit && p.fitW !== 1 ? (p.fitW == null ? 0 : p.fitW) : undefined, fitToHeight: p.fit && p.fitH !== 1 ? (p.fitH == null ? 0 : p.fitH) : undefined, pageOrder: p.pageOrder === 'overThenDown' ? 'overThenDown' : undefined, orientation: p.orientation === 'landscape' ? 'landscape' : 'portrait', blackAndWhite: p.bw ? '1' : undefined, draft: p.draft ? '1' : undefined, cellComments: p.comments && p.comments !== 'none' ? p.comments : undefined, useFirstPageNumber: p.firstPage != null ? '1' : undefined, errors: p.errors && p.errors !== 'displayed' ? p.errors : undefined };
    s += `<pageSetup${attrs(ps)}/>`;
    const hf = [['oddHeader', p.header], ['oddFooter', p.footer], ['evenHeader', p.diffOddEven && p.evenHeader], ['evenFooter', p.diffOddEven && p.evenFooter], ['firstHeader', p.diffFirst && p.firstHeader], ['firstFooter', p.diffFirst && p.firstFooter]].filter(([, v]) => v);
    if (hf.length || p.diffFirst || p.diffOddEven) s += `<headerFooter${attrs({ differentOddEven: p.diffOddEven ? '1' : undefined, differentFirst: p.diffFirst ? '1' : undefined, scaleWithDoc: p.hfScale === false ? '0' : undefined, alignWithMargins: p.hfAlign === false ? '0' : undefined })}>` + hf.map(([k, v]) => `<${k}>${escT(v)}</${k}>`).join('') + '</headerFooter>';
    const rb = (p.rowBreaks || []).filter((b) => b > 0), cb = (p.colBreaks || []).filter((b) => b > 0);
    if (rb.length) s += `<rowBreaks count="${rb.length}" manualBreakCount="${rb.length}">` + rb.map((b) => `<brk id="${b}" max="16383" man="1"/>`).join('') + '</rowBreaks>';
    if (cb.length) s += `<colBreaks count="${cb.length}" manualBreakCount="${cb.length}">` + cb.map((b) => `<brk id="${b}" max="1048575" man="1"/>`).join('') + '</colBreaks>';
    return s;
  }
  const keepXml = (sh, name) => { const k = sh.extra && sh.extra.keep && sh.extra.keep.find((x) => x.name === name); return k ? k.xml : ''; };

  /* comments: the comments part and its VML drawing */
  function commentsXml(sh) {
    const list = Array.from(sh.comments.values()).filter((c) => c.r < M.MAXR && c.c < M.MAXC).sort((a, b) => a.r - b.r || a.c - b.c);
    const authors = [];
    const aid = (a) => { a = a || ''; let i = authors.indexOf(a); if (i < 0) { i = authors.length; authors.push(a); } return i; };
    const items = list.map((cm) => {
      const id = aid(cm.author);
      let text;
      if (cm.runs && cm.runs.length) text = cm.runs.map((r) => '<r>' + (r.font ? '<rPr>' + fontXml(Object.assign({ sz: 9, name: 'Tahoma', family: 2 }, r.font), true) + '</rPr>' : '<rPr><sz val="9"/><rFont val="Tahoma"/><family val="2"/></rPr>') + tEl(r.t || '') + '</r>').join('');
      else {
        const t = cm.text || '';
        const au = cm.author && t.startsWith(cm.author + ':') ? cm.author + ':' : '';
        text = (au ? `<r><rPr><b/><sz val="9"/><rFont val="Tahoma"/><family val="2"/></rPr>${tEl(au)}</r>` : '') + `<r><rPr><sz val="9"/><rFont val="Tahoma"/><family val="2"/></rPr>${tEl(t.slice(au.length))}</r>`;
      }
      return `<comment ref="${cellName(cm.r, cm.c)}" authorId="${id}"><text>${text}</text></comment>`;
    });
    if (!authors.length) aid('');
    return HDR + `<comments xmlns="${NS_MAIN}"><authors>` + authors.map((a) => `<author>${escT(a)}</author>`).join('') + '</authors><commentList>' + items.join('') + '</commentList></comments>';
  }
  function vmlXml(sh, idmap) {
    const list = Array.from(sh.comments.values()).filter((c) => c.r < M.MAXR && c.c < M.MAXC);
    let s = '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
      `<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${idmap}"/></o:shapelayout>` +
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';
    list.forEach((cm, i) => {
      const w = cm.w || 108, hgt = cm.h || 59.25;
      const anchor = cm.anchor && cm.anchor.length === 8 ? cm.anchor.join(', ') : `${cm.c + 1}, 15, ${Math.max(0, cm.r - 1)}, 10, ${cm.c + 3}, 15, ${Math.max(0, cm.r - 1) + 4}, 4`;
      const fill = cm.fill && /^#[0-9a-f]{6}$/i.test(cm.fill) ? cm.fill : '#ffffe1';
      s += `<v:shape id="_x0000_s${idmap * 1024 + i + 1}" type="#_x0000_t202" style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:${w}pt;height:${hgt}pt;z-index:${i + 1};visibility:${cm.visible ? 'visible' : 'hidden'}" fillcolor="${fill}" o:insetmode="auto">` +
        `<v:fill color2="${fill}"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>` +
        `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${anchor}</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${cm.r}</x:Row><x:Column>${cm.c}</x:Column>${cm.visible ? '<x:Visible/>' : ''}</x:ClientData></v:shape>`;
    });
    return s + '</xml>';
  }

  function tableXml(t, wb) {
    const ref = t.ref;
    const a = { xmlns: NS_MAIN, id: t.id, name: t.dname || t.name, displayName: t.name, ref: rangeName(ref) };
    if (t.header === false) a.headerRowCount = '0';
    if (t.totals) a.totalsRowCount = '1'; else a.totalsRowShown = '0';
    if (t.comment) a.comment = t.comment;
    let s = HDR + `<table${attrs(a)}>`;
    if (t.autoFilter !== false && t.header !== false) {
      const afRef = { r1: ref.r1, c1: ref.c1, r2: t.totals ? ref.r2 - 1 : ref.r2, c2: ref.c2 };
      s += autoFilterXml(t.filter || { cols: [] }, afRef);
    }
    s += `<tableColumns count="${t.columns.length}">` + t.columns.map((c, i) => {
      const ca = { id: i + 1, name: c.name, totalsRowFunction: t.totals && c.totalsFn && c.totalsFn !== 'none' ? c.totalsFn : undefined, totalsRowLabel: t.totals && c.totalsLabel != null ? c.totalsLabel : undefined, dataDxfId: c.dataDxf };
      const inner = (c.calc ? `<calculatedColumnFormula>${esc(c.calc)}</calculatedColumnFormula>` : '') + (t.totals && c.totalsFn === 'custom' && c.totalsFormula ? `<totalsRowFormula>${esc(c.totalsFormula)}</totalsRowFormula>` : '');
      return `<tableColumn${attrs(ca)}${inner ? '>' + inner + '</tableColumn>' : '/>'}`;
    }).join('') + '</tableColumns>';
    const st = t.style;
    if (st) s += `<tableStyleInfo${attrs({ name: st.name || undefined, showFirstColumn: st.first ? '1' : '0', showLastColumn: st.last ? '1' : '0', showRowStripes: st.rowStripes ? '1' : '0', showColumnStripes: st.colStripes ? '1' : '0' })}/>`;
    void wb;
    return s + '</table>';
  }
  /** header cell text of every table column (unique, non-empty, as Excel requires) */
  function syncTableHeaders(wb, sh) {
    const out = new Map();
    for (const t of wb.tables) {
      if (t.sheet !== sh) continue;
      const ncol = t.ref.c2 - t.ref.c1 + 1;
      while (t.columns.length < ncol) t.columns.push({ name: 'Column' + (t.columns.length + 1), id: t.columns.length + 1 });
      if (t.columns.length > ncol) t.columns.length = ncol;
      const used = new Set();
      t.columns.forEach((col, i) => {
        let name = col.name;
        if (t.header !== false) {
          const v = sh.val(t.ref.r1, t.ref.c1 + i);
          name = v == null || v === '' ? '' : M.isErr(v) ? v.e : typeof v === 'number' ? L.editor && L.editor.cellText ? String(L.editor.cellText(sh, t.ref.r1, t.ref.c1 + i)) : String(v) : String(v);
        }
        name = String(name || '').replace(/[\r\n]+/g, ' ');
        if (!name) name = 'Column' + (i + 1);
        let n = name, k = 2;
        while (used.has(n.toLowerCase())) n = name + k++;
        used.add(n.toLowerCase());
        col.name = n;
        if (t.header !== false) { const v = sh.val(t.ref.r1, t.ref.c1 + i); const cell = sh.get(t.ref.r1, t.ref.c1 + i); if (v !== n || (cell && cell.f != null)) out.set(t.ref.r1 * 16384 + t.ref.c1 + i, n); }
      });
    }
    return out;
  }

  /* ------------------------------------------------------------ drawings */
  const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing', NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main', NS_C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const emu = (pt) => Math.max(0, Math.round((pt || 0) * EMU));
  const ptXml = (tag, p) => `<xdr:${tag}><xdr:col>${p.c}</xdr:col><xdr:colOff>${emu(p.cOff)}</xdr:colOff><xdr:row>${p.r}</xdr:row><xdr:rowOff>${emu(p.rOff)}</xdr:rowOff></xdr:${tag}>`;
  /** size in points of an anchor (for xfrm / ext) */
  function anchorSize(sh, a) {
    if (a.type !== 'two' || !a.to) return { w: a.w || 72, h: a.h || 72 };
    const g = L.layout && L.layout.geo ? L.layout.geo(sh) : null;
    if (!g) return { w: 72, h: 72 };
    const x1 = g.cols.pos(a.from.c) * 0.75 + (a.from.cOff || 0), x2 = g.cols.pos(a.to.c) * 0.75 + (a.to.cOff || 0);
    const y1 = g.rows.pos(a.from.r) * 0.75 + (a.from.rOff || 0), y2 = g.rows.pos(a.to.r) * 0.75 + (a.to.rOff || 0);
    return { w: Math.max(1, x2 - x1), h: Math.max(1, y2 - y1) };
  }
  function shapeXml(d, id, sz) {
    const fill = d.fill === null || d.fill === 'none' ? '<a:noFill/>' : `<a:solidFill><a:srgbClr val="${String(d.fill || '#FFFFFF').replace('#', '').toUpperCase()}"/></a:solidFill>`;
    const ln = d.line === null || d.line === 'none' ? '<a:ln><a:noFill/></a:ln>' : `<a:ln w="${Math.round((d.lineW || 0.75) * EMU)}"><a:solidFill><a:srgbClr val="${String(d.line || '#000000').replace('#', '').toUpperCase()}"/></a:solidFill></a:ln>`;
    const paras = String(d.text || '').split('\n');
    const f = d.font || {};
    const rpr = `<a:rPr lang="en-US" sz="${Math.round((f.sz || 10) * 100)}"${f.b ? ' b="1"' : ''}${f.i ? ' i="1"' : ''}>${f.color ? `<a:solidFill><a:srgbClr val="${String(f.color).replace('#', '').toUpperCase()}"/></a:solidFill>` : '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'}<a:latin typeface="${esc(f.name || 'Arial')}"/></a:rPr>`;
    const body = d.text != null && d.text !== '' ? `<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="${d.vAlign || 't'}"/><a:lstStyle/>` + paras.map((p) => `<a:p><a:pPr algn="${d.align || 'l'}"/>${p ? `<a:r>${rpr}<a:t>${escT(p)}</a:t></a:r>` : `<a:endParaRPr lang="en-US" sz="${Math.round((f.sz || 10) * 100)}"/>`}</a:p>`).join('') + '</xdr:txBody>'
      : '<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" rtlCol="0" anchor="t"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" sz="1100"/></a:p></xdr:txBody>';
    return `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${esc(d.name || 'Shape ' + id)}"/><xdr:cNvSpPr${d.txBox ? ' txBox="1"' : ''}/></xdr:nvSpPr>` +
      `<xdr:spPr><a:xfrm${d.rot ? ` rot="${Math.round(d.rot * 60000)}"` : ''}><a:off x="0" y="0"/><a:ext cx="${emu(sz.w)}" cy="${emu(sz.h)}"/></a:xfrm><a:prstGeom prst="${esc(d.geom || 'rect')}"><a:avLst/></a:prstGeom>${fill}${ln}</xdr:spPr>${body}</xdr:sp>`;
  }
  /** the chart part for a chart drawing */
  function chartPartXml(d, wb) {
    if (d.chart && (d.chart.dirty || !d.xml) && L.xchart && L.xchart.toXml) return L.xchart.toXml(d.chart, wb, d);
    let x = d.xml || '';
    if (d.chart && d.chart.refs && L.xchart && L.xchart.patchRefs) x = L.xchart.patchRefs(x, d.chart.refs);
    /* parts the chart points at that we do not carry (user shapes, embedded data) */
    x = x.replace(/<(\w+:)?userShapes\b[^>]*\/>/g, '').replace(/<(\w+:)?externalData\b[^>]*\/>/g, '').replace(/<(\w+:)?externalData\b[\s\S]*?<\/(\w+:)?externalData>/g, '');
    x = x.replace(/\sr:(embed|link|id)="[^"]*"/g, '');
    return x;
  }

  /* ------------------------------------------------------------ default theme */
  function themeXml(wb) {
    const th = wb.theme || M.DEFAULT_THEME;
    const c = th.colors || M.DEFAULT_THEME.colors;
    const hex = (i) => String(c[i] || M.DEFAULT_THEME.colors[i]).replace('#', '').toUpperCase();
    const clr = (tag, i) => `<a:${tag}><a:srgbClr val="${hex(i)}"/></a:${tag}>`;
    const phSolid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
    return HDR + `<a:theme xmlns:a="${NS_A}" name="${esc(th.name || 'Office Theme')}"><a:themeElements>` +
      `<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="${hex(1)}"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="${hex(0)}"/></a:lt1>${clr('dk2', 3)}${clr('lt2', 2)}${clr('accent1', 4)}${clr('accent2', 5)}${clr('accent3', 6)}${clr('accent4', 7)}${clr('accent5', 8)}${clr('accent6', 9)}${clr('hlink', 10)}${clr('folHlink', 11)}</a:clrScheme>` +
      `<a:fontScheme name="Office"><a:majorFont><a:latin typeface="${esc(th.major || 'Cambria')}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="${esc(th.minor || 'Calibri')}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>` +
      `<a:fmtScheme name="Office"><a:fillStyleLst>${phSolid}${phSolid}${phSolid}</a:fillStyleLst>` +
      '<a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst>' +
      '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
      `<a:bgFillStyleLst>${phSolid}${phSolid}${phSolid}</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`;
  }
  const META_XML = HDR + `<metadata xmlns="${NS_MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;

  /* ------------------------------------------------------------ workbook */
  W.bytes = async function (wb, opts) {
    opts = opts || {};
    const type = opts.type || 'xlsx';
    const macro = (type === 'xlsm' || type === 'xltm') && wb.extra && wb.extra.vba;
    const files = [];
    const overrides = new Map();
    const defaults = new Map([['rels', 'application/vnd.openxmlformats-package.relationships+xml'], ['xml', 'application/xml']]);
    const add = (name, data, ct) => { files.push({ name, data }); if (ct) overrides.set('/' + name, ct); };
    const wbRels = new Rels();
    const sst = new SST();
    const ctx = { sst, dynamic: false };
    const mediaPath = new Map(); /* media id → part name */
    let nImg = 0, nDrawing = 0, nChart = 0, nComment = 0, nTable = 0, nVml = 0;
    const mediaPart = (id) => {
      if (mediaPath.has(id)) return mediaPath.get(id);
      const m = wb.media && wb.media.get(id);
      if (!m) return null;
      let ext = String(m.ext || 'png').toLowerCase();
      if (ext === 'jpg') ext = 'jpeg';
      const name = `xl/media/image${++nImg}.${ext}`;
      defaults.set(ext, IMG_CT[ext] || m.type || 'application/octet-stream');
      add(name, m.bytes);
      mediaPath.set(id, name);
      return name;
    };
    /* table ids unique across the workbook */
    let tid = 0;
    for (const t of wb.tables) { if (!(t.id > tid)) t.id = tid + 1; tid = Math.max(tid, t.id); }
    const usedT = new Set();
    for (const t of wb.tables) { while (usedT.has(t.id)) t.id = ++tid; usedT.add(t.id); }
    const sheetEntries = [];
    for (let si = 0; si < wb.sheets.length; si++) {
      const sh = wb.sheets[si];
      const rels = new Rels();
      if (sh.kind === 'chartsheet') {
        const part = `xl/chartsheets/sheet${si + 1}.xml`;
        let body = HDR + `<chartsheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">` + (sh.tabColor ? `<sheetPr>${colorEl('tabColor', sh.tabColor)}</sheetPr>` : '') + `<sheetViews><sheetView${si === wb.active ? ' tabSelected="1"' : ''}${sh.view.zoom && sh.view.zoom !== 100 ? ` zoomScale="${Math.round(sh.view.zoom)}"` : ''}${sh.view.zoomToFit ? ' zoomToFit="1"' : ''} workbookViewId="0"/></sheetViews>`;
        const m = sh.print.margins;
        body += `<pageMargins left="${m.l}" right="${m.r}" top="${m.t}" bottom="${m.b}" header="${m.header}" footer="${m.footer}"/><pageSetup orientation="${sh.print.orientation === 'portrait' ? 'portrait' : 'landscape'}"/>`;
        const dr = await drawingPart(sh, rels);
        if (dr) body += `<drawing r:id="${dr}"/>`;
        body += '</chartsheet>';
        add(part, body, CT.chartsheet);
        if (rels.list.length) add(relsPath(part), rels.xml());
        sheetEntries.push({ sh, part, kind: 'chartsheet' });
        continue;
      }
      const part = `xl/worksheets/sheet${si + 1}.xml`;
      const headers = syncTableHeaders(wb, sh);
      const out = [HDR, `<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">`];
      /* sheetPr */
      const op = sh.outline || {};
      const prIn = (sh.tabColor ? colorEl('tabColor', sh.tabColor) : '') + (op.below === false || op.right === false || op.applyStyles ? `<outlinePr${attrs({ applyStyles: op.applyStyles ? '1' : undefined, summaryBelow: op.below === false ? '0' : undefined, summaryRight: op.right === false ? '0' : undefined })}/>` : '') + (sh.print.fit ? '<pageSetUpPr fitToPage="1"/>' : '');
      const prA = attrs({ codeName: sh.codeName, filterMode: sh.filterMode || (sh.rows.some((r) => r && r.filtered)) ? '1' : undefined });
      if (prIn || prA) out.push(`<sheetPr${prA}${prIn ? '>' + prIn + '</sheetPr>' : '/>'}`);
      const maxR = Math.max(0, sh.maxR), maxC = Math.max(0, sh.maxC);
      out.push(`<dimension ref="${sh.maxR < 0 ? 'A1' : rangeName({ r1: 0, c1: 0, r2: Math.min(maxR, M.MAXR - 1), c2: Math.min(maxC, M.MAXC - 1) })}"/>`);
      out.push(sheetViewXml(sh, wb, si));
      const fp = { baseColWidth: sh.baseColW != null && sh.baseColW !== 8 ? sh.baseColW : undefined, defaultColWidth: sh.defColW != null ? sh.defColW : undefined, defaultRowHeight: sh.defRowH != null ? sh.defRowH : M.defaultRowPt(sh), customHeight: sh.defRowH != null ? '1' : undefined, zeroHeight: sh.zeroHeight ? '1' : undefined, outlineLevelRow: op.levelRow || undefined, outlineLevelCol: op.levelCol || undefined };
      out.push(`<sheetFormatPr${attrs(fp)}/>`);
      out.push(colsXml(sh));
      out.push(sheetDataXml(sh, ctx, headers));
      if (sh.protection) {
        const a = Object.assign({}, sh.protection.attrs || {});
        a.sheet = '1';
        out.push('<sheetProtection' + Object.keys(a).map((k) => ` ${k}="${esc(a[k])}"`).join('') + '/>');
      }
      out.push(keepXml(sh, 'protectedRanges'));
      if (sh.scenarios && sh.scenarios.length) {
        const all = [];
        for (const x of sh.scenarios) for (const c of x.cells) if (!all.some((q) => q.r === c.r && q.c === c.c)) all.push(c);
        out.push(`<scenarios current="0" show="0" sqref="${all.map((c) => cellName(c.r, c.c)).join(' ')}">` + sh.scenarios.map((x) => `<scenario${attrs({ name: x.name, locked: '1', count: x.cells.length, user: x.user || undefined, comment: x.comment || undefined })}>` + x.cells.map((c) => `<inputCells r="${cellName(c.r, c.c)}" val="${esc(c.v == null ? '' : c.v)}"/>`).join('') + '</scenario>').join('') + '</scenarios>');
      } else if (!sh.scenarios) out.push(keepXml(sh, 'scenarios'));
      if (sh.autoFilter && sh.autoFilter.ref) out.push(autoFilterXml(sh.autoFilter));
      if (sh.sortState) out.push(sortStateXml(sh.sortState));
      out.push(keepXml(sh, 'dataConsolidate'), keepXml(sh, 'customSheetViews'));
      const merges = sh.merges.filter((m) => m.r1 !== m.r2 || m.c1 !== m.c2);
      if (merges.length) out.push(`<mergeCells count="${merges.length}">` + merges.map((m) => `<mergeCell ref="${rangeName(m)}"/>`).join('') + '</mergeCells>');
      out.push(keepXml(sh, 'phoneticPr'));
      out.push(cfXml(sh));
      out.push(dvXml(sh));
      const links = sh.links.filter((l) => l.ref);
      if (links.length) {
        out.push('<hyperlinks>' + links.map((l) => {
          const a = { ref: rangeName(l.ref) };
          let loc = l.location;
          if (l.target && /^#/.test(l.target)) loc = l.target.slice(1);
          else if (l.target) a['r:id'] = rels.add('hyperlink', l.target, true);
          if (loc != null) a.location = loc;
          if (l.display != null) a.display = l.display;
          if (l.tooltip != null) a.tooltip = l.tooltip;
          return `<hyperlink${attrs(a)}/>`;
        }).join('') + '</hyperlinks>');
      }
      out.push(printXml(sh));
      out.push(keepXml(sh, 'cellWatches'), keepXml(sh, 'ignoredErrors'), keepXml(sh, 'smartTags'));
      const dr = await drawingPart(sh, rels);
      if (dr) out.push(`<drawing r:id="${dr}"/>`);
      if (sh.comments.size) {
        const cPart = `xl/comments${++nComment}.xml`;
        add(cPart, commentsXml(sh), CT.comments);
        rels.add('comments', '../' + cPart.slice(3));
        const vPart = `xl/drawings/vmlDrawing${++nVml}.vml`;
        defaults.set('vml', CT.vml);
        add(vPart, vmlXml(sh, si + 1));
        out.push(`<legacyDrawing r:id="${rels.add('vmlDrawing', '../drawings/' + vPart.split('/').pop())}"/>`);
      }
      if (sh.background) { const mp = mediaPart(sh.background); if (mp) out.push(`<picture r:id="${rels.add('image', '../media/' + mp.split('/').pop())}"/>`); }
      out.push(keepXml(sh, 'webPublishItems'));
      const tables = wb.tables.filter((t) => t.sheet === sh);
      if (tables.length) {
        const ids = tables.map((t) => {
          const tp = `xl/tables/table${++nTable}.xml`;
          add(tp, tableXml(t, wb), CT.table);
          return rels.add('table', '../tables/' + tp.split('/').pop());
        });
        out.push(`<tableParts count="${ids.length}">` + ids.map((id) => `<tablePart r:id="${id}"/>`).join('') + '</tableParts>');
      }
      out.push(sparklineXml(sh, wb));
      out.push('</worksheet>');
      add(part, out.join(''), CT.sheet);
      if (rels.list.length) add(relsPath(part), rels.xml());
      sheetEntries.push({ sh, part, kind: 'worksheet' });
    }

    async function drawingPart(sh, rels) {
      const list = sh.drawings.filter((d) => d && d.anchor && (d.kind === 'image' ? wb.media && wb.media.get(d.media) : d.kind === 'chart' ? d.xml || d.chart : d.kind === 'shape'));
      if (!list.length) return null;
      const dPart = `xl/drawings/drawing${++nDrawing}.xml`;
      const drels = new Rels();
      let nextId = 2;
      for (const d of list) if (d.id >= nextId) nextId = d.id + 1;
      const usedIds = new Set();
      let x = HDR + `<xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${NS_A}">`;
      for (const d of list) {
        let id = d.id > 0 && !usedIds.has(d.id) ? d.id : nextId++;
        usedIds.add(id);
        const a = d.anchor;
        const sz = anchorSize(sh, a);
        let open, close = '<xdr:clientData' + (d.unlocked ? ' fLocksWithSheet="0"' : '') + (d.noPrint ? ' fPrintsWithSheet="0"' : '') + '/>';
        if (a.type === 'two' && a.from && a.to) { open = `<xdr:twoCellAnchor${a.editAs && a.editAs !== 'twoCell' ? ` editAs="${a.editAs}"` : ''}>` + ptXml('from', a.from) + ptXml('to', a.to); close += '</xdr:twoCellAnchor>'; }
        else if (a.type === 'one' && a.from) { open = '<xdr:oneCellAnchor>' + ptXml('from', a.from) + `<xdr:ext cx="${emu(a.w)}" cy="${emu(a.h)}"/>`; close += '</xdr:oneCellAnchor>'; }
        else { open = `<xdr:absoluteAnchor><xdr:pos x="${emu(a.x)}" y="${emu(a.y)}"/><xdr:ext cx="${emu(a.w)}" cy="${emu(a.h)}"/>`; close += '</xdr:absoluteAnchor>'; }
        let obj = '';
        const nameAttr = esc(d.name || (d.kind === 'chart' ? 'Chart ' : d.kind === 'image' ? 'Picture ' : 'Shape ') + (id - 1));
        if (d.kind === 'image') {
          const mp = mediaPart(d.media);
          const rid = drels.add('image', '../media/' + mp.split('/').pop());
          const crop = d.crop ? `<a:srcRect${attrs({ l: d.crop.l ? Math.round(d.crop.l * 1000) : undefined, t: d.crop.t ? Math.round(d.crop.t * 1000) : undefined, r: d.crop.r ? Math.round(d.crop.r * 1000) : undefined, b: d.crop.b ? Math.round(d.crop.b * 1000) : undefined })}/>` : '';
          obj = `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="${nameAttr}"${d.descr ? ` descr="${esc(d.descr)}"` : ''}/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
            `<xdr:blipFill><a:blip xmlns:r="${NS_R}" r:embed="${rid}"/>${crop}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
            `<xdr:spPr><a:xfrm${d.rot ? ` rot="${Math.round(d.rot * 60000)}"` : ''}><a:off x="0" y="0"/><a:ext cx="${emu(sz.w)}" cy="${emu(sz.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${d.lineXml || ''}</xdr:spPr></xdr:pic>`;
        } else if (d.kind === 'chart') {
          const cPart = `xl/charts/chart${++nChart}.xml`;
          add(cPart, chartPartXml(d, wb), CT.chart);
          if (d.chartParts && !(d.chart && d.chart.dirty)) {
            const crels = new Rels();
            for (const p of d.chartParts) {
              if (p.type === 'chartStyle') { const n = `xl/charts/style${nChart}.xml`; add(n, p.bytes, CT.chartStyle); crels.add('http://schemas.microsoft.com/office/2011/relationships/chartStyle', 'style' + nChart + '.xml'); }
              else if (p.type === 'chartColorStyle') { const n = `xl/charts/colors${nChart}.xml`; add(n, p.bytes, CT.chartColors); crels.add('http://schemas.microsoft.com/office/2011/relationships/chartColorStyle', 'colors' + nChart + '.xml'); }
            }
            if (crels.list.length) add(relsPath(cPart), crels.xml());
          }
          const rid = drels.add('chart', '../charts/' + cPart.split('/').pop());
          obj = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${nameAttr}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${NS_C}"><c:chart xmlns:c="${NS_C}" xmlns:r="${NS_R}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
        } else {
          if (d.xml && !d.dirty) obj = d.xml.replace(/\sr:(embed|link|id)="[^"]*"/g, '').replace(/(<(?:\w+:)?cNvPr\b[^>]*?\bid=")\d+(")/, '$1' + id + '$2');
          else obj = shapeXml(d, id, sz);
        }
        x += open + obj + close;
      }
      x += '</xdr:wsDr>';
      add(dPart, x, CT.drawing);
      if (drels.list.length) add(relsPath(dPart), drels.xml());
      return rels.add('drawing', '../drawings/' + dPart.split('/').pop());
    }

    /* workbook-level parts */
    const sheetRids = sheetEntries.map((e) => wbRels.add(e.kind === 'chartsheet' ? 'chartsheet' : 'worksheet', e.part.slice(3)));
    wbRels.add('theme', 'theme/theme1.xml');
    add('xl/theme/theme1.xml', wb.themeXml || themeXml(wb), CT.theme);
    wbRels.add('styles', 'styles.xml');
    const extRids = [];
    if (wb.externalLinks && wb.externalLinks.length) {
      wb.externalLinks.forEach((e, i) => {
        const p = `xl/externalLinks/externalLink${i + 1}.xml`;
        add(p, e.xml, CT.ext);
        if (e.relTarget != null) {
          const r = new Rels();
          const t = e.relType === 'xlPathMissing' ? 'http://schemas.microsoft.com/office/2006/relationships/xlExternalLinkPath/xlPathMissing' : e.relType || 'externalLinkPath';
          r.add(t, e.relTarget, e.relExternal !== false);
          add(relsPath(p), r.xml());
        }
        extRids.push(wbRels.add('externalLink', 'externalLinks/externalLink' + (i + 1) + '.xml'));
      });
    }
    if (macro) { add('xl/vbaProject.bin', wb.extra.vba, CT.vba); wbRels.add('http://schemas.microsoft.com/office/2006/relationships/vbaProject', 'vbaProject.bin'); }
    /* workbook.xml */
    let x = HDR + `<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`;
    x += '<fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/>';
    x += `<workbookPr${attrs({ date1904: wb.date1904 ? '1' : undefined, codeName: wb.codeName, filterPrivacy: wb.extra && (wb.extra.filterPrivacy || wb.extra.removePersonal) ? '1' : undefined, defaultThemeVersion: '124226' })}/>`;
    if (wb.protection) { const a = Object.assign({}, wb.protection.attrs || {}); if (wb.protection.structure) a.lockStructure = '1'; if (wb.protection.windows) a.lockWindows = '1'; x += '<workbookProtection' + Object.keys(a).map((k) => ` ${k}="${esc(a[k])}"`).join('') + '/>'; }
    const visible = wb.sheets.map((s, i) => (s.state === 'visible' ? i : -1)).filter((i) => i >= 0);
    let active = wb.sheets[wb.active] && wb.sheets[wb.active].state === 'visible' ? wb.active : visible[0] || 0;
    x += `<bookViews><workbookView${attrs({ xWindow: 240, yWindow: 120, windowWidth: 18195, windowHeight: 8505, tabRatio: wb.tabRatio || undefined, firstSheet: wb.firstTab && wb.firstTab < wb.sheets.length ? wb.firstTab : undefined, activeTab: active || undefined, showSheetTabs: wb.hideTabs ? '0' : undefined, showHorizontalScroll: wb.hideHScroll ? '0' : undefined, showVerticalScroll: wb.hideVScroll ? '0' : undefined })}/></bookViews>`;
    const usedIds = new Set();
    x += '<sheets>' + sheetEntries.map((e, i) => {
      let id = e.sh.sheetId > 0 && !usedIds.has(e.sh.sheetId) ? e.sh.sheetId : 0;
      if (!id) { id = 1; while (usedIds.has(id) || wb.sheets.some((s) => s !== e.sh && s.sheetId === id)) id++; e.sh.sheetId = id; }
      usedIds.add(id);
      return `<sheet${attrs({ name: e.sh.name, sheetId: id, state: e.sh.state !== 'visible' ? e.sh.state : undefined, 'r:id': sheetRids[i] })}/>`;
    }).join('') + '</sheets>';
    if (extRids.length) x += '<externalReferences>' + extRids.map((id) => `<externalReference r:id="${id}"/>`).join('') + '</externalReferences>';
    /* defined names: the AutoFilter database name Excel keeps for filtered sheets */
    const names = wb.names.filter((n) => n && n.name && n.ref != null).slice();
    wb.sheets.forEach((s, i) => {
      if (s.autoFilter && s.autoFilter.ref && !names.some((n) => n.scope === i && /^_xlnm\._FilterDatabase$/i.test(n.name))) names.push({ name: '_xlnm._FilterDatabase', ref: F.quoteSheet(s.name) + '!' + F.absRangeName(s.autoFilter.ref), scope: i, hidden: true });
    });
    const okNames = names.filter((n) => n.scope == null || (n.scope >= 0 && n.scope < wb.sheets.length));
    if (okNames.length) {
      x += '<definedNames>' + okNames.map((n) => `<definedName${attrs({ name: n.name, comment: n.comment, localSheetId: n.scope != null ? n.scope : undefined, hidden: n.hidden ? '1' : undefined })}>${escT(String(n.ref).replace(/^=/, ''))}</definedName>`).join('') + '</definedNames>';
    }
    const cp = wb.calcPr || {};
    x += `<calcPr${attrs({ calcId: '124519', calcMode: cp.mode && cp.mode !== 'auto' ? cp.mode : undefined, iterate: cp.iterate ? '1' : undefined, iterateCount: cp.iterate && cp.iterateCount !== 100 ? cp.iterateCount : undefined, iterateDelta: cp.iterate && cp.iterateDelta !== 0.001 ? cp.iterateDelta : undefined, fullPrecision: cp.fullPrecision === false ? '0' : undefined, refMode: wb.r1c1 ? 'R1C1' : undefined, fullCalcOnLoad: '1' })}/>`;
    x += '</workbook>';
    const wbCT = type === 'xltx' ? CT.wbt : type === 'xltm' ? CT.wbtm : macro ? CT.wbm : CT.wb;
    add('xl/workbook.xml', x, wbCT);
    add('xl/styles.xml', buildStyles(wb), CT.styles);
    if (sst.list.length) { wbRels.add('sharedStrings', 'sharedStrings.xml'); add('xl/sharedStrings.xml', sst.xml(), CT.sst); }
    if (ctx.dynamic) { wbRels.add('sheetMetadata', 'metadata.xml'); add('xl/metadata.xml', META_XML, CT.meta); }
    add('xl/_rels/workbook.xml.rels', wbRels.xml());
    /* document properties */
    /* Tools ▸ Options ▸ Security ▸ Remove personal information: no author, last author or manager */
    const p = wb.extra && (wb.extra.removePersonal || wb.extra.filterPrivacy) ? Object.assign({}, wb.props, { creator: '', lastModifiedBy: ' ', manager: '' }) : wb.props || {};
    const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
    const iso = (s) => { if (!s) return null; const d = new Date(s); return isNaN(d) ? null : d.toISOString().replace(/\.\d+Z$/, 'Z'); };
    const dc = (tag, v) => (v ? `<${tag}>${escT(v)}</${tag}>` : '');
    add('docProps/core.xml', HDR + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      dc('dc:title', p.title) + dc('dc:subject', p.subject) + dc('dc:creator', p.creator) + dc('cp:keywords', p.keywords) + dc('dc:description', p.description) + dc('cp:lastModifiedBy', p.lastModifiedBy || (L.app && L.app.opts && L.app.opts.userName) || p.creator) + dc('cp:category', p.category) +
      `<dcterms:created xsi:type="dcterms:W3CDTF">${iso(p.created) || now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`, CT.core);
    const titles = wb.sheets.map((s) => s.name);
    add('docProps/app.xml', HDR + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>' +
      `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${titles.length}</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
      `<TitlesOfParts><vt:vector size="${titles.length}" baseType="lpstr">${titles.map((t) => `<vt:lpstr>${escT(t)}</vt:lpstr>`).join('')}</vt:vector></TitlesOfParts>` +
      dc('Manager', p.manager) + dc('Company', p.company) + '<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>11.0000</AppVersion></Properties>', CT.app);
    const rootRels = new Rels();
    rootRels.add('officeDocument', 'xl/workbook.xml');
    rootRels.add('http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', 'docProps/core.xml');
    rootRels.add('extended-properties', 'docProps/app.xml');
    if (p.custom && p.custom.length) {
      add('docProps/custom.xml', HDR + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
        p.custom.map((c, i) => `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="${i + 2}" name="${esc(c.name)}"><vt:${c.type || 'lpwstr'}>${escT(c.value)}</vt:${c.type || 'lpwstr'}></property>`).join('') + '</Properties>', CT.custom);
      rootRels.add('custom-properties', 'docProps/custom.xml');
    }
    add('_rels/.rels', rootRels.xml());
    /* content types first in the archive */
    let ct = HDR + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
    for (const [e, t] of defaults) ct += `<Default Extension="${e}" ContentType="${t}"/>`;
    for (const [n, t] of overrides) ct += `<Override PartName="${esc(n)}" ContentType="${t}"/>`;
    ct += '</Types>';
    files.unshift({ name: '[Content_Types].xml', data: ct });
    const blob = await L.zip.write(files, 'application/octet-stream');
    let bytes = new Uint8Array(await blob.arrayBuffer());
    if (opts.password) bytes = await L.officeCrypto.encrypt(bytes, opts.password);
    return bytes;
  };
  W.MIME = { xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.12', xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template', xltm: 'application/vnd.ms-excel.template.macroEnabled.12' };
  W.write = async function (wb, opts) {
    const bytes = await W.bytes(wb, opts);
    return new Blob([bytes], { type: W.MIME[(opts && opts.type) || 'xlsx'] || W.MIME.xlsx });
  };
  W.themeXml = themeXml;
})(typeof window !== 'undefined' ? window : globalThis);
