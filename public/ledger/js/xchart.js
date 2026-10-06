/* Ledger — worksheet charts: read c:chartSpace, live data from cell references, SVG rendering,
 * chart XML for new or edited charts, reference maintenance, and the Chart Wizard. */
(function (root) {
  'use strict';
  const L = root.L;
  const M = L.model, F = L.formula, C = L.calc;
  const XC = (L.xchart = {});
  const PAL2003 = ['#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF', '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF'];
  XC.PAL2003 = PAL2003;
  const themeCtx = (wb) => {
    const t = wb && wb.theme;
    const names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
    const colors = {};
    names.forEach((n, i) => { colors[n] = (t && t.colors && t.colors[i]) || M.DEFAULT_THEME.colors[i]; });
    return { theme: { colors } };
  };

  /* ------------------------------------------------------------ reading */
  const kids = (el, n) => (el ? el.children.filter((c) => !n || c.localName === n) : []);
  const kid = (el, n) => (el ? el.children.find((c) => c.localName === n) || null : null);
  const fOf = (el) => { if (!el) return null; const all = el.getElementsByTagName('*'); const f = all.find((e) => e.localName === 'f'); return f ? f.textContent.trim() : null; };
  XC.read = function (doc, wb, sh) {
    const model = L.dml && L.dml.chart ? L.dml.chart(doc, themeCtx(wb)) : null;
    if (!model) return null;
    /* every formula in document order, so edited references can be written back in place */
    model.refs = doc.getElementsByTagName('*').filter((e) => e.localName === 'f').map((e) => e.textContent.trim());
    const plot = kid(kid(doc, 'chart'), 'plotArea');
    const types = /^(bar|bar3D|line|line3D|pie|pie3D|ofPie|doughnut|area|area3D|scatter|radar|bubble|stock|surface|surface3D)Chart$/;
    const sers = [];
    for (const ct of kids(plot).filter((c) => types.test(c.localName))) for (const s of kids(ct, 'ser')) sers.push(s);
    sers.forEach((s, i) => {
      const sr = model.series[i];
      if (!sr) return;
      sr.nameRef = fOf(kid(s, 'tx'));
      sr.catRef = fOf(kid(s, 'cat') || kid(s, 'xVal'));
      sr.valRef = fOf(kid(s, 'val') || kid(s, 'yVal'));
      sr.cacheVals = sr.vals.slice();
    });
    model.cacheCats = (model.cats || []).slice();
    model.imported = true;
    const v3 = kid(kid(doc, 'chart'), 'view3D');
    if (v3) model.view3D = true;
    const spPr = kid(plot, 'spPr');
    const pf = spPr && kid(spPr, 'solidFill');
    if (pf) { const c = L.dml.colorOf(pf, themeCtx(wb)); if (c) model.plotBg = { t: 'solid', color: c.c }; }
    const csp = kid(doc, 'spPr');
    const cf = csp && kid(csp, 'solidFill');
    if (cf) { const c = L.dml.colorOf(cf, themeCtx(wb)); if (c) model.bg = { t: 'solid', color: c.c }; }
    if (csp && kid(csp, 'noFill')) model.bg = { t: 'none' };
    void sh;
    return model;
  };

  /* ------------------------------------------------------------ live values */
  /** evaluate a chart reference ("Sheet1!$B$2:$B$5") → {vals: [...], texts: [...]} or null */
  function evalRef(text, wb, sh) {
    if (!text) return null;
    let ast;
    try { ast = F.parse(text.replace(/^\(|\)$/g, '')); } catch (e) { return null; }
    let v;
    try { v = C.ev(ast, C.ctx(wb, sh, 0, 0, true)); } catch (e) { return null; }
    const areas = v instanceof C.RefList ? v.list : v instanceof C.Ref ? [v] : null;
    if (!areas) return null;
    const vals = [], texts = [];
    for (const a of areas) {
      if (!a.sheet || a.sheet.external) return null;
      const byCol = a.r2 - a.r1 >= a.c2 - a.c1;
      const n = byCol ? a.r2 - a.r1 + 1 : a.c2 - a.c1 + 1;
      if (n > 4000) return null;
      for (let i = 0; i < n; i++) {
        const r = byCol ? a.r1 + i : a.r1, c = byCol ? a.c1 : a.c1 + i;
        /* hidden rows and columns are left out, as Excel's default "plot visible cells only" */
        const row = a.sheet.rows[r], col = a.sheet.cols[c];
        if ((row && row.hidden) || (col && col.hidden)) continue;
        const val = C.cellValue(a.sheet, r, c);
        vals.push(val);
        /* multi-column category ranges join the levels */
        if (!byCol || a.c1 === a.c2) texts.push(L.csv ? L.csv.cellText(a.sheet, r, c) : val == null ? '' : String(val));
        else { const parts = []; for (let cc = a.c1; cc <= a.c2; cc++) parts.push(L.csv ? L.csv.cellText(a.sheet, r, cc) : String(C.cellValue(a.sheet, r, cc) || '')); texts.push(parts.filter(Boolean).join(' ')); }
      }
    }
    return { vals, texts };
  }
  XC.evalRef = evalRef;
  /** the model to draw: cached values replaced by the cells' current values */
  XC.resolve = function (chart, wb, sh) {
    const m = Object.assign({}, chart);
    let cats = null;
    m.series = (chart.series || []).map((s, i) => {
      const o = Object.assign({}, s);
      const v = evalRef(s.valRef, wb, sh);
      if (v) o.vals = v.vals.map((x) => (typeof x === 'number' ? x : typeof x === 'boolean' ? +x : x == null || x === '' ? null : 0)).map((x) => (x == null ? 0 : x));
      const nm = evalRef(s.nameRef, wb, sh);
      if (nm && nm.texts.length) o.name = nm.texts.join(' ');
      if (!cats && s.catRef) { const c = evalRef(s.catRef, wb, sh); if (c) cats = chart.kind === 'scatter' ? c.vals.map((x) => (typeof x === 'number' ? x : 0)) : c.texts; }
      if (!o.color) o.color = (chart.imported ? null : PAL2003[i % 8]) || L.chart.PALETTE[i % L.chart.PALETTE.length];
      return o;
    });
    if (cats) m.cats = cats;
    else if (!chart.series.some((s) => s.catRef)) { const n = Math.max(0, ...m.series.map((s) => s.vals.length)); if (!chart.cats || chart.cats.length !== n) m.cats = Array.from({ length: n }, (_, i) => String(i + 1)); }
    return m;
  };
  /** an <svg> for chart drawing d at w × h CSS pixels */
  XC.render = function (chart, wb, sh, w, h, z) {
    if (L.render && !L.render.fillCSS) L.render.fillCSS = (f) => (f && f.t === 'solid' ? f.color || f.c || '#FFFFFF' : 'none');
    const m = XC.resolve(chart, wb, sh);
    m.fsz = Math.max(6, (chart.fsz || 10) * (96 / 72) * (z || 1));
    if (!m.bg) m.bg = { t: 'solid', color: '#FFFFFF' };
    if (!chart.imported && !m.plotBg) m.plotBg = { t: 'solid', color: '#C0C0C0' };
    return L.chart.render(m, w, h, null);
  };

  /* ------------------------------------------------------------ reference maintenance */
  function mapRefs(chart, fn) {
    const memo = new Map();
    const conv = (t) => {
      if (t == null) return t;
      if (memo.has(t)) return memo.get(t);
      let out = t;
      try { const ast = F.parse(t.replace(/^\(|\)$/g, '')); const r = fn(ast); if (r && r.changed) out = (/^\(/.test(t) ? '(' : '') + F.toText(r.ast) + (/^\(/.test(t) ? ')' : ''); } catch (e) { out = t; }
      memo.set(t, out);
      return out;
    };
    if (chart.refs) chart.refs = chart.refs.map(conv);
    for (const s of chart.series || []) { s.nameRef = conv(s.nameRef); s.catRef = conv(s.catRef); s.valRef = conv(s.valRef); }
  }
  XC.adjust = (chart, fn) => mapRefs(chart, fn);
  XC.rename = (chart, from, to) => mapRefs(chart, (ast) => F.renameSheet(ast, from, to));
  XC.patchRefs = function (xml, refs) {
    let i = 0;
    return xml.replace(/<((?:\w+:)?f)>([\s\S]*?)<\/\1>/g, (m, tag, body) => { const r = refs[i++]; return r == null ? m : `<${tag}>${L.xml.esc(r)}</${tag}>`; });
  };

  /* ------------------------------------------------------------ writing */
  const esc = (s) => L.xml.esc(String(s));
  const hex = (c) => String(c || '#000000').replace('#', '').toUpperCase();
  function cacheXml(kind, vals, fmt) {
    if (kind === 'num') return `<c:numCache><c:formatCode>${esc(fmt || 'General')}</c:formatCode><c:ptCount val="${vals.length}"/>` + vals.map((v, i) => (typeof v === 'number' && isFinite(v) ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('') + '</c:numCache>';
    return `<c:strCache><c:ptCount val="${vals.length}"/>` + vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v == null ? '' : v)}</c:v></c:pt>`).join('') + '</c:strCache>';
  }
  const richTitle = (text, sz, bold) => `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="${sz}" b="${bold ? 1 : 0}"/></a:pPr><a:r><a:rPr lang="en-US" sz="${sz}" b="${bold ? 1 : 0}"/><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
  const richTitleV = (text) => `<c:title><c:tx><c:rich><a:bodyPr rot="-5400000" vert="horz"/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="1"/></a:pPr><a:r><a:rPr lang="en-US" sz="1000" b="1"/><a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>`;
  XC.toXml = function (chart, wb, d) {
    const sh = d && d.sheet ? d.sheet : wb.sheets[wb.active];
    const m = XC.resolve(chart, wb, sh);
    const k = chart.kind || 'col';
    const isPie = k === 'pie' || k === 'doughnut', isScatter = k === 'scatter' || k === 'scatterLines';
    const barDir = /^bar/.test(k) ? 'bar' : 'col';
    const grouping = /Pct$/.test(k) ? 'percentStacked' : /Stacked$/.test(k) ? 'stacked' : /^(line|area)/.test(k) ? 'standard' : 'clustered';
    let ser = '';
    m.series.forEach((s, i) => {
      const col = hex(s.color || PAL2003[i % 8]);
      const name = s.nameRef ? `<c:tx><c:strRef><c:f>${esc(s.nameRef)}</c:f>${cacheXml('str', [s.name])}</c:strRef></c:tx>` : `<c:tx><c:v>${esc(s.name || 'Series' + (i + 1))}</c:v></c:tx>`;
      const isLine = /^line/.test(k) || isScatter;
      const sp = k === 'scatter' ? '<c:spPr><a:ln w="28575"><a:noFill/></a:ln></c:spPr>' : isLine ? `<c:spPr><a:ln w="28575"><a:solidFill><a:srgbClr val="${col}"/></a:solidFill></a:ln></c:spPr>` : `<c:spPr><a:solidFill><a:srgbClr val="${col}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>`;
      let marker = '';
      if (k === 'line' || k === 'lineStacked') marker = '<c:marker><c:symbol val="none"/></c:marker>';
      else if (/^line|scatter/.test(k)) marker = `<c:marker><c:symbol val="square"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${col}"/></a:solidFill></c:spPr></c:marker>`;
      let dpts = '';
      if (isPie) dpts = (m.cats || []).map((_, j) => `<c:dPt><c:idx val="${j}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${hex((chart.pieColors && chart.pieColors[j]) || PAL2003[j % PAL2003.length])}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('');
      const cats = m.cats || [];
      const catEl = isScatter ? 'c:xVal' : 'c:cat';
      const valEl = isScatter ? 'c:yVal' : 'c:val';
      const catNum = isScatter || cats.every((x) => typeof x === 'number');
      const catXml = s.catRef ? `<${catEl}>${catNum ? `<c:numRef><c:f>${esc(s.catRef)}</c:f>${cacheXml('num', cats.map(Number))}</c:numRef>` : `<c:strRef><c:f>${esc(s.catRef)}</c:f>${cacheXml('str', cats)}</c:strRef>`}</${catEl}>` : '';
      const valXml = `<${valEl}>${s.valRef ? `<c:numRef><c:f>${esc(s.valRef)}</c:f>${cacheXml('num', s.vals)}</c:numRef>` : `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${s.vals.length}"/>${s.vals.map((v, j) => `<c:pt idx="${j}"><c:v>${v}</c:v></c:pt>`).join('')}</c:numLit>`}</${valEl}>`;
      const labels = chart.labels ? '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="' + (isPie ? 0 : 1) + '"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="' + (isPie ? 1 : 0) + '"/><c:showBubbleSize val="0"/></c:dLbls>' : '';
      ser += `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${name}${sp}${isLine ? marker : ''}${dpts}${labels}${catXml}${valXml}${isLine && !isScatter ? '<c:smooth val="0"/>' : isScatter ? '<c:smooth val="0"/>' : ''}</c:ser>`;
    });
    const AX1 = 50010001, AX2 = 50010002;
    let plotType;
    if (isPie) plotType = k === 'pie' ? `<c:pieChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/></c:pieChart>` : `<c:doughnutChart><c:varyColors val="1"/>${ser}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`;
    else if (isScatter) plotType = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${ser}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:scatterChart>`;
    else if (/^line/.test(k)) plotType = `<c:lineChart><c:grouping val="${k === 'lineStacked' ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${ser}<c:marker val="1"/><c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:lineChart>`;
    else if (/^area/.test(k)) plotType = `<c:areaChart><c:grouping val="${k === 'areaStacked' ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${ser}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:areaChart>`;
    else plotType = `<c:barChart><c:barDir val="${barDir}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${ser}<c:gapWidth val="150"/>${grouping !== 'clustered' ? '<c:overlap val="100"/>' : ''}<c:axId val="${AX1}"/><c:axId val="${AX2}"/></c:barChart>`;
    const grid = chart.gridY !== false ? '<c:majorGridlines><c:spPr><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' : '';
    const axLine = '<c:spPr><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>';
    const horizBar = barDir === 'bar' && !isScatter && !/^(line|area)/.test(k);
    let axes = '';
    if (!isPie) {
      const catPos = horizBar ? 'l' : 'b', valPos = horizBar ? 'b' : 'l';
      const xTitle = chart.axTitleX ? (horizBar ? richTitleV(chart.axTitleX) : richTitle(chart.axTitleX, 1000, true)) : '';
      const yTitle = chart.axTitleY ? (horizBar ? richTitle(chart.axTitleY, 1000, true) : richTitleV(chart.axTitleY)) : '';
      if (isScatter) {
        axes = `<c:valAx><c:axId val="${AX1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${xTitle}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axLine}<c:crossAx val="${AX2}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>` +
          `<c:valAx><c:axId val="${AX2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="l"/>${grid}${yTitle}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axLine}<c:crossAx val="${AX1}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`;
      } else {
        axes = `<c:catAx><c:axId val="${AX1}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/>${xTitle}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axLine}<c:crossAx val="${AX2}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>` +
          `<c:valAx><c:axId val="${AX2}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/>${grid}${yTitle}<c:numFmt formatCode="${/Pct$/.test(k) ? '0%' : 'General'}" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axLine}<c:crossAx val="${AX1}"/><c:crosses val="autoZero"/><c:crossBetween val="${/^area/.test(k) ? 'midCat' : 'between'}"/></c:valAx>`;
      }
    }
    const plotFill = chart.plotBg && chart.plotBg.t === 'solid' ? `<c:spPr><a:solidFill><a:srgbClr val="${hex(chart.plotBg.color)}"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr>` : !chart.imported && !isPie ? '<c:spPr><a:solidFill><a:srgbClr val="C0C0C0"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="808080"/></a:solidFill></a:ln></c:spPr>' : '';
    const legend = chart.legend && chart.legend !== 'none' ? `<c:legend><c:legendPos val="${chart.legend}"/><c:overlay val="0"/><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr></c:legend>` : '';
    const title = chart.title ? richTitle(chart.title, 1200, true) : '';
    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<c:roundedCorners val="0"/><c:chart>' + title + `<c:autoTitleDeleted val="${title ? 0 : 1}"/>` + `<c:plotArea><c:layout/>${plotType}${axes}${plotFill}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
      `<c:spPr><a:solidFill><a:srgbClr val="${hex((chart.bg && chart.bg.color) || '#FFFFFF')}"/></a:solidFill><a:ln w="3175"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></c:spPr>` +
      '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000" b="0" i="0"><a:latin typeface="Arial"/><a:cs typeface="Arial"/></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr>' +
      '<c:printSettings><c:headerFooter/><c:pageMargins b="1" l="0.75" r="0.75" t="1" header="0.5" footer="0.5"/><c:pageSetup/></c:printSettings></c:chartSpace>';
  };

  /* ------------------------------------------------------------ building a chart from a range */
  /** series from a data range: header row / column detection like the Chart Wizard */
  XC.fromRange = function (sh, rg, byRows) {
    const wb = sh.wb;
    const q = (r1, c1, r2, c2) => F.quoteSheet(sh.name) + '!' + F.absRangeName({ r1, c1, r2, c2 });
    const isNum = (r, c) => typeof C.cellValue(sh, r, c) === 'number';
    const blank = (r, c) => { const v = C.cellValue(sh, r, c); return v == null || v === ''; };
    /* the top-left cell blank or the first row / column text → labels */
    let hasHeadRow = false, hasHeadCol = false;
    if (rg.r2 > rg.r1) { let txt = 0; for (let c = rg.c1; c <= rg.c2; c++) if (!isNum(rg.r1, c) && !blank(rg.r1, c)) txt++; hasHeadRow = txt > 0 && (blank(rg.r1, rg.c1) || txt >= (rg.c2 - rg.c1 + 1) / 2); }
    if (rg.c2 > rg.c1) { let txt = 0; for (let r = rg.r1 + (hasHeadRow ? 1 : 0); r <= rg.r2; r++) if (!isNum(r, rg.c1) && !blank(r, rg.c1)) txt++; hasHeadCol = txt > 0 || (hasHeadRow && blank(rg.r1, rg.c1)); }
    const r0 = rg.r1 + (hasHeadRow ? 1 : 0), c0 = rg.c1 + (hasHeadCol ? 1 : 0);
    if (byRows == null) byRows = rg.c2 - c0 > rg.r2 - r0; /* Excel: series in the longer direction's perpendicular */
    const series = [];
    if (!byRows) {
      for (let c = c0; c <= rg.c2; c++) series.push({ nameRef: hasHeadRow ? q(rg.r1, c, rg.r1, c) : null, name: hasHeadRow ? String(C.cellValue(sh, rg.r1, c) == null ? '' : C.cellValue(sh, rg.r1, c)) : 'Series' + (c - c0 + 1), catRef: hasHeadCol ? q(r0, rg.c1, rg.r2, rg.c1) : null, valRef: q(r0, c, rg.r2, c), vals: [] });
    } else {
      for (let r = r0; r <= rg.r2; r++) series.push({ nameRef: hasHeadCol ? q(r, rg.c1, r, rg.c1) : null, name: hasHeadCol ? String(C.cellValue(sh, r, rg.c1) == null ? '' : C.cellValue(sh, r, rg.c1)) : 'Series' + (r - r0 + 1), catRef: hasHeadRow ? q(rg.r1, c0, rg.r1, rg.c2) : null, valRef: q(r, c0, r, rg.c2), vals: [] });
    }
    series.forEach((s, i) => { s.color = PAL2003[i % 8]; });
    void wb;
    return { series, byRows };
  };

  /* ------------------------------------------------------------ Chart Wizard */
  const TYPES = [
    { id: 'column', label: 'Column', subs: [['col', 'Clustered Column'], ['colStacked', 'Stacked Column'], ['colPct', '100% Stacked Column']] },
    { id: 'bar', label: 'Bar', subs: [['bar', 'Clustered Bar'], ['barStacked', 'Stacked Bar']] },
    { id: 'line', label: 'Line', subs: [['line', 'Line'], ['lineMarkers', 'Line with markers displayed at each data value']] },
    { id: 'pie', label: 'Pie', subs: [['pie', 'Pie']] },
    { id: 'scatter', label: 'XY (Scatter)', subs: [['scatter', 'Scatter. Compares pairs of values.']] },
    { id: 'area', label: 'Area', subs: [['area', 'Area'], ['areaStacked', 'Stacked Area']] },
    { id: 'doughnut', label: 'Doughnut', subs: [['doughnut', 'Doughnut']] },
  ];
  const typeIcon = (kind) => {
    const s = 40;
    const bars = (pts, horiz) => pts.map(([x, y, w, hh, col]) => (horiz ? `<rect x="${y}" y="${x}" width="${hh}" height="${w}" fill="${col}" stroke="#000" stroke-width=".5"/>` : `<rect x="${x}" y="${s - 4 - hh - y}" width="${w}" height="${hh}" fill="${col}" stroke="#000" stroke-width=".5"/>`)).join('');
    const c1 = '#9999FF', c2 = '#993366', c3 = '#FFFFCC';
    let body = '';
    switch (kind) {
      case 'col': body = bars([[6, 0, 5, 14, c1], [11, 0, 5, 22, c2], [20, 0, 5, 18, c1], [25, 0, 5, 28, c2]]); break;
      case 'colStacked': body = bars([[8, 0, 8, 12, c1], [8, 12, 8, 8, c2], [22, 0, 8, 18, c1], [22, 18, 8, 10, c2]]); break;
      case 'colPct': body = bars([[8, 0, 8, 18, c1], [8, 18, 8, 14, c2], [22, 0, 8, 10, c1], [22, 10, 8, 22, c2]]); break;
      case 'bar': body = bars([[6, 4, 5, 14, c1], [11, 4, 5, 22, c2], [20, 4, 5, 18, c1], [25, 4, 5, 28, c2]], true); break;
      case 'barStacked': body = bars([[8, 4, 8, 12, c1], [8, 16, 8, 8, c2], [22, 4, 8, 18, c1], [22, 22, 8, 10, c2]], true); break;
      case 'line': case 'lineMarkers': body = `<polyline points="4,30 14,20 24,24 36,8" fill="none" stroke="#000080" stroke-width="1.5"/><polyline points="4,34 14,28 24,30 36,22" fill="none" stroke="#FF00FF" stroke-width="1.5"/>` + (kind === 'lineMarkers' ? [[4, 30], [14, 20], [24, 24], [36, 8]].map(([x, y]) => `<rect x="${x - 2}" y="${y - 2}" width="4" height="4" fill="#000080"/>`).join('') : ''); break;
      case 'pie': body = `<circle cx="20" cy="20" r="14" fill="${c1}" stroke="#000" stroke-width=".5"/><path d="M20,20 L20,6 A14,14 0 0 1 33,25 Z" fill="${c2}" stroke="#000" stroke-width=".5"/><path d="M20,20 L33,25 A14,14 0 0 1 14,33 Z" fill="${c3}" stroke="#000" stroke-width=".5"/>`; break;
      case 'doughnut': body = `<circle cx="20" cy="20" r="14" fill="${c1}" stroke="#000" stroke-width=".5"/><path d="M20,20 L20,6 A14,14 0 0 1 33,25 Z" fill="${c2}" stroke="#000" stroke-width=".5"/><circle cx="20" cy="20" r="6" fill="#fff" stroke="#000" stroke-width=".5"/>`; break;
      case 'scatter': body = [[6, 30], [10, 24], [16, 26], [22, 14], [28, 16], [34, 8]].map(([x, y]) => `<rect x="${x - 2}" y="${y - 2}" width="4" height="4" fill="#000080"/>`).join(''); break;
      case 'area': case 'areaStacked': body = `<polygon points="4,36 4,20 16,12 26,18 36,6 36,36" fill="${c1}" stroke="#000" stroke-width=".5"/><polygon points="4,36 4,28 16,24 26,28 36,20 36,36" fill="${c2}" stroke="#000" stroke-width=".5"/>`; break;
      default: break;
    }
    return `<svg width="${s}" height="${s}" viewBox="0 0 ${s} ${s}"><rect x="0" y="0" width="${s}" height="${s}" fill="#fff"/>${body}</svg>`;
  };
  const CW = (L.chartWizard = {});
  CW.TYPES = TYPES;
  CW.typeIcon = typeIcon;
  CW.typeMenu = function (at) {
    const d = L.grid.selectedObject();
    const kinds = TYPES.flatMap((t) => t.subs.map((s) => s[0]));
    L.ui.openMenu([{ custom: (close) => {
      const g = L.h('div', { class: 'ctype-grid' });
      for (const k of kinds) { const b = L.h('button', { type: 'button', class: 'ctype-b', html: typeIcon(k), 'data-tip': k, 'aria-label': k }); b.addEventListener('click', () => { close(); if (d && d.kind === 'chart' && d.chart) CW.setKind(d, k); }); g.appendChild(b); }
      return g;
    } }], at);
  };
  CW.setKind = function (d, kind) {
    const sh = L.grid.sheet();
    const nc = Object.assign({}, d.chart, { kind, dirty: true });
    const nd = Object.assign({}, d, { chart: nc });
    L.ops.tx(sh.wb, 'Chart Type', () => L.ops.setDrawings(sh, sh.drawings.map((x) => (x === d ? nd : x))));
    L.grid.selectObject(nd); L.grid.syncObjects(true);
  };
  /**
   * The four-step Chart Wizard. opts: {sheet: true to start with "As new sheet", edit: drawing to edit}
   */
  CW.open = function (opts) {
    opts = opts || {};
    const ui = L.ui, h = L.h, G = L.grid;
    const sh = G.sheet(), wb = G.wb;
    const editing = opts.edit || null;
    let rg = O_clip(sh, G.range());
    const single = rg.r1 === rg.r2 && rg.c1 === rg.c2;
    if (single) rg = L.ops.currentRegion(sh, rg.r1, rg.c1);
    const st = {
      kind: editing ? editing.chart.kind : 'col', step: 1,
      range: editing ? null : F.quoteSheet(sh.name) + '!' + F.absRangeName(rg), byRows: null,
      title: editing ? editing.chart.title || '' : '', axX: editing ? editing.chart.axTitleX || '' : '', axY: editing ? editing.chart.axTitleY || '' : '',
      gridY: editing ? editing.chart.gridY !== false : true, legend: editing ? editing.chart.legend || 'none' : 'r', labels: editing ? !!editing.chart.labels : false,
      where: opts.sheet ? 'sheet' : 'object', sheetName: wb.nextSheetName('Chart'),
      series: editing ? editing.chart.series.map((s) => Object.assign({}, s)) : null,
    };
    const body = h('div', { class: 'cw' });
    const stepBox = h('div', { class: 'cw-step' });
    const preview = h('div', { class: 'cw-preview' });
    body.append(stepBox, preview);
    const model = () => {
      let series = st.series;
      if (!editing || st.rangeChanged) {
        const p = parseRange(st.range);
        if (p) { const fr = XC.fromRange(p.sheet, p.rg, st.byRows); series = fr.series; if (st.byRows == null) st.byRows = fr.byRows; }
        else series = [];
      }
      return { kind: st.kind, title: st.title, axTitleX: st.axX, axTitleY: st.axY, gridY: st.gridY, legend: st.legend, labels: st.labels, series: series || [], cats: [], fsz: 10 };
    };
    function parseRange(t) {
      if (!t) return null;
      try {
        const ast = F.parse(String(t).replace(/^=/, ''));
        const v = C.ev(ast, C.ctx(wb, sh, 0, 0));
        const a = v instanceof C.RefList ? v.list[0] : v;
        if (!(a instanceof C.Ref)) return null;
        return { sheet: a.sheet, rg: { r1: a.r1, c1: a.c1, r2: Math.min(a.r2, a.sheet.maxR), c2: Math.min(a.c2, a.sheet.maxC) } };
      } catch (e) { return null; }
    }
    function drawPreview() {
      preview.textContent = '';
      const m = model();
      if (!m.series.length) { preview.appendChild(h('div', { class: 'cw-empty', text: 'Select the data range for the chart.' })); return; }
      try { preview.appendChild(XC.render(m, wb, sh, 300, 200, 0.75)); } catch (e) { preview.appendChild(h('div', { text: String(e.message || e) })); }
    }
    function step1() {
      stepBox.textContent = '';
      const list = h('div', { class: 'cw-types', role: 'listbox' });
      const subs = h('div', { class: 'cw-subs' });
      const desc = h('div', { class: 'cw-desc' });
      let curType = TYPES.find((t) => t.subs.some((s) => s[0] === st.kind)) || TYPES[0];
      const drawSubs = () => {
        subs.textContent = '';
        for (const [k, label] of curType.subs) {
          const b = h('button', { type: 'button', class: 'cw-sub' + (k === st.kind ? ' on' : ''), html: typeIcon(k), 'aria-label': label });
          b.addEventListener('click', () => { st.kind = k; desc.textContent = label; drawSubs(); drawPreview(); });
          subs.appendChild(b);
        }
        const cur = curType.subs.find((s) => s[0] === st.kind);
        desc.textContent = cur ? cur[1] : '';
      };
      for (const t of TYPES) {
        const it = h('div', { class: 'cw-type' + (t === curType ? ' on' : ''), role: 'option', html: typeIcon(t.subs[0][0]).replace(/width="40" height="40"/, 'width="20" height="20"') + `<span>${t.label}</span>` });
        it.addEventListener('click', () => { curType = t; if (!t.subs.some((s) => s[0] === st.kind)) st.kind = t.subs[0][0]; L.$$('.cw-type', list).forEach((x) => x.classList.toggle('on', x === it)); drawSubs(); drawPreview(); });
        list.appendChild(it);
      }
      stepBox.append(h('div', { class: 'cw-cols' }, h('div', { class: 'col' }, h('label', { text: 'Chart type:' }), list), h('div', { class: 'col' }, h('label', { text: 'Chart sub-type:' }), subs, desc)));
      drawSubs();
    }
    function step2() {
      stepBox.textContent = '';
      const inp = h('input', { type: 'text', value: st.range || '', style: 'width:100%' });
      inp.addEventListener('change', () => { st.range = inp.value; st.rangeChanged = true; st.byRows = null; drawPreview(); });
      const rows = ui.radio('cw-by', '&Rows', st.byRows === true, () => { st.byRows = true; st.rangeChanged = true; drawPreview(); });
      const cols = ui.radio('cw-by', 'Colu&mns', st.byRows !== true, () => { st.byRows = false; st.rangeChanged = true; drawPreview(); });
      stepBox.append(ui.field('&Data range:', inp), h('div', { class: 'row', style: 'gap:16px' }, h('span', { text: 'Series in:' }), rows, cols));
      if (editing && !st.rangeChanged) stepBox.appendChild(h('div', { class: 'hint', text: 'Leave the range empty to keep the chart\'s current series.' }));
    }
    function step3() {
      stepBox.textContent = '';
      const tIn = (key, label) => { const i = h('input', { type: 'text', value: st[key], style: 'width:100%' }); i.addEventListener('input', () => { st[key] = i.value; drawPreview(); }); return ui.field(label, i); };
      const titles = h('div', { class: 'col' }, tIn('title', 'Chart &title:'), tIn('axX', '&Category (X) axis:'), tIn('axY', '&Value (Y) axis:'));
      const grid = h('div', { class: 'col' }, ui.check('&Major gridlines (Value axis)', st.gridY, (v) => { st.gridY = v; drawPreview(); }));
      const leg = h('div', { class: 'col' }, ui.check('&Show legend', st.legend !== 'none', (v) => { st.legend = v ? 'r' : 'none'; drawPreview(); }),
        ...[['b', '&Bottom'], ['tr', 'C&orner'], ['t', 'To&p'], ['r', '&Right'], ['l', '&Left']].map(([k, l]) => ui.radio('cw-leg', l, st.legend === k, () => { st.legend = k; drawPreview(); })));
      const lab = h('div', { class: 'col' }, ui.check('&Value', st.labels, (v) => { st.labels = v; drawPreview(); }));
      stepBox.appendChild(ui.tabs([{ label: 'Titles', body: titles }, { label: 'Gridlines', body: grid }, { label: 'Legend', body: leg }, { label: 'Data Labels', body: lab }]));
    }
    function step4() {
      stepBox.textContent = '';
      const nm = h('input', { type: 'text', value: st.sheetName, style: 'width:180px' });
      nm.addEventListener('input', () => { st.sheetName = nm.value; });
      stepBox.append(h('div', { text: 'Place chart:' }),
        h('div', { class: 'row', style: 'gap:8px;margin:6px 0' }, ui.radio('cw-where', 'As new &sheet:', st.where === 'sheet', () => { st.where = 'sheet'; }), nm),
        h('div', { class: 'row', style: 'gap:8px;margin:6px 0' }, ui.radio('cw-where', 'As &object in:', st.where === 'object', () => { st.where = 'object'; }), h('span', { text: sh.name })));
    }
    const steps = editing ? [step1, step2, step3] : [step1, step2, step3, step4];
    const titles = ['Chart Type', 'Chart Source Data', 'Chart Options', 'Chart Location'];
    return new Promise((resolve) => {
      let ok = false;
      const d = ui.dialog({
        title: 'Chart Wizard - Step 1 of ' + steps.length + ' - Chart Type', body, width: 640,
        buttons: [
          { label: 'Cancel' },
          { label: '< &Back', onClick: () => { if (st.step > 1) { st.step--; show(); } return false; } },
          { label: '&Next >', onClick: () => { if (st.step < steps.length) { st.step++; show(); } return false; } },
          { label: '&Finish', primary: true, onClick: () => { ok = true; } },
        ],
      });
      function show() { steps[st.step - 1](); d.el.querySelector('.dlg-ttl').textContent = `Chart Wizard - Step ${st.step} of ${steps.length} - ${titles[st.step - 1]}`; d.buttons[1].disabled = st.step === 1; d.buttons[2].disabled = st.step === steps.length; drawPreview(); }
      show();
      d.done.then(() => {
        if (!ok) { resolve(null); return; }
        const m = model();
        if (!m.series.length) { ui.msg('The chart has no data. Select a range with numbers first.', { icon: 'warn' }); resolve(null); return; }
        const chart = Object.assign({}, editing ? editing.chart : {}, m, { dirty: true, imported: editing ? editing.chart.imported : false });
        delete chart.refs;
        if (editing) {
          const nd = Object.assign({}, editing, { chart });
          L.ops.tx(wb, 'Chart Options', () => L.ops.setDrawings(sh, sh.drawings.map((x) => (x === editing ? nd : x))));
          G.selectObject(nd); G.syncObjects(true);
          resolve(nd); return;
        }
        if (st.where === 'sheet') {
          let name = (st.sheetName || '').trim() || wb.nextSheetName('Chart');
          if (wb.sheetByName(name)) name = wb.nextSheetName('Chart');
          let ns;
          L.ops.tx(wb, 'Insert Chart', () => {
            ns = L.ops.addSheet(wb, name, wb.active);
            ns.kind = 'chartsheet';
            ns.view.grid = false; ns.view.headings = false;
            L.ops.setDrawings(ns, [{ kind: 'chart', name: 'Chart 1', id: 2, chart, anchor: { type: 'abs', x: 0, y: 0, w: 640, h: 440 } }]);
          });
          wb.active = wb.sheets.indexOf(ns);
          L.app.renderTabs(); G.paint(true); G.syncObjects(true);
          resolve(ns); return;
        }
        const s = G.sel();
        const g = L.layout.geo(sh);
        const startC = Math.min(rg.c2 + 2, M.MAXC - 8), startR = Math.max(0, rg.r1);
        const dr = { kind: 'chart', name: 'Chart ' + (sh.drawings.filter((x) => x.kind === 'chart').length + 1), id: Math.max(1, ...sh.drawings.map((x) => x.id || 0)) + 1, chart, anchor: { type: 'two', editAs: 'twoCell', from: { r: startR, c: startC, rOff: 0, cOff: 0 }, to: { r: startR, c: startC, rOff: 0, cOff: 0 } } };
        /* about 5 × 3 inches, like Excel's default */
        const wpx = 480, hpx = 288;
        let cc = startC, acc = 0; while (acc + g.cols.size(cc) <= wpx && cc < M.MAXC - 1) { acc += g.cols.size(cc); cc++; }
        dr.anchor.to.c = cc; dr.anchor.to.cOff = (wpx - acc) * 0.75;
        let rr = startR; acc = 0; while (acc + g.rows.size(rr) <= hpx && rr < M.MAXR - 1) { acc += g.rows.size(rr); rr++; }
        dr.anchor.to.r = rr; dr.anchor.to.rOff = (hpx - acc) * 0.75;
        void s;
        L.ops.tx(wb, 'Insert Chart', () => L.ops.setDrawings(sh, sh.drawings.concat([dr])));
        G.selectObject(dr); G.syncObjects(true);
        L.app.updateToolbars();
        resolve(dr);
      });
    });
  };
  const O_clip = (sh, rg) => L.ops.clip(sh, rg);
})(typeof window !== 'undefined' ? window : globalThis);
