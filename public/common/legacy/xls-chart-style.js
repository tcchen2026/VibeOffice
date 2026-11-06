/* Binary Excel chart formatting and axes.
 * Derived from Apache OpenOffice main/sc/source/filter/excel/{xichart,xlchart}.cxx (Apache-2.0).
 * Derived from Apache POI hssf/record/chart/{AreaFormat,LineFormat,ValueRange,AxisOptions}Record.java (Apache-2.0).
 * ECMA-376 DrawingML chart vocabulary; explicit binary colours avoid theme changes. */
(function (root) {
  'use strict';
  const B = root.L.legacy, X = B.xls, C = X.charts, R = C.record, el = B.el;
  const val = C.val = (name, value) => el('c:' + name, { val: value });
  const rgb = n => [n & 255, n >>> 8 & 255, n >>> 16 & 255]
    .map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase();
  const solid = value => el('a:solidFill', {}, el('a:srgbClr', { val: value }));
  const indexed = (styles, index, fallback) => styles.palette[index] || fallback;
  const lineColors = [32, 33, 34, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47,
    48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 8, 9, 10, 11, 12, 13,
    14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 63];
  C.autoColor = (styles, index, line = false) => indexed(styles,
    line ? lineColors[index % lineColors.length] : (index % 56 + 16) % 56 + 8, '000000');
  C.line = function (r, styles, fallback = '000000', defaultWidth = 12700) {
    if (!r) return el('a:ln', { w: defaultWidth }, solid(fallback));
    const pattern = r.u16(4), weight = r.i16(6), auto = r.u16(8) & 1;
    const color = auto ? fallback : r.length >= 12 ? indexed(styles, r.u16(10), rgb(r.u32(0))) : rgb(r.u32(0));
    const width = weight < 0 ? 3175 : Math.max(0, Math.min(3, weight) + 1) * 12700;
    const dash = ['solid', 'dash', 'sysDot', 'dashDot', 'sysDashDotDot'][pattern] || 'solid';
    return el('a:ln', { w: width }, pattern === 5 ? el('a:noFill') : solid(color) +
      el('a:prstDash', { val: dash }));
  };
  C.fill = function (r, styles, fallback = 'FFFFFF') {
    if (!r) return solid(fallback);
    const pattern = r.u16(8), auto = r.u16(10) & 1;
    if (!pattern) return el('a:noFill');
    const front = auto ? fallback : r.length >= 16 ? indexed(styles, r.u16(12), rgb(r.u32(0))) : rgb(r.u32(0));
    const back = r.length >= 16 ? indexed(styles, r.u16(14), rgb(r.u32(4))) : rgb(r.u32(4));
    if (pattern === 1) return solid(front);
    const preset = ['pct50', 'pct75', 'pct25', 'horz', 'vert', 'dnDiag', 'upDiag', 'diagCross',
      'cross', 'ltHorz', 'ltVert', 'ltDnDiag', 'ltUpDiag', 'ltCross', 'ltDiagCross', 'pct10', 'pct5'][pattern - 2];
    return preset ? el('a:pattFill', { prst: preset }, el('a:fgClr', {}, el('a:srgbClr', { val: front })) +
      el('a:bgClr', {}, el('a:srgbClr', { val: back }))) : solid(front);
  };
  C.frame = function (node, styles, defaults = {}, chart) {
    let fill = C.fill(C.reader(node, R.AREA_FORMAT), styles, defaults.fill || 'FFFFFF');
    const gel = C.reader(node, R.GEL_FRAME);
    if (gel) {
      try {
        const properties = new Map();
        for (const rec of B.officeart.records(gel.bytes)) {
          if (rec.type === 0xf00b) B.officeart.mergeProperties(properties, B.officeart.properties(rec));
        }
        const palette = { ...styles.palette, 64: '000000', 65: 'FFFFFF' };
        fill = B.shapes.style({ type: 1, properties }, { palette, scheme: palette, components: true }).fill;
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        chart?.issues.push({ kind: 'format', error });
      }
    }
    const line = C.line(C.reader(node, R.LINE_FORMAT), styles, defaults.line || '000000');
    return el('c:spPr', {}, fill + line);
  };
  C.marker = function (node, styles, index) {
    const r = C.reader(node, R.MARKER);
    if (!r) return '';
    const auto = r.u16(10) & 1, kind = auto ? [2, 1, 3, 4, 5, 8, 9, 6, 7][index % 9] : r.u16(8);
    const name = ['none', 'square', 'diamond', 'triangle', 'x', 'star', 'dash', 'dot', 'circle', 'plus'][kind];
    const defaultColor = C.autoColor(styles, index, true);
    const front = auto ? defaultColor : r.length >= 20 ? indexed(styles, r.u16(14), rgb(r.u32(4))) : rgb(r.u32(4));
    const line = auto ? defaultColor : r.length >= 20 ? indexed(styles, r.u16(12), rgb(r.u32(0))) : rgb(r.u32(0));
    const size = r.length >= 20 ? Math.max(2, Math.min(72, Math.round(r.u32(16) / 20))) : 5;
    return el('c:marker', {}, val('symbol', name || 'none') + (kind ? val('size', size) : '') +
      el('c:spPr', {}, (r.u16(10) & 16 ? el('a:noFill') : solid(front)) + el('a:ln', {},
        r.u16(10) & 32 ? el('a:noFill') : solid(line))));
  };
  C.numFormat = (styles, id) => styles.formats.get(id) || ({
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%',
    11: '0.00E+00', 12: '# ?/?', 13: '# ??/??', 14: 'm/d/yy', 15: 'd-mmm-yy',
    16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm',
    21: 'h:mm:ss', 22: 'm/d/yy h:mm', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mmss.0', 49: '@'
  })[id] || 'General';

  C.text = function (node, book, si, styles, formulas, fallback = '') {
    if (!node && !fallback) return null;
    let value = '', formula = '', font = C.reader(node, R.FONT)?.u16(0) || 0;
    const link = C.reader(node, R.LINK);
    const string = C.reader(node, R.STRING);
    if (string) value = X.string(string, 2, book.version === 'BIFF8', book.codepage, true).value;
    if (link?.u16(6)) {
      try {
        formula = C.formula(link.sub(8, link.u16(6)), book, si, formulas);
        if (!value) value = String(C.referenceValues(formula, book, si, formulas)[0]?.value ?? '');
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
      }
    }
    value ||= fallback == null ? '' : String(fallback);
    const textRecord = node?.r;
    const rotation = textRecord?.length >= 32 ? textRecord.u16(30) : 0;
    const body = { rot: rotation > 0 && rotation <= 90 ? -rotation * 60000 :
      rotation > 90 && rotation <= 180 ? (rotation - 90) * 60000 : null,
    vert: rotation === 255 ? 'wordArtVert' : null };
    const fontXml = (id, tag) => {
      let xml = X.drawingFont(styles, id, tag);
      if (textRecord?.length >= 26 && !(textRecord.u16(24) & 1)) {
        const color = textRecord.length >= 28 ? indexed(styles, textRecord.u16(26), rgb(textRecord.u32(4))) :
          rgb(textRecord.u32(4));
        xml = xml.replace(/<a:solidFill>.*?<\/a:solidFill>/, solid(color));
      }
      return xml;
    };
    const runs = C.reader(node, R.FORMAT_RUNS), boundaries = new Map([[0, font]]);
    if (runs) {
      const count = runs.u16(0);
      runs.check(2, count * 4);
      for (let i = 0; i < count; i++) boundaries.set(runs.u16(2 + i * 4), runs.u16(4 + i * 4));
    }
    const positions = [...boundaries.keys()].filter(p => p <= value.length).sort((a, b) => a - b);
    const content = positions.map((start, i) => el('a:r', {}, fontXml(boundaries.get(start), 'a:rPr') +
      el('a:t', {}, B.xml(value.slice(start, positions[i + 1] ?? value.length))))).join('');
    const rich = el('a:bodyPr', body) + el('a:lstStyle') + el('a:p', {}, content || fontXml(font, 'a:endParaRPr'));
    const txPr = el('c:txPr', {}, el('a:bodyPr', body) + el('a:lstStyle') + el('a:p', {},
      el('a:pPr', {}, fontXml(font, 'a:defRPr')) + el('a:endParaRPr')));
    return { value, formula, rich, txPr };
  };
  C.title = function (node, book, si, styles, formulas, fallback) {
    const text = C.text(node, book, si, styles, formulas, fallback);
    if (!text?.value && !text?.formula || node?.r.length >= 26 && node.r.u16(24) & 64) return '';
    const tx = text.formula ? el('c:strRef', {}, el('c:f', {}, B.xml(text.formula)) +
      el('c:strCache', {}, val('ptCount', 1) + el('c:pt', { idx: 0 }, el('c:v', {}, B.xml(text.value))))) :
      el('c:rich', {}, text.rich);
    return el('c:title', {}, el('c:tx', {}, tx) + el('c:layout') + val('overlay', 0) + text.txPr);
  };

  C.axis = function (axis, ids, other, type, styles, title, font = 0, secondary = false,
    horizontal = false, sourceFormat = 0, crossingAxis) {
    const node = axis || { children: [] }, r = C.reader(node, R.VALUE_RANGE);
    const cat = C.reader(node, R.CATEGORY_RANGE), options = C.reader(node, R.AXIS_OPTIONS);
    const tick = C.reader(node, R.TICK), number = C.reader(node, R.NUMBER_FORMAT);
    const date = type === 'catAx' && options && options.u16(16) & 16;
    if (date) type = 'dateAx';
    const flags = r?.u16(40) || 0, categoryFlags = cat?.u16(6) || 0;
    let scaling = flags & 32 ? val('logBase', 10) : '';
    scaling += val('orientation', (r ? flags & 64 : categoryFlags & 4) ? 'maxMin' : 'minMax');
    if (r && !(flags & 2) && Number.isFinite(r.f64(8))) scaling += val('max', r.f64(8));
    if (r && !(flags & 1) && Number.isFinite(r.f64(0))) scaling += val('min', r.f64(0));
    if (date) {
      if (!(options.u16(16) & 2)) scaling += val('max', options.u16(2));
      if (!(options.u16(16) & 1)) scaling += val('min', options.u16(0));
    }
    let grid = '', line = null;
    const children = node.children || [];
    for (let i = 0; i < children.length; i++) {
      if (children[i].id !== R.AXIS_LINE) continue;
      const kind = children[i].r.u16(0), format = children[i + 1]?.id === R.LINE_FORMAT ? children[i + 1].r : null;
      if (kind === 0) line = format;
      else if (kind === 1 || kind === 2) grid += el(kind === 1 ? 'c:majorGridlines' : 'c:minorGridlines', {},
        el('c:spPr', {}, C.line(format, styles, '000000')));
    }
    const xAxis = node.r ? node.r.u16(0) === 0 : type !== 'valAx';
    const pos = xAxis !== horizontal ? secondary ? 't' : 'b' : secondary ? 'r' : 'l';
    const ticks = ['', 'in', 'out', 'cross'];
    const rotation = tick?.length >= 30 && !(tick.u16(24) & 32) ? tick.u16(28) : 0;
    const body = { rot: rotation > 0 && rotation <= 90 ? -rotation * 60000 :
      rotation > 90 && rotation <= 180 ? (rotation - 90) * 60000 : null,
    vert: rotation === 255 ? 'wordArtVert' : null };
    let fontXml = X.drawingFont(styles, C.reader(node, R.FONT)?.u16(0) ?? font, 'a:defRPr');
    if (tick?.length >= 26 && !(tick.u16(24) & 1)) {
      const color = tick.length >= 28 ? indexed(styles, tick.u16(26), rgb(tick.u32(20))) : rgb(tick.u32(20));
      fontXml = fontXml.replace(/<a:solidFill>.*?<\/a:solidFill>/, solid(color));
    }
    const text = el('c:txPr', {}, el('a:bodyPr', body) + el('a:lstStyle') + el('a:p', {},
      el('a:pPr', {}, fontXml)));
    // BIFF stores where the other axis crosses this one. DrawingML places the
    // crossing setting on that other axis (AOO XclImpChAxis::CreateAxis).
    const crossingValue = C.reader(crossingAxis, R.VALUE_RANGE);
    const crossingCategory = C.reader(crossingAxis, R.CATEGORY_RANGE);
    const crossingFlags = crossingValue?.u16(40) || 0, crossingCategoryFlags = crossingCategory?.u16(6) || 0;
    const atMaximum = crossingFlags & 128 || !crossingValue && crossingCategoryFlags & 2;
    let cross = val('crosses', atMaximum ? 'max' : 'autoZero');
    if (crossingValue && !(crossingFlags & 16) && !(crossingFlags & 128) && Number.isFinite(crossingValue.f64(32))) {
      cross = val('crossesAt', crossingValue.f64(32));
    } else if (crossingCategory && !(crossingCategoryFlags & 2) && crossingCategory.u16(0) !== 1) {
      cross = val('crossesAt', crossingCategory.u16(0));
    }
    let tail = '';
    if (type === 'valAx') {
      const crossing = C.reader(crossingAxis, R.CATEGORY_RANGE);
      tail = val('crossBetween', crossing?.u16(6) & 1 ? 'between' : 'midCat');
      if (r && !(flags & 4) && r.f64(16) > 0) tail += val('majorUnit', r.f64(16));
      if (r && !(flags & 8) && r.f64(24) > 0) tail += val('minorUnit', r.f64(24));
    } else if (type === 'catAx') {
      tail = val('auto', 1) + val('lblAlgn', 'ctr') + val('lblOffset', 100);
      if (cat?.u16(2) > 1) tail += val('tickLblSkip', cat.u16(2));
      if (cat?.u16(4) > 1) tail += val('tickMarkSkip', cat.u16(4));
      tail += val('noMultiLvlLbl', 0);
    } else if (type === 'dateAx') {
      const units = ['days', 'months', 'years'], f = options.u16(16);
      tail = val('auto', f & 128 ? 1 : 0) + val('lblOffset', 100) +
        val('baseTimeUnit', units[options.u16(12)] || 'days');
      if (!(f & 4)) tail += val('majorUnit', options.u16(4) || 1) +
        val('majorTimeUnit', units[options.u16(6)] || 'days');
      if (!(f & 8)) tail += val('minorUnit', options.u16(8) || 1) +
        val('minorTimeUnit', units[options.u16(10)] || 'days');
    }
    const numberFormat = el('c:numFmt', { formatCode: C.numFormat(styles, number ? number.u16(0) : sourceFormat),
      sourceLinked: number ? 0 : 1 });
    return el('c:' + type, {}, val('axId', ids) + el('c:scaling', {}, scaling) + val('delete', axis ? 0 : 1) +
      val('axPos', pos) + grid + title + numberFormat + val('majorTickMark', ticks[tick?.u8(0)] || 'none') +
      val('minorTickMark', ticks[tick?.u8(1)] || 'none') +
      val('tickLblPos', ['none', 'low', 'high', 'nextTo'][tick?.u8(2)] || 'nextTo') +
      el('c:spPr', {}, C.line(line, styles)) + text + val('crossAx', other) + cross + tail);
  };
})(typeof window !== 'undefined' ? window : globalThis);
