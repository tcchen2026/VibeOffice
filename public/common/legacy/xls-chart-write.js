/* Native DrawingML charts from BIFF chart records.
 * Derived from Apache OpenOffice main/sc/source/filter/excel/xichart.cxx (Apache-2.0).
 * Derived from Apache POI hssf/record/chart/ (Apache-2.0).
 * Chart element order follows ECMA-376 and Office-authored corpus charts. */
(function (root) {
  'use strict';
  const B = root.L.legacy, X = B.xls, C = X.charts, R = C.record, el = B.el, val = C.val, N = s => s;
  const NS = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  // IDs are local to a chart part. These pairs occur in Office-authored
  // WithChartSheet.xlsx and chartTitle_withTitle.xlsx; binary axes have no IDs.
  const AXIS_IDS = [[53641216, 53642752, 51001600], [410637520, 410639088, 51003392]];
  const TYPES = new Set([R.BAR, R.LINE, R.PIE, R.AREA, R.SCATTER, R.RADAR, R.RADAR_AREA, R.SURFACE, R.OF_PIE]);
  const finite = n => Number.isFinite(n) ? n : null;

  function cache(data, numeric, styles) {
    const body = (numeric ? el('c:formatCode', {}, B.xml(C.numFormat(styles, data.format))) : '') +
      val('ptCount', data.count) + data.points.map((point, index) => {
        if (point.type === 'blank' || point.value == null || numeric && point.type === 's') return '';
        return el('c:pt', { idx: index }, el('c:v', {}, B.xml(point.value)));
      }).join('');
    if (data.formula) return el(numeric ? 'c:numRef' : 'c:strRef', {},
      el('c:f', {}, B.xml(data.formula)) + el(numeric ? 'c:numCache' : 'c:strCache', {}, body));
    return el(numeric ? 'c:numLit' : 'c:strLit', {}, body);
  }
  function labels(node) {
    const r = C.reader(node, R.LABELS);
    if (!r) return '';
    const flags = r.u16(0);
    if (!(flags & 0x37)) return '';
    return el('c:dLbls', {}, val('showLegendKey', 0) + val('showVal', flags & 1 ? 1 : 0) +
      val('showCatName', flags & 20 ? 1 : 0) + val('showSerName', 0) +
      val('showPercent', flags & 6 ? 1 : 0) + val('showBubbleSize', flags & 32 ? 1 : 0));
  }
  function trendlines(chart, series, styles) {
    return chart.series.filter(s => s.parent === series.index && C.child(s.node, R.TREND)).map(s => {
      const r = C.reader(s.node, R.TREND), kind = r.u8(0), order = r.u8(1);
      const type = kind === 0 ? order === 1 ? 'linear' : 'poly' :
        ['linear', 'exp', 'log', 'power', 'movingAvg'][kind];
      if (!type) {
        chart.issues.push({ kind: 'format', error: 'Unknown trendline type' });
        return '';
      }
      const format = C.child(s.node, R.DATA_FORMAT);
      return el('c:trendline', {}, el('c:spPr', {}, C.line(C.reader(format, R.LINE_FORMAT), styles)) +
        val('trendlineType', type) + (type === 'poly' ? val('order', Math.max(2, Math.min(6, order))) : '') +
        (type === 'movingAvg' ? val('period', Math.max(2, order)) : '') +
        (r.f64(12) > 0 ? val('forward', finite(r.f64(12))) : '') +
        (r.f64(20) > 0 ? val('backward', finite(r.f64(20))) : '') +
        (Number.isFinite(r.f64(2)) ? val('intercept', r.f64(2)) : '') +
        val('dispRSqr', r.u8(11) ? 1 : 0) + val('dispEq', r.u8(10) ? 1 : 0));
    }).join('');
  }
  function errorBars(chart, series, book, si, styles, formulas) {
    const auxiliary = chart.series.filter(s => s.parent === series.index && C.child(s.node, R.ERROR_BAR));
    return ['x', 'y'].map((direction, axis) => {
      const positive = auxiliary.find(s => C.reader(s.node, R.ERROR_BAR).u8(0) === axis * 2 + 1);
      const negative = auxiliary.find(s => C.reader(s.node, R.ERROR_BAR).u8(0) === axis * 2 + 2);
      if (!positive && !negative) return '';
      const source = positive || negative, r = C.reader(source.node, R.ERROR_BAR);
      const type = [null, 'percentage', 'fixedVal', 'stdDev', 'cust', 'stdErr'][r.u8(1)];
      if (!type) {
        chart.issues.push({ kind: 'format', error: 'Unknown error bar source' });
        return '';
      }
      let body = val('errDir', direction) +
        val('errBarType', positive && negative ? 'both' : positive ? 'plus' : 'minus') +
        val('errValType', type) + val('noEndCap', r.u8(2) ? 0 : 1);
      if (type === 'cust') {
        for (const [tag, series] of [['plus', positive], ['minus', negative]]) {
          if (series) body += el('c:' + tag, {}, cache(C.data(chart, series, 1, book, si, formulas), true, styles));
        }
      } else if (type !== 'stdErr') body += val('val', finite(r.f64(4)) ?? 0);
      return el('c:errBars', {}, body + el('c:spPr', {},
        C.line(C.reader(C.child(source.node, R.DATA_FORMAT), R.LINE_FORMAT), styles)));
    }).join('');
  }

  function seriesXml(chart, series, group, kind, styles, book, si, formulas) {
    const local = C.children(series.node, R.DATA_FORMAT);
    const formats = [...local, ...C.children(chart.node, R.DATA_FORMAT).filter(n => n.r.u16(2) === series.index)];
    const format = formats.find(n => n.r.u16(0) === 65535) || C.child(group.node, R.DATA_FORMAT);
    const order = format?.r.u16(4) ?? series.index;
    const values = C.data(chart, series, 1, book, si, formulas);
    const categories = C.data(chart, series, 2, book, si, formulas);
    const title = C.data(chart, series, 0, book, si, formulas);
    const numericCategory = !categories.points.some(p => p.type === 's');
    let text = '';
    if (title.formula) text = el('c:tx', {}, cache(title, false, styles));
    else if (title.points[0]?.value != null) text = el('c:tx', {}, el('c:v', {}, B.xml(title.points[0].value)));
    const line = kind === 'line' || kind === 'scatter' || kind === 'radar';
    const color = C.autoColor(styles, order, line);
    const style = C.frame(format, styles, { fill: color, line: line ? color : '000000' }, chart);
    let body = val('idx', series.index) + val('order', order) + text + style;
    if (['bar', 'bubble'].includes(kind)) {
      const area = C.reader(format, R.AREA_FORMAT);
      body += val('invertIfNegative', area?.u16(10) & 2 ? 1 : 0);
    }
    if (['line', 'scatter', 'radar'].includes(kind)) body += C.marker(format, styles, order);
    if (['pie', 'doughnut', 'ofPie'].includes(kind)) {
      const explosion = C.reader(format, R.PIE_FORMAT)?.u16(0);
      if (explosion) body += val('explosion', Math.min(400, explosion));
    }
    const points = new Map(formats.filter(n => n.r.u16(0) < values.count).map(n => [n.r.u16(0), n]));
    if (kind === 'surface' && points.size) {
      chart.issues.push({ kind: 'format', error: 'Surface chart point formatting could not be kept' });
      points.clear();
    }
    if (group.flags & 1 && ['bar', 'pie', 'doughnut', 'ofPie'].includes(kind)) {
      for (let i = 0; i < values.count; i++) if (!points.has(i)) points.set(i, null);
    }
    body += [...points].sort(([a], [b]) => a - b).map(([index, format]) => {
      const color = C.autoColor(styles, index, line);
      let item = val('idx', index);
      if (['line', 'scatter', 'radar'].includes(kind)) item += C.marker(format, styles, index);
      const explosion = C.reader(format, R.PIE_FORMAT)?.u16(0);
      if (explosion && ['pie', 'doughnut', 'ofPie'].includes(kind)) item += val('explosion', Math.min(400, explosion));
      return el('c:dPt', {}, item + C.frame(format, styles, { fill: color, line: '000000' }, chart));
    }).join('');
    if (kind !== 'surface') body += labels(format);
    if (['bar', 'line', 'area', 'scatter', 'bubble'].includes(kind)) {
      body += trendlines(chart, series, styles) + errorBars(chart, series, book, si, styles, formulas);
    }
    if (kind === 'scatter' || kind === 'bubble') {
      body += el('c:xVal', {}, cache(categories, true, styles)) + el('c:yVal', {}, cache(values, true, styles));
      if (kind === 'bubble') body += el('c:bubbleSize', {},
        cache(C.data(chart, series, 3, book, si, formulas), true, styles)) +
        val('bubble3D', C.reader(format, R.SERIES_FORMAT)?.u16(0) & 2 ? 1 : 0);
    } else body += el('c:cat', {}, cache(categories, numericCategory, styles)) +
      el('c:val', {}, cache(values, true, styles));
    if (kind === 'line' || kind === 'scatter') {
      const flags = C.reader(format, R.SERIES_FORMAT)?.u16(0) || 0;
      body += val('smooth', flags & 1 ? 1 : 0);
    }
    return el('c:ser', {}, body);
  }

  function groupInfo(group) {
    const record = group.node.children.find(r => TYPES.has(r.id));
    if (!record) B.fail('XLS_CHART_TYPE', N('This chart type cannot be converted yet'));
    const r = record.r;
    let kind, flags = 0;
    switch (record.id) {
      case R.BAR:
        kind = 'bar';
        flags = r.u16(4);
        break;
      case R.LINE:
        kind = 'line';
        flags = r.u16(0);
        break;
      case R.AREA:
        kind = 'area';
        flags = r.u16(0);
        break;
      case R.PIE:
        kind = r.u16(2) ? 'doughnut' : 'pie';
        break;
      case R.SCATTER:
        kind = r.length >= 6 && r.u16(4) & 1 ? 'bubble' : 'scatter';
        break;
      case R.RADAR:
      case R.RADAR_AREA:
        kind = 'radar';
        break;
      case R.SURFACE:
        kind = 'surface';
        break;
      case R.OF_PIE:
        kind = 'ofPie';
        break;
    }
    const threeD = C.reader(group.node, R.CHART_3D);
    const stock = kind === 'line' && !threeD && group.node.children.some(r =>
      r.id === R.CHART_LINE && r.r.u16(0) === 1);
    if (stock) kind = 'stock';
    return { kind, flags, r, record, threeD, horizontal: kind === 'bar' && !!(flags & 1) };
  }

  C.write = function (chart, book, si, styles, formulas) {
    let plot = '', axesXml = '', threeD = null, legend = null;
    const titleFor = (parent, link) => C.children(parent, R.TEXT)
      .find(n => C.reader(n, R.OBJECT_LINK)?.u16(0) === link);
    const defaultText = C.children(chart.node, R.TEXT).find(n => !C.child(n, R.OBJECT_LINK));
    const font = C.reader(defaultText, R.FONT)?.u16(0) || 0;
    const regular = chart.series.filter(series => series.parent == null);
    for (const series of chart.series) {
      if (series.parent == null) continue;
      const parent = chart.series[series.parent];
      if (!parent || parent.parent != null) {
        chart.issues.push({ kind: 'data', error: 'A chart trendline or error bar has an invalid parent series' });
      }
    }
    const automaticTitle = regular.length === 1 && chart.axes.length === 1 ?
      C.data(chart, regular[0], 0, book, si, formulas).points[0]?.value : '';
    const title = C.title(titleFor(chart.node, 1), book, si, styles, formulas, automaticTitle);
    for (const axes of chart.axes) {
      const ids = AXIS_IDS[axes.index & 1], writtenGroups = [];
      let scatter = false, horizontal = false, needAxes = false, zAxis = false;
      for (const group of axes.groups) {
        const info = groupInfo(group), { kind, r, flags } = info;
        const series = chart.series.filter(s => s.parent == null && s.group === group.index);
        const is3D = info.threeD && ['bar', 'line', 'area', 'pie', 'surface'].includes(kind);
        let tag = kind + (is3D ? '3D' : '') + 'Chart';
        if (kind === 'surface' && !is3D) tag = 'surfaceChart';
        let body = '';
        if (kind === 'bar') body += val('barDir', flags & 1 ? 'bar' : 'col') +
          val('grouping', flags & 4 ? 'percentStacked' : flags & 2 ? 'stacked' : 'clustered');
        else if (kind === 'line' || kind === 'area') body +=
          val('grouping', flags & 2 ? 'percentStacked' : flags & 1 ? 'stacked' : 'standard');
        else if (kind === 'scatter') body += val('scatterStyle', 'lineMarker');
        else if (kind === 'radar') body += val('radarStyle', info.record.id === R.RADAR_AREA ? 'filled' : 'marker');
        else if (kind === 'surface') body += val('wireframe', r.u16(0) & 1 ? 0 : 1);
        else if (kind === 'ofPie') body += val('ofPieType', r.u8(0) === 1 ? 'pie' : 'bar');
        if (!['surface', 'stock'].includes(kind)) body += val('varyColors', group.flags & 1 ? 1 : 0);
        body += series.map(s => seriesXml(chart, s, group, kind, styles, book, si, formulas)).join('');
        if (kind === 'bar') {
          body += val('gapWidth', Math.min(500, r.u16(2)));
          if (!is3D) body += val('overlap', Math.max(-100, Math.min(100, r.i16(0))));
        }
        if (kind === 'pie' && !is3D || kind === 'doughnut') body += val('firstSliceAng', r.u16(0) % 360);
        if (kind === 'doughnut') body += val('holeSize', Math.max(10, Math.min(90, r.u16(2))));
        if (kind === 'bubble') body += val('bubbleScale', Math.max(0, Math.min(300, r.u16(0)))) +
          val('showNegBubbles', r.u16(4) & 2 ? 1 : 0) + val('sizeRepresents', r.u16(2) === 2 ? 'w' : 'area');
        for (let i = 0; i < group.node.children.length; i++) {
          const child = group.node.children[i];
          if (child.id !== R.CHART_LINE) continue;
          const type = child.r.u16(0), next = group.node.children[i + 1];
          const allowed = type === 0 && ['line', 'area'].includes(kind) ||
            type === 1 && ['line', 'stock'].includes(kind);
          if (allowed) body += el(type ? 'c:hiLowLines' : 'c:dropLines', {}, el('c:spPr', {},
            C.line(next?.id === R.LINE_FORMAT ? next.r : null, styles)));
        }
        const dropBars = C.children(group.node, R.DROP_BAR);
        if (dropBars.length && (kind === 'stock' || kind === 'line')) {
          body += el('c:upDownBars', {}, val('gapWidth', dropBars[0].r.u16(0)) +
            el('c:upBars', {}, C.frame(dropBars[0], styles, { fill: 'FFFFFF' }, chart)) +
            el('c:downBars', {}, C.frame(dropBars[1], styles, { fill: '000000' }, chart)));
        }
        if (!['pie', 'doughnut', 'ofPie'].includes(kind)) {
          needAxes = true;
          body += val('axId', ids[0]) + val('axId', ids[1]);
          if (is3D && kind !== 'bar' || kind === 'surface') {
            body += val('axId', ids[2]);
            zAxis = true;
          }
        }
        writtenGroups.push(el('c:' + tag, {}, body));
        scatter ||= kind === 'scatter' || kind === 'bubble';
        horizontal ||= info.horizontal;
        threeD ||= info.threeD;
        legend ||= C.child(group.node, R.LEGEND);
      }
      plot += writtenGroups.join('');
      if (needAxes) {
        const axis = index => C.children(axes.node, R.AXIS).find(n => n.r.u16(0) === index);
        const firstSeries = regular.find(series => axes.groups.some(group => group.index === series.group));
        const sourceFormat = kind => firstSeries?.links.get(kind)?.format || 0;
        axesXml += C.axis(axis(0), ids[0], ids[1], scatter ? 'valAx' : 'catAx', styles,
          C.title(titleFor(axes.node, 3), book, si, styles, formulas), font, !!axes.index, horizontal,
          sourceFormat(2), axis(1));
        axesXml += C.axis(axis(1), ids[1], ids[0], 'valAx', styles,
          C.title(titleFor(axes.node, 2), book, si, styles, formulas), font, !!axes.index, horizontal,
          sourceFormat(1), axis(0));
        if (zAxis) axesXml += C.axis(axis(2), ids[2], ids[1], 'serAx', styles,
          C.title(titleFor(axes.node, 7), book, si, styles, formulas), font, !!axes.index, horizontal, 0, axis(1));
      }
    }
    if (!plot) B.fail('XLS_CHART_TYPE', N('The chart has no convertible chart groups'));
    const primary = chart.axes[0]?.node;
    const plotFrame = C.child(primary, R.FRAME);
    const plotStyle = plotFrame ? C.frame(plotFrame, styles, { fill: 'C0C0C0' }, chart) :
      el('c:spPr', {}, el('a:noFill') + el('a:ln', {}, el('a:noFill')));
    const properties = C.reader(chart.node, R.PROPERTIES), flags = properties?.u16(0) || 0;
    const plotPosition = C.reader(primary, R.POSITION);
    const manualPlot = book.version === 'BIFF5/7' || !!(flags & 16);
    let layout = '';
    if (manualPlot) {
      const usePosition = plotPosition && plotPosition.u16(0) === 2 && plotPosition.u16(2) === 2;
      const r = usePosition ? plotPosition : primary.r, start = usePosition ? 4 : 2;
      const values = [0, 1, 2, 3].map(i => (usePosition ? r.i16(start + i * 4) : r.i32(start + i * 4)) / 4000);
      layout = el('c:manualLayout', {}, val('layoutTarget', usePosition ? 'outer' : 'inner') +
        val('xMode', 'edge') + val('yMode', 'edge') + val('wMode', 'factor') + val('hMode', 'factor') +
        ['x', 'y', 'w', 'h'].map((tag, i) => val(tag, values[i])).join(''));
    }
    let view = '';
    if (threeD) {
      const rotation = threeD.u16(0) % 360, elevation = Math.max(-90, Math.min(90, threeD.i16(2)));
      view = el('c:view3D', {}, val('rotX', elevation) + val('hPercent', Math.max(5, Math.min(500, threeD.u16(6)))) +
        val('rotY', rotation) + val('depthPercent', Math.max(20, Math.min(2000, threeD.u16(8)))) +
        val('rAngAx', threeD.u16(12) & 1 ? 0 : 1) + val('perspective', Math.min(240, threeD.u16(4))));
    }
    let legendXml = '';
    if (legend) {
      const position = ['b', 'tr', 't', 'r', 'l'][legend.r.u8(16)] || 'r';
      const text = C.text(C.child(legend, R.TEXT) || defaultText, book, si, styles, formulas);
      const pos = C.reader(legend, R.POSITION);
      let manual = '';
      if (pos && (manualPlot || legend.r.u8(16) === 7)) {
        manual = val('xMode', 'edge') + val('yMode', 'edge') + val('x', pos.i16(4) / 4000) +
          val('y', pos.i16(8) / 4000);
        if (pos.u16(2) === 1 && pos.i16(12) > 0 && pos.i16(16) > 0) {
          manual += val('w', pos.i16(12) * 65536 / chart.node.r.u32(8)) +
            val('h', pos.i16(16) * 65536 / chart.node.r.u32(12));
        }
      }
      legendXml = el('c:legend', {}, val('legendPos', position) +
        el('c:layout', {}, manual ? el('c:manualLayout', {}, manual) : undefined) + val('overlay', manual ? 1 : 0) +
        (C.child(legend, R.FRAME) ? C.frame(C.child(legend, R.FRAME), styles, {}, chart) : '') + (text?.txPr || ''));
    }
    const chartXml = el('c:chart', {}, title + val('autoTitleDeleted', title ? 0 : 1) + view +
      el('c:plotArea', {}, el('c:layout', {}, layout) + plot + axesXml + plotStyle) + legendXml +
      val('plotVisOnly', flags & 2 ? 1 : 0) + val('dispBlanksAs', ['gap', 'zero', 'span'][properties?.u8(2)] || 'gap'));
    return B.xmlDeclaration + el('c:chartSpace', { 'xmlns:c': NS, 'xmlns:a': B.ns.a,
      'xmlns:r': B.ns.officeRel.slice(0, -1) }, (book.date1904 ? val('date1904', 1) : '') + chartXml +
      C.frame(C.child(chart.node, R.FRAME), styles, { fill: 'FFFFFF' }, chart) +
      (C.text(defaultText, book, si, styles, formulas)?.txPr || ''));
  };

  X.prepareCharts = function (book, losses) {
    book.charts = book.sheets.map(sheet => {
      const result = [];
      let streams;
      try { streams = C.streams(sheet); }
      catch (error) {
        if (error.name !== 'LegacyError') throw error;
        X.drawingLoss(losses, sheet, 'chart', null, error);
        return result;
      }
      for (const stream of streams) {
        try { result.push(C.read(book, stream)); }
        catch (error) {
          if (error.name !== 'LegacyError') throw error;
          X.drawingLoss(losses, sheet, 'chart', null, error);
        }
      }
      return result;
    });
  };
  X.writeObjectChart = function (book, si, object, styles, pkg, drawingPart, formulas, losses) {
    const chart = book.charts[si].find(c => object ? c.objectOffset === object.record.offset : c.sheet);
    if (!chart) return '';
    const name = 'xl/charts/chart' + ((book.chartSerial || 0) + 1) + '.xml';
    const xml = C.write(chart, book, si, styles, formulas);
    pkg.add(name, xml, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
    book.chartSerial = (book.chartSerial || 0) + 1;
    if (!chart.sheet && chart.records.some(record => record.id === R.OBJ)) {
      X.drawingLoss(losses, book.sheets[si], 'chartDrawing', object?.shape,
        'Drawing objects inside the chart were omitted');
    }
    for (const issue of chart.issues) {
      const kind = issue.kind === 'data' ? 'chartData' : issue.kind === 'reference' ? 'chartReference' : 'format';
      X.drawingLoss(losses, book.sheets[si], kind, object?.shape, issue.error);
    }
    const rid = pkg.rel(drawingPart, 'chart', '../charts/' + name.split('/').at(-1));
    chart.written = true;
    return el('c:chart', { 'xmlns:c': NS, 'xmlns:r': B.ns.officeRel.slice(0, -1), 'r:id': rid });
  };
})(typeof window !== 'undefined' ? window : globalThis);
