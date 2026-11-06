/* BIFF chart substreams, record groups and saved data.
 * Derived from Apache OpenOffice main/sc/source/filter/excel/{xichart,excform}.cxx (Apache-2.0).
 * Derived from Apache POI hssf/record/chart/{LinkedData,Series,SeriesIndex}Record.java (Apache-2.0).
 * MS-XLS 2.2.3.2: cached rows are point indices, columns are series indices. */
(function (root) {
  'use strict';
  const B = root.L.legacy, X = B.xls, C = X.charts = {};
  const R = C.record = Object.freeze({
    BOF: 0x809, EOF: 0xa, OBJ: 0x5d, NUMBER: 0x203, LABEL: 0x204, BOOLERR: 0x205, BLANK: 0x201,
    CHART: 0x1002, SERIES: 0x1003, DATA_FORMAT: 0x1006, LINE_FORMAT: 0x1007, MARKER: 0x1009,
    AREA_FORMAT: 0x100a, PIE_FORMAT: 0x100b, LABELS: 0x100c, STRING: 0x100d, GROUP: 0x1014,
    LEGEND: 0x1015, BAR: 0x1017, LINE: 0x1018, PIE: 0x1019, AREA: 0x101a, SCATTER: 0x101b,
    CHART_LINE: 0x101c, AXIS: 0x101d, TICK: 0x101e, VALUE_RANGE: 0x101f, CATEGORY_RANGE: 0x1020,
    AXIS_LINE: 0x1021, DROP_BAR: 0x103d, DEFAULT_TEXT: 0x1024, TEXT: 0x1025, FONT: 0x1026,
    OBJECT_LINK: 0x1027, FRAME: 0x1032, BEGIN: 0x1033, END: 0x1034, PLOT_AREA: 0x1035,
    CHART_3D: 0x103a, RADAR: 0x103e, SURFACE: 0x103f, RADAR_AREA: 0x1040, AXIS_PARENT: 0x1041,
    PROPERTIES: 0x1044, SERIES_GROUP: 0x1045, SERIES_PARENT: 0x104a, TREND: 0x104b,
    NUMBER_FORMAT: 0x104e, POSITION: 0x104f, FORMAT_RUNS: 0x1050, LINK: 0x1051,
    ERROR_BAR: 0x105b, SERIES_FORMAT: 0x105d, AXIS_OPTIONS: 0x1062, CACHE_INDEX: 0x1065,
    GEL_FRAME: 0x1066, OF_PIE: 0x1061, LABEL_OPTIONS: 0x86b
  });
  C.child = (node, type) => node?.children?.find(record => record.id === type);
  C.children = (node, type) => node?.children?.filter(record => record.id === type) || [];
  C.reader = (node, type) => C.child(node, type)?.r;
  const fail = text => B.fail('XLS_CHART', text);

  C.streams = function (sheet) {
    const result = [], records = sheet.records;
    let depth = 0, object = null;
    for (let i = 0; i < records.length; i++) {
      const record = records[i];
      if (record.id === R.BOF) {
        if (record.r.u16(2) === 0x20 && (depth === 1 || sheet.kind === 2 && depth === 0)) {
          const start = i;
          let level = 1;
          while (++i < records.length && level) {
            if (records[i].id === R.BOF) level++;
            else if (records[i].id === R.EOF) level--;
          }
          if (level) fail('The chart substream is incomplete');
          i--;
          result.push({ records: records.slice(start + 1, i), objectOffset: object?.offset,
            sheet: sheet.kind === 2 && depth === 0 });
          object = null;
        } else depth++;
      } else if (record.id === R.EOF) depth--;
      else if (record.id === R.OBJ && depth === 1) object = record;
    }
    return result;
  };

  function groups(records) {
    const result = { children: [] }, stack = [result];
    for (const record of records) {
      if (record.id === R.BEGIN) {
        const parent = stack.at(-1).children.at(-1);
        if (!parent || stack.length >= B.limits.depth) fail('Invalid chart record nesting');
        stack.push(parent);
      } else if (record.id === R.END) {
        if (stack.length < 2) fail('A chart record group ends before it starts');
        stack.pop();
      } else stack.at(-1).children.push({ ...record, children: [] });
    }
    if (stack.length !== 1) fail('A chart record group is incomplete');
    return result;
  }

  function caches(records, book) {
    const result = new Map();
    let kind = 0;
    for (const { id, r } of records) {
      if (id === R.CACHE_INDEX) {
        kind = r.u16(0);
        continue;
      }
      if (!kind || ![R.NUMBER, R.LABEL, R.BOOLERR, R.BLANK].includes(id)) continue;
      const point = r.u16(0), series = r.u16(2);
      if (point > 65535 || series > 255) fail('Invalid chart cache address');
      let value = null, type = 'blank';
      if (id === R.NUMBER) {
        value = r.f64(6);
        type = 'n';
      } else if (id === R.LABEL) {
        value = X.string(r, 6, book.version === 'BIFF8', book.codepage).value;
        type = 's';
      } else if (id === R.BOOLERR) {
        type = r.u8(7) ? 'e' : 'n';
        value = type === 'e' ? X.formulaErrors[r.u8(6)] : r.u8(6) ? 1 : 0;
      }
      if (type === 'n' && !Number.isFinite(value)) {
        type = 'invalid';
        value = null;
      }
      const key = kind + ':' + series;
      if (!result.has(key)) result.set(key, new Map());
      result.get(key).set(point, { value, type });
    }
    return result;
  }

  C.read = function (book, stream) {
    const grouped = groups(stream.records), node = C.child(grouped, R.CHART);
    if (!node) fail('The chart record is missing');
    const series = C.children(node, R.SERIES).map((node, index) => {
      const r = node.r, links = new Map();
      for (let i = 0; i < node.children.length; i++) {
        const rec = node.children[i];
        if (rec.id !== R.LINK) continue;
        const r = rec.r, n = r.u16(6), link = { kind: r.u8(0), referenceType: r.u8(1),
          flags: r.u16(2), format: r.u16(4), tokens: r.sub(8, n) };
        if (node.children[i + 1]?.id === R.STRING) {
          link.text = X.string(node.children[++i].r, 2, book.version === 'BIFF8', book.codepage, true).value;
        }
        links.set(link.kind, link);
      }
      const parent = C.reader(node, R.SERIES_PARENT), group = C.reader(node, R.SERIES_GROUP);
      return { node, index, links, categoryType: r.u16(0), valueType: r.u16(2),
        categoryCount: r.u16(4), valueCount: r.u16(6),
        bubbleCount: r.length >= 12 ? r.u16(10) : 0, group: group?.u16(0) ?? 0,
        parent: parent?.u16(0) ? parent.u16(0) - 1 : null };
    });
    const axes = C.children(node, R.AXIS_PARENT).map(node => ({ node, index: node.r.u16(0),
      groups: C.children(node, R.GROUP).map(group => ({ node: group, index: group.r.u16(18),
        flags: group.r.u16(16) })) }));
    return { ...stream, node, series, axes, caches: caches(grouped.children, book), issues: [] };
  };

  function splitAreas(text) {
    const result = [];
    let quoted = false, start = 0, depth = 0;
    for (let i = 0; i < text.length; i++) {
      if (text[i] === "'") {
        if (quoted && text[i + 1] === "'") i++;
        else quoted = !quoted;
      } else if (!quoted) {
        if (text[i] === '(') depth++;
        else if (text[i] === ')') depth--;
        else if (text[i] === ',' && depth === 0) {
          result.push(text.slice(start, i));
          start = i + 1;
        }
      }
    }
    result.push(text.slice(start));
    return result;
  }
  C.referenceValues = function (text, book, si, formulas, seen = new Set()) {
    if (seen.has(text) || seen.size >= B.limits.depth) fail('Chart data references contain a loop');
    seen.add(text);
    const key = text;
    try {
      const output = [];
      if (text.startsWith('(') && text.endsWith(')')) text = text.slice(1, -1);
      for (const area of splitAreas(text)) {
        const match = /^(?:(.+)!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(area);
        if (!match) {
          const n = formulas.names.find(n => n.name === area && (n.scope === si || n.scope < 0));
          if (!n?.text) fail('The chart data cannot be resolved from worksheet cells');
          output.push(...C.referenceValues(n.text, book, n.scope >= 0 ? n.scope : si, formulas, seen));
          continue;
        }
        const name = match[1]?.replace(/^'|'$/g, '').replace(/''/g, "'");
        const sheet = name ? book.sheets.find(s => s.name === name) : book.sheets[si];
        if (!sheet) fail('The chart data refers to another workbook');
        const col = s => [...s.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
        const firstCol = col(match[2]), lastCol = col(match[4] || match[2]);
        const firstRow = +match[3] - 1, lastRow = +(match[5] || match[3]) - 1;
        const count = (lastCol - firstCol + 1) * (lastRow - firstRow + 1);
        if (firstCol < 0 || lastCol > 255 || firstRow < 0 || lastRow > 65535 ||
            lastCol < firstCol || lastRow < firstRow ||
            output.length + count > B.limits.cells) fail('The chart data range is too large');
        sheet.chartCells ||= new Map(sheet.cells.map(cell => [cell.row * 256 + cell.col, cell]));
        for (let row = firstRow; row <= lastRow; row++) for (let col = firstCol; col <= lastCol; col++) {
          const cell = sheet.chartCells.get(row * 256 + col);
          output.push(cell ? { type: cell.type === 'b' ? 'n' : cell.type,
            value: cell.type === 'e' ? X.formulaErrors[cell.value] : cell.type === 'b' ? +cell.value : cell.value } :
          { type: 'blank', value: null });
        }
      }
      return output;
    } finally { seen.delete(key); }
  };

  C.data = function (chart, series, kind, book, si, formulas) {
    const link = series.links.get(kind), count = kind === 1 ? series.valueCount :
      kind === 2 ? series.categoryCount : kind === 3 ? series.bubbleCount : 1;
    let formula = '', points = [];
    if (link?.tokens.length) {
      try { formula = C.formula(link.tokens, book, si, formulas); }
      catch (error) {
        if (error.name !== 'LegacyError') throw error;
        chart.issues.push({ kind: 'reference', error });
      }
    }
    const cache = chart.caches.get(kind + ':' + series.index);
    if (kind === 0 && link?.text != null) points = [{ type: 's', value: link.text }];
    else if (cache?.size) {
      const length = Math.max(count, ...[...cache.keys()].map(index => index + 1));
      points = Array.from({ length }, (_, i) => cache.get(i) || { type: 'blank', value: null });
      if (points.some(point => point.type === 'invalid')) {
        let cells = [];
        if (formula) {
          try { cells = C.referenceValues(formula, book, si, formulas); }
          catch (error) { if (error.name !== 'LegacyError') throw error; }
        }
        points = points.map((point, index) => {
          if (point.type !== 'invalid') return point;
          if (cells[index]) return cells[index];
          chart.issues.push({ kind: 'data', error: 'An invalid chart cache value was omitted' });
          return { type: 'blank', value: null };
        });
      }
    } else if (formula) {
      try { points = C.referenceValues(formula, book, si, formulas); }
      catch (error) {
        if (error.name !== 'LegacyError') throw error;
        if (kind !== 0) chart.issues.push({ kind: 'data', error });
      }
    } else if (link?.text != null) points = [{ type: 's', value: link.text }];
    else if (kind === 2 && (!link || link.referenceType === 0)) {
      points = Array.from({ length: series.valueCount }, (_, i) => ({ type: 'n', value: i + 1 }));
    }
    return { formula, points, format: link?.format || 0, count: Math.max(points.length, count) };
  };

  C.formula = function (tokens, book, si, formulas) {
    if (book.version !== 'BIFF5/7') return X.formula(tokens, new Uint8Array(), formulas.context(si));
    // BIFF5 chart sources normally contain one 3-D reference or a union of
    // those references. Its direct sheet indices precede the short cell address.
    const r = new B.Reader(tokens), stack = [];
    for (let p = 0; p < r.length;) {
      const raw = r.u8(p++), id = raw >= 32 ? (raw & 31) | 32 : raw;
      if (id === 0x10) {
        if (stack.length < 2) fail('Invalid chart reference union');
        const right = stack.pop(), left = stack.pop();
        stack.push(left + ',' + right);
      } else if (id === 0x15) {
        if (!stack.length) fail('Invalid chart reference parentheses');
        stack[stack.length - 1] = '(' + stack.at(-1) + ')';
      } else if (id === 0x3a || id === 0x3b) {
        const external = r.i16(p), first = r.u16(p + 10), last = r.u16(p + 12);
        if (external >= 0 || first !== last || !book.sheets[first]) fail('External BIFF5 chart data is unavailable');
        p += 14;
        const row1 = r.u16(p) & 16383;
        const row2 = id === 0x3b ? r.u16(p + 2) & 16383 : row1;
        const col1 = r.u8(p + (id === 0x3b ? 4 : 2)), col2 = id === 0x3b ? r.u8(p + 5) : col1;
        p += id === 0x3b ? 6 : 3;
        const ref = (row, col) => X.address(row, col).replace(/([A-Z]+)(\d+)/, '$$$1$$$2');
        stack.push(X.quoteSheet(book.sheets[first].name) + '!' + ref(row1, col1) +
          (row1 === row2 && col1 === col2 ? '' : ':' + ref(row2, col2)));
      } else return X.formula(tokens, new Uint8Array(), formulas.context(si));
    }
    if (stack.length !== 1) fail('Invalid chart reference expression');
    return stack[0];
  };
})(typeof window !== 'undefined' ? window : globalThis);
