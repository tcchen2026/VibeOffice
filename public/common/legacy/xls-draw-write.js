/* SpreadsheetDrawing pictures, shapes, groups and text from binary Excel.
 * Derived from Apache OpenOffice main/sc/source/filter/excel/xiescher.cxx (Apache-2.0).
 * Derived from Apache POI hssf/usermodel/{HSSFShape,HSSFClientAnchor,HSSFTextbox}.java (Apache-2.0).
 * MS-ODRAW shape properties; output follows Office-authored DrawingML drawings. */
(function (root) {
  'use strict';
  const B = root.L.legacy, X = B.xls, A = B.officeart, S = B.shapes, el = B.el, N = s => s;
  const P = Object.freeze({
    rotation: 4, textLeft: 129, textTop: 130, textRight: 131, textBottom: 132, textWrap: 133,
    textAnchor: 135, textFlow: 136, textBoolean: 191, wordArtText: 192, wordArtAlign: 194,
    wordArtSize: 195, wordArtFont: 197, wordArtBoolean: 255, cropTop: 256, cropBottom: 257,
    cropLeft: 258, cropRight: 259, picture: 260, transparentColor: 263, contrast: 264,
    brightness: 265, gamma: 266, pictureBoolean: 319, vertices: 325, segments: 326,
    connections: 342, fillType: 384, fillPicture: 390,
    shadowBoolean: 575, name: 896, description: 897, groupBoolean: 959
  });
  const OBJECT = Object.freeze({ GROUP: 0, CHART: 5, PICTURE: 8, DROPDOWN: 20, COMMENT: 25 });
  const ns = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const get = (shape, id, fallback = 0) => shape.properties.get(id)?.value ?? fallback;
  const flag = (shape, id, bit, fallback) => {
    const value = get(shape, id);
    return value & 1 << (bit + 16) ? !!(value & 1 << bit) : fallback;
  };
  const string = (shape, id) => {
    const bytes = shape.properties.get(id)?.data;
    return bytes ? B.decode(bytes.subarray(0, bytes.length & ~1), 1200).replace(/\0+$/, '') : '';
  };
  const localError = error => ['LegacyError', 'CFBError'].includes(error.name);

  X.drawingLoss = function (losses, sheet, kind, shape, error) {
    const messages = {
      picture: N('A picture could not be kept'), shape: N('A drawing object could not be kept'),
      format: N('Some drawing formatting could not be kept'), text: N('Some drawing text could not be kept'),
      control: N('A form control was kept only as a drawing object'),
      object: N('An embedded object was kept only as a picture'),
      pictureLink: N('A picture no longer updates from its linked cells'),
      chart: N('A chart could not be kept'), chartData: N('Some chart values could not be kept'),
      chartReference: N('A chart link to its source cells was not kept'),
      chartDrawing: N('Drawing objects inside a chart were not kept'),
      connection: N('A connector is no longer attached to its shapes')
    };
    const name = shape && string(shape, P.name);
    const where = sheet.name + (name ? ', ' + name : shape?.id ? ', object ' + shape.id : '');
    const loss = B.loss('legacy-xls-' + kind, messages[kind] || messages.shape, where, 'drop', where, true);
    loss.detail = { code: error?.code || String(error || ''), shape: shape?.id };
    losses.push(loss);
  };

  function color(styles, index) {
    return styles.palette[index] || (index === 65 ? 'FFFFFF' : '000000');
  }
  X.drawingFont = function (styles, id = 0, tag = 'a:rPr') {
    const font = styles.fonts[styles.fontIndex(id)] || styles.fonts[0];
    const attrs = { sz: Math.max(100, Math.min(400000, Math.round(font.size * 100))),
      b: font.weight >= 700 ? 1 : null, i: font.flags & 2 ? 1 : null,
      strike: font.flags & 8 ? 'sngStrike' : null,
      u: { 1: 'sng', 2: 'dbl', 33: 'sng', 34: 'dbl' }[font.underline] || null,
      baseline: font.script === 1 ? 30000 : font.script === 2 ? -25000 : null };
    const fill = el('a:solidFill', {}, el('a:srgbClr', { val: color(styles, font.color) }));
    const face = { typeface: font.name || 'Arial',
      charset: font.charset ? font.charset > 127 ? font.charset - 256 : font.charset : null };
    return el(tag, attrs, fill + el('a:latin', face) + el('a:ea', face) + el('a:cs', face));
  };

  X.drawingText = function (item, styles, shape) {
    const value = String(item?.value || ''), runs = (item?.runs || []).filter(r => r.at < value.length);
    const boundaries = new Map(runs.map(r => [r.at, r.font]));
    const paragraphs = [];
    let text = '', content = '', font = 0;
    const flush = () => {
      if (text) content += el('a:r', {}, X.drawingFont(styles, font) + el('a:t', {}, B.xml(text)));
      text = '';
    };
    const alignment = ['l', 'l', 'ctr', 'r', 'just', 'dist'][(item?.options >>> 1) & 7] || 'l';
    const paragraph = () => {
      flush();
      paragraphs.push(el('a:p', {}, el('a:pPr', { algn: alignment }, el('a:buNone')) + content +
        X.drawingFont(styles, font, 'a:endParaRPr')));
      content = '';
    };
    for (let i = 0; i < value.length; i++) {
      if (boundaries.has(i)) {
        flush();
        font = boundaries.get(i);
      }
      if (value[i] === '\r' || value[i] === '\n') {
        paragraph();
        if (value[i] === '\r' && value[i + 1] === '\n') i++;
      } else text += value[i];
    }
    paragraph();
    const orientation = item?.orientation || 0;
    const attrs = { wrap: get(shape, P.textWrap) === 2 ? 'none' : 'square',
      lIns: get(shape, P.textLeft, 91440) | 0, tIns: get(shape, P.textTop, 45720) | 0,
      rIns: get(shape, P.textRight, 91440) | 0, bIns: get(shape, P.textBottom, 45720) | 0,
      anchor: ['t', 't', 'ctr', 'b', 'just'][(item?.options >>> 4) & 7] || 't',
      vert: orientation === 1 ? 'wordArtVert' : get(shape, P.textFlow) === 1 ? 'vert' : null,
      rot: orientation === 2 ? 5400000 : orientation === 3 ? -5400000 : null };
    return el('xdr:txBody', {}, el('a:bodyPr', attrs,
      flag(shape, P.textBoolean, 1, false) ? el('a:spAutoFit') : el('a:noAutofit')) +
      el('a:lstStyle') + paragraphs.join(''));
  };

  function pictureEffects(shape, context, report) {
    let out = '';
    const transparent = get(shape, P.transparentColor, 0xffffffff);
    if (transparent !== 0xffffffff) {
      const rgb = S.color(transparent, context, shape);
      out += el('a:clrChange', {}, el('a:clrFrom', {}, el('a:srgbClr', { val: rgb })) +
        el('a:clrTo', {}, el('a:srgbClr', { val: rgb }, el('a:alpha', { val: 0 }))));
    }
    const gain = get(shape, P.contrast, 65536) / 65536;
    const contrast = Math.round((gain > 1 ? 1 - 1 / gain : gain - 1) * 100000);
    const bright = Math.max(-100000, Math.min(100000, Math.round((get(shape, P.brightness) | 0) / 32768 * 100000)));
    if (contrast || bright) out += el('a:lum', { bright, contrast });
    const bits = get(shape, P.pictureBoolean);
    if (bits & 4) out += el('a:grayscl');
    if (bits & 2) out += el('a:biLevel', { thresh: 50000 });
    if (get(shape, P.gamma, 65536) !== 65536) report('format', shape, 'Picture gamma');
    return out;
  }

  X.sheetDrawings = async function (book, si, meta, styles, pkg, part, losses, formulas) {
    const sheet = book.sheets[si], drawing = X.readDrawingRecords(meta);
    const stats = { drawingSource: 0, pictures: 0, shapes: 0, groups: 0, textboxes: 0, charts: 0, connectors: 0 };
    const report = (kind, shape, error) => X.drawingLoss(losses, sheet, kind, shape, error);
    const source = drawing.shapes.flatMap(s => s.flags & 4 ? s.children : [s]);
    for (const object of drawing.objects.values()) {
      if (!object.shape && object.type === OBJECT.CHART && object.anchor) {
        const shape = { id: object.id, type: 75, flags: 0xa00, properties: new Map(),
          children: [], anchor: object.anchor, object };
        object.shape = shape;
        source.push(shape);
      }
    }
    if (!source.length && !drawing.objects.size && sheet.kind !== 2) return { xml: '', stats };
    const name = 'xl/drawings/drawing' + (si + 1) + '.xml', group = X.drawingGroup(book);
    const grid = X.drawingGrid(meta), images = new Map(), media = book.drawingMedia || (book.drawingMedia = []);
    const context = { palette: { ...styles.palette, 64: '000000', 65: 'FFFFFF' },
      scheme: { ...styles.palette, 64: '000000', 65: 'FFFFFF' },
      system: { 5: 'FFFFFF', 8: '000000', 24: 'FFFFE1', 23: '000000' },
      imageRel: index => images.get(index) };
    const sourceIds = new Map(), shapes = [], references = new Set();
    const ignored = shape => shape.object?.type === OBJECT.COMMENT || shape.object?.autoFilter;
    const prepare = list => {
      for (const shape of list) {
        const properties = A.mergeProperties(new Map(group.defaults), shape.properties);
        shape.properties = properties;
        if (!ignored(shape)) {
          shapes.push(shape);
          sourceIds.set(shape.id, shape);
        }
        for (const id of [P.picture, P.fillPicture]) if (get(shape, id)) references.add(get(shape, id));
        prepare(shape.children || []);
      }
    };
    prepare(source);
    stats.drawingSource = shapes.length;
    for (const index of references) {
      try {
        const entry = group.images[index - 1];
        if (!entry) continue;
        const image = await entry.read();
        const crc = root.L.zip.crc32(image.bytes);
        let stored = media.find(m => m.crc === crc && m.ext === image.ext && m.bytes.length === image.bytes.length &&
          m.bytes.every((v, i) => v === image.bytes[i]));
        if (!stored) {
          stored = { ...image, crc, name: 'xl/media/image' + (media.length + 1) + '.' + image.ext };
          pkg.add(stored.name, image.bytes, image.mime);
          media.push(stored);
        }
        images.set(index, pkg.rel(name, 'image', '../media/' + stored.name.split('/').at(-1)));
      } catch (error) {
        if (!localError(error)) throw error;
        // The visible object reports this failure; unused corrupt BLIPs stay silent.
      }
    }
    const ids = new Set(), written = new Set(), connections = [];
    const sheetChart = sheet.kind === 2 ? book.charts[si].find(c => c.sheet) : null;
    const chartRect = sheetChart ? {
      x: 0, y: 0, cx: Math.max(1, Math.round(sheetChart.node.r.u32(8) * 12700 / 65536)),
      cy: Math.max(1, Math.round(sheetChart.node.r.u32(12) * 12700 / 65536))
    } : null;
    const anchorFor = shape => {
      if (!chartRect) return X.drawingAnchor(shape.anchor, grid);
      const a = shape.anchor;
      if (!a || a.kind !== 'cells') B.fail('XLS_DRAWING_ANCHOR', 'Missing chart drawing anchor');
      // Chart hosts use 0..4000 coordinates in the cell-address components.
      // Old non-Escher objects store those coordinates in the offset components.
      const old = shape.object?.modern === false;
      const x = Math.round((old ? a.dx1 : a.col1) * chartRect.cx / 4000);
      const y = Math.round((old ? a.dy1 : a.row1) * chartRect.cy / 4000);
      const right = Math.round((old ? a.dx2 : a.col2) * chartRect.cx / 4000);
      const bottom = Math.round((old ? a.dy2 : a.row2) * chartRect.cy / 4000);
      return { rect: { x: Math.min(x, right), y: Math.min(y, bottom),
        cx: Math.abs(right - x), cy: Math.abs(bottom - y) } };
    };
    const cNvPr = shape => el('xdr:cNvPr', { id: shape.id,
      name: string(shape, P.name) || 'Object ' + (shape.object?.id ?? shape.id),
      descr: string(shape, P.description) || null, hidden: flag(shape, P.groupBoolean, 1, false) ? 1 : null });
    const style = shape => {
      try { return S.style(shape, context); }
      catch (error) {
        if (!localError(error)) throw error;
        report('format', shape, error);
        const properties = new Map(shape.properties);
        properties.set(P.fillType, { value: 0 });
        properties.set(P.shadowBoolean, { value: 0x20000 });
        try { return S.style({ ...shape, properties }, context); }
        catch (error) {
          if (!localError(error)) throw error;
          return el('a:noFill') + el('a:ln', {}, el('a:noFill'));
        }
      }
    };
    const write = (shape, rect, depth = 0) => {
      if (depth >= B.limits.depth) B.fail('XLS_DRAWING_DEPTH', 'Drawing nesting exceeds the limit');
      if (ignored(shape)) return '';
      if (!shape.id || ids.has(shape.id)) B.fail('XLS_DRAWING_ID', 'Invalid or duplicate drawing identity');
      ids.add(shape.id);
      if (shape.object?.textError) report('text', shape, shape.object.textError);
      if (shape.object?.type === OBJECT.CHART) {
        const chart = X.writeObjectChart?.(book, si, shape.object, styles, pkg, name, formulas, losses);
        if (!chart) {
          report('chart', shape, 'Chart data unavailable');
          return '';
        }
        stats.charts++;
        written.add(shape);
        return el('xdr:graphicFrame', { macro: '' }, el('xdr:nvGraphicFramePr', {}, cNvPr(shape) +
          el('xdr:cNvGraphicFramePr')) + el('xdr:xfrm', {}, el('a:off', { x: 0, y: 0 }) +
          el('a:ext', { cx: 0, cy: 0 })) + el('a:graphic', {}, el('a:graphicData',
          { uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart' }, chart)));
      }
      if (shape.flags & 1 || shape.children?.length) {
        const children = [];
        for (const child of shape.children || []) {
          try {
            const a = child.anchor;
            const childRect = a?.kind === 'child' ? { x: a.left, y: a.top,
              cx: a.right - a.left, cy: a.bottom - a.top } : anchorFor(child).rect;
            children.push(write(child, childRect, depth + 1));
          } catch (error) {
            if (!localError(error)) throw error;
            report('shape', child, error);
          }
        }
        stats.groups++;
        written.add(shape);
        return el('xdr:grpSp', {}, el('xdr:nvGrpSpPr', {}, cNvPr(shape) + el('xdr:cNvGrpSpPr')) +
          el('xdr:grpSpPr', {}, S.transform(shape, rect, true)) + children.join(''));
      }
      const picture = get(shape, P.picture), rid = images.get(picture);
      if (picture && !rid) report('picture', shape, 'The picture data could not be read');
      if (rid) {
        if (shape.object?.embedded || shape.object?.linked) report('object', shape, 'The preview remains');
        else if (shape.object?.control) report('control', shape, 'The control preview remains');
        else if (shape.object?.pictureLink) report('pictureLink', shape, 'The saved picture remains');
        const crop = {};
        for (const [key, id] of [['l', P.cropLeft], ['t', P.cropTop], ['r', P.cropRight], ['b', P.cropBottom]]) {
          if (shape.properties.has(id)) crop[key] = Math.round((get(shape, id) | 0) * 100000 / 65536);
        }
        stats.pictures++;
        written.add(shape);
        return el('xdr:pic', {}, el('xdr:nvPicPr', {}, cNvPr(shape) +
          el('xdr:cNvPicPr', {}, el('a:picLocks', { noChangeAspect: 1 }))) + el('xdr:blipFill', {},
          el('a:blip', { 'r:embed': rid }, pictureEffects(shape, context, report)) +
          (Object.keys(crop).length ? el('a:srcRect', crop) : '') + el('a:stretch', {}, el('a:fillRect'))) +
          el('xdr:spPr', {}, S.transform(shape, rect) + el('a:prstGeom', { prst: 'rect' }, el('a:avLst')) +
          style(shape)));
      }
      let text = shape.object?.text, geometry;
      const wordArt = string(shape, P.wordArtText);
      if (wordArt) {
        text = { value: wordArt, runs: [] };
        report('format', shape, 'WordArt was kept as ordinary text');
      }
      const sourceConnector = shape.type >= 32 && shape.type <= 40;
      const connector = sourceConnector && ![P.vertices, P.segments, P.connections].some(id =>
        shape.properties.has(id)) && !text?.value;
      try { geometry = connector ? S.connectorGeometry(shape) : S.geometry(shape); }
      catch (error) {
        if (!localError(error)) throw error;
        report(text?.value ? 'format' : 'shape', shape, error);
        if (!text?.value) return '';
        geometry = el('a:prstGeom', { prst: 'rect' }, el('a:avLst'));
      }
      if (shape.object && ![0, 1, 2, 3, 4, 5, 6, 8, 9, 25, 30].includes(shape.object.type)) {
        report('control', shape, 'The drawing appearance remains; the control action is unavailable');
      }
      let connectionXml = '';
      const rule = drawing.connectors.get(shape.id);
      if (rule && !connector) report('connection', shape, 'The connector remains a drawing shape');
      if (rule && connector) for (const [tag, key] of [['a:stCxn', 'start'], ['a:endCxn', 'end']]) {
        if (!rule[key] || rule[key + 'Site'] === 0xffffffff) continue;
        const target = sourceIds.get(rule[key]);
        let site = null;
        try { if (target) site = S.connectionSite(target, rule[key + 'Site']); }
        catch (error) { if (!localError(error)) throw error; }
        if (site == null) report('connection', shape, 'Missing target or connection site');
        else {
          const xml = el(tag, { id: target.id, idx: site });
          connections.push({ xml, target, shape });
          connectionXml += xml;
        }
      }
      if (text?.formattingError) report('format', shape, 'Text formatting runs could not be read');
      if (text) stats.textboxes++;
      stats.shapes++;
      if (connector) stats.connectors++;
      written.add(shape);
      const nv = connector ? el('xdr:nvCxnSpPr', {}, cNvPr(shape) +
        el('xdr:cNvCxnSpPr', {}, connectionXml)) : el('xdr:nvSpPr', {}, cNvPr(shape) +
        el('xdr:cNvSpPr', { txBox: shape.type === 202 ? 1 : null }));
      return el(connector ? 'xdr:cxnSp' : 'xdr:sp', {}, nv + el('xdr:spPr', {}, S.transform(shape, rect) +
        geometry + style(shape)) + (text || shape.type === 202 ? X.drawingText(text, styles, shape) : ''));
    };
    const output = [];
    for (const shape of source) {
      if (ignored(shape)) continue;
      try {
        const anchor = anchorFor(shape);
        const body = write(shape, anchor.rect);
        if (!body) continue;
        const marker = (tag, values) => el(tag, {}, Object.entries(values).map(([key, value]) =>
          el('xdr:' + key, {}, String(value))).join(''));
        output.push(chartRect ? el('xdr:absoluteAnchor', {}, el('xdr:pos', { x: anchor.rect.x, y: anchor.rect.y }) +
          el('xdr:ext', { cx: anchor.rect.cx, cy: anchor.rect.cy }) + body + el('xdr:clientData')) :
          el('xdr:twoCellAnchor', { editAs: anchor.editAs }, marker('xdr:from', anchor.from) +
            marker('xdr:to', anchor.to) + body + el('xdr:clientData')));
      } catch (error) {
        if (!localError(error)) throw error;
        report(shape.object?.type === OBJECT.CHART ? 'chart' : 'shape', shape, error);
      }
    }
    for (const object of drawing.objects.values()) {
      if (object.type === OBJECT.COMMENT || object.autoFilter || object.shape && written.has(object.shape)) continue;
      if (object.type === OBJECT.GROUP && !object.shape) continue;
      if (!object.shape) report(object.type === OBJECT.CHART ? 'chart' : 'shape', null, 'Object ' + object.id);
    }
    if (drawing.errors.length) report('shape', null, drawing.errors[0]);
    if (sheet.kind === 2) {
      try {
        const body = X.writeObjectChart(book, si, null, styles, pkg, name, formulas, losses);
        if (body) {
          const frame = el('xdr:graphicFrame', { macro: '' }, el('xdr:nvGraphicFramePr', {},
            el('xdr:cNvPr', { id: 2, name: 'Chart 1' }) +
            el('xdr:cNvGraphicFramePr', {}, el('a:graphicFrameLocks', { noGrp: 1 }))) +
            el('xdr:xfrm', {}, el('a:off', { x: 0, y: 0 }) + el('a:ext', { cx: 0, cy: 0 })) +
            el('a:graphic', {}, el('a:graphicData',
              { uri: 'http://schemas.openxmlformats.org/drawingml/2006/chart' }, body)));
          output.unshift(el('xdr:absoluteAnchor', {}, el('xdr:pos', { x: 0, y: 0 }) +
            el('xdr:ext', { cx: chartRect.cx, cy: chartRect.cy }) + frame + el('xdr:clientData')));
          stats.charts++;
        }
      } catch (error) {
        if (!localError(error)) throw error;
        report('chart', null, error);
      }
    }
    if (!output.length && sheet.kind !== 2) return { xml: '', stats };
    let content = output.join('');
    for (const connection of connections) if (!written.has(connection.target)) {
      content = content.replaceAll(connection.xml, '');
      report('connection', connection.shape, 'The target drawing object was not kept');
    }
    pkg.add(name, B.xmlDeclaration + el('xdr:wsDr', { 'xmlns:xdr': ns, 'xmlns:a': B.ns.a,
      'xmlns:r': B.ns.officeRel.slice(0, -1) }, content),
    'application/vnd.openxmlformats-officedocument.drawing+xml');
    return { xml: el('drawing', { 'r:id': pkg.rel(part, 'drawing', '../drawings/drawing' + (si + 1) + '.xml') }),
      stats };
  };
})(typeof window !== 'undefined' ? window : globalThis);
