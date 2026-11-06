/* Binary PowerPoint drawings and media.
 * Derived from Apache POI hslf/usermodel/{HSLFShape,HSLFGroupShape,
 * HSLFPictureShape,HSLFConnectorShape}.java (Apache-2.0).
 * Derived from Apache OpenOffice filter/source/msfilter/{svdfppt,msdffimp}.cxx
 * (Apache-2.0): anchors, solver connections and local drawing recovery.
 * MS-ODRAW 2.2; MS-PPT 2.7; PresentationML uses the source shape identities. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    P = B.ppt,
    T = P.text,
    A = B.officeart,
    S = B.shapes,
    el = B.el,
    N = s => s,
    D = (P.drawings = {});
  const emu = n => Math.round(n * 1587.5),
    get = (s, id, fallback = 0) => s.properties.get(id)?.value ?? fallback;
  D.pages = book =>
    [...book.masters, ...book.slides, ...book.notes, book.notesMaster, book.handoutMaster].filter(Boolean);
  D.string = (shape, id) => {
    const data = shape.properties.get(id)?.data;
    return data ? B.decode(data.subarray(0, data.length & ~1), 1200).replace(/\0+$/, '') : '';
  };
  const messages = {
    picture: N('Some pictures could not be kept'),
    'picture-adjustments': N('Some picture color adjustments could not be kept'),
    shape: N('Some drawing objects could not be kept'),
    format: N('Some drawing formatting could not be kept'),
    table: N('A table was kept as separate drawing objects'),
    link: N('Some hyperlinks or actions could not be kept'),
    background: N('The background fill could not be kept'),
    header: N('Some headers and footers could not be kept'),
    transition: N('Some slide transition settings could not be kept'),
    object: N('An embedded object was kept only as a picture'),
    sound: N('Some sounds could not be kept'),
    connection: N('Some connectors are no longer attached to their shapes')
  };
  D.report = function (book, page, kind, shape, reason) {
    page.drawingIssues ||= new Map();
    let loss = page.drawingIssues.get(kind);
    if (!loss) {
      loss = B.loss(
        kind === 'background' ? 'legacy-slide-background' : 'legacy-presentation-' + kind,
        messages[kind] || messages.shape,
        page.where,
        'drop',
        page.where,
        true
      );
      loss.detail = [];
      book.losses.push(loss);
      page.drawingIssues.set(kind, loss);
    }
    if (loss.detail.length < 12)
      loss.detail.push({
        shape: shape?.id,
        name: shape ? D.string(shape, 896) || undefined : undefined,
        reason: String(reason?.code || reason || '')
      });
  };
  D.relative = function (from, to) {
    const a = from.split('/');
    a.pop();
    const b = to.split('/');
    while (a.length && b.length && a[0] === b[0]) {
      a.shift();
      b.shift();
    }
    return '../'.repeat(a.length) + b.join('/');
  };
  D.mediaPool = function (book, pkg) {
    const buckets = new Map(),
      rels = new Map();
    let number = 0;
    book.addImage = image => {
      const key = image.mime + ':' + image.bytes.length + ':' + root.L.zip.crc32(image.bytes),
        list = buckets.get(key) || [];
      for (const old of list) if (old.bytes.every((v, i) => v === image.bytes[i])) return old.name;
      let name;
      do {
        name = 'ppt/media/image' + ++number + '.' + image.ext;
      } while (pkg.parts.has(name));
      pkg.add(name, image.bytes, image.mime);
      list.push({ name, bytes: image.bytes });
      buckets.set(key, list);
      return name;
    };
    book.mediaRel = (name, from = book.textPart) => {
      if (!name || !from) return null;
      const key = from + ':' + name;
      if (!rels.has(key)) rels.set(key, pkg.rel(from, 'image', D.relative(from, name)));
      return rels.get(key);
    };
  };
  D.prepare = async function (book, cfb, pkg) {
    D.mediaPool(book, pkg);
    book.drawingStats = {
      sourceShapes: 0,
      sourcePictures: 0,
      sourceGroups: 0,
      sourceTables: 0,
      wordArt: 0,
      drawingShapes: 0,
      pictures: 0,
      groups: 0,
      connectors: 0,
      tables: 0,
      images: 0,
      backgrounds: 0,
      hyperlinks: 0,
      transitions: 0
    };
    const references = new Set(),
      images = new Map(),
      pages = D.pages(book);
    const visit = (shapes, depth = 0) => {
      if (depth > B.limits.depth)
        B.fail('PPT_DRAWING_LIMIT', N('Presentation drawing nesting exceeds the limit'));
      for (const s of shapes) {
        book.drawingStats.sourceShapes++;
        book.drawingStats.sourcePictures += +!!get(s, 260);
        book.drawingStats.sourceGroups += +!!(s.flags & 1);
        book.drawingStats.sourceTables += +!!(get(s, 927) & 1);
        if (get(s, 260)) references.add(get(s, 260));
        if ([1, 2, 3].includes(get(s, 384)) && get(s, 390)) references.add(get(s, 390));
        visit(s.children || [], depth + 1);
      }
    };
    for (const page of pages) {
      visit(page.textShapes || []);
      if (page.backgroundShape && [1, 2, 3].includes(get(page.backgroundShape, 384)))
        references.add(get(page.backgroundShape, 390));
      page.connectorRules = new Map();
      for (const drawing of page.children.filter(r => r.type === 1036))
        P.walk(
          drawing,
          r => {
            if (r.type === 0xf012)
              try {
                r.r.check(0, 24);
                page.connectorRules.set(r.r.u32(12), {
                  start: r.r.u32(4),
                  end: r.r.u32(8),
                  startSite: r.r.u32(16),
                  endSite: r.r.u32(20)
                });
              } catch (e) {
                if (e.name !== 'LegacyError') throw e;
                D.report(book, page, 'format', null, e);
              }
          },
          e => D.report(book, page, 'shape', null, e)
        );
    }
    let store = [],
      delayed;
    if (references.size) {
      try {
        delayed = cfb.stream('Pictures');
      } catch (e) {
        if (!['CFBError', 'LegacyError'].includes(e.name)) throw e;
      }
      for (const record of book.documentRecords.filter(r => r.type === 1035))
        try {
          store.push(...A.store(A.records(record.r), delayed));
        } catch (e) {
          if (e.name !== 'LegacyError') throw e;
        }
      for (const index of references) {
        const entry = store[index - 1];
        if (!entry) continue;
        try {
          images.set(index, book.addImage(await entry.read()));
        } catch (e) {
          if (!['CFBError', 'LegacyError'].includes(e.name)) throw e;
        }
      }
    }
    book.drawingStats.images = images.size;
    book.drawingImageRel = index => book.mediaRel(images.get(index));
    P.tables?.prepare(book);
    return book;
  };
  D.rect = function (shape, book, page) {
    let s = shape;
    if (!s.anchor && s.ph) {
      let parent = page.kind === 'notes' ? book.notesMaster : page.parent;
      const seen = new Set();
      while (parent && !seen.has(parent)) {
        seen.add(parent);
        const match = parent.textAll?.find(x => x.ph?.type === s.ph.type && x.anchor);
        if (match) {
          s = match;
          break;
        }
        parent = parent.parent;
      }
    }
    if (!s.anchor) return null;
    const scale = s.anchor.kind === 'child' ? n => n : emu,
      a = s.anchor;
    return { x: scale(a.left), y: scale(a.top), cx: scale(a.right - a.left), cy: scale(a.bottom - a.top) };
  };
  D.style = function (shape, book, page) {
    const context = { scheme: page.scheme, imageRel: book.drawingImageRel };
    try {
      return S.style(shape, context);
    } catch (e) {
      if (e.name !== 'LegacyError') throw e;
      D.report(book, page, 'format', shape, e);
      // Keep text and geometry when a fill/effect fails independently.
      const properties = new Map(shape.properties);
      properties.set(384, { value: 0 });
      properties.set(575, { value: 0x20000 });
      try {
        return S.style({ ...shape, properties }, context);
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        return el('a:noFill') + el('a:ln', {}, el('a:noFill'));
      }
    }
  };
  D.background = function (book, page) {
    if (!page.backgroundShape)
      return el(
        'p:bg',
        {},
        el('p:bgPr', {}, el('a:solidFill', {}, el('a:schemeClr', { val: 'bg1' })) + el('a:effectLst'))
      );
    try {
      const parts = S.style(
        { ...page.backgroundShape, type: 1 },
        { scheme: page.backgroundScheme || page.scheme, imageRel: book.drawingImageRel, components: true }
      );
      page.backgroundWritten = true;
      book.drawingStats.backgrounds++;
      return el('p:bg', {}, el('p:bgPr', {}, parts.fill + (parts.effects || el('a:effectLst'))));
    } catch (e) {
      if (e.name !== 'LegacyError') throw e;
      D.report(book, page, 'background', null, e);
      return el(
        'p:bg',
        {},
        el('p:bgPr', {}, el('a:solidFill', {}, el('a:schemeClr', { val: 'bg1' })) + el('a:effectLst'))
      );
    }
  };
  D.wordArt = function (book, page, s) {
    const text = D.string(s, 192);
    if (!text) return null;
    // Legacy WordArt keeps its letters in gtextUNICODE, outside TextCharsAtom.
    // Plain Text is the measured corpus form. Other envelopes keep readable
    // editable text and report the lost warp rather than silently lose letters.
    if (s.type !== 136) D.report(book, page, 'format', s, 'WordArt text envelope ' + s.type);
    const bits = get(s, 255),
      face = D.string(s, 197) || book.fonts.get(0)?.name || 'Arial',
      size = Math.max(100, Math.min(400000, Math.round((get(s, 195, 36 * 65536) / 65536) * 100)));
    let style;
    try {
      style = S.style(s, { scheme: page.scheme, imageRel: book.drawingImageRel, components: true });
    } catch (e) {
      if (e.name !== 'LegacyError') throw e;
      D.report(book, page, 'format', s, e);
      style = { fill: el('a:solidFill', {}, el('a:srgbClr', { val: '000000' })), line: '', effects: '' };
    }
    const attrs = {
      sz: size,
      b: bits & 32 ? 1 : null,
      i: bits & 16 ? 1 : null,
      u: bits & 8 ? 'sng' : null,
      strike: bits & 1 ? 'sngStrike' : null,
      kern: bits & 4096 ? 0 : null
    };
    // AOO ReadObjText treats the last CR/LF as the paragraph terminator, not
    // another empty warped line (which would squash all the visible letters).
    const rpr = el(
        'a:rPr',
        attrs,
        style.line +
          style.fill +
          style.effects +
          el('a:latin', { typeface: face }) +
          el('a:ea', { typeface: face }) +
          el('a:cs', { typeface: face })
      ),
      paragraphs = text
        .replace(/\r\n|\n\r|\n/g, '\r')
        .replace(/\r$/, '')
        .split('\r');
    const body = paragraphs
      .map(t =>
        el(
          'a:p',
          {},
          el(
            'a:pPr',
            {
              algn: ['ctr', 'l', 'r', 'just', 'dist', 'just'][get(s, 194)] || 'ctr',
              rtl: /[\u0590-\u08ff]/.test(t) ? 1 : null
            },
            el('a:buNone')
          ) + el('a:r', {}, rpr + el('a:t', {}, B.xml(t)))
        )
      )
      .join('');
    book.drawingStats.wordArt++;
    book.textStats.textShapes++;
    book.textStats.textCharacters += text.length;
    book.textStats.paragraphs += paragraphs.length;
    s.textWritten = true;
    return el(
      'p:txBody',
      {},
      el(
        'a:bodyPr',
        {
          lIns: 0,
          tIns: 0,
          rIns: 0,
          bIns: 0,
          wrap: 'none',
          anchor: 'ctr',
          fromWordArt: 1,
          vert: bits & 8192 ? 'vert' : null
        },
        el('a:prstTxWarp', { prst: 'textPlain' }, el('a:avLst')) + el('a:noAutofit')
      ) +
        el('a:lstStyle') +
        body
    );
  };
  D.connectorGeometry = S.connectorGeometry;
  D.tree = function (book, page, { layout, shapes, usedIds = [], fragment = false } = {}) {
    if (!page) return P.emptyShapeTree();
    const source = shapes || (layout ? layout.textPlaceholders : page.textShapes) || [],
      used = new Set([1, ...usedIds]),
      reserved = new Set(),
      ids = new Map();
    let next = 2;
    const collect = (list, depth = 0) => {
      if (depth > B.limits.depth)
        B.fail('PPT_DRAWING_LIMIT', N('Presentation drawing nesting exceeds the limit'));
      for (const s of list) {
        if (s.id > 1 && s.id <= 0xffffffff) reserved.add(s.id);
        collect(s.children || [], depth + 1);
      }
    };
    collect(source);
    const assign = list => {
      for (const s of list) {
        let id = s.id;
        if (!(id > 1 && id <= 0xffffffff) || used.has(id)) {
          while (used.has(next) || reserved.has(next)) next++;
          id = next++;
        }
        used.add(id);
        ids.set(s, id);
        assign(s.children || []);
      }
    };
    assign(source);
    const bySource = new Map(),
      sourceShape = new Map();
    for (const [s, id] of ids)
      if (!bySource.has(s.id)) {
        bySource.set(s.id, id);
        sourceShape.set(s.id, s);
      }
    for (const [s, id] of ids)
      if (s.tableData && !get(s, 4) && !(s.flags & 192))
        for (const child of s.children || []) bySource.delete(child.id);
    if (!layout && !fragment) page.timingShapes = new Map();
    const cNvPr = s => {
      if (!layout && !fragment && !page.timingShapes.has(s.id))
        page.timingShapes.set(s.id, { id: ids.get(s), shape: s });
      return el(
        'p:cNvPr',
        {
          id: ids.get(s),
          name:
            D.string(s, 896) || (s.flags & 1 ? 'Group ' : get(s, 260) ? 'Picture ' : 'Shape ') + ids.get(s),
          descr: D.string(s, 897) || null,
          hidden: (get(s, 959) & 0x20002) === 0x20002 ? 1 : null
        },
        P.links?.shape(book, page, s) || undefined
      );
    };
    const nvPr = s =>
      el(
        'p:nvPr',
        {},
        (s.ph ? el('p:ph', { type: s.ph.type, idx: s.ph.idx, sz: s.ph.sz, orient: s.ph.orient }) : '') +
          (!layout ? P.media?.shape(book, page, s) || '' : '')
      );
    const write = (s, depth = 0) => {
      if (depth > B.limits.depth)
        B.fail('PPT_DRAWING_LIMIT', N('Presentation drawing nesting exceeds the limit'));
      if (s.headerHidden) {
        s.drawingWritten = true;
        return '';
      }
      let rect = D.rect(s, book, page);
      if (!rect && s.coord)
        rect = {
          x: emu(s.coord.left),
          y: emu(s.coord.top),
          cx: emu(s.coord.right - s.coord.left),
          cy: emu(s.coord.bottom - s.coord.top)
        };
      if (!rect) {
        D.report(book, page, 'shape', s, 'Missing anchor');
        return '';
      }
      // Binary object placeholders can be groups (notably old tables).
      // PowerPoint rejects p:ph on p:grpSp. Its layout slot is an ordinary
      // empty shape; the actual grouped content keeps its own coordinates.
      if (layout && s.ph && (s.flags & 1 || s.children?.length))
        return el(
          'p:sp',
          {},
          el('p:nvSpPr', {}, cNvPr(s) + el('p:cNvSpPr', {}, el('a:spLocks', { noGrp: 1 })) + nvPr(s)) +
            el(
              'p:spPr',
              {},
              S.transform(s, rect) +
                el('a:prstGeom', { prst: 'rect' }, el('a:avLst')) +
                el('a:noFill') +
                el('a:ln', {}, el('a:noFill'))
            ) +
            T.body(s, book, page, true)
        );
      if (get(s, 927) & 1 && P.tables && !layout)
        try {
          const table = P.tables.write(book, page, s, rect, ids.get(s), cNvPr(s));
          if (table) {
            s.drawingWritten = true;
            book.drawingStats.tables++;
            return table;
          }
        } catch (e) {
          if (e.name !== 'LegacyError') throw e;
          D.report(book, page, 'table', s, e);
        }
      if (s.flags & 1 || s.children?.length) {
        const children = (s.children || []).map(c => write(c, depth + 1)).join('');
        s.drawingWritten = true;
        book.drawingStats.groups++;
        return el(
          'p:grpSp',
          {},
          el('p:nvGrpSpPr', {}, cNvPr(s) + el('p:cNvGrpSpPr') + el('p:nvPr')) +
            el('p:grpSpPr', {}, S.transform(s, rect, true)) +
            children
        );
      }
      const picture = get(s, 260),
        rid = picture ? book.drawingImageRel(picture) : null;
      if (picture && !rid) D.report(book, page, 'picture', s, 'Unavailable picture ' + picture);
      if (rid && !layout) {
        for (const [id, fallback, name] of [
          [263, null, 'Transparent color'],
          [264, 65536, 'Contrast'],
          [265, 0, 'Brightness'],
          [266, 65536, 'Gamma']
        ])
          if (s.properties.has(id) && get(s, id) !== fallback)
            D.report(book, page, 'picture-adjustments', s, name);
        if (get(s, 319) & 6) D.report(book, page, 'picture-adjustments', s, 'Grayscale or black and white');
        const crop = {};
        for (const [key, id] of [
          ['l', 258],
          ['t', 256],
          ['r', 259],
          ['b', 257]
        ])
          if (s.properties.has(id)) crop[key] = Math.round(((get(s, id) | 0) * 100000) / 65536);
        const blip = el('a:blip', { 'r:embed': rid });
        const style = D.style(s, book, page);
        s.drawingWritten = true;
        book.drawingStats.pictures++;
        const client = s.records?.find(r => r.type === 0xf011);
        if (
          client && !s.media &&
          P.children(client, e => D.report(book, page, 'object', s, e)).some(r => r.type === 3009)
        )
          D.report(book, page, 'object', s, 'The object preview was retained');
        return el(
          'p:pic',
          {},
          el(
            'p:nvPicPr',
            {},
            cNvPr(s) + el('p:cNvPicPr', {}, el('a:picLocks', { noChangeAspect: 1 })) + nvPr(s)
          ) +
            el(
              'p:blipFill',
              {},
              blip +
                (Object.keys(crop).length ? el('a:srcRect', crop) : '') +
                el('a:stretch', {}, el('a:fillRect'))
            ) +
            el('p:spPr', {}, S.transform(s, rect) + el('a:prstGeom', { prst: 'rect' }, el('a:avLst')) + style)
        );
      }
      const wordArt = !layout && D.wordArt(book, page, s);
      if (wordArt) {
        s.drawingWritten = true;
        book.drawingStats.drawingShapes++;
        return el(
          'p:sp',
          {},
          el('p:nvSpPr', {}, cNvPr(s) + el('p:cNvSpPr') + nvPr(s)) +
            el(
              'p:spPr',
              {},
              S.transform(s, rect) +
                el('a:prstGeom', { prst: 'rect' }, el('a:avLst')) +
                el('a:noFill') +
                el('a:ln', {}, el('a:noFill'))
            ) +
            wordArt
        );
      }
      const sourceConnector = s.type >= 32 && s.type <= 40,
        connector = sourceConnector && ![325, 326, 342].some(id => s.properties.has(id));
      if (sourceConnector && !connector)
        D.report(book, page, 'connection', s, 'The custom connector was kept as a drawing shape');
      let geometry;
      try {
        geometry = connector ? D.connectorGeometry(s) : S.geometry(s);
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        D.report(book, page, 'shape', s, e);
        if (!s.text && !s.ph) return '';
        geometry = el('a:prstGeom', { prst: 'rect' }, el('a:avLst'));
      }
      const rule = page.connectorRules?.get(s.id);
      let cnx = '';
      if (connector && rule)
        for (const [tag, key, site] of [
          ['a:stCxn', 'start', 'startSite'],
          ['a:endCxn', 'end', 'endSite']
        ]) {
          const id = bySource.get(rule[key]);
          if (rule[key] && rule[site] !== 0xffffffff) {
            const target = sourceShape.get(rule[key]);
            let idx = rule[site],
              valid = !!id && !!D.rect(target, book, page) && !target.headerHidden;
            if (valid) {
              try {
                idx = S.connectionSite(target, idx);
                valid = idx != null;
              } catch (error) {
                if (error.name !== 'LegacyError') throw error;
                valid = false;
              }
            }
            if (valid) cnx += el(tag, { id, idx });
            else
              D.report(
                book,
                page,
                'connection',
                s,
                'Missing shape or connection site ' + rule[key] + ':' + rule[site]
              );
          }
        }
      const nv = connector
        ? el('p:nvCxnSpPr', {}, cNvPr(s) + el('p:cNvCxnSpPr', {}, cnx) + nvPr(s))
        : el(
            'p:nvSpPr',
            {},
            cNvPr(s) +
              el(
                'p:cNvSpPr',
                { txBox: s.type === 202 ? 1 : null },
                s.ph ? el('a:spLocks', { noGrp: 1 }) : ''
              ) +
              nvPr(s)
          );
      // An empty binary textbox still needs p:txBody. Office opens a master
      // textbox without it, but can render every dependent slide as blank.
      const text = !connector && (s.text || s.ph || s.type === 202) ? T.body(s, book, page, !!layout) : '';
      s.drawingWritten = true;
      s.frameWritten = true;
      book.drawingStats.drawingShapes++;
      if (connector) book.drawingStats.connectors++;
      return el(
        connector ? 'p:cxnSp' : 'p:sp',
        {},
        nv + el('p:spPr', {}, S.transform(s, rect) + geometry + D.style(s, book, page)) + text
      );
    };
    const content = source.map(s => write(s)).join('');
    return fragment ? content : P.emptyShapeTree().replace('</p:spTree>', () => content + '</p:spTree>');
  };
})(typeof window !== 'undefined' ? window : globalThis);
