/* Word pictures, drawing anchors and text-box stories (MS-DOC 2.8/2.9).
 * Derived from Apache POI hwpf/model/{OfficeArtContent,PICFAndOfficeArtData,
 * types/PICFAbstractType,types/FSPAAbstractType}.java (Apache-2.0).
 * Derived from Apache OpenOffice sw/source/filter/ww8/{ww8graf,ww8graf2,ww8par4}.cxx
 * (Apache-2.0): inline pictures, delayed BLIPs, OLE previews and text-box chains. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    D = B.doc,
    A = B.officeart,
    N = s => s,
    G = (D.draw = {});
  const P = (G.properties = {
    rotation: 4,
    textId: 128,
    textLeft: 129,
    textAnchor: 135,
    textFlow: 136,
    textDirection: 137,
    textBoolean: 191,
    wordArtText: 192,
    wordArtAlign: 194,
    wordArtSize: 195,
    wordArtSpacing: 196,
    wordArtFont: 197,
    wordArtBoolean: 255,
    cropTop: 256,
    cropBottom: 257,
    cropLeft: 258,
    cropRight: 259,
    picture: 260,
    transparentColor: 263,
    contrast: 264,
    brightness: 265,
    gamma: 266,
    pictureBoolean: 319,
    fillPicture: 390,
    fillBoolean: 447,
    name: 896,
    description: 897,
    hyperlink: 898,
    wrapPolygon: 899,
    wrapLeft: 900,
    wrapTop: 901,
    wrapRight: 902,
    wrapBottom: 903,
    horizontalAlign: 911,
    horizontalRelative: 912,
    verticalAlign: 913,
    verticalRelative: 914,
    groupBoolean: 959
  });
  const RECORD = {
    blipStore: 61441,
    shapeContainer: 61444,
    blipStoreEntry: 61447,
    properties: 61451,
    tertiaryProperties: 61730
  };
  const TABLE = { drawingContent: 50, mainAnchors: 40, headerAnchors: 41, textboxes: 56, headerTextboxes: 58 };

  G.get = (shape, id, fallback = 0) => shape.properties.get(id)?.value ?? fallback;
  G.flag = (shape, id, bit, fallback) => {
    const v = G.get(shape, id);
    return v & (1 << (bit + 16)) ? !!(v & (1 << bit)) : fallback;
  };
  G.string = (shape, id) => {
    const bytes = shape.properties.get(id)?.data;
    return bytes ? D.unicode(new B.Reader(bytes), 0, bytes.length & ~1).replace(/\0.*$/s, '') : '';
  };
  const failure = (book, where, e) => {
    if (!['LegacyError', 'CFBError'].includes(e.name)) throw e;
    D.notice(book, 'legacy-word-picture', N('A picture or drawing object could not be kept'), where, true, e.code);
  };
  async function optional(book, where, read) {
    try {
      return await read();
    } catch (e) {
      failure(book, where, e);
      return null;
    }
  }
  function content(book, state) {
    const r = D.table(book, TABLE.drawingContent, 'OfficeArtContent');
    if (!r.length) return;
    const first = A.record(r),
      group = A.records(r.reader(0, first.end));
    for (const rec of group[0].children || [])
      if ([RECORD.properties, RECORD.tertiaryProperties].includes(rec.type))
        D.optional(book, 'Document', N('Some drawing formatting could not be read'), () =>
          A.mergeProperties(state.defaults, A.properties(rec))
        );
    // The host delay stream is WordDocument. AOO also tries Data as a fallback.
    state.store = A.store(group, [book.r.bytes, () => book.cfb.stream('Data')]);
    for (let p = first.end; p < r.length; ) {
      const label = r.u8(p++),
        record = A.record(r, p),
        tree = A.records(r.reader(p, record.end - p));
      p = record.end;
      if (label > 1) {
        D.notice(book, 'legacy-word-drawing-metadata', 'Unknown drawing story', 'Document', false);
        continue;
      }
      const visit = shape => {
        shape.properties = A.mergeProperties(new Map(state.defaults), shape.properties);
        state.shapes[label].set(shape.id, shape);
        shape.story = label;
        shape.children.forEach(visit);
      };
      A.shapes(tree, 'doc').forEach(visit);
    }
  }
  function anchors(book, state, id, story, label) {
    const r = D.table(book, id, 'PlcfSpa');
    if (!r.length) return;
    const plc = D.plc(r, 26),
      range = book.storyRanges[story];
    for (let i = 0; i < plc.n; i++)
      D.optional(book, story, N('A drawing with a damaged anchor could not be kept'), () => {
        const p = plc.at(i),
          cp = range.start + plc.cp[i],
          shape = state.shapes[label].get(r.u32(p));
        if (cp < range.start || cp >= range.end || !shape)
          B.fail('DOC_DRAW_ANCHOR', N('A drawing anchor has no matching shape'));
        const x = r.i32(p + 4) * 635,
          y = r.i32(p + 8) * 635,
          cx = (r.i32(p + 12) - r.i32(p + 4)) * 635,
          cy = (r.i32(p + 16) - r.i32(p + 8)) * 635;
        const node = {
          cp,
          shape,
          rect: { x, y, cx, cy },
          flags: r.u16(p + 20),
          inline: false,
          where: (label ? 'Header or footer' : 'Document') + ', drawing ' + (i + 1)
        };
        const list = state.at.get(cp) || [];
        state.at.set(cp, list);
        list.push(node);
        state.nodes.push(node);
      });
  }
  function textboxes(book, state, id, story, label) {
    const r = D.table(book, id, 'PlcfTxbxTxt');
    if (!r.length) return;
    const plc = D.plc(r, 22),
      base = book.storyRanges[story];
    state.textboxes[label] = Array.from({ length: plc.n }, (_, i) => ({
      index: i,
      start: base.start + plc.cp[i],
      end: base.start + plc.cp[i + 1],
      reusable: !!r.u16(plc.at(i) + 8),
      shapeId: r.u32(plc.at(i) + 14),
      story
    }));
  }
  function inline(book, cp, c) {
    if (c.ole) return objectPreview(book, cp, c);
    book.data ??= new B.Reader(book.cfb.stream('Data') || book.r.bytes, 'Data');
    const at = c.pictureOffset;
    if (at == null || c.data) B.fail('DOC_PICTURE_DATA', N('The picture has no readable image data'));
    const size = book.data.u32(at),
      r = book.data.reader(at, size, 'PICF'),
      header = r.u16(4),
      mm = r.u16(6);
    if (header < 68 || header > size) B.fail('DOC_PICTURE_HEADER', N('The picture header is damaged'));
    if (mm !== 100 && mm !== 102) B.fail('DOC_PICTURE_TYPE', N('This picture format cannot be converted yet'));
    let p = header;
    if (mm === 102) p += 1 + r.u8(p);
    const records = A.records(r.reader(p, r.length - p)),
      shape = A.shapes(records, 'doc')[0];
    if (!shape) B.fail('DOC_PICTURE_SHAPE', N('The picture shape is missing'));
    const entries = records.filter(rec => rec.type === RECORD.blipStoreEntry),
      direct = records.filter(rec => rec.type !== RECORD.shapeContainer && rec.type !== RECORD.blipStoreEntry);
    // Inline stores are private to this PICF. Never fall back to unrelated
    // global BSE indices, even when the recorded blip number exceeds one.
    const store = A.store([{ type: RECORD.blipStore, children: entries }], null);
    if (!store.length && direct.length) store.push({ read: () => A.blip(direct[0]) });
    const cx = Math.round((r.i16(28) * r.u16(32) * 635) / 1000),
      cy = Math.round((r.i16(30) * r.u16(34) * 635) / 1000);
    if (cx <= 0 || cy <= 0) B.fail('DOC_PICTURE_SIZE', N('The picture has no usable size'));
    return {
      cp,
      shape,
      rect: { x: 0, y: 0, cx, cy },
      inline: true,
      store,
      where: 'Document, picture at character ' + (cp + 1),
      borders: [46, 50, 54, 58].map(p => D.border(r.reader(p, 4)))
    };
  }
  function objectPreview(book, cp, c) {
    // Word stores the ObjectPool storage ID in CPicLocation for an OLE run.
    // Keep its existing WMF preview; the embedded application's data is not run.
    const path = 'ObjectPool/_' + (c.pictureOffset | 0) + '/',
      meta = book.cfb.stream(path + '\x03META'),
      pic = book.cfb.stream(path + '\x03PIC');
    if (!meta || !pic) B.fail('DOC_OBJECT_PREVIEW', N('The embedded object has no readable picture'));
    const m = new B.Reader(meta, 'Object preview'),
      r = new B.Reader(pic, 'Object picture size');
    m.check(0, 26);
    r.check(0, 68);
    if ([94, 99].includes(m.u16(0)) || !m.i16(2) || !m.i16(4) || ![1, 2].includes(m.u16(8)) || m.u16(10) !== 9)
      B.fail('DOC_OBJECT_PREVIEW', N('The embedded object picture is damaged'));
    const width = r.i32(20),
      height = r.i32(24),
      sx = r.i32(44),
      sy = r.i32(48),
      crop = [r.i32(52), r.i32(56), r.i32(60), r.i32(64)];
    if (width <= 0 || height <= 0 || sx < 10 || sx > 65536 || sy < 10 || sy > 65536)
      B.fail('DOC_PICTURE_SIZE', N('The picture has no usable size'));
    const cx = Math.round(((width - crop[0] - crop[2]) * sx * 635) / 1000),
      cy = Math.round(((height - crop[1] - crop[3]) * sy * 635) / 1000);
    if (cx <= 0 || cy <= 0) B.fail('DOC_PICTURE_SIZE', N('The picture has no usable size'));
    const properties = new Map([[P.picture, { id: P.picture, value: 1 }]]);
    for (const [i, id] of [P.cropLeft, P.cropTop, P.cropRight, P.cropBottom].entries())
      if (crop[i])
        properties.set(id, { id, value: Math.round((crop[i] * 65536) / (i % 2 ? height : width)) >>> 0 });
    const media = { bytes: m.sub(8, m.length - 8), ext: 'wmf', mime: 'image/x-wmf' },
      shape = { id: 0, type: 75, flags: 0xa00, properties, children: [] };
    return {
      cp,
      shape,
      rect: { x: 0, y: 0, cx, cy },
      inline: true,
      store: [{ read: () => media }],
      where: 'Document, embedded object at character ' + (cp + 1)
    };
  }
  async function images(book, state, shape, store, where, privateStore = false, seen = new Set()) {
    if (seen.has(shape)) B.fail('DOC_DRAW_CYCLE', N('Drawing objects contain a loop'));
    if (seen.size >= B.limits.depth) B.fail('DOC_DRAW_DEPTH', N('Drawing objects are nested too deeply'));
    seen.add(shape);
    shape.images ??= new Map();
    for (const id of [P.picture, P.fillPicture]) {
      const index = G.get(shape, id);
      if (!index || shape.images.has(index)) continue;
      const entry = store[index - 1] || (privateStore && store.length === 1 ? store[0] : null);
      if (!entry) {
        D.notice(book, 'legacy-word-picture', N('A picture or picture fill is missing'), where, true);
        continue;
      }
      const media = await optional(book, where, () => entry.read());
      if (media) shape.images.set(index, media);
    }
    for (const child of shape.children) await images(book, state, child, store, where, privateStore, seen);
    seen.delete(shape);
  }
  G.prepare = async function (book, pkg) {
    const state = (book.draw = {
      pkg,
      at: new Map(),
      nodes: [],
      defaults: new Map(),
      shapes: [new Map(), new Map()],
      textboxes: [[], []],
      store: [],
      media: new Map(),
      relationships: new Map(),
      active: new Set(),
      usedIds: new Set(),
      reservedIds: new Set(),
      serial: 0,
      mediaSerial: 0,
      chainShapes: new Map(),
      emittedText: new Set(),
      handledMarkers: new Set()
    });
    book.drawStats = { pictures: 0, shapes: 0, groups: 0, textboxes: 0, drawingFallbacks: 0 };
    D.optional(book, 'Document', N('Some drawing objects could not be read'), () => content(book, state));
    for (const [id, story, label] of [
      [TABLE.textboxes, 'textboxes', 0],
      [TABLE.headerTextboxes, 'headerTextboxes', 1]
    ])
      D.optional(book, story, N('Some text-box positions could not be read'), () =>
        textboxes(book, state, id, story, label)
      );
    for (const [id, story, label] of [
      [TABLE.mainAnchors, 'main', 0],
      [TABLE.headerAnchors, 'headers', 1]
    ])
      D.optional(book, story, N('Some drawing anchors could not be read'), () =>
        anchors(book, state, id, story, label)
      );
    const forms = new Set();
    for (const f of book.fieldStarts.values()) {
      // The last 0x01 in these instructions points to field data, not content.
      const end = f.separator ?? f.end;
      if (
        end != null &&
        /^\s*(?:FORM(?:TEXT|CHECKBOX|DROPDOWN)|HYPERLINK)\b/i.test(
          book.text.slice(f.start + 1, Math.min(end, f.start + 40))
        )
      ) {
        forms.add(end - 1);
        state.handledMarkers.add(end - 1);
      }
    }
    for (let cp = 0; cp < book.text.length; cp++)
      if (book.text[cp] === '\x01' && !forms.has(cp)) {
        const c = D.optional(
          book,
          'Document',
          N('A picture could not be read'),
          () => D.properties(book, cp, 2).props,
          {}
        );
        if (c.data) continue;
        state.handledMarkers.add(cp);
        let node;
        try {
          node = inline(book, cp, c);
          node.shape.properties = A.mergeProperties(new Map(state.defaults), node.shape.properties);
        } catch (e) {
          if (!['LegacyError', 'CFBError'].includes(e.name)) throw e;
          node = { cp, error: e, where: 'Document, picture at character ' + (cp + 1) };
        }
        // Delay a damaged PICF's notice until its story is actually written.
        // Deleted/reusable text-box data is not visible document content.
        state.at.set(cp, [...(state.at.get(cp) || []), node]);
        state.nodes.push(node);
      }
    const visited = new Set(),
      index = shape => {
        if (visited.has(shape)) return;
        visited.add(shape);
        if (shape.id) state.reservedIds.add(shape.id);
        const label = shape.story || 0,
          id = shape.textbox?.length >= 4 ? new B.Reader(shape.textbox).u32(0) : G.get(shape, P.textId);
        if (id >>> 16) {
          shape.textId = id;
          const key = label + ':' + (id >>> 16),
            list = state.chainShapes.get(key) || [];
          state.chainShapes.set(key, list);
          list.push(shape);
        }
        shape.children.forEach(index);
      };
    for (const [order, node] of state.nodes.entries()) {
      node.order = order;
      if (node.error) continue;
      index(node.shape);
      await optional(book, node.where, () =>
        images(book, state, node.shape, node.store || state.store, node.where, !!node.store)
      );
    }
    for (const list of state.chainShapes.values()) list.sort((a, b) => (a.textId & 65535) - (b.textId & 65535));
  };
  G.textRange = function (book, shape, where) {
    if (!shape.textId) return null;
    const state = book.draw,
      label = shape.story || 0,
      key = label + ':' + (shape.textId >>> 16),
      chain = state.chainShapes.get(key) || [shape];
    const index = (shape.textId >>> 16) - 1,
      table = state.textboxes[label];
    let range = table[index];
    // AOO tolerates a reusable slot by advancing to the next live entry.
    for (let i = index; range?.reusable && i < table.length; ) range = table[++i];
    if (
      !range ||
      range.start < book.storyRanges[range.story].start ||
      range.end > book.storyRanges[range.story].end + 1
    )
      B.fail('DOC_TEXTBOX_RANGE', N('A text box has damaged text positions'));
    const next = chain[chain.indexOf(shape) + 1];
    if (chain[0] !== shape) return { start: range.start, end: range.start, key, next, empty: true };
    let end = Math.min(range.end, book.storyRanges[range.story].end);
    if (end > range.start && book.text[end - 1] === '\r') end--;
    return { ...range, end, key, next, rangeKey: label + ':' + range.index };
  };
  G.finish = function (book) {
    const state = book.draw;
    for (const node of state.nodes)
      if (!node.emitted && !node.error)
        D.notice(
          book,
          'legacy-word-picture',
          N('A picture or drawing outside the readable text could not be kept'),
          node.where,
          true
        );
    for (const label of [0, 1])
      if (book.stories[label ? 'headerTextboxes' : 'textboxes'] && !state.textboxes[label].length)
        D.notice(
          book,
          'legacy-word-textboxes',
          N('Text boxes with missing text positions could not be kept'),
          label ? 'Headers and footers' : 'Document',
          true
        );
    for (const label of [0, 1])
      for (const range of state.textboxes[label])
        if (!range.reusable && !state.emittedText.has(label + ':' + range.index)) {
          const text = book.text.slice(range.start, Math.min(range.end, book.storyRanges[range.story].end)),
            shape = state.shapes[label].get(range.shapeId),
            where = (label ? 'Header or footer' : 'Document') + ', text box ' + (range.index + 1);
          // An OLE preview shape has no ClientTextbox, but still has an FTXBXS
          // slot containing EMBED/LINK's single picture result. Its picture is
          // already written from the shape; only the object functionality is lost.
          if (
            shape?.wordWritten &&
            G.get(shape, P.picture) &&
            /^\x13\s*(?:EMBED|LINK)\b[^\x14]*\x14\x01\x15[\r\x07]*$/.test(text)
          )
            D.notice(book, 'legacy-word-object', N('An embedded object was kept only as a picture'), where, true);
          else if (text.replace(/[\r\x07]/g, ''))
            D.notice(
              book,
              'legacy-word-textboxes',
              N('A text box with no readable drawing could not be kept'),
              where,
              true
            );
        }
  };
})(typeof window !== 'undefined' ? window : globalThis);
