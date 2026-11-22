/* Native Word drawings and VML shape/text-box groups.
 * Derived from Apache OpenOffice sw/source/filter/ww8/ww8graf.cxx
 * (Apache-2.0): floating placement and text-box content ownership.
 * Derived from Open XML SDK data/schemas/schemas-microsoft-com_office_word.json
 * and schemas-microsoft-com_vml.json (MIT): VML border vocabulary.
 * MS-DOC/MS-ODRAW and ECMA-376; output vocabulary follows Office-authored
 * drawing.docx, VariousPictures.docx and shapes-with-text.docx in POI's corpus. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    D = B.doc,
    G = D.draw,
    S = B.shapes,
    el = B.el,
    N = s => s,
    get = G.get,
    P = G.properties,
    w = (...a) => D.w(...a);
  const ns = {
    v: 'urn:schemas-microsoft-com:vml',
    o: 'urn:schemas-microsoft-com:office:office',
    w10: 'urn:schemas-microsoft-com:office:word',
    wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    a: B.ns.a,
    pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture'
  };
  const point = n => Math.round((n / 12700) * 10000) / 10000;
  function mediaRel(book, media, part) {
    if (!media) return null;
    const s = book.draw;
    let name = s.media.get(media);
    if (!name) {
      name = 'word/media/image' + ++s.mediaSerial + '.' + media.ext;
      s.pkg.add(name, media.bytes, media.mime);
      s.media.set(media, name);
    }
    const key = part + '\0' + name;
    if (!s.relationships.has(key)) s.relationships.set(key, s.pkg.rel(part, 'image', name.slice(5)));
    return s.relationships.get(key);
  }
  function id(book, shape) {
    if (shape.wordId) return shape.wordId;
    const s = book.draw;
    let n = shape.id;
    if (!n || s.usedIds.has(n)) {
      n = 1025;
      while (s.usedIds.has(n) || s.reservedIds.has(n)) n++;
    }
    s.usedIds.add(n);
    return (shape.wordId = '_x0000_s' + n);
  }
  function position(shape, node) {
    // Word's raw values follow AOO ProcessEscherAlign (zero-based), including
    // its default 2 = text. MS-ODRAW's generic enumeration is one-based.
    const flags = node.flags || 0,
      h = get(shape, P.horizontalRelative, (flags >>> 1) & 3),
      v = get(shape, P.verticalRelative, (flags >>> 3) & 3);
    return {
      hAlign:
        ['absolute', 'left', 'center', 'right', 'inside', 'outside'][get(shape, P.horizontalAlign)] || 'absolute',
      vAlign:
        ['absolute', 'top', 'center', 'bottom', 'inside', 'outside'][get(shape, P.verticalAlign)] || 'absolute',
      hRel: ['margin', 'page', 'text', 'char'][h] || 'page',
      vRel: ['margin', 'page', 'text', 'line'][v] || 'page'
    };
  }
  function style(book, shape, rect, node, group) {
    const r = S.anchorRect(shape, rect),
      parts = [
        'position:absolute',
        (group ? 'left:' : 'margin-left:') + (group ? r.x : point(r.x)) + (group ? '' : 'pt'),
        (group ? 'top:' : 'margin-top:') + (group ? r.y : point(r.y)) + (group ? '' : 'pt'),
        'width:' + (group ? r.cx : point(r.cx)) + (group ? '' : 'pt'),
        'height:' + (group ? r.cy : point(r.cy)) + (group ? '' : 'pt')
      ];
    const angle = (get(shape, P.rotation) | 0) / 65536;
    if (angle) parts.push('rotation:' + angle);
    if (shape.flags & 192) parts.push('flip:' + (shape.flags & 64 ? 'x' : '') + (shape.flags & 128 ? 'y' : ''));
    if (G.flag(shape, P.groupBoolean, 1, false)) parts.push('visibility:hidden');
    if (!group && !node.inline) {
      const p = position(shape, node);
      parts.push(
        'z-index:' + (node.flags & 16384 ? -1 : 1) * (251658240 + (node.order || 0)),
        'mso-position-horizontal:' + p.hAlign,
        'mso-position-horizontal-relative:' + p.hRel,
        'mso-position-vertical:' + p.vAlign,
        'mso-position-vertical-relative:' + p.vRel,
        'mso-wrap-style:square'
      );
      for (const [id, side, defaultValue] of [
        [P.wrapLeft, 'left', 114300],
        [P.wrapTop, 'top', 0],
        [P.wrapRight, 'right', 114300],
        [P.wrapBottom, 'bottom', 0]
      ])
        parts.push('mso-wrap-distance-' + side + ':' + point(get(shape, id, defaultValue) | 0) + 'pt');
    }
    const anchor = ['top', 'middle', 'bottom', 'top-center', 'middle-center', 'bottom-center'][
      get(shape, P.textAnchor)
    ];
    if (anchor) parts.push('v-text-anchor:' + anchor);
    if (get(shape, P.textFlow)) parts.push('layout-flow:vertical');
    return (node.inline && !group ? parts.filter(p => !/^position:|^margin-/.test(p)) : parts).join(';');
  }
  function wrap(node) {
    const wr = (node.flags >>> 5) & 15,
      side = ['both', 'left', 'right', 'largest'][(node.flags >>> 9) & 15] || 'both';
    if (node.inline) return '';
    return el('w10:wrap', {
      type: { 0: 'square', 1: 'topAndBottom', 2: 'square', 3: 'none', 4: 'tight', 5: 'through' }[wr] || 'square',
      side
    });
  }
  function textbox(book, shape, part, where) {
    const range = G.textRange(book, shape, where);
    if (!range) return null;
    if (book.draw.active.has(range.key)) B.fail('DOC_TEXTBOX_CYCLE', N('Text boxes contain a loop'));
    book.draw.active.add(range.key);
    try {
      const body = range.empty ? w('p') : D.storyXML(book, range, where + ', text box', { part });
      if (!range.empty) book.draw.emittedText.add(range.rangeKey);
      book.drawStats.textboxes++;
      return { text: w('txbxContent', {}, body), next: range.next ? id(book, range.next) : null };
    } finally {
      book.draw.active.delete(range.key);
    }
  }
  function imageAttributes(shape) {
    const attrs = {};
    for (const [id, name] of [
      [P.cropTop, 'croptop'],
      [P.cropBottom, 'cropbottom'],
      [P.cropLeft, 'cropleft'],
      [P.cropRight, 'cropright']
    ])
      if (shape.properties.has(id)) attrs[name] = (get(shape, id) | 0) / 65536;
    if (shape.properties.has(P.contrast)) attrs.gain = get(shape, P.contrast) / 65536;
    if (shape.properties.has(P.brightness)) attrs.blacklevel = (get(shape, P.brightness) | 0) / 65536;
    if (shape.properties.has(P.gamma)) attrs.gamma = get(shape, P.gamma) / 65536;
    if (shape.properties.has(P.transparentColor) && get(shape, P.transparentColor) !== 0xffffffff)
      attrs.chromakey = '#' + S.color(get(shape, P.transparentColor));
    if (G.flag(shape, P.pictureBoolean, 2, false)) attrs.grayscale = 't';
    if (G.flag(shape, P.pictureBoolean, 1, false)) attrs.bilevel = 't';
    return attrs;
  }
  const borderTypes = [
    'none',
    'single',
    'thick',
    'double',
    'hairline',
    'dot',
    'dash',
    'dotDash',
    'dashDotDot',
    'triple',
    'thinThickSmall',
    'thickThinSmall',
    'thickBetweenThinSmall',
    'thinThick',
    'thickThin',
    'thickBetweenThin',
    'thinThickLarge',
    'thickThinLarge',
    'thickBetweenThinLarge',
    'wave',
    'doubleWave',
    'dashedSmall',
    'dashDotStroked',
    'threeDEmboss',
    'threeDEngrave',
    'outset',
    'inset'
  ];
  function borders(book, node, where) {
    const attributes = {};
    let xml = '';
    for (const [i, b] of (node.borders || []).entries()) {
      if (!b.type || b.type === 255) continue;
      const type = borderTypes[b.type],
        side = ['top', 'left', 'bottom', 'right'][i];
      if (!type) {
        D.notice(
          book,
          'legacy-word-picture-border',
          N('A decorative picture border could not be kept'),
          where,
          true
        );
        continue;
      }
      attributes['o:border' + side + 'color'] = b.color === 'auto' ? 'black' : '#' + b.color;
      xml += el('w10:border' + side, { type, width: Math.max(1, b.size), shadow: b.shadow ? 't' : null });
      if (b.space)
        D.notice(
          book,
          'legacy-word-picture-border-space',
          N('The spacing around a picture border could not be kept'),
          where,
          true
        );
    }
    return { attributes, xml };
  }
  function vector(book, shape, rect, node, part, where, group = false, depth = 0) {
    if (depth >= B.limits.depth) B.fail('DOC_DRAW_DEPTH', N('Drawing objects are nested too deeply'));
    if (rect.cx < 0 || rect.cy < 0) B.fail('DOC_DRAW_SIZE', N('The drawing has invalid dimensions'));
    const shapeId = id(book, shape);
    let css = style(book, shape, rect, node, group);
    if (shape.flags & 1) {
      const c = shape.coord;
      if (!c || c.right === c.left || c.bottom === c.top)
        B.fail('DOC_DRAW_GROUP', N('A drawing group has damaged bounds'));
      let body = '';
      for (const child of shape.children) {
        const a = child.anchor;
        if (!a || a.kind !== 'child') {
          D.notice(book, 'legacy-word-picture', N('A drawing in a group has no usable position'), where, true);
          continue;
        }
        body += D.optional(
          book,
          where,
          N('A drawing in a group could not be kept'),
          () =>
            vector(
              book,
              child,
              { x: a.left, y: a.top, cx: a.right - a.left, cy: a.bottom - a.top },
              node,
              part,
              where,
              true,
              depth + 1
            ),
          ''
        );
      }
      book.drawStats.groups++;
      return el(
        'v:group',
        {
          id: shapeId,
          'o:spid': shapeId,
          style: css,
          coordorigin: c.left + ',' + c.top,
          coordsize: c.right - c.left + ',' + (c.bottom - c.top)
        },
        body + (!group ? wrap(node) : '')
      );
    }
    const range = G.textRange(book, shape, where);
    if (range?.next) css += ';mso-next-textbox:#' + id(book, range.next);
    const wordArt = G.string(shape, P.wordArtText);
    let tail = group ? '' : wrap(node),
      text = range ? () => textbox(book, shape, part, where)?.text : null;
    if (wordArt) {
      // VML's textpath keeps the original OfficeArt WordArt outline and text.
      const font = G.string(shape, P.wordArtFont) || 'Arial',
        f = (bit, fallback = false) => G.flag(shape, P.wordArtBoolean, bit, fallback),
        fontStyle = [
          'font-family:"' + font.replace(/[";{}]/g, '') + '"',
          'font-size:' + get(shape, P.wordArtSize, 36 * 65536) / 65536 + 'pt',
          f(5) ? 'font-weight:bold' : '',
          f(4) ? 'font-style:italic' : '',
          'v-text-kern:' + (f(12) ? 't' : 'f'),
          'v-same-letter-heights:' + (f(7) ? 't' : 'f'),
          'v-rotate-letters:' + (f(13) ? 't' : 'f'),
          'v-text-reverse:' + (f(15) ? 't' : 'f'),
          'v-text-align:' +
            (['left', 'right', 'center', 'letter-justify', 'word-justify', 'stretch-justify'][
              get(shape, P.wordArtAlign, 2)
            ] || 'center'),
          'v-text-spacing:' + get(shape, P.wordArtSpacing, 65536) + 'f',
          'v-text-spacing-mode:' + (f(11) ? 'tracking' : 'tightening')
        ]
          .filter(Boolean)
          .join(';');
      tail =
        el('v:path', { textpathok: 't' }) +
        el('v:textpath', {
          on: f(14, true) ? 't' : 'f',
          fitshape: f(10, ![144, 145, 146, 147].includes(shape.type)) ? 't' : 'f',
          fitpath: f(8) ? 't' : 'f',
          trim: f(9) ? 't' : 'f',
          xscale: f(6) ? 't' : 'f',
          string: wordArt.replace(/\r$/, ''),
          style: fontStyle
        }) +
        tail;
    }
    if (shape.properties.has(P.hyperlink))
      D.notice(book, 'legacy-word-drawing-link', N('A hyperlink on a drawing could not be kept'), where, true);
    const vrect = group
      ? { x: rect.x * 12700, y: rect.y * 12700, cx: rect.cx * 12700, cy: rect.cy * 12700 }
      : rect;
    const border = group ? { attributes: {}, xml: '' } : borders(book, node, where);
    const context = {
      id: shapeId,
      style: css,
      text,
      tail: tail + border.xml,
      imageRel: index => mediaRel(book, shape.images?.get(index), part),
      imageAttributes: imageAttributes(shape),
      attributes: {
        alt: G.string(shape, P.description) || null,
        filled: G.flag(shape, P.fillBoolean, 4, ![20, 32, 33, 34, 35, 36, 37, 38, 39, 40].includes(shape.type))
          ? 't'
          : 'f',
        'o:allowincell': G.flag(shape, P.groupBoolean, 15, true) ? 't' : 'f',
        'o:allowoverlap': G.flag(shape, P.groupBoolean, 5, true) ? 't' : 'f',
        ...border.attributes
      },
      textboxAttributes: { style: G.flag(shape, P.textBoolean, 1, false) ? 'mso-fit-shape-to-text:t' : null }
    };
    let result;
    try {
      result = S.vml(shape, vrect, context);
    } catch (e) {
      if (e.name !== 'LegacyError' || !range) throw e;
      D.notice(
        book,
        'legacy-word-drawing-style',
        N('A text box lost its shape formatting; its text was kept'),
        where,
        true,
        e.code
      );
      const plain = {
        ...shape,
        type: 202,
        properties: new Map([...shape.properties].filter(([id]) => id >= P.textLeft && id <= P.textDirection))
      };
      result = S.vml(plain, vrect, { ...context, tail: group ? '' : wrap(node) });
    }
    if (get(shape, P.picture)) book.drawStats.pictures++;
    else book.drawStats.shapes++;
    shape.wordWritten = true;
    return result;
  }
  function picture(book, node, part) {
    const shape = node.shape,
      media = shape.images?.get(get(shape, P.picture));
    if (!media) B.fail('DOC_PICTURE_DATA', N('The picture data is missing'));
    if (shape.properties.has(P.hyperlink))
      D.notice(
        book,
        'legacy-word-drawing-link',
        N('A hyperlink on a drawing could not be kept'),
        node.where,
        true
      );
    const rect = S.anchorRect(shape, node.rect),
      rid = mediaRel(book, media, part),
      number = ++book.draw.serial,
      name = G.string(shape, P.name) || 'Picture ' + number;
    const crop = {};
    for (const [id, key] of [
      [P.cropTop, 't'],
      [P.cropBottom, 'b'],
      [P.cropLeft, 'l'],
      [P.cropRight, 'r']
    ])
      if (shape.properties.has(id)) crop[key] = Math.round(((get(shape, id) | 0) * 100000) / 65536);
    let effects = '';
    if (G.flag(shape, P.pictureBoolean, 2, false)) effects += el('a:grayscl');
    if (G.flag(shape, P.pictureBoolean, 1, false)) effects += el('a:biLevel', { thresh: 50000 });
    if (
      (shape.properties.has(P.contrast) && get(shape, P.contrast) !== 65536) ||
      get(shape, P.brightness) ||
      (shape.properties.has(P.gamma) && get(shape, P.gamma) !== 65536)
    )
      D.notice(
        book,
        'legacy-word-picture-adjustments',
        N('Some picture color adjustments could not be kept'),
        node.where,
        true
      );
    const transform = S.transform(shape, { x: 0, y: 0, cx: node.rect.cx, cy: node.rect.cy }, false, true),
      style = D.optional(
        book,
        node.where,
        N('A picture lost its border or shadow formatting'),
        () =>
          S.style(shape, { components: true, imageRel: index => mediaRel(book, shape.images?.get(index), part) }),
        { line: el('a:ln', {}, el('a:noFill')), effects: '' }
      );
    const pic = el(
      'pic:pic',
      {},
      el(
        'pic:nvPicPr',
        {},
        el('pic:cNvPr', { id: number, name }) + el('pic:cNvPicPr', {}, el('a:picLocks', { noChangeAspect: 1 }))
      ) +
        el(
          'pic:blipFill',
          {},
          el('a:blip', { 'r:embed': rid }, effects || undefined) +
            (Object.keys(crop).length ? el('a:srcRect', crop) : '') +
            el('a:stretch', {}, el('a:fillRect'))
        ) +
        el('pic:spPr', {}, transform + S.geometry(shape) + style.line + style.effects)
    );
    const common =
      el('wp:extent', { cx: Math.round(rect.cx), cy: Math.round(rect.cy) }) +
      el('wp:effectExtent', { l: 0, t: 0, r: 0, b: 0 });
    const end =
      el('wp:docPr', { id: number, name, descr: G.string(shape, P.description) || null }) +
      el('wp:cNvGraphicFramePr', {}, el('a:graphicFrameLocks', { noChangeAspect: 1 })) +
      el('a:graphic', {}, el('a:graphicData', { uri: ns.pic }, pic));
    let xml;
    if (node.inline) xml = el('wp:inline', { distT: 0, distB: 0, distL: 0, distR: 0 }, common + end);
    else {
      const p = position(shape, node),
        flags = node.flags,
        wr = (flags >>> 5) & 15;
      const pos = (name, rel, align, offset) =>
        el(
          'wp:' + name,
          { relativeFrom: rel },
          align === 'absolute' ? el('wp:posOffset', {}, Math.round(offset)) : el('wp:align', {}, align)
        );
      const polygon = B.officeart.array(shape.properties.get(P.wrapPolygon)),
        points = polygon.map(r => (r.length === 4 ? [r.i16(0), r.i16(2)] : [r.i32(0), r.i32(4)]));
      let wrapping =
        wr === 3
          ? el('wp:wrapNone')
          : wr === 1
            ? el('wp:wrapTopAndBottom')
            : el('wp:wrapSquare', {
                wrapText: ['bothSides', 'left', 'right', 'largest'][(flags >>> 9) & 15] || 'bothSides'
              });
      if ((wr === 4 || wr === 5) && points.length >= 3)
        wrapping = el(
          wr === 4 ? 'wp:wrapTight' : 'wp:wrapThrough',
          { wrapText: ['bothSides', 'left', 'right', 'largest'][(flags >>> 9) & 15] || 'bothSides' },
          el(
            'wp:wrapPolygon',
            { edited: 1 },
            points.map(([x, y], i) => el(i ? 'wp:lineTo' : 'wp:start', { x, y })).join('')
          )
        );
      const attrs = {
        distT: get(shape, P.wrapTop),
        distB: get(shape, P.wrapBottom),
        distL: get(shape, P.wrapLeft, 114300),
        distR: get(shape, P.wrapRight, 114300),
        simplePos: 0,
        relativeHeight: 251658240 + node.order,
        behindDoc: flags & 16384 ? 1 : 0,
        locked: flags & 32768 ? 1 : 0,
        layoutInCell: G.flag(shape, P.groupBoolean, 15, true) ? 1 : 0,
        allowOverlap: G.flag(shape, P.groupBoolean, 5, true) ? 1 : 0
      };
      xml = el(
        'wp:anchor',
        attrs,
        el('wp:simplePos', { x: 0, y: 0 }) +
          pos('positionH', { text: 'column', char: 'character' }[p.hRel] || p.hRel, p.hAlign, rect.x) +
          pos('positionV', p.vRel === 'text' ? 'paragraph' : p.vRel, p.vAlign, rect.y) +
          common +
          wrapping +
          end
      );
    }
    book.drawStats.pictures++;
    shape.wordWritten = true;
    return w('drawing', Object.fromEntries(['wp', 'a', 'pic'].map(k => ['xmlns:' + k, ns[k]])), xml);
  }
  G.at = function (book, cp, part, where) {
    let out = '';
    for (const node of book.draw?.at.get(cp) || []) {
      if (node.emitted) continue;
      node.emitted = true;
      const result = D.optional(
        book,
        where,
        N('A picture or drawing object could not be kept'),
        () => {
          if (node.error) throw node.error;
          if (node.rect.cx < 0 || node.rect.cy < 0)
            B.fail('DOC_DRAW_SIZE', N('The drawing has invalid dimensions'));
          const adjusted =
            (node.shape.properties.has(P.transparentColor) &&
              get(node.shape, P.transparentColor) !== 0xffffffff) ||
            (node.shape.properties.has(P.contrast) && get(node.shape, P.contrast) !== 65536) ||
            get(node.shape, P.brightness) ||
            (node.shape.properties.has(P.gamma) && get(node.shape, P.gamma) !== 65536);
          if (
            get(node.shape, P.picture) &&
            !node.shape.textId &&
            !node.shape.children.length &&
            !adjusted &&
            !node.borders?.some(b => b.type && b.type !== 255)
          )
            return picture(book, node, part);
          return w(
            'pict',
            Object.fromEntries(['v', 'o', 'w10'].map(k => ['xmlns:' + k, ns[k]])),
            vector(book, node.shape, node.rect, node, part, where)
          );
        },
        ''
      );
      if (!result) book.drawStats.drawingFallbacks++;
      out += result;
    }
    return out;
  };
})(typeof window !== 'undefined' ? window : globalThis);
