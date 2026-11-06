/* Derived from Apache OpenOffice sc/source/filter/excel/xistyle.cxx and
 * Derived from Apache POI hssf/record/{Font,Format,ExtendedFormat,Style,Palette}Record.java
 * (Apache-2.0). BIFF5/8 styles -> SpreadsheetML, independent of any app model. */
/* Font child order: Derived from Open XML SDK v3.3.0 schema data (MIT). */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    X = B.xls,
    el = B.el,
    N = s => s;
  const patterns = [
    'none',
    'solid',
    'mediumGray',
    'darkGray',
    'lightGray',
    'darkHorizontal',
    'darkVertical',
    'darkDown',
    'darkUp',
    'darkGrid',
    'darkTrellis',
    'lightHorizontal',
    'lightVertical',
    'lightDown',
    'lightUp',
    'lightGrid',
    'lightTrellis',
    'gray125',
    'gray0625'
  ];
  const lines = [
    '',
    'thin',
    'medium',
    'dashed',
    'dotted',
    'thick',
    'double',
    'hair',
    'mediumDashed',
    'dashDot',
    'mediumDashDot',
    'dashDotDot',
    'mediumDashDotDot',
    'slantDashDot'
  ];
  const hAlign = ['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'],
    vAlign = ['top', 'center', 'bottom', 'justify', 'distributed'];
  const builtins = [
    'Normal',
    'RowLevel_',
    'ColLevel_',
    'Comma',
    'Currency',
    'Percent',
    'Comma [0]',
    'Currency [0]',
    'Hyperlink',
    'Followed Hyperlink'
  ];
  const defaultPalette = (
    '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 ' +
    'C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF ' +
    '800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC ' +
    '99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333'
  )
    .split(' ')
    .map(c => 'FF' + c);
  const color = (tag, index) =>
    el(
      tag,
      typeof index === 'object'
        ? index
        : index === 0x7fff
          ? { auto: 1 }
          : { indexed: index <= 65 ? index : 64 }
    );
  X.stylePatterns = patterns;
  X.borderLines = lines;
  X.styleColor = color;
  // ECMA-376 CT_Font/CT_RPrElt order, also used in Office-authored comments.xlsx.
  X.fontOrder = [
    'b',
    'i',
    'strike',
    'condense',
    'extend',
    'outline',
    'shadow',
    'u',
    'vertAlign',
    'sz',
    'color',
    'name',
    'family',
    'charset',
    'scheme'
  ];
  X.styles = function (book, losses) {
    const fonts = [],
      formats = new Map(),
      xfs = [],
      named = [];
    let palette;
    const warn = (code, where) =>
      losses.push(
        B.loss(code, N('Some workbook formatting could not be read'), where, 'drop', 'Workbook formatting')
      );
    const b8 = book.version === 'BIFF8';
    for (const rec of book.globals) {
      const r = rec.r;
      try {
        if (rec.id === 0x31) {
          const f = {
            size: r.u16(0) / 20,
            flags: r.u16(2),
            color: r.u16(4),
            weight: r.u16(6),
            script: r.u16(8),
            underline: r.u8(10),
            family: r.u8(11),
            charset: r.u8(12)
          };
          f.name = X.string(r, 14, b8, book.codepage, true).value;
          fonts.push(f);
        } else if (rec.id === 0x41e) formats.set(r.u16(0), X.string(r, 2, b8, book.codepage, !b8).value);
        else if (rec.id === 0xe0) {
          const flags = r.u16(4),
            align = r.u16(6),
            misc = b8 ? r.u16(8) : align,
            one = b8 ? r.u32(10) : r.u32(12),
            two = b8 ? r.u32(14) : r.u32(8),
            fill = b8 ? r.u16(18) : two;
          xfs.push({
            font: r.u16(0),
            numFmt: r.u16(2),
            flags,
            parent: flags >>> 4,
            style: !!(flags & 4),
            used: misc >>> 10,
            align: {
              horizontal: hAlign[align & 7],
              vertical: vAlign[(align >>> 4) & 7] || 'bottom',
              wrapText: align & 8 ? 1 : 0,
              textRotation: b8 ? align >>> 8 : [0, 255, 90, 180][(align >>> 8) & 3],
              indent: b8 ? misc & 15 : 0,
              shrinkToFit: b8 && misc & 16 ? 1 : 0,
              readingOrder: b8 ? Math.min(2, (misc >>> 6) & 3) : 0,
              justifyLastLine: b8 && align & 128 ? 1 : 0
            },
            border: b8
              ? {
                  left: [one & 15, (one >>> 16) & 127],
                  right: [(one >>> 4) & 15, (one >>> 23) & 127],
                  top: [(one >>> 8) & 15, two & 127],
                  bottom: [(one >>> 12) & 15, (two >>> 7) & 127],
                  diagonal: [(two >>> 21) & 15, (two >>> 14) & 127],
                  down: !!(one & 0x40000000),
                  up: !!(one & 0x80000000)
                }
              : {
                  left: [(one >>> 3) & 7, (one >>> 16) & 127],
                  right: [(one >>> 6) & 7, (one >>> 23) & 127],
                  top: [one & 7, (one >>> 9) & 127],
                  bottom: [(two >>> 22) & 7, (two >>> 25) & 127]
                },
            pattern: b8 ? two >>> 26 : (two >>> 16) & 63,
            foreground: fill & 127,
            background: (fill >>> 7) & 127
          });
        } else if (rec.id === 0x293) {
          const id = r.u16(0),
            builtin = !!(id & 0x8000),
            kind = builtin ? r.u8(2) : null,
            level = builtin ? r.u8(3) : null;
          const name = builtin
            ? (builtins[kind] || 'Style ' + kind) + (kind === 1 || kind === 2 ? level + 1 : '')
            : X.string(r, 2, b8, book.codepage, !b8).value;
          named.push({
            source: id & 4095,
            name,
            builtinId: kind,
            level: kind === 1 || kind === 2 ? level : null
          });
        } else if (rec.id === 0x892 && named.length) {
          const n = named.at(-1),
            flags = r.u8(12);
          n.hidden = flags & 2 ? 1 : null;
          n.customBuiltin = flags & 4 ? 1 : null;
          if (flags & 1) {
            n.builtinId = r.u8(14);
            n.level = [1, 2].includes(n.builtinId) ? r.u8(15) : null;
          }
        } else if (rec.id === 0x92) {
          const count = r.u16(0);
          if (count > 56) B.fail('PALETTE_LIMIT', N('The workbook palette is too large'));
          r.check(2, count * 4);
          palette = defaultPalette.map((c, i) =>
            i < count
              ? 'FF' +
                [r.u8(2 + i * 4), r.u8(3 + i * 4), r.u8(4 + i * 4)]
                  .map(n => n.toString(16).padStart(2, '0'))
                  .join('')
                  .toUpperCase()
              : c
          );
        }
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        // Keep indexes stable if an optional formatting record is damaged.
        if (rec.id === 0x31) fonts.push(null);
        if (rec.id === 0xe0) xfs.push(null);
        warn('legacy-xls-format', rec.offset);
      }
    }
    X.styleExtensions(book, xfs, losses);
    const baseFont = {
      name: 'Arial',
      size: 10,
      flags: 0,
      color: 0x7fff,
      weight: 400,
      script: 0,
      underline: 0,
      family: 2,
      charset: 0
    };
    if (!fonts.length) fonts.push(baseFont);
    for (let i = 0; i < fonts.length; i++) fonts[i] ||= baseFont;
    const boldFont = fonts.length;
    fonts.push({ ...fonts[0], weight: 700 });
    const fontIndex = id =>
      id === 4
        ? boldFont
        : id >= 5 && id - 1 < boldFont
          ? id - 1
          : id >= 0 && id < 4 && id < boldFont
            ? id
            : 0;
    const fontBody = (f, rich = false) => {
      const parts = {
        name: el(rich ? 'rFont' : 'name', { val: f.name || 'Arial' }),
        sz: el('sz', { val: f.size > 0 && f.size <= 409 ? f.size : 10 })
      };
      if (f.weight >= 700) parts.b = el('b');
      if (f.flags & 2) parts.i = el('i');
      if (f.flags & 8) parts.strike = el('strike');
      if (f.flags & 16) parts.outline = el('outline');
      if (f.flags & 32) parts.shadow = el('shadow');
      const underline = { 1: 'single', 2: 'double', 33: 'singleAccounting', 34: 'doubleAccounting' }[
        f.underline
      ];
      if (underline) parts.u = el('u', { val: underline });
      if (f.script === 1 || f.script === 2)
        parts.vertAlign = el('vertAlign', { val: f.script === 1 ? 'superscript' : 'subscript' });
      parts.color = color('color', f.color);
      if (f.family <= 5) parts.family = el('family', { val: f.family });
      if (f.charset) parts.charset = el('charset', { val: f.charset });
      if (f.scheme) parts.scheme = el('scheme', { val: f.scheme });
      return X.fontOrder.map(key => parts[key] || '').join('');
    };
    const defaultXF = {
      font: 0,
      numFmt: 0,
      flags: 1,
      parent: 0,
      style: false,
      used: 63,
      align: { vertical: 'bottom' },
      border: {},
      pattern: 0,
      foreground: 64,
      background: 65
    };
    if (!xfs.length) xfs.push(defaultXF);
    for (let i = 0; i < xfs.length; i++) xfs[i] ||= defaultXF;
    const def =
      xfs[15] && !xfs[15].style
        ? 15
        : Math.max(
            0,
            xfs.findIndex(x => !x.style)
          );
    const sourceOrder = [def, ...xfs.map((_, i) => i).filter(i => i !== def)],
      xfMap = new Map(sourceOrder.map((id, i) => [id, i]));
    const normal =
      named.find(n => n.builtinId === 0)?.source ??
      Math.max(
        0,
        xfs.findIndex(x => x.style)
      );
    const styleOrder = [normal, ...xfs.map((x, i) => (x.style ? i : -1)).filter(i => i >= 0 && i !== normal)],
      styleMap = new Map(styleOrder.map((id, i) => [id, i]));
    const fills = [
        el('fill', {}, el('patternFill', { patternType: 'none' })),
        el('fill', {}, el('patternFill', { patternType: 'gray125' }))
      ],
      fillMap = new Map(fills.map((s, i) => [s, i]));
    const borders = [
        el('border', {}, ['left', 'right', 'top', 'bottom', 'diagonal'].map(tag => el(tag)).join(''))
      ],
      borderMap = new Map([[borders[0], 0]]);
    const intern = (list, map, value) => {
      if (!map.has(value)) {
        map.set(value, list.length);
        list.push(value);
      }
      return map.get(value);
    };
    const extendedFonts = new Map();
    const xfXml = (xf, isStyle) => {
      let fontId = fontIndex(xf.font);
      if (xf.fontColor || xf.scheme) {
        const key = JSON.stringify([fontId, xf.fontColor, xf.scheme]);
        if (!extendedFonts.has(key)) {
          extendedFonts.set(key, fonts.length);
          fonts.push({
            ...fonts[fontId],
            ...(xf.fontColor ? { color: xf.fontColor } : {}),
            scheme: xf.scheme
          });
        }
        fontId = extendedFonts.get(key);
      }
      const patt = patterns[xf.pattern] || 'none',
        fill = el(
          'fill',
          {},
          xf.gradient ||
            el(
              'patternFill',
              { patternType: patt },
              patt === 'none' || patt === 'gray125'
                ? undefined
                : color('fgColor', xf.foreground) + color('bgColor', xf.background)
            )
        );
      const border = el(
        'border',
        { diagonalDown: xf.border.down ? 1 : null, diagonalUp: xf.border.up ? 1 : null },
        ['left', 'right', 'top', 'bottom', 'diagonal']
          .map(tag => {
            const [style = 0, index = 64] = xf.border[tag] || [];
            return el(
              tag,
              { style: style ? lines[style] || 'thin' : null },
              style ? color('color', index) : undefined
            );
          })
          .join('')
      );
      const alignment = { ...xf.align };
      if (alignment.textRotation > 180 && alignment.textRotation !== 255) delete alignment.textRotation;
      for (const key of Object.keys(alignment))
        if (
          alignment[key] === 0 ||
          alignment[key] === 'general' ||
          (key === 'vertical' && alignment[key] === 'bottom')
        )
          delete alignment[key];
      const body =
        (Object.keys(alignment).length ? el('alignment', alignment) : '') +
        ((xf.flags & 3) !== 1
          ? el('protection', { locked: xf.flags & 1 ? 1 : 0, hidden: xf.flags & 2 ? 1 : 0 })
          : '');
      return el(
        'xf',
        {
          numFmtId: xf.numFmt,
          fontId,
          fillId: intern(fills, fillMap, fill),
          borderId: intern(borders, borderMap, border),
          xfId: isStyle ? null : (styleMap.get(xf.parent) ?? 0),
          applyNumberFormat: !isStyle && xf.used & 1 ? 1 : null,
          applyFont: !isStyle && xf.used & 2 ? 1 : null,
          applyAlignment: !isStyle && xf.used & 4 ? 1 : null,
          applyBorder: !isStyle && xf.used & 8 ? 1 : null,
          applyFill: !isStyle && xf.used & 16 ? 1 : null,
          applyProtection: !isStyle && xf.used & 32 ? 1 : null,
          quotePrefix: xf.flags & 8 ? 1 : null
        },
        body || undefined
      );
    };
    const styleXfs = styleOrder.map(i => xfXml(xfs[i] || defaultXF, true)),
      cellXfs = sourceOrder.map(i => xfXml(xfs[i], false));
    const names = new Set(),
      styles = [];
    if (!named.some(n => n.builtinId === 0)) named.unshift({ source: normal, name: 'Normal', builtinId: 0 });
    for (const n of named) {
      if (!n.name || names.has(n.name.toLowerCase()) || !styleMap.has(n.source)) continue;
      names.add(n.name.toLowerCase());
      styles.push(
        el('cellStyle', {
          name: n.name,
          xfId: styleMap.get(n.source),
          builtinId: n.builtinId <= 53 ? n.builtinId : null,
          iLevel: n.level,
          hidden: n.hidden,
          customBuiltin: n.customBuiltin
        })
      );
    }
    const section = (tag, items) => el(tag, { count: items.length }, items.join(''));
    let xml =
      B.xmlDeclaration +
      el(
        'styleSheet',
        { xmlns: B.ns.s },
        (formats.size
          ? section(
              'numFmts',
              [...formats].map(([numFmtId, formatCode]) => el('numFmt', { numFmtId, formatCode }))
            )
          : '') +
          section(
            'fonts',
            fonts.map(f => el('font', {}, fontBody(f)))
          ) +
          section('fills', fills) +
          section('borders', borders) +
          section('cellStyleXfs', styleXfs) +
          section('cellXfs', cellXfs) +
          section('cellStyles', styles) +
          (palette
            ? el(
                'colors',
                {},
                el(
                  'indexedColors',
                  {},
                  [
                    'FF000000',
                    'FFFFFFFF',
                    'FFFF0000',
                    'FF00FF00',
                    'FF0000FF',
                    'FFFFFF00',
                    'FFFF00FF',
                    'FF00FFFF',
                    ...palette
                  ]
                    .map(rgb => el('rgbColor', { rgb }))
                    .join('')
                )
              )
            : '')
      );
    const dxfs = [],
      dxfMap = new Map(),
      at = xml.includes('<colors>') ? xml.indexOf('<colors>') : xml.lastIndexOf('</styleSheet>');
    return {
      get xml() {
        return dxfs.length ? xml.slice(0, at) + section('dxfs', dxfs) + xml.slice(at) : xml;
      },
      dxf: body => intern(dxfs, dxfMap, el('dxf', {}, body)),
      palette: [
        '000000',
        'FFFFFF',
        'FF0000',
        '00FF00',
        '0000FF',
        'FFFF00',
        'FF00FF',
        '00FFFF',
        ...(palette || defaultPalette).map(c => c.slice(2))
      ],
      formats,
      xfMap,
      fontIndex,
      fontBody: id => fontBody(fonts[fontIndex(id)], true),
      fonts,
      xfs,
      defaultXF: def
    };
  };
})(typeof window !== 'undefined' ? window : globalThis);
