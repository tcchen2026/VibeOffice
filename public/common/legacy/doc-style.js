/* Word fonts, stylesheet and property inheritance.
 * Derived from Apache POI hwpf/model/{StyleSheet,StyleDescription,FontTable,Ffn}
 * and hwpf/sprm/CharacterSprmUncompressor.java (Apache-2.0).
 * MS-DOC: Stsh, StdfBase, StdfPost2000, SttbfFfn and §2.4.6 properties. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    D = B.doc,
    N = s => s;
  function fonts(book) {
    const r = D.table(book, 15, 'SttbfFfn');
    if (!r.length) return [];
    const count = r.u16(0),
      extra = r.u16(2),
      out = [];
    let p = 4;
    for (let i = 0; i < count; i++) {
      const size = r.u8(p) + 1,
        f = r.reader(p, size, 'FFN');
      p += size + extra;
      const names = f.text(40, size - 40, 1200).split('\0');
      out.push({
        name: names[0],
        alternate: f.u8(5) ? f.text(40 + f.u8(5) * 2, size - 40 - f.u8(5) * 2, 1200).split('\0')[0] : null,
        pitch: f.u8(1) & 3,
        family: (f.u8(1) >>> 4) & 7,
        weight: f.u16(2),
        charset: f.u8(4),
        panose: f.sub(6, 10),
        signature: f.sub(16, 24)
      });
    }
    return out;
  }
  const builtin = {
    0: 'Normal',
    10: 'Index1',
    19: 'TOC1',
    28: 'NormalIndent',
    29: 'FootnoteText',
    30: 'AnnotationText',
    31: 'Header',
    32: 'Footer',
    33: 'IndexHeading',
    34: 'Caption',
    35: 'TableofFigures',
    36: 'EnvelopeAddress',
    37: 'EnvelopeReturn',
    38: 'FootnoteReference',
    39: 'AnnotationReference',
    40: 'LineNumber',
    41: 'PageNumber',
    42: 'EndnoteReference',
    43: 'EndnoteText',
    44: 'TableofAuthorities',
    45: 'MacroText',
    46: 'TOAHeading',
    47: 'List',
    48: 'ListBullet',
    49: 'ListNumber',
    50: 'List2',
    51: 'List3',
    52: 'List4',
    53: 'List5',
    54: 'ListBullet2',
    55: 'ListBullet3',
    56: 'ListBullet4',
    57: 'ListBullet5',
    58: 'ListNumber2',
    59: 'ListNumber3',
    60: 'ListNumber4',
    61: 'ListNumber5',
    62: 'Title',
    63: 'Closing',
    64: 'Signature',
    65: 'DefaultParagraphFont',
    66: 'BodyText',
    67: 'BodyTextIndent',
    68: 'ListContinue',
    69: 'ListContinue2',
    70: 'ListContinue3',
    71: 'ListContinue4',
    72: 'ListContinue5',
    73: 'MessageHeader',
    74: 'Subtitle',
    85: 'Hyperlink',
    86: 'FollowedHyperlink',
    90: 'DocumentMap',
    91: 'PlainText'
  };
  for (let i = 1; i <= 9; i++) {
    builtin[i] = 'Heading' + i;
    builtin[9 + i] = 'Index' + i;
    builtin[18 + i] = 'TOC' + i;
  }
  function styles(book) {
    const r = D.table(book, 1, 'STSH');
    if (!r.length) return [];
    const header = r.u16(0),
      h = r.reader(2, header, 'STSHI'),
      count = h.u16(0),
      base = h.u16(2),
      out = [],
      ids = new Set();
    let p = header + 2;
    if (count > 4095 || base < 10) B.fail('DOC_STYLES', N('The document stylesheet is damaged'));
    if (header >= 18)
      Object.assign(book.defaults, { font: h.u16(12), fontEastAsia: h.u16(14), fontOther: h.u16(16) });
    for (let i = 0; i < count; i++) {
      const size = r.u16(p);
      p += 2;
      if (!size) {
        out.push(null);
        continue;
      }
      const bytes = r.reader(p, size, 'STD');
      p += size;
      out.push(
        D.optional(
          book,
          'Style ' + i,
          N('A damaged text style was not kept'),
          () => {
            const info = bytes.u16(0),
              b = bytes.u16(2),
              c = bytes.u16(4),
              flags = bytes.u16(8),
              n = bytes.u16(base),
              name = bytes.text(base + 2, n * 2, 1200).split(',')[0];
            let q = base + 2 + (n + 1) * 2;
            const upx = [];
            for (let j = 0; j < (c & 15); j++) {
              const length = bytes.u16(q);
              q += 2;
              upx.push(bytes.reader(q, length, 'UPX'));
              q += length + (length & 1);
            }
            const kind = b & 15,
              sti = info & 4095;
            let id = builtin[sti] || name.replace(/[^\p{L}\p{N}]/gu, '').slice(0, 200) || 'Style' + i;
            if (ids.has(id)) id = 'Style' + i;
            ids.add(id);
            const pap = kind === 1 ? upx[0] : null,
              chp = kind === 1 ? upx[1] : kind === 2 ? upx[0] : null;
            return {
              index: i,
              id,
              name: name || id,
              sti,
              kind,
              basedOn: b >>> 4,
              next: c >>> 4,
              flags,
              link: base >= 18 ? bytes.u16(10) & 4095 : 4095,
              pOps: pap && pap.length >= 2 ? D.sprms(pap.reader(2, pap.length - 2)) : [],
              cOps: chp ? D.sprms(chp) : []
            };
          },
          null,
          false
        )
      );
    }
    return out;
  }
  D.styles = function (book) {
    book.defaults = {
      font: 0,
      fontEastAsia: 0,
      fontOther: 0,
      size: 20,
      sizeCs: 20,
      color: 'auto',
      language: book.r.u16(6)
    };
    book.fonts = D.optional(book, 'Fonts', N('Some font choices could not be read'), () => fonts(book), []);
    book.styles = D.optional(book, 'Styles', N('Text styles could not be read'), () => styles(book), []);
    if (!book.styles[0])
      book.styles[0] = {
        index: 0,
        id: 'Normal',
        name: 'Normal',
        sti: 0,
        kind: 1,
        basedOn: 4095,
        next: 0,
        flags: 0,
        pOps: [],
        cOps: []
      };
    const visiting = new Set();
    function resolve(style, depth = 0) {
      if (style.resolved) return style;
      if (visiting.has(style.index) || depth >= B.limits.depth) {
        style.basedOn = 4095;
        style.p = {};
        style.c = { ...book.defaults };
        style.resolved = true;
        style.damaged = true;
        return style;
      }
      visiting.add(style.index);
      const parent = book.styles[style.basedOn],
        base = parent && parent.kind === style.kind ? resolve(parent, depth + 1) : null;
      style.p = D.optional(
        book,
        'Style ' + style.name,
        N('Some paragraph style formatting could not be read'),
        () => D.apply(style.pOps, 1, base?.p || {}),
        {},
        false
      );
      style.c = D.optional(
        book,
        'Style ' + style.name,
        N('Some character style formatting could not be read'),
        () => D.apply(style.cOps, 2, base?.c || book.defaults),
        { ...book.defaults },
        false
      );
      visiting.delete(style.index);
      style.resolved = true;
      return style;
    }
    book.styles.filter(Boolean).forEach(s => resolve(s));
    return book.styles;
  };
  D.style = function (book, index, kind, where) {
    const style = book.styles[index];
    if (style && style.kind === kind) {
      if (style.damaged)
        D.notice(
          book,
          'legacy-word-style',
          N('A damaged style was replaced with default formatting'),
          where,
          true
        );
      return style;
    }
    if (index != null && index !== 4095 && index !== 0)
      D.notice(
        book,
        'legacy-word-style',
        N('A missing text style was replaced with default formatting'),
        where,
        true,
        String(index)
      );
    return kind === 1 ? book.styles[0] : null;
  };
  D.properties = function (book, cp, kind, paragraph, where) {
    const { piece, run, end } = D.at(book, kind === 1 ? book.papx : book.chpx, cp),
      ops = run?.ops || [];
    const prm = piece
      ? D.optional(
          book,
          where,
          N('Some text formatting could not be read'),
          () => D.pieceSprms(book, piece),
          []
        )
      : [];
    if (kind === 1) {
      let index = run?.style || 0;
      for (const op of [...ops, ...prm]) if (op.code === 0x4600) index = D.operand(op).u16(0);
      const style = D.style(book, index, 1, where),
        initial = D.apply([...ops, ...prm], 1, style.p),
        list = D.listFor(book, initial, where);
      const base = list?.definition && initial.list > 0 ? { ...style.p, ...list.definition.p } : style.p;
      const properties = D.apply([...ops, ...prm], 1, base);
      properties.style = index;
      return { props: properties, style, end };
    }
    const base = paragraph?.c || book.defaults,
      cache = new Map();
    // CIstd acts at its position in the property sequence. Apply its ancestry
    // to the paragraph style; later direct properties can override it.
    function characterStyle(index) {
      if (cache.has(index)) return cache.get(index);
      const chain = [],
        seen = new Set();
      let s = D.style(book, index, 2, where),
        current = { ...base };
      while (s && s.kind === 2 && !seen.has(s.index) && chain.length < B.limits.depth) {
        seen.add(s.index);
        chain.push(s);
        s = book.styles[s.basedOn];
      }
      for (const s of chain.reverse()) current = D.apply(s.cOps, 2, current);
      cache.set(index, current);
      return current;
    }
    const props = D.apply([...ops, ...prm], 2, base, {
      toggleBase: base,
      paragraphBase: base,
      characterStyle
    });
    return { props, base: cache.get(props.charStyle) || base, end };
  };
})(typeof window !== 'undefined' ? window : globalThis);
