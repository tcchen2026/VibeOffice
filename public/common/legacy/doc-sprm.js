/* Binary Word property operands (MS-DOC §§2.6, 2.9).
 * Derived from Apache OpenOffice sw/source/filter/ww8/ww8scan.cxx (Apache-2.0).
 * Derived from Apache POI hwpf/sprm/{SprmOperation,CharacterSprmUncompressor,
 * ParagraphSprmUncompressor,SectionSprmUncompressor}.java (Apache-2.0).
 * The byte parser is bounded; property application keeps unsupported values local. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    D = B.doc,
    N = s => s;
  D.sprms = function (bytes) {
    const r = bytes instanceof B.Reader ? bytes : new B.Reader(bytes, 'Word properties'),
      out = [];
    let p = 0;
    while (p < r.length) {
      // A single trailing zero is padding in some producer-written grpprls.
      if (p === r.length - 1 && r.u8(p) === 0) break;
      const code = r.u16(p),
        form = code >>> 13;
      let n = [1, 1, 2, 4, 2, 2, 0, 3][form];
      p += 2;
      if (!code) continue;
      if (form === 6) {
        if (code === 0xd608) {
          n = r.u16(p) - 1;
          p += 2;
        } else {
          n = r.u8(p++);
          if (code === 0xc615 && n === 255) {
            const del = r.u8(p),
              add = r.u8(p + 1 + 4 * del);
            n = 2 + 4 * del + 3 * add;
          }
        }
      }
      r.check(p, n);
      out.push({ code, kind: (code >>> 10) & 7, id: code & 511, r, p, n });
      p += n;
      if (out.length > B.limits.records)
        B.fail('DOC_PROPERTIES', N('The document contains too many formatting properties'));
    }
    return out;
  };
  D.operand = op => op.r.reader(op.p, op.n, 'Word property ' + op.code.toString(16));
  const one = {
    0: 0x2879,
    4: 0x2602,
    5: 0x2403,
    7: 0x2405,
    8: 0x2406,
    9: 0x2407,
    12: 0x260a,
    13: 0x2470,
    14: 0x240c,
    15: 0x2471,
    24: 0x2416,
    25: 0x2417,
    29: 0x261b,
    37: 0x2423,
    44: 0x242a,
    50: 0x2430,
    51: 0x2431,
    53: 0x2433,
    54: 0x2434,
    55: 0x2435,
    56: 0x2436,
    57: 0x2437,
    58: 0x2438,
    65: 0x0800,
    66: 0x0801,
    67: 0x0802,
    71: 0x0806,
    75: 0x080a,
    77: 0x2a0c,
    78: 0x0858,
    79: 0x2859,
    80: 0x0811,
    81: 0x0818,
    83: 0x2a33,
    85: 0x0835,
    86: 0x0836,
    87: 0x0837,
    88: 0x0838,
    89: 0x0839,
    90: 0x083a,
    91: 0x083b,
    92: 0x083c,
    94: 0x2a3e,
    98: 0x2a42,
    104: 0x2a48,
    115: 0x2a53,
    116: 0x0854,
    117: 0x0855,
    118: 0x0856,
    120: 0x2640,
    123: 0x2a90,
    124: 0x2a86,
    126: 0x2443
  };
  D.pieceSprms = function (book, piece) {
    if (piece.ops) return piece.ops;
    const prm = piece.prm;
    if (!prm) return (piece.ops = []);
    if (prm & 1) {
      const bytes = book.prcs[prm >>> 1];
      if (!bytes) B.fail('DOC_PRM', N('A text piece refers to missing formatting'));
      return (piece.ops = D.sprms(bytes));
    }
    const code = one[(prm >>> 1) & 127];
    if (!code) B.fail('DOC_PRM', N('A text piece uses unsupported formatting'));
    return (piece.ops = D.sprms(new Uint8Array([code & 255, code >>> 8, prm >>> 8])));
  };
  // MS-DOC COLORREF.fAuto occupies the high byte; its RGB bytes are ignored.
  D.color = n =>
    n >>> 24 === 255
      ? 'auto'
      : [n & 255, (n >>> 8) & 255, (n >>> 16) & 255]
          .map(v => v.toString(16).padStart(2, '0'))
          .join('')
          .toUpperCase();
  D.palette = [
    'auto',
    '000000',
    '0000FF',
    '00FFFF',
    '00FF00',
    'FF00FF',
    'FF0000',
    'FFFF00',
    'FFFFFF',
    '000080',
    '008080',
    '008000',
    '800080',
    '800000',
    '808000',
    '808080',
    'C0C0C0'
  ];
  D.border = function (r) {
    if (r.length === 8)
      return {
        color: D.color(r.u32(0)),
        size: r.u8(4),
        type: r.u8(5),
        space: r.u8(6) & 31,
        shadow: !!(r.u8(6) & 32),
        frame: !!(r.u8(6) & 64)
      };
    const v = r.u32(0);
    return {
      color: D.palette[(v >>> 16) & 255] || 'auto',
      size: v & 255,
      type: (v >>> 8) & 255,
      space: (v >>> 24) & 31,
      shadow: !!(v & 0x20000000),
      frame: !!(v & 0x40000000)
    };
  };
  D.shading = function (r) {
    if (r.length >= 10) return { color: D.color(r.u32(0)), fill: D.color(r.u32(4)), pattern: r.u16(8) };
    const v = r.u16(0);
    return {
      color: D.palette[v & 31] || 'auto',
      fill: D.palette[(v >>> 5) & 31] || 'auto',
      pattern: v >>> 10
    };
  };
  function tabs(r, previous, wide) {
    const out = new Map((previous || []).map(t => [t.position, { ...t }]));
    let p = 0;
    const del = r.u8(p++),
      positions = [];
    for (let i = 0; i < del; i++) positions.push(r.i16(p + i * 2));
    p += del * 2;
    const tolerances = wide ? positions.map((_, i) => r.u16(p + i * 2)) : positions.map(() => 0);
    if (wide) p += del * 2;
    for (let i = 0; i < del; i++) {
      for (const key of out.keys()) if (Math.abs(key - positions[i]) <= tolerances[i]) out.delete(key);
      out.set(positions[i], { position: positions[i], clear: true });
    }
    const add = r.u8(p++);
    r.check(p, add * 3);
    for (let i = 0; i < add; i++) {
      const position = r.i16(p + i * 2),
        flags = r.u8(p + add * 2 + i);
      out.set(position, { position, align: flags & 7, leader: (flags >>> 3) & 7 });
    }
    return [...out.values()].sort((a, b) => a.position - b.position);
  }
  const toggles = {
    0x35: 'bold',
    0x36: 'italic',
    0x37: 'strike',
    0x38: 'outline',
    0x39: 'shadow',
    0x3a: 'smallCaps',
    0x3b: 'caps',
    0x3c: 'vanish',
    0x54: 'imprint',
    0x58: 'emboss',
    0x5a: 'rtl',
    0x5c: 'boldCs',
    0x5d: 'italicCs',
    0x68: 'snapToGrid',
    0x75: 'noProof',
    0x82: 'complexScript'
  };
  const cpNumbers = {
    0x3d: 'font',
    0x3e: 'underline',
    0x40: 'spacing',
    0x41: 'language',
    0x43: 'size',
    0x45: 'position',
    0x48: 'vertAlign',
    0x4b: 'kern',
    0x4f: 'font',
    0x50: 'fontEastAsia',
    0x51: 'fontOther',
    0x52: 'scale',
    0x5e: 'fontCs',
    0x5f: 'languageCs',
    0x61: 'sizeCs',
    0x6b: 'codepage',
    0x6d: 'language',
    0x6e: 'languageEastAsia',
    0x6f: 'hint',
    0x73: 'language',
    0x74: 'languageEastAsia'
  };
  const ppNumbers = {
    0: 'style',
    3: 'align',
    0xa: 'level',
    0xb: 'list',
    0xe: 'right',
    0xf: 'left',
    0x11: 'first',
    0x13: 'before',
    0x14: 'after',
    0x18: 'frameX',
    0x19: 'frameY',
    0x1a: 'frameWidth',
    0x1b: 'frameAnchor',
    0x22: 'frameHSpace',
    0x23: 'frameWrap',
    0x2b: 'frameHeight',
    0x2c: 'dropCap',
    0x2e: 'frameVSpace',
    0x2f: 'frameHSpace',
    0x39: 'textAlignment',
    0x3a: 'textDirection',
    0x40: 'outlineLevel',
    0x49: 'tableDepth',
    0x5d: 'right',
    0x5e: 'left',
    0x60: 'first',
    0x61: 'alignLogical',
    0x55: 'rightChars',
    0x56: 'leftChars',
    0x57: 'firstChars',
    0x58: 'beforeLines',
    0x59: 'afterLines'
  };
  const ppFlags = {
    5: 'keepLines',
    6: 'keepNext',
    7: 'pageBreakBefore',
    0xc: 'suppressLineNumbers',
    0x16: 'inTable',
    0x17: 'tableEnd',
    0x2a: 'suppressAutoHyphens',
    0x30: 'locked',
    0x31: 'widowControl',
    0x33: 'kinsoku',
    0x34: 'wordWrap',
    0x35: 'overflowPunct',
    0x36: 'topLinePunct',
    0x37: 'autoSpaceDE',
    0x38: 'autoSpaceDN',
    0x41: 'bidi',
    0x47: 'snapToGrid',
    0x48: 'adjustRightInd',
    0x4b: 'innerCell',
    0x4c: 'innerRow',
    0x5b: 'beforeAuto',
    0x5c: 'afterAuto',
    0x6d: 'contextualSpacing',
    0x70: 'mirrorIndents'
  };
  const spNumbers = {
    0: 'chapterSep',
    1: 'chapterStyle',
    7: 'paperFirst',
    8: 'paperOther',
    9: 'breakType',
    0xb: 'columnsMinusOne',
    0xc: 'columnSpace',
    0xe: 'pageNumberFormat',
    0x13: 'lineRestart',
    0x14: 'headerFlags',
    0x15: 'lineCountBy',
    0x16: 'lineDistance',
    0x17: 'headerDistance',
    0x18: 'footerDistance',
    0x1a: 'verticalAlign',
    0x1b: 'lineStart',
    0x1c: 'pageStart',
    0x1f: 'width',
    0x20: 'height',
    0x21: 'left',
    0x22: 'right',
    0x23: 'top',
    0x24: 'bottom',
    0x25: 'gutter',
    0x26: 'paperSize',
    0x2f: 'pageBorderFlags',
    0x30: 'charSpace',
    0x31: 'linePitch',
    0x32: 'gridType',
    0x33: 'textDirection',
    0x3b: 'footnotePosition',
    0x3c: 'footnoteRestart',
    0x3e: 'endnoteRestart',
    0x3f: 'footnoteStart',
    0x40: 'footnoteFormat',
    0x41: 'endnoteStart',
    0x42: 'endnoteFormat',
    0x44: 'pageStart'
  };
  const spFlags = {
    5: 'evenColumns',
    6: 'unlocked',
    0xa: 'titlePage',
    0x11: 'pageRestart',
    0x12: 'endnotes',
    0x19: 'columnLine',
    0x1d: 'landscape',
    0x28: 'bidi',
    0x29: 'facingColumns',
    0x2a: 'rtlGutter'
  };
  // Non-content producer/revision IDs are retained by later phases, not reported as lost formatting.
  const metadata = new Set([
    0x4804, 0x6805, 0x4807, 0x480b, 0x2a10, 0x6815, 0x6816, 0x6817, 0x4863, 0x6864, 0x4867, 0x681e, 0x6869,
    0x681f, 0x6870, 0x6877, 0x6879, 0x687a, 0x687b, 0x687c, 0x6479, 0x6467, 0x6468, 0x6469, 0x646a, 0x646b,
    0x2462, 0x2465, 0x2466, 0x2a86, 0x301d, 0x301e, 0x703a, 0x484e, 0xc66c, 0xc669
  ]);
  D.apply = function (ops, kind, base = {}, options = {}) {
    let out = { ...base };
    const unsupported = new Set(base.unsupported || []),
      toggleBase = options.toggleBase || base;
    // MS-DOC sprmCIstd/sprmCPlain reset direct formatting but retain these
    // content/markup properties. In particular, clearing formatting must not
    // make deleted text visible again.
    const reset = next => {
      const kept = { ...next };
      for (const key of [
        'deleted',
        'inserted',
        'insertAuthor',
        'insertDate',
        'deleteAuthor',
        'deleteDate',
        'rtl',
        'complexScript',
        'fieldHidden',
        'special',
        'data',
        'ole',
        'webHidden',
        'object',
        'pictureOffset',
        'revisionFormatting',
        'revisionBefore',
        'formatRevision',
        'symbol',
        'hint',
        'highlight'
      ]) {
        if (key in out) kept[key] = out[key];
        else delete kept[key];
      }
      return kept;
    };
    for (const op of ops) {
      if (op.kind !== kind) continue;
      const r = D.operand(op),
        id = op.id,
        v = op.n === 1 ? r.u8(0) : op.n === 2 ? r.u16(0) : op.n >= 4 ? r.u32(0) : 0,
        s = op.n === 1 ? (v > 127 ? v - 256 : v) : op.n === 2 ? r.i16(0) : v | 0;
      if (op.code === 0x2a83 || op.code === 0x2664) {
        if (v) {
          const previous = { ...out };
          delete previous.revisionBefore;
          delete previous.revisionFormatting;
          delete previous.formatRevision;
          out.revisionBefore = previous;
        } else delete out.revisionBefore;
        continue;
      }
      if ([0xca57, 0xca89, 0xc63f, 0xc66f].includes(op.code)) {
        out.revisionFormatting = !!r.u8(0);
        out.formatRevision = { author: r.i16(1), date: r.u32(3) };
        continue;
      }
      if (kind === 2) {
        if ([0x4804, 0x6805, 0x4863, 0x6864].includes(op.code)) {
          out[
            { 0x4804: 'insertAuthor', 0x6805: 'insertDate', 0x4863: 'deleteAuthor', 0x6864: 'deleteDate' }[
              op.code
            ]
          ] = v;
          continue;
        }
        if (toggles[id]) {
          const key = toggles[id];
          out[key] = v === 128 ? !!toggleBase[key] : v === 129 ? !toggleBase[key] : !!v;
          continue;
        }
        if (cpNumbers[id]) {
          out[cpNumbers[id]] = [0x40, 0x45].includes(id) ? s : v;
          if (id === 0x3d) out.fontOther = out.fontEastAsia = v;
          continue;
        }
        if ([0, 1, 2, 6, 0xa, 0x11, 0x18, 0x55, 0x56].includes(id)) {
          out[
            {
              0: 'deleted',
              1: 'inserted',
              2: 'fieldHidden',
              6: 'data',
              10: 'ole',
              17: 'webHidden',
              24: 'specVanish',
              85: 'special',
              86: 'object'
            }[id]
          ] = !!v;
          continue;
        }
        if (id === 3 || id === 0xe) {
          out.pictureOffset = v;
          out.special = true;
          continue;
        }
        if (id === 9) {
          out.symbol = { font: r.u16(0), character: r.u16(2) };
          out.special = true;
          continue;
        }
        if (id === 0xc) {
          out.highlight = v;
          continue;
        }
        if (id === 0x30) {
          if (options.characterStyle) out = reset(options.characterStyle(v));
          out.charStyle = v;
          continue;
        }
        if (id === 0x32) {
          for (const key of ['bold', 'italic', 'strike', 'outline', 'shadow', 'smallCaps', 'caps', 'vanish'])
            out[key] = false;
          out.underline = 0;
          out.color = 'auto';
          continue;
        }
        if (id === 0x33) {
          out = reset(options.paragraphBase || toggleBase);
          continue;
        }
        if (id === 0x42 || id === 0x60) {
          out.color = D.palette[v] || 'auto';
          continue;
        }
        if (id === 0x70) {
          out.color = D.color(v);
          continue;
        }
        if (id === 0x53) {
          out.doubleStrike = !!v;
          continue;
        }
        if (id === 0x65 || id === 0x72) {
          out.border = D.border(r);
          continue;
        }
        if (id === 0x66 || id === 0x71) {
          out.shading = D.shading(r);
          continue;
        }
        if (id === 0x77) {
          out.underlineColor = D.color(v);
          continue;
        }
        if (op.code === 0x2879) {
          out.breakClear = v;
          continue;
        }
        if (id === 0x44) {
          out.size = Math.max(2, (out.size || 20) + s * 2);
          continue;
        }
        if (id === 0x49) {
          out.size = r.u16(0);
          continue;
        }
        if (id === 0x4a) {
          out.size = Math.max(8, (out.size || 20) + r.i16(0));
          continue;
        }
        if (id === 0x4d) {
          out.size = Math.max(2, Math.round((out.size || 20) * (1 + v / 100)));
          continue;
        }
        if (id === 0x62) {
          out.fieldRevision = !!r.u8(0);
          continue;
        }
        if (id === 0x59 && v === 0) continue;
      } else if (kind === 1) {
        if (ppNumbers[id]) {
          out[ppNumbers[id]] = [
            0xb, 0xe, 0xf, 0x11, 0x18, 0x19, 0x22, 0x2e, 0x2f, 0x5d, 0x5e, 0x60, 0x55, 0x56, 0x57
          ].includes(id)
            ? s
            : v;
          continue;
        }
        if (ppFlags[id]) {
          out[ppFlags[id]] = !!v;
          continue;
        }
        if (id === 0xd || id === 0x15) {
          out.tabs = tabs(r, out.tabs, id === 0x15);
          continue;
        }
        if (id === 0x12) {
          out.line = { value: r.i16(0), multiple: !!r.u16(2) };
          continue;
        }
        if (id === 0x10) {
          out.left = Math.max(0, (out.left || 0) + s);
          continue;
        }
        if (id === 2) {
          out.style = Math.min(9, Math.max(1, (out.style || 0) + s));
          out.outlineLevel = Math.min(9, Math.max(0, (out.outlineLevel || 0) + s));
          continue;
        }
        const side = id >= 0x24 && id <= 0x29 ? id - 0x24 : id >= 0x4e && id <= 0x53 ? id - 0x4e : -1;
        if (side >= 0) {
          out.borders = {
            ...out.borders,
            [['top', 'left', 'bottom', 'right', 'between', 'bar'][side]]: D.border(r)
          };
          continue;
        }
        if (id === 0x2d || id === 0x4d) {
          out.shading = D.shading(r);
          continue;
        }
        if (id === 0x4a) {
          out.tableDepth = (out.tableDepth || 0) + s;
          continue;
        }
        if (id === 0x3f || id === 0x43 || id === 0x6f || op.code === 0xc645) {
          out.revisionFormatting = !!r.u8(0);
          continue;
        }
        if (id === 0x3e) {
          out.oldNumbering = r.bytes;
          continue;
        }
      } else if (kind === 4) {
        if (spNumbers[id]) {
          out[spNumbers[id]] = [0x23, 0x24, 0x30].includes(id) ? s : v;
          continue;
        }
        if (spFlags[id]) {
          out[spFlags[id]] = !!v;
          continue;
        }
        if (id === 3 || id === 4) {
          const index = r.u8(0);
          out.columnInfo = {
            ...out.columnInfo,
            [index]: { ...out.columnInfo?.[index], [id === 3 ? 'width' : 'space']: r.u16(1) }
          };
          continue;
        }
        const side = id >= 0x2b && id <= 0x2e ? id - 0x2b : id >= 0x34 && id <= 0x37 ? id - 0x34 : -1;
        if (side >= 0) {
          out.borders = { ...out.borders, [['top', 'left', 'bottom', 'right'][side]]: D.border(r) };
          continue;
        }
        if (id === 0x27 || id === 0x39 || id === 0x43) {
          out.revisionFormatting = !!r.u8(0);
          continue;
        }
        if ([0xd, 0xf, 0x10, 0x1e].includes(id)) continue;
      }
      if (!metadata.has(op.code)) unsupported.add(op.code);
    }
    if (unsupported.size) out.unsupported = [...unsupported];
    return out;
  };
})(typeof window !== 'undefined' ? window : globalThis);
