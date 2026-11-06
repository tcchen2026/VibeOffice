/* Word DOP/settings and section output.
 * Derived from Apache OpenOffice sw/source/filter/ww8/ww8scan.cxx and
 * ww8par.cxx (Apache-2.0). MS-DOC DopBase/Dop95/Dop97/Dop2000 and Copts.
 * Compatibility bits are translated individually, including inverted flags.
 * Derived from Open XML SDK data/schemas/schemas_openxmlformats_org_wordprocessingml_2006_main.json (MIT):
 * compatibility ordering. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    D = B.doc,
    N = s => s,
    w = (...args) => D.w(...args),
    v = (name, val) => (val == null ? '' : w(name, { val }));
  const first = [
    'noTabHangInd',
    'noSpaceRaiseLower',
    'suppressSpBfAfterPgBrk',
    'wrapTrailSpaces',
    'printColBlack',
    'noColumnBalance',
    'convMailMergeEsc',
    'suppressTopSpacing',
    'useSingleBorderforContiguousCells',
    null,
    'showBreaksInFrames',
    'swapBordersFacingPages',
    'doNotLeaveBackslashAlone',
    'doNotExpandShiftReturn',
    'ulTrailSpace',
    'balanceSingleByteDoubleByteWidth',
    'suppressSpacingAtTopOfPage',
    'spacingInWholePoints',
    'printBodyTextBeforeHeader',
    'noLeading',
    'spaceForUL',
    'mwSmallCaps',
    'suppressTopSpacingWP',
    'truncateFontHeightsLikeWP6',
    'subFontBySize',
    'lineWrapLikeWord6',
    'doNotSuppressParagraphBorders',
    'noExtraLineSpacing',
    'suppressBottomSpacing',
    'wpSpaceWidth',
    'wpJustification',
    'usePrinterMetrics'
  ];
  const second = [
    'shapeLayoutLikeWW8',
    'footnoteLayoutLikeWW8',
    'doNotUseHTMLParagraphAutoSpacing',
    'adjustLineHeightInTable',
    'forgetLastTabAlignment',
    'autoSpaceLikeWord95',
    'alignTablesRowByRow',
    'layoutRawTableWidth',
    'layoutTableRowsApart',
    'useWord97LineBreakRules',
    'doNotBreakWrappedTables',
    'doNotSnapToGridInCell',
    'selectFldWithFirstOrLastChar',
    'applyBreakingRules',
    'doNotWrapTextWithPunct',
    'doNotUseEastAsianBreakRules',
    'useWord2002TableStyleRules',
    'growAutofit',
    'useNormalStyleForList',
    'doNotUseIndentAsNumberingTabStop',
    'useAltKinsokuLineBreakRules',
    'allowSpaceOfSameStyleInTable',
    'doNotSuppressIndentation',
    'doNotAutofitConstrainedTables',
    'autofitToFirstFixedWidthCell',
    'underlineTabInNumList',
    'displayHangulFixedWidth',
    'splitPgBreakAndParaMark',
    'doNotVertAlignCellWithSp',
    'doNotBreakConstrainedForcedTable',
    'doNotVertAlignInTxbx',
    'useAnsiKerningPairs'
  ];
  // ECMA-376 CT_Compat sequence (also checked against Open XML SDK v3.3.0).
  const compatOrder = [
    'useSingleBorderforContiguousCells',
    'wpJustification',
    'noTabHangInd',
    'noLeading',
    'spaceForUL',
    'noColumnBalance',
    'balanceSingleByteDoubleByteWidth',
    'noExtraLineSpacing',
    'doNotLeaveBackslashAlone',
    'ulTrailSpace',
    'doNotExpandShiftReturn',
    'spacingInWholePoints',
    'lineWrapLikeWord6',
    'printBodyTextBeforeHeader',
    'printColBlack',
    'wpSpaceWidth',
    'showBreaksInFrames',
    'subFontBySize',
    'suppressBottomSpacing',
    'suppressTopSpacing',
    'suppressSpacingAtTopOfPage',
    'suppressTopSpacingWP',
    'suppressSpBfAfterPgBrk',
    'swapBordersFacingPages',
    'convMailMergeEsc',
    'truncateFontHeightsLikeWP6',
    'mwSmallCaps',
    'usePrinterMetrics',
    'doNotSuppressParagraphBorders',
    'wrapTrailSpaces',
    'footnoteLayoutLikeWW8',
    'shapeLayoutLikeWW8',
    'alignTablesRowByRow',
    'forgetLastTabAlignment',
    'adjustLineHeightInTable',
    'autoSpaceLikeWord95',
    'noSpaceRaiseLower',
    'doNotUseHTMLParagraphAutoSpacing',
    'layoutRawTableWidth',
    'layoutTableRowsApart',
    'useWord97LineBreakRules',
    'doNotBreakWrappedTables',
    'doNotSnapToGridInCell',
    'selectFldWithFirstOrLastChar',
    'applyBreakingRules',
    'doNotWrapTextWithPunct',
    'doNotUseEastAsianBreakRules',
    'useWord2002TableStyleRules',
    'growAutofit',
    'useFELayout',
    'useNormalStyleForList',
    'doNotUseIndentAsNumberingTabStop',
    'useAltKinsokuLineBreakRules',
    'allowSpaceOfSameStyleInTable',
    'doNotSuppressIndentation',
    'doNotAutofitConstrainedTables',
    'autofitToFirstFixedWidthCell',
    'underlineTabInNumList',
    'displayHangulFixedWidth',
    'splitPgBreakAndParaMark',
    'doNotVertAlignCellWithSp',
    'doNotBreakConstrainedForcedTable',
    'doNotVertAlignInTxbx',
    'useAnsiKerningPairs',
    'cachedColBalance',
    'compatSetting'
  ];
  D.settings = function (book) {
    const read = () => {
      const r = D.table(book, 31, 'DOP'),
        u16 = p => (p + 2 <= r.length ? r.u16(p) : 0),
        u32 = p => (p + 4 <= r.length ? r.u32(p) : 0),
        byte = p => (p < r.length ? r.u8(p) : 0);
      const f = u16(0),
        notes = u16(54),
        view = u16(82),
        compat = r.length >= 512 ? u32(508) : r.length >= 88 ? u32(84) : u16(8);
      return {
        evenHeaders: !!(f & 1),
        widow: !!(f & 2),
        footnotePosition: (f >>> 5) & 3,
        footnoteRestart: u16(2) & 3,
        footnoteStart: u16(2) >>> 2,
        endnoteRestart: u16(52) & 3,
        endnoteStart: u16(52) >>> 2,
        endnotePosition: notes & 3,
        footnoteFormat: r.length >= 496 ? u16(492) : (notes >>> 2) & 15,
        endnoteFormat: r.length >= 496 ? u16(494) : (notes >>> 6) & 15,
        tab: u16(10) || 720,
        hyphenationZone: u16(14),
        consecutiveHyphens: u16(16),
        autoHyphen: !!(byte(5) & 16),
        hyphenateCaps: !!(byte(5) & 8),
        mirrorMargins: !!(byte(6) & 32),
        printHidden: !!(byte(6) & 4),
        printResults: !!(byte(6) & 8),
        gutterAtTop: !!(view & 32768),
        zoom: (view >>> 3) & 511,
        view: view & 7,
        compat,
        compat2: r.length >= 516 ? u32(512) : 7,
        compat3: r.length >= 520 ? u32(516) : 0,
        storedPages: u16(46),
        storedWords: u32(38),
        storedCharacters: u32(42),
        protected: !!(byte(7) & 2),
        trackRevisions: !!(byte(5) & 128),
        embeddedFonts: !!(byte(7) & 128)
      };
    };
    book.settings = D.optional(
      book,
      'Document settings',
      N('Some document layout settings could not be read'),
      read,
      { tab: 720, compat: 0, compat2: 7 }
    );
    return book.settings;
  };
  D.noteSettings = function (kind, p, global = false) {
    let out = '';
    const prefix = kind === 'footnote' ? 'footnote' : 'endnote';
    const position = p[prefix + 'Position'];
    if (position != null)
      out += v(
        'pos',
        (kind === 'footnote' ? ['pageBottom', 'beneathText'] : ['sectEnd', 'docEnd'])[position]
      );
    out += v('numFmt', D.numberFormats[p[prefix + 'Format']]);
    const start = p[prefix + 'Start'];
    if (start > 0 && start <= 32767) out += v('numStart', start);
    out += v('numRestart', ['continuous', 'eachSect', 'eachPage'][p[prefix + 'Restart']]);
    if (global) out += w(kind, { id: -1 }) + w(kind, { id: 0 });
    return out ? w(kind + 'Pr', {}, out) : '';
  };
  D.settingsXML = function (book) {
    const p = book.settings;
    let out = v('view', { 1: 'print', 2: 'outline', 3: 'masterPages', 4: 'normal', 5: 'web' }[p.view]);
    if (p.zoom) out += w('zoom', { percent: D.bound(book, p.zoom, 10, 500, 'Document view') });
    if (p.mirrorMargins) out += w('mirrorMargins');
    if (p.gutterAtTop) out += w('gutterAtTop');
    if (p.trackRevisions) out += w('trackRevisions');
    out += v('defaultTabStop', D.bound(book, p.tab, 1, 31680, 'Document settings'));
    if (p.autoHyphen) out += w('autoHyphenation');
    if (p.consecutiveHyphens) out += v('consecutiveHyphenLimit', p.consecutiveHyphens);
    if (p.hyphenationZone) out += v('hyphenationZone', p.hyphenationZone);
    if (p.hyphenateCaps === false) out += w('doNotHyphenateCaps');
    if (p.evenHeaders) out += w('evenAndOddHeaders');
    if (book.footnotes.length) out += D.noteSettings('footnote', p, true);
    if (book.endnotes.length) out += D.noteSettings('endnote', p, true);
    const enabledFlags = new Set();
    first.forEach((key, i) => {
      if (!key) return;
      const bit = !!(p.compat & (1 << i)),
        enabled = [12, 13, 14, 15, 20].includes(i) ? !bit : bit;
      if (enabled) enabledFlags.add(key);
    });
    second.forEach((key, i) => {
      const bit = !!(p.compat2 & (1 << i));
      if (i === 3 ? !bit : bit) enabledFlags.add(key);
    });
    if (p.compat3 & 1) enabledFlags.add('cachedColBalance');
    let compat = compatOrder
      .filter(name => enabledFlags.has(name))
      .map(name => w(name))
      .join('');
    compat += w('compatSetting', {
      name: 'compatibilityMode',
      uri: 'http://schemas.microsoft.com/office/word',
      val: 11
    });
    out += w('compat', {}, compat);
    if (book.pair(60).length)
      out += D.optional(
        book,
        'Document variables',
        N('Some document variables used by fields were not kept'),
        () => {
          const r = D.table(book, 60, 'Document variables'),
            state = {},
            names = D.markup.strings(r, state),
            seen = new Set();
          let p = state.end,
            body = '';
          for (const entry of names) {
            const n = r.u16(p);
            p += 2;
            const value = D.unicode(r, p, n * 2);
            p += n * 2;
            const name = entry.text,
              key = name.toLowerCase();
            // These three entries contain binary VBA signatures, not field values.
            if (['sign', 'sigagile', 'sigv3'].includes(key)) continue;
            if (!name || seen.has(key)) {
              D.notice(
                book,
                'legacy-word-variables',
                N('A damaged document variable used by fields was not kept'),
                'Document variables',
                true
              );
              continue;
            }
            seen.add(key);
            body += w('docVar', { name, val: value });
          }
          return body ? w('docVars', {}, body) : '';
        },
        ''
      );
    return D.part('settings', out);
  };
  D.sectionXML = function (book, section) {
    const p = {
        width: 12240,
        height: 15840,
        left: 1800,
        right: 1800,
        top: 1440,
        bottom: 1440,
        gutter: 0,
        headerDistance: 720,
        footerDistance: 720,
        columnSpace: 720,
        breakType: 2,
        ...section.props
      },
      where = section.where || 'Section';
    if (p.revisionFormatting)
      D.notice(
        book,
        'legacy-word-section-revision',
        N('An earlier section formatting change was not kept; the current formatting was retained'),
        where,
        true
      );
    let out = section.references || '';
    if (book.footnotes.length) out += D.noteSettings('footnote', p);
    if (book.endnotes.length) out += D.noteSettings('endnote', p);
    out += v(
      'type',
      ['continuous', 'nextColumn', 'nextPage', 'evenPage', 'oddPage'][p.breakType] || 'nextPage'
    );
    out += w('pgSz', {
      w: D.bound(book, p.width, 1, 31680, where),
      h: D.bound(book, p.height, 1, 31680, where),
      orient: p.landscape ? 'landscape' : null,
      code: p.paperSize
    });
    out += w('pgMar', {
      top: D.bound(book, p.top, -31680, 31680, where),
      right: D.bound(book, p.right, 0, 31680, where),
      bottom: D.bound(book, p.bottom, -31680, 31680, where),
      left: D.bound(book, p.left, 0, 31680, where),
      header: D.bound(book, p.headerDistance, 0, 31680, where),
      footer: D.bound(book, p.footerDistance, 0, 31680, where),
      gutter: D.bound(book, p.gutter, 0, 31680, where)
    });
    if (p.paperFirst != null || p.paperOther != null)
      out += w('paperSrc', { first: p.paperFirst, other: p.paperOther });
    if (p.borders)
      out += w(
        'pgBorders',
        {
          offsetFrom: p.pageBorderFlags & 32 ? 'text' : 'page',
          display: ['allPages', 'firstPage', 'notFirstPage'][p.pageBorderFlags & 7],
          zOrder: p.pageBorderFlags & 8 ? 'back' : 'front'
        },
        ['top', 'left', 'bottom', 'right']
          .filter(side => p.borders[side])
          .map(side => D.borderXML(book, side, p.borders[side], where))
          .join('')
      );
    if (p.lineCountBy)
      out += w('lnNumType', {
        countBy: p.lineCountBy,
        start: p.lineStart,
        distance: p.lineDistance,
        restart: ['newPage', 'newSection', 'continuous'][p.lineRestart]
      });
    if (p.pageRestart || p.pageNumberFormat != null || p.chapterStyle)
      out += w('pgNumType', {
        fmt: D.numberFormats[p.pageNumberFormat],
        start: p.pageRestart ? D.bound(book, p.pageStart ?? 0, 0, 2147483646, where) : null,
        chapStyle: p.chapterStyle ? Math.min(9, p.chapterStyle) : null,
        chapSep: p.chapterStyle ? ['hyphen', 'period', 'colon', 'emDash', 'enDash'][p.chapterSep] : null
      });
    const count = D.bound(book, (p.columnsMinusOne || 0) + 1, 1, 45, where),
      cols = [];
    if (p.evenColumns === false && p.columnInfo) {
      for (let i = 0; i < count; i++) {
        const c = p.columnInfo[i] || {};
        cols.push(
          w('col', {
            w: D.bound(
              book,
              c.width ?? Math.max(1, (p.width - p.left - p.right - (count - 1) * p.columnSpace) / count),
              1,
              31680,
              where
            ),
            space: i + 1 < count ? D.bound(book, c.space ?? p.columnSpace, 0, 31680, where) : null
          })
        );
      }
    }
    out += w(
      'cols',
      {
        num: count > 1 ? count : null,
        space: D.bound(book, p.columnSpace, 0, 31680, where),
        sep: p.columnLine ? 1 : null,
        equalWidth: cols.length ? 0 : null
      },
      cols.length ? cols.join('') : undefined
    );
    if (p.unlocked != null) out += v('formProt', p.unlocked ? 0 : 1);
    out += v('vAlign', ['top', 'center', 'both', 'bottom'][p.verticalAlign]);
    if (p.endnotes === false) out += w('noEndnote');
    if (p.titlePage) out += w('titlePg');
    out += v('textDirection', ['lrTb', 'tbRl', 'btLr', 'lrTbV', 'tbRlV'][p.textDirection]);
    if (p.bidi) out += w('bidi');
    if (p.rtlGutter) out += w('rtlGutter');
    if (p.gridType || p.linePitch || p.charSpace)
      out += w('docGrid', {
        type: ['default', 'linesAndChars', 'lines', 'snapToChars'][p.gridType || 0],
        linePitch: p.linePitch,
        charSpace: p.charSpace
      });
    D.formatWarnings(book, p, where);
    return w('sectPr', {}, out);
  };
})(typeof window !== 'undefined' ? window : globalThis);
