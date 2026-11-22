/* Derived from Apache OpenOffice main/sc/source/filter/excel/{xipage,xicontent,excimp8}.cxx
 * and Derived from Apache POI hssf/record/{Hyperlink,PrintSetup,HeaderFooter,Protect,
 * PageBreak}Record.java (Apache-2.0). MS-XLS: Setup, HeaderFooter, AFDOper, FeatHdr.
 * Optional worksheet records: a damaged feature never discards the worksheet. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    X = B.xls,
    el = B.el,
    N = s => s;
  const bad = message => B.fail('XLS_FEATURE', message);
  X.featureLoss = function (losses, id, what, sheet, error, notify = true) {
    const loss = B.loss(id, what, sheet, 'drop', sheet, notify);
    if (error) loss.detail = { code: error.code, message: error.message };
    losses.push(loss);
  };
  X.cellRange = function (r, p) {
    const a = [r.u16(p), r.u16(p + 4), r.u16(p + 2), r.u16(p + 6)];
    if (a[0] > a[2] || a[1] > a[3] || a[3] > 255) bad('Invalid worksheet range');
    return a;
  };
  // Binary continuations (TXO/SST have their own character-width boundaries).
  X.continuedRecord = function (records, i) {
    let n = records[i].r.length,
      j = i + 1;
    while (records[j]?.id === 0x3c) {
      n += records[j++].r.length;
      if (n > B.limits.text) bad('Continued record is too large');
    }
    if (j === i + 1) return records[i].r;
    const bytes = new Uint8Array(n);
    let p = 0;
    for (let k = i; k < j; k++) {
      bytes.set(records[k].r.bytes, p);
      p += records[k].r.length;
    }
    return new B.Reader(bytes, 'Continued worksheet record');
  };
  const hash = n => (n ? n.toString(16).padStart(4, '0').toUpperCase() : null);
  X.workbookProtection = function (book, losses) {
    const attrs = {};
    for (const rec of book.globals) {
      if (![0x12, 0x19, 0x13, 0x1af, 0x1bc].includes(rec.id)) continue;
      try {
        const n = rec.r.u16(0);
        if (rec.id === 0x13) attrs.workbookPassword = hash(n);
        else if (rec.id === 0x1bc) attrs.revisionsPassword = hash(n);
        else
          attrs[rec.id === 0x12 ? 'lockStructure' : rec.id === 0x19 ? 'lockWindows' : 'lockRevision'] = n
            ? 1
            : null;
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        X.featureLoss(
          losses,
          'legacy-workbook-protection',
          N('Some workbook protection settings were not kept'),
          'Workbook',
          e
        );
      }
    }
    return attrs.lockStructure || attrs.lockWindows || attrs.lockRevision
      ? el('workbookProtection', attrs)
      : '';
  };
  const trimNull = s => s.split('\0')[0];
  X.hyperlink = function (r, cp) {
    const range = X.cellRange(r, 0),
      flags = r.u32(28);
    let p = 32;
    const unicode = () => {
      const n = r.u32(p);
      p += 4;
      const s = trimNull(r.text(p, n * 2, 1200));
      p += n * 2;
      return s;
    };
    const out = { ref: X.range(...range) };
    if (flags & 0x14) out.display = unicode();
    if (flags & 0x80) out.targetFrame = unicode();
    if (flags & 1) {
      if (flags & 0x100) out.target = unicode();
      else {
        const moniker = r.u32(p);
        r.check(p, 16);
        p += 16;
        if (moniker === 0x79eac9e0) {
          const n = r.u32(p);
          p += 4;
          out.target = trimNull(r.text(p, n, 1200));
          p += n;
        } else if (moniker === 0x303) {
          const up = r.u16(p),
            n = r.u32(p + 2);
          p += 6;
          let path = trimNull(r.text(p, n, cp));
          p += n + 24;
          const size = r.u32(p);
          p += 4;
          if (size) {
            const bytes = r.u32(p);
            p += 6;
            path = trimNull(r.text(p, bytes, 1200));
            p += bytes;
          }
          out.target = '../'.repeat(Math.min(up, 1024)) + path;
        } else if (moniker === 0x305) {
          const n = r.u32(p + 2);
          p += 6;
          out.target = trimNull(r.text(p, n, 65001));
          p += n;
        } else bad('Unsupported hyperlink moniker');
      }
      out.target = out.target.replace(/\\/g, '/');
      if (/^[a-z]:\//i.test(out.target)) out.target = 'file:///' + out.target;
      else if (out.target.startsWith('//')) out.target = 'file:' + out.target;
      if (
        !out.target ||
        /[\x00-\x1f]/.test(out.target) ||
        /^(?:javascript|vbscript|data):/i.test(out.target.trim())
      )
        bad('Unsupported hyperlink target');
      out.target = out.target.replace(/ /g, '%20');
    }
    if (flags & 8) out.location = unicode();
    if (!out.target && !out.location) bad('Hyperlink destination is missing');
    return out;
  };
  const operations = [
    null,
    'lessThan',
    'equal',
    'lessThanOrEqual',
    'greaterThan',
    'notEqual',
    'greaterThanOrEqual'
  ];
  const filterColumn = function (r, b8, cp) {
    const colId = r.u16(0),
      flags = r.u16(2);
    if (colId > 255) bad('Invalid filter column');
    if (flags & 16) {
      const n = flags >>> 7;
      if (!n || n > 500) bad('Invalid Top 10 filter');
      return {
        colId,
        body: el('top10', { top: flags & 32 ? 1 : 0, percent: flags & 64 ? 1 : null, val: n })
      };
    }
    let p = 24;
    const conditions = [];
    for (let i = 0; i < 2; i++) {
      const at = 4 + 10 * i,
        type = r.u8(at),
        op = r.u8(at + 1);
      r.check(at, 10);
      if (!type) continue;
      let val,
        operator = operations[op];
      if (!operator) bad('Unknown filter comparison');
      if (type === 2) val = X.rk(r.u32(at + 2));
      else if (type === 4) val = r.f64(at + 2);
      else if (type === 6) {
        const n = r.u8(at + 6),
          wide = b8 ? r.u8(p++) & 1 : false;
        val = b8 && !wide ? B.latin16(r.sub(p, n)) : r.text(p, n * (wide ? 2 : 1), wide ? 1200 : cp);
        p += n * (wide ? 2 : 1);
      } else if (type === 8)
        val = r.u8(at + 2) ? X.formulaErrors[r.u8(at + 3)] : r.u8(at + 3) ? 'TRUE' : 'FALSE';
      else if (type === 12 || type === 14) {
        val = '';
        operator = type === 12 ? 'equal' : 'notEqual';
      } else bad('Unsupported filter operand');
      if (val == null || (typeof val === 'number' && !Number.isFinite(val))) bad('Invalid filter operand');
      conditions.push({ operator, val });
    }
    const body = conditions.length
      ? el(
          'customFilters',
          { 'and': conditions.length > 1 && !(flags & 1) ? 1 : null },
          conditions.map(c => el('customFilter', c)).join('')
        )
      : '';
    return { colId, body };
  };
  X.sheetExtras = function (book, si, meta, formulas, pkg, part, losses) {
    const sheet = book.sheets[si],
      b8 = book.version === 'BIFF8',
      margins = { left: 0.75, right: 0.75, top: 1, bottom: 1, header: 0.5, footer: 0.5 },
      print = {},
      setup = {},
      hf = {},
      hfAttrs = {},
      protect = {},
      links = [],
      tips = new Map(),
      filters = new Map(),
      breaks = {};
    let hasMargins = false,
      hasSetup = false,
      fit = false,
      autoPageBreaks = null,
      filterCount = 0,
      filterMode = false,
      protectionOptions;
    const issue = (id, what, e, notify = true) => X.featureLoss(losses, id, what, sheet.name, e, notify);
    const records = meta.records;
    for (let i = 0; i < records.length; i++) {
      const rec = records[i],
        id = rec.id;
      if (
        ![
          0x12, 0x13, 0x63, 0xdd, 0x867, 0x81, 0x26, 0x27, 0x28, 0x29, 0x83, 0x84, 0x2a, 0x2b, 0x82, 0xa1,
          0x14, 0x15, 0x89c, 0x1a, 0x1b, 0x1b8, 0x800, 0x9d, 0x9e, 0x9b, 0x87e
        ].includes(id)
      )
        continue;
      const feature =
        id === 0x1b8 || id === 0x800
          ? 'hyperlinks'
          : [0x9d, 0x9e, 0x87e].includes(id)
            ? 'filters'
            : [0x12, 0x13, 0x63, 0xdd, 0x867].includes(id)
              ? 'protection'
              : 'print-settings';
      try {
        const r = X.continuedRecord(records, i);
        if (id === 0x12 || id === 0x63 || id === 0xdd)
          protect[id === 0x12 ? 'sheet' : id === 0x63 ? 'objects' : 'scenarios'] = r.u16(0) ? 1 : 0;
        else if (id === 0x13) protect.password = hash(r.u16(0));
        else if (id === 0x867) {
          if (r.u16(12) === 2 && r.u32(15) === 0xffffffff) protectionOptions = r.u16(19);
        } else if (id === 0x81) {
          const flags = r.u16(0);
          fit = !!(flags & 0x100);
          autoPageBreaks = flags & 1 ? 1 : 0;
        } else if (id >= 0x26 && id <= 0x29) {
          const n = r.f64(0);
          if (!Number.isFinite(n) || n < 0 || n > 49) bad('Invalid page margin');
          margins[['left', 'right', 'top', 'bottom'][id - 0x26]] = n;
          hasMargins = true;
        } else if ([0x83, 0x84, 0x2a, 0x2b, 0x82].includes(id))
          print[
            {
              131: 'horizontalCentered',
              132: 'verticalCentered',
              42: 'headings',
              43: 'gridLines',
              130: 'gridLinesSet'
            }[id]
          ] = r.u16(0) ? 1 : 0;
        else if (id === 0xa1) {
          const flags = r.u16(10),
            valid = !(flags & 4);
          hasSetup = true;
          Object.assign(setup, {
            paperSize: valid ? r.u16(0) || null : null,
            scale: valid && r.u16(2) >= 10 && r.u16(2) <= 400 ? r.u16(2) : null,
            firstPageNumber: flags & 128 ? r.u16(4) : null,
            fitToWidth: r.u16(6),
            fitToHeight: r.u16(8),
            pageOrder: flags & 1 ? 'overThenDown' : null,
            orientation: valid ? (flags & 0x40 || flags & 2 ? 'portrait' : 'landscape') : null,
            blackAndWhite: flags & 8 ? 1 : null,
            draft: flags & 16 ? 1 : null,
            cellComments: flags & 32 ? (flags & 512 ? 'atEnd' : 'asDisplayed') : null,
            useFirstPageNumber: flags & 128 ? 1 : null,
            errors: ['displayed', 'blank', 'dash', 'NA'][(flags >>> 10) & 3],
            horizontalDpi: valid ? r.u16(12) || null : null,
            verticalDpi: valid ? r.u16(14) || null : null,
            copies: valid ? r.u16(32) || null : null
          });
          for (const [key, at] of [
            ['header', 16],
            ['footer', 24]
          ]) {
            const n = r.f64(at);
            if (!Number.isFinite(n) || n < 0 || n > 49) bad('Invalid header/footer margin');
            margins[key] = n;
          }
          hasMargins = true;
        } else if (id === 0x14 || id === 0x15)
          hf[id === 0x14 ? 'oddHeader' : 'oddFooter'] = r.length
            ? X.string(r, 0, b8, book.codepage, !b8).value
            : '';
        else if (id === 0x89c) {
          if (r.sub(12, 16).some(v => v)) continue;
          const flags = r.u16(28);
          Object.assign(hfAttrs, {
            differentOddEven: flags & 1 ? 1 : 0,
            differentFirst: flags & 2 ? 1 : 0,
            scaleWithDoc: flags & 4 ? 1 : 0,
            alignWithMargins: flags & 8 ? 1 : 0
          });
          let p = 38;
          for (const [j, key] of ['evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'].entries())
            if (r.u16(30 + 2 * j)) {
              const s = X.string(r, p);
              hf[key] = s.value;
              p = s.end;
            }
        } else if (id === 0x1a || id === 0x1b) {
          const list = [],
            n = r.u16(0),
            step = b8 ? 6 : 2;
          r.check(2, n * step);
          for (let j = 0; j < n; j++) {
            const at = 2 + j * step,
              id = r.u16(at),
              min = b8 ? r.u16(at + 2) : 0,
              max = b8 ? r.u16(at + 4) : rec.id === 0x1a ? 65535 : 255;
            if (id && min <= max) list.push({ id, min, max, man: 1 });
          }
          breaks[id === 0x1b ? 'rowBreaks' : 'colBreaks'] = list;
        } else if (id === 0x1b8) links.push(X.hyperlink(r, book.codepage));
        else if (id === 0x800) {
          const ref = X.range(...X.cellRange(r, 2));
          tips.set(ref, trimNull(r.text(10, r.length - 10, 1200)));
        } else if (id === 0x9d) filterCount = r.u16(0);
        else if (id === 0x9b) filterMode = true;
        else if (id === 0x9e) {
          const c = filterColumn(r, b8, book.codepage);
          filters.set(c.colId, c);
        } else if (id === 0x87e) bad('Extended AutoFilter criteria are not supported');
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        issue(
          'legacy-' + feature,
          {
            hyperlinks: N('A cell hyperlink was not kept'),
            filters: N('Some AutoFilter settings were not kept'),
            protection: N('Some worksheet protection settings were not kept'),
            'print-settings': 'Some print settings could not be read'
          }[feature],
          e,
          feature !== 'print-settings'
        );
      }
    }
    if (protectionOptions != null) {
      const names = [
        'objects',
        'scenarios',
        'formatCells',
        'formatColumns',
        'formatRows',
        'insertColumns',
        'insertRows',
        'insertHyperlinks',
        'deleteColumns',
        'deleteRows',
        'selectLockedCells',
        'sort',
        'autoFilter',
        'pivotTables',
        'selectUnlockedCells'
      ];
      names.forEach((n, i) => (protect[n] = protectionOptions & (1 << i) ? 0 : 1));
    }
    let autoFilter = '';
    if (filterCount) {
      const name = formulas.names.find(
          n => n.name === '_xlnm._FilterDatabase' && n.scope === si && !n.invalid
        ),
        ref = name?.text?.split('!').at(-1)?.replace(/\$/g, '');
      if (ref && /^[A-Z]+\d+(?::[A-Z]+\d+)?$/.test(ref))
        autoFilter = el(
          'autoFilter',
          { ref },
          [...filters.values()]
            .filter(c => c.colId < filterCount)
            .map(c => el('filterColumn', { colId: c.colId }, c.body || undefined))
            .join('')
        );
      else
        issue(
          'legacy-filters',
          N('AutoFilter settings were not kept because their cell range could not be read')
        );
    }
    const hyperlinks = [];
    for (const link of links) {
      try {
        const { target, targetFrame, ...attrs } = link;
        if (target) attrs['r:id'] = pkg.rel(part, 'hyperlink', target, { external: true });
        attrs.tooltip = tips.get(link.ref);
        hyperlinks.push(el('hyperlink', attrs));
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        issue('legacy-hyperlinks', N('A cell hyperlink was not kept'), e);
      }
    }
    if (Object.values(hf).some(s => /(?:^|[^&])(?:&&)*&G/.test(s)))
      issue('legacy-header-footer-pictures', N('Header and footer pictures were not kept'));
    const headerFooter = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter']
      .filter(k => hf[k])
      .map(k => el(k, {}, B.xstring(hf[k])))
      .join('');
    return {
      sheetPr: { filterMode: filterMode ? 1 : null },
      pagePr: hasSetup || fit ? el('pageSetUpPr', { autoPageBreaks, fitToPage: fit ? 1 : null }) : '',
      protection: protect.sheet ? el('sheetProtection', protect) : '',
      autoFilter,
      hyperlinks: hyperlinks.length ? el('hyperlinks', {}, hyperlinks.join('')) : '',
      print:
        (Object.keys(print).length ? el('printOptions', print) : '') +
        (hasMargins ? el('pageMargins', margins) : '') +
        (hasSetup ? el('pageSetup', setup) : '') +
        (headerFooter || Object.keys(hfAttrs).length ? el('headerFooter', hfAttrs, headerFooter) : '') +
        ['rowBreaks', 'colBreaks']
          .filter(tag => breaks[tag]?.length)
          .map(tag => {
            const a = breaks[tag];
            return el(
              tag,
              { count: a.length, manualBreakCount: a.length },
              a.map(b => el('brk', b)).join('')
            );
          })
          .join(''),
      stats: {
        hyperlinks: hyperlinks.length,
        filterColumns: filters.size,
        autoFilter: !!autoFilter,
        protection: !!protect.sheet,
        print: hasSetup
      }
    };
  };
})(typeof window !== 'undefined' ? window : globalThis);
