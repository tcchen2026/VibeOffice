/* Ledger — ownership of original OOXML parts and merged workbook settings. */
(function (root) {
  'use strict';
  const L = root.L, K = L.opc, N = K.NS, P = (L.preserve = {});
  const RT = name => N.rel + '/' + name;
  const kids = el => Array.from(el?.children || []);
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  const preview2006 = 'http://schemas.microsoft.com/office/excel/2006/2';
  const previewDrawing = 'http://schemas.openxmlformats.org/drawingml/2006/3/main';
  const previewXML = P.previewXML = xml => !xml.includes(preview2006) && !xml.includes(previewDrawing) ? xml : xml
    .replace(/(\bxmlns(?::[\w.-]+)?\s*=\s*)(["'])http:\/\/schemas\.microsoft\.com\/office\/excel\/2006\/2\2/g, (_, key, quote) => key + quote + N.s + quote)
    .replace(/(\bxmlns(?::[\w.-]+)?\s*=\s*)(["'])http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/3\/main\2/g, (_, key, quote) => key + quote + N.a + quote)
    .replace(/(<\/?)([\w.-]+:)?sstItem(?=[\s/>])/g, '$1$2si');
  P.filterValues = c => JSON.stringify([c.values, c.blank, c.dates]);
  P.sheetFormatValues = sh => ({
    baseColWidth: sh.baseColW != null && sh.baseColW !== 8 ? sh.baseColW : undefined,
    defaultColWidth: sh.defColW != null ? sh.defColW : undefined,
    defaultRowHeight: sh.defRowH != null ? sh.defRowH : L.model.defaultRowPt(sh),
    customHeight: sh.defRowH != null ? '1' : undefined, zeroHeight: sh.zeroHeight ? '1' : undefined,
    outlineLevelRow: sh.outline?.levelRow || undefined, outlineLevelCol: sh.outline?.levelCol || undefined,
  });
  function keepSheetProperty(sh, el, values, tag) {
    if (!sh.wb.pkg) return;
    const ctx = { pkg: sh.wb.pkg, part: sh.extra.ooxmlPart };
    const keep = { values, fragment: el ? K.fragment(el, ctx) : null };
    // Pre-release Excel used a different schema (including different units).
    // Its format element is not a valid child of the current worksheet schema.
    if (['http://schemas.microsoft.com/office/excel/2005/8/worksheet',
      'http://schemas.microsoft.com/office/excel/2006/2'].includes(el?.namespaceURI)) {
      keep.fragment = null; keep.legacyNamespace = true;
    }
    if (keep.fragment?.record) {
      const tree = K.parse(keep.fragment.xml);
      // A wrapper that also owns columns/views cannot be emitted by this one
      // property: that would duplicate or resurrect separately edited content.
      if ([...tree.getElementsByTagName('*')].some(e => e.namespaceURI !== N.mc && e.localName !== tag)) {
        keep.fragment = K.fragment(K.parse(K.raw(el)), ctx); keep.converted = true;
      }
    }
    return keep;
  }
  function sheetPropertyXML(sh, writer, owner, tag, keep, current, generated, notice) {
    if (!keep) return generated();
    if (keep.legacyNamespace) {
      writer.loss({ id: notice.id + '-legacy:' + sh.id, where: sh.name, action: 'conversion', what: notice.legacy });
      return generated();
    }
    if (keep.converted) writer.loss({ id: notice.id + '-alternative:' + sh.id, where: sh.name, action: 'conversion', what: notice.alternative });
    const changes = {};
    for (const [name, value] of Object.entries(current)) if (!same(value, keep.values[name])) changes[name] = value;
    if (!keep.fragment) return Object.keys(changes).length ? generated(changes) : '';
    // A file's automatic row height is authoritative too. Comparing model
    // fields avoids replacing it merely because our font metrics differ.
    let xml = keep.fragment.xml;
    if (Object.keys(changes).length) {
      const tree = K.parse(xml), edits = [];
      for (const el of [tree, ...tree.getElementsByTagName('*')]) if (el.localName === tag && [N.s, 'http://purl.oclc.org/ooxml/spreadsheetml/main'].includes(el.namespaceURI)) {
        const pos = L.xmlTree.source.get(el);
        edits.push({ start: pos.start, end: pos.end, value: K.attributes(K.raw(el), changes) });
      }
      xml = K.patch(xml, edits);
    }
    return writer.emit(K.slice(keep.fragment, xml), owner);
  }
  P.keepSheetFormat = (sh, el) => { sh.extra.formatKeep = keepSheetProperty(sh, el, P.sheetFormatValues(sh), 'sheetFormatPr'); };
  P.sheetFormatXML = (sh, writer, owner) => sheetPropertyXML(sh, writer, owner, 'sheetFormatPr', sh.extra.formatKeep,
    P.sheetFormatValues(sh), () => K.attributes('<sheetFormatPr xmlns="' + N.s + '"/>', P.sheetFormatValues(sh)), {
      id: 'sheet-format', legacy: 'Worksheet size settings from a pre-release Excel format were converted to current Excel settings.',
      alternative: 'A worksheet-format alternative that also contained other worksheet settings was converted to its displayed properties.' });
  P.pageSetupValues = sh => {
    const p = sh.print, dpi = sh.extra.pageSetupKeep?.legacyNamespace && !(p.dpi > 0) ? undefined : p.dpi;
    const values = { paperSize: p.paper || 1, scale: Math.round(p.scale || 100),
      firstPageNumber: p.firstPage != null ? p.firstPage : undefined,
      fitToWidth: p.fit ? (p.fitW == null ? 0 : p.fitW) : undefined, fitToHeight: p.fit ? (p.fitH == null ? 0 : p.fitH) : undefined,
      pageOrder: p.pageOrder === 'overThenDown' ? 'overThenDown' : 'downThenOver',
      orientation: p.orientation === 'landscape' ? 'landscape' : 'portrait', blackAndWhite: p.bw ? '1' : '0',
      draft: p.draft ? '1' : '0', cellComments: p.comments || 'none', useFirstPageNumber: p.firstPage != null ? '1' : '0',
      errors: p.errors || 'displayed', horizontalDpi: dpi, verticalDpi: dpi };
    if (sh.kind === 'chartsheet') for (const key of ['scale', 'fitToWidth', 'fitToHeight', 'pageOrder', 'cellComments', 'errors']) delete values[key];
    return values;
  };
  P.keepPageSetup = (sh, el) => { sh.extra.pageSetupKeep = keepSheetProperty(sh, el, P.pageSetupValues(sh), 'pageSetup'); };
  P.pageSetupXML = (sh, writer, owner) => {
    const values = P.pageSetupValues(sh), generated = changes => {
      const attrs = changes || { ...values };
      if (!changes) for (const [key, value] of Object.entries({ paperSize: 1, scale: 100, fitToWidth: 1, fitToHeight: 1,
        pageOrder: 'downThenOver', blackAndWhite: '0', draft: '0', cellComments: 'none', useFirstPageNumber: '0', errors: 'displayed' })) {
        if (attrs[key] === value) delete attrs[key];
      }
      return K.attributes('<pageSetup xmlns="' + N.s + '"/>', attrs);
    };
    return sheetPropertyXML(sh, writer, owner, 'pageSetup', sh.extra.pageSetupKeep, values, generated, {
      id: 'page-setup', legacy: 'Page settings from a pre-release Excel format were converted to current Excel settings.',
      alternative: 'A page-setup alternative that also contained other worksheet settings was converted to its displayed properties.' });
  };
  P.parts = Object.freeze([
    ['xl/styles.xml', RT('styles'), 'regenerated'],
    ['xl/sharedStrings.xml', RT('sharedStrings'), 'regenerated'],
    ['xl/theme/theme1.xml', RT('theme'), 'opaque'],
    ['xl/metadata.xml', RT('sheetMetadata'), 'regenerated'],
    ['xl/vbaProject.bin', 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', 'opaque'],
    ['docProps/core.xml', N.pkg + '/metadata/core-properties', 'merged', true],
    ['docProps/app.xml', RT('extended-properties'), 'merged', true],
    ['docProps/custom.xml', RT('custom-properties'), 'merged', true],
  ]);
  P.values = wb => JSON.parse(JSON.stringify({ props: wb.props, date1904: wb.date1904, codeName: wb.codeName,
    filterPrivacy: !!(wb.extra.filterPrivacy || wb.extra.removePersonal), calcPr: wb.calcPr, r1c1: wb.r1c1,
    active: wb.active, firstTab: wb.firstTab, hideTabs: wb.hideTabs, hideHScroll: wb.hideHScroll, hideVScroll: wb.hideVScroll,
    tabRatio: wb.tabRatio, protection: wb.protection, names: wb.names.map(({ _ast, ...n }) => n), themeXml: wb.themeXml }));
  // Existing style indices have opaque referrers (macro sheets, caches and
  // extensions). Keep their definitions and append only styles used by edits.
  P.styles = function (wb, generated, writer) {
    const source = P.related(wb, wb.pkg?.main, 'styles'), baseline = wb.extra.styleBaseline;
    if (!source || !baseline) return { id: i => i || 0, xml: () => generated };
    let original;
    try {
      original = wb.pkg.xml(source);
      if (original?.localName !== 'styleSheet') throw new Error('The original style table has no valid root element.');
      if (original.namespaceURI === preview2006) original = K.parse(previewXML(K.raw(original)));
    } catch (error) {
      writer.loss({ id: 'styles:malformed', what: 'The damaged style table was rebuilt from the readable cell formatting.', where: source, action: 'conversion' });
      return { id: i => i || 0, xml: () => generated };
    }
    const fresh = K.parse(generated), root = K.raw(original);
    const groups = {}, appended = {}, maps = {}, modelIds = new Map();
    for (const tag of ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'dxfs']) {
      groups[tag] = kids(kids(original).find(e => e.localName === tag)); appended[tag] = []; maps[tag] = new Map();
    }
    baseline.xf.forEach((model, index) => { if (!modelIds.has(model) && same(wb.styles.list[model], baseline.list[model])) modelIds.set(model, index); });
    const generatedGroup = tag => kids(kids(fresh).find(e => e.localName === tag));
    const formatIds = new Map(groups.numFmts.map(e => [e.getAttribute('formatCode'), +e.getAttribute('numFmtId')]));
    let nextFormat = Math.max(163, ...formatIds.values());
    const numberFormat = id => {
      if (+id < 164) return +id || 0;
      const el = generatedGroup('numFmts').find(e => +e.getAttribute('numFmtId') === +id);
      if (!el) return +id;
      const code = el.getAttribute('formatCode');
      if (!formatIds.has(code)) {
        const next = ++nextFormat; formatIds.set(code, next);
        appended.numFmts.push(K.attributes(K.raw(el), { numFmtId: next }));
      }
      return formatIds.get(code);
    };
    const component = (tag, id) => {
      id = +id || 0;
      if (!maps[tag].has(id)) {
        const el = generatedGroup(tag)[id];
        if (!el) return 0;
        maps[tag].set(id, groups[tag].length + appended[tag].length);
        appended[tag].push(K.raw(el));
      }
      return maps[tag].get(id);
    };
    const xf = (el, named) => {
      const changes = {};
      for (const [attr, group] of [['fontId', 'fonts'], ['fillId', 'fills'], ['borderId', 'borders']]) if (el.hasAttribute(attr)) changes[attr] = component(group, el.getAttribute(attr));
      if (el.hasAttribute('numFmtId')) changes.numFmtId = numberFormat(el.getAttribute('numFmtId'));
      if (named && el.hasAttribute('xfId')) {
        const id = +el.getAttribute('xfId');
        if (!maps.cellStyleXfs.has(id)) {
          const node = generatedGroup('cellStyleXfs')[id];
          maps.cellStyleXfs.set(id, groups.cellStyleXfs.length + appended.cellStyleXfs.length);
          appended.cellStyleXfs.push(node ? xf(node, false) : '<xf/>');
        }
        changes.xfId = maps.cellStyleXfs.get(id);
      }
      return K.attributes(K.raw(el), changes);
    };
    return {
      id(i = 0, keep) {
        if (keep?.source === wb.pkg.id && keep.style != null && keep.modelStyle === i && same(wb.styles.list[i], baseline.list[i])) return keep.style;
        if (!modelIds.has(i)) {
          const el = generatedGroup('cellXfs')[i] || generatedGroup('cellXfs')[0];
          modelIds.set(i, groups.cellXfs.length + appended.cellXfs.length);
          appended.cellXfs.push(xf(el, true));
        }
        return modelIds.get(i);
      },
      xml() {
        const replacements = {};
        if (wb.extra.previewThemeColors) {
          const colors = kids(original).find(e => e.localName === 'colors');
          if (colors) replacements['{' + N.s + '}colors'] = kids(colors).some(e => e.localName !== 'themeColors') ?
            K.mergeBag(K.raw(colors), { ['{' + N.s + '}themeColors']: '' }) : [];
        }
        for (const [tag, added] of Object.entries(appended)) if (added.length) {
          const el = kids(original).find(e => e.localName === tag);
          const raw = el ? K.raw(el) : '<' + tag + ' xmlns="' + N.s + '"/>';
          replacements['{' + N.s + '}' + tag] = K.attributes(K.mergeBag(raw, { __append: added }), { count: groups[tag].length + added.length });
        }
        // Differential formats are indexed directly by the model. Preserve the
        // raw definition unless that particular format was changed.
        const dxfs = generatedGroup('dxfs'), kept = groups.dxfs;
        if (!same(wb.dxfs, baseline.dxfs)) {
          const values = dxfs.map((el, i) => same(wb.dxfs[i], baseline.dxfs[i]) && kept[i] ? K.raw(kept[i]) : K.raw(el));
          replacements['{' + N.s + '}dxfs'] = '<dxfs xmlns="' + N.s + '" count="' + values.length + '">' + values.join('') + '</dxfs>';
        }
        return K.partXML(wb.pkg.text(source), K.merge(root, replacements, 's:CT_Stylesheet'));
      },
    };
  };
  P.related = (wb, owner, type) => wb.pkg?.rels(owner).find(r => !r.external && K.relationshipType(r.type) === RT(type))?.part;
  P.sheetContent = sh => JSON.stringify({
    rows: sh.rows.map(row => row && { ht: row.ht, s: row.s, hidden: row.hidden, level: row.level, collapsed: row.collapsed,
      cells: row.cells.map(c => c && { v: c.f ? undefined : c.v, f: c.f, s: c.s, rt: c.rt, dt: c.dt, af: c.af }) }),
    cols: sh.cols, merges: sh.merges, print: sh.print, tabColor: sh.tabColor,
  });
  P.begin = function (wb, format) {
    const pkg = wb.pkg;
    const previewBook = pkg?.main && pkg.xml(pkg.main)?.namespaceURI === preview2006;
    const previewTheme = /http:\/\/schemas\.microsoft\.com\/office\/officeart\/2005\/8\/oartml/.test(wb.themeXml || '');
    const consumed = new Set([RT('officeDocument'), ...P.parts.map(p => p[1]),
      RT('dialogsheet'), 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet', 'http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet',
      ...['worksheet', 'chartsheet', 'externalLink', 'drawing', 'image', 'chart', 'hyperlink', 'comments', 'vmlDrawing', 'table', 'control', 'ctrlProp', 'oleObject', 'customProperty'].map(RT)]);
    const output = K.output(pkg, { doc: wb, format, consumes: (base, rel, type) => consumed.has(type) || L.slicers.consumes(type) ||
      type === 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment' ||
      (type === RT('queryTable') && wb.extra.queries?.some(q => q.part === rel.part)) ||
      (type === RT('pivotTable') && wb.extra.pivots?.tables.some(t => t.part === rel.part)),
      convert: base => base === 'xl/theme/theme1.xml' ? L.xlsxWrite.themeXml(wb) : null,
      merge: (base, source, data, writer) => P.merge(wb, base, source, data, writer),
      audit: writer => {
        if (!previewBook && !previewTheme) return;
        for (const [part, data] of writer.parts) {
          if (!part.startsWith('xl/') || !part.endsWith('.xml')) continue;
          let text = typeof data === 'string' ? data : L.xml.decode(data);
          if (previewBook && (text.includes(preview2006) || text.includes(previewDrawing))) {
            const converted = previewXML(text);
            if (converted !== text) { writer.put(part, converted); text = converted; }
          }
          if (previewTheme && /\/theme\//.test(part) && text.includes('http://schemas.microsoft.com/office/officeart/2005/8/oartml')) writer.put(part, L.xlsxWrite.themeXml(wb));
          if (previewBook && wb.extra.previewThemeColors && /\/theme\//.test(part)) {
            const scheme = [...K.parse(text).getElementsByTagName('*')].find(e => e.localName === 'clrScheme');
            if (scheme) {
              const names = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'], replacements = {};
              names.forEach((name, i) => { const color = wb.theme.colors[i]; if (color) replacements['{' + N.a + '}' + name] = '<a:' + name + ' xmlns:a="' + N.a + '"><a:srgbClr val="' + color.slice(-6) + '"/></a:' + name + '>'; });
              const pos = L.xmlTree.source.get(scheme);
              writer.put(part, K.hoistNamespaces(K.patch(text, [{ start: pos.start, end: pos.end, value: K.mergeBag(K.raw(scheme), replacements) }])));
            }
          }
        }
        if (previewBook) writer.loss({ id: 'preview-workbook-schema', where: pkg.main, action: 'conversion', what: 'Workbook settings, styles, theme colours, shared text and calculation metadata from a pre-release Excel format were converted to current Excel markup.' });
        if (previewTheme) writer.loss({ id: 'preview-theme', where: 'Workbook theme', action: 'conversion', what: 'A pre-release Excel theme was rebuilt from its readable colours and fonts; unsupported theme details were removed.' });
      } });
    output.bind('xl/workbook.xml', pkg?.main, 'merged');
    for (const sh of wb.sheets) if (sh.extra.invalidVisibility) output.writer.loss({
      id: 'sheet-visibility:' + sh.id, where: sh.name, action: 'conversion',
      what: 'An invalid worksheet visibility setting (' + sh.extra.invalidVisibility + ') was replaced with ' + sh.state + '. This can change which sheets are printed.' });
    for (const [base, type, mode, root] of P.parts) {
      const rel = pkg?.rels(root ? '' : pkg.main).find(r => K.relationshipType(r.type) === type && !r.external);
      output.bind(base, rel?.part, base === 'xl/theme/theme1.xml' && !same(wb.themeXml, wb.extra.keepValues?.themeXml) ? 'regenerated' : mode);
    }
    return output;
  };
  P.merge = function (wb, base, source, generated, writer) {
    let root = writer.pkg.xml(source);
    if (base === 'xl/workbook.xml' && root.namespaceURI === preview2006) root = K.parse(previewXML(K.raw(root)));
    const fragment = K.fragment(root, { pkg: writer.pkg, part: source });
    const old = wb.extra.keepValues || {}, current = P.values(wb), fresh = K.parse(generated);
    let xml = fragment.xml;
    if (base === 'xl/workbook.xml') {
      const tree = K.parse(xml), replacements = {};
      const set = (tag, value) => { replacements['{' + N.s + '}' + tag] = value; };
      const generatedChild = tag => kids(fresh).filter(e => e.localName === tag).map(K.raw);
      // These lists are generated by the model. Other workbook children (data
      // connections, pivot caches, custom views, extension lists) stay in place.
      for (const tag of ['sheets', 'externalReferences']) set(tag, generatedChild(tag));
      const pivotCaches = L.pivots.workbook(wb, tree, writer);
      if (pivotCaches !== undefined) set('pivotCaches', pivotCaches);
      if (!same(current.names, old.names)) set('definedNames', generatedChild('definedNames'));
      else { const names = L.slicers.definedNames(tree, writer.slicers); if (names !== undefined) set('definedNames', names); }
      if (!same(current.protection, old.protection)) set('workbookProtection', generatedChild('workbookProtection'));
      const patch = (tag, fields, forced = {}) => {
        const before = kids(tree).find(e => e.localName === tag), after = kids(fresh).find(e => e.localName === tag);
        const changes = { ...forced };
        for (const [key, name] of Object.entries(fields)) if (!same(current[key], old[key])) changes[name] = after?.getAttribute(name) ?? null;
        if (!Object.keys(changes).length) return;
        set(tag, before ? K.attributes(K.raw(before), changes) : after ? K.raw(after) : []);
      };
      patch('workbookPr', { date1904: 'date1904', codeName: 'codeName', filterPrivacy: 'filterPrivacy' });
      if (!same(current.calcPr, old.calcPr) || !same(current.r1c1, old.r1c1)) set('calcPr', generatedChild('calcPr'));
      else patch('calcPr', {}, { fullCalcOnLoad: '1' });
      const viewFields = { active: 'activeTab', firstTab: 'firstSheet', hideTabs: 'showSheetTabs', hideHScroll: 'showHorizontalScroll', hideVScroll: 'showVerticalScroll', tabRatio: 'tabRatio' };
      if (Object.keys(viewFields).some(k => !same(current[k], old[k]))) {
        const views = kids(tree).find(e => e.localName === 'bookViews'), view = kids(views).find(e => e.localName === 'workbookView');
        const next = kids(kids(fresh).find(e => e.localName === 'bookViews')).find(e => e.localName === 'workbookView');
        if (view && next) {
          const changes = {};
          for (const [key, name] of Object.entries(viewFields)) if (!same(current[key], old[key])) changes[name] = next.getAttribute(name);
          // Only the first window is the editor's active view; other windows stay.
          const bag = K.raw(views), root = K.parse(bag), first = kids(root).find(e => e.localName === 'workbookView'), pos = L.xmlTree.source.get(first);
          set('bookViews', K.patch(bag, [{ start: pos.start, end: pos.end, value: K.attributes(K.raw(first), changes) }]));
        } else set('bookViews', generatedChild('bookViews'));
      }
      xml = K.merge(xml, replacements, 's:CT_Workbook');
      xml = L.slicers.workbook(fragment, xml, writer);
    } else if (base === 'docProps/custom.xml') {
      const named = props => Object.fromEntries((props?.custom || []).map(p => [p.name, { type: p.type, value: p.value }]));
      xml = K.customProperties(xml, generated, named(wb.props), named(old.props));
    } else {
      const owned = base === 'docProps/core.xml' ? { title: 'title', subject: 'subject', creator: 'creator', keywords: 'keywords', description: 'description', lastModifiedBy: 'lastModifiedBy', category: 'category', created: 'created', modified: 'modified' } : { manager: 'Manager', company: 'Company' };
      const props = wb.extra.removePersonal || wb.extra.filterPrivacy ? { ...wb.props, creator: '', lastModifiedBy: ' ', manager: '' } : wb.props;
      xml = K.properties(xml, generated, props, old.props, owned);
    }
    return writer.emit(K.slice(fragment, xml), source);
  };
})(typeof window !== 'undefined' ? window : globalThis);
