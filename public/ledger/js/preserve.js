/* Ledger — ownership of original OOXML parts and merged workbook settings. */
(function (root) {
  'use strict';
  const L = root.L, K = L.opc, N = K.NS, P = (L.preserve = {});
  const RT = name => N.rel + '/' + name;
  const kids = el => Array.from(el?.children || []);
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  P.filterValues = c => JSON.stringify([c.values, c.blank, c.dates]);
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
    const consumed = new Set([RT('officeDocument'), ...P.parts.map(p => p[1]),
      RT('dialogsheet'), 'http://schemas.microsoft.com/office/2006/relationships/xlMacrosheet', 'http://schemas.microsoft.com/office/2006/relationships/xlIntlMacrosheet',
      ...['worksheet', 'chartsheet', 'externalLink', 'drawing', 'image', 'chart', 'hyperlink', 'comments', 'vmlDrawing', 'table', 'control', 'ctrlProp', 'oleObject', 'customProperty'].map(RT)]);
    const output = K.output(pkg, { doc: wb, format, consumes: (base, rel, type) => consumed.has(type) || L.slicers.consumes(type) ||
      type === 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment' ||
      (type === RT('queryTable') && wb.extra.queries?.some(q => q.part === rel.part)) ||
      (type === RT('pivotTable') && wb.extra.pivots?.tables.some(t => t.part === rel.part)),
      convert: base => base === 'xl/theme/theme1.xml' ? L.xlsxWrite.themeXml(wb) : null,
      merge: (base, source, data, writer) => P.merge(wb, base, source, data, writer) });
    output.bind('xl/workbook.xml', pkg?.main, 'merged');
    for (const [base, type, mode, root] of P.parts) {
      const rel = pkg?.rels(root ? '' : pkg.main).find(r => K.relationshipType(r.type) === type && !r.external);
      output.bind(base, rel?.part, base === 'xl/theme/theme1.xml' && !same(wb.themeXml, wb.extra.keepValues?.themeXml) ? 'regenerated' : mode);
    }
    return output;
  };
  P.merge = function (wb, base, source, generated, writer) {
    const fragment = K.fragment(writer.pkg.xml(source), { pkg: writer.pkg, part: source });
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
