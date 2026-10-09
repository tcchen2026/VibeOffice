/* Ledger — retained worksheet extensions belong to their rules and sparkline groups. */
(function (root) {
  'use strict';
  const L = root.L, K = L.opc, E = (L.sheetExtensions = {}), N = K.NS;
  const X = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
  const XM = 'http://schemas.microsoft.com/office/excel/2006/main';
  const kids = e => Array.from(e?.children || []), all = e => e ? [e, ...e.getElementsByTagName('*')] : [];
  const key = e => '{' + e.namespaceURI + '}' + e.localName, at = (e, n) => e?.getAttribute(n);
  const clone = x => JSON.parse(JSON.stringify(x)), same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  const values = ({ extKeep, baseKeep, ...o }) => clone(o);
  const fragment = (el, sh) => K.fragment(el, { pkg: sh.wb.pkg, part: sh.extra.ooxmlPart });
  const captured = new WeakSet();
  const types = { conditionalFormattings: 'conditionalFormatting', dataValidations: 'dataValidation', sparklineGroups: 'sparklineGroup' };
  E.NS = X; E.XM = XM;
  E.copy = o => K.duplicate(o, { guidAttributes: ['uid'], guidFields: ['x14id'] });
  E.keep = function (o, el, sh, group) {
    captured.add(el);
    o.extKeep = { fragment: fragment(el, sh), values: values(o) };
    if (group) o.extKeep.group = { fragment: fragment(group, sh), ranges: clone(sh.cf.find(c => c.rules.includes(o))?.ranges || []) };
    if (o.dxf != null && sh.wb.dxfs[o.dxf]) o.extKeep.dxf = clone(sh.wb.dxfs[o.dxf]);
  };
  E.base = (o, el, sh) => { if (o.x14id) o.baseKeep = { fragment: fragment(el, sh), values: values(o) }; };
  E.read = function (sh, ext) {
    // A skeleton retains attributes and unconsumed children even in a known URI.
    // Model-owned items are inserted anew; deleting an item cannot resurrect it.
    sh.extra.extensions = kids(ext).map(el => {
      const f = fragment(el, sh), tree = K.parse(f.xml), replacements = {};
      for (const c of kids(el)) if (c.namespaceURI === X && types[c.localName]) {
        const raw = kids(tree).find(e => e.localName === c.localName), remaining = [];
        for (const [i, item] of kids(c).entries()) {
          const tokenItem = kids(raw)[i];
          if (item.localName !== types[c.localName] || item.namespaceURI !== X) continue;
          if (c.localName === 'conditionalFormattings' && item.localName === 'conditionalFormatting') {
            const rules = kids(item).filter(e => e.localName === 'cfRule');
            if (rules.some(e => captured.has(e))) {
              const unowned = rules.filter(e => !captured.has(e));
              if (unowned.length) remaining.push(K.mergeBag(K.raw(tokenItem), { ['{' + X + '}cfRule']: kids(tokenItem).filter(e => e.localName === 'cfRule' && unowned.some(u => at(u, 'id') === at(e, 'id'))).map(K.raw) }));
            } else remaining.push(K.raw(tokenItem));
          } else if (!captured.has(item)) remaining.push(K.raw(tokenItem));
        }
        replacements[key(raw)] = K.mergeBag(K.raw(raw), { ['{' + X + '}' + types[c.localName]]: [], __append: remaining });
      }
      return { uri: at(el, 'uri'), fragment: K.slice(f, K.mergeBag(f.xml, replacements)), original: f };
    });
    for (const cf of sh.cf) for (const r of cf.rules) if (r.baseKeep) r.baseKeep.values = values(r);
  };
  E.begin = function (wb, writer) {
    const state = { writer, ids: new Set(), rules: new WeakMap(), uids: new WeakMap() };
    const allocate = original => {
      let id = original;
      if (!id || state.ids.has(id.toUpperCase())) do { id = K.guid(); } while (state.ids.has(id));
      state.ids.add(id.toUpperCase()); return id;
    };
    state.uid = (o, id) => {
      if (!state.uids.has(o)) state.uids.set(o, new Map());
      const map = state.uids.get(o); if (!map.has(id)) map.set(id, allocate(id)); return map.get(id);
    };
    for (const sh of wb.sheets) for (const cf of sh.cf) for (const r of cf.rules) state.rules.set(r, allocate(r.x14id));
    return state;
  };
  E.id = (r, state) => state.rules.get(r);
  E.bar = r => !!r.bar && (!!r.x14id || ['gradient', 'border', 'direction', 'negColor', 'borderColor', 'axisColor', 'axis'].some(k => r.bar[k] != null));
  const children = (xml, ns, name) => kids(K.parse(xml)).filter(e => e.namespaceURI === ns && e.localName === name).map(K.raw);
  function properties(original, generated, before, after, attrs, childFields, forced = {}) {
    const fresh = K.parse(generated), changes = { ...forced }, replacements = {};
    for (const [name, field] of Object.entries(attrs)) if (!same(before?.[field], after?.[field])) changes[name] = at(fresh, name);
    for (const [name, fields] of Object.entries(childFields)) if (fields.some(f => !same(before?.[f], after?.[f]))) {
      replacements[name] = kids(fresh).filter(e => key(e) === name).map(K.raw);
    }
    // Insert regenerated children in the order of the fresh Office vocabulary.
    let xml = K.attributes(original, changes);
    const known = Object.keys(childFields), order = known;
    for (const name of known) if (name in replacements) xml = K.mergeBag(xml, { [name]: replacements[name] });
    const tree = K.parse(xml), ordered = kids(tree).filter(e => known.includes(key(e)));
    if (Object.keys(replacements).length && ordered.length > 1) {
      const sorted = ordered.slice().sort((a, b) => order.indexOf(key(a)) - order.indexOf(key(b)));
      if (ordered.some((e, i) => e !== sorted[i])) {
        const source = L.xmlTree.source;
        xml = K.patch(xml, ordered.map((e, i) => ({ start: source.get(e).start, end: source.get(e).end, value: K.raw(sorted[i]) })));
      }
    }
    return xml;
  }
  function emit(o, f, xml, state, owner) {
    return state.writer.emit(K.slice(f, K.remapGuids(xml, id => state.uid(o, id), ['uid'])), owner);
  }
  E.rule = function (r, generated, state, owner, wb, base = false) {
    const keep = base ? r.baseKeep : r.extKeep; if (!keep) return generated;
    const before = keep.values, after = values(r), ns = base ? N.s : X, f = keep.fragment;
    if (r.type !== before.type) return generated;
    const attrs = { type: 'type', stopIfTrue: 'stop', operator: 'op', text: 'text', timePeriod: 'period', rank: 'rank', percent: 'percent', bottom: 'bottom', aboveAverage: 'below', equalAverage: 'equal', stdDev: 'stdDev' };
    const fields = { ['{' + (base ? N.s : XM) + '}' + (base ? 'formula' : 'f')]: ['f'],
      ['{' + ns + '}colorScale']: ['scale'], ['{' + ns + '}dataBar']: [], ['{' + ns + '}iconSet']: ['icons'],
      ['{' + ns + '}dxf']: [], ['{' + ns + '}extLst']: [] };
    const fresh = K.parse(generated), forced = {};
    if (base) forced.priority = at(fresh, 'priority');
    else { forced.id = E.id(r, state); if (at(K.parse(f.xml), 'priority') != null || !r.bar) forced.priority = at(fresh, 'priority'); }
    let xml = properties(f.xml, generated, before, after, attrs, fields, forced);
    const db = kids(K.parse(xml)).find(e => e.localName === 'dataBar');
    const next = kids(fresh).find(e => e.localName === 'dataBar');
    if (db && next && before.bar && r.bar) {
      let bar = properties(K.raw(db), K.raw(next), before.bar, r.bar,
        { minLength: 'minLength', maxLength: 'maxLength', showValue: 'showValue', gradient: 'gradient', border: 'border', direction: 'direction', axisPosition: 'axis' },
        { ['{' + ns + '}cfvo']: ['cfvo'], ['{' + ns + '}' + (base ? 'color' : 'fillColor')]: ['color'],
          ['{' + ns + '}borderColor']: ['borderColor', 'border'], ['{' + ns + '}negativeFillColor']: ['negColor'],
          ['{' + ns + '}negativeBorderColor']: [], ['{' + ns + '}axisColor']: ['axisColor', 'axis'], ['{' + ns + '}extLst']: [] });
      if (base && ['minLength', 'maxLength'].some(k => !same(before.bar[k], r.bar[k]))) {
        bar = K.attributes(bar, { minLength: at(next, 'minLength'), maxLength: at(next, 'maxLength') });
      }
      if (!base && !same(before.bar.negColor, r.bar.negColor)) bar = K.attributes(bar, { negativeBarColorSameAsPositive: '0' });
      if (!base && before.bar.border !== r.bar.border && !r.bar.border) bar = K.mergeBag(bar, { ['{' + X + '}negativeBorderColor']: [] });
      xml = K.mergeBag(xml, { ['{' + ns + '}dataBar']: bar });
    } else if (!same(before.bar, r.bar)) xml = K.mergeBag(xml, { ['{' + ns + '}dataBar']: next ? K.raw(next) : [] });
    if (!base && !same(keep.dxf, wb.dxfs[r.dxf])) xml = K.mergeBag(xml, { ['{' + X + '}dxf']: children(generated, X, 'dxf') });
    // The base and extension use the same identity, including after sheet copies.
    if (base) {
      const changes = [];
      for (const el of all(K.parse(xml))) if (el.namespaceURI === X && el.localName === 'id') {
        const p = L.xmlTree.source.get(el);
        changes.push({ start: p.start, end: p.end, value: '<x14:id xmlns:x14="' + X + '">' + L.xml.esc(E.id(r, state)) + '</x14:id>' });
      }
      xml = K.patch(xml, changes);
    }
    return emit(r, f, xml, state, owner);
  };
  E.groups = function (items, state, owner) {
    const groups = new Map();
    for (const { cf, r, xml } of items) {
      const keep = r.extKeep?.group, ranges = L.model.sqref ? L.model.sqref(cf.ranges) : cf.ranges.map(g => L.formula.rangeName(g)).join(' ');
      const k = (keep?.fragment.xml || '') + ':' + ranges;
      if (!groups.has(k)) groups.set(k, { keep, cf, items: [] });
      groups.get(k).items.push(xml);
    }
    return [...groups.values()].map(({ keep, cf, items }) => {
      const ranges = cf.ranges.map(g => L.formula.rangeName(g)).join(' '), sq = '<xm:sqref xmlns:xm="' + XM + '">' + L.xml.esc(ranges) + '</xm:sqref>';
      if (!keep) return '<x14:conditionalFormatting xmlns:x14="' + X + '" xmlns:xm="' + XM + '">' + items.join('') + sq + '</x14:conditionalFormatting>';
      const replacements = { ['{' + X + '}cfRule']: items };
      if (!same(cf.ranges, keep.ranges)) replacements['{' + XM + '}sqref'] = sq;
      return state.writer.emit(K.slice(keep.fragment, K.mergeBag(keep.fragment.xml, replacements)), owner);
    }).join('');
  };
  E.validation = function (d, generated, state, owner) {
    const keep = d.extKeep; if (!keep) return generated;
    const attrs = { type: 'type', operator: 'op', allowBlank: 'allowBlank', showDropDown: 'showDrop', showInputMessage: 'showInput', showErrorMessage: 'showError', errorStyle: 'errorStyle', imeMode: 'ime' };
    for (const k of ['promptTitle', 'prompt', 'errorTitle', 'error']) attrs[k] = k;
    const xml = properties(keep.fragment.xml, generated, keep.values, d, attrs,
      { ['{' + X + '}formula1']: ['f1'], ['{' + X + '}formula2']: ['f2'], ['{' + XM + '}sqref']: ['ranges'] });
    return emit(d, keep.fragment, xml, state, owner);
  };
  E.sparkline = function (s, generated, state, owner) {
    const keep = s.extKeep; if (!keep) return generated;
    const attrs = {}, fields = {};
    for (const k of ['manualMax', 'manualMin', 'lineWeight', 'type', 'dateAxis', 'displayEmptyCellsAs', 'markers', 'high', 'low', 'first', 'last', 'negative', 'displayXAxis', 'displayHidden', 'minAxisType', 'maxAxisType', 'rightToLeft']) attrs[k] = k;
    for (const [name, field] of Object.entries({ colorSeries: 'color', colorNegative: 'negColor', colorAxis: 'axisColor', colorMarkers: 'markerColor', colorFirst: 'firstColor', colorLast: 'lastColor', colorHigh: 'highColor', colorLow: 'lowColor' })) fields['{' + X + '}' + name] = [field];
    fields['{' + XM + '}f'] = []; fields['{' + X + '}sparklines'] = ['items'];
    return emit(s, keep.fragment, properties(keep.fragment.xml, generated, keep.values, s, attrs, fields), state, owner);
  };
  E.sheet = function (sh, generated, state, owner) {
    const entries = kids(K.parse(generated || '<extLst/>')), fresh = new Map(entries.map(e => [at(e, 'uri'), K.raw(e)])), out = [];
    for (const entry of sh.extra.extensions || []) {
      const f = entry.fragment, tree = K.parse(f.xml), next = fresh.get(entry.uri), newTree = next && K.parse(next);
      fresh.delete(entry.uri);
      const replacements = {}; let known = false;
      for (const c of kids(tree)) if (c.namespaceURI === X && types[c.localName]) {
        known = true;
        const newList = kids(newTree).find(e => key(e) === key(c));
        let list = K.mergeBag(K.raw(c), { __append: kids(newList).map(K.raw) });
        if (c.localName === 'dataValidations') list = K.attributes(list, { count: kids(K.parse(list)).filter(e => e.localName === 'dataValidation').length });
        const originalList = kids(K.parse(entry.original.xml)).find(e => key(e) === key(c));
        replacements[key(c)] = kids(K.parse(list)).length ? list : kids(originalList).length ? [] : K.raw(c);
      }
      // Slicer/timeline entries are owned by their separate source-aware writer.
      if (!known && kids(tree).some(e => /^(slicerList|timelineRefs)$/.test(e.localName))) { if (next) out.push(next); continue; }
      const xml = K.mergeBag(f.xml, replacements);
      if (kids(K.parse(xml)).length) out.push(state.writer.emit(K.slice(f, xml), owner));
    }
    out.push(...fresh.values());
    return out.length ? '<extLst>' + out.join('') + '</extLst>' : '';
  };
})(typeof window !== 'undefined' ? window : globalThis);
