/* Ledger — slicer/timeline views retain their frames, definitions and source caches. */
(function (root) {
  'use strict';
  const L = root.L, K = L.opc, S = (L.slicers = {}), N = K.NS;
  const kids = el => Array.from(el?.children || []), all = el => el ? [el, ...Array.from(el.getElementsByTagName('*'))] : [];
  const at = (el, name) => el?.getAttribute(name), clone = x => JSON.parse(JSON.stringify(x));
  const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
  const guid = () => '{' + root.crypto.randomUUID().toUpperCase() + '}';
  const viewRel = type => /\/(slicer|timeline)$/.test(type || '');
  const cacheRel = type => /\/(slicerCache|timelineCache)$/.test(type || '');
  const rid = el => el?.getAttributeNS(N.rel, 'id') || el?.getAttributeNS(N.strictRel, 'id');
  const payload = el => /\/(slicer|timeslicer)$/.test(el.namespaceURI || '') && ['slicer', 'timeslicer'].includes(el.localName);
  const values = d => ({ xml: d.xml, text: d.text, dirty: d.dirty, media: d.media, geom: d.geom,
    fill: d.fill, line: d.line, lineW: d.lineW, font: d.font });
  const intact = d => same(values(d), d.slicerKeep?.values);
  const sheet = (wb, id) => wb.sheets.find(s => s.id === id);
  const byName = (list, name) => list.find(c => c.name === name);
  const xmlPart = (pkg, part) => { try { return pkg.xml(part); } catch (_) { return null; } };
  S.consumes = type => viewRel(type) || cacheRel(type);

  S.captureDrawing = function (tree, sh, part) {
    const records = new WeakMap(), pkg = sh.wb.pkg;
    if (!pkg) return records;
    for (const anchor of all(tree).filter(e => /Anchor$/.test(e.localName))) {
      const views = all(anchor).filter(payload); if (!views.length) continue;
      const nv = all(anchor).find(e => e.localName === 'cNvPr');
      records.set(anchor, { source: pkg.id, part, sheet: sh.id, id: +at(nv, 'id'),
        names: [...new Set(views.map(e => at(e, 'name')))], fragment: K.fragment(anchor, { pkg, part }) });
    }
    return records;
  };
  S.attach = (d, record) => {
    d.id = record.id; d.name ||= record.names[0];
    d.slicerKeep = { ...record, anchor: clone(d.anchor), values: clone(values(d)) };
  };
  S.copy = function (d, toSheet) {
    if (!d.slicerKeep) return;
    d.slicerKeep = { ...clone(d.slicerKeep), copy: guid(), copySheet: toSheet?.id };
  };
  S.read = function (wb) {
    const pkg = wb.pkg; if (!pkg) return;
    const caches = [], groups = [];
    for (const rel of pkg.rels(pkg.main).filter(r => cacheRel(r.type) && !r.external)) {
      const tree = xmlPart(pkg, rel.part); if (!tree) continue;
      const table = all(tree).find(e => e.localName === 'tableSlicerCache');
      const modelTable = table && wb.tables.find(t => t.id === +at(table, 'tableId'));
      const cacheId = all(tree).find(e => e.hasAttribute('pivotCacheId'))?.getAttribute('pivotCacheId');
      const pivot = cacheId && wb.extra.pivots?.caches.find(c => all(pkg.xml(c.part)).some(e => at(e, 'pivotCacheId') === cacheId));
      const links = all(tree).filter(e => e.localName === 'pivotTable').map(e => {
        const sh = wb.sheets.find(s => s.sheetId === +at(e, 'tabId'));
        const p = wb.extra.pivots?.tables.find(t => t.sheet === sh?.id && t.name === at(e, 'name'));
        return { name: at(e, 'name'), tabId: at(e, 'tabId'), part: p?.part, sheet: sh?.id };
      });
      caches.push({ part: rel.part, name: at(tree, 'name'), rel, pivot: pivot?.id, field: pivot?.fields.indexOf(at(tree, 'sourceName')), links,
        table: modelTable && { part: modelTable.ooxmlPart, sheet: modelTable.sheet.id, column: +at(table, 'column') } });
    }
    for (const sh of wb.sheets) {
      const rels = pkg.rels(sh.extra.ooxmlPart || '').filter(r => viewRel(r.type) && !r.external);
      if (!rels.length) continue;
      const root = pkg.xml(sh.extra.ooxmlPart), extList = kids(root).find(e => e.localName === 'extLst');
      for (const rel of rels) {
        const tree = xmlPart(pkg, rel.part); if (!tree) continue;
        const ext = kids(extList).find(e => all(e).some(c => rid(c) === rel.id));
        groups.push({ sheet: sh.id, part: rel.part, rel, ext: ext && K.fragment(ext, { pkg, part: sh.extra.ooxmlPart }),
          definitions: kids(tree).filter(e => ['slicer', 'timeline'].includes(e.localName)).map(e => ({ name: at(e, 'name'), cache: at(e, 'cache') })) });
      }
    }
    if (!caches.length && !groups.length) return;
    wb.extra.slicers = { caches, groups };
    for (const sh of wb.sheets) for (const d of sh.drawings) if (d.slicerKeep) {
      d.slicerKeep.views = d.slicerKeep.names.map(name => {
        const g = groups.find(g => g.sheet === sh.id && byName(g.definitions, name));
        return g && { ...byName(g.definitions, name), part: g.part, ext: g.ext, rel: g.rel };
      }).filter(Boolean);
    }
    for (const g of groups) g.bound = wb.sheets.flatMap(sh => sh.drawings.flatMap(d =>
      (d.slicerKeep?.views || []).filter(v => v.part === g.part).map(v => v.name)));
  };
  S.begin = function (wb, pack) {
    const w = pack.writer, original = wb.extra.slicers, p = L.pivots.current(wb);
    const state = { wb, pack, caches: [], groups: [], frames: new Map(), omitted: new Set() };
    w.slicers = state;
    const aliveTables = (p?.tables || []).filter(t => !t.drop), aliveCaches = (p?.caches || []).filter(c => !c.drop);
    for (const c of original?.caches || []) {
      const table = c.table && wb.tables.find(t => t.ooxmlPart === c.table.part && t.sheet.id === c.table.sheet);
      const pivot = aliveCaches.find(p => p.id === c.pivot);
      const linked = c.links.flatMap(link => link.part ? aliveTables.filter(t => t.part === link.part).map(t => ({ ...link, table: t })) : [link]);
      const missing = c.table && (!table || !table.columns.some(col => col.id === c.table.column)) ||
        c.pivot && (!pivot || c.field >= 0 && pivot.columns && !pivot.columns.includes(c.field)) || c.links.length && !linked.length;
      state.caches.push({ ...c, linked, modelTable: table, drop: !!missing, target: w.target(wb.pkg, c.part) });
    }
    for (const sh of wb.sheets) for (const d of sh.drawings) if (d.slicerKeep) {
      const keep = d.slicerKeep, entry = { d, sh, status: 'keep', names: new Map() }; state.frames.set(d, entry);
      if (!intact(d) || keep.source !== wb.pkg?.id || !keep.views?.length) {
        entry.status = 'convert';
        w.loss({ id: 'slicer-conversion:' + sh.id + ':' + d.id, what: keep.source !== wb.pkg?.id ?
          'The pasted filter has no source table or pivot in this workbook. Its displayed shape is kept.' :
          'The edited slicer or timeline was converted to its displayed shape.', where: sh.name, action: 'conversion' });
        continue;
      }
      if (keep.views.some(v => {
        const c = byName(state.caches, v.cache);
        return c?.drop && !(c.table && keep.copySheet === sh.id && wb.tables.some(t => t.ooxmlPart === c.table.part && t.sheet === sh));
      })) { entry.status = 'omit'; continue; }
      for (const view of keep.views) {
        const name = keep.copy ? view.name + '_Copy' + keep.copy.slice(1, 9) : view.name;
        entry.names.set(view.name, name);
        let cache = byName(state.caches, view.cache);
        if (cache?.table && keep.copySheet === sh.id) {
          const table = wb.tables.find(t => t.ooxmlPart === cache.table.part && t.sheet === sh);
          if (table) {
            const copyName = cache.name + '_Copy' + keep.copy.slice(1, 9);
            let copied = byName(state.caches, copyName);
            if (!copied) { copied = { ...cache, name: copyName, copy: keep.copy, modelTable: table, drop: false,
              target: w.name('xl/slicerCaches', 'slicerCache', 'xml') }; state.caches.push(copied); }
            cache = copied;
          }
        }
        const key = sh.id + ':' + view.part;
        let group = state.groups.find(g => g.key === key);
        if (!group) {
          const originalSheet = (original?.groups || []).find(g => g.part === view.part)?.sheet;
          const slash = view.part.lastIndexOf('/');
          group = { ...view, key, sheet: sh, members: [], target: sh.id === originalSheet ? w.target(wb.pkg, view.part) :
            w.name(view.part.slice(0, slash), view.part.slice(slash + 1).replace(/\d*\.xml$/, ''), 'xml') };
          state.groups.push(group);
        }
        group.members.push({ name, original: view.name, cache: cache?.name || view.cache, uid: keep.copy });
      }
    }
    // Definitions without a readable drawing have no editable owner to delete.
    // Retain them while their worksheet and source still exist.
    for (const g of original?.groups || []) {
      const sh = sheet(wb, g.sheet); if (!sh) continue;
      const members = g.definitions.filter(v => !g.bound.includes(v.name) && !byName(state.caches, v.cache)?.drop)
        .map(v => ({ ...v, original: v.name }));
      if (!members.length) continue;
      const key = sh.id + ':' + g.part, existing = state.groups.find(g => g.key === key);
      if (existing) existing.members.push(...members);
      else state.groups.push({ ...g, key, sheet: sh, members, target: w.target(wb.pkg, g.part) });
    }
    // Only caches that originally had visible views are pruned by view deletion.
    for (const c of state.caches) {
      const hadView = original?.groups.some(g => g.definitions.some(v => v.cache === c.name));
      if (hadView && !state.groups.some(g => g.members.some(v => v.cache === c.name))) c.drop = true;
      if (c.drop) { state.omitted.add(c.part); w.omit(c.part, 'The slicer/timeline source or all its views were deleted.'); }
    }
    for (const g of original?.groups || []) if (!state.groups.some(next => next.part === g.part && next.target === w.target(wb.pkg, g.part))) {
      state.omitted.add(g.part); w.omit(g.part, 'The slicer/timeline views on this worksheet were deleted.');
    }
    return state;
  };
  S.frame = function (d, state, owner, id) {
    const entry = state.frames.get(d); if (!entry || entry.status === 'convert') return null;
    if (entry.status === 'omit') return '';
    const keep = d.slicerKeep;
    let f = keep.fragment;
    let xml = f.xml.replace(/\u0001id:(\d+)\u0001/g, (token, n) => f.ids[+n].kind === 'shape' && +f.ids[+n].id === keep.id ? String(id) : token);
    if (!same(d.anchor, keep.anchor)) xml = K.setBox(xml, d.anchor.type === 'two' ? { from: d.anchor.from, to: d.anchor.to } : d.anchor);
    if (keep.copy) {
      const patches = [];
      for (const e of all(K.parse(xml))) {
        const attrs = payload(e) ? { name: entry.names.get(at(e, 'name')) } : e.localName === 'creationId' ? { id: keep.copy } : {};
        if (Object.keys(attrs).length) { const pos = L.xmlTree.source.get(e); patches.push({ start: pos.start, end: pos.end, value: K.attributes(K.raw(e), attrs) }); }
      }
      xml = K.patch(xml, patches);
    }
    f = K.slice(f, xml); if (keep.copy) f = K.duplicate(f);
    return state.pack.writer.emit(f, owner);
  };
  S.sheet = function (sh, generated, state, owner) {
    const extensions = new Map(), w = state.pack.writer;
    for (const group of state.groups.filter(g => g.sheet === sh)) {
      if (!group.ext) continue;
      const rId = w.rels(owner).add(group.rel.type, K.relative(owner, group.target), false, owner === group.ext.part ? group.rel.id : undefined);
      const f = group.ext, tree = K.parse(f.xml), uri = at(tree, 'uri');
      let list = extensions.get(uri);
      if (!list) { list = { f, references: [] }; extensions.set(uri, list); }
      const node = all(tree).find(e => rid(e)?.startsWith('\u0001rel:'));
      if (node) list.references.push(K.attributes(K.raw(node), { [Array.from(node.attributes).find(a => a.value === rid(node)).name]: rId }));
    }
    const xml = [];
    for (const { f, references } of extensions.values()) {
      const tree = K.parse(f.xml), list = kids(tree)[0], entry = kids(list)[0];
      if (!entry) continue;
      const value = K.mergeBag(K.raw(list), { ['{' + entry.namespaceURI + '}' + entry.localName]: references });
      const merged = K.mergeBag(f.xml, { ['{' + list.namespaceURI + '}' + list.localName]: value });
      xml.push(w.emit(K.slice(f, merged), owner));
    }
    return xml.length ? K.mergeBag(generated || '<extLst xmlns="' + N.s + '"/>', { __append: xml }) : generated;
  };
  S.write = function (state) {
    const { wb, pack } = state, pkg = wb.pkg, w = pack.writer; if (!pkg) return;
    for (const g of state.groups) {
      const original = pkg.text(g.part), tree = K.parse(original), nodes = kids(tree).filter(e => ['slicer', 'timeline'].includes(e.localName));
      const unchanged = g.target === g.part && same(g.members.map(v => [v.name, v.cache]), nodes.map(e => [at(e, 'name'), at(e, 'cache')]));
      if (unchanged) { w.carry(pkg, g.part); continue; }
      const entries = g.members.map(v => {
        const el = nodes.find(e => at(e, 'name') === v.original), attrs = { name: v.name, cache: v.cache };
        if (v.uid) for (const a of Array.from(el.attributes)) if (a.localName === 'uid') attrs[a.name] = v.uid;
        return K.attributes(K.raw(el), attrs);
      });
      const xml = K.partXML(original, K.mergeBag(K.raw(tree), { ['{' + nodes[0].namespaceURI + '}' + nodes[0].localName]: entries }));
      if (g.target === g.part) w.claim(g.part, 'merged', g.target);
      w.put(g.target, K.hoistNamespaces(xml), pkg.type(g.part)); w.carryRels(pkg, g.part, g.target);
    }
    for (const c of state.caches.filter(c => !c.drop)) {
      const original = pkg.text(c.part), tree = K.parse(original), patches = [];
      const links = kids(tree).find(e => e.localName === 'pivotTables');
      if (links) {
        const old = kids(links), kept = c.linked.map(link => {
          const node = old.find(e => at(e, 'tabId') === link.tabId && at(e, 'name') === link.name);
          return link.table ? K.attributes(K.raw(node), { tabId: sheet(wb, link.table.sheet).sheetId, name: link.table.name }) : K.raw(node);
        });
        if (!same(kept, old.map(K.raw))) { const pos = L.xmlTree.source.get(links); patches.push({ start: pos.start, end: pos.end, value: K.mergeBag(K.raw(links), { ['{' + old[0].namespaceURI + '}pivotTable']: kept }) }); }
      }
      let xml = K.patch(original, patches);
      if (c.copy) {
        const root = K.parse(xml), attrs = { name: c.name };
        for (const a of Array.from(root.attributes)) if (a.localName === 'uid') attrs[a.name] = c.copy;
        xml = K.attributes(xml, attrs);
        const t = all(K.parse(xml)).find(e => e.localName === 'tableSlicerCache'), pos = L.xmlTree.source.get(t);
        xml = K.patch(xml, [{ start: pos.start, end: pos.end, value: K.attributes(K.raw(t), { tableId: c.modelTable.id }) }]);
      }
      if (xml === original) w.carry(pkg, c.part);
      else {
        if (!c.copy) w.claim(c.part, 'merged', c.target);
        w.put(c.target, K.hoistNamespaces(xml), pkg.type(c.part)); w.carryRels(pkg, c.part, c.target);
      }
    }
  };
  S.workbook = function (fragment, xml, writer) {
    const state = writer.slicers; if (!state || !state.omitted.size && !state.caches.some(c => c.copy)) return xml;
    function change(el) {
      const ref = rid(el), token = /^\u0001rel:(\d+)\u0001$/.exec(ref || '');
      if (token && cacheRel(fragment.deps[+token[1]]?.type)) {
        const part = fragment.deps[+token[1]].part, attrs = Array.from(el.attributes).find(a => a.value === ref);
        const copies = state.caches.filter(c => c.part === part && c.copy && !c.drop).map(c => K.attributes(K.raw(el), {
          [attrs.name]: writer.rels(writer.pkg.main).add(c.rel.type, K.relative(writer.pkg.main, c.target)) }));
        return (state.omitted.has(part) ? '' : K.raw(el)) + copies.join('');
      }
      const raw = K.raw(el), tree = K.parse(raw), patches = [];
      for (const child of kids(tree)) { const value = change(child); if (value !== K.raw(child)) { const p = L.xmlTree.source.get(child); patches.push({ start: p.start, end: p.end, value }); } }
      const result = K.patch(raw, patches);
      if (['ext', 'extLst', 'slicerCaches', 'timelineCacheRefs'].includes(el.localName) && !kids(K.parse(result)).length) return '';
      return result;
    }
    return change(K.parse(xml));
  };
})(typeof window !== 'undefined' ? window : globalThis);
