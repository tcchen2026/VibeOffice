/* Ledger — table/column identities and retained external query definitions. */
(function (root) {
  'use strict';
  const L = root.L, K = L.opc, F = L.formula, T = (L.tableKeep = {}), N = K.NS;
  const kids = (e, name) => Array.from(e?.children || []).filter(c => !name || c.localName === name);
  const kid = (e, name) => kids(e, name)[0], at = (e, n) => e?.getAttribute(n);
  const all = e => e ? [e, ...e.getElementsByTagName('*')] : [];
  const clone = x => x == null ? x : JSON.parse(JSON.stringify(x));
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  const values = ({ sheet, columns, keep, queries, queryOps, queryWrites, copyKeep, ...t }) => clone(t);
  const colValues = ({ keep, ...c }) => clone(c);
  const cellValues = c => ({ v: c?.f == null ? c?.v ?? null : null, f: c?.f, rt: c?.rt });
  const rels = (pkg, owner) => pkg.rels(owner).filter(r => !r.external && K.relationshipType(r.type) === N.rel + '/queryTable');
  const fragment = (el, wb, part) => K.fragment(el, { pkg: wb.pkg, part });
  T.read = function (t, el) {
    const wb = t.sheet.wb, part = t.ooxmlPart;
    if (!wb.pkg) return; // The reader already reports an unrecoverable package graph.
    t.keep = { fragment: fragment(el, wb, part), values: values(t) };
    kids(kid(el, 'tableColumns'), 'tableColumn').forEach((c, i) => {
      const col = t.columns[i];
      if (at(c, 'uniqueName') != null) col.uniqueName = at(c, 'uniqueName');
      if (at(c, 'queryTableFieldId') != null) col.queryField = +at(c, 'queryTableFieldId');
      col.keep = { fragment: fragment(c, wb, part), values: colValues(col), header: clone(cellValues(t.sheet.get(t.ref.r1, t.ref.c1 + i))) };
    });
    t.queries = rels(wb.pkg, part);
  };
  T.readWorkbook = function (wb) {
    if (!wb.pkg) return;
    const original = [];
    for (const t of wb.tables) for (const r of t.queries || []) original.push({ part: r.part, owner: t.ooxmlPart });
    for (const sh of wb.sheets) {
      sh.extra.queryParts = rels(wb.pkg, sh.extra.ooxmlPart || '');
      for (const r of sh.extra.queryParts) original.push({ part: r.part, owner: sh.extra.ooxmlPart });
    }
    wb.extra.queries = original;
  };
  T.sameHeader = (t, c, i) => c.keep && same(c.keep.header, cellValues(t.sheet.get(t.ref.r1, t.ref.c1 + i))) && c.name === c.keep.values.name;
  T.copy = function (object, wb) {
    const { sheet, ...data } = object, ids = {}, previous = object.copyKeep?.ids;
    const sources = (object.queries || []).map(r => wb.pkg?.text(r.part)).filter(Boolean).map(xml => K.remapGuids(xml, previous));
    const copy = K.duplicate(data, { guidMap: ids, guidAttributes: ['uid'], guidSources: sources });
    // Opaque query parts still come from their source package, even when copying
    // a copy. Compose that original-to-current mapping with this duplication.
    for (const [old, current] of Object.entries(previous || {})) if (ids[current]) ids[old] = ids[current];
    copy.copyKeep = { id: K.guid(), ids };
    return sheet ? { ...copy, sheet } : copy;
  };
  T.copySheet = function (from, to) {
    to.extra.queryParts = clone(from.extra.queryParts || []);
    if (to.extra.queryParts.length) to.extra.queryCopy = { ...T.copy({ queries: to.extra.queryParts, copyKeep: from.extra.queryCopy }, from.wb).copyKeep, keepName: true };
  };
  const copyIds = (xml, copy) => K.remapGuids(xml, copy?.ids);
  function attributes(raw, fresh, before, after, fields) {
    const changes = {};
    for (const [name, field] of Object.entries(fields)) if (!same(before?.[field], after?.[field])) changes[name] = at(fresh, name);
    return K.attributes(raw, changes);
  }
  T.shift = function (sh, axis, at, n, shift) {
    sh.wb.tables = sh.wb.tables.filter(t => {
      if (t.sheet !== sh) return true;
      const old = t.ref, next = shift(old, axis, at, n); if (!next) return false;
      if (axis === 'c') {
        if (n > 0 && at > old.c1 && at <= old.c2) {
          // Never reuse a column id from the file: query fields and slicers bind to it.
          const original = t.keep ? kids(kid(K.parse(t.keep.fragment.xml), 'tableColumns'), 'tableColumn').map(c => +c.getAttribute('id') || 0) : [];
          let id = Math.max(0, ...original, ...t.columns.map(c => c.id || 0));
          const used = new Set(t.columns.map(c => c.name.toLowerCase()));
          const added = Array.from({ length: n }, () => {
            let name, k = 1; do { name = 'Column' + k++; } while (used.has(name.toLowerCase()));
            used.add(name.toLowerCase()); return { id: ++id, name };
          });
          t.columns.splice(at - old.c1, 0, ...added);
        } else if (n < 0) t.columns = t.columns.filter((c, i) => old.c1 + i < at || old.c1 + i >= at - n);
        if (t.filter?.cols) t.filter.cols = t.filter.cols.flatMap(c => {
          const p = shift({ r1: old.r1, r2: old.r1, c1: old.c1 + c.col, c2: old.c1 + c.col }, axis, at, n);
          return p ? [{ ...c, col: p.c1 - next.c1 }] : [];
        });
      }
      if (!same(old, next)) t.queryOps = (t.queryOps || []).concat({ axis, at, n });
      t.ref = next;
      return true;
    });
  };
  function ranges(xml, ops) {
    if (!ops?.length) return xml;
    const patches = [];
    for (const e of all(K.parse(xml))) if (['autoFilter', 'sortState', 'sortCondition'].includes(e.localName) && at(e, 'ref')) {
      const position = L.xmlTree.source.get(e);
      if (patches.some(q => q.start <= position.start && q.end >= position.end)) continue;
      let ref = F.parseRange(at(e, 'ref')); if (!ref) continue;
      const before = ref;
      for (const op of ops) if (ref) ref = L.ops.shiftRange(ref, op.axis, op.at, op.n);
      const p = L.xmlTree.source.get(e);
      // Patch attributes rather than nested nodes, so parent/child edits do not overlap.
      if (same(before, ref)) continue;
      if (ref) {
        const a = p.attrs.find(a => a.name === 'ref'); patches.push({ start: a.start, end: a.end, value: F.rangeName(ref) });
      } else if (!patches.some(q => q.start <= p.start && q.end >= p.end)) patches.push({ start: p.start, end: p.end, value: '' });
    }
    return K.patch(xml, patches);
  }
  T.begin = function (wb, pack) {
    const used = new Set();
    for (const t of wb.tables) if (!t.copyKeep) for (const r of t.queries || []) used.add(r.part);
    for (const sh of wb.sheets) if (!sh.extra.queryCopy) for (const r of sh.extra.queryParts || []) used.add(r.part);
    for (const q of wb.extra.queries || []) if (!used.has(q.part)) pack.writer.omit(q.part, 'The table or worksheet that owned this query was deleted.');
    const names = new Set([...used].map(part => (at(wb.pkg.xml(part), 'name') || '').toLowerCase())), copies = new Map();
    pack.writer.queryName = (name, copy, part) => {
      const key = copy.id + ':' + part; if (copies.has(key)) return copies.get(key);
      let next = name, n = +(name.match(/\d+$/)?.[0] || 0), base = name.replace(/\d+$/, '');
      if (base === name) base += '_';
      while (names.has(next.toLowerCase())) next = base + ++n;
      names.add(next.toLowerCase()); copies.set(key, next); return next;
    };
  };
  T.prepare = function (t, columns) {
    const current = { ...t, columns }, pkg = t.sheet.wb.pkg;
    current.queryWrites = (t.queries || []).map(rel => {
      const original = pkg.text(rel.part), tree = K.parse(original), refresh = kid(tree, 'queryTableRefresh');
      if (!refresh) return { rel, original, xml: original };
      const fields = kid(refresh, 'queryTableFields'), source = kids(fields, 'queryTableField');
      const originalColumns = kids(kid(K.parse(t.keep.fragment.xml), 'tableColumns'), 'tableColumn');
      const removed = [], output = [];
      let nextId = Math.max(+at(refresh, 'nextId') || 1, ...source.map(e => +at(e, 'id') + 1));
      const found = new Set();
      for (const col of columns) {
        const before = source.find(e => col.queryField != null && +at(e, 'id') === col.queryField || col.keep && +at(e, 'tableColumnId') === col.id);
        if (before) { output.push(K.raw(before)); found.add(before); }
        else if (!col.keep) {
          col.queryField = nextId++; col.uniqueName = String(col.id);
          output.push('<queryTableField xmlns="' + N.s + '" id="' + col.queryField + '" dataBound="0" tableColumnId="' + col.id + '"/>');
        }
      }
      for (const e of source) if (!found.has(e)) {
        const wasColumn = originalColumns.some(c => +at(c, 'id') === +at(e, 'tableColumnId'));
        if (!wasColumn) output.push(K.raw(e)); // clipped/unmodeled fields still belong to the external result
        else if (at(e, 'dataBound') !== '0' && at(e, 'name')) removed.push(at(e, 'name'));
      }
      const changed = output.length !== source.length || output.some((x, i) => x !== (source[i] && K.raw(source[i])));
      let xml = original;
      if (changed) {
        const list = K.attributes(K.mergeBag(fields ? K.raw(fields) : '<queryTableFields xmlns="' + N.s + '"/>', { ['{' + (fields?.namespaceURI || N.s) + '}queryTableField']: output }), { count: output.length });
        const children = output.map(K.parse), bound = children.map(e => at(e, 'dataBound') !== '0');
        const first = bound.indexOf(true), last = bound.lastIndexOf(true);
        let rxml = K.attributes(K.raw(refresh), { nextId, unboundColumnsLeft: first < 0 ? children.length || null : first || null,
          unboundColumnsRight: last < 0 ? null : children.length - last - 1 || null });
        const replacements = { ['{' + N.s + '}queryTableFields']: list };
        if (removed.length) {
          const old = kid(refresh, 'queryTableDeletedFields'), prior = kids(old, 'deletedField').map(e => at(e, 'name'));
          const gone = [...new Set([...prior, ...removed])];
          replacements['{' + N.s + '}queryTableDeletedFields'] = '<queryTableDeletedFields xmlns="' + N.s + '" count="' + gone.length + '">' +
            gone.map(name => '<deletedField name="' + L.xml.esc(name) + '"/>').join('') + '</queryTableDeletedFields>';
        }
        rxml = K.merge(rxml, replacements, 's:CT_QueryTableRefresh');
        xml = K.partXML(original, K.mergeBag(K.raw(tree), { ['{' + tree.namespaceURI + '}queryTableRefresh']: rxml }));
      }
      xml = ranges(xml, t.queryOps);
      return { rel, original, xml };
    });
    return current;
  };
  function writeQuery(wb, info, pack, owner, copy) {
    const w = pack.writer, pkg = wb.pkg, { rel, original } = info;
    let xml = info.xml, target;
    if (copy) {
      target = w.name('xl/queryTables', 'queryTable', 'xml');
      w.copyPart(pkg, rel.part, target);
      const tree = K.parse(xml), name = at(tree, 'name');
      xml = K.partXML(xml, copyIds(K.attributes(K.raw(tree), { name: copy.keepName ? name : w.queryName(name, copy, rel.part) }), copy));
    } else target = w.target(pkg, rel.part);
    if (!copy && xml === original) w.carry(pkg, rel.part);
    else {
      if (!copy) { w.claim(rel.part, 'regenerated', target); w.carryRels(pkg, rel.part, target); }
      w.put(target, K.hoistNamespaces(xml), pkg.type(rel.part));
    }
    w.rels(owner).add(rel.type, K.relative(owner, target), false, copy ? undefined : rel.id);
  }
  T.sheet = function (sh, pack, owner) {
    for (const rel of sh.extra.queryParts || []) {
      const original = sh.wb.pkg.text(rel.part);
      writeQuery(sh.wb, { rel, original, xml: original }, pack, owner, sh.extra.queryCopy);
    }
  };
  T.xml = function (t, generated, pack, owner) {
    for (const q of t.queryWrites || []) writeQuery(t.sheet.wb, q, pack, owner, t.copyKeep);
    if (!t.keep) return generated;
    if (t.copyKeep) pack.writer.carryRels(t.sheet.wb.pkg, t.ooxmlPart, owner, r => K.relationshipType(r.type) !== N.rel + '/queryTable');
    const f = t.keep.fragment, fresh = K.parse(generated), before = t.keep.values;
    let xml = attributes(f.xml, fresh, before, values(t), { id: 'id', name: 'dname', displayName: 'name', ref: 'ref',
      headerRowCount: 'header', totalsRowCount: 'totals', totalsRowShown: 'totals', tableType: 'type', comment: 'comment' });
    const original = K.parse(xml), replacements = {};
    const auto = kid(original, 'autoFilter'), nextAuto = kid(fresh, 'autoFilter');
    if (!same(before.filter, t.filter) || before.autoFilter !== t.autoFilter || before.header !== t.header) replacements['{' + N.s + '}autoFilter'] = nextAuto ? K.raw(nextAuto) : [];
    else if (auto && !same(before.ref, t.ref)) replacements['{' + N.s + '}autoFilter'] = K.attributes(ranges(K.raw(auto), t.queryOps), { ref: at(nextAuto, 'ref') });
    if (!same(before.style, t.style)) replacements['{' + N.s + '}tableStyleInfo'] = kids(fresh, 'tableStyleInfo').map(K.raw);
    const cols = kid(fresh, 'tableColumns');
    const merged = t.columns.map((c, i) => {
      const keep = c.keep, next = cols.children[i]; if (!keep) return K.raw(next);
      let xml = attributes(keep.fragment.xml, next, keep.values, colValues(c), { id: 'id', name: 'name', uniqueName: 'uniqueName', queryTableFieldId: 'queryField',
        totalsRowFunction: 'totalsFn', totalsRowLabel: 'totalsLabel', dataDxfId: 'dataDxf' });
      const children = {};
      for (const [tag, keys] of Object.entries({ calculatedColumnFormula: ['calc'], totalsRowFormula: ['totalsFn', 'totalsFormula'] })) if (keys.some(k => !same(keep.values[k], c[k]))) children['{' + N.s + '}' + tag] = kids(next, tag).map(K.raw);
      xml = K.merge(xml, children, 's:CT_TableColumn');
      return pack.writer.emit(K.slice(keep.fragment, copyIds(xml, t.copyKeep)), owner);
    });
    const oldCols = kid(original, 'tableColumns');
    replacements['{' + N.s + '}tableColumns'] = K.attributes(K.mergeBag(K.raw(oldCols), { ['{' + oldCols.namespaceURI + '}tableColumn']: merged }), { count: merged.length });
    const sort = kid(original, 'sortState'); if (sort && t.queryOps?.length) replacements['{' + N.s + '}sortState'] = ranges(K.raw(sort), t.queryOps);
    xml = copyIds(K.merge(xml, replacements, 's:CT_Table'), t.copyKeep);
    return K.partXML(t.sheet.wb.pkg.text(t.ooxmlPart), pack.writer.emit(K.slice(f, xml), owner));
  };
})(typeof window !== 'undefined' ? window : globalThis);
