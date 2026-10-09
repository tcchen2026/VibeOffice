/* Ledger — edit locations around preserved pivot definitions, never rebuild their contents. */
(function (root) {
  'use strict';
  const L = root.L, K = L.opc, F = L.formula, P = (L.pivots = {});
  const kids = (el, name) => Array.from(el?.children || []).filter(e => !name || e.localName === name);
  const kid = (el, name) => kids(el, name)[0];
  const at = (el, name) => el?.getAttribute(name);
  const clone = value => value == null ? null : JSON.parse(JSON.stringify(value));
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  const state = wb => wb.extra.pivots;
  const sheet = (wb, id) => wb.sheets.find(s => s.id === id);
  const relation = (pkg, part, type) => pkg.rels(part).filter(r => K.relationshipType(r.type) === K.NS.rel + '/' + type && !r.external);
  const xmlPart = (pkg, part) => { try { return pkg.xml(part); } catch (_) { return null; } };
  const rid = el => el?.getAttributeNS(K.NS.rel, 'id') || el?.getAttributeNS(K.NS.strictRel, 'id');
  const inside = (rg, r, c) => rg && r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2;

  function definedName(wb, src) {
    return wb.names.find(n => n.name.toLowerCase() === src.name.toLowerCase() &&
      (n.scope == null ? null : wb.sheets[n.scope]?.id) === (src.scope || null));
  }
  function namedArea(wb, n, local, seen = new Set()) {
    if (!n || seen.has(n)) return null;
    seen.add(n);
    let a; try { a = F.parse(String(n.ref).replace(/^=/, '')); } catch (_) { return null; }
    if (a.book != null || a.sheet2) return null;
    if (a.t === 'referr') return { missing: true };
    if (a.t === 'name') {
      const sh = a.sheet ? wb.sheetByName(a.sheet) : local;
      const next = wb.names.find(x => x.name.toLowerCase() === a.name.toLowerCase() && x.scope != null && wb.sheets[x.scope] === sh) ||
        wb.names.find(x => x.name.toLowerCase() === a.name.toLowerCase() && x.scope == null);
      return namedArea(wb, next, sh, seen);
    }
    if (a.t !== 'ref' && a.t !== 'area') return null;
    const sh = a.sheet ? wb.sheetByName(a.sheet) : n.scope != null ? wb.sheets[n.scope] : local;
    if (!sh) return a.sheet ? { missing: true } : null;
    return { sheet: sh.id, ref: a.t === 'ref' ? { r1: a.r, c1: a.c, r2: a.r, c2: a.c } : { r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2 } };
  }
  function area(wb, src) {
    if (src.kind === 'range') return { sheet: src.sheet, ref: src.ref, missing: !sheet(wb, src.sheet) || !src.ref };
    if (src.kind === 'table') {
      const t = wb.tables.find(t => t.id === src.table && t.sheet.id === src.sheet);
      return t ? { sheet: t.sheet.id, ref: t.ref, name: t.name } : { missing: true };
    }
    if (src.kind === 'name') {
      const n = definedName(wb, src);
      return n ? { ...namedArea(wb, n, sheet(wb, src.sheet)), formula: n.ref } : { missing: true };
    }
    return {};
  }
  function source(wb, el) {
    const name = at(el, 'name'), local = wb.sheetByName(at(el, 'sheet') || '');
    const out = { kind: 'unknown', originalSheet: at(el, 'sheet'), originalRef: at(el, 'ref'), name };
    // A worksheet source may be an external workbook, even when a local table
    // happens to have the same name. Its relationship takes precedence.
    if (rid(el)) return { ...out, kind: 'external' };
    if (name) {
      const t = wb.findTable(name);
      if (t) return { ...out, kind: 'table', table: t.id, sheet: t.sheet.id };
      const n = wb.names.find(n => n.name.toLowerCase() === name.toLowerCase() && n.scope != null && wb.sheets[n.scope] === local) ||
        wb.names.find(n => n.name.toLowerCase() === name.toLowerCase() && n.scope == null);
      if (n) return { ...out, kind: 'name', sheet: local?.id, scope: n.scope == null ? null : wb.sheets[n.scope]?.id };
    }
    const ref = F.parseRange(at(el, 'ref') || '');
    if (local && ref) return { ...out, kind: 'range', sheet: local.id, ref };
    return out;
  }
  P.read = function (wb) {
    const pkg = wb.pkg;
    if (!pkg) return;
    const caches = [], tables = [];
    for (const el of kids(kid(pkg.xml(pkg.main), 'pivotCaches'), 'pivotCache')) {
      const rel = pkg.rels(pkg.main).find(r => r.id === rid(el) && !r.external), root = rel && xmlPart(pkg, rel.part);
      if (!root) {
        // Damaged opaque cache bytes must not prevent opening the readable cells.
        if (rel && pkg.has(rel.part)) caches.push({ id: at(el, 'cacheId'), part: rel.part, kind: 'unknown', sources: [], fields: [], opaque: true });
        continue; // The ordinary graph carry reports missing parts.
      }
      const cs = kid(root, 'cacheSource'), kind = at(cs, 'type');
      const sources = kind === 'worksheet' ? kids(cs, 'worksheetSource') : kind === 'consolidation' ? kids(kid(kid(cs, 'consolidation'), 'rangeSets'), 'rangeSet') : [];
      caches.push({ id: at(el, 'cacheId'), part: rel.part, kind, sources: sources.map(s => source(wb, s, rel.part)),
        fields: kids(kid(root, 'cacheFields'), 'cacheField').map(f => at(f, 'name')) });
    }
    for (const sh of wb.sheets) for (const rel of relation(pkg, sh.extra.ooxmlPart || '', 'pivotTable')) {
      const root = xmlPart(pkg, rel.part), location = kid(root, 'location');
      if (!root || !location) continue;
      const used = new Set();
      kids(kid(root, 'pivotFields'), 'pivotField').forEach((f, i) => { if (at(f, 'axis') || /^(1|true)$/.test(at(f, 'dataField'))) used.add(i); });
      for (const [tag, child, key] of [['rowFields', 'field', 'x'], ['colFields', 'field', 'x'], ['pageFields', 'pageField', 'fld'], ['dataFields', 'dataField', 'fld'], ['filters', 'filter', 'fld']])
        for (const e of kids(kid(root, tag), child)) if (+at(e, key) >= 0) used.add(+at(e, key));
      tables.push({ part: rel.part, sheet: sh.id, name: at(root, 'name'), cache: at(root, 'cacheId'),
        ref: F.parseRange(at(location, 'ref')), offsets: Object.fromEntries(['firstHeaderRow', 'firstDataRow', 'firstDataCol'].filter(k => location.hasAttribute(k)).map(k => [k, +at(location, k)])), used: [...used] });
    }
    if (!caches.length && !tables.length) return;
    wb.extra.pivots = { caches, tables };
    for (const cache of caches) {
      const a = cache.sources.length === 1 && area(wb, cache.sources[0]);
      if (a?.ref && cache.kind !== 'consolidation') cache.columns = Array.from({ length: a.ref.c2 - a.ref.c1 + 1 }, (_, i) => i);
      cache.formulas = cache.sources.some(s => hasFormula(wb, area(wb, s)));
    }
  };
  function hasFormula(wb, a) {
    if (!a.ref) return false;
    return sheet(wb, a.sheet)?.rows.some((row, r) => r >= a.ref.r1 && r <= a.ref.r2 && row?.cells.some((c, i) => i >= a.ref.c1 && i <= a.ref.c2 && c?.f != null)) || false;
  }
  P.snapshot = wb => clone(state(wb));
  P.restore = (wb, value) => { if (value) wb.extra.pivots = clone(value); else delete wb.extra.pivots; };
  const areas = wb => (state(wb)?.caches || []).map(c => c.sources.map(s => clone(area(wb, s))));
  const sourcesUnbounded = (wb, c) => c.formulas || c.sources.some(s => s.kind === 'name' && !area(wb, s).ref);
  const names = wb => JSON.stringify([wb.sheets.map(s => [s.id, s.name]), wb.names.map(n => [n.name, n.scope, n.ref])]);
  P.begin = wb => state(wb) && { data: P.snapshot(wb), areas: areas(wb), names: names(wb) };
  function validate(wb, p = state(wb)) {
    if (!p) return;
    for (const c of p.caches) if (!c.drop && c.sources.some(s => area(wb, s).missing)) c.drop = 'The pivot source was deleted.';
    for (const t of p.tables) {
      const c = p.caches.find(c => c.id === t.cache);
      if (!t.drop && (!sheet(wb, t.sheet) || c?.drop)) { t.drop = !sheet(wb, t.sheet) ? 'The worksheet containing the pivot table was deleted.' : c.drop; t.notify = !!sheet(wb, t.sheet); }
    }
  }
  P.current = wb => { const p = clone(state(wb)); validate(wb, p); return p; };
  P.end = function (wb, before) {
    if (!before) return;
    validate(wb);
    const afterAreas = areas(wb), p = state(wb);
    p.caches.forEach((c, i) => {
      if (!same(before.areas[i], afterAreas[i]) || (before.names !== names(wb) && sourcesUnbounded(wb, c))) c.refresh = true;
    });
    const after = P.snapshot(wb);
    if (!same(before.data, after)) wb.undo.op(() => P.restore(wb, before.data), () => P.restore(wb, after));
  };
  P.cell = function (sh, r, c, old, value) {
    const input = cell => cell?.f != null ? ['f', cell.f] : cell?.v ?? null;
    if (same(input(old), input(value))) return;
    const p = state(sh.wb); if (!p) return;
    for (const cache of p.caches) if (!cache.drop) {
      if (cache.formulas) cache.refresh = true; // A precedent may be on another sheet.
      for (const src of cache.sources) {
        const a = area(sh.wb, src);
        if ((src.kind === 'name' && !a.ref) || (a.sheet === sh.id && inside(a.ref, r, c))) {
          cache.refresh = true;
          if (value?.f != null) cache.formulas = true;
        }
      }
    }
    for (const t of p.tables) if (!t.drop && t.sheet === sh.id && inside(t.ref, r, c)) t.edited = true;
  };
  function fingerprint(wb, a) {
    const sh = sheet(wb, a.sheet), rg = a.ref;
    if (!sh || !rg) return null;
    const cells = [];
    sh.rows.forEach((row, r) => { if (r < rg.r1 || r > rg.r2 || !row) return;
      row.cells.forEach((cell, c) => { if (c < rg.c1 || c > rg.c2 || !cell || (cell.v == null && cell.f == null)) return;
        cells.push([r - rg.r1, c - rg.c1, cell.f == null ? cell.v : ['f', cell.f]]); }); });
    return JSON.stringify(cells);
  }
  // Sort, cut/paste and partial shifts write cells directly. Compare just the
  // affected pivot areas once around that structural transaction.
  P.structural = sh => {
    const p = state(sh.wb); if (!p) return null;
    return { caches: p.caches.map(c => c.sources.map(s => fingerprint(sh.wb, area(sh.wb, s)))),
      precedents: p.caches.some(c => sourcesUnbounded(sh.wb, c)) ? fingerprint(sh.wb, { sheet: sh.id, ref: { r1: 0, c1: 0, r2: L.model.MAXR, c2: L.model.MAXC } }) : null,
      tables: p.tables.map(t => ({ ref: clone(t.ref), value: fingerprint(sh.wb, t) })) };
  };
  P.afterStructural = function (sh, before) {
    if (!before) return;
    const p = state(sh.wb);
    const changed = before.precedents != null && before.precedents !== fingerprint(sh.wb, { sheet: sh.id, ref: { r1: 0, c1: 0, r2: L.model.MAXR, c2: L.model.MAXC } });
    p.caches.forEach((c, i) => {
      if ((changed && sourcesUnbounded(sh.wb, c)) || c.sources.some((s, j) => fingerprint(sh.wb, area(sh.wb, s)) !== before.caches[i]?.[j])) c.refresh = true;
      c.formulas = c.sources.some(s => hasFormula(sh.wb, area(sh.wb, s)));
    });
    p.tables.forEach((t, i) => {
      if (same(t.ref, before.tables[i]?.ref) && fingerprint(sh.wb, t) !== before.tables[i]?.value) t.edited = true;
    });
  };
  P.lines = function (sh, axis, at, n) {
    const p = state(sh.wb), O = L.ops; if (!p) return;
    const k1 = axis === 'r' ? 'r1' : 'c1', k2 = axis === 'r' ? 'r2' : 'c2';
    // A whole-column (A:D) or whole-row source stays whole along that axis, as in Excel.
    const whole = rg => axis === 'r' ? rg.r1 === 0 && rg.r2 === L.model.MAXR - 1 : rg.c1 === 0 && rg.c2 === L.model.MAXC - 1;
    for (const c of p.caches) if (sourcesUnbounded(sh.wb, c)) c.refresh = true;
    for (const cache of p.caches) for (const src of cache.sources) {
      const a = area(sh.wb, src);
      if (a.sheet !== sh.id || !a.ref) continue;
      const rg = a.ref, kept = whole(rg), next = kept ? rg : O.shiftRange(rg, axis, at, n);
      if (!kept && same(rg, next)) continue;
      if (!next) cache.drop = 'The complete pivot source range was deleted.';
      if (src.kind === 'range' && !kept) src.ref = next;
      cache.refresh = true;
      if (axis === 'c' && cache.columns && cache.sources.length === 1) {
        if (n < 0) {
          const from = Math.max(at, rg.c1) - rg.c1, to = Math.min(at - n - 1, rg.c2) - rg.c1;
          const removed = to >= from ? cache.columns.slice(from, to + 1) : [];
          for (const t of p.tables) if (!t.drop && t.cache === cache.id && t.used.some(i => removed.includes(i))) {
            t.drop = 'A source column used by the pivot table was deleted. Its remaining result cells are kept.'; t.notify = true;
          }
          if (to >= from) cache.columns.splice(from, to - from + 1);
        } else if (at > rg.c1 && at <= rg.c2 || kept && at - rg.c1 < cache.columns.length) cache.columns.splice(at - rg.c1, 0, ...Array(n).fill(null));
      }
    }
    for (const t of p.tables) if (!t.drop && t.sheet === sh.id && t.ref) {
      const rg = t.ref, next = O.shiftRange(rg, axis, at, n);
      if (!next) { t.drop = 'The complete pivot table output was deleted.'; continue; }
      // Lines removed from or inserted into the result change its layout; Excel rebuilds it on open.
      const through = n < 0 ? at <= rg[k2] && at - n - 1 >= rg[k1] : at > rg[k1] && at <= rg[k2];
      if (through) { t.reshaped = true; const c = p.caches.find(c => c.id === t.cache); if (c) c.refresh = true; }
      for (const key of axis === 'r' ? ['firstHeaderRow', 'firstDataRow'] : ['firstDataCol']) if (key in t.offsets) {
        const pos = rg[k1] + t.offsets[key];
        const moved = n > 0 ? pos >= at ? pos + n : pos : pos >= at - n ? pos + n : pos >= at ? at : pos;
        t.offsets[key] = Math.max(0, Math.min(next[k2] - next[k1], moved - next[k1]));
      }
      t.ref = next;
    }
  };
  P.copySheet = function (wb, from, to) {
    const p = state(wb); if (!p) return;
    for (const t of p.tables.slice()) if (!t.drop && t.sheet === from.id) {
      let i = 1, name; do { name = t.name + '_' + i++; } while (p.tables.some(t => t.name === name));
      p.tables.push({ ...clone(t), sheet: to.id, name, copy: true, uid: '{' + root.crypto.randomUUID().toUpperCase() + '}' });
    }
  };
  // Patch only requested attributes, without serializing the rest of a part.
  function patchXML(xml, edits) {
    const patches = [];
    for (const [el, attrs] of edits) {
      if (!el) continue;
      const pos = L.xmlTree.source.get(el);
      for (const [key, value] of Object.entries(attrs)) {
        if (String(at(el, key)) === String(value)) continue;
        const a = pos.attrs.find(a => a.name === key), escaped = L.xml.esc(String(value));
        if (a) patches.push({ start: a.start, end: a.end, value: escaped });
        else patches.push({ start: pos.openEnd - (xml[pos.openEnd - 2] === '/' ? 2 : 1), end: pos.openEnd - (xml[pos.openEnd - 2] === '/' ? 2 : 1), value: ' ' + key + '="' + escaped + '"' });
      }
    }
    return K.patch(xml, patches);
  }
  P.write = function (wb, pack) {
    const p = clone(state(wb)), pkg = wb.pkg, w = pack.writer; if (!p || !pkg) return;
    validate(wb, p);
    const retained = new Set(p.tables.filter(t => !t.drop).map(t => t.cache));
    w.pivotOmitted = new Set();
    for (const c of p.caches) {
      c.omit = !!c.drop || !retained.has(c.id) && p.tables.some(t => t.cache === c.id);
      if (c.omit) { w.pivotOmitted.add(c.id); w.omit(c.part, c.drop || 'The cache has no remaining pivot tables.'); continue; }
      if (c.opaque) { w.carry(pkg, c.part); continue; }
      let xml = pkg.text(c.part), tree = K.parse(xml), edits = [], cs = kid(tree, 'cacheSource');
      const elements = c.kind === 'worksheet' ? kids(cs, 'worksheetSource') : kids(kid(kid(cs, 'consolidation'), 'rangeSets'), 'rangeSet');
      c.sources.forEach((s, i) => {
        const attrs = {}, a = area(wb, s);
        if (s.kind === 'range' && a.ref && !same(a.ref, F.parseRange(s.originalRef))) attrs.ref = F.rangeName(a.ref);
        if (s.originalSheet && sheet(wb, s.sheet) && sheet(wb, s.sheet).name !== s.originalSheet) attrs.sheet = sheet(wb, s.sheet).name;
        if (s.kind === 'table' && a.name !== s.name) attrs.name = a.name;
        edits.push([elements[i], attrs]);
      });
      if (c.refresh) {
        edits.push([tree, { refreshOnLoad: '1' }]);
        w.loss({ id: 'pivot-refresh:' + c.part, what: 'The pivot source changed. Excel will refresh its cached result when the file opens.', where: c.part, action: 'conversion' });
      }
      xml = patchXML(xml, edits);
      if (xml === pkg.text(c.part)) w.carry(pkg, c.part);
      else {
        w.claim(c.part, 'merged', w.target(pkg, c.part));
        w.put(w.target(pkg, c.part), xml, pkg.type(c.part)); w.carryRels(pkg, c.part);
      }
    }
    for (const t of p.tables) {
      if (t.drop) {
        // Told only when the result stays behind as plain cells; a deleted sheet or output takes the pivot with it
        const notice = t.notify ? { what: 'This pivot table will be saved as plain cells, because ' + (/source column/.test(t.drop) ? 'a column it uses was deleted from its source data.' : 'its source data was deleted.'), place: (sheet(wb, t.sheet)?.name || '') + ': ' + t.name, notify: true } : { what: t.drop, notify: false };
        if (!t.copy) w.omit(t.part, notice.what, { notify: notice.notify, place: notice.place }); else w.loss({ id: 'pivot-copy:' + t.sheet + ':' + t.name, what: notice.what, place: notice.place, notify: notice.notify, where: t.name, action: 'drop' });
        continue;
      }
      let xml = pkg.text(t.part), tree = K.parse(xml), loc = kid(tree, 'location');
      const attrs = {}, originalRef = F.parseRange(at(loc, 'ref'));
      if (!same(t.ref, originalRef)) attrs.ref = F.rangeName(t.ref);
      for (const [key, value] of Object.entries(t.offsets)) if (+at(loc, key) !== value) attrs[key] = value;
      const edits = [[loc, attrs]];
      let target = w.target(pkg, t.part);
      if (t.copy) {
        target = w.name('xl/pivotTables', 'pivotTable', 'xml');
        const rootAttrs = { name: t.name };
        for (const a of Array.from(tree.attributes)) if (a.localName === 'uid') rootAttrs[a.name] = t.uid;
        edits.push([tree, rootAttrs]);
      }
      xml = patchXML(xml, edits);
      if (!t.copy && xml === pkg.text(t.part)) w.carry(pkg, t.part);
      else {
        if (!t.copy) w.claim(t.part, 'merged', target);
        w.put(target, xml, pkg.type(t.part)); w.carryRels(pkg, t.part, target);
      }
      const sh = sheet(wb, t.sheet), owner = pack.sheetParts.get(sh);
      const rel = relation(pkg, sh.extra.ooxmlPart || '', 'pivotTable').find(r => r.part === t.part);
      w.rels(owner).add(K.NS.rel + '/pivotTable', K.relative(owner, target), false, t.copy ? undefined : rel?.id);
      const changed = 'Changes made inside this pivot table will be replaced when Excel refreshes it.', place = sh.name + ': ' + t.name;
      if (t.reshaped) w.loss({ id: 'pivot-layout:' + t.sheet + ':' + t.name, what: changed, place, notify: true, where: sh.name + '!' + F.rangeName(t.ref), action: 'conversion' });
      if (t.edited) w.loss({ id: 'pivot-output:' + t.sheet + ':' + t.name, what: changed, place, notify: true, where: sh.name + '!' + F.rangeName(t.ref), action: 'conversion' });
    }
  };
  P.workbook = function (wb, tree, writer) {
    const p = state(wb), list = kid(tree, 'pivotCaches'); if (!p || !list) return undefined;
    const omitted = writer.pivotOmitted || new Set();
    if (!omitted.size) return undefined;
    // The settings merge receives a tokenised fragment; cacheId is an identity
    // token there, so consult the same original child for its source identity.
    const original = kids(kid(wb.pkg.xml(wb.pkg.main), 'pivotCaches'), 'pivotCache');
    const kept = kids(list, 'pivotCache').filter((e, i) => !omitted.has(at(original[i], 'cacheId')));
    return kept.length ? K.mergeBag(K.raw(list), { ['{' + K.NS.s + '}pivotCache']: kept.map(K.raw) }) : [];
  };
})(typeof window !== 'undefined' ? window : globalThis);
