/* Lectern — keep unsupported frames attached to the previews the editor displays. */
(function () {
  'use strict';
  const L = window.L, K = L.opc, X = L.xmlTree, F = (L.frames = {}), N = K.NS;
  const kids = e => Array.from(e?.children || []);
  const all = e => e ? [e, ...kids(e).flatMap(all)] : [];
  const at = (e, key) => e?.getAttribute(key);
  const GEOMETRY = new Set(['x', 'y', 'w', 'h', 'rot', 'flipH', 'flipV']);
  const box = s => Object.fromEntries([...GEOMETRY].map(k => [k, s[k] || 0]));
  const key = f => [f.source, f.part, f.record || f.key, f.copy || ''].join('|');
  const rootNames = new Set(['sp', 'pic', 'cxnSp', 'grpSp', 'graphicFrame', 'contentPart', 'model3d']);
  function label(root) {
    const nodes = all(root);
    if (nodes.some(e => e.localName === 'relIds' && /\/diagram$/.test(e.namespaceURI))) return 'SmartArt';
    if (nodes.some(e => e.localName === 'oleObj')) return 'Embedded object';
    if (nodes.some(e => e.localName === 'contentPart')) return 'Ink';
    if (nodes.some(e => e.localName === 'model3d')) return '3-D model';
    if (nodes.some(e => e.localName === 'chart' && /\/chartex$/.test(e.namespaceURI))) return 'New chart';
    return null;
  }
  function source(working, original, parent) {
    if (!rootNames.has(working.localName) && working.localName !== 'AlternateContent') return null;
    const alternate = K.alternate(working);
    // Descendants of the chosen group belong to that group's outer wrapper.
    if (alternate && alternate === K.alternate(parent)) return null;
    const tree = alternate ? K.parse(alternate.xml) : original;
    if (!tree || !alternate && tree.localName === 'grpSp') return null;
    const kind = label(tree); return kind ? { tree, label: kind, alternate } : null;
  }
  F.placeholder = function (working, original, parent, ctx) {
    const src = source(working, original, parent); if (!src) return null;
    const b = K.getBox(K.raw(src.tree)), map = ctx.map || (v => v);
    const cnv = all(src.tree).find(e => e.localName === 'cNvPr');
    return { id: L.uid('s'), type: 'image', name: at(cnv, 'name') || src.label, numId: +(at(cnv, 'id') || 0),
      ...map({ x: b.x || 0, y: b.y || 0, w: b.w ?? 100, h: b.h ?? 60 }), rot: b.rot || 0, flipH: !!b.flipH, flipV: !!b.flipV,
      missingLabel: src.label, media: null, geom: 'rect', crop: { l: 0, t: 0, r: 0, b: 0 }, img: {}, line: { t: 'none' } };
  };
  F.attach = function (models, working, original, parent, pkg, part) {
    const src = source(working, original, parent); if (!src) return;
    const fragment = K.fragment(src.alternate ? working : src.tree, { pkg, part });
    fragment.key = fragment.record || L.uid('frame');
    const related = [];
    for (const dep of fragment.deps) if (/\/diagramData$/.test(dep.type || '')) {
      let tree; try { tree = pkg.xml(dep.part); } catch (_) { continue; }
      for (const el of all(tree).filter(e => e.localName === 'dataModelExt')) {
        const rel = pkg.rels(part).find(r => r.id === at(el, 'relId'));
        if (rel) related.push({ source: pkg.id, owner: part, dataPart: dep.part, ...rel });
      }
    }
    if (src.label === 'Embedded object' && all(src.tree).some(e => e.localName === 'oleObj' && e.hasAttribute('spid'))) {
      for (const rel of pkg.rels(part).filter(r => /\/vmlDrawing$/.test(r.type))) related.push({ source: pkg.id, owner: part, ...rel });
    }
    for (const s of models) {
      s.keep ||= { source: pkg.id, part };
      s.keep.frame = { fragment: JSON.parse(JSON.stringify(fragment)), label: src.label, related, localBox: K.getBox(K.raw(working)) };
      // After ungrouping, the children still explain the conversion of their original frame.
      L.model.walk(s.kids || [], child => { child.keep ||= {}; child.keep.frameOrigin = { key: fragment.key, label: src.label, part }; return true; });
    }
  };
  function containers(pres, visit) {
    const walk = list => { visit(list); for (const s of list || []) if (s.kids) walk(s.kids); };
    for (const slide of pres.slides) walk(slide.shapes);
    for (const d of Object.values(pres.designs)) for (const list of [d.deco, d.titleDeco, ...Object.values(d.layoutDecos || {})]) if (list) walk(list);
  }
  const union = shapes => L.unionBounds(shapes.map(s => ({ x: s.x, y: s.y, w: s.w, h: s.h })));
  function signature(shape, outer) {
    const visit = (v, name, isShape) => {
      if (name === 'keep' || name === 'id' || name === 'numId' || name === 'view') return undefined;
      if (isShape && GEOMETRY.has(name)) return undefined;
      if (Array.isArray(v)) return v.map(x => visit(x, '', false));
      if (v && typeof v === 'object') {
        const out = {}, ownShape = rootNames.has(v.type) || ['image', 'shape', 'text', 'group', 'chart', 'table', 'wordart', 'line'].includes(v.type);
        for (const k of Object.keys(v).sort()) {
          if (ownShape && ['x', 'y', 'w', 'h'].includes(k)) {
            const axis = k === 'x' || k === 'w' ? 'w' : 'h', origin = k === 'x' || k === 'y' ? outer[k] : 0;
            out[k] = Math.round((v[k] - origin) / (outer[axis] || 1) * 1e6) / 1e6;
          } else {
            const value = visit(v[k], k, ownShape && v === shape); if (value !== undefined) out[k] = value;
          }
        }
        return out;
      }
      return v;
    };
    return JSON.stringify(visit(shape, '', false));
  }
  F.seal = async function (pres, media) {
    const records = [], images = [], imageParts = new Map(Array.from(media, ([part, id]) => [id, part]));
    containers(pres, list => {
      for (const s of list) if (!s.keep?.frame && s.keep?.properties) {
        const previews = [], collect = v => {
          if (!v || typeof v !== 'object') return;
          for (const [k, value] of Object.entries(v)) {
            if (k === 'keep') continue;
            if (typeof value === 'string' && L.media.has(value)) previews.push({ id: value, source: imageParts.has(value) ? pres.pkg.id : null, part: imageParts.get(value), size: L.media.get(value).size });
            else collect(value);
          }
        };
        collect(s); s.keep.properties.previews = previews; images.push(...previews);
      }
      const groups = new Map();
      for (const s of list) if (s.keep?.frame) {
        const k = key(s.keep.frame.fragment); if (!groups.has(k)) groups.set(k, []); groups.get(k).push(s);
      }
      for (const members of groups.values()) {
        const bounds = union(members);
        const localBoxes = members.map(s => s.keep.frame.localBox);
        const localBounds = localBoxes.every(b => ['x', 'y', 'w', 'h'].every(k => Number.isFinite(b[k]))) ? union(localBoxes) : K.getBox(members[0].keep.frame.fragment.xml);
        members.forEach((s, index) => {
          const keep = s.keep.frame, previews = [];
          const collect = v => {
            if (!v || typeof v !== 'object') return;
            for (const [k, value] of Object.entries(v)) {
              if (k === 'keep') continue;
              if (typeof value === 'string' && L.media.has(value)) previews.push({ id: value, source: imageParts.has(value) ? keep.fragment.source : null, part: imageParts.get(value), size: L.media.get(value).size });
              else collect(value);
            }
          };
          collect(s); Object.assign(keep, { index, count: members.length, bounds, localBounds, box: box(s), signature: signature(s, bounds), previews });
          images.push(...previews);
          if (!records.some(r => r.key === keep.fragment.key && r.part === keep.fragment.part)) records.push({ key: keep.fragment.key, label: keep.label, part: keep.fragment.part, rels: [...keep.fragment.deps, ...keep.related].map(d => d.id) });
        });
      }
    });
    // Display renditions (metafiles, duotones) travel with the clipboard too, but
    // their bytes remain outside the document's JSON history.
    const parts = [], bundled = new Map();
    const data = async blob => { const bytes = new Uint8Array(await blob.arrayBuffer()); let text = ''; for (let i = 0; i < bytes.length; i += 32768) text += String.fromCharCode(...bytes.subarray(i, i + 32768)); return btoa(text); };
    for (const ref of images) {
      if (!bundled.has(ref.id)) {
        const m = L.media.get(ref.id), entry = {};
        for (const field of ['blob', 'view']) if (field === 'blob' ? !ref.source : m.view !== m.blob) {
          const name = 'preview/' + (parts.length + 1) + '.bin';
          parts.push({ name, type: m[field].type || 'image/png', data: await data(m[field]) }); entry[field] = name;
        }
        bundled.set(ref.id, entry);
      }
    }
    if (parts.length) {
      const source = K.import([{ source: 'previews', parts }]).previews.id;
      for (const ref of images) {
        const entry = bundled.get(ref.id);
        if (entry.blob) { ref.source = source; ref.part = entry.blob; }
        if (entry.view) ref.view = { source, part: entry.view };
      }
    }
    pres.keep.frames = records;
  };
  F.consumes = (pres, part, rel) => (pres.keep?.frames || []).some(f => f.part === part && f.rels.includes(rel.id));
  // What the user is told when a kept object can only be saved as its preview
  const savedAs = (label, why) => 'This ' + (({ 'New chart': 'chart', Ink: 'ink drawing', 'Embedded object': 'embedded object' })[label] || label) +
    ' will be saved as ' + (label === 'Embedded object' ? 'a picture' : 'ordinary shapes') + why;
  const slidePlace = (pres, part) => { const i = pres.slides.findIndex(s => s.keep?.part === part); return i < 0 ? undefined : 'Slide ' + (i + 1); };
  F.prepare = function (pres, writer) {
    const groups = new Map(), seen = new Set(), converted = new Map();
    containers(pres, list => {
      list.forEach((shape, index) => {
        const keep = shape.keep?.frame;
        if (keep) {
          const k = key(keep.fragment); if (!groups.has(k)) groups.set(k, { entries: [], emitted: false });
          groups.get(k).entries.push({ shape, list, index }); seen.add(keep.fragment.key);
        }
        const origin = shape.keep?.frameOrigin; if (origin) converted.set(origin.key, origin);
      });
    });
    for (const group of groups.values()) {
      const first = group.entries[0], record = first.shape.keep.frame, shapes = group.entries.map(e => e.shape);
      group.bounds = union(shapes);
      group.intact = shapes.length === record.count && group.entries.every((e, i) => e.list === first.list && e.index === first.index + i && e.shape.keep.frame.index === i && signature(e.shape, group.bounds) === e.shape.keep.frame.signature);
      group.moved = group.entries.some(e => JSON.stringify(box(e.shape)) !== JSON.stringify(e.shape.keep.frame.box));
    }
    for (const record of [...(pres.keep?.frames || []), ...converted.values()]) if (!seen.has(record.key)) {
      // Removed objects were deleted by the user; converted ones were ungrouped or edited
      if (converted.has(record.key)) writer.loss({ id: 'frame:' + record.key, what: savedAs(record.label, ', because it was edited here.'), where: record.part, place: slidePlace(pres, record.part), action: 'conversion', notify: true });
      else writer.loss({ id: 'frame:' + record.key, what: record.label + ' was removed.', where: record.part, action: 'drop', notify: false });
    }
    return groups;
  };
  function transform(xml, old, next) {
    const root = K.parse(xml), shapes = [];
    const walk = el => { if (rootNames.has(el.localName)) { shapes.push(el); return; } kids(el).forEach(walk); };
    walk(root);
    if (shapes.length <= 1) return K.setBox(xml, next);
    // Alternatives may have one frame in Choice and several members in Fallback.
    const sx = old.w ? next.w / old.w : 1, sy = old.h ? next.h / old.h : 1, edits = [];
    for (const el of shapes) {
      const raw = K.raw(el), b = K.getBox(raw); if (b.x == null || b.y == null) continue;
      const p = X.source.get(el);
      edits.push({ start: p.start, end: p.end, value: K.setBox(raw, { x: next.x + (b.x - old.x) * sx, y: next.y + (b.y - old.y) * sy, w: b.w * sx, h: b.h * sy }) });
    }
    return K.patch(xml, edits);
  }
  function vml(record, rel, shape, ctx, changed) {
    const w = ctx.writer, pkg = K.package(record.fragment.source), copied = !!record.fragment.copy || pkg !== w.pkg;
    if (!copied && !changed) return { id: w.keepRel(ctx.part, rel), spids: new Map() };
    const cache = ctx.frames.vml || (ctx.frames.vml = new Map());
    let state = cache.get(ctx.part);
    if (!state) {
      const original = w.pkg?.rels(ctx.part).find(r => /\/vmlDrawing$/.test(r.type) && !r.external);
      const target = original ? w.target(w.pkg, original.part) : w.name('ppt/drawings', 'vmlDrawing', 'vml');
      const xml = original ? w.pkg.text(original.part) : '<xml xmlns:v="' + K.KNOWN_NS.v + '" xmlns:o="' + K.KNOWN_NS.o + '"/>';
      state = { target, xml, original }; cache.set(ctx.part, state);
      if (original) { w.claim(original.part, 'merged', target); w.carryRels(w.pkg, original.part, target); }
      w.ids.reserve(target, 'vml:_x0000_s', 1024);
    }
    const spids = new Map(), tree = pkg.xml(rel.part);
    const references = all(K.parse(record.fragment.xml)).filter(e => e.localName === 'oleObj' && e.hasAttribute('spid')).map(e => at(e, 'spid'));
    for (const id of new Set(references)) {
      const node = all(tree).find(e => e.namespaceURI === K.KNOWN_NS.v && at(e, 'id') === id);
      if (!node) throw Error('Missing embedded-object preview ' + id);
      let f = K.fragment(node, { pkg, part: rel.part }), types = [];
      const typeIds = new Set(all(node).filter(e => e.namespaceURI === K.KNOWN_NS.v && e.hasAttribute('type')).map(e => at(e, 'type').replace(/^#/, '')));
      for (const type of typeIds) {
        const definition = kids(tree).find(e => e.localName === 'shapetype' && at(e, 'id') === type);
        const prior = kids(K.parse(state.xml)).find(e => e.localName === 'shapetype' && at(e, 'id') === type);
        if (prior && definition && K.hoistNamespaces(K.raw(prior)) === K.hoistNamespaces(K.raw(definition))) {
          f = K.slice(f, f.xml.replace(/\u0001id:(\d+)\u0001/g, (token, i) => {
            const ref = f.ids[+i]; return !ref.definition && ref.kind.startsWith('vmlType') && (ref.prefix || '').replace(/^#/, '') + ref.id === type ? '#' + type : token;
          }));
        } else if (definition) types.push(K.fragment(definition, { pkg, part: rel.part }));
      }
      if (copied) { const bundle = K.duplicate({ f, types }); f = bundle.f; types = bundle.types; }
      const added = types.map(t => w.emit(t, state.target));
      let xml = w.emit(f, state.target);
      if (changed || copied) xml = K.setBox(xml, box(shape));
      spids.set(id, at(K.parse(xml), 'id'));
      if (!copied) {
        const original = all(K.parse(state.xml)).find(e => e.namespaceURI === K.KNOWN_NS.v && at(e, 'id') === id);
        if (original) { const p = X.source.get(original); state.xml = K.patch(state.xml, [{ start: p.start, end: p.end, value: xml }]); }
        else added.push(xml);
      } else added.push(xml);
      if (added.length) state.xml = K.mergeBag(state.xml, { __append: added });
    }
    w.put(state.target, K.hoistNamespaces(state.xml), pkg.type(rel.part));
    const id = w.rels(ctx.part).add(rel.type, K.relative(ctx.part, state.target), false, state.original?.id);
    return { id, spids };
  }
  function dependencies(record, shape, ctx, changed) {
    const w = ctx.writer, f = record.fragment, pkg = K.package(f.source), privateCopy = !!f.copy || f.source !== w.pkg?.id;
    const parts = new Map(), relIds = new Map();
    function copy(part) {
      if (parts.has(part)) return parts.get(part);
      const slash = part.lastIndexOf('/'), dot = part.lastIndexOf('.');
      const target = w.name(part.slice(0, slash), part.slice(slash + 1, dot).replace(/\d+$/, ''), part.slice(dot + 1)); parts.set(part, target);
      const refs = {};
      for (const r of pkg.rels(part)) if (!r.external) refs[r.part] = copy(r.part);
      w.copyPart(pkg, part, target, refs); return target;
    }
    const relation = dep => {
      if (!dep.type) throw Error('Missing relationship ' + dep.id);
      if (!privateCopy) { const id = w.keepRel(ctx.part, dep); if (!id) throw Error('Removed dependency ' + dep.id); return id; }
      const target = dep.external ? dep.target : K.relative(ctx.part, copy(dep.part)) + (dep.fragment || '');
      return w.rels(ctx.part).add(dep.type, target, dep.external);
    };
    const spids = new Map();
    for (const dep of record.related) {
      if (dep.dataPart && /\/diagramDrawing$/.test(dep.type) && !pkg.has(dep.part)) {
        // Some authored diagrams have data/layout/style but a dangling optional
        // drawing-cache reference. Keep the diagram; remove only that reference.
        const tree = pkg.xml(dep.dataPart), edits = [], remove = new Set();
        for (const el of all(tree).filter(e => e.localName === 'dataModelExt' && at(e, 'relId') === dep.id)) {
          let node = el;
          if (node.parentNode?.localName === 'ext' && kids(node.parentNode).length === 1) node = node.parentNode;
          if (node.parentNode?.localName === 'extLst' && kids(node.parentNode).length === 1) node = node.parentNode;
          remove.add(node);
        }
        for (const el of remove) { const p = X.source.get(el); edits.push({ start: p.start, end: p.end, value: '' }); }
        const target = privateCopy ? copy(dep.dataPart) : w.target(pkg, dep.dataPart);
        if (!privateCopy) { w.claim(dep.dataPart, 'merged', target, pkg); w.carryRels(pkg, dep.dataPart, target); }
        w.put(target, K.patch(X.source.get(tree).text, edits), pkg.type(dep.dataPart));
        w.loss({ id: 'diagram-cache:' + ctx.part + ':' + dep.dataPart, what: 'A reference to a SmartArt drawing cache missing from the original file was removed; its data and layout are retained.', where: dep.dataPart, action: 'conversion' });
      } else if (/\/vmlDrawing$/.test(dep.type)) {
        const kept = vml(record, dep, shape, ctx, changed); relIds.set(dep.id, kept.id);
        for (const entry of kept.spids) spids.set(...entry);
      } else relIds.set(dep.id, relation(dep));
    }
    let xml = f.xml.replace(/\u0001rel:(\d+)\u0001/g, (_, i) => relation(f.deps[+i]));
    if (spids.size) {
      const edits = [];
      for (const el of all(K.parse(xml))) if (el.localName === 'oleObj' && spids.has(at(el, 'spid'))) {
        const p = X.source.get(el), a = p.attrs.find(a => a.name === 'spid'); edits.push({ start: a.start, end: a.end, value: spids.get(at(el, 'spid')) });
      }
      xml = K.patch(xml, edits);
    }
    for (const rel of record.related.filter(r => r.dataPart && relIds.has(r.id) && relIds.get(r.id) !== r.id)) {
      let target = parts.get(rel.dataPart);
      if (!target) { target = copy(rel.dataPart); }
      const text = typeof w.parts.get(target) === 'string' ? w.parts.get(target) : X.decode(w.parts.get(target));
      const tree = K.parse(text), edits = [];
      for (const el of all(tree).filter(e => e.localName === 'dataModelExt' && at(e, 'relId') === rel.id)) {
        const p = X.source.get(el); edits.push({ start: p.start, end: p.end, value: K.attributes(K.raw(el), { relId: relIds.get(rel.id) }) });
      }
      w.put(target, K.hoistNamespaces(K.patch(text, edits)), pkg.type(rel.dataPart));
    }
    return K.slice(f, xml);
  }
  F.emit = function (shape, ctx) {
    const record = shape.keep?.frame; if (!record) return null;
    const group = ctx.frames?.get(key(record.fragment));
    const loss = (what, error) => ctx.writer.loss({ id: 'frame:' + key(record.fragment), what, detail: error?.message, where: ctx.part, place: ctx.place, action: 'conversion', notify: !K.sourceMissing(error) });
    if (!group?.intact) { loss(savedAs(record.label, ', because it was edited here.')); return null; }
    if (group.emitted) return '';
    try {
      const fragment = dependencies(record, shape, ctx, group.moved || ctx.inGroup);
      let xml = ctx.writer.emit(fragment, ctx.part);
      if (group.moved || ctx.inGroup) {
        const current = group.entries.length === 1 ? box(shape) : group.bounds;
        xml = transform(xml, record.localBounds, current);
      }
      xml = L.comments.shapeIdentity(shape, ctx, xml);
      group.emitted = true; return xml;
    } catch (error) { group.intact = false; loss(savedAs(record.label, ', because the original couldn\'t be kept.'), error); return null; }
  };
  F.references = function (item) {
    const refs = [];
    const visit = v => {
      if (!v || typeof v !== 'object') return;
      if (v.frame?.fragment) refs.push(...(v.frame.related || []).filter(d => !d.external), ...(v.frame.previews || []).flatMap(r => [r, ...(r.view ? [r.view] : [])]));
      if (v.properties?.previews) refs.push(...v.properties.previews.flatMap(r => [r, ...(r.view ? [r.view] : [])]));
      Object.values(v).forEach(visit);
    };
    visit(item); return refs;
  };
  F.restore = function (item) {
    const images = new Map();
    const lists = item.kind === 'slides' ? item.slides.map(s => s.shapes).concat(Object.values(item.designs || {}).flatMap(d => [d.deco, d.titleDeco, ...Object.values(d.layoutDecos || {})]).filter(Boolean)) : [item.shapes || []];
    for (const list of lists) L.model.walk(list, shape => {
      const record = shape.keep?.frame || shape.keep?.properties; if (!record) return true;
      const replacements = new Map();
      for (const ref of record.previews || []) {
        const pkg = K.package(ref.source), key = ref.source + ':' + ref.part;
        if (!images.has(key)) {
          const viewPkg = ref.view && K.package(ref.view.source), view = viewPkg && new Blob([viewPkg.bytes(ref.view.part)], { type: viewPkg.type(ref.view.part) });
          const id = L.media.add(new Blob([pkg.bytes(ref.part)], { type: pkg.type(ref.part) || 'image/png' }), ref.part.split('/').pop(), view);
          if (ref.size) L.media.get(id).size = ref.size;
          images.set(key, id);
        }
        replacements.set(ref.id, images.get(key)); ref.id = images.get(key);
      }
      const replace = v => {
        if (!v || typeof v !== 'object') return;
        for (const [k, value] of Object.entries(v)) {
          if (k === 'keep') continue;
          if (typeof value === 'string' && replacements.has(value)) v[k] = replacements.get(value); else replace(value);
        }
      };
      replace(shape);
      if (record.signature) { const signature = JSON.parse(record.signature); replace(signature); record.signature = JSON.stringify(signature); }
      if (record.before) replace(record.before);
      if (shape.keep.designValue) { const signature = JSON.parse(shape.keep.designValue); replace(signature); shape.keep.designValue = JSON.stringify(signature); }
      return true;
    });
  };
})();
