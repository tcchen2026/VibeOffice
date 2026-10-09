/* Ledger — worksheet objects and their DrawingML/VML representations share an owner. */
(function (root) {
  'use strict';
  const L = root.L, K = L.opc, M = L.model, X = (L.sheetObjects = {}), N = K.NS;
  const V = K.KNOWN_NS.v, EX = K.KNOWN_NS.x;
  const kids = el => Array.from(el?.children || []);
  const all = el => [el, ...Array.from(el?.getElementsByTagName('*') || [])].filter(Boolean);
  const at = (el, name) => el?.getAttribute(name);
  const clone = v => JSON.parse(JSON.stringify(v));
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  const uid = () => root.crypto.randomUUID();
  const shapeId = value => /^_x0000_s\d+$/.test(value || '') ? +value.slice(8) : +value || 0;
  const vmlId = el => shapeId(at(el, 'id')) || shapeId(el.getAttributeNS(K.KNOWN_NS.o, 'spid'));
  const values = d => ({ xml: d.xml, media: d.media, imgLink: d.imgLink, crop: d.crop, text: d.text,
    fill: d.fill, line: d.line, lineW: d.lineW, geom: d.geom, font: d.font, dirty: d.dirty });
  const noteValues = cm => ({ r: cm.r, c: cm.c, w: cm.w, h: cm.h, anchor: cm.anchor, visible: cm.visible, fill: cm.fill });
  const active = d => d.objectKeep && same(values(d), d.objectKeep.values);
  const objectNames = new Set(['controls', 'oleObjects', 'customProperties']);
  const item = el => ['control', 'oleObject'].includes(el.localName) && at(el, 'shapeId') != null;
  const objectIds = el => [...new Set(all(el).filter(item).map(e => +at(e, 'shapeId')))];

  X.capture = function (sh, part, xml) {
    if (!sh.wb.pkg) return;
    const tree = K.parse(xml), groups = [];
    for (const el of kids(tree)) {
      const names = [...new Set(all(el).filter(e => objectNames.has(e.localName)).map(e => e.localName))];
      if (!names.length) continue;
      groups.push({ key: uid(), names, ids: objectIds(el), fragment: K.fragment(el, { pkg: sh.wb.pkg, part }) });
    }
    if (groups.length) sh.extra.objectGroups = groups;
  };
  function xmlAnchor(el) {
    const a = all(el).find(e => e.localName === 'anchor'), from = a && kids(a).find(e => e.localName === 'from'), to = a && kids(a).find(e => e.localName === 'to');
    const point = e => {
      const n = name => +(kids(e).find(c => c.localName === name)?.textContent || 0);
      return { c: n('col'), r: n('row'), cOff: n('colOff') / 12700, rOff: n('rowOff') / 12700 };
    };
    return from && to ? { type: 'two', from: point(from), to: point(to), editAs: at(a, 'sizeWithCells') === '1' ? 'twoCell' : at(a, 'moveWithCells') === '1' ? 'oneCell' : 'absolute' } : null;
  }
  function vmlAnchor(el) {
    const data = all(el).find(e => e.localName === 'ClientData' && e.namespaceURI === EX), anchor = kids(data).find(e => e.localName === 'Anchor');
    const p = anchor?.textContent.split(',').map(Number);
    if (p?.length === 8 && p.every(Number.isFinite)) return { type: 'two', from: { c: p[0], cOff: p[1] * 0.75, r: p[2], rOff: p[3] * 0.75 }, to: { c: p[4], cOff: p[5] * 0.75, r: p[6], rOff: p[7] * 0.75 },
      editAs: kids(data).some(e => e.localName === 'SizeWithCells') ? 'twoCell' : kids(data).some(e => e.localName === 'MoveWithCells') ? 'oneCell' : 'absolute' };
    const box = K.getBox(K.raw(el));
    return { type: 'abs', x: box.x || 0, y: box.y || 0, w: box.w || 72, h: box.h || 24 };
  }
  function owner(sh, id, anchor) {
    let d = sh.drawings.find(d => d.id === id);
    if (!d) { d = { kind: 'legacy', id, name: 'Preserved object ' + id, hidden: true, anchor: anchor || { type: 'abs', x: 0, y: 0, w: 72, h: 24 } }; sh.drawings.push(d); }
    return d;
  }
  X.read = function (sh) {
    const pkg = sh.wb.pkg, part = sh.extra.ooxmlPart; if (!pkg) return;
    for (const g of sh.extra.objectGroups || []) {
      const tree = K.parse(g.fragment.xml);
      for (const id of g.ids) {
        const el = all(tree).find(e => item(e) && +at(e, 'shapeId') === id), d = owner(sh, id, xmlAnchor(el));
        if (!d.objectKeep) d.objectKeep = { source: pkg.id, part, id, groups: [], values: clone(values(d)), anchor: clone(d.anchor) };
        d.objectKeep.groups.push(g);
      }
    }
    const rel = pkg.rels(part).find(r => /\/vmlDrawing$/.test(r.type) && !r.external);
    if (!rel) return;
    const state = sh.extra.vml = { source: pkg.id, part: rel.part, sheet: sh.id, relId: rel.id };
    let tree; try { tree = pkg.xml(rel.part); } catch (_) { state.unparsed = true; state.notes = clone([...sh.comments.values()].map(noteValues)); return; }
    const shapes = kids(tree).filter(e => e.namespaceURI === V && !['shapetype', 'background'].includes(e.localName));
    state.shapes = shapes.map(e => vmlId(e));
    for (const el of shapes) {
      const client = all(el).find(e => e.localName === 'ClientData' && e.namespaceURI === EX);
      const vml = { source: pkg.id, part: rel.part, id: vmlId(el), fragment: K.fragment(el, { pkg, part: rel.part }) };
      if (at(client, 'ObjectType') === 'Note') {
        const r = +(kids(client).find(e => e.localName === 'Row')?.textContent || 0), c = +(kids(client).find(e => e.localName === 'Column')?.textContent || 0), cm = sh.comments.get(M.key(r, c));
        if (cm) { vml.values = clone(noteValues(cm)); cm.keepVml = vml; continue; }
      }
      const d = owner(sh, vml.id, vmlAnchor(el));
      if (!d.objectKeep) d.objectKeep = { source: pkg.id, part, id: vml.id, groups: [], values: clone(values(d)), anchor: clone(d.anchor) };
      d.objectKeep.vml = vml;
    }
  };
  X.copy = d => {
    const value = clone(d);
    if (value.objectKeep) value.objectKeep.copy = uid();
    if (value.keep?.picture) value.keep.picture = K.duplicate(value.keep.picture);
    L.slicers.copy(value);
    return value;
  };
  X.copySheet = (from, to) => {
    if (from.extra.vml) to.extra.vml = clone(from.extra.vml);
    if (from.extra.objectGroups) to.extra.objectGroups = clone(from.extra.objectGroups);
    const copies = new Map();
    for (const d of to.drawings) if (d.objectKeep) {
      const prior = d.objectKeep.copy || '';
      if (!copies.has(prior)) copies.set(prior, uid());
      d.objectKeep = { ...d.objectKeep, copy: copies.get(prior) };
    }
    for (const d of to.drawings) L.slicers.copy(d, to);
    for (const d of to.drawings) if (d.keep?.picture) d.keep.picture = K.duplicate(d.keep.picture);
  };
  X.begin = function (sh, pack, owner, vml) {
    const ids = new Map(vml.objects), used = new Set(ids.values()), w = pack.writer;
    const names = new Map(), named = [], takenNames = new Set(), activeX = new Set();
    for (const d of sh.drawings) if (active(d)) {
      const entries = d.objectKeep.groups.flatMap(g => all(K.parse(g.fragment.xml)).filter(e => item(e) && +at(e, 'shapeId') === d.objectKeep.id).map(el => ({ el, g })));
      const el = entries.find(e => at(e.el, 'name'))?.el; if (!el) continue;
      if (entries.some(({ el, g }) => K.slice(g.fragment, K.raw(el)).deps.some(r => /\/control$/.test(r.type)))) activeX.add(d);
      const name = at(el, 'name'); named.push({ d, name });
      if (!d.objectKeep.copy) { names.set(d, name); takenNames.add(name.toLowerCase()); }
    }
    for (const { d, name } of named) if (!names.has(d)) {
      let stem = activeX.has(d) ? name.replace(/[^\p{L}\p{N}_]/gu, '_') : name;
      if (activeX.has(d) && !/^\p{L}/u.test(stem)) stem = 'Control' + stem;
      let fresh = stem, n = +(stem.match(/\d+$/)?.[0] || 0), base = stem.replace(/\d+$/, '');
      if (base === stem && !activeX.has(d)) base += ' ';
      while (takenNames.has(fresh.toLowerCase())) fresh = base + ++n;
      takenNames.add(fresh.toLowerCase()); names.set(d, fresh);
    }
    let next = Math.max(1024, ...used, ...sh.drawings.map(d => d.id || 0));
    for (const d of sh.drawings) {
      let id = ids.get(d) || d.id;
      if (!ids.has(d) && (!(id > 0) || used.has(id) || d.objectKeep && (id < 1025 || id > 268435456))) {
        do { next = next < 268435456 ? next + 1 : 1025; } while (used.has(next));
        id = next;
      }
      used.add(id); ids.set(d, id);
      w.ids.reserve(owner, 'shape', id);
      if (d.objectKeep && !active(d)) w.loss({ id: 'worksheet-object:' + sh.id + ':' + id, what: 'An embedded object or control you edited will be saved as a picture and will no longer work.', where: sh.name + ': ' + d.name, place: sh.name + ': ' + d.name, action: 'conversion', notify: true });
    }
    return { sh, pack, ids, used, next, owner, vml, names, activeX, copiedParts: new Map() };
  };
  // Strip only the linked top-level identity from generic ID remapping. Its
  // drawing, worksheet entry and VML preview must use the same chosen value.
  function linked(fragment, old, id, vml) {
    const match = i => vml ? /^vml:/.test(i.kind) && +i.id === old : i.kind === 'shape' && +i.id === old;
    return K.slice(fragment, fragment.xml.replace(/\u0001id:(\d+)\u0001/g, (token, n) => match(fragment.ids[+n]) ? (fragment.ids[+n].prefix || '') + id : token));
  }
  X.drawing = function (d, fragment, state, size) {
    if (!d.objectKeep) return fragment;
    const id = state.ids.get(d), old = d.objectKeep.id;
    let f = linked(fragment, old, id, false);
    const tree = K.parse(f.xml), patches = [];
    const name = state.names.get(d);
    if (name && d.objectKeep.copy) for (const e of all(tree)) if (e.localName === 'cNvPr' && +at(e, 'id') === id && at(e, 'name') !== name) {
      const p = L.xmlTree.source.get(e), updated = L.xmlTree.source.get(K.parse(K.attributes(K.raw(e), { name })));
      patches.push({ start: p.start, end: p.openEnd, value: updated.text.slice(updated.start, updated.openEnd) });
    }
    for (const e of all(tree)) if (e.localName === 'compatExt' && at(e, 'spid') === '_x0000_s' + old) {
      const p = L.xmlTree.source.get(e), a = p.attrs.find(a => a.name === 'spid'); patches.push({ start: a.start, end: a.end, value: '_x0000_s' + id });
    }
    f = K.slice(f, K.patch(f.xml, patches));
    if (size && !same(d.anchor, d.objectKeep.anchor)) {
      const pos = d.anchor.type === 'abs' ? d.anchor : position(state.sh, d.anchor.from);
      f = K.slice(f, K.setBox(f.xml, { ...pos, ...size }));
    }
    if (d.objectKeep.copy) f = K.duplicate(f);
    return f;
  };
  X.frame = function (d, state, owner, size) {
    if (!active(d) || !d.keep?.frame) return '';
    return state.pack.writer.emit(X.drawing(d, d.keep.frame, state, size), owner);
  };
  const pictureValues = d => Object.fromEntries(['media', 'imgLink', 'crop', 'lineXml', 'rot', 'flipH', 'flipV', 'name', 'descr', 'hidden', 'link'].map(key => [key, d[key]]));
  X.keepPicture = (d, fragment) => { d.keep.picture = { fragment, id: d.id, anchor: clone(d.anchor), values: clone(pictureValues(d)) }; };
  X.picture = function (d, generated, state, owner) {
    const keep = d.keep?.picture; if (!keep) return generated;
    let f = linked(keep.fragment, keep.id, state.ids.get(d), false), xml = K.transitionalXML(f.xml);
    const current = pictureValues(d), changed = key => !same(current[key], keep.values[key]);
    const fresh = K.parse('<root xmlns:xdr="' + K.KNOWN_NS.xdr + '" xmlns:a="' + N.a + '" xmlns:r="' + N.rel + '">' + generated + '</root>');
    const generatedChild = tag => { const el = all(fresh).find(e => e.localName === tag); return el ? K.raw(el) : ''; };
    const patch = (tag, fn) => {
      const tree = K.parse(xml), el = all(tree).find(e => e.localName === tag); if (!el) return;
      const p = L.xmlTree.source.get(el); xml = K.patch(xml, [{ start: p.start, end: p.end, value: fn(K.raw(el)) }]);
    };
    patch('cNvPr', raw => {
      const attrs = {}; for (const key of ['name', 'descr', 'hidden']) if (changed(key)) attrs[key] = key === 'hidden' ? (d.hidden ? '1' : undefined) : d[key];
      raw = K.attributes(raw, attrs);
      if (changed('link')) {
        const id = d.link && state.pack.writer.rels(owner).add(N.rel + '/hyperlink', d.link, true);
        raw = K.merge(raw, { ['{' + N.a + '}hlinkClick']: id ? '<a:hlinkClick xmlns:a="' + N.a + '" xmlns:r="' + N.rel + '" r:id="' + L.xml.esc(id) + '"/>' : '' }, 'a:CT_NonVisualDrawingProps');
      }
      return raw;
    });
    const fill = {};
    if (changed('crop')) fill['{' + N.a + '}srcRect'] = generatedChild('srcRect');
    if (changed('media') || changed('imgLink')) {
      fill['{' + N.a + '}blip'] = generatedChild('blip');
      const blip = all(K.parse(xml)).find(e => e.localName === 'blip');
      if (kids(blip).length) state.pack.writer.loss({ id: 'picture-content:' + state.sh.id + ':' + state.ids.get(d), what: 'Replacing this picture removes effects and extensions attached to its original image.', where: state.sh.name + ': ' + (d.name || 'Picture'), action: 'conversion' });
    }
    if (Object.keys(fill).length) patch('blipFill', raw => K.merge(raw, fill, 'a:CT_BlipFillProperties'));
    if (changed('lineXml')) patch('spPr', raw => K.merge(raw, { ['{' + N.a + '}ln']: d.lineXml || '' }, 'a:CT_ShapeProperties'));
    const box = {};
    if (!same(d.anchor, keep.anchor)) {
      const bounds = a => { const p = a.type === 'abs' ? a : position(state.sh, a.from), q = a.type === 'two' ? position(state.sh, a.to) : null;
        return { x: p.x, y: p.y, w: q ? q.x - p.x : a.w, h: q ? q.y - p.y : a.h }; };
      const before = bounds(keep.anchor), after = bounds(d.anchor), source = K.getBox(xml);
      for (const key of ['x', 'y', 'w', 'h']) if (before[key] !== after[key]) box[key] = (source[key] ?? before[key]) + after[key] - before[key];
    }
    for (const key of ['rot', 'flipH', 'flipV']) if (changed(key)) box[key] = d[key] || 0;
    if (Object.keys(box).length) xml = K.setBox(xml, box);
    f = K.slice(f, xml);
    return state.pack.writer.emit(f, owner);
  };
  function textPatch(xml, el, value, patches) {
    if (!el || el.textContent === String(value)) return;
    const p = L.xmlTree.source.get(el), self = xml[p.openEnd - 2] === '/';
    patches.push(self ? { start: p.openEnd - 2, end: p.openEnd, value: '>' + L.xml.esc(String(value)) + '</' + el.nodeName + '>' } : { start: p.openEnd, end: xml.lastIndexOf('</', p.end - 1), value: L.xml.esc(String(value)) });
  }
  function anchorXML(xml, anchor) {
    if (anchor?.type !== 'two') return xml;
    const tree = K.parse(xml), patches = [];
    for (const a of all(tree).filter(e => e.localName === 'anchor')) for (const side of ['from', 'to']) {
      const p = anchor[side], el = kids(a).find(e => e.localName === side); if (!p || !el) continue;
      for (const [tag, val] of [['col', p.c], ['row', p.r], ['colOff', Math.round((p.cOff || 0) * 12700)], ['rowOff', Math.round((p.rOff || 0) * 12700)]]) textPatch(xml, kids(el).find(e => e.localName === tag), val, patches);
    }
    return K.patch(xml, patches);
  }
  function project(xml, members, state) {
    function change(el) {
      let raw = K.raw(el);
      if (item(el)) {
        const d = members.find(d => d.objectKeep.id === +at(el, 'shapeId')); if (!d) return '';
        const id = state.ids.get(d), attrs = { shapeId: id };
        if (state.names.has(d)) attrs.name = state.names.get(d);
        raw = K.attributes(raw, attrs);
        return same(d.anchor, d.objectKeep.anchor) ? raw : anchorXML(raw, d.anchor);
      }
      const tree = K.parse(raw), patches = [];
      for (const e of kids(tree)) { const value = change(e); if (value === K.raw(e)) continue; const p = L.xmlTree.source.get(e); patches.push({ start: p.start, end: p.end, value }); }
      raw = K.patch(raw, patches);
      if (['controls', 'oleObjects'].includes(el.localName) && !all(K.parse(raw)).some(item)) return '';
      return raw;
    }
    return change(K.parse(xml));
  }
  function copiedParts(fragment, members, state) {
    const copy = members.find(d => d.objectKeep.copy)?.objectKeep.copy;
    if (!copy) return fragment;
    const w = state.pack.writer, replacements = new Map();
    fragment.deps.forEach((dep, i) => {
      if (dep.external || !/\/(control|ctrlProp|oleObject)$/.test(dep.type || '')) return;
      const pkg = K.package(dep.source), key = copy + ':' + dep.source + ':' + dep.part;
      if (!state.copiedParts.has(key)) {
        const targets = {};
        for (const part of K.subtree(pkg, dep.part).parts) {
          const slash = part.lastIndexOf('/'), dot = part.lastIndexOf('.');
          targets[part] = w.name(part.slice(0, slash), part.slice(slash + 1, dot > slash ? dot : undefined).replace(/\d+$/, ''), dot > slash ? part.slice(dot + 1) : '');
        }
        for (const [part, target] of Object.entries(targets)) w.copyPart(pkg, part, target, targets);
        state.copiedParts.set(key, targets[dep.part]);
      }
      replacements.set(i, w.rels(state.owner).add(dep.type, K.relative(state.owner, state.copiedParts.get(key))));
    });
    return K.slice(fragment, fragment.xml.replace(/\u0001rel:(\d+)\u0001/g, (token, n) => replacements.get(+n) || token));
  }
  // Combine copied lists without creating an empty controls/oleObjects element
  // for older Office versions. Keep the alternatives outside the list, as authored.
  function combineLists(values, name) {
    function branches(el, required = []) {
      if (el.localName === name) return [{ required, body: kids(el).map(K.raw).join('') }];
      if (el.localName !== 'AlternateContent') return [{ required, body: '' }];
      const out = [];
      for (const choice of kids(el)) {
        const req = (at(choice, 'Requires') || '').split(/\s+/).filter(Boolean).map(p => choice.lookupNamespaceURI(p));
        const content = kids(choice)[0];
        out.push(...(content ? branches(content, required.concat(req)) : [{ required, body: '' }]));
      }
      if (!kids(el).some(e => e.localName === 'Fallback')) out.push({ required, body: '' });
      return out;
    }
    let combined = [{ required: [], body: '' }];
    for (const xml of values) {
      const next = new Map();
      for (const a of combined) for (const b of branches(K.parse(xml))) {
        const required = [...new Set(a.required.concat(b.required))].sort(), key = required.join(' ');
        if (!next.has(key)) next.set(key, { required, body: a.body + b.body });
      }
      combined = [...next.values()];
    }
    const list = body => body ? '<' + name + ' xmlns="' + N.s + '">' + body + '</' + name + '>' : '';
    if (combined.length === 1 && !combined[0].required.length) return list(combined[0].body);
    const uris = [...new Set(combined.flatMap(b => b.required))];
    return '<mc:AlternateContent xmlns:mc="' + N.mc + '"' + uris.map((u, i) => ' xmlns:kr' + i + '="' + L.xml.esc(u) + '"').join('') + '>' + combined.map(b => b.required.length ?
      '<mc:Choice Requires="' + b.required.map(u => 'kr' + uris.indexOf(u)).join(' ') + '">' + list(b.body) + '</mc:Choice>' :
      '<mc:Fallback>' + list(b.body) + '</mc:Fallback>').join('') + '</mc:AlternateContent>';
  }
  X.sheetXML = function (name, state) {
    const { sh, pack } = state, groups = new Map();
    for (const d of sh.drawings) if (active(d)) for (const g of d.objectKeep.groups) if (g.names[0] === name) {
      const key = g.key + ':' + (d.objectKeep.copy || '');
      if (!groups.has(key)) groups.set(key, { group: g, members: [] }); groups.get(key).members.push(d);
    }
    if (name === 'customProperties') for (const g of sh.extra.objectGroups || []) if (g.names[0] === name && !groups.has(g.key + ':')) groups.set(g.key + ':', { group: g, members: [] });
    const values = [];
    for (const { group: g, members } of groups.values()) {
      const unchanged = g.ids.length === members.length && members.every(d => !d.objectKeep.copy && state.ids.get(d) === d.objectKeep.id && same(d.anchor, d.objectKeep.anchor));
      const xml = unchanged || !g.ids.length ? g.fragment.xml : project(g.fragment.xml, members, state);
      if (!xml) continue;
      try { values.push(pack.writer.emit(copiedParts(K.slice(g.fragment, xml), members, state), state.owner)); }
      catch (error) { pack.writer.loss({ id: 'sheet-object:' + g.key, what: 'An object on this sheet couldn\'t be kept and will be removed.', detail: error.message, where: sh.name, place: sh.name, action: 'drop', notify: !K.sourceMissing(error) }); }
    }
    if (values.length <= 1) return values.join('');
    return combineLists(values, name);
  };
  X.hasVML = sh => !!sh.extra.vml || sh.drawings.some(d => active(d) && d.objectKeep.vml) || [...sh.comments.values()].some(cm => cm.keepVml);
  function position(sh, p) {
    let x = p.cOff || 0, y = p.rOff || 0;
    for (let c = 0; c < p.c; c++) x += M.colPx(sh, c) * 0.75;
    for (let r = 0; r < p.r; r++) y += M.rowPt(sh, r);
    return { x, y };
  }
  function vmlGeometry(xml, anchor, before, sh) {
    if (same(anchor, before)) return xml;
    const tree = K.parse(xml), patches = [], client = all(tree).find(e => e.localName === 'ClientData' && e.namespaceURI === EX);
    if (anchor?.type === 'two') {
      const a = anchor.from, b = anchor.to, values = [a.c, Math.round((a.cOff || 0) / 0.75), a.r, Math.round((a.rOff || 0) / 0.75), b.c, Math.round((b.cOff || 0) / 0.75), b.r, Math.round((b.rOff || 0) / 0.75)];
      textPatch(xml, kids(client).find(e => e.localName === 'Anchor'), values.join(', '), patches);
      xml = K.patch(xml, patches);
      const p = position(sh, a), q = position(sh, b);
      return K.setBox(xml, { x: p.x, y: p.y, w: Math.max(1, q.x - p.x), h: Math.max(1, q.y - p.y) });
    }
    return anchor?.type === 'abs' ? K.setBox(xml, anchor) : xml;
  }
  // VML shape ids come in blocks of 1024 named by o:idmap; each sheet's drawing owns its block(s).
  const blocks = tree => (at(all(tree).find(e => e.localName === 'idmap' && e.namespaceURI === K.KNOWN_NS.o), 'data') || '').split(/[\s,]+/).map(Number).filter(n => n > 0);
  const vmlKey = (v, copy, id = v.id) => JSON.stringify([v.source, v.part, copy || '', +id]);
  function vmlLinked(fragment, copy, state) {
    return K.slice(fragment, fragment.xml.replace(/\u0001id:(\d+)\u0001/g, (token, n) => {
      const i = fragment.ids[+n], id = i.kind === 'vml:_x0000_s' && state.identities.get(vmlKey({ source: i.source, part: i.scope }, copy, i.id));
      return id ? (i.prefix || '') + id : token;
    }));
  }
  // Allocate linked drawing/worksheet/VML identities together. Original notes and
  // objects keep their ids; only new or copied members need fresh slots.
  X.idmaps = function (wb) {
    const own = new Map(), taken = new Set(), out = new Map();
    for (const sh of wb.sheets) {
      const v = sh.extra.vml; if (!v || v.unparsed || v.sheet !== sh.id || v.source !== wb.pkg?.id) continue;
      let b = [], reserved = []; try { const tree = K.package(v.source).xml(v.part); b = blocks(tree); reserved = all(tree).filter(e => e.namespaceURI === V && e.localName !== 'shapetype').map(vmlId); } catch (_) {}
      if (b.length && !b.some(n => taken.has(n))) { own.set(sh, { original: v, list: b, reserved }); b.forEach(n => taken.add(n)); }
    }
    const fresh = after => { let b = after + 1; while (taken.has(b)) b++; taken.add(b); return b; };
    for (const sh of wb.sheets) {
      const o = own.get(sh), list = o ? o.list.slice() : [fresh(0)], used = new Set(o?.reserved), claimed = new Set();
      let next = Math.max(list[0] * 1024, ...used);
      const state = { blocks: list, objects: new Map(), notes: new Map(), identities: new Map() }, requests = new Map(), bindings = [];
      const member = (object, v, copied, map) => {
        const key = v?.id > 0 ? vmlKey(v, copied) : object;
        if (map) bindings.push({ object, map, key });
        if (requests.has(key)) return;
        const keep = o && !copied && v?.source === o.original.source && v.part === o.original.part && v.id > 0 && !claimed.has(v.id) && list.includes(Math.floor(v.id / 1024));
        requests.set(key, keep ? v.id : 0);
        if (keep) claimed.add(v.id);
      };
      for (const d of sh.drawings) if (active(d) && d.objectKeep.vml) {
        const v = d.objectKeep.vml, copy = d.objectKeep.copy;
        member(d, v, copy, state.objects);
        for (const i of v.fragment.ids) if (i.definition && i.kind === 'vml:_x0000_s') member(null, { source: i.source, part: i.scope, id: +i.id }, copy);
      }
      // A nested control has its own worksheet entry but its VML is owned by a
      // group. Give the entry the identity already allocated to that child.
      for (const d of sh.drawings) if (active(d) && !d.objectKeep.vml && sh.extra.vml) {
        const v = { ...sh.extra.vml, id: d.objectKeep.id }, key = vmlKey(v, d.objectKeep.copy);
        if (requests.has(key)) bindings.push({ object: d, map: state.objects, key });
      }
      for (const cm of sh.comments.values()) member(cm, cm.keepVml, cm.keepVml?.fragment.copy, state.notes);
      for (const [key, retained] of requests) {
        if (retained) { state.identities.set(key, retained); continue; }
        do {
          next++;
          if (next % 1024 === 0 || !list.includes(Math.floor(next / 1024))) {
            let block = Math.min(...list.filter(b => b * 1024 + 1 > next));
            if (!Number.isFinite(block)) { block = fresh(Math.max(...list)); list.push(block); }
            next = block * 1024 + 1;
          }
        } while (used.has(next));
        used.add(next); state.identities.set(key, next);
      }
      for (const { object, map, key } of bindings) map.set(object, state.identities.get(key));
      out.set(sh, state);
    }
    return out;
  };
  X.vml = function (generated, base, state) {
    const { sh, pack } = state, w = pack.writer, original = sh.extra.vml, source = original && K.package(original.source);
    const members = sh.drawings.filter(d => active(d) && d.objectKeep.vml), notes = [...sh.comments.values()];
    if (original?.unparsed && same(original.notes, notes.map(noteValues))) { pack.bind(base, original.part, 'opaque'); return source.bytes(original.part); }
    if (!members.length && !notes.length) return null;
    const sameOwner = original && original.sheet === sh.id && original.source === sh.wb.pkg?.id;
    const originalRoot = original && !original.unparsed && source.xml(original.part);
    const from = sameOwner ? original.part : null;
    pack.bind(base, from, 'regenerated');
    const target = pack.part(base);
    // A copied or moved drawing takes this sheet's blocks; its notes are renumbered into them.
    const alloc = state.vml.blocks, own = originalRoot ? blocks(originalRoot) : [];
    const rebase = !!originalRoot && !(sameOwner && own[0] === alloc[0]), extend = !!originalRoot && alloc.join(',') !== own.join(',');
    const shapes = [], supports = new Map(), rawNotes = kids(K.parse(generated)).filter(e => e.namespaceURI === V && e.localName === 'shape');
    const support = (pkg, part) => {
      const key = pkg.id + ':' + part; if (supports.has(key)) return;
      supports.set(key, { pkg, part, tree: pkg.xml(part) });
    };
    let unchanged = !!sameOwner && !original.unparsed;
    const controlNames = new Map([...state.activeX].filter(d => d.objectKeep.copy).map(d => [state.ids.get(d), state.names.get(d)]));
    for (const d of members) {
      const v = d.objectKeep.vml, pkg = K.package(v.source); support(pkg, v.part);
      let f = vmlLinked(v.fragment, d.objectKeep.copy, state.vml);
      // ActiveX binds its worksheet name to v:shape@id; o:spid remains the
      // numeric preview identity. This also covers controls inside VML groups.
      const patches = [];
      for (const el of all(K.parse(f.xml))) if (el.namespaceURI === V && el.localName !== 'shapetype') {
        const name = controlNames.get(vmlId(el)); if (!name) continue;
        const pos = L.xmlTree.source.get(el), next = L.xmlTree.source.get(K.parse(K.attributes(K.raw(el), { id: name })));
        patches.push({ start: pos.start, end: pos.openEnd, value: next.text.slice(next.start, next.openEnd) });
      }
      f = K.slice(f, K.patch(f.xml, patches));
      f = K.slice(f, vmlGeometry(f.xml, d.anchor, d.objectKeep.anchor, sh));
      if (d.objectKeep.copy) f = K.duplicate(f);
      unchanged &&= v.part === original?.part && v.source === source?.id && v.id === state.ids.get(d) && same(d.anchor, d.objectKeep.anchor) && !d.objectKeep.copy;
      shapes.push({ f });
    }
    for (let i = 0; i < notes.length; i++) {
      const cm = notes[i], v = cm.keepVml;
      const id = state.vml.notes.get(cm);
      if (v) {
        const pkg = K.package(v.source); support(pkg, v.part);
        let f = vmlLinked(v.fragment, v.fragment.copy, state.vml), xml = f.xml;
        const old = v.values, tree = K.parse(xml), patches = [];
        const client = all(tree).find(e => e.localName === 'ClientData' && e.namespaceURI === EX);
        textPatch(xml, kids(client).find(e => e.localName === 'Row'), cm.r, patches);
        textPatch(xml, kids(client).find(e => e.localName === 'Column'), cm.c, patches);
        const a = cm.anchor?.slice();
        if (a?.length === 8) { a[0] += cm.c - old.c; a[4] += cm.c - old.c; a[2] += cm.r - old.r; a[6] += cm.r - old.r; textPatch(xml, kids(client).find(e => e.localName === 'Anchor'), a.join(', '), patches); }
        xml = K.patch(xml, patches);
        const attrs = {}, shape = K.parse(xml);
        if (cm.fill !== old.fill && cm.fill) attrs.fillcolor = cm.fill;
        if (cm.visible !== old.visible) {
          const style = at(shape, 'style') || '', visibility = 'visibility:' + (cm.visible ? 'visible' : 'hidden');
          attrs.style = /visibility\s*:/.test(style) ? style.replace(/visibility\s*:[^;]+/i, visibility) : style + ';' + visibility;
          const data = all(shape).find(e => e.localName === 'ClientData' && e.namespaceURI === EX), p = data && L.xmlTree.source.get(data);
          if (p) xml = K.patch(xml, [{ start: p.start, end: p.end, value: K.mergeBag(K.raw(data), { ['{' + EX + '}Visible']: cm.visible ? '<x:Visible xmlns:x="' + EX + '"/>' : '' }) }]);
        }
        xml = K.attributes(xml, attrs);
        if (cm.w !== old.w || cm.h !== old.h) xml = K.setBox(xml, { w: cm.w, h: cm.h });
        shapes.push({ f: K.slice(f, xml) });
        unchanged &&= v.part === original?.part && v.source === source?.id && v.id === id && same(noteValues(cm), old);
      } else { shapes.push({ xml: K.attributes(K.raw(rawNotes[i]), { id: '_x0000_s' + id }) }); unchanged = false; }
    }
    const originalIds = original?.shapes || [];
    unchanged &&= shapes.length === originalIds.length && !rebase && !extend;
    if (unchanged) { pack.bind(base, original.part, 'opaque'); return source.bytes(original.part); }
    if (original?.unparsed) w.loss({ id: 'vml:' + sh.id, what: 'The damaged VML drawing could not be merged; its readable notes were rebuilt.', where: sh.name, action: 'conversion' });
    let xml = originalRoot ? K.raw(originalRoot) : K.raw(K.parse(generated));
    const tree = K.parse(xml), changes = [];
    for (const e of kids(tree)) if (e.namespaceURI === V && !['shapetype', 'background'].includes(e.localName)) { const p = L.xmlTree.source.get(e); changes.push({ start: p.start, end: p.end, value: '' }); }
    const idmap = (rebase || extend) && all(tree).find(e => e.localName === 'idmap' && e.namespaceURI === K.KNOWN_NS.o);
    if (idmap) { const p = L.xmlTree.source.get(idmap); changes.push({ start: p.start, end: p.end, value: K.attributes(K.raw(idmap), { data: alloc.join(',') }) }); }
    xml = K.patch(xml, changes);
    const template = K.parse(xml), known = new Set(kids(template).filter(e => e.localName === 'shapetype').map(e => at(e, 'id'))), added = [];
    for (const { pkg, part, tree } of supports.values()) for (const el of kids(tree)) if (el.localName === 'shapetype') {
      if (part === from && pkg === source) continue;
      const f = K.fragment(el, { pkg, part });
      // Bind an identical original type already in the destination root.
      const prior = kids(template).find(e => e.localName === 'shapetype' && at(e, 'id') === at(el, 'id'));
      if (prior && K.hoistNamespaces(K.raw(prior)) === K.hoistNamespaces(K.raw(el))) {
        for (const id of f.ids.filter(i => i.definition)) w.ids.bind(id.source, id.scope, id.kind, id.id, id.id, target);
      } else { const value = w.emit(f, target); added.push(value); known.add(at(K.parse(value), 'id')); }
    }
    if (shapes.some(s => s.xml) && !known.has('_x0000_t202')) added.push(K.raw(kids(K.parse(generated)).find(e => e.localName === 'shapetype')));
    // Original header/shape types can have their own relationships too.
    if (source && originalRoot) {
      const f = K.fragment(K.parse(xml), { pkg: source, part: original.part });
      xml = w.emit(f, target);
    }
    return K.hoistNamespaces(K.mergeBag(xml, { __append: added.concat(shapes.map(s => s.xml || w.emit(s.f, target))) }));
  };
})(typeof window !== 'undefined' ? window : globalThis);
