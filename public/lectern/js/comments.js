/* Lectern — comment parts belong to slides; authors are an implicit dependency. */
(function () {
  'use strict';
  const L = window.L, K = L.opc, X = L.xmlTree, C = (L.comments = {}), N = K.NS;
  const MODERN = 'http://schemas.microsoft.com/office/powerpoint/2018/8/main';
  const PC = 'http://schemas.microsoft.com/office/powerpoint/2013/main/command';
  const AC = 'http://schemas.microsoft.com/office/drawing/2013/main/command';
  const RT = 'http://schemas.microsoft.com/office/2018/10/relationships/';
  const types = [N.rel + '/comments', RT + 'comments'];
  const authors = [N.rel + '/commentAuthors', RT + 'authors'];
  C.types = types;
  const kids = (e, tag) => Array.from(e?.children || []).filter(c => !tag || c.localName === tag);
  const all = e => e ? [e, ...kids(e).flatMap(all)] : [];
  const at = (e, key) => e?.getAttribute(key);
  const kind = rel => types.indexOf(K.relationshipType(rel.type));
  const related = (pkg, type) => pkg?.rels(pkg.main).find(r => K.relationshipType(r.type) === type && !r.external);
  function guid(seed) {
    const bytes = L.sha.sha1(new TextEncoder().encode(seed)).slice(0, 16);
    bytes[6] = (bytes[6] & 15) | 80; bytes[8] = (bytes[8] & 63) | 128;
    const h = Array.from(bytes, b => b.toString(16).padStart(2, '0')).join('').toUpperCase();
    return '{' + [h.slice(0, 8), h.slice(8, 12), h.slice(12, 16), h.slice(16, 20), h.slice(20)].join('-') + '}';
  }
  function rewrite(xml, change) {
    const edits = [], root = K.parse(xml);
    for (const el of all(root)) {
      const attrs = change(el); if (!attrs || !Object.keys(attrs).length) continue;
      const p = X.source.get(el);
      const updated = X.source.get(K.parse(K.attributes(K.raw(el), attrs)));
      edits.push({ start: p.start, end: p.openEnd, value: updated.text.slice(updated.start, updated.openEnd) });
    }
    return edits.length ? K.patch(X.source.get(root).text, edits) : xml;
  }
  C.read = function (slide, original, pkg, part) {
    const records = pkg.rels(part).filter(r => kind(r) >= 0 && !r.external).map(r => {
      const author = related(pkg, authors[kind(r)]);
      return { source: pkg.id, part: r.part, type: r.type, id: r.id, author: author && { source: pkg.id, part: author.part, type: author.type } };
    });
    if (!records.length) return;
    slide.keep.comments = records;
    const ext = kids(kids(original, 'extLst')[0]).filter(e => all(e).some(n => n.namespaceURI === MODERN && n.localName === 'commentRel'));
    slide.keep.commentExt = ext.map(e => K.fragment(e, { pkg, part }));
    if (!records.some(r => kind(r) === 1)) return;
    const creation = kids(kids(kids(original, 'cSld')[0], 'extLst')[0]).filter(e => all(e).some(n => n.localName === 'creationId'));
    slide.keep.commentCreation = creation.map(K.raw);
    const refs = new Set();
    for (const r of records.filter(r => kind(r) === 1)) {
      let tree; try { tree = pkg.xml(r.part); } catch (_) { continue; }
      for (const el of all(tree)) if (el.namespaceURI === AC && /^(sp|grpSp|graphicFrame|cxnSp|pic|ink)Mk$/.test(el.localName)) refs.add(at(el, 'id'));
    }
    L.model.walk(slide.shapes, s => {
      const id = s.keep?.identity?.ids.find(i => i.kind === 'shape' && i.definition)?.id;
      if (refs.has(id)) s.keep.commentAnchor = { owner: s.id, text: s.tx ? L.txt.plain(s.tx) : null };
      return true;
    });
  };
  C.references = function (item) {
    return (item.slides || []).flatMap(s => (s.keep?.comments || []).flatMap(r => [r, ...(r.author ? [r.author] : [])]));
  };
  C.shapeIdentity = function (shape, ctx, xml) {
    const f = shape.keep?.identity;
    if (!f || !f.copy && f.source === ctx.writer?.pkg?.id && f.part === ctx.part) return xml;
    return rewrite(xml, el => el.localName === 'creationId' && /^\{/.test(at(el, 'id') || '') ? { id: guid('shape:' + shape.id + ':' + at(el, 'id')) } : null);
  };
  C.begin = function (pres, pack) {
    const w = pack.writer, tables = new Map(), creationIds = new Map(), usedCreation = new Set();
    for (const r of pres.pkg?.rels(pres.pkg.main) || []) if (K.relationshipType(r.type) === N.rel + '/slide' && !r.external) {
      let root; try { root = pres.pkg.xml(r.part); } catch (_) { continue; }
      for (const e of all(root)) if (e.localName === 'creationId' && e.hasAttribute('val')) usedCreation.add(+at(e, 'val'));
    }
    function creationId(slide) {
      if (!creationIds.has(slide.id)) {
        const original = !slide.keep?.copy && slide.keep?.source === pres.pkg?.id && (slide.keep.commentCreation || []).flatMap(xml => all(K.parse(xml))).find(e => e.localName === 'creationId' && e.hasAttribute('val'));
        if (original) { creationIds.set(slide.id, +at(original, 'val')); return creationIds.get(slide.id); }
        let id = parseInt(guid('slide:' + slide.id).slice(1, 9), 16);
        while (!id || usedCreation.has(id)) id = (id + 1) >>> 0;
        usedCreation.add(id); creationIds.set(slide.id, id);
      }
      return creationIds.get(slide.id);
    }
    function table(n) {
      if (tables.has(n)) return tables.get(n);
      const rel = related(pres.pkg, authors[n]), source = rel?.part;
      const xml = source ? pres.pkg.text(source) : '<' + (n ? 'authorLst' : 'cmAuthorLst') + ' xmlns="' + (n ? MODERN : N.p) + '"/>';
      const root = K.parse(xml), entries = new Map(kids(root).map(e => [at(e, 'id'), { xml: K.raw(e), last: +(at(e, 'lastIdx') || 0) }]));
      const t = { n, rel, source, xml, ns: root.namespaceURI, entries, maps: new Map(), changed: false };
      // A stale source counter must not allocate an already-used author/index pair.
      if (!n) for (const name of pres.pkg?.names || []) if (pres.pkg.type(name)?.endsWith('.comments+xml')) {
        let root; try { root = pres.pkg.xml(name); } catch (_) { continue; }
        for (const e of kids(root, 'cm')) { const author = entries.get(at(e, 'authorId')); if (author) author.last = Math.max(author.last, +at(e, 'idx') || 0); }
      }
      tables.set(n, t); return t;
    }
    function author(t, record, id) {
      if (record.source === pres.pkg?.id) return id;
      const key = record.source + ':' + record.author?.part + ':' + id;
      if (t.maps.has(key)) return t.maps.get(key);
      const pkg = K.package(record.source), original = kids(record.author && pkg.xml(record.author.part)).find(e => at(e, 'id') === id);
      if (!original) throw Error('Missing comment author ' + id);
      let fresh = t.n ? id : String(Math.max(-1, ...Array.from(t.entries.keys(), Number)) + 1);
      if (t.n && t.entries.has(fresh)) {
        // Reuse a modern author only when its entire record agrees.
        if (K.serialize(K.parse(t.entries.get(fresh).xml)) === K.serialize(K.parse(K.raw(original)))) { t.maps.set(key, fresh); return fresh; }
        fresh = guid('author:' + key);
      }
      t.entries.set(fresh, { xml: K.attributes(K.raw(original), { id: fresh, ...(t.n ? {} : { lastIdx: 0 }) }), last: 0 });
      t.maps.set(key, fresh); t.changed = true; return fresh;
    }
    return { pres, pack, w, table, author, tables, creationId };
  };
  function modernXML(state, slide, record, ctx, xml, copy) {
    const { w } = state, t = state.table(1), shapeMap = new Map();
    L.model.walk(slide.shapes, s => {
      const ref = s.keep?.identity?.ids.find(i => i.definition && i.kind === 'shape');
      if (ref && s.keep.commentAnchor?.owner === s.id && s.keep.source === record.source && s.keep.part === slide.keep.part) shapeMap.set(String(ref.id), s);
      return true;
    });
    const anchorTags = new Set(['sldMkLst', 'deMkLst', 'txMkLst']);
    const tree = K.parse(xml), changes = [];
    for (const cm of kids(tree, 'cm')) {
      const anchors = kids(cm).filter(e => anchorTags.has(e.localName));
      const refs = anchors.flatMap(all).filter(e => e.namespaceURI === AC && /^(sp|grpSp|graphicFrame|cxnSp|pic|ink)Mk$/.test(e.localName));
      const absent = refs.some(e => !shapeMap.has(at(e, 'id')));
      const textChanged = anchors.some(e => e.localName === 'txMkLst') && refs.some(e => {
        const shape = shapeMap.get(at(e, 'id')); return shape?.keep.commentAnchor && shape.keep.commentAnchor.text !== (shape.tx ? L.txt.plain(shape.tx) : null);
      });
      if (absent || textChanged) {
        const sldId = w.slideIds.get(slide.id);
        anchors.forEach((anchor, i) => {
          const p = X.source.get(anchor);
          changes.push({ start: p.start, end: p.end, value: i ? '' : '<pc:sldMkLst xmlns:pc="' + PC + '"><pc:docMk/><pc:sldMk cId="' + state.creationId(slide) + '" sldId="' + sldId + '"/></pc:sldMkLst>' });
        });
        w.loss({ id: 'comment-anchor:' + slide.id + ':' + at(cm, 'id'), what: absent ? 'A comment on a deleted object is now attached to its slide.' : 'A comment on edited text is now attached to its slide.', where: 'Slide comments', action: 'conversion' });
      }
    }
    if (changes.length) xml = K.patch(X.source.get(tree).text, changes);
    return rewrite(xml, el => {
      const attrs = {};
      if (copy && el.namespaceURI === MODERN && ['cm', 'reply'].includes(el.localName)) attrs.id = guid('comment:' + slide.id + ':' + at(el, 'id'));
      if (record.source !== state.pres.pkg?.id) {
        if (el.hasAttribute('authorId')) attrs.authorId = state.author(t, record, at(el, 'authorId'));
        if (el.hasAttribute('assignedTo')) attrs.assignedTo = at(el, 'assignedTo').split(/\s+/).filter(Boolean).map(id => state.author(t, record, id)).join(' ');
      }
      if (el.namespaceURI === PC && el.localName === 'sldMk') {
        if (at(el, 'sldId') !== String(w.slideIds.get(slide.id))) attrs.sldId = w.slideIds.get(slide.id);
        if (copy || !el.hasAttribute('cId') || !slide.keep?.commentCreation?.length) attrs.cId = state.creationId(slide);
      }
      if (el.namespaceURI === AC && /^(sp|grpSp|graphicFrame|cxnSp|pic|ink)Mk$/.test(el.localName)) {
        const shape = shapeMap.get(at(el, 'id'));
        if (shape) {
          if (at(el, 'id') !== String(ctx.nextId(shape.id))) attrs.id = ctx.nextId(shape.id);
          if (shape.keep.identity.copy && el.hasAttribute('creationId')) attrs.creationId = guid('shape:' + shape.id + ':' + at(el, 'creationId'));
        }
      }
      return attrs;
    });
  }
  C.slide = function (state, slide, ctx) {
    const { w, pres } = state, comments = slide.keep?.comments || [], relIds = new Map();
    for (const record of comments) {
      const pkg = K.package(record.source), n = kind(record), copy = !!slide.keep.copy || record.source !== pres.pkg?.id;
      try {
        let xml = pkg.text(record.part), changed = false;
        if (n === 1) { const next = modernXML(state, slide, record, ctx, xml, copy); changed = next !== xml; xml = next; }
        else if (copy) {
          const t = state.table(0), ids = new Map();
          for (const cm of kids(K.parse(xml), 'cm')) {
            const id = state.author(t, record, at(cm, 'authorId')), entry = t.entries.get(id);
            if (!entry) throw Error('Missing comment author ' + id);
            ids.set(at(cm, 'authorId') + ':' + at(cm, 'idx'), { authorId: id, idx: ++entry.last }); t.changed = true;
          }
          xml = rewrite(xml, el => ['cm', 'parentCm'].includes(el.localName) ? ids.get(at(el, 'authorId') + ':' + at(el, 'idx')) : null); changed = true;
        }
        const target = copy ? w.name('ppt/comments', 'comment', 'xml') : w.target(pkg, record.part);
        if (copy) w.copyPart(pkg, record.part, target);
        else w.carry(pkg, record.part);
        if (changed) w.put(target, K.hoistNamespaces(xml), pkg.type(record.part));
        const id = w.rels(ctx.part).add(record.type, K.relative(ctx.part, target), false, copy ? undefined : record.id);
        relIds.set(record.id, id);
      } catch (error) {
        w.loss({ id: 'comments:' + slide.id + ':' + record.part, what: 'Comments on this slide couldn\'t be kept and will be removed.', detail: error.message, where: record.part, place: ctx.place, action: 'drop', notify: !L.opc.sourceMissing(error) });
      }
    }
    const extensions = (slide.keep?.commentExt || []).flatMap(f => {
      // The relationship belongs to this particular slide copy, not the shared source part.
      if (f.deps.some(d => !relIds.has(d.id))) return [];
      return [f.xml.replace(/\u0001rel:(\d+)\u0001/g, (_, i) => relIds.get(f.deps[+i].id))];
    }).join('');
    let creation = (slide.keep?.commentCreation || []).map(xml => slide.keep.copy || slide.keep.source !== pres.pkg?.id
      ? rewrite(xml, e => e.localName === 'creationId' ? { val: state.creationId(slide) } : null) : xml).join('');
    if (!creation && comments.some(r => kind(r) === 1 && relIds.has(r.id))) creation = '<p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="' + state.creationId(slide) + '"/></p:ext>';
    return { ext: extensions ? '<p:extLst>' + extensions + '</p:extLst>' : '', creation: creation ? '<p:extLst>' + creation + '</p:extLst>' : '' };
  };
  C.finish = function (state) {
    const { pres, w, pack } = state;
    for (const t of state.tables.values()) if (t.changed) {
      const target = t.source ? w.target(pres.pkg, t.source) : w.name('ppt', t.n ? 'authors' : 'commentAuthors', 'xml');
      if (t.source) w.claim(t.source, 'merged', target);
      const entries = [...t.entries.values()].map(e => t.n ? e.xml : K.attributes(e.xml, { lastIdx: e.last }));
      const xml = K.mergeBag(t.xml, { ['{' + t.ns + '}' + (t.n ? 'author' : 'cmAuthor')]: entries });
      w.put(target, K.hoistNamespaces(xml), pres.pkg?.type(t.source) || (t.n ? 'application/vnd.ms-powerpoint.authors+xml' : 'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml'));
      if (t.source) w.carryRels(pres.pkg, t.source, target);
      const owner = pack.part('ppt/presentation.xml');
      w.rels(owner).add(t.rel?.type || authors[t.n], K.relative(owner, target), false, t.rel?.id);
    }
    // A source part is dropped only when none of its original owners survives.
    const live = new Set(pres.slides.filter(s => s.keep?.source && s.keep.source === pres.pkg?.id && !s.keep.copy).flatMap(s => (s.keep.comments || []).map(r => r.part)));
    for (const name of pres.pkg?.names || []) for (const r of pres.pkg.rels(name).filter(r => kind(r) >= 0 && !r.external)) {
      if (!live.has(r.part)) w.omit(r.part, 'The slide containing these comments was deleted.');
    }
  };
})();
