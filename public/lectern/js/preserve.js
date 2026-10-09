/* Lectern — model-attached OOXML properties. Content parts remain model-generated. */
(function () {
  'use strict';
  const L = window.L, K = L.opc, XML = L.xmlTree, P = (L.preserve = {}), N = K.NS;
  const kids = el => Array.from(el?.children || []);
  const kid = (el, name) => kids(el).find(e => e.localName === name);
  const all = el => { const out = []; const visit = e => { out.push(e); kids(e).forEach(visit); }; if (el) visit(el); return out; };
  const find = (el, name) => all(el).find(e => e.localName === name);
  const value = (o, keys) => JSON.stringify(keys.map(k => o[k] ?? null));
  const GEOMETRY = ['x', 'y', 'w', 'h', 'rot', 'flipH', 'flipV'];
  const CONTENT = ['type', 'media', 'svgMedia', 'linkUrl'];
  const PROPERTIES = ['crop', 'img', 'geom', 'adj', 'path', 'line', 'shadow', 'name', 'alt', 'hidden', 'link'];
  const RT = name => N.rel + '/' + name;
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  P.parts = Object.freeze([
    ['ppt/presProps.xml', RT('presProps'), 'merged'],
    ['ppt/viewProps.xml', RT('viewProps'), 'opaque'],
    ['ppt/tableStyles.xml', RT('tableStyles'), 'opaque'],
    ['ppt/notesMasters/notesMaster1.xml', RT('notesMaster'), 'opaque'],
    ['docProps/core.xml', N.pkg + '/metadata/core-properties', 'merged', true],
    ['docProps/app.xml', RT('extended-properties'), 'merged', true],
  ]);
  P.values = pres => JSON.parse(JSON.stringify({ props: pres.props, W: pres.W, H: pres.H, firstNum: pres.firstNum, show: pres.show }));
  P.readSlideLists = function (pres) {
    const pkg = pres.pkg; if (!pkg) return;
    const tree = pkg.xml(pkg.main), slides = pres.slides, rels = pkg.rels(pkg.main);
    const sections = all(tree).filter(e => e.localName === 'section' && /powerpoint\/2010\/main$/.test(e.namespaceURI));
    pres.keep.sections = sections.map(section => {
      const id = section.getAttribute('id');
      for (const member of kids(kid(section, 'sldIdLst'))) {
        const slide = slides.find(s => s.keep?.sldId === member.getAttribute('id'));
        if (slide) slide.section = id;
      }
      return id;
    });
    pres.keep.customShows = kids(kid(tree, 'custShowLst')).map(show => ({
      id: show.getAttribute('id'), name: show.getAttribute('name'),
      slides: kids(kid(show, 'sldLst')).map(member => {
        const id = member.getAttributeNS(N.rel, 'id') || member.getAttributeNS(N.strictRel, 'id');
        const part = rels.find(r => r.id === id)?.part;
        return slides.find(s => s.keep?.part === part)?.id || null;
      }),
    }));
  };
  P.joinSection = function (pres, rest, at, slides) {
    const sections = pres.keep?.sections;
    const id = sections?.length ? rest[at - 1]?.section || rest[at]?.section || sections[0] : null;
    for (const slide of slides) { if (id) slide.section = id; else delete slide.section; }
  };
  const emptyShows = pres => (pres.keep?.customShows || []).filter(show => show.slides.length && !show.slides.some(id => pres.slides.some(s => s.id === id)));
  P.mergeSlideLists = function (pres, xml, writer) {
    const tree = K.parse(xml), edits = [], live = new Set(pres.slides.map(s => s.id));
    const members = new Map();
    let prior;
    for (let i = 0; i < pres.slides.length; i++) {
      const slide = pres.slides[i];
      // Importers may append directly; ordinary UI insert/move operations assign at edit time.
      const id = slide.section || (!(slide.keep?.source === pres.pkg?.id && !slide.keep.copy) &&
        (prior || pres.slides.slice(i + 1).find(s => s.section)?.section || pres.keep?.sections?.[0]));
      if (id) { if (!members.has(id)) members.set(id, []); members.get(id).push(String(writer.slideIds.get(slide.id))); prior = id; }
    }
    const replace = (el, value) => { const p = XML.source.get(el); edits.push({ start: p.start, end: p.end, value }); };
    for (const section of all(tree).filter(e => e.localName === 'section' && /powerpoint\/2010\/main$/.test(e.namespaceURI))) {
      const list = kid(section, 'sldIdLst'), ns = section.namespaceURI, ids = members.get(section.getAttribute('id')) || [];
      const old = kids(list), before = old.map(e => e.getAttribute('id'));
      if (same(ids, before)) continue;
      const children = ids.map(id => { const el = old.find(e => e.getAttribute('id') === id); return el ? K.raw(el) : '<sldId xmlns="' + ns + '" id="' + id + '"/>'; });
      const next = K.mergeBag(list ? K.raw(list) : '<sldIdLst xmlns="' + ns + '"/>', { ['{' + ns + '}sldId']: children });
      replace(section, K.mergeBag(K.raw(section), { ['{' + ns + '}sldIdLst']: next }));
    }
    const shows = kid(tree, 'custShowLst'), records = pres.keep?.customShows || [];
    if (shows) {
      const children = kids(shows).flatMap((show, i) => {
        const record = records[i], list = kid(show, 'sldLst'); if (!record || !list) return [K.raw(show)];
        const kept = kids(list).filter((el, j) => live.has(record.slides[j]));
        if (kept.length === kids(list).length) return [K.raw(show)];
        if (!kept.length) {
          writer.loss({ id: 'custom-show:' + record.id, what: 'The custom show "' + record.name + '" was removed because all its slides were deleted.', where: 'Custom shows', action: 'drop' });
          return [];
        }
        return [K.mergeBag(K.raw(show), { ['{' + show.namespaceURI + '}sldLst']: K.mergeBag(K.raw(list), { ['{' + list.namespaceURI + '}sld']: kept.map(K.raw) }) })];
      });
      replace(shows, children.length ? K.mergeBag(K.raw(shows), { ['{' + shows.namespaceURI + '}custShow']: children }) : '');
    }
    return K.patch(xml, edits);
  };
  P.designValues = design => JSON.stringify(Object.fromEntries(Object.entries(design).filter(([key]) => key !== 'keep')));
  P.begin = function (pres, format) {
    const pkg = pres.pkg;
    const consumed = new Set([RT('officeDocument'), ...P.parts.map(p => p[1]), ...L.comments.types, ...['slide', 'slideMaster', 'slideLayout', 'notesSlide', 'theme', 'image', 'chart', 'hyperlink', 'tags'].map(RT)]);
    const output = K.output(pkg, { doc: pres, format, contentType: P.mediaType,
      consumes: (base, rel, type) => consumed.has(type) || L.frames.consumes(pres, base, rel),
      convert: (base, source, data) => /^ppt\/(theme\/|notesMasters\/|notesSlides\/|viewProps\.xml$|tableStyles\.xml$)/.test(base) ? data : null,
      merge: (base, source, data, writer) => P.mergePackage(pres, base, source, data, writer), audit: P.reportLosses });
    output.bind('ppt/presentation.xml', pkg?.main, 'merged');
    for (const [base, type, mode, root] of P.parts) {
      const rel = pkg?.rels(root ? '' : pkg.main).find(r => K.relationshipType(r.type) === type && !r.external);
      output.bind(base, rel?.part, mode);
    }
    return output;
  };
  P.reportLosses = writer => {
    const pkg = writer.pkg, losses = writer.doc.losses || [];
    if (pkg) for (const record of writer.doc.keep?.frames || []) {
      const notice = losses.find(e => e.phase === 'save' && (e.id === 'frame:' + record.key || e.id.startsWith('frame:' + pkg.id + '|' + record.part + '|' + record.key + '|')));
      if (notice) writer.coverLoss(notice.id, pkg.rels(record.part).filter(r => !r.external && record.rels.includes(r.id)).map(r => r.part));
    }
    K.reportFeatures(writer, {
      accepts: el => ['sld', 'notes', 'sldMaster', 'sldLayout'].includes(el.localName) && [N.p, 'http://purl.oclc.org/ooxml/presentationml/main'].includes(el.namespaceURI),
      classify: el => ['http://purl.oclc.org/ooxml/presentationml/main', N.p].includes(el.namespaceURI) && el.localName === 'control' ? 'controls' :
        ['http://purl.oclc.org/ooxml/drawingml/main', N.a].includes(el.namespaceURI) && el.localName === 'fld' ? 'fields' : null,
      labels: { controls: 'Some original form controls were converted to previews or removed', fields: 'Some original text fields were converted to ordinary text' },
    });
    if (writer.doc.repaired?.parts?.length) writer.loss({ id: 'read:damaged', what: 'Part of this file is damaged. Whatever couldn\'t be read from it won\'t be saved.', notify: true, where: writer.doc.repaired.parts.join(', '), action: 'conversion' });
  };
  P.mergePackage = function (pres, base, source, generated, writer) {
    const fragment = K.fragment(writer.pkg.xml(source), { pkg: writer.pkg, part: source });
    const before = pres.keep?.values || {}, fresh = K.parse(generated);
    let xml = fragment.xml;
    if (base === 'ppt/presentation.xml') {
      const replacements = {};
      const replace = tag => { replacements['{' + N.p + '}' + tag] = kids(fresh).filter(e => e.localName === tag).map(K.raw); };
      for (const tag of ['sldMasterIdLst', 'notesMasterIdLst', 'sldIdLst']) replace(tag);
      if (pres.W !== before.W || pres.H !== before.H) replace('sldSz');
      xml = K.merge(xml, replacements, 'p:CT_Presentation');
      xml = P.mergeSlideLists(pres, xml, writer);
      if (pres.firstNum !== before.firstNum) xml = K.attributes(xml, { firstSlideNum: pres.firstNum || 1 });
    } else if (base === 'ppt/presProps.xml') {
      if (!same(pres.show, before.show)) xml = K.merge(xml, { ['{' + N.p + '}showPr']: kids(fresh).filter(e => e.localName === 'showPr').map(K.raw) }, 'p:CT_PresentationProperties');
      const show = kid(K.parse(xml), 'showPr'), custom = kid(show, 'custShow');
      if (custom && emptyShows(pres).some(s => s.id === custom.getAttribute('id'))) {
        const updated = K.mergeBag(K.raw(show), { ['{' + show.namespaceURI + '}custShow']: '<p:sldAll xmlns:p="' + N.p + '"/>' });
        xml = K.mergeBag(xml, { ['{' + show.namespaceURI + '}showPr']: updated });
      }
    } else {
      const owned = base === 'docProps/core.xml' ? { title: 'title', subject: 'subject', author: 'creator', keywords: 'keywords', comments: 'description', category: 'category', created: 'created', revision: 'revision' } : { company: 'Company' };
      xml = K.properties(xml, generated, pres.props, before.props, owned);
    }
    return writer.emit(K.slice(fragment, xml), source);
  };
  P.mutations = Object.freeze({ text: 'keep', geometry: 'update', crop: 'replace crop', shadow: 'replace outerShdw', picture: 'conversion', duplicate: 'remap', delete: 'drop', undo: 'restore' });
  P.anims = slide => JSON.stringify(slide.anims || []);
  P.mediaType = (pkg, part) => ({ avi: 'video/x-msvideo', mp3: 'audio/mpeg', mp4: 'video/mp4', wav: 'audio/wav', wmv: 'video/x-ms-wmv', wma: 'audio/x-ms-wma', mov: 'video/quicktime' })[part.split('.').pop().toLowerCase()];
  P.isMedia = el => all(el).some(e => ['videoFile', 'audioFile', 'wavAudioFile', 'audioCd'].includes(e.localName) || e.localName === 'media' && /powerpoint/.test(e.namespaceURI || ''));
  P.originalShapes = function (pkg, part) {
    const tree = pkg.xml(part), map = new Map();
    for (const el of all(tree)) if (el.localName === 'cNvPr') {
      const shape = el.parentNode?.parentNode;
      if (shape?.localName === 'spTree') continue;
      const prior = map.get(el.getAttribute('id'));
      // An OLE preview can repeat its outer frame's ID. The frame owns that
      // identity; sibling Choice/Fallback definitions still prefer the fallback.
      if (shape && (!prior || !all(prior).includes(shape))) map.set(el.getAttribute('id'), shape);
    }
    return { tree, map };
  };
  P.shape = function (shape, original, working, pkg, part) {
    if (!original) return;
    // SmartArt drawing children are converted into slide shapes. Their ID space
    // belongs to the diagram part (and often repeats 0), not the slide. Only the
    // original PresentationML frame carries the identity referenced by the slide.
    if (![N.p, 'http://purl.oclc.org/ooxml/presentationml/main'].includes(original.namespaceURI)) return;
    const cnv = kid(kids(original).find(e => /^nv/.test(e.localName)), 'cNvPr');
    if (!cnv) return;
    shape.keep = { source: pkg.id, part, element: original.localName, identity: K.fragment(cnv, { pkg, part }) };
    L.properties.shape(shape, original, pkg, part);
    if (/\/slide(?:Masters|Layouts)\//.test(part)) shape.keep.designFrame = K.fragment(working && K.alternate(working) ? working : original, { pkg, part });
    if (find(cnv, 'snd')) shape.keep.sound = true;
    const alternate = K.alternate(working);
    if (shape.type === 'image' && (P.isMedia(original) || alternate && P.isMedia(K.parse(alternate.xml)))) {
      shape.keep.media = { frame: K.fragment(alternate ? working : original, { pkg, part }) };
    }
  };
  P.seal = function (pres) {
    L.properties.seal(pres);
    L.designs.seal(pres);
    for (const d of Object.values(pres.designs)) if (d.keep) {
      d.keep.themeValues = JSON.stringify([d.colors, d.fonts]);
      d.keep.values = P.designValues(d);
    }
    for (const list of L.properties.lists(pres)) L.model.walk(list, shape => {
      if (shape.keep) shape.keep.metadata = Object.fromEntries(['name', 'alt', 'hidden', 'link'].map(k => [k, JSON.stringify(shape[k] ?? null)]));
      if (shape.keep?.media) {
        const keep = shape.keep.media;
        keep.content = value(shape, CONTENT); keep.box = value(shape, GEOMETRY);
        keep.properties = Object.fromEntries(PROPERTIES.map(k => [k, JSON.stringify(shape[k] ?? null)]));
      }
      return true;
    });
  };
  const missing = fragment => fragment.deps.find(d => !d.type || !d.external && !K.package(d.source)?.has(d.part));
  P.valid = shape => !!shape.keep?.media && shape.keep.media.content === value(shape, CONTENT) && !missing(shape.keep.media.frame);
  const shapeRefs = el => all(el).flatMap(e => e.hasAttribute('spid') ? [e.getAttribute('spid')] : []);
  // Targets that exist only inside a graphic frame: a diagram, chart or OLE part, or the frame's own build entry.
  const partRefs = el => all(el).flatMap(e => e.localName === 'spTgt' && kids(e).some(k => ['graphicEl', 'oleChrtEl', 'subSp'].includes(k.localName)) || ['bldGraphic', 'bldDgm', 'bldOleChart'].includes(e.localName) && e.hasAttribute('spid') ? [e.getAttribute('spid')] : []);
  // Prune the timing property, never the slide content DOM. Atomic behaviours and
  // their now-empty containers go with their deleted target; sibling sequences stay.
  // converted: ids of graphic frames now written as plain shapes, whose part-level targets are gone.
  P.pruneTiming = function (xml, allowed, media = allowed, modeled = false, converted = new Set()) {
    const root = K.parse(xml), text = XML.source.get(root).text;
    const atomic = new Set(['video', 'audio', 'cmd', 'set', 'anim', 'animClr', 'animEffect', 'animMotion', 'animRot', 'animScale', 'bldP', 'bldDgm', 'bldOleChart', 'bldGraphic']);
    function visit(el, inMain = false) {
      const pos = XML.source.get(el), tag = el.localName, refs = shapeRefs(el), ctn = kid(el, 'cTn');
      inMain ||= tag === 'cTn' && el.getAttribute('nodeType') === 'mainSeq';
      const cls = ctn?.getAttribute('presetClass');
      if (atomic.has(tag) && (refs.some(id => !allowed.has(id)) || partRefs(el).some(id => converted.has(id)))) return '';
      if (['video', 'audio', 'cmd'].includes(tag) && refs.some(id => !media.has(id))) return '';
      if (tag === 'par' && cls && (modeled && inMain && ['entr', 'exit', 'emph', 'path'].includes(cls) || cls === 'mediacall' && refs.some(id => !media.has(id)))) return '';
      if (tag === 'seq' && ctn?.getAttribute('nodeType') === 'interactiveSeq' && shapeRefs(kid(ctn, 'stCondLst')).some(id => !allowed.has(id))) return '';
      const changes = [], alive = [];
      for (const child of kids(el)) {
        const cp = XML.source.get(child), output = visit(child, inMain);
        if (output) alive.push(child.localName);
        if (output !== text.slice(cp.start, cp.end)) changes.push({ start: cp.start - pos.start, end: cp.end - pos.start, value: output });
      }
      if (['childTnLst', 'subTnLst', 'tnLst', 'bldLst'].includes(tag) && !alive.length) return '';
      if (['par', 'seq'].includes(tag) && !alive.includes('cTn')) return '';
      if (tag === 'cTn' && kid(el, 'childTnLst') && !alive.includes('childTnLst') && !['tmRoot', 'mainSeq'].includes(el.getAttribute('nodeType'))) return '';
      return K.patch(text.slice(pos.start, pos.end), changes);
    }
    let output = visit(root);
    if (modeled && output) {
      const tree = K.parse(output), groups = new Set(all(tree).filter(e => e.localName === 'cTn' && e.hasAttribute('grpId')).map(e => e.getAttribute('grpId'))), edits = [];
      for (const e of all(tree).filter(e => /^bld/.test(e.localName) && e.hasAttribute('grpId') && !groups.has(e.getAttribute('grpId')))) {
        const p = XML.source.get(e); edits.push({ start: p.start, end: p.end, value: '' });
      }
      output = K.patch(output, edits);
    }
    return output;
  };
  function tokenIds(fragment, kind, id) {
    return fragment.ids.flatMap((r, i) => r.kind === kind && String(r.id) === String(id) ? ['\u0001id:' + i + '\u0001'] : []);
  }
  P.slide = function (slide, original, pkg, part, sldId, working) {
    const tree = kid(kid(original, 'cSld'), 'spTree'), cnv = kid(kid(tree, 'nvGrpSpPr'), 'cNvPr');
    slide.keep = { source: pkg.id, part, sldId, rootId: cnv?.getAttribute('id') || '1' };
    L.comments.read(slide, original, pkg, part);
    const media = []; L.model.walk(slide.shapes, s => { if (s.keep?.media) media.push(s); return true; });
    const timing = kid(original, 'timing');
    if (timing) {
      const fragment = K.fragment(timing, { pkg, part });
      slide.keep.timing = fragment; slide.keep.anims = P.anims(slide);
      const raw = K.parse(fragment.xml), root = all(raw).find(e => e.localName === 'cTn' && e.getAttribute('nodeType') === 'tmRoot');
      const main = all(raw).find(e => e.localName === 'cTn' && e.getAttribute('nodeType') === 'mainSeq');
      for (const shape of media) {
        const id = shape.keep.identity.ids.find(i => i.kind === 'shape')?.id;
        const allowed = new Set(tokenIds(fragment, 'shape', id));
        const split = nodes => nodes.flatMap(node => {
          if (!shapeRefs(node).some(ref => allowed.has(ref))) return [];
          const xml = P.pruneTiming(K.raw(node), allowed), f = K.slice(fragment, xml);
          return xml ? [f] : [];
        });
        const mainToken = /^\u0001id:(\d+)\u0001$/.exec(main?.getAttribute('id') || '');
        shape.keep.media.timing = { mainRef: mainToken ? fragment.ids[+mainToken[1]] : null, main: split(kids(kid(main, 'childTnLst'))),
          root: split(kids(kid(root, 'childTnLst')).filter(e => !all(e).includes(main))) };
      }
    }
    const normalized = kid(working, 'transition');
    const transition = normalized && K.alternate(normalized) ? normalized : kid(original, 'transition');
    if (transition) {
      slide.keep.transition = K.fragment(transition, { pkg, part }); slide.keep.trans = JSON.stringify(slide.trans);
    }
    L.properties.slide(slide, original, pkg, part);
  };
  // Change a known property in every alternative while leaving the rest lexical.
  function rewrite(xml, select, change) {
    const root = K.parse(xml), edits = [];
    for (const el of all(root).filter(select)) {
      const pos = XML.source.get(el); edits.push({ start: pos.start, end: pos.end, value: change(el) });
    }
    return K.patch(xml, edits);
  }
  function replaceChildren(xml, replacements, type) { return K.merge(xml, Object.fromEntries(Object.entries(replacements).map(([k, v]) => ['{' + N.a + '}' + k, v])), type); }
  function metadata(xml, shape, baseline, generated) {
    const clone = K.parse(xml), changed = key => baseline[key] !== JSON.stringify(shape[key] ?? null);
    for (const [key, attr] of [['name', 'name'], ['alt', 'descr'], ['hidden', 'hidden']]) if (changed(key)) {
      if (generated.hasAttribute(attr)) clone.setAttribute(attr, generated.getAttribute(attr)); else clone.removeAttribute(attr);
    }
    if (changed('link')) {
      for (const c of kids(clone).filter(e => e.localName === 'hlinkClick')) c.remove();
      const link = kid(generated, 'hlinkClick'); if (link) { const c = link.cloneNode(true); c.parentNode = clone; clone.children.unshift(c); clone.childNodes.unshift(c); }
    }
    return K.serialize(clone);
  }
  P.nonVisual = function (shape, ctx, generated) {
    if (!ctx.writer || !shape.keep?.identity) return generated;
    const absent = missing(shape.keep.identity);
    if (absent) {
      ctx.writer.loss({ id: 'action-sound:' + shape.id, what: 'The action sound refers to a part missing from the original file.', where: shape.name + ': ' + (absent.part || absent.id), action: 'drop' });
      return generated;
    }
    const identity = L.properties.action(shape.keep.identity, shape.link, ctx, generated);
    let xml = L.properties.emit(identity, ctx);
    if (xml == null) return generated;
    if (['name', 'alt', 'hidden', 'link'].some(k => shape.keep.metadata[k] !== JSON.stringify(shape[k] ?? null))) {
      const g = K.parse('<root xmlns:p="' + N.p + '" xmlns:a="' + N.a + '" xmlns:r="' + N.rel + '">' + generated + '</root>');
      xml = metadata(xml, shape, shape.keep.metadata, find(g, 'cNvPr'));
    }
    return L.comments.shapeIdentity(shape, ctx, xml);
  };
  P.emitShape = function (shape, ctx, generated) {
    const keep = shape.keep?.media;
    if (!keep) return null;
    if (!P.valid(shape)) {
      ctx.writer.loss({ id: 'media:' + shape.id, ...(missing(keep.frame) ? { what: 'The media object refers to data missing from the original file.', notify: false } : { what: 'This video or sound will be saved as a picture and will no longer play, because its picture was replaced.', notify: true, place: ctx.place }), where: shape.name, action: 'conversion' });
      return null;
    }
    let xml = ctx.writer.emit(keep.frame, ctx.part);
    if (!xml) return '';
    xml = L.comments.shapeIdentity(shape, ctx, xml);
    if (ctx.inGroup || keep.box !== value(shape, GEOMETRY)) xml = K.setBox(xml, Object.fromEntries(GEOMETRY.map(k => [k, shape[k] || 0])));
    const changed = key => keep.properties[key] !== JSON.stringify(shape[key] ?? null);
    if (!PROPERTIES.some(changed)) return xml;
    const g = K.parse('<root xmlns:p="' + N.p + '" xmlns:a="' + N.a + '" xmlns:r="' + N.rel + '">' + generated() + '</root>');
    const gsp = find(g, 'spPr'), gpic = find(g, 'blipFill'), gn = find(g, 'cNvPr');
    const raw = e => e ? K.raw(e) : [];
    if (changed('line') || changed('geom') || changed('adj') || changed('path') || changed('shadow')) {
      xml = rewrite(xml, e => e.localName === 'spPr', e => {
        const changes = {};
        if (changed('line')) changes.ln = raw(kid(gsp, 'ln'));
        if (changed('geom') || changed('adj') || changed('path')) { changes.prstGeom = raw(kid(gsp, 'prstGeom')); changes.custGeom = raw(kid(gsp, 'custGeom')); }
        if (changed('shadow')) {
          const fx = kid(e, 'effectLst');
          changes.effectLst = replaceChildren(fx ? K.raw(fx) : '<a:effectLst xmlns:a="' + N.a + '"/>', { outerShdw: raw(kid(kid(gsp, 'effectLst'), 'outerShdw')) }, 'a:CT_EffectList');
        }
        return replaceChildren(K.raw(e), changes, 'a:CT_ShapeProperties');
      });
    }
    if (changed('crop')) xml = rewrite(xml, e => e.localName === 'blipFill', e => {
      const old = kid(e, 'srcRect'), replacement = raw(kid(gpic, 'srcRect'));
      const pos = XML.source.get(e), cp = old && XML.source.get(old), after = kid(e, 'blip') && XML.source.get(kid(e, 'blip')).end;
      return K.patch(pos.text.slice(pos.start, pos.end), [{ start: (cp ? cp.start : after || pos.openEnd) - pos.start, end: (cp ? cp.end : after || pos.openEnd) - pos.start, value: Array.isArray(replacement) ? '' : replacement }]);
    });
    if (changed('img')) xml = rewrite(xml, e => e.localName === 'blip', e => {
      const changes = {}; for (const name of ['alphaModFix', 'clrChange', 'grayscl', 'biLevel', 'lum']) changes[name] = raw(kid(find(gpic, 'blip'), name));
      return replaceChildren(K.raw(e), changes, 'a:CT_Blip');
    });
    if (['name', 'alt', 'hidden', 'link'].some(changed)) xml = rewrite(xml, e => e.localName === 'cNvPr', e => metadata(K.raw(e), shape, keep.properties, gn));
    return xml;
  };
  function append(xml, element, content) {
    if (!content) return xml;
    const pos = XML.source.get(element), self = pos.text[pos.openEnd - 2] === '/';
    return K.patch(xml, [self ? { start: pos.openEnd - 2, end: pos.openEnd, value: '>' + content + '</' + element.nodeName + '>' } : { start: pos.end - element.nodeName.length - 3, end: pos.end - element.nodeName.length - 3, value: content }]);
  }
  P.timing = function (slide, ctx, generate) {
    const allowed = new Set(), media = new Set(), extras = [];
    L.model.walk(slide.shapes, shape => {
      const id = String(ctx.idOf(shape.id)); allowed.add(id);
      if (P.valid(shape)) {
        media.add(id); const f = shape.keep.media.frame, base = slide.keep?.timing;
        if (!base || f.source !== base.source || f.part !== base.part || f.copy !== base.copy) extras.push(shape.keep.media.timing);
      }
      return true;
    });
    const kept = slide.keep?.timing;
    if (!kept && !extras.some(Boolean)) return generate();
    let xml = kept ? ctx.writer.emit(kept, ctx.part) : '';
    const live = new Set(); L.model.walk(slide.shapes, shape => { live.add(shape.id); return true; });
    const previous = slide.keep?.anims ? JSON.parse(slide.keep.anims).filter(a => live.has(a.sid)) : [];
    const edited = !!kept && JSON.stringify(previous) !== P.anims(slide);
    const converted = ctx.convertedFrames || new Map();
    if (xml) {
      const targets = new Set(partRefs(K.parse(xml)));
      for (const [id, record] of converted) if (targets.has(id)) ctx.writer.loss({ id: 'frame-timing:' + record.key, what: record.label + ': its animation by parts was removed with the original object.', where: ctx.part, action: 'drop' });
      xml = P.pruneTiming(xml, allowed, media, edited, converted);
    }
    if (!xml) {
      const rootId = ctx.writer.ids.fresh(ctx.part, 'timing'), mainId = ctx.writer.ids.fresh(ctx.part, 'timing');
      xml = '<p:timing xmlns:p="' + N.p + '"><p:tnLst><p:par><p:cTn id="' + rootId + '" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="' + mainId + '" dur="indefinite" nodeType="mainSeq"><p:childTnLst/></p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst></p:timing>';
    }
    let tree = K.parse(xml);
    const root = all(tree).find(e => e.localName === 'cTn' && e.getAttribute('nodeType') === 'tmRoot');
    const main = all(tree).find(e => e.localName === 'cTn' && e.getAttribute('nodeType') === 'mainSeq');
    ctx.timingRootId = root?.getAttribute('id'); ctx.timingMainId = main?.getAttribute('id');
    ctx.nextTimingId = () => ctx.writer.ids.fresh(ctx.part, 'timing');
    ctx.nextBuildId = () => ctx.writer.ids.fresh(ctx.part, 'build');
    const add = { main: '', root: '' };
    for (const extra of extras.filter(Boolean)) for (const where of ['main', 'root']) for (const f of extra[where]) {
      const references = f.ids.filter(i => !i.definition && i.kind === 'timing' && extra.mainRef && i.source === extra.mainRef.source && i.scope === extra.mainRef.scope && i.id === extra.mainRef.id);
      for (const ref of references) if (main) ctx.writer.ids.bind(ref.source, ref.scope, ref.kind, ref.id, main.getAttribute('id'), ctx.part);
      try { add[where] += ctx.writer.emit(f, ctx.part, { references }); }
      catch (error) {
        // The core records the uncopyable reference. Keep the media and all
        // independent behaviours; that one sequence uses the converted fallback.
        if (error.code !== 'OOXML_COPY_REFERENCE') throw error;
      }
    }
    if (edited || !kept) {
      const fresh = generate();
      if (fresh) {
        const g = K.parse('<root xmlns:p="' + N.p + '">' + fresh + '</root>'), seq = all(g).find(e => e.localName === 'cTn' && e.getAttribute('nodeType') === 'mainSeq');
        add.main += kids(kid(seq, 'childTnLst')).map(K.raw).join('');
        const builds = find(g, 'bldLst');
        if (builds) {
          const old = kid(K.parse(xml), 'bldLst');
          if (old) {
            const p = XML.source.get(old), combined = append(K.raw(old), K.parse(K.raw(old)), kids(builds).map(K.raw).join(''));
            xml = K.patch(xml, [{ start: p.start, end: p.end, value: combined }]);
          } else xml = append(xml, K.parse(xml), K.raw(builds));
        }
      }
    }
    for (const where of ['main', 'root']) if (add[where]) {
      tree = K.parse(xml);
      const node = all(tree).find(e => e.localName === 'cTn' && e.getAttribute('nodeType') === (where === 'main' ? 'mainSeq' : 'tmRoot'));
      if (!node) throw new Error('Cannot attach preserved media timing: missing ' + where + ' sequence');
      const list = kid(node, 'childTnLst');
      xml = append(xml, list || node, list ? add[where] : '<p:childTnLst>' + add[where] + '</p:childTnLst>');
    }
    xml = P.pruneTiming(xml, allowed, media, false, converted);
    if (!media.size && !all(K.parse(xml)).some(e => ['audio', 'video', 'cmd', 'set', 'anim', 'animEffect', 'animClr', 'animMotion', 'animRot', 'animScale'].includes(e.localName))) return '';
    return xml;
  };
  P.clipboard = function (item) {
    const refs = [...L.comments.references(item), ...L.frames.references(item)], visit = value => {
      if (!value || typeof value !== 'object') return;
      if (typeof value.xml === 'string' && Array.isArray(value.deps)) {
        for (const dep of value.deps) if (!dep.external && dep.part) refs.push({ source: dep.source, part: dep.part });
        return;
      }
      Object.values(value).forEach(visit);
    };
    visit(item);
    const sources = K.export(refs);
    for (const source of sources) for (const part of source.parts) if (!part.type) part.type = P.mediaType(null, part.name) || 'application/octet-stream';
    return { version: 1, item, sources };
  };
  P.pasteboard = function (packet) {
    if (packet.version !== 1 || !packet.item || !Array.isArray(packet.sources)) throw new Error('Unsupported Lectern clipboard');
    const sources = K.import(packet.sources), item = K.remapSources(packet.item, sources), media = new Map();
    const restore = shape => {
      if (!P.valid(shape)) return true;
      const frame = shape.keep.media.frame, root = K.parse(frame.xml), blip = find(root, 'blip');
      const embed = blip?.getAttributeNS(N.rel, 'embed') || blip?.getAttribute('r:embed'), match = /^\u0001rel:(\d+)\u0001$/.exec(embed || '');
      const dep = match && frame.deps[+match[1]], pkg = dep && K.package(dep.source);
      if (pkg && dep.part) {
        const key = dep.source + ':' + dep.part;
        if (!media.has(key)) media.set(key, L.media.add(new Blob([pkg.bytes(dep.part)], { type: pkg.type(dep.part) || 'image/png' }), dep.part.split('/').pop()));
        shape.media = media.get(key);
        if (shape.img?.view) { delete shape.img.view; shape.keep.media.properties.img = JSON.stringify(shape.img); }
        shape.keep.media.content = value(shape, CONTENT);
      }
      return true;
    };
    L.model.walk(item.kind === 'slides' ? item.slides.flatMap(s => s.shapes) : item.shapes || [], restore);
    L.frames.restore(item);
    return item;
  };
})();
