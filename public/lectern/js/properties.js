/* Lectern — unsupported details stay with the property or object that owns them. */
(function () {
  'use strict';
  const L = window.L, K = L.opc, X = L.xmlTree, Q = (L.properties = {}), N = K.NS;
  const kids = e => Array.from(e?.children || []), kid = (e, name) => kids(e).find(c => c.localName === name);
  const all = e => e ? [e, ...kids(e).flatMap(all)] : [];
  const find = (e, name) => all(e).find(c => c.localName === name);
  const raw = e => e ? K.raw(e) : '';
  const plain = o => JSON.parse(JSON.stringify(o ?? null, (k, v) => k === 'keep' ? undefined : v));
  const same = (a, b) => JSON.stringify(plain(a)) === JSON.stringify(plain(b));
  const replace = (xml, el, next) => { const p = X.source.get(el); return K.patch(xml, [{ start: p.start, end: p.end, value: next }]); };
  const fillNames = 'noFill solidFill gradFill blipFill pattFill grpFill';
  const GEOMETRY = ['x', 'y', 'w', 'h', 'rot', 'flipH', 'flipV'];
  const shapeFields = [...GEOMETRY, 'type', 'fill', 'line', 'shadow', 'geom', 'adj', 'path', 'crop', 'img', 'media', 'svgMedia', 'linkUrl', 'lockAspect', 'ph', 'wa'];
  function snapshot(o, fields) { return Object.fromEntries(fields.map(k => [k, plain(o[k])])); }
  function capture(el, ctx) {
    const ac = K.alternate(el);
    if (ac) {
      const root = K.parse(ac.xml), branches = kids(root).flatMap(kids);
      if (!branches.every(e => e.localName === el.localName)) {
        const candidates = all(root).filter(e => e.localName === el.localName);
        el = candidates.at(-1) || K.parse(K.raw(el));
      }
    }
    const f = K.fragment(el, { pkg: ctx.pkg, part: ctx.partPath });
    return K.slice(f, f.xml);
  }
  Q.capture = function (object, el, ctx, kind) {
    if (!el) return;
    object.keep = { ...(object.keep || {}), kind, record: L.uid('property'), fragment: capture(el, ctx), before: plain(object) };
    if (kind === 'para') object.keep.before.lvl = +(el.getAttribute('lvl') || 0);
  };
  Q.shape = function (shape, original, pkg, part) {
    const keep = shape.keep.properties = {}, ctx = { pkg, partPath: part };
    for (const tag of ['spPr', 'grpSpPr', 'blipFill', 'style']) {
      const el = kid(original, tag); if (el) keep[tag] = K.fragment(el, { pkg, part });
    }
    const nv = kids(original).find(e => /^nv/.test(e.localName));
    keep.nonVisual = kids(nv).filter(e => e.localName !== 'cNvPr').map(el => K.fragment(el, { pkg, part }));
    keep.attributes = Array.from(original.attributes).filter(a => a.prefix !== 'xmlns' && a.name !== 'xmlns').map(a => ({ name: a.name, value: a.value }));
    if (shape.wa && kid(original, 'txBody')) keep.wordArt = K.fragment(kid(original, 'txBody'), { pkg, part });
    // Lines can have an empty, formatted text body even though the editor does
    // not expose their text. It still belongs to the source shape.
    if (!shape.tx && kid(original, 'txBody')) keep.textBody = K.fragment(kid(original, 'txBody'), { pkg, part });
    if (shape.tx) {
      const body = kid(original, 'txBody');
      Q.capture(shape.tx, kid(body, 'bodyPr'), ctx, 'body');
      const list = kid(body, 'lstStyle');
      if (list) { shape.tx.keep ||= {}; shape.tx.keep.list = K.fragment(list, { pkg, part }); }
    }
  };
  Q.lists = pres => pres.slides.map(s => s.shapes).concat(Object.values(pres.designs).flatMap(d => [d.deco, d.titleDeco, ...Object.values(d.layoutDecos || {})]).filter(Boolean));
  Q.seal = function (pres) {
    for (const list of Q.lists(pres)) L.model.walk(list, shape => {
      if (shape.keep?.properties) shape.keep.properties.before = snapshot(shape, shapeFields);
      return true;
    });
    const seen = new Set(), visit = o => {
      if (!o || typeof o !== 'object' || seen.has(o)) return;
      seen.add(o);
      const keep = o.keep;
      if (keep?.kind) {
        const before = plain(o); delete before.t; delete before.ps;
        if (keep.kind === 'para') before.lvl = keep.before.lvl;
        keep.before = before;
        if (keep.list) keep.listBefore = plain(o.lst);
      }
      for (const [k, v] of Object.entries(o)) if (k !== 'keep') visit(v);
    };
    visit(pres);
  };
  Q.emit = function (fragment, ctx) {
    if (!fragment) return '';
    let writer = ctx.writer;
    // Each copied tag list belongs to its copied slide/shape. Other immutable
    // dependencies may be shared; future tag edits in Office must be independent.
    if (fragment.copy && fragment.deps.some(d => /\/tags$/.test(d.type || ''))) {
      const local = Object.create(writer), base = writer;
      local.keepRel = (owner, dep) => {
        if (!/\/tags$/.test(dep.type || '')) return base.keepRel(owner, dep);
        const map = base.tagCopies || (base.tagCopies = new Map()), key = [dep.source, dep.part, fragment.copy].join('|');
        if (!map.has(key)) {
          const pkg = K.package(dep.source), target = base.name('ppt/tags', 'tag', 'xml');
          base.copyPart(pkg, dep.part, target); map.set(key, target);
        }
        return base.rels(owner).add(dep.type, K.relative(owner, map.get(key)));
      };
      writer = local;
    }
    try { return writer.emit(fragment, ctx.part); }
    catch (error) {
      writer.loss({ id: 'property:' + fragment.source + ':' + fragment.part + ':' + fragment.xml.slice(0, 80), what: 'A retained property could not be written: ' + error.message, where: ctx.part, action: 'drop' });
      return null;
    }
  };
  Q.action = function (fragment, link, ctx, generated) {
    // A known slide jump follows the live model, including a deleted target.
    // Remove its old dependency before emission so a link cannot revive a slide.
    if (!link?.slide) return fragment;
    const root = K.parse(fragment.xml), fresh = K.parse('<root xmlns:a="' + N.a + '" xmlns:p="' + N.p + '" xmlns:r="' + N.rel + '">' + generated + '</root>');
    let xml = fragment.xml;
    const old = find(root, 'hlinkClick'), next = find(fresh, 'hlinkClick');
    if (old) {
      const changed = next ? K.attributes(raw(old), { 'r:id': next.getAttributeNS(N.rel, 'id') || next.getAttribute('r:id') }) : '';
      xml = replace(xml, old, changed);
      if (!next) ctx.writer.loss({ id: 'deleted-link:' + fragment.source + ':' + fragment.part + ':' + link.slide, what: 'A link to a deleted slide was removed.', where: ctx.part, action: 'drop' });
    }
    return K.slice(fragment, xml);
  };
  function shadow(xml, generated) {
    const old = kid(K.parse(xml), 'effectLst'), next = kid(K.parse(generated), 'effectLst');
    const effects = K.merge(raw(old) || '<a:effectLst xmlns:a="' + N.a + '"/>', { ['{' + N.a + '}outerShdw']: raw(kid(next, 'outerShdw')) }, 'a:CT_EffectList');
    return effects;
  }
  Q.text = function (object, ctx, generated, kind, tag, lvl) {
    const keep = object?.keep;
    if (!ctx?.writer || !keep?.fragment || keep.kind !== kind) return generated;
    let f = Q.action(keep.fragment, object.link, ctx, generated);
    let xml = Q.emit(f, ctx); if (xml == null) return generated;
    const fields = kind === 'run' ? { ...L.designs.runFields, font: 'latin ea cs', shd: '', shdX: '', link: 'hlinkClick' } : kind === 'para' ? { ...L.designs.paraFields, lvl: '@lvl' } : {
      anchor: '@anchor', anchorCtr: '@anchorCtr', wrap: '@wrap', ins: '@lIns @tIns @rIns @bIns', vert: '@vert', rot: '@rot', cols: '@numCol', colGap: '@spcCol', autofit: 'noAutofit normAutofit spAutoFit', fontScale: 'normAutofit', lnSpcRed: 'normAutofit', warp: 'prstTxWarp',
    };
    const current = kind === 'para' ? { ...object, lvl } : object;
    const fresh = raw(kids(K.parse('<root xmlns:a="' + N.a + '" xmlns:r="' + N.rel + '">' + (generated || '<' + tag + '/>') + '</root>'))[0]);
    const apply = el => {
      let out = L.designs.property(raw(el), fresh, current, keep.before, fields, kind === 'run' ? 'a:CT_TextCharacterProperties' : kind === 'para' ? 'a:CT_TextParagraphProperties' : 'a:CT_TextBodyProperties');
      if (kind === 'para' && !same(object.rPr, keep.before.rPr)) out = K.merge(out, { ['{' + N.a + '}defRPr']: raw(kid(K.parse(fresh), 'defRPr')) }, 'a:CT_TextParagraphProperties');
      if (kind === 'run' && (!same(object.shd, keep.before.shd) || !same(object.shdX, keep.before.shdX))) out = K.merge(out, { ['{' + N.a + '}effectLst']: shadow(out, fresh) }, 'a:CT_TextCharacterProperties');
      if (el.localName !== tag.split(':').pop()) {
        const tree = K.parse(out), p = X.source.get(tree);
        out = K.patch(out, [{ start: p.start + 1, end: p.start + 1 + tree.nodeName.length, value: tag }, ...(p.text[p.openEnd - 2] === '/' ? [] : [{ start: p.end - tree.nodeName.length - 1, end: p.end - 1, value: tag }])]);
      }
      return out;
    };
    const root = K.parse(xml);
    if (root.localName !== 'AlternateContent') return apply(root);
    const edits = all(root).filter(e => ['rPr', 'endParaRPr', 'defRPr', 'pPr', 'bodyPr'].includes(e.localName)).map(e => { const p = X.source.get(e); return { start: p.start, end: p.end, value: apply(e) }; });
    return K.patch(xml, edits);
  };
  Q.list = (tx, ctx, generated) => ctx?.writer && tx.keep?.list && same(tx.lst, tx.keep.listBefore) ? Q.emit(tx.keep.list, ctx) ?? generated : generated;
  Q.apply = function (shape, ctx, generated) {
    const keep = shape.keep?.properties; if (!keep?.before || !ctx.writer) return generated;
    const before = keep.before, changed = k => !same(shape[k], before[k]);
    let xml = generated, root = K.parse('<root xmlns:a="' + N.a + '" xmlns:p="' + N.p + '" xmlns:r="' + N.rel + '">' + generated + '</root>');
    const replaceIn = (tag, text, parentName) => {
      if (text == null) return;
      const tree = K.parse('<root xmlns:a="' + N.a + '" xmlns:p="' + N.p + '" xmlns:r="' + N.rel + '">' + xml + '</root>'), shape = kids(tree)[0], parent = parentName ? kids(shape).find(e => e.localName === parentName) : shape;
      if (!parent) return;
      const target = kid(parent, tag);
      const type = { sp: 'CT_Shape', pic: 'CT_Picture', grpSp: 'CT_GroupShape', graphicFrame: 'CT_GraphicalObjectFrame', cxnSp: 'CT_Connector' }[parent.localName];
      const changed = target ? replace(raw(parent), kid(K.parse(raw(parent)), tag), text) : type ? K.merge(raw(parent), { ['{' + N.p + '}' + tag]: text }, 'p:' + type) : K.mergeBag(raw(parent), { __append: [text] });
      if (parent === shape) xml = changed;
      else xml = replace(raw(shape), kid(K.parse(raw(shape)), parentName), changed);
    };
    for (const tag of ['spPr', 'grpSpPr']) if (keep[tag]) {
      let original = Q.emit(keep[tag], ctx); if (original == null) continue;
      const fresh = find(root, tag), replacements = {};
      const fromFresh = name => { replacements['{' + N.a + '}' + name] = raw(kid(fresh, name)); };
      // Regenerated group members use the model's common coordinate space.
      // Their parent must use that same space even when its visible box is unchanged.
      if (tag === 'grpSpPr' || ctx.inGroup || GEOMETRY.some(changed)) fromFresh('xfrm');
      if (changed('fill')) fillNames.split(' ').forEach(fromFresh);
      if (changed('line')) fromFresh('ln');
      if (['geom', 'adj', 'path'].some(changed)) ['prstGeom', 'custGeom'].forEach(fromFresh);
      if (changed('shadow')) replacements['{' + N.a + '}effectLst'] = shadow(original, raw(fresh));
      original = K.merge(original, replacements, tag === 'grpSpPr' ? 'a:CT_GroupShapeProperties' : 'a:CT_ShapeProperties');
      replaceIn(tag, original);
    }
    if (keep.blipFill) {
      let original = Q.emit(keep.blipFill, ctx);
      if (original != null) {
        const fresh = find(root, 'blipFill');
        if (changed('crop')) original = K.merge(original, { ['{' + N.a + '}srcRect']: raw(kid(fresh, 'srcRect')) }, 'a:CT_BlipFillProperties');
        let oldBlip = kid(K.parse(original), 'blip'); const newBlip = kid(fresh, 'blip');
        if (oldBlip) {
          let b = raw(oldBlip), replacements = {};
          if (['media', 'svgMedia', 'linkUrl'].some(changed)) {
            b = K.attributes(b, { 'r:embed': newBlip?.getAttributeNS(N.rel, 'embed'), 'r:link': newBlip?.getAttributeNS(N.rel, 'link') });
            replacements['{' + N.a + '}extLst'] = raw(kid(newBlip, 'extLst'));
          }
          const properties = { alpha: ['alphaModFix'], clear: ['clrChange'], mode: ['grayscl', 'biLevel', 'lum'], bright: ['lum'], contrast: ['lum'] };
          for (const [key, names] of Object.entries(properties)) if (!same(shape.img?.[key], before.img?.[key])) for (const name of names) replacements['{' + N.a + '}' + name] = raw(kid(newBlip, name));
          b = K.merge(b, replacements, 'a:CT_Blip'); original = replace(original, oldBlip, b);
        }
        replaceIn('blipFill', original);
      }
    }
    if (keep.style) replaceIn('style', Q.emit(keep.style, ctx));
    if (keep.wordArt && !changed('wa')) replaceIn('txBody', Q.emit(keep.wordArt, ctx));
    if (keep.textBody && !shape.tx) replaceIn('txBody', Q.emit(keep.textBody, ctx));
    const nv = kids(kids(root)[0]).find(e => /^nv/.test(e.localName));
    for (const f of keep.nonVisual || []) {
      const sourceTag = K.parse(f.xml).localName;
      if (sourceTag.startsWith('cNv') && nv && !kid(nv, sourceTag)) {
        ctx.writer.loss({ id: 'shape-form:' + shape.id, what: 'This object uses its converted shape type; incompatible shape locks were removed.', where: ctx.part, action: 'conversion' });
        continue;
      }
      let text = Q.emit(f, ctx); if (text == null) continue;
      if (ctx.liveShapeIds) {
        const edits = [];
        for (const e of all(K.parse(text)).filter(e => ['stCxn', 'endCxn'].includes(e.localName))) if (!ctx.liveShapeIds.has(e.getAttribute('id'))) {
          const p = X.source.get(e); edits.push({ start: p.start, end: p.end, value: '' });
          ctx.writer.loss({ id: 'connector:' + shape.id + ':' + e.localName, what: 'A connector endpoint attached to a deleted shape was detached.', where: ctx.part, action: 'conversion' });
        }
        text = K.patch(text, edits);
      }
      const tag = K.parse(text).localName;
      if (tag === 'nvPr' && changed('ph')) text = K.mergeBag(text, { ['{' + N.p + '}ph']: raw(find(nv, 'ph')) });
      if (changed('lockAspect')) {
        const lock = all(K.parse(text)).find(e => /Locks$/.test(e.localName));
        if (lock) text = replace(text, lock, K.attributes(raw(lock), { noChangeAspect: shape.lockAspect === false ? '0' : '1' }));
      }
      replaceIn(tag, text, nv?.localName);
    }
    const attrs = Object.fromEntries((keep.attributes || []).filter(a => a.name !== 'useBgFill').map(a => [a.name, a.value]));
    return K.attributes(xml, attrs);
  };
  Q.slide = function (slide, root, pkg, part) {
    const common = kid(root, 'cSld'), capture = e => K.fragment(e, { pkg, part });
    slide.keep.details = {
      common: kids(common).filter(e => ['custDataLst', 'extLst'].includes(e.localName)).map(capture),
      root: kids(root).filter(e => ['clrMapOvr', 'extLst'].includes(e.localName)).map(capture),
      bg: kid(common, 'bg') && capture(kid(common, 'bg')), background: plain(slide.bg), design: slide.design,
      attributes: Object.fromEntries(Array.from(root.attributes).filter(a => !['show', 'showMasterSp'].includes(a.localName)).map(a => [a.name, a.value])),
      commonAttributes: Object.fromEntries(Array.from(common.attributes).map(a => [a.name, a.value])),
    };
  };
  function extensions(original, generated) {
    const old = K.parse(original), next = generated && K.parse(generated), replacements = {};
    for (const ext of kids(next)) replacements[ext.getAttribute('uri')] = raw(ext);
    return K.mergeBag(original, replacements, e => e.getAttribute('uri'));
  }
  Q.finishSlide = function (slide, ctx, xml) {
    const keep = slide.keep?.details; if (!keep) return xml;
    for (const where of ['common', 'root']) for (const f of keep[where]) {
      const tag = K.parse(f.xml).localName, tree = K.parse(xml), parent = where === 'common' ? kid(tree, 'cSld') : tree;
      if (tag === 'clrMapOvr' && slide.design !== keep.design) continue;
      // Replaced extensions must not first carry their old dependencies. A copied
      // comment extension otherwise adds a second relationship to the source thread.
      const generated = tag === 'extLst' && kid(parent, tag);
      // Leave an empty slot at each original URI so replacement keeps its order.
      const retained = generated ? K.slice(f, K.mergeBag(f.xml, Object.fromEntries(kids(generated).map(e => [e.getAttribute('uri'),
        '<p:ext xmlns:p="' + N.p + '" uri="' + K.esc(e.getAttribute('uri')) + '"/>'])), e => e.getAttribute('uri'))) : f;
      let text = Q.emit(retained, ctx); if (text == null) continue;
      if (tag === 'extLst') text = extensions(text, raw(kid(parent, tag)));
      const next = K.merge(raw(parent), { ['{' + N.p + '}' + tag]: text }, where === 'common' ? 'p:CT_CommonSlideData' : 'p:CT_Slide');
      xml = where === 'common' ? replace(xml, parent, next) : next;
    }
    if (keep.bg && same(slide.bg, keep.background)) {
      const text = Q.emit(keep.bg, ctx), parent = kid(K.parse(xml), 'cSld');
      if (text != null) xml = replace(xml, parent, K.merge(raw(parent), { ['{' + N.p + '}bg']: text }, 'p:CT_CommonSlideData'));
    }
    const common = kid(K.parse(xml), 'cSld');
    xml = replace(xml, common, K.attributes(raw(common), keep.commonAttributes || {}));
    return K.attributes(xml, keep.attributes || {});
  };
  Q.transition = function (slide, ctx, generated) {
    const keep = slide.keep; if (!keep?.transition) return generated;
    const before = JSON.parse(keep.trans), current = slide.trans;
    if (same(current, before)) return Q.emit(keep.transition, ctx) ?? generated;
    const effect = o => Object.fromEntries(Object.entries(o || {}).filter(([k]) => !['spd', 'click', 'after'].includes(k)));
    if (same(effect(current), effect(before))) {
      let xml = Q.emit(keep.transition, ctx); if (xml == null) return generated;
      const changes = {};
      if (current.spd !== before.spd) changes.spd = current.spd || 'fast';
      if (current.click !== before.click) changes.advClick = current.click === false ? '0' : '1';
      if (current.after !== before.after) changes.advTm = current.after == null || current.after === '' ? null : Math.round(current.after);
      const edits = all(K.parse(xml)).filter(e => e.localName === 'transition').map(e => { const p = X.source.get(e); return { start: p.start, end: p.end, value: K.attributes(raw(e), changes) }; });
      return K.patch(xml, edits);
    }
    ctx.writer.loss({ id: 'transition:' + slide.id, what: 'The original transition was replaced with the selected transition.', where: ctx.part, action: 'conversion' });
    return generated;
  };
})();
