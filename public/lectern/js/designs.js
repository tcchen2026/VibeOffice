/* Lectern — imported design libraries and property-scoped design changes. */
(function () {
  'use strict';
  const L = window.L, K = L.opc, X = L.xmlTree, D = (L.designs = {}), N = K.NS;
  const kids = e => Array.from(e?.children || []), kid = (e, name) => kids(e).find(c => c.localName === name);
  const all = e => e ? [e, ...kids(e).flatMap(all)] : [];
  const find = (e, name) => all(e).find(c => c.localName === name);
  const raw = e => e ? K.raw(e) : '';
  const value = o => JSON.stringify(o ?? null, (k, v) => k === 'keep' ? undefined : v);
  const same = (a, b) => value(a) === value(b);
  const geometry = new Set(['x', 'y', 'w', 'h', 'rot', 'flipH', 'flipV']);
  const shapeValue = s => value(Object.fromEntries(Object.entries(s).filter(([k]) => !geometry.has(k))));
  const namespaced = e => '{' + e.namespaceURI + '}' + e.localName;
  const changeChild = (xml, name, child, type) => K.merge(xml, { ['{' + N.p + '}' + name]: child }, type);
  const changeAt = (xml, node, next) => { const p = X.source.get(node); return K.patch(xml, [{ start: p.start, end: p.end, value: next }]); };
  D.seal = function (pres) {
    for (const d of Object.values(pres.designs)) if (d.keep) {
      d.keep.design = d.id;
      d.keep.before = JSON.parse(value(d));
      d.keep.decorations = {};
      for (const record of [{ part: d.keep.part, fragment: d.keep.master }, ...(d.keep.layoutParts || [])]) {
        if (!record.fragment) continue;
        const shapes = (record.lkey ? d.layoutDecos?.[record.lkey] : d.deco) || [];
        const identities = shapes.map(s => ({ id: s.id, ids: (s.keep?.identity || s.keep?.designFrame || s.keep?.frame?.fragment)?.ids.filter(i => i.definition && i.kind === 'shape').map(i => i.id) || [] }));
        const tree = find(K.parse(record.fragment.xml), 'spTree');
        d.keep.decorations[record.part] = kids(tree).flatMap((node, index) => {
          const ids = all(node).filter(e => e.localName === 'cNvPr').map(e => {
            const token = /^\u0001id:(\d+)\u0001$/.exec(e.getAttribute('id') || '');
            return token ? record.fragment.ids[+token[1]].id : e.getAttribute('id');
          });
          const models = identities.filter(s => s.ids.some(id => ids.includes(id))).map(s => s.id);
          return models.length ? [{ index, models }] : [];
        });
      }
      for (const list of [d.deco, d.titleDeco, ...Object.values(d.layoutDecos || {})]) L.model.walk(list || [], s => {
        if (s.keep?.designFrame) s.keep.designValue = shapeValue(s);
        return true;
      });
    }
  };
  D.shape = function (shape, ctx) {
    const f = shape.keep?.designFrame;
    if (!f || shape.keep.designValue !== shapeValue(shape)) return null;
    return ctx.writer.emit(f, ctx.part, { box: Object.fromEntries([...geometry].map(k => [k, shape[k] || 0])) });
  };
  // Apply only the fields edited in the model, retaining unknown siblings and
  // attributes. Used for design text styles as well as individual font changes.
  D.property = function (xml, generated, current, before, mapping, type) {
    const fresh = K.parse(generated), changes = {}, attrs = {};
    for (const [field, names] of Object.entries(mapping)) if (!same(current?.[field], before?.[field])) {
      for (const name of names.split(' ')) {
        if (!name) continue;
        if (name.startsWith('@')) attrs[name.slice(1)] = fresh.getAttribute(name.slice(1));
        else if (name === 'effectLst' && type === 'a:CT_TextCharacterProperties') {
          const old = kid(K.parse(xml), name), next = kid(fresh, name);
          changes['{' + N.a + '}' + name] = K.merge(raw(old) || '<a:effectLst xmlns:a="' + N.a + '"/>', { ['{' + N.a + '}outerShdw']: raw(kid(next, 'outerShdw')) }, 'a:CT_EffectList');
        } else changes['{' + N.a + '}' + name] = kids(fresh).filter(e => e.localName === name).map(raw);
      }
    }
    return K.attributes(K.merge(xml, changes, type), attrs);
  };
  D.runFields = { b: '@b', i: '@i', u: '@u', strike: '@strike', sz: '@sz', spc: '@spc', kern: '@kern', base: '@baseline', lang: '@lang', cap: '@cap', font: 'latin', hl: 'highlight', color: 'solidFill noFill gradFill blipFill pattFill grpFill', fill: 'solidFill noFill gradFill blipFill pattFill grpFill', ln: 'ln', shd: 'effectLst', shdX: 'effectLst' };
  D.paraFields = { algn: '@algn', marL: '@marL', marR: '@marR', indent: '@indent', rtl: '@rtl', defTab: '@defTabSz', spcBef: 'spcBef', spcAft: 'spcAft', lnSpc: 'lnSpc', tabs: 'tabLst', bu: 'buClrTx buClr buSzTx buSzPct buSzPts buFontTx buFont buNone buAutoNum buChar buBlip' };
  function levels(xml, generated, current, before) {
    const fresh = K.parse(generated), replacements = {};
    for (let i = 0; i < 9; i++) {
      if (same(current?.[i], before?.[i])) continue;
      const name = 'lvl' + (i + 1) + 'pPr', old = kid(K.parse(xml), name), next = kid(fresh, name);
      if (!next) continue;
      let out = old ? D.property(raw(old), raw(next), current?.[i], before?.[i], D.paraFields, 'a:CT_TextParagraphProperties') : raw(next);
      if (old && !same(current?.[i]?.rPr, before?.[i]?.rPr)) {
        const r = kid(old, 'defRPr'), g = kid(next, 'defRPr');
        if (g) out = K.merge(out, { ['{' + N.a + '}defRPr']: r ? D.property(raw(r), raw(g), current?.[i]?.rPr, before?.[i]?.rPr, D.runFields, 'a:CT_TextCharacterProperties') : raw(g) }, 'a:CT_TextParagraphProperties');
      }
      replacements['{' + N.a + '}' + name] = out;
    }
    return K.mergeBag(xml, replacements);
  }
  function theme(d, fragment, generated) {
    const before = d.keep.before, fresh = K.parse(generated);
    let xml = fragment.xml;
    for (const slot of Object.keys(d.colors)) if (!same(d.colors[slot], before.colors[slot])) {
      const root = K.parse(xml), scheme = find(root, 'clrScheme'), next = kid(find(fresh, 'clrScheme'), slot);
      if (scheme && next) xml = changeAt(xml, scheme, K.mergeBag(raw(scheme), { [namespaced(next)]: raw(next) }));
    }
    for (const kind of ['major', 'minor']) if (!same(d.fonts[kind], before.fonts[kind])) {
      const original = find(K.parse(xml), kind + 'Font'), next = find(fresh, kind + 'Font');
      if (original && next) xml = changeAt(xml, original, K.mergeBag(raw(original), { ['{' + N.a + '}latin']: raw(kid(next, 'latin')) }));
    }
    return K.slice(fragment, xml);
  }
  function decorations(xml, d, fragment, current, before, ctx, emitShape) {
    const tree = find(K.parse(xml), 'spTree'), children = kids(tree);
    const slots = d.keep.decorations?.[fragment.part] || [], owners = new Map();
    for (const slot of slots) for (const id of slot.models) owners.set(id, slot);
    const old = new Map((before || []).map(s => [s.id, s])), units = [];
    for (const shape of current || []) {
      const owner = owners.get(shape.id), last = units[units.length - 1];
      if (owner && last?.owner === owner) last.shapes.push(shape);
      else units.push({ owner, shapes: [shape] });
    }
    for (const unit of units) {
      const unchanged = unit.owner && same(unit.shapes.map(s => s.id), unit.owner.models) && unit.shapes.every(s => same(s, old.get(s.id)));
      unit.xml = unchanged ? raw(children[unit.owner.index]) : unit.shapes.map(s => emitShape(s, ctx, d)).join('');
    }
    // Ordinary property edits and deletions stay in their original slots, with
    // unread shapes and placeholders between them untouched. Only a deliberate
    // reorder moves surviving decorations between those slots.
    const ordered = units.filter(u => u.owner), reordered = ordered.some((u, i) => i && u.owner.index < ordered[i - 1].owner.index);
    const live = slots.filter(slot => ordered.some(u => u.owner === slot));
    const replacements = new Map(slots.map(s => [s.index, '']));
    let pending = '', previous, cursor = 0;
    for (const unit of units) {
      if (!unit.owner) { pending += unit.xml; continue; }
      const slot = reordered ? live[Math.min(cursor++, live.length - 1)] : unit.owner;
      replacements.set(slot.index, replacements.get(slot.index) + pending + unit.xml);
      pending = ''; previous = slot;
    }
    if (previous) { replacements.set(previous.index, replacements.get(previous.index) + pending); pending = ''; }
    const edits = [...replacements].map(([index, value]) => { const p = X.source.get(children[index]); return { start: p.start, end: p.end, value }; });
    if (pending) {
      const ext = children.find(e => e.localName === 'extLst'), pos = X.source.get(tree);
      const at = ext ? X.source.get(ext).start : pos.end - tree.nodeName.length - 3;
      edits.push({ start: at, end: at, value: pending });
    }
    return K.patch(xml, edits);
  }
  function content(d, record, generated, ctx, emitShape, pres) {
    const before = d.keep.before, master = !record.lkey, fragment = record.fragment;
    const fresh = K.parse(generated); let xml = fragment.xml;
    const currentDeco = master ? d.deco : d.layoutDecos?.[record.lkey];
    const oldDeco = master ? before.deco : before.layoutDecos?.[record.lkey];
    const title = !master && record.key === 'title';
    const titleChanged = title && !same(d.titleDeco, before.titleDeco);
    if (!same(currentDeco, oldDeco) || titleChanged) {
      xml = decorations(xml, d, fragment, titleChanged && d.titleDeco ? d.titleDeco : currentDeco, oldDeco, ctx, emitShape);
    }
    const background = master ? d.bg : title && !same(d.titleBg, before.titleBg) ? d.titleBg : d.layoutBgs?.[record.lkey];
    const oldBackground = master ? before.bg : title && !same(d.titleBg, before.titleBg) ? before.titleBg : before.layoutBgs?.[record.lkey];
    if (!same(background, oldBackground)) {
      const common = kid(K.parse(xml), 'cSld'), bg = kid(kid(fresh, 'cSld'), 'bg');
      xml = changeAt(xml, common, changeChild(raw(common), 'bg', raw(bg), 'p:CT_CommonSlideData'));
    }
    if (titleChanged) xml = K.attributes(xml, { showMasterSp: d.titleDeco ? '0' : null });
    const originalSize = pres.keep?.values, sx = originalSize?.W ? pres.W / originalSize.W : 1, sy = originalSize?.H ? pres.H / originalSize.H : 1;
    for (const shape of kids(find(K.parse(xml), 'spTree')).filter(e => find(e, 'ph'))) {
      const ph = find(shape, 'ph'), type = ph.getAttribute('type') || 'body';
      const key = type === 'obj' ? 'body' : type, frame = d.ph?.[key];
      let next = raw(shape);
      if (frame && (master || ['ctrTitle', 'subTitle'].includes(key)) && !same(frame, before.ph?.[key])) next = K.setBox(next, frame);
      else if (!master && (sx !== 1 || sy !== 1)) {
        const box = K.getBox(next);
        if (box.x != null) next = K.setBox(next, { ...box, x: box.x * sx, y: box.y * sy, w: box.w * sx, h: box.h * sy });
      }
      const footer = ['dt', 'ftr', 'sldNum'].includes(key), alignment = key === 'ctrTitle' ? 'ctrTitleAlgn' : key === 'subTitle' ? 'subTitleAlgn' : null;
      if (footer && !same(d.footer, before.footer) || alignment && !same(d[alignment], before[alignment])) {
        const body = kid(K.parse(next), 'txBody');
        if (body) {
          const list = kid(body, 'lstStyle'), level = kid(list, 'lvl1pPr');
          let lv = raw(level) || '<a:lvl1pPr xmlns:a="' + N.a + '"/>';
          if (alignment) lv = K.attributes(lv, { algn: d[alignment] || null });
          if (footer) {
            const freshShape = kids(find(fresh, 'spTree')).find(e => find(e, 'ph')?.getAttribute('type') === key);
            const r = kid(level, 'defRPr'), g = find(freshShape, 'defRPr');
            if (g) lv = K.merge(lv, { ['{' + N.a + '}defRPr']: r ? D.property(raw(r), raw(g), d.footer?.rPr, before.footer?.rPr, D.runFields, 'a:CT_TextCharacterProperties') : raw(g) }, 'a:CT_TextParagraphProperties');
          }
          const lst = K.mergeBag(raw(list) || '<a:lstStyle xmlns:a="' + N.a + '"/>', { ['{' + N.a + '}lvl1pPr']: lv });
          next = changeAt(next, body, K.merge(raw(body), { ['{' + N.a + '}lstStyle']: lst }, 'a:CT_TextBody'));
        }
      }
      if (next !== raw(shape)) {
        const current = kids(find(K.parse(xml), 'spTree')).find(e => raw(e) === raw(shape));
        if (current) xml = changeAt(xml, current, next);
      }
    }
    if (master) {
      if (!same(d.clrMap, before.clrMap)) {
        const cm = kid(K.parse(xml), 'clrMap');
        if (cm) xml = changeAt(xml, cm, K.attributes(raw(cm), d.clrMap || {}));
      }
      for (const kind of ['title', 'body', 'other']) if (!same(d.tx?.[kind], before.tx?.[kind])) {
        const old = find(K.parse(xml), kind + 'Style'), next = find(fresh, kind + 'Style');
        if (old && next) xml = changeAt(xml, old, levels(raw(old), raw(next), d.tx[kind], before.tx?.[kind]));
      }
    }
    return K.slice(fragment, xml);
  }
  D.write = function (d, pres, pack, api) {
    const keep = d.keep, pkg = keep && K.package(keep.source); if (!keep?.master || !pkg) return null;
    // Reject an incomplete graph before claiming any output parts, so the
    // existing recovery writer can still produce a complete converted design.
    try { K.subtree(pkg, keep.part); }
    catch (error) { error.code = 'OOXML_DESIGN_DEPENDENCY'; throw error; }
    const writer = pack.writer, copied = keep.design !== d.id || pkg !== pres.pkg, parts = new Map();
    const allocate = part => {
      if (!part) return null;
      if (parts.has(part)) return parts.get(part);
      const slash = part.lastIndexOf('/'), dot = part.lastIndexOf('.');
      const target = copied ? api.name(part.slice(0, slash), part.slice(slash + 1, dot).replace(/\d+$/, ''), part.slice(dot + 1)) : writer.target(pkg, part);
      parts.set(part, target); if (!copied) writer.claim(part, 'regenerated', target, pkg);
      return target;
    };
    const master = allocate(keep.part), layouts = keep.layoutParts || [], targets = new Map();
    layouts.forEach(r => allocate(r.part));
    const themeChanged = !same([d.colors, d.fonts], [keep.before.colors, keep.before.fonts]);
    const themePart = keep.theme && pkg.has(keep.theme) ? themeChanged || copied ? api.name('ppt/theme', 'theme', 'xml') : writer.target(pkg, keep.theme) : null;
    if (themePart) parts.set(keep.theme, themePart);
    const local = Object.create(writer);
    local.keepRel = (owner, dep) => {
      if (dep.source === pkg.id && !dep.external && parts.has(dep.part)) return writer.rels(owner).add(dep.type, K.relative(owner, parts.get(dep.part)) + (dep.fragment || ''), false, dep.id);
      return writer.keepRel(owner, dep);
    };
    const ctxFor = (part, shapes) => {
      const ctx = api.context(part, shapes, local);
      return ctx;
    };
    const emit = (fragment, target) => local.emit(fragment, target, { references: fragment.ids });
    const carryRels = (part, target, whole) => {
      for (const r of pkg.rels(part)) {
        const type = K.relationshipType(r.type);
        if (!whole && (/\/(slideMaster|slideLayout|theme|image|chart|hyperlink|tags)$/.test(type) || L.frames.consumes(pres, part, r))) continue;
        try { local.keepRel(target, { source: pkg.id, owner: part, ...r }); }
        catch (error) { writer.loss({ id: 'design-dependency:' + part + ':' + r.id, what: 'A design dependency could not be retained: ' + error.message, where: part, action: 'drop' }); }
      }
    };
    function write(fragment, target, part, original = false) {
      const xml = original && !copied ? pkg.bytes(part) : K.hoistNamespaces(emit(fragment, target));
      // Unchanged content is copied exactly; relations still follow copied designs.
      if (original && !copied) for (const dep of fragment.deps) local.keepRel(target, dep);
      writer.put(target, xml, pkg.type(part)); carryRels(part, target, original);
    }
    if (themePart) {
      const fragment = K.fragment(pkg.xml(keep.theme), { pkg, part: keep.theme });
      write(themeChanged ? theme(d, fragment, api.theme(d)) : fragment, themePart, keep.theme, !themeChanged);
      if (themeChanged && !Object.values(pres.designs).some(other => other !== d && other.keep?.theme === keep.theme && same([other.colors, other.fonts], [keep.before.colors, keep.before.fonts]))) writer.classes.set(writer.key(pkg, keep.theme), 'regenerated');
    }
    const masterTree = pkg.xml(keep.part), oldIds = kids(kid(masterTree, 'sldLayoutIdLst'));
    const masterRels = pkg.rels(keep.part), layoutIds = [];
    for (const record of layouts) {
      const target = parts.get(record.part), ctx = ctxFor(target, d.layoutDecos?.[record.lkey] || []);
      const generated = api.layout(d, record.key, record.lkey, { ...ctx, designTemplate: true });
      const changed = content(d, record, generated, ctx, api.shape, pres);
      write(changed, target, record.part, changed.xml === record.fragment.xml);
      local.keepRel(target, { source: pkg.id, owner: record.part, ...pkg.rels(record.part).find(r => /\/slideMaster$/.test(r.type)) });
      const rel = masterRels.find(r => r.part === record.part && /\/slideLayout$/.test(r.type));
      const oldId = oldIds.find(e => e.getAttributeNS(N.rel, 'id') === rel?.id || e.getAttributeNS(N.strictRel, 'id') === rel?.id)?.getAttribute('id');
      const id = !copied && oldId ? oldId : writer.ids.fresh('presentation', 'sldLayoutId');
      const rid = writer.rels(master).add(RT('slideLayout'), K.relative(master, target), false, rel?.id);
      layoutIds.push({ id, rid }); targets.set(record.key + '|' + record.lkey, target);
    }
    for (const [key, part] of Object.entries(keep.layouts)) if (parts.has(part)) targets.set(key, parts.get(part));
    for (const slide of pres.slides.filter(s => s.design === d.id)) {
      const key = slide.layout + (slide.lkey ? '|' + slide.lkey : ''); if (targets.has(key)) continue;
      const target = api.name('ppt/slideLayouts', 'slideLayout', 'xml'), ctx = ctxFor(target, []);
      writer.put(target, api.layout(d, slide.layout, slide.lkey, ctx), 'application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml');
      writer.rels(target).add(RT('slideMaster'), K.relative(target, master));
      layoutIds.push({ id: writer.ids.fresh('presentation', 'sldLayoutId'), rid: writer.rels(master).add(RT('slideLayout'), K.relative(master, target)) });
      targets.set(key, target);
    }
    const ctx = ctxFor(master, d.deco || []), generated = api.master(d, { ...ctx, designTemplate: true }, layoutIds);
    let fragment = content(d, { fragment: keep.master }, generated, ctx, api.shape, pres);
    if (copied || layoutIds.length !== oldIds.length) {
      const list = kid(K.parse(generated), 'sldLayoutIdLst');
      fragment = K.slice(fragment, changeChild(fragment.xml, 'sldLayoutIdLst', raw(list), 'p:CT_SlideMaster'));
    }
    write(fragment, master, keep.part, fragment.xml === keep.master.xml);
    if (themePart) writer.rels(master).add(RT('theme'), K.relative(master, themePart), false, masterRels.find(r => /\/theme$/.test(r.type))?.id);
    const rel = pkg.rels(pkg.main).find(r => r.part === keep.part && /\/slideMaster$/.test(r.type));
    const entry = all(pkg.xml(pkg.main)).find(e => e.localName === 'sldMasterId' && (e.getAttributeNS(N.rel, 'id') || e.getAttributeNS(N.strictRel, 'id')) === rel?.id);
    const id = !copied && entry?.getAttribute('id') || writer.ids.fresh('presentation', 'sldMasterId');
    return { part: master, theme: themePart, id, rid: !copied ? rel?.id : null, layouts: targets };
  };
  const RT = name => N.rel + '/' + name;
})();
