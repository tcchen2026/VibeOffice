/* Quire — retain drawing properties independently of the fields an edit owns. */
(function () {
  'use strict';
  const L = window.L, P = L.preserve, K = L.opc, N = K.NS;
  const kids = e => Array.from(e?.children || []), child = (e, name) => kids(e).find(c => c.localName === name);
  const raw = e => e ? K.raw(e) : '';
  const value = v => JSON.stringify(v ?? null);
  const names = new Set(['pic', 'sp', 'wsp', 'wgp', 'wpc', 'grpSp', 'lockedCanvas', 'shape', 'rect', 'roundrect', 'oval', 'line', 'group', 'image', 'polyline', 'arc']);
  const vml = e => e.namespaceURI === K.KNOWN_NS.v;
  const decl = Object.entries(K.KNOWN_NS).map(([p, ns]) => 'xmlns:' + p + '="' + ns + '"').join(' ');
  const find = (e, name) => [e, ...Array.from(e?.getElementsByTagName('*') || [])].find(c => c?.localName === name);
  function outer(tree) {
    const frames = [], anchors = [], metadata = [];
    function scan(e) {
      if (names.has(e.localName)) { frames.push(e); return; }
      if (['anchor', 'inline'].includes(e.localName) && /wordprocessingDrawing$/.test(e.namespaceURI || '')) anchors.push(e);
      if (e.localName === 'docPr') metadata.push(e);
      kids(e).forEach(scan);
    }
    scan(tree); return { frames, anchors, metadata };
  }
  const meta = ['name', 'alt', 'title', 'hidden', 'link'];
  const picture = ['crop', 'border', 'gray', 'bw', 'bright', 'contrast', 'washout', 'alpha', 'geom'];
  const shape = ['fill', 'line', 'shadow', 'geom', 'adj', 'paths', 'ins', 'anchor', 'autofit', 'noWrap', 'vert'];
  P.drawingProperties = function (object, fragment) {
    if (!['img', 'shape', 'group', 'chart'].includes(object.t)) return null;
    const nodes = outer(K.parse(fragment.xml));
    let fields = nodes.anchors.length ? [...meta, 'float'] : [];
    if (nodes.frames.length && !object.diagram && !object.wordart) {
      const editable = nodes.frames.every(e => vml(e) ? e.localName !== 'group' : ['pic', 'wsp', 'sp'].includes(e.localName));
      if (editable && object.t === 'img' && nodes.frames.every(e => vml(e) ? !!child(e, 'imagedata') : e.localName === 'pic')) fields.push(...picture);
      if (editable && object.t === 'shape' && nodes.frames.every(e => vml(e) ? !child(e, 'imagedata') : ['wsp', 'sp'].includes(e.localName))) fields.push(...shape);
    }
    if (nodes.frames.some(vml)) {
      // VML geometry and text layout have different models. Those edits still
      // convert; ordinary fill/line/shadow and picture adjustments have adapters.
      fields = fields.filter(k => !['name', 'float', 'geom', 'adj', 'paths', 'alpha', 'ins', 'anchor', 'autofit', 'noWrap', 'vert'].includes(k));
      fields.push('alt', 'title', 'hidden', 'link');
    }
    return Object.fromEntries([...new Set(fields)].map(k => [k, value(object[k])]));
  };
  function rewrite(xml, select, replace) {
    const tree = K.parse(xml);
    return K.patch(xml, select(tree).map(e => ({ ...L.xmlTree.source.get(e), value: replace(e) })));
  }
  function merge(xml, changes, type) {
    const alternative = e => e.namespaceURI === N.mc && ['AlternateContent', 'Choice', 'Fallback'].includes(e.localName);
    const contents = e => kids(e).flatMap(c => alternative(c) ? contents(c) : [c]);
    const choices = [['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'], ['prstGeom', 'custGeom'], ['noAutofit', 'normAutofit', 'spAutoFit']];
    const groups = new Map();
    for (const name of Object.keys(changes)) {
      const group = choices.find(g => g.includes(name))?.join(',') || name;
      if (!groups.has(group)) groups.set(group, {});
      groups.get(group)[name] = changes[name];
    }
    for (const group of groups.values()) {
      const relevant = e => Object.hasOwn(group, e.localName) && (e.namespaceURI === N.a || e.namespaceURI === 'http://purl.oclc.org/ooxml/drawingml/main');
      function branch(text, outer = false) {
        let tree = K.parse(text);
        const present = contents(tree).some(relevant);
        text = K.patch(text, kids(tree).filter(alternative).map(e => ({ ...L.xmlTree.source.get(e), value: branch(raw(e)) })));
        tree = K.parse(text);
        if (tree.namespaceURI === N.mc && tree.localName === 'AlternateContent') return text;
        if (!kids(tree).some(relevant) && !(outer && !present)) return text;
        const replacements = Object.fromEntries(Object.entries(group).map(([k, v]) => ['{' + N.a + '}' + k, (typeof v === 'function' ? v(child(tree, k)) : v) || []]));
        return K.merge(text, replacements, type);
      }
      xml = branch(xml, true);
    }
    return xml;
  }
  function metadata(xml, object, changed, generated) {
    const attrs = {};
    for (const [k, a] of [['name', 'name'], ['alt', 'descr'], ['title', 'title'], ['hidden', 'hidden']]) if (changed(k)) attrs[a] = k === 'hidden' ? object.hidden ? '1' : null : object[k] || null;
    xml = K.attributes(xml, attrs);
    return changed('link') ? merge(xml, { hlinkClick: raw(child(generated, 'hlinkClick')) }, 'a:CT_NonVisualDrawingProps') : xml;
  }
  function style(xml, changes) {
    const el = K.parse(xml);
    let text = el.getAttribute('style') || '';
    for (const [key, v] of Object.entries(changes)) {
      const re = new RegExp('(^|;)(\\s*' + key + '\\s*:)[^;]*');
      text = re.test(text) ? text.replace(re, (_, before, label) => before + (v == null ? '' : label + v)) : v == null ? text : text + (text && !text.endsWith(';') ? ';' : '') + key + ':' + v;
    }
    return K.attributes(xml, { style: text });
  }
  function vmlChild(xml, name, replacement) {
    const root = K.parse(xml), el = child(root, name), p = L.xmlTree.source.get(el || root);
    if (el) return K.patch(xml, [{ ...p, value: replacement }]);
    if (!replacement) return xml;
    const self = xml[p.openEnd - 2] === '/';
    const at = self ? p.openEnd - 2 : xml.lastIndexOf('</', p.end - 1);
    return K.patch(xml, [{ start: at, end: self ? p.openEnd : at, value: self ? '>' + replacement + '</' + root.nodeName + '>' : replacement }]);
  }
  const tag = (name, attrs) => K.attributes('<v:' + name + ' xmlns:v="' + K.KNOWN_NS.v + '"/>', attrs);
  function vmlProperties(el, object, changed) {
    let xml = raw(el), attrs = {};
    if (changed('alt')) attrs.alt = object.alt || null;
    if (changed('title')) attrs.title = object.title || null;
    if (changed('link')) attrs.href = object.link?.url || null;
    if (changed('hidden')) xml = style(xml, { visibility: object.hidden ? 'hidden' : null });
    if (changed('fill')) {
      const f = object.fill || { t: 'none' };
      if (!['none', 'solid'].includes(f.t)) throw new Error('This VML fill requires conversion.');
      attrs.filled = f.t === 'none' ? 'f' : 't'; attrs.fillcolor = f.c || null;
      xml = vmlChild(xml, 'fill', f.t === 'none' ? '' : tag('fill', { color: f.c, opacity: f.a == null ? '1' : f.a }));
    }
    if (changed('line') || changed('border')) {
      const border = object.border;
      const line = object.t === 'img' ? border && border.val !== 'nil' && { c: '#' + (border.color === 'auto' ? '000000' : (border.color || '000000').replace(/^#/, '')), w: border.sz, dash: border.val === 'dashed' ? 'dash' : 'solid' } : object.line;
      attrs.stroked = !line || line.t === 'none' ? 'f' : 't'; attrs.strokecolor = line?.c || null; attrs.strokeweight = line?.w != null ? line.w + 'pt' : null;
      const arrow = a => ({ triangle: 'block', stealth: 'classic', arrow: 'open', diamond: 'diamond', oval: 'oval' }[a?.type] || 'none');
      const dash = { sysDot: 'shortdot', sysDash: 'shortdash', lgDash: 'longdash', lgDashDot: 'longdashdot', lgDashDotDot: 'longdashdotdot', sysDashDot: 'shortdashdot', sysDashDotDot: 'shortdashdotdot' }[line?.dash] || line?.dash || 'solid';
      xml = vmlChild(xml, 'stroke', attrs.stroked === 'f' ? '' : tag('stroke', { color: line.c, weight: (line.w || 0.75) + 'pt', dashstyle: dash, startarrow: arrow(line.head), endarrow: arrow(line.tail) }));
    }
    if (changed('shadow')) {
      const s = object.shadow;
      xml = vmlChild(xml, 'shadow', s ? tag('shadow', { on: 't', color: s.c || '#000000', opacity: s.a ?? 0.5, offset: (s.dx ?? 3) + 'pt,' + (s.dy ?? 3) + 'pt' }) : '');
    }
    const image = child(K.parse(xml), 'imagedata');
    if (image) {
      const a = {};
      if (changed('crop')) for (const [side, key] of [['left', 'l'], ['right', 'r'], ['top', 't'], ['bottom', 'b']]) a['crop' + side] = object.crop?.[key] || null;
      if (changed('gray')) a.grayscale = object.gray ? 't' : null;
      if (changed('bw')) a.bilevel = object.bw ? 't' : null;
      if (changed('bright') || changed('washout')) a.blacklevel = object.washout ? '22938f' : Math.round((object.bright || 0) * 65536) + 'f';
      if (changed('contrast') || changed('washout')) a.gain = object.washout ? '19661f' : Math.round((1 + (object.contrast || 0)) * 65536) + 'f';
      if (Object.keys(a).length) xml = vmlChild(xml, 'imagedata', K.attributes(raw(image), a));
    }
    return K.attributes(xml, attrs);
  }
  P.mergeDrawing = function (xml, object, previous, ctx) {
    const changed = k => Object.hasOwn(previous, k) && previous[k] !== value(object[k]);
    if (!Object.keys(previous).some(changed)) return xml;
    const fresh = K.parse('<root ' + decl + '>' + ctx.drawingXML(object, ctx) + '</root>');
    const generated = outer(fresh), frame = generated.frames[0], sp = child(frame, 'spPr'), pic = child(frame, 'blipFill'), body = child(frame, 'bodyPr');
    xml = rewrite(xml, tree => outer(tree).frames, el => {
      if (vml(el)) return vmlProperties(el, object, changed);
      let result = raw(el);
      if (['fill', 'line', 'border', 'shadow', 'geom', 'adj', 'paths'].some(changed)) result = rewrite(result, root => kids(root).filter(e => e.localName === 'spPr'), e => {
        const changes = {};
        if (changed('fill')) for (const name of ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']) changes[name] = raw(child(sp, name));
        if (changed('line') || changed('border')) changes.ln = raw(child(sp, 'ln'));
        if (['geom', 'adj', 'paths'].some(changed)) for (const name of ['prstGeom', 'custGeom']) changes[name] = raw(child(sp, name));
        if (changed('shadow')) {
          if (find(e, 'effectDag')) throw new Error('This effect graph requires conversion when its shadow changes.');
          changes.effectLst = original => merge(raw(original) || '<a:effectLst xmlns:a="' + N.a + '"/>', { outerShdw: raw(child(child(sp, 'effectLst'), 'outerShdw')) }, 'a:CT_EffectList');
        }
        return merge(raw(e), changes, 'a:CT_ShapeProperties');
      });
      if (picture.some(changed)) result = rewrite(result, root => kids(root).filter(e => e.localName === 'blipFill'), e => {
        let b = raw(e);
        if (changed('crop')) b = merge(b, { srcRect: raw(child(pic, 'srcRect')) }, 'a:CT_BlipFillProperties');
        const changes = {};
        for (const [keys, name] of [[['gray'], 'grayscl'], [['bw'], 'biLevel'], [['bright', 'contrast', 'washout'], 'lum'], [['alpha'], 'alphaModFix']]) if (keys.some(changed)) changes[name] = raw(child(child(pic, 'blip'), name));
        if (Object.keys(changes).length) b = rewrite(b, root => kids(root).filter(c => c.localName === 'blip'), c => merge(raw(c), changes, 'a:CT_Blip'));
        return b;
      });
      if (['ins', 'anchor', 'autofit', 'noWrap', 'vert'].some(changed)) result = rewrite(result, root => kids(root).filter(e => e.localName === 'bodyPr'), e => {
        const attrs = {};
        for (const [field, names] of [['ins', ['lIns', 'rIns', 'tIns', 'bIns']], ['anchor', ['anchor']], ['noWrap', ['wrap']], ['vert', ['vert']]]) if (changed(field)) for (const name of names) attrs[name] = body?.getAttribute(name);
        let b = K.attributes(raw(e), attrs);
        if (changed('autofit')) b = merge(b, Object.fromEntries(['noAutofit', 'normAutofit', 'spAutoFit'].map(n => [n, raw(child(body, n))])), 'a:CT_TextBodyProperties');
        return b;
      });
      if (meta.some(changed)) result = rewrite(result, root => {
        const nv = child(root, 'nvPicPr') || child(root, 'nvSpPr');
        return [child(root, 'cNvPr'), child(nv, 'cNvPr')].filter(Boolean);
      }, e => metadata(raw(e), object, changed, generated.metadata[0]));
      return result;
    });
    if (meta.some(changed)) xml = rewrite(xml, tree => outer(tree).metadata, e => metadata(raw(e), object, changed, generated.metadata[0]));
    if (changed('float')) xml = rewrite(xml, tree => outer(tree).anchors, e => {
      const next = generated.anchors[0], before = JSON.parse(previous.float), after = object.float;
      if (!next) throw new Error('The drawing has no generated anchor.');
      if (!before || !after) {
        // Changing inline/floating status replaces the anchor shell, while its
        // original graphic, identities, extensions and effects remain attached.
        const changes = {};
        for (const name of ['extent', 'effectExtent', 'docPr', 'cNvGraphicFramePr', 'graphic']) {
          const old = child(e, name); if (old) changes['{' + (name === 'graphic' ? N.a : K.KNOWN_NS.wp) + '}' + name] = raw(old);
        }
        return K.merge(raw(next), changes, 'wp:CT_' + (after ? 'Anchor' : 'Inline'));
      }
      const attrs = {}, changes = {}, eq = k => value(before[k]) === value(after[k]);
      for (const [field, a] of [['z', 'relativeHeight'], ['locked', 'locked'], ['layoutInCell', 'layoutInCell'], ['allowOverlap', 'allowOverlap']]) if (!eq(field)) attrs[a] = next.getAttribute(a);
      if (!eq('dist')) for (const a of ['distT', 'distB', 'distL', 'distR']) attrs[a] = next.getAttribute(a);
      if (!eq('wrap') || !eq('behind') || !eq('side')) {
        attrs.behindDoc = next.getAttribute('behindDoc');
        for (const n of ['wrapSquare', 'wrapTight', 'wrapThrough', 'wrapTopAndBottom', 'wrapNone']) changes['{' + K.KNOWN_NS.wp + '}' + n] = raw(child(next, n)) || [];
      }
      for (const [field, name] of [['posH', 'positionH'], ['posV', 'positionV']]) if (!eq(field)) changes['{' + K.KNOWN_NS.wp + '}' + name] = raw(child(next, name));
      return K.merge(K.attributes(raw(e), attrs), changes, 'wp:CT_Anchor');
    });
    return xml;
  };
})();
