/* PowerPoint animation behaviors and PresentationML equivalents.
 * Derived from Apache OpenOffice sd/source/filter/ppt/pptinanimations.cxx (Apache-2.0).
 * MS-PPT 2.8.29–2.8.72; ECMA-376 common behavior and animation elements.
 * Formula/path strings are document data, never JavaScript. */
(function (root) {
  'use strict';
  const B = root.L.legacy, P = B.ppt, A = P.timing, R = A.records, Q = A.behaviorProperties, el = B.el;
  const kinds = new Map([[R.animate, 'anim'], [R.color, 'animClr'], [R.effect, 'animEffect'],
    [R.motion, 'animMotion'], [R.rotation, 'animRot'], [R.scale, 'animScale'],
    [R.set, 'set'], [R.command, 'cmd']]);
  const atomTypes = new Map([[R.animate, R.animateAtom], [R.color, R.colorAtom],
    [R.effect, R.effectAtom], [R.motion, R.motionAtom], [R.rotation, R.rotationAtom],
    [R.scale, R.scaleAtom], [R.set, R.setAtom], [R.command, R.commandAtom]]);
  A.readBehavior = function (record) {
    const children = P.children(record), atom = children.find(c => c.type === atomTypes.get(record.type));
    const common = children.find(c => c.type === R.behavior);
    if (!common || !atom) A.invalid('BEHAVIOR');
    const cc = P.children(common), r = cc.find(c => c.type === R.behaviorAtom)?.r;
    if (!r) A.invalid('BEHAVIOR');
    r.check(0, 16);
    const bits = r.u32(0), properties = A.propertySet(cc.find(c => c.type === R.properties));
    const behavior = { kind: kinds.get(record.type), properties,
      target: A.target(cc.find(c => c.type === R.target)), names: [], values: new Map(), keys: [] };
    if (bits & 1) behavior.additive = r.u32(4);
    if (bits & 2) behavior.accumulate = r.u32(8);
    if (bits & 8) behavior.transform = r.u32(12);
    if (bits & 4) {
      const names = cc.find(c => c.type === R.names);
      for (const item of names ? P.children(names) : [])
        if (item.type === R.variant) behavior.names.push(String(A.variant(item).value));
    }
    for (const item of children)
      if (item.type === R.variant) behavior.values.set(item.instance, A.variant(item));
    const keyList = children.find(c => c.type === R.values);
    if (keyList) {
      const list = P.children(keyList);
      for (let i = 0; i < list.length; i++) {
        if (list[i].type !== R.time) continue;
        const time = list[i].r.i32(0), value = list[++i];
        if (!value || value.type !== R.variant) A.invalid('KEYFRAME');
        const key = { time, value: A.variant(value) };
        if (list[i + 1]?.type === R.variant) key.formula = A.variant(list[++i]).value;
        behavior.keys.push(key);
      }
    }
    const d = atom.r;
    const floats = (offset, count) => Array.from({ length: count }, (_, i) => A.number(d.f32(offset + 4 * i)));
    if (record.type === R.animate) {
      d.check(0, 12);
      behavior.flags = d.u32(4);
      behavior.calc = behavior.flags & 8 ? d.u32(0) : 1;
      behavior.valueType = behavior.flags & 32 ? d.u32(8) : 1;
    } else {
      behavior.flags = d.u32(0);
      if ([R.effect, R.set, R.command].includes(record.type)) behavior.mode = d.u32(4);
      if (record.type === R.motion || record.type === R.scale) {
        d.check(0, 32);
        for (let i = 0; i < 3; i++)
          if (behavior.flags & (1 << i)) behavior[['by', 'from', 'to'][i]] = floats(4 + i * 8, 2);
        behavior.mode = d.u32(28);
      } else if (record.type === R.rotation) {
        d.check(0, 20);
        for (let i = 0; i < 3; i++)
          if (behavior.flags & (1 << i)) behavior[['by', 'from', 'to'][i]] = floats(4 + i * 4, 1)[0];
      } else if (record.type === R.color) {
        d.check(0, 52);
        for (let i = 0; i < 3; i++)
          if (behavior.flags & (1 << i))
            behavior[['by', 'from', 'to'][i]] = Array.from({ length: 4 }, (_, j) => d.i32(4 + 16 * i + 4 * j));
      }
    }
    return behavior;
  };
  const integer = (n, scale = 1) => {
    const value = Math.round(A.number(n) * scale);
    if (Math.abs(value) > 2147483647) A.invalid('RANGE');
    return value;
  };
  const scalar = value => value == null ? null : typeof value.value === 'boolean' ? +value.value : value.value;
  A.valueXml = value => {
    const tag = ['boolVal', 'intVal', 'fltVal', 'strVal'][value.type];
    if (!tag) A.invalid('VALUE_TYPE');
    return el('p:' + tag, { val: scalar(value) });
  };
  const color = (value, page, delta = false) => {
    if (!value) return '';
    const [mode, a, b, c] = value;
    if (delta) {
      if (mode === 0) return el('p:rgb', { r: integer(a / 255, 100000),
        g: integer(b / 255, 100000), b: integer(c / 255, 100000) });
      if (mode === 1) return el('p:hsl', { h: integer(a, 60000), s: integer(b, 1000), l: integer(c, 1000) });
      A.invalid('COLOR');
    }
    if (mode === 0) {
      if ([a, b, c].some(v => v < 0 || v > 255)) A.invalid('COLOR');
      return el('a:srgbClr', { val: [a, b, c].map(v => v.toString(16).padStart(2, '0')).join('').toUpperCase() });
    }
    if (mode === 2 && page.scheme?.[a]) return el('a:srgbClr', { val: page.scheme[a] });
    A.invalid('COLOR');
  };
  A.behaviorXml = function (node, context) {
    const v = node.behavior, flags = v.flags, attrs = {}, values = v.values;
    const props = v.properties, kind = v.kind;
    const runtime = props.get(Q.context);
    if (runtime != null && runtime !== 'PPT') A.invalid('CONTEXT');
    const commonAttrs = {};
    if (v.additive != null) {
      commonAttrs.additive = ['base', 'sum', 'repl', 'mult', 'none'][v.additive];
      if (!commonAttrs.additive) A.invalid('ADDITIVE');
    }
    if (v.accumulate != null) commonAttrs.accumulate = v.accumulate ? 'always' : 'none';
    if (v.transform != null) commonAttrs.xfrmType = v.transform ? 'img' : 'pt';
    if (props.has(Q.override)) commonAttrs.override = props.get(Q.override) ? 'childStyle' : 'normal';
    const target = A.targetXml(v.target, context);
    if (!target) A.invalid('TARGET');
    const common = el('p:cBhvr', commonAttrs, A.commonXml(node, context) + target +
      (v.names.length ? el('p:attrNameLst', {}, v.names.map(name => el('p:attrName', {}, B.xml(name))).join('')) : ''));
    let body = '';
    if (kind === 'set') {
      const to = values.get(1) || values.get(0);
      if (!to || !(flags & 1)) A.invalid('SET');
      body = el('p:to', {}, A.valueXml(to));
    } else if (kind === 'anim') {
      attrs.calcmode = ['discrete', 'lin', 'fmla'][v.calc];
      attrs.valueType = ['str', 'num', 'clr'][v.valueType];
      if (!attrs.calcmode || !attrs.valueType) A.invalid('ANIMATE');
      for (let i = 0; i < 3; i++)
        if (flags & (1 << i)) attrs[['by', 'from', 'to'][i]] = scalar(values.get(i + 1));
      if (flags & 16 && v.keys.length)
        body = el('p:tavLst', {}, v.keys.map(key => {
          if ((key.time < 0 && key.time !== -1000 && key.time !== -1) || key.time > 1000) A.invalid('KEY_TIME');
          const time = key.time === -1 ? null : key.time === -1000 ? 'indefinite' : key.time * 100;
          return el('p:tav', { tm: time, fmla: key.formula || null },
            el('p:val', {}, A.valueXml(key.value)));
        }).join(''));
    } else if (kind === 'animEffect') {
      attrs.transition = flags & 1 && v.mode ? 'out' : 'in';
      if (flags & 2) attrs.filter = scalar(values.get(1));
      if (flags & 4 && values.has(2)) body = el('p:progress', {}, A.valueXml(values.get(2)));
    } else if (kind === 'cmd') {
      attrs.type = flags & 1 ? ['evt', 'call', 'verb'][v.mode] : 'call';
      attrs.cmd = flags & 2 ? scalar(values.get(1)) : null;
      if (!attrs.type || attrs.cmd == null) A.invalid('COMMAND');
    } else if (kind === 'animMotion' || kind === 'animScale') {
      if (kind === 'animMotion') {
        attrs.origin = flags & 8 && v.mode < 2 ? 'parent' : 'layout';
        if (flags & 16) attrs.path = scalar(values.get(1));
        attrs.pathEditMode = props.get(Q.relativePath) === false ? 'fixed' : 'relative';
        if (flags & 128) attrs.ptsTypes = props.get(Q.pointsTypes);
        if (flags & 64 && props.has(Q.pathRotation)) attrs.rAng = integer(props.get(Q.pathRotation), 60000);
      } else if (flags & 8) attrs.zoomContents = +!!v.mode;
      for (const name of ['by', 'from', 'to'])
        if (v[name]) body += el('p:' + name, { x: integer(v[name][0], kind === 'animScale' ? 1000 : 100000),
          y: integer(v[name][1], kind === 'animScale' ? 1000 : 100000) });
      if (kind === 'animMotion' && flags & 64 && props.has(Q.pathCenterX) && props.has(Q.pathCenterY))
        body += el('p:rCtr', { x: integer(props.get(Q.pathCenterX), 100000),
          y: integer(props.get(Q.pathCenterY), 100000) });
    } else if (kind === 'animRot') {
      for (const name of ['by', 'from', 'to']) if (v[name] != null) attrs[name] = integer(v[name], 60000);
    } else if (kind === 'animClr') {
      attrs.clrSpc = flags & 8 && props.get(Q.colorModel) ? 'hsl' : 'rgb';
      attrs.dir = flags & 16 && props.get(Q.colorDirection) ? 'ccw' : 'cw';
      for (const name of ['by', 'from', 'to'])
        if (v[name]) body += el('p:' + name, {}, color(v[name], context.page, name === 'by'));
    } else A.invalid('BEHAVIOR_TYPE');
    return el('p:' + kind, attrs, common + body);
  };
})(typeof window !== 'undefined' ? window : globalThis);
