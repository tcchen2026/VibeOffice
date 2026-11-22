/* PresentationML timing trees, targets and paragraph builds.
 * Derived from Apache OpenOffice sd/source/filter/ppt/pptinanimations.cxx (Apache-2.0).
 * ECMA-376 19.5; native PowerPoint files supply ordering and omitted defaults. */
(function (root) {
  'use strict';
  const B = root.L.legacy, P = B.ppt, A = P.timing, Q = A.properties, el = B.el;
  const events = [null, 'onBegin', 'onEnd', 'begin', 'end', 'onClick', 'onDblClick',
    'onMouseOver', 'onMouseOut', 'onNext', 'onPrev', 'onStopAudio'];
  const roles = [null, 'clickEffect', 'withEffect', 'afterEffect', 'mainSeq', 'interactiveSeq',
    'clickPar', 'withGroup', 'afterGroup', 'tmRoot'];
  const classes = ['cust', 'entr', 'exit', 'emph', 'path', 'verb', 'mediacall'];
  const unsigned = value => {
    if (!Number.isInteger(value) || value < 0 || value > 0xffffffff) A.invalid('INTEGER');
    return value;
  };
  const duration = n => n < 0 ? 'indefinite' : Math.round(A.number(n));
  const percent = n => Math.round(A.number(n) * 100000);
  const paragraphs = text => {
    const ranges = [];
    let start = 0;
    for (let i = 0; i < text.length; i++)
      if (text[i] === '\r') {
        ranges.push({ start, end: i + 1 });
        start = i + 1;
      }
    ranges.push({ start, end: text.length + 1 });
    return ranges;
  };
  A.paragraphs = paragraphs;
  A.targetXml = function (target, context) {
    if (!target) return '';
    if (target.kind === 'slide') return el('p:tgtEl', {}, el('p:sldTgt'));
    if (target.kind === 'sound') return P.media?.target(context.book, target.id) || '';
    const found = context.page.timingShapes?.get(target.id);
    if (!found) return '';
    const { id, shape } = found;
    let body = '';
    if (target.mode === 6) body = el('p:bg');
    else if (target.mode === 2 || target.mode === 8) {
      const text = shape.text?.text;
      if (text == null) return '';
      const ranges = paragraphs(text);
      let start = target.start, end = target.end;
      if (target.mode === 8 || (start === -1 && end === -1)) {
        start = 0;
        end = text.length + 1;
      }
      if (start < 0 || end <= start || end > text.length + 1) return '';
      const first = ranges.findIndex(r => r.start === start), last = ranges.findIndex(r => r.end === end);
      body = el('p:txEl', {}, first >= 0 && last >= first
        ? el('p:pRg', { st: first, end: last })
        : el('p:charRg', { st: start, end: end - 1 }));
    } else if (![0, 3, 4].includes(target.mode)) return '';
    if ((target.mode === 3 || target.mode === 4) && !shape.media) return '';
    return el('p:tgtEl', {}, el('p:spTgt', { spid: id }, body || undefined));
  };
  const condition = (item, context, tag = 'p:cond') => {
    const attrs = { delay: duration(item.delay) };
    if (item.event) {
      if (!events[item.event]) A.invalid('EVENT');
      attrs.evt = events[item.event];
    }
    let body;
    if (item.object === 1) {
      body = A.targetXml(item.target, context);
      if (!body) A.invalid('EVENT_TARGET');
    } else if (item.object === 2) {
      if (!context.ids.has(item.id)) A.invalid('TIME_REFERENCE');
      body = el('p:tn', { val: item.id });
    } else if (item.object === 3) body = el('p:rtn', { val: 'all' });
    else if (item.object !== 0) A.invalid('EVENT_TARGET');
    return el(tag, attrs, body);
  };
  const conditionList = (node, type, tag, context) => {
    const items = node.conditions.filter(c => c.list === type);
    return items.length ? el(tag, {}, items.map(c => condition(c, context)).join('')) : '';
  };
  A.commonXml = function (node, context) {
    const props = node.properties, attrs = { id: node.id };
    if (node.duration != null) attrs.dur = duration(node.duration);
    if (node.restart) attrs.restart = [null, 'always', 'whenNotActive', 'never'][node.restart];
    if (node.fill) attrs.fill = [null, 'remove', 'freeze', 'hold', 'transition'][node.fill];
    const copy = (key, attr, transform = v => v) => {
      if (props.has(key)) {
        const value = transform(props.get(key));
        if (value == null) A.invalid('PROPERTY');
        attrs[attr] = value;
      }
    };
    copy(Q.presetId, 'presetID', unsigned);
    copy(Q.presetSubtype, 'presetSubtype', unsigned);
    copy(Q.presetClass, 'presetClass', n => classes[n]);
    copy(Q.groupId, 'grpId', unsigned);
    copy(Q.nodeType, 'nodeType', n => roles[n]);
    copy(Q.afterEffect, 'afterEffect', n => +!!n);
    copy(Q.display, 'display', n => n ? 0 : 1);
    copy(Q.placeholder, 'nodePh', n => +!!n);
    copy(Q.timeFilter, 'tmFilter');
    copy(Q.eventFilter, 'evtFilter');
    copy(Q.masterRelation, 'masterRel', n => n === 0 ? 'sameClick' : n === 2 ? 'nextClick' : 'lastClick');
    const modifiers = node.modifiers;
    if (modifiers.has(0)) attrs.repeatCount = modifiers.get(0) < 0 ? 'indefinite' :
      Math.round(modifiers.get(0) * 1000);
    if (modifiers.has(1)) attrs.repeatDur = duration(modifiers.get(1) * 1000);
    if (modifiers.has(2)) attrs.spd = percent(modifiers.get(2));
    if (modifiers.has(3)) attrs.accel = percent(modifiers.get(3));
    if (modifiers.has(4)) attrs.decel = percent(modifiers.get(4));
    if (modifiers.has(5)) attrs.autoRev = +modifiers.get(5);
    // The one-based condition instances in real PPT files and AOO are
    // begin/end/next/previous/endSync. MS-PPT's enum table is off by one.
    let body = conditionList(node, 1, 'p:stCondLst', context) + conditionList(node, 2, 'p:endCondLst', context);
    const sync = node.conditions.find(c => c.list === 5);
    if (sync) body += condition(sync, context, 'p:endSync');
    if (node.iterate) {
      const i = node.iterate, type = ['el', 'wd', 'lt'][i.type];
      if (!type) A.invalid('ITERATE');
      body += el('p:iterate', { type, backwards: i.backwards ? 1 : null },
        i.percentage ? el('p:tmPct', { val: i.interval * 100 }) : el('p:tmAbs', { val: i.interval }));
    }
    const children = context.checkOnly ? '' : node.children.map(n => A.nodeXml(n, context)).join('');
    if (children) body += el('p:childTnLst', {}, children);
    const sub = context.checkOnly ? '' : node.sub.map(n => A.nodeXml(n, context)).join('');
    if (sub) body += el('p:subTnLst', {}, sub);
    return el('p:cTn', attrs, body || undefined);
  };
  A.nodeXml = function (node, context) {
    if (!context.checkOnly && !context.ids.has(node.id)) return '';
    try {
      if (node.invalid) A.invalid('NODE');
      if (node.kind === 3) return A.behaviorXml(node, context);
      if (node.kind === 4) {
        const target = A.targetXml(node.target, context);
        if (!target) A.invalid('MEDIA_TARGET');
        const props = node.properties, attrs = {};
        if (props.has(Q.volume)) attrs.vol = percent(props.get(Q.volume));
        if (props.has(Q.mute)) attrs.mute = +!!props.get(Q.mute);
        if (props.has(Q.slideCount)) attrs.numSld = unsigned(props.get(Q.slideCount));
        if (props.has(Q.hideWhenStopped)) attrs.showWhenStopped = props.get(Q.hideWhenStopped) ? 0 : 1;
        const tag = node.target.mode === 4 ? 'p:video' : 'p:audio';
        const outer = {};
        const media = context.page.timingShapes?.get(node.target.id)?.shape.media;
        if (tag === 'p:audio' && media?.flags & 4) outer.isNarration = 1;
        if (tag === 'p:video' && props.has(Q.fullScreen)) outer.fullScrn = +!!props.get(Q.fullScreen);
        return el(tag, outer, el('p:cMediaNode', attrs, A.commonXml(node, context) + target));
      }
      const common = A.commonXml(node, context);
      if (!context.checkOnly && !common.includes('<p:childTnLst>') && !common.includes('<p:subTnLst>')) return '';
      if (node.kind === 0) return el('p:par', {}, common);
      if (node.kind === 1) {
        const seq = node.sequence || {}, attrs = {};
        if (seq.concurrent != null) attrs.concurrent = +!!seq.concurrent;
        if (seq.next != null) attrs.nextAc = seq.next ? 'seek' : 'none';
        if (seq.previous != null && seq.previous) attrs.prevAc = 'skipTimed';
        return el('p:seq', attrs, common + conditionList(node, 4, 'p:prevCondLst', context) +
          conditionList(node, 3, 'p:nextCondLst', context));
      }
      A.invalid('NODE_TYPE');
    } catch (error) {
      if (error.name !== 'LegacyError') throw error;
      A.report(context.book, context.page, error, node.behavior?.target?.id || node.target?.id);
      return '';
    }
  };
  A.prepare = function (book) {
    for (const page of [...book.masters, ...book.slides]) {
      try {
        page.animation = A.read(book, page);
        if (!page.animation.roots.length && A.legacy) page.animation = A.legacy(book, page);
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        A.report(book, page, error);
      }
    }
  };
  A.write = function (book, page) {
    const animation = page.animation;
    if (!animation) return '';
    const context = { book, page, ids: new Set(), checkOnly: true };
    const entries = new Map(), references = new Map(), remove = [];
    const collect = (nodes, parent = null) => {
      for (const node of nodes) {
        context.ids.add(node.id);
        entries.set(node.id, { node, parent, children: node.children.length + node.sub.length });
        for (const condition of node.conditions) if (condition.object === 2) {
          if (!references.has(condition.id)) references.set(condition.id, []);
          references.get(condition.id).push(node.id);
        }
        collect(node.children, node.id);
        collect(node.sub, node.id);
      }
    };
    collect(animation.roots);
    for (const build of animation.builds) for (const template of build.templates) collect([template.node]);
    for (const [id, entry] of entries) {
      if (!A.nodeXml(entry.node, context) || (entry.node.kind < 2 && !entry.children)) remove.push(id);
    }
    // Remove dependents and empty containers in one queue. A lost target must
    // never leave a dangling time reference or silently become an immediate effect.
    for (let at = 0; at < remove.length; at++) {
      const id = remove[at];
      if (!context.ids.delete(id)) continue;
      const entry = entries.get(id), node = entry.node;
      for (const child of [...node.children, ...node.sub]) remove.push(child.id);
      for (const dependent of references.get(id) || []) {
        if (context.ids.has(dependent)) A.report(book, page, 'PPT_TIMING_TIME_REFERENCE');
        remove.push(dependent);
      }
      const parent = entries.get(entry.parent);
      if (parent && --parent.children === 0 && parent.node.kind < 2) remove.push(parent.node.id);
    }
    context.checkOnly = false;
    const nodes = animation.roots.map(node => A.nodeXml(node, context)).join('');
    const builds = A.buildXml?.(animation.builds, context) || '';
    return nodes || builds ? el('p:timing', {}, (nodes ? el('p:tnLst', {}, nodes) : '') + builds) : '';
  };
})(typeof window !== 'undefined' ? window : globalThis);
