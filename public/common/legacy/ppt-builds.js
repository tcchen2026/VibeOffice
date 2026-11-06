/* PowerPoint 97 animation builds and paragraph build lists.
 * Derived from Apache OpenOffice sd/source/filter/ppt/ppt97animations.cxx (Apache-2.0).
 * Derived from Apache OpenOffice sd/xml/effects.xml (Apache-2.0).
 * Derived from Apache POI hslf/record/AnimationInfoAtom.java (Apache-2.0).
 * MS-PPT 2.8.2–2.8.10. Effects use PresentationML behaviors, not a preview image. */
(function (root) {
  'use strict';
  const B = root.L.legacy, P = B.ppt, A = P.timing, R = A.records, Q = A.properties, el = B.el;
  const CLIENT_DATA = 0xf011;
  const F = { reverse: 1, automatic: 4, sound: 16, stopSound: 64, play: 256,
    synchronous: 1024, hidden: 4096, background: 16384 };
  const delay = ms => ({ list: 1, object: 0, event: 0, id: 0, delay: ms, target: null });
  const node = (kind = 0) => ({ id: 0, kind, children: [], sub: [], conditions: [],
    properties: new Map(), modifiers: new Map(), fill: 3 });
  const value = v => ({ type: typeof v === 'string' ? 3 : typeof v === 'boolean' ? 0 : 2, value: v });
  const behavior = (kind, target, names = []) => ({ kind, target, names, properties: new Map(),
    values: new Map(), keys: [], flags: 0 });
  const set = (target, attribute, to) => {
    const n = node(3);
    n.duration = 1;
    n.conditions = [delay(0)];
    n.behavior = behavior('set', target, [attribute]);
    n.behavior.flags = 1;
    n.behavior.values.set(1, value(to));
    return n;
  };
  const animate = (target, attribute, from, to, ms = 500, formula = '') => {
    const n = node(3), v = behavior('anim', target, [attribute]);
    n.duration = ms;
    v.flags = 56;
    v.additive = 0;
    v.calc = 1;
    v.valueType = 1;
    v.keys = [{ time: 0, value: value(from), formula }, { time: 1000, value: value(to) }];
    n.behavior = v;
    return n;
  };
  const filter = (target, name, ms = 500) => {
    const n = node(3);
    delete n.fill;
    n.duration = ms;
    n.behavior = behavior('animEffect', target);
    n.behavior.flags = 3;
    n.behavior.mode = 0;
    n.behavior.values.set(1, value(name));
    return n;
  };
  const command = (target, text, type = 1, ms = 1) => {
    const n = node(3);
    n.duration = ms;
    n.behavior = behavior('cmd', target);
    n.behavior.flags = 3;
    n.behavior.mode = type;
    n.behavior.values.set(1, value(text));
    return n;
  };
  A.oldEffect = function (old, target) {
    const n = node(), d = old.direction, effect = old.effect;
    n.conditions.push(delay(0));
    n.properties.set(Q.presetClass, 1);
    let preset = 1, subtype = 0;
    n.children.push(set(target, 'style.visibility', 'visible'));
    const addFilter = (id, sub, name) => {
      preset = id;
      subtype = sub;
      n.children.push(filter(target, name));
    };
    if (effect === 0) {
      // PowerPoint's binary-to-OOXML conversion gives the legacy Appear
      // effect 500 ms: its one-millisecond visibility set starts at 499 ms.
      n.children[0].conditions = [delay(499)];
    }
    else if (effect === 1) {
      preset = 24;
      n.children.push(filter(target, 'random'));
    } else if (effect === 2) addFilter(3, d ? 10 : 5, 'blinds(' + (d ? 'horizontal' : 'vertical') + ')');
    else if (effect === 3) addFilter(5, d ? 5 : 10, 'checkerboard(' + (d ? 'down' : 'across') + ')');
    else if (effect === 5) addFilter(9, 0, 'dissolve');
    else if (effect === 8) addFilter(14, d ? 5 : 10, 'randombar(' + (d ? 'vertical' : 'horizontal') + ')');
    else if (effect === 9 && d >= 4 && d <= 7)
      addFilter(18, [9, 3, 12, 6][d - 4], 'strips(' + ['upLeft', 'upRight', 'downLeft', 'downRight'][d - 4] + ')');
    else if (effect === 10 && d <= 3) addFilter(22, [2, 4, 8, 1][d], 'wipe(' + ['right', 'down', 'left', 'up'][d] + ')');
    else if (effect === 11) addFilter(4, d ? 16 : 32, 'box(' + (d ? 'in' : 'out') + ')');
    else if (effect === 13 && d <= 3)
      addFilter(16, [42, 26, 37, 21][d], 'barn(' + (d & 1 ? 'in' : 'out') +
        (d < 2 ? 'Horizontal' : 'Vertical') + ')');
    else if (effect === 14 && d <= 2) {
      preset = 11;
      n.children[0].duration = [75, 500, 1000][d];
      delete n.children[0].fill;
    } else if (effect === 12) {
      if (d < 8 || (d >= 12 && d < 16)) {
        const direction = d < 8 ? d : d - 12, ms = d >= 12 ? 5000 : 500;
        preset = d >= 12 ? 7 : 2;
        subtype = [8, 1, 2, 4, 9, 3, 12, 6][direction];
        const x = [0, 4, 6].includes(direction) ? '-#ppt_w/2' :
          [2, 5, 7].includes(direction) ? '1+#ppt_w/2' : '#ppt_x';
        const y = [1, 4, 5].includes(direction) ? '-#ppt_h/2' :
          [3, 6, 7].includes(direction) ? '1+#ppt_h/2' : '#ppt_y';
        n.children.push(animate(target, 'ppt_x', x, '#ppt_x', ms), animate(target, 'ppt_y', y, '#ppt_y', ms));
      } else if (d < 12) {
        const axis = d & 1 ? 'y' : 'x', size = d & 1 ? 'h' : 'w';
        const motion = animate(target, 'ppt_' + axis, '#ppt_' + axis +
          (d === 8 || d === 11 ? '-' : '+') + '#ppt_' + size + '*1.125000', '#ppt_' + axis);
        delete motion.fill;
        n.children.push(motion);
        addFilter(12, [8, 4, 2, 1][d - 8], 'wipe(' + ['right', 'up', 'left', 'down'][d - 8] + ')');
      } else if (d <= 21) {
        preset = 23;
        subtype = [16, 272, 32, 288, 528, 36][d - 16];
        const factor = ['0', '2/3', '4', '4/3', '0', '(6*min(max(#ppt_w*#ppt_h,.3),1)-7.4)/-.7'][d - 16];
        n.children.push(animate(target, 'ppt_w', factor + '*#ppt_w', '#ppt_w'),
          animate(target, 'ppt_h', factor + '*#ppt_h', '#ppt_h'));
        if (d >= 20) n.children.push(animate(target, 'ppt_x', '0.5', '#ppt_x'),
          animate(target, 'ppt_y', d === 20 ? '0.5' : '1+' + factor + '*#ppt_h/2', '#ppt_y'));
      } else if (d <= 26) {
        preset = 17;
        subtype = [10, 8, 1, 2, 4][d - 22];
        const horizontal = d === 22 || d === 23 || d === 25;
        if (d !== 22) {
          const sign = d === 23 || d === 24 ? '-' : '+';
          n.children.push(animate(target, 'ppt_x', horizontal ? '#ppt_x' + sign + '#ppt_w/2' : '#ppt_x', '#ppt_x'),
            animate(target, 'ppt_y', horizontal ? '#ppt_y' : '#ppt_y' + sign + '#ppt_h/2', '#ppt_y'));
        }
        n.children.push(animate(target, 'ppt_w', horizontal ? 0 : '#ppt_w', '#ppt_w'),
          animate(target, 'ppt_h', horizontal ? '#ppt_h' : 0, '#ppt_h'));
      } else if (d === 27) {
        preset = 19;
        subtype = 10;
        n.children.push(animate(target, 'ppt_w', '0', '1', 5000, '#ppt_w*sin(2.5*pi*$)'),
          animate(target, 'ppt_h', '#ppt_h', '#ppt_h', 5000));
      } else if (d === 28) {
        preset = 15;
        n.children.push(animate(target, 'ppt_w', '0', '#ppt_w', 1000),
          animate(target, 'ppt_h', '0', '#ppt_h', 1000),
          animate(target, 'ppt_x', '0', '1', 1000,
            '#ppt_x+(cos(-2*pi*(1-$))*-#ppt_x-sin(-2*pi*(1-$))*(1-#ppt_y))*(1-$)'),
          animate(target, 'ppt_y', '0', '1', 1000,
            '#ppt_y+(sin(-2*pi*(1-$))*-#ppt_x+cos(-2*pi*(1-$))*(1-#ppt_y))*(1-$)'));
      } else A.invalid('OLD_EFFECT');
    } else A.invalid('OLD_EFFECT');
    n.properties.set(Q.presetId, preset);
    n.properties.set(Q.presetSubtype, subtype);
    if (old.textMode) n.iterate = { type: old.textMode === 2 ? 2 : 1,
      interval: old.textMode === 2 ? 75 : 300, backwards: false, percentage: false };
    return n;
  };
  A.oldRecords = function (book, page) {
    const result = [];
    const visit = (shapes, depth = 0) => {
      if (depth > B.limits.depth) A.invalid('LIMIT');
      for (const shape of shapes) {
        try {
          const client = shape.records?.find(r => r.type === CLIENT_DATA);
          if (client) {
            const container = P.children(client).find(r => r.type === R.oldContainer);
            const r = container && P.children(container).find(r => r.type === R.oldAnimation)?.r;
            if (r) {
              r.check(0, 28);
              result.push({ shape, color: r.u32(0), flags: r.u32(4), sound: r.u32(8), delay: r.i32(12),
                order: r.i16(16), slideCount: r.u16(18), build: r.u8(20), effect: r.u8(21),
                direction: r.u8(22), after: r.u8(23), textMode: r.u8(24), verb: r.u8(25) });
            }
          }
        } catch (error) {
          if (error.name !== 'LegacyError') throw error;
          A.report(book, page, error, shape.id);
        }
        visit(shape.children || [], depth + 1);
      }
    };
    visit(page.textShapes || []);
    return result.sort((a, b) => a.order - b.order);
  };
  A.legacy = function (book, page) {
    const timing = node(), sequence = node(1), builds = [], effects = [];
    timing.duration = -1;
    timing.restart = 3;
    delete timing.fill;
    timing.properties.set(Q.nodeType, 9);
    sequence.duration = -1;
    delete sequence.fill;
    sequence.properties.set(Q.nodeType, 4);
    sequence.sequence = { concurrent: 1, next: 1 };
    sequence.conditions = [3, 4].map(list => ({ list, object: 1, event: list === 3 ? 9 : 10,
      id: 0, delay: 0, target: { kind: 'slide' } }));
    timing.children.push(sequence);
    for (const old of A.oldRecords(book, page)) {
      if (old.order === -2 || (!old.build && !old.shape.media)) continue;
      try {
        const shape = old.shape, text = shape.text?.text || '';
        const target = { kind: 'shape', type: 1, mode: 0, id: shape.id, start: -1, end: -1 };
        if (shape.media) {
          const media = node(4), video = shape.media.kind === 'video';
          media.target = { ...target, mode: video ? 4 : 3 };
          media.fill = shape.media.flags & 2 ? 1 : 3;
          media.conditions = [delay(-1)];
          media.properties = new Map([[Q.volume, 0.8], [Q.display, 1],
            [Q.hideWhenStopped, !!(old.flags & F.hidden)]]);
          if (old.slideCount > 0) media.properties.set(Q.slideCount, old.slideCount);
          if (shape.media.flags & 1) media.modifiers.set(0, -1);
          if (!video) media.conditions.push({ ...delay(0), list: 2, object: 1,
            event: 11, target: { kind: 'slide' } });
          timing.children.push(media);
          if (video) {
            const interactive = node(1), click = node(), group = node(), effect = node();
            interactive.restart = 2;
            interactive.sequence = { concurrent: 1, next: 1 };
            interactive.properties = new Map([[Q.nodeType, 5], [Q.eventFilter, 'cancelBubble']]);
            interactive.conditions = [1, 3].map(list => ({ ...delay(0), list, object: 1,
              event: 5, target }));
            interactive.conditions.push({ ...delay(0), list: 5, object: 3, event: 4 });
            click.conditions = group.conditions = effect.conditions = [delay(0)];
            click.properties.set(Q.nodeType, 6);
            group.properties.set(Q.nodeType, 7);
            effect.properties = new Map([[Q.presetId, 2], [Q.presetClass, 6],
              [Q.presetSubtype, 0], [Q.nodeType, 1]]);
            effect.children.push(command(media.target, 'togglePause'));
            group.children.push(effect);
            click.children.push(group);
            interactive.children.push(click);
            timing.children.push(interactive);
          }
          if (!(old.flags & F.play) && !old.build) continue;
        }
        const targets = [], group = 0;
        if (old.build > 1 && text) {
          const ranges = A.paragraphs(text);
          if (old.flags & F.reverse) ranges.reverse();
          if (old.flags & F.background) targets.push({ ...target, mode: 6, background: true });
          for (const range of ranges)
            if (text.slice(range.start, range.end).trim()) targets.push({ ...target, mode: 2, ...range });
          builds.push({ shape: shape.id, id: group, expanded: true, uiExpanded: false,
            type: 1, level: old.build - 1, background: !!(old.flags & F.background),
            reverse: !!(old.flags & F.reverse), auto: !!(old.flags & F.automatic),
            delay: Math.max(0, old.delay === 0x7fffffff ? 0 : old.delay), templates: [] });
        } else {
          targets.push(target);
          if (text && old.build) builds.push({ shape: shape.id, id: group, type: 3, background: true,
            templates: [] });
        }
        let previousTarget;
        for (const t of targets) {
          let effect;
          if (shape.media && old.flags & F.play) {
            if (old.flags & F.synchronous && !shape.media.duration)
              A.report(book, page, 'PPT_TIMING_MEDIA_DURATION', shape.id);
            effect = node();
            effect.conditions = [delay(0)];
            effect.properties = new Map([[Q.presetId, 1], [Q.presetClass, 6], [Q.presetSubtype, 0]]);
            effect.children.push(command({ ...t, mode: shape.media.kind === 'video' ? 4 : 3 },
              'playFrom(0.0)', 1, old.flags & F.synchronous ? Math.max(1, shape.media.duration || 1) : 1));
          } else effect = A.oldEffect(old, t);
          let role = old.flags & F.automatic ? 3 : 1;
          const format = shape.text?.paragraphs.find(p => p.start <= t.start && p.end > t.start)?.props;
          if (previousTarget?.background || (previousTarget && (format?.level || 0) >= old.build - 1))
            role = old.textMode ? 3 : 2;
          effect.properties.set(Q.nodeType, role);
          if (old.build && text) effect.properties.set(Q.groupId, group);
          effect.oldDelay = role === 3 ? Math.max(0, old.delay === 0x7fffffff ? 0 : old.delay) : 0;
          if (old.after) {
            let after;
            if (old.after === 1) {
              after = node(3);
              after.duration = 1;
              after.behavior = behavior('animClr', t, ['ppt_c']);
              after.behavior.flags = 4;
              const color = old.color, index = color >>> 24;
              after.behavior.to = index < 8 ? [2, index, 0, 0] :
                [0, color & 255, color >>> 8 & 255, color >>> 16 & 255];
            } else after = set(t, 'style.visibility', 'hidden');
            after.properties.set(Q.afterEffect, true);
            after.properties.set(Q.masterRelation, old.after === 3 ? 2 : 0);
            effect.sub.push(after);
          }
          if (old.flags & F.sound && old.sound) {
            const audio = node(4);
            audio.target = { kind: 'sound', id: old.sound };
            audio.properties.set(Q.display, 1);
            audio.properties.set(Q.masterRelation, 0);
            if (old.slideCount) audio.properties.set(Q.slideCount, old.slideCount);
            effect.sub.push(audio);
          }
          if (old.flags & F.stopSound) effect.children.unshift(command({ kind: 'slide' }, 'onstopaudio', 0));
          effects.push(effect);
          previousTarget = t;
        }
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        A.report(book, page, error, old.shape.id);
      }
    }
    let click, group, previous;
    for (const effect of effects) {
      const role = effect.properties.get(Q.nodeType);
      if (!click || role === 1) {
        click = node();
        click.properties.set(Q.nodeType, 6);
        click.conditions = [delay(role === 1 ? -1 : 0)];
        sequence.children.push(click);
        group = null;
      }
      if (!group || role !== 2) {
        group = node();
        group.properties.set(Q.nodeType, 7);
        group.conditions = [delay(effect.oldDelay)];
        if (role === 3 && previous) group.conditions[0] = {
          ...delay(effect.oldDelay), object: 2, event: 4, reference: previous
        };
        click.children.push(group);
      }
      group.children.push(effect);
      previous = effect;
    }
    let id = 0;
    const assign = n => {
      n.id = ++id;
      n.children.forEach(assign);
      n.sub.forEach(assign);
    };
    assign(timing);
    const references = n => {
      for (const c of n.conditions) if (c.reference) {
        c.id = c.reference.id;
        delete c.reference;
      }
      n.children.forEach(references);
      n.sub.forEach(references);
    };
    references(timing);
    return { roots: effects.length || timing.children.length > 1 ? [timing] : [], builds, count: id, nextId: id };
  };
  A.buildXml = function (builds, context) {
    const rows = [];
    for (const b of builds) {
      try {
        const shape = context.page.timingShapes?.get(b.shape);
        if (!shape || !shape.shape.text) A.invalid('BUILD_TARGET');
        const type = ['allAtOnce', 'p', 'cust', 'whole'][b.type];
        if (!type) A.invalid('BUILD_TYPE');
        const attrs = { spid: shape.id, grpId: b.id, uiExpand: b.uiExpanded ? 1 : null,
          build: b.type === 3 ? null : type, bldLvl: b.type === 1 ? b.level : null,
          animBg: b.background ? 1 : null, autoUpdateAnimBg: 0, rev: b.reverse ? 1 : null,
          advAuto: b.auto ? b.delay : null };
        const templates = b.templates.map(t => {
          const content = A.nodeXml(t.node, context);
          return content ? el('p:tmpl', { lvl: t.level }, el('p:tnLst', {}, content)) : '';
        }).join('');
        rows.push(el('p:bldP', attrs, templates ? el('p:tmplLst', {}, templates) : undefined));
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        A.report(context.book, context.page, error, b.shape);
      }
    }
    return rows.length ? el('p:bldLst', {}, rows.join('')) : '';
  };
})(typeof window !== 'undefined' ? window : globalThis);
