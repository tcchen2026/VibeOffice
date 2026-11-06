/* Binary PowerPoint timing records.
 * Derived from Apache OpenOffice sd/source/filter/ppt/pptinanimations.cxx (Apache-2.0).
 * MS-PPT 2.8.15–2.8.28, 2.8.73–2.8.82: bounded time nodes, conditions and variants.
 * Source flags select meaningful fields; reserved fields do not reject a deck. */
(function (root) {
  'use strict';
  const B = root.L.legacy, P = B.ppt, N = s => s;
  const A = P.timing = {};
  const FLOAT_MAX = 3.4028234663852886e38;
  const R = A.records = {
    condition: 0xf125, node: 0xf127, trigger: 0xf128, modifier: 0xf129,
    behavior: 0xf12a, animate: 0xf12b, color: 0xf12c, effect: 0xf12d,
    motion: 0xf12e, rotation: 0xf12f, scale: 0xf130, set: 0xf131, command: 0xf132,
    behaviorAtom: 0xf133, animateAtom: 0xf134, colorAtom: 0xf135, effectAtom: 0xf136,
    motionAtom: 0xf137, rotationAtom: 0xf138, scaleAtom: 0xf139, setAtom: 0xf13a,
    commandAtom: 0xf13b, target: 0xf13c, properties: 0xf13d, names: 0xf13e,
    values: 0xf13f, iterate: 0xf140, sequence: 0xf141, variant: 0xf142,
    time: 0xf143, group: 0xf144, sub: 0xf145, shape: 0x2afb, slide: 0x2b01,
    buildList: 0x2b02, build: 0x2b03, paragraphBuild: 0x2b08,
    paragraphAtom: 0x2b09, level: 0x2b0a, oldAnimation: 4081, oldContainer: 4116
  };
  A.properties = {
    display: 2, masterRelation: 5, subType: 6, presetId: 9, presetSubtype: 10,
    presetClass: 11, afterEffect: 13, slideCount: 15, timeFilter: 16, eventFilter: 17,
    hideWhenStopped: 18, groupId: 19, nodeType: 20, placeholder: 21,
    volume: 22, mute: 23, fullScreen: 26
  };
  A.behaviorProperties = {
    context: 2, relativePath: 3, colorModel: 4, colorDirection: 5, override: 6,
    pathRotation: 7, pathCenterX: 8, pathCenterY: 9, pointsTypes: 10
  };
  A.report = function (book, page, reason, shape) {
    if (!page.animationLoss) {
      page.animationLoss = B.loss('legacy-slide-animation', N('Some animations could not be kept'),
        page.where, 'drop', page.where, true);
      page.animationLoss.detail = [];
      book.losses.push(page.animationLoss);
    }
    if (page.animationLoss.detail.length < 20)
      page.animationLoss.detail.push({ shape, reason: String(reason?.code || reason) });
  };
  A.invalid = code => B.fail('PPT_TIMING_' + code, 'Invalid presentation animation data');
  A.number = value => {
    if (!Number.isFinite(value) || Math.abs(value) > 1e12) A.invalid('NUMBER');
    return value;
  };
  A.variant = function (record) {
    const r = record.r, type = r.u8(0);
    let value;
    if (type === 0) value = !!r.u8(1);
    else if (type === 1) value = r.i32(1);
    else if (type === 2) value = A.number(r.f32(1));
    else if (type === 3) {
      if (r.length > B.limits.text) A.invalid('STRING_LIMIT');
      value = r.text(1, (r.length - 1) & ~1, 1200).split('\0')[0];
    } else A.invalid('VALUE_TYPE');
    return { type, value };
  };
  A.propertySet = function (record) {
    const result = new Map();
    for (const child of record ? P.children(record) : [])
      if (child.type === R.variant) result.set(child.instance, A.variant(child).value);
    return result;
  };
  A.target = function (record) {
    if (!record) return null;
    const children = P.children(record);
    if (children.some(c => c.type === R.slide)) return { kind: 'slide' };
    const r = children.find(c => c.type === R.shape)?.r;
    if (!r) return null;
    r.check(0, 20);
    return { kind: r.u32(4) === 2 ? 'sound' : 'shape', mode: r.u32(0), type: r.u32(4),
      id: r.u32(8), start: r.i32(12), end: r.i32(16) };
  };
  A.condition = function (record) {
    const children = P.children(record), r = children.find(c => c.type === R.trigger)?.r;
    if (!r) A.invalid('CONDITION');
    r.check(0, 16);
    return { list: record.instance, object: r.u32(0), event: r.u32(4), id: r.u32(8),
      delay: r.i32(12), target: A.target(children.find(c => c.type === R.target)) };
  };
  A.readNode = function (record, context, depth = 0) {
    if (depth > B.limits.depth || ++context.count > B.limits.records) A.invalid('LIMIT');
    const children = P.children(record), r = children.find(c => c.type === R.node)?.r;
    const node = { id: ++context.id, kind: 0, children: [], sub: [], conditions: [],
      modifiers: new Map(), properties: new Map() };
    try {
      if (!r) A.invalid('NODE');
      r.check(0, 32);
      const flags = r.u32(28);
      node.kind = flags & 8 ? r.u32(8) : 0;
      if (flags & 1) node.fill = r.u32(12);
      if (flags & 2) node.restart = r.u32(4);
      if (flags & 16) node.duration = r.i32(24);
    } catch (error) {
      if (error.name !== 'LegacyError') throw error;
      context.report(error);
      node.invalid = true;
    }
    for (const child of children) {
      try {
        const data = child.r;
        if (child.type === R.group || child.type === R.sub) {
          const list = child.type === R.sub ? node.sub : node.children;
          list.push(A.readNode(child, context, depth + 1));
        } else if (child.type === R.condition) node.conditions.push(A.condition(child));
        else if (child.type === R.properties) node.properties = A.propertySet(child);
        else if (child.type === R.modifier) {
          data.check(0, 8);
          // Office stores these values as IEEE single precision, including
          // repeat count. The auto-reverse field is the one integer flag.
          const key = data.u32(0);
          const value = data.f32(4);
          node.modifiers.set(key, key === 5 ? !!data.u32(4) :
            key < 2 && value === FLOAT_MAX ? -1 : A.number(value));
        } else if (child.type === R.iterate) {
          data.check(0, 20);
          const bits = data.u32(16);
          node.iterate = { interval: bits & 4 ? data.u32(0) : 0,
            type: bits & 2 ? data.u32(4) : 0, backwards: !!(bits & 1) && !data.u32(8),
            percentage: !!(bits & 8) && !!data.u32(12) };
        } else if (child.type === R.sequence) {
          data.check(0, 20);
          const bits = data.u32(16);
          node.sequence = { concurrent: bits & 1 ? data.u32(0) : 0,
            next: bits & 2 ? data.u32(4) : 0, previous: bits & 4 ? data.u32(8) : 0 };
        } else if (child.type === R.target) node.target = A.target(child);
        else if (child.type >= R.animate && child.type <= R.command)
          node.behavior = A.readBehavior(child);
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        context.report(error);
        if (child.type !== R.group && child.type !== R.sub) node.invalid = true;
      }
    }
    if (![0, 1, 3, 4].includes(node.kind) || (node.kind === 3 && !node.behavior)) node.invalid = true;
    return node;
  };
  A.readBuilds = function (record, context) {
    const out = [];
    for (const entry of record ? P.children(record) : []) {
      try {
        if (entry.type !== R.paragraphBuild) {
          context.report('Unsupported animation build ' + entry.type);
          continue;
        }
        const children = P.children(entry), r = children.find(c => c.type === R.build)?.r,
          p = children.find(c => c.type === R.paragraphAtom)?.r;
        if (!r || !p) A.invalid('BUILD');
        r.check(0, 16);
        p.check(0, 16);
        const build = { id: r.u32(4), shape: r.u32(8), expanded: !!r.u8(12), uiExpanded: !!r.u8(13),
          type: p.u32(0), level: p.u32(4), background: !!p.u8(8), reverse: !!p.u8(9),
          auto: !!p.u8(11), delay: p.u32(12), templates: [] };
        let level = 0;
        for (const child of children) {
          if (child.type === R.level) level = child.r.u32(0);
          else if (child.type === R.group) build.templates.push({ level, node: A.readNode(child, context) });
        }
        out.push(build);
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        context.report(error);
      }
    }
    return out;
  };
  A.read = function (book, page) {
    const context = { id: 0, count: 0, report: error => A.report(book, page, error) };
    const extension = P.text.extensions(page.children || [], context.report);
    const roots = [], buildRecords = [];
    for (const record of extension) {
      try {
        if (record.type === R.group) roots.push(A.readNode(record, context));
        if (record.type === R.buildList) buildRecords.push(record);
      } catch (error) {
        if (error.name !== 'LegacyError') throw error;
        context.report(error);
      }
    }
    // Template timing IDs follow the slide timeline's IDs, as in Office output.
    const builds = buildRecords.flatMap(record => A.readBuilds(record, context));
    return { roots, builds, count: context.count, nextId: context.id };
  };
})(typeof window !== 'undefined' ? window : globalThis);
