/* Binary PowerPoint shape actions and UTF-16 text-range hyperlinks.
 * Derived from Apache POI hslf/usermodel/HSLFHyperlink.java (Apache-2.0).
 * Derived from Apache OpenOffice filter/source/msfilter/svdfppt.cxx and
 * sd/source/filter/ppt/pptin.cxx (Apache-2.0).
 * MS-PPT 2.6.10, 2.10.16–24 and MS-OSHARED 2.3.3.1.21:
 * targets, property-set hyperlink tables, locations and ScreenTips. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    P = B.ppt,
    D = P.drawings,
    T = P.text,
    el = B.el,
    H = (P.links = {});
  const string = r => r.r.text(0, r.r.length & ~1, 1200).replace(/\0+$/, '');
  const jumps = ['', 'nextslide', 'previousslide', 'firstslide', 'lastslide', 'lastslideviewed', 'endshow'];
  const bad = reason => B.fail('PPT_HYPERLINK', reason);
  H.propertyTargets = function (cfb) {
    const links = [];
    if (!cfb) return links;
    try {
      const sets = B.properties(cfb.stream('\x05DocumentSummaryInformation'), { kind: 'documentSummary' });
      for (const set of sets) {
        const id = Object.keys(set.dictionary).find(id => set.dictionary[id] === '_PID_HLINKS'),
          prop = set.values[id];
        if (prop?.type !== 65 || !Array.isArray(prop.value)) continue;
        const r = new B.Reader(Uint8Array.from(prop.value), 'Presentation hyperlink properties'),
          count = r.u32(0);
        if (count % 6 || count > 100000) bad('Invalid hyperlink property count');
        let at = 4;
        const value = () => {
          const type = r.u32(at);
          at += 4;
          if (type === 3) {
            const value = r.i32(at);
            at += 4;
            return value;
          }
          if (type !== 30 && type !== 31) bad('Invalid hyperlink property value');
          const n = r.u32(at),
            bytes = n * (type === 31 ? 2 : 1);
          at += 4;
          if (bytes > B.limits.text) bad('Hyperlink property text is too large');
          const text = r.text(at, bytes, type === 31 ? 1200 : set.codepage).replace(/\0+$/, '');
          at = Math.ceil((at + bytes) / 4) * 4;
          return text;
        };
        for (let i = 0; i < count / 6; i++) {
          for (let j = 0; j < 4; j++)
            if (typeof value() !== 'number') bad('Invalid hyperlink property header');
          const target = value(),
            location = value();
          if (typeof target !== 'string' || typeof location !== 'string')
            bad('Invalid hyperlink property target');
          links.push({ target, location });
        }
        break;
      }
    } catch (e) {
      if (!['LegacyError', 'CFBError'].includes(e.name)) throw e;
    }
    // Preserve a valid prefix after optional damage; each unresolved used link
    // reports its own loss. ExHyperlink records provide the IDs in list order.
    return links;
  };
  H.prepare = function (book, pkg, cfb) {
    const definitions = new Map(),
      extras = new Map(),
      slides = new Map(book.slides.map((s, i) => [s.id, 'ppt/slides/slide' + (i + 1) + '.xml'])),
      relationships = new Map(),
      written = new Set();
    const records = book.documentRecords
      .filter(r => r.type === 1033)
      .flatMap(rec => P.children(rec, () => {}).filter(r => r.type === 4055));
    const targets = records.length ? H.propertyTargets(cfb) : [];
    let targetIndex = 0;
    // Definitions are optional until used. One damaged entry cannot disable
    // another link or refuse the presentation.
    const definition = rec => {
      const target = targets[targetIndex++];
      try {
        const children = P.children(rec, () => {}),
          id = children.find(r => r.type === 4051)?.r.u32(0);
        if (id == null) return;
        const data = {};
        for (const r of children) if (r.type === 4026) data[r.instance] = string(r);
        if (target) {
          data[1] = target.target;
          data[3] = target.location;
        }
        definitions.set(id, data);
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
      }
    };
    for (const rec of records) definition(rec);
    for (const rec of T.extensions(book.documentRecords, () => {}))
      if (rec.type === 4068)
        try {
          const children = P.children(rec, () => {}),
            id = children.find(r => r.type === 4051)?.r.u32(0),
            tip = children.find(r => r.type === 4026);
          if (id != null) extras.set(id, { tooltip: tip ? string(tip) : '' });
        } catch (e) {
          if (e.name !== 'LegacyError') throw e;
        }
    const resolve = (rec, page, s) => {
      const atom = P.children(rec, e => D.report(book, page, 'link', s, e)).find(r => r.type === 4083)?.r;
      if (!atom) return null;
      atom.check(0, 16);
      const id = atom.u32(4),
        action = atom.u8(8),
        jump = atom.u8(10),
        flags = atom.u8(11),
        type = atom.u8(12),
        definition = definitions.get(id),
        info = {
          kind: rec.instance === 1 ? 'over' : 'click',
          attrs: {
            tooltip: extras.get(id)?.tooltip || D.string(s, 909) || null,
            highlightClick: flags & 1 ? 1 : null,
            endSnd: !atom.u32(0) && flags & 2 ? 1 : null
          }
        };
      if (action || atom.u32(0) || flags & 2) book.drawingStats.sourceHyperlinks++;
      if (atom.u32(0)) {
        if (P.media) info.sound = atom.u32(0);
        else D.report(book, page, 'sound', s, 'Action sound');
      }
      if (action === 0) return flags & 2 || info.sound ? info : null;
      if (action === 6 && s.media) {
        info.attrs.action = 'ppaction://media';
        return info;
      }
      if (action === 3) {
        if (!jumps[jump]) bad('Unknown slide jump');
        info.attrs.action = 'ppaction://hlinkshowjump?jump=' + jumps[jump];
        return info;
      }
      if (action !== 4) bad('Unsupported action ' + action);
      if (type <= 3) {
        info.attrs.action = 'ppaction://hlinkshowjump?jump=' + jumps[type + 1];
        return info;
      }
      if (!definition) bad('Missing hyperlink ' + id);
      const location = definition[3] || '';
      if (type === 7) {
        // Location uses source slide ID, slide number, title. POI also writes
        // this tuple in Target. Never treat the persist reference as a slide.
        const address = location || definition[1] || '',
          match = /^(-?\d+),(-?\d+),/.exec(address);
        if (!match || !slides.has(+match[1])) bad('Missing hyperlink slide ' + address);
        info.target = slides.get(+match[1]);
        info.type = 'slide';
        info.attrs.action = 'ppaction://hlinksldjump';
        return info;
      }
      if (![8, 9, 10].includes(type)) bad('Unsupported hyperlink type ' + type);
      let target = definition[1] || '';
      if (!target) bad('Missing hyperlink destination');
      target = target.replace(/\\/g, '/');
      if (/^[a-z]:\//i.test(target)) target = 'file:///' + target;
      else if (target.startsWith('//')) target = 'file:' + target;
      if (location) target += '#' + location;
      if (/[\x00-\x1f]/.test(target) || /^(?:javascript|vbscript|data):/i.test(target.trim()))
        bad('Unsupported hyperlink destination');
      info.target = target.replace(/%(?![0-9a-f]{2})/gi, '%25').replace(/ /g, '%20');
      info.type = 'hyperlink';
      info.external = true;
      if (type === 9) info.attrs.action = 'ppaction://hlinkpres';
      else if (type === 10) info.attrs.action = 'ppaction://hlinkfile';
      return info;
    };
    const parse = (records, page, s, text) => {
      const result = [];
      for (let i = 0; i < records.length; i++) {
        const rec = records[i];
        if (rec.type !== 4082) continue;
        try {
          const info = resolve(rec, page, s);
          if (!info) continue;
          if (text != null) {
            const range = records[i + 1];
            if (range?.type !== 4063) bad('Missing hyperlink text range');
            info.start = range.r.i32(0);
            info.end = range.r.i32(4);
            if (
              info.start < 0 ||
              info.end <= info.start ||
              info.start >= text.length ||
              info.end > text.length + 1
            )
              bad('Invalid hyperlink text range');
            info.end = Math.min(info.end, text.length);
            // A range may end at the implicit paragraph mark, but never split
            // a Unicode scalar when its two UTF-16 units span a boundary.
            for (const key of ['start', 'end'])
              if (
                info[key] > 0 &&
                info[key] < text.length &&
                /[\ud800-\udbff]/.test(text[info[key] - 1]) &&
                /[\udc00-\udfff]/.test(text[info[key]])
              )
                info[key]++;
          }
          result.push(info);
        } catch (e) {
          if (e.name !== 'LegacyError') throw e;
          D.report(book, page, 'link', s, e);
        }
      }
      return result;
    };
    book.drawingStats.sourceHyperlinks = 0;
    for (const page of D.pages(book)) {
      const walk = shapes => {
        for (const s of shapes) {
          const client = s.records?.find(r => r.type === 0xf011);
          s.shapeLinks = client
            ? parse(
                P.children(client, e => D.report(book, page, 'link', s, e)),
                page,
                s
              )
            : [];
          s.textLinks = { click: [], over: [] };
          for (const info of parse(s.text?.interactiveRecords || [], page, s, s.text?.text))
            s.textLinks[info.kind].push(info);
          for (const kind of ['click', 'over']) {
            const list = s.textLinks[kind];
            list.sort((a, b) => a.start - b.start);
            let end = -1;
            s.textLinks[kind] = list.filter(info => {
              if (info.start < end) {
                D.report(book, page, 'link', s, 'Overlapping text hyperlinks');
                return false;
              }
              end = info.end;
              return true;
            });
          }
          walk(s.children || []);
        }
      };
      walk(page.textShapes || []);
    }
    book.linkXml = (info, page, s) => {
      if (!info) return '';
      try {
        const attrs = { ...info.attrs, 'r:id': '' };
        if (info.target) {
          const key = book.textPart + '|' + info.type + '|' + info.target;
          let rid = relationships.get(key);
          if (!rid) {
            rid = pkg.rel(
              book.textPart,
              info.type,
              info.external ? info.target : D.relative(book.textPart, info.target),
              { external: !!info.external }
            );
            relationships.set(key, rid);
          }
          attrs['r:id'] = rid;
        }
        if (!written.has(info)) {
          written.add(info);
          book.drawingStats.hyperlinks++;
        }
        const sound = info.sound ? P.media?.sound(book, info.sound) : '';
        if (info.sound && !sound) P.media?.report(book, page, s, 'Missing action sound');
        return el(info.kind === 'over' ? 'a:hlinkMouseOver' : 'a:hlinkClick', attrs, sound || undefined);
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        D.report(book, page, 'link', s, e);
        return '';
      }
    };
  };
  H.shape = (book, page, s) =>
    ['click', 'over']
      .map(
        kind =>
          book.linkXml?.(
            s.shapeLinks?.find(h => h.kind === kind),
            page,
            s
          ) || ''
      )
      .join('');
  H.cursor = function (book, page, s) {
    const offsets = { click: 0, over: 0 };
    return pos => {
      let xml = '',
        next = Infinity;
      for (const kind of ['click', 'over']) {
        const list = s.textLinks?.[kind] || [];
        while (offsets[kind] < list.length && list[offsets[kind]].end <= pos) offsets[kind]++;
        const info = list[offsets[kind]];
        if (!info) continue;
        if (info.start <= pos) {
          xml += book.linkXml(info, page, s);
          next = Math.min(next, info.end);
        } else next = Math.min(next, info.start);
      }
      return { xml, next };
    };
  };
})(typeof window !== 'undefined' ? window : globalThis);
