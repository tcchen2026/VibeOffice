/* Derived from Apache OpenOffice main/filter/source/msfilter/svdfppt.cxx (Apache-2.0).
 * Derived from Apache POI hslf/record/{DocumentAtom,SlideAtom,NotesAtom,FontEntityAtom,
 * ColorSchemeAtom,SSSlideInfoAtom}.java (Apache-2.0).
 * MS-PPT 2.4–2.5: source identities, master inheritance and page associations.
 * The record byte views are retained for text and drawing conversion. */
(function (root) {
  'use strict';
  const B = root.L.legacy,
    P = B.ppt,
    N = s => s;
  P.defaultScheme = ['FFFFFF', '000000', '808080', '000000', 'BBE0E3', '333399', '009999', '99CC00'];
  P.notice = (losses, id, what, where, notify = false, detail) => {
    const loss = B.loss(id, what, where, 'drop', where, notify);
    if (detail) loss.detail = detail;
    losses.push(loss);
  };
  const problem = (losses, where, e) =>
    P.notice(
      losses,
      'legacy-presentation-record',
      N('Damaged presentation content could not be read'),
      where,
      true,
      e.code
    );
  P.structure = function (cfb, losses = []) {
    const book = P.read(cfb, { content: false, recover: true });
    if (book.encrypted) return book;
    const dc = P.children(book.document),
      atom = dc.find(x => x.type === 1001)?.r;
    if (!atom) B.fail('PPT_DOCUMENT', N('The presentation size and settings are missing'));
    atom.check(0, 40);
    const dimension = (n, where) => {
      const emu = Math.round(n * 1587.5);
      if (emu < 914400 || emu > 51206400)
        B.fail('PPT_SIZE', N('The presentation page size is outside the supported range'), where);
      return emu;
    };
    book.size = { cx: dimension(atom.i32(0), 'Slides'), cy: dimension(atom.i32(4), 'Slides') };
    book.notesSize = { cx: dimension(atom.i32(8), 'Notes'), cy: dimension(atom.i32(12), 'Notes') };
    book.firstSlideNumber = atom.u16(32);
    book.sizeType = atom.u16(34);
    book.rtl = !!atom.u8(38);
    book.omitTitlePlaceholders = !!atom.u8(37);
    book.documentRecords = dc;
    book.fonts = new Map();
    book.losses = losses;
    if (book.slides.length + book.masters.length + book.notes.length > 20000)
      B.fail('PPT_LIMIT', N('The presentation has too many slides and masters'));
    const safe = (fn, where, fallback) => {
      try {
        return fn();
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        problem(losses, where, e);
        return fallback;
      }
    };
    let scanned = 0,
      textBytes = 0;
    const content = item => {
      const found = { text: [], shapes: 0, pictures: 0, tables: 0, animations: 0, links: 0 };
      const visit = node => {
        if (++scanned > B.limits.records) B.fail('RECORD_LIMIT', N('Too many presentation records'));
        if (node.type === 0xf004) {
          const children = P.children(node, e => problem(losses, item.where, e)),
            sp = children.find(r => r.type === 0xf00a);
          if (sp && sp.r.length >= 8 && sp.r.u32(4) & 0x400) {
            const properties = new Map();
            for (const r of children.filter(r => [0xf00b, 0xf121, 0xf122].includes(r.type)))
              for (const [key, value] of B.officeart.properties(r)) properties.set(key, value);
            item.backgroundShape = { properties };
          }
        }
        if (node.type === 4000 || node.type === 4008) {
          textBytes += node.r.length;
          if (textBytes > B.limits.text) B.fail('TEXT_LIMIT', N('Presentation text is too large'));
          found.text.push(node.type === 4000 ? node.r.text(0, node.r.length, 1200) : B.latin16(node.r.bytes));
        }
        if (node.type === 0xf00a && node.r.length >= 8) {
          const flags = node.r.u32(4);
          if (!(flags & 1) && !(flags & 0x400)) found.shapes++;
        }
        if ([0xf00b, 0xf121, 0xf122].includes(node.type) && node.instance * 6 <= node.r.length)
          for (let i = 0; i < node.instance; i++) {
            const id = node.r.u16(i * 6) & 0x3fff,
              val = node.r.u32(i * 6 + 2);
            if (id === 260 && val) found.pictures++;
            if (id === 927 && val & 1) found.tables++;
          }
        if (node.type === 4081 || node.type === 4116) found.animations++;
        if (node.type === 4051) found.links++;
      };
      for (const rec of [...item.listRecords, ...(item.record ? [item.record] : [])])
        P.walk(rec, visit, e => problem(losses, item.where, e));
      return found;
    };
    const page = (item, kind, index) => {
      item.kind = kind;
      item.where =
        kind === 'slide'
          ? 'Slide ' + (index + 1)
          : kind === 'notes'
            ? 'Notes ' + (index + 1)
            : kind === 'master'
              ? 'Slide master ' + (index + 1)
              : kind === 'notesMaster'
                ? 'Notes master'
                : 'Handout master';
      item.children = item.record ? P.children(item.record, e => problem(losses, item.where, e)) : [];
      item.layout = 16;
      item.placeholders = [];
      item.masterId = 0;
      item.notesId = 0;
      item.slideId = 0;
      item.flags = 0;
      safe(() => {
        const a = item.children.find(x => [1007, 1009].includes(x.type));
        if (!a) return;
        const r = a.r;
        if (a.type === 1007) {
          item.layout = r.u32(0);
          item.placeholders = [...r.sub(4, 8)];
          item.masterId = r.u32(12);
          item.notesId = r.u32(16);
          item.flags = r.u16(20);
        } else {
          item.slideId = r.u32(0);
          item.flags = r.u16(4);
        }
      }, item.where);
      const schemes = item.children.filter(x => x.type === 2032),
        selected = schemes.find(x => x.instance === 1);
      // Instance 6 is the gallery, not the current scheme (instance 1).
      item.localScheme = selected
        ? safe(
            () =>
              Array.from({ length: 8 }, (_, i) =>
                [...selected.r.sub(i * 4, 3)]
                  .map(v => v.toString(16).padStart(2, '0'))
                  .join('')
                  .toUpperCase()
              ),
            item.where,
            null
          )
        : null;
      item.schemeGallery = schemes
        .filter(x => x.instance === 6 && x.r.length >= 32)
        .map(x =>
          Array.from({ length: 8 }, (_, i) =>
            [...x.r.sub(i * 4, 3)]
              .map(v => v.toString(16).padStart(2, '0'))
              .join('')
              .toUpperCase()
          )
        );
      const name = item.children.find(x => x.type === 4026 && x.instance === 3);
      item.name = name
        ? safe(() => name.r.text(0, name.r.length, 1200).replace(/\0+$/, ''), item.where, '')
        : '';
      const show = item.children.find(x => x.type === 1017)?.r;
      item.hidden = !!(show && show.length >= 12 && show.u16(10) & 4);
      item.transition = !!(show && show.length >= 12 && (show.u8(9) || show.u16(10) & 0x550));
      item.content = safe(() => content(item), item.where, {
        text: [],
        shapes: 0,
        pictures: 0,
        tables: 0,
        animations: 0,
        links: 0
      });
      item.text = item.content.text;
      return item;
    };
    book.slides.forEach((s, i) => page(s, 'slide', i));
    book.masters.forEach((s, i) => page(s, 'master', i));
    book.notes.forEach((s, i) => page(s, 'notes', i));
    const byRef = [book.slides, book.masters, book.notes].map(
      items => new Map(items.map(item => [item.ref, item]))
    );
    for (const issue of book.issues) {
      const item = byRef[issue.list].get(issue.ref);
      problem(losses, item?.where || 'Presentation', issue);
    }
    const special = (ref, kind) =>
      ref
        ? safe(() => page({ ref, id: 0, record: book.get(ref), listRecords: [] }, kind, 0), kind, null)
        : null;
    book.notesMaster = special(atom.u32(24), 'notesMaster');
    book.handoutMaster = special(atom.u32(28), 'handoutMaster');
    const masters = new Map(),
      notes = new Map();
    for (const m of book.masters) {
      if (masters.has(m.id))
        P.notice(losses, 'legacy-master-identity', 'A duplicate master identifier was reassigned', m.where);
      else masters.set(m.id, m);
    }
    for (const n of book.notes) if (!notes.has(n.id)) notes.set(n.id, n);
    for (const m of book.masters) m.parent = masters.get(m.masterId) || null;
    const resolve = (page, bit, field) => {
      let current = page;
      const seen = new Set();
      while (current.flags & bit && current.parent) {
        if (seen.has(current) || seen.size >= B.limits.depth) {
          P.notice(
            losses,
            'legacy-master-cycle',
            N('A damaged master relationship was not kept'),
            page.where,
            true
          );
          return page[field] || P.defaultScheme;
        }
        seen.add(current);
        current = current.parent;
      }
      return current[field] || P.defaultScheme;
    };
    for (const s of book.slides) {
      s.parent = masters.get(s.masterId) || null;
      s.notes = notes.get(s.notesId) || null;
      if (s.notes) s.notes.slide = s;
    }
    // NotesAtom supplies a slide ID, not a persist ID. Recover a one-way link,
    // but never attach an already-associated notes page to a different slide.
    const slidesById = new Map(book.slides.map(s => [s.id, s]));
    for (const n of book.notes) {
      const s = slidesById.get(n.slideId);
      if (!n.slide && s && !s.notes) {
        n.slide = s;
        s.notes = n;
      }
      n.parent = book.notesMaster;
    }
    for (const page of [
      ...book.masters,
      ...book.slides,
      ...book.notes,
      book.notesMaster,
      book.handoutMaster
    ].filter(Boolean)) {
      page.scheme = resolve(page, 2, 'localScheme');
      page.backgroundScheme = resolve(page, 4, 'localScheme');
    }
    const env = dc.find(x => x.type === 1010),
      fonts = env
        ? P.children(env, e => problem(losses, 'Presentation fonts', e)).find(x => x.type === 2005)
        : null;
    if (fonts)
      for (const font of P.children(fonts, e => problem(losses, 'Presentation fonts', e))) {
        if (font.type === 4023)
          safe(() => {
            const r = font.r;
            r.check(0, 68);
            book.fonts.set(font.instance, {
              id: font.instance,
              name: r.text(0, 64, 1200).split('\0')[0],
              charset: r.u8(64),
              flags: r.u8(65),
              type: r.u8(66),
              pitchFamily: r.u8(67)
            });
          }, 'Presentation fonts');
        else if (font.type === 4024) book.embeddedFonts = true;
      }
    return book;
  };
})(typeof window !== 'undefined' ? window : globalThis);
