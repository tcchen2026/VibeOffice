/* Ledger — preserve complete comment threads beside their displayed notes. */
(function (root) {
  'use strict';
  const L = root.L, K = L.opc, F = L.formula, M = L.model, T = (L.threads = {});
  const NS = 'http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments';
  const RT = 'http://schemas.microsoft.com/office/2017/10/relationships/';
  const XR = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';
  const kids = (el, name) => Array.from(el?.children || []).filter(e => !name || e.localName === name);
  const at = (el, key) => el?.getAttribute(key);
  const clone = value => JSON.parse(JSON.stringify(value));
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  const note = cm => ({ text: cm.text, author: cm.author, runs: cm.runs });
  const values = t => ({ id: t.id, parent: t.parent, personId: t.personId, date: t.date, text: t.text, done: t.done, mentions: t.mentions });
  const related = (pkg, part, kind) => pkg?.rels(part).find(r => r.type === RT + kind && !r.external);
  const attrNS = (el, a) => a.prefix ? el.lookupNamespaceURI(a.prefix) : '';
  function personKey(xml) {
    const content = e => [e.namespaceURI, e.localName,
      Array.from(e.attributes).filter(a => a.name !== 'xmlns' && a.prefix !== 'xmlns').map(a => [attrNS(e, a), a.localName, a.value]).sort(),
      e.childNodes.map(n => n.nodeType === 1 ? content(n) : n.data)];
    return JSON.stringify(content(K.parse(xml)));
  }
  function people(wb) {
    const rel = related(wb.pkg, wb.pkg?.main, 'person');
    return kids(rel && wb.pkg.xml(rel.part), 'person').map(p => ({ id: at(p, 'id'), xml: K.raw(p) }));
  }
  T.notes = function (sh, part) {
    const pkg = sh.wb.pkg; if (!pkg) return;
    for (const el of kids(kids(pkg.xml(part), 'commentList')[0], 'comment')) {
      const pos = F.parseCell(at(el, 'ref') || ''), cm = pos && sh.comments.get(M.key(pos.r, pos.c));
      if (cm) cm.keepNote = { fragment: K.fragment(el, { pkg, part }), values: clone(note(cm)) };
    }
  };
  T.read = function (sh, part, authors) {
    const pkg = sh.wb.pkg, tree = pkg?.xml(part); if (!tree) return;
    sh.extra.threadPart = part;
    const list = kids(tree, 'threadedComment'), byId = new Map(list.map(e => [at(e, 'id'), e]));
    const persons = people(sh.wb), touched = new Set();
    for (const el of list) {
      let ref = at(el, 'ref'), parent = el, seen = new Set();
      while (!ref && parent && !seen.has(parent)) { seen.add(parent); parent = byId.get(at(parent, 'parentId')); ref = at(parent, 'ref'); }
      const pos = F.parseCell(ref || ''); if (!pos) continue;
      const key = M.key(pos.r, pos.c), pid = at(el, 'personId');
      let cm = sh.comments.get(key);
      if (!cm) { cm = { r: pos.r, c: pos.c, text: '', author: authors?.get(pid) || '', visible: false }; sh.comments.set(key, cm); }
      if (!cm.threadKeep) cm.threadKeep = { source: pkg.id, part, synthetic: !cm.keepNote, people: [] };
      const mentions = kids(kids(el, 'mentions')[0], 'mention').map(e => Object.fromEntries(Array.from(e.attributes).filter(a => a.prefix !== 'xmlns' && a.name !== 'xmlns').map(a => [a.name, a.value])));
      const t = { author: authors?.get(pid) || '', text: kids(el, 'text')[0]?.textContent || '', date: at(el, 'dT') || '',
        id: at(el, 'id'), parent: at(el, 'parentId') || null, personId: pid, done: /^(1|true)$/.test(at(el, 'done')), mentions };
      t.keep = { fragment: K.fragment(el, { pkg, part }), values: clone(values(t)), ref: at(el, 'ref') };
      (cm.thread || (cm.thread = [])).push(t); touched.add(cm);
    }
    for (const cm of touched) {
      if (cm.threadKeep.synthetic) cm.text = cm.thread.map(t => t.text).join('\n\n');
      const ids = new Set(cm.thread.flatMap(t => [t.personId, ...t.mentions.map(m => m.mentionpersonId)]));
      cm.threadKeep.people = persons.filter(p => ids.has(p.id));
      cm.threadKeep.note = clone(note(cm));
    }
  };
  T.active = cm => cm.thread?.length && same(note(cm), cm.threadKeep?.note);
  const rootComment = cm => cm.thread?.find(t => !t.parent);
  const rootId = cm => rootComment(cm)?.id;
  T.editText = cm => T.active(cm) ? rootComment(cm)?.text ?? cm.text : cm.text;
  T.editNote = function (cm, rootText) {
    const first = cm && rootComment(cm), previous = cm?.threadKeep?.note;
    if (!first || !previous || rootText === undefined && cm.text === previous.text) return cm;
    const text = String(rootText ?? cm.text ?? ''), before = first.text;
    let start = 0, end = before.length, nextEnd = text.length;
    while (start < end && start < nextEnd && before[start] === text[start]) start++;
    while (end > start && nextEnd > start && before[end - 1] === text[nextEnd - 1]) { end--; nextEnd--; }
    const intact = before ? text.indexOf(before) : -1;
    // Mentions outside the replaced text keep their identities; those after it
    // move with their text. Replacing a mentioned name removes that mention only.
    const mentions = first.mentions.flatMap(m => {
      const at = Number(m.startIndex), length = Number(m.length);
      if (!Number.isInteger(at) || !Number.isInteger(length) || at < 0 || length < 1 || at + length > before.length) return [];
      if (intact >= 0 && text.indexOf(before, intact + 1) < 0) return [{ ...m, startIndex: String(at + intact) }];
      if (at + length <= start) return [{ ...m }];
      if (at >= end) return [{ ...m, startIndex: String(at + nextEnd - end) }];
      // Several edits can surround an untouched name. Keep its unique match;
      // ambiguous repeated names cannot safely acquire a mention by guessing.
      const name = before.slice(at, at + length), match = text.indexOf(name);
      if (before.indexOf(name) === at && before.indexOf(name, at + 1) < 0 && match >= 0 && text.indexOf(name, match + 1) < 0) return [{ ...m, startIndex: String(match) }];
      return [];
    });
    const edited = { ...cm, thread: cm.thread.map(t => t === first ? { ...t, text, mentions } : t), runs: undefined };
    edited.text = edited.thread.map(t => t.text).join('\n\n');
    edited.threadKeep = { ...cm.threadKeep, note: clone(note(edited)) };
    return edited;
  };
  T.author = cm => T.active(cm) && rootId(cm) ? 'tc=' + rootId(cm) : cm.thread?.length && /^tc=\{/.test(cm.author || '') ? cm.thread[0].author : cm.author;
  T.noteAttributes = cm => T.active(cm) && rootId(cm) ? ' xmlns:xr="' + XR + '" xr:uid="' + rootId(cm) + '"' : '';
  T.noteNamespaces = ' xmlns:mc="' + K.NS.mc + '" xmlns:xr="' + XR + '" mc:Ignorable="xr"';
  T.noteXML = function (cm, authorId, writer, owner) {
    if (!cm.keepNote || !same(note(cm), cm.keepNote.values)) return null;
    const f = cm.keepNote.fragment, tree = K.parse(f.xml), attrs = { ref: F.cellName(cm.r, cm.c), authorId };
    const uid = Array.from(tree.attributes).find(a => a.localName === 'uid' && attrNS(tree, a) === XR);
    if (T.active(cm)) { if (!uid) attrs['xmlns:xr'] = XR; attrs[uid?.name || 'xr:uid'] = rootId(cm); }
    return writer.emit(K.slice(f, K.attributes(f.xml, attrs)), owner);
  };
  // Copied definitions and parent/mention references change together. Person
  // IDs are shared unless the destination uses the same ID for another person.
  T.copy = function (cm, wb) {
    const definitions = (cm.thread || []).flatMap(t => [t.id, ...t.mentions.map(m => m.mentionId)]);
    const existing = new Map(people(wb).map(p => [p.id, p.xml]));
    for (const sh of wb.sheets) for (const c of sh.comments.values()) for (const p of c.threadKeep?.people || []) existing.set(p.id, p.xml);
    for (const p of cm.threadKeep?.people || []) if (existing.has(p.id) && personKey(existing.get(p.id)) !== personKey(p.xml)) definitions.push(p.id);
    return K.duplicate(cm, { guidValues: definitions, guidAttributes: ['uid'],
      guidFields: ['id', 'parent', 'personId', 'mentionId', 'mentionpersonId'] });
  };
  function commentXML(cm, t, writer, owner) {
    const f = t.keep.fragment, old = t.keep.values, attrs = {};
    const ref = F.cellName(cm.r, cm.c); if (t.keep.ref !== ref) attrs.ref = ref;
    for (const [key, attr] of [['id', 'id'], ['parent', 'parentId'], ['personId', 'personId'], ['date', 'dT'], ['done', 'done']])
      if (!same(t[key], old[key])) attrs[attr] = key === 'done' ? t.done ? '1' : '0' : t[key] || null;
    let xml = K.attributes(f.xml, attrs), replacements = {};
    if (t.text !== old.text) {
      replacements['{' + NS + '}text'] = '<text xmlns="' + NS + '" xml:space="preserve">' + L.xml.esc(t.text) + '</text>';
    }
    if (!same(t.mentions, old.mentions)) {
      const list = kids(K.parse(xml), 'mentions')[0], entries = kids(list, 'mention');
      const byId = new Map(entries.map(el => [at(el, 'mentionId'), el]));
      replacements['{' + NS + '}mentions'] = t.mentions.length ? K.mergeBag(list ? K.raw(list) : '<mentions xmlns="' + NS + '"/>', {
        ['{' + NS + '}mention']: t.mentions.map(m => K.attributes(byId.has(m.mentionId) ? K.raw(byId.get(m.mentionId)) : '<mention xmlns="' + NS + '"/>', m))
      }) : [];
      if (old.mentions.some(m => !t.mentions.some(n => n.mentionId === m.mentionId))) writer.loss({ id: 'thread-mentions:' + t.id, what: 'Some mentions could not be kept after the comment text was edited.', where: ref, action: 'drop' });
    }
    if (Object.keys(replacements).length) xml = K.mergeBag(xml, replacements);
    return writer.emit(K.slice(f, xml), owner);
  }
  T.write = function (wb, pack) {
    const w = pack.writer, pkg = wb.pkg, needed = new Map();
    for (const sh of wb.sheets) {
      const source = sh.extra.threadPart, originals = source && pkg?.xml(source), items = [];
      for (const cm of sh.comments.values()) {
        if (T.active(cm)) {
          for (const t of cm.thread) items.push({ cm, t });
          for (const p of cm.threadKeep.people) needed.set(p.id, p.xml);
        } else if (cm.thread?.length) w.loss({ id: 'thread-conversion:' + sh.id + ':' + rootId(cm), what: 'A comment thread you edited will be saved as an ordinary note. Excel will show it as a note, not as a conversation.', where: sh.name + '!' + F.cellName(cm.r, cm.c), place: sh.name + '!' + F.cellName(cm.r, cm.c), action: 'conversion', notify: true });
      }
      if (!items.length) { if (source) w.omit(source, 'The comments on this sheet were deleted or converted to ordinary notes.'); continue; }
      const old = kids(originals, 'threadedComment'), positions = new Map(old.map((e, index) => [at(e, 'id'), index]));
      items.sort((a, b) => (positions.get(a.t.id) ?? Infinity) - (positions.get(b.t.id) ?? Infinity));
      const target = source ? w.target(pkg, source) : w.name('xl/threadedComments', 'threadedComment', 'xml');
      const unchanged = source && old.length === items.length && items.every(({ cm, t }, j) => t.keep.fragment.source === pkg.id && t.keep.fragment.part === source && t.id === at(old[j], 'id') && same(values(t), t.keep.values) && F.cellName(cm.r, cm.c) === t.keep.ref);
      if (unchanged) w.carry(pkg, source);
      else {
        if (source) w.claim(source, 'merged', target);
        const xml = originals ? pkg.text(source) : '<ThreadedComments xmlns="' + NS + '"/>';
        w.put(target, K.hoistNamespaces(K.mergeBag(xml, { ['{' + NS + '}threadedComment']: items.map(({ cm, t }) => commentXML(cm, t, w, target)) })), pkg?.type(source) || 'application/vnd.ms-excel.threadedcomments+xml');
        if (source) w.carryRels(pkg, source, target);
      }
      const owner = pack.sheetParts.get(sh), rel = related(pkg, sh.extra.ooxmlPart, 'threadedComment');
      w.rels(owner).add(RT + 'threadedComment', K.relative(owner, target), false, rel?.id);
    }
    const rel = related(pkg, pkg?.main, 'person'), source = rel?.part, tree = source && pkg.xml(source);
    const existing = new Set(kids(tree, 'person').map(p => at(p, 'id'))), added = [...needed].filter(([id]) => !existing.has(id)).map(([, xml]) => xml);
    if (added.length) {
      const target = source ? w.target(pkg, source) : w.name('xl/persons', 'person', 'xml');
      if (source) w.claim(source, 'merged', target);
      const xml = source ? pkg.text(source) : '<personList xmlns="' + NS + '"/>';
      w.put(target, K.hoistNamespaces(K.mergeBag(xml, { __append: added })), pkg?.type(source) || 'application/vnd.ms-excel.person+xml');
      if (source) w.carryRels(pkg, source, target);
      const owner = pack.part('xl/workbook.xml');
      w.rels(owner).add(RT + 'person', K.relative(owner, target), false, rel?.id);
    }
  };
})(typeof window !== 'undefined' ? window : globalThis);
