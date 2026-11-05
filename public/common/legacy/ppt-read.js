/* Derived from Apache POI poi-scratchpad/src/main/java/org/apache/poi/hslf/record/
 * CurrentUserAtom, UserEditAtom, PersistPtrHolder, SlidePersistAtom (Apache-2.0).
 * MS-PPT §§2.2–2.5: current persist records, not stale objects from incremental saves. */
(function (root) {
  'use strict';
  const N = s => s; // UI refusal text, translated where it is shown.
  const B = root.L.legacy, P = B.ppt || (B.ppt = {});
  P.record = function (r, offset, end = r.length) {
    r.check(offset, 8); const word = r.u16(offset), size = r.u32(offset + 4), next = offset + 8 + size;
    if (next > end) B.fail('PPT_RECORD', N('A presentation record extends beyond its container'), offset);
    r.check(offset + 8, size);
    return { type: r.u16(offset + 2), version: word & 15, instance: word >>> 4, offset, next, r: r.reader(offset + 8, size) };
  };
  P.children = function (rec, damaged) {
    const children = []; let p = 0;
    while (p < rec.r.length) {
      if (children.length >= B.limits.records) B.fail('RECORD_LIMIT', N('Too many presentation records'));
      let child; try { child = P.record(rec.r, p); }
      catch (e) { if (!damaged || e.name !== 'LegacyError') throw e; damaged(e); break; }
      children.push(child); p = child.next;
    }
    return children;
  };
  P.walk = function (rec, visit, damaged) {
    const pending = [{ rec, depth: 0 }]; let count = 0;
    while (pending.length) {
      const { rec: node, depth } = pending.pop();
      if (++count > B.limits.records || depth > B.limits.depth) B.fail('RECORD_LIMIT', N('Presentation record nesting exceeds the limit'));
      visit(node);
      if (node.version === 15 || [1036, 1035, 0xf00d, 0xf011].includes(node.type)) {
        const children = P.children(node, damaged); for (let i = children.length - 1; i >= 0; i--) pending.push({ rec: children[i], depth: depth + 1 });
      }
    }
  };
  P.read = function (cfb, options = {}) {
    const bytes = cfb.stream('PowerPoint Document'), userBytes = cfb.stream('Current User');
    if (!bytes) B.fail('PPT_STREAM', N('The PowerPoint Document stream is missing'));
    const r = new B.Reader(bytes, 'PowerPoint Document'), user = new B.Reader(userBytes || new Uint8Array(0), 'Current User');
    const encrypted = user.u32(12) === 0xf3d1c4df;
    if (encrypted) return { version: 'PowerPoint97+', encrypted, slides: [] };
    if (user.u32(12) !== 0xe391c05f) B.fail('PPT_USER', N('Unknown PowerPoint Current User header'));
    const persists = new Map(), edits = [], seen = new Set(); let at = user.u32(16), docRef;
    while (at) {
      if (seen.has(at)) B.fail('PPT_CYCLE', N('Cyclic PowerPoint edit history'));
      if (seen.size >= 65536) B.fail('PPT_LIMIT', N('Too many presentation edits')); seen.add(at);
      const edit = P.record(r, at); if (edit.type !== 4085) B.fail('PPT_EDIT', N('Invalid presentation edit record'));
      if (docRef == null) docRef = edit.r.u32(16);
      const ptr = P.record(r, edit.r.u32(12)); if (![6001, 6002].includes(ptr.type)) B.fail('PPT_PERSIST', N('Invalid presentation persist directory'));
      let p = 0;
      while (p < ptr.r.length) {
        const head = ptr.r.u32(p), id = head & 0xfffff, count = head >>> 20; p += 4; ptr.r.check(p, count * 4);
        if (!id || !count || id + count > 0x100000) B.fail('PPT_PERSIST', N('Invalid presentation persist range'));
        // Obsolete and optional persist objects need not be readable. Check the
        // selected object's offset in get(), not while building the directory.
        for (let i = 0; i < count; i++, p += 4) { const offset = ptr.r.u32(p); if (!persists.has(id + i)) persists.set(id + i, offset); }
      }
      edits.push(at); at = edit.r.u32(8);
    }
    const get = id => { const offset = persists.get(id); if (offset == null) B.fail('PPT_PERSIST', N('A presentation object is missing'), id); return P.record(r, offset); };
    const document = get(docRef); if (document.type !== 1000) B.fail('PPT_DOCUMENT', N('The current presentation document is missing'));
    const result = { version: 'PowerPoint97+', encrypted: false, r, user, persists, edits, document, get, slides: [], masters: [], notes: [], features: {}, issues: [] };
    const text = rec => { const parts = []; P.walk(rec, node => { result.features[node.type] = (result.features[node.type] || 0) + 1;
      if ([0xf00b, 0xf121, 0xf122].includes(node.type) && node.instance * 6 <= node.r.length) {
        for (let i = 0; i < node.instance; i++) if ((node.r.u16(i * 6) & 0x3fff) === 927 && (node.r.u32(i * 6 + 2) & 1)) result.features.tables = true;
      }
      if (node.type === 4000 || node.type === 4008) { if (node.r.length > B.limits.text) B.fail('TEXT_LIMIT', N('Presentation text is too large')); parts.push(node.type === 4000 ? node.r.text(0, node.r.length, 1200) : B.latin16(node.r.bytes)); } }); return parts; };
    for (const list of P.children(document).filter(x => x.type === 4080)) {
      const target = [result.slides, result.masters, result.notes][list.instance]; if (!target) continue;
      const first = target.length; let slide;
      for (const rec of P.children(list)) {
        if (rec.type === 1011) {
          const ref = rec.r.u32(0), id = rec.r.u32(12); let content;
          try { content = get(ref); } catch (e) { if (!options.recover || e.name !== 'LegacyError') throw e; result.issues.push({ code: e.code, ref, id, list: list.instance }); }
          slide = { ref, id, record: content, text: [], listRecords: [] }; target.push(slide);
        }
        else if (slide) { slide.listRecords.push(rec); if (options.content !== false) slide.text.push(...text(rec)); }
      }
      if (options.content !== false) for (const item of target.slice(first)) if (item.record) item.text.push(...text(item.record));
    }
    return result;
  };
})(typeof window !== 'undefined' ? window : globalThis);
