/* OfficeArt records, properties and BLIPs (MS-ODRAW §§2.2–2.3).
 * Derived from Apache POI ddf/{EscherPropertyFactory,EscherArrayProperty,EscherBSERecord,
 * EscherMetafileBlip}.java (Apache-2.0).
 * Derived from Apache OpenOffice filter/source/msfilter/msdffimp.cxx (Apache-2.0).
 * Kept independent of the document host and of DOM/rendering. */
(function (root) {
  'use strict';
  const N = s => s; // UI refusal text, translated where it is shown.
  const B = root.L.legacy, A = B.officeart = {};
  A.record = function (bytes, at = 0) {
    const r = bytes instanceof B.Reader ? bytes : new B.Reader(bytes, 'OfficeArt');
    r.check(at, 8); const options = r.u16(at), length = r.u32(at + 4); r.check(at + 8, length);
    return { at, end: at + 8 + length, type: r.u16(at + 2), version: options & 15, instance: options >>> 4, r: r.reader(at + 8, length) };
  };
  A.records = function (bytes, visit, depth = 0, budget = { left: B.limits.records }) {
    const r = bytes instanceof B.Reader ? bytes : new B.Reader(bytes, 'OfficeArt'), records = [];
    if (depth >= B.limits.depth) B.fail('OFFICEART_DEPTH', N('OfficeArt nesting is too deep'));
    for (let p = 0; p < r.length;) {
      if (--budget.left < 0) B.fail('OFFICEART_LIMIT', N('Too many OfficeArt records'));
      const rec = A.record(r, p); p = rec.end; records.push(rec); if (visit) visit(rec, depth);
      if (rec.version === 15) rec.children = A.records(rec.r, visit, depth + 1, budget);
    }
    return records;
  };
  const arrays = new Set([325, 326, 337, 338, 341, 342, 343, 345, 407, 463, 899, 928]);
  A.properties = function (rec) {
    const r = rec.r, count = rec.instance, props = new Map(); r.check(0, count * 6); let p = count * 6;
    for (let i = 0; i < count; i++) {
      const flags = r.u16(i * 6), id = flags & 0x3fff, value = r.u32(i * 6 + 2), prop = { id, value, complex: !!(flags & 0x8000), blip: !!(flags & 0x4000) };
      if (prop.complex) {
        let len = value;
        if (arrays.has(id) && len && p <= r.length - 6) {
          const size = r.i16(p + 4), actual = size < 0 ? -size >>> 2 : size;
          if (actual * r.u16(p) === len) len += 6; // Some Office writers omit the array header from cb.
        }
        prop.data = r.sub(p, len); p += len;
      }
      props.set(id, prop);
    }
    return props;
  };
  A.array = function (prop) {
    if (!prop?.data?.length) return [];
    const r = new B.Reader(prop.data, 'OfficeArt array'), count = r.u16(0), storedSize = r.i16(4), size = storedSize < 0 ? -storedSize >>> 2 : storedSize;
    if (!size && count) B.fail('OFFICEART_ARRAY', N('Invalid OfficeArt array element size'));
    r.check(6, count * size); return Array.from({ length: count }, (_, i) => r.reader(6 + i * size, size));
  };
  A.blip = async function (record) {
    const r = record.r, type = record.type, secondUID = record.instance & 1; let p = secondUID ? 32 : 16; r.check(0, p);
    if ([0xf01a, 0xf01b, 0xf01c].includes(type)) {
      const expected = r.u32(p), bounds = [r.i32(p + 4), r.i32(p + 8), r.i32(p + 12), r.i32(p + 16)], width = r.i32(p + 20), height = r.i32(p + 24), size = r.u32(p + 28), compression = r.u8(p + 32); p += 34;
      const compressed = r.sub(p, size);
      if (![0, 0xfe].includes(compression)) B.fail('BLIP_COMPRESSION', N('Unsupported metafile compression'));
      const bytes = compression === 0 ? await B.images.inflate(compressed, expected) : compressed;
      if (bytes.length > 64 * 1024 * 1024) B.fail('IMAGE_LIMIT', N('The metafile is too large'));
      const ext = type === 0xf01a ? 'emf' : type === 0xf01b ? 'wmf' : 'pict';
      if (ext === 'pict') B.fail('BLIP_FORMAT', N('PICT pictures cannot be converted yet'));
      return { bytes, ext, mime: 'image/x-' + ext, bounds, width, height };
    }
    const types = { 0xf01d: ['jpeg', 'image/jpeg'], 0xf01e: ['png', 'image/png'], 0xf01f: ['dib', 'image/bmp'], 0xf029: ['tif', 'image/tiff'], 0xf02a: ['jpeg', 'image/jpeg'] };
    if (!types[type]) B.fail('BLIP_FORMAT', N('This picture uses an unsupported format'), type);
    p++; const bytes = r.sub(p, r.length - p), [ext, mime] = types[type];
    return ext === 'dib' ? B.images.dib(bytes) : { bytes, ext, mime };
  };
  A.store = function (records, delayed) {
    const entries = [];
    const add = rec => {
      if (rec.type === 0xf001) {
        for (const entry of rec.children || A.records(rec.r)) {
          if (entry.type !== 0xf007) B.fail('BLIP_STORE', N('Invalid picture-store entry'));
          const r = entry.r, size = r.length >= 36 ? r.u32(20) : 0, references = r.length >= 36 ? r.u32(24) : 0, offset = r.length >= 36 ? r.u32(28) : 0;
          const index = entries.length + 1; let promise;
          entries.push({ index, references, size, offset, read: () => promise || (promise = (async () => {
            // A damaged optional picture must not invalidate the indices or
            // readable entries that follow it in the shared picture store.
            r.check(0, 36); const nameLength = r.u8(33);
            const inlineAt = 36 + nameLength;
            if (r.length >= inlineAt + 8) return A.blip(A.record(r, inlineAt));
            // Word permits a fallback stream for pictures; the host supplies it explicitly.
            const streams = Array.isArray(delayed) ? delayed : [delayed]; let last;
            for (const bytes of streams) if (bytes) {
              try { return await A.blip(A.record(bytes, offset)); }
              catch (e) { if (e.name !== 'LegacyError') throw e; last = e; }
            }
            if (last) throw last;
            B.fail('BLIP_MISSING', N('The picture data is missing'), index);
          })()) });
        }
      } else for (const child of rec.children || []) add(child);
    };
    for (const rec of records) add(rec); return entries;
  };
  A.anchor = function (record, host) {
    const r = record.r;
    if (record.type === 0xf00f) return { kind: 'child', left: r.i32(0), top: r.i32(4), right: r.i32(8), bottom: r.i32(12) };
    if (host === 'ppt') {
      const small = r.length === 8, get = p => small ? r.i16(p * 2) : r.i32(p * 4);
      return { kind: 'rect', top: get(0), left: get(1), right: get(2), bottom: get(3), units: 576 };
    }
    if (host === 'xls') return { kind: 'cells', flags: r.u16(0), col1: r.u16(2), dx1: r.u16(4), row1: r.u16(6), dy1: r.u16(8), col2: r.u16(10), dx2: r.u16(12), row2: r.u16(14), dy2: r.u16(16) };
    // Word's ClientAnchor is just an index into its host drawing data (PlcSpa supplies the rectangle).
    return { kind: 'host', data: r.bytes };
  };
  A.shapes = function (records, host) {
    const read = rec => {
      const children = rec.children || [], shape = { properties: new Map(), children: [], record: rec };
      for (const c of children) {
        if (c.type === 0xf00a) { shape.id = c.r.u32(0); shape.flags = c.r.u32(4); shape.type = c.instance; }
        else if ([0xf00b, 0xf121, 0xf122].includes(c.type)) for (const [id, p] of A.properties(c)) shape.properties.set(id, p);
        else if (c.type === 0xf009) shape.coord = { left: c.r.i32(0), top: c.r.i32(4), right: c.r.i32(8), bottom: c.r.i32(12) };
        else if ([0xf00f, 0xf010].includes(c.type)) shape.anchor = A.anchor(c, host);
        else if (c.type === 0xf00d) shape.textbox = c.r.bytes;
        else if (c.type === 0xf011) shape.clientData = c.r.bytes;
      }
      shape.table = !!(shape.properties.get(927)?.value & 1); return shape;
    };
    const visit = rec => {
      if (rec.type === 0xf004) return [read(rec)];
      if (rec.type === 0xf003) {
        const kids = rec.children || [], groupRecord = kids.find(c => c.type === 0xf004);
        if (!groupRecord) return kids.flatMap(visit);
        const group = read(groupRecord); group.children = kids.filter(c => c !== groupRecord).flatMap(visit); return [group];
      }
      return (rec.children || []).flatMap(visit);
    };
    return records.flatMap(visit);
  };
})(typeof window !== 'undefined' ? window : globalThis);
