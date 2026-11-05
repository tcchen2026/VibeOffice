/* Derived from Apache POI poi/src/main/java/org/apache/poi/hssf/record/
 * (BOFRecord, BoundSheetRecord, SSTRecord, RKRecord, FormulaRecord; Apache-2.0).
 * BIFF records are interpreted per MS-XLS §2.4; byte bounds and data model are original. */
(function (root) {
  'use strict';
  const B = root.L.legacy, X = B.xls || (B.xls = {});
  X.records = function (bytes) {
    const r = new B.Reader(bytes, 'BIFF'), records = []; let p = 0;
    while (p < r.length) {
      if (r.length - p < 4) { if (r.sub(p, r.length - p).every(v => v === 0)) break; r.check(p, 4); }
      const id = r.u16(p), n = r.u16(p + 2);
      if (!id && !n) break;
      if (records.length >= B.limits.records) B.fail('RECORD_LIMIT', 'Too many BIFF records');
      records.push({ id, offset: p, r: r.reader(p + 4, n, 'BIFF ' + id.toString(16) + ':' + p) }); p += n + 4;
    }
    return records;
  };
  X.version = function (records) {
    const first = records[0];
    if (!first) B.fail('BIFF_HEADER', 'The workbook stream is empty');
    const v = { 9: 'BIFF2', 0x209: 'BIFF3', 0x409: 'BIFF4' }[first.id];
    if (v) return v;
    if (first.id !== 0x809) B.fail('BIFF_HEADER', 'The workbook has no BIFF header');
    const version = first.r.u16(0);
    return version === 0x600 ? 'BIFF8' : version === 0x500 ? 'BIFF5/7' : 'BIFF-' + version.toString(16);
  };
  X.string = function (r, p, biff8 = true, cp = 1252, short = false) {
    const n = short ? r.u8(p++) : r.u16(p); if (!short) p += 2;
    if (!biff8) return { value: r.text(p, n, cp), end: p + n };
    const flags = r.u8(p++), rich = flags & 8 ? r.u16(p) : 0; if (flags & 8) p += 2;
    const ext = flags & 4 ? r.u32(p) : 0; if (flags & 4) p += 4;
    const len = n * (flags & 1 ? 2 : 1), value = flags & 1 ? r.text(p, len, 1200) : B.latin16(r.sub(p, len));
    p += len; r.check(p, rich * 4 + ext); return { value, end: p + rich * 4 + ext };
  };
  // CONTINUE switches character width only while continuing the character array.
  X.sst = function (records, index, single = false) {
    const chunks = [records[index].r.bytes];
    for (let j = index + 1; records[j]?.id === 0x3c; j++) chunks.push(records[j].r.bytes);
    let ci = 0, pos = 0;
    const advance = () => { while (ci < chunks.length && pos === chunks[ci].length) { ci++; pos = 0; } if (ci >= chunks.length) B.fail('SST_TRUNCATED', 'Truncated shared string table'); };
    const byte = () => { advance(); return chunks[ci][pos++]; };
    const word = () => byte() | byte() << 8, dword = () => (word() + word() * 65536) >>> 0;
    const skip = n => { while (n) { advance(); const m = Math.min(n, chunks[ci].length - pos); pos += m; n -= m; } };
    const total = single ? 1 : dword(), unique = single ? 1 : dword();
    // Some real writers leave the use count stale; only the string array bounds matter.
    if (unique > B.limits.cells) B.fail('SST_LIMIT', 'Invalid shared string count');
    const strings = []; let characters = 0;
    for (let i = 0; i < unique; i++) {
      const n = word(); let flags = byte(); const rich = flags & 8 ? word() : 0, ext = flags & 4 ? dword() : 0;
      characters += n; if (characters > B.limits.text) B.fail('TEXT_LIMIT', 'The shared string table is too large');
      let value = '', remaining = n;
      while (remaining) {
        if (ci >= chunks.length) B.fail('SST_TRUNCATED', 'Truncated string');
        if (pos === chunks[ci].length) { ci++; pos = 0; flags = byte(); }
        const width = flags & 1 ? 2 : 1, take = Math.min(remaining, Math.floor((chunks[ci].length - pos) / width));
        if (!take) B.fail('SST_TRUNCATED', 'A shared string ends inside a character');
        const bytes = chunks[ci].subarray(pos, pos + take * width);
        value += width === 2 ? B.decode(bytes, 1200) : B.latin16(bytes); pos += take * width; remaining -= take;
      }
      skip(rich * 4 + ext); strings.push(value);
    }
    return strings;
  };
  X.rk = function (value) {
    let n;
    if (value & 2) n = value >> 2;
    else { const v = new DataView(new ArrayBuffer(8)); v.setUint32(4, value & 0xfffffffc, true); n = v.getFloat64(0, true); }
    return value & 1 ? n / 100 : n;
  };
  X.read = function (bytes) {
    const records = X.records(bytes), version = X.version(records), b8 = version === 'BIFF8';
    const result = { version, records, sheets: [], encrypted: records.some(x => x.id === 0x2f), features: {} };
    for (const rec of records) result.features[rec.id] = (result.features[rec.id] || 0) + 1;
    if (result.encrypted) return result;
    let cp = 1252; for (const rec of records) if (rec.id === 0x42) { cp = rec.r.u16(0); break; }
    const sstIndex = records.findIndex(x => x.id === 0xfc), strings = sstIndex < 0 ? [] : X.sst(records, sstIndex);
    result.strings = strings; result.codepage = cp;
    for (const rec of records) if (rec.id === 0x85) {
      const r = rec.r; result.sheets.push({ name: X.string(r, 6, b8, cp, true).value, offset: r.u32(0), state: r.u8(4), kind: r.u8(5), cells: [], records: [] });
    }
    if (!result.sheets.length && version !== 'BIFF8' && version !== 'BIFF5/7') return result;
    const byOffset = new Map(records.map((r, i) => [r.offset, i]));
    let count = 0;
    for (const sheet of result.sheets) {
      const begin = byOffset.get(sheet.offset); if (begin == null) B.fail('SHEET_OFFSET', 'A worksheet has an invalid stream offset', sheet.name);
      let lastFormula = null, depth = 0;
      for (let i = begin; i < records.length; i++) {
        const rec = records[i], r = rec.r, id = rec.id;
        sheet.records.push(rec);
        if (id === 0x809) { depth++; continue; } if (id === 0xa && --depth === 0) break; if (depth !== 1) continue;
        const add = (row, col, xf, type, value, extra = {}) => {
          if (++count > B.limits.cells) B.fail('CELL_LIMIT', 'Too many cells');
          if (row >= 65536 || col >= 256) B.fail('CELL_ADDRESS', 'Invalid BIFF cell address');
          const cell = { row, col, xf, type, value, offset: rec.offset, ...extra }; sheet.cells.push(cell); return cell;
        };
        if ([0x203, 0x27e, 0xfd, 0x205, 0x201, 0x204, 6].includes(id)) {
          const row = r.u16(0), col = r.u16(2), xf = r.u16(4);
          switch (id) {
            case 0x203: add(row, col, xf, 'n', r.f64(6)); break;
            case 0x27e: add(row, col, xf, 'n', X.rk(r.u32(6))); break;
            case 0xfd: { const at = r.u32(6); if (at >= strings.length) B.fail('SST_INDEX', 'Invalid shared string index'); add(row, col, xf, 's', strings[at]); break; }
            case 0x205: add(row, col, xf, r.u8(7) ? 'e' : 'b', r.u8(6)); break;
            case 0x201: add(row, col, xf, 'blank', null); break;
            case 0x204: add(row, col, xf, 's', X.string(r, 6, b8, cp).value); break;
            case 6: {
              const special = r.u16(12) === 0xffff, type = special ? ['s', 'b', 'e', 'empty'][r.u8(6)] : 'n';
              if (!type) B.fail('FORMULA_CACHE', 'Invalid formula result type');
              lastFormula = add(row, col, xf, type, special ? type === 's' || type === 'empty' ? '' : r.u8(8) : r.f64(6), { formula: r.sub(22, r.u16(20)), formulaRecord: rec }); break;
            }
          }
        } else if (id === 0xbd || id === 0xbe) {
          const row = r.u16(0), first = r.u16(2), last = r.u16(r.length - 2), step = id === 0xbd ? 6 : 2;
          if (last < first || r.length !== 6 + (last - first + 1) * step) B.fail('MUL_CELL', 'Invalid multiple-cell record');
          for (let col = first, p = 4; col <= last; col++, p += step) add(row, col, r.u16(p), id === 0xbd ? 'n' : 'blank', id === 0xbd ? X.rk(r.u32(p + 2)) : null);
        } else if (id === 0x207 && lastFormula?.type === 's') { lastFormula.value = b8 ? X.sst(records, i, true)[0] : X.string(r, 0, false, cp).value; lastFormula = null; }
      }
    }
    return result;
  };
})(typeof window !== 'undefined' ? window : globalThis);
