/* Derived from Apache POI poi-scratchpad/src/main/java/org/apache/poi/hwpf/model/
 * FileInformationBlock, FIBFieldHandler, ComplexFileTable, PieceDescriptor (Apache-2.0).
 * FIB and text pieces: MS-DOC §§2.5, 2.8, 2.9. Original bounded byte implementation. */
(function (root) {
  'use strict';
  const N = s => s; // UI refusal text, translated where it is shown.
  const B = root.L.legacy, D = B.doc || (B.doc = {});
  // CPs address UTF-16 code units. TextDecoder consumes a BOM and replaces a
  // lone surrogate at a piece boundary, although the following piece can hold
  // its low half. Decode units first; XML escaping handles truly broken pairs.
  D.unicode = function(r, offset, bytes) {
    r.check(offset,bytes);if(bytes%2)B.fail('DOC_TEXT',N('A Unicode text piece has an incomplete character'));
    let text='';for(let p=0;p<bytes;p+=8192){const n=Math.min(8192,bytes-p),units=[];for(let i=0;i<n;i+=2)units.push(r.u16(offset+p+i));text+=String.fromCharCode(...units);}return text;
  };
  D.fib = function (bytes) {
    const r = new B.Reader(bytes, 'WordDocument');
    if (![0xa5ec, 0xa5dc].includes(r.u16(0))) B.fail('DOC_HEADER', N('Invalid Word binary header'));
    if (r.u16(0) === 0xa5dc && r.u16(2) === 0xa5db) return { version: 'Word2', nFib: r.u16(2), encrypted: false, r };
    const nFib = r.u16(2), flags = r.u16(10), result = { nFib, version: nFib < 0xc0 ? 'Word6/95' : 'Word97+',
      encrypted: !!(flags & 0x100), obfuscated: !!(flags & 0x8000), template: !!(flags & 1), table: flags & 0x200 ? '1Table' : '0Table', flags, r };
    if (nFib < 0xc0 || result.encrypted) return result;
    let p = 32; const csw = r.u16(p); p += 2; r.check(p, csw * 2); result.words = r.reader(p, csw * 2); p += csw * 2;
    const cslw = r.u16(p); p += 2; result.longs = r.reader(p, cslw * 4); p += cslw * 4;
    const pairs = r.u16(p); p += 2; result.pairs = r.reader(p, pairs * 8);
    result.pair = id => id * 8 + 8 <= result.pairs.length ? { offset: result.pairs.u32(id * 8), length: result.pairs.u32(id * 8 + 4) } : { offset: 0, length: 0 };
    result.stories = {};
    for (const [i, name] of ['main', 'footnotes', 'headers', 'macros', 'comments', 'endnotes', 'textboxes', 'headerTextboxes'].entries())
      result.stories[name] = result.longs.u32((i + 3) * 4);
    return result;
  };
  // Direct insertion/deletion/character-format revision flags in CHPX FKPs.
  // Derived from Apache POI hwpf/model/CHPFormattedDiskPage.java (Apache-2.0).
  D.revisionFlags = function (fib, table) {
    const pair = fib.pair(12); if (!pair.length) return false;
    const bte = table.reader(pair.offset, pair.length, 'PlcfBteChpx');
    if (bte.length < 4 || (bte.length - 4) % 8) B.fail('DOC_BTE', N('Invalid Word character-property table'));
    const n = (bte.length - 4) / 8, seen = new Set();
    for (let i = 0; i < n; i++) {
      const page = bte.u32((n + 1) * 4 + i * 4) & 0x3fffff;
      if (seen.has(page)) continue; seen.add(page);
      const fkp = fib.r.reader(page * 512, 512, 'CHPX FKP'), runs = fkp.u8(511);
      fkp.check(0, (runs + 1) * 4 + runs);
      for (let j = 0; j < runs; j++) {
        const at = fkp.u8((runs + 1) * 4 + j) * 2; if (!at) continue;
        const props = fkp.reader(at + 1, fkp.u8(at), 'CHPX sprms'); let p = 0;
        while (p < props.length) {
          const op = props.u16(p); p += 2; const form = op >>> 13;
          let length = [1, 1, 2, 4, 2, 2, 0, 3][form];
          if (form === 6) length = 1 + props.u8(p);
          props.check(p, length);
          if (((op === 0x0800 || op === 0x0801) && props.u8(p)) || (op === 0xca57 && length > 1 && props.u8(p + 1))) return true;
          p += length;
        }
      }
    }
    return false;
  };
  D.read = function (cfb) {
    const bytes = cfb.stream('WordDocument'); if (!bytes) B.fail('DOC_STREAM', N('The WordDocument stream is missing'));
    const fib = D.fib(bytes), result = { ...fib, pieces: [], text: '', storyText: {} };
    if (fib.encrypted || fib.version !== 'Word97+') return result;
    const tableBytes = cfb.stream(fib.table); if (!tableBytes) B.fail('DOC_TABLE', N('The Word text table is missing'));
    const table = new B.Reader(tableBytes, fib.table), clx = fib.pair(33), r = table.reader(clx.offset, clx.length, 'CLX');
    result.tableReader = table; result.prcs = [];
    if (r.length) {
      let p = 0;
      while (p < r.length && r.u8(p) === 1) { const n = r.u16(p + 1); result.prcs.push(r.sub(p + 3, n)); p += 3 + n; }
      if (r.u8(p) !== 2) B.fail('DOC_PIECES', N('The Word text piece table is missing'));
      const size = r.u32(p + 1); p += 5; const plc = r.reader(p, size, 'PlcPcd');
      if (size < 4 || (size - 4) % 12) B.fail('DOC_PIECES', N('Invalid Word text piece table size'));
      const n = (size - 4) / 12; let total = 0;
      if (n > 1000000) B.fail('DOC_PIECES', N('Too many Word text pieces'));
      for (let i = 0; i < n; i++) {
        const start = plc.u32(i * 4), end = plc.u32((i + 1) * 4), at = (n + 1) * 4 + i * 8, fc = plc.u32(at + 2), compressed = !!(fc & 0x40000000);
        if (end < start || (i === 0 && start !== 0)) B.fail('DOC_PIECES', N('Invalid Word character positions'));
        const offset = (fc & 0x3fffffff) / (compressed ? 2 : 1), chars = end - start;
        if ((total += chars) > B.limits.text) B.fail('TEXT_LIMIT', N('The Word document has too much text'));
        const text = compressed ? fib.r.text(offset, chars, 1252) : D.unicode(fib.r,offset,chars*2);
        result.pieces.push({ start, end, offset, compressed, prm: plc.u16(at + 6), text });
      }
      result.text = result.pieces.map(x => x.text).join('');
    } else {
      const start = fib.r.u32(24), end = fib.r.u32(28); if (end < start) B.fail('DOC_TEXT', N('Invalid Word text range'));
      // Non-complex files can carry a contiguous text range instead of CLX.
      const unicode = !!(fib.flags & 0x1000), text = unicode ? D.unicode(fib.r,start,end-start) : fib.r.text(start,end-start,1252);
      result.pieces.push({ start: 0, end: text.length, offset: start, compressed: !unicode, prm: 0, text }); result.text = text;
    }
    let pos = 0;
    for (const [name, n] of Object.entries(fib.stories)) { if (pos + n > result.text.length) B.fail('DOC_STORIES', N('Word story lengths exceed the text piece table')); result.storyText[name] = result.text.slice(pos, pos + n); pos += n; }
    result.features = { fields: [16, 17, 18, 19, 20, 48, 57, 59].some(id => fib.pair(id).length > 0),
      comments: fib.pair(4).length > 0, bookmarks: fib.pair(22).length > 0, lists: fib.pair(73).length > 0,
      drawings: fib.pair(50).length > 0 || fib.pair(40).length > 0, tableCellMarkers: (result.storyText.main.match(/\x07/g) || []).length,
      footnotes: fib.stories.footnotes > 0, endnotes: fib.stories.endnotes > 0, textboxes: fib.stories.textboxes > 0,
      sections: fib.pair(6).length >= 4 ? (fib.pair(6).length - 4) / 16 : 0 };
    try { result.features.revisionFlags = D.revisionFlags(fib, table); }
    catch (e) { if (e.name !== 'LegacyError') throw e; result.features.revisionFlags = null; result.features.revisionScanError = e.code; }
    result.features.tables = result.features.tableCellMarkers > 0;
    result.features.inlinePictureMarkers = (result.storyText.main.match(/\x01/g) || []).length;
    return result;
  };
})(typeof window !== 'undefined' ? window : globalThis);
