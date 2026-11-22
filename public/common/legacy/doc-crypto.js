/* Word stream encryption, MS-DOC 2.2.6: 512-byte RC4 blocks; 68 clear FIB bytes.
 * Derived from Apache OpenOffice main/sw/source/filter/ww8/ww8par.cxx (Apache-2.0).
 * Derived from Apache POI hwpf/HWPFDocumentCore.java (Apache-2.0).
 * The original compound file and unrelated streams are left intact. */
(function (root) {
  'use strict';
  const B = root.L.legacy, D = B.doc, E = B.encryption, N = s => s;
  const FIB_CLEAR = 68, FLAGS = 10, KEY = 14, CHARSET_TABLE = 22, BLOCK_SIZE = 512;
  const CHARSETS = { 0: 1252, 128: 932, 129: 949, 134: 936, 136: 950, 161: 1253, 162: 1254,
    163: 1258, 177: 1255, 178: 1256, 186: 1257, 204: 1251, 222: 874, 238: 1250, 255: 850 };
  D.decrypt = function (cfb, options = {}) {
    const raw = cfb.stream('WordDocument');
    if (!raw) B.fail('DOC_STREAM', N('The WordDocument stream is missing'));
    const fib = D.fib(raw);
    if (!fib.encrypted || fib.version !== 'Word97+') return cfb;
    const length = fib.r.u32(KEY), tableName = fib.table.toLowerCase();
    const table = cfb.stream(fib.table);
    if (!table) B.fail('DOC_TABLE', N('The Word text table is missing'));
    const info = fib.obfuscated ? { mode: 'xor', word: true, key: length >>> 16, hash: length & 0xffff }
      : E.header(new B.Reader(table, fib.table).sub(0, length));
    const session = E.unlock(info, options.password, 'WordDocument', CHARSETS[fib.r.u16(CHARSET_TABLE)]);
    const view = E.overlay(cfb, (name, key) => {
      const bytes = key === 'worddocument' ? raw : key === tableName ? table : cfb.stream(name);
      if (!bytes || !['worddocument', tableName, 'data'].includes(key)) return bytes;
      const clear = key === 'worddocument' ? FIB_CLEAR : key === tableName && !fib.obfuscated ? length : 0;
      new B.Reader(bytes, name).check(0, clear);
      const plain = session.mode === 'xor' ? E.xorBytes(bytes, session.array, 0, true)
        : E.blocks(bytes, session, BLOCK_SIZE);
      plain.set(bytes.subarray(0, clear));
      if (key === 'worddocument') {
        const v = new DataView(plain.buffer);
        v.setUint16(FLAGS, fib.flags & ~0x8100, true);
        v.setUint32(KEY, 0, true);
      }
      return plain;
    });
    return E.summary(view, session, 'Encryption');
  };
})(typeof window !== 'undefined' ? window : globalThis);
