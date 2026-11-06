/* Excel FILEPASS and record-aware XOR/RC4 decryption, MS-XLS 2.2.10.
 * Derived from Apache OpenOffice main/sc/source/filter/excel/xistream.cxx (Apache-2.0).
 * Derived from Apache POI hssf/record/crypto/Biff8DecryptingStream.java (Apache-2.0).
 * Record headers and excluded fields consume RC4 keystream positions too. */
(function (root) {
  'use strict';
  const B = root.L.legacy, X = B.xls, E = B.encryption, N = s => s;
  const FILEPASS = 0x2f, BOUNDSHEET = 0x85, EOF = 0x0a, CODEPAGE = 0x42;
  const BOF = 0x809, BOF2 = 0x09, BOF3 = 0x209, BOF4 = 0x409;
  const INTERFACEHDR = 0xe1, USREXCL = 0x194, FILELOCK = 0x195;
  const RRDINFO = 0x196, RRDHEAD = 0x138, BLOCK_SIZE = 1024;
  const PLAIN = new Set([BOF, BOF2, BOF3, BOF4, FILEPASS, INTERFACEHDR, USREXCL, FILELOCK, RRDINFO, RRDHEAD]);
  X.decrypt = function (cfb, options = {}) {
    const name = cfb.entries.find(e => /^(workbook|book)$/i.test(e.path))?.path;
    if (!name) B.fail('XLS_STREAM', N('This file has no Excel workbook stream'));
    const stream = cfb.stream(name), r = new B.Reader(stream, name);
    let pass, at = 0, codepage = 1252, records = 0;
    while (at + 4 <= r.length) {
      if (++records > B.limits.records) B.fail('RECORD_LIMIT', N('Too many BIFF records'));
      const id = r.u16(at), size = r.u16(at + 2);
      if (!id && !size || id === EOF) break;
      r.check(at + 4, size);
      if (id === CODEPAGE && size >= 2) codepage = r.u16(at + 4);
      if (id === FILEPASS) {
        pass = r.reader(at + 4, size, 'FILEPASS');
        break;
      }
      at += 4 + size;
    }
    if (!pass) return { cfb, stream, decrypted: false };
    const old = r.u16(0) !== BOF || r.u16(4) !== 0x600;
    let info;
    if (old || pass.u16(0) === 0) {
      const start = old ? 0 : 2;
      info = { mode: 'xor', key: pass.u16(start), hash: pass.u16(start + 2) };
    } else if (pass.u16(0) === 1) info = E.header(pass.sub(2, pass.length - 2));
    else B.fail('ENCRYPTION_UNSUPPORTED', N('This workbook uses an unsupported password protection method'));
    const session = E.unlock(info, options.password, name, codepage);
    const plain = session.mode === 'xor' ? stream.slice() : E.blocks(stream, session, BLOCK_SIZE);
    const first = at + 4 + pass.length;
    plain.set(stream.subarray(0, first));
    const all = X.records(stream);
    for (const rec of all) {
      const start = rec.offset + 4, end = start + rec.r.length;
      if (rec.offset < first) continue;
      plain.set(stream.subarray(rec.offset, start), rec.offset);
      if (PLAIN.has(rec.id)) plain.set(rec.r.bytes, start);
      else if (session.mode === 'xor') plain.set(E.xorBytes(rec.r.bytes, session.array, end), start);
      if (rec.id === BOUNDSHEET) {
        rec.r.check(0, 4);
        plain.set(rec.r.sub(0, 4), start);
      }
    }
    const last = all.at(-1), end = last ? last.offset + 4 + last.r.length : 0;
    plain.set(stream.subarray(end), end);
    const view = E.overlay(cfb, (requested, key) => key === name.toLowerCase() ? plain : cfb.stream(requested));
    return { cfb: E.summary(view, session, 'Encryption'), stream: plain, decrypted: true };
  };
})(typeof window !== 'undefined' ? window : globalThis);
