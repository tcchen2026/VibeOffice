/* VBA project preservation and bounded metadata decoding, MS-OVBA 2.2–2.4.
 * Derived from Apache OpenOffice main/oox/source/ole/{vbainputstream,vbaproject}.cxx (Apache-2.0).
 * Module source, compiled caches, protection settings and designer streams are copied unchanged.
 * No macro is evaluated or executed. Metadata decoding never rewrites the project. */
(function (root) {
  'use strict';
  const B = root.L.legacy;
  const V = B.vba;
  const N = s => s;
  const CHUNK_SIZE = 4096;
  const PROJECT_CODEPAGE = 0x0003;
  const PROJECT_VERSION = 0x0009;
  const PROJECT_MODULES = 0x000f;
  const PROJECT_END = 0x0010;
  const bad = detail => B.fail('VBA_DATA', 'Invalid VBA project data', detail);

  V.decompress = function (bytes, limit = B.limits.text) {
    const r = new B.Reader(bytes, 'VBA compressed stream');
    if (r.u8(0) !== 1) bad('Invalid compressed stream signature');
    const chunks = [];
    let p = 1;
    let length = 0;
    while (p < r.length) {
      if (chunks.length > Math.ceil(limit / CHUNK_SIZE)) bad('Too many compressed chunks');
      const header = r.u16(p);
      const end = p + (header & 0xfff) + 3;
      if ((header & 0x7000) !== 0x3000 || end > r.length) bad('Invalid compressed chunk');
      p += 2;
      const chunk = new Uint8Array(CHUNK_SIZE);
      let size = 0;
      if (!(header & 0x8000)) {
        if (end - p !== CHUNK_SIZE) bad('Invalid uncompressed chunk');
        chunk.set(r.sub(p, CHUNK_SIZE));
        size = CHUNK_SIZE;
        p = end;
      } else {
        while (p < end) {
          const flags = r.u8(p++);
          for (let bit = 0; bit < 8 && p < end; bit++) {
            if (flags & (1 << bit)) {
              if (p + 2 > end) bad('Truncated copy token');
              const token = r.u16(p);
              p += 2;
              let bits = 4;
              while (2 ** bits < size) bits++;
              const count = (token & (0xffff >>> bits)) + 3;
              const offset = (token >>> (16 - bits)) + 1;
              if (offset > size || size + count > CHUNK_SIZE) bad('Invalid copy token');
              for (let i = 0; i < count; i++) {
                chunk[size] = chunk[size - offset];
                size++;
              }
            } else {
              if (size >= CHUNK_SIZE) bad('Oversized compressed chunk');
              chunk[size++] = r.u8(p++);
            }
          }
        }
      }
      length += size;
      if (length > limit) bad('Decompressed metadata exceeds the byte limit');
      chunks.push(chunk.subarray(0, size));
    }
    const out = new Uint8Array(length);
    let at = 0;
    for (const chunk of chunks) {
      out.set(chunk, at);
      at += chunk.length;
    }
    return out;
  };

  V.notice = function (losses, what, where, detail, notify = true) {
    const entry = B.loss('legacy-macros', what, where, 'drop', where, notify);
    if (detail) entry.detail = detail;
    losses.push(entry);
  };
  V.damaged = (losses, where, error) => {
    if (!['LegacyError', 'CFBError'].includes(error.name)) throw error;
    V.notice(losses, N('Damaged macros were not kept'), where, error.code);
    return null;
  };

  V.metadata = function (source, prefix) {
    const dir = new B.Reader(V.decompress(source.stream(prefix + 'VBA/dir')), 'VBA directory');
    let codepage = 1252;
    // The code page precedes references and module records. Stop there; future
    // reference records or module versions must not prevent opaque preservation.
    for (let p = 0; p < dir.length;) {
      const id = dir.u16(p);
      const size = id === PROJECT_VERSION ? 6 : dir.u32(p + 2);
      p += 6;
      dir.check(p, size);
      if (id === PROJECT_CODEPAGE) {
        codepage = dir.u16(p);
        break;
      }
      if (id === PROJECT_MODULES || id === PROJECT_END) break;
      p += size;
    }
    const text = B.decode(source.stream(prefix + 'PROJECT'), codepage);
    const names = new Map();
    const mapping = source.stream(prefix + 'PROJECTwm');
    if (mapping) {
      const r = new B.Reader(mapping, 'VBA module name map');
      let p = 0;
      while (p + 1 < r.length && r.u8(p)) {
        const start = p;
        while (p < r.length && r.u8(p)) p++;
        const ansi = B.decode(r.sub(start, p - start), codepage);
        p++;
        const unicodeStart = p;
        while (p + 1 < r.length && r.u16(p)) p += 2;
        const unicode = r.text(unicodeStart, p - unicodeStart, 1200);
        r.check(p, 2);
        p += 2;
        names.set(ansi, unicode);
      }
    }
    const documents = [];
    const forms = [];
    for (const line of text.split(/\r\n|\r|\n/)) {
      if (line.startsWith('[')) break;
      const match = /^(Document|BaseClass)=(.*)$/.exec(line);
      if (!match) continue;
      const name = match[2].replace(/\/&H[0-9a-f]+$/i, '');
      (match[1] === 'Document' ? documents : forms).push(names.get(name) || name);
    }
    return { codepage, documents, forms };
  };

  V.read = function (source, prefix, losses, where) {
    const storage = source.entries.find(e => e.type === 1 && e.path.toUpperCase() === prefix.toUpperCase());
    if (prefix && !storage) return null;
    prefix = storage ? storage.path + '/' : '';
    const entries = source.entries.filter(e => e.path.toUpperCase().startsWith(prefix.toUpperCase()));
    if (!entries.length) return null;
    try {
      for (const name of ['PROJECT', 'VBA/dir', 'VBA/_VBA_PROJECT']) {
        const entry = entries.find(e => e.type === 2 && e.path.toUpperCase() === (prefix + name).toUpperCase());
        if (!entry?.size) bad('Missing ' + name);
      }
      const bytes = V.pack(source, prefix);
      const project = { bytes, streams: entries.filter(e => e.type === 2).length, metadata: null };
      try {
        project.metadata = V.metadata(source, prefix);
      } catch (error) {
        if (!['LegacyError', 'CFBError'].includes(error.name)) throw error;
        V.notice(losses, 'Some macro project information could not be read', where, error.code, false);
      }
      return project;
    } catch (error) {
      return V.damaged(losses, where, error);
    }
  };

  V.add = function (pkg, main, project) {
    if (!project) return false;
    const folder = main.slice(0, main.lastIndexOf('/') + 1);
    pkg.add(folder + 'vbaProject.bin', project.bytes, 'application/vnd.ms-office.vbaProject');
    pkg.rel(main, 'http://schemas.microsoft.com/office/2006/relationships/vbaProject', 'vbaProject.bin');
    return true;
  };
})(typeof window !== 'undefined' ? window : globalThis);
