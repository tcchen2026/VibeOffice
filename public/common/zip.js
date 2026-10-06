/* VibeOffice — ZIP container engine (read + write), no dependencies; shared by Quire, Ledger and Lectern.
 * Uses the browser's Compression Streams when present and ships a small pure-JS inflater, so Office
 * Open XML packages open in any modern browser (and in Node, for the tests).
 */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const te = new TextEncoder();
  const td = new TextDecoder('utf-8');

  /* ---------- CRC-32 ---------- */
  const CRC_TABLE = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(buf) {
    let c = 0xffffffff;
    for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  }

  /* ---------- pure-JS inflate (RFC 1951) ---------- */
  const LBASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
  const LEXT = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
  const DBASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
  const DEXT = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
  const CLORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

  function buildHuff(lengths, n) {
    const counts = new Uint16Array(16), offs = new Uint16Array(16), syms = new Uint16Array(n);
    for (let i = 0; i < n; i++) counts[lengths[i]]++;
    counts[0] = 0;
    for (let i = 1; i < 15; i++) offs[i + 1] = offs[i] + counts[i];
    for (let i = 0; i < n; i++) if (lengths[i]) syms[offs[lengths[i]]++] = i;
    return { counts, syms };
  }
  let FIXED_L = null, FIXED_D = null;
  function fixedTables() {
    if (FIXED_L) return;
    const l = new Uint8Array(288);
    for (let i = 0; i < 144; i++) l[i] = 8;
    for (let i = 144; i < 256; i++) l[i] = 9;
    for (let i = 256; i < 280; i++) l[i] = 7;
    for (let i = 280; i < 288; i++) l[i] = 8;
    FIXED_L = buildHuff(l, 288);
    const d = new Uint8Array(30).fill(5);
    FIXED_D = buildHuff(d, 30);
  }

  function inflateRawJS(data, sizeHint) {
    let pos = 0, bitbuf = 0, bitcnt = 0;
    let out = new Uint8Array(Math.max(1024, sizeHint || data.length * 4)), op = 0;
    const need = (n) => {
      if (op + n <= out.length) return;
      let sz = out.length * 2; while (sz < op + n) sz *= 2;
      const nb = new Uint8Array(sz); nb.set(out.subarray(0, op)); out = nb;
    };
    const bits = (n) => {
      while (bitcnt < n) {
        if (pos >= data.length) throw new Error('inflate: unexpected end of data');
        bitbuf |= data[pos++] << bitcnt; bitcnt += 8;
      }
      const v = bitbuf & ((1 << n) - 1);
      bitbuf >>>= n; bitcnt -= n;
      return v;
    };
    const decode = (h) => {
      let code = 0, first = 0, index = 0;
      for (let len = 1; len < 16; len++) {
        code |= bits(1);
        const count = h.counts[len];
        if (code - count < first) return h.syms[index + (code - first)];
        index += count; first += count; first <<= 1; code <<= 1;
      }
      throw new Error('inflate: bad code');
    };
    const codes = (lt, dt) => {
      for (;;) {
        let sym = decode(lt);
        if (sym < 256) { need(1); out[op++] = sym; }
        else if (sym === 256) return;
        else {
          sym -= 257;
          if (sym >= 29) throw new Error('inflate: bad length');
          const len = LBASE[sym] + bits(LEXT[sym]);
          const ds = decode(dt);
          const dist = DBASE[ds] + bits(DEXT[ds]);
          if (dist > op) throw new Error('inflate: distance too far');
          need(len);
          for (let i = 0; i < len; i++) { out[op] = out[op - dist]; op++; }
        }
      }
    };
    let last;
    do {
      last = bits(1);
      const type = bits(2);
      if (type === 0) {
        bitbuf = 0; bitcnt = 0;
        const len = data[pos] | (data[pos + 1] << 8);
        pos += 4;
        need(len);
        out.set(data.subarray(pos, pos + len), op);
        op += len; pos += len;
      } else if (type === 1) {
        fixedTables();
        codes(FIXED_L, FIXED_D);
      } else if (type === 2) {
        const nlen = bits(5) + 257, ndist = bits(5) + 1, ncode = bits(4) + 4;
        const cl = new Uint8Array(19);
        for (let i = 0; i < ncode; i++) cl[CLORDER[i]] = bits(3);
        const clh = buildHuff(cl, 19);
        const lens = new Uint8Array(nlen + ndist);
        let i = 0;
        while (i < nlen + ndist) {
          const sym = decode(clh);
          if (sym < 16) lens[i++] = sym;
          else {
            let rep = 0, val = 0;
            if (sym === 16) { if (!i) throw new Error('inflate: repeat with no prior'); val = lens[i - 1]; rep = 3 + bits(2); }
            else if (sym === 17) rep = 3 + bits(3);
            else rep = 11 + bits(7);
            while (rep--) lens[i++] = val;
          }
        }
        codes(buildHuff(lens.subarray(0, nlen), nlen), buildHuff(lens.subarray(nlen), ndist));
      } else throw new Error('inflate: bad block type');
    } while (!last);
    return out.slice(0, op);
  }

  async function streamThrough(data, stream) {
    const res = new Response(new Blob([data]).stream().pipeThrough(stream));
    return new Uint8Array(await res.arrayBuffer());
  }
  async function inflateRaw(data, size) {
    if (typeof DecompressionStream === 'function') {
      try { return await streamThrough(data, new DecompressionStream('deflate-raw')); } catch (e) { /* fall back */ }
    }
    return inflateRawJS(data, size);
  }
  async function deflateRaw(data) {
    if (typeof CompressionStream === 'function') {
      try { return await streamThrough(data, new CompressionStream('deflate-raw')); } catch (e) { /* store */ }
    }
    return null;
  }

  /* ---------- reader ---------- */
  function makeEntry(u8, dv, name, method, csize, usize, loff, dataStart) {
    return {
      name, method, csize, usize, loff,
      async bytes() {
        if (this._cache) return this._cache;
        let start = dataStart;
        if (start == null) {
          const lh = this.loff;
          if (lh + 30 > u8.length || dv.getUint32(lh, true) !== 0x04034b50) throw new Error('Damaged entry ' + this.name);
          start = lh + 30 + dv.getUint16(lh + 26, true) + dv.getUint16(lh + 28, true);
        }
        const raw = u8.subarray(start, Math.min(u8.length, start + this.csize));
        let out;
        if (this.method === 0) out = raw.slice();
        else if (this.method === 8) out = await inflateRaw(raw, this.usize);
        else throw new Error('Unsupported compression method ' + this.method + ' in ' + this.name);
        this._cache = out;
        return out;
      },
      async text() { return td.decode(await this.bytes()); },
    };
  }
  const decodeName = (u8, flags) => (flags & 0x800 ? td.decode(u8) : Array.from(u8, (b) => String.fromCharCode(b)).join(''));
  function addEntry(files, entry) {
    if (entry.name.endsWith('/')) return;
    files.set(entry.name, entry);
    /* tolerate packages written with backslashes */
    const norm = entry.name.replace(/\\/g, '/');
    if (norm !== entry.name && !files.has(norm)) files.set(norm, entry);
  }
  function readCentral(u8, dv) {
    let eocd = -1;
    for (let i = u8.length - 22; i >= Math.max(0, u8.length - 65557); i--) {
      if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) return null;
    let count = dv.getUint16(eocd + 10, true);
    let cdOff = dv.getUint32(eocd + 16, true);
    /* ZIP64 end of central directory */
    if (cdOff === 0xffffffff || count === 0xffff) {
      const loc = eocd - 20;
      if (loc >= 0 && dv.getUint32(loc, true) === 0x07064b50) {
        const z64 = Number(dv.getBigUint64(loc + 8, true));
        count = Number(dv.getBigUint64(z64 + 32, true));
        cdOff = Number(dv.getBigUint64(z64 + 48, true));
      }
    }
    const files = new Map();
    let p = cdOff;
    for (let n = 0; n < count; n++) {
      if (p + 46 > u8.length || dv.getUint32(p, true) !== 0x02014b50) break;
      const flags = dv.getUint16(p + 8, true);
      const method = dv.getUint16(p + 10, true);
      let csize = dv.getUint32(p + 20, true);
      let usize = dv.getUint32(p + 24, true);
      const nlen = dv.getUint16(p + 28, true), xlen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true);
      let loff = dv.getUint32(p + 42, true);
      const name = decodeName(u8.subarray(p + 46, p + 46 + nlen), flags);
      /* zip64 extra field */
      let xp = p + 46 + nlen; const xend = Math.min(u8.length, xp + xlen);
      while (xp + 4 <= xend) {
        const id = dv.getUint16(xp, true), sz = dv.getUint16(xp + 2, true);
        if (id === 1) {
          let q = xp + 4;
          if (usize === 0xffffffff) { usize = Number(dv.getBigUint64(q, true)); q += 8; }
          if (csize === 0xffffffff) { csize = Number(dv.getBigUint64(q, true)); q += 8; }
          if (loff === 0xffffffff) { loff = Number(dv.getBigUint64(q, true)); }
        }
        xp += 4 + sz;
      }
      p += 46 + nlen + xlen + clen;
      addEntry(files, makeEntry(u8, dv, name, method, csize, usize, loff));
    }
    return files.size ? files : null;
  }
  /** Repair path: walk the local file headers when the central directory is missing or damaged. */
  function readLocal(u8, dv) {
    const files = new Map();
    let p = 0;
    while (p + 30 <= u8.length) {
      if (dv.getUint32(p, true) !== 0x04034b50) {
        /* resynchronise on the next local header */
        let q = p + 1;
        while (q + 4 <= u8.length && !(u8[q] === 0x50 && u8[q + 1] === 0x4b && u8[q + 2] === 0x03 && u8[q + 3] === 0x04)) q++;
        if (q + 30 > u8.length) break;
        p = q;
        continue;
      }
      const flags = dv.getUint16(p + 6, true), method = dv.getUint16(p + 8, true);
      let csize = dv.getUint32(p + 18, true), usize = dv.getUint32(p + 22, true);
      const nlen = dv.getUint16(p + 26, true), xlen = dv.getUint16(p + 28, true);
      const name = decodeName(u8.subarray(p + 30, p + 30 + nlen), flags);
      const start = p + 30 + nlen + xlen;
      let next = start + csize;
      if ((flags & 8) || csize === 0xffffffff || (csize === 0 && method === 8)) {
        /* sizes live in a trailing data descriptor: find one whose size field matches */
        let q = start, found = false;
        while (q + 16 <= u8.length) {
          if (u8[q] === 0x50 && u8[q + 1] === 0x4b) {
            const sig = dv.getUint32(q, true);
            if (sig === 0x08074b50 && dv.getUint32(q + 8, true) === q - start) { csize = q - start; usize = dv.getUint32(q + 12, true); next = q + 16; found = true; break; }
            if (sig === 0x04034b50 || sig === 0x02014b50) { csize = q - start; usize = 0; next = q; found = true; break; }
          }
          q++;
        }
        if (!found) { csize = u8.length - start; usize = 0; next = u8.length; }
      }
      if (start > u8.length) break;
      addEntry(files, makeEntry(u8, dv, name, method, csize, usize, p, start));
      p = Math.max(next, p + 30);
    }
    return files;
  }
  async function read(buffer) {
    const u8 = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (u8.length < 22) throw new Error('This file is not a valid Office Open XML (ZIP) package.');
    if (dv.getUint32(0, true) !== 0x04034b50) {
      if (dv.getUint32(0, true) === 0xe011cfd0) {
        /* OLE compound file: an encrypted OOXML package or a legacy binary .ppt */
        const probe = new TextDecoder('utf-16le').decode(u8.subarray(0, Math.min(u8.length, 1 << 20)));
        const enc = /EncryptedPackage/.test(probe);
        const err = new Error(enc
          ? 'This file is password-protected.'
          : 'This is a binary Office 97–2003 file, not an Office Open XML package.');
        err.code = 'ole';
        err.encrypted = enc;
        err.bytes = u8;
        throw err;
      }
    }
    let files = null;
    try { files = readCentral(u8, dv); } catch (e) { files = null; }
    if (!files) {
      try { files = readLocal(u8, dv); } catch (e) { files = null; }
      if (!files || !files.size) throw new Error(dv.getUint32(0, true) === 0x04034b50 ? 'This file is damaged and could not be repaired.' : 'This file is not a valid Office Open XML (ZIP) package.');
      files.repaired = true;
    }
    return files;
  }

  /* ---------- writer ---------- */
  function dosDateTime(d) {
    const time = (d.getHours() << 11) | (d.getMinutes() << 5) | Math.floor(d.getSeconds() / 2);
    const date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
    return { time, date };
  }
  /** files: [{name, data: Uint8Array|string, store?: bool}] → Blob */
  async function write(files, mime) {
    const parts = [];
    const central = [];
    let offset = 0;
    const { time, date } = dosDateTime(new Date());
    for (const f of files) {
      const nameBytes = te.encode(f.name);
      const data = typeof f.data === 'string' ? te.encode(f.data) : f.data instanceof Uint8Array ? f.data : new Uint8Array(f.data);
      const crc = crc32(data);
      let comp = null;
      if (!f.store && data.length > 64) comp = await deflateRaw(data);
      const method = comp && comp.length < data.length ? 8 : 0;
      const body = method === 8 ? comp : data;
      const lh = new Uint8Array(30 + nameBytes.length);
      const ldv = new DataView(lh.buffer);
      ldv.setUint32(0, 0x04034b50, true);
      ldv.setUint16(4, 20, true);
      ldv.setUint16(6, 0x0800, true);
      ldv.setUint16(8, method, true);
      ldv.setUint16(10, time, true);
      ldv.setUint16(12, date, true);
      ldv.setUint32(14, crc, true);
      ldv.setUint32(18, body.length, true);
      ldv.setUint32(22, data.length, true);
      ldv.setUint16(26, nameBytes.length, true);
      ldv.setUint16(28, 0, true);
      lh.set(nameBytes, 30);
      parts.push(lh, body);
      const ch = new Uint8Array(46 + nameBytes.length);
      const cdv = new DataView(ch.buffer);
      cdv.setUint32(0, 0x02014b50, true);
      cdv.setUint16(4, 20, true);
      cdv.setUint16(6, 20, true);
      cdv.setUint16(8, 0x0800, true);
      cdv.setUint16(10, method, true);
      cdv.setUint16(12, time, true);
      cdv.setUint16(14, date, true);
      cdv.setUint32(16, crc, true);
      cdv.setUint32(20, body.length, true);
      cdv.setUint32(24, data.length, true);
      cdv.setUint16(28, nameBytes.length, true);
      cdv.setUint32(42, offset, true);
      ch.set(nameBytes, 46);
      central.push(ch);
      offset += lh.length + body.length;
    }
    const cdSize = central.reduce((a, c) => a + c.length, 0);
    const end = new Uint8Array(22);
    const edv = new DataView(end.buffer);
    edv.setUint32(0, 0x06054b50, true);
    edv.setUint16(8, files.length, true);
    edv.setUint16(10, files.length, true);
    edv.setUint32(12, cdSize, true);
    edv.setUint32(16, offset, true);
    return new Blob([...parts, ...central, end], { type: mime || 'application/zip' });
  }

  L.zip = { read, write, crc32, inflateRawJS, encode: (s) => te.encode(s), decode: (b) => td.decode(b) };
})(typeof window !== 'undefined' ? window : globalThis);
