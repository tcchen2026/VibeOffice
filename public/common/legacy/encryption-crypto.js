/* Primitives used only to read encrypted Office 97-2003 files.
 * MD5: independent implementation of RFC 1321 section 3 (not its sample code).
 * RC4: independent byte permutation and keystream implementation; RFC 6229 test vectors.
 * Derived from Apache OpenOffice main/filter/source/msfilter/mscodec.cxx (Apache-2.0): XOR95.
 * These old ciphers are not used to encrypt newly saved files. */
(function (root) {
  'use strict';
  const B = root.L.legacy, E = B.encryption = {};
  const SHIFT = [[7, 12, 17, 22], [5, 9, 14, 20], [4, 11, 16, 23], [6, 10, 15, 21]];
  const SINE = Uint32Array.from({ length: 64 }, (_, i) => Math.floor(Math.abs(Math.sin(i + 1)) * 0x100000000));
  const PADDING = [0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0, 0xbe, 0x0f, 0, 0xbf, 0x0f, 0];
  const rotate = (x, n, bits) => ((x << n) | (x >>> (bits - n))) & (2 ** bits - 1);
  E.join = function (...arrays) {
    const size = arrays.reduce((n, a) => n + a.length, 0);
    if (size > B.limits.bytes) B.fail('SIZE_LIMIT', 'Encryption buffer exceeds the import limit');
    const out = new Uint8Array(size);
    let p = 0;
    for (const a of arrays) {
      out.set(a, p);
      p += a.length;
    }
    return out;
  };
  E.utf16 = function (text) {
    const out = new Uint8Array(text.length * 2);
    for (let i = 0; i < text.length; i++) {
      const c = text.charCodeAt(i);
      out[i * 2] = c;
      out[i * 2 + 1] = c >>> 8;
    }
    return out;
  };
  E.md5 = function (bytes) {
    const r = new B.Reader(bytes), size = Math.ceil((r.length + 9) / 64) * 64;
    const padded = new Uint8Array(size), words = new DataView(padded.buffer);
    padded.set(bytes);
    padded[r.length] = 0x80;
    words.setUint32(size - 8, r.length * 8, true);
    words.setUint32(size - 4, Math.floor(r.length / 0x20000000), true);
    const state = [0x67452301, 0xefcdab89, 0x98badcfe, 0x10325476];
    for (let p = 0; p < size; p += 64) {
      let [a, b, c, d] = state;
      for (let i = 0; i < 64; i++) {
        const round = i >>> 4;
        const f = [() => (b & c) | (~b & d), () => (d & b) | (~d & c),
          () => b ^ c ^ d, () => c ^ (b | ~d)][round]();
        const index = [i, 5 * i + 1, 3 * i + 5, 7 * i][round] & 15;
        const n = SHIFT[round][i & 3], value = (a + f + SINE[i] + words.getUint32(p + index * 4, true)) | 0;
        const next = (b + ((value << n) | (value >>> (32 - n)))) | 0;
        a = d;
        d = c;
        c = b;
        b = next;
      }
      for (const [i, value] of [a, b, c, d].entries()) state[i] = (state[i] + value) | 0;
    }
    const out = new Uint8Array(16), view = new DataView(out.buffer);
    state.forEach((v, i) => view.setUint32(i * 4, v, true));
    return out;
  };
  E.rc4 = function (key) {
    if (!key.length || key.length > 256) B.fail('ENCRYPTION_KEY', 'Invalid RC4 key length');
    const state = Uint8Array.from({ length: 256 }, (_, i) => i);
    let x = 0, y = 0;
    for (let i = 0, j = 0; i < 256; i++) {
      j = (j + state[i] + key[i % key.length]) & 255;
      [state[i], state[j]] = [state[j], state[i]];
    }
    return function (bytes) {
      new B.Reader(bytes);
      const out = new Uint8Array(bytes.length);
      for (let i = 0; i < bytes.length; i++) {
        x = (x + 1) & 255;
        y = (y + state[x]) & 255;
        [state[x], state[y]] = [state[y], state[x]];
        out[i] = bytes[i] ^ state[(state[x] + state[y]) & 255];
      }
      return out;
    };
  };
  // XOR95 uses single-byte password data. Both the verifier and the array use
  // exactly the same bytes. The caller chooses the encoding whose verifier matches.
  E.xorKey = function (password, rotation = 2) {
    const end = password.indexOf(0), n = Math.min(end < 0 ? password.length : end, 15);
    let key = 0, base = 0x8000, last = 0xffff, hash = n ? n ^ 0xce4b : 0;
    for (let i = n - 1; i >= 0; i--) {
      let c = password[i] & 0x7f;
      for (let bit = 0; bit < 8; bit++) {
        base = rotate(base, 1, 16);
        if (base & 1) base ^= 0x1020;
        if (c & 1) key ^= base;
        c >>>= 1;
        last = rotate(last, 1, 16);
        if (last & 1) last ^= 0x1020;
      }
    }
    key = n ? key ^ last : 0;
    for (let i = 0; i < n; i++) hash ^= rotate(password[i], (i + 1) % 15, 15);
    const array = new Uint8Array(16);
    for (let i = 0; i < 16; i++) {
      const value = i < n ? password[i] : PADDING[i - n] || 0;
      array[i] = rotate(value ^ (i & 1 ? key >>> 8 : key & 255), rotation, 8);
    }
    return { key, hash, array };
  };
  E.xorBytes = function (bytes, array, offset, word = false) {
    new B.Reader(bytes);
    const out = new Uint8Array(bytes.length);
    for (let i = 0; i < bytes.length; i++) {
      const value = bytes[i], mask = array[(offset + i) & 15];
      out[i] = word ? (value && (value ^ mask) ? value ^ mask : value) : rotate(value, 3, 8) ^ mask;
    }
    return out;
  };
})(typeof window !== 'undefined' ? window : globalThis);
