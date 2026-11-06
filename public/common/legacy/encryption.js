/* Binary RC4 and RC4 CryptoAPI headers, password verifiers and stream keys.
 * Derived from Apache POI poi/src/main/java/org/apache/poi/poifs/crypt/
 * binaryrc4/BinaryRC4Decryptor.java, cryptoapi/CryptoAPIDecryptor.java (Apache-2.0).
 * MS-OFFCRYPTO sections 2.3.5–2.3.7. Optional encrypted streams remain lazy. */
(function (root) {
  'use strict';
  const L = root.L, B = L.legacy, E = B.encryption, N = s => s;
  const DEFAULT_PASSWORD = 'VelvetSweatshop';
  const RC4 = 0x6801, SHA1 = 0x8004, EXTERNAL = 0x10, AES = 0x20;
  const blockBytes = block => {
    const bytes = new Uint8Array(4);
    new DataView(bytes.buffer).setUint32(0, block, true);
    return bytes;
  };
  const equal = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);
  E.header = function (bytes) {
    const r = new B.Reader(bytes, 'Encryption header'), major = r.u16(0), minor = r.u16(2);
    if (major === 1 && minor === 1) {
      return { mode: 'rc4', salt: r.sub(4, 16), verifier: r.sub(20, 16), hash: r.sub(36, 16), length: 52 };
    }
    if (![2, 3, 4].includes(major) || minor !== 2) {
      B.fail('ENCRYPTION_UNSUPPORTED', N('This file uses an unsupported password protection method'));
    }
    const size = r.u32(8), header = r.reader(12, size), flags = header.u32(0);
    const cipher = header.u32(8), hash = header.u32(12), keyBits = header.u32(16) || 40;
    header.check(0, 32);
    if ((flags & (EXTERNAL | AES)) || ![0, RC4].includes(cipher) || ![0, SHA1].includes(hash)
      || keyBits < 40 || keyBits > 128 || keyBits % 8) {
      B.fail('ENCRYPTION_UNSUPPORTED', N('This file uses an unsupported password protection method'));
    }
    const at = 12 + size;
    if (r.u32(at) !== 16) B.fail('ENCRYPTION_HEADER', N('The file has a damaged password verifier'));
    // Reserved fields, provider names and the hash-size hint do not affect RC4.
    // Office/POI use the fixed SHA-1 result size; do not reject an unused hint.
    return { mode: 'cryptoapi', flags, keyBits, salt: r.sub(at + 4, 16), verifier: r.sub(at + 20, 16),
      hash: r.sub(at + 40, 20), length: at + 60 };
  };
  function xorCandidates(password, codepage) {
    const text = password.slice(0, 15), candidates = [];
    // MS-OFFCRYPTO 2.3.7.4 permits the low/high-byte method as well as the
    // application's ANSI code page. The matching verifier decides which was used.
    candidates.push(Uint8Array.from(text.split(''), c => (c.charCodeAt(0) & 255) || (c.charCodeAt(0) >>> 8)));
    for (const cp of new Set([codepage, 1252].filter(Boolean))) {
      const inverse = new Map();
      try {
        for (let i = 0; i < 256; i++) inverse.set(B.decode(Uint8Array.of(i), cp), i);
      } catch (e) {
        if (e.name !== 'LegacyError') throw e;
        continue;
      }
      candidates.push(Uint8Array.from(text.split(''), c => inverse.get(c) ?? 63));
    }
    return candidates;
  }
  E.unlock = function (info, password, where, codepage) {
    const supplied = password != null;
    const text = String(supplied ? password : DEFAULT_PASSWORD).slice(0, 255);
    if (info.mode === 'xor') {
      for (const bytes of xorCandidates(text, codepage)) {
        const key = E.xorKey(bytes, info.word ? 7 : 2);
        if (key.key === info.key && key.hash === info.hash) return { ...info, array: key.array };
      }
    } else {
      const hash = info.mode === 'rc4' ? E.md5 : L.sha.sha1;
      let secret;
      if (info.mode === 'rc4') {
        const first = hash(E.utf16(text)).subarray(0, 5), input = E.join(first, info.salt);
        const repeated = new Uint8Array(input.length * 16);
        for (let i = 0; i < 16; i++) repeated.set(input, i * input.length);
        secret = hash(repeated).subarray(0, 5);
      } else secret = hash(E.join(info.salt, E.utf16(text)));
      const key = block => {
        const digest = hash(E.join(secret, blockBytes(block)));
        if (info.mode === 'rc4') return digest;
        const bytes = new Uint8Array(info.keyBits === 40 ? 16 : info.keyBits / 8);
        bytes.set(digest.subarray(0, info.keyBits / 8));
        return bytes;
      };
      const cipher = E.rc4(key(0)), verifier = cipher(info.verifier), verifierHash = cipher(info.hash);
      if (equal(hash(verifier), verifierHash)) return { ...info, key };
    }
    B.fail(supplied ? 'WRONG_PASSWORD' : 'PASSWORD_REQUIRED', supplied
      ? N('The password is incorrect. Please try again.') : N('This file needs a password'), where);
  };
  E.blocks = function (bytes, session, size) {
    new B.Reader(bytes);
    const out = new Uint8Array(bytes.length);
    for (let p = 0; p < bytes.length; p += size) {
      out.set(E.rc4(session.key(p / size))(bytes.subarray(p, p + size)), p);
    }
    return out;
  };
  // A stream wrapper over the existing CFB reader, not a second storage reader.
  // A corrupt optional stream is read and rejected only when its caller asks for it.
  E.overlay = function (cfb, read) {
    const cache = new Map();
    return { ...cfb, stream(name) {
      const key = name.toLowerCase();
      if (!cache.has(key)) cache.set(key, read(name, key));
      return cache.get(key);
    } };
  };
  E.summary = function (cfb, session, streamName) {
    if (session.mode !== 'cryptoapi') return cfb;
    let descriptors, source;
    const directory = () => {
      if (descriptors) return descriptors;
      source = cfb.stream(streamName);
      if (!source) return descriptors = new Map();
      const raw = new B.Reader(source, streamName);
      const head = new B.Reader(E.rc4(session.key(0))(raw.sub(0, 8)), streamName);
      const offset = head.u32(0), size = head.u32(4);
      const list = new B.Reader(E.rc4(session.key(0))(raw.sub(offset, size)), streamName + ' directory');
      const count = list.u32(0), entries = new Map();
      if (count > Math.floor((list.length - 4) / 18)) {
        B.fail('ENCRYPTION_HEADER', N('The encrypted file properties are damaged'), streamName);
      }
      for (let p = 4, i = 0; i < count; i++) {
        const entry = { offset: list.u32(p), size: list.u32(p + 4), block: list.u16(p + 8) };
        const chars = list.u8(p + 10), name = B.latin16(list.sub(p + 16, chars * 2));
        // Decode code units directly: descriptor names can start with control characters.
        let decoded = '';
        for (let j = 0; j < name.length; j += 2) decoded += String.fromCharCode(name.charCodeAt(j)
          | (name.charCodeAt(j + 1) << 8));
        list.check(p, 18 + chars * 2);
        if (!entries.has(decoded.toLowerCase())) entries.set(decoded.toLowerCase(), entry);
        p += 18 + chars * 2;
      }
      return descriptors = entries;
    };
    return E.overlay(cfb, (name, key) => {
      if (!key.startsWith('\x05')) return cfb.stream(name);
      const entry = directory().get(key);
      if (!entry) return cfb.stream(name);
      const raw = new B.Reader(source, streamName);
      return E.rc4(session.key(entry.block))(raw.sub(entry.offset, entry.size));
    });
  };
})(typeof window !== 'undefined' ? window : globalThis);
