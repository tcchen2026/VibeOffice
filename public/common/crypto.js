/* VibeOffice — password-protected Office files (ECMA-376 Part 2 "Office Document Cryptography"):
 * .docx, .xlsx and .pptx alike. An encrypted package is an OLE compound file holding EncryptionInfo and
 * EncryptedPackage streams. Supported: Agile encryption (Office 2010 and later, AES with SHA-1/SHA-256/
 * SHA-384/SHA-512) and Standard encryption (Office 2007, AES-ECB with SHA-1). The password hashing loop uses
 * sha.js when it is loaded (fast), else WebCrypto; AES is done here so that unpadded CBC and ECB blocks can
 * be decrypted directly. Nothing leaves the browser. */
(function (root) {
  'use strict';
  const L = root.L || (root.L = {});
  const OC = (L.officeCrypto = {});

  /* ================= compound file (CFB) reader ================= */
  function cfb(u8) {
    const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
    if (dv.getUint32(0, true) !== 0xe011cfd0 || dv.getUint32(4, true) !== 0xe11ab1a1) throw new Error('not a compound file');
    const ssz = 1 << dv.getUint16(0x1e, true), mssz = 1 << dv.getUint16(0x20, true);
    const nFat = dv.getUint32(0x2c, true), dirStart = dv.getUint32(0x30, true), cutoff = dv.getUint32(0x38, true);
    const miniFatStart = dv.getUint32(0x3c, true), difatStart = dv.getUint32(0x44, true), nDifat = dv.getUint32(0x48, true);
    const off = (n) => (n + 1) * ssz; /* sector n follows the header sector */
    /* FAT sector list: 109 in the header, the rest in the DIFAT chain */
    const fatSecs = [];
    for (let i = 0; i < 109 && fatSecs.length < nFat; i++) { const v = dv.getUint32(0x4c + i * 4, true); if (v < 0xfffffffa) fatSecs.push(v); }
    let d = difatStart, guard = 0;
    while (d < 0xfffffffa && guard++ < nDifat + 1 && fatSecs.length < nFat) {
      const base = off(d), per = ssz / 4 - 1;
      for (let i = 0; i < per && fatSecs.length < nFat; i++) { const v = dv.getUint32(base + i * 4, true); if (v < 0xfffffffa) fatSecs.push(v); }
      d = dv.getUint32(base + per * 4, true);
    }
    const fat = new Uint32Array(fatSecs.length * (ssz / 4));
    fatSecs.forEach((s, i) => { const b = off(s); for (let k = 0; k < ssz / 4; k++) fat[i * (ssz / 4) + k] = dv.getUint32(b + k * 4, true); });
    const chain = (start, limit) => {
      const out = [];
      let s = start;
      while (s < 0xfffffffa && out.length <= limit) { out.push(s); s = fat[s]; }
      return out;
    };
    const readChain = (start, size) => {
      const secs = chain(start, Math.ceil((size || 1e9) / ssz) + 1);
      const buf = new Uint8Array(secs.length * ssz);
      secs.forEach((s, i) => buf.set(u8.subarray(off(s), Math.min(u8.length, off(s) + ssz)), i * ssz));
      return size != null ? buf.subarray(0, size) : buf;
    };
    /* directory */
    const dirBytes = readChain(dirStart, null);
    const entries = [];
    for (let p = 0; p + 128 <= dirBytes.length; p += 128) {
      const nameLen = dirBytes[p + 64] | (dirBytes[p + 65] << 8);
      const type = dirBytes[p + 66];
      if (!type) { entries.push(null); continue; }
      const name = new TextDecoder('utf-16le').decode(dirBytes.subarray(p, p + Math.max(0, nameLen - 2)));
      const ed = new DataView(dirBytes.buffer, dirBytes.byteOffset + p, 128);
      entries.push({ name, type, start: ed.getUint32(116, true), size: ed.getUint32(120, true) });
    }
    const rootEntry = entries[0];   // the compound file's root storage
    let mini = null, miniFat = null;
    const getMini = () => {
      if (!mini) {
        mini = readChain(rootEntry.start, rootEntry.size);
        const mf = miniFatStart < 0xfffffffa ? readChain(miniFatStart, null) : new Uint8Array(0);
        miniFat = new Uint32Array(mf.buffer, mf.byteOffset, mf.byteLength >> 2);
      }
    };
    return {
      stream(name) {
        const e = entries.find((x) => x && x.type === 2 && x.name.toLowerCase() === name.toLowerCase());
        if (!e) return null;
        if (e.size < cutoff) {
          getMini();
          const out = new Uint8Array(e.size);
          let s = e.start, pos = 0, g = 0;
          while (s < 0xfffffffa && pos < e.size && g++ < 1e7) { out.set(mini.subarray(s * mssz, s * mssz + Math.min(mssz, e.size - pos)), pos); pos += mssz; s = miniFat[s]; }
          return out;
        }
        return readChain(e.start, e.size);
      },
    };
  }

  /* ================= AES (decryption) ================= */
  const SBOX = new Uint8Array(256), INV = new Uint8Array(256);
  (function () {
    const rotl = (x, s) => ((x << s) | (x >> (8 - s))) & 0xff;
    let p = 1, q = 1;
    do {
      p = p ^ ((p << 1) & 0xff) ^ (p & 0x80 ? 0x1b : 0);
      q ^= q << 1; q ^= q << 2; q ^= q << 4; q &= 0xff; if (q & 0x80) q ^= 0x09;
      SBOX[p] = (q ^ rotl(q, 1) ^ rotl(q, 2) ^ rotl(q, 3) ^ rotl(q, 4) ^ 0x63) & 0xff;
    } while (p !== 1);
    SBOX[0] = 0x63;
    for (let i = 0; i < 256; i++) INV[SBOX[i]] = i;
  })();
  const xtime = (b) => ((b << 1) ^ (b & 0x80 ? 0x1b : 0)) & 0xff;
  const gmul = (a, b) => { let r = 0; while (b) { if (b & 1) r ^= a; a = xtime(a); b >>= 1; } return r; };
  const M9 = new Uint8Array(256), M11 = new Uint8Array(256), M13 = new Uint8Array(256), M14 = new Uint8Array(256);
  for (let i = 0; i < 256; i++) { M9[i] = gmul(i, 9); M11[i] = gmul(i, 11); M13[i] = gmul(i, 13); M14[i] = gmul(i, 14); }
  function expandKey(key) {
    const Nk = key.length / 4, Nr = Nk + 6, W = new Uint8Array(16 * (Nr + 1));
    W.set(key);
    let rcon = 1;
    for (let i = Nk; i < 4 * (Nr + 1); i++) {
      let t0 = W[4 * i - 4], t1 = W[4 * i - 3], t2 = W[4 * i - 2], t3 = W[4 * i - 1];
      if (i % Nk === 0) { const t = t0; t0 = SBOX[t1] ^ rcon; t1 = SBOX[t2]; t2 = SBOX[t3]; t3 = SBOX[t]; rcon = xtime(rcon); }
      else if (Nk > 6 && i % Nk === 4) { t0 = SBOX[t0]; t1 = SBOX[t1]; t2 = SBOX[t2]; t3 = SBOX[t3]; }
      W[4 * i] = W[4 * (i - Nk)] ^ t0; W[4 * i + 1] = W[4 * (i - Nk) + 1] ^ t1; W[4 * i + 2] = W[4 * (i - Nk) + 2] ^ t2; W[4 * i + 3] = W[4 * (i - Nk) + 3] ^ t3;
    }
    return { W, Nr };
  }
  function decryptBlock(ks, inp, io, out, oo, s, t) {
    const { W, Nr } = ks;
    for (let i = 0; i < 16; i++) s[i] = inp[io + i] ^ W[Nr * 16 + i];
    for (let round = Nr - 1; round >= 0; round--) {
      /* inverse shift rows + inverse substitution */
      for (let c = 0; c < 4; c++) for (let r = 0; r < 4; r++) t[r + 4 * c] = INV[s[r + 4 * ((c - r + 4) & 3)]];
      const k = round * 16;
      if (round === 0) { for (let i = 0; i < 16; i++) out[oo + i] = t[i] ^ W[k + i]; break; }
      for (let c = 0; c < 4; c++) {
        const a0 = t[4 * c] ^ W[k + 4 * c], a1 = t[4 * c + 1] ^ W[k + 4 * c + 1], a2 = t[4 * c + 2] ^ W[k + 4 * c + 2], a3 = t[4 * c + 3] ^ W[k + 4 * c + 3];
        s[4 * c] = M14[a0] ^ M11[a1] ^ M13[a2] ^ M9[a3];
        s[4 * c + 1] = M9[a0] ^ M14[a1] ^ M11[a2] ^ M13[a3];
        s[4 * c + 2] = M13[a0] ^ M9[a1] ^ M14[a2] ^ M11[a3];
        s[4 * c + 3] = M11[a0] ^ M13[a1] ^ M9[a2] ^ M14[a3];
      }
    }
  }
  /** AES-CBC (iv given) or AES-ECB (iv null) decryption of whole blocks, no padding */
  OC.aesDecrypt = function (key, data, iv) {
    const ks = expandKey(key);
    const n = data.length - (data.length % 16);
    const out = new Uint8Array(n);
    const s = new Uint8Array(16), t = new Uint8Array(16);
    for (let p = 0; p < n; p += 16) {
      decryptBlock(ks, data, p, out, p, s, t);
      if (iv) { const prev = p ? data.subarray(p - 16, p) : iv; for (let i = 0; i < 16; i++) out[p + i] ^= prev[i]; }
    }
    return out;
  };

  /* ================= hashing ================= */
  const HASH = { SHA1: 'SHA-1', 'SHA-1': 'SHA-1', SHA256: 'SHA-256', 'SHA-256': 'SHA-256', SHA384: 'SHA-384', 'SHA-384': 'SHA-384', SHA512: 'SHA-512', 'SHA-512': 'SHA-512' };
  const digest = async (alg, ...parts) => {
    let len = 0; for (const p of parts) len += p.length;
    const buf = new Uint8Array(len); let o = 0; for (const p of parts) { buf.set(p, o); o += p.length; }
    return new Uint8Array(await crypto.subtle.digest(alg, buf));
  };
  const u32 = (n) => new Uint8Array([n & 255, (n >>> 8) & 255, (n >>> 16) & 255, (n >>> 24) & 255]);
  const utf16 = (s) => { const b = new Uint8Array(s.length * 2); for (let i = 0; i < s.length; i++) { b[2 * i] = s.charCodeAt(i) & 255; b[2 * i + 1] = s.charCodeAt(i) >> 8; } return b; };
  const b64 = (s) => { const bin = atob(String(s || '').replace(/\s+/g, '')); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
  const fix = (h, n, pad) => { if (h.length >= n) return h.slice(0, n); const o = new Uint8Array(n).fill(pad); o.set(h); return o; };
  const same = (a, b) => a.length === b.length && a.every((x, i) => x === b[i]);
  /** the password hash iterated spinCount times (the slow part, by design) */
  async function spin(alg, salt, password, count) {
    let h = await digest(alg, salt, utf16(password));
    /* the synchronous hashes in sha.js run the loop several times faster than a promise per round */
    const f = L.sha && L.sha.get(alg);
    if (f) {
      const b = new Uint8Array(4 + h.length);
      for (let i = 0; i < count; i++) { b[0] = i & 255; b[1] = (i >>> 8) & 255; b[2] = (i >>> 16) & 255; b[3] = i >>> 24; b.set(h, 4); h = f(b); }
      return h;
    }
    for (let i = 0; i < count; i++) h = await digest(alg, u32(i), h);
    return h;
  }

  /* ================= the two encryption schemes ================= */
  async function agile(info, pkg, password) {
    const xml = new TextDecoder('utf-8').decode(info.subarray(8));
    const x = L.xml ? L.xml.parse(xml).parentNode : new DOMParser().parseFromString(xml, 'application/xml');
    const byName = (n) => Array.from(x.getElementsByTagName('*')).find((e) => e.localName === n);
    const kd = byName('keyData');
    const ek = Array.from(x.getElementsByTagName('*')).find((e) => e.localName === 'encryptedKey' && e.getAttribute('spinCount'));
    if (!kd || !ek) throw new Error('unsupported');
    /* Office writes AES with SHA-1 / SHA-2; other ciphers (DES, 3DES, RC2) and MD5 are not supported */
    for (const e of [ek, kd]) {
      if (!/^AES$/i.test(e.getAttribute('cipherAlgorithm') || 'AES')) throw new Error('unsupported');
      if (e.getAttribute('hashAlgorithm') && !HASH[e.getAttribute('hashAlgorithm')]) throw new Error('unsupported');
    }
    const alg = HASH[ek.getAttribute('hashAlgorithm')] || 'SHA-1';
    const keyBytes = (+ek.getAttribute('keyBits') || 128) / 8;
    const salt = b64(ek.getAttribute('saltValue'));
    const H = await spin(alg, salt, password, +ek.getAttribute('spinCount') || 100000);
    const derive = async (block) => fix(await digest(alg, H, block), keyBytes, 0x36);
    const B_IN = new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]);
    const B_VAL = new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]);
    const B_KEY = new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]);
    /* password check: the verifier's hash must match */
    const vin = OC.aesDecrypt(await derive(B_IN), b64(ek.getAttribute('encryptedVerifierHashInput')), salt);
    const vval = OC.aesDecrypt(await derive(B_VAL), b64(ek.getAttribute('encryptedVerifierHashValue')), salt);
    const vin2 = vin.slice(0, +ek.getAttribute('saltSize') || 16);
    const vh = await digest(alg, vin2);
    if (!same(vh, vval.slice(0, vh.length))) { const e = new Error('badpass'); e.code = 'badpass'; throw e; }
    const secret = OC.aesDecrypt(await derive(B_KEY), b64(ek.getAttribute('encryptedKeyValue')), salt).slice(0, (+kd.getAttribute('keyBits') || 128) / 8);
    /* the package: an 8-byte size, then 4096-byte segments each with its own IV */
    const dv = new DataView(pkg.buffer, pkg.byteOffset, pkg.byteLength);
    const size = dv.getUint32(0, true) + dv.getUint32(4, true) * 4294967296;
    const kdSalt = b64(kd.getAttribute('saltValue'));
    const kdAlg = HASH[kd.getAttribute('hashAlgorithm')] || alg;
    const bs = +kd.getAttribute('blockSize') || 16;
    const out = new Uint8Array(Math.ceil((pkg.length - 8) / 16) * 16);
    for (let seg = 0, p = 8; p < pkg.length; seg++, p += 4096) {
      const iv = fix(await digest(kdAlg, kdSalt, u32(seg)), bs, 0x36);
      const chunk = pkg.subarray(p, Math.min(pkg.length, p + 4096));
      out.set(OC.aesDecrypt(secret, chunk, iv), p - 8);
    }
    return out.subarray(0, size);
  }
  async function standard(info, pkg, password) {
    const dv = new DataView(info.buffer, info.byteOffset, info.byteLength);
    const hSize = dv.getUint32(8, true);
    const h = 12;
    const algID = dv.getUint32(h + 8, true), keyBits = dv.getUint32(h + 16, true) || 128;
    if (algID && algID !== 0x660e && algID !== 0x660f && algID !== 0x6610) throw new Error('unsupported'); /* RC4 is not supported */
    const v = h + hSize;
    const saltSize = dv.getUint32(v, true);
    const salt = info.slice(v + 4, v + 4 + saltSize);
    const encVerifier = info.slice(v + 4 + saltSize, v + 20 + saltSize);
    const encVerifierHash = info.slice(v + 24 + saltSize, v + 24 + saltSize + 32);
    let hh = await spin('SHA-1', salt, password, 50000);
    hh = await digest('SHA-1', hh, u32(0));
    const x1 = new Uint8Array(64).fill(0x36), x2 = new Uint8Array(64).fill(0x5c);
    for (let i = 0; i < hh.length; i++) { x1[i] ^= hh[i]; x2[i] ^= hh[i]; }
    const k = new Uint8Array(40); k.set(await digest('SHA-1', x1)); k.set(await digest('SHA-1', x2), 20);
    const key = k.slice(0, keyBits / 8);
    const ver = OC.aesDecrypt(key, encVerifier, null);
    const verHash = OC.aesDecrypt(key, encVerifierHash, null).slice(0, 20);
    if (!same(await digest('SHA-1', ver), verHash)) { const e = new Error('badpass'); e.code = 'badpass'; throw e; }
    const pdv = new DataView(pkg.buffer, pkg.byteOffset, pkg.byteLength);
    const size = pdv.getUint32(0, true) + pdv.getUint32(4, true) * 4294967296;
    return OC.aesDecrypt(key, pkg.subarray(8), null).subarray(0, size);
  }

  /* ================= writing: Agile encryption (AES-256, SHA-512) in a compound file ================= */
  const hex = (h) => new Uint8Array(h.match(/../g).map((x) => parseInt(x, 16)));
  /* the \x06DataSpaces streams Office writes for its standard encryption transform (constant) */
  const DS = {
    Version: '3c0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e004400610074006100530070006100630065007300010000000100000001000000',
    DataSpaceMap: '08000000010000006800000001000000000000002000000045006e0063007200790070007400650064005000610063006b00610067006500320000005300740072006f006e00670045006e006300720079007000740069006f006e004400610074006100530070006100630065000000',
    StrongEncryptionDataSpace: '0800000001000000320000005300740072006f006e00670045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d000000',
    Primary: '58000000010000004c0000007b00460046003900410033004600300033002d0035003600450046002d0034003600310033002d0042004400440035002d003500410034003100430031004400300037003200340036007d004e0000004d006900630072006f0073006f00660074002e0043006f006e007400610069006e00650072002e0045006e006300720079007000740069006f006e005400720061006e00730066006f0072006d00000001000000010000000100000000000000000000000000000004000000',
  };
  /** a version-3 compound file with Office's directory layout for an encrypted package */
  function cfbWrite(info, pkg) {
    const SS = 512, MS = 64, CUT = 4096, END = 0xfffffffe, FREE = 0xffffffff, FATSEC = 0xfffffffd, DIFSEC = 0xfffffffc, NONE = 0xffffffff;
    /* entries: name, type (1 storage, 2 stream, 5 root), data, left, right, child */
    const E = [
      { n: 'Root Entry', t: 5, c: 1 },
      { n: 'EncryptionInfo', t: 2, d: info, l: 2, r: 3 },
      { n: '\u0006DataSpaces', t: 1, c: 5 },
      { n: 'EncryptedPackage', t: 2, d: pkg },
      { n: 'Version', t: 2, d: hex(DS.Version) },
      { n: 'DataSpaceMap', t: 2, d: hex(DS.DataSpaceMap), l: 4, r: 6 },
      { n: 'DataSpaceInfo', t: 1, c: 7, r: 8 },
      { n: 'StrongEncryptionDataSpace', t: 2, d: hex(DS.StrongEncryptionDataSpace) },
      { n: 'TransformInfo', t: 1, c: 9 },
      { n: 'StrongEncryptionTransform', t: 1, c: 10 },
      { n: '\u0006Primary', t: 2, d: hex(DS.Primary) },
    ];
    /* small streams live in the mini stream */
    const miniParts = [];
    let miniLen = 0;
    for (const e of E) if (e.t === 2 && e.d.length < CUT) { e.mini = true; e.start = e.d.length ? miniLen / MS : END; miniParts.push(e); miniLen += Math.ceil(e.d.length / MS) * MS; }
    const nMini = miniLen / MS;
    const miniFatSecs = Math.ceil((nMini * 4) / SS);
    const miniStreamSecs = Math.ceil(miniLen / SS);
    const dirSecs = Math.ceil(E.length / 4);
    const big = E.filter((e) => e.t === 2 && !e.mini);
    const bigSecs = big.reduce((a, e) => a + Math.ceil(e.d.length / SS), 0);
    let fatSecs = 1, difSecs = 0;
    for (let k = 0; k < 20; k++) {
      const total = difSecs + fatSecs + dirSecs + miniFatSecs + miniStreamSecs + bigSecs;
      const f = Math.ceil(total / (SS / 4)), d = f > 109 ? Math.ceil((f - 109) / (SS / 4 - 1)) : 0;
      if (f === fatSecs && d === difSecs) break;
      fatSecs = f; difSecs = d;
    }
    const total = difSecs + fatSecs + dirSecs + miniFatSecs + miniStreamSecs + bigSecs;
    const out = new Uint8Array(SS * (1 + total));
    const dv = new DataView(out.buffer);
    const fat = new Uint32Array(fatSecs * (SS / 4)).fill(FREE);
    let next = 0;
    const run = (count, mark) => { const start = next; for (let i = 0; i < count; i++) fat[next + i] = mark != null ? mark : i === count - 1 ? END : next + i + 1; next += count; return count ? start : END; };
    const difStart = run(difSecs, DIFSEC);
    const fatStart = run(fatSecs, FATSEC);
    const dirStart = run(dirSecs);
    const miniFatStart = run(miniFatSecs);
    const miniStreamStart = run(miniStreamSecs);
    for (const e of big) { const n = Math.ceil(e.d.length / SS); e.start = run(n); }
    const at = (sec) => SS * (sec + 1);
    /* header */
    out.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1], 0);
    dv.setUint16(0x18, 0x3e, true); dv.setUint16(0x1a, 3, true); dv.setUint16(0x1c, 0xfffe, true);
    dv.setUint16(0x1e, 9, true); dv.setUint16(0x20, 6, true);
    dv.setUint32(0x2c, fatSecs, true); dv.setUint32(0x30, dirStart, true); dv.setUint32(0x38, CUT, true);
    dv.setUint32(0x3c, miniFatSecs ? miniFatStart : END, true); dv.setUint32(0x40, miniFatSecs, true);
    dv.setUint32(0x44, difSecs ? difStart : END, true); dv.setUint32(0x48, difSecs, true);
    for (let i = 0; i < 109; i++) dv.setUint32(0x4c + 4 * i, i < fatSecs ? fatStart + i : FREE, true);
    /* DIFAT sectors for FAT sectors beyond the first 109 */
    for (let k = 0; k < difSecs; k++) {
      const b = at(difStart + k), per = SS / 4 - 1;
      for (let i = 0; i < per; i++) { const idx = 109 + k * per + i; dv.setUint32(b + 4 * i, idx < fatSecs ? fatStart + idx : FREE, true); }
      dv.setUint32(b + 4 * per, k + 1 < difSecs ? difStart + k + 1 : END, true);
    }
    /* FAT */
    for (let i = 0; i < fat.length; i++) dv.setUint32(at(fatStart) + 4 * i, fat[i], true);
    /* mini FAT and mini stream */
    const mf = new Uint32Array(miniFatSecs * (SS / 4)).fill(FREE);
    for (const e of miniParts) {
      const n = Math.ceil(e.d.length / MS);
      for (let i = 0; i < n; i++) mf[e.start + i] = i === n - 1 ? END : e.start + i + 1;
      out.set(e.d, at(miniStreamStart) + e.start * MS);
    }
    for (let i = 0; i < mf.length; i++) dv.setUint32(at(miniFatStart) + 4 * i, mf[i], true);
    for (const e of big) out.set(e.d, at(e.start));
    /* directory */
    for (let i = 0; i < dirSecs * 4; i++) {
      const b = at(dirStart) + 128 * i;
      const e = E[i];
      if (!e) { dv.setUint32(b + 68, NONE, true); dv.setUint32(b + 72, NONE, true); dv.setUint32(b + 76, NONE, true); continue; }
      for (let k = 0; k < e.n.length; k++) dv.setUint16(b + 2 * k, e.n.charCodeAt(k), true);
      dv.setUint16(b + 64, (e.n.length + 1) * 2, true);
      out[b + 66] = e.t; out[b + 67] = 1;
      dv.setUint32(b + 68, e.l != null ? e.l : NONE, true); dv.setUint32(b + 72, e.r != null ? e.r : NONE, true); dv.setUint32(b + 76, e.c != null ? e.c : NONE, true);
      if (e.t === 5) { dv.setUint32(b + 116, miniLen ? miniStreamStart : END, true); dv.setUint32(b + 120, miniLen, true); }
      else if (e.t === 2) { dv.setUint32(b + 116, e.d.length ? e.start : END, true); dv.setUint32(b + 120, e.d.length, true); }
    }
    return out;
  }
  const aesCbcEncrypt = async (key, iv, data) => {
    const k = await crypto.subtle.importKey('raw', key, 'AES-CBC', false, ['encrypt']);
    const padded = data.length % 16 ? (() => { const p = new Uint8Array(Math.ceil(data.length / 16) * 16); p.set(data); return p; })() : data;
    return new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-CBC', iv }, k, padded)).slice(0, padded.length); /* drop WebCrypto's padding block */
  };
  const b64e = (u) => { let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); };
  /** encrypt package bytes with a password to open, as Office 2013+ does (AES-256, SHA-512); resolves to the file bytes */
  OC.encrypt = async function (pkgBytes, password) {
    if (!(root.crypto && root.crypto.subtle)) throw new Error('Saving with a password needs a secure (https) page.');
    const rnd = (n) => crypto.getRandomValues(new Uint8Array(n));
    const alg = 'SHA-512', keyDataSalt = rnd(16), pwSalt = rnd(16), secret = rnd(32), verifier = rnd(16), hmacKey = rnd(64);
    /* the package: an 8-byte size, then 4096-byte segments, each with its own IV */
    const enc = new Uint8Array(8 + Math.ceil(pkgBytes.length / 16) * 16);
    new DataView(enc.buffer).setUint32(0, pkgBytes.length, true);
    for (let seg = 0, p = 0; p < pkgBytes.length; seg++, p += 4096) {
      const iv = fix(await digest(alg, keyDataSalt, u32(seg)), 16, 0x36);
      enc.set(await aesCbcEncrypt(secret, iv, pkgBytes.subarray(p, Math.min(pkgBytes.length, p + 4096))), 8 + p);
    }
    /* the password's key encryptor */
    const H = await spin(alg, pwSalt, password, 100000);
    const derive = async (block) => fix(await digest(alg, H, block), 32, 0x36);
    const encIn = await aesCbcEncrypt(await derive(new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79])), pwSalt, verifier);
    const encVal = await aesCbcEncrypt(await derive(new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e])), pwSalt, await digest(alg, verifier));
    const encKey = await aesCbcEncrypt(await derive(new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6])), pwSalt, secret);
    /* data integrity: an HMAC of the encrypted package */
    const ivHK = fix(await digest(alg, keyDataSalt, new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6])), 16, 0x36);
    const ivHV = fix(await digest(alg, keyDataSalt, new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33])), 16, 0x36);
    const hk = await crypto.subtle.importKey('raw', hmacKey, { name: 'HMAC', hash: alg }, false, ['sign']);
    const mac = new Uint8Array(await crypto.subtle.sign('HMAC', hk, enc));
    const encHmacKey = await aesCbcEncrypt(secret, ivHK, hmacKey), encHmacVal = await aesCbcEncrypt(secret, ivHV, mac);
    const xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate">' +
      `<keyData saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64e(keyDataSalt)}"/>` +
      `<dataIntegrity encryptedHmacKey="${b64e(encHmacKey)}" encryptedHmacValue="${b64e(encHmacVal)}"/>` +
      '<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">' +
      `<p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512" saltValue="${b64e(pwSalt)}" encryptedVerifierHashInput="${b64e(encIn)}" encryptedVerifierHashValue="${b64e(encVal)}" encryptedKeyValue="${b64e(encKey)}"/>` +
      '</keyEncryptor></keyEncryptors></encryption>';
    const xb = new TextEncoder().encode(xml);
    const info = new Uint8Array(8 + xb.length);
    info.set([4, 0, 4, 0, 0x40, 0, 0, 0]);
    info.set(xb, 8);
    return cfbWrite(info, enc);
  };

  /** decrypt an encrypted Office package; resolves to the inner .xlsx bytes */
  OC.decrypt = async function (u8, password) {
    if (!(root.crypto && root.crypto.subtle)) throw new Error('Opening password-protected files needs a secure (https) page.');
    const c = cfb(u8);
    const info = c.stream('EncryptionInfo'), pkg = c.stream('EncryptedPackage');
    if (!info || !pkg) throw new Error('unsupported');
    const major = info[0] | (info[1] << 8), minor = info[2] | (info[3] << 8);
    if (major === 4 && minor === 4) return agile(info, pkg, password);
    if ((major === 3 || major === 4) && minor === 2) return standard(info, pkg, password);
    throw new Error('unsupported');
  };
})(typeof window !== 'undefined' ? window : globalThis);
