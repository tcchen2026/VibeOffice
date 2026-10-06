/* SHA implementations in public/common/sha.js against WebCrypto. Usage: node sha.test.js */
globalThis.L = globalThis.L || {};
require('../../../public/common/sha.js');
const L = globalThis.L;
const hex = (u) => Buffer.from(u).toString('hex');
(async () => {
  let pass = 0, fail = 0;
  const lens = [0, 1, 3, 55, 56, 63, 64, 65, 111, 112, 119, 120, 127, 128, 129, 200, 1000, 5000];
  for (const alg of ['SHA-1', 'SHA-256', 'SHA-384', 'SHA-512']) {
    for (const n of lens) {
      const m = new Uint8Array(n); for (let i = 0; i < n; i++) m[i] = (i * 131 + n * 7) & 255;
      const want = hex(new Uint8Array(await crypto.subtle.digest(alg, m)));
      const got = hex(L.sha.get(alg)(m));
      if (want === got) pass++; else { fail++; console.log('FAIL', alg, n, want, got); }
    }
  }
  /* the Office spin loop: 100,000 rounds of SHA-512 */
  const t0 = Date.now();
  let h = L.sha.sha512(new Uint8Array(40));
  const b = new Uint8Array(4 + 64);
  for (let i = 0; i < 100000; i++) { b[0] = i & 255; b[1] = (i >>> 8) & 255; b[2] = (i >>> 16) & 255; b[3] = i >>> 24; b.set(h, 4); h = L.sha.sha512(b); }
  console.log(`${pass}/${pass + fail} digests match WebCrypto; 100,000 SHA-512 rounds in ${Date.now() - t0} ms`);
  process.exit(fail ? 1 : 0);
})();
