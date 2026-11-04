// A copy of the suite served over plain http (not https, localhost or 127.0.0.1) is not a secure context: the browser
// gives its pages no crypto.randomUUID and no crypto.subtle. The apps take both from shared helpers that fall back
// (K.uuid and K.guid in common/opc.js, L.sha.digest in common/sha.js); the Start Center says what stays off
// (docs/suite.md, Plain http):  node --test tests/insecure.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
for (const f of ['xml', 'opc', 'sha', 'crypto']) require(`../public/common/${f}.js`);
const { opc: K, officeCrypto: OC } = globalThis.L;
globalThis.L.xml = globalThis.L.xmlTree;   // crypto.js's XML parser in Node, as on Ledger's page (ledger/js/xml.js)
const pub = path.join(path.dirname(new URL(import.meta.url).pathname), '../public');

/** fn run with crypto as an http page has it: getRandomValues only */
async function insecure(fn) {
  const real = Object.getOwnPropertyDescriptor(globalThis, 'crypto'), web = globalThis.crypto;
  Object.defineProperty(globalThis, 'crypto', { configurable: true, value: { getRandomValues: (a) => web.getRandomValues(a) } });
  try { return await fn(); } finally { Object.defineProperty(globalThis, 'crypto', real); }
}

test('only the shared helpers call randomUUID or crypto.subtle', () => {
  /* saving with a password needs WebCrypto's AES, and says so on an http page */
  const ALLOWED = { randomUUID: ['common/opc.js'], 'crypto.subtle': ['common/sha.js', 'common/crypto.js'] };
  const scripts = (dir) => fs.readdirSync(dir, { withFileTypes: true })
    .flatMap((e) => (e.isDirectory() ? scripts(path.join(dir, e.name)) : e.name.endsWith('.js') ? [path.join(dir, e.name)] : []));
  const found = [];
  for (const file of scripts(pub)) {
    const rel = path.relative(pub, file), text = fs.readFileSync(file, 'utf8');
    if (rel === 'sw.js') continue;   // a service worker runs only on secure pages
    for (const [name, allowed] of Object.entries(ALLOWED)) if (text.includes(name) && !allowed.includes(rel)) found.push(`${rel}: ${name}`);
  }
  assert.deepEqual(found, []);
});

test('K.uuid and K.guid give random version 4 UUIDs without randomUUID', async () => {
  await insecure(() => {
    const seen = new Set();
    for (let i = 0; i < 500; i++) seen.add(K.uuid());
    assert.equal(seen.size, 500);
    for (const u of seen) assert.match(u, /^[\da-f]{8}-[\da-f]{4}-4[\da-f]{3}-[89ab][\da-f]{3}-[\da-f]{12}$/);
    assert.match(K.guid(), /^\{[\dA-F]{8}-[\dA-F]{4}-4[\dA-F]{3}-[89AB][\dA-F]{3}-[\dA-F]{12}\}$/);
  });
});

test('a password-protected file opens without crypto.subtle; a wrong password is refused, saving with one says why', async () => {
  const pkg = new TextEncoder().encode('PK\u0003\u0004 the package bytes '.repeat(400));   // several 4096-byte segments
  const file = await OC.encrypt(pkg, 'Secret 1');   // as an https page saves it
  await insecure(async () => {
    assert.equal(Buffer.compare(Buffer.from(await OC.decrypt(file, 'Secret 1')), Buffer.from(pkg)), 0);
    await assert.rejects(OC.decrypt(file, 'Secret 2'), (e) => e.code === 'badpass');
    await assert.rejects(OC.encrypt(pkg, 'Secret 1'), /secure \(https\) page/);
  });
});
