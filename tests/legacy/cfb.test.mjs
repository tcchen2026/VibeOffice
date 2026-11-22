import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
vm.runInThisContext(readFileSync(process.env.CRYPTO_SOURCE || 'public/common/crypto.js', 'utf8'));
const cfb = globalThis.L.officeCrypto.cfb;
const END = 0xfffffffe, FREE = 0xffffffff;
// Small independent fixture writer: root mini-stream + two storages with the same leaf name.
function fixture() {
  const b = new Uint8Array(512 * 14), d = new DataView(b.buffer);
  const u16 = (p, n) => d.setUint16(p, n, true), u32 = (p, n) => d.setUint32(p, n, true);
  u32(0, 0xe011cfd0); u32(4, 0xe11ab1a1); u16(24, 0x3e); u16(26, 3); u16(28, 0xfffe); u16(30, 9); u16(32, 6);
  u32(44, 1); u32(48, 1); u32(56, 4096); u32(60, 3); u32(64, 1); u32(68, END);
  for (let i = 0; i < 109; i++) u32(76 + i * 4, i ? FREE : 0);
  for (let i = 0; i < 128; i++) u32(512 + i * 4, FREE);
  [0xfffffffd, 2, END, END, END, 6, 7, 8, 9, 10, 11, 12, END].forEach((n, i) => u32(512 + i * 4, n));
  const entry = (id, name, type, child, right, start = END, size = 0) => {
    const p = 1024 + id * 128;
    for (let i = 0; i < name.length; i++) u16(p + i * 2, name.charCodeAt(i));
    u16(p + 64, (name.length + 1) * 2); b[p + 66] = type; b[p + 67] = 1;
    u32(p + 68, FREE); u32(p + 72, right); u32(p + 76, child); u32(p + 116, start); u32(p + 120, size);
  };
  entry(0, 'Root Entry', 5, 1, FREE, 4, 128);
  entry(1, 'A', 1, 2, 3); entry(2, 'data', 2, FREE, FREE, 0, 3);
  entry(3, 'B', 1, 4, 5); entry(4, 'data', 2, FREE, FREE, 1, 3);
  entry(5, 'large', 2, FREE, FREE, 5, 4096);
  for (let i = 0; i < 128; i++) u32(2048 + i * 4, i < 2 ? END : FREE);
  b.set([11, 12, 13], 2560); b.set([21, 22, 23], 2624); b.fill(42, 3072);
  return { b, d, u16, u32 };
}
const refuses = (fn, code) => assert.throws(fn, e => e.name === 'CFBError' && (!code || e.code === code));
test('shared CFB API resolves storage-qualified streams without ambiguous leaves', () => {
  const { b } = fixture(), f = cfb(b);
  assert.deepEqual([...f.stream('a/DATA')], [11, 12, 13]);
  assert.deepEqual([...f.stream('B/data')], [21, 22, 23]);
  assert.equal(f.stream('data'), null); assert.equal(f.stream('absent'), null);
  assert.equal(f.stream('large').length, 4096); assert.equal(f.stream('large')[4095], 42);
  assert.equal(f.entries.length, 5);
});
test('v3 ignores the historical uninitialized high size DWORD', () => {
  const { b, u32 } = fixture(); u32(1024 + 2 * 128 + 124, 0xffff1234);
  assert.equal(cfb(b).stream('A/data').length, 3);
});
test('FAT and mini-FAT cycles stop only the affected stream, without repeating bytes', () => {
  let f = fixture(); f.u32(512 + 6 * 4, 5);
  assert.deepEqual([...cfb(f.b).stream('A/data')], [11, 12, 13]);
  { const c = cfb(f.b); assert.equal(c.stream('large').length, 1024); assert.ok(c.issues.some(x => x.code === 'CFB_CYCLE')); }
  f = fixture(); f.u32(2048, 0); f.u32(1024 + 2 * 128 + 120, 70);
  assert.equal(cfb(f.b).stream('large').length, 4096);
  { const c = cfb(f.b); assert.equal(c.stream('A/data').length, 64); assert.ok(c.issues.some(x => x.code === 'CFB_CYCLE')); }
});
test('short chains and stale sizes return their available bytes; invalid shifts still refuse', () => {
  let f = fixture(); f.u32(512 + 5 * 4, END); assert.equal(cfb(f.b).stream('large').length, 512);
  assert.equal(cfb(f.b).stream('B/data')[0], 21);
  f = fixture(); f.u32(1024 + 5 * 128 + 120, 0xffffffff);
  assert.equal(cfb(f.b).stream('A/data')[0], 11); assert.equal(cfb(f.b).stream('large').length, 4096);
  f = fixture(); f.u16(30, 31); refuses(() => cfb(f.b), 'CFB_HEADER');
});
test('a short final sector is zero-padded, and advisory header fields do not reject the document', () => {
  const f = fixture(); f.u16(26, 99); f.u16(28, 0); f.u16(32, 3); f.u32(56, 1);
  const data = cfb(f.b.subarray(0, f.b.length - 8)).stream('large');
  assert.equal(data.length, 4096); assert.equal(data[4087], 42); assert.deepEqual([...data.subarray(4088)], Array(8).fill(0));
});
test('a bad entry, duplicate path, bad pointer or cyclic branch keeps the sound siblings', () => {
  let f = fixture(); f.u16(1024 + 2 * 128 + 64, 1000);
  assert.equal(cfb(f.b).stream('A/data'), null); assert.equal(cfb(f.b).stream('B/data')[0], 21);
  f = fixture(); f.u16(1024 + 3 * 128, 65); // B renamed to duplicate A
  assert.equal(cfb(f.b).stream('large').length, 4096); assert.ok(cfb(f.b).issues.some(x => x.code === 'CFB_DUPLICATE'));
  f = fixture(); f.u32(1024 + 2 * 128 + 68, 10000); assert.equal(cfb(f.b).stream('large').length, 4096);
  f = fixture(); f.u32(1024 + 128 + 68, 1); assert.equal(cfb(f.b).stream('large').length, 4096);
});
test('trailing sectors beyond FAT coverage do not hide existing streams', () => {
  const f = fixture(), b = new Uint8Array(512 * 150); b.set(f.b);
  assert.equal(cfb(b).stream('large')[0], 42); assert.equal(cfb(b).stream('A/data')[0], 11);
});
test('all truncated lengths terminate without native exceptions', () => {
  const { b } = fixture();
  for (let n = 0; n < b.length; n++) {
    try { cfb(b.subarray(0, n)).stream('large'); }
    catch (e) { assert.equal(e.name, 'CFBError', 'length ' + n + ': ' + e.message); }
  }
});
test('an over-depth directory branch is skipped without following it', () => {
  const f = fixture(); f.u32(1024 + 2 * 128 + 76, 4); f.b[1024 + 2 * 128 + 66] = 1;
  const read = cfb(f.b, { maxDepth: 1 }); assert.equal(read.stream('large').length, 4096); assert.ok(read.issues.length);
});

test('stale FAT counts and a broken directory tail retain the usable directory prefix', () => {
  let f = fixture(); f.u32(44, 2105376); assert.equal(cfb(f.b).stream('A/data')[0], 11);
  f = fixture(); f.u32(512 + 2 * 4, 1000); assert.equal(cfb(f.b).stream('A/data')[0], 11);
});
test('a DIFAT terminator may be FREE, and unused FAT entries need not cover trailing data', () => {
  const f = fixture(), b = new Uint8Array(512 * (110 * 128 + 1)); b.set(f.b);
  const d = new DataView(b.buffer), put = (p, n) => d.setUint32(p, n, true);
  put(44, 110); put(68, 121); put(72, 1);
  for (let i = 1; i < 109; i++) put(76 + i * 4, i + 12);
  for (let i = 0; i < 128; i++) put(122 * 512 + i * 4, i ? FREE : 122);
  assert.equal(cfb(b).stream('large')[0], 42);
});
