import { test } from 'node:test';
import assert from 'node:assert/strict';
import { load } from '../../tools/legacy/load.mjs';
const B = load();
const record = (id, data) => Buffer.concat([Buffer.from([id & 255, id >> 8, data.length & 255, data.length >> 8]), Buffer.from(data)]);
const u32 = n => { const b = Buffer.alloc(4); b.writeUInt32LE(n); return b; };

test('byte views check their own bounds and preserve subarray offsets', () => {
  const r = new B.Reader(new Uint8Array([99, 1, 2, 3, 88]).subarray(1, 4));
  assert.equal(r.u16(0), 513); assert.equal(r.u8(2), 3);
  for (const p of [-1, 1.5, 3, 2 ** 40]) assert.throws(() => r.u16(p), { name: 'LegacyError', code: 'TRUNCATED' });
});
test('BIFF compressed Unicode is not CP1252; XML cannot turn text into markup', () => {
  assert.equal(B.latin16(new Uint8Array([0x80, 65])), '\x80A');
  assert.equal(B.decode(new Uint8Array([0x80]), 1252), '€');
  assert.equal(B.xml('<&"\x00'), '&lt;&amp;&quot;�');
  assert.equal(B.xstring('_x000A_\x07\r'), '_x005F_x000A__x0007__x000D_');
});
test('SST continuation changes character width but not following rich-text bytes', () => {
  const a = Buffer.concat([u32(2), u32(2), Buffer.from([4, 0, 0, 65, 66])]);
  const b = Buffer.from([1, 0x2d, 0x4e, 0x87, 0x65, 1, 0, 0, 90]);
  const records = B.xls.records(Buffer.concat([record(0xfc, a), record(0x3c, b)]));
  assert.deepEqual(B.xls.sst(records, 0), ['AB中文', 'Z']);
  assert.throws(() => B.xls.sst(B.xls.records(record(0xfc, a)), 0), { name: 'LegacyError', code: 'SST_TRUNCATED' });
});
test('RK integers, fractions and IEEE values decode without changing bits', () => {
  assert.equal(B.xls.rk((-42 << 2) | 2), -42); assert.equal(B.xls.rk((1234 << 2) | 3), 12.34);
  const d = new DataView(new ArrayBuffer(8)); d.setFloat64(0, 1.5, true); assert.equal(B.xls.rk(d.getUint32(4, true)), 1.5);
});
test('truncated BIFF and nested presentation records fail with typed errors', () => {
  assert.throws(() => B.xls.records(new Uint8Array([9, 8, 8, 0, 0])), { name: 'LegacyError', code: 'TRUNCATED' });
  const b = Buffer.alloc(8); b.writeUInt16LE(15, 0); b.writeUInt32LE(0xffffffff, 4);
  assert.throws(() => B.ppt.record(new B.Reader(b), 0), { name: 'LegacyError', code: 'PPT_RECORD' });
  let rec = { version: 15, type: 1000, r: new B.Reader(new Uint8Array(0)) };
  for (let i = 0; i < 140; i++) { const header = Buffer.alloc(8); header.writeUInt16LE(15); header.writeUInt16LE(1000, 2); header.writeUInt32LE(rec.r.length, 4); rec = { version: 15, type: 1000, r: new B.Reader(Buffer.concat([header, rec.r.bytes])) }; }
  assert.throws(() => B.ppt.walk(rec, () => {}), { name: 'LegacyError', code: 'RECORD_LIMIT' });
});
test('Word compressed and Unicode pieces retain story boundaries and reject out-of-file FCs', () => {
  const doc = Buffer.alloc(2048), table = Buffer.alloc(33);
  doc.writeUInt16LE(0xa5ec); doc.writeUInt16LE(0xc1, 2); doc.writeUInt16LE(14, 32); doc.writeUInt16LE(22, 62);
  doc.writeUInt32LE(4, 76); doc.writeUInt16LE(93, 152); doc.writeUInt32LE(33, 154 + 33 * 8 + 4);
  table[0] = 2; table.writeUInt32LE(28, 1); table.writeUInt32LE(2, 9); table.writeUInt32LE(4, 13);
  table.writeUInt32LE(0x40000800, 19); table.writeUInt32LE(1030, 27);
  doc.write('AB', 1024); doc.write('中\r', 1030, 'utf16le');
  const cfb = { stream: name => name === 'WordDocument' ? doc : table };
  assert.equal(B.doc.read(cfb).storyText.main, 'AB中\r');
  table.writeUInt32LE(0x3fffffff, 27);
  assert.throws(() => B.doc.read(cfb), { name: 'LegacyError', code: 'TRUNCATED' });
});
test('OLE property page/word/character counts are read by ID', () => {
  const b = Buffer.alloc(104); b.writeUInt16LE(0xfffe); b.writeUInt32LE(1, 24); b.writeUInt32LE(48, 44);
  b.writeUInt32LE(56, 48); b.writeUInt32LE(3, 52);
  for (let i = 0; i < 3; i++) { b.writeUInt32LE(14 + i, 56 + i * 8); b.writeUInt32LE(32 + i * 8, 60 + i * 8); b.writeUInt32LE(3, 80 + i * 8); b.writeInt32LE([3, 120, 650][i], 84 + i * 8); }
  const p = B.properties(b)[0]; assert.equal(p.values[14].value, 3); assert.equal(p.values[15].value, 120); assert.equal(p.values[16].value, 650);
});

test('PowerPoint follows the latest persist map and rejects an edit-history cycle', () => {
  const rec = (type, bytes, instance = 0, container = false) => {
    const header = Buffer.alloc(8); header.writeUInt16LE(instance * 16 + (container ? 15 : 0)); header.writeUInt16LE(type, 2); header.writeUInt32LE(bytes.length, 4); return Buffer.concat([header, bytes]);
  };
  const doc = Buffer.alloc(1024), user = Buffer.alloc(28);
  const persist = Buffer.alloc(20); persist.writeUInt32LE(2); persist.writeUInt32LE(256, 12);
  const latest = rec(1000, rec(4080, Buffer.concat([rec(1011, persist), rec(4008, Buffer.from('current'))]), 0, true), 0, true);
  doc.set(latest, 40); doc.set(rec(1006, rec(4008, Buffer.from('shape')), 0, true), 160);
  const ptr = Buffer.concat([u32(0x00200001), u32(40), u32(160)]); doc.set(rec(6002, ptr), 240);
  const edit = Buffer.alloc(28); edit.writeUInt32LE(400, 8); edit.writeUInt32LE(240, 12); edit.writeUInt32LE(1, 16); doc.set(rec(4085, edit), 300);
  const previous = Buffer.alloc(28); previous.writeUInt32LE(460, 12); previous.writeUInt32LE(1, 16); doc.set(rec(4085, previous), 400);
  doc.set(rec(6002, Buffer.concat([u32(0x00100001), u32(600)])), 460);
  doc.set(rec(1000, rec(4008, Buffer.from('obsolete')), 0, true), 600);
  user.writeUInt32LE(0xe391c05f, 12); user.writeUInt32LE(300, 16);
  const cfb = { stream: name => name === 'Current User' ? user : doc }, parsed = B.ppt.read(cfb);
  assert.equal(parsed.slides.length, 1); assert.deepEqual(parsed.slides[0].text, ['current', 'shape']); assert.equal(parsed.edits.length, 2);
  doc.writeUInt32LE(300, 416); assert.throws(() => B.ppt.read(cfb), { name: 'LegacyError', code: 'PPT_CYCLE' });
});
test('Word 6/95 and Word 2 signatures remain distinct', () => {
  const b = Buffer.alloc(32); b.writeUInt16LE(0xa5dc); b.writeUInt16LE(0x65, 2);
  assert.equal(B.doc.fib(b).version, 'Word6/95');
  b.writeUInt16LE(0xa5db, 2); assert.equal(B.doc.fib(b).version, 'Word2');
});
test('Word revision inventory reads bounded character-property runs', () => {
  const bytes = Buffer.alloc(1024), table = Buffer.alloc(12); table.writeUInt32LE(1, 8);
  bytes[1023] = 1; bytes[520] = 80; bytes[672] = 3; bytes.writeUInt16LE(0x0801, 673); bytes[675] = 1;
  const fib = { r: new B.Reader(bytes), pair: () => ({ offset: 0, length: 12 }) }, t = new B.Reader(table);
  assert.equal(B.doc.revisionFlags(fib, t), true); bytes[675] = 0; assert.equal(B.doc.revisionFlags(fib, t), false);
  table.writeUInt32LE(99999, 8); assert.throws(() => B.doc.revisionFlags(fib, t), { name: 'LegacyError', code: 'TRUNCATED' });
});
