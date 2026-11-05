import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { load } from '../../tools/legacy/load.mjs';
const B = load();
if (process.env.BASE_SOURCE) vm.runInThisContext(fs.readFileSync(process.env.BASE_SOURCE, 'utf8'));
if (process.env.PROPERTIES_SOURCE) vm.runInThisContext(fs.readFileSync(process.env.PROPERTIES_SOURCE, 'utf8'));
if (process.env.PROPERTY_PARTS_SOURCE) vm.runInThisContext(fs.readFileSync(process.env.PROPERTY_PARTS_SOURCE, 'utf8'));
const u32 = n => { const b = Buffer.alloc(4); b.writeUInt32LE(n); return b; };
const typed = (t, ...data) => Buffer.concat([u32(t), ...data]);
const str = s => { const b = Buffer.from(s + '\0'); return Buffer.concat([u32(b.length), b]); };
function properties(entries) {
  const parts = entries.map(([id, data]) => [id, data]), header = Buffer.alloc(48), section = Buffer.alloc(8 + parts.length * 8);
  header.writeUInt16LE(0xfffe); header.writeUInt32LE(1, 24); header.writeUInt32LE(48, 44);
  let offset = section.length; section.writeUInt32LE(parts.length, 4);
  parts.forEach(([id, data], i) => { section.writeUInt32LE(id, 8 + i * 8); section.writeUInt32LE(offset, 12 + i * 8); offset += data.length; });
  section.writeUInt32LE(offset); return Buffer.concat([header, section, ...parts.map(p => p[1])]);
}
test('Office unaligned document-part strings and heading variants retain every element', () => {
  const b = properties([[12, typed(0x100c, u32(4), typed(30, str('Title')), typed(3, u32(1)), typed(30, str('Headings')), typed(3, u32(2)))], [13, typed(0x101e, u32(3), str('One'), str('AB'), str('Last'))]]);
  const p = B.properties(b, { kind: 'documentSummary' })[0];
  assert.deepEqual(p.values[12].value, ['Title', 1, 'Headings', 2]); assert.deepEqual(p.values[13].value, ['One', 'AB', 'Last']);
});
test('packed scalar vectors and UTF16 codepage byte lengths are distinct from LPWSTR characters', () => {
  const b = properties([[1, typed(2, Buffer.from([0xb0, 4, 0, 0]))], [2, typed(30, u32(4), Buffer.from('中\0', 'utf16le'))], [3, typed(0x1002, u32(3), Buffer.from([1, 0, 2, 0, 3, 0]))]]);
  const p = B.properties(b)[0]; assert.equal(p.values[2].value, '中'); assert.deepEqual(p.values[3].value, [1, 2, 3]);
});
test('one corrupt optional property cannot consume or discard its siblings', () => {
  const b = properties([[2, typed(30, u32(99999))], [14, typed(3, u32(5))]]), p = B.properties(b)[0];
  assert.equal(p.values[14].value, 5); assert.equal(p.values[2].unsupported, true); assert.equal(p.issues[0].id, 2);
});
test('property parts preserve editing duration, metadata escapes and custom values', () => {
  const ticks = Buffer.alloc(8); ticks.writeBigUInt64LE(1800000000n);
  const summary = properties([[2, typed(30, str('<&title'))], [10, typed(64, ticks)]]);
  const result = B.propertyParts({ stream: name => name === '\x05SummaryInformation' ? summary : null });
  assert.match(result.parts[0].data, /&lt;&amp;title/); assert.match(result.parts[1].data, /<TotalTime>3<\/TotalTime>/);
  assert.equal(result.losses.length, 0);
});
test('loss notification requires explicit content-loss opt-in', () => {
  assert.equal(B.loss('metadata', 'Document properties', 'docProps', 'drop').notify, false);
  assert.equal(B.loss('picture', 'A picture was dropped', 'slide1', 'drop', 'Slide 1', true).notify, true);
  assert.equal(B.loss('metadata', 'Document properties', 'docProps', 'drop', undefined, 'true').notify, false);
});
test('editing durations outside the OOXML integer range lose only that metadata', () => {
  const ticks=Buffer.alloc(8);ticks.writeBigUInt64LE(2147483648n*600000000n);
  const summary=properties([[10,typed(64,ticks)],[14,typed(3,u32(7))]]),result=B.propertyParts({stream:name=>name==='\x05SummaryInformation'?summary:null});
  const app=result.parts.find(p=>p.name==='docProps/app.xml').data;
  assert.doesNotMatch(app,/<TotalTime>/);assert.match(app,/<Pages>7<\/Pages>/);assert.ok(result.losses.some(l=>l.id==='legacy-property-value'&&!l.notify));
});
test('OEM codepages decode legacy characters and non-ASCII-compatible bytes', () => {
  assert.equal(B.decode(Uint8Array.of(0x80, 0xdb), 437), 'Ç█');
  assert.equal(B.decode(Uint8Array.of(0x9b, 0xd5), 850), 'øı');
  assert.equal(B.decode(Uint8Array.of(0xa5, 0xa9), 852), 'ąę');
  assert.equal(B.decode(Uint8Array.of(0xd5), 858), '€');
  assert.equal(B.decode(Uint8Array.of(0x25), 864), '٪');
  const summary = properties([[1, typed(2, Buffer.from([0xb5, 1, 0, 0]))], [2, typed(30, u32(3), Buffer.from([0x80,0xdb,0]))]]);
  assert.equal(B.properties(summary)[0].values[2].value, 'Ç█');
});
test('an unknown property codepage loses only that string and reports silently', () => {
  const cp = Buffer.alloc(4); cp.writeUInt16LE(49999);
  const summary = properties([[1, typed(2, cp)], [2, typed(30, str('Title'))], [14, typed(3, u32(7))]]);
  const parsed = B.properties(summary)[0];
  assert.equal(parsed.values[14].value, 7); assert.equal(parsed.values[2].error, 'CODEPAGE_UNSUPPORTED');
  const result = B.propertyParts({ stream: name => name === '\x05SummaryInformation' ? summary : null });
  assert.ok(result.losses.length); assert.ok(result.losses.every(l => l.notify === false));
  assert.match(result.parts.find(p => p.name === 'docProps/app.xml').data, /<Pages>7<\/Pages>/);
});
test('package writer reuses ZIP, escapes attributes and rejects duplicate or unsafe names', async () => {
  const p = new B.Package(); p.add('word/document.xml', '<test/>', 'application/xml'); p.rel('', 'officeDocument', 'word/document.xml');
  assert.throws(() => p.add('../outside', 'bad'), { code: 'PART_NAME' });
  assert.throws(() => p.add('word/document.xml', 'bad'), { code: 'PART_DUPLICATE' });
  const bytes = await p.finish(), zip = await globalThis.L.zip.read(bytes.buffer);
  assert.ok(zip.has('word/document.xml')); assert.ok(zip.has('[Content_Types].xml'));
  const x = new B.XML(); x.start('a', { title: 'a\nb"' }).text('<&').end();
  assert.match(x.finish(), /title="a&#10;b&quot;">&lt;&amp;<\/a>/);
  assert.throws(() => new B.XML().start('a').finish(), { code: 'XML_STRUCTURE' });
});
