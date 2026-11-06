import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { words, dword, rec, str, workbook, cfb } from './xls-fixtures.mjs';

import { B, X, cat, escher, shape, container, drawing, pictureStore, convert, png } from './xls-draw-fixtures.mjs';
if (process.env.LEGACY_BEFORE_C4) {
  for (const file of ['xls-read', 'xls-sheet', 'xls-comments', 'xls-write']) {
    vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_C4 + '/' + file + '.js', 'utf8'));
  }
}
test('Excel pictures retain original bytes, crop, sheet relationship and fractional cell anchors', async () => {
  const { output, zip, xml } = await convert(drawing(shape({ type: 75, objectType: 8,
    props: [[260, 1], [258, 16384], [896, 'Portrait']] })), [pictureStore()]);
  assert.ok(zip.has('xl/media/image1.png'));
  assert.deepEqual(Buffer.from(await zip.get('xl/media/image1.png').bytes()), png);
  assert.match(xml, /<xdr:pic>/);
  assert.match(xml, /name="Portrait"/);
  assert.match(xml, /<a:srcRect l="25000"\/>/);
  assert.match(xml, /<xdr:from><xdr:col>1<\/xdr:col><xdr:colOff>152400<\/xdr:colOff>/);
  assert.match(xml, /<xdr:row>2<\/xdr:row><xdr:rowOff>80963<\/xdr:rowOff>/);
  assert.match(await zip.get('xl/worksheets/sheet1.xml').text(), /<drawing r:id="rId1"\/>/);
  assert.match(await zip.get('[Content_Types].xml').text(), /<Default Extension="png" ContentType="image\/png"\/>/);
  assert.equal(output.losses.filter(loss => loss.notify).length, 0);
});

test('Excel TXO text, groups, child coordinates, rotation and one-cell sizing remain native', async () => {
  const group = shape({ id: 1025, flags: 0xa01, objectType: 0, coord: [0, 0, 1000, 800],
    anchor: words(2, 0, 0, 0, 0, 8, 0, 12, 0), props: [[4, 90 * 65536]] });
  const child = shape({ id: 1026, type: 202, objectType: 6, text: 'First\n第二', child: true,
    anchor: cat(dword(100), dword(150), dword(600), dword(450)) });
  const { output, xml } = await convert(drawing(container(0xf003, [...group, ...child])));
  assert.match(xml, /<xdr:twoCellAnchor editAs="oneCell">/);
  assert.match(xml, /<xdr:grpSp>/);
  assert.match(xml, /rot="5400000"/);
  assert.match(xml, /<a:chExt cx="1000" cy="800"\/>/);
  assert.match(xml, /<a:off x="100" y="150"\/><a:ext cx="500" cy="300"\/>/);
  assert.match(xml, /<a:t>First<\/a:t>/);
  assert.match(xml, /<a:t>第二<\/a:t>/);
  assert.equal(output.stats.features[0].groups, 1);
  assert.equal(output.stats.features[0].textboxes, 1);
});

test('a corrupt picture is a local loss; the following text object and cells survive', async () => {
  const records = drawing([...shape({ id: 1025, type: 75, objectType: 8, props: [[260, 99]] }),
    ...shape({ id: 1026, type: 202, objectType: 6, text: 'Keep this text' })]);
  records.push(rec(0x204, cat(words(0, 0, 0), str('Keep this cell'))));
  const { output, zip, xml } = await convert(records);
  assert.match(xml, /<a:t>Keep this text<\/a:t>/);
  assert.match(await zip.get('xl/sharedStrings.xml').text(), /Keep this cell/);
  assert.ok(output.losses.some(loss => loss.id === 'legacy-xls-picture' && loss.notify));
});

test('comment and drawing client markers share a stream without duplicating the comment', async () => {
  const records = drawing([...shape({ id: 1025, type: 202, objectType: 25, text: 'Comment' }),
    ...shape({ id: 1026, type: 202, objectType: 6, text: 'Drawing text' })]);
  records.push(rec(0x1c, cat(words(2, 1, 0, 1), str('Author'))));
  const { zip, xml, output } = await convert(records);
  assert.match(await zip.get('xl/comments1.xml').text(), />Comment<\/t>/);
  assert.match(xml, />Drawing text<\/a:t>/);
  assert.doesNotMatch(xml, />Comment</);
  assert.equal(output.losses.filter(loss => loss.notify).length, 0);
});

test('unused damaged BLIPs do not cause a document loss or a refusal', async () => {
  const { output, xml } = await convert(drawing(shape({ text: 'Readable' })), [rec(0xeb, Buffer.from([1]))]);
  assert.match(xml, />Readable<\/a:t>/);
  assert.equal(output.losses.filter(loss => loss.notify).length, 0);
});

test('known OfficeArt containers retain pictures when their version nibble is zero', async () => {
  const records = container(0xf002, container(0xf003,
    shape({ type: 75, objectType: 8, props: [[260, 1]] })));
  records[0].writeUInt16LE(0, 0);
  const { output, zip, xml } = await convert(drawing(records), [pictureStore()]);
  assert.match(xml, /<xdr:pic>/);
  assert.deepEqual(Buffer.from(await zip.get('xl/media/image1.png').bytes()), png);
  assert.equal(output.losses.filter(loss => loss.notify).length, 0);
});

test('embedded object previews keep their picture and report the lost working object', async () => {
  const link = cat(words(12, 5), dword(0), [2], dword(1), [0]);
  const { output, xml } = await convert(drawing(shape({ type: 75, objectType: 8, props: [[260, 1]],
    objectParts: [rec(9, link)] })), [pictureStore()]);
  assert.match(xml, /<xdr:pic>/);
  assert.ok(output.losses.some(loss => loss.id === 'legacy-xls-object' && loss.notify));
});

test('AutoFilter dropdown objects do not become extra control shapes or false loss notices', async () => {
  const list = rec(19, words(0, 0, 0, 0, 0, 2, 8, 0));
  const { output, xml } = await convert(drawing([...shape({ objectType: 20, objectParts: [list] }),
    ...shape({ id: 1026, text: 'Keep me' })]));
  assert.match(xml, />Keep me<\/a:t>/);
  assert.doesNotMatch(xml, /id="1025"/);
  assert.equal(output.losses.filter(loss => loss.notify).length, 0);
});

test('connectors retain preset adjustments and attachment sites in the same drawing', async () => {
  const solver = escher(0xf005, escher(0xf012,
    cat(dword(0), dword(1025), dword(1026), dword(1027), dword(1), dword(3))), 15);
  const { output, xml } = await convert(drawing([...shape(), ...shape({ id: 1026 }),
    ...shape({ id: 1027, type: 34, objectType: 1, props: [[327, 5400]] }), solver]));
  assert.match(xml, /<xdr:cxnSp>/);
  assert.match(xml, /<a:stCxn id="1025" idx="3"\/>/);
  assert.match(xml, /<a:endCxn id="1026" idx="1"\/>/);
  assert.match(xml, /<a:gd name="adj1" fmla="val 25000"\/>/);
  assert.equal(output.stats.features[0].connectors, 1);
  assert.equal(output.losses.filter(loss => loss.notify).length, 0);
});

test('monochrome pattern masks use the shape foreground and background colors', async () => {
  const { xml } = await convert(drawing(shape({ props: [[384, 1], [385, 0x0000ff],
    [387, 0xffffff], [390, 1]] })), [pictureStore()]);
  assert.match(xml, /<a:duotone><a:srgbClr val="FFFFFF"\/><a:srgbClr val="FF0000"\/><\/a:duotone>/);
});

