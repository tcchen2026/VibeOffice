import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { load } from '../../tools/legacy/load.mjs';
import { source, project, word, excel, powerpoint, commands } from './vba-fixtures.mjs';
import { deck, shape, container, rec } from './ppt-animation-fixtures.mjs';
import { encrypted } from './encrypted-fixtures.mjs';

const B = load();
const pack = B.vba.pack;
if (process.env.LEGACY_BEFORE_G) {
  for (const name of ['package', 'doc-write', 'xls-write', 'ppt-links', 'ppt-write']) {
    const file = process.env.LEGACY_BEFORE_G + '/' + name + '.js';
    vm.runInThisContext(fs.readFileSync(file, 'utf8'), { filename: file });
  }
}
if (process.env.LEGACY_BEFORE_G_HOSTS)
  vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_G_HOSTS, 'utf8'));
const snapshot = cfb => cfb.entries.filter(e => e.type === 2).map(e =>
  [e.path, Buffer.from(cfb.stream(e.path)).toString('hex')]).sort(([a], [b]) => a.localeCompare(b));
const convert = async (format, bytes, options) => {
  const result = await B[format].convert(bytes, options);
  return { ...result, zip: await L.zip.read(result.pkg) };
};
const macroLoss = r => r.losses.filter(l => l.id === 'legacy-macros' && l.notify);

test('VBA packaging preserves short, cutoff, empty and multi-FAT streams plus form storage metadata', () => {
  const streams = project({ empty: Buffer.alloc(0), short: Buffer.from([1, 2, 3]),
    cutoff: Buffer.alloc(4096, 0x71), long: Buffer.alloc(7 * 1024 * 1024, 0x53),
    'Form1/f': Buffer.from([8, 9]), 'Form1/i01/o': Buffer.from([4, 5, 6]) });
  const clsid = Buffer.from('0123456789abcdef0123456789abcdef', 'hex');
  const input = source(streams, { Form1: { clsid } });
  const bytes = pack(input);
  const output = B.cfb(bytes);
  assert.deepEqual(snapshot(output), snapshot(input));
  assert.deepEqual(Buffer.from(output.entries.find(e => e.path === 'Form1').clsid), clsid);
  assert.ok(new DataView(bytes.buffer).getUint32(72, true) > 0, 'DIFAT extension is exercised');
  assert.deepEqual(output.issues, []);
});

test('VBA storage extraction excludes neighboring host and embedded projects', () => {
  const streams = Object.fromEntries(Object.entries(project()).map(([p, b]) => ['Macros/' + p, b]));
  streams.WordDocument = Buffer.from('Document bytes');
  streams['Other/VBA/dir'] = Buffer.from('Unrelated embedded project');
  const result = B.vba.read(source(streams), 'Macros', [], 'Document');
  assert.deepEqual(snapshot(B.cfb(result.bytes)), snapshot(source(project())));
});

test('VBA directory ordering uses CFB lengths and case without losing non-ASCII names', () => {
  const input = source({ Z: Buffer.from([1]), aa: Buffer.from([2]), B: Buffer.from([3]),
    é: Buffer.from([4]), '表单/f': Buffer.from([5]) });
  const bytes = pack(input);
  assert.deepEqual(snapshot(B.cfb(bytes)), snapshot(input));
  const view = new DataView(bytes.buffer);
  const names = new Map();
  for (let id = 0; id <= input.entries.length; id++) {
    const p = 512 + id * 128;
    names.set(id, new TextDecoder('utf-16le').decode(bytes.subarray(p, p + view.getUint16(p + 64, true) - 2)));
  }
  const walk = id => id === 0xffffffff ? [] : [
    ...walk(view.getUint32(512 + id * 128 + 68, true)), names.get(id),
    ...walk(view.getUint32(512 + id * 128 + 72, true))
  ];
  assert.deepEqual(walk(view.getUint32(512 + 76, true)), ['B', 'Z', 'é', 'aa', '表单']);
});

test('VBA malformed or oversized streams cannot allocate from their declared size', () => {
  const input = source(project());
  input.entries.find(e => e.path === 'VBA/dir').size = Number.MAX_SAFE_INTEGER;
  const losses = [];
  assert.equal(B.vba.read(input, '', losses, 'Workbook'), null);
  assert.equal(losses.length, 1);
  assert.equal(losses[0].notify, true);
  assert.throws(() => pack({ ...input, entries: [...input.entries, input.entries[0]] }), /VBA project storage/);
});

test('VBA decompression handles overlapping copies and bounds chunks, offsets and output', () => {
  // Three literals followed by a six-byte, distance-three copy token: abcabcabc.
  const data = Buffer.from([1, 0x05, 0xb0, 0x08, 97, 98, 99, 0x03, 0x20]);
  assert.equal(new TextDecoder().decode(B.vba.decompress(data)), 'abcabcabc');
  assert.throws(() => B.vba.decompress(data, 8), /VBA project data/);
  assert.throws(() => B.vba.decompress(Buffer.from([1, 2, 0xb0, 1, 0, 0])), /VBA project data/);
  assert.throws(() => B.vba.decompress(data.subarray(0, 7)), /VBA project data/);
  assert.throws(() => B.vba.decompress(Buffer.from([1, 0, 0x30, 0])), /VBA project data/);
});

test('VBA undecodable metadata stays opaque without dropping the project', () => {
  const streams = project({ 'VBA/dir': Buffer.from([0]) });
  const losses = [];
  const result = B.vba.read(source(streams), '', losses, 'Workbook');
  assert.deepEqual(snapshot(B.cfb(result.bytes)), snapshot(source(streams)));
  assert.equal(result.metadata, null);
  assert.equal(losses.length, 1);
  assert.equal(losses[0].notify, false);
});

test('VBA PROJECTwm supplies Unicode document and designer names in their source encoding', () => {
  const name = 'Résumé';
  const streams = project({ PROJECT: Buffer.from('Document=Resume/&H00000000\r\nBaseClass=Form1\r\n'),
    PROJECTwm: Buffer.concat([Buffer.from('Resume\0'), Buffer.from(name + '\0', 'utf16le'), Buffer.alloc(2)]) });
  const metadata = B.vba.read(source(streams), '', [], 'Workbook').metadata;
  assert.deepEqual(metadata.documents, [name]);
  assert.deepEqual(metadata.forms, ['Form1']);
  assert.equal(metadata.codepage, 1252);
});

test('Word keeps the project, command names and DOP event mask in a DOCM package', async () => {
  const dop = Buffer.alloc(616);
  dop.writeUInt32LE(0x4017, 0x1b2);
  const result = await convert('doc', word(B, pack, { dop, tables: [[24, commands()]] }));
  assert.equal(result.type, 'docm');
  const data = await result.zip.get('word/vbaData.xml').text();
  assert.match(data, /<wne:eventDocNew\/><wne:eventDocOpen\/><wne:eventDocClose\/>/);
  assert.match(data, /<wne:eventDocXmlAfterInsert\/>/);
  assert.match(data, /<wne:eventDocBuildingBlockAfterInsert\/>/);
  assert.match(data, /wne:macroName="PROJECT.MODULE1.RUN" wne:name="Project.Module1.Run"/);
  assert.match(await result.zip.get('word/_rels/vbaProject.bin.rels').text(), /wordVbaData/);
  assert.deepEqual(snapshot(B.cfb(await result.zip.get('word/vbaProject.bin').bytes())), snapshot(source(project())));
  assert.deepEqual(macroLoss(result), []);
});

test('Word macro templates become DOTM without changing the template intent', async () => {
  const result = await convert('doc', word(B, pack, { template: true }));
  assert.equal(result.type, 'dotm');
  assert.match(await result.zip.get('[Content_Types].xml').text(), /ms-word.template.macroEnabled.main\+xml/);
});

test('Word damaged supplemental commands lose only that binding and keep VBA and document text', async () => {
  const result = await convert('doc', word(B, pack, { text: 'Keep this text\r', tables: [[24, Buffer.from([255, 1])]] }));
  assert.equal(result.type, 'docm');
  assert.match(await result.zip.get('word/document.xml').text(), /Keep this text/);
  assert.ok(result.zip.get('word/vbaProject.bin'));
  assert.ok(macroLoss(result).some(l => l.what.includes('command names')));
});

test('Excel preserves workbook and worksheet code names independently of visible sheet titles', async () => {
  const result = await convert('xls', excel(pack, { workbookName: 'DieseArbeitsmappe', sheetName: 'Tabelle7' }));
  assert.equal(result.type, 'xlsm');
  assert.match(await result.zip.get('xl/workbook.xml').text(), /codeName="DieseArbeitsmappe"/);
  assert.match(await result.zip.get('xl/workbook.xml').text(), /name="Visible title"/);
  assert.match(await result.zip.get('xl/worksheets/sheet1.xml').text(), /codeName="Tabelle7"/);
  assert.deepEqual(snapshot(B.cfb(await result.zip.get('xl/vbaProject.bin').bytes())), snapshot(source(project())));
  assert.deepEqual(macroLoss(result), []);
});

test('Excel macro templates use XLTM and the Office VBA content-type Default and relationship', async () => {
  const result = await convert('xls', excel(pack, { template: true }));
  assert.equal(result.type, 'xltm');
  const types = await result.zip.get('[Content_Types].xml').text();
  assert.match(types, /<Default Extension="bin" ContentType="application\/vnd.ms-office.vbaProject"\/>/);
  assert.match(types, /ms-excel.template.macroEnabled.main\+xml/);
  assert.match(await result.zip.get('xl/_rels/workbook.xml.rels').text(),
    /Type="http:\/\/schemas.microsoft.com\/office\/2006\/relationships\/vbaProject" Target="vbaProject.bin"/);
});

test('Excel an incomplete optional project leaves a readable workbook with a precise loss', async () => {
  const streams = project();
  delete streams['VBA/dir'];
  const result = await convert('xls', excel(pack, {}, streams));
  assert.equal(result.type, 'xlsx');
  assert.equal(result.zip.get('xl/vbaProject.bin'), undefined);
  assert.ok(macroLoss(result).some(l => l.what === 'Damaged macros were not kept'));
});

test('PowerPoint follows nested active VBA persist references for compressed and plain storage', async () => {
  for (const uncompressed of [false, true]) {
    const bytes = powerpoint(pack, { uncompressed });
    const result = await convert('ppt', bytes);
    assert.equal(result.type, 'pptm');
    const projectBytes = await result.zip.get('ppt/vbaProject.bin').bytes();
    assert.deepEqual(Buffer.from(projectBytes), Buffer.from(pack(source(project()))));
    assert.deepEqual(macroLoss(result), []);
  }
});

test('PowerPoint macro shows and templates retain PPSM and POTM intent', async () => {
  for (const [extension, type] of [['pps', 'ppsm'], ['pot', 'potm']]) {
    const result = await convert('ppt', powerpoint(pack), { extension });
    assert.equal(result.type, type);
    assert.ok(result.zip.get('ppt/vbaProject.bin'));
  }
});

test('PowerPoint keeps the macro click action alongside the project without an external relationship', async () => {
  const atom = Buffer.alloc(16);
  atom[8] = 1;
  const action = container(4082, [rec(4083, atom), rec(4026, Buffer.from('Module1.Run', 'utf16le'), 2)]);
  const pages = deck([shape(2050, { client: [action] })]).pages;
  const result = await convert('ppt', powerpoint(pack, { pages }));
  assert.equal(result.type, 'pptm');
  assert.match(await result.zip.get('ppt/slides/slide1.xml').text(), /action="ppaction:\/\/macro\?name=Module1.Run"/);
  assert.doesNotMatch(await result.zip.get('ppt/slides/_rels/slide1.xml.rels').text(), /TargetMode="External"/);
});

test('PowerPoint ignores an explicitly empty VBA storage and isolates a missing active persist object', async () => {
  const empty = await convert('ppt', powerpoint(pack, { empty: true }));
  assert.equal(empty.type, 'pptx');
  assert.deepEqual(macroLoss(empty), []);
  const broken = await convert('ppt', powerpoint(pack, { reference: 999 }));
  assert.equal(broken.type, 'pptx');
  assert.equal(broken.stats.slides, 2);
  assert.ok(macroLoss(broken).some(l => l.what === 'Damaged macros were not kept'));
});

test('PowerPoint unrelated damaged document information does not invent a macro loss', async () => {
  const book = { documentRecords: [{ type: 2000, r: new B.Reader(new Uint8Array([1])) }], losses: [] };
  assert.equal(await B.vba.powerpoint(book), null);
  assert.deepEqual(book.losses, []);
});

test('VBA signatures report their loss without treating optional property damage as lost macros', () => {
  const bytes = Buffer.alloc(76);
  bytes.writeUInt16LE(0xfffe, 0);
  bytes.writeUInt32LE(1, 24);
  Buffer.from('02d5cdd59c2e1b10939708002b2cf9ae', 'hex').copy(bytes, 28);
  bytes.writeUInt32LE(48, 44);
  bytes.writeUInt32LE(28, 48);
  bytes.writeUInt32LE(1, 52);
  bytes.writeUInt32LE(24, 56);
  bytes.writeUInt32LE(16, 60);
  bytes.writeUInt32LE(65, 64);
  bytes.writeUInt32LE(4, 68);
  bytes.writeUInt32LE(1234, 72);
  const losses = [];
  B.vba.signatureNotice({ stream: () => bytes }, losses, 'Workbook');
  assert.equal(losses.length, 1);
  assert.equal(losses[0].what, 'The macro digital signature was not kept');
  assert.equal(losses[0].notify, true);
  const damaged = [];
  B.vba.signatureNotice({ stream: () => bytes.subarray(0, 10) }, damaged, 'Workbook');
  assert.deepEqual(damaged, []);
});

test('Word empty signature variables do not claim a lost digital signature', async () => {
  const names = Buffer.from([255, 255, 1, 0, 4, 0, 4, 0, 83, 0, 105, 0, 103, 0, 110, 0, 0, 0, 0, 0]);
  for (const signed of [false, true]) {
    const value = signed ? Buffer.from([1, 0, 42, 0]) : Buffer.from([0, 0]);
    const result = await convert('doc', word(B, pack, { tables: [[60, Buffer.concat([names, value])]] }));
    assert.equal(result.type, 'docm');
    assert.equal(macroLoss(result).some(l => l.what.includes('digital signature')), signed);
  }
});

test('Word and Excel password handling leaves the separate VBA storage unchanged', async t => {
  const files = encrypted.filter(([format, name]) =>
    format === 'doc' && name.includes('binaryrc4') || format === 'xls' && name.endsWith('__password.xls'));
  if (files.some(([, name]) => !fs.existsSync('/work/corpora/' + name)))
    return t.skip('Pinned legacy corpus is not installed');
  for (const [format, name, password] of files) {
    const original = B.cfb(fs.readFileSync('/work/corpora/' + name));
    const streams = Object.fromEntries(original.entries.filter(e => e.type === 2)
      .map(e => [e.path, original.stream(e.path)]));
    const prefix = format === 'doc' ? 'Macros/' : '_VBA_PROJECT_CUR/';
    for (const [path, bytes] of Object.entries(project())) streams[prefix + path] = bytes;
    const bytes = pack(source(streams));
    await assert.rejects(B[format].convert(bytes, { password: 'wrong' }), { code: 'WRONG_PASSWORD' });
    const result = await convert(format, bytes, { password });
    assert.equal(result.type, format === 'doc' ? 'docm' : 'xlsm');
    const part = (format === 'doc' ? 'word' : 'xl') + '/vbaProject.bin';
    assert.deepEqual(snapshot(B.cfb(await result.zip.get(part).bytes())), snapshot(source(project())));
  }
});
