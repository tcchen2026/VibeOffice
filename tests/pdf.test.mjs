// The shared PDF writer: check its bytes with independent PDF readers, not just our object strings.
import { test, before, after } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import vm from 'node:vm';
import { execFileSync } from 'node:child_process';
import { inflateSync } from 'node:zlib';

const jpeg = new Uint8Array(Buffer.from([
  '/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMU',
  'FRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQU',
  'FBQUFBQUFBT/wAARCAAEAAQDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUF',
  'BAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVW',
  'V1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi',
  '4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAEC',
  'AxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVm',
  'Z2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq',
  '8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD7rooorxD2D//Z',
].join(''), 'base64'));
const fixed = { title: 'PDF 日本語 😀', author: 'VibeOffice tests', date: '2020-01-02T03:04:05Z', lang: 'en-US',
  outline: [{ title: 'First', page: 0, top: 20, level: 1 },
    { title: 'Child 日本語', page: 1, top: 40, level: 2 }, { title: 'Second', page: 1, level: 1 }] };
const pages = [
  { jpeg, iw: 4, ih: 4, pw: 240, ph: 320, lines: [
    [{ text: 'Hello', x: 20, y: 20, w: 40, h: 12 }, { text: 'PDF', x: 66, y: 20, w: 24, h: 12 }],
    [{ text: '日本語', x: 20, y: 42, w: 36, h: 12 }, { text: '😀', x: 62, y: 42, w: 12, h: 12 }],
    [{ text: 'שלום', x: 20, y: 64, w: 32, h: 12 }],
    [{ text: 'Turn', x: 200, y: 100, w: 32, h: 12, angle: 90 }],
  ], links: [{ x: 20, y: 20, w: 70, h: 12, url: 'https://vibeoffice.work/a%20b?q=(c)' },
    { x: 20, y: 42, w: 54, h: 12, page: 1, top: 40 }] },
  { jpeg, iw: 4, ih: 4, pw: 240, ph: 320, lines: [
    [{ text: 'Second', x: 20, y: 20, w: 48, h: 12 }, { text: 'needle', x: 74, y: 20, w: 48, h: 12 }],
  ] },
];
let dir, L, bytes, plainBytes;
async function load(compress = true) {
  const root = process.env.PDF_PUBLIC || 'public';
  const c = { Blob, Response, TextEncoder, TextDecoder, Uint8Array, Date, URL,
    CompressionStream: compress ? globalThis.CompressionStream : undefined };
  c.window = c;
  for (const file of ['common/core.js', 'common/sha.js']) {
    let source;
    try { source = await fs.readFile(path.join(root, file), 'utf8'); }
    catch { source = await fs.readFile(path.join('public', file), 'utf8'); }
    vm.runInNewContext(source, c, { filename: file });
  }
  // A browser's first export loads this through core.js. The Node fixture has no DOM.
  if (!process.env.PDF_PUBLIC) {
    try { vm.runInNewContext(await fs.readFile('public/common/pdf.js', 'utf8'), c, { filename: 'pdf.js' }); }
    catch (e) { if (e.code !== 'ENOENT') throw e; }
  }
  return c.L;
}
const buffer = async blob => Buffer.from(await blob.arrayBuffer());
function objects(data) {
  const text = data.toString('latin1'), result = new Map();
  const start = Number(/startxref\s+(\d+)/.exec(text)[1]);
  assert.equal(text.slice(start, start + 4), 'xref');
  const entries = text.slice(start).split('\n'), count = Number(entries[1].split(' ')[1]);
  for (let id = 1; id < count; id++) {
    const offset = Number(entries[id + 2].slice(0, 10));
    assert.equal(text.slice(offset, offset + String(id).length + 6), id + ' 0 obj');
    const end = text.indexOf('\nendobj', offset);
    result.set(id, data.subarray(offset + String(id).length + 7, end));
  }
  return result;
}
function stream(body) {
  const text = body.toString('latin1'), begin = text.indexOf('stream\n') + 7;
  const length = Number(/\/Length (\d+)/.exec(text)[1]);
  const data = body.subarray(begin, begin + length);
  return text.slice(0, begin).includes('/FlateDecode') ? inflateSync(data) : data;
}
const matching = (data, rx) => [...objects(data).values()].filter(b => rx.test(b.toString('latin1')));

before(async () => {
  dir = await fs.mkdtemp(path.join(os.tmpdir(), 'vibeoffice-pdf-'));
  L = await load();
  bytes = await buffer(await L.buildPDF(pages, fixed));
  plainBytes = await buffer(await L.buildPDF(pages.map(({ lines, ...p }) => p), fixed));
  await fs.writeFile(path.join(dir, 'text.pdf'), bytes);
  await fs.writeFile(path.join(dir, 'pictures.pdf'), plainBytes);
});
after(async () => { if (dir) await fs.rm(dir, { recursive: true, force: true }); });

test('PDF 1.7 is asynchronous, deterministic, and every xref offset is correct', async () => {
  const pending = L.buildPDF(pages, fixed);
  assert.equal(typeof pending.then, 'function');
  assert.equal(bytes.subarray(0, 8).toString(), '%PDF-1.7');
  assert.deepEqual(await buffer(await pending), bytes);
  assert.ok(objects(bytes).size > 10);
  const text = bytes.toString('latin1');
  assert.match(text, /\/CreationDate \(D:20200102030405Z\)/);
  const id = /\/ID \[<([0-9a-f]+)> <\1>\]/i.exec(text);
  assert.ok(id && id[1].length >= 32);
  const changed = await buffer(await L.buildPDF(pages, { ...fixed, title: 'Changed' }));
  assert.notEqual(/\/ID \[<([^>]+)/.exec(changed.toString('latin1'))[1], id[1]);
});

test('pdftotext extracts words, single spaces, Unicode scalars and lines in reading order', () => {
  const text = execFileSync('pdftotext', ['-raw', path.join(dir, 'text.pdf'), '-'], { encoding: 'utf8' })
    .replace(/[\u202a-\u202e\u2066-\u2069]/g, '');
  assert.ok(text.startsWith('Hello PDF\n日本語 😀\n'), JSON.stringify(text));
  assert.ok(text.includes('שלום'), JSON.stringify(text));
  assert.ok(text.includes('Turn'), JSON.stringify(text));
  assert.ok(text.endsWith('Second needle\n\f'), JSON.stringify(text));
});

test('pdffonts sees one embedded Unicode font with a valid empty TrueType program', () => {
  const font = execFileSync('pdffonts', [path.join(dir, 'text.pdf')], { encoding: 'utf8' });
  assert.match(font, /VibeOfficeText\s+CID TrueType\s+Identity-H\s+yes\s+no\s+yes/);
  assert.equal(font.trim().split('\n').length, 3);
  const program = stream(matching(bytes, /\/Length1 /)[0]);
  let sum = 0;
  for (let n = 0; n < program.length; n += 4) sum = (sum + program.readUInt32BE(n)) >>> 0;
  assert.equal(sum, 0xb1b0afba, 'OpenType whole-font checksum');
  const count = program.readUInt16BE(4), tables = new Map();
  for (let n = 0; n < count; n++) {
    const p = 12 + n * 16, tag = program.toString('ascii', p, p + 4);
    tables.set(tag, program.subarray(program.readUInt32BE(p + 8),
      program.readUInt32BE(p + 8) + program.readUInt32BE(p + 12)));
  }
  assert.deepEqual([...tables.keys()].sort(),
    ['OS/2', 'cmap', 'glyf', 'head', 'hhea', 'hmtx', 'loca', 'maxp', 'name', 'post']);
  assert.equal(tables.get('maxp').readUInt16BE(4), 2);
  assert.equal(tables.get('hhea').readInt16BE(4), 800);
  assert.equal(tables.get('hhea').readInt16BE(6), -200);
  assert.equal(tables.get('glyf').readInt16BE(0), 0);
  assert.equal(tables.get('glyf').readInt16BE(12), 1);
});

test('PDFium extracts single characters, emoji and vertical words from the generated empty font', () => {
  const output = execFileSync('python3', ['-I', '-c',
    'import pypdfium2 as p,sys,json; d=p.PdfDocument(sys.argv[1]); ' +
    'print(json.dumps([page.get_textpage().get_text_bounded() for page in d]))', path.join(dir, 'text.pdf')],
  { encoding: 'utf8' });
  const text = JSON.parse(output);
  assert.ok(text[0].includes('日本語 😀'), JSON.stringify(text));
  assert.ok(text[0].includes('שלום'), JSON.stringify(text));
  assert.ok(text[0].includes('Turn'), JSON.stringify(text));
  assert.ok(text[1].includes('Second needle'), JSON.stringify(text));
});

test('the generated metric contour paints nothing even if visible text mode is requested', async () => {
  const uncompressed = await buffer(await (await load(false)).buildPDF(pages, fixed));
  // Replacing the one-byte rendering mode leaves all xref offsets intact.
  const visible = Buffer.from(uncompressed.toString('latin1').replaceAll('BT 3 Tr', 'BT 0 Tr'), 'latin1');
  await fs.writeFile(path.join(dir, 'empty-font.pdf'), visible);
  for (const name of ['empty-font', 'pictures']) {
    execFileSync('pdftoppm', ['-r', '72', path.join(dir, name + '.pdf'), path.join(dir, name)]);
  }
  for (const page of [1, 2]) assert.deepEqual(await fs.readFile(path.join(dir, 'empty-font-' + page + '.ppm')),
    await fs.readFile(path.join(dir, 'pictures-' + page + '.ppm')));
  const same = execFileSync('python3', ['-I', '-c',
    'import pypdfium2 as p,sys; a=p.PdfDocument(sys.argv[1]); b=p.PdfDocument(sys.argv[2]); ' +
    'assert all(bytes(x.render().buffer)==bytes(y.render().buffer) for x,y in zip(a,b))',
    path.join(dir, 'empty-font.pdf'), path.join(dir, 'pictures.pdf')]);
  assert.equal(same.length, 0);
});

test('the invisible layer changes no pixels and the original JPEG bytes survive', async () => {
  for (const name of ['text', 'pictures']) {
    execFileSync('pdftoppm', ['-r', '72', path.join(dir, name + '.pdf'), path.join(dir, name)]);
  }
  for (const page of [1, 2]) assert.deepEqual(
    await fs.readFile(path.join(dir, 'text-' + page + '.ppm')),
    await fs.readFile(path.join(dir, 'pictures-' + page + '.ppm')));
  const images = matching(bytes, /\/DCTDecode/);
  assert.equal(images.length, 2);
  for (const image of images) assert.deepEqual(stream(image), Buffer.from(jpeg));
});

test('links, destinations, nested outlines, language and document title are written', () => {
  const all = [...objects(bytes).values()].map(b => b.toString('latin1')).join('\n');
  assert.match(all, /\/Type \/Catalog[^]*\/Lang /);
  assert.match(all, /\/DisplayDocTitle true/);
  assert.equal(matching(bytes, /\/Subtype \/Link/).length, 2);
  assert.match(all, /\/Rect \[20 288 90 300\]/);
  assert.match(all, /\/S \/URI/);
  assert.ok(all.includes('https://vibeoffice.work/a%20b?q=\\(c\\)'), 'URI escapes survive exactly once');
  assert.match(all, /\/Dest \[\d+ 0 R \/XYZ 0 280 null\]/);
  assert.match(all, /\/Type \/Outlines \/First \d+ 0 R \/Last \d+ 0 R \/Count 3/);
  assert.match(all, /\/Parent \d+ 0 R[^]*\/Count 1/);
});

test('selection boxes have the supplied point coordinates and text height', () => {
  const xml = execFileSync('pdftotext', ['-bbox', path.join(dir, 'text.pdf'), '-'], { encoding: 'utf8' });
  const word = /<word xMin="([^"]+)" yMin="([^"]+)" xMax="([^"]+)" yMax="([^"]+)">Hello<\/word>/.exec(xml);
  assert.ok(word, xml);
  assert.deepEqual(word.slice(1).map(Number), [20, 20, 60, 32]);
});

test('right-to-left words retain letters, marks, numbers and embedded Latin text', async () => {
  const words = ['שָׁלוֹם', 'مرحبا', 'שלום123', 'אב12גד', 'שלוםABC', 'abcשלום', 'مرحبا١٢٣'];
  const page = { jpeg, iw: 4, ih: 4, pw: 240, ph: 320, lines: words.map((text, n) =>
    [{ text, x: 20, y: 20 + n * 30, w: 150, h: 16 }]) };
  const file = path.join(dir, 'rtl.pdf');
  await fs.writeFile(file, await buffer(await L.buildPDF([page], fixed)));
  const text = execFileSync('pdftotext', ['-raw', file, '-'], { encoding: 'utf8' })
    .replace(/[\u202a-\u202e\u2066-\u2069]/g, '').trim();
  assert.deepEqual(text.split('\n'), words);
});

test('right-to-left lines copy in reading order with a space between their word boxes', async () => {
  const page = { jpeg, iw: 4, ih: 4, pw: 240, ph: 320, lines: [[
    { text: 'שלום', x: 140, y: 20, w: 40, h: 20 }, { text: 'עולם', x: 90, y: 20, w: 40, h: 20 },
  ]] };
  const file = path.join(dir, 'rtl-line.pdf');
  await fs.writeFile(file, await buffer(await L.buildPDF([page], fixed)));
  const text = execFileSync('pdftotext', [file, '-'], { encoding: 'utf8' })
    .replace(/[\u202a-\u202e\u2066-\u2069]/g, '').trim();
  assert.equal(text, 'שלום עולם');
});

test('the writer loads only on export, shares concurrent loads and retries a failed load', async () => {
  const c = { Blob, Response, CompressionStream, TextEncoder, TextDecoder, Uint8Array, Date,
    VO: { base: 'http://127.0.0.1:8769/' } };
  c.window = c;
  const scripts = [], source = await fs.readFile('public/common/pdf.js', 'utf8');
  c.document = { documentElement: { dataset: {}, getAttribute: () => 'en' },
    createElement: () => ({ remove() {} }), head: { append: script => scripts.push(script) } };
  for (const file of ['core', 'sha']) {
    let body;
    try { body = await fs.readFile(path.join(process.env.PDF_PUBLIC || 'public', 'common', file + '.js'), 'utf8'); }
    catch { body = await fs.readFile('public/common/' + file + '.js', 'utf8'); }
    vm.runInNewContext(body, c);
  }
  c.L.t = s => s;
  assert.equal(scripts.length, 0);
  const failed = c.L.buildPDF(pages, fixed);
  assert.equal(scripts.length, 1);
  const rejection = assert.rejects(failed, /exporter could not be loaded/);
  scripts[0].onerror();
  await rejection;
  const first = c.L.buildPDF(pages, fixed), second = c.L.buildPDF(pages, fixed);
  assert.equal(scripts.length, 2);
  assert.equal(scripts[1].src, 'http://127.0.0.1:8769/common/pdf.js');
  vm.runInNewContext(source, c);
  scripts[1].onload();
  const result = await Promise.all([first, second]);
  assert.deepEqual(await buffer(result[0]), await buffer(result[1]));
});

test('Flate streams have a working plain fallback when CompressionStream is absent', async () => {
  assert.ok(matching(bytes, /\/FlateDecode/).length >= 5);
  const plain = await load(false), uncompressed = await buffer(await plain.buildPDF(pages, fixed));
  assert.equal(matching(uncompressed, /\/FlateDecode/).length, 0);
  const file = path.join(dir, 'uncompressed.pdf');
  await fs.writeFile(file, uncompressed);
  assert.match(execFileSync('pdftotext', ['-raw', file, '-'], { encoding: 'utf8' }), /Hello PDF/);
  assert.deepEqual(matching(bytes, /\/Length1 /).map(stream), matching(uncompressed, /\/Length1 /).map(stream));
});
