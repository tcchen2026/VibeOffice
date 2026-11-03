/* Objects the reader can't show (SmartArt, an add-in shown by its picture) come back as they were. */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} };
require('../../public/ledger/js/ops.js');
const K = L.opc, O = L.ops;
const corpus = process.env.VO_CORPORA || path.join(os.homedir(), 'corpora');
const fixtures = { smartart: 'libreoffice__12275492e8fc__tdf83671_SmartArt_import.xlsx', addin: 'openxml-sdk__03f833d57306__Youtube.xlsx',
  strict: 'openxml-sdk__97ec4a7d7ff8__Photo Formats - CGM-O12-XL-Pictures.xlsx', empty: 'libreoffice__541359181211__tdf140968.xlsx',
  commentShape: 'libreoffice__761105c9e678__universal-content-strict.xlsx' };
async function open(t, kind) {
  const file = path.join(corpus, 'excel', fixtures[kind]);
  if (!fs.existsSync(file)) { t.skip('Fetch pinned corpus: ' + file); return {}; }
  const bytes = fs.readFileSync(file);
  return { wb: await L.xlsxRead.read(bytes), original: await K.open(bytes) };
}
async function save(wb) { const bytes = await L.xlsxWrite.bytes(wb); return { pkg: await K.open(bytes), losses: wb.losses || [] }; }
const ids = losses => losses.map(e => e.id.split(':')[0]);
const text = (pkg, part) => new TextDecoder().decode(pkg.bytes(part));
const drawingOf = pkg => pkg.names.find(n => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
const diagrams = pkg => pkg.names.filter(n => /^xl\/diagrams\/[a-zA-Z]+\d+\.xml$/.test(n)).sort();

test('an unedited save keeps a SmartArt frame and its diagram parts byte for byte', async t => {
  const { wb, original } = await open(t, 'smartart'); if (!wb) return;
  const d = wb.sheets[0].drawings[0];
  assert.equal(d.kind, 'frame'); assert.equal(d.label, 'IntelliArt');
  const { pkg, losses } = await save(wb);
  assert.ok(!ids(losses).includes('unreferenced'), losses.map(e => e.id).join('\n'));
  for (const part of diagrams(original)) assert.deepEqual(pkg.bytes(part), original.bytes(part), part);
  const frame = [...pkg.xml(drawingOf(pkg)).getElementsByTagName('*')].find(e => e.localName === 'relIds');
  const rels = Object.fromEntries(pkg.rels(drawingOf(pkg)).map(r => [r.id, r.part]));
  assert.equal(rels[frame.getAttributeNS(K.NS.rel, 'dm')], 'xl/diagrams/data1.xml');
  assert.match(text(pkg, drawingOf(pkg)), /creationId[^>]*\{A50947B9-D0A6-1C6F-F56C-6F4847174B4A\}/);
  // the data part names the drawing cache through dataModelExt/@relId, a relationship of the sheet drawing
  const ext = [...pkg.xml('xl/diagrams/data1.xml').getElementsByTagName('*')].find(e => e.localName === 'dataModelExt');
  assert.equal(rels[ext.getAttribute('relId')], 'xl/diagrams/drawing1.xml');
});

test('a moved SmartArt keeps its frame; a copy gets diagram parts of its own; a deleted one takes them along', async t => {
  const { wb } = await open(t, 'smartart'); if (!wb) return;
  const sh = wb.sheets[0];
  O.tx(wb, 'Move', () => { const list = L.clone(sh.drawings); list[0].anchor.from.c += 2; list[0].anchor.to.c += 2; O.setDrawings(sh, list); });
  let { pkg, losses } = await save(wb);
  assert.ok(!ids(losses).includes('unreferenced'));
  assert.match(text(pkg, drawingOf(pkg)), /<xdr:col>2<\/xdr:col>.*<dgm:relIds/s);
  O.tx(wb, 'Duplicate', () => O.setDrawings(sh, sh.drawings.concat(L.clone(sh.drawings[0]))));
  ({ pkg, losses } = await save(wb));
  assert.ok(!ids(losses).includes('unreferenced'));
  assert.deepEqual(diagrams(pkg).map(n => n.replace(/\d+\.xml$/, '')), ['colors', 'colors', 'data', 'data', 'drawing', 'drawing', 'layout', 'layout', 'quickStyle', 'quickStyle'].map(n => 'xl/diagrams/' + n));
  const drawing = drawingOf(pkg), rels = Object.fromEntries(pkg.rels(drawing).map(r => [r.id, r.part]));
  const frames = [...pkg.xml(drawing).getElementsByTagName('*')].filter(e => e.localName === 'relIds');
  const data = frames.map(f => rels[f.getAttributeNS(K.NS.rel, 'dm')]);
  assert.equal(new Set(data).size, 2, 'each frame has its own data part');
  for (const part of data) {
    const ext = [...pkg.xml(part).getElementsByTagName('*')].find(e => e.localName === 'dataModelExt');
    assert.ok(rels[ext.getAttribute('relId')]?.startsWith('xl/diagrams/drawing'), part + ' names a drawing cache of this sheet drawing');
  }
  assert.equal(new Set(data.map(part => [...pkg.xml(part).getElementsByTagName('*')].find(e => e.localName === 'dataModelExt').getAttribute('relId'))).size, 2);
  const cnv = [...pkg.xml(drawing).getElementsByTagName('*')].filter(e => e.localName === 'cNvPr').map(e => e.getAttribute('id'));
  assert.equal(new Set(cnv).size, 2, 'the copy has an id of its own');
  O.tx(wb, 'Delete', () => O.setDrawings(sh, []));
  ({ pkg, losses } = await save(wb));
  assert.deepEqual(diagrams(pkg), []);
  assert.ok(ids(losses).includes('frame-deleted')); assert.ok(!ids(losses).includes('unreferenced'), losses.map(e => e.id).join('\n'));
});

test('an add-in shown by its picture keeps its AlternateContent until the picture is edited', async t => {
  const { wb } = await open(t, 'addin'); if (!wb) return;
  const sh = wb.sheets[0], d = sh.drawings[0];
  assert.equal(d.kind, 'image');
  let { pkg, losses } = await save(wb);
  assert.match(text(pkg, drawingOf(pkg)), /<mc:Choice Requires="a15">.*webextension.*<mc:Fallback><xdr:pic>/s);
  assert.ok(pkg.has('xl/webextensions/webextension1.xml')); assert.ok(!ids(losses).includes('unreferenced'));
  O.tx(wb, 'Move', () => { const list = L.clone(sh.drawings); list[0].anchor.from.r += 1; list[0].anchor.to.r += 1; O.setDrawings(sh, list); });
  ({ pkg } = await save(wb));
  assert.match(text(pkg, drawingOf(pkg)), /webextension/, 'moving keeps it');
  O.tx(wb, 'Crop', () => { const list = L.clone(sh.drawings); list[0].crop = { l: 10, t: 0, r: 0, b: 0 }; O.setDrawings(sh, list); });
  ({ pkg, losses } = await save(wb));
  assert.doesNotMatch(text(pkg, drawingOf(pkg)), /AlternateContent|webextension/);
  const notice = losses.find(e => e.id.startsWith('alternate-edit:'));
  assert.equal(notice?.notify, true);
});

test('Strict SmartArt comes back as Transitional with its picture; an empty drawing goes quietly', async t => {
  const { wb, original } = await open(t, 'strict'); if (!wb) return;
  assert.equal(wb.sheets.flatMap(s => s.drawings).filter(d => d.kind === 'frame' && d.label === 'IntelliArt').length, 2);
  const { pkg, losses } = await save(wb);
  assert.ok(!ids(losses).includes('unreferenced'), losses.filter(e => e.id.startsWith('unreferenced:')).map(e => e.where).join(' '));
  assert.deepEqual(diagrams(pkg), diagrams(original));
  assert.ok(pkg.has('xl/media/image6.emf'));
  assert.doesNotMatch(diagrams(pkg).map(n => text(pkg, n)).join(''), /purl\.oclc\.org/);
  const empty = await open(t, 'empty'); if (!empty.wb) return;
  const saved = await save(empty.wb);
  assert.ok(!ids(saved.losses).includes('unreferenced')); assert.ok(saved.losses.some(e => e.id.startsWith('part:') && /drawing/.test(e.where) && e.notify === false));
});

// LibreOffice draws a comment in VML and, as the Fallback of an AlternateContent with an empty VML Choice, in DrawingML;
// the comment's commentPr names that shape, and Excel 2016 refuses the file without it (tools/ooxml/package.py,
// comment-shape). The reader took the empty Choice, so the shape went.
test('a comment\'s DrawingML shape in an AlternateContent stays, also when the drawing\'s other objects go', async t => {
  const { wb } = await open(t, 'commentShape'); if (!wb) return;
  const fallbackShapes = pkg => {
    const ac = [...pkg.xml(drawingOf(pkg)).children].filter(e => e.localName === 'AlternateContent');
    return ac.flatMap(e => [...e.getElementsByTagName('*')].filter(c => c.localName === 'cNvPr').map(c => c.getAttribute('id')));
  };
  let { pkg } = await save(wb);
  assert.deepEqual(fallbackShapes(pkg), ['1025']);
  assert.match(text(pkg, 'xl/comments1.xml'), /shapeId="1025"/);
  O.tx(wb, 'Delete', () => O.setDrawings(wb.sheets[0], []));
  ({ pkg } = await save(wb));
  assert.deepEqual(fallbackShapes(pkg), ['1025'], 'the comment\'s shape outlives the drawing\'s other objects');
});
