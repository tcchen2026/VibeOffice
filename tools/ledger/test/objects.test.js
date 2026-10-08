/* Worksheet entries, hidden previews, VML and dependent parts remain one object. */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} };
require('../../../public/ledger/js/ops.js');
const K = L.opc, O = L.ops, output = process.env.OBJECT_RESULTS, results = [];
const corpus = process.env.VO_CORPORA || path.join(os.homedir(), 'corpora');
const fixtures = { control: 'libreoffice__1b04d082bd6e__button-form-control.xlsx', ole: 'poi__2a89697e0dc6__58325_lt.xlsx',
  named: 'libreoffice__8d005aca2c1d__tdf161365.xlsx', mixed: 'poi__9028285d4065__bug66827.xlsx',
  vml: 'libreoffice__9f0dba934b74__tdf166724_cellAnchor.xlsx', properties: 'poi__448351c00314__60512.xlsm' };
async function open(t, kind) {
  const file = path.join(corpus, 'excel', fixtures[kind]);
  if (!fs.existsSync(file)) { t.skip('Fetch pinned corpus: ' + file); return; }
  const wb = await L.xlsxRead.read(fs.readFileSync(file)); Object.defineProperty(wb, '_input', { value: file }); return wb;
}
async function save(wb, scenario) {
  const bytes = await L.xlsxWrite.bytes(wb), pkg = await K.open(bytes);
  if (output) {
    const dir = path.join(output, scenario); fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, path.basename(wb._input)), bytes);
    results.push({ file: path.basename(wb._input), original: wb._input, scenario, status: 'ok' });
    fs.writeFileSync(path.join(output, 'results.jsonl'), results.map(r => JSON.stringify(r)).join('\n') + '\n');
  }
  return { bytes, pkg, losses: bytes.dropped };
}
const vmlPart = pkg => pkg.names.find(n => n.endsWith('.vml'));
const all = (pkg, part) => { const e = pkg.xml(part); return [e, ...e.getElementsByTagName('*')]; };
const sheet = pkg => pkg.names.find(n => /^xl\/worksheets\/sheet\d+\.xml$/.test(n));
const references = (pkg, tag) => all(pkg, sheet(pkg)).filter(e => e.localName === tag).map(e => e.getAttribute('shapeId'));
const caption = pkg => all(pkg, vmlPart(pkg)).filter(e => e.localName === 'ClientData').map(e => e.getAttribute('ObjectType'));
const equalBytes = (a, b, message) => assert.deepEqual(Buffer.from(a), Buffer.from(b), message);

test('controls, OLE, mixed VML and sheet custom properties retain XML dependencies on ordinary saves', async t => {
  for (const kind of Object.keys(fixtures)) {
    const wb = await open(t, kind); if (!wb) return;
    const s = await save(wb, 'save');
    if (vmlPart(wb.pkg)) equalBytes(s.pkg.bytes(vmlPart(wb.pkg)), wb.pkg.bytes(vmlPart(wb.pkg)), kind + ' VML');
    for (const part of wb.pkg.names.filter(n => /\/embeddings\/|\/ctrlProps\/|\/activeX\//.test(n))) equalBytes(s.pkg.bytes(part), wb.pkg.bytes(part), part);
    for (const tag of ['control', 'oleObject', 'customPr']) assert.equal(all(s.pkg, sheet(s.pkg)).filter(e => e.localName === tag).length, all(wb.pkg, sheet(wb.pkg)).filter(e => e.localName === tag).length, kind + ' ' + tag);
    assert.equal(s.losses.length, 0, kind);
  }
});
for (const kind of ['control', 'ole']) test(kind + ': row changes update worksheet, drawing and VML anchors; undo restores them', async t => {
  const wb = await open(t, kind); if (!wb) return;
  const sh = wb.sheets[0], d = sh.drawings.find(d => d.objectKeep), r = d.anchor.from.r, end = d.anchor.to.r;
  O.insertLines(sh, 'r', 0, 2); const s = await save(wb, 'rows');
  assert.equal(d.anchor.to.r, end + 2, 'fixed-size objects retain their lower anchor when moved');
  const anchors = all(s.pkg, sheet(s.pkg)).filter(e => e.localName === 'anchor');
  assert(anchors.length); assert(anchors.every(e => +e.getElementsByTagName('*').find(n => n.localName === 'row').textContent === r + 2));
  const va = all(s.pkg, vmlPart(s.pkg)).find(e => e.localName === 'Anchor'); assert.equal(+va.textContent.split(',')[2], r + 2);
  const reread = await L.xlsxRead.read(s.bytes); assert.equal(reread.sheets[0].drawings[0].anchor.from.r, r + 2);
  wb.undo.undo(); const u = await save(wb, 'rows-undo'); equalBytes(u.pkg.bytes(vmlPart(wb.pkg)), wb.pkg.bytes(vmlPart(wb.pkg)));
  wb.undo.redo(); await save(wb, 'rows-redo');
});
test('deleting or changing a displayed object removes its control/embedding and VML; undo restores payloads', async t => {
  const wb = await open(t, 'ole'); if (!wb) return;
  const sh = wb.sheets[0], d = sh.drawings[0];
  O.tx(wb, 'Delete object', () => O.setDrawings(sh, []));
  let s = await save(wb, 'delete'); assert.equal(references(s.pkg, 'oleObject').length, 0); assert(!vmlPart(s.pkg));
  assert(!s.pkg.names.some(n => /\/embeddings\//.test(n)));
  wb.undo.undo(); s = await save(wb, 'delete-undo'); assert.equal(references(s.pkg, 'oleObject').length, 2);
  O.tx(wb, 'Edit object', () => O.setDrawings(sh, [{ ...d, dirty: true, text: 'Converted preview' }]));
  s = await save(wb, 'content-edit'); assert.equal(references(s.pkg, 'oleObject').length, 0);
  assert(s.losses.some(e => e.id.startsWith('worksheet-object:')));
});
for (const kind of ['control', 'ole', 'named']) test(kind + ': copying an object keeps independent payloads and remaps linked identities', async t => {
  const wb = await open(t, kind); if (!wb) return;
  const sh = wb.sheets[0], d = L.sheetObjects.copy(sh.drawings[0]); d.id = Math.max(...sh.drawings.map(d => d.id)) + 1;
  const tag = kind === 'ole' ? 'oleObject' : 'control', ids = [...new Set(references(wb.pkg, tag)), String(d.id)];
  d.anchor.from.r += 10; d.anchor.to.r += 10;
  O.tx(wb, 'Copy object', () => O.setDrawings(sh, sh.drawings.concat(d)));
  const s = await save(wb, 'copy');
  assert.deepEqual(new Set(references(s.pkg, tag)), new Set(ids));
  const shapes = all(s.pkg, vmlPart(s.pkg)).filter(e => e.localName === 'shape');
  assert.deepEqual(new Set(shapes.map(e => e.getAttributeNS(K.KNOWN_NS.o, 'spid') || e.getAttribute('id'))), new Set(ids.map(id => '_x0000_s' + id)));
  assert.equal(new Set(shapes.map(e => e.getAttribute('id'))).size, shapes.length);
  const drawing = s.pkg.names.find(n => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
  assert.deepEqual(all(s.pkg, drawing).filter(e => e.localName === 'cNvPr').map(e => e.getAttribute('id')), ids);
  assert.deepEqual(all(s.pkg, drawing).filter(e => e.localName === 'compatExt').map(e => e.getAttribute('spid')), ids.map(id => '_x0000_s' + id));
  const payloads = pkg => pkg.names.filter(n => /\/(ctrlProps|embeddings)\//.test(n));
  assert.equal(payloads(s.pkg).length, payloads(wb.pkg).length + 1, 'the copied payload is independently owned');
  wb.undo.undo(); const u = await save(wb, 'copy-undo'); equalBytes(u.pkg.bytes(vmlPart(wb.pkg)), wb.pkg.bytes(vmlPart(wb.pkg)));
});
test('copied sheets have independent VML parts; deleting the source retains the copied controls', async t => {
  const wb = await open(t, 'control'); if (!wb) return;
  const original = wb.sheets[0]; O.copySheet(wb, original, 1);
  let s = await save(wb, 'sheet-copy'); assert.equal(s.pkg.names.filter(n => n.endsWith('.vml')).length, 2);
  O.deleteSheet(wb, original); s = await save(wb, 'source-sheet-delete');
  assert.equal(references(s.pkg, 'control').length, 1); assert.equal(caption(s.pkg)[0], 'Button');
});
test('non-comment VML survives note changes and removal without rebuilding its shape', async t => {
  const wb = await open(t, 'mixed'); if (!wb) return;
  const sh = wb.sheets[0], note = [...sh.comments.values()][0];
  O.tx(wb, 'Delete note', () => O.setComment(sh, note.r, note.c, null));
  let s = await save(wb, 'note-delete'); assert.deepEqual(caption(s.pkg), ['Text']);
  wb.undo.undo(); s = await save(wb, 'note-delete-undo'); equalBytes(s.pkg.bytes(vmlPart(wb.pkg)), wb.pkg.bytes(vmlPart(wb.pkg)));
  O.tx(wb, 'Insert note', () => O.setComment(sh, 9, 3, { r: 9, c: 3, author: 'Test', text: 'New note' }));
  s = await save(wb, 'note-add'); assert.deepEqual(caption(s.pkg).sort(), ['Note', 'Note', 'Text']);
});
