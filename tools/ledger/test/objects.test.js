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
  vml: 'libreoffice__9f0dba934b74__tdf166724_cellAnchor.xlsx', properties: 'poi__448351c00314__60512.xlsm',
  notes: 'poi__c613ea5875fa__SimpleWithComments.xlsx', grouped: 'poi__90ce92f078bd__45540_form_Header.xlsx',
  activex: 'libreoffice__e80921817d6f__activex_checkbox.xlsx' };
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

// Spreadsheet VML: every part has its own idmap blocks and numbers its shapes inside them.
const blocks = pkg => pkg.names.filter(n => n.endsWith('.vml')).map(part => {
  const els = all(pkg, part), idmap = els.find(e => e.localName === 'idmap').getAttribute('data').split(',').map(Number);
  return { part, idmap, ids: els.filter(e => e.namespaceURI === K.KNOWN_NS.v && e.localName !== 'shapetype').map(e => e.getAttributeNS(K.KNOWN_NS.o, 'spid') || e.getAttribute('id')).filter(id => /^_x0000_s\d+$/.test(id || '')).map(id => +id.slice(8)) };
});
const check = pkg => {
  const list = blocks(pkg), maps = list.flatMap(b => b.idmap), ids = list.flatMap(b => b.ids);
  assert.equal(new Set(maps).size, maps.length, 'unique idmap blocks ' + JSON.stringify(list.map(b => [b.part, b.idmap])));
  assert.equal(new Set(ids).size, ids.length, 'unique note shape ids');
  for (const b of list) for (const id of b.ids) assert(b.idmap.includes(Math.floor(id / 1024)), id + ' outside ' + b.part + ' block ' + b.idmap);
  return list;
};
test('each sheet\'s notes use their own VML id block, across new, original and copied sheets', async t => {
  const wb = await open(t, 'notes'); if (!wb) return;
  const original = await save(wb, 'notes-save');
  assert.deepEqual(check(original.pkg).map(b => b.idmap), [[1]]);
  const fresh = O.addSheet(wb, 'New notes', 0);
  O.setComment(fresh, 0, 0, { author: 'Test', text: 'first', visible: false });
  O.setComment(fresh, 1, 1, { author: 'Test', text: 'second', visible: false });
  O.copySheet(wb, wb.sheets[1], wb.sheets.length);
  const saved = await save(wb, 'notes-new-and-copied');
  assert.equal(check(saved.pkg).length, 3);
  equalBytes(saved.pkg.bytes(vmlPart(wb.pkg)), wb.pkg.bytes(vmlPart(wb.pkg)), 'the original sheet keeps its VML bytes');
});

test('a sheet with more notes than one VML block holds gets further blocks, never another sheet\'s', async t => {
  const wb = await open(t, 'notes'); if (!wb) return;
  const many = O.addSheet(wb, 'Many notes', 0);
  for (let r = 0; r < 1025; r++) many.comments.set(L.model.key(r, 0), { r, c: 0, author: 'Test', text: 'note ' + r, visible: false });
  const original = wb.sheets[1];
  for (let r = 0; r < 1100; r++) original.comments.set(L.model.key(r, 9), { r, c: 9, author: 'Test', text: 'more ' + r, visible: false });
  O.copySheet(wb, wb.sheets[1], wb.sheets.length);
  const list = check((await save(wb, 'notes-many-blocks')).pkg);
  assert.equal(list.reduce((n, b) => n + b.ids.length, 0), 1025 + 2 * (1100 + 3));
  assert(list.some(b => b.idmap.length > 1));
});

test('retained notes need no new ids: a reopened sheet keeps its single block and saves unchanged', async t => {
  const wb = await open(t, 'notes'); if (!wb) return;
  const sh = O.addSheet(wb, 'Half block', 0);
  for (let r = 0; r < 512; r++) sh.comments.set(L.model.key(r, 0), { r, c: 0, author: 'Test', text: 'note ' + r, visible: false });
  const first = await save(wb, 'notes-512');
  const reopened = await L.xlsxRead.read(first.bytes); Object.defineProperty(reopened, '_input', { value: wb._input });
  const second = await save(reopened, 'notes-512-resave');
  for (const part of first.pkg.names.filter(n => n.endsWith('.vml'))) equalBytes(second.pkg.bytes(part), first.pkg.bytes(part), part);
  assert.deepEqual(check(second.pkg).map(b => b.idmap.length), [1, 1]);
});

test('copied controls and OLE previews use their destination sheet blocks alongside notes', async t => {
  for (const kind of ['control', 'ole']) {
    const wb = await open(t, kind); if (!wb) return;
    O.copySheet(wb, wb.sheets[0], wb.sheets.length);
    O.copySheet(wb, wb.sheets[0], wb.sheets.length);
    for (const sh of wb.sheets) O.setComment(sh, 100, 10, { r: 100, c: 10, author: 'Test', text: 'note' });
    const saved = await save(wb, 'copy-blocks-' + kind), pkg = saved.pkg;
    check(pkg);
    for (const part of pkg.names.filter(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n))) {
      const vml = pkg.rels(part).find(r => /\/vmlDrawing$/.test(r.type)); if (!vml) continue;
      const previewIds = new Set(all(pkg, vml.part).map(e => e.getAttributeNS(K.KNOWN_NS.o, 'spid') || e.getAttribute('id')));
      for (const e of all(pkg, part).filter(e => ['control', 'oleObject'].includes(e.localName))) assert(previewIds.has('_x0000_s' + e.getAttribute('shapeId')), 'sheet entry and VML preview use the same id');
    }
  }
});

test('copied control names follow the worksheet numeric sequence', async t => {
  const wb = await open(t, 'control'); if (!wb) return;
  const sh = wb.sheets[0], d = L.sheetObjects.copy(sh.drawings[0]); d.id++;
  O.setDrawings(sh, sh.drawings.concat(d));
  const pkg = (await save(wb, 'control-names')).pkg;
  assert.deepEqual(all(pkg, sheet(pkg)).filter(e => e.localName === 'control').map(e => e.getAttribute('name')), ['Button 1', 'Button 2']);
});

test('ActiveX copies keep valid control names as VML ids, including names without a numeric suffix', async t => {
  for (const renamed of [false, true]) {
    let wb = await open(t, 'activex'); if (!wb) return;
    if (renamed) {
      // Derived from the public control: change only its worksheet/display name.
      const entries = wb.pkg.names.map(name => ({ name, data: /\.(xml|vml)$/.test(name) ?
        wb.pkg.text(name).replace(/CheckBox1343/g, 'cmdOK') : wb.pkg.bytes(name) }));
      const bytes = new Uint8Array(await (await L.zip.write(entries)).arrayBuffer());
      const dir = output ? path.join(output, 'original') : fs.mkdtempSync(path.join(os.tmpdir(), 'vo-activex-'));
      fs.mkdirSync(dir, { recursive: true });
      const file = path.join(dir, 'activex-named.xlsx'); fs.writeFileSync(file, bytes);
      if (!output) t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
      wb = await L.xlsxRead.read(bytes); Object.defineProperty(wb, '_input', { value: file });
    }
    const sh = wb.sheets[0], first = sh.drawings[0], name = renamed ? 'cmdOK' : 'CheckBox1343';
    const verify = pkg => {
      for (const part of pkg.names.filter(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n))) {
        const controls = all(pkg, part).filter(e => e.localName === 'control');
        const names = new Map();
        const vml = pkg.rels(part).find(r => /\/vmlDrawing$/.test(r.type));
        for (const e of controls) {
          const id = e.getAttribute('shapeId'), name = e.getAttribute('name');
          assert.match(name, /^[A-Za-z][A-Za-z0-9_]*$/, 'VBA-safe control name');
          assert(!names.has(name) || names.get(name) === id, 'unique name outside alternate branches'); names.set(name, id);
          const shape = all(pkg, vml.part).find(e => e.namespaceURI === K.KNOWN_NS.v && e.getAttributeNS(K.KNOWN_NS.o, 'spid') === '_x0000_s' + id);
          assert.equal(shape?.getAttribute('id'), name, 'VML shape id equals its ActiveX control name');
        }
      }
      check(pkg);
    };
    const addCopy = d => O.tx(wb, 'Copy control', () => {
      const copy = L.sheetObjects.copy(d); copy.id = Math.max(...sh.drawings.map(d => d.id)) + 1;
      copy.anchor.from.r += 5; copy.anchor.to.r += 5; O.setDrawings(sh, sh.drawings.concat(copy)); return copy;
    });
    addCopy(first); addCopy(sh.drawings[1]);
    const copied = await save(wb, 'activex-copy'); verify(copied.pkg);
    const names = new Set(all(copied.pkg, sheet(copied.pkg)).filter(e => e.localName === 'control').map(e => e.getAttribute('name')));
    assert.deepEqual(names, new Set(renamed ? ['cmdOK', 'cmdOK1', 'cmdOK2'] : [name, 'CheckBox1344', 'CheckBox1345']));
    wb.undo.undo(); wb.undo.undo();
    equalBytes((await save(wb, 'activex-copy-undo')).pkg.bytes(vmlPart(wb.pkg)), wb.pkg.bytes(vmlPart(wb.pkg)));
    wb.undo.redo(); wb.undo.redo(); verify((await save(wb, 'activex-copy-redo')).pkg);
    O.copySheet(wb, sh, wb.sheets.length); verify((await save(wb, 'activex-sheet-copy')).pkg);
    const reopened = await L.xlsxRead.read(copied.bytes); Object.defineProperty(reopened, '_input', { value: wb._input });
    verify((await save(reopened, 'activex-recovered')).pkg);
  }
});

test('nested VML controls and their worksheet entries share the copied sheet identities', async t => {
  const wb = await open(t, 'grouped'); if (!wb) return;
  const original = wb.sheets[0]; O.copySheet(wb, original, 1); O.copySheet(wb, original, 2);
  const saved = await save(wb, 'nested-copy'), pkg = saved.pkg;
  check(pkg);
  equalBytes(pkg.bytes(vmlPart(wb.pkg)), wb.pkg.bytes(vmlPart(wb.pkg)), 'the original group stays byte-identical');
  for (const part of pkg.names.filter(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n))) {
    const vml = pkg.rels(part).find(r => /\/vmlDrawing$/.test(r.type)); if (!vml) continue;
    const ids = new Set(all(pkg, vml.part).map(e => e.getAttributeNS(K.KNOWN_NS.o, 'spid') || e.getAttribute('id')));
    const entries = all(pkg, part).filter(e => ['control', 'oleObject'].includes(e.localName));
    assert(entries.length > 20, 'real fixture includes controls inside its VML group');
    for (const e of entries) assert(ids.has('_x0000_s' + e.getAttribute('shapeId')), 'nested entry points to its preview');
  }
  wb.undo.undo(); wb.undo.undo();
  const undone = await save(wb, 'nested-copy-undo');
  equalBytes(undone.pkg.bytes(vmlPart(wb.pkg)), wb.pkg.bytes(vmlPart(wb.pkg)));
});
