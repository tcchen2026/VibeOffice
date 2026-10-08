const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} }; require('../../../public/ledger/js/ops.js');
const K = L.opc, O = L.ops, results = [], output = process.env.SLICER_RESULTS;
const dir = process.env.LEDGER_FEATURES || path.join(os.homedir(), 'corpora', 'ledger-features');
const files = { slicers: 'SlicersOnPivotAndTable.xlsx', timeline: 'Timelines_Missing_21232.xlsx', shared: 'coffeeSalesProject.xlsx' };
const all = root => [root, ...root.getElementsByTagName('*')], attr = (e, k) => e.getAttribute(k);
const parts = pkg => pkg.names.filter(n => /\/(slicerCaches|slicers|timelineCaches|timelines)\/.*\.xml$/.test(n));
const drawings = pkg => pkg.names.filter(n => /^xl\/drawings\/drawing\d+\.xml$/.test(n));
const views = pkg => drawings(pkg).flatMap(n => all(pkg.xml(n)).filter(e => /\/(slicer|timeslicer)$/.test(e.namespaceURI || '') && ['slicer', 'timeslicer'].includes(e.localName)));
async function open(t, kind) {
  const file = path.join(dir, files[kind]); if (!fs.existsSync(file)) { t.skip('Fetch: sh tools/corpora.sh ~/corpora ledger-features'); return; }
  const wb = await L.xlsxRead.read(fs.readFileSync(file)); Object.defineProperty(wb, '_input', { value: file }); return wb;
}
function refs(pkg) {
  const caches = parts(pkg).filter(n => /Caches\//.test(n)).map(n => pkg.xml(n));
  const definitions = parts(pkg).filter(n => !/Caches\//.test(n)).flatMap(n => all(pkg.xml(n)).filter(e => ['slicer', 'timeline'].includes(e.localName)));
  const names = definitions.map(e => attr(e, 'name')); assert.equal(new Set(names).size, names.length, 'unique view names');
  for (const e of views(pkg)) assert(names.includes(attr(e, 'name')), 'drawing has a view definition');
  for (const e of definitions) assert(caches.some(c => attr(c, 'name') === attr(e, 'cache')), 'view has a cache');
  const ids = pkg.names.filter(n => /^xl\/tables\/.*\.xml$/.test(n)).map(n => attr(pkg.xml(n), 'id'));
  const sheets = all(pkg.xml(pkg.main)).filter(e => e.localName === 'sheet');
  for (const cache of caches) for (const el of all(cache)) {
    if (el.localName === 'tableSlicerCache') assert(ids.includes(attr(el, 'tableId')), 'table cache source exists');
    if (el.localName === 'pivotTable') {
      const sh = sheets.find(e => attr(e, 'sheetId') === attr(el, 'tabId')); assert(sh, 'pivot sheet exists');
      const rid = sh.getAttributeNS(K.NS.rel, 'id'), part = pkg.rels(pkg.main).find(r => r.id === rid).part;
      assert(pkg.rels(part).filter(r => /\/pivotTable$/.test(r.type)).some(r => attr(pkg.xml(r.part), 'name') === attr(el, 'name')), 'cache pivot exists on its declared sheet');
    }
  }
}
async function save(wb, scenario) {
  const bytes = await L.xlsxWrite.bytes(wb), pkg = await K.open(bytes);
  if (output) {
    const dest = path.join(output, scenario); fs.mkdirSync(dest, { recursive: true }); fs.writeFileSync(path.join(dest, path.basename(wb._input)), bytes);
    results.push({ file: path.basename(wb._input), scenario, status: 'ok' });
    fs.writeFileSync(path.join(output, 'results.jsonl'), results.map(r => JSON.stringify(r)).join('\n') + '\n');
  }
  refs(pkg); return { bytes, pkg, losses: bytes.dropped };
}
test('all seven slicer/timeline views retain frames and twelve original definition/cache parts', async t => {
  let count = 0, kept = 0;
  for (const kind of Object.keys(files)) {
    const wb = await open(t, kind); if (!wb) return;
    const saved = await save(wb, 'save'); count += views(saved.pkg).length;
    for (const p of parts(wb.pkg)) { assert.deepEqual(Buffer.from(saved.pkg.bytes(p)), Buffer.from(wb.pkg.bytes(p)), p); kept++; }
    const reopened = await L.xlsxRead.read(saved.bytes);
    assert.equal(reopened.sheets.flatMap(s => s.drawings.filter(d => d.slicerKeep)).length, views(wb.pkg).length);
  }
  assert.equal(count, 7); assert.equal(kept, 12);
});
test('timeline source edits refresh the pivot while moves and history preserve its frame', async t => {
  const wb = await open(t, 'timeline'); if (!wb) return;
  const sh = wb.sheets[0], d = sh.drawings[0], row = d.anchor.from.r;
  O.insertLines(sh, 'r', 0, 2); const moved = await save(wb, 'rows');
  assert.equal((await L.xlsxRead.read(moved.bytes)).sheets[0].drawings[0].anchor.from.r, row + 2);
  wb.undo.undo(); await save(wb, 'rows-undo'); wb.undo.redo(); await save(wb, 'rows-redo');
  const source = wb.sheets[1]; O.tx(wb, 'Change source', () => O.put(source, 2, 1, { v: 73 }));
  const edited = await save(wb, 'source-edit'); assert(edited.losses.some(e => e.id.startsWith('pivot-refresh:')));
  assert.equal(views(edited.pkg).length, 1);
});
test('copying a view keeps its source and remaps both the drawing and definition names', async t => {
  const wb = await open(t, 'slicers'); if (!wb) return;
  const sh = wb.sheets[0], d = L.sheetObjects.copy(sh.drawings[0]); d.id++;
  d.anchor.from.r += 16; d.anchor.to.r += 16;
  O.tx(wb, 'Copy view', () => O.setDrawings(sh, sh.drawings.concat(d)));
  const saved = await save(wb, 'copy-view'); assert.equal(views(saved.pkg).length, 3);
  wb.undo.undo(); await save(wb, 'copy-view-undo'); wb.undo.redo(); await save(wb, 'copy-view-redo');
});
test('a copied table sheet has an independent slicer cache that survives source-sheet deletion', async t => {
  const wb = await open(t, 'slicers'); if (!wb) return;
  const source = wb.sheets[0]; O.copySheet(wb, source, 1);
  let saved = await save(wb, 'copy-sheet'); assert.equal(views(saved.pkg).length, 3);
  O.deleteSheet(wb, source); saved = await save(wb, 'copy-sheet-delete-source');
  assert.equal(views(saved.pkg).length, 1); assert(saved.losses.some(e => e.action === 'drop'));
  wb.undo.undo(); await save(wb, 'copy-sheet-delete-source-undo');
});
test('shared filter caches remove only the deleted pivot and undo restores all connections', async t => {
  const wb = await open(t, 'shared'); if (!wb) return;
  O.deleteSheet(wb, wb.sheetByName('PerCountry'));
  const saved = await save(wb, 'delete-shared-pivot'); assert.equal(views(saved.pkg).length, 4);
  for (const n of parts(saved.pkg).filter(n => /Caches\//.test(n))) assert.equal(all(saved.pkg.xml(n)).filter(e => e.localName === 'pivotTable').length, 2);
  wb.undo.undo(); const undone = await save(wb, 'delete-shared-pivot-undo');
  for (const n of parts(wb.pkg)) assert.deepEqual(Buffer.from(undone.pkg.bytes(n)), Buffer.from(wb.pkg.bytes(n)), n);
});
test('deleting a source removes its timeline, cache, workbook entry and frame; undo restores them', async t => {
  const wb = await open(t, 'timeline'); if (!wb) return;
  O.deleteSheet(wb, wb.sheets[1]); let saved = await save(wb, 'delete-source');
  assert.equal(views(saved.pkg).length, 0); assert.equal(parts(saved.pkg).length, 0);
  wb.undo.undo(); saved = await save(wb, 'delete-source-undo'); assert.equal(views(saved.pkg).length, 1);
});
test('editing the fallback converts only that view, with an undoable notice', async t => {
  const wb = await open(t, 'slicers'); if (!wb) return;
  const sh = wb.sheets[0], d = sh.drawings[0];
  O.tx(wb, 'Edit fallback', () => O.setDrawings(sh, [{ ...d, dirty: true, text: 'Edited fallback' }]));
  const saved = await save(wb, 'content-edit'); assert.equal(views(saved.pkg).length, 1);
  assert(saved.losses.some(e => e.id.startsWith('slicer-conversion:')));
  wb.undo.undo(); const undone = await save(wb, 'content-edit-undo'); assert.equal(views(undone.pkg).length, 2);
  assert(!undone.losses.some(e => e.id.startsWith('slicer-conversion:')));
});
