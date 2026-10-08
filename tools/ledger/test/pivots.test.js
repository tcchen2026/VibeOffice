/* Real-package edit tests. Set VO_CORPORA to the pinned corpus root. */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} };
require('../../../public/ledger/js/ops.js');
const O = L.ops, K = L.opc;
const corpus = process.env.VO_CORPORA || path.join(os.homedir(), 'corpora');
const rangeFile = path.join(corpus, 'excel/libreoffice__0bef51ea4c70__Pivot1_Row.xlsx');
const sources = path.join(corpus, 'results/ledger-preservation-2026-10-08/source-fixtures');
const output = process.env.PIVOT_RESULTS;
const rows = [];
async function open(t, file = rangeFile) {
  if (!fs.existsSync(file)) { t.skip('Fetch pinned corpus and generate pivot source fixtures first: ' + file); return null; }
  const wb = await L.xlsxRead.read(fs.readFileSync(file));
  Object.defineProperty(wb, '_testFile', { value: file });
  return wb;
}
async function save(wb, scenario) {
  const bytes = await L.xlsxWrite.bytes(wb);
  if (output) {
    const folder = path.join(output, scenario); fs.mkdirSync(folder, { recursive: true });
    fs.writeFileSync(path.join(folder, path.basename(wb._testFile)), bytes);
    rows.push({ file: path.basename(wb._testFile), original: wb._testFile, scenario, status: 'ok' });
    fs.writeFileSync(path.join(output, 'results.jsonl'), rows.map(r => JSON.stringify(r)).join('\n') + '\n');
  }
  return { bytes, pkg: await K.open(bytes), losses: bytes.dropped };
}
const pivot = wb => wb.extra.pivots.tables[0];
const cache = wb => wb.extra.pivots.caches[0];
const src = wb => wb.sheets.find(s => s.id === cache(wb).sources[0].sheet) || wb.sheetByName('Sheet1');
const dst = wb => wb.sheets.find(s => s.id === pivot(wb).sheet);
const rootAttr = (pkg, part, key) => pkg.xml(part)?.getAttribute(key);
const childAttr = (pkg, part, tag, key) => pkg.xml(part)?.getElementsByTagName(tag)[0]?.getAttribute(key);
const pivotParts = pkg => pkg.names.filter(n => /\/pivot(?:Cache|Tables)\//.test(n));

test('unchanged pivots, cache records and relationships keep all original bytes', async t => {
  const wb = await open(t); if (!wb) return;
  const { pkg, losses } = await save(wb, 'save');
  for (const n of pivotParts(wb.pkg)) assert.deepEqual(pkg.bytes(n), wb.pkg.bytes(n), n);
  assert.equal(losses.length, 0);
});

test('source edits request refresh once; formatting and unrelated edits do not; undo restores flags', async t => {
  const wb = await open(t); if (!wb) return;
  const sh = src(wb), original = { ...sh.get(1, 2) };
  O.tx(wb, 'format', () => O.put(sh, 1, 2, { ...original, s: wb.styles.add({ b: true }) }));
  assert.equal(cache(wb).refresh, undefined); wb.undo.undo();
  O.tx(wb, 'outside', () => O.put(sh, 30, 20, { v: 44 }));
  assert.equal(cache(wb).refresh, undefined); wb.undo.undo();
  O.tx(wb, 'source', () => O.put(sh, 1, 2, { ...original, v: 999 }));
  let s = await save(wb, 'source-edit');
  assert.equal(rootAttr(s.pkg, cache(wb).part, 'refreshOnLoad'), '1');
  assert.equal(s.losses.filter(x => x.id.startsWith('pivot-refresh:')).length, 1);
  const records = wb.pkg.rels(cache(wb).part).find(r => /\/pivotCacheRecords$/.test(r.type)).part;
  assert.deepEqual(s.pkg.bytes(records), wb.pkg.bytes(records));
  wb.undo.undo(); s = await save(wb, 'source-edit-undo');
  assert.deepEqual(s.pkg.bytes(cache(wb).part), wb.pkg.bytes(cache(wb).part));
  assert.equal(s.losses.length, 0);
  wb.undo.redo(); s = await save(wb, 'source-edit-redo');
  assert.equal(rootAttr(s.pkg, cache(wb).part, 'refreshOnLoad'), '1');
  const recovered = await L.xlsxRead.read(s.bytes), second = await L.xlsxWrite.bytes(recovered), pkg = await K.open(second);
  assert.equal(rootAttr(pkg, cache(recovered).part, 'refreshOnLoad'), '1');
});

test('source and output geometry, renames, partial deletions and shared snapshots follow undo', async t => {
  const wb = await open(t); if (!wb) return;
  O.insertLines(src(wb), 'r', 1, 2);
  let s = await save(wb, 'source-rows');
  assert.equal(childAttr(s.pkg, cache(wb).part, 'worksheetSource', 'ref'), 'A1:C8');
  wb.undo.undo(); assert.equal(cache(wb).sources[0].ref.r2, 5);
  O.renameSheet(wb, src(wb), 'Renamed & data');
  s = await save(wb, 'rename-source');
  assert.equal(childAttr(s.pkg, cache(wb).part, 'worksheetSource', 'sheet'), 'Renamed & data');
  wb.undo.undo();
  O.insertLines(dst(wb), 'r', 0, 2);
  s = await save(wb, 'output-rows');
  assert.equal(childAttr(s.pkg, pivot(wb).part, 'location', 'ref'), 'A5:C10');
  assert.equal(childAttr(s.pkg, pivot(wb).part, 'location', 'firstDataRow'), '1');
  wb.undo.undo();
  O.insertLines(dst(wb), 'r', 7, -1);
  s = await save(wb, 'output-shrink');
  assert.equal(childAttr(s.pkg, pivot(wb).part, 'location', 'ref'), 'A3:C7');
  wb.undo.undo();
  const snap = O.snapSheet(src(wb)); O.insertLines(src(wb), 'c', 1, 1); O.restoreSheet(src(wb), snap);
  assert.equal(cache(wb).refresh, undefined);
  assert.deepEqual(cache(wb).columns, [0, 1, 2]);
  O.insertLines(src(wb), 'c', 0, 3);
  O.insertLines(src(wb), 'c', 1, -1);
  assert.equal(pivot(wb).drop, undefined, 'Deleting columns before the source does not delete pivot fields');
  assert.deepEqual(cache(wb).columns, [0, 1, 2]);
  s = await save(wb, 'source-columns-before');
  assert.equal(childAttr(s.pkg, cache(wb).part, 'worksheetSource', 'ref'), 'C1:E6');
});

test('deleting a source or a used field removes definitions; output edits remain and are reported', async t => {
  const wb = await open(t); if (!wb) return;
  O.tx(wb, 'output', () => O.put(dst(wb), 3, 2, { v: 1234 }));
  let s = await save(wb, 'output-edit');
  assert(s.losses.some(x => x.id.startsWith('pivot-output:')));
  assert(s.pkg.has(pivot(wb).part)); wb.undo.undo();
  O.insertLines(src(wb), 'c', 2, -1);
  s = await save(wb, 'delete-field'); assert(!s.pkg.has(pivot(wb).part));
  assert(!s.pkg.has(cache(wb).part)); assert(s.losses.some(x => /source column/.test(x.what)));
  wb.undo.undo(); s = await save(wb, 'delete-field-undo'); assert(s.pkg.has(pivot(wb).part));
  O.deleteSheet(wb, src(wb)); s = await save(wb, 'delete-source');
  assert(!s.pkg.has(pivot(wb).part)); assert(!s.pkg.has(cache(wb).part));
  assert.equal(s.pkg.xml(s.pkg.main).getElementsByTagName('pivotCache').length, 0);
  wb.undo.undo(); s = await save(wb, 'delete-source-undo'); assert(s.pkg.has(pivot(wb).part));
  wb.undo.redo(); s = await save(wb, 'delete-source-redo'); assert(!s.pkg.has(pivot(wb).part));
});

test('a copied pivot has its own part and name while sharing the cache; deletion retains the survivor', async t => {
  const wb = await open(t); if (!wb) return;
  const source = dst(wb); O.copySheet(wb, source, wb.sheets.length);
  let s = await save(wb, 'copy-output');
  let reread = await L.xlsxRead.read(s.bytes);
  assert.equal(reread.extra.pivots.tables.length, 2);
  assert.equal(reread.extra.pivots.caches.length, 1);
  assert.equal(new Set(reread.extra.pivots.tables.map(t => t.name)).size, 2);
  assert.equal(new Set(reread.extra.pivots.tables.map(t => t.part)).size, 2);
  O.deleteSheet(wb, source); s = await save(wb, 'delete-original-output');
  reread = await L.xlsxRead.read(s.bytes);
  assert.equal(reread.extra.pivots.tables.length, 1); assert.equal(reread.extra.pivots.caches.length, 1);
  wb.undo.undo(); wb.undo.undo(); s = await save(wb, 'copy-output-undo');
  assert.equal((await L.xlsxRead.read(s.bytes)).extra.pivots.tables.length, 1);
});

test('Office-authored shared caches survive until the last pivot is removed, including undo', async t => {
  const wb = await open(t, path.join(corpus, 'excel/libreoffice__c0d5d627d116__test_diff_aggregation.xlsx')); if (!wb) return;
  const c = cache(wb), tables = wb.extra.pivots.tables.slice();
  assert.equal(tables.length, 2); assert.equal(wb.extra.pivots.caches.length, 1);
  O.deleteSheet(wb, dst(wb));
  let s = await save(wb, 'shared-delete-one');
  assert.deepEqual(s.pkg.bytes(c.part), wb.pkg.bytes(c.part));
  assert(!s.pkg.has(tables[0].part)); assert(s.pkg.has(tables[1].part));
  O.deleteSheet(wb, wb.sheets.find(sh => sh.id === tables[1].sheet));
  s = await save(wb, 'shared-delete-last'); assert(!s.pkg.has(c.part));
  wb.undo.undo(); wb.undo.undo(); s = await save(wb, 'shared-delete-undo');
  for (const part of pivotParts(wb.pkg)) assert.deepEqual(s.pkg.bytes(part), wb.pkg.bytes(part), part);
});

test('format-only structural edits do not refresh formula-based sources; changed precedents do', async t => {
  const wb = await open(t, path.join(corpus, 'excel/libreoffice__979853188c7d__pivot_table_first_header_row.xlsx')); if (!wb) return;
  assert(cache(wb).formulas);
  const sh = src(wb);
  O.structural(sh, 'Row height', () => { sh.rows[1].ht = 30; });
  assert.equal(cache(wb).refresh, undefined);
  let s = await save(wb, 'formula-format');
  assert.deepEqual(s.pkg.bytes(cache(wb).part), wb.pkg.bytes(cache(wb).part));
  O.structural(dst(wb), 'Precedent value', () => dst(wb).put(100, 30, { v: 55 }));
  s = await save(wb, 'formula-precedent');
  assert.equal(rootAttr(s.pkg, cache(wb).part, 'refreshOnLoad'), '1');
  wb.undo.undo(); assert.equal(cache(wb).refresh, undefined);
});

for (const [file, kind] of [['named-source.xlsx', 'name'], ['consolidation-source.xlsx', 'range']])
test(file + ': sources follow shifts and deletion, retaining their original source kind', async t => {
  const wb = await open(t, path.join(sources, file)); if (!wb) return;
  assert.equal(cache(wb).sources[0].kind, kind);
  O.insertLines(src(wb), 'r', 1, 1);
  let s = await save(wb, 'source-rows');
  assert.equal(rootAttr(s.pkg, cache(wb).part, 'refreshOnLoad'), '1');
  if (kind === 'name') {
    assert.equal(childAttr(s.pkg, cache(wb).part, 'worksheetSource', 'name'), 'PivotSource');
    assert.match(wb.names.find(n => n.name === 'PivotSource').ref, /7/);
  } else {
    assert.equal(childAttr(s.pkg, cache(wb).part, 'rangeSet', 'ref'), 'A1:C5');
    assert.equal(s.pkg.xml(cache(wb).part).getElementsByTagName('rangeSet')[1].getAttribute('ref'), 'E1:G4');
  }
  wb.undo.undo(); s = await save(wb, 'source-rows-undo');
  assert.deepEqual(s.pkg.bytes(cache(wb).part), wb.pkg.bytes(cache(wb).part));
  if (kind === 'name') O.tx(wb, 'delete name', () => O.setNames(wb, wb.names.filter(n => n.name !== 'PivotSource')));
  else O.insertLines(src(wb), 'c', 4, -3);
  s = await save(wb, 'delete-source'); assert(!s.pkg.has(pivot(wb).part));
});

test('table sources keep the table name; external caches stay byte-identical after local edits', async t => {
  const wb = await open(t, path.join(corpus, 'excel/poi__809b1637fb9f__ExcelPivotTableSample.xlsx')); if (!wb) return;
  const c = wb.extra.pivots.caches.find(c => c.sources[0]?.kind === 'table'); assert(c);
  const sh = wb.sheets.find(s => s.id === c.sources[0].sheet), name = c.sources[0].name;
  O.insertLines(sh, 'r', 2, 1); let s = await save(wb, 'table-rows');
  assert.equal(childAttr(s.pkg, c.part, 'worksheetSource', 'name'), name);
  assert.equal(rootAttr(s.pkg, c.part, 'refreshOnLoad'), '1');
  const ext = await open(t, path.join(corpus, 'excel/openxml-sdk__5dc27e92b3e0__Pivot4.xlsx')); if (!ext) return;
  O.tx(ext, 'local edit', () => O.put(ext.sheets[0], 100, 20, { v: 55 })); s = await save(ext, 'external-edit');
  for (const c of ext.extra.pivots.caches) assert.deepEqual(s.pkg.bytes(c.part), ext.pkg.bytes(c.part));
});
