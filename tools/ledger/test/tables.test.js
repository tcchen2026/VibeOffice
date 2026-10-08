const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} }; require('../../../public/ledger/js/ops.js');
const K = L.opc, O = L.ops, output = process.env.QUERY_RESULTS, results = [];
const dir = process.env.EXCEL_CORPUS || path.join(os.homedir(), 'corpora', 'excel');
const files = { basic: 'openxml-sdk__a5ef25d2c713__QueryTable.xlsx', ids: 'libreoffice__3ee733eedc60__tdf141547.xlsx',
  empty: 'libreoffice__35d821135eaa__TableEmptyHeaders.xlsx', range: 'libreoffice__a278bf3eff36__tdf166712.xlsx' };
const all = tree => [tree, ...tree.getElementsByTagName('*')], queryParts = pkg => pkg.names.filter(n => /^xl\/queryTables\/[^/]+\.xml$/.test(n));
async function open(t, kind) {
  const file = path.join(dir, files[kind]); if (!fs.existsSync(file)) { t.skip('Fetch pinned Excel corpus'); return; }
  const wb = await L.xlsxRead.read(fs.readFileSync(file)); Object.defineProperty(wb, '_input', { value: file }); return wb;
}
function refs(pkg) {
  for (const part of pkg.names.filter(n => /^xl\/tables\/[^/]+\.xml$/.test(n))) {
    const tree = pkg.xml(part), columns = all(tree).filter(e => e.localName === 'tableColumn');
    const ids = columns.map(c => c.getAttribute('id')); assert.equal(ids.length, new Set(ids).size, 'unique table column IDs');
    for (const rel of pkg.rels(part).filter(r => /\/queryTable$/.test(r.type))) {
      assert.equal(tree.getAttribute('tableType'), 'queryTable');
      const fields = all(pkg.xml(rel.part)).filter(e => e.localName === 'queryTableField');
      for (const c of columns) if (c.hasAttribute('queryTableFieldId')) assert(fields.some(f => f.getAttribute('id') === c.getAttribute('queryTableFieldId') && f.getAttribute('tableColumnId') === c.getAttribute('id')), 'query and table field identities match');
    }
  }
}
async function save(wb, scenario) {
  const bytes = await L.xlsxWrite.bytes(wb), pkg = await K.open(bytes);
  if (output) {
    fs.mkdirSync(path.join(output, scenario), { recursive: true }); fs.writeFileSync(path.join(output, scenario, path.basename(wb._input)), bytes);
    results.push({ file: path.basename(wb._input), scenario, status: 'ok' });
    fs.writeFileSync(path.join(output, 'results.jsonl'), results.map(r => JSON.stringify(r)).join('\n') + '\n');
  }
  refs(pkg); return { bytes, pkg, losses: bytes.dropped };
}
test('query definitions, original field IDs and custom column attributes survive ordinary saves', async t => {
  for (const kind of ['basic', 'ids', 'empty', 'range']) {
    const wb = await open(t, kind); if (!wb) return;
    const saved = await save(wb, 'save');
    for (const p of queryParts(wb.pkg)) assert.deepEqual(Buffer.from(saved.pkg.bytes(p)), Buffer.from(wb.pkg.bytes(p)), p);
    for (const tab of wb.tables) {
      const a = wb.pkg.xml(tab.ooxmlPart), b = saved.pkg.xml(tab.ooxmlPart);
      assert.equal(b.getAttribute('tableType'), a.getAttribute('tableType'));
      for (const prop of ['id', 'uniqueName', 'queryTableFieldId', 'name']) assert.deepEqual(all(b).filter(e => e.localName === 'tableColumn').map(e => e.getAttribute(prop)), all(a).filter(e => e.localName === 'tableColumn').map(e => e.getAttribute(prop)));
    }
  }
});
test('inserted query-table columns are unbound; retained columns keep their identities through history', async t => {
  const wb = await open(t, 'basic'); if (!wb) return;
  const tab = wb.tables[0], sh = tab.sheet, ids = tab.columns.map(c => c.id), before = JSON.stringify(tab.columns);
  O.insertLines(sh, 'c', tab.ref.c1 + 2, 1);
  assert.deepEqual(tab.columns.map(c => c.id), [ids[0], ids[1], Math.max(...ids) + 1, ...ids.slice(2)]);
  const saved = await save(wb, 'insert-column'), q = saved.pkg.xml(queryParts(saved.pkg)[0]);
  const added = all(q).find(e => e.localName === 'queryTableField' && e.getAttribute('dataBound') === '0');
  assert(added); assert.equal(added.getAttribute('tableColumnId'), String(Math.max(...ids) + 1));
  wb.undo.undo(); assert.equal(JSON.stringify(tab.columns), before); const undone = await save(wb, 'insert-column-undo');
  assert.deepEqual(Buffer.from(undone.pkg.bytes(queryParts(wb.pkg)[0])), Buffer.from(wb.pkg.bytes(queryParts(wb.pkg)[0])));
  wb.undo.redo(); await save(wb, 'insert-column-redo');
});
test('deleted fields stay excluded from query refresh and full-table deletion is undoable', async t => {
  const wb = await open(t, 'ids'); if (!wb) return;
  const tab = wb.tables[0], sh = tab.sheet, name = tab.columns[0].name, remaining = tab.columns[1].id;
  O.insertLines(sh, 'c', tab.ref.c1, -1); let saved = await save(wb, 'delete-column');
  assert.equal(tab.columns[0].id, remaining);
  assert(all(saved.pkg.xml(queryParts(saved.pkg)[0])).some(e => e.localName === 'deletedField' && e.getAttribute('name') === name));
  O.insertLines(sh, 'c', tab.ref.c1, -1); saved = await save(wb, 'delete-table');
  assert.equal(wb.tables.length, 0); assert.equal(queryParts(saved.pkg).length, 0);
  assert(saved.losses.some(e => e.action === 'drop'));
  wb.undo.undo(); assert.equal(wb.tables.length, 1); await save(wb, 'delete-table-undo');
  wb.undo.undo(); await save(wb, 'delete-column-undo');
});
test('editing query results or a header keeps the external field and undo does not retain save mutations', async t => {
  const wb = await open(t, 'basic'); if (!wb) return;
  const tab = wb.tables[0], sh = tab.sheet, name = tab.columns[0].name;
  O.tx(wb, 'Edit query result', () => { O.put(sh, tab.ref.r1, tab.ref.c1, { v: 'Local label' }); O.put(sh, tab.ref.r1 + 1, tab.ref.c1, { v: 77 }); });
  let saved = await save(wb, 'edit-result');
  assert.equal(tab.columns[0].name, name, 'save does not mutate history');
  assert.equal(all(saved.pkg.xml(tab.ooxmlPart)).find(e => e.localName === 'tableColumn').getAttribute('name'), 'Local label');
  assert.deepEqual(Buffer.from(saved.pkg.bytes(queryParts(wb.pkg)[0])), Buffer.from(wb.pkg.bytes(queryParts(wb.pkg)[0])));
  wb.undo.undo(); saved = await save(wb, 'edit-result-undo');
  assert.equal(all(saved.pkg.xml(tab.ooxmlPart)).find(e => e.localName === 'tableColumn').getAttribute('name'), name);
});
test('copied tables get independent query parts and survive deleting the original sheet', async t => {
  const wb = await open(t, 'empty'); if (!wb) return;
  const sh = wb.tables[0].sheet; O.copySheet(wb, sh, wb.sheets.length);
  let saved = await save(wb, 'copy-sheet'); assert.equal(queryParts(saved.pkg).length, 2);
  const uid = [];
  for (const p of [...queryParts(saved.pkg), ...saved.pkg.names.filter(n => /^xl\/tables\/[^/]+\.xml$/.test(n))]) for (const e of all(saved.pkg.xml(p))) for (const a of e.attributes) if (a.localName === 'uid') uid.push(a.value);
  assert.equal(new Set(uid).size, uid.length);
  O.deleteSheet(wb, sh); saved = await save(wb, 'copy-delete-source'); assert.equal(queryParts(saved.pkg).length, 1);
  wb.undo.undo(); await save(wb, 'copy-delete-source-undo');
});
test('range query copies retain sheet-scoped destination names through moves, undo and drafts', async t => {
  const wb = await open(t, 'range'); if (!wb) return;
  const sh = wb.sheets[0], before = JSON.stringify(wb.names);
  const copy = O.copySheet(wb, sh, 0, 'Copied query');
  const local = wb.names.find(n => n.scope === 0); assert(local.ref.includes('Copied query'));
  let saved = await save(wb, 'copy-range-query'); assert.equal(queryParts(saved.pkg).length, 2);
  wb.undo.undo(); assert.equal(JSON.stringify(wb.names), before); await save(wb, 'copy-range-query-undo'); wb.undo.redo();
  O.insertLines(copy, 'r', 0, 2); saved = await save(wb, 'range-rows');
  const draft = await L.xlsxRead.read(saved.bytes); Object.defineProperty(draft, '_input', { value: wb._input });
  await save(draft, 'draft-recovered'); assert(draft.names.find(n => n.scope === 0).ref.includes('$A$3'));
});
