const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} }; require('../../../public/ledger/js/ops.js');
const K = L.opc, E = L.sheetExtensions, O = L.ops, output = process.env.EXTENSION_RESULTS, results = [];
const dir = process.env.EXCEL_CORPUS || path.join(os.homedir(), 'corpora', 'excel');
const files = { bars: 'poi__48c9cc6ad6e4__NewStyleConditionalFormattings.xlsx',
  simple: 'libreoffice__824231349a68__condformat_databar.xlsx',
  validation: 'libreoffice__18cc6e52b38e__tdf152037.xlsx',
  sparks: 'libreoffice__057bf1a72c9e__Sparklines.xlsx',
  cross: 'libreoffice__03ced28f3bd6__conditional_fmt_checkpriority.xlsx',
  opaque: 'poi__0b6bc73c95be__sheetProtection_not_protected.xlsx' };
const all = e => [e, ...e.getElementsByTagName('*')];
const xrules = tree => all(tree).filter(e => e.namespaceURI === E.NS && e.localName === 'cfRule');
function content(el) {
  return [el.namespaceURI, el.localName, Array.from(el.attributes).filter(a => a.name !== 'xmlns' && a.prefix !== 'xmlns' && el.lookupNamespaceURI(a.prefix) !== K.NS.mc)
    .map(a => [a.prefix ? el.lookupNamespaceURI(a.prefix) : '', a.localName, a.value]).sort(),
  el.childNodes.map(c => c.nodeType === 1 ? content(c) : c.textContent).filter(c => typeof c !== 'string' || c.trim())];
}
async function open(t, kind) {
  const file = path.join(dir, files[kind]); if (!fs.existsSync(file)) { t.skip('Fetch pinned Excel corpus'); return; }
  const wb = await L.xlsxRead.read(fs.readFileSync(file)); Object.defineProperty(wb, '_input', { value: file }); return wb;
}
async function save(wb, scenario) {
  const bytes = await L.xlsxWrite.bytes(wb), pkg = await K.open(bytes);
  if (output) {
    fs.mkdirSync(path.join(output, scenario), { recursive: true }); fs.writeFileSync(path.join(output, scenario, path.basename(wb._input)), bytes);
    results.push({ file: path.basename(wb._input), scenario, status: 'ok' });
    fs.writeFileSync(path.join(output, 'results.jsonl'), results.map(r => JSON.stringify(r)).join('\n') + '\n');
  }
  const ids = [];
  for (const p of pkg.names.filter(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n))) {
    const tree = pkg.xml(p), rules = xrules(tree);
    ids.push(...rules.map(e => e.getAttribute('id')));
    for (const link of all(tree).filter(e => e.namespaceURI === E.NS && e.localName === 'id')) assert(rules.some(r => r.getAttribute('id') === link.textContent), 'base rule resolves to its extension');
    const ext = tree.children.find(e => e.localName === 'extLst');
    if (ext) assert.equal(new Set(ext.children.map(e => e.getAttribute('uri'))).size, ext.children.length, 'unique extension URI');
  }
  assert.equal(new Set(ids).size, ids.length, 'unique extended rule IDs');
  return { bytes, pkg, tree: pkg.xml('xl/worksheets/sheet1.xml') };
}
test('Excel data bars and custom icons keep extended properties and original links', async t => {
  const wb = await open(t, 'bars'); if (!wb) return;
  const source = xrules(wb.pkg.xml('xl/worksheets/sheet1.xml')), saved = await save(wb, 'save');
  assert.deepEqual(xrules(saved.tree).map(content), source.map(content));
  const reopened = await L.xlsxRead.read(saved.bytes);
  assert.equal(reopened.sheets[0].cf.flatMap(c => c.rules).filter(r => r.extKeep).length, 3);
});
test('editing one bar property retains negative borders, automatic thresholds and schema order', async t => {
  const wb = await open(t, 'bars'); if (!wb) return;
  const sh = wb.sheets[0], before = xrules(wb.pkg.xml('xl/worksheets/sheet1.xml'))[0];
  const cf = structuredClone(sh.cf), rule = cf.flatMap(c => c.rules).find(r => r.bar && r.x14id);
  rule.bar.gradient = false; rule.bar.axisColor = { rgb: 'FF123456' };
  O.tx(wb, 'Bar appearance', () => O.setCF(sh, cf));
  const saved = await save(wb, 'bar-edit'), bar = xrules(saved.tree)[0].children[0];
  assert.equal(bar.getAttribute('gradient'), '0');
  assert.deepEqual(bar.children.map(e => e.localName), before.children[0].children.map(e => e.localName));
  for (const name of ['cfvo', 'negativeBorderColor', 'negativeFillColor']) assert.deepEqual(bar.children.filter(e => e.localName === name).map(content), before.children[0].children.filter(e => e.localName === name).map(content));
  assert.equal(bar.children.find(e => e.localName === 'axisColor').getAttribute('rgb'), 'FF123456');
  wb.undo.undo(); assert.deepEqual(xrules((await save(wb, 'bar-edit-undo')).tree).map(content), xrules(wb.pkg.xml('xl/worksheets/sheet1.xml')).map(content));
  wb.undo.redo(); await save(wb, 'bar-edit-redo');
  const length = structuredClone(sh.cf); length.flatMap(c => c.rules).find(r => r.bar && r.x14id).bar.minLength = 20;
  O.tx(wb, 'Bar length', () => O.setCF(sh, length));
  const resized = await save(wb, 'bar-length');
  const base = all(resized.tree).find(e => e.localName === 'dataBar' && e.namespaceURI === K.NS.s);
  assert.equal(base.getAttribute('minLength'), '20'); assert.equal(base.getAttribute('maxLength'), '100');
});
test('moving, copying and deleting bar ranges keeps links balanced through undo and drafts', async t => {
  const wb = await open(t, 'simple'); if (!wb) return;
  const sh = wb.sheets[0]; O.insertLines(sh, 'r', 0, 2);
  const moved = await save(wb, 'rows');
  assert.equal(all(moved.tree).find(e => e.namespaceURI === E.XM && e.localName === 'sqref').textContent, 'A3');
  O.copySheet(wb, sh, 1); const copied = await save(wb, 'copy-sheet');
  assert.equal((await L.xlsxRead.read(copied.bytes)).sheets[1].cf[0].rules[0].x14id === sh.cf[0].rules[0].x14id, false);
  O.tx(wb, 'Remove formatting', () => O.removeRangeRules(sh, { r1: 2, c1: 0, r2: 2, c2: 0 }, 'cf'));
  assert.equal(xrules((await save(wb, 'delete-rule')).tree).length, 0);
  wb.undo.undo(); const undone = await save(wb, 'delete-rule-undo'); assert.equal(xrules(undone.tree).length, 1);
  const draft = await L.xlsxRead.read(undone.bytes); Object.defineProperty(draft, '_input', { value: wb._input });
  await save(draft, 'draft-recovered');
});
test('extended validations retain identities and metadata while row references move', async t => {
  const wb = await open(t, 'validation'); if (!wb) return;
  const original = all(wb.pkg.xml('xl/worksheets/sheet1.xml')).filter(e => e.namespaceURI === E.NS && e.localName === 'dataValidation');
  let saved = await save(wb, 'save');
  assert.deepEqual(all(saved.tree).filter(e => e.namespaceURI === E.NS && e.localName === 'dataValidation').map(content), original.map(content));
  assert.equal(all(saved.tree).filter(e => e.namespaceURI === K.NS.s && e.localName === 'dataValidation').length, 0);
  O.insertLines(wb.sheets[0], 'r', 0, 2); saved = await save(wb, 'rows');
  assert.equal(all(saved.tree).find(e => e.namespaceURI === E.XM && e.localName === 'sqref').textContent, 'C4');
  wb.undo.undo(); await save(wb, 'rows-undo'); wb.undo.redo(); await save(wb, 'rows-redo');
});
test('sparkline properties survive save, relocation and undo without changing their uid', async t => {
  const wb = await open(t, 'sparks'); if (!wb) return;
  const original = all(wb.pkg.xml('xl/worksheets/sheet1.xml')).filter(e => e.namespaceURI === E.NS && e.localName === 'sparklineGroup');
  let saved = await save(wb, 'save');
  assert.deepEqual(all(saved.tree).filter(e => e.namespaceURI === E.NS && e.localName === 'sparklineGroup').map(content), original.map(content));
  O.insertLines(wb.sheets[0], 'r', 0, 1); await save(wb, 'rows'); wb.undo.undo(); saved = await save(wb, 'rows-undo');
  assert.deepEqual(all(saved.tree).filter(e => e.namespaceURI === E.NS && e.localName === 'sparklineGroup').map(content), original.map(content));
});
test('unknown worksheet extension URIs survive unrelated edits and copied sheets', async t => {
  const wb = await open(t, 'opaque'); if (!wb) return;
  const sh = wb.sheets[0], uri = 'http://schemas.microsoft.com/office/mac/excel/2008/main';
  const find = tree => all(tree).find(e => e.localName === 'ext' && e.getAttribute('uri') === uri);
  const original = content(find(wb.pkg.xml(sh.extra.ooxmlPart)));
  O.tx(wb, 'Edit cell', () => O.put(sh, 0, 0, { v: 'Edited' }));
  assert.deepEqual(content(find((await save(wb, 'cell-edit')).tree)), original);
  O.copySheet(wb, sh, wb.sheets.length); const saved = await save(wb, 'copy-sheet');
  const copyPart = saved.pkg.names.filter(n => /^xl\/worksheets\/[^/]+\.xml$/.test(n)).at(-1);
  assert.deepEqual(content(find(saved.pkg.xml(copyPart))), original);
});
test('a cross-sheet extension formula follows source rows and is restored by undo', async t => {
  const wb = await open(t, 'cross'); if (!wb) return;
  const sh = wb.sheets[0], before = sh.cf.flatMap(c => c.rules).filter(r => r.x14Only).map(r => r.f);
  O.insertLines(wb.sheets[1], 'r', 0, 2); await save(wb, 'source-rows');
  assert.notDeepEqual(sh.cf.flatMap(c => c.rules).filter(r => r.x14Only).map(r => r.f), before);
  wb.undo.undo(); await save(wb, 'source-rows-undo');
  assert.deepEqual(sh.cf.flatMap(c => c.rules).filter(r => r.x14Only).map(r => r.f), before);
  wb.undo.redo(); await save(wb, 'source-rows-redo');
});
