const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} }; require('../../../public/ledger/js/ops.js');
const K = L.opc, O = L.ops, output = process.env.SHEET_PROPERTY_RESULTS, results = [];
const dir = process.env.EXCEL_CORPUS || path.join(os.homedir(), 'corpora', 'excel');
const files = { height: 'openxml-sdk__0e3494804149__Styles.xlsx',
  descent: 'libreoffice__55180c991855__tdf118668.xlsx', rtl: 'libreoffice__47ec633dafe5__tdf66668.xlsx',
  beta: 'poi__d6227fd3154b__sample-beta.xlsx', early: 'openxml-sdk__9f7806671fc8__ProjectStatusReport_TP10094814.xltx' };
const child = (e, name) => e.children.find(c => c.localName === name);
const format = (pkg, part = 'xl/worksheets/sheet1.xml') => child(pkg.xml(part), 'sheetFormatPr');
const attributes = e => Object.fromEntries(Array.from(e.attributes).filter(a => a.name !== 'xmlns' && a.prefix !== 'xmlns' && a.namespaceURI !== K.NS.mc)
  .map(a => [a.prefix ? '{' + (a.namespaceURI || e.lookupNamespaceURI(a.prefix)) + '}' + a.localName : a.name, a.value]));
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
  return { bytes, pkg };
}
test('automatic worksheet heights survive cell edits, history, drafts and sheet copies', async t => {
  const wb = await open(t, 'height'); if (!wb) return;
  const sh = wb.sheets[0], original = attributes(format(wb.pkg));
  assert.equal(original.defaultRowHeight, '16.5'); assert.notEqual(L.model.defaultRowPt(sh), 16.5);
  assert.deepEqual(attributes(format((await save(wb, 'save')).pkg)), original);
  O.tx(wb, 'Edit cell', () => O.put(sh, 0, 0, { v: 'Edited' }));
  assert.deepEqual(attributes(format((await save(wb, 'text')).pkg)), original);
  wb.undo.undo(); assert.deepEqual(attributes(format((await save(wb, 'text-undo')).pkg)), original);
  wb.undo.redo(); const changed = await save(wb, 'text-redo');
  const draft = await L.xlsxRead.read(changed.bytes); Object.defineProperty(draft, '_input', { value: wb._input });
  assert.deepEqual(attributes(format((await save(draft, 'draft-recovered')).pkg)), original);
  O.copySheet(wb, sh, wb.sheets.length); const copy = await save(wb, 'copy-sheet');
  assert.deepEqual(attributes(format(copy.pkg, 'xl/worksheets/sheet4.xml')), original);
});
test('worksheet dimension edits keep sibling attributes and undo restores exact source values', async t => {
  const wb = await open(t, 'descent'); if (!wb) return;
  const sh = wb.sheets[0], original = attributes(format(wb.pkg)), descent = '{' + K.KNOWN_NS.x14ac + '}dyDescent';
  assert.equal(original[descent], '0.2');
  assert.deepEqual(attributes(format((await save(wb, 'save')).pkg)), original);
  O.sheetProp(wb, sh, 'defColW', 20, 'Default column width');
  assert.deepEqual(attributes(format((await save(wb, 'column-width')).pkg)), { ...original, defaultColWidth: '20' });
  wb.undo.undo(); assert.deepEqual(attributes(format((await save(wb, 'column-width-undo')).pkg)), original);
  O.sheetProp(wb, sh, 'defRowH', 24, 'Default row height');
  assert.deepEqual(attributes(format((await save(wb, 'row-height')).pkg)), { ...original, defaultRowHeight: '24', customHeight: '1' });
  wb.undo.undo(); assert.deepEqual(attributes(format((await save(wb, 'row-height-undo')).pkg)), original);
  wb.undo.redo(); const changed = await save(wb, 'row-height-redo');
  const reopened = await L.xlsxRead.read(changed.bytes); assert.equal(reopened.sheets[0].defRowH, 24);
});
test('literal true and false worksheet view flags keep the authored reading direction', async t => {
  const wb = await open(t, 'rtl'); if (!wb) return;
  assert.equal(wb.sheets[0].view.rtl, true);
  const saved = await save(wb, 'save'), reopened = await L.xlsxRead.read(saved.bytes);
  assert.equal(reopened.sheets[0].view.rtl, true);
  const view = child(child(saved.pkg.xml('xl/worksheets/sheet1.xml'), 'sheetViews'), 'sheetView');
  assert.equal(view.getAttribute('rightToLeft'), '1');
});
test('pre-release worksheet sizes convert to the current schema with a notice through edits and undo', async t => {
  for (const kind of ['beta', 'early']) {
    const wb = await open(t, kind); if (!wb) continue;
    const sh = wb.sheets[0], initial = await save(wb, 'legacy-save');
    assert.equal(format(initial.pkg).namespaceURI, K.NS.s);
    assert(wb.losses.some(e => e.id === 'sheet-format-legacy:' + sh.id));
    O.sheetProp(wb, sh, 'defRowH', 26, 'Default row height');
    const edited = await save(wb, 'legacy-height');
    assert.equal(format(edited.pkg).getAttribute('defaultRowHeight'), '26');
    assert.equal((await L.xlsxRead.read(edited.bytes)).sheets[0].defRowH, 26);
    wb.undo.undo();
    assert.deepEqual(attributes(format((await save(wb, 'legacy-undo')).pkg)), attributes(format(initial.pkg)));
  }
});
test('missing worksheet format stays absent and alternatives keep unread siblings', () => {
  const wb = new L.model.Workbook(), sh = wb.addSheet('Sheet1');
  sh.extra.formatKeep = { values: L.preserve.sheetFormatValues(sh), fragment: null };
  const writer = new K.Writer(), owner = 'xl/worksheets/sheet1.xml';
  assert.equal(L.preserve.sheetFormatXML(sh, writer, owner), '');
  const xml = '<mc:AlternateContent xmlns:mc="' + K.NS.mc + '" xmlns="' + K.NS.s + '" xmlns:x14ac="' + K.KNOWN_NS.x14ac + '"><mc:Choice Requires="x14ac"><sheetFormatPr defaultRowHeight="16.5" x14ac:dyDescent="0.2"/></mc:Choice><mc:Fallback><sheetFormatPr defaultRowHeight="16.5"/></mc:Fallback></mc:AlternateContent>';
  sh.extra.formatKeep.fragment = K.fragment(K.parse(xml));
  sh.defColW = 21;
  const tree = K.parse(L.preserve.sheetFormatXML(sh, writer, owner));
  const formats = tree.getElementsByTagName('sheetFormatPr');
  assert.equal(formats.length, 2);
  for (const e of formats) { assert.equal(e.getAttribute('defaultRowHeight'), '16.5'); assert.equal(e.getAttribute('defaultColWidth'), '21'); }
  assert.equal(formats[0].getAttribute('x14ac:dyDescent'), '0.2');
});
test('a mixed worksheet alternative cannot resurrect columns through the format property', async t => {
  const wb = await open(t, 'height'); if (!wb) return;
  const sh = wb.sheets[0], xml = '<mc:AlternateContent xmlns:mc="' + K.NS.mc + '" xmlns="' + K.NS.s + '"><mc:Fallback><sheetFormatPr defaultRowHeight="16.5"/><cols><col min="1" max="1" width="80"/></cols></mc:Fallback></mc:AlternateContent>';
  const tree = K.parse(xml); K.captureAC(tree);
  L.preserve.keepSheetFormat(sh, tree.getElementsByTagName('sheetFormatPr')[0]);
  const writer = new K.Writer(wb.pkg, { doc: wb });
  const kept = K.parse(L.preserve.sheetFormatXML(sh, writer, sh.extra.ooxmlPart));
  assert.equal(kept.localName, 'sheetFormatPr');
  assert.equal(kept.getElementsByTagName('col').length, 0);
  assert(wb.losses.some(e => e.id === 'sheet-format-alternative:' + sh.id));
});
