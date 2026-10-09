const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} }; require('../../../public/ledger/js/ops.js');
const K = L.opc, O = L.ops, output = process.env.PAGE_SETUP_RESULTS, results = [];
const corpus = process.env.EXCEL_CORPUS || path.join(os.homedir(), 'corpora', 'excel');
const files = { absent: 'poi__9b403550715d__testSharedFormulasRangeSetBlankBug.xlsx',
  properties: 'libreoffice__d5ec3c1759df__tdf55417.xlsx', printer: 'poi__b79382456457__1_NoIden.xlsx',
  chart: 'openxml-sdk__f73c5d3ccfd4__ChartSheet.xlsx' };
const setup = (pkg, part = 'xl/worksheets/sheet1.xml') => pkg.xml(part).children.find(e => e.localName === 'pageSetup');
const attrs = e => Object.fromEntries(Array.from(e.attributes).filter(a => a.name !== 'xmlns' && a.prefix !== 'xmlns' && a.namespaceURI !== K.NS.mc).map(a => [a.name, a.value]));
async function open(t, kind) {
  const input = path.join(corpus, files[kind]); if (!fs.existsSync(input)) { t.skip('Fetch pinned Excel corpus'); return; }
  const wb = await L.xlsxRead.read(fs.readFileSync(input)); Object.defineProperty(wb, '_input', { value: input }); return wb;
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
test('absent page setup survives cell edits, copies and history until a page setting changes', async t => {
  const wb = await open(t, 'absent'); if (!wb) return;
  const sh = wb.sheets[0]; assert.equal(setup(wb.pkg), undefined);
  assert.equal(setup((await save(wb, 'save')).pkg), undefined);
  O.tx(wb, 'Edit cell', () => O.put(sh, 0, 0, { v: 'Edited' }));
  assert.equal(setup((await save(wb, 'text')).pkg), undefined);
  O.tx(wb, 'Page Setup', () => O.setPrint(sh, { ...sh.print, orientation: 'landscape' }));
  assert.deepEqual(attrs(setup((await save(wb, 'orientation')).pkg)), { orientation: 'landscape' });
  wb.undo.undo(); assert.equal(setup((await save(wb, 'orientation-undo')).pkg), undefined);
  wb.undo.redo(); const changed = await save(wb, 'orientation-redo');
  assert.equal((await L.xlsxRead.read(changed.bytes)).sheets[0].print.orientation, 'landscape');
  wb.undo.undo(); O.copySheet(wb, sh, wb.sheets.length);
  assert.equal(setup((await save(wb, 'copy-sheet')).pkg, 'xl/worksheets/sheet' + wb.sheets.length + '.xml'), undefined);
});
test('editing paper size and resolution keeps unrelated printer options and unused page-number settings', async t => {
  const wb = await open(t, 'properties'); if (!wb) return;
  const sh = wb.sheets[0], original = attrs(setup(wb.pkg));
  assert.equal(original.paperSize, '9'); assert.equal(original.firstPageNumber, '0');
  assert.deepEqual(attrs(setup((await save(wb, 'save')).pkg)), original);
  O.tx(wb, 'Page Setup', () => O.setPrint(sh, { ...sh.print, paper: 1 }));
  assert.deepEqual(attrs(setup((await save(wb, 'paper')).pkg)), { ...original, paperSize: '1' });
  O.tx(wb, 'Page Setup', () => O.setPrint(sh, { ...sh.print, dpi: 600 }));
  assert.deepEqual(attrs(setup((await save(wb, 'quality')).pkg)), { ...original, paperSize: '1', horizontalDpi: '600', verticalDpi: '600' });
  wb.undo.undo(); wb.undo.undo(); assert.deepEqual(attrs(setup((await save(wb, 'quality-undo')).pkg)), original);
});
test('retained printer-setting relationships and binary bytes survive page edits and sheet copies', async t => {
  const wb = await open(t, 'printer'); if (!wb) return;
  const sh = wb.sheets[0], ref = setup(wb.pkg).getAttribute('r:id');
  const source = wb.pkg.rels(sh.extra.ooxmlPart).find(r => r.id === ref); assert(source);
  const check = (pkg, part) => {
    const id = setup(pkg, part).getAttribute('r:id'), relation = pkg.rels(part).find(r => r.id === id);
    assert(relation); assert.equal(relation.type, source.type);
    assert.deepEqual(pkg.bytes(relation.part), wb.pkg.bytes(source.part));
  };
  check((await save(wb, 'save')).pkg, 'xl/worksheets/sheet1.xml');
  O.tx(wb, 'Page Setup', () => O.setPrint(sh, { ...sh.print, orientation: 'landscape' }));
  check((await save(wb, 'orientation')).pkg, 'xl/worksheets/sheet1.xml');
  O.copySheet(wb, sh, wb.sheets.length); const copy = await save(wb, 'copy-sheet');
  check(copy.pkg, 'xl/worksheets/sheet' + wb.sheets.length + '.xml');
});
test('chart-sheet page edits use chart-compatible settings and undo restores absent setup', async t => {
  const wb = await open(t, 'chart'); if (!wb) return;
  const sh = wb.sheets.find(s => s.kind === 'chartsheet'), part = sh.extra.ooxmlPart;
  assert.equal(setup(wb.pkg, part), undefined);
  assert.equal(setup((await save(wb, 'save')).pkg, part), undefined);
  O.tx(wb, 'Page Setup', () => O.setPrint(sh, { ...sh.print, paper: 9 }));
  assert.deepEqual(attrs(setup((await save(wb, 'paper')).pkg, part)), { paperSize: '9' });
  wb.undo.undo(); assert.equal(setup((await save(wb, 'paper-undo')).pkg, part), undefined);
});
