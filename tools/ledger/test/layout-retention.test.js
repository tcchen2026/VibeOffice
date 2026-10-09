/* Real-file regressions: drawing coordinates and seemingly default source styles. */
const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
require('../../../public/ledger/js/layout.js');
require('../../../public/ledger/js/ops.js');
const O = L.ops, output = process.env.LAYOUT_RESULTS, results = [];
const corpus = process.env.VO_CORPORA || path.join(os.homedir(), 'corpora');
async function open(t, name) {
  const file = path.join(corpus, 'excel', name);
  if (!fs.existsSync(file)) { t.skip('Fetch pinned corpus: ' + file); return; }
  const wb = await L.xlsxRead.read(fs.readFileSync(file));
  Object.defineProperty(wb, '_input', { value: file });
  return wb;
}
async function save(wb, scenario) {
  const bytes = await L.xlsxWrite.bytes(wb), pkg = await L.opc.open(bytes);
  if (output) {
    const dir = path.join(output, scenario); fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(path.join(dir, path.basename(wb._input)), bytes);
    results.push({ file: path.basename(wb._input), original: wb._input, scenario, status: 'ok' });
    fs.writeFileSync(path.join(output, 'results.jsonl'), results.map(r => JSON.stringify(r)).join('\n') + '\n');
  }
  return { bytes, pkg, wb: await L.xlsxRead.read(bytes) };
}

test('signed drawing anchors survive saving, movement and undo; extents stay nonnegative', async t => {
  const wb = await open(t, 'libreoffice__47ec633dafe5__tdf66668.xlsx'); if (!wb) return;
  const sh = wb.sheets[0], anchors = sh.drawings.map(d => d.anchor);
  assert(anchors.some(a => a.from.cOff < 0));
  assert.deepEqual((await save(wb, 'save')).wb.sheets[0].drawings.map(d => d.anchor), anchors);
  const drawings = JSON.parse(JSON.stringify(sh.drawings));
  drawings[0].anchor.from.cOff -= 3;
  drawings[0].anchor.to.cOff -= 3;
  O.tx(wb, 'Move', () => O.setDrawings(sh, drawings));
  assert.deepEqual((await save(wb, 'move')).wb.sheets[0].drawings[0].anchor, drawings[0].anchor);
  wb.undo.undo();
  assert.deepEqual((await save(wb, 'move-undo')).wb.sheets[0].drawings.map(d => d.anchor), anchors);
  wb.undo.redo();
  assert.deepEqual((await save(wb, 'move-redo')).wb.sheets[0].drawings[0].anchor, drawings[0].anchor);

  // The same signed-coordinate contract also applies to one-cell and absolute anchors.
  const forms = JSON.parse(JSON.stringify(sh.drawings.slice(0, 2)));
  forms[0].anchor = { type: 'one', from: { c: 1, cOff: -12, r: 2, rOff: -4 }, w: 60, h: 40 };
  forms[1].anchor = { type: 'abs', x: -20, y: -10, w: 60, h: 40 };
  O.tx(wb, 'Anchor types', () => O.setDrawings(sh, forms));
  assert.deepEqual((await save(wb, 'anchor-types')).wb.sheets[0].drawings.map(d => d.anchor), forms.map(d => d.anchor));
});

test('blank cells keep source style identities through edits, clearing, row shifts and history', async t => {
  const wb = await open(t, 'libreoffice__9b556f43c795__pivottable_tabular_mode.xlsx'); if (!wb) return;
  const sh = wb.sheets[0], range = { r1: 0, r2: 0, c1: 7, c2: 7 };
  assert.equal(sh.get(0, 7)?.keep?.style, 1, 'H1 has an explicit source style even though its model style is default');
  assert.equal(sh.get(0, 7).s || 0, 0);
  sh.recalcBounds(); assert.equal(sh.maxC, 7);
  const check = async (scenario, r = 0, exists = true) => {
    const saved = await save(wb, scenario), cell = saved.wb.sheets[0].get(r, 7);
    assert.equal(cell?.keep?.style, exists ? 1 : undefined, scenario);
    return saved;
  };
  await check('save');
  O.tx(wb, 'Type', () => O.enter(sh, 0, 7, 'temporary'));
  O.clear(sh, [range], 'contents'); await check('clear-contents');
  wb.undo.undo(); assert.equal(sh.val(0, 7), 'temporary');
  wb.undo.undo(); await check('text-undo');
  O.insertLines(sh, 'r', 0, 1); await check('insert-row', 1);
  wb.undo.undo(); await check('insert-row-undo');
  O.clear(sh, [range], 'formats'); await check('clear-formats', 0, false);
  wb.undo.undo(); await check('clear-formats-undo');
  O.clear(sh, [range], 'all'); await check('clear-all', 0, false);
  wb.undo.undo(); await check('clear-all-undo');
});

test('authored table blank cells retain their explicit default-looking style', async t => {
  const wb = await open(t, 'poi__a3232f095b4d__tableStyle.xlsx'); if (!wb) return;
  const saved = await save(wb, 'table-blank-styles');
  for (const c of [1, 2]) {
    const original = wb.sheets[0].get(6, c), reopened = saved.wb.sheets[0].get(6, c);
    assert.equal(original?.keep?.style, 1, 'B7/C7 have an explicit source style');
    assert.equal(original.s || 0, 0);
    assert.equal(reopened?.keep?.style, 1, 'blank cell and style survive the writer');
    assert.equal(reopened.v, original.v);
  }
});
