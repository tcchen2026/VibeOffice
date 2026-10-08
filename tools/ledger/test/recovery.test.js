const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs'), path = require('node:path'), os = require('node:os');
const L = require('./load.js');
L.layout = { invalidate() {}, touchMerges() {} }; require('../../../public/ledger/js/ops.js');
const dir = process.env.EXCEL_CORPUS || path.join(os.homedir(), 'corpora/excel');
test('unreadable preservation metadata does not prevent recovery of readable workbook cells', async t => {
  const files = ['libreoffice__b3bd447d4d4b__tdf124525.xlsx', 'libreoffice__e2d45126d898__different-column-width-excel2010.xlsx',
    'poi__af72d5c76a83__clusterfuzz-testcase-minimized-POIXSSFFuzzer-5089447305609216.xlsx',
    'poi__074eb560824d__clusterfuzz-testcase-minimized-POIXSSFFuzzer-4828727001088000.xlsx'];
  if (!files.every(f => fs.existsSync(path.join(dir, f)))) return t.skip('Fetch pinned Excel corpus');
  for (const file of files) {
    const wb = await L.xlsxRead.read(fs.readFileSync(path.join(dir, file)));
    const saved = await L.xlsxWrite.bytes(wb), reopened = await L.xlsxRead.read(saved);
    assert.equal(reopened.sheets.length, wb.sheets.length, file);
    if (!wb.pkg) assert(wb.losses.some(e => e.id === 'package'), 'failed package retention is reported');
    else {
      const pkg = await L.opc.open(saved);
      for (const c of wb.extra.pivots?.caches || []) if (c.opaque) assert.deepEqual(pkg.bytes(c.part), wb.pkg.bytes(c.part));
    }
  }
});
test('unchanged mixed date/text filters retain their source ordering until their criteria change', async t => {
  const file = path.join(dir, 'libreoffice__f25f3dea692d__tdf164417.xlsx');
  if (!fs.existsSync(file)) return t.skip('Fetch pinned Excel corpus');
  const wb = await L.xlsxRead.read(fs.readFileSync(file)), sh = wb.sheets[0];
  const filters = async () => {
    const pkg = await L.opc.open(await L.xlsxWrite.bytes(wb));
    return pkg.xml('xl/worksheets/sheet1.xml').getElementsByTagName('filters')[0];
  };
  const original = ['dateGroupItem', 'filter'];
  assert.deepEqual((await filters()).children.map(e => e.localName), original);
  const changed = structuredClone(sh.autoFilter); changed.cols[0].values = ['changed']; delete changed.cols[0].dates;
  L.ops.tx(wb, 'Change filter', () => L.ops.setAutoFilter(sh, changed));
  assert.deepEqual((await filters()).children.map(e => [e.localName, e.getAttribute('val')]), [['filter', 'changed']]);
  wb.undo.undo(); assert.deepEqual((await filters()).children.map(e => e.localName), original);
});
