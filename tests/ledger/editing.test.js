// Range metadata through the same operations used by the grid, including undo and redo.
const { test } = require('node:test');
const assert = require('node:assert/strict');
const path = require('node:path');
const L = require('./load');
const root = process.env.VO_TEST_PUBLIC || path.resolve(__dirname, '../../public');
L.N = s => s;
L.t = (s, ...args) => s.replace(/\{(\d+)\}/g, (_, n) => args[n]);
for (const script of ['layout', 'xchart', 'ops', 'clipboard']) require(path.join(root, 'ledger/js', script));
const { model: M, formula: F, ops: O, clip: CL } = L;
const range = F.parseRange;
const refs = entries => entries.flatMap(x => x.ranges || [x.ref]).map(F.rangeName);
function book() {
  const wb = new M.Workbook(), sh = wb.addSheet('Data'), other = wb.addSheet('Report');
  let sheet = sh, sel;
  const select = (s, rg) => { sheet = s; sel = { r: rg.r1, c: rg.c1, active: 0, ranges: [rg] }; };
  L.grid = { sheet: () => sheet, sel: () => sel, selectedObject() {}, state: {}, paint() {}, guardProtect() {}, selectRange: rg => select(sheet, rg) };
  L.app = { status() {} };
  const copy = (s, rg, cut) => {
    const data = {};
    select(s, rg); CL.copy({ clipboardData: { setData: (k, v) => { data[k] = v; } } }, cut);
    CL.clearMarquee(); return data;
  };
  return { wb, sh, other, select, copy };
}
function chart(sh, ref) {
  sh.drawings.push({ kind: 'chart', anchor: { type: 'abs' }, chart: { refs: [ref], series: [{ valRef: ref }] } });
}
test('Sort moves hyperlink destinations in rows and columns, including a link crossing the selection', () => {
  for (const byCols of [false, true]) {
    const { wb, sh } = book();
    for (let i = 0; i < 3; i++) sh.cell(byCols ? 0 : i, byCols ? i : 0).v = [2, 3, 1][i];
    const link = (ref, target) => ({ ref: range(ref), target, display: 'stale' });
    sh.links = [link(byCols ? 'A2:A3' : 'B1:C1', 'two'), link(byCols ? 'C2' : 'B3', 'one')];
    O.sort(sh, range(byCols ? 'A1:C2' : 'A1:B3'), [{ index: 0 }], { byCols });
    const dest = addr => sh.links.find(l => M.rangeContains(l.ref, range(addr).r1, range(addr).c1));
    assert.equal(dest(byCols ? 'A2' : 'B1').target, 'one');
    assert.equal(dest('B2').target, 'two');
    assert.equal(dest('B2').display, undefined);
    assert.equal(dest(byCols ? 'A3' : 'C1').target, 'two');
    wb.undo.undo(); assert.equal(sh.links[0].display, 'stale');
  }
});
test('Copy/Paste takes a snapshot of a ranged link and replaces only the covered destination', () => {
  const { wb, sh, other, select, copy } = book();
  sh.cell(0, 0).v = 'Open'; sh.links.push({ ref: range('A1:B2'), target: 'source' });
  copy(sh, range('A1:A2')); sh.links[0].target = 'changed later';
  other.links.push({ ref: range('D4:F4'), target: 'original' });
  select(other, range('E4')); CL.pasteInternal({});
  assert.deepEqual(other.links.map(l => [F.rangeName(l.ref), l.target]), [['D4', 'original'], ['F4', 'original'], ['E4:E5', 'source']]);
  wb.undo.undo(); assert.deepEqual(refs(other.links), ['D4:F4']);
});
test('Cut and drag move ranges, names and chart series between sheets, and invalidate overwritten cell references', () => {
  const { wb, sh, other } = book(), report = wb.addSheet('Summary');
  sh.cell(0, 0).v = 1; sh.cell(1, 0).v = 2;
  sh.links = [{ ref: range('A1:A3'), target: 'keep' }];
  sh.dv = [{ ranges: [range('A1:A2')], type: 'whole', f1: '1' }];
  sh.cf = [{ ranges: [range('A1:A3')], rules: [{ type: 'expression', f: ['1'] }] }];
  other.dv = [{ ranges: [range('C1:C4')], type: 'whole', f1: '9' }];
  wb.names.push({ name: 'Total', ref: 'Data!$A$1:$A$2' });
  chart(report, 'Data!$A$1:$A$2'); report.cell(0, 0).f = 'Report!C2*2';
  CL.moveRange(sh, range('A1:A2'), range('C2:C3'), false, other);
  assert.deepEqual(refs(sh.links), ['A3']); assert.deepEqual(refs(other.links), ['C2:C3']);
  assert.deepEqual(refs(sh.cf), ['A3']); assert.deepEqual(refs(other.cf), ['C2:C3']);
  assert.deepEqual(refs(sh.dv), []); assert.deepEqual(refs(other.dv), ['C1', 'C4', 'C2:C3']);
  assert.equal(wb.names[0].ref, 'Report!$C$2:$C$3');
  assert.equal(report.drawings[0].chart.series[0].valRef, wb.names[0].ref);
  assert.match(report.get(0, 0).f, /#REF!/);
  wb.undo.undo(); assert.deepEqual(refs(sh.cf), ['A1:A3']); assert.equal(report.drawings[0].chart.refs[0], 'Data!$A$1:$A$2');
  wb.undo.redo(); assert.deepEqual(refs(other.links), ['C2:C3']); assert.equal(report.drawings[0].chart.refs[0], wb.names[0].ref);
});
test('Insert and Delete Cells move links, validation, CF, names and chart ranges beyond the last used cell', () => {
  const { wb, sh, other } = book(); sh.cell(0, 0).v = 1;
  sh.links = [{ ref: range('C2'), location: 'A1' }];
  sh.cf = [{ ranges: [range('C2')], rules: [] }]; sh.dv = [{ ranges: [range('C2')] }];
  wb.names.push({ name: 'Price', ref: 'Data!$C$2' }); chart(other, 'Data!$C$2:$D$2');
  O.shiftCells(sh, range('B2'), 'right', true);
  for (const records of [sh.links, sh.cf, sh.dv]) assert.deepEqual(refs(records), ['D2']);
  assert.equal(wb.names[0].ref, 'Data!$D$2'); assert.equal(other.drawings[0].chart.refs[0], 'Data!$D$2:$E$2');
  O.shiftCells(sh, range('D2'), 'left', false);
  for (const records of [sh.links, sh.cf, sh.dv]) assert.deepEqual(refs(records), []);
  assert.match(wb.names[0].ref, /#REF!/); wb.undo.undo(); assert.deepEqual(refs(sh.cf), ['D2']);
});
test('Undo and redo restore chart references on the edited sheet and on another sheet', () => {
  const { wb, sh, other } = book();
  chart(sh, 'Data!$B$1:$B$4'); chart(other, 'Data!$B$1:$B$4');
  for (const [axis, at, n, changed] of [['r', 1, 1, 'Data!$B$1:$B$5'], ['c', 0, -1, 'Data!$A$1:$A$4']]) {
    const check = expected => { for (const s of [sh, other]) { const ch = s.drawings[0].chart; assert.deepEqual(ch.refs, [expected]); assert.equal(ch.series[0].valRef, expected); } };
    O.insertLines(sh, axis, at, n); check(changed);
    wb.undo.undo(); check('Data!$B$1:$B$4'); wb.undo.redo(); check(changed); wb.undo.undo();
  }
});
test('Filtered clipboard compresses rows in every flavour and preserves formula origins and attached metadata', () => {
  const { sh, other, copy, select } = book();
  for (const [r, value] of ['Ready', 'Waitlist', 'Hand-hidden', 'Delivery'].entries()) sh.cell(r, 0).v = value;
  Object.assign(sh.rows[1], { filtered: true, hidden: true }); sh.rows[2].hidden = true; sh.rows[3].ht = 37;
  sh.cell(3, 1).f = 'A4'; sh.comments.set(M.key(3, 0), { r: 3, c: 0, text: 'last' });
  sh.dv.push({ type: 'list', ranges: [range('B2:B4')] }); sh.merges.push(range('C1:C3')); sh.cell(0, 2).v = 'Merged';
  const data = copy(sh, { r1: 0, c1: 0, r2: 3, c2: M.MAXC - 1 });
  for (const mime of ['text/plain', 'text/html']) { assert.ok(data[mime]?.includes('Hand-hidden'), mime); assert.ok(!data[mime].includes('Waitlist')); }
  assert.match(data['text/html'], /rowspan="2"/); select(other, range('A8')); CL.pasteInternal({});
  assert.deepEqual([7, 8, 9].map(r => other.val(r, 0)), ['Ready', 'Hand-hidden', 'Delivery']);
  assert.equal(other.get(9, 1).f, 'A10'); assert.equal(other.comments.get(M.key(9, 0)).text, 'last');
  assert.deepEqual(refs(other.dv), ['B9:B10']); assert.equal(F.rangeName(other.merges[0]), 'C8:C9'); assert.equal(other.rows[9].ht, 37);
});
test('CSV and text same-format settings retain delimiter, text qualifier and BOM, including UTF-16', async () => {
  for (const [name, enc, bom] of [['stock.csv', 'utf8', ''], ['stock.txt', 'utf16le', '\ufeff']]) {
    const text = "Item;Place\r\nBolts;'North;West'\r\n";
    const wb = L.csv.importBytes(Buffer.from(bom + text, enc), name);
    const bytes = await L.csv.write(wb, wb.sheets[0], name.endsWith('.csv') ? ',' : '\t', wb.textImport).arrayBuffer();
    assert.equal(Buffer.from(bytes).toString(enc), bom + text);
    const fresh = Buffer.from(await L.csv.write(wb, wb.sheets[0], ',').arrayBuffer()).toString();
    assert.ok(fresh.startsWith('\ufeffItem,Place'));
  }
});
test('Cross-workbook copy renews sheet-owned data and reports unsupported features and incoming references on move', () => {
  const { wb, sh, other } = book(), dest = new M.Workbook(); dest.addSheet('Data');
  sh.comments.set(M.key(1, 1), { r: 1, c: 1, text: 'For review' });
  wb.dxfs = [{ nf: '0.00' }]; dest.dxfs = [{ nf: 'General' }];
  sh.cf = [{ ranges: [range('A1:B2')], rules: [{ type: 'expression', dxf: 0, f: ['1'] }] }];
  sh.dv = [{ type: 'list', ranges: [range('A2')] }]; sh.links = [{ ref: range('A2'), target: 'example' }];
  sh.defRowH = 28; sh.defColW = 16; sh.outline.below = false;
  wb.names.push({ name: 'OriginalRate', ref: 'Report!A1' });
  sh.cell(0, 0).f = 'OriginalRate+Report!A1'; other.cell(1, 0).f = 'Data!A1';
  chart(sh, 'Data!A1'); sh.drawings.push({ kind: 'picture' }); sh.sparklines = [{ items: [] }]; wb.tables.push({ sheet: sh, name: 'DataTable' });
  assert.match(O.sheetTransferWarnings([sh], false).join(' '), /Charts.*Pictures.*Tables.*Sparklines.*OriginalRate.*other sheets/);
  assert.ok(!O.sheetTransferWarnings([sh], false).join(' ').includes('#REF!'));
  assert.match(O.sheetTransferWarnings([sh], true).join(' '), /#REF!/);
  const copied = O.copyToWorkbook(sh, dest, 1);
  assert.deepEqual(copied.links, sh.links); assert.deepEqual(copied.dv, sh.dv);
  assert.equal(dest.dxfs[copied.cf[0].rules[0].dxf].nf, '0.00');
  assert.equal(copied.defRowH, 28); assert.equal(copied.defColW, 16); assert.equal(copied.outline.below, false);
  assert.equal(copied.comments.get(M.key(1, 1)).text, 'For review'); assert.notEqual(copied.comments.get(M.key(1, 1)), sh.comments.get(M.key(1, 1)));
});
