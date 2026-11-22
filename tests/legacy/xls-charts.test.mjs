import { test } from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import vm from 'node:vm';
import { words, dword, rec, str, bof, workbook, cfb } from './xls-fixtures.mjs';
import { B, X, cat, shape, drawing, convert } from './xls-draw-fixtures.mjs';

if (process.env.LEGACY_BEFORE_C4) {
  for (const file of ['xls-read', 'xls-sheet', 'xls-comments', 'xls-write']) {
    vm.runInThisContext(fs.readFileSync(process.env.LEGACY_BEFORE_C4 + '/' + file + '.js', 'utf8'));
  }
}
if (process.env.LEGACY_CHART_BEFORE_NOTICE) {
  for (const file of ['xls-chart-read', 'xls-chart-write']) {
    vm.runInThisContext(fs.readFileSync(process.env.LEGACY_CHART_BEFORE_NOTICE + '/' + file + '.js', 'utf8'));
  }
}
if (process.env.LEGACY_CHART_BEFORE_CROSSING) {
  for (const file of ['xls-chart-read', 'xls-chart-style', 'xls-chart-write']) {
    vm.runInThisContext(fs.readFileSync(process.env.LEGACY_CHART_BEFORE_CROSSING + '/' + file + '.js', 'utf8'));
  }
}
const double = value => { const b = Buffer.alloc(8); b.writeDoubleLE(value); return b; };
const number = (row, col, value) => rec(0x203, cat(words(row, col, 0), double(value)));
const group = (id, data, children = []) => [rec(id, data), rec(0x1033), ...children, rec(0x1034)];
const link = (kind, tokens = [], reference = 2) => rec(0x1051,
  cat([kind, reference], words(0, 0, tokens.length), tokens));
const area = col => cat([0x3b], words(0, 0, 2, col, col));
const series = (index, values, categories = [], extra = []) => group(0x1003, words(1, 1, 3, 3, 1, 0), [
  link(0, [], 1), rec(0x100d, cat(words(0), [6, 0], Buffer.from('Series'))),
  link(1, values), link(2, categories, categories.length ? 2 : 0),
  ...group(0x1006, words(65535, index, index, 0)), rec(0x1045, words(0)), ...extra
]);
function chart({ source = area(0), cached = false, kind = 0x1017,
  extra = [], seriesExtra = [], firstCache = 17, categoryFlags = 1 } = {}) {
  const xAxis = group(0x101d, cat(words(0), Buffer.alloc(16)), [rec(0x1020, words(1, 1, 1, categoryFlags))]);
  const yAxis = group(0x101d, cat(words(1), Buffer.alloc(16)), [
    rec(0x101f, cat(double(0), double(0), double(0), double(0), double(0), words(0x11f)))
  ]);
  const chartGroup = group(0x1014, cat(Buffer.alloc(16), words(0, 0)), [
    rec(kind, kind === 0x1017 ? words(0, 150, 0) : kind === 0x101b ? words(100, 1, 0) : words(0))
  ]);
  const records = [bof(0x20), ...extra, ...group(0x1002, cat(dword(0), dword(0),
    dword(400 * 65536), dword(250 * 65536)), [
    ...series(0, source, [], seriesExtra),
    ...group(0x1041, cat(words(0), Buffer.alloc(16)), [...xAxis, ...yAxis, ...chartGroup])
  ])];
  if (cached) records.push(rec(0x1065, words(1)), number(0, 0, firstCache), number(1, 0, 29), number(2, 0, 41));
  records.push(rec(10));
  return records;
}
const globals = [rec(0x1ae, words(1, 0x401)), rec(0x17, words(1, 0, 0, 0))];

test('embedded BIFF charts retain native series references and values from the worksheet', async () => {
  const records = [number(0, 0, 2), number(1, 0, 5), number(2, 0, 11),
    ...drawing(shape({ type: 75, objectType: 5 })), ...chart()];
  const { output, zip, xml } = await convert(records, globals);
  assert.match(xml, /<xdr:graphicFrame/);
  assert.ok(zip.has('xl/charts/chart1.xml'));
  const content = await zip.get('xl/charts/chart1.xml').text();
  assert.match(content, /<c:barChart>/);
  assert.match(content, /<c:f>Data!\$A\$1:\$A\$3<\/c:f>/);
  assert.match(content, /<c:pt idx="0"><c:v>2<\/c:v>/);
  assert.match(content, /<c:pt idx="1"><c:v>5<\/c:v>/);
  assert.match(content, /<c:pt idx="2"><c:v>11<\/c:v>/);
  assert.match(content, /<c:title>.*?<a:t>Series<\/a:t>/);
  assert.match(content, /<c:crossBetween val="between"\/>/);
  assert.match(content, /<c:spPr><a:noFill\/><a:ln><a:noFill\/><\/a:ln><\/c:spPr><\/c:plotArea>/);
  assert.equal(output.stats.features[0].charts, 1);
  assert.equal(output.losses.filter(loss => loss.notify).length, 0);
});

test('saved chart values remain authoritative when a separate cache is present', async () => {
  const { output, zip } = await convert([number(0, 0, 1),
    ...drawing(shape({ type: 75, objectType: 5 })), ...chart({ cached: true })], globals);
  assert.ok(zip.has('xl/charts/chart1.xml'));
  const content = await zip.get('xl/charts/chart1.xml').text();
  assert.match(content, /<c:pt idx="0"><c:v>17<\/c:v>/);
  assert.match(content, /<c:pt idx="2"><c:v>41<\/c:v>/);
  assert.equal(output.stats.cells, 1);
});

test('BIFF crossing settings apply to the perpendicular DrawingML axis', async () => {
  const { zip } = await convert([...drawing(shape({ type: 75, objectType: 5 })),
    ...chart({ cached: true, categoryFlags: 3 })], globals);
  const content = await zip.get('xl/charts/chart1.xml').text();
  assert.match(content.match(/<c:valAx>.*?<\/c:valAx>/s)?.[0] || '', /<c:crosses val="max"\/>/);
  assert.match(content.match(/<c:catAx>.*?<\/c:catAx>/s)?.[0] || '', /<c:crosses val="autoZero"\/>/);
});

test('a nonfinite saved chart point falls back to its worksheet value without losing other cache points', async () => {
  const { output, zip } = await convert([number(0, 0, 7),
    ...drawing(shape({ type: 75, objectType: 5 })), ...chart({ cached: true, firstCache: NaN })], globals);
  const content = await zip.get('xl/charts/chart1.xml').text();
  assert.match(content, /<c:pt idx="0"><c:v>7<\/c:v>/);
  assert.match(content, /<c:pt idx="1"><c:v>29<\/c:v>/);
  assert.equal(output.losses.filter(loss => loss.notify).length, 0);
});

test('chart sheets keep their type, chart and view without importing cache numbers as cells', async () => {
  const setup = rec(0xa1, cat(words(9, 125, 1, 1, 1, 0, 300, 300), double(0.3), double(0.3), words(1)));
  const raw = workbook([], [{ name: 'Chart1', records: [] }]);
  const bound = X.records(raw).find(r => r.id === 0x85);
  raw[bound.offset + 9] = 2;
  const offset = bound.r.u32(0);
  const stream = cat(raw.subarray(0, offset), ...chart({ source: [], cached: true, extra: [setup] }));
  const output = await X.convert(cfb(stream));
  const zip = await L.zip.read(output.pkg);
  assert.ok(zip.has('xl/chartsheets/sheet1.xml'));
  assert.ok(zip.has('xl/charts/chart1.xml'));
  assert.ok(!zip.has('xl/worksheets/sheet1.xml'));
  assert.equal(output.stats.cells, 0);
  const xml = await zip.get('xl/chartsheets/sheet1.xml').text();
  assert.match(xml, /<sheetView workbookViewId="0" tabSelected="1"/);
  assert.doesNotMatch(xml, /fitToWidth|fitToHeight|scale=|errors=/);
  assert.match(await zip.get('xl/drawings/drawing1.xml').text(), /<xdr:absoluteAnchor>/);
});

test('a broken embedded chart loses only that chart and preserves its worksheet', async () => {
  const { output, zip } = await convert([number(0, 0, 42),
    ...drawing(shape({ type: 75, objectType: 5 })), bof(0x20), rec(0x1034), rec(10)]);
  assert.ok(!zip.has('xl/charts/chart1.xml'));
  assert.match(await zip.get('xl/worksheets/sheet1.xml').text(), /<v>42<\/v>/);
  assert.ok(output.losses.some(loss => loss.id === 'legacy-xls-chart' && loss.notify));
});

test('invalid chart parents report local data loss while parent zero means a regular series', async () => {
  for (const parent of [1, 99, 0]) {
    const { output, zip } = await convert([number(0, 0, 42),
      ...drawing(shape({ type: 75, objectType: 5 })),
      ...chart({ cached: true, seriesExtra: [rec(0x104a, words(parent))] })], globals);
    assert.match(await zip.get('xl/worksheets/sheet1.xml').text(), /<v>42<\/v>/);
    const content = await zip.get('xl/charts/chart1.xml').text();
    if (parent) {
      assert.ok(output.losses.some(loss => loss.id === 'legacy-xls-chartData' && loss.notify));
      assert.doesNotMatch(content, /<c:ser>/);
    } else {
      assert.match(content, /<c:pt idx="0"><c:v>17<\/c:v>/);
      assert.equal(output.losses.filter(loss => loss.notify).length, 0);
    }
  }
});

test('BIFF5 Cyrillic chart ranges retain their local worksheet rather than losing the series link', () => {
  const tokens = cat([0x3b], words(65535), Buffer.alloc(8), words(0, 0, 1, 4), [1, 1]);
  assert.equal(X.charts.formula(tokens, { version: 'BIFF5/7', sheets: [{ name: 'Лист1' }] }, 0, {}),
    'Лист1!$B$2:$B$5');
});

test('chart axis formatting preserves date labels and explicit label rotation', () => {
  const tick = Buffer.alloc(30);
  tick.writeUInt16LE(1, 24);
  tick.writeUInt16LE(45, 28);
  const node = { r: new B.Reader(words(0)), children: [{ id: X.charts.record.TICK, r: new B.Reader(tick) }] };
  const styles = X.styles(X.read(workbook([], [{ name: 'Data', records: [] }])), []);
  const xml = X.charts.axis(node, 1, 2, 'catAx', styles, '', 0, false, false, 14);
  assert.match(xml, /<c:numFmt formatCode="m\/d\/yy" sourceLinked="1"\/>/);
  assert.match(xml, /<a:bodyPr rot="-2700000"\/>/);
});
